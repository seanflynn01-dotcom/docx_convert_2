from flask import Flask, request, jsonify
import docx
import zipfile
from lxml import etree
import io
import re

app = Flask(__name__)

# XML Constants
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TEXT = WORD_NAMESPACE + 't'
COMMENT = WORD_NAMESPACE + 'comment'

def clean_text(text):
    if not text:
        return ""
    return re.sub(r'\s+', ' ', text).strip()

def get_comments_from_xml(docx_file_stream):
    comments_list = []
    try:
        with zipfile.ZipFile(docx_file_stream) as zf:
            if 'word/comments.xml' not in zf.namelist():
                return []
            
            xml_content = zf.read('word/comments.xml')
            tree = etree.fromstring(xml_content)

            for comment in tree.iter(COMMENT):
                author = comment.get(WORD_NAMESPACE + 'author', 'Unknown')
                c_id = comment.get(WORD_NAMESPACE + 'id')
                raw_text = ''.join([node.text for node in comment.iter(TEXT) if node.text])
                
                comments_list.append({
                    "id": c_id,
                    "author": author,
                    "content": clean_text(raw_text)
                })
    except Exception:
        return []
    return comments_list

@app.route('/', methods=['POST'])
def extract_word_data():
    # check if the post request has the file part
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    try:
        # Read file into memory once
        file_bytes = file.read()

        # 1. Structure-Aware Text Extraction
        doc_stream = io.BytesIO(file_bytes)
        doc = docx.Document(doc_stream)
        
        structured_content = []
        for para in doc.paragraphs:
            cleaned = clean_text(para.text)
            if not cleaned: continue

            style_name = para.style.name.lower()
            
            if 'heading' in style_name:
                try:
                    level = int(style_name.replace('heading', '').strip())
                except ValueError:
                    level = 1
                structured_content.append({"type": "heading", "level": level, "text": cleaned})
            elif 'title' in style_name:
                structured_content.append({"type": "heading", "level": 0, "text": cleaned})
            elif 'list' in style_name:
                structured_content.append({"type": "list_item", "text": cleaned})
            else:
                structured_content.append({"type": "paragraph", "text": cleaned})

        # 2. Comment Extraction
        zip_stream = io.BytesIO(file_bytes)
        comments_data = get_comments_from_xml(zip_stream)

        return jsonify({
            "filename": file.filename,
            "structure": structured_content,
            "comments": comments_data,
            "metadata": {
                "total_blocks": len(structured_content),
                "comment_count": len(comments_data)
            }
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # This is used when running locally only. 
    # production usage relies on Gunicorn (see below).
    app.run(debug=True, host='0.0.0.0', port=5000)
