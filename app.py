import functions_framework
import docx
import zipfile
from lxml import etree
import io
import json
import re

# XML Namespace constants
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TEXT = WORD_NAMESPACE + 't'
COMMENT = WORD_NAMESPACE + 'comment'

def clean_text(text):
    """
    1. Removes leading/trailing whitespace.
    2. Collapses multiple internal spaces into one (e.g. "Hello    World" -> "Hello World").
    """
    if not text:
        return ""
    # Replace multiple whitespace characters (spaces, tabs, newlines) with a single space
    return re.sub(r'\s+', ' ', text).strip()

def get_comments_from_xml(docx_file_stream):
    """Parses inner XML for comments."""
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
                # Join text nodes, then clean the result
                raw_text = ''.join([node.text for node in comment.iter(TEXT) if node.text])
                
                comments_list.append({
                    "id": c_id,
                    "author": author,
                    "content": clean_text(raw_text)
                })
    except Exception:
        return []
    return comments_list

@functions_framework.http
def extract_word_data(request):
    if request.method == 'OPTIONS':
        headers = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Max-Age': '3600'
        }
        return ('', 204, headers)

    if request.method != 'POST':
        return ('Only POST requests are accepted', 405)

    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return ('No file part in the request', 400)

    try:
        file_bytes = uploaded_file.read()

        # --- 1. Structure-Aware Text Extraction ---
        doc_stream = io.BytesIO(file_bytes)
        doc = docx.Document(doc_stream)
        
        structured_content = []
        
        for para in doc.paragraphs:
            cleaned = clean_text(para.text)
            if not cleaned:
                continue # Skip empty lines

            style_name = para.style.name.lower()
            
            # Categorize based on Word Styles
            if 'heading' in style_name:
                # Extract level number if possible (e.g., "Heading 1" -> 1)
                try:
                    level = int(style_name.replace('heading', '').strip())
                except ValueError:
                    level = 1 # Fallback
                
                structured_content.append({
                    "type": "heading",
                    "level": level,
                    "text": cleaned
                })
            elif 'title' in style_name:
                structured_content.append({
                    "type": "heading",
                    "level": 0, # Top level
                    "text": cleaned
                })
            elif 'list' in style_name:
                structured_content.append({
                    "type": "list_item",
                    "text": cleaned
                })
            else:
                structured_content.append({
                    "type": "paragraph",
                    "text": cleaned
                })

        # --- 2. Comment Extraction ---
        zip_stream = io.BytesIO(file_bytes)
        comments_data = get_comments_from_xml(zip_stream)

        # Construct final JSON response
        response_data = {
            "filename": uploaded_file.filename,
            "structure": structured_content, # The hierarchy-aware list
            "comments": comments_data,
            "metadata": {
                "total_blocks": len(structured_content),
                "comment_count": len(comments_data)
            }
        }

        return (json.dumps(response_data), 200, {'Content-Type': 'application/json'})

    except Exception as e:
        return (f"Error processing file: {str(e)}", 500)
