import base64
import os
import fitz
import io
import logging
from io import BytesIO
from typing import Dict, Any, List, Optional
from datetime import datetime
from .common import logger, SHP_DOC_LIBRARY, sp_context

logger = logging.getLogger(__name__)

# Helper function to safely convert to ISO format
def _to_iso_optional(dt_obj: Optional[datetime]) -> Optional[str]:
    """Converts a datetime object to ISO format string, or returns None if the object is None."""
    if dt_obj is not None:
        return dt_obj.isoformat()
    return None

def _get_sp_path(sub_path: Optional[str] = None) -> str:
    """Create a properly formatted SharePoint path"""
    return f"{SHP_DOC_LIBRARY}/{sub_path or ''}".rstrip('/')

def list_folders(parent_folder: Optional[str] = None) -> List[Dict[str, Any]]:
    """List folders in the specified directory or root if not specified"""
    path = _get_sp_path(parent_folder)
    log_location = parent_folder or "root directory"
    logger.info(f"Listing folders in {log_location}")
    
    # Use the ClientObject.get_items() method which handles loading automatically
    parent = sp_context.web.get_folder_by_server_relative_url(path)
    folders = parent.folders
    sp_context.load(folders, ["ServerRelativeUrl", "Name", "TimeCreated", "TimeLastModified"])
    sp_context.execute_query()
    
    # Convert directly to the required format
    return [{
        "name": f.name,
        "url": f.properties.get("ServerRelativeUrl"),
        "created": _to_iso_optional(f.properties.get("TimeCreated")),
        "modified": _to_iso_optional(f.properties.get("TimeLastModified"))
    } for f in folders]

def list_documents(folder_name: str) -> List[Dict[str, Any]]:
    """List all documents in a specified folder"""
    logger.info(f"Listing documents in folder: {folder_name}")
    path = _get_sp_path(folder_name)
    
    # Load files with specific properties to reduce data transfer
    folder = sp_context.web.get_folder_by_server_relative_url(path)
    files = folder.files
    sp_context.load(files, ["ServerRelativeUrl", "Name", "Length", "TimeCreated", "TimeLastModified"])
    sp_context.execute_query()
    
    # Convert directly to the required format
    return [{
        "name": f.name,
        "url": f.properties.get("ServerRelativeUrl"),
        "size": f.properties.get("Length"),
        "created": _to_iso_optional(f.properties.get("TimeCreated")),
        "modified": _to_iso_optional(f.properties.get("TimeLastModified"))
    } for f in files]

def extract_text_from_pdf(pdf_content):
    """Extract text from PDF using PyMuPDF with fallback methods."""
    
    # Process the PDF
    try:
        pdf_document = fitz.open(stream=pdf_content, filetype="pdf")
        page_count = len(pdf_document)
        text_content = ""
        
        for page_num in range(page_count):
            page = pdf_document[page_num]
            text_content += page.get_text() + "\n"
        
        pdf_document.close()
        return text_content.strip(), page_count
    except Exception as e:
        logger.error(f"Error extracting text from PDF: {e}")
        raise

def get_folder_tree(parent_folder: Optional[str] = None) -> Dict[str, Any]:
    """Recursively list all folders and files from a parent folder."""
    path = _get_sp_path(parent_folder)
    logger.info(f"Building tree for {parent_folder or 'root'} at path: {path}")

    folder = _load_sp_folder(path)
    if not folder:
        return {"name": os.path.basename(path), "path": path, "type": "folder", "error": "Could not access folder", "children": []}

    children = _get_folder_children(parent_folder)

    return {
        "name": folder.name,
        "path": folder.properties.get("ServerRelativeUrl"),
        "type": "folder",
        "created": _to_iso_optional(folder.properties.get("TimeCreated")),
        "modified": _to_iso_optional(folder.properties.get("TimeLastModified")),
        "children": children
    }

def _load_sp_folder(path: str):
    """Load folder object from SharePoint."""
    try:
        folder = sp_context.web.get_folder_by_server_relative_url(path)
        sp_context.load(folder, ["Name", "ServerRelativeUrl", "TimeCreated", "TimeLastModified"])
        sp_context.execute_query()
        return folder
    except Exception as e:
        logger.error(f"Failed to load folder '{path}': {e}")
        return None

def _get_folder_children(parent_folder: Optional[str]):
    """Return list of child folders and files."""
    children = []
    try:
        site_title = sp_context.web.properties.get('Title') or _load_web_title()
        lib_prefix = f"/sites/{site_title}/{SHP_DOC_LIBRARY}"

        # Subfolders
        for f in list_folders(parent_folder):
            if f["url"].startswith(lib_prefix):
                children.append(get_folder_tree(f["url"][len(lib_prefix):].lstrip('/')))
            else:
                logger.warning(f"Cannot determine relative path for {f['url']}")

        # Files
        children.extend({
            "name": f["name"],
            "path": f["url"],
            "type": "file",
            "size": f.get("size"),
            "created": f.get("created"),
            "modified": f.get("modified"),
        } for f in list_documents(parent_folder))

    except Exception as e:
        logger.error(f"Error listing children for '{parent_folder}': {e}")
    return children

def _load_web_title() -> str:
    if not sp_context.web.is_property_available('Title'):
        sp_context.load(sp_context.web, ['Title'])
        sp_context.execute_query()
    return sp_context.web.properties['Title']

def get_document_content(folder_name: str, file_name: str) -> dict:
    """Retrieve document content; supports PDF text extraction."""
    file_path = _get_sp_path(f"{folder_name}/{file_name}")
    file = sp_context.web.get_file_by_server_relative_url(file_path)
    sp_context.load(file, ["Exists", "Length", "Name"])
    sp_context.execute_query()
    logger.info(f"File exists: {file.exists}, size: {file.length}")

    content = io.BytesIO()
    file.download(content)
    sp_context.execute_query()
    content_bytes = content.getvalue()

    return _process_file_content(file_name, content_bytes)

def _process_file_content(file_name: str, content: bytes) -> dict:
    lower_name = file_name.lower()
    if lower_name.endswith('.pdf'):
        try:
            text, pages = extract_text_from_pdf(content)
            return {"name": file_name, "content_type": "text", "content": text, "original_type": "pdf", "page_count": pages, "size": len(content)}
        except Exception as e:
            logger.warning(f"PDF processing failed: {e}")
            return _binary_file_response(file_name, content, "pdf")

    if lower_name.endswith(('.txt', '.csv', '.json', '.xml', '.html', '.md', '.js', '.css', '.py')):
        try:
            return {"name": file_name, "content_type": "text", "content": content.decode('utf-8'), "size": len(content)}
        except UnicodeDecodeError:
            return _binary_file_response(file_name, content)

    return _binary_file_response(file_name, content)

def _binary_file_response(file_name: str, content: bytes, original_type: Optional[str] = None) -> dict:
    return {"name": file_name, "content_type": "binary", "content_base64": base64.b64encode(content).decode(), "original_type": original_type, "size": len(content)}
