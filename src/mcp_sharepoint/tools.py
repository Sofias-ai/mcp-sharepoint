import base64, os
from functools import wraps
from typing import Optional, Dict, Any
from .common import logger, mcp, SHP_DOC_LIBRARY, sp_context
from .resources import list_folders, list_documents, get_document_content, get_folder_tree, download_document

# Helper functions to reduce code duplication
def _get_path(folder: str = "", file: Optional[str] = None) -> str:
    """Construct SharePoint path from components"""
    path = f"{SHP_DOC_LIBRARY}/{folder}".rstrip('/')
    return f"{path}/{file}" if file else path

def _handle_sp_operation(func):
    """Decorator for SharePoint operations with error handling"""
    @wraps(func)
    async def wrapper(*args, **kwargs):
        try:
            return await func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error in {func.__name__}: {str(e)}")
            return {"success": False, "message": f"Operation failed: {str(e)}"}
    return wrapper

def _file_success_response(file_obj, message: str) -> Dict[str, Any]:
    """Standard success response for file operations"""
    return {
        "success": True,
        "message": message,
        "file": {"name": file_obj.name, "url": file_obj.serverRelativeUrl}
    }

# Tool implementations
@mcp.tool(name="List_SharePoint_Folders", description="List folders in the specified SharePoint directory or root if not specified")
async def list_folders_tool(parent_folder: Optional[str] = None):
    """List folders in the specified SharePoint directory or root if not specified"""
    return list_folders(parent_folder)

@mcp.tool(name="List_SharePoint_Documents", description="List all documents in a specified SharePoint folder")
async def list_documents_tool(folder_name: str):
    """List all documents in a specified SharePoint folder"""
    return list_documents(folder_name)

@mcp.tool(name="Get_SharePoint_Tree", description="Get a recursive tree view of a SharePoint folder")
async def get_sharepoint_tree_tool(parent_folder: Optional[str] = None):
    """Get a recursive tree view of a SharePoint folder."""
    return get_folder_tree(parent_folder)

@mcp.tool(name="Get_Document_Content", description="Get content of a document in SharePoint")
async def get_document_content_tool(folder_name: str, file_name: str):
    """Get content of a document in SharePoint"""
    return get_document_content(folder_name, file_name)

@mcp.tool(name="Create_Folder", description="Create a new folder in the specified directory or root if not specified")
@_handle_sp_operation
async def create_folder(folder_name: str, parent_folder: Optional[str] = None):
    """Create a new folder in the specified directory or root if not specified"""
    parent_path = _get_path(parent_folder or "")
    logger.info(f"Creating folder '{folder_name}' in {parent_folder or 'root directory'}")
    
    # Check for existing folder
    if any(f["name"] == folder_name for f in list_folders(parent_folder)):
        return {"success": False, "message": f"Folder {folder_name} already exists"}
    
    # Create folder
    parent = sp_context.web.get_folder_by_server_relative_url(parent_path)
    new_folder = parent.folders.add(folder_name)
    sp_context.execute_query()
    
    return _file_success_response(new_folder, f"Folder {folder_name} created successfully")

@mcp.tool(name="Upload_Document", description="Upload a new file to a SharePoint directory")
@_handle_sp_operation
async def upload_document(folder_name: str, file_name: str, content: str, is_base64: bool = False):
    """Upload a new file to a directory"""
    logger.info(f"Uploading document {file_name} to folder {folder_name}")
    
    # Convert content and upload
    file_content = base64.b64decode(content) if is_base64 else content.encode('utf-8')
    folder = sp_context.web.get_folder_by_server_relative_url(_get_path(folder_name))
    uploaded_file = folder.upload_file(file_name, file_content)
    sp_context.execute_query()
    
    return _file_success_response(uploaded_file, f"File {file_name} uploaded successfully")

@mcp.tool(name="Upload_Document_From_Path", description="Upload a file directly from a file path to SharePoint")
@_handle_sp_operation
async def upload_document_from_path(folder_name: str, file_path: str, new_file_name: Optional[str] = None):
    """Upload a file directly from a path without needing to convert to base64 first"""
    logger.info(f"Uploading document from path {file_path} to folder {folder_name}")
    
    try:
        with open(file_path, "rb") as file:
            file_content = file.read()
        
        if not new_file_name:
            new_file_name = os.path.basename(file_path)
            
        folder = sp_context.web.get_folder_by_server_relative_url(_get_path(folder_name))
        uploaded_file = folder.upload_file(new_file_name, file_content)
        sp_context.execute_query()
        
        return _file_success_response(uploaded_file, f"File {new_file_name} uploaded successfully")
    except Exception as e:
        logger.error(f"Error uploading file from path: {str(e)}")
        raise

@mcp.tool(name="Update_Document", description="Update an existing document in a SharePoint directory")
@_handle_sp_operation
async def update_document(folder_name: str, file_name: str, content: str, is_base64: bool = False):
    """Update an existing document in a SharePoint directory"""
    logger.info(f"Updating document {file_name} in folder {folder_name}")
    
    # Check if file exists
    file_path = _get_path(folder_name, file_name)
    file = sp_context.web.get_file_by_server_relative_url(file_path)
    sp_context.load(file, ["Exists", "Name", "ServerRelativeUrl"])
    sp_context.execute_query()
    
    if not file.exists:
        return {"success": False, "message": f"File {file_name} does not exist in folder {folder_name}"}
    
    # Update file using upload method
    file_content = base64.b64decode(content) if is_base64 else content.encode('utf-8')
    folder = sp_context.web.get_folder_by_server_relative_url(_get_path(folder_name))
    updated_file = folder.upload_file(file_name, file_content)
    sp_context.execute_query()
    
    return _file_success_response(updated_file, f"File {file_name} updated successfully")

@mcp.tool(name="Delete_Document", description="Delete a document from a SharePoint directory")
@_handle_sp_operation
async def delete_document(folder_name: str, file_name: str):
    """Delete a document from a directory"""
    logger.info(f"Deleting document {file_name} from folder {folder_name}")
    
    # Check if file exists and delete
    file = sp_context.web.get_file_by_server_relative_url(_get_path(folder_name, file_name))
    sp_context.load(file, ["Exists"])
    sp_context.execute_query()
    
    if not file.exists:
        return {"success": False, "message": f"File {file_name} does not exist in folder {folder_name}"}
    
    file.delete_object()
    sp_context.execute_query()
    return {"success": True, "message": f"File {file_name} deleted successfully"}

@mcp.tool(name="Delete_Folder", description="Delete an empty folder from SharePoint")
@_handle_sp_operation
async def delete_folder(folder_path: str):
    """Delete an empty folder from SharePoint"""
    logger.info(f"Deleting folder: {folder_path}")
    
    # Get folder and check if it exists and is empty
    full_path = _get_path(folder_path)
    folder = sp_context.web.get_folder_by_server_relative_url(full_path)
    sp_context.load(folder)
    sp_context.load(folder.files)
    sp_context.load(folder.folders)
    sp_context.execute_query()
    
    if not hasattr(folder, 'exists') or not folder.exists:
        return {"success": False, "message": f"Folder '{folder_path}' does not exist"}
    
    if len(folder.files) > 0:
        return {"success": False, "message": f"Folder contains {len(folder.files)} files"}
    
    if len(folder.folders) > 0:
        return {"success": False, "message": f"Folder contains {len(folder.folders)} subfolders"}
    
    # Delete the empty folder
    folder.delete_object()
    sp_context.execute_query()
    return {"success": True, "message": f"Folder '{folder_path}' deleted successfully"}

@mcp.tool(name="Download_Document", description="Download a document from SharePoint to local filesystem")
@_handle_sp_operation
async def download_document_tool(folder_name: str, file_name: str, local_path: str):
    """Download a document from SharePoint to local filesystem with fallback support"""
    return download_document(folder_name, file_name, local_path)

@mcp.tool(name="Update_File_Metadata", description="Update metadata fields for a SharePoint file")
@_handle_sp_operation
async def update_file_metadata(folder_name: str, file_name: str, metadata: dict):
    """Update metadata fields for a SharePoint file"""
    file_path = _get_path(folder_name, file_name)
    file_obj = sp_context.web.get_file_by_server_relative_url(file_path)
    list_item = file_obj.listItemAllFields
    
    # Convert values to strings
    form_values = {}
    for key, value in metadata.items():
        if value is not None:
            if isinstance(value, bool):
                form_values[key] = "1" if value else "0"
            elif isinstance(value, list):
                form_values[key] = ";".join(str(v) for v in value)
            else:
                form_values[key] = str(value)
    
    if not form_values:
        return {"success": True, "message": "No fields to update"}
    
    list_item.validate_update_list_item(form_values, new_document_update=True)
    sp_context.execute_query()
    
    return {"success": True, "message": f"Updated {len(form_values)} fields for {file_name}"}

@mcp.tool(name="Get_File_Metadata", description="Get metadata fields from a SharePoint file")
@_handle_sp_operation
async def get_file_metadata(folder_name: str, file_name: str):
    """Get metadata fields from a SharePoint file"""
    file_path = _get_path(folder_name, file_name)
    
    # Get the file
    file_obj = sp_context.web.get_file_by_server_relative_url(file_path)
    
    # Access list item directly
    list_item = file_obj.listItemAllFields
    sp_context.load(list_item)
    sp_context.execute_query()
    
    # Extract metadata
    metadata = {}
    for prop_name, prop_value in list_item.properties.items():
        if prop_value is not None:
            metadata[prop_name] = str(prop_value)
    
    return {
        "success": True,
        "message": f"Metadata retrieved for {file_name}",
        "metadata": metadata,
        "file": {"name": file_name, "path": file_path}
    }
