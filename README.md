# SharePoint MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A lightweight MCP Server for seamless integration with Microsoft SharePoint, enabling MCP clients to interact with documents, folders and other SharePoint resources. Developed by [sofias tech](https://github.com/sofias/mcp-sharepoint/).

<a href="https://glama.ai/mcp/servers/@Sofias-ai/mcp-sharepoint">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@Sofias-ai/mcp-sharepoint/badge" alt="SharePoint Server MCP server" />
</a>

## Features

This server provides a clean interface to SharePoint resources through the Model Context Protocol (MCP), with optimized operations for document management.

### Tools

The server implements the following tools:

- `List_SharePoint_Folders`: Lists all folders in a specified directory or root
- `List_SharePoint_Documents`: Fetches all documents within a specified folder
- `Get_Document_Content`: Retrieves the content of a document (as text or base64-encoded binary)
- `Create_Folder`: Creates a new folder in the specified directory or root
- `Upload_Document`: Uploads a new document to a specified folder
- `Upload_Document`: Uploads large documents from path.
- `Update_Document`: Updates the content of an existing document
- `Delete_Document`: Removes a document from a specified folder
- `Delete_Folder`: Deletes an empty folder from SharePoint

## Architecture

The server is built with resource efficiency in mind:

- Efficient SharePoint API usage with selective property loading
- Error handling through decorators for cleaner code
- Clear separation between resource management and tool implementation
- Optimized content handling for both text and binary files

## Setup

1. Register an app in Azure AD with appropriate SharePoint permissions
2. Obtain the client ID and client secret for the registered app
3. Identify your SharePoint site URL and the document library path you want to work with

## Environment Variables

The server requires these environment variables:

- `SHP_ID_APP`: Your Azure AD application client ID
- `SHP_ID_APP_SECRET`: Your Azure AD application client secret
- `SHP_SITE_URL`: The URL of your SharePoint site
- `SHP_DOC_LIBRARY`: Path to the document library (default: "Shared Documents/mcp_server")
- `SHP_TENANT_ID`: Your Microsoft tenant ID

## Quickstart

### Installation

```bash
pip install -e .
```

Or install from PyPI once published:

```bash
pip install mcp-sharepoint-server
```

Using uv:

```bash
uv pip install mcp-sharepoint-server
```

### Claude Desktop Integration

To integrate with Claude Desktop, update the configuration file:

On Windows: `%APPDATA%/Claude/claude_desktop_config.json`
On macOS: `~/Library/Application\ Support/Claude/claude_desktop_config.json`

#### Standard Integration

```json
"mcpServers": {
  "sharepoint": {
    "command": "mcp-sharepoint",
    "env": {
      "SHP_ID_APP": "your-app-id",
      "SHP_ID_APP_SECRET": "your-app-secret",
      "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
      "SHP_DOC_LIBRARY": "Shared Documents/your-folder",
      "SHP_TENANT_ID": "your-tenant-id"
    }
  }
}
```

#### Using uvx

```json
"mcpServers": {
  "sharepoint": {
    "command": "uvx",
    "args": [
      "mcp-sharepoint"
    ],
    "env": {
      "SHP_ID_APP": "your-app-id",
      "SHP_ID_APP_SECRET": "your-app-secret",
      "SHP_SITE_URL": "https://your-tenant.sharepoint.com/sites/your-site",
      "SHP_DOC_LIBRARY": "Shared Documents/your-folder",
      "SHP_TENANT_ID": "your-tenant-id"
    }
  }
}
```

## Development

### Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt` and `pyproject.toml`

### Local Development

1. Clone the repository
2. Create a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```
3. Install development dependencies:
   ```bash
   pip install -e .
   ```
4. Create a `.env` file with your SharePoint credentials:
   ```
   SHP_ID_APP=your-app-id
   SHP_ID_APP_SECRET=your-app-secret
   SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
   SHP_DOC_LIBRARY=Shared Documents/your-folder
   SHP_TENANT_ID=your-tenant-id
   ```
5. Run the server:
   ```bash
   python -m mcp_sharepoint
   ```

### Debugging

For debugging the MCP server, you can use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector):

```bash
npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

Copyright (c) 2025 sofias tech