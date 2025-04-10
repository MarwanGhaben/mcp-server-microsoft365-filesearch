# Microsoft 365 File Search MCP Server (SharePoint & OneDrive)

## Overview

A Model Context Protocol (MCP) server implementation that provides advanced file search capabilities within Microsoft 365. This server enables efficient file discovery, metadata analysis, and integration with business workflows by making available the content from SharePoint/OneDrive.



https://github.com/user-attachments/assets/bbe63c02-f6d9-4c9b-8f98-36fc22a081cc



## Components

### Tools

The server offers 2 core tools:

- **search_m365_files**  
  Perform a file search within the M365 environment.  
  **Input**:
  - `query` (string): The search term or criteria.  
  **Returns**: Array of file metadata objects. Metadata includes the file content summary, *drive ID*, and *file ID*, among other details.

- **get_file_content**  
  Retrieve content from a specific file.  
  **Input**:
  - `driveid` (string): The unique identifier of the parent drive.  
  - `fileid` (string): The unique identifier of the file.  
  **Returns**: File content as a binary stream.  
  **Note**: Uses a local cache to speed up repeat access.

### Caching

To improve performance and reduce redundant API calls, the server caches downloaded files locally. This is particularly useful when working with large documents or frequently accessed files.

- Cached files are stored in the `./src/mcp_m365_filesearch/.local/downloads` directory (relative to the project root).
- When a file is requested via `get_file_content`, the server first checks the cache.
- If the file is already cached, it is returned directly from disk without a new API call.

This feature ensures faster response times and efficient use of API rate limits.

## Usage with Claude Desktop

To integrate the server with Claude Desktop, update your `claude_desktop_config.json`:

```json
"mcpServers": {
  "M365 File Search (SharePoint/OneDrive)": {
    "command": "uv",
    "args": [
      "--directory",
      "full_path_to_parent_directory",
      "run",
      ".\\src\\mcp_m365_filesearch\\server.py"
    ],
    "env": {
      "CLIENT_ID": "MSGraph Client ID",
      "CLIENT_SECRET": "MS Graph Client Secret",
      "TENANT_ID": "TENANT ID",
      "REGION": "SEARCH REGION"
    }
  }
}
