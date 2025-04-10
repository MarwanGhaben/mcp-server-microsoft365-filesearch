from logger_config import setup_logger
from typing import Literal
import os
from mcp.server.fastmcp import FastMCP
import mcp.types as types
from pydantic import AnyUrl
from msal_auth import get_token_client_credentials
from msgraph_util import search_graph, parse_search_response, download_file

# Initialize logger
logger = setup_logger()

# Create MCP server
mcp = FastMCP("M365 File Search (SharePoint/OneDrive)")

# ----------------------
# Configuration
# ----------------------
VALID_REGIONS = {"NAM", "EUR", "APC", "AUS", "IND"}
REGION = os.getenv("REGION", "NAM").upper()
if REGION not in VALID_REGIONS:
    logger.warning(f"Invalid REGION '{REGION}' specified. Defaulting to 'NAM'.")
    REGION = "NAM"

logger.info("Server started successfully.")

# ----------------------
# MCP Resource
# ----------------------

#@mcp.resource("microsoft365://{driveid}/{fileid}", name="Get File Content", description="Get content of a Microsoft 365 file by drive id and file id.")
@mcp.tool()
async def get_file_content(driveid: str, fileid: str) -> str:
    """
    Get content of Microsoft 365 file by drive id and file id.
    """
    access_token = get_token_client_credentials()
    if not access_token:
        logger.error("Failed to acquire access token.")
        return {"error": "Authentication failed."}
    return await download_file(driveid, fileid, access_token) 

# @mcp.list_resources()
# async def handle_list_resources() -> list[types.Resource]:
#     logger.debug("Handling list_resources request")
#     return [
#         types.Resource(
#             uri=AnyUrl("microsoft365://{driveid}/{fileid}"),
#             name="Get File Content",
#             description="Get content of a Microsoft 365 file by drive id and file id.",
#             mimeType="text/plain",
#         )
#     ]

# @mcp.read_resource()
# async def handle_read_resource(uri: AnyUrl) -> str:
#     logger.debug(f"Handling read_resource request for URI: {uri}")
#     if uri.scheme != "microsoft365":
#         logger.error(f"Unsupported URI scheme: {uri.scheme}")
#         raise ValueError(f"Unsupported URI scheme: {uri.scheme}")

#     path = str(uri).replace("microsoft365://", "")
#     if not path or path.split("/").count != 2:
#         logger.error(f"Unknown resource path: {path}")
#         raise ValueError(f"Unknown resource path: {path}")
#     uri_parts = path.split("/")
#     driveid = uri_parts[0]  
#     fileid = uri_parts[1]
#     logger.debug(f"Extracted driveid: {driveid}, fileid: {fileid}")
#     return get_file_content(driveid, fileid)
# ----------------------
# MCP Tool
# ----------------------

@mcp.tool()
async def search_m365_files(
    query: str,
    file_type: Literal["all", "document", "spreadsheet", "presentation", "image"] = "all",
    max_results: int = 10,
) -> dict:
    """
    Search Microsoft 365 files by query and file type. If you need content of a file, you can make use of resource microsoft365://{driveid}/{fileid}
    """
    logger.info(f"Tool invoked: search_m365_files with query='{query}', file_type='{file_type}', max_results={max_results}")

    if not query.strip():
        logger.warning("Empty query received.")
        return {"count": 0, "files": [], "message": "Please provide a valid search query."}

    file_types = {
        "all": None,
        "document": ["docx", "doc", "txt", "pdf"],
        "spreadsheet": ["xlsx", "xls"],
        "presentation": ["pptx"],
        "image": ["jpg", "png"],
    }
    file_extension = file_types[file_type]

    access_token = get_token_client_credentials()
    if not access_token:
        return {"count": 0, "files": [], "message": "Authentication failed."}

    search_results = search_graph(query, access_token, REGION, size=max_results, from_index=0)
    if not search_results:
        return {"count": 0, "files": [], "message": "No results found."}

    results = parse_search_response(search_results, file_type, file_extension)
    logger.info(f"Returning {len(results)} results.")
    return {"count": len(results), "files": results}

# ----------------------
# Run Server
# ----------------------
if __name__ == "__main__":
    logger.info("Starting MCP server...")
    mcp.run(transport="stdio") 
