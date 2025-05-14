from fastapi import FastAPI, Query
from typing import Literal
from msgraph_util import search_graph, parse_search_response, download_file
from msal_auth import get_token_client_credentials
from fastapi.responses import JSONResponse
import os
import logging

app = FastAPI()
logging.basicConfig(level=logging.INFO)

VALID_REGIONS = {"NAM", "EUR", "APC", "AUS", "IND", "CAN"}
REGION = os.getenv("REGION", "NAM").upper()
print(f"⚠️ DEBUG: REGION detected: {REGION}")
if REGION not in VALID_REGIONS:
    REGION = "NAM"

@app.get("/search")
async def search_m365_files(
    query: str = Query(..., description="Search query"),
    file_type: Literal["all", "document", "spreadsheet", "presentation", "image"] = "all",
    max_results: int = 10
):
    access_token = get_token_client_credentials()
    if not access_token:
        return {"count": 0, "files": [], "message": "Authentication failed."}

    file_types = {
        "all": None,
        "document": ["docx", "doc", "txt", "pdf"],
        "spreadsheet": ["xlsx", "xls"],
        "presentation": ["pptx"],
        "image": ["jpg", "png"],
    }
    file_extension = file_types[file_type]

    # ✅ Add fallback keyword if query is empty
    if not query.strip():
        if file_type == "spreadsheet":
            query = "xlsx"
        elif file_type == "document":
            query = "docx"
        elif file_type == "presentation":
            query = "pptx"
        elif file_type == "image":
            query = "jpg"
        else:
            query = "file"

    search_results = search_graph(query, access_token, REGION, size=max_results, from_index=0)
    if not search_results:
        return {"count": 0, "files": [], "message": "No results found."}

    results = parse_search_response(search_results, file_type, file_extension)
    return {"count": len(results), "files": results}

@app.get("/get_file_content")
async def get_file_content(driveid: str, fileid: str):
    access_token = get_token_client_credentials()
    if not access_token:
        return JSONResponse(status_code=401, content={"error": "Failed to authenticate with Microsoft Graph."})

    try:
        content = await download_file(driveid, fileid, access_token)
        return {"content": content}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
