from fastapi import FastAPI, Query
from fastapi.responses import JSONResponse
from typing import Literal
import os
import logging

from msgraph_util import (
    search_graph,
    parse_search_response,
    download_file,
    crawl_drive_items
)
from msal_auth import get_token_client_credentials

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

    search_results = search_graph(query, access_token, REGION, size=max_results, from_index=0)
    if not search_results:
        return {"count": 0, "files": [], "message": "No results found."}

    results = parse_search_response(search_results, file_type, file_extension)
    return {"count": len(results), "files": results}

@app.get("/get_file_content")
async def get_file_content(driveid: str, fileid: str, offset: int = 0, limit: int = 50):
    access_token = get_token_client_credentials()
    if not access_token:
        return JSONResponse(status_code=401, content={"error": "Failed to authenticate with Microsoft Graph."})

    try:
        content = await download_file(driveid, fileid, access_token, offset=offset, limit=limit)
        if isinstance(content, list) and len(content) > 0 and "text" in content[0]:
            return {"content": content[0]["text"]}
        elif isinstance(content, str):
            return {"content": content}
        else:
            return {"content": "[No readable content found]"}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.get("/crawl")
def crawl_files(driveid: str, file_extension: str = None):
    access_token = get_token_client_credentials()
    if not access_token:
        return JSONResponse(status_code=401, content={"error": "Authentication failed."})

    try:
        files = crawl_drive_items(access_token, driveid, file_extension=file_extension)
        return {"count": len(files), "files": files}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

from .msgraph_util import SITE_NAME_TO_ID

@app.get("/search_site")
async def search_files_in_site(site_name: str, query: str, max_results: int = 10):
    from .msgraph_util import SITE_NAME_TO_ID
    access_token = get_token_client_credentials()
    print(f"site_name: {site_name}, access_token: {bool(access_token)}")
    if not access_token:
        print("No access token!")
        return {"count": 0, "files": [], "message": "Authentication failed."}
    site_id = SITE_NAME_TO_ID.get(site_name)
    print(f"Resolved site_id: {site_id}")
    if not site_id:
        print(f"Unknown site name: {site_name}")
        return {"count": 0, "files": [], "message": f"Unknown site name: {site_name}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{query}')"
    print(f"Graph URL: {url}")
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    print(f"Graph response: {response.status_code} {response.text}")
    if response.status_code != 200:
        return {"count": 0, "files": [], "message": f"Search failed: {response.status_code} {response.text}"}
    data = response.json()
    files = [
        {
            "name": item.get("name"),
            "id": item.get("id"),
            "webUrl": item.get("webUrl")
        }
        for item in data.get("value", [])
    ]
    return {"count": len(files), "files": files}
