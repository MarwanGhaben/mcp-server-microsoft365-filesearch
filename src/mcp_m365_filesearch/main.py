from fastapi import FastAPI, Query
from typing import Literal
from msgraph_util import search_graph, parse_search_response
from msal_auth import get_token_client_credentials
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

    search_results = search_graph(query, access_token, REGION, size=max_results, from_index=0)
    if not search_results:
        return {"count": 0, "files": [], "message": "No results found."}

    results = parse_search_response(search_results, file_type, file_extension)
    return {"count": len(results), "files": results}
