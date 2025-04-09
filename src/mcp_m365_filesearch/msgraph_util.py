import os
import requests
import time
import json  # For saving and loading the documents object as JSON
from llama_index.core import SimpleDirectoryReader
from logger_config import setup_logger

# Initialize logger
logger = setup_logger()

# Microsoft Graph API URL
GRAPH_URL = "https://graph.microsoft.com/v1.0"  

# ----------------------
# Graph Search
# ----------------------
def search_graph(query_text, access_token, region, size=20, from_index=0):
    logger.info(f"Searching Microsoft Graph for query: {query_text} (from={from_index}, size={size})")
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    url = f"{GRAPH_URL}/search/query"
    body = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": query_text},
                "fields": [
                    "name", "webUrl", "id", "parentReference",
                    "createdBy", "createdDateTime", "lastModifiedBy", "lastModifiedDateTime"
                ],
                "from": from_index,
                "size": size,
                "region": region
            }
        ]
    }
    response = requests.post(url, headers=headers, json=body)
    if response.status_code == 200:
        logger.info("Search completed successfully.")
        return response.json()
    else:
        logger.error(f"Search failed: {response.status_code} - {response.text}")
        return None

# ----------------------
# Response Parser
# ----------------------
def parse_search_response(search_results, file_type, file_extension):
    results = []
    hits_containers = search_results.get("value", [])
    if hits_containers and isinstance(hits_containers, list):
        hits = hits_containers[0].get("hitsContainers", [])
        if hits and isinstance(hits, list):
            for result in hits[0].get("hits", []):
                resource = result.get("resource", {})
                az_search_rank = result.get("rank")
                summary = result.get("summary", "")
                file_name = resource.get("name", "")
                file_url = resource.get("webUrl")

                if file_name and (
                    file_type == "all" or any(file_name.endswith(f".{ext}") for ext in file_extension)
                ):
                    results.append({
                        "name": file_name,
                        "url": file_url,
                        "summary": summary,
                        "rank": az_search_rank,
                        "source": classify_source(file_url),
                        "created_by": resource.get("createdBy", {}).get("user", {}),
                        "created_date": resource.get("createdDateTime"),
                        "last_modified_by": resource.get("lastModifiedBy", {}).get("user", {}),
                        "last_modified_date": resource.get("lastModifiedDateTime"),
                        "fileid": resource.get("id"),
                        "parent_reference": resource.get("parentReference", {}),
                        "drive_id": resource.get("parentReference", {}).get("driveId"),
                    })
    return results

# ----------------------
# Download Helpers
# ----------------------
def classify_source(web_url):
    logger.debug(f"Classifying source for URL: {web_url}")
    if "-my.sharepoint.com" in web_url or "/personal/" in web_url:
        return "OneDrive"
    return "SharePoint"

def download_file(drive_id, item_id, access_token):
    logger.info(f"Downloading file with ID: {item_id} from drive: {drive_id}")
    headers = {"Authorization": f"Bearer {access_token}"}
    metadata_url = f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}"

    # Fetch metadata to get the file name
    metadata_response = requests.get(metadata_url, headers=headers)
    if metadata_response.status_code == 200:
        metadata = metadata_response.json()
        file_name = metadata.get("name", f"{item_id}.bin")
    else:
        logger.error(f"Failed to fetch metadata: {metadata_response.status_code} - {metadata_response.text}")
        return None
    current_dir = os.path.dirname(os.path.abspath(__file__))
    local_dir = os.path.join(current_dir, ".local")
    item_folder = os.path.join(local_dir, "downloads", drive_id, item_id)
    os.makedirs(item_folder, exist_ok=True)

    # Check if the folder contains a file
    existing_files = os.listdir(item_folder)
    if existing_files:
        existing_file_path = os.path.join(item_folder, existing_files[0])
        file_age = time.time() - os.path.getmtime(existing_file_path)
        if file_age < 24 * 3600:  # File is less than 24 hours old
            logger.info(f"Using existing file: {existing_file_path}")
            return _read_file_content(existing_file_path)
        else:
            logger.info(f"Deleting old file: {existing_file_path}")
            os.remove(existing_file_path)

    # Full file path for the new file
    file_path = os.path.join(item_folder, file_name)

    # Download the file
    content_url = f"{metadata_url}/content"
    response = requests.get(content_url, headers=headers, stream=True)
    if response.status_code == 200:
        with open(file_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        logger.info(f"File downloaded: {file_path}")
        return _read_file_content(file_path)
    else:
        logger.error(f"Download failed: {response.status_code} - {response.text}")
        return None

def _read_file_content(file_path):
    """
    Read the content of the file using llamaindex SimpleDirectoryReader.
    Cache the processed documents object to a JSON file for reuse.
    """
    try:
        # Define the path for the cached documents object
        cache_file_path = f"{file_path}.cache.json"

        # Check if the cache file exists and is not older than 24 hours
        if os.path.exists(cache_file_path):
            cache_age = time.time() - os.path.getmtime(cache_file_path)
            if cache_age < 24 * 3600:  # Cache is less than 24 hours old
                logger.info(f"Using cached documents from: {cache_file_path}")
                with open(cache_file_path, "r", encoding="utf-8") as cache_file:
                    documents = json.load(cache_file)
                return documents  # Return the entire documents object

        # Process the file content if no valid cache exists
        reader = SimpleDirectoryReader(input_files=[file_path])
        documents = reader.load_data()
        logger.info("File processed using SimpleDirectoryReader.")

        if documents:
            # Convert documents to a serializable format
            serializable_documents = [{"text": doc.text, **doc.metadata} for doc in documents]

            # Save the documents object to the cache file as JSON
            with open(cache_file_path, "w", encoding="utf-8") as cache_file:
                json.dump(serializable_documents, cache_file, ensure_ascii=False, indent=4)
            logger.debug(f"Processed documents cached at: {cache_file_path}")
            return serializable_documents  # Return the entire documents object
        else:
            logger.warning("No documents found in the file.")
            return None
    except Exception as e:
        logger.error(f"Failed to process file using SimpleDirectoryReader: {e}")
        return None

