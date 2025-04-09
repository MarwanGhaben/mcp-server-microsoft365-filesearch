from msal import ConfidentialClientApplication
import os
from logger_config import setup_logger

# Initialize logger
logger = setup_logger()

# ----------------------
# Configuration
# ----------------------
CLIENT_ID = os.getenv("CLIENT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")
TENANT_ID = os.getenv("TENANT_ID", "")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# ----------------------
# Authentication
# ----------------------
def get_token_client_credentials():
    logger.info("Acquiring token using client credentials...")
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        logger.info("Token acquired successfully.")
    else:
        logger.error(f"Failed to acquire token: {result.get('error_description', 'Unknown error')}")
    return result.get("access_token")
