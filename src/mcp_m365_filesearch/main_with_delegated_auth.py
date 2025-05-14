from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, JSONResponse
from starlette.middleware.sessions import SessionMiddleware
from starlette.responses import HTMLResponse
import os
import msal
import requests

from msgraph_util import (
    search_graph,
    parse_search_response,
    download_file,
    crawl_drive_items
)
from msal_auth import get_token_client_credentials

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=os.urandom(24))

# ENV config for delegated auth
CLIENT_ID = os.getenv("DELEGATED_CLIENT_ID")
CLIENT_SECRET = os.getenv("DELEGATED_CLIENT_SECRET")
TENANT_ID = os.getenv("DELEGATED_TENANT_ID")
REDIRECT_URI = os.getenv("DELEGATED_REDIRECT_URI", "http://localhost:8000/auth/callback")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["User.Read", "Files.Read.All"]

@app.get("/")
def home(request: Request):
    if request.session.get("user"):
        return HTMLResponse(f"""
            <h3>Welcome {request.session['user'].get('name')}</h3>
            <a href='/me/files'>View My OneDrive</a>
        """)
    return HTMLResponse("<a href='/auth/login'>Sign in with Microsoft</a>")

@app.get("/auth/login")
def auth_login(request: Request):
    flow = _build_auth_code_flow()
    request.session["flow"] = flow
    return RedirectResponse(flow["auth_uri"])

@app.get("/auth/callback")
def auth_callback(request: Request):
    flow = request.session.get("flow")
    result = _build_msal_app().acquire_token_by_auth_code_flow(flow, dict(request.query_params))

    if "error" in result:
        return JSONResponse({"error": result.get("error_description")}, status_code=400)

    request.session["user"] = result.get("id_token_claims")
    request.session["access_token"] = result.get("access_token")
    return RedirectResponse("/")

@app.get("/me/files")
def list_my_files(request: Request):
    access_token = request.session.get("access_token")
    if not access_token:
        return RedirectResponse("/auth/login")

    headers = {"Authorization": f"Bearer {access_token}"}
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()
    else:
        return JSONResponse({"error": response.text}, status_code=response.status_code)

# ------------------
# MSAL helper funcs
# ------------------

def _build_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

def _build_auth_code_flow():
    return _build_msal_app().initiate_auth_code_flow(
        scopes=SCOPE,
        redirect_uri=REDIRECT_URI
    )
