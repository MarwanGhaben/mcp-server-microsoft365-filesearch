from src.mcp_m365_filesearch.server import app  # adjust if actual app object is in another file

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
