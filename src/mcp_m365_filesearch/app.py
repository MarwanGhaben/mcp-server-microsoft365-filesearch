import uvicorn
from server import mcp

if __name__ == "__main__":
    uvicorn.run(mcp.app, host="0.0.0.0", port=10000)
