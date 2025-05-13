from server import mcp
import logging

if __name__ == "__main__":
    logging.info("Running MCP via app.py...")
    mcp.run(host="0.0.0.0", port=10000)
