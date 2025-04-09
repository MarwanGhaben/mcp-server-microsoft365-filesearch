import logging
import os

# ----------------------
# Logger Configuration
# ----------------------
def setup_logger():
    # Get the directory of the current file
    current_dir = os.path.dirname(os.path.abspath(__file__))
    local_dir = os.path.join(current_dir, ".local")
    os.makedirs(local_dir, exist_ok=True)  # Ensure the .local folder exists
    log_file_path = os.path.join(local_dir, "server.log")

    # Configure logging
    logging.basicConfig(
        level=logging.DEBUG,  # Set the logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.StreamHandler(),  # Output logs to the console
            logging.FileHandler(log_file_path, mode="a", encoding="utf-8")  # Write logs to a file
        ]
    )
    return logging.getLogger(__name__)
