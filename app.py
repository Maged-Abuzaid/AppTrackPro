import logging
import os
from config.settings_manager import base_path
from src.gui.main_window import AppTrackPro
# Configure logging
log_file_path = os.path.join(base_path, "apptrackpro.log")
os.makedirs(os.path.dirname(log_file_path), exist_ok=True)  # Ensure the directory exists

logging.basicConfig(
    filename=log_file_path,
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)
logging.debug("Application is starting.")

if __name__ == "__main__":
    try:
        app = AppTrackPro()
        logging.debug("AppTrackPro initialized successfully.")
        app.mainloop()
    except Exception as e:
        logging.error(f"Application failed to start: {str(e)}")
