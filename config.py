# gme_compliance_pipeline/config.py
import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env once
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)

# Folder paths
FOLDER_PATH_gme_compliance = os.getenv("FOLDER_PATH_gme_compliance")
if not FOLDER_PATH_gme_compliance:
    raise ValueError("FOLDER_PATH_gme_compliance not found in .env file")
