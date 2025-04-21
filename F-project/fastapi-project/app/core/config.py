import os
from dotenv import load_dotenv

load_dotenv()

class Settings:
    PROJECT_NAME: str = os.getenv("PROJECT_NAME", "FastAPI Settlement Project")
    VERSION: str = os.getenv("VERSION", "1.0.0")
    DESCRIPTION: str = os.getenv("DESCRIPTION", "A FastAPI project for generating settlement files.")
    INPUT_FILE_PATH: str = os.getenv("INPUT_FILE_PATH", "/path/to/input/file.xlsx")
    OUTPUT_FOLDER_PATH: str = os.getenv("OUTPUT_FOLDER_PATH", "/path/to/output/folder")
    SPECIFIED_DATE: str = os.getenv("SPECIFIED_DATE", None)

settings = Settings()