import os
from pathlib import Path

from pydantic import BaseModel

# APP SETTINGS
APP_NAME = os.getenv("APP_NAME", "corpus-collector-backend")
APP_VERSION = os.getenv("APP_VERSION", "1.0.0")

# LOGGING SETTINGS
LOGGER_LEVEL = os.getenv("LOG_LEVEL", "INFO")

# BLOB STORAGE SETTINGS
BLOB_URL = os.getenv('BLOB_URL', 'https://eviestrgstrdev01.blob.core.windows.net')
CONTAINER_NAME = os.getenv("AZURE_STORAGE_CONTAINER_NAME", "corpus-collector-data")


class Configuration(BaseModel):
    app_name: str = APP_NAME
    app_version: str = APP_VERSION
    logger_level: str = LOGGER_LEVEL
    blob_url: str = BLOB_URL
    container_name: str = CONTAINER_NAME


app_settings = Configuration()
