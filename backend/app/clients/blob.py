from azure.storage.blob import BlobServiceClient
from typing import IO
from azure.identity import DefaultAzureCredential
from pathlib import Path

from app.lib.configurations import app_settings
from app.lib.logging_helpers import AppLogger

logger, tracer = AppLogger.get_logger_and_tracer()

class BlobClient:
    def __init__(self):
        self.service_client = BlobServiceClient(account_url=app_settings.blob_url, credential=DefaultAzureCredential())

    def fetch_blob(
        self,
        file_name: str,
        container_name: str,
    ) -> bytes:
        blob_client = self.service_client.get_blob_client(
            container=container_name,
            blob=file_name,
        )
        logger.debug(f'Fetching {blob_client.url}')
        blob = blob_client.download_blob()

        return blob.readall()

    def save_to_blob(
        self,
        data: bytes | str | IO[bytes],
        file_name: str,
        container_name: str,
    ):
        blob_client = self.service_client.get_blob_client(
            container=container_name,
            blob=file_name,
        )
        blob_client.upload_blob(data, overwrite=True)
        file_extension = Path(file_name).suffix.replace('.', '').upper()
        logger.info(f'{file_extension} saved to {blob_client.url}')