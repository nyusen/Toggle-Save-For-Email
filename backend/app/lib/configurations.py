import os
from pathlib import Path

from pydantic import BaseModel

# APP SETTINGS
APP_NAME = os.getenv("APP_NAME", "corpus-collector-backend")
APP_VERSION = os.getenv("APP_VERSION", "1.0.0")
ENV = os.getenv("ENV", "dev")

# LOGGING SETTINGS
LOGGER_LEVEL = os.getenv("LOG_LEVEL", "INFO")

# BLOB STORAGE SETTINGS
BLOB_URL = os.getenv('BLOB_URL', 'https://eviestrgstrdev01.blob.core.windows.net')
CONTAINER_NAME = os.getenv("AZURE_STORAGE_CONTAINER_NAME", "corpus-collector-data")

class SQLServerSettings(BaseModel):
    odbc_version: str = os.getenv('ODBC_VERSION','18')
    query_parameter_limit: int = 2000  #Maximum query parameter limit is 2100
    pool_min_size: int = 20 if ENV == 'prd' else 5
    pool_max_size: int = 50 if ENV == 'prd' else 20
    pool_recycle: int = 480  #If set to a value other than -1, number of seconds between connection recycling, which means upon checkout, if this timeout is surpassed the connection will be closed and replaced with a newly opened connection. EXEC sp_configure 'remote query timeout' can be run to see default connection timeout in sqlserver
    echo: bool = False # if True, the connection pool will log informational output such as when connections are invalidated as well as when connections are recycled to the default log handler, which defaults to sys.stdout for output.. If set to the string "debug", the logging will include pool checkouts and checkins.
    autocommit: bool = False

class Configuration(BaseModel):
    app_name: str = APP_NAME
    app_version: str = APP_VERSION
    env: str = ENV
    logger_level: str = LOGGER_LEVEL
    blob_url: str = BLOB_URL
    container_name: str = CONTAINER_NAME
    sql_server_settings: SQLServerSettings = SQLServerSettings()


app_settings = Configuration()
