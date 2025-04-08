import logging
import os

from azure.monitor.opentelemetry import configure_azure_monitor
from opentelemetry import trace
from opentelemetry.sdk.resources import Resource

from app.lib.configurations import app_settings


class AppLogger(object):
    _LOGGER = None
    _TRACER = None

    @classmethod
    def init_logger_and_tracer(cls):
        # Azure is very chatty - limit those logs
        azure_http_logger = logging.getLogger(
            "azure.core.pipeline.policies.http_logging_policy"
        )
        azure_http_logger.setLevel(logging.WARNING)
        azure_logger = logging.getLogger("azure")
        azure_logger.setLevel(logging.WARNING)

        # limit snowflake logs
        snowflake_connector = logging.getLogger("snowflake.connector")
        snowflake_connector.setLevel(logging.WARNING)

        # limit urllib3 logs
        urllib3_logger = logging.getLogger("urllib3")
        urllib3_logger.setLevel(logging.WARNING)

        logger = logging.getLogger(app_settings.app_name)
        logger.setLevel(app_settings.logger_level)

        if os.getenv("APPLICATIONINSIGHTS_CONNECTION_STRING"):
            configure_azure_monitor(
                logger_name=app_settings.app_name,
                resource=Resource(
                    attributes={
                        "service.version": app_settings.app_version,
                        "service.name": app_settings.app_name,
                    }
                ),
            )
        else:
            formatter = logging.Formatter(
                "[%(asctime)s] {%(name)s} %(levelname)s - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
            lh = logging.StreamHandler()
            lh.setFormatter(formatter)
            logger.addHandler(lh)

        cls._LOGGER = logger
        cls._TRACER = trace.get_tracer(app_settings.app_name)

    @classmethod
    def get_logger_and_tracer(cls):
        if cls._LOGGER is None or cls._TRACER is None:
            cls.init_logger_and_tracer()
        return cls._LOGGER, cls._TRACER
