from contextlib import asynccontextmanager
from itertools import chain
from typing import Any
import uuid

from azure.identity import DefaultAzureCredential
import pandas as pd
from aioodbc import create_pool, Connection, Cursor
import struct as s

from app.lib.configurations import app_settings
from app.lib.logging_helpers import AppLogger

logger, tracer = AppLogger.get_logger_and_tracer()

SQL_COPT_SS_ACCESS_TOKEN = 1256


class SqlServerManagerPool:
    def __init__(
        self,
        env: str,
        odbc_version: int = 18,
    ) -> None:
        """
        Creates a connection pool that
        handles SELECT,INSERT,UPDATE into the sql server database.

        Args:
            env (str): dev or prd.
            odbc_version (int, optional): Needed for different driver versions. Defaults to 18.
            credential_type (str, optional): Authentication method. Options are Environment, AzureCLI, Default. Defaults to "Default".
        """
        self.env = env
        self.odbc_version = odbc_version

    def get_token(
        self,
    ):
        """
        Get the token.

        Returns:
            token: Token to be used for connection.
        """
        default_credential = DefaultAzureCredential()
        raw_token = default_credential.get_token(
            "https://database.windows.net/.default"
        ).token.encode("utf-16-le")
        return s.pack(f"<I{len(raw_token)}s", len(raw_token), raw_token)

    async def get_pool(
        self,
    ):
        """
        Create the connection pool.

        Returns:
            pool: aioodbc connection pool (based of pyodbc).
        """

        if self.env == "prod":
            server_env = "-prod.2c4097eb9c0c"
        else:
            server_env = "mi-dev.2096c36ffe64"

        connstr = f"Driver=ODBC Driver {self.odbc_version} for SQL Server;Server=tcp:evt-sql{server_env}.database.windows.net,1433;Database=DS_ArsenalFC;Encrypt=yes;TrustServerCertificate=no;"
        token_struct = self.get_token()
        self.pool = await create_pool(
            dsn=connstr,
            minsize=app_settings.sql_server_settings.pool_min_size,
            maxsize=app_settings.sql_server_settings.pool_max_size,
            pool_recycle=app_settings.sql_server_settings.pool_recycle,
            echo=app_settings.sql_server_settings.echo,
            attrs_before={SQL_COPT_SS_ACCESS_TOKEN: token_struct},
            autocommit=app_settings.sql_server_settings.autocommit,
        )

    def update_pool_token(self):
        """
        Update the token in the pool.
        """
        token_struct = self.get_token()
        self.pool._conn_kwargs["attrs_before"] = {
            SQL_COPT_SS_ACCESS_TOKEN: token_struct
        }

    @asynccontextmanager
    async def get_conn(
        self,
    ):
        """
        Generator for connection. Gets a connection from the pool
        and releases it once complete.

        Yields:
            conn: Connection is acquired.
        """
        max_attempts = 2
        num_attempts = 0
        while num_attempts < max_attempts:
            try:
                logger.debug(
                    f"Pool Stats: free {self.pool.freesize} | used {len(self.pool._used)} | acquiring {self.pool._acquiring} | size {self.pool.size}"
                )
                conn = await self.pool.acquire()
                yield conn
            except Exception as e:
                if num_attempts < max_attempts:
                    logger.error(
                        f"Failed to acquire connection. Refreshing tokens. Attempt {num_attempts + 1} of {max_attempts}. Exception: {e}"
                    )
                    self.update_pool_token()
                    num_attempts += 1
                    continue
                raise
            else:
                await self.pool.release(conn)
                break

    @asynccontextmanager
    async def get_cur(self, conn: Connection):
        """
        Generator for cursor.

        Args:
            conn (Connection): Connection acquired from the pool.

        Yields:
            c: Cursor is initialized and comes with temp table based off uuid.
        """
        try:
            c: Cursor = await conn.cursor()
            yield c
        except Exception as e:
            logger.error(f"Failed to acquire cursor. Exception: {e}")
            raise
        finally:
            await c.close()

    async def select_sql(self, query: str, params: tuple):
        """
        Perform select statements. Uses cursor.executemany().

        Args:
            query (str): Query with ? parameters.
            params (list[tuple]): Parameters to replace in the query.

        Returns:
            df: Dataframe with select results.
        """
        logger.info("Selecting from database.")
        logger.debug(f"Executing \n{query}")
        async with self.get_conn() as conn:
            async with self.get_cur(conn) as c:
                try:
                    await c.execute(query, params)
                    rows = await c.fetchall()
                    df = pd.DataFrame(
                        [tuple(row) for row in rows],
                        columns=[column[0] for column in c.description],
                    )
                except Exception as e:
                    logger.error(f"Failed to execute select query. Exception: {e}")

        return df

    async def insert_sql(
        self,
        query: str,
        params: tuple,
        output: bool = False,
    ):
        """
        Perform insert statements. Uses cursor.executemany().

        Note that executemany still executes 1 by 1.

        Args:
            query (str): Query with ? parameters.
            params (list[tuple]): Parameters to replace in the query.
            output (bool, optional): If the query uses "OUTPUT INTO" to return values. Defaults to False.

        Returns:
            df: Dataframe is returned only if output is set to True
        """
        logger.info("Inserting to database.")
        async with self.get_conn() as conn:
            async with self.get_cur(conn) as c:
                try:
                    logger.debug(f"Executing \n{query}")
                    await c.executemany(query, params)
                    if output:
                        new_id = await c.fetchone()
                        await conn.commit()
                        logger.info("Committing Transaction.")
                        return new_id[0] if new_id else None
                    await conn.commit()
                    logger.info("Committing Transaction.")
                except Exception as e:
                    await conn.rollback()
                    logger.error(e)
                    logger.error("Failed to execute insert query. Rolling back.")

    async def insert_dynamic_sql(
        self, query: str, query_values: list[tuple], output: bool = False
    ):
        """
        Perform insert statements. Takes in a base insert statement up until the VALUES keyword.
        The passed query_values will be extracted into a string with form:
            (?, ?),(?, ?),(?, ?),...
        This gets appended to the pass and sent to cursor.execute()

        Args:
            query (str): Query with ? parameters.
            query_values (list[tuple]): Tuples of values to insert.
            output (bool, optional): If the id of the query needs to be returned. Defaults to False.

        Returns:
            df: Dataframe is returned only if output is set to True
        """

        logger.info("Inserting to database.")
        async with self.get_conn() as conn:
            async with self.get_cur(conn) as c:
                try:
                    tuple_length = len(query_values[0])
                    chunk_size = int(
                        app_settings.sql_server_settings.query_parameter_limit
                        / tuple_length
                    )
                    chunked_query_values = [
                        query_values[i : i + chunk_size]
                        for i in range(0, len(query_values), chunk_size)
                    ]
                    for chunk in chunked_query_values:
                        full_query = (
                            query.format(c.temp_label)
                            + " ({values})".format(
                                values="),(".join(
                                    ",".join("?" for val in term) for term in chunk
                                )
                            )
                            + ";"
                        )
                        logger.debug(f"Executing \n{full_query}")
                        await c.execute(full_query, tuple(chain(*chunk)))
                    if output:
                        await c.execute("SELECT * FROM #temp_{}".format(c.temp_label))
                        rows = await c.fetchall()
                        await conn.commit()
                        logger.info("Committing Transaction.")
                        df = pd.DataFrame(
                            [tuple(row) for row in rows],
                            columns=[column[0] for column in c.description],
                        )
                        return df
                    await conn.commit()
                    logger.info("Committing Transaction.")
                except:
                    await conn.rollback()
                    logger.error("Failed to execute insert query. Rolling back.")
                    raise

    async def close(self):
        """
        Bundles up kicking off the closing of the pool
        and waiting for all connections to close.
        """
        self.pool.close()
        await self.pool.wait_closed()
