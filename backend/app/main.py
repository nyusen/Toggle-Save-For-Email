from contextlib import asynccontextmanager
from datetime import datetime
from typing import List, Optional
import os
import json
import jwt
from jwt.algorithms import RSAAlgorithm
import requests
import base64
from uuid import uuid4

from fastapi import FastAPI, BackgroundTasks, HTTPException, Depends, Header, Request
from fastapi.responses import JSONResponse
from fastapi.middleware import Middleware
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
from dotenv import load_dotenv
import pandas as pd

from app.clients.blob import BlobClient
from app.lib.configurations import app_settings
from app.lib.database_manager import SqlServerManagerPool
from app.lib.logging_helpers import AppLogger

load_dotenv()

logger, tracer = AppLogger.get_logger_and_tracer()



CLIENT_ID = "6d7e781e-9cf5-48ff-8c05-b697ca1a90e3"
JWKS_URL = "https://login.microsoftonline.com/common/discovery/v2.0/keys"

security = HTTPBearer()

def make_middleware() -> List[Middleware]:
    middleware = [
        # Configure CORS
        Middleware(
            CORSMiddleware,
            allow_origins=['*'],
            allow_methods=['*'],
            allow_headers=['*'],
        ),
    ]
    return middleware

@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info('Starting instance.')
    logger.info('Creating connection pool.')
    pool = SqlServerManagerPool(env=app_settings.env, odbc_version=app_settings.sql_server_settings.odbc_version)
    await pool.get_pool()
    yield {'pool': pool}
    logger.info('Closing connection pool.')
    await pool.close()
    logger.info('Shutting down instance.')

def create_app() -> FastAPI:
    app_ = FastAPI(title="Email Training Saver API", middleware=make_middleware(), lifespan=lifespan)
    return app_

app = create_app()

# Initialize the BlobServiceClient
blob_client = BlobClient()

class EmailData(BaseModel):
    subject: str
    body: str
    sender: str
    recipients: List[str]
    tags: List[int]
    timestamp: str

class Tag(BaseModel):
    description: str

def get_ms_public_keys():
    """Fetch Microsoft's public keys for token validation"""
    try:
        jwks = requests.get(JWKS_URL).json()
        return jwks
    except Exception as e:
        logger.error(f"Error fetching Microsoft public keys: {str(e)}")
        raise HTTPException(status_code=500, detail="Failed to fetch Microsoft public keys")

def validate_token(token: str) -> dict:
    """Validate the Exchange identity token against Microsoft's public keys"""
    try:
        # Decode the token header without verification
        header = jwt.get_unverified_header(token)
        kid = header.get('kid')
        
        if not kid:
            logger.error("Token validation failed: No key ID in token header")
            raise HTTPException(status_code=401, detail="Invalid token: No key ID")
        
        # Get the matching public key
        jwks = get_ms_public_keys()
        matching_key = next((k for k in jwks["keys"] if k['kid'] == kid), None)
        
        if not matching_key:
            logger.error(f"Token validation failed: No matching key found for kid {kid}")
            raise HTTPException(status_code=401, detail="Invalid token: Key not found")

        # Convert the JWK to PEM format
        public_key = RSAAlgorithm.from_jwk(matching_key)
        
        try:
            decoded_token = jwt.decode(
                token,
                key=public_key,
                algorithms=["RS256"],
                audience=CLIENT_ID,
            )
            logger.info("Token successfully validated")
            return decoded_token
            
        except jwt.ExpiredSignatureError:
            logger.error("Token validation failed: Token has expired")
            raise HTTPException(status_code=401, detail="Token has expired")
        except jwt.InvalidIssuerError:
            logger.error("Token validation failed: Invalid issuer")
            raise HTTPException(status_code=401, detail="Invalid token issuer")
        except jwt.InvalidTokenError as e:
            logger.error(f"Token validation failed: {str(e)}")
            raise HTTPException(status_code=401, detail=f"Invalid token: {str(e)}")
            
    except Exception as e:
        logger.error(f"Unexpected error during token validation: {str(e)}")
        raise HTTPException(status_code=401, detail="Token validation failed")

async def verify_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Dependency to validate token and return user claims"""
    token = credentials.credentials
    claims = validate_token(token)
    # if True or not email or not (email.endswith("@eventellect.com") or email.endswith("@unifiedevents.com")):
    #     raise HTTPException(status_code=403, detail="User not authorized to access this API")
    return claims

async def persist_email_data(email_data: EmailData, request_metadata: Request):
    email_uuid = uuid4()

    try:
        # Save email data to sql server
        pool: SqlServerManagerPool = request_metadata.state.pool
        
        insert_email_query = """
        insert into ds_experimentation.dbo.training_email (uuid, sender, recipients, subject)
        output inserted.id
        values (?, ?, ?, ?)
        """
        email_id = await pool.insert_sql(insert_email_query, [(email_uuid, email_data.sender, ",".join(email_data.recipients), email_data.subject)], output=True)

        if email_data.tags:
            insert_tag_query = """
            insert into ds_experimentation.dbo.training_email_tag (email_id, tag_id)
            values (?, ?)
            """
            await pool.insert_sql(insert_tag_query, [(email_id, tag) for tag in email_data.tags])

        # Save email to blob
        # Create a unique filename using timestamp and user info
        filename = f"{email_uuid}.json"
        
        # Add user information to the saved data
        email_data_dict = email_data.model_dump()
        
        # Convert email data to JSON string
        email_json = json.dumps(email_data_dict)
        
        # Upload to blob storage
        blob_client.save_to_blob(data=email_json, file_name=filename, container_name=app_settings.container_name)

        logger.info(f"Successfully saved file to blob storage: {filename}")
        return {"message": "Email saved successfully", "filename": filename}
    
    except Exception as e:
        logger.error(f"Failed to save email data to sql server: {str(e)}")
        raise

@app.get("/tag")
async def get_tags(request_metadata: Request):
    pool: SqlServerManagerPool = request_metadata.state.pool
    tag_select_query = """
    select * from ds_experimentation.dbo.training_tag
    """
    tag_df: pd.DataFrame = await pool.select_sql(tag_select_query, ())
    tag_records = tag_df.to_dict(orient='records')

    # tags = [
    #     {"id": 1, "description": "Pulse Report"},
    #     {"id": 2, "description": "Internal Communication"},
    #     {"id": 3, "description": "External Communication"},
    #     {"id": 4, "description": "Playoff Preview"},
    #     {"id": 5, "description": "Underwriting"},
    # ]

    return JSONResponse(content=tag_records, media_type="application/json")
    

@app.post("/tag")
async def create_tag(tag: Tag, request_metadata: Request,):
    pool: SqlServerManagerPool = request_metadata.state.pool
    insert_tag_query = """
    insert into ds_experimentation.dbo.training_tag (description)
    output inserted.id
    values (?)
    """

    inserted_id = await pool.insert_sql(insert_tag_query, [(tag.description,)], output=True)
    logger.info(f"Inserted tag with ID: {inserted_id}")
    return JSONResponse(content={"id": inserted_id, "description": tag.description}, media_type="application/json")
    
@app.post("/save-email")
async def save_email(
    email_data: EmailData,
    request_metadata: Request,
    background_tasks: BackgroundTasks,
    user_claims: dict = Depends(verify_user)
):
    background_tasks.add_task(persist_email_data, email_data, request_metadata)
    return JSONResponse(
        status_code=202,
        content={"message": "Email is saving to blob"},
    )

@app.get("/health")
async def health_check():
    return {"status": "healthy"}
