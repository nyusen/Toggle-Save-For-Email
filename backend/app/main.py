from fastapi import FastAPI, BackgroundTasks, HTTPException, Depends, Header
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel
from datetime import datetime
from typing import List, Optional
import os
import json
import jwt
import requests
import base64
from dotenv import load_dotenv
from app.clients.blob import BlobClient
from app.lib.configurations import app_settings
from app.lib.logging_helpers import AppLogger

load_dotenv()

logger, tracer = AppLogger.get_logger_and_tracer()

JWKS_URL = "https://login.microsoftonline.com/common/discovery/keys"

app = FastAPI(title="Email Training Saver API")
security = HTTPBearer()

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with specific origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize the BlobServiceClient
blob_client = BlobClient()

class EmailData(BaseModel):
    subject: str
    body: str
    sender: str
    recipients: List[str]
    timestamp: str

def get_ms_public_keys():
    """Fetch Microsoft's public keys for token validation"""
    try:
        jwks = requests.get(JWKS_URL).json()
        return jwks['keys']
    except Exception as e:
        logger.error(f"Error fetching Microsoft public keys: {str(e)}")
        raise HTTPException(status_code=500, detail="Failed to fetch Microsoft public keys")

def validate_token(token: str) -> dict:
    """Validate the Exchange identity token"""
    try:
        # Decode the token header without verification
        header = jwt.get_unverified_header(token)
        kid = header.get('kid')
        
        if not kid:
            raise HTTPException(status_code=401, detail="Invalid token: No key ID")
        
        # Get the matching public key
        keys = get_ms_public_keys()
        matching_key = next((k for k in keys if k['kid'] == kid), None)
        
        if not matching_key:
            raise HTTPException(status_code=401, detail="Invalid token: Key not found")
        
        try:
            decoded_token = jwt.decode(token, matching_key, algorithms=["RS256"], options={"verify_aud": False})
            print("Token is valid:", decoded_token)
            return decoded_token
        except jwt.ExpiredSignatureError:
            print("Token has expired")
        except jwt.InvalidTokenError as e:
            print("Invalid token:", str(e))
            
            return claims
        
        return claims
    
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token has expired")
    except jwt.InvalidTokenError as e:
        raise HTTPException(status_code=401, detail=f"Invalid token: {str(e)}")
    except Exception as e:
        logger.error(f"Token validation error: {str(e)}")
        raise HTTPException(status_code=401, detail="Token validation failed")

async def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    """Dependency to validate token and return user claims"""
    token = credentials.credentials
    claims = validate_token(token)
    return claims

def save_to_blob(email_data: EmailData, user_claims: dict):
    try:
        # Create a unique filename using timestamp and user info
        user_id = user_claims.get('unique_name', 'unknown')
        filename = f"email_{datetime.fromisoformat(email_data.timestamp).strftime('%Y%m%d_%H%M%S')}_{user_id}_{hash(email_data.subject)}.json"
        
        # Add user information to the saved data
        email_data_dict = email_data.model_dump()
        email_data_dict['user_info'] = {
            'email': user_claims.get('unique_name'),
            'app_id': user_claims.get('appid'),
            'audience': user_claims.get('aud'),
            'tenant_id': user_claims.get('tid')
        }
        
        # Convert email data to JSON string
        email_json = json.dumps(email_data_dict)
        
        # Upload to blob storage
        blob_client.save_to_blob(data=email_json, file_name=filename, container_name=app_settings.container_name)

        logger.info(f"Successfully saved file to blob storage: {filename}")
        return {"message": "Email saved successfully", "filename": filename}
    
    except Exception as e:
        logger.error(f"Failed to save file to blob storage: {filename}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/save-email")
async def save_email(
    email_data: EmailData,
    background_tasks: BackgroundTasks,
    user_claims: dict = Depends(get_current_user)
):
    background_tasks.add_task(save_to_blob, email_data, user_claims)
    return JSONResponse(
        status_code=202,
        content={"message": "Email is saving to blob"},
    )

@app.get("/health")
async def health_check():
    return {"status": "healthy"}
