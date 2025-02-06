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
from jwt.algorithms import RSAAlgorithm
import requests
import base64
from dotenv import load_dotenv
from app.clients.blob import BlobClient
from app.lib.configurations import app_settings
from app.lib.logging_helpers import AppLogger

load_dotenv()

logger, tracer = AppLogger.get_logger_and_tracer()

CLIENT_ID = "6d7e781e-9cf5-48ff-8c05-b697ca1a90e3"
JWKS_URL = "https://login.microsoftonline.com/common/discovery/v2.0/keys"

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


def save_to_blob(email_data: EmailData, user_claims: dict):
    try:
        # Create a unique filename using timestamp and user info
        user_email = user_claims.get('email', 'unknown')
        filename = f"email_{datetime.fromisoformat(email_data.timestamp).strftime('%Y%m%d_%H%M%S')}_{user_email}_{hash(email_data.subject)}.json"
        
        # Add user information to the saved data
        email_data_dict = email_data.model_dump()
        
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
    user_claims: dict = Depends(verify_user)
):
    background_tasks.add_task(save_to_blob, email_data, user_claims)
    return JSONResponse(
        status_code=202,
        content={"message": "Email is saving to blob"},
    )

@app.get("/health")
async def health_check():
    return {"status": "healthy"}
