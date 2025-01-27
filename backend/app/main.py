from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime
from typing import List
import os
import json
from dotenv import load_dotenv
from app.clients.blob import BlobClient
from app.lib.configurations import app_settings
from app.lib.logging_helpers import AppLogger

load_dotenv()

logger, tracer = AppLogger.get_logger_and_tracer()

app = FastAPI(title="Email Training Saver API")

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

def save_to_blob(email_data: EmailData):
    try:
        # Create a unique filename using timestamp
        filename = f"email_{datetime.fromisoformat(email_data.timestamp).strftime('%Y%m%d_%H%M%S')}_{hash(email_data.subject)}.json"
        
        # Convert email data to JSON string
        email_json = email_data.model_dump_json()
        
        # Upload to blob storage
        blob_client.save_to_blob(data=email_json, file_name=filename, container_name=app_settings.container_name)

        logger.info(f"Successfully saved file to blob storage: {filename}")
        return {"message": "Email saved successfully", "filename": filename}
    
    except Exception as e:
        logger.error(f"Failed to save file to blob storage: {filename}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/save-email")
async def save_email(email_data: EmailData, background_tasks: BackgroundTasks):
    background_tasks.add_task(save_to_blob, email_data)
    return JSONResponse(
        status_code=202,
        content={"message": "Email is saving to blob"},
    )
    

@app.get("/health")
async def health_check():
    return {"status": "healthy"}
