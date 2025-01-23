from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime
from typing import List
import os
import json
from dotenv import load_dotenv
from app.clients.blob import BlobClient
from app.lib.configurations import app_settings

load_dotenv()

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

@app.post("/api/save-email")
async def save_email(email_data: EmailData):
    try:
        # Create a unique filename using timestamp
        filename = f"email_{datetime.fromisoformat(email_data.timestamp).strftime('%Y%m%d_%H%M%S')}_{hash(email_data.subject)}.json"
        
        # Convert email data to JSON string
        email_json = email_data.model_dump_json()
        
        # Upload to blob storage
        blob_client.save_to_blob(data=email_json, file_name=filename, container_name=app_settings.container_name)
        
        return {"message": "Email saved successfully", "filename": filename}
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/health")
async def health_check():
    return {"status": "healthy"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
