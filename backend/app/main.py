from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime
from typing import List
import os
import json
from azure.storage.blob import BlobServiceClient
from dotenv import load_dotenv

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

# Get Azure Storage connection string from environment variable
connection_string = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
container_name = os.getenv("AZURE_STORAGE_CONTAINER_NAME", "email-training-data")

# Initialize the BlobServiceClient
blob_service_client = BlobServiceClient.from_connection_string(connection_string)
container_client = blob_service_client.get_container_client(container_name)

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
        blob_client = container_client.get_blob_client(filename)
        blob_client.upload_blob(email_json, overwrite=True)
        
        return {"message": "Email saved successfully", "filename": filename}
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/health")
async def health_check():
    return {"status": "healthy"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
