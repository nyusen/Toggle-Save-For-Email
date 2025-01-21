# Email Training Saver Backend

This is a FastAPI-based backend service that handles saving emails to Azure Blob Storage. It's containerized using Docker for easy deployment.

## Prerequisites

- Docker and Docker Compose
- Azure Storage Account and Container

## Setup

1. Create a `.env` file from the template:
```bash
cp .env.example .env
```

2. Update the `.env` file with your Azure Storage credentials:
```
AZURE_STORAGE_CONNECTION_STRING=your_connection_string_here
AZURE_STORAGE_CONTAINER_NAME=email-training-data
```

## Running the Service

Using Docker Compose:
```bash
docker-compose up --build
```

The API will be available at `http://localhost:8000`

## API Endpoints

- `POST /api/save-email`: Save email data to Azure Blob Storage
  - Request body: JSON object containing email data (subject, body, sender, recipients, timestamp)
  - Returns: JSON object with success message and filename

- `GET /api/health`: Health check endpoint
  - Returns: JSON object with service status

## Development

To run the service without Docker:

1. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the service:
```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## API Documentation

When the service is running, you can access the auto-generated API documentation at:
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`
