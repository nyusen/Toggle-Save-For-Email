# Email Training Saver Backend

This is a FastAPI-based backend service that handles saving emails to Azure Blob Storage. It's containerized using Docker for easy deployment and uses uv for fast, reliable Python package management.

## Prerequisites

- Docker and Docker Compose
- Azure Storage Account and Container
- Python 3.11+ (for local development)

## Setup

1. Create a `.env` file from the template:
```bash
cp .env.example .env
```

2. Update the `.env` file with your Azure Storage credentials:
```
AZURE_STORAGE_CONNECTION_STRING=your_connection_string_here
AZURE_STORAGE_CONTAINER_NAME=corpus-collector-data
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

1. Install uv (if not already installed):
```bash
pip install uv
```

2. Create a virtual environment and install dependencies:
```bash
uv venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
uv pip install .
```

3. Run the service:
```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## API Documentation

When the service is running, you can access the auto-generated API documentation at:
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## Why uv?

We use uv for package management because it offers several advantages:
- Much faster package installation and dependency resolution
- Reliable and reproducible builds
- Built-in virtual environment management
- Compatible with standard Python tooling
