.PHONY: help install dev build run clean test lint

CWD = $(shell pwd)
export DOCKER_IMAGE_NAME_BASE = backend-email-saver-backend

help: ## Show this help message
	@echo 'Usage:'
	@echo '  make <target>'
	@echo ''
	@echo 'Targets:'
	@awk 'BEGIN {FS = ":.*?## "} /^[a-zA-Z_-]+:.*?## / {printf "  %-20s %s\n", $$1, $$2}' $(MAKEFILE_LIST)

az_login_local:
	-mkdir .azure-for-docker
	docker run -v $(CWD)/.azure-for-docker:/root/.azure $(DOCKER_IMAGE_NAME_BASE):latest az login

test_run:
	docker run $(DOCKER_IMAGE_NAME_BASE):latest python

dev: ## Run the development server
	uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

build: ## Build Docker image
	docker build -t $(DOCKER_IMAGE_NAME_BASE):latest .

run: ## Run Docker container
	docker run -p 8000:8000 --env-file .env $(DOCKER_IMAGE_NAME_BASE):latest

docker-compose: ## Run with Docker Compose
	docker-compose up -d

docker-compose-build: ## Run with Docker Compose
	docker-compose up --build

docker-compose-build-detach: ## Run with Docker Compose in detached mode
	docker-compose up --build -d

clean: ## Clean up generated files and Docker artifacts
	find . -type d -name "__pycache__" -exec rm -r {} +
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
	find . -type f -name "*.pyd" -delete
	find . -type f -name ".coverage" -delete
	find . -type d -name "*.egg-info" -exec rm -r {} +
	find . -type d -name "*.egg" -exec rm -r {} +
	find . -type d -name ".pytest_cache" -exec rm -r {} +
	find . -type d -name ".mypy_cache" -exec rm -r {} +
	docker-compose down --rmi all --volumes --remove-orphans

test: ## Run tests
	pytest

lint: ## Run linters
	ruff check .
	mypy .
