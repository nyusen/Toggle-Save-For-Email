.PHONY: help install dev build run clean test lint

CWD = $(shell pwd)
ACR_NAME_DEV = evieacrmldev01
ACR_NAME_PROD = evieacrmlprd01
ACCT_NAME_DEV = EVIE-DataLandingZone-Dev
ACCT_NAME_PROD = EVIE-DataLandingZone-Prod
export VERSION = 0.1.0
export DOCKER_IMAGE_NAME_BASE = backend-email-saver-backend

az_login_local:
	-mkdir .azure-for-docker
	docker run -v $(CWD)/.azure-for-docker:/root/.azure $(DOCKER_IMAGE_NAME_BASE):latest az login

test_run:
	docker run $(DOCKER_IMAGE_NAME_BASE):latest python

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

build_acr_dev:
	az account set --subscription "$(ACCT_NAME_DEV)"
	az acr login -n $(ACR_NAME_DEV)
	docker build -f ./Dockerfile -t $(ACR_NAME_DEV).azurecr.io/corpus-collector-backend:$(VERSION) -t $(ACR_NAME_DEV).azurecr.io/corpus-collector-backend:latest --target deployment .
	docker push $(ACR_NAME_DEV).azurecr.io/corpus-collector-backend:$(VERSION)
	docker push $(ACR_NAME_DEV).azurecr.io/corpus-collector-backend:latest

run_from_acr:
	docker pull $(ACR_NAME_DEV).azurecr.io/corpus-collector-backend:latest
	docker run -p 8502:8502 --env-file .env $(ACR_NAME_DEV).azurecr.io/corpus-collector-backend:latest


deploy_aks_dev:
	kubectl config use-context evie-aks-mlinf-dev-01-admin
	kubectl config set-context --current --namespace=pulse
	helm upgrade corpus-collector-backend ../_charts/backend --install --values ../_charts/backend/values.dev.yaml

