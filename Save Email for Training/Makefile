.PHONY: help install dev build run clean

CWD = $(shell pwd)
ACR_NAME_DEV = evieacrmldev01
ACR_NAME_PROD = evieacrmlprd01
ACCT_NAME_DEV = EVIE-DataLandingZone-Dev
ACCT_NAME_PROD = EVIE-DataLandingZone-Prod
export VERSION = 0.1.0
export DOCKER_IMAGE_NAME_BASE = outlook-save-for-training

help: ## Display this help message
	@echo "Available commands:"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-30s\033[0m %s\n", $$1, $$2}'

install: ## Install dependencies
	npm ci

dev: ## Run development server
	npm start

build: ## Build Docker image
	docker build -t $(DOCKER_IMAGE_NAME_BASE):latest .

run: ## Run Docker container
	docker run -p 3000:3000 $(DOCKER_IMAGE_NAME_BASE):latest

build_acr_dev: ## Build and push to Azure Container Registry (Dev)
	az account set --subscription "$(ACCT_NAME_DEV)"
	az acr login -n $(ACR_NAME_DEV)
	docker build -f ./Dockerfile -t $(ACR_NAME_DEV).azurecr.io/$(DOCKER_IMAGE_NAME_BASE):$(VERSION) -t $(ACR_NAME_DEV).azurecr.io/$(DOCKER_IMAGE_NAME_BASE):latest .
	docker push $(ACR_NAME_DEV).azurecr.io/$(DOCKER_IMAGE_NAME_BASE):$(VERSION)
	docker push $(ACR_NAME_DEV).azurecr.io/$(DOCKER_IMAGE_NAME_BASE):latest

deploy_aks_dev: ## Deploy to AKS Dev environment
	kubectl config use-context evie-aks-mlinf-dev-01-admin
	kubectl config set-context --current --namespace=pulse
	helm upgrade outlook-save-for-training ../_charts/outlook-save-for-training --install --values ../_charts/outlook-save-for-training/values.dev.yaml

run_from_acr:
	docker pull $(ACR_NAME_DEV).azurecr.io/$(DOCKER_IMAGE_NAME_BASE):latest
	docker run -p 3000:3000 $(ACR_NAME_DEV).azurecr.io/$(DOCKER_IMAGE_NAME_BASE):latest
