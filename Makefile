.PHONY: help install install-dev test lint format type-check clean run run-http

help: ## Show this help message
	@echo "Available commands:"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-20s\033[0m %s\n", $$1, $$2}'

install: ## Install the package
	uv venv
	uv pip install -e .

install-dev: ## Install the package with development dependencies
	uv venv
	uv pip install -e ".[dev]"

sync: ## Sync dependencies and create lock file
	uv sync

test: ## Run tests
	pytest

test-cov: ## Run tests with coverage
	pytest --cov=auth --cov=core --cov=teams --cov-report=html --cov-report=term

lint: ## Run linting checks
	flake8 auth core teams main.py
	black --check auth core teams main.py
	isort --check-only auth core teams main.py

format: ## Format code
	black auth core teams main.py
	isort auth core teams main.py

type-check: ## Run type checking
	mypy auth core teams main.py

clean: ## Clean up build artifacts
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	find . -type d -name __pycache__ -delete
	find . -type f -name "*.pyc" -delete
	rm -f .coverage
	rm -rf htmlcov/
	rm -f teams_mcp_server_debug.log

run: ## Run the server with stdio transport
	export $$(cat .env | grep -v '^#' | grep -v '^$$' | xargs) && python main.py --transport stdio

run-http: ## Run the server with HTTP transport
	export $$(cat .env | grep -v '^#' | grep -v '^$$' | xargs) && python main.py --transport streamable-http --port 8000

run-dev: ## Run the server in development mode
	export $$(cat .env | grep -v '^#' | grep -v '^$$' | xargs) && SKIP_AUTH=true python main.py --transport streamable-http --port 8000

env-check: ## Check environment configuration
	@echo "Checking environment variables..."
	@export $$(cat .env | grep -v '^#' | grep -v '^$$' | xargs) && python -c "import os; print('MICROSOFT_OAUTH_CLIENT_ID:', 'SET' if os.getenv('MICROSOFT_OAUTH_CLIENT_ID') else 'NOT SET')"
	@export $$(cat .env | grep -v '^#' | grep -v '^$$' | xargs) && python -c "import os; print('MICROSOFT_OAUTH_CLIENT_SECRET:', 'SET' if os.getenv('MICROSOFT_OAUTH_CLIENT_SECRET') else 'NOT SET')"

deps: ## Show dependency tree
	uv pip list

build: ## Build the package
	uv build

dev-setup: install-dev sync ## Complete development setup
