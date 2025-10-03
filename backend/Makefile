.PHONY: install run test clean help

help:  ## Show this help message
	@echo "Translation Pipeline Backend - Makefile"
	@echo ""
	@awk 'BEGIN {FS = ":.*##"; printf "%-20s %s\n", "Target", "Description"} /^[a-zA-Z_-]+:.*?##/ { printf "%-20s %s\n", $$1, $$2 }' $(MAKEFILE_LIST)

install:  ## Install dependencies
	python -m venv venv
	. venv/bin/activate && pip install -r requirements.txt
	@echo "Dependencies installed successfully"

run:  ## Run the development server
	@echo "Starting FastAPI server on http://localhost:8000"
	@echo "API docs available at http://localhost:8000/docs"
	. venv/bin/activate && uvicorn app.main:app --reload --host 0.0.0.0 --port 8000

test:  ## Run tests
	. venv/bin/activate && pytest -v

test-cov:  ## Run tests with coverage
	. venv/bin/activate && pytest --cov=app --cov-report=html --cov-report=term

lint:  ## Run linting (placeholder)
	@echo "Linting not configured yet"

format:  ## Format code (placeholder)
	@echo "Code formatting not configured yet"

clean:  ## Clean up temporary files
	@find . -type d -name __pycache__ -exec rm -rf {} +
	@find . -type f -name "*.pyc" -delete
	@rm -rf htmlcov/
	@rm -rf .pytest_cache/
	@echo "Cleaned up temporary files"

clean-all:  ## Clean up everything including virtual environment
	$(MAKE) clean
	rm -rf venv/
	rm -rf uploads/
	rm -rf outputs/
	@echo "All cleaned up"