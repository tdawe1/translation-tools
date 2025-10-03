# Environment Setup Guide

This guide explains how to set up your environment to run the PDF translation pipeline.

## Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

## Installation

### 1. Install Dependencies

Install all required dependencies using pip:

```bash
# Install core dependencies
pip install -r requirements.txt

# Install PDF-specific dependencies
pip install -r requirements_pdf.txt
```

### 2. Set up Environment Variables

Copy the example environment file and customize it:

```bash
cp .env.example .env
```

Then edit `.env` to add your API keys and configuration:

```bash
# Required for translation
OPENAI_API_KEY=sk-your-openai-key

# Optional overrides
OPENAI_MODEL=gpt-4o-2024-08-06
```

### 3. Verify Installation

Run the verification script to ensure all dependencies are properly installed:

```bash
make verify-deps
```

## Running Tests

### Run All Tests

```bash
make test-all
```

### Run Specific Test Suites

```bash
# Run unit tests
make test

# Run PDF-specific tests
make test-pdf

# Run quality metrics tests
make test-quality

# Run integration tests
make test-integration

# Run clean environment tests
make test-clean
```

## Running Translations

### Translate a PDF

```bash
make translate-pdf INPUT=document.pdf OUTPUT=translated_document.pdf
```

### Translate with Specific Options

```bash
make translate-pdf INPUT=document.pdf OUTPUT=translated_document.pdf MODEL=gpt-4o-mini PAGES=1-5
```

## Required Services

The following services may be required depending on your usage:

1. **OpenAI API** - Required for translation if not using cache-only mode
2. **PDF Processing Libraries** - Installed via requirements_pdf.txt
3. **Python 3.8+** - Runtime environment

## Troubleshooting

### Missing Dependencies

If you encounter import errors, ensure all dependencies are installed:

```bash
make setup
```

### Permission Issues

If you encounter permission issues, you may need to run with sudo or use a virtual environment:

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
make setup
```

### Test Failures

If tests fail due to missing test data, you can run only the unit tests that don't require external files:

```bash
make test-clean
```