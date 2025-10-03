# PDF Translation Dependencies Setup Guide

## Overview

This document provides instructions for setting up the required dependencies for PDF translation functionality in the translation pipeline.

## Required Dependencies

The PDF translation functionality requires the following Python packages:

- **PyMuPDF** (>= 1.22.0) - Primary PDF processing library
- **pdfplumber** (>= 0.9.0) - Fallback text extraction for complex layouts
- **pdfminer.six** (>= 20221105) - PDF text extraction and analysis
- **pypdf** (>= 3.0.0) - PDF manipulation and analysis
- **Pillow** (>= 9.0.0) - Image processing support

## Installation Methods

### Method 1: Using Requirements File (Recommended)

```bash
# Install all PDF dependencies at once
pip install -r requirements_pdf.txt
```

### Method 2: Individual Package Installation

```bash
# Install core PDF libraries
pip install PyMuPDF>=1.22.0
pip install pdfplumber>=0.9.0
pip install pdfminer.six>=20221105
pip install pypdf>=3.0.0
pip install Pillow>=9.0.0
```

### Method 3: Virtual Environment (Recommended for Development)

```bash
# Create virtual environment
python -m venv pdf-env
source pdf-env/bin/activate  # On Windows: pdf-env\Scripts\activate

# Install dependencies
pip install -r requirements_pdf.txt
```

## Verification

After installation, verify that all dependencies are working correctly:

```bash
# Run the dependency verification script
python scripts/verify_pdf_deps.py

# For verbose output
python scripts/verify_pdf_deps.py --verbose

# Check-only mode (useful for CI/CD)
python scripts/verify_pdf_deps.py --check-only
```

## Expected Output

Successful verification should show:

```
🔍 Verifying PDF dependencies...
==================================================

Required Dependencies:
-------------------------
✅ PyMuPDF: 1.22.3 (>= 1.22.0)
✅ pdfplumber: 0.9.0 (>= 0.9.0)
✅ pdfminer.six: installed (version unknown)
✅ pypdf: 3.0.1 (>= 3.0.0)
✅ Pillow: 10.0.0 (>= 9.0.0)

==================================================
🎉 All required PDF dependencies are properly installed!

🧪 Testing basic PDF functionality...
✅ PyMuPDF: Basic functionality working
✅ pdfplumber: Import successful
✅ pypdf: Basic functionality working
```

## Troubleshooting

### Common Issues

1. **ImportError: No module named 'fitz'**
   ```bash
   pip install PyMuPDF>=1.22.0
   ```

2. **ImportError: No module named 'pdfplumber'**
   ```bash
   pip install pdfplumber>=0.9.0
   ```

3. **ImportError: No module named 'pdfminer'**
   ```bash
   pip install pdfminer.six>=20221105
   ```

4. **Externally Managed Environment Error**
   - Use a virtual environment as shown in Method 3
   - Or use `pip install --break-system-packages` (not recommended)

5. **Permission Errors**
   ```bash
   pip install --user PyMuPDF>=1.22.0 pdfplumber>=0.9.0
   ```

### Platform-Specific Notes

#### Linux (Debian/Ubuntu)
```bash
# System dependencies that might be needed
sudo apt-get update
sudo apt-get install libpng-dev libjpeg-dev libtiff-dev
```

#### macOS
```bash
# Use Homebrew for system dependencies
brew install libpng libjpeg libtiff
```

#### Windows
- No special system dependencies typically required
- Ensure you have the latest Microsoft Visual C++ Build Tools

## Optional Dependencies

The following optional dependencies are useful for development:

- **pytest** (>= 7.0.0) - Testing framework
- **black** (>= 22.0.0) - Code formatter
- **flake8** (>= 5.0.0) - Linter
- **mypy** (>= 1.0.0) - Type checker

Install optional dependencies:
```bash
pip install pytest>=7.0.0 black>=22.0.0 flake8>=5.0.0 mypy>=1.0.0
```

## Integration with PDF Components

The following scripts will use these dependencies:

1. **extract_pdf.py** - PDF text extraction
   - Uses: PyMuPDF (primary), pdfplumber (fallback)

2. **audit_pdf.py** - PDF quality assessment
   - Uses: pypdf, pdfminer.six

3. **test_audit_pdf.py** - Testing utilities
   - Uses: pytest, all PDF libraries

## Security Considerations

- Only install packages from trusted sources (PyPI)
- Keep dependencies updated to get security patches
- Use virtual environments to isolate dependencies
- Consider using `pip freeze` to lock versions in production

## Performance Notes

- PyMuPDF is the fastest for PDF processing
- pdfplumber provides better text extraction for complex layouts but is slower
- pypdf is lightweight but may have limitations with complex PDFs
- pdfminer.six is comprehensive but can be slow for large files