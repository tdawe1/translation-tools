#!/usr/bin/env python3
"""
verify_pdf_deps.py

Script to verify that all PDF translation dependencies are properly installed.
"""

import sys
import importlib

def check_import(module_name, description):
    """Check if a module can be imported."""
    try:
        importlib.import_module(module_name)
        print(f"✅ {description} - OK")
        return True
    except ImportError as e:
        print(f"❌ {description} - MISSING: {e}")
        return False

def main():
    """Main function to verify dependencies."""
    print("Verifying PDF translation dependencies...")
    print("=" * 50)
    
    # Core PDF processing libraries
    deps = [
        ("fitz", "PyMuPDF (fitz) - Core PDF processing"),
        ("pdfplumber", "pdfplumber - Fallback text extraction"),
        ("pdfminer", "pdfminer - PDF text analysis"),
        ("pypdf", "pypdf - PDF manipulation"),
        ("PIL", "Pillow - Image processing"),
    ]
    
    # Testing dependencies
    test_deps = [
        ("pytest", "pytest - Testing framework"),
        ("pytest_asyncio", "pytest-asyncio - Async testing support"),
    ]
    
    # WebSocket and web dependencies
    web_deps = [
        ("websockets", "websockets - Real-time updates"),
        ("aiohttp", "aiohttp - HTTP client/server"),
        ("aiohttp_cors", "aiohttp-cors - CORS support"),
    ]
    
    # Development dependencies
    dev_deps = [
        ("black", "black - Code formatting"),
        ("flake8", "flake8 - Code linting"),
        ("mypy", "mypy - Type checking"),
    ]
    
    # Check all dependencies
    all_passed = True
    
    print("\nCore PDF Dependencies:")
    for module, description in deps:
        if not check_import(module, description):
            all_passed = False
    
    print("\nTesting Dependencies:")
    for module, description in test_deps:
        if not check_import(module, description):
            all_passed = False
    
    print("\nWeb Dependencies:")
    for module, description in web_deps:
        if not check_import(module, description):
            all_passed = False
    
    print("\nDevelopment Dependencies:")
    for module, description in dev_deps:
        if not check_import(module, description):
            all_passed = False
    
    print("\n" + "=" * 50)
    if all_passed:
        print("✅ All dependencies are properly installed!")
        return 0
    else:
        print("❌ Some dependencies are missing. Please install them.")
        return 1

if __name__ == "__main__":
    sys.exit(main())