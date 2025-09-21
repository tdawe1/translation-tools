#!/usr/bin/env python3
"""
test_clean_environment.py

Simple test to verify basic functionality in a clean environment.
"""

import sys
import os
import tempfile
import json
from pathlib import Path

def test_imports():
    """Test that all required modules can be imported."""
    # Add the scripts directory to the path
    scripts_path = str(Path(__file__).parent / "scripts")
    if scripts_path not in sys.path:
        sys.path.insert(0, scripts_path)
    
    try:
        # Try to import the main translate_pdf module
        import scripts.translate_pdf
        print("✅ translate_pdf module import - OK")
    except Exception as e:
        print(f"❌ translate_pdf module import - FAILED: {e}")
        return False
    
    return True

def test_basic_syntax():
    """Test that Python files have valid syntax."""
    scripts_dir = Path(__file__).parent / "scripts"
    
    # List of key Python files to check
    key_files = [
        "translate_pdf.py",
        "extract_pdf.py",
        "apply_pdf_translation.py",
        "pdf_layout_engine.py"
    ]
    
    for filename in key_files:
        filepath = scripts_dir / filename
        if filepath.exists():
            try:
                # Try to compile the file to check for syntax errors
                with open(filepath, 'r', encoding='utf-8') as f:
                    compile(f.read(), str(filepath), 'exec')
                print(f"✅ {filename} syntax - OK")
            except Exception as e:
                print(f"❌ {filename} syntax - FAILED: {e}")
                return False
        else:
            print(f"⚠️  {filename} not found")
    
    return True

def main():
    """Main test function."""
    print("Testing basic functionality in clean environment...")
    print("=" * 60)
    
    # Test imports
    if not test_imports():
        return 1
    
    # Test basic syntax
    if not test_basic_syntax():
        return 1
    
    print("\n" + "=" * 60)
    print("✅ All basic tests passed!")
    return 0

if __name__ == "__main__":
    sys.exit(main())
