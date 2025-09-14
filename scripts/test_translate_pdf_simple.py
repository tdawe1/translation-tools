#!/usr/bin/env python3
"""
test_translate_pdf_simple.py

Simple test of the PDF translator CLI interface without full dependencies.
"""

import sys
import os
import tempfile
from pathlib import Path

# Add the scripts directory to Python path
sys.path.insert(0, '/home/thomas/translation-tools/translations-pptx-pipeline/scripts')

def test_cli_parsing():
    """Test CLI argument parsing without full dependencies."""
    
    # Mock all the problematic imports
    import unittest.mock as mock
    
    with mock.patch.dict('sys.modules', {
        'fitz': mock.MagicMock(),
        'pdfplumber': mock.MagicMock(),
        'pypdf': mock.MagicMock(),
        'pdfminer': mock.MagicMock(),
        'pdfminer.high_level': mock.MagicMock(),
        'pdfminer.layout': mock.MagicMock(),
        'openai': mock.MagicMock(),
    }):
        # Mock the modules to avoid import errors
        mock_fit = mock.MagicMock()
        mock_pdfplumber = mock.MagicMock()
        mock_pypdf = mock.MagicMock()
        mock_pdfminer = mock.MagicMock()
        mock_openai = mock.MagicMock()
        
        sys.modules['fitz'] = mock_fit
        sys.modules['pdfplumber'] = mock_pdfplumber
        sys.modules['pypdf'] = mock_pypdf
        sys.modules['pdfminer'] = mock_pdfminer
        sys.modules['pdfminer.high_level'] = mock.MagicMock()
        sys.modules['pdfminer.layout'] = mock.MagicMock()
        sys.modules['openai'] = mock_openai
        
        try:
            # Now try to import and test the CLI
            from translate_pdf import PDFTranslator, parse_page_range
            
            print("✓ Successfully imported PDFTranslator with mocked dependencies")
            
            # Test page range parsing
            result = parse_page_range("1-5")
            assert result == (1, 5), f"Expected (1, 5), got {result}"
            
            result = parse_page_range("3")
            assert result == (3, 3), f"Expected (3, 3), got {result}"
            
            result = parse_page_range(None)
            assert result is None, f"Expected None, got {result}"
            
            print("✓ Page range parsing works correctly")
            
            # Test basic translator initialization
            with tempfile.TemporaryDirectory() as temp_dir:
                input_pdf = os.path.join(temp_dir, "test.pdf")
                output_pdf = os.path.join(temp_dir, "output.pdf")
                Path(input_pdf).touch()
                
                translator = PDFTranslator(
                    input_path=input_pdf,
                    output_path=output_pdf,
                    offline_mode=True,
                    cache_only=False
                )
                
                assert translator.input_path == input_pdf
                assert translator.output_path == output_pdf
                assert translator.offline_mode == True
                assert translator.cache_only == False
                
                print("✓ Translator initialization works correctly")
            
            print("\n✅ All basic tests passed!")
            return True
            
        except Exception as e:
            print(f"❌ Test failed: {e}")
            import traceback
            traceback.print_exc()
            return False

def test_cli_help():
    """Test CLI help functionality."""
    
    import subprocess
    import sys
    
    # Test if the script can be invoked without crashing
    try:
        result = subprocess.run([
            sys.executable, 
            '/home/thomas/translation-tools/translations-pptx-pipeline/scripts/translate_pdf.py',
            '--help'
        ], capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            print("✓ CLI help command works")
            print(f"Help output length: {len(result.stdout)} characters")
            return True
        else:
            print(f"❌ CLI help failed with return code {result.returncode}")
            print(f"Error: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("❌ CLI help command timed out")
        return False
    except Exception as e:
        print(f"❌ CLI help test failed: {e}")
        return False

if __name__ == "__main__":
    print("Testing PDF translator CLI interface...")
    print("=" * 50)
    
    success = True
    
    print("\n1. Testing basic functionality with mocked dependencies...")
    success &= test_cli_parsing()
    
    print("\n2. Testing CLI help command...")
    success &= test_cli_help()
    
    if success:
        print("\n🎉 All tests passed!")
        sys.exit(0)
    else:
        print("\n💥 Some tests failed!")
        sys.exit(1)