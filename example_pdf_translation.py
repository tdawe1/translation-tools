#!/usr/bin/env python3
"""
Example usage of the PDF back-projector script.

This script demonstrates how to use the apply_pdf_translation.py script
to replace Japanese text with English translations in a PDF document.
"""

import json
import os
import sys
from pathlib import Path

def create_sample_translations():
    """Create sample translation data for demonstration."""
    translations = [
        {
            "original": "こんにちは世界",
            "translated": "Hello World",
            "font_scaling": 0.9
        },
        {
            "original": "日本語のテキスト",
            "translated": "Japanese text",
            "font_scaling": 1.0
        },
        {
            "original": "これはテストです",
            "translated": "This is a test",
            "font_scaling": 0.85
        },
        {
            "original": "PDFドキュメント",
            "translated": "PDF document",
            "font_scaling": 0.95
        },
        {
            "original": "翻訳システム",
            "translated": "Translation system",
            "font_scaling": 0.9
        }
    ]
    
    with open("example_translations.json", "w", encoding="utf-8") as f:
        json.dump(translations, f, ensure_ascii=False, indent=2)
    
    print("Created example_translations.json")
    return "example_translations.json"

def run_example():
    """Run the PDF back-projector example."""
    print("PDF Back-Projector Example")
    print("=" * 40)
    
    # Create sample translations
    translations_file = create_sample_translations()
    
    # Example usage
    print("\nExample Usage:")
    print("python scripts/apply_pdf_translation.py \\")
    print("  --input original.pdf \\")
    print("  --output translated.pdf \\")
    print(f"  --translations {translations_file}")
    
    print("\nFeatures demonstrated:")
    print("- Japanese text detection and replacement")
    print("- Font scaling for text expansion")
    print("- Formatting preservation")
    print("- Layout adjustments")
    
    print("\nNote: This example requires a PDF file with Japanese text.")
    print("Replace 'original.pdf' with your actual PDF file path.")
    
    # Clean up
    if os.path.exists(translations_file):
        os.remove(translations_file)
        print(f"\nCleaned up {translations_file}")

if __name__ == "__main__":
    run_example()