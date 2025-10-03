#!/usr/bin/env python3
"""
Example usage of the PDF text extraction component.

This script demonstrates how to use the extract_pdf.py module
for extracting Japanese text from PDF files with layout preservation.
"""

import sys
import os
from pathlib import Path

# Add the scripts directory to Python path
sys.path.insert(0, str(Path(__file__).parent / "scripts"))

from extract_pdf import PDFExtractor, save_extraction_result

def example_basic_extraction():
    """Example: Basic PDF text extraction."""
    print("=== Basic PDF Text Extraction ===")
    
    # Initialize extractor with default settings
    extractor = PDFExtractor()
    
    # Extract text from PDF
    result = extractor.extract_text_blocks("example.pdf")
    
    # Print summary
    print(f"Extracted {result.total_blocks} text blocks from {len(result.pages)} pages")
    print(f"Found {result.total_japanese_blocks} Japanese text blocks")
    print(f"Extraction completed in {result.extraction_time:.2f} seconds")
    
    # Save result
    save_extraction_result(result, "extracted_text.json")
    print("Results saved to extracted_text.json")

def example_japanese_only_extraction():
    """Example: Extract only Japanese text blocks."""
    print("\n=== Japanese-Only Extraction ===")
    
    extractor = PDFExtractor()
    
    # Extract all text first
    full_result = extractor.extract_text_blocks("example.pdf")
    
    # Filter to Japanese text only
    japanese_result = extractor.filter_japanese_text(full_result)
    
    print(f"Filtered to {japanese_result.total_blocks} Japanese text blocks")
    
    # Save filtered result
    save_extraction_result(japanese_result, "japanese_only.json")

def example_translation_pipeline_format():
    """Example: Convert to translation pipeline format."""
    print("\n=== Translation Pipeline Format ===")
    
    extractor = PDFExtractor()
    
    # Extract text
    result = extractor.extract_text_blocks("example.pdf")
    
    # Convert to translation-compatible format
    translation_data = extractor.to_translation_format(result)
    
    print(f"Prepared {len(translation_data['japanese_texts'])} texts for translation")
    print(f"Layout info for {len(translation_data['layout_info']['pages'])} pages")
    
    # Save translation format
    import json
    with open("translation_input.json", "w", encoding="utf-8") as f:
        json.dump(translation_data, f, ensure_ascii=False, indent=2)
    
    print("Translation input saved to translation_input.json")

def example_detailed_extraction():
    """Example: Detailed extraction with metadata."""
    print("\n=== Detailed Extraction with Metadata ===")
    
    # Initialize with detailed mode
    extractor = PDFExtractor(use_fallback=True, min_confidence=0.7)
    
    # Extract with detailed metadata
    result = extractor.extract_text_blocks("example.pdf", detailed=True)
    
    # Show detailed statistics
    stats = result.metadata.get("extraction_stats", {})
    print(f"Extraction methods used: {result.extraction_methods}")
    print(f"Failed pages: {stats.get('failed_pages', 0)}")
    
    # Show sample text block with metadata
    if result.pages and result.pages[0].text_blocks:
        first_block = result.pages[0].text_blocks[0]
        print(f"\nSample text block:")
        print(f"  ID: {first_block.id}")
        print(f"  Text: {first_block.text}")
        print(f"  Font: {first_block.font_name} ({first_block.font_size}pt)")
        print(f"  Position: ({first_block.x0}, {first_block.y0}) - ({first_block.x1}, {first_block.y1})")
        print(f"  Confidence: {first_block.confidence}")
        print(f"  Metadata: {first_block.metadata}")

def example_batch_processing():
    """Example: Process multiple PDF files."""
    print("\n=== Batch Processing ===")
    
    pdf_files = ["document1.pdf", "document2.pdf", "presentation.pdf"]
    extractor = PDFExtractor()
    
    all_results = []
    
    for pdf_file in pdf_files:
        if os.path.exists(pdf_file):
            print(f"Processing {pdf_file}...")
            result = extractor.extract_text_blocks(pdf_file)
            all_results.append(result)
            
            # Save individual result
            output_file = f"{Path(pdf_file).stem}_extracted.json"
            save_extraction_result(result, output_file)
            print(f"  -> {output_file}")
        else:
            print(f"  File not found: {pdf_file}")
    
    print(f"\nProcessed {len(all_results)} files")
    total_blocks = sum(r.total_blocks for r in all_results)
    total_japanese = sum(r.total_japanese_blocks for r in all_results)
    print(f"Total text blocks: {total_blocks}")
    print(f"Total Japanese blocks: {total_japanese}")

def example_error_handling():
    """Example: Error handling and fallback."""
    print("\n=== Error Handling and Fallback ===")
    
    # Initialize with fallback enabled
    extractor = PDFExtractor(use_fallback=True, min_confidence=0.6)
    
    try:
        # Try to extract from a potentially problematic PDF
        result = extractor.extract_text_blocks("problematic.pdf")
        
        if result.total_blocks == 0:
            print("No text blocks extracted - PDF may be image-based or encrypted")
        else:
            print(f"Successfully extracted {result.total_blocks} blocks")
            
            # Check if fallback was used
            if "pdfplumber" in result.extraction_methods:
                print("Used pdfplumber fallback for some pages")
    
    except Exception as e:
        print(f"Extraction failed: {e}")
        print("Consider checking if the PDF is password-protected or corrupted")

def main():
    """Main function to run examples."""
    print("PDF Text Extraction Examples")
    print("=" * 40)
    
    # Note: These examples assume you have PDF files to test with
    # In a real scenario, you would replace "example.pdf" with actual PDF files
    
    print("This script demonstrates various usage patterns for PDF text extraction.")
    print("Replace 'example.pdf' with actual PDF files to run the examples.")
    
    # Check if we have sample PDFs
    sample_files = ["example.pdf", "document1.pdf", "problematic.pdf"]
    existing_files = [f for f in sample_files if os.path.exists(f)]
    
    if not existing_files:
        print(f"\nNo sample PDF files found. Looking for: {sample_files}")
        print("Please place a PDF file in the current directory to test the extraction.")
        return
    
    print(f"\nFound {len(existing_files)} sample file(s): {existing_files}")
    
    # Run examples if we have files
    if "example.pdf" in existing_files:
        example_basic_extraction()
        example_japanese_only_extraction()
        example_translation_pipeline_format()
        example_detailed_extraction()
    
    # Always show batch processing example
    example_batch_processing()
    
    if "problematic.pdf" in existing_files:
        example_error_handling()
    
    print("\n" + "=" * 40)
    print("Examples completed! Check the generated JSON files for extraction results.")

if __name__ == "__main__":
    main()