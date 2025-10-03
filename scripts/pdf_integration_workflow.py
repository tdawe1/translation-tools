#!/usr/bin/env python3
"""
Integration script demonstrating PDF extraction with the existing translation pipeline.

This script shows how to use extract_pdf.py with the existing translate_pptx_inplace.py
functionality for a complete PDF translation workflow.
"""

import json
import os
import sys
import argparse
from pathlib import Path

# Add scripts to path
sys.path.insert(0, str(Path(__file__).parent / "scripts"))

def extract_pdf_workflow(pdf_path: str, output_dir: str = "output"):
    """Complete PDF extraction and translation workflow."""
    print(f"📄 Starting PDF extraction workflow for: {pdf_path}")
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Step 1: Extract text from PDF
    print("🔍 Step 1: Extracting text from PDF...")
    try:
        from extract_pdf import PDFExtractor
        
        extractor = PDFExtractor(use_fallback=True, min_confidence=0.8)
        result = extractor.extract_text_blocks(pdf_path, detailed=True)
        
        # Convert to translation format
        translation_data = extractor.to_translation_format(result)
        
        # Save extraction results
        extraction_file = os.path.join(output_dir, "pdf_extraction.json")
        with open(extraction_file, 'w', encoding='utf-8') as f:
            json.dump(translation_data, f, ensure_ascii=False, indent=2)
        
        print(f"✅ Extracted {len(translation_data['japanese_texts'])} Japanese text blocks")
        print(f"📊 Results saved to: {extraction_file}")
        
    except ImportError as e:
        print(f"❌ PDF extraction failed: {e}")
        print("💡 Install required dependencies: pip install PyMuPDF pdfplumber")
        return None
    
    return translation_data

def prepare_for_translation(translation_data: dict, output_dir: str = "output"):
    """Prepare extracted data for translation pipeline."""
    print("\n🔄 Step 2: Preparing for translation...")
    
    # Create a simple cache file format compatible with existing pipeline
    cache_file = os.path.join(output_dir, "pdf_translation_cache.json")
    
    # Extract unique Japanese texts
    unique_texts = translation_data['unique_texts']
    
    # Create cache entries (empty for now, will be filled by translation)
    cache = {}
    for text in unique_texts:
        cache[text] = ""  # Empty translation placeholder
    
    # Save cache file
    with open(cache_file, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    
    print(f"📝 Prepared {len(unique_texts)} unique texts for translation")
    print(f"💾 Cache file created: {cache_file}")
    
    return cache, unique_texts

def simulate_translation(unique_texts: list, output_dir: str = "output"):
    """Simulate translation process (placeholder for actual translation)."""
    print("\n🌐 Step 3: Translating text (simulation)...")
    
    # Load existing cache if available
    cache_file = os.path.join(output_dir, "pdf_translation_cache.json")
    if os.path.exists(cache_file):
        with open(cache_file, 'r', encoding='utf-8') as f:
            cache = json.load(f)
    else:
        cache = {}
    
    # Simulate translation (in real scenario, use translate_batch)
    translated_count = 0
    for text in unique_texts:
        if text not in cache or not cache[text]:
            # Mock translation for demonstration
            mock_translation = f"[EN] {text[:20]}..." if len(text) > 20 else f"[EN] {text}"
            cache[text] = mock_translation
            translated_count += 1
    
    # Save updated cache
    with open(cache_file, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    
    print(f"✅ Translated {translated_count} texts")
    print(f"💾 Updated cache: {cache_file}")
    
    return cache

def create_bilingual_output(translation_data: dict, cache: dict, output_dir: str = "output"):
    """Create bilingual output for review."""
    print("\n📋 Step 4: Creating bilingual output...")
    
    bilingual_file = os.path.join(output_dir, "pdf_bilingual.csv")
    
    # Create CSV with original and translated text
    with open(bilingual_file, 'w', encoding='utf-8', newline='') as f:
        import csv
        writer = csv.writer(f)
        writer.writerow(["Page", "Block ID", "Japanese", "English", "Block Type", "Confidence"])
        
        for text_info in translation_data['text_mapping'].values():
            japanese_text = next(
                (text for text, mapping in translation_data['text_mapping'].items() 
                 if mapping['block_id'] == text_info['block_id']), 
                None
            )
            
            if japanese_text:
                english_text = cache.get(japanese_text, "")
                writer.writerow([
                    text_info['page'],
                    text_info['block_id'],
                    japanese_text,
                    english_text,
                    text_info['block_type'],
                    text_info['confidence']
                ])
    
    print(f"📄 Bilingual CSV created: {bilingual_file}")
    return bilingual_file

def generate_summary_report(translation_data: dict, output_dir: str = "output"):
    """Generate summary report of the extraction process."""
    print("\n📊 Step 5: Generating summary report...")
    
    report_file = os.path.join(output_dir, "pdf_extraction_report.json")
    
    # Calculate statistics
    total_pages = len(translation_data['layout_info']['pages'])
    total_texts = len(translation_data['japanese_texts'])
    unique_texts = len(translation_data['unique_texts'])
    
    # Analyze by block type
    block_types = {}
    for text_info in translation_data['text_mapping'].values():
        block_type = text_info['block_type']
        block_types[block_type] = block_types.get(block_type, 0) + 1
    
    # Calculate average confidence
    confidences = [info['confidence'] for info in translation_data['text_mapping'].values()]
    avg_confidence = sum(confidences) / len(confidences) if confidences else 0
    
    report = {
        "extraction_summary": {
            "source_file": translation_data['source_file'],
            "total_pages": total_pages,
            "total_text_instances": total_texts,
            "unique_texts": unique_texts,
            "average_confidence": round(avg_confidence, 3),
            "extraction_methods": list(set(
                page['extraction_method'] 
                for page in translation_data['layout_info']['pages']
            ))
        },
        "block_type_distribution": block_types,
        "layout_info": translation_data['layout_info'],
        "recommendations": generate_recommendations(block_types, avg_confidence)
    }
    
    with open(report_file, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    
    print(f"📈 Summary report generated: {report_file}")
    return report

def generate_recommendations(block_types: dict, avg_confidence: float) -> list:
    """Generate recommendations based on extraction results."""
    recommendations = []
    
    # Confidence-based recommendations
    if avg_confidence < 0.8:
        recommendations.append("Consider lowering confidence threshold or checking PDF quality")
    
    # Block type-based recommendations
    if block_types.get('table', 0) > 0:
        recommendations.append("Tables detected - consider specialized table extraction for better accuracy")
    
    if block_types.get('title', 0) > 0:
        recommendations.append("Titles detected - ensure consistent terminology in glossary")
    
    # Volume-based recommendations
    total_blocks = sum(block_types.values())
    if total_blocks > 100:
        recommendations.append("Large document - consider processing in batches")
    
    return recommendations

def main():
    """Main workflow function."""
    parser = argparse.ArgumentParser(
        description="Complete PDF extraction and translation workflow",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  python scripts/pdf_integration_workflow.py --input document.pdf
  python scripts/pdf_integration_workflow.py -i presentation.pdf -o results/
  python scripts/pdf_integration_workflow.py --input report.pdf --simulate-only
        """
    )
    
    parser.add_argument('--input', '-i', required=True,
                       help='Input PDF file path')
    parser.add_argument('--output', '-o', default='output',
                       help='Output directory (default: output)')
    parser.add_argument('--simulate-only', action='store_true',
                       help='Only simulate translation, don\'t apply changes')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose output')
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.input):
        print(f"❌ Input file not found: {args.input}")
        sys.exit(1)
    
    print("🚀 PDF Translation Pipeline Integration")
    print("=" * 50)
    
    try:
        # Step 1: Extract text from PDF
        translation_data = extract_pdf_workflow(args.input, args.output)
        if not translation_data:
            sys.exit(1)
        
        # Step 2: Prepare for translation
        cache, unique_texts = prepare_for_translation(translation_data, args.output)
        
        # Step 3: Translate (simulate)
        cache = simulate_translation(unique_texts, args.output)
        
        # Step 4: Create bilingual output
        bilingual_file = create_bilingual_output(translation_data, cache, args.output)
        
        # Step 5: Generate summary report
        report = generate_summary_report(translation_data, args.output)
        
        # Print final summary
        print("\n" + "=" * 50)
        print("🎉 Workflow Complete!")
        print(f"📁 Output directory: {args.output}")
        print(f"📄 Bilingual CSV: {bilingual_file}")
        print(f"📊 Summary report: {os.path.join(args.output, 'pdf_extraction_report.json')}")
        
        if args.verbose:
            print("\n📋 Detailed Results:")
            print(f"  Pages processed: {report['extraction_summary']['total_pages']}")
            print(f"  Text blocks found: {report['extraction_summary']['total_text_instances']}")
            print(f"  Unique texts: {report['extraction_summary']['unique_texts']}")
            print(f"  Average confidence: {report['extraction_summary']['average_confidence']:.3f}")
            print(f"  Block types: {list(report['block_type_distribution'].keys())}")
            
            if report['recommendations']:
                print("\n💡 Recommendations:")
                for rec in report['recommendations']:
                    print(f"  • {rec}")
        
    except Exception as e:
        print(f"❌ Workflow failed: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()