#!/bin/bash

# Demo script for PDF translation pipeline
# Shows how to use the complete PDF translation system

set -e

echo "=== PDF Translation Pipeline Demo ==="
echo ""

# Check if we have the required components
echo "🔍 Checking system requirements..."

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 not found"
    exit 1
fi
echo "✅ Python 3: $(python3 --version)"

# Check OpenAI API key
if [ -z "$OPENAI_API_KEY" ]; then
    echo "⚠️  OPENAI_API_KEY not set - will use offline mode"
    OFFLINE_MODE=true
else
    echo "✅ OPENAI_API_KEY: configured"
    OFFLINE_MODE=false
fi

# Check PDF translation script
if [ ! -f "scripts/translate_pdf.py" ]; then
    echo "❌ PDF translation script not found"
    exit 1
fi
echo "✅ PDF translation script: found"

# Check PDF translation components
COMPONENTS=("extract_pdf.py" "pdf_layout_engine.py" "apply_pdf_translation.py" "audit_pdf.py")
for component in "${COMPONENTS[@]}"; do
    if [ -f "scripts/$component" ]; then
        echo "✅ Component: $component"
    else
        echo "❌ Component: $component - not found"
        exit 1
    fi
done

echo ""
echo "🚀 Starting PDF translation demo..."
echo ""

# Create demo directory
DEMO_DIR="pdf_demo"
mkdir -p "$DEMO_DIR"

# Check if we have a sample PDF
SAMPLE_PDF="$DEMO_DIR/sample_japanese.pdf"
if [ ! -f "$SAMPLE_PDF" ]; then
    echo "📝 Creating sample PDF for demo..."
    
    # Create a simple text file with Japanese content
    cat > "$DEMO_DIR/sample_japanese.txt" << 'EOF'
テスト文書

これは日本語のテスト文書です。
PDF翻訳パイプラインのデモンストレーション用に作成されました。

主な機能：
- 日本語から英語への翻訳
- レイアウトの保持
- キャッシュシステムの共有
- 品質監査レポート

この文書は、PDF翻訳システムの能力を示すために使用されます。

注意：これはデモ用のサンプルテキストです。
実際の翻訳品質は、使用するモデルや内容によって異なります。
EOF

    # Try to create a PDF using available tools
    if command -v pandoc &> /dev/null; then
        pandoc "$DEMO_DIR/sample_japanese.txt" -o "$SAMPLE_PDF"
        echo "✅ Created sample PDF using pandoc"
    elif command -v textutil &> /dev/null; then
        # macOS textutil
        textutil -convert pdf "$DEMO_DIR/sample_japanese.txt" -output "$SAMPLE_PDF"
        echo "✅ Created sample PDF using textutil"
    else
        echo "⚠️  Cannot create PDF - please provide a Japanese PDF file in $DEMO_DIR/sample_japanese.pdf"
        echo "   Skipping translation demo..."
        exit 0
    fi
fi

# Run translation demo
echo ""
echo "🔄 Running translation demo..."
echo ""

OUTPUT_PDF="$DEMO_DIR/sample_english.pdf"

if [ "$OFFLINE_MODE" = true ]; then
    echo "📝 Running in OFFLINE mode (no API calls)..."
    python3 scripts/translate_pdf.py \
        --in "$SAMPLE_PDF" \
        --out "$OUTPUT_PDF" \
        --offline \
        --verbose
else
    echo "📝 Running with API translation..."
    python3 scripts/translate_pdf.py \
        --in "$SAMPLE_PDF" \
        --out "$OUTPUT_PDF" \
        --model gpt-4o-mini \
        --verbose
fi

echo ""
echo "📊 Translation Results:"
echo ""

# Check if output files were created
if [ -f "$OUTPUT_PDF" ]; then
    echo "✅ Translated PDF: $OUTPUT_PDF"
    file_size=$(stat -c%s "$OUTPUT_PDF" 2>/dev/null || stat -f%z "$OUTPUT_PDF" 2>/dev/null || echo "unknown")
    echo "   Size: $file_size bytes"
else
    echo "❌ Translated PDF not created"
    exit 1
fi

# Check bilingual CSV
BILINGUAL_CSV="${OUTPUT_PDF%.pdf}_bilingual.csv"
if [ -f "$BILINGUAL_CSV" ]; then
    echo "✅ Bilingual CSV: $BILINGUAL_CSV"
    line_count=$(wc -l < "$BILINGUAL_CSV" 2>/dev/null || echo "unknown")
    echo "   Lines: $line_count"
else
    echo "⚠️  Bilingual CSV not created"
fi

# Check audit report
AUDIT_JSON="${OUTPUT_PDF%.pdf}_audit.json"
if [ -f "$AUDIT_JSON" ]; then
    echo "✅ Audit Report: $AUDIT_JSON"
    
    # Try to parse and show key metrics
    if command -v jq &> /dev/null; then
        echo "   Quality Score: $(jq -r '.quality_assessment.overall_quality_score' "$AUDIT_JSON" 2>/dev/null || echo "N/A")"
        echo "   Residual JP: $(jq -r '.quality_assessment.residual_japanese_count' "$AUDIT_JSON" 2>/dev/null || echo "N/A") chars"
    fi
else
    echo "⚠️  Audit report not created"
fi

# Check cache
CACHE_FILE="translation_cache.json"
if [ -f "$CACHE_FILE" ]; then
    cache_entries=$(grep -c '"[^"]*":' "$CACHE_FILE" 2>/dev/null || echo "0")
    echo "✅ Translation Cache: $CACHE_FILE ($cache_entries entries)"
else
    echo "⚠️  Cache file not found"
fi

echo ""
echo "🎯 Demo completed successfully!"
echo ""
echo "Generated files:"
echo "  - $OUTPUT_PDF"
if [ -f "$BILINGUAL_CSV" ]; then echo "  - $BILINGUAL_CSV"; fi
if [ -f "$AUDIT_JSON" ]; then echo "  - $AUDIT_JSON"; fi

echo ""
echo "📖 To learn more:"
echo "  - Read: docs/PDF_TRANSLATION_README.md"
echo "  - Try: make translate-pdf INPUT=your.pdf OUTPUT=your_en.pdf"
echo "  - Test: make test-pdf"

echo ""
echo "💡 Tips for production use:"
echo "  - Use gpt-4o-2024-08-06 for best quality"
echo "  - Create a glossary.json for consistent terminology"
echo "  - Use --pages for large documents"
echo "  - Check audit reports for quality assurance"