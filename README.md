# 🚀 PowerPoint Translation Pipeline (JA→EN)

A production-ready translation system for converting Japanese PowerPoint presentations to English while preserving layout, formatting, and visual elements.

## ✨ Features

### 🎯 **Production-Ready Translation**
- **Smart batch sizing**: Auto-optimizes API requests per model
- **Comprehensive logging**: Real-time progress with ETA estimates  
- **Robust error handling**: Auto-retry with intelligent backoff
- **Layout preservation**: Maintains original formatting and design

### 🧠 **AI-Powered Quality**
- **Style consistency**: Unified tone and terminology across slides
- **Content-aware processing**: Handles titles, bullets, tables differently
- **Expansion management**: Prevents text overflow with smart compression
- **Glossary integration**: Ensures consistent translation of key terms

### 📊 **Advanced Features**
- **Translation caching**: Avoids re-translating identical content
- **Bilingual output**: CSV mapping for quality assurance
- **Performance metrics**: Detailed audit reports and statistics
- **Webhook integration**: Real-time progress tracking (optional)

## 🚀 Quick Start

### Prerequisites
```bash
export OPENAI_API_KEY=your_key_here
```

### Basic Usage
```bash
# Production presets (recommended)
python scripts/translate_pptx_inplace.py \
  --in input.pptx \
  --out output_en.pptx \
  --model gpt-4o-2024-08-06

# Cost-optimized option
python scripts/translate_pptx_inplace.py \
  --in input.pptx \
  --out output_en.pptx \
  --model gpt-4o-mini
```

## 🎛️ Production Presets

| Preset | Model | Batch Size | Use Case |
|--------|-------|------------|----------|
| **Conservative** | `gpt-4o-2024-08-06` | 8-12 (auto) | Maximum reliability |
| **Balanced** | `gpt-4o-2024-08-06` | 10-14 (auto) | **Recommended** |
| **Cost-lean** | `gpt-4o-mini` | 12-16 (auto) | Good quality, lower cost |

*Batch sizes are automatically calculated based on content complexity and token limits.*

## 📋 Command Line Options

```bash
python scripts/translate_pptx_inplace.py [OPTIONS]

Required:
  --in INPUT.pptx          Input PowerPoint file
  --out OUTPUT.pptx        Output translated file

Optional:
  --model MODEL           AI model (default: auto-optimized)
  --batch N               Batch size (default: auto-calculated)
  --cache FILE            Translation cache (default: translation_cache.json)
  --glossary FILE         Terminology glossary (default: glossary.json)
  --slides RANGE          Process specific slides (e.g., "1-10")
  --style-preset PRESET   Style guide preset (gengo, minimal)
```

## 📁 Project Structure

```
├── scripts/
│   ├── translate_pptx_inplace.py  # Main translation engine
│   ├── style_checker.py           # Style consistency system
│   ├── eta.py                     # Progress estimation
│   ├── webhook_server.py          # Real-time progress tracking
│   └── audit_style.py            # Quality analysis
├── tools/
│   ├── derive_deck_tone.py       # Tone analysis
│   └── estimate_cost.py          # Cost estimation
├── inputs/                       # Source presentations
├── outputs/                      # Translated results
└── data/                        # Glossaries and configs
```

## 🔧 Advanced Configuration

### Custom Glossary
Create `glossary.json` for consistent terminology:
```json
{
  "株式会社": "Corporation",
  "取締役": "Director",
  "戦略": "Strategy"
}
```

### Style Consistency
Configure tone and style preferences:
```json
{
  "formality": "business_formal",
  "technical_terms": "preserve_english",
  "bullet_style": "concise_fragments"
}
```

### Webhook Progress Tracking
Run the webhook server for real-time updates:
```bash
# Terminal 1: Start webhook server
uvicorn scripts.webhook_server:app --port 8000

# Terminal 2: Run translation
python scripts/translate_pptx_inplace.py --in input.pptx --out output.pptx
```

## 📊 Output Files

Each translation run generates:

| File | Description |
|------|-------------|
| `output_en.pptx` | Translated presentation |
| `bilingual.csv` | Side-by-side translation mapping |
| `audit.json` | Translation statistics and metrics |
| `translation_cache.json` | Cached translations for efficiency |
| `translation.log` | Detailed execution log |

## 🛠️ System Architecture

### Smart Batch Processing
- **Token-aware sizing**: Calculates optimal batch sizes based on model limits
- **Dynamic adjustment**: Reduces batch size automatically on high retry rates
- **Content analysis**: Adjusts for complex content (tables, technical text)

### Style Consistency Engine
- **Multi-stage processing**: Pre-translation normalization → Translation → Post-processing
- **Authority corrections**: Deterministic style fixes based on diagnostics
- **Tone preservation**: Maintains consistent voice across the document

### Error Resilience
- **Progressive backoff**: 1s, 2s, 3s delays on retries
- **Graceful degradation**: Falls back to smaller batches on failures
- **Cache recovery**: Preserves work through interruptions

## 📈 Performance Optimization

### Batch Size Guidelines
- **gpt-4o models**: 8-14 items (10k token target)
- **gpt-4o-mini**: 12-18 items (8k token target)
- **Complex content**: Use lower end of ranges
- **Simple text**: Can use higher batch sizes

### Cost Management
- **Cache efficiency**: ~90% cache hit rate on re-runs
- **Model selection**: gpt-4o-mini offers 10x cost savings
- **Batch optimization**: Reduces API call overhead

## 🚨 Troubleshooting

### Common Issues

**High retry rates (>5%)**
- System automatically reduces batch size
- Check API key limits and quotas
- Consider using gpt-4o-mini for better stability

**Text overflow in slides**
- Enable PowerPoint's "Shrink text on overflow"
- Use style presets for more concise translations
- Adjust font sizes manually if needed

**Cache corruption**
- Delete `translation_cache.json` to reset
- Use `--cache new_cache.json` for fresh cache

### Debug Mode
```bash
# Enable verbose logging
export PYTHONPATH=scripts
python -u scripts/translate_pptx_inplace.py --in input.pptx --out output.pptx 2>&1 | tee debug.log
```

## 🔮 Future Enhancements

- **OCR integration**: Translate text in images
- **Multi-language support**: Beyond JA→EN
- **Real-time collaboration**: Shared translation sessions  
- **Template management**: Reusable style configurations
- **Quality scoring**: Automatic translation assessment

## 📄 License

MIT License - see LICENSE file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests and documentation
5. Submit a pull request

---

*Built with ❤️ for efficient, high-quality presentation translation.*