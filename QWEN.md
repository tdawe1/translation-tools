# Document Translation Pipeline (JA→EN) - Context for Qwen

## Project Overview

This project is a sophisticated translation pipeline designed to convert Japanese documents (primarily PowerPoint presentations and PDFs) into English while meticulously preserving the original document's layout, formatting, and visual elements. It's structured as a production-ready system with a strong emphasis on quality, efficiency, and automation.

The core functionality revolves around:
1.  Extracting text content from documents.
2.  Leveraging AI (primarily OpenAI models like GPT-4o/4o-mini) for high-quality translation.
3.  Implementing intelligent caching to avoid re-translating identical content, reducing costs.
4.  Performing back-projection to apply translations directly onto the original document structure, ensuring layout integrity.
5.  Incorporating layout adjustment techniques (like font scaling) to handle text expansion during translation.
6.  Providing comprehensive auditing and quality assurance tools.

## Key Technologies

*   **Python**: The primary programming language for all scripts and tools.
*   **OpenAI API**: Used for the core machine translation engine.
*   **PyMuPDF (fitz)**: The core library for PDF manipulation, including text extraction and replacement for PDF back-projection.
*   **XML Parsing (Python standard library `xml.etree.ElementTree`)**: Used for direct manipulation of PowerPoint (PPTX) XML structures for text replacement.
*   **Make**: Used for defining and running common command-line workflows via a `Makefile`.

## Project Structure

The project root contains several key directories and files:

*   `scripts/`: Contains the core implementation scripts for translation, extraction, back-projection, and related utilities.
    *   `translate_pptx_inplace.py`: The main orchestrator for PPTX translation.
    *   `translate_pdf.py`: The main orchestrator for PDF translation.
    *   `apply_pdf_translation.py`: The core PDF back-projector implementation.
    *   `extract_pdf.py`, `pdf_layout_engine.py`, `audit_pdf.py`: Supporting components for the PDF pipeline.
    *   Other helper scripts for style checking, auditing, etc.
*   `tools/`: Contains utility scripts for tasks like cost estimation (`estimate_cost.py`, `estimate_cost_pdf.py`) and tone derivation (`derive_deck_tone.py`).
*   `tests/`: Houses the project's test suite.
*   `inputs/`: Intended directory for source documents to be translated.
*   `outputs/`: Directory where translated documents and artifacts are saved.
*   `data/`: Likely contains configuration files like glossaries (`glossary.json`) and pricing data (`pricing.json`).
*   `requirements.txt` / `requirements_pdf.txt`: Lists Python dependencies.
*   `Makefile`: Defines standard commands for translation, estimation, testing, and cleanup.
*   `README.md`: The main project documentation.
*   `IMPLEMENTATION_SUMMARY.md`, `PDF_TRANSLATION_PLAN.md`: Detailed documentation on specific components, especially PDF handling.
*   `translation_cache.json` (likely): The default cache file storing previously translated segments.
*   `CLAUDE.md`: A log of interactions or notes, possibly from a previous AI assistant.

## Core Workflows

### PPTX Translation

1.  **Command**: `python scripts/translate_pptx_inplace.py --in <input.pptx> --out <output.pptx> [--model MODEL]`
2.  **Process**:
    *   Parses the PPTX file's internal XML.
    *   Extracts Japanese text blocks.
    *   Queries the AI model (or cache) for translations in optimized batches.
    *   Applies the English translations back into the PPTX XML, preserving formatting.
    *   Performs layout adjustments (e.g., font scaling via `normAutofit`) to handle text expansion.
    *   Generates output files: translated PPTX, bilingual CSV, audit JSON, updated cache.

### PDF Translation

1.  **Command**: `python scripts/translate_pdf.py --in <input.pdf> --out <output.pdf> [--model MODEL]` or `make translate-pdf INPUT=<input.pdf> OUTPUT=<output.pdf>`
2.  **Process**:
    *   Uses `PyMuPDF` to extract Japanese text blocks along with their precise positions and formatting.
    *   Reuses the same AI translation and caching logic as the PPTX pipeline.
    *   Applies the English translations back onto the PDF at the original text positions using `PyMuPDF`, preserving fonts and colors.
    *   Adjusts font sizes to accommodate text expansion.
    *   Generates output files: translated PDF, bilingual CSV, audit JSON, updated cache.

### Cost Estimation

*   **PPTX**: `python tools/estimate_cost.py <input.pptx>`
*   **PDF**: `python tools/estimate_cost_pdf.py <input.pdf>` or `make estimate-pdf PDF_INPUT=<input.pdf>`

Estimates the cost of translating a document based on its content and the selected AI model.

## Building, Running, and Testing

### Prerequisites

*   Python 3.x
*   Required Python packages listed in `requirements.txt` (install via `pip install -r requirements.txt`).
*   An OpenAI API key set as the environment variable `OPENAI_API_KEY`.
*   For PDF processing, `PyMuPDF` is essential.

### Main Commands (via Makefile)

*   `make translate-pptx INPUT=<file> OUTPUT=<file>`: Translate a PPTX file.
*   `make translate-pdf INPUT=<file> OUTPUT=<file>`: Translate a PDF file.
*   `make estimate-pdf PDF_INPUT=<file>`: Estimate the cost of translating a PDF.
*   `make test`: Run the main test suite.
*   `make test-pdf`: Run PDF-specific tests.
*   `make clean`: Clean up temporary files.

### Direct Script Execution

Most functionality is available by running the Python scripts in `scripts/` and `tools/` directly with appropriate command-line arguments. Refer to the `README.md` or the script's help (`--help`) for details.

## Development Conventions

*   **Caching**: Translations are cached in `translation_cache.json` by default to prevent redundant API calls. The cache key is typically the normalized Japanese text.
*   **Glossaries**: Custom term translations can be provided via `glossary.json`.
*   **Style Consistency**: Tools like `style_checker.py` are used to ensure consistent translation style.
*   **Auditing**: Scripts like `audit_translated_only.py` are used to verify translation quality by checking for residual untranslated Japanese characters.
*   **Modularity**: The system is designed in modular components (extractor, translator, back-projector) to allow for future expansion to other document types (as outlined in the roadmap).

## Roadmap & Future Enhancements

According to the `README.md`, the project has aspirations for:
*   A zero-touch, fully automated pipeline triggered by file drops (e.g., Google Drive).
*   Supporting newer models like GPT-5.
*   Expanding to other document formats (DOCX, XLSX, Markdown) using a shared Document Abstraction Layer (DAL).
*   Implementing shared translation caches and better configuration defaults.
