.PHONY: estimate estimate-pdf derive-tone translate-pdf test-pdf translate-pptx test clean help clean-pdf test-all estimate-all translate-all clean-all setup verify-deps

help:
	@echo "Targets:"
	@echo "  setup           - Install all dependencies"
	@echo "  verify-deps     - Verify all dependencies are installed"
	@echo "  estimate        - Estimate translation cost for PPTX"
	@echo "  estimate-pdf    - Estimate translation cost for PDF"
	@echo "  derive-tone     - Derive tone/style from PPTX"
	@echo "  translate-pdf   - Translate PDF using the complete pipeline"
	@echo "  test-pdf        - Test PDF translation pipeline"
	@echo "  test-quality    - Test PDF quality metrics"
	@echo "  test-integration - Test PDF integration"
	@echo "  translate-pptx  - Translate PPTX using the standard pipeline"
	@echo "  test            - Run unit tests"
	@echo "  clean           - Clean up temporary files"
	@echo "  clean-pdf       - Clean up PDF translation artifacts"
	@echo ""
	@echo "Combined targets:"
	@echo "  estimate-all    - Estimate costs for both PPTX and PDF"
	@echo "  translate-all   - Translate both PPTX and PDF"
	@echo "  test-all        - Run all tests (PPTX + PDF)"
	@echo "  clean-all       - Clean all artifacts"

# Setup target to install all dependencies
setup:
	@echo "Installing all dependencies..."
	pip install -r requirements.txt
	pip install -r requirements_pdf.txt
	@echo "✅ Dependencies installed successfully"

# Verify dependencies
verify-deps:
	@echo "Verifying dependencies..."
	@python scripts/verify_pdf_deps.py

all: test

# Run all tests with proper PYTHONPATH
test:
	@echo "Running unit tests..."
	@PYTHONPATH=. python -m pytest tests/test_translate_pdf.py -v
	@echo "✅ PDF translation orchestrator tests passed"

# Run PDF-specific tests with proper PYTHONPATH
test-pdf:
	@echo "Testing PDF translation pipeline..."
	@PYTHONPATH=. python -m pytest tests/test_translate_pdf.py -v
	@echo "✅ PDF translation tests passed"

# Run tests in a clean environment
test-clean:
	@echo "Testing in clean environment..."
	@PYTHONPATH=. python tests/test_clean_environment.py
	@echo "✅ Clean environment tests passed"

# Run quality metrics tests (more reliable)
test-quality:
	@echo "Testing PDF quality metrics..."
	@PYTHONPATH=. python -m pytest tests/test_pdf_quality_metrics.py -v
	@echo "✅ PDF quality metrics tests passed"

# Run integration tests with proper PYTHONPATH
test-integration:
	@echo "Testing PDF integration..."
	@PYTHONPATH=. python -m pytest tests/test_pdf_integration.py -v; \
	exit_code=$?; \
	if [ ${exit_code} -eq 0 ]; then \
		echo "✅ PDF integration tests passed"; \
	else \
		echo "❌ PDF integration tests failed"; \
		exit 1; \
	fi

translate-pdf:
	@if [ -z "$(INPUT)" ] || [ -z "$(OUTPUT)" ]; then \
		echo "Usage: make translate-pdf INPUT=input.pdf OUTPUT=output.pdf [OPTIONS]"; \
		echo "Example: make translate-pdf INPUT=document.pdf OUTPUT=document_en.pdf MODEL=gpt-4o-mini"; \
		echo ""; \
		echo "Available options:"; \
		echo "  MODEL           - Translation model (default: gpt-4o-2024-08-06)"; \
		echo "  PAGES           - Page range (e.g., 1-10, 5)"; \
		echo "  GLOSSARY        - Path to glossary file"; \
		echo "  CACHE           - Cache file path (default: translation_cache.json)"; \
		echo "  CACHE_ONLY      - Use only cached translations (true/false)"; \
		echo "  OFFLINE         - Run in offline mode (true/false)"; \
		echo "  VERBOSE         - Enable verbose logging (true/false)"; \
		exit 1; \
	fi
	@echo "Translating PDF: $(INPUT) -> $(OUTPUT)"
	@python scripts/translate_pdf.py \
		--in "$(INPUT)" \
		--out "$(OUTPUT)" \
		--model "$(MODEL)" \
		--pages "$(PAGES)" \
		--glossary "$(GLOSSARY)" \
		--cache "$(CACHE)" \
		--cache-only "$(CACHE_ONLY)" \
		--offline "$(OFFLINE)" \
		--verbose "$(VERBOSE)"

translate-pptx:
	@if [ -z "$(INPUT)" ] || [ -z "$(OUTPUT)" ]; then \
		echo "Usage: make translate-pptx INPUT=input.pptx OUTPUT=output.pptx [OPTIONS]"; \
		echo "Example: make translate-pptx INPUT=presentation.pptx OUTPUT=presentation_en.pptx MODEL=gpt-4o-mini"; \
		exit 1; \
	fi
	@echo "Translating PPTX: $(INPUT) -> $(OUTPUT)"
	@python scripts/translate_pptx_inplace.py \
		--in "$(INPUT)" \
		--out "$(OUTPUT)" \
		--model "$(MODEL)"

clean:
	@./scripts/cleanup.sh aggressive

# PDF-specific targets
estimate-pdf:
	@if [ -z "$(PDF_INPUT)" ]; then \
		echo "Usage: make estimate-pdf PDF_INPUT=document.pdf [OPTIONS]"; \
		echo "Example: make estimate-pdf PDF_INPUT=document.pdf MODEL=openai:gpt-5 PAGES=1-10"; \
		echo ""; \
		echo "Available options:"; \
		echo "  MODEL           - Translation model (default: openai:gpt-5)"; \
		echo "  REVIEWER        - Reviewer model (default: openai:gpt-5-mini)"; \
		echo "  PAGES           - Page range to process (e.g., 1-10, 5)"; \
		echo "  BATCH_SIZE      - Blocks per request (default: 16)"; \
		echo "  PRICING        - Path to pricing JSON file"; \
		echo "  PREFIX_FILE     - Path to prefix file for caching calculations"; \
		echo "  ALSO            - Additional models to estimate"; \
		echo "  NO_CACHE        - Disable caching calculations (true/false)"; \
		echo "  VERBOSE         - Enable verbose output (true/false)"; \
		exit 1; \
	fi
	@echo "Estimating PDF translation cost: $(PDF_INPUT)"
	@python tools/estimate_cost_pdf.py "$(PDF_INPUT)" \
		--producer "$(MODEL)" \
		--reviewer "$(REVIEWER)" \
		--pages "$(PAGES)" \
		--batch-size "$(BATCH_SIZE)" \
		--pricing "$(PRICING)" \
		--prefix-file "$(PREFIX_FILE)" \
		$(if $(filter true,$(NO_CACHE)),--no-cache,) \
		$(if $(filter true,$(ANTHROPIC_CACHE_WRITE)),--anthropic-cache-write,) \
		$(if $(ALSO),$(foreach model,$(ALSO),--also "$(model)"),) \
		$(if $(filter true,$(VERBOSE)),--verbose,)

clean-pdf:
	@echo "Cleaning PDF translation artifacts..."
	@rm -f outputs/*.pdf outputs/*_pdf*.json outputs/*pdf*.csv outputs/*pdf*.log
	@rm -rf pdf_extraction.log
	@echo "✅ PDF artifacts cleaned"

# Combined targets
estimate-all: estimate estimate-pdf

translate-all: translate-pptx translate-pdf

test-all: test test-pdf test-quality test-integration

clean-all: clean clean-pdf

estimate:
	@./tools/estimate_cost.py inputs/68b42f175c652_f711fcda865b11f0b6cecace4a312dcf_en_final_offline_v2.pptx --pricing pricing.example.json --producer openai:gpt-5 --reviewer openai:gpt-5-mini --batch-size 16 --prefix-file ./scripts/translate_pptx_inplace.py --also anthropic:claude-sonnet-4 google:gemini-1.5-pro

derive-tone:
	@./tools/derive_deck_tone.py inputs/68b42f175c652_f711fcda865b11f0b6cecace4a312dcf_en_final_offline_v2.pptx