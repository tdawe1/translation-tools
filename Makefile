.PHONY: estimate derive-tone docx-ci
.PHONY: all clean test help

help:
	@echo "Targets: estimate, derive-tone, docx-ci, test, clean"

all: test

test:
	@echo "No tests wired yet."

clean:
	@./scripts/cleanup.sh aggressive

estimate:
	@./tools/estimate_cost.py inputs/68b42f175c652_f711fcda865b11f0b6cecace4a312dcf_en_final_offline_v2.pptx --pricing pricing.example.json --producer openai:gpt-5 --reviewer openai:gpt-5-mini --batch-size 16 --prefix-file ./scripts/translate_pptx_inplace.py --also anthropic:claude-sonnet-4 google:gemini-1.5-pro

derive-tone:
	@./tools/derive_deck_tone.py inputs/68b42f175c652_f711fcda865b11f0b6cecace4a312dcf_en_final_offline_v2.pptx

docx-ci:
	@echo "Running DOCX CI pipeline..."
	# Install additional dependencies if needed
	python -c "import pytest, jsonschema, docx" || pip install pytest jsonschema python-docx defusedxml
	# Run adapter tests
	PYTHONPATH=. python -m pytest tests/test_docx_adapter_basic.py -v
	# Run smoke test with dummy fixture
	@mkdir -p tests/fixtures
	@echo "Creating dummy fixture for smoke test..."
	@python scripts/create_dummy_docx.py tests/fixtures/dummy.docx
	@OPENAI_API_KEY=dummy python scripts/smoke_translate_docx.py --input tests/fixtures/dummy.docx --output /tmp/test_output.docx
	@rm -f /tmp/test_output.docx tests/fixtures/dummy.docx
