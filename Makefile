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
	@python test_cf5_validation.py
	@echo "DOCX CI validation completed"
