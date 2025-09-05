.PHONY: estimate derive-tone
.PHONY: all clean test help

help:
	@echo "Targets: estimate, derive-tone, test, clean"

all: test

test:
	@echo "No tests wired yet."

clean:
	@rm -f deck_tone.json

estimate:
	@./tools/estimate_cost.py --pricing pricing.example.json --producer openai:gpt-5 --reviewer openai:gpt-5-mini --batch-size 16 --prefix-file ./scripts/translate_pptx_inplace.py --also anthropic:claude-sonnet-4 google:gemini-1.5-pro

derive-tone:
	@./tools/derive_deck_tone.py
