.PHONY: estimate derive-tone

estimate:
	@echo "Estimating translation cost..."
	@python tools/estimate_cost.py --producer openai:gpt-5 --reviewer openai:gpt-5-mini --batch-size 16 --also openai:gpt-4o google:gemini-1.5-pro

derive-tone:
	@echo "Deriving deck tone fingerprint..."
	@python tools/derive_deck_tone.py
