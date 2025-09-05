.PHONY: estimate

estimate:
	@echo "Estimating translation cost..."
	@python tools/estimate_cost.py --producer openai:gpt-5 --reviewer openai:gpt-5-mini --batch-size 16 --also openai:gpt-4o google:gemini-1.5-pro
