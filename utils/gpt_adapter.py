import os
import json
import time
import re
from typing import List, Dict, Any
from openai import OpenAI

def mock_translate(items: List[str]) -> List[str]:
    """Generate mock translations for offline testing."""
    return [f"Mock EN: {item[:20]}..." if len(item) > 20 else f"Mock EN: {item}" for item in items]

def _extract_json_array(s: str, expected_len: int) -> List[str] | None:
    """Extract JSON array from response string, tolerant to markdown fences."""
    s = re.sub(r"^```(?:json)?|```$", "", s.strip(), flags=re.M)
    try:
        data = json.loads(s)
        if isinstance(data, list) and len(data) >= expected_len:
            return [str(item) for item in data[:expected_len]]
    except json.JSONDecodeError:
        pass
    return None

class GPTAdapter:
    def __init__(self, api_key: str = None, base_url: str = None, primary_model: str = "gpt-5"):
        self.client = OpenAI(
            api_key=api_key or os.getenv("OPENAI_API_KEY"),
            base_url=base_url or os.getenv("OPENAI_BASE_URL", "")
        )
        self.primary_model = primary_model
        self.fallback_models = ["gpt-4o-2024-08-06", "gpt-4o-mini"]

    def batch_translate(
        self, 
        items: List[str], 
        glossary: Dict[str, str] = None, 
        offline: bool = False, 
        temperature: float = 0.6
    ) -> List[str]:
        """Translate batch of Japanese strings to English with fallback chain.
        
        Handles unsupported params by stripping them for GPT-5 (e.g., response_format, verbosity).
        Chunks large payloads (>10MB total input size).
        Falls back on errors like endpoint mismatches or size limits.
        """
        if offline:
            return mock_translate(items)

        if not items:
            return []

        # Simple style guide (in production, load from file)
        style_guide = (
            "Translate naturally to professional English. "
            "Preserve numbers, URLs, and structure. Use double quotes. "
            "Follow glossary terms exactly."
        )
        sys_prompt = (
            f"{style_guide}\n\n"
            f"GLOSSARY: {json.dumps(glossary or {}, ensure_ascii=False)}\n\n"
            "Return ONLY a JSON array of translated strings in the same order. "
            "No explanations or markdown."
        )

        models = [self.primary_model] + self.fallback_models
        total_input_size = sum(len(item.encode('utf-8')) for item in items)

        for model in models:
            try:
                if total_input_size > 10 * 1024 * 1024:  # >10MB, chunk input
                    return self._chunk_and_translate(model, items, sys_prompt, temperature)
                else:
                    return self._single_call_translate(model, items, sys_prompt, temperature)
            except Exception as e:
                error_str = str(e).lower()
                if any(keyword in error_str for keyword in ["endpoint", "mismatch", "size limit", "rate limit", "invalid parameter"]):
                    print(f"Falling back from {model} due to: {e}")
                    continue
                else:
                    # Non-recoverable error, raise
                    raise

        raise RuntimeError(f"All models failed to translate batch of {len(items)} items.")

    def _single_call_translate(self, model: str, items: List[str], sys_prompt: str, temperature: float) -> List[str]:
        user_payload = {
            "strings": items,
            "instructions": "Respond with ONLY the JSON array."
        }
        user_content = json.dumps(user_payload, ensure_ascii=False)

        content = self._call_api(model, sys_prompt, user_content, temperature)
        data = _extract_json_array(content, len(items))
        if data:
            return data
        raise ValueError("Failed to extract valid JSON array from response")

    def _chunk_and_translate(self, model: str, items: List[str], sys_prompt: str, temperature: float) -> List[str]:
        """Chunk large inputs and translate in parallel (simplified sequential for now)."""
        # Simple chunking: aim for ~3MB per chunk
        chunk_size_bytes = 3 * 1024 * 1024
        chunks = []
        current_chunk = []
        current_size = 0

        for item in items:
            item_bytes = len(item.encode('utf-8'))
            if current_size + item_bytes > chunk_size_bytes and current_chunk:
                chunks.append(current_chunk)
                current_chunk = [item]
                current_size = item_bytes
            else:
                current_chunk.append(item)
                current_size += item_bytes
        if current_chunk:
            chunks.append(current_chunk)

        all_translations = []
        for chunk in chunks:
            user_payload = {"strings": chunk, "instructions": "Respond with ONLY the JSON array."}
            user_content = json.dumps(user_payload, ensure_ascii=False)
            content = self._call_api(model, sys_prompt, user_content, temperature)
            chunk_data = _extract_json_array(content, len(chunk))
            if chunk_data:
                all_translations.extend(chunk_data)
            else:
                # Fallback: summarize output if too large, but for now raise
                raise ValueError("Chunk translation failed")

        # If output too large, could summarize here, but skip for simplicity
        return all_translations

    def _call_api(self, model: str, sys_prompt: str, user_content: str, temperature: float) -> str:
        """Make API call, stripping unsupported params for GPT-5 (e.g., no response_format, verbosity)."""
        use_responses = model.startswith("gpt-5")
        max_retries = 3

        for attempt in range(max_retries):
            try:
                if use_responses:
                    # GPT-5 Responses API: strip response_format and verbosity
                    input_data = [
                        {"role": "system", "content": [{"type": "input_text", "text": sys_prompt}]},
                        {"role": "user", "content": [{"type": "input_text", "text": user_content}]}
                    ]
                    kwargs = {
                        "model": model,
                        "input": input_data,
                        "temperature": temperature,
                        # No response_format={"type": "json"} - enforce via prompt
                        # No text={"verbosity": "low"}
                    }
                    resp = self.client.responses.create(**kwargs)
                    # Extract content
                    if hasattr(resp, "output_text"):
                        content = resp.output_text
                    elif hasattr(resp, "choices") and resp.choices:
                        content = resp.choices[0].message.content
                    elif hasattr(resp, "output") and resp.output:
                        content = resp.output[0].content[0].text
                    else:
                        content = str(resp)
                else:
                    # Chat Completions: use prompt for JSON, no response_format if unsupported
                    messages = [
                        {"role": "system", "content": sys_prompt},
                        {"role": "user", "content": user_content}
                    ]
                    kwargs = {
                        "model": model,
                        "messages": messages,
                        "temperature": temperature,
                        # Strip response_format - use prompt enforcement
                    }
                    resp = self.client.chat.completions.create(**kwargs)
                    content = resp.choices[0].message.content

                return content.strip()

            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                else:
                    raise

# For backward compatibility, provide function alias
def batch_translate(client=None, model="gpt-5", items=None, glossary=None, offline_mode=False, temperature=0.6):
    """Backward compatible function. Creates adapter if no client provided."""
    if client is None:
        adapter = GPTAdapter(primary_model=model)
        return adapter.batch_translate(items, glossary, offline_mode, temperature)
    else:
        # If client provided, fallback to simple call (no adapter logic)
        # This maintains compatibility without forcing refactor
        raise NotImplementedError("Direct client usage not implemented in adapter; use adapter for full features.")
