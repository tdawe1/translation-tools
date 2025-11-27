import os
import json
from openai import OpenAI, AsyncOpenAI
import asyncio

def _use_responses_api(model: str) -> bool:
    m = (model or "").lower()
    # Prefer Responses API for latest models like gpt-5 family
    return m.startswith("gpt-5") or os.getenv("OPENAI_USE_RESPONSES") == "1"

def make_array_schema(expected_len: int | None):
    """Build a strict JSON Schema for string arrays."""
    return {
        "name": "BatchArrayOfStrings",
        "schema": {
            "type": "array",
            "items": {"type": "string"},
            "minItems": 1
        },
        "strict": True
    }

def _responses_create(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    # OpenAI Responses API with GPT-5 reasoning model
    try:
        # Configure reasoning effort based on model - high for main translation, minimal for reviews
        if model.startswith("gpt-5-mini"):
            effort = "minimal"  # Fast reviewer
        else:
            effort = os.getenv("OPENAI_REASONING_EFFORT", "high")  # Deep thinking for translation

        resp = client.responses.create(
            model=model,
            input=[
                {"role": "system", "content": [{"type": "input_text", "text": sys_prompt}]},
                {"role": "user", "content": [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]}
            ],
            reasoning={"effort": effort},
            text={"verbosity": "low"},  # Concise responses, avoid chatty prose
            temperature=temperature,
            response_format={"type": "json"},
        )
        # New SDKs expose output_text; fall back if absent
        content = getattr(resp, "output_text", None)
        if not content:
            # Fallback to choices/message style if present
            if getattr(resp, "choices", None):
                content = resp.choices[0].message.content
        if not content and getattr(resp, "output", None):
            try:
                # Attempt to read the first text content
                content = resp.output[0].content[0].text
            except Exception:
                content = None
        return content.strip() if content else ""
    except Exception:
        raise

def _chat_create(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    """Sync version with response_format fallback."""
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
            response_format={"type": "json_object"},
        )
    except Exception:
        # Fallback: schema in prompt
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt + "\nReturn ONLY a JSON array."},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
        )
    return resp.choices[0].message.content.strip()

async def _chat_create_async(client, model: str, sys_prompt: str, user_payload: dict, temperature: float):
    """Async version with response_format fallback."""
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
            response_format={"type": "json_object"},
        )
    except Exception:
        # Fallback: schema in prompt
        resp = await client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": sys_prompt + "\nReturn ONLY a JSON array."},
                {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
            ],
            temperature=temperature,
        )
    return resp.choices[0].message.content.strip()

async def _responses_create_compat_async(aclient, *, model, input, temperature, json_schema, max_output_tokens):
    """Async Responses API wrapper with JSON schema fallback."""
    try:
        resp = await aclient.responses.create(
            model=model,
            input=input,
            temperature=temperature,
            max_output_tokens=max_output_tokens,
            response_format={"type": "json_schema", "json_schema": json_schema, "strict": True},
        )
    except TypeError as e:
        if "response_format" in str(e):
            # Fallback: inline schema in prompt
            schema_text = f"Return ONLY a valid JSON value matching this JSON Schema:\n{json.dumps(json_schema, indent=2)}"
            fallback_input = input.copy()
            if fallback_input and len(fallback_input) > 0:
                fallback_input[0]["content"] = schema_text + "\n\n" + fallback_input[0]["content"]

            resp = await aclient.responses.create(
                model=model,
                input=fallback_input,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
            )
        else:
            raise

    # Extract content from response
    content = getattr(resp, "output_text", None)
    if not content and getattr(resp, "output", None):
        try:
            content = resp.output[0].content[0].text
        except Exception:
            content = None
    if not content and getattr(resp, "choices", None):
        content = resp.choices[0].message.content

    return content.strip() if content else ""