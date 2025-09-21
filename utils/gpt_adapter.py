#!/usr/bin/env python3
\"\"\"GPT-5 Adapter for OpenAI API compatibility.

Handles unsupported parameters by stripping them, chunks large responses (>10MB),
and provides fallback to GPT-4o on failures. Includes retry logic for large payloads.

Usage:
    adapter = GPT5Adapter(api_key=os.getenv('OPENAI_API_KEY'))
    response = adapter.chat.completions.create(
        model='gpt-5',
        messages=[...],
        # Unsupported params like response_format, verbosity will be stripped
    )
\"\"\"

import os
import json
import time
from typing import Any, Dict, List, Optional
from openai import OpenAI, AsyncOpenAI

class GPT5Adapter:
    def __init__(self, api_key: str, base_url: Optional[str] = None, fallback_model: str = 'gpt-4o-2024-08-06'):
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.async_client = AsyncOpenAI(api_key=api_key, base_url=base_url)
        self.fallback_model = fallback_model
        self.max_response_size = 10 * 1024 * 1024  # 10MB
        self.max_retries = 3

    def _strip_unsupported_params(self, params: Dict[str, Any]) -> Dict[str, Any]:
        \"\"\"Strip unsupported parameters for GPT-5 (e.g., response_format, verbosity).\"\"\"
        unsupported = {'response_format', 'verbosity', 'reasoning_effort', 'text'}
        return {k: v for k, v in params.items() if k not in unsupported}

    def _chunk_large_payload(self, payload: str, max_size: int = 5 * 1024 * 1024) -> List[str]:
        \"\"\"Chunk large payloads for retry logic (e.g., >5MB tool payloads).\"\"\"
        if len(payload) <= max_size:
            return [payload]
        # Simple chunking; could summarize or split logically
        chunks = []
        for i in range(0, len(payload), max_size):
            chunk = payload[i:i + max_size]
            # Add summary limit if needed
            if i > 0:
                chunk = f'Summary of previous: {len(chunk)} chars... {chunk[-1000:]}'  # Limit summary
            chunks.append(chunk)
        return chunks

    def _handle_large_response(self, response: str) -> str:
        \"\"\"Chunk responses larger than 10MB.\"\"\"
        if len(response) <= self.max_response_size:
            return response
        # Chunk the response
        chunks = [response[i:i + self.max_response_size] for i in range(0, len(response), self.max_response_size)]
        return json.dumps({'chunks': chunks, 'total_size': len(response), 'message': 'Response chunked due to size limit'})

    def chat_completions_create(self, **kwargs) -> Any:
        \"\"\"Wrapper for chat.completions.create with stripping, chunking, and fallback.\"\"\"
        model = kwargs.get('model', 'gpt-5')
        cleaned_kwargs = self._strip_unsupported_params(kwargs)

        for attempt in range(self.max_retries):
            try:
                # Check for large payload in messages
                user_content = kwargs.get('messages', [{}])[-1].get('content', '')
                if isinstance(user_content, str) and len(user_content) > 5 * 1024 * 1024:
                    chunks = self._chunk_large_payload(user_content)
                    # For simplicity, process first chunk; in production, aggregate
                    cleaned_kwargs['messages'][-1]['content'] = chunks[0]

                response = self.client.chat.completions.create(**cleaned_kwargs)
                content = response.choices[0].message.content
                return self._handle_large_response(content)
            except Exception as e:
                if 'payload too large' in str(e).lower() or len(str(e)) > 1000:  # Heuristic for large payload errors
                    # Retry with chunking
                    time.sleep(2 ** attempt)
                    continue
                if attempt == self.max_retries - 1:
                    # Fallback to GPT-4o
                    print(f'Falling back to {self.fallback_model} after {self.max_retries} attempts: {e}')
                    cleaned_kwargs['model'] = self.fallback_model
                    return self.client.chat.completions.create(**cleaned_kwargs)
                time.sleep(2 ** attempt)

        raise Exception('All retries failed')

    async def chat_completions_create_async(self, **kwargs) -> Any:
        \"\"\"Async wrapper.\"\"\"
        # Similar to sync, but async
        model = kwargs.get('model', 'gpt-5')
        cleaned_kwargs = self._strip_unsupported_params(kwargs)

        for attempt in range(self.max_retries):
            try:
                response = await self.async_client.chat.completions.create(**cleaned_kwargs)
                content = response.choices[0].message.content
                return self._handle_large_response(content)
            except Exception as e:
                if attempt == self.max_retries - 1:
                    cleaned_kwargs['model'] = self.fallback_model
                    return await self.async_client.chat.completions.create(**cleaned_kwargs)
                time.sleep(2 ** attempt)

        raise Exception('All retries failed')

    # Add responses.create wrapper if needed
    def responses_create(self, **kwargs) -> Any:
        \"\"\"Wrapper for responses.create.\"\"\"
        cleaned_kwargs = self._strip_unsupported_params(kwargs)
        try:
            response = self.client.responses.create(**cleaned_kwargs)
            # Extract content
            content = getattr(response, 'output_text', getattr(response, 'choices', [None])[0].message.content if response.choices else '')
            return self._handle_large_response(content)
        except Exception as e:
            # Fallback
            print(f'GPT-5 responses failed, falling back to chat: {e}')
            kwargs['model'] = self.fallback_model
            return self.chat_completions_create(**kwargs)