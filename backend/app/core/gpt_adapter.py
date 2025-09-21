import json
import os
from typing import List, Dict, Any, Optional
from openai import OpenAI, AsyncOpenAI

class GPTAdapter:
    """
    Central adapter for OpenAI GPT calls, handling GPT-5 compatibility
    and large payload chunking via summarization or splitting.
    """
    
    def __init__(self, api_key: str, base_url: Optional[str] = None):
        self.api_key = api_key
        self.base_url = base_url
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.async_client = AsyncOpenAI(api_key=api_key, base_url=base_url)
    
    def is_gpt5(self, model: str) -> bool:
        """Check if model is GPT-5 variant."""
        return model.lower().startswith('gpt-5')
    
    def _strip_unsupported_params(self, params: Dict[str, Any], model: str) -> Dict[str, Any]:
        """Strip unsupported parameters for GPT-5 models."""
        if not self.is_gpt5(model):
            return params.copy()
        
        cleaned = params.copy()
        # Remove response_format for GPT-5 (not supported)
        if 'response_format' in cleaned:
            del cleaned['response_format']
        
        # Remove verbosity from text param
        if 'text' in cleaned and isinstance(cleaned['text'], dict) and 'verbosity' in cleaned['text']:
            if 'verbosity' in cleaned['text']:
                del cleaned['text']['verbosity']
            if not cleaned['text']:
                del cleaned['text']
        
        # Add more stripping as needed
        unsupported = ['verbosity', 'reasoning_effort']  # Example
        for key in list(cleaned.keys()):
            if key in unsupported:
                del cleaned[key]
        
        return cleaned
    
    def _check_payload_size(self, payload: Dict[str, Any]) -> bool:
        """Check if payload exceeds 10MB."""
        try:
            payload_json = json.dumps(payload)
            return len(payload_json.encode('utf-8')) > 10 * 1024 * 1024
        except:
            return False
    
    async def _chunk_large_payload(self, model: str, messages: List[Dict[str, Any]], 
                                   params: Dict[str, Any], batch_key: str = 'strings') -> List[Any]:
        """Chunk large payloads by splitting batches and translating separately."""
        results = []
        try:
            # Assume last message is user with JSON payload
            user_msg = messages[-1]
            if isinstance(user_msg['content'], str):
                payload = json.loads(user_msg['content'])
                if batch_key in payload and isinstance(payload[batch_key], list):
                    items = payload[batch_key]
                    chunk_size = max(1, len(items) // 4)  # Split into 4 chunks max
                    for i in range(0, len(items), chunk_size):
                        chunk = items[i:i + chunk_size]
                        chunk_payload = payload.copy()
                        chunk_payload[batch_key] = chunk
                        chunk_messages = messages.copy()
                        chunk_messages[-1]['content'] = json.dumps(chunk_payload)
                        
                        chunk_params = self._strip_unsupported_params(params.copy(), model)
                        
                        response = await self.async_client.chat.completions.create(
                            model=model,
                            messages=chunk_messages,
                            **chunk_params
                        )
                        
                        content = response.choices[0].message.content
                        # Assume JSON array response
                        chunk_results = json.loads(content)
                        results.extend(chunk_results)
                    
                    # If needed, summarize combined results, but for translation, just concat
                    return results[:len(items)]  # Trim if extra
        except Exception:
            pass  # Fallback to normal call
        return []
    
    async def chat_completion(self, model: str, messages: List[Dict[str, Any]], 
                              **kwargs) -> Any:
        """Async chat completion with param stripping and chunking."""
        params = self._strip_unsupported_params(kwargs, model)
        
        payload = {
            "model": model,
            "messages": messages,
            **params
        }
        
        if self._check_payload_size(payload):
            # Chunk if large
            chunked = await self._chunk_large_payload(model, messages, params)
            if chunked:
                # Return simulated full response, adjust as needed
                return type('obj', (object,), {'choices': [{'message': {'content': json.dumps(chunked)}}]})()
        
        return await self.async_client.chat.completions.create(
            model=model,
            messages=messages,
            **params
        )
    
    def sync_chat_completion(self, model: str, messages: List[Dict[str, Any]], 
                             **kwargs) -> Any:
        """Sync chat completion."""
        params = self._strip_unsupported_params(kwargs, model)
        
        payload = {
            "model": model,
            "messages": messages,
            **params
        }
        
        if self._check_payload_size(payload):
            # For sync, raise or handle differently
            raise ValueError("Payload too large for sync call; use async")
        
        return self.client.chat.completions.create(
            model=model,
            messages=messages,
            **params
        )
    
    # Similar for responses.create if used
    async def responses_create(self, model: str, input: List[Dict], **kwargs) -> Any:
        """Async responses create with stripping."""
        params = self._strip_unsupported_params(kwargs, model)
        return await self.async_client.responses.create(
            model=model,
            input=input,
            **params
        )
    
    async def batch_translate(self, model: str, items: List[str], 
                              glossary: Dict[str, str] = None, 
                              offline_mode: bool = False,
                              temperature: float = 0.6) -> List[str]:
        """Batch translate items using adapter."""
        if offline_mode:
            return [f"Mock translation: {item[:20]}..." for item in items]
        
        if not glossary:
            glossary = {}
        
        # Build prompt similar to existing
        sys_prompt = "Translate the following Japanese strings to English naturally, preserving structure. Return ONLY a JSON array of translations in the same order."
        user_payload = {
            "strings": items,
            "glossary": glossary
        }
        messages = [
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)}
        ]
        
        response = await self.chat_completion(
            model=model,
            messages=messages,
            temperature=temperature,
            response_format={"type": "json_object"}
        )
        
        content = response.choices[0].message.content
        try:
            data = json.loads(content)
            return data.get("translations", data) if isinstance(data, dict) else data
        except:
            # Fallback extraction
            import re
            array_match = re.search(r'\[.*\]', content, re.DOTALL)
            if array_match:
                return json.loads(array_match.group())
            return [item for item in items]  # Fallback to original