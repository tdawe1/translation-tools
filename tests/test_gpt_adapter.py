import pytest
import json
from unittest.mock import Mock, patch, MagicMock, AsyncMock
from openai import OpenAI, AsyncOpenAI

from backend.app.core.gpt_adapter import GPTAdapter

@pytest.fixture
def adapter():
    return GPTAdapter(api_key="fake_key")

def test_is_gpt5(adapter):
    assert adapter.is_gpt5("gpt-5-preview") is True
    assert adapter.is_gpt5("gpt-4o") is False
    assert adapter.is_gpt5("gpt-5-mini") is True
    assert adapter.is_gpt5("other-model") is False

def test_strip_unsupported_params(adapter):
    params = {
        "response_format": {"type": "json_object"},
        "text": {"verbosity": "low"},
        "temperature": 0.6,
        "verbosity": "high",
        "reasoning_effort": "high"
    }
    
    # For GPT-5
    cleaned = adapter._strip_unsupported_params(params, "gpt-5")
    assert "response_format" not in cleaned
    assert "text" not in cleaned  # Since empty after del
    assert "verbosity" not in cleaned
    assert "reasoning_effort" not in cleaned
    assert cleaned["temperature"] == 0.6
    
    # For non-GPT-5, preserve
    preserved = adapter._strip_unsupported_params(params, "gpt-4o")
    assert "response_format" in preserved
    assert preserved["text"]["verbosity"] == "low"
    assert "temperature" in preserved

def test_check_payload_size(adapter):
    small_payload = {"model": "gpt-5", "messages": [{"role": "user", "content": "small"}]}
    assert adapter._check_payload_size(small_payload) is False
    
    # Simulate large
    with patch('json.dumps') as mock_dumps:
        mock_dumps.return_value = 'a' * (11 * 1024 * 1024)
        large_payload = {"model": "gpt-5", "messages": [{"role": "user", "content": "large"}]}
        assert adapter._check_payload_size(large_payload) is True

@pytest.mark.asyncio
@patch('backend.app.core.gpt_adapter.AsyncOpenAI')
async def test_chat_completion_stripping(mock_async_openai, adapter):
    mock_client = AsyncMock(spec=AsyncOpenAI)
    mock_response = MagicMock()
    mock_response.choices = [MagicMock(message=MagicMock(content='{"result": "ok"}'))]
    mock_client.chat.completions.create.return_value = mock_response
    mock_async_openai.return_value = mock_client
    
    messages = [{"role": "user", "content": "test"}]
    result = await adapter.chat_completion("gpt-5", messages, 
                                           response_format={"type": "json"}, 
                                           text={"verbosity": "low"},
                                           temperature=0.6)
    
    mock_client.chat.completions.create.assert_called_once()
    call_args = mock_client.chat.completions.create.call_args[1]
    assert "response_format" not in call_args
    assert call_args["temperature"] == 0.6
    assert result == mock_response

@pytest.mark.asyncio
@patch('backend.app.core.gpt_adapter.AsyncOpenAI')
async def test_chat_completion_large_payload_chunking(mock_async_openai, adapter):
    mock_client = AsyncMock(spec=AsyncOpenAI)
    mock_response1 = MagicMock()
    mock_response1.choices = [MagicMock(message=MagicMock(content='["Hello"]'))]
    mock_response2 = MagicMock()
    mock_response2.choices = [MagicMock(message=MagicMock(content='["World"]'))]
    mock_client.chat.completions.create.side_effect = [mock_response1, mock_response2]
    mock_async_openai.return_value = mock_client
    
    # Large batch
    items = ["こんにちは", "世界", "テスト"] * 10  # Assume large
    user_payload = {"strings": items}
    messages = [
        {"role": "system", "content": "Translate"},
        {"role": "user", "content": json.dumps(user_payload)}
    ]
    
    # Mock size check to true
    with patch.object(adapter, '_check_payload_size', return_value=True):
        result = await adapter.chat_completion("gpt-5", messages, temperature=0.6)
    
    assert mock_client.chat.completions.create.call_count >= 2  # Chunked calls
    assert result.choices[0].message.content == '["Hello", "World"]'  # Simulated concat

@pytest.mark.asyncio
async def test_batch_translate_offline(adapter):
    items = ["こんにちは", "世界"]
    result = await adapter.batch_translate("gpt-4o", items, offline_mode=True)
    assert len(result) == 2
    for r in result:
        assert r.startswith("Mock translation:")

@pytest.mark.asyncio
@patch.object(GPTAdapter, 'chat_completion')
async def test_batch_translate_online(mock_chat, adapter):
    items = ["こんにちは", "世界"]
    mock_response = MagicMock()
    mock_response.choices = [MagicMock(message=MagicMock(content='["Hello", "World"]'))]
    mock_chat.return_value = mock_response
    
    result = await adapter.batch_translate("gpt-4o", items, temperature=0.7)
    
    mock_chat.assert_called_once()
    assert result == ["Hello", "World"]

@pytest.mark.asyncio
@patch.object(GPTAdapter, 'chat_completion')
async def test_batch_translate_error_handling(mock_chat, adapter):
    items = ["test"]
    mock_chat.side_effect = Exception("API Error")
    
    result = await adapter.batch_translate("gpt-4o", items, temperature=0.6)
    assert result == ["test"]  # Fallback to original

@pytest.mark.asyncio
@patch('backend.app.core.gpt_adapter.AsyncOpenAI')
async def test_responses_create_stripping(mock_async_openai, adapter):
    mock_client = AsyncMock(spec=AsyncOpenAI)
    mock_response = MagicMock()
    mock_client.responses.create.return_value = mock_response
    mock_async_openai.return_value = mock_client
    
    input_data = [{"role": "user", "content": "test"}]
    await adapter.responses_create("gpt-5", input_data, text={"verbosity": "low"})
    
    mock_client.responses.create.assert_called_once()
    call_args = mock_client.responses.create.call_args[1]
    assert "text" not in call_args or "verbosity" not in call_args.get("text", {})

def test_sync_chat_completion(adapter):
    params = {"response_format": {"type": "json_object"}}
    with patch.object(adapter.client.chat.completions, 'create') as mock_create:
        mock_response = MagicMock()
        mock_response.choices = [MagicMock(message=MagicMock(content='{"ok": true}'))]
        mock_create.return_value = mock_response
        
        result = adapter.sync_chat_completion("gpt-5", [{"role": "user", "content": "test"}], **params)
        
        mock_create.assert_called_once()
        call_args = mock_create.call_args[1]
        assert "response_format" not in call_args
        assert result.choices[0].message.content == '{"ok": true}'

# Edge cases
def test_strip_empty_params(adapter):
    params = {"text": {}}
    cleaned = adapter._strip_unsupported_params(params, "gpt-5")
    assert "text" not in cleaned  # Empty dict removed

@pytest.mark.asyncio
async def test_batch_translate_empty(adapter):
    result = await adapter.batch_translate("gpt-4o", [], offline_mode=False)
    assert result == []

# Run with coverage
if __name__ == "__main__":
    pytest.main(["-v", __file__])