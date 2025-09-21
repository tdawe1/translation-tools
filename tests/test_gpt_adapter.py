import pytest
import os
from unittest.mock import Mock, patch
from utils.gpt_adapter import GPT5Adapter

@pytest.fixture
def adapter():
    api_key = 'test_key'
    return GPT5Adapter(api_key=api_key)

def test_strip_unsupported_params(adapter):
    params = {
        'model': 'gpt-5',
        'messages': [{'role': 'user', 'content': 'test'}],
        'response_format': {'type': 'json'},
        'verbosity': 'low',
        'temperature': 0.6
    }
    cleaned = adapter._strip_unsupported_params(params)
    assert 'response_format' not in cleaned
    assert 'verbosity' not in cleaned
    assert 'temperature' in cleaned

def test_chunk_large_payload(adapter):
    large_payload = 'a' * 6 * 1024 * 1024  # >5MB
    chunks = adapter._chunk_large_payload(large_payload)
    assert len(chunks) > 1
    assert len(chunks[0]) <= 5 * 1024 * 1024

@patch('openai.OpenAI.chat.completions.create')
def test_chat_completions_create_fallback(mock_create, adapter):
    mock_create.side_effect = Exception('API error')
    params = {
        'model': 'gpt-5',
        'messages': [{'role': 'user', 'content': 'test'}]
    }
    with pytest.raises(Exception):
        adapter.chat_completions_create(**params)
    # Check fallback model called
    mock_create.assert_any_call(model='gpt-4o-2024-08-06', messages=[{'role': 'user', 'content': 'test'}])

def test_handle_large_response(adapter):
    large_response = 'a' * 11 * 1024 * 1024  # >10MB
    chunked = adapter._handle_large_response(large_response)
    assert 'chunks' in json.loads(chunked)

# Integration test simulation
@patch.dict(os.environ, {'OPENAI_API_KEY': 'test'})
def test_adapter_integration():
    adapter = GPT5Adapter(api_key='test')
    # Mock the client calls
    with patch.object(adapter.client, 'chat.completions.create', return_value=Mock(choices=[Mock(message=Mock(content='response'))])):
        result = adapter.chat_completions_create(model='gpt-5', messages=[{'role': 'user', 'content': 'test'}])
        assert result == 'response'