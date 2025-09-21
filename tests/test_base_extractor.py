import pytest
from scripts.base_extractor import BaseExtractor

def test_base_extractor_abstract():
    with pytest.raises(TypeError):
        BaseExtractor()