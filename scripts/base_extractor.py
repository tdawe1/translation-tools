from abc import ABC, abstractmethod
from typing import List, Tuple, Optional, Dict

class BaseExtractor(ABC):
    @abstractmethod
    def extract(self, file_path: str, page_range: Optional[str] = None) -> List[Tuple[str, int, str, Dict[str, any]]]:
        \"\"\"
        Extract translatable text blocks.
        Returns list of (section, index, text, layout_info)
        layout_info: dict with keys like 'font_size', 'position', 'type', etc.
        \"\"\"
        pass