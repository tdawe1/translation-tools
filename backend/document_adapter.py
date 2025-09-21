"""
Document adapter base classes and data structures.
"""

from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, List, Optional


class SegmentType(Enum):
    """Types of document segments."""
    PARAGRAPH = "paragraph"
    HEADING = "heading"
    TABLE = "table"
    FOOTNOTE = "footnote"
    HEADER = "header"
    FOOTER = "footer"


@dataclass
class Segment:
    """A text segment from a document."""
    id: str
    text: str
    segment_type: SegmentType
    metadata: Dict[str, Any]
    has_japanese: bool = False
    word_count: int = 0
    position: Optional[str] = None


@dataclass
class DocumentMetadata:
    """Metadata about a document."""
    filename: str
    file_size: int
    page_count: Optional[int] = None
    word_count: Optional[int] = None
    character_count: Optional[int] = None
    language: Optional[str] = None
    created_at: Optional[str] = None
    modified_at: Optional[str] = None


@dataclass
class TranslationResult:
    """Result of a translation operation."""
    output_path: str
    segments_translated: int
    total_segments: int
    words_translated: int
    total_words: int
    cache_hits: int
    processing_time: float
    warnings: List[str]
    artifacts: Dict[str, Any]


class BaseDocumentAdapter:
    """Base class for document adapters."""

    def extract_segments(self, file_path: str) -> List[Segment]:
        """Extract text segments from document."""
        raise NotImplementedError

    def apply_translations(self, input_path: str, segments: List[Segment], output_path: str) -> None:
        """Apply translations to document."""
        raise NotImplementedError

    def get_metadata(self, file_path: str) -> DocumentMetadata:
        """Get document metadata."""
        raise NotImplementedError