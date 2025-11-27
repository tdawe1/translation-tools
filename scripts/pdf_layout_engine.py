#!/usr/bin/env python3
"""
PDF Layout Engine for Japanese-to-English Translation

Handles text expansion during translation by optimizing font sizes and layout
to prevent overflow while preserving readability and formatting.
"""

import re
import math
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Union
from enum import Enum
import logging

logger = logging.getLogger(__name__)


class ContentType(Enum):
    """Types of content with different expansion handling priorities."""
    TITLE = "title"
    HEADING = "heading"
    BODY = "body"
    CAPTION = "caption"
    TABLE = "table"
    FOOTER = "footer"
    HEADER = "header"


class LayoutConstraint(Enum):
    """Types of layout constraints for text blocks."""
    FIXED = "fixed"          # Cannot expand (e.g., table cells)
    FLEXIBLE = "flexible"    # Can expand within limits
    FREE = "free"           # No constraints


@dataclass
class TextBlock:
    """Represents a block of text with layout information."""
    id: str
    text: str
    jp_text: str  # Original Japanese text
    en_text: str  # Translated English text
    content_type: ContentType
    constraint: LayoutConstraint
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0
    font_size: float = 12.0
    font_name: str = "Arial"
    line_spacing: float = 1.2
    char_spacing: float = 1.0
    priority: int = 1  # Higher = more important to preserve
    min_font_size: float = 8.0
    max_font_size: float = 72.0
    expansion_factor: float = 1.0
    optimized_font_size: float = 12.0
    overflow_detected: bool = False
    adjustment_applied: bool = False


@dataclass
class LayoutAdjustments:
    """Represents layout adjustments to be applied."""
    font_scalings: Dict[str, float] = field(default_factory=dict)
    spacing_adjustments: Dict[str, float] = field(default_factory=dict)
    overflow_blocks: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    success: bool = True


class PDFLayoutEngine:
    """
    Layout engine for handling text expansion in PDF translation.

    Implements intelligent font scaling and layout optimization to
    accommodate English text expansion while preserving readability.
    """

    def __init__(self,
                 min_font_scale: float = 0.7,
                 max_expansion_ratio: float = 1.5,
                 line_spacing_reduction: float = 0.9,
                 char_spacing_compression: float = 0.95):
        """
        Initialize layout engine with configuration parameters.

        Args:
            min_font_scale: Minimum font size as ratio of original (0.7 = 70%)
            max_expansion_ratio: Maximum allowed expansion before intervention
            line_spacing_reduction: Factor to reduce line spacing
            char_spacing_compression: Factor to compress character spacing
        """
        self.min_font_scale = min_font_scale
        self.max_expansion_ratio = max_expansion_ratio
        self.line_spacing_reduction = line_spacing_reduction
        self.char_spacing_compression = char_spacing_compression

        # Content type priorities (higher = more important)
        self.content_priorities = {
            ContentType.TITLE: 5,
            ContentType.HEADING: 4,
            ContentType.CAPTION: 3,
            ContentType.BODY: 2,
            ContentType.TABLE: 2,
            ContentType.HEADER: 1,
            ContentType.FOOTER: 1,
        }

        # Expansion thresholds by content type
        self.expansion_thresholds = {
            ContentType.TITLE: 1.3,
            ContentType.HEADING: 1.4,
            ContentType.BODY: 1.5,
            ContentType.CAPTION: 1.4,
            ContentType.TABLE: 1.2,
            ContentType.HEADER: 1.3,
            ContentType.FOOTER: 1.3,
        }

        logger.info(f"PDFLayoutEngine initialized with min_font_scale={min_font_scale}, "
                   f"max_expansion_ratio={max_expansion_ratio}")

    def calculate_expansion_factor(self, jp_text: str, en_text: str) -> float:
        """
        Calculate expansion factor between Japanese and English text.

        Args:
            jp_text: Original Japanese text
            en_text: Translated English text

        Returns:
            Expansion factor (en_text_length / jp_text_length)
        """
        if not jp_text or not en_text:
            return 1.0

        # Normalize by removing extra whitespace
        jp_clean = re.sub(r'\s+', ' ', jp_text.strip())
        en_clean = re.sub(r'\s+', ' ', en_text.strip())

        # Calculate character counts
        jp_chars = len(jp_clean)
        en_chars = len(en_clean)

        if jp_chars == 0:
            return 1.0

        expansion_factor = en_chars / jp_chars

        # Apply smoothing for very short texts
        if jp_chars < 10:
            # Use weighted average with typical expansion
            typical_expansion = 1.25
            weight = jp_chars / 10.0
            expansion_factor = (weight * expansion_factor +
                             (1 - weight) * typical_expansion)

        logger.debug(f"Expansion factor: {jp_chars} JP chars -> {en_chars} EN chars = {expansion_factor:.3f}")
        return expansion_factor

    def estimate_text_dimensions(self, text: str, font_size: float,
                               width: float, font_name: str = "Arial") -> Tuple[float, float]:
        """
        Estimate text dimensions based on font size and content.

        Args:
            text: Text content
            font_size: Font size in points
            width: Available width in points
            font_name: Font family name

        Returns:
            Tuple of (estimated_height, estimated_width)
        """
        # Simple estimation based on character count and font size
        # In a real implementation, this would use font metrics

        # Average character width (approximation)
        char_width = font_size * 0.6  # 60% of font size for most fonts

        # Average line height
        line_height = font_size * 1.2

        # Estimate number of lines needed
        chars_per_line = max(1, int(width / char_width))
        lines_needed = max(1, math.ceil(len(text) / chars_per_line))

        estimated_height = lines_needed * line_height
        estimated_width = min(width, len(text) * char_width)

        return estimated_height, estimated_width

    def calculate_optimal_font_size(self, text_block: TextBlock) -> float:
        """
        Calculate optimal font size to prevent overflow.

        Args:
            text_block: Text block with layout information

        Returns:
            Optimal font size in points
        """
        if text_block.constraint == LayoutConstraint.FREE:
            return text_block.font_size

        # Calculate required scaling based on expansion
        expansion_factor = text_block.expansion_factor

        # Get threshold for this content type
        threshold = self.expansion_thresholds.get(text_block.content_type, 1.4)

        if expansion_factor <= threshold:
            return text_block.font_size  # No scaling needed

        # Calculate required font scaling
        required_scale = threshold / expansion_factor

        # Apply minimum scale constraint
        required_scale = max(required_scale, self.min_font_scale)

        # Calculate new font size
        optimal_size = text_block.font_size * required_scale

        # Ensure within bounds
        optimal_size = max(text_block.min_font_size,
                          min(optimal_size, text_block.max_font_size))

        logger.debug(f"Block {text_block.id}: expansion={expansion_factor:.2f}, "
                    f"scale={required_scale:.2f}, optimal_size={optimal_size:.1f}pt")

        return optimal_size

    def optimize_font_sizes(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Optimize font sizes for all text blocks to prevent overflow.

        Args:
            text_blocks: List of text blocks to optimize

        Returns:
            List of optimized text blocks
        """
        logger.info(f"Optimizing font sizes for {len(text_blocks)} text blocks")

        # First pass: calculate expansion factors and optimal sizes
        for block in text_blocks:
            block.expansion_factor = self.calculate_expansion_factor(
                block.jp_text, block.en_text)
            block.priority = self.content_priorities.get(block.content_type, 1)
            block.optimized_font_size = self.calculate_optimal_font_size(block)

            # Mark if optimization was applied
            if block.optimized_font_size != block.font_size:
                block.adjustment_applied = True

        # Second pass: handle conflicts between adjacent blocks
        text_blocks = self._resolve_font_conflicts(text_blocks)

        # Third pass: apply spacing optimizations
        text_blocks = self._apply_spacing_optimizations(text_blocks)

        # Final pass: detect any remaining overflow
        text_blocks = self._detect_overflow(text_blocks)

        optimized_count = sum(1 for block in text_blocks
                             if block.optimized_font_size != block.font_size)

        logger.info(f"Font optimization complete: {optimized_count}/{len(text_blocks)} blocks adjusted")

        return text_blocks

    def _resolve_font_conflicts(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Resolve conflicts between adjacent text blocks.

        Args:
            text_blocks: List of text blocks

        Returns:
            List with resolved conflicts
        """
        # Group blocks by proximity and content type
        groups = self._group_adjacent_blocks(text_blocks)

        for group in groups:
            if len(group) <= 1:
                continue

            # Find minimum font size in group based on highest priority
            min_size_in_group = min(block.optimized_font_size for block in group)
            max_priority = max(block.priority for block in group)

            # Apply consistent scaling within group
            for block in group:
                if block.priority == max_priority:
                    # High priority blocks maintain their size
                    continue
                else:
                    # Lower priority blocks scale to match
                    scale_factor = min_size_in_group / block.optimized_font_size
                    if scale_factor < 1.0:
                        block.optimized_font_size = min_size_in_group
                        block.adjustment_applied = True

        return text_blocks

    def _group_adjacent_blocks(self, text_blocks: List[TextBlock]) -> List[List[TextBlock]]:
        """
        Group text blocks that are adjacent or related.

        Args:
            text_blocks: List of text blocks

        Returns:
            List of grouped text blocks
        """
        groups = []
        current_group = []

        # Sort by position (top to bottom, left to right)
        sorted_blocks = sorted(text_blocks, key=lambda b: (b.y, b.x))

        for block in sorted_blocks:
            if not current_group:
                current_group.append(block)
            else:
                # Check if block is adjacent to last block in current group
                last_block = current_group[-1]

                # Simple proximity check (within 50 points vertically or horizontally)
                is_adjacent = (abs(block.x - last_block.x) < 50 or
                              abs(block.y - last_block.y) < 50)

                if is_adjacent:
                    current_group.append(block)
                else:
                    groups.append(current_group)
                    current_group = [block]

        if current_group:
            groups.append(current_group)

        return groups

    def _apply_spacing_optimizations(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Apply line spacing and character spacing optimizations.

        Args:
            text_blocks: List of text blocks

        Returns:
            List with spacing optimizations applied
        """
        for block in text_blocks:
            if block.expansion_factor > self.expansion_thresholds.get(
                block.content_type, 1.4):

                # Reduce line spacing
                block.line_spacing *= self.line_spacing_reduction

                # Compress character spacing slightly
                block.char_spacing *= self.char_spacing_compression

                block.adjustment_applied = True

                logger.debug(f"Applied spacing optimization to block {block.id}: "
                           f"line_spacing={block.line_spacing:.2f}, "
                           f"char_spacing={block.char_spacing:.2f}")

        return text_blocks

    def _detect_overflow(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Detect remaining overflow after optimizations.

        Args:
            text_blocks: List of text blocks

        Returns:
            List with overflow detection applied
        """
        for block in text_blocks:
            if block.constraint != LayoutConstraint.FREE:
                # Estimate dimensions with optimized font size
                est_height, est_width = self.estimate_text_dimensions(
                    block.en_text, block.optimized_font_size, block.width, block.font_name)

                # Check if overflow would occur
                if est_height > block.height * 1.1:  # 10% tolerance
                    block.overflow_detected = True
                    logger.warning(f"Overflow detected in block {block.id}: "
                                 f"est_height={est_height:.1f} > height={block.height:.1f}")

        return text_blocks

    def handle_overflow(self, text_blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Handle blocks that still have overflow after optimization.

        Args:
            text_blocks: List of text blocks with potential overflow

        Returns:
            List with overflow handled
        """
        overflow_blocks = [block for block in text_blocks if block.overflow_detected]

        if not overflow_blocks:
            return text_blocks

        logger.info(f"Handling overflow in {len(overflow_blocks)} blocks")

        for block in overflow_blocks:
            # Strategy 1: Further reduce font size (as last resort)
            if block.optimized_font_size > block.min_font_size:
                emergency_scale = 0.9
                new_size = block.optimized_font_size * emergency_scale

                if new_size >= block.min_font_size:
                    block.optimized_font_size = new_size
                    block.adjustment_applied = True
                    logger.info(f"Emergency font scaling for block {block.id}: {new_size:.1f}pt")
                    continue

            # Strategy 2: Truncate text with ellipsis
            if block.content_type in [ContentType.BODY, ContentType.CAPTION]:
                truncated_text = self._truncate_text(block.en_text, block.width,
                                                   block.optimized_font_size)
                if truncated_text != block.en_text:
                    block.en_text = truncated_text
                    block.adjustment_applied = True
                    logger.info(f"Text truncated in block {block.id}")

        return text_blocks

    def _truncate_text(self, text: str, width: float, font_size: float) -> str:
        """
        Truncate text to fit within width constraints.

        Args:
            text: Text to truncate
            width: Available width
            font_size: Font size

        Returns:
            Truncated text with ellipsis
        """
        # Simple truncation based on character count
        # In a real implementation, this would use more sophisticated text measurement

        char_width = font_size * 0.6
        max_chars = int(width / char_width) - 3  # Leave room for ellipsis

        if len(text) <= max_chars:
            return text

        # Try to truncate at word boundary
        truncated = text[:max_chars]
        last_space = truncated.rfind(' ')

        if last_space > max_chars * 0.8:  # If we found a good breaking point
            truncated = truncated[:last_space]

        return truncated + "..."

    def suggest_layout_adjustments(self, pdf_path: str) -> LayoutAdjustments:
        """
        Analyze PDF and suggest layout adjustments.

        Args:
            pdf_path: Path to PDF file

        Returns:
            LayoutAdjustments with suggested changes
        """
        # This would typically involve PDF parsing
        # For now, return empty adjustments

        logger.info(f"Analyzing layout adjustments for {pdf_path}")

        adjustments = LayoutAdjustments()

        # In a real implementation, this would:
        # 1. Extract text blocks from PDF
        # 2. Calculate expansion factors
        # 3. Determine optimal adjustments
        # 4. Return comprehensive suggestions

        return adjustments

    def generate_layout_report(self, text_blocks: List[TextBlock]) -> Dict[str, any]:
        """
        Generate a report of layout optimizations applied.

        Args:
            text_blocks: List of processed text blocks

        Returns:
            Dictionary with layout optimization statistics
        """
        total_blocks = len(text_blocks)
        optimized_blocks = sum(1 for block in text_blocks if block.adjustment_applied)
        overflow_blocks = sum(1 for block in text_blocks if block.overflow_detected)

        avg_expansion = sum(block.expansion_factor for block in text_blocks) / total_blocks if total_blocks > 0 else 0

        avg_scale = sum(block.optimized_font_size / block.font_size
                      for block in text_blocks) / total_blocks if total_blocks > 0 else 0

        report = {
            "total_blocks": total_blocks,
            "optimized_blocks": optimized_blocks,
            "overflow_blocks": overflow_blocks,
            "average_expansion_factor": avg_expansion,
            "average_font_scale": avg_scale,
            "optimization_rate": optimized_blocks / total_blocks if total_blocks > 0 else 0,
            "overflow_rate": overflow_blocks / total_blocks if total_blocks > 0 else 0,
        }

        logger.info(f"Layout report generated: {report}")

        return report


# Utility functions
def create_text_block(id: str, jp_text: str, en_text: str,
                     content_type: ContentType, **kwargs) -> TextBlock:
    """
    Create a TextBlock with sensible defaults.

    Args:
        id: Block identifier
        jp_text: Japanese text
        en_text: English text
        content_type: Type of content
        **kwargs: Additional properties

    Returns:
        TextBlock instance
    """
    defaults = {
        'constraint': LayoutConstraint.FLEXIBLE,
        'x': 0.0, 'y': 0.0,
        'width': 400.0, 'height': 100.0,
        'font_size': 12.0,
        'font_name': 'Arial',
        'line_spacing': 1.2,
        'char_spacing': 1.0,
        'priority': 1,
        'min_font_size': 8.0,
        'max_font_size': 72.0,
    }

    # Override defaults with provided kwargs
    defaults.update(kwargs)

    return TextBlock(
        id=id,
        text=en_text,
        jp_text=jp_text,
        en_text=en_text,
        content_type=content_type,
        **defaults
    )


def analyze_layout_health(text_blocks: List[TextBlock]) -> Dict[str, any]:
    """
    Analyze overall layout health after optimization.

    Args:
        text_blocks: List of processed text blocks

    Returns:
        Dictionary with health metrics
    """
    if not text_blocks:
        return {"status": "empty", "issues": []}

    issues = []

    # Check for remaining overflow
    overflow_blocks = [block for block in text_blocks if block.overflow_detected]
    if overflow_blocks:
        issues.append(f"{len(overflow_blocks)} blocks still have overflow")

    # Check for excessive font reduction
    excessive_scaling = [block for block in text_blocks
                        if block.optimized_font_size / block.font_size < 0.8]
    if excessive_scaling:
        issues.append(f"{len(excessive_scaling)} blocks scaled below 80%")

    # Check for inconsistent scaling
    font_sizes = [block.optimized_font_size for block in text_blocks]
    if font_sizes:
        size_variation = (max(font_sizes) - min(font_sizes)) / max(font_sizes)
        if size_variation > 0.5:
            issues.append(f"High font size variation: {size_variation:.1%}")

    # Overall health score
    health_score = 1.0
    health_score -= len(overflow_blocks) * 0.2
    health_score -= len(excessive_scaling) * 0.1
    health_score -= len(issues) * 0.05

    health_score = max(0.0, min(1.0, health_score))

    return {
        "status": "healthy" if health_score > 0.8 else "needs_attention",
        "health_score": health_score,
        "issues": issues,
        "total_blocks": len(text_blocks),
        "overflow_blocks": len(overflow_blocks),
        "excessively_scaled": len(excessive_scaling),
    }


# Example usage
if __name__ == "__main__":
    # Example usage of the PDF layout engine
    engine = PDFLayoutEngine()

    # Create sample text blocks
    blocks = [
        create_text_block("title", "タイトル", "This is a much longer title that needs optimization",
                         ContentType.TITLE, font_size=24.0, constraint=LayoutConstraint.FIXED, width=300.0),
        create_text_block("body1", "これは本文です", "This is body text that will expand significantly when translated from Japanese to English",
                         ContentType.BODY, font_size=12.0, width=400.0),
        create_text_block("body2", "短い文", "Short text",
                         ContentType.BODY, font_size=12.0, width=400.0),
    ]

    # Optimize layout
    optimized_blocks = engine.optimize_font_sizes(blocks)

    # Generate report
    report = engine.generate_layout_report(optimized_blocks)
    health = analyze_layout_health(optimized_blocks)

    print("Layout Optimization Report:")
    print(f"  Total blocks: {report['total_blocks']}")
    print(f"  Optimized: {report['optimized_blocks']}")
    print(f"  Average expansion: {report['average_expansion_factor']:.2f}x")
    print(f"  Health score: {health['health_score']:.2f}")

    for block in optimized_blocks:
        print(f"  {block.id}: {block.font_size:.1f}pt -> {block.optimized_font_size:.1f}pt "
              f"({block.optimized_font_size/block.font_size:.1%})")