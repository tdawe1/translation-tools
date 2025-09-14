#!/usr/bin/env python3
"""
PDF audit tool for quality assessment of translated PDF documents.

This tool provides comprehensive quality assessment for PDF translations, including:
- Residual Japanese character detection
- Layout integrity verification  
- Translation quality assessment
- Structured audit reports

Usage:
  python audit_pdf.py translated.pdf [original.pdf] [--report REPORT.csv]
  python audit_pdf.py --help
"""
import re
import csv
import json
import sys
import argparse
from typing import Dict, List, Tuple, Any, Optional
from dataclasses import dataclass, asdict
from pathlib import Path
from collections import Counter, defaultdict

# Required PDF processing libraries
try:
    import pypdf
except ImportError:
    print("ERROR: pypdf is required. Install via: pip install pypdf>=3.0.0", file=sys.stderr)
    sys.exit(1)

try:
    from pdfminer.high_level import extract_text_to_fp
    from pdfminer.layout import LAParams
except ImportError:
    print("ERROR: pdfminer.six is required. Install via: pip install pdfminer.six>=20221105", file=sys.stderr)
    sys.exit(1)

import io


@dataclass
class LayoutCheckResult:
    """Result of layout integrity check."""
    score: float  # 0.0 to 1.0
    issues: List[str]
    page_count_match: bool
    similar_structure: bool


@dataclass 
class QualityAssessment:
    """Translation quality assessment metrics."""
    residual_japanese_count: int
    residual_japanese_percentage: float
    text_completeness_score: float
    formatting_consistency_score: float
    overall_quality_score: float
    recommendations: List[str]


@dataclass
class AuditReport:
    """Comprehensive audit report."""
    file_path: str
    original_file_path: Optional[str]
    timestamp: str
    total_pages: int
    extracted_text_length: int
    layout_check: Optional[LayoutCheckResult]
    quality_assessment: QualityAssessment
    page_details: List[Dict[str, Any]]


class PDFAuditor:
    """PDF quality auditor for translated documents."""
    
    def __init__(self):
        # Japanese character ranges (Hiragana, Katakana, Kanji, CJK punctuation)
        self.jp_pattern = re.compile(
            r'[\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]'
        )
        # Core Japanese letters (excluding punctuation)
        self.jp_core_pattern = re.compile(
            r'[\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff]'
        )
        
    def count_residual_jp(self, pdf_path: str) -> int:
        """
        Count residual Japanese characters in translated PDF.
        
        Args:
            pdf_path: Path to the translated PDF file
            
        Returns:
            Number of Japanese characters found
        """
        text = self._extract_text(pdf_path)
        jp_chars = self.jp_core_pattern.findall(text)
        return len(jp_chars)
    
    def check_layout_integrity(self, original: str, translated: str) -> LayoutCheckResult:
        """
        Verify layout integrity between original and translated PDFs.
        
        Args:
            original: Path to original PDF
            translated: Path to translated PDF
            
        Returns:
            LayoutCheckResult with score and issues
        """
        try:
            # Extract basic layout information
            orig_pages = self._get_page_count(original)
            trans_pages = self._get_page_count(translated)
            
            issues = []
            score = 1.0
            
            # Check page count
            if orig_pages != trans_pages:
                issues.append(f"Page count mismatch: original={orig_pages}, translated={trans_pages}")
                score -= 0.3
            
            # Extract text for structural comparison
            orig_text = self._extract_text(original)
            trans_text = self._extract_text(translated)
            
            # Check text completeness (translated should be longer due to English expansion)
            if len(trans_text) < len(orig_text) * 0.7:
                issues.append("Translated text appears incomplete")
                score -= 0.2
            
            # Check for excessive text expansion
            if len(trans_text) > len(orig_text) * 3.0:
                issues.append("Excessive text expansion may indicate layout issues")
                score -= 0.1
            
            # Analyze text structure (paragraphs, line breaks)
            orig_structure = self._analyze_text_structure(orig_text)
            trans_structure = self._analyze_text_structure(trans_text)
            
            structure_similarity = self._compare_structures(orig_structure, trans_structure)
            if structure_similarity < 0.7:
                issues.append("Text structure differs significantly from original")
                score -= 0.2
            
            return LayoutCheckResult(
                score=max(0.0, score),
                issues=issues,
                page_count_match=orig_pages == trans_pages,
                similar_structure=structure_similarity >= 0.7
            )
            
        except Exception as e:
            return LayoutCheckResult(
                score=0.0,
                issues=[f"Layout check failed: {str(e)}"],
                page_count_match=False,
                similar_structure=False
            )
    
    def assess_translation_quality(self, pdf_path: str) -> QualityAssessment:
        """
        Assess translation quality of a PDF document.
        
        Args:
            pdf_path: Path to the translated PDF
            
        Returns:
            QualityAssessment with metrics and recommendations
        """
        text = self._extract_text(pdf_path)
        jp_chars = self.jp_core_pattern.findall(text)
        total_chars = len(re.sub(r'\s', '', text))
        
        # Calculate residual Japanese percentage
        jp_count = len(jp_chars)
        jp_percentage = (jp_count / total_chars * 100) if total_chars > 0 else 0
        
        # Text completeness score (basic heuristics)
        completeness_score = 1.0
        if total_chars < 100:  # Very short document
            completeness_score = 0.5
        elif '\n' not in text and ' ' not in text:  # No word separation
            completeness_score = 0.3
        
        # Formatting consistency score
        formatting_score = self._assess_formatting_consistency(text)
        
        # Overall quality score
        weights = {
            'japanese_penalty': min(jp_percentage / 5.0, 0.5),  # Lose up to 0.5 for Japanese
            'completeness': completeness_score * 0.3,
            'formatting': formatting_score * 0.2
        }
        overall_score = max(0.0, 1.0 - weights['japanese_penalty'] + weights['completeness'] + weights['formatting'])
        
        # Generate recommendations
        recommendations = []
        if jp_percentage > 1.0:
            recommendations.append(f"Remove residual Japanese characters ({jp_count} found, {jp_percentage:.1f}%)")
        if completeness_score < 0.7:
            recommendations.append("Check for incomplete translation or missing content")
        if formatting_score < 0.7:
            recommendations.append("Review formatting consistency and text layout")
        if overall_score < 0.8:
            recommendations.append("Overall quality below threshold - manual review recommended")
        
        return QualityAssessment(
            residual_japanese_count=jp_count,
            residual_japanese_percentage=jp_percentage,
            text_completeness_score=completeness_score,
            formatting_consistency_score=formatting_score,
            overall_quality_score=overall_score,
            recommendations=recommendations
        )
    
    def generate_audit_report(self, pdf_path: str, original_pdf_path: Optional[str] = None) -> AuditReport:
        """
        Generate comprehensive audit report for a translated PDF.
        
        Args:
            pdf_path: Path to translated PDF
            original_pdf_path: Optional path to original PDF for comparison
            
        Returns:
            Complete AuditReport
        """
        from datetime import datetime
        
        # Basic file information
        pdf_path = Path(pdf_path)
        timestamp = datetime.now().isoformat()
        
        try:
            total_pages = self._get_page_count(str(pdf_path))
            text = self._extract_text(str(pdf_path))
            extracted_text_length = len(text)
        except Exception as e:
            raise RuntimeError(f"Failed to process PDF: {e}")
        
        # Layout integrity check (if original provided)
        layout_check = None
        if original_pdf_path:
            layout_check = self.check_layout_integrity(original_pdf_path, str(pdf_path))
        
        # Quality assessment
        quality_assessment = self.assess_translation_quality(str(pdf_path))
        
        # Page-level details
        page_details = self._analyze_pages(str(pdf_path))
        
        return AuditReport(
            file_path=str(pdf_path),
            original_file_path=original_pdf_path,
            timestamp=timestamp,
            total_pages=total_pages,
            extracted_text_length=extracted_text_length,
            layout_check=layout_check,
            quality_assessment=quality_assessment,
            page_details=page_details
        )
    
    def compare_with_original(self, original: str, translated: str) -> Dict[str, Any]:
        """
        Compare translated PDF with original Japanese PDF.
        
        Args:
            original: Path to original PDF
            translated: Path to translated PDF
            
        Returns:
            Comparison results with detailed metrics
        """
        orig_text = self._extract_text(original)
        trans_text = self._extract_text(translated)
        
        # Basic metrics
        orig_jp_count = len(self.jp_core_pattern.findall(orig_text))
        trans_jp_count = len(self.jp_core_pattern.findall(trans_text))
        
        # Text expansion analysis
        orig_clean = re.sub(r'\s+', '', orig_text)
        trans_clean = re.sub(r'\s+', '', trans_text)
        
        expansion_ratio = len(trans_clean) / len(orig_clean) if len(orig_clean) > 0 else 0
        
        # Character type analysis
        orig_char_types = self._analyze_character_types(orig_text)
        trans_char_types = self._analyze_character_types(trans_text)
        
        return {
            'original_stats': {
                'total_chars': len(orig_clean),
                'japanese_chars': orig_jp_count,
                'japanese_percentage': (orig_jp_count / len(orig_clean) * 100) if len(orig_clean) > 0 else 0,
                'char_types': orig_char_types
            },
            'translated_stats': {
                'total_chars': len(trans_clean),
                'japanese_chars': trans_jp_count,
                'japanese_percentage': (trans_jp_count / len(trans_clean) * 100) if len(trans_clean) > 0 else 0,
                'char_types': trans_char_types
            },
            'expansion_ratio': expansion_ratio,
            'translation_completeness': min(expansion_ratio / 1.5, 1.0),  # Assuming 1.5x expansion is good
            'residual_japanese_removed': ((orig_jp_count - trans_jp_count) / orig_jp_count * 100) if orig_jp_count > 0 else 100
        }
    
    def _extract_text(self, pdf_path: str) -> str:
        """Extract text from PDF using pdfminer for better accuracy."""
        try:
            # Try pdfminer first for better text extraction
            output_string = io.StringIO()
            with open(pdf_path, 'rb') as pdf_file:
                laparams = LAParams()
                extract_text_to_fp(pdf_file, output_string, laparams=laparams)
            return output_string.getvalue()
        except Exception:
            # Fallback to pypdf
            return self._extract_text_fallback(pdf_path)
    
    def _extract_text_fallback(self, pdf_path: str) -> str:
        """Extract text using pypdf as fallback."""
        text = ""
        with open(pdf_path, 'rb') as file:
            reader = pypdf.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() + "\n"
        return text
    
    def _get_page_count(self, pdf_path: str) -> int:
        """Get total page count of PDF."""
        with open(pdf_path, 'rb') as file:
            reader = pypdf.PdfReader(file)
            return len(reader.pages)
    
    def _analyze_text_structure(self, text: str) -> Dict[str, Any]:
        """Analyze text structure for comparison."""
        lines = text.split('\n')
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
        
        return {
            'line_count': len(lines),
            'paragraph_count': len(paragraphs),
            'avg_line_length': sum(len(line) for line in lines) / len(lines) if lines else 0,
            'avg_paragraph_length': sum(len(p) for p in paragraphs) / len(paragraphs) if paragraphs else 0
        }
    
    def _compare_structures(self, orig: Dict[str, Any], trans: Dict[str, Any]) -> float:
        """Compare text structures and return similarity score."""
        # Simple similarity based on structural metrics
        line_ratio = min(orig['line_count'], trans['line_count']) / max(orig['line_count'], trans['line_count']) if max(orig['line_count'], trans['line_count']) > 0 else 0
        para_ratio = min(orig['paragraph_count'], trans['paragraph_count']) / max(orig['paragraph_count'], trans['paragraph_count']) if max(orig['paragraph_count'], trans['paragraph_count']) > 0 else 0
        
        return (line_ratio + para_ratio) / 2
    
    def _assess_formatting_consistency(self, text: str) -> float:
        """Assess formatting consistency in extracted text."""
        score = 1.0
        
        # Check for consistent line endings
        if '\r\n' in text and '\n' in text:
            score -= 0.1
        
        # Check for unusual spacing patterns
        if re.search(r'[ ]{3,}', text):  # Multiple spaces
            score -= 0.1
        
        # Check for broken words or unusual hyphenation
        if re.search(r'\w+-\s*\w+', text):  # Hyphenated words across lines
            score -= 0.1
        
        return max(0.0, score)
    
    def _analyze_character_types(self, text: str) -> Dict[str, int]:
        """Analyze types of characters in text."""
        char_types = {
            'japanese': len(self.jp_core_pattern.findall(text)),
            'english': len(re.findall(r'[a-zA-Z]', text)),
            'digits': len(re.findall(r'\d', text)),
            'punctuation': len(re.findall(r'[^\w\s\u3040-\u9fff]', text)),
            'whitespace': len(re.findall(r'\s', text))
        }
        return char_types
    
    def _analyze_pages(self, pdf_path: str) -> List[Dict[str, Any]]:
        """Analyze each page individually."""
        page_details = []
        
        try:
            with open(pdf_path, 'rb') as file:
                reader = pypdf.PdfReader(file)
                
                for i, page in enumerate(reader.pages):
                    page_text = page.extract_text()
                    jp_count = len(self.jp_core_pattern.findall(page_text))
                    word_count = len(page_text.split())
                    
                    page_details.append({
                        'page_number': i + 1,
                        'word_count': word_count,
                        'japanese_chars': jp_count,
                        'text_length': len(page_text)
                    })
        except Exception:
            # If page-by-page analysis fails, provide basic info
            pass
        
        return page_details


def save_report_csv(report: AuditReport, output_path: str) -> None:
    """Save audit report as CSV file."""
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'metric', 'value', 'details'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        # File information
        writer.writerow({'metric': 'file_path', 'value': report.file_path, 'details': ''})
        writer.writerow({'metric': 'timestamp', 'value': report.timestamp, 'details': ''})
        writer.writerow({'metric': 'total_pages', 'value': report.total_pages, 'details': ''})
        writer.writerow({'metric': 'extracted_text_length', 'value': report.extracted_text_length, 'details': ''})
        
        # Quality assessment
        qa = report.quality_assessment
        writer.writerow({'metric': 'residual_japanese_count', 'value': qa.residual_japanese_count, 'details': f'{qa.residual_japanese_percentage:.2f}%'})
        writer.writerow({'metric': 'text_completeness_score', 'value': f'{qa.text_completeness_score:.2f}', 'details': ''})
        writer.writerow({'metric': 'formatting_consistency_score', 'value': f'{qa.formatting_consistency_score:.2f}', 'details': ''})
        writer.writerow({'metric': 'overall_quality_score', 'value': f'{qa.overall_quality_score:.2f}', 'details': ''})
        
        # Layout check (if available)
        if report.layout_check:
            lc = report.layout_check
            writer.writerow({'metric': 'layout_integrity_score', 'value': f'{lc.score:.2f}', 'details': ''})
            writer.writerow({'metric': 'page_count_match', 'value': str(lc.page_count_match), 'details': ''})
            for issue in lc.issues:
                writer.writerow({'metric': 'layout_issue', 'value': 'issue', 'details': issue})
        
        # Recommendations
        for i, rec in enumerate(qa.recommendations):
            writer.writerow({'metric': f'recommendation_{i+1}', 'value': 'recommendation', 'details': rec})


def save_report_json(report: AuditReport, output_path: str) -> None:
    """Save audit report as JSON file."""
    report_dict = asdict(report)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(report_dict, f, indent=2, ensure_ascii=False)


def main():
    """CLI interface for PDF audit tool."""
    parser = argparse.ArgumentParser(
        description="PDF audit tool for quality assessment of translated documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python audit_pdf.py translated.pdf
  python audit_pdf.py translated.pdf original.pdf --report audit.csv
  python audit_pdf.py translated.pdf --json --output audit.json
        """
    )
    
    parser.add_argument("translated_pdf", help="Path to translated PDF file")
    parser.add_argument("original_pdf", nargs="?", help="Path to original Japanese PDF (optional)")
    parser.add_argument("--report", "-r", help="Output CSV report path", default="PDF_AUDIT_REPORT.csv")
    parser.add_argument("--json", "-j", action="store_true", help="Output JSON format instead of CSV")
    parser.add_argument("--output", "-o", help="Output file path (for JSON)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    parser.add_argument("--threshold", "-t", type=float, default=0.8, help="Quality threshold (0.0-1.0)")
    
    args = parser.parse_args()
    
    # Validate files exist
    if not Path(args.translated_pdf).exists():
        print(f"Error: Translated PDF not found: {args.translated_pdf}")
        sys.exit(1)
    
    if args.original_pdf and not Path(args.original_pdf).exists():
        print(f"Error: Original PDF not found: {args.original_pdf}")
        sys.exit(1)
    
    # Initialize auditor
    auditor = PDFAuditor()
    
    try:
        # Generate audit report
        report = auditor.generate_audit_report(args.translated_pdf, args.original_pdf)
        
        # Print summary
        print(f"\n=== PDF Audit Report ===")
        print(f"File: {report.file_path}")
        if report.original_file_path:
            print(f"Original: {report.original_file_path}")
        print(f"Pages: {report.total_pages}")
        print(f"Text length: {report.extracted_text_length}")
        
        qa = report.quality_assessment
        print(f"\n=== Quality Assessment ===")
        print(f"Residual Japanese: {qa.residual_japanese_count} chars ({qa.residual_japanese_percentage:.2f}%)")
        print(f"Completeness Score: {qa.text_completeness_score:.2f}")
        print(f"Formatting Score: {qa.formatting_consistency_score:.2f}")
        print(f"Overall Quality: {qa.overall_quality_score:.2f}")
        
        if qa.recommendations:
            print(f"\n=== Recommendations ===")
            for i, rec in enumerate(qa.recommendations, 1):
                print(f"{i}. {rec}")
        
        if report.layout_check:
            lc = report.layout_check
            print(f"\n=== Layout Integrity ===")
            print(f"Score: {lc.score:.2f}")
            print(f"Page Count Match: {lc.page_count_match}")
            if lc.issues:
                print("Issues:")
                for issue in lc.issues:
                    print(f"  - {issue}")
        
        # Save report
        if args.json:
            output_path = args.output or "PDF_AUDIT_REPORT.json"
            save_report_json(report, output_path)
            print(f"\nJSON report saved to: {output_path}")
        else:
            save_report_csv(report, args.report)
            print(f"\nCSV report saved to: {args.report}")
        
        # Exit with error code if quality below threshold
        if qa.overall_quality_score < args.threshold:
            print(f"\nQuality score {qa.overall_quality_score:.2f} below threshold {args.threshold}")
            sys.exit(1)
        else:
            print(f"\nQuality check passed (score: {qa.overall_quality_score:.2f})")
            sys.exit(0)
            
    except Exception as e:
        print(f"Error during audit: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()