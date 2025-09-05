#!/usr/bin/env python3
"""
Deck-level consistency audit for style drift detection.
Fast, deterministic checks that can fail CI on style violations.
"""
import re
import csv
import json
from typing import List, Dict, Tuple, Any
from collections import Counter, defaultdict
from style_normalize import BANNED_PHRASES, title_case

def audit_capitalization(rows: List[Tuple]) -> List[Dict[str, Any]]:
    """Check for Title Case violations in title content."""
    issues = []
    
    for slide_xml, idx, jp, en, kind in rows:
        if kind == "title" and en.strip():
            # Remove formatting tags for analysis
            clean_text = re.sub(r'\[/?[^\]]+\]', '', en)
            expected = title_case(clean_text)
            
            if clean_text != expected:
                issues.append({
                    "type": "title_case",
                    "slide": slide_xml,
                    "index": idx,
                    "text": en[:50] + ("..." if len(en) > 50 else ""),
                    "expected": expected[:50] + ("..." if len(expected) > 50 else "")
                })
    
    return issues

def audit_bullet_punctuation(rows: List[Tuple]) -> List[Dict[str, Any]]:
    """Check for inappropriate terminal punctuation in bullets."""
    issues = []
    
    for slide_xml, idx, jp, en, kind in rows:
        if kind == "bullet" and en.strip():
            clean_text = re.sub(r'\[/?[^\]]+\]', '', en).strip()
            
            # Check for terminal punctuation in fragments
            if re.search(r'[.;:]\s*$', clean_text):
                # Allow if genuinely multiple sentences
                sentence_count = len(re.findall(r'[.!?]+', clean_text))
                if sentence_count <= 1:
                    issues.append({
                        "type": "bullet_punct",
                        "slide": slide_xml,
                        "index": idx,
                        "text": clean_text[-20:],  # Show end of text
                        "issue": "bullet ends with terminal punctuation"
                    })
    
    return issues

def audit_banned_words(rows: List[Tuple]) -> List[Dict[str, Any]]:
    """Check for banned phrases that indicate tone drift."""
    issues = []
    
    banned_pattern = '|'.join(r'\b' + re.escape(phrase) + r'\b' for phrase in BANNED_PHRASES.keys())
    regex = re.compile(banned_pattern, re.IGNORECASE)
    
    for slide_xml, idx, jp, en, kind in rows:
        clean_text = re.sub(r'\[/?[^\]]+\]', '', en)
        matches = regex.findall(clean_text)
        
        for match in matches:
            suggested = BANNED_PHRASES.get(match.lower(), "review")
            issues.append({
                "type": "banned_word",
                "slide": slide_xml,
                "index": idx,
                "word": match,
                "suggested": suggested,
                "context": clean_text[:60] + ("..." if len(clean_text) > 60 else "")
            })
    
    return issues

def audit_terminology_consistency(rows: List[Tuple], glossary: Dict[str, str] = None) -> List[Dict[str, Any]]:
    """Check for terminology consistency across the deck."""
    issues = []
    
    if not glossary:
        return issues
    
    # Track usage of terms across slides
    term_usage = defaultdict(list)
    
    for slide_xml, idx, jp, en, kind in rows:
        clean_text = re.sub(r'\[/?[^\]]+\]', '', en).lower()
        
        # Check for each glossary term
        for jp_term, expected_en in glossary.items():
            expected_lower = expected_en.lower()
            
            # Simple term detection (could be enhanced with NLP)
            if expected_lower in clean_text:
                term_usage[jp_term].append({
                    "slide": slide_xml,
                    "index": idx,
                    "usage": expected_en,
                    "context": clean_text
                })
    
    # Check for inconsistent usage patterns
    for jp_term, usages in term_usage.items():
        if len(usages) > 1:
            usage_variants = set(usage["usage"] for usage in usages)
            if len(usage_variants) > 1:
                issues.append({
                    "type": "terminology_inconsistency",
                    "term": jp_term,
                    "variants": list(usage_variants),
                    "locations": [(u["slide"], u["index"]) for u in usages]
                })
    
    return issues

def audit_japanese_residual(rows: List[Tuple]) -> List[Dict[str, Any]]:
    """Check for remaining Japanese characters in translations."""
    issues = []
    jp_pattern = re.compile(r'[\u3040-\u309f\u30a0-\u30ff\u31f0-\u31ff\u3400-\u4dbf\u4e00-\u9fff]')
    
    for slide_xml, idx, jp, en, kind in rows:
        if jp_pattern.search(en):
            jp_chars = jp_pattern.findall(en)
            issues.append({
                "type": "japanese_residual",
                "slide": slide_xml,
                "index": idx,
                "jp_chars": ''.join(jp_chars),
                "text": en[:50] + ("..." if len(en) > 50 else "")
            })
    
    return issues

def audit_formatting_tags(rows: List[Tuple]) -> List[Dict[str, Any]]:
    """Check for malformed or unbalanced formatting tags."""
    issues = []
    
    for slide_xml, idx, jp, en, kind in rows:
        # Check for unbalanced tags
        open_tags = re.findall(r'\[([^/][^\]]*)\]', en)
        close_tags = re.findall(r'\[/([^\]]+)\]', en)
        
        open_counter = Counter(open_tags)
        close_counter = Counter(close_tags)
        
        # Check for unmatched tags
        for tag in set(open_tags + close_tags):
            open_count = open_counter.get(tag, 0)
            close_count = close_counter.get(tag, 0)
            
            if open_count != close_count:
                issues.append({
                    "type": "unbalanced_tags",
                    "slide": slide_xml,
                    "index": idx,
                    "tag": tag,
                    "open_count": open_count,
                    "close_count": close_count,
                    "text": en[:50] + ("..." if len(en) > 50 else "")
                })
    
    return issues

def audit_length_violations(rows: List[Tuple], max_title_words: int = 12, max_bullet_chars: int = 200) -> List[Dict[str, Any]]:
    """Check for content that exceeds recommended length limits."""
    issues = []
    
    for slide_xml, idx, jp, en, kind in rows:
        clean_text = re.sub(r'\[/?[^\]]+\]', '', en).strip()
        
        if kind == "title":
            word_count = len(clean_text.split())
            if word_count > max_title_words:
                issues.append({
                    "type": "title_too_long",
                    "slide": slide_xml,
                    "index": idx,
                    "word_count": word_count,
                    "max_words": max_title_words,
                    "text": clean_text[:50] + ("..." if len(clean_text) > 50 else "")
                })
        
        elif kind == "bullet":
            char_count = len(clean_text)
            if char_count > max_bullet_chars:
                issues.append({
                    "type": "bullet_too_long",
                    "slide": slide_xml,
                    "index": idx,
                    "char_count": char_count,
                    "max_chars": max_bullet_chars,
                    "text": clean_text[:50] + ("..." if len(clean_text) > 50 else "")
                })
    
    return issues

def run_full_audit(csv_path: str, glossary: Dict[str, str] = None) -> Dict[str, List[Dict[str, Any]]]:
    """
    Run comprehensive style audit on bilingual CSV output.
    
    Args:
        csv_path: Path to bilingual CSV file from translation
        glossary: Optional glossary for terminology checking
        
    Returns:
        Dictionary of audit results by category
    """
    # Load CSV data
    rows = []
    try:
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader)  # Skip header
            
            for row in reader:
                if len(row) >= 4:
                    slide_xml, idx, jp, en = row[:4]
                    # Detect content type (simplified)
                    kind = "title" if len(en.split()) <= 12 else "bullet"
                    rows.append((slide_xml, int(idx), jp, en, kind))
                    
    except FileNotFoundError:
        print(f"Bilingual CSV not found: {csv_path}")
        return {}
    
    # Run all audit checks
    audit_results = {
        "capitalization_issues": audit_capitalization(rows),
        "bullet_punctuation_issues": audit_bullet_punctuation(rows), 
        "banned_word_issues": audit_banned_words(rows),
        "terminology_issues": audit_terminology_consistency(rows, glossary),
        "japanese_residual": audit_japanese_residual(rows),
        "formatting_tag_issues": audit_formatting_tags(rows),
        "length_violations": audit_length_violations(rows)
    }
    
    return audit_results

def generate_audit_report(audit_results: Dict[str, List[Dict[str, Any]]], output_path: str = "STYLE_REPORT.csv"):
    """Generate human-readable audit report as CSV."""
    
    all_issues = []
    for category, issues in audit_results.items():
        for issue in issues:
            issue_row = {
                "category": category,
                "slide": issue.get("slide", ""),
                "index": issue.get("index", ""),
                "issue_type": issue.get("type", ""),
                "description": _format_issue_description(issue),
                "suggested_fix": _get_suggested_fix(issue)
            }
            all_issues.append(issue_row)
    
    # Write CSV report
    if all_issues:
        with open(output_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=["category", "slide", "index", "issue_type", "description", "suggested_fix"])
            writer.writeheader()
            writer.writerows(all_issues)
    
    return len(all_issues)

def _format_issue_description(issue: Dict[str, Any]) -> str:
    """Format issue description for report."""
    issue_type = issue.get("type", "")
    
    if issue_type == "title_case":
        return f"Title not in Title Case: '{issue.get('text', '')}'"
    elif issue_type == "bullet_punct":
        return f"Bullet ends with punctuation: '{issue.get('text', '')}'"
    elif issue_type == "banned_word":
        return f"Banned word '{issue.get('word', '')}' found"
    elif issue_type == "japanese_residual":
        return f"Japanese characters remain: {issue.get('jp_chars', '')}"
    elif issue_type == "unbalanced_tags":
        return f"Unbalanced tag [{issue.get('tag', '')}]: {issue.get('open_count', 0)} open, {issue.get('close_count', 0)} close"
    elif issue_type == "title_too_long":
        return f"Title too long: {issue.get('word_count', 0)} words (max {issue.get('max_words', 12)})"
    elif issue_type == "bullet_too_long":
        return f"Bullet too long: {issue.get('char_count', 0)} chars (max {issue.get('max_chars', 200)})"
    
    return str(issue)

def _get_suggested_fix(issue: Dict[str, Any]) -> str:
    """Get suggested fix for issue."""
    issue_type = issue.get("type", "")
    
    if issue_type == "title_case":
        return issue.get("expected", "Apply Title Case")
    elif issue_type == "bullet_punct":
        return "Remove terminal punctuation"
    elif issue_type == "banned_word":
        return f"Replace with '{issue.get('suggested', 'alternative')}'"
    elif issue_type == "japanese_residual":
        return "Re-translate remaining Japanese"
    elif issue_type == "unbalanced_tags":
        return f"Balance [{issue.get('tag', '')}] tags"
    elif issue_type in ["title_too_long", "bullet_too_long"]:
        return "Condense text or apply expansion policy"
    
    return "Review manually"

def should_fail_ci(audit_results: Dict[str, List[Dict[str, Any]]], 
                   max_issues: int = 10, 
                   critical_categories: List[str] = None) -> Tuple[bool, str]:
    """
    Determine if CI should fail based on audit results.
    
    Args:
        audit_results: Results from run_full_audit
        max_issues: Maximum total issues before failing
        critical_categories: Categories that cause immediate failure
        
    Returns:
        Tuple of (should_fail, reason)
    """
    if critical_categories is None:
        critical_categories = ["japanese_residual", "unbalanced_tags"]
    
    total_issues = sum(len(issues) for issues in audit_results.values())
    
    # Check for critical issues
    for category in critical_categories:
        if audit_results.get(category, []):
            return True, f"Critical issues found in {category}: {len(audit_results[category])} issues"
    
    # Check total issue count
    if total_issues > max_issues:
        return True, f"Too many style issues: {total_issues} (max allowed: {max_issues})"
    
    return False, f"Style audit passed: {total_issues} issues found"

# CLI interface
def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="Style consistency audit for PPTX translations")
    parser.add_argument("csv_path", help="Path to bilingual CSV file")
    parser.add_argument("--glossary", help="Path to glossary JSON file")
    parser.add_argument("--report", default="STYLE_REPORT.csv", help="Output report path")
    parser.add_argument("--max-issues", type=int, default=10, help="Maximum issues before CI failure")
    parser.add_argument("--fail-on-critical", action="store_true", help="Fail CI on critical issues")
    
    args = parser.parse_args()
    
    # Load glossary if provided
    glossary = {}
    if args.glossary:
        try:
            with open(args.glossary, 'r', encoding='utf-8') as f:
                glossary = json.load(f)
        except FileNotFoundError:
            print(f"Warning: Glossary file not found: {args.glossary}")
    
    # Run audit
    results = run_full_audit(args.csv_path, glossary)
    
    # Generate report
    issue_count = generate_audit_report(results, args.report)
    
    # Print summary
    print(f"Style Audit Results:")
    for category, issues in results.items():
        if issues:
            print(f"  {category}: {len(issues)} issues")
    
    print(f"\nTotal issues: {issue_count}")
    if issue_count > 0:
        print(f"Detailed report: {args.report}")
    
    # Check if CI should fail
    critical_categories = ["japanese_residual", "unbalanced_tags"] if args.fail_on_critical else []
    should_fail, reason = should_fail_ci(results, args.max_issues, critical_categories)
    
    print(f"\nCI Status: {reason}")
    
    if should_fail:
        exit(1)  # Fail CI
    else:
        exit(0)  # Pass CI

if __name__ == "__main__":
    main()