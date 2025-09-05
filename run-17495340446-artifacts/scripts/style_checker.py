#!/usr/bin/env python3
"""
Model-based style checker with JSON diagnostics.
Returns structured feedback for deterministic authority fixes.
"""
import json
import re
from typing import List, Dict, Any
from style_normalize import SMALL_WORDS, BANNED_PHRASES, title_case, get_style_guide

def create_style_checker_prompt(translations: List[str], glossary: Dict[str, str] = None) -> str:
    """Create prompt for style checking with JSON diagnostics output."""
    
    glossary_section = ""
    if glossary:
        glossary_items = [f'"{jp}" → "{en}"' for jp, en in list(glossary.items())[:10]]
        glossary_section = f"""
**Key glossary terms:**
{'; '.join(glossary_items)}
"""
    
    style_guide = get_style_guide()
    
    return f"""You are a style reviewer for marketing slide translations. Review the provided English translations and return ONLY a JSON object with style diagnostics.

{style_guide}
{glossary_section}

**Analysis Instructions:**
1. Check each translation for style violations
2. Return structured JSON diagnostics for deterministic fixes
3. Focus on objective rule violations, not subjective preferences
4. Preserve all formatting tags [b][i][u][sup][sub][li-lN] and placeholders ⟦...⟧

**Required JSON format:**
```json
{{
  "style": {{
    "title_case_violations": [{{
      "index": 0,
      "text": "original text",
      "issue": "not in Title Case",
      "suggested_fix": "Corrected Title Case Text"
    }}],
    "bullet_terminal_punctuation": [{{
      "index": 2, 
      "text": "bullet text ending.",
      "issue": "bullet ends with period",
      "suggested_fix": "bullet text ending"
    }}],
    "parallelism_issues": [{{
      "indices": [3, 4, 5],
      "issue": "inconsistent verb forms",
      "note": "Mix of gerunds and imperatives in bullet group"
    }}],
    "glossary_violations": [{{
      "index": 1,
      "term": "導入", 
      "expected": "implementation",
      "found": "introduction",
      "context": "system introduction process"
    }}],
    "banned_phrases": [{{
      "index": 6,
      "phrase": "utilize",
      "suggested": "use", 
      "context": "utilize advanced features"
    }}],
    "punctuation_errors": [{{
      "index": 7,
      "issue": "JP punctuation not converted",
      "original": "、",
      "correct": ", "
    }}]
  }}
}}
```

**Review these translations:**
"""

def check_title_case_violations(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect Title Case violations in likely title text."""
    violations = []
    
    for i, text in enumerate(translations):
        # Skip if contains formatting tags that complicate analysis
        clean_text = re.sub(r'\[/?[^\].]+?\]', '', text)
        
        # Heuristic: likely title if short and doesn't end with sentence punctuation
        if len(clean_text.split()) <= 12 and not clean_text.strip().endswith(('.', ':', ';')):
            
            # Apply Title Case while preserving formatting tags
            parts = re.split(r'(\[/?[^\].]+?\])', text)
            corrected_parts = []
            for part in parts:
                if part.startswith('[') and part.endswith(']'):
                    corrected_parts.append(part)
                else:
                    corrected_parts.append(title_case(part))
            corrected = "".join(corrected_parts)

            if text != corrected:
                violations.append({
                    "index": i,
                    "text": text,
                    "issue": "not in Title Case",
                    "suggested_fix": corrected
                })
    
    return violations


```python
def check_bullet_punctuation(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect bullets ending with inappropriate punctuation."""
    violations = []
    
    for i, text in enumerate(translations):
        clean_text = re.sub(r'\[/?[^\\\]]+\]', '', text).strip()
        
        # Check if likely bullet content (not title, has bullet indicators, or is fragment-like)
        has_bullet_tags = '[li-' in text or '•' in text
        is_fragment_like = len(clean_text.split()) < 15 and not clean_text.count('.') > 1
        
        if (has_bullet_tags or is_fragment_like) and re.search(r'[.;:]\s*

def check_glossary_violations(translations: List[str], glossary: Dict[str, str]) -> List[Dict[str, Any]]:
    """Detect glossary term violations."""
    if not glossary:
        return []
    
    violations = []
    
    for i, text in enumerate(translations):
        clean_text = re.sub(r'\[/?[^\]]+\]', '', text).lower()
        
        for jp_term, expected_en in glossary.items():
            expected_lower = expected_en.lower()
            
            # Look for alternative translations that should use glossary term
            # This is simplified - could be more sophisticated with semantic matching
            if jp_term in ["導入", "実装", "システム"]:  # Example key terms
                alternatives = {
                    "導入": ["introduction", "deployment", "rollout"],
                    "実装": ["development", "creation", "building"], 
                    "システム": ["platform", "solution", "tool"]
                }
                
                if jp_term in alternatives:
                    for alt in alternatives[jp_term]:
                        if alt in clean_text and expected_lower not in clean_text:
                            violations.append({
                                "index": i,
                                "term": jp_term,
                                "expected": expected_en,
                                "found": alt,
                                "context": text[:50] + "..."
                            })
                            break
    
    return violations

def check_banned_phrases(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect banned phrases that should be replaced."""
    violations = []
    
    for i, text in enumerate(translations):
        clean_text = re.sub(r'\[/?[^\]]+\]', '', text)
        
        for banned, suggested in BANNED_PHRASES.items():
            pattern = re.compile(r'\b' + re.escape(banned) + r'\b', re.IGNORECASE)
            if pattern.search(clean_text):
                violations.append({
                    "index": i,
                    "phrase": banned,
                    "suggested": suggested,
                    "context": clean_text[:80] + "..."
                })
    
    return violations

def check_punctuation_errors(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect unconverted Japanese punctuation."""
    violations = []
    
    jp_punct_patterns = {
        "、": ", ",
        "。": ". ", 
        "「": '"',
        "」": '"',
        "／": "/",
        "・": "•",
        "　": " "
    }
    
    for i, text in enumerate(translations):
        for jp_char, en_char in jp_punct_patterns.items():
            if jp_char in text:
                violations.append({
                    "index": i,
                    "issue": "Japanese punctuation not converted",
                    "original": jp_char,
                    "correct": en_char,
                    "context": text[:50] + "..."
                })
    
    return violations

def analyze_parallelism(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect parallelism issues in bullet groups (simplified heuristic)."""
    issues = []
    
    # Group consecutive bullets (simplified - assumes bullets are sequential)
    bullet_groups = []
    current_group = []
    
    for i, text in enumerate(translations):
        if '[li-' in text or '•' in text or (len(text.split()) < 15 and not text.endswith('.')):
            current_group.append((i, text))
        else:
            if len(current_group) > 1:
                bullet_groups.append(current_group)
            current_group = []
    
    if len(current_group) > 1:
        bullet_groups.append(current_group)
    
    # Check each group for parallelism
    for group in bullet_groups:
        if len(group) >= 3:  # Only check groups of 3+ bullets
            indices = [idx for idx, _ in group]
            texts = [text for _, text in group]
            
            # Simplified parallelism check: look for mixed verb forms
            starts_with_gerund = sum(1 for t in texts if re.match(r'^\w+ing\b', re.sub(r'\[/?[^\]]+\]', '', t)))
            starts_with_verb = sum(1 for t in texts if re.match(r'^\w+\b', re.sub(r'\[/?[^\]]+\]', '', t)))
            
            if starts_with_gerund > 0 and starts_with_verb > 0 and starts_with_gerund != len(texts):
                issues.append({
                    "indices": indices,
                    "issue": "inconsistent verb forms", 
                    "note": f"Mix of gerunds ({starts_with_gerund}) and other forms in bullet group"
                })
    
    return issues

def run_style_check(translations: List[str], glossary: Dict[str, str] = None) -> Dict[str, Any]:
    """
    Run comprehensive style check and return JSON diagnostics.
    
    Args:
        translations: List of English translation strings
        glossary: Optional Japanese -> English glossary mapping
    
    Returns:
        Dictionary with style diagnostics in structured format
    """
    
    return {
        "style": {
            "title_case_violations": check_title_case_violations(translations),
            "bullet_terminal_punctuation": check_bullet_punctuation(translations),
            "parallelism_issues": analyze_parallelism(translations),
            "glossary_violations": check_glossary_violations(translations, glossary or {}),
            "banned_phrases": check_banned_phrases(translations),
            "punctuation_errors": check_punctuation_errors(translations)
        }
    }

def apply_style_fixes(translations: List[str], diagnostics: Dict[str, Any]) -> List[str]:
    """
    Apply deterministic fixes based on style diagnostics.
    
    Args:
        translations: Original translation list
        diagnostics: Style diagnostics from run_style_check or model
    
    Returns:
        Fixed translations with style issues resolved
    """
    fixed = translations.copy()
    style_issues = diagnostics.get("style", {})
    
    # Apply title case fixes
    for violation in style_issues.get("title_case_violations", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            fixed[index] = violation.get("suggested_fix", fixed[index])
    
    # Fix bullet punctuation
    for violation in style_issues.get("bullet_terminal_punctuation", []):
        index = violation["index"] 
        if 0 <= index < len(fixed):
            fixed[index] = violation.get("suggested_fix", fixed[index])
    
    # Replace banned phrases
    for violation in style_issues.get("banned_phrases", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            old_phrase = violation["phrase"]
            new_phrase = violation["suggested"]
            fixed[index] = re.sub(
                r'\b' + re.escape(old_phrase) + r'\b', 
                new_phrase, 
                fixed[index], 
                flags=re.IGNORECASE
            )
    
    # Fix punctuation errors
    for violation in style_issues.get("punctuation_errors", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            original = violation["original"]
            correct = violation["correct"]
            fixed[index] = fixed[index].replace(original, correct)
    
    # Glossary fixes (simple token replacement)
    for violation in style_issues.get("glossary_violations", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            found_term = violation["found"]
            expected_term = violation["expected"]
            fixed[index] = re.sub(
                r'\b' + re.escape(found_term) + r'\b',
                expected_term,
                fixed[index],
                flags=re.IGNORECASE
            )
    
    return fixed

# Model-based checking (integration with OpenAI)
def model_style_check(client, translations: List[str], glossary: Dict[str, str] = None) -> Dict[str, Any]:
    """
    Use model to perform style checking and return JSON diagnostics.
    
    Args:
        client: OpenAI client instance
        translations: List of translations to check
        glossary: Optional glossary for term consistency
    
    Returns:
        Structured style diagnostics
    """
    prompt = create_style_checker_prompt(translations, glossary)
    
    # Add translations to prompt
    numbered_translations = []
    for i, translation in enumerate(translations):
        numbered_translations.append(f"{i}: {translation}")
    
    full_prompt = prompt + "\n\n" + "\n".join(numbered_translations)
    
    try:
        response = client.responses.create(
            model="gpt-5",
            reasoning_effort="medium",  # Balanced effort for style analysis
            text={"verbosity": "low"},
            input=[{"role": "user", "content": [{"type": "input_text", "text": full_prompt}]}],
            response_format={"type": "json_object"},
            temperature=0.0  # Deterministic for consistent diagnostics
        )
        
        content = getattr(response, "output_text", None)
        if content and callable(content):
            content = content()
        elif hasattr(response, 'choices') and response.choices:
            content = response.choices[0].message.content
        
        return json.loads(content) if content else {"style": {}}
        
    except Exception as e:
        print(f"Style check failed: {e}")
        # Fallback to local style checking
        return run_style_check(translations, glossary), clean_text):
            # Check if it's genuinely multiple sentences that need punctuation
            sentence_count = len(re.findall(r'[.!?]+', clean_text))
            if sentence_count <= 1:
                # Preserve tags by removing punctuation from the original string
                fixed_text = re.sub(r'[.;:]\s*

def check_glossary_violations(translations: List[str], glossary: Dict[str, str]) -> List[Dict[str, Any]]:
    """Detect glossary term violations."""
    if not glossary:
        return []
    
    violations = []
    
    for i, text in enumerate(translations):
        clean_text = re.sub(r'\[/?[^\]]+\]', '', text).lower()
        
        for jp_term, expected_en in glossary.items():
            expected_lower = expected_en.lower()
            
            # Look for alternative translations that should use glossary term
            # This is simplified - could be more sophisticated with semantic matching
            if jp_term in ["導入", "実装", "システム"]:  # Example key terms
                alternatives = {
                    "導入": ["introduction", "deployment", "rollout"],
                    "実装": ["development", "creation", "building"], 
                    "システム": ["platform", "solution", "tool"]
                }
                
                if jp_term in alternatives:
                    for alt in alternatives[jp_term]:
                        if alt in clean_text and expected_lower not in clean_text:
                            violations.append({
                                "index": i,
                                "term": jp_term,
                                "expected": expected_en,
                                "found": alt,
                                "context": text[:50] + "..."
                            })
                            break
    
    return violations

def check_banned_phrases(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect banned phrases that should be replaced."""
    violations = []
    
    for i, text in enumerate(translations):
        clean_text = re.sub(r'\[/?[^\]]+\]', '', text)
        
        for banned, suggested in BANNED_PHRASES.items():
            pattern = re.compile(r'\b' + re.escape(banned) + r'\b', re.IGNORECASE)
            if pattern.search(clean_text):
                violations.append({
                    "index": i,
                    "phrase": banned,
                    "suggested": suggested,
                    "context": clean_text[:80] + "..."
                })
    
    return violations

def check_punctuation_errors(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect unconverted Japanese punctuation."""
    violations = []
    
    jp_punct_patterns = {
        "、": ", ",
        "。": ". ", 
        "「": '"',
        "」": '"',
        "／": "/",
        "・": "•",
        "　": " "
    }
    
    for i, text in enumerate(translations):
        for jp_char, en_char in jp_punct_patterns.items():
            if jp_char in text:
                violations.append({
                    "index": i,
                    "issue": "Japanese punctuation not converted",
                    "original": jp_char,
                    "correct": en_char,
                    "context": text[:50] + "..."
                })
    
    return violations

def analyze_parallelism(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect parallelism issues in bullet groups (simplified heuristic)."""
    issues = []
    
    # Group consecutive bullets (simplified - assumes bullets are sequential)
    bullet_groups = []
    current_group = []
    
    for i, text in enumerate(translations):
        if '[li-' in text or '•' in text or (len(text.split()) < 15 and not text.endswith('.')):
            current_group.append((i, text))
        else:
            if len(current_group) > 1:
                bullet_groups.append(current_group)
            current_group = []
    
    if len(current_group) > 1:
        bullet_groups.append(current_group)
    
    # Check each group for parallelism
    for group in bullet_groups:
        if len(group) >= 3:  # Only check groups of 3+ bullets
            indices = [idx for idx, _ in group]
            texts = [text for _, text in group]
            
            # Simplified parallelism check: look for mixed verb forms
            starts_with_gerund = sum(1 for t in texts if re.match(r'^\w+ing\b', re.sub(r'\[/?[^\]]+\]', '', t)))
            starts_with_verb = sum(1 for t in texts if re.match(r'^\w+\b', re.sub(r'\[/?[^\]]+\]', '', t)))
            
            if starts_with_gerund > 0 and starts_with_verb > 0 and starts_with_gerund != len(texts):
                issues.append({
                    "indices": indices,
                    "issue": "inconsistent verb forms", 
                    "note": f"Mix of gerunds ({starts_with_gerund}) and other forms in bullet group"
                })
    
    return issues

def run_style_check(translations: List[str], glossary: Dict[str, str] = None) -> Dict[str, Any]:
    """
    Run comprehensive style check and return JSON diagnostics.
    
    Args:
        translations: List of English translation strings
        glossary: Optional Japanese -> English glossary mapping
    
    Returns:
        Dictionary with style diagnostics in structured format
    """
    
    return {
        "style": {
            "title_case_violations": check_title_case_violations(translations),
            "bullet_terminal_punctuation": check_bullet_punctuation(translations),
            "parallelism_issues": analyze_parallelism(translations),
            "glossary_violations": check_glossary_violations(translations, glossary or {}),
            "banned_phrases": check_banned_phrases(translations),
            "punctuation_errors": check_punctuation_errors(translations)
        }
    }

def apply_style_fixes(translations: List[str], diagnostics: Dict[str, Any]) -> List[str]:
    """
    Apply deterministic fixes based on style diagnostics.
    
    Args:
        translations: Original translation list
        diagnostics: Style diagnostics from run_style_check or model
    
    Returns:
        Fixed translations with style issues resolved
    """
    fixed = translations.copy()
    style_issues = diagnostics.get("style", {})
    
    # Apply title case fixes
    for violation in style_issues.get("title_case_violations", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            fixed[index] = violation.get("suggested_fix", fixed[index])
    
    # Fix bullet punctuation
    for violation in style_issues.get("bullet_terminal_punctuation", []):
        index = violation["index"] 
        if 0 <= index < len(fixed):
            fixed[index] = violation.get("suggested_fix", fixed[index])
    
    # Replace banned phrases
    for violation in style_issues.get("banned_phrases", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            old_phrase = violation["phrase"]
            new_phrase = violation["suggested"]
            fixed[index] = re.sub(
                r'\b' + re.escape(old_phrase) + r'\b', 
                new_phrase, 
                fixed[index], 
                flags=re.IGNORECASE
            )
    
    # Fix punctuation errors
    for violation in style_issues.get("punctuation_errors", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            original = violation["original"]
            correct = violation["correct"]
            fixed[index] = fixed[index].replace(original, correct)
    
    # Glossary fixes (simple token replacement)
    for violation in style_issues.get("glossary_violations", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            found_term = violation["found"]
            expected_term = violation["expected"]
            fixed[index] = re.sub(
                r'\b' + re.escape(found_term) + r'\b',
                expected_term,
                fixed[index],
                flags=re.IGNORECASE
            )
    
    return fixed

# Model-based checking (integration with OpenAI)
def model_style_check(client, translations: List[str], glossary: Dict[str, str] = None) -> Dict[str, Any]:
    """
    Use model to perform style checking and return JSON diagnostics.
    
    Args:
        client: OpenAI client instance
        translations: List of translations to check
        glossary: Optional glossary for term consistency
    
    Returns:
        Structured style diagnostics
    """
    prompt = create_style_checker_prompt(translations, glossary)
    
    # Add translations to prompt
    numbered_translations = []
    for i, translation in enumerate(translations):
        numbered_translations.append(f"{i}: {translation}")
    
    full_prompt = prompt + "\n\n" + "\n".join(numbered_translations)
    
    try:
        response = client.responses.create(
            model="gpt-5",
            reasoning_effort="medium",  # Balanced effort for style analysis
            text={"verbosity": "low"},
            input=[{"role": "user", "content": [{"type": "input_text", "text": full_prompt}]}],
            response_format={"type": "json_object"},
            temperature=0.0  # Deterministic for consistent diagnostics
        )
        
        content = getattr(response, "output_text", None)
        if content and callable(content):
            content = content()
        elif hasattr(response, 'choices') and response.choices:
            content = response.choices[0].message.content
        
        return json.loads(content) if content else {"style": {}}
        
    except Exception as e:
        print(f"Style check failed: {e}")
        # Fallback to local style checking
        return run_style_check(translations, glossary), '', text.rstrip()).rstrip()
                violations.append({
                    "index": i,
                    "text": text,
                    "issue": "bullet ends with terminal punctuation", 
                    "suggested_fix": fixed_text
                })
    
    return violations
```

def check_glossary_violations(translations: List[str], glossary: Dict[str, str]) -> List[Dict[str, Any]]:
    """Detect glossary term violations."""
    if not glossary:
        return []
    
    violations = []
    
    for i, text in enumerate(translations):
        clean_text = re.sub(r'\[/?[^\]]+\]', '', text).lower()
        
        for jp_term, expected_en in glossary.items():
            expected_lower = expected_en.lower()
            
            # Look for alternative translations that should use glossary term
            # This is simplified - could be more sophisticated with semantic matching
            if jp_term in ["導入", "実装", "システム"]:  # Example key terms
                alternatives = {
                    "導入": ["introduction", "deployment", "rollout"],
                    "実装": ["development", "creation", "building"], 
                    "システム": ["platform", "solution", "tool"]
                }
                
                if jp_term in alternatives:
                    for alt in alternatives[jp_term]:
                        if alt in clean_text and expected_lower not in clean_text:
                            violations.append({
                                "index": i,
                                "term": jp_term,
                                "expected": expected_en,
                                "found": alt,
                                "context": text[:50] + "..."
                            })
                            break
    
    return violations

def check_banned_phrases(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect banned phrases that should be replaced."""
    violations = []
    
    for i, text in enumerate(translations):
        clean_text = re.sub(r'\[/?[^\]]+\]', '', text)
        
        for banned, suggested in BANNED_PHRASES.items():
            pattern = re.compile(r'\b' + re.escape(banned) + r'\b', re.IGNORECASE)
            if pattern.search(clean_text):
                violations.append({
                    "index": i,
                    "phrase": banned,
                    "suggested": suggested,
                    "context": clean_text[:80] + "..."
                })
    
    return violations

def check_punctuation_errors(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect unconverted Japanese punctuation."""
    violations = []
    
    jp_punct_patterns = {
        "、": ", ",
        "。": ". ", 
        "「": '"',
        "」": '"',
        "／": "/",
        "・": "•",
        "　": " "
    }
    
    for i, text in enumerate(translations):
        for jp_char, en_char in jp_punct_patterns.items():
            if jp_char in text:
                violations.append({
                    "index": i,
                    "issue": "Japanese punctuation not converted",
                    "original": jp_char,
                    "correct": en_char,
                    "context": text[:50] + "..."
                })
    
    return violations

def analyze_parallelism(translations: List[str]) -> List[Dict[str, Any]]:
    """Detect parallelism issues in bullet groups (simplified heuristic)."""
    issues = []
    
    # Group consecutive bullets (simplified - assumes bullets are sequential)
    bullet_groups = []
    current_group = []
    
    for i, text in enumerate(translations):
        if '[li-' in text or '•' in text or (len(text.split()) < 15 and not text.endswith('.')):
            current_group.append((i, text))
        else:
            if len(current_group) > 1:
                bullet_groups.append(current_group)
            current_group = []
    
    if len(current_group) > 1:
        bullet_groups.append(current_group)
    
    # Check each group for parallelism
    for group in bullet_groups:
        if len(group) >= 3:  # Only check groups of 3+ bullets
            indices = [idx for idx, _ in group]
            texts = [text for _, text in group]
            
            # Simplified parallelism check: look for mixed verb forms
            starts_with_gerund = sum(1 for t in texts if re.match(r'^\w+ing\b', re.sub(r'\[/?[^\]]+\]', '', t)))
            starts_with_verb = sum(1 for t in texts if re.match(r'^\w+\b', re.sub(r'\[/?[^\]]+\]', '', t)))
            
            if starts_with_gerund > 0 and starts_with_verb > 0 and starts_with_gerund != len(texts):
                issues.append({
                    "indices": indices,
                    "issue": "inconsistent verb forms", 
                    "note": f"Mix of gerunds ({starts_with_gerund}) and other forms in bullet group"
                })
    
    return issues

def run_style_check(translations: List[str], glossary: Dict[str, str] = None) -> Dict[str, Any]:
    """
    Run comprehensive style check and return JSON diagnostics.
    
    Args:
        translations: List of English translation strings
        glossary: Optional Japanese -> English glossary mapping
    
    Returns:
        Dictionary with style diagnostics in structured format
    """
    
    return {
        "style": {
            "title_case_violations": check_title_case_violations(translations),
            "bullet_terminal_punctuation": check_bullet_punctuation(translations),
            "parallelism_issues": analyze_parallelism(translations),
            "glossary_violations": check_glossary_violations(translations, glossary or {}),
            "banned_phrases": check_banned_phrases(translations),
            "punctuation_errors": check_punctuation_errors(translations)
        }
    }

def apply_style_fixes(translations: List[str], diagnostics: Dict[str, Any]) -> List[str]:
    """
    Apply deterministic fixes based on style diagnostics.
    
    Args:
        translations: Original translation list
        diagnostics: Style diagnostics from run_style_check or model
    
    Returns:
        Fixed translations with style issues resolved
    """
    fixed = translations.copy()
    style_issues = diagnostics.get("style", {})
    
    # Apply title case fixes
    for violation in style_issues.get("title_case_violations", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            fixed[index] = violation.get("suggested_fix", fixed[index])
    
    # Fix bullet punctuation
    for violation in style_issues.get("bullet_terminal_punctuation", []):
        index = violation["index"] 
        if 0 <= index < len(fixed):
            fixed[index] = violation.get("suggested_fix", fixed[index])
    
    # Replace banned phrases
    for violation in style_issues.get("banned_phrases", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            old_phrase = violation["phrase"]
            new_phrase = violation["suggested"]
            fixed[index] = re.sub(
                r'\b' + re.escape(old_phrase) + r'\b', 
                new_phrase, 
                fixed[index], 
                flags=re.IGNORECASE
            )
    
    # Fix punctuation errors
    for violation in style_issues.get("punctuation_errors", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            original = violation["original"]
            correct = violation["correct"]
            fixed[index] = fixed[index].replace(original, correct)
    
    # Glossary fixes (simple token replacement)
    for violation in style_issues.get("glossary_violations", []):
        index = violation["index"]
        if 0 <= index < len(fixed):
            found_term = violation["found"]
            expected_term = violation["expected"]
            fixed[index] = re.sub(
                r'\b' + re.escape(found_term) + r'\b',
                expected_term,
                fixed[index],
                flags=re.IGNORECASE
            )
    
    return fixed

# Model-based checking (integration with OpenAI)
def model_style_check(client, translations: List[str], glossary: Dict[str, str] = None) -> Dict[str, Any]:
    """
    Use model to perform style checking and return JSON diagnostics.
    
    Args:
        client: OpenAI client instance
        translations: List of translations to check
        glossary: Optional glossary for term consistency
    
    Returns:
        Structured style diagnostics
    """
    prompt = create_style_checker_prompt(translations, glossary)
    
    # Add translations to prompt
    numbered_translations = []
    for i, translation in enumerate(translations):
        numbered_translations.append(f"{i}: {translation}")
    
    full_prompt = prompt + "\n\n" + "\n".join(numbered_translations)
    
    try:
        response = client.responses.create(
            model="gpt-5",
            reasoning_effort="medium",  # Balanced effort for style analysis
            text={"verbosity": "low"},
            input=[{"role": "user", "content": [{"type": "input_text", "text": full_prompt}]}],
            response_format={"type": "json_object"},
            temperature=0.0  # Deterministic for consistent diagnostics
        )
        
        content = getattr(response, "output_text", None)
        if content and callable(content):
            content = content()
        elif hasattr(response, 'choices') and response.choices:
            content = response.choices[0].message.content
        
        return json.loads(content) if content else {"style": {}}
        
    except Exception as e:
        print(f"Style check failed: {e}")
        # Fallback to local style checking
        return run_style_check(translations, glossary)