#!/usr/bin/env python3
"""
Single PPTX formatting profile for consistent slide appearance.
Applies uniform fonts, sizes, spacing, and layout across entire deck.
"""
import xml.etree.ElementTree as ET

# Namespace constants
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

def apply_textframe_profile_xml(root, is_title=False):
    """
    Apply consistent formatting profile to XML textframe elements.
    Works directly with PPTX XML structure for maximum control.
    
    Args:
        root: XML root element containing text content
        is_title: True for title text frames, False for body content
    """
    # Find all text bodies and apply formatting
    for txBody in root.iter(A_NS + "txBody"):
        _apply_body_formatting(txBody, is_title)
        
        # Apply paragraph and run formatting
        for para in txBody.iter(A_NS + "p"):
            _apply_paragraph_formatting(para, is_title)
            
            # Apply run formatting for font consistency
            for run in para.iter(A_NS + "r"):
                _apply_run_formatting(run, is_title)

def _apply_body_formatting(txBody, is_title=False):
    """Apply text body properties for margins and autofit."""
    bodyPr = txBody.find(A_NS + "bodyPr")
    if bodyPr is None:
        bodyPr = ET.SubElement(txBody, A_NS + "bodyPr")
    
    # Tight margins for maximum text space (values in EMUs: 1pt = 12700 EMUs)
    bodyPr.set("lIns", "25400")   # Left margin: 2pt  
    bodyPr.set("rIns", "25400")   # Right margin: 2pt
    bodyPr.set("tIns", "12700")   # Top margin: 1pt
    bodyPr.set("bIns", "12700")   # Bottom margin: 1pt
    bodyPr.set("wrap", "square")  # Enable text wrapping
    bodyPr.set("rtlCol", "0")     # Left-to-right text direction
    
    # Enable shrink-to-fit with sensible limits
    normAutofit = bodyPr.find(A_NS + "normAutofit")
    if normAutofit is None:
        normAutofit = ET.SubElement(bodyPr, A_NS + "normAutofit")
    
    # Set scaling limits to prevent unreadable text
    if is_title:
        normAutofit.set("fontScale", "90000")      # Minimum 90% font scaling for titles
        normAutofit.set("lnSpcReduction", "10000") # Maximum 10% line spacing reduction
    else:
        normAutofit.set("fontScale", "85000")      # Minimum 85% font scaling for body
        normAutofit.set("lnSpcReduction", "15000") # Maximum 15% line spacing reduction

def _apply_paragraph_formatting(para, is_title=False):
    """Apply paragraph-level formatting for spacing and indentation."""
    pPr = para.find(A_NS + "pPr")
    if pPr is None:
        pPr = ET.SubElement(para, A_NS + "pPr")
    
    # Optimize line spacing (110% vs default ~120%)
    lnSpc = pPr.find(A_NS + "lnSpc")
    if lnSpc is None:
        lnSpc = ET.SubElement(pPr, A_NS + "lnSpc")
    
    spcPct = lnSpc.find(A_NS + "spcPct")
    if spcPct is None:
        spcPct = ET.SubElement(lnSpc, A_NS + "spcPct")
    
    line_spacing = "105000" if is_title else "110000"  # 105% for titles, 110% for body
    spcPct.set("val", line_spacing)
    
    # Remove paragraph spacing before/after
    for spacing_elem in [A_NS + "spcBef", A_NS + "spcAft"]:
        existing = pPr.find(spacing_elem)
        if existing is not None:
            pPr.remove(existing)
    
    # Apply bullet indentation based on level
    level = int(pPr.get("lvl", "0"))
    _apply_bullet_indentation(pPr, level, is_title)

def _apply_bullet_indentation(pPr, level, is_title=False):
    """Apply consistent bullet indentation based on level."""
    if is_title:
        # Titles typically don't have bullets, but if they do, minimal indent
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        return
    
    # Bullet indentation in EMUs (1 inch = 914400 EMUs)
    base_indent = 274320  # ~0.3 inch base indent
    hanging_indent = 182880  # ~0.2 inch hanging indent
    
    if level == 0:
        # First level bullets (hanging indent)
        pPr.set("marL", "0")
        pPr.set("indent", f"-{hanging_indent}")
    elif level == 1:
        # Second level bullets  
        left_margin = base_indent
        pPr.set("marL", str(left_margin))
        pPr.set("indent", f"-{hanging_indent}")
    else:
        # Third level and beyond
        left_margin = base_indent + (level - 1) * 182880
        pPr.set("marL", str(left_margin))
        pPr.set("indent", f"-{hanging_indent}")

def _apply_run_formatting(run, is_title=False):
    """Apply run-level formatting for fonts and sizes."""
    rPr = run.find(A_NS + "rPr")
    if rPr is None:
        rPr = ET.SubElement(run, A_NS + "rPr")
    
    # Set minimum font sizes (in half-points: 1pt = 100 half-points)
    current_size = rPr.get("sz")
    min_size = 1800 if is_title else 1100  # 18pt for titles, 11pt for body
    if current_size is None:
        rPr.set("sz", str(min_size))
    else:
        try:
            if int(current_size) < min_size:
                rPr.set("sz", str(min_size))
        except ValueError:
            rPr.set("sz", str(min_size))

    # Ensure font family is set for consistency
    latin = rPr.find(A_NS + "latin")
    if latin is None:
        latin = ET.SubElement(rPr, A_NS + "latin")

    # Use brand font if not already specified
    if not latin.get("typeface"):
        brand_font = "Inter"  # Default professional font, can be customized
        latin.set("typeface", brand_font)
def apply_deck_formatting_profile(root):
    """
    Apply formatting profile to entire slide, detecting content types.
    
    Args:
        root: XML root element of slide
    """
    # Track if we've found title elements
    title_found = False
    
    # Look for title placeholders and apply title formatting
    for shape in root.iter():
        if shape.tag.endswith("}sp"):  # Shape element
            # Check for title indicators in shape properties
            nvSpPr = shape.find(".//" + P_NS + "nvSpPr")
            if nvSpPr is not None:
                nvPr = nvSpPr.find(P_NS + "nvPr")
                if nvPr is not None:
                    ph = nvPr.find(P_NS + "ph")
                    if ph is not None and ph.get("type") in ["title", "ctrTitle"]:
                        title_found = True
                        # Apply title formatting to this text frame
                        txBody = shape.find(".//" + A_NS + "txBody")
                        if txBody is not None:
                            apply_textframe_profile_xml(shape, is_title=True)
    
    # Apply body formatting to all other text frames
    for txBody in root.iter(A_NS + "txBody"):
        # Skip if we already processed this as a title
        parent_shape = txBody
        while parent_shape is not None and not parent_shape.tag.endswith("}sp"):
            parent_shape = parent_shape.getparent() if hasattr(parent_shape, 'getparent') else None
        
        is_title_frame = False
        if parent_shape is not None:
            nvSpPr = parent_shape.find(".//" + P_NS + "nvSpPr")
            if nvSpPr is not None:
                nvPr = nvSpPr.find(P_NS + "nvPr") 
                if nvPr is not None:
                    ph = nvPr.find(P_NS + "ph")
                    if ph is not None and ph.get("type") in ["title", "ctrTitle"]:
                        is_title_frame = True
        
        if not is_title_frame:
            apply_textframe_profile_xml(parent_shape or txBody.getparent(), is_title=False)

def get_formatting_statistics(root) -> dict:
    """
    Analyze formatting consistency across slide.
    Returns statistics for audit purposes.
    """
    stats = {
        "text_frames": 0,
        "title_frames": 0,
        "body_frames": 0,
        "font_families": set(),
        "font_sizes": [],
        "margin_settings": [],
        "line_spacings": []
    }
    
    for txBody in root.iter(A_NS + "txBody"):
        stats["text_frames"] += 1
        
        # Check body properties
        bodyPr = txBody.find(A_NS + "bodyPr")
        if bodyPr is not None:
            margins = {
                "left": bodyPr.get("lIns", "default"),
                "right": bodyPr.get("rIns", "default"), 
                "top": bodyPr.get("tIns", "default"),
                "bottom": bodyPr.get("bIns", "default")
            }
            stats["margin_settings"].append(margins)
        
        # Check paragraph properties
        for para in txBody.iter(A_NS + "p"):
            pPr = para.find(A_NS + "pPr")
            if pPr is not None:
                lnSpc = pPr.find(A_NS + "lnSpc")
                if lnSpc is not None:
                    spcPct = lnSpc.find(A_NS + "spcPct")
                    if spcPct is not None:
                        stats["line_spacings"].append(spcPct.get("val", "default"))
            
            # Check run properties
            for run in para.iter(A_NS + "r"):
                rPr = run.find(A_NS + "rPr")
                if rPr is not None:
                    size = rPr.get("sz")
                    if size:
                        stats["font_sizes"].append(int(size))
                    
                    latin = rPr.find(A_NS + "latin")
                    if latin is not None:
                        font_family = latin.get("typeface")
                        if font_family:
                            stats["font_families"].add(font_family)
    
    return stats

# Integration helpers
def format_slide_xml(slide_xml_data: bytes) -> bytes:
    """
    Apply formatting profile to slide XML data.
    
    Args:
        slide_xml_data: Raw slide XML bytes
        
    Returns:
        Formatted slide XML bytes
    """
    root = ET.fromstring(slide_xml_data)
    apply_deck_formatting_profile(root)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def should_apply_formatting(content_changed: bool = True) -> bool:
    """Determine if formatting should be applied based on content changes."""
    # Always apply formatting when content has changed to ensure consistency
    return content_changed