import re
from xml.etree import ElementTree as ET

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

def normalize_para_text(p_el):
    """Extract full visible text for a paragraph (concatenate runs, insert '\n' for a:br)."""
    br_tag = A_NS + "br"
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"

    parts = []
    for node in p_el:
        if node.tag == r_tag:
            t = node.find(t_tag)
            parts.append("" if t is None or t.text is None else t.text)
        elif node.tag == br_tag:
            parts.append("\n")
        else:
            t = node.find(f".//{t_tag}")
            if t is not None and t.text:
                parts.append(t.text)

    return "".join(parts)

def set_para_text(p_el, new_text: str):
    """Word-aware replacement. Preserves word boundaries and turns '\n' into <a:br/>."""
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"
    br_tag = A_NS + "br"

    # Collect runs (preserve overall styling distribution), clear <a:br/> and run text
    runs = [child for child in p_el if child.tag == r_tag]
    if not runs:
        r = ET.Element(r_tag)
        ET.SubElement(r, t_tag).text = ""
        p_el.insert(0, r)
        runs = [r]

    for child in list(p_el):
        if child.tag == br_tag:
            p_el.remove(child)
    for r in runs:
        t = r.find(t_tag)
        if t is None:
            t = ET.SubElement(r, t_tag)
        t.text = ""

    # Tokenize: keep whitespace; use None sentinel for newline
    def tokenize(s): return re.findall(r"\S+|\s+", s)
    tokens = []
    lines = new_text.split("\n")
    for i, line in enumerate(lines):
        tokens.extend(tokenize(line))
        if i < len(lines) - 1:
            tokens.append(None)  # newline marker

    # Single run: dump text, insert <a:br/> at markers
    if len(runs) == 1:
        t = runs[0].find(t_tag)
        buf = []
        br_count = 0
        for tok in tokens:
            if tok is None:
                # Insert <a:br/> after the run
                br = ET.Element(br_tag)
                run_idx = list(p_el).index(runs[0])
                p_el.insert(run_idx + 1 + br_count, br)
                br_count += 1
            else:
                buf.append(tok)
        t.text = "".join(buf).strip()
        return

    # Multi-run: distribute on word boundaries proportional to original text lengths
    orig_lens = [len((r.find(t_tag).text or "")) for r in runs]
    total_words = sum(len(x) for x in tokens if isinstance(x, str))
    total_base = sum(orig_lens) or total_words or 1
    targets = []
    acc = 0
    for L in orig_lens:
        share = round(total_words * (L / total_base))
        targets.append(share)
        acc += share
    if targets:
        targets[-1] += (total_words - acc)  # fix rounding drift

    def consume(n_chars):
        taken = []
        count = 0
        while tokens:
            tok = tokens[0]
            if tok is None:  # stop before newline; caller will insert <a:br/>
                break
            need = len(tok)
            # respect word boundaries
            if count > 0 and not tok.isspace() and count + need > n_chars:
                break
            taken.append(tokens.pop(0))
            count += need
            if tokens and tokens[0] is None:
                break
        return "".join(taken)

    # Fill each run, inserting <a:br/> exactly where newlines occur
    for r, n in zip(runs, targets):
        t = r.find(t_tag)
        t.text = consume(n)
        while tokens and tokens[0] is None:
            tokens.pop(0)
            br = ET.Element(br_tag)
            run_idx = list(p_el).index(r)
            p_el.insert(run_idx + 1, br)

    # Any leftovers go into the last run
    if tokens:
        tail = "".join(tok for tok in tokens if isinstance(tok, str))
        last_t = runs[-1].find(t_tag)
        last_t.text = (last_t.text or "") + tail

def _ensure_autofit(root):
    # For every txBody, ensure <a:bodyPr><a:normAutofit/></a:bodyPr>
    for tx in root.iter(A_NS + "txBody"):
        bodyPr = tx.find(A_NS + "bodyPr")
        if bodyPr is None:
            bodyPr = ET.SubElement(tx, A_NS + "bodyPr")
        if bodyPr.find(A_NS + "normAutofit") is None and bodyPr.find(A_NS + "spAutoFit") is None:
            ET.SubElement(bodyPr, A_NS + "normAutofit")

def _ensure_autofit_on_tree(root, mode: str, font_scale_min: int, line_spacing_pct: int, tight_margins: bool):
    """Enable slide-wide text autofit & spacing. Safe, mechanical; does not change wording."""
    bodyPr_tag = A_NS + "bodyPr"
    norm_tag   = A_NS + "normAutofit"
    shape_tag  = A_NS + "spAutoFit"
    no_tag     = A_NS + "noAutofit"
    p_tag      = A_NS + "p"
    pPr_tag    = A_NS + "pPr"
    lnSpc_tag  = A_NS + "lnSpc"
    spcPct_tag = A_NS + "spcPct"

    # 1) Per text frame: set autofit and tighten insets if requested
    for bp in list(root.iter(bodyPr_tag)):
        # remove conflicting children
        for ch in list(bp):
            if ch.tag in (norm_tag, shape_tag, no_tag):
                bp.remove(ch)
        if mode == "none":
            bp.append(ET.Element(no_tag))
        elif mode == "shape":
            bp.append(ET.Element(shape_tag))
        else:
            na = ET.Element(norm_tag)
            # allow PowerPoint to shrink fonts and line spacing to fit
            na.set("fontScale", str(font_scale_min))           # e.g., 90000 = 90%
            na.set("lnSpcReduction", "12000")                  # ~12% line-space reduction headroom
            bp.append(na)
        if tight_margins:
            # Insets: values are EMUs; these are conservative, non-zero margins
            bp.set("lIns", "45720")   # ~0.05"
            bp.set("rIns", "45720")
            bp.set("tIns", "22860")   # ~0.025"
            bp.set("bIns", "22860")

    # 2) Normalize line spacing to a fixed percentage (e.g., 100%)
    for p in list(root.iter(p_tag)):
        pPr = p.find(pPr_tag)
        if pPr is None:
            pPr = ET.SubElement(p, pPr_tag)
        ln = pPr.find(lnSpc_tag)
        if ln is None:
            ln = ET.SubElement(pPr, lnSpc_tag)
        sp = ln.find(spcPct_tag)
        if sp is None:
            sp = ET.SubElement(ln, spcPct_tag)
        sp.set("val", str(line_spacing_pct))
