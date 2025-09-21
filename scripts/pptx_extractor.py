import re
import zipfile
from xml.etree import ElementTree as ET
from .base_extractor import BaseExtractor
from typing import List, Tuple, Optional, Dict, Set

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

def normalize_para_text(p_el) -> str:
    br_tag = A_NS + "br"
    t_tag = A_NS + "t"
    r_tag = A_NS + "r"

    parts = []
    for node in p_el:
        if node.tag == r_tag:
            t = node.find(t_tag)
            parts.append("" if t is None or t.text is None else t.text)
        elif node.tag == br_tag:
            parts.append("\\n")
        else:
            t = node.find(f".//{t_tag}")
            if t is not None and t.text:
                parts.append(t.text)

    return "".join(parts)

class PptxExtractor(BaseExtractor):
    def extract(self, file_path: str, page_range: Optional[str] = None) -> List[Tuple[str, int, str, Dict]]:
        with zipfile.ZipFile(file_path, "r") as z:
            paras = []
            slide_files = sorted([n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")])

            slide_range_set: Set[int] = set()
            if page_range:
                parts = page_range.split('-')
                if len(parts) == 2:
                    start, end = int(parts[0]), int(parts[1])
                    slide_range_set = set(range(start, end + 1))

            filtered_slides = []
            for sf in slide_files:
                match = re.search(r'slide(\\d+)\\.xml', sf)
                if match and (not slide_range_set or int(match.group(1)) in slide_range_set):
                    filtered_slides.append(sf)
            slide_files = filtered_slides

            for sf in slide_files:
                root = ET.fromstring(z.read(sf))
                for idx, p_el in enumerate(root.iter(A_NS + "p")):
                    text = normalize_para_text(p_el)
                    if text.strip():
                        layout = {}  # Can add font info from rPr if needed
                        paras.append((sf, idx, text, layout))
            return paras