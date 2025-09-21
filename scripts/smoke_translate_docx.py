#!/usr/bin/env python3
"""
Smoke test CLI for translate_docx.py on fixtures.
Verifies XML structure parity >95% after translation.
"""

import argparse
import subprocess
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from difflib import SequenceMatcher

def get_structure_xml(docx_path: Path) -> str:
    """Extract and serialize XML structure ignoring text content."""
    xml_contents = []
    with zipfile.ZipFile(docx_path) as z:
        for name in sorted(z.namelist()):
            if name.endswith('.xml') and not name.startswith('word/media/'):
                try:
                    content = z.read(name).decode('utf-8')
                    root = ET.fromstring(content)
                    def remove_text(elem):
                        elem.text = None
                        elem.tail = None
                        for child in elem:
                            remove_text(child)
                    remove_text(root)
                    struct = ET.tostring(root, encoding='unicode', method='xml')
                    xml_contents.append(f"--- {name} ---\n{struct}")
                except Exception:
                    # Skip invalid XML
                    pass
    return '\n'.join(xml_contents)

def compute_parity(input_path: Path, output_path: Path) -> float:
    """Compute structure similarity ratio between input and output docx."""
    struct_in = get_structure_xml(input_path)
    struct_out = get_structure_xml(output_path)
    return SequenceMatcher(None, struct_in, struct_out).ratio()

def main():
    parser = argparse.ArgumentParser(description="Smoke test for DOCX translation")
    parser.add_argument('--input', required=True, help='Input fixture DOCX path')
    parser.add_argument('--output', required=True, help='Output path for translated DOCX')
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return 1

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Run translation
    print("Running translation...")
    cmd = [
        'python', 'scripts/translate_docx.py',
        '--in', str(input_path),
        '--out', str(output_path),
        '--batch', '1'  # Small batch for testing
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, env=os.environ)
    if result.returncode != 0:
        print(f"Translation failed:\n{result.stderr}")
        return 1

    if not output_path.exists():
        print("Output file not created")
        return 1

    # Compute parity
    print("Computing XML structure parity...")
    parity = compute_parity(input_path, output_path)
    print(f"XML structure parity: {parity:.2%}")

    if parity < 0.95:
        print("FAIL: Structure parity below 95%")
        return 1

    # Collect samples
    smoke_dir = Path('tmp/smoke_out')
    smoke_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = input_path.stem
    in_sample = smoke_dir / f"{base_name}_input_{timestamp}.docx"
    out_sample = smoke_dir / f"{base_name}_output_{timestamp}.docx"
    shutil.copy2(input_path, in_sample)
    shutil.copy2(output_path, out_sample)
    print(f"Samples collected in {smoke_dir}")
    print("Smoke test PASSED")
    return 0

if __name__ == "__main__":
    sys.exit(main())
