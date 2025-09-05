#!/usr/bin/env python3
"""
Derive a deck's tone fingerprint from a sample of Japanese slides.
Writes the result to deck_tone.json.
"""

import argparse, json, os, re, sys, zipfile
from openai import OpenAI

# This is a simplified version of the text extraction logic from
# estimate_pptx_translation_cost.py
JP_RX = re.compile(r"[\u3040-\u30ff\u3400-\u9fff々〆ヵヶ]")
BR_RX = re.compile(r"\n||\n")

def extract_text_blocks(pptx_path: str):
    """Return a list of visible text blocks (titles, bullets, notes) from the PPTX."""
    A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    blocks = []
    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_names = sorted([n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")])
        for name in slide_names:
            xml = z.read(name).decode("utf-8", errors="ignore")
            texts = re.findall(rf"<a:t>(.*?)</a:t>", xml, flags=re.S)
            if texts:
                s = BR_RX.sub("\n",  ".join(t.replace("&lt;","<").replace("&gt;">").replace("&amp;",&) for t in texts))
                parts = [p.strip() for p in re.split(r"\n{2,}", s) if p.strip()]
                blocks.extend(parts)
    return [b for b in blocks if b.strip()]

def derive_tone_fingerprint(client, text_sample):
    """
    Uses an LLM to derive a tone fingerprint from a sample of Japanese text.
    """
    prompt = f"""Analyze the following Japanese text from a presentation and return a JSON object describing its tone and style. Focus on the overall impression a native speaker would have.

    **Required JSON format:**
    ```json
    {{
      "register": "desu_masu" | "de_aru" | "plain",
      "formality": 1-5,
      "directness": 1-5,
      "persuasiveness": 1-5,
      "technicality": 1-5,
      "audience": "internal_exec | internal_staff | external_clients | general",
      "style_notes": ["short headlines","dense bullets","directive","neutral"]
    }}
    ```

    **Text sample:**
    {text_sample}
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a linguistic analyst specializing in Japanese business communication."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.0
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"Error calling OpenAI API: {e}", file=sys.stderr)
        return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", help="Path to PPTX file to analyze")
    ap.add_argument("--sample-size", type=int, default=10, help="Number of text blocks to sample")
    ap.add_argument("--out", default="deck_tone.json", help="Output JSON file path")
    args = ap.parse_args()

    if not os.environ.get("OPENAI_API_KEY"):
        print("ERROR: OPENAI_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    client = OpenAI()

    blocks = extract_text_blocks(args.pptx)
    if not blocks:
        print("No text found in the presentation.", file=sys.stderr)
        sys.exit(1)

    jp_blocks = [b for b in blocks if JP_RX.search(b)]
    if not jp_blocks:
        print("No Japanese text found in the presentation.", file=sys.stderr)
        sys.exit(1)

    sample = "\n\n---\n\n".join(jp_blocks[:args.sample_size])

    print(f"Deriving tone fingerprint from {len(jp_blocks[:args.sample_size])} Japanese text blocks...")
    fingerprint = derive_tone_fingerprint(client, sample)

    if fingerprint:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(fingerprint, f, ensure_ascii=False, indent=2)
        print(f"Successfully wrote tone fingerprint to {args.out}")
    else:
        print("Failed to derive tone fingerprint.", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
