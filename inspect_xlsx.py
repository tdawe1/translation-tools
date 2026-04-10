import pandas as pd
import sys

file_path = '/home/thomas/translation-tools/translations-pptx-pipeline/outputs/6938c4aec799f_baee40f0d56211f0b374ca6fa6c9eb86.xlsx'

try:
    df = pd.read_excel(file_path)
    
    # Iterate through rows and extract relevant content
    # Assuming the structure observed: Col 2 is JP, Col 3 is AI En, Col 4 is Native En
    # Column indices in df (0-based) match Unnamed: x (where x is the index if we account for the index col?)
    # Let's use iloc to be safe. 
    # Col 0: Unnamed: 0
    # Col 1: Unnamed: 1
    # Col 2: Unnamed: 2 (JP)
    # Col 3: Unnamed: 3 (En AI)
    # Col 4: Unnamed: 4 (En Native)
    
    results = []
    for index, row in df.iterrows():
        jp_text = row['Unnamed: 2']
        en_ai = row['Unnamed: 3']
        en_native = row['Unnamed: 4']
        
        # Check if jp_text is valid and not a header
        if pd.notna(jp_text) and str(jp_text).strip() != "日本語":
            results.append({
                "row": index,
                "japanese": str(jp_text).strip(),
                "english_ai": str(en_ai).strip() if pd.notna(en_ai) else "N/A",
                "english_native": str(en_native).strip() if pd.notna(en_native) else "N/A"
            })

    # Print results clearly
    for item in results:
        print(f"--- Segment (Row {item['row']}) ---")
        print(f"[Japanese]:\n{item['japanese']}\n")
        print(f"[English AI]:\n{item['english_ai']}\n")
        print(f"[English Native]:\n{item['english_native']}\n")
        print("="*40)

except Exception as e:
    print(f"Error reading file: {e}")
