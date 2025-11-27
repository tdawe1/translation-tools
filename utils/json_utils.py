import json
import re

def extract_json_array(s: str, expected_len: int):
    s = re.sub(r"^```(?:json)?|```$", "", s.strip(), flags=re.M)
    dec = json.JSONDecoder()
    in_str = esc = False
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if esc:
            esc = False
        elif ch == '\\' and in_str:
            esc = True
        elif ch == '"' and in_str:
            in_str = not in_str
        elif not in_str and ch == '[':
            try:
                obj, end = dec.raw_decode(s, i)
            except json.JSONDecodeError:
                i += 1
                continue
            if isinstance(obj, list) and (expected_len == 0 or len(obj) >= expected_len):
                return obj[:expected_len] if expected_len else obj
            i = end
            continue
        i += 1
    return None
