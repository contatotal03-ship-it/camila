import json
import os

try:
    v_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.json")
    with open(v_path, "r", encoding="utf-8") as f:
        print(json.load(f).get("version", "1.1.0"))
except Exception:
    print("1.1.0")
