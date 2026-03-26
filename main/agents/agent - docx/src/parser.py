import json
import os

def parse_input(file_path="input.json"):
    base_dir = os.path.dirname(__file__)
    full_path = os.path.join(base_dir, file_path)

    with open(full_path, "r") as f:
        data = json.load(f)

    return data.get("title"), data.get("sections")