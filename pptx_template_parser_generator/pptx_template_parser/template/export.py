import json
from pathlib import Path


def save_template_json(template, path):
    """Save an independently reusable description of the source layouts."""
    output = Path(path)
    with output.open("x", encoding="utf-8") as stream:
        json.dump(template, stream, ensure_ascii=False, indent=2)
        stream.write("\n")
    return output
