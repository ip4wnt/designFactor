from .builder import build_raw_template
from .normalizer import normalize_template


def build_template(pptx_path):
    raw = build_raw_template(
        pptx_path
    )

    return normalize_template(
        raw
    )


__all__ = [
    "build_template",
]