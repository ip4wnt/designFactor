"""Generate the three layout variants required by the specification."""

import argparse
import json
import os
from pathlib import Path

from pptx_template_parser.model import ModelError, OpenAICompatibleModel
from scripts.generate_ai import run


PROJECT_ROOT = Path(__file__).resolve().parents[1]
VARIANTS = (
    ("variant_a", "balanced", "01_balanced"),
    ("variant_b", "compact", "02_compact"),
    ("variant_c", "visual", "03_visual"),
)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Создать три варианта презентации по одному брифу"
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("--brief", required=True)
    parser.add_argument("--purpose", default="business")
    parser.add_argument("--slides", type=int, default=10)
    parser.add_argument("--output-dir", type=Path,
                        default=PROJECT_ROOT / "output" / "ai_suite")
    parser.add_argument("--base-url", default=os.getenv("PPTX_LLM_BASE_URL"))
    parser.add_argument("--model", default=os.getenv("PPTX_LLM_MODEL"))
    parser.add_argument("--api-key", default=os.getenv("PPTX_LLM_API_KEY"))
    parser.add_argument("--image", type=Path)
    parser.add_argument("--mock", action="store_true")
    args = parser.parse_args(argv)
    if not args.source.is_file():
        parser.error(f"Шаблон не найден: {args.source}")
    if not 10 <= args.slides <= 15:
        parser.error("Укажи --slides от 10 до 15")
    model = None if args.mock else OpenAICompatibleModel(
        base_url=args.base_url, model=args.model, api_key=args.api_key
    )
    results = []
    try:
        for name, strategy, mock_profile in VARIANTS:
            path, audit = run(
                args.source,
                args.brief,
                f"{args.purpose}; layout_strategy={strategy}",
                args.slides,
                args.output_dir / name,
                model=model,
                mock=args.mock,
                image=args.image,
                mock_profile=mock_profile,
            )
            results.append({"variant": name, "presentation": str(path.resolve()),
                            "audit": audit})
    except (OSError, ValueError, ModelError) as exc:
        parser.error(str(exc))
    report = {
        "brief": args.brief,
        "slides_per_deck": args.slides,
        "passed": sum(item["audit"]["passed"] for item in results),
        "total": len(results),
        "results": results,
    }
    args.output_dir.mkdir(parents=True, exist_ok=True)
    (args.output_dir / "suite_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["passed"] == report["total"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
