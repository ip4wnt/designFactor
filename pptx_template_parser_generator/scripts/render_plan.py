import argparse
import json
from pathlib import Path

from pptx_template_parser import build_template, validate_generation_plan
from scripts.generate import create_presentation


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Собрать PPTX из шаблона и JSON-плана модели"
    )
    parser.add_argument("source", help="Путь к исходному PPTX-шаблону")
    parser.add_argument("plan", help="Путь к JSON-плану со slides")
    parser.add_argument("--output", help="Путь к итоговому PPTX")
    parser.add_argument("--check-images", action="store_true",
                        help="Проверить все пути изображений до рендеринга")
    args = parser.parse_args(argv)

    source, plan_path = Path(args.source), Path(args.plan)
    if not source.is_file():
        parser.error(f"Шаблон не найден: {source}")
    if not plan_path.is_file():
        parser.error(f"План не найден: {plan_path}")
    output = (Path(args.output) if args.output else
              DEFAULT_OUTPUT_DIR / f"{plan_path.stem}.pptx")
    if output.suffix.lower() != ".pptx":
        parser.error("Выходной файл должен иметь расширение .pptx")
    if output.exists():
        parser.error(f"PPTX уже существует: {output}")

    try:
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        parser.error(f"Не удалось прочитать JSON-план: {exc}")
    template = build_template(source)
    errors = validate_generation_plan(
        template, plan, check_image_files=args.check_images,
    )
    if errors:
        print(json.dumps({"valid": False, "errors": errors},
                         ensure_ascii=False, indent=2))
        return 1

    slides = plan.get("slides") if isinstance(plan, dict) else plan
    output.parent.mkdir(parents=True, exist_ok=True)
    create_presentation(
        template, slides, output, None, source_file=source,
    )
    print(json.dumps({
        "valid": True,
        "slides": len(slides),
        "output": str(output.resolve()),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
