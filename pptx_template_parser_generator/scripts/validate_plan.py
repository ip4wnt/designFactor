import argparse
import json
from pathlib import Path

from pptx_template_parser import build_template, validate_generation_plan


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Проверить JSON-план модели по generation_contract шаблона"
    )
    parser.add_argument("source", help="Путь к исходному PPTX")
    parser.add_argument("plan", help="Путь к JSON-плану со slides")
    parser.add_argument("--check-images", action="store_true",
                        help="Проверить существование файлов изображений")
    args = parser.parse_args(argv)
    source, plan_path = Path(args.source), Path(args.plan)
    if not source.is_file():
        parser.error(f"Шаблон не найден: {source}")
    if not plan_path.is_file():
        parser.error(f"План не найден: {plan_path}")
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    errors = validate_generation_plan(
        build_template(source), plan, check_image_files=args.check_images,
    )
    if errors:
        print(json.dumps({"valid": False, "errors": errors}, ensure_ascii=False, indent=2))
        return 1
    print(json.dumps({"valid": True, "errors": []}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
