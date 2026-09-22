import argparse
from pathlib import Path

from pptx_template_parser import build_template
from pptx_template_parser.template.export import save_template_json


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "output"


def main(argv=None):
    parser = argparse.ArgumentParser(description="Сохранить описание образцов PPTX в JSON")
    parser.add_argument("source", help="Путь к исходному PPTX")
    parser.add_argument("--output", help="Путь для JSON (по умолчанию в output)")
    args = parser.parse_args(argv)

    source = Path(args.source)
    if not source.is_file():
        parser.error(f"Шаблон не найден: {source}")
    output = (Path(args.output) if args.output else
              DEFAULT_OUTPUT_DIR / f"{source.stem}_template.json")
    output.parent.mkdir(parents=True, exist_ok=True)
    template = build_template(source)
    if output.exists():
        parser.error(f"JSON уже существует: {output}")
    save_template_json(template, output)
    print(f"Сохранено {len(template['variants'])} макетов: {output}")
    return output


if __name__ == "__main__":
    main()
