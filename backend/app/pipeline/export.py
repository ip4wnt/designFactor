"""Экспорт готового .pptx в .pdf/.html.

.pptx уже полностью готов на выходе Компоновщика (app/pipeline/template_engine.py) —
здесь только конвертация форматов, без какой-либо генерации контента.

* pdf — headless LibreOffice (общая очередь app/pipeline/soffice.py);
* html — автономная HTML-страница-слайдшоу: слайды рендерятся через PDF в PNG
  и встраиваются как data-URI. Штатный HTML-фильтр LibreOffice на шаблонах
  с крупными фото падает («Text node too long») и даёт неаккуратную вёрстку,
  поэтому не используется.
"""
from __future__ import annotations

import asyncio
import base64
import html
import io
import shutil
from pathlib import Path

from pdf2image import convert_from_path

from app.pipeline import soffice

_SUPPORTED_FORMATS = {"pptx", "pdf", "html"}


class ExportError(RuntimeError):
    pass


async def export_presentation(pptx_path: str | Path, output_dir: str | Path, fmt: str) -> Path:
    if fmt not in _SUPPORTED_FORMATS:
        raise ValueError(f"Неподдерживаемый формат экспорта: {fmt}")

    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if fmt == "pptx":
        destination = output_dir / pptx_path.name
        if destination.resolve() != pptx_path.resolve():
            shutil.copyfile(pptx_path, destination)
        return destination

    if fmt == "html":
        return await _export_html(pptx_path, output_dir)
    return await _convert_with_libreoffice(pptx_path, output_dir, fmt)


_HTML_DPI = 120

_HTML_PAGE = """<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
  html,body{{margin:0;height:100%;background:#0b0f19;color:#e5e7eb;font:14px/1.4 -apple-system,Segoe UI,Roboto,sans-serif}}
  .stage{{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;padding:24px 24px 56px;box-sizing:border-box}}
  .stage img{{max-width:100%;max-height:100%;box-shadow:0 20px 60px rgba(0,0,0,.6);border-radius:6px;display:none}}
  .stage img.on{{display:block}}
  .bar{{position:fixed;left:0;right:0;bottom:0;height:44px;display:flex;align-items:center;justify-content:center;gap:16px;background:rgba(11,15,25,.9)}}
  button{{background:#1f2937;color:#e5e7eb;border:0;border-radius:6px;padding:6px 14px;cursor:pointer;font:inherit}}
  button:hover{{background:#374151}}
  .grid{{display:none;padding:24px;gap:16px;grid-template-columns:repeat(auto-fill,minmax(260px,1fr))}}
  .grid img{{width:100%;border-radius:4px;cursor:pointer}}
  body.overview .stage{{display:none}} body.overview .grid{{display:grid}}
</style></head><body>
<div class="stage">{images}</div>
<div class="grid">{thumbs}</div>
<div class="bar"><button id="prev">←</button><span id="counter"></span><button id="next">→</button><button id="grid">Все слайды</button></div>
<script>
(function(){{
  var imgs=[].slice.call(document.querySelectorAll('.stage img')),i=0,n=imgs.length;
  function show(k){{i=(k+n)%n;imgs.forEach(function(im,j){{im.classList.toggle('on',j===i)}});document.getElementById('counter').textContent=(i+1)+' / '+n;document.body.classList.remove('overview')}}
  document.getElementById('prev').onclick=function(){{show(i-1)}};
  document.getElementById('next').onclick=function(){{show(i+1)}};
  document.getElementById('grid').onclick=function(){{document.body.classList.toggle('overview')}};
  document.querySelector('.stage').onclick=function(){{show(i+1)}};
  [].slice.call(document.querySelectorAll('.grid img')).forEach(function(im,j){{im.onclick=function(){{show(j)}}}});
  document.addEventListener('keydown',function(e){{if(e.key==='ArrowRight'||e.key===' '||e.key==='PageDown')show(i+1);else if(e.key==='ArrowLeft'||e.key==='PageUp')show(i-1);else if(e.key==='Escape')document.body.classList.toggle('overview')}});
  show(0);
}})();
</script></body></html>
"""


async def _export_html(pptx_path: Path, output_dir: Path) -> Path:
    destination = output_dir / (pptx_path.stem + ".html")
    pdf_dir = output_dir / "_html_tmp"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    try:
        pdf_path = await soffice.convert(pptx_path, pdf_dir, "pdf")
    except soffice.SofficeError as exc:
        raise ExportError(str(exc)) from exc
    images = await asyncio.to_thread(convert_from_path, str(pdf_path), dpi=_HTML_DPI)
    shutil.rmtree(pdf_dir, ignore_errors=True)
    if not images:
        raise ExportError("Не удалось отрендерить слайды для HTML-экспорта")

    def encode(image) -> str:
        buffer = io.BytesIO()
        image.convert("RGB").save(buffer, "JPEG", quality=88, optimize=True)
        return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")

    sources = await asyncio.to_thread(lambda: [encode(image) for image in images])
    imgs = "".join(f'<img src="{src}" alt="Слайд {idx}">' for idx, src in enumerate(sources, start=1))
    thumbs = "".join(f'<img src="{src}" alt="Слайд {idx}" loading="lazy">' for idx, src in enumerate(sources, start=1))
    destination.write_text(
        _HTML_PAGE.format(title=html.escape(pptx_path.stem), images=imgs, thumbs=thumbs),
        encoding="utf-8",
    )
    return destination


async def _convert_with_libreoffice(pptx_path: Path, output_dir: Path, fmt: str) -> Path:
    """Конвертация через общий сериализованный запуск soffice (см. app/pipeline/soffice.py)."""
    try:
        return await soffice.convert(pptx_path, output_dir, fmt)
    except soffice.SofficeError as exc:
        raise ExportError(str(exc)) from exc
