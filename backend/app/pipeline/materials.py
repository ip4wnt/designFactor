"""Извлечение текста из материалов, которые пользователь добавляет на втором шаге
брифования: .txt/.md/.csv, .docx, .xlsx, .pdf и ссылки на веб-страницы.

Текстовые материалы возвращаются клиенту в рамках сессии. DOCX/XLSX/PDF загружаются
через серверный реестр таблиц в routes.py и сохраняются вместе с исходником.
Если файл или страницу
прочитать не удалось, это явно возвращается как ошибка — молча считать материал
прочитанным нельзя.
"""
from __future__ import annotations

import ipaddress
import re
import socket
import subprocess
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import urlparse

import httpx
from lxml import etree, html

MAX_CHARS = 30_000  # на один материал
MAX_DOWNLOAD = 3 * 1024 * 1024
TEXT_SUFFIXES = {".txt", ".md", ".csv", ".json", ".rtf"}
SUPPORTED = TEXT_SUFFIXES | {".docx", ".xlsx", ".pdf"}


class MaterialError(Exception):
    """Материал не удалось прочитать (сообщение показывается пользователю)."""


def _clip(text: str) -> str:
    text = re.sub(r"[ \t\u00a0]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text[:MAX_CHARS]


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1251", "utf-16"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return data.decode("utf-8", errors="ignore")


def _docx_text(path: Path) -> str:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    try:
        with zipfile.ZipFile(path) as archive:
            xml = archive.read("word/document.xml")
    except (KeyError, zipfile.BadZipFile) as exc:
        raise MaterialError("Файл .docx поврежден или не является документом Word") from exc
    root = etree.fromstring(xml)
    lines: list[str] = []
    for paragraph in root.iter("{%s}p" % ns["w"]):
        parts = [node.text or "" for node in paragraph.iter("{%s}t" % ns["w"])]
        line = "".join(parts).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def _xlsx_text(path: Path) -> str:
    from openpyxl import load_workbook

    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 — не .xlsx или поврежден
        raise MaterialError("Не удалось открыть таблицу .xlsx") from exc
    out: list[str] = []
    for sheet in workbook.worksheets[:6]:
        out.append(f"Лист «{sheet.title}»")
        for row_index, row in enumerate(sheet.iter_rows(values_only=True)):
            if row_index >= 80:
                out.append("…")
                break
            cells = ["" if value is None else str(value).strip() for value in row]
            if any(cells):
                out.append(" | ".join(cells).rstrip(" |"))
    return "\n".join(out)


def _pdf_text(path: Path) -> str:
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", "-enc", "UTF-8", str(path), "-"],
            capture_output=True, timeout=40, check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise MaterialError("Не удалось прочитать PDF") from exc
    text = result.stdout.decode("utf-8", errors="ignore")
    if result.returncode != 0 or not text.strip():
        raise MaterialError("В PDF нет текста, который можно прочитать (возможно, это скан)")
    return text


def extract_file(filename: str, data: bytes) -> tuple[str, str]:
    """Возвращает (вид материала, текст). MaterialError — если прочитать нельзя."""
    suffix = Path(filename or "").suffix.lower()
    if suffix not in SUPPORTED:
        raise MaterialError("Поддерживаются .txt, .md, .csv, .docx, .xlsx и .pdf")
    if suffix in TEXT_SUFFIXES:
        text = _decode(data)
        kind = "Текст"
    else:
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(data)
            tmp_path = Path(tmp.name)
        try:
            if suffix == ".docx":
                text, kind = _docx_text(tmp_path), "Документ"
            elif suffix == ".xlsx":
                text, kind = _xlsx_text(tmp_path), "Таблица"
            else:
                text, kind = _pdf_text(tmp_path), "PDF"
        finally:
            tmp_path.unlink(missing_ok=True)
    text = _clip(text)
    if not text:
        raise MaterialError("В файле не нашлось текста")
    return kind, text


def _check_public_host(host: str) -> None:
    """Не даём серверу ходить во внутреннюю сеть по ссылке из формы (SSRF)."""
    try:
        infos = socket.getaddrinfo(host, None)
    except socket.gaierror as exc:
        raise MaterialError("Не удалось найти такой адрес") from exc
    for info in infos:
        address = ipaddress.ip_address(info[4][0])
        if address.is_private or address.is_loopback or address.is_link_local or address.is_reserved or address.is_multicast:
            raise MaterialError("Эта ссылка недоступна для чтения: адрес не публичный")


def extract_link(url: str) -> tuple[str, str]:
    parsed = urlparse(url.strip())
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise MaterialError("Укажите адрес, начинающийся с https:// или http://")
    _check_public_host(parsed.hostname)
    headers = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36", "Accept": "text/html,text/plain,*/*;q=0.5"}
    try:
        with httpx.Client(timeout=15.0, follow_redirects=True, max_redirects=4, headers=headers) as client:
            with client.stream("GET", url.strip()) as response:
                if response.status_code in {401, 403}:
                    raise MaterialError("Страница закрыта: нужен вход или у нас нет доступа")
                if response.status_code >= 400:
                    raise MaterialError(f"Страница не открылась (ответ {response.status_code})")
                final_host = response.url.host
                if final_host and final_host != parsed.hostname:
                    _check_public_host(final_host)
                content_type = response.headers.get("content-type", "")
                body = bytearray()
                for chunk in response.iter_bytes():
                    body.extend(chunk)
                    if len(body) > MAX_DOWNLOAD:
                        break
    except MaterialError:
        raise
    except httpx.HTTPError as exc:
        raise MaterialError("Не удалось открыть страницу: " + (str(exc) or type(exc).__name__)[:120]) from exc
    if "pdf" in content_type.lower():
        return "Ссылка", extract_file("page.pdf", bytes(body))[1]
    raw = _decode(bytes(body))
    if "html" in content_type.lower() or raw.lstrip().startswith("<"):
        try:
            tree = html.fromstring(raw)
        except (etree.ParserError, ValueError) as exc:
            raise MaterialError("Страница открылась, но текст из нее извлечь не удалось") from exc
        for bad in tree.xpath("//script|//style|//noscript|//nav|//footer|//header|//form|//svg"):
            bad.getparent().remove(bad)
        title = (tree.findtext(".//title") or "").strip()
        blocks = [" ".join(node.text_content().split()) for node in tree.xpath("//h1|//h2|//h3|//p|//li")]
        text = "\n".join(block for block in blocks if len(block) > 20)
        text = (title + "\n" + text).strip()
    else:
        text = raw
    text = _clip(text)
    if len(text) < 80:
        raise MaterialError("Страница открылась, но осмысленного текста на ней нет (возможно, она собирается скриптами)")
    return "Ссылка", text
