from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from pptx_template_parser.core.units import emu_to_cm
from pptx_template_parser.core.xml_utils import NS


REL_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/relationships"
)


def shape_position(shape):
    return {
        "x": emu_to_cm(shape.left),
        "y": emu_to_cm(shape.top),
        "w": emu_to_cm(shape.width),
        "h": emu_to_cm(shape.height),
    }


def placeholder_type(shape):
    if not shape.is_placeholder:
        return None

    try:
        return str(shape.placeholder_format.type)
    except Exception:
        return None


def get_text(shape):
    if not shape.has_text_frame:
        return ""

    return shape.text


def _first_text_run(shape):
    if not shape.has_text_frame:
        return None

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            return run

    return None


def _xml_font_info(shape):
    result = {
        "family": None,
        "size": None,
        "weight": "regular",
    }

    element = shape._element

    candidates = [
        element.find(".//a:rPr", NS),
        element.find(".//a:defRPr", NS),
        element.find(".//a:endParaRPr", NS),
    ]

    rpr = next(
        (item for item in candidates if item is not None),
        None,
    )

    if rpr is None:
        return result

    size = rpr.attrib.get("sz")

    if size:
        try:
            result["size"] = round(int(size) / 100, 2)
        except ValueError:
            pass

    latin = rpr.find("a:latin", NS)

    if latin is not None:
        typeface = latin.attrib.get("typeface")

        if typeface:
            result["family"] = typeface

    if rpr.attrib.get("b") == "1":
        result["weight"] = "bold"

    return result


def get_font_info(shape):
    result = _xml_font_info(shape)

    run = _first_text_run(shape)

    if run is not None:
        font = run.font

        if font.name:
            result["family"] = font.name

        if font.size:
            result["size"] = round(font.size.pt, 2)

        if font.bold:
            result["weight"] = "bold"

    return result


def get_vertical_align(shape):
    if not shape.has_text_frame:
        return None

    try:
        value = shape.text_frame.vertical_anchor

        if value is None:
            return None

        return str(value).split(".")[-1].lower()

    except Exception:
        return None



def get_horizontal_align(shape):
    if not shape.has_text_frame:
        return None

    element = shape._element

    # 1. Явное выравнивание конкретного абзаца
    paragraphs = element.findall(
        ".//a:p",
        NS,
    )

    for paragraph in paragraphs:
        p_pr = paragraph.find(
            "a:pPr",
            NS,
        )

        if p_pr is not None:
            algn = p_pr.get("algn")

            if algn:
                return _normalize_xml_horizontal_align(algn)

    # 2. Выравнивание по умолчанию для уровня текста
    lvl1p_pr = element.find(
        ".//a:lvl1pPr",
        NS,
    )

    if lvl1p_pr is not None:
        algn = lvl1p_pr.get("algn")

        if algn:
            return _normalize_xml_horizontal_align(algn)

    # 3. Fallback через python-pptx
    for paragraph in shape.text_frame.paragraphs:
        if paragraph.alignment is not None:
            value = str(
                paragraph.alignment
            ).split(".")[-1].lower()

            return value

    return None


def _normalize_xml_horizontal_align(value):
    mapping = {
        "l": "left",
        "ctr": "center",
        "r": "right",
        "just": "justify",
        "dist": "distributed",
        "thaiDist": "thai_distributed",
    }

    return mapping.get(value, value)





def get_margins(shape):
    if not shape.has_text_frame:
        return None

    tf = shape.text_frame

    return {
        "left": emu_to_cm(tf.margin_left),
        "right": emu_to_cm(tf.margin_right),
        "top": emu_to_cm(tf.margin_top),
        "bottom": emu_to_cm(tf.margin_bottom),
    }



# ============================================================
# IMAGE
# ============================================================

def get_picture_rid(shape):
    """
    Получает r:embed из p:blip.
    """

    try:
        blip = shape._element.find(".//a:blip", NS)

        if blip is None:
            return None

        return blip.attrib.get(
            f"{{{REL_NS}}}embed"
        )

    except Exception:
        return None


def get_picture_target(shape):
    """
    Получает реальный путь файла внутри PPTX.
    """

    rid = get_picture_rid(shape)

    if not rid:
        return None

    try:
        rel = shape.part.rels.get(rid)

        if rel is None:
            return None

        if not hasattr(rel, "target_part"):
            return None

        target_part = rel.target_part

        return str(
            target_part.partname
        ).lstrip("/")

    except Exception:
        return None


def get_picture_metadata(shape):
    """
    Достаёт всё, что PowerPoint хранит у cNvPr:

        name
        descr
        title
    """

    result = {
        "name": shape.name,
        "description": None,
        "title": None,
    }

    try:
        c_nv_pr = shape._element.find(
            ".//p:cNvPr",
            NS,
        )

        if c_nv_pr is None:
            return result

        result["description"] = (
            c_nv_pr.attrib.get("descr")
        )

        result["title"] = (
            c_nv_pr.attrib.get("title")
        )

    except Exception:
        pass

    return result


def is_full_slide_image(
    image,
    slide_width,
    slide_height,
    tolerance=0.05,
):
    """
    Картинка считается фоном, только если она
    практически полностью совпадает со слайдом.

    Никаких фиксированных координат.
    """

    def close(a, b):
        return abs(a - b) <= tolerance

    return (
        close(image["x"], 0)
        and close(image["y"], 0)
        and close(image["w"], slide_width)
        and close(image["h"], slide_height)
    )


def parse_layout_shapes(layout):
    text_shapes = []
    image_shapes = []
    image_placeholders = []

    for shape_index, shape in enumerate(layout.shapes):

        position = shape_position(shape)

        if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            image_placeholders.append({
                "name": shape.name,
                "shape_index": shape_index,
                "placeholder_idx": shape.placeholder_format.idx,
                **position,
            })
            continue

        # --------------------------------------------------
        # IMAGE
        # --------------------------------------------------

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:

            metadata = get_picture_metadata(shape)

            image_shapes.append({
                **metadata,
                "shape_index": shape_index,

                "rId": get_picture_rid(shape),
                "target": get_picture_target(shape),

                **position,
            })

            continue

        # --------------------------------------------------
        # TEXT
        # --------------------------------------------------

        if shape.has_text_frame:
            # Date and slide-number placeholders are managed by PowerPoint.
            # Footer is intentionally retained: some templates use it as an
            # editable text field at the top of the slide.
            if shape.is_placeholder and shape.placeholder_format.type in {
                PP_PLACEHOLDER.DATE,
                PP_PLACEHOLDER.SLIDE_NUMBER,
                PP_PLACEHOLDER.HEADER,
            }:
                continue

            text_shapes.append({
                "name": shape.name,
                "shape_index": shape_index,
                "placeholder_idx": (
                    shape.placeholder_format.idx
                    if shape.is_placeholder else None
                ),

                "placeholder_type": placeholder_type(
                    shape
                ),

                "text": get_text(shape),

                "font": get_font_info(shape),

                "horizontal_align": get_horizontal_align(
                    shape
                ),

                "vertical_align": get_vertical_align(
                    shape
                ),

                "margins": get_margins(shape),

                **position,
            })

    return {
        "text_shapes": text_shapes,
        "images": image_shapes,
        "image_placeholders": image_placeholders,
    }
