from collections import Counter


def placeholder_to_slot(value):
    if not value:
        return None

    value = value.lower()

    if "title" in value:
        return "title"

    if "footer" in value:
        return "footer"

    if "subtitle" in value:
        return "subtitle"

    if "body" in value:
        return "text"

    if "object" in value:
        return "text"

    return None


def _most_common(values):
    values = [
        value
        for value in values
        if value not in (None, "")
    ]

    if not values:
        return None

    return Counter(values).most_common(1)[0][0]


def layout_type(text_slots, image_slots):
    """Describe editable structure, not an inferred business purpose."""
    has_title = "title" in text_slots
    has_body = any(name != "title" for name in text_slots)
    has_image = bool(image_slots)
    if has_title and has_body and has_image:
        return "title_text_image"
    if has_title and has_body:
        return "title_text"
    if has_title and has_image:
        return "title_image"
    if has_body and has_image:
        return "text_image"
    if has_title:
        return "title"
    if has_body:
        return "text"
    if has_image:
        return "image"
    return "decorative"


def editable_regions(slots):
    return {
        name: {key: slot[key] for key in ("x", "y", "w", "h")}
        for name, slot in slots.items()
        if slot.get("placeholder_idx") is not None
    }


def normalize_variant(
    raw_variant,
    slide_width,
    slide_height,
):
    slots = {}
    text_index = 0

    # --------------------------------------------------
    # TEXT
    # --------------------------------------------------

    for shape in raw_variant["slots_raw"]:

        explicit_slot = placeholder_to_slot(
            shape.get("placeholder_type")
        )

        if explicit_slot:
            slot = explicit_slot

        else:
            text_index += 1

            if text_index == 1:
                slot = "text"
            else:
                slot = f"text_{text_index}"

        if slot in slots:
            base = slot
            index = 2

            while f"{base}_{index}" in slots:
                index += 1

            slot = f"{base}_{index}"

        slots[slot] = {
            "placeholder_idx": shape.get("placeholder_idx"),
            "shape_index": shape.get("shape_index"),
            "x": shape["x"],
            "y": shape["y"],
            "w": shape["w"],
            "h": shape["h"],

            "horizontal_align": _normalize_horizontal_align(
                shape.get("horizontal_align")
            ),

            "vertical_align": _normalize_vertical_align(
                shape.get("vertical_align")
            ),

            "font": shape.get("font"),
        }


    # --------------------------------------------------
    # IMAGES
    # --------------------------------------------------

    normalized_images = normalize_images(
        raw_variant["images_raw"],
        slide_width,
        slide_height,
    )

    from_slides = raw_variant.get("slide_index") is not None
    source_images = raw_variant.get("image_placeholders_raw", [])
    if from_slides:
        source_images = source_images + [
            image for image in raw_variant["images_raw"]
            if not _is_full_slide(image, slide_width, slide_height)
        ]
    image_slots = {
        ("image" if index == 1 else f"image_{index}"): {
            "placeholder_idx": shape.get("placeholder_idx"),
            "shape_index": shape.get("shape_index"),
            "x": shape["x"],
            "y": shape["y"],
            "w": shape["w"],
            "h": shape["h"],
        }
        for index, shape in enumerate(source_images, start=1)
    }

    background = normalized_images["background"]

    master_background = raw_variant.get("background_raw") or {}

    if background is None and master_background and not from_slides:
        background = {}
        if master_background.get("target"):
            background.update({
                "target": master_background["target"],
                "x": 0,
                "y": 0,
                "w": slide_width,
                "h": slide_height,
            })

    if master_background.get("color"):
        if background is None:
            background = {}
        background["color"] = master_background["color"]

    text_regions = (editable_regions(slots) if raw_variant.get("slide_index") is None
                    else {name: {key: spec[key] for key in ("x", "y", "w", "h")}
                          for name, spec in slots.items()})
    image_regions = (editable_regions(image_slots) if not from_slides
                     else {name: {key: spec[key] for key in ("x", "y", "w", "h")}
                           for name, spec in image_slots.items()})

    return {
        "id": raw_variant["name"],
        "type": layout_type(text_regions, image_regions),
        "layout_file": raw_variant.get("layout_file"),
        "master_index": raw_variant.get("master_index"),
        "slide_index": raw_variant.get("slide_index"),
        "slots": slots,
        "image_slots": image_slots,
        "background": background,
        "images": [] if from_slides else normalized_images["images"],
        "constraints": {
            "slide_bounds_cm": {
                "x": 0, "y": 0, "w": slide_width, "h": slide_height,
            },
            "text_regions_cm": text_regions,
            "image_regions_cm": image_regions,
        },
    }


def _normalize_vertical_align(value):
    if not value:
        return None

    value = str(value)

    if "bottom" in value:
        return "bottom"

    if "top" in value:
        return "top"

    if "middle" in value:
        return "middle"

    if "center" in value:
        return "middle"

    return value


def normalize_fonts(raw_variants):
    title_fonts = []
    subtitle_fonts = []

    for variant in raw_variants:

        for shape in variant["slots_raw"]:

            slot = placeholder_to_slot(
                shape.get("placeholder_type")
            )

            if slot == "title":
                title_fonts.append(shape["font"])

            elif slot == "subtitle":
                subtitle_fonts.append(shape["font"])

    def build(fonts):

        family = _most_common(
            [font.get("family") for font in fonts]
        )

        size = _most_common(
            [font.get("size") for font in fonts]
        )

        weight = _most_common(
            [font.get("weight") for font in fonts]
        )

        return {
            "family": family or "",
            "size": size,
            "weight": weight or "regular",
        }

    return {
        "title": build(title_fonts),
        "subtitle": build(subtitle_fonts),
    }


def normalize_margins(raw_variants):
    margins = []

    for variant in raw_variants:

        for shape in variant["slots_raw"]:

            value = shape.get("margins")

            if value:
                margins.append(
                    tuple(value.items())
                )

    if not margins:
        return {
            "left": None,
            "right": None,
            "top": None,
            "bottom": None,
        }

    most_common = Counter(margins).most_common(1)[0][0]

    return dict(most_common)


def normalize_template(raw_template):
    raw_variants = raw_template["variants"]

    slide_width = raw_template["slide_size"]["width"]
    slide_height = raw_template["slide_size"]["height"]

    variants = [
        normalize_variant(
            variant,
            slide_width,
            slide_height,
        )
        for variant in raw_variants
    ]

    return {
        "layout_id": raw_template["layout_id"],
        "type": "presentation_template",
        "source_mode": raw_template.get("source_mode", "layouts"),
        "slide_size": raw_template["slide_size"],
        "colors": raw_template.get("colors", {}),

        "text_box_defaults": {
            "margins": normalize_margins(
                raw_variants
            ),
        },

        "fonts": normalize_fonts(
            raw_variants
        ),

        "variants": variants,

        "allowed_content": raw_template.get(
            "allowed_content",
            [],
        ),

        "disallowed_content": raw_template.get(
            "disallowed_content",
            [],
        ),
    }


def normalize_images(
    raw_images,
    slide_width,
    slide_height,
):
    result = []
    seen = set()
    background = None

    for image in raw_images:

        target = image.get("target")

        if not target:
            continue

        # Один файл может намеренно стоять в нескольких местах макета.
        identity = (target, image["x"], image["y"], image["w"], image["h"])
        if identity in seen:
            continue

        seen.add(identity)

        normalized = {
            "name": image.get("name"),
            "rId": image.get("rId"),
            "target": target,
            "x": image["x"],
            "y": image["y"],
            "w": image["w"],
            "h": image["h"],
        }
        if image.get("shape_index") is not None:
            normalized["shape_index"] = image["shape_index"]

        # --------------------------------------------------
        # BACKGROUND
        # --------------------------------------------------

        if (
            background is None
            and _is_full_slide(
                image,
                slide_width,
                slide_height,
            )
        ):
            background = normalized
            continue

        result.append(normalized)

    return {
        "background": background,
        "images": result,
    }


def _is_full_slide(
    image,
    slide_width,
    slide_height,
    tolerance=0.05,
):
    def close(a, b):
        return abs(a - b) <= tolerance

    return (
        close(image["x"], 0)
        and close(image["y"], 0)
        and close(image["w"], slide_width)
        and close(image["h"], slide_height)
    )


def _normalize_horizontal_align(value):
    if not value:
        return None

    value = str(value).lower()

    if "center" in value:
        return "center"

    if "right" in value:
        return "right"

    if "justify" in value:
        return "justify"

    if "left" in value:
        return "left"

    return value

