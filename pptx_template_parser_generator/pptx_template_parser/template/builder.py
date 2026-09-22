from pptx import Presentation

from pptx_template_parser.core.zip_reader import (
    ZipReader,
)

from pptx_template_parser.master.colors import (
    parse_theme_colors,
    parse_background,
)

from pptx_template_parser.master.fonts import (
    parse_theme_fonts,
)
from pptx_template_parser.media.relationships import parse_relationships

from pptx_template_parser.layout.variants import (
    parse_variants,
    parse_slide_variants,
    has_default_layouts,
)


def find_theme(zip_reader):
    theme_files = zip_reader.list(
        "ppt/theme/"
    )

    for filename in theme_files:

        if filename.endswith(".xml"):
            return filename

    return None


def build_raw_template(
    pptx_path,
):
    presentation = Presentation(
        pptx_path
    )

    # --------------------------------------------------------
    # SIZE
    # --------------------------------------------------------

    slide_size = {
        "width": round(
            presentation.slide_width / 360000,
            2,
        ),

        "height": round(
            presentation.slide_height / 360000,
            2,
        ),

        "unit": "cm",
    }

    # --------------------------------------------------------
    # VARIANTS
    # --------------------------------------------------------

    source_mode = (
        "slides" if len(presentation.slides) and has_default_layouts(presentation)
        else "layouts"
    )
    variants = (parse_slide_variants(presentation) if source_mode == "slides"
                else parse_variants(presentation))

    # --------------------------------------------------------
    # THEME + MASTERS + MEDIA
    # --------------------------------------------------------

    with ZipReader(
        pptx_path
    ) as reader:

        # ----------------------------------------------------
        # THEME
        # ----------------------------------------------------

        theme_file = find_theme(
            reader
        )

        if theme_file:

            theme_xml = reader.read(
                theme_file
            )

            colors = parse_theme_colors(
                theme_xml
            )

            fonts = parse_theme_fonts(
                theme_xml
            )

        else:

            colors = {}
            fonts = {}

        # ----------------------------------------------------
        # MASTER BACKGROUNDS
        # ----------------------------------------------------

        master_cache = {}

        for variant in variants:

            master_file = variant.get(
                "master_file"
            )

            if master_file and master_file not in master_cache:
                master_cache[master_file] = (
                    parse_background(reader.read(master_file))
                    if reader.exists(master_file) else {}
                )

            background = dict(master_cache.get(master_file, {}))

            # A layout background overrides the master background.
            layout_file = variant.get("layout_file")
            layout_background = {}
            if layout_file and reader.exists(layout_file):
                layout_background = parse_background(reader.read(layout_file))
                if layout_background:
                    background = layout_background

            slide_file = variant.get("slide_file")
            slide_background = {}
            if slide_file and reader.exists(slide_file):
                slide_background = parse_background(reader.read(slide_file))
                if slide_background:
                    background = slide_background

            background_file = (slide_file if slide_background.get("rId") else
                               layout_file if layout_background.get("rId") else master_file)
            if background.get("rId") and background_file:
                relation = parse_relationships(reader, background_file).get(
                    background["rId"], {}
                )
                if relation.get("target"):
                    background["target"] = relation["target"]

            # ------------------------------------------------
            # RESOLVE THEME COLOR
            # ------------------------------------------------

            if (
                "color" not in background
                and "scheme" in background
            ):
                scheme = background["scheme"]

                color = colors.get(
                    scheme
                )

                if color:
                    background["color"] = color

            # Сохраняем результат
            variant["background_raw"] = (
                background
            )

        # ----------------------------------------------------
        # MEDIA
        # ----------------------------------------------------

        assets = {
            "images": []
        }

        media_files = reader.list(
            "ppt/media/"
        )

        for media_file in media_files:

            assets["images"].append(
                media_file
            )

    # --------------------------------------------------------
    # RESULT
    # --------------------------------------------------------

    return {
        "layout_id": (
            presentation.slide_layouts[0].name
            if presentation.slide_layouts
            else "template"
        ),

        "type": "presentation_template",
        "source_mode": source_mode,

        "slide_size": slide_size,

        "colors": colors,

        "assets": assets,

        "fonts": fonts,

        "variants": variants,
    }
