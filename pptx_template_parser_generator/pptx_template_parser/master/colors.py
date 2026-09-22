from pptx_template_parser.core.xml_utils import (
    parse_xml,
    NS,
    local_name,
)


def parse_theme_colors(theme_xml):
    root = parse_xml(
        theme_xml
    )

    clr_scheme = root.find(
        ".//a:clrScheme",
        NS,
    )

    if clr_scheme is None:
        return {}

    colors = {}

    for child in clr_scheme:

        name = local_name(
            child.tag
        )

        srgb = child.find(
            "a:srgbClr",
            NS,
        )

        if srgb is not None:
            value = srgb.get(
                "val"
            )

            if value:
                colors[name] = (
                    "#" + value.upper()
                )

            continue

        sysclr = child.find(
            "a:sysClr",
            NS,
        )

        if sysclr is not None:
            value = sysclr.get(
                "lastClr"
            )

            if value:
                colors[name] = (
                    "#" + value.upper()
                )

    return colors


def parse_background(master_xml):
    root = parse_xml(
        master_xml
    )

    bg = root.find(
        ".//p:bg",
        NS,
    )

    if bg is None:
        return {}

    bg_pr = bg.find(
        "p:bgPr",
        NS,
    )

    if bg_pr is None:
        return {}

    result = {}

    # ========================================================
    # COLOR
    # ========================================================

    solid_fill = bg_pr.find(
        "a:solidFill",
        NS,
    )

    if solid_fill is not None:

        srgb = solid_fill.find(
            "a:srgbClr",
            NS,
        )

        if srgb is not None:
            value = srgb.get(
                "val"
            )

            if value:
                result["color"] = (
                    "#" + value.upper()
                )

        if "color" not in result:

            scheme = solid_fill.find(
                "a:schemeClr",
                NS,
            )

            if scheme is not None:
                value = scheme.get(
                    "val"
                )

                if value:
                    result["scheme"] = value

    # ========================================================
    # BACKGROUND IMAGE
    # ========================================================

    blip = bg_pr.find(
        "a:blipFill/a:blip",
        NS,
    )

    if blip is not None:
        result["rId"] = blip.get(
            "{%s}embed" % NS["r"]
        )

    return result