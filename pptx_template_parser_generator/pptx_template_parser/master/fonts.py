from pptx_template_parser.core.xml_utils import parse_xml, NS


def parse_theme_fonts(theme_xml):
    root = parse_xml(theme_xml)

    result = {
        "major": {},
        "minor": {},
    }

    major = root.find(".//a:majorFont", NS)
    minor = root.find(".//a:minorFont", NS)

    if major is not None:
        latin = major.find("a:latin", NS)

        if latin is not None:
            result["major"]["latin"] = latin.attrib.get(
                "typeface",
                ""
            )

    if minor is not None:
        latin = minor.find("a:latin", NS)

        if latin is not None:
            result["minor"]["latin"] = latin.attrib.get(
                "typeface",
                ""
            )

    return result