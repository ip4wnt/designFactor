import posixpath
import xml.etree.ElementTree as ET

from pptx_template_parser.core.xml_utils import NS


def relationships_path(xml_path):
    directory = posixpath.dirname(xml_path)
    filename = posixpath.basename(xml_path)

    return posixpath.join(
        directory,
        "_rels",
        filename + ".rels",
    )


def parse_relationships(
    zip_reader,
    xml_path,
):
    rels_path = relationships_path(
        xml_path
    )

    if not zip_reader.exists(
        rels_path
    ):
        return {}

    root = ET.fromstring(
        zip_reader.read(rels_path)
    )

    result = {}

    for rel in root.findall(
        "pr:Relationship",
        NS
    ):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")

        if not rid or not target:
            continue

        base = posixpath.dirname(
            xml_path
        )

        target_path = posixpath.normpath(
            posixpath.join(
                base,
                target
            )
        )

        while target_path.startswith("../"):
            target_path = target_path[3:]

        result[rid] = {
            "target": target_path,
            "type": rel_type,
        }

    return result