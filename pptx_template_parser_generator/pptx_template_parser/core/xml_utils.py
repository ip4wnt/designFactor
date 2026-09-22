import xml.etree.ElementTree as ET


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def parse_xml(data):
    if isinstance(data, bytes):
        return ET.fromstring(data)

    return ET.parse(data).getroot()


def local_name(tag):
    return tag.split("}")[-1]