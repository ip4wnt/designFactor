"""Deterministic spatial graph and conservative composition patterns."""

import math


def _bounds(slot):
    return {key: float(slot.get(key, 0) or 0) for key in ("x", "y", "w", "h")}


def _overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def _ratio(value, denominator):
    return value / max(denominator, 0.001)


def build_spatial_graph(variant, slide_width, slide_height):
    """Return geometry-derived nodes and edges without semantic assumptions."""
    nodes = []
    for kind, collection in (("text", variant.get("slots", {})),
                             ("image", variant.get("image_slots", {}))):
        for name, slot in collection.items():
            bounds = _bounds(slot)
            nodes.append({
                "id": f"{kind}:{name}",
                "kind": kind,
                "slot": name,
                "bounds_cm": bounds,
                "center_cm": {
                    "x": round(bounds["x"] + bounds["w"] / 2, 3),
                    "y": round(bounds["y"] + bounds["h"] / 2, 3),
                },
            })

    edges = []
    x_tolerance = max(0.15, slide_width * 0.015)
    y_tolerance = max(0.15, slide_height * 0.025)
    near_distance = math.hypot(slide_width, slide_height) * 0.08
    for index, first in enumerate(nodes):
        a = first["bounds_cm"]
        ac = first["center_cm"]
        for second in nodes[index + 1:]:
            b = second["bounds_cm"]
            bc = second["center_cm"]
            relations = []
            if abs(a["x"] - b["x"]) <= x_tolerance:
                relations.append("aligned_left")
            if abs((a["x"] + a["w"]) - (b["x"] + b["w"])) <= x_tolerance:
                relations.append("aligned_right")
            if abs(ac["x"] - bc["x"]) <= x_tolerance:
                relations.append("aligned_center_x")
            if abs(a["y"] - b["y"]) <= y_tolerance:
                relations.append("aligned_top")
            if abs(ac["y"] - bc["y"]) <= y_tolerance:
                relations.append("aligned_center_y")

            horizontal_overlap = _overlap(a["x"], a["x"] + a["w"], b["x"], b["x"] + b["w"])
            vertical_overlap = _overlap(a["y"], a["y"] + a["h"], b["y"], b["y"] + b["h"])
            if horizontal_overlap and vertical_overlap:
                relations.append("overlaps")
            if _ratio(horizontal_overlap, min(a["w"], b["w"])) >= 0.4:
                relations.append("same_column")
                relations.append("above" if ac["y"] < bc["y"] else "below")
            if _ratio(vertical_overlap, min(a["h"], b["h"])) >= 0.4:
                relations.append("same_row")
                relations.append("left_of" if ac["x"] < bc["x"] else "right_of")

            dx = max(a["x"] - (b["x"] + b["w"]), b["x"] - (a["x"] + a["w"]), 0)
            dy = max(a["y"] - (b["y"] + b["h"]), b["y"] - (a["y"] + a["h"]), 0)
            distance = math.hypot(dx, dy)
            if distance <= near_distance:
                relations.append("near")
            if relations:
                edges.append({
                    "source": first["id"],
                    "target": second["id"],
                    "relations": sorted(set(relations)),
                    "distance_cm": round(distance, 3),
                })
    return {"nodes": nodes, "edges": edges}


def detect_geometry_patterns(graph):
    """Infer reusable composition patterns from the spatial graph."""
    patterns = []
    nodes = graph["nodes"]
    text_nodes = [node for node in nodes if node["kind"] == "text"]
    image_nodes = [node for node in nodes if node["kind"] == "image"]

    rows = _components(graph, "same_row")
    columns = _components(graph, "same_column")
    repeated_rows = [group for group in rows if len(group) >= 2]
    repeated_columns = [group for group in columns if len(group) >= 2]
    if repeated_rows:
        patterns.append(_pattern("horizontal_sequence", 0.78, repeated_rows[0],
                                 "two_or_more_nodes_share_a_row"))
    if len(repeated_columns) >= 2:
        members = sorted({member for group in repeated_columns for member in group})
        patterns.append(_pattern("multi_column", 0.8, members,
                                 "two_or_more_vertical_groups"))

    similar = []
    for index, first in enumerate(nodes):
        a = first["bounds_cm"]
        for second in nodes[index + 1:]:
            b = second["bounds_cm"]
            if (_ratio(abs(a["w"] - b["w"]), max(a["w"], b["w"])) <= 0.12 and
                    _ratio(abs(a["h"] - b["h"]), max(a["h"], b["h"])) <= 0.12):
                similar.extend([first["id"], second["id"]])
    if len(set(similar)) >= 3:
        patterns.append(_pattern("repeated_blocks", 0.82, sorted(set(similar)),
                                 "three_or_more_similarly_sized_nodes"))

    for image in image_nodes:
        ib = image["bounds_cm"]
        candidates = []
        for text in text_nodes:
            tb = text["bounds_cm"]
            overlap = _overlap(ib["x"], ib["x"] + ib["w"], tb["x"], tb["x"] + tb["w"])
            gap = tb["y"] - (ib["y"] + ib["h"])
            if _ratio(overlap, min(ib["w"], tb["w"])) >= 0.35 and -0.1 <= gap <= 1.5:
                candidates.append((gap, text["id"]))
        if candidates:
            text_id = min(candidates)[1]
            patterns.append(_pattern("image_caption_pair", 0.84,
                                     [image["id"], text_id],
                                     "text_is_directly_below_image"))
    return patterns


def _components(graph, relation):
    adjacency = {node["id"]: set() for node in graph["nodes"]}
    for edge in graph["edges"]:
        if relation in edge["relations"]:
            adjacency[edge["source"]].add(edge["target"])
            adjacency[edge["target"]].add(edge["source"])
    result, visited = [], set()
    for node, neighbors in adjacency.items():
        if node in visited or not neighbors:
            continue
        stack, group = [node], []
        while stack:
            current = stack.pop()
            if current in visited:
                continue
            visited.add(current)
            group.append(current)
            stack.extend(adjacency[current] - visited)
        result.append(sorted(group))
    return result


def _pattern(kind, confidence, members, evidence):
    return {
        "type": kind,
        "confidence": confidence,
        "members": members,
        "evidence": [evidence],
    }
