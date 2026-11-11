"""Read DXF evidence and propose source LINEs using explicit plan settings."""

import argparse
from collections import Counter
import hashlib
import json
import math
from pathlib import Path

import ezdxf


def source_handle(entity):
    seen = set()
    while getattr(entity, "source_of_copy", None) is not None:
        if id(entity) in seen:
            break
        seen.add(id(entity))
        entity = entity.source_of_copy
    return entity.dxf.get("handle", "UNRESOLVED")


def inventory(doc):
    records = []

    def walk(entity, path=(), inherited_layer=None):
        kind = entity.dxftype()
        handle = source_handle(entity)
        layer = entity.dxf.get("layer", "0")
        if layer == "0" and inherited_layer is not None:
            layer = inherited_layer
        row = {"kind": kind, "handle": handle, "block_path": list(path), "layer": layer}
        if kind == "INSERT":
            # virtual_entities applies insertion, scale, rotation, and nested transforms.
            # MINSERT must expand its array before virtual_entities is called.
            inserts = entity.multi_insert() if entity.mcount > 1 else [entity]
            for index, insert in enumerate(inserts):
                child_path = (*path, f"{handle}[{index}]")
                for attr in insert.attribs:
                    walk(attr, child_path, layer)
                def skipped(child, reason):
                    records.append({**row, "skipped_kind": child.dxftype(), "reason": str(reason)})
                for child in insert.virtual_entities(skipped_entity_callback=skipped):
                    walk(child, child_path, layer)
            return
        if kind == "LINE":
            row.update(start=list(entity.dxf.start), end=list(entity.dxf.end))
        elif kind in {"TEXT", "MTEXT", "ATTRIB"}:
            row.update(text=entity.plain_text(), insert=list(entity.dxf.insert))
        elif kind in {"CIRCLE", "ARC"}:
            row.update(center=list(entity.ocs().to_wcs(entity.dxf.center)), radius=entity.dxf.radius)
            if kind == "ARC":
                row.update(start_angle=entity.dxf.start_angle, end_angle=entity.dxf.end_angle)
        elif kind == "LWPOLYLINE":
            row.update(points=[list(p) for p in entity.vertices_in_wcs()], closed=entity.closed,
                       bulges=[p[4] for p in entity.get_points()])
        elif kind == "POLYLINE":
            row.update(points=[list(p) for p in entity.points_in_wcs()], closed=entity.is_closed)
        records.append(row)

    for entity in doc.modelspace():
        walk(entity)
    return records


def candidates(records, config):
    result = []
    keys = set()
    for plan in config["plans"]:
        if plan["key"] in keys:
            raise ValueError("Plan keys must be unique")
        keys.add(plan["key"])
        scale = float(plan["scale_to_mm"])
        if scale <= 0 or not math.isfinite(scale):
            raise ValueError("scale_to_mm must be positive and finite")
        x0, y0 = plan["origin"]
        angle = math.radians(plan["rotation_deg"])
        xlo, ylo, xhi, yhi = plan["bbox"]
        if xhi <= xlo or yhi <= ylo or not plan["line_layers"]:
            raise ValueError("Require a nonempty layer selection and a valid source bbox")
        def transform(point):
            x, y = (point[0] - x0) * scale, (point[1] - y0) * scale
            return [x * math.cos(angle) - y * math.sin(angle),
                    x * math.sin(angle) + y * math.cos(angle), plan["level_m"] * 1000]
        def inside(point):
            return xlo <= point[0] <= xhi and ylo <= point[1] <= yhi
        members, crossings, nonplanar = [], [], []
        for row in records:
            if row["kind"] != "LINE" or row["layer"] not in plan["line_layers"]:
                continue
            a, b = row["start"], row["end"]
            # Segment bbox overlap also catches lines crossing the region with both ends outside.
            overlap = (max(a[0], b[0]) >= xlo and min(a[0], b[0]) <= xhi
                       and max(a[1], b[1]) >= ylo and min(a[1], b[1]) <= yhi)
            if not overlap:
                continue
            if not (inside(a) and inside(b)):
                crossings.append(row)
                continue
            if abs(a[2] - b[2]) * scale > 0.01:
                nonplanar.append(row)
                continue
            if math.dist(a, b) * scale <= 1e-6:
                continue
            identity = "/".join([plan["key"], *row["block_path"], row["handle"]])
            members.append({"source_id": identity, "start": transform(a), "end": transform(b),
                            "source": row, "status": "candidate; connectivity unconfirmed"})
        result.append({"plan": plan, "members": members, "crossing_boundary": crossings,
                       "nonplanar_lines": nonplanar})
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("drawing", type=Path)
    parser.add_argument("--config", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    if args.drawing.suffix.lower() != ".dxf":
        parser.error("Convert DWG to a verified DXF first")
    doc = ezdxf.readfile(args.drawing)
    records = inventory(doc)
    output = {"source": str(args.drawing.resolve()),
              "sha256": hashlib.sha256(args.drawing.read_bytes()).hexdigest(),
              "declared_units": doc.header.get("$INSUNITS", 0),
              "counts": dict(Counter(f'{r["layer"]} / {r["kind"]}' for r in records)),
              "entities": records}
    if args.config:
        output["plans"] = candidates(records, json.loads(args.config.read_text(encoding="utf-8-sig")))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("x", encoding="utf-8") as stream:
        json.dump(output, stream, ensure_ascii=False, indent=2, allow_nan=False)
    print(json.dumps({"entities": len(records), "output": str(args.output)}, ensure_ascii=True))


if __name__ == "__main__":
    main()
