"""Read S2K section assignments and match collinear frame coverage to member rows."""

import argparse
from collections import Counter
import hashlib
import json
import math
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import parameters


FIELDS = re.compile(r'(\w+)=("[^"\n]*"|\S+)')
TABLES = {
    "PROGRAM CONTROL", "JOINT COORDINATES", "CONNECTIVITY - FRAME",
    "FRAME SECTION ASSIGNMENTS", "FRAME SECTION PROPERTIES 01 - GENERAL",
    "SECTION DESIGNER PROPERTIES 01 - GENERAL",
    "SECTION DESIGNER PROPERTIES 09 - SHAPE BOX/TUBE",
    "SECTION DESIGNER PROPERTIES 12 - SHAPE SOLID RECTANGLE",
}


def read_tables(path):
    tables = {name: [] for name in TABLES}
    table, pieces, start = None, [], None
    with Path(path).open(encoding="utf-8-sig") as stream:
        for number, line in enumerate(stream, 1):
            if line.strip() == "END TABLE DATA":
                break
            if line.startswith("TABLE:"):
                table = line.split('"')[1]
            elif table in TABLES and line.strip():
                if not pieces:
                    start = number
                text = line.strip()
                continued = text.endswith(" _")
                pieces.append(text[:-2] if continued else text)
                if not continued:
                    row = {key: value.strip('"') for key, value in FIELDS.findall(" ".join(pieces))}
                    row["_line"] = start
                    tables[table].append(row)
                    pieces = []
    if pieces:
        raise ValueError("Unfinished S2K continuation record")
    return tables


def read_model(path):
    tables = read_tables(path)
    units = tables["PROGRAM CONTROL"][0]["CurrUnits"]
    length_unit = units.split(",")[1].strip()
    to_mm = {"mm": 1, "cm": 10, "m": 1000, "in": 25.4, "ft": 304.8}[length_unit]
    nodes = {}
    for row in tables["JOINT COORDINATES"]:
        if all(key in row for key in ("GlobalX", "GlobalY", "GlobalZ")):
            xyz = [float(row[key]) * to_mm for key in ("GlobalX", "GlobalY", "GlobalZ")]
        elif row.get("CoordSys") == "GLOBAL" and row.get("CoordType") == "Cartesian":
            xyz = [float(row[key]) * to_mm for key in ("XorR", "Y", "Z")]
        else:
            raise ValueError("Missing global joint coordinates: " + row["Joint"])
        nodes[row["Joint"]] = xyz
    assignments = {r["Frame"]: r for r in tables["FRAME SECTION ASSIGNMENTS"]}
    sections = {r["SectionName"]: r for r in tables["FRAME SECTION PROPERTIES 01 - GENERAL"]}
    frames = []
    for row in tables["CONNECTIVITY - FRAME"]:
        if row.get("IsCurved", "No") != "No":
            continue
        frames.append({"id": row["Frame"], "a": nodes[row["JointI"]], "b": nodes[row["JointJ"]],
                       "assignment": assignments.get(row["Frame"], {}), "connectivity_line": row["_line"]})
    return {"units": units, "to_mm": to_mm, "frames": frames, "sections": sections, "tables": tables}


def match_member(row, frames, tolerance):
    a = [row[k] for k in ("xi", "yi", "zi")]
    b = [row[k] for k in ("xj", "yj", "zj")]
    length = math.dist(a, b)
    if length <= tolerance:
        return {"status": "invalid_member", "pieces": []}
    axis = [(b[i] - a[i]) / length for i in range(3)]
    pieces = []
    for frame in frames:
        if any(max(frame["a"][i], frame["b"][i]) < min(a[i], b[i]) - tolerance
               or min(frame["a"][i], frame["b"][i]) > max(a[i], b[i]) + tolerance for i in range(3)):
            continue
        projections, offsets = [], []
        for point in (frame["a"], frame["b"]):
            distance = sum((point[i] - a[i]) * axis[i] for i in range(3))
            projections.append(distance)
            offsets.append(math.dist(point, [a[i] + distance * axis[i] for i in range(3)]))
        if max(offsets) > tolerance:
            continue
        lo, hi = max(0, min(projections)), min(length, max(projections))
        if hi - lo <= tolerance:
            continue
        assignment = frame["assignment"]
        pieces.append({"frame": frame["id"], "start_mm": lo, "end_mm": hi,
                       "section": assignment.get("AnalSect"), "material_override": assignment.get("MatProp"),
                       "auto_select": assignment.get("AutoSelect"), "section_type": assignment.get("SectionType"),
                       "axis_error_mm": max(offsets), "assignment_line": assignment.get("_line"),
                       "connectivity_line": frame["connectivity_line"]})
    pieces.sort(key=lambda p: (p["start_mm"], p["end_mm"], p["frame"]))
    cursor, gaps, overlaps = 0, [], []
    for piece in pieces:
        lo, hi = piece["start_mm"], piece["end_mm"]
        if lo > cursor + tolerance:
            gaps.append([cursor, lo])
        if lo < cursor - tolerance:
            overlaps.append([lo, min(cursor, hi)])
        cursor = max(cursor, hi)
    if cursor < length - tolerance:
        gaps.append([cursor, length])
    status = "matched"
    if not pieces:
        status = "unmatched"
    elif overlaps:
        status = "overlapping_frames"
    elif gaps:
        status = "incomplete_coverage"
    elif any(not p["section"] or p["section"] in {"None", "N.A."} for p in pieces):
        status = "missing_assignment"
    return {"status": status, "length_mm": length, "gaps_mm": gaps, "overlaps_mm": overlaps, "pieces": pieces}


def inspect(workbook, source, tolerance):
    book = parameters.read_book(workbook)
    model = read_model(source)
    rows = []
    for row in parameters.records(book, "members"):
        if parameters.active(row):
            rows.append({"id": row["id"], "sheet": row["_sheet"], "row": row["_row"],
                         **match_member(row, model["frames"], tolerance)})
    return model, {"source": str(Path(source).resolve()),
                   "source_sha256": hashlib.sha256(Path(source).read_bytes()).hexdigest(),
                   "workbook": str(Path(workbook).resolve()),
                   "workbook_sha256": hashlib.sha256(Path(workbook).read_bytes()).hexdigest(),
                   "source_units": model["units"], "tolerance_mm": tolerance,
                   "counts": dict(Counter(r["status"] for r in rows)), "members": rows}


def section_schedule(pieces, tolerance):
    schedule = []
    for piece in pieces:
        if schedule and piece["section"] == schedule[-1]["section"] and abs(piece["start_mm"] - schedule[-1]["end_mm"]) <= tolerance:
            schedule[-1]["end_mm"] = piece["end_mm"]
        else:
            schedule.append({k: piece[k] for k in ("start_mm", "end_mm", "section")})
    return schedule


def dimension_label(section, to_mm):
    def size(key):
        return f'{float(section[key]) * to_mm:.9f}'.rstrip("0").rstrip(".")
    if section["Shape"] == "I/Wide Flange":
        if size("t2") == size("t2b") and size("tf") == size("tfb"):
            return "工字形 " + "x".join(size(k) for k in ("t3", "t2", "tw", "tf")) + " mm"
        return (f'工字形 高{size("t3")}; 上翼缘{size("t2")}x{size("tf")}; '
                f'腹板厚{size("tw")}; 下翼缘{size("t2b")}x{size("tfb")} mm')
    if section["Shape"] == "Box/Tube":
        if size("tw") == size("tf"):
            return "箱形 " + "x".join(size(k) for k in ("t3", "t2", "tw")) + " mm"
        return "箱形 " + "x".join(size(k) for k in ("t3", "t2", "tw", "tf")) + " mm（腹板厚/翼缘厚）"
    raise ValueError("Dimension display needs section-specific handling: " + section["SectionName"])


def section_value(names, model):
    labels = list(dict.fromkeys(dimension_label(model["sections"][name], model["to_mm"]) for name in names))
    return labels[0] if len(labels) == 1 else "分段：" + " / ".join(labels)


def style_table(ws, widths, header=2):
    ws.freeze_panes = f"C{header + 1}"
    ws.auto_filter.ref = f"A{header}:{ws.cell(ws.max_row, len(widths)).coordinate}"
    for column, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(header, column).column_letter].width = width
    for cell in ws[header]:
        cell.fill = PatternFill("solid", fgColor="254D61")
        cell.font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True)
    for row in ws.iter_rows(min_row=header + 1):
        for cell in row:
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if parameters.numeric(cell.value):
                cell.number_format = "0.###"


def save_enriched(workbook, output, model, report):
    output = Path(output)
    if output.exists():
        raise FileExistsError(output)
    wb = load_workbook(workbook)
    book = parameters.read_book(workbook)
    trace = next(s for s in book["sheets"] if s["role"] == "trace")
    trace_ws = wb[trace["name"]]
    trace_keys = [k for k, _ in parameters.SCHEMAS["trace"]]
    next_trace = max((r["_row"] for r in trace["rows"]), default=2) + 1

    def add_trace(identity, table, line, detail):
        nonlocal next_trace
        values = {"id": identity, "source": report["source"], "sheet": table, "row": line,
                  "detail": json.dumps(detail, ensure_ascii=False, allow_nan=False)}
        for col, key in enumerate(trace_keys, 1):
            trace_ws.cell(next_trace, col, values.get(key))
        next_trace += 1

    used, imported, unresolved, conflicts = set(), 0, [], []
    for item in report["members"]:
        ws = wb[item["sheet"]]
        columns = {key: index for index, (key, _) in enumerate(parameters.SCHEMAS["members"], 1)}
        number = item["row"]
        old_note = ws.cell(number, columns["note"]).value or ""
        if item["status"] != "matched":
            reason = {"unmatched": "无同轴模型段；现有模型几何可能已改版",
                      "incomplete_coverage": "模型段未完整覆盖当前梁",
                      "overlapping_frames": "存在重叠模型段，无法唯一匹配"}.get(item["status"], item["status"])
            ws.cell(number, columns["note"], old_note + "; S2K截面待对应：" + reason)
            unresolved.append(item["id"])
            add_trace("S2K-" + item["id"], "FRAME SECTION ASSIGNMENTS", None, item)
            continue
        names = sorted({p["section"] for p in item["pieces"]})
        for name in names:
            if name not in model["sections"]:
                raise ValueError("Missing section definition: " + name)
        schedule = section_schedule(item["pieces"], report["tolerance_mm"])
        # Keep one physical-member row even if model subsegments have different sections.
        value = section_value(names, model)
        legacy_value = names[0] if len(names) == 1 else "分段：" + " / ".join(names)
        old = ws.cell(number, columns["section"]).value
        if old not in (None, "", value, legacy_value):
            conflicts.append({"id": item["id"], "existing": old, "s2k": value})
            continue
        ws.cell(number, columns["section"], value)
        materials = {p["material_override"] if p["material_override"] not in (None, "Default", "N.A.")
                     else model["sections"][p["section"]].get("Material") for p in item["pieces"]}
        if len(materials) == 1 and None not in materials:
            material = next(iter(materials))
            existing = ws.cell(number, columns["material"]).value
            if existing not in (None, "", material):
                raise ValueError("Existing material differs: " + item["id"])
            ws.cell(number, columns["material"], material)
        frames = [p["frame"] for p in item["pieces"]]
        detail = f'S2K截面：{Path(report["source"]).stem}; 模型段=' + ",".join(frames)
        detail += "; 原S2K截面名=" + " / ".join(names)
        detail += f'; 轴线容差{report["tolerance_mm"]:g}mm; 最大偏差{max(p["axis_error_mm"] for p in item["pieces"]):.6f}mm'
        if len(names) > 1:
            detail += "; 距梁起点分段/m：" + "; ".join(
                f'{p["start_mm"] / 1000:.6f}~{p["end_mm"] / 1000:.6f}={dimension_label(model["sections"][p["section"]], model["to_mm"])}' for p in schedule)
        ws.cell(number, columns["note"], old_note + "; " + detail)
        ws.column_dimensions[ws.cell(2, columns["section"]).column_letter].width = 43
        ws.cell(number, columns["section"]).alignment = Alignment(vertical="center", wrap_text=True)
        item["section_schedule"] = schedule
        used.update(names)
        imported += 1
        add_trace("S2K-" + item["id"], "FRAME SECTION ASSIGNMENTS", item["pieces"][0]["assignment_line"], item)
    if conflicts:
        raise ValueError("Existing sections differ; saved edits were preserved: " + json.dumps(conflicts, ensure_ascii=False))

    definitions = wb.create_sheet("截面定义_S2K")
    definitions.append(["S2K截面实际参数 | 尺寸以数值字段为准，不能从截面名称反推 | 来源：" + Path(report["source"]).name])
    definitions.merge_cells(start_row=1, start_column=1, end_row=1, end_column=16)
    definitions.append(["截面实际尺寸/mm", "类型", "材料", "高度/mm", "上宽/mm", "上翼缘厚/mm", "腹板厚/mm",
                        "下宽/mm", "下翼缘厚/mm", "面积/mm²", "I22/mm⁴", "I33/mm⁴", "J/mm⁴", "S2K行号", "说明", "原S2K截面名"])
    for name in sorted(used):
        section = model["sections"][name]
        def scaled(key, power=1):
            return float(section[key]) * model["to_mm"] ** power if key in section else None
        definitions.append([dimension_label(section, model["to_mm"]), section.get("Shape"), section.get("Material"),
                            *[scaled(k) for k in ("t3", "t2", "tf", "tw", "t2b", "tfb")],
                            scaled("Area", 2), scaled("I22", 4), scaled("I33", 4), scaled("TorsConst", 4),
                            section["_line"], section.get("Notes"), name])
        add_trace("S2K-SECTION-" + name, "FRAME SECTION PROPERTIES 01 - GENERAL", section["_line"], section)
    style_table(definitions, [43, 22, 14, 14, 14, 16, 16, 14, 16, 18, 22, 22, 22, 14, 60, 32])
    report["section_definitions"] = {name: model["sections"][name] for name in sorted(used)}
    report["imported"] = imported
    report["unresolved_ids"] = unresolved
    trace_ws.auto_filter.ref = f"A2:{trace_ws.cell(trace_ws.max_row, len(trace_keys)).coordinate}"
    settings = wb["项目设置"]
    settings.append(["member_sections_source", report["source"]])
    settings.append(["member_sections_source_sha256", report["source_sha256"]])
    settings.append(["section_match_tolerance_mm", report["tolerance_mm"]])
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    wb.close()
    result = parameters.analyze(parameters.read_book(output))
    wb = load_workbook(output)
    pending = wb["待核清单"] if "待核清单" in wb else wb.create_sheet("待核清单")
    pending.delete_rows(1, pending.max_row)
    pending.append(["级别", "类型", "工作表", "行号", "记录编号", "问题"])
    for finding in result["issues"]:
        if finding["code"] != "pending":
            pending.append([finding[k] for k in ("severity", "code", "sheet", "row", "id", "message")])
    style_table(pending, [12, 30, 24, 10, 28, 90], header=1)
    wb.save(output)
    wb.close()
    report["output_workbook"] = str(output.resolve())
    report["parameter_findings"] = result["issues"]


def format_existing(workbook, source, output):
    model = read_model(source)
    source_hash = hashlib.sha256(Path(source).read_bytes()).hexdigest()
    book = parameters.read_book(workbook)
    if book["settings"].get("member_sections_source_sha256") != source_hash:
        raise ValueError("The workbook's imported section source does not match this S2K")
    wb = load_workbook(workbook)
    changes, used, empty = [], set(), 0
    for row in parameters.records(book, "members"):
        original = row.get("section")
        if not original:
            empty += 1
            continue
        names = original.removeprefix("分段：").split(" / ")
        if any(name not in model["sections"] for name in names):
            raise ValueError("Section was edited or has no source definition: " + row["id"])
        value = section_value(names, model)
        ws = wb[row["_sheet"]]
        columns = {key: i for i, (key, _) in enumerate(parameters.SCHEMAS["members"], 1)}
        cell = ws.cell(row["_row"], columns["section"], value)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[cell.column_letter].width = 43
        original_names = "; 原S2K截面名=" + " / ".join(names)
        note = row.get("note") or ""
        if original.startswith("分段："):
            for name in sorted(names, key=len, reverse=True):
                note = note.replace("=" + name, "=" + dimension_label(model["sections"][name], model["to_mm"]))
            ws.row_dimensions[row["_row"]].height = max(ws.row_dimensions[row["_row"]].height or 26, 48)
        ws.cell(row["_row"], columns["note"], note + original_names)
        changes.append({"id": row["id"], "sheet": row["_sheet"], "row": row["_row"],
                        "original": original, "dimensions": value})
        used.update(names)
    for sheet in book["sheets"]:
        if sheet["role"] == "members":
            ws = wb[sheet["name"]]
            ws.cell(1, 1, ws.cell(1, 1).value + " | 截面单位mm：工字形高x宽x腹板厚x翼缘厚；箱形高x宽x壁厚")
    definitions = wb["截面定义_S2K"]
    definitions.cell(2, 1, "截面实际尺寸/mm")
    definitions.cell(2, 16, "原S2K截面名")
    for number in range(3, definitions.max_row + 1):
        name = definitions.cell(number, 1).value
        if name:
            definitions.cell(number, 1, dimension_label(model["sections"][name], model["to_mm"]))
            definitions.cell(number, 16, name)
    style_table(definitions, [43, 22, 14, 14, 14, 16, 16, 14, 16, 18, 22, 22, 22, 14, 60, 32])
    wb["项目设置"].append(["section_display", "实际尺寸/mm；工字形高x宽x腹板厚x翼缘厚；等厚箱形高x宽x壁厚"])
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    wb.close()
    result = parameters.analyze(parameters.read_book(output))
    return {"source": str(Path(source).resolve()), "source_sha256": source_hash,
            "workbook": str(Path(workbook).resolve()),
            "workbook_sha256": hashlib.sha256(Path(workbook).read_bytes()).hexdigest(),
            "output_workbook": str(Path(output).resolve()), "source_units": model["units"],
            "counts": {"dimension_display": len(changes), "blank_pending": empty}, "imported": len(changes),
            "section_definitions": {name: model["sections"][name] for name in sorted(used)},
            "changes": changes, "parameter_findings": result["issues"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("s2k", type=Path)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--output", type=Path, help="Save an enriched workbook copy; omit for a matching report only")
    parser.add_argument("--format-existing", action="store_true", help="Display existing imported section names as actual dimensions without rematching geometry")
    parser.add_argument("--tolerance-mm", default=0.01, type=float)
    args = parser.parse_args()
    if args.tolerance_mm <= 0:
        parser.error("Tolerance must be positive")
    if args.report.exists() or (args.output and args.output.exists()):
        raise FileExistsError("Output workbook or report already exists")
    if args.format_existing:
        if not args.output:
            parser.error("--format-existing requires --output")
        report = format_existing(args.workbook, args.s2k, args.output)
    else:
        model, report = inspect(args.workbook, args.s2k, args.tolerance_mm)
        if args.output:
            save_enriched(args.workbook, args.output, model, report)
    parameters.write_json(args.report, report)
    print(json.dumps({"matches": report["counts"], "imported": report.get("imported"),
                      "section_definitions": len(report.get("section_definitions", {}))}, ensure_ascii=False))


if __name__ == "__main__":
    main()
