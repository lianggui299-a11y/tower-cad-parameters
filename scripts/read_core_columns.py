"""Add S2K core-column section intervals and actual composite constituent sizes."""

import argparse
from collections import Counter
import hashlib
import json
import math
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import parameters as p
import read_s2k_sections as s


def describe_section(model, name):
    section = model["sections"][name]
    scale = model["to_mm"]
    if section["Shape"] != "SD Section":
        return {"label": s.dimension_label(section, scale), "material": section["Material"],
                "height_mm": float(section["t3"]) * scale, "width_mm": float(section["t2"]) * scale,
                "constituents": [], "general": section}
    general = next(r for r in model["tables"]["SECTION DESIGNER PROPERTIES 01 - GENERAL"] if r["SectionName"] == name)
    tubes = [r for r in model["tables"]["SECTION DESIGNER PROPERTIES 09 - SHAPE BOX/TUBE"] if r["SectionName"] == name]
    rects = [r for r in model["tables"]["SECTION DESIGNER PROPERTIES 12 - SHAPE SOLID RECTANGLE"] if r["SectionName"] == name]
    shapes = tubes + rects
    if len(tubes) != 1 or len(shapes) != int(general["nTotalShp"]) or any(float(r["Rotation"]) != 0 for r in shapes):
        raise ValueError("Composite shape layout needs additional interpretation: " + name)
    def value(row, key):
        return float(row[key]) * scale
    def fmt(number):
        return f"{number:.9f}".rstrip("0").rstrip(".")
    tube = tubes[0]
    h, b, tf, tw = [value(tube, k) for k in ("Height", "Width", "FlngThick", "WebThick")]
    if not math.isclose(tf, tw, abs_tol=1e-6):
        raise ValueError("Unequal tube thickness needs explicit display: " + name)
    core = [r for r in rects if all(math.isclose(value(r, k), value(tube, k), abs_tol=1e-6) for k in ("XCenter", "YCenter"))
            and math.isclose(value(r, "Height"), h - 2 * tf, abs_tol=1e-6)
            and math.isclose(value(r, "Width"), b - 2 * tw, abs_tol=1e-6)]
    if len(core) != 1:
        raise ValueError("Cannot identify tube concrete core: " + name)
    xmin = min(value(r, "XCenter") - value(r, "Width") / 2 for r in shapes)
    xmax = max(value(r, "XCenter") + value(r, "Width") / 2 for r in shapes)
    ymin = min(value(r, "YCenter") - value(r, "Height") / 2 for r in shapes)
    ymax = max(value(r, "YCenter") + value(r, "Height") / 2 for r in shapes)
    height, width = ymax - ymin, xmax - xmin
    label = "钢管混凝土：" if len(rects) == 1 else f"外包混凝土{fmt(height)}x{fmt(width)}；内置"
    label += f"箱形钢管{fmt(h)}x{fmt(b)}x{fmt(tw)}；填芯{fmt(h - 2 * tf)}x{fmt(b - 2 * tw)} mm"
    steel_area = h * b - (h - 2 * tf) * (b - 2 * tw)
    concrete_area = sum(value(r, "Height") * value(r, "Width") for r in rects)
    return {"label": label, "material": "+".join(sorted({r["ShapeMat"] for r in shapes}, reverse=True)),
            "height_mm": height, "width_mm": width, "gross_area_mm2": steel_area + concrete_area,
            "steel_area_mm2": steel_area, "concrete_area_mm2": concrete_area,
            "constituents": shapes, "sd_general": general, "general": section}


def collect_columns(book, model, tolerance):
    columns, used = [], set()
    for column in p.records(book, "core_column"):
        frames = [f for f in model["frames"] if all(abs(f[end][i] - column[key]) <= tolerance
                  for end in ("a", "b") for i, key in enumerate(("x", "y")))
                  and abs(f["b"][2] - f["a"][2]) > tolerance]
        if not frames:
            raise ValueError("No vertical column on saved XY: " + column["id"])
        frames.sort(key=lambda f: min(f["a"][2], f["b"][2]))
        groups = []
        for frame in frames:
            lo, hi = sorted((frame["a"][2], frame["b"][2]))
            assignment = frame["assignment"]
            name = assignment["AnalSect"]
            if assignment.get("MatProp") not in (None, "Default", "N.A."):
                raise ValueError("Column material override requires constituent verification: " + frame["id"])
            if groups and abs(groups[-1]["z1"] - lo) > tolerance:
                raise ValueError("Column segments have a gap or overlap: " + column["id"])
            if groups and groups[-1]["source_section"] == name:
                groups[-1]["z1"] = hi
                groups[-1]["frames"].append(frame)
            else:
                groups.append({"z0": lo, "z1": hi, "source_section": name, "frames": [frame]})
            used.add(name)
        if abs(groups[-1]["z1"] - column["z"]) > tolerance:
            raise ValueError("S2K column top differs from saved top: " + column["id"])
        columns.append({"original": column, "intervals": groups})
    if not columns:
        raise ValueError("No original core-column locations in workbook")
    return columns, {name: describe_section(model, name) for name in sorted(used)}


def enrich(workbook, source, output, tolerance):
    book = p.read_book(workbook)
    model = s.read_model(source)
    source_hash = hashlib.sha256(Path(source).read_bytes()).hexdigest()
    if source_hash != book["settings"].get("member_sections_source_sha256"):
        raise ValueError("S2K does not match the existing section source")
    columns, definitions = collect_columns(book, model, tolerance)
    wb = load_workbook(workbook)
    sheet_name = columns[0]["original"]["_sheet"]
    ws = wb[sheet_name]
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    ws.delete_rows(1, ws.max_row)
    ws.data_validations.dataValidation = []
    schema = p.SCHEMAS["core_column_section"]
    ws.append(["核心筒柱 | 每个连续截面区段一行；尺寸及标高mm；实际S2K指派，待人工核查"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(schema))
    ws.append([label for _, label in schema])
    trace = next(r for r in book["sheets"] if r["role"] == "trace")
    trace_ws = wb[trace["name"]]
    def add_trace(identity, table, line, detail):
        trace_ws.append([identity, str(Path(source).resolve()), table, line, None, json.dumps(detail, ensure_ascii=False)])
    for column in columns:
        original = column["original"]
        add_trace("S2K-CORE-ORIGINAL-" + original["id"], "original workbook core_column", original["_row"], original)
        for index, group in enumerate(column["intervals"], 1):
            description = definitions[group["source_section"]]
            identity = f'{original["id"]}-SEC{index:02d}'
            note = f'原S2K截面名={group["source_section"]}; 模型段=' + ",".join(f["id"] for f in group["frames"])
            row = {"id": identity, "column": original["id"], "x": original["x"], "y": original["y"],
                   "z0": group["z0"], "z1": group["z1"], "section": description["label"],
                   "material": description["material"], "status": original["status"], "note": note}
            ws.append([row.get(key) for key, _ in schema])
            add_trace("S2K-CORE-" + identity, "FRAME SECTION ASSIGNMENTS", group["frames"][0]["assignment"]["_line"], group)
    s.style_table(ws, [22, 14, 14, 14, 18, 18, 75, 18, 16, 65])
    for row in range(3, ws.max_row + 1):
        ws.row_dimensions[row].height = 48
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor="EEF6FC")
    for row in wb["__schema"].iter_rows(min_row=2):
        if row[0].value == sheet_name:
            row[1].value = "core_column_section"
            row[2].value = json.dumps([key for key, _ in schema])
    summary = wb["截面定义_S2K"]
    for name, detail in definitions.items():
        general = detail["general"]
        def scaled(key, power):
            return float(general[key]) * model["to_mm"] ** power if key in general else None
        composite = bool(detail["constituents"])
        explanation = detail["label"]
        if composite:
            explanation += "; SD面积/惯性矩为基准材料" + detail["sd_general"]["BaseMat"] + "的等效参数；组成形状及位置见来源记录"
        summary.append([detail["label"], general["Shape"], detail["material"], detail["height_mm"], detail["width_mm"],
                        None if composite else scaled("tf", 1), None if composite else scaled("tw", 1), None, None,
                        scaled("Area", 2), scaled("I22", 4), scaled("I33", 4), scaled("TorsConst", 4), general["_line"], explanation, name])
        add_trace("S2K-CORE-SECTION-" + name, "FRAME SECTION PROPERTIES / SECTION DESIGNER", general["_line"], detail)
    s.style_table(summary, [75, 22, 18, 14, 14, 16, 16, 14, 16, 18, 22, 22, 22, 14, 80, 32])
    trace_ws.auto_filter.ref = trace_ws.dimensions.replace("A1:", "A2:")
    wb["项目设置"].append(["core_column_sections", "S2K按同XY竖向连续区段读取；合并连续同截面；保留SD组成形状，尺寸mm"])
    wb.save(output)
    wb.close()
    result = p.analyze(p.read_book(output))
    wb = load_workbook(output)
    pending = wb["待核清单"] if "待核清单" in wb else wb.create_sheet("待核清单")
    pending.delete_rows(1, pending.max_row)
    pending.append(["级别", "类型", "工作表", "行号", "记录编号", "问题"])
    for finding in result["issues"]:
        if finding["code"] != "pending":
            pending.append([finding[k] for k in ("severity", "code", "sheet", "row", "id", "message")])
    s.style_table(pending, [12, 30, 24, 10, 28, 90], header=1)
    wb.save(output)
    wb.close()
    return {"source": str(Path(source).resolve()), "source_sha256": source_hash,
            "workbook": str(Path(workbook).resolve()), "workbook_sha256": hashlib.sha256(Path(workbook).read_bytes()).hexdigest(),
            "output_workbook": str(Path(output).resolve()), "tolerance_mm": tolerance,
            "columns": columns, "section_definitions": definitions, "parameter_findings": result["issues"],
            "counts": {"columns": len(columns), "source_frames": sum(len(g["frames"]) for c in columns for g in c["intervals"]),
                       "section_intervals": sum(len(c["intervals"]) for c in columns), "section_types": len(definitions)}}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("s2k", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--tolerance-mm", default=0.01, type=float)
    args = parser.parse_args()
    if args.output.exists() or args.report.exists():
        raise FileExistsError("Output already exists")
    report = enrich(args.workbook, args.s2k, args.output, args.tolerance_mm)
    p.write_json(args.report, report)
    print(json.dumps(report["counts"]))


if __name__ == "__main__":
    main()
