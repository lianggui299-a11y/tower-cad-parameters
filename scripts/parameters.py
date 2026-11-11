"""Create, check, and draw editable software-neutral tower parameter workbooks."""

import argparse
from collections import defaultdict
import colorsys
import hashlib
import json
import math
from pathlib import Path
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.properties import CalcProperties


SCHEMAS = {
    "members": "id:构件编号 level:楼层/m type:构件类型 section:截面 material:材料 ni:起点编号 xi:起X/mm yi:起Y/mm zi:起Z/mm nj:终点编号 xj:终X/mm yj:终Y/mm zj:终Z/mm ci:起端连接 cj:终端连接 keep:保留/排除 status:核查状态 note:备注",
    "wall": "id:区段编号 z0:下标高/mm z1:上标高/mm r0:下端外半径/mm r1:上端外半径/mm t0:下端法向厚度/mm t1:上端法向厚度/mm material:材料 status:核查状态 note:备注",
    "point": "id:荷载编号 level:楼层/m node:节点编号 x:X/mm y:Y/mm z:Z/mm member:受力构件 case:荷载工况 fx:FX/kN fy:FY/kN fz:FZ/kN mx:MX/(kN·m) my:MY/(kN·m) mz:MZ/(kN·m) status:核查状态 note:备注",
    "line": "id:荷载编号 level:楼层/m member:构件编号 a:起始距离/m b:终止距离/m category:墙类别 density:容重/(kN/m³) thickness:厚度/m height:高度/m q:q/(kN/m) case:荷载工况 direction:作用方向 status:核查状态 note:备注",
    "surface": "id:面编号 level:楼层/m nodes:边界节点序列 category:区域类别 dead:恒载/(kN/m²) live:活载/(kN/m²) status:核查状态 note:备注",
    "category": "id:区域类别 dead:恒载/(kN/m²) live:活载/(kN/m²) note:取值说明",
    "surface_node": "id:节点编号 x:X/mm y:Y/mm z:Z/mm member:所在构件 status:核查状态 note:备注",
    "pile": "id:桩编号 node:节点编号 x:X/mm y:Y/mm z:支承Z/mm diameter:桩径/mm system:支承坐标系 kx:Kx/(kN/m) ky:Ky/(kN/m) kz:Kz/(kN/m) krx:Krx/(kN·m/rad) kry:Kry/(kN·m/rad) krz:Krz/(kN·m/rad) behavior:支承特性 status:核查状态 note:备注",
    "slab": "id:底板编号 material:材料 thickness:厚度/mm z:中面Z/mm status:核查状态 note:备注",
    "boundary": "slab:底板编号 loop:轮廓编号 kind:轮廓类型 order:顶点顺序 node:节点编号 x:X/mm y:Y/mm z:Z/mm status:核查状态 note:备注",
    "core_column": "id:柱编号 level:标高/m x:X/mm y:Y/mm z:Z/mm section:截面 status:核查状态 note:备注",
    "core_column_section": "id:区段编号 column:柱编号 x:X/mm y:Y/mm z0:下标高/mm z1:上标高/mm section:截面 material:材料 status:核查状态 note:备注",
    "wall_audit": "id:对照编号 level:标高/m cad_inner:CAD内壁半径/mm cad_outer:CAD外壁半径/mm model_outer:模型外壁半径/mm model_inner_lower:模型下段内壁半径/mm model_inner_upper:模型上段内壁半径/mm delta_inner_lower:内壁差下侧/mm delta_inner_upper:内壁差上侧/mm delta_outer:外壁差/mm status:核查状态 note:说明 source:CAD识别数据来源",
    "trace": "id:记录编号 source:来源路径 sheet:来源工作表 row:来源行 handle:源实体句柄 detail:来源说明",
}
SCHEMAS = {role: [tuple(pair.split(":", 1)) for pair in text.split()] for role, text in SCHEMAS.items()}
DERIVED = {"line": {"q"}, "surface": {"dead", "live"}}


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def nodes_list(value):
    return [s.strip() for s in re.split(r"[,，;；\s]+", str(value or "")) if s.strip()]


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2, allow_nan=False)


def formula(role, key, row, category_sheet="荷载类别取值"):
    if role == "line" and key == "q":
        return f'=IF(COUNT(G{row}:I{row})<3,"",G{row}*H{row}*I{row})'
    if role == "surface" and key in {"dead", "live"}:
        sheet = "'" + category_sheet.replace("'", "''") + "'"
        column = "B" if key == "dead" else "C"
        # INDEX/MATCH returns a blank cell as zero; COUNT tests that matched cell first.
        cell = f"INDEX({sheet}!{column}$3:{column}$10002,MATCH(D{row},{sheet}!A$3:A$10002,0))"
        return f'=IFERROR(IF(COUNT({cell})=0,"",{cell}),"")'
    return None


def create(data, output):
    if Path(output).exists():
        raise FileExistsError(output)
    wb = Workbook()
    help_sheet = wb.active
    help_sheet.title = "使用说明"
    help_rows = [
        [data["project"], "参数核查草案；工程参数确认完成后方可交接建模。"],
        ["编辑方式", "浅蓝为可编辑输入，浅灰为计算公式；在构件行直接改端点坐标。"],
        ["编号", "实际连接处共用节点号；同节点坐标冲突需逐项核实，不能任取一条。"],
        ["单位", "坐标/筒壁尺寸/桩径：mm；点力：kN；点矩：kN·m；线区间：m。"],
        ["荷载单位", "容重：kN/m³；墙厚/墙高：m；线载：kN/m；面载：kN/m²。"],
        ["筒壁厚度", "法向厚度；原有径向厚度需核实换算，不能直接改列名。"],
        ["点荷载", "每作用点每工况一行；同点不同工况不合并。"],
        ["线荷载", "距构件起点定义作用区间；q由容重×厚度×高度自动计算。"],
        ["面荷载", "填写四个有序角点节点号，逗号分隔，直线闭合为一个四边形；曲边以弦代替，不加点拟合。"],
        ["基础", "维护桩位、等效支承、底板参数及边界；不维护有限元壳网格。"],
        ["空白与零", "未知留空；零表示已明确为零。未确认不能作为无荷载或无约束。"],
        ["修改后", "保存Excel，再由Codex检查当前文件并重新生成PDF；不会覆盖本表。"],
        ["追溯", "详见末尾来源记录；历史样例不等同于当前已确认参数。"],
    ]
    for row in help_rows:
        help_sheet.append(row)
    help_sheet.column_dimensions["A"].width = 22
    help_sheet.column_dimensions["B"].width = 100
    settings = wb.create_sheet("项目设置")
    settings.append(["设置项", "当前值"])
    for key, value in {"node_tolerance_mm": 0.01, **data.get("settings", {})}.items():
        settings.append([key, value if isinstance(value, (str, int, float)) else json.dumps(value, ensure_ascii=False)])
    settings.column_dimensions["A"].width = 32
    settings.column_dimensions["B"].width = 100
    meta = wb.create_sheet("__schema")
    meta.append(["schema_version", 1])
    category_sheets = [s["name"] for s in data["sheets"] if s["role"] == "category"]
    if len(category_sheets) > 1:
        raise ValueError("Use one category table")
    category_name = category_sheets[0] if category_sheets else "荷载类别取值"
    for definition in data["sheets"]:
        role, name = definition["role"], definition["name"]
        schema = SCHEMAS[role]
        if name in wb.sheetnames:
            raise ValueError(f"Duplicate/reserved sheet: {name}")
        ws = wb.create_sheet(name)
        ws.append([name + " | 浅蓝输入，浅灰计算；坐标可直接修改"])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(schema))
        ws.append([label for _, label in schema])
        keys = [key for key, _ in schema]
        meta.append([name, role, json.dumps(keys)])
        source_rows = definition.get("rows", [])
        for number in range(3, max(13, len(source_rows) + 3)):
            source = source_rows[number - 3] if number - 3 < len(source_rows) else {}
            for col, (key, _) in enumerate(schema, 1):
                value = source.get(key)
                generated = formula(role, key, number, category_name)
                if generated is not None:
                    value = generated
                cell = ws.cell(number, col, value)
                if isinstance(value, str) and len(value) > 32767:
                    raise ValueError(f"Excel cell text is too long: {name}/{number}/{key}")
                # Treat imported text as text, never as a formula from drawing annotations.
                if isinstance(value, str) and value.startswith("=") and generated is None:
                    cell.data_type = "s"
                cell.fill = PatternFill("solid", fgColor="F0F1F2" if generated else "EEF6FC")
                cell.font = Font(name="Microsoft YaHei", size=10, color="243746")
                cell.alignment = Alignment(vertical="center", wrap_text=key in {"note", "detail", "nodes"})
                if numeric(value) or generated:
                    cell.number_format = "0.###"
            ws.row_dimensions[number].height = 26
        for col, (key, label) in enumerate(schema, 1):
            width = 15
            if key in {"id", "member", "ni", "nj", "node", "category", "type"}:
                width = 21
            if key in {"note", "detail", "source", "nodes", "behavior"}:
                width = 48
            ws.column_dimensions[ws.cell(2, col).column_letter].width = width
            ws.cell(2, col).fill = PatternFill("solid", fgColor="254D61")
            ws.cell(2, col).font = Font(name="Microsoft YaHei", bold=True, color="FFFFFF", size=10)
            ws.cell(2, col).alignment = Alignment(wrap_text=True, vertical="center")
            if key in {"status", "keep", "kind"}:
                choices = {"status": "待核,待补边界,已确认,排除", "keep": "保留,排除", "kind": "外边界,洞口"}[key]
                validation = DataValidation(type="list", formula1='"' + choices + '"', allow_blank=True)
                ws.add_data_validation(validation)
                validation.add(f"{ws.cell(2, col).column_letter}3:{ws.cell(2, col).column_letter}10002")
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 38
        ws.cell(1, 1).font = Font(name="Microsoft YaHei", bold=True, size=13, color="254D61")
        ws.freeze_panes = "C3"
        ws.auto_filter.ref = f"A2:{ws.cell(ws.max_row, len(schema)).coordinate}"
        ws.sheet_view.zoomScale = 85
        ws.print_title_rows = "1:2"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    meta.sheet_state = "hidden"
    wb.calculation = CalcProperties(calcId=191029, fullCalcOnLoad=True)
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def read_book(path):
    wb = load_workbook(path, data_only=False)
    if "__schema" not in wb or wb["__schema"].cell(1, 2).value != 1:
        raise ValueError("Not a supported parameter workbook")
    sheets = []
    for name, role, encoded in wb["__schema"].iter_rows(min_row=2, values_only=True):
        keys = json.loads(encoded)
        ws = wb[name]
        expected = [label for _, label in SCHEMAS[role]]
        actual = [ws.cell(2, i).value for i in range(1, len(expected) + 1)]
        if actual != expected:
            raise ValueError(f"Headers changed in {name}; reconcile schema before reading")
        rows = []
        for number, values in enumerate(ws.iter_rows(min_row=3, max_col=len(keys), values_only=True), 3):
            row = dict(zip(keys, values))
            if not any(v not in (None, "") for k, v in row.items() if k not in DERIVED.get(role, set())):
                continue
            rows.append({**row, "_sheet": name, "_row": number, "_role": role})
        sheets.append({"name": name, "role": role, "rows": rows})
    settings = dict(wb["项目设置"].iter_rows(min_row=2, max_col=2, values_only=True))
    derived_nodes = []
    if "__topology" in wb.sheetnames:
        topo = wb["__topology"]
        topo_headers = [topo.cell(2, i).value for i in range(1, 8)]
        if topo_headers == ["节点编号", "X/mm", "Y/mm", "Z/mm", "节点类型", "核查状态", "说明"]:
            for values in topo.iter_rows(min_row=3, max_col=7, values_only=True):
                if values[0]:
                    derived_nodes.append(dict(zip(("node_id", "x", "y", "z", "kind", "status", "detail"), values)))
    result = {"project": wb["使用说明"].cell(1, 1).value, "settings": settings,
              "sheets": sheets, "_derived_nodes": derived_nodes}
    wb.close()
    return result


def records(book, role):
    return [r for s in book["sheets"] if s["role"] == role for r in s["rows"]]


def active(row):
    return row.get("keep") != "排除" and row.get("status") != "排除"


def xy_area(points):
    return sum(a[0] * b[1] - b[0] * a[1] for a, b in zip(points, points[1:] + points[:1])) / 2


def orientation(a, b, c):
    return (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])


def on_segment(p, a, b, tol=1e-6):
    length = math.dist(a[:2], b[:2])
    return (abs(orientation(a, b, p)) <= tol * max(length, 1)
            and min(a[0], b[0]) - tol <= p[0] <= max(a[0], b[0]) + tol
            and min(a[1], b[1]) - tol <= p[1] <= max(a[1], b[1]) + tol)


def intersects(a, b, c, d, proper=False):
    v = [orientation(a, b, c), orientation(a, b, d), orientation(c, d, a), orientation(c, d, b)]
    epsilon = 1e-7 * max(math.dist(a[:2], b[:2]), math.dist(c[:2], d[:2]), 1)
    signs = [0 if abs(value) <= epsilon else (1 if value > 0 else -1) for value in v]
    if signs[0] * signs[1] < 0 and signs[2] * signs[3] < 0:
        return True
    if proper:
        return False
    return not proper and (on_segment(c, a, b) or on_segment(d, a, b)
                           or on_segment(a, c, d) or on_segment(b, c, d))


def inside_polygon(point, points, strict=False):
    if any(on_segment(point, a, b) for a, b in zip(points, points[1:] + points[:1])):
        return not strict
    inside = False
    for a, b in zip(points, points[1:] + points[:1]):
        if (a[1] > point[1]) != (b[1] > point[1]):
            if point[0] < (b[0] - a[0]) * (point[1] - a[1]) / (b[1] - a[1]) + a[0]:
                inside = not inside
    return inside


def polygon_problem(points, tolerance):
    if len(points) < 3 or len(set(tuple(p) for p in points)) != len(points):
        return "边界至少需要三个不同顶点，末点不重复首点"
    if max(p[2] for p in points) - min(p[2] for p in points) > tolerance:
        return "边界不在同一水平标高"
    if abs(xy_area(points)) <= 1e-6:
        return "边界面积为零"
    n = len(points)
    for i in range(n):
        for j in range(i + 1, n):
            if j == i + 1 or (i == 0 and j == n - 1):
                continue
            if intersects(points[i], points[(i + 1) % n], points[j], points[(j + 1) % n]):
                return "边界存在自交或非相邻边接触"
    return None


def polygons_overlap(first, second):
    edges_a = list(zip(first, first[1:] + first[:1]))
    edges_b = list(zip(second, second[1:] + second[:1]))
    if any(intersects(a, b, c, d, proper=True) for a, b in edges_a for c, d in edges_b):
        return True
    if any(inside_polygon(p, second, strict=True) for p in first):
        return True
    if any(inside_polygon(p, first, strict=True) for p in second):
        return True
    # Interior probes also detect identical boundaries and overlapping collinear edges.
    for polygon, other in ((first, second), (second, first)):
        sign = 1 if xy_area(polygon) > 0 else -1
        for a, b in zip(polygon, polygon[1:] + polygon[:1]):
            length = math.dist(a[:2], b[:2])
            if length <= 1e-6:
                continue
            offset = min(length * 1e-4, .001)
            p = [(a[0] + b[0]) / 2 - sign * (b[1] - a[1]) / length * offset,
                 (a[1] + b[1]) / 2 + sign * (b[0] - a[0]) / length * offset]
            if inside_polygon(p, polygon, strict=True) and inside_polygon(p, other, strict=True):
                return True
    return False


def node_map(book, issues):
    tolerance = book["settings"].get("node_tolerance_mm", 0.01)
    if not numeric(tolerance) or tolerance < 0:
        raise ValueError("Invalid node_tolerance_mm")
    nodes, conflicts = {}, set()
    for role in ("members", "point", "pile", "boundary", "surface_node", "core_column"):
        for row in records(book, role):
            if not active(row):
                continue
            endpoints = [(row.get("node"), [row.get(k) for k in ("x", "y", "z")])]
            if role == "surface_node":
                endpoints = [(row.get("id"), [row.get(k) for k in ("x", "y", "z")])]
            if role == "core_column":
                endpoints = [(row.get("id"), [row.get(k) for k in ("x", "y", "z")])]
            if role == "members":
                endpoints = [(row.get("n" + e), [row.get(k + e) for k in ("x", "y", "z")]) for e in "ij"]
            for node, point in endpoints:
                if not node or not all(numeric(v) for v in point):
                    issue(issues, row, "error", "node_incomplete", "节点编号或XYZ缺失/非数值")
                    continue
                if node in nodes and math.dist(nodes[node], point) > tolerance:
                    conflicts.add(node)
                    issue(issues, row, "error", "node_conflict", f"节点{node}坐标冲突：{nodes[node]} / {point}")
                else:
                    nodes.setdefault(node, point)
    for row in book.get("_derived_nodes", []):
        node, point = row.get("node_id"), [row.get(k) for k in ("x", "y", "z")]
        if not node or not all(numeric(v) for v in point):
            continue
        if node in nodes and math.dist(nodes[node], point) > tolerance:
            conflicts.add(node)
            issue(issues, {**row, "_sheet": "__topology", "_row": None}, "error",
                  "derived_node_conflict", f"派生节点{node}坐标冲突：{nodes[node]} / {point}")
        else:
            nodes.setdefault(node, point)
    for node in conflicts:
        nodes.pop(node, None)
    return nodes, conflicts


def issue(issues, row, severity, code, message):
    issues.append({"severity": severity, "code": code, "sheet": row.get("_sheet"),
                   "row": row.get("_row"), "id": row.get("id", row.get("loop")), "message": message})


def calculate(book, row, key, categories, issues):
    value = row.get(key)
    if not isinstance(value, str) or not value.startswith("="):
        return value
    category_name = next((s["name"] for s in book["sheets"] if s["role"] == "category"), "荷载类别取值")
    if value != formula(row["_role"], key, row["_row"], category_name):
        issue(issues, row, "error", "unsupported_formula", f"{key}公式已改变，需核实后读取")
        return None
    if row["_role"] == "line":
        factors = [row.get(k) for k in ("density", "thickness", "height")]
        return math.prod(factors) if all(numeric(v) for v in factors) else None
    category = categories.get(row.get("category"), {})
    return category.get(key)


def analyze(book):
    issues = []
    nodes, conflicts = node_map(book, issues)
    tolerance = book["settings"].get("node_tolerance_mm", 0.01)
    allrows = [r for s in book["sheets"] for r in s["rows"] if active(r)]
    allrows.extend({**row, "id": row.get("node_id"), "_role": "surface_node", "_sheet": "__topology", "_row": i}
                   for i, row in enumerate(book.get("_derived_nodes", []), 1))
    seen = set()
    for row in allrows:
        role = row["_role"]
        if role not in {"trace", "boundary"}:
            namespace = "load" if role in {"point", "line", "surface"} else role
            identity = (namespace, row.get("id"))
            if not row.get("id") or identity in seen:
                issue(issues, row, "error", "duplicate_or_missing_id", "编号缺失或重复")
            seen.add(identity)
        if "status" in row and row["status"] != "已确认":
            issue(issues, row, "warning", "pending", "尚未人工确认")
    members = {r["id"]: r for r in records(book, "members") if active(r) and r.get("id")}
    categories = {r["id"]: r for r in records(book, "category") if r.get("id")}
    for row in allrows:
        role = row["_role"]
        required = {
            "members": ("section", "material", "ci", "cj"),
            "wall": ("material",), "point": ("case", "member"),
            "core_column_section": ("column", "x", "y", "z0", "z1", "section", "material"),
            "line": ("case", "direction", "density", "thickness", "height"),
            "pile": ("diameter", "system", "kx", "ky", "kz", "krx", "kry", "krz", "behavior"),
            "slab": ("material", "thickness", "z"),
        }.get(role, ())
        missing = [k for k in required if row.get(k) in (None, "")]
        if missing:
            issue(issues, row, "warning", "missing_parameters", "待补参数：" + ", ".join(missing))
        if role == "core_column_section":
            if all(numeric(row.get(k)) for k in ("z0", "z1")) and row["z1"] <= row["z0"]:
                issue(issues, row, "error", "column_interval", "柱区段上标高必须大于下标高")
        elif role == "members":
            if row.get("ni") in nodes and row.get("nj") in nodes:
                if math.dist(nodes[row["ni"]], nodes[row["nj"]]) <= tolerance:
                    issue(issues, row, "error", "zero_length", "构件长度为零或小于节点核对容差")
        elif role == "point":
            if not any(numeric(row.get(k)) for k in ("fx", "fy", "fz", "mx", "my", "mz")):
                issue(issues, row, "warning", "missing_load", "未填写任何力或力矩分量")
            if row.get("member") and row["member"] not in members:
                issue(issues, row, "error", "missing_member", "受力构件不存在或被排除")
            beam = members.get(row.get("member"))
            point = [row.get(k) for k in ("x", "y", "z")]
            if beam and beam.get("ni") in nodes and beam.get("nj") in nodes and all(numeric(v) for v in point):
                a, b = nodes[beam["ni"]], nodes[beam["nj"]]
                delta = [b[i] - a[i] for i in range(3)]
                length2 = sum(v * v for v in delta)
                if length2 > 0:
                    ratio = sum((point[i] - a[i]) * delta[i] for i in range(3)) / length2
                    distance = math.dist(point, [a[i] + max(0, min(1, ratio)) * delta[i] for i in range(3)])
                    if distance > tolerance:
                        issue(issues, row, "error" if row.get("status") == "已确认" else "warning",
                              "point_off_member", f"点位距指定构件{distance:.3f}mm，不自动投影")
        elif role == "line":
            row["_q"] = calculate(book, row, "q", categories, issues)
            if not numeric(row["_q"]):
                issue(issues, row, "warning", "missing_q", "容重/厚度/高度未齐，q未知")
            beam = members.get(row.get("member"))
            if beam is None:
                issue(issues, row, "error", "missing_member", "受荷构件不存在或被排除")
            elif beam.get("ni") in nodes and beam.get("nj") in nodes:
                length = math.dist(nodes[beam["ni"]], nodes[beam["nj"]]) / 1000
                if not (numeric(row.get("a")) and numeric(row.get("b"))
                        and -tolerance / 1000 <= row["a"] < row["b"] <= length + tolerance / 1000):
                    issue(issues, row, "error", "line_interval", "区间须满足0≤起始距离<终止距离≤构件长度")
        elif role == "surface":
            ids = nodes_list(row.get("nodes"))
            if not ids and row.get("status") == "待补边界":
                issue(issues, row, "warning", "surface_boundary_pending", "边界节点尚待补充，不按缺载或零荷载处理")
                for key in ("dead", "live"):
                    row["_" + key] = calculate(book, row, key, categories, issues)
                    if not numeric(row["_" + key]):
                        issue(issues, row, "warning", "missing_intensity", f"{key}强度未知，不按零处理")
                continue
            triangle_ids = nodes_list(book["settings"].get("triangle_surface_ids"))
            expected_count = 3 if row.get("id") in triangle_ids else 4
            if len(ids) != expected_count or len(set(ids)) != expected_count:
                issue(issues, row, "error", "surface_corners", f"荷载面需{expected_count}个不同的有序角点；直线闭合，不插点拟合曲边")
            elif any(n not in nodes for n in ids):
                issue(issues, row, "error", "surface_nodes", "角点须为已定义且无坐标冲突的节点")
            else:
                row["_points"] = [nodes[n] for n in ids]
                problem = polygon_problem(row["_points"], tolerance)
                if problem:
                    issue(issues, row, "error", "surface_polygon", problem)
                else:
                    row["_area_m2"] = abs(xy_area(row["_points"])) / 1e6
            if row.get("category") not in categories:
                issue(issues, row, "warning", "missing_category", "区域类别未在取值表定义")
            for key in ("dead", "live"):
                row["_" + key] = calculate(book, row, key, categories, issues)
                if not numeric(row["_" + key]):
                    issue(issues, row, "warning", "missing_intensity", f"{key}强度未知，不按零处理")
        elif role == "wall":
            if not all(numeric(row.get(k)) for k in ("z0", "z1", "r0", "r1", "t0", "t1")):
                issue(issues, row, "error", "wall_incomplete", "区段几何或法向厚度缺失/非数值")
            elif row["z1"] <= row["z0"] or min(row[k] for k in ("r0", "r1", "t0", "t1")) <= 0:
                issue(issues, row, "error", "wall_geometry", "区段标高顺序或尺寸无效")
        elif role == "pile":
            if any(numeric(row.get(k)) and row[k] < 0 for k in ("kx", "ky", "kz", "krx", "kry", "krz")):
                issue(issues, row, "error", "negative_stiffness", "等效线性支承刚度不能为负")
        elif role == "slab":
            if not numeric(row.get("thickness")) or row["thickness"] <= 0 or not numeric(row.get("z")):
                issue(issues, row, "error", "slab_parameters", "底板厚度须为正数，中面标高须为数值")
    surfaces = [r for r in allrows if r["_role"] == "surface" and "_area_m2" in r]
    for i, first in enumerate(surfaces):
        for second in surfaces[i + 1:]:
            if abs(first["_points"][0][2] - second["_points"][0][2]) > tolerance:
                continue
            if polygons_overlap(first["_points"], second["_points"]):
                issue(issues, second, "error", "surface_overlap", f'受荷范围与{first["id"]}重叠')
    slabs = {r["id"]: r for r in records(book, "slab") if active(r) and r.get("id")}
    loops = defaultdict(list)
    for row in records(book, "boundary"):
        if active(row):
            loops[(row.get("slab"), row.get("loop"))].append(row)
    valid_loops = defaultdict(list)
    for (slab_id, loop), rows in loops.items():
        row = rows[0]
        if slab_id not in slabs or not loop:
            issue(issues, row, "error", "boundary_reference", "底板不存在或轮廓编号缺失")
            continue
        orders = [r.get("order") for r in rows]
        kinds = {r.get("kind") for r in rows}
        if (not all(numeric(v) and v == int(v) and v > 0 for v in orders)
                or len(set(orders)) != len(orders) or len(kinds) != 1
                or next(iter(kinds)) not in {"外边界", "洞口"}):
            issue(issues, row, "error", "boundary_order", "顶点顺序须为不重复正整数，轮廓类型须一致")
            continue
        if any(r.get("node") not in nodes for r in rows):
            continue
        points = [nodes[r["node"]] for r in sorted(rows, key=lambda r: r["order"])]
        problem = polygon_problem(points, tolerance)
        if problem:
            issue(issues, row, "error", "boundary_polygon", problem)
            continue
        if numeric(slabs[slab_id].get("z")) and any(abs(p[2] - slabs[slab_id]["z"]) > tolerance for p in points):
            issue(issues, row, "error", "boundary_elevation", "轮廓标高与底板中面标高不一致")
        valid_loops[slab_id].append((row, points))
    for slab_id, slab in slabs.items():
        outer = [(r, p) for r, p in valid_loops[slab_id] if r["kind"] == "外边界"]
        holes = [(r, p) for r, p in valid_loops[slab_id] if r["kind"] == "洞口"]
        if len(outer) != 1:
            issue(issues, slab, "error", "slab_outer", "每个底板需一个完整外边界")
            continue
        for row, hole in holes:
            crossing = any(intersects(a, b, c, d) for a, b in zip(hole, hole[1:] + hole[:1])
                           for c, d in zip(outer[0][1], outer[0][1][1:] + outer[0][1][:1]))
            if crossing or not all(inside_polygon(p, outer[0][1], strict=True) for p in hole):
                issue(issues, row, "error", "slab_hole", "洞口须完整位于底板外边界内")
        for i, (row, hole) in enumerate(holes):
            for other_row, other in holes[i + 1:]:
                if polygons_overlap(hole, other):
                    issue(issues, other_row, "error", "slab_hole_overlap", "洞口相互重叠")
    return {"issues": issues, "nodes": nodes, "conflicts": sorted(conflicts),
            "members": members, "categories": categories}


def wall_section_at(book, level):
    """Return the saved wall's radii in mm, taking the lower segment at a joint."""
    z = level * 1000
    walls = [r for r in records(book, "wall") if active(r)
             and all(numeric(r.get(k)) for k in ("z0", "z1", "r0", "r1", "t0", "t1"))
             and r["z1"] > r["z0"]]
    base = min((r["z0"] for r in walls), default=None)
    matches = [r for r in walls if r["z0"] < z <= r["z1"]
               or z == base == r["z0"]]
    if not matches:
        return None
    if len(matches) != 1:
        raise ValueError(f"Overlapping wall segments at {level:g} m")
    row = matches[0]
    fraction = (z - row["z0"]) / (row["z1"] - row["z0"])
    outer = row["r0"] + fraction * (row["r1"] - row["r0"])
    normal = row["t0"] + fraction * (row["t1"] - row["t0"])
    slope = (row["r1"] - row["r0"]) / (row["z1"] - row["z0"])
    inner = outer - normal * math.sqrt(1 + slope * slope)
    if not 0 < inner < outer:
        raise ValueError(f"Invalid wall radii at {level:g} m")
    return {"outer": outer, "inner": inner, "z0": row["z0"], "z1": row["z1"]}


def review(book, result, output, font_path=None):
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A3, landscape

    if Path(output).exists():
        raise FileExistsError(output)
    if result["conflicts"]:
        raise ValueError("Resolve node coordinate conflicts before drawing: " + ", ".join(result["conflicts"]))
    default_font = Path("C:/Windows/Fonts/simhei.ttf")
    selected = Path(font_path) if font_path else default_font
    if selected.is_file():
        pdfmetrics.registerFont(TTFont("ReviewCJK", str(selected)))
        font = "ReviewCJK"
    else:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        font = "STSong-Light"
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output), pagesize=landscape(A3))
    c.setTitle(str(book["project"]) + " 参数核对图")
    width, height = landscape(A3)
    nodes, members = result["nodes"], result["members"]
    pages = 0

    def member_color(member_id):
        # Stable per physical member ID, so labels and repeated review pages match.
        value = int.from_bytes(hashlib.sha256(str(member_id).encode("utf-8")).digest()[:4], "big")
        hue = value / 2**32
        red, green, blue = colorsys.hsv_to_rgb(hue, .72, .62)
        return red, green, blue

    def label(text, x, y, size=7, color="#344454"):
        c.setFont(font, size)
        c.setFillColor(color)
        c.drawString(x, y, str(text))

    def page(title, points, draw, wall=None):
        nonlocal pages
        if wall:
            radius = wall["outer"]
            points = points + [[-radius, -radius], [radius, radius]]
        if not points:
            return
        pages += 1
        label(book["project"], 35, height - 35, 13)
        label(title, 35, height - 63, 19)
        label("核查草案 | 坐标 mm | 每根梁使用独立颜色 | 节点圆点标记梁端 | 颜色不代表连接结论", 35, height - 83, 9)
        if wall:
            label(f'外壁：深蓝实线，外径 {2 * wall["outer"]:.3f} mm', 35, height - 101, 9, "#234C73")
            label(f'内壁：橙色虚线，内径 {2 * wall["inner"]:.3f} mm', 360, height - 101, 9, "#BD5C17")
            label(f'筒段 {wall["z0"] / 1000:g}–{wall["z1"] / 1000:g} m（分界标高取下段）', 695, height - 101, 9)
        xlo, xhi = min(p[0] for p in points), max(p[0] for p in points)
        ylo, yhi = min(p[1] for p in points), max(p[1] for p in points)
        scale = min((width - 110) / max(xhi - xlo, 1), (height - 170) / max(yhi - ylo, 1))
        def pt(p):
            return (width / 2 + (p[0] - (xlo + xhi) / 2) * scale,
                    (height - 120) / 2 + 25 + (p[1] - (ylo + yhi) / 2) * scale)
        c.setStrokeColorRGB(.82, .85, .87)
        c.setLineWidth(.5)
        for coordinate in range(math.ceil(xlo / 5000) * 5000, int(xhi) + 1, 5000):
            a, b = pt([coordinate, ylo]), pt([coordinate, yhi])
            c.line(*a, *b)
            label(str(coordinate), a[0], a[1] - 12, 7, "#81909B")
        for coordinate in range(math.ceil(ylo / 5000) * 5000, int(yhi) + 1, 5000):
            a, b = pt([xlo, coordinate]), pt([xhi, coordinate])
            c.line(*a, *b)
            label(str(coordinate), a[0] - 30, a[1], 7, "#81909B")
        occupied = []
        def tag(text, point, size=7, color="#344454"):
            x, y = pt(point)
            tw = pdfmetrics.stringWidth(str(text), font, size) + 4
            candidates = [(5, 5), (5, -11), (-tw - 5, 5), (-tw - 5, -11),
                          (12, 17), (12, -23), (-tw - 12, 17), (-tw - 12, -23)]
            for dx, dy in candidates:
                rect = (x + dx, y + dy - 2, x + dx + tw, y + dy + size + 2)
                if not any(rect[0] < b[2] and rect[2] > b[0] and rect[1] < b[3] and rect[3] > b[1] for b in occupied):
                    break
            occupied.append(rect)
            c.setFillColorRGB(1, 1, 1)
            c.rect(rect[0], rect[1], tw, size + 4, fill=1, stroke=0)
            label(text, rect[0] + 2, rect[1] + 2, size, color)
        draw(pt, tag)
        if wall:
            c.saveState()
            for key, color, dash in (("outer", "#234C73", []), ("inner", "#BD5C17", [6, 3])):
                c.setStrokeColor(color)
                c.setLineWidth(1.4)
                c.setDash(dash)
                c.circle(*pt([0, 0]), wall[key] * scale, stroke=1, fill=0)
            c.restoreState()
        label(f"第 {pages} 页 | 修改 Excel 后重新生成；待核参数见检查记录。", 35, 20, 9)
        c.showPage()

    floors = sorted({r["level"] for r in members.values() if numeric(r.get("level"))})
    for floor in floors:
        wall = wall_section_at(book, floor)
        beams = [r for r in members.values() if r.get("level") == floor
                 and r.get("ni") in nodes and r.get("nj") in nodes]
        points = [nodes[r[k]] for r in beams for k in ("ni", "nj")]
        def background(pt, tag, detailed=False):
            c.setLineWidth(1.2 if detailed else .7)
            for beam in beams:
                a, b = nodes[beam["ni"]], nodes[beam["nj"]]
                c.setStrokeColorRGB(*member_color(beam["id"])) if detailed else c.setStrokeColorRGB(.78, .81, .83)
                c.line(*pt(a), *pt(b))
                if detailed:
                    color = "#%02X%02X%02X" % tuple(round(v * 255) for v in member_color(beam["id"]))
                    tag(beam["id"], [(a[i] + b[i]) / 2 for i in range(3)], 7, color)
            if detailed:
                ids = dict.fromkeys(n for b in beams for n in (b["ni"], b["nj"]))
                for node in ids:
                    x, y = pt(nodes[node])
                    c.setFillColorRGB(.22, .25, .28)
                    c.circle(x, y, 2, stroke=0, fill=1)
                    tag(node, nodes[node], 6, "#A33C25")
        page(f"{floor:g} m 几何：构件号 / 共用节点号", points, lambda pt, tag: background(pt, tag, True), wall)
        loads = [r for r in records(book, "point") if active(r) and r.get("level") == floor
                 and all(numeric(r.get(k)) for k in ("x", "y", "z"))]
        if loads:
            def point_view(pt, tag):
                background(pt, tag)
                for row in loads:
                    position = [row[k] for k in ("x", "y", "z")]
                    x, y = pt(position)
                    c.setFillColorRGB(.75, .27, .18)
                    c.circle(x, y, 3, fill=1, stroke=0)
                    tag(f'{row["id"]} / {row.get("node", "?")}', position, 7, "#A33C25")
            page(f"{floor:g} m 点荷载位置：不同工况分别核查", points + [[r[k] for k in ("x", "y", "z")] for r in loads], point_view, wall)
        lines = [r for r in records(book, "line") if active(r) and r.get("level") == floor]
        if lines:
            def line_view(pt, tag):
                background(pt, tag)
                for row in lines:
                    beam = members.get(row.get("member"))
                    if not beam or beam.get("ni") not in nodes or beam.get("nj") not in nodes:
                        continue
                    if not all(numeric(row.get(k)) for k in ("a", "b")):
                        continue
                    a, b = nodes[beam["ni"]], nodes[beam["nj"]]
                    length = math.dist(a, b) / 1000
                    if length <= 0 or not (0 <= row["a"] < row["b"] <= length + 1e-5):
                        continue
                    p = [[a[i] + (b[i] - a[i]) * row[k] / length for i in range(3)] for k in ("a", "b")]
                    c.setLineWidth(3)
                    c.setStrokeColorRGB(.7, .43, .1)
                    c.line(*pt(p[0]), *pt(p[1]))
                    q = f'{row["_q"]:g}' if numeric(row.get("_q")) else "待定"
                    tag(f'{row["id"]} q={q}', [(p[0][i] + p[1][i]) / 2 for i in range(3)], 8)
            page(f"{floor:g} m 线荷载：作用区间 / q(kN/m)", points, line_view, wall)
        areas = [r for r in records(book, "surface") if active(r) and r.get("level") == floor and "_area_m2" in r]
        if areas:
            def area_view(pt, tag):
                background(pt, tag)
                for row in areas:
                    polygon = row["_points"]
                    path = c.beginPath()
                    path.moveTo(*pt(polygon[0]))
                    for vertex in polygon[1:]:
                        path.lineTo(*pt(vertex))
                    path.close()
                    c.setStrokeColorRGB(.1, .5, .4)
                    c.setLineWidth(1.5)
                    c.drawPath(path)
                    tag(row["id"], [sum(p[i] for p in polygon) / len(polygon) for i in range(3)], 9)
            page(f"{floor:g} m 面荷载：用户定义的区域边界", points + [p for r in areas for p in r["_points"]], area_view, wall)
    piles = [r for r in records(book, "pile") if active(r) and r.get("node") in nodes]
    boundaries = [r for r in records(book, "boundary") if active(r) and r.get("node") in nodes]
    if piles or boundaries:
        def foundation(pt, tag):
            loops = defaultdict(list)
            for row in boundaries:
                loops[(row["slab"], row["loop"], row["kind"])].append(row)
            for (_, _, kind), rows in loops.items():
                polygon = [nodes[r["node"]] for r in sorted(rows, key=lambda r: r["order"])]
                c.setStrokeColorRGB(.2, .4, .45) if kind == "外边界" else c.setStrokeColorRGB(.7, .25, .15)
                c.setLineWidth(1.2)
                for a, b in zip(polygon, polygon[1:] + polygon[:1]):
                    c.line(*pt(a), *pt(b))
            for row in piles:
                position = nodes[row["node"]]
                c.setStrokeColorRGB(.15, .4, .5)
                c.circle(*pt(position), 3, stroke=1, fill=0)
                tag(row["id"], position, 6)
        page("基础平面：桩位 / 底板外轮廓 / 洞口", [nodes[r["node"]] for r in piles + boundaries], foundation)
    if not pages:
        page("尚无完整几何可绘制；请先补齐构件或基础坐标。", [[0, 0, 0]], lambda pt, tag: None)
    c.save()
    return pages


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=["create", "check", "review"])
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--font", type=Path)
    args = parser.parse_args()
    if args.command == "create":
        create(read_json(args.input), args.output)
        print(json.dumps({"workbook": str(args.output)}, ensure_ascii=True))
        return
    book = read_book(args.input)
    result = analyze(book)
    if args.command == "check":
        counts = {severity: sum(r["severity"] == severity for r in result["issues"]) for severity in ("error", "warning")}
        report = {"status": "checked_draft" if any(counts.values()) else "checks_passed",
                  "source": str(args.input.resolve()), "counts": counts, "issues": result["issues"]}
        write_json(args.output, report)
        print(json.dumps(counts))
    else:
        pages = review(book, result, args.output, args.font)
        print(json.dumps({"pages": pages, "findings": len(result["issues"])}))


if __name__ == "__main__":
    main()
