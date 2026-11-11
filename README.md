# 吸热塔参数识别

用于 Codex 的结构工程参数整理 Skill。将 CAD 图纸中的吸热塔几何、基础和荷载
整理为可人工修改的 Excel 总表，并生成带编号、梁颜色和内外壁圆的 PDF 核对图。
输出用于建模前参数交接，可供不同结构软件使用。

## 参数来源

| 信息 | 来源 | 处理方式 |
| --- | --- | --- |
| 几何 | CAD 图纸 DWG/DXF | 识别轴线、标高、筒壁、梁柱、桩位及底板；核实连接和坐标 |
| 截面 | 用户直接指定，或用户指定的结构软件导出文件 | 对应到具体构件，记录真实尺寸、材料、来源及截面变化 |
| 荷载 | 用户指定的图纸、荷载表和说明 | 整理点、线、面荷载；保留工况、方向、强度及作用范围 |

截面栏显示真实尺寸，例如 `工字形 600x300x12x20 mm`、`箱形 450x450x20 mm`。
原软件截面名称仅用于追溯。模型几何不覆盖 CAD 识别几何；匹配不上时标为待核。

## 安装与调用

本仓库根目录就是 Skill 目录，`SKILL.md`、`scripts/`、`references/`并列。
可通过 CC Switch 按 GitHub 仓库安装，仓库内路径为 `.`，同步到 Codex；
也可将整个目录以 `tower-cad-parameters` 为名放入主机的 Skills 目录。
显示名称为“吸热塔参数识别”，调用名为 `$tower-cad-parameters`。

示例请求：

```text
使用 $tower-cad-parameters 识别我指定的吸热塔CAD图纸。
几何按CAD识别，截面按我指定的S2K文件读取，荷载按图纸说明整理。
先生成一个楼层的核对图，再按确认规则整理全塔Excel总表。
```

用户直接指定截面时，可以写“这些梁采用工字形600x300x12x20mm、Q355”。
不用软件导出文件也可整理参数。

## 文件结构

```text
SKILL.md
README.md
agents/openai.yaml
references/cad-workflow.md
references/section-sources.md
references/workbook-schema.md
scripts/scan_dxf.py
scripts/parameters.py
scripts/read_s2k_sections.py
scripts/read_core_columns.py
scripts/requirements.txt
examples/parameters.json
examples/model.s2k
```

## 依赖与运行

Python 3.10+，运行 `python -m pip install -r scripts/requirements.txt`。
DWG 需可用的 AutoCAD/Core Console 或经过验证的转换工具；附带脚本不直接解析DWG。
PDF 默认尝试 Windows 黑体，其他系统可通过 `review --font` 指定中文 TrueType 字体。

以下为仓库内虚构数据的演示，不包含实际项目资料。`model.s2k`是用于验证解析的
最小表片段，不是完整的分析模型，也不应作为工程设计输入。

```shell
python scripts/parameters.py create examples/parameters.json --output demo.xlsx
python scripts/parameters.py check demo.xlsx --output demo-check.json
python scripts/parameters.py review demo.xlsx --output demo.pdf
python scripts/read_s2k_sections.py demo.xlsx examples/model.s2k --output demo-sections.xlsx --report demo-sections.json
python scripts/read_core_columns.py demo-sections.xlsx examples/model.s2k --output demo-columns.xlsx --report demo-columns.json
```

演示预期：一根梁完整匹配两个模型子段，梁截面显示真实尺寸；一根柱展开为两个
截面区段，下段为钢管混凝土、上段为钢管。工程确认仍为待核。已有输出需换新文件名。

## 当前能力范围

- DXF 实体盘点、图块变换及指定图层/区域内的候选梁轴线提取。
- Excel 创建、参数检查及几何/点线面荷载/基础 PDF 核对图。
- S2K 直线框架截面指派、同轴分段覆盖及工字形/箱形尺寸读取。
- 核心筒柱连续截面区段、单根箱管与未旋转混凝土矩形组成的组合截面读取。

标注、荷载含义和真实连接仍需解释与核查。其他软件输出文件需先验证格式并
补充适配；当前没有通用二进制模型解析器。此 Skill 不执行有限元建模或分析。

## 上传 GitHub

1. 新建仓库，建议仓库名 `tower-cad-parameters`，描述填“吸热塔参数识别”。
2. 上传本目录的内容到仓库根目录，使 `SKILL.md` 位于根目录。
3. 自行选择仓库可见性和发布许可证；本包没有预设许可证。
4. 提交前核对文件，项目 DWG、S2K 模型、工作簿、生成报告、临时文件和缓存均不随包发布。
5. 上传后，在 CC Switch 中填写仓库地址和仓库内路径 `.`，用于后续安装或更新。

`examples/`只含虚构演示数据。`.gitignore`排除生成物，但不会替代对已追踪文件的检查。
