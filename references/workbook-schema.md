# Workbook and normalized input

The helper consumes UTF-8 JSON with a `project` string, optional `settings` object,
and a `sheets` list. Each sheet is `{ "name": ..., "role": ..., "rows": [...] }`.
Rows use the keys listed below. Omitted values stay blank. Unknown numeric data
must not be filled with zero. User-facing headers are Chinese with explicit units.
The hidden `__schema` sheet stores header-to-key mappings; source trace is separate
from editing columns.

```json
{
  "project": "example-tower",
  "settings": {"node_tolerance_mm": 0.01},
  "sheets": [{
    "name": "9-57米梁",
    "role": "members",
    "rows": [{
      "id": "9.000-L001", "level": 9, "type": "floor beam",
      "ni": "N001", "xi": 0, "yi": 0, "zi": 9000,
      "nj": "N002", "xj": 5000, "yj": 0, "zj": 9000,
      "keep": "保留", "status": "待核"
    }]
  }]
}
```

## Roles and row keys

All coordinates, radii, shell thicknesses, and pile diameters are mm. Elevation
filter columns named `level` are m. Point forces are kN and point moments kN*m.
Equivalent translational support stiffness is kN/m; rotational stiffness is
kN*m/rad. Wall-density inputs use kN/m^3, m thickness, and m height. Line intervals
are m from the beam's start node. Surface intensities are kN/m^2.

- `members`: `id, level, type, section, material, ni, xi, yi, zi, nj, xj, yj,
  zj, ci, cj, keep, status, note`. Use for beams, columns, chords, webs, braces,
  and outriggers; preserve their type and parent grouping in `type`/`note`.
- `wall`: `id, z0, z1, r0, r1, t0, t1, material, status, note`. Thickness is
  normal to the wall. For a constant-thickness straight cone, radial thickness
  divided by `sqrt(1+slope^2)` gives normal thickness. Variable-thickness and
  transition geometry require checking the actual surfaces before conversion;
  do not indiscriminately apply that constant parallel-surface formula.
- `core_column`: legacy location/top records using `id, level, x, y, z, section,
  status, note`; do not infer a missing bottom elevation from these alone.
- `core_column_section`: `id, column, x, y, z0, z1, section, material, status,
  note`. One continuous assigned section interval per row; XY/elevations are mm.
  `column` preserves the parent column ID. Merge adjacent S2K segments only when
  their assignments agree and their elevations are continuous. Preserve the
  original location/top records in trace when upgrading an existing core-column
  sheet. Composite section text states actual steel tube, concrete core and
  encasement dimensions, with constituent materials; raw SD shapes and positions
  remain in trace. General SD area and inertia are equivalent properties relative
  to its base material, not automatically physical geometry properties.
- `point`: `id, level, node, x, y, z, member, case, fx, fy, fz, mx, my, mz,
  status, note`. One row per action point per case, not a merged cross-case total.
- `line`: `id, level, member, a, b, category, density, thickness, height, q,
  case, direction, status, note`. `q` is generated as a blank-aware product of
  density, thickness, and height. No missing factor becomes zero.
- `surface`: `id, level, nodes, category, dead, live, status, note`. `nodes` is
  a comma-separated sequence of exactly four distinct, ordered, coplanar corner
  nodes. Join consecutive nodes and close the fourth to the first with straight
  edges. Do not repeat the first node at the end. No self intersections or holes.
  A curved wall-side boundary is replaced by the chord between its two corner
  nodes, with no extra vertices, subdivisions, or curve-fitting child faces.
  Compute the load area from this quadrilateral and retain the input intensity.
  Record any omitted curved strip and its qA difference when source geometry is
  available; never increase intensity or add faces to recover the old total.
  This convention does not change structural shell or foundation geometry. Category lookup
  distinguishes an empty cell from zero. Dead/live cells are formulas; changing
  the category or category table updates them.
- `category`: `id, dead, live, note`.
- `surface_node`: `id, x, y, z, member, status, note`. Supplemental corner nodes
  at documented beam-interior intersections when they are not existing member
  endpoints. Existing endpoints retain their member-row node IDs and coordinates.
  Supplemental nodes do not split or renumber physical members.
- `pile`: `id, node, x, y, z, diameter, system, kx, ky, kz, krx, kry, krz,
  behavior, status, note`. Do not infer unspecified rotational or nonlinear
  behavior from a missing value.
- `slab`: `id, material, thickness, z, status, note`. Z is the physical reference
  midsurface elevation, not an unnamed top/bottom elevation.
- `boundary`: `slab, loop, kind, order, node, x, y, z, status, note`. `kind` is
  `外边界` or `洞口`. Each loop is ordered and closes implicitly at its first
  vertex. Foundation slab openings are valid and independent of load polygons.
- `trace`: `id, source, sheet, row, handle, detail`. Store verbose JSON and raw
  text here, not in the editing tables. Split detail longer than Excel's 32767
  character cell limit before passing it in.

## Formula and revision behavior

Only line q and surface-category lookups are generated formulas. User-editable
inputs use a light blue background and computed cells a light gray background.
The helper does not run Excel's calculation engine. Checks and PDFs independently
calculate these supported formulas from the current saved inputs. A replacement
formula outside the generated forms is flagged instead of treated as its cached
or former result. Excel should calculate normally when the workbook opens.

`check` reports missing values, node coordinate conflicts, duplicate IDs, invalid
line intervals, surface geometry, and incomplete foundation definitions. Missing
materials, sections, cases, or unconfirmed rows mean the workbook is still a draft.
An internal node coordinate map is derived from member, point, pile, and boundary
rows. Users still edit coordinates in those rows, not in a separate node table.
Explicitly approved triangular regions are listed as comma-separated load IDs in
`triangle_surface_ids`; only those IDs accept three corners.

PDFs read the workbook and draw floor geometry with member/node labels, separate
point/line/area views when records exist, and foundation contours/pile locations.
Conflicting node coordinates are never silently resolved to draw a load area.
The helper refuses a PDF if a referenced shared node has inconsistent coordinates.

Surface rows use beam-axis intersections with the verified wall inner boundary.
Member rows and line-load geometry use the corresponding confirmed outer-surface
intersections. Keep any computed outer-node mapping in a machine manifest and do
not append it as extra manual node rows. At a story elevation select the section
immediately below that elevation: at 9 m select the 0-9 m section. Retain source
CAD radii and signed differences as evidence, without substituting them for the
verified wall profile.
