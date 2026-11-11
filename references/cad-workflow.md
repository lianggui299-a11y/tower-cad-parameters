# Drawing extraction and project settings

## Intake

Inspect actual filenames, revisions, DWG units, layers, blocks, annotations, and
existing user edits. A historic handoff may describe an earlier state. Do not
assume the newest filename or a model-derived spreadsheet is authoritative for
all engineering parameters.

Ask for the missing origin or drawing scale when evidence cannot determine it.
Record a drawing-local origin for each plan, an XY rotation, a drawing-unit to mm
scale, its elevation, and the selected source region. Do not assume a circle is
the origin merely because the current project's plans use one.

## DWG

Use an installed, verified AutoCAD/Core Console or licensed DWG reader. Discover
the executable rather than hardcoding a different user's installation. Perform
conversion in a working copy with a new output. A real DXF conversion must retain
coordinates, text, layers, and block transforms. Compare those items with the
original and record the converter. Renaming a file extension is not conversion.
If no converter is available, identify that dependency; do not manufacture CAD
geometry from a screenshot or a similarly named old DXF.

## DXF inventory and candidate LINE extraction

`scan_dxf.py` reads modelspace, expands INSERT transformations, and inventories
LINE, text, circle, arc, and polyline evidence with source handles and block paths.
It deliberately does not explode curved boundaries into guessed straight beams.
Attributes and unsupported geometry remain evidence requiring inspection.

Without `--config`, it produces an entity inventory and layer/type counts. With
a config, candidate members require an explicit layer list, a plan source region,
unit scale, and coordinate transform. Only complete LINEs inside the source region
are candidates. Crossing-boundary LINEs are listed separately, never clipped into
new physical members. Do not automatically merge collinear lines or split beams
at intersections.

Example configuration (coordinates below are illustrative, not tower defaults):

```json
{
  "project": "example-tower",
  "plans": [{
    "key": "F09",
    "level_m": 9,
    "origin": [100000, 200000],
    "rotation_deg": 0,
    "scale_to_mm": 1,
    "bbox": [85000, 185000, 115000, 215000],
    "line_layers": ["CONFIRMED_BEAM_CENTERLINES"]
  }]
}
```

Inspect both text and the proposed lines before assigning member types,
end conditions, or actual common nodes. A geometric endpoint coincidence is not
proof that members at that location connect. Preserve the original handle/block
path mapping when assigning the final user-facing IDs. Keep published IDs stable
on revisions, matching source identity and showing additions/deletions rather
than renumbering every member.

## Section sources

Geometry is identified from CAD. Sections come from explicit user assignments or
user-selected software exports. Software coordinates are matching evidence, not
replacement CAD geometry. Read [section-sources.md](section-sources.md) for manual
assignments, software format inspection, S2K helpers, and composite sections.

## Load interpretation

- Associate a load table with its actual elevation, marker, leaders, and support
  count. Preserve raw signed text and units. Do not collapse plus/minus choices
  or redistribute a total to supports without a project-confirmed rule.
- Record the actual receiving member and location. Do not shift a point to the
  nearest beam merely because it is nearby.
- For wall line loads, retain the actual density, thickness, and height inputs.
  Confirm the load interval follows the chosen member orientation.
- Confirm category dead/live loads, self-weight inclusion, and action directions.
  Distinguish unknown, not applicable, and a specified numerical zero.
- Four user-entered load-area corner nodes define one straight-edged quadrilateral.
  At a curved tower edge, connect the two corner nodes by a chord rather than
  inserting vertices or creating small curve-fitting faces. Internal beams do not
  subdivide this load region. Check planarity, ordered boundary, area, and overlaps.
  Use the quadrilateral area with the specified intensity; report the omitted
  curved strip and resulting load difference if source geometry permits computing
  them. Do not compensate automatically. Structural shells retain their actual
  geometry, and structural slab holes remain separate contours.

## Shell-reference connection

The load-area corner uses the beam-axis intersection with the verified inner-wall
profile, selecting the lower segment at a story boundary. Member records use the corresponding
beam-axis intersection with the outer-surface reference. Generate/share the outer
node after the confirmed wall profile is selected. Record source and model radii
separately when they differ; do not silently substitute one for the other.

## Iteration

User modifications remain the maintained workbook data. A changed CAD input
creates a candidate revision with a difference list. Keep the current workbook
intact until revisions have been reconciled. Do not run a historical whole-book
generation script over the maintained workbook.

## Wall-profile reconciliation

Before computing an endpoint at a wall, compare the source plan's inner/outer
circles with the selected model wall profile at the same elevation. At a story
boundary use the lower story segment's inner and outer diameters unless the user
specifies otherwise; keep both inside radii when the wall thickness steps. Convert
radial thickness to normal thickness only after identifying the segment slope.
Record CAD and model radii, signed differences, source, and selected basis. A
large mismatch remains a review finding; do not silently relocate recognized
members or load boundaries to make the figures agree.
