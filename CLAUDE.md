# solidworks-mcp — notes for Claude

MCP server (Python, pywin32 late binding) that drives SOLIDWORKS 2025 SP2
(rev 33.2, **Korean UI**) through its COM API. `server.py` = tools,
`sw.py` = COM plumbing. User-facing docs: README.md (Korean). Talk to the
user in Korean.

## How work has been done so far

- Developed in a cloud session without Windows/SOLIDWORKS. Every change was
  checked with the fake-COM scripts in `tests/fake_com/` (`run_all.py`), then
  the user ran a test prompt in **Claude Desktop** on their PC (MCP
  registered from `C:\mcp\solidworks-mcp`) and pasted back results and error
  text. Treat "passes fakes" as unverified until a real run confirms it.
- On this PC you can run the server against the real SOLIDWORKS directly
  (e.g. `python -c "import server; ..."` or call tool functions via
  `server.<tool>.__wrapped__` inside `sw.call`). Close test documents
  without saving; the user keeps real files in `Desktop\sw_test`.
- Commit style: imperative subject, body explains the observed failure and
  the fix. Push to `main`.
- Install: `install.cmd` → `install.ps1` (venv, pip, registration, check);
  `doctor.py` checks the setup and merges the Claude Desktop config
  (`--register` / `--unregister`, backups, Store-app path too). Test the
  installer on a copy of the repo with APPDATA/LOCALAPPDATA pointed at a
  scratch folder, never against the real config. `.ps1` files must stay
  UTF-8 *with* BOM (Windows PowerShell 5.1 misreads Korean otherwise) and
  CRLF (`.gitattributes`).

## Hard-won rules (each one was a real bug)

1. **Parameterless COM methods** (GetTitle, FirstFeature, GetOutline, …)
   often arrive as their *value* under late binding. Always call them via
   `sw.m0(obj, "Name")`; methods whose args are all optional (IMeasure
   .Calculate, EvaluateAll) via `sw.invoke`. Never write `obj.Method()`.
2. Indexed property writes (VBA `obj.Prop(i) = v`) need `sw.put_indexed`.
3. ByRef out-params: `sw.out_int()` / `sw.out_bstr()`, read `.value`.
4. Signatures differ between releases → `_call_first(obj, [(name, args), …])`
   tries known variants and only moves on for signature errors.
5. Enum values: read them with `sw.constant("swXxx")` (from swconst.tlb),
   don't hard-code guesses.
6. Default planes are localized (정면/윗면/우측면): use `sw.standard_plane`.
   Select planes by name with `SelectByID2(name, "PLANE")`; Feature.Select2
   was ignored by InsertSketch on new parts. `sketch()` verifies the
   sketch normal and redraws once if wrong.
7. `SelectByID2` with coordinates picks through the current view (misses
   hidden edges). In parts use `sw.select_at` / `sw.nearest` (body search).
8. InsertSketch / Insert3DSketch **toggle**: call `sw.close_open_sketch`
   first and check `ActiveSketch` after opening.
9. `dim.SystemValue = x` changes all configurations; use SetSystemValue3
   (done in set_dimension). Same for suppression (SetSuppression2).
10. One COM thread; `sw.call` times out (SW_TIMEOUT, 90 s) and rejects new
    work while a stuck call runs — usually a modal dialog in SOLIDWORKS.
11. API units are metres/radians; tools take mm/degrees (`mm()`, `deg()`).
12. Success must be checked, not assumed (export file mtime, built length
    from the body box, wizard hole position count, annotation counts).
    Some setters return nothing even when they work
    (Face.SetMaterialPropertyValues2): read the value back instead.
13. **Selection marks** decide what an API call uses, and they are rarely
    documented. Found by trying: angle plane = plane mark 0 + axis mark 1;
    InsertDome = face mark 1; miter flange = profile sketch mark 0 + edges
    mark 1; hole table = origin vertex mark 1 + face mark 2; component
    pattern = definition created first, then components mark 1 + direction
    mark 2. Pass marks with `CreateSelectData` + `Select4`.
14. MathUtility.CreatePoint / MathPoint.MultiplyTransform must go through
    `sw.invoke` (late binding reports "member not found" otherwise).
15. While a drawing view is active, SketchManager coordinates are the
    view's own sketch space (relative to the view, unscaled), not sheet
    coordinates: map model points with `_to_sheet(view, p, sketch=True)`.
16. Standard model view names are localized ("*등각 보기"): take them from
    `GetModelViewNames` by position (`_model_view_names`).
17. InsertCombineFeature takes swBodyOperationType_e (SWBODYADD/CUT/
    INTERSECT); add/common want every body as a tool and no main body.
    SketchUseEdge3's second argument means "inner loops only".
    Tolerance.SetValues2 returns False here; SetValues works.
18. With Windows' UTF-8 code page (ACP 65001) SOLIDWORKS' PDF export drops
    every text run containing Korean; drawing PDFs are printed to
    "Microsoft Print to PDF" instead (`export`, pdf_method).
19. Weldments and annotations, found by trying: InsertWeldmentTrimFeature(4)
    with members mark 1 and boundary bodies mark 2 trims (other end
    conditions and the ...2 variant make a feature that trims nothing —
    check the volume); InsertGussetFeature2's BIsProfile True = polygon
    (needs both chamfer legs > 0), False = triangle; InsertWeldmentTable
    wants configuration "" (the view's "기본<가공 상태>" gives None) while
    InsertBomTable4 needs the real name; InsertDatumTag2/InsertGtol are
    parameterless (`sw.invoke`); IGtol.SetFrameSymbols takes swGcs/swMc
    shorts (SetFrameSymbols2 wants text codes like "<GTOL-POSI>"); the
    roughness value of a surface finish symbol is text slot
    swSFSymbolRoughnessValue1; the Korean revision template has a title
    row above the column headings and calls the description "비고".
    Table positions are their anchor corner (`_table_box`).
20. Drawing views' visible *vertices* report GetPoint shifted by the view's
    sheet Position (edges and circle centres do not): `_vertex_point`.
    IView.InsertBreak3 takes positions relative to the view centre in
    sheet metres, with the view activated (not selected), then BreakView.
    Variable fillet = FeatureFillet3 type 1 with per-vertex radii; face
    fillet = type 2 with both faces mark 0. Weld symbols are text codes
    "<WELD-FILL>" etc. from gtol.sym (SetText's last argument is a long).
    MirrorComponents3 with ComponentsToInstance mirrors instances only.
    Sketched bends work via CreateDefinition(swFmSketchBend); SetFixedFace
    rejects every argument form, so SOLIDWORKS keeps the origin side, and a
    bend on an already bent flange is refused. IModelDoc2.
    InsertSheetMetalJog created nothing in 48 argument/selection variants.
21. IDimension.GetType is swDimensionParamType_e: 0 linear, 1 angular,
    2 integer (pattern counts). It used to be read as "2 = angular", so
    angles came out as mm and counts as degrees, in list_dimensions and
    set_dimension. Automatic annotation positions stay 20 mm inside the
    sheet edges (`_inside_border`) and stack per side of the view.
22. Grey dimensions/callouts in drawing PDFs are SOLIDWORKS' colour for
    driven (reference) annotations (swSystemColorsImportedDrivenAnnotation
    = 0x808080). Printing with PageSetup.DrawingColor = black and white
    removes them from grey but also drops every model edge from the PDF,
    so drawing PDFs stay in colour; SaveAs with swPDFExportInColor off is
    black but loses Korean text (rule 18).

## Verified on the real SOLIDWORKS

Run on this PC (SOLIDWORKS 2025, Korean UI) through the server's own
functions, each result checked (volumes, positions, PDFs rendered):

- Parts: sketch (+ design-intent dims), extrude/cut (blind, through, up to
  surface/plane, offset from surface, thin), fillet, linear pattern,
  draft, dome, cosmetic_thread, angle_plane, sketch_from_face (convert,
  offset), sketch_corners, combine_bodies (add/subtract/common),
  move_body (move, copy), face_color, import_file (STEP), materials,
  equations, configurations, measure, model_info (body box).
- Holes: Hole Wizard counterbore, tapped (cosmetic thread), clearance,
  all without the near-side countersink.
- Sheet metal: base flange, edge_flange (90° and 60°), miter_flange along
  two edges, hem.
- Weldments: structural_member (iso square tube, mitre corners),
  trim_members, end_cap, gusset (triangle, polygon), cut_list.
- More parts/sheet metal/assemblies: variable_fillet, face_fillet,
  sketched_bend, mirror_components.
- Assemblies: insert_component, add_mate (coincident, distance; faces
  found through component bodies), list_mates, set_component,
  component_pattern (linear, circular), exploded_view, replace_component.
- Drawings: make_drawing with model dimensions and hole callouts,
  drawing_view (isometric, exploded), section_view, detail_view,
  projected_view, set_drawing_scale, drawing_note, center_marks,
  add_dimension, set_tolerance, hole_table, title_block ("a3 - iso"),
  bom_table, auto_balloons, drawing_sheets, datum, geometric_tolerance,
  surface_finish, revision_table, broken_out_section, auxiliary_view,
  cut_list_table, weld_symbol, ordinate_dimensions, break_view, export
  PDF (all sheets, Korean text) and DXF.

Macro recordings and argument tables: `docs/recordings/HANDOFF.md`.

## Open items

1. Hole position dimensions in drawings are missing (wizard placement
   sketch points have no dimensions); add_dimension with entity "center"
   is the workaround.
2. Weldments: the cut list table's description column ("비고") stays empty
   with the Korean template (the cut list property is "설명").
3. Not rechecked on the real SOLIDWORKS: revolve, chamfer,
   circular_pattern, mirror, shell, loft, rib, check_interference,
   export_flat_pattern options, Hole Wizard sizes other than M4/M6/M8.
4. Sheet metal jog: no working API path found (see rule 20); a second
   sketched bend on a bent flange fails for the same fixed-face reason.
5. mirror_components makes instances only, not opposite-hand parts.
