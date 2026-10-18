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
    from GetPartBox, wizard hole position count, annotation counts).

## Verified on the real SOLIDWORKS

sketch (+ design-intent dims, symmetry, 4X Ø prefix), extrude/cut, fillet,
linear pattern, materials, equations/global variables, configurations,
measure, make_drawing + PDF, export, export_folder, hole via Hole Wizard
(counterbore, tapped with cosmetic thread, clearance without near-side
countersink — `countersink diameter 0` variant worked), base flange via
feature definition (one direction), sketch3d + tube sweep (outer
diameter), engrave.

Macro recordings and argument tables: `docs/recordings/HANDOFF.md`.

## Open items (as of commit after c3f763e)

1. Retest drawing: hole callouts after the GetOutline fix
   (`hole_callouts`, `callout_problems`), and why model dimensions stopped
   importing (`dimensions_imported` is now reported; was 0 on bracket_hw).
2. `open_document` on an already-open file now activates it — retest.
3. Hole position dimensions in drawings are missing (wizard placement
   sketch points have no dimensions).
4. Unverified on real SOLIDWORKS: revolve, chamfer, circular_pattern,
   mirror, shell, loft (once ok), rib, assemblies (insert_component,
   add_mate, check_interference, bom), set_color, export_flat_pattern
   options, Hole Wizard sizes other than the recorded ones (M8 worked).
