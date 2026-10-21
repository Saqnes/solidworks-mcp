"""Regression run against the real SOLIDWORKS on this PC.

    .venv\\Scripts\\python tests\\live\\run_live.py            all groups
    .venv\\Scripts\\python tests\\live\\run_live.py parts holes   only some groups

Groups: parts, holes, sheetmetal, weldment, assembly, drawing.

Builds small test models in a temporary folder through the server's own
tool functions and compares what comes out with hand-computed numbers
(volumes, lengths, counts). Prints one PASS/FAIL line per check and a
summary; the exit code is the number of failures. It opens and closes only
its own documents (anything already open in SOLIDWORKS is left alone) and
writes nothing outside the temporary folder (set SWMCP_KEEP=1 to keep it
for a look). Takes a few minutes.
"""

from __future__ import annotations

import math
import os
import shutil
import sys
import tempfile
import time
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

import server as S  # noqa: E402
import sw  # noqa: E402

WORK = tempfile.mkdtemp(prefix="swmcp_live_")
OPENED: list[str] = []
RESULTS: list[tuple[str, bool, str]] = []


def path(name: str) -> str:
    # a prefix keeps titles unique: SOLIDWORKS refuses a second open file with the same title
    return os.path.join(WORK, f"live_{name}")


def check(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, ok, detail))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail else ""), flush=True)


def near(name: str, got: float, want: float, tol: float = 0.05) -> None:
    check(name, got is not None and abs(got - want) <= tol, f"got {got:.3f}, want {want:.3f}" if got is not None else "no value")


def new(kind: str = "part") -> dict:
    info = S.new_document(kind)
    OPENED.append(info["title"])
    return info


def opened(info: dict) -> dict:
    OPENED.append(info["title"])
    return info


def volume() -> float:
    return S.model_info()["volume_mm3"]


def close_mine() -> None:
    """Close what this run opened: documents created here (by their title at
    creation, or their drawings) and anything saved into the work folder."""
    def work():
        app = sw.app()
        # newest first (drawings and assemblies before the parts they hold open), a few passes;
        # closing a document can invalidate the other handles, so re-list every pass
        for _ in range(4):
            mine = []
            for d in reversed(list(sw.m0(app, "GetDocuments") or ())):
                try:
                    title, where = sw.m0(d, "GetTitle"), sw.m0(d, "GetPathName") or ""
                except Exception:
                    continue
                if os.path.normcase(where).startswith(os.path.normcase(WORK)) or any(
                        title == t or title.startswith(f"{t} - ") for t in OPENED):
                    mine.append(title)
            if not mine:
                return
            for title in mine:
                app.CloseDoc(title)
    sw.call(work)
    OPENED.clear()


# --- groups --------------------------------------------------------------------


def parts() -> None:
    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 60, "height": 40}], plane="top")
    feat = S.extrude(depth=30)["feature"]
    near("extrude box", volume(), 72000)
    S.rename_feature(feat, "base block")
    check("rename feature", "base block" in [f["name"] for f in S.list_features()])
    S.shell(2, [[0, 30, 0]])
    near("shell open top", volume(), 72000 - 56 * 36 * 28)

    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 40, "height": 40}], plane="top")
    S.extrude(depth=20)
    S.chamfer(3, [[0, 20, 20]])
    near("chamfer 3x45", volume(), 32000 - 4.5 * 40)
    S.fillet(5, [[0, 20, -20]])
    near("fillet R5", volume(), 32000 - 4.5 * 40 - (25 - math.pi * 25 / 4) * 40, 0.5)

    new()
    S.sketch([{"type": "centerline", "x1": 0, "y1": -10, "x2": 0, "y2": 50},
              {"type": "rectangle", "x1": 10, "y1": 0, "x2": 20, "y2": 30}], plane="front", dimension=False)
    S.revolve(360)
    near("revolve ring", volume(), math.pi * (400 - 100) * 30, 0.5)

    new()
    S.sketch([{"type": "circle", "cx": 0, "cy": 0, "r": 40}], plane="top")
    S.extrude(depth=10)
    S.sketch([{"type": "circle", "cx": 25, "cy": 0, "r": 3}], face_point=[0, 10, 0])
    cut = S.extrude(depth=10, cut=True, end="through_all")["feature"]
    S.circular_pattern([cut], 6, "y")
    near("circular pattern 6 holes", volume(), math.pi * 1600 * 10 - 6 * math.pi * 9 * 10, 0.5)

    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 40, "height": 40}], plane="top")
    S.extrude(depth=10)
    S.sketch([{"type": "circle", "cx": 12, "cy": 0, "r": 4}], face_point=[0, 10, 0])
    boss = S.extrude(depth=10)["feature"]
    S.mirror([boss], "right")
    near("mirror boss", volume(), 16000 + 2 * math.pi * 16 * 10, 0.5)

    new()
    S.sketch([{"type": "rectangle", "x1": -30, "y1": -20, "x2": 30, "y2": 20}], plane="top")
    S.extrude(depth=5)
    S.sketch([{"type": "rectangle", "x1": -30, "y1": -20, "x2": -25, "y2": 20}], face_point=[-27, 5, 0])
    S.extrude(depth=35)
    S.sketch([{"type": "line", "x1": -25, "y1": 30, "x2": 15, "y2": 5}], plane="front", dimension=False)
    near("rib (25 x 40 triangle, 4 thick)", S.rib(4)["volume_added_mm3"], 2000, 0.5)

    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 40, "height": 40}], plane="top")
    plane = S.offset_plane(30, "top")["feature"]
    S.sketch([{"type": "circle", "cx": 0, "cy": 0, "r": 10}], plane=plane)
    sketches = [f["name"] for f in S.list_features() if f["type"] == "ProfileFeature"]
    S.loft(sketches[-2:])
    bb = S.model_info()["bounding_box_mm"]
    near("loft height", bb["max"][1] - bb["min"][1], 30, 0.01)

    new()
    S.sketch([{"type": "rectangle", "x1": -30, "y1": -20, "x2": 30, "y2": 20}], plane="top")
    S.sketch_corners([[-30, -20], [30, 20]], radius=5)
    S.extrude(depth=10)
    near("sketch corners R5", volume(), 24000 - 2 * (25 - math.pi * 25 / 4) * 10, 0.5)
    S.sketch([{"type": "circle", "cx": 30, "cy": 0, "r": 8}], plane="top")
    S.extrude(depth=30, merge=False)
    check("two bodies", len(S.list_bodies()) == 2)
    S.combine_bodies("subtract", [-20, 5, 20], [[38, 20, 0]])
    check("combine subtract", len(S.list_bodies()) == 1)

    new()
    h = S.helix(pitch=10, revolutions=5, diameter=20)
    S.sweep(h["feature"], circle_diameter=2)
    near("spring (helix sweep) volume", volume(), math.pi * 1 * h["length_mm"], 0.5)
    cs = S.coordinate_system([10, 20, 30], [90, 0, 0])
    check("coordinate system origin and turned y axis", cs["origin_mm"] == [10.0, 20.0, 30.0] and cs["y_axis"] == [0.0, 0.0, 1.0], str(cs))

    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 60, "height": 40}], plane="top")
    S.extrude(depth=5)
    S.sketch([{"type": "circle", "cx": -20, "cy": 0, "r": 3}], face_point=[0, 5, 0])
    cut = S.extrude(depth=5, cut=True, end="through_all")["feature"]
    pts = S.sketch([{"type": "point", "x": 0, "y": 0}, {"type": "point", "x": 20, "y": 10}],
                   face_point=[0, 5, 0], dimension=False)["sketch"]
    near("sketch pattern: two more holes", S.sketch_pattern([cut], pts)["volume_change_mm3"], -2 * math.pi * 9 * 5, 0.05)
    rp = S.reference_point([-17, 5, 0], entity="edge")
    check("reference point at the hole centre", rp["point_mm"] == [-20.0, 5.0, 0.0], str(rp))

    # one run_steps call: plate, hole on the top face found by list_faces, fillet on an edge from list_edges
    r = S.run_steps([
        {"tool": "new_document", "args": {"kind": "part"}},
        {"tool": "sketch", "args": {"entities": [{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 60, "height": 40}], "plane": "top"}},
        {"tool": "extrude", "args": {"depth": 10}},
        {"tool": "list_faces", "args": {"type": "plane"}},
        {"tool": "sketch", "args": {"entities": [{"type": "circle", "cx": 15, "cy": 0, "r": 5}], "face_point": "$3.faces.4.point_mm"}},
        {"tool": "extrude", "args": {"depth": 10, "cut": True, "end": "through_all"}},
        {"tool": "fillet", "args": {"radius": 5, "edge_points": [[30, 5, 20]]}},
    ])
    OPENED.append(r["results"][0]["title"])
    check("run_steps ran every step", r["completed"] == r["of"] == 7, str(r.get("failed")))
    near("run_steps model volume", volume(), 24000 - math.pi * 25 * 10 - (25 - math.pi * 25 / 4) * 10, 0.5)
    cyl = S.list_faces(type="cylinder")["faces"]
    hole = [f for f in cyl if f["radius_mm"] == 5.0 and f["axis_point_mm"][0] == 15.0]
    check("list_faces finds the hole (r5 at x 15)", len(hole) == 1, str(cyl))
    top = [f for f in S.list_faces(type="plane")["faces"] if f["normal"] == [0.0, 1.0, 0.0]]
    check("top face point lies on the face, off the hole",
          len(top) == 1 and top[0]["point_mm"][1] == 10.0 and math.dist(top[0]["point_mm"][::2], [15, 0]) > 5, str(top))
    circles = S.list_edges(type="circle")["edges"]
    check("list_edges: two hole circles and two fillet arcs", sorted(e["radius_mm"] for e in circles) == [5.0, 5.0, 5.0, 5.0], str(circles))
    fil = [f["name"] for f in S.list_features() if f["type"] == "Fillet"][-1]
    broken = S.set_dimension(f"D1@{fil}", 100)
    check("a fillet too big is reported as a rebuild error", any(e["feature"] == fil for e in broken.get("rebuild_errors", [])), str(broken))
    check("fixed again: no rebuild errors", "rebuild_errors" not in S.set_dimension(f"D1@{fil}", 5))
    units = S.document_units(decimals=0)
    check("document units: whole mm", units["decimals"] == 0 and units["dimension_decimals"] == 0, str(units))
    S.save(path("lefthand.SLDPRT"))
    com = S.model_info()["center_of_mass_mm"]
    m = S.mirror_part()
    OPENED.append(m["part"])
    check("mirror_part flips the centre of mass in x", abs(m["center_of_mass_mm"][0] + com[0]) < 0.01 and abs(com[0]) > 0.1,
          f"{com} -> {m['center_of_mass_mm']}")


def holes() -> None:
    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 200, "height": 120}], plane="top")
    S.extrude(depth=20)
    for kind, size, x in (("clearance", "M8", -70), ("tapped", "M6", -25), ("counterbore", "M6", 20), ("countersink", "M5", 65)):
        r = S.hole(size, [[x, 0], [x, -30]], kind=kind, face_point=[x, 20, 0], depth=15 if kind == "tapped" else 0)
        check(f"hole {kind} {size} by Hole Wizard", r.get("method") == "hole_wizard", r.get("wizard_error", ""))
    m = S.measure([-70 + 4.5, 10, 0])
    near("M8 clearance diameter", m.get("diameter_mm"), 9.0, 0.01)


def sheetmetal() -> None:
    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 80, "height": 50}], plane="top")
    S.sheet_metal_base_flange(2, 2)
    S.edge_flange([0, 0, -25], 40)
    bb = S.model_info()["bounding_box_mm"]
    near("edge flange height", bb["max"][1], 40, 0.01)
    S.hole("M6", [[-25, -10], [25, -10]], face_point=[0, 0, 10])
    S.save(path("sm.SLDPRT"))
    r = S.export_flat_pattern(path("sm_flat.dxf"))
    lines = open(r["path"], encoding="latin-1").read().splitlines()
    pairs = list(zip(lines[0::2], lines[1::2]))
    start = next(i for i, (c, v) in enumerate(pairs) if c.strip() == "2" and v.strip() == "ENTITIES")
    xs = []
    for code, val in pairs[start:]:
        if code.strip() == "0" and val.strip() == "ENDSEC":
            break
        if code.strip() in ("10", "11"):
            xs.append(float(val))
    near("flat pattern length (50-4 + 42-4 + pi/2*3)", max(xs) - min(xs), 46 + 38 + math.pi / 2 * 3, 0.02)

    # a sharp-cornered L of 2 mm plate (like an imported STEP) made into sheet metal
    new()
    pts = [(0, 0), (40, 0), (40, 2), (2, 2), (2, 30), (0, 30), (0, 0)]
    S.sketch([{"type": "line", "x1": a[0], "y1": a[1], "x2": b[0], "y2": b[1]} for a, b in zip(pts, pts[1:])],
             plane="front", dimension=False)
    S.extrude(depth=30)
    bent = S.insert_bends([20, 2, 15], 2)
    check("insert_bends makes sheet metal", bent["sheet_metal"], str(bent))
    near("insert_bends flat length (36 + 26 + pi/2*3)", bent.get("flat_size_mm", [0])[0], 36 + 26 + math.pi / 2 * 3, 0.01)


def weldment() -> None:
    new()
    S.sketch3d([{"type": "polyline", "points": [[0, 0, 0], [500, 0, 0], [500, 0, 300], [0, 0, 300], [0, 0, 0]]}])
    S.sketch3d([{"type": "line", "p1": [0, 0, 0], "p2": [0, -400, 0]}, {"type": "line", "p1": [500, 0, 0], "p2": [500, -400, 0]}])
    S.structural_member("square tube", "40 x 40 x 4", sketch_name="3D스케치1")
    S.structural_member("square tube", "30 x 30 x 2.6", sketch_name="3D스케치2", corner="none")
    check("six members", len(S.list_bodies()) == 6)
    r = S.trim_members([[-15, -200, 0], [515, -200, 0]], [[0, -20, 150], [500, -20, 150], [250, -20, 0]])
    near("trim legs to the frame (2 x 20 mm of tube)", r["volume_removed_mm3"], 107020.6 / 400 * 20 * 2, 1)
    lengths = sorted(float(c["properties"].get("길이") or c["properties"].get("LENGTH") or 0) for c in S.cut_list()
                     if (c["properties"].get("길이") or c["properties"].get("LENGTH")))
    check("cut list lengths", lengths == [340.0, 380.0, 540.0], str(lengths))


def assembly() -> None:
    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 20, "height": 20}], plane="top")
    S.extrude(depth=20)
    S.save(path("blk.SLDPRT"))
    new("assembly")
    S.insert_component(path("blk.SLDPRT"), [0, 0, 0])
    S.insert_component(path("blk.SLDPRT"), [10, 0, 0])
    inter = S.check_interference()
    near("interference volume", inter[0]["volume_mm3"] if inter else None, 4000, 0.5)
    S.component_pattern(["live_blk-2"], 3, 30, "x")
    check("component pattern +2", len(S.list_components()) == 4)
    S.save(path("asm.SLDASM"))

    # advanced mates on three free 20 mm blocks: A fixed at the origin, B to the right, C behind
    new("assembly")
    for p in ([0, 0, 0], [40, 0, 0], [-40, 0, 40]):
        S.insert_component(path("blk.SLDPRT"), p)
    a, b, c = sorted(x["name"] for x in S.list_components())  # -1, -2, -3: the insertion order
    S.set_component(a, fixed=True)
    org = lambda name: sw.call(lambda: [v / sw.MM for v in S._origin(S._component(sw.active_doc(), name))])
    ax, ay, az = org(a)
    cx, cy, cz = org(c)
    S.add_mate("width", [ax - 10, ay + 10, az], [ax + 10, ay + 10, az],
               extra_points=[[cx - 10, cy + 10, cz], [cx + 10, cy + 10, cz]])
    near("width mate centres the tab", org(c)[0], ax, 0.01)
    bx, by, bz = org(b)
    limit = S.add_mate("distance", [bx, by, bz], [ax, ay + 20, az], value=5, min_value=0, max_value=30)
    near("limit distance starts at its value", org(b)[1] - ay, 25, 0.01)
    bx, by, bz = org(b)
    S.add_mate("symmetric", [bx - 10, by + 10, bz], [org(c)[0] - 10, cy + 10, cz], plane="right")  # the left faces
    near("symmetric mate mirrors the faces", org(b)[0] - 10, -(org(c)[0] - 10), 0.01)
    S.rename_feature(limit["mate"], "lift range")
    mates = {m["name"]: m for m in S.list_mates()}
    check("limit mate renamed and listed", mates.get("lift range", {}).get("limits") == [0.0, 30.0], str(mates.get("lift range")))

    # mechanical mates: a plate (slot x -30..-10 and a 10 mm hole at x 20), 8 mm pins and a rack bar
    new()
    S.sketch([{"type": "center_rectangle", "cx": 0, "cy": 0, "width": 80, "height": 40}], plane="top")
    S.extrude(depth=10)
    S.sketch([{"type": "slot", "x1": -30, "y1": 0, "x2": -10, "y2": 0, "width": 8},
              {"type": "circle", "cx": 20, "cy": 0, "r": 5}], face_point=[0, 10, 0], dimension=False)
    S.extrude(depth=10, cut=True, end="through_all")
    S.save(path("plate.SLDPRT"))
    new()
    S.sketch([{"type": "circle", "cx": 0, "cy": 0, "r": 4}], plane="top", dimension=False)
    S.extrude(depth=30)
    S.save(path("pin.SLDPRT"))
    new()
    S.sketch([{"type": "rectangle", "x1": -50, "y1": -5, "x2": 50, "y2": 5}], plane="top", dimension=False)
    S.extrude(depth=10)
    S.save(path("rack.SLDPRT"))
    new("assembly")
    plate = S.insert_component(path("plate.SLDPRT"))["component"]
    pin1 = S.insert_component(path("pin.SLDPRT"), [0, 30, 0])
    check("insert puts the origin at the position", pin1["origin_mm"] == [0.0, 30.0, 0.0], str(pin1))
    pin2 = S.insert_component(path("pin.SLDPRT"), [50, 30, 30], [0, 0, 90])["component"]
    turned = S.set_component(pin2)["y_axis"]
    check("insert turns the component (z 90: y axis -> -x)", turned == [-1.0, 0.0, 0.0], str(turned))
    axis = S.list_faces(component=pin2, type="cylinder")["faces"][0]["axis"]
    check("list_faces gives a turned component's axis in assembly space", abs(axis[0]) == 1.0, str(axis))
    S.set_component(pin2, rotation=[0, 0, 0])
    S.insert_component(path("rack.SLDPRT"), [0, 0, 60])
    S.add_mate("slot", [-20, 5, 4], [4, 45, 0])
    near("slot mate centres the pin (x)", org(pin1["component"])[0], -20, 0.01)
    S.add_mate("hinge", [54, 45, 30], [25, 5, 0], extra_points=[[50, 30, 30], [0, 10, 10]])
    near("hinge: pin in the hole (x)", org(pin2)[0], 20, 0.01)
    near("hinge: pin on the plate (y)", org(pin2)[1], 10, 0.01)
    S.add_mate("rack_pinion", [0, 10, 65], [-16, 45, 0], value=20)
    check("rack and pinion mate", any(m.get("type") == "rack_pinion" for m in S.list_mates()))
    # a part with a second, shorter configuration, switched in the assembly
    new()
    S.sketch([{"type": "circle", "cx": 0, "cy": 0, "r": 5}], plane="top", dimension=False)
    S.extrude(depth=40)
    S.configurations(add="short")
    S.set_dimension("D1@보스-돌출1", 10)
    S.configurations(activate="기본")
    S.save(path("rod.SLDPRT"))
    new("assembly")
    rod = S.insert_component(path("rod.SLDPRT"))["component"]
    switched = S.set_component(rod, configuration="short")
    check("set_component switches the configuration", switched["configuration"] == "short", str(switched))
    near("the short configuration is 10 mm long", max(f["point_mm"][1] for f in S.list_faces(type="plane")["faces"]), 10, 0.01)

    new("assembly")
    S.insert_component(path("plate.SLDPRT"))
    pin = S.insert_component(path("pin.SLDPRT"), [50, 30, 30])["component"]
    S.add_mate("screw", [54, 45, 30], [25, 5, 0], value=2)
    o = org(pin)
    check("screw mate puts the pin on the hole axis", abs(o[0] - 20) < 0.01 and abs(o[2]) < 0.01, str(o))


def drawing() -> None:
    opened(S.open_document(path("sm.SLDPRT")))
    d = S.make_drawing()
    check("model dimensions imported", d.get("dimensions_imported", 0) >= 3, str(d.get("dimensions_imported")))
    check("hole callout", bool(d.get("hole_callouts")))
    S.title_block({"Description": "live test"}, sheet_format="a3 - iso")
    r = S.add_dimension([-25, 0, 10], [25, 0, 10], entity1="center", entity2="center", view="도면뷰2", direction="horizontal")
    near("hole spacing dimension", r["value_mm"], 50, 0.01)
    S.set_tolerance(r["dimension"], "symmetric", upper=0.1)
    holes_dims = S.dimension_holes(view="도면뷰2")
    check("dimension_holes", holes_dims["horizontal_dimensions"] >= 3, str(holes_dims))
    styles = S.view_style(display="hidden_visible", tangent_edges="hidden")["views"]
    check("view style read back", all(v["display"] == "hidden_visible" and v["tangent_edges"] == "hidden"
                                      for v in styles.values()), str(styles))
    lines = S.centerlines()["centerlines_added"]
    check("centerlines on side views", sum(lines.values()) + d.get("centerlines", 0) >= 1, f"{lines}, make_drawing {d.get('centerlines')}")
    flat = S.flat_pattern_view()
    near("flat pattern view width", flat["dimensions"].get("horizontal"), 46 + 38 + math.pi / 2 * 3, 0.02)
    out = S.export(path("sm.pdf"))
    check("drawing PDF written", out["bytes"] > 10000, f"{out['bytes']} bytes, {out.get('method', 'export')}")


GROUPS = {"parts": parts, "holes": holes, "sheetmetal": sheetmetal, "weldment": weldment,
          "assembly": assembly, "drawing": drawing}


def main() -> int:
    wanted = [a for a in sys.argv[1:] if a in GROUPS] or list(GROUPS)
    if "drawing" in wanted and "sheetmetal" not in wanted:
        wanted.insert(wanted.index("drawing"), "sheetmetal")  # the drawing uses the sheet metal part
    print(f"SOLIDWORKS live regression in {WORK}\n")
    start = time.time()
    for name in wanted:
        print(f"== {name}")
        try:
            GROUPS[name]()
        except Exception as exc:
            check(f"{name} group ran", False, f"{type(exc).__name__}: {exc}")
            traceback.print_exc()
        finally:
            if name != "sheetmetal":  # keep the part open for the drawing group
                close_mine()
    close_mine()
    failed = sum(1 for _, ok, _ in RESULTS if not ok)
    print(f"\n{len(RESULTS) - failed} passed, {failed} failed in {time.time() - start:.0f} s")
    if os.environ.get("SWMCP_KEEP"):
        print(f"files kept in {WORK}")
    else:
        shutil.rmtree(WORK, ignore_errors=True)
    return failed


if __name__ == "__main__":
    sys.exit(main())
