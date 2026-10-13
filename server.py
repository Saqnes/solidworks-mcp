"""MCP server that drives SOLIDWORKS through its COM API.

Lengths are millimetres and angles degrees in every tool; points are model
coordinates of the active document (sketch tools take 2D sketch coordinates).
Anything the dedicated tools do not cover can be done with run_api_script,
which runs Python against the live API object.
"""

from __future__ import annotations

import contextlib
import functools
import io
import math
import os
import tempfile
import traceback
from pathlib import Path
from typing import Any, Literal

from mcp.server.mcpserver import Image, MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import sw
from sw import DOC_ASSEMBLY, DOC_DRAWING, DOC_PART, MM, call, deg, mm

mcp = MCPServer(
    "solidworks",
    instructions=(
        "Drives SOLIDWORKS via its API. Units: millimetres and degrees. Typical flow: "
        "status -> new_document or open_document -> sketch -> extrude/revolve -> "
        "screenshot to check -> save. Feature and plane names are localized (the UI may "
        "be Korean), so read them from list_features instead of guessing. For anything "
        "without a dedicated tool, use run_api_script."
    ),
)


def tool(fn):
    """Register fn as a tool. Its body runs on the COM thread, and any failure
    (including COM errors) comes back to the agent as a readable error."""

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return call(lambda: fn(*args, **kwargs))
        except ToolError:
            raise
        except Exception as exc:
            raise ToolError(f"{type(exc).__name__}: {exc}") from exc

    return mcp.tool()(wrapper)


def _rebuild(doc: Any) -> None:
    doc.ForceRebuild3(False)


def _feature_result(feat: Any, what: str) -> dict:
    if feat is None:
        raise RuntimeError(
            f"SOLIDWORKS did not create the {what}. Check the selection/geometry "
            "(screenshot, list_features) or build it with run_api_script."
        )
    return {"feature": feat.Name, "type": feat.GetTypeName2()}


# --- Session and documents ------------------------------------------------


@tool
def status() -> dict:
    """SOLIDWORKS version and the active document, if any."""
    app = sw.app()
    doc = app.ActiveDoc
    return {"revision": app.RevisionNumber(), "active_document": sw.doc_info(doc) if doc else None}


@tool
def list_documents() -> list[dict]:
    """All documents currently open in SOLIDWORKS."""
    docs = sw.app().GetDocuments() or ()
    return [sw.doc_info(d) for d in docs]


@tool
def new_document(kind: Literal["part", "assembly", "drawing"] = "part") -> dict:
    """Create a document from the default template set in SOLIDWORKS options."""
    app = sw.app()
    doc_type = {"part": DOC_PART, "assembly": DOC_ASSEMBLY, "drawing": DOC_DRAWING}[kind]
    template = app.GetUserPreferenceStringValue(sw.TEMPLATE_PREF[doc_type])
    if not template:
        raise RuntimeError(f"No default {kind} template configured (Tools > Options > Default Templates)")
    doc = app.NewDocument(template, 8, 0.42, 0.297)  # paper args only matter for drawings (A3)
    if doc is None:
        raise RuntimeError(f"Could not create a {kind} from {template}")
    return sw.doc_info(doc)


@tool
def open_document(path: str) -> dict:
    """Open a .sldprt/.sldasm/.slddrw file and make it active."""
    ext = Path(path).suffix.lower()
    if ext not in sw.DOC_TYPES:
        raise ValueError("Only .sldprt, .sldasm and .slddrw can be opened; use run_api_script (LoadFile4) to import STEP etc.")
    errors, warnings = sw.out_int(), sw.out_int()
    doc = sw.app().OpenDoc6(str(Path(path)), sw.DOC_TYPES[ext], 1, "", errors, warnings)  # 1 = silent
    if doc is None:
        raise RuntimeError(f"Open failed (swFileLoadError {errors.value}) for {path}")
    return {**sw.doc_info(doc), "warnings": warnings.value}


@tool
def activate_document(title: str) -> dict:
    """Bring an open document (by title, as in list_documents) to the front."""
    errors = sw.out_int()
    doc = sw.app().ActivateDoc3(title, False, 0, errors)
    if doc is None:
        raise LookupError(f"No open document titled {title!r} (error {errors.value})")
    return sw.doc_info(doc)


@tool
def save(path: str = "") -> dict:
    """Save the active document; with a path (same file type), save it there
    and keep working on the new file."""
    doc = sw.active_doc()
    errors, warnings = sw.out_int(), sw.out_int()
    if path:
        ok = doc.Extension.SaveAs3(str(Path(path)), 0, 1, sw.nothing(), sw.nothing(), errors, warnings)
    else:
        if not doc.GetPathName():
            raise ValueError("Document was never saved; give a path")
        ok = doc.Save3(1, errors, warnings)
    if not ok:
        raise RuntimeError(f"Save failed (swFileSaveError {errors.value})")
    return {**sw.doc_info(doc), "warnings": warnings.value}


@tool
def export(path: str) -> dict:
    """Export a copy of the active document; the format follows the extension
    (.step/.stp, .igs, .x_t, .stl, .3mf, .pdf, .dxf/.dwg (drawings), .png/.jpg,
    .eDrawings). The open document keeps its own file."""
    doc = sw.active_doc()
    errors, warnings = sw.out_int(), sw.out_int()
    target = str(Path(path))
    ok = doc.Extension.SaveAs3(target, 0, 1 | 2, sw.nothing(), sw.nothing(), errors, warnings)  # silent | copy
    if not ok or not os.path.exists(target):
        raise RuntimeError(f"Export failed (swFileSaveError {errors.value}) for {target}")
    return {"path": target, "bytes": os.path.getsize(target), "warnings": warnings.value}


@tool
def close_document(title: str = "", discard_changes: bool = False) -> str:
    """Close a document (default: the active one). Refuses if it has unsaved
    changes unless discard_changes=true."""
    app = sw.app()
    doc = sw.active_doc() if not title else next((d for d in (app.GetDocuments() or ()) if d.GetTitle() == title), None)
    if doc is None:
        raise LookupError(f"No open document titled {title!r}")
    if doc.GetSaveFlag() and not discard_changes:
        raise RuntimeError("Document has unsaved changes: save first or pass discard_changes=true")
    name = doc.GetTitle()
    app.CloseDoc(name)
    return f"closed {name}"


# --- Inspecting -----------------------------------------------------------


@tool
def list_features(include_hidden: bool = False) -> list[dict]:
    """Feature tree of the active document, in order. Folders and other
    system entries are skipped unless include_hidden=true."""
    hidden = {
        "HistoryFolder", "SensorFolder", "DocsFolder", "DetailCabinet", "SurfaceBodyFolder",
        "SolidBodyFolder", "EnvFolder", "InkMarkupFolder", "EqnFolder", "MaterialFolder",
        "CommentsFolder", "FavoriteFolder", "SelectionSetFolder", "MateReferenceGroupFolder",
        "AmbientLight", "DirectionLight", "PointLight", "SpotLight", "MateGroup", "BlockFolder",
        "LiveSectionFolder", "MarkupCommentFolder", "EnvironmentFolder", "CutListFolder",
    }
    out = []
    for f in sw.iter_features(sw.active_doc()):
        t = f.GetTypeName2()
        if include_hidden or t not in hidden:
            out.append({"name": f.Name, "type": t, "suppressed": bool(f.IsSuppressed())})
    return out


def _dim_value(dim: Any) -> tuple[float, str]:
    if dim.GetType() == 2:  # swDimensionParamTypeDoubleAngular
        return round(math.degrees(dim.SystemValue), 6), "deg"
    return round(dim.SystemValue / MM, 6), "mm"


@tool
def list_dimensions(feature: str = "") -> list[dict]:
    """Dimensions of all features (or one feature), with names usable in
    set_dimension, e.g. "D1@Sketch1" (localized: "D1@스케치1")."""
    doc = sw.active_doc()
    feats = [sw.feature(doc, feature)] if feature else list(sw.iter_features(doc))
    out = []
    for f in feats:
        dd = f.GetFirstDisplayDimension()
        while dd is not None:
            dim = dd.GetDimension2(0)
            value, unit = _dim_value(dim)
            out.append({"name": "@".join(dim.FullName.split("@")[:2]), "value": value, "unit": unit, "feature": f.Name})
            dd = f.GetNextDisplayDimension(dd)
    return out


@tool
def model_info() -> dict:
    """Mass properties, bounding box and material of the active part or
    assembly (mass uses the assigned material's density)."""
    doc = sw.active_doc()
    if doc.GetType() == DOC_DRAWING:
        raise RuntimeError("model_info needs a part or assembly")
    info: dict = {"document": sw.doc_info(doc)}
    mp = doc.Extension.CreateMassProperty()
    if mp is not None:
        info.update(
            mass_kg=round(mp.Mass, 6),
            volume_mm3=round(mp.Volume / MM**3, 3),
            surface_area_mm2=round(mp.SurfaceArea / MM**2, 3),
            density_kg_m3=round(mp.Density, 3),
            center_of_mass_mm=[round(v / MM, 4) for v in mp.CenterOfMass],
        )
    if doc.GetType() == DOC_PART:
        box = doc.GetPartBox(True)
        if box:
            info["bounding_box_mm"] = {"min": [round(v / MM, 4) for v in box[:3]], "max": [round(v / MM, 4) for v in box[3:]]}
        db = sw.out_bstr()
        info["material"] = doc.GetMaterialPropertyName2("", db) or None
    return info


VIEWS = {"front": 1, "back": 2, "left": 3, "right": 4, "top": 5, "bottom": 6, "isometric": 7, "trimetric": 8, "dimetric": 9}


@tool
def screenshot(view: Literal["current", "front", "back", "left", "right", "top", "bottom", "isometric", "trimetric", "dimetric"] = "isometric", zoom_to_fit: bool = True) -> list:
    """Image of the active document's graphics area, optionally reoriented."""
    from PIL import Image as PILImage

    doc = sw.active_doc()
    if view != "current" and doc.GetType() != DOC_DRAWING:
        doc.ShowNamedView2("", VIEWS[view])
    if zoom_to_fit:
        doc.ViewZoomtofit2()
    doc.GraphicsRedraw2()
    with tempfile.TemporaryDirectory() as tmp:
        bmp = os.path.join(tmp, "view.bmp")
        if not doc.SaveBMP(bmp, 1280, 960) or not os.path.exists(bmp):
            raise RuntimeError("SOLIDWORKS could not save the view image")
        img = PILImage.open(bmp).convert("RGB")
        img.thumbnail((1280, 1280))
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=85)
    return [f"{doc.GetTitle()} ({view})", Image(data=buf.getvalue(), format="jpeg")]


# --- Editing --------------------------------------------------------------


@tool
def set_dimension(name: str, value: float) -> dict:
    """Change a dimension (mm, or degrees for angles) and rebuild.
    `name` is as returned by list_dimensions, e.g. "D1@Sketch1"."""
    doc = sw.active_doc()
    dim = doc.Parameter(name)
    if dim is None:
        raise LookupError(f"No dimension {name!r}; use list_dimensions")
    dim.SystemValue = deg(value) if dim.GetType() == 2 else mm(value)
    _rebuild(doc)
    new, unit = _dim_value(dim)
    return {"name": name, "value": new, "unit": unit}


@tool
def rebuild() -> str:
    """Force a full rebuild of the active document."""
    _rebuild(sw.active_doc())
    return "rebuilt"


@tool
def undo(steps: int = 1) -> str:
    """Undo the last `steps` operations."""
    sw.active_doc().EditUndo2(steps)
    return f"undid {steps}"


@tool
def delete_feature(name: str) -> str:
    """Delete a feature together with its absorbed sketch."""
    doc = sw.active_doc()
    doc.ClearSelection2(True)
    sw.select_feature(doc, name)
    if not doc.Extension.DeleteSelection2(1):  # swDelete_Absorbed
        raise RuntimeError(f"Could not delete {name!r} (other features may depend on it)")
    return f"deleted {name}"


@tool
def suppress_feature(name: str, suppress: bool = True) -> str:
    """Suppress or unsuppress a feature."""
    doc = sw.active_doc()
    doc.ClearSelection2(True)
    sw.select_feature(doc, name)
    ok = doc.EditSuppress2() if suppress else doc.EditUnsuppress2()
    if not ok:
        raise RuntimeError(f"Could not change suppression of {name!r}")
    _rebuild(doc)
    return f"{'suppressed' if suppress else 'unsuppressed'} {name}"


@tool
def set_material(material: str, database: str = "SOLIDWORKS Materials") -> dict:
    """Assign a material to the active part, e.g. "AISI 304", "6061-T6 (SS)",
    "ABS". Names must match the material library exactly."""
    doc = sw.active_doc(DOC_PART)
    doc.SetMaterialPropertyName2("", database, material)
    db = sw.out_bstr()
    applied = doc.GetMaterialPropertyName2("", db)
    if applied != material:
        raise LookupError(f"Material {material!r} not found in {database!r}")
    return {"material": applied, "density_kg_m3": round(doc.Extension.CreateMassProperty().Density, 3)}


@tool
def custom_properties(set_values: dict[str, str] | None = None) -> dict[str, str]:
    """Read the document's custom properties, after first adding/replacing any
    given in set_values (e.g. {"PartNo": "A-001", "Description": "Bracket"})."""
    cpm = sw.active_doc().Extension.CustomPropertyManager("")
    for k, v in (set_values or {}).items():
        cpm.Add3(k, 30, str(v), 2)  # text, replace existing
    return {n: cpm.Get(n) for n in (cpm.GetNames() or ())}


# --- Part modelling -------------------------------------------------------


def _sketch_entity(sm: Any, e: dict) -> Any:
    t = e.get("type")
    g = lambda k, d=None: mm(e[k]) if k in e else d  # noqa: E731
    if t == "line":
        return sm.CreateLine(g("x1"), g("y1"), 0, g("x2"), g("y2"), 0)
    if t == "centerline":
        return sm.CreateCenterLine(g("x1"), g("y1"), 0, g("x2"), g("y2"), 0)
    if t == "rectangle":
        return sm.CreateCornerRectangle(g("x1"), g("y1"), 0, g("x2"), g("y2"), 0)
    if t == "center_rectangle":
        cx, cy = g("cx", 0.0), g("cy", 0.0)
        return sm.CreateCenterRectangle(cx, cy, 0, cx + g("width") / 2, cy + g("height") / 2, 0)
    if t == "circle":
        return sm.CreateCircleByRadius(g("cx", 0.0), g("cy", 0.0), 0, g("r"))
    if t == "arc":
        return sm.CreateArc(g("cx"), g("cy"), 0, g("x1"), g("y1"), 0, g("x2"), g("y2"), 0, -1 if e.get("clockwise") else 1)
    if t == "polygon":
        cx, cy, r = g("cx", 0.0), g("cy", 0.0), g("r")
        return sm.CreatePolygon(cx, cy, 0, cx + r, cy, 0, int(e["sides"]), True)
    if t == "slot":
        return sm.CreateSketchSlot(0, 0, g("width"), g("x1"), g("y1"), 0, g("x2"), g("y2"), 0, 0, 0, 0, 1, False)
    if t == "point":
        return sm.CreatePoint(g("x"), g("y"), 0)
    raise ValueError(f"Unknown sketch entity type {t!r}")


@tool
def sketch(entities: list[dict], plane: str = "front", face_point: list[float] | None = None) -> dict:
    """Create a sketch and close it. Place it on "front"/"top"/"right", on any
    plane feature by name, or on the face at face_point ([x,y,z] mm).
    Coordinates are 2D sketch coordinates in mm. Entity types:
      {"type":"line","x1","y1","x2","y2"}  {"type":"centerline",...same}
      {"type":"rectangle","x1","y1","x2","y2"}  {"type":"center_rectangle","cx","cy","width","height"}
      {"type":"circle","cx","cy","r"}  {"type":"polygon","cx","cy","r","sides"} (inscribed)
      {"type":"arc","cx","cy","x1","y1","x2","y2","clockwise":false}
      {"type":"slot","x1","y1","x2","y2","width"}  {"type":"point","x","y"}
    Profiles for extrude must be closed; a revolve profile needs one centerline as its axis."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    if face_point:
        sw.select_at(doc, "FACE", face_point, False)
    else:
        planes = sw.standard_planes(doc)
        target = planes.get(plane.lower()) or sw.feature(doc, plane)
        if not target.Select2(False, 0):
            raise RuntimeError(f"Could not select plane {plane!r}")
    sm = doc.SketchManager
    sm.InsertSketch(True)
    sm.AddToDB = True  # no snapping/inference to nearby geometry
    failed = []
    try:
        for i, e in enumerate(entities):
            if _sketch_entity(sm, e) is None:
                failed.append(i)
    finally:
        sm.AddToDB = False
        sm.InsertSketch(True)
    feat = sw.last_feature_of_type(doc, "ProfileFeature")
    result = {"sketch": feat.Name, "entities": len(entities) - len(failed)}
    if failed:
        result["failed_entities"] = failed
    return result


END_CONDITIONS = {"blind": 0, "through_all": 1, "up_to_next": 2, "mid_plane": 6}


@tool
def extrude(
    depth: float = 10.0,
    sketch_name: str = "",
    end: Literal["blind", "through_all", "up_to_next", "mid_plane"] = "blind",
    reverse: bool = False,
    cut: bool = False,
    depth2: float = 0.0,
    draft_angle: float = 0.0,
    merge: bool = True,
) -> dict:
    """Boss-extrude (or cut-extrude with cut=true) a sketch, default the most
    recent one. depth2 > 0 also extrudes blind in the opposite direction;
    mid_plane splits `depth` evenly about the sketch plane."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    name = sketch_name or sw.last_feature_of_type(doc, "ProfileFeature").Name
    sw.select_feature(doc, name)
    fm = doc.FeatureManager
    single = depth2 <= 0
    t1, d1 = END_CONDITIONS[end], mm(depth)
    t2, d2 = 0, mm(depth2)
    draft = draft_angle != 0
    da = deg(abs(draft_angle))
    if cut:
        feat = fm.FeatureCut4(
            single, False, reverse, t1, t2, d1, d2, draft, draft, draft_angle < 0, draft_angle < 0, da, da,
            False, False, False, False, False, True, True, False, False, False, 0, 0, False, False,
        )
    else:
        feat = fm.FeatureExtrusion3(
            single, False, reverse, t1, t2, d1, d2, draft, draft, draft_angle < 0, draft_angle < 0, da, da,
            False, False, False, False, merge, True, True, 0, 0, False,
        )
    return _feature_result(feat, "cut-extrude" if cut else "extrude")


@tool
def revolve(angle: float = 360.0, sketch_name: str = "", cut: bool = False, reverse: bool = False) -> dict:
    """Revolve a sketch about its centerline (default the most recent sketch).
    cut=true removes material instead."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    name = sketch_name or sw.last_feature_of_type(doc, "ProfileFeature").Name
    sw.select_feature(doc, name)
    feat = doc.FeatureManager.FeatureRevolve2(
        True, True, False, cut, reverse, False, 0, 0, deg(angle), 0,
        False, False, 0, 0, 0, 0, 0, True, True, True,
    )
    return _feature_result(feat, "revolve")


@tool
def fillet(radius: float, edge_points: list[list[float]]) -> dict:
    """Constant-radius fillet on the edges passing through the given model
    points ([[x,y,z], ...] in mm; a point on the edge, e.g. its midpoint)."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    for i, p in enumerate(edge_points):
        sw.select_at(doc, "EDGE", p, append=i > 0, mark=1)
    e = sw.empty
    feat = doc.FeatureManager.FeatureFillet3(195, mm(radius), 0, 0, 0, 0, e(), e(), e(), e(), e(), e())
    return _feature_result(feat, "fillet")


@tool
def chamfer(distance: float, edge_points: list[list[float]], angle: float = 45.0) -> dict:
    """Angle-distance chamfer on the edges through the given points (mm)."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    for i, p in enumerate(edge_points):
        sw.select_at(doc, "EDGE", p, append=i > 0, mark=1)
    feat = doc.FeatureManager.InsertFeatureChamfer(4, 1, mm(distance), deg(angle), 0, 0, 0, 0)
    return _feature_result(feat, "chamfer")


# --- Assemblies -----------------------------------------------------------


@tool
def insert_component(path: str, position: list[float] | None = None) -> dict:
    """Insert a part/assembly file into the active assembly at [x,y,z] mm."""
    app = sw.app()
    asm = sw.active_doc(DOC_ASSEMBLY)
    asm_title = asm.GetTitle()
    ext = Path(path).suffix.lower()
    if ext not in (".sldprt", ".sldasm"):
        raise ValueError("Component must be a .sldprt or .sldasm file")
    errors, warnings = sw.out_int(), sw.out_int()
    if app.OpenDoc6(str(Path(path)), sw.DOC_TYPES[ext], 1, "", errors, warnings) is None:
        raise RuntimeError(f"Could not open {path} (error {errors.value})")
    asm = app.ActivateDoc3(asm_title, False, 0, sw.out_int())
    x, y, z = (mm(v) for v in (position or [0, 0, 0]))
    comp = asm.AddComponent5(str(Path(path)), 0, "", False, "", x, y, z)
    if comp is None:
        raise RuntimeError(f"Could not insert {path}")
    return {"component": comp.Name2}


@tool
def list_components(top_level_only: bool = True) -> list[dict]:
    """Components of the active assembly."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    return [
        {"name": c.Name2, "path": c.GetPathName(), "suppressed": bool(c.IsSuppressed()), "fixed": bool(c.IsFixed())}
        for c in (asm.GetComponents(top_level_only) or ())
    ]


MATES = {"coincident": 0, "concentric": 1, "perpendicular": 2, "parallel": 3, "tangent": 4, "distance": 5, "angle": 6}
ALIGN = {"aligned": 0, "anti_aligned": 1, "closest": 2}


@tool
def add_mate(
    mate: Literal["coincident", "concentric", "perpendicular", "parallel", "tangent", "distance", "angle"],
    point1: list[float],
    point2: list[float],
    entity1: Literal["face", "edge", "vertex"] = "face",
    entity2: Literal["face", "edge", "vertex"] = "face",
    value: float = 0.0,
    align: Literal["aligned", "anti_aligned", "closest"] = "closest",
    flip: bool = False,
) -> dict:
    """Mate two entities picked by assembly-space points (mm). `value` is the
    distance (mm) for distance mates or the angle (deg) for angle mates."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    asm.ClearSelection2(True)
    sw.select_at(asm, entity1, point1, append=False, mark=1)
    sw.select_at(asm, entity2, point2, append=True, mark=1)
    dist = mm(value) if mate == "distance" else 0
    ang = deg(value) if mate == "angle" else 0
    err = sw.out_int()
    feat = asm.AddMate5(MATES[mate], ALIGN[align], flip, dist, dist, dist, 1, 1, ang, ang, ang, False, False, 0, err)
    if feat is None:
        raise RuntimeError(f"Mate failed (swAddMateError {err.value})")
    asm.EditRebuild3()
    return {"mate": feat.Name}


# --- Drawings -------------------------------------------------------------


@tool
def make_drawing(model_path: str = "", projection: Literal["third_angle", "first_angle"] = "third_angle", import_dimensions: bool = True) -> dict:
    """Create a drawing with the standard three views (plus model dimensions)
    of a saved part/assembly; default the active document. Export it as PDF
    or DXF afterwards with export."""
    app = sw.app()
    if not model_path:
        src = sw.active_doc()
        model_path = src.GetPathName()
        if not model_path:
            raise RuntimeError("Save the model first; a drawing references the model file")
    template = app.GetUserPreferenceStringValue(sw.TEMPLATE_PREF[DOC_DRAWING])
    drw = app.NewDocument(template, 8, 0.42, 0.297)
    if drw is None:
        raise RuntimeError("Could not create a drawing; check the default drawing template")
    create = drw.Create3rdAngleViews2 if projection == "third_angle" else drw.Create1stAngleViews2
    if not create(model_path):
        raise RuntimeError(f"Could not create views of {model_path}")
    if import_dimensions:
        # all views, marked-for-drawing dimensions, then rebuild
        drw.InsertModelAnnotations3(0, 32768, True, True, False, False)
    drw.ForceRebuild3(False)
    return sw.doc_info(drw)


# --- Escape hatches -------------------------------------------------------


@tool
def run_api_script(code: str) -> dict:
    """Run Python against the live SOLIDWORKS API for anything the other
    tools do not cover. Available names: swApp, doc (active document or
    None), mm (0.001), deg (degrees -> radians), nothing(), empty(),
    out_int() (ByRef Long; read .value), select_at(doc, "FACE", [x,y,z]mm,
    append), win32com, pythoncom, math. Assign to `result` to return a value;
    print() output is returned too. The API uses metres and radians."""
    app = sw.app()
    env = {
        "swApp": app, "doc": app.ActiveDoc, "mm": MM, "deg": math.radians, "math": math,
        "nothing": sw.nothing, "empty": sw.empty, "out_int": sw.out_int, "select_at": sw.select_at,
        "result": None,
    }
    import pythoncom
    import win32com

    env.update(win32com=win32com, pythoncom=pythoncom)
    out = io.StringIO()
    try:
        with contextlib.redirect_stdout(out):
            exec(compile(code, "<run_api_script>", "exec"), env)
    except Exception:
        return {"ok": False, "stdout": out.getvalue(), "error": traceback.format_exc(limit=3)}
    return {"ok": True, "stdout": out.getvalue(), "result": repr(env["result"]) if env["result"] is not None else None}


@tool
def run_macro(path: str, module: str, procedure: str) -> str:
    """Run a VBA macro (.swp), e.g. module "Module1", procedure "main"."""
    err = sw.out_int()
    if not sw.app().RunMacro2(str(Path(path)), module, procedure, 1, err):
        raise RuntimeError(f"Macro failed (swRunMacroError {err.value})")
    return "macro finished"


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
