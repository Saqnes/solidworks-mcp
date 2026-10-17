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
import time
import traceback
import xml.etree.ElementTree as ET
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
    return {"feature": feat.Name, "type": sw.m0(feat, "GetTypeName2")}


# --- Session and documents ------------------------------------------------


@tool
def status() -> dict:
    """SOLIDWORKS version and the active document, if any."""
    app = sw.app()
    doc = app.ActiveDoc
    return {"revision": sw.m0(app, "RevisionNumber"), "active_document": sw.doc_info(doc) if doc else None}


@tool
def list_documents() -> list[dict]:
    """All documents currently open in SOLIDWORKS."""
    docs = sw.m0(sw.app(), "GetDocuments") or ()
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


SAVE_ERRORS = {
    1: "generic save error", 2: "file is read-only", 4: "file name is empty",
    8: "file name contains '@'", 16: "file is locked by another program or user",
    32: "format not available", 128: "target exists and may not be overwritten",
    256: "invalid file extension", 1024: "bad eDrawings version",
    2048: "path is too long", 4096: "this document type cannot be saved in that format",
    8192: "referenced documents must be saved first",
}


def _save_error(code: int) -> str:
    reasons = [text for bit, text in SAVE_ERRORS.items() if code & bit]
    return f"swFileSaveError {code}: {', '.join(reasons) or 'unknown'}"


def _check_not_open_elsewhere(doc: Any, target: str) -> None:
    """A path open in SOLIDWORKS as another document cannot be written."""
    norm = os.path.normcase(os.path.abspath(target))
    for other in sw.m0(sw.app(), "GetDocuments") or ():
        path = sw.m0(other, "GetPathName")
        if path and os.path.normcase(os.path.abspath(path)) == norm and sw.m0(other, "GetTitle") != sw.m0(doc, "GetTitle"):
            raise RuntimeError(
                f"{target} is open in SOLIDWORKS as {sw.m0(other, 'GetTitle')!r}; "
                "close it first (close_document) or choose another name"
            )
    if os.path.exists(target) and not os.access(target, os.W_OK):
        raise RuntimeError(f"{target} is read-only")


@tool
def save(path: str = "") -> dict:
    """Save the active document; with a path (same file type), save it there
    and keep working on the new file."""
    doc = sw.active_doc()
    errors, warnings = sw.out_int(), sw.out_int()
    if path:
        _check_not_open_elsewhere(doc, str(Path(path)))
        ok = doc.Extension.SaveAs3(str(Path(path)), 0, 1, sw.nothing(), sw.nothing(), errors, warnings)
    else:
        if not sw.m0(doc, "GetPathName"):
            raise ValueError("Document was never saved; give a path")
        ok = doc.Save3(1, errors, warnings)
    if not ok:
        raise RuntimeError(f"Save failed ({_save_error(errors.value)})")
    return {**sw.doc_info(doc), "warnings": warnings.value}


@tool
def export(path: str) -> dict:
    """Export a copy of the active document; the format follows the extension
    (.step/.stp, .igs, .x_t, .stl, .3mf, .pdf, .dxf/.dwg (drawings), .png/.jpg,
    .eDrawings). The open document keeps its own file."""
    doc = sw.active_doc()
    errors, warnings = sw.out_int(), sw.out_int()
    target = str(Path(path))
    _check_not_open_elsewhere(doc, target)
    ok = doc.Extension.SaveAs3(target, 0, 1 | 2, sw.nothing(), sw.nothing(), errors, warnings)  # silent | copy
    if not ok or not os.path.exists(target):
        raise RuntimeError(f"Export failed ({_save_error(errors.value)}) for {target}")
    return {"path": target, "bytes": os.path.getsize(target), "warnings": warnings.value}


@tool
def close_document(title: str = "", discard_changes: bool = False) -> str:
    """Close a document (default: the active one). Refuses if it has unsaved
    changes unless discard_changes=true."""
    app = sw.app()
    doc = sw.active_doc() if not title else next((d for d in (sw.m0(app, "GetDocuments") or ()) if sw.m0(d, "GetTitle") == title), None)
    if doc is None:
        raise LookupError(f"No open document titled {title!r}")
    if sw.m0(doc, "GetSaveFlag") and not discard_changes:
        raise RuntimeError("Document has unsaved changes: save first or pass discard_changes=true")
    name = sw.m0(doc, "GetTitle")
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
        t = sw.m0(f, "GetTypeName2")
        if include_hidden or t not in hidden:
            out.append({"name": f.Name, "type": t, "suppressed": bool(sw.m0(f, "IsSuppressed"))})
    return out


def _dim_value(dim: Any) -> tuple[float, str]:
    if sw.m0(dim, "GetType") == 2:  # swDimensionParamTypeDoubleAngular
        return round(math.degrees(dim.SystemValue), 6), "deg"
    return round(dim.SystemValue / MM, 6), "mm"


@tool
def list_dimensions(feature: str = "") -> list[dict]:
    """Dimensions of all features (or one feature), with names usable in
    set_dimension, e.g. "D1@Sketch1" (localized: "D1@스케치1"). In a drawing,
    lists the dimensions shown in each view instead."""
    doc = sw.active_doc()
    if sw.m0(doc, "GetType") == DOC_DRAWING:
        return _drawing_dimensions(doc)
    feats = [sw.feature(doc, feature)] if feature else list(sw.iter_features(doc))
    out = []
    for f in feats:
        dd = sw.m0(f, "GetFirstDisplayDimension")
        while dd is not None:
            dim = dd.GetDimension2(0)
            value, unit = _dim_value(dim)
            out.append({"name": "@".join(dim.FullName.split("@")[:2]), "value": value, "unit": unit, "feature": f.Name})
            dd = f.GetNextDisplayDimension(dd)
    return out


def _drawing_dimensions(drw: Any) -> list[dict]:
    out = []
    view = sw.m0(drw, "GetFirstView")  # the sheet itself, then each view
    while view is not None:
        dd = sw.m0(view, "GetFirstDisplayDimension5")
        while dd is not None:
            dim = dd.GetDimension2(0)
            value, unit = _dim_value(dim)
            out.append({"view": view.Name, "name": "@".join(dim.FullName.split("@")[:2]), "value": value, "unit": unit})
            dd = sw.m0(dd, "GetNext5")
        view = sw.m0(view, "GetNextView")
    return out


@tool
def model_info() -> dict:
    """Mass properties, bounding box and material of the active part or
    assembly (mass uses the assigned material's density)."""
    doc = sw.active_doc()
    if sw.m0(doc, "GetType") == DOC_DRAWING:
        raise RuntimeError("model_info needs a part or assembly")
    info: dict = {"document": sw.doc_info(doc)}
    mp = sw.m0(doc.Extension, "CreateMassProperty")
    if mp is not None:
        info.update(
            mass_kg=round(mp.Mass, 6),
            volume_mm3=round(mp.Volume / MM**3, 3),
            surface_area_mm2=round(mp.SurfaceArea / MM**2, 3),
            density_kg_m3=round(mp.Density, 3),
            center_of_mass_mm=[round(v / MM, 4) for v in mp.CenterOfMass],
        )
    if sw.m0(doc, "GetType") == DOC_PART:
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
    if view != "current" and sw.m0(doc, "GetType") != DOC_DRAWING:
        doc.ShowNamedView2("", VIEWS[view])
    if zoom_to_fit:
        sw.m0(doc, "ViewZoomtofit2")
    sw.m0(doc, "GraphicsRedraw2")
    with tempfile.TemporaryDirectory() as tmp:
        bmp = os.path.join(tmp, "view.bmp")
        if not doc.SaveBMP(bmp, 1280, 960) or not os.path.exists(bmp):
            raise RuntimeError("SOLIDWORKS could not save the view image")
        img = PILImage.open(bmp).convert("RGB")
        img.thumbnail((1280, 1280))
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=85)
    return [f"{sw.m0(doc, 'GetTitle')} ({view})", Image(data=buf.getvalue(), format="jpeg")]


# --- Editing --------------------------------------------------------------


@tool
def set_dimension(name: str, value: float) -> dict:
    """Change a dimension (mm, or degrees for angles) and rebuild.
    `name` is as returned by list_dimensions, e.g. "D1@Sketch1"."""
    doc = sw.active_doc()
    dim = doc.Parameter(name)
    if dim is None:
        raise LookupError(f"No dimension {name!r}; use list_dimensions")
    dim.SystemValue = deg(value) if sw.m0(dim, "GetType") == 2 else mm(value)
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
def delete_feature(name: str, keep_sketches: bool = False) -> dict:
    """Delete a feature. Sketches it absorbed (used only by this feature) are
    deleted with it unless keep_sketches=true."""
    doc = sw.active_doc()
    feat = sw.feature(doc, name)
    absorbed = []
    if not keep_sketches:
        for parent in sw.m0(feat, "GetParents") or ():
            if sw.m0(parent, "GetTypeName2") == "ProfileFeature" and len(sw.m0(parent, "GetChildren") or ()) <= 1:
                absorbed.append(parent.Name)
    doc.ClearSelection2(True)
    sw.select_feature(doc, name)
    if not doc.Extension.DeleteSelection2(1):  # swDelete_Absorbed
        raise RuntimeError(f"Could not delete {name!r} (other features may depend on it)")
    deleted = [name]
    for sk in absorbed:
        if doc.FeatureByName(sk) is not None:  # not taken along by the delete above
            doc.ClearSelection2(True)
            sw.select_feature(doc, sk)
            doc.Extension.DeleteSelection2(0)
        if doc.FeatureByName(sk) is None:
            deleted.append(sk)
    _rebuild(doc)
    return {"deleted": deleted}


@tool
def suppress_feature(name: str, suppress: bool = True) -> str:
    """Suppress or unsuppress a feature."""
    doc = sw.active_doc()
    doc.ClearSelection2(True)
    sw.select_feature(doc, name)
    ok = sw.m0(doc, "EditSuppress2") if suppress else sw.m0(doc, "EditUnsuppress2")
    if not ok:
        raise RuntimeError(f"Could not change suppression of {name!r}")
    _rebuild(doc)
    return f"{'suppressed' if suppress else 'unsuppressed'} {name}"


def _material_library() -> list[dict]:
    """Every material in the material databases SOLIDWORKS is configured with."""
    out = []
    try:
        paths = sw.m0(sw.app(), "GetMaterialDatabases") or ()
    except Exception:
        return out  # set_material then just tries the name as given
    for path in paths:
        try:
            root = ET.parse(path).getroot()
        except (OSError, ET.ParseError):
            continue
        for cls in root.iter():
            if cls.tag.rsplit("}", 1)[-1] != "classification":
                continue
            for m in cls:
                if m.tag.rsplit("}", 1)[-1] == "material" and m.get("name"):
                    out.append({"name": m.get("name"), "category": cls.get("name"), "database": Path(path).stem})
    return out


def _norm(text: str) -> str:
    return "".join(c for c in text.lower() if c.isalnum())


@tool
def list_materials(query: str = "") -> list[dict]:
    """Search the material library by name or category (e.g. "6061", "stainless",
    "ABS"); returns the exact names set_material accepts."""
    q = _norm(query)
    return [m for m in _material_library() if q in _norm(m["name"]) or q in _norm(m["category"] or "")][:60]


@tool
def set_material(material: str, database: str = "") -> dict:
    """Assign a material to the active part by its library name, e.g.
    "AISI 304", "6061-T6 (SS)", "ABS". Unknown names fail with suggestions;
    use list_materials to search."""
    doc = sw.active_doc(DOC_PART)
    library = _material_library()
    exact = [m for m in library if m["name"].lower() == material.lower() and (not database or m["database"].lower() == database.lower())]
    if exact:
        material, database = exact[0]["name"], exact[0]["database"]
    elif library:
        q = _norm(material)
        close = [m["name"] for m in library if q and (q in _norm(m["name"]) or _norm(m["name"]) in q)]
        raise LookupError(f"Material {material!r} is not in the library. Closest names: {close[:15] or 'none'}; see list_materials")
    doc.SetMaterialPropertyName2("", database or "SOLIDWORKS Materials", material)
    applied = doc.GetMaterialPropertyName2("", sw.out_bstr())
    if applied != material:
        raise LookupError(f"SOLIDWORKS did not accept material {material!r} from {database!r}")
    return {"material": applied, "database": database, "density_kg_m3": round(sw.m0(doc.Extension, "CreateMassProperty").Density, 3)}


@tool
def custom_properties(set_values: dict[str, str] | None = None) -> dict[str, str]:
    """Read the document's custom properties, after first adding/replacing any
    given in set_values (e.g. {"PartNo": "A-001", "Description": "Bracket"})."""
    cpm = sw.active_doc().Extension.CustomPropertyManager("")
    for k, v in (set_values or {}).items():
        cpm.Add3(k, 30, str(v), 2)  # text, replace existing
    return {n: cpm.Get(n) for n in (sw.m0(cpm, "GetNames") or ())}


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


def _flatten(created: list[Any]) -> list[Any]:
    out = []
    for c in created:
        if isinstance(c, (tuple, list)):
            out.extend(x for x in c if x is not None)
        elif c is not None and not isinstance(c, bool):
            out.append(c)
    return out


def _datum_lines(segments: list[Any]) -> tuple[Any, Any]:
    """The leftmost vertical and lowest horizontal solid lines, used as
    dimension datums so a rectangle gets its overall width/height and
    features are located from its edges, as on a drawing."""
    left = bottom = None
    for seg in segments:
        try:
            if sw.m0(seg, "GetType") != 0 or seg.ConstructionGeometry:  # swSketchLINE
                continue
            a, b = sw.m0(seg, "GetStartPoint2"), sw.m0(seg, "GetEndPoint2")
            (x1, y1), (x2, y2) = (a.X, a.Y), (b.X, b.Y)
        except Exception:
            continue
        if abs(x1 - x2) < 1e-9 and (left is None or x1 < left[0]):
            left = (x1, seg)
        elif abs(y1 - y2) < 1e-9 and (bottom is None or y1 < bottom[0]):
            bottom = (y1, seg)
    return (left[1] if left else None), (bottom[1] if bottom else None)


def _fully_define(sm: Any, segments: list[Any]) -> str:
    """Let SOLIDWORKS add the dimensions/relations that fully define the
    open sketch (so they show up in drawings). Returns the resulting state."""
    sk = sm.ActiveSketch
    left, bottom = _datum_lines(segments)
    datums = [(left or sw.nothing(), bottom or sw.nothing())]
    if left is not None or bottom is not None:
        datums.append((sw.nothing(), sw.nothing()))  # fall back to the origin
    # FullyDefineSketch(Relations, Dimensions, RelationsToAdd, VerticalAndHorizontal,
    #   HorizontalScheme, HorizontalDatum, VerticalScheme, VerticalDatum,
    #   HorizontalPlacement, VerticalPlacement); datum Nothing = sketch origin.
    # Enum values are tried from baseline/below/left down to all-zero defaults.
    error = None
    for (hdatum, vdatum), (scheme, hplace, vplace) in ((d, e) for d in datums for e in ((1, 1, 1), (0, 0, 0))):
        try:
            sm.FullyDefineSketch(True, True, 1023, True, scheme, hdatum, scheme, vdatum, hplace, vplace)
            error = None
            break
        except Exception as exc:
            error = exc
    if error is not None:
        return f"not dimensioned ({error})"
    states = {1: "unknown", 2: "under_defined", 3: "fully_defined", 4: "over_defined"}
    try:
        return states.get(sw.m0(sk, "GetConstrainedStatus"), "unknown")
    except Exception:
        return "unknown"


def _sketch_normal_axis(feat: Any) -> int | None:
    """Model axis (0=x, 1=y, 2=z) a finished sketch faces, or None if tilted."""
    try:
        data = sw.m0(sw.m0(sw.m0(feat, "GetSpecificFeature2"), "ModelToSketchTransform"), "ArrayData")
    except Exception:
        return None
    normal = data[6:9]
    axis = max(range(3), key=lambda i: abs(normal[i]))
    return axis if abs(abs(normal[axis]) - 1) < 1e-6 else None


def _select_plane(doc: Any, target: Any) -> None:
    # Name-based selection is what a recorded macro does; Feature.Select2
    # was seen to be ignored by InsertSketch on a freshly created part.
    if doc.Extension.SelectByID2(target.Name, "PLANE", 0, 0, 0, False, 0, sw.nothing(), 0):
        return
    if not target.Select2(False, 0):
        raise RuntimeError(f"Could not select plane {target.Name!r}")


class SketchNotOpened(Exception):
    pass


def _draw(doc: Any, entities: list[dict], dimension: bool) -> tuple[Any, list[int], str | None]:
    """Open a sketch on the current selection, draw, optionally dimension, close."""
    sm = doc.SketchManager
    sm.InsertSketch(True)
    if sm.ActiveSketch is None:
        raise SketchNotOpened()
    sm.AddToDB = True  # no snapping/inference to nearby geometry
    failed: list[int] = []
    created: list[Any] = []
    state = None
    try:
        for i, e in enumerate(entities):
            made = _sketch_entity(sm, e)
            if made is None:
                failed.append(i)
            created.append(made)
        sm.AddToDB = False
        if dimension and len(failed) < len(entities):
            state = _fully_define(sm, _flatten(created))
    finally:
        sm.AddToDB = False
        sm.InsertSketch(True)
    return sw.last_feature_of_type(doc, "ProfileFeature"), failed, state


@tool
def sketch(entities: list[dict], plane: str = "front", face_point: list[float] | None = None, dimension: bool = True) -> dict:
    """Create a sketch and close it. Place it on "front"/"top"/"right", on any
    plane feature by name, or on the face at face_point ([x,y,z] mm).
    Coordinates are 2D sketch coordinates in mm. Entity types:
      {"type":"line","x1","y1","x2","y2"}  {"type":"centerline",...same}
      {"type":"rectangle","x1","y1","x2","y2"}  {"type":"center_rectangle","cx","cy","width","height"}
      {"type":"circle","cx","cy","r"}  {"type":"polygon","cx","cy","r","sides"} (inscribed)
      {"type":"arc","cx","cy","x1","y1","x2","y2","clockwise":false}
      {"type":"slot","x1","y1","x2","y2","width"}  {"type":"point","x","y"}
    Profiles for extrude must be closed; a revolve profile needs one centerline as its axis.
    dimension=true (default) lets SOLIDWORKS add the dimensions that fully
    define the sketch, so they can be changed later and appear in drawings.
    The result's "normal" is the model axis the sketch actually faces."""
    doc = sw.active_doc(DOC_PART)
    target, expected = None, None
    if not face_point:
        key = plane.lower()
        if key in sw.PLANE_NAMES:
            target, expected = sw.standard_plane(doc, key), sw.PLANE_AXIS[key]
        else:
            target = sw.feature(doc, plane)
            expected = sw._plane_normal_axis(target)
    # A freshly created part sometimes is not ready for sketching yet: make
    # sure it is the active window and retry opening the sketch briefly.
    sw.app().ActivateDoc3(sw.m0(doc, "GetTitle"), False, 0, sw.out_int())
    for attempt in range(2):
        for opening in range(3):
            doc.ClearSelection2(True)
            if face_point:
                sw.select_at(doc, "FACE", face_point, False)
            else:
                _select_plane(doc, target)
            try:
                feat, failed, state = _draw(doc, entities, dimension)
                break
            except SketchNotOpened:
                if opening == 2:
                    raise RuntimeError("SOLIDWORKS did not open a sketch on the selected plane/face (3 tries)")
                time.sleep(0.5)
        actual = _sketch_normal_axis(feat)
        if expected is None or actual is None or actual == expected:
            break
        # Landed on the wrong plane: remove it and try once more.
        bad = feat.Name
        doc.ClearSelection2(True)
        sw.select_feature(doc, bad)
        doc.Extension.DeleteSelection2(0)
        if attempt == 1:
            raise RuntimeError(
                f"Sketch kept landing on the plane facing {'xyz'[actual]} instead of "
                f"{'xyz'[expected]} ({target.Name}); deleted {bad}"
            )
    result = {
        "sketch": feat.Name,
        "on": f"face at {face_point} mm" if face_point else target.Name,
        "normal": "xyz"[actual] if actual is not None else "unknown",
        "entities": len(entities) - len(failed),
    }
    if state:
        result["definition"] = state
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
    asm_title = sw.m0(asm, "GetTitle")
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
        {"name": c.Name2, "path": sw.m0(c, "GetPathName"), "suppressed": bool(sw.m0(c, "IsSuppressed")), "fixed": bool(sw.m0(c, "IsFixed"))}
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
    sw.m0(asm, "EditRebuild3")
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
        model_path = sw.m0(src, "GetPathName")
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
    None), m0(obj, "Name") (call a parameterless method: late binding often
    returns e.g. doc.GetTitle as a value, so doc.GetTitle() fails), mm (0.001), deg (degrees -> radians), nothing(), empty(),
    out_int() (ByRef Long; read .value), select_at(doc, "FACE", [x,y,z]mm,
    append), win32com, pythoncom, math. Assign to `result` to return a value;
    print() output is returned too. The API uses metres and radians."""
    app = sw.app()
    env = {
        "swApp": app, "doc": app.ActiveDoc, "mm": MM, "deg": math.radians, "math": math,
        "m0": sw.m0, "nothing": sw.nothing, "empty": sw.empty, "out_int": sw.out_int, "select_at": sw.select_at,
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
