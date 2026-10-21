"""MCP server that drives SOLIDWORKS through its COM API.

Lengths are millimetres and angles degrees in every tool; points are model
coordinates of the active document (sketch tools take 2D sketch coordinates).
Anything the dedicated tools do not cover can be done with run_api_script,
which runs Python against the live API object.
"""

from __future__ import annotations

import contextlib
import functools
import glob
import io
import math
import os
import re
import shutil
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
        "Drives SOLIDWORKS via its API. Units: millimetres and degrees; points are model "
        "coordinates [x, y, z] in mm.\n"
        "Parts: status -> new_document or open_document -> sketch -> extrude/revolve/hole/... "
        "-> model_info (volume, bounding box) or screenshot to check -> save. Drawings: "
        "make_drawing (views, model dimensions, hole callouts) -> more views and annotations "
        "-> title_block -> export to PDF/DXF. Nothing is saved unless you call save.\n"
        "Sketch coordinates (x, y) on the default planes are model (x, y, 0) on front, "
        "(x, 0, -y) on top and (0, y, -x) on right; extrudes go along +z, +y and +x.\n"
        "Faces, edges and vertices are picked by a model point on them (within 0.5 mm; "
        "hidden ones too): take the points from list_faces / list_edges (they also give "
        "normals, axes, radii and the feature that made each face).\n"
        "Chain several tools in one run_steps call (\"$N.key\" passes step N's result on): "
        "much faster than one call per step. Results carry rebuild_errors when a feature "
        "fails to rebuild.\n"
        "Names of planes, features, views and components are localized (the UI may be "
        "Korean: 정면/윗면/우측면, 도면뷰1): take them from list_features, list_views and "
        "list_components instead of guessing.\n"
        "Tools check their result and raise an error that says why; fix the input rather "
        "than repeating the same call. A timeout usually means a dialog is open in "
        "SOLIDWORKS: ask the user to close it. For anything without a dedicated tool, use "
        "run_api_script."
    ),
)


def tool(fn=None, *, timeout: float | None = None):
    """Register fn as a tool. Its body runs on the COM thread, and any failure
    (including COM errors and timeouts) comes back to the agent as a
    readable error. Long-running tools pass a larger timeout (seconds)."""

    def register(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                return call(lambda: _with_rebuild_errors(fn(*args, **kwargs)), timeout)
            except ToolError:
                raise
            except Exception as exc:
                raise ToolError(f"{type(exc).__name__}: {exc}") from exc

        TOOLS[fn.__name__] = fn
        return mcp.tool()(wrapper)

    return register(fn) if fn is not None else register


TOOLS: dict[str, Any] = {}  # tool name -> plain function (runs on the COM thread), for run_steps


def _feature_errors(doc: Any) -> list[dict]:
    """Features SOLIDWORKS flags with rebuild errors or warnings (its
    What's Wrong list)."""
    import pythoncom
    from win32com.client import VARIANT

    if not sw.m0(doc.Extension, "GetWhatsWrongCount"):
        return []
    feats, codes, warns = (VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, None) for _ in range(3))
    doc.Extension.GetWhatsWrong(feats, codes, warns)
    names = sw.enum_names("swFeatureError_e")
    out = []
    for feat, code, warning in zip(feats.value or (), codes.value or (), warns.value or ()):
        name = names.get(code, str(code)).replace("swFeatureError", "") or "Unknown"
        out.append({"feature": feat.Name, "error": name, "warning": bool(warning)})
    return out


def _with_rebuild_errors(result: Any) -> Any:
    """Add the active document's rebuild errors to a tool's dict result, so
    a change that broke a feature (a fillet too big after set_dimension) is
    reported instead of looking like a success."""
    if not isinstance(result, dict) or "rebuild_errors" in result:
        return result
    try:
        doc = sw.app().ActiveDoc
        errors = _feature_errors(doc) if doc is not None else []
    except Exception:
        return result
    return {**result, "rebuild_errors": errors} if errors else result


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
    """SOLIDWORKS version, the active document and how much memory and
    Windows handles SOLIDWORKS uses (it slows down and finally blocks every
    command with a "resources low" dialog when they run out)."""
    app = sw.app()
    doc = app.ActiveDoc
    result = {"revision": sw.m0(app, "RevisionNumber"), "active_document": sw.doc_info(doc) if doc else None,
              "open_documents": len(sw.m0(app, "GetDocuments") or ())}
    try:
        usage = sw.process_usage(sw.m0(app, "GetProcessID"))
    except Exception:
        return result
    result["solidworks_process"] = usage
    if usage["memory_mb"] > 6000 or max(usage["gdi_objects"], usage["user_objects"]) > 7000:
        result["warning"] = ("SOLIDWORKS is using a lot of memory/handles; save your work and restart it "
                             "before long jobs (it froze behind a 'Windows resources low' dialog at ~9 GB)")
    return result


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
    app = sw.app()
    target = os.path.normcase(os.path.abspath(str(Path(path))))
    for d in sw.m0(app, "GetDocuments") or ():
        p = sw.m0(d, "GetPathName")
        if p and os.path.normcase(os.path.abspath(p)) == target:
            # Already open: OpenDoc6 would return it without switching windows.
            active = app.ActivateDoc3(sw.m0(d, "GetTitle"), False, 0, sw.out_int())
            return {**sw.doc_info(active or d), "already_open": True}
    errors, warnings = sw.out_int(), sw.out_int()
    doc = app.OpenDoc6(str(Path(path)), sw.DOC_TYPES[ext], 1, "", errors, warnings)  # 1 = silent
    if doc is None:
        raise RuntimeError(f"Open failed: {sw.load_error(errors.value)} - {path}")
    active = app.ActivateDoc3(sw.m0(doc, "GetTitle"), False, 0, sw.out_int())
    return {**sw.doc_info(active or doc), "warnings": warnings.value}


@tool
def activate_document(title: str) -> dict:
    """Bring an open document (by title, as in list_documents) to the front."""
    app = sw.app()
    docs = {sw.m0(d, "GetTitle"): d for d in sw.m0(app, "GetDocuments") or ()}
    # exact title first; "bracket_hw" also names "bracket_hw.SLDPRT" when extensions are shown
    exact = title if title in docs else next((t for t in docs if Path(t).stem == title and not t.startswith(f"{title} - ")), None)
    if exact is None:
        raise LookupError(f"No open document titled {title!r}; open: {sorted(docs)}")
    for _ in range(2):  # SOLIDWORKS sometimes leaves the previous window in front
        doc = app.ActivateDoc3(exact, False, 0, sw.out_int())
        active = sw.m0(app, "ActiveDoc")
        if active is not None and sw.m0(active, "GetTitle") == exact:
            return sw.doc_info(active)
    raise RuntimeError(f"Could not bring {exact!r} to the front; the active document is "
                       f"{sw.m0(active, 'GetTitle') if active is not None else None!r}")


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


PDF_PRINTER = "Microsoft Print to PDF"
# Windows paper sizes (DMPAPER_*) by sheet size in mm, long side first
PAPER_SIZES = {(297, 210): 9, (420, 297): 8, (594, 420): 66, (279, 216): 1, (432, 279): 3}


def _utf8_codepage() -> bool:
    """Whether Windows runs with "Beta: Use Unicode UTF-8" (ANSI code page
    65001). SOLIDWORKS' PDF export then leaves out every text run with
    Korean (or other non-Latin) characters, though the screen shows them."""
    import ctypes

    return ctypes.windll.kernel32.GetACP() == 65001


def _print_pdf(drw: Any, target: str, before: float | None) -> None:
    """Print the drawing to "Microsoft Print to PDF", scaled to fit paper of
    the sheet's size; Windows' printing keeps the text SaveAs drops."""
    import pythoncom
    from win32com.client import VARIANT

    props = sw.m0(sw.m0(drw, "GetCurrentSheet"), "GetProperties2")  # ..., width, height (m)
    w, h = round(props[5] * 1000), round(props[6] * 1000)
    paper = PAPER_SIZES.get((max(w, h), min(w, h)), 8)
    ps = drw.PageSetup
    saved = (ps.ScaleToFit, ps.Orientation, ps.PrinterPaperSize)
    ps.ScaleToFit, ps.Orientation, ps.PrinterPaperSize = True, 2 if w >= h else 1, paper
    try:
        pages = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [1, max(1, sw.m0(drw, "GetSheetCount"))])
        drw.Extension.PrintOut2(pages, 1, False, PDF_PRINTER, target)
    finally:
        ps.ScaleToFit, ps.Orientation, ps.PrinterPaperSize = saved
    # the spooler writes the file after PrintOut2 returns
    size, deadline = -1, time.time() + 60
    while time.time() < deadline:
        if sw.file_written(target, before):
            now = os.path.getsize(target)
            if now > 0 and now == size:
                return
            size = now
        time.sleep(0.5)
    raise RuntimeError(f"{PDF_PRINTER} did not write {target}")


@tool
def export(path: str, pdf_method: Literal["auto", "export", "print"] = "auto") -> dict:
    """Export a copy of the active document; the format follows the extension
    (.step/.stp, .igs, .x_t, .stl, .3mf, .pdf, .dxf/.dwg (drawings), .png/.jpg,
    .eDrawings). The open document keeps its own file.

    Drawing PDFs: with Windows' UTF-8 code page SOLIDWORKS' own PDF export
    drops Korean text (e.g. "관통" in hole callouts), so "auto" then prints
    to "Microsoft Print to PDF" instead; "export"/"print" force a method."""
    doc = sw.active_doc()
    errors, warnings = sw.out_int(), sw.out_int()
    target = str(Path(path))
    _check_not_open_elsewhere(doc, target)
    before = sw.mtime(target)
    if target.lower().endswith(".pdf") and sw.m0(doc, "GetType") == DOC_DRAWING:
        import win32print

        printers = [p[2] for p in win32print.EnumPrinters(2 | 4)]  # local and network printers
        use_print = pdf_method == "print" or (pdf_method == "auto" and _utf8_codepage() and PDF_PRINTER in printers)
        if use_print:
            if PDF_PRINTER not in printers:
                raise RuntimeError(f'Printer "{PDF_PRINTER}" is not installed')
            _print_pdf(doc, target, before)
            return {"path": target, "bytes": os.path.getsize(target), "method": "print"}
    ok = doc.Extension.SaveAs3(target, 0, 1 | 2, sw.nothing(), sw.nothing(), errors, warnings)  # silent | copy
    if not ok or not sw.file_written(target, before):
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


def _dim_kind(dim: Any) -> str:
    """"linear", "angular" or "integer" (e.g. pattern counts), from
    swDimensionParamType_e (0, 1, 2 as read from swconst.tlb in 2025)."""
    kind = sw.m0(dim, "GetType")
    try:
        angular, integer = sw.constant("swDimensionParamTypeDoubleAngular"), sw.constant("swDimensionParamTypeInteger")
    except Exception:
        angular, integer = 1, 2  # swconst unreadable (tests without SOLIDWORKS)
    return {angular: "angular", integer: "integer"}.get(kind, "linear")


def _dim_value(dim: Any) -> tuple[float, str]:
    kind = _dim_kind(dim)
    if kind == "angular":
        return round(math.degrees(dim.SystemValue), 6), "deg"
    if kind == "integer":
        return round(dim.SystemValue, 6), ""
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
    out, seen = [], set()
    for f in feats:
        dd = sw.m0(f, "GetFirstDisplayDimension")
        while dd is not None:
            dim = dd.GetDimension2(0)
            value, unit = _dim_value(dim)
            name = "@".join(dim.FullName.split("@")[:2])
            if name not in seen:  # an extrude also shows its sketch's dimensions
                seen.add(name)
                out.append({"name": name, "value": value, "unit": unit, "feature": name.split("@")[1]})
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
        box = _body_box(doc)
        if box:
            info["bounding_box_mm"] = {"min": [round(v / MM, 4) for v in box[:3]], "max": [round(v / MM, 4) for v in box[3:]]}
        db = sw.out_bstr()
        info["material"] = doc.GetMaterialPropertyName2("", db) or None
    return info


VIEWS = {"front": 1, "back": 2, "left": 3, "right": 4, "top": 5, "bottom": 6, "isometric": 7, "trimetric": 8, "dimetric": 9}


def _body_box(doc: Any) -> list[float] | None:
    """Box (m) around the part's solid bodies; GetPartBox also counts
    sketches and reference geometry."""
    try:
        boxes = [b.GetBodyBox() for b in doc.GetBodies2(0, True) or ()]
    except Exception:
        boxes = []  # bodies that cannot report a box: fall back to the part box
    if not boxes:
        return doc.GetPartBox(True)
    return [min(b[i] for b in boxes) for i in range(3)] + [max(b[i] for b in boxes) for i in range(3, 6)]


def _mm3(p: Any) -> list[float]:
    return [round(v / MM, 3) + 0.0 for v in p]


def _unit3(v: Any) -> list[float]:
    return [round(c, 4) + 0.0 for c in v]


class _Placement:
    """Maps a component's geometry (in its own part space) into the
    assembly: ArrayData holds the component's X, Y, Z axes, then its origin."""

    def __init__(self, comp: Any = None):
        data = list(sw.m0(comp.Transform2, "ArrayData")) if comp is not None else None
        self.axes = [data[0:3], data[3:6], data[6:9]] if data else None
        self.origin = data[9:12] if data else None

    def direction(self, v: Any) -> list[float]:
        if self.axes is None:
            return list(v)
        return [sum(v[i] * self.axes[i][k] for i in range(3)) for k in range(3)]

    def point(self, p: Any) -> list[float]:
        if self.axes is None:
            return list(p)
        return [c + o for c, o in zip(self.direction(p), self.origin)]


def _bodies_with_placement(doc: Any, component: str) -> list[tuple[str | None, Any, _Placement]]:
    if sw.m0(doc, "GetType") == DOC_PART:
        return [(None, b, _Placement()) for b in doc.GetBodies2(0, False) or ()]
    comps = [_component(doc, component)] if component else list(doc.GetComponents(False) or ())
    out = []
    for comp in comps:
        if sw.m0(comp, "IsSuppressed"):
            continue
        place = _Placement(comp)
        out.extend((comp.Name2, b, place) for b in comp.GetBodies2(0) or ())
    return out


def _inside_point(face: Any) -> list[float] | None:
    """A point well inside a face (m): the centre of its largest display
    triangle, so a hole or a narrow strip of the face does not mislead."""
    tri = face.GetTessTriangles(True)
    best, area = None, -1.0
    for i in range(0, len(tri or ()) - 8, 9):
        a, b, c = tri[i:i + 3], tri[i + 3:i + 6], tri[i + 6:i + 9]
        u = [b[k] - a[k] for k in range(3)]
        v = [c[k] - a[k] for k in range(3)]
        cross = [u[1] * v[2] - u[2] * v[1], u[2] * v[0] - u[0] * v[2], u[0] * v[1] - u[1] * v[0]]
        size = math.sqrt(sum(x * x for x in cross))
        if size > area:
            area, best = size, [(a[k] + b[k] + c[k]) / 3 for k in range(3)]
    if best is None:  # no display data: the face point nearest its box centre
        box = sw.m0(face, "GetBox")
        best = list(face.GetClosestPointOn((box[0] + box[3]) / 2, (box[1] + box[4]) / 2, (box[2] + box[5]) / 2))[:3]
    return best


def _face_info(face: Any, place: _Placement) -> dict:
    surf = sw.m0(face, "GetSurface")
    feat = sw.m0(face, "GetFeature")
    info: dict = {"type": "other"}
    if sw.invoke(surf, "IsPlane"):
        info = {"type": "plane", "normal": _unit3(place.direction(face.Normal))}
    elif sw.invoke(surf, "IsCylinder"):
        p = list(surf.CylinderParams)
        info = {"type": "cylinder", "radius_mm": round(p[6] / MM, 3), "axis": _unit3(place.direction(p[3:6])),
                "axis_point_mm": _mm3(place.point(p[0:3]))}
    elif sw.invoke(surf, "IsCone"):
        p = list(surf.ConeParams)
        info = {"type": "cone", "axis": _unit3(place.direction(p[3:6])), "apex_side_point_mm": _mm3(place.point(p[0:3]))}
    elif sw.invoke(surf, "IsSphere"):
        p = list(surf.SphereParams)
        info = {"type": "sphere", "radius_mm": round(p[3] / MM, 3), "center_mm": _mm3(place.point(p[0:3]))}
    elif sw.invoke(surf, "IsTorus"):
        info = {"type": "torus"}
    point = _inside_point(face)
    info["point_mm"] = _mm3(place.point(point)) if point else None
    info["area_mm2"] = round(sw.m0(face, "GetArea") / MM**2, 3)
    info["feature"] = feat.Name if feat is not None else None
    return info


@tool
def list_faces(feature: str = "", type: Literal["", "plane", "cylinder", "cone", "sphere", "torus", "other"] = "",
               component: str = "", limit: int = 60) -> dict:
    """Faces of the active part (or, in an assembly, of every component or
    one `component`, in assembly coordinates), each with point_mm: a point
    well inside the face to pass to fillet, sketch(face_point=...),
    add_mate and the other tools that pick by point. Planes give their
    outward normal, cylinders their radius and axis. Filter by the feature
    that made the faces (names from list_features) and/or by type."""
    doc = sw.active_doc()
    if sw.m0(doc, "GetType") == DOC_DRAWING:
        raise ValueError("Open the part or assembly; drawings have views, not faces")
    faces, total = [], 0
    for comp, body, place in _bodies_with_placement(doc, component):
        for face in body.GetFaces() or ():
            info = _face_info(face, place)
            if (feature and info["feature"] != feature) or (type and info["type"] != type):
                continue
            total += 1
            if len(faces) < limit:
                if comp:
                    info["component"] = comp
                faces.append(info)
    return {"faces": faces, "total": total, **({"truncated": True} if total > len(faces) else {})}


@tool
def list_edges(feature: str = "", type: Literal["", "line", "circle", "other"] = "", component: str = "",
               limit: int = 80) -> dict:
    """Edges of the active part (or of the components of an assembly, in
    assembly coordinates) with point_mm, the edge's midpoint to pass to
    fillet/chamfer and other edge picks, its length, and for circles the
    centre, radius and axis. feature= keeps edges of faces that feature
    made (e.g. the top edges of a boss)."""
    doc = sw.active_doc()
    if sw.m0(doc, "GetType") == DOC_DRAWING:
        raise ValueError("Open the part or assembly; drawings have views, not edges")
    edges, total = [], 0
    for comp, body, place in _bodies_with_placement(doc, component):
        for edge in body.GetEdges() or ():
            if feature:
                owners = {getattr(sw.m0(f, "GetFeature"), "Name", None) for f in sw.m0(edge, "GetTwoAdjacentFaces2") or () if f is not None}
                if feature not in owners:
                    continue
            curve = sw.m0(edge, "GetCurve")
            params = sw.m0(edge, "GetCurveParams2")
            t0, t1 = params[6], params[7]
            kind = "line" if sw.invoke(curve, "IsLine") else "circle" if sw.invoke(curve, "IsCircle") else "other"
            if type and kind != type:
                continue
            total += 1
            if len(edges) >= limit:
                continue
            info = {"type": kind, "point_mm": _mm3(place.point(list(curve.Evaluate2((t0 + t1) / 2, 0))[:3])),
                    "length_mm": round(curve.GetLength3(t0, t1) / MM, 3)}
            if kind == "line":
                info["ends_mm"] = [_mm3(place.point(params[0:3])), _mm3(place.point(params[3:6]))]
            elif kind == "circle":
                c = list(curve.CircleParams)
                info.update({"center_mm": _mm3(place.point(c[0:3])), "radius_mm": round(c[6] / MM, 3),
                             "axis": _unit3(place.direction(c[3:6]))})
            if comp:
                info["component"] = comp
            edges.append(info)
    return {"edges": edges, "total": total, **({"truncated": True} if total > len(edges) else {})}


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
def set_dimension(name: str, value: float, configurations: Literal["this", "all"] = "this") -> dict:
    """Change a dimension (mm, or degrees for angles) and rebuild.
    `name` is as returned by list_dimensions, e.g. "D1@Sketch1". By default
    only the active configuration changes; configurations="all" changes
    every configuration. A dimension driven by an equation must be changed
    through set_equation instead."""
    doc = sw.active_doc()
    dim = doc.Parameter(name)
    if dim is None:
        raise LookupError(f"No dimension {name!r}; use list_dimensions")
    kind = _dim_kind(dim)
    system = deg(value) if kind == "angular" else float(round(value)) if kind == "integer" else mm(value)
    which = 1 if configurations == "this" else 2  # swSetValue_InThisConfiguration / _InAllConfigurations
    status = _call_first(dim, [
        ("SetSystemValue3", (system, which, sw.nothing())),
        ("SetSystemValue2", (system, which)),
    ])
    if status not in (0, None):  # swSetValue_Successful = 0
        raise RuntimeError(f"SOLIDWORKS refused the value (swSetValueReturnStatus {status}); is it driven by an equation?")
    _rebuild(doc)
    new, unit = _dim_value(dim)
    return {
        "name": name, "value": new, "unit": unit,
        "configuration": doc.ConfigurationManager.ActiveConfiguration.Name if configurations == "this" else "all",
    }


@tool
def rebuild() -> dict:
    """Force a full rebuild of the active document and list the features
    that fail (SOLIDWORKS' What's Wrong list: feature, error, warning).
    Every other tool adds the same list as rebuild_errors when there is any."""
    doc = sw.active_doc()
    _rebuild(doc)
    return {"rebuilt": sw.m0(doc, "GetTitle"), "rebuild_errors": _feature_errors(doc)}


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
    sw.close_open_sketch(doc)  # a sketch left in edit mode blocks deletion
    feat = sw.feature(doc, name)
    absorbed = []
    if not keep_sketches:
        for parent in sw.m0(feat, "GetParents") or ():
            if sw.m0(parent, "GetTypeName2") in ("ProfileFeature", "3DProfileFeature") and len(sw.m0(parent, "GetChildren") or ()) <= 1:
                absorbed.append(parent.Name)
    doc.ClearSelection2(True)
    sw.select_feature(doc, name)
    if not doc.Extension.DeleteSelection2(1):  # swDelete_Absorbed
        raise RuntimeError(f"SOLIDWORKS refused to delete {name!r}; check list_features and a screenshot for what still uses it")
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
def suppress_feature(name: str, suppress: bool = True, configurations: Literal["this", "all"] = "this") -> str:
    """Suppress or unsuppress a feature, in the active configuration only
    (default) or in all configurations."""
    doc = sw.active_doc()
    feat = sw.feature(doc, name)
    # swSuppressFeature / swUnSuppressFeature; swThisConfiguration / swAllConfiguration
    ok = _call_first(feat, [("SetSuppression2", (0 if suppress else 1, 1 if configurations == "this" else 2, sw.nothing()))])
    if not ok:
        raise RuntimeError(f"Could not change suppression of {name!r}")
    _rebuild(doc)
    return f"{'suppressed' if suppress else 'unsuppressed'} {name} ({configurations} configuration{'s' if configurations == 'all' else ''})"


@tool
def rename_feature(name: str, new_name: str) -> dict:
    """Rename a feature, sketch, plane or mate (names from list_features or
    list_mates), e.g. to give the tree meaningful names."""
    doc = sw.active_doc()
    feat = sw.feature(doc, name)
    if new_name != name and doc.FeatureByName(new_name) is not None:
        raise ValueError(f"{new_name!r} is already used in this document")
    feat.Name = new_name
    if feat.Name != new_name:
        raise RuntimeError(f"SOLIDWORKS kept the name {feat.Name!r}; the new name may contain characters it does not allow")
    return {"renamed": name, "name": feat.Name}


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


LENGTH_UNITS = {"mm": "swMM", "cm": "swCM", "m": "swMETER", "in": "swINCHES"}


@tool
def document_units(length: Literal["mm", "cm", "m", "in", ""] = "", decimals: int | None = None,
                   angle_decimals: int | None = None) -> dict:
    """Read, and optionally change, the active document's length unit and
    the decimal places shown for lengths and angles (values in the model
    and the dimensions of a drawing), e.g. decimals=0 for whole-mm
    drawings. Tools always take mm whatever the document unit is."""
    doc = sw.active_doc()
    ext = doc.Extension
    pref = lambda name: sw.constant(name)  # noqa: E731
    if length:
        ext.SetUserPreferenceInteger(pref("swUnitSystem"), 0, pref("swUnitSystem_Custom"))
        ext.SetUserPreferenceInteger(pref("swUnitsLinear"), 0, pref(LENGTH_UNITS[length]))
    if decimals is not None:
        ext.SetUserPreferenceInteger(pref("swUnitsLinearDecimalPlaces"), 0, decimals)
        ext.SetUserPreferenceInteger(pref("swDetailingLinearDimPrecision"), 0, decimals)
    if angle_decimals is not None:
        ext.SetUserPreferenceInteger(pref("swUnitsAngularDecimalPlaces"), 0, angle_decimals)
        ext.SetUserPreferenceInteger(pref("swDetailingAngularDimPrecision"), 0, angle_decimals)
    if length or decimals is not None or angle_decimals is not None:
        doc.ForceRebuild3(False)
    read = lambda name: ext.GetUserPreferenceInteger(pref(name), 0)  # noqa: E731
    unit = sw.enum_names("swLengthUnit_e").get(read("swUnitsLinear"), "")
    result = {
        "length": next((k for k, v in LENGTH_UNITS.items() if v == unit), unit),
        "unit_system": sw.enum_names("swUnitSystem_e").get(read("swUnitSystem"), "").replace("swUnitSystem_", ""),
        "decimals": read("swUnitsLinearDecimalPlaces"),
        "dimension_decimals": read("swDetailingLinearDimPrecision"),
        "angle_decimals": read("swUnitsAngularDecimalPlaces"),
    }
    if (length and result["length"] != length) or (decimals is not None and result["decimals"] != decimals):
        result["problem"] = "SOLIDWORKS kept a different setting"
    return result


# --- Part modelling -------------------------------------------------------


def _sketch_entity(sm: Any, e: dict, doc: Any = None) -> Any:
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
    if t == "ellipse":
        cx, cy = g("cx", 0.0), g("cy", 0.0)
        a, b, ang = g("major_radius"), g("minor_radius"), math.radians(float(e.get("angle", 0)))
        return sm.CreateEllipse(cx, cy, 0, cx + a * math.cos(ang), cy + a * math.sin(ang), 0,
                                cx - b * math.sin(ang), cy + b * math.cos(ang), 0)
    if t == "spline":
        pts = [c for x, y in e["points"] for c in (mm(x), mm(y), 0.0)]
        return sm.CreateSpline2(sw.doubles(pts), True)
    if t == "text":
        # IModelDoc2.InsertSketchText (not the sketch manager); x, y = lower-left corner
        st = doc.InsertSketchText(g("x", 0.0), g("y", 0.0), 0, str(e["text"]), 0, 0, 0, 100, 100)
        if st is not None and "height" in e:
            try:
                fmt = sw.m0(st, "GetTextFormat")
                fmt.CharHeight = g("height")
                st.SetTextFormat(False, fmt)
            except Exception as exc:
                raise RuntimeError(f"text placed, but its height could not be set: {exc}") from exc
        return st
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


def _rect_sides(segments: list[Any]) -> dict[str, Any]:
    """left/right/bottom/top solid lines of a rectangle's segments."""
    vert, horiz = [], []
    for seg in segments:
        try:
            if sw.m0(seg, "GetType") != 0 or seg.ConstructionGeometry:
                continue
            a, b = sw.m0(seg, "GetStartPoint2"), sw.m0(seg, "GetEndPoint2")
        except Exception:
            continue
        if abs(a.X - b.X) < 1e-9:
            vert.append((a.X, seg))
        elif abs(a.Y - b.Y) < 1e-9:
            horiz.append((a.Y, seg))
    if len(vert) < 2 or len(horiz) < 2:
        return {}
    vert.sort(key=lambda v: v[0])
    horiz.sort(key=lambda h: h[0])
    return {"left": vert[0][1], "right": vert[-1][1], "bottom": horiz[0][1], "top": horiz[-1][1]}


def _close(a: float, b: float) -> bool:
    return abs(a - b) < 1e-9


INPUT_DIM_ON_CREATE = 10  # swUserPreferenceToggle_e.swInputDimValOnCreate


def _to_model(sk: Any, x: float, y: float) -> tuple[float, float, float]:
    """Sketch coordinates (m) to model coordinates, for placing dimension text."""
    try:
        mu = sw.m0(sw.app(), "GetMathUtility")
        pt = sw.invoke(mu, "CreatePoint", sw.doubles((x, y, 0.0)))
        inv = sw.m0(sw.m0(sk, "ModelToSketchTransform"), "Inverse")
        return tuple(sw.m0(sw.invoke(pt, "MultiplyTransform", inv), "ArrayData"))[:3]
    except Exception:
        return (x, y, 0.0)


def _symmetry(doc: Any, sm: Any, rects: list, circles: list, select: Any, notes: list[str]) -> None:
    """Keep geometry drawn symmetric about the sketch origin symmetric when
    dimensions change: rectangles centred on the origin and holes in mirror
    pairs get symmetric relations about fixed construction centerlines, so a
    wider plate keeps its hole pattern centred."""
    v_pairs, h_pairs, sym_rects = [], [], []
    for spec, sides in rects:
        if not sides:
            continue
        t = spec.get("type")
        centred = (
            t == "center_rectangle" and _close(spec.get("cx", 0), 0) and _close(spec.get("cy", 0), 0)
        ) or (
            t == "rectangle" and _close(spec["x1"], -spec["x2"]) and _close(spec["y1"], -spec["y2"])
        )
        if centred:
            sym_rects.append(sides)
    for a in circles:
        if a[0] < 0:  # mirror partner across the vertical axis
            b = next((c for c in circles if _close(c[0], -a[0]) and _close(c[1], a[1]) and _close(c[2], a[2])), None)
            if b is not None:
                v_pairs.append((a, b))
            # across the horizontal axis, only on the left half: the right half
            # follows from the vertical pairs, and more would over-define
            if a[1] < 0:
                c = next((c for c in circles if _close(c[0], a[0]) and _close(c[1], -a[1]) and _close(c[2], a[2])), None)
                if c is not None:
                    h_pairs.append((a, c))
    on_v = [c for c in circles if _close(c[0], 0)]  # centred left-right
    on_h = [c for c in circles if _close(c[1], 0)]  # centred top-bottom
    if not (sym_rects or v_pairs or h_pairs or on_v or on_h):
        return
    # centerlines a little longer than the geometry they mirror
    extent = [abs(c[0]) + c[2] for c in circles] + [abs(c[1]) + c[2] for c in circles]
    for spec, _ in rects:
        if spec.get("type") == "center_rectangle":
            extent += [mm(spec["width"]) / 2, mm(spec["height"]) / 2]
        else:
            extent += [abs(mm(spec[k])) for k in ("x1", "y1", "x2", "y2")]
    reach = 1.2 * max(extent, default=0.05) + 5 * MM
    axes = {}
    try:
        for name, (x1, y1, x2, y2) in (("v", (0, -reach, 0, reach)), ("h", (-reach, 0, reach, 0))):
            line = sm.CreateCenterLine(x1, y1, 0, x2, y2, 0)
            if line is None or not select(line):
                raise RuntimeError(f"could not create the {name} centerline")
            doc.SketchAddConstraints("sgFIXED")
            axes[name] = line
        for sides in sym_rects:
            for a, b, axis in (("left", "right", "v"), ("bottom", "top", "h")):
                if select(sides[a], sides[b], axes[axis]):
                    doc.SketchAddConstraints("sgSYMMETRIC")
        for pairs, axis in ((v_pairs, "v"), (h_pairs, "h")):
            for a, b in pairs:
                if select(sw.m0(a[3], "GetCenterPoint2"), sw.m0(b[3], "GetCenterPoint2"), axes[axis]):
                    doc.SketchAddConstraints("sgSYMMETRIC")
        for group, axis in ((on_v, "v"), (on_h, "h")):
            for c in group:
                if select(sw.m0(c[3], "GetCenterPoint2"), axes[axis]):
                    doc.SketchAddConstraints("sgCOINCIDENT")
    except Exception as exc:
        notes.append(f"symmetry: {exc}")
    doc.ClearSelection2(True)


def _dimension_intent(doc: Any, sm: Any, entities: list[dict], created: list[Any]) -> tuple[int, list[str]]:
    """Add the dimensions a designer would: overall width/height of
    rectangles, one diameter per hole size (equal relations for the rest)
    and the spacing between the outermost holes. FullyDefineSketch then only
    has to locate things. Returns (dimensions added, problems)."""
    sk = sm.ActiveSketch
    off = 8 * MM  # how far dimension text sits from the geometry
    added, notes = 0, []

    def place(kind: str, x: float, y: float) -> Any:
        nonlocal added
        add = {"size": doc.AddDimension2, "h": doc.AddHorizontalDimension2, "v": doc.AddVerticalDimension2}[kind]
        dd = add(*_to_model(sk, x, y))
        if dd is not None:
            added += 1
        else:
            notes.append(f"{kind} dimension at ({x / MM:.1f}, {y / MM:.1f}) mm not added")
        doc.ClearSelection2(True)
        return dd

    def select(*ents: Any) -> bool:
        doc.ClearSelection2(True)
        return all(sw._select_entity(doc, e, i > 0, 0) for i, e in enumerate(ents))

    circles, rects = [], []
    for spec, made in zip(entities, created):
        try:
            if spec.get("type") in ("rectangle", "center_rectangle") and made:
                rects.append((spec, _rect_sides(_flatten([made]))))
                left, bottom = _datum_lines(_flatten([made]))
                if bottom is not None and select(bottom):
                    a, b = sw.m0(bottom, "GetStartPoint2"), sw.m0(bottom, "GetEndPoint2")
                    place("size", (a.X + b.X) / 2, a.Y - off)
                if left is not None and select(left):
                    a, b = sw.m0(left, "GetStartPoint2"), sw.m0(left, "GetEndPoint2")
                    place("size", a.X - off, (a.Y + b.Y) / 2)
            elif spec.get("type") == "circle" and made is not None:
                circles.append((mm(spec.get("cx", 0)), mm(spec.get("cy", 0)), mm(spec["r"]), made))
        except Exception as exc:
            notes.append(f"{spec.get('type')}: {exc}")
            doc.ClearSelection2(True)

    _symmetry(doc, sm, rects, circles, select, notes)

    by_size: dict[float, list] = {}
    for c in circles:
        by_size.setdefault(round(c[2], 9), []).append(c)
    for r, group in by_size.items():
        try:
            cx, cy, _, seg = group[0]
            if select(seg):
                dd = place("size", cx + 1.5 * r, cy + 1.5 * r)
                if dd is not None and len(group) > 1:
                    # The prefix already holds the Ø symbol; keep it after the count.
                    try:
                        prefix = dd.GetText(1) or ""  # swDimensionTextPrefix
                    except Exception:
                        prefix = ""
                    if "<MOD-DIAM>" not in prefix:
                        prefix = "<MOD-DIAM>" + prefix
                    dd.SetText(1, f"{len(group)}X {prefix}")  # "4X Ø6.60"
            if len(group) > 1 and select(*(g[3] for g in group)):
                doc.SketchAddConstraints("sgSAMELENGTH")  # equal radius
                doc.ClearSelection2(True)
        except Exception as exc:
            notes.append(f"circle diameter: {exc}")
            doc.ClearSelection2(True)

    if len(circles) > 1:
        try:
            low = min(c[1] for c in circles)
            row = [c for c in circles if abs(c[1] - low) < 1e-9]  # bottom row
            a, b = min(row, key=lambda c: c[0]), max(row, key=lambda c: c[0])
            if b[0] - a[0] > 1e-9 and select(sw.m0(a[3], "GetCenterPoint2"), sw.m0(b[3], "GetCenterPoint2")):
                place("h", (a[0] + b[0]) / 2, low - a[2] - off)
            left_x = min(c[0] for c in circles)
            col = [c for c in circles if abs(c[0] - left_x) < 1e-9]  # left column
            a, b = min(col, key=lambda c: c[1]), max(col, key=lambda c: c[1])
            if b[1] - a[1] > 1e-9 and select(sw.m0(a[3], "GetCenterPoint2"), sw.m0(b[3], "GetCenterPoint2")):
                place("v", left_x - a[2] - off, (a[1] + b[1]) / 2)
        except Exception as exc:
            notes.append(f"hole spacing: {exc}")
            doc.ClearSelection2(True)
    return added, notes


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
            made = _sketch_entity(sm, e, doc)
            if made is None:
                failed.append(i)
            created.append(made)
        sm.AddToDB = False
        if dimension and len(failed) < len(entities):
            app = sw.app()
            prompt = app.GetUserPreferenceToggle(INPUT_DIM_ON_CREATE)
            app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, False)  # no "Modify" dialog per dimension
            try:
                added, notes = _dimension_intent(doc, sm, entities, created)
            finally:
                app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, prompt)
            state = _fully_define(sm, _flatten(created))
            if notes:
                state = f"{state}; added {added} dimensions, problems: {'; '.join(notes)}"
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
      {"type":"ellipse","cx","cy","major_radius","minor_radius","angle":0}
      {"type":"spline","points":[[x,y],...]}  {"type":"text","x","y","text","height"} (or use engrave)
    Profiles for extrude must be closed; a revolve profile needs one centerline as its axis.
    dimension=true (default) adds design-intent dimensions (overall
    rectangle width/height, one diameter per hole size with equal relations,
    spacing between the outer holes) and lets SOLIDWORKS add the rest needed
    to fully define the sketch, so they can be changed later and appear in
    drawings. If a rectangle is the first entity, its width/height become D1/D2.
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
    sw.close_open_sketch(doc)  # InsertSketch toggles: a leftover open sketch would be closed instead
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


END_CONDITIONS = {"blind": 0, "through_all": 1, "up_to_next": 2, "mid_plane": 6,
                  "up_to_surface": 4, "offset_from_surface": 5}  # swEndConditions_e


@tool
def extrude(
    depth: float = 10.0,
    sketch_name: str = "",
    end: Literal["blind", "through_all", "up_to_next", "mid_plane", "up_to_surface", "offset_from_surface"] = "blind",
    reverse: bool = False,
    cut: bool = False,
    depth2: float = 0.0,
    draft_angle: float = 0.0,
    merge: bool = True,
    face_point: list[float] | None = None,
    up_to_plane: str = "",
    thin: float = 0.0,
    thin_reverse: bool = False,
) -> dict:
    """Boss-extrude (or cut-extrude with cut=true) a sketch, default the most
    recent one. depth2 > 0 also extrudes blind in the opposite direction;
    mid_plane splits `depth` evenly about the sketch plane. up_to_surface
    stops at the face through face_point ([x,y,z] mm; it must cover the
    profile) or at the plane up_to_plane (a name or "front"/"top"/"right");
    offset_from_surface stops `depth` mm short of it. thin > 0 makes a thin feature with that
    wall thickness (mm) instead of a solid: an open profile becomes a wall,
    a closed one a hollow tube (thin_reverse puts the wall on the other side)."""
    doc = sw.active_doc(DOC_PART)
    if end in ("up_to_surface", "offset_from_surface") and not (face_point or up_to_plane):
        raise ValueError(f"end={end!r} needs face_point (a point on the face) or up_to_plane")
    doc.ClearSelection2(True)
    name = sketch_name or sw.last_feature_of_type(doc, "ProfileFeature").Name
    sw.select_feature(doc, name)
    if end in ("up_to_surface", "offset_from_surface"):  # the end face or plane with mark 1
        if up_to_plane:
            key = up_to_plane.lower()
            target = sw.standard_plane(doc, key) if key in sw.PLANE_NAMES else sw.feature(doc, up_to_plane)
            target.Select2(True, 1)
        else:
            sw.select_at(doc, "FACE", face_point, True, 1)
    fm = doc.FeatureManager
    single = depth2 <= 0
    t1, d1 = END_CONDITIONS[end], mm(depth)
    t2, d2 = 0, mm(depth2)
    draft = draft_angle != 0
    da = deg(abs(draft_angle))
    if thin > 0:
        head = (single, False, reverse, t1, t2, d1, d2, draft, draft, draft_angle < 0, draft_angle < 0, da, da,
                False, False, False, False)
        wall = (mm(thin), 0, 0, thin_reverse, False, False, 0)  # one-direction wall, no caps
        if cut:
            feat = fm.FeatureCutThin2(*head, *wall, True, True, 0, 0, False)
        else:
            feat = fm.FeatureExtrusionThin2(*head, merge, *wall, True, True, 0, 0, False)
        return _feature_result(feat, "thin cut" if cut else "thin extrude")
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
def insert_component(path: str, position: list[float] | None = None, rotation: list[float] | None = None) -> dict:
    """Insert a part/assembly file into the active assembly with its origin
    at position [x,y,z] mm (default the assembly origin), turned by rotation
    [about x, about y, about z] degrees (applied in that order)."""
    app = sw.app()
    asm = sw.active_doc(DOC_ASSEMBLY)
    asm_title = sw.m0(asm, "GetTitle")
    ext = Path(path).suffix.lower()
    if ext not in (".sldprt", ".sldasm"):
        raise ValueError("Component must be a .sldprt or .sldasm file")
    errors, warnings = sw.out_int(), sw.out_int()
    was_open = app.GetOpenDocumentByName(str(Path(path))) is not None
    part = app.OpenDoc6(str(Path(path)), sw.DOC_TYPES[ext], 1, "", errors, warnings)
    if part is None:
        raise RuntimeError(f"Could not open {path}: {sw.load_error(errors.value)}")
    asm = app.ActivateDoc3(asm_title, False, 0, sw.out_int())
    x, y, z = (mm(v) for v in (position or [0, 0, 0]))
    comp = asm.AddComponent5(str(Path(path)), 0, "", False, "", x, y, z)
    if not was_open:
        # AddComponent5 needs the file open; its own window is not needed afterwards (the
        # model stays loaded for the assembly). Left open, every insert kept a window and
        # long sessions ran SOLIDWORKS out of Windows resources.
        app.CloseDoc(sw.m0(part, "GetTitle"))
        asm = app.ActivateDoc3(asm_title, False, 0, sw.out_int())
    if comp is None:
        raise RuntimeError(f"Could not insert {path}")
    # AddComponent5 puts the component's bounding box centre at x, y, z: move its origin there
    data = _place_component(asm, comp, position or [0, 0, 0], rotation)
    return {"component": comp.Name2, **_axes_info(data)}


def _place_component(asm: Any, comp: Any, position: list[float] | None, rotation: list[float] | None) -> list[float]:
    """Set a component's origin (mm) and/or orientation; returns the
    transform read back (mates may hold it elsewhere)."""
    data = list(sw.m0(comp.Transform2, "ArrayData"))
    if rotation is not None:
        data[0:9] = _rotation_axes(rotation)
    if position is not None:
        data[9:12] = [mm(v) for v in position]
    mu = sw.m0(sw.app(), "GetMathUtility")
    comp.Transform2 = sw.invoke(mu, "CreateTransform", sw.doubles(data))
    sw.m0(asm, "EditRebuild3")
    return list(sw.m0(comp.Transform2, "ArrayData"))


@tool
def list_components(top_level_only: bool = True) -> list[dict]:
    """Components of the active assembly."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    return [
        {"name": c.Name2, "path": sw.m0(c, "GetPathName"), "suppressed": bool(sw.m0(c, "IsSuppressed")), "fixed": bool(sw.m0(c, "IsFixed"))}
        for c in (asm.GetComponents(top_level_only) or ())
    ]


MATES = {"coincident": 0, "concentric": 1, "perpendicular": 2, "parallel": 3, "tangent": 4, "distance": 5, "angle": 6,
         "symmetric": 8, "gear": 10, "lock": 16}
# mates whose AddMate5 form fails or does not exist: built from a mate data object
DATA_MATES = {"width": "swMateWIDTH", "slot": "swMateSLOT", "hinge": "swMateHINGE", "screw": "swMateSCREW",
              "rack_pinion": "swMateRACKPINION"}
ALIGN = {"aligned": 0, "anti_aligned": 1, "closest": 2}


def _picked(asm: Any, entity: str, point: list[float]) -> Any:
    """The face/edge/vertex nearest an assembly point (mm), as an object."""
    asm.ClearSelection2(True)
    sw.select_at(asm, entity, point, append=False)
    obj = asm.SelectionManager.GetSelectedObject6(1, -1)
    asm.ClearSelection2(True)
    if obj is None:
        raise LookupError(f"Nothing selectable at {point} mm")
    return obj


def _data_mate(asm: Any, mate: str, first: Any, second: Any, more: list[Any], value: float, centered: bool) -> Any:
    import pythoncom
    from win32com.client import VARIANT

    def objects(items: list[Any]) -> Any:
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, items)

    data = asm.CreateMateData(sw.constant(DATA_MATES[mate]))
    if data is None:
        raise RuntimeError(f"SOLIDWORKS could not start a {mate} mate")
    if mate == "width":
        data.WidthSelection = objects([first, second])
        data.TabSelection = objects(more)
        data.ConstraintType = sw.constant("swMateWidth_Centered" if centered else "swMateWidth_Free")
    elif mate == "slot":
        data.EntitiesToMate = objects([first, second])
        data.Constraint = sw.constant("swSlotMateConstraintOption_Centered" if centered else "swSlotMateConstraintOption_Free")
    elif mate == "hinge":
        sw.put_indexed(data, "EntitiesToMate", sw.constant("swHingeMateEntityType_Concentric"), objects([first, second]))
        sw.put_indexed(data, "EntitiesToMate", sw.constant("swHingeMateEntityType_Coincident"), objects(more))
    elif mate == "screw":
        data.EntitiesToMate = objects([first, second])
        data.RevolutionType = sw.constant("swDistancePerRevolution")
        data.RevolutionVal = mm(value)
    else:  # rack_pinion: one plain object per role (an array is refused)
        sw.put_indexed(data, "EntitiesToMate", sw.constant("swRackPinionMateEntityType_Rack"), first)
        sw.put_indexed(data, "EntitiesToMate", sw.constant("swRackPinionMateEntityType_Pinion"), second)
        data.DiameterType = sw.constant("swPinionPitchDiameter")
        data.DiameterVal = mm(value)
    return asm.CreateMate(data)


@tool
def add_mate(
    mate: Literal["coincident", "concentric", "perpendicular", "parallel", "tangent", "distance", "angle",
                  "symmetric", "width", "gear", "lock", "slot", "hinge", "screw", "rack_pinion"],
    point1: list[float],
    point2: list[float],
    entity1: Literal["face", "edge", "vertex"] = "face",
    entity2: Literal["face", "edge", "vertex"] = "face",
    value: float = 0.0,
    align: Literal["aligned", "anti_aligned", "closest"] = "closest",
    flip: bool = False,
    min_value: float | None = None,
    max_value: float | None = None,
    plane: str = "right",
    extra_points: list[list[float]] | None = None,
    ratio: float = 1.0,
    centered: bool = True,
) -> dict:
    """Mate two entities picked by assembly-space points (mm).
    - distance / angle: `value` is the distance (mm) or angle (deg); give
      min_value and max_value too for a limit mate (value is where it starts).
    - symmetric: the two entities mirror each other about `plane`
      ("front"/"top"/"right" or a plane name from list_features).
    - width: point1/point2 are the two width faces (a slot's walls) and
      extra_points one point on each of the tab's two side faces; the tab is
      centred between the walls (centered=false: free between them).
    - slot: point1 on a face of a slot, point2 on the pin (cylinder) in it;
      the pin sits at the slot's middle (centered=false: slides along it).
    - hinge: point1/point2 on the two cylinders that share an axis and
      extra_points on the two faces that touch (pin through a hole, sitting
      on the plate).
    - screw: two cylinders on one axis; `value` = advance per turn (mm).
    - rack_pinion: point1 on the rack's straight edge (entity1="edge"),
      point2 on the pinion's cylinder; `value` = pinion pitch diameter (mm).
    - gear: two cylindrical faces or circular edges; entity2 turns `ratio`
      times per turn of entity1 (e.g. teeth 20/40 -> 0.5).
    - lock: fixes the two components to each other (any face of each)."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    limit = min_value is not None or max_value is not None
    if limit and mate not in ("distance", "angle"):
        raise ValueError("min_value/max_value only apply to distance and angle mates")
    if limit and (min_value is None or max_value is None or not min_value <= value <= max_value):
        raise ValueError("A limit mate needs min_value <= value <= max_value")
    if mate in DATA_MATES:
        if mate in ("width", "hinge") and len(extra_points or []) != 2:
            # width with a cylinder as the tab: refused alone, and given twice the mate moves nothing
            what = "tab's two side faces" if mate == "width" else "two faces that touch"
            raise ValueError(f"A {mate} mate needs extra_points: one point on each of the {what}")
        if mate in ("screw", "rack_pinion") and value <= 0:
            what = "advance per turn" if mate == "screw" else "pinion pitch diameter"
            raise ValueError(f"A {mate} mate needs value > 0 ({what}, mm)")
        if mate == "rack_pinion" and entity1 == "face":
            entity1 = "edge"
        first, second = _picked(asm, entity1, point1), _picked(asm, entity2, point2)
        more = [_picked(asm, "face", p) for p in extra_points or []]
        feat = _data_mate(asm, mate, first, second, more, value, centered)
        if feat is None:
            raise RuntimeError(f"{mate} mate failed; check what the points pick (see this tool's description)")
        sw.m0(asm, "EditRebuild3")
        return {"mate": feat.Name, "type": mate}
    asm.ClearSelection2(True)
    sw.select_at(asm, entity1, point1, append=False, mark=1)
    sw.select_at(asm, entity2, point2, append=True, mark=1)
    if mate == "symmetric":  # the symmetry plane goes in with mark 2
        key = plane.lower()
        target = sw.standard_plane(asm, key) if key in sw.PLANE_NAMES else sw.feature(asm, plane)
        if not target.Select2(True, 2):
            raise RuntimeError(f"Could not select plane {target.Name!r}")
    unit = mm if mate == "distance" else deg
    nominal = unit(value) if mate in ("distance", "angle") else 0
    high, low = (unit(max_value), unit(min_value)) if limit else (nominal, nominal)
    dist = (nominal, high, low) if mate == "distance" else (0, 0, 0)
    ang = (nominal, high, low) if mate == "angle" else (0, 0, 0)
    gear = (0.01 * ratio, 0.01) if mate == "gear" else (1, 1)  # the two sizes, entity1 : entity2
    err = sw.out_int()
    feat = asm.AddMate5(MATES[mate], ALIGN[align], flip, *dist, *gear, *ang, False, False, 0, err)
    if feat is None:
        raise RuntimeError(f"Mate failed (swAddMateError {err.value})")
    sw.m0(asm, "EditRebuild3")
    result = {"mate": feat.Name, "type": mate}
    if limit:
        result["limits"] = [min_value, max_value]
    return result


# --- Drawings -------------------------------------------------------------


CALLOUT_PITCH = 0.024  # m on the sheet between stacked hole callouts


def _hole_callouts(drw: Any) -> dict:
    """Add one hole callout per Hole Wizard feature, on a view where its
    holes appear as circles; SOLIDWORKS adds the instance count ("2X")."""
    edge_kind = sw.constant("swViewEntityType_Edge")
    done: dict[str, str] = {}
    notes: list[str] = []
    view = sw.m0(drw, "GetFirstView")
    view = sw.m0(view, "GetNextView") if view is not None else None  # skip the sheet
    while view is not None:
        stack = 0
        for edge in view.GetVisibleEntities2(sw.nothing(), edge_kind) or ():
            try:
                curve = sw.m0(edge, "GetCurve")
                if not sw.invoke(curve, "IsCircle"):
                    continue
                faces = sw.m0(edge, "GetTwoAdjacentFaces2") or ()
                feats = [sw.m0(f, "GetFeature") for f in faces if f is not None]
                feat = next((f for f in feats if f is not None and sw.m0(f, "GetTypeName2") == "HoleWzd"), None)
                if feat is None or feat.Name in done:
                    continue
                drw.ActivateView(view.Name)
                drw.ClearSelection2(True)
                if not view.SelectEntity(edge, False):
                    continue
                # stack the callout texts beside the view's right edge, one
                # slot per callout tall enough for three lines (hole, seat
                # or thread, countersink)
                outline = sw.m0(view, "GetOutline")  # xmin, ymin, xmax, ymax on the sheet (m)
                x = outline[2] + 0.015
                y = outline[3] - CALLOUT_PITCH * stack
                if drw.AddHoleCallout2(x, y, 0) is not None:
                    done[feat.Name] = view.Name
                    stack += 1
            except Exception as exc:
                notes.append(f"{view.Name}: {exc}")
        view = sw.m0(view, "GetNextView")
    drw.ClearSelection2(True)
    out: dict = {"hole_callouts": done}
    if notes:
        out["callout_problems"] = notes[:5]
    return out


def _annotation_types(marked_only: bool = True) -> int:
    """Marked-for-drawing dimensions (plus the unmarked ones unless
    marked_only) and Hole Wizard locations, looked up in swconst where
    available (32768 = swInsertDimensionsMarkedForDrawing)."""
    types = 32768
    names = ["swInsertDimensionsMarkedForDrawing", "swInsertHoleWizardLocations"]
    if not marked_only:
        names.append("swInsertDimensionsNotMarkedForDrawing")
    for name in names:
        try:
            types |= sw.constant(name)
        except Exception:
            continue  # constant missing in this release: import what we can
    return types


def _import_model_dimensions(drw: Any) -> dict:
    """Import model dimensions into all views and report how many arrived.
    Dimensions marked for drawing come first; if there are none (parts built
    through the API often have no marks), all dimensions are imported, and
    as a last resort view by view, before warning."""
    def run(all_views: bool, marked_only: bool) -> int:
        made = drw.InsertModelAnnotations3(0, _annotation_types(marked_only), all_views, True, False, False)
        return len(made) if made else 0

    # Right after the views are created the call imports nothing, even after
    # a rebuild: the views' geometry is only computed once something asks
    # for it, such as a PDF export or listing a view's visible edges.
    edge_kind = sw.constant("swViewEntityType_Edge")
    view = sw.m0(drw, "GetFirstView")
    view = sw.m0(view, "GetNextView") if view is not None else None  # skip the sheet
    while view is not None:
        view.GetVisibleEntities2(sw.nothing(), edge_kind)
        view = sw.m0(view, "GetNextView")
    source = "marked for drawing"
    count = run(True, True)
    if count == 0:
        source = "all model dimensions"
        count = run(True, False)
    if count == 0:
        view = sw.m0(drw, "GetFirstView")
        view = sw.m0(view, "GetNextView") if view is not None else None  # skip the sheet
        while view is not None:
            drw.ActivateView(view.Name)
            count += run(False, False)
            view = sw.m0(view, "GetNextView")
    out: dict = {"dimensions_imported": count}
    if count:
        out["dimension_source"] = source
    else:
        out["dimension_warning"] = "No model dimensions were imported, marked for drawing or not"
    return out


@tool
def add_hole_callouts() -> dict:
    """Add hole callouts (e.g. "2X Ø6.6 ⌴Ø11 ↧6.4") for every Hole Wizard
    feature to the active drawing."""
    drw = sw.active_doc(DOC_DRAWING)
    result = _hole_callouts(drw)
    drw.ForceRebuild3(False)
    return result


@tool(timeout=300)
def make_drawing(
    model_path: str = "",
    projection: Literal["third_angle", "first_angle"] = "third_angle",
    import_dimensions: bool = True,
    hole_callouts: bool = True,
    centerlines: bool = True,
) -> dict:
    """Create a drawing with the standard three views (plus model dimensions,
    hole callouts for Hole Wizard holes and centerlines on cylinders seen
    from the side) of a saved part/assembly; default the active document.
    Export it as PDF or DXF afterwards with export."""
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
    info = sw.doc_info(drw)
    if import_dimensions:
        info.update(_import_model_dimensions(drw))
    if hole_callouts:
        try:
            info.update(_hole_callouts(drw))
        except Exception as exc:
            info["callout_problems"] = [f"{type(exc).__name__}: {exc}"]
    if centerlines:
        try:
            info["centerlines"] = sum(_centerlines(drw, _drawing_views(drw)).values())
        except Exception as exc:
            info["centerline_problems"] = [f"{type(exc).__name__}: {exc}"]
    drw.ForceRebuild3(False)
    return info


# --- Drawing views, scale and annotations -----------------------------------

# Standard model views in the order IModelDoc2.GetModelViewNames lists them
# (the names are localized, e.g. "*등각 보기"); the English names work in
# English installations only.
STANDARD_VIEWS = ["normal", "front", "back", "left", "right", "top", "bottom", "isometric", "trimetric", "dimetric"]


def _model_view_names(model_path: str, orientation: str) -> list[str]:
    names = [f"*{orientation.capitalize()}"]
    target = os.path.normcase(os.path.abspath(model_path))
    for d in sw.m0(sw.app(), "GetDocuments") or ():
        p = sw.m0(d, "GetPathName")
        if p and os.path.normcase(os.path.abspath(p)) == target:
            local = sw.m0(d, "GetModelViewNames") or ()
            if len(local) >= len(STANDARD_VIEWS):
                names.insert(0, local[STANDARD_VIEWS.index(orientation)])
    return names


def _drawing_views(drw: Any) -> list:
    """Views on the current sheet, without the sheet itself."""
    view = sw.m0(drw, "GetFirstView")
    view = sw.m0(view, "GetNextView") if view is not None else None
    out = []
    while view is not None:
        out.append(view)
        view = sw.m0(view, "GetNextView")
    return out


def _drawing_view(drw: Any, name: str = "") -> Any:
    """A view by name, default the first one (the front view of make_drawing)."""
    views = _drawing_views(drw)
    if not views:
        raise LookupError("The drawing has no views; use make_drawing or drawing_view first")
    if not name:
        return views[0]
    for v in views:
        if v.Name == name:
            return v
    raise LookupError(f"No view {name!r}; views: {[v.Name for v in views]}")


def _to_sheet(view: Any, point_mm: list[float], sketch: bool = False) -> tuple[float, float]:
    """Model point (mm) to sheet coordinates (m) as shown in `view`, or with
    sketch=true to the coordinates of the view's own sketch, which are
    relative to the view and unscaled: what SketchManager takes while the
    view is active."""
    mu = sw.m0(sw.app(), "GetMathUtility")
    pt = sw.invoke(mu, "CreatePoint", sw.doubles([mm(v) for v in point_mm]))
    pt = sw.invoke(pt, "MultiplyTransform", sw.m0(view, "ModelToViewTransform"))
    if sketch:
        pt = sw.invoke(pt, "MultiplyTransform", sw.m0(sw.m0(view, "GetSketch"), "ModelToSketchTransform"))
    x, y, _ = sw.m0(pt, "ArrayData")
    return x, y


def _sheet_size(drw: Any) -> tuple[float, float]:
    return tuple(sw.m0(sw.m0(drw, "GetCurrentSheet"), "GetProperties2")[5:7])


def _free_spot(drw: Any, width: float, height: float, exclude: str = "",
               near: tuple[float, float] | None = None, axis: str = "",
               extra: list | None = None, margin: float = 0.03) -> tuple[float, float]:
    """Centre (m) of a free area of the given size on the sheet, clear of the
    views (and their annotations) by a gap; the free spot closest to `near`
    (default the sheet's top right). axis "x" keeps near's y (search along
    x), "y" keeps near's x: for views SOLIDWORKS keeps aligned. When nothing
    is free, the spot that overlaps the least."""
    sheet_w, sheet_h = _sheet_size(drw)
    boxes = [sw.m0(v, "GetOutline") for v in _drawing_views(drw) if v.Name != exclude] + list(extra or [])
    gap, step = 0.012, 0.005
    hw, hh = width / 2 + gap, height / 2 + gap
    target = near or (sheet_w, sheet_h)
    ys = [target[1]] if axis == "x" else None
    xs = [target[0]] if axis == "y" else None
    y_vals = ys or [margin + height / 2 + i * step for i in range(int((sheet_h - 2 * margin - height) / step) + 1)]
    x_vals = xs or [margin + width / 2 + i * step for i in range(int((sheet_w - 2 * margin - width) / step) + 1)]
    best, best_key = None, (math.inf, math.inf)
    for y in y_vals:
        for x in x_vals:
            # overlap with the obstacles (0 = free), then distance to the target
            overlap = sum(max(0.0, min(x + hw, b[2]) - max(x - hw, b[0])) * max(0.0, min(y + hh, b[3]) - max(y - hh, b[1]))
                          for b in boxes)
            key = (round(overlap, 9), math.hypot(x - target[0], y - target[1]))
            if key < best_key:
                best, best_key = (x, y), key
    return best or (sheet_w - margin - width / 2, sheet_h - margin - height / 2)  # too big for the sheet


def _parse_scale(scale: str) -> tuple[float, float]:
    num, _, den = scale.replace(" ", "").partition(":")
    return float(num), float(den or 1)


def _view_info(view: Any) -> dict:
    x1, y1, x2, y2 = sw.m0(view, "GetOutline")
    return {
        "view": view.Name,
        "scale": "{:g}:{:g}".format(*view.ScaleRatio),
        "center_mm": [round((x1 + x2) / 2 / MM, 1), round((y1 + y2) / 2 / MM, 1)],
        "size_mm": [round((x2 - x1) / MM, 1), round((y2 - y1) / MM, 1)],
    }


def _set_view_scale(view: Any, scale: str) -> None:
    num, den = _parse_scale(scale)
    view.UseSheetScale = False
    view.ScaleRatio = sw.doubles([num, den])


@tool
def list_views() -> list[dict]:
    """Views on the active drawing sheet with their scale, centre and size on
    the sheet (mm), for placing new views and annotations."""
    drw = sw.active_doc(DOC_DRAWING)
    return [_view_info(v) for v in _drawing_views(drw)]


@tool
def drawing_view(
    orientation: Literal["front", "back", "top", "bottom", "left", "right", "isometric", "trimetric", "dimetric"] = "isometric",
    position: list[float] | None = None,
    scale: str = "",
    model_path: str = "",
    exploded: bool = False,
) -> dict:
    """Add a model view (e.g. an isometric view) to the active drawing.
    position is the view centre on the sheet in mm (default a free spot
    beside the other views); scale like "1:2" (default the sheet scale). The
    model is the one the drawing already shows unless model_path is given.
    exploded=true shows an assembly's exploded view (see exploded_view)."""
    drw = sw.active_doc(DOC_DRAWING)
    if not model_path:
        views = _drawing_views(drw)
        if not views:
            raise ValueError("The drawing has no views yet; pass model_path")
        model_path = sw.m0(views[0], "GetReferencedModelName")
    sheet_w, sheet_h = _sheet_size(drw)
    x, y = (mm(position[0]), mm(position[1])) if position else (sheet_w / 2, sheet_h / 2)
    view = None
    for name in _model_view_names(model_path, orientation):
        view = drw.CreateDrawViewFromModelView3(model_path, name, x, y, 0)
        if view is not None:
            break
    if view is None:
        raise RuntimeError(f"Could not create the {orientation} view of {model_path}")
    if scale:
        _set_view_scale(view, scale)
    if exploded:
        sw.invoke(view, "ShowExploded", True)
    drw.ForceRebuild3(False)
    if not position:  # now that its size is known
        x1, y1, x2, y2 = sw.m0(view, "GetOutline")
        view.Position = sw.doubles(_free_spot(drw, x2 - x1, y2 - y1, exclude=view.Name))
        drw.ForceRebuild3(False)
    return _view_info(view)


@tool
def section_view(
    through: list[float],
    direction: Literal["horizontal", "vertical"] = "vertical",
    view: str = "",
    label: str = "A",
    position: list[float] | None = None,
    flip: bool = False,
) -> dict:
    """Section view cut through a model point ([x,y,z] mm) of a drawing view
    (default the first, the front view). The cutting line runs vertically or
    horizontally across that view; flip reverses the viewing direction.
    position is the new view's centre on the sheet (mm); by default it is
    placed in line with the parent view."""
    drw = sw.active_doc(DOC_DRAWING)
    parent = _drawing_view(drw, view)
    px, py = _to_sheet(parent, through, sketch=True)
    # the parent's outline in its sketch coordinates (relative, unscaled)
    x1, y1, x2, y2 = sw.m0(parent, "GetOutline")
    cx, cy = parent.Position
    num, den = parent.ScaleRatio
    k = den / num
    pad = 0.005
    drw.ActivateView(parent.Name)
    drw.ClearSelection2(True)
    sm = drw.SketchManager
    if direction == "vertical":
        line = sm.CreateLine(px, (y1 - cy) * k - pad, 0, px, (y2 - cy) * k + pad, 0)
    else:
        line = sm.CreateLine((x1 - cx) * k - pad, py, 0, (x2 - cx) * k + pad, py, 0)
    if line is None:
        raise RuntimeError("Could not draw the cutting line")
    drw.ClearSelection2(True)
    line.Select4(False, sw.nothing())
    w, h = x2 - x1, y2 - y1
    if position:
        x, y = mm(position[0]), mm(position[1])
    else:  # SOLIDWORKS keeps a section aligned: beside a vertical cut, above/below a horizontal one
        x, y = _free_spot(drw, w, h, near=((x1 + x2) / 2, (y1 + y2) / 2), axis="x" if direction == "vertical" else "y")
    options = 4 if flip else 0  # swCreateSectionView_ChangeDirection
    new = drw.CreateSectionViewAt5(x, y, 0, label, options, sw.empty(), 0)
    if new is None:
        sw.m0(drw, "EditDelete")  # the cutting line
        raise RuntimeError("SOLIDWORKS did not create the section view")
    drw.ClearSelection2(True)
    drw.ForceRebuild3(False)
    return _view_info(new)


@tool
def detail_view(
    center: list[float],
    radius: float,
    scale: str = "2:1",
    view: str = "",
    label: str = "B",
    position: list[float] | None = None,
) -> dict:
    """Enlarged detail of the circle of `radius` (model mm) around a model
    point ([x,y,z] mm) in a drawing view (default the first). scale is the
    detail view's scale, e.g. "2:1"; position its centre on the sheet (mm)."""
    drw = sw.active_doc(DOC_DRAWING)
    parent = _drawing_view(drw, view)
    cx, cy = _to_sheet(parent, center, sketch=True)
    drw.ActivateView(parent.Name)
    drw.ClearSelection2(True)
    # view sketches are unscaled, so the radius stays in model size
    circle = drw.SketchManager.CreateCircleByRadius(cx, cy, 0, mm(radius))
    if circle is None:
        raise RuntimeError("Could not draw the detail circle")
    drw.ClearSelection2(True)
    circle.Select4(False, sw.nothing())
    s1, s2 = _parse_scale(scale)
    size = 2 * mm(radius) * s1 / s2
    ox1, oy1, ox2, oy2 = sw.m0(parent, "GetOutline")
    near = ((ox1 + ox2) / 2, (oy1 + oy2) / 2)
    x, y = (mm(position[0]), mm(position[1])) if position else _free_spot(drw, size, size, near=near)
    # standard style, circle shown, full outline
    new = drw.CreateDetailViewAt4(x, y, 0, 0, s1, s2, label, 1, True, False, False, 5)
    if new is None:
        sw.m0(drw, "EditDelete")  # the circle
        raise RuntimeError("SOLIDWORKS did not create the detail view")
    drw.ClearSelection2(True)
    drw.ForceRebuild3(False)
    return _view_info(new)


@tool
def set_drawing_scale(scale: str, view: str = "") -> dict:
    """Set the scale ("1:2", "2:1", ...) of the active sheet, or of one view
    (by name, see list_views) so it no longer follows the sheet."""
    drw = sw.active_doc(DOC_DRAWING)
    if view:
        target = _drawing_view(drw, view)
        _set_view_scale(target, scale)
        drw.ForceRebuild3(False)
        return _view_info(target)
    num, den = _parse_scale(scale)
    sheet = sw.m0(drw, "GetCurrentSheet")
    sheet.SetScale(num, den, True, True)
    drw.ForceRebuild3(False)
    return {"sheet": sw.m0(sheet, "GetName"), "scale": scale, "views": [_view_info(v) for v in _drawing_views(drw)]}


@tool
def drawing_note(text: str, position: list[float], height: float = 0.0) -> dict:
    """Put a note (several lines with \\n) on the active drawing sheet at
    [x,y] mm from the sheet's lower left corner; height = text height in mm
    (0 = document default)."""
    drw = sw.active_doc(DOC_DRAWING)
    drw.ClearSelection2(True)
    drw.ActivateSheet(sw.m0(sw.m0(drw, "GetCurrentSheet"), "GetName"))
    note = drw.InsertNote(text.replace("\n", "\r\n"))
    if note is None:
        raise RuntimeError("Could not insert the note")
    ann = sw.m0(note, "GetAnnotation")
    ann.SetPosition2(mm(position[0]), mm(position[1]), 0)
    if height > 0:
        tf = ann.GetTextFormat(0)
        tf.CharHeight = mm(height)
        ann.SetTextFormat(0, False, tf)
    drw.ClearSelection2(True)
    sw.m0(drw, "GraphicsRedraw2")
    return {"note": sw.m0(note, "GetName"), "position_mm": position}


@tool
def center_marks(view: str = "", holes: bool = True, fillets: bool = False, slots: bool = True) -> dict:
    """Add center marks to the holes (and optionally fillets and slots) of a
    drawing view, default every view."""
    drw = sw.active_doc(DOC_DRAWING)
    kinds = (1 if holes else 0) | (2 if fillets else 0) | (4 if slots else 0)  # swAutoInsertCenterMarkType_e
    views = [_drawing_view(drw, view)] if view else _drawing_views(drw)
    done = {}
    for v in views:
        drw.ActivateView(v.Name)
        # single marks, slot centres, document size and gap
        done[v.Name] = bool(v.AutoInsertCenterMarks2(kinds, 2, True, True, True, 0.0025, 0.0005, True, True, 0))
    drw.ClearSelection2(True)
    drw.ForceRebuild3(False)
    return {"center_marks": done}


def _centerlines(drw: Any, views: list) -> dict:
    """Centerlines on every cylinder seen from the side, per view: selecting
    the whole view and inserting a centerline inserts them all."""
    added = {}
    for v in views:
        before = sw.m0(v, "GetCenterLineCount") or 0
        drw.ActivateView(v.Name)
        drw.ClearSelection2(True)
        if drw.Extension.SelectByID2(v.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, sw.nothing(), 0):
            sw.invoke(drw, "InsertCenterLine2")
        added[v.Name] = (sw.m0(v, "GetCenterLineCount") or 0) - before
    drw.ClearSelection2(True)
    drw.ForceRebuild3(False)
    return added


@tool
def centerlines(view: str = "") -> dict:
    """Add centerlines along the cylinders (holes, shafts, bosses) that a
    drawing view shows from the side, default every view; views that show
    them as circles get center marks instead (center_marks)."""
    drw = sw.active_doc(DOC_DRAWING)
    return {"centerlines_added": _centerlines(drw, [_drawing_view(drw, view)] if view else _drawing_views(drw))}


VIEW_DISPLAY = {"hidden_visible": "swHIDDEN_GREYED", "hidden_removed": "swHIDDEN", "wireframe": "swWIREFRAME",
                "shaded": "swSHADED", "shaded_edges": "swSHADED_EDGES"}
TANGENT_EDGES = {"visible": "swTangentEdgesVisible", "fonted": "swTangentEdgesVisibleAndFonted",
                 "hidden": "swTangentEdgesHidden"}


@tool
def view_style(view: str = "", display: Literal["hidden_visible", "hidden_removed", "wireframe", "shaded",
                                                "shaded_edges", ""] = "",
               tangent_edges: Literal["visible", "fonted", "hidden", ""] = "") -> dict:
    """Display style of a drawing view (default every view): hidden lines
    shown dashed (hidden_visible), removed, wireframe or shaded; tangent
    edges (where fillets meet faces) visible, in a thin font or hidden."""
    drw = sw.active_doc(DOC_DRAWING)
    if not display and not tangent_edges:
        raise ValueError("Give display and/or tangent_edges")
    names = {v: k for k, v in VIEW_DISPLAY.items()}
    tangent_names = {v: k for k, v in TANGENT_EDGES.items()}
    out = {}
    for v in [_drawing_view(drw, view)] if view else _drawing_views(drw):
        if display:
            v.SetDisplayMode3(False, sw.constant(VIEW_DISPLAY[display]), False, True)
        if tangent_edges:
            v.SetDisplayTangentEdges2(sw.constant(TANGENT_EDGES[tangent_edges]))
        out[v.Name] = (sw.m0(v, "GetDisplayMode2"), sw.m0(v, "GetDisplayTangentEdges2"))
    drw.ForceRebuild3(False)
    result = {}
    for name, (mode, tangent) in out.items():
        mode_name = names.get(sw.enum_names("swDisplayMode_e").get(mode), mode)
        tangent_name = tangent_names.get(sw.enum_names("swDisplayTangentEdges_e").get(tangent), tangent)
        result[name] = {"display": mode_name, "tangent_edges": tangent_name}
        if (display and mode_name != display) or (tangent_edges and tangent_name != tangent_edges):
            result[name]["problem"] = "SOLIDWORKS kept a different style"
    return {"views": result}


def _table_boxes(drw: Any) -> list[tuple[float, float, float, float]]:
    """Sheet boxes (m) of the tables on the current sheet."""
    out = []
    ta = sw.m0(sw.m0(drw, "GetFirstView"), "GetFirstTableAnnotation")
    while ta is not None:
        try:
            out.append(_table_box(ta))
        except Exception:
            pass  # a table we cannot measure is not an obstacle we can avoid
        ta = sw.m0(ta, "GetNext")
    return out


def _table_anchor(table: Any) -> int:
    """swBOMConfigurationAnchorType_e: 1 top left, 2 top right, 3 bottom left, 4 bottom right."""
    try:
        return int(sw.m0(table, "AnchorType")) or 1
    except Exception:
        return 1


def _table_box(table: Any) -> tuple[float, float, float, float]:
    w = sum(table.GetColumnWidth(i) for i in range(sw.m0(table, "ColumnCount")))
    h = sum(table.GetRowHeight(i) for i in range(sw.m0(table, "RowCount")))
    x, y = list(sw.m0(sw.m0(table, "GetAnnotation"), "GetPosition"))[:2]  # the anchor corner
    anchor = _table_anchor(table)
    left = x if anchor in (1, 3) else x - w
    top = y if anchor in (1, 2) else y + h
    return (left, top - h, left + w, top)


def _place_table(drw: Any, table: Any, others: list) -> list[float]:
    """Move a table to a free spot near the sheet's top right, clear of
    views and other tables; returns its top left (mm)."""
    x1, y1, x2, y2 = _table_box(table)
    w, h = x2 - x1, y2 - y1
    sheet_w, sheet_h = _sheet_size(drw)
    cx, cy = _free_spot(drw, w, h, near=(sheet_w, sheet_h), extra=others, margin=0.015)
    left, top = cx - w / 2, cy + h / 2
    anchor = _table_anchor(table)
    ax = left if anchor in (1, 3) else left + w
    ay = top if anchor in (1, 2) else top - h
    sw.m0(table, "GetAnnotation").SetPosition2(ax, ay, 0)
    return [round(left / MM, 1), round(top / MM, 1)]


def _bom_template() -> str:
    """The standard BOM template of the installed SOLIDWORKS."""
    exe = str(sw.invoke(sw.app(), "GetExecutablePath"))
    root = Path(os.path.dirname(exe) if exe.lower().endswith(".exe") else exe)
    for lang in ("korean", "english"):
        p = root / "lang" / lang / "bom-standard.sldbomtbt"
        if p.exists():
            return str(p)
    found = sorted(root.glob("lang/*/bom-standard.sldbomtbt"))
    if found:
        return str(found[0])
    raise FileNotFoundError("No bom-standard.sldbomtbt in the SOLIDWORKS installation; pass template")


@tool
def bom_table(
    view: str = "",
    position: list[float] | None = None,
    kind: Literal["top_level", "parts_only", "indented"] = "top_level",
    template: str = "",
) -> dict:
    """Insert a bill of materials table for the assembly in a drawing view
    (default the first) with its top left corner at position [x,y] mm
    (default a free spot near the sheet's top right). template is a
    .sldbomtbt file, default the standard template SOLIDWORKS ships."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    sheet_w, sheet_h = _sheet_size(drw)
    x, y = (mm(position[0]), mm(position[1])) if position else (0.02, sheet_h - 0.02)
    bom_type = {"top_level": 2, "parts_only": 1, "indented": 3}[kind]  # swBomType_e
    template = template or _bom_template()
    others = _table_boxes(drw)
    drw.ActivateView(target.Name)
    # anchored at the top left corner (swBOMConfigurationAnchor_TopLeft)
    config = target.ReferencedConfiguration  # an empty name gives a table without rows
    table = target.InsertBomTable4(False, x, y, 1, bom_type, config, template, False, 0, False)
    if table is None:
        raise RuntimeError(f"Could not insert a BOM table (template {template}); does the view show an assembly?")
    corner = [x / MM, y / MM] if position else _place_table(drw, table, others)
    drw.ForceRebuild3(False)
    try:
        rows = sw.m0(table, "RowCount") - 1  # without the header
    except Exception:
        rows = None
    return {"rows": rows, "configuration": config, "top_left_mm": corner, "template": template}


@tool
def auto_balloons(view: str = "") -> dict:
    """Number the components of an assembly drawing with balloons (matching
    the BOM table), in one view (default every view)."""
    drw = sw.active_doc(DOC_DRAWING)
    views = [_drawing_view(drw, view)] if view else _drawing_views(drw)
    made = {}
    for v in views:
        before = _note_count(v)
        drw.ClearSelection2(True)
        drw.Extension.SelectByID2(v.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, sw.nothing(), 0)
        drw.AutoBalloon5(sw.m0(drw, "CreateAutoBalloonOptions"))
        made[v.Name] = _note_count(v) - before
    drw.ClearSelection2(True)
    drw.ForceRebuild3(False)
    return {"balloons": made}


def _note_count(view: Any) -> int:
    """Notes in a view; balloons are notes."""
    n, note = 0, sw.m0(view, "GetFirstNote")
    while note is not None:
        n += 1
        note = sw.m0(note, "GetNext")
    return n


def _view_pick(drw: Any, view: Any, ent: Any, mark: int, append: bool) -> bool:
    """Select a model entity shown in a drawing view, with a mark."""
    data = sw.m0(drw.SelectionManager, "CreateSelectData")
    data.Mark = mark
    data.View = view
    return bool(ent.Select4(append, data))


@tool
def hole_table(
    view: str = "",
    origin_point: list[float] | None = None,
    face_point: list[float] | None = None,
    position: list[float] | None = None,
    tags: Literal["letters", "numbers"] = "letters",
) -> dict:
    """Hole table (tag, X/Y position, size of every hole) for the holes of
    a face in a drawing view (default the first view that shows holes face
    on). origin_point is the model vertex ([x,y,z] mm) the X/Y positions
    are measured from, default the view's lower left corner; face_point a
    point on the face with the holes, default the face with most holes.
    position is the table's top left corner on the sheet (mm), default a
    free spot near the sheet's top right."""
    drw = sw.active_doc(DOC_DRAWING)
    views = [_drawing_view(drw, view)] if view else _drawing_views(drw)
    vk, fk = sw.constant("swViewEntityType_Vertex"), sw.constant("swViewEntityType_Face")
    target = face = None
    for v in views:
        faces = list(v.GetVisibleEntities2(sw.nothing(), fk) or ())
        if face_point:
            p = tuple(mm(c) for c in face_point)
            near = [(math.dist(p, tuple(f.GetClosestPointOn(*p)[:3])), f) for f in faces]
            near = [n for n in near if n[0] <= sw.TOLERANCE]
            cand = min(near, key=lambda n: n[0])[1] if near else None
        else:  # the face with most loops: an outline plus a loop per hole
            cand = max(faces, key=lambda f: sw.m0(f, "GetLoopCount"), default=None)
            cand = cand if cand is not None and sw.m0(cand, "GetLoopCount") > 1 else None
        if cand is not None:
            target, face = v, cand
            break
    if target is None:
        raise LookupError("No view shows a face with holes" + (f" at {face_point} mm" if face_point else ""))
    vertices = list(target.GetVisibleEntities2(sw.nothing(), vk) or ())
    if origin_point:
        p = tuple(mm(c) for c in origin_point)
        origin = min(vertices, key=lambda vx: math.dist(p, _vertex_point(target, vx, p)), default=None)
    else:  # lower left on the sheet
        def sheet_xy(vx):
            x, y = _to_sheet(target, [c / MM for c in _vertex_point(target, vx)])
            return x + y
        origin = min(vertices, key=sheet_xy, default=None)
    if origin is None:
        raise LookupError(f"No vertex in {target.Name} to measure from")
    drw.ActivateView(target.Name)
    drw.ClearSelection2(True)
    _view_pick(drw, target, origin, 1, False)  # origin with mark 1, faces with mark 2
    _view_pick(drw, target, face, 2, True)
    sheet_w, sheet_h = _sheet_size(drw)
    x, y = (mm(position[0]), mm(position[1])) if position else (0.02, sheet_h - 0.02)
    template = _table_template(f"standard hole table--{tags}.sldholtbt")
    others = _table_boxes(drw)
    table = target.InsertHoleTable2(False, x, y, 1, "A" if tags == "letters" else "1", template)  # anchored top left
    drw.ClearSelection2(True)
    if table is None:
        raise RuntimeError("SOLIDWORKS did not create the hole table")
    corner = [x / MM, y / MM] if position else _place_table(drw, table, others)
    drw.ForceRebuild3(False)
    try:
        rows = sw.m0(table, "RowCount") - 1
    except Exception:
        rows = None
    return {"view": target.Name, "holes": rows, "top_left_mm": corner, "template": template}


def _table_template(name: str) -> str:
    exe = str(sw.invoke(sw.app(), "GetExecutablePath"))
    root = Path(os.path.dirname(exe) if exe.lower().endswith(".exe") else exe)
    for lang in ("korean", "english"):
        p = root / "lang" / lang / name
        if p.exists():
            return str(p)
    found = sorted(root.glob(f"lang/*/{name}"))
    if found:
        return str(found[0])
    raise FileNotFoundError(f"No {name} in the SOLIDWORKS installation")


def _sheet_format_path(name: str) -> str:
    if name.lower().endswith(".slddrt") and Path(name).exists():
        return name
    stem = name[:-7] if name.lower().endswith(".slddrt") else name
    roots = [Path(os.environ.get("ProgramData", r"C:\ProgramData")) / "SOLIDWORKS"]
    exe = str(sw.invoke(sw.app(), "GetExecutablePath"))
    roots.append(Path(os.path.dirname(exe) if exe.lower().endswith(".exe") else exe))
    for root in roots:
        for lang in ("korean", "english"):
            for p in root.glob(f"*/lang/{lang}/sheetformat/{stem}.slddrt"):
                return str(p)
            p = root / "lang" / lang / "sheetformat" / f"{stem}.slddrt"
            if p.exists():
                return str(p)
    raise FileNotFoundError(f"No sheet format {name!r}; try e.g. \"a3 - iso\" or \"a4 - iso\"")


PROPERTY_LINK = re.compile(r'\$PRP(SHEET)?:"([^"]+)"', re.IGNORECASE)


def _title_block_fields(drw: Any) -> dict[str, str]:
    """Property names the sheet format's notes show: {name: "drawing" for
    $PRP, "model" for $PRPSHEET}."""
    fields: dict[str, str] = {}
    sw.m0(drw, "EditTemplate")  # the sheet format's notes are only reachable while editing it
    try:
        note = sw.m0(sw.m0(drw, "GetFirstView"), "GetFirstNote")
        while note is not None:
            try:
                text = sw.m0(note, "PropertyLinkedText") or ""
            except Exception:
                text = ""
            for sheet, prop in PROPERTY_LINK.findall(text):
                fields.setdefault(prop, "model" if sheet else "drawing")
            note = sw.m0(note, "GetNext")
    finally:
        sw.m0(drw, "EditSheet")
    return fields


@tool
def title_block(fields: dict[str, str] | None = None, sheet_format: str = "") -> dict:
    """Fill in the drawing's title block. sheet_format first replaces the
    current sheet's format, e.g. "a3 - iso" or "a4 - iso" (the standard
    formats with a title block) or a .slddrt path. fields sets the
    properties the title block shows ({"Description": "브래킷",
    "DrawnBy": "..."}): those linked with $PRPSHEET go to the model, the
    rest to the drawing. Returns the title block's fields and their values."""
    drw = sw.active_doc(DOC_DRAWING)
    sheet = sw.m0(drw, "GetCurrentSheet")
    if sheet_format:
        path = _sheet_format_path(sheet_format)
        paper, _, s1, s2, first_angle, w, h = list(sw.m0(sheet, "GetProperties2"))[:7]
        ok = drw.SetupSheet5(sw.m0(sheet, "GetName"), int(paper), 12, s1, s2, bool(first_angle), path, w, h, "Default", True)
        if not ok:
            raise RuntimeError(f"Could not apply the sheet format {path}")
        drw.ForceRebuild3(False)
    linked = _title_block_fields(drw)
    views = _drawing_views(drw)
    model = None
    if views:
        model_path = os.path.normcase(sw.m0(views[0], "GetReferencedModelName") or "")
        model = next((d for d in sw.m0(sw.app(), "GetDocuments") or ()
                      if os.path.normcase(sw.m0(d, "GetPathName") or "") == model_path), None)
    managers = {"drawing": drw.Extension.CustomPropertyManager("")}
    if model is not None:
        managers["model"] = model.Extension.CustomPropertyManager("")
    for name, value in (fields or {}).items():
        where = linked.get(name, "drawing")
        cpm = managers.get(where, managers["drawing"])
        cpm.Add3(name, 30, str(value), 2)  # text, replace
    drw.ForceRebuild3(False)
    out = {}
    for name, where in linked.items():
        cpm = managers.get(where)
        out[name] = {"from": where, "value": cpm.Get(name) if cpm is not None else None}
    result = {"sheet_format": sw.m0(sheet, "GetTemplateName"), "fields": out}
    unknown = [n for n in (fields or {}) if n not in linked]
    if unknown:
        result["not_in_title_block"] = unknown  # set on the drawing, but nothing shows them
    return result


TOLERANCE_TYPES = {"none": 0, "basic": 1, "bilateral": 2, "limit": 3, "symmetric": 4, "min": 5, "max": 6,
                   "fit": 7, "fit_with_tolerance": 8, "fit_only": 9}  # swTolType_e


def _find_dimension(doc: Any, name: str) -> Any:
    """A dimension by name: "D1@스케치1" in a part/assembly, or as list_dimensions
    shows it in a drawing (model dimensions and the drawing's own)."""
    if sw.m0(doc, "GetType") != DOC_DRAWING:
        dim = doc.Parameter(name)
        if dim is None:
            raise LookupError(f"No dimension {name!r}; see list_dimensions")
        return dim
    for view in [sw.m0(doc, "GetFirstView")] + _drawing_views(doc):
        dd = sw.m0(view, "GetFirstDisplayDimension5")
        while dd is not None:
            dim = dd.GetDimension2(0)
            full = dim.FullName
            if name in (full, "@".join(full.split("@")[:2])):
                return dim
            dd = sw.m0(dd, "GetNext5")
    raise LookupError(f"No dimension {name!r} in the drawing; see list_dimensions")


@tool
def set_tolerance(
    dimension: str,
    kind: Literal["none", "basic", "bilateral", "limit", "symmetric", "min", "max", "fit", "fit_with_tolerance", "fit_only"],
    upper: float = 0.0,
    lower: float = 0.0,
    hole_fit: str = "",
    shaft_fit: str = "",
) -> dict:
    """Tolerance of a dimension (name from list_dimensions, e.g. "D1@스케치1";
    in a drawing as listed there). symmetric uses upper (±); bilateral and
    limit use upper and lower deviations in mm (e.g. upper=0.1, lower=-0.05);
    fit uses ISO fit classes, hole_fit like "H7" and/or shaft_fit like "g6".
    Model dimensions keep it in every drawing that shows them."""
    doc = sw.active_doc()
    dim = _find_dimension(doc, dimension)
    tol = sw.m0(dim, "Tolerance")
    tol.Type = TOLERANCE_TYPES[kind]
    # SetValues2 (per configuration) returns False here; SetValues works
    if kind == "symmetric":
        tol.SetValues(-mm(abs(upper)), mm(abs(upper)))
    elif kind in ("bilateral", "limit", "min", "max", "fit_with_tolerance"):
        tol.SetValues(mm(lower), mm(upper))
    if kind.startswith("fit"):
        tol.SetFitValues(hole_fit, shaft_fit)
    doc.ForceRebuild3(False)
    sw.m0(doc, "GraphicsRedraw2")
    out = {"dimension": dim.FullName, "value_mm": round(dim.Value, 4),
           "tolerance": {v: k for k, v in TOLERANCE_TYPES.items()}.get(tol.Type, tol.Type)}
    try:
        out["upper_mm"], out["lower_mm"] = round(sw.m0(tol, "GetMaxValue") / MM, 4), round(sw.m0(tol, "GetMinValue") / MM, 4)
    except Exception:
        pass
    if kind.startswith("fit"):
        out["fit"] = [sw.m0(tol, "GetHoleFitValue"), sw.m0(tol, "GetShaftFitValue")]
    return out


def _vertex_point(view: Any, vertex: Any, near: tuple | None = None) -> tuple:
    """Model point (m) of a vertex from a view's visible entities. Those
    report the model point shifted by the view's sheet position (edges do
    not), so the shift is taken off; with `near`, whichever reading is
    closer wins in case a release reports it plainly."""
    raw = tuple(sw.m0(vertex, "GetPoint")[:3])
    px, py = view.Position
    fixed = (raw[0] - px, raw[1] - py, raw[2])
    if near is None:
        return fixed
    return min((fixed, raw), key=lambda q: math.dist(q, near))


def _view_entity(view: Any, point_mm: list[float], kind: str) -> Any:
    """The edge or vertex shown in a view that is nearest a model point;
    kind "center" is the circular edge whose centre is nearest (a hole)."""
    types = {"edge": sw.constant("swViewEntityType_Edge"), "vertex": sw.constant("swViewEntityType_Vertex")}
    p = tuple(mm(c) for c in point_mm)
    best, best_d = None, math.inf
    for ent in view.GetVisibleEntities2(sw.nothing(), types["edge" if kind == "center" else kind]) or ():
        try:
            if kind == "vertex":
                q = _vertex_point(view, ent, p)
            elif kind == "center":
                curve = sw.m0(ent, "GetCurve")
                if not sw.invoke(curve, "IsCircle"):
                    continue
                q = curve.CircleParams[:3]
            else:
                q = ent.GetClosestPointOn(*p)
        except Exception:
            continue
        d = math.dist(p, tuple(q[:3]))
        if d < best_d:
            best, best_d = ent, d
    if best is None or best_d > sw.TOLERANCE:
        raise LookupError(f"No {kind} of {view.Name} within 0.5 mm of {point_mm} mm")
    return best


@tool
def add_dimension(
    point1: list[float],
    point2: list[float] | None = None,
    entity1: Literal["edge", "vertex", "center"] = "edge",
    entity2: Literal["edge", "vertex", "center"] = "edge",
    view: str = "",
    direction: Literal["aligned", "horizontal", "vertical"] = "aligned",
    position: list[float] | None = None,
) -> dict:
    """Add a dimension to a drawing view (default the first) on model
    geometry picked by model points ([x,y,z] mm): one edge (its length, or
    a circle's diameter) or the distance between two edges/vertices.
    entity "center" picks the hole (circular edge) centred at the point, for
    centre-to-centre distances.
    direction forces a horizontal or vertical distance; position is the
    dimension text on the sheet (mm), default just beside the view."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    first = _view_entity(target, point1, entity1)
    second = _view_entity(target, point2, entity2) if point2 else None
    drw.ActivateView(target.Name)
    drw.ClearSelection2(True)
    _view_pick(drw, target, first, 0, False)
    if second is not None:
        _view_pick(drw, target, second, 0, True)
    if position:
        x, y = mm(position[0]), mm(position[1])
    else:  # above the midpoint of what is measured, outside the view
        pts = [_to_sheet(target, point1)] + ([_to_sheet(target, point2)] if point2 else [])
        x1, y1, x2, y2 = sw.m0(target, "GetOutline")
        side = "left" if direction == "vertical" else "top"
        stack = 0.008 + 0.007 * _dimension_count(target, side)  # one row further out per dimension there
        along = sum(p[1] for p in pts) / len(pts) if direction == "vertical" else sum(p[0] for p in pts) / len(pts)
        x, y = _beside_view(drw, target, side, stack, along)
    add = {"aligned": "AddDimension2", "horizontal": "AddHorizontalDimension2", "vertical": "AddVerticalDimension2"}[direction]
    app = sw.app()
    prompt = app.GetUserPreferenceToggle(INPUT_DIM_ON_CREATE)
    app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, False)  # no "Modify" dialog
    try:
        dim = getattr(drw, add)(x, y, 0)
    finally:
        app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, prompt)
    drw.ClearSelection2(True)
    if dim is None:
        raise RuntimeError("SOLIDWORKS did not add the dimension; check the picked geometry")
    d = dim.GetDimension2(0)
    return {"dimension": "@".join(d.FullName.split("@")[:2]), "value_mm": round(d.Value, 4), "view": target.Name}


def _dimension_count(view: Any, side: str = "") -> int:
    """Dimensions shown in a view; with side ("top"/"left"/...), only those
    whose text sits beyond that side of the view's outline."""
    x1, y1, x2, y2 = sw.m0(view, "GetOutline")
    n, dd = 0, sw.m0(view, "GetFirstDisplayDimension5")
    while dd is not None:
        if not side:
            n += 1
        else:
            try:
                tx, ty = list(sw.m0(sw.m0(dd, "GetAnnotation"), "GetPosition"))[:2]
            except Exception:
                tx, ty = (x1 + x2) / 2, (y1 + y2) / 2
            n += {"top": ty > y2, "bottom": ty < y1, "left": tx < x1, "right": tx > x2}[side]
        dd = sw.m0(dd, "GetNext5")
    return n


BORDER = 0.02  # m kept clear along the sheet edges (the sheet format's frame and zone labels)


def _inside_border(drw: Any, x: float, y: float) -> tuple[float, float]:
    sheet_w, sheet_h = _sheet_size(drw)
    return min(max(x, BORDER), sheet_w - BORDER), min(max(y, BORDER), sheet_h - BORDER)


def _beside_view(drw: Any, view: Any, side: str, distance: float, along: float) -> tuple[float, float]:
    """A sheet point `distance` beyond one side ("top"/"bottom"/"left"/
    "right") of a view's outline, at `along` (x for top/bottom, y for
    left/right). When the sheet border would push it back onto the view,
    the opposite side is used instead."""
    x1, y1, x2, y2 = sw.m0(view, "GetOutline")
    flip = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
    for s in (side, flip[side]):
        spot = {"top": (along, y2 + distance), "bottom": (along, y1 - distance),
                "left": (x1 - distance, along), "right": (x2 + distance, along)}[s]
        x, y = _inside_border(drw, *spot)
        clear = {"top": y > y2, "bottom": y < y1, "left": x < x1, "right": x > x2}[s]
        if clear:
            return x, y
    return _inside_border(drw, *{"top": (along, y2 + distance), "bottom": (along, y1 - distance),
                                 "left": (x1 - distance, along), "right": (x2 + distance, along)}[side])


@tool
def projected_view(view: str = "", side: Literal["right", "left", "top", "bottom"] = "right",
                   position: list[float] | None = None) -> dict:
    """Add a view projected from a drawing view (default the first) to its
    right/left/top/bottom, aligned with it. position is its centre on the
    sheet (mm); by default the nearest free spot on that side."""
    drw = sw.active_doc(DOC_DRAWING)
    parent = _drawing_view(drw, view)
    x1, y1, x2, y2 = sw.m0(parent, "GetOutline")
    cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
    sheet_w, sheet_h = _sheet_size(drw)
    if position:
        x, y = mm(position[0]), mm(position[1])
    else:
        w, h, gap = x2 - x1, y2 - y1, 0.02
        x, y = {"right": (x2 + gap + w / 2, cy), "left": (x1 - gap - w / 2, cy),
                "top": (cx, y2 + gap + h / 2), "bottom": (cx, y1 - gap - h / 2)}[side]
        if not (0 < x < sheet_w and 0 < y < sheet_h):
            raise ValueError(f"No room {side} of {parent.Name} on the sheet; move the view (or pass position) first")
    drw.ClearSelection2(True)
    drw.Extension.SelectByID2(parent.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, sw.nothing(), 0)
    new = drw.CreateUnfoldedViewAt3(x, y, 0, False)
    drw.ClearSelection2(True)
    if new is None:
        raise RuntimeError("SOLIDWORKS did not create the projected view")
    drw.ForceRebuild3(False)
    if not position:  # landed on another view: slide it further out on the same side
        nx1, ny1, nx2, ny2 = sw.m0(new, "GetOutline")
        ncx, ncy = (nx1 + nx2) / 2, (ny1 + ny2) / 2
        nx, ny = _free_spot(drw, nx2 - nx1, ny2 - ny1, exclude=new.Name, near=(ncx, ncy),
                            axis="x" if side in ("right", "left") else "y")
        on_side = {"right": nx > cx, "left": nx < cx, "top": ny > cy, "bottom": ny < cy}[side]
        if (nx, ny) != (ncx, ncy) and on_side:
            new.Position = sw.doubles([new.Position[0] + nx - ncx, new.Position[1] + ny - ncy])
            drw.ForceRebuild3(False)
    return _view_info(new)


@tool
def drawing_sheets(add: str = "", activate: str = "", sheet_format: str = "", scale: str = "") -> dict:
    """List the drawing's sheets; add one (named `add`, same paper as the
    current sheet, optionally another sheet_format such as "a3 - iso" and
    scale "1:2") or activate one. New views go on the active sheet; a PDF
    export contains every sheet."""
    drw = sw.active_doc(DOC_DRAWING)
    if add:
        cur = sw.m0(drw, "GetCurrentSheet")
        paper, _, s1, s2, first_angle, w, h = list(sw.m0(cur, "GetProperties2"))[:7]
        if scale:
            s1, s2 = _parse_scale(scale)
        template = _sheet_format_path(sheet_format) if sheet_format else (sw.m0(cur, "GetTemplateName") or "")
        kind = 12 if template else 13  # swDwgTemplateCustom / swDwgTemplateNone
        if not drw.NewSheet4(add, int(paper), kind, s1, s2, bool(first_angle), template, w, h, "Default",
                             0, 0, 0, 0, 0, 0):
            raise RuntimeError(f"Could not add sheet {add!r}")
    if activate and not drw.ActivateSheet(activate):
        raise LookupError(f"No sheet {activate!r}")
    return {"sheets": list(sw.m0(drw, "GetSheetNames") or ()),
            "active": sw.m0(sw.m0(drw, "GetCurrentSheet"), "GetName")}


def _annotation_spot(view: Any, point_mm: list[float], distance: float = 0.012) -> tuple[float, float]:
    """A sheet position (m) just outside the view, beyond the side of its
    outline nearest the model point, with room for earlier annotations."""
    px, py = _to_sheet(view, point_mm)
    x1, y1, x2, y2 = sw.m0(view, "GetOutline")
    sides = {"left": px - x1, "right": x2 - px, "bottom": py - y1, "top": y2 - py}
    side = min(sides, key=sides.get)
    n = 0
    ann = sw.m0(view, "GetFirstAnnotation3")
    while ann is not None:
        if sw.m0(ann, "GetType") != 4:  # symbols, not dimensions
            try:
                ax, ay = list(sw.m0(ann, "GetPosition"))[:2]
                n += {"left": ax < x1, "right": ax > x2, "bottom": ay < y1, "top": ay > y2}[side]
            except Exception:
                pass
        ann = sw.m0(ann, "GetNext3")
    d = distance + 0.006 * (n % 4)  # stagger so symbols on the same side do not sit on each other
    spot = {"left": (x1 - d, py), "right": (x2 + d, py), "bottom": (px, y1 - d), "top": (px, y2 + d)}[side]
    return _inside_border(sw.active_doc(DOC_DRAWING), *spot)


def _attach(drw: Any, view: Any, point: list[float], entity: str) -> None:
    ent = _view_entity(view, point, entity)
    drw.ActivateView(view.Name)
    drw.ClearSelection2(True)
    _view_pick(drw, view, ent, 0, False)


def _place(ann: Any, view: Any, point: list[float], position: list[float] | None) -> list[float]:
    x, y = (mm(position[0]), mm(position[1])) if position else _annotation_spot(view, point)
    ann.SetPosition2(x, y, 0)
    return [round(x / MM, 1), round(y / MM, 1)]


@tool
def datum(label: str, point: list[float], view: str = "", entity: Literal["edge", "vertex"] = "edge",
          position: list[float] | None = None) -> dict:
    """Datum feature symbol (e.g. "A") on the model edge through `point`
    ([x,y,z] mm) in a drawing view (default the first). position is the
    symbol on the sheet (mm); by default just outside the view."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    _attach(drw, target, point, entity)
    tag = sw.invoke(drw, "InsertDatumTag2")
    drw.ClearSelection2(True)
    if tag is None:
        raise RuntimeError("SOLIDWORKS did not insert the datum symbol")
    tag.SetLabel(label)
    at = _place(sw.m0(tag, "GetAnnotation"), target, point, position)
    sw.m0(drw, "GraphicsRedraw2")
    return {"datum": label, "view": target.Name, "position_mm": at}


GTOL_SYMBOLS = {
    "straightness": "swGcsSTRAIGHT", "flatness": "swGcsFLAT", "circularity": "swGcsROUND",
    "cylindricity": "swGcsCYL", "line_profile": "swGcsLINEPROF", "surface_profile": "swGcsSURFPROF",
    "angularity": "swGcsANG", "perpendicularity": "swGcsPERP", "parallelism": "swGcsPARALLEL",
    "position": "swGcsPOSITION", "concentricity": "swGcsCONC", "symmetry": "swGcsSYMMETRY",
    "circular_runout": "swGcsCIRCRUNOUT", "total_runout": "swGcsTOTALRUNOUT",
}


@tool
def geometric_tolerance(
    symbol: Literal["straightness", "flatness", "circularity", "cylindricity", "line_profile", "surface_profile",
                    "angularity", "perpendicularity", "parallelism", "position", "concentricity", "symmetry",
                    "circular_runout", "total_runout"],
    tolerance: float,
    point: list[float],
    datums: list[str] | None = None,
    diameter: bool = False,
    material_condition: Literal["", "MMC", "LMC"] = "",
    view: str = "",
    entity: Literal["edge", "vertex", "center"] = "edge",
    position: list[float] | None = None,
) -> dict:
    """Geometric tolerance frame (e.g. flatness 0.05, or position Ø0.1 Ⓜ
    to datums A and B) attached to the model edge through `point` ([x,y,z]
    mm; entity "center" picks the hole centred there) in a drawing view
    (default the first). position is the frame on the sheet (mm)."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    _attach(drw, target, point, entity)
    gtol = sw.invoke(drw, "InsertGtol")
    drw.ClearSelection2(True)
    if gtol is None:
        raise RuntimeError("SOLIDWORKS did not insert the tolerance frame")
    mc = {"": "swMcNONE", "MMC": "swMcMMC", "LMC": "swMcLMC"}[material_condition]
    # the short-code variant; SetFrameSymbols2 wants the symbols as text codes
    gtol.SetFrameSymbols(1, sw.constant(GTOL_SYMBOLS[symbol]), diameter, sw.constant(mc), False, 0, 0, 0, 0)
    d = list(datums or []) + ["", "", ""]
    gtol.SetFrameValues2(1, f"{tolerance:g}", "", d[0], d[1], d[2])
    at = _place(sw.m0(gtol, "GetAnnotation"), target, point, position)
    sw.m0(drw, "GraphicsRedraw2")
    return {"symbol": symbol, "tolerance": tolerance, "datums": datums or [], "view": target.Name, "position_mm": at,
            "frame": list(sw.invoke(gtol, "GetFrameSymbols3", 1) or ())[:2]}


@tool
def surface_finish(point: list[float], roughness: str = "Ra 1.6", kind: Literal["machined", "basic", "no_machining"] = "machined",
                   view: str = "", position: list[float] | None = None) -> dict:
    """Surface texture symbol with a roughness value (e.g. "Ra 1.6") on the
    model edge through `point` ([x,y,z] mm) in a drawing view (default the
    first). kind: machined (material removal required), basic, or
    no_machining (removal not allowed)."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    _attach(drw, target, point, "edge")
    x, y = (mm(position[0]), mm(position[1])) if position else _annotation_spot(target, point, 0.015)
    sym = {"basic": 0, "machined": 1, "no_machining": 2}[kind]  # swSFSymType_e
    # symbol type, straight leader, position, no lay, arrow, then empty texts
    sf = drw.Extension.InsertSurfaceFinishSymbol3(sym, 1, x, y, 0, 0, 1, "", "", "", "", "", "", "")
    drw.ClearSelection2(True)
    if sf is None:
        raise RuntimeError("SOLIDWORKS did not insert the surface finish symbol")
    sf.SetText(sw.constant("swSFSymbolRoughnessValue1"), roughness)  # the value under the symbol's bar
    sw.m0(drw, "GraphicsRedraw2")
    return {"roughness": roughness, "kind": kind, "view": target.Name, "position_mm": [round(x / MM, 1), round(y / MM, 1)]}


@tool
def revision_table(revision: str = "", description: str = "", approved: str = "") -> dict:
    """Add a revision row (e.g. "A", "개정 내용") to the drawing's revision
    table, creating the table at the sheet's top right first if there is
    none. Without arguments it only lists the revisions."""
    drw = sw.active_doc(DOC_DRAWING)
    sheet = sw.m0(drw, "GetCurrentSheet")
    table = sw.m0(sheet, "RevisionTable")
    created = False
    if table is None:
        others = _table_boxes(drw)
        sheet_w, sheet_h = _sheet_size(drw)
        template = _table_template("standard revision block.sldrevtbt")
        # top left anchor, circle revision symbols, update zone cells
        table = sheet.InsertRevisionTable2(False, sheet_w - 0.2, sheet_h - 0.015, 1, template, 1, True)
        if table is None:
            raise RuntimeError("Could not insert a revision table")
        _place_table(drw, table, others)
        created = True
    cols = sw.m0(table, "ColumnCount")
    # the template may start with a title row ("수정본") above the column headings
    head_row = next((r for r in range(min(3, sw.m0(table, "RowCount")))
                     if sum(1 for c in range(cols) if (table.Text(r, c) or "").strip()) > 1), 0)
    heads = [(table.Text(head_row, c) or "").strip().lower() for c in range(cols)]
    if revision:
        if not table.AddRevision(revision):
            raise RuntimeError(f"Could not add revision {revision!r}")
        row = sw.m0(table, "RowCount") - 1  # appended at the bottom
        for col, head in enumerate(heads):
            if description and any(k in head for k in ("설명", "비고", "description", "내용")):
                _set_cell(table, row, col, description)
            if approved and any(k in head for k in ("승인", "approved")):
                _set_cell(table, row, col, approved)
    drw.ForceRebuild3(False)
    rows = [[table.Text(r, c) for c in range(cols)] for r in range(head_row + 1, sw.m0(table, "RowCount"))]
    return {"created": created, "columns": [table.Text(head_row, c) for c in range(cols)], "revisions": rows}


def _set_cell(table: Any, row: int, col: int, text: str) -> None:
    """ITableAnnotation.Text(row, col) = text: an indexed property write."""
    import pythoncom

    dispid = table._oleobj_.GetIDsOfNames("Text")
    table._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, False, row, col, text)


@tool
def broken_out_section(center: list[float], radius: float, depth: float, view: str = "") -> dict:
    """Broken-out (partial) section: cut away a circular region of `radius`
    mm around a model point ([x,y,z] mm) in a drawing view (default the
    first), `depth` mm deep from the front of the view, to show what is
    inside."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    cx, cy = _to_sheet(target, center, sketch=True)
    drw.ActivateView(target.Name)
    drw.ClearSelection2(True)
    circle = drw.SketchManager.CreateCircleByRadius(cx, cy, 0, mm(radius))
    if circle is None:
        raise RuntimeError("Could not draw the section outline")
    drw.ClearSelection2(True)
    circle.Select4(False, sw.nothing())
    before = sw.m0(target, "GetBreakOutSectionCount")
    sw.invoke(drw, "CreateBreakOutSection", mm(depth))
    drw.ClearSelection2(True)
    drw.ForceRebuild3(False)
    if sw.m0(target, "GetBreakOutSectionCount") <= before:
        raise RuntimeError("SOLIDWORKS did not create the broken-out section; is the depth inside the part?")
    return {"view": target.Name, "broken_out_sections": sw.m0(target, "GetBreakOutSectionCount")}


@tool
def auxiliary_view(edge_point: list[float], view: str = "", label: str = "C",
                   position: list[float] | None = None) -> dict:
    """Auxiliary view looking square onto an inclined face: projected
    perpendicular to the model edge through edge_point ([x,y,z] mm) as seen
    in a drawing view (default the first). position is its centre on the
    sheet (mm); by default a free spot on the side the edge faces."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    edge = _view_entity(target, edge_point, "edge")
    drw.ActivateView(target.Name)
    drw.ClearSelection2(True)
    _view_pick(drw, target, edge, 0, False)
    x1, y1, x2, y2 = sw.m0(target, "GetOutline")
    # the view is projected perpendicular to the edge, away from the parent view
    px, py = _to_sheet(target, edge_point)
    try:
        ends = [_to_sheet(target, [c / MM for c in sw.m0(sw.m0(edge, v), "GetPoint")]) for v in ("GetStartVertex", "GetEndVertex")]
        ex, ey = ends[1][0] - ends[0][0], ends[1][1] - ends[0][1]
        nx, ny = -ey, ex
    except Exception:  # no end vertices: fall back to "away from the view centre"
        nx, ny = px - (x1 + x2) / 2, py - (y1 + y2) / 2
    k = math.hypot(nx, ny) or 1
    nx, ny = nx / k, ny / k
    if nx * (px - (x1 + x2) / 2) + ny * (py - (y1 + y2) / 2) < 0:
        nx, ny = -nx, -ny
    if position:
        x, y = mm(position[0]), mm(position[1])
    else:
        reach = 0.02 + max(x2 - x1, y2 - y1) / 2
        x, y = px + nx * reach, py + ny * reach
    new = drw.CreateAuxiliaryViewAt2(x, y, 0, False, label, True, False)
    drw.ClearSelection2(True)
    if new is None:
        raise RuntimeError("SOLIDWORKS did not create the auxiliary view; pick a straight edge")
    drw.ForceRebuild3(False)
    if not position:  # slide it out along the projection direction until it is clear
        sheet_w, sheet_h = _sheet_size(drw)
        others = [sw.m0(v, "GetOutline") for v in _drawing_views(drw) if v.Name != new.Name] + _table_boxes(drw)

        def clash() -> bool:
            a1, b1, a2, b2 = sw.m0(new, "GetOutline")
            return any(a2 + 0.01 > o[0] and a1 - 0.01 < o[2] and b2 + 0.01 > o[1] and b1 - 0.01 < o[3] for o in others)

        for _ in range(80):
            a1, b1, a2, b2 = sw.m0(new, "GetOutline")
            if not clash() or a1 < 0 or b1 < 0 or a2 > sheet_w or b2 > sheet_h:
                break
            pos = new.Position
            new.Position = sw.doubles([pos[0] + nx * 0.005, pos[1] + ny * 0.005])
            drw.ForceRebuild3(False)
        if clash():  # no room in line with the edge: recreate it unaligned in a free spot
            a1, b1, a2, b2 = sw.m0(new, "GetOutline")
            fx, fy = _free_spot(drw, a2 - a1, b2 - b1, exclude=new.Name, extra=_table_boxes(drw))
            name = new.Name
            drw.ClearSelection2(True)
            drw.Extension.SelectByID2(name, "DRAWINGVIEW", 0, 0, 0, False, 0, sw.nothing(), 0)
            drw.Extension.DeleteSelection2(0)
            drw.ActivateView(target.Name)
            drw.ClearSelection2(True)
            _view_pick(drw, target, _view_entity(target, edge_point, "edge"), 0, False)
            new = drw.CreateAuxiliaryViewAt2(fx, fy, 0, True, label, True, False)
            drw.ClearSelection2(True)
            if new is None:
                raise RuntimeError("SOLIDWORKS did not recreate the auxiliary view off its alignment")
            drw.ForceRebuild3(False)
    return _view_info(new)


@tool
def cut_list_table(view: str = "", position: list[float] | None = None) -> dict:
    """Weldment cut list table (item, quantity, description, length) for the
    weldment part in a drawing view (default the first), at position (top
    left, mm) or a free spot near the sheet's top right."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    sheet_w, sheet_h = _sheet_size(drw)
    x, y = (mm(position[0]), mm(position[1])) if position else (0.02, sheet_h - 0.02)
    others = _table_boxes(drw)
    drw.ActivateView(target.Name)
    template = _table_template("cut list.sldwldtbt")
    # unlike BOM tables, the view's configuration name ("기본<가공 상태>") gives nothing; "" works
    table = (target.InsertWeldmentTable(False, x, y, 1, "", template)
             or target.InsertWeldmentTable(False, x, y, 1, target.ReferencedConfiguration, template))
    if table is None:
        raise RuntimeError("Could not insert a cut list table; does the view show a weldment part?")
    corner = [x / MM, y / MM] if position else _place_table(drw, table, others)
    drw.ForceRebuild3(False)
    try:
        rows = sw.m0(table, "RowCount") - 1
    except Exception:
        rows = None
    return {"rows": rows, "top_left_mm": corner}


WELD_SYMBOLS = {"fillet": "FILL", "butt": "BUTT", "square": "BUSQ", "v": "BUSV", "bevel": "BUSB", "u": "BUSU",
                "j": "BUSJ", "plug": "PLUG", "spot": "SPOT", "seam": "SEAM", "backing": "BACK"}


@tool
def weld_symbol(
    point: list[float],
    symbol: Literal["fillet", "butt", "square", "v", "bevel", "u", "j", "plug", "spot", "seam", "backing"] = "fillet",
    size: str = "",
    length: str = "",
    other_side: bool = False,
    all_around: bool = False,
    field_weld: bool = False,
    view: str = "",
    position: list[float] | None = None,
) -> dict:
    """ISO weld symbol pointing at the model edge through `point` ([x,y,z]
    mm) in a drawing view (default the first): symbol type, size (e.g. "5"
    for a 5 mm fillet, left of the symbol) and length/pitch (right of it),
    on the arrow side or the other side; all_around draws the circle,
    field_weld the flag."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    _attach(drw, target, point, "edge")
    ws = sw.invoke(drw, "InsertWeldSymbol3")
    drw.ClearSelection2(True)
    if ws is None:
        raise RuntimeError("SOLIDWORKS did not insert the weld symbol")
    # SetText(arrow side?, left text, symbol code, right text, stagger, contour)
    ws.SetText(not other_side, size, f"<WELD-{WELD_SYMBOLS[symbol]}>", length, "", 0)
    if all_around:
        ws.SetPeripheral(True)
    if field_weld:
        ws.SetFieldWeld(True)
    at = _place(sw.m0(ws, "GetAnnotation"), target, point, position)
    sw.m0(drw, "GraphicsRedraw2")
    return {"symbol": symbol, "size": size, "length": length, "view": target.Name, "position_mm": at}


@tool
def ordinate_dimensions(
    points: list[list[float]],
    direction: Literal["horizontal", "vertical"] = "horizontal",
    entity: Literal["vertex", "edge", "center"] = "vertex",
    view: str = "",
    position: list[float] | None = None,
) -> dict:
    """Ordinate (baseline) dimensions in a drawing view (default the
    first): the first model point ([x,y,z] mm) is the zero, the others get
    their distance from it, all along one line. entity picks vertices,
    edges or hole centres at the points. position is where the dimension
    line sits on the sheet (mm), default just outside the view."""
    if len(points) < 2:
        raise ValueError("Give the zero point and at least one more")
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    ents = [_view_entity(target, p, entity) for p in points]
    drw.ActivateView(target.Name)
    drw.ClearSelection2(True)
    for i, ent in enumerate(ents):
        _view_pick(drw, target, ent, 0, i > 0)
    x1, y1, x2, y2 = sw.m0(target, "GetOutline")
    if position:
        x, y = mm(position[0]), mm(position[1])
    elif direction == "horizontal":
        x, y = _beside_view(drw, target, "top", 0.01 + 0.007 * _dimension_count(target, "top"), (x1 + x2) / 2)
    else:
        x, y = _beside_view(drw, target, "left", 0.01 + 0.007 * _dimension_count(target, "left"), (y1 + y2) / 2)
    kind = sw.constant("swHorizontalOrdinate" if direction == "horizontal" else "swVerticalOrdinate")
    app = sw.app()
    prompt = app.GetUserPreferenceToggle(INPUT_DIM_ON_CREATE)
    app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, False)
    try:
        before = _dimension_count(target)
        drw.AddOrdinateDimension2(kind, x, y, 0)
    finally:
        app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, prompt)
    drw.ClearSelection2(True)
    made = _dimension_count(target) - before
    if made <= 0:
        raise RuntimeError("SOLIDWORKS added no ordinate dimensions; check the picked geometry")
    return {"view": target.Name, "dimensions": made, "direction": direction}


@tool
def break_view(start: float, end: float, axis: Literal["x", "y", "z"] = "x", view: str = "",
               style: Literal["straight", "zigzag", "curve", "small_zigzag", "jagged"] = "zigzag") -> dict:
    """Shorten a long part in a drawing view (default the first): the stretch
    between model coordinates start and end (mm) along `axis` is cut out and
    the ends are drawn with break lines; dimensions still show true length."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    k = "xyz".index(axis)
    p1, p2 = [0.0, 0.0, 0.0], [0.0, 0.0, 0.0]
    p1[k], p2[k] = start, end
    a, b = _to_sheet(target, p1), _to_sheet(target, p2)
    cx, cy = target.Position
    # break lines across the direction the axis runs on the sheet; positions relative to the view centre
    horizontal_run = abs(b[0] - a[0]) >= abs(b[1] - a[1])
    orient = 2 if horizontal_run else 1  # swBreakLineVertical / swBreakLineHorizontal
    pos = (a[0] - cx, b[0] - cx) if horizontal_run else (a[1] - cy, b[1] - cy)
    kinds = {"straight": 1, "zigzag": 2, "curve": 3, "small_zigzag": 4, "jagged": 5}  # swBreakLineStyle_e
    drw.ClearSelection2(True)
    drw.ActivateView(target.Name)
    line = target.InsertBreak3(orient, min(pos), max(pos), kinds[style], 0, False)
    if line is None:
        raise RuntimeError("SOLIDWORKS did not place the break lines; are start/end inside the part?")
    drw.ClearSelection2(True)
    drw.Extension.SelectByID2(target.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, sw.nothing(), 0)
    sw.invoke(drw, "BreakView")
    drw.ClearSelection2(True)
    drw.ForceRebuild3(False)
    return {"view": target.Name, "removed_mm": abs(end - start), "axis": axis, "style": style}


def _edge_ends(edge: Any) -> tuple[tuple, tuple] | None:
    """Model start and end points (m) of an edge, from its curve parameters."""
    try:
        p = sw.m0(edge, "GetCurveParams2")
        return tuple(p[0:3]), tuple(p[3:6])
    except Exception:
        return None


def _full_circles(view: Any) -> list[tuple[Any, tuple[float, float]]]:
    """(edge, sheet centre) of each hole seen in a view: full circles only
    (fillet arcs are skipped), one per centre (both faces of a plate show
    the same circle)."""
    out: list = []
    for edge in view.GetVisibleEntities2(sw.nothing(), sw.constant("swViewEntityType_Edge")) or ():
        try:
            curve = sw.m0(edge, "GetCurve")
            if not sw.invoke(curve, "IsCircle"):
                continue
            params = sw.m0(edge, "GetCurveParams2")
            if abs(abs(params[7] - params[6]) - 2 * math.pi) > 1e-3:  # start/end parameter: a partial arc
                continue
            centre = _to_sheet(view, [c / MM for c in curve.CircleParams[:3]])
        except Exception:
            continue
        if all(math.dist(centre, c) > 1e-4 for _, c in out):
            out.append((edge, centre))
    return out


def _lower_left_vertex(view: Any) -> Any:
    best, key = None, math.inf
    for vx in view.GetVisibleEntities2(sw.nothing(), sw.constant("swViewEntityType_Vertex")) or ():
        x, y = _to_sheet(view, [c / MM for c in _vertex_point(view, vx)])
        if x + y < key:
            best, key = vx, x + y
    return best


def _ordinate(drw: Any, view: Any, origin: Any, ents: list, direction: str, x: float, y: float) -> int:
    drw.ActivateView(view.Name)
    drw.ClearSelection2(True)
    _view_pick(drw, view, origin, 0, False)
    for ent in ents:
        _view_pick(drw, view, ent, 0, True)
    before = _dimension_count(view)
    kind = sw.constant("swHorizontalOrdinate" if direction == "horizontal" else "swVerticalOrdinate")
    app = sw.app()
    prompt = app.GetUserPreferenceToggle(INPUT_DIM_ON_CREATE)
    app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, False)
    try:
        drw.AddOrdinateDimension2(kind, x, y, 0)
    finally:
        app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, prompt)
    drw.ClearSelection2(True)
    return _dimension_count(view) - before


@tool
def dimension_holes(view: str = "") -> dict:
    """Position every hole seen in a drawing view (default the first view
    that shows holes face on) with ordinate dimensions from the view's lower
    left corner: one horizontal row above the view and one vertical row to
    its left, each hole centre once (holes sharing an x or y share it)."""
    drw = sw.active_doc(DOC_DRAWING)
    views = [_drawing_view(drw, view)] if view else _drawing_views(drw)
    target, holes = None, []
    for v in views:
        holes = _full_circles(v)
        if holes:
            target = v
            break
    if target is None:
        raise LookupError("No view shows holes as circles" + (f" ({view})" if view else ""))
    origin = _lower_left_vertex(target)
    if origin is None:
        raise LookupError(f"No corner in {target.Name} to measure from")

    def distinct(axis: int) -> list:
        picked, seen = [], []
        for edge, c in sorted(holes, key=lambda h: h[1][axis]):
            if all(abs(c[axis] - s) > 1e-4 for s in seen):
                picked.append(edge)
                seen.append(c[axis])
        return picked

    x1, y1, x2, y2 = sw.m0(target, "GetOutline")
    # beyond any dimensions already on that side, or SOLIDWORKS drops the new row
    top = 0.012 + 0.008 * _dimension_count(target, "top")
    left = 0.012 + 0.008 * _dimension_count(target, "left")
    across = _ordinate(drw, target, origin, distinct(0), "horizontal",
                       *_beside_view(drw, target, "top", top, (x1 + x2) / 2))
    up = _ordinate(drw, target, origin, distinct(1), "vertical",
                   *_beside_view(drw, target, "left", left, (y1 + y2) / 2))
    drw.ForceRebuild3(False)
    return {"view": target.Name, "holes": len(holes), "horizontal_dimensions": across, "vertical_dimensions": up}


@tool
def flat_pattern_view(position: list[float] | None = None, scale: str = "", hide_bend_lines: bool = False,
                      flip: bool = False, dimension: bool = True, model_path: str = "") -> dict:
    """Add the flat pattern of a sheet metal part (the one the drawing
    shows, or model_path) as a drawing view, with bend lines unless
    hide_bend_lines. dimension adds the overall flat width and height.
    position is the view centre on the sheet (mm), default a free spot."""
    drw = sw.active_doc(DOC_DRAWING)
    views = _drawing_views(drw)
    if not model_path:
        if not views:
            raise ValueError("The drawing has no views yet; pass model_path")
        model_path = sw.m0(views[0], "GetReferencedModelName")
    config = views[0].ReferencedConfiguration if views else ""
    sheet_w, sheet_h = _sheet_size(drw)
    x, y = (mm(position[0]), mm(position[1])) if position else (sheet_w / 2, sheet_h / 2)
    new = _call_first(drw, [
        ("CreateFlatPatternViewFromModelView3", (model_path, config, x, y, 0, hide_bend_lines, flip)),
        ("CreateFlatPatternViewFromModelView2", (model_path, config, x, y, 0, hide_bend_lines)),
    ])
    if new is None:
        raise RuntimeError("Could not create a flat pattern view; is the model a sheet metal part?")
    if scale:
        _set_view_scale(new, scale)
    drw.ForceRebuild3(False)
    if not position:
        a1, b1, a2, b2 = sw.m0(new, "GetOutline")
        new.Position = sw.doubles(_free_spot(drw, a2 - a1, b2 - b1, exclude=new.Name, extra=_table_boxes(drw)))
        drw.ForceRebuild3(False)
    result = _view_info(new)
    if dimension:
        result["dimensions"] = _overall_dimensions(drw, new)
    return result


def _overall_dimensions(drw: Any, view: Any) -> dict:
    """Overall width and height of what a view shows, between its outermost
    straight edges."""
    lines = []
    for edge in view.GetVisibleEntities2(sw.nothing(), sw.constant("swViewEntityType_Edge")) or ():
        try:
            if not sw.invoke(sw.m0(edge, "GetCurve"), "IsLine"):
                continue
            ends = _edge_ends(edge)
            if ends is None:
                continue
            a, b = (_to_sheet(view, [c / MM for c in p]) for p in ends)
        except Exception:
            continue
        lines.append((edge, a, b))
    out = {}
    x1, y1, x2, y2 = sw.m0(view, "GetOutline")
    for axis, direction, add in ((0, "horizontal", "AddHorizontalDimension2"), (1, "vertical", "AddVerticalDimension2")):
        # edges running across this axis: the two lying furthest apart along it
        across = [l for l in lines if abs(l[1][axis] - l[2][axis]) < 1e-5]
        if len(across) < 2:
            continue
        low = min(across, key=lambda l: l[1][axis])
        high = max(across, key=lambda l: l[1][axis])
        drw.ActivateView(view.Name)
        drw.ClearSelection2(True)
        _view_pick(drw, view, low[0], 0, False)
        _view_pick(drw, view, high[0], 0, True)
        px, py = (_beside_view(drw, view, "top", 0.01, (x1 + x2) / 2) if axis == 0
                  else _beside_view(drw, view, "left", 0.01, (y1 + y2) / 2))
        app = sw.app()
        prompt = app.GetUserPreferenceToggle(INPUT_DIM_ON_CREATE)
        app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, False)
        try:
            dim = getattr(drw, add)(px, py, 0)
        finally:
            app.SetUserPreferenceToggle(INPUT_DIM_ON_CREATE, prompt)
        drw.ClearSelection2(True)
        if dim is not None:
            out[direction] = round(dim.GetDimension2(0).Value, 3)
    return out


# --- Patterns, reference geometry and other features ------------------------

# COM errors meaning "this API has a different signature here": bad parameter
# count, member not found, type mismatch.
_SIGNATURE_ERRORS = {-2147352562, -2147352573, -2147352571}


def _call_first(obj: Any, variants: list[tuple[str, tuple]]) -> Any:
    """Call the first (method, args) the installed SOLIDWORKS accepts. Some
    feature methods gained arguments between releases."""
    last: Exception | None = None
    for name, args in variants:
        try:
            return getattr(obj, name)(*args)
        except AttributeError as exc:
            last = exc
        except Exception as exc:
            code = getattr(exc, "hresult", None) or (exc.args[0] if exc.args else None)
            if code not in _SIGNATURE_ERRORS:
                raise
            last = exc
    raise RuntimeError(f"none of {[v[0] for v in variants]} accepted these arguments ({last})")


def _select_direction(doc: Any, direction: str, mark: int, append: bool) -> None:
    """Select a direction reference: "x"/"y"/"z" (via the default plane whose
    normal it is), a plane/axis feature name, or "edge:x,y,z" (mm)."""
    if direction.startswith("edge:"):
        sw.select_at(doc, "EDGE", [float(v) for v in direction[5:].split(",")], append, mark)
        return
    plane = {"x": "right", "y": "top", "z": "front"}.get(direction.lower())
    feat = sw.standard_plane(doc, plane) if plane else sw.feature(doc, direction)
    if not feat.Select2(append, mark):
        raise RuntimeError(f"Could not select direction {direction!r}")


def _select_features(doc: Any, names: list[str], mark: int, append: bool = True) -> None:
    for name in names:
        sw.select_feature(doc, name, append=append, mark=mark)
        append = True


@tool
def linear_pattern(
    features: list[str],
    direction: str,
    count: int,
    spacing: float,
    direction2: str = "",
    count2: int = 1,
    spacing2: float = 10.0,
    reverse: bool = False,
    reverse2: bool = False,
) -> dict:
    """Repeat features (by name, e.g. ["컷-돌출1"]) in a line. direction is
    "x"/"y"/"z", a plane/axis name, or "edge:x,y,z" (a model edge through
    that point, mm). count includes the original; spacing in mm. An optional
    second direction makes a grid."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    _select_direction(doc, direction, 1, False)
    if direction2:
        _select_direction(doc, direction2, 2, True)
    _select_features(doc, features, 4)
    n2 = count2 if direction2 else 1
    base = (count, mm(spacing), n2, mm(spacing2), reverse, reverse2, "NULL", "NULL",
            False, False, False, False, False, False, True, True, False, False, 0, 0)
    feat = _call_first(doc.FeatureManager, [
        ("FeatureLinearPattern5", base + (False, False)),
        ("FeatureLinearPattern4", base),
    ])
    return _feature_result(feat, "linear pattern")


def _axis(doc: Any, axis: str) -> Any:
    """A new reference axis along x/y/z through the origin, at the
    intersection of two default planes."""
    planes = {"x": ("top", "front"), "y": ("front", "right"), "z": ("top", "right")}[axis]
    doc.ClearSelection2(True)
    for i, p in enumerate(planes):
        sw.standard_plane(doc, p).Select2(i > 0, 0)
    if not doc.InsertAxis2(True):
        raise RuntimeError(f"Could not create the {axis} axis")
    doc.ClearSelection2(True)
    return sw.last_feature_of_type(doc, "RefAxis")


@tool
def circular_pattern(features: list[str], count: int, axis: str = "y", angle: float = 360.0, reverse: bool = False) -> dict:
    """Repeat features around an axis: "x"/"y"/"z" (through the origin; a
    reference axis is created), an axis feature name, or "face:x,y,z" (the
    axis of the cylindrical face through that point, mm). count includes the
    original; the instances are spread evenly over `angle` degrees."""
    doc = sw.active_doc(DOC_PART)
    if axis.lower() in ("x", "y", "z"):
        axis_feat = _axis(doc, axis.lower())
        doc.ClearSelection2(True)
        axis_feat.Select2(False, 1)
    elif axis.startswith("face:"):
        doc.ClearSelection2(True)
        sw.select_at(doc, "FACE", [float(v) for v in axis[5:].split(",")], False, 1)
    else:
        doc.ClearSelection2(True)
        sw.select_feature(doc, axis, mark=1)
    _select_features(doc, features, 4)
    total = deg(angle)
    feat = _call_first(doc.FeatureManager, [
        ("FeatureCircularPattern5", (count, total, reverse, "NULL", False, True, False, False, False, False, 1, total, "NULL", False)),
        ("FeatureCircularPattern4", (count, total, reverse, "NULL", False, True, False)),
    ])
    return _feature_result(feat, "circular pattern")


@tool
def sketch_pattern(features: list[str], sketch_name: str, use_centroid: bool = True) -> dict:
    """Copy features (e.g. a hole cut) to every point of a sketch (made with
    sketch() "point" entities): the seed's centroid (or, use_centroid=false,
    its sketch origin) lands on each point. For irregular hole layouts."""
    doc = sw.active_doc(DOC_PART)
    before = sw.m0(doc, "GetMassProperties")
    doc.ClearSelection2(True)
    _select_features(doc, features, 4, append=False)
    sw.select_feature(doc, sketch_name, append=True, mark=64)
    feat = doc.FeatureManager.FeatureSketchDrivenPattern(use_centroid, False)
    result = _feature_result(feat, "sketch driven pattern")
    after = sw.m0(doc, "GetMassProperties")
    if before and after:
        result["volume_change_mm3"] = round((after[3] - before[3]) / MM**3, 2)
    return result


@tool
def mirror_part(plane: str = "right", face_point: list[float] | None = None, break_link: bool = False) -> dict:
    """Make the opposite-hand version of the active (saved) part as a new
    part, mirrored about "front"/"top"/"right", a plane name or the planar
    face at face_point. It stays linked to the original (changes follow)
    unless break_link. The new part is active and unsaved: save it."""
    import pythoncom
    from win32com.client import VARIANT

    doc = sw.active_doc(DOC_PART)
    if not sw.m0(doc, "GetPathName"):
        raise ValueError("Save the part first; the mirrored part refers to its file")
    source = sw.m0(doc, "GetTitle")
    doc.ClearSelection2(True)
    if face_point:
        sw.select_at(doc, "FACE", face_point, False)
    else:
        key = plane.lower()
        _select_plane(doc, sw.standard_plane(doc, key) if key in sw.PLANE_NAMES else sw.feature(doc, plane))
    options = 0
    for name in ("ImportSolids", "ImportPlanes", "ImportAxes", "ImportCosmeticThreads", "ImportCustomProperties",
                 "ImportHoleWizardData", "ImportPartMaterial"):
        options |= sw.constant(f"swMirrorPartOptions_{name}")
    made = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_DISPATCH, None)
    doc.MirrorPart2(break_link, options, made)
    new = sw.app().ActiveDoc
    if new is None or sw.m0(new, "GetTitle") == source:
        raise RuntimeError("SOLIDWORKS did not make the mirrored part; check the plane or face")
    result = {"part": sw.m0(new, "GetTitle"), "mirrored_from": source}
    props = sw.m0(new.Extension, "CreateMassProperty")
    if props is not None:
        result["volume_mm3"] = round(props.Volume / MM**3, 3)
        result["center_of_mass_mm"] = _mm3(props.CenterOfMass)
    return result


@tool
def mirror(features: list[str], plane: str = "right", face_point: list[float] | None = None) -> dict:
    """Mirror features (by name) about "front"/"top"/"right", a plane name,
    or the planar face at face_point (mm)."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    if face_point:
        sw.select_at(doc, "FACE", face_point, False, 2)
    else:
        key = plane.lower()
        target = sw.standard_plane(doc, key) if key in sw.PLANE_NAMES else sw.feature(doc, plane)
        target.Select2(False, 2)
    _select_features(doc, features, 1)
    feat = _call_first(doc.FeatureManager, [
        ("InsertMirrorFeature2", (False, False, False, False, 0)),
        ("InsertMirrorFeature", (False, False, False, False)),
    ])
    return _feature_result(feat, "mirror")


@tool
def shell(thickness: float, open_face_points: list[list[float]] | None = None, outward: bool = False) -> dict:
    """Hollow the part to a wall `thickness` (mm), removing the faces through
    the given points (e.g. the top face to make an open box)."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    for i, p in enumerate(open_face_points or []):
        sw.select_at(doc, "FACE", p, append=i > 0, mark=0)
    before = doc.FeatureByPositionReverse(0).Name
    doc.InsertFeatureShell(mm(thickness), outward)
    feat = doc.FeatureByPositionReverse(0)
    return _feature_result(feat if feat.Name != before else None, "shell")


@tool
def offset_plane(distance: float, plane: str = "top", face_point: list[float] | None = None, reverse: bool = False) -> dict:
    """Create a reference plane `distance` mm from a plane ("front"/"top"/
    "right" or a name) or from the planar face at face_point; sketch on it by
    passing the returned name to sketch(plane=...)."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    if face_point:
        sw.select_at(doc, "FACE", face_point, False, 0)
    else:
        key = plane.lower()
        target = sw.standard_plane(doc, key) if key in sw.PLANE_NAMES else sw.feature(doc, plane)
        target.Select2(False, 0)
    # swRefPlaneReferenceConstraint_Distance, plus OptionFlip to go the other way
    feat = doc.FeatureManager.InsertRefPlane(8 | (256 if reverse else 0), mm(distance), 0, 0, 0, 0)
    return _feature_result(feat, "reference plane")


@tool
def reference_axis(axis: Literal["x", "y", "z"]) -> dict:
    """Create a reference axis along x/y/z through the origin (for patterns,
    revolves or mates)."""
    return _feature_result(_axis(sw.active_doc(DOC_PART), axis), "reference axis")


def _rotation_axes(rotation: list[float]) -> list[float]:
    """The X, Y and Z axes (one after another, as in MathTransform
    ArrayData) of a rotation about the X, then Y, then Z axis, in degrees."""
    cx, cy, cz = (math.cos(math.radians(a)) for a in rotation)
    sx, sy, sz = (math.sin(math.radians(a)) for a in rotation)
    r = [[cz * cy, cz * sy * sx - sz * cx, cz * sy * cx + sz * sx],
         [sz * cy, sz * sy * sx + cz * cx, sz * sy * cx - cz * sx],
         [-sy, cy * sx, cy * cx]]
    return [r[row][col] for col in range(3) for row in range(3)]


def _axes_info(data: list[float]) -> dict:
    return {"origin_mm": [round(v / MM, 3) + 0.0 for v in data[9:12]],
            **{f"{n}_axis": [round(v, 4) + 0.0 for v in data[3 * i:3 * i + 3]] for i, n in enumerate("xyz")}}


@tool
def coordinate_system(origin: list[float] | None = None, rotation: list[float] | None = None) -> dict:
    """Reference coordinate system at origin [x,y,z] mm, turned by
    rotation [about x, about y, about z] degrees (applied in that order),
    e.g. for exports or measurements relative to it."""
    doc = sw.active_doc(DOC_PART)
    x, y, z = (mm(v) for v in (origin or [0, 0, 0]))
    ax, ay, az = (deg(v) for v in (rotation or [0, 0, 0]))
    doc.ClearSelection2(True)
    feat = doc.FeatureManager.CreateCoordinateSystemUsingNumericalValues(True, x, y, z, True, ax, ay, az)
    result = _feature_result(feat, "coordinate system")
    xf = doc.Extension.GetCoordinateSystemTransformByName(feat.Name)
    if xf is not None:
        result.update(_axes_info(list(sw.m0(xf, "ArrayData"))))
    return result


@tool
def reference_point(point: list[float], entity: Literal["face", "edge"] = "face") -> dict:
    """Reference point at the centre of the face, or of the circular edge,
    through point ([x,y,z] mm), e.g. to mate to or measure from."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_at(doc, entity, point, False)
    kind = sw.constant("swRefPointFaceCenter" if entity == "face" else "swRefPointCenterEdge")
    last = doc.FeatureByPositionReverse(0).Name
    doc.FeatureManager.InsertReferencePoint(kind, 0, 0, 1)
    newest = doc.FeatureByPositionReverse(0)  # the call returns an array of points, not the feature
    if newest.Name == last or sw.m0(newest, "GetTypeName2") != "RefPoint":
        raise RuntimeError(f"SOLIDWORKS made no reference point at the {entity} through {point} mm")
    result = _feature_result(newest, "reference point")
    p = sw.m0(sw.m0(newest, "GetSpecificFeature2"), "GetRefPoint")
    result["point_mm"] = [round(v / MM, 3) + 0.0 for v in list(sw.m0(p, "ArrayData"))[:3]]
    return result


@tool
def draft(angle: float, neutral_face_point: list[float], face_points: list[list[float]], reverse: bool = False) -> dict:
    """Neutral-plane draft: tilt the faces through face_points ([[x,y,z], ...]
    mm) by `angle` degrees, pivoting about the planar face (or plane) at
    neutral_face_point. reverse flips the pull direction."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_at(doc, "FACE", neutral_face_point, False, 1)
    for p in face_points:
        sw.select_at(doc, "FACE", p, True, 2)
    feat = doc.FeatureManager.InsertMultiFaceDraft(deg(angle), reverse, False, 0, False, False)
    return _feature_result(feat, "draft")


@tool
def cosmetic_thread(edge_point: list[float], size: str, depth: float = 0.0, note: str = "") -> dict:
    """Cosmetic thread on a cylinder (a shaft's outside or a plain hole),
    starting at the circular edge through edge_point ([x,y,z] mm). size is
    ISO metric like "M8" (coarse pitch) or "M8x1"; depth in mm, 0 = to the
    end of the cylinder. Drawings show it as a thread with a callout."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_at(doc, "EDGE", edge_point, False, 0)
    edge = doc.SelectionManager.GetSelectedObject6(1, -1)
    radius = sw.m0(edge, "GetCurve").CircleParams[6] if edge is not None else 0
    if not radius:
        raise ValueError(f"The edge at {edge_point} mm is not circular")
    nominal = float(size.upper().lstrip("M").split("X")[0])
    pitch = float(size.upper().split("X")[1]) if "X" in size.upper() else ISO_PITCH.get(f"M{nominal:g}", 0)
    outside = abs(radius * 2 / MM - nominal) < abs(radius * 2 / MM - (nominal - pitch))
    # the thread's other diameter: minor for a shaft, major for a hole
    other = nominal - 1.0825 * pitch if outside else nominal
    length = depth or _cylinder_length(edge)
    callout = note or (f"M{nominal:g}x{pitch:g}" if pitch else f"M{nominal:g}")
    feat = doc.FeatureManager.InsertCosmeticThread3(-2, "", "", mm(other), 0, mm(length), callout)
    return _feature_result(feat, "cosmetic thread") | {"thread": callout, "on": "shaft" if outside else "hole",
                                                         "length_mm": round(length, 3)}


def _cylinder_length(edge: Any) -> float:
    """Length (mm) of the cylindrical face behind a circular edge."""
    for face in sw.m0(edge, "GetTwoAdjacentFaces2") or ():
        surf = sw.m0(face, "GetSurface")
        if face is not None and sw.invoke(surf, "IsCylinder"):
            box = sw.m0(face, "GetBox")
            axis = surf.CylinderParams[3:6]
            spans = [abs(box[i + 3] - box[i]) for i in range(3)]
            k = max(range(3), key=lambda i: abs(axis[i]))
            return round(spans[k] / MM, 3)
    raise ValueError("No cylindrical face next to that edge")


# ISO 261 coarse pitches
ISO_PITCH = {"M2": 0.4, "M2.5": 0.45, "M3": 0.5, "M4": 0.7, "M5": 0.8, "M6": 1.0, "M8": 1.25, "M10": 1.5,
             "M12": 1.75, "M14": 2.0, "M16": 2.0, "M20": 2.5, "M24": 3.0, "M30": 3.5}


@tool
def angle_plane(angle: float, axis: str = "y", plane: str = "front") -> dict:
    """Reference plane through an axis at `angle` degrees to a plane. axis is
    "x"/"y"/"z" (through the origin), an axis feature name or "edge:x,y,z"
    (a straight model edge through that point, mm); plane is
    "front"/"top"/"right" or a plane name, containing or parallel to the axis."""
    doc = sw.active_doc(DOC_PART)
    axis_feat = _axis(doc, axis.lower()) if axis.lower() in ("x", "y", "z") else None
    doc.ClearSelection2(True)
    # first reference (mark 0): the plane, at an angle; second (mark 1): the axis, coincident
    key = plane.lower()
    target = sw.standard_plane(doc, key) if key in sw.PLANE_NAMES else sw.feature(doc, plane)
    target.Select2(False, 0)
    if axis_feat is not None:
        axis_feat.Select2(True, 1)
    elif axis.startswith("edge:"):
        sw.select_at(doc, "EDGE", [float(v) for v in axis[5:].split(",")], True, 1)
    else:
        sw.select_feature(doc, axis, append=True, mark=1)
    feat = doc.FeatureManager.InsertRefPlane(16, deg(angle), 4, 0, 0, 0)
    return _feature_result(feat, "angled reference plane")


@tool
def dome(face_point: list[float], height: float, reverse: bool = False) -> dict:
    """Dome of `height` mm on the planar face at face_point (reverse makes a
    dent instead)."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_at(doc, "FACE", face_point, False, 1)  # InsertDome takes the face with mark 1
    before = doc.FeatureByPositionReverse(0).Name
    doc.InsertDome(mm(height), reverse, False)
    feat = doc.FeatureByPositionReverse(0)
    return _feature_result(feat if feat.Name != before else None, "dome")


@tool
def sketch_from_face(face_point: list[float], offset: float = 0.0, plane: str = "",
                     inner_loops: bool = True, both_directions: bool = False) -> dict:
    """Sketch the outline of the face at face_point ([x,y,z] mm): offset 0
    copies its edges (convert entities), otherwise they are offset by
    `offset` mm (positive = outward, negative = inward). The sketch goes on
    that face, or on `plane` ("front"/"top"/"right" or a plane name) when
    given, projected onto it. inner_loops also takes the face's holes.
    Extrude or cut the result like any sketch."""
    doc = sw.active_doc(DOC_PART)
    sw.close_open_sketch(doc)
    doc.ClearSelection2(True)
    if plane:
        key = plane.lower()
        _select_plane(doc, sw.standard_plane(doc, key) if key in sw.PLANE_NAMES else sw.feature(doc, plane))
    else:
        sw.select_at(doc, "FACE", face_point, False)
    sm = doc.SketchManager
    sm.InsertSketch(True)
    if sm.ActiveSketch is None:
        raise RuntimeError("SOLIDWORKS did not open a sketch there")
    try:
        doc.ClearSelection2(True)
        # SketchUseEdge3's second argument takes the inner loops *instead of* the outer one
        sw.select_at(doc, "FACE", face_point, False)
        if not sm.SketchUseEdge3(False, False):
            raise RuntimeError(f"Could not convert the edges of the face at {face_point} mm")
        if inner_loops:
            doc.ClearSelection2(True)
            sw.select_at(doc, "FACE", face_point, False)
            sm.SketchUseEdge3(False, True)  # False when the face has no holes
        if offset:
            converted = list(sw.m0(sm.ActiveSketch, "GetSketchSegments") or ())
            doc.ClearSelection2(True)
            for i, seg in enumerate(converted):
                seg.Select4(i > 0, sw.nothing())
            # positive offsets go outward from the closed loops
            if not sm.SketchOffset2(mm(offset), both_directions, True, 0, False, True):
                raise RuntimeError("SOLIDWORKS could not offset the outline by that distance")
            doc.ClearSelection2(True)
            for i, seg in enumerate(converted):  # keep only the offset copy
                seg.Select4(i > 0, sw.nothing())
            sw.m0(doc, "EditDelete")
        count = len(sw.m0(sm.ActiveSketch, "GetSketchSegments") or ())
    finally:
        doc.ClearSelection2(True)
        sm.InsertSketch(True)
    feat = sw.last_feature_of_type(doc, "ProfileFeature")
    return {"sketch": feat.Name, "segments": count, "offset_mm": offset}


@tool
def list_bodies() -> list[dict]:
    """Solid bodies of the active part (several after extrude(merge=false),
    split or mirror), with volume and bounding box, for combine_bodies."""
    doc = sw.active_doc(DOC_PART)
    out = []
    for b in doc.GetBodies2(0, True) or ():
        box = b.GetBodyBox()
        props = b.GetMassProperties(1.0)  # density 1: [cx, cy, cz, volume, area, mass, ...]
        out.append({
            "name": b.Name,
            "volume_mm3": round(props[3] / MM**3, 3) if props else None,
            "box_mm": {"min": [round(v / MM, 3) for v in box[:3]], "max": [round(v / MM, 3) for v in box[3:]]},
        })
    return out


def _body_at(doc: Any, point: list[float]) -> Any:
    dist, face = sw.nearest(doc, "face", point)
    if face is None or dist > sw.TOLERANCE:
        raise LookupError(f"No body surface within 0.5 mm of {point} mm")
    return sw.m0(face, "GetBody")


@tool
def combine_bodies(operation: Literal["add", "subtract", "common"], main_body_point: list[float],
                   tool_body_points: list[list[float]]) -> dict:
    """Combine solid bodies of a multibody part. Each body is picked by a
    point on its surface ([x,y,z] mm). add merges them all; subtract removes
    the tool bodies from the main body; common keeps only the overlap."""
    import pythoncom
    from win32com.client import VARIANT

    doc = sw.active_doc(DOC_PART)
    main = _body_at(doc, main_body_point)
    tools = [_body_at(doc, p) for p in tool_body_points]
    if any(t.Name == main.Name for t in tools):
        raise ValueError("A tool body point picks the main body itself")
    # swBodyOperationType_e (SWBODYADD/CUT/INTERSECT), not swCombineBodiesOperation_e
    op = sw.constant({"add": "SWBODYADD", "subtract": "SWBODYCUT", "common": "SWBODYINTERSECT"}[operation])
    before = len(doc.GetBodies2(0, True) or ())
    doc.ClearSelection2(True)
    def bodies(items):
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, items)

    # subtract takes the main body apart; add and common want every body as a tool
    attempts = [(main, tools)] if operation == "subtract" else [(sw.nothing(), [main] + tools), (main, tools)]
    feat = None
    for first, rest in attempts:
        feat = doc.FeatureManager.InsertCombineFeature(op, first, bodies(rest))
        if feat is not None:
            break
    result = _feature_result(feat, f"combine ({operation})")
    result["bodies_before"], result["bodies_after"] = before, len(doc.GetBodies2(0, True) or ())
    return result


@tool
def import_file(path: str) -> dict:
    """Open a neutral CAD file (.step/.stp, .igs/.iges, .x_t/.x_b, .sat,
    .stl) as a new SOLIDWORKS part or assembly. Save it afterwards with
    save(path=...sldprt) to keep it."""
    target = str(Path(path))
    if not os.path.exists(target):
        raise FileNotFoundError(target)
    app = sw.app()
    data = sw.invoke(app, "GetImportFileData", target)
    errors = sw.out_int()
    doc = app.LoadFile4(target, "r", data, errors)
    if doc is None:
        raise RuntimeError(f"Import failed: {sw.load_error(errors.value)} - {target}")
    info = sw.doc_info(doc)
    if sw.m0(doc, "GetType") == DOC_PART:
        info["bodies"] = len(doc.GetBodies2(0, True) or ())
    elif sw.m0(doc, "GetType") == DOC_ASSEMBLY:
        info["components"] = len(doc.GetComponents(True) or ())
    return info


@tool
def sketch_corners(corner_points: list[list[float]], radius: float = 0.0, chamfer: float = 0.0,
                   sketch_name: str = "") -> dict:
    """Round (radius mm) or bevel (chamfer mm, equal distances) corners of an
    existing sketch (default the most recent), picked by their 2D sketch
    coordinates [[x,y], ...] mm, like the coordinates given to sketch()."""
    if (radius > 0) == (chamfer > 0):
        raise ValueError("Give either radius or chamfer")
    doc = sw.active_doc(DOC_PART)
    feat = sw.feature(doc, sketch_name) if sketch_name else sw.last_feature_of_type(doc, "ProfileFeature")
    sk = sw.m0(feat, "GetSpecificFeature2")
    points = list(sw.m0(sk, "GetSketchPoints2") or ())
    picked = []
    for x, y in corner_points:
        best = min(points, key=lambda p: math.hypot(p.X - mm(x), p.Y - mm(y)), default=None)
        if best is None or math.hypot(best.X - mm(x), best.Y - mm(y)) > sw.TOLERANCE:
            raise LookupError(f"No sketch corner at [{x}, {y}] mm in {feat.Name}")
        picked.append(best)
    sw.close_open_sketch(doc)
    doc.ClearSelection2(True)
    feat.Select2(False, 0)
    sw.m0(doc, "EditSketch")
    sm = doc.SketchManager
    try:
        doc.ClearSelection2(True)
        for i, p in enumerate(picked):
            p.Select4(i > 0, sw.nothing())
        if radius > 0:
            # keep the corner's constrained geometry (swConstrainedCornerKeepGeometry)
            made = sm.CreateFillet(mm(radius), 1)
        else:
            made = sm.CreateChamfer(2, mm(chamfer), mm(chamfer))  # swSketchChamfer_DistanceEqual
    finally:
        doc.ClearSelection2(True)
        sm.InsertSketch(True)
    if not made:
        raise RuntimeError("SOLIDWORKS could not round/bevel those corners (is the radius too big?)")
    _rebuild(doc)
    return {"sketch": feat.Name, "corners": len(picked), "radius_mm" if radius else "chamfer_mm": radius or chamfer}


@tool
def move_body(body_point: list[float], translate: list[float] | None = None, rotate: list[float] | None = None,
              about: list[float] | None = None, copy: bool = False, copies: int = 1) -> dict:
    """Move (or copy, copies times) the solid body through body_point
    ([x,y,z] mm on its surface) by translate [dx,dy,dz] mm and/or rotate
    [rx,ry,rz] degrees about the point `about` (default the origin)."""
    doc = sw.active_doc(DOC_PART)
    body = _body_at(doc, body_point)
    dx, dy, dz = (mm(v) for v in (translate or [0, 0, 0]))
    rx, ry, rz = (deg(v) for v in (rotate or [0, 0, 0]))
    px, py, pz = (mm(v) for v in (about or [0, 0, 0]))
    doc.ClearSelection2(True)
    data = sw.m0(doc.SelectionManager, "CreateSelectData")
    data.Mark = 1
    body.Select2(False, data)
    before = len(doc.GetBodies2(0, True) or ())
    feat = doc.FeatureManager.InsertMoveCopyBody2(dx, dy, dz, 0, px, py, pz, rx, ry, rz, copy, copies if copy else 0)
    doc.ClearSelection2(True)
    result = _feature_result(feat, "copy body" if copy else "move body")
    result["bodies"] = {"before": before, "after": len(doc.GetBodies2(0, True) or ())}
    return result


@tool
def face_color(face_points: list[list[float]], rgb: list[int]) -> dict:
    """Colour single faces of the active part (picked by points on them,
    [[x,y,z], ...] mm), e.g. a machined face [180, 180, 190]."""
    doc = sw.active_doc(DOC_PART)
    r, g, b = (max(0, min(255, int(v))) / 255 for v in rgb)
    done = 0
    for p in face_points:
        dist, face = sw.nearest(doc, "face", p)
        if face is None or dist > sw.TOLERANCE:
            raise LookupError(f"No face within 0.5 mm of {p} mm")
        # red, green, blue, ambient, diffuse, specular, shininess, transparency, emission; this configuration
        face.SetMaterialPropertyValues2(sw.doubles((r, g, b, 1, 1, 0.5, 0.4, 0, 0)), 1, sw.empty())  # returns nothing
        now = sw.m0(face, "MaterialPropertyValues") or ()
        if len(now) >= 3 and all(abs(now[i] - c) < 0.01 for i, c in enumerate((r, g, b))):
            done += 1
    sw.m0(doc, "GraphicsRedraw2")
    return {"faces_coloured": done, "rgb": list(rgb)}


@tool
def variable_fillet(edge_point: list[float], start_radius: float, end_radius: float) -> dict:
    """Variable-radius fillet on the edge through edge_point ([x,y,z] mm),
    growing from start_radius at one end of the edge to end_radius at the
    other (mm)."""
    import pythoncom
    from win32com.client import VARIANT

    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_at(doc, "EDGE", edge_point, False, 1)
    e = sw.empty
    radii = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [mm(start_radius), mm(end_radius)])  # per vertex
    before = sw.m0(doc, "GetMassProperties")
    feat = doc.FeatureManager.FeatureFillet3(195, mm(start_radius), 0, 0, 1, 0, 0, radii, e(), e(), e(), e(), e(), e())
    result = _feature_result(feat, "variable fillet")
    after = sw.m0(doc, "GetMassProperties")
    if before and after:
        result["volume_removed_mm3"] = round((before[3] - after[3]) / MM**3, 2)
    return result


@tool
def face_fillet(face1_point: list[float], face2_point: list[float], radius: float) -> dict:
    """Face fillet of `radius` mm blending two faces (points on each,
    [x,y,z] mm); unlike an edge fillet the faces need not share an edge."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_at(doc, "FACE", face1_point, False, 0)
    sw.select_at(doc, "FACE", face2_point, True, 0)
    e = sw.empty
    feat = doc.FeatureManager.FeatureFillet3(195, mm(radius), 0, 0, 2, 0, 0, e(), e(), e(), e(), e(), e(), e())  # face fillet
    return _feature_result(feat, "face fillet")


@tool
def edit_sketch(sketch_name: str, entities: list[dict]) -> dict:
    """Add entities to an existing sketch (name from list_features), in its
    own 2D coordinates (mm), with the same entity types as sketch(). The
    features built on the sketch update; e.g. add a circle to a plate's
    sketch to cut another hole in the same extrude."""
    doc = sw.active_doc(DOC_PART)
    feat = sw.feature(doc, sketch_name)
    sw.close_open_sketch(doc)
    doc.ClearSelection2(True)
    feat.Select2(False, 0)
    sw.m0(doc, "EditSketch")
    sm = doc.SketchManager
    if sm.ActiveSketch is None:
        raise RuntimeError(f"Could not open {sketch_name} for editing")
    failed = []
    sm.AddToDB = True
    try:
        for i, e in enumerate(entities):
            if _sketch_entity(sm, e, doc) is None:
                failed.append(i)
    finally:
        sm.AddToDB = False
        doc.ClearSelection2(True)
        sm.InsertSketch(True)  # leave the sketch; the model rebuilds
    _rebuild(doc)
    result = {"sketch": feat.Name, "added": len(entities) - len(failed)}
    if failed:
        result["failed_entities"] = failed
    return result


# --- Equations, global variables and configurations -------------------------


def _equations(eqm: Any) -> list[dict]:
    out = []
    for i in range(sw.m0(eqm, "GetCount")):
        try:
            value = eqm.Value(i)
        except Exception:
            value = None
        out.append({"index": i, "equation": eqm.Equation(i), "value": value})
    return out


@tool
def list_equations() -> list[dict]:
    """Global variables and equations of the active document."""
    return _equations(sw.m0(sw.active_doc(), "GetEquationMgr"))


def _eq_norm(text: str) -> str:
    """Equation text as SOLIDWORKS may rewrite it: spacing, case, units."""
    return re.sub(r"(?<=\d)mm\b", "", "".join(text.split()).lower())


def _replace_equation(eqm: Any, index: int, old: str, text: str) -> None:
    """Edit equation `index` in place (so dimensions referencing a global
    variable stay linked); if SOLIDWORKS refuses, leave the old one intact.
    Success means the stored text changed away from the old one, since
    SOLIDWORKS may reformat what it stores."""
    errors = []
    attempts = (
        ("Equation property", lambda: sw.put_indexed(eqm, "Equation", index, text)),
        ("SetEquationAndConfigurationOption",
         lambda: _call_first(eqm, [("SetEquationAndConfigurationOption", (index, text, 0, sw.nothing()))])),
    )
    for label, attempt in attempts:
        try:
            attempt()
        except Exception as exc:
            errors.append(f"{label}: {exc}")
            continue
        stored = eqm.Equation(index)
        if _eq_norm(stored) == _eq_norm(text) or _eq_norm(stored) != _eq_norm(old):
            return
        errors.append(f"{label}: no change")
    if _eq_norm(eqm.Equation(index)) != _eq_norm(old):
        sw.put_indexed(eqm, "Equation", index, old)
    raise RuntimeError(f"SOLIDWORKS rejected {text!r}; kept {old!r} ({'; '.join(errors)})")


@tool
def set_equation(lhs: str, rhs: str) -> list[dict]:
    """Add or replace an equation. Global variable: lhs="Width", rhs="120".
    Drive a dimension: lhs="D1@스케치2", rhs='"Width"' or '"Width" / 2 + 5'.
    Names are quoted automatically on the left; quote variable/dimension
    names yourself inside rhs. Values are in document units (mm)."""
    doc = sw.active_doc()
    eqm = sw.m0(doc, "GetEquationMgr")
    left = lhs if lhs.startswith('"') else f'"{lhs}"'
    text = f"{left} = {rhs}"
    for e in _equations(eqm):
        if e["equation"].split("=", 1)[0].strip().lower() == left.lower():
            _replace_equation(eqm, e["index"], e["equation"], text)
            break
    else:
        if eqm.Add2(-1, text, True) < 0:
            raise RuntimeError(f"SOLIDWORKS rejected {text!r}")
    try:
        sw.invoke(eqm, "EvaluateAll")
    except Exception:
        pass  # optional: the rebuild below evaluates equations as well
    _rebuild(doc)
    return _equations(eqm)


@tool
def delete_equation(lhs: str) -> list[dict]:
    """Remove the equation or global variable whose left side is `lhs`."""
    doc = sw.active_doc()
    eqm = sw.m0(doc, "GetEquationMgr")
    left = (lhs if lhs.startswith('"') else f'"{lhs}"').lower()
    for e in _equations(eqm):
        if e["equation"].split("=", 1)[0].strip().lower() == left:
            eqm.Delete(e["index"])
            _rebuild(doc)
            return _equations(eqm)
    raise LookupError(f"No equation for {lhs!r}")


@tool
def configurations(add: str = "", activate: str = "") -> dict:
    """List configurations; optionally add one (it becomes active) or
    activate an existing one. While a configuration is active, set_dimension
    and suppress_feature change only that configuration."""
    doc = sw.active_doc()
    if add:
        if doc.AddConfiguration3(add, "", "", 0) is None:
            raise RuntimeError(f"Could not add configuration {add!r}")
    if activate and not doc.ShowConfiguration2(activate):
        raise LookupError(f"No configuration {activate!r}")
    return {
        "configurations": list(sw.m0(doc, "GetConfigurationNames") or ()),
        "active": doc.ConfigurationManager.ActiveConfiguration.Name,
    }


# --- Checking -----------------------------------------------------------------


@tool
def measure(
    point1: list[float],
    point2: list[float] | None = None,
    entity1: Literal["face", "edge", "vertex"] = "face",
    entity2: Literal["face", "edge", "vertex"] = "face",
) -> dict:
    """Measure one entity (length/area/diameter) or the distance and angle
    between two, picked by model points (mm). Distances in mm, angles in
    degrees; values SOLIDWORKS does not report for this pair are omitted.
    Between holes/arcs, center_to_center_mm and minimum_mm (wall to wall)
    are both given."""
    doc = sw.active_doc()
    doc.ClearSelection2(True)
    sw.select_at(doc, entity1, point1, False)
    if point2:
        sw.select_at(doc, entity2, point2, True)
    m = sw.m0(doc.Extension, "CreateMeasure")
    if not sw.invoke(m, "Calculate", sw.nothing()):  # Nothing = use the selection
        raise RuntimeError("SOLIDWORKS could not measure this selection")
    out = {}
    for key, attr, scale in (
        ("distance_mm", "Distance", 1 / MM), ("normal_distance_mm", "NormalDistance", 1 / MM),
        ("length_mm", "Length", 1 / MM), ("diameter_mm", "Diameter", 1 / MM),
        ("radius_mm", "Radius", 1 / MM), ("area_mm2", "Area", 1 / MM**2),
        ("perimeter_mm", "Perimeter", 1 / MM), ("angle_deg", "Angle", 180 / math.pi),
    ):
        try:
            v = getattr(m, attr)
        except Exception:
            continue
        if v is not None and v >= 0:  # -1 = not applicable
            out[key] = round(v * scale, 4)
    if point2:
        # Between arcs/cylinders the distance depends on IMeasure.ArcOption
        # (0 centre, 1 minimum); report both under unambiguous names.
        by_option = {}
        for option in (0, 1):
            try:
                m.ArcOption = option
                if sw.invoke(m, "Calculate", sw.nothing()) and m.Distance >= 0:
                    by_option[option] = round(m.Distance / MM, 4)
            except Exception:
                continue  # this option is unsupported for the selection; report the other
        if len(set(by_option.values())) == 2:
            centre, minimum = by_option[0], by_option[1]
            out["center_to_center_mm"], out["minimum_mm"] = max(centre, minimum), min(centre, minimum)
    doc.ClearSelection2(True)
    return out


@tool
def check_interference() -> list[dict]:
    """Interferences between components of the active assembly (touching
    faces are not counted), with the overlap volume."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    idm = asm.InterferenceDetectionManager
    idm.TreatCoincidenceAsInterference = False
    out = []
    try:
        for intf in sw.m0(idm, "GetInterferences") or ():
            out.append({
                "components": [c.Name2 for c in (intf.Components or ())],
                "volume_mm3": round(intf.Volume / MM**3, 3),
            })
    finally:
        sw.m0(idm, "Done")
    return out


@tool
def bom() -> list[dict]:
    """Bill of materials of the active assembly: each part/subassembly file
    and configuration with its quantity (suppressed components excluded)."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    counts: dict[tuple[str, str], int] = {}
    for c in asm.GetComponents(False) or ():
        if sw.m0(c, "IsSuppressed"):
            continue
        key = (sw.m0(c, "GetPathName"), c.ReferencedConfiguration)
        counts[key] = counts.get(key, 0) + 1
    return [
        {"file": Path(path).name, "configuration": cfg, "quantity": n, "path": path}
        for (path, cfg), n in sorted(counts.items(), key=lambda kv: Path(kv[0][0]).name.lower())
    ]


# --- Appearance and batch work ------------------------------------------------


@tool
def set_color(rgb: list[int]) -> str:
    """Colour the whole active part, e.g. [200, 30, 30]."""
    doc = sw.active_doc(DOC_PART)
    r, g, b = (max(0, min(255, int(v))) / 255 for v in rgb)
    # red, green, blue, ambient, diffuse, specular, shininess, transparency, emission
    doc.MaterialPropertyValues = sw.doubles((r, g, b, 1, 1, 0.5, 0.4, 0, 0))
    sw.m0(doc, "GraphicsRedraw2")
    return f"coloured {sw.m0(doc, 'GetTitle')} {list(rgb)}"


@tool(timeout=1800)
def export_folder(folder: str, extension: str = "step", recursive: bool = False, overwrite: bool = False) -> list[dict]:
    """Export every .sldprt/.sldasm in a folder to another format (step,
    igs, x_t, stl, 3mf, pdf ...), next to the originals. Files already open
    stay open; others are opened silently and closed again."""
    app = sw.app()
    previous = app.ActiveDoc
    open_docs = {os.path.normcase(sw.m0(d, "GetPathName")): d for d in (sw.m0(app, "GetDocuments") or ()) if sw.m0(d, "GetPathName")}
    pattern = "**/*" if recursive else "*"
    files = [p for p in Path(folder).glob(pattern) if p.suffix.lower() in (".sldprt", ".sldasm") and not p.name.startswith("~$")]
    if not files:
        raise LookupError(f"No .sldprt/.sldasm files in {folder}")
    results = []
    for src in sorted(files):
        target = src.with_suffix("." + extension.lstrip("."))
        if target.exists() and not overwrite:
            results.append({"file": src.name, "skipped": f"{target.name} exists"})
            continue
        doc = open_docs.get(os.path.normcase(str(src)))
        opened_here = doc is None
        try:
            if opened_here:
                errors, warnings = sw.out_int(), sw.out_int()
                doc = app.OpenDoc6(str(src), sw.DOC_TYPES[src.suffix.lower()], 1, "", errors, warnings)
                if doc is None:
                    results.append({"file": src.name, "error": f"open failed: {sw.load_error(errors.value)}"})
                    continue
            errors, warnings = sw.out_int(), sw.out_int()
            before = sw.mtime(str(target))
            ok = doc.Extension.SaveAs3(str(target), 0, 1 | 2, sw.nothing(), sw.nothing(), errors, warnings)
            if ok and sw.file_written(str(target), before):
                results.append({"file": src.name, "exported": target.name, "bytes": target.stat().st_size})
            else:
                results.append({"file": src.name, "error": _save_error(errors.value)})
        finally:
            if opened_here and doc is not None:
                app.CloseDoc(sw.m0(doc, "GetTitle"))
    if previous is not None:
        app.ActivateDoc3(sw.m0(previous, "GetTitle"), False, 0, sw.out_int())
    return results


def _dependencies(path: str, traverse: bool = True, search: bool = True) -> list[str]:
    """Files a SOLIDWORKS document references (paths), all levels with
    traverse. search=True resolves them the way opening would (documents
    already open win); False gives the paths stored in the file."""
    pairs = sw.app().GetDocumentDependencies2(path, traverse, search, False) or ()
    return [pairs[i + 1] for i in range(0, len(pairs) - 1, 2)]


@tool(timeout=600)
def pack_and_go(folder: str, prefix: str = "", suffix: str = "", include_drawings: bool = True) -> dict:
    """Copy the active part/assembly (saved) with every file it references
    into `folder`, optionally renamed with a prefix/suffix, plus the
    drawings next to those files that show them: to hand a design to
    someone else. The copies reference each other only, never the
    originals; the open documents keep pointing at the originals. Unsaved
    changes are not included (save first)."""
    doc = sw.active_doc()
    if sw.m0(doc, "GetType") == DOC_DRAWING:
        raise ValueError("Activate the part or assembly; its drawings come along with include_drawings")
    root = sw.m0(doc, "GetPathName")
    if not root:
        raise ValueError("Save the document first")
    target = os.path.abspath(folder)
    files = [root] + [f for f in _dependencies(root) if f.lower() != root.lower()]
    files = list(dict.fromkeys(os.path.normpath(f) for f in files if os.path.exists(f)))
    if include_drawings:
        wanted = {f.lower() for f in files}
        for d in sorted({os.path.dirname(f) for f in files}):
            for drw in glob.glob(os.path.join(d, "*.slddrw")):
                if any(f.lower() in wanted for f in _dependencies(drw, traverse=False)):
                    files.append(os.path.normpath(drw))
    mapping: dict[str, str] = {}
    for f in files:
        stem, ext = os.path.splitext(os.path.basename(f))
        new = os.path.join(target, f"{prefix}{stem}{suffix}{ext}")
        if any(v.lower() == new.lower() for v in mapping.values()):
            raise ValueError(f"Two referenced files would both become {os.path.basename(new)}; use a prefix or rename one")
        mapping[f] = new
    clash = [v for v in mapping.values() if os.path.exists(v)]
    if clash:
        raise FileExistsError(f"{len(clash)} files already exist in {target} (e.g. {os.path.basename(clash[0])}); choose an empty folder")
    os.makedirs(target, exist_ok=True)
    for old, new in mapping.items():
        shutil.copy2(old, new)
    # point every copied assembly/drawing at the copies (works on closed files)
    app = sw.app()
    lower = {k.lower(): v for k, v in mapping.items()}
    repointed, problems = 0, []
    for old, new in mapping.items():
        if not old.lower().endswith((".sldasm", ".slddrw")):
            continue
        for ref in _dependencies(new, traverse=False, search=False):
            dest = lower.get(os.path.normpath(ref).lower())
            if dest and os.path.normcase(dest) != os.path.normcase(ref):
                if app.ReplaceReferencedDocument(new, ref, dest):
                    repointed += 1
                else:
                    problems.append(f"{os.path.basename(new)} -> {os.path.basename(ref)}")
    # (reading the copies' references back is no check: while the originals are
    # open, SOLIDWORKS reports the open documents' paths even for stored ones)
    result = {"folder": target, "files": [os.path.basename(v) for v in mapping.values()], "references_updated": repointed}
    if sw.m0(doc, "GetSaveFlag"):
        result["warning"] = "the active document has unsaved changes; the copy holds the saved version"
    if problems:
        result["problems"] = [f"could not repoint {p}" for p in problems]
    return result


# --- Holes (ISO metric) -------------------------------------------------------

# ISO 273 medium clearance, ISO coarse tap drills, ISO 4762 counterbores
# (diameter, depth) and 90° countersinks for ISO 10642, in mm.
CLEARANCE = {"M2": 2.4, "M2.5": 2.9, "M3": 3.4, "M4": 4.5, "M5": 5.5, "M6": 6.6, "M8": 9.0,
             "M10": 11.0, "M12": 13.5, "M16": 17.5, "M20": 22.0}
TAP_DRILL = {"M2": 1.6, "M2.5": 2.05, "M3": 2.5, "M4": 3.3, "M5": 4.2, "M6": 5.0, "M8": 6.8,
             "M10": 8.5, "M12": 10.2, "M16": 14.0, "M20": 17.5}
COUNTERBORE = {"M2": (4.4, 2.3), "M2.5": (5.5, 2.8), "M3": (6.5, 3.3), "M4": (8.0, 4.4),
               "M5": (10.0, 5.4), "M6": (11.0, 6.4), "M8": (15.0, 8.6), "M10": (18.0, 10.6),
               "M12": (20.0, 12.6), "M16": (26.0, 16.6), "M20": (33.0, 20.6)}
COUNTERSINK = {"M2": 4.4, "M2.5": 5.5, "M3": 6.72, "M4": 8.96, "M5": 11.2, "M6": 13.44, "M8": 17.92,
               "M10": 22.4, "M12": 26.88, "M16": 33.6, "M20": 40.32}
HOLE_NAMES = {"clearance": "관통 구멍", "tapped": "탭 구멍", "counterbore": "카운터보어", "countersink": "카운터싱크"}


def _circles(positions: list[list[float]], diameter: float) -> list[dict]:
    return [{"type": "circle", "cx": x, "cy": y, "r": diameter / 2} for x, y in positions]


def _rename(feat_name: str, new: str) -> str:
    try:
        sw.feature(sw.active_doc(), feat_name).Name = new
        return new
    except Exception:
        return feat_name  # name taken or not allowed: keep SOLIDWORKS' name


HALF_PI = math.pi / 2
# HoleWizard5 GenericHoleType / FastenerTypeIndex with StandardIndex 8 (ISO),
# from macros recorded in SOLIDWORKS 2025 (see docs/recordings/HANDOFF.md).
WIZARD_TYPES = {"counterbore": (0, 139), "countersink": (1, 140), "clearance": (2, 144), "tapped": (4, 147)}


def _wizard_args(kind: str, size: str, through: bool, depth: float, cosmetic: bool, thread_depth: float) -> tuple:
    """The 27 HoleWizard5 arguments, following the recorded calls and
    scaling the size-dependent values from the ISO tables."""
    generic, fastener = WIZARD_TYPES[kind]
    nominal = mm(float(size[1:]))
    end_type = 1 if through else 0  # 1 through all, 0 blind
    blind = mm(depth) if not through else 0.01  # ignored when through
    tail = ("", False, True, True, True, True, False)  # thread class, reverse, scopes, propagate
    if kind == "counterbore":
        cb_d, cb_h = COUNTERBORE[size]
        values = (mm(cb_d), mm(cb_h), 0, 1, 0, mm(cb_d) * 11.05 / 11, HALF_PI, 0, 0, 0, 0, 0)
        return (generic, 8, fastener, size, end_type, mm(CLEARANCE[size]), blind, -1) + values + tail
    if kind == "countersink":
        values = (mm(COUNTERSINK[size]), HALF_PI, 0, 1, 0, 0, 0, 0, 0, -1, -1, -1)
        return (generic, 8, fastener, size, end_type, mm(CLEARANCE[size]), blind, -1) + values + tail
    if kind == "clearance":
        # V1 = fit (1 normal); V3/V4 = the near-side countersink SOLIDWORKS adds by default
        values = (1, 0, mm(CLEARANCE[size]) * 6.65 / 6.6, HALF_PI, 0, 0, 0, -1, -1, -1, -1, -1)
        return (generic, 8, fastener, size, end_type, mm(CLEARANCE[size]), blind, -1) + values + tail
    near = nominal * 4.05 / 4  # near-side countersink, as recorded for M4
    if through:  # V1 thread depth (unused when the thread runs through), V8 = 1 thread through
        values = (nominal * 9.625 / 4, near, HALF_PI, 0, 0, 0, 1 if cosmetic else 0, 1, 0, 0, -1, -1)
    else:  # V6 = 118° drill point, V8 = 0 blind thread
        values = (mm(thread_depth), near, HALF_PI, 0, 0, math.radians(118), 1 if cosmetic else 0, 0, 0, 0, -1, -1)
    return (generic, 8, fastener, size, end_type, mm(TAP_DRILL[size]), blind, True) + values + tail


def _no_countersink_variants(kind: str, size: str, recorded: tuple) -> list[tuple[str, tuple]]:
    """Argument sets that may switch off the near-side countersink SOLIDWORKS
    adds by default to clearance, tapped and counterbore holes. Which one
    works is not documented, so each is tried and checked on the geometry
    (no conical face at the hole entry). Keys are HoleWizard5 argument
    indexes; the values start at index 8 (V1)."""
    if kind == "clearance":  # V1 fit, V3/V4 countersink diameter/angle
        drill = mm(CLEARANCE[size])
        changes = [("countersink diameter 0", {10: 0}), ("countersink flag off", {8: 0, 10: 0}),
                   ("countersink = drill", {10: drill})]
    elif kind == "tapped":  # V2/V3 countersink diameter/angle
        changes = [("countersink diameter 0", {9: 0}), ("countersink diameter and angle 0", {9: 0, 10: 0}),
                   ("countersink = tap drill", {9: mm(TAP_DRILL[size])})]
    elif kind == "counterbore":  # V4 flag?, V6/V7 countersink diameter/angle on the bore
        changes = [("countersink diameter 0", {13: 0}), ("countersink flag off", {11: 0, 13: 0}),
                   ("countersink = bore", {13: mm(COUNTERBORE[size][0])})]
    else:
        return []
    return [(label, tuple(new.get(i, v) for i, v in enumerate(recorded))) for label, new in changes]


def _has_entry_cone(feat: Any, entry: tuple, normal: tuple) -> bool:
    """Whether the hole has a conical face reaching the entry face (a
    near-side countersink); a blind hole's drill point lies deeper down."""
    level = sum(entry[i] * normal[i] for i in range(3))
    for face in sw.m0(feat, "GetFaces") or ():
        try:
            if not sw.invoke(sw.m0(face, "GetSurface"), "IsCone"):
                continue
            box = sw.m0(face, "GetBox")  # x1, y1, z1, x2, y2, z2 (m)
        except Exception:
            continue  # a face whose surface type cannot be read is not a countersink we can see
        corners = [(box[0 + 3 * a], box[1 + 3 * b], box[2 + 3 * c]) for a in (0, 1) for b in (0, 1) for c in (0, 1)]
        top = max(sum(pt[i] * normal[i] for i in range(3)) for pt in corners)
        if top > level - 0.05 * MM:
            return True
    return False


def _delete_feature_obj(doc: Any, feat: Any) -> None:
    doc.ClearSelection2(True)
    feat.Select2(False, 0)
    doc.Extension.DeleteSelection2(1)


def _select_face_by_ray(doc: Any, face_point: list[float]) -> tuple:
    """Select the face at face_point the way the Hole Wizard recording does:
    a ray shot onto it from just outside, which also sets the pick point.
    Returns the face's outward normal."""
    dist, face = sw.nearest(doc, "face", face_point)
    if face is None or dist > sw.TOLERANCE:
        raise LookupError(f"No face within 0.5 mm of {face_point} mm")
    n = tuple(face.Normal)
    p = tuple(mm(v) for v in face_point)
    start = tuple(p[i] + n[i] * MM for i in range(3))
    doc.ClearSelection2(True)
    if not doc.Extension.SelectByRay(*start, -n[0], -n[1], -n[2], 0.0002, 2, False, 0, 0):  # 2 = faces
        raise RuntimeError(f"Could not pick the face at {face_point} mm")
    return n


def _hole_wizard(doc: Any, kind: str, size: str, positions: list[list[float]], face_point: list[float],
                 depth: float, cosmetic: bool, thread_depth: float, near_csk: bool = True) -> dict:
    recorded = _wizard_args(kind, size, depth <= 0, depth, cosmetic, thread_depth)
    attempts = [("as recorded", recorded)]
    switch_off = not near_csk and kind != "countersink"
    if switch_off:
        attempts = _no_countersink_variants(kind, size, recorded) + attempts
    entry = tuple(mm(v) for v in face_point)
    feat, variant = None, None
    for label, args in attempts:
        normal = _select_face_by_ray(doc, face_point)
        try:
            candidate = doc.FeatureManager.HoleWizard5(*args)
        except Exception:
            continue
        if candidate is None:
            continue
        if switch_off and label != "as recorded" and _has_entry_cone(candidate, entry, normal):
            _delete_feature_obj(doc, candidate)  # this variant kept the countersink
            continue
        feat, variant = candidate, label
        break
    if feat is None:
        raise RuntimeError("HoleWizard5 returned no feature")
    try:
        # Move the holes: replace the point the wizard placed with ours.
        placement = sw.m0(feat, "GetFirstSubFeature")
        doc.ClearSelection2(True)
        placement.Select2(False, 0)
        sw.m0(doc, "EditSketch")
        sm = doc.SketchManager
        if sm.ActiveSketch is None:
            raise RuntimeError("could not edit the hole position sketch")
        points = sw.m0(sw.m0(placement, "GetSpecificFeature2"), "GetSketchPoints2") or ()
        doc.ClearSelection2(True)
        for i, pt in enumerate(points):
            sw._select_entity(doc, pt, i > 0, 0)
        if points:
            sw.m0(doc, "EditDelete")
        sm.AddToDB = True
        try:
            for x, y in positions:
                sm.CreatePoint(mm(x), mm(y), 0)
        finally:
            sm.AddToDB = False
            sm.InsertSketch(True)
        _rebuild(doc)
        placed = len(sw.m0(sw.m0(placement, "GetSpecificFeature2"), "GetSketchPoints2") or ())
        if placed != len(positions):
            raise RuntimeError(f"the wizard hole has {placed} positions instead of {len(positions)}")
    except Exception:
        sw.close_open_sketch(doc)
        doc.ClearSelection2(True)
        feat.Select2(False, 0)
        doc.Extension.DeleteSelection2(1)  # don't leave a half-placed wizard hole behind
        raise
    result = {"features": [feat.Name], "method": "hole_wizard"}
    if kind != "countersink":
        result["near_side_countersink"] = near_csk or variant == "as recorded"
        if not near_csk:
            result["countersink_off_by"] = variant if variant != "as recorded" else "none worked; SOLIDWORKS default kept"
    return result


@tool
def hole(
    size: str,
    positions: list[list[float]],
    kind: Literal["clearance", "tapped", "counterbore", "countersink"] = "clearance",
    face_point: list[float] | None = None,
    plane: str = "top",
    depth: float = 0.0,
    cosmetic_thread: bool = True,
    thread_depth: float = 0.0,
    wizard: bool = True,
    near_side_countersink: bool = False,
) -> dict:
    """ISO metric holes (M2-M20) at several positions at once, on the face at
    face_point ([x,y,z] mm, e.g. the top face). positions are 2D sketch
    coordinates like sketch() (on a top face, sketch y = -model z). depth 0
    = through all. kind: clearance (ISO 273 normal), tapped, counterbore
    (ISO 4762 socket head) or countersink (ISO 10642).

    By default this is a real Hole Wizard hole, so drawings get hole
    callouts; tapped holes get a cosmetic thread (cosmetic_thread=false for
    the tap drill only) and, when blind, a thread depth (default 75 % of the
    depth). Clearance, tapped and counterbore holes come without
    SOLIDWORKS' default near-side countersink unless
    near_side_countersink=true. If the wizard fails, or
    without face_point, or with
    wizard=false, the holes are sketched and cut instead (no callouts);
    "method" in the result says which happened."""
    size = size.upper().replace(" ", "")
    if size not in CLEARANCE:
        raise ValueError(f"Unknown size {size!r}; supported: {sorted(CLEARANCE, key=lambda k: float(k[1:]))}")
    if kind == "tapped" and depth > 0:
        thread_depth = thread_depth or round(depth * 0.75, 3)
        if thread_depth > depth:
            raise ValueError("thread_depth cannot exceed depth")
    wizard_error = None
    if wizard and face_point:
        doc = sw.active_doc(DOC_PART)
        try:
            info = _hole_wizard(doc, kind, size, positions, face_point, depth, cosmetic_thread, thread_depth,
                                near_side_countersink)
            info.update(count=len(positions), drill_mm=TAP_DRILL[size] if kind == "tapped" else CLEARANCE[size])
            if kind == "tapped":
                info["thread"] = {"cosmetic": cosmetic_thread, "depth_mm": thread_depth if depth > 0 else "through"}
            return info
        except Exception as exc:
            wizard_error = f"{type(exc).__name__}: {exc}"
    info = _sketched_holes(size, positions, kind, face_point, plane, depth)
    info["method"] = "sketch_cut"
    if wizard_error:
        info["wizard_error"] = wizard_error
    elif wizard and not face_point:
        info["note"] = "Hole Wizard needs face_point; holes were sketched and cut"
    return info


def _sketched_holes(size: str, positions: list[list[float]], kind: str, face_point, plane: str, depth: float) -> dict:
    where = {"face_point": face_point} if face_point else {"plane": plane}
    sketch_fn, extrude_fn = sketch.__wrapped__, extrude.__wrapped__
    d = TAP_DRILL[size] if kind == "tapped" else CLEARANCE[size]
    # Sketch everything while the face is still intact: after the first cut
    # face_point may sit over a hole and pick a different face.
    seat = None
    if kind == "counterbore":
        seat = sketch_fn(_circles(positions, COUNTERBORE[size][0]), **where)["sketch"]
    elif kind == "countersink":
        seat = sketch_fn(_circles(positions, COUNTERSINK[size]), **where)["sketch"]
    drill = sketch_fn(_circles(positions, d), **where)["sketch"]
    made = []
    if kind == "counterbore":
        made.append(_rename(extrude_fn(depth=COUNTERBORE[size][1], sketch_name=seat, cut=True)["feature"], f"{size} 카운터보어 자리"))
    elif kind == "countersink":
        # a 45° inward draft from the countersink diameter down to the hole
        sink = round((COUNTERSINK[size] - d) / 2, 4)
        made.append(_rename(extrude_fn(depth=sink, sketch_name=seat, cut=True, draft_angle=45)["feature"], f"{size} 카운터싱크 자리"))
    end = "blind" if depth > 0 else "through_all"
    hole_feat = extrude_fn(depth=depth or 10.0, sketch_name=drill, cut=True, end=end)["feature"]
    made.append(_rename(hole_feat, f"{size} {HOLE_NAMES[kind]}"))
    info = {"features": made, "count": len(positions), "drill_mm": d}
    if kind == "counterbore":
        info["counterbore_mm"] = {"diameter": COUNTERBORE[size][0], "depth": COUNTERBORE[size][1]}
    if kind == "countersink":
        info["countersink_mm"] = COUNTERSINK[size]
    return info


@tool
def engrave(
    text: str,
    face_point: list[float],
    x: float = 0.0,
    y: float = 0.0,
    height: float = 5.0,
    depth: float = 0.5,
    emboss: bool = False,
) -> dict:
    """Engrave (cut) or emboss (raise) text on the planar face at face_point
    (mm). x, y = lower-left corner of the text in sketch coordinates like
    sketch(); height = character height, depth = cut/raise depth (mm)."""
    sk = sketch.__wrapped__([{"type": "text", "x": x, "y": y, "text": text, "height": height}],
                            face_point=face_point, dimension=False)
    feat = extrude.__wrapped__(depth=depth, sketch_name=sk["sketch"], cut=not emboss)
    result = {"sketch": sk["sketch"], **feat}
    doc = sw.active_doc(DOC_PART)
    try:
        doc.ClearSelection2(True)
        sw.select_feature(doc, sk["sketch"])
        sw.m0(doc, "BlankSketch")  # hide the text sketch so only the lettering shows
    except Exception as exc:
        result["note"] = f"lettering made, but the sketch could not be hidden: {exc}"
    doc.ClearSelection2(True)
    return result


# --- Sweeps, lofts, ribs and 3D sketches -------------------------------------


def _open_3d_sketch(doc: Any) -> Any:
    sm = doc.SketchManager
    sw.close_open_sketch(doc)
    doc.ClearSelection2(True)
    for opener in (lambda: sm.Insert3DSketch(True), lambda: doc.Insert3DSketch2(True)):
        try:
            opener()
        except Exception:
            continue
        sk = sm.ActiveSketch
        if sk is not None:
            return sk
    raise RuntimeError("SOLIDWORKS did not open a 3D sketch (Insert3DSketch left no active sketch)")


def _create_3d(sm: Any, e: dict) -> int:
    """Draw one 3D entity; returns how many segments were created."""
    t = e.get("type")
    if t == "line":
        pairs = [(e["p1"], e["p2"])]
    elif t == "polyline":
        pts = e["points"]
        pairs = list(zip(pts, pts[1:]))
    elif t == "spline":
        return int(sm.CreateSpline2(sw.doubles([mm(v) for p in e["points"] for v in p]), True) is not None)
    else:
        raise ValueError(f"Unknown 3D sketch entity type {t!r}")
    return sum(sm.CreateLine(*(mm(v) for v in a), *(mm(v) for v in b)) is not None for a, b in pairs)


@tool
def sketch3d(entities: list[dict]) -> dict:
    """3D sketch in model coordinates (mm), e.g. a sweep path. Entities:
      {"type":"polyline","points":[[x,y,z],...]}  (connected lines)
      {"type":"line","p1":[x,y,z],"p2":[x,y,z]}
      {"type":"spline","points":[[x,y,z],...]}
    An empty result is deleted again, and failures say why."""
    doc = sw.active_doc(DOC_PART)
    sm = doc.SketchManager
    _open_3d_sketch(doc)
    made, problems = 0, []
    try:
        for i, e in enumerate(entities):
            try:
                n = 0
                for add_to_db in (True, False):  # retry without AddToDB if nothing was drawn
                    sm.AddToDB = add_to_db
                    n = _create_3d(sm, e)
                    if n:
                        break
                if not n:
                    problems.append({"entity": i, "error": "SOLIDWORKS created nothing", "sketch_open": sm.ActiveSketch is not None})
                made += n
            except Exception as exc:
                problems.append({"entity": i, "error": f"{type(exc).__name__}: {exc}", "sketch_open": sm.ActiveSketch is not None})
    finally:
        sm.AddToDB = False
        if sm.ActiveSketch is not None:
            sm.Insert3DSketch(True)
    feat = sw.last_feature_of_type(doc, "3DProfileFeature")
    if not made:
        doc.ClearSelection2(True)
        feat.Select2(False, 0)
        doc.Extension.DeleteSelection2(0)
        raise RuntimeError(f"No 3D sketch entities could be created; removed the empty sketch. Details: {problems}")
    result = {"sketch": feat.Name, "segments": made}
    if problems:
        result["problems"] = problems
    return result


def _last_sketch_name(doc: Any) -> str:
    for i in range(200):
        f = doc.FeatureByPositionReverse(i)
        if f is None:
            break
        if sw.m0(f, "GetTypeName2") in ("ProfileFeature", "3DProfileFeature"):
            return f.Name
    raise LookupError("No sketch found")


@tool
def sweep(
    path: str,
    profile: str = "",
    circle_diameter: float = 0.0,
    wall: float = 0.0,
    cut: bool = False,
    diameter_is: Literal["outer", "inner"] = "outer",
) -> dict:
    """Sweep a closed profile sketch along a path sketch (2D or 3D, by name).
    Without a profile, sweep a circle of circle_diameter mm along the path
    (pipes, wires, grooves). wall > 0 makes it a hollow tube of that wall
    thickness; circle_diameter is then the outer diameter (wall grows inward)
    unless diameter_is="inner". cut=true removes material instead."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    circular = not profile
    if circular and circle_diameter <= 0:
        raise ValueError("Give a profile sketch or a circle_diameter")
    if profile:
        sw.select_feature(doc, profile, mark=1)
    sw.select_feature(doc, path, append=True, mark=4)
    thin, t1 = wall > 0, mm(wall)
    dia = mm(circle_diameter)
    fm = doc.FeatureManager
    # Propagate, Alignment, TwistCtrlType, KeepTangency, AdvancedSmoothing, StartMatching,
    # EndMatching, IsThin, Thickness1, Thickness2, ThinType, PathAlign, ...
    # ThinType 0 grows outward from the profile (seen in testing), 1 the opposite way
    thin_type = 1 if (thin and diameter_is == "outer") else 0
    head = (False, False, 0, False, False, 0, 0, thin, t1, t1, thin_type, 0)
    if cut:
        tail = (True, True, 0, True)  # UseFeatScope, UseAutoSelect, TwistAngle, MergeSmoothFaces
        feat = _call_first(fm, [
            ("InsertCutSwept5", head + tail + (False, False, False, circular, dia, 0)),
            ("InsertCutSwept4", head + tail + (False, False, False)),
        ])
    else:
        tail = (True, True, True, 0, True)  # Merge, UseFeatScope, UseAutoSelect, TwistAngle, MergeSmoothFaces
        feat = _call_first(fm, [
            ("InsertProtrusionSwept4", head + tail + (circular, dia, 0)),
            ("InsertProtrusionSwept3", head + tail),
        ])
    result = _feature_result(feat, "sweep cut" if cut else "sweep")
    if circular:
        if thin:
            outer = circle_diameter if diameter_is == "outer" else circle_diameter + 2 * wall
            result["tube_mm"] = {"outer_diameter": outer, "inner_diameter": round(outer - 2 * wall, 4), "wall": wall}
        else:
            result["diameter_mm"] = circle_diameter
    return result


@tool
def loft(profiles: list[str], guides: list[str] | None = None, cut: bool = False, closed: bool = False) -> dict:
    """Loft through profile sketches (names, in order, on different planes,
    e.g. made with offset_plane), optionally steered by guide curve sketches.
    cut=true removes material."""
    if len(profiles) < 2:
        raise ValueError("A loft needs at least two profiles")
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    _select_features(doc, profiles, 1, append=False)
    for g in guides or []:
        sw.select_feature(doc, g, append=True, mark=2)
    fm = doc.FeatureManager
    if cut:
        base = (closed, True, False, 1, 0, 0, False, 0, 0, 0, True, True)
        feat = _call_first(fm, [("InsertCutBlend", base + (0,)), ("InsertCutBlend", base)])
    else:
        base = (closed, True, False, 1, 0, 0, 1, 1, True, True, False, 0, 0, 0, True, True, True)
        feat = _call_first(fm, [("InsertProtrusionBlend2", base + (0,)), ("InsertProtrusionBlend", base)])
    return _feature_result(feat, "loft cut" if cut else "loft")


@tool
def helix(pitch: float, revolutions: float, diameter: float = 0.0, plane: str = "top",
          center: list[float] | None = None, sketch_name: str = "", clockwise: bool = False,
          reverse: bool = False, start_angle: float = 0.0, taper_angle: float = 0.0) -> dict:
    """Helix curve (mm, degrees) for springs, threads and coils: around a
    circle of `diameter` drawn on `plane` at center [x,y] (sketch
    coordinates), or around the circle in an existing sketch_name. It rises
    pitch per turn along the plane's normal (reverse = the other way).
    Sweep a circle along it for a spring: sweep(path=<helix>,
    circle_diameter=wire)."""
    doc = sw.active_doc(DOC_PART)
    if not sketch_name:
        if diameter <= 0:
            raise ValueError("Give the helix diameter, or a sketch_name with one circle")
        cx, cy = center or [0, 0]
        sketch_name = sketch.__wrapped__([{"type": "circle", "cx": cx, "cy": cy, "r": diameter / 2}], plane=plane,
                                         dimension=False)["sketch"]
    doc.ClearSelection2(True)
    sw.select_feature(doc, sketch_name)
    last = doc.FeatureByPositionReverse(0).Name
    defined_by = sw.constant("swHelixDefinedByPitchAndRevolution")
    doc.InsertHelix(reverse, clockwise, taper_angle != 0, False, defined_by, mm(pitch * revolutions),
                    mm(pitch), revolutions, deg(taper_angle), deg(start_angle))
    newest = doc.FeatureByPositionReverse(0)  # InsertHelix returns nothing
    if newest.Name == last or sw.m0(newest, "GetTypeName2") != "Helix":
        raise RuntimeError(f"SOLIDWORKS made no helix from {sketch_name}; it needs a sketch with one circle")
    result = _feature_result(newest, "helix")
    result.update({"height_mm": round(pitch * revolutions, 4), "length_mm": round(
        revolutions * math.hypot(math.pi * (diameter or _circle_diameter(doc, sketch_name)), pitch), 3)})
    return result


def _circle_diameter(doc: Any, sketch_name: str) -> float:
    sk = sw.m0(sw.feature(doc, sketch_name), "GetSpecificFeature2")
    for seg in sw.m0(sk, "GetSketchSegments") or ():
        if sw.m0(seg, "GetType") == sw.constant("swSketchARC"):
            return 2 * sw.m0(seg, "GetRadius") / MM
    return 0.0


@tool
def rib(thickness: float, sketch_name: str = "", both_sides: bool = True, flip_material: bool = False) -> dict:
    """Rib from an open sketch (e.g. one line between two walls), default
    the most recent sketch; thickness in mm."""
    doc = sw.active_doc(DOC_PART)
    sketch = sketch_name or _last_sketch_name(doc)

    def attempt(flip: bool) -> Any:
        doc.ClearSelection2(True)
        sw.select_feature(doc, sketch)
        # two-sided, reverse thickness, thickness, reference edge, reverse material,
        # drafted, draft outward, draft angle, normal to sketch(, drafted from wall)
        args = (both_sides, False, mm(thickness), 0, flip, False, False, 0.0, False)
        last = doc.FeatureByPositionReverse(0).Name
        made = _call_first(doc.FeatureManager, [("InsertRib", args + (False,)), ("InsertRib", args)])
        if made is None:  # InsertRib builds the rib but returns nothing: look at the tree
            newest = doc.FeatureByPositionReverse(0)
            made = newest if newest.Name != last and sw.m0(newest, "GetTypeName2") == "Rib" else None
        return made

    before = sw.m0(doc, "GetMassProperties")
    feat = attempt(flip_material)
    after = sw.m0(doc, "GetMassProperties")
    if feat is None or (before and after and after[3] <= before[3] + 1e-12):
        if feat is not None:  # material went the wrong way (into nothing): try the other side
            _delete_feature_obj(doc, feat)
        feat = attempt(not flip_material)
    result = _feature_result(feat, "rib")
    after = sw.m0(doc, "GetMassProperties")
    if before and after:
        result["volume_added_mm3"] = round((after[3] - before[3]) / MM**3, 2)
    return result


# --- Sheet metal --------------------------------------------------------------


@tool
def sheet_metal_base_flange(
    thickness: float,
    bend_radius: float,
    length: float = 0.0,
    sketch_name: str = "",
    reverse: bool = False,
    mid_plane: bool = False,
    k_factor: float = 0.5,
) -> dict:
    """Start a sheet metal part from a sketch (default the most recent): a
    closed sketch becomes a flat plate; an open profile (an L or U of lines)
    is extruded `length` mm in one direction (reverse=true flips it,
    mid_plane=true centres it on the sketch plane), with bends of
    bend_radius and the given K-factor. Reports the length actually built."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    profile = sketch_name or _last_sketch_name(doc)
    sw.select_feature(doc, profile)
    t, r, l = mm(thickness), mm(bend_radius), mm(length or 10)
    method, error = "feature_definition", None
    try:
        # As recorded in SOLIDWORKS 2025 (docs/recordings/HANDOFF.md).
        fm = doc.FeatureManager
        allowance = sw.m0(fm, "CreateCustomBendAllowance")
        allowance.KFactor = k_factor
        data = fm.CreateDefinition(sw.constant("swFmBaseFlange"))
        sw.invoke(data, "Initialize", False, True, allowance, True, 1, True, 0.5, 0.0001, 0.0001)
        data.Thickness = t
        data.BendRadius = r
        data.ReverseThickness = False
        data.ReverseDirection = reverse
        data.D1EndConditionType = 1  # blind, as recorded
        data.D1ReverseOffset = False
        data.D2EndConditionType = 1
        data.D2ReverseOffset = False
        if mid_plane:
            data.OffsetDirections = 2  # both directions, half each
            data.D1EndConditionDistance = l / 2
            data.D2EndConditionDistance = l / 2
        else:
            data.OffsetDirections = 1  # one direction
            data.D1EndConditionDistance = l
            data.D2EndConditionDistance = 0.01
        feat = fm.CreateFeature(data)
        if feat is None:
            raise RuntimeError("CreateFeature returned no feature")
    except Exception as exc:
        # Older path: these arguments extrude both ways by the distance
        # each, so pass half per side (always centred).
        error = f"{type(exc).__name__}: {exc}"
        method = "InsertSheetMetalBaseFlange2 (centred)"
        doc.ClearSelection2(True)
        sw.select_feature(doc, profile)
        h = l / 2
        feat = _call_first(doc.FeatureManager, [
            ("InsertSheetMetalBaseFlange2", (t, False, r, h, h, reverse, 0, 0, 0, sw.nothing(), False, 0.5, 0.0001, 0.0001, 0.5, True, False, True, True)),
            ("InsertSheetMetalBaseFlange2", (t, False, r, h, h, reverse, 0, 0, 0, sw.nothing(), False, 0.5, 0.0001, 0.0001, 0.5, True, False, True)),
        ])
    result = _feature_result(feat, "base flange")
    result["method"] = method
    if error:
        result["definition_error"] = error
    axis = _sketch_normal_axis(sw.feature(doc, profile))
    box = _body_box(doc)
    if axis is not None and box:
        lo, hi = box[axis] / MM, box[3 + axis] / MM
        result["length_built_mm"] = round(hi - lo, 3)
        result["extent_mm"] = [round(lo, 3), round(hi, 3)]
        if length and abs((hi - lo) - length) > 0.01:
            result["warning"] = f"asked for {length} mm along {'xyz'[axis]}, got {round(hi - lo, 3)} mm"
    return result


@tool
def flatten(flat: bool = True) -> str:
    """Show a sheet metal part flat (true) or folded (false), in the active
    configuration."""
    doc = sw.active_doc(DOC_PART)
    flat_pattern = next((f for f in sw.iter_features(doc) if sw.m0(f, "GetTypeName2") == "FlatPattern"), None)
    if flat_pattern is None:
        raise LookupError("No flat pattern: this is not a sheet metal part")
    if not _call_first(flat_pattern, [("SetSuppression2", (1 if flat else 0, 1, sw.nothing()))]):  # 1 = this configuration
        raise RuntimeError("Could not change the flat pattern state")
    _rebuild(doc)
    return "flattened" if flat else "folded"


@tool
def export_flat_pattern(path: str, bend_lines: bool = True) -> dict:
    """Export the flat pattern of the active (saved) sheet metal part to
    DXF/DWG for laser cutting, with bend lines by default."""
    doc = sw.active_doc(DOC_PART)
    model = sw.m0(doc, "GetPathName")
    if not model:
        raise RuntimeError("Save the part first; the export reads the saved model")
    target = str(Path(path))
    options = 1 | (4 if bend_lines else 0)  # geometry, bend lines
    before = sw.mtime(target)
    ok = doc.ExportToDWG2(target, model, 1, True, sw.nothing(), False, False, options, sw.nothing())  # 1 = sheet metal
    if not ok or not sw.file_written(target, before):
        raise RuntimeError(f"Flat pattern export failed for {target}")
    return {"path": target, "bytes": os.path.getsize(target)}


def _transform_points(xform: Any, points: list[tuple]) -> list[tuple]:
    mu = sw.m0(sw.app(), "GetMathUtility")
    out = []
    for p in points:
        pt = sw.invoke(mu, "CreatePoint", sw.doubles(p))
        out.append(tuple(sw.m0(sw.invoke(pt, "MultiplyTransform", xform), "ArrayData")))
    return out


@tool
def edge_flange(edge_point: list[float], length: float, angle: float = 90.0, flip: bool = False,
                bend_radius: float = 0.0) -> dict:
    """Sheet metal edge flange along the straight edge through edge_point
    ([x,y,z] mm), `length` mm long measured from the face the edge lies on,
    bent `angle` degrees (90 = perpendicular) towards that face's outside
    (flip=true for the other side). bend_radius 0 = the part's default."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_at(doc, "EDGE", edge_point, False, 0)
    edge = doc.SelectionManager.GetSelectedObject6(1, -1)
    a = sw.m0(sw.m0(edge, "GetStartVertex"), "GetPoint")
    b = sw.m0(sw.m0(edge, "GetEndVertex"), "GetPoint")
    # the broad face the edge lies on and the thin side face
    faces = sorted(sw.m0(edge, "GetTwoAdjacentFaces2") or (), key=lambda f: sw.m0(f, "GetArea"), reverse=True)
    if len(faces) < 2:
        raise ValueError(f"The edge at {edge_point} mm is not a sheet edge")
    n, t = tuple(faces[0].Normal), tuple(faces[1].Normal)
    sign = -1 if flip else 1
    th = deg(angle)
    d = tuple(mm(length) * (sign * n[i] * math.sin(th) + t[i] * math.cos(th)) for i in range(3))
    sketch_feat = doc.InsertSketchForEdgeFlange(edge, th, flip)
    if sketch_feat is None:
        raise RuntimeError("Could not create the flange profile sketch")
    # profile: the edge plus `length` in the flange direction, in the sketch's coordinates
    sk = sw.m0(sketch_feat, "GetSpecificFeature2")
    corners = _transform_points(sw.m0(sk, "ModelToSketchTransform"),
                                [a, b, tuple(b[i] + d[i] for i in range(3)), tuple(a[i] + d[i] for i in range(3))])
    doc.ClearSelection2(True)
    sketch_feat.Select2(False, 0)
    sw.m0(doc, "EditSketch")
    sm = doc.SketchManager
    sm.AddToDB = True
    try:
        for i in range(4):
            p, q = corners[i], corners[(i + 1) % 4]
            sm.CreateLine(p[0], p[1], 0, q[0], q[1], 0)
    finally:
        sm.AddToDB = False
        sm.InsertSketch(True)
    import pythoncom
    from win32com.client import VARIANT

    def objects(items):
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, items)

    options = 128 | (1 if bend_radius <= 0 else 0)  # default relief, default radius
    before = sw.m0(doc, "GetMassProperties")
    feat = doc.FeatureManager.InsertSheetMetalEdgeFlange2(
        objects([edge]), objects([sketch_feat]), options, th, mm(bend_radius), 1, 0, 0, 0.5, 0, 0, 0, sw.nothing())
    if feat is None:
        _delete_feature_obj(doc, sketch_feat)
        raise RuntimeError("SOLIDWORKS did not create the edge flange; check the edge and the length")
    after = sw.m0(doc, "GetMassProperties")
    result = _feature_result(feat, "edge flange")
    if before and after and after[3] <= before[3] + 1e-12:
        result["warning"] = "the part volume did not grow; the flange may be empty"
    return result


HEM_TYPES = {"open": 0, "closed": 1, "teardrop": 2, "rolled": 3, "double": 4}  # swHemType_e


@tool
def hem(edge_points: list[list[float]], kind: Literal["closed", "open", "teardrop", "rolled", "double"] = "closed",
        length: float = 10.0, gap: float = 0.5, angle: float = 270.0, radius: float = 1.0,
        outside: bool = False, reverse: bool = False) -> dict:
    """Sheet metal hem (folded-back edge) on the edges through edge_points
    ([[x,y,z], ...] mm). closed/open use length (mm) and, for open, gap;
    teardrop and rolled use angle (deg) and radius (mm). outside puts the hem
    outside the material line; reverse folds it the other way."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    for i, p in enumerate(edge_points):
        sw.select_at(doc, "EDGE", p, i > 0, 0)
    feat = doc.FeatureManager.InsertSheetMetalHem2(
        HEM_TYPES[kind], 1 if outside else 0, reverse, mm(length), mm(gap), deg(angle), mm(radius), 0,
        sw.nothing(), True, 0, 0, True, 0.5, 0, 0)  # default relief
    return _feature_result(feat, "hem")


def _chain_end(edges: list) -> tuple[Any, Any]:
    """(edge, vertex) at one free end of a chain of edges, where a miter
    flange profile starts."""
    def key(v):
        return tuple(round(c, 7) for c in sw.m0(v, "GetPoint"))

    count: dict = {}
    for e in edges:
        for v in (sw.m0(e, "GetStartVertex"), sw.m0(e, "GetEndVertex")):
            count[key(v)] = count.get(key(v), 0) + 1
    for e in edges:
        for v in (sw.m0(e, "GetStartVertex"), sw.m0(e, "GetEndVertex")):
            if count[key(v)] == 1:
                return e, v
    return edges[0], sw.m0(edges[0], "GetStartVertex")  # a closed loop: any vertex


@tool
def miter_flange(edge_points: list[list[float]], length: float, angle: float = 90.0, flip: bool = False) -> dict:
    """Sheet metal miter flange: one flange running along a chain of
    connected edges (through edge_points, [[x,y,z], ...] mm), mitred at the
    corners. `length` mm, bent `angle` degrees (90 = perpendicular) towards
    the outside of the face the edges lie on (flip for the other side)."""
    doc = sw.active_doc(DOC_PART)
    sel = doc.SelectionManager
    edges = []
    for p in edge_points:
        doc.ClearSelection2(True)
        sw.select_at(doc, "EDGE", p, False, 0)
        edges.append(sel.GetSelectedObject6(1, -1))
    edge, vertex = _chain_end(edges)
    faces = sorted(sw.m0(edge, "GetTwoAdjacentFaces2") or (), key=lambda f: sw.m0(f, "GetArea"), reverse=True)
    if len(faces) < 2:
        raise ValueError("The edges are not sheet edges")
    n, t = tuple(faces[0].Normal), tuple(faces[1].Normal)
    sign, th = (-1 if flip else 1), deg(angle)
    d = tuple(mm(length) * (sign * n[i] * math.sin(th) + t[i] * math.cos(th)) for i in range(3))
    a = sw.m0(vertex, "GetPoint")

    def pick(obj, mark, append):
        data = sw.m0(sel, "CreateSelectData")
        data.Mark = mark
        obj.Select4(append, data)

    # profile plane: normal to the end edge at the free vertex
    doc.ClearSelection2(True)
    pick(edge, 0, False)
    pick(vertex, 1, True)
    plane = doc.FeatureManager.InsertRefPlane(2, 0, 4, 0, 0, 0)  # perpendicular, coincident
    if plane is None:
        raise RuntimeError("Could not create the profile plane at the end of the edges")
    doc.ClearSelection2(True)
    _select_plane(doc, plane)
    sm = doc.SketchManager
    sm.InsertSketch(True)
    if sm.ActiveSketch is None:
        raise RuntimeError("Could not open the profile sketch")
    p0, p1 = _transform_points(sw.m0(sm.ActiveSketch, "ModelToSketchTransform"), [a, tuple(a[i] + d[i] for i in range(3))])
    sm.AddToDB = True
    try:
        sm.CreateLine(p0[0], p0[1], 0, p1[0], p1[1], 0)
    finally:
        sm.AddToDB = False
        sm.InsertSketch(True)
    profile = sw.last_feature_of_type(doc, "ProfileFeature")
    # the profile sketch with mark 0, the edges with mark 1
    doc.ClearSelection2(True)
    profile.Select2(False, 0)
    for e in edges:
        pick(e, 1, True)
    feat = doc.FeatureManager.InsertSheetMetalMiterFlange(
        True, 0, mm(0.1), True, True, 0.5, 0, 0, 0, False, 1, 0, 0, sw.nothing())
    doc.ClearSelection2(True)
    if feat is None:
        _delete_feature_obj(doc, profile)
        _delete_feature_obj(doc, plane)
        raise RuntimeError("SOLIDWORKS did not create the miter flange; are the edges connected?")
    plane.Select2(False, 0)
    sw.m0(doc, "BlankRefGeom")  # hide the helper plane
    doc.ClearSelection2(True)
    return _feature_result(feat, "miter flange") | {"profile_plane": plane.Name}


# --- Weldments ----------------------------------------------------------------


def _profile_root() -> Path:
    try:
        folder = sw.app().GetUserPreferenceStringValue(29)  # swFileLocationsWeldmentProfiles
    except Exception:
        folder = ""
    if folder and os.path.isdir(folder):
        return Path(folder)
    return Path(os.environ.get("ProgramData", r"C:\ProgramData")) / "SOLIDWORKS" / "SOLIDWORKS 2025" / "Weldment Profiles"


@tool
def weldment_profiles(standard: str = "iso") -> dict:
    """Structural profiles available for structural_member: profile types
    per standard (iso, din, ansi, jis, gb, ...) and the sizes of each."""
    root = _profile_root()
    folder = next((d for d in root.iterdir() if d.is_dir() and d.name.lower() == standard.lower()), None) if root.exists() else None
    if folder is None:
        raise LookupError(f"No standard {standard!r} in {root}; have: {sorted(d.name for d in root.iterdir() if d.is_dir()) if root.exists() else []}")
    app = sw.app()
    out = {}
    for f in sorted(folder.glob("*.sldlfp")):
        out[f.stem] = list(sw.invoke(app, "GetConfigurationNames", str(f)) or ())
    return {"standard": folder.name, "profiles": out}


@tool
def structural_member(
    profile: str,
    size: str,
    standard: str = "iso",
    sketch_name: str = "",
    segment_points: list[list[float]] | None = None,
    corner: Literal["miter", "butt1", "butt2", "none"] = "miter",
    rotate: float = 0.0,
) -> dict:
    """Weldment structural member (e.g. profile "square tube", size
    "40 x 40 x 4") along the lines of a 2D or 3D sketch (default the most
    recent). segment_points ([[x,y,z], ...] model mm, one point on each
    line) limits it to some lines; by default every line of the sketch is
    used. corner sets the corner treatment where members meet; rotate turns
    the profile about the path (deg). The part becomes a weldment; see
    weldment_profiles for what is installed."""
    import pythoncom
    from win32com.client import VARIANT

    doc = sw.active_doc(DOC_PART)
    root = _profile_root()
    folder = next((d for d in root.iterdir() if d.is_dir() and d.name.lower() == standard.lower()), None)
    path = folder / f"{profile}.sldlfp" if folder else None
    if path is None or not path.exists():
        raise LookupError(f"No profile {profile!r} for {standard!r}; see weldment_profiles")
    sizes = list(sw.invoke(sw.app(), "GetConfigurationNames", str(path)) or ())
    if size not in sizes:
        raise LookupError(f"No size {size!r} for {profile}; sizes: {sizes}")
    if sketch_name:
        feat = sw.feature(doc, sketch_name)
    else:
        feats = [f for f in sw.iter_features(doc) if sw.m0(f, "GetTypeName2") in ("ProfileFeature", "3DProfileFeature")]
        if not feats:
            raise LookupError("No sketch to place the member on; draw the path with sketch or sketch3d first")
        feat = feats[-1]
    sk = sw.m0(feat, "GetSpecificFeature2")
    segments = [s for s in sw.m0(sk, "GetSketchSegments") or () if not sw.m0(s, "ConstructionGeometry")]
    if segment_points:
        chosen = []
        for p in segment_points:
            q = tuple(mm(v) for v in p)
            seg = min(segments, key=lambda s: _segment_distance(s, sk, q))
            if _segment_distance(seg, sk, q) > sw.TOLERANCE:
                raise LookupError(f"No line of {feat.Name} within 0.5 mm of {p} mm")
            chosen.append(seg)
        segments = chosen
    if not segments:
        raise ValueError(f"{feat.Name} has no lines")
    fm = doc.FeatureManager
    if not any(sw.m0(f, "GetTypeName2") == "WeldmentFeature" for f in sw.iter_features(doc)):
        sw.m0(fm, "InsertWeldmentFeature")  # the part becomes a weldment
    group = sw.m0(fm, "CreateStructuralMemberGroup")
    group.Segments = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, segments)
    group.ApplyCornerTreatment = corner != "none"
    group.CornerTreatmentType = {"none": 0, "miter": 1, "butt1": 2, "butt2": 3}[corner]  # swSolidworksWeldmentEndCondOptions
    group.Angle = deg(rotate)
    before = len(doc.GetBodies2(0, True) or ())
    member = fm.InsertStructuralWeldment5(str(path), 1, True, VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [group]), size)
    result = _feature_result(member, "structural member")
    result.update(profile=f"{standard}/{profile} {size}", segments=len(segments),
                  bodies_added=len(doc.GetBodies2(0, True) or ()) - before)
    return result


def _segment_distance(seg: Any, sk: Any, point: tuple) -> float:
    """Distance (m) from a model point to a sketch line."""
    a = sw.m0(seg, "GetStartPoint2")
    b = sw.m0(seg, "GetEndPoint2")
    to_model = sw.m0(sw.m0(sk, "ModelToSketchTransform"), "Inverse")
    pa, pb = _transform_points(to_model, [(a.X, a.Y, a.Z), (b.X, b.Y, b.Z)])
    ab = [pb[i] - pa[i] for i in range(3)]
    ap = [point[i] - pa[i] for i in range(3)]
    t = max(0.0, min(1.0, sum(ab[i] * ap[i] for i in range(3)) / (sum(c * c for c in ab) or 1)))
    return math.dist(point, tuple(pa[i] + t * ab[i] for i in range(3)))


@tool
def trim_members(trim_points: list[list[float]], boundary_points: list[list[float]]) -> dict:
    """Trim weldment members where they run into others: the bodies through
    trim_points ([[x,y,z], ...] mm, a point on each member's surface) are
    cut back to the boundary bodies through boundary_points, keeping the
    longer side (e.g. legs that poke into a frame)."""
    doc = sw.active_doc(DOC_PART)
    trim = [_body_at(doc, p) for p in trim_points]
    bounds = [_body_at(doc, p) for p in boundary_points]
    before = {b.Name: b.GetMassProperties(1.0)[3] for b in trim}
    doc.ClearSelection2(True)
    # members to trim with mark 1, the boundary with mark 2; end condition 4 trims to bodies
    for i, b in enumerate(trim):
        _pick_body(doc, b, 1, i > 0)
    for b in bounds:
        _pick_body(doc, b, 2, True)
    feat = doc.FeatureManager.InsertWeldmentTrimFeature(4)
    doc.ClearSelection2(True)
    result = _feature_result(feat, "trim/extend")
    _rebuild(doc)
    after = sum(b.GetMassProperties(1.0)[3] for b in doc.GetBodies2(0, True) or () if sw.m0(feat, "Name") in b.Name)
    removed = sum(before.values()) - after if after else 0
    result["volume_removed_mm3"] = round(removed / MM**3, 1)
    if removed <= 0:
        result["warning"] = "nothing was trimmed; do the members actually overlap the boundary?"
    return result


def _pick_body(doc: Any, body: Any, mark: int, append: bool) -> bool:
    data = sw.m0(doc.SelectionManager, "CreateSelectData")
    data.Mark = mark
    return bool(body.Select2(append, data))


@tool
def end_cap(face_points: list[list[float]], thickness: float = 3.0, inward: bool = False, chamfer: float = 0.0) -> dict:
    """Close the open ends of weldment tubes with end caps `thickness` mm
    thick. face_points are points on the end faces (the tube wall ring,
    [[x,y,z], ...] mm). inward sinks the cap into the tube; chamfer bevels
    its corners (mm)."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    for i, p in enumerate(face_points):
        sw.select_at(doc, "FACE", p, i > 0, 0)
    before = len(doc.GetBodies2(0, True) or ())
    feat = doc.FeatureManager.InsertEndCapFeature3(
        mm(thickness), True, chamfer > 0, 0.0, 0.5, mm(chamfer), False, 0.0, False, inward)
    doc.ClearSelection2(True)
    result = _feature_result(feat, "end cap")
    result["caps"] = len(doc.GetBodies2(0, True) or ()) - before
    return result


@tool
def gusset(face1_point: list[float], face2_point: list[float], size1: float = 60.0, size2: float = 60.0,
           thickness: float = 5.0, shape: Literal["triangle", "polygon"] = "triangle", chamfer: float = 15.0,
           location: Literal["center", "start", "end"] = "center",
           thickness_side: Literal["both", "inner", "outer"] = "both") -> dict:
    """Gusset plate in the corner between two planar faces of a weldment
    (points on each, [x,y,z] mm), size1 x size2 mm along them, `thickness`
    mm thick. polygon cuts the outer corner off by `chamfer` mm on both
    legs."""
    import pythoncom
    from win32com.client import VARIANT

    doc = sw.active_doc(DOC_PART)
    faces = []
    for p in (face1_point, face2_point):
        dist, face = sw.nearest(doc, "face", p)
        if face is None or dist > sw.TOLERANCE:
            raise LookupError(f"No face within 0.5 mm of {p} mm")
        faces.append(face)
    poly = shape == "polygon"
    c = mm(chamfer) if poly else 0.0
    side = {"inner": 0, "both": 1, "outer": 2}[thickness_side]  # swGussetThickness*
    loc = {"start": 0, "center": 1, "end": 2}[location]  # swGussetProfileLocation*
    before = len(doc.GetBodies2(0, True) or ())
    doc.ClearSelection2(True)
    # BIsProfile: True = polygon (needs both chamfer legs), False = triangle
    feat = doc.FeatureManager.InsertGussetFeature2(
        mm(thickness), side, loc, poly, mm(size1), mm(size2), c, math.pi / 4, c, False, 0.0, 0, False, False, False,
        VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, faces))
    result = _feature_result(feat, "gusset")
    result["plates"] = len(doc.GetBodies2(0, True) or ()) - before
    return result


def _resolved(cpm: Any, name: str) -> str:
    import pythoncom
    from win32com.client import VARIANT

    raw, value = sw.out_bstr(), sw.out_bstr()
    was = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, False)
    cpm.Get5(name, False, raw, value, was)
    return value.value


@tool
def cut_list() -> list[dict]:
    """The weldment cut list of the active part: one entry per group of
    identical members (description, quantity, length, end angles, total
    length) plus other bodies such as gussets and end caps, refreshed
    first."""
    doc = sw.active_doc(DOC_PART)
    out = []
    for f in sw.iter_features(doc):
        if sw.m0(f, "GetTypeName2") != "SolidBodyFolder":
            continue
        sw.m0(sw.m0(f, "GetSpecificFeature2"), "UpdateCutList")
        sub = sw.m0(f, "GetFirstSubFeature")
        while sub is not None:
            if sw.m0(sub, "GetTypeName2") == "CutListFolder":
                folder = sw.m0(sub, "GetSpecificFeature2")
                cpm = sub.CustomPropertyManager
                props = {n: _resolved(cpm, n) for n in sw.m0(cpm, "GetNames") or ()}
                out.append({"item": sub.Name, "bodies": sw.m0(folder, "GetBodyCount"), "properties": props})
            sub = sw.m0(sub, "GetNextSubFeature")
    if not out:
        raise LookupError("No cut list: is this a weldment (structural_member)?")
    return out


def _line_sketch_on_face(doc: Any, face: Any, p1: tuple, p2: tuple) -> Any:
    """A sketch on `face` holding one line between two model points (m)."""
    sw.close_open_sketch(doc)
    doc.ClearSelection2(True)
    face.Select4(False, sw.nothing())
    sm = doc.SketchManager
    sm.InsertSketch(True)
    if sm.ActiveSketch is None:
        raise RuntimeError("Could not open a sketch on the face")
    a, b = _transform_points(sw.m0(sm.ActiveSketch, "ModelToSketchTransform"), [p1, p2])
    sm.AddToDB = True
    try:
        sm.CreateLine(a[0], a[1], 0, b[0], b[1], 0)
    finally:
        sm.AddToDB = False
        sm.InsertSketch(True)
    return sw.last_feature_of_type(doc, "ProfileFeature")


def _sketched_bend(doc: Any, sketch: Any, angle: float, reverse: bool, radius: float) -> Any:
    data = doc.FeatureManager.CreateDefinition(sw.constant("swFmSketchBend"))
    data.BendAngle = deg(angle)
    data.ReverseDirection = reverse
    if radius > 0:
        data.UseDefaultBendRadius = False
        data.BendRadius = mm(radius)
    doc.ClearSelection2(True)
    sketch.Select2(False, 0)
    feat = doc.FeatureManager.CreateFeature(data)
    doc.ClearSelection2(True)
    return feat


def _face_at(doc: Any, point_m: tuple) -> Any:
    dist, face = sw.nearest(doc, "face", [c / MM for c in point_m])
    if face is None or dist > sw.TOLERANCE:
        raise LookupError(f"No face within 0.5 mm of {[round(c / MM, 3) for c in point_m]} mm")
    return face


@tool
def sketched_bend(line_start: list[float], line_end: list[float], angle: float = 90.0, reverse: bool = False,
                  bend_radius: float = 0.0) -> dict:
    """Sheet metal bend along a line drawn on a flat face, from line_start to
    line_end ([x,y,z] mm, both on the face; the line may overrun the
    face). The side containing the part's origin (the larger side) stays;
    the other folds up by `angle` degrees, or down with reverse=true.
    bend_radius 0 = the part's default. Works on the flat plate; a second
    bend on an already bent flange is refused (the fixed face cannot be set
    through the API here), so build steps as separate flanges instead."""
    doc = sw.active_doc(DOC_PART)
    p1, p2 = tuple(mm(v) for v in line_start), tuple(mm(v) for v in line_end)
    mid = tuple((p1[i] + p2[i]) / 2 for i in range(3))
    sketch = _line_sketch_on_face(doc, _face_at(doc, mid), p1, p2)
    before = _body_box(doc)
    feat = _sketched_bend(doc, sketch, angle, reverse, bend_radius)
    if feat is None:
        _delete_feature_obj(doc, sketch)
        raise RuntimeError("SOLIDWORKS did not bend along that line")
    result = _feature_result(feat, "sketched bend")
    result["box_mm"] = {"before": [round(v / MM, 2) for v in before], "after": [round(v / MM, 2) for v in _body_box(doc)]}
    return result


# --- Assembly management ----------------------------------------------------

MATE_TYPES = {0: "coincident", 1: "concentric", 2: "perpendicular", 3: "parallel", 4: "tangent", 5: "distance",
              6: "angle", 7: "unknown", 8: "symmetric", 9: "cam", 10: "gear", 11: "width", 12: "lock_to_sketch",
              13: "rack_pinion", 14: "max_mates", 15: "path", 16: "lock", 17: "screw", 18: "linear_coupler",
              19: "universal_joint", 20: "coordinate", 21: "slot", 22: "hinge", 23: "slider", 24: "profile_center",
              25: "magnetic"}


@tool
def list_mates() -> list[dict]:
    """Mates of the active assembly with their type, the components they
    connect and their value (distance mm / angle deg) where they have one."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    out = []
    for f in sw.iter_features(asm):
        if sw.m0(f, "GetTypeName2") != "MateGroup":
            continue
        sub = sw.m0(f, "GetFirstSubFeature")
        while sub is not None:
            mate = sw.m0(sub, "GetSpecificFeature2")
            info: dict = {"name": sub.Name, "suppressed": bool(sw.m0(sub, "IsSuppressed"))}
            try:
                kind = mate.Type
                info["type"] = MATE_TYPES.get(kind, kind)
                comps = []
                for i in range(sw.m0(mate, "GetMateEntityCount")):
                    comp = sw.m0(mate.MateEntity(i), "ReferenceComponent")
                    comps.append(comp.Name2 if comp is not None else None)
                info["components"] = comps
                if kind in (5, 6):
                    value = sw.m0(mate, "DisplayDimension").GetDimension2(0).Value
                    info["distance_mm" if kind == 5 else "angle_deg"] = round(value, 4)
                    high, low = mate.MaximumVariation, mate.MinimumVariation  # limits, relative to the value
                    if high != low:  # a limit mate
                        scale = 1 / MM if kind == 5 else 180 / math.pi
                        info["limits"] = [round(value + low * scale, 4) + 0.0, round(value + high * scale, 4) + 0.0]
            except Exception as exc:
                info.setdefault("problem", f"{type(exc).__name__}: {exc}")
            out.append(info)
            sub = sw.m0(sub, "GetNextSubFeature")
    return out


def _component(asm: Any, name: str) -> Any:
    comps = list(asm.GetComponents(False) or ())
    for c in comps:
        if c.Name2 == name or c.Name2.split("/")[-1] == name:
            return c
    starts = [c for c in comps if c.Name2.split("/")[-1].rsplit("-", 1)[0] == name]
    if len(starts) == 1:
        return starts[0]
    raise LookupError(f"No component {name!r}; components: {[c.Name2 for c in comps]}")


@tool
def set_component(
    name: str,
    fixed: bool | None = None,
    visible: bool | None = None,
    suppressed: bool | None = None,
    position: list[float] | None = None,
    rotation: list[float] | None = None,
) -> dict:
    """Change a component of the active assembly (name as in
    list_components, e.g. "bracket-1"): fix or float it, hide or show it,
    suppress or resolve it, move its origin to position [x,y,z] mm, or turn
    it to rotation [about x, about y, about z] degrees from the assembly
    axes (applied in that order). Moving only works where mates do not
    hold it."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    comp = _component(asm, name)
    if suppressed is not None:
        comp.SetSuppression2(0 if suppressed else 2)  # swComponentSuppressed / swComponentResolved
        comp = _component(asm, comp.Name2)
    if fixed is not None:
        asm.ClearSelection2(True)
        comp.Select4(False, sw.nothing(), False)
        sw.m0(asm, "FixComponent" if fixed else "UnfixComponent")
    if visible is not None:
        comp.Visible = 1 if visible else 0
    if position is not None or rotation is not None:
        _place_component(asm, comp, position, rotation)
    asm.ClearSelection2(True)
    sw.m0(asm, "GraphicsRedraw2")
    result = {
        "component": comp.Name2,
        "fixed": bool(sw.m0(comp, "IsFixed")),
        "visible": bool(comp.Visible),
        "suppressed": bool(sw.m0(comp, "IsSuppressed")),
    }
    if not result["suppressed"]:
        result.update(_axes_info(list(sw.m0(comp.Transform2, "ArrayData"))))
    return result


def _pick(doc: Any, obj: Any, mark: int, append: bool) -> bool:
    data = sw.m0(doc.SelectionManager, "CreateSelectData")
    data.Mark = mark
    try:
        return bool(obj.Select4(append, data, False))  # components take a third argument
    except Exception:
        return bool(obj.Select4(append, data))


def _select_assembly_direction(asm: Any, direction: str, mark: int) -> None:
    """A direction in an assembly: "x"/"y"/"z" (a reference axis through
    the origin is created), an axis name, or "edge:x,y,z" (assembly mm)."""
    if direction.startswith("edge:"):
        sw.select_at(asm, "EDGE", [float(v) for v in direction[5:].split(",")], True, mark)
    elif direction.lower() in ("x", "y", "z"):
        axis = _axis(asm, direction.lower())
        asm.ClearSelection2(True)
        axis.Select2(True, mark)
    else:
        sw.select_feature(asm, direction, append=True, mark=mark)


def _make_component_pattern(asm: Any, comps: list, count: int, spacing: float, direction: str,
                            linear: bool, reverse: bool) -> tuple[Any, list]:
    fm = asm.FeatureManager
    # the definition first, then components (mark 1) and the direction (mark 2)
    data = fm.CreateDefinition(sw.constant("swFmLocalLPattern" if linear else "swFmLocalCirPattern"))
    asm.ClearSelection2(True)
    for i, c in enumerate(comps):
        _pick(asm, c, 1, i > 0)
    _select_assembly_direction(asm, direction, 2)
    if linear:
        data.D1Spacing = mm(spacing)
        data.D1TotalInstances = count
        data.D1ReverseDirection = reverse
    else:
        data.Spacing = deg(spacing)
        data.TotalInstances = count
        data.ReverseDirection = reverse
    before = {c.Name2 for c in asm.GetComponents(False) or ()}
    feat = fm.CreateFeature(data)
    asm.ClearSelection2(True)
    return feat, [c for c in asm.GetComponents(False) or () if c.Name2 not in before]


def _origin(comp: Any) -> list[float]:
    return list(sw.m0(comp.Transform2, "ArrayData"))[9:12]


@tool
def component_pattern(
    components: list[str],
    count: int,
    spacing: float,
    direction: str = "x",
    kind: Literal["linear", "circular"] = "linear",
    reverse: bool = False,
) -> dict:
    """Pattern assembly components (names as in list_components). linear:
    `count` copies (including the original) `spacing` mm apart along
    direction ("x"/"y"/"z" = the positive assembly axis, an axis name or
    "edge:x,y,z" in assembly mm); reverse goes the other way. circular:
    around that axis, `spacing` is the angle in degrees between copies
    (360/count spreads them evenly)."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    comps = [_component(asm, n) for n in components]
    linear = kind == "linear"
    axis = "xyz".find(direction.lower()) if direction.lower() in ("x", "y", "z") else -1
    if axis >= 0:  # creating the axis clears the selection, so before picking
        direction = _axis(asm, direction.lower()).Name
    feat, added = _make_component_pattern(asm, comps, count, spacing, direction, linear, reverse)
    if feat is not None and linear and axis >= 0 and added:
        # a reference axis may point either way: make "x" mean +x
        seed = _origin(comps[0])[axis]
        nearest = min(added, key=lambda c: abs(_origin(c)[axis] - seed))
        if (_origin(nearest)[axis] < seed) != reverse:
            _delete_feature_obj(asm, feat)
            feat, added = _make_component_pattern(asm, comps, count, spacing, direction, linear, not reverse)
    result = _feature_result(feat, f"{kind} component pattern")
    result["components_added"] = len(added)
    return result


AXIS_INDEX = {"x": 0, "y": 1, "z": 2}  # swExplodeDirectionIndex_e


@tool
def exploded_view(steps: list[dict], show: bool = True) -> dict:
    """Exploded view of the active assembly. steps is a list of
    {"components": ["name-1", ...], "direction": "x"|"y"|"z", "distance": mm}
    (negative distance = the other way), applied in order. Adds to the
    configuration's exploded view, creating one if needed; show=false
    collapses it again afterwards. Drawing views can show it with
    drawing_view(exploded=true)."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    cfg = sw.m0(asm, "GetActiveConfiguration")
    if not (sw.m0(asm, "GetExplodedViewNames") or ()):
        sw.m0(asm, "CreateExplodedView")
        # it starts with an automatic step that spreads every component: drop it
        for i in reversed(range(sw.m0(cfg, "GetNumberOfExplodeSteps") or 0)):
            cfg.DeleteExplodeStep(sw.m0(cfg.GetExplodeStep(i), "Name"))
    names = list(sw.m0(asm, "GetExplodedViewNames") or ())
    made = []
    for i, step in enumerate(steps):
        comps = [_component(asm, n) for n in step["components"]]
        asm.ClearSelection2(True)
        for j, c in enumerate(comps):
            _pick(asm, c, 1, j > 0)  # components with mark 1
        dist = float(step["distance"])
        err = sw.out_int()
        added = cfg.AddExplodeStep2(mm(abs(dist)), AXIS_INDEX[step.get("direction", "y").lower()], dist < 0,
                                    0, -1, False, False, False, err)
        if added is None:
            raise RuntimeError(f"Step {i + 1} was not added (swAddExplodeStepError {err.value})")
        made.append(sw.m0(added, "Name") if hasattr(added, "Name") else f"step {i + 1}")
    asm.ClearSelection2(True)
    sw.invoke(asm, "ShowExploded2", show, names[0] if names else "")
    sw.m0(asm, "GraphicsRedraw2")
    return {"exploded_view": names[0] if names else None, "steps_added": made,
            "total_steps": sw.m0(cfg, "GetNumberOfExplodeSteps"), "shown": show}


@tool
def replace_component(name: str, path: str, all_instances: bool = False, configuration: str = "") -> dict:
    """Swap an assembly component (name as in list_components) for another
    part/assembly file, keeping its mates where the new geometry allows.
    all_instances replaces every instance of the same file."""
    asm = sw.active_doc(DOC_ASSEMBLY)
    comp = _component(asm, name)
    target = str(Path(path))
    if not os.path.exists(target):
        raise FileNotFoundError(target)
    old = sw.m0(comp, "GetPathName")
    asm.ClearSelection2(True)
    _pick(asm, comp, 0, False)
    # configuration: by name match unless one is given; re-attach mates
    ok = asm.ReplaceComponents2(target, configuration, all_instances, 1 if configuration else 0, True)
    asm.ClearSelection2(True)
    if not ok:
        raise RuntimeError(f"SOLIDWORKS did not replace {name} with {target}")
    sw.m0(asm, "EditRebuild3")
    now = [c.Name2 for c in asm.GetComponents(True) or () if os.path.normcase(sw.m0(c, "GetPathName") or "") == os.path.normcase(target)]
    return {"replaced": name, "old_file": old, "new_file": target, "components_now": now}


@tool
def mirror_components(components: list[str], plane: str = "right") -> dict:
    """Mirror assembly components (names as in list_components) about a
    plane ("front"/"top"/"right" or a plane name): adds mirrored instances
    of the same files (no opposite-hand parts are made)."""
    import pythoncom
    from win32com.client import VARIANT

    asm = sw.active_doc(DOC_ASSEMBLY)
    comps = [_component(asm, n) for n in components]
    key = plane.lower()
    target = sw.standard_plane(asm, key) if key in sw.PLANE_NAMES else sw.feature(asm, plane)
    before = {c.Name2 for c in asm.GetComponents(True) or ()}
    orient = sw.constant("swMirrorComponentOrientation_None")
    asm.MirrorComponents3(
        target, VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, comps),
        VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [orient] * len(comps)), False,
        sw.empty(), False, sw.empty(), 0, "", "", 0, False, False, False)
    added = [c for c in asm.GetComponents(True) or () if c.Name2 not in before]
    if not added:
        raise RuntimeError("SOLIDWORKS did not mirror the components")
    return {"mirrored": [{"component": c.Name2, "origin_mm": [round(v / MM, 3) for v in _origin(c)]} for c in added]}


# --- Escape hatches -------------------------------------------------------


STEP_REF = re.compile(r"^\$(\d+)((?:\.[^.]+)*)$")


def _step_value(results: list, text: str) -> Any:
    match = STEP_REF.match(text)
    if not match:
        return text
    index = int(match.group(1))
    if index >= len(results):
        raise ValueError(f"{text!r} refers to step {index}, which has not run yet")
    value = results[index]
    for key in match.group(2).split(".")[1:]:
        if isinstance(value, list):
            value = value[int(key)]
        elif isinstance(value, dict) and key in value:
            value = value[key]
        else:
            raise ValueError(f"{text!r}: step {index}'s result has no {key!r}")
    return value


def _resolve_steps(value: Any, results: list) -> Any:
    if isinstance(value, str):
        return _step_value(results, value)
    if isinstance(value, list):
        return [_resolve_steps(v, results) for v in value]
    if isinstance(value, dict):
        return {k: _resolve_steps(v, results) for k, v in value.items()}
    return value


@tool(timeout=900)
def run_steps(steps: list[dict], stop_on_error: bool = True) -> dict:
    """Run several tools in one call, in order, which is much faster than a
    call each: steps = [{"tool": "sketch", "args": {...}}, {"tool":
    "extrude", "args": {"depth": 10}}, ...]. An argument "$N.key" is
    replaced by `key` from step N's result (N counts from 0; keys separated
    by dots, list positions as numbers), e.g. {"features": ["$3.feature"]}
    or "$0.faces.2.point_mm". Returns every step's result; stops at the
    first failing step (reported with its error) unless stop_on_error=false.
    screenshot is left out (call it on its own)."""
    results: list = []
    failed = []
    for i, step in enumerate(steps):
        name = step.get("tool", "")
        try:
            fn = TOOLS.get(name)
            if fn is None or name in ("run_steps", "screenshot"):
                raise ValueError(f"Unknown or unsupported tool {name!r}")
            out = fn(**_resolve_steps(step.get("args") or {}, results))
            results.append(out)
        except Exception as exc:
            error = {"step": i, "tool": name, "error": f"{type(exc).__name__}: {exc}"}
            results.append(error)
            failed.append(error)
            if stop_on_error:
                break
    done = len(results) - len(failed)
    return {"completed": done, "of": len(steps), "results": results, **({"failed": failed} if failed else {})}


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
    import sys

    if any(a in sys.argv[1:] for a in ("--check", "--register", "--unregister")):
        import doctor  # setup check / Claude Desktop registration, see doctor.py

        sys.exit(doctor.main([a for a in sys.argv[1:] if a != "--check"]))
    mcp.run()


if __name__ == "__main__":
    main()
