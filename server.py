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
import re
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


def tool(fn=None, *, timeout: float | None = None):
    """Register fn as a tool. Its body runs on the COM thread, and any failure
    (including COM errors and timeouts) comes back to the agent as a
    readable error. Long-running tools pass a larger timeout (seconds)."""

    def register(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                return call(lambda: fn(*args, **kwargs), timeout)
            except ToolError:
                raise
            except Exception as exc:
                raise ToolError(f"{type(exc).__name__}: {exc}") from exc

        return mcp.tool()(wrapper)

    return register(fn) if fn is not None else register


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
        raise RuntimeError(f"Open failed (swFileLoadError {errors.value}) for {path}")
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
    boxes = [b.GetBodyBox() for b in doc.GetBodies2(0, True) or ()]
    if not boxes:
        return doc.GetPartBox(True)
    return [min(b[i] for b in boxes) for i in range(3)] + [max(b[i] for b in boxes) for i in range(3, 6)]


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
    system = deg(value) if sw.m0(dim, "GetType") == 2 else mm(value)
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
) -> dict:
    """Create a drawing with the standard three views (plus model dimensions
    and hole callouts for Hole Wizard holes) of a saved part/assembly;
    default the active document. Export it as PDF or DXF afterwards with
    export."""
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
               near: tuple[float, float] | None = None, axis: str = "") -> tuple[float, float]:
    """Centre (m) of a free area of the given size on the sheet, clear of the
    views (and their annotations) by a gap; the free spot closest to `near`
    (default the sheet's top right). axis "x" keeps near's y (search along
    x), "y" keeps near's x: for views SOLIDWORKS keeps aligned. Falls back to
    the sheet's top right corner."""
    sheet_w, sheet_h = _sheet_size(drw)
    boxes = [sw.m0(v, "GetOutline") for v in _drawing_views(drw) if v.Name != exclude]
    gap, margin, step = 0.012, 0.01, 0.005
    hw, hh = width / 2 + gap, height / 2 + gap
    target = near or (sheet_w, sheet_h)
    best, best_d = None, math.inf
    ys = [target[1]] if axis == "x" else None
    xs = [target[0]] if axis == "y" else None
    y_vals = ys or [margin + height / 2 + i * step for i in range(int((sheet_h - 2 * margin - height) / step) + 1)]
    x_vals = xs or [margin + width / 2 + i * step for i in range(int((sheet_w - 2 * margin - width) / step) + 1)]
    for y in y_vals:
        for x in x_vals:
            if all(x + hw <= b[0] or x - hw >= b[2] or y + hh <= b[1] or y - hh >= b[3] for b in boxes):
                d = math.hypot(x - target[0], y - target[1])
                if d < best_d:
                    best, best_d = (x, y), d
    return best or (sheet_w - margin - width / 2, sheet_h - margin - height / 2)


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
) -> dict:
    """Add a model view (e.g. an isometric view) to the active drawing.
    position is the view centre on the sheet in mm (default a free spot
    beside the other views); scale like "1:2" (default the sheet scale). The
    model is the one the drawing already shows unless model_path is given."""
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
    (default the first) with its top right corner at position [x,y] mm
    (default the sheet's top right). template is a .sldbomtbt file, default
    the standard template SOLIDWORKS ships."""
    drw = sw.active_doc(DOC_DRAWING)
    target = _drawing_view(drw, view)
    sheet_w, sheet_h = _sheet_size(drw)
    x, y = (mm(position[0]), mm(position[1])) if position else (sheet_w - 0.01, sheet_h - 0.01)
    bom_type = {"top_level": 2, "parts_only": 1, "indented": 3}[kind]  # swBomType_e
    template = template or _bom_template()
    drw.ActivateView(target.Name)
    # anchored at the top right corner (swBOMConfigurationAnchor_TopRight)
    config = target.ReferencedConfiguration  # an empty name gives a table without rows
    table = target.InsertBomTable4(False, x, y, 2, bom_type, config, template, False, 0, False)
    if table is None:
        raise RuntimeError(f"Could not insert a BOM table (template {template}); does the view show an assembly?")
    drw.ForceRebuild3(False)
    try:
        rows = sw.m0(table, "RowCount") - 1  # without the header
    except Exception:
        rows = None
    return {"rows": rows, "configuration": config, "template": template}


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
                    results.append({"file": src.name, "error": f"open failed (swFileLoadError {errors.value})"})
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
def rib(thickness: float, sketch_name: str = "", both_sides: bool = True, flip_material: bool = False) -> dict:
    """Rib from an open sketch (e.g. one line between two walls), default
    the most recent sketch; thickness in mm."""
    doc = sw.active_doc(DOC_PART)
    doc.ClearSelection2(True)
    sw.select_feature(doc, sketch_name or _last_sketch_name(doc))
    feat = _call_first(doc.FeatureManager, [
        ("InsertRib", (both_sides, False, mm(thickness), 0, flip_material, False, True, 0, False)),
    ])
    return _feature_result(feat, "rib")


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
                if kind == 5:
                    info["distance_mm"] = round(sw.m0(mate, "DisplayDimension").GetDimension2(0).Value, 4)
                elif kind == 6:
                    info["angle_deg"] = round(sw.m0(mate, "DisplayDimension").GetDimension2(0).Value, 4)
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
) -> dict:
    """Change a component of the active assembly (name as in
    list_components, e.g. "bracket-1"): fix or float it, hide or show it,
    suppress or resolve it, or move its origin to position [x,y,z] mm (only
    moves freely if mates do not hold it)."""
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
    if position is not None:
        xf = comp.Transform2
        data = list(sw.m0(xf, "ArrayData"))
        data[9:12] = [mm(v) for v in position]
        mu = sw.m0(sw.app(), "GetMathUtility")
        comp.Transform2 = sw.invoke(mu, "CreateTransform", sw.doubles(data))
        sw.m0(asm, "EditRebuild3")
    asm.ClearSelection2(True)
    sw.m0(asm, "GraphicsRedraw2")
    origin = list(sw.m0(comp.Transform2, "ArrayData"))[9:12] if not sw.m0(comp, "IsSuppressed") else None
    return {
        "component": comp.Name2,
        "fixed": bool(sw.m0(comp, "IsFixed")),
        "visible": bool(comp.Visible),
        "suppressed": bool(sw.m0(comp, "IsSuppressed")),
        "origin_mm": [round(v / MM, 3) for v in origin] if origin else None,
    }


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
