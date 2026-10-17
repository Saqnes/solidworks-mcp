"""COM plumbing for the SOLIDWORKS API.

COM objects belong to the thread that created them, so every API call runs
on one dedicated worker thread (`call`). The SOLIDWORKS API works in metres
and radians; the helpers here convert from the millimetres and degrees the
tools expose.
"""

from __future__ import annotations

import math
import os
import inspect
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable, TypeVar

T = TypeVar("T")

MM = 0.001

# swDocumentTypes_e
DOC_PART, DOC_ASSEMBLY, DOC_DRAWING = 1, 2, 3
DOC_TYPES = {".sldprt": DOC_PART, ".sldasm": DOC_ASSEMBLY, ".slddrw": DOC_DRAWING}
DOC_TYPE_NAMES = {0: "none", DOC_PART: "part", DOC_ASSEMBLY: "assembly", DOC_DRAWING: "drawing"}

# swUserPreferenceStringValue_e default templates
TEMPLATE_PREF = {DOC_PART: 8, DOC_ASSEMBLY: 9, DOC_DRAWING: 10}


def _init_com() -> None:
    import pythoncom

    pythoncom.CoInitialize()


_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="sw-com", initializer=_init_com)


def call(fn: Callable[[], T]) -> T:
    """Run fn on the COM thread and return its result (or raise its error)."""
    return _executor.submit(fn).result()


# --- COM argument helpers -------------------------------------------------


def nothing() -> Any:
    """VBA `Nothing` for optional object arguments (e.g. SelectByID2 Callout)."""
    import pythoncom
    from win32com.client import VARIANT

    return VARIANT(pythoncom.VT_DISPATCH, None)


def empty() -> Any:
    """An empty Variant, for optional array arguments."""
    import pythoncom
    from win32com.client import VARIANT

    return VARIANT(pythoncom.VT_EMPTY, None)


def out_int() -> Any:
    """A ByRef Long for the Errors/Warnings out-parameters."""
    import pythoncom
    from win32com.client import VARIANT

    return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)


def out_bstr() -> Any:
    """A ByRef String out-parameter; read .value after the call."""
    import pythoncom
    from win32com.client import VARIANT

    return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")


def doubles(values: Any) -> Any:
    """A SAFEARRAY of doubles, e.g. for MathUtility.CreatePoint."""
    import pythoncom
    from win32com.client import VARIANT

    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(v) for v in values))


def mm(v: float) -> float:
    return float(v) * MM


def deg(v: float) -> float:
    return math.radians(float(v))


def m0(obj: Any, name: str) -> Any:
    """Call a parameterless API method. Under late binding pywin32 often
    resolves these (GetTitle, FirstFeature, ...) to their value instead of a
    method, so call only when it is actually a method."""
    attr = getattr(obj, name)
    return attr() if inspect.isroutine(attr) else attr


# --- Connection -----------------------------------------------------------

_app: Any = None


def app() -> Any:
    """The running SOLIDWORKS instance, starting one if none is running.
    Must be called on the COM thread."""
    global _app
    import pythoncom
    import win32com.client
    from win32com.client import dynamic

    if _app is not None:
        try:
            m0(_app, "RevisionNumber")
            return _app
        except Exception:
            _app = None
    progid = os.environ.get("SW_PROGID", "SldWorks.Application")
    try:
        obj = dynamic.Dispatch(win32com.client.GetActiveObject(progid)._oleobj_)
    except pythoncom.com_error:
        obj = dynamic.Dispatch(progid)
    obj.Visible = True
    _app = obj
    return _app


def active_doc(kind: int | None = None) -> Any:
    doc = app().ActiveDoc
    if doc is None:
        raise RuntimeError("No document is open in SOLIDWORKS")
    if kind is not None and m0(doc, "GetType") != kind:
        raise RuntimeError(
            f"Active document is a {DOC_TYPE_NAMES.get(m0(doc, 'GetType'), '?')}, "
            f"this needs a {DOC_TYPE_NAMES[kind]}"
        )
    return doc


def doc_info(doc: Any) -> dict:
    return {
        "title": m0(doc, "GetTitle"),
        "path": m0(doc, "GetPathName") or None,
        "type": DOC_TYPE_NAMES.get(m0(doc, "GetType"), "?"),
        "unsaved_changes": bool(m0(doc, "GetSaveFlag")),
    }


# --- Features and selection -----------------------------------------------


def iter_features(doc: Any):
    f = m0(doc, "FirstFeature")
    while f is not None:
        yield f
        f = m0(f, "GetNextFeature")


def feature(doc: Any, name: str) -> Any:
    f = doc.FeatureByName(name)
    if f is None:
        raise LookupError(f"No feature named {name!r}. Use list_features to see names.")
    return f


def last_feature_of_type(doc: Any, type_name: str) -> Any:
    i = 0
    while True:
        f = doc.FeatureByPositionReverse(i)
        if f is None:
            raise LookupError(f"Document has no {type_name} feature")
        if m0(f, "GetTypeName2") == type_name:
            return f
        i += 1


PLANE_NAMES = {
    "front": ("Front Plane", "정면", "Front"),
    "top": ("Top Plane", "윗면", "Top"),
    "right": ("Right Plane", "우측면", "Right"),
}
PLANE_AXIS = {"right": 0, "top": 1, "front": 2}  # index of the plane normal (x, y, z)


def _plane_normal_axis(f: Any) -> int | None:
    """Axis (0=x, 1=y, 2=z) a reference plane through the origin faces, or
    None if it is offset or tilted."""
    try:
        data = m0(m0(f, "GetSpecificFeature2").Transform, "ArrayData")
    except Exception:
        return None
    normal, origin = data[6:9], data[9:12]
    if any(abs(v) > 1e-9 for v in origin):
        return None
    axis = max(range(3), key=lambda i: abs(normal[i]))
    return axis if abs(abs(normal[axis]) - 1) < 1e-6 else None


def standard_plane(doc: Any, which: str) -> Any:
    """The default Front/Top/Right plane. Tried by its English or Korean
    name first, then by geometry (the plane through the origin facing the
    right axis), since names depend on the template's language."""
    for name in PLANE_NAMES[which]:
        f = doc.FeatureByName(name)
        if f is not None:
            return f
    for f in iter_features(doc):
        if m0(f, "GetTypeName2") == "RefPlane" and _plane_normal_axis(f) == PLANE_AXIS[which]:
            return f
    raise LookupError(f"Could not find the {which} plane; pass its name from list_features")


def select_feature(doc: Any, name: str, append: bool = False, mark: int = 0) -> None:
    if not feature(doc, name).Select2(append, mark):
        raise RuntimeError(f"Could not select {name!r}")


def _select_entity(doc: Any, ent: Any, append: bool, mark: int) -> bool:
    data = m0(doc.SelectionManager, "CreateSelectData")
    data.Mark = mark
    return bool(ent.Select4(append, data))


def _nearest(entities: Any, point: tuple[float, float, float], where: Callable[[Any], Any]) -> tuple[float, Any]:
    best = (math.inf, None)
    for ent in entities or ():
        try:
            q = where(ent)
        except Exception:
            continue
        d = math.dist(point, tuple(q[:3]))
        if d < best[0]:
            best = (d, ent)
    return best


TOLERANCE = 0.5 * MM


def select_at(doc: Any, entity: str, point_mm: list[float], append: bool, mark: int = 0) -> None:
    """Select the face/edge/vertex nearest a model-space point (mm).

    In parts this searches the solid bodies directly, so it finds hidden
    edges too; SelectByID2 only picks what is visible in the current view.
    Assemblies fall back to SelectByID2."""
    entity = entity.lower()
    p = tuple(mm(v) for v in point_mm)
    if m0(doc, "GetType") == DOC_PART:
        best = (math.inf, None)
        for body in doc.GetBodies2(0, False) or ():  # solid bodies, hidden ones included
            if entity == "edge":
                cand = _nearest(m0(body, "GetEdges"), p, lambda e: e.GetClosestPointOn(*p))
            elif entity == "face":
                cand = _nearest(m0(body, "GetFaces"), p, lambda f: f.GetClosestPointOn(*p))
            elif entity == "vertex":
                cand = _nearest(m0(body, "GetVertices"), p, lambda v: m0(v, "GetPoint"))
            else:
                raise ValueError(f"Unknown entity type {entity!r}")
            best = min(best, cand, key=lambda c: c[0])
        if best[1] is not None and best[0] <= TOLERANCE:
            if not _select_entity(doc, best[1], append, mark):
                raise RuntimeError(f"Found the {entity} at {point_mm} mm but could not select it")
            return
        if best[1] is not None:
            raise LookupError(f"No {entity} within 0.5 mm of {point_mm} mm (nearest is {best[0] / MM:.2f} mm away)")
    x, y, z = p
    if not doc.Extension.SelectByID2("", entity.upper(), x, y, z, append, mark, nothing(), 0):
        raise LookupError(f"No {entity} found at {point_mm} mm")
