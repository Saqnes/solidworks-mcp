"""COM plumbing for the SOLIDWORKS API.

COM objects belong to the thread that created them, so every API call runs
on one dedicated worker thread (`call`). The SOLIDWORKS API works in metres
and radians; the helpers here convert from the millimetres and degrees the
tools expose.
"""

from __future__ import annotations

import math
import os
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


def mm(v: float) -> float:
    return float(v) * MM


def deg(v: float) -> float:
    return math.radians(float(v))


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
            _app.RevisionNumber()
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
    if kind is not None and doc.GetType() != kind:
        raise RuntimeError(
            f"Active document is a {DOC_TYPE_NAMES.get(doc.GetType(), '?')}, "
            f"this needs a {DOC_TYPE_NAMES[kind]}"
        )
    return doc


def doc_info(doc: Any) -> dict:
    return {
        "title": doc.GetTitle(),
        "path": doc.GetPathName() or None,
        "type": DOC_TYPE_NAMES.get(doc.GetType(), "?"),
        "unsaved_changes": bool(doc.GetSaveFlag()),
    }


# --- Features and selection -----------------------------------------------


def iter_features(doc: Any):
    f = doc.FirstFeature()
    while f is not None:
        yield f
        f = f.GetNextFeature()


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
        if f.GetTypeName2() == type_name:
            return f
        i += 1


def standard_planes(doc: Any) -> dict[str, Any]:
    """The default Front/Top/Right planes, found by position rather than name
    because their names are localized (e.g. 정면/윗면/우측면 in Korean)."""
    planes = [f for f in iter_features(doc) if f.GetTypeName2() == "RefPlane"][:3]
    if len(planes) < 3:
        raise RuntimeError("Could not find the three default planes")
    return dict(zip(("front", "top", "right"), planes))


def select_feature(doc: Any, name: str, append: bool = False, mark: int = 0) -> None:
    if not feature(doc, name).Select2(append, mark):
        raise RuntimeError(f"Could not select {name!r}")


def select_at(doc: Any, entity: str, point_mm: list[float], append: bool, mark: int = 0) -> None:
    """Select the face/edge/vertex at a model-space point (mm)."""
    x, y, z = (mm(v) for v in point_mm)
    if not doc.Extension.SelectByID2("", entity.upper(), x, y, z, append, mark, nothing(), 0):
        raise LookupError(f"No {entity.lower()} found at {point_mm} mm")
