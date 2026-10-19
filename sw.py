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
from concurrent.futures import Future, ThreadPoolExecutor
from concurrent.futures import TimeoutError as FutureTimeout
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


DEFAULT_TIMEOUT = float(os.environ.get("SW_TIMEOUT", "90"))
_pending: Future | None = None


def call(fn: Callable[[], T], timeout: float | None = None) -> T:
    """Run fn on the COM thread and return its result (or raise its error).

    There is one COM thread, so a call stuck in SOLIDWORKS (typically behind
    a modal dialog) would block every later tool. Give up after `timeout`
    seconds, and refuse new work while the stuck call is still running."""
    global _pending
    if _pending is not None and not _pending.done():
        raise RuntimeError(
            "SOLIDWORKS is still busy with an earlier request that timed out. "
            "Look for an open dialog in SOLIDWORKS, close it, then retry."
        )
    _pending = _executor.submit(fn)
    limit = timeout or DEFAULT_TIMEOUT
    try:
        return _pending.result(timeout=limit)
    except FutureTimeout:
        raise RuntimeError(
            f"SOLIDWORKS did not answer within {limit:.0f} s. A dialog may be waiting "
            "in SOLIDWORKS; close it. The request may still finish in the background."
        ) from None


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


def invoke(obj: Any, name: str, *args: Any) -> Any:
    """Call an API method whose arguments are all optional. Late binding may
    already have invoked it (with no arguments) and handed back its result,
    e.g. IMeasure.Calculate arriving as a bool."""
    flag = getattr(obj, "_FlagAsMethod", None)  # pywin32: treat `name` as a method from now on
    if flag is not None:
        try:
            flag(name)
        except Exception:
            pass  # not a dynamic dispatch object; inspect what getattr returns instead
    attr = getattr(obj, name)
    return attr(*args) if inspect.isroutine(attr) else attr


def close_open_sketch(doc: Any) -> bool:
    """Leave sketch edit mode if a 2D or 3D sketch is open (a leftover edit
    blocks deleting features and opening new sketches)."""
    sm = doc.SketchManager
    sk = sm.ActiveSketch
    if sk is None:
        return False
    try:
        is3d = bool(invoke(sk, "Is3D"))
    except Exception:
        is3d = False
    (sm.Insert3DSketch if is3d else sm.InsertSketch)(True)
    return True


def file_written(path: str, before: float | None) -> bool:
    """True if `path` exists and was (re)written after mtime `before`
    (None = did not exist), so a stale file is not taken for success."""
    if not os.path.exists(path):
        return False
    return before is None or os.path.getmtime(path) > before


def mtime(path: str) -> float | None:
    return os.path.getmtime(path) if os.path.exists(path) else None


def put_indexed(obj: Any, name: str, index: Any, value: Any) -> None:
    """Assign an indexed property (VBA: obj.Name(index) = value), which
    pywin32's late binding has no syntax for."""
    import pythoncom

    dispid = obj._oleobj_.GetIDsOfNames(name)
    obj._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, False, index, value)


# --- SOLIDWORKS constants (swconst.tlb) ---------------------------------------

SWCONST_TYPELIB = "{4687F359-55D0-4CD3-B6CF-2EB42C11F989}"
_constants: dict[str, int] | None = None


def _swconst_paths() -> list[str]:
    paths = []
    try:
        exe = invoke(app(), "GetExecutablePath")
        folder = os.path.dirname(exe) if str(exe).lower().endswith(".exe") else str(exe)
        paths.append(os.path.join(folder, "swconst.tlb"))
    except Exception:
        pass
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"TypeLib\{SWCONST_TYPELIB}") as root:
            versions = [winreg.EnumKey(root, i) for i in range(winreg.QueryInfoKey(root)[0])]
            for ver in sorted(versions, key=lambda v: [int(x, 16) for x in v.split(".")], reverse=True):
                for arch in ("win64", "win32"):
                    try:
                        paths.append(winreg.QueryValue(root, rf"{ver}\0\{arch}"))
                    except OSError:
                        continue
    except OSError:
        pass
    paths.append(r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\swconst.tlb")
    return paths


def constant(name: str) -> int:
    """Value of a SOLIDWORKS enum member (e.g. "swFmBaseFlange"), read from
    the installed swconst type library instead of hard-coding numbers."""
    global _constants
    import pythoncom

    _load_constants()
    if name not in _constants:
        raise LookupError(f"SOLIDWORKS has no constant {name!r}")
    return _constants[name]


_enums: dict[str, dict[int, str]] = {}


def enum_names(enum: str) -> dict[int, str]:
    """{value: member name} of one swconst enum, e.g. "swFileLoadError_e"."""
    _load_constants()
    return _enums.get(enum, {})


def _load_constants() -> None:
    global _constants
    import pythoncom

    if _constants is not None:
        return
    table: dict[str, int] = {}
    for path in _swconst_paths():
        if not os.path.exists(path):
            continue
        tlb = pythoncom.LoadTypeLib(path)
        for i in range(tlb.GetTypeInfoCount()):
            ti = tlb.GetTypeInfo(i)
            attr = ti.GetTypeAttr()
            if attr.typekind != pythoncom.TKIND_ENUM:
                continue
            members = _enums.setdefault(ti.GetDocumentation(-1)[0], {})
            for j in range(attr.cVars):
                vd = ti.GetVarDesc(j)
                member = ti.GetNames(vd.memid)[0]
                table[member] = vd.value
                members.setdefault(vd.value, member)
        break
    if not table:
        raise RuntimeError("Could not read swconst.tlb from the SOLIDWORKS installation")
    _constants = table


LOAD_HINTS = {
    "swFileNotFoundError": "파일이 없습니다",
    "swFileWithSameTitleAlreadyOpen": "같은 이름의 다른 파일이 이미 열려 있습니다; 그 문서를 닫거나 파일 이름을 바꾸세요",
    "swFutureVersion": "더 새 버전의 SOLIDWORKS에서 저장된 파일입니다",
    "swInvalidFileTypeError": "지원하지 않는 파일 형식입니다",
    "swLowResourcesError": "메모리가 부족합니다",
}


def load_error(code: int) -> str:
    """swFileLoadError_e bits as names (with a hint for the common ones)."""
    names = enum_names("swFileLoadError_e") if code else {}
    parts = [names.get(bit, str(bit)) for bit in (1 << i for i in range(32)) if code & bit] or [str(code)]
    hints = [LOAD_HINTS[n] for n in parts if n in LOAD_HINTS]
    return ", ".join(parts) + (f" ({'; '.join(hints)})" if hints else "")


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


def _nearest_on_bodies(bodies: Any, entity: str, p: tuple[float, float, float]) -> tuple[float, Any]:
    best: tuple[float, Any] = (math.inf, None)
    for body in bodies or ():
        if entity == "edge":
            cand = _nearest(m0(body, "GetEdges"), p, lambda e: e.GetClosestPointOn(*p))
        elif entity == "face":
            cand = _nearest(m0(body, "GetFaces"), p, lambda f: f.GetClosestPointOn(*p))
        elif entity == "vertex":
            cand = _nearest(m0(body, "GetVertices"), p, lambda v: m0(v, "GetPoint"))
        else:
            raise ValueError(f"Unknown entity type {entity!r}")
        best = min(best, cand, key=lambda c: c[0])
    return best


def nearest(doc: Any, entity: str, point_mm: list[float]) -> tuple[float, Any]:
    """(distance m, entity) of the part's face/edge/vertex nearest a point (mm)."""
    return _nearest_on_bodies(doc.GetBodies2(0, False), entity, tuple(mm(v) for v in point_mm))


def nearest_in_assembly(asm: Any, entity: str, point_mm: list[float]) -> tuple[float, Any]:
    """Like nearest, over the bodies of every resolved component. Their
    geometry is in component space, so the point is moved there first; the
    entities found select in the assembly's context."""
    mu = m0(app(), "GetMathUtility")
    best: tuple[float, Any] = (math.inf, None)
    for comp in asm.GetComponents(False) or ():
        if m0(comp, "IsSuppressed"):
            continue
        try:
            pt = invoke(mu, "CreatePoint", doubles([mm(v) for v in point_mm]))
            local = tuple(m0(invoke(pt, "MultiplyTransform", m0(comp.Transform2, "Inverse")), "ArrayData"))[:3]
            cand = _nearest_on_bodies(comp.GetBodies2(0), entity, local)
        except Exception:
            continue  # a lightweight or empty component: nothing to pick there
        best = min(best, cand, key=lambda c: c[0])
    return best


def select_at(doc: Any, entity: str, point_mm: list[float], append: bool, mark: int = 0) -> None:
    """Select the face/edge/vertex nearest a model-space point (mm).

    This searches the solid bodies directly (in assemblies, those of every
    component), so it finds hidden edges too; SelectByID2, the fallback,
    only picks what is visible in the current view."""
    entity = entity.lower()
    p = tuple(mm(v) for v in point_mm)
    kind = m0(doc, "GetType")
    if kind in (DOC_PART, DOC_ASSEMBLY):
        finder = nearest if kind == DOC_PART else nearest_in_assembly
        best = finder(doc, entity, point_mm)  # solid bodies, hidden ones included
        if best[1] is not None and best[0] <= TOLERANCE:
            if not _select_entity(doc, best[1], append, mark):
                raise RuntimeError(f"Found the {entity} at {point_mm} mm but could not select it")
            return
        if best[1] is not None:
            raise LookupError(f"No {entity} within 0.5 mm of {point_mm} mm (nearest is {best[0] / MM:.2f} mm away)")
    x, y, z = p
    if not doc.Extension.SelectByID2("", entity.upper(), x, y, z, append, mark, nothing(), 0):
        raise LookupError(f"No {entity} found at {point_mm} mm")
