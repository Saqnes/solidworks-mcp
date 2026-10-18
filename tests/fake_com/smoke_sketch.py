import sys, types, asyncio, os, json
sys.path.insert(0, sys.argv[1])
pc = types.ModuleType("pythoncom"); pc.VT_DISPATCH, pc.VT_EMPTY, pc.VT_BYREF, pc.VT_I4, pc.VT_BSTR = 9, 0, 0x4000, 3, 8
pc.CoInitialize = lambda: None; pc.com_error = type("com_error", (Exception,), {}); sys.modules["pythoncom"] = pc
class VARIANT:
    def __init__(s, vt, v): s.varianttype, s.value = vt, v
w = types.ModuleType("win32com"); wc = types.ModuleType("win32com.client"); wd = types.ModuleType("win32com.client.dynamic")
wc.VARIANT = VARIANT; w.client = wc; wc.dynamic = wd
sys.modules.update({"win32com": w, "win32com.client": wc, "win32com.client.dynamic": wd})
LOCALIZED = os.environ.get("NAMES", "ko")
log = []
def tf(normal):  # MathTransform.ArrayData with the given z axis, at origin
    return types.SimpleNamespace(ArrayData=(0,)*6 + tuple(normal) + (0, 0, 0, 1, 0, 0, 0))
class Feat:
    def __init__(s, name, t, normal=None, parents=(), children=()):
        s.Name, s.t, s.normal, s.parents, s.children = name, t, normal, list(parents), list(children)
    def GetTypeName2(s): return s.t
    def GetSpecificFeature2(s): return types.SimpleNamespace(Transform=tf(s.normal), ModelToSketchTransform=tf(s.normal))
    def Select2(s, a, m): log.append(("select", s.Name)); return True
    def GetParents(s): return s.parents
    def GetChildren(s): return s.children
    def GetNextFeature(s):
        i = d.feats.index(s); return d.feats[i + 1] if i + 1 < len(d.feats) else None
    IsSuppressed = lambda s: False
class Edge:
    def __init__(s, pts): s.pts = pts
    def GetClosestPointOn(s, x, y, z):  # nearest of sampled points
        import math; return min(s.pts, key=lambda p: math.dist(p, (x, y, z))) + (0, 0)
    def Select4(s, append, data): log.append(("edge", s.pts[0], append, data.Mark)); return True
def seg(a, b, n=20): return [tuple(a[k] + (b[k] - a[k]) * i / n for k in range(3)) for i in range(n + 1)]
# 100x60x8 box on the top plane: vertical edges at x=+-0.05, z=+-0.03, y 0..0.008
edges = [Edge(seg((sx*0.05, 0, sz*0.03), (sx*0.05, 0.008, sz*0.03))) for sx in (1, -1) for sz in (1, -1)]
class Body:
    def GetEdges(s): return edges
    def GetFaces(s): return []
    def GetVertices(s): return []
class Pt:
    def __init__(s, x, y): s.X, s.Y = x, y
class Circle:
    def __init__(s, cx, cy, r): s.c, s.r = (cx, cy), r
    def GetCenterPoint2(s): return types.SimpleNamespace(X=s.c[0], Y=s.c[1], Select4=lambda a, d_: sel.append(("center", s.c)) or True)
    def Select4(s, a, d_): sel.append(("circle", s.c)); return True
    def GetType(s): return 1
sel = []
class Line:
    ConstructionGeometry = False
    def Select4(s, a, d_): sel.append(("line", s.name)); return True
    def __init__(s, a, b, name): s.a, s.b, s.name = a, b, name
    def GetType(s): return 0
    def GetStartPoint2(s): return Pt(*s.a)
    def GetEndPoint2(s): return Pt(*s.b)
class SM:
    AddToDB = False
    def __init__(s): s.open = False; s.calls = 0; s.wrong_plane_done = False
    @property
    def ActiveSketch(s): return types.SimpleNamespace(GetConstrainedStatus=lambda: 3) if s.open else None
    def InsertSketch(s, a):
        s.calls += 1
        if s.calls == 1: log.append("insert ignored"); return  # first call on a new part does nothing
        if not s.open:
            s.open = True
            if not s.wrong_plane_done: s.wrong_plane_done = True; s.open_normal = (0, 0, 1)  # then lands on front once
            else: s.open_normal = d.selected.normal
        else:
            s.open = False; d.feats.append(Feat(f"스케치{len(d.feats)}", "ProfileFeature", s.open_normal))
    def CreateCenterRectangle(s, cx, cy, cz, x, y, z):
        w, h = x - cx, y - cy
        return (Line((cx-w, cy-h), (cx+w, cy-h), "bottom"), Line((cx+w, cy-h), (cx+w, cy+h), "right"),
                Line((cx+w, cy+h), (cx-w, cy+h), "top"), Line((cx-w, cy+h), (cx-w, cy-h), "left"))
    def CreateCircleByRadius(s, cx, cy, cz, r): return Circle(cx, cy, r)
    def CreateCenterLine(s, x1, y1, z1, x2, y2, z2):
        l = Line((x1, y1), (x2, y2), "CL-v" if x1 == x2 else "CL-h"); l.ConstructionGeometry = True; log.append(("centerline", l.name, round(max(abs(x1), abs(y1), abs(x2), abs(y2)) * 1000, 1))); return l
    def FullyDefineSketch(s, *a):
        assert len(a) == 10 and a[3] is True
        log.append(("datums", getattr(a[5], "name", "origin"), getattr(a[7], "name", "origin"))); return 0
class Ext:
    def DeleteSelection2(s, opt):
        name = log[-1][1]; log.append(("deleted", name)); d.feats = [f for f in d.feats if f.Name != name]; return True
    def SelectByID2(s, name, typ, *a):
        f = d.FeatureByName(name)
        if typ == "PLANE" and f: d.selected = f; log.append(("plane", name)); return True
        log.append("SelectByID2"); return False
class Doc:
    def __init__(s):
        names = {"ko": ("정면", "윗면", "우측면"), "other": ("Plane-A", "Plane-B", "Plane-C")}[LOCALIZED]
        # deliberately odd order: an extra origin plane first, and top/front swapped
        s.feats = [Feat("Extra", "RefPlane", (0, 0, 1)) if LOCALIZED == "other" else Feat("Origin", "OriginProfileFeature"),
                   Feat(names[1], "RefPlane", (0, 1, 0)), Feat(names[0], "RefPlane", (0, 0, 1)), Feat(names[2], "RefPlane", (1, 0, 0))]
        if LOCALIZED == "other": s.feats[0].GetSpecificFeature2 = lambda: types.SimpleNamespace(Transform=types.SimpleNamespace(ArrayData=(0,)*6 + (0, 0, 1) + (0, 0, 0.01, 1, 0, 0, 0)))
        s.SketchManager = SM(); s.Extension = Ext()
        s.SelectionManager = types.SimpleNamespace(CreateSelectData=lambda: types.SimpleNamespace(Mark=0))
        s.FeatureManager = types.SimpleNamespace(FeatureFillet3=lambda *a: Feat("필렛1", "Fillet"))
    GetType = lambda s: 1; GetTitle = lambda s: "bracket"; GetPathName = lambda s: ""; GetSaveFlag = lambda s: True
    def ClearSelection2(s, a): sel.clear()
    def FirstFeature(s): return s.feats[0]
    def FeatureByName(s, n): return next((f for f in s.feats if f.Name == n), None)
    def FeatureByPositionReverse(s, i): return s.feats[-1 - i] if i < len(s.feats) else None
    def GetBodies2(s, t, vis): return (Body(),)
    def _dim(s, kind, x, y, z): log.append((kind, [x[1] if x[0] != "center" else "c" + str(tuple(round(v*1000) for v in x[1])) for x in sel], round(x*1000), round(y*1000))) if False else None
    def AddDimension2(s, x, y, z):
        log.append(("dim", list(sel), round(x*1000, 1), round(y*1000, 1))); sel.clear()
        return types.SimpleNamespace(SetText=lambda part, txt: log.append(("prefix", part, txt)))
    def AddHorizontalDimension2(s, x, y, z): log.append(("hdim", list(sel))); sel.clear(); return object()
    def AddVerticalDimension2(s, x, y, z): log.append(("vdim", list(sel))); sel.clear(); return object()
    def SketchAddConstraints(s, c): log.append(("constraint", c, len(sel))); sel.clear()
    def ForceRebuild3(s, a): pass
    def SetMaterialPropertyName2(s, c, db, name): s.mat = name
    def GetMaterialPropertyName2(s, c, db): return getattr(s, "mat", "")
d = Doc()
class App:
    ActiveDoc = d
    pref = True
    def GetUserPreferenceToggle(s, k): assert k == 10; return App.pref
    def SetUserPreferenceToggle(s, k, v): assert k == 10; App.pref = v; log.append(("pref", v))
    def ActivateDoc3(s, *a): return d
    def GetMaterialDatabases(s): return (os.path.join(os.path.dirname(os.path.abspath(__file__)), "matdb", "solidworks materials.sldmat"),)
import sw; sw.app = lambda: App()
d.Extension.CreateMassProperty = lambda: types.SimpleNamespace(Density=2700.0)
import server
from mcp.client import Client
async def go():
    async with Client(server.mcp) as c:
        async def t(name, args={}):
            r = await c.call_tool(name, args)
            print(("ERR " if r.is_error else "ok  ") + name, (r.content[0].text if r.content else "").replace("\n", " ")[:230])
        await t("sketch", {"plane": "top", "entities": [{"type": "center_rectangle", "width": 100, "height": 60}]})
        await t("sketch", {"plane": "top", "entities": [{"type": "circle", "cx": x, "cy": y, "r": 3.3} for x in (-40, 40) for y in (-20, 20)] + [{"type": "circle", "cx": 0, "cy": 0, "r": 5}]})
        await t("fillet", {"radius": 5, "edge_points": [[50, 4, -30], [-50, 4, 30]]})
        await t("fillet", {"radius": 5, "edge_points": [[50, 4, -20]]})
        other = Doc(); other.GetTitle = lambda: "bracket3.SLDPRT"; other.GetPathName = lambda: r"C:\\sw_test\\bracket3.SLDPRT"
        App.GetDocuments = lambda self: (d, other)
        await t("save", {"path": r"C:\\sw_test\\bracket3.SLDPRT"})
        await t("set_material", {"material": "6061 alloy"})
        await t("set_material", {"material": "6061"})
        await t("list_materials", {"query": "6061"})
        # delete: feature with absorbed sketch
        sk = sw.last_feature_of_type(d, "ProfileFeature"); ext = Feat("보스-돌출1", "Extrusion", parents=[sk]); sk.children = [ext]; d.feats.append(ext)
        await t("delete_feature", {"name": "보스-돌출1"})
asyncio.run(go())
print(*[x for x in log if x != "SelectByID2" and x[0] in ("dim","hdim","vdim","constraint","pref","datums","centerline","prefix")], sep="\n")
