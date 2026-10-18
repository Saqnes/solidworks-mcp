import sys, types, asyncio, os
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_smoke3.py")).read().split("import server")[0])
import server
ev = []
class Face:
    Normal = (0.0, 1.0, 0.0)
    def GetClosestPointOn(s, x, y, z): return (x, 0.01, z, 0, 0)
class Body:
    def GetFaces(s): return [Face()]
d.GetBodies2 = lambda t, v: [Body()]
d.Extension.SelectByRay = lambda *a: ev.append(("ray", tuple(round(v, 4) for v in a[:6]))) or True
d.Extension.DeleteSelection2 = lambda o: ev.append(("delete", o)) or True
d.SelectionManager = types.SimpleNamespace(CreateSelectData=lambda: types.SimpleNamespace(Mark=0))
class Pt:
    def __init__(s, x, y): s.x, s.y = x, y
    def Select4(s, a, data): ev.append(("selpt", s.x, s.y, a)); return True
sk_pts = [Pt(0.001, 0.002)]
class SM:
    AddToDB = False; ActiveSketch = None
    def CreatePoint(s, x, y, z): sk_pts.append(Pt(x, y)); ev.append(("point", round(x*1000, 1), round(y*1000, 1)))
    def InsertSketch(s, a): s.ActiveSketch = None; ev.append("exit sketch")
d.SketchManager = SM()
def edit_sketch(): d.SketchManager.ActiveSketch = object(); ev.append("edit sketch")
d.EditSketch = edit_sketch
def edit_delete():
    sk_pts.clear(); ev.append("EditDelete")
d.EditDelete = edit_delete
placement = types.SimpleNamespace(Select2=lambda a, m: True, GetSpecificFeature2=lambda: types.SimpleNamespace(GetSketchPoints2=lambda: tuple(sk_pts)))
def hw5(*a):
    assert len(a) == 27; ev.append(("HoleWizard5", a[:5])); f = newfeat("CBORE for M6", "HoleWzd"); f.GetFirstSubFeature = lambda: placement; return f
d.FeatureManager.HoleWizard5 = hw5
server.sketch.__wrapped__ = lambda entities, **kw: ev.append(("fallback sketch", len(entities))) or {"sketch": "S"}
server.extrude.__wrapped__ = lambda **kw: {"feature": "컷"}
# base flange via definition
sw.constant = lambda n: {"swFmBaseFlange": 61}[n]
class Def:
    def Initialize(s, *a): ev.append(("Initialize", len(a)))
data = Def()
d.FeatureManager.CreateCustomBendAllowance = lambda: types.SimpleNamespace(KFactor=0)
d.FeatureManager.CreateDefinition = lambda k: ev.append(("CreateDefinition", k)) or data
d.FeatureManager.CreateFeature = lambda dd: ev.append(("CreateFeature", dd.OffsetDirections, round(dd.D1EndConditionDistance*1000, 1))) or newfeat("베이스-플랜지1", "SMBaseFlange")
d.feats.append(Feat("L자", "ProfileFeature"))
server._sketch_normal_axis = lambda f: 2
d.GetPartBox = lambda a: (0, 0, 0, 0.05, 0.03, 0.04)
from mcp.client import Client
async def go():
    async with Client(server.mcp) as c:
        async def t(name, args={}):
            r = await c.call_tool(name, args)
            print(("ERR " if r.is_error else "ok  ") + name, (r.content[0].text if r.content else "").replace("\n", " ")[:230])
        await t("hole", {"size": "M6", "kind": "counterbore", "positions": [[-35, 0], [35, 0]], "face_point": [0, 10, 0]})
        print(*ev, sep="\n"); ev.clear()
        sk_pts[:] = [Pt(0, 0)]
        d.SketchManager.CreatePoint = lambda x, y, z: ev.append("CreatePoint fails")  # positions not placed
        await t("hole", {"size": "M6", "kind": "counterbore", "positions": [[-35, 0], [35, 0]], "face_point": [0, 10, 0]})
        print(*ev, sep="\n"); ev.clear()
        await t("sheet_metal_base_flange", {"thickness": 1.5, "bend_radius": 1, "length": 40, "sketch_name": "L자"})
        print(*ev, sep="\n")
asyncio.run(go())
