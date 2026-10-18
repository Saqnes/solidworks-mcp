import sys, types, asyncio, os
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_smoke3.py")).read().split("import server")[0])
import server
class SM:
    AddToDB = False
    def __init__(s): s.active = None; s.n3d = 0
    @property
    def ActiveSketch(s): return s.active
    def Insert3DSketch(s, a):
        s.n3d += 1
        if s.active is not None: s.active = None; return
        if s.n3d == 1: log.append("3d open ignored"); return       # first opener does nothing
        s.active = types.SimpleNamespace(Is3D=lambda: True); d.feats.append(Feat("3D스케치1", "3DProfileFeature"))
    def CreateLine(s, *a):
        assert s.active is not None
        return None if s.AddToDB else object()                     # only works without AddToDB
d.SketchManager = SM()
def ins3d2(a):
    log.append("Insert3DSketch2"); d.SketchManager.active = types.SimpleNamespace(Is3D=lambda: True); d.feats.append(Feat("3D스케치1", "3DProfileFeature"))
d.Insert3DSketch2 = ins3d2
calls = []
server.sketch.__wrapped__ = lambda entities, **kw: calls.append(("sketch", [e.get("r", e.get("text")) for e in entities], kw)) or {"sketch": f"S{len(calls)}"}
server.extrude.__wrapped__ = lambda **kw: calls.append(("cut", kw.get("sketch_name"), kw.get("depth"), kw.get("cut"))) or {"feature": f"F{len(calls)}", "type": "ICE"}
# sheet metal: SOLIDWORKS builds each direction by the distance given
def flange(*a):
    d.box = (0, 0, -a[3], 0.05, 0.03, a[4]); return newfeat("베이스-플랜지1", "SMBaseFlange")
d.FeatureManager.InsertSheetMetalBaseFlange2 = flange
d.GetPartBox = lambda a: d.box
d.feats.append(Feat("L자", "ProfileFeature"))
server._sketch_normal_axis = lambda f: 2
from mcp.client import Client
async def go():
    async with Client(server.mcp) as c:
        async def t(name, args={}):
            r = await c.call_tool(name, args)
            print(("ERR " if r.is_error else "ok  ") + name, (r.content[0].text if r.content else "").replace("\n", " ")[:220])
        await t("sketch3d", {"entities": [{"type": "polyline", "points": [[0,0,0],[0,50,0],[50,50,0]]}]})
        await t("engrave", {"text": "MCP", "face_point": [0, 10, 0], "height": 6, "depth": 0.5})
        await t("hole", {"size": "M6", "positions": [[-35, 0], [35, 0]], "kind": "counterbore", "face_point": [-35, 10, 0]})
        await t("sheet_metal_base_flange", {"thickness": 1.5, "bend_radius": 1, "length": 40, "sketch_name": "L자"})
        d.eq.eqs = ['"Width" = 100', '"D1@스케치2" = "Width"']
        orig = sw.put_indexed
        def reformat(obj, name, i, v): d.eq.eqs[i] = v.replace(" = ", "= ").replace("120", "120mm")
        sw.put_indexed = reformat
        await t("set_equation", {"lhs": "Width", "rhs": "120"})
        print("eqs:", d.eq.eqs)
asyncio.run(go())
print(*[c for c in calls], sep="\n"); print([x for x in log if isinstance(x, str)])
