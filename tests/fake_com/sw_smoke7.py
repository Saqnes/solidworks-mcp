import sys, types, asyncio, os
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_smoke6.py")).read().split("from mcp.client import Client")[0])
# clearance variants: the first still has a cone, the second does not
cone_by_call = iter([True, False])
def hw5(*a):
    assert len(a) == 27, len(a)
    cone = next(cone_by_call, False)
    ev.append(("HoleWizard5", a[0], a[8:12], "cone" if cone else "no cone"))
    f = newfeat("M8 여유 구멍1", "HoleWzd"); f.GetFirstSubFeature = lambda: placement
    surf = types.SimpleNamespace(IsCone=lambda: cone)
    f.GetFaces = lambda: [types.SimpleNamespace(GetSurface=lambda: surf)]
    return f
d.FeatureManager.HoleWizard5 = hw5
sw.constant = lambda n: {"swFmBaseFlange": 61, "swViewEntityType_Edge": 1}[n]
# drawing with a top view showing two circles of one wizard hole and one plain circle
wz = types.SimpleNamespace(Name="M6 카운터보어1", GetTypeName2=lambda: "HoleWzd")
other = types.SimpleNamespace(Name="보스-돌출1", GetTypeName2=lambda: "Extrusion")
def circle_edge(feat):
    face = types.SimpleNamespace(GetFeature=lambda: feat)
    return types.SimpleNamespace(GetCurve=lambda: types.SimpleNamespace(IsCircle=lambda: True), GetTwoAdjacentFaces2=lambda: (face, None))
line = types.SimpleNamespace(GetCurve=lambda: types.SimpleNamespace(IsCircle=lambda: False))
top = types.SimpleNamespace(Name="도면뷰2", GetVisibleEntities2=lambda c, k: [line, circle_edge(wz), circle_edge(wz), circle_edge(other)],
                            SelectEntity=lambda e, a: True, GetOutline=(0.1, 0.1, 0.2, 0.2), GetNextView=lambda: None)
sheet = types.SimpleNamespace(Name="시트1", GetNextView=lambda: top)
drw = types.SimpleNamespace(GetType=lambda: 3, GetTitle=lambda: "도면1", GetPathName=lambda: "", GetSaveFlag=lambda: True,
    GetFirstView=lambda: sheet, ActivateView=lambda n: True, ClearSelection2=lambda a: None, ForceRebuild3=lambda a: None,
    AddHoleCallout2=lambda x, y, z: ev.append(("callout", round(x, 3), round(y, 3))) or object())
from mcp.client import Client
async def go():
    async with Client(server.mcp) as c:
        async def t(name, args={}):
            r = await c.call_tool(name, args)
            print(("ERR " if r.is_error else "ok  ") + name, (r.content[0].text if r.content else "").replace("\n", " ")[:260])
        sk_pts[:] = [Pt(0, 0)]
        await t("hole", {"size": "M8", "kind": "clearance", "positions": [[0, 0]], "face_point": [0, 10, 0]})
        App.ActiveDoc = drw
        await t("add_hole_callouts")
        App.ActiveDoc = d
        await t("sweep", {"path": "L자", "circle_diameter": 10, "wall": 1})
asyncio.run(go())
print(*[e for e in ev if e[0] in ("HoleWizard5", "delete", "callout")], sep="\n")
