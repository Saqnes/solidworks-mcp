import sys, types, asyncio, os
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "sw_smoke3.py")).read().split("import server")[0])  # reuse fakes from smoke3
import server
calls = []
server.sketch.__wrapped__ = lambda entities, **kw: calls.append(("sketch", [round(e["r"]*2, 2) for e in entities], kw)) or {"sketch": f"스케치{len(calls)}"}
server.extrude.__wrapped__ = lambda **kw: calls.append(("cut", {k: kw[k] for k in ("depth", "end", "draft_angle") if k in kw})) or {"feature": f"컷-돌출{len(calls)}"}
def fm(name, n):
    def f(*a):
        assert len(a) == n, f"{name}: {len(a)} args"
        calls.append((name, len(a))); return newfeat(name, name)
    return f
d.FeatureManager.InsertProtrusionSwept4 = fm("InsertProtrusionSwept4", 20)
d.FeatureManager.InsertCutSwept5 = fm("InsertCutSwept5", 22)
d.FeatureManager.InsertProtrusionBlend2 = fm("InsertProtrusionBlend2", 18)
d.FeatureManager.InsertRib = fm("InsertRib", 10)
d.GetMassProperties = lambda: (0, 0, 0, 1e-6 * len(calls), 0, 0)  # rib checks that the volume grew
d.FeatureManager.InsertSheetMetalBaseFlange2 = fm("InsertSheetMetalBaseFlange2", 19)
d.feats += [Feat("경로", "3DProfileFeature"), Feat("단면", "ProfileFeature"), Feat("단면2", "ProfileFeature")]
from mcp.client import Client
async def go():
    async with Client(server.mcp) as c:
        async def t(name, args={}):
            r = await c.call_tool(name, args)
            print(("ERR " if r.is_error else "ok  ") + name, (r.content[0].text if r.content else "").replace("\n", " ")[:170])
        await t("hole", {"size": "m6", "positions": [[-40, -20], [40, -20]], "kind": "counterbore", "face_point": [0, 8, 0]})
        await t("hole", {"size": "M4", "positions": [[0, 0]], "kind": "countersink"})
        await t("hole", {"size": "M5", "positions": [[0, 0]], "kind": "tapped", "depth": 12})
        await t("hole", {"size": "M7", "positions": [[0, 0]]})
        await t("sweep", {"path": "경로", "circle_diameter": 6, "wall": 1})
        await t("sweep", {"path": "경로", "profile": "단면", "cut": True})
        await t("loft", {"profiles": ["단면", "단면2"]})
        await t("rib", {"thickness": 3})
        await t("sheet_metal_base_flange", {"thickness": 1.5, "bend_radius": 1, "length": 40})
asyncio.run(go())
print(*calls, sep="\n")
