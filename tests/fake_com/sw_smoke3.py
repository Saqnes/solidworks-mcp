import sys, types, asyncio, os, tempfile, pathlib
sys.path.insert(0, sys.argv[1])
pc = types.ModuleType("pythoncom"); pc.VT_DISPATCH, pc.VT_EMPTY, pc.VT_BYREF, pc.VT_I4, pc.VT_BSTR, pc.VT_ARRAY, pc.VT_R8 = 9, 0, 0x4000, 3, 8, 0x2000, 5
pc.CoInitialize = lambda: None; pc.DISPATCH_PROPERTYPUT = 4
class com_error(Exception):
    @property
    def hresult(s): return s.args[0]
pc.com_error = com_error; sys.modules["pythoncom"] = pc
class VARIANT:
    def __init__(s, vt, v): s.varianttype, s.value = vt, v
w = types.ModuleType("win32com"); wc = types.ModuleType("win32com.client"); wd = types.ModuleType("win32com.client.dynamic")
wc.VARIANT = VARIANT; w.client = wc; wc.dynamic = wd
sys.modules.update({"win32com": w, "win32com.client": wc, "win32com.client.dynamic": wd})
log = []
class Feat:
    def __init__(s, name, t, normal=None): s.Name, s.t, s.normal = name, t, normal
    def GetTypeName2(s): return s.t
    def Select2(s, a, m): log.append(("sel", s.Name, a, m)); return True
    def GetNextFeature(s):
        i = d.feats.index(s); return d.feats[i + 1] if i + 1 < len(d.feats) else None
def newfeat(name, t): f = Feat(name, t); d.feats.append(f); return f
class FM:
    def FeatureLinearPattern5(s, *a): raise com_error(-2147352562, "bad param count", None, None)
    def FeatureLinearPattern4(s, *a): assert len(a) == 20; log.append(("LP4", a[:4])); return newfeat("선형 패턴1", "LPattern")
    def FeatureCircularPattern5(s, *a): assert len(a) == 14; log.append(("CP5", a[:2])); return newfeat("원형 패턴1", "CirPattern")
    def InsertMirrorFeature2(s, *a): assert len(a) == 5; return newfeat("대칭복사1", "MirrorPattern")
    def InsertRefPlane(s, *a): assert len(a) == 6; log.append(("plane", a[:2])); return newfeat("평면1", "RefPlane")
class EQ:
    def __init__(s): s.eqs = []
    def GetCount(s): return len(s.eqs)
    def Equation(s, i): return s.eqs[i]
    def Value(s, i): return float(s.eqs[i].split("=")[1].strip().strip('"')) if s.eqs[i].split("=")[1].strip().replace(".", "").isdigit() else None
    def Add2(s, i, t, solve):
        name = t.split("=")[0].strip()
        if any(name in e.split("=", 1)[1] for e in s.eqs):  # re-adding a referenced variable is refused
            return -1
        s.eqs.insert(len(s.eqs) if i < 0 else i, t); return 0
    @property
    def _oleobj_(s):
        eq = s
        class O:
            def GetIDsOfNames(self, n): return 7
            def Invoke(self, dispid, lcid, flags, want, i, v):
                if EQ.put_fails: raise com_error(-2147352567, "exception", None, None)
                eq.eqs[i] = v
        return O()
    put_fails = False
    def Delete(s, i): s.eqs.pop(i)
class Doc:
    def __init__(s):
        s.feats = [Feat("정면", "RefPlane"), Feat("윗면", "RefPlane"), Feat("우측면", "RefPlane"), Feat("컷-돌출1", "ICE")]
        s.FeatureManager = FM(); s.eq = EQ(); s.cfgs = ["기본"]
        s.ConfigurationManager = types.SimpleNamespace(ActiveConfiguration=types.SimpleNamespace(Name="기본"))
        class M:
            ArcOption = 0; Calculate = True; Angle = -1; Length = -1; Area = -1; CenterDistance = 0.0734; NormalDistance = -1; Diameter = -1; Radius = -1; Perimeter = -1
            @property
            def Distance(self): return 0.08 if self.ArcOption == 0 else 0.0734
        s.Extension = types.SimpleNamespace(CreateMeasure=lambda: M())
    GetType = lambda s: 1; GetTitle = lambda s: "bracket6"; GetPathName = lambda s: ""
    def ClearSelection2(s, a): pass
    def FirstFeature(s): return s.feats[0]
    def FeatureByName(s, n): return next((f for f in s.feats if f.Name == n), None)
    def FeatureByPositionReverse(s, i): return s.feats[-1 - i] if i < len(s.feats) else None
    def InsertAxis2(s, a): newfeat("기준축1", "RefAxis"); return True
    def GetEquationMgr(s): return s.eq
    def ForceRebuild3(s, a): pass
    def AddConfiguration3(s, n, *a): s.cfgs.append(n); s.ConfigurationManager.ActiveConfiguration.Name = n; return object()
    def ShowConfiguration2(s, n):
        if n in s.cfgs: s.ConfigurationManager.ActiveConfiguration.Name = n; return True
        return False
    def GetConfigurationNames(s): return tuple(s.cfgs)
    def GraphicsRedraw2(s): pass
    def GetBodies2(s, *a): return ()
    def Parameter(s, n):
        doc = s
        class Dim:
            def GetType(self): return 1
            @property
            def SystemValue(self): return doc.vals.get(doc.ConfigurationManager.ActiveConfiguration.Name, 0.008)
            def SetSystemValue3(self, v, which, names):
                for c in (doc.cfgs if which == 2 else [doc.ConfigurationManager.ActiveConfiguration.Name]): doc.vals[c] = v
                return 0
        return Dim()
    vals = {}
d = Doc()
class App:
    ActiveDoc = d
import sw; sw.app = lambda: App()
sw.select_at = lambda doc, ent, p, append, mark=0: log.append(("pick", ent, p, mark))
import server
from mcp.client import Client
async def go():
    async with Client(server.mcp) as c:
        async def t(name, args={}):
            r = await c.call_tool(name, args)
            print(("ERR " if r.is_error else "ok  ") + name, (r.content[0].text if r.content else "").replace("\n", " ")[:200])
        await t("linear_pattern", {"features": ["컷-돌출1"], "direction": "x", "count": 3, "spacing": 25})
        await t("circular_pattern", {"features": ["컷-돌출1"], "count": 6})
        await t("mirror", {"features": ["컷-돌출1"], "plane": "right"})
        await t("offset_plane", {"distance": 20, "plane": "top"})
        await t("configurations", {"add": "긴형"})
        await t("configurations", {"activate": "없음"})
        await t("set_dimension", {"name": "D1@보스-돌출1", "value": 5})
        await t("configurations", {"activate": "기본"})
        print("default thickness:", d.vals.get("기본", 0.008), "thin:", d.vals.get("긴형"))
        await t("set_equation", {"lhs": "Width", "rhs": "100"})
        await t("set_equation", {"lhs": "D1@스케치2", "rhs": '"Width"'})
        d.eq.eqs = ['"Width" = 100', '"D1@스케치2" = "Width"']
        await t("set_equation", {"lhs": "Width", "rhs": "120"})
        print("eqs:", d.eq.eqs)
        EQ.put_fails = True
        await t("set_equation", {"lhs": "Width", "rhs": "130"})
        print("eqs after refused edit:", d.eq.eqs)
        await t("measure", {"point1": [-40, 4, -20], "point2": [40, 4, -20]})
        d.MaterialPropertyValues = None
        await t("set_color", {"rgb": [200, 30, 30]})
asyncio.run(go())

print("color:", d.MaterialPropertyValues.value)
