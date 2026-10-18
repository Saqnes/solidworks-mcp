import sys, types
sys.path.insert(0, sys.argv[1])
for m in ("pythoncom", "win32com", "win32com.client", "win32com.client.dynamic"):
    sys.modules[m] = types.ModuleType(m)
sys.modules["pythoncom"].CoInitialize = lambda: None
import server
class Dim:
    def __init__(s, n, v): s.FullName, s.SystemValue = n, v
    def GetType(s): return 1
class DD:
    def __init__(s, dims): s.dims = dims
    def GetDimension2(s, i): return s.dims[0]
    def GetNext5(s): return DD(s.dims[1:]) if len(s.dims) > 1 else None
class View:
    def __init__(s, name, dims, nxt=None): s.Name, s.dims, s.nxt = name, dims, nxt
    def GetFirstDisplayDimension5(s): return DD(s.dims) if s.dims else None
    def GetNextView(s): return s.nxt
top = View("도면뷰1", [Dim("D1@스케치2@bracket3.Part", 0.1), Dim("D2@스케치2@bracket3.Part", 0.06)], View("도면뷰2", [Dim("D1@보스-돌출1@bracket3.Part", 0.008)]))
drw = types.SimpleNamespace(GetFirstView=lambda: View("시트1", [], top))
print(server._drawing_dimensions(drw))
