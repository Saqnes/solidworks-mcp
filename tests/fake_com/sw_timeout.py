import sys, types, time
sys.path.insert(0, sys.argv[1])
pc = types.ModuleType("pythoncom"); pc.CoInitialize = lambda: None; sys.modules["pythoncom"] = pc
import sw
print(sw.call(lambda: 42))
try: sw.call(lambda: time.sleep(2), timeout=0.3)
except RuntimeError as e: print("timeout:", str(e)[:60])
try: sw.call(lambda: 1)
except RuntimeError as e: print("busy:", str(e)[:60])
time.sleep(2); print("after:", sw.call(lambda: 7))
