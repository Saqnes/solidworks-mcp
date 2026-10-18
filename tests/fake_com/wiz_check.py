import sys, types, math
sys.path.insert(0, sys.argv[1])
for m in ("pythoncom", "win32com", "win32com.client", "win32com.client.dynamic"): sys.modules[m] = types.ModuleType(m)
sys.modules["pythoncom"].CoInitialize = lambda: None
import server
T, F, PI2 = True, False, 1.5707963267949
rec = {
 "A cbore M6 thru": ((0, 8, 139, "M6", 1, 0.0066, 0.01, -1, 0.011, 0.0064, 0, 1, 0, 0.01105, PI2, 0, 0, 0, 0, 0, "", F, T, T, T, T, F), ("counterbore", "M6", True, 0, True, 0)),
 "B tap M4 thru cosm": ((4, 8, 147, "M4", 1, 0.0033, 0.01, T, 0.009625, 0.00405, PI2, 0, 0, 0, 1, 1, 0, 0, -1, -1, "", F, T, T, T, T, F), ("tapped", "M4", True, 0, True, 0)),
 "C tap M4 blind cosm": ((4, 8, 147, "M4", 0, 0.0033, 0.008, T, 0.006, 0.00405, PI2, 0, 0, 2.05948851735331, 1, 0, 0, 0, -1, -1, "", F, T, T, T, T, F), ("tapped", "M4", False, 8, True, 6)),
 "D tap M4 blind drill": ((4, 8, 147, "M4", 0, 0.0033, 0.008, T, 0.006, 0.00405, PI2, 0, 0, 2.05948851735331, 0, 0, 0, 0, -1, -1, "", F, T, T, T, T, F), ("tapped", "M4", False, 8, False, 6)),
 "E clear M6 thru": ((2, 8, 144, "M6", 1, 0.0066, 0.01, -1, 1, 0, 0.00665, PI2, 0, 0, 0, -1, -1, -1, -1, -1, "", F, T, T, T, T, F), ("clearance", "M6", True, 0, True, 0)),
 "F csk M5 thru": ((1, 8, 140, "M5", 1, 0.0055, 0.01, -1, 0.0112, PI2, 0, 1, 0, 0, 0, 0, 0, -1, -1, -1, "", F, T, T, T, T, F), ("countersink", "M5", True, 0, True, 0)),
}
for name, (expected, (kind, size, thru, depth, cosm, td)) in rec.items():
    got = server._wizard_args(kind, size, thru, depth, cosm, td)
    bad = [(i + 1, e, g) for i, (e, g) in enumerate(zip(expected, got))
           if not (isinstance(e, str) and e == g or not isinstance(e, str) and abs(float(e) - float(g)) < 1e-6)]
    print(f"{name:22s} len {len(got)}  {'MATCH' if not bad and len(got) == 27 else bad}")
