"""Run every fake-COM smoke script against the server in the repository root.

    python -I tests/fake_com/run_all.py

Each script replaces pywin32 and SOLIDWORKS with small fakes, drives the
MCP tools through a real MCP client and prints one "ok"/"ERR" line per call.
Some ERR lines are intended (see README.md); anything else is a regression.
"""

import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
SCRIPTS = ["smoke_sketch.py", "sw_smoke3.py", "sw_smoke4.py", "sw_smoke5.py", "sw_smoke6.py",
           "sw_smoke7.py", "wiz_check.py", "sw_timeout.py", "drw_test.py"]

for name in SCRIPTS:
    env = dict(os.environ, NAMES=os.environ.get("NAMES", "ko"))
    out = subprocess.run([sys.executable, "-I", os.path.join(HERE, name), ROOT], cwd=HERE,
                         capture_output=True, text=True, env=env, timeout=300).stdout
    lines = [l for l in out.splitlines() if l.startswith(("ok ", "ERR ")) or "MATCH" in l]
    print(f"== {name}")
    for line in lines or out.splitlines()[-5:]:
        print("  " + line[:160])
