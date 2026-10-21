"""Build the Claude Desktop extension: dist/solidworks-mcp-<version>.mcpb

    .venv\\Scripts\\python scripts\\build_mcpb.py

Copies the version from pyproject.toml into manifest.json, lists the tools
(name + first sentence) for Claude Desktop's install screen, then packs the
repository with the official packer (`npx @anthropic-ai/mcpb pack`, needs
Node.js). The bundle holds only the server files (see .mcpbignore); Claude
Desktop installs Python and the packages itself with uv when it is
installed, so the person receiving it needs nothing but SOLIDWORKS and
Claude Desktop.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tomllib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)


def first_sentence(text: str) -> str:
    text = " ".join((text or "").split())
    match = re.match(r"(.+?[.:])(\s|$)", text)
    return (match.group(1) if match else text)[:200]


def main() -> int:
    with open(os.path.join(ROOT, "pyproject.toml"), "rb") as f:
        version = tomllib.load(f)["project"]["version"]
    path = os.path.join(ROOT, "manifest.json")
    with open(path, encoding="utf-8") as f:
        manifest = json.load(f)
    import server  # noqa: E402  (the registered tools)

    manifest["version"] = version
    manifest["tools"] = [{"name": name, "description": first_sentence(fn.__doc__)}
                         for name, fn in sorted(server.TOOLS.items())]
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")

    npx = shutil.which("npx")
    if not npx:
        print("Node.js (npx) is needed to pack: https://nodejs.org/ or `winget install OpenJS.NodeJS.LTS`")
        return 1
    out = os.path.join(ROOT, "dist", f"solidworks-mcp-{version}.mcpb")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    if os.path.exists(out):
        os.remove(out)
    done = subprocess.run([npx, "-y", "@anthropic-ai/mcpb", "pack", ROOT, out], cwd=ROOT)
    if done.returncode != 0 or not os.path.exists(out):
        print("packing failed")
        return 1
    print(f"\n{out} ({os.path.getsize(out) / 1024:.0f} KB, {len(manifest['tools'])} tools)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
