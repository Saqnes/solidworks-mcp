"""Setup check and Claude registration for solidworks-mcp.

    python doctor.py               check this PC (Python, packages, SOLIDWORKS, PDF printer, Claude config)
    python doctor.py --register    add/update the "solidworks" server in Claude Desktop's config
    python doctor.py --unregister  remove it again

The check never starts SOLIDWORKS; it only looks at a running instance.
The config is merged (other servers and settings are kept) and backed up
next to itself before every change.
"""

from __future__ import annotations

import glob
import json
import os
import shutil
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SERVER = os.path.join(HERE, "server.py")
NAME = "solidworks"
SWCONST_TYPELIB = "{4687F359-55D0-4CD3-B6CF-2EB42C11F989}"
PDF_PRINTER = "Microsoft Print to PDF"

results: list[tuple[str, str, str]] = []  # (status, item, detail)


def report(status: str, item: str, detail: str = "") -> None:
    results.append((status, item, detail))
    mark = {"ok": "[ OK ]", "warn": "[주의]", "fail": "[실패]", "info": "[정보]"}[status]
    print(f"{mark} {item}" + (f" - {detail}" if detail else ""))


def venv_python() -> str:
    """The interpreter the server should run with: this one (normally .venv)."""
    return os.path.abspath(sys.executable)


# --- Claude Desktop config -------------------------------------------------


def desktop_configs(existing_only: bool = False) -> list[str]:
    """Claude Desktop config files: the classic one under %APPDATA% and the
    Microsoft Store (MSIX) one, whose folder only exists for Store installs."""
    paths = [os.path.expandvars(r"%APPDATA%\Claude\claude_desktop_config.json")]
    paths += glob.glob(os.path.expandvars(r"%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\claude_desktop_config.json"))
    store_dirs = glob.glob(os.path.expandvars(r"%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude"))
    for d in store_dirs:
        p = os.path.join(d, "claude_desktop_config.json")
        if p not in paths:
            paths.append(p)
    if existing_only:
        return [p for p in paths if os.path.exists(p)]
    # write the classic path, plus the Store path when Claude is a Store install
    return [p for p in paths if os.path.exists(p) or os.path.isdir(os.path.dirname(p))] or paths[:1]


def load_config(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    with open(path, "rb") as f:
        raw = f.read()
    text = raw.decode("utf-8-sig").strip()
    return json.loads(text) if text else {}


def save_config(path: str, cfg: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if os.path.exists(path):
        stamp = time.strftime("%Y%m%d-%H%M%S")
        backup, n = f"{path}.{stamp}.bak", 1
        while os.path.exists(backup):  # two changes in the same second keep both backups
            n += 1
            backup = f"{path}.{stamp}-{n}.bak"
        shutil.copy2(path, backup)
        print(f"       백업: {backup}")
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:  # UTF-8 without BOM, as Claude writes it
        json.dump(cfg, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(tmp, path)


def server_entry() -> dict:
    return {"command": venv_python(), "args": [SERVER]}


def register() -> int:
    done = 0
    for path in desktop_configs():
        try:
            cfg = load_config(path)
        except ValueError as exc:
            report("fail", f"Claude Desktop 설정을 읽지 못함: {path}", f"JSON 오류 ({exc}); 직접 고친 뒤 다시 실행하세요")
            continue
        servers = cfg.setdefault("mcpServers", {})
        entry = servers.get(NAME, {})
        new = dict(entry)
        new.update(server_entry())
        if entry == new:
            report("ok", f"Claude Desktop에 이미 등록됨: {path}")
        else:
            servers[NAME] = new
            save_config(path, cfg)
            report("ok", f"Claude Desktop에 등록: {path}")
        done += 1
    if done:
        report("info", "Claude Desktop을 완전히 종료(트레이 아이콘 > 종료)한 뒤 다시 실행해야 적용됩니다")
    return 0 if done else 1


def unregister() -> int:
    for path in desktop_configs(existing_only=True):
        cfg = load_config(path)
        if NAME in cfg.get("mcpServers", {}):
            del cfg["mcpServers"][NAME]
            save_config(path, cfg)
            report("ok", f"Claude Desktop에서 제거: {path}")
    return 0


# --- Checks ------------------------------------------------------------------


def check_python() -> None:
    v = sys.version_info
    if v >= (3, 11):
        report("ok", f"Python {v.major}.{v.minor}.{v.micro}", venv_python())
    else:
        report("fail", f"Python {v.major}.{v.minor}", "3.11 이상이 필요합니다")
    if ".venv" not in venv_python().lower():
        report("warn", "가상환경(.venv) 밖의 Python입니다", "install.cmd로 설치하면 .venv를 만듭니다")


def check_packages() -> None:
    for module, pip_name in (("mcp", "mcp"), ("win32com.client", "pywin32"), ("PIL", "pillow")):
        try:
            __import__(module)
            report("ok", f"패키지 {pip_name}")
        except Exception as exc:
            report("fail", f"패키지 {pip_name} 없음", f"{exc}; '.venv\\Scripts\\pip install -e .' 실행")


def check_solidworks() -> None:
    try:
        import winreg
    except ImportError:
        report("fail", "Windows가 아닙니다", "SOLIDWORKS가 설치된 Windows PC에서 실행하세요")
        return
    progid = os.environ.get("SW_PROGID", "SldWorks.Application")
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{progid}\CLSID"):
            pass
        versions = []
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as root:
            i = 0
            while True:
                try:
                    k = winreg.EnumKey(root, i)
                except OSError:
                    break
                if k.startswith("SldWorks.Application.") and k.rsplit(".", 1)[-1].isdigit():
                    versions.append(int(k.rsplit(".", 1)[-1]))
                i += 1
        years = ", ".join(f"{v + 1992}" for v in sorted(versions)) or "?"
        report("ok", "SOLIDWORKS COM 등록됨", f"{progid}; 설치된 버전 {years}")
    except OSError:
        report("fail", "SOLIDWORKS COM이 등록되어 있지 않음", f"{progid}; SOLIDWORKS가 설치된 PC인지 확인하세요")
        return
    # swconst.tlb: the enum values the server reads instead of hard-coding them
    found = None
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"TypeLib\{SWCONST_TYPELIB}") as root:
            for j in range(winreg.QueryInfoKey(root)[0]):
                ver = winreg.EnumKey(root, j)
                for arch in ("win64", "win32"):
                    try:
                        p = winreg.QueryValue(root, rf"{ver}\0\{arch}")
                    except OSError:
                        continue
                    if os.path.exists(p):
                        found = p
    except OSError:
        pass
    found = found or next((p for p in (r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\swconst.tlb",) if os.path.exists(p)), None)
    if found:
        report("ok", "swconst.tlb", found)
    else:
        report("warn", "swconst.tlb를 찾지 못함", "SOLIDWORKS가 실행 중이면 그 설치 폴더에서 다시 찾습니다")
    # a running instance: version and default templates (never start one here)
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        app = win32com.client.dynamic.Dispatch(win32com.client.GetActiveObject(progid)._oleobj_)
    except Exception:
        report("info", "SOLIDWORKS가 실행 중이 아님", "도구를 처음 쓸 때 서버가 SOLIDWORKS를 띄웁니다")
        return
    rev = app.RevisionNumber() if callable(getattr(app, "RevisionNumber", None)) else app.RevisionNumber
    report("ok", "실행 중인 SOLIDWORKS에 연결됨", f"revision {rev}")
    for pref, label in ((8, "파트"), (9, "어셈블리"), (10, "도면")):
        path = app.GetUserPreferenceStringValue(pref)
        if path and os.path.exists(path):
            report("ok", f"기본 {label} 템플릿", path)
        else:
            report("fail", f"기본 {label} 템플릿이 없음", "도구 > 옵션 > 시스템 옵션 > 기본 템플릿에서 지정하세요")


def check_pdf() -> None:
    try:
        import win32print

        printers = [p[2] for p in win32print.EnumPrinters(2 | 4)]
    except Exception:
        printers = []
    import ctypes

    utf8 = ctypes.windll.kernel32.GetACP() == 65001
    if PDF_PRINTER in printers:
        report("ok", f"프린터 '{PDF_PRINTER}'", "한글이 든 도면 PDF에 씁니다" if utf8 else "")
    elif utf8:
        report("warn", f"프린터 '{PDF_PRINTER}' 없음",
               "Windows UTF-8 모드라 SOLIDWORKS 자체 PDF 내보내기는 한글을 빠뜨립니다; Windows 기능에서 'Microsoft Print to PDF'를 켜세요")
    else:
        report("info", f"프린터 '{PDF_PRINTER}' 없음", "도면 PDF는 SOLIDWORKS 내보내기로 만듭니다")


def check_registration() -> None:
    configs = desktop_configs(existing_only=True)
    if not configs:
        report("warn", "Claude Desktop 설정 파일이 없음", "Claude Desktop을 쓰면 'python doctor.py --register'")
    for path in configs:
        try:
            entry = load_config(path).get("mcpServers", {}).get(NAME)
        except ValueError as exc:
            report("fail", f"Claude Desktop 설정 JSON 오류: {path}", str(exc))
            continue
        if not entry:
            report("warn", f"Claude Desktop에 '{NAME}' 미등록", f"{path}; 'install.cmd' 또는 'python doctor.py --register'")
        elif os.path.normcase(entry.get("command", "")) != os.path.normcase(venv_python()) or \
                [os.path.normcase(a) for a in entry.get("args", [])] != [os.path.normcase(SERVER)]:
            report("warn", f"Claude Desktop의 '{NAME}'가 다른 경로를 가리킴", f"{entry.get('command')} {entry.get('args')}; '--register'로 고칩니다")
        else:
            report("ok", f"Claude Desktop에 등록됨", path)
    claude = shutil.which("claude")
    if claude:
        try:
            out = subprocess.run([claude, "mcp", "get", NAME], capture_output=True, text=True, timeout=30,
                                 encoding="utf-8", errors="replace")
            if out.returncode == 0:
                report("ok", "Claude Code에 등록됨")
            else:
                report("info", "Claude Code에는 미등록", f"claude mcp add {NAME} --scope user -- \"{venv_python()}\" \"{SERVER}\"")
        except Exception:
            report("info", "Claude Code 등록 여부를 확인하지 못함")


def check_server_imports() -> None:
    """Import the server module the way Claude starts it, without running it."""
    try:
        code = "import asyncio, server; print(len(asyncio.run(server.mcp.list_tools())))"
        out = subprocess.run([venv_python(), "-c", code],
                             cwd=HERE, capture_output=True, text=True, timeout=120, encoding="utf-8", errors="replace")
    except Exception as exc:
        report("fail", "server.py를 불러오지 못함", str(exc))
        return
    if out.returncode == 0:
        tools = out.stdout.strip().splitlines()[-1] if out.stdout.strip() else "?"
        report("ok", "server.py 불러오기", f"도구 {tools}개" if tools.isdigit() else "")
    else:
        report("fail", "server.py를 불러오지 못함", (out.stderr or out.stdout).strip().splitlines()[-1] if (out.stderr or out.stdout) else "")


def check() -> int:
    print("solidworks-mcp 설치 점검\n")
    check_python()
    check_packages()
    check_server_imports()
    check_solidworks()
    check_pdf()
    check_registration()
    fails = sum(1 for s, *_ in results if s == "fail")
    warns = sum(1 for s, *_ in results if s == "warn")
    print(f"\n결과: 실패 {fails}, 주의 {warns}")
    return 1 if fails else 0


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")  # never die on a console that cannot show a character
        except Exception:
            pass
    if "--register" in args:
        return register()
    if "--unregister" in args:
        return unregister()
    return check()


if __name__ == "__main__":
    sys.exit(main())
