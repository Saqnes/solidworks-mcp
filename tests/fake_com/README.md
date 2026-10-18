# Fake-COM smoke tests

These run without Windows or SOLIDWORKS: each script swaps pywin32 and the
SOLIDWORKS objects for small Python fakes, then drives the real MCP tools
through an MCP client. They check call shapes (argument counts, selection
order, retries, error handling), not real geometry.

```powershell
.venv\Scripts\python -I tests\fake_com\run_all.py
```

(Also runs on Linux or macOS with just `mcp` and `pillow` installed;
pywin32 is not needed.)

## ERR lines that are expected

| Script | Expected ERR | Why |
|---|---|---|
| smoke_sketch | fillet `No edge within 0.5 mm of [50, 4, -20]` | deliberate wrong point |
| smoke_sketch | save `... is open in SOLIDWORKS as ...` | refuses to overwrite an open file |
| smoke_sketch | set_material `'6061' is not in the library` | suggests close names |
| sw_smoke3 | configurations `No configuration '없음'` | unknown configuration |
| sw_smoke3 | set_equation `rejected '"Width" = 130'; kept ...` | refused edit keeps the old equation |
| sw_smoke4 | hole `Unknown size 'M7'` | size table check |
| sw_smoke4 | sheet_metal_base_flange `no attribute 'GetPartBox'` | gap in that fake |
| sw_smoke7 | sweep `none of [...] accepted` | that fake has no sweep API |

Anything else marked ERR is a regression. The fakes grew test by test, so
some are rough; `sw_smoke3.py` is the base the later scripts reuse.
