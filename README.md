# solidworks-mcp — SOLIDWORKS API를 Claude에서 쓰는 MCP 서버

SOLIDWORKS의 공식 COM API를 호출해서 Claude가 모델링, 치수 수정, 도면 생성, 내보내기를 할 수 있게 해 줍니다.
SOLIDWORKS 2025 SP2 기준으로 작성했습니다.

- 모든 도구는 **mm, 도(°)** 단위를 씁니다. API 내부 단위인 m/rad 변환은 서버가 합니다.
- 한글 UI에서는 평면 이름이 `정면/윗면/우측면`, 피처 이름이 `보스-돌출1`처럼 현지화되어 있습니다. 기본 평면은 이름이 아니라 순서로 찾으니 `plane="front"`처럼 영어로 지정하면 됩니다.

## 도구

| 분류 | 도구 |
|---|---|
| 문서 | `status`, `list_documents`, `new_document`, `open_document`, `activate_document`, `save`, `export`, `close_document` |
| 확인 | `list_features`, `list_dimensions`, `model_info`(질량·부피·무게중심·바운딩박스·재질), `screenshot`(뷰 방향 지정 가능) |
| 수정 | `set_dimension`, `rebuild`, `undo`, `delete_feature`, `suppress_feature`, `list_materials`, `set_material`, `custom_properties` |
| 파트 모델링 | `sketch`(선·중심선·사각형·원·호·다각형·슬롯·점, 평면 또는 면 위, 치수 자동 기입), `extrude`(보스/컷, 블라인드·관통·중간평면·구배), `revolve`, `fillet`, `chamfer` |
| 어셈블리 | `insert_component`, `list_components`, `add_mate`(일치·동심·거리·각도 등) |
| 도면 | `make_drawing`(3도 투상, 모델 치수 가져오기) → `export`로 PDF/DXF |
| 만능 | `run_api_script`(SOLIDWORKS API를 Python으로 직접 호출), `run_macro`(.swp 실행) |

모서리나 면은 **그 위의 점 좌표(mm)**로 지정합니다. 예를 들어 50×30×10 박스의 위쪽 앞 모서리는 `[0, 10, 15]`입니다.
파트에서는 바디를 직접 검색하므로 화면에서 가려진 모서리도 찾습니다(0.5 mm 이내).
재질 이름은 라이브러리 이름과 정확히 같아야 하므로, 모르면 `list_materials("6061")`처럼 먼저 검색하세요.
`model_info`의 바운딩박스와 `screenshot`으로 Claude가 좌표를 잡습니다.
전용 도구가 없는 기능(패턴, 구멍 가공 마법사, 판금, 방정식 등)은 `run_api_script`로 처리합니다.

## 설치 (SOLIDWORKS가 깔린 Windows PC)

Python 3.11 이상이 필요합니다.

```powershell
cd solidworks-mcp
python -m venv .venv
.venv\Scripts\pip install -e .
```

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "solidworks": {
      "command": "C:\\path\\to\\personalproject\\solidworks-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\personalproject\\solidworks-mcp\\server.py"]
    }
  }
}
```

### Claude Code

```powershell
claude mcp add solidworks -- C:\path\to\personalproject\solidworks-mcp\.venv\Scripts\python.exe C:\path\to\personalproject\solidworks-mcp\server.py
```

SOLIDWORKS가 실행 중이면 그 창에 연결하고, 실행 중이 아니면 새로 띄웁니다.
`도구 > 옵션 > 기본 템플릿`에 파트/어셈블리/도면 템플릿이 지정되어 있어야 `new_document`와 `make_drawing`이 동작합니다.

## 예시 요청

- "100×60×8 mm 알루미늄 6061 브래킷 만들고, 모서리 R5 필렛, 지름 6.6 관통 구멍 4개 넣고 STEP이랑 도면 PDF로 뽑아 줘."
- "열려 있는 파트에서 두께 치수 찾아서 12 mm로 바꾸고 질량 알려 줘."
- "이 폴더의 sldprt 전부 STEP으로 변환해 줘."

## 주의

- `run_api_script`는 Claude가 쓴 Python 코드를 이 PC에서 그대로 실행합니다. Claude Desktop/Code의 도구 승인 창에서 코드를 확인하고 허용하세요.
- `close_document`는 저장하지 않은 변경이 있으면 거부합니다(`discard_changes=true`일 때만 버림).
- 실행 중인 SOLIDWORKS가 2개 이상이면 먼저 연결된 쪽을 씁니다. 특정 버전을 지정하려면 환경변수 `SW_PROGID=SldWorks.Application.33`(2025)을 설정하세요.
