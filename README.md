# solidworks-mcp — SOLIDWORKS API를 Claude에서 쓰는 MCP 서버

SOLIDWORKS의 공식 COM API를 호출해서 Claude가 모델링, 치수 수정, 도면 생성, 내보내기를 할 수 있게 해 줍니다.
SOLIDWORKS 2025 SP2 기준으로 작성했습니다.

- 모든 도구는 **mm, 도(°)** 단위를 씁니다. API 내부 단위인 m/rad 변환은 서버가 합니다.
- 한글 UI에서는 평면 이름이 `정면/윗면/우측면`, 피처 이름이 `보스-돌출1`처럼 현지화되어 있습니다. 기본 평면은 한글·영어 이름으로 찾고, 없으면 원점을 지나는 평면의 방향으로 찾으니 `plane="top"`처럼 영어로 지정하면 됩니다.

## 도구

| 분류 | 도구 |
|---|---|
| 문서 | `status`, `list_documents`, `new_document`, `open_document`, `import_file`(STEP·IGES·Parasolid 등 가져오기), `activate_document`, `save`, `export`, `close_document` |
| 확인 | `list_features`, `list_dimensions`, `model_info`(질량·부피·무게중심·바운딩박스·재질), `screenshot`(뷰 방향 지정 가능) |
| 수정 | `set_dimension`, `rebuild`, `undo`, `delete_feature`, `suppress_feature`, `list_materials`, `set_material`, `custom_properties` |
| 파트 모델링 | `sketch`(선·중심선·사각형·원·호·다각형·슬롯·점·타원·스플라인·글자, 평면 또는 면 위, 설계 의도 치수·대칭 구속 자동), `extrude`(보스/컷, 블라인드·관통·중간평면·면/평면까지·면에서 오프셋·구배·얇은 피처), `sketch_corners`(스케치 모서리 필렛·모따기), `revolve`, `fillet`, `chamfer`, `shell`, `draft`(중립면 구배), `dome`, `sketch_from_face`(면 윤곽 요소 변환·오프셋), `list_bodies`, `combine_bodies`(멀티바디 더하기·빼기·공통), `move_body`(바디 이동·회전·복사) |
| 구멍·각인 | `hole`(Hole Wizard: ISO M2~M20 관통·탭·카운터보어·카운터싱크, 여러 위치 한 번에, 나사산 표시), `cosmetic_thread`(축·구멍에 나사산 표시), `engrave`(글자 각인·양각) |
| 스윕·로프트 | `sketch3d`(3D 경로), `sweep`(단면 또는 원형 파이프·튜브, 컷 가능), `loft`(가이드 커브, 컷 가능), `rib` |
| 용접 구조물 | `weldment_profiles`(설치된 형강 종류·크기), `structural_member`(스케치 선을 따라 각관·파이프·채널 등, 모서리 마이터/맞대기), `trim_members`(부재 트림), `end_cap`(엔드캡), `gusset`(삼각·다각 거싯), `cut_list`(컷 리스트: 수량·길이·각도) |
| 판금 | `sheet_metal_base_flange`, `edge_flange`(모서리 플랜지, 길이·각도 지정), `miter_flange`(이어진 모서리를 따라 도는 플랜지), `hem`(닫힘·열림·물방울·롤·이중 헴), `flatten`(펼치기/접기), `export_flat_pattern`(전개도 DXF, 벤드 라인 포함) |
| 패턴·참조 | `linear_pattern`, `circular_pattern`, `mirror`, `offset_plane`, `angle_plane`(축을 지나는 각도 평면), `reference_axis` |
| 설계 변수 | `list_equations`, `set_equation`(전역 변수·치수 연결), `delete_equation`, `configurations` |
| 검사 | `measure`(거리·각도·길이·면적·지름), `check_interference`(어셈블리 간섭), `bom`(부품 수량표) |
| 기타 | `set_color`, `face_color`(면 단위 색상), `export_folder`(폴더 일괄 STEP/STL 등 변환) |
| 어셈블리 | `insert_component`, `list_components`, `add_mate`(일치·동심·거리·각도 등), `list_mates`, `set_component`(고정/유동·숨기기·억제·이동), `component_pattern`(부품 직선·원형 패턴), `exploded_view`(분해 단계), `replace_component`(부품 교체) |
| 도면 | `make_drawing`(3도 투상, 모델 치수·Hole Wizard 구멍 표기 자동), `add_hole_callouts`, `list_views`, `drawing_view`(등각 등 모델 뷰, 분해 상태 표시 가능), `section_view`(단면도), `detail_view`(상세도), `set_drawing_scale`, `drawing_note`, `center_marks`, `bom_table`, `auto_balloons`, `hole_table`(구멍 표), `title_block`(시트 형식 적용·표제란 속성 입력), `set_tolerance`(대칭·양방향·한계·맞춤 공차), `add_dimension`(뷰의 모서리·꼭짓점·구멍 중심에 치수), `projected_view`(투상도 추가), `auxiliary_view`(보조 투상도), `broken_out_section`(부분 단면도), `drawing_sheets`(시트 추가·전환), `datum`, `geometric_tolerance`(기하공차 틀), `surface_finish`(표면 거칠기), `revision_table`(개정 표), `cut_list_table`(용접 컷 리스트 표) → `export`로 PDF/DXF(모든 시트) |
| 만능 | `run_api_script`(SOLIDWORKS API를 Python으로 직접 호출), `run_macro`(.swp 실행) |

모서리나 면은 **그 위의 점 좌표(mm)**로 지정합니다. 예를 들어 50×30×10 박스의 위쪽 앞 모서리는 `[0, 10, 15]`입니다.
파트와 어셈블리 모두 바디를 직접 검색하므로 화면에서 가려진 모서리도 찾습니다(0.5 mm 이내). 어셈블리에서는 어셈블리 좌표로 지정합니다.
단면도·상세도는 부모 뷰에서 **모델 좌표**로 위치를 지정하고(예: `section_view([35, 5, 0])`), 새 뷰는 시트의 빈 자리에 자동으로 놓입니다.
재질 이름은 라이브러리 이름과 정확히 같아야 하므로, 모르면 `list_materials("6061")`처럼 먼저 검색하세요.
`model_info`의 바운딩박스와 `screenshot`으로 Claude가 좌표를 잡습니다.
전용 도구가 없는 기능(판금 엣지 플랜지, 구배, 서피스, 용접 구조물 등)은 `run_api_script`로 처리합니다.
`hole`은 SOLIDWORKS Hole Wizard로 만들어 도면 구멍 표기(callout)와 나사산 표시가 붙습니다. 인자는 녹화한 매크로(`docs/recordings/`)에서 가져왔습니다. Hole Wizard가 실패하거나 `face_point` 없이 부르면 내장 ISO 표 기준의 스케치+컷으로 대신 만들고, 결과의 `method`로 알려 줍니다.
Windows에서 "Beta: 세계 언어 지원을 위해 Unicode UTF-8 사용"이 켜져 있으면 SOLIDWORKS의 PDF 내보내기가 한글이 든 글자(구멍 표기의 "관통", 주석 등)를 통째로 빠뜨립니다. 이때 도면 PDF는 자동으로 "Microsoft Print to PDF" 프린터로 인쇄해 만듭니다(`export`의 `pdf_method`로 바꿀 수 있음).

도구 하나가 90초(`SW_TIMEOUT` 환경변수로 변경) 안에 끝나지 않으면 오류로 돌아오고, 그 호출이 끝날 때까지 다음 호출은 바로 거절됩니다. 대개 SOLIDWORKS에 대화상자가 떠 있는 경우이니 닫고 다시 시도하세요.

## 설치 (SOLIDWORKS가 깔린 Windows PC)

업데이트는 `cd C:\mcp\solidworks-mcp` 후 `git pull`, 그리고 Claude Desktop 재시작입니다.

Python 3.11 이상이 필요합니다.

```powershell
git clone https://github.com/Saqnes/solidworks-mcp C:\mcp\solidworks-mcp
cd C:\mcp\solidworks-mcp
python -m venv .venv
.venv\Scripts\pip install -e .
```

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "solidworks": {
      "command": "C:\\mcp\\solidworks-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\mcp\\solidworks-mcp\\server.py"]
    }
  }
}
```

### Claude Code

```powershell
claude mcp add solidworks -- C:\mcp\solidworks-mcp\.venv\Scripts\python.exe C:\mcp\solidworks-mcp\server.py
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
