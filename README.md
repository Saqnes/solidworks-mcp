# solidworks-mcp — SOLIDWORKS API를 Claude에서 쓰는 MCP 서버

SOLIDWORKS의 공식 COM API를 호출해서 Claude가 모델링, 치수 수정, 도면 생성, 내보내기를 할 수 있게 해 줍니다.
SOLIDWORKS 2025 SP2 기준으로 작성했습니다.

- 모든 도구는 **mm, 도(°)** 단위를 씁니다. API 내부 단위인 m/rad 변환은 서버가 합니다.
- 한글 UI에서는 평면 이름이 `정면/윗면/우측면`, 피처 이름이 `보스-돌출1`처럼 현지화되어 있습니다. 기본 평면은 한글·영어 이름으로 찾고, 없으면 원점을 지나는 평면의 방향으로 찾으니 `plane="top"`처럼 영어로 지정하면 됩니다.

## 도구

| 분류 | 도구 |
|---|---|
| 문서 | `status`(버전·열린 문서 수·SOLIDWORKS 메모리/핸들 사용량 경고), `list_documents`, `new_document`, `open_document`, `import_file`(STEP·IGES·Parasolid 등 가져오기), `activate_document`, `save`, `export`, `close_document` |
| 확인 | `list_features`, `list_faces`·`list_edges`(면·모서리의 안쪽 점·법선·축·반경·만든 피처: 점으로 고르는 도구에 바로 사용), `list_dimensions`, `model_info`(질량·부피·무게중심·바운딩박스·재질), `screenshot`(뷰 방향, 크기 small/medium/large) |
| 수정 | `set_dimension`, `rebuild`(재빌드 오류 목록. 다른 도구도 결과에 `rebuild_errors`로 알려 줌), `undo`, `delete_feature`, `suppress_feature`, `rename_feature`(피처·스케치·메이트 이름 바꾸기), `list_materials`, `set_material`, `custom_properties`, `document_units`(길이 단위·소수 자릿수) |
| 파트 모델링 | `sketch`(선·중심선·사각형·원·호·다각형·슬롯·점·타원·스플라인·글자, 평면 또는 면 위, 설계 의도 치수·대칭 구속 자동), `extrude`(보스/컷, 블라인드·관통·중간평면·면/평면까지·면에서 오프셋·구배·얇은 피처), `sketch_corners`(스케치 모서리 필렛·모따기), `edit_sketch`(기존 스케치에 요소 추가), `revolve`, `fillet`, `variable_fillet`(변수 반경), `face_fillet`(면 필렛), `chamfer`, `shell`, `draft`(중립면 구배), `dome`, `sketch_from_face`(면 윤곽 요소 변환·오프셋), `list_bodies`, `combine_bodies`(멀티바디 더하기·빼기·공통), `move_body`(바디 이동·회전·복사) |
| 구멍·각인 | `hole`(Hole Wizard: ISO M2~M20 관통·탭·카운터보어·카운터싱크, 여러 위치 한 번에, 나사산 표시. 카운터싱크 M2·M2.5는 ISO 10642에 없는 크기라 스케치 컷으로 만듦), `cosmetic_thread`(축·구멍에 나사산 표시), `engrave`(글자 각인·양각) |
| 스윕·로프트 | `sketch3d`(3D 경로), `helix`(나선: 스프링·코일 경로, 피치·회전수·테이퍼), `sweep`(단면 또는 원형 파이프·튜브, 컷 가능), `loft`(가이드 커브, 컷 가능), `rib` |
| 용접 구조물 | `weldment_profiles`(설치된 형강 종류·크기), `structural_member`(스케치 선을 따라 각관·파이프·채널 등, 모서리 마이터/맞대기), `trim_members`(부재 트림), `end_cap`(엔드캡), `gusset`(삼각·다각 거싯), `cut_list`(컷 리스트: 수량·길이·각도) |
| 판금 | `sheet_metal_base_flange`, `edge_flange`(모서리 플랜지, 길이·각도 지정), `miter_flange`(이어진 모서리를 따라 도는 플랜지), `hem`(닫힘·열림·물방울·롤·이중 헴), `sketched_bend`(판 위의 선을 따라 굽힘), `insert_bends`(가져온 STEP처럼 두께 일정한 솔리드를 판금으로: 굽힘 반경 지정, 전개 크기 확인), `flatten`(펼치기/접기), `export_flat_pattern`(전개도 DXF, 벤드 라인 포함) |
| 패턴·참조 | `linear_pattern`, `circular_pattern`, `sketch_pattern`(스케치 점 위치마다 복사), `mirror`, `mirror_part`(반대손 부품을 새 파트로), `offset_plane`, `angle_plane`(축을 지나는 각도 평면), `reference_axis`, `coordinate_system`(위치·회전 지정 좌표계), `reference_point`(면·원형 모서리 중심점) |
| 설계 변수 | `list_equations`, `set_equation`(전역 변수·치수 연결), `delete_equation`, `configurations` |
| 검사 | `measure`(거리·각도·길이·면적·지름), `check_interference`(어셈블리 간섭), `bom`(부품 수량표) |
| 기타 | `pack_and_go`(참조 파일·도면까지 한 폴더로 복사, 접두사/접미사), `set_color`, `face_color`(면 단위 색상), `export_folder`(폴더 일괄 STEP/STL 등 변환) |
| 어셈블리 | `insert_component`(원점 위치·회전 지정), `list_components`(구성 포함), `add_mate`(일치·동심·평행·직각·접선·거리·각도, 대칭·폭·기어·잠금·슬롯·힌지·나사·랙피니언, 거리/각도 한계 메이트), `list_mates`, `set_component`(고정/유동·숨기기·억제·이동·회전·구성 전환), `component_pattern`(부품 직선·원형 패턴), `exploded_view`(분해 단계), `replace_component`(부품 교체), `mirror_components`(평면 대칭 인스턴스) |
| 도면 | `make_drawing`(3도 투상, 모델 치수·Hole Wizard 구멍 표기·중심선 자동), `view_style`(은선 표시/제거·음영·접선 모서리), `centerlines`(원통 측면 중심선), `add_hole_callouts`, `list_views`, `drawing_view`(등각 등 모델 뷰, 분해 상태 표시 가능), `section_view`(단면도), `detail_view`(상세도), `set_drawing_scale`, `drawing_note`, `center_marks`, `bom_table`, `auto_balloons`, `hole_table`(구멍 표), `title_block`(시트 형식 적용·표제란 속성 입력), `set_tolerance`(대칭·양방향·한계·맞춤 공차), `add_dimension`(뷰의 모서리·꼭짓점·구멍 중심에 치수), `dimension_holes`(모든 구멍 위치를 좌표 치수로), `flat_pattern_view`(판금 전개도 뷰 + 외곽 치수), `projected_view`(투상도 추가), `auxiliary_view`(보조 투상도), `broken_out_section`(부분 단면도), `drawing_sheets`(시트 추가·전환), `datum`, `geometric_tolerance`(기하공차 틀), `surface_finish`(표면 거칠기), `revision_table`(개정 표), `cut_list_table`(용접 컷 리스트 표), `weld_symbol`(ISO 용접 기호), `ordinate_dimensions`(좌표 치수), `break_view`(단축/파단 뷰) → `export`로 PDF/DXF(모든 시트) |
| 만능 | `run_steps`(여러 도구를 한 번에 순서대로 실행, 앞 단계 결과를 `"$N.키"`로 전달: 도구를 하나씩 부르는 것보다 훨씬 빠름), `run_api_script`(SOLIDWORKS API를 Python으로 직접 호출), `run_macro`(.swp 실행) |

모서리나 면은 **그 위의 점 좌표(mm)**로 지정합니다. 예를 들어 50×30×10 박스의 위쪽 앞 모서리는 `[0, 10, 15]`입니다.
파트와 어셈블리 모두 바디를 직접 검색하므로 화면에서 가려진 모서리도 찾습니다(0.5 mm 이내). 어셈블리에서는 어셈블리 좌표로 지정합니다.
단면도·상세도는 부모 뷰에서 **모델 좌표**로 위치를 지정하고(예: `section_view([35, 5, 0])`), 새 뷰는 시트의 빈 자리에 자동으로 놓입니다.
재질 이름은 라이브러리 이름과 정확히 같아야 하므로, 모르면 `list_materials("6061")`처럼 먼저 검색하세요.
Claude는 `list_faces`/`list_edges`(면 안쪽 점, 모서리 중점, 법선·축·반경)로 좌표를 잡고 `screenshot`으로 확인합니다.
전용 도구가 없는 기능(서피스 모델링 등)은 `run_api_script`로 처리합니다.
`hole`은 SOLIDWORKS Hole Wizard로 만들어 도면 구멍 표기(callout)와 나사산 표시가 붙습니다. 인자는 녹화한 매크로(`docs/recordings/`)에서 가져왔습니다. Hole Wizard가 실패하거나 `face_point` 없이 부르면 내장 ISO 표 기준의 스케치+컷으로 대신 만들고, 결과의 `method`로 알려 줍니다.
Windows에서 "Beta: 세계 언어 지원을 위해 Unicode UTF-8 사용"이 켜져 있으면 SOLIDWORKS의 PDF 내보내기가 한글이 든 글자(구멍 표기의 "관통", 주석 등)를 통째로 빠뜨립니다. 이때 도면 PDF는 자동으로 "Microsoft Print to PDF" 프린터로 인쇄해 만듭니다(`export`의 `pdf_method`로 바꿀 수 있음).

도구 하나가 90초(`SW_TIMEOUT` 환경변수로 변경) 안에 끝나지 않으면 오류로 돌아오고, 그 호출이 끝날 때까지 다음 호출은 바로 거절됩니다. 대개 SOLIDWORKS에 대화상자가 떠 있는 경우이니 닫고 다시 시도하세요.

## 설치 (SOLIDWORKS가 깔린 Windows PC)

두 방법 중 **하나만** 쓰세요. 둘 다 하면 Claude에 도구가 두 번 보입니다. 어느 쪽이든 SOLIDWORKS(2025 SP2에서 확인)와 Claude Desktop(또는 Claude Code)이 있어야 합니다.

### 방법 A: Claude Desktop 확장 파일(.mcpb) — 가장 간단, Python 필요 없음

1. `solidworks-mcp-<버전>.mcpb` 파일을 받습니다. 배포한 사람에게 받거나, 저장소에서 `.venv\Scripts\python scripts\build_mcpb.py`로 만듭니다(`dist\` 폴더에 생김).
2. 파일을 더블클릭하거나, Claude Desktop 설정의 **확장(Extensions)** 화면에 끌어다 놓고 **설치**를 누릅니다.
3. 끝입니다. 처음 켤 때 Claude Desktop이 Python과 패키지를 직접 내려받느라 1~2분 걸릴 수 있고, 그다음부터는 1초 안에 뜹니다. 확장 설정에서 응답 대기 시간과 SOLIDWORKS 버전(ProgID)을 바꿀 수 있습니다.

업데이트는 새 버전 `.mcpb`를 같은 방법으로 설치하면 되고, 제거는 확장 화면에서 합니다.

### 방법 B: 설치 스크립트 — Claude Code에서도 쓰거나 코드를 고칠 때

1. **받기**: 아래 중 하나. 폴더 위치는 어디든 됩니다(한글 경로 가능).
   - git: `git clone https://github.com/Saqnes/solidworks-mcp C:\mcp\solidworks-mcp`
   - git이 없으면: GitHub 페이지의 **Code > Download ZIP**을 받아 원하는 폴더에 압축을 풉니다.
2. **설치**: 그 폴더의 `install.cmd`를 더블클릭합니다. 30초~1분쯤 걸리며 아래를 알아서 합니다.
   - Python 3.11+ 찾기. 없으면 **winget으로 Python 3.12를 설치할지 물어보고**(Enter = 예) 설치합니다.
   - `.venv` 가상환경 → 패키지 설치(mcp, pywin32, pillow)
   - Claude Desktop 설정(`claude_desktop_config.json`)에 `solidworks` 등록. 다른 서버·설정은 그대로 두고, 바꾸기 전에 `.bak` 백업을 만듭니다. Microsoft Store판 Claude도 지원합니다.
   - `claude` 명령(Claude Code)이 있으면 Claude Code에도 등록
   - 설치 점검: Python, 패키지, SOLIDWORKS COM 등록, 기본 템플릿, PDF 프린터, 등록 상태를 `[ OK ]` / `[주의]` / `[실패]`로 보여 줍니다.
3. **Claude Desktop을 완전히 종료**(작업 표시줄 트레이 아이콘 > 종료)했다가 다시 켭니다. 창만 닫으면 서버가 다시 읽히지 않습니다.

| 할 일 | 방법 |
|---|---|
| 업데이트 | `update.cmd` 더블클릭(`git pull` 후 재설치·점검). ZIP으로 받았다면 새 ZIP을 같은 폴더에 덮어쓰고 `install.cmd` |
| 점검만 다시 | 폴더에서 `.venv\Scripts\python doctor.py` |
| 제거 | `.venv\Scripts\python doctor.py --unregister`(Claude Desktop), `claude mcp remove solidworks --scope user`(Claude Code), 그다음 폴더 삭제 |
| 옵션 | `powershell -ExecutionPolicy Bypass -File install.ps1 -Yes -NoDesktop -NoClaudeCode -Python C:\경로\python.exe` (`-Yes`: 묻지 않고 Python 설치) |
| 다른 사람에게 줄 `.mcpb` 만들기 | `.venv\Scripts\python scripts\build_mcpb.py` (Node.js 필요, `dist\`에 생김) |

SOLIDWORKS가 실행 중이면 그 창에 연결하고, 실행 중이 아니면 도구를 처음 쓸 때 새로 띄웁니다.
`도구 > 옵션 > 시스템 옵션 > 기본 템플릿`에 파트/어셈블리/도면 템플릿이 지정되어 있어야 `new_document`와 `make_drawing`이 동작합니다(점검에서 확인).

### 문제 해결

| 증상 | 해결 |
|---|---|
| Claude에 solidworks 도구가 안 보임 | Claude Desktop을 트레이에서 **종료** 후 재실행. 그래도 없으면 `doctor.py`로 등록 경로 확인 |
| "Python 3.11 이상이 없습니다" | `install.cmd`가 winget 설치를 제안합니다. winget이 없으면 python.org에서 설치(PATH 체크). `python`을 치면 Microsoft Store가 열리는 PC는 Store용 바로가기만 있는 상태입니다. 방법 A(.mcpb)는 Python이 필요 없습니다 |
| 확장(.mcpb) 설치 후 첫 실행이 오래 걸림/실패 | 처음 한 번은 Python과 패키지(약 30 MB)를 내려받습니다. 회사 프록시·방화벽이 막으면 실패하니 방법 B를 쓰세요 |
| 패키지 설치 실패 | 인터넷/회사 프록시 확인. 프록시면 `set HTTPS_PROXY=http://프록시:포트` 후 `install.cmd` |
| `install.cmd`가 바로 닫힘/차단 | 파일 속성에서 *차단 해제* 체크, 또는 PowerShell에서 위 옵션 줄로 실행 |
| 기본 템플릿 [실패] | SOLIDWORKS 옵션에서 기본 템플릿 지정 |
| SOLIDWORKS가 "Windows 리소스가 매우 부족합니다" 창을 띄우고 멈춤 | 문서를 아주 많이 만들고 닫으면 생깁니다(시험에서 약 140개, 메모리 9 GB). 작업을 저장하고 SOLIDWORKS를 재시작하세요. `status`가 사용량이 많을 때 미리 경고합니다 |
| 도면 PDF에 한글이 빠짐 | Windows "UTF-8 사용(베타)" 설정 때문입니다. 서버가 자동으로 "Microsoft Print to PDF"로 인쇄하므로 그 프린터가 켜져 있어야 합니다(점검에서 확인) |

<details>
<summary>수동 설정(설치 스크립트를 쓰지 않을 때)</summary>

```powershell
cd C:\mcp\solidworks-mcp
python -m venv .venv
.venv\Scripts\pip install -e .
```

`%APPDATA%\Claude\claude_desktop_config.json`의 `mcpServers`에 추가:

```json
"solidworks": {
  "command": "C:\\mcp\\solidworks-mcp\\.venv\\Scripts\\python.exe",
  "args": ["C:\\mcp\\solidworks-mcp\\server.py"]
}
```

Claude Code:

```powershell
claude mcp add solidworks --scope user -- C:\mcp\solidworks-mcp\.venv\Scripts\python.exe C:\mcp\solidworks-mcp\server.py
```
</details>

## 빠르게 쓰는 요령

- 여러 단계 작업은 "**run_steps로 한 번에** 해 줘"라고 하면 Claude가 도구를 하나씩 부르지 않고 한 호출로 처리해 훨씬 빠릅니다.
- 확인용 그림은 `screenshot`의 `size="small"`이면 충분한 경우가 많습니다(기본 medium, 세부는 large).
- 모서리·면 좌표는 `list_faces`/`list_edges`가 바로 주므로, 바운딩박스로 추측하게 하지 않아도 됩니다.
- 서버는 도구를 실행하는 동안 SOLIDWORKS에 `CommandInProgress`를 알려 같은 모델링 작업이 약 30% 빨라집니다.
- 긴 세션 뒤 느려지면 `status`의 메모리·핸들 경고를 보고 SOLIDWORKS를 재시작하세요.

## 예시 요청

- "100×60×8 mm 알루미늄 6061 브래킷 만들고, 모서리 R5 필렛, 지름 6.6 관통 구멍 4개 넣고 STEP이랑 도면 PDF로 뽑아 줘."
- "열려 있는 파트에서 두께 치수 찾아서 12 mm로 바꾸고 질량 알려 줘."
- "이 폴더의 sldprt 전부 STEP으로 변환해 줘."
- "어셈블리에서 핀을 브래킷 구멍에 동심으로 넣고, 슬라이더는 레일 사이 폭 메이트에 0~30 mm 거리 한계로 움직이게 해 줘."

## 개발자용 테스트

| 테스트 | 명령 | 내용 |
|---|---|---|
| 가짜 COM | `.venv\Scripts\python -I tests\fake_com\run_all.py` | SOLIDWORKS 없이 도구 호출 흐름만 확인 |
| 실제 SOLIDWORKS | `.venv\Scripts\python tests\live\run_live.py [parts holes sheetmetal weldment assembly drawing]` | 임시 폴더에 작은 모델을 만들어 부피·길이·개수·위치를 손으로 계산한 값과 비교(약 2분). 이미 열린 문서는 건드리지 않고, `SWMCP_KEEP=1`이면 결과 파일을 남김 |

## 주의

- `run_api_script`는 Claude가 쓴 Python 코드를 이 PC에서 그대로 실행합니다. Claude Desktop/Code의 도구 승인 창에서 코드를 확인하고 허용하세요.
- `close_document`는 저장하지 않은 변경이 있으면 거부합니다(`discard_changes=true`일 때만 버림).
- 실행 중인 SOLIDWORKS가 2개 이상이면 먼저 연결된 쪽을 씁니다. 특정 버전을 지정하려면 환경변수 `SW_PROGID=SldWorks.Application.33`(2025)을 설정하세요.
