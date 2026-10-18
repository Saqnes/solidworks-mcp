# Handoff: Hole Wizard / base flange 재작성 (2026-10-18 녹화 기준)

녹화 원본: `rec_holewizard.swp`, `rec_holes2.swp`, `rec_baseflange.swp` (이 폴더). SOLIDWORKS 2025 (rev 33.2), 한글판. 녹화 코드는 VBA 편집기에서 복사한 원문.

## 1. `hole` → Hole Wizard (`FeatureManager.HoleWizard5`)

녹화 원본: `rec_holewizard.swp`(M6 CBORE 관통, M4 탭 관통), `rec_holes2.swp`(M4 막힌 탭 cosmetic on/off, M6 clearance, M5 CSK).

### 공통 절차 (녹화 그대로)
```vb
' 1) 면 선택: 구멍 위치 위에서 -Y 방향 ray (모델 좌표 x, 윗면 높이, z)
boolstatus = Part.Extension.SelectByRay(x, 0.01, z, 0, -1, 0, r, 2, False, 0, 0)
' 2) HoleWizard5(...)  ← 아래 표
' 3) 위치 스케치 편집: 자동 생성 점 삭제 → 원하는 위치에 점 → 종료
Set sk = f.GetFirstSubFeature : sk.Select2 False, 0 : Part.EditSketch
pts = sk.GetSpecificFeature2().GetSketchPoints2()
'   각 점: SelectionManager.AddSelectionListObject pt, Nothing : Part.EditDelete
Part.SketchManager.CreatePoint(x, -z, 0)   ' 윗면 스케치: sketch y = -model z (녹화로 확인)
Part.SketchManager.InsertSketch True
```

### 녹화된 HoleWizard5 호출 6종
```vb
' A. M6 카운터보어 ISO 4762, 관통
HoleWizard5(0, 8, 139, "M6", 1, 0.0066, 0.01, -1,  0.011, 0.0064, 0, 1, 0, 0.01105, 1.5707963267949, 0, 0, 0, 0, 0, "", False, True, True, True, True, False)
' B. M4 탭, 관통, cosmetic thread
HoleWizard5(4, 8, 147, "M4", 1, 0.0033, 0.01, True, 0.009625, 0.00405, 1.5707963267949, 0, 0, 0, 1, 1, 0, 0, -1, -1, "", False, True, True, True, True, False)
' C. M4 탭, 막힌 구멍 8 / 나사산 6, cosmetic thread
HoleWizard5(4, 8, 147, "M4", 0, 0.0033, 0.008, True, 0.006, 0.00405, 1.5707963267949, 0, 0, 2.05948851735331, 1, 0, 0, 0, -1, -1, "", False, True, True, True, True, False)
' D. M4 탭, 막힌 구멍 8 / 나사산 6, 탭 드릴 지름만 (cosmetic 없음)
HoleWizard5(4, 8, 147, "M4", 0, 0.0033, 0.008, True, 0.006, 0.00405, 1.5707963267949, 0, 0, 2.05948851735331, 0, 0, 0, 0, -1, -1, "", False, True, True, True, True, False)
' E. M6 나사 여유 구멍(clearance, 맞춤 보통), 관통, near-side csk 기본 ON
HoleWizard5(2, 8, 144, "M6", 1, 0.0066, 0.01, -1,  1, 0, 0.00665, 1.5707963267949, 0, 0, 0, -1, -1, -1, -1, -1, "", False, True, True, True, True, False)
' F. M5 카운터싱크 ISO 10642, 관통
HoleWizard5(1, 8, 140, "M5", 1, 0.0055, 0.01, -1,  0.0112, 1.5707963267949, 0, 1, 0, 0, 0, 0, 0, -1, -1, -1, "", False, True, True, True, True, False)
```

### 인자 해석 (C/D diff 등으로 확인한 것은 "확인", 나머지는 "추정")
| # | 이름 | 값 | 판정 |
|---|---|---|---|
| 1 | GenericHoleType | 0=counterbore, 1=countersink, 2=hole(clearance), 4=tap | 확인 |
| 2 | StandardIndex | 8 = ISO | 확인 |
| 3 | FastenerTypeIndex | 139=ISO 4762 소켓 머리(CBORE), 140=ISO 10642 소켓 접시(CSK), 144=나사 여유값(clearance), 147=탭 구멍 | 확인 |
| 5 | EndType | 0=블라인드, 1=관통 | 확인 |
| 6 | Diameter | 드릴(탭 드릴) 지름 m | 확인 |
| 7 | Depth | 블라인드 깊이 m (관통이면 무시) | 확인 |
| 8 | Length | -1 (VB True = -1) | |
| CBORE V1/V2 | 보어 지름 0.011 / 깊이 0.0064 | 확인 |
| CBORE V6/V7 | near-side csk 지름 0.01105 / 각도 π/2 | 추정 |
| TAP V1 | 나사산 깊이 (C/D: 0.006 = 입력한 6 mm) | 확인 |
| TAP V2/V3 | near-side csk 지름 0.00405 / 각도 π/2 | 추정 |
| TAP V6 | 드릴 끝 각도 2.0594885 rad = 118° (블라인드일 때만) | 확인 |
| TAP V7 | 1=cosmetic thread, 0=탭 드릴 지름만 (C vs D 유일한 차이) | 확인 |
| TAP V8 | 1=나사산 관통, 0=나사산 블라인드 (B vs C) | 추정 |
| CLEAR V1 | 1 = 맞춤 "보통"으로 추정 | 추정 |
| CLEAR V3/V4 | near-side csk 지름 0.00665 / 각도 π/2 | 추정 |
| CSK V1/V2 | 카운터싱크 지름 0.0112 / 각도 π/2 | 확인 |
| 21~27 | ThreadClass "", RevDir False, FeatureScope True, AutoSelect True, AsmFeatureScope True, AutoSelectComponents True, PropagateFeatureToParts False | 녹화값 그대로 |

### 구현 메모
- 기존 sketch+cut 경로는 `wizard=False` 폴백으로 남기고, 결과에 `"method": "hole_wizard" | "sketch_cut"` 반환.
- 첫 위치용 면 선택은 기존 `select_at(doc,"FACE",face_point)` 재사용 가능. 이후 위치 스케치에서 자동 점을 지우고 `positions` 전부를 `CreatePoint` (위치 스케치 종료 전 `ActiveSketch` 확인).
- `tapped`에 `cosmetic_thread: bool = True`(V7), `thread_depth`(V1) 노출. 블라인드면 V6=118°.
- clearance는 GUI 기본으로 near-side csk가 켜져 있었음 → API에서는 끄는 값이 필요할 수 있음(재녹화 또는 V3/V4=0 실험).
- 크기별 FastenerTypeIndex는 크기와 무관(규격·종류별)이라 위 4개로 충분할 가능성이 큼. 다른 크기로 1회 검증 권장.
- 테스트: 도면 Hole Callout(`M6` CBORE, `M4x0.7 ↧6` 등), 탭 C에 cosmetic thread 존재, D에는 없음.

## 2. `sheet_metal_base_flange` → feature definition 방식

### 녹화된 호출 (L 스케치, 두께 1.5, R1, 한 방향 블라인드 40)
```vb
Set ba = Part.FeatureManager.CreateCustomBendAllowance()
ba.KFactor = 0.5
Set d = swFeatMgr.CreateDefinition(swFeatureNameID_e.swFmBaseFlange)
d.Initialize False, True, ba, True, 1, True, 0.5, 0.0001, 0.0001
d.BendRadius = 0.001
d.D1EndConditionDistance = 0.04
d.D1EndConditionType = 1
d.D1ReverseOffset = False
d.D2EndConditionDistance = 0.01
d.D2EndConditionType = 1
d.D2ReverseOffset = False
d.OffsetDirections = 1        ' 한 방향
d.ReverseDirection = False
d.ReverseThickness = False
d.Thickness = 0.0015
Set swFeat = swFeatMgr.CreateFeature(d)
```
- 현재 코드(InsertSheetMetalBaseFlange2 + 길이/2 대칭)를 이것으로 교체. 기본은 한 방향, `mid_plane` 옵션으로 대칭 선택 가능하게.
- `swFmBaseFlange` 숫자값은 추측하지 말고 typelib 상수로 얻을 것 (`win32com.client.gencache.EnsureDispatch` 후 `win32com.client.constants.swFmBaseFlange`). `Initialize`는 late binding에서 `_FlagAsMethod` 필요할 수 있음.
- 테스트: 길이 40 → 바운딩박스 해당 축 40, 대칭 아님.

## 3. 그 외 남은 항목
- `sweep(wall=...)`: Ø10 + wall 1 이 외경 Ø12/내경 Ø10으로 생성됨(두께가 바깥쪽). `circle_diameter`를 외경 기준으로 할지 옵션화하거나 docstring에 명시.
- `engrave`: 성공. 끝나면 텍스트 스케치 숨기기 권장.
- `export_flat_pattern`: `ExportToDWG2`의 Alignment/Views 배열 자리에 `nothing()` 대신 `sw.empty()` 검토.
- `flatten`: `SetSuppression2(..., 2, ...)`는 모든 구성 적용 — 의도 확인.
- 녹화 매크로의 한글 피처 이름은 인코딩이 깨져 저장됨 → 이름 기반 선택 대신 객체/좌표 선택 유지.
