# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

본 파일은 UnivRPA 고유 사항만 다룬다. VBA 모듈/프로시저 시그니처 규약(`@desc`, `@depends`, 길이 제한)은 부모 `Project/CLAUDE.md` § 2 참조 — 중복 기재하지 않는다.

## 1. 프로젝트 정체성

UnivRPA는 Excel VBA 기반 범용 RPA 인터프리터. 시트의 표(COMMAND/PARAM1/PARAM2/PARAM3) 행을 명령어로 해석해 Win32 API로 실행한다. 픽셀 비전, 라벨/CALL/RETURN 흐름 제어를 지원. 운영 자산은 `../UnivRPA.xlsb` 단일 워크북.

## 2. 4-Layer 아키텍처

엄격한 단방향 의존:

```
modEngine   (Layer 1) ─ 인터프리터/오케스트레이터  RunRPA() 진입
   └─ modContext (Layer 2) ─ 상태 (Memory[0..99], CallStack, Labels, Loop counters)
   └─ modActions (Layer 3) ─ 행동 (클릭/입력/대기/픽셀)
         └─ modWinAPI (Layer 4) ─ Win32 thin wrapper
```

규칙(`modEngine.bas:8-13`에 명시):

- modEngine은 modWinAPI를 직접 호출하지 않는다 → modActions 경유.
- modEngine은 변수를 직접 조작하지 않는다 → modContext 경유.
- modWinAPI는 Pure 기술 레이어. 상태/루프/판단 로직 추가 금지.

부모 디렉터리 `../UnivRPA.md`에 자동 생성된 의존 그래프, API Reference, 모듈 인벤토리(21 모듈, ~4236 LoC)가 있다. 단 export 시점이 오래됐을 수 있어 코드를 우선한다.

## 3. 비자명한 동작 규칙

- **시트 → 배열 박제**: `RunRPA` 시작 시 시트 전체를 `m_arrMain` 2D Variant 배열로 캐시. 실행 중 시트 재참조 금지. 사용자가 행을 추가/수정해도 다음 실행에서야 반영.
- **파라미터 보간 (`ResolveParam`, dispatch 시점)**:
  - `{n}` — Memory 슬롯(0~99)
  - `{CLIP}` — 클립보드
  - `{PIXEL:x,y}` — 해당 좌표 픽셀 색
- **들여쓰기 시각화**: 명령어 앞 `>` 접두는 중첩 깊이. `modDevTools.FormatRPAScript_Final`이 자동 정렬.
- **인터럽트 키**: `VK_ESCAPE`(`&H1B`)=하드 종료, `VK_PAUSE_RESUME`(`&H13`)=처음부터 재시작. `BetterWait`이 `DoEvents` + 50ms(`DELAY_MICRO`) 단위로 검사.
- **타이밍 상수** (`modActions.bas` Private Const): `DELAY_MICRO=50`, `DELAY_SHORT=100`, `DELAY_MEDIUM=200`, `TIMEOUT_DEFAULT=30000`.
- **색상 표기**: `GetColorHexAt`은 6자리 대문자 hex(예: `"FF0000"`). 비교 시 `UCase()` 적용.

## 4. 운영 원칙 — MsgBox 금지(자동화 자산 한정)

운영 매크로(`Action_*` 계열, `modEngine` 본체)는 모달 다이얼로그 호출 금지. 야간/무인 실행 시 프로세스를 차단한다. Task A(`../issues/A_wait_pixel_esc.md`)에서 H2로 확정된 규칙:

```vb
' 금지
MsgBox "[TimeOut] " & sType & " 실패", vbCritical

' 권장
LogShim.LogJ "wait_timeout", _
    "{""type"":""" & sType & """,""color"":""" & hexColor & """}"
```

`ErrHandler`도 동일 패턴으로 LogJ 위임. 사용자 확인이 본질적으로 필요한 dev 도구(`modDevTools`, `modHelper`의 HUD)는 예외.

## 5. LogShim 브리지

`LogShim.bas`는 외부 애드인 `_UnivDbg.xlam`(`Application.UserLibraryPath`에 설치)으로 위임하는 옵션 shim. 애드인 부재 시 1회 Ping → `dbgEnabled=False` 캐시 → 모든 후속 호출 무해 무시. 호출자는 가용 여부를 검사할 필요 없다. 출력 위치는 `../_dbg/UnivRPA.jsonl`(애드인 측 책임).

## 6. 진입점

| 진입점 | 위치 | 용도 |
|---|---|---|
| `RunRPA()` | `modEngine` | 메인 RPA 실행 (사용자 버튼) |
| `StartHUD_Final()` | `modHelper` | 픽셀 좌표 캡처 HUD (Caps Lock 토글로 기록) |
| `FormatRPAScript_Final()` | `modDevTools` | 스크립트 신택스 컬러링/들여쓰기/검증 |
| `Run_System_SelfDiagnosis()` | `modTester` | 내장 회귀 테스트(Memory/Type/Flow) |

빌드/린트/CI 없음. 테스트는 `modTester` 매크로를 VBE 또는 sbroenne `vba.run`으로 수동 실행한다.

## 7. 코드 동기 — `src/*.bas|cls|frm` ↔ `../UnivRPA.xlsb`

`CodeManager.bas`가 양방향 브리지(자동 아님, 매크로 호출 필요):

- **import**: `ImportVBA_From_ImportFolder` — `src/IMPORT/*.bas|cls|frm`(별도 폴더)을 읽어 `UpdateModuleCode`로 VBE 메모리에 주입. 워크북 저장 전까지 디스크 미반영.
- **export**: `ExportAllVBAElements` — VBE 모듈을 `src/*.bas|cls|frm`(UTF-8)로 출력. git 추적용.

권장 흐름:

1. VBE에서 편집 → 워크북 저장 → `ExportAllVBAElements` → `git add src/*.bas` → commit, **또는**
2. `src/*.bas` 직접 편집 → 사본을 `IMPORT/`에 복사 → `ImportVBA_From_ImportFolder` 실행 → 워크북 저장.

`src/*.bas`만 수정하고 워크북을 동기화하지 않으면 git에는 반영되나 운영 자산은 옛 코드. 반대 방향도 동일 위험.

## 8. 작업 워크플로 (사이클 기반)

`../issues/[TaskID]_[summary].md`에 task spec. 진행 중 task의 commit 컨벤션:

```
task-<ID> cycle <N>: <단계 요약>
```

각 사이클 종료 시 1회 commit, commit 없이 다음 사이클 진입 금지. 일반 흐름: cycle 0 셋업 → cycle 1 베이스라인 → 가설 검증 사이클들 → 회귀 → 회수(임시 디버그 코드 제거) → 결과 보고. 완료 사례: `../issues/A_wait_pixel_esc.md`.

가드레일 발동 시 자율 루프 즉시 정지하고 사용자에게 보고, 임의 재시작 금지.

## 9. 디렉터리 인접 자산 (src 외부)

- `../UnivRPA.xlsb` — 운영 자산. 런타임의 단일 source of truth.
- `../UnivRPA.md` — 자동 생성 코드베이스 레퍼런스(헤더의 export date 확인 필요).
- `../issues/` — task spec.
- `../_dbg/UnivRPA.jsonl` — 디버그 로그 스트림(LogShim 출력).
- `../archive/[YYYYMMDD_HHMMSS]/` — 이전 세션 스냅샷.
- `../old/` — deprecated 아카이브(참조용, 사용 금지).
