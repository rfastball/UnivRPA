UnivRPA Script Assist — 구현 spec (v5)

v4 대비 변경 (핵심):

* Assist의 셀 서식 페인팅을 전면 제거 — 시각 렌더링은 배치 포매터 단독 소유
* 오류 가시성을 Assist Note 채널(레거시 Comment)로 이전 + StatusBar 폴백
* Negative Scope에 "가시성 승격(Private→Public) 예외" 도입 (ValidateLine, EnsureHeaders)
* ApplyCommandStyle / ApplyErrorStyle 폐기, ClearAssistState 의미 축소
* GetCmdDef 정규화 일치 규약 명문화
* SheetChange 다중 셀 입력 처리 규약 추가
* 재진입 안전성 근거 명시

⸻

목표

기존 modCmdRegistry(SSOT)를 읽기 전용으로 소비하는 별도 편집 보조 서브시스템(Script Assist)을 구축한다.
Assist는 시각 렌더링을 소유하지 않으며, 비시각 의미 피드백(검증·힌트·오타 후보·오류 Note)만 담당한다.

철학:

* UnivRPA 명령어셋은 약 30개 내외이며 암기 가능한 규모이다.
* 명령어 입력은 자연 타이핑을 기본으로 한다.
* 편집기는 자동완성 UI보다 안정성, 가벼움, Excel-native 경험을 우선한다.
* 편집기는 코드 에디터가 아니라 DSL 작성 보조 도구이다.
* 편집기는 셀 서식(색·배경·테두리·글꼴)을 칠하지 않는다. 입력 중에는 평범한 셀이고, 시각 렌더링은 포맷 실행 시점에 배치 포매터가 일괄 수행한다(IDE의 save/compile 루프).

⸻

범위

신규

modScriptAssist

편집 보조 어댑터.

역할:

* COMMAND열 판별
* 명령어 유효성 검사
* 파라미터 스키마 힌트 표시 (StatusBar)
* 오타 후보 제안
* 오류 Note 부착 및 복구 (Assist 전용 Comment)

비역할 (명시적 제외):

* 셀 색·배경·테두리·글꼴 등 일체의 서식 적용
* 블록 깊이(depth) 계산
* 명령어 색상 그룹/페인트 스코프 렌더링

제공 기능:

ResolveCmdCol(Sh)
IsScriptSheet(Sh)
ValidateCommandCell(Target)
ShowParamHint(Target)
GetClosestCommand(rawCmd)
ApplyErrorNote(Target, rawCmd, suggestion)
ClearAssistNote(Target)
ClearAssistState(Target)
SetupScriptSheet(Sh)
IsBlankCommand(rawCmd)
IsCommentCommand(rawCmd)
NormalizeCmd(rawCmd)

폐기된 v4 기능:

ApplyCommandStyle   ← 페인팅 전면 제거
ApplyErrorStyle     ← 오류 표현을 Note 채널로 이전

ClearAssistState는 "Assist가 생성한 Note 제거"로 의미를 축소한다. 셀 서식은 손대지 않는다.

⸻

편집

ThisWorkbook

필요 최소 수준 후킹만 허용.

후킹 후보:

Workbook_SheetSelectionChange
Workbook_SheetChange

역할:

* COMMAND 선택 시 StatusBar 힌트 갱신
* COMMAND 수정 시 유효성 검사
* 오류 Note 부착 및 복구

⸻

읽기 전용 소비 (수정 없음)

modCmdRegistry.AllCommands
modCmdRegistry.GetCmdDef
CmdDef
CmdDef.Param
ParamSpec
ParamSpec.Meaning
ParamSpec.DataType
ParamSpec.Required

참고: ParamSpec은 Meaning / DataType / Required를 Public 필드로 노출한다. ShowParamHint는 이를 읽기만 하며 ParamSpec을 수정하지 않는다.

⸻

가시성 승격 예외 (Negative Scope의 타입 한정 예외)

규약:

Negative Scope 모듈에 대해 "시그니처 보존 가시성 승격(Private→Public)"만 예외로 허용한다.
함수 본문·시그니처·동작 변경은 여전히 금지한다.
각 승격은 modScriptAssist의 소비자를 명시해야 한다. 소비자 없는 승격은 금지한다.

승격 대상 (2건):

modDevTools.ValidateLine
→ 소비자: modScriptAssist.ValidateCommandCell
→ 사유: IF/WHILE 연산자 슬롯의 "맨 '=' 금지(== 강제)" 검사를 엔진과 단일 기준으로 공유. GetCmdDef 단독 호출로는 이 뉘앙스가 누락됨.

modDevTools.EnsureHeaders
→ 소비자: modScriptAssist.SetupScriptSheet
→ 사유: 헤더 스키마(열 라벨/열폭) 및 시트 시그니처를 단일 출처로 유지. SetupScriptSheet 자체 작성 시 헤더 정의가 이원화되어 드리프트 발생.

승격하지 않는 대상 (소비자 부재):

modDevTools.IsBlockStarter   ← depth 계산용. Assist가 페인팅을 안 하므로 소비자 없음. Private 유지.
modDevTools.IsBlockEnder     ← 동상. Private 유지.

무관:

modDevTools.ApplyUnifiedStyle ← 이미 Public이나 Assist는 호출하지 않음(페인팅 비역할). 본 예외와 무관.

⸻

Negative Scope

다음 구성요소의 함수 본문(behavior)은 절대 수정하지 않는다.

modCmdRegistry
CmdDef
ParamSpec
modEngine
modContext
modActions
clsActionCommands
modHelper
StartHUD_Final
modDevTools
ValidateLine
ApplyUnifiedStyle
IsBlockStarter
IsBlockEnder
EnsureHeaders

단, 위 "가시성 승격 예외"에 명시된 ValidateLine / EnsureHeaders의 Private→Public 승격은 허용한다.
승격은 가시성 키워드 1개 변경에 한하며, 본문·인자·반환을 보존한다.

편집기는 위 구성요소를 읽기 전용으로만 소비한다.

⸻

폐기된 설계

다음 설계는 v5에서 사용하지 않는다.

[v4에서 이미 폐기]
ActiveX ComboBox
OLEObject("TempCombo")
clsComboHandler
WithEvents Combo 이벤트 (KeyDown/KeyUp/Change/DropDown/Activate)
Combo 재배선
실시간 필터링 / 컨텍스트 후보 축소 / 자동완성 드롭다운
Excel Validation 목록

[v5 신규 폐기]
Assist의 셀 서식 페인팅 (배경·색·테두리·글꼴)
Assist의 depth 계산
ApplyCommandStyle / ApplyErrorStyle
Assist에 의한 ColorGroup / PaintScope / BlockRole 렌더링

편집기는 명령어 선택 UI도, 시각 렌더링도 제공하지 않는다.

⸻

시각 렌더링 단일 소유

모든 셀 서식(색·중첩 배경·들여쓰기 가이드·오류 색)은 배치 포매터가 단독 소유한다.

소유자:

modDevTools.FormatRPAScript_Final
modDevTools.ApplyUnifiedStyle

Assist는 색 영역에 일절 개입하지 않는다.
따라서 라이브 입력 피드백(Assist)과 시각 렌더링(배치)은 채널이 분리되어 충돌하지 않는다.

⸻

컬럼 규약

A = ROW
B = COMMAND
C = PARAM1
D = PARAM2
E = PARAM3
F = COMMENT

⸻

시트 판별

시그니처:

A1 = "ROW"
B1 = "COMMAND"

위 조건을 동시에 만족하는 경우만 스크립트 시트로 간주한다.
그 외 모든 시트에서 Script Assist는 동작하지 않는다.

제외 예: RPA_LOG / 일반 데이터 시트 / 피벗 시트 / 설정 시트

⸻

초기화

SetupScriptSheet는 반자동 IDE 초기화 루틴에서 명시 호출한다.

동작:

빈 시트
→ modDevTools.EnsureHeaders 호출 (헤더 멱등 생성, 단일 출처)
→ 스크립트 시트 시그니처 확립

SelectionChange는 헤더를 자동 생성하지 않는다.

⸻

명령어 정규화 규약

Assist의 모든 빈/주석/미등록 판별은 게이트와 동일한 정규화를 사용한다.

NormalizeCmd(rawCmd):

Trim → Replace(">", "") → Trim → UCase

근거: modCmdRegistry.GetCmdDef가 동일 순서로 정규화한다. Assist가 이를 복제하지 않으면 들여쓰기된 "> # 주석"이나 공백 케이스에서 게이트와 판정이 어긋난다.
IsBlankCommand / IsCommentCommand는 NormalizeCmd 결과를 기준으로 판정한다.

⸻

COMMAND 입력 모델

COMMAND열은 자유 입력을 사용한다.

직접 타이핑 → Tab → PARAM1 이동 → 파라미터 입력

자동완성은 제공하지 않는다.

⸻

파라미터 힌트

COMMAND 셀 선택 시:

GetCmdDef(cmd).Param(1).Meaning / .DataType / .Required
GetCmdDef(cmd).Param(2)...
GetCmdDef(cmd).Param(3)...

를 읽어 StatusBar에 표시한다.

예:

CLICK
P1: x (num, required)
P2: y (num, required)
P3: 사후대기ms (num)

출력은 Application.StatusBar 만 사용한다. UserForm UI는 사용하지 않는다.

⸻

StatusBar 규약

다음 경우 즉시 해제한다 (Application.StatusBar = False):

COMMAND 셀 이외
COMMAND열 이탈
비스크립트 시트
빈 COMMAND

편집기는 StatusBar를 점유하지 않는다. 필요 시점에만 잠시 표시하고 즉시 반환한다.

⸻

유효성 검사

COMMAND 수정 후 ValidateCommandCell(Target)이 다음 순서로 판정한다.

1. NormalizeCmd 후 빈/주석 → 오류 아님 → ClearAssistNote 후 종료
2. ValidateLine(cmd, p1, p2, p3) 호출 (승격된 Public)
   - True  → 유효 → ClearAssistNote (오류 복구)
   - False → 미등록 또는 연산자 위반 → ApplyErrorNote

⸻

명령어 부재 상태 (오류 아님)

빈 셀 ("")
→ Assist Note 제거
→ StatusBar 해제
→ 오류 표시 없음

주석 (# 설명)
→ Assist Note 제거
→ StatusBar 해제
→ 오류 표시 없음

빈/주석의 셀 색은 배치 포매터 소관이며 Assist는 관여하지 않는다.

⸻

미등록 / 위반 명령어 (오류)

예: CLCIK / FOUCS / ENF_IF / IF의 맨 '=' 위반

→ Assist Note 부착 (오류 마커)
→ 오타 후보 계산
→ 입력 자체는 막지 않는다 (자동 수정 금지)

⸻

오류 표시 채널 — Assist Note (레거시 Comment)

오류 가시성은 셀의 레거시 Comment(Range.Comment)로 표현한다. 셀 모서리 마커로 수정 전까지 지속된다.
Assist는 색을 칠하지 않으므로 배치 포매터와 색 충돌이 없다.

Assist Note 본문 규약:

[ScriptAssist]
Unknown command: CLCIK
Did you mean CLICK?

소유권 식별:

코멘트 본문이 "[ScriptAssist]" 프리픽스로 시작하면 Assist 소유로 간주한다.

ApplyErrorNote(Target, rawCmd, suggestion) 동작:

Target.Comment Is Nothing
  → AddComment (신규 생성)
IsAssistNote(Target.Comment) = True
  → Comment.Text 본문 교체
사용자 Note 존재 (태그 없음)
  → 덮어쓰지 않음 → StatusBar 폴백으로만 오류 알림

ClearAssistNote(Target) 동작:

IsAssistNote(Target.Comment) = True 인 경우에만 Delete.
사용자 작성 Note(태그 없음)는 절대 삭제하지 않는다.

⸻

셀당 코멘트 1개 제약 — StatusBar 폴백

레거시 Comment는 한 셀에 하나만 존재한다.
사용자가 해당 COMMAND 셀에 이미 자기 메모를 달아둔 경우, Assist는 오류 Note를 부착할 자리가 없다.

규약:

사용자 Note가 점유 중인 셀
→ Assist는 Note 부착을 포기한다 (사용자 메모 보존)
→ 오류는 Application.StatusBar 로만 표시한다

즉 (A)는 "Note 우선, 사용자 Note 점유 시 StatusBar 폴백" 하이브리드다.

⸻

오타 후보 계산

후보 원천:

modCmdRegistry.AllCommands()

알고리즘:

Levenshtein Distance

규칙:

길이 1~4 → 거리 1까지
길이 5 이상 → 거리 2까지

최소 거리 후보가 하나일 때만 제안한다. 동률이면 제안하지 않는다.

출력 위치:

Application.StatusBar 또는 Assist Note (오류 채널과 동일 경로)

⸻

자동 수정 금지

편집기는 셀 값을 자동 변경하지 않는다.

CLCIK → CLICK 자동 치환 금지. 사용자가 직접 수정한다.

이 규약은 재진입 안전성의 근거이기도 하다(아래).

⸻

오류 복구

명령어가 정상화되면:

Assist Note 제거 (ClearAssistNote)
StatusBar 갱신/해제

사용자 Note는 유지한다.
셀 색 복구는 배치 포매터 소관이며 Assist는 관여하지 않는다.

⸻

이벤트 규율

SheetSelectionChange

스크립트 시트 판별
COMMAND열 판별 (ResolveCmdCol)
COMMAND 셀이면 StatusBar 힌트 갱신
COMMAND 이외/이탈/비스크립트 시트면 StatusBar 해제

SheetChange

Target을 COMMAND열과 Intersect
교집합이 없으면 즉시 반환
교집합 셀을 순회하며 셀 단위로 ValidateCommandCell 호출
(열 전체 붙여넣기·다중 행 입력 시 Target이 범위로 들어오므로 단일 셀 가정 금지)

이벤트는 가볍고 무상태(stateless)하게 유지한다.

재진입 안전성:

Assist는 셀 값을 변경하지 않는다 (자동 수정 금지, 색 비페인팅).
Note/StatusBar만 조작하므로 SheetChange가 재귀 발화하지 않는다.
따라서 EnableEvents 재진입 제어는 불필요하다.
(주의: 향후 Assist가 셀 값을 쓰는 코드를 추가하면 이 전제가 깨진다.)

⸻

수락 기준

* A1="ROW" 그리고 B1="COMMAND"인 시트에서만 Script Assist가 작동한다.
* COMMAND열 이외에서는 침묵한다.
* COMMAND는 자유 입력 가능하다.
* Tab으로 PARAM1 입력 흐름이 자연스럽다.
* COMMAND 선택 시 ParamSpec 기반 StatusBar 힌트가 표시된다.
* 빈 COMMAND와 주석 COMMAND는 오류 처리되지 않는다.
* COMMAND열 이탈/비스크립트 시트에서 StatusBar가 해제된다.
* 미등록·연산자 위반 명령어는 Assist Note(또는 폴백 StatusBar)로 오류가 표시된다.
* Assist는 어떤 경우에도 셀 색·배경·테두리·글꼴을 변경하지 않는다.
* Levenshtein 기반 최근접 후보가 유일하면 제안되고, 동률이면 제안되지 않는다.
* 셀 값 자동 수정은 발생하지 않는다.
* 오류 수정 시 Assist가 생성한 Note만 제거된다. 사용자 Note는 유지된다.
* 사용자 Note가 점유한 COMMAND 셀에서는 Assist가 Note를 덮지 않고 StatusBar로만 오류를 알린다.
* SheetChange가 다중 셀 입력(열 붙여넣기 등)에서도 누락 없이 동작한다.
* 레지스트리에 새 명령어를 추가하면 검증·힌트·오타 후보가 자동 반영된다.
* modDevTools의 함수 본문은 수정하지 않는다. ValidateLine / EnsureHeaders의 Private→Public 가시성 승격만 예외로 허용하며, 본문·시그니처는 보존된다.
* IsBlockStarter / IsBlockEnder는 승격하지 않고 Private을 유지한다.
* modCmdRegistry, modEngine, modContext, modActions, modHelper 어느 줄도 수정하지 않는다.

⸻

구현 순서

1. modScriptAssist 기반
   * NormalizeCmd
   * ResolveCmdCol
   * IsScriptSheet
   * SetupScriptSheet (EnsureHeaders 승격 호출)
2. 명령어 검증
   * IsBlankCommand
   * IsCommentCommand
   * ValidateCommandCell (ValidateLine 승격 호출)
3. StatusBar 스키마 힌트
   * ShowParamHint
   * ClearStatusBar
4. 오타 후보
   * GetClosestCommand
   * Levenshtein 연동
5. 오류 Note 채널
   * IsAssistNote
   * ApplyErrorNote (사용자 Note 점유 시 StatusBar 폴백)
   * ClearAssistNote
   * ClearAssistState (Note 제거로 축소)
6. 가시성 승격 (modDevTools)
   * ValidateLine: Private → Public
   * EnsureHeaders: Private → Public
   * (본문 무변경, 가시성 키워드만)
7. ThisWorkbook 후킹
   * Workbook_SheetSelectionChange
   * Workbook_SheetChange (Intersect + 셀 순회)

각 단계는 독립적이며 되돌림 가능하다.
가역성 주의: Assist 본체는 통째 제거 가능하나, 6단계의 가시성 승격 2건은 modDevTools에 흔적으로 남는다(무해, 한 줄 복귀로 가역).
