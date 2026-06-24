# UnivRPA — 명령어 메타데이터 레지스트리 구현 설계문서

> **상태:** 설계 확정 (구현 전)
> **대상 프로젝트:** UnivRPA.xlsb
> **목적:** 명령어 정의의 단일 출처(SSOT) 확립 및 디스패치 구조 정식화

> **[Revisit 2026-06-24]** 본 문서를 `UnivRPA.xlsb` 현행 명령어셋(2026-06-24 export, **총 30개**)에 맞춰 재정렬함. 구버전 전제였던 패턴 명령어 가족(`IF_PIXEL`/`IF_CLIP`/`IFNOT_PIXEL`), `SET_MEM`, `MOUSE-*`, 그리고 `ValidateLine` **부분일치**는 현행 코드에 **존재하지 않는다**. 조건의 "대상 종류"는 명령어 이름이 아니라 토크나이저(`{$PIXEL:x,y}`, `{$CLIP}`, `{name}`)와 연산자(`==`/`LIKE`/…)로 이미 분리되었고, `ValidateLine`은 이미 완전일치 화이트리스트다. 관련 논거는 §1.2·§4.2·§5.2·§5.3·§10에서 갱신했으며, 현행 명령어 전수 인벤토리는 **부록 A**에 신설했다.
>
> 아울러 본 개정은 `clsMemoryCommands`의 정체를 **"메모리"에서 "lvalue-writer"로** 재정의하면서, (1) `modEngine`의 `ParseLValue`를 이 클래스로 흡수하고 (2) `{$LOOP}`로 읽기만 가능하던 *이름 붙은 정수 상태*의 읽기-쓰기 비대칭을 닫기 위해 `INC`/`DEC`를 **신규 등록**한다(§4.2·부록 A.3). 반대로 산술·표현식 평가는 **비범위(out of scope)로 명문화**한다 — 토크나이저는 치환기이지 평가기가 아니다(§5.3·§10).

---

## 0. 한 줄 요약

명령어에 대한 지식이 5개 모듈에 흩어져 있는 **동기화 부채**를, 명령어 메타데이터 **레지스트리**(Layer 0)로 통합한다. SSOT는 목표가 아니라 레지스트리 도입의 **부수효과**다. 디스패치는 카테고리별 `Implements` 클래스로 정식화한다.

---

## 1. 문제 정의 — 청산할 빚

### 1.1 흩어진 5조각

명령어 1개에 대한 지식이 아래 5곳에 분산되어 있다. 새 명령어 1개 추가 시 5군데를 동시에 수정해야 하며, 이 5곳은 서로를 모르므로 하나라도 빠뜨리면 조용히 어긋난다.

| 위치 | 담고 있는 명령어 지식 |
|---|---|
| `modDevTools.ValidateLine` | 유효 명령어 목록 (완전일치 화이트리스트) |
| `modEngine.ProcessAction` Select Case | 액션 디스패치 |
| `modEngine.ProcessFlowControl` Select Case | 흐름제어 디스패치 |
| `modDevTools.ApplyUnifiedStyle` | 카테고리별 색상 그룹 분류 |
| `modDevTools.IsBlockStarter` / `IsBlockEnder` | 블록 여닫이 판정 |

> **[참고] 추가 매립 지점.** 위 5곳 외에도 명령어 리터럴이 박혀 있는 곳이 더 있다: `ProcessFlowControl`의 `SELECT` 전방 스캔(`"SELECT"`/`"CASE"`/`"END_SELECT"` 하드코딩), `FindBlockPair`의 `SearchType` 분기(`IF`/`WHILE`/`REPEAT`/`SELECT`), `PreScanLabels`의 `"LABEL"`. 1차 범위는 위 5조각으로 한정하되, 레지스트리 정착 후 이들도 게이트 기반으로 흡수 가능.

### 1.2 추가로 식별된 구조적 결함

- **검증·디스패치 목록의 이중 관리:** `ValidateLine`의 화이트리스트와 두 디스패처(`ProcessAction`/`ProcessFlowControl`)의 `Select Case`는 *현재는 우연히 일치*하지만(양쪽 모두 30개), 서로를 모르는 **별도 사본**이라 구조적으로 표류를 막지 못한다. 한쪽에만 명령어를 추가하면 — 검증 통과 후 디스패치 거부, 혹은 그 반대 — 조용히 어긋난다. *(구버전의 `InStr(cmd,"IF_")=1` 부분일치 결함은 완전일치 전환으로 이미 해소됨. 이제 막을 대상은 "버그"가 아니라 "두 목록의 미래 표류"다.)*
- **폴백 체인의 군더더기:** `ExecuteSmartAction_Core`는 `ProcessFlowControl` → `ProcessAction` 순으로 물어보고 `CMD_NOT_HANDLED`로 넘기는 순차 폴백 구조다. `CMD_NOT_HANDLED`라는 특수 반환값의 존재 자체가 이 어색함의 증상이다.
- **서식의 숨은 축:** `ApplyUnifiedStyle`은 "무슨 색"(ColorGroup) 외에 "칠하는 범위"(줄 전체 vs B열만)라는 별도 축을 우연히 겸하고 있다. (현행 코드에서 이 두 축의 불일치가 실측됨 — 부록 A 참조.)
- **파라미터 해석의 암묵지:** `SET`의 P1(좌변 주소)은 원본(Raw, `ParseLValue`가 직접 파싱)을 쓰고 P3(값)은 해석본(Resolve)을 쓴다. 같은 명령어 안에서도 슬롯마다 "해석 여부"가 갈리는데, 이 규칙이 코드에 암묵적으로만 존재한다.

---

## 2. 정체와 명명

| 항목 | 내용 |
|---|---|
| 정확한 정체 | **명령어 메타데이터 레지스트리** |
| 표준 용어 | 인터프리터/VM의 instruction table (opcode table) 패턴 |
| 아키텍처 위치 | **Layer 0 / 명령어 정의 계층** (기존 Layer 1~4를 횡단하는 참조 계층) |
| SSOT와의 관계 | SSOT는 목표가 아니라 레지스트리 도입의 **결과적 속성**. "레지스트리를 만들면 그 부수효과로 SSOT가 달성된다"가 정확한 인과 |

기존 5조각은 전부 이 레지스트리를 읽는 **파생 함수**로 환원된다.
예: `IsBlockStarter(cmd)` → `GetCmdDef(cmd).BlockRole = STARTER` 한 줄.

---

## 3. 아키텍처 — 3층 구조

```
┌─────────────────────────────────────────────┐
│  [1층] 데이터: 레지스트리                       │
│   Dictionary( cmd → CmdDef 객체 )             │  ← 유일한 진실
└────────────────────┬────────────────────────┘
                     │ 오직 이 문으로만 접근
                     ▼
┌─────────────────────────────────────────────┐
│  [2층] 게이트: GetCmdDef(cmd)                 │
│   - 명령어 정규화(">", Trim, UCase)           │  ← 모든 조회의 단일 통로
│   - 등록 확인 → CmdDef 반환 or Nothing        │
└────────────────────┬────────────────────────┘
                     │ 모두가 여기서만 묻는다
   ┌──────────┬──────┴───────┬──────────────┐
   ▼          ▼              ▼              ▼
ValidateLine IsBlockStarter ApplyStyle  디스패치
[3층] 소비자: 스스로 판단하지 않고 게이트에 질문만
```

**규율:** 레지스트리(Dictionary)에 직접 접근하는 코드는 없다. 모든 접근은 `GetCmdDef` 게이트 하나로 강제된다.

---

## 4. 디스패치 전략 — 2단 디스패치

### 4.1 함수 포인터 테이블을 쓰지 않는 이유 (검증 완료)

VBA에는 네이티브 함수 포인터 호출이 없다(`AddressOf`로 주소 전달은 되나 Basic 내부에서 그 주소로 호출 불가). 세 가지 우회로 모두 부적합:

| 방식 | 문제 |
|---|---|
| `DispCallFunc` (OleAut32) | 인자 있는 함수는 매 호출마다 VarType/VarPtr 마샬링. 64비트 취약. 핫루프에 부적합 |
| `Application.Run "이름"` | Private 호출 불가, 문자열 오타 런타임까지 미검출, 오버헤드 |
| `Implements` 1:1 | 명령어 수만큼 클래스(30+). 과잉 |

**결론:** `Select Case`는 VBA에서 가장 빠르고 타입 안전한 디스패치다. 진짜 포인터 테이블로 바꾸면 성능·안전성이 둘 다 후퇴한다.

### 4.2 채택안 — 카테고리별 클래스 (Implements 묶음)

클래스의 단위는 "명령어"가 아니라 **"실행 방식이 같은 명령어 묶음"**. 현행 30개(+ 본 설계 추가 `INC`/`DEC` = 32개)를 4개 클래스로 묶는다.

```
ICommand (인터페이스)
  └ Execute(args) As Long   ' 공통 약속

clsFlowCommands   Implements ICommand   ' IF/WHILE/REPEAT/SELECT/GOTO/CALL/RETURN/BREAK/CONTINUE/LABEL/...
clsActionCommands Implements ICommand   ' CLICK/RCLICK/KEY/COMBO/TYPE  (+ SYSTEM: FOCUS/SLEEP/NOTIFY/EXIT)
clsVisionCommands Implements ICommand   ' WAIT_APPEAR/WAIT_VANISH/CLICK_SYNC/ASSERT
clsMemoryCommands Implements ICommand   ' SET / INC / DEC — 이름 붙은 슬롯 쓰기(lvalue-writer); ParseLValue 소유
```

> **[갱신] SYSTEM 카테고리의 귀속.** 의미상 카테고리는 5개(FLOW/ACTION/VISION/MEMORY/SYSTEM)이지만, 현행 코드가 시스템 명령어(FOCUS/SLEEP/NOTIFY/EXIT)를 `ProcessAction` 안에서 함께 처리하는 것과 동일하게 **`clsActionCommands`가 SYSTEM도 처리**한다. 핸들러 맵에서 `"SYSTEM"`과 `"ACTION"` 키가 같은 인스턴스를 가리키게 두면, *의미상 5카테고리 / 디스패치 4클래스*가 양립한다("30→4" 헤드라인 유지). 1:1을 선호하면 `clsSystemCommands` 신설 — §10 미결.

> **[갱신] `clsMemoryCommands`의 정체 — "메모리"가 아니라 "lvalue-writer".** 이 클래스가 얇아 보였던 건 카테고리 오류가 아니라 lvalue-writer가 `SET` 하나뿐이었기 때문이다. 본 설계는 두 조치로 이를 닫는다:
> 1. **`ParseLValue` 흡수** — "어떤 좌변이 쓰기 가능한가"(`{$CLIP}`=가능, 그 외 `$…`=읽기전용, `{name}`=가능)는 메모리 명령어의 고유 지식인데 현재 `modEngine`에 얹혀 있다(계층 위반). 이를 클래스로 내린다.
> 2. **`INC`/`DEC` 추가** — `{$LOOP}`로 *읽기만* 가능하던 "이름 붙은 정수 상태"의 **읽기-쓰기 비대칭을 닫는다**(read-modify-write). compute 도입이 아니라 *닫힘(closure)*이 근거 — §10의 산술 비범위 선언과 한 쌍이다.
>
> 그 결과 클래스는 명령어 개수가 아니라 *"쓰기 대상의 합법성 + 정수 슬롯 변이를 소유한 단위"*라는 책임으로 완결된다. (효용이 아니라 닫힘으로 정당화 — 클래스를 채우려 명령어를 발명하는 것은 거꾸로다.)

### 4.3 2단 디스패치 흐름

```vba
' 1단 (카테고리 → 클래스 선택): 핸들러 Dictionary
Dim cat As String: cat = GetCmdDef(cmd).Category    ' SYSTEM도 ACTION 핸들러로 매핑
Dim handler As ICommand
Set handler = m_handlers(cat)

' 2단 (명령어 분기): 선택된 클래스 내부의 Select Case
nextLine = handler.Execute(cmd, args)
```

거대 `Select Case` 하나가 4개의 작은 `Select Case`로 쪼개져 각 클래스 안으로 들어간다. 기존 `ProcessAction`/`ProcessFlowControl`의 분리는 이 구조의 절반을 이미 구현한 상태.

### 4.4 이 구조가 제거하는 것

- **`CMD_NOT_HANDLED` 폴백 체인 소멸** — 테이블이 카테고리를 알고 있으니 "물어보고 아니면 넘기고"가 불필요. 담당자에게 직행.
- **관심사의 물리적 분리** — `clsFlowCommands`엔 흐름제어만, 마우스/타이핑 코드 없음.
- **변경의 국소성** — 새 명령어 추가가 한 카테고리 안에 갇힌다.

---

## 5. 자료구조 — 전용 클래스

### 5.1 레지스트리 레코드

순수 Dictionary + 중첩 배열은 파라미터 스펙이 4속성×3슬롯으로 펴지는 순간 `d(7)(2)` 인덱스 지옥이 된다. 디스패치용 클래스를 이미 만드는 김에 레코드도 전용 클래스로 둔다.

```
레지스트리 = Scripting.Dictionary
             key:   명령어 정식명 (String)
             value: CmdDef 객체

CmdDef (클래스 모듈)
  ├ [식별]   Canonical, Aliases, IsPattern
  ├ [구조]   Category, BlockRole, BlockGroup
  ├ [서식]   ColorGroup, PaintScope
  └ [파라미터] P1, P2, P3   ← 각각 ParamSpec 객체

ParamSpec (클래스 모듈 — 중첩)
  └ Meaning, DataType, ResolveMode(Resolve/Raw), Required
```

`def.P1.IsRaw` 형태의 접근이 `d(7)(2)`를 대체한다.

### 5.2 스키마 전체 명세

원안 5컬럼 `{Category, BlockRole, BlockPair, ColorGroup, IsPattern}`은 불충분. 코드 요구에 역대조하여 보강한 확정안:

| 그룹 | 속성 | 용도 | 파생/환원되는 기존 조각 |
|---|---|---|---|
| 식별 | `Canonical` | 명령어 정식명 | (키) |
| 식별 | `Aliases` | 별칭 배열 | (신규 — 별칭 흩어짐 방지. 현행 명령어엔 별칭 없음, §10) |
| 식별 | `IsPattern` | 별칭/패턴 여부 | `ValidateLine` (현행: 패턴 명령어 부재 → 사실상 사문화, §5.3) |
| 구조 | `Category` | FLOW/ACTION/VISION/MEMORY/SYSTEM | 1단 디스패치 키(SYSTEM→ACTION 핸들러), `ApplyUnifiedStyle` |
| 구조 | `BlockRole` | STARTER/ENDER/MID/NONE | `IsBlockStarter`/`IsBlockEnder` |
| 구조 | `BlockGroup` | IF/WHILE/REPEAT/SELECT/— | (`BlockPair` 개명) · `FindBlockPair`의 `SearchType` |
| 서식 | `ColorGroup` | 색상 키 | `ApplyUnifiedStyle` 색상 분류 |
| 서식 | `PaintScope` | FULLROW / CMDCELL | **신규** — 칠하는 범위 |
| 파라미터 | `P1/P2/P3` | 각 `{Meaning, DataType, ResolveMode, Required}` | 파라미터 힌트, decode |

**핵심 보강 둘:**
1. **`PaintScope` 분리** — `ApplyUnifiedStyle`은 Flow류/System류(줄 전체 칠)와 Action/Vision류(B열만 칠 + 중첩 시 C~E 연녹 배경)를 구분한다. ColorGroup이 이 둘을 우연히 겸하던 것을 명시 컬럼으로 분리. **현행 코드가 이 분리를 실증한다**: `SET`은 `MEMORY` 카테고리지만 ColorGroup은 ACTION(연청), `FOCUS`/`SLEEP`은 `SYSTEM`이지만 ColorGroup은 ACTION, `NOTIFY`/`EXIT`만 SYSTEM(연회색)+FULLROW. 즉 Category·ColorGroup·PaintScope는 서로 독립 축이다(부록 A).
2. **`ResolveMode` 축** — `SET`/`INC`/`DEC`의 P1=Raw(좌변=주소), 나머지=Resolve. **규칙으로 일반화: 모든 lvalue 슬롯과 연산자 슬롯은 RAW, 그 외는 RESOLVE.** 파라미터 스펙에 해석 여부를 명시. 이는 표준 VM의 **decode/execute 분리**와 정확히 일치하며, 파라미터 스펙 도입은 곧 decode 단계를 데이터 주도로 정식화하는 것. (현행 RAW 슬롯 전수는 부록 A.1.)

### 5.3 "패턴류 명령어"는 이미 토크나이저로 해소됨 (전제 폐기)

**[갱신]** 본 절의 원래 전제 — `IF_PIXEL`/`IF_CLIP`/`IF_MEM`/`IFNOT_PIXEL` 같은 패턴 명령어 가족과 `ValidateLine`의 부분일치 — 는 현행 명령어셋에 **더 이상 존재하지 않는다**. 조건 판정의 "대상 종류"(픽셀/클립/메모리/리터럴)는 명령어 이름이 아니라 **토크나이저**(`{$PIXEL:x,y}`, `{$CLIP}`, `{name}`, `{$LOOP}` 등)와 **연산자**(`==` `!=` `>` `<` `>=` `<=` `LIKE` `STARTS` `ENDS` `CONTAINS`)로 분리되었다. 따라서 `IF`/`WHILE`는 가족이 아니라 **단일 명령어**이며, `ValidateLine`은 **이미 완전일치 화이트리스트**다.

결과적으로:
- **검증·디스패치 불일치는 이미 봉쇄됨.** 부분일치가 사라졌으므로 `IF_BANANA`류 통과 경로가 원천적으로 없다. 레지스트리의 역할은 *이미 일치하는* 두 목록(`ValidateLine` ↔ 디스패치)이 앞으로도 **독립 표류하지 않게 강제**하는 것으로 재정의된다(§7 역검증이 이 파수꾼).
- **`IsPattern` 컬럼은 명령어 차원에서 사실상 사문화.** 현행 명령어셋에 표기 별칭을 가진 명령어가 없다(§10). 컬럼은 향후 별칭 도입 여지를 위한 예약 슬롯으로만 잔존.

**[이관된 산개 지식]** 원래 5.3이 겨냥하던 "유한한 변종을 화이트리스트로 박제" 문제는, 명령어가 아니라 **토큰·연산자·키 이름**이라는 형제 테이블로 자리를 옮겼다:

| 산개된 표 | 위치 | 성격 |
|---|---|---|
| 토큰 테이블 | `modEngine.ResolveToken` | `$BLANK/$TAB/$CR/$LF/$SPACE/$CLIP/$LOOP/$PIXEL:` 하드코딩 |
| 연산자 테이블 | `modActions.Action_CheckCondition` | 비교 연산자 + 맨 `=` 금지 |
| 키 이름 테이블 | `modActions.ParseKeyCode` | `ENTER/ESC/F1..F12/방향키…` 키 별칭 |

> **[예외] `ParseLValue`는 형제가 아니라 흡수 대상.** "어떤 좌변이 쓰기 가능한가"는 토큰·연산자·키와 달리 **명령어(lvalue-writer)의 고유 지식**이다. 따라서 별도 형제 레지스트리로 빼지 않고 `clsMemoryCommands`로 내린다(§4.2). 위 표의 셋만이 명령어 레지스트리와 진짜로 직교한다.

이들은 **명령어 레지스트리의 범위 밖**이다(별도 SSOT 후보). 본 설계는 명령어 레지스트리에 집중하되, 동일 패턴의 형제 레지스트리(특히 토큰 테이블)를 후속 과제로 남긴다(§10 미결). 인벤토리는 부록 A.2.

> **[경계] 토큰은 *치환기*지 *평가기*가 아니다.** 위 토큰들은 전부 "외부/상태 값을 문자열로 가져와 끼워넣기"다. 산술식(`{$ADD:..}`)·문자열 함수(`{$SUBSTR:..}`) 같은 *평가* 토큰은 토크나이저의 정체를 치환기에서 평가기로 바꾸고 연산자 우선순위·타입 강제·에러 전파라는 새 책임을 통째로 끌어들인다. 절반만(예: `{$ADD}`만) 넣으면 *그게* 진짜 미완성이다. 본 설계는 평가 계층을 **명시적 비범위로 선언**한다(§10) — "전부 제대로 하거나, 손대지 않거나". `INC`/`DEC`가 *명령어*이고 *토큰*이 아닌 것도 같은 선 위에 있다(lvalue 변이는 닫힘, 식 평가는 비범위).

---

## 6. 의존성 주입 (Init 패턴)

VBA는 생성자가 없으므로, 클래스가 `m_arrMain`/`modContext`/`modActions`에 접근하려면 `Init` 메서드로 수동 주입한다. **UnivDB 프로젝트의 기존 Init 주입 패턴을 재사용**(검증된 패턴).

적용 대상:
- 디스패치 클래스 4개: 실행에 필요한 `m_arrMain`, 컨텍스트 등을 `Init`으로 주입
- `CmdDef`/`ParamSpec`: `New` 후 `Init(...)` 팩토리 메서드로 속성 일괄 세팅

### 6.1 한 줄 등록 유지

객체를 만들고 속성을 하나씩 세팅하면 레지스트리 정의부가 명령어당 십수 줄로 부풀어 오히려 안 읽힌다. 헬퍼로 **한 줄 등록**을 유지한다:

```vba
' WHILE: P1=좌변(해석) · P2=연산자(원본 RAW!) · P3=우변(해석)
'   Reg(정식명, Category, BlockRole, BlockGroup, ColorGroup, PaintScope, P1, P2, P3)
Reg "WHILE", FLOW, STARTER, "WHILE", FLOW_COLOR, FULLROW, _
    P("좌변",   STR_T, RESOLVE, True), _
    P("연산자", STR_T, RAW,     True), _
    P("우변",   STR_T, RESOLVE, True)
```

- `Reg(...)` — 내부에서 `New CmdDef` + `Init` 호출 후 Dictionary에 등록
- `P(...)` — `New ParamSpec` + `Init` 후 반환하는 헬퍼

연산자 슬롯이 `RAW`로 박히는 점에 주목 — decode 규칙이 데이터로 드러난다. 클래스의 가독성(`def.P1.IsRaw`)과 정의부의 간결함(한 줄)을 동시에 확보. 디스패치 클래스 Init 주입과 같은 철학으로 일관성 유지.

---

## 7. 검증 자동화 — 역검증

레지스트리와 디스패치의 일체성을 구조가 공짜로 보장하지 못한다(테이블엔 등록했으나 클래스 Select Case 추가를 누락하는 경우). `modTester`에 역검증을 추가:

```
등록된 모든 명령어 순회
  → 각각 디스패치 시도 (빈 인자 등)
  → "알 수 없는 명령어" 응답 = 테이블엔 있는데 핸들러 없음 → FAIL 보고
```

레지스트리가 "정의의 단일 출처"라면, 이 테스트는 "정의와 구현의 동기화 파수꾼"이다. (현행 `ValidateLine`↔디스패치가 우연히 일치하는 상태를 *구조적으로* 못박는 장치이기도 하다 — §1.2.)

---

## 8. 모듈 구성 (신설/변경)

### 신설

| 종류 | 이름 | 역할 |
|---|---|---|
| 표준 모듈 | `modCmdRegistry` | 레지스트리 빌드 + `GetCmdDef` 게이트 + 핸들러 Dictionary 보유 |
| 클래스 | `CmdDef` | 레코드 1개 |
| 클래스 | `ParamSpec` | 파라미터 슬롯 1개 (CmdDef가 3개 보유) |
| 클래스 | `ICommand` | 디스패치 인터페이스 |
| 클래스 | `clsFlowCommands` | 흐름제어 핸들러 |
| 클래스 | `clsActionCommands` | 액션 핸들러 (+ SYSTEM 명령어 겸임, §4.2) |
| 클래스 | `clsVisionCommands` | 비전 핸들러 |
| 클래스 | `clsMemoryCommands` | lvalue-writer 핸들러 (SET/INC/DEC) + `ParseLValue` 흡수 |

### 변경

| 모듈 | 변경 내용 |
|---|---|
| `modEngine` | `ProcessAction`/`ProcessFlowControl` 거대 함수 → 얇은 2단 디스패처로 축소. `CMD_NOT_HANDLED` 폴백 제거. `ParseLValue`는 `clsMemoryCommands`로 이관 |
| `modDevTools` | `ValidateLine`/`ApplyUnifiedStyle`/`IsBlockStarter`/`IsBlockEnder` → `GetCmdDef` 읽는 파생 함수로 환원 |
| `modTester` | 역검증 테스트 추가 |

---

## 9. 마이그레이션 순서

1. **`modCmdRegistry` 신설** — `CmdDef`/`ParamSpec` 클래스 골격 + `Reg`/`P` 헬퍼 작성
2. **레지스트리 채우기** — 전 명령어를 한 줄 등록(부록 A를 초기 적재값으로 사용 + 부록 A.3의 `INC`/`DEC` 신규 등록). 파라미터 스펙(decode 테이블) 완성
3. **`GetCmdDef` 게이트 + 파생 함수** — 5조각을 게이트 위에서 재작성
4. **기존 5조각 교체** — 파생 함수 호출로 치환 (검증·서식·블록판정 먼저, 디스패치는 마지막)
5. **디스패치 클래스화** — `ICommand` + 카테고리 클래스 4개, Init 주입 배선. `clsMemoryCommands`에 `ParseLValue` 흡수 + `INC`/`DEC` 구현
6. **역검증 테스트 추가** — `modTester`에 동기화 파수꾼 투입

---

## 10. 미결 항목

| 항목 | 상태 |
|---|---|
| `BlockGroup`과 `FindBlockPair`의 관계 | `FindBlockPair`는 현재 `SearchType`(IF/WHILE/REPEAT/SELECT)으로 동적 탐색. `BlockGroup`을 테이블에 박으면 중복 가능 — 컨텍스트 필터 용도로 살릴지, `FindBlockPair`를 테이블 기반으로 바꿀지 미확정 |
| 별칭(Aliases) 포함 범위 | **현행 명령어셋엔 표기 별칭을 가진 명령어가 없다**(과거 `MOUSE-L`/`MOUSEL` 류 소멸). 따라서 `Aliases`/`IsPattern`은 당장은 빈 슬롯. 산개한 별칭 지식은 토큰(`ResolveToken`)·연산자(`Action_CheckCondition`)·키(`ParseKeyCode`) 테이블에 있으며, 이를 명령어 레지스트리에 흡수할지 형제 레지스트리로 분리할지 미확정(§5.3 이관표) |
| SYSTEM 카테고리의 디스패치 귀속 | 의미상 카테고리는 5개(FLOW/ACTION/VISION/MEMORY/SYSTEM)이나 핸들러 클래스는 4개. 현행처럼 SYSTEM을 `clsActionCommands`에 합칠지(핸들러 맵에서 `"SYSTEM"`→ACTION 별칭), 아니면 `clsSystemCommands`를 신설해 카테고리=클래스 1:1 불변식을 둘지 미확정. 전자는 "30→4" 헤드라인 유지, 후자는 불변식 명료 |
| 산술·표현식 평가 (비범위 **확정**) | **토크나이저는 치환기이며 평가기가 아니다**를 불변식으로 선언. `{$ADD}`/`{$MUL}`/`{$SUBSTR}`/`{$LEN}` 등 평가 토큰은 도입하지 않는다 — 연산자 우선순위·타입 강제·에러 전파를 절반만 들이면 그게 미완성이므로. 스크립트 내 계산이 필요하면 호스트 시트/사전계산에 위임. 재론 시 "전부 제대로 or 손대지 않기" 원칙 적용. (`INC`/`DEC`는 식 평가가 아니라 정수 슬롯의 read-modify-write라 이 선의 안쪽이다.) |
| `DEC`의 위치 (슈가) | `DEC {i} n`은 `INC {i} -n`으로 환원 가능 — 닫힘(closure)에 필수는 아니고 **가독성 슈가**다. 본 설계는 대칭/가독성을 위해 등록하되, 미니멀을 택하면 `DEC`를 빼고 `INC` 음수 스텝만 남겨도 닫힘은 유지된다 |
| `INC`/`DEC` ColorGroup | 현행 `SET`을 따라 ACTION(연청)으로 둠. MEMORY가 이제 실명령어 3개를 가지므로 전용 색(예: 연보라) 신설을 고려할 수 있으나 `ApplyUnifiedStyle` 팔레트 변경 필요 — 미적용 |
| 형제 레지스트리 분리 여부 | 토큰 테이블(`ResolveToken`)은 명령어 레지스트리와 같은 결함(산개)을 공유. 별도 `modTokenRegistry`로 뽑을지, 본 레지스트리 정착 후 후속 PR로 둘지 미확정 |

---

## 부록 A — 현행 명령어 인벤토리 (2026-06-24 export 기준)

> 레지스트리 **초기 적재값**을 현행 소스에서 역도출한 표. `ValidateLine`(정식 목록) · 두 디스패처 · `ApplyUnifiedStyle`(ColorGroup/PaintScope) · `IsBlockStarter`/`IsBlockEnder`(BlockRole)를 교차 대조함.
> **총 30개** = FLOW 16 · ACTION 5 · MEMORY 1 · SYSTEM 4 · VISION 4.

**[주목] 세 축의 독립성이 코드로 확인됨.** 현행 `ApplyUnifiedStyle`의 색 분류는 의미 카테고리와 일치하지 않는다 — `SET`(MEMORY)·`FOCUS`/`SLEEP`(SYSTEM)이 ACTION 색을 쓰고, `NOTIFY`/`EXIT`만 SYSTEM 색을 쓴다. 즉 `Category`·`ColorGroup`·`PaintScope`는 별개 축이며, 이를 분리한 §5.2 결정이 실측으로 정당화된다.

| 명령어 | Category | BlockRole | BlockGroup | ColorGroup | PaintScope |
|---|---|---|---|---|---|
| IF | FLOW | STARTER | IF | FLOW | FULLROW |
| ELSE | FLOW | MID | IF | FLOW | FULLROW |
| END_IF | FLOW | ENDER | IF | FLOW | FULLROW |
| WHILE | FLOW | STARTER | WHILE | FLOW | FULLROW |
| END_WHILE | FLOW | ENDER | WHILE | FLOW | FULLROW |
| REPEAT | FLOW | STARTER | REPEAT | FLOW | FULLROW |
| END_REPEAT | FLOW | ENDER | REPEAT | FLOW | FULLROW |
| SELECT | FLOW | STARTER | SELECT | FLOW | FULLROW |
| CASE | FLOW | MID | SELECT | FLOW | FULLROW |
| END_SELECT | FLOW | ENDER | SELECT | FLOW | FULLROW |
| BREAK | FLOW | NONE | — | FLOW | FULLROW |
| CONTINUE | FLOW | NONE | — | FLOW | FULLROW |
| GOTO | FLOW | NONE | — | FLOW | FULLROW |
| CALL | FLOW | NONE | — | FLOW | FULLROW |
| RETURN | FLOW | NONE | — | FLOW | FULLROW |
| LABEL | FLOW | NONE | — | FLOW | FULLROW |
| CLICK | ACTION | NONE | — | ACTION | CMDCELL |
| RCLICK | ACTION | NONE | — | ACTION | CMDCELL |
| KEY | ACTION | NONE | — | ACTION | CMDCELL |
| COMBO | ACTION | NONE | — | ACTION | CMDCELL |
| TYPE | ACTION | NONE | — | ACTION | CMDCELL |
| SET | MEMORY | NONE | — | ACTION | CMDCELL |
| FOCUS | SYSTEM | NONE | — | ACTION | CMDCELL |
| SLEEP | SYSTEM | NONE | — | ACTION | CMDCELL |
| NOTIFY | SYSTEM | NONE | — | SYSTEM | FULLROW |
| EXIT | SYSTEM | NONE | — | SYSTEM | FULLROW |
| WAIT_APPEAR | VISION | NONE | — | VISION | CMDCELL |
| WAIT_VANISH | VISION | NONE | — | VISION | CMDCELL |
| CLICK_SYNC | VISION | NONE | — | VISION | CMDCELL |
| ASSERT | VISION | NONE | — | VISION | CMDCELL |

### A.1 파라미터 decode 슬롯 (대표 예 — ResolveMode 검증)

| 명령어 | P1 | P2 | P3 |
|---|---|---|---|
| SET | 좌변/주소 · **RAW** | — | 값 · RESOLVE |
| IF / WHILE | 좌변 · RESOLVE | 연산자 · **RAW** | 우변 · RESOLVE |
| TYPE | 딜레이ms · RESOLVE | — | 텍스트 · RESOLVE |
| CLICK / RCLICK | x · RESOLVE | y · RESOLVE | 사후대기ms · RESOLVE(opt) |
| KEY | 키코드/이름 · RESOLVE | — | 사후대기ms · RESOLVE(opt) |
| COMBO | key1 · RESOLVE | key2 · RESOLVE | 사후대기ms · RESOLVE(opt) |
| WAIT_APPEAR / WAIT_VANISH | x · RESOLVE | y · RESOLVE | hexColor · RESOLVE |
| CLICK_SYNC | x · RESOLVE | y · RESOLVE | — |
| ASSERT | x · RESOLVE | y · RESOLVE | 조건(`!`접두=금지색) · RESOLVE |
| REPEAT | 반복횟수 · RESOLVE | — | — |
| GOTO / CALL | — | — | 라벨명 또는 줄번호 · RESOLVE |

> **RAW 슬롯의 규칙:** 토큰 치환을 적용하면 의미가 깨지는 두 부류뿐이다 — ① **lvalue**(쓰기 대상 주소: `SET`/`INC`/`DEC`의 P1)와 ② **연산자**(`IF`/`WHILE`의 P2). 그 외 모든 슬롯은 RESOLVE. 클래스화 시 슬롯별 ResolveMode를 데이터로 못박아 암묵지를 제거한다. (신규 `INC`/`DEC` 슬롯은 부록 A.3.)

### A.2 토큰 인벤토리 (형제 레지스트리 후보 — §5.3 / §10)

| 토큰 | 의미 | 좌변(쓰기) 가능 |
|---|---|---|
| `{name}` | 사용자 메모리 변수 | O |
| `{$CLIP}` | 클립보드 | O |
| `{$LOOP}` | 가장 안쪽 루프 현재 회차(루프 밖=0) | X |
| `{$PIXEL:x,y}` | 좌표 색상 Hex | X |
| `{$BLANK}` | 빈 문자열 | X |
| `{$TAB}` `{$CR}` `{$LF}` `{$SPACE}` | 리터럴 제어문자 | X |

연산자(`==` `!=` `>` `<` `>=` `<=` `LIKE` `STARTS` `ENDS` `CONTAINS`, 맨 `=` 금지)와 키 이름(`ENTER/ESC/TAB/F1..F12/방향키…`)도 동일한 산개 지식 — 명령어 레지스트리와 분리하여 별도 표로 관리할지 §10 미결.

### A.3 본 설계로 추가되는 명령어 (현행 export 미포함)

> 부록 A 본표는 2026-06-24 export의 **현행 30개**다. 아래 2개는 본 설계가 신규 등록하는 lvalue-writer로, MEMORY 카테고리를 1→3개로 채운다(§4.2). 디스패치 구조(4클래스)는 불변, 합계 32개.

| 명령어 | Category | BlockRole | BlockGroup | ColorGroup | PaintScope | P1 | P2 | P3 |
|---|---|---|---|---|---|---|---|---|
| INC | MEMORY | NONE | — | ACTION\* | CMDCELL | 좌변 · **RAW** | — | 증분 · RESOLVE(기본 1) |
| DEC | MEMORY | NONE | — | ACTION\* | CMDCELL | 좌변 · **RAW** | — | 감분 · RESOLVE(기본 1) |

\* ColorGroup은 현행 `SET`과 동일하게 ACTION(연청)을 따른다. MEMORY 전용 색 신설은 `ApplyUnifiedStyle` 팔레트 변경이 필요하므로 별도 결정(§10, 미적용).

**시맨틱 — 정수 read-modify-write (`{$LOOP}`와의 대칭):**

- **대상(P1):** 쓰기 가능한 좌변(`{name}` 또는 `{$CLIP}`). `ParseLValue` 판정을 그대로 재사용 — 읽기전용 토큰(`{$LOOP}`/`{$PIXEL}` 등)을 좌변으로 주면 `CMD_ERROR`. (lvalue 합법성 지식이 `clsMemoryCommands` 안에 있으므로 `SET`과 규칙 공유.)
- **미정의/비숫자 기존값:** `GetVar`의 *missing == empty* 불변식을 **깨지 않는다**. 기존값은 `CLng(Val(...))`로 강제되어 미정의·빈 문자열·비숫자는 **0으로 취급**. 따라서 `INC {new}`(한 번도 안 쓴 변수) → `"1"`. 이 때문에 `EXISTS`/`CLEAR`류가 불필요하다(불변식 보존).
- **증분(P3):** RESOLVE되므로 변수 사용 가능(`INC {total} {delta}`). 생략 시 1. 음수 허용(`INC {i} -2`) → `DEC`는 환원 가능한 슈가(§10).
- **타입/범위:** 정수(Long). 카운터 의미(`{$LOOP}`)와의 대칭을 위해 정수 시맨틱. Long 오버플로(±2³¹)는 기존 코드 관례대로 특별 가드 없음(`On Error Resume Next` 흐름).
- **결과:** 변이값을 `CStr`로 슬롯에 재기록 후 `CMD_NEXT_LINE` 반환.

**예 — `{$LOOP}`로는 불가능한 두 가지(수동 WHILE 카운터 · 회차 무관 누산기):**

```
SET   {errs}  =  0          ' 누산기 초기화
SET   {row}   =  1          ' 수동 카운터 (WHILE에는 내장 카운터가 없음)
WHILE {row}   <=  100
  IF    {$PIXEL:50,{row}}  ==  FF0000
    INC {errs}              ' 루프 구조와 무관한 누적 — {$LOOP}로는 못 함
  END_IF
  INC {row}                 ' WHILE을 N회 루프로 만드는 진행
END_WHILE
NOTIFY  "에러 행 수: {errs}"
```

> `{$LOOP}`는 *가장 안쪽 루프의 회차*만 읽는다. `{errs}`처럼 조건부로만 증가하는 누산기, `{row}`처럼 WHILE을 세는 외부 카운터는 쓰기 가능한 정수 슬롯이 있어야 표현된다 — 이것이 `INC`가 닫는 비대칭이다.
