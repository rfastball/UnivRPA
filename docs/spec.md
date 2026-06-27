# UnivRPA — FOREACH_ROW 데이터 레이어 설계 결정서

> **목적**: 외부 엑셀 그리드를 순회하며 입력하는 `FOREACH_ROW` 기능의 데이터 레이어 설계 확정본.
> **상태**: 데이터 레이어 *내부* + 인터프리터 *노출 표면*(토큰·블록·매처) 설계 **전부 완료(closed)**. 구현 미착수. 확정 로그는 §11.
> **갱신**: 2세션차 — §9.2 닫힘, §10 작업큐→확정 결과로 전환, §11 확정 로그 신설.
> **전제 코드**: `clsTable.cls` (UnivDB v2.5.1), `modFilterUtils.bas`, UnivRPA 매뉴얼(`ccad021`).
> **이 문서 하나로 다음 세션이 컨텍스트를 복원할 수 있도록 작성됨.**

---

## 0. 한 줄 요약

UnivRPA에는 데이터 레이어가 없어서(전역 스칼라 변수뿐) "정해진 표를 루프 돌며 입력"을 언어 차원에서 표현할 수 없다. 이를 **`clsTable`을 얕게 포크해 채택**하고, **표시 텍스트 변환은 `modFilterUtils.GetSafeFormattedText` 발췌**에 위임하며, 그 위에 **`FOREACH_ROW` 커서**를 얹어 해결한다.

---

## 1. 문제 정의 (왜 이 설계가 필요한가)

인터프리터형 RPA의 정석 패턴은 **Interpreter + Iterator + Context(스코프)**. UnivRPA에 대보면:

| 정석 기둥 | UnivRPA 현황 |
|-----------|--------------|
| Interpreter + Composite (블록 트리) | **있음** — 줄 단위 인터프리터 + 런타임 블록 짝 매칭(IF/END_IF 등) |
| Context / 늦은 바인딩 | **부분** — `SET` 변수 + RESOLVE 토큰 치환 |
| Iterator over DataSource (행→스코프 바인딩) | **없음** ← 이 빈칸이 문제 |

"데이터 레이어"란 *제어 흐름·동작과 분리되어, 구조화된 데이터를 일급 값으로 표현·주소지정·순회하는 부분*. UnivRPA는 변수가 전역 스칼라뿐이라 ① 구조화된 값(행) ② 주소지정(`행.컬럼`) ③ 컬렉션 추상화 ④ 커서 ⑤ 데이터 식이 전부 부재. 그래서 "표를 들고 한 행씩 꺼내 쓰는" 일을 표현할 수단이 없다.

핵심 통찰: **UnivRPA가 접근 가능한 "데이터"는 화면 픽셀·클립보드뿐인데 이는 데이터가 아니라 표현(presentation)이다.** 좌표를 `INC`로 밀어 행을 흉내 내는 트릭은 진짜 행 커서의 부재를 메우는 편법이었다. 소스가 "엑셀 그리드"로 확정되면서, 이미 구조화된 진짜 데이터를 인터프리터에 노출하는 통로만 뚫으면 된다.

---

## 2. 채택 결정: `clsTable` 얕은 포크 (vendoring)

### 2.1 세 갈래 중 B 선택

- **A. 전체 의존** (UnivDB 통째 참조) — ✗ 외부 설치 전제가 UnivRPA의 자기완결성을 깨고, UnivDB의 라이프사이클 규약(`InitDB`→`ClearDB`, "clsTable이 서브루틴 밖에 살면 안 됨")이 FOREACH 커서 수명(블록 스택에 매달려 여러 줄 생존)과 정면 충돌.
- **B. 클래스 단위 차용** ← **채택**. `clsTable.cls` 한 파일만 복사. 파사드·`clsQuery`·`clsReportBuilder`·전역 라이프사이클은 안 가져옴.
- **C. 경량 재구현** — ✗ `LoadFromSheet`가 이미 푼 헤더 식별·빈 셀·중복 헤더 등을 재발명하게 됨.

### 2.2 B가 성립하는 근거 (코드로 검증됨)

`clsTable.cls`는 **의존성 0**. 외부 모듈(`UnivDB`/`clsQuery`/`modLogger`) 호출 없음. 에러는 자체 `Err.Raise 9999`. 외부 의존은 표준 `Scripting.Dictionary`뿐. 자기참조(`GroupBy`/`SelectColumns`가 `New clsTable` 반환)는 한 파일 내 완결. → **한 파일 복사로 컴파일됨.**

### 2.3 "얕은 포크"의 의미

`clsTable`의 정체성(2D 배열 + 헤더 Dictionary + `GetValue`)은 그대로 두고, **데이터 입구(`LoadFromSheet`의 로드 루프)만 RPA용으로 교체**한다(아래 3장). 포크 범위는 `LoadFromSheet` 메서드 하나로 국한. 이후 `GetValue`/`headers`/순회는 전부 무수정으로 동작.

---

## 3. 읽기 규칙: "보이는 그대로" 강제 (핵심 결정)

### 3.1 `Value2` vs `Value` vs `Text` — 정체성 차이

`clsTable`은 원래 `rngTarget.Value2`로 읽는다(DB 엔진으로서 올바름 — 수만 행 조인·집계·정렬에 연산 가능한 원시값이 필요, 서식 변환 오버헤드 회피). 그러나 **RPA의 목적은 정반대**: "화면에 보이던 값을 다른 화면에 그대로 옮겨 친다". `Value2`의 원시값(날짜→시리얼 `45352`, 통화→순수 숫자)은 전부 틀린 답. 규모도 정반대(DB 수만 행 ↔ RPA 수백 행)라 서식 변환 비용이 무의미. → **RPA 맥락에선 다른 읽기 모드가 필요.**

| 방식 | 속도 | 표시 정확도 | 비고 |
|------|------|-------------|------|
| `.Value2` | 빠름·일괄 | ✗ 날짜=시리얼, 서식 소실 | DB용. RPA 부적합 |
| `.Value` | 빠름·일괄 | △ 타입 보존, 표시형식은 로케일 의존 | 타입 우회로 |
| `.Text` | 느림·셀별 | ✓ 보이는 그대로 | `####` 잘림 약점 |

### 3.2 확정 규칙

> **각 셀을 `.Text`로 읽어 화면 렌더 문자열을 그대로 취한다. 타입 판단은 하지 않는다 — 모든 값이 출발부터 문자열.**
> **단, `.Text`가 `####`류 렌더 실패(열 너비 부족)면 그 셀만 `GetSafeFormattedText(cell.Value, cell.NumberFormat)`로 복원한다.** (이 폴백은 셀 너비와 무관하게 올바른 표시문자열 생성)

### 3.3 이 규칙이 공짜로 해결하는 것

- **`"007"` 앞자리 0 보존** — 타입을 안 거치므로 숫자 `7`로 소실될 경로 자체가 없음.
- **날짜 시리얼·통화 원시값 노출 방지** — 화면 텍스트를 그대로 가져오므로.
- **숨은 반올림 자동 정합** — 값 10.5·서식 `0` → 화면 `11` → 입력도 `11`. (`modFilterUtils`의 `fdRounded`/`fdPercentRounded` 개념이 이 케이스의 중요성을 방증)

### 3.4 받아들이는 비용

`.Text`는 배열 일괄 읽기 불가 → 셀별 순회. DB라면 치명적이나 RPA 입력 규모(수백 행)에선 무시 가능. "RPA 규모니 서식째로 뜨는 게 맞다"는 판단의 직접 귀결.

---

## 4. 변환 함수 차용: `modFilterUtils` 발췌

### 4.1 그대로 차용

| 함수 | 역할 |
|------|------|
| `GetSafeFormattedText(vVal, sFmt)` | **변환 엔진 / `####` 폴백의 핵심.** 4단 폴백: `WorksheetFunction.Text` → `VBA Format` → `yyyy-mm-dd` → `CStr`. 1순위가 셀 `.Text`와 동일 렌더 결과를 값+서식 입력으로 생성(너비 무관). `Empty`→빈 문자열, 에러→`CStr` 처리도 우리 규칙과 일치. |

### 4.2 의존성 함정 (주의 — clsTable과 다름)

`modFilterUtils`는 **self-contained 아님**. 모듈 전체에 `Log "...", lvlDebug` 호출이 깔려 있음(외부 로깅 모듈 의존). 그대로 복사 시 `Log`/`lvlDebug` 미정의 컴파일 에러.

→ **대응: `GetSafeFormattedText`만 발췌 + `Log` 줄 제거.** `GetSafeFormattedText`엔 `Log` 1줄(2단계 실패 시)이 있으므로 이 줄만 제거한다. 발췌하면 `ExtractNumberAggressive`(RegExp)·`GetTypeName` 등 내부 의존 사슬에서 깨끗이 끊어짐.

### 4.3 개념만 / 배제

- `GetColumnIdentity` — 컬럼 타입 정밀 판정 엔진. **과함** — "보이는 대로 입력"은 타입 판정 불필요. 단 `fdRounded` 개념은 3.3의 근거로 차용.
- `CleanValue` / `ExtractNumberAggressive` — **역방향**(오염 문자열→원시 숫자). v1 불필요. **v2 되쓰기 정규화에 재활용 가능**.
- `GetUniqueCleanData` / `GetAccountingFilterString` / `GetPercentageFilterString` / `GetCleanFormat` — **전부 엑셀 AutoFilter 매칭 전용**(필터가 값이 아닌 화면 텍스트와 매칭하는 특유 동작 우회). FOREACH와 무관. **안 가져옴.**

---

## 5. 확정된 전체 구조

```
clsTable (얕은 포크)              ← LoadFromSheet 로드 루프만 .Text 기반 교체
   +
modFilterUtils 발췌                ← GetSafeFormattedText
   (Log 호출 제거, 발췌로 의존 사슬 차단)
   +
셀별 원본값/서식 스냅샷           ← .Text 실패 셀의 ####폴백용 cell.Value + cell.NumberFormat
   +
FOREACH 커서 + {$ROW.x} 리졸버    ← 신규. 커서는 GetValue에 위임 → 현재행 정수 r 하나로 환원
```

리졸버 개념식:
```
{$ROW.컬럼} = ( .Text 직독 결과 )
              └ 만약 #### → GetSafeFormattedText(clsTable이 보관한 cell.Value, cell.NumberFormat)
```

---

## 6. clsTable 정합 — 공짜로 따라온 것들

- **헤더 정규화** `m_dicHeaders(UCase(Trim(name)))` → UnivRPA 변수의 **대소문자 무시** 규칙과 이미 일치. `{$ROW.이름}` 추가 정합 불필요. (한글은 `UCase` 무영향)
- **빈 셀** `Empty` → 빈 문자열. UnivRPA "미정의 변수=빈 문자열" 규칙과 일치.
- **행 접근 메서드 기존재** `GetValue(lRow, sColName)`, `GetRecord(lRow)` → **커서 구현 거의 공짜**. 커서 = 현재행 정수 r + END_FOREACH까지 (헤더다음 ~ RowCount) 순회.
- **전체 셀 Trim** `LoadFromArray`가 문자열 셀에 `Trim$` → 앞뒤 공백 제거(대체로 축복). "FOREACH는 소스 텍스트를 trim한다"를 문서화된 동작으로 명시. (`.Text`로 가면 전부 문자열이라 전부 trim되나, 값 변경이 아닌 공백 제거라 무해)

---

## 7. 경계 — 의도적 배제 (v1)

- `clsTable` **변형 메서드**(`SelectColumns`/`RemoveColumn`/`OrderBy`/`DeleteBatch`)는 **호출 안 함**. v1은 비파괴 읽기 전용 순회. (이들은 `Me`를 in-place 변형 — 레퍼런스의 "Copy-on-Write"는 `clsReportBuilder` 계층 얘기, `clsTable` 자체는 파괴적)
- `UnivDB` 파사드·전역 라이프사이클(`InitDB`/`ClearDB`)은 안 가져옴. FOREACH 커서는 `New clsTable` 일회성 인스턴스로 블록 스택 수명에 매닮(템플릿 8.1-A "RegisterTable 불필요, 일회성" 경로).
- `Value` Property **Let**의 silent-fail(컬럼 없으면 `Debug.Print`만) 주의 — 되쓰기 시 `headers.exists` 선검증 필요.

---

## 8. v2가 거의 공짜인 지점 (지금 만들지 않음, 길만 열어둠)

- **되쓰기**: `clsTable.Value` Property Let 기존재 → 메모리 마킹 후 `DumpToSheet`. 외부 워크북도 "읽고 닫은 뒤" 메모리에 들고 있다 원하는 시트에 덤프 가능.
- **순회 전 필터**: `SelectRows`가 O(1) 인덱스 검색으로 이미 지원 → `WHERE` 문법만 얹으면 됨. (비파괴 `SelectRows` ↔ 파괴적 `DeleteBatch` 구분, 비파괴 경로만 사용)
- **역방향 변환**: `CleanValue`/`ExtractNumberAggressive` 재활용.

---

## 9. 워크북 소스 지정 (원칙 + 매처 구현 확정 — 닫힘)

크로스 워크북(UnivRPA가 사는 곳과 *다른* 워크북) 지원은 사용성의 전제. 외부 파일을 라이브로 읽으면 포커스 이탈·스크롤·모달 팝업에 취약 → **"열고 → 배열 스냅샷 → 닫고 → 배열 위 순회"** 모델.

### 9.1 확정 원칙

- **소스 지정자**: 파일을 P1, `시트명` 또는 `시트!주소`를 P2로 **분리**(토큰 치환 살림 — 파일 목록 바깥루프 × 행 안쪽루프 패턴 가능).
- **P2 입력 다형성**: `clsTable.LoadFromSheet`의 기존 입력 형태를 그대로 살린다. `Sheet1`은 시트의 실제 데이터 범위 자동 탐색, `Sheet1!A1`은 A1부터 실제 마지막 데이터까지 자동 탐색, `Sheet1!A1:C10`은 지정 범위 강제 로드.
- **이름→경로 역참조**: 사용자가 경로 대신 **열린 워크북 이름**을 주면 `Workbooks` 컬렉션에서 `.Name`으로 조회 → 경로 은폐 + 파일 열기 생략 동시 달성.
- **이름/경로 자동판별**: 입력값에 구분자(`\` `/` `:`) 있으면 경로 모드, 없으면 열린 워크북 이름 모드. 같은 P1 슬롯이 양쪽 수용.
- **매칭 강도**: **부분 일치 허용**(데이터 파일명이 으레 길어서 — `2024_Q3_거래처별_발주현황_최종_v3.xlsx`). 대소문자 무시(UnivRPA 관습 일치).
- **모호성 처리**(부분 일치의 필수 안전벨트): 0개→하드 에러(경로 폴백 없음), 1개→진행, **2개 이상→에러로 멈추고 후보 FullName 나열**. (ScriptAssist의 "유일할 때만 제안" 철학과 동일)
- **내가 연 것만 닫는다**: 이름 모드(이미 열림)는 `iOpened=False`→안 닫음(사용자 워크북 보호). 경로 모드로 내가 열었으면 닫음. `ReadOnly:=True`, `UpdateLinks:=False`로 조용히 열기.
- **StatusBar 힌트**: `FOREACH_ROW` P1 셀 선택 시 **현재 열린 워크북 목록**을 StatusBar에 우선 표시한다. 그 외 선택은 기존 파라미터 힌트를 유지한다. 비시각 피드백이라 ScriptAssist "서식 안 건드림" 원칙 유지.

### 9.2 확정 (닫힘 — 2세션차)

부분 일치 시 **0매칭 → 하드 에러로 멈춤(경로 폴백 안 함)**. 근거: 이름 모드의 존재 이유는 "이미 열려 있다"는 전제 위의 경로 은폐인데, 0매칭은 그 전제의 퇴화 케이스다. 맨이름(디렉터리·확장자 없음)을 경로로 날조해 여는 것은 정확히 "엉뚱한 파일 열기" 사고를 부른다. 디스크에서 열려면 사용자가 구분자(`\` `/` `:`)를 넣어 **결정론적으로 경로 모드로 라우팅**하면 되므로 탈출구는 이미 명시적이다. "유일할 때만 진행·모호하면 멈춤" 보수성과 일치(StatusBar가 이미 열린 목록을 보여줬으므로 사용자는 후보를 봤음).

### 9.3 매처 구현 확정 (2세션차)

`SnapshotSource(p1, p2) As clsTable` = `ResolveWorkbook(p1, ByRef iOpened)` + `ResolveTableSourceSpec(wb, p2)` + 포크된 `LoadFromSheet`(.Text 스냅샷+셀별 원본값/서식) + `If iOpened Then wb.Close SaveChanges:=False`. `ResolveTableSourceSpec`은 P2를 `Worksheet` 또는 `Range` 객체로 해석하고, 그 객체를 `LoadFromSheet`에 넘긴다. 스냅샷 시점에 `.Text`/`cell.Value`/`cell.NumberFormat`을 셀 단위로 전부 배열 적재 → **닫은 뒤에도** `####`폴백·`{$ROW.x}`가 워크북 없이 동작(close가 스냅샷 *후*라 안전).

- **모드 판별**: `p1`에 `\` `/` `:` 있으면 경로 모드, 없으면 이름 모드.
- **이름 모드**: `Workbooks` 순회 `InStr(1, wb.Name, p1, vbTextCompare)>0` 수집. `iOpened=False`(사용자 것·안 닫음). 0개→에러(§9.2) / 1개→진행 / 2+개→에러+후보 `FullName` 나열.
- **경로 모드**: 열기 *전* basename이 이미 열렸는지 선검사(`Workbooks`에서 `.Name` 매칭). **이 선검사가 핵심 안전장치** — `Workbooks.Open`은 이미 열린 파일에 에러 없이 기존 인스턴스를 반환하므로, 선검사 없이 `iOpened=True`로 박으면 사용자가 띄워둔 워크북을 끝나고 닫아버린다. 열려 있음→보호(`iOpened=False`), 아니면 `Open(ReadOnly:=True, UpdateLinks:=False)`(`iOpened=True`). 엑셀이 동명 동시오픈을 금지하므로 `.Name` 매칭으로 충분. → 두 모드가 같은 불변식("내가 연 것만 닫는다")으로 수렴.
- **`ResolveTableSourceSpec`**: `!`가 없으면 `wb.Worksheets(p2)`를 반환해 시트 전체 실제 데이터 범위를 자동 탐색한다. `!`가 있으면 마지막 `!` 기준 split(따옴표 시트명 안전) → `wb.Worksheets(시트).Range(주소)`를 반환한다. 단일 셀(`A1`)은 시작점 자동 탐색, 명시 범위(`A1:C10`)는 강제 로드. **과대범위(`A:D` 등)는 `Intersect(range, sheet.UsedRange)`로 클램프** — 100만 행 스냅샷 함정 방지. 명명 범위(`wb.Names`)는 v2.

### 9.4 미결 (없음)

§9 전부 닫힘.

---

## 10. 확정 결과 (이번 세션 닫힘 — 인터프리터 노출 표면)

데이터 레이어 *내부*에 더해 *노출 표면*(토큰·블록·매처)까지 전부 확정됨. 아래는 구현 청사진(코드 선례 검증 완료).

### 10.1 `{$ROW.x}` 토큰 문법 — 닫힘

| 토큰 | 해석 | 배치 | 부재 시 |
|------|------|------|---------|
| `{$ROW.컬럼}` | 최근접 FOREACH 프레임 `tbl.GetValue(dataRow, 컬럼)` | `ResolveToken`에 `Left$(up,5)="$ROW."` 분기(=`$PIXEL:` 선례) | 커서 없음→리터럴 통과 / 컬럼 없음→빈값 |
| `{$ROW.{변수}}` | 자식 먼저 해석(`ResolveParam`이 `ResolveToken(ResolveParam(inner))`) 후 동일 | 공짜(추가 코드 0) | 동일 |
| `{$RIDX}` | 최근접 FOREACH 프레임 `iteration`(1-base) | `Select Case`에 `Case "$RIDX"` → `GetActiveForeachOrdinal()` | 커서 없음→리터럴 통과 |
| `{$LOOP}` | **불변** — 최내곽 루프 회차 | 손 안 댐(0줄) | 현행대로 0 |

- **`{$LOOP}`/`{$RIDX}` 관계 = C2(분리, 공짜 일치)로 확정.** `{$LOOP}`은 현행 "최내곽 회차" 의미 그대로. `{$RIDX}`는 "가장 가까운 FOREACH 행". FOREACH 최내곽이면 둘이 자동 동일, 중첩(FOREACH 안 WHILE)이면 각자 올바름 → `{$RIDX}`가 별도 토큰일 이유 성립.
- `GetActiveForeachOrdinal()`은 LoopStack을 **top→down 스캔**해 첫 `kind="FOREACH"` 프레임을 집음(리졸버는 중첩 WHILE 안에서도 fire되므로 top이 아니라 스캔).
- **부재/미스 비대칭(의도)**: 커서 없음(구조 오류)→`$PIXEL` 미지토큰 선례(리졸버 3756–3758행)대로 리터럴 통과로 가시화. 컬럼 없음(데이터 미스)→`GetVar` 선례대로 조용히 빈값.

### 10.2 워크북 매처 — 닫힘

→ §9.2~9.3 참조. `SnapshotSource`가 §10.3 STARTER와의 유일한 seam. 3번은 "테이블을 받거나 에러난다"까지만 알면 닫히고, 0매칭 폴백은 매처 안쪽에 캡슐화됨.

### 10.3 `FOREACH_ROW` 블록 인터프리터 통합 — 닫힘

**등록(3곳, 전부 1줄):**
```vba
Reg "FOREACH_ROW", FLOW, STARTER, "FOREACH", clrFLOW, FULLROW, _
    P("워크북/경로", DT_STR, RESOLVE, True), P("시트 또는 시트!주소", DT_STR, RESOLVE, True)
Reg "END_FOREACH", FLOW, ENDER, "FOREACH", clrFLOW, FULLROW
```
- `clsScriptProgram.FindBlockPair` 초반의 허용 그룹 목록에 `"FOREACH"` **한 단어 추가** — "자동 상속"의 **유일한 예외**. 나머지(들여쓰기·색상·ScriptAssist 검증)는 `STARTER/ENDER`+`BlockGroup`+`clrFLOW`로 따라옴. SELECT는 별도 스캐너라 흉내 불요(FOREACH는 MIDBLOCK 없는 순수 짝→IF/WHILE 경로에 탑승).
- `clsFlowCommands.ICommand_Execute`에 `Case "FOREACH_ROW"`/`Case "END_FOREACH"` 2개 추가. 디스패치는 `Category=FLOW`라 자동 라우팅.
- P1/P2 둘 다 `RESOLVE` → `{변수}` 소스 지정 동작(§9.1 "토큰 치환 살림"). **이 두 줄이 곧 §10.4(레지스트리 엔트리) — 별도 결정 없음(subsumed).**

**실행 의미(WHILE 빈-스킵 + REPEAT 본문-재실행 합성, 둘 다 선례 있음):**

| 줄 | 동작 | 점프 |
|----|------|------|
| `FOREACH_ROW`(STARTER) | `Set tbl = SnapshotSource(args(1),args(2))`. 실패→`CMD_ERROR`. 빈 표(`tbl.RowCount < tbl.FirstDataRow`)→스킵+StatusBar "0행" 힌트. 아니면 `LoopPushForeach(R, end, tbl)`(dataRow=FirstDataRow, iteration=1) | 빈 표: `END_FOREACH+1` / 아니면: 다음 줄 |
| `END_FOREACH`(ENDER) | top 프레임(구조상 이 FOREACH 보장)에서 `dataRow < tbl.RowCount`면 `dataRow++`+`LoopIncrement`. 아니면 `Set tbl=Nothing`+`LoopPop` | 다음 행: **`header+1`** / 소진: 다음 줄 |

⚠️ **로드베어링**: ENDER 백점프는 WHILE처럼 `header`가 아니라 REPEAT처럼 **`header+1`**. `header`(=FOREACH_ROW 줄)로 가면 워크북을 매 행 재오픈/재스냅샷. `header+1`이라야 "열고·닫고"가 1회로 끝나고 본문만 순회.

**LoopFrame 확장(2필드):**
```vba
Private Type LoopFrame
    ' ... 기존 5필드(kind/headerLine/endLine/iteration/maxCount) ...
    tbl As Object       ' FOREACH 전용: 스냅샷 clsTable
    dataRow As Long     ' FOREACH 전용: 현재 테이블 내부 행 인덱스
End Type
```
- `dataRow`를 파생(FirstDataRow+iteration-1) 안 하고 **명시 필드**로 둔 이유: v2 필터(`SelectRows`가 비연속 행 컬렉션 반환, §8)에서 iteration(서수)과 dataRow(배열 인덱스)가 갈라짐. 명시 필드라야 v2가 "dataRow를 필터 리스트에서 집기"만으로 드롭인. v1은 둘이 함께 전진.
- 신규: `LoopPushForeach(header, end, tbl)`(기존 `LoopPush` 호출부 불변), 접근자 `GetActiveForeachOrdinal()`·`ActiveForeachTable()`·`ActiveForeachRow()`(전부 최근접-FOREACH 스캔), `ForeachAdvance()`.
- **`LoopPop`이 FOREACH 프레임 버릴 때 `Set tbl=Nothing` 추가**(현행은 decrement만) → 스냅샷 메모리 즉시 해제.

**BREAK/CONTINUE — 최내곽에 FOREACH 포함: 자동 ✅.** 둘 다 `LoopTopEnd()`/`LoopPop()`을 읽는데 FOREACH가 같은 스택에 올라타므로 최내곽이 FOREACH면 그대로 동작(기존 스택 확장의 직접 배당).

**clsTable 포크 2번째 터치(LoadFromSheet 외 유일):** `Public Property Get FirstDataRow() As Long`(=`m_lHeaderRowIndex+1`) 게터 노출 — 커서가 데이터 시작 위치를 알아야 하는데 `m_lHeaderRowIndex`가 private. 정체성 불변이라 얕은 포크 범위 내.

### 10.4 레지스트리 엔트리 — subsumed (§10.3의 두 Reg 줄)

추가 결정 없음.

### 10.5 커서 중첩 — 닫힘

v1 = 익명 최내곽 평면 한 겹(`{$ROW.x}`/`{$RIDX}`, `{$LOOP}` 미러링). 명명 커서(`{별칭.컬럼}`)는 P3 슬롯 **생략**으로 후방 확장 여지만 남김(빈 슬롯이 ScriptAssist 힌트를 어지럽힘 방지). 현실적 중첩 수요는 *파일들 × 행들* 형태이며 그때 P3 추가.

### 10.6 StatusBar 작성-시점 힌트(§9.1) — 분리 확인

런타임 매처와 독립된 *작성 시점* 기능. `Workbook_SheetSelectionChange`에서 "선택 셀이 FOREACH_ROW 행의 P1인가" 판정 → 열린 워크북 목록을 기존 파라미터 힌트보다 우선 표시. 그 외 셀 선택은 현행 `ShowParamHint` 흐름을 그대로 유지한다. `FindWorkbooksByPartialName` 스캔 재사용하되 modScriptAssist 경로 소속 → `SnapshotSource`와 독립. 구현 시 같이 얹음.

---

## 11. 확정 로그 (세션 결정 추적)

| # | 결정 | 선택 | 근거 한 줄 |
|---|------|------|-----------|
| D1 | `{$LOOP}`↔`{$RIDX}` | **C2 분리** | 기존 `{$LOOP}` 의미 0줄 보존 + FOREACH 최내곽 시 자동 일치 |
| D2 | 커서 상태 보관 | **기존 LoopStack 확장**(+`tbl`,+`dataRow`) | BREAK/CONTINUE/`{$LOOP}` 공짜 상속 (평행 스택 신설 회피) |
| D3 | ENDER 백점프 | **`header+1`**(REPEAT형) | 워크북 재스냅샷 방지(1회 오픈 보장) |
| D4 | `dataRow` 명시 vs 파생 | **명시 필드** | v2 비연속 필터 드롭인 |
| D5 | clsTable 포크 범위 | **LoadFromSheet + FirstDataRow 게터**만 | 정체성 불변, 얕은 포크 유지 |
| D6 | 빈 표 | **본문 스킵 + StatusBar 힌트** | §9.1 비시각 피드백 철학 |
| D7 | v1 P3(별칭) 슬롯 | **생략** | YAGNI, 후방 확장 여지만 |
| D8 | §9.2 0매칭 | **하드 에러(B)** | 이름 모드 전제(이미 열림) 퇴화 시 경로 날조 금지 |
| D9 | 경로 모드 already-open | **선검사 후 보호**(`iOpened=False`) | "내가 연 것만 닫는다" 불변식 수렴 |
| D10 | 과대 범위 | **`Intersect(UsedRange)` 클램프** | 100만 행 스냅샷 함정 방지 |
| D11 | 명명 범위(`wb.Names`) | **v2 연기** | v1은 `시트명`/`시트!주소`만 |

### 다음 세션 = 구현 착수

설계 결정은 전부 닫힘. 남은 것은 코딩: ① clsTable 포크(LoadFromSheet `.Text`화 + 셀별 원본값/서식 스냅샷 + FirstDataRow) ② `modFilterUtils.GetSafeFormattedText` 발췌(Log 제거) ③ LoopFrame 확장+접근자 ④ `clsFlowCommands` 2 Case ⑤ `ResolveToken` `$ROW.`/`$RIDX` 분기 ⑥ `SnapshotSource`/`ResolveWorkbook`/`ResolveTableSourceSpec` ⑦ `FindBlockPair` 화이트리스트 1단어 ⑧ Reg 2줄 ⑨ StatusBar 힌트. 의존 순서: ①②→③→④⑤→⑥→⑦⑧⑨.

---

## 부록 A. 설계 원칙 요약 (한눈에)

1. `clsTable.cls` 얕은 포크 벤더링 — 의존성 0 확인. 입구(LoadFromSheet)만 RPA용 교체.
2. 읽기 = `.Text` 1순위(보이는 그대로, 타입 판단 없음) + `####`면 `GetSafeFormattedText` 폴백.
3. 변환은 `modFilterUtils.GetSafeFormattedText` 발췌(Log 제거)에 위임 — 변환 엔진 자작 안 함.
4. 헤더 대소문자 무시가 UnivRPA 변수 규칙과 우연히 일치 → 추가 정합 불필요.
5. v1은 비파괴 읽기 전용. 되쓰기·필터는 클래스가 이미 지원하므로 v2 확장이 공짜에 가까움.
6. 크로스 워크북은 "열고·스냅샷·닫고·순회". 이름 부분일치 역참조 + 다중매칭 거부로 경로 은폐.
7. `UnivDB` 전역 파사드/라이프사이클은 건드리지 않음 — 일회성 `clsTable` 인스턴스만.

## 부록 B. 차용 출처 매핑

| 우리가 필요한 것 | 출처 | 방식 |
|------------------|------|------|
| 데이터 컨테이너(2D배열+헤더Dict) | `clsTable` | 얕은 포크 |
| 스냅샷(시트/범위→배열) | `clsTable.LoadFromSheet` | 로드 루프만 교체 |
| 헤더→인덱스 맵 | `clsTable.headers` | 그대로 |
| 행 값 접근 | `clsTable.GetValue` | 그대로 |
| 표시 텍스트 변환 / ####폴백 | `modFilterUtils.GetSafeFormattedText` | 발췌(Log제거) |
| 워크북 이름 역참조 | — | 신규 |
| FOREACH 커서/블록 | — | 신규 |
| `{$ROW.x}` 리졸버 | `clsTable.GetValue` 활용 | 얇은 래퍼(신규) |
| (v2) 순회 전 필터 | `clsTable.SelectRows` | 문법만 신규 |
| (v2) 되쓰기 | `clsTable.Value Let` + `DumpToSheet` | 문법만 신규 |
| (v2) 역방향 정규화 | `modFilterUtils.CleanValue` | 재활용 |
