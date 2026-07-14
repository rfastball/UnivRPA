# UnivRPA

**Excel 시트를 IDE처럼 사용하는 저코드 RPA(업무 자동화) 프레임워크.**

스크립트는 6열짜리 명령어 테이블로 작성하고, 단축키 한 번으로 줄 단위 실행합니다.
프로그래밍 지식 없이 스프레드시트의 행·열만으로 마우스·키보드·창 제어·화면(픽셀)
동기화를 자동화할 수 있습니다. 자동화 한 줄이 곧 한 행(row)입니다.

```
A=ROW   B=COMMAND   C=PARAM1   D=PARAM2   E=PARAM3   F=COMMENT
 1       CLICK       542        375        500        # 버튼 클릭
```

- **런타임 없음** — 순수 Excel VBA. `.xlsm` 워크북 하나로 동작하며 별도 설치가 필요 없습니다.
- **저코드** — 5개 카테고리 36개 명령어. 셀에 명령어와 파라미터를 채우는 방식.
- **작성 보조 내장** — 실시간 파라미터 힌트, 구조·라벨 검증, 자동 포맷·색상 서식, 좌표·색상 캡처 HUD.

---

## 한눈에 보기

### 6열 구조

| 열 | 헤더 | 용도 |
|----|------|------|
| A | `ROW` | 자동 줄번호 (포맷터가 생성) |
| B | `COMMAND` | 명령어 (CLICK, IF, WHILE, TYPE …) |
| C | `PARAM1` | 첫 번째 파라미터 |
| D | `PARAM2` | 두 번째 파라미터 |
| E | `PARAM3` | 세 번째 파라미터 |
| F | `COMMENT` | 자유 메모 (실행 시 무시) |

> 스크립트 시트 판별 시그니처: **A1 = `ROW`**, **B1 = `COMMAND`** 를 동시에
> 만족하는 시트에서만 작성 보조 기능이 동작합니다.

### 고정 단축키

| 단축키 | 기능 |
|--------|------|
| `Ctrl`+`Shift`+`W` | RPA 실행 |
| `Ctrl`+`Shift`+`F` | 스크립트 포맷(서식·줄번호·들여쓰기) |
| `Ctrl`+`Shift`+`/` | 선택 줄 주석 토글 |
| `Ctrl`+`Shift`+`E` | 색상·좌표 캡처 HUD |

> 실행 중 제어: `Pause` = 일시정지/재개, `Esc` = 중단

### 명령어 카테고리 (총 36개)

| 카테고리 | 개수 | 명령어 |
|----------|------|--------|
| 흐름 제어 (FLOW) | 19 | IF · ELSE_IF · ELSE · END_IF · WHILE · END_WHILE · REPEAT · END_REPEAT · FOREACH_ROW · END_FOREACH · SELECT · CASE · END_SELECT · BREAK · CONTINUE · GOTO · CALL · RETURN · LABEL |
| 동작 (ACTION) | 6 | MOVE · CLICK · RCLICK · KEY · COMBO · TYPE |
| 메모리 (MEMORY) | 3 | SET · INC · DEC |
| 시스템 (SYSTEM) | 4 | FOCUS · SLEEP · NOTIFY · EXIT |
| 비전 (VISION) | 4 | WAIT_APPEAR · WAIT_VANISH · CLICK_SYNC · ASSERT |

---

## 예제

버튼 클릭 → 화면이 흰색으로 바뀔 때까지 대기 → 이름 입력 → 확인 → 완료 알림:

| ROW | COMMAND | PARAM1 | PARAM2 | PARAM3 | COMMENT |
|-----|---------|--------|--------|--------|---------|
| 1 | CLICK | 542 | 375 | 500 | # 버튼 클릭 후 0.5초 대기 |
| 2 | WAIT_APPEAR | 100 | 100 | FFFFFF | # 대화상자(흰색)가 뜰 때까지 대기 |
| 3 | TYPE | 50 |  | 홍길동 | # 50ms 간격으로 이름 입력 |
| 4 | CLICK | 680 | 400 | 500 | # 확인 클릭 |
| 5 | SLEEP |  |  | 2000 | # 2초 대기 |
| 6 | NOTIFY | 3 |  | 완료! | # 3초간 완료 알림 |

데이터 행 순회 — 표나 외부 워크북을 한 줄씩 처리:

```
FOREACH_ROW | 고객목록.xlsx | Sheet1!A1:C100 | VISIBLE_ROWS
  TYPE   | 50 |  | {$ROW.이름}
  KEY    | TAB
  TYPE   | 50 |  | {$ROW.전화번호}
END_FOREACH
```

반복 안에서 `{$ROW.헤더명}`으로 현재 행의 값을, `{$RIDX}`로 현재 순번을,
`{$RTOTAL}`로 전체 행수를 읽습니다.

---

## 시작하기

1. `UnivRPA.xlsm` 을 엽니다 (매크로 사용 허용).
2. 빈 시트에 명령어를 입력하고 `Ctrl`+`Shift`+`F` 로 포맷합니다.
3. 좌표·색상은 `Ctrl`+`Shift`+`E` 캡처 HUD로 채웁니다.
4. `Ctrl`+`Shift`+`W` 로 실행합니다.

전체 절차는 사용자 매뉴얼을 참고하세요:

| 문서 | 내용 |
|------|------|
| [01. 시작하기](src/docs/reference/01-getting-started.md) | 시트 구조, 작성·실행, 상태 폼, 좌표 캡처, 첫 스크립트 |
| [02. 명령어 레퍼런스](src/docs/reference/02-command-reference.md) | 36개 명령어 전체, 조건 연산자, 토큰·변수, 행 순회 |
| [03. 작성 보조 도구](src/docs/reference/03-authoring-tools.md) | 단축키, 포맷터, 실시간 검증·힌트, 라벨 이동, 캡처 HUD |

---

## 아키텍처

`RunRPA`(Ctrl+Shift+W) 기준 한 실행 사이클:

1. ActiveSheet 스냅샷 → `clsScriptProgram` 정적 모델로 승격
2. 구조·라벨 정적 검증 (실패 시 실행 거부)
3. 줄 루프: 명령어 정규화·조회 → 파라미터 디코드 → 2단 디스패치(카테고리 → 핸들러)
4. 핸들러 반환값이 다음 줄번호 (다음 줄 / 점프 / 중단 / 종료)

레이어 규율 (위가 아래만 호출):

| Layer | 모듈 | 역할 |
|---|---|---|
| 0 | `modCmdRegistry` (+`CmdDef`/`ParamSpec`) | 명령어 지식의 SSOT |
| 1 | `modEngine` | 인터프리터 |
| 2 | `modContext` | 변수·콜스택·실행 상태 |
| 3 | `modActions` | 구체 동작(클릭·입력·대기) |
| 4 | `modWinAPI` | Win32 래퍼 (64비트 PtrSafe) |

명령어 핸들러 5클래스(`clsCommandsFlow/Action/Memory/System/Vision`)는 모두
`ICommand` 를 구현하고 레지스트리가 카테고리별로 배선합니다. 정적 모델
`clsScriptProgram` 은 파사드이며, 소스·인덱스·분석 세 서비스로 분할됩니다.

---

## 프로젝트 구조

```
UnivRPA/
├─ src/                  VBA 소스 (.bas/.cls/.frm) — 편집·버전 관리 대상
│  └─ docs/reference/    사용자 매뉴얼 (01~03)
├─ UnivRPA.xlsm          대상 워크북 (런타임 아티팩트, git 미추적)
├─ CLAUDE.md / AGENTS.md 아키텍처·개발 규약
└─ UnivRPA_DependencyGraph.md  모듈 의존 그래프
```

`.xlsm` 워크북은 실행 아티팩트라 버전 관리에서 제외되며, 곁의 `src/` 가
진실의 원본입니다. VBA 소스는 UTF-8(BOM 없음) + CRLF 규약을 따릅니다.
