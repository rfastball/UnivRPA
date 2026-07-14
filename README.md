# UnivRPA

**Excel 시트를 IDE처럼 사용하는 저코드 RPA(업무 자동화) 프레임워크**

VBA 매크로로 구현된 순수 Excel 워크북(`UnivRPA.xlsm`) 하나로 동작한다. 스크립트는
6열짜리 명령어 테이블로 작성하고, 단축키 한 번으로 줄 단위 실행한다.
프로그래밍 지식 없이 스프레드시트의 행·열만으로 마우스·키보드·창 제어·화면(픽셀)
동기화를 자동화할 수 있다.

```
A=ROW   B=COMMAND   C=PARAM1   D=PARAM2   E=PARAM3   F=COMMENT
 1       CLICK       542        375        500        # 버튼 클릭
```

## 특징

- **저코드**: 셀에 명령어와 파라미터만 채우면 자동화 한 줄 완성. 자동완성 없이도
  암기 가능한 36개 명령어셋
- **흐름 제어**: `IF`/`WHILE`/`REPEAT`/`SELECT`/`FOREACH_ROW`/`CALL`/`GOTO` 등
  중첩 가능한 블록 구조와 서브루틴을 지원하는 완전한 인터프리터
- **화면 동기화**: 픽셀 색상 기반 `WAIT_APPEAR`/`WAIT_VANISH`/`CLICK_SYNC`/`ASSERT`로
  타이밍에 의존하지 않는 안정적인 자동화
- **데이터 순회**: `FOREACH_ROW`로 열린 워크북 또는 외부 워크북의 표를 한 줄씩 순회,
  `{$ROW.헤더명}` 토큰으로 값 참조
- **작성 보조**: 실시간 서식·검증·힌트, 좌표/색상 캡처 HUD(`Ctrl+Shift+E`),
  라벨 이동, 주석 토글 등 시트 안에서 완결되는 개발 경험

## 빠른 시작

1. `UnivRPA.xlsm`을 연다.
2. 빈 시트에 `ROW | COMMAND | PARAM1 | PARAM2 | PARAM3 | COMMENT` 형식으로 명령을 입력한다.
3. `Ctrl+Shift+F` : 스크립트 서식 정리(헤더 생성, 줄번호, 들여쓰기, 색상).
4. `Ctrl+Shift+W` : 실행. 실행 중 `Pause`로 일시정지/재개, `Esc`로 중단.

자세한 안내는 [`src/docs/reference/01-getting-started.md`](src/docs/reference/01-getting-started.md)
를 참고한다.

## 명령어 카테고리 (총 36개)

| 카테고리 | 개수 | 명령어 |
|----------|------|--------|
| 흐름 제어 (FLOW) | 19 | IF · ELSE_IF · ELSE · END_IF · WHILE · END_WHILE · REPEAT · END_REPEAT · FOREACH_ROW · END_FOREACH · SELECT · CASE · END_SELECT · BREAK · CONTINUE · GOTO · CALL · RETURN · LABEL |
| 동작 (ACTION) | 6 | MOVE · CLICK · RCLICK · KEY · COMBO · TYPE |
| 메모리 (MEMORY) | 3 | SET · INC · DEC |
| 시스템 (SYSTEM) | 4 | FOCUS · SLEEP · NOTIFY · EXIT |
| 비전 (VISION) | 4 | WAIT_APPEAR · WAIT_VANISH · CLICK_SYNC · ASSERT |

전체 구문과 예제는 [`src/docs/reference/02-command-reference.md`](src/docs/reference/02-command-reference.md)
를 참고한다.

## 문서

| 문서 | 내용 |
|------|------|
| [01. 시작하기](src/docs/reference/01-getting-started.md) | 시트 구조, 스크립트 작성·실행, 상태 폼, 좌표 캡처, 첫 스크립트 따라하기 |
| [02. 명령어 레퍼런스](src/docs/reference/02-command-reference.md) | 5개 카테고리 36개 명령어 전체, 조건 연산자, 토큰·변수 시스템, 행 단위 데이터 순회 |
| [03. 작성 보조 도구](src/docs/reference/03-authoring-tools.md) | 고정 단축키, 포맷터, 실시간 검증·힌트, 라벨 이동, 주석 토글, 색상 캡처 HUD |

## 아키텍처

`RunRPA`(`modEngine`, `Ctrl+Shift+W`) 기준 한 실행 사이클:

1. ActiveSheet 스냅샷 → `clsScriptProgram.LoadFromRange`로 ProgramImage 승격
2. 구조/라벨 정적 검증 (실패 시 실행 거부)
3. 줄 루프: 명령어 레지스트리 조회 → 파라미터 디코드(RAW/RESOLVE) →
   카테고리별 `ICommand` 핸들러로 2단 디스패치
4. 핸들러 반환값으로 다음 줄 결정(다음 줄/점프/중단/정상 종료)

레이어는 상위가 하위만 호출하는 단방향 규율을 따른다.

| Layer | 모듈 | 역할 |
|---|---|---|
| 0 | `modCmdRegistry` + `CmdDef`/`ParamSpec` | 명령어 지식의 단일 출처(SSOT) |
| 1 | `modEngine` | 인터프리터 (API/변수/동작 직접 수행 금지) |
| 2 | `modContext` | 변수(Memory)·콜스택·실행 상태 |
| 3 | `modActions` | 구체 동작(클릭·입력·대기) |
| 4 | `modWinAPI` | Win32 래퍼(64비트 PtrSafe, 무상태) |

명령어 핸들러 5클래스(`clsCommandsFlow`/`Action`/`Memory`/`System`/`Vision`)는
모두 `ICommand`를 구현하며, 레지스트리가 카테고리별로 1회 배선한다.

## 개발

VBA 소스는 `src/*.bas` · `src/*.cls` · `src/*.frm`에 있으며, 워크북과의 동기화는
`vbaflow` CLI(`vbaflow.toml` 참조)로 수행한다:

```
vbaflow session start   # 세션용 Excel 오픈(멱등)
vbaflow push            # src -> 워크북 in-place 주입
vbaflow test             # Test_* 전체 실행 + 결과 판독
```

테스트는 `Public Sub Test_*()`(인자 없음)를 `vbaflow test`가 열거 실행하며,
단언은 `LogShim.AssertEq`를 사용한다. 자세한 절차는 프로젝트 루트의
`CLAUDE.md`/`AGENTS.md`를 참고한다.
