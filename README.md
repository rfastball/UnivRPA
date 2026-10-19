# UnivRPA

**자동화 한 줄이 곧 Excel의 한 행.**

Excel 시트에 명령어를 적고 단축키로 실행하는 저코드 업무 자동화 프레임워크입니다.
마우스 클릭, 키보드 입력, 창 활성화, 픽셀 색상 대기를 6열 테이블로 구성합니다.

[시작하기](#시작하기) · [명령어 레퍼런스](src/docs/reference/02-command-reference.md) · [작성 보조 도구](src/docs/reference/03-authoring-tools.md)

## 시트가 곧 스크립트

버튼을 클릭하고, 지정한 픽셀이 흰색이 되면 이름을 입력하는 예제입니다.
좌표와 색상은 자동화할 화면에 맞게 바꿉니다.

| ROW | COMMAND | PARAM1 | PARAM2 | PARAM3 | COMMENT |
| --- | --- | --- | --- | --- | --- |
| 1 | CLICK | 542 | 375 | 500 | 클릭 후 0.5초 대기 |
| 2 | WAIT_APPEAR | 100 | 100 | FFFFFF | 지정 픽셀의 흰색 대기 |
| 3 | TYPE | 50 | | 홍길동 | 이름 입력 |
| 4 | NOTIFY | 3 | | 완료! | 3초간 알림 |

A열은 줄번호, B열은 명령어, C~E열은 파라미터, F열은 메모입니다.
`TYPE`의 입력 내용처럼 **명령어마다 값을 넣는 파라미터 위치가 다릅니다.**
셀을 선택하면 상태표시줄에서 해당 명령어의 파라미터 힌트를 확인할 수 있습니다.

| 할 일 | UnivRPA에서 쓰는 방식 |
| --- | --- |
| 화면 조작 | 좌표로 클릭하고 키·문자열 입력 |
| 화면 변화 대기 | 지정 좌표의 픽셀 색상으로 동기화 |
| 조건·반복·재사용 | IF, WHILE, REPEAT, CALL 등의 흐름 제어 |
| Excel 데이터 입력 | 표의 행을 순회하며 헤더 이름으로 값 참조 |
| 스크립트 작성 | 파라미터 힌트, 구조·라벨 검증, 자동 서식, 좌표·색상 캡처 |

## 시작하기

**실행 환경:** Windows 데스크톱 Excel과 매크로 실행이 가능한 `UnivRPA.xlsm`이 필요합니다.
VBA와 Win32 API를 사용하며, 별도의 RPA 런타임 없이 Excel에서 실행합니다.
좌표·픽셀 기반이므로 대상 창의 위치나 화면 배율이 달라지면 값을 다시 확인해야 합니다.

> **이 저장소에는 실행 워크북이 포함되어 있지 않습니다.**
> `.xlsm`은 Git 관리 대상에서 제외됩니다. 아래 절차는 소스와 폼이 준비된
> `UnivRPA.xlsm`을 보유한 경우를 기준으로 합니다. 저장소를 복제하는 것만으로 실행되지는 않습니다.

1. `UnivRPA.xlsm`을 열고 매크로를 허용합니다.
2. 스크립트 시트의 A1:F1에 `ROW`, `COMMAND`, `PARAM1`, `PARAM2`, `PARAM3`, `COMMENT`를 넣고, 2행부터 명령어를 입력합니다.
3. `Ctrl+Shift+E`로 캡처 HUD를 열어 대상 화면의 좌표·색상을 확인합니다.
4. `Ctrl+Shift+F`로 줄번호와 서식을 정리합니다.
5. 스크립트 시트를 활성화한 상태에서 `Ctrl+Shift+W`로 실행합니다.

작성 보조는 **A1=`ROW`, B1=`COMMAND`**인 시트에서 동작합니다.
자세한 화면 안내와 예제는 [시작하기 매뉴얼](src/docs/reference/01-getting-started.md)을 참고하세요.

| 단축키 | 기능 |
| --- | --- |
| `Ctrl+Shift+W` | 스크립트 실행 |
| `Ctrl+Shift+F` | 서식·줄번호·들여쓰기 정리 |
| `Ctrl+Shift+/` | 선택 줄 주석 토글 |
| `Ctrl+Shift+E` | 좌표·색상 캡처 HUD |
| `Pause` | 실행 중 일시정지·재개 |
| `Esc` | 실행 중 중단 |

## Excel 데이터를 한 행씩 입력하기

`FOREACH_ROW`는 첫 행을 헤더로 읽고, 이후 데이터 행을 순회합니다.
다음 예제는 고객 목록의 이름과 전화번호를 대상 화면에 입력합니다.
실행 전에 입력할 창과 첫 입력 칸에 포커스가 오도록 준비해야 합니다.

| ROW | COMMAND | PARAM1 | PARAM2 | PARAM3 | COMMENT |
| --- | --- | --- | --- | --- | --- |
| 1 | FOREACH_ROW | 고객목록.xlsx | Sheet1!A1:C100 | VISIBLE_ROWS | 보이는 행만 순회 |
| 2 | TYPE | 50 | | {$ROW.이름} | 이름 입력 |
| 3 | KEY | TAB | | | 다음 입력 칸 |
| 4 | TYPE | 50 | | {$ROW.전화번호} | 전화번호 입력 |
| 5 | KEY | TAB | | | 다음 이름 칸 |
| 6 | END_FOREACH | | | | 반복 끝 |

이 예제는 `Tab`으로 이름 → 전화번호 → 다음 이름 칸으로 이동하는 화면을 가정합니다.
실제 화면에 맞게 이동·저장 명령을 구성하세요.

- `{$ROW.헤더명}`: 현재 행의 값
- `{$RIDX}`: 현재 행 순번
- `{$RTOTAL}`: 전체 순회 행수

워크북 지정 방식과 순회 모드는 [명령어 레퍼런스](src/docs/reference/02-command-reference.md)에서 확인할 수 있습니다.

## 5개 카테고리, 36개 명령어

| 카테고리 | 명령어 |
| --- | --- |
| 흐름 제어 · 19 | IF · ELSE_IF · ELSE · END_IF · WHILE · END_WHILE · REPEAT · END_REPEAT · FOREACH_ROW · END_FOREACH · SELECT · CASE · END_SELECT · BREAK · CONTINUE · GOTO · CALL · RETURN · LABEL |
| 동작 · 6 | MOVE · CLICK · RCLICK · KEY · COMBO · TYPE |
| 메모리 · 3 | SET · INC · DEC |
| 시스템 · 4 | FOCUS · SLEEP · NOTIFY · EXIT |
| 비전 · 4 | WAIT_APPEAR · WAIT_VANISH · CLICK_SYNC · ASSERT |

명령어 구문·조건 연산자·변수는 [02. 명령어 레퍼런스](src/docs/reference/02-command-reference.md),
포맷터·힌트·캡처 사용법은 [03. 작성 보조 도구](src/docs/reference/03-authoring-tools.md)에 정리되어 있습니다.

## 개발자를 위한 구조

실행 흐름은 **시트 스냅샷 → 구조·라벨 검증 → 명령어 조회·파라미터 해석 → 핸들러 실행**입니다.
검증에 실패하면 실행을 거부하고, 실행 중에는 각 핸들러가 다음 줄·점프·중단·종료를 결정합니다.

| 구성 요소 | 역할 |
| --- | --- |
| `modCmdRegistry` | 명령어 정의와 카테고리별 핸들러 연결 |
| `clsScriptProgram` | 원본·블록·라벨·분석을 제공하는 정적 모델 |
| `modEngine` | 줄 단위 해석과 실행 흐름 |
| `clsCommands*` | `ICommand`를 구현한 5개 카테고리 핸들러 |
| `modContext` | 변수·콜스택·실행 상태 |
| `modActions` / `modWinAPI` | 구체 동작과 Win32 호출 |

편집할 VBA 소스는 `src/`, 사용자 문서는 `src/docs/reference/`에 있습니다.
레이어 규율과 명령어 추가 절차는 [AGENTS.md](AGENTS.md)를 참고하세요.

개발용 워크북을 준비한 환경에서는 `vbaloop`로 소스 주입과 테스트를 수행합니다.

```powershell
vbaloop init --check
vbaloop push
vbaloop test
```

`push`는 `src/`를 워크북에 반영하고, `test`는 인자 없는 `Public Sub Test_*()`를 실행합니다.
소스는 UTF-8(BOM 없음)·CRLF 규약을 따릅니다.
`src/LogShim.bas`는 vbaloop 관리 사본이므로 직접 수정하지 않습니다.

## 라이선스

[MIT License](LICENSE) · Copyright (c) 2026 JM
