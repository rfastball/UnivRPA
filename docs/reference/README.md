# UnivRPA 사용자 매뉴얼

> 기준 소스 버전: `ccad021`

UnivRPA는 **Excel 시트를 IDE처럼 사용하는 저코드 RPA(업무 자동화) 프레임워크**입니다.
스크립트는 6열짜리 명령어 테이블로 작성하고, 단축키 한 번으로 줄 단위 실행합니다.
프로그래밍 지식 없이 스프레드시트의 행·열만으로 마우스·키보드·창 제어·화면(픽셀) 동기화를 자동화할 수 있습니다.

자동화 한 줄은 한 행(row)이며, 셀에 명령어와 파라미터를 채우는 방식입니다.

```
A=ROW   B=COMMAND   C=PARAM1   D=PARAM2   E=PARAM3   F=COMMENT
 1       CLICK       542        375        500        # 버튼 클릭
```

---

## 목차

| 문서 | 내용 |
|------|------|
| **[01. 시작하기](01-getting-started.md)** | 시트 구조, 스크립트 작성·실행, 상태 폼, 좌표 캡처, 첫 스크립트 따라하기 |
| **[02. 명령어 레퍼런스](02-command-reference.md)** | 5개 카테고리 33개 명령어 전체, 조건 연산자, 토큰·변수 시스템 |
| **[03. 작성 보조 도구](03-authoring-tools.md)** | 고정 단축키, 포맷터, 실시간 검증·힌트, 주석 토글, 색상 캡처 HUD |

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

> 스크립트 시트 판별 시그니처: **A1 = `ROW`**, **B1 = `COMMAND`** 를 동시에 만족하는 시트에서만 작성 보조 기능이 동작합니다.

### 고정 단축키

| 단축키 | 기능 |
|--------|------|
| `Ctrl`+`Shift`+`W` | RPA 실행 |
| `Ctrl`+`Shift`+`F` | 스크립트 포맷(서식·줄번호·들여쓰기) |
| `Ctrl`+`Shift`+`/` | 선택 줄 주석 토글 |
| `Ctrl`+`Shift`+`E` | 색상·좌표 캡처 HUD |

> 실행 중 제어: `Pause` = 일시정지/재개, `Esc` = 중단

### 명령어 카테고리 (총 33개)

| 카테고리 | 개수 | 명령어 |
|----------|------|--------|
| 흐름 제어 (FLOW) | 16 | IF · ELSE · END_IF · WHILE · END_WHILE · REPEAT · END_REPEAT · SELECT · CASE · END_SELECT · BREAK · CONTINUE · GOTO · CALL · RETURN · LABEL |
| 동작 (ACTION) | 6 | MOVE · CLICK · RCLICK · KEY · COMBO · TYPE |
| 메모리 (MEMORY) | 3 | SET · INC · DEC |
| 시스템 (SYSTEM) | 4 | FOCUS · SLEEP · NOTIFY · EXIT |
| 비전 (VISION) | 4 | WAIT_APPEAR · WAIT_VANISH · CLICK_SYNC · ASSERT |

자세한 구문과 예제는 [02. 명령어 레퍼런스](02-command-reference.md)를 참고하세요.
