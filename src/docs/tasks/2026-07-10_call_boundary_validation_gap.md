# CALL 흐름 경계 검사 누락 (2026-07-10)

2026-07-10 core 엔진 워크플로 코드리뷰에서 발견, 수동 재검증으로 CONFIRMED 확정.
당시 리뷰 10건 중 유일한 신규 correctness 버그. 코드 수정 없이 백로그로만 기록한다
(수정 여부/방식은 설계 검토 후 결정).

## 문제

GOTO는 흐름 경계(루프/블록 내부, sub-area 교차)를 넘는 점프를 정적/런타임 양쪽에서
차단하지만, CALL은 양쪽 다 검사가 없다. 동일한 안전망의 비대칭.

| 검사 지점 | GOTO | CALL |
|---|---|---|
| 정적: `clsScriptProgramAnalysis.cls:473` ValidateJumpTargets | `IsGotoAllowed` 검사 | 분기 자체가 없음 (라벨 존재만 확인) |
| 런타임: modEngine | `ResolveGotoTarget`(:409)이 `IsGotoAllowed` + 루프/IF 스택 재조정 | `ResolveCallTarget`(:424)은 라벨 해석만 수행 |

## 실패 시나리오

1. WHILE 본문 안에 LABEL이 있고, 루프 밖에서 그 LABEL을 CALL
   (ValidateStructure/ValidateLabels는 통과. GOTO만 경계 검사하므로)
2. 루프 헤더를 거치지 않고 본문에 진입하므로 LoopStack에 프레임이 쌓이지 않음
3. END_WHILE은 프레임 유무와 무관하게 헤더로 무조건 백점프 (`clsCommandsFlow.cls:151`)
4. CALL이 저장한 호출 시점 루프 깊이(`clsCommandsFlow.cls:79` PushStack)와 현재 깊이가
   어긋나, RETURN에서 "루프 깊이 불일치" 실패 또는 의도치 않은 무한루프

즉 오류가 원인 지점(CALL)이 아니라 한참 뒤(RETURN/END_WHILE)에서 터져 진단이 어렵다.

## 조치안 (향후 구현 시)

- `ResolveCallTarget`에 GOTO와 동일한 `IsGotoAllowed(sourceLine, targetLine)` 검사 추가
- `ValidateJumpTargets`의 GOTO 분기를 CALL에도 확장 (실행 전 정적 차단)
- 선행 설계 검토: CALL 대상은 통상 EXIT 아래 서브루틴 영역(sub-area)이므로,
  `IsGotoAllowed`의 "같은 sub-area 강제" 규칙을 그대로 적용하면 CALL의 존재 이유
  (다른 영역의 서브루틴 호출)와 충돌한다. CALL용 허용 규칙은 "sub-area 교차는 허용,
  루프/블록 본문 중간 진입만 금지" 형태로 분리 정의가 필요할 수 있음
  (`clsScriptProgramIndex.cls:253` IsGotoAllowed, :258 IsLoopTransferAllowed 참고)

## 비고

- 수정 전까지 회피책: 작성 규율로 루프/IF/SELECT 본문 내부의 LABEL을 CALL 대상으로
  쓰지 않는다
- 관련: `2026-07-08_design_review_recommendations.md` B1(CMD_ERROR 사유 채널)이 해결되면
  이 버그의 "뒤늦게 터지는" 증상 진단도 함께 쉬워진다
