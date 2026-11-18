# 값 보존과 편집 오류 및 Pause 수정 계획

기준: 작업 트리 깨끗함. 직전 감사의 기존 테스트 21개 통과, 워크북과 src 실질 차이 없음.

## 계획

- [x] 1. 스크립트 스냅샷의 Currency 셀만 Value2로 보정하여 정밀도와 기존 날짜 해석을 함께 보존한다.
- [x] 2. HUD 자동 입력의 색상 셀을 텍스트로 기록하여 앞자리 0과 지수처럼 보이는 색상을 보존한다.
- [x] 3. 작성 보조 스냅샷 오류를 상태표시줄로 안내하고 정상 편집으로 복구한다. 실행 시 오류값 거부는 유지한다.
- [x] 4. 숫자 범위 초과를 기본값으로 대체하지 않고 오류로 처리한다. 비숫자 현재값의 기존 기본값 계약은 유지한다.
- [x] 5. 비전 대기의 제한시간에서 실제 Pause 시간을 제외한다. 중단과 실제 제한시간 도달은 계속 유효하다.
- [x] 기존 소유 테스트 확장, 사용 계약 문서 갱신, lint/push/관련 테스트/전체 테스트/diff 및 lease 해제.

## Test delta brief

- Unique risk: 통화 서식 파라미터의 숫자 손실, HUD 색상 자동 숫자 변환, 오류 수식이 있는 시트의 편집 이벤트 예외, 정수 범위 초과 시 변수 덮어쓰기, Pause 직후 비전 시간 초과. 기존 테스트는 표의 통화 서식과 배열 입력, 실행 거부만 확인하고 실제 편집 이벤트 및 Pause 시간이 흐르는 조건은 검증하지 않는다.
- Existing owner: modTester.TC_EngineLifecycle, TC_ScriptAssistFormatterRegression, TC_MemoryResolverHeadless, TC_AbortLatch.
- Selected layer: 기존 임시 시트로 Excel 서식/이벤트 경계를 검증한다. 메모리는 배열 기반 실행, Pause는 기존 clsActionTestDriver에 가상 시간과 키 상태를 추가하여 실제 대기 없이 검증한다. HUD 셀 기록은 폼 표시 없는 시트 어댑터로 호출한다.
- Expected cost delta: 공개 Test_* 및 테스트 파일 추가 0개, 추가 Excel/외부 프로세스 0개. 실제 키 입력과 장시간 대기 0회. 기존 임시 시트를 재사용한다.
- Evidence to retire: 기존 검증은 서로 다른 정상/실패 경로를 지키므로 유지한다. 같은 경로의 별도 스모크나 소스 문자열 검사는 추가하지 않는다.

## 결과

5개 수정과 사용 계약 문서 갱신을 완료하고 UnivRPA.xlsm에 반영했다.

- `vbaloop init --check --json`: 통과.
- `vbaloop lint --json`: 41개 파일, 지적 0개.
- `vbaloop push --json`: VBE 컴파일 통과, 워크북 저장. 백업 `20261117-221703-521-push`.
- `vbaloop run modTester.Test_AbortLatch --json`: 실패 단언 0개, 매크로 실행 549ms.
- `vbaloop test --json`: 21개 통과, 실패/건너뜀 0개. 명령 전체 약 8.3초(1회 관측, 성능 비교 아님).
- `vbaloop diff --json`: clean=true. 실질 차이 0개, VBE 식별자 대소문자 차이 6개.
- `git diff --check`: 통과.
- `vbaloop lease release --json`: released.

회귀 검증은 기존 4개 소유 테스트를 확장했다. 통화 정밀도와 날짜 호환, HUD 색상의 문자열 보존, 실제 시트 편집 이벤트에서 수식 오류와 복구, 정수 범위 초과 시 변수 보존, TYPE 지연값 범위 초과 시 키 입력 차단을 확인했다. 비전 3개 명령은 60초 가상 Pause 후 재개, Pause 없는 30초 제한, Pause 중 중단을 확인했다.

테스트 비용 변화: 공개 Test_* 및 테스트 파일 증가 0개, 전체 스위트의 Excel/외부 프로세스 증가 0개, 새 임시 시트 0개, 실제 장시간 대기 0회. 기존 검증은 다른 계약을 보호하므로 삭제하지 않았다. 실행시간 증감은 측정하지 않았다.

검증 한계: 실제 대상 프로그램에 대한 마우스/키보드 및 HUD 수동 조작은 이번 자동 검증에 포함하지 않았다.
