# 감사 후속 수정 계획 (2026-11-17)

범위: 후속 감사 8개 항목. 기존 작업 트리 변경을 보존한다.

## 구현 계획

- [x] 1: ASSERT 색상을 정규화하고 잘못된 16진수 색상을 거부한다.
- [x] 2: 토큰 치환 후 빈 FOCUS 창 제목을 거부한다.
- [x] 3: 클립보드 읽기 실패를 전달하고 모든 경로에서 잠금과 핸들을 정리한다.
- [x] 4: 커서 이동 실패 시 클릭 전에 중단한다.
- [x] 5: 실행 중 도달한 CASE는 값 해석 없이 END_SELECT로 이동한다.
- [x] 6: IF 조회와 제거를 현재 CALL의 프레임으로 제한한다.
- [x] 7: 루프 제어와 재진입을 현재 CALL의 프레임으로 제한한다.
- [x] 8: 기존 레지스트리 스위트에서 명령별 실제 실행 분기를 검증한다.
- [x] 관련 사용 계약 문서 갱신, lint, push 컴파일, 좁은 테스트, 전체 테스트, diff.

## Test delta brief

- Unique risk: 잘못된 화면 검사 통과, 빈 창 제목의 오입력, 클립보드 실패의 빈값 변환, 커서 이동 실패 후 클릭, 건너뛴 CASE 값 해석, CALL 간 IF/루프 상태 오염, 미구현 명령의 테스트 통과.
- Existing owner: modTester.TC_AbortLatch, TC_FlowHappyPaths, TC_FlowNestedInteractions, TC_RegistryDispatchSync 및 clsActionTestDriver.
- Selected layer: 배열 기반 실행 및 컨텍스트 검증, 기존 드라이버의 API 실패 주입. 실제 클릭/타이핑/클립보드 변경 없이 검증한다.
- Expected cost delta: 새 테스트 파일과 공개 Test_* 진입점 없음. 기존 Excel 세션을 재사용하며 실시간 대기와 외부 프로세스 추가 없음.
- Evidence to retire: 단순 흰색 ASSERT 성공 검증을 정규화/금지색 사례로 확장하고, 레지스트리의 핸들러 존재만 확인하던 검증을 실제 분기 검증으로 강화한다. 기존 서로 다른 실패 경로 검증은 유지한다.

## 검증 결과

작업 전 기준: 기존 테스트 21개 통과, 실패/건너뜀 0. 소스와 워크북 일치.

- `vbaloop init --check`: 통과, 실행 파일과 관리 버전 0.6.0 일치.
- `vbaloop lint`: 41개 파일, 지적 0건.
- `vbaloop push`: VBE 컴파일 통과 후 UnivRPA.xlsm 저장. 백업 `20261117-131639-930-push`.
- `modTester.Test_FlowNestedInteractions`: 실패 단언 0건.
- `vbaloop test`: 21개 통과, 실패 0개, 건너뜀 0개.
- `git diff --check`: 통과.
- `vbaloop diff`: clean=true, 실질 변경 0개. VBE 대소문자 차이만 5개 모듈에 존재.
- `vbaloop lease release`: 작업 잠금 해제 완료.
- 작업 전 `_dbg/audit_followup_before/src` 백업과 대조하여 이번 변경 범위를 확인했다.

## Test delta

- 기존 소유자 4개 스위트에 회귀 사례를 통합했다. 새 테스트 파일과 공개 진입점은 없으며 수집 개수는 21개로 동일하다.
- 단순 ASSERT 성공 검증을 정규화 및 금지색 검증으로 확장했다. 레지스트리 검증은 메타데이터 검사와 함께 36개 명령의 실행 분기를 호출한다.
- 실제 API 장애는 기존 드라이버에 실패를 주입하여 검증했다. 추가 실클릭, 타이핑, 클립보드 쓰기는 없다.
- 스위트별 Excel 프로세스 추가는 0개다. 레지스트리 검증에는 공유 임시 시트 1개를 사용하고 정리한다.
- 전체 테스트 명령의 관측 시간은 약 6.6초에서 7.1초였다. Excel 시작 비용을 포함한 단일 실행이므로 성능 회귀나 개선의 근거로 삼지 않는다.
- coverage 기준은 변경하지 않았다. 중복 공개 테스트나 별도 통합 테스트 계층은 추가하지 않았다.

## 동작 변경과 검증 한계

CALL 내부의 BREAK/CONTINUE는 해당 호출이 만든 루프만 제어한다. 자체 루프가 없으면 오류로 중단하며 호출자의 루프를 변경하지 않는다. 재귀 IF/WHILE도 호출별 상태를 유지한다.

실제 Windows API 장애 상황과 배포 산출물은 검증하지 않았다. 커밋은 만들지 않았으며 기존 작업 트리 변경은 보존했다.
