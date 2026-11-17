# 감사 8개 항목 수정 계획

작업 전 기준: git 작업 트리 깨끗함, 기존 감사 테스트 21개 통과, 소스와 워크북 일치.

## 계획

- [x] 1. KEY/COMBO의 한 글자 키 변환을 영문자로 제한한다.
- [x] 2. WAIT와 ASSERT가 동일한 색상 검증을 사용한다.
- [x] 3. PIXEL 토큰의 잘못된 좌표를 기본값으로 대체하지 않고 거부한다.
- [x] 4. 실행 엔진에서 스크립트 시그니처를 검사한다.
- [x] 5. 표 데이터의 양끝 공백을 보존하고 헤더만 정리한다.
- [x] 6. 일반 LABEL은 서브루틴 종료 경계에서 제외한다.
- [x] 7. 전체 시트 선택을 CountLarge로 안전하게 처리한다.
- [x] 8. 문서의 줄번호 점프 설명을 정적 LABEL 계약으로 수정한다.
- [x] lint, push 컴파일, 관련 테스트, 전체 테스트, 워크북 diff 확인 후 커밋한다.

## Test delta brief

- Unique risk: 문장부호의 Delete 키 변환, 잘못된 색상/좌표의 성공 처리, 일반 시트 실행, 공백 손실, 내부 LABEL의 정상 서브루틴 거부, 전체 선택 오버플로. 기존 테스트에는 이 입력들이 없다.
- Existing owner: modTester의 TC_AbortLatch, TC_MemoryResolverHeadless, TC_EngineLifecycle, TC_ForeachRowExecution, TC_TableSnapshotReadOnly, TC_ScriptProgramModel, TC_ScriptAssistFormatterRegression.
- Selected layer: 기존 입력 드라이버, 배열/상태 검증, 기존 임시 워크시트 통합 검증. 실제 키 입력이나 픽셀 읽기 없이 오류 경로를 확인한다.
- Expected cost delta: 공개 Test_* 및 파일 추가 없음. 기존 Excel 테스트 실행과 임시 시트를 재사용하며 실제 대기와 외부 프로세스 추가 없음.
- Evidence to retire: 일반 문자열/헤더 사례를 공백 보존 사례로 강화하고 기존 색상 검증 사례를 공유 규칙으로 확장한다. 서로 다른 API 실패 검증은 유지한다.

## 검증 및 Test delta

- `vbaloop init --check`: 관리 파일과 실행 파일 버전 0.6.0 일치.
- `vbaloop lint`: 41개 소스 파일, 지적 0건.
- `vbaloop push`: VBE 컴파일 통과, UnivRPA.xlsm 저장. 마지막 백업: 20261117-140941-107-push.
- 좁은 검증: Test_EngineLifecycle, Test_ScriptAssistFormatterRegression, Test_ForeachRowExecution의 실패 단언 0건.
- 최종 `vbaloop test`: 21개 통과, 실패 0개, 건너뜀 0개. Excel 비정상 종료 없음.
- `vbaloop diff`: clean=true, 변경 모듈 0개. VBE 식별자 대소문자 차이만 5개.
- `git diff --check`: 통과.

첫 전체 실행에서는 필터 기준 열에 공백을 넣은 테스트 데이터 때문에 예상 행수가 달랐다. 필터 기준값은 원래대로 유지하고 메모 열에서 공백 보존을 검증하도록 조정한 뒤 전체 통과했다.

### Test delta

- Unique risk: 잘못된 키/색상/좌표가 실제 동작으로 이어지는 문제, 비스크립트 시트 실행, 데이터 공백 손실, 내부 LABEL 경계 오류, 전체 선택 오버플로.
- Owner and layer: 기존 7개 TC_* 소유자. 입력 드라이버, 배열/상태 검증, 기존 임시 시트를 이용한 실행 및 실제 선택 이벤트 검증.
- Added / extended / removed: 공개 진입점과 테스트 파일 추가 없음. 일반 헤더/문자열 검증을 공백 사례로 강화하고 기존 ASSERT 검증을 공유 색상 규칙으로 갱신했다.
- Collected-case delta: 0 (21개 유지).
- Runtime delta: not measured. 같은 환경에서 단일 전체 명령 관측은 감사 기준 약 11.1초, 최종 약 9.8초였지만 Excel 시작 비용과 변동을 포함하므로 성능 비교 근거로 사용하지 않는다.
- Heavy-resource launch delta: 전체 스위트당 추가 Excel/외부 프로세스 0개. 새 사례의 실제 타이핑, 클릭, 대기는 없다.
- Coverage-floor impact: 기준 변경 없음. 수치 coverage는 측정하지 않았다.

### 호환성과 검증 한계

- 문장부호 키 이름, 잘못된 색상/PIXEL 좌표, 시그니처 없는 시트 실행은 명시 오류가 된다. 숫자 VK 코드와 영문 키 이름은 유지한다.
- 표의 양끝 공백이 보존되므로 기존 자동 공백 제거에 의존하던 입력값은 원본 데이터에서 정리해야 한다.
- 시그니처 판별은 modHelper로 공유하고 기존 modScriptAssist 공개 함수는 유지하여 실행 엔진이 작성 보조에 의존하지 않게 했다.
- 실제 대상 프로그램의 화면 입력과 OS 장애, 배포 산출물, 대용량 성능은 검증하지 않았다.
