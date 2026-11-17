# 감사 5개 항목 수정 계획

기준: 작업 트리 깨끗함. 직전 감사에서 기존 테스트 21개 통과, 워크북과 src 실질 차이 없음.

## 계획

- [x] 1. 스크립트 명령어/파라미터의 Excel 오류값을 로딩 단계에서 거부하고 셀 주소를 표시한다.
- [x] 2. GOTO의 IF 분기 외부 진입과 형제 분기 간 이동을 거부한다. 같은 분기 및 바깥으로의 이동은 유지한다.
- [x] 3. 좁은 통화 열의 표시값 복원에서 Value2로 원래 정밀도를 보존한다.
- [x] 4. FOREACH_ROW의 비연속 범위를 읽기 전에 거부한다.
- [x] 5. 쓰기 변수 주소 내부의 중괄호를 거부한다.
- [x] 기존 테스트에 재현 사례 통합, 사용 계약 문서 갱신, lint/push/관련 테스트/전체 테스트/diff.

## Test delta brief

- Unique risk: 수식 오류가 무한 반복으로 변환됨, GOTO가 IF의 여러 분기를 실행함, 좁은 통화 열에서 정밀도가 손실됨, 비연속 범위의 뒤쪽 행 누락, 읽을 수 없는 변수 주소에 쓰기 성공. 기존 테스트에는 이 입력 조건들이 없다.
- Existing owner: modTester.TC_EngineLifecycle, TC_ScriptProgramModel, TC_FlowNestedInteractions, TC_ForeachRowExecution, TC_ParamGuardRuntime.
- Selected layer: 배열 기반 분기 검증과 기존 임시 시트/입력 드라이버. Excel 고유의 오류값/통화/범위 동작은 기존 시트 통합 검증에서 확인한다.
- Expected cost delta: 공개 Test_*와 테스트 파일 추가 없음. 기존 임시 시트 재사용, 스위트당 추가 Excel/외부 프로세스 0개, 실제 입력과 시간 대기 추가 없음.
- Evidence to retire: 기존 정상 값/분기 검증은 서로 다른 허용 경로를 지키므로 유지한다. 추가 사례는 기존 소유자에 통합하며 별도 중복 스위트는 만들지 않는다.

## 결과

- `vbaloop init --check`: 관리 상태 정상, 버전 0.6.0 일치.
- `vbaloop lint`: 41개 파일, 지적 0건.
- `vbaloop push`: VBE 컴파일 통과, UnivRPA.xlsm 저장. 마지막 백업: `20261117-214636-227-push`.
- `modTester.Test_ForeachRowExecution`: 실패 단언 0건.
- `vbaloop test`: 21개 통과, 실패 0개, 건너뜀 0개. Excel 비정상 종료 없음.
- `vbaloop diff`: clean=true, 실질 변경 모듈 0개. VBE 식별자 대소문자 차이 5개.
- `git diff --check`: 통과.

통화 테스트 준비 과정에서 한국어 Excel의 표시 기호와 숨겨진 Excel의 빈 Text 반환을 확인했다. 최종 테스트는 `$#,##0.000000` 형식을 쓰고, 실제 Value가 Currency 타입인지와 넓은/좁은 열 모두 소수 여섯 자리를 보존하는지 검사한다. 첫 두 시도의 실패는 이 테스트 전제를 보정한 뒤 해소했다.

## Test delta

- Unique risk: 오류값으로 인한 무한 반복, IF 다중 분기 실행, 통화 정밀도 손실, 비연속 범위 누락, 잘못된 변수 주소 쓰기.
- Owner and layer: 기존 5개 TC_* 소유자. 배열 기반 상태/정적 검증, 기존 임시 시트와 입력 드라이버를 사용하는 통합 검증.
- Added / extended / removed: 기존 테스트만 확장. 공개 진입점/테스트 파일/의존성 추가 없음. IF 정적 검증은 실행 전 차단을, 런타임 검증은 사전 검증 없는 ExecuteSmartAction 호출의 방어를 각각 확인한다.
- Collected-case delta: 0 (21개 유지).
- Runtime delta: not measured. 직전 감사 전체 명령 약 7.8초, 최종 약 7.9초는 Excel 시작 비용을 포함한 단일 관측으로 성능 비교 근거가 아니다.
- Heavy-resource launch delta: 전체 스위트당 추가 Excel/외부 프로세스 0개. 새 사례의 실클릭/타이핑/대기 없음.
- Coverage-floor impact: 기준 변경 없음, 수치 coverage는 측정하지 않음.

## 호환성과 검증 한계

- 분기 외부에서 IF 본문으로 들어가던 GOTO, 비연속 표 주소, 중괄호가 포함된 변수명은 이제 명시 오류다.
- 실제 대상 앱 입력, 대용량 성능, 32비트 Excel과 배포 산출물은 검증하지 않았다.
