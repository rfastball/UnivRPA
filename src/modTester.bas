Attribute VB_Name = "modTester"
' @desc 명령어셋 통합 자가진단 및 레지스트리 역검증 실행 담당
' @depends modEngine, modContext, modActions, modWinAPI, modCmdRegistry, modStringDistance
Option Explicit

Private Const DEFAULT_MAX_STEPS As Long = 10000

Public Enum eRunResult
    rrCompleted = 0
    rrHalted = 1
    rrErrored = 2
    rrExhausted = 3
End Enum

'==============================================================================
' [System] Full Integration Test Runner (v2 명령어셋)
'==============================================================================
' 주의: 단일 라인 하니스는 배열을 clsScriptProgram으로 로드한 뒤 ExecuteSmartAction(1)을 호출한다.
'   여기서는 액션/토큰/메모리와 ProgramImage 구조 인덱스를 함께 검증한다.
'   워크시트 UI 동작은 별도 수동 검증 (spec Verification 참조).
'   여기서는 SET/메모리, 토크나이저, TYPE/KEY, CALL/RETURN,
'   비전(WAIT_VANISH/ASSERT), FOCUS, IF-참, 그리고 연산자 단위검증을 다룬다.
'==============================================================================
Public Sub Run_Flow_Diagnosis()
    Dim failCount As Long
    failCount = 0

    Debug.Print "=== [Flow Diagnosis] Headless unit suite ==="
    failCount = failCount + TC_RunProgramSmoke()
    failCount = failCount + TC_FlowHappyPaths()
    failCount = failCount + TC_TableSnapshotReadOnly()
    failCount = failCount + TC_ForeachRowExecution()
    failCount = failCount + TC_FlowNestedInteractions()
    failCount = failCount + TC_FlowNegativePaths()
    failCount = failCount + TC_MemoryResolverHeadless()
    failCount = failCount + TC_ParamSpecValidation()
    failCount = failCount + TC_ParamGuardRuntime()
    failCount = failCount + TC_AbortLatch()
    failCount = failCount + TC_StaticVariableAnalysis()
    failCount = failCount + TC_ScriptAssistFormatterRegression()
    failCount = failCount + TC_StringDistance()
    failCount = failCount + TC_CommandNormalization()

    Debug.Print "--- [Regression] Registry/ScriptProgram ---"
    failCount = failCount + TC_RegistryDispatchSync()
    failCount = failCount + TC_ScriptProgramModel()

    If failCount = 0 Then
        Debug.Print "[OK] Headless flow diagnosis passed."
    Else
        Debug.Print "[FAIL] Headless flow diagnosis failed: " & failCount & " item(s)."
    End If

    On Error Resume Next
    modContext.ResetContext
    On Error GoTo 0
End Sub

Public Sub Run_All()
    Debug.Print "=== [Run_All] Headless flow suite first ==="
    Run_Flow_Diagnosis
    Debug.Print "=== [Run_All] Starting semi-manual system diagnosis (I/O included) ==="
    Run_System_SelfDiagnosis
End Sub

'==============================================================================
' [vbaflow test 진입점] 인자 없는 Public Sub Test_* 만 열거된다.
' 각 TC_* 스위트를 개별 래핑하고, 판정은 LogShim A레코드(ok:false = fail).
' 집계 래퍼는 두지 않는다 (개별 래퍼와 이중 실행되므로).
'==============================================================================

' @desc 스위트 실패 건수를 AssertEq로 기록하고 컨텍스트 정리함
Private Sub ReportSuiteResult(ByVal suiteName As String, ByVal fails As Long)
    LogShim.AssertEq suiteName & ".fail_count", 0, fails
    On Error Resume Next
    modContext.ResetContext
    On Error GoTo 0
End Sub

' @desc RunProgram 스모크 스위트 실행해 실패 0건 단언함
Public Sub Test_RunProgramSmoke()
    ReportSuiteResult "flow.run_program_smoke", TC_RunProgramSmoke()
End Sub

' @desc 흐름 해피패스 스위트 실행해 실패 0건 단언함
Public Sub Test_FlowHappyPaths()
    ReportSuiteResult "flow.happy_paths", TC_FlowHappyPaths()
End Sub

' @desc TableSnapshot 읽기전용 스위트 실행해 실패 0건 단언함
Public Sub Test_TableSnapshotReadOnly()
    ReportSuiteResult "flow.table_snapshot", TC_TableSnapshotReadOnly()
End Sub

' @desc FOREACH_ROW 실행 스위트 실행해 실패 0건 단언함
Public Sub Test_ForeachRowExecution()
    ReportSuiteResult "flow.foreach_row", TC_ForeachRowExecution()
End Sub

' @desc 중첩 흐름 상호작용 스위트 실행해 실패 0건 단언함
Public Sub Test_FlowNestedInteractions()
    ReportSuiteResult "flow.nested_interactions", TC_FlowNestedInteractions()
End Sub

' @desc 흐름 부정경로 스위트 실행해 실패 0건 단언함
Public Sub Test_FlowNegativePaths()
    ReportSuiteResult "flow.negative_paths", TC_FlowNegativePaths()
End Sub

' @desc 메모리/토큰 해석 스위트 실행해 실패 0건 단언함
Public Sub Test_MemoryResolverHeadless()
    ReportSuiteResult "flow.memory_resolver", TC_MemoryResolverHeadless()
End Sub

' @desc ParamSpec 작성보조 검증 스위트 실행해 실패 0건 단언함
Public Sub Test_ParamSpecValidation()
    ReportSuiteResult "flow.param_spec", TC_ParamSpecValidation()
End Sub

' @desc ParamSpec 런타임 게이트 스위트 실행해 실패 0건 단언함
Public Sub Test_ParamGuardRuntime()
    ReportSuiteResult "flow.param_guard_runtime", TC_ParamGuardRuntime()
End Sub

' @desc 중단 래치 스위트 실행해 실패 0건 단언함
Public Sub Test_AbortLatch()
    ReportSuiteResult "flow.abort_latch", TC_AbortLatch()
End Sub

' @desc 정적 변수 분석 스위트 실행해 실패 0건 단언함
Public Sub Test_StaticVariableAnalysis()
    ReportSuiteResult "flow.static_variable_analysis", TC_StaticVariableAnalysis()
End Sub

' @desc 포맷터 주석 회귀 스위트 실행해 실패 0건 단언함
Public Sub Test_ScriptAssistFormatterRegression()
    ReportSuiteResult "flow.formatter_regression", TC_ScriptAssistFormatterRegression()
End Sub

' @desc 문자열 거리 유틸 스위트 실행해 실패 0건 단언함
Public Sub Test_StringDistance()
    ReportSuiteResult "util.string_distance", TC_StringDistance()
End Sub

' @desc 명령어 정규화 규칙 스위트 실행해 실패 0건 단언함
Public Sub Test_CommandNormalization()
    ReportSuiteResult "registry.command_normalization", TC_CommandNormalization()
End Sub

' @desc 레지스트리-핸들러 동기화 역검증 실행해 실패 0건 단언함
Public Sub Test_RegistryDispatchSync()
    ReportSuiteResult "registry.dispatch_sync", TC_RegistryDispatchSync()
End Sub

' @desc ScriptProgram 정적 모델 스위트 실행해 실패 0건 단언함
Public Sub Test_ScriptProgramModel()
    ReportSuiteResult "model.script_program", TC_ScriptProgramModel()
End Sub

Public Function RunProgram(ByRef arr As Variant, ByVal firstRow As Long, ByVal maxSteps As Long, _
                           ByRef outTrace() As Long, ByRef outLastCode As Long, _
                           ByRef outLastLine As Long) As eRunResult
    RunProgram = RunProgramFromLine(arr, firstRow, 1, maxSteps, outTrace, outLastCode, outLastLine)
End Function

Private Function RunProgramFromLine(ByRef arr As Variant, ByVal firstRow As Long, ByVal startLine As Long, _
                                    ByVal maxSteps As Long, ByRef outTrace() As Long, _
                                    ByRef outLastCode As Long, ByRef outLastLine As Long) As eRunResult
    On Error GoTo ErrHandler

    Dim curLine As Long
    Dim totalLines As Long
    Dim steps As Long
    Dim code As Long
    Dim program As clsScriptProgram

    Erase outTrace
    outLastCode = 0
    outLastLine = 0
    If maxSteps <= 0 Then maxSteps = DEFAULT_MAX_STEPS

    modActions.SetUnattendedMode True   ' 테스트 하니스 = 무인 모드 (팝업 억제)
    modActions.ClearAbort               ' 이전 실행의 중단 래치 해제
    modContext.InitContext
    modEngine.LoadProgramFromArray arr, firstRow
    Set program = modEngine.ActiveProgram()
    totalLines = program.totalLines

    curLine = startLine
    Do While curLine >= 1 And curLine <= totalLines
        If steps >= maxSteps Then
            outLastLine = curLine
            RunProgramFromLine = rrExhausted
            Exit Function
        End If

        AppendTrace outTrace, steps + 1, curLine
        code = modEngine.ExecuteSmartAction(curLine)
        outLastCode = code
        steps = steps + 1

        ' [B2] 중단 래치 우선: 엔진 루프와 동일하게 중단으로 수렴
        If modActions.AbortPending() Then
            outLastLine = curLine
            RunProgramFromLine = rrHalted
            Exit Function
        End If

        Select Case code
            Case CMD_NEXT_LINE
                curLine = curLine + 1
            Case CMD_HALT
                outLastLine = curLine
                RunProgramFromLine = rrHalted
                Exit Function
            Case CMD_ERROR
                outLastLine = curLine
                RunProgramFromLine = rrErrored
                Exit Function
            Case Is > 0
                curLine = code
            Case Else
                outLastLine = curLine
                RunProgramFromLine = rrErrored
                Exit Function
        End Select
    Loop

    outLastLine = curLine
    RunProgramFromLine = rrCompleted
    Exit Function

ErrHandler:
    outLastCode = CMD_ERROR
    outLastLine = curLine
    RunProgramFromLine = rrErrored
End Function

Public Function Step1(ByRef arr As Variant, ByVal R As Long) As Long
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long

    Call RunProgramFromLine(arr, 1, R, 1, trace, lastCode, lastLine)
    Step1 = lastCode
End Function

Sub Run_System_SelfDiagnosis()
    Dim wsTest As Worksheet
    Dim failCount As Long
    Dim arrCmd(1 To 60, 1 To 5) As Variant
    Dim nextLine As Long

    ' 1. 테스트 환경 준비
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("RPA_TEST_SUITE").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsTest = Sheets.Add
    wsTest.name = "RPA_TEST_SUITE"

    ' 컨텍스트 및 변수 초기화
    modContext.InitContext
    failCount = 0

    ' 엑셀 창 활성화 (입력 테스트 필수)
    AppActivate Application.Caption
    modActions.BetterWait 500

    '==========================================================================
    ' [Test 1] 메모리 SET + 비우기 ({$BLANK})
    '==========================================================================
    Debug.Print "--- [Test 1] Memory (SET / $BLANK) ---"

    Call MakeCmd(arrCmd, 1, "SET", "{score}", 0, "10")
    ExecuteArrayLine arrCmd, 1

    If modContext.GetVar("score") = "10" Then
        LogPass "SET (mem write)"
    Else
        LogFail "SET", "10", modContext.GetVar("score"): failCount = failCount + 1
    End If

    Call MakeCmd(arrCmd, 1, "SET", "{score}", 0, "{$BLANK}")
    ExecuteArrayLine arrCmd, 1

    If modContext.GetVar("score") = "" Then
        LogPass "SET {$BLANK} (clear)"
    Else
        LogFail "SET {$BLANK}", "(empty)", modContext.GetVar("score"): failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 2] 토크나이저: 문장 내 {name} 보간 + {$TAB} 리터럴
    '==========================================================================
    Debug.Print "--- [Test 2] Tokenizer (interp + literal) ---"

    Call MakeCmd(arrCmd, 1, "SET", "{a}", 0, "hello")
    ExecuteArrayLine arrCmd, 1
    Call MakeCmd(arrCmd, 1, "SET", "{b}", 0, "{a}_world")
    ExecuteArrayLine arrCmd, 1

    If modContext.GetVar("b") = "hello_world" Then
        LogPass "Interpolation ({a}_world)"
    Else
        LogFail "Interpolation", "hello_world", modContext.GetVar("b"): failCount = failCount + 1
    End If

    Call MakeCmd(arrCmd, 1, "SET", "{t}", 0, "A{$TAB}B")
    ExecuteArrayLine arrCmd, 1

    If modContext.GetVar("t") = "A" & vbTab & "B" Then
        LogPass "Literal token ({$TAB})"
    Else
        LogFail "{$TAB}", "A<TAB>B", modContext.GetVar("t"): failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 3] {$LOOP} 루프 밖에서는 0
    '==========================================================================
    Debug.Print "--- [Test 3] {$LOOP} outside loop = 0 ---"

    Call MakeCmd(arrCmd, 1, "SET", "{L}", 0, "{$LOOP}")
    ExecuteArrayLine arrCmd, 1

    If modContext.GetVar("L") = "0" Then
        LogPass "{$LOOP} (no loop -> 0)"
    Else
        LogFail "{$LOOP}", "0", modContext.GetVar("L"): failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 4] 시스템 토큰: {$CLIP} (읽기/쓰기) + {$PIXEL:x,y}
    '==========================================================================
    Debug.Print "--- [Test 4] System tokens ($CLIP / $PIXEL) ---"

    ' 쓰기: SET {$CLIP} 값  -> 읽기: SET {c} {$CLIP}
    Call MakeCmd(arrCmd, 1, "SET", "{$CLIP}", 0, "CLIP_RT")
    ExecuteArrayLine arrCmd, 1
    modActions.BetterWait 80
    Call MakeCmd(arrCmd, 1, "SET", "{c}", 0, "{$CLIP}")
    ExecuteArrayLine arrCmd, 1

    If modContext.GetVar("c") = "CLIP_RT" Then
        LogPass "SET {$CLIP} (write/read)"
    Else
        LogFail "{$CLIP}", "CLIP_RT", modContext.GetVar("c"): failCount = failCount + 1
    End If

    Dim screenX As Long: screenX = 300
    Dim screenY As Long: screenY = 300
    Dim actualColor As String
    actualColor = modWinAPI.GetColorHexAt(screenX, screenY)

    Call MakeCmd(arrCmd, 1, "SET", "{px}", 0, "{$PIXEL:" & screenX & "," & screenY & "}")
    ExecuteArrayLine arrCmd, 1

    If modContext.GetVar("px") = actualColor Then
        LogPass "{$PIXEL:x,y} (color read)"
    Else
        LogFail "{$PIXEL}", actualColor, modContext.GetVar("px"): failCount = failCount + 1
    End If

    If Left$(modContext.GetVar("px"), 1) <> "#" Then
        LogPass "{$PIXEL:x,y} (RRGGBB without #)"
    Else
        LogFail "{$PIXEL}", "RRGGBB without #", modContext.GetVar("px"): failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 5] 타이핑 (TYPE 텍스트 + KEY Enter)
    '==========================================================================
    Debug.Print "--- [Test 5] Input (TYPE & Enter) ---"

    wsTest.Range("A1").ClearContents
    wsTest.Range("A1").Select

    Call MakeCmd(arrCmd, 1, "TYPE", 10, 0, "HelloRPA") ' P1=딜레이10ms, P3=텍스트
    ExecuteArrayLine arrCmd, 1
    Call MakeCmd(arrCmd, 1, "KEY", 13, 0, "")
    ExecuteArrayLine arrCmd, 1

    DoEvents
    If wsTest.Range("A1").value = "HelloRPA" Then
        LogPass "TYPE + KEY (Cell Input)"
    Else
        LogFail "TYPE", "HelloRPA", wsTest.Range("A1").value: failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 6] 메모리 변수 타이핑 (TYPE {name})
    '==========================================================================
    Debug.Print "--- [Test 6] TYPE {name} ---"

    Call MakeCmd(arrCmd, 1, "SET", "{dyn}", 0, "DynamicValue")
    ExecuteArrayLine arrCmd, 1

    wsTest.Range("A2").Select
    Call MakeCmd(arrCmd, 1, "TYPE", 10, 0, "{dyn}")
    ExecuteArrayLine arrCmd, 1
    Call MakeCmd(arrCmd, 1, "KEY", 13, 0, "")
    ExecuteArrayLine arrCmd, 1

    DoEvents
    If wsTest.Range("A2").value = "DynamicValue" Then
        LogPass "TYPE {name} (Var -> Cell)"
    Else
        LogFail "TYPE {name}", "DynamicValue", wsTest.Range("A2").value: failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 7] 흐름 제어 (CALL / RETURN)
    '==========================================================================
    Debug.Print "--- [Test 7] Flow (CALL/RETURN) ---"

    Call MakeCmd(arrCmd, 1, "CALL", 0, 0, "SUB_TEST")
    Call MakeCmd(arrCmd, 2, "EXIT", 0, 0, "")
    Call MakeCmd(arrCmd, 3, "LABEL", 0, 0, "SUB_TEST")
    Call MakeCmd(arrCmd, 4, "RETURN", 0, 0, "")
    nextLine = ExecuteArrayLine(arrCmd, 1)

    If nextLine = 3 And modContext.GetStackDepth() = 1 Then
        LogPass "CALL (Stack Push & Jump)"

        nextLine = ExecuteArrayLine(arrCmd, 4)

        If nextLine = 2 And modContext.GetStackDepth() = 0 Then
            LogPass "RETURN (Stack Pop & Return)"
        Else
            LogFail "RETURN", "Return 2", "Line=" & nextLine: failCount = failCount + 1
        End If
    Else
        LogFail "CALL", "Jump SUB_TEST", "Line=" & nextLine: failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 8] IF (참 갈래는 최소 블록으로 검증: 다음 줄 반환)
    '   IF {$PIXEL:x,y} EQ <현재색> / END_IF  -> 참 -> 0 반환
    '==========================================================================
    Debug.Print "--- [Test 8] IF (true branch) ---"

    Call MakeCmd(arrCmd, 1, "IF", "{$PIXEL:" & screenX & "," & screenY & "}", "EQ", actualColor)
    Call MakeCmd(arrCmd, 2, "END_IF", 0, 0, "")
    nextLine = ExecuteArrayLine(arrCmd, 1)

    If nextLine = 0 Then
        LogPass "IF EQ (true -> next line)"
        nextLine = ExecuteArrayLine(arrCmd, 2)
        If nextLine <> 0 Then
            LogFail "END_IF", "0 (next)", "Line=" & nextLine: failCount = failCount + 1
        End If
    Else
        LogFail "IF EQ", "0 (next)", "Line=" & nextLine: failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 9] 소멸 대기 (WAIT_VANISH 로직 검증 - 즉시 통과)
    '==========================================================================
    Debug.Print "--- [Test 9] WAIT_VANISH (logic) ---"

    wsTest.Range("E1").Interior.color = RGB(255, 255, 255) ' 흰색

    Dim targetRange As Range
    Dim px As Long, py As Long
    Set targetRange = wsTest.Range("E1")
    px = ActiveWindow.PointsToScreenPixelsX(targetRange.Left) + 10
    py = ActiveWindow.PointsToScreenPixelsY(targetRange.Top) + 5

    ' (px,py)가 검정(000000)이 아니면 즉시 통과해야 함 (현재 흰색)
    Call MakeCmd(arrCmd, 1, "WAIT_VANISH", px, py, "000000")

    Dim tStart As Long, tEnd As Long
    tStart = modWinAPI.GetTickCount()
    nextLine = ExecuteArrayLine(arrCmd, 1)
    tEnd = modWinAPI.GetTickCount()

    If (tEnd - tStart) < 1000 And nextLine <> -1 Then
        LogPass "WAIT_VANISH (Logic Check)"
    Else
        LogFail "WAIT_VANISH", "Immediate Pass", "Took " & (tEnd - tStart) & "ms": failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 10] 창 활성화 (FOCUS)
    '==========================================================================
    Debug.Print "--- [Test 10] FOCUS ---"

    Call MakeCmd(arrCmd, 1, "FOCUS", 0, 0, "Excel")
    nextLine = ExecuteArrayLine(arrCmd, 1)

    If nextLine <> -1 Then
        LogPass "FOCUS (Find Window)"
    Else
        LogFail "FOCUS", "Success", "Failed (-1)": failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 11] 조건 연산자 단위 검증 (modCondition 직접 호출)
    '   EQ 정확비교 / LIKE 패턴 / STARTS / ENDS / CONTAINS / 맨 '=' 금지
    '   Excel 셀에는 ==를 그대로 입력할 수 없고, 사용하려면 '==로 입력해야 한다.
    '==========================================================================
    Debug.Print "--- [Test 11] Condition operators ---"

    AssertCond "num >", "5", ">", "3", True, failCount
    AssertCond "EQ exact", "abc", "EQ", "abc", True, failCount
    AssertCond "EQ star literal", "abc", "EQ", "ab*", False, failCount   ' * 는 리터럴
    AssertCond "== compatibility", "abc", "==", "abc", True, failCount
    AssertCond "LIKE pattern", "abc", "LIKE", "a*", True, failCount
    AssertCond "STARTS", "abcdef", "STARTS", "abc", True, failCount
    AssertCond "ENDS", "abcdef", "ENDS", "def", True, failCount
    AssertCond "CONTAINS", "abcdef", "CONTAINS", "cd", True, failCount
    AssertCond "bare = banned", "5", "=", "5", False, failCount          ' 맨 '=' 금지 -> False
    AssertCond "unknown operator", "5", "EQUALS", "5", False, failCount

    If modDevTools.ValidateLine("IF", "abc", "EQ", "abc") Then
        LogPass "ValidateLine EQ"
    Else
        LogFail "ValidateLine EQ", "True", "False": failCount = failCount + 1
    End If

    If modDevTools.ValidateLine("IF", "abc", "==", "abc") Then
        LogPass "ValidateLine == compatibility"
    Else
        LogFail "ValidateLine ==", "True", "False": failCount = failCount + 1
    End If

    If Not modDevTools.ValidateLine("IF", "abc", "EQUALS", "abc") Then
        LogPass "ValidateLine unknown operator"
    Else
        LogFail "ValidateLine unknown operator", "False", "True": failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 12] INC / DEC (정수 read-modify-write, {$LOOP} 비대칭 닫힘)
    '==========================================================================
    Debug.Print "--- [Test 12] INC / DEC ---"

    ' 미정의 변수 INC -> 오류
    Dim incMissing As Long
    Call MakeCmd(arrCmd, 1, "INC", "{n}", 0, "")
    incMissing = ExecuteArrayLine(arrCmd, 1)
    If incMissing = -1 Then
        LogPass "INC (undefined -> ERROR)"
    Else
        LogFail "INC undef", "-1", CStr(incMissing): failCount = failCount + 1
    End If

    Call MakeCmd(arrCmd, 1, "SET", "{n}", 0, "0")
    ExecuteArrayLine arrCmd, 1
    Call MakeCmd(arrCmd, 1, "INC", "{n}", 0, "")
    ExecuteArrayLine arrCmd, 1
    If modContext.GetVar("n") = "1" Then
        LogPass "INC (defined 0 -> 1)"
    Else
        LogFail "INC defined", "1", modContext.GetVar("n"): failCount = failCount + 1
    End If

    ' INC {n} 5 -> 6 (P3 RESOLVE)
    Call MakeCmd(arrCmd, 1, "INC", "{n}", 0, "5")
    ExecuteArrayLine arrCmd, 1
    If modContext.GetVar("n") = "6" Then
        LogPass "INC (+5 -> 6)"
    Else
        LogFail "INC +5", "6", modContext.GetVar("n"): failCount = failCount + 1
    End If

    ' DEC {n} -> 5 (기본 1 감분)
    Call MakeCmd(arrCmd, 1, "DEC", "{n}", 0, "")
    ExecuteArrayLine arrCmd, 1
    If modContext.GetVar("n") = "5" Then
        LogPass "DEC (-1 -> 5)"
    Else
        LogFail "DEC", "5", modContext.GetVar("n"): failCount = failCount + 1
    End If

    ' 읽기전용 좌변 INC {$LOOP} -> CMD_ERROR(-1)
    Dim incErr As Long
    Call MakeCmd(arrCmd, 1, "INC", "{$LOOP}", 0, "")
    incErr = ExecuteArrayLine(arrCmd, 1)
    If incErr = -1 Then
        LogPass "INC {$LOOP} (read-only -> ERROR)"
    Else
        LogFail "INC {$LOOP}", "-1", CStr(incErr): failCount = failCount + 1
    End If

    ' [정리] Test 13(레지스트리 동기화)/Test 14(ScriptProgram 모델)는 헤드리스
    ' 스위트(Run_Flow_Diagnosis)와 중복이라 제거됨. 여기는 I/O 검증 전용.

    '==========================================================================
    ' 종료 처리
    '==========================================================================
    Application.ScreenUpdating = True

    Dim resMsg As String
    If failCount = 0 Then
        resMsg = "[OK] 모든 시스템 검증 통과 (All Systems Go)"
    Else
        resMsg = "[FAIL] " & failCount & "개 항목 실패. Immediate 창을 확인하세요."
    End If

    MsgBox resMsg, IIf(failCount = 0, vbInformation, vbCritical), "Self Diagnosis (v2)"
End Sub

' @desc 등록 명령어 전수 순회로 레지스트리<->핸들러 라우팅 동기화 검증, 실패 개수 반환
'   '테이블엔 있는데 핸들러 없음'(카테고리 미배선/오타)을 구조적으로 막는 파수꾼 (설계 §7).
'   부작용 없이 라우팅 무결성만 확인 ? 실제 실행(클릭/대기)은 자가진단/시트 검증이 담당.
Public Function TC_RegistryDispatchSync() As Long
    Dim cmds As Variant: cmds = modCmdRegistry.AllCommands()
    Dim i As Long, fails As Long
    Dim cmd As String, def As CmdDef, h As ICommand

    For i = LBound(cmds) To UBound(cmds)
        cmd = CStr(cmds(i))
        Set def = modCmdRegistry.GetCmdDef(cmd)
        If def Is Nothing Then
            Debug.Print "[FAIL] 게이트 미해석: " & cmd
            fails = fails + 1
        Else
            Set h = modCmdRegistry.handler(def.Category)
            If h Is Nothing Then
                Debug.Print "[FAIL] 핸들러 없음 (category=" & def.Category & "): " & cmd
                fails = fails + 1
            End If
        End If
    Next i

    Debug.Print "[역검증] 총 " & (UBound(cmds) - LBound(cmds) + 1) & "개 명령어 라우팅 확인, 실패 " & fails
    TC_RegistryDispatchSync = fails
End Function

' [Helper] 연산자 단위 검증
Private Sub AssertCond(ByVal name As String, ByVal v1 As String, ByVal op As String, _
                       ByVal v2 As String, ByVal expected As Boolean, ByRef failCount As Long)
    Dim actual As Boolean
    actual = modCondition.EvaluateCondition(v1, op, v2)
    If actual = expected Then
        LogPass "Cond " & name
    Else
        LogFail "Cond " & name, CStr(expected), CStr(actual)
        failCount = failCount + 1
    End If
End Sub

Private Sub AssertElseIfOperator(ByVal name As String, ByVal v1 As String, ByVal op As String, _
                                 ByVal v2 As String, ByRef failCount As Long)
    Dim arr(1 To 6, 1 To 5) As Variant
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult

    Call MakeCmd(arr, 1, "IF", "never", "EQ", "match")
    Call MakeCmd(arr, 2, "SET", "{elseif_op_hit}", 0, "bad")
    Call MakeCmd(arr, 3, "ELSE_IF", v1, op, v2)
    Call MakeCmd(arr, 4, "SET", "{elseif_op_hit}", 0, name)
    Call MakeCmd(arr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 6, "EXIT", 0, 0, "")

    result = RunProgram(arr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, failCount
    ExpectVisits trace, 4, 1, failCount
    ExpectVar "elseif_op_hit", name, failCount
    ExpectIfBranchDepth 0, failCount
End Sub

Private Function TC_RunProgramSmoke() As Long
    Dim fails As Long
    Dim arr(1 To 3, 1 To 5) As Variant
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult

    Debug.Print "--- [Flow 1] RunProgram smoke ---"
    Call MakeCmd(arr, 1, "SET", "{a}", 0, "alpha")
    Call MakeCmd(arr, 2, "SET", "{b}", 0, "{a}_beta")
    Call MakeCmd(arr, 3, "EXIT", 0, 0, "")

    result = RunProgram(arr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "a", "alpha", fails
    ExpectVar "b", "alpha_beta", fails
    ExpectTrace trace, Array(1, 2, 3), fails
    ExpectLong "Smoke last line", 3, lastLine, fails
    ExpectLong "Smoke last code", CMD_HALT, lastCode, fails

    TC_RunProgramSmoke = fails
End Function

Private Function TC_FlowHappyPaths() As Long
    Dim fails As Long
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult

    Debug.Print "--- [Flow 2] Happy paths ---"

    Dim gotoArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(gotoArr, 1, "LABEL", 0, 0, "START")
    Call MakeCmd(gotoArr, 2, "SET", "{mark}", 0, "before")
    Call MakeCmd(gotoArr, 3, "GOTO", 0, 0, "DONE")
    Call MakeCmd(gotoArr, 4, "SET", "{mark}", 0, "bad")
    Call MakeCmd(gotoArr, 5, "LABEL", 0, 0, "DONE")
    Call MakeCmd(gotoArr, 6, "EXIT", 0, 0, "")
    result = RunProgram(gotoArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "mark", "before", fails
    ExpectVisits trace, 4, 0, fails
    ExpectTrace trace, Array(1, 2, 3, 5, 6), fails

    Dim gotoLoopExitArr(1 To 8, 1 To 5) As Variant
    Call MakeCmd(gotoLoopExitArr, 1, "REPEAT", "5", 0, "")
    Call MakeCmd(gotoLoopExitArr, 2, "SET", "{loop_before_goto}", 0, "0")
    Call MakeCmd(gotoLoopExitArr, 3, "INC", "{loop_before_goto}", 0, "")
    Call MakeCmd(gotoLoopExitArr, 4, "GOTO", 0, 0, "OUT")
    Call MakeCmd(gotoLoopExitArr, 5, "INC", "{loop_after_goto}", 0, "")
    Call MakeCmd(gotoLoopExitArr, 6, "END_REPEAT", 0, 0, "")
    Call MakeCmd(gotoLoopExitArr, 7, "LABEL", 0, 0, "OUT")
    Call MakeCmd(gotoLoopExitArr, 8, "EXIT", 0, 0, "")
    result = RunProgram(gotoLoopExitArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "loop_before_goto", "1", fails
    ExpectVar "loop_after_goto", "", fails
    ExpectLoopDepth 0, fails

    Dim callArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(callArr, 1, "CALL", 0, 0, "SUB1")
    Call MakeCmd(callArr, 2, "SET", "{after}", 0, "yes")
    Call MakeCmd(callArr, 3, "EXIT", 0, 0, "")
    Call MakeCmd(callArr, 4, "LABEL", 0, 0, "SUB1")
    Call MakeCmd(callArr, 5, "SET", "{sub}", 0, "ran")
    Call MakeCmd(callArr, 6, "RETURN", 0, 0, "")
    result = RunProgram(callArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "sub", "ran", fails
    ExpectVar "after", "yes", fails
    ExpectStackDepth 0, fails
    ExpectTrace trace, Array(1, 4, 5, 6, 2, 3), fails

    Dim ifTrueArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(ifTrueArr, 1, "IF", "1", "EQ", "1")
    Call MakeCmd(ifTrueArr, 2, "SET", "{if_true}", 0, "yes")
    Call MakeCmd(ifTrueArr, 3, "ELSE", 0, 0, "")
    Call MakeCmd(ifTrueArr, 4, "SET", "{if_else}", 0, "bad")
    Call MakeCmd(ifTrueArr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(ifTrueArr, 6, "EXIT", 0, 0, "")
    result = RunProgram(ifTrueArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 2, 1, fails
    ExpectVisits trace, 4, 0, fails
    ExpectVar "if_true", "yes", fails

    Dim ifFalseArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(ifFalseArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(ifFalseArr, 2, "SET", "{if_body}", 0, "bad")
    Call MakeCmd(ifFalseArr, 3, "ELSE", 0, 0, "")
    Call MakeCmd(ifFalseArr, 4, "SET", "{else_body}", 0, "yes")
    Call MakeCmd(ifFalseArr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(ifFalseArr, 6, "EXIT", 0, 0, "")
    result = RunProgram(ifFalseArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 2, 0, fails
    ExpectVisits trace, 4, 1, fails
    ExpectVar "else_body", "yes", fails

    Dim elseIfNoElseArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(elseIfNoElseArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(elseIfNoElseArr, 2, "SET", "{elseif_if_body}", 0, "bad")
    Call MakeCmd(elseIfNoElseArr, 3, "ELSE_IF", "2", "EQ", "0")
    Call MakeCmd(elseIfNoElseArr, 4, "SET", "{elseif_body}", 0, "bad")
    Call MakeCmd(elseIfNoElseArr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(elseIfNoElseArr, 6, "EXIT", 0, 0, "")
    result = RunProgram(elseIfNoElseArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 2, 0, fails
    ExpectVisits trace, 4, 0, fails
    ExpectIfBranchDepth 0, fails

    Dim elseIfElseArr(1 To 8, 1 To 5) As Variant
    Call MakeCmd(elseIfElseArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(elseIfElseArr, 2, "SET", "{elseif_first}", 0, "bad")
    Call MakeCmd(elseIfElseArr, 3, "ELSE_IF", "2", "EQ", "0")
    Call MakeCmd(elseIfElseArr, 4, "SET", "{elseif_second}", 0, "bad")
    Call MakeCmd(elseIfElseArr, 5, "ELSE", 0, 0, "")
    Call MakeCmd(elseIfElseArr, 6, "SET", "{elseif_else}", 0, "yes")
    Call MakeCmd(elseIfElseArr, 7, "END_IF", 0, 0, "")
    Call MakeCmd(elseIfElseArr, 8, "EXIT", 0, 0, "")
    result = RunProgram(elseIfElseArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 2, 0, fails
    ExpectVisits trace, 4, 0, fails
    ExpectVisits trace, 6, 1, fails
    ExpectVar "elseif_else", "yes", fails
    ExpectIfBranchDepth 0, fails

    Dim elseIfChainArr(1 To 10, 1 To 5) As Variant
    Call MakeCmd(elseIfChainArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(elseIfChainArr, 2, "SET", "{elseif_chain_if}", 0, "bad")
    Call MakeCmd(elseIfChainArr, 3, "ELSE_IF", "2", "EQ", "2")
    Call MakeCmd(elseIfChainArr, 4, "SET", "{elseif_chain_hit}", 0, "yes")
    Call MakeCmd(elseIfChainArr, 5, "ELSE_IF", "3", "EQ", "3")
    Call MakeCmd(elseIfChainArr, 6, "SET", "{elseif_chain_late}", 0, "bad")
    Call MakeCmd(elseIfChainArr, 7, "ELSE", 0, 0, "")
    Call MakeCmd(elseIfChainArr, 8, "SET", "{elseif_chain_else}", 0, "bad")
    Call MakeCmd(elseIfChainArr, 9, "END_IF", 0, 0, "")
    Call MakeCmd(elseIfChainArr, 10, "EXIT", 0, 0, "")
    result = RunProgram(elseIfChainArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 4, 1, fails
    ExpectVisits trace, 6, 0, fails
    ExpectVisits trace, 8, 0, fails
    ExpectVar "elseif_chain_hit", "yes", fails
    ExpectIfBranchDepth 0, fails

    AssertElseIfOperator "LIKE", "abc", "LIKE", "a*", fails
    AssertElseIfOperator "STARTS", "abcdef", "STARTS", "abc", fails
    AssertElseIfOperator "ENDS", "abcdef", "ENDS", "def", fails
    AssertElseIfOperator "CONTAINS", "abcdef", "CONTAINS", "cd", fails

    Dim whileArr(1 To 13, 1 To 5) As Variant
    Call MakeCmd(whileArr, 1, "SET", "{i}", 0, "0")
    Call MakeCmd(whileArr, 2, "WHILE", "{i}", "<", "3")
    Call MakeCmd(whileArr, 3, "INC", "{i}", 0, "")
    Call MakeCmd(whileArr, 4, "SELECT", "{$LOOP}", 0, "")
    Call MakeCmd(whileArr, 5, "CASE", "1", 0, "")
    Call MakeCmd(whileArr, 6, "SET", "{L1}", 0, "{$LOOP}")
    Call MakeCmd(whileArr, 7, "CASE", "2", 0, "")
    Call MakeCmd(whileArr, 8, "SET", "{L2}", 0, "{$LOOP}")
    Call MakeCmd(whileArr, 9, "CASE", "3", 0, "")
    Call MakeCmd(whileArr, 10, "SET", "{L3}", 0, "{$LOOP}")
    Call MakeCmd(whileArr, 11, "END_SELECT", 0, 0, "")
    Call MakeCmd(whileArr, 12, "END_WHILE", 0, 0, "")
    Call MakeCmd(whileArr, 13, "EXIT", 0, 0, "")
    result = RunProgram(whileArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 3, 3, fails
    ExpectVar "i", "3", fails
    ExpectVar "L1", "1", fails
    ExpectVar "L2", "2", fails
    ExpectVar "L3", "3", fails
    ExpectLoopDepth 0, fails

    Dim zeroWhileArr(1 To 4, 1 To 5) As Variant
    Call MakeCmd(zeroWhileArr, 1, "WHILE", "1", "EQ", "0")
    Call MakeCmd(zeroWhileArr, 2, "SET", "{body}", 0, "bad")
    Call MakeCmd(zeroWhileArr, 3, "END_WHILE", 0, 0, "")
    Call MakeCmd(zeroWhileArr, 4, "EXIT", 0, 0, "")
    result = RunProgram(zeroWhileArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 2, 0, fails
    ExpectLoopDepth 0, fails

    Dim repeatArr(1 To 5, 1 To 5) As Variant
    Call MakeCmd(repeatArr, 1, "SET", "{r}", 0, "0")
    Call MakeCmd(repeatArr, 2, "REPEAT", "3", 0, "")
    Call MakeCmd(repeatArr, 3, "INC", "{r}", 0, "")
    Call MakeCmd(repeatArr, 4, "END_REPEAT", 0, 0, "")
    Call MakeCmd(repeatArr, 5, "EXIT", 0, 0, "")
    result = RunProgram(repeatArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 3, 3, fails
    ExpectVar "r", "3", fails
    ExpectLoopDepth 0, fails

    Dim selectArr(1 To 9, 1 To 5) As Variant
    Call MakeCmd(selectArr, 1, "SELECT", "B", 0, "")
    Call MakeCmd(selectArr, 2, "CASE", "A", 0, "")
    Call MakeCmd(selectArr, 3, "SET", "{caseA}", 0, "bad")
    Call MakeCmd(selectArr, 4, "CASE", "B", 0, "")
    Call MakeCmd(selectArr, 5, "SET", "{caseB}", 0, "yes")
    Call MakeCmd(selectArr, 6, "CASE", "ELSE", 0, "")
    Call MakeCmd(selectArr, 7, "SET", "{fallback}", 0, "yes")
    Call MakeCmd(selectArr, 8, "END_SELECT", 0, 0, "")
    Call MakeCmd(selectArr, 9, "EXIT", 0, 0, "")
    result = RunProgram(selectArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 3, 0, fails
    ExpectVisits trace, 5, 1, fails
    ExpectVisits trace, 7, 0, fails
    ExpectVar "caseB", "yes", fails

    selectArr(1, 3) = "Z"
    result = RunProgram(selectArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 7, 1, fails
    ExpectVar "fallback", "yes", fails

    TC_FlowHappyPaths = fails
End Function

Private Function TC_TableSnapshotReadOnly() As Long
    On Error GoTo ErrHandler

    Dim fails As Long
    Dim data(1 To 4, 1 To 3) As Variant
    Dim tbl As clsTableSnapshot
    Dim rec As Variant
    Dim rows As Collection
    Dim keys As Variant

    Debug.Print "--- [Flow 2A] TableSnapshot read-only surface ---"

    data(1, 1) = "Code": data(1, 2) = "Name": data(1, 3) = "Group"
    data(2, 1) = "001": data(2, 2) = "Alice": data(2, 3) = "A"
    data(3, 1) = "002": data(3, 2) = "Bob": data(3, 3) = "B"
    data(4, 1) = "003": data(4, 2) = "Cara": data(4, 3) = "A"

    Set tbl = New clsTableSnapshot
    tbl.LoadFromArray data

    ExpectLong "Snapshot RowCount", 4, tbl.RowCount, fails
    ExpectLong "Snapshot FirstDataRow", 2, tbl.FirstDataRow, fails
    ExpectLong "Snapshot ColCount", 3, tbl.ColCount, fails
    ExpectBool "Snapshot Headers CODE", True, tbl.headers.exists("CODE"), fails
    ExpectText "Snapshot GetValue", "Alice", CStr(tbl.GetValue(2, "Name")), fails

    rec = tbl.GetRecord(3)
    ExpectText "Snapshot GetRecord", "Bob", CStr(rec(2)), fails

    Set rows = tbl.SelectRows("Group", "A")
    ExpectLong "Snapshot SelectRows count", 2, rows.count, fails
    ExpectLong "Snapshot SelectRows first", 2, CLng(rows(1)), fails
    ExpectLong "Snapshot SelectRows second", 4, CLng(rows(2)), fails

    Set rows = tbl.SelectRowsByIndex(1, "002")
    ExpectLong "Snapshot SelectRowsByIndex count", 1, rows.count, fails
    ExpectLong "Snapshot SelectRowsByIndex row", 3, CLng(rows(1)), fails

    keys = tbl.GetUniqueKeys("Group")
    ExpectLong "Snapshot GetUniqueKeys count", 2, UBound(keys) - LBound(keys) + 1, fails

    keys = tbl.GetUniqueKeysByIndex(1)
    ExpectLong "Snapshot GetUniqueKeysByIndex count", 3, UBound(keys) - LBound(keys) + 1, fails

    tbl.Clear
    ExpectLong "Snapshot Clear RowCount", 0, tbl.RowCount, fails
    ExpectLong "Snapshot Clear ColCount", 0, tbl.ColCount, fails

    TC_TableSnapshotReadOnly = fails
    Exit Function

ErrHandler:
    fails = fails + 1
    LogFail "TableSnapshot read-only test", "no error", Err.Description
    TC_TableSnapshotReadOnly = fails
End Function

Private Function TC_ForeachRowExecution() As Long
    On Error GoTo ErrHandler

    Dim fails As Long
    Dim ws As Worksheet
    Dim sheetName As String
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult

    Debug.Print "--- [Flow 2B] FOREACH_ROW table iteration ---"

    sheetName = "FOREACH_DATA_TEST"
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = sheetName
    ws.Range("A1:D1").value = Array("Code", "Rounded", "DateNarrow", "Note")
    ws.Range("A2").NumberFormat = "@"
    ws.Range("A2").value = "007"
    ws.Range("B2").value = 10.5
    ws.Range("B2").NumberFormat = "0"
    ws.Range("C2").value = DateSerial(2024, 2, 1)
    ws.Range("C2").NumberFormat = "yyyy-mm-dd"
    ws.Columns("C").ColumnWidth = 1
    ws.Range("D2").value = "first"
    ws.Range("A3:D3").value = Array("B", 20, DateSerial(2024, 2, 2), "skip")
    ws.Range("A4:D4").value = Array("C", 30, DateSerial(2024, 2, 3), "break")

    Dim arr(1 To 23, 1 To 5) As Variant
    Call MakeCmd(arr, 1, "SET", "{processed}", 0, "0")
    Call MakeCmd(arr, 2, "FOREACH_ROW", ThisWorkbook.FullName, sheetName, "")
    Call MakeCmd(arr, 3, "IF", "{$RIDX}", "EQ", "1")
    Call MakeCmd(arr, 4, "SET", "{first_code}", 0, "{$ROW.Code}")
    Call MakeCmd(arr, 5, "SET", "{first_rounded}", 0, "{$ROW.Rounded}")
    Call MakeCmd(arr, 6, "SET", "{first_date}", 0, "{$ROW.DateNarrow}")
    Call MakeCmd(arr, 7, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 8, "SET", "{last_code}", 0, "{$ROW.Code}")
    Call MakeCmd(arr, 9, "SET", "{last_idx}", 0, "{$RIDX}")
    Call MakeCmd(arr, 10, "SET", "{last_total}", 0, "{$RTOTAL}")
    Call MakeCmd(arr, 11, "IF", "{$ROW.Code}", "EQ", "B")
    Call MakeCmd(arr, 12, "CONTINUE", 0, 0, "")
    Call MakeCmd(arr, 13, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 14, "INC", "{processed}", 0, "")
    Call MakeCmd(arr, 15, "REPEAT", "2", 0, "")
    Call MakeCmd(arr, 16, "SET", "{nested_ridx}", 0, "{$RIDX}")
    Call MakeCmd(arr, 17, "SET", "{nested_total}", 0, "{$RTOTAL}")
    Call MakeCmd(arr, 18, "END_REPEAT", 0, 0, "")
    Call MakeCmd(arr, 19, "IF", "{$ROW.Code}", "EQ", "C")
    Call MakeCmd(arr, 20, "BREAK", 0, 0, "")
    Call MakeCmd(arr, 21, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 22, "END_FOREACH", 0, 0, "")
    Call MakeCmd(arr, 23, "EXIT", 0, 0, "")

    result = RunProgram(arr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "first_code", "007", fails
    ExpectVar "first_rounded", "11", fails
    ExpectVar "first_date", "2024-02-01", fails
    ExpectVar "last_code", "C", fails
    ExpectVar "last_idx", "3", fails
    ExpectVar "last_total", "3", fails
    ExpectVar "processed", "2", fails
    ExpectVar "nested_ridx", "3", fails
    ExpectVar "nested_total", "3", fails
    ExpectLoopDepth 0, fails

    ws.Range("A1:D4").AutoFilter Field:=1, Criteria1:="<>B"
    Dim visibleArr(1 To 11, 1 To 5) As Variant
    Call MakeCmd(visibleArr, 1, "FOREACH_ROW", ThisWorkbook.FullName, sheetName, "VISIBLE_ROWS")
    Call MakeCmd(visibleArr, 2, "SET", "{visible_last_idx}", 0, "{$RIDX}")
    Call MakeCmd(visibleArr, 3, "SET", "{visible_total}", 0, "{$RTOTAL}")
    Call MakeCmd(visibleArr, 4, "IF", "{$RIDX}", "EQ", "1")
    Call MakeCmd(visibleArr, 5, "SET", "{visible1}", 0, "{$ROW.Code}")
    Call MakeCmd(visibleArr, 6, "END_IF", 0, 0, "")
    Call MakeCmd(visibleArr, 7, "IF", "{$RIDX}", "EQ", "2")
    Call MakeCmd(visibleArr, 8, "SET", "{visible2}", 0, "{$ROW.Code}")
    Call MakeCmd(visibleArr, 9, "END_IF", 0, 0, "")
    Call MakeCmd(visibleArr, 10, "END_FOREACH", 0, 0, "")
    Call MakeCmd(visibleArr, 11, "EXIT", 0, 0, "")
    result = RunProgram(visibleArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "visible_last_idx", "2", fails
    ExpectVar "visible_total", "2", fails
    ExpectVar "visible1", "007", fails
    ExpectVar "visible2", "C", fails
    ExpectLoopDepth 0, fails
    ws.AutoFilterMode = False

    ws.Range("A1:D4").AutoFilter Field:=1, Criteria1:="<>C"
    Dim visibleTailArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(visibleTailArr, 1, "FOREACH_ROW", ThisWorkbook.FullName, sheetName, "VISIBLE_ROWS")
    Call MakeCmd(visibleTailArr, 2, "SET", "{visible_tail_last_idx}", 0, "{$RIDX}")
    Call MakeCmd(visibleTailArr, 3, "IF", "{$RIDX}", "EQ", "2")
    Call MakeCmd(visibleTailArr, 4, "SET", "{visible_tail2}", 0, "{$ROW.Code}")
    Call MakeCmd(visibleTailArr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(visibleTailArr, 6, "END_FOREACH", 0, 0, "")
    Call MakeCmd(visibleTailArr, 7, "EXIT", 0, 0, "")
    result = RunProgram(visibleTailArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "visible_tail_last_idx", "2", fails
    ExpectVar "visible_tail2", "B", fails
    ExpectLoopDepth 0, fails
    ws.AutoFilterMode = False

    ws.Range("A2:D4").ClearContents
    Dim emptyArr(1 To 5, 1 To 5) As Variant
    Call MakeCmd(emptyArr, 1, "FOREACH_ROW", ThisWorkbook.FullName, sheetName, "")
    Call MakeCmd(emptyArr, 2, "SET", "{empty_body}", 0, "bad")
    Call MakeCmd(emptyArr, 3, "END_FOREACH", 0, 0, "")
    Call MakeCmd(emptyArr, 4, "SET", "{empty_after}", 0, "yes")
    Call MakeCmd(emptyArr, 5, "EXIT", 0, 0, "")
    result = RunProgram(emptyArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "empty_body", "", fails
    ExpectVar "empty_after", "yes", fails
    ExpectLoopDepth 0, fails

CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    TC_ForeachRowExecution = fails
    Exit Function

ErrHandler:
    fails = fails + 1
    LogFail "FOREACH_ROW test", "no error", Err.Description
    Resume CleanUp
End Function

Private Function TC_FlowNestedInteractions() As Long
    Dim fails As Long
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult

    Debug.Print "--- [Flow 3] Nested and interaction paths ---"

    Dim breakArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(breakArr, 1, "SET", "{before_break}", 0, "0")
    Call MakeCmd(breakArr, 2, "REPEAT", "5", 0, "")
    Call MakeCmd(breakArr, 3, "INC", "{before_break}", 0, "")
    Call MakeCmd(breakArr, 4, "BREAK", 0, 0, "")
    Call MakeCmd(breakArr, 5, "INC", "{after_break}", 0, "")
    Call MakeCmd(breakArr, 6, "END_REPEAT", 0, 0, "")
    Call MakeCmd(breakArr, 7, "EXIT", 0, 0, "")
    result = RunProgram(breakArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "before_break", "1", fails
    ExpectVar "after_break", "", fails
    ExpectVisits trace, 5, 0, fails
    ExpectLoopDepth 0, fails

    Dim continueArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(continueArr, 1, "SET", "{continue_hits}", 0, "0")
    Call MakeCmd(continueArr, 2, "REPEAT", "3", 0, "")
    Call MakeCmd(continueArr, 3, "INC", "{continue_hits}", 0, "")
    Call MakeCmd(continueArr, 4, "CONTINUE", 0, 0, "")
    Call MakeCmd(continueArr, 5, "INC", "{continue_skipped}", 0, "")
    Call MakeCmd(continueArr, 6, "END_REPEAT", 0, 0, "")
    Call MakeCmd(continueArr, 7, "EXIT", 0, 0, "")
    result = RunProgram(continueArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "continue_hits", "3", fails
    ExpectVar "continue_skipped", "", fails
    ExpectVisits trace, 5, 0, fails
    ExpectLoopDepth 0, fails

    Dim peakArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(peakArr, 1, "REPEAT", "2", 0, "")
    Call MakeCmd(peakArr, 2, "REPEAT", "2", 0, "")
    Call MakeCmd(peakArr, 3, "SET", "{inner_loop}", 0, "{$LOOP}")
    Call MakeCmd(peakArr, 4, "END_REPEAT", 0, 0, "")
    Call MakeCmd(peakArr, 5, "END_REPEAT", 0, 0, "")
    Call MakeCmd(peakArr, 6, "EXIT", 0, 0, "")
    result = RunProgram(peakArr, 1, 3, trace, lastCode, lastLine)
    ExpectTerminal rrExhausted, result, fails
    ExpectLoopDepth 2, fails

    Dim nestedArr(1 To 15, 1 To 5) As Variant
    Call MakeCmd(nestedArr, 1, "SET", "{outer1}", 0, "0")
    Call MakeCmd(nestedArr, 2, "SET", "{outer2}", 0, "0")
    Call MakeCmd(nestedArr, 3, "SET", "{inner_hits}", 0, "0")
    Call MakeCmd(nestedArr, 4, "REPEAT", "2", 0, "")
    Call MakeCmd(nestedArr, 5, "SELECT", "{$LOOP}", 0, "")
    Call MakeCmd(nestedArr, 6, "CASE", "1", 0, "")
    Call MakeCmd(nestedArr, 7, "INC", "{outer1}", 0, "")
    Call MakeCmd(nestedArr, 8, "CASE", "2", 0, "")
    Call MakeCmd(nestedArr, 9, "INC", "{outer2}", 0, "")
    Call MakeCmd(nestedArr, 10, "END_SELECT", 0, 0, "")
    Call MakeCmd(nestedArr, 11, "REPEAT", "2", 0, "")
    Call MakeCmd(nestedArr, 12, "INC", "{inner_hits}", 0, "")
    Call MakeCmd(nestedArr, 13, "END_REPEAT", 0, 0, "")
    Call MakeCmd(nestedArr, 14, "END_REPEAT", 0, 0, "")
    Call MakeCmd(nestedArr, 15, "EXIT", 0, 0, "")
    result = RunProgram(nestedArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "outer1", "1", fails
    ExpectVar "outer2", "1", fails
    ExpectVar "inner_hits", "4", fails
    ExpectLoopDepth 0, fails

    Dim callLoopArr(1 To 8, 1 To 5) As Variant
    Call MakeCmd(callLoopArr, 1, "SET", "{call_loop_hits}", 0, "0")
    Call MakeCmd(callLoopArr, 2, "REPEAT", "2", 0, "")
    Call MakeCmd(callLoopArr, 3, "CALL", 0, 0, "SUB")
    Call MakeCmd(callLoopArr, 4, "END_REPEAT", 0, 0, "")
    Call MakeCmd(callLoopArr, 5, "EXIT", 0, 0, "")
    Call MakeCmd(callLoopArr, 6, "LABEL", 0, 0, "SUB")
    Call MakeCmd(callLoopArr, 7, "INC", "{call_loop_hits}", 0, "")
    Call MakeCmd(callLoopArr, 8, "RETURN", 0, 0, "")
    result = RunProgram(callLoopArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "call_loop_hits", "2", fails
    ExpectStackDepth 0, fails
    ExpectLoopDepth 0, fails

    Dim nestedCallArr(1 To 10, 1 To 5) As Variant
    Call MakeCmd(nestedCallArr, 1, "CALL", 0, 0, "A")
    Call MakeCmd(nestedCallArr, 2, "SET", "{done}", 0, "yes")
    Call MakeCmd(nestedCallArr, 3, "EXIT", 0, 0, "")
    Call MakeCmd(nestedCallArr, 4, "LABEL", 0, 0, "A")
    Call MakeCmd(nestedCallArr, 5, "CALL", 0, 0, "B")
    Call MakeCmd(nestedCallArr, 6, "SET", "{after_b}", 0, "yes")
    Call MakeCmd(nestedCallArr, 7, "RETURN", 0, 0, "")
    Call MakeCmd(nestedCallArr, 8, "LABEL", 0, 0, "B")
    Call MakeCmd(nestedCallArr, 9, "SET", "{in_b}", 0, "yes")
    Call MakeCmd(nestedCallArr, 10, "RETURN", 0, 0, "")
    result = RunProgram(nestedCallArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "in_b", "yes", fails
    ExpectVar "after_b", "yes", fails
    ExpectVar "done", "yes", fails
    ExpectStackDepth 0, fails
    ExpectTrace trace, Array(1, 4, 5, 8, 9, 10, 6, 7, 2, 3), fails

    Dim nestedElseIfArr(1 To 10, 1 To 5) As Variant
    Call MakeCmd(nestedElseIfArr, 1, "IF", "outer", "EQ", "outer")
    Call MakeCmd(nestedElseIfArr, 2, "IF", "inner", "EQ", "miss")
    Call MakeCmd(nestedElseIfArr, 3, "SET", "{nested_inner_if}", 0, "bad")
    Call MakeCmd(nestedElseIfArr, 4, "ELSE_IF", "inner", "EQ", "inner")
    Call MakeCmd(nestedElseIfArr, 5, "SET", "{nested_inner_hit}", 0, "yes")
    Call MakeCmd(nestedElseIfArr, 6, "END_IF", 0, 0, "")
    Call MakeCmd(nestedElseIfArr, 7, "ELSE_IF", "outer2", "EQ", "outer2")
    Call MakeCmd(nestedElseIfArr, 8, "SET", "{nested_outer_late}", 0, "bad")
    Call MakeCmd(nestedElseIfArr, 9, "END_IF", 0, 0, "")
    Call MakeCmd(nestedElseIfArr, 10, "EXIT", 0, 0, "")
    result = RunProgram(nestedElseIfArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 5, 1, fails
    ExpectVisits trace, 8, 0, fails
    ExpectVar "nested_inner_hit", "yes", fails
    ExpectIfBranchDepth 0, fails

    Dim gotoIfArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(gotoIfArr, 1, "IF", "1", "EQ", "1")
    Call MakeCmd(gotoIfArr, 2, "GOTO", 0, 0, "OUT_IF")
    Call MakeCmd(gotoIfArr, 3, "ELSE_IF", "1", "EQ", "1")
    Call MakeCmd(gotoIfArr, 4, "SET", "{goto_if_bad}", 0, "bad")
    Call MakeCmd(gotoIfArr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(gotoIfArr, 6, "LABEL", 0, 0, "OUT_IF")
    Call MakeCmd(gotoIfArr, 7, "EXIT", 0, 0, "")
    result = RunProgram(gotoIfArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 4, 0, fails
    ExpectIfBranchDepth 0, fails

    Dim breakIfArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(breakIfArr, 1, "REPEAT", "3", 0, "")
    Call MakeCmd(breakIfArr, 2, "IF", "1", "EQ", "1")
    Call MakeCmd(breakIfArr, 3, "BREAK", 0, 0, "")
    Call MakeCmd(breakIfArr, 4, "END_IF", 0, 0, "")
    Call MakeCmd(breakIfArr, 5, "INC", "{break_if_bad}", 0, "")
    Call MakeCmd(breakIfArr, 6, "END_REPEAT", 0, 0, "")
    Call MakeCmd(breakIfArr, 7, "EXIT", 0, 0, "")
    result = RunProgram(breakIfArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 5, 0, fails
    ExpectLoopDepth 0, fails
    ExpectIfBranchDepth 0, fails

    Dim continueIfArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(continueIfArr, 1, "REPEAT", "1", 0, "")
    Call MakeCmd(continueIfArr, 2, "IF", "1", "EQ", "1")
    Call MakeCmd(continueIfArr, 3, "CONTINUE", 0, 0, "")
    Call MakeCmd(continueIfArr, 4, "END_IF", 0, 0, "")
    Call MakeCmd(continueIfArr, 5, "INC", "{continue_if_bad}", 0, "")
    Call MakeCmd(continueIfArr, 6, "END_REPEAT", 0, 0, "")
    Call MakeCmd(continueIfArr, 7, "EXIT", 0, 0, "")
    result = RunProgram(continueIfArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 5, 0, fails
    ExpectLoopDepth 0, fails
    ExpectIfBranchDepth 0, fails

    Dim returnIfArr(1 To 8, 1 To 5) As Variant
    Call MakeCmd(returnIfArr, 1, "CALL", 0, 0, "SUB_IF")
    Call MakeCmd(returnIfArr, 2, "EXIT", 0, 0, "")
    Call MakeCmd(returnIfArr, 3, "LABEL", 0, 0, "SUB_IF")
    Call MakeCmd(returnIfArr, 4, "IF", "1", "EQ", "1")
    Call MakeCmd(returnIfArr, 5, "RETURN", 0, 0, "")
    Call MakeCmd(returnIfArr, 6, "ELSE_IF", "1", "EQ", "1")
    Call MakeCmd(returnIfArr, 7, "SET", "{return_if_bad}", 0, "bad")
    Call MakeCmd(returnIfArr, 8, "END_IF", 0, 0, "")
    result = RunProgram(returnIfArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 7, 0, fails
    ExpectStackDepth 0, fails
    ExpectIfBranchDepth 0, fails

    TC_FlowNestedInteractions = fails
End Function

Private Function TC_FlowNegativePaths() As Long
    Dim fails As Long
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult

    Debug.Print "--- [Flow 4] Negative paths ---"

    Dim arr1(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr1, 1, "GOTO", 0, 0, "MISSING")
    result = RunProgram(arr1, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr2(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr2, 1, "CALL", 0, 0, "MISSING")
    result = RunProgram(arr2, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr3(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr3, 1, "RETURN", 0, 0, "")
    result = RunProgram(arr3, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr4(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr4, 1, "BREAK", 0, 0, "")
    result = RunProgram(arr4, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr5(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr5, 1, "CONTINUE", 0, 0, "")
    result = RunProgram(arr5, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr6(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr6, 1, "SELECT", "A", 0, "")
    result = RunProgram(arr6, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr7(1 To 2, 1 To 5) As Variant
    Call MakeCmd(arr7, 1, "WHILE", "1", "EQ", "1")
    Call MakeCmd(arr7, 2, "END_WHILE", 0, 0, "")
    result = RunProgram(arr7, 1, 5, trace, lastCode, lastLine)
    ExpectTerminal rrExhausted, result, fails
    ExpectVisits trace, 1, 3, fails

    Dim arr8(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr8, 1, "SET", "{$LOOP}", 0, "bad")
    result = RunProgram(arr8, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr9(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr9, 1, "SET", "LOOP", 0, "bad")
    result = RunProgram(arr9, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr10(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr10, 1, "GOTO", 0, 0, "1")
    result = RunProgram(arr10, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr11(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr11, 1, "GOTO", 0, 0, "{target}")
    result = RunProgram(arr11, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr12(1 To 5, 1 To 5) As Variant
    Call MakeCmd(arr12, 1, "GOTO", 0, 0, "INSIDE")
    Call MakeCmd(arr12, 2, "REPEAT", "1", 0, "")
    Call MakeCmd(arr12, 3, "LABEL", 0, 0, "INSIDE")
    Call MakeCmd(arr12, 4, "END_REPEAT", 0, 0, "")
    Call MakeCmd(arr12, 5, "EXIT", 0, 0, "")
    result = RunProgram(arr12, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr13(1 To 8, 1 To 5) As Variant
    Call MakeCmd(arr13, 1, "CALL", 0, 0, "SUB")
    Call MakeCmd(arr13, 2, "EXIT", 0, 0, "")
    Call MakeCmd(arr13, 3, "LABEL", 0, 0, "SUB")
    Call MakeCmd(arr13, 4, "GOTO", 0, 0, "OUTSIDE")
    Call MakeCmd(arr13, 5, "RETURN", 0, 0, "")
    Call MakeCmd(arr13, 6, "LABEL", 0, 0, "OUTSIDE")
    Call MakeCmd(arr13, 7, "SET", "{bad_sub_goto}", 0, "bad")
    Call MakeCmd(arr13, 8, "EXIT", 0, 0, "")
    result = RunProgram(arr13, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr14(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr14, 1, "SET", "{out}", 0, "{missing}")
    result = RunProgram(arr14, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr15(1 To 1, 1 To 5) As Variant
    Call MakeCmd(arr15, 1, "INC", "{n}", 0, "")
    result = RunProgram(arr15, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails

    Dim arr16(1 To 2, 1 To 5) As Variant
    Call MakeCmd(arr16, 1, "SET", "{score}", 0, "1")
    Call MakeCmd(arr16, 2, "SET", "{out}", 0, "{scroe}")
    result = RunProgram(arr16, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails
    ExpectBool "Undefined var suggestion", True, (InStr(1, modEngine.LastRuntimeError(), "{score}", vbTextCompare) > 0), fails

    TC_FlowNegativePaths = fails
End Function

Private Function TC_MemoryResolverHeadless() As Long
    Dim fails As Long
    Dim arr(1 To 22, 1 To 5) As Variant
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult
    Dim pxExpected As String

    Debug.Print "--- [Flow 5] Memory and resolver ---"
    Call MakeCmd(arr, 1, "SET", "{a}", 0, "hello")
    Call MakeCmd(arr, 2, "SET", "{b}", 0, "{a}_world")
    Call MakeCmd(arr, 3, "SET", "{a}", 0, "overwritten")
    Call MakeCmd(arr, 4, "SET", "{a}", 0, "{$BLANK}")
    Call MakeCmd(arr, 5, "SET", "{empty_copy}", 0, "{a}")
    Call MakeCmd(arr, 6, "SET", "{cr}", 0, "A{$CR}B")
    Call MakeCmd(arr, 7, "SET", "{lf}", 0, "A{$LF}B")
    Call MakeCmd(arr, 8, "SET", "{sp}", 0, "A{$SPACE}B")
    Call MakeCmd(arr, 9, "SET", "{unknown}", 0, "{$XXX}")
    Call MakeCmd(arr, 10, "SET", "{ridx_out}", 0, "{$RIDX}")
    Call MakeCmd(arr, 11, "SET", "{rtotal_out}", 0, "{$RTOTAL}")
    Call MakeCmd(arr, 12, "SET", "{row_out}", 0, "{$ROW.Code}")
    Call MakeCmd(arr, 13, "SET", "{open}", 0, "A{broken")
    Call MakeCmd(arr, 14, "SET", "{x}", 0, "300")
    Call MakeCmd(arr, 15, "SET", "{y}", 0, "300")
    Call MakeCmd(arr, 16, "SET", "{px}", 0, "{$PIXEL:{x},{y}}")
    Call MakeCmd(arr, 17, "SET", "{delta}", 0, "-2")
    Call MakeCmd(arr, 18, "SET", "{n}", 0, "0")
    Call MakeCmd(arr, 19, "INC", "{n}", 0, "")
    Call MakeCmd(arr, 20, "INC", "{n}", 0, "{delta}")
    Call MakeCmd(arr, 21, "DEC", "{n}", 0, "-3")
    Call MakeCmd(arr, 22, "EXIT", 0, 0, "")

    result = RunProgram(arr, 1, 0, trace, lastCode, lastLine)
    pxExpected = modWinAPI.GetColorHexAt(300, 300)

    ExpectTerminal rrHalted, result, fails
    ExpectVar "a", "", fails
    ExpectBool "Defined empty var", True, modContext.HasVar("empty_copy"), fails
    ExpectVar "empty_copy", "", fails
    ExpectVar "b", "hello_world", fails
    ExpectVar "cr", "A" & vbCr & "B", fails
    ExpectVar "lf", "A" & vbLf & "B", fails
    ExpectVar "sp", "A B", fails
    ExpectVar "unknown", "{$XXX}", fails
    ExpectVar "ridx_out", "{$RIDX}", fails
    ExpectVar "rtotal_out", "{$RTOTAL}", fails
    ExpectVar "row_out", "{$ROW.Code}", fails
    ExpectVar "open", "A{broken", fails
    ExpectVar "px", pxExpected, fails
    ExpectVar "n", "2", fails
    ExpectTrace trace, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22), fails

    TC_MemoryResolverHeadless = fails
End Function

Private Function TC_ParamSpecValidation() As Long
    Dim fails As Long

    Debug.Print "--- [Flow 6] ParamSpec validation ---"

    ExpectBool "Validate SLEEP P3", True, modDevTools.ValidateLine("SLEEP", "", "", "1000"), fails
    ExpectBool "Reject SLEEP shifted P1", False, modDevTools.ValidateLine("SLEEP", "1000", "", ""), fails
    ExpectBool "Validate NOTIFY seconds/message", True, modDevTools.ValidateLine("NOTIFY", "3", "", "done"), fails
    ExpectBool "Validate NOTIFY default seconds", True, modDevTools.ValidateLine("NOTIFY", "", "", "done"), fails
    ExpectBool "Validate NOTIFY wait for OK", True, modDevTools.ValidateLine("NOTIFY", "0", "", "done"), fails
    ExpectBool "Reject NOTIFY non-numeric seconds", False, modDevTools.ValidateLine("NOTIFY", "abc", "", "done"), fails
    ExpectBool "Allow NOTIFY literal pipe", True, modDevTools.ValidateLine("NOTIFY", "", "", "done|3"), fails
    ExpectBool "Validate TYPE P1/P3", True, modDevTools.ValidateLine("TYPE", "10", "", "hello"), fails
    ExpectBool "Reject TYPE shifted P2", False, modDevTools.ValidateLine("TYPE", "", "hello", ""), fails
    ExpectBool "Reject numeric literal mismatch", False, modDevTools.ValidateLine("MOVE", "10", "20", "later"), fails
    ExpectBool "Allow numeric token", True, modDevTools.ValidateLine("MOVE", "{x}", "20", ""), fails
    ExpectBool "Reject unused slot", False, modDevTools.ValidateLine("KEY", "ENTER", "unused", ""), fails
    ExpectBool "Validate IF operator", True, modDevTools.ValidateLine("IF", "abc", "EQ", "abc"), fails
    ExpectBool "Reject IF operator", False, modDevTools.ValidateLine("IF", "abc", "EQUALS", "abc"), fails
    ExpectBool "Validate SET memory lvalue", True, modDevTools.ValidateLine("SET", "{name}", "", "value"), fails
    ExpectBool "Validate SET clipboard lvalue", True, modDevTools.ValidateLine("SET", "{$CLIP}", "", "value"), fails
    ExpectBool "Reject read-only lvalue", False, modDevTools.ValidateLine("INC", "{$LOOP}", "", ""), fails
    ExpectBool "Reject raw label token", False, modDevTools.ValidateLine("LABEL", "", "", "{dynamic}"), fails
    ExpectBool "Validate GOTO static label", True, modDevTools.ValidateLine("GOTO", "", "", "DONE"), fails
    ExpectBool "Reject GOTO numeric target", False, modDevTools.ValidateLine("GOTO", "", "", "12"), fails
    ExpectBool "Reject CALL dynamic target", False, modDevTools.ValidateLine("CALL", "", "", "{target}"), fails

    TC_ParamSpecValidation = fails
End Function

' [Flow 5C] ParamSpec 런타임 게이트 (C1): Required=원문 기준, DT_NUM=해석값 기준
Private Function TC_ParamGuardRuntime() As Long
    Dim fails As Long
    Dim arrCmd(1 To 1, 1 To 5) As Variant
    Dim code As Long

    Debug.Print "--- [Flow 5C] ParamSpec runtime guard ---"
    modContext.InitContext

    ' 필수 누락: SET P3(값) 원문 공백 -> 게이트 차단
    Call MakeCmd(arrCmd, 1, "SET", "{pg}", 0, "")
    code = ExecuteArrayLine(arrCmd, 1)
    ExpectLong "Guard: SET empty P3", CMD_ERROR, code, fails

    ' 의도된 빈 값: {$BLANK}는 원문이 비어있지 않으므로 통과 (회귀 가드)
    Call MakeCmd(arrCmd, 1, "SET", "{pg}", 0, "{$BLANK}")
    code = ExecuteArrayLine(arrCmd, 1)
    ExpectLong "Guard: SET {$BLANK} passes", CMD_NEXT_LINE, code, fails
    ExpectVar "pg", "", fails

    ' 필수 숫자 슬롯: SLEEP P3 비숫자 차단 / 숫자 통과
    Call MakeCmd(arrCmd, 1, "SLEEP", 0, 0, "abc")
    ExpectLong "Guard: SLEEP non-numeric P3", CMD_ERROR, ExecuteArrayLine(arrCmd, 1), fails
    Call MakeCmd(arrCmd, 1, "SLEEP", 0, 0, "10")
    ExpectLong "Guard: SLEEP numeric P3", CMD_NEXT_LINE, ExecuteArrayLine(arrCmd, 1), fails

    ' 좌표 슬롯: CLICK P1 비숫자 / P2 공백 -> 핸들러 진입 전 차단 (마우스 미동작)
    Call MakeCmd(arrCmd, 1, "CLICK", "abc", "10", "")
    ExpectLong "Guard: CLICK non-numeric P1", CMD_ERROR, ExecuteArrayLine(arrCmd, 1), fails
    Call MakeCmd(arrCmd, 1, "CLICK", "10", "", "")
    ExpectLong "Guard: CLICK empty P2", CMD_ERROR, ExecuteArrayLine(arrCmd, 1), fails

    ' 선택 숫자 슬롯: INC 빈 P3 허용(기본 1) / 비숫자 P3 차단
    Call MakeCmd(arrCmd, 1, "SET", "{pgn}", 0, "1")
    ExpectLong "Guard: SET pgn=1", CMD_NEXT_LINE, ExecuteArrayLine(arrCmd, 1), fails
    Call MakeCmd(arrCmd, 1, "INC", "{pgn}", 0, "")
    ExpectLong "Guard: INC empty P3 allowed", CMD_NEXT_LINE, ExecuteArrayLine(arrCmd, 1), fails
    ExpectVar "pgn", "2", fails
    Call MakeCmd(arrCmd, 1, "INC", "{pgn}", 0, "xyz")
    ExpectLong "Guard: INC non-numeric P3", CMD_ERROR, ExecuteArrayLine(arrCmd, 1), fails

    ' 토큰 해석 후 숫자 판정: {n}이 숫자면 통과, 비숫자면 차단
    Call MakeCmd(arrCmd, 1, "SET", "{pgv}", 0, "250")
    ExpectLong "Guard: SET pgv=250", CMD_NEXT_LINE, ExecuteArrayLine(arrCmd, 1), fails
    Call MakeCmd(arrCmd, 1, "SLEEP", 0, 0, "{pgv}")
    ExpectLong "Guard: SLEEP {token} numeric", CMD_NEXT_LINE, ExecuteArrayLine(arrCmd, 1), fails
    Call MakeCmd(arrCmd, 1, "SET", "{pgv}", 0, "not_a_number")
    ExpectLong "Guard: SET pgv=text", CMD_NEXT_LINE, ExecuteArrayLine(arrCmd, 1), fails
    Call MakeCmd(arrCmd, 1, "SLEEP", 0, 0, "{pgv}")
    ExpectLong "Guard: SLEEP {token} non-numeric", CMD_ERROR, ExecuteArrayLine(arrCmd, 1), fails

    ' RAW 필수 슬롯: IF 연산자(P2) 공백 -> rrErrored
    Dim arr3(1 To 3, 1 To 5) As Variant
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult
    Call MakeCmd(arr3, 1, "IF", "1", "", "1")
    Call MakeCmd(arr3, 2, "END_IF", 0, 0, "")
    Call MakeCmd(arr3, 3, "EXIT", 0, 0, "")
    result = RunProgram(arr3, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrErrored, result, fails
    ExpectLong "Guard: IF empty operator line", 1, lastLine, fails

    TC_ParamGuardRuntime = fails
End Function

' [Flow 5E] 중단 래치 (B2): 물리키 없이 RequestAbort 주입으로 조기 탈출 검증
Private Function TC_AbortLatch() As Long
    Dim fails As Long
    Dim t0 As Long
    Dim ok As Boolean

    Debug.Print "--- [Flow 5E] Abort latch ---"

    ' 래치 기본 상태 (물리 ESC 미입력 가정)
    modActions.ClearAbort
    ExpectBool "Abort: clear state", False, modActions.AbortPending(), fails

    ' 래치 설정 후 유지 (순간 신호 유실 방지의 핵심)
    modActions.RequestAbort
    ExpectBool "Abort: latched", True, modActions.AbortPending(), fails
    ExpectBool "Abort: latch persists", True, modActions.AbortPending(), fails

    ' BetterWait 조기 탈출: 1000ms 요청이 즉시 반환
    t0 = modWinAPI.GetTickCount()
    modActions.BetterWait 1000
    ExpectBool "Abort: BetterWait early exit", True, _
        (modWinAPI.GetTickCount() - t0) < 500, fails

    ' Action_Wait 조기 탈출: 30초 타임아웃 대신 즉시 False
    t0 = modWinAPI.GetTickCount()
    ok = modActions.Action_Wait("WAIT_APPEAR", 0, 0, "#ZZZZZZ")
    ExpectBool "Abort: Action_Wait returns False", False, ok, fails
    ExpectBool "Abort: Action_Wait early exit", True, _
        (modWinAPI.GetTickCount() - t0) < 1000, fails

    ' 해제 확인
    modActions.ClearAbort
    ExpectBool "Abort: cleared", False, modActions.AbortPending(), fails

    TC_AbortLatch = fails
End Function

Private Function TC_StaticVariableAnalysis() As Long
    On Error GoTo ErrHandler

    Dim fails As Long
    Dim ws As Worksheet
    Dim sheetName As String

    Debug.Print "--- [Flow 6A] Static variable analysis ---"

    sheetName = "RPA_VAR_ANALYSIS_TEST"
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = sheetName
    modDevTools.EnsureHeaders ws

    ws.Cells(2, 2).value = "SET"
    ws.Cells(2, 3).value = "{score}"
    ws.Cells(2, 5).value = "1"

    ws.Cells(3, 2).value = "SET"
    ws.Cells(3, 3).value = "{out}"
    ws.Cells(3, 5).value = "{scroe}"

    ws.Cells(4, 2).value = "SET"
    ws.Cells(4, 3).value = "{unused}"
    ws.Cells(4, 5).value = "value"

    ws.Cells(5, 2).value = "SET"
    ws.Cells(5, 3).value = "{sys}"
    ws.Cells(5, 5).value = "{$BLANK}{$LOOP}{$RIDX}{$RTOTAL}{$ROW.Code}{$PIXEL:1,1}{$XXX}"

    ExpectBool "Static undefined var", False, modDevTools.ValidateSheetVariablesAt(ws, 3), fails
    ExpectBool "Static var suggestion", True, (InStr(1, modDevTools.LastValidationMessage(), "{score}", vbTextCompare) > 0), fails

    ExpectBool "Unused var is warning only", True, modDevTools.ValidateSheetVariablesAt(ws, 4), fails
    ExpectBool "Unused var warning text", True, (InStr(1, modDevTools.LastValidationWarning(), "{unused}", vbTextCompare) > 0), fails

    ExpectBool "System tokens ignored by var analysis", True, modDevTools.ValidateSheetVariablesAt(ws, 5), fails

CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    TC_StaticVariableAnalysis = fails
    Exit Function

ErrHandler:
    fails = fails + 1
    LogFail "Static variable analysis", "no error", Err.Description
    Resume CleanUp
End Function

Private Function TC_ScriptAssistFormatterRegression() As Long
    On Error GoTo ErrHandler

    Dim fails As Long
    Dim ws As Worksheet
    Dim sheetName As String

    Debug.Print "--- [Flow 6B] Script Assist formatter comments ---"

    sheetName = "RPA_FORMAT_COMMENT_TEST"

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = sheetName
    modDevTools.EnsureHeaders ws

    ws.Cells(2, 2).value = "IF"
    ws.Cells(2, 3).value = "1"
    ws.Cells(2, 4).value = "EQ"
    ws.Cells(2, 5).value = "1"
    ws.Cells(3, 2).value = "#> TYPE"
    ws.Cells(3, 3).value = "0"
    ws.Cells(3, 5).value = "inside"
    ws.Cells(4, 2).value = "END_IF"

    modDevTools.FormatRPAScript ws
    ExpectText "Formatter normalizes #> indent", "# > TYPE", CStr(ws.Cells(3, 2).value), fails

    modDevTools.FormatRPAScript ws
    ExpectText "Formatter keeps comment stable", "# > TYPE", CStr(ws.Cells(3, 2).value), fails

    Dim program As clsScriptProgram
    Set program = New clsScriptProgram
    program.LoadFromRange ws.Range(ws.Cells(2, 1), ws.Cells(4, 6))
    ExpectBool "Formatter comment remains skippable", True, program.IsCommentLine(2), fails

    modScriptAssist.ToggleSelectedCommandComment ws.Cells(3, 2)
    modDevTools.FormatRPAScript ws
    ExpectText "Formatter restores uncommented indent", "> TYPE", CStr(ws.Cells(3, 2).value), fails

CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    Application.StatusBar = False
    On Error GoTo 0
    TC_ScriptAssistFormatterRegression = fails
    Exit Function

ErrHandler:
    fails = fails + 1
    LogFail "Script Assist formatter comments", "no error", Err.Description
    Resume CleanUp
End Function

Private Function TC_StringDistance() As Long
    Dim fails As Long

    Debug.Print "--- [Util 1A] String distance ---"

    ExpectLong "Distance: identical", 0, modStringDistance.LevenshteinDistance("CLICK", "CLICK"), fails
    ExpectLong "Distance: one edit", 1, modStringDistance.LevenshteinDistance("CLIK", "CLICK"), fails
    ExpectLong "Distance: case and trim ignored", 0, modStringDistance.LevenshteinDistance("  click ", "CLICK"), fails

    ExpectText "Closest: unique match", "CLICK", modStringDistance.ClosestName("CLIK", Array("CLICK", "MOVE")), fails
    ExpectText "Closest: tie returns empty", "", modStringDistance.ClosestName("AB", Array("ABC", "ABD")), fails
    ExpectText "Closest: out of threshold", "", modStringDistance.ClosestName("ZZZZZZ", Array("CLICK")), fails

    ' GetClosestCommand 위임 회귀: 오타는 제안, 정확 일치는 빈값
    ExpectText "Assist: typo suggestion", "CLICK", modScriptAssist.GetClosestCommand("CLIK"), fails
    ExpectText "Assist: exact match no hint", "", modScriptAssist.GetClosestCommand("CLICK"), fails

    TC_StringDistance = fails
End Function

Private Function TC_CommandNormalization() As Long
    Dim fails As Long

    Debug.Print "--- [Util 1B] Command normalization ---"

    ' 선행 들여쓰기 마커(공백/탭/">")는 벗긴다
    ExpectText "Norm: leading marker", "CLICK", modCmdRegistry.NormalizeCommand(">CLICK"), fails
    ExpectText "Norm: double marker", "CLICK", modCmdRegistry.NormalizeCommand(">> CLICK"), fails
    ExpectText "Norm: spaced markers", "IF", modCmdRegistry.NormalizeCommand("> > if"), fails
    ExpectText "Norm: tab mixed", "MOVE", modCmdRegistry.NormalizeCommand(vbTab & "> MOVE"), fails
    ExpectText "Norm: marker only", "", modCmdRegistry.NormalizeCommand(">"), fails

    ' 내부/후행 ">"는 명령 문자열의 일부로 보존한다
    ExpectText "Norm: internal preserved", "A>B", modCmdRegistry.NormalizeCommand("a>b"), fails
    ExpectText "Norm: comment preserved", "#> TYPE", modCmdRegistry.NormalizeCommand("#> type"), fails

    ' GetCmdDef 게이트 회귀: 오타는 미등록, 들여쓰기 명령은 정상 조회
    ExpectBool "Gate: corrupted name rejected", True, (modCmdRegistry.GetCmdDef("CLI>CK") Is Nothing), fails
    ExpectBool "Gate: indented command resolves", False, (modCmdRegistry.GetCmdDef(">> CLICK") Is Nothing), fails

    TC_CommandNormalization = fails
End Function

' [Helper] 가상 명령어 배열 생성기 (B=cmd, C=P1, D=P2, E=P3)
Private Sub MakeCmd(ByRef arr As Variant, R As Long, cmd As String, p1 As Variant, p2 As Variant, p3 As String)
    arr(R, 2) = cmd
    arr(R, 3) = p1
    arr(R, 4) = p2
    arr(R, 5) = p3
End Sub

Private Function ExecuteArrayLine(ByRef arr As Variant, ByVal R As Long) As Long
    modActions.SetUnattendedMode True   ' 단일 라인 하니스도 무인 모드
    modActions.ClearAbort               ' 래치 누수 방지
    modEngine.LoadProgramFromArray arr, 1
    ExecuteArrayLine = modEngine.ExecuteSmartAction(R)
End Function

Private Sub AppendTrace(ByRef trace() As Long, ByVal traceCount As Long, ByVal lineNo As Long)
    ReDim Preserve trace(1 To traceCount)
    trace(traceCount) = lineNo
End Sub

Private Function TC_ScriptProgramModel() As Long
    Dim fails As Long
    Dim arr(1 To 26, 1 To 5) As Variant

    Call MakeCmd(arr, 1, "> IF", "1", "EQ", "0")
    Call MakeCmd(arr, 2, "TYPE", 0, 0, "if-body")
    Call MakeCmd(arr, 3, "ELSE", 0, 0, "")
    Call MakeCmd(arr, 4, "TYPE", 0, 0, "else-body")
    Call MakeCmd(arr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 6, "WHILE", "1", "EQ", "1")
    Call MakeCmd(arr, 7, "END_WHILE", 0, 0, "")
    Call MakeCmd(arr, 8, "REPEAT", 2, 0, "")
    Call MakeCmd(arr, 9, "END_REPEAT", 0, 0, "")
    Call MakeCmd(arr, 10, "SELECT", "A", 0, "")
    Call MakeCmd(arr, 11, "CASE", "A", 0, "")
    Call MakeCmd(arr, 12, "TYPE", 0, 0, "A")
    Call MakeCmd(arr, 13, "SELECT", "B", 0, "")
    Call MakeCmd(arr, 14, "CASE", "B", 0, "")
    Call MakeCmd(arr, 15, "END_SELECT", 0, 0, "")
    Call MakeCmd(arr, 16, "CASE", "ELSE", 0, "")
    Call MakeCmd(arr, 17, "TYPE", 0, 0, "fallback")
    Call MakeCmd(arr, 18, "END_SELECT", 0, 0, "")
    Call MakeCmd(arr, 19, "LABEL", 0, 0, "LOGIN")
    Call MakeCmd(arr, 20, "GOTO", 0, 0, "LOGIN")
    Call MakeCmd(arr, 21, "# real comment", 0, 0, "")
    Call MakeCmd(arr, 22, "> # not comment", 0, 0, "")
    Call MakeCmd(arr, 23, ">", 0, 0, "")
    Call MakeCmd(arr, 24, "", 0, 0, "")
    Call MakeCmd(arr, 25, "FOREACH_ROW", "Book", "Sheet", "")
    Call MakeCmd(arr, 26, "END_FOREACH", 0, 0, "")

    Dim program As clsScriptProgram
    Set program = New clsScriptProgram
    program.LoadFromArray arr, 5
    If Not program.ValidateStructure() Then
        LogFail "Program ValidateStructure", "True", program.ValidationErrors(): fails = fails + 1
    Else
        LogPass "Program ValidateStructure"
    End If
    If Not program.ValidateLabels() Then
        LogFail "Program ValidateLabels", "True", program.ValidationErrors(): fails = fails + 1
    Else
        LogPass "Program ValidateLabels"
    End If

    ExpectLong "TotalLines", 26, program.totalLines, fails
    ExpectLong "SheetRowFromRuntimeLine", 5, program.SheetRowFromRuntimeLine(1), fails
    ExpectLong "RuntimeLineFromSheetRow", 1, program.RuntimeLineFromSheetRow(5), fails
    ExpectText "CmdAt indent", "IF", program.CmdAt(1), fails
    ExpectBool "raw comment true", True, program.IsCommentLine(21), fails
    ExpectBool "indented comment false", False, program.IsCommentLine(22), fails
    ExpectBool "raw > is not blank", False, program.IsBlankLine(23), fails
    ExpectBool "empty raw blank", True, program.IsBlankLine(24), fails

    ExpectLong "IF -> ELSE", 3, program.FindBlockPair(1, "IF", 1), fails
    ExpectLong "ELSE -> END_IF", 5, program.FindBlockPair(3, "IF", 1), fails
    ExpectLong "WHILE -> END_WHILE", 7, program.FindBlockPair(6, "WHILE", 1), fails
    ExpectLong "END_WHILE -> WHILE", 6, program.FindBlockPair(7, "WHILE", -1), fails
    ExpectLong "REPEAT -> END_REPEAT", 9, program.FindBlockPair(8, "REPEAT", 1), fails
    ExpectLong "END_REPEAT -> REPEAT", 8, program.FindBlockPair(9, "REPEAT", -1), fails
    ExpectLong "FOREACH_ROW -> END_FOREACH", 26, program.FindBlockPair(25, "FOREACH", 1), fails
    ExpectLong "END_FOREACH -> FOREACH_ROW", 25, program.FindBlockPair(26, "FOREACH", -1), fails
    ExpectLong "FindBlockPair rejects SELECT", CMD_ERROR, program.FindBlockPair(10, "SELECT", 1), fails

    Dim outerBlock As clsSelectBlock
    Set outerBlock = program.SelectBlock(10)
    If outerBlock Is Nothing Then
        LogFail "Outer SelectBlock", "object", "Nothing": fails = fails + 1
    Else
        ExpectLong "Outer SELECT end", 18, outerBlock.EndSelectLine, fails
        ExpectLong "Outer SELECT case count", 2, outerBlock.CaseCount, fails
        ExpectLong "Outer SELECT case 1", 11, outerBlock.CaseLineAt(1), fails
        ExpectLong "Outer SELECT case 2", 16, outerBlock.CaseLineAt(2), fails
    End If
    ExpectLong "CASE fallthrough outer", 18, program.EndSelectOfCase(11), fails
    ExpectLong "CASE fallthrough fallback", 18, program.EndSelectOfCase(16), fails

    Dim innerBlock As clsSelectBlock
    Set innerBlock = program.SelectBlock(13)
    If innerBlock Is Nothing Then
        LogFail "Inner SelectBlock", "object", "Nothing": fails = fails + 1
    Else
        ExpectLong "Inner SELECT end", 15, innerBlock.EndSelectLine, fails
        ExpectLong "Inner SELECT case count", 1, innerBlock.CaseCount, fails
    End If

    ExpectBool "HasLabel LOGIN", True, program.HasLabel("login"), fails
    ExpectLong "LabelLine LOGIN", 19, program.LabelLine("LOGIN"), fails
    ExpectLong "ResolveJumpTarget label", 19, program.ResolveJumpTarget("LOGIN"), fails
    ExpectLong "ResolveJumpTarget sheet row", 23, program.ResolveJumpTargetSheetRow("LOGIN"), fails
    ExpectText "LabelRole LOGIN", "JumpLabel", program.LabelRoleAt(19), fails
    ExpectLong "DisplayDepth IF body", 1, program.DisplayDepthAt(2), fails
    ExpectLong "DisplayDepth ELSE corrected", 0, program.DisplayDepthAt(3), fails
    ExpectLong "StyleDepth ELSE guide", 1, program.StyleDepthAt(3), fails
    ExpectLong "ResolveJumpTarget numeric rejected", CMD_ERROR, program.ResolveJumpTarget("20"), fails
    ExpectLong "ResolveJumpTarget low out of range", CMD_ERROR, program.ResolveJumpTarget("0"), fails
    ExpectLong "ResolveJumpTarget high out of range", CMD_ERROR, program.ResolveJumpTarget("999"), fails

    Dim ifChainModelArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(ifChainModelArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(ifChainModelArr, 2, "TYPE", 0, 0, "if")
    Call MakeCmd(ifChainModelArr, 3, "ELSE_IF", "2", "EQ", "2")
    Call MakeCmd(ifChainModelArr, 4, "TYPE", 0, 0, "elseif")
    Call MakeCmd(ifChainModelArr, 5, "ELSE_IF", "3", "EQ", "3")
    Call MakeCmd(ifChainModelArr, 6, "ELSE", 0, 0, "")
    Call MakeCmd(ifChainModelArr, 7, "END_IF", 0, 0, "")
    Set program = New clsScriptProgram
    program.LoadFromArray ifChainModelArr, 1
    If Not program.ValidateStructure() Then
        LogFail "ELSE_IF ValidateStructure", "True", program.ValidationErrors(): fails = fails + 1
    Else
        LogPass "ELSE_IF ValidateStructure"
    End If
    ExpectLong "IF -> ELSE_IF", 3, program.FindBlockPair(1, "IF", 1), fails
    ExpectLong "ELSE_IF -> ELSE_IF", 5, program.FindBlockPair(3, "IF", 1), fails
    ExpectLong "ELSE_IF -> ELSE", 6, program.FindBlockPair(5, "IF", 1), fails
    ExpectLong "ELSE -> END_IF model", 7, program.FindBlockPair(6, "IF", 1), fails
    ExpectLong "EndIfOf IF", 7, program.EndIfOf(1), fails
    ExpectLong "EndIfOf ELSE_IF", 7, program.EndIfOf(3), fails
    ExpectLong "EndIfOf ELSE", 7, program.EndIfOf(6), fails
    ExpectLong "EndIfOf END_IF", 7, program.EndIfOf(7), fails
    ExpectLong "DisplayDepth ELSE_IF corrected", 0, program.DisplayDepthAt(3), fails
    ExpectLong "StyleDepth ELSE_IF guide", 1, program.StyleDepthAt(3), fails

    Dim roleArr(1 To 8, 1 To 5) As Variant
    Call MakeCmd(roleArr, 1, "CALL", 0, 0, "SUB")
    Call MakeCmd(roleArr, 2, "GOTO", 0, 0, "JMP")
    Call MakeCmd(roleArr, 3, "CALL", 0, 0, "MIX")
    Call MakeCmd(roleArr, 4, "GOTO", 0, 0, "MIX")
    Call MakeCmd(roleArr, 5, "LABEL", 0, 0, "SUB")
    Call MakeCmd(roleArr, 6, "LABEL", 0, 0, "JMP")
    Call MakeCmd(roleArr, 7, "LABEL", 0, 0, "MIX")
    Call MakeCmd(roleArr, 8, "LABEL", 0, 0, "DEAD")
    Set program = New clsScriptProgram
    program.LoadFromArray roleArr, 10
    ExpectText "LabelRole subroutine", "SubroutineLabel", program.LabelRoleAt(5), fails
    ExpectText "LabelRole jump", "JumpLabel", program.LabelRoleAt(6), fails
    ExpectText "LabelRole mixed", "MixedLabel", program.LabelRoleAt(7), fails
    ExpectText "LabelRole dead", "DeadLabel", program.LabelRoleAt(8), fails
    ExpectBool "Mixed is call target", True, program.IsCallTargetLabelAt(7), fails
    ExpectBool "Mixed warning", True, (InStr(1, program.LabelRoleWarningAt(7), "MixedLabel", vbTextCompare) > 0), fails
    ExpectLong "First CALL source SUB", 10, program.FirstCallSourceSheetRowForLabel("SUB"), fails
    ExpectLong "First CALL source MIX", 12, program.FirstCallSourceSheetRowForLabel("MIX"), fails
    ExpectLong "Resolve mixed sheet row", 16, program.ResolveJumpTargetSheetRow("MIX"), fails

    Dim subArr(1 To 5, 1 To 5) As Variant
    Call MakeCmd(subArr, 1, "CALL", 0, 0, "SUB")
    Call MakeCmd(subArr, 2, "EXIT", 0, 0, "")
    Call MakeCmd(subArr, 3, "LABEL", 0, 0, "SUB")
    Call MakeCmd(subArr, 4, "TYPE", 0, 0, "inside")
    Call MakeCmd(subArr, 5, "RETURN", 0, 0, "")
    Set program = New clsScriptProgram
    program.LoadFromArray subArr, 20
    ExpectBool "Subroutine start", True, program.IsSubroutineStartAt(3), fails
    ExpectBool "Subroutine end", True, program.IsSubroutineEndAt(5), fails
    ExpectBool "Subroutine valid label", False, program.IsFlowInvalidAt(3), fails
    ExpectLong "Subroutine body display depth", 1, program.DisplayDepthAt(4), fails

    Dim varArr(1 To 3, 1 To 5) As Variant
    Call MakeCmd(varArr, 1, "SET", "{score}", 0, "1")
    Call MakeCmd(varArr, 2, "SET", "{out}", 0, "{scroe}")
    Call MakeCmd(varArr, 3, "SET", "{unused}", 0, "value")
    Set program = New clsScriptProgram
    program.LoadFromArray varArr, 30
    ExpectBool "Program variable suggestion", True, (InStr(1, program.VariableErrorAt(2), "{score}", vbTextCompare) > 0), fails
    ExpectBool "Program unused warning", True, (InStr(1, program.VariableWarningAt(3), "{unused}", vbTextCompare) > 0), fails

    modEngine.LoadProgramFromArray arr, 5
    ExpectLong "ResolveSelectTarget first match", 12, modEngine.ResolveSelectTarget(10, "A"), fails
    ExpectLong "ResolveSelectTarget last fallback", 17, modEngine.ResolveSelectTarget(10, "Z"), fails
    ExpectLong "CASE execute fallthrough", 18, modEngine.ExecuteSmartAction(11), fails

    Dim dupArr(1 To 2, 1 To 5) As Variant
    Call MakeCmd(dupArr, 1, "LABEL", 0, 0, "DUP")
    Call MakeCmd(dupArr, 2, "LABEL", 0, 0, "dup")
    Set program = New clsScriptProgram
    program.LoadFromArray dupArr, 1
    If program.ValidateLabels() Then
        LogFail "Duplicate LABEL validation", "False", "True": fails = fails + 1
    Else
        LogPass "Duplicate LABEL validation"
    End If

    Dim tokenArr(1 To 1, 1 To 5) As Variant
    Call MakeCmd(tokenArr, 1, "LABEL", 0, 0, "{target}")
    Set program = New clsScriptProgram
    program.LoadFromArray tokenArr, 1
    If program.ValidateLabels() Then
        LogFail "Token LABEL validation", "False", "True": fails = fails + 1
    Else
        LogPass "Token LABEL validation"
    End If

    Dim numericJumpArr(1 To 1, 1 To 5) As Variant
    Call MakeCmd(numericJumpArr, 1, "GOTO", 0, 0, "1")
    Set program = New clsScriptProgram
    program.LoadFromArray numericJumpArr, 1
    If program.ValidateLabels() Then
        LogFail "Numeric GOTO validation", "False", "True": fails = fails + 1
    Else
        LogPass "Numeric GOTO validation"
    End If

    Dim dynamicJumpArr(1 To 1, 1 To 5) As Variant
    Call MakeCmd(dynamicJumpArr, 1, "CALL", 0, 0, "{target}")
    Set program = New clsScriptProgram
    program.LoadFromArray dynamicJumpArr, 1
    If program.ValidateLabels() Then
        LogFail "Dynamic CALL validation", "False", "True": fails = fails + 1
    Else
        LogPass "Dynamic CALL validation"
    End If

    Dim orphanElseIfArr(1 To 1, 1 To 5) As Variant
    Call MakeCmd(orphanElseIfArr, 1, "ELSE_IF", "1", "EQ", "1")
    Set program = New clsScriptProgram
    program.LoadFromArray orphanElseIfArr, 1
    If program.ValidateStructure() Then
        LogFail "Orphan ELSE_IF validation", "False", "True": fails = fails + 1
    Else
        LogPass "Orphan ELSE_IF validation"
    End If
    ExpectBool "Orphan ELSE_IF flow invalid", True, program.IsFlowInvalidAt(1), fails

    Dim afterElseIfArr(1 To 4, 1 To 5) As Variant
    Call MakeCmd(afterElseIfArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(afterElseIfArr, 2, "ELSE", 0, 0, "")
    Call MakeCmd(afterElseIfArr, 3, "ELSE_IF", "1", "EQ", "1")
    Call MakeCmd(afterElseIfArr, 4, "END_IF", 0, 0, "")
    Set program = New clsScriptProgram
    program.LoadFromArray afterElseIfArr, 1
    If program.ValidateStructure() Then
        LogFail "ELSE_IF after ELSE validation", "False", "True": fails = fails + 1
    Else
        LogPass "ELSE_IF after ELSE validation"
    End If
    ExpectBool "ELSE_IF after ELSE flow invalid", True, program.IsFlowInvalidAt(3), fails

    Dim badElseIfOpArr(1 To 3, 1 To 5) As Variant
    Call MakeCmd(badElseIfOpArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(badElseIfOpArr, 2, "ELSE_IF", "1", "EQUALS", "1")
    Call MakeCmd(badElseIfOpArr, 3, "END_IF", 0, 0, "")
    Set program = New clsScriptProgram
    program.LoadFromArray badElseIfOpArr, 1
    If program.ValidateStructure() Then
        LogFail "ELSE_IF operator validation", "False", "True": fails = fails + 1
    Else
        LogPass "ELSE_IF operator validation"
    End If
    ExpectBool "ELSE_IF operator flow invalid", True, program.IsFlowInvalidAt(2), fails

    Dim missingElseIfParamArr(1 To 3, 1 To 5) As Variant
    Call MakeCmd(missingElseIfParamArr, 1, "IF", "1", "EQ", "0")
    Call MakeCmd(missingElseIfParamArr, 2, "ELSE_IF", "", "", "")
    Call MakeCmd(missingElseIfParamArr, 3, "END_IF", 0, 0, "")
    Set program = New clsScriptProgram
    program.LoadFromArray missingElseIfParamArr, 1
    If program.ValidateStructure() Then
        LogFail "ELSE_IF missing params validation", "False", "True": fails = fails + 1
    Else
        LogPass "ELSE_IF missing params validation"
    End If
    ExpectBool "ELSE_IF missing params flow invalid", True, program.IsFlowInvalidAt(2), fails

    TC_ScriptProgramModel = fails
End Function

Private Sub ExpectLong(ByVal testName As String, ByVal expected As Long, ByVal actual As Long, ByRef failCount As Long)
    If expected = actual Then
        LogPass testName
    Else
        LogFail testName, CStr(expected), CStr(actual)
        failCount = failCount + 1
    End If
End Sub

Private Sub ExpectText(ByVal testName As String, ByVal expected As String, ByVal actual As String, ByRef failCount As Long)
    If expected = actual Then
        LogPass testName
    Else
        LogFail testName, expected, actual
        failCount = failCount + 1
    End If
End Sub

Private Sub ExpectBool(ByVal testName As String, ByVal expected As Boolean, ByVal actual As Boolean, ByRef failCount As Long)
    If expected = actual Then
        LogPass testName
    Else
        LogFail testName, CStr(expected), CStr(actual)
        failCount = failCount + 1
    End If
End Sub

Private Sub ExpectVar(ByVal name As String, ByVal expected As String, ByRef failCount As Long)
    ExpectText "Var " & name, expected, modContext.GetVar(name), failCount
End Sub

Private Sub ExpectStackDepth(ByVal expected As Long, ByRef failCount As Long)
    ExpectLong "StackDepth", expected, modContext.GetStackDepth(), failCount
End Sub

Private Sub ExpectLoopDepth(ByVal expected As Long, ByRef failCount As Long)
    ExpectLong "LoopDepth", expected, modContext.LoopDepth(), failCount
End Sub

Private Sub ExpectIfBranchDepth(ByVal expected As Long, ByRef failCount As Long)
    ExpectLong "IfBranchDepth", expected, modContext.IfBranchDepth(), failCount
End Sub

Private Sub ExpectTerminal(ByVal expected As eRunResult, ByVal actual As eRunResult, ByRef failCount As Long)
    If expected = actual Then
        LogPass "Terminal " & RunResultName(expected)
    Else
        LogFail "Terminal", RunResultName(expected), RunResultName(actual)
        failCount = failCount + 1
    End If
End Sub

Private Sub ExpectVisits(ByRef trace() As Long, ByVal lineNo As Long, ByVal expected As Long, ByRef failCount As Long)
    Dim actual As Long
    Dim i As Long

    On Error GoTo EmptyTrace
    For i = LBound(trace) To UBound(trace)
        If trace(i) = lineNo Then actual = actual + 1
    Next i
    On Error GoTo 0

    ExpectLong "Visits line " & lineNo, expected, actual, failCount
    Exit Sub

EmptyTrace:
    On Error GoTo 0
    ExpectLong "Visits line " & lineNo, expected, 0, failCount
End Sub

Private Sub ExpectTrace(ByRef trace() As Long, ByVal expected As Variant, ByRef failCount As Long)
    Dim actualLen As Long
    Dim expectedLen As Long
    Dim i As Long

    actualLen = TraceLength(trace)
    expectedLen = UBound(expected) - LBound(expected) + 1
    If actualLen <> expectedLen Then
        LogFail "Trace length", CStr(expectedLen), CStr(actualLen)
        failCount = failCount + 1
        Exit Sub
    End If

    For i = 1 To actualLen
        If trace(i) <> CLng(expected(LBound(expected) + i - 1)) Then
            LogFail "Trace item " & i, CStr(expected(LBound(expected) + i - 1)), CStr(trace(i))
            failCount = failCount + 1
            Exit Sub
        End If
    Next i

    LogPass "Trace exact"
End Sub

Private Function TraceLength(ByRef trace() As Long) As Long
    On Error GoTo EmptyTrace
    TraceLength = UBound(trace) - LBound(trace) + 1
    Exit Function

EmptyTrace:
    TraceLength = 0
End Function

Private Function RunResultName(ByVal result As eRunResult) As String
    Select Case result
        Case rrCompleted: RunResultName = "rrCompleted"
        Case rrHalted: RunResultName = "rrHalted"
        Case rrErrored: RunResultName = "rrErrored"
        Case rrExhausted: RunResultName = "rrExhausted"
        Case Else: RunResultName = "Unknown(" & CStr(result) & ")"
    End Select
End Function

Private Sub LogPass(testName As String)
    Debug.Print "[PASS] " & testName
End Sub

Private Sub LogFail(testName As String, expected As String, actual As String)
    Debug.Print "[FAIL] " & testName & " (Exp: " & expected & " / Act: " & actual & ")"
    ' jsonl에도 ok:false A레코드 기록 (vbaflow test가 실패 상세를 인라인하도록)
    On Error Resume Next
    LogShim.AssertEq testName, expected, actual
    On Error GoTo 0
End Sub
