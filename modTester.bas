Attribute VB_Name = "modTester"
' @desc 명령어셋 통합 자가진단 및 레지스트리 역검증 실행 담당
' @depends modEngine, modContext, modActions, modWinAPI, modCmdRegistry
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
    failCount = failCount + Test_RunProgramSmoke()
    failCount = failCount + Test_FlowHappyPaths()
    failCount = failCount + Test_TableSnapshotReadOnly()
    failCount = failCount + Test_ForeachRowExecution()
    failCount = failCount + Test_FlowNestedInteractions()
    failCount = failCount + Test_FlowNegativePaths()
    failCount = failCount + Test_MemoryResolverHeadless()
    failCount = failCount + Test_ParamSpecValidation()

    Debug.Print "--- [Regression] Registry/ScriptProgram ---"
    failCount = failCount + Test_RegistryDispatchSync()
    failCount = failCount + Test_ScriptProgramModel()

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

    Call MakeCmd(arrCmd, 1, "CALL", 0, 0, "50")
    nextLine = ExecuteArrayLine(arrCmd, 1)

    If nextLine = 50 And modContext.GetStackDepth() = 1 Then
        LogPass "CALL (Stack Push & Jump)"

        Call MakeCmd(arrCmd, 1, "RETURN", 0, 0, "")
        nextLine = ExecuteArrayLine(arrCmd, 1)

        If nextLine = 2 And modContext.GetStackDepth() = 0 Then
            LogPass "RETURN (Stack Pop & Return)"
        Else
            LogFail "RETURN", "Return 2", "Line=" & nextLine: failCount = failCount + 1
        End If
    Else
        LogFail "CALL", "Jump 50", "Line=" & nextLine: failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 8] IF (참 갈래는 단일 라인으로 검증: 다음 줄 반환)
    '   IF {$PIXEL:x,y} EQ <현재색>  -> 참 -> 0 반환
    '==========================================================================
    Debug.Print "--- [Test 8] IF (true branch) ---"

    Call MakeCmd(arrCmd, 1, "IF", "{$PIXEL:" & screenX & "," & screenY & "}", "EQ", actualColor)
    nextLine = ExecuteArrayLine(arrCmd, 1)

    If nextLine = 0 Then
        LogPass "IF EQ (true -> next line)"
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

    ' 미정의 변수 INC -> 1 (missing empty -> 0 +1)
    Call MakeCmd(arrCmd, 1, "INC", "{n}", 0, "")
    ExecuteArrayLine arrCmd, 1
    If modContext.GetVar("n") = "1" Then
        LogPass "INC (undefined -> 1)"
    Else
        LogFail "INC undef", "1", modContext.GetVar("n"): failCount = failCount + 1
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

    '==========================================================================
    ' [Test 13] 역검증 ? 레지스트리<->핸들러 동기화 파수꾼 (설계 §7)
    '==========================================================================
    Debug.Print "--- [Test 13] Registry/Handler sync ---"
    Dim syncFails As Long
    syncFails = Test_RegistryDispatchSync()
    If syncFails = 0 Then
        LogPass "Registry dispatch sync (all routed)"
    Else
        LogFail "Registry sync", "0 fails", CStr(syncFails) & " fails": failCount = failCount + syncFails
    End If

    '==========================================================================
    ' [Test 14] ScriptProgram 구조/SELECT/LABEL 인덱스
    '==========================================================================
    Debug.Print "--- [Test 14] ScriptProgram model ---"
    failCount = failCount + Test_ScriptProgramModel()

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
Public Function Test_RegistryDispatchSync() As Long
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
    Test_RegistryDispatchSync = fails
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

Private Function Test_RunProgramSmoke() As Long
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

    Test_RunProgramSmoke = fails
End Function

Private Function Test_FlowHappyPaths() As Long
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

    Dim repeatArr(1 To 4, 1 To 5) As Variant
    Call MakeCmd(repeatArr, 1, "REPEAT", "3", 0, "")
    Call MakeCmd(repeatArr, 2, "INC", "{r}", 0, "")
    Call MakeCmd(repeatArr, 3, "END_REPEAT", 0, 0, "")
    Call MakeCmd(repeatArr, 4, "EXIT", 0, 0, "")
    result = RunProgram(repeatArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVisits trace, 2, 3, fails
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

    Test_FlowHappyPaths = fails
End Function

Private Function Test_TableSnapshotReadOnly() As Long
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
    ExpectBool "Snapshot Headers CODE", True, tbl.Headers.exists("CODE"), fails
    ExpectText "Snapshot GetValue", "Alice", CStr(tbl.GetValue(2, "Name")), fails

    rec = tbl.GetRecord(3)
    ExpectText "Snapshot GetRecord", "Bob", CStr(rec(2)), fails

    Set rows = tbl.SelectRows("Group", "A")
    ExpectLong "Snapshot SelectRows count", 2, rows.Count, fails
    ExpectLong "Snapshot SelectRows first", 2, CLng(rows(1)), fails
    ExpectLong "Snapshot SelectRows second", 4, CLng(rows(2)), fails

    Set rows = tbl.SelectRowsByIndex(1, "002")
    ExpectLong "Snapshot SelectRowsByIndex count", 1, rows.Count, fails
    ExpectLong "Snapshot SelectRowsByIndex row", 3, CLng(rows(1)), fails

    keys = tbl.GetUniqueKeys("Group")
    ExpectLong "Snapshot GetUniqueKeys count", 2, UBound(keys) - LBound(keys) + 1, fails

    keys = tbl.GetUniqueKeysByIndex(1)
    ExpectLong "Snapshot GetUniqueKeysByIndex count", 3, UBound(keys) - LBound(keys) + 1, fails

    tbl.Clear
    ExpectLong "Snapshot Clear RowCount", 0, tbl.RowCount, fails
    ExpectLong "Snapshot Clear ColCount", 0, tbl.ColCount, fails

    Test_TableSnapshotReadOnly = fails
    Exit Function

ErrHandler:
    fails = fails + 1
    LogFail "TableSnapshot read-only test", "no error", Err.Description
    Test_TableSnapshotReadOnly = fails
End Function

Private Function Test_ForeachRowExecution() As Long
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

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
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

    Dim arr(1 To 21, 1 To 5) As Variant
    Call MakeCmd(arr, 1, "FOREACH_ROW", ThisWorkbook.FullName, sheetName, "")
    Call MakeCmd(arr, 2, "IF", "{$RIDX}", "EQ", "1")
    Call MakeCmd(arr, 3, "SET", "{first_code}", 0, "{$ROW.Code}")
    Call MakeCmd(arr, 4, "SET", "{first_rounded}", 0, "{$ROW.Rounded}")
    Call MakeCmd(arr, 5, "SET", "{first_date}", 0, "{$ROW.DateNarrow}")
    Call MakeCmd(arr, 6, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 7, "SET", "{last_code}", 0, "{$ROW.Code}")
    Call MakeCmd(arr, 8, "SET", "{last_idx}", 0, "{$RIDX}")
    Call MakeCmd(arr, 9, "IF", "{$ROW.Code}", "EQ", "B")
    Call MakeCmd(arr, 10, "CONTINUE", 0, 0, "")
    Call MakeCmd(arr, 11, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 12, "INC", "{processed}", 0, "")
    Call MakeCmd(arr, 13, "REPEAT", "2", 0, "")
    Call MakeCmd(arr, 14, "SET", "{nested_ridx}", 0, "{$RIDX}")
    Call MakeCmd(arr, 15, "END_REPEAT", 0, 0, "")
    Call MakeCmd(arr, 16, "IF", "{$ROW.Code}", "EQ", "C")
    Call MakeCmd(arr, 17, "BREAK", 0, 0, "")
    Call MakeCmd(arr, 18, "END_IF", 0, 0, "")
    Call MakeCmd(arr, 19, "END_FOREACH", 0, 0, "")
    Call MakeCmd(arr, 20, "SET", "{after_loop}", 0, "done")
    Call MakeCmd(arr, 21, "EXIT", 0, 0, "")

    result = RunProgram(arr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "first_code", "007", fails
    ExpectVar "first_rounded", "11", fails
    ExpectVar "first_date", "2024-02-01", fails
    ExpectVar "last_code", "C", fails
    ExpectVar "last_idx", "3", fails
    ExpectVar "processed", "2", fails
    ExpectVar "nested_ridx", "3", fails
    ExpectVar "after_loop", "done", fails
    ExpectLoopDepth 0, fails

    ws.Range("A1:D4").AutoFilter Field:=1, Criteria1:="<>B"
    Dim visibleArr(1 To 10, 1 To 5) As Variant
    Call MakeCmd(visibleArr, 1, "FOREACH_ROW", ThisWorkbook.FullName, sheetName, "VISIBLE_ROWS")
    Call MakeCmd(visibleArr, 2, "SET", "{visible_last_idx}", 0, "{$RIDX}")
    Call MakeCmd(visibleArr, 3, "IF", "{$RIDX}", "EQ", "1")
    Call MakeCmd(visibleArr, 4, "SET", "{visible1}", 0, "{$ROW.Code}")
    Call MakeCmd(visibleArr, 5, "END_IF", 0, 0, "")
    Call MakeCmd(visibleArr, 6, "IF", "{$RIDX}", "EQ", "2")
    Call MakeCmd(visibleArr, 7, "SET", "{visible2}", 0, "{$ROW.Code}")
    Call MakeCmd(visibleArr, 8, "END_IF", 0, 0, "")
    Call MakeCmd(visibleArr, 9, "END_FOREACH", 0, 0, "")
    Call MakeCmd(visibleArr, 10, "EXIT", 0, 0, "")
    result = RunProgram(visibleArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "visible_last_idx", "2", fails
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
    Test_ForeachRowExecution = fails
    Exit Function

ErrHandler:
    fails = fails + 1
    LogFail "FOREACH_ROW test", "no error", Err.Description
    Resume CleanUp
End Function

Private Function Test_FlowNestedInteractions() As Long
    Dim fails As Long
    Dim trace() As Long
    Dim lastCode As Long
    Dim lastLine As Long
    Dim result As eRunResult

    Debug.Print "--- [Flow 3] Nested and interaction paths ---"

    Dim breakArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(breakArr, 1, "REPEAT", "5", 0, "")
    Call MakeCmd(breakArr, 2, "INC", "{before_break}", 0, "")
    Call MakeCmd(breakArr, 3, "BREAK", 0, 0, "")
    Call MakeCmd(breakArr, 4, "INC", "{after_break}", 0, "")
    Call MakeCmd(breakArr, 5, "END_REPEAT", 0, 0, "")
    Call MakeCmd(breakArr, 6, "EXIT", 0, 0, "")
    result = RunProgram(breakArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "before_break", "1", fails
    ExpectVar "after_break", "", fails
    ExpectVisits trace, 4, 0, fails
    ExpectLoopDepth 0, fails

    Dim continueArr(1 To 6, 1 To 5) As Variant
    Call MakeCmd(continueArr, 1, "REPEAT", "3", 0, "")
    Call MakeCmd(continueArr, 2, "INC", "{continue_hits}", 0, "")
    Call MakeCmd(continueArr, 3, "CONTINUE", 0, 0, "")
    Call MakeCmd(continueArr, 4, "INC", "{continue_skipped}", 0, "")
    Call MakeCmd(continueArr, 5, "END_REPEAT", 0, 0, "")
    Call MakeCmd(continueArr, 6, "EXIT", 0, 0, "")
    result = RunProgram(continueArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "continue_hits", "3", fails
    ExpectVar "continue_skipped", "", fails
    ExpectVisits trace, 4, 0, fails
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

    Dim nestedArr(1 To 12, 1 To 5) As Variant
    Call MakeCmd(nestedArr, 1, "REPEAT", "2", 0, "")
    Call MakeCmd(nestedArr, 2, "SELECT", "{$LOOP}", 0, "")
    Call MakeCmd(nestedArr, 3, "CASE", "1", 0, "")
    Call MakeCmd(nestedArr, 4, "INC", "{outer1}", 0, "")
    Call MakeCmd(nestedArr, 5, "CASE", "2", 0, "")
    Call MakeCmd(nestedArr, 6, "INC", "{outer2}", 0, "")
    Call MakeCmd(nestedArr, 7, "END_SELECT", 0, 0, "")
    Call MakeCmd(nestedArr, 8, "REPEAT", "2", 0, "")
    Call MakeCmd(nestedArr, 9, "INC", "{inner_hits}", 0, "")
    Call MakeCmd(nestedArr, 10, "END_REPEAT", 0, 0, "")
    Call MakeCmd(nestedArr, 11, "END_REPEAT", 0, 0, "")
    Call MakeCmd(nestedArr, 12, "EXIT", 0, 0, "")
    result = RunProgram(nestedArr, 1, 0, trace, lastCode, lastLine)
    ExpectTerminal rrHalted, result, fails
    ExpectVar "outer1", "1", fails
    ExpectVar "outer2", "1", fails
    ExpectVar "inner_hits", "4", fails
    ExpectLoopDepth 0, fails

    Dim callLoopArr(1 To 7, 1 To 5) As Variant
    Call MakeCmd(callLoopArr, 1, "REPEAT", "2", 0, "")
    Call MakeCmd(callLoopArr, 2, "CALL", 0, 0, "SUB")
    Call MakeCmd(callLoopArr, 3, "END_REPEAT", 0, 0, "")
    Call MakeCmd(callLoopArr, 4, "EXIT", 0, 0, "")
    Call MakeCmd(callLoopArr, 5, "LABEL", 0, 0, "SUB")
    Call MakeCmd(callLoopArr, 6, "INC", "{call_loop_hits}", 0, "")
    Call MakeCmd(callLoopArr, 7, "RETURN", 0, 0, "")
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

    Test_FlowNestedInteractions = fails
End Function

Private Function Test_FlowNegativePaths() As Long
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

    Test_FlowNegativePaths = fails
End Function

Private Function Test_MemoryResolverHeadless() As Long
    Dim fails As Long
    Dim arr(1 To 17, 1 To 5) As Variant
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
    Call MakeCmd(arr, 5, "SET", "{cr}", 0, "A{$CR}B")
    Call MakeCmd(arr, 6, "SET", "{lf}", 0, "A{$LF}B")
    Call MakeCmd(arr, 7, "SET", "{sp}", 0, "A{$SPACE}B")
    Call MakeCmd(arr, 8, "SET", "{unknown}", 0, "{$XXX}")
    Call MakeCmd(arr, 9, "SET", "{open}", 0, "A{broken")
    Call MakeCmd(arr, 10, "SET", "{x}", 0, "300")
    Call MakeCmd(arr, 11, "SET", "{y}", 0, "300")
    Call MakeCmd(arr, 12, "SET", "{px}", 0, "{$PIXEL:{x},{y}}")
    Call MakeCmd(arr, 13, "SET", "{delta}", 0, "-2")
    Call MakeCmd(arr, 14, "INC", "{n}", 0, "")
    Call MakeCmd(arr, 15, "INC", "{n}", 0, "{delta}")
    Call MakeCmd(arr, 16, "DEC", "{n}", 0, "-3")
    Call MakeCmd(arr, 17, "EXIT", 0, 0, "")

    result = RunProgram(arr, 1, 0, trace, lastCode, lastLine)
    pxExpected = modWinAPI.GetColorHexAt(300, 300)

    ExpectTerminal rrHalted, result, fails
    ExpectVar "a", "", fails
    ExpectVar "b", "hello_world", fails
    ExpectVar "cr", "A" & vbCr & "B", fails
    ExpectVar "lf", "A" & vbLf & "B", fails
    ExpectVar "sp", "A B", fails
    ExpectVar "unknown", "{$XXX}", fails
    ExpectVar "open", "A{broken", fails
    ExpectVar "px", pxExpected, fails
    ExpectVar "n", "2", fails
    ExpectTrace trace, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17), fails

    Test_MemoryResolverHeadless = fails
End Function

Private Function Test_ParamSpecValidation() As Long
    Dim fails As Long

    Debug.Print "--- [Flow 6] ParamSpec validation ---"

    ExpectBool "Validate SLEEP P3", True, modDevTools.ValidateLine("SLEEP", "", "", "1000"), fails
    ExpectBool "Reject SLEEP shifted P1", False, modDevTools.ValidateLine("SLEEP", "1000", "", ""), fails
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

    Test_ParamSpecValidation = fails
End Function

' [Helper] 가상 명령어 배열 생성기 (B=cmd, C=P1, D=P2, E=P3)
Private Sub MakeCmd(ByRef arr As Variant, R As Long, cmd As String, p1 As Variant, p2 As Variant, p3 As String)
    arr(R, 2) = cmd
    arr(R, 3) = p1
    arr(R, 4) = p2
    arr(R, 5) = p3
End Sub

Private Function ExecuteArrayLine(ByRef arr As Variant, ByVal R As Long) As Long
    modEngine.LoadProgramFromArray arr, 1
    ExecuteArrayLine = modEngine.ExecuteSmartAction(R)
End Function

Private Sub AppendTrace(ByRef trace() As Long, ByVal traceCount As Long, ByVal lineNo As Long)
    ReDim Preserve trace(1 To traceCount)
    trace(traceCount) = lineNo
End Sub

Private Function Test_ScriptProgramModel() As Long
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
    ExpectLong "ResolveJumpTarget numeric", 20, program.ResolveJumpTarget("20"), fails
    ExpectLong "ResolveJumpTarget low out of range", CMD_ERROR, program.ResolveJumpTarget("0"), fails
    ExpectLong "ResolveJumpTarget high out of range", CMD_ERROR, program.ResolveJumpTarget("999"), fails

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

    Test_ScriptProgramModel = fails
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
End Sub
