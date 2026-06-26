Attribute VB_Name = "modTester"
' @desc 명령어셋 통합 자가진단 및 레지스트리 역검증 실행 담당
' @depends modEngine, modContext, modActions, modWinAPI, modCmdRegistry
Option Explicit

'==============================================================================
' [System] Full Integration Test Runner (v2 명령어셋)
'==============================================================================
' 주의: 단일 라인 하니스는 배열을 clsScriptProgram으로 로드한 뒤 ExecuteSmartAction(1)을 호출한다.
'   여기서는 액션/토큰/메모리와 ProgramImage 구조 인덱스를 함께 검증한다.
'   워크시트 UI 동작은 별도 수동 검증 (spec Verification 참조).
'   여기서는 SET/메모리, 토크나이저, TYPE/KEY, CALL/RETURN,
'   비전(WAIT_VANISH/ASSERT), FOCUS, IF-참, 그리고 연산자 단위검증을 다룬다.
'==============================================================================
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
    ' [Test 11] 조건 연산자 단위 검증 (Action_CheckCondition 직접 호출)
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
    actual = modActions.Action_CheckCondition(v1, op, v2)
    If actual = expected Then
        LogPass "Cond " & name
    Else
        LogFail "Cond " & name, CStr(expected), CStr(actual)
        failCount = failCount + 1
    End If
End Sub

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

Private Function Test_ScriptProgramModel() As Long
    Dim fails As Long
    Dim arr(1 To 24, 1 To 5) As Variant

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

    ExpectLong "TotalLines", 24, program.TotalLines, fails
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

Private Sub LogPass(testName As String)
    Debug.Print "[PASS] " & testName
End Sub

Private Sub LogFail(testName As String, expected As String, actual As String)
    Debug.Print "[FAIL] " & testName & " (Exp: " & expected & " / Act: " & actual & ")"
End Sub
