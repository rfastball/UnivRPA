Attribute VB_Name = "modTester"
' @desc 명령어셋 통합 자가진단 및 레지스트리 역검증 실행 담당
' @depends modEngine, modContext, modActions, modWinAPI, modCmdRegistry
Option Explicit

'==============================================================================
' [System] Full Integration Test Runner (v2 명령어셋)
'==============================================================================
' 주의: 단일 라인 하니스(ExecuteSmartAction(arrCmd,1))는 m_arrMain 을 채우지 못하므로
'   블록 흐름(IF 거짓/WHILE/REPEAT/SELECT, BREAK/CONTINUE)은 여기서 검증하지 않는다.
'   -> 해당 항목은 워크시트 프로그램으로 수동 검증 (spec Verification 참조).
'   여기서는 m_arrMain 불필요한 것만: SET/메모리, 토크나이저, TYPE/KEY, CALL/RETURN,
'   비전(WAIT_VANISH/ASSERT), FOCUS, IF-참, 그리고 연산자 단위검증을 다룬다.
'==============================================================================
Sub Run_System_SelfDiagnosis()
    Dim wsTest As Worksheet
    Dim failCount As Long
    Dim arrCmd(1 To 10, 1 To 5) As Variant
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
    modEngine.ExecuteSmartAction arrCmd, 1

    If modContext.GetVar("score") = "10" Then
        LogPass "SET (mem write)"
    Else
        LogFail "SET", "10", modContext.GetVar("score"): failCount = failCount + 1
    End If

    Call MakeCmd(arrCmd, 1, "SET", "{score}", 0, "{$BLANK}")
    modEngine.ExecuteSmartAction arrCmd, 1

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
    modEngine.ExecuteSmartAction arrCmd, 1
    Call MakeCmd(arrCmd, 1, "SET", "{b}", 0, "{a}_world")
    modEngine.ExecuteSmartAction arrCmd, 1

    If modContext.GetVar("b") = "hello_world" Then
        LogPass "Interpolation ({a}_world)"
    Else
        LogFail "Interpolation", "hello_world", modContext.GetVar("b"): failCount = failCount + 1
    End If

    Call MakeCmd(arrCmd, 1, "SET", "{t}", 0, "A{$TAB}B")
    modEngine.ExecuteSmartAction arrCmd, 1

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
    modEngine.ExecuteSmartAction arrCmd, 1

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
    modEngine.ExecuteSmartAction arrCmd, 1
    modActions.BetterWait 80
    Call MakeCmd(arrCmd, 1, "SET", "{c}", 0, "{$CLIP}")
    modEngine.ExecuteSmartAction arrCmd, 1

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
    modEngine.ExecuteSmartAction arrCmd, 1

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
    modEngine.ExecuteSmartAction arrCmd, 1
    Call MakeCmd(arrCmd, 1, "KEY", 13, 0, "")
    modEngine.ExecuteSmartAction arrCmd, 1

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
    modEngine.ExecuteSmartAction arrCmd, 1

    wsTest.Range("A2").Select
    Call MakeCmd(arrCmd, 1, "TYPE", 10, 0, "{dyn}")
    modEngine.ExecuteSmartAction arrCmd, 1
    Call MakeCmd(arrCmd, 1, "KEY", 13, 0, "")
    modEngine.ExecuteSmartAction arrCmd, 1

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
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)

    If nextLine = 50 And modContext.GetStackDepth() = 1 Then
        LogPass "CALL (Stack Push & Jump)"

        Call MakeCmd(arrCmd, 1, "RETURN", 0, 0, "")
        nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)

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
    '   IF {$PIXEL:x,y} == <현재색>  -> 참 -> 0 반환 (m_arrMain 불필요)
    '==========================================================================
    Debug.Print "--- [Test 8] IF (true branch) ---"

    Call MakeCmd(arrCmd, 1, "IF", "{$PIXEL:" & screenX & "," & screenY & "}", "==", actualColor)
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)

    If nextLine = 0 Then
        LogPass "IF == (true -> next line)"
    Else
        LogFail "IF ==", "0 (next)", "Line=" & nextLine: failCount = failCount + 1
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
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
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
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)

    If nextLine <> -1 Then
        LogPass "FOCUS (Find Window)"
    Else
        LogFail "FOCUS", "Success", "Failed (-1)": failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 11] 조건 연산자 단위 검증 (Action_CheckCondition 직접 호출)
    '   == 정확비교 / LIKE 패턴 / STARTS / ENDS / CONTAINS / 맨 '=' 금지
    '==========================================================================
    Debug.Print "--- [Test 11] Condition operators ---"

    AssertCond "num >", "5", ">", "3", True, failCount
    AssertCond "== exact", "abc", "==", "abc", True, failCount
    AssertCond "== star literal", "abc", "==", "ab*", False, failCount   ' * 는 리터럴
    AssertCond "LIKE pattern", "abc", "LIKE", "a*", True, failCount
    AssertCond "STARTS", "abcdef", "STARTS", "abc", True, failCount
    AssertCond "ENDS", "abcdef", "ENDS", "def", True, failCount
    AssertCond "CONTAINS", "abcdef", "CONTAINS", "cd", True, failCount
    AssertCond "bare = banned", "5", "=", "5", False, failCount          ' 맨 '=' 금지 -> False

    '==========================================================================
    ' [Test 12] INC / DEC (정수 read-modify-write, {$LOOP} 비대칭 닫힘)
    '==========================================================================
    Debug.Print "--- [Test 12] INC / DEC ---"

    ' 미정의 변수 INC -> 1 (missing==empty -> 0 +1)
    Call MakeCmd(arrCmd, 1, "INC", "{n}", 0, "")
    modEngine.ExecuteSmartAction arrCmd, 1
    If modContext.GetVar("n") = "1" Then
        LogPass "INC (undefined -> 1)"
    Else
        LogFail "INC undef", "1", modContext.GetVar("n"): failCount = failCount + 1
    End If

    ' INC {n} 5 -> 6 (P3 RESOLVE)
    Call MakeCmd(arrCmd, 1, "INC", "{n}", 0, "5")
    modEngine.ExecuteSmartAction arrCmd, 1
    If modContext.GetVar("n") = "6" Then
        LogPass "INC (+5 -> 6)"
    Else
        LogFail "INC +5", "6", modContext.GetVar("n"): failCount = failCount + 1
    End If

    ' DEC {n} -> 5 (기본 1 감분)
    Call MakeCmd(arrCmd, 1, "DEC", "{n}", 0, "")
    modEngine.ExecuteSmartAction arrCmd, 1
    If modContext.GetVar("n") = "5" Then
        LogPass "DEC (-1 -> 5)"
    Else
        LogFail "DEC", "5", modContext.GetVar("n"): failCount = failCount + 1
    End If

    ' 읽기전용 좌변 INC {$LOOP} -> CMD_ERROR(-1)
    Dim incErr As Long
    Call MakeCmd(arrCmd, 1, "INC", "{$LOOP}", 0, "")
    incErr = modEngine.ExecuteSmartAction(arrCmd, 1)
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

Private Sub LogPass(testName As String)
    Debug.Print "[PASS] " & testName
End Sub

Private Sub LogFail(testName As String, expected As String, actual As String)
    Debug.Print "[FAIL] " & testName & " (Exp: " & expected & " / Act: " & actual & ")"
End Sub
