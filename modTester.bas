Attribute VB_Name = "modTester"
Option Explicit

'==============================================================================
' [System] Full Integration Test Runner (Enhanced)
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
    wsTest.Name = "RPA_TEST_SUITE"
    
    ' 컨텍스트 및 변수 초기화
    modContext.InitContext
    failCount = 0
    
    ' 엑셀 창 활성화 (입력 테스트 필수)
    AppActivate Application.Caption
    modActions.BetterWait 500
    
    '==========================================================================
    ' [Test 1] 메모리 및 연산 검증 (SET_MEM, INC)
    '==========================================================================
    Debug.Print "--- [Test 1] Memory Logic ---"
    
    Call MakeCmd(arrCmd, 1, "SET_MEM", 0, 0, "10")
    modEngine.ExecuteSmartAction arrCmd, 1
    
    Call MakeCmd(arrCmd, 1, "INC", 0, 0, "")
    modEngine.ExecuteSmartAction arrCmd, 1
    
    If modContext.GetVar(0) = "11" Then
        LogPass "Memory & Math (10+1=11)"
    Else
        LogFail "Memory & Math", "11", modContext.GetVar(0)
        failCount = failCount + 1
    End If
    
    '==========================================================================
    ' [Test 2] 타이핑 검증 (TYPE, KEY)
    '==========================================================================
    Debug.Print "--- [Test 2] Input (TYPE & Enter) ---"
    
    wsTest.Range("A1").ClearContents
    wsTest.Range("A1").Select
    
    ' 타이핑 수행
    Call MakeCmd(arrCmd, 1, "TYPE", 10, 0, "HelloRPA") ' 딜레이 10ms (고속)
    modEngine.ExecuteSmartAction arrCmd, 1
    
    ' 엔터키 입력 (입력 확정)
    Call MakeCmd(arrCmd, 1, "KEY", 13, 0, "")
    modEngine.ExecuteSmartAction arrCmd, 1
    
    DoEvents
    If wsTest.Range("A1").Value = "HelloRPA" Then
        LogPass "TYPE Action (Cell Input)"
    Else
        LogFail "TYPE Action", "HelloRPA", wsTest.Range("A1").Value
        failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 3] 화면 인식 로직 검증 (IF_PIXEL)
    '==========================================================================
    Debug.Print "--- [Test 3] Vision Logic (IF_PIXEL) ---"
    
    ' 테스트용 좌표 설정 (화면 중앙 쯤 임의 설정)
    Dim screenX As Long: screenX = 300
    Dim screenY As Long: screenY = 300
    
    ' 현재 해당 좌표의 색상을 가져옴 (Self-Reference)
    Dim actualColor As String
    actualColor = modWinAPI.GetColorHexAt(screenX, screenY)
    
    ' 로직: 색상이 맞으면 99번 줄로 가라
    Call MakeCmd(arrCmd, 1, "IF_PIXEL", screenX, screenY, actualColor & "|99")
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
    
    If nextLine = 99 Then
        LogPass "IF_PIXEL (Condition Match -> Jump)"
    Else
        LogFail "IF_PIXEL", "Jump 99", "Line=" & nextLine
        failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 4] 스택 및 흐름 제어 (CALL/RETURN)
    '==========================================================================
    Debug.Print "--- [Test 4] Flow Control (CALL/RETURN) ---"
    
    Call MakeCmd(arrCmd, 1, "CALL", 0, 0, "50")
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
    
    If nextLine = 50 And modContext.GetStackDepth() = 1 Then
        LogPass "CALL (Stack Push & Jump)"
        
        Call MakeCmd(arrCmd, 1, "RETURN", 0, 0, "")
        nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
        
        ' 1번줄에서 호출했으니 복귀주소는 r+1인 2번줄이어야 함
        If nextLine = 2 And modContext.GetStackDepth() = 0 Then
            LogPass "RETURN (Stack Pop & Return)"
        Else
             LogFail "RETURN", "Return 2", "Line=" & nextLine
             failCount = failCount + 1
        End If
    Else
        LogFail "CALL", "Jump 50", "Line=" & nextLine
        failCount = failCount + 1
    End If
    
    '==========================================================================
    ' [Test 5] 클립보드 분기 검증 (IF_CLIP) - *NEW*
    '==========================================================================
    Debug.Print "--- [Test 5] Clipboard Logic ---"
    
    ' 1. 클립보드 강제 설정 (API 사용)
    modWinAPI.SetClipboardText "SECRET_CODE_1234"
    modActions.BetterWait 100
    
    ' 2. IF_CLIP: "CODE"가 포함되어 있으면 77번으로 점프
    Call MakeCmd(arrCmd, 1, "IF_CLIP", 0, 0, "CODE|77")
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
    
    If nextLine = 77 Then
        LogPass "IF_CLIP (Found Text -> Jump)"
    Else
        LogFail "IF_CLIP", "Jump 77", "Line=" & nextLine
        failCount = failCount + 1
    End If
    
    ' 3. IFNOT_CLIP: "BANANA"가 없으면 88번으로 점프
    Call MakeCmd(arrCmd, 1, "IFNOT_CLIP", 0, 0, "BANANA|88")
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
    
    If nextLine = 88 Then
        LogPass "IFNOT_CLIP (Not Found -> Jump)"
    Else
        LogFail "IFNOT_CLIP", "Jump 88", "Line=" & nextLine
        failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 6] 다중 색상 분기 (SWITCH_PIXEL) - *NEW*
    '==========================================================================
    Debug.Print "--- [Test 6] Switch Logic ---"
    
    ' 아까 구한 actualColor를 재사용
    ' 시나리오: "검정색|10 ; 현재색|20 ; 흰색|30" 중 어디로 가는가?
    Dim switchCond As String
    switchCond = "000000|10;" & actualColor & "|20;FFFFFF|30"
    
    Call MakeCmd(arrCmd, 1, "SWITCH_PIXEL", screenX, screenY, switchCond)
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
    
    If nextLine = 20 Then
        LogPass "SWITCH_PIXEL (Multi-Condition Match)"
    Else
        LogFail "SWITCH_PIXEL", "Jump 20", "Line=" & nextLine
        failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 7] 메모리 변수 타이핑 (TYPE_MEM) - *NEW*
    '==========================================================================
    Debug.Print "--- [Test 7] Memory Typing ---"
    
    ' 1. 메모리 99번에 값 저장
    modContext.SetVar 99, "DynamicValue"
    
    ' 2. A2 셀 선택
    wsTest.Range("A2").Select
    
    ' 3. TYPE_MEM 실행 (p1:인덱스99, p2:딜레이10)
    Call MakeCmd(arrCmd, 1, "TYPE_MEM", 99, 10, "")
    modEngine.ExecuteSmartAction arrCmd, 1
    
    ' 4. 엔터키
    Call MakeCmd(arrCmd, 1, "KEY", 13, 0, "")
    modEngine.ExecuteSmartAction arrCmd, 1
    
    DoEvents
    If wsTest.Range("A2").Value = "DynamicValue" Then
        LogPass "TYPE_MEM (Var -> Cell)"
    Else
        LogFail "TYPE_MEM", "DynamicValue", wsTest.Range("A2").Value
        failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 9] 소멸 대기 검증 (WAIT_VANISH)
    '==========================================================================
    Debug.Print "--- [Test 9] Wait Vanish ---"
    
    ' 시나리오: E1 셀이 빨간색이다. "빨간색이 사라질 때까지(아니게 될 때까지)" 대기하라.
    ' 하지만 테스트 자동화를 위해 "이미 빨간색이 아닌 상황"을 주고 즉시 통과하는지 확인 (Logic check)
    ' (TimeOut 테스트는 시간이 걸리므로 Logic만 검증)
    
    wsTest.Range("E1").Interior.color = RGB(255, 255, 255) ' 흰색 설정
    
    Dim whiteColor As String
    ' 흰색이 사라지길 기다려라 (근데 이미 흰색임 -> 대기해야 함 -> Timeout 발생 예상)
    ' 역발상 테스트: "검정색(000000)이 사라지길 기다려라" -> 이미 흰색이므로 검정이 아님 -> 즉시 통과
    Dim px, py As Long
    ' E1 좌표 구하기
    Dim targetRange As Range
    Set targetRange = wsTest.Range("E1")
    px = ActiveWindow.PointsToScreenPixelsX(targetRange.Left) + 10
    py = ActiveWindow.PointsToScreenPixelsY(targetRange.Top) + 5
    
    ' 명령: (px, py)가 검정색(000000)이 아니면 통과하라. (현재 흰색이므로 즉시 통과해야 함)
    Call MakeCmd(arrCmd, 1, "WAIT_VANISH", px, py, "000000")
    
    Dim tStart As Long, tEnd As Long
    tStart = modWinAPI.GetTickCount()
    
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1) ' 성공 시 0 리턴일 것임 (WAIT 계열은 리턴값이 0 또는 -1)
    
    tEnd = modWinAPI.GetTickCount()
    
    ' 검증: 1초 이내에 끝났으면 로직 성공 (대기하지 않았으므로)
    If (tEnd - tStart) < 1000 And nextLine <> -1 Then
        LogPass "WAIT_VANISH (Logic Check)"
    Else
        LogFail "WAIT_VANISH", "Immediate Pass", "Took " & (tEnd - tStart) & "ms"
        failCount = failCount + 1
    End If

    '==========================================================================
    ' [Test 10] 창 활성화 (FOCUS)
    '==========================================================================
    Debug.Print "--- [Test 10] Window Focus ---"
    
    ' 시나리오: 엑셀 타이틀의 일부("Excel")를 주면 찾아오는가?
    Dim winTitlePart As String
    winTitlePart = "Excel" ' 보통 캡션에 Excel이 포함됨
    
    Call MakeCmd(arrCmd, 1, "FOCUS", 0, 0, winTitlePart)
    nextLine = modEngine.ExecuteSmartAction(arrCmd, 1)
    
    If nextLine <> -1 Then
        LogPass "FOCUS (Find Window)"
    Else
        LogFail "FOCUS", "Success", "Failed (-1)"
        failCount = failCount + 1
    End If
    '==========================================================================
    ' 종료 처리
    '==========================================================================
    Application.ScreenUpdating = True
    
    Dim resMsg As String
    If failCount = 0 Then
        resMsg = "? 모든 시스템 검증 통과 (All Systems Go)"
    Else
        resMsg = "? " & failCount & "개 항목 실패. 직접 확인하세요."
    End If
    
    MsgBox resMsg, IIf(failCount = 0, vbInformation, vbCritical), "Self Diagnosis"
    
    ' 자동 삭제 (개발 중엔 주석 처리 추천)
    ' Application.DisplayAlerts = False
    ' wsTest.Delete
    ' Application.DisplayAlerts = True
End Sub

' [Helper] 가상 명령어 배열 생성기
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
