Attribute VB_Name = "modActions"
Option Explicit

'==============================================================================
' [Layer 3] modActions: RPA Activity Implementation
'==============================================================================
' 역할: modWinAPI를 사용하여 구체적인 동작(클릭, 입력, 대기)을 수행
' 원칙:
' 1. 모든 동작에 "비상 탈출(ESC)" 로직 포함 (BetterWait)
' 2. 입력 방식의 이원화 (Paste vs Type) 지원
' 3. 비즈니스 로직(판단)은 최소화하고 "행동"에 집중
'==============================================================================

' [상수 정의] 매직 넘버 제거
Private Const DELAY_MICRO As Long = 50
Private Const DELAY_SHORT As Long = 100
Private Const DELAY_MEDIUM As Long = 200
Private Const TIMEOUT_DEFAULT As Long = 30000

'==============================================================================
' 1. Core Utilities (핵심 유틸리티)
'==============================================================================

' [스마트 대기] Sleep 대신 사용. 엑셀 응답없음 방지 + ESC 비상종료 감지
Public Sub BetterWait(ByVal totalWait As Long)
    On Error GoTo ErrHandler
    Dim i As Long, stepWait As Long: stepWait = DELAY_MICRO
    
    If totalWait < stepWait Then
        If IsKeyDown(VK_ESCAPE) Then Exit Sub
        If IsKeyDown(VK_PAUSE_RESUME) Then Exit Sub
        Sleep totalWait
        DoEvents
        GoTo CleanUp
    End If
    For i = 1 To totalWait \ stepWait
        If IsKeyDown(VK_ESCAPE) Then Exit Sub
        If IsKeyDown(VK_PAUSE_RESUME) Then Exit Sub
        Sleep stepWait
        DoEvents
    Next i

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

'==============================================================================
' 2. Keyboard Input Actions (키보드 제어)
'==============================================================================

' [저수준 입력] Unicode 지원 + KeyUp/Down 분리 구현
Public Sub Action_InputText(ByVal txt As String, ByVal delayMs As Long)
    On Error GoTo ErrHandler
    Dim i As Long
    Dim charCode As Long
    
    For i = 1 To Len(txt)
        If IsKeyDown(VK_ESCAPE) Then Exit Sub
        If IsKeyDown(VK_PAUSE_RESUME) Then Exit Sub
        charCode = AscW(Mid(txt, i, 1))
        
        ' 1. Key Down (Unicode Flag 사용)
        SendKeyboardInput 0, charCode, KEYEVENTF_UNICODE
        
        ' 2. Key Up
        SendKeyboardInput 0, charCode, KEYEVENTF_UNICODE Or KEYEVENTF_KEYUP
        
        If delayMs > 0 Then BetterWait delayMs
    Next i

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

' [고수준 래퍼] Action_InputText를 안전하게 호출 (딜레이 보정)
Public Sub Action_SmartType(delayParam As Long, text As String)
    On Error GoTo ErrHandler
    Dim d As Long: d = delayParam
    
    ' [핵심] 딜레이가 0이면 너무 빨라서 씹히거나 비동기처럼 보일 수 있음
    ' 따라서 0일 경우 강제로 50ms를 주어 "타이핑 중 대기"를 보장함
    If d <= 0 Then d = DELAY_MICRO
    
    If Len(text) > 0 Then
        ' 님께서 작성하신 원본 함수 호출 (건드리지 않음)
        Action_InputText text, d
    End If

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

' [특수키 입력] Enter, Tab, Esc 등 단일 키 입력
Public Sub Action_PressKey(ByVal vKey As Variant)
    On Error GoTo ErrHandler
    Dim keyCode As Long
    keyCode = ParseKeyCode(vKey)
    modWinAPI.SendKey_Click keyCode
    BetterWait DELAY_MICRO ' 키 입력 후 미세 대기

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

' [조합키 입력] Ctrl+C 등 (순차적 Down -> 순차적 Up)
Public Sub Action_PressCombo(k1 As Variant, k2 As Variant)
    On Error GoTo ErrHandler
    Dim c1 As Long, c2 As Long
    c1 = ParseKeyCode(k1)
    c2 = ParseKeyCode(k2)
    modWinAPI.SendKey_Down c1
    modWinAPI.Sleep DELAY_MICRO
    modWinAPI.SendKey_Down c2
    modWinAPI.Sleep DELAY_MICRO
    modWinAPI.SendKey_Up c2
    modWinAPI.Sleep DELAY_MICRO
    modWinAPI.SendKey_Up c1
    modWinAPI.Sleep DELAY_MICRO

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

' [내부 함수] 키 코드 파싱 (숫자 or 텍스트 -> Virtual Key Code)
Private Function ParseKeyCode(ByVal vKey As Variant) As Long
    If IsNumeric(vKey) Then
        ParseKeyCode = CLng(vKey)
        Exit Function
    End If
    
    Dim sKey As String
    sKey = UCase(Trim(CStr(vKey)))
    
    Select Case sKey
        Case "ENTER", "RETURN": ParseKeyCode = 13
        Case "ESC", "ESCAPE":   ParseKeyCode = 27
        Case "TAB":             ParseKeyCode = 9
        Case "CTRL", "CONTROL": ParseKeyCode = 17
        Case "ALT", "MENU":     ParseKeyCode = 18
        Case "SHIFT":           ParseKeyCode = 16
        Case "WIN", "WINDOWS":  ParseKeyCode = 91
        Case "SPACE", "SPC":    ParseKeyCode = 32
        Case "BS", "BACKSPACE": ParseKeyCode = 8
        Case "DEL", "DELETE":   ParseKeyCode = 46
        Case "INS", "INSERT":   ParseKeyCode = 45
        Case "HOME":            ParseKeyCode = 36
        Case "END":             ParseKeyCode = 35
        Case "PGUP", "PAGEUP":  ParseKeyCode = 33
        Case "PGDN", "PAGEDOWN": ParseKeyCode = 34
        Case "UP":              ParseKeyCode = 38
        Case "DOWN":            ParseKeyCode = 40
        Case "LEFT":            ParseKeyCode = 37
        Case "RIGHT":           ParseKeyCode = 39
        Case "PRTSC":           ParseKeyCode = 44
        Case "CAPS":            ParseKeyCode = 20
        Case "F1": ParseKeyCode = 112: Case "F2": ParseKeyCode = 113: Case "F3": ParseKeyCode = 114
        Case "F4": ParseKeyCode = 115: Case "F5": ParseKeyCode = 116: Case "F6": ParseKeyCode = 117
        Case "F7": ParseKeyCode = 118: Case "F8": ParseKeyCode = 119: Case "F9": ParseKeyCode = 120
        Case "F10": ParseKeyCode = 121: Case "F11": ParseKeyCode = 122: Case "F12": ParseKeyCode = 123
        Case Else
            ' [알파벳/숫자] 1글자인 경우 ASCII 값 사용 (A-Z, 0-9는 VK코드와 일치)
            ' UCase로 변환했으므로 소문자 a도 A(65)로 변환되어 정상 동작함
            If Len(sKey) = 1 Then ParseKeyCode = Asc(sKey) Else ParseKeyCode = 0
    End Select
End Function

'==============================================================================
' 3. Mouse Actions (마우스 제어)
'==============================================================================

' [마우스 클릭] 이동 -> 대기 -> 클릭 -> 대기
Public Sub Action_MouseClick(ByVal x As Long, ByVal y As Long, Optional ByVal isRightClick As Boolean = False)
    On Error GoTo ErrHandler
    ' 1. 좌표 이동
    modWinAPI.SetCursorPos x, y
    
    ' 2. 이동 후 안정화 (중요: 너무 빠르면 클릭 씹힘)
    BetterWait DELAY_SHORT
    
    ' 3. 클릭 수행
    If isRightClick Then
        modWinAPI.mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        BetterWait DELAY_MICRO ' 누르고 잠시 대기
        modWinAPI.mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Else
        modWinAPI.mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        BetterWait DELAY_MICRO
        modWinAPI.mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    End If
    
    BetterWait DELAY_SHORT

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub

'==============================================================================
' 4. Vision Actions (화면 인식 및 대기)
'==============================================================================

' [통합 대기] 특정 색상 출현(WAIT_PIXEL) 또는 소멸(WAIT_VANISH) 대기
Public Function Action_Wait(sType As String, x As Long, y As Long, hexColor As String) As Boolean
    On Error GoTo ErrHandler
    Dim startTime As Long
    Dim timeLimit As Long: timeLimit = TIMEOUT_DEFAULT ' 30초 제한
    Dim isVanish As Boolean: isVanish = (sType = "WAIT_VANISH")
    Dim curColor As String
    
    startTime = modWinAPI.GetTickCount()
    
    Do
        curColor = modWinAPI.GetColorHexAt(x, y)
        
        ' 조건 충족 확인
        If isVanish Then
            If curColor <> UCase(hexColor) Then Action_Wait = True: Exit Function
        Else
            If curColor = UCase(hexColor) Then Action_Wait = True: Exit Function
        End If
        
        ' 타임아웃
        If (modWinAPI.GetTickCount() - startTime) > timeLimit Then
            LogShim.LogJ "wait_timeout", _
                "{""type"":""" & sType & """,""color"":""" & hexColor & """,""x"":" & x & ",""y"":" & y & "}"
            Action_Wait = False
            Exit Function
        End If
        
        ' 비상 탈출
        If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then Action_Wait = False: Exit Function
        
        BetterWait DELAY_SHORT
    Loop

CleanUp:
    Exit Function
ErrHandler:
    LogShim.LogJ "wait_error", _
        "{""num"":" & Err.Number & ",""src"":""Action_Wait""}"
    Action_Wait = False
    Resume CleanUp
End Function

' [단순 대기] 특정 좌표가 특정 색상이 될 때까지 대기 (Timeout 적용 - 구버전 호환용)
Public Function Action_WaitForColor(ByVal x As Long, ByVal y As Long, ByVal hexColor As String, ByVal TimeoutSec As Long) As Boolean
    On Error GoTo ErrHandler
    Dim startTime As Long
    Dim currColor As String
    
    startTime = modWinAPI.GetTickCount()
    
    Do
        ' 1. 현재 색상 추출
        currColor = modWinAPI.GetColorHexAt(x, y)
        
        ' 2. 일치 여부 확인
        If currColor = hexColor Then
            BetterWait DELAY_MEDIUM ' 깜빡임 방지용 2차 확인 대기
            If modWinAPI.GetColorHexAt(x, y) = hexColor Then
                Action_WaitForColor = True
                Exit Function
            End If
        End If
        
        ' 3. 타임아웃 체크
        If (modWinAPI.GetTickCount() - startTime) > (TimeoutSec * 1000) Then
            Action_WaitForColor = False
            Exit Function
        End If
        
        BetterWait DELAY_SHORT ' 루프 주기
    Loop

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [화면 검증] Assert (기대값과 다르면 실패 처리)
Public Function Action_Assert(x As Long, y As Long, condition As String) As Boolean
    On Error GoTo ErrHandler
    Dim targetColor As String, nowColor As String
    Dim isForbidden As Boolean
    
    targetColor = condition
    isForbidden = (Left(targetColor, 1) = "!")
    If isForbidden Then targetColor = Mid(targetColor, 2)
    
    nowColor = modWinAPI.GetColorHexAt(x, y)
    
    If isForbidden Then
        If nowColor = targetColor Then
            MsgBox "[Assert Fail] 금지된 색상 감지!", vbCritical
            Action_Assert = False: Exit Function
        End If
    Else
        If nowColor <> targetColor Then
            MsgBox "[Assert Fail] 색상 불일치 (Exp:" & targetColor & " Act:" & nowColor & ")", vbCritical
            Action_Assert = False: Exit Function
        End If
    End If
    Action_Assert = True

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function



'==============================================================================
' 5. Logic Helpers (판단 로직 위임), 파서 로직
'==============================================================================

' [화면 조건 파싱] IF_PIXEL, SWITCH 등 복합 조건 문자열 파싱 및 결과 반환
Public Function GetJumpTarget_Pixel(sType As String, x As Long, y As Long, conditionStr As String) As Long
    On Error GoTo ErrHandler
    Dim parts() As String, subParts() As String
    Dim curColor As String
    Dim i As Long
    
    curColor = modWinAPI.GetColorHexAt(x, y)
    GetJumpTarget_Pixel = 0
    
    Select Case sType
        Case "IF_PIXEL"
            parts = Split(conditionStr, "|")
            ' 잠시 대기하며 색상 확인 (Wait Logic 재사용 가능)
            If Action_WaitForColor(x, y, parts(0), 0.2) Then GetJumpTarget_Pixel = val(parts(1))
            
        Case "IFNOT_PIXEL"
            parts = Split(conditionStr, "|")
            BetterWait DELAY_MEDIUM ' 안정화
            If curColor <> UCase(parts(0)) Then GetJumpTarget_Pixel = val(parts(1))
            
        Case "SWITCH_PIXEL"
            parts = Split(conditionStr, ";")
            For i = LBound(parts) To UBound(parts)
                subParts = Split(parts(i), "|")
                If UBound(subParts) >= 1 Then
                    If curColor = UCase(subParts(0)) Then
                        GetJumpTarget_Pixel = val(subParts(1))
                        Exit Function
                    End If
                End If
            Next i
    End Select

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [텍스트 조건 파싱] IF_CLIP, IF_MEM 등 문자열 포함 여부 확인
Public Function GetJumpTarget_Text(sType As String, sourceText As String, conditionStr As String) As Long
    On Error GoTo ErrHandler
    Dim parts() As String
    Dim targetText As String, jumpLine As Long
    Dim isMatch As Boolean
    
    If InStr(conditionStr, "|") = 0 Then Exit Function
    
    parts = Split(conditionStr, "|")
    targetText = parts(0)
    jumpLine = val(parts(1))
    
    ' 포함 여부 검사
    isMatch = (InStr(1, sourceText, targetText, vbTextCompare) > 0)
    
    If sType = "IFNOT_CLIP" Or sType = "IFNOT_MEM" Then isMatch = Not isMatch
    
    If isMatch Then GetJumpTarget_Text = jumpLine Else GetJumpTarget_Text = 0

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

'==============================================================================
' 6. Window & System Actions (윈도우 제어)
'==============================================================================

' [창 활성화] 제목의 일부만으로 창을 찾아 맨 앞으로 가져옴
Public Function Action_FocusWindow(ByVal partialTitle As String) As Boolean
    On Error GoTo ErrHandler
    Dim hwnd As LongPtr
    Dim winTitle As String
    Dim lenTitle As Long
    
    ' 현재 실행 중인 첫 번째 윈도우 핸들
    hwnd = FindWindow(vbNullString, vbNullString)
    
    Do While hwnd <> 0
        lenTitle = GetWindowTextLength(hwnd)
        If lenTitle > 0 Then
            winTitle = String(lenTitle, vbNullChar)
            GetWindowText hwnd, winTitle, lenTitle + 1
            winTitle = Left(winTitle, lenTitle)
            
            ' 부분 일치 검사
            If InStr(1, winTitle, partialTitle, vbTextCompare) > 0 Then
                ' 최소화 상태면 복구
                If IsIconic(hwnd) Then
                    ShowWindow hwnd, SW_RESTORE
                Else
                    SetForegroundWindow hwnd
                End If
                Action_FocusWindow = True
                Exit Function
            End If
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    
    Action_FocusWindow = False

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Logic] 조건 비교 (개선됨)
Public Function Action_CheckCondition(ByVal v1 As String, ByVal op As String, ByVal v2 As String) As Boolean
    On Error GoTo ErrHandler
    Dim d1 As Double, d2 As Double
    Dim s1 As String, s2 As String
    Dim isNum As Boolean
    Dim isExplicit As Boolean
    Dim q1 As Boolean, q2 As Boolean
    
    ' 1. 입력값 정리
    v1 = Trim(v1)
    v2 = Trim(v2)
    
    ' 2. 명시적 텍스트(쌍따옴표) 감지 및 제거
    If Left(v1, 1) = """" And Right(v1, 1) = """" And Len(v1) >= 2 Then
        s1 = Mid(v1, 2, Len(v1) - 2)
        q1 = True
    Else
        s1 = v1
    End If
    
    If Left(v2, 1) = """" And Right(v2, 1) = """" And Len(v2) >= 2 Then
        s2 = Mid(v2, 2, Len(v2) - 2)
        q2 = True
    Else
        s2 = v2
    End If
    
    ' 3. 모드 결정
    ' 하나라도 쌍따옴표가 있으면 -> 텍스트 모드 (와일드카드 지원, 대소문자 구분)
    If q1 Or q2 Then
        isExplicit = True
        isNum = False
    Else
        ' 둘 다 따옴표가 없으면 -> 숫자 변환 시도
        ' [Smart Check] "0123" 처럼 0으로 시작하는 정수형 텍스트(0#*)는 문자열로 취급 (0.xx는 제외)
        If IsNumeric(s1) And IsNumeric(s2) And Not (s1 Like "0#*" Or s2 Like "0#*") Then
            d1 = val(s1)
            d2 = val(s2)
            isNum = True
        Else
            ' 변환 실패 또는 식별자형 숫자(0123 등) -> 암묵적 텍스트 모드
            isNum = False
            s1 = UCase(s1)
            s2 = UCase(s2)
        End If
    End If
    
    op = UCase(Trim(op))
    
    Select Case op
        ' [Equal] 같다 (Like 연산자 사용하여 와일드카드 지원)
        Case "==", "=", "IS", "EQ", ""
            If isNum Then
                Action_CheckCondition = (d1 = d2)
            Else
                ' 텍스트 모드: Like 사용 (s2에 *, ?가 있으면 패턴 매칭됨)
                Action_CheckCondition = (s1 Like s2)
            End If
            
        ' [Not Equal] 다르다 (!=, <>, ISNOT, NEQ)
        Case "!=", "<>", "ISNOT", "NEQ", "NOT"
            If isNum Then
                Action_CheckCondition = (d1 <> d2)
            Else
                Action_CheckCondition = Not (s1 Like s2)
            End If
            
        ' [Greater Than] 크다 (>, GT)
        Case ">", "GT"
            If isNum Then Action_CheckCondition = (d1 > d2) Else Action_CheckCondition = (s1 > s2)
            
        ' [Less Than] 작다 (<, LT)
        Case "<", "LT"
            If isNum Then Action_CheckCondition = (d1 < d2) Else Action_CheckCondition = (s1 < s2)
            
        ' [Greater or Equal] 크거나 같다 (>=, GE, GTE)
        Case ">=", "GE", "GTE"
            If isNum Then Action_CheckCondition = (d1 >= d2) Else Action_CheckCondition = (s1 >= s2)
            
        ' [Less or Equal] 작거나 같다 (<=, LE, LTE)
        Case "<=", "LE", "LTE"
            If isNum Then Action_CheckCondition = (d1 <= d2) Else Action_CheckCondition = (s1 <= s2)
            
        ' [String Contains] 포함 (CONTAINS, HAS)
        Case "CONTAINS", "HAS", "IN"
            ' 명시적이면 BinaryCompare(대소문자 구분), 암묵적이면 TextCompare(이미 UCase됨)
            If isExplicit Then
                Action_CheckCondition = (InStr(1, s1, s2, vbBinaryCompare) > 0)
            Else
                Action_CheckCondition = (InStr(1, s1, s2, vbTextCompare) > 0)
            End If
            
        Case Else
            Action_CheckCondition = False
    End Select

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [알림 표시] 비차단형 메시지박스 (실행 흐름을 막지 않음)
Public Sub Action_ShowNotification(ByVal msg As String, Optional ByVal durationSec As Long = 2)
    On Error GoTo ErrHandler
    Dim flags As Long
    flags = MB_ICONINFO + MB_TOPMOST + MB_SETFOREGROUND
    modWinAPI.MessageBoxTimeout 0, msg, "RPA Notification", flags, 0, durationSec * 1000

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub
' [modActions] 사용자 알림 (ms 단위 준수)
' p1 (durationMs): 0이면 무한대기(MsgBox), 0보다 크면 자동종료(ms 단위)
' p2 (iconType): 아이콘 번호 (16, 32, 48, 64 등)
' p3 (msg): 출력할 메시지
Public Sub Action_Notify(ByVal durationMs As Long, ByVal iconType As Long, ByVal msg As String)
    On Error GoTo ErrHandler
    Dim flags As Long
    ' 기본 플래그: 항상 위 + 포커스
    flags = iconType + MB_TOPMOST + MB_SETFOREGROUND
    
    If durationMs <= 0 Then
        ' [차단형] 확인 누를 때까지 정지
        MsgBox msg, flags, "RPA Alert"
    Else
        ' [비차단형] p1(ms) 만큼 대기 후 자동 종료
        ' 별도의 연산 없이 durationMs를 그대로 전달합니다.
        modWinAPI.MessageBoxTimeout 0, msg, "RPA Notification", flags, 0, durationMs
    End If

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub
