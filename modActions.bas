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

' [CLICK_SYNC 전용] 이 시간 내 미안정 색변화는 호버로 간주하지 않음.
' = "호버의 정의" = 커서 진입 후 이 시간 내 전이. 늘리면 클릭반응을 호버로 오인하기 시작함.
Private Const HOVER_SETTLE_MS As Long = 150

Private Function WaitIfPaused() As Boolean
    If Not IsKeyDown(VK_PAUSE_RESUME) Then Exit Function

    Action_ShowNotification "일시정지됨. [Pause] 재개, [ESC] 종료", 3
    Sleep 500

    Do
        If IsKeyDown(VK_ESCAPE) Then
            WaitIfPaused = True
            Exit Function
        End If

        If IsKeyDown(VK_PAUSE_RESUME) Then
            Sleep 200
            Exit Function
        End If

        DoEvents
        Sleep DELAY_MICRO
    Loop
End Function

'==============================================================================
' 1. Core Utilities (핵심 유틸리티)
'==============================================================================

' [스마트 대기] Sleep 대신 사용. 엑셀 응답없음 방지 + ESC 비상종료 감지
Public Sub BetterWait(ByVal totalWait As Long)
    On Error GoTo ErrHandler
    Dim i As Long, stepWait As Long: stepWait = DELAY_MICRO
    
    If totalWait < stepWait Then
        If IsKeyDown(VK_ESCAPE) Then Exit Sub
        If WaitIfPaused() Then Exit Sub
        Sleep totalWait
        DoEvents
        GoTo CleanUp
    End If
    For i = 1 To totalWait \ stepWait
        If IsKeyDown(VK_ESCAPE) Then Exit Sub
        If WaitIfPaused() Then Exit Sub
        Sleep stepWait
        DoEvents
    Next i

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.BetterWait", Err.Description
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
        If WaitIfPaused() Then Exit Sub
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
    Err.Raise Err.Number, "modActions.Action_InputText", Err.Description
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
    Err.Raise Err.Number, "modActions.Action_SmartType", Err.Description
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
    Err.Raise Err.Number, "modActions.Action_PressKey", Err.Description
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
    Err.Raise Err.Number, "modActions.Action_PressCombo", Err.Description
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

' [프리미티브] 좌표 이동 + 안정화 대기
Public Sub MoveAndSettle(ByVal x As Long, ByVal y As Long)
    On Error GoTo ErrHandler
    modWinAPI.SetCursorPos x, y
    BetterWait DELAY_SHORT   ' 이동 후 안정화 (너무 빠르면 클릭 씹힘)

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.MoveAndSettle", Err.Description
    Resume CleanUp
End Sub

' [프리미티브] 현재 커서 위치에서 Down/Up (좌/우)
Public Sub FireClick(Optional ByVal isRightClick As Boolean = False)
    On Error GoTo ErrHandler
    If isRightClick Then
        modWinAPI.mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        BetterWait DELAY_MICRO ' 누르고 잠시 대기
        modWinAPI.mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Else
        modWinAPI.mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        BetterWait DELAY_MICRO
        modWinAPI.mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    End If

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.FireClick", Err.Description
    Resume CleanUp
End Sub

' [마우스 클릭] 이동 -> 안정화 -> 클릭 -> 안정화 (CLICK/RCLICK = fire-and-forget)
Public Sub Action_MouseClick(ByVal x As Long, ByVal y As Long, Optional ByVal isRightClick As Boolean = False)
    On Error GoTo ErrHandler
    MoveAndSettle x, y
    FireClick isRightClick
    BetterWait DELAY_SHORT

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_MouseClick", Err.Description
    Resume CleanUp
End Sub

' [확인된 좌클릭] CLICK_SYNC (§2.1): 이동 -> 호버안정 -> 기준색캡처 -> 좌클릭 -> 색변화까지 대기
'   타임아웃 시 조용히 False 반환 (모달 없음). 좌클릭 전용 (RCLICK_SYNC/KEY_SYNC 확장 금지).
Public Function Action_ClickSync(ByVal x As Long, ByVal y As Long) As Boolean
    On Error GoTo ErrHandler
    Dim baseColor As String
    Dim startTime As Long

    ' 1. 이동 -> 호버 전이가 끝난 뒤 기준색 캡처 (호버색을 기준선에 포함 = 빼버림)
    MoveAndSettle x, y
    BetterWait HOVER_SETTLE_MS
    baseColor = modWinAPI.GetColorHexAt(x, y)

    ' 2. 좌클릭 발사
    FireClick False

    ' 3. (x,y) 색이 기준색에서 벗어날 때까지 폴링. 동일점 가드(누른 곳=확인 곳).
    startTime = modWinAPI.GetTickCount()
    Do
        If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then Action_ClickSync = False: Exit Function
        If modWinAPI.GetColorHexAt(x, y) <> baseColor Then Action_ClickSync = True: Exit Function
        If (modWinAPI.GetTickCount() - startTime) > TIMEOUT_DEFAULT Then
            Action_ClickSync = False    ' 조용히 실패 (모달 금지)
            Exit Function
        End If
        BetterWait DELAY_SHORT
    Loop

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_ClickSync", Err.Description
    Resume CleanUp
End Function

'==============================================================================
' 4. Vision Actions (화면 인식 및 대기)
'==============================================================================

' [통합 대기] 특정 색상 출현(WAIT_APPEAR) 또는 소멸(WAIT_VANISH) 대기
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
        
        ' 타임아웃 (무인 환경: 모달 금지, 조용히 실패 반환 → 알림 여부는 엔진이 결정)
        If (modWinAPI.GetTickCount() - startTime) > timeLimit Then
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
    Err.Raise Err.Number, "modActions.Action_Wait", Err.Description
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
    Err.Raise Err.Number, "modActions.Action_WaitForColor", Err.Description
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
    
    ' 무인 환경: 모달 금지, 불일치 시 조용히 False 반환 (알림은 엔진이 결정)
    If isForbidden Then
        If nowColor = targetColor Then Action_Assert = False: Exit Function
    Else
        If nowColor <> targetColor Then Action_Assert = False: Exit Function
    End If
    Action_Assert = True

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_Assert", Err.Description
    Resume CleanUp
End Function



'==============================================================================
' 5. Window & System Actions (윈도우 제어)
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
    Err.Raise Err.Number, "modActions.Action_FocusWindow", Err.Description
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
        ' [Equal] 정확히 같다 (== : *, ? 도 리터럴로 취급,  아님)
        Case "=="
            If isNum Then
                Action_CheckCondition = (d1 = d2)
            Else
                ' 명시 텍스트=대소문자 구분(원본), 암묵 텍스트=UCase 정규화됨
                Action_CheckCondition = (s1 = s2)
            End If

        ' [Not Equal] 다르다 (!=)
        Case "!="
            If isNum Then
                Action_CheckCondition = (d1 <> d2)
            Else
                Action_CheckCondition = (s1 <> s2)
            End If

        ' [Greater Than] 크다 (>)
        Case ">"
            If isNum Then Action_CheckCondition = (d1 > d2) Else Action_CheckCondition = (s1 > s2)

        ' [Less Than] 작다 (<)
        Case "<"
            If isNum Then Action_CheckCondition = (d1 < d2) Else Action_CheckCondition = (s1 < s2)

        ' [Greater or Equal] 크거나 같다 (>=)
        Case ">="
            If isNum Then Action_CheckCondition = (d1 >= d2) Else Action_CheckCondition = (s1 >= s2)

        ' [Less or Equal] 작거나 같다 (<=)
        Case "<="
            If isNum Then Action_CheckCondition = (d1 <= d2) Else Action_CheckCondition = (s1 <= s2)

        ' [Pattern] LIKE (* 임의 / ? 한글자 / # 한숫자) - 와일드카드는 여기서만
        Case "LIKE"
            Action_CheckCondition = (s1 Like s2)

        ' [Starts With] ~로 시작
        Case "STARTS"
            Action_CheckCondition = (Len(s2) > 0 And Left$(s1, Len(s2)) = s2)

        ' [Ends With] ~로 끝
        Case "ENDS"
            Action_CheckCondition = (Len(s2) > 0 And Right$(s1, Len(s2)) = s2)

        ' [String Contains] 포함 (CONTAINS)
        Case "CONTAINS"
            ' 명시적이면 BinaryCompare(대소문자 구분), 암묵적이면 TextCompare(이미 UCase됨)
            If isExplicit Then
                Action_CheckCondition = (InStr(1, s1, s2, vbBinaryCompare) > 0)
            Else
                Action_CheckCondition = (InStr(1, s1, s2, vbTextCompare) > 0)
            End If

        Case Else
            ' 알 수 없는 연산자 (맨 '=' 포함) -> False
            Action_CheckCondition = False
    End Select

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_CheckCondition", Err.Description
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
    Debug.Print "[ERROR] Action_ShowNotification: " & Err.Description
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
    Err.Raise Err.Number, "modActions.Action_Notify", Err.Description
    Resume CleanUp
End Sub
