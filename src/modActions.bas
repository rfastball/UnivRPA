Attribute VB_Name = "modActions"
' @desc 클릭/입력/대기 등 구체 동작 수행 담당
' @depends modWinAPI, LogShim
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
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

' [무인 모드] True면 알림을 팝업 대신 LogShim jsonl로 강등 (vbaflow test 등 헤드리스 실행)
Private m_unattended As Boolean

' [중단 래치 (B2)] ESC 물리키 폴링의 단일 소유자. 한 번 감지되면 래치되어
' 실행 전체가 중단 경로로 수렴한다 (순간 키입력 유실 방지).
Private m_abortRequested As Boolean
Private m_inputTestDriver As Object

' @desc 무인 모드 여부 설정함 (True=팝업/일시정지 억제, 로그 기록만)
Public Sub SetUnattendedMode(ByVal enabled As Boolean)
    m_unattended = enabled
    If Not enabled Then Set m_inputTestDriver = Nothing
End Sub


' @desc 무인 입력 검증용 드라이버 설정함, Nothing이면 실제 API 사용
Public Sub SetInputTestDriver(ByVal driver As Object)
    If Not m_unattended And Not driver Is Nothing Then Err.Raise 5, "modActions", "Unattended mode required."
    Set m_inputTestDriver = driver
End Sub

Private Sub SleepSlice(ByVal milliseconds As Long)
    If m_inputTestDriver Is Nothing Then
        modWinAPI.Sleep milliseconds
        DoEvents
    Else
        m_inputTestDriver.SleepMilliseconds milliseconds
    End If
End Sub

Private Sub SendMouseButton(ByVal flags As Long)
    If m_inputTestDriver Is Nothing Then
        modWinAPI.mouse_event flags, 0, 0, 0, 0
    Else
        m_inputTestDriver.MouseButton flags
    End If
End Sub

Private Sub SendKeyboardEvent(ByVal vk As Integer, ByVal scan As Integer, ByVal flags As Long)
    modWinAPI.SendKeyboardInput vk, scan, flags, m_inputTestDriver
End Sub

' Down 성공 후에는 중단과 무관하게 Up을 시도한다. 정리 실패는 원래 오류를 덮지 않는다.
Private Sub SendKeyStroke(ByVal vk As Integer, ByVal scan As Integer, ByVal flags As Long)
    Dim held As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    If AbortPending() Then Exit Sub
    SendKeyboardEvent vk, scan, flags
    held = True
    SendKeyboardEvent vk, scan, flags Or KEYEVENTF_KEYUP
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If held Then SendKeyboardEvent vk, scan, flags Or KEYEVENTF_KEYUP
    On Error GoTo 0
    Err.Raise errorNumber, "modActions.SendKeyStroke", errorText
End Sub

' @desc 중단 래치를 설정함 (ESC 감지 또는 외부 요청)
Public Sub RequestAbort()
    m_abortRequested = True
End Sub

' @desc 중단 래치를 해제함 (실행 시작 시 엔진/하니스가 호출)
Public Sub ClearAbort()
    m_abortRequested = False
End Sub

' @desc 중단 여부 반환: 래치 우선 확인, 미설정 시 ESC 1회 폴링해 래치함
Public Function AbortPending() As Boolean
    If m_abortRequested Then
        AbortPending = True
        Exit Function
    End If
    If Not m_inputTestDriver Is Nothing Then Exit Function
    If IsKeyDown(VK_ESCAPE) Then
        m_abortRequested = True
        AbortPending = True
    End If
End Function

' @desc Pause 해제까지 대기하고 ESC 중단 시 True 반환함
Public Function WaitIfPaused(Optional ByRef resumed As Boolean = False) As Boolean
    resumed = False
    If AbortPending() Then
        WaitIfPaused = True
        Exit Function
    End If
    If m_unattended Then Exit Function
    If Not IsKeyDown(VK_PAUSE_RESUME) Then Exit Function

    Action_ShowNotification "일시정지됨. [Pause] 재개, [ESC] 종료", 3
    Sleep 500

    Do
        If AbortPending() Then
            WaitIfPaused = True
            Exit Function
        End If

        If IsKeyDown(VK_PAUSE_RESUME) Then
            Sleep 200
            resumed = True
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
    If totalWait < 0 Then Err.Raise 5, "modActions.BetterWait", "Wait milliseconds must not be negative."
    Dim remaining As Long, slice As Long
    remaining = totalWait
    Do
        If WaitIfPaused() Then Exit Sub
        slice = remaining
        If slice > DELAY_MICRO Then slice = DELAY_MICRO
        SleepSlice slice
        remaining = remaining - slice
    Loop While remaining > 0
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
        If AbortPending() Then Exit Sub
        If WaitIfPaused() Then Exit Sub
        charCode = AscW(Mid(txt, i, 1))
        
        SendKeyStroke 0, CInt(charCode), KEYEVENTF_UNICODE

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
    SendKeyStroke CInt(keyCode), 0, 0
    BetterWait DELAY_MICRO ' 키 입력 후 미세 대기

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_PressKey", Err.Description
    Resume CleanUp
End Sub

' [조합키 입력] Ctrl+C 등 (순차적 Down -> 순차적 Up)
Public Sub Action_PressCombo(k1 As Variant, k2 As Variant)
    Dim c1 As Long, c2 As Long
    Dim held1 As Boolean, held2 As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    c1 = ParseKeyCode(k1)
    c2 = ParseKeyCode(k2)
    If AbortPending() Then Exit Sub
    SendKeyboardEvent CInt(c1), 0, 0
    held1 = True
    SleepSlice DELAY_MICRO
    If Not AbortPending() Then
        SendKeyboardEvent CInt(c2), 0, 0
        held2 = True
        SleepSlice DELAY_MICRO
    End If
    If held2 Then
        SendKeyboardEvent CInt(c2), 0, KEYEVENTF_KEYUP
        held2 = False
        SleepSlice DELAY_MICRO
    End If
    SendKeyboardEvent CInt(c1), 0, KEYEVENTF_KEYUP
    held1 = False
    BetterWait DELAY_MICRO
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If held2 Then SendKeyboardEvent CInt(c2), 0, KEYEVENTF_KEYUP
    If held1 Then SendKeyboardEvent CInt(c1), 0, KEYEVENTF_KEYUP
    On Error GoTo 0
    Err.Raise errorNumber, "modActions.Action_PressCombo", errorText
End Sub

' [내부 함수] 키 코드 파싱 (숫자 or 텍스트 -> Virtual Key Code)
Private Function ParseKeyCode(ByVal vKey As Variant) As Long
    If IsNumeric(vKey) Then
        ParseKeyCode = CLng(vKey)
        If ParseKeyCode <= 0 Or ParseKeyCode > 254 Then Err.Raise 5, "modActions.ParseKeyCode", "잘못된 키 코드: " & CStr(vKey)
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
            ' Only letters share their character code with the intended virtual key.
            If sKey Like "[A-Z]" Then ParseKeyCode = Asc(sKey)
    End Select

    ' 알 수 없는 키 이름은 VK 0으로 조용히 발사되지 않도록 명시 오류로 전환
    If ParseKeyCode <= 0 Then Err.Raise 5, "modActions.ParseKeyCode", "알 수 없는 키: " & CStr(vKey)
End Function

'==============================================================================
' 3. Mouse Actions (마우스 제어)
'==============================================================================

' [프리미티브] 좌표 이동 + 안정화 대기
Public Sub MoveAndSettle(ByVal x As Long, ByVal y As Long)
    On Error GoTo ErrHandler
    If AbortPending() Then Exit Sub
    Dim moved As Long
    If m_inputTestDriver Is Nothing Then
        moved = modWinAPI.SetCursorPos(x, y)
    Else
        moved = m_inputTestDriver.MoveCursor(x, y)
    End If
    If moved = 0 Then Err.Raise 5, "modActions.MoveAndSettle", "SetCursorPos failed."
    BetterWait DELAY_SHORT   ' 이동 후 안정화 (너무 빠르면 클릭 씹힘)

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.MoveAndSettle", Err.Description
    Resume CleanUp
End Sub

' [프리미티브] 현재 커서 위치에서 Down/Up (좌/우)
Public Sub FireClick(Optional ByVal isRightClick As Boolean = False)
    Dim downFlag As Long, upFlag As Long
    Dim held As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    If AbortPending() Then Exit Sub
    downFlag = MOUSEEVENTF_LEFTDOWN: upFlag = MOUSEEVENTF_LEFTUP
    If isRightClick Then
        downFlag = MOUSEEVENTF_RIGHTDOWN: upFlag = MOUSEEVENTF_RIGHTUP
    End If
    SendMouseButton downFlag
    held = True
    ' 버튼을 놓기 전에는 Pause 대기에 진입하지 않는다.
    SleepSlice DELAY_MICRO
    SendMouseButton upFlag
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If held Then SendMouseButton upFlag
    On Error GoTo 0
    Err.Raise errorNumber, "modActions.FireClick", errorText
End Sub

' [마우스 클릭] 이동 -> 안정화 -> 클릭 -> 안정화 (CLICK/RCLICK = fire-and-forget)
Public Sub Action_MouseClick(ByVal x As Long, ByVal y As Long, Optional ByVal isRightClick As Boolean = False)
    On Error GoTo ErrHandler
    MoveAndSettle x, y
    If AbortPending() Then Exit Sub
    FireClick isRightClick
    BetterWait DELAY_SHORT

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_MouseClick", Err.Description
    Resume CleanUp
End Sub

' [마우스 이동] 클릭 없이 좌표 이동 -> 안정화
Public Sub Action_MouseMove(ByVal x As Long, ByVal y As Long)
    On Error GoTo ErrHandler
    MoveAndSettle x, y

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_MouseMove", Err.Description
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
    If AbortPending() Then Exit Function
    baseColor = modWinAPI.GetColorHexAt(x, y, m_inputTestDriver)

    ' 2. 좌클릭 발사
    FireClick False

    ' 3. (x,y) 색이 기준색에서 벗어날 때까지 폴링. 동일점 가드(누른 곳=확인 곳).
    startTime = modWinAPI.GetTickCount()
    Do
        If AbortPending() Then Action_ClickSync = False: Exit Function
        If modWinAPI.GetColorHexAt(x, y, m_inputTestDriver) <> baseColor Then Action_ClickSync = True: Exit Function
        If modWinAPI.ElapsedMilliseconds(startTime, modWinAPI.GetTickCount()) > TIMEOUT_DEFAULT Then
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
    Dim targetColor As String
    
    If AbortPending() Then Exit Function
    targetColor = NormalizeHexColor(hexColor)
    startTime = modWinAPI.GetTickCount()
    
    Do
        curColor = modWinAPI.GetColorHexAt(x, y, m_inputTestDriver)
        
        ' 조건 충족 확인
        If isVanish Then
            If curColor <> targetColor Then Action_Wait = True: Exit Function
        Else
            If curColor = targetColor Then Action_Wait = True: Exit Function
        End If
        
        ' 타임아웃 (무인 환경: 모달 금지, 조용히 실패 반환 → 알림 여부는 엔진이 결정)
        If modWinAPI.ElapsedMilliseconds(startTime, modWinAPI.GetTickCount()) > timeLimit Then
            Action_Wait = False
            Exit Function
        End If
        
        ' 비상 탈출 (래치: 감지 시 실행 전체가 중단 경로로 수렴)
        If AbortPending() Then Action_Wait = False: Exit Function
        
        BetterWait DELAY_SHORT
    Loop

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_Wait", Err.Description
    Resume CleanUp
End Function

' [화면 검증] Assert (기대값과 다르면 실패 처리)
Public Function Action_Assert(x As Long, y As Long, condition As String) As Boolean
    On Error GoTo ErrHandler
    Dim targetColor As String, nowColor As String
    Dim isForbidden As Boolean
    
    targetColor = UCase$(Trim$(condition))
    isForbidden = (Left(targetColor, 1) = "!")
    If isForbidden Then targetColor = Mid(targetColor, 2)
    targetColor = NormalizeHexColor(targetColor)
    
    nowColor = modWinAPI.GetColorHexAt(x, y, m_inputTestDriver)
    
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

Private Function NormalizeHexColor(ByVal colorText As String) As String
    colorText = UCase$(Trim$(colorText))
    If Len(colorText) <> 6 Or colorText Like "*[!0-9A-F]*" Then _
        Err.Raise 5, "modActions.NormalizeHexColor", "Invalid pixel color: " & colorText
    NormalizeHexColor = colorText
End Function


'==============================================================================
' 5. Window & System Actions (윈도우 제어)
'==============================================================================

' [창 활성화] 제목의 일부만으로 창을 찾아 맨 앞으로 가져옴
Public Function Action_FocusWindow(ByVal partialTitle As String) As Boolean
    On Error GoTo ErrHandler
    If Len(Trim$(partialTitle)) = 0 Then Err.Raise 5, "modActions.Action_FocusWindow", "FOCUS title must not be empty."
    Dim hwnd As LongPtr
    Dim winTitle As String
    Dim lenTitle As Long

    If Not m_inputTestDriver Is Nothing Then
        Action_FocusWindow = modWinAPI.ActivateWindow(m_inputTestDriver.WindowHandle, m_inputTestDriver)
        Exit Function
    End If
    
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
                Action_FocusWindow = modWinAPI.ActivateWindow(hwnd)
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

' [알림 표시] 비차단형 알림. Action_Notify 단일 게이트로 위임 (항상 자동 종료 보장)
' @desc 메시지와 표시시간초 받아 비차단 알림 표시함
Public Sub Action_ShowNotification(ByVal msg As String, Optional ByVal durationSec As Long = 2)
    On Error GoTo ErrHandler
    Dim ms As Long
    ms = durationSec * 1000
    If ms <= 0 Then ms = 2000   ' 비차단 계약 유지: 0 이하도 자동 종료로 강제
    Action_Notify ms, MB_ICONINFO, msg

CleanUp:
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] Action_ShowNotification: " & Err.Description
    Resume CleanUp
End Sub

' [알림 단일 게이트] 모든 사용자 알림은 이 프로시저를 경유한다
' p1 (durationMs): 0이면 확인대기(MsgBox), 0보다 크면 자동종료(ms 단위)
' p2 (iconType): 아이콘 번호 (16, 32, 48, 64 등)
' p3 (msg): 출력할 메시지
' 무인 모드: 차단형 포함 전부 LogShim jsonl로 강등 (헤드리스 hang 차단)
' @desc 지속ms/아이콘/메시지 받아 알림 표시함, 무인 모드 시 로그로 강등
Public Sub Action_Notify(ByVal durationMs As Long, ByVal iconType As Long, ByVal msg As String)
    On Error GoTo ErrHandler
    If m_unattended Then
        LogShim.LogEvent "notify", Array("msg", "durationMs"), Array(msg, durationMs)
        GoTo CleanUp
    End If

    Dim flags As Long
    ' 기본 플래그: 항상 위 + 포커스
    flags = iconType + MB_TOPMOST + MB_SETFOREGROUND

    If durationMs <= 0 Then
        ' [차단형] 확인 누를 때까지 정지 (대화형 전용)
        MsgBox msg, flags, "RPA Alert"
    Else
        ' [비차단형] p1(ms) 만큼 대기 후 자동 종료
        modWinAPI.MessageBoxTimeout 0, msg, "RPA Notification", flags, 0, durationMs
    End If

CleanUp:
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "modActions.Action_Notify", Err.Description
    Resume CleanUp
End Sub
