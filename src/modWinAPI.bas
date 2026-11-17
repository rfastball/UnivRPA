Attribute VB_Name = "modWinAPI"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

'==============================================================================
' [Layer 4] modWinAPI: Windows API & Technical Wrappers
'==============================================================================
' 역할: OS 레벨의 하드웨어 제어, 메모리 관리, 64비트 호환성 처리
' 원칙: 상태(변수)를 가지지 않으며, 비즈니스 로직(Loop, Wait)을 포함하지 않음.
'==============================================================================

'==============================================================================
' 1. Windows API 선언부 (Public, 64비트 호환 PtrSafe 적용)
'==============================================================================

' [입력 제어]
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function SendInput Lib "user32" (ByVal nInputs As Long, ByRef pInputs As Any, ByVal cbSize As Long) As Long

' [화면 인식]
Public Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetProcessDPIAware Lib "user32" () As Long
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

' [창 제어]
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' [메시지 박스 (Timer)]
Public Declare PtrSafe Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wLanguageId As Integer, ByVal dwMilliseconds As Long) As Long

' [클립보드 제어 (안정성 강화)]
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
Public Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

' [키보드 토글 상태 확인용 API 추가]
Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'==============================================================================
' 2. 구조체 및 상수 정의 (Public)
'==============================================================================

' [좌표 구조체]
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' [SendInput 관련 구조체]
Public Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

' [상수 모음]
Public Const INPUT_KEYBOARD = 1
Public Const INPUT_MOUSE = 0

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_UNICODE = &H4

' 가상 키코드 (Virtual Keys)
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public Const VK_MENU = &H12 ' ALT key
Public Const VK_PAUSE = &H13
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_RETURN = &HD
Public Const VK_PAUSE_RESUME = &H13
Public Const VK_CAPITAL = &H14  ' Caps Lock 키 코드

' 마우스 이벤트 플래그
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_ABSOLUTE = &H8000

' 윈도우 상태 상수
Public Const GW_HWNDNEXT = 2
Public Const SW_RESTORE = 9

' 작업영역/DPI/창 배치 상수 (폼 배치 조합은 modHelper가 소유 - L4는 선언만)
Public Const SPI_GETWORKAREA As Long = 48
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOACTIVATE As Long = &H10

' 메시지박스 옵션
Public Const MB_ICONINFO = &H40
Public Const MB_TOPMOST = &H40000
Public Const MB_SETFOREGROUND = &H10000

' 클립보드 및 메모리 상수
Public Const CF_UNICODETEXT = 13
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40

'화면 해상도
Public Const SM_CXSCREEN = 0        ' 주 모니터 가로 해상도
Public Const SM_CYSCREEN = 1        ' 주 모니터 세로 해상도
Public Const SM_CXVIRTUALSCREEN = 78 ' 전체 모니터(가상) 가로 해상도
Public Const SM_CYVIRTUALSCREEN = 79 ' 전체 모니터(가상) 세로 해상도

'==============================================================================
' 3. Technical Wrappers (기술적 복잡성 은닉)
'==============================================================================

' [Vision] 색상 Long 값을 6자리 Hex String으로 포맷 (관용구의 단일 소유)
Public Function FormatColorHex(ByVal colorValue As Long) As String
    FormatColorHex = Right("000000" & hex(colorValue), 6)
End Function

' [Vision] 화면의 특정 좌표 색상을 Hex String으로 반환 (DC 관리 포함)
Public Function GetColorHexAt(x As Long, y As Long, Optional ByVal inputTestDriver As Object = Nothing) As String
    GetColorHexAt = FormatColorHex(GetColorAt(x, y, inputTestDriver))
End Function

Public Function GetColorAt(ByVal x As Long, ByVal y As Long, Optional ByVal inputTestDriver As Object = Nothing) As Long
    Dim hdc As LongPtr
    Dim colorValue As Long
    If inputTestDriver Is Nothing Then
        hdc = GetDC(0)
        If hdc = 0 Then Err.Raise 5, "modWinAPI.GetColorAt", "GetDC failed."
        colorValue = GetPixel(hdc, x, y)
        ReleaseDC 0, hdc
    Else
        colorValue = inputTestDriver.ReadPixel(x, y)
    End If
    If colorValue = -1 Then Err.Raise 5, "modWinAPI.GetColorAt", "GetPixel failed at (" & x & "," & y & ")."
    GetColorAt = colorValue
End Function

' @desc 창 핸들을 받아 전면화 후 실제 포커스 확인, 실패 시 False 반환
Public Function ActivateWindow(ByVal hwnd As LongPtr, Optional ByVal inputTestDriver As Object = Nothing) As Boolean
    If hwnd = 0 Then Exit Function
    Dim activated As Long, foreground As LongPtr
    If inputTestDriver Is Nothing Then
        If IsIconic(hwnd) Then ShowWindow hwnd, SW_RESTORE
        activated = SetForegroundWindow(hwnd)
        foreground = GetForegroundWindow()
    Else
        activated = inputTestDriver.ActivateWindow(hwnd)
        foreground = inputTestDriver.ForegroundHandle
    End If
    ActivateWindow = (activated <> 0 And foreground = hwnd)
End Function

' [Input] 키보드 상태 확인 (GetAsyncKeyState Wrapper)
Public Function IsKeyDown(ByVal vKey As Long) As Boolean
    IsKeyDown = (GetAsyncKeyState(vKey) And &H8000) <> 0
End Function

Public Function IsToggleKeyOn(ByVal vKey As Long) As Boolean
    IsToggleKeyOn = (GetKeyState(vKey) And 1) = 1
End Function

' [Clipboard] 텍스트 가져오기 (API 방식 - 안정적)
Public Function GetClipboardText(Optional ByVal clipboardApi As Object = Nothing) As String
    Dim hMem As LongPtr, lpData As LongPtr
    Dim opened As Boolean, result As Long, nSize As Long
    Dim errorNumber As Long, errorText As String, strBuffer As String
    On Error GoTo Failed
    If clipboardApi Is Nothing Then
        result = OpenClipboard(0)
    Else
        result = clipboardApi.OpenClipboard(0)
    End If
    If result = 0 Then Err.Raise 5, "modWinAPI.GetClipboardText", "OpenClipboard failed."
    opened = True
    If clipboardApi Is Nothing Then
        result = IsClipboardFormatAvailable(CF_UNICODETEXT)
    Else
        result = clipboardApi.IsClipboardFormatAvailable(CF_UNICODETEXT)
    End If
    ' No text format is a valid empty value; a failed read is not.
    If result = 0 Then GoTo CleanUp
    If clipboardApi Is Nothing Then
        hMem = GetClipboardData(CF_UNICODETEXT)
    Else
        hMem = clipboardApi.GetClipboardData(CF_UNICODETEXT)
    End If
    If hMem = 0 Then Err.Raise 5, "modWinAPI.GetClipboardText", "GetClipboardData failed."
    If clipboardApi Is Nothing Then
        lpData = GlobalLock(hMem)
    Else
        lpData = clipboardApi.GlobalLock(hMem)
    End If
    If lpData = 0 Then Err.Raise 5, "modWinAPI.GetClipboardText", "GlobalLock failed."
    If clipboardApi Is Nothing Then
        nSize = lstrlenW(lpData)
        strBuffer = String$(nSize, vbNullChar)
        If nSize > 0 Then CopyMemory ByVal StrPtr(strBuffer), ByVal lpData, LenB(strBuffer)
    Else
        strBuffer = clipboardApi.ReadClipboardText(lpData)
    End If
    GetClipboardText = strBuffer
CleanUp:
    On Error Resume Next
    If clipboardApi Is Nothing Then
        If lpData <> 0 Then GlobalUnlock hMem
        If opened Then CloseClipboard
    Else
        If lpData <> 0 Then clipboardApi.GlobalUnlock hMem
        If opened Then clipboardApi.CloseClipboard
    End If
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "modWinAPI.GetClipboardText", errorText
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume CleanUp
End Function

' [Clipboard] 텍스트 넣기 (API 방식 - DataObject 대체, 윈도우 10/11 호환성 해결)
Public Sub SetClipboardText(ByVal text As String, Optional ByVal clipboardApi As Object = Nothing)
    Dim hMem As LongPtr, result As LongPtr, opened As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed

    ' Prepare memory before clearing the user's clipboard.
    If clipboardApi Is Nothing Then
        hMem = AllocateClipboardText(text)
        result = OpenClipboard(Application.hwnd)
    Else
        hMem = clipboardApi.AllocateClipboardText(text)
        result = clipboardApi.OpenClipboard(Application.hwnd)
    End If
    If result = 0 Then Err.Raise 5, "modWinAPI.SetClipboardText", "OpenClipboard failed."
    opened = True
    If clipboardApi Is Nothing Then
        result = EmptyClipboard()
    Else
        result = clipboardApi.EmptyClipboard()
    End If
    If result = 0 Then Err.Raise 5, "modWinAPI.SetClipboardText", "EmptyClipboard failed."
    If clipboardApi Is Nothing Then
        result = SetClipboardData(CF_UNICODETEXT, hMem)
    Else
        result = clipboardApi.SetClipboardData(CF_UNICODETEXT, hMem)
    End If
    If result = 0 Then Err.Raise 5, "modWinAPI.SetClipboardText", "SetClipboardData failed."
    hMem = 0 ' Ownership passes to Windows only after a successful write.

CleanUp:
    On Error Resume Next
    If clipboardApi Is Nothing Then
        If opened Then CloseClipboard
        If hMem <> 0 Then GlobalFree hMem
    Else
        If opened Then clipboardApi.CloseClipboard
        If hMem <> 0 Then clipboardApi.FreeClipboardMemory hMem
    End If
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "modWinAPI.SetClipboardText", errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume CleanUp
End Sub

Private Function AllocateClipboardText(ByVal text As String) As LongPtr
    Dim hMem As LongPtr, lpData As LongPtr
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(text) + 2)
    If hMem = 0 Then Err.Raise 7, "modWinAPI.SetClipboardText", "GlobalAlloc failed."
    lpData = GlobalLock(hMem)
    If lpData = 0 Then
        GlobalFree hMem
        Err.Raise 5, "modWinAPI.SetClipboardText", "GlobalLock failed."
    End If
    If LenB(text) > 0 Then CopyMemory ByVal lpData, ByVal StrPtr(text), LenB(text)
    GlobalUnlock hMem
    AllocateClipboardText = hMem
End Function
'==============================================================================
' 4. SendInput Technical Implementation (64-bit Padding)
'==============================================================================

' [Internal] 안전한 SendInput 래퍼 (64비트 패딩 자동 계산)
Public Sub SendKeyboardInput(ByVal wVk As Integer, ByVal wScan As Integer, ByVal dwFlags As Long, Optional ByVal inputTestDriver As Object = Nothing)
    Dim sent As Long
    Dim inputSize As Long
    Dim buffer() As Byte
    Dim kbInput As KEYBDINPUT

    If inputTestDriver Is Nothing Then
        ' 데이터 준비
        kbInput.wVk = wVk
        kbInput.wScan = wScan
        kbInput.dwFlags = dwFlags
        kbInput.time = 0
        kbInput.dwExtraInfo = 0

        ' 아키텍처별 메모리 구조 결정
        #If Win64 Then
            ' [64비트 환경] Type(4) + Padding(4) + KEYBDINPUT(24) + Padding(8) = 40 bytes
            inputSize = 40
            ReDim buffer(0 To 39)

            Dim dwType As Long: dwType = INPUT_KEYBOARD
            CopyMemory buffer(0), dwType, 4

            ' 64비트 정렬을 위한 Offset 8 위치에 데이터 복사
            CopyMemory buffer(8), kbInput, LenB(kbInput)

        #Else
            ' [32비트 환경] Type(4) + KEYBDINPUT(20) = 28 bytes
            inputSize = 28
            ReDim buffer(0 To 27)

            Dim dwType32 As Long: dwType32 = INPUT_KEYBOARD
            CopyMemory buffer(0), dwType32, 4

            CopyMemory buffer(4), kbInput, LenB(kbInput)
        #End If

        sent = SendInput(1, buffer(0), inputSize)
    Else
        sent = inputTestDriver.KeyboardInput(wVk, wScan, dwFlags)
    End If
    If sent <> 1 Then Err.Raise vbObjectError + 2101, "modWinAPI.SendKeyboardInput", "SendInput failed: inserted " & sent & " of 1 event."
End Sub

' @desc 두 32비트 tick 받아 경과ms 반환, 한 주기(49.7일) 이내 측정용
Public Function ElapsedMilliseconds(ByVal startTick As Long, ByVal endTick As Long) As Double
    Dim startValue As Double, endValue As Double
    startValue = CDbl(startTick): endValue = CDbl(endTick)
    If startValue < 0 Then startValue = startValue + 4294967296#
    If endValue < 0 Then endValue = endValue + 4294967296#
    ElapsedMilliseconds = endValue - startValue
    If ElapsedMilliseconds < 0 Then ElapsedMilliseconds = ElapsedMilliseconds + 4294967296#
End Function
