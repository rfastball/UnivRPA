Attribute VB_Name = "modWinAPI"
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
Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
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

' 64비트 엑셀 크래시 방지를 위한 패딩 적용 구조체 (Union 대체)
Public Type GENERALINPUT
    dwType As Long
    ' 64비트 환경에서 구조체 정렬(Alignment)을 맞추기 위해 넉넉한 바이트 배열 사용
    xi(0 To 39) As Byte
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
Public Function GetColorHexAt(x As Long, y As Long) As String
    GetColorHexAt = FormatColorHex(GetColorAt(x, y))
End Function

Public Function GetColorAt(ByVal x As Long, ByVal y As Long) As Long
    Dim hdc As LongPtr
    
    hdc = GetDC(0)
    GetColorAt = GetPixel(hdc, x, y)
    ReleaseDC 0, hdc
End Function


' [Input] 키보드 상태 확인 (GetAsyncKeyState Wrapper)
Public Function IsKeyDown(ByVal vKey As Long) As Boolean
    IsKeyDown = (GetAsyncKeyState(vKey) And &H8000) <> 0
End Function

Public Function IsToggleKeyOn(ByVal vKey As Long) As Boolean
    IsToggleKeyOn = (GetKeyState(vKey) And 1) = 1
End Function

' [Clipboard] 텍스트 가져오기 (API 방식 - 안정적)
Public Function GetClipboardText() As String
    Dim hMem As LongPtr
    Dim lpData As LongPtr
    Dim nSize As Long
    Dim strBuffer As String
    
    If OpenClipboard(0) = 0 Then Exit Function
    
    hMem = GetClipboardData(CF_UNICODETEXT)
    If hMem <> 0 Then
        lpData = GlobalLock(hMem)
        If lpData <> 0 Then
            nSize = lstrlenW(lpData)
            strBuffer = String(nSize, vbNullChar)
            lstrcpyW StrPtr(strBuffer), lpData
            GlobalUnlock hMem
        End If
    End If
    
    CloseClipboard
    GetClipboardText = strBuffer
End Function

' [Clipboard] 텍스트 넣기 (API 방식 - DataObject 대체, 윈도우 10/11 호환성 해결)
Public Sub SetClipboardText(ByVal text As String)
    Dim hMem As LongPtr
    Dim lpData As LongPtr
    Dim nSize As Long
    
    If OpenClipboard(0) = 0 Then Exit Sub
    EmptyClipboard
    
    ' 유니코드 문자열 크기 계산 (Null 종료 문자 포함)
    nSize = LenB(text) + 2
    
    ' 전역 메모리 할당
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, nSize)
    If hMem <> 0 Then
        lpData = GlobalLock(hMem)
        If lpData <> 0 Then
            ' 메모리에 문자열 복사
            CopyMemory ByVal lpData, ByVal StrPtr(text), nSize
            GlobalUnlock hMem
            
            ' 클립보드에 데이터 설정
            SetClipboardData CF_UNICODETEXT, hMem
        End If
    End If
    
    CloseClipboard
End Sub

'==============================================================================
' 4. SendInput Technical Implementation (64-bit Padding)
'==============================================================================

' [Public Interface] 단일 키 클릭
Public Sub SendKey_Click(ByVal vKey As Integer)
    SendKey_Down vKey
    SendKey_Up vKey
End Sub

' [Public Interface] 키 누름 (Down)
Public Sub SendKey_Down(ByVal vKey As Integer)
    SendKeyboardInput vKey, 0, 0
End Sub

' [Public Interface] 키 뗌 (Up)
Public Sub SendKey_Up(ByVal vKey As Integer)
    SendKeyboardInput vKey, 0, KEYEVENTF_KEYUP
End Sub

' [Internal] 안전한 SendInput 래퍼 (64비트 패딩 자동 계산)
Public Sub SendKeyboardInput(ByVal wVk As Integer, ByVal wScan As Integer, ByVal dwFlags As Long)
    Dim inputSize As Long
    Dim buffer() As Byte
    Dim kbInput As KEYBDINPUT
    
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
    
    SendInput 1, buffer(0), inputSize
End Sub
