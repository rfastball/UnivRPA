VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsActionTestDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @dist exclude
' @desc 입력 호출 기록과 중단 및 API 실패 주입 담당
' @depends modActions, modEngine
Option Explicit

Public SleepTotal As Long
Public SleepCalls As Long
Public TickMilliseconds As Long
Public EscapeOnPoll As Long
Public PauseAtTick As Long
Public ResumeAtTick As Long
Private m_pauseState As Long
Private m_cursorX As Long, m_cursorY As Long
Public CursorShiftX As Long
Public MoveCalls As Long
Public MoveResult As Long
Public MouseEvents As String
Public KeyEvents As String
Public KeyCalls As Long
Public AbortOnSleep As Long
Public FailOnSleep As Long
Public FailOnKey As Long
Public ReentrySheet As Worksheet
Public ReentryStatus As String
Public AttemptReentry As Boolean
Public WindowHandle As LongPtr
Public ForegroundHandle As LongPtr
Public FocusResult As Long
Public FocusCalls As Long
Public PixelColor As Long
Public PixelReads As Long
Public ChangePixelOnRead As Long
Public ChangePixelAfterClick As Boolean
Public ClipboardFailure As String
Public ClipboardContents As String
Public ClipboardClosed As Boolean
Public ClipboardFreed As Boolean
Public ClipboardOwner As LongPtr
Public ClipboardHasText As Boolean
Public ClipboardUnlocked As Boolean
Private m_clipboardText As String
Public OpenSecurity As Long
Public OpenEvents As Boolean
Public FailOpen As Boolean

' @desc 호출 기록과 주입 조건 및 중단 래치를 초기화함
Public Sub Reset()
    SleepTotal = 0: SleepCalls = 0: MoveCalls = 0: KeyCalls = 0
    EscapeOnPoll = 0
    TickMilliseconds = 0: PauseAtTick = 0: ResumeAtTick = 0: m_pauseState = 0
    CursorShiftX = 0
    MoveResult = 1: FocusCalls = 0
    MouseEvents = "": KeyEvents = ""
    AbortOnSleep = 0: FailOnSleep = 0: FailOnKey = 0
    Set ReentrySheet = Nothing
    ReentryStatus = "": AttemptReentry = False
    WindowHandle = 123: ForegroundHandle = 123: FocusResult = 1
    PixelColor = &HFFFFFF
    PixelReads = 0: ChangePixelOnRead = 0
    ChangePixelAfterClick = False
    ClipboardFailure = "": ClipboardContents = "old": m_clipboardText = ""
    ClipboardClosed = False: ClipboardFreed = False: ClipboardOwner = 0
    ClipboardHasText = True: ClipboardUnlocked = False
    OpenSecurity = 0: OpenEvents = True: FailOpen = False
    modActions.ClearAbort
End Sub

' @desc 실제 대기 없이 요청 시간 기록하고 지정 호출에서 중단 또는 오류 주입함
Public Sub SleepMilliseconds(ByVal milliseconds As Long)
    SleepTotal = SleepTotal + milliseconds
    TickMilliseconds = TickMilliseconds + milliseconds
    SleepCalls = SleepCalls + 1
    If AttemptReentry Then
        AttemptReentry = False
        ReentryStatus = modEngine.RunScriptSheet(ReentrySheet, True)
    End If
    If SleepCalls = AbortOnSleep Then modActions.RequestAbort
    If SleepCalls = FailOnSleep Then Err.Raise 5, "clsActionTestDriver", "Injected wait failure."
End Sub

Public Function EscapeKeyDown() As Boolean
    If EscapeOnPoll <= 0 Then Exit Function
    EscapeOnPoll = EscapeOnPoll - 1
    EscapeKeyDown = (EscapeOnPoll = 0)
End Function

' @desc 지정한 가상 시각에 Pause와 재개 키 누름을 각 1회 반환
Public Function PauseKeyDown() As Boolean
    If m_pauseState = 0 And PauseAtTick > 0 And TickMilliseconds >= PauseAtTick Then
        m_pauseState = 1
        PauseKeyDown = True
    ElseIf m_pauseState = 1 And TickMilliseconds >= ResumeAtTick Then
        m_pauseState = 2
        PauseKeyDown = True
    End If
End Function

' @desc 커서 이동 요청 횟수를 기록하고 지정한 성공 여부 반환
Public Function MoveCursor(ByVal x As Long, ByVal y As Long) As Long
    MoveCalls = MoveCalls + 1
    m_cursorX = x: m_cursorY = y
    MoveCursor = MoveResult
End Function

Public Function ReadCursorPos(ByRef x As Long, ByRef y As Long) As Long
    x = m_cursorX + CursorShiftX
    y = m_cursorY
    ReadCursorPos = 1
End Function

' @desc 마우스 버튼 플래그 순서를 기록함
Public Sub MouseButton(ByVal flags As Long)
    MouseEvents = MouseEvents & flags & ";"
End Sub

' @desc 키 요청 기록 후 성공 개수 반환, 지정 호출에서 0 반환
Public Function KeyboardInput(ByVal vk As Integer, ByVal scan As Integer, ByVal flags As Long) As Long
    KeyCalls = KeyCalls + 1
    KeyEvents = KeyEvents & vk & ":" & flags & ";"
    If KeyCalls <> FailOnKey Then KeyboardInput = 1
End Function

' @desc 전면화 API의 성공 또는 실패 결과를 반환함
Public Function ActivateWindow(ByVal hwnd As LongPtr) As Long
    FocusCalls = FocusCalls + 1
    ActivateWindow = FocusResult
End Function

' @desc 지정한 픽셀 판독값 또는 실패값을 반환함
Public Function ReadPixel(ByVal x As Long, ByVal y As Long) As Long
    PixelReads = PixelReads + 1
    ReadPixel = PixelColor
    If ChangePixelAfterClick And Len(MouseEvents) > 0 Then ReadPixel = PixelColor Xor &HFFFFFF
    If ChangePixelOnRead > 0 And PixelReads >= ChangePixelOnRead Then ReadPixel = PixelColor Xor &HFFFFFF
End Function

' @desc 가상 클립보드의 텍스트 형식 존재 여부 반환
Public Function IsClipboardFormatAvailable(ByVal formatId As Long) As Long
    If ClipboardHasText Then IsClipboardFormatAvailable = 1
End Function

' @desc 읽기용 가상 메모리 핸들 반환, 주입 실패 시 0 반환
Public Function GetClipboardData(ByVal formatId As Long) As LongPtr
    If ClipboardFailure <> "get" Then GetClipboardData = 456
End Function

' @desc 읽기용 가상 포인터 반환, 주입 실패 시 0 반환
Public Function GlobalLock(ByVal hMem As LongPtr) As LongPtr
    If ClipboardFailure <> "lock" Then GlobalLock = 789
End Function

' @desc 가상 클립보드 텍스트 반환, 주입 조건이면 오류 발생
Public Function ReadClipboardText(ByVal pointer As LongPtr) As String
    If ClipboardFailure = "read" Then Err.Raise 5, "clsActionTestDriver", "Injected clipboard read failure."
    ReadClipboardText = ClipboardContents
End Function

' @desc 읽기 메모리 잠금 해제 시도를 기록함
Public Sub GlobalUnlock(ByVal hMem As LongPtr)
    ClipboardUnlocked = True
End Sub

' @desc 실제 클립보드를 변경하지 않고 쓰기 메모리를 준비함
Public Function AllocateClipboardText(ByVal text As String) As LongPtr
    If ClipboardFailure = "allocate" Then Err.Raise 7, "clsActionTestDriver", "Injected allocation failure."
    m_clipboardText = text
    AllocateClipboardText = 456
End Function

' @desc 클립보드 소유 창을 기록하고 열기 결과를 반환함
Public Function OpenClipboard(ByVal hwnd As LongPtr) As Long
    ClipboardOwner = hwnd
    If ClipboardFailure <> "open" Then OpenClipboard = 1
End Function

' @desc 가상 클립보드를 비우고 성공 여부를 반환함
Public Function EmptyClipboard() As Long
    If ClipboardFailure = "empty" Then Exit Function
    ClipboardContents = ""
    EmptyClipboard = 1
End Function

' @desc 가상 클립보드에 텍스트를 기록하고 성공 핸들을 반환함
Public Function SetClipboardData(ByVal formatId As Long, ByVal hMem As LongPtr) As LongPtr
    If ClipboardFailure = "set" Then Exit Function
    ClipboardContents = m_clipboardText
    SetClipboardData = hMem
End Function

' @desc 클립보드 닫기 시도를 기록함
Public Sub CloseClipboard()
    ClipboardClosed = True
End Sub

' @desc 이전하지 못한 메모리 해제를 기록함
Public Sub FreeClipboardMemory(ByVal hMem As LongPtr)
    ClipboardFreed = True
End Sub

' @desc 외부 파일 열기 시 보안 상태를 기록하고 성공 또는 실패를 주입함
Public Function OpenWorkbook(ByVal filePath As String) As Workbook
    OpenSecurity = Application.AutomationSecurity
    OpenEvents = Application.EnableEvents
    If FailOpen Then Err.Raise 53, "clsActionTestDriver", "Injected workbook open failure."
    Set OpenWorkbook = ThisWorkbook
End Function
