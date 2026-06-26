Attribute VB_Name = "modContext"
Option Explicit

'==============================================================================
' [Layer 2] modContext: State & Memory Management
'==============================================================================
' 역할: 변수(Memory), 콜스택(CallStack), 실행 상태(State) 관리
' 원칙:
' 1. 모든 변수는 Private으로 숨기고, 함수(Get/Set)를 통해서만 접근
' 2. 스택 오버플로우/언더플로우 방지 로직 포함
' 3. 로그 및 상태바 업데이트 중앙화
'==============================================================================

' [1] 핵심 저장소
Private Memory As Object            ' 변수 저장 공간 (Scripting.Dictionary, 자유 명명 {name})
Private CallStack As Collection     ' 서브루틴 복귀 주소 저장 (LIFO 구조)

' [2] 실행 상태 플래그
Private IsRunning As Boolean        ' 현재 매크로 실행 중 여부

' [3] 루프 런타임: 활성 루프 스택 ({$LOOP}/BREAK/CONTINUE 통합)
Private Type LoopFrame
    kind As String      ' "WHILE" | "REPEAT"
    headerLine As Long  ' 점프 복귀 대상 (헤더 줄)
    endLine As Long     ' 종료(ENDER) 줄. BREAK 대상 = endLine + 1
    iteration As Long   ' 1-base 현재 회차 -> {$LOOP}
    maxCount As Long    ' REPEAT 전용 최대 횟수 (0=WHILE 또는 무한 REPEAT)
End Type
Private LoopStack() As LoopFrame
Private LoopTop As Long             ' 0 = 빈 스택, n = 활성 프레임 개수

'==============================================================================
' 1. Initialization & Reset (초기화)
'==============================================================================

' [초기화] 매크로 시작 시 반드시 호출
Public Sub InitContext()
    ' 1. 메모리 초기화 (Dictionary, 대소문자 무시)
    Set Memory = CreateObject("Scripting.Dictionary")
    Memory.CompareMode = vbTextCompare

    ' 2. 스택 초기화
    Set CallStack = New Collection

    ' 3. 상태 플래그 설정
    IsRunning = True

    ' 4. 상태바 초기화
    Application.StatusBar = "RPA Ready..."

    ' 5. 루프 스택 초기화
    LoopReset
End Sub

' [종료] 매크로 중단/완료 시 호출
Public Sub ResetContext()
    IsRunning = False
    Set CallStack = Nothing
    Set Memory = Nothing
    LoopReset
    Application.StatusBar = False ' 엑셀 기본 상태로 복귀
End Sub

'==============================================================================
' 2. Memory Management (변수 제어)
'==============================================================================

' [변수 쓰기] 자유 명명 키. 좌변=주소, 우변=값
Public Sub SetVar(ByVal key As String, ByVal val As String)
    Memory(Trim(key)) = val   ' 존재하면 갱신, 없으면 자동 추가
End Sub

' [변수 읽기] 미정의 키는 빈 문자열({$BLANK}) 반환 (오류 아님)
Public Function GetVar(ByVal key As String) As String
    Dim k As String: k = Trim(key)
    If Memory.exists(k) Then GetVar = Memory(k) Else GetVar = ""
End Function

'==============================================================================
' 3. Call Stack Management (서브루틴 제어)
'==============================================================================

' [PUSH] 서브루틴 진입 시 복귀할 줄 번호 저장
Public Sub PushStack(ByVal returnRow As Long)
    CallStack.Add returnRow
End Sub

' [POP] 서브루틴 종료(RETURN) 시 복귀할 줄 번호 반환
Public Function PopStack() As Long
    If CallStack.count > 0 Then
        ' 컬렉션의 마지막 아이템을 가져옴 (LIFO)
        PopStack = CallStack.item(CallStack.count)
        CallStack.Remove CallStack.count
    Else
        LogStatus "빈 콜스택 RETURN: 복귀할 CALL이 없습니다."
        PopStack = 0
    End If
End Function

' [PEEK] 현재 스택 깊이 확인 (무한 재귀 방지용)
Public Function GetStackDepth() As Long
    If CallStack Is Nothing Then
        GetStackDepth = 0
    Else
        GetStackDepth = CallStack.count
    End If
End Function

'==============================================================================
' 4. Logging & Status (로그 및 알림)
'==============================================================================

' [상태 표시] 엑셀 하단 상태바에 메시지 표시
Public Sub LogStatus(ByVal msg As String)
    Application.StatusBar = "[RPA] " & msg
    DoEvents
End Sub

' [실행 상태 확인] 외부(Engine)에서 루프 계속 여부 확인용
Public Function CheckRunningState() As Boolean
    CheckRunningState = IsRunning
End Function

'--------------------------------------------------
' 5. [Loop Management] 활성 루프 스택
'   {$LOOP}(가장 안쪽 회차) / BREAK / CONTINUE 를 한 묶음으로 해결
'--------------------------------------------------

' [초기화] 스택 비우기
Public Sub LoopReset()
    Erase LoopStack
    LoopTop = 0
End Sub

' [깊이] 현재 활성 루프 개수
Public Function LoopDepth() As Long
    LoopDepth = LoopTop
End Function

' [진입] 새 루프 프레임 PUSH (iteration = 1 로 시작)
Public Sub LoopPush(ByVal kind As String, ByVal headerLine As Long, ByVal endLine As Long, ByVal maxCount As Long)
    LoopTop = LoopTop + 1
    ReDim Preserve LoopStack(1 To LoopTop)
    LoopStack(LoopTop).kind = kind
    LoopStack(LoopTop).headerLine = headerLine
    LoopStack(LoopTop).endLine = endLine
    LoopStack(LoopTop).iteration = 1
    LoopStack(LoopTop).maxCount = maxCount
End Sub

' [이탈] 최상위 루프 프레임 POP
Public Sub LoopPop()
    If LoopTop > 0 Then
        LoopTop = LoopTop - 1
        If LoopTop > 0 Then
            ReDim Preserve LoopStack(1 To LoopTop)
        Else
            Erase LoopStack
        End If
    End If
End Sub

' [{$LOOP}] 가장 안쪽 루프의 현재 회차 (루프 밖이면 0)
Public Function GetActiveLoopIteration() As Long
    If LoopTop > 0 Then GetActiveLoopIteration = LoopStack(LoopTop).iteration Else GetActiveLoopIteration = 0
End Function

' [회차+1] 최상위 프레임의 회차 증가
Public Sub LoopIncrement()
    If LoopTop > 0 Then LoopStack(LoopTop).iteration = LoopStack(LoopTop).iteration + 1
End Sub

' [접근자] 최상위 프레임 필드 읽기 (빈 스택이면 기본값)
Public Function LoopTopKind() As String
    If LoopTop > 0 Then LoopTopKind = LoopStack(LoopTop).kind Else LoopTopKind = ""
End Function
Public Function LoopTopHeader() As Long
    If LoopTop > 0 Then LoopTopHeader = LoopStack(LoopTop).headerLine Else LoopTopHeader = 0
End Function
Public Function LoopTopEnd() As Long
    If LoopTop > 0 Then LoopTopEnd = LoopStack(LoopTop).endLine Else LoopTopEnd = 0
End Function
Public Function LoopTopMax() As Long
    If LoopTop > 0 Then LoopTopMax = LoopStack(LoopTop).maxCount Else LoopTopMax = 0
End Function
