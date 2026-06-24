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

' [3] 라벨 저장소
Private LabelMap As Collection      ' 라벨 저장소 (Key:Name -> Value:LineNumber)

' [4] 루프 런타임: 활성 루프 스택 ({$LOOP}/BREAK/CONTINUE 통합)
Private Type LoopFrame
    kind As String      ' "WHILE" | "REPEAT"
    headerLine As Long  ' 점프 복귀 대상 (헤더 줄)
    endLine As Long     ' 종료(ENDER) 줄. BREAK 대상 = endLine + 1
    iteration As Long   ' 1-base 현재 회차 -> {$LOOP}
    maxCount As Long    ' REPEAT 전용 최대 횟수
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

    ' 5. 라벨 맵 초기화
    Set LabelMap = New Collection

    ' 6. 루프 스택 초기화
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

' [전체 메모리 덤프] 디버깅용
Public Sub DumpMemory()
    Dim k As Variant
    Debug.Print "=== Memory Dump ==="
    For Each k In Memory.Keys
        Debug.Print "Var(" & k & "): [" & Memory(k) & "]"
    Next k
    Debug.Print "==================="
End Sub

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
        PopStack = CallStack.Item(CallStack.count)
        CallStack.Remove CallStack.count
    Else
        LogError "Call Stack Underflow: 돌아갈 곳이 없습니다. (RETURN 매칭 오류)"
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

' [에러 로그] 치명적 오류 발생 시 (메시지박스 + 중단 권장)
Public Sub LogError(ByVal msg As String)
    Debug.Print "[ERROR] " & msg
    modActions.Action_ShowNotification "오류: " & msg, 3
    ' 필요 시 여기에 에러 기록 시트 작성 로직 추가 가능
End Sub

' [실행 상태 확인] 외부(Engine)에서 루프 계속 여부 확인용
Public Function CheckRunningState() As Boolean
    CheckRunningState = IsRunning
End Function

' -------------------------------------------------------------------------
' 5. [Label Management] 라벨(핸들) 관리
' -------------------------------------------------------------------------

' [등록] 라벨 추가 (중복 시 에러 발생)
Public Sub AddLabel(ByVal sName As String, ByVal iLine As Long)
    Dim cleanName As String
    cleanName = UCase(Trim(sName))
    
    If cleanName = "" Then Exit Sub
    
    ' 중복 체크: 이미 있으면 에러 (Collection 특성상 키 중복 시 런타임 에러 발생하므로 선제적 방어는 선택사항이나, 명확한 에러 메시지를 위해 추가)
    On Error Resume Next
    Dim check As Variant
    check = LabelMap.Item(cleanName)
    If Err.Number = 0 Then
        MsgBox "오류: 중복된 라벨 이름이 있습니다 -> " & sName, vbCritical
        End ' 강제 종료
    End If
    On Error GoTo 0
    
    LabelMap.Add iLine, cleanName ' Key를 이름으로 저장
End Sub

' [검색] 라벨 이름으로 줄 번호 반환
Public Function GetLabel(ByVal sName As String) As Long
    Dim cleanName As String
    cleanName = UCase(Trim(sName))
    
    On Error Resume Next
    GetLabel = LabelMap.Item(cleanName)
    
    If Err.Number <> 0 Then
        ' 라벨을 못 찾았을 경우
        GetLabel = -1
    End If
    On Error GoTo 0
End Function

'--------------------------------------------------
' 6. [Loop Management] 활성 루프 스택
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

' [modContext] 타입 오류 없는 안전한 로거
Public Sub WriteLog(ByVal v1 As Variant, ByVal v2 As Variant, ByVal v3 As Variant, ByVal v4 As Variant)
    Dim wsLog As Worksheet
    Dim NextRow As Long
    
    ' 1. 로그 시트 확보 (없으면 생성)

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("RPA_LOG")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsLog.name = "RPA_LOG"
        wsLog.Range("A1:E1").Value = Array("Time", "Loop", "Line", "Command", "Message")
        wsLog.Range("A1:E1").Font.Bold = True
    End If
    On Error GoTo 0

    ' 2. 마지막 행 찾기 (속도를 위해 가급적 효율적으로)
    NextRow = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).Row + 1
    
    With wsLog
        .Cells(NextRow, 1).Value = Now
        ' 어떤 타입이 들어와도 & "" 가 문자로 안전하게 변환해줌
        .Cells(NextRow, 2).Value = v1 & ""
        .Cells(NextRow, 3).Value = v2 & ""
        .Cells(NextRow, 4).Value = v3 & ""
        .Cells(NextRow, 5).Value = v4 & ""
    End With
End Sub
