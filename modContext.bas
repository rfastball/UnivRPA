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
Private Memory(0 To 99) As String   ' 변수 저장 공간 (0~99번 슬롯)
Private CallStack As Collection     ' 서브루틴 복귀 주소 저장 (LIFO 구조)

' [2] 실행 상태 플래그
Private IsRunning As Boolean        ' 현재 매크로 실행 중 여부

' [3] 라벨 저장소
Private LabelMap As Collection      ' 라벨 저장소 (Key:Name -> Value:LineNumber)

' [4] 루프카운터 관리
Private LoopCounters As New Collection ' Key: "L10" (Line 10) -> Value: Current Count

'==============================================================================
' 1. Initialization & Reset (초기화)
'==============================================================================

' [초기화] 매크로 시작 시 반드시 호출
Public Sub InitContext()
    Dim i As Integer

    ' 1. 메모리 초기화
    For i = 0 To 99
        Memory(i) = ""
    Next i
    
    ' 2. 스택 초기화
    Set CallStack = New Collection
    
    ' 3. 상태 플래그 설정
    IsRunning = True
    
    ' 4. 상태바 초기화
    Application.StatusBar = "RPA Ready..."
    
    ' 5. 라벨 맵 초기화
    Set LabelMap = New Collection
End Sub

' [종료] 매크로 중단/완료 시 호출
Public Sub ResetContext()
    IsRunning = False
    Set CallStack = Nothing
    Application.StatusBar = False ' 엑셀 기본 상태로 복귀
End Sub

'==============================================================================
' 2. Memory Management (변수 제어)
'==============================================================================

' [변수 쓰기]
Public Sub SetVar(ByVal index As Integer, ByVal val As String)
    If index < 0 Or index > 99 Then
        LogError "메모리 인덱스 초과: " & index
        Exit Sub
    End If
    Memory(index) = val
    ' 디버깅용: 변수 변경 시 이미디에이트 창에 출력
    ' Debug.Print "[MEM] Var(" & index & ") = " & val
End Sub

' [변수 읽기]
Public Function GetVar(ByVal index As Integer) As String
    If index < 0 Or index > 99 Then
        LogError "메모리 인덱스 초과: " & index
        GetVar = ""
        Exit Function
    End If
    GetVar = Memory(index)
End Function

' [전체 메모리 덤프] 디버깅용
Public Sub DumpMemory()
    Dim i As Integer
    Debug.Print "=== Memory Dump ==="
    For i = 0 To 99
        If Memory(i) <> "" Then
            Debug.Print "Var(" & i & "): [" & Memory(i) & "]"
        End If
    Next i
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
' 6. [Loop Management] 완전 자동 카운터 로직
'--------------------------------------------------

' 루프 카운터 가져오기 (없으면 0 반환)
Public Function GetLoopCount(ByVal lineIdx As Long) As Long
    On Error Resume Next
    GetLoopCount = LoopCounters.Item("L" & lineIdx)
    If Err.Number <> 0 Then GetLoopCount = 0
    On Error GoTo 0
End Function

' 루프 카운터 설정/업데이트
Public Sub SetLoopCount(ByVal lineIdx As Long, ByVal val As Long)
    Dim sKey As String: sKey = "L" & lineIdx
    On Error Resume Next
    LoopCounters.Remove sKey ' 기존 값 제거
    On Error GoTo 0
    LoopCounters.Add val, sKey
End Sub

' 루프 종료 시 카운터 제거
Public Sub RemoveLoopCount(ByVal lineIdx As Long)
    On Error Resume Next
    LoopCounters.Remove "L" & lineIdx
    On Error GoTo 0
End Sub

' [modContext] 타입 오류 없는 안전한 로거
Public Sub WriteLog(ByVal v1 As Variant, ByVal v2 As Variant, ByVal v3 As Variant, ByVal v4 As Variant)
    Dim wsLog As Worksheet
    Dim NextRow As Long
    
    ' 1. 로그 시트 확보 (없으면 생성)

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("RPA_LOG")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsLog.Name = "RPA_LOG"
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
