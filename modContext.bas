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
    kind As String      ' "WHILE" | "REPEAT" | "FOREACH"
    headerLine As Long  ' 점프 복귀 대상 (헤더 줄)
    endLine As Long     ' 종료(ENDER) 줄. BREAK 대상 = endLine + 1
    iteration As Long   ' 1-base 현재 회차 -> {$LOOP}
    maxCount As Long    ' REPEAT 전용 최대 횟수 (0=WHILE 또는 무한 REPEAT)
    dataRow As Long     ' FOREACH 전용: 현재 테이블 내부 행 인덱스
End Type
Private LoopStack() As LoopFrame
Private LoopTables() As Object      ' LoopStack과 같은 인덱스의 FOREACH 테이블 스냅샷
Private LoopTop As Long             ' 0 = 빈 스택, n = 활성 프레임 개수

' [4] IF 런타임: 이미 선택된 IF/ELSE_IF 분기 추적
Private Type IfBranchFrame
    headerLine As Long
    endIfLine As Long
End Type
Private IfBranchStack() As IfBranchFrame
Private IfBranchTop As Long

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

    ' 5. 런타임 스택 초기화
    LoopReset
    IfBranchReset
End Sub

' [종료] 매크로 중단/완료 시 호출
Public Sub ResetContext()
    IsRunning = False
    Set CallStack = Nothing
    Set Memory = Nothing
    LoopReset
    IfBranchReset
    Application.StatusBar = False ' 엑셀 기본 상태로 복귀
End Sub

'==============================================================================
' 2. Memory Management (변수 제어)
'==============================================================================

' [변수 쓰기] 자유 명명 키. 좌변=주소, 우변=값
Public Sub SetVar(ByVal key As String, ByVal val As String)
    Memory(Trim(key)) = val   ' 존재하면 갱신, 없으면 자동 추가
End Sub

' [변수 읽기] 호환 조회용. strict 실행 경로는 TryGetVar로 정의 여부를 확인한다.
Public Function GetVar(ByVal key As String) As String
    Dim k As String: k = Trim(key)
    If Memory.exists(k) Then GetVar = Memory(k) Else GetVar = ""
End Function

' [변수 정의 여부] 빈 문자열 값도 정의된 변수로 취급
Public Function HasVar(ByVal key As String) As Boolean
    Dim k As String: k = Trim(key)
    If Memory Is Nothing Then Exit Function
    HasVar = Memory.exists(k)
End Function

' [변수 안전 읽기] 정의 여부와 값을 함께 반환
Public Function TryGetVar(ByVal key As String, ByRef value As String) As Boolean
    Dim k As String: k = Trim(key)
    If Memory Is Nothing Then Exit Function
    If Not Memory.exists(k) Then Exit Function

    value = CStr(Memory(k))
    TryGetVar = True
End Function

' [정의된 변수명 목록] 추천/진단용
Public Function DefinedVarNames() As Variant
    If Memory Is Nothing Then
        DefinedVarNames = Array()
    Else
        DefinedVarNames = Memory.keys
    End If
End Function

'==============================================================================
' 3. Call Stack Management (서브루틴 제어)
'==============================================================================

' [PUSH] 서브루틴 진입 시 복귀할 줄 번호 저장
Public Sub PushStack(ByVal returnRow As Long, Optional ByVal loopDepthAtCall As Long = -1)
    If loopDepthAtCall < 0 Then loopDepthAtCall = LoopDepth()
    CallStack.Add Array(returnRow, loopDepthAtCall)
End Sub

' [POP] 서브루틴 종료(RETURN) 시 복귀할 줄 번호 반환
Public Function PopStack(Optional ByRef loopDepthRestored As Long = -1) As Long
    If CallStack.count > 0 Then
        Dim frame As Variant
        frame = CallStack.item(CallStack.count)
        CallStack.Remove CallStack.count

        Dim returnRow As Long
        Dim expectedLoopDepth As Long
        returnRow = CLng(frame(0))
        expectedLoopDepth = CLng(frame(1))

        If LoopDepth() < expectedLoopDepth Then
            LogStatus "RETURN 루프 깊이 불일치: 호출 시점보다 루프 스택이 얕습니다."
            PopStack = 0
            Exit Function
        End If

        Do While LoopDepth() > expectedLoopDepth
            LoopPop "return_restore"
        Loop

        loopDepthRestored = LoopDepth()
        PopStack = returnRow
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
    Erase LoopTables
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
    ReDim Preserve LoopTables(1 To LoopTop)
    LoopStack(LoopTop).kind = kind
    LoopStack(LoopTop).headerLine = headerLine
    LoopStack(LoopTop).endLine = endLine
    LoopStack(LoopTop).iteration = 1
    LoopStack(LoopTop).maxCount = maxCount
    Set LoopTables(LoopTop) = Nothing
End Sub

' [이탈] 최상위 루프 프레임 POP
Public Sub LoopPop(Optional ByVal reason As String = "end")
    If LoopTop > 0 Then
        LogShim.LogEvent "loop.pop", _
            Array("line", "kind", "iteration", "reason"), _
            Array(LoopStack(LoopTop).headerLine, LoopStack(LoopTop).kind, LoopStack(LoopTop).iteration, reason)
        Set LoopTables(LoopTop) = Nothing
        LoopTop = LoopTop - 1
        If LoopTop > 0 Then
            ReDim Preserve LoopStack(1 To LoopTop)
            ReDim Preserve LoopTables(1 To LoopTop)
        Else
            Erase LoopStack
            Erase LoopTables
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

Public Sub LoopPushForeach(ByVal headerLine As Long, ByVal endLine As Long, ByVal tbl As Object)
    If tbl Is Nothing Then Err.Raise 5, "modContext.LoopPushForeach", "FOREACH table is required."

    LoopPush "FOREACH", headerLine, endLine, 0
    Set LoopTables(LoopTop) = tbl
    LoopStack(LoopTop).dataRow = tbl.FirstDataRow
End Sub

Public Function ForeachAdvance() As Boolean
    If LoopTop = 0 Then Exit Function
    If LoopStack(LoopTop).kind <> "FOREACH" Then Exit Function
    If LoopTables(LoopTop) Is Nothing Then Exit Function

    If LoopStack(LoopTop).dataRow < LoopTables(LoopTop).RowCount Then
        LoopStack(LoopTop).dataRow = LoopStack(LoopTop).dataRow + 1
        LoopStack(LoopTop).iteration = LoopStack(LoopTop).iteration + 1
        ForeachAdvance = True
    End If
End Function

Public Function GetActiveForeachOrdinal() As Long
    Dim i As Long
    For i = LoopTop To 1 Step -1
        If LoopStack(i).kind = "FOREACH" Then
            GetActiveForeachOrdinal = LoopStack(i).iteration
            Exit Function
        End If
    Next i
End Function

Public Function ActiveForeachRow() As Long
    Dim i As Long
    For i = LoopTop To 1 Step -1
        If LoopStack(i).kind = "FOREACH" Then
            ActiveForeachRow = LoopStack(i).dataRow
            Exit Function
        End If
    Next i
End Function

Public Function ActiveForeachTable() As Object
    Dim i As Long
    For i = LoopTop To 1 Step -1
        If LoopStack(i).kind = "FOREACH" Then
            Set ActiveForeachTable = LoopTables(i)
            Exit Function
        End If
    Next i
End Function

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

Public Function ReconcileLoopStackForJump(ByVal targetLine As Long, ByVal targetStaticLoopDepth As Long) As Long
    Dim keepDepth As Long
    Dim i As Long

    keepDepth = 0
    For i = LoopTop To 1 Step -1
        If targetLine >= LoopStack(i).headerLine And targetLine <= LoopStack(i).endLine Then
            keepDepth = i
            Exit For
        End If
    Next i

    If keepDepth = 0 And targetStaticLoopDepth > 0 Then
        ReconcileLoopStackForJump = -1
        Exit Function
    End If

    Do While LoopTop > keepDepth
        LoopPop "goto_reconcile"
        ReconcileLoopStackForJump = ReconcileLoopStackForJump + 1
    Loop
End Function

'--------------------------------------------------
' 6. [IF Branch Management] 선택된 IF/ELSE_IF 분기 스택
'--------------------------------------------------

Public Sub IfBranchReset()
    Erase IfBranchStack
    IfBranchTop = 0
End Sub

Public Function IfBranchDepth() As Long
    IfBranchDepth = IfBranchTop
End Function

Public Sub IfBranchPush(ByVal headerLine As Long, ByVal endIfLine As Long)
    If endIfLine <= 0 Then Exit Sub

    IfBranchTop = IfBranchTop + 1
    ReDim Preserve IfBranchStack(1 To IfBranchTop)
    IfBranchStack(IfBranchTop).headerLine = headerLine
    IfBranchStack(IfBranchTop).endIfLine = endIfLine
End Sub

Public Sub IfBranchPop()
    If IfBranchTop <= 0 Then Exit Sub

    IfBranchTop = IfBranchTop - 1
    If IfBranchTop > 0 Then
        ReDim Preserve IfBranchStack(1 To IfBranchTop)
    Else
        Erase IfBranchStack
    End If
End Sub

Public Function IfBranchPeekEnd() As Long
    If IfBranchTop > 0 Then IfBranchPeekEnd = IfBranchStack(IfBranchTop).endIfLine
End Function

Public Function ReconcileIfBranchStackForJump(ByVal targetLine As Long) As Long
    Dim keepDepth As Long
    Dim i As Long

    keepDepth = 0
    For i = IfBranchTop To 1 Step -1
        If targetLine >= IfBranchStack(i).headerLine And targetLine <= IfBranchStack(i).endIfLine Then
            keepDepth = i
            Exit For
        End If
    Next i

    Do While IfBranchTop > keepDepth
        IfBranchPop
        ReconcileIfBranchStackForJump = ReconcileIfBranchStackForJump + 1
    Loop
End Function
