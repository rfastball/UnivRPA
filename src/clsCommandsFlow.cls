VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandsFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 흐름제어 명령어(분기/반복/점프) 19종 디스패치 실행 담당
' @depends modEngine, modContext, modCmdRegistry, modCondition, modTokenResolver, modParse, modHelper, ICommand
Option Explicit

Implements ICommand

Private Function ExecuteIfBranch(ByVal R As Long, ByRef args() As String) As Long
    Dim endIfLine As Long
    endIfLine = modEngine.ActiveProgram().EndIfOf(R)
    If endIfLine <= 0 Then
        modContext.SetLastError "IF 라인 " & R & ": END_IF 짝을 찾지 못했습니다."
        ExecuteIfBranch = CMD_ERROR
        Exit Function
    End If

    If EvalCondition(args(1), args(2), args(3)) Then
        modContext.IfBranchPush R, endIfLine
        ExecuteIfBranch = CMD_NEXT_LINE
    Else
        ExecuteIfBranch = NextIfBranchTarget(R)
    End If
End Function

Private Function NextIfBranchTarget(ByVal R As Long) As Long
    Dim targetLine As Long
    targetLine = modEngine.ActiveProgram().FindBlockPair(R, "IF", 1)
    If targetLine <= 0 Then
        modContext.SetLastError "IF 라인 " & R & ": 다음 분기(ELSE_IF/ELSE/END_IF)를 찾지 못했습니다."
        NextIfBranchTarget = CMD_ERROR
        Exit Function
    End If

    Select Case modEngine.ActiveProgram().CmdAt(targetLine)
        Case "ELSE"
            NextIfBranchTarget = targetLine + 1
        Case "ELSE_IF", "END_IF"
            NextIfBranchTarget = targetLine
        Case Else
            NextIfBranchTarget = targetLine
    End Select
End Function

' @desc 흐름제어 명령어와 디코드된 args 받아 다음 줄/점프/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    ' args(1)=좌변(해석) · args(2)=연산자(RAW) · args(3)=우변/대상(해석)
    ICommand_Execute = CMD_NEXT_LINE

    Select Case cmd
         Case "LABEL": ' Pass (라벨은 clsScriptProgram에서 사전 인덱싱됨)

         Case "GOTO"
            Dim gotoFramesPopped As Long
            Dim gotoTarget As Long
            Dim gotoDepthBefore As Long
            gotoDepthBefore = modContext.LoopDepth()
            gotoTarget = ResolveGotoTarget(R, args(3), gotoFramesPopped)
            If gotoTarget > 0 Then
                LogShim.LogEvent "jump.taken", _
                    Array("line", "cmd", "target", "to", "subAreaId", "loopDepth_before", "loopDepth_after", "framesPopped"), _
                    Array(R, cmd, args(3), gotoTarget, modEngine.ActiveProgram().SubAreaIdAt(gotoTarget), gotoDepthBefore, modContext.LoopDepth(), gotoFramesPopped)
                ICommand_Execute = gotoTarget
            Else
                ICommand_Execute = CMD_ERROR
            End If

         Case "CALL"
            Dim callTarget As Long
            callTarget = ResolveCallTarget(R, args(3))
            If callTarget > 0 Then
                modContext.PushStack R + 1, modContext.LoopDepth()
                LogShim.LogEvent "call.push", _
                    Array("line", "target", "to", "returnRow", "stackDepth_after", "loopDepth_at_call"), _
                    Array(R, args(3), callTarget, R + 1, modContext.GetStackDepth(), modContext.LoopDepth())
                ICommand_Execute = callTarget
            Else
                ICommand_Execute = CMD_ERROR
            End If

         Case "RETURN"
            Dim retLine As Long
            Dim restoredDepth As Long
            retLine = modContext.PopStack(restoredDepth)
            If retLine > 0 Then
                modContext.ReconcileIfBranchStackForJump retLine
                LogShim.LogEvent "return.pop", _
                    Array("line", "returnRow", "stackDepth_after", "loopDepth_restored"), _
                    Array(R, retLine, modContext.GetStackDepth(), restoredDepth)
                ICommand_Execute = retLine
            Else
                ICommand_Execute = CMD_ERROR
            End If

         Case "IF"
            ICommand_Execute = ExecuteIfBranch(R, args)

         Case "ELSE_IF"
            Dim elseIfEnd As Long
            elseIfEnd = modEngine.ActiveProgram().EndIfOf(R)
            If elseIfEnd <= 0 Then
                modContext.SetLastError "ELSE_IF 라인 " & R & ": END_IF 짝을 찾지 못했습니다."
                ICommand_Execute = CMD_ERROR
            ElseIf modContext.IfBranchPeekEnd() = elseIfEnd Then
                ICommand_Execute = elseIfEnd
            Else
                ICommand_Execute = ExecuteIfBranch(R, args)
            End If

         Case "ELSE"
            Dim elseTarget As Long
            elseTarget = modEngine.ActiveProgram().FindBlockPair(R, "IF", 1)
            If elseTarget > 0 Then
                ICommand_Execute = elseTarget
            Else
                modContext.SetLastError "ELSE 라인 " & R & ": END_IF 짝을 찾지 못했습니다."
                ICommand_Execute = CMD_ERROR
            End If

         Case "END_IF"
            If modContext.IfBranchPeekEnd() = R Then modContext.IfBranchPop
            ICommand_Execute = CMD_NEXT_LINE

         Case "END_SELECT"
            ICommand_Execute = CMD_NEXT_LINE

         '--- [Loop] WHILE / END_WHILE (활성 루프 스택) ---
         Case "WHILE"
            Dim whileEnd As Long
            ' 백엣지 재진입이면 프레임이 이미 있다(top.header=R) - 짝 스캔 없이 프레임 값 사용.
            ' 첫 진입은 프레임이 없으므로 END_WHILE 짝을 전방 스캔해야 한다(제거 불가).
            Dim whileReentry As Boolean
            whileReentry = (modContext.LoopDepth() > 0 And modContext.LoopTopHeader() = R)

            If EvalCondition(args(1), args(2), args(3)) Then
                ' 참: 재진입이면 회차+1, 첫 진입이면 짝 스캔 후 PUSH
                If whileReentry Then
                    modContext.LoopIncrement
                Else
                    whileEnd = modEngine.ActiveProgram().FindBlockPair(R, "WHILE", 1)
                    If whileEnd <= 0 Then
                        modContext.SetLastError "WHILE 라인 " & R & ": END_WHILE 짝을 찾지 못했습니다."
                        ICommand_Execute = CMD_ERROR
                        GoTo CleanUp
                    End If
                    modContext.LoopPush "WHILE", R, whileEnd, 0
                End If
                ICommand_Execute = CMD_NEXT_LINE
            Else
                ' 거짓: 재진입이면 POP 전에 프레임 end 포착 후 탈출, 첫 진입이면 짝 스캔으로 탈출
                If whileReentry Then
                    whileEnd = modContext.LoopTopEnd()
                    modContext.LoopPop "end"
                    ICommand_Execute = whileEnd + 1
                Else
                    whileEnd = modEngine.ActiveProgram().FindBlockPair(R, "WHILE", 1)
                    If whileEnd > 0 Then
                        ICommand_Execute = whileEnd + 1
                    Else
                        modContext.SetLastError "WHILE 라인 " & R & ": END_WHILE 짝을 찾지 못했습니다."
                        ICommand_Execute = CMD_ERROR
                    End If
                End If
            End If

         Case "END_WHILE"
            ' 헤더로 백점프. POP은 WHILE 거짓 갈래가 담당 (백엣지 동안 {$LOOP} 유지)
            ' 헤더 줄은 활성 프레임에 이미 있다 - 후방 스캔 없이 조회 (구조상 최상위가 이 WHILE)
            If modContext.LoopDepth() > 0 And modContext.LoopTopKind() = "WHILE" Then
                ICommand_Execute = modContext.LoopTopHeader()
            Else
                modContext.SetLastError "END_WHILE 라인 " & R & ": 활성 WHILE 프레임이 없습니다."
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Loop] REPEAT / END_REPEAT (N회 또는 무한 반복) ---
         Case "REPEAT"
            Dim repeatArg As String: repeatArg = Trim$(args(1))
            Dim maxCnt As Long
            Dim rEndRepeat As Long: rEndRepeat = modEngine.ActiveProgram().FindBlockPair(R, "REPEAT", 1)
            If rEndRepeat <= 0 Then
                modContext.SetLastError "REPEAT 라인 " & R & ": END_REPEAT 짝을 찾지 못했습니다."
                ICommand_Execute = CMD_ERROR
                GoTo CleanUp
            End If

            If Len(repeatArg) = 0 Then
                ' 빈 반복횟수 -> 무한 반복 (BREAK/ESC로 종료)
                maxCnt = 0
                modContext.LoopPush "REPEAT", R, rEndRepeat, maxCnt
                ICommand_Execute = CMD_NEXT_LINE
            Else
                maxCnt = modParse.ParseLongStrict(repeatArg)
                If maxCnt <= 0 Then
                    ' 0 이하 -> 무한 반복 (반복을 끄려면 REPEAT 블록을 주석 처리)
                    maxCnt = 0
                    modContext.LoopPush "REPEAT", R, rEndRepeat, maxCnt
                    ICommand_Execute = CMD_NEXT_LINE
                Else
                    modContext.LoopPush "REPEAT", R, rEndRepeat, maxCnt
                    ICommand_Execute = CMD_NEXT_LINE
                End If
            End If

         Case "END_REPEAT"
            ' 헤더 줄은 활성 프레임에서 조회 - 후방 스캔 제거 (기존엔 스캔값을 프레임과 등가 검증만 했음)
            Dim rHead As Long: rHead = modContext.LoopTopHeader()
            If modContext.LoopDepth() > 0 And modContext.LoopTopKind() = "REPEAT" Then
                If modContext.LoopTopMax() = 0 Then
                    modContext.LoopIncrement
                    ICommand_Execute = rHead + 1   ' 무한 반복: 본문 재실행
                ElseIf modContext.GetActiveLoopIteration() < modContext.LoopTopMax() Then
                    modContext.LoopIncrement
                    ICommand_Execute = rHead + 1   ' 본문 재실행 (REPEAT 줄은 1회만)
                Else
                    modContext.LoopPop "end"
                    ICommand_Execute = CMD_NEXT_LINE
                End If
            Else
                modContext.SetLastError "END_REPEAT 라인 " & R & ": 활성 REPEAT 프레임이 없습니다."
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Loop] FOREACH_ROW / END_FOREACH (table snapshot cursor) ---
         Case "FOREACH_ROW"
            Dim fEnd As Long
            fEnd = modEngine.ActiveProgram().FindBlockPair(R, "FOREACH", 1)
            If fEnd <= 0 Then
                modContext.SetLastError "FOREACH_ROW 라인 " & R & ": END_FOREACH 짝을 찾지 못했습니다."
                ICommand_Execute = CMD_ERROR
                GoTo CleanUp
            End If

            Dim tbl As clsTableSnapshot
            Set tbl = modTableSource.SnapshotSource(args(1), args(2), args(3))

            Dim totalRows As Long
            totalRows = tbl.RowCount - tbl.FirstDataRow + 1
            If totalRows < 0 Then totalRows = 0

            If tbl.RowCount < tbl.FirstDataRow Then
                modHelper.SetStatusBar "FOREACH_ROW: 0행"
                LogShim.LogEvent "foreach.empty", _
                    Array("headerLine", "rtotal"), _
                    Array(R, totalRows)
                ICommand_Execute = fEnd + 1
            Else
                modContext.LoopPushForeach R, fEnd, tbl
                ICommand_Execute = CMD_NEXT_LINE
            End If

         Case "END_FOREACH"
            ' 헤더 줄은 활성 프레임에서 조회 - 행마다 반복되던 후방 스캔 제거 (핫패스)
            Dim fHead As Long: fHead = modContext.LoopTopHeader()
            If modContext.LoopDepth() > 0 And modContext.LoopTopKind() = "FOREACH" Then
                If modContext.ForeachAdvance() Then
                    ICommand_Execute = fHead + 1
                Else
                    modContext.LoopPop "end"
                    ICommand_Execute = CMD_NEXT_LINE
                End If
            Else
                modContext.SetLastError "END_FOREACH 라인 " & R & ": 활성 FOREACH 프레임이 없습니다."
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Loop Control] BREAK / CONTINUE ---
         Case "BREAK"
            If modContext.LoopDepth() > 0 Then
                Dim brkEnd As Long: brkEnd = modContext.LoopTopEnd()
                Dim brkTarget As Long: brkTarget = brkEnd + 1
                modContext.LoopPop "break" ' 카운터 정리
                modContext.ReconcileIfBranchStackForJump brkTarget
                ICommand_Execute = brkTarget
            Else
                modContext.SetLastError "BREAK 라인 " & R & ": 활성 루프가 없습니다."
                ICommand_Execute = CMD_ERROR
            End If

         Case "CONTINUE"
            If modContext.LoopDepth() > 0 Then
                ' 감싸는 루프의 ENDER로 점프 → WHILE/REPEAT 각각의 ENDER가 알맞게 처리
                Dim continueTarget As Long: continueTarget = modContext.LoopTopEnd()
                modContext.ReconcileIfBranchStackForJump continueTarget
                ICommand_Execute = continueTarget
            Else
                modContext.SetLastError "CONTINUE 라인 " & R & ": 활성 루프가 없습니다."
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Structured Flow] SELECT / CASE / END_SELECT ---
         Case "SELECT"
            ' SELECT 실행 의미(CASE 매칭)도 이 핸들러 소유, 구조 위치는 ActiveProgram 인덱스가 담당
            ICommand_Execute = ResolveSelectTarget(R, args(1))

         Case "CASE"
            ' 실행 중 CASE 도달 = 폴스루 -> END_SELECT 로 점프
            Dim endSelLine As Long
            endSelLine = modEngine.ActiveProgram().EndSelectOfCase(R)
            If endSelLine > 0 Then
                ICommand_Execute = endSelLine
            Else
                modContext.SetLastError "CASE 라인 " & R & ": END_SELECT 짝을 찾지 못했습니다."
                ICommand_Execute = CMD_ERROR
            End If

         Case Else
            modContext.SetLastError "FLOW 핸들러에 구현되지 않은 명령어: " & cmd
            ICommand_Execute = CMD_ERROR
    End Select

CleanUp:
    Exit Function
ErrHandler:
    ' 사유만 기록 - 알림은 RunRPA가 LastRuntimeError로 1회 조립 (B1)
    modContext.SetLastError "FLOW " & cmd & " 라인 " & R & ": " & Err.Description
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function

'==============================================================================
' [이관 - P2] 흐름 해석 서비스 (구 modEngine 소유 - 흐름 정책을 이 핸들러로 일원화)
'   에러 사유는 modContext.SetLastError로 기록, 엔진 합류점이 1회 표면화 (B1)
'==============================================================================

' [조건 평가] IF/WHILE 조건절 판단 (v1: 해석된 P1, op: P2 RAW 연산자, v2: 해석된 P3)
Private Function EvalCondition(ByVal v1 As String, ByVal op As String, ByVal v2 As String) As Boolean
    EvalCondition = modCondition.EvaluateCondition(v1, UCase$(Trim$(op)), v2)
End Function

' [GOTO 해석] 정적 라벨 -> 줄번호, 흐름 경계 검사 + 루프/분기 스택 정합까지 수행
Private Function ResolveGotoTarget(ByVal sourceLine As Long, ByVal Target As String, ByRef framesPopped As Long) As Long
    On Error GoTo ErrHandler
    ResolveGotoTarget = CMD_ERROR
    framesPopped = 0

    Dim prog As clsScriptProgram
    Set prog = modEngine.ActiveProgram()
    If prog Is Nothing Then
        modContext.SetLastError "GOTO 라인 " & sourceLine & ": 실행할 ProgramImage가 없습니다."
        Exit Function
    End If

    Dim targetLine As Long
    targetLine = prog.ResolveJumpTarget(Target)
    If targetLine <= 0 Then
        modContext.SetLastError "GOTO 라인 " & sourceLine & ": 대상 라벨 해석 실패: '" & Target & "' (정적 라벨명만 허용)"
        Exit Function
    End If

    If Not prog.IsGotoAllowed(sourceLine, targetLine) Then
        modContext.SetLastError "GOTO 라인 " & sourceLine & ": 라인 " & targetLine & "으로의 점프가 흐름 경계(루프/영역)를 침범합니다."
        Exit Function
    End If

    Dim popped As Long
    popped = modContext.ReconcileLoopStackForJump(targetLine, prog.LoopDepthAt(targetLine))
    If popped < 0 Then
        modContext.SetLastError "GOTO 라인 " & sourceLine & ": 대상 라인 " & targetLine & "의 루프 깊이와 활성 루프 스택이 불일치합니다."
        Exit Function
    End If
    modContext.ReconcileIfBranchStackForJump targetLine

    framesPopped = popped
    ResolveGotoTarget = targetLine
    Exit Function

ErrHandler:
    modContext.SetLastError "GOTO 라인 " & sourceLine & " 해석 오류: " & Err.Description
    ResolveGotoTarget = CMD_ERROR
End Function

' [CALL 해석] 정적 라벨 -> 줄번호, 루프 본문 밖->안 진입 차단
Private Function ResolveCallTarget(ByVal sourceLine As Long, ByVal Target As String) As Long
    On Error GoTo ErrHandler
    Dim prog As clsScriptProgram
    Set prog = modEngine.ActiveProgram()
    If prog Is Nothing Then
        modContext.SetLastError "CALL 라인 " & sourceLine & ": 실행할 ProgramImage가 없습니다."
        ResolveCallTarget = CMD_ERROR
    Else
        ResolveCallTarget = prog.ResolveJumpTarget(Target)
        If ResolveCallTarget <= 0 Then
            modContext.SetLastError "CALL 라인 " & sourceLine & ": 대상 라벨 해석 실패: '" & Target & "' (정적 라벨명만 허용)"
        ElseIf Not prog.IsCallAllowed(sourceLine, ResolveCallTarget) Then
            ' [CALL 경계] 루프 헤더를 거치지 않은 본문 진입은 LoopStack 프레임 부재로
            '   END_* 백점프/RETURN 깊이 복원이 원인에서 먼 곳에서 깨진다 - 여기서 차단
            modContext.SetLastError "CALL 라인 " & sourceLine & ": 라인 " & ResolveCallTarget & "은 루프 본문 내부라 밖에서 진입할 수 없습니다."
            ResolveCallTarget = CMD_ERROR
        End If
    End If
    Exit Function

ErrHandler:
    modContext.SetLastError "CALL 라인 " & sourceLine & " 해석 오류: " & Err.Description
    ResolveCallTarget = CMD_ERROR
End Function

' [SELECT 해석] CASE 매칭 -> 일치 CASE 다음 줄 / 폴백 / END_SELECT 반환
Private Function ResolveSelectTarget(ByVal R As Long, ByVal selVal As String) As Long
    On Error GoTo ErrHandler
    Dim prog As clsScriptProgram
    Set prog = modEngine.ActiveProgram()
    If prog Is Nothing Then
        modContext.SetLastError "SELECT 라인 " & R & ": 실행할 ProgramImage가 없습니다."
        ResolveSelectTarget = CMD_ERROR
        GoTo CleanUp
    End If

    Dim block As clsSelectBlock
    Set block = prog.SelectBlock(R)
    If block Is Nothing Then
        modContext.SetLastError "SELECT 라인 " & R & ": SELECT 블록 인덱스를 찾지 못했습니다."
        ResolveSelectTarget = CMD_ERROR
        GoTo CleanUp
    End If
    If block.EndSelectLine <= 0 Then
        modContext.SetLastError "SELECT 라인 " & R & ": END_SELECT 짝이 없습니다."
        ResolveSelectTarget = CMD_ERROR
        GoTo CleanUp
    End If

    Dim caseVal As String
    Dim fallbackLine As Long
    fallbackLine = 0

    Dim i As Long
    For i = 1 To block.CaseCount
        Dim caseLine As Long
        caseLine = block.CaseLineAt(i)
        caseVal = modTokenResolver.ResolveParam(prog.InstructionAt(caseLine).p1)

        ' Case ELSE (또는 빈 값) = 폴백. 여러 개면 마지막 폴백이 이긴다.
        If UCase$(Trim$(caseVal)) = "ELSE" Or Len(caseVal) = 0 Then
            fallbackLine = caseLine
        ElseIf modCondition.EvaluateCondition(selVal, "EQ", caseVal) Then
            ResolveSelectTarget = caseLine + 1
            GoTo CleanUp
        End If
    Next i

    If fallbackLine > 0 Then
        ResolveSelectTarget = fallbackLine + 1
    Else
        ResolveSelectTarget = block.EndSelectLine
    End If

CleanUp:
    Exit Function
ErrHandler:
    modContext.SetLastError "SELECT 라인 " & R & " 해석 오류: " & Err.Description
    ResolveSelectTarget = CMD_ERROR
    Resume CleanUp
End Function
