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
' @depends modEngine, modContext, modCmdRegistry, ICommand
Option Explicit

Implements ICommand

Private Function ExecuteIfBranch(ByVal R As Long, ByRef args() As String) As Long
    Dim endIfLine As Long
    endIfLine = modEngine.ActiveProgram().EndIfOf(R)
    If endIfLine <= 0 Then
        ExecuteIfBranch = CMD_ERROR
        Exit Function
    End If

    If modEngine.CheckIfSyntax(args(1), args(2), args(3)) Then
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
            gotoTarget = modEngine.ResolveGotoTarget(R, args(3), gotoFramesPopped)
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
            callTarget = modEngine.ResolveCallTarget(R, args(3))
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
                ICommand_Execute = CMD_ERROR
            ElseIf modContext.IfBranchPeekEnd() = elseIfEnd Then
                ICommand_Execute = elseIfEnd
            Else
                ICommand_Execute = ExecuteIfBranch(R, args)
            End If

         Case "ELSE"
            ICommand_Execute = modEngine.ActiveProgram().FindBlockPair(R, "IF", 1)

         Case "END_IF"
            If modContext.IfBranchPeekEnd() = R Then modContext.IfBranchPop
            ICommand_Execute = CMD_NEXT_LINE

         Case "END_SELECT"
            ICommand_Execute = CMD_NEXT_LINE

         '--- [Loop] WHILE / END_WHILE (활성 루프 스택) ---
         Case "WHILE"
            Dim whileEnd As Long
            If modEngine.CheckIfSyntax(args(1), args(2), args(3)) Then
                whileEnd = modEngine.ActiveProgram().FindBlockPair(R, "WHILE", 1)
                If whileEnd <= 0 Then
                    ICommand_Execute = CMD_ERROR
                    GoTo CleanUp
                End If
                ' 참: 첫 진입이면 PUSH, 백엣지 재진입(top.header=R)이면 회차+1
                If modContext.LoopDepth() > 0 And modContext.LoopTopHeader() = R Then
                    modContext.LoopIncrement
                Else
                    modContext.LoopPush "WHILE", R, whileEnd, 0
                End If
                ICommand_Execute = CMD_NEXT_LINE
            Else
                ' 거짓: 활성 프레임이 이 WHILE이면 POP 후 탈출
                If modContext.LoopDepth() > 0 And modContext.LoopTopHeader() = R Then modContext.LoopPop "end"
                whileEnd = modEngine.ActiveProgram().FindBlockPair(R, "WHILE", 1)
                If whileEnd > 0 Then ICommand_Execute = whileEnd + 1 Else ICommand_Execute = CMD_ERROR
            End If

         Case "END_WHILE"
            ' 헤더로 백점프. POP은 WHILE 거짓 갈래가 담당 (백엣지 동안 {$LOOP} 유지)
            Dim wHead As Long: wHead = modEngine.ActiveProgram().FindBlockPair(R, "WHILE", -1)
            If wHead > 0 Then ICommand_Execute = wHead Else ICommand_Execute = CMD_ERROR

         '--- [Loop] REPEAT / END_REPEAT (N회 또는 무한 반복) ---
         Case "REPEAT"
            Dim repeatArg As String: repeatArg = Trim$(args(1))
            Dim maxCnt As Long
            Dim rEndRepeat As Long: rEndRepeat = modEngine.ActiveProgram().FindBlockPair(R, "REPEAT", 1)
            If rEndRepeat <= 0 Then
                ICommand_Execute = CMD_ERROR
                GoTo CleanUp
            End If

            If Len(repeatArg) = 0 Then
                ' 빈 반복횟수 -> 무한 반복 (BREAK/ESC로 종료)
                maxCnt = 0
                modContext.LoopPush "REPEAT", R, rEndRepeat, maxCnt
                ICommand_Execute = CMD_NEXT_LINE
            Else
                maxCnt = CLng(val(repeatArg))
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
            Dim rHead As Long: rHead = modEngine.ActiveProgram().FindBlockPair(R, "REPEAT", -1)
            If rHead > 0 And modContext.LoopDepth() > 0 And _
               modContext.LoopTopKind() = "REPEAT" And modContext.LoopTopHeader() = rHead Then
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
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Loop] FOREACH_ROW / END_FOREACH (table snapshot cursor) ---
         Case "FOREACH_ROW"
            Dim fEnd As Long
            fEnd = modEngine.ActiveProgram().FindBlockPair(R, "FOREACH", 1)
            If fEnd <= 0 Then
                ICommand_Execute = CMD_ERROR
                GoTo CleanUp
            End If

            Dim tbl As clsTableSnapshot
            Set tbl = modTableSource.SnapshotSource(args(1), args(2), args(3))

            Dim totalRows As Long
            totalRows = tbl.RowCount - tbl.FirstDataRow + 1
            If totalRows < 0 Then totalRows = 0

            If tbl.RowCount < tbl.FirstDataRow Then
                modContext.LogStatus "FOREACH_ROW: 0행"
                LogShim.LogEvent "foreach.empty", _
                    Array("headerLine", "rtotal"), _
                    Array(R, totalRows)
                ICommand_Execute = fEnd + 1
            Else
                modContext.LoopPushForeach R, fEnd, tbl
                ICommand_Execute = CMD_NEXT_LINE
            End If

         Case "END_FOREACH"
            Dim fHead As Long
            fHead = modEngine.ActiveProgram().FindBlockPair(R, "FOREACH", -1)
            If fHead > 0 And modContext.LoopDepth() > 0 And _
               modContext.LoopTopKind() = "FOREACH" And modContext.LoopTopHeader() = fHead Then
                If modContext.ForeachAdvance() Then
                    ICommand_Execute = fHead + 1
                Else
                    modContext.LoopPop "end"
                    ICommand_Execute = CMD_NEXT_LINE
                End If
            Else
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
                ICommand_Execute = CMD_ERROR
            End If

         Case "CONTINUE"
            If modContext.LoopDepth() > 0 Then
                ' 감싸는 루프의 ENDER로 점프 → WHILE/REPEAT 각각의 ENDER가 알맞게 처리
                Dim continueTarget As Long: continueTarget = modContext.LoopTopEnd()
                modContext.ReconcileIfBranchStackForJump continueTarget
                ICommand_Execute = continueTarget
            Else
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Structured Flow] SELECT / CASE / END_SELECT ---
         Case "SELECT"
            ' SELECT 실행 의미는 Engine, 구조 위치는 ActiveProgram 인덱스가 담당
            ICommand_Execute = modEngine.ResolveSelectTarget(R, args(1))

         Case "CASE"
            ' 실행 중 CASE 도달 = 폴스루 -> END_SELECT 로 점프
            Dim endSelLine As Long
            endSelLine = modEngine.ActiveProgram().EndSelectOfCase(R)
            If endSelLine > 0 Then ICommand_Execute = endSelLine Else ICommand_Execute = CMD_ERROR

         Case Else
            ICommand_Execute = CMD_ERROR
    End Select

CleanUp:
    Exit Function
ErrHandler:
    modActions.Action_ShowNotification "[에러] FLOW " & cmd & ": " & Err.Description, 3
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function
