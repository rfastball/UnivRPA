VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFlowCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 흐름제어 명령어(분기/반복/점프) 16종 디스패치 실행 담당
' @depends modEngine, modContext, modCmdRegistry, ICommand
Option Explicit

Implements ICommand

' @desc 흐름제어 명령어와 디코드된 args 받아 다음 줄/점프/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    ' args(1)=좌변(해석) · args(2)=연산자(RAW) · args(3)=우변/대상(해석)
    ICommand_Execute = CMD_NEXT_LINE

    Select Case cmd
         Case "LABEL": ' Pass (라벨은 PreScanLabels에서 사전 등록됨)

         Case "GOTO"
            ICommand_Execute = modEngine.ResolveJumpTarget(args(3))

         Case "CALL"
            modContext.PushStack R + 1
            ICommand_Execute = modEngine.ResolveJumpTarget(args(3))

         Case "RETURN"
            Dim retLine As Long
            retLine = modContext.PopStack()
            If retLine > 0 Then ICommand_Execute = retLine Else ICommand_Execute = CMD_HALT

         Case "IF"
            If modEngine.CheckIfSyntax(args(1), args(2), args(3)) Then
                ICommand_Execute = CMD_NEXT_LINE
            Else
                Dim targetLine As Long: targetLine = modEngine.FindBlockPair(R, "IF", 1)
                If targetLine > 0 Then
                    ' ELSE 면 그 다음 줄로, END_IF 면 그 줄로
                    If modEngine.CmdAt(targetLine) = "ELSE" Then
                        ICommand_Execute = targetLine + 1
                    Else
                        ICommand_Execute = targetLine
                    End If
                Else
                    ICommand_Execute = CMD_ERROR
                End If
            End If

         Case "ELSE"
            ICommand_Execute = modEngine.FindBlockPair(R, "IF", 1)

         Case "END_IF", "END_SELECT"
            ICommand_Execute = CMD_NEXT_LINE

         '--- [Loop] WHILE / END_WHILE (활성 루프 스택) ---
         Case "WHILE"
            Dim whileEnd As Long
            If modEngine.CheckIfSyntax(args(1), args(2), args(3)) Then
                ' 참: 첫 진입이면 PUSH, 백엣지 재진입(top.header=R)이면 회차+1
                If modContext.LoopDepth() > 0 And modContext.LoopTopHeader() = R Then
                    modContext.LoopIncrement
                Else
                    modContext.LoopPush "WHILE", R, modEngine.FindBlockPair(R, "WHILE", 1), 0
                End If
                ICommand_Execute = CMD_NEXT_LINE
            Else
                ' 거짓: 활성 프레임이 이 WHILE이면 POP 후 탈출
                If modContext.LoopDepth() > 0 And modContext.LoopTopHeader() = R Then modContext.LoopPop
                whileEnd = modEngine.FindBlockPair(R, "WHILE", 1)
                If whileEnd > 0 Then ICommand_Execute = whileEnd + 1 Else ICommand_Execute = CMD_ERROR
            End If

         Case "END_WHILE"
            ' 헤더로 백점프. POP은 WHILE 거짓 갈래가 담당 (백엣지 동안 {$LOOP} 유지)
            Dim wHead As Long: wHead = modEngine.FindBlockPair(R, "WHILE", -1)
            If wHead > 0 Then ICommand_Execute = wHead Else ICommand_Execute = CMD_ERROR

         '--- [Loop] REPEAT / END_REPEAT (N회 반복) ---
         Case "REPEAT"
            Dim maxCnt As Long: maxCnt = CLng(val(args(1)))
            If maxCnt <= 0 Then
                ' 0회 -> 본문 건너뛰기 (PUSH 안 함)
                Dim rEndSkip As Long: rEndSkip = modEngine.FindBlockPair(R, "REPEAT", 1)
                If rEndSkip > 0 Then ICommand_Execute = rEndSkip + 1 Else ICommand_Execute = CMD_ERROR
            Else
                modContext.LoopPush "REPEAT", R, modEngine.FindBlockPair(R, "REPEAT", 1), maxCnt
                ICommand_Execute = CMD_NEXT_LINE
            End If

         Case "END_REPEAT"
            Dim rHead As Long: rHead = modEngine.FindBlockPair(R, "REPEAT", -1)
            If rHead > 0 Then
                If modContext.GetActiveLoopIteration() < modContext.LoopTopMax() Then
                    modContext.LoopIncrement
                    ICommand_Execute = rHead + 1   ' 본문 재실행 (REPEAT 줄은 1회만)
                Else
                    modContext.LoopPop
                    ICommand_Execute = CMD_NEXT_LINE
                End If
            Else
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Loop Control] BREAK / CONTINUE ---
         Case "BREAK"
            If modContext.LoopDepth() > 0 Then
                Dim brkEnd As Long: brkEnd = modContext.LoopTopEnd()
                modContext.LoopPop ' 카운터 정리
                ICommand_Execute = brkEnd + 1
            Else
                ICommand_Execute = CMD_ERROR
            End If

         Case "CONTINUE"
            If modContext.LoopDepth() > 0 Then
                ' 감싸는 루프의 ENDER로 점프 → WHILE/REPEAT 각각의 ENDER가 알맞게 처리
                ICommand_Execute = modContext.LoopTopEnd()
            Else
                ICommand_Execute = CMD_ERROR
            End If

         '--- [Structured Flow] SELECT / CASE / END_SELECT ---
         Case "SELECT"
            ' 전방 스캔(m_arrMain 의존)은 modEngine 엔진 서비스에 위임
            ICommand_Execute = modEngine.ResolveSelectTarget(R, args(1))

         Case "CASE"
            ' 실행 중 CASE 도달 = 폴스루 -> END_SELECT 로 점프
            Dim endSelLine As Long
            endSelLine = modEngine.FindBlockPair(R, "SELECT", 1)
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
