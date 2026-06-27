VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystemCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 시스템 명령어(창포커스/대기/알림/종료) 4종 디스패치 실행 담당
' @depends modActions, modCmdRegistry, ICommand
Option Explicit

Implements ICommand

' @desc 시스템 명령어와 디코드된 args 받아 동작 수행 후 다음 줄/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    ICommand_Execute = CMD_NEXT_LINE

    Select Case cmd
        Case "FOCUS"
            If Not modActions.Action_FocusWindow(args(3)) Then
                ICommand_Execute = CMD_ERROR
                GoTo CleanUp
            End If
            modActions.BetterWait 500

        Case "SLEEP"
            modActions.BetterWait CLng(val(args(3)))

        Case "NOTIFY"
            ' P1 = 표시시간초(빈값=2, 0=확인대기), P3 = 메시지
            Dim nMsg As String, nSec As Long, durMs As Long
            nMsg = args(3)
            If Len(Trim$(args(1))) = 0 Then
                durMs = 2000
            Else
                nSec = CLng(val(args(1)))
                durMs = nSec * 1000
            End If
            modActions.Action_Notify durMs, 64, nMsg

        Case "EXIT"
            ICommand_Execute = CMD_HALT

        Case Else
            ICommand_Execute = CMD_ERROR
    End Select

CleanUp:
    Exit Function
ErrHandler:
    modActions.Action_ShowNotification "[에러] SYSTEM " & cmd & ": " & Err.Description, 3
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function
