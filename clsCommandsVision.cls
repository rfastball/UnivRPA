VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandsVision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 비전 명령어(픽셀 대기/동기 클릭/단언) 4종 디스패치 실행 담당
' @depends modActions, modCmdRegistry, ICommand
Option Explicit

Implements ICommand

' @desc 비전 명령어와 디코드된 args 받아 타임아웃 동작 수행 후 다음 줄/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    Dim p1 As Long, p2 As Long
    p1 = CLng(val(args(1)))
    p2 = CLng(val(args(2)))

    ICommand_Execute = CMD_NEXT_LINE

    Select Case cmd
        Case "WAIT_APPEAR", "WAIT_VANISH"
            If Not modActions.Action_Wait(cmd, p1, p2, args(3)) Then ICommand_Execute = CMD_ERROR

        Case "CLICK_SYNC"
            If Not modActions.Action_ClickSync(p1, p2) Then ICommand_Execute = CMD_ERROR

        Case "ASSERT"
            If Not modActions.Action_Assert(p1, p2, args(3)) Then ICommand_Execute = CMD_ERROR

        Case Else
            ICommand_Execute = CMD_ERROR
    End Select

CleanUp:
    Exit Function
ErrHandler:
    modActions.Action_ShowNotification "[에러] VISION " & cmd & ": " & Err.Description, 3
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function
