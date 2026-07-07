VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandsAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 입력 액션 명령어(마우스/키보드) 6종 디스패치 실행 담당
' @depends modActions, modCmdRegistry, ICommand
Option Explicit

Implements ICommand

' @desc 입력 명령어와 디코드된 args 받아 동작 수행 후 다음 줄/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    Dim p1 As Long, p2 As Long
    p1 = CLng(val(args(1)))
    p2 = CLng(val(args(2)))

    ICommand_Execute = CMD_NEXT_LINE

    Select Case cmd
        Case "MOVE":   modActions.Action_MouseMove p1, p2
        Case "CLICK":  modActions.Action_MouseClick p1, p2, False
        Case "RCLICK": modActions.Action_MouseClick p1, p2, True
        Case "KEY":    modActions.Action_PressKey args(1)
        Case "COMBO":  modActions.Action_PressCombo args(1), args(2)
        Case "TYPE"
            ' P1 = 딜레이ms, P3 = 텍스트
            modActions.Action_SmartType CLng(val(args(1))), args(3)
        Case Else
            ICommand_Execute = CMD_ERROR
            GoTo CleanUp
    End Select

    ' [사후 대기] MOVE/CLICK/RCLICK/KEY/COMBO 만 P3에 대기시간(ms)을 가질 수 있음
    '   (TYPE=텍스트 이므로 제외)
    If IsNumeric(args(3)) And val(args(3)) > 0 Then
        Select Case cmd
            Case "MOVE", "CLICK", "RCLICK", "KEY", "COMBO"
                modActions.BetterWait CLng(val(args(3)))
        End Select
    End If

CleanUp:
    Exit Function
ErrHandler:
    modActions.Action_ShowNotification "[에러] ACTION " & cmd & ": " & Err.Description, 3
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function
