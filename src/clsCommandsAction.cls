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
' @depends modActions, modContext, modCmdRegistry, modParse, ICommand
Option Explicit

Implements ICommand

' @desc 입력 명령어와 디코드된 args 받아 동작 수행 후 다음 줄/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    ICommand_Execute = CMD_NEXT_LINE

    ' 좌표(P1/P2)는 DT_NUM 슬롯인 분기에서만 파싱한다
    ' (KEY/COMBO의 P1/P2는 문자열 - 과거의 무조건 CLng(Val())는 데드 파스였음)
    Select Case cmd
        Case "MOVE":   modActions.Action_MouseMove modParse.ParseLongStrict(args(1)), modParse.ParseLongStrict(args(2))
        Case "CLICK":  modActions.Action_MouseClick modParse.ParseLongStrict(args(1)), modParse.ParseLongStrict(args(2)), False
        Case "RCLICK": modActions.Action_MouseClick modParse.ParseLongStrict(args(1)), modParse.ParseLongStrict(args(2)), True
        Case "KEY":    modActions.Action_PressKey args(1)
        Case "COMBO":  modActions.Action_PressCombo args(1), args(2)
        Case "TYPE"
            ' P1 = 딜레이ms(선택, 빈값=0), P3 = 텍스트
            modActions.Action_SmartType modParse.ParseLongOrDefault(args(1), 0), args(3)
        Case Else
            modContext.SetLastError "ACTION 핸들러에 구현되지 않은 명령어: " & cmd
            ICommand_Execute = CMD_ERROR
            GoTo CleanUp
    End Select

    ' [사후 대기] MOVE/CLICK/RCLICK/KEY/COMBO 만 P3에 대기시간(ms)을 가질 수 있음
    '   (TYPE=텍스트 이므로 제외)
    Dim waitMs As Long
    If modParse.TryParseLongStrict(args(3), waitMs) And waitMs > 0 Then
        Select Case cmd
            Case "MOVE", "CLICK", "RCLICK", "KEY", "COMBO"
                modActions.BetterWait waitMs
        End Select
    End If

CleanUp:
    Exit Function
ErrHandler:
    ' 사유만 기록 - 알림은 RunRPA가 LastRuntimeError로 1회 조립 (B1)
    modContext.SetLastError "ACTION " & cmd & " 라인 " & R & ": " & Err.Description
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function
