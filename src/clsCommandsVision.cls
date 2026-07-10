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
' @depends modActions, modContext, modCmdRegistry, modParse, ICommand
Option Explicit

Implements ICommand

' @desc 비전 명령어와 디코드된 args 받아 타임아웃 동작 수행 후 다음 줄/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    ' 좌표 P1/P2는 전 명령 공통 DT_NUM 필수 슬롯 (게이트 통과 보장)
    Dim p1 As Long, p2 As Long
    p1 = modParse.ParseLongStrict(args(1))
    p2 = modParse.ParseLongStrict(args(2))

    ICommand_Execute = CMD_NEXT_LINE

    Select Case cmd
        Case "WAIT_APPEAR", "WAIT_VANISH"
            If Not modActions.Action_Wait(cmd, p1, p2, args(3)) Then
                modContext.SetLastError cmd & " 라인 " & R & ": (" & p1 & "," & p2 & ") 대기 실패 (타임아웃/중단)"
                ICommand_Execute = CMD_ERROR
            End If

        Case "CLICK_SYNC"
            If Not modActions.Action_ClickSync(p1, p2) Then
                modContext.SetLastError "CLICK_SYNC 라인 " & R & ": (" & p1 & "," & p2 & ") 클릭 후 화면 변화 감지 실패 (타임아웃/중단)"
                ICommand_Execute = CMD_ERROR
            End If

        Case "ASSERT"
            If Not modActions.Action_Assert(p1, p2, args(3)) Then
                modContext.SetLastError "ASSERT 라인 " & R & ": (" & p1 & "," & p2 & ") 색상 단언 실패: 기대값 '" & args(3) & "'"
                ICommand_Execute = CMD_ERROR
            End If

        Case Else
            modContext.SetLastError "VISION 핸들러에 구현되지 않은 명령어: " & cmd
            ICommand_Execute = CMD_ERROR
    End Select

CleanUp:
    Exit Function
ErrHandler:
    ' 사유만 기록 - 알림은 RunRPA가 LastRuntimeError로 1회 조립 (B1)
    modContext.SetLastError "VISION " & cmd & " 라인 " & R & ": " & Err.Description
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function
