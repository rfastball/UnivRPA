VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandsMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 이름 붙은 슬롯 쓰기(SET/INC/DEC) 실행 담당
' @depends modContext, modWinAPI, modCmdRegistry, modParse, ICommand
Option Explicit

Implements ICommand

' @desc 쓰기 명령어와 디코드된 args 받아 슬롯에 기록/변이 후 다음 줄/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    ICommand_Execute = CMD_NEXT_LINE

    ' args(1)=좌변(RAW 주소) · args(3)=값/증분(해석본)
    Dim kind As String, key As String
    If Not modParse.TryParseLValue(args(1), kind, key) Then
        ' 좌변이 쓰기 가능한 토큰이 아님 (읽기전용/형식오류)
        modContext.SetLastError cmd & " 라인 " & R & ": P1이 쓰기 가능한 대상이 아닙니다: '" & args(1) & "'"
        ICommand_Execute = CMD_ERROR
        GoTo CleanUp
    End If

    Select Case cmd
        Case "SET"
            WriteSlot kind, key, args(3)

        Case "INC", "DEC"
            If kind = "MEM" And Not modContext.HasVar(key) Then
                Err.Raise 5, "clsCommandsMemory.Execute", "Undefined variable: {" & key & "}. Initialize it with SET first."
            End If

            ' [정수 read-modify-write] 기존값 비숫자는 0 (기존 계약 유지)
            Dim cur As Long, delta As Long
            cur = modParse.ParseLongOrDefault(ReadSlot(kind, key), 0)

            ' 증분 P3: 생략 시 1, 변수/음수 허용 (비어있지 않으면 게이트 통과 보장)
            If Trim$(args(3)) = "" Then delta = 1 Else delta = modParse.ParseLongStrict(args(3))
            If cmd = "DEC" Then delta = -delta

            WriteSlot kind, key, CStr(cur + delta)

        Case Else
            modContext.SetLastError "MEMORY 핸들러에 구현되지 않은 명령어: " & cmd
            ICommand_Execute = CMD_ERROR
    End Select

CleanUp:
    Exit Function
ErrHandler:
    ' 사유만 기록 - 알림은 RunRPA가 LastRuntimeError로 1회 조립 (B1)
    modContext.SetLastError "MEMORY " & cmd & " 라인 " & R & ": " & Err.Description
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function

' [헬퍼] 쓰기 대상에서 현재값 읽기 (INC/DEC read 단계)
Private Function ReadSlot(ByVal kind As String, ByVal key As String) As String
    If kind = "CLIP" Then
        ReadSlot = modWinAPI.GetClipboardText()
    Else
        ReadSlot = modContext.GetVar(key)
    End If
End Function

' [헬퍼] 쓰기 대상에 값 기록 (SET / INC/DEC write 단계)
Private Sub WriteSlot(ByVal kind As String, ByVal key As String, ByVal val As String)
    If kind = "CLIP" Then
        modWinAPI.SetClipboardText val
    Else
        modContext.SetVar key, val
    End If
End Sub
