VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandsMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 이름 붙은 슬롯 쓰기(SET/INC/DEC) 실행 및 lvalue 합법성 소유 담당
' @depends modContext, modWinAPI, modCmdRegistry, ICommand
Option Explicit

Implements ICommand

' @desc 쓰기 명령어와 디코드된 args 받아 슬롯에 기록/변이 후 다음 줄/에러 코드 반환
Private Function ICommand_Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
    On Error GoTo ErrHandler

    ICommand_Execute = CMD_NEXT_LINE

    ' args(1)=좌변(RAW 주소) · args(3)=값/증분(해석본)
    Dim kind As String, key As String
    If Not ParseLValue(args(1), kind, key) Then
        ' 좌변이 쓰기 가능한 토큰이 아님 (읽기전용/형식오류)
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

            ' [정수 read-modify-write] 기존값을 CLng(Val(...))로 강제 -> 비숫자는 0
            Dim cur As Long, delta As Long
            cur = CLng(val(ReadSlot(kind, key)))

            ' 증분 P3: 생략 시 1, 변수/음수 허용
            If Trim$(args(3)) = "" Then delta = 1 Else delta = CLng(val(args(3)))
            If cmd = "DEC" Then delta = -delta

            WriteSlot kind, key, CStr(cur + delta)

        Case Else
            ICommand_Execute = CMD_ERROR
    End Select

CleanUp:
    Exit Function
ErrHandler:
    modActions.Action_ShowNotification "[에러] MEMORY " & cmd & ": " & Err.Description, 3
    ICommand_Execute = CMD_ERROR
    Resume CleanUp
End Function

' [흡수] SET 좌변(주소) 파싱. 좌변은 RAW (값 해석 금지)
'   반환 True 시 kind="MEM"|"CLIP", key=메모리명. 읽기전용 토큰은 좌변 불가(False).
'   "어떤 좌변이 쓰기 가능한가"는 메모리 명령어의 고유 지식 -> 이 클래스가 소유.
Private Function ParseLValue(ByVal rawP1 As String, ByRef kind As String, ByRef key As String) As Boolean
    Dim s As String: s = Trim(rawP1)
    ParseLValue = False
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "{" Or Right$(s, 1) <> "}" Then Exit Function

    Dim inner As String: inner = Trim(Mid$(s, 2, Len(s) - 2))
    If UCase$(inner) = "$CLIP" Then
        kind = "CLIP": key = ""
        ParseLValue = True
    ElseIf Left$(inner, 1) = "$" Then
        ' 읽기 전용 시스템 토큰은 좌변(쓰기) 불가
        ParseLValue = False
    Else
        kind = "MEM": key = inner
        ParseLValue = True
    End If
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
