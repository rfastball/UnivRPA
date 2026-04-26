Attribute VB_Name = "LogShim"
' @desc 본 코드용 디버그 shim, 애드인 부재 시 무해 무시 동작 노출
Option Explicit

Private Const ADDIN_FILENAME As String = "_UnivDbg.xlam"
Private Const ADDIN_PREFIX As String = "_UnivDbg.xlam!UnivDbg."

' Empty=미확인, True=가용, False=미가용 (1회 Ping 후 캐시)
Private dbgEnabled As Variant

' @desc level과 msg 받아 애드인 LogD에 위임함, 애드인 부재 시 무해 무시
Public Sub LogD(level As String, msg As String)
    If Not EnsureEnabled() Then Exit Sub
    On Error Resume Next
    Application.Run ADDIN_PREFIX & "LogD", level, msg
End Sub

' @desc eventName과 dataJson 받아 애드인 LogJ에 위임함, 애드인 부재 시 무해 무시
Public Sub LogJ(eventName As String, dataJson As String)
    If Not EnsureEnabled() Then Exit Sub
    On Error Resume Next
    Application.Run ADDIN_PREFIX & "LogJ", eventName, dataJson
End Sub

' @desc 애드인 Reset에 위임하여 jsonl 파일을 초기화함, 애드인 부재 시 무해 무시
Public Sub Reset()
    If Not EnsureEnabled() Then Exit Sub
    On Error Resume Next
    Application.Run ADDIN_PREFIX & "Reset"
End Sub

' @desc 가용성 판정 단계: Ping→Workbooks.Open 폴백→Ping 재시도. sbroenne hidden 세션의 추가기능 자동 로드 부재 대응
Private Function EnsureEnabled() As Boolean
    If Not IsEmpty(dbgEnabled) Then
        EnsureEnabled = (dbgEnabled = True)
        Exit Function
    End If

    On Error Resume Next

    ' 1차 시도: 이미 로드된 애드인 Ping
    Err.Clear
    Application.Run ADDIN_PREFIX & "Ping"
    If Err.Number = 0 Then
        dbgEnabled = True
        EnsureEnabled = True
        Exit Function
    End If

    ' 2차 시도: UserLibraryPath에서 명시 로드 후 Ping 재시도
    Err.Clear
    Dim addinPath As String
    addinPath = Application.UserLibraryPath & ADDIN_FILENAME
    If LenB(Dir(addinPath)) > 0 Then
        Workbooks.Open fileName:=addinPath
        Err.Clear
        Application.Run ADDIN_PREFIX & "Ping"
        If Err.Number = 0 Then
            dbgEnabled = True
            EnsureEnabled = True
            Exit Function
        End If
    End If

    dbgEnabled = False
    EnsureEnabled = False
End Function

