Attribute VB_Name = "dbgTemp"
Option Explicit

Public g_AbortWait As Boolean

Public Sub RaiseAbort()
    g_AbortWait = True
End Sub

Public Sub LogInvariant_Wait(result As Boolean, elapsed_ms As Long, exit_reason As String)
    LogShim.LogJ "expected_invariant", _
        "{""result"":" & LCase(CStr(result)) & _
        ",""elapsed_ms"":" & elapsed_ms & _
        ",""exit_reason"":""" & exit_reason & """}"
End Sub

Public Sub Scenario_S1_ImmediateMatch()
    On Error GoTo Fail
    Dim startTime As Long: startTime = modWinAPI.GetTickCount()
    Dim curHex As String: curHex = modWinAPI.GetColorHexAt(100, 100)
    Dim r As Boolean: r = modActions.Action_Wait("WAIT_PIXEL", 100, 100, curHex)
    LogInvariant_Wait r, modWinAPI.GetTickCount() - startTime, "matched"
    Exit Sub
Fail:
    LogShim.LogJ "macro_error", "{""scenario"":""S1"",""num"":" & Err.Number & "}"
End Sub

Public Sub Scenario_S2_EscDuring()
    On Error GoTo Fail
    g_AbortWait = False
    Application.OnTime Now + TimeValue("00:00:01"), "'UnivRPA.xlsm'!dbgTemp.RaiseAbort"
    Dim startTime As Long: startTime = modWinAPI.GetTickCount()
    Dim r As Boolean: r = modActions.Action_Wait("WAIT_PIXEL", 100, 100, "FFFFFFFFFF")
    LogInvariant_Wait r, modWinAPI.GetTickCount() - startTime, "esc"
    Exit Sub
Fail:
    LogShim.LogJ "macro_error", "{""scenario"":""S2"",""num"":" & Err.Number & "}"
End Sub

Public Sub Scenario_S3_Timeout()
    On Error GoTo Fail
    g_AbortWait = False
    Dim startTime As Long: startTime = modWinAPI.GetTickCount()
    Dim r As Boolean: r = modActions.Action_Wait("WAIT_PIXEL", 100, 100, "FFFFFFFFFF")
    LogInvariant_Wait r, modWinAPI.GetTickCount() - startTime, "timeout"
    Exit Sub
Fail:
    LogShim.LogJ "macro_error", "{""scenario"":""S3"",""num"":" & Err.Number & "}"
End Sub
