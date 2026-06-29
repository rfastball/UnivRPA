Attribute VB_Name = "modFilterUtils"
Option Explicit

' Applies an Excel number format with graceful fallback for narrow-column "####" snapshots.
Public Function GetSafeFormattedText(ByVal vVal As Variant, ByVal sFmt As String) As String
    Dim sResult As String
    Dim bSuccess As Boolean

    If IsError(vVal) Then
        GetSafeFormattedText = CStr(vVal)
        Exit Function
    End If
    If IsEmpty(vVal) Or IsNull(vVal) Then
        GetSafeFormattedText = ""
        Exit Function
    End If

    If sFmt <> "" Then
        On Error Resume Next
        sResult = Application.WorksheetFunction.text(vVal, sFmt)
        bSuccess = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If

    If Not bSuccess And sFmt <> "" Then
        On Error Resume Next
        sResult = Format(vVal, sFmt)
        bSuccess = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If

    If Not bSuccess Then
        If IsDate(vVal) Then
            sResult = Format(vVal, "yyyy-mm-dd")
        Else
            sResult = CStr(vVal)
        End If
    End If

    GetSafeFormattedText = sResult
End Function
