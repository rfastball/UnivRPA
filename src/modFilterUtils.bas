Attribute VB_Name = "modFilterUtils"
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' @desc 일괄 읽은 셀 값의 타입을 보존하고 통화값만 원래 정밀도로 반환함
Public Function SnapshotCellValue(ByVal typedValue As Variant, ByVal source As Range, _
                                  ByVal rowIndex As Long, ByVal columnIndex As Long) As Variant
    If VarType(typedValue) = vbCurrency Then
        SnapshotCellValue = source.Cells(rowIndex, columnIndex).Value2
    Else
        SnapshotCellValue = typedValue
    End If
End Function

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
