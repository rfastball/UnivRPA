Attribute VB_Name = "modCondition"
' @desc Condition comparison/operator SSOT for the script runtime layer.
' @depends modParse
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' [Logic] condition comparison
Public Function EvaluateCondition(ByVal v1 As String, ByVal op As String, ByVal v2 As String) As Boolean
    Dim s1 As String, s2 As String, q1 As Boolean, q2 As Boolean
    s1 = NormalizeComparable(Trim$(v1), q1)
    s2 = NormalizeComparable(Trim$(v2), q2)
    If Not (q1 Or q2) Then
        s1 = UCase$(s1)
        s2 = UCase$(s2)
    End If

    op = UCase$(Trim$(op))
    Select Case op
        Case "LIKE": EvaluateCondition = (s1 Like s2)
        Case "STARTS": EvaluateCondition = (Len(s2) > 0 And Left$(s1, Len(s2)) = s2)
        Case "ENDS": EvaluateCondition = (Len(s2) > 0 And Right$(s1, Len(s2)) = s2)
        Case "CONTAINS"
            EvaluateCondition = (InStr(1, s1, s2, IIf(q1 Or q2, vbBinaryCompare, vbTextCompare)) > 0)
        Case "EQ", "==", "!=", ">", "<", ">=", "<="
            Dim cmp As Long
            If Not (q1 Or q2) And IsNumericPair(s1, s2) Then
                cmp = CompareDecimalText(s1, s2)
            Else
                cmp = StrComp(s1, s2, vbBinaryCompare)
            End If
            Select Case op
                Case "EQ", "==": EvaluateCondition = (cmp = 0)
                Case "!=": EvaluateCondition = (cmp <> 0)
                Case ">": EvaluateCondition = (cmp > 0)
                Case "<": EvaluateCondition = (cmp < 0)
                Case ">=": EvaluateCondition = (cmp >= 0)
                Case "<=": EvaluateCondition = (cmp <= 0)
            End Select
    End Select
End Function

Private Function IsNumericPair(ByVal s1 As String, ByVal s2 As String) As Boolean
    If s1 Like "0#*" Or s2 Like "0#*" Then Exit Function
    IsNumericPair = modParse.IsStrictNumberSyntax(s1) And modParse.IsStrictNumberSyntax(s2)
End Function

' Equal-width whole parts can be compared digit by digit, including the fractional part.
Private Function CompareDecimalText(ByVal s1 As String, ByVal s2 As String) As Long
    Dim sign1 As Long, sign2 As Long, width1 As Long, width2 As Long
    s1 = DecimalMagnitude(s1, sign1, width1)
    s2 = DecimalMagnitude(s2, sign2, width2)
    If sign1 <> sign2 Then
        CompareDecimalText = Sgn(sign1 - sign2)
    ElseIf width1 <> width2 Then
        CompareDecimalText = sign1 * Sgn(width1 - width2)
    Else
        CompareDecimalText = sign1 * StrComp(s1, s2, vbBinaryCompare)
    End If
End Function

Private Function DecimalMagnitude(ByVal value As String, ByRef numberSign As Long, ByRef wholeWidth As Long) As String
    numberSign = 1
    If Left$(value, 1) = "-" Then numberSign = -1
    Dim first As Long, dot As Long, last As Long
    first = 1
    If Left$(value, 1) = "-" Or Left$(value, 1) = "+" Then first = 2
    dot = InStr(value, ".")
    last = Len(value)
    If dot = 0 Then dot = last + 1
    Do While first < dot
        If Mid$(value, first, 1) <> "0" Then Exit Do
        first = first + 1
    Loop
    Do While last > dot
        If Mid$(value, last, 1) <> "0" Then Exit Do
        last = last - 1
    Loop
    If last = dot Then last = dot - 1
    wholeWidth = dot - first
    If last < first Then
        numberSign = 0
    Else
        DecimalMagnitude = Mid$(value, first, last - first + 1)
    End If
End Function

Public Function IsConditionOperator(ByVal op As String) As Boolean
    Select Case UCase$(Trim$(op))
        Case "EQ", "==", "!=", ">", "<", ">=", "<=", "LIKE", "STARTS", "ENDS", "CONTAINS"
            IsConditionOperator = True
    End Select
End Function

Public Function NormalizeComparable(ByVal value As String, ByRef isQuoted As Boolean) As String
    isQuoted = False
    If Len(value) >= 2 And Left$(value, 1) = """" And Right$(value, 1) = """" Then
        NormalizeComparable = Mid$(value, 2, Len(value) - 2)
        isQuoted = True
    Else
        NormalizeComparable = value
    End If
End Function

Public Function TryCoerceNumber(ByVal s1 As String, ByVal s2 As String, _
                                ByRef d1 As Double, ByRef d2 As Double) As Boolean
    If s1 Like "0#*" Or s2 Like "0#*" Then Exit Function
    If modParse.TryParseDoubleStrict(s1, d1) And modParse.TryParseDoubleStrict(s2, d2) Then
        TryCoerceNumber = True
    End If
End Function
