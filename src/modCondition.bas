Attribute VB_Name = "modCondition"
' @desc Condition comparison/operator SSOT for the script runtime layer.
' @depends modParse
Option Explicit

' [Logic] condition comparison
Public Function EvaluateCondition(ByVal v1 As String, ByVal op As String, ByVal v2 As String) As Boolean
    On Error GoTo ErrHandler
    Dim d1 As Double, d2 As Double
    Dim s1 As String, s2 As String
    Dim isNum As Boolean
    Dim isExplicit As Boolean
    Dim q1 As Boolean, q2 As Boolean

    v1 = Trim(v1)
    v2 = Trim(v2)

    s1 = NormalizeComparable(v1, q1)
    s2 = NormalizeComparable(v2, q2)

    If q1 Or q2 Then
        isExplicit = True
        isNum = False
    ElseIf TryCoerceNumber(s1, s2, d1, d2) Then
        isNum = True
    Else
        isNum = False
        s1 = UCase(s1)
        s2 = UCase(s2)
    End If

    op = UCase(Trim(op))

    Select Case op
        Case "EQ", "=="
            If isNum Then
                EvaluateCondition = (d1 = d2)
            Else
                EvaluateCondition = (s1 = s2)
            End If

        Case "!="
            If isNum Then
                EvaluateCondition = (d1 <> d2)
            Else
                EvaluateCondition = (s1 <> s2)
            End If

        Case ">"
            If isNum Then EvaluateCondition = (d1 > d2) Else EvaluateCondition = (s1 > s2)

        Case "<"
            If isNum Then EvaluateCondition = (d1 < d2) Else EvaluateCondition = (s1 < s2)

        Case ">="
            If isNum Then EvaluateCondition = (d1 >= d2) Else EvaluateCondition = (s1 >= s2)

        Case "<="
            If isNum Then EvaluateCondition = (d1 <= d2) Else EvaluateCondition = (s1 <= s2)

        Case "LIKE"
            EvaluateCondition = (s1 Like s2)

        Case "STARTS"
            EvaluateCondition = (Len(s2) > 0 And Left$(s1, Len(s2)) = s2)

        Case "ENDS"
            EvaluateCondition = (Len(s2) > 0 And Right$(s1, Len(s2)) = s2)

        Case "CONTAINS"
            If isExplicit Then
                EvaluateCondition = (InStr(1, s1, s2, vbBinaryCompare) > 0)
            Else
                EvaluateCondition = (InStr(1, s1, s2, vbTextCompare) > 0)
            End If

        Case Else
            EvaluateCondition = False
    End Select

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modCondition.EvaluateCondition", Err.Description
    Resume CleanUp
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
    ' zero-pad 가드: "007" 같은 코드 문자열은 숫자 비교로 승격하지 않는다
    If s1 Like "0#*" Or s2 Like "0#*" Then Exit Function
    If modParse.TryParseDoubleStrict(s1, d1) And modParse.TryParseDoubleStrict(s2, d2) Then
        TryCoerceNumber = True
    End If
End Function
