Attribute VB_Name = "modScriptAssist"
' @desc Script Assist: COMMAND validation, parameter hints, typo suggestions, and Assist-owned notes
' @depends modCmdRegistry, modDevTools, CmdDef, ParamSpec
Option Explicit

Private Const ASSIST_PREFIX As String = "[ScriptAssist]"
Private Const CMD_COL As Long = 2

Public Function ResolveCmdCol(ByVal Sh As Worksheet) As Long
    If IsScriptSheet(Sh) Then ResolveCmdCol = CMD_COL
End Function

Public Function IsScriptSheet(ByVal Sh As Worksheet) As Boolean
    On Error GoTo SafeExit
    IsScriptSheet = (UCase$(Trim$(CStr(Sh.Cells(1, 1).Value))) = "ROW" _
                  And UCase$(Trim$(CStr(Sh.Cells(1, 2).Value))) = "COMMAND")
SafeExit:
End Function

Public Sub SetupScriptSheet(ByVal Sh As Worksheet)
    modDevTools.EnsureHeaders Sh
End Sub

Public Sub ValidateCommandCell(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Target.Cells.count <> 1 Then Exit Sub
    If ResolveCmdCol(Target.Worksheet) <> Target.Column Then Exit Sub

    Dim rawCmd As Variant
    rawCmd = Target.Value

    If IsBlankCommand(rawCmd) Or IsCommentCommand(rawCmd) Then
        ClearAssistNote Target
        ClearStatusBar
        Exit Sub
    End If

    Dim cmd As String
    cmd = NormalizeCmd(rawCmd)

    If modDevTools.ValidateLine(cmd, Target.offset(0, 1).Value, Target.offset(0, 2).Value, Target.offset(0, 3).Value) Then
        ClearAssistNote Target
    Else
        Dim suggestion As String
        Dim def As CmdDef
        Set def = modCmdRegistry.GetCmdDef(cmd)
        If def Is Nothing Then suggestion = GetClosestCommand(rawCmd)
        ApplyErrorNote Target, rawCmd, suggestion
    End If
End Sub

Public Sub ShowParamHint(ByVal Target As Range)
    If Target Is Nothing Then
        ClearStatusBar
        Exit Sub
    End If

    If Target.Cells.count <> 1 Then
        ClearStatusBar
        Exit Sub
    End If

    If ResolveCmdCol(Target.Worksheet) <> Target.Column Then
        ClearStatusBar
        Exit Sub
    End If

    If IsBlankCommand(Target.Value) Or IsCommentCommand(Target.Value) Then
        ClearStatusBar
        Exit Sub
    End If

    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(Target.Value)
    If def Is Nothing Then
        ClearStatusBar
        Exit Sub
    End If

    Application.StatusBar = FormatParamHint(def)
End Sub

Public Function GetClosestCommand(ByVal rawCmd As Variant) As String
    Dim cmd As String
    cmd = NormalizeCmd(rawCmd)
    If Len(cmd) = 0 Or Left$(cmd, 1) = "#" Then Exit Function

    Dim exactDef As CmdDef
    Set exactDef = modCmdRegistry.GetCmdDef(cmd)
    If Not exactDef Is Nothing Then Exit Function

    Dim threshold As Long
    If Len(cmd) <= 4 Then
        threshold = 1
    Else
        threshold = 2
    End If

    Dim commands As Variant
    commands = modCmdRegistry.AllCommands()

    Dim bestDistance As Long
    bestDistance = 32767

    Dim bestCommand As String
    Dim bestCount As Long
    Dim item As Variant
    For Each item In commands
        Dim candidate As String
        candidate = UCase$(CStr(item))

        Dim distance As Long
        distance = LevenshteinDistance(cmd, candidate)

        If distance <= threshold Then
            If distance < bestDistance Then
                bestDistance = distance
                bestCommand = candidate
                bestCount = 1
            ElseIf distance = bestDistance Then
                bestCount = bestCount + 1
            End If
        End If
    Next item

    If bestCount = 1 Then GetClosestCommand = bestCommand
End Function

Public Sub ApplyErrorNote(ByVal Target As Range, ByVal rawCmd As Variant, ByVal suggestion As String)
    If Target Is Nothing Then Exit Sub

    Dim body As String
    body = BuildErrorMessage(rawCmd, suggestion)

    If Target.Comment Is Nothing Then
        Target.AddComment Text:=body
    ElseIf IsAssistNote(Target.Comment) Then
        Target.Comment.Text Text:=body
    Else
        Application.StatusBar = Replace(body, vbCrLf, " ")
    End If
End Sub

Public Sub ClearAssistNote(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Target.Comment Is Nothing Then Exit Sub
    If IsAssistNote(Target.Comment) Then Target.Comment.Delete
End Sub

Public Sub ClearAssistState(ByVal Target As Range)
    ClearAssistNote Target
End Sub

Public Function IsBlankCommand(ByVal rawCmd As Variant) As Boolean
    IsBlankCommand = (Len(NormalizeCmd(rawCmd)) = 0)
End Function

Public Function IsCommentCommand(ByVal rawCmd As Variant) As Boolean
    Dim cmd As String
    cmd = NormalizeCmd(rawCmd)
    IsCommentCommand = (Len(cmd) > 0 And Left$(cmd, 1) = "#")
End Function

Public Function NormalizeCmd(ByVal rawCmd As Variant) As String
    Dim text As String
    text = SafeText(rawCmd)
    NormalizeCmd = UCase$(Trim$(Replace(Trim$(text), ">", "")))
End Function

Private Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

Private Function IsAssistNote(ByVal note As Comment) As Boolean
    If note Is Nothing Then Exit Function
    IsAssistNote = (Left$(note.Text, Len(ASSIST_PREFIX)) = ASSIST_PREFIX)
End Function

Private Function BuildErrorMessage(ByVal rawCmd As Variant, ByVal suggestion As String) As String
    Dim cmd As String
    cmd = NormalizeCmd(rawCmd)

    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(cmd)

    If def Is Nothing Then
        BuildErrorMessage = ASSIST_PREFIX & vbCrLf & "Unknown command: " & cmd
    Else
        BuildErrorMessage = ASSIST_PREFIX & vbCrLf & "Invalid command: " & cmd
    End If

    If Len(suggestion) > 0 Then
        BuildErrorMessage = BuildErrorMessage & vbCrLf & "Did you mean " & suggestion & "?"
    End If
End Function

Private Function FormatParamHint(ByVal def As CmdDef) As String
    Dim parts As Collection
    Set parts = New Collection
    parts.Add def.Canonical

    Dim idx As Long
    For idx = 1 To 3
        Dim spec As ParamSpec
        Set spec = def.Param(idx)
        If Not spec Is Nothing Then
            parts.Add "P" & idx & ": " & spec.Meaning & " (" & DataTypeLabel(spec.DataType) & IIf(spec.Required, ", required", "") & ")"
        End If
    Next idx

    Dim result As String
    For idx = 1 To parts.count
        If Len(result) > 0 Then result = result & " | "
        result = result & CStr(parts(idx))
    Next idx

    FormatParamHint = result
End Function

Private Function DataTypeLabel(ByVal dataType As eDataType) As String
    Select Case dataType
        Case DT_NUM
            DataTypeLabel = "num"
        Case Else
            DataTypeLabel = "str"
    End Select
End Function

Private Function LevenshteinDistance(ByVal leftText As String, ByVal rightText As String) As Long
    Dim leftValue As String
    Dim rightValue As String
    leftValue = UCase$(Trim$(leftText))
    rightValue = UCase$(Trim$(rightText))

    Dim leftLen As Long
    Dim rightLen As Long
    leftLen = Len(leftValue)
    rightLen = Len(rightValue)

    If leftLen = 0 Then
        LevenshteinDistance = rightLen
        Exit Function
    End If
    If rightLen = 0 Then
        LevenshteinDistance = leftLen
        Exit Function
    End If
    If leftValue = rightValue Then
        LevenshteinDistance = 0
        Exit Function
    End If

    Dim dist() As Long
    ReDim dist(0 To leftLen, 0 To rightLen) As Long

    Dim i As Long
    Dim j As Long
    For i = 0 To leftLen
        dist(i, 0) = i
    Next i
    For j = 0 To rightLen
        dist(0, j) = j
    Next j

    For i = 1 To leftLen
        For j = 1 To rightLen
            Dim cost As Long
            If Mid$(leftValue, i, 1) = Mid$(rightValue, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            dist(i, j) = Min3Long(dist(i - 1, j) + 1, _
                         dist(i, j - 1) + 1, _
                         dist(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = dist(leftLen, rightLen)
End Function

Private Function Min3Long(ByVal firstValue As Long, ByVal secondValue As Long, ByVal thirdValue As Long) As Long
    Min3Long = firstValue
    If secondValue < Min3Long Then Min3Long = secondValue
    If thirdValue < Min3Long Then Min3Long = thirdValue
End Function

Private Function SafeText(ByVal value As Variant) As String
    On Error GoTo SafeExit
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    SafeText = CStr(value)
SafeExit:
End Function
