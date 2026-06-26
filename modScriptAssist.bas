Attribute VB_Name = "modScriptAssist"
' @desc Script Assist: COMMAND validation, parameter hints, typo suggestions, and Assist-owned notes
' @depends modCmdRegistry, modDevTools, CmdDef, ParamSpec
Option Explicit

Private Const ASSIST_PREFIX As String = "[ScriptAssist]"
Private Const CMD_COL As Long = 2
Private Const COMMENT_PREFIX As String = "# "

Public Function ResolveCmdCol(ByVal Sh As Worksheet) As Long
    If IsScriptSheet(Sh) Then ResolveCmdCol = CMD_COL
End Function

Public Function IsScriptSheet(ByVal Sh As Worksheet) As Boolean
    On Error GoTo SafeExit
    IsScriptSheet = (UCase$(Trim$(CStr(Sh.Cells(1, 1).value))) = "ROW" _
                  And UCase$(Trim$(CStr(Sh.Cells(1, 2).value))) = "COMMAND")
SafeExit:
End Function

Public Sub SetupScriptSheet(ByVal Sh As Worksheet)
    modDevTools.EnsureHeaders Sh
End Sub

Public Sub RegisterScriptAssistHotkeys()
    Application.OnKey "^+/", "ToggleSelectedCommandComment"
    Application.OnKey "^+F", "FormatRPAScript"
    Application.OnKey "^+W", "RunRPA"
    Application.OnKey "^+E", "StartHUD"
End Sub

Public Sub UnregisterScriptAssistHotkeys()
    Application.OnKey "^+/"
    Application.OnKey "^+F"
    Application.OnKey "^+W"
    Application.OnKey "^+E"
End Sub

Public Sub ToggleSelectedCommandComment()
    Dim activeObject As Object
    Set activeObject = ActiveSheet

    If Not TypeOf activeObject Is Worksheet Then
        Application.StatusBar = "Script Assist: worksheet에서만 사용할 수 있습니다."
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = activeObject

    Dim cmdCol As Long
    cmdCol = ResolveCmdCol(ws)
    If cmdCol = 0 Then
        Application.StatusBar = "Script Assist: 스크립트 시트에서만 사용할 수 있습니다."
        Exit Sub
    End If

    Dim selectedObject As Object
    Set selectedObject = Selection

    If Not TypeOf selectedObject Is Range Then
        Application.StatusBar = "Script Assist: 셀 범위를 선택해 주세요."
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = selectedObject

    If selectedRange.Areas.count <> 1 Then
        Application.StatusBar = "Script Assist: 연속된 범위만 주석 토글할 수 있습니다."
        Exit Sub
    End If

    Dim commandRange As Range
    Set commandRange = CommandRangeFromSelection(ws, selectedRange.Areas(1), cmdCol)
    If commandRange Is Nothing Then
        Application.StatusBar = "Script Assist: 토글할 COMMAND 셀이 없습니다."
        Exit Sub
    End If

    Dim cell As Range
    Dim targetCount As Long
    Dim allCommented As Boolean
    allCommented = True

    For Each cell In commandRange.Cells
        If Not IsBlankCommand(cell.value) Then
            targetCount = targetCount + 1
            If Not IsCommentCommand(cell.value) Then allCommented = False
        End If
    Next cell

    If targetCount = 0 Then
        Application.StatusBar = "Script Assist: 토글할 COMMAND 셀이 없습니다."
        Exit Sub
    End If

    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim changedCount As Long
    Dim failNumber As Long
    Dim failSource As String
    Dim failDescription As String
    Dim failHelpFile As String
    Dim failHelpContext As Long

    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    On Error GoTo CleanFail
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each cell In commandRange.Cells
        If Not IsBlankCommand(cell.value) Then
            If allCommented Then
                cell.value = UncommentCommandText(cell.value)
            ElseIf Not IsCommentCommand(cell.value) Then
                cell.value = CommentCommandText(cell.value)
            End If
            ValidateCommandCell cell
            changedCount = changedCount + 1
        End If
    Next cell

CleanExit:
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    If changedCount > 0 Then
        If allCommented Then
            Application.StatusBar = "Script Assist: " & changedCount & "줄 주석 해제"
        Else
            Application.StatusBar = "Script Assist: " & changedCount & "줄 주석 처리"
        End If
    End If
    Exit Sub

CleanFail:
    failNumber = Err.Number
    failSource = Err.Source
    failDescription = Err.Description
    failHelpFile = Err.HelpFile
    failHelpContext = Err.HelpContext
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    On Error GoTo 0
    Err.Raise failNumber, failSource, failDescription, failHelpFile, failHelpContext
End Sub

Public Sub ValidateCommandCell(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Target.Cells.count <> 1 Then Exit Sub
    If ResolveCmdCol(Target.Worksheet) <> Target.Column Then Exit Sub

    Dim rawCmd As Variant
    rawCmd = Target.value

    If IsBlankCommand(rawCmd) Or IsCommentCommand(rawCmd) Then
        ClearAssistNote Target
        ClearStatusBar
        Exit Sub
    End If

    Dim cmd As String
    cmd = NormalizeCmd(rawCmd)

    If modDevTools.ValidateLine(cmd, Target.offset(0, 1).value, Target.offset(0, 2).value, Target.offset(0, 3).value) _
       And modDevTools.ValidateSheetFlowAt(Target.Worksheet, Target.Row) Then
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

    Dim cmdCell As Range
    Set cmdCell = AssistCommandCell(Target)
    If cmdCell Is Nothing Then
        ClearStatusBar
        Exit Sub
    End If

    If IsBlankCommand(cmdCell.value) Or IsCommentCommand(cmdCell.value) Then
        ClearStatusBar
        Exit Sub
    End If

    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(cmdCell.value)
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
        Target.AddComment text:=body
    ElseIf IsAssistNote(Target.Comment) Then
        Target.Comment.text text:=body
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
    IsBlankCommand = (Len(Trim$(SafeText(rawCmd))) = 0)
End Function

Public Function IsCommentCommand(ByVal rawCmd As Variant) As Boolean
    Dim text As String
    text = Trim$(SafeText(rawCmd))
    IsCommentCommand = (Len(text) > 0 And Left$(text, 1) = "#")
End Function

Public Function NormalizeCmd(ByVal rawCmd As Variant) As String
    NormalizeCmd = modCmdRegistry.NormalizeCommand(rawCmd)
End Function

Private Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

Private Function AssistCommandCell(ByVal Target As Range) As Range
    If Target Is Nothing Then Exit Function
    If Target.Cells.count <> 1 Then Exit Function
    If Target.Row <= 1 Then Exit Function

    Dim cmdCol As Long
    cmdCol = ResolveCmdCol(Target.Worksheet)
    If cmdCol = 0 Then Exit Function

    If Target.Column < cmdCol Or Target.Column > cmdCol + 3 Then Exit Function
    Set AssistCommandCell = Target.Worksheet.Cells(Target.Row, cmdCol)
End Function

Private Function CommandRangeFromSelection(ByVal ws As Worksheet, ByVal selectedRange As Range, ByVal cmdCol As Long) As Range
    Dim firstRow As Long
    Dim lastSelectedRow As Long
    Dim lastCommandRow As Long
    Dim lastRow As Long

    firstRow = selectedRange.Row
    If firstRow < 2 Then firstRow = 2

    lastSelectedRow = selectedRange.Row + selectedRange.Rows.count - 1
    lastCommandRow = ws.Cells(ws.Rows.count, cmdCol).End(xlUp).Row
    lastRow = lastSelectedRow
    If lastRow > lastCommandRow Then lastRow = lastCommandRow

    If firstRow > lastRow Then Exit Function
    Set CommandRangeFromSelection = ws.Range(ws.Cells(firstRow, cmdCol), ws.Cells(lastRow, cmdCol))
End Function

Private Function CommentCommandText(ByVal rawCmd As Variant) As String
    CommentCommandText = SafeText(rawCmd)
    If IsBlankCommand(CommentCommandText) Or IsCommentCommand(CommentCommandText) Then Exit Function
    CommentCommandText = COMMENT_PREFIX & CommentCommandText
End Function

Private Function UncommentCommandText(ByVal rawCmd As Variant) As String
    Dim text As String
    text = SafeText(rawCmd)
    UncommentCommandText = text
    If Not IsCommentCommand(text) Then Exit Function

    Dim markerPos As Long
    markerPos = CommentMarkerPos(text)
    If markerPos = 0 Then Exit Function

    text = Left$(text, markerPos - 1) & Mid$(text, markerPos + 1)
    If Mid$(text, markerPos, 1) = " " Then
        text = Left$(text, markerPos - 1) & Mid$(text, markerPos + 1)
    End If
    UncommentCommandText = text
End Function

Private Function CommentMarkerPos(ByVal text As String) As Long
    Dim idx As Long
    For idx = 1 To Len(text)
        Select Case Mid$(text, idx, 1)
            Case " ", vbTab, ">"
                ' keep scanning through visual indentation
            Case "#"
                CommentMarkerPos = idx
                Exit Function
            Case Else
                Exit Function
        End Select
    Next idx
End Function

Private Function IsAssistNote(ByVal note As Comment) As Boolean
    If note Is Nothing Then Exit Function
    IsAssistNote = (Left$(note.text, Len(ASSIST_PREFIX)) = ASSIST_PREFIX)
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

Private Function DataTypeLabel(ByVal DataType As eDataType) As String
    Select Case DataType
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
