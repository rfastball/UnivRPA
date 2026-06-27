VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScriptProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const COL_ROW As Long = 1
Private Const COL_CMD As Long = 2
Private Const COL_P1 As Long = 3
Private Const COL_P2 As Long = 4
Private Const COL_P3 As Long = 5

Private Const LINE_NOT_FOUND As Long = CMD_ERROR
Private Const END_SELECT_NOT_FOUND As Long = CMD_ERROR
Private Const LABEL_NOT_FOUND As Long = CMD_ERROR

Private m_lines As Variant
Private m_TotalLines As Long
Private m_firstSheetRow As Long
Private m_selectBlocks As Object
Private m_caseEndSelectLines As Object
Private m_labels As Object
Private m_duplicateLabels As Object
Private m_validationErrors As Collection
Private m_instructions As Collection
Private m_loopDepths() As Long
Private m_subAreaIds() As Long
Private m_loopRanges As Collection

Private Sub Class_Initialize()
    m_firstSheetRow = 1
    ResetIndexes
End Sub

Private Sub ResetIndexes()
    Set m_selectBlocks = CreateObject("Scripting.Dictionary")
    Set m_caseEndSelectLines = CreateObject("Scripting.Dictionary")
    Set m_labels = CreateObject("Scripting.Dictionary")
    Set m_duplicateLabels = CreateObject("Scripting.Dictionary")
    Set m_validationErrors = New Collection
    Set m_instructions = New Collection
    Set m_loopRanges = New Collection
    m_labels.CompareMode = vbTextCompare
    m_duplicateLabels.CompareMode = vbTextCompare
End Sub

Public Sub LoadFromRange(ByVal sourceRange As Range)
    If sourceRange Is Nothing Then Err.Raise 5, "clsScriptProgram.LoadFromRange", "Source range is required."
    LoadFromArray sourceRange.value, sourceRange.Row
End Sub

Public Sub LoadFromArray(ByVal sourceArray As Variant, Optional ByVal firstSheetRow As Long = 1)
    If Not IsArray(sourceArray) Then Err.Raise 5, "clsScriptProgram.LoadFromArray", "Program source must be a 2D array."
    m_lines = sourceArray
    m_TotalLines = UBound(m_lines, 1)
    m_firstSheetRow = firstSheetRow
    ResetIndexes
    BuildInstructionCache
    BuildStateCoordinates
End Sub

Public Property Get totalLines() As Long
    totalLines = m_TotalLines
End Property

Public Function RawCmdAt(ByVal lineNo As Long) As String
    RawCmdAt = ValueAt(lineNo, COL_CMD)
End Function

Public Function CmdAt(ByVal lineNo As Long) As String
    CmdAt = modCmdRegistry.NormalizeCommand(RawCmdAt(lineNo))
End Function

Public Function P1At(ByVal lineNo As Long) As String
    P1At = ValueAt(lineNo, COL_P1)
End Function

Public Function P2At(ByVal lineNo As Long) As String
    P2At = ValueAt(lineNo, COL_P2)
End Function

Public Function P3At(ByVal lineNo As Long) As String
    P3At = ValueAt(lineNo, COL_P3)
End Function

Public Function IsBlankLine(ByVal lineNo As Long) As Boolean
    IsBlankLine = (Len(Trim$(RawCmdAt(lineNo))) = 0)
End Function

Public Function IsCommentLine(ByVal lineNo As Long) As Boolean
    Dim raw As String
    raw = Trim$(RawCmdAt(lineNo))
    IsCommentLine = (Len(raw) > 0 And Left$(raw, 1) = "#")
End Function

Public Function IsSkippable(ByVal lineNo As Long) As Boolean
    IsSkippable = IsBlankLine(lineNo) Or IsCommentLine(lineNo)
End Function

Public Function SheetRowFromRuntimeLine(ByVal lineNo As Long) As Long
    ValidateLineRange lineNo
    SheetRowFromRuntimeLine = m_firstSheetRow + lineNo - 1
End Function

Public Function RuntimeLineFromSheetRow(ByVal sheetRow As Long) As Long
    RuntimeLineFromSheetRow = sheetRow - m_firstSheetRow + 1
    ValidateLineRange RuntimeLineFromSheetRow
End Function

Public Function InstructionAt(ByVal lineNo As Long) As clsInstruction
    ValidateLineRange lineNo
    Set InstructionAt = m_instructions.item(lineNo)
End Function

Public Function FindBlockPair(ByVal startLine As Long, ByVal group As String, Optional ByVal Direction As Long = 1) As Long
    On Error GoTo ErrHandler
    ValidateLineRange startLine

    group = UCase$(Trim$(group))
    If group <> "IF" And group <> "WHILE" And group <> "REPEAT" And group <> "FOREACH" Then
        FindBlockPair = LINE_NOT_FOUND
        Exit Function
    End If
    If Direction <> 1 And Direction <> -1 Then
        FindBlockPair = LINE_NOT_FOUND
        Exit Function
    End If

    Dim R As Long
    Dim endLine As Long
    Dim depth As Long
    If Direction = 1 Then endLine = m_TotalLines Else endLine = 1

    For R = startLine + Direction To endLine Step Direction
        If IsSkippable(R) Then GoTo NextRow

        Dim def As CmdDef
        Set def = modCmdRegistry.GetCmdDef(CmdAt(R))
        If def Is Nothing Then GoTo NextRow
        If UCase$(def.BlockGroup) <> group Then GoTo NextRow

        If Direction = 1 Then
            Select Case def.BlockRole
                Case STARTER
                    depth = depth + 1
                Case MIDBLOCK
                    If group = "IF" And depth = 0 Then
                        FindBlockPair = R
                        Exit Function
                    End If
                Case ENDER
                    If depth = 0 Then
                        FindBlockPair = R
                        Exit Function
                    End If
                    depth = depth - 1
            End Select
        Else
            Select Case def.BlockRole
                Case ENDER
                    depth = depth + 1
                Case STARTER
                    If depth = 0 Then
                        FindBlockPair = R
                        Exit Function
                    End If
                    depth = depth - 1
            End Select
        End If

NextRow:
    Next R

    FindBlockPair = LINE_NOT_FOUND
    Exit Function
ErrHandler:
    FindBlockPair = LINE_NOT_FOUND
End Function

Public Sub BuildStructureIndex()
    Set m_selectBlocks = CreateObject("Scripting.Dictionary")
    Set m_caseEndSelectLines = CreateObject("Scripting.Dictionary")

    Dim R As Long
    For R = 1 To m_TotalLines
        If CmdAt(R) = "SELECT" Then BuildSelectBlock R
    Next R
End Sub

Public Function SelectBlock(ByVal SelectLine As Long) As clsSelectBlock
    If m_selectBlocks.exists(CStr(SelectLine)) Then Set SelectBlock = m_selectBlocks(CStr(SelectLine))
End Function

Public Function EndSelectOfCase(ByVal caseLine As Long) As Long
    If m_caseEndSelectLines.exists(CStr(caseLine)) Then
        EndSelectOfCase = CLng(m_caseEndSelectLines(CStr(caseLine)))
    Else
        EndSelectOfCase = END_SELECT_NOT_FOUND
    End If
End Function

Public Sub BuildLabelIndex()
    Set m_labels = CreateObject("Scripting.Dictionary")
    Set m_duplicateLabels = CreateObject("Scripting.Dictionary")
    m_labels.CompareMode = vbTextCompare
    m_duplicateLabels.CompareMode = vbTextCompare

    Dim R As Long
    For R = 1 To m_TotalLines
        If CmdAt(R) = "LABEL" Then
            Dim labelName As String
            labelName = NormalizeLabelName(P3At(R))
            If Len(labelName) = 0 Or ContainsTokenSyntax(P3At(R)) Then GoTo NextRow

            If m_labels.exists(labelName) Then
                m_duplicateLabels(labelName) = True
            Else
                m_labels.Add labelName, R
            End If
        End If
NextRow:
    Next R
End Sub

Public Function HasLabel(ByVal name As String) As Boolean
    HasLabel = m_labels.exists(NormalizeLabelName(name))
End Function

Public Function LabelLine(ByVal name As String) As Long
    Dim key As String
    key = NormalizeLabelName(name)
    If m_labels.exists(key) Then LabelLine = CLng(m_labels(key)) Else LabelLine = LABEL_NOT_FOUND
End Function

Public Function ResolveJumpTarget(ByVal Target As String) As Long
    Dim trimmedTarget As String
    trimmedTarget = Trim$(Target)

    If Len(trimmedTarget) = 0 Or IsNumeric(trimmedTarget) Or ContainsTokenSyntax(trimmedTarget) Then
        ResolveJumpTarget = LABEL_NOT_FOUND
    Else
        ResolveJumpTarget = LabelLine(trimmedTarget)
    End If
End Function

Public Function LoopDepthAt(ByVal lineNo As Long) As Long
    ValidateLineRange lineNo
    LoopDepthAt = m_loopDepths(lineNo)
End Function

Public Function SubAreaIdAt(ByVal lineNo As Long) As Long
    ValidateLineRange lineNo
    SubAreaIdAt = m_subAreaIds(lineNo)
End Function

Public Function IsGotoAllowed(ByVal sourceLine As Long, ByVal targetLine As Long) As Boolean
    ValidateLineRange sourceLine
    ValidateLineRange targetLine

    If SubAreaIdAt(sourceLine) <> SubAreaIdAt(targetLine) Then Exit Function
    IsGotoAllowed = IsLoopTransferAllowed(sourceLine, targetLine)
End Function

Public Function HasDuplicateLabels() As Boolean
    HasDuplicateLabels = (m_duplicateLabels.count > 0)
End Function

Public Function DuplicateLabels() As String
    DuplicateLabels = JoinDictionaryKeys(m_duplicateLabels)
End Function

Public Function ValidateStructure() As Boolean
    ClearValidationErrors
    BuildStructureIndex

    Dim key As Variant
    For Each key In m_selectBlocks.Keys
        Dim block As clsSelectBlock
        Set block = m_selectBlocks(key)
        If block.EndSelectLine <= 0 Then AddValidationError "SELECT line " & key & " has no END_SELECT."
    Next key

    ValidateStructure = (m_validationErrors.count = 0)
End Function

Public Function ValidateLabels() As Boolean
    BuildLabelIndex

    Dim R As Long
    For R = 1 To m_TotalLines
        If CmdAt(R) = "LABEL" Then
            Dim rawLabel As String
            rawLabel = P3At(R)
            If Len(NormalizeLabelName(rawLabel)) = 0 Then
                AddValidationError "LABEL line " & R & " has empty name."
            ElseIf ContainsTokenSyntax(rawLabel) Then
                AddValidationError "LABEL line " & R & " must be a static symbol."
            End If
        End If
    Next R

    If HasDuplicateLabels Then AddValidationError "Duplicate LABEL: " & DuplicateLabels()
    ValidateJumpTargets
    ValidateLabels = (m_validationErrors.count = 0)
End Function

Public Function ValidationErrors() As String
    Dim i As Long
    For i = 1 To m_validationErrors.count
        If Len(ValidationErrors) > 0 Then ValidationErrors = ValidationErrors & vbCrLf
        ValidationErrors = ValidationErrors & CStr(m_validationErrors.item(i))
    Next i
End Function

Public Function ValidateLineRange(ByVal lineNo As Long) As Boolean
    If lineNo < 1 Or lineNo > m_TotalLines Then Err.Raise 9, "clsScriptProgram.ValidateLineRange", "Runtime line out of range: " & lineNo
    ValidateLineRange = True
End Function

Private Function ValueAt(ByVal lineNo As Long, ByVal columnNo As Long) As String
    ValidateLineRange lineNo
    If columnNo > UBound(m_lines, 2) Then Exit Function
    If IsError(m_lines(lineNo, columnNo)) Or IsNull(m_lines(lineNo, columnNo)) Or IsEmpty(m_lines(lineNo, columnNo)) Then Exit Function
    ValueAt = CStr(m_lines(lineNo, columnNo))
End Function

Private Sub BuildInstructionCache()
    Set m_instructions = New Collection

    Dim R As Long
    For R = 1 To m_TotalLines
        Dim inst As clsInstruction
        Set inst = New clsInstruction
        inst.Init R, SheetRowFromRuntimeLine(R), RawCmdAt(R), CmdAt(R), P1At(R), P2At(R), P3At(R), modCmdRegistry.GetCmdDef(CmdAt(R))
        m_instructions.Add inst
    Next R
End Sub

Private Sub BuildStateCoordinates()
    ReDim m_loopDepths(1 To m_TotalLines)
    ReDim m_subAreaIds(1 To m_TotalLines)
    Set m_loopRanges = New Collection

    BuildLoopCoordinates
    BuildSubAreaCoordinates
End Sub

Private Sub BuildLoopCoordinates()
    Dim depth As Long
    Dim R As Long

    For R = 1 To m_TotalLines
        Dim cmd As String
        cmd = CmdAt(R)

        If IsLoopEnder(cmd) Then depth = IIf(depth > 0, depth - 1, 0)
        m_loopDepths(R) = depth

        If IsLoopStarter(cmd) Then
            Dim endLine As Long
            endLine = FindBlockPair(R, LoopGroupOf(cmd), 1)
            If endLine > 0 Then m_loopRanges.Add Array(R, endLine)
            depth = depth + 1
        End If
    Next R
End Sub

Private Sub BuildSubAreaCoordinates()
    Dim labelRows As Object
    Set labelRows = CreateObject("Scripting.Dictionary")
    labelRows.CompareMode = vbTextCompare

    Dim callTargets As Object
    Set callTargets = CreateObject("Scripting.Dictionary")
    callTargets.CompareMode = vbTextCompare

    Dim R As Long
    For R = 1 To m_TotalLines
        Dim cmd As String
        cmd = CmdAt(R)
        If cmd = "LABEL" Then
            Dim labelName As String
            labelName = NormalizeLabelName(P3At(R))
            If Len(labelName) > 0 And Not ContainsTokenSyntax(P3At(R)) Then
                If Not labelRows.exists(labelName) Then labelRows.Add labelName, R
            End If
        ElseIf cmd = "CALL" Then
            Dim targetName As String
            targetName = NormalizeLabelName(P3At(R))
            If IsStaticLabelSymbol(targetName) Then callTargets(targetName) = True
        End If
    Next R

    Dim areaId As Long
    Dim key As Variant
    For Each key In callTargets.Keys
        If labelRows.exists(CStr(key)) Then
            areaId = areaId + 1
            Dim startLine As Long
            Dim endLine As Long
            startLine = CLng(labelRows(CStr(key)))
            endLine = FindSubroutineReturnLine(startLine)
            If endLine <= 0 Then endLine = startLine
            MarkSubArea startLine, endLine, areaId
        End If
    Next key
End Sub

Private Function FindSubroutineReturnLine(ByVal startLine As Long) As Long
    Dim labelDepth As Long
    labelDepth = LoopDepthAt(startLine)

    Dim R As Long
    For R = startLine + 1 To m_TotalLines
        If IsSkippable(R) Then GoTo NextRow

        Dim cmd As String
        cmd = CmdAt(R)
        If Len(cmd) = 0 Then GoTo NextRow
        If LoopDepthAt(R) < labelDepth Then Exit Function

        If LoopDepthAt(R) = labelDepth Then
            Select Case cmd
                Case "RETURN"
                    FindSubroutineReturnLine = R
                    Exit Function
                Case "LABEL"
                    Exit Function
            End Select
        End If

NextRow:
    Next R
End Function

Private Sub MarkSubArea(ByVal startLine As Long, ByVal endLine As Long, ByVal areaId As Long)
    Dim R As Long
    For R = startLine To endLine
        If m_subAreaIds(R) = 0 Then m_subAreaIds(R) = areaId
    Next R
End Sub

Private Sub BuildSelectBlock(ByVal SelectLine As Long)
    If m_selectBlocks.exists(CStr(SelectLine)) Then Exit Sub

    Dim block As clsSelectBlock
    Set block = New clsSelectBlock
    block.SelectLine = SelectLine

    Dim depth As Long
    Dim R As Long
    For R = SelectLine + 1 To m_TotalLines
        If IsSkippable(R) Then GoTo NextRow

        Dim def As CmdDef
        Set def = modCmdRegistry.GetCmdDef(CmdAt(R))
        If def Is Nothing Then GoTo NextRow
        If UCase$(def.BlockGroup) <> "SELECT" Then GoTo NextRow

        Select Case def.BlockRole
            Case STARTER
                depth = depth + 1
            Case MIDBLOCK
                If depth = 0 Then block.AddCaseLine R
            Case ENDER
                If depth = 0 Then
                    block.EndSelectLine = R
                    Exit For
                End If
                depth = depth - 1
        End Select

NextRow:
    Next R

    m_selectBlocks.Add CStr(SelectLine), block
    If block.EndSelectLine > 0 Then
        Dim i As Long
        For i = 1 To block.CaseCount
            m_caseEndSelectLines(CStr(block.CaseLineAt(i))) = block.EndSelectLine
        Next i
    End If
End Sub

Private Function NormalizeLabelName(ByVal name As String) As String
    NormalizeLabelName = UCase$(Trim$(name))
End Function

Private Function ContainsTokenSyntax(ByVal value As String) As Boolean
    ContainsTokenSyntax = (InStr(1, value, "{", vbTextCompare) > 0 _
                        Or InStr(1, value, "}", vbTextCompare) > 0 _
                        Or InStr(1, value, "$", vbTextCompare) > 0)
End Function

Private Function IsStaticLabelSymbol(ByVal value As String) As Boolean
    value = Trim$(value)
    IsStaticLabelSymbol = (Len(value) > 0 And Not IsNumeric(value) And Not ContainsTokenSyntax(value))
End Function

Private Sub ValidateJumpTargets()
    Dim R As Long
    For R = 1 To m_TotalLines
        Dim cmd As String
        cmd = CmdAt(R)
        If cmd <> "GOTO" And cmd <> "CALL" Then GoTo NextRow

        Dim rawTarget As String
        rawTarget = Trim$(P3At(R))

        If Len(rawTarget) = 0 Then
            AddValidationError cmd & " line " & R & " has empty target."
        ElseIf IsNumeric(rawTarget) Then
            AddValidationError cmd & " line " & R & " must target a LABEL, not a line number."
        ElseIf ContainsTokenSyntax(rawTarget) Then
            AddValidationError cmd & " line " & R & " must target a static LABEL."
        ElseIf Not HasLabel(rawTarget) Then
            AddValidationError cmd & " line " & R & " target not found: " & rawTarget
        ElseIf cmd = "GOTO" Then
            Dim targetLine As Long
            targetLine = LabelLine(rawTarget)
            If targetLine > 0 Then
                If Not IsGotoAllowed(R, targetLine) Then AddValidationError "GOTO line " & R & " crosses invalid flow boundary."
            End If
        End If

NextRow:
    Next R
End Sub

Private Function IsLoopStarter(ByVal cmd As String) As Boolean
    IsLoopStarter = (cmd = "WHILE" Or cmd = "REPEAT" Or cmd = "FOREACH_ROW")
End Function

Private Function IsLoopEnder(ByVal cmd As String) As Boolean
    IsLoopEnder = (cmd = "END_WHILE" Or cmd = "END_REPEAT" Or cmd = "END_FOREACH")
End Function

Private Function LoopGroupOf(ByVal cmd As String) As String
    Select Case cmd
        Case "WHILE": LoopGroupOf = "WHILE"
        Case "REPEAT": LoopGroupOf = "REPEAT"
        Case "FOREACH_ROW": LoopGroupOf = "FOREACH"
    End Select
End Function

Private Function IsLoopTransferAllowed(ByVal sourceLine As Long, ByVal targetLine As Long) As Boolean
    Dim targetRange As Variant
    targetRange = InnermostLoopRangeAt(targetLine)

    If IsEmpty(targetRange) Then
        IsLoopTransferAllowed = True
    Else
        IsLoopTransferAllowed = (sourceLine >= CLng(targetRange(0)) And sourceLine <= CLng(targetRange(1)))
    End If
End Function

Private Function InnermostLoopRangeAt(ByVal lineNo As Long) As Variant
    Dim best As Variant
    Dim bestSize As Long
    Dim item As Variant

    bestSize = 0
    For Each item In m_loopRanges
        If lineNo >= CLng(item(0)) And lineNo <= CLng(item(1)) Then
            Dim size As Long
            size = CLng(item(1)) - CLng(item(0))
            If bestSize = 0 Or size < bestSize Then
                best = item
                bestSize = size
            End If
        End If
    Next item

    If bestSize > 0 Then InnermostLoopRangeAt = best
End Function


Private Sub ClearValidationErrors()
    Set m_validationErrors = New Collection
End Sub

Private Sub AddValidationError(ByVal message As String)
    m_validationErrors.Add message
End Sub

Private Function JoinDictionaryKeys(ByVal dict As Object) As String
    Dim key As Variant
    For Each key In dict.Keys
        If Len(JoinDictionaryKeys) > 0 Then JoinDictionaryKeys = JoinDictionaryKeys & ", "
        JoinDictionaryKeys = JoinDictionaryKeys & CStr(key)
    Next key
End Function
