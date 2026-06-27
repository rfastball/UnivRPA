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

    If IsNumeric(trimmedTarget) Then
        Dim lineNo As Long
        lineNo = CLng(val(trimmedTarget))
        If lineNo >= 1 And lineNo <= m_TotalLines Then
            ResolveJumpTarget = lineNo
        Else
            ResolveJumpTarget = LABEL_NOT_FOUND
        End If
    Else
        ResolveJumpTarget = LabelLine(trimmedTarget)
    End If
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
