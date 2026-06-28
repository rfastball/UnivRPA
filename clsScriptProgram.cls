VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScriptProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 정적 스크립트 의미 모델/분석 SSOT. 구조·라벨·흐름·변수 판단은 이 클래스에 질의한다.
' @depends modCmdRegistry, CmdDef, ParamSpec, clsInstruction, clsSelectBlock
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
Private m_labelRows As Object
Private m_duplicateLabels As Object
Private m_callLabelTargets As Object
Private m_gotoLabelTargets As Object
Private m_callSourceRows As Object
Private m_labelRoles As Object
Private m_labelRoleWarnings As Object
Private m_subroutineStartRows As Object
Private m_subroutineEndRows As Object
Private m_flowInvalidRows As Object
Private m_variableErrors As Object
Private m_variableWarnings As Object
Private m_validationErrors As Collection
Private m_instructions As Collection
Private m_loopDepths() As Long
Private m_subAreaIds() As Long
Private m_displayDepths() As Long
Private m_styleDepths() As Long
Private m_subDepths() As Long
Private m_loopRanges As Collection

Private Sub Class_Initialize()
    m_firstSheetRow = 1
    ResetIndexes
End Sub

Private Sub ResetIndexes()
    Set m_selectBlocks = CreateObject("Scripting.Dictionary")
    Set m_caseEndSelectLines = CreateObject("Scripting.Dictionary")
    Set m_labels = CreateObject("Scripting.Dictionary")
    Set m_labelRows = CreateObject("Scripting.Dictionary")
    Set m_duplicateLabels = CreateObject("Scripting.Dictionary")
    Set m_callLabelTargets = CreateObject("Scripting.Dictionary")
    Set m_gotoLabelTargets = CreateObject("Scripting.Dictionary")
    Set m_callSourceRows = CreateObject("Scripting.Dictionary")
    Set m_labelRoles = CreateObject("Scripting.Dictionary")
    Set m_labelRoleWarnings = CreateObject("Scripting.Dictionary")
    Set m_subroutineStartRows = CreateObject("Scripting.Dictionary")
    Set m_subroutineEndRows = CreateObject("Scripting.Dictionary")
    Set m_flowInvalidRows = CreateObject("Scripting.Dictionary")
    Set m_variableErrors = CreateObject("Scripting.Dictionary")
    Set m_variableWarnings = CreateObject("Scripting.Dictionary")
    Set m_validationErrors = New Collection
    Set m_instructions = New Collection
    Set m_loopRanges = New Collection
    m_labels.CompareMode = vbTextCompare
    m_labelRows.CompareMode = vbTextCompare
    m_duplicateLabels.CompareMode = vbTextCompare
    m_callLabelTargets.CompareMode = vbTextCompare
    m_gotoLabelTargets.CompareMode = vbTextCompare
    m_callSourceRows.CompareMode = vbTextCompare
    m_labelRoles.CompareMode = vbTextCompare
    m_labelRoleWarnings.CompareMode = vbTextCompare
    m_subroutineStartRows.CompareMode = vbTextCompare
    m_subroutineEndRows.CompareMode = vbTextCompare
    m_flowInvalidRows.CompareMode = vbTextCompare
    m_variableErrors.CompareMode = vbTextCompare
    m_variableWarnings.CompareMode = vbTextCompare
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
    BuildLabelReferenceIndex
    BuildStateCoordinates
    BuildStaticAnalysis
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
    BuildLabelReferenceIndex
    BuildStateCoordinates
    BuildStaticAnalysis
End Sub

Public Function LabelRoleAt(ByVal lineNo As Long) As String
    ValidateLineRange lineNo
    If m_labelRoles.exists(CStr(lineNo)) Then LabelRoleAt = CStr(m_labelRoles(CStr(lineNo)))
End Function

Public Function LabelRoleAtSheetRow(ByVal sheetRow As Long) As String
    LabelRoleAtSheetRow = LabelRoleAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function LabelRoleWarningAt(ByVal lineNo As Long) As String
    ValidateLineRange lineNo
    If m_labelRoleWarnings.exists(CStr(lineNo)) Then LabelRoleWarningAt = CStr(m_labelRoleWarnings(CStr(lineNo)))
End Function

Public Function LabelRoleWarningAtSheetRow(ByVal sheetRow As Long) As String
    LabelRoleWarningAtSheetRow = LabelRoleWarningAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsCallTargetLabelAt(ByVal lineNo As Long) As Boolean
    Dim roleName As String
    roleName = LabelRoleAt(lineNo)
    IsCallTargetLabelAt = (roleName = "SubroutineLabel" Or roleName = "MixedLabel")
End Function

Public Function IsCallTargetLabelAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsCallTargetLabelAtSheetRow = IsCallTargetLabelAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsJumpTargetMissingAt(ByVal lineNo As Long) As Boolean
    ValidateLineRange lineNo

    Dim cmd As String
    cmd = CmdAt(lineNo)
    If cmd <> "GOTO" And cmd <> "CALL" Then Exit Function

    Dim targetName As String
    targetName = NormalizeLabelName(P3At(lineNo))
    If IsStaticLabelSymbol(targetName) Then IsJumpTargetMissingAt = Not m_labels.exists(targetName)
End Function

Public Function IsJumpTargetMissingAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsJumpTargetMissingAtSheetRow = IsJumpTargetMissingAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function ResolveJumpTargetSheetRow(ByVal Target As String) As Long
    Dim lineNo As Long
    lineNo = ResolveJumpTarget(Target)
    If lineNo > 0 Then ResolveJumpTargetSheetRow = SheetRowFromRuntimeLine(lineNo)
End Function

Public Function FirstCallSourceLineForLabel(ByVal labelName As String) As Long
    Dim key As String
    key = NormalizeLabelName(labelName)
    If Not m_callSourceRows.exists(key) Then Exit Function

    Dim rows As Collection
    Set rows = m_callSourceRows(key)
    If rows.count > 0 Then FirstCallSourceLineForLabel = CLng(rows.item(1))
End Function

Public Function FirstCallSourceSheetRowForLabel(ByVal labelName As String) As Long
    Dim lineNo As Long
    lineNo = FirstCallSourceLineForLabel(labelName)
    If lineNo > 0 Then FirstCallSourceSheetRowForLabel = SheetRowFromRuntimeLine(lineNo)
End Function

Public Function DisplayDepthAt(ByVal lineNo As Long) As Long
    ValidateLineRange lineNo
    DisplayDepthAt = m_displayDepths(lineNo)
End Function

Public Function DisplayDepthAtSheetRow(ByVal sheetRow As Long) As Long
    DisplayDepthAtSheetRow = DisplayDepthAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function StyleDepthAt(ByVal lineNo As Long) As Long
    ValidateLineRange lineNo
    StyleDepthAt = m_styleDepths(lineNo)
End Function

Public Function StyleDepthAtSheetRow(ByVal sheetRow As Long) As Long
    StyleDepthAtSheetRow = StyleDepthAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsSubroutineStartAt(ByVal lineNo As Long) As Boolean
    ValidateLineRange lineNo
    IsSubroutineStartAt = m_subroutineStartRows.exists(CStr(lineNo))
End Function

Public Function IsSubroutineStartAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsSubroutineStartAtSheetRow = IsSubroutineStartAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsSubroutineEndAt(ByVal lineNo As Long) As Boolean
    ValidateLineRange lineNo
    IsSubroutineEndAt = m_subroutineEndRows.exists(CStr(lineNo))
End Function

Public Function IsSubroutineEndAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsSubroutineEndAtSheetRow = IsSubroutineEndAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsFlowInvalidAt(ByVal lineNo As Long) As Boolean
    ValidateLineRange lineNo
    IsFlowInvalidAt = m_flowInvalidRows.exists(CStr(lineNo))
End Function

Public Function IsFlowInvalidAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsFlowInvalidAtSheetRow = IsFlowInvalidAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function VariableErrorAt(ByVal lineNo As Long) As String
    ValidateLineRange lineNo
    If m_variableErrors.exists(CStr(lineNo)) Then VariableErrorAt = CStr(m_variableErrors(CStr(lineNo)))
End Function

Public Function VariableErrorAtSheetRow(ByVal sheetRow As Long) As String
    VariableErrorAtSheetRow = VariableErrorAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function VariableWarningAt(ByVal lineNo As Long) As String
    ValidateLineRange lineNo
    If m_variableWarnings.exists(CStr(lineNo)) Then VariableWarningAt = CStr(m_variableWarnings(CStr(lineNo)))
End Function

Public Function VariableWarningAtSheetRow(ByVal sheetRow As Long) As String
    VariableWarningAtSheetRow = VariableWarningAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Private Sub BuildLabelReferenceIndex()
    Set m_labels = CreateObject("Scripting.Dictionary")
    Set m_labelRows = CreateObject("Scripting.Dictionary")
    Set m_duplicateLabels = CreateObject("Scripting.Dictionary")
    Set m_callLabelTargets = CreateObject("Scripting.Dictionary")
    Set m_gotoLabelTargets = CreateObject("Scripting.Dictionary")
    Set m_callSourceRows = CreateObject("Scripting.Dictionary")
    Set m_labelRoles = CreateObject("Scripting.Dictionary")
    Set m_labelRoleWarnings = CreateObject("Scripting.Dictionary")
    m_labels.CompareMode = vbTextCompare
    m_labelRows.CompareMode = vbTextCompare
    m_duplicateLabels.CompareMode = vbTextCompare
    m_callLabelTargets.CompareMode = vbTextCompare
    m_gotoLabelTargets.CompareMode = vbTextCompare
    m_callSourceRows.CompareMode = vbTextCompare
    m_labelRoles.CompareMode = vbTextCompare
    m_labelRoleWarnings.CompareMode = vbTextCompare

    Dim R As Long
    For R = 1 To m_TotalLines
        If IsSkippable(R) Then GoTo NextRow

        Dim cmd As String
        cmd = CmdAt(R)

        If cmd = "LABEL" Then
            AddLabelReferenceRow R
        ElseIf cmd = "CALL" Then
            AddJumpReferenceTarget m_callLabelTargets, P3At(R)
            AddCallSourceRow P3At(R), R
        ElseIf cmd = "GOTO" Then
            AddJumpReferenceTarget m_gotoLabelTargets, P3At(R)
        End If
NextRow:
    Next R

    BuildLabelRoles
End Sub

Private Sub AddLabelReferenceRow(ByVal lineNo As Long)
    Dim labelName As String
    labelName = NormalizeLabelName(P3At(lineNo))
    If Len(labelName) = 0 Or ContainsTokenSyntax(P3At(lineNo)) Then Exit Sub

    Dim rows As Collection
    If m_labelRows.exists(labelName) Then
        Set rows = m_labelRows(labelName)
        m_duplicateLabels(labelName) = True
    Else
        Set rows = New Collection
        m_labelRows.Add labelName, rows
        m_labels.Add labelName, lineNo
    End If

    rows.Add lineNo
End Sub

Private Sub AddJumpReferenceTarget(ByVal targets As Object, ByVal rawTarget As String)
    Dim targetName As String
    targetName = NormalizeLabelName(rawTarget)
    If IsStaticLabelSymbol(targetName) Then targets(targetName) = True
End Sub

Private Sub AddCallSourceRow(ByVal rawTarget As String, ByVal lineNo As Long)
    Dim targetName As String
    targetName = NormalizeLabelName(rawTarget)
    If Not IsStaticLabelSymbol(targetName) Then Exit Sub

    Dim rows As Collection
    If m_callSourceRows.exists(targetName) Then
        Set rows = m_callSourceRows(targetName)
    Else
        Set rows = New Collection
        m_callSourceRows.Add targetName, rows
    End If
    rows.Add lineNo
End Sub

Private Sub BuildLabelRoles()
    Dim labelName As Variant
    For Each labelName In m_labelRows.Keys
        Dim isCallTarget As Boolean
        Dim isGotoTarget As Boolean
        isCallTarget = m_callLabelTargets.exists(CStr(labelName))
        isGotoTarget = m_gotoLabelTargets.exists(CStr(labelName))

        Dim roleName As String
        If isCallTarget And isGotoTarget Then
            roleName = "MixedLabel"
        ElseIf isCallTarget Then
            roleName = "SubroutineLabel"
        ElseIf isGotoTarget Then
            roleName = "JumpLabel"
        Else
            roleName = "DeadLabel"
        End If

        Dim rows As Collection
        Set rows = m_labelRows(CStr(labelName))

        Dim rowItem As Variant
        For Each rowItem In rows
            m_labelRoles(CStr(rowItem)) = roleName
            If roleName = "MixedLabel" Then
                m_labelRoleWarnings(CStr(rowItem)) = _
                    "Warning: MixedLabel '" & CStr(labelName) & "' is referenced by both CALL and GOTO."
            End If
        Next rowItem
    Next labelName
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
    Dim areaId As Long
    Dim key As Variant
    For Each key In m_callLabelTargets.Keys
        If m_labels.exists(CStr(key)) Then
            areaId = areaId + 1
            Dim startLine As Long
            Dim endLine As Long
            startLine = CLng(m_labels(CStr(key)))
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

Private Function IsFormatBlockStarter(ByVal cmd As String) As Boolean
    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(cmd)
    If Not def Is Nothing Then IsFormatBlockStarter = (def.BlockRole = STARTER)
End Function

Private Function IsFormatBlockEnder(ByVal cmd As String) As Boolean
    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(cmd)
    If Not def Is Nothing Then IsFormatBlockEnder = (def.BlockRole = ENDER)
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

Private Sub BuildStaticAnalysis()
    ReDim m_displayDepths(1 To m_TotalLines)
    ReDim m_styleDepths(1 To m_TotalLines)
    ReDim m_subDepths(1 To m_TotalLines)

    Set m_subroutineStartRows = CreateObject("Scripting.Dictionary")
    Set m_subroutineEndRows = CreateObject("Scripting.Dictionary")
    Set m_flowInvalidRows = CreateObject("Scripting.Dictionary")
    Set m_variableErrors = CreateObject("Scripting.Dictionary")
    Set m_variableWarnings = CreateObject("Scripting.Dictionary")
    m_subroutineStartRows.CompareMode = vbTextCompare
    m_subroutineEndRows.CompareMode = vbTextCompare
    m_flowInvalidRows.CompareMode = vbTextCompare
    m_variableErrors.CompareMode = vbTextCompare
    m_variableWarnings.CompareMode = vbTextCompare

    BuildDisplayDepths
    BuildSubroutineFlowAnalysis
    BuildVariableAnalysis
End Sub

Private Sub BuildDisplayDepths()
    Dim depth As Long
    Dim R As Long

    For R = 1 To m_TotalLines
        Dim cmd As String
        cmd = CmdAt(R)

        If IsFormatBlockEnder(cmd) Then depth = IIf(depth > 0, depth - 1, 0)

        m_styleDepths(R) = depth
        m_displayDepths(R) = depth
        If cmd = "ELSE" Or cmd = "CASE" Then m_displayDepths(R) = IIf(depth > 0, depth - 1, 0)

        If IsFormatBlockStarter(cmd) Then depth = depth + 1
    Next R
End Sub

Private Sub BuildSubroutineFlowAnalysis()
    MarkMissingJumpTargets

    Dim inSubroutineRows As Object
    Set inSubroutineRows = CreateObject("Scripting.Dictionary")
    inSubroutineRows.CompareMode = vbTextCompare

    Dim R As Long
    For R = 1 To m_TotalLines
        If IsCommentLine(R) Then GoTo NextRow

        If CmdAt(R) = "LABEL" Then
            If IsCallTargetLabelAt(R) Then
                m_subroutineStartRows(CStr(R)) = True

                Dim endLine As Long
                endLine = FindSubroutineReturnLineForFlow(R)

                If endLine > 0 Then
                    m_subroutineEndRows(CStr(endLine)) = True
                    inSubroutineRows(CStr(R)) = True
                    inSubroutineRows(CStr(endLine)) = True

                    Dim innerLine As Long
                    For innerLine = R + 1 To endLine - 1
                        m_subDepths(innerLine) = m_subDepths(innerLine) + 1
                        inSubroutineRows(CStr(innerLine)) = True
                    Next innerLine
                Else
                    m_flowInvalidRows(CStr(R)) = True
                End If

                If Not HasFlowBreakBefore(R) Then m_flowInvalidRows(CStr(R)) = True
            End If
        End If
NextRow:
    Next R

    For R = 1 To m_TotalLines
        If IsCommentLine(R) Then GoTo NextReturnRow
        If CmdAt(R) = "RETURN" Then
            If Not inSubroutineRows.exists(CStr(R)) Then m_flowInvalidRows(CStr(R)) = True
        End If
NextReturnRow:
    Next R

    For R = 1 To m_TotalLines
        m_displayDepths(R) = m_displayDepths(R) + m_subDepths(R)
        m_styleDepths(R) = m_styleDepths(R) + m_subDepths(R)
    Next R
End Sub

Private Sub MarkMissingJumpTargets()
    Dim R As Long
    For R = 1 To m_TotalLines
        If IsJumpTargetMissingAt(R) Then m_flowInvalidRows(CStr(R)) = True
    Next R
End Sub

Private Function FindSubroutineReturnLineForFlow(ByVal startLine As Long) As Long
    Dim labelDepth As Long
    labelDepth = m_displayDepths(startLine)

    Dim R As Long
    For R = startLine + 1 To m_TotalLines
        If IsCommentLine(R) Then GoTo NextRow

        Dim cmd As String
        cmd = CmdAt(R)
        If Len(cmd) = 0 Then GoTo NextRow
        If m_displayDepths(R) < labelDepth Then Exit Function

        If m_displayDepths(R) = labelDepth Then
            Select Case cmd
                Case "RETURN"
                    FindSubroutineReturnLineForFlow = R
                    Exit Function
                Case "LABEL"
                    Exit Function
            End Select
        End If

NextRow:
    Next R
End Function

Private Function HasFlowBreakBefore(ByVal lineNo As Long) As Boolean
    Dim R As Long
    For R = lineNo - 1 To 1 Step -1
        If IsCommentLine(R) Then GoTo NextRow

        Dim cmd As String
        cmd = CmdAt(R)
        If Len(cmd) > 0 Then
            HasFlowBreakBefore = (cmd = "EXIT" Or cmd = "GOTO" Or cmd = "RETURN")
            Exit Function
        End If
NextRow:
    Next R
End Function

Private Sub BuildVariableAnalysis()
    Dim defined As Object
    Dim used As Object
    Dim defRows As Object
    Set defined = CreateObject("Scripting.Dictionary")
    Set used = CreateObject("Scripting.Dictionary")
    Set defRows = CreateObject("Scripting.Dictionary")
    defined.CompareMode = vbTextCompare
    used.CompareMode = vbTextCompare
    defRows.CompareMode = vbTextCompare

    Dim R As Long
    For R = 1 To m_TotalLines
        If IsCommentLine(R) Then GoTo NextRow

        Dim cmd As String
        cmd = CmdAt(R)
        If Len(cmd) = 0 Then GoTo NextRow

        Dim def As CmdDef
        Set def = modCmdRegistry.GetCmdDef(cmd)
        If def Is Nothing Then GoTo NextRow

        Dim idx As Long
        For idx = 1 To 3
            Dim spec As ParamSpec
            Set spec = def.Param(idx)
            If Not spec Is Nothing Then
                If Not spec.IsRaw Then
                    Dim uses As Collection
                    Set uses = New Collection
                    CollectUserVarsFromText ValueAt(R, idx + 2), uses

                    Dim varName As Variant
                    For Each varName In uses
                        used(CStr(varName)) = True
                        If Not defined.exists(CStr(varName)) Then
                            AddRowMessage m_variableErrors, R, UndefinedVarMessage(CStr(varName), defined.Keys)
                        End If
                    Next varName
                End If
            End If
        Next idx

        If cmd = "SET" Or cmd = "INC" Or cmd = "DEC" Then
            Dim lvalueName As String
            lvalueName = MemoryLValueName(P1At(R))
            If Len(lvalueName) > 0 Then
                If (cmd = "INC" Or cmd = "DEC") And Not defined.exists(lvalueName) Then
                    AddRowMessage m_variableErrors, R, UndefinedVarMessage(lvalueName, defined.Keys) & ". Initialize it with SET first."
                End If

                If Not defined.exists(lvalueName) Then defined.Add lvalueName, True
                AddDefRow defRows, lvalueName, R
            End If
        End If

NextRow:
    Next R

    Dim nameKey As Variant
    For Each nameKey In defRows.Keys
        If Not used.exists(CStr(nameKey)) Then
            Dim rows As Collection
            Set rows = defRows(CStr(nameKey))

            Dim rowItem As Variant
            For Each rowItem In rows
                If Not m_variableErrors.exists(CStr(rowItem)) Then
                    AddRowMessage m_variableWarnings, CLng(rowItem), "Warning: defined variable is not used: {" & CStr(nameKey) & "}"
                End If
            Next rowItem
        End If
    Next nameKey
End Sub

Private Sub CollectUserVarsFromText(ByVal text As String, ByVal result As Collection)
    Dim i As Long
    Dim n As Long
    n = Len(text)
    i = 1

    Do While i <= n
        If Mid$(text, i, 1) = "{" Then
            Dim depth As Long
            Dim j As Long
            Dim found As Long
            depth = 1
            For j = i + 1 To n
                Select Case Mid$(text, j, 1)
                    Case "{"
                        depth = depth + 1
                    Case "}"
                        depth = depth - 1
                        If depth = 0 Then
                            found = j
                            Exit For
                        End If
                End Select
            Next j

            If found = 0 Then Exit Do

            Dim inner As String
            inner = Mid$(text, i + 1, found - i - 1)
            CollectUserVarsFromText inner, result

            If IsUserVariableToken(inner) Then result.Add Trim$(inner)
            i = found + 1
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Function IsUserVariableToken(ByVal inner As String) As Boolean
    Dim key As String
    key = Trim$(inner)
    If Len(key) = 0 Then Exit Function
    If Left$(key, 1) = "$" Then Exit Function
    If InStr(1, key, "{", vbTextCompare) > 0 Or InStr(1, key, "}", vbTextCompare) > 0 Then Exit Function

    IsUserVariableToken = True
End Function

Private Function MemoryLValueName(ByVal rawP1 As String) As String
    Dim s As String
    s = Trim$(rawP1)
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "{" Or Right$(s, 1) <> "}" Then Exit Function

    Dim inner As String
    inner = Trim$(Mid$(s, 2, Len(s) - 2))
    If Len(inner) = 0 Then Exit Function
    If UCase$(inner) = "$CLIP" Then Exit Function
    If Left$(inner, 1) = "$" Then Exit Function

    MemoryLValueName = inner
End Function

Private Sub AddDefRow(ByVal defRows As Object, ByVal varName As String, ByVal lineNo As Long)
    Dim rows As Collection
    If defRows.exists(varName) Then
        Set rows = defRows(varName)
    Else
        Set rows = New Collection
        defRows.Add varName, rows
    End If
    rows.Add lineNo
End Sub

Private Sub AddRowMessage(ByVal rows As Object, ByVal lineNo As Long, ByVal message As String)
    Dim key As String
    key = CStr(lineNo)

    If rows.exists(key) Then
        If InStr(1, CStr(rows(key)), message, vbTextCompare) = 0 Then
            rows(key) = CStr(rows(key)) & vbCrLf & message
        End If
    Else
        rows.Add key, message
    End If
End Sub

Private Function UndefinedVarMessage(ByVal varName As String, ByVal candidates As Variant) As String
    Dim msg As String
    msg = "Undefined variable: {" & varName & "}"

    Dim suggestion As String
    suggestion = ClosestName(varName, candidates)
    If Len(suggestion) > 0 Then msg = msg & ". Did you mean {" & suggestion & "}?"

    UndefinedVarMessage = msg
End Function

Private Function ClosestName(ByVal target As String, ByVal candidates As Variant) As String
    On Error GoTo SafeExit

    Dim threshold As Long
    If Len(Trim$(target)) <= 4 Then
        threshold = 1
    Else
        threshold = 2
    End If

    Dim bestDistance As Long
    bestDistance = 32767

    Dim bestName As String
    Dim bestCount As Long
    Dim item As Variant
    For Each item In candidates
        Dim candidate As String
        candidate = CStr(item)

        Dim distance As Long
        distance = LevenshteinDistance(target, candidate)
        If distance <= threshold Then
            If distance < bestDistance Then
                bestDistance = distance
                bestName = candidate
                bestCount = 1
            ElseIf distance = bestDistance Then
                bestCount = bestCount + 1
            End If
        End If
    Next item

    If bestCount = 1 Then ClosestName = bestName
SafeExit:
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
