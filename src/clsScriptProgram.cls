VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScriptProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc Static ScriptProgram facade. External modules should query only this class.
' @depends clsScriptProgramSource, clsScriptProgramIndex, clsScriptProgramAnalysis
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

Private m_source As clsScriptProgramSource
Private m_index As clsScriptProgramIndex
Private m_analysis As clsScriptProgramAnalysis

Private Sub Class_Initialize()
    Set m_source = New clsScriptProgramSource
    Set m_index = New clsScriptProgramIndex
    Set m_analysis = New clsScriptProgramAnalysis
End Sub

Private Sub WireInternals()
    m_index.Init m_source
    m_analysis.Init m_source, m_index
End Sub

Public Sub LoadFromRange(ByVal sourceRange As Range)
    If sourceRange Is Nothing Then Err.Raise 5, "clsScriptProgram.LoadFromRange", "Source range is required."
    LoadFromArray sourceRange.value, sourceRange.Row
End Sub

Public Sub LoadFromArray(ByVal sourceArray As Variant, Optional ByVal firstSheetRow As Long = 1)
    If Not IsArray(sourceArray) Then Err.Raise 5, "clsScriptProgram.LoadFromArray", "Program source must be a 2D array."
    m_source.LoadFromArray sourceArray, firstSheetRow
    WireInternals
    m_index.BuildStructureIndex
    m_index.BuildLabelIndex
    m_index.BuildStateCoordinates
    m_analysis.BuildStaticAnalysis
End Sub

Public Property Get totalLines() As Long
    totalLines = m_source.totalLines
End Property

Public Function RawCmdAt(ByVal lineNo As Long) As String
    RawCmdAt = m_source.RawCmdAt(lineNo)
End Function

Public Function CmdAt(ByVal lineNo As Long) As String
    CmdAt = m_source.CmdAt(lineNo)
End Function

Public Function P1At(ByVal lineNo As Long) As String
    P1At = m_source.P1At(lineNo)
End Function

Public Function P2At(ByVal lineNo As Long) As String
    P2At = m_source.P2At(lineNo)
End Function

Public Function P3At(ByVal lineNo As Long) As String
    P3At = m_source.P3At(lineNo)
End Function

Public Function IsBlankLine(ByVal lineNo As Long) As Boolean
    IsBlankLine = m_source.IsBlankLine(lineNo)
End Function

Public Function IsCommentLine(ByVal lineNo As Long) As Boolean
    IsCommentLine = m_source.IsCommentLine(lineNo)
End Function

Public Function IsSkippable(ByVal lineNo As Long) As Boolean
    IsSkippable = m_source.IsSkippable(lineNo)
End Function

Public Function SheetRowFromRuntimeLine(ByVal lineNo As Long) As Long
    SheetRowFromRuntimeLine = m_source.SheetRowFromRuntimeLine(lineNo)
End Function

Public Function RuntimeLineFromSheetRow(ByVal sheetRow As Long) As Long
    RuntimeLineFromSheetRow = m_source.RuntimeLineFromSheetRow(sheetRow)
End Function

Public Function InstructionAt(ByVal lineNo As Long) As clsInstruction
    Set InstructionAt = m_source.InstructionAt(lineNo)
End Function

Public Function FindBlockPair(ByVal startLine As Long, ByVal group As String, Optional ByVal Direction As Long = 1) As Long
    FindBlockPair = m_index.FindBlockPair(startLine, group, Direction)
End Function

Public Function SelectBlock(ByVal SelectLine As Long) As clsSelectBlock
    Set SelectBlock = m_index.SelectBlock(SelectLine)
End Function

Public Function EndSelectOfCase(ByVal caseLine As Long) As Long
    EndSelectOfCase = m_index.EndSelectOfCase(caseLine)
End Function

Public Function EndIfOf(ByVal lineNo As Long) As Long
    EndIfOf = m_index.EndIfOf(lineNo)
End Function

Public Function LabelRoleAt(ByVal lineNo As Long) As String
    LabelRoleAt = m_index.LabelRoleAt(lineNo)
End Function

Public Function LabelRoleAtSheetRow(ByVal sheetRow As Long) As String
    LabelRoleAtSheetRow = LabelRoleAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function LabelRoleWarningAt(ByVal lineNo As Long) As String
    LabelRoleWarningAt = m_index.LabelRoleWarningAt(lineNo)
End Function

Public Function LabelRoleWarningAtSheetRow(ByVal sheetRow As Long) As String
    LabelRoleWarningAtSheetRow = LabelRoleWarningAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsCallTargetLabelAt(ByVal lineNo As Long) As Boolean
    IsCallTargetLabelAt = m_index.IsCallTargetLabelAt(lineNo)
End Function

Public Function IsJumpTargetMissingAt(ByVal lineNo As Long) As Boolean
    IsJumpTargetMissingAt = m_index.IsJumpTargetMissingAt(lineNo)
End Function

Public Function ResolveJumpTargetSheetRow(ByVal Target As String) As Long
    Dim lineNo As Long
    lineNo = ResolveJumpTarget(Target)
    If lineNo > 0 Then ResolveJumpTargetSheetRow = SheetRowFromRuntimeLine(lineNo)
End Function

Public Function FirstCallSourceLineForLabel(ByVal labelName As String) As Long
    FirstCallSourceLineForLabel = m_index.FirstCallSourceLineForLabel(labelName)
End Function

Public Function FirstCallSourceSheetRowForLabel(ByVal labelName As String) As Long
    Dim lineNo As Long
    lineNo = FirstCallSourceLineForLabel(labelName)
    If lineNo > 0 Then FirstCallSourceSheetRowForLabel = SheetRowFromRuntimeLine(lineNo)
End Function

Public Function DisplayDepthAt(ByVal lineNo As Long) As Long
    DisplayDepthAt = m_analysis.DisplayDepthAt(lineNo)
End Function

Public Function DisplayDepthAtSheetRow(ByVal sheetRow As Long) As Long
    DisplayDepthAtSheetRow = DisplayDepthAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function StyleDepthAt(ByVal lineNo As Long) As Long
    StyleDepthAt = m_analysis.StyleDepthAt(lineNo)
End Function

Public Function StyleDepthAtSheetRow(ByVal sheetRow As Long) As Long
    StyleDepthAtSheetRow = StyleDepthAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsSubroutineStartAt(ByVal lineNo As Long) As Boolean
    IsSubroutineStartAt = m_analysis.IsSubroutineStartAt(lineNo)
End Function

Public Function IsSubroutineStartAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsSubroutineStartAtSheetRow = IsSubroutineStartAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsSubroutineEndAt(ByVal lineNo As Long) As Boolean
    IsSubroutineEndAt = m_analysis.IsSubroutineEndAt(lineNo)
End Function

Public Function IsSubroutineEndAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsSubroutineEndAtSheetRow = IsSubroutineEndAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function IsFlowInvalidAt(ByVal lineNo As Long) As Boolean
    IsFlowInvalidAt = m_analysis.IsFlowInvalidAt(lineNo)
End Function

Public Function IsFlowInvalidAtSheetRow(ByVal sheetRow As Long) As Boolean
    IsFlowInvalidAtSheetRow = IsFlowInvalidAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function VariableErrorAt(ByVal lineNo As Long) As String
    VariableErrorAt = m_analysis.VariableErrorAt(lineNo)
End Function

Public Function VariableErrorAtSheetRow(ByVal sheetRow As Long) As String
    VariableErrorAtSheetRow = VariableErrorAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function VariableWarningAt(ByVal lineNo As Long) As String
    VariableWarningAt = m_analysis.VariableWarningAt(lineNo)
End Function

Public Function VariableWarningAtSheetRow(ByVal sheetRow As Long) As String
    VariableWarningAtSheetRow = VariableWarningAt(RuntimeLineFromSheetRow(sheetRow))
End Function

Public Function HasLabel(ByVal name As String) As Boolean
    HasLabel = m_index.HasLabel(name)
End Function

Public Function LabelLine(ByVal name As String) As Long
    LabelLine = m_index.LabelLine(name)
End Function

Public Function ResolveJumpTarget(ByVal Target As String) As Long
    ResolveJumpTarget = m_index.ResolveJumpTarget(Target)
End Function

Public Function LoopDepthAt(ByVal lineNo As Long) As Long
    LoopDepthAt = m_index.LoopDepthAt(lineNo)
End Function

Public Function SubAreaIdAt(ByVal lineNo As Long) As Long
    SubAreaIdAt = m_index.SubAreaIdAt(lineNo)
End Function

Public Function IsGotoAllowed(ByVal sourceLine As Long, ByVal targetLine As Long) As Boolean
    IsGotoAllowed = m_index.IsGotoAllowed(sourceLine, targetLine)
End Function

Public Function IsCallAllowed(ByVal sourceLine As Long, ByVal targetLine As Long) As Boolean
    IsCallAllowed = m_index.IsCallAllowed(sourceLine, targetLine)
End Function

Public Function HasDuplicateLabels() As Boolean
    HasDuplicateLabels = m_index.HasDuplicateLabels()
End Function

Public Function DuplicateLabels() As String
    DuplicateLabels = m_index.DuplicateLabels()
End Function

Public Function ValidateStructure() As Boolean
    ValidateStructure = m_analysis.ValidateStructure()
End Function

Public Function ValidateLabels() As Boolean
    ValidateLabels = m_analysis.ValidateLabels()
End Function

Public Function ValidationErrors() As String
    ValidationErrors = m_analysis.ValidationErrors()
End Function

Public Function ValidateLineRange(ByVal lineNo As Long) As Boolean
    ValidateLineRange = m_source.ValidateLineRange(lineNo)
End Function
