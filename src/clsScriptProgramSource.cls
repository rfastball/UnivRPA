VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScriptProgramSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @internal ScriptProgram implementation detail.
' @access Only clsScriptProgram should create or call this class.
' @desc ScriptProgram source, row, command access, and instruction cache.
' @depends modCmdRegistry, CmdDef, clsInstruction
Option Explicit

Private Const COL_CMD As Long = 2
Private Const COL_P1 As Long = 3
Private Const COL_P2 As Long = 4
Private Const COL_P3 As Long = 5

Private m_lines As Variant
Private m_TotalLines As Long
Private m_firstSheetRow As Long
Private m_instructions As Collection

Private Sub Class_Initialize()
    m_firstSheetRow = 1
    Set m_instructions = New Collection
End Sub

Public Sub LoadFromArray(ByVal sourceArray As Variant, Optional ByVal firstSheetRow As Long = 1)
    If Not IsArray(sourceArray) Then Err.Raise 5, "clsScriptProgramSource.LoadFromArray", "Program source must be a 2D array."
    m_lines = sourceArray
    m_TotalLines = UBound(m_lines, 1)
    m_firstSheetRow = firstSheetRow
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

Public Function ValidateLineRange(ByVal lineNo As Long) As Boolean
    If lineNo < 1 Or lineNo > m_TotalLines Then Err.Raise 9, "clsScriptProgramSource.ValidateLineRange", "Runtime line out of range: " & lineNo
    ValidateLineRange = True
End Function

Public Function ValueAt(ByVal lineNo As Long, ByVal columnNo As Long) As String
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
