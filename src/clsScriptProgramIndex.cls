VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScriptProgramIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @internal ScriptProgram implementation detail.
' @access Only clsScriptProgram should create or call this class.
' @desc ScriptProgram structure, label, and coordinate indexes.
' @depends modCmdRegistry, CmdDef, clsSelectBlock, clsScriptProgramSource
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit


Private m_source As clsScriptProgramSource
Private m_selectBlocks As Object
Private m_caseEndSelectLines As Object
Private m_ifEndLines As Object
Private m_labels As Object
Private m_labelRows As Object
Private m_duplicateLabels As Object
Private m_callLabelTargets As Object
Private m_gotoLabelTargets As Object
Private m_callSourceRows As Object
Private m_labelRoles As Object
Private m_labelRoleWarnings As Object
Private m_loopDepths() As Long
Private m_subAreaIds() As Long
Private m_loopRanges As Collection
' [C3] 루프 짝/고아 CASE 구조 오류 (라인 -> 메시지). BuildStructureIndex의
'   통합 스택 스캔이 채우고 Analysis.ValidateStructure가 차단 오류로 승격한다.
Private m_blockPairErrors As Object
' [D3] IF 구조 오류 (라인 -> 메시지). 같은 통합 스캔이 채우고 Analysis가
'   서식(m_flowInvalidRows)과 차단 오류 양쪽으로 소비한다 (스캔 단일 소유).
Private m_ifStructureErrors As Object

' [D2] 빌드 완료 플래그: 소스는 로드 후 불변이라 재빌드는 항상 동일 결과.
'   로드가 만든 인덱스를 검증(ValidateStructure/ValidateLabels)이 재구축하지 않게 한다.
Private m_structureBuilt As Boolean
Private m_labelIndexBuilt As Boolean
Private m_coordinatesBuilt As Boolean

Private Sub Class_Initialize()
    ResetIndexes
End Sub

Public Sub Init(ByVal Source As clsScriptProgramSource)
    Set m_source = Source
    ResetIndexes
End Sub

Private Sub ResetIndexes()
    Set m_selectBlocks = CreateObject("Scripting.Dictionary")
    Set m_caseEndSelectLines = CreateObject("Scripting.Dictionary")
    Set m_ifEndLines = CreateObject("Scripting.Dictionary")
    Set m_labels = CreateObject("Scripting.Dictionary")
    Set m_labelRows = CreateObject("Scripting.Dictionary")
    Set m_duplicateLabels = CreateObject("Scripting.Dictionary")
    Set m_callLabelTargets = CreateObject("Scripting.Dictionary")
    Set m_gotoLabelTargets = CreateObject("Scripting.Dictionary")
    Set m_callSourceRows = CreateObject("Scripting.Dictionary")
    Set m_labelRoles = CreateObject("Scripting.Dictionary")
    Set m_labelRoleWarnings = CreateObject("Scripting.Dictionary")
    Set m_loopRanges = New Collection
    Set m_blockPairErrors = CreateObject("Scripting.Dictionary")
    Set m_ifStructureErrors = CreateObject("Scripting.Dictionary")
    m_labels.CompareMode = vbTextCompare
    m_labelRows.CompareMode = vbTextCompare
    m_duplicateLabels.CompareMode = vbTextCompare
    m_callLabelTargets.CompareMode = vbTextCompare
    m_gotoLabelTargets.CompareMode = vbTextCompare
    m_callSourceRows.CompareMode = vbTextCompare
    m_labelRoles.CompareMode = vbTextCompare
    m_labelRoleWarnings.CompareMode = vbTextCompare
    m_structureBuilt = False
    m_labelIndexBuilt = False
    m_coordinatesBuilt = False
End Sub

Public Property Get SelectBlocks() As Object
    Set SelectBlocks = m_selectBlocks
End Property

Public Function FindBlockPair(ByVal startLine As Long, ByVal group As String, Optional ByVal Direction As Long = 1) As Long
    On Error GoTo ErrHandler
    m_source.ValidateLineRange startLine

    group = UCase$(Trim$(group))
    If group <> "IF" And group <> "WHILE" And group <> "REPEAT" And group <> "FOREACH" Then
        FindBlockPair = CMD_ERROR
        Exit Function
    End If
    If Direction <> 1 And Direction <> -1 Then
        FindBlockPair = CMD_ERROR
        Exit Function
    End If

    Dim R As Long
    Dim endLine As Long
    Dim depth As Long
    If Direction = 1 Then endLine = m_source.totalLines Else endLine = 1

    For R = startLine + Direction To endLine Step Direction
        If m_source.IsSkippable(R) Then GoTo NextRow

        Dim def As CmdDef
        Set def = modCmdRegistry.GetCmdDef(m_source.CmdAt(R))
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

    FindBlockPair = CMD_ERROR
    Exit Function
ErrHandler:
    FindBlockPair = CMD_ERROR
End Function

Public Sub BuildStructureIndex()
    If m_structureBuilt Then Exit Sub

    Set m_selectBlocks = CreateObject("Scripting.Dictionary")
    Set m_caseEndSelectLines = CreateObject("Scripting.Dictionary")
    Set m_ifEndLines = CreateObject("Scripting.Dictionary")

    Dim R As Long
    For R = 1 To m_source.totalLines
        If m_source.CmdAt(R) = "SELECT" Then BuildSelectBlock R
    Next R
    BuildBlockPairIndex
    m_structureBuilt = True
End Sub

Public Property Get BlockPairErrors() As Object
    Set BlockPairErrors = m_blockPairErrors
End Property

Public Property Get IfStructureErrors() As Object
    Set IfStructureErrors = m_ifStructureErrors
End Property

Public Function SelectBlock(ByVal SelectLine As Long) As clsSelectBlock
    If m_selectBlocks.exists(CStr(SelectLine)) Then Set SelectBlock = m_selectBlocks(CStr(SelectLine))
End Function

Public Function EndSelectOfCase(ByVal caseLine As Long) As Long
    If m_caseEndSelectLines.exists(CStr(caseLine)) Then
        EndSelectOfCase = CLng(m_caseEndSelectLines(CStr(caseLine)))
    Else
        EndSelectOfCase = CMD_ERROR
    End If
End Function

Public Function EndIfOf(ByVal lineNo As Long) As Long
    m_source.ValidateLineRange lineNo
    If m_ifEndLines.exists(CStr(lineNo)) Then
        EndIfOf = CLng(m_ifEndLines(CStr(lineNo)))
    Else
        EndIfOf = CMD_ERROR
    End If
End Function

Public Sub BuildLabelIndex()
    If m_labelIndexBuilt Then Exit Sub
    BuildLabelReferenceIndex
    m_labelIndexBuilt = True
End Sub

Public Sub BuildStateCoordinates()
    If m_coordinatesBuilt Then Exit Sub

    ReDim m_loopDepths(1 To m_source.totalLines)
    ReDim m_subAreaIds(1 To m_source.totalLines)
    Set m_loopRanges = New Collection

    BuildLoopCoordinates
    BuildSubAreaCoordinates
    m_coordinatesBuilt = True
End Sub

Public Function LabelRoleAt(ByVal lineNo As Long) As String
    m_source.ValidateLineRange lineNo
    If m_labelRoles.exists(CStr(lineNo)) Then LabelRoleAt = CStr(m_labelRoles(CStr(lineNo)))
End Function

Public Function LabelRoleWarningAt(ByVal lineNo As Long) As String
    m_source.ValidateLineRange lineNo
    If m_labelRoleWarnings.exists(CStr(lineNo)) Then LabelRoleWarningAt = CStr(m_labelRoleWarnings(CStr(lineNo)))
End Function

Public Function IsCallTargetLabelAt(ByVal lineNo As Long) As Boolean
    Dim roleName As String
    roleName = LabelRoleAt(lineNo)
    IsCallTargetLabelAt = (roleName = "SubroutineLabel" Or roleName = "MixedLabel")
End Function

Public Function IsJumpTargetMissingAt(ByVal lineNo As Long) As Boolean
    m_source.ValidateLineRange lineNo

    Dim cmd As String
    cmd = m_source.CmdAt(lineNo)
    If cmd <> "GOTO" And cmd <> "CALL" Then Exit Function

    Dim targetName As String
    targetName = NormalizeLabelName(m_source.P3At(lineNo))
    If IsStaticLabelSymbol(targetName) Then IsJumpTargetMissingAt = Not m_labels.exists(targetName)
End Function

Public Function FirstCallSourceLineForLabel(ByVal labelName As String) As Long
    Dim key As String
    key = NormalizeLabelName(labelName)
    If Not m_callSourceRows.exists(key) Then Exit Function

    Dim rows As Collection
    Set rows = m_callSourceRows(key)
    If rows.count > 0 Then FirstCallSourceLineForLabel = CLng(rows.item(1))
End Function

Public Function HasLabel(ByVal name As String) As Boolean
    HasLabel = m_labels.exists(NormalizeLabelName(name))
End Function

Public Function LabelLine(ByVal name As String) As Long
    Dim key As String
    key = NormalizeLabelName(name)
    If m_labels.exists(key) Then LabelLine = CLng(m_labels(key)) Else LabelLine = CMD_ERROR
End Function

Public Function ResolveJumpTarget(ByVal Target As String) As Long
    Dim trimmedTarget As String
    trimmedTarget = Trim$(Target)

    If Len(trimmedTarget) = 0 Or IsNumeric(trimmedTarget) Or ContainsTokenSyntax(trimmedTarget) Then
        ResolveJumpTarget = CMD_ERROR
    Else
        ResolveJumpTarget = LabelLine(trimmedTarget)
    End If
End Function

Public Function LoopDepthAt(ByVal lineNo As Long) As Long
    m_source.ValidateLineRange lineNo
    LoopDepthAt = m_loopDepths(lineNo)
End Function

Public Function SubAreaIdAt(ByVal lineNo As Long) As Long
    m_source.ValidateLineRange lineNo
    SubAreaIdAt = m_subAreaIds(lineNo)
End Function

Public Function IsGotoAllowed(ByVal sourceLine As Long, ByVal targetLine As Long) As Boolean
    m_source.ValidateLineRange sourceLine
    m_source.ValidateLineRange targetLine

    If SubAreaIdAt(sourceLine) <> SubAreaIdAt(targetLine) Then Exit Function
    IsGotoAllowed = IsLoopTransferAllowed(sourceLine, targetLine)
End Function

' [CALL 경계] CALL의 존재 이유가 다른 sub-area의 서브루틴 호출이므로 IsGotoAllowed의
'   동일 영역 강제는 적용하지 않는다. 루프 본문 중간 진입만 금지 - 헤더를 거치지 않은
'   진입은 LoopStack 프레임 부재로 END_* 백점프/RETURN 깊이 복원이 깨진다.
Public Function IsCallAllowed(ByVal sourceLine As Long, ByVal targetLine As Long) As Boolean
    m_source.ValidateLineRange sourceLine
    m_source.ValidateLineRange targetLine

    IsCallAllowed = IsLoopTransferAllowed(sourceLine, targetLine)
End Function

Public Function HasDuplicateLabels() As Boolean
    HasDuplicateLabels = (m_duplicateLabels.count > 0)
End Function

Public Function DuplicateLabels() As String
    DuplicateLabels = JoinDictionaryKeys(m_duplicateLabels)
End Function

Public Function NormalizeLabelName(ByVal name As String) As String
    NormalizeLabelName = UCase$(Trim$(name))
End Function

Public Function ContainsTokenSyntax(ByVal value As String) As Boolean
    ContainsTokenSyntax = (InStr(1, value, "{", vbTextCompare) > 0 _
                        Or InStr(1, value, "}", vbTextCompare) > 0 _
                        Or InStr(1, value, "$", vbTextCompare) > 0)
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
    For R = 1 To m_source.totalLines
        If m_source.IsSkippable(R) Then GoTo NextRow

        Dim cmd As String
        cmd = m_source.CmdAt(R)

        If cmd = "LABEL" Then
            AddLabelReferenceRow R
        ElseIf cmd = "CALL" Then
            AddJumpReferenceTarget m_callLabelTargets, m_source.P3At(R)
            AddCallSourceRow m_source.P3At(R), R
        ElseIf cmd = "GOTO" Then
            AddJumpReferenceTarget m_gotoLabelTargets, m_source.P3At(R)
        End If
NextRow:
    Next R

    BuildLabelRoles
End Sub

Private Sub AddLabelReferenceRow(ByVal lineNo As Long)
    Dim labelName As String
    labelName = NormalizeLabelName(m_source.P3At(lineNo))
    If Len(labelName) = 0 Or ContainsTokenSyntax(m_source.P3At(lineNo)) Then Exit Sub

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
    For Each labelName In m_labelRows.keys
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

' [D2] 단일 패스 스택 스캔: 라인마다 FindBlockPair 재스캔(O(N^2) 잠재)을 제거.
'   교차 중첩 등 비정합 짝은 BuildBlockPairIndex가 실행 전 차단하므로 여기서는
'   가장 가까운 STARTER와 짝짓는 것으로 충분하다.
Private Sub BuildLoopCoordinates()
    Dim depth As Long
    Dim R As Long
    Dim starterLines() As Long
    Dim Top As Long

    For R = 1 To m_source.totalLines
        Dim cmd As String
        cmd = m_source.CmdAt(R)

        If IsLoopEnder(cmd) Then
            depth = IIf(depth > 0, depth - 1, 0)
            If Top > 0 Then
                m_loopRanges.Add Array(starterLines(Top), R)
                Top = Top - 1
            End If
        End If
        m_loopDepths(R) = depth

        If IsLoopStarter(cmd) Then
            Top = Top + 1
            ReDim Preserve starterLines(1 To Top)
            starterLines(Top) = R
            depth = depth + 1
        End If
    Next R
End Sub

Private Sub BuildSubAreaCoordinates()
    Dim areaId As Long
    Dim key As Variant
    For Each key In m_callLabelTargets.keys
        If m_labels.exists(CStr(key)) Then
            areaId = areaId + 1
            Dim startLine As Long
            Dim endLine As Long
            startLine = CLng(m_labels(CStr(key)))
            endLine = FindSubroutineEndAtDepth(startLine, m_loopDepths)
            If endLine <= 0 Then endLine = startLine
            MarkSubArea startLine, endLine, areaId
        End If
    Next key
End Sub

' [D3] 서브루틴 종료행 스캔의 단일 소유: 라벨과 같은 깊이의 첫 RETURN을 찾고
'   같은 깊이의 다른 LABEL 또는 더 얕은 깊이를 만나면 중단(0 반환).
'   깊이 모델은 호출자가 선택 - Index(sub-area)는 루프 깊이, Analysis(서식/검증)는
'   표시 깊이를 넘긴다 (두 모델은 의미가 달라 의도적으로 분리 유지).
Public Function FindSubroutineEndAtDepth(ByVal startLine As Long, ByRef depths() As Long) As Long
    Dim labelDepth As Long
    labelDepth = depths(startLine)

    Dim R As Long
    For R = startLine + 1 To m_source.totalLines
        If m_source.IsSkippable(R) Then GoTo NextRow

        Dim cmd As String
        cmd = m_source.CmdAt(R)
        If Len(cmd) = 0 Then GoTo NextRow
        If depths(R) < labelDepth Then Exit Function

        If depths(R) = labelDepth Then
            Select Case cmd
                Case "RETURN"
                    FindSubroutineEndAtDepth = R
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
    For R = SelectLine + 1 To m_source.totalLines
        If m_source.IsSkippable(R) Then GoTo NextRow

        Dim def As CmdDef
        Set def = modCmdRegistry.GetCmdDef(m_source.CmdAt(R))
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

' [C3+D3] 블록 구조 통합 스택 스캔 (단일 패스, Index가 스캔의 단일 소유자):
'   - 루프 3그룹(WHILE/REPEAT/FOREACH): 고아 ENDER / 미폐쇄 STARTER / 그룹 교차
'   - SELECT 깊이: 고아 CASE / END_SELECT
'   - IF 그룹: 멤버->END_IF 짝 인덱스(m_ifEndLines) + 구조 오류(고아 분기,
'     ELSE 순서/중복, 미폐쇄 IF)를 m_ifStructureErrors에 기록
'   SELECT 미폐쇄는 SelectBlocks 검증, ELSE_IF 파라미터는 Analysis가 담당.
Private Sub BuildBlockPairIndex()
    Set m_blockPairErrors = CreateObject("Scripting.Dictionary")
    Set m_ifStructureErrors = CreateObject("Scripting.Dictionary")

    Dim stackGroups() As String
    Dim stackLines() As Long
    Dim Top As Long
    Dim selectDepth As Long

    Dim ifMembers() As Collection
    Dim ifSawElse() As Boolean
    Dim ifLines() As Long
    Dim ifTop As Long

    Dim R As Long
    For R = 1 To m_source.totalLines
        If m_source.IsSkippable(R) Then GoTo NextRow

        Dim cmd As String
        cmd = m_source.CmdAt(R)

        Select Case cmd
            Case "SELECT"
                selectDepth = selectDepth + 1
            Case "CASE"
                If selectDepth = 0 Then AddBlockPairError R, "CASE line " & R & " has no enclosing SELECT."
            Case "END_SELECT"
                If selectDepth = 0 Then
                    AddBlockPairError R, "END_SELECT line " & R & " has no matching SELECT."
                Else
                    selectDepth = selectDepth - 1
                End If

            Case "IF"
                ifTop = ifTop + 1
                ReDim Preserve ifMembers(1 To ifTop)
                ReDim Preserve ifSawElse(1 To ifTop)
                ReDim Preserve ifLines(1 To ifTop)
                Set ifMembers(ifTop) = New Collection
                ifMembers(ifTop).Add R
                ifSawElse(ifTop) = False
                ifLines(ifTop) = R

            Case "ELSE_IF"
                If ifTop = 0 Then
                    AddIfStructureError R, "ELSE_IF line " & R & " has no matching IF."
                ElseIf ifSawElse(ifTop) Then
                    AddIfStructureError R, "ELSE_IF line " & R & " cannot appear after ELSE."
                    ifMembers(ifTop).Add R
                Else
                    ifMembers(ifTop).Add R
                End If

            Case "ELSE"
                If ifTop = 0 Then
                    AddIfStructureError R, "ELSE line " & R & " has no matching IF."
                ElseIf ifSawElse(ifTop) Then
                    AddIfStructureError R, "ELSE line " & R & " duplicates an ELSE in the same IF block."
                    ifMembers(ifTop).Add R
                Else
                    ifSawElse(ifTop) = True
                    ifMembers(ifTop).Add R
                End If

            Case "END_IF"
                If ifTop = 0 Then
                    AddIfStructureError R, "END_IF line " & R & " has no matching IF."
                Else
                    ifMembers(ifTop).Add R
                    Dim memberLine As Variant
                    For Each memberLine In ifMembers(ifTop)
                        m_ifEndLines(CStr(memberLine)) = R
                    Next memberLine
                    ifTop = ifTop - 1
                End If

            Case Else
                If IsLoopStarter(cmd) Then
                    Top = Top + 1
                    ReDim Preserve stackGroups(1 To Top)
                    ReDim Preserve stackLines(1 To Top)
                    stackGroups(Top) = LoopGroupOf(cmd)
                    stackLines(Top) = R
                ElseIf IsLoopEnder(cmd) Then
                    If Top = 0 Then
                        AddBlockPairError R, cmd & " line " & R & " has no matching loop starter."
                    Else
                        If stackGroups(Top) <> LoopEnderGroupOf(cmd) Then
                            AddBlockPairError R, cmd & " line " & R & " crosses loop nesting: innermost open loop is " _
                                & stackGroups(Top) & " (line " & stackLines(Top) & ")."
                        End If
                        Top = Top - 1
                    End If
                End If
        End Select

NextRow:
    Next R

    Do While Top > 0
        AddBlockPairError stackLines(Top), stackGroups(Top) & " line " & stackLines(Top) & " has no matching loop ender."
        Top = Top - 1
    Loop

    Do While ifTop > 0
        AddIfStructureError ifLines(ifTop), "IF line " & ifLines(ifTop) & " has no END_IF."
        ifTop = ifTop - 1
    Loop
End Sub

Private Sub AddIfStructureError(ByVal lineNo As Long, ByVal message As String)
    Dim key As String
    key = CStr(lineNo)
    If m_ifStructureErrors.exists(key) Then
        m_ifStructureErrors(key) = m_ifStructureErrors(key) & " / " & message
    Else
        m_ifStructureErrors(key) = message
    End If
End Sub

Private Sub AddBlockPairError(ByVal lineNo As Long, ByVal message As String)
    Dim key As String
    key = CStr(lineNo)
    If m_blockPairErrors.exists(key) Then
        m_blockPairErrors(key) = m_blockPairErrors(key) & " / " & message
    Else
        m_blockPairErrors(key) = message
    End If
End Sub

Private Function LoopEnderGroupOf(ByVal cmd As String) As String
    Select Case cmd
        Case "END_WHILE": LoopEnderGroupOf = "WHILE"
        Case "END_REPEAT": LoopEnderGroupOf = "REPEAT"
        Case "END_FOREACH": LoopEnderGroupOf = "FOREACH"
    End Select
End Function

Private Function IsStaticLabelSymbol(ByVal value As String) As Boolean
    value = Trim$(value)
    IsStaticLabelSymbol = (Len(value) > 0 And Not IsNumeric(value) And Not ContainsTokenSyntax(value))
End Function

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

Private Function JoinDictionaryKeys(ByVal dict As Object) As String
    Dim key As Variant
    For Each key In dict.keys
        If Len(JoinDictionaryKeys) > 0 Then JoinDictionaryKeys = JoinDictionaryKeys & ", "
        JoinDictionaryKeys = JoinDictionaryKeys & CStr(key)
    Next key
End Function
