VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScriptProgramAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @internal ScriptProgram implementation detail.
' @access Only clsScriptProgram should create or call this class.
' @desc ScriptProgram flow, variable, and validation analysis.
' @depends modCmdRegistry, CmdDef, ParamSpec, clsScriptProgramSource, clsScriptProgramIndex, modStringDistance
Option Explicit

Private Type IfValidationFrame
    headerLine As Long
    sawElse As Boolean
End Type

Private m_source As clsScriptProgramSource
Private m_index As clsScriptProgramIndex
Private m_subroutineStartRows As Object
Private m_subroutineEndRows As Object
Private m_flowInvalidRows As Object
' 서브루틴 흐름 갭(RETURN 누락/흐름차단 부재/고아 RETURN)만 담는 라인->메시지 사전.
'   m_flowInvalidRows는 서식(빨간 표시)용 전수 집합, 이쪽은 실행 전 차단으로 승격할 부분집합.
'   점프대상/IF 오류는 각각 ValidateJumpTargets/BuildIfStructureAnalysis가 이미 차단하므로 제외.
Private m_flowBlockingErrors As Object
Private m_variableErrors As Object
Private m_variableWarnings As Object
Private m_validationErrors As Collection
Private m_displayDepths() As Long
Private m_styleDepths() As Long
Private m_subDepths() As Long
' [D2] 정적 분석 빌드 완료 플래그 (소스 불변 - 재빌드 스킵)
Private m_staticBuilt As Boolean

Private Sub Class_Initialize()
    ResetAnalysis
End Sub

Public Sub Init(ByVal Source As clsScriptProgramSource, ByVal index As clsScriptProgramIndex)
    Set m_source = Source
    Set m_index = index
    ResetAnalysis
End Sub

Private Sub ResetAnalysis()
    m_staticBuilt = False
    Set m_subroutineStartRows = CreateObject("Scripting.Dictionary")
    Set m_subroutineEndRows = CreateObject("Scripting.Dictionary")
    Set m_flowInvalidRows = CreateObject("Scripting.Dictionary")
    Set m_flowBlockingErrors = CreateObject("Scripting.Dictionary")
    Set m_variableErrors = CreateObject("Scripting.Dictionary")
    Set m_variableWarnings = CreateObject("Scripting.Dictionary")
    Set m_validationErrors = New Collection
    m_subroutineStartRows.CompareMode = vbTextCompare
    m_subroutineEndRows.CompareMode = vbTextCompare
    m_flowInvalidRows.CompareMode = vbTextCompare
    m_flowBlockingErrors.CompareMode = vbTextCompare
    m_variableErrors.CompareMode = vbTextCompare
    m_variableWarnings.CompareMode = vbTextCompare
End Sub

Public Sub BuildStaticAnalysis()
    ' [D2] 소스는 로드 후 불변 - 로드가 만든 분석을 ValidateLabels가 재구축하지 않는다
    If m_staticBuilt Then Exit Sub

    ReDim m_displayDepths(1 To m_source.totalLines)
    ReDim m_styleDepths(1 To m_source.totalLines)
    ReDim m_subDepths(1 To m_source.totalLines)

    Set m_subroutineStartRows = CreateObject("Scripting.Dictionary")
    Set m_subroutineEndRows = CreateObject("Scripting.Dictionary")
    Set m_flowInvalidRows = CreateObject("Scripting.Dictionary")
    Set m_flowBlockingErrors = CreateObject("Scripting.Dictionary")
    Set m_variableErrors = CreateObject("Scripting.Dictionary")
    Set m_variableWarnings = CreateObject("Scripting.Dictionary")
    m_subroutineStartRows.CompareMode = vbTextCompare
    m_subroutineEndRows.CompareMode = vbTextCompare
    m_flowInvalidRows.CompareMode = vbTextCompare
    m_flowBlockingErrors.CompareMode = vbTextCompare
    m_variableErrors.CompareMode = vbTextCompare
    m_variableWarnings.CompareMode = vbTextCompare

    BuildDisplayDepths
    BuildIfStructureAnalysis False
    BuildSubroutineFlowAnalysis
    BuildVariableAnalysis
    m_staticBuilt = True
End Sub

Public Function DisplayDepthAt(ByVal lineNo As Long) As Long
    m_source.ValidateLineRange lineNo
    DisplayDepthAt = m_displayDepths(lineNo)
End Function

Public Function StyleDepthAt(ByVal lineNo As Long) As Long
    m_source.ValidateLineRange lineNo
    StyleDepthAt = m_styleDepths(lineNo)
End Function

Public Function IsSubroutineStartAt(ByVal lineNo As Long) As Boolean
    m_source.ValidateLineRange lineNo
    IsSubroutineStartAt = m_subroutineStartRows.exists(CStr(lineNo))
End Function

Public Function IsSubroutineEndAt(ByVal lineNo As Long) As Boolean
    m_source.ValidateLineRange lineNo
    IsSubroutineEndAt = m_subroutineEndRows.exists(CStr(lineNo))
End Function

Public Function IsFlowInvalidAt(ByVal lineNo As Long) As Boolean
    m_source.ValidateLineRange lineNo
    IsFlowInvalidAt = m_flowInvalidRows.exists(CStr(lineNo))
End Function

Public Function VariableErrorAt(ByVal lineNo As Long) As String
    m_source.ValidateLineRange lineNo
    If m_variableErrors.exists(CStr(lineNo)) Then VariableErrorAt = CStr(m_variableErrors(CStr(lineNo)))
End Function

Public Function VariableWarningAt(ByVal lineNo As Long) As String
    m_source.ValidateLineRange lineNo
    If m_variableWarnings.exists(CStr(lineNo)) Then VariableWarningAt = CStr(m_variableWarnings(CStr(lineNo)))
End Function

Public Function ValidateStructure() As Boolean
    ClearValidationErrors
    m_index.BuildStructureIndex
    BuildIfStructureAnalysis True

    Dim blocks As Object
    Set blocks = m_index.SelectBlocks

    Dim key As Variant
    For Each key In blocks.keys
        Dim block As clsSelectBlock
        Set block = blocks(key)
        If block.EndSelectLine <= 0 Then AddValidationError "SELECT line " & key & " has no END_SELECT."
    Next key

    ' [C3] 루프 짝 불균형/교차, 고아 CASE·END_SELECT 승격 (Index 통합 스택 스캔이 기록)
    Dim pairKey As Variant
    For Each pairKey In m_index.BlockPairErrors.keys
        AddValidationError CStr(m_index.BlockPairErrors(pairKey))
    Next pairKey

    ValidateStructure = (m_validationErrors.count = 0)
End Function

Public Function ValidateLabels() As Boolean
    m_index.BuildLabelIndex
    m_index.BuildStateCoordinates
    BuildStaticAnalysis

    Dim R As Long
    For R = 1 To m_source.totalLines
        If m_source.CmdAt(R) = "LABEL" Then
            Dim rawLabel As String
            rawLabel = m_source.P3At(R)
            If Len(m_index.NormalizeLabelName(rawLabel)) = 0 Then
                AddValidationError "LABEL line " & R & " has empty name."
            ElseIf m_index.ContainsTokenSyntax(rawLabel) Then
                AddValidationError "LABEL line " & R & " must be a static symbol."
            End If
        End If
    Next R

    If m_index.HasDuplicateLabels Then AddValidationError "Duplicate LABEL: " & m_index.DuplicateLabels()
    ValidateJumpTargets
    PromoteFlowBlockingErrors
    ValidateLabels = (m_validationErrors.count = 0)
End Function

' 서브루틴 흐름 갭을 실행 전 차단 오류로 승격 (BuildStaticAnalysis가 채운 뒤 호출)
Private Sub PromoteFlowBlockingErrors()
    Dim key As Variant
    For Each key In m_flowBlockingErrors.keys
        AddValidationError CStr(m_flowBlockingErrors(key))
    Next key
End Sub

' 흐름 차단 오류를 서식용 집합과 승격용 메시지 사전에 함께 기록 (같은 라인 다중 사유는 이어붙임)
Private Sub MarkFlowBlocking(ByVal lineNo As Long, ByVal message As String)
    m_flowInvalidRows(CStr(lineNo)) = True
    If m_flowBlockingErrors.exists(CStr(lineNo)) Then
        m_flowBlockingErrors(CStr(lineNo)) = CStr(m_flowBlockingErrors(CStr(lineNo))) & vbCrLf & message
    Else
        m_flowBlockingErrors(CStr(lineNo)) = message
    End If
End Sub

Public Function ValidationErrors() As String
    Dim i As Long
    For i = 1 To m_validationErrors.count
        If Len(ValidationErrors) > 0 Then ValidationErrors = ValidationErrors & vbCrLf
        ValidationErrors = ValidationErrors & CStr(m_validationErrors.item(i))
    Next i
End Function

Private Sub BuildIfStructureAnalysis(ByVal reportValidation As Boolean)
    Dim stack() As IfValidationFrame
    Dim Top As Long
    Dim R As Long

    For R = 1 To m_source.totalLines
        If m_source.IsSkippable(R) Then GoTo NextRow

        Select Case m_source.CmdAt(R)
            Case "IF"
                Top = Top + 1
                ReDim Preserve stack(1 To Top)
                stack(Top).headerLine = R
                stack(Top).sawElse = False

            Case "ELSE_IF"
                If Top = 0 Then
                    MarkIfStructureIssue R, "ELSE_IF line " & R & " has no matching IF.", reportValidation
                ElseIf stack(Top).sawElse Then
                    MarkIfStructureIssue R, "ELSE_IF line " & R & " cannot appear after ELSE.", reportValidation
                End If
                ValidateElseIfParams R, reportValidation

            Case "ELSE"
                If Top = 0 Then
                    MarkIfStructureIssue R, "ELSE line " & R & " has no matching IF.", reportValidation
                ElseIf stack(Top).sawElse Then
                    MarkIfStructureIssue R, "ELSE line " & R & " duplicates an ELSE in the same IF block.", reportValidation
                Else
                    stack(Top).sawElse = True
                End If

            Case "END_IF"
                If Top = 0 Then
                    MarkIfStructureIssue R, "END_IF line " & R & " has no matching IF.", reportValidation
                Else
                    Top = Top - 1
                End If
        End Select

NextRow:
    Next R

    Do While Top > 0
        MarkIfStructureIssue stack(Top).headerLine, "IF line " & stack(Top).headerLine & " has no END_IF.", reportValidation
        Top = Top - 1
    Loop
End Sub

Private Sub ValidateElseIfParams(ByVal lineNo As Long, ByVal reportValidation As Boolean)
    If Len(Trim$(m_source.P1At(lineNo))) = 0 Then
        MarkIfStructureIssue lineNo, "ELSE_IF line " & lineNo & " has empty P1.", reportValidation
    End If
    If Len(Trim$(m_source.P2At(lineNo))) = 0 Then
        MarkIfStructureIssue lineNo, "ELSE_IF line " & lineNo & " has empty P2.", reportValidation
    ElseIf Not modCondition.IsConditionOperator(m_source.P2At(lineNo)) Then
        MarkIfStructureIssue lineNo, "ELSE_IF line " & lineNo & " has invalid condition operator: " & m_source.P2At(lineNo), reportValidation
    End If
    If Len(Trim$(m_source.P3At(lineNo))) = 0 Then
        MarkIfStructureIssue lineNo, "ELSE_IF line " & lineNo & " has empty P3.", reportValidation
    End If
End Sub

Private Sub MarkIfStructureIssue(ByVal lineNo As Long, ByVal message As String, ByVal reportValidation As Boolean)
    If lineNo > 0 Then m_flowInvalidRows(CStr(lineNo)) = True
    If reportValidation Then AddValidationError message
End Sub

Private Sub BuildDisplayDepths()
    Dim depth As Long
    Dim R As Long

    For R = 1 To m_source.totalLines
        Dim cmd As String
        cmd = m_source.CmdAt(R)

        If IsFormatBlockEnder(cmd) Then depth = IIf(depth > 0, depth - 1, 0)

        m_styleDepths(R) = depth
        m_displayDepths(R) = depth
        If cmd = "ELSE" Or cmd = "ELSE_IF" Or cmd = "CASE" Then m_displayDepths(R) = IIf(depth > 0, depth - 1, 0)

        If IsFormatBlockStarter(cmd) Then depth = depth + 1
    Next R
End Sub

Private Sub BuildSubroutineFlowAnalysis()
    MarkMissingJumpTargets

    Dim inSubroutineRows As Object
    Set inSubroutineRows = CreateObject("Scripting.Dictionary")
    inSubroutineRows.CompareMode = vbTextCompare

    Dim R As Long
    For R = 1 To m_source.totalLines
        If m_source.IsCommentLine(R) Then GoTo NextRow

        If m_source.CmdAt(R) = "LABEL" Then
            If m_index.IsCallTargetLabelAt(R) Then
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
                    MarkFlowBlocking R, "CALL 대상 LABEL line " & R & " 에 RETURN이 없습니다."
                End If

                If Not HasFlowBreakBefore(R) Then _
                    MarkFlowBlocking R, "CALL 대상 LABEL line " & R & " 앞에 흐름 차단(EXIT/GOTO/RETURN)이 없어 fall-through로 진입됩니다."
            End If
        End If
NextRow:
    Next R

    For R = 1 To m_source.totalLines
        If m_source.IsCommentLine(R) Then GoTo NextReturnRow
        If m_source.CmdAt(R) = "RETURN" Then
            If Not inSubroutineRows.exists(CStr(R)) Then _
                MarkFlowBlocking R, "RETURN line " & R & " 이 어떤 서브루틴에도 속하지 않습니다."
        End If
NextReturnRow:
    Next R

    For R = 1 To m_source.totalLines
        m_displayDepths(R) = m_displayDepths(R) + m_subDepths(R)
        m_styleDepths(R) = m_styleDepths(R) + m_subDepths(R)
    Next R
End Sub

Private Sub MarkMissingJumpTargets()
    Dim R As Long
    For R = 1 To m_source.totalLines
        If m_index.IsJumpTargetMissingAt(R) Then m_flowInvalidRows(CStr(R)) = True
    Next R
End Sub

Private Function FindSubroutineReturnLineForFlow(ByVal startLine As Long) As Long
    Dim labelDepth As Long
    labelDepth = m_displayDepths(startLine)

    Dim R As Long
    For R = startLine + 1 To m_source.totalLines
        If m_source.IsCommentLine(R) Then GoTo NextRow

        Dim cmd As String
        cmd = m_source.CmdAt(R)
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
        If m_source.IsCommentLine(R) Then GoTo NextRow

        Dim cmd As String
        cmd = m_source.CmdAt(R)
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
    For R = 1 To m_source.totalLines
        If m_source.IsCommentLine(R) Then GoTo NextRow

        Dim cmd As String
        cmd = m_source.CmdAt(R)
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
                    CollectUserVarsFromText m_source.ValueAt(R, idx + 2), uses

                    Dim varName As Variant
                    For Each varName In uses
                        used(CStr(varName)) = True
                        If Not defined.exists(CStr(varName)) Then
                            AddRowMessage m_variableErrors, R, UndefinedVarMessage(CStr(varName), defined.keys)
                        End If
                    Next varName
                End If
            End If
        Next idx

        If cmd = "SET" Or cmd = "INC" Or cmd = "DEC" Then
            Dim lvalueName As String
            lvalueName = MemoryLValueName(m_source.P1At(R))
            If Len(lvalueName) > 0 Then
                If (cmd = "INC" Or cmd = "DEC") And Not defined.exists(lvalueName) Then
                    AddRowMessage m_variableErrors, R, UndefinedVarMessage(lvalueName, defined.keys) & ". Initialize it with SET first."
                End If

                If Not defined.exists(lvalueName) Then defined.Add lvalueName, True
                AddDefRow defRows, lvalueName, R
            End If
        End If

NextRow:
    Next R

    Dim nameKey As Variant
    For Each nameKey In defRows.keys
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

Private Sub ValidateJumpTargets()
    Dim R As Long
    For R = 1 To m_source.totalLines
        Dim cmd As String
        cmd = m_source.CmdAt(R)

        ' [C2] 점프 명령 판별은 레지스트리 RawKind 태그 질의 (GOTO/CALL 이름 하드코딩 제거)
        Dim jumpDef As CmdDef
        Set jumpDef = modCmdRegistry.GetCmdDef(cmd)
        If jumpDef Is Nothing Then GoTo NextRow
        If jumpDef.p3 Is Nothing Then GoTo NextRow
        If jumpDef.p3.RawKind <> RK_JUMP_TARGET Then GoTo NextRow

        Dim rawTarget As String
        rawTarget = Trim$(m_source.P3At(R))

        If Len(rawTarget) = 0 Then
            AddValidationError cmd & " line " & R & " has empty target."
        ElseIf IsNumeric(rawTarget) Then
            AddValidationError cmd & " line " & R & " must target a LABEL, not a line number."
        ElseIf m_index.ContainsTokenSyntax(rawTarget) Then
            AddValidationError cmd & " line " & R & " must target a static LABEL."
        ElseIf Not m_index.HasLabel(rawTarget) Then
            AddValidationError cmd & " line " & R & " target not found: " & rawTarget
        Else
            Dim targetLine As Long
            targetLine = m_index.LabelLine(rawTarget)
            If targetLine > 0 Then
                If cmd = "GOTO" Then
                    If Not m_index.IsGotoAllowed(R, targetLine) Then AddValidationError "GOTO line " & R & " crosses invalid flow boundary."
                Else
                    ' [CALL 경계] sub-area 교차는 CALL의 존재 이유라 허용, 루프 본문 중간 진입만 차단
                    If Not m_index.IsCallAllowed(R, targetLine) Then AddValidationError "CALL line " & R & " enters a loop body from outside."
                End If
            End If
        End If

NextRow:
    Next R
End Sub

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
    suggestion = modStringDistance.ClosestName(varName, candidates)
    If Len(suggestion) > 0 Then msg = msg & ". Did you mean {" & suggestion & "}?"

    UndefinedVarMessage = msg
End Function

Private Sub ClearValidationErrors()
    Set m_validationErrors = New Collection
End Sub

Private Sub AddValidationError(ByVal message As String)
    m_validationErrors.Add message
End Sub
