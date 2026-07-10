Attribute VB_Name = "modDevTools"
' @desc 시트 어댑터 및 서식/표시 담당. 정적 의미 판단은 clsScriptProgram에 질의한다.
' @depends modCmdRegistry, modParse, clsScriptProgram
Option Explicit

Private m_lastValidationMessage As String
Private m_lastValidationWarning As String

' @desc 대상 시트(기본 ActiveSheet)의 스크립트를 검증 및 서식 적용함
Public Sub FormatRPAScript(Optional ByVal targetSheet As Worksheet = Nothing)
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim failNumber As Long
    Dim failSource As String
    Dim failDescription As String
    Dim failHelpFile As String
    Dim failHelpContext As Long
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    If targetSheet Is Nothing Then Set ws = ActiveSheet Else Set ws = targetSheet

    ' 0. 헤더 자동 생성
    Call EnsureHeaders(ws)

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    Dim R As Long, pureCmd As String
    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)

    ' [추가] 번호가 없는 부분(데이터 끝 이후) 서식 및 내용 초기화
    Dim clearLastRow As Long
    Dim clearCol As Long
    For clearCol = 1 To 6
        clearLastRow = Application.Max(clearLastRow, ws.Cells(ws.rows.count, clearCol).End(xlUp).Row)
    Next clearCol
    If clearLastRow > lastRow Then
        ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(clearLastRow, 6)).Clear
    End If

    If lastRow < 2 Then GoTo CleanExit

    ' 1. 기본 스타일 및 명확한 격자 초기화
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6))
        .Font.name = "돋움체"
        .Font.size = 10
        .Font.Bold = False
        .Font.Underline = xlUnderlineStyleNone
        .Interior.ColorIndex = xlNone

        ' --- [줄바꿈 방지 및 행 정렬] ---
        .WrapText = False        ' 셀 내부 줄바꿈 방지 (핵심!)
        .VerticalAlignment = xlCenter ' 텍스트를 셀 중앙에 정렬

        ' --- [격자 설정] ---
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .color = RGB(210, 210, 210)
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .color = RGB(210, 210, 210)
            .Weight = xlThin
        End With

        ' 외곽 테두리 제거
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
    End With

    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 6)).HorizontalAlignment = xlLeft

    ' 2. [추가] 행 높이 일괄 고정 (IDE의 일정한 간격 구현)
    ' 18~20 정도가 Consolas 10pt에 가장 보기 좋은 높이입니다.
    ws.Range("2:" & lastRow).RowHeight = 19

    For R = 2 To lastRow
        ' 0. 줄번호 자동 생성
        ws.Cells(R, 1).value = R - 1

        pureCmd = modCmdRegistry.NormalizeCommand(ws.Cells(R, 2).value)
        If Len(Trim$(CStr(ws.Cells(R, 2).value))) = 0 Then
            Call ApplyJumpTargetLinkStyle(ws.Cells(R, 5), pureCmd)
            Call ApplyLabelRoleDisplay(ws, R, pureCmd, program)
            GoTo NextRow
        End If
        Dim rowIsComment As Boolean
        Dim rowLooksLikeIndentedComment As Boolean
        rowIsComment = IsCommentRow(ws, R)
        rowLooksLikeIndentedComment = (Not rowIsComment And Left$(pureCmd, 1) = "#")

                Dim printDepth As Long
                Dim styleDepth As Long
                If Not program Is Nothing Then
                    printDepth = program.DisplayDepthAtSheetRow(R)
                    styleDepth = program.StyleDepthAtSheetRow(R)
                End If

                    ' 3. 명령어 텍스트 및 IDE 스타일 적용
                    If rowIsComment Then
                        ws.Cells(R, 2).value = FormatCommentCommandText(ws.Cells(R, 2).value, printDepth)
                    ElseIf Not rowLooksLikeIndentedComment Then
                        ws.Cells(R, 2).value = FormatCommandText(pureCmd, printDepth)
                    End If

                    ' [추가] 문법 및 파라미터 검사
                    Dim isValid As Boolean
                    If rowIsComment Then
                        isValid = True
                    Else
                        isValid = ValidateLine(pureCmd, ws.Cells(R, 3).value, ws.Cells(R, 4).value, ws.Cells(R, 5).value)
                    End If
                    If Not program Is Nothing Then
                        If program.IsFlowInvalidAtSheetRow(R) Then isValid = False
                        If Len(program.VariableErrorAtSheetRow(R)) > 0 Then isValid = False
                    End If

                    Dim hasWarning As Boolean
                    If Not program Is Nothing Then
                        hasWarning = (Len(program.VariableWarningAtSheetRow(R)) > 0 Or HasLabelRoleWarning(program, R))
                    End If

                    ' [통합 스타일 호출] 테두리 가이드와 배경색 처리
                    Dim isSubStart As Boolean
                    Dim isSubEnd As Boolean
                    If Not program Is Nothing Then
                        isSubStart = program.IsSubroutineStartAtSheetRow(R)
                        isSubEnd = program.IsSubroutineEndAtSheetRow(R)
                    End If

                    Call ApplyUnifiedStyle(ws.Cells(R, 2).Resize(1, 4), pureCmd, styleDepth, isValid, rowIsComment, hasWarning)
                    Call ApplySubroutineBoundaryStyle(ws.Cells(R, 2).Resize(1, 4), _
                                                       isSubStart, _
                                                       isSubEnd, _
                                                       Not isValid)
                    Call ApplyJumpTargetLinkStyle(ws.Cells(R, 5), pureCmd)
                    If rowIsComment Then
                        Call ClearLabelRoleDisplay(ws, R)
                    Else
                        Call ApplyLabelRoleDisplay(ws, R, pureCmd, program)
                    End If

NextRow:
                    Next R

CleanExit:
                    Application.Calculation = prevCalculation
                    Application.EnableEvents = prevEnableEvents
                    Application.ScreenUpdating = prevScreenUpdating
                    Exit Sub

CleanFail:
                    failNumber = Err.Number
                    failSource = Err.Source
                    failDescription = Err.Description
                    failHelpFile = Err.HelpFile
                    failHelpContext = Err.HelpContext
                    On Error Resume Next
                    Application.Calculation = prevCalculation
                    Application.EnableEvents = prevEnableEvents
                    Application.ScreenUpdating = prevScreenUpdating
                    On Error GoTo 0
                    Err.Raise failNumber, failSource, failDescription, failHelpFile, failHelpContext
End Sub

' [신규] 헤더가 없으면 생성
Public Sub EnsureHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("ROW", "COMMAND", "PARAM1", "PARAM2", "PARAM3", "COMMENT")

    With ws.Range("A1:F1")
        .value = headers
        .Font.name = "돋움체"
        .Font.size = 10
        .Font.Bold = True
        .Font.color = vbBlack
        .Interior.color = RGB(240, 240, 240)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C:E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 30
End Sub

' [파생] 명령어 및 파라미터 유효성 검사 → GetCmdDef 게이트로 환원
Public Function ValidateLine(cmd As String, p1 As Variant, p2 As Variant, p3 As Variant) As Boolean
    m_lastValidationMessage = ""
    m_lastValidationWarning = ""

    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(cmd)

    ' 1. 명령어 존재 여부 = 레지스트리 등록 여부
    If def Is Nothing Then
        SetValidationError "Unknown command: " & modCmdRegistry.NormalizeCommand(cmd)
        ValidateLine = False
        Exit Function
    End If

    ValidateLine = True

    If Not ValidateParamSlot(def, 1, p1) Then ValidateLine = False
    If Not ValidateParamSlot(def, 2, p2) Then ValidateLine = False
    If Not ValidateParamSlot(def, 3, p3) Then ValidateLine = False

    If Not ValidateParamShift(def, p1, p2, p3) Then ValidateLine = False
End Function

Public Function LastValidationMessage() As String
    LastValidationMessage = m_lastValidationMessage
End Function

Public Function LastValidationWarning() As String
    LastValidationWarning = m_lastValidationWarning
End Function

Private Function ValidateParamSlot(ByVal def As CmdDef, ByVal idx As Long, ByVal rawValue As Variant) As Boolean
    ValidateParamSlot = True

    Dim text As String
    Dim numProbe As Double
    text = ParamText(rawValue)

    Dim spec As ParamSpec
    Set spec = def.Param(idx)

    If spec Is Nothing Then
        If Len(text) > 0 Then
            SetValidationError def.Canonical & " does not use P" & idx & "."
            ValidateParamSlot = False
        End If
        Exit Function
    End If

    If spec.Required And Len(text) = 0 Then
        SetValidationError def.Canonical & " requires P" & idx & " (" & spec.Meaning & ")."
        ValidateParamSlot = False
        Exit Function
    End If

    If Len(text) = 0 Then Exit Function

    ' [C2] RAW 슬롯 시맨틱은 레지스트리(ParamSpec.RawKind)가 선언 - 명령어명 하드코딩 금지
    Select Case spec.RawKind
        Case RK_JUMP_TARGET
            If modParse.TryParseDoubleStrict(text, numProbe) Or ContainsTokenSyntax(text) Then
                SetValidationError def.Canonical & " target must be a static LABEL."
                ValidateParamSlot = False
            End If
        Case RK_CONDITION
            If Not modCondition.IsConditionOperator(text) Then
                SetValidationError "Invalid condition operator: " & text
                ValidateParamSlot = False
            End If
        Case RK_LVALUE
            If Not IsWritableLValue(text) Then
                SetValidationError "Invalid writable target: " & text
                ValidateParamSlot = False
            End If
        Case RK_TEXT
            If ContainsTokenSyntax(text) Then
                SetValidationError def.Canonical & " P" & idx & " must be raw text, not a token."
                ValidateParamSlot = False
            End If
    End Select

    If spec.DataType = DT_NUM Then
        If Not ContainsTokenSyntax(text) Then
            ' 런타임 게이트(ValidateDecodedSlots)와 동일 파서로 판정
            If Not modParse.TryParseDoubleStrict(text, numProbe) Then
                SetValidationError def.Canonical & " P" & idx & " must be numeric."
                ValidateParamSlot = False
            End If
        End If
    End If
End Function

Private Function ValidateParamShift(ByVal def As CmdDef, ByVal p1 As Variant, ByVal p2 As Variant, ByVal p3 As Variant) As Boolean
    ValidateParamShift = True

    Dim vals(1 To 3) As String
    vals(1) = ParamText(p1)
    vals(2) = ParamText(p2)
    vals(3) = ParamText(p3)

    Dim idx As Long, later As Long
    For idx = 1 To 2
        Dim spec As ParamSpec
        Set spec = def.Param(idx)
        If Not spec Is Nothing Then
            If spec.Required And Len(vals(idx)) = 0 Then
                For later = idx + 1 To 3
                    If Len(vals(later)) > 0 Then
                        SetValidationError def.Canonical & " P" & idx & " is blank, but a later parameter has a value."
                        ValidateParamShift = False
                        Exit Function
                    End If
                Next later
            End If
        End If
    Next idx
End Function

Private Function ParamText(ByVal rawValue As Variant) As String
    On Error GoTo SafeExit
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    ParamText = Trim$(CStr(rawValue))
SafeExit:
End Function

Private Function ContainsTokenSyntax(ByVal text As String) As Boolean
    ContainsTokenSyntax = (InStr(1, text, "{", vbTextCompare) > 0 Or InStr(1, text, "}", vbTextCompare) > 0)
End Function

Private Function IsWritableLValue(ByVal rawP1 As String) As Boolean
    Dim s As String
    s = Trim$(rawP1)
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "{" Or Right$(s, 1) <> "}" Then Exit Function

    Dim inner As String
    inner = Trim$(Mid$(s, 2, Len(s) - 2))
    If Len(inner) = 0 Then Exit Function

    If UCase$(inner) = "$CLIP" Then
        IsWritableLValue = True
    ElseIf Left$(inner, 1) <> "$" Then
        IsWritableLValue = True
    End If
End Function

Private Sub SetValidationError(ByVal msg As String)
    If Len(m_lastValidationMessage) = 0 Then m_lastValidationMessage = msg
End Sub

Public Function ValidateSheetFlowAt(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    ValidateSheetFlowAt = True
    m_lastValidationWarning = ""
    If ws Is Nothing Then Exit Function
    If rowNum < 2 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    If rowNum > lastRow Then Exit Function

    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)
    If Not program Is Nothing Then
        ValidateSheetFlowAt = Not program.IsFlowInvalidAtSheetRow(rowNum)
        m_lastValidationWarning = program.LabelRoleWarningAtSheetRow(rowNum)
    End If
End Function

Public Function ValidateSheetVariablesAt(ByVal ws As Worksheet, ByVal rowNum As Long, Optional ByVal preserveExistingWarning As Boolean = False) As Boolean
    Dim existingWarning As String
    If preserveExistingWarning Then existingWarning = m_lastValidationWarning

    m_lastValidationMessage = ""
    m_lastValidationWarning = ""
    ValidateSheetVariablesAt = True

    If ws Is Nothing Then Exit Function
    If rowNum < 2 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    If rowNum > lastRow Then Exit Function

    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)
    If program Is Nothing Then Exit Function

    Dim variableError As String
    variableError = program.VariableErrorAtSheetRow(rowNum)
    If Len(variableError) > 0 Then
        SetValidationError variableError
        ValidateSheetVariablesAt = False
    End If

    Dim variableWarning As String
    variableWarning = program.VariableWarningAtSheetRow(rowNum)
    If Len(variableWarning) > 0 Then
        If Len(existingWarning) > 0 Then
            m_lastValidationWarning = existingWarning & vbCrLf & variableWarning
        Else
            m_lastValidationWarning = variableWarning
        End If
    ElseIf Len(existingWarning) > 0 Then
        m_lastValidationWarning = existingWarning
    End If
End Function

Public Function ResolveLabelTargetRow(ByVal ws As Worksheet, ByVal targetName As String) As Long
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row

    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)
    If Not program Is Nothing Then ResolveLabelTargetRow = program.ResolveJumpTargetSheetRow(targetName)
End Function

Public Function ResolveFirstCallSourceRow(ByVal ws As Worksheet, ByVal labelName As String) As Long
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row

    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)
    If Not program Is Nothing Then ResolveFirstCallSourceRow = program.FirstCallSourceSheetRowForLabel(labelName)
End Function

Public Function RefreshLabelRoleDisplay(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    If ws Is Nothing Then Exit Function
    If rowNum < 2 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    If rowNum > lastRow Then Exit Function

    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)
    Call ApplyLabelRoleDisplay(ws, rowNum, PureCommandAt(ws, rowNum), program)
    RefreshLabelRoleDisplay = True
End Function

Public Sub ClearLabelRoleDisplay(ByVal ws As Worksheet, ByVal rowNum As Long)
    If ws Is Nothing Then Exit Sub
    If rowNum < 2 Then Exit Sub

    modScriptAssist.ClearLabelRoleNote ws.Cells(rowNum, 2)
    ClearEditorLinkStyle ws.Cells(rowNum, 5), True
End Sub

Private Function ProgramSnapshotFromSheet(ByVal ws As Worksheet, ByVal lastRow As Long) As clsScriptProgram
    If ws Is Nothing Then Exit Function
    If lastRow < 2 Then Exit Function

    Dim program As clsScriptProgram
    Set program = New clsScriptProgram
    program.LoadFromRange ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6))
    Set ProgramSnapshotFromSheet = program
End Function

Private Function HasLabelRoleWarning(ByVal program As clsScriptProgram, ByVal sheetRow As Long) As Boolean
    If program Is Nothing Then Exit Function
    HasLabelRoleWarning = (Len(program.LabelRoleWarningAtSheetRow(sheetRow)) > 0)
End Function

' 통합 스타일 엔진: 하이라이트 강제 초기화 및 스타일 재적용
Public Sub ApplyUnifiedStyle(rng As Range, cmd As String, depth As Long, Optional isValid As Boolean = True, Optional isComment As Boolean = False, Optional hasWarning As Boolean = False)
    Dim cmdCell As Range: Set cmdCell = rng.Cells(1, 1)

    '----------------------------------------------------------
    ' [0. 서식 초기화] 하이라이트(노란 배경, 빨간 테두리) 흔적 지우기
    '----------------------------------------------------------
    With rng
        .Interior.ColorIndex = xlNone  ' 배경색 투명화 (노란색 제거)
        .Font.Bold = False             ' 굵기 초기화
        .Font.color = vbBlack          ' 글자색 초기화

        ' 테두리 초기화 (격자 색상 RGB 210으로 복구)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(210, 210, 210)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(210, 210, 210)
        End With
        ' 세로 격자선 복구 (셀 사이 줄)
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(210, 210, 210)
        End With
        ' 왼쪽/오른쪽 테두리 일단 제거
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
    End With

    '----------------------------------------------------------
    ' 1. [구조] 수직 가이드라인 (더 두꺼운 테두리)
    '----------------------------------------------------------
    If depth > 0 Then
        With cmdCell.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .color = RGB(130, 130, 130)
        End With
    End If

    '----------------------------------------------------------
    ' [오류 처리] 유효하지 않은 명령어는 빨간색 경고
    '----------------------------------------------------------
    If Not isValid Then
        cmdCell.Font.color = vbRed
        cmdCell.Font.Bold = True
    End If

    '----------------------------------------------------------
    ' 2. [주석 처리]
    '----------------------------------------------------------
    If isComment Then
        rng.Font.color = RGB(180, 180, 180)
     Exit Sub
    End If

    '----------------------------------------------------------
    ' 3. [그룹별 컬러링] 다시 입히기
    '----------------------------------------------------------
    cmdCell.Font.Bold = True ' 명령어는 다시 굵게

    ' [파생] ColorGroup(무슨 색) / PaintScope(칠하는 범위) ? 두 독립 축을 게이트에서 조회
    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(cmd)
    If def Is Nothing Then
        ' 미등록 명령어: 블록 내부면 연녹 배경만 유지 (구 Case Else)
        If depth > 0 Then rng.Interior.color = RGB(245, 255, 245)
        Exit Sub
    End If

    Dim clr As Long: clr = ColorOf(def.ColorGroup)
    If def.PaintScope = FULLROW Then
        ' [Flow/System] 줄 전체
        rng.Interior.color = clr
    Else
        ' [Action/Vision/Memory] B열 강조 + 중첩 시 C-E 연녹 배경
        cmdCell.Interior.color = clr
        If depth > 0 Then rng.offset(0, 1).Resize(1, 3).Interior.color = RGB(245, 255, 245)
    End If

    If isValid And hasWarning Then
        cmdCell.Font.color = RGB(156, 101, 0)
    End If
End Sub

' [파생 헬퍼] ColorGroup enum -> 현행 팔레트 RGB (값 보존)
Private Function ColorOf(ByVal cg As eColorGroup) As Long
    Select Case cg
        Case clrFLOW:   ColorOf = RGB(226, 239, 218)  ' 연녹
        Case clrACTION: ColorOf = RGB(221, 235, 247)  ' 연청
        Case clrVISION: ColorOf = RGB(255, 242, 204)  ' 연황
        Case clrSYSTEM: ColorOf = RGB(242, 242, 242)  ' 연회
        Case Else:      ColorOf = RGB(245, 255, 245)
    End Select
End Function

Private Sub ApplySubroutineBoundaryStyle(ByVal rng As Range, ByVal isStart As Boolean, _
                                         ByVal isEnd As Boolean, ByVal isInvalid As Boolean)
    If Not isStart And Not isEnd Then Exit Sub

    Dim borderColor As Long
    If isInvalid Then
        borderColor = RGB(192, 0, 0)
    Else
        borderColor = RGB(90, 90, 90)
    End If

    If isStart Then
        With rng.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .color = borderColor
        End With
    End If

    If isEnd Then
        With rng.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .color = borderColor
        End With
    End If
End Sub

Private Sub ApplyJumpTargetLinkStyle(ByVal targetCell As Range, ByVal cmd As String)
    If targetCell Is Nothing Then Exit Sub

    If targetCell.Hyperlinks.count > 0 Then targetCell.Hyperlinks.Delete

    ClearEditorLinkStyle targetCell
    If cmd = "CALL" Or cmd = "GOTO" Then ApplyEditorLinkStyle targetCell
End Sub

Private Sub ApplyLabelRoleDisplay(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal cmd As String, _
                                  ByVal program As clsScriptProgram)
    Dim roleCell As Range
    Set roleCell = ws.Cells(rowNum, 2)

    Dim labelCell As Range
    Set labelCell = ws.Cells(rowNum, 5)

    If cmd <> "LABEL" Then
        modScriptAssist.ClearLabelRoleNote roleCell
        ClearEditorLinkStyle labelCell
        Exit Sub
    End If

    If program Is Nothing Then
        ClearEditorLinkStyle labelCell, True
        Exit Sub
    End If

    Dim roleName As String
    roleName = program.LabelRoleAtSheetRow(rowNum)
    If Len(roleName) = 0 Then
        modScriptAssist.ClearLabelRoleNote roleCell
        ClearEditorLinkStyle labelCell, True
        Exit Sub
    End If

    Dim warning As String
    warning = program.LabelRoleWarningAtSheetRow(rowNum)

    If roleName = "SubroutineLabel" And Len(warning) = 0 Then
        modScriptAssist.ClearLabelRoleNote roleCell
    Else
        modScriptAssist.ApplyLabelRoleNote roleCell, roleName, warning
    End If

    roleCell.Font.Italic = (roleName = "DeadLabel")
    If roleName = "MixedLabel" Then
        roleCell.Font.color = RGB(156, 101, 0)
    ElseIf roleName = "DeadLabel" Then
        roleCell.Font.color = RGB(128, 128, 128)
    End If

    If roleName = "SubroutineLabel" Or roleName = "MixedLabel" Then
        ApplyEditorLinkStyle labelCell
    Else
        ClearEditorLinkStyle labelCell, True
    End If
End Sub

Private Sub ApplyEditorLinkStyle(ByVal targetCell As Range)
    If targetCell Is Nothing Then Exit Sub
    If Len(ParamText(targetCell.value)) = 0 Then Exit Sub

    With targetCell.Font
        .color = RGB(5, 99, 193)
        .Underline = xlUnderlineStyleSingle
    End With
End Sub

Private Sub ClearEditorLinkStyle(ByVal targetCell As Range, Optional ByVal resetColor As Boolean = False)
    If targetCell Is Nothing Then Exit Sub
    With targetCell.Font
        If resetColor Then .color = vbBlack
        .Underline = xlUnderlineStyleNone
    End With
End Sub

Private Function PureCommandAt(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    PureCommandAt = modCmdRegistry.NormalizeCommand(ws.Cells(rowNum, 2).value)
End Function

Private Function IsCommentRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim rawCmd As String
    rawCmd = Trim$(CStr(ws.Cells(rowNum, 2).value))
    IsCommentRow = (Len(rawCmd) > 0 And Left$(rawCmd, 1) = "#")
End Function

Private Function FormatCommandText(ByVal cmd As String, ByVal depth As Long) As String
    FormatCommandText = String(depth, ">") & IIf(depth > 0, " ", "") & cmd
End Function

Private Function FormatCommentCommandText(ByVal rawCmd As Variant, ByVal depth As Long) As String
    Dim body As String
    body = modCmdRegistry.NormalizeCommand(rawCmd)
    ' 주석은 "#"이 마커보다 앞이라 NormalizeCommand가 못 벗긴 들여쓰기가 남는다
    If Left$(body, 1) = "#" Then body = modCmdRegistry.StripIndentMarkers(Mid$(body, 2))

    If Len(body) = 0 Then
        FormatCommentCommandText = "#"
    ElseIf depth > 0 Then
        FormatCommentCommandText = "# " & String(depth, ">") & " " & body
    Else
        FormatCommentCommandText = "# " & body
    End If
End Function
