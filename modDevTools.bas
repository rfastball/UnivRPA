Attribute VB_Name = "modDevTools"
' @desc RPA 스크립트 시트 서식·검증을 레지스트리 파생으로 적용 담당
' @depends modCmdRegistry
Option Explicit

Private m_lastValidationMessage As String
Private m_lastValidationWarning As String

Public Sub FormatRPAScript()
Attribute FormatRPAScript.VB_ProcData.VB_Invoke_Func = "F\n14"
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

    Dim ws As Worksheet: Set ws = ActiveSheet

    ' 0. 헤더 자동 생성
    Call EnsureHeaders(ws)

    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    Dim R As Long, pureCmd As String, depth As Integer: depth = 0
    Dim baseDepths() As Long
    Dim subDepths() As Long
    Dim subStartRows As Object
    Dim subEndRows As Object
    Dim invalidRows As Object
    Dim varInvalidRows As Object
    Dim varWarningRows As Object
    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)
    Call AnalyzeSubroutineFlow(ws, lastRow, program, baseDepths, subDepths, subStartRows, subEndRows, invalidRows)
    Call AnalyzeSheetVariables(ws, lastRow, varInvalidRows, varWarningRows)

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
        .Font.Size = 10
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

            ' 2. 깊이 계산
            If IsBlockEnder(pureCmd) Then depth = IIf(depth > 0, depth - 1, 0)

                ' [수정] Else 문 들여쓰기 보정 (출력 시에만 깊이 -1 적용)
                Dim printDepth As Integer: printDepth = depth
                If pureCmd = "ELSE" Or pureCmd = "CASE" Then printDepth = IIf(depth > 0, depth - 1, 0)
                printDepth = printDepth + subDepths(R)

                    ' 3. 명령어 텍스트 및 IDE 스타일 적용
                    If Not rowLooksLikeIndentedComment Then
                        ws.Cells(R, 2).value = String(printDepth, ">") & (IIf(printDepth > 0, " ", "")) & pureCmd
                    End If

                    ' [추가] 문법 및 파라미터 검사
                    Dim isValid As Boolean
                    If rowIsComment Then
                        isValid = True
                    Else
                        isValid = ValidateLine(pureCmd, ws.Cells(R, 3).value, ws.Cells(R, 4).value, ws.Cells(R, 5).value)
                    End If
                    If invalidRows.exists(CStr(R)) Then isValid = False
                    If varInvalidRows.exists(CStr(R)) Then isValid = False

                    Dim hasWarning As Boolean
                    hasWarning = varWarningRows.exists(CStr(R)) Or HasLabelRoleWarning(program, R)

                    ' [통합 스타일 호출] 테두리 가이드와 배경색 처리
                    Call ApplyUnifiedStyle(ws.Cells(R, 2).Resize(1, 4), pureCmd, depth + subDepths(R), isValid, rowIsComment, hasWarning)
                    Call ApplySubroutineBoundaryStyle(ws.Cells(R, 2).Resize(1, 4), _
                                                       subStartRows.exists(CStr(R)), _
                                                       subEndRows.exists(CStr(R)), _
                                                       Not isValid)
                    Call ApplyJumpTargetLinkStyle(ws.Cells(R, 5), pureCmd)
                    Call ApplyLabelRoleDisplay(ws, R, pureCmd, program)

                    ' 4. 깊이 증가
                    If IsBlockStarter(pureCmd) Then depth = depth + 1

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
        .Font.Size = 10
        .Font.Bold = True
        .Font.color = vbBlack
        .Interior.color = RGB(240, 240, 240)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 35
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

    If IsJumpTargetSlot(def, idx) Then
        If IsNumeric(text) Or ContainsTokenSyntax(text) Then
            SetValidationError def.Canonical & " target must be a static LABEL."
            ValidateParamSlot = False
        End If
    End If

    If spec.IsRaw Then
        If IsConditionOperatorSlot(def, idx) Then
            If Not modCondition.IsConditionOperator(text) Then
                SetValidationError "Invalid condition operator: " & text
                ValidateParamSlot = False
            End If
        ElseIf IsMemoryLValueSlot(def, idx) Then
            If Not IsWritableLValue(text) Then
                SetValidationError "Invalid writable target: " & text
                ValidateParamSlot = False
            End If
        ElseIf ContainsTokenSyntax(text) Then
            SetValidationError def.Canonical & " P" & idx & " must be raw text, not a token."
            ValidateParamSlot = False
        End If
    End If

    If spec.DataType = DT_NUM Then
        If Not ContainsTokenSyntax(text) Then
            If Not IsNumeric(text) Then
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

Private Function IsConditionOperatorSlot(ByVal def As CmdDef, ByVal idx As Long) As Boolean
    IsConditionOperatorSlot = ((def.Canonical = "IF" Or def.Canonical = "WHILE") And idx = 2)
End Function

Private Function IsMemoryLValueSlot(ByVal def As CmdDef, ByVal idx As Long) As Boolean
    IsMemoryLValueSlot = ((def.Canonical = "SET" Or def.Canonical = "INC" Or def.Canonical = "DEC") And idx = 1)
End Function

Private Function IsJumpTargetSlot(ByVal def As CmdDef, ByVal idx As Long) As Boolean
    IsJumpTargetSlot = ((def.Canonical = "GOTO" Or def.Canonical = "CALL") And idx = 3)
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

    Dim baseDepths() As Long
    Dim subDepths() As Long
    Dim subStartRows As Object
    Dim subEndRows As Object
    Dim invalidRows As Object
    Dim program As clsScriptProgram
    Set program = ProgramSnapshotFromSheet(ws, lastRow)
    Call AnalyzeSubroutineFlow(ws, lastRow, program, baseDepths, subDepths, subStartRows, subEndRows, invalidRows)

    ValidateSheetFlowAt = Not invalidRows.exists(CStr(rowNum))
    If Not program Is Nothing Then m_lastValidationWarning = program.LabelRoleWarningAtSheetRow(rowNum)
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

    Dim invalidRows As Object
    Dim warningRows As Object
    Call AnalyzeSheetVariables(ws, lastRow, invalidRows, warningRows)

    If invalidRows.exists(CStr(rowNum)) Then
        SetValidationError CStr(invalidRows(CStr(rowNum)))
        ValidateSheetVariablesAt = False
    End If

    If warningRows.exists(CStr(rowNum)) Then
        If Len(existingWarning) > 0 Then
            m_lastValidationWarning = existingWarning & vbCrLf & CStr(warningRows(CStr(rowNum)))
        Else
            m_lastValidationWarning = CStr(warningRows(CStr(rowNum)))
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

Public Function ScrollToJumpTarget(ByVal sourceParamCell As Range) As Boolean
    If sourceParamCell Is Nothing Then Exit Function
    If sourceParamCell.Cells.count <> 1 Then Exit Function

    Dim ws As Worksheet
    Set ws = sourceParamCell.Worksheet

    Dim cmd As String
    cmd = PureCommandAt(ws, sourceParamCell.Row)
    If cmd <> "CALL" And cmd <> "GOTO" Then Exit Function

    Dim targetRow As Long
    targetRow = ResolveLabelTargetRow(ws, CStr(sourceParamCell.value))
    If targetRow = 0 Then Exit Function

    Application.Goto ws.Cells(targetRow, 5), True
    ScrollToJumpTarget = True
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

Private Function IsCallTargetLabelRow(ByVal program As clsScriptProgram, ByVal sheetRow As Long) As Boolean
    If program Is Nothing Then Exit Function
    IsCallTargetLabelRow = program.IsCallTargetLabelAtSheetRow(sheetRow)
End Function

' 통합 스타일 엔진: 하이라이트 강제 초기화 및 스타일 재적용
Public Sub ApplyUnifiedStyle(rng As Range, cmd As String, depth As Integer, Optional isValid As Boolean = True, Optional isComment As Boolean = False, Optional hasWarning As Boolean = False)
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

Private Sub AnalyzeSheetVariables(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                  ByRef invalidRows As Object, ByRef warningRows As Object)
    Set invalidRows = CreateObject("Scripting.Dictionary")
    Set warningRows = CreateObject("Scripting.Dictionary")
    invalidRows.CompareMode = vbTextCompare
    warningRows.CompareMode = vbTextCompare

    If ws Is Nothing Or lastRow < 2 Then Exit Sub

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
    For R = 2 To lastRow
        If IsCommentRow(ws, R) Then GoTo NextRow

        Dim cmd As String
        cmd = PureCommandAt(ws, R)
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
                    CollectUserVarsFromText ParamText(ws.Cells(R, idx + 2).value), uses

                    Dim varName As Variant
                    For Each varName In uses
                        used(CStr(varName)) = True
                        If Not defined.exists(CStr(varName)) Then
                            AddRowMessage invalidRows, R, UndefinedVarMessage(CStr(varName), defined.Keys)
                        End If
                    Next varName
                End If
            End If
        Next idx

        If cmd = "SET" Or cmd = "INC" Or cmd = "DEC" Then
            Dim lvalueName As String
            lvalueName = MemoryLValueName(ParamText(ws.Cells(R, 3).value))
            If Len(lvalueName) > 0 Then
                If (cmd = "INC" Or cmd = "DEC") And Not defined.exists(lvalueName) Then
                    AddRowMessage invalidRows, R, UndefinedVarMessage(lvalueName, defined.Keys) & ". Initialize it with SET first."
                End If

                If Not defined.exists(lvalueName) Then
                    defined.Add lvalueName, True
                End If
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
                If Not invalidRows.exists(CStr(rowItem)) Then
                    AddRowMessage warningRows, CLng(rowItem), "Warning: defined variable is not used: {" & CStr(nameKey) & "}"
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

Private Sub AddDefRow(ByVal defRows As Object, ByVal varName As String, ByVal rowNum As Long)
    Dim rows As Collection
    If defRows.exists(varName) Then
        Set rows = defRows(varName)
    Else
        Set rows = New Collection
        defRows.Add varName, rows
    End If
    rows.Add rowNum
End Sub

Private Sub AddRowMessage(ByVal rows As Object, ByVal rowNum As Long, ByVal message As String)
    Dim key As String
    key = CStr(rowNum)

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

Private Sub AnalyzeSubroutineFlow(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal program As clsScriptProgram, _
                                  ByRef baseDepths() As Long, _
                                  ByRef subDepths() As Long, _
                                  ByRef subStartRows As Object, _
                                  ByRef subEndRows As Object, _
                                  ByRef invalidRows As Object)
    ReDim baseDepths(1 To lastRow)
    ReDim subDepths(1 To lastRow)

    Set subStartRows = CreateObject("Scripting.Dictionary")
    Set subEndRows = CreateObject("Scripting.Dictionary")
    Set invalidRows = CreateObject("Scripting.Dictionary")
    subStartRows.CompareMode = vbTextCompare
    subEndRows.CompareMode = vbTextCompare
    invalidRows.CompareMode = vbTextCompare

    If lastRow < 2 Then Exit Sub

    Call ComputeBaseDepths(ws, lastRow, baseDepths)
    Call MarkMissingJumpTargets(ws, lastRow, program, invalidRows)

    Dim inSubroutineRows As Object
    Set inSubroutineRows = CreateObject("Scripting.Dictionary")
    inSubroutineRows.CompareMode = vbTextCompare

    Dim R As Long
    For R = 2 To lastRow
        If IsCommentRow(ws, R) Then GoTo NextRow
        If PureCommandAt(ws, R) = "LABEL" Then
            If IsCallTargetLabelRow(program, R) Then
                subStartRows(CStr(R)) = True

                Dim endRow As Long
                endRow = FindSubroutineReturn(ws, R, lastRow, baseDepths(R), baseDepths)

                If endRow > 0 Then
                    subEndRows(CStr(endRow)) = True
                    inSubroutineRows(CStr(R)) = True
                    inSubroutineRows(CStr(endRow)) = True

                    Dim innerRow As Long
                    For innerRow = R + 1 To endRow - 1
                        subDepths(innerRow) = subDepths(innerRow) + 1
                        inSubroutineRows(CStr(innerRow)) = True
                    Next innerRow
                Else
                    invalidRows(CStr(R)) = True
                End If

                If Not HasFlowBreakBefore(ws, R) Then invalidRows(CStr(R)) = True
            End If
        End If
NextRow:
    Next R

    For R = 2 To lastRow
        If IsCommentRow(ws, R) Then GoTo NextReturnRow
        If PureCommandAt(ws, R) = "RETURN" Then
            If Not inSubroutineRows.exists(CStr(R)) Then invalidRows(CStr(R)) = True
        End If
NextReturnRow:
    Next R
End Sub

Private Sub ComputeBaseDepths(ByVal ws As Worksheet, ByVal lastRow As Long, ByRef baseDepths() As Long)
    Dim depth As Long
    Dim R As Long

    For R = 2 To lastRow
        Dim cmd As String
        cmd = PureCommandAt(ws, R)
        If IsCommentRow(ws, R) Then GoTo NextRow

        If IsBlockEnder(cmd) Then depth = IIf(depth > 0, depth - 1, 0)

        baseDepths(R) = depth
        If cmd = "ELSE" Or cmd = "CASE" Then baseDepths(R) = IIf(depth > 0, depth - 1, 0)

        If IsBlockStarter(cmd) Then depth = depth + 1
NextRow:
    Next R
End Sub

Private Sub MarkMissingJumpTargets(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                   ByVal program As clsScriptProgram, ByVal invalidRows As Object)
    If program Is Nothing Then Exit Sub

    Dim R As Long
    For R = 2 To lastRow
        If program.IsJumpTargetMissingAtSheetRow(R) Then invalidRows(CStr(R)) = True
    Next R
End Sub

Private Function FindSubroutineReturn(ByVal ws As Worksheet, ByVal startRow As Long, _
                                      ByVal lastRow As Long, ByVal labelDepth As Long, _
                                      ByRef baseDepths() As Long) As Long
    Dim R As Long
    For R = startRow + 1 To lastRow
        If IsCommentRow(ws, R) Then GoTo NextRow
        Dim cmd As String
        cmd = PureCommandAt(ws, R)
        If Not IsExecutableCommand(cmd) Then GoTo NextRow

        If baseDepths(R) < labelDepth Then Exit Function

        If baseDepths(R) = labelDepth Then
            Select Case cmd
                Case "RETURN"
                    FindSubroutineReturn = R
                    Exit Function
                Case "LABEL"
                    Exit Function
            End Select
        End If

NextRow:
    Next R
End Function

Private Function HasFlowBreakBefore(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim R As Long
    For R = rowNum - 1 To 2 Step -1
        If IsCommentRow(ws, R) Then GoTo NextRow
        Dim cmd As String
        cmd = PureCommandAt(ws, R)
        If IsExecutableCommand(cmd) Then
            HasFlowBreakBefore = (cmd = "EXIT" Or cmd = "GOTO" Or cmd = "RETURN")
            Exit Function
        End If
NextRow:
    Next R
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

    Dim targetText As String
    targetText = ParamText(targetCell.value)

    With targetCell.Font
        .Underline = xlUnderlineStyleNone
        If cmd = "CALL" Or cmd = "GOTO" Then
            If Len(targetText) > 0 Then
                .color = RGB(5, 99, 193)
                .Underline = xlUnderlineStyleSingle
            End If
        End If
    End With
End Sub

Private Sub ApplyLabelRoleDisplay(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal cmd As String, _
                                  ByVal program As clsScriptProgram)
    Dim roleCell As Range
    Set roleCell = ws.Cells(rowNum, 2)

    If cmd <> "LABEL" Then
        modScriptAssist.ClearLabelRoleNote roleCell
        Exit Sub
    End If

    If program Is Nothing Then Exit Sub

    Dim roleName As String
    roleName = program.LabelRoleAtSheetRow(rowNum)
    If Len(roleName) = 0 Then
        modScriptAssist.ClearLabelRoleNote roleCell
        Exit Sub
    End If

    Dim warning As String
    warning = program.LabelRoleWarningAtSheetRow(rowNum)

    modScriptAssist.ApplyLabelRoleNote roleCell, roleName, warning

    roleCell.Font.Italic = (roleName = "DeadLabel")
    If roleName = "MixedLabel" Then
        roleCell.Font.color = RGB(156, 101, 0)
    ElseIf roleName = "DeadLabel" Then
        roleCell.Font.color = RGB(128, 128, 128)
    End If
End Sub

Private Function PureCommandAt(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    PureCommandAt = modCmdRegistry.NormalizeCommand(ws.Cells(rowNum, 2).value)
End Function

Private Function IsExecutableCommand(ByVal cmd As String) As Boolean
    IsExecutableCommand = (Len(cmd) > 0)
End Function

Private Function IsCommentRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim rawCmd As String
    rawCmd = Trim$(CStr(ws.Cells(rowNum, 2).value))
    IsCommentRow = (Len(rawCmd) > 0 And Left$(rawCmd, 1) = "#")
End Function

' [파생] 블록 시작 판정 → BlockRole = STARTER
Private Function IsBlockStarter(cmd As String) As Boolean
    Dim def As CmdDef: Set def = modCmdRegistry.GetCmdDef(cmd)
    If Not def Is Nothing Then IsBlockStarter = (def.BlockRole = STARTER)
End Function

' [파생] 블록 종료 판정 → BlockRole = ENDER
Private Function IsBlockEnder(cmd As String) As Boolean
    Dim def As CmdDef: Set def = modCmdRegistry.GetCmdDef(cmd)
    If Not def Is Nothing Then IsBlockEnder = (def.BlockRole = ENDER)
End Function



