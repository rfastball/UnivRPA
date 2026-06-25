Attribute VB_Name = "modDevTools"
' @desc RPA 스크립트 시트 서식·검증을 레지스트리 파생으로 적용 담당
' @depends modCmdRegistry
Option Explicit

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

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    Dim R As Long, pureCmd As String, depth As Integer: depth = 0
    Dim baseDepths() As Long
    Dim subDepths() As Long
    Dim subStartRows As Object
    Dim subEndRows As Object
    Dim invalidRows As Object
    Call AnalyzeSubroutineFlow(ws, lastRow, baseDepths, subDepths, subStartRows, subEndRows, invalidRows)

    ' [추가] 번호가 없는 부분(데이터 끝 이후) 서식 및 내용 초기화
    Dim clearLastRow As Long
    Dim clearCol As Long
    For clearCol = 1 To 6
        clearLastRow = Application.Max(clearLastRow, ws.Cells(ws.Rows.count, clearCol).End(xlUp).Row)
    Next clearCol
    If clearLastRow > lastRow Then
        ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(clearLastRow, 6)).Clear
    End If

    If lastRow < 2 Then GoTo CleanExit

    ' 1. 기본 스타일 및 명확한 격자 초기화
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6))
        .Font.name = "돋움체"
        .Font.Size = 10
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

    ' 2. [추가] 행 높이 일괄 고정 (IDE의 일정한 간격 구현)
    ' 18~20 정도가 Consolas 10pt에 가장 보기 좋은 높이입니다.
    ws.Range("2:" & lastRow).RowHeight = 19

    For R = 2 To lastRow
        ' 0. 줄번호 자동 생성
        ws.Cells(R, 1).value = R - 1

        pureCmd = UCase(Trim(Replace(ws.Cells(R, 2).value, ">", "")))
        If pureCmd = "" Then GoTo NextRow

            ' 2. 깊이 계산
            If IsBlockEnder(pureCmd) Then depth = IIf(depth > 0, depth - 1, 0)

                ' [수정] Else 문 들여쓰기 보정 (출력 시에만 깊이 -1 적용)
                Dim printDepth As Integer: printDepth = depth
                If pureCmd = "ELSE" Or pureCmd = "CASE" Then printDepth = IIf(depth > 0, depth - 1, 0)
                printDepth = printDepth + subDepths(R)

                    ' 3. 명령어 텍스트 및 IDE 스타일 적용
                    ws.Cells(R, 2).value = String(printDepth, ">") & (IIf(printDepth > 0, " ", "")) & pureCmd

                    ' [추가] 문법 및 파라미터 검사
                    Dim isValid As Boolean
                    isValid = ValidateLine(pureCmd, ws.Cells(R, 3).value, ws.Cells(R, 4).value, ws.Cells(R, 5).value)
                    If invalidRows.exists(CStr(R)) Then isValid = False

                    ' [통합 스타일 호출] 테두리 가이드와 배경색 처리
                    Call ApplyUnifiedStyle(ws.Cells(R, 2).Resize(1, 4), pureCmd, depth + subDepths(R), isValid)
                    Call ApplySubroutineBoundaryStyle(ws.Cells(R, 2).Resize(1, 4), _
                                                       subStartRows.exists(CStr(R)), _
                                                       subEndRows.exists(CStr(R)), _
                                                       Not isValid)

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
    If ws.Cells(1, 2).value = "" Then
        Dim headers As Variant
        headers = Array("ROW", "COMMAND", "PARAM1", "PARAM2", "PARAM3", "COMMENT")

        With ws.Range("A1:F1")
            .value = headers
            .Font.name = "돋움체"
            .Font.Size = 10
            .Font.Bold = True
            .Interior.color = RGB(240, 240, 240)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        ws.Columns("A").ColumnWidth = 5
        ws.Columns("B").ColumnWidth = 35
        ws.Columns("C:E").ColumnWidth = 15
        ws.Columns("F").ColumnWidth = 30
    End If
End Sub

' [파생] 명령어 및 파라미터 유효성 검사 → GetCmdDef 게이트로 환원
Public Function ValidateLine(cmd As String, p1 As Variant, p2 As Variant, p3 As Variant) As Boolean
    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(cmd)

    ' 1. 명령어 존재 여부 = 레지스트리 등록 여부
    If def Is Nothing Then
        ValidateLine = False
        Exit Function
    End If

    ValidateLine = True

    ' 2. 조건 연산자: 맨 '=' 금지 (== 강제)
    '    연산자 슬롯(P2=RAW)을 가진 명령어(IF/WHILE)에 한해 검사 ? 데이터로 일반화
    If Not def.p2 Is Nothing Then
        If def.p2.IsRaw Then
            If Trim(CStr(p2)) = "=" Then ValidateLine = False
        End If
    End If
End Function

Public Function ValidateSheetFlowAt(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    ValidateSheetFlowAt = True
    If ws Is Nothing Then Exit Function
    If rowNum < 2 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    If rowNum > lastRow Then Exit Function

    Dim baseDepths() As Long
    Dim subDepths() As Long
    Dim subStartRows As Object
    Dim subEndRows As Object
    Dim invalidRows As Object
    Call AnalyzeSubroutineFlow(ws, lastRow, baseDepths, subDepths, subStartRows, subEndRows, invalidRows)

    ValidateSheetFlowAt = Not invalidRows.exists(CStr(rowNum))
End Function

' 통합 스타일 엔진: 하이라이트 강제 초기화 및 스타일 재적용
Public Sub ApplyUnifiedStyle(rng As Range, cmd As String, depth As Integer, Optional isValid As Boolean = True)
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
    If Left(cmd, 1) = "#" Then
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

Private Sub AnalyzeSubroutineFlow(ByVal ws As Worksheet, ByVal lastRow As Long, _
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

    Dim labelRows As Object
    Dim callTargets As Object
    Set labelRows = BuildLabelRowMap(ws, lastRow)
    Set callTargets = BuildCallTargetMap(ws, lastRow)

    Call MarkMissingJumpTargets(ws, lastRow, labelRows, invalidRows)

    Dim inSubroutineRows As Object
    Set inSubroutineRows = CreateObject("Scripting.Dictionary")
    inSubroutineRows.CompareMode = vbTextCompare

    Dim R As Long
    For R = 2 To lastRow
        If PureCommandAt(ws, R) = "LABEL" Then
            Dim labelName As String
            labelName = LabelNameAt(ws, R)

            If callTargets.exists(labelName) Then
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
    Next R

    For R = 2 To lastRow
        If PureCommandAt(ws, R) = "RETURN" Then
            If Not inSubroutineRows.exists(CStr(R)) Then invalidRows(CStr(R)) = True
        End If
    Next R
End Sub

Private Sub ComputeBaseDepths(ByVal ws As Worksheet, ByVal lastRow As Long, ByRef baseDepths() As Long)
    Dim depth As Long
    Dim R As Long

    For R = 2 To lastRow
        Dim cmd As String
        cmd = PureCommandAt(ws, R)

        If IsBlockEnder(cmd) Then depth = IIf(depth > 0, depth - 1, 0)

        baseDepths(R) = depth
        If cmd = "ELSE" Or cmd = "CASE" Then baseDepths(R) = IIf(depth > 0, depth - 1, 0)

        If IsBlockStarter(cmd) Then depth = depth + 1
    Next R
End Sub

Private Function BuildLabelRowMap(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim R As Long
    For R = 2 To lastRow
        If PureCommandAt(ws, R) = "LABEL" Then
            Dim labelName As String
            labelName = LabelNameAt(ws, R)
            If Len(labelName) > 0 Then
                Dim rows As Collection
                If Not result.exists(labelName) Then
                    Set rows = New Collection
                    result.Add labelName, rows
                Else
                    Set rows = result(labelName)
                End If
                rows.Add R
            End If
        End If
    Next R

    Set BuildLabelRowMap = result
End Function

Private Function BuildCallTargetMap(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim R As Long
    For R = 2 To lastRow
        If PureCommandAt(ws, R) = "CALL" Then
            Dim target As String
            target = JumpTargetAt(ws, R)
            If IsStaticLabelTarget(target) Then result(target) = True
        End If
    Next R

    Set BuildCallTargetMap = result
End Function

Private Sub MarkMissingJumpTargets(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                   ByVal labelRows As Object, ByVal invalidRows As Object)
    Dim R As Long
    For R = 2 To lastRow
        Dim cmd As String
        cmd = PureCommandAt(ws, R)

        If cmd = "CALL" Or cmd = "GOTO" Then
            Dim target As String
            target = JumpTargetAt(ws, R)
            If IsStaticLabelTarget(target) Then
                If Not labelRows.exists(target) Then invalidRows(CStr(R)) = True
            End If
        End If
    Next R
End Sub

Private Function FindSubroutineReturn(ByVal ws As Worksheet, ByVal startRow As Long, _
                                      ByVal lastRow As Long, ByVal labelDepth As Long, _
                                      ByRef baseDepths() As Long) As Long
    Dim R As Long
    For R = startRow + 1 To lastRow
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
        Dim cmd As String
        cmd = PureCommandAt(ws, R)
        If IsExecutableCommand(cmd) Then
            HasFlowBreakBefore = (cmd = "EXIT" Or cmd = "GOTO" Or cmd = "RETURN")
            Exit Function
        End If
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

Private Function PureCommandAt(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    PureCommandAt = UCase$(Trim$(Replace(Trim$(CStr(ws.Cells(rowNum, 2).value)), ">", "")))
End Function

Private Function LabelNameAt(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    LabelNameAt = UCase$(Trim$(CStr(ws.Cells(rowNum, 5).value)))
End Function

Private Function JumpTargetAt(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    JumpTargetAt = UCase$(Trim$(CStr(ws.Cells(rowNum, 5).value)))
End Function

Private Function IsExecutableCommand(ByVal cmd As String) As Boolean
    IsExecutableCommand = (Len(cmd) > 0 And Left$(cmd, 1) <> "#")
End Function

Private Function IsStaticLabelTarget(ByVal target As String) As Boolean
    IsStaticLabelTarget = (Len(target) > 0 And Not IsNumeric(target) And InStr(1, target, "{", vbTextCompare) = 0)
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



