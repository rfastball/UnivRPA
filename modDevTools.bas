Attribute VB_Name = "modDevTools"
Option Explicit
Public Sub FormatRPAScript_Final()
    Dim ws As Worksheet: Set ws = ActiveSheet

    ' 0. 헤더 자동 생성
    Call EnsureHeaders(ws)

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    Dim R As Long, pureCmd As String, depth As Integer: depth = 0

    Application.ScreenUpdating = False

    ' [추가] 번호가 없는 부분(데이터 끝 이후) 서식 및 내용 초기화
    If lastRow < ws.Rows.count Then
        ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(ws.Rows.count, 6)).Clear
    End If

    ' 1. 기본 스타일 및 명확한 격자 초기화
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6))
        .Font.Name = "돋움체"
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
        ws.Cells(R, 1).Value = R - 1

        pureCmd = UCase(Trim(Replace(ws.Cells(R, 2).Value, ">", "")))
        If pureCmd = "" Then GoTo NextRow

            ' 2. 깊이 계산
            If IsBlockEnder(pureCmd) Then depth = IIf(depth > 0, depth - 1, 0)

                ' [수정] Else 문 들여쓰기 보정 (출력 시에만 깊이 -1 적용)
                Dim printDepth As Integer: printDepth = depth
                If pureCmd = "ELSE" Or pureCmd = "CASE" Then printDepth = IIf(depth > 0, depth - 1, 0)

                    ' 3. 명령어 텍스트 및 IDE 스타일 적용
                    ws.Cells(R, 2).Value = String(printDepth, ">") & (IIf(printDepth > 0, " ", "")) & pureCmd

                    ' [추가] 문법 및 파라미터 검사
                    Dim isValid As Boolean
                    isValid = ValidateLine(pureCmd, ws.Cells(R, 3).Value, ws.Cells(R, 4).Value, ws.Cells(R, 5).Value)

                    ' [통합 스타일 호출] 테두리 가이드와 배경색 처리
                    Call ApplyUnifiedStyle(ws.Cells(R, 2).Resize(1, 4), pureCmd, depth, isValid)

                    ' 4. 깊이 증가
                    If IsBlockStarter(pureCmd) Then depth = depth + 1

NextRow:
                    Next R

                    Application.ScreenUpdating = True
End Sub

' [신규] 헤더가 없으면 생성
Private Sub EnsureHeaders(ws As Worksheet)
    If ws.Cells(1, 2).Value = "" Then
        Dim headers As Variant
        headers = Array("ROW", "COMMAND", "PARAM1", "PARAM2", "PARAM3", "COMMENT")

        With ws.Range("A1:F1")
            .Value = headers
            .Font.Name = "돋움체"
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

' [신규] 명령어 및 파라미터 유효성 검사
Private Function ValidateLine(cmd As String, p1 As Variant, p2 As Variant, p3 As Variant) As Boolean
    ValidateLine = True

    ' 1. 명령어 존재 여부 확인
    Select Case True
     Case cmd = "IF", cmd = "ELSE", cmd = "END_IF", cmd = "WHILE", cmd = "WEND", _
        cmd = "FOR", cmd = "NEXT", cmd = "FOR_TABLE", cmd = "NEXT_TABLE", _
        cmd = "LOAD_TABLE", cmd = "GOTO", cmd = "CALL", cmd = "RETURN", cmd = "LABEL", _
        cmd = "SWITCH_PIXEL", cmd = "ASSERT", cmd = "KEY", cmd = "COMBO", _
        cmd = "FOCUS", cmd = "INC", cmd = "EXIT", cmd = "NOTIFY", _
        cmd = "SELECT", cmd = "CASE", cmd = "END_SELECT"
        ' 정상 명령어

     Case InStr(cmd, "IF_") = 1, InStr(cmd, "IFNOT_") = 1, _
        InStr(cmd, "WAIT") = 1, InStr(cmd, "MOUSE") > 0, _
        InStr(cmd, "TYPE") > 0, InStr(cmd, "MEM") > 0, _
        InStr(cmd, "CLIP") > 0
        ' 패턴 매칭 명령어

     Case Else
        ValidateLine = False
     Exit Function
    End Select
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

    Select Case True
        ' [Logic & Flow] - 연녹색 (줄 전체)
     Case IsBlockStarter(cmd), IsBlockEnder(cmd), cmd = "ELSE", cmd = "CASE", _
        cmd = "GOTO", cmd = "CALL", cmd = "RETURN", cmd = "LABEL", _
        InStr(cmd, "IF_") > 0, InStr(cmd, "IFNOT_") > 0, cmd = "SWITCH_PIXEL", cmd = "ASSERT"
        rng.Interior.color = RGB(226, 239, 218)

        ' [Wait & Vision] - 연황색 (B열 강조, C-E 연녹 배경)
     Case InStr(cmd, "WAIT") = 1
        cmdCell.Interior.color = RGB(255, 242, 204)
        If depth > 0 Then rng.offset(0, 1).Resize(1, 3).Interior.color = RGB(245, 255, 245)

            ' [Action & Input] - 연청색 (B열 강조, C-E 연녹 배경)
         Case InStr(cmd, "MOUSE") > 0, cmd = "KEY", cmd = "COMBO", _
            InStr(cmd, "TYPE") > 0, InStr(cmd, "MEM") > 0, cmd = "FOCUS", cmd = "INC", InStr(cmd, "CLIP") > 0
            cmdCell.Interior.color = RGB(221, 235, 247)
            If depth > 0 Then rng.offset(0, 1).Resize(1, 3).Interior.color = RGB(245, 255, 245)

                ' [System] - 연회색
             Case cmd = "EXIT", cmd = "NOTIFY"
                rng.Interior.color = RGB(242, 242, 242)

                ' 블록 내부인데 위 케이스에 해당 없는 경우 연녹색 배경 유지
             Case Else
                If depth > 0 Then rng.Interior.color = RGB(245, 255, 245)
                End Select
End Sub

' [수정] 블록 시작 명령어 정의
Private Function IsBlockStarter(cmd As String) As Boolean
    Select Case True
     Case InStr(cmd, "IF") = 1, InStr(cmd, "WHILE") = 1, _
        cmd = "FOR", cmd = "FOR_TABLE", cmd = "SELECT"
        IsBlockStarter = True
     Case Else: IsBlockStarter = False
    End Select
End Function

' [수정] 블록 종료 명령어 정의
Private Function IsBlockEnder(cmd As String) As Boolean
    Select Case True
     Case InStr(cmd, "END") = 1, cmd = "WEND", _
        cmd = "NEXT", cmd = "NEXT_TABLE"
        IsBlockEnder = True
     Case Else: IsBlockEnder = False
    End Select
End Function



