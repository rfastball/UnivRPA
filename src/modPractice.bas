Attribute VB_Name = "modPractice"
' @desc 배포 사본에 기능별 연습 시트를 구성함
' @depends modDevTools, clsScriptProgram
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' @desc 배포 사본 앞에 실행 가능한 연습 시트 여섯 장을 추가함
Public Sub BuildPracticeSheets()
    Dim anchor As Worksheet, firstLesson As Worksheet, tableLesson As Worksheet
    Set anchor = ThisWorkbook.Worksheets(1)

    Set firstLesson = AddLesson(anchor, "연습01 시스템", Array( _
        "SLEEP|||500|Ctrl+Shift+W로 실행하세요. 0.5초 후 다음 줄로 이동합니다.", _
        "NOTIFY|2||첫 실행 성공!|알림 문구를 바꾸고 다시 실행해 보세요.", _
        "EXIT||||여기서 실행을 마칩니다."))

    AddLesson anchor, "연습02 변수", Array( _
        "SET|{재고}||10|재고를 10개로 시작합니다.", _
        "INC|{재고}||5|입고 5개를 더합니다.", _
        "DEC|{재고}||2|출고 2개를 뺍니다.", _
        "NOTIFY|2||남은 재고 {재고}개|예상 결과는 13개입니다. 입고량을 바꿔 보세요.")

    AddLesson anchor, "연습03 흐름 제어", Array( _
        "SET|{합계}||0|누적할 값을 0으로 초기화합니다.", _
        "REPEAT|3|||반복 횟수를 4로 바꿔 결과를 비교해 보세요.", _
        "INC|{합계}||{$LOOP}|현재 회차를 더합니다. 3회면 1+2+3입니다.", _
        "END_REPEAT||||반복을 마치면 합계를 검사합니다.", _
        "IF|{합계}|>=|10|합계가 10 이상인지 확인합니다.", _
        "NOTIFY|2||목표 달성 ({합계})|4회 반복하면 이 알림이 표시됩니다.", _
        "ELSE||||합계가 10 미만인 경우입니다.", _
        "NOTIFY|2||목표 미달 ({합계})|기본 3회에서는 목표 미달 (6)이 표시됩니다.", _
        "END_IF||||조건 분기를 마칩니다.")

    AddLesson anchor, "연습04 마우스", Array( _
        "MOVE|80|160|700|포인터가 첫 좌표에 도착해 0.7초간 머뭅니다.", _
        "MOVE|320|160|700|포인터가 오른쪽으로 이동합니다. X값을 바꿔 보세요.", _
        "NOTIFY|2||포인터 이동 완료|Ctrl+Shift+E로 연 캡처 HUD에서 다른 좌표를 확인할 수 있습니다.")

    AddLesson anchor, "연습05 화면 색상", Array( _
        "SET|{x}||80|검사할 화면 X좌표입니다. 고정된 배경을 고르세요.", _
        "SET|{y}||160|검사할 화면 Y좌표입니다.", _
        "SET|{기준색}||{$PIXEL:{x},{y}}|실행 시점의 픽셀 색상을 읽어 저장합니다.", _
        "SLEEP|||500|같은 위치를 다시 읽기 전 0.5초 기다립니다.", _
        "ASSERT|{x}|{y}|{기준색}|색이 달라지면 이 줄에서 멈춥니다.", _
        "NOTIFY|2||확인한 색상 {기준색}|색이 같으면 여섯 자리 색상값이 표시됩니다.")

    Set tableLesson = AddLesson(anchor, "연습06 표 순회", Array( _
        "SET|{총수량}||0|오른쪽 H:I열에 있는 주문 세 건의 수량을 더합니다.", _
        "FOREACH_ROW|UnivRPA|연습06 표 순회!H1:I4||파일명을 바꿨거나 여러 버전을 열었다면 P1에 파일명을 입력하세요.", _
        "INC|{총수량}||{$ROW.수량}|표의 수량 열을 머리글 이름으로 읽습니다.", _
        "END_FOREACH||||다음 주문으로 이동합니다.", _
        "NOTIFY|2||총 주문 수량 {총수량}개|예상 결과는 9개입니다. I열의 수량을 바꿔 보세요."))
    With tableLesson
        .Range("H1:I1").Value = Array("품목", "수량")
        .Range("H2:I2").Value = Array("노트", 2)
        .Range("H3:I3").Value = Array("펜", 4)
        .Range("H4:I4").Value = Array("파일", 3)
        .Range("H1:I1").Font.Bold = True
        .Columns("H:I").AutoFit
    End With
    firstLesson.Activate
End Sub

Private Function AddLesson(ByVal anchor As Worksheet, ByVal title As String, ByVal lines As Variant) As Worksheet
    Dim ws As Worksheet, i As Long, program As clsScriptProgram
    Set ws = ThisWorkbook.Worksheets.Add(Before:=anchor)
    ws.Name = title
    modDevTools.EnsureHeaders ws
    ws.Range("B2:F" & UBound(lines) + 2).NumberFormat = "@"
    For i = 0 To UBound(lines)
        ws.Cells(i + 2, 2).Resize(1, 5).Value = Split(lines(i), "|")
    Next i
    modDevTools.FormatRPAScript ws
    ws.Columns("B:F").AutoFit
    Set program = modDevTools.ProgramSnapshotFromSheet(ws)
    If Not program.ValidateStructure() Or Not program.ValidateLabels() Then _
        Err.Raise 5, "modPractice", title & vbCrLf & program.ValidationErrors()
    Set AddLesson = ws
End Function
