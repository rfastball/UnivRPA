Attribute VB_Name = "modPractice"
' @desc 배포 사본에 기능별 연습 시트를 구성함
' @depends modDevTools, clsScriptProgram, modActions
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' @desc 배포 사본에 브라우저 실습과 기능별 연습 시트를 추가함
Public Sub BuildPracticeSheets()
    Dim anchor As Worksheet, firstLesson As Worksheet, tableLesson As Worksheet
    Set anchor = ThisWorkbook.Worksheets(1)

    Set firstLesson = AddLesson(anchor, "연습01 하나 쏙", Array( _
        "FOCUS|||UnivRPA Lab 01|H2에서 연습실을 열고 창을 나란히 두세요.", _
        "SLEEP|||1000|1초 뒤 모모의 바구니를 보세요.", _
        "CLICK|{간식X}|{간식Y}||간식 주기를 누르면 바나나 1개가 담깁니다."))
    AddLabControls firstLesson, "간식 주기 버튼"
    firstLesson.Range("H8").Value2 = "다시 시작을 누르고 CLICK 줄을 주석 처리해 보세요. 모모의 바구니가 비어 있으면 성공입니다."
    Dim entryLesson As Worksheet, waitLesson As Worksheet
    Set entryLesson = AddLesson(anchor, "연습02 파파야 둘", Array( _
        "FOCUS|||UnivRPA Lab 02|다시 시작하면 바나나 1개가 선택됩니다.", _
        "CLICK|{파파야X}|{파파야Y}||파파야 그림 버튼 중앙을 캡처하세요.", _
        "CLICK|{개수X}|{개수Y}||개수 칸을 클릭합니다.", _
        "COMBO|CTRL|A||기존 숫자 1을 모두 선택합니다.", _
        "TYPE|30||2|개수를 2로 바꿉니다.", _
        "KEY|ENTER|||모모의 바구니에 파파야 2개가 담깁니다."))
    AddLabControls entryLesson, "파파야 버튼과 개수 칸"
    Set waitLesson = AddLesson(anchor, "연습03 코코 오는 중", Array( _
        "FOCUS|||UnivRPA Lab 03|빠른 걸음으로 실행한 뒤 느린 걸음으로 바꿔 보세요.", _
        "CLICK|{부르기X}|{부르기Y}||코코 부르기를 누른 순간 출발합니다.", _
        "SLEEP|||1000|느린 걸음에서 아직 오는 중!이 뜨면 이 줄을 주석 처리하세요.", _
        "# WAIT_APPEAR|{초록불X}|{초록불Y}|408000|초록불의 중앙을 캡처하고 이 줄의 주석을 해제하세요.", _
        "CLICK|{간식X}|{간식Y}||코코의 바구니에 바나나 1개가 담겼는지 확인하세요."))
    AddLabControls waitLesson, "코코 부르기 버튼과 간식 주기 버튼"
    waitLesson.Range("H8").Value2 = "코코를 손으로 부르세요. 도착한 뒤 초록불 중앙을 좌표/색상 모드로 C5에 캡처하세요."
    waitLesson.Range("H9").Value2 = "느린 걸음에서 B4와 B5의 주석을 각각 전환하고 다시 시작한 뒤 실행하세요."

    AddLesson anchor, "연습04 시스템", Array( _
        "SLEEP|||500|Ctrl+Shift+W로 실행하세요. 0.5초 후 다음 줄로 이동합니다.", _
        "NOTIFY|2||첫 실행 성공!|알림 문구를 바꾸고 다시 실행해 보세요.", _
        "EXIT||||여기서 실행을 마칩니다.")

    AddLesson anchor, "연습05 변수", Array( _
        "SET|{재고}||10|재고를 10개로 시작합니다.", _
        "INC|{재고}||5|입고 5개를 더합니다.", _
        "DEC|{재고}||2|출고 2개를 뺍니다.", _
        "NOTIFY|2||남은 재고 {재고}개|예상 결과는 13개입니다. 입고량을 바꿔 보세요.")

    AddLesson anchor, "연습06 흐름 제어", Array( _
        "SET|{합계}||0|누적할 값을 0으로 초기화합니다.", _
        "REPEAT|3|||반복 횟수를 4로 바꿔 결과를 비교해 보세요.", _
        "INC|{합계}||{$LOOP}|현재 회차를 더합니다. 3회면 1+2+3입니다.", _
        "END_REPEAT||||반복을 마치면 합계를 검사합니다.", _
        "IF|{합계}|>=|10|합계가 10 이상인지 확인합니다.", _
        "NOTIFY|2||목표 달성 ({합계})|4회 반복하면 이 알림이 표시됩니다.", _
        "ELSE||||합계가 10 미만인 경우입니다.", _
        "NOTIFY|2||목표 미달 ({합계})|기본 3회에서는 목표 미달 (6)이 표시됩니다.", _
        "END_IF||||조건 분기를 마칩니다.")

    AddLesson anchor, "연습07 마우스", Array( _
        "MOVE|80|160|700|포인터가 첫 좌표에 도착해 0.7초간 머뭅니다.", _
        "MOVE|320|160|700|포인터가 오른쪽으로 이동합니다. X값을 바꿔 보세요.", _
        "NOTIFY|2||포인터 이동 완료|Ctrl+Shift+E로 연 캡처 HUD에서 다른 좌표를 확인할 수 있습니다.")

    AddLesson anchor, "연습08 화면 색상", Array( _
        "SET|{x}||80|검사할 화면 X좌표입니다. 고정된 배경을 고르세요.", _
        "SET|{y}||160|검사할 화면 Y좌표입니다.", _
        "SET|{기준색}||{$PIXEL:{x},{y}}|실행 시점의 픽셀 색상을 읽어 저장합니다.", _
        "SLEEP|||500|같은 위치를 다시 읽기 전 0.5초 기다립니다.", _
        "ASSERT|{x}|{y}|{기준색}|색이 달라지면 이 줄에서 멈춥니다.", _
        "NOTIFY|2||확인한 색상 {기준색}|색이 같으면 여섯 자리 색상값이 표시됩니다.")

    Set tableLesson = AddLesson(anchor, "연습09 표 순회", Array( _
        "SET|{총수량}||0|오른쪽 H:I열에 있는 주문 세 건의 수량을 더합니다.", _
        "FOREACH_ROW|UnivRPA|연습09 표 순회!H1:I4||파일명을 바꿨거나 여러 버전을 열었다면 P1에 파일명을 입력하세요.", _
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

Private Sub AddLabControls(ByVal ws As Worksheet, ByVal targetName As String)
    Dim button As Button
    Set button = ws.Buttons.Add(ws.Range("H2").Left, ws.Range("H2").Top, 110, 22)
    button.Caption = "연습실 열기"
    button.OnAction = "OpenPracticeLab"
    ws.Range("H4").Value2 = "창을 나란히 둔 뒤 창 위치와 확대율을 유지하세요. 대상: " & targetName
    ws.Range("H5").Value2 = "CLICK 행의 C셀 선택 > Ctrl+Shift+E > 좌표만 (X/Y) > 자동 입력 모드를 켭니다."
    ws.Range("H6").Value2 = "대상에 마우스를 올리고 CapsLock을 켭니다. 다음 캡처 전에는 끄고, Esc로 HUD를 닫습니다."
    ws.Range("H7").Value2 = "스크립트 시트에서 Ctrl+Shift+W로 실행합니다. Ctrl+Shift+/로 선택한 명령의 주석을 전환합니다."
End Sub

' @desc 현재 연습 시트에 맞는 로컬 브라우저 실습 창을 엶
Public Sub OpenPracticeLab()
    Dim lesson As String, page As String, edge As String, url As String, folder As Variant
    On Error GoTo Failed
    lesson = Mid$(ActiveSheet.Name, 3, 2)
    Select Case lesson
        Case "01", "02", "03"
        Case Else: Exit Sub
    End Select
    page = ThisWorkbook.Path & Application.PathSeparator & "연습실.html"
    If Len(Dir$(page)) = 0 Then
        MsgBox "배포 ZIP을 모두 풀고 연습실.html을 통합 문서와 같은 폴더에 두세요.", vbInformation
        Exit Sub
    End If
    If modActions.Action_FocusWindow("UnivRPA Lab " & lesson) Then Exit Sub
    url = "file:///" & Replace(page, "\", "/")
    url = Replace(Replace(Replace(url, "%", "%25"), "#", "%23"), " ", "%20") & "#" & lesson
    For Each folder In Array(Environ$("ProgramFiles(x86)"), Environ$("ProgramFiles"), Environ$("LocalAppData"))
        edge = folder & "\Microsoft\Edge\Application\msedge.exe"
        If Len(Dir$(edge)) > 0 Then
            Shell """" & edge & """ --app=""" & url & """", vbNormalFocus
            Exit Sub
        End If
    Next
    ThisWorkbook.FollowHyperlink Address:=page, SubAddress:=lesson, NewWindow:=True
    Exit Sub
Failed:
    MsgBox "연습실.html을 브라우저에서 열고 " & lesson & "강을 선택하세요." & vbCrLf & Err.Description, vbInformation
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
