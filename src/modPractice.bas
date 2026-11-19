Attribute VB_Name = "modPractice"
' @desc 원숭이 군도 배송 시연 통합 문서 구성
' @depends modDevTools, modEngine, modActions, modWinAPI, clsScriptProgram
Option Explicit

' @desc 배포 사본에 시연 시작, 배송표, 작동 스크립트를 구성함
Public Sub BuildPracticeSheets()
    Dim anchor As Worksheet, front As Worksheet, orders As Worksheet, ws As Worksheet
    Dim oldSheet As Worksheet, row As Long, program As clsScriptProgram
    Set anchor = ThisWorkbook.Worksheets(1)
    Set front = ThisWorkbook.Worksheets.Add(Before:=anchor)
    front.Name = "시연 시작"
    Set orders = ThisWorkbook.Worksheets.Add(Before:=anchor)
    orders.Name = "배송표"
    Set ws = ThisWorkbook.Worksheets.Add(Before:=anchor)
    ws.Name = "작동 스크립트"

    front.Range("B2").Value2 = "원숭이 군도 특급배송"
    front.Range("B2").Font.Size = 22
    front.Range("B2").Font.Bold = True
    front.Range("B4").Value2 = "화면을 열고 배송표 다섯 행을 운항합니다. 종료까지 약 1분 30초가 걸립니다."
    front.Range("B6").Value2 = "코코가 짐을 받은 뒤 선착장의 신호와 섬 수령함을 지켜보세요."
    front.Range("B11").Value2 = "첫 운항이 끝나면 군도 창에서 파파야 +1 보충을 누르고 이어서 배송하세요."
    front.Range("B13").Value2 = "Pause: 일시정지 / 재개   Esc: 중단   같은 군도 창을 유지하면 수령 기록을 이어 씁니다."
    front.Range("B15").Value2 = "창을 직접 열 때는 같은 폴더의 군도.html을 사용하세요. 화면 응답 판독에는 텍스트 클립보드를 사용합니다."
    front.Columns("B").ColumnWidth = 100
    Dim btn As Button
    Set btn = front.Buttons.Add(front.Range("B8").Left, front.Range("B8").Top, 120, 28)
    btn.Caption = "처음부터 배송"
    btn.OnAction = "StartIslandShipping"
    Set btn = front.Buttons.Add(front.Range("B8").Left + 132, front.Range("B8").Top, 120, 28)
    btn.Caption = "이어서 배송"
    btn.OnAction = "ResumeIslandShipping"

    orders.Range("A1:D1").Value = Array("번호", "원숭이", "과일", "수량")
    orders.Range("A2:D2").Value = Array("A01", "모모", "바나나", 2)
    orders.Range("A3:D3").Value = Array("A02", "코코", "파파야", 3)
    orders.Range("A4:D4").Value = Array("A03", "루루", "바나나", 4)
    orders.Range("A5:D5").Value = Array("A04", "모모", "파파야", 2)
    orders.Range("A6:D6").Value = Array("A02", "코코", "파파야", 3)
    orders.Range("A1:D1").Font.Bold = True
    orders.Columns("A:D").ColumnWidth = 15

    modDevTools.EnsureHeaders ws
    ws.Columns("B:F").NumberFormat = "@"
    row = 2
    AddStep ws, row, "FOCUS|||UnivRPA Islands|군도 창에서만 입력합니다."
    AddStep ws, row, "SET|{명령}||이어배송|시연 시작 버튼이 새배송 또는 이어배송을 선택합니다."
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "FOREACH_ROW|UnivRPA.xlsm|배송표||"
    AddStep ws, row, "SET|{명령}||목록 {$ROW.번호} {$ROW.원숭이} {$ROW.과일} {$ROW.수량}|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "END_FOREACH||||"
    AddStep ws, row, "SLEEP|||2500|관객이 오늘의 배송표를 먼저 봅니다."
    AddStep ws, row, "FOREACH_ROW|UnivRPA.xlsm|배송표||"
    AddStep ws, row, "SET|{명령}||조회 {$RIDX}|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "SLEEP|||2300|현재 행과 목적지를 보여줍니다."
    AddStep ws, row, "IF|{응답}|EQ|OPEN|"
    AddStep ws, row, "SET|{명령}||잔량|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "SET|{남은짐}||{응답}|"
    AddStep ws, row, "SET|{명령}||재고|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "IF|{응답}|<|{남은짐}|"
    AddStep ws, row, "SET|{명령}||보류|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "SLEEP|||3500|"
    AddStep ws, row, "ELSE||||"
    AddStep ws, row, "WHILE|{남은짐}|>|0|"
    AddStep ws, row, "SET|{적재}||{남은짐}|"
    AddStep ws, row, "IF|{적재}|>|3|"
    AddStep ws, row, "SET|{적재}||3|뗏목 한 척의 적재 한도입니다."
    AddStep ws, row, "END_IF||||"
    AddStep ws, row, "CALL|||한 척 운항|"
    AddStep ws, row, "DEC|{남은짐}||{적재}|"
    AddStep ws, row, "END_WHILE||||"
    AddStep ws, row, "SET|{명령}||확정|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "SLEEP|||2200|"
    AddStep ws, row, "END_IF||||"
    AddStep ws, row, "END_IF||||"
    AddStep ws, row, "END_FOREACH||||"
    AddStep ws, row, "SET|{명령}||마감|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "EXIT||||"
    AddStep ws, row, "LABEL|||한 척 운항|"
    AddStep ws, row, "SET|{명령}||출항 {적재}|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "SET|{상자번호}||{응답}|"
    AddStep ws, row, "SLEEP|||4600|3초 항해 후 섬의 도착 장면을 잠시 유지합니다."
    AddStep ws, row, "SET|{명령}||신호|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "IF|{응답}|!=|{상자번호}|"
    AddStep ws, row, "SLEEP|||2500|도착 신호가 없으면 재발송 전에 수령함을 엽니다."
    AddStep ws, row, "SET|{명령}||수령함 {상자번호}|"
    AddStep ws, row, "CALL|||선착장 조작|"
    AddStep ws, row, "SLEEP|||3500|수령증 번호가 보이는 동안 멈춥니다."
    AddStep ws, row, "IF|{응답}|!=|{상자번호}|"
    AddStep ws, row, "NOTIFY|4||수령증 불일치. 배송을 중단합니다.|"
    AddStep ws, row, "EXIT||||"
    AddStep ws, row, "END_IF||||"
    AddStep ws, row, "END_IF||||"
    AddStep ws, row, "RETURN||||"
    AddStep ws, row, "LABEL|||선착장 조작|"
    AddStep ws, row, "KEY|F2|||"
    AddStep ws, row, "TYPE|10||{명령}|"
    AddStep ws, row, "KEY|ENTER||120|"
    AddStep ws, row, "SET|{$CLIP}||응답없음|이전 응답을 새 수령증으로 오인하지 않도록 합니다."
    AddStep ws, row, "COMBO|CTRL|C|150|"
    AddStep ws, row, "SET|{응답}||{$CLIP}|"
    AddStep ws, row, "IF|{응답}|EQ|응답없음|"
    AddStep ws, row, "COMBO|CTRL|C|250|마지막 응답만 한 번 더 읽습니다."
    AddStep ws, row, "SET|{응답}||{$CLIP}|"
    AddStep ws, row, "END_IF||||"
    AddStep ws, row, "IF|{응답}|EQ|응답없음|"
    AddStep ws, row, "NOTIFY|3||군도 창의 응답을 복사하지 못했습니다.|"
    AddStep ws, row, "EXIT||||"
    AddStep ws, row, "ELSE_IF|{응답}|EQ|ERROR|"
    AddStep ws, row, "NOTIFY|3||선착장에서 거절했습니다. 군도 화면의 안내를 확인하세요.|"
    AddStep ws, row, "EXIT||||"
    AddStep ws, row, "END_IF||||"
    AddStep ws, row, "RETURN||||"
    modDevTools.FormatRPAScript ws
    ws.Columns("B:F").AutoFit
    Set program = modDevTools.ProgramSnapshotFromSheet(ws)
    If Not program.ValidateStructure() Or Not program.ValidateLabels() Then Err.Raise 5, "modPractice", program.ValidationErrors()
    front.Activate
    For Each oldSheet In ThisWorkbook.Worksheets
        If oldSheet.Index >= anchor.Index Then oldSheet.Visible = xlSheetVeryHidden
    Next
End Sub

Private Sub AddStep(ByVal ws As Worksheet, ByRef row As Long, ByVal line As String)
    ws.Cells(row, 2).Resize(1, 5).Value = Split(line, "|")
    row = row + 1
End Sub

' @desc 재고와 수령 기록을 초기화하고 운항함
Public Sub StartIslandShipping()
    RunIslandShipping "새배송"
End Sub

' @desc 기존 수령 기록과 현재 재고로 운항함
Public Sub ResumeIslandShipping()
    RunIslandShipping "이어배송"
End Sub

Private Sub RunIslandShipping(ByVal mode As String)
    Dim ws As Worksheet, page As String, url As String, edge As String, folder As Variant, attempt As Long
    Static busy As Boolean
    If busy Then Exit Sub
    busy = True
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("작동 스크립트")
    If Not modActions.Action_FocusWindow("UnivRPA Islands") Then
        page = ThisWorkbook.Path & Application.PathSeparator & "군도.html"
        If Len(Dir$(page)) = 0 Then Err.Raise 53, , "통합 문서 옆에 군도.html이 필요합니다. ZIP 전체를 풀어 주세요."
        url = Application.WorksheetFunction.EncodeURL(Replace(page, "\", "/"))
        url = "file:///" & Replace(Replace(url, "%2F", "/", 1, -1, vbTextCompare), "%3A", ":", 1, -1, vbTextCompare)
        For Each folder In Array(Environ$("ProgramFiles(x86)"), Environ$("ProgramFiles"), Environ$("LocalAppData"))
            edge = folder & "\Microsoft\Edge\Application\msedge.exe"
            If Len(Dir$(edge)) > 0 Then Exit For
        Next
        If Len(Dir$(edge)) > 0 Then
            Shell """" & edge & """ --app=""" & url & """", vbNormalFocus
        Else
            ThisWorkbook.FollowHyperlink Address:=page, NewWindow:=True
        End If
    End If
    For attempt = 1 To 60
        If modActions.Action_FocusWindow("UnivRPA Islands") Then Exit For
        If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then GoTo Finished
        modWinAPI.Sleep 150
        DoEvents
    Next
    If attempt > 60 Then Err.Raise 5, , "군도.html을 브라우저에서 열고 다시 실행하세요."
    ws.Range("E3").Value2 = mode
    modEngine.RunScriptSheet ws, False, False
Finished:
    If Not ws Is Nothing Then ws.Range("E3").Value2 = "이어배송"
    busy = False
    Exit Sub
Failed:
    MsgBox Err.Description, vbInformation
    Resume Finished
End Sub
