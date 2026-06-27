Attribute VB_Name = "modTableSource"
Option Explicit

Public Function SnapshotSource(ByVal pWorkbook As String, ByVal pSource As String, Optional ByVal pMode As String = "") As clsTableSnapshot
    On Error GoTo ErrHandler

    Dim openedHere As Boolean
    Dim wb As Workbook
    Set wb = ResolveWorkbook(Trim$(pWorkbook), openedHere)

    Dim visibleOnly As Boolean
    visibleOnly = ParseSnapshotMode(pMode)

    Dim sourceObject As Object
    Set sourceObject = ResolveTableSourceSpec(wb, Trim$(pSource))

    Dim tbl As clsTableSnapshot
    Set tbl = New clsTableSnapshot
    tbl.LoadFromSheet sourceObject, visibleOnly
    Set SnapshotSource = tbl

CleanUp:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

ErrHandler:
    Dim desc As String
    desc = Err.Description
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, "modTableSource.SnapshotSource", desc
End Function

Private Function ParseSnapshotMode(ByVal modeText As String) As Boolean
    Dim mode As String
    mode = UCase$(Trim$(modeText))

    Select Case mode
        Case "", "ALL_ROWS"
            ParseSnapshotMode = False
        Case "VISIBLE_ROWS"
            ParseSnapshotMode = True
        Case Else
            Err.Raise 5, "modTableSource.ParseSnapshotMode", "FOREACH_ROW P3는 비워두거나 ALL_ROWS/VISIBLE_ROWS만 사용할 수 있습니다: " & modeText
    End Select
End Function

Public Function OpenWorkbookHint(ByVal filterText As String) As String
    Dim matches As Collection
    Set matches = FindWorkbooksByPartialName(filterText)

    If matches.Count = 0 Then
        OpenWorkbookHint = "FOREACH_ROW: 열린 워크북 없음"
        Exit Function
    End If

    Dim i As Long
    Dim text As String
    For i = 1 To matches.Count
        If Len(text) > 0 Then text = text & " | "
        text = text & matches(i).Name
    Next i

    OpenWorkbookHint = "FOREACH_ROW P1 후보: " & text
End Function

Public Function ResolveWorkbook(ByVal workbookSpec As String, ByRef openedHere As Boolean) As Workbook
    If Len(workbookSpec) = 0 Then Err.Raise 5, "modTableSource.ResolveWorkbook", "워크북 이름 또는 경로가 필요합니다."

    openedHere = False

    If IsPathSpec(workbookSpec) Then
        Dim baseName As String
        baseName = Dir$(workbookSpec)
        If Len(baseName) = 0 Then baseName = FileNameFromPath(workbookSpec)

        Dim wb As Workbook
        For Each wb In Application.Workbooks
            If StrComp(wb.Name, baseName, vbTextCompare) = 0 Then
                Set ResolveWorkbook = wb
                Exit Function
            End If
        Next wb

        Set ResolveWorkbook = Application.Workbooks.Open(Filename:=workbookSpec, ReadOnly:=True, UpdateLinks:=False)
        openedHere = True
        Exit Function
    End If

    Dim matches As Collection
    Set matches = FindWorkbooksByPartialName(workbookSpec)

    Select Case matches.Count
        Case 1
            Set ResolveWorkbook = matches(1)
        Case 0
            Err.Raise 5, "modTableSource.ResolveWorkbook", "열린 워크북을 찾을 수 없습니다: " & workbookSpec
        Case Else
            Err.Raise 5, "modTableSource.ResolveWorkbook", "워크북 이름이 모호합니다: " & workbookSpec & vbCrLf & JoinWorkbookFullNames(matches)
    End Select
End Function

Public Function ResolveTableSourceSpec(ByVal wb As Workbook, ByVal sourceSpec As String) As Object
    If wb Is Nothing Then Err.Raise 5, "modTableSource.ResolveTableSourceSpec", "워크북이 필요합니다."
    If Len(sourceSpec) = 0 Then Err.Raise 5, "modTableSource.ResolveTableSourceSpec", "시트 또는 시트!주소가 필요합니다."

    Dim bangPos As Long
    bangPos = InStrRev(sourceSpec, "!")

    If bangPos = 0 Then
        Set ResolveTableSourceSpec = wb.Worksheets(UnquoteSheetName(sourceSpec))
        Exit Function
    End If

    Dim sheetName As String
    Dim addressText As String
    sheetName = UnquoteSheetName(Left$(sourceSpec, bangPos - 1))
    addressText = Mid$(sourceSpec, bangPos + 1)
    If Len(Trim$(addressText)) = 0 Then Err.Raise 5, "modTableSource.ResolveTableSourceSpec", "범위 주소가 비어 있습니다: " & sourceSpec

    Set ResolveTableSourceSpec = wb.Worksheets(sheetName).Range(addressText)
End Function

Public Function FindWorkbooksByPartialName(ByVal filterText As String) As Collection
    Dim result As New Collection
    Dim wb As Workbook
    Dim needle As String

    needle = Trim$(filterText)
    For Each wb In Application.Workbooks
        If Len(needle) = 0 Or InStr(1, wb.Name, needle, vbTextCompare) > 0 Then
            result.Add wb
        End If
    Next wb

    Set FindWorkbooksByPartialName = result
End Function

Private Function IsPathSpec(ByVal value As String) As Boolean
    IsPathSpec = (InStr(1, value, "\", vbTextCompare) > 0 _
               Or InStr(1, value, "/", vbTextCompare) > 0 _
               Or InStr(1, value, ":", vbTextCompare) > 0)
End Function

Private Function FileNameFromPath(ByVal value As String) As String
    Dim slashPos As Long
    Dim backslashPos As Long
    Dim pos As Long

    slashPos = InStrRev(value, "/")
    backslashPos = InStrRev(value, "\")
    pos = slashPos
    If backslashPos > pos Then pos = backslashPos

    If pos > 0 Then
        FileNameFromPath = Mid$(value, pos + 1)
    Else
        FileNameFromPath = value
    End If
End Function

Private Function UnquoteSheetName(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) >= 2 And Left$(value, 1) = "'" And Right$(value, 1) = "'" Then
        value = Mid$(value, 2, Len(value) - 2)
    End If
    UnquoteSheetName = Replace(value, "''", "'")
End Function

Private Function JoinWorkbookFullNames(ByVal matches As Collection) As String
    Dim i As Long
    For i = 1 To matches.Count
        If Len(JoinWorkbookFullNames) > 0 Then JoinWorkbookFullNames = JoinWorkbookFullNames & vbCrLf
        JoinWorkbookFullNames = JoinWorkbookFullNames & matches(i).FullName
    Next i
End Function
