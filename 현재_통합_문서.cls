VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Target Is Nothing Then
        Application.StatusBar = False
        Exit Sub
    End If

    If Not TypeOf Sh Is Worksheet Then
        Application.StatusBar = False
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Sh

    Dim cmdCol As Long
    cmdCol = modScriptAssist.ResolveCmdCol(ws)
    If cmdCol = 0 Then
        Application.StatusBar = False
        Exit Sub
    End If

    If Target.Cells.count <> 1 Or Target.Column <> cmdCol Then
        Application.StatusBar = False
        Exit Sub
    End If

    modScriptAssist.ShowParamHint Target
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    Dim cmdCol As Long
    cmdCol = modScriptAssist.ResolveCmdCol(ws)
    If cmdCol = 0 Then Exit Sub

    Dim commandChanges As Range
    Set commandChanges = Intersect(Target, ws.Columns(cmdCol))
    If commandChanges Is Nothing Then Exit Sub

    Dim cell As Range
    For Each cell In commandChanges.Cells
        If cell.Row > 1 Then modScriptAssist.ValidateCommandCell cell
    Next cell
End Sub
