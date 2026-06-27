VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const MAX_VALIDATE_CELLS As Long = 500

Private Sub Workbook_Open()
    modScriptAssist.RegisterScriptAssistHotkeys
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modScriptAssist.UnregisterScriptAssistHotkeys
End Sub

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

    If Target.Cells.count <> 1 Or Target.Column < cmdCol Or Target.Column > cmdCol + 3 Then
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

    Dim assistChanges As Range
    Set assistChanges = Intersect(Target, ws.Range(ws.Columns(cmdCol), ws.Columns(cmdCol + 3)))
    If assistChanges Is Nothing Then Exit Sub
    Set assistChanges = Intersect(assistChanges, ws.UsedRange)
    If assistChanges Is Nothing Then Exit Sub
    If assistChanges.Cells.CountLarge > MAX_VALIDATE_CELLS Then Exit Sub

    Dim rowsToValidate As Object
    Set rowsToValidate = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In assistChanges.Cells
        If cell.Row > 1 Then rowsToValidate(CStr(cell.Row)) = True
    Next cell

    Dim key As Variant
    For Each key In rowsToValidate.Keys
        modScriptAssist.ValidateCommandCell ws.Cells(CLng(key), cmdCol)
    Next key
End Sub






