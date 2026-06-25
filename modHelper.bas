Attribute VB_Name = "modHelper"
Option Explicit

'==============================================================================
' Form controllers
'==============================================================================

Public Sub StartHUD()
    RunCaptureHud frmCapture
End Sub

Public Sub RunCaptureHud(ByVal view As frmCapture)
    Dim pt As POINTAPI
    Dim colorValue As Long
    Dim hexColor As String
    Dim strInfo As String
    Dim isRecorded As Boolean

    view.Prepare
    isRecorded = False

    Do While IsFormVisible(view)
        DoEvents
        If Not IsFormVisible(view) Then Exit Do

        If modWinAPI.IsToggleKeyOn(modWinAPI.VK_CAPITAL) Then
            If Not isRecorded Then
                modWinAPI.GetCursorPos pt
                colorValue = modWinAPI.GetColorAt(pt.x, pt.y)
                hexColor = Right("000000" & hex(colorValue), 6)
                strInfo = pt.x & vbTab & pt.y & vbTab & hexColor

                If view.IsAutoMode Then
                    FillActiveCell pt.x, pt.y, hexColor
                    view.RenderAutoFilled pt.x, pt.y, hexColor
                Else
                    modWinAPI.SetClipboardText strInfo
                    view.RenderCopied
                End If

                isRecorded = True
            End If
        Else
            isRecorded = False
            modWinAPI.GetCursorPos pt
            colorValue = modWinAPI.GetColorAt(pt.x, pt.y)
            hexColor = Right("000000" & hex(colorValue), 6)
            view.RenderScanning pt.x, pt.y, colorValue, hexColor
        End If

        modWinAPI.Sleep 50
    Loop
End Sub

Public Sub ShowRpaStatus(ByVal view As frmRPAStatus, ByVal totalLines As Long)
    view.Prepare totalLines
End Sub

Public Sub UpdateRpaStatus(ByVal view As frmRPAStatus, ByVal curLine As Long, ByVal detailText As String)
    view.RenderProgress curLine, detailText
End Sub

Public Sub CloseRpaStatus(ByVal view As frmRPAStatus)
    Unload view
End Sub

Private Sub FillActiveCell(ByVal x As Long, ByVal y As Long, ByVal hexColor As String)
    With ActiveCell
        .value = x
        .offset(0, 1).value = y
        .offset(0, 2).value = hexColor
        .Interior.color = RGB(220, 255, 220)
        .offset(0, 1).Interior.color = RGB(220, 255, 220)
        .offset(1, 0).Select
    End With
End Sub

Private Function IsFormVisible(ByVal view As Object) As Boolean
    On Error GoTo SafeExit
    IsFormVisible = view.Visible
SafeExit:
End Function
