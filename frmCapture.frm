VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCapture 
   Caption         =   "UserForm1"
   ClientHeight    =   2768
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4448
   OleObjectBlob   =   "frmCapture.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Const CAPTURE_MODE_COORDS As String = "좌표/색상"
Private Const CAPTURE_MODE_PIXEL_IF As String = "픽셀 IF 조건"

Public Sub Prepare()
    Me.Caption = "매크로 헬퍼"
    chkContinuous.Caption = "자동 입력 모드 (Auto-Fill)"
    chkContinuous.value = False
    PrepareSnippetMode
    Me.Show vbModeless
End Sub

Public Function IsAutoMode() As Boolean
    IsAutoMode = CBool(chkContinuous.value)
End Function

Public Function IsPixelIfMode() As Boolean
    IsPixelIfMode = (SelectedSnippetMode() = CAPTURE_MODE_PIXEL_IF)
End Function

Public Function ModeCaption() As String
    ModeCaption = SelectedSnippetMode()
End Function

Public Sub RenderScanning(ByVal x As Long, ByVal y As Long, ByVal colorValue As Long, ByVal hexColor As String)
    Me.BackColor = &H8000000F
    Label1.Caption = "탐색 중..." & vbCrLf & _
                     "MODE: " & ModeCaption() & vbCrLf & _
                     "X: " & x & vbCrLf & _
                     "Y: " & y & vbCrLf & _
                     "CLR: " & hexColor
    Label2.BackColor = colorValue
End Sub

Public Sub RenderAutoFilled(ByVal x As Long, ByVal y As Long, ByVal hexColor As String)
    Label1.Caption = "▶ 시트 입력 완료! ◀" & vbCrLf & _
                     "MODE: " & ModeCaption() & vbCrLf & _
                     "X:" & x & " Y:" & y & " CLR: " & hexColor
    Me.BackColor = RGB(200, 255, 200)
End Sub

Public Sub RenderCopied()
    Label1.Caption = "▶ 클립보드 복사됨 ◀" & vbCrLf & _
                     "MODE: " & ModeCaption() & vbCrLf & _
                     "(붙여넣기 하세요)"
    Me.BackColor = RGB(255, 200, 200)
End Sub

Private Sub PrepareSnippetMode()
    On Error GoTo SafeExit
    With cboSnippetMode
        .Clear
        .AddItem CAPTURE_MODE_COORDS
        .AddItem CAPTURE_MODE_PIXEL_IF
        .ListIndex = 0
    End With
SafeExit:
End Sub

Private Function SelectedSnippetMode() As String
    On Error GoTo SafeExit
    SelectedSnippetMode = CStr(cboSnippetMode.value)
    If Len(Trim$(SelectedSnippetMode)) = 0 Then SelectedSnippetMode = CAPTURE_MODE_COORDS
    Exit Function
SafeExit:
    SelectedSnippetMode = CAPTURE_MODE_COORDS
End Function

