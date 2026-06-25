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

Public Sub Prepare()
    Me.Caption = "매크로 헬퍼"
    chkContinuous.Caption = "자동 입력 모드 (Auto-Fill)"
    chkContinuous.value = False
    Me.Show vbModeless
End Sub

Public Function IsAutoMode() As Boolean
    IsAutoMode = CBool(chkContinuous.value)
End Function

Public Sub RenderScanning(ByVal x As Long, ByVal y As Long, ByVal colorValue As Long, ByVal hexColor As String)
    Me.BackColor = &H8000000F
    Label1.Caption = "탐색 중..." & vbCrLf & _
                     "X: " & x & vbCrLf & _
                     "Y: " & y & vbCrLf & _
                     "CLR: " & hexColor
    Label2.BackColor = colorValue
End Sub

Public Sub RenderAutoFilled(ByVal x As Long, ByVal y As Long, ByVal hexColor As String)
    Label1.Caption = "▶ 시트 입력 완료! ◀" & vbCrLf & _
                     "X:" & x & " Y:" & y & " CLR: " & hexColor
    Me.BackColor = RGB(200, 255, 200)
End Sub

Public Sub RenderCopied()
    Label1.Caption = "▶ 클립보드 복사됨 ◀" & vbCrLf & _
                     "(붙여넣기 하세요)"
    Me.BackColor = RGB(255, 200, 200)
End Sub


Private Sub cboSnippetMode_Change()

End Sub

