VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRPAStatus 
   Caption         =   "UserForm2"
   ClientHeight    =   2848
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4288
   OleObjectBlob   =   "frmRPAStatus.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmRPAStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

' 동적 컨트롤 변수
Private lblProgress As Object
Private lblCmd As Object
Private m_TotalLines As Long

Private Sub UserForm_Initialize()
    ' [디자인] 슬림 & 모던 스타일
    Me.Caption = "RPA Running"
    Me.Width = 240  ' 너비 축소
    Me.Height = 80  ' 높이 축소 (슬림)
    Me.BackColor = &HFFFFFF ' 흰색 배경 (Clean Look)
    
    ' 1. 진행 상태 (좌측 상단)
    Set lblProgress = Me.Controls.Add("Forms.Label.1", "lblProgress")
    With lblProgress
        .Top = 6: .Left = 8: .Width = 150: .Height = 14
        .Font.name = "Malgun Gothic": .Font.Size = 9: .Font.Bold = True
        .ForeColor = &H404040 ' 진한 회색
        .Caption = "Ready"
    End With
    
    ' 3. 현재 실행 정보 (하단)
    Set lblCmd = Me.Controls.Add("Forms.Label.1", "lblCmd")
    With lblCmd
        .Top = 24: .Left = 8: .Width = 220: .Height = 30
        .Font.name = "Consolas": .Font.Size = 8
        .ForeColor = &H808080 ' 연한 회색
        .WordWrap = False ' 한 줄로 표시 (잘리면 잘리는대로)
        .Caption = ""
    End With
End Sub

Public Sub Prepare(ByVal total As Long)
    m_TotalLines = total
    Me.Show vbModeless
    modWinAPI.MoveFormToWorkAreaTopLeftTopMost Me
    lblProgress.Caption = "Ready (Total: " & total & ")"
End Sub

Public Sub RenderProgress(ByVal curLine As Long, ByVal detailText As String)
    ' 진행률 표시
    lblProgress.Caption = "Line " & curLine & "/" & m_TotalLines
    ' 실행중인 열 정보 (명령어 + 파라미터)
    lblCmd.Caption = detailText
    Me.Repaint
End Sub
