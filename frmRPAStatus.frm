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

' [API] 화면 해상도 및 작업 영역 감지
#If VBA7 Then
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const SPI_GETWORKAREA As Long = 48
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10

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

' [신규] 외부에서 호출 가능한 위치 이동 메서드
Public Sub MoveToBottomRight()
    #If VBA7 Then
        Dim hdc As LongPtr
        Dim hWndForm As LongPtr
    #Else
        Dim hdc As Long
        Dim hWndForm As Long
    #End If
    Dim R As RECT
    Dim dpiX As Long, dpiY As Long
    
    ' 1. 작업 영역(WorkArea) 및 DPI 가져오기
    SystemParametersInfo SPI_GETWORKAREA, 0, R, 0
    hdc = GetDC(0)
    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc

    ' 2. 좌표 계산 (폼이 Show된 후 호출해야 정확함)
    ' [수정] 좌상단으로 이동 (여백 최소화)
    Me.Left = (R.Left * 72 / dpiX)
    Me.Top = (R.Top * 72 / dpiY)

    ' 3. [New] 항상 위에 표시 (TopMost)
    hWndForm = FindWindow("ThunderDFrame", Me.Caption)
    If hWndForm <> 0 Then
        SetWindowPos hWndForm, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    End If
End Sub

Public Sub InitStatus(total As Long)
    m_TotalLines = total
    lblProgress.Caption = "Ready (Total: " & total & ")"
End Sub

Public Sub UpdateInfo(curLine As Long, detailText As String)
    ' 진행률 표시
    lblProgress.Caption = "Line " & curLine & "/" & m_TotalLines
    ' 실행중인 열 정보 (명령어 + 파라미터)
    lblCmd.Caption = detailText
    Me.Repaint
End Sub
