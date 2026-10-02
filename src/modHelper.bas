Attribute VB_Name = "modHelper"
' @desc 폼 컨트롤러(캡처 HUD/상태 폼) 구동 및 상태바/화면 배치 담당
' @depends modWinAPI
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' 캡처 자동입력 채움색 (RGB 220,255,220 연녹)
Private Const CLR_CAPTURE_FILL As Long = &HDCFFDC

'==============================================================================
' StatusBar gate (UI 쓰기의 단일 통로 - 상태 모듈(modContext)은 UI 무접촉)
'==============================================================================

' @desc 상태바에 RPA 접두 메시지를 표시함
'   DoEvents 없음 - 메시지 펌핑은 엔진 스로틀 블록/일시정지 루프가 소유
Public Sub SetStatusBar(ByVal msg As String)
    Application.StatusBar = "[RPA] " & msg
End Sub

' @desc 상태바를 엑셀 기본 상태로 복원함
Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

'==============================================================================
' Form controllers
'==============================================================================

Public Sub StartHUD()
    RunCaptureHud frmCapture
End Sub

' [폼 배치] 작업영역 좌상단으로 이동 + 최상위 고정 (폼 개념 조합은 여기가 소유,
'   modWinAPI는 선언/프리미티브만 제공)
Public Sub MoveFormToWorkAreaTopLeftTopMost(ByVal formView As Object)
    Dim workArea As modWinAPI.RECT
    Dim hdc As LongPtr
    Dim dpiX As Long
    Dim dpiY As Long
    Dim hWndForm As LongPtr

    modWinAPI.SystemParametersInfo modWinAPI.SPI_GETWORKAREA, 0, workArea, 0
    hdc = modWinAPI.GetDC(0)
    dpiX = modWinAPI.GetDeviceCaps(hdc, modWinAPI.LOGPIXELSX)
    dpiY = modWinAPI.GetDeviceCaps(hdc, modWinAPI.LOGPIXELSY)
    modWinAPI.ReleaseDC 0, hdc

    If dpiX > 0 Then formView.Left = (workArea.Left * 72 / dpiX)
    If dpiY > 0 Then formView.Top = (workArea.Top * 72 / dpiY)

    hWndForm = modWinAPI.FindWindow("ThunderDFrame", formView.Caption)
    If hWndForm <> 0 Then
        modWinAPI.SetWindowPos hWndForm, modWinAPI.HWND_TOPMOST, 0, 0, 0, 0, _
            modWinAPI.SWP_NOMOVE Or modWinAPI.SWP_NOSIZE Or modWinAPI.SWP_NOACTIVATE
    End If
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
        If ShouldStopCapture(view) Then Exit Do

        If modWinAPI.IsToggleKeyOn(modWinAPI.VK_CAPITAL) Then
            If Not isRecorded Then
                modWinAPI.GetCursorPos pt
                colorValue = modWinAPI.GetColorAt(pt.x, pt.y)
                hexColor = modWinAPI.FormatColorHex(colorValue)
                strInfo = CaptureSnippet(view, pt.x, pt.y, hexColor)

                If view.IsAutoMode Then
                    FillActiveCell view, pt.x, pt.y, hexColor
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
            hexColor = modWinAPI.FormatColorHex(colorValue)
            view.RenderScanning pt.x, pt.y, colorValue, hexColor
        End If

        If WaitCaptureInterval(view, 50) Then Exit Do
    Loop
End Sub

Public Sub ShowRpaStatus(ByVal view As frmRPAStatus, ByVal totalLines As Long)
    view.Prepare totalLines
End Sub

Public Sub UpdateRpaStatus(ByVal view As frmRPAStatus, ByVal curLine As Long, ByVal detailText As String, _
                           Optional ByVal rowIndex As Long = 0, Optional ByVal rowTotal As Long = 0)
    view.RenderProgress curLine, detailText, rowIndex, rowTotal
End Sub

Public Sub CloseRpaStatus(ByVal view As frmRPAStatus)
    Unload view
End Sub

Private Function CaptureSnippet(ByVal view As frmCapture, ByVal x As Long, ByVal y As Long, ByVal hexColor As String) As String
    If view.IsPixelIfMode Then
        CaptureSnippet = "{$PIXEL:" & x & "," & y & "}" & vbTab & "EQ" & vbTab & hexColor
    Else
        CaptureSnippet = x & vbTab & y & vbTab & hexColor
    End If
End Function

Private Sub FillActiveCell(ByVal view As frmCapture, ByVal x As Long, ByVal y As Long, ByVal hexColor As String)
    With ActiveCell
        If view.IsPixelIfMode Then
            .value = "{$PIXEL:" & x & "," & y & "}"
            .offset(0, 1).value = "EQ"
            .offset(0, 2).value = hexColor
        Else
            .value = x
            .offset(0, 1).value = y
            .offset(0, 2).value = hexColor
        End If
        .Interior.color = CLR_CAPTURE_FILL
        .offset(0, 1).Interior.color = CLR_CAPTURE_FILL
        .offset(0, 2).Interior.color = CLR_CAPTURE_FILL
        .offset(1, 0).Select
    End With
End Sub

Private Function IsFormVisible(ByVal view As Object) As Boolean
    On Error GoTo SafeExit
    IsFormVisible = view.Visible
SafeExit:
End Function

Private Function ShouldStopCapture(ByVal view As Object) As Boolean
    If Not IsFormVisible(view) Then
        ShouldStopCapture = True
    ElseIf modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then
        Unload view
        ShouldStopCapture = True
    End If
End Function

Private Function WaitCaptureInterval(ByVal view As Object, ByVal totalMs As Long) As Boolean
    Dim elapsedMs As Long
    Do While elapsedMs < totalMs
        If ShouldStopCapture(view) Then
            WaitCaptureInterval = True
            Exit Function
        End If
        modWinAPI.Sleep 10
        elapsedMs = elapsedMs + 10
        DoEvents
    Loop
End Function
