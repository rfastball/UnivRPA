Attribute VB_Name = "modHelper"
'==============================================================================
' [UI 설계 가이드] UserForm1 구성 명세서
'==============================================================================
' ※ 이 코드를 실행하기 위해 VBE에서 [삽입] -> [사용자 정의 폼]을 생성하고
'    아래 속성(Properties)에 맞춰 컨트롤을 배치하십시오.
'
' 1. UserForm (폼 자체)
'    - (이름)       : UserForm1
'    - Caption      : 매크로 헬퍼
'    - ShowModal    : False  ★중요★ (True면 루프가 돌지 않아 먹통이 됨)
'    - Height / Width : 적당히 작게 (예: 150 / 200) - 화면 구석에 띄우기 위함
'
' 2. Label (정보 표시용 텍스트)
'    - (이름)       : Label1
'    - Caption      : (비워둠)
'    - TextAlign    :
'    - Font         : 맑은 고딕, 굵게 (가독성 확보)
'    - 위치         : 폼 상단에 큼직하게 배치
'
' 3. Label (색상 미리보기 박스)
'    - (이름)       : Label2
'    - Caption      : (비워둠)
'    - BorderStyle  : 1 - fmBorderStyleSingle (테두리가 있어야 흰색 배경에서도 보임)
'    - Width/Height : 정사각형 권장 (예: 30 x 30)
'    - 위치         : Label1 우하단
'
' 4. CheckBox (자동 입력 모드 스위치)
'    - (이름)       : chkContinuous
'    - Caption      : 자동 입력 모드
'    - Value        : False
'    - 위치         : 폼 좌하단
'==============================================================================

'==============================================================================
' API 선언부
'==============================================================================
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Type POINTAPI
    x As Long
    y As Long
End Type

Const VK_CAPITAL = &H14

'==============================================================================
' [최종 수정] HUD 모드 (자동 입력 vs 수동 복사 선택)
'==============================================================================
Sub StartHUD_Final()
    Dim pt As POINTAPI
    Dim hdc As LongPtr
    Dim pColor As Long
    Dim hexColor As String
    Dim strInfo As String
    
    Dim isRecorded As Boolean ' 이미 기록했는지 확인하는 플래그
    Dim capsOn As Boolean     ' Caps Lock 상태
    Dim isAutoMode As Boolean ' 자동 기록 모드(체크박스) 여부
    
    ' 유저폼 초기화
    UserForm1.chkContinuous.Caption = "자동 입력 모드 (Auto-Fill)"
    UserForm1.chkContinuous.value = False ' 기본은 수동
    UserForm1.Show
    UserForm1.Caption = "매크로 헬퍼"
    
    isRecorded = False
    
    ' 유저폼 실행 중 루프
    Do While UserForm1.Visible
        DoEvents
        
        ' 1. 상태 및 좌표 확인
        capsOn = (GetKeyState(VK_CAPITAL) And 1) = 1
        isAutoMode = UserForm1.chkContinuous.value
        
        ' 2. Caps Lock이 켜져 있는 경우 (캡처 시도)
        If capsOn Then
            
            ' 아직 기록하지 않은 상태라면 (키를 누른 직후 1회만 실행)
            If isRecorded = False Then
                
                ' 좌표값 확정
                GetCursorPos pt
                hdc = GetDC(0)
                pColor = GetPixel(hdc, pt.x, pt.y)
                ReleaseDC 0, hdc
                hexColor = Right("000000" & hex(pColor), 6)
                strInfo = pt.x & vbTab & pt.y & vbTab & hexColor
                
                ' === [모드 분기] ===
                If isAutoMode Then
                    ' [A] 자동 입력 모드: 엑셀 셀에 값 꽂아넣기
                    With ActiveCell
                        .value = pt.x
                        .offset(0, 1).value = pt.y
                        .offset(0, 2).value = hexColor
                        
                        ' 시각적 피드백 (배경색 살짝 변경)
                        .Interior.color = RGB(220, 255, 220)
                        .offset(0, 1).Interior.color = RGB(220, 255, 220)
                        
                        ' 다음 입력을 위해 아랫줄 선택
                        .offset(1, 0).Select
                    End With
                    
                    UserForm1.Label1.Caption = "▶ 시트 입력 완료! ◀" & vbCrLf & _
                                               "X:" & pt.x & " Y:" & pt.y & " CLR: " & hexColor
                    UserForm1.BackColor = RGB(200, 255, 200) ' 녹색 피드백
                    
                Else
                    ' [B] 수동 모드: 클립보드 복사만
                    Dim DataObj As Object
                    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
                    DataObj.SetText strInfo
                    DataObj.PutInClipboard
                    
                    UserForm1.Label1.Caption = "▶ 클립보드 복사됨 ◀" & vbCrLf & _
                                               "(붙여넣기 하세요)"
                    UserForm1.BackColor = RGB(255, 200, 200) ' 붉은색 피드백
                End If
                
                ' 기록 완료 플래그 세움 (Caps Lock 끌 때까지 재실행 방지)
                isRecorded = True
            End If
            
        Else
            ' 3. Caps Lock이 꺼져 있는 경우 (탐색 모드)
            
            ' 기록 플래그 초기화 (다시 찍을 준비)
            isRecorded = False
            
            ' 실시간 좌표 갱신
            GetCursorPos pt
            hdc = GetDC(0)
            pColor = GetPixel(hdc, pt.x, pt.y)
            ReleaseDC 0, hdc
            hexColor = Right("000000" & hex(pColor), 6)
            
            UserForm1.BackColor = &H8000000F ' 기본색 복귀
            UserForm1.Label1.Caption = "탐색 중..." & vbCrLf & _
                                       "X: " & pt.x & vbCrLf & _
                                       "Y: " & pt.y & vbCrLf & _
                                       "CLR: " & hexColor
            UserForm1.Label2.BackColor = pColor
        End If
        
        Sleep 50
    Loop
End Sub





