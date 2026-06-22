Attribute VB_Name = "modEngine"

Option Explicit

'==============================================================================
' [Layer 1] modEngine: RPA Core Logic & Interpreter
'==============================================================================
' 역할: 엑셀 시트의 명령어를 해석하고, 하위 모듈(API, Actions, Context)을 조립하여 실행
' 원칙:
' 1. 직접적인 API 호출 금지 (modWinAPI 사용)
' 2. 직접적인 변수 조작 금지 (modContext 사용)
' 3. 직접적인 동작 수행 금지 (modActions 사용)
'==============================================================================
' [수정] 배열을 모듈 전역 변수로 선언하여 모든 함수에서 접근 가능하게 함
Private m_arrMain As Variant

' [상수 정의] 매직 넘버 제거
Private Const CMD_NEXT_LINE As Long = 0
Private Const CMD_ERROR As Long = -1
Private Const CMD_NOT_HANDLED As Long = -2

'==============================================================================
' 1. Main Entry Point (RunRPA)
'==============================================================================
Sub RunRPA()
Attribute RunRPA.VB_ProcData.VB_Invoke_Func = "W\n14"
    Dim tb1 As ListObject
    Dim currentLine As Long
    Dim nextLine As Long
    Dim currentLoop As Long
    ' [wsSource] 실행 시점의 시트를 박제할 변수 추가
    Dim wsSource As Worksheet
    Dim rngTarget As Range
    Dim totalLines As Long

    ' [복구용 변수 추가]
    Dim pureCmd As String
    Dim depth As Long

    ' [Context] 1. 실행 환경 초기화
    modContext.InitContext
    currentLoop = 0

    '==========================================================
    ' [핵심: 스냅샷] 버튼을 누른 "바로 그 시트"를 고정합니다.
    ' 이후 모든 로직은 ActiveSheet가 아닌 wsSource를 바라봅니다.
    '==========================================================
    Set wsSource = ActiveSheet

    ' [데이터 로드] wsSource를 기준으로 영역을 잡습니다.
    On Error Resume Next
    ' 이제 실행 중에 포커스가 바뀌어도 wsSource는 원래 시트를 잡고 있습니다.
    Set rngTarget = wsSource.Range("A1").CurrentRegion
    On Error GoTo 0

        ' 유효성 검사: 데이터가 없거나 헤더만 있는 경우
        If rngTarget Is Nothing Or rngTarget.Rows.count < 2 Then
            MsgBox "실행할 데이터가 없습니다. (현재 시트: " & wsSource.Name & ")", vbExclamation
            modContext.ResetContext
         Exit Sub
        End If

        ' [수정] 헤더(1행) 제외: 실제 데이터는 2행부터 시작
        Set rngTarget = rngTarget.offset(1, 0).Resize(rngTarget.Rows.count - 1)

        ' [중요] 2차원 배열로 메모리에 완전히 적재 (메모리 맵 생성)
        ' 이후 루프 안에서는 Cells나 Range를 쓰지 말고 오직 m_arrMain만 써야 합니다.
        m_arrMain = rngTarget.Value
        totalLines = UBound(m_arrMain, 1)

        ' [New] 0. 라벨 프리 스캔 (Pre-Scan)
        ' 메모리에 올라온 m_arrMain 지도를 보고 미리 위치를 파악합니다.
        PreScanLabels m_arrMain, totalLines

        ' [UI] 상태 표시 폼 로드
        frmRPAStatus.Show vbModeless
        frmRPAStatus.MoveToBottomRight ' [Fix] 폼이 뜬 직후 위치 강제 이동
        frmRPAStatus.InitStatus totalLines

        ' [재시작 포인트]
RestartPoint:
        Do
            ' [Context] 실행 상태 확인 (Loop 시작 전)
            If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then GoTo ExitMacro

                currentLoop = currentLoop + 1
                modContext.LogStatus "▶ 매크로 실행 중... (" & currentLoop & " 회차)"

                ' 줄 번호 초기화
                currentLine = 1

                ' [내부 루프] 한 줄씩 실행
                Do While currentLine <= totalLines

                    ' 1. ESC 체크 (비상 탈출)
                    If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then GoTo ExitMacro

                        ' 2. 일시정지(~) 체크 및 대기 로직
                        If modWinAPI.IsKeyDown(modWinAPI.VK_PAUSE_RESUME) Then
                            ' [WinAPI] 메시지박스 타임아웃 사용 (3초 후 자동 닫힘 Or 사용자 확인)
                            modActions.Action_ShowNotification "일시정지됨. [Pause] 재시작, [ESC] 종료", 3
                            modWinAPI.Sleep 500

                            Do
                                ' 다시 Pause 키를 누르면 재시작(처음부터)
                                If modWinAPI.IsKeyDown(modWinAPI.VK_PAUSE_RESUME) Then
                                    modWinAPI.Sleep 200
                                    modContext.LogStatus "처음부터 다시 시작합니다."
                                    GoTo RestartPoint
                                    End If

                                    ' ESC 누르면 완전 종료
                                    If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then GoTo ExitMacro

                                        DoEvents
                                        modWinAPI.Sleep 50
                                    Loop
                                End If

                                '----------------------------------------------------------
                                ' [명령어 파싱 및 상태 업데이트]
                                '----------------------------------------------------------
                                ' 현재 실행할 행의 정보를 미리 파악 (B열 데이터 기준)
                                Dim fullCmdText As String
                                fullCmdText = m_arrMain(currentLine, 2)
                                pureCmd = UCase(Trim(Replace(fullCmdText, ">", "")))
                                depth = Len(fullCmdText) - Len(Replace(fullCmdText, ">", ""))

                                ' [UI] 표시할 상세 정보 조합 (명령어 + 파라미터)
                                Dim stepInfo As String
                                stepInfo = pureCmd & " " & m_arrMain(currentLine, 3) & " " & m_arrMain(currentLine, 4) & " " & m_arrMain(currentLine, 5)

                                ' [UI] 폼 업데이트
                                frmRPAStatus.UpdateInfo currentLine, currentLoop, stepInfo
                                DoEvents

                                '----------------------------------------------------------
                                ' 3. 액션 실행 및 에러 감시
                                '----------------------------------------------------------
                                On Error Resume Next ' 런타임 에러 발생 시 중단 방지
                                nextLine = ExecuteSmartAction(m_arrMain, currentLine)

                                ' 에러 체크 (VBA 에러 혹은 커스텀 에러 코드 -99 가정)
                                If Err.Number <> 0 Or nextLine = -99 Then
                                    ' 알림창 띄우기 (여기서 코드가 멈춤)
                                    MsgBox "?? [에러 발생] 라인: " & currentLine & "행" & vbCrLf & _
                                    "명령어: " & pureCmd & vbCrLf & _
                                    "내용: " & Err.Description, vbCritical, "JM_RPA Error Tracker"

                                    Err.Clear
                                    On Error GoTo 0
                                        GoTo ExitMacro ' 에러 발생 시 안전하게 종료
                                        End If
                                        On Error GoTo 0


                                            If nextLine = -1 Then GoTo ExitMacro

                                                ' 4. 액션 직후 체크
                                                If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then GoTo ExitMacro
                                                    If modWinAPI.IsKeyDown(modWinAPI.VK_PAUSE_RESUME) Then DoEvents

                                                        ' 5. 분기 처리 (Goto, Call, RETURN 등)
                                                        If nextLine > 0 Then
                                                            currentLine = nextLine
                                                            DoEvents
                                                        Else
                                                            currentLine = currentLine + 1
                                                        End If

                                                        DoEvents
                                                    Loop
                                                Loop

ExitMacro:
                                                ' [UI] 폼 종료
                                                Unload frmRPAStatus

                                                ' [Context] 종료 처리 (스택/메모리 해제, 상태바 복귀)
                                                modContext.ResetContext

                                                ' 완료 알림
                                                modActions.Action_ShowNotification "매크로가 종료되었습니다.", 1
End Sub

' [진짜 엔진이 호출하는 입구]
' 로그 기록 ON/OFF 용
Public Function ExecuteSmartAction(arr As Variant, R As Long) As Long
    On Error GoTo ErrHandler
    Dim res As Long

    ' 1. 원래 하던 대로 실행 (Core 호출)
    res = ExecuteSmartAction_Core(arr, R)

    Dim status As String
    If res = CMD_ERROR Then status = "FAIL" Else status = "SUCCESS"

        ' rP3 등의 변수를 가져오기 번거로우면 그냥 원본(arr(r, 5))을 남겨도 충분합니다.
        'Call modContext.WriteLog(r, arr(r, 2), arr(r, 5), status)

        ' 3. 결과값 그대로 반환
        ExecuteSmartAction = res

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function


Public Function ExecuteSmartAction_Core(arr As Variant, R As Long) As Long
    On Error GoTo ErrHandler
    Dim sType As String
    Dim sP1 As String, sP2 As String, sP3 As String
    Dim rP1 As String, rP2 As String, rP3 As String
    Dim p1 As Long, p2 As Long, p3 As String

    ' 1. 명령어(Type) 먼저 추출
    sType = UCase(Trim(CStr(arr(R, 2))))

    ' [New] 주석 처리: #으로 시작하면 파라미터 해석 없이 즉시 다음 줄로 이동
    If Left(sType, 1) = "#" Then
        ExecuteSmartAction_Core = CMD_NEXT_LINE ' 0 = 다음 줄 실행
     GoTo CleanUp
    End If

    ' [New: 들여쓰기 제거] ">" 기호를 모두 제거하고 다시 공백 제거
    ' 예: ">> TYPE" -> "TYPE"
    sType = Trim(Replace(sType, ">", ""))

    ' 2. 나머지 파라미터 추출 (주석이 아닐 때만 실행됨 - 성능 최적화)
    sP1 = CStr(arr(R, 3))
    sP2 = CStr(arr(R, 4))
    sP3 = CStr(arr(R, 5))

    ' 2. 파라미터 해석 (Centralized Resolution)
    '    "{n}" 형태면 메모리 값을 가져오고, 아니면 원본 유지
    rP1 = ResolveParam(sP1)
    rP2 = ResolveParam(sP2)
    rP3 = ResolveParam(sP3)

    ' 3. 타입 변환 (For Compatibility)
    '    대부분의 명령어는 해석된 값(rP)을 숫자로 변환하여 사용
    ' LONG 타입 오버플로우 방지 ON Error Resume Next
    On Error Resume Next
    p1 = val(rP1)
    p2 = val(rP2)
    On Error GoTo ErrHandler
    p3 = rP3     ' 문자열은 그대로 사용

    ' [Refactoring] 로직 분리: 흐름 제어 vs 액션 수행
    Dim result As Long
    
    ' 3. 흐름 제어 (Flow Control) 처리
    '    - IF, FOR, WHILE, GOTO 등 라인 점프가 주 목적
    result = ProcessFlowControl(sType, R, sP1, rP1, rP3, p3)
    If result <> CMD_NOT_HANDLED Then
        ExecuteSmartAction_Core = result
        GoTo CleanUp
    End If

    ' 4. 액션 (Action) 처리
    '    - 입력, 비전, 메모리 조작 등 동작 수행이 주 목적
    result = ProcessAction(sType, R, sP1, p1, p2, p3, rP1, rP2)
    If result <> CMD_NOT_HANDLED Then
        ExecuteSmartAction_Core = result
        GoTo CleanUp
    End If

    ' 5. 처리되지 않은 명령어
    MsgBox "알 수 없는 명령어입니다: " & sType, vbCritical
    ExecuteSmartAction_Core = CMD_ERROR

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Internal] 액션 명령어 처리 (입력, 비전, 메모리 등)
Private Function ProcessAction(sType As String, R As Long, sP1 As String, p1 As Long, p2 As Long, p3 As String, Optional rP1 As String = "", Optional rP2 As String = "") As Long
    On Error GoTo ErrHandler
    
    ProcessAction = CMD_NEXT_LINE ' 기본값: 성공 시 다음 줄

    Select Case sType
        '--------------------------------------------------
        ' [Vision & Logic] 판단 및 대기
        '--------------------------------------------------
     Case "WAIT_PIXEL", "WAIT_VANISH"
        ' 좌표(p1, p2), 색상(p3) 모두 변수 사용 가능
        If Not modActions.Action_Wait(sType, p1, p2, p3) Then
            ProcessAction = CMD_ERROR
         GoTo CleanUp
        End If

     Case "ASSERT"
        If Not modActions.Action_Assert(p1, p2, p3) Then
            ProcessAction = CMD_ERROR
         GoTo CleanUp
        End If
        
        '-------------------------------
        ' [Input] 단순 동작
        '--------------------------------------------------
     Case "MOUSE-L": modActions.Action_MouseClick p1, p2, False
     Case "MOUSE-R": modActions.Action_MouseClick p1, p2, True
     Case "KEY":     modActions.Action_PressKey rP1
     Case "COMBO":   modActions.Action_PressCombo rP1, rP2

     Case "TYPE"
        ' TYPE {1} -> 1번 메모리 내용 입력 가능
        modActions.Action_SmartType p1, p3

     Case "TYPE_CLIP":
        modActions.Action_SmartType p1, modWinAPI.GetClipboardText()

     Case "FOCUS"
        ' FOCUS {1} -> 1번 메모리에 저장된 창 제목으로 포커스 가능
        If Not modActions.Action_FocusWindow(p3) Then
            MsgBox "창을 찾을 수 없습니다: " & p3, vbCritical
            ProcessAction = CMD_ERROR
         GoTo CleanUp
        End If
        modActions.BetterWait 500

     Case "CLEAR_CLIP"
        modWinAPI.SetClipboardText ""

        '--------------------------------------------------
        ' [Memory]
        '--------------------------------------------------
     Case "WAIT"
        ' WAIT {1} -> 가변 대기 가능
        modActions.BetterWait val(p3)

     Case "EXIT"
        ProcessAction = CMD_ERROR
     GoTo CleanUp

        ' [중요] 메모리 쓰기 명령어들은 '주소'는 원본(sP1), '값'은 해석본(p3)을 사용
     Case "SET_MEM"
        If UCase(p3) = "{CLIP}" Then
            modContext.SetVar val(sP1), modWinAPI.GetClipboardText()
        Else
            ' p3는 이미 Resolve 되었으므로, {1}을 넣으면 값이 복사됨
            modContext.SetVar val(sP1), p3
        End If

     Case "TYPE_MEM"
        ' 기존 명령어 유지 (내부적으로 TYPE {n}과 동일해짐)
        ' 입력 대상 인덱스는 sP1 사용
        modActions.Action_SmartType 0, modContext.GetVar(val(sP1))

     Case "CLEAR_MEM": modContext.SetVar val(sP1), ""
     Case "INC": IncrementVariable p1: ProcessAction = CMD_NEXT_LINE

         Case "NOTIFY"
            modActions.Action_Notify val(p1), val(p2), p3
            
         Case Else
            ProcessAction = CMD_NOT_HANDLED
            GoTo CleanUp
    End Select

    ' [사후 대기] Input 계열만 p3에 대기시간이 있을 수 있음
    If IsNumeric(p3) And val(p3) > 0 Then
        Select Case sType
            Case "MOUSE-L", "MOUSE-R", "KEY", "COMBO", "FOCUS"
            modActions.BetterWait val(p3)
        End Select
    End If

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Internal] 흐름 제어 명령어 처리 (분기, 반복, 점프)
Private Function ProcessFlowControl(sType As String, R As Long, sP1 As String, rP1 As String, rP3 As String, p3 As String) As Long
    On Error GoTo ErrHandler
    Dim jumpTarget As Long
    
    ProcessFlowControl = CMD_NEXT_LINE

    Select Case sType
         Case "LABEL": ' Pass

         Case "GOTO"
            If IsNumeric(p3) Then
                ProcessFlowControl = val(p3)
            Else
                jumpTarget = modContext.GetLabel(p3)
                If jumpTarget > 0 Then ProcessFlowControl = jumpTarget Else ProcessFlowControl = CMD_ERROR
            End If

         Case "CALL"
            modContext.PushStack R + 1
            If IsNumeric(p3) Then
                ProcessFlowControl = val(p3)
            Else
                jumpTarget = modContext.GetLabel(p3)
                If jumpTarget > 0 Then ProcessFlowControl = jumpTarget Else ProcessFlowControl = CMD_ERROR
            End If

         Case "RETURN"
            Dim retLine As Long
            retLine = modContext.PopStack()
            If retLine > 0 Then ProcessFlowControl = retLine Else ProcessFlowControl = CMD_ERROR

         Case "IF_PIXEL", "IFNOT_PIXEL", "SWITCH_PIXEL"
            ' p1, p2는 좌표이므로 재해석 필요할 수 있으나 여기선 단순화
            ' 주의: ProcessFlowControl 인자에는 p1, p2가 없으므로 rP1 등을 파싱해야 함
            ' 편의상 여기서는 modActions 호출 시 val(ResolveParam...) 사용 권장되나
            ' 기존 로직 유지를 위해 rP1(좌표X), rP2(좌표Y)가 필요함.
            ' [Fix] 인자 부족으로 인해 여기서는 rP1, rP2를 다시 파싱하거나 인자를 추가해야 함.
            ' -> 간단히 해결하기 위해 rP1을 콤마로 나누거나, 상위에서 p1, p2를 넘겨받아야 함.
            ' -> 이번 리팩토링에서는 IF_PIXEL 등은 Action 성격이 강하므로 ProcessAction으로 넘기는 게 맞을 수도 있으나
            ' -> 점프를 수행하므로 FlowControl이 맞음.
            ' -> [Solution] 이 함수 내에서 필요한 값만 다시 파싱 (성능 영향 미미)
            Dim x As Long, y As Long
            x = val(rP1): y = val(ResolveParam(CStr(m_arrMain(R, 4)))) ' rP2
            jumpTarget = modActions.GetJumpTarget_Pixel(sType, x, y, p3)
            If jumpTarget > 0 Then ProcessFlowControl = jumpTarget

         Case "IF_CLIP", "IFNOT_CLIP"
            jumpTarget = modActions.GetJumpTarget_Text(sType, modWinAPI.GetClipboardText(), p3)
            If jumpTarget > 0 Then ProcessFlowControl = jumpTarget

         Case "IF_MEM", "IFNOT_MEM"
            jumpTarget = modActions.GetJumpTarget_Text(sType, modContext.GetVar(val(sP1)), p3)
            If jumpTarget > 0 Then ProcessFlowControl = jumpTarget

         Case "IF"
            If CheckIfSyntax(rP1, rP3) Then
                ProcessFlowControl = CMD_NEXT_LINE
            Else
                Dim targetLine As Long: targetLine = FindBlockPair(R, "IF", 1)
                If targetLine > 0 Then
                    Dim foundCmd As String
                    foundCmd = UCase(Trim(Replace(m_arrMain(targetLine, 2), ">", "")))
                    ProcessFlowControl = IIf(foundCmd = "ELSE", targetLine + 1, targetLine)
                Else
                    ProcessFlowControl = CMD_ERROR
                End If
            End If

         Case "ELSE"
            ProcessFlowControl = FindBlockPair(R, "IF", 1)

         Case "END_IF", "END_SELECT"
            ProcessFlowControl = CMD_NEXT_LINE

         Case "WHILE"
            If CheckIfSyntax(rP1, rP3) Then
                ProcessFlowControl = CMD_NEXT_LINE
            Else
                Dim endLoop As Long: endLoop = FindBlockPair(R, "WHILE", 1)
                If endLoop > 0 Then ProcessFlowControl = endLoop + 1 Else ProcessFlowControl = CMD_ERROR
            End If

         Case "WEND"
            Dim startLoop As Long: startLoop = FindBlockPair(R, "WHILE", -1)
            If startLoop > 0 Then ProcessFlowControl = startLoop Else ProcessFlowControl = CMD_ERROR

         Case "FOR"
            If modContext.GetLoopCount(R) = 0 Then modContext.SetLoopCount R, 1
            ProcessFlowControl = CMD_NEXT_LINE

         Case "NEXT"
            Dim forLine As Long: forLine = FindBlockPair(R, "FOR", -1)
            If forLine > 0 Then
                Dim currentCnt As Long: currentCnt = modContext.GetLoopCount(forLine)
                Dim maxCnt As Long: maxCnt = val(ResolveParam(CStr(m_arrMain(forLine, 3))))
                If currentCnt < maxCnt Then
                    modContext.SetLoopCount forLine, currentCnt + 1
                    ProcessFlowControl = forLine + 1
                Else
                    modContext.RemoveLoopCount forLine
                    ProcessFlowControl = CMD_NEXT_LINE
                End If
            Else
                ProcessFlowControl = CMD_ERROR
            End If

            '--------------------------------------------------
            ' [Structured Flow] Select / Case / End_Select
            '--------------------------------------------------
         Case "SELECT"
            Dim selVal As String, caseVal As String
            Dim scanR As Long, scanDepth As Long
            Dim scanCmd As String
            Dim fallbackLine As Long

            selVal = rP1 ' Select Value (e.g. {CLIP}, {PIXEL:100,100})
            scanDepth = 0
            fallbackLine = 0

            ' Scan forward For matching Case
            For scanR = R + 1 To UBound(m_arrMain, 1)
                scanCmd = UCase(Trim(Replace(m_arrMain(scanR, 2), ">", "")))

                If scanCmd = "SELECT" Then
                    scanDepth = scanDepth + 1
                ElseIf scanCmd = "END_SELECT" Then
                    If scanDepth > 0 Then
                        scanDepth = scanDepth - 1
                    Else
                        ' End of Select block reached
                        If fallbackLine > 0 Then
                            ProcessFlowControl = fallbackLine + 1
                        Else
                            ProcessFlowControl = scanR ' Jump To End_Select
                        End If
                     GoTo CleanUp
                    End If
                ElseIf scanCmd = "CASE" Then
                    If scanDepth = 0 Then
                        caseVal = ResolveParam(CStr(m_arrMain(scanR, 3)))

                        ' Check For Else (Case | Else Or Empty)
                        If UCase(caseVal) = "ELSE" Or caseVal = "" Then
                            fallbackLine = scanR
                        Else
                            If modActions.Action_CheckCondition(selVal, "==", caseVal) Then
                                ProcessFlowControl = scanR + 1
                             GoTo CleanUp
                            End If
                        End If
                    End If
                End If
            Next scanR
            ProcessFlowControl = CMD_ERROR

         Case "CASE"
            ' Case encountered during execution means fallthrough -> Jump To End_Select
            Dim endSelLine As Long
            endSelLine = FindBlockPair(R, "SELECT", 1)
            If endSelLine > 0 Then ProcessFlowControl = endSelLine Else ProcessFlowControl = CMD_ERROR

         Case Else
            ProcessFlowControl = CMD_NOT_HANDLED
    End Select

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function
' [Internal] 파라미터 해석기 (Interpolation 지원 버전)
Private Function ResolveParam(ByVal rawVal As String) As String
    On Error GoTo ErrHandler
    Dim strVal As String: strVal = Trim(rawVal)
    Dim result As String: result = strVal
    Dim parts() As String
    Dim i As Long

    ' [Case A] 파이프(|) 처리 (기존 로직 유지)
    If InStr(strVal, "|") > 0 Then
        parts = Split(strVal, "|")
        For i = LBound(parts) To UBound(parts)
            parts(i) = ResolveParam(parts(i))
        Next i
        ResolveParam = Join(parts, "|")
     GoTo CleanUp
    End If

    ' [Case B] 문장 내 변수 치환 (Interpolation)
    ' 1. 시스템 예약어 {CLIP} 치환
    If InStr(UCase(result), "{CLIP}") > 0 Then
        result = Replace(result, "{CLIP}", modWinAPI.GetClipboardText(), , , vbTextCompare)
    End If

    ' 2. 일반 메모리 변수 {0}~{99} 치환
    ' 문장 전체에서 {n} 패턴을 찾아 실제 값으로 갈아끼웁니다.
    For i = 0 To 99
        Dim token As String: token = "{" & i & "}"
        If InStr(result, token) > 0 Then
            result = Replace(result, token, modContext.GetVar(i))
        End If
    Next i

    ' 3. 특수 토큰 처리 (기존의 {PIXEL:x,y} 등 단일 토큰 로직)
    ' 픽셀 값은 보통 문장 중간보다는 "단일 값"으로 쓰이므로 기존 로직을 보호합니다.
    If Left(result, 1) = "{" And Right(result, 1) = "}" Then
        Dim inner As String: inner = Mid(result, 2, Len(result) - 2)
        If InStr(UCase(inner), "PIXEL:") = 1 Then
            result = ResolvePixelTokenInside(inner)
        End If
    End If

    ResolveParam = result

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Helper] PIXEL 토큰 내부 파싱 (예: "PIXEL:100,200")
Private Function ResolvePixelTokenInside(ByVal innerText As String) As String
    On Error GoTo ErrHandler
    Dim content As String, coords() As String
    Dim x As Long, y As Long

    ' "PIXEL:" 이후의 문자열 추출
    content = Mid(innerText, 7)
    coords = Split(content, ",")

    If UBound(coords) = 1 Then
        ' 좌표값에 변수가 섞여있을 수 있으므로 재귀적으로 한 번 더 해석
        x = val(ResolveParam(Trim(coords(0))))
        y = val(ResolveParam(Trim(coords(1))))

        ' ★ 핵심: 엔진은 modWinAPI의 래퍼만 호출합니다.
        ResolvePixelTokenInside = modWinAPI.GetColorHexAt(x, y)
    Else
        ' 형식이 맞지 않으면 원본 반환
        ResolvePixelTokenInside = "{" & innerText & "}"
    End If

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

Private Sub PreScanLabels(arr As Variant, total As Long)
    On Error GoTo ErrHandler
    Dim R As Long
    Dim sType As String
    Dim sLabelName As String

    For R = 1 To total
        sType = UCase(Trim(CStr(arr(R, 2))))

        ' 들여쓰기(>) 및 주석(#) 제거 후 판단
        If Left(sType, 1) = "#" Then GoTo ContinueLoop
            sType = Trim(Replace(sType, ">", ""))

            If sType = "LABEL" Then
                ' P3 (5번째 열)를 라벨 이름으로 사용
                sLabelName = CStr(arr(R, 5))
                modContext.AddLabel sLabelName, R
            End If

ContinueLoop:
        Next R

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub
' [Internal] 블록의 짝을 찾는 통합 탐색 함수
' StartLine: 탐색 시작 위치 (현재 줄)
' SearchType: 찾으려는 블록 종류 ("If", "While", "For")
' Direction: 1 (아래로 탐색 - 시작점에서 끝 찾기), -1 (위로 탐색 - 끝점에서 시작 찾기)
' Return: 찾은 줄 번호 (못 찾으면 -1)
Private Function FindBlockPair(ByVal startLine As Long, ByVal SearchType As String, Optional ByVal Direction As Long = 1) As Long
    On Error GoTo ErrHandler
    Dim R As Long
    Dim depth As Long
    Dim sCmd As String
    Dim endLine As Long

    ' 탐색 범위 설정
    If Direction = 1 Then
        endLine = UBound(m_arrMain, 1) ' 아래로 끝까지
    Else
        endLine = 1 ' 위로 첫 줄까지
    End If

    depth = 0

    ' Step Direction을 통해 r이 ++ 또는 -- 되도록 설정
    For R = startLine + Direction To endLine Step Direction
        ' 명령어 추출 및 주석 처리
        sCmd = CStr(m_arrMain(R, 2))
        
        If Trim(sCmd) = "" Or Left(Trim(sCmd), 1) = "#" Then GoTo NextRow
        
        sCmd = UCase(Trim(Replace(sCmd, ">", "")))

        Select Case UCase(SearchType)
            Case "IF"
                ' [If]는 역방향 탐색이 필요 없으나 구조 유지를 위해 포함
                If sCmd = "IF" Then depth = depth + 1
                If sCmd = "END_IF" Or (sCmd = "ELSE" And depth = 0) Then
                    If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                    If sCmd = "END_IF" Then depth = depth - 1
                End If

            Case "WHILE"
                If Direction = 1 Then ' 아래로 WEND 찾기
                    If sCmd = "WHILE" Then depth = depth + 1
                    If sCmd = "WEND" Then
                        If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                        depth = depth - 1
                    End If
                Else ' 위로 While 찾기
                    If sCmd = "WEND" Then depth = depth + 1
                    If sCmd = "WHILE" Then
                        If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                        depth = depth - 1
                    End If
                End If

            Case "FOR"
                If Direction = 1 Then ' 아래로 Next 찾기
                    If sCmd = "FOR" Then depth = depth + 1
                    If sCmd = "NEXT" Then
                        If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                        depth = depth - 1
                    End If
                Else ' 위로 For 찾기
                    If sCmd = "NEXT" Then depth = depth + 1
                    If sCmd = "FOR" Then
                        If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                        depth = depth - 1
                    End If
                End If

            Case "SELECT"
                ' [수정] IF문과 동일하게 정방향(End 찾기)만 지원 (루프 아님)
                If sCmd = "SELECT" Then depth = depth + 1
                If sCmd = "END_SELECT" Then
                    If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                    depth = depth - 1
                End If
        End Select

NextRow:
    Next R

    FindBlockPair = -1 ' 짝을 못 찾음

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Internal] If/While 문의 조건절(P1, P3)을 해석하여 True/False 반환
' v1: P1에서 해석된 값
' p3Raw: P3 셀의 원본 문자열 (예: "GT | 100" 또는 "Success")
' [modEngine] CheckIfSyntax 수정
Private Function CheckIfSyntax(ByVal v1 As String, ByVal p3Resolved As String) As Boolean
    On Error GoTo ErrHandler
    Dim parts() As String
    Dim op As String, v2 As String

    ' p3Resolved에 | 가 있는지 확인
    If InStr(p3Resolved, "|") > 0 Then
        parts = Split(p3Resolved, "|")
        ' Trim을 확실히 하여 공백 문제를 방지
        op = UCase(Trim(parts(0)))
        v2 = Trim(parts(1))
    Else
        ' 파이프가 없으면 전체를 값으로 보고 "EQ" 처리
        op = "EQ"
        v2 = Trim(p3Resolved)
    End If

    ' modActions에 최종 판단 위임
    CheckIfSyntax = modActions.Action_CheckCondition(v1, op, v2)

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function
' [Internal] 변수 값을 1 증가시키는 공통 로직
Private Sub IncrementVariable(ByVal varIdx As Long)
    On Error GoTo ErrHandler
    Dim currentVal As Double
    currentVal = val(modContext.GetVar(varIdx)) ' 현재 값 가져오기
    modContext.SetVar varIdx, CStr(currentVal + 1) ' 1 증가시켜 저장

CleanUp:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Sub


