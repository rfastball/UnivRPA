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
                                    ' 무인 환경: 비차단 알림 후 안전 종료 (모달 정지 금지)
                                    modActions.Action_ShowNotification _
                                    "[에러] 라인 " & currentLine & " / " & pureCmd & vbCrLf & Err.Description, 3

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
    '    - IF, WHILE, REPEAT, GOTO 등 라인 점프가 주 목적
    '    - sP2(raw)=연산자, rP3=우변
    result = ProcessFlowControl(sType, R, rP1, sP2, rP3)
    If result <> CMD_NOT_HANDLED Then
        ExecuteSmartAction_Core = result
        GoTo CleanUp
    End If

    ' 4. 액션 (Action) 처리
    '    - 입력, 비전, 메모리 조작 등 동작 수행이 주 목적
    '    - sP1(raw)=SET 좌변(주소), rP1/rP2/rP3=해석된 값
    result = ProcessAction(sType, R, sP1, rP1, rP2, rP3)
    If result <> CMD_NOT_HANDLED Then
        ExecuteSmartAction_Core = result
        GoTo CleanUp
    End If

    ' 5. 처리되지 않은 명령어 (무인 환경: 비차단 알림)
    modActions.Action_ShowNotification "알 수 없는 명령어: " & sType, 3
    ExecuteSmartAction_Core = CMD_ERROR

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Internal] 액션 명령어 처리 (v2: 입력 / 메모리 / 시스템 / 비전)
'   sP1Raw = SET 좌변(주소, RAW), rP1/rP2/rP3 = 해석된 P1/P2/P3
Private Function ProcessAction(sType As String, R As Long, sP1Raw As String, rP1 As String, rP2 As String, rP3 As String) As Long
    On Error GoTo ErrHandler

    Dim p1 As Long, p2 As Long
    p1 = CLng(val(rP1))
    p2 = CLng(val(rP2))

    ProcessAction = CMD_NEXT_LINE ' 기본값: 성공 시 다음 줄

    Select Case sType
        '--------------------------------------------------
        ' [Input] 마우스 / 키보드
        '--------------------------------------------------
        Case "CLICK":  modActions.Action_MouseClick p1, p2, False
        Case "RCLICK": modActions.Action_MouseClick p1, p2, True
        Case "KEY":    modActions.Action_PressKey rP1
        Case "COMBO":  modActions.Action_PressCombo rP1, rP2

        Case "TYPE"
            ' P1 = 딜레이ms, P3 = 텍스트
            modActions.Action_SmartType CLng(val(rP1)), rP3

        '--------------------------------------------------
        ' [Memory / Clipboard] SET (좌변 RAW = 주소, 우변 rP3 = 값)
        '--------------------------------------------------
        Case "SET"
            Dim kind As String, key As String
            If ParseLValue(sP1Raw, kind, key) Then
                If kind = "CLIP" Then
                    modWinAPI.SetClipboardText rP3
                Else
                    modContext.SetVar key, rP3
                End If
            Else
                ' 좌변이 쓰기 가능한 토큰이 아님 (읽기전용/형식오류)
                ProcessAction = CMD_ERROR
                GoTo CleanUp
            End If

        '--------------------------------------------------
        ' [System]
        '--------------------------------------------------
        Case "FOCUS"
            If Not modActions.Action_FocusWindow(rP3) Then
                ProcessAction = CMD_ERROR
                GoTo CleanUp
            End If
            modActions.BetterWait 500

        Case "SLEEP"
            modActions.BetterWait CLng(val(rP3))

        Case "NOTIFY"
            ' P3 = "메시지[|초]"
            Dim nMsg As String, nSec As Long, durMs As Long
            nMsg = rP3: nSec = 0
            If InStr(rP3, "|") > 0 Then
                Dim np() As String: np = Split(rP3, "|")
                nMsg = np(0)
                nSec = CLng(val(Trim(np(1))))
            End If
            If nSec > 0 Then durMs = nSec * 1000 Else durMs = 2000
            modActions.Action_Notify durMs, 64, nMsg

        Case "EXIT"
            ProcessAction = CMD_ERROR
            GoTo CleanUp

        '--------------------------------------------------
        ' [Vision] 타임아웃 가능 (실패 시 CMD_ERROR)
        '--------------------------------------------------
        Case "WAIT_APPEAR", "WAIT_VANISH"
            If Not modActions.Action_Wait(sType, p1, p2, rP3) Then
                ProcessAction = CMD_ERROR
                GoTo CleanUp
            End If

        Case "CLICK_SYNC"
            If Not modActions.Action_ClickSync(p1, p2) Then
                ProcessAction = CMD_ERROR
                GoTo CleanUp
            End If

        Case "ASSERT"
            If Not modActions.Action_Assert(p1, p2, rP3) Then
                ProcessAction = CMD_ERROR
                GoTo CleanUp
            End If

        Case Else
            ProcessAction = CMD_NOT_HANDLED
            GoTo CleanUp
    End Select

    ' [사후 대기] CLICK/RCLICK/KEY/COMBO 만 P3에 대기시간(ms)을 가질 수 있음
    '   (TYPE=텍스트, FOCUS=제목 이므로 제외)
    If IsNumeric(rP3) And val(rP3) > 0 Then
        Select Case sType
            Case "CLICK", "RCLICK", "KEY", "COMBO"
                modActions.BetterWait CLng(val(rP3))
        End Select
    End If

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Internal] 흐름 제어 명령어 처리 (v2: 분기, 반복, 점프)
'   rP1 = 해석된 좌변, sP2 = 연산자(raw), rP3 = 해석된 우변/대상
Private Function ProcessFlowControl(sType As String, R As Long, rP1 As String, sP2 As String, rP3 As String) As Long
    On Error GoTo ErrHandler

    ProcessFlowControl = CMD_NEXT_LINE

    Select Case sType
         Case "LABEL": ' Pass

         Case "GOTO"
            ProcessFlowControl = ResolveJumpTarget(rP3)

         Case "CALL"
            modContext.PushStack R + 1
            ProcessFlowControl = ResolveJumpTarget(rP3)

         Case "RETURN"
            Dim retLine As Long
            retLine = modContext.PopStack()
            If retLine > 0 Then ProcessFlowControl = retLine Else ProcessFlowControl = CMD_ERROR

         Case "IF"
            If CheckIfSyntax(rP1, sP2, rP3) Then
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

         '--------------------------------------------------
         ' [Loop] WHILE / END_WHILE (활성 루프 스택)
         '--------------------------------------------------
         Case "WHILE"
            If CheckIfSyntax(rP1, sP2, rP3) Then
                ' 참: 첫 진입이면 PUSH, 백엣지 재진입(top.header=R)이면 회차+1
                If modContext.LoopDepth() > 0 And modContext.LoopTopHeader() = R Then
                    modContext.LoopIncrement
                Else
                    modContext.LoopPush "WHILE", R, FindBlockPair(R, "WHILE", 1), 0
                End If
                ProcessFlowControl = CMD_NEXT_LINE
            Else
                ' 거짓: 활성 프레임이 이 WHILE이면 POP 후 탈출
                If modContext.LoopDepth() > 0 And modContext.LoopTopHeader() = R Then modContext.LoopPop
                Dim wEnd As Long: wEnd = FindBlockPair(R, "WHILE", 1)
                If wEnd > 0 Then ProcessFlowControl = wEnd + 1 Else ProcessFlowControl = CMD_ERROR
            End If

         Case "END_WHILE"
            ' 헤더로 백점프. POP은 WHILE 거짓 갈래가 담당 (백엣지 동안 {$LOOP} 유지)
            Dim wHead As Long: wHead = FindBlockPair(R, "WHILE", -1)
            If wHead > 0 Then ProcessFlowControl = wHead Else ProcessFlowControl = CMD_ERROR

         '--------------------------------------------------
         ' [Loop] REPEAT / END_REPEAT (N회 반복)
         '--------------------------------------------------
         Case "REPEAT"
            Dim maxCnt As Long: maxCnt = CLng(val(rP1))
            If maxCnt <= 0 Then
                ' 0회 -> 본문 건너뛰기 (PUSH 안 함)
                Dim rEndSkip As Long: rEndSkip = FindBlockPair(R, "REPEAT", 1)
                If rEndSkip > 0 Then ProcessFlowControl = rEndSkip + 1 Else ProcessFlowControl = CMD_ERROR
            Else
                modContext.LoopPush "REPEAT", R, FindBlockPair(R, "REPEAT", 1), maxCnt
                ProcessFlowControl = CMD_NEXT_LINE
            End If

         Case "END_REPEAT"
            Dim rHead As Long: rHead = FindBlockPair(R, "REPEAT", -1)
            If rHead > 0 Then
                If modContext.GetActiveLoopIteration() < modContext.LoopTopMax() Then
                    modContext.LoopIncrement
                    ProcessFlowControl = rHead + 1   ' 본문 재실행 (REPEAT 줄은 1회만)
                Else
                    modContext.LoopPop
                    ProcessFlowControl = CMD_NEXT_LINE
                End If
            Else
                ProcessFlowControl = CMD_ERROR
            End If

         '--------------------------------------------------
         ' [Loop Control] BREAK / CONTINUE
         '--------------------------------------------------
         Case "BREAK"
            If modContext.LoopDepth() > 0 Then
                Dim brkEnd As Long: brkEnd = modContext.LoopTopEnd()
                modContext.LoopPop ' 카운터 정리
                ProcessFlowControl = brkEnd + 1
            Else
                ProcessFlowControl = CMD_ERROR
            End If

         Case "CONTINUE"
            If modContext.LoopDepth() > 0 Then
                ' 감싸는 루프의 ENDER로 점프 → WHILE/REPEAT 각각의 ENDER가 알맞게 처리
                ProcessFlowControl = modContext.LoopTopEnd()
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

' [Helper] GOTO/CALL 점프 대상 해석 (숫자=줄번호, 아니면 라벨명)
Private Function ResolveJumpTarget(ByVal target As String) As Long
    If IsNumeric(target) Then
        ResolveJumpTarget = val(target)
    Else
        Dim jt As Long: jt = modContext.GetLabel(target)
        If jt > 0 Then ResolveJumpTarget = jt Else ResolveJumpTarget = CMD_ERROR
    End If
End Function

' [Internal] 파라미터 해석기 (v2 토크나이저)
'   {...} 스팬을 깊이 인식으로 스캔 -> 내용으로 분기 ($시스템/$PIXEL/리터럴/{name})
'   중첩 중괄호 지원: {$PIXEL:{cx},{cy}}. 값 치환만 수행하며 *? 와일드카드는 건드리지 않음.
Private Function ResolveParam(ByVal rawVal As String) As String
    On Error GoTo ErrHandler
    Dim src As String: src = rawVal
    Dim out As String
    Dim i As Long, n As Long
    Dim ch As String

    n = Len(src)
    i = 1
    Do While i <= n
        ch = Mid$(src, i, 1)
        If ch = "{" Then
            ' 짝이 되는 '}' 를 깊이 카운팅으로 탐색 (중첩 허용)
            Dim depth As Long, j As Long, found As Long
            depth = 1: found = 0
            For j = i + 1 To n
                Select Case Mid$(src, j, 1)
                    Case "{": depth = depth + 1
                    Case "}"
                        depth = depth - 1
                        If depth = 0 Then found = j: Exit For
                End Select
            Next j

            If found = 0 Then
                ' 짝 없음 -> 남은 문자열을 리터럴로 그대로 출력 (graceful fallback)
                out = out & Mid$(src, i)
                Exit Do
            End If

            Dim inner As String
            inner = Mid$(src, i + 1, found - i - 1)
            ' 내부의 자식 토큰을 먼저 해석한 뒤 분류
            out = out & ResolveToken(ResolveParam(inner))
            i = found + 1
        Else
            out = out & ch
            i = i + 1
        End If
    Loop

    ResolveParam = out

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Helper] 단일 토큰(중괄호 내부 문자열)을 값으로 분류/치환
Private Function ResolveToken(ByVal inner As String) As String
    On Error GoTo ErrHandler
    Dim key As String: key = Trim(inner)
    Dim up As String: up = UCase$(key)

    ' [리터럴/시스템 고정 토큰]
    Select Case up
        Case "$BLANK": ResolveToken = "": Exit Function
        Case "$TAB":   ResolveToken = vbTab: Exit Function
        Case "$CR":    ResolveToken = vbCr: Exit Function
        Case "$LF":    ResolveToken = vbLf: Exit Function
        Case "$SPACE": ResolveToken = " ": Exit Function
        Case "$CLIP":  ResolveToken = modWinAPI.GetClipboardText(): Exit Function
        Case "$LOOP":  ResolveToken = CStr(modContext.GetActiveLoopIteration()): Exit Function
    End Select

    ' [시스템 동적 토큰] {$PIXEL:x,y} (좌표는 이미 해석됨)
    If Left$(up, 7) = "$PIXEL:" Then
        ResolveToken = ResolvePixelToken(Mid$(key, 8))
        Exit Function
    End If

    ' [미지의 시스템 토큰] -> 오타 노출 위해 리터럴 그대로 반환
    If Left$(key, 1) = "$" Then
        ResolveToken = "{" & inner & "}"
        Exit Function
    End If

    ' [사용자 메모리] {name}
    ResolveToken = modContext.GetVar(key)

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Helper] $PIXEL 좌표 -> 색상 헥스 (coordText 는 이미 해석된 "x,y")
Private Function ResolvePixelToken(ByVal coordText As String) As String
    On Error GoTo ErrHandler
    Dim coords() As String
    coords = Split(coordText, ",")
    If UBound(coords) = 1 Then
        ResolvePixelToken = modWinAPI.GetColorHexAt(CLng(val(Trim(coords(0)))), CLng(val(Trim(coords(1)))))
    Else
        ResolvePixelToken = "{$PIXEL:" & coordText & "}"  ' 형식 불일치 시 원본
    End If

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function

' [Helper] SET 좌변(주소) 파싱. 좌변은 RAW (값 해석 금지)
'   반환 True 시 kind="MEM"|"CLIP", key=메모리명. 읽기전용 토큰은 좌변 불가(False).
Private Function ParseLValue(ByVal rawP1 As String, ByRef kind As String, ByRef key As String) As Boolean
    Dim s As String: s = Trim(rawP1)
    ParseLValue = False
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "{" Or Right$(s, 1) <> "}" Then Exit Function

    Dim inner As String: inner = Trim(Mid$(s, 2, Len(s) - 2))
    If UCase$(inner) = "$CLIP" Then
        kind = "CLIP": key = ""
        ParseLValue = True
    ElseIf Left$(inner, 1) = "$" Then
        ' 읽기 전용 시스템 토큰은 좌변(쓰기) 불가
        ParseLValue = False
    Else
        kind = "MEM": key = inner
        ParseLValue = True
    End If
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
                If Direction = 1 Then ' 아래로 END_WHILE 찾기
                    If sCmd = "WHILE" Then depth = depth + 1
                    If sCmd = "END_WHILE" Then
                        If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                        depth = depth - 1
                    End If
                Else ' 위로 WHILE 찾기
                    If sCmd = "END_WHILE" Then depth = depth + 1
                    If sCmd = "WHILE" Then
                        If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                        depth = depth - 1
                    End If
                End If

            Case "REPEAT"
                If Direction = 1 Then ' 아래로 END_REPEAT 찾기
                    If sCmd = "REPEAT" Then depth = depth + 1
                    If sCmd = "END_REPEAT" Then
                        If depth = 0 Then FindBlockPair = R: GoTo CleanUp
                        depth = depth - 1
                    End If
                Else ' 위로 REPEAT 찾기
                    If sCmd = "END_REPEAT" Then depth = depth + 1
                    If sCmd = "REPEAT" Then
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

' [Internal] IF/WHILE 조건절 판단 (v2: 좌변 P1 | 연산자 P2 | 우변 P3)
'   v1: 해석된 P1, op: P2(연산자), v2: 해석된 P3
Private Function CheckIfSyntax(ByVal v1 As String, ByVal op As String, ByVal v2 As String) As Boolean
    On Error GoTo ErrHandler
    CheckIfSyntax = modActions.Action_CheckCondition(v1, UCase(Trim(op)), v2)

CleanUp:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume CleanUp
End Function


