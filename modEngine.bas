Attribute VB_Name = "modEngine"
' @desc 시트 명령어를 해석해 레지스트리 2단 디스패처로 실행 조립 담당
' @depends modCmdRegistry, modContext, modActions, modWinAPI
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

' [상수] 디스패치 반환 코드는 modCmdRegistry에 공개 정의 (CMD_NEXT_LINE / CMD_ERROR).
'   CMD_NOT_HANDLED 폴백은 레지스트리 도입으로 폐기됨.

'==============================================================================
' 1. Main Entry Point (RunRPA)
'==============================================================================
Private Enum eControlSignal
    ControlContinue = 0
    ControlPause = 1
    ControlAbort = 2
End Enum

Sub RunRPA()
Attribute RunRPA.VB_ProcData.VB_Invoke_Func = "W\n14"
    On Error GoTo FatalError

    Dim currentLine As Long
    Dim nextLine As Long
    Dim wsSource As Worksheet
    Dim rngTarget As Range
    Dim totalLines As Long
    Dim pureCmd As String
    Dim finalMessage As String

    finalMessage = "매크로가 종료되었습니다."
    modContext.InitContext

    '==========================================================
    ' [핵심: 스냅샷] 버튼을 누른 "바로 그 시트"를 고정합니다.
    ' 이후 모든 로직은 ActiveSheet가 아닌 wsSource를 바라봅니다.
    '==========================================================
    Set wsSource = ActiveSheet
    Set rngTarget = wsSource.Range("A1").CurrentRegion

    ' 유효성 검사: 데이터가 없거나 헤더만 있는 경우
    If rngTarget Is Nothing Then
        modActions.Action_ShowNotification "실행할 데이터가 없습니다. (현재 시트: " & wsSource.name & ")", 3
        GoTo CleanUp
    End If

    If rngTarget.Rows.count < 2 Then
        modActions.Action_ShowNotification "실행할 데이터가 없습니다. (현재 시트: " & wsSource.name & ")", 3
        GoTo CleanUp
    End If

    ' [수정] 헤더(1행) 제외: 실제 데이터는 2행부터 시작
    Set rngTarget = rngTarget.offset(1, 0).Resize(rngTarget.Rows.count - 1)

    ' [중요] 2차원 배열로 메모리에 완전히 적재 (메모리 맵 생성)
    ' 이후 루프 안에서는 Cells나 Range를 쓰지 말고 오직 m_arrMain만 써야 합니다.
    m_arrMain = rngTarget.value
    totalLines = UBound(m_arrMain, 1)

    ' [New] 0. 라벨 프리 스캔 (Pre-Scan)
    ' 메모리에 올라온 m_arrMain 지도를 보고 미리 위치를 파악합니다.
    PreScanLabels m_arrMain, totalLines

    ' [UI] 상태 표시 폼 로드
    modHelper.ShowRpaStatus frmRPAStatus, totalLines

    modContext.LogStatus "매크로 실행 중..."
    currentLine = 1

    Do While currentLine >= 1 And currentLine <= totalLines
        If CheckControlSignals() = ControlAbort Then
            finalMessage = "매크로가 중단되었습니다."
            GoTo CleanUp
        End If

        '----------------------------------------------------------
        ' [명령어 파싱 및 상태 업데이트]
        '----------------------------------------------------------
        Dim fullCmdText As String
        fullCmdText = m_arrMain(currentLine, 2)
        pureCmd = UCase(Trim(Replace(fullCmdText, ">", "")))

        ' [UI] 표시할 상세 정보 조합 (명령어 + 파라미터)
        Dim stepInfo As String
        stepInfo = pureCmd & " " & m_arrMain(currentLine, 3) & " " & m_arrMain(currentLine, 4) & " " & m_arrMain(currentLine, 5)

        ' [UI] 폼 업데이트
        modHelper.UpdateRpaStatus frmRPAStatus, currentLine, stepInfo
        DoEvents

        '----------------------------------------------------------
        ' [액션 실행 및 반환 계약 해석]
        '----------------------------------------------------------
        nextLine = ExecuteSmartAction(m_arrMain, currentLine)

        Select Case nextLine
            Case CMD_NEXT_LINE
                currentLine = NextPointer(currentLine, nextLine)

            Case CMD_HALT
                GoTo CleanUp

            Case CMD_ERROR
                modActions.Action_ShowNotification "[에러] 라인 " & currentLine & " / " & pureCmd, 3
                GoTo CleanUp

            Case Is > 0
                currentLine = NextPointer(currentLine, nextLine)

            Case Else
                modActions.Action_ShowNotification "[에러] 알 수 없는 반환 코드: " & nextLine, 3
                GoTo CleanUp
        End Select

        DoEvents
    Loop

CleanUp:
    On Error Resume Next
    modHelper.CloseRpaStatus frmRPAStatus
    modContext.ResetContext
    modActions.Action_ShowNotification finalMessage, 1
    On Error GoTo 0
    Exit Sub

FatalError:
    modActions.Action_ShowNotification _
        "[에러] 라인 " & currentLine & " / " & pureCmd & vbCrLf & Err.Description, 3
    finalMessage = "매크로가 종료되었습니다."
    Resume CleanUp
End Sub

Private Function CheckControlSignals() As eControlSignal
    CheckControlSignals = ControlContinue

    If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then
        CheckControlSignals = ControlAbort
        Exit Function
    End If

    If Not modWinAPI.IsKeyDown(modWinAPI.VK_PAUSE_RESUME) Then Exit Function

    CheckControlSignals = ControlPause
    modActions.Action_ShowNotification "일시정지됨. [Pause] 재개, [ESC] 종료", 3
    modWinAPI.Sleep 500

    Do
        If modWinAPI.IsKeyDown(modWinAPI.VK_ESCAPE) Then
            CheckControlSignals = ControlAbort
            Exit Function
        End If

        If modWinAPI.IsKeyDown(modWinAPI.VK_PAUSE_RESUME) Then
            modWinAPI.Sleep 200
            modContext.LogStatus "일시정지를 해제하고 현재 줄에서 재개합니다."
            CheckControlSignals = ControlContinue
            Exit Function
        End If

        DoEvents
        modWinAPI.Sleep 50
    Loop
End Function

Private Function NextPointer(ByVal curLine As Long, ByVal retCode As Long) As Long
    If retCode = CMD_NEXT_LINE Then
        NextPointer = curLine + 1
    Else
        NextPointer = retCode
    End If
End Function

' @desc 명령어 줄 받아 레지스트리 2단 디스패치로 실행, 미등록 시 CMD_ERROR 반환
Public Function ExecuteSmartAction(arr As Variant, R As Long) As Long
    On Error GoTo ErrHandler
    Dim sType As String

    ' 1. 명령어(Type) 추출
    sType = UCase(Trim(CStr(arr(R, 2))))

    ' [주석] #으로 시작하면 파라미터 해석 없이 즉시 다음 줄로 이동
    If Left(sType, 1) = "#" Then
        ExecuteSmartAction = CMD_NEXT_LINE
     GoTo CleanUp
    End If

    ' [들여쓰기 제거] ">" 기호 모두 제거 후 공백 제거 (예: ">> TYPE" -> "TYPE")
    sType = Trim(Replace(sType, ">", ""))

    ' 2. 레지스트리 게이트 조회 (모든 디스패치의 단일 통로)
    Dim def As CmdDef
    Set def = modCmdRegistry.GetCmdDef(sType)
    If def Is Nothing Then
        ' 미등록 명령어 (무인 환경: 비차단 알림)
        modActions.Action_ShowNotification "알 수 없는 명령어: " & sType, 3
        ExecuteSmartAction = CMD_ERROR
     GoTo CleanUp
    End If

    ' 3. decode: 슬롯별 ResolveMode(RAW/RESOLVE)로 args 구성 (decode/execute 분리)
    '    RAW = lvalue 주소(SET/INC/DEC P1) · 연산자(IF/WHILE P2). 그 외 = RESOLVE.
    Dim args(1 To 3) As String
    args(1) = DecodeSlot(def.p1, CStr(arr(R, 3)))
    args(2) = DecodeSlot(def.p2, CStr(arr(R, 4)))
    args(3) = DecodeSlot(def.p3, CStr(arr(R, 5)))

    ' 4. 2단 디스패치: 카테고리 -> 핸들러 클래스 -> 내부 명령어 분기
    Dim handler As ICommand
    Set handler = modCmdRegistry.handler(def.Category)
    ExecuteSmartAction = handler.Execute(sType, R, args)

CleanUp:
    Exit Function
ErrHandler:
    modActions.Action_ShowNotification "[에러] 라인 " & R & ": " & Err.Description, 3
    ExecuteSmartAction = CMD_ERROR
    Resume CleanUp
End Function

' [Helper] 슬롯 ResolveMode에 따라 RAW(원본) / RESOLVE(토큰치환) 디코드.
'   슬롯 미정의(Nothing)는 RESOLVE 기본.
Private Function DecodeSlot(ByVal spec As ParamSpec, ByVal rawVal As String) As String
    If spec Is Nothing Then
        DecodeSlot = ResolveParam(rawVal)
    ElseIf spec.IsRaw Then
        DecodeSlot = rawVal
    Else
        DecodeSlot = ResolveParam(rawVal)
    End If
End Function

' [엔진 서비스] GOTO/CALL 점프 대상 해석 (숫자=줄번호, 아니면 라벨명)
' @desc 점프 대상 문자열 받아 줄번호 반환, 미해석 라벨 시 CMD_ERROR 반환
Public Function ResolveJumpTarget(ByVal Target As String) As Long
    If IsNumeric(Target) Then
        ResolveJumpTarget = val(Target)
    Else
        Dim jt As Long: jt = modContext.GetLabel(Target)
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
    Err.Raise Err.Number, "modEngine.ResolveParam", Err.Description
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
    Err.Raise Err.Number, "modEngine.ResolveToken", Err.Description
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
    Err.Raise Err.Number, "modEngine.ResolvePixelToken", Err.Description
    Resume CleanUp
End Function

' [이관됨] ParseLValue는 clsMemoryCommands로 흡수됨 (lvalue 합법성 = 메모리 명령어 고유 지식)

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
    Err.Raise Err.Number, "modEngine.PreScanLabels", Err.Description
    Resume CleanUp
End Sub
' [엔진 서비스] 블록의 짝을 찾는 통합 탐색 함수 (m_arrMain 의존 -> modEngine 보유)
' StartLine: 탐색 시작 위치 (현재 줄)
' SearchType: 찾으려는 블록 종류 ("IF"/"WHILE"/"REPEAT"/"SELECT")
' Direction: 1 (아래로 탐색 - 시작점에서 끝 찾기), -1 (위로 탐색 - 끝점에서 시작 찾기)
' @desc 시작줄/블록종류/방향 받아 짝 줄번호 반환, 못 찾으면 -1 반환
Public Function FindBlockPair(ByVal startLine As Long, ByVal SearchType As String, Optional ByVal Direction As Long = 1) As Long
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
    Err.Raise Err.Number, "modEngine.FindBlockPair", Err.Description
    Resume CleanUp
End Function

' [엔진 서비스] IF/WHILE 조건절 판단 (v2: 좌변 P1 | 연산자 P2 | 우변 P3)
'   v1: 해석된 P1, op: P2(연산자, RAW), v2: 해석된 P3
' @desc 좌변/연산자/우변 받아 조건 참거짓 반환
Public Function CheckIfSyntax(ByVal v1 As String, ByVal op As String, ByVal v2 As String) As Boolean
    On Error GoTo ErrHandler
    CheckIfSyntax = modActions.Action_CheckCondition(v1, UCase(Trim(op)), v2)

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modEngine.CheckIfSyntax", Err.Description
    Resume CleanUp
End Function

' [엔진 서비스] m_arrMain 한 줄의 정규화된 명령어 반환 (clsFlowCommands IF/ELSE 판정용)
' @desc 줄번호 받아 들여쓰기 제거된 대문자 명령어 반환
Public Function CmdAt(ByVal R As Long) As String
    CmdAt = UCase(Trim(Replace(CStr(m_arrMain(R, 2)), ">", "")))
End Function

' [엔진 서비스] SELECT 전방 스캔 (m_arrMain 의존) -> 일치 CASE 다음 줄 / 폴백 / END_SELECT 반환
'   원 ProcessFlowControl SELECT 분기에서 추출. clsFlowCommands가 호출.
' @desc 현재 SELECT 줄과 대상값 받아 점프 줄번호 반환, 일치/폴백 없으면 END_SELECT 또는 CMD_ERROR
Public Function ResolveSelectTarget(ByVal R As Long, ByVal selVal As String) As Long
    On Error GoTo ErrHandler
    Dim scanR As Long, scanDepth As Long
    Dim scanCmd As String, caseVal As String
    Dim fallbackLine As Long

    scanDepth = 0
    fallbackLine = 0

    For scanR = R + 1 To UBound(m_arrMain, 1)
        scanCmd = UCase(Trim(Replace(m_arrMain(scanR, 2), ">", "")))

        If scanCmd = "SELECT" Then
            scanDepth = scanDepth + 1
        ElseIf scanCmd = "END_SELECT" Then
            If scanDepth > 0 Then
                scanDepth = scanDepth - 1
            Else
                ' SELECT 블록 끝 도달
                If fallbackLine > 0 Then
                    ResolveSelectTarget = fallbackLine + 1
                Else
                    ResolveSelectTarget = scanR ' END_SELECT 로 점프
                End If
             GoTo CleanUp
            End If
        ElseIf scanCmd = "CASE" Then
            If scanDepth = 0 Then
                caseVal = ResolveParam(CStr(m_arrMain(scanR, 3)))
                ' Case ELSE (또는 빈 값) = 폴백
                If UCase(caseVal) = "ELSE" Or caseVal = "" Then
                    fallbackLine = scanR
                Else
                    If modActions.Action_CheckCondition(selVal, "EQ", caseVal) Then
                        ResolveSelectTarget = scanR + 1
                     GoTo CleanUp
                    End If
                End If
            End If
        End If
    Next scanR

    ResolveSelectTarget = CMD_ERROR

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modEngine.ResolveSelectTarget", Err.Description
    Resume CleanUp
End Function


