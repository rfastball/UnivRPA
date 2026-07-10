Attribute VB_Name = "modEngine"
' @desc 시트 명령어를 해석해 레지스트리 2단 디스패처로 실행 조립 담당
' @depends modCmdRegistry, modContext, modActions, modTokenResolver, modHelper, modWinAPI, modCondition, modParse, LogShim
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
' [ProgramImage] 로드된 스크립트 프로그램 이미지
Private m_program As clsScriptProgram
Private m_lastRuntimeError As String

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
    On Error GoTo FatalError

    Dim currentLine As Long
    Dim nextLine As Long
    Dim wsSource As Worksheet
    Dim rngTarget As Range
    Dim totalLines As Long
    Dim pureCmd As String
    Dim finalMessage As String
    Dim fatalDesc As String

    finalMessage = "매크로가 종료되었습니다."
    modActions.SetUnattendedMode False   ' 수동 실행 = 대화형 알림
    modActions.ClearAbort                ' 이전 실행의 중단 래치 해제
    modContext.InitContext

    '==========================================================
    ' [핵심: 스냅샷] 버튼을 누른 "바로 그 시트"를 고정합니다.
    ' 이후 모든 로직은 ActiveSheet가 아닌 wsSource를 바라봅니다.
    '==========================================================
    Set wsSource = ActiveSheet
    Set rngTarget = ResolveScriptRange(wsSource)

    ' 유효성 검사: 데이터가 없거나 헤더만 있는 경우
    If rngTarget Is Nothing Then
        modActions.Action_ShowNotification "실행할 데이터가 없습니다. (현재 시트: " & wsSource.name & ")", 3
        GoTo CleanUp
    End If

    ' [ProgramImage] 시트 스냅샷을 구조 인덱스/라벨 인덱스를 가진 Program으로 승격
    Set m_program = New clsScriptProgram
    m_program.LoadFromRange rngTarget
    totalLines = m_program.totalLines

    Dim isProgramValid As Boolean
    isProgramValid = m_program.ValidateStructure()
    isProgramValid = m_program.ValidateLabels() And isProgramValid
    If Not isProgramValid Then
        modActions.Action_ShowNotification "[에러] 스크립트 구조 오류" & vbCrLf & m_program.ValidationErrors(), 3
        GoTo CleanUp
    End If

    ' [UI] 상태 표시 폼 로드
    modHelper.ShowRpaStatus frmRPAStatus, totalLines

    ' [운용 로그] 세션 시작: 이전 jsonl 비우고 run_start 기록
    LogShim.Reset
    LogShim.LogEvent "run_start", _
        Array("sheet", "total"), Array(wsSource.name, totalLines)

    modContext.LogStatus "매크로 실행 중..."
    currentLine = 1

    Do While currentLine >= 1 And currentLine <= totalLines
        If CheckControlSignals() = ControlAbort Then
            finalMessage = "매크로가 중단되었습니다."
            LogForeachResumeHint "foreach.interrupted"
            If Len(ForeachResumeHint()) > 0 Then finalMessage = finalMessage & vbCrLf & ForeachResumeHint()
            GoTo CleanUp
        End If

        '----------------------------------------------------------
        ' [명령어 파싱 및 상태 업데이트]
        '----------------------------------------------------------
        Dim inst As clsInstruction
        Set inst = m_program.InstructionAt(currentLine)
        pureCmd = inst.cmd

        ' [UI] 표시할 상세 정보 조합 (명령어 + 파라미터)
        Dim stepInfo As String
        stepInfo = inst.cmd & " " & inst.p1 & " " & inst.p2 & " " & inst.p3

        ' [UI] 폼 업데이트
        Dim activeRidx As Long
        Dim activeRtotal As Long
        activeRidx = modContext.GetActiveForeachOrdinal()
        activeRtotal = modContext.GetActiveForeachTotalRows()
        modHelper.UpdateRpaStatus frmRPAStatus, currentLine, stepInfo, activeRidx, activeRtotal
        UpdateStatusBarProgress currentLine, totalLines, activeRidx, activeRtotal
        DoEvents

        '----------------------------------------------------------
        ' [액션 실행 및 반환 계약 해석]
        '----------------------------------------------------------
        ' [운용 로그] 줄 실행 (DEBUG: 운용 시 MIN_LEVEL=INFO로 꺼짐)
        LogShim.LogD "DEBUG", "line=" & inst.RuntimeLine & " cmd=" & inst.cmd

        nextLine = ExecuteSmartAction(currentLine)

        ' [B2] 중단 래치 우선 판정: 액션 내부에서 ESC로 탈출한 경우
        ' cmd_error가 아니라 "사용자 중단"으로 즉시 종료한다
        If modActions.AbortPending() Then
            finalMessage = "매크로가 중단되었습니다."
            LogForeachResumeHint "foreach.interrupted"
            If Len(ForeachResumeHint()) > 0 Then finalMessage = finalMessage & vbCrLf & ForeachResumeHint()
            GoTo CleanUp
        End If

        Select Case nextLine
            Case CMD_NEXT_LINE
                currentLine = NextPointer(currentLine, nextLine)

            Case CMD_HALT
                GoTo CleanUp

            Case CMD_ERROR
                ' 실패 사유는 ExecuteSmartAction이 m_lastRuntimeError에 남긴다 - 알림은 여기서 1회
                Dim errDetail As String
                errDetail = LastRuntimeError()
                If Len(errDetail) = 0 Then errDetail = "명령 실행 실패"
                LogShim.LogEvent "cmd_error", _
                    Array("line", "sheetRow", "cmd", "reason"), _
                    Array(inst.RuntimeLine, inst.sheetRow, inst.cmd, errDetail)
                LogForeachResumeHint "foreach.cmd_error"
                modActions.Action_ShowNotification _
                    "[에러] 라인 " & inst.RuntimeLine & " / 시트 행 " & inst.sheetRow & " / " & inst.cmd & _
                    vbCrLf & errDetail & ForeachHintSuffix(), 3
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
    ' [운용 로그] 세션 종료 (finalMessage가 상태를 담음)
    LogShim.LogEvent "run_end", Array("status"), Array(finalMessage)
    modHelper.CloseRpaStatus frmRPAStatus
    modContext.ResetContext
    Set m_program = Nothing
    modActions.Action_ShowNotification finalMessage, 1
    On Error GoTo 0
    Exit Sub

FatalError:
    fatalDesc = Err.Description
    LogShim.LogEvent "fatal_error", _
        Array("line", "cmd", "desc"), _
        Array(currentLine, pureCmd, fatalDesc)
    LogForeachResumeHint "foreach.fatal_error"
    modActions.Action_ShowNotification _
        "[에러] 라인 " & currentLine & " / " & pureCmd & vbCrLf & fatalDesc & ForeachHintSuffix(), 3
    finalMessage = "매크로가 종료되었습니다."
    Resume CleanUp
End Sub

Private Sub UpdateStatusBarProgress(ByVal curLine As Long, ByVal totalLines As Long, _
                                    ByVal ridx As Long, ByVal rtotal As Long)
    Dim msg As String
    msg = "Line " & curLine & "/" & totalLines
    If ridx > 0 Then msg = msg & " | FOREACH_ROW RIDX " & ridx & "/" & rtotal
    modContext.LogStatus msg
End Sub

Private Function ForeachResumeHint() As String
    Dim ridx As Long
    ridx = modContext.GetActiveForeachOrdinal()
    If ridx <= 0 Then Exit Function

    ForeachResumeHint = "FOREACH_ROW RIDX " & ridx & "/" & _
                        modContext.GetActiveForeachTotalRows() & " 처리 중"
End Function

Private Function ForeachHintSuffix() As String
    Dim hint As String
    hint = ForeachResumeHint()
    If Len(hint) > 0 Then ForeachHintSuffix = vbCrLf & hint
End Function

Private Sub LogForeachResumeHint(ByVal eventName As String)
    Dim ridx As Long
    ridx = modContext.GetActiveForeachOrdinal()
    If ridx <= 0 Then Exit Sub

    LogShim.LogEvent eventName, _
        Array("headerLine", "ridx", "rtotal"), _
        Array(modContext.GetActiveForeachHeaderLine(), ridx, modContext.GetActiveForeachTotalRows())
End Sub

Private Function CheckControlSignals() As eControlSignal
    CheckControlSignals = ControlContinue

    If modActions.AbortPending() Then
        CheckControlSignals = ControlAbort
        Exit Function
    End If

    If Not modWinAPI.IsKeyDown(modWinAPI.VK_PAUSE_RESUME) Then Exit Function

    CheckControlSignals = ControlPause
    modActions.Action_ShowNotification "일시정지됨. [Pause] 재개, [ESC] 종료", 3
    modWinAPI.Sleep 500

    Do
        If modActions.AbortPending() Then
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

Public Function ActiveProgram() As clsScriptProgram
    Set ActiveProgram = m_program
End Function

' @desc 스크립트 시트 받아 실행 범위 A2:F{lastRow} 반환, 데이터 없으면 Nothing 반환
Public Function ResolveScriptRange(ByVal ws As Worksheet) As Range
    ' CurrentRegion은 중간 빈 행/열에서 조용히 잘려 스크립트가 반만 실행된다.
    ' B열(COMMAND) 기준 last row로 A~F 6열 고정 범위를 산정한다 (빈 행은 실행기가 스킵).
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    Set ResolveScriptRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6))
End Function

Public Function LastRuntimeError() As String
    LastRuntimeError = m_lastRuntimeError
End Function

Public Sub LoadProgramFromArray(ByVal sourceArray As Variant, Optional ByVal firstSheetRow As Long = 1)
    ' LoadFromArray가 구조/라벨 인덱스와 정적 분석을 이미 전부 빌드한다 (재호출 불필요)
    Set m_program = New clsScriptProgram
    m_program.LoadFromArray sourceArray, firstSheetRow
End Sub

' @desc 명령어 줄 받아 레지스트리 2단 디스패치로 실행, 미등록 시 CMD_ERROR 반환
Public Function ExecuteSmartAction(ByVal R As Long) As Long
    On Error GoTo ErrHandler
    m_lastRuntimeError = ""
    modContext.ClearLastError
    If m_program Is Nothing Then
        m_lastRuntimeError = "실행할 ProgramImage가 없습니다."
        ExecuteSmartAction = CMD_ERROR
        GoTo CleanUp
    End If

    Dim inst As clsInstruction
    Set inst = m_program.InstructionAt(R)

    ' [주석] #으로 시작하면 파라미터 해석 없이 즉시 다음 줄로 이동
    If m_program.IsBlankLine(R) Or m_program.IsCommentLine(R) Then
        ExecuteSmartAction = CMD_NEXT_LINE
     GoTo CleanUp
    End If

    ' 2. 레지스트리 게이트 조회 (모든 디스패치의 단일 통로)
    If inst.def Is Nothing Then
        ' 미등록 명령어: 사유만 기록, 알림은 호출자(RunRPA)가 1회 조립
        m_lastRuntimeError = "알 수 없는 명령어: " & inst.cmd
        ExecuteSmartAction = CMD_ERROR
     GoTo CleanUp
    End If

    ' 3. decode: 슬롯별 ResolveMode(RAW/RESOLVE)로 args 구성 (decode/execute 분리)
    '    RAW 슬롯 4종은 레지스트리가 RawKind로 선언 (연산자/lvalue/점프대상/라벨명)
    Dim args(1 To 3) As String
    args(1) = DecodeSlot(inst.def.p1, inst.p1)
    args(2) = DecodeSlot(inst.def.p2, inst.p2)
    args(3) = DecodeSlot(inst.def.p3, inst.p3)

    ' 3.5 ParamSpec 게이트: 레지스트리가 선언한 Required/DataType을 실행 전 강제 (C1)
    Dim guardMsg As String
    guardMsg = ValidateDecodedSlots(inst, args)
    If Len(guardMsg) > 0 Then
        m_lastRuntimeError = guardMsg
        LogShim.LogEvent "param_guard", _
            Array("line", "sheetRow", "cmd", "reason"), _
            Array(inst.RuntimeLine, inst.sheetRow, inst.cmd, guardMsg)
        ExecuteSmartAction = CMD_ERROR
        GoTo CleanUp
    End If

    ' 4. 2단 디스패치: 카테고리 -> 핸들러 클래스 -> 내부 명령어 분기
    Dim handler As ICommand
    Set handler = modCmdRegistry.handler(inst.def.Category)
    ExecuteSmartAction = handler.Execute(inst.cmd, inst.RuntimeLine, args)

    ' 핸들러 실패 사유 합류 (B1): 핸들러는 modContext.SetLastError로 기록하고 -1만 반환한다
    If ExecuteSmartAction = CMD_ERROR And Len(m_lastRuntimeError) = 0 Then
        m_lastRuntimeError = modContext.LastError()
    End If

CleanUp:
    Exit Function
ErrHandler:
    m_lastRuntimeError = Err.Description
    ExecuteSmartAction = CMD_ERROR
    Resume CleanUp
End Function

' [Helper] ParamSpec 게이트 (C1). Required는 원문 기준으로 판정해
'   {$BLANK} 등 "의도적으로 빈 값" 작성을 허용하고, DT_NUM은 핸들러가
'   실제로 소비할 해석값 기준으로 판정한다.
' @desc 명령줄과 해석된 args 받아 첫 위반 메시지 반환, 통과 시 빈 문자열 반환
Private Function ValidateDecodedSlots(ByVal inst As clsInstruction, ByRef args() As String) As String
    Dim i As Long
    Dim spec As ParamSpec
    Dim rawVal As String
    Dim numCheck As Double

    For i = 1 To 3
        Set spec = inst.def.Param(i)
        If Not spec Is Nothing Then
            Select Case i
                Case 1: rawVal = inst.p1
                Case 2: rawVal = inst.p2
                Case 3: rawVal = inst.p3
            End Select

            If spec.Required And Len(Trim$(rawVal)) = 0 Then
                ValidateDecodedSlots = "P" & i & "(" & spec.Meaning & ") 필수값 누락"
                Exit Function
            End If

            If spec.DataType = DT_NUM Then
                ' 판정은 핸들러가 소비할 파서(modParse)와 동일해야 한다
                If Len(Trim$(args(i))) > 0 And Not modParse.TryParseDoubleStrict(args(i), numCheck) Then
                    ValidateDecodedSlots = "P" & i & "(" & spec.Meaning & ") 숫자 아님: '" & args(i) & "'"
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' [Helper] 슬롯 ResolveMode에 따라 RAW(원본) / RESOLVE(토큰치환) 디코드.
'   슬롯 미정의(Nothing)는 RESOLVE 기본.
Private Function DecodeSlot(ByVal spec As ParamSpec, ByVal rawVal As String) As String
    If spec Is Nothing Then
        DecodeSlot = modTokenResolver.ResolveParam(rawVal)
    ElseIf spec.IsRaw Then
        DecodeSlot = rawVal
    Else
        DecodeSlot = modTokenResolver.ResolveParam(rawVal)
    End If
End Function

' [엔진 서비스] GOTO/CALL 점프 대상 해석 (정적 라벨명만 허용, 숫자/토큰은 거부)
' @desc 점프 대상 문자열 받아 줄번호 반환, 미해석 라벨 시 CMD_ERROR 반환
Public Function ResolveJumpTarget(ByVal Target As String) As Long
    If m_program Is Nothing Then
        ResolveJumpTarget = CMD_ERROR
    Else
        ResolveJumpTarget = m_program.ResolveJumpTarget(Target)
    End If
End Function

Public Function ResolveGotoTarget(ByVal sourceLine As Long, ByVal Target As String, ByRef framesPopped As Long) As Long
    On Error GoTo ErrHandler
    ResolveGotoTarget = CMD_ERROR
    framesPopped = 0

    If m_program Is Nothing Then
        m_lastRuntimeError = "GOTO 라인 " & sourceLine & ": 실행할 ProgramImage가 없습니다."
        Exit Function
    End If

    Dim targetLine As Long
    targetLine = m_program.ResolveJumpTarget(Target)
    If targetLine <= 0 Then
        m_lastRuntimeError = "GOTO 라인 " & sourceLine & ": 대상 라벨 해석 실패: '" & Target & "' (정적 라벨명만 허용)"
        Exit Function
    End If

    If Not m_program.IsGotoAllowed(sourceLine, targetLine) Then
        m_lastRuntimeError = "GOTO 라인 " & sourceLine & ": 라인 " & targetLine & "으로의 점프가 흐름 경계(루프/영역)를 침범합니다."
        Exit Function
    End If

    Dim popped As Long
    popped = modContext.ReconcileLoopStackForJump(targetLine, m_program.LoopDepthAt(targetLine))
    If popped < 0 Then
        m_lastRuntimeError = "GOTO 라인 " & sourceLine & ": 대상 라인 " & targetLine & "의 루프 깊이와 활성 루프 스택이 불일치합니다."
        Exit Function
    End If
    modContext.ReconcileIfBranchStackForJump targetLine

    framesPopped = popped
    ResolveGotoTarget = targetLine
    Exit Function

ErrHandler:
    m_lastRuntimeError = "GOTO 라인 " & sourceLine & " 해석 오류: " & Err.Description
    ResolveGotoTarget = CMD_ERROR
End Function

Public Function ResolveCallTarget(ByVal sourceLine As Long, ByVal Target As String) As Long
    On Error GoTo ErrHandler
    If m_program Is Nothing Then
        m_lastRuntimeError = "CALL 라인 " & sourceLine & ": 실행할 ProgramImage가 없습니다."
        ResolveCallTarget = CMD_ERROR
    Else
        ResolveCallTarget = m_program.ResolveJumpTarget(Target)
        If ResolveCallTarget <= 0 Then
            m_lastRuntimeError = "CALL 라인 " & sourceLine & ": 대상 라벨 해석 실패: '" & Target & "' (정적 라벨명만 허용)"
        ElseIf Not m_program.IsCallAllowed(sourceLine, ResolveCallTarget) Then
            ' [CALL 경계] 루프 헤더를 거치지 않은 본문 진입은 LoopStack 프레임 부재로
            '   END_* 백점프/RETURN 깊이 복원이 원인에서 먼 곳에서 깨진다 - 여기서 차단
            m_lastRuntimeError = "CALL 라인 " & sourceLine & ": 라인 " & ResolveCallTarget & "은 루프 본문 내부라 밖에서 진입할 수 없습니다."
            ResolveCallTarget = CMD_ERROR
        End If
    End If
    Exit Function

ErrHandler:
    m_lastRuntimeError = "CALL 라인 " & sourceLine & " 해석 오류: " & Err.Description
    ResolveCallTarget = CMD_ERROR
End Function

' [이관됨 - D4] ResolveParam/ResolveToken 계열 토큰 해석기는 modTokenResolver로 추출됨
' [이관됨] ParseLValue는 clsCommandsMemory로 흡수됨 (lvalue 합법성 = 메모리 명령어 고유 지식)

' [호환 엔진 서비스] 블록의 짝 탐색은 ActiveProgram으로 위임
' StartLine: 탐색 시작 위치 (현재 줄)
' SearchType: 찾으려는 블록 종류 ("IF"/"WHILE"/"REPEAT"/"SELECT")
' Direction: 1 (아래로 탐색 - 시작점에서 끝 찾기), -1 (위로 탐색 - 끝점에서 시작 찾기)
' @desc 시작줄/블록종류/방향 받아 짝 줄번호 반환, 못 찾으면 -1 반환
Public Function FindBlockPair(ByVal startLine As Long, ByVal SearchType As String, Optional ByVal Direction As Long = 1) As Long
    If m_program Is Nothing Then
        FindBlockPair = CMD_ERROR
    Else
        FindBlockPair = m_program.FindBlockPair(startLine, SearchType, Direction)
    End If
End Function

' [엔진 서비스] IF/WHILE 조건절 판단 (v2: 좌변 P1 | 연산자 P2 | 우변 P3)
'   v1: 해석된 P1, op: P2(연산자, RAW), v2: 해석된 P3
' @desc 좌변/연산자/우변 받아 조건 참거짓 반환
Public Function CheckIfSyntax(ByVal v1 As String, ByVal op As String, ByVal v2 As String) As Boolean
    On Error GoTo ErrHandler
    CheckIfSyntax = modCondition.EvaluateCondition(v1, UCase(Trim(op)), v2)

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modEngine.CheckIfSyntax", Err.Description
    Resume CleanUp
End Function

' [호환 엔진 서비스] ActiveProgram 한 줄의 정규화된 명령어 반환
' @desc 줄번호 받아 들여쓰기 제거된 대문자 명령어 반환
Public Function CmdAt(ByVal R As Long) As String
    If Not m_program Is Nothing Then CmdAt = m_program.CmdAt(R)
End Function

' [엔진 서비스] SELECT 인덱스 조회 -> 일치 CASE 다음 줄 / 폴백 / END_SELECT 반환
' @desc 현재 SELECT 줄과 대상값 받아 점프 줄번호 반환, 일치/폴백 없으면 END_SELECT 또는 CMD_ERROR
Public Function ResolveSelectTarget(ByVal R As Long, ByVal selVal As String) As Long
    On Error GoTo ErrHandler
    If m_program Is Nothing Then
        m_lastRuntimeError = "SELECT 라인 " & R & ": 실행할 ProgramImage가 없습니다."
        ResolveSelectTarget = CMD_ERROR
        GoTo CleanUp
    End If

    Dim block As clsSelectBlock
    Set block = m_program.SelectBlock(R)
    If block Is Nothing Then
        m_lastRuntimeError = "SELECT 라인 " & R & ": SELECT 블록 인덱스를 찾지 못했습니다."
        ResolveSelectTarget = CMD_ERROR
        GoTo CleanUp
    End If
    If block.EndSelectLine <= 0 Then
        m_lastRuntimeError = "SELECT 라인 " & R & ": END_SELECT 짝이 없습니다."
        ResolveSelectTarget = CMD_ERROR
        GoTo CleanUp
    End If

    Dim caseVal As String
    Dim fallbackLine As Long
    fallbackLine = 0

    Dim i As Long
    For i = 1 To block.CaseCount
        Dim caseLine As Long
        caseLine = block.CaseLineAt(i)
        caseVal = modTokenResolver.ResolveParam(m_program.InstructionAt(caseLine).p1)

        ' Case ELSE (또는 빈 값) = 폴백. 여러 개면 마지막 폴백이 이긴다.
        If UCase$(Trim$(caseVal)) = "ELSE" Or Len(caseVal) = 0 Then
            fallbackLine = caseLine
        ElseIf modCondition.EvaluateCondition(selVal, "EQ", caseVal) Then
            ResolveSelectTarget = caseLine + 1
            GoTo CleanUp
        End If
    Next i

    If fallbackLine > 0 Then
        ResolveSelectTarget = fallbackLine + 1
    Else
        ResolveSelectTarget = block.EndSelectLine
    End If

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modEngine.ResolveSelectTarget", Err.Description
    Resume CleanUp
End Function

