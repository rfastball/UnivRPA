Attribute VB_Name = "modCmdRegistry"
' @desc 명령어 메타데이터 레지스트리 빌드 및 GetCmdDef 단일 게이트 제공 담당
' @depends CmdDef, ParamSpec
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

'==============================================================================
' [Layer 0] modCmdRegistry: 명령어 정의 단일 출처(SSOT)
'==============================================================================
' 역할: 명령어 1개의 모든 지식(검증/디스패치/서식/블록판정/decode)을 CmdDef 레코드로
'       집약하고, 모든 조회를 GetCmdDef 게이트 하나로 강제한다.
' 규율: m_reg(Dictionary)에 직접 접근하는 코드는 없다. 오직 GetCmdDef 경유.
'==============================================================================

' [상수 enum] 세 축은 서로 독립 (Category != ColorGroup != PaintScope)
Public Enum eCategory
    FLOW = 1
    ACTION = 2
    VISION = 3
    Memory = 4
    SYSTEM = 5
End Enum

Public Enum eBlockRole
    NOBLOCK = 0
    STARTER = 1
    ENDER = 2
    MIDBLOCK = 3
End Enum

Public Enum eColorGroup
    clrFLOW = 1
    clrACTION = 2
    clrVISION = 3
    clrSYSTEM = 4
End Enum

Public Enum ePaintScope
    FULLROW = 1     ' 줄 전체(B~E) 칠
    cmdCell = 2     ' B열만 칠 + 중첩 시 C~E 연녹 배경
End Enum

Public Enum eResolveMode
    RESOLVE = 0     ' 토큰 치환 적용
    raw = 1         ' 원본 유지 (lvalue 주소 / 연산자)
End Enum

Public Enum eDataType
    DT_STR = 0
    DT_NUM = 1
End Enum

' [RAW 슬롯 시맨틱] 같은 RAW라도 검증 로직이 다르다 - 어떤 종류인지 레지스트리가 선언 (C2).
'   소비처(modDevTools 포맷터, Analysis 점프 검증)는 명령어명 하드코딩 대신 이 태그를 질의한다.
Public Enum eRawKind
    RK_NONE = 0         ' RESOLVE 슬롯 (시맨틱 태그 없음)
    RK_TEXT = 1         ' 일반 RAW 텍스트 (LABEL명 등 - 토큰만 거부)
    RK_CONDITION = 2    ' 조건 연산자 (modCondition이 합법성 소유)
    RK_LVALUE = 3       ' 쓰기 주소 (lvalue 합법성 검증)
    RK_JUMP_TARGET = 4  ' 점프 대상 (정적 라벨 계약: 숫자/토큰 거부)
End Enum

' [디스패치 반환 코드] 모든 모듈/핸들러 클래스가 공유 (CMD_NOT_HANDLED는 폐기됨)
Public Const CMD_NEXT_LINE As Long = 0   ' 성공 -> 다음 줄
Public Const CMD_ERROR As Long = -1      ' 실패 -> 중단
Public Const CMD_HALT As Long = -2       ' 의도된 종료

' [레지스트리 저장소] 정식명(String) -> CmdDef
Private m_reg As Object

' [핸들러 저장소] Category(eCategory) -> ICommand 인스턴스 (1단 디스패치)
Private m_handlers As Object

'==============================================================================
' 1. Gate (단일 통로)
'==============================================================================

' @desc 명령어 문자열 받아 정규화 후 CmdDef 반환, 미등록/주석이면 Nothing 반환
Public Function GetCmdDef(ByVal rawCmd As String) As CmdDef
    EnsureRegistry

    ' 정규화: 선행 들여쓰기 마커 제거 -> Trim -> UCase
    Dim cmd As String
    cmd = NormalizeCommand(rawCmd)

    ' 주석(#) 또는 빈 줄은 명령어 아님
    If Len(cmd) = 0 Then Exit Function
    If Left$(cmd, 1) = "#" Then Exit Function

    If m_reg.exists(cmd) Then Set GetCmdDef = m_reg(cmd)
End Function

' @desc 명령어 문자열의 단일 정규화 규칙: 선행 마커 제거 -> Trim -> UCase
Public Function NormalizeCommand(ByVal rawCmd As Variant) As String
    On Error GoTo SafeExit
    If IsError(rawCmd) Or IsNull(rawCmd) Or IsEmpty(rawCmd) Then Exit Function
    ' 선행 들여쓰기 마커만 벗긴다 - 내부/후행 ">"는 명령 문자열의 일부로 보존
    NormalizeCommand = UCase$(StripIndentMarkers(CStr(rawCmd)))
SafeExit:
End Function

' @desc 선행 들여쓰기 마커(공백/탭/">") 제거한 나머지 반환, 전부 마커면 빈 문자열 반환
Public Function StripIndentMarkers(ByVal text As String) As String
    ' 마커 문자 집합은 modScriptAssist.CommentMarkerPos와 동일해야 한다
    Dim idx As Long
    For idx = 1 To Len(text)
        Select Case Mid$(text, idx, 1)
            Case " ", vbTab, ">"
                ' 시각적 들여쓰기 - 계속 스캔
            Case Else
                Exit For
        End Select
    Next idx
    StripIndentMarkers = Trim$(Mid$(text, idx))
End Function

' @desc 레지스트리가 비어있으면 1회 빌드함 (지연 초기화)
Public Sub EnsureRegistry()
    If Not m_reg Is Nothing Then Exit Sub
    Set m_reg = CreateObject("Scripting.Dictionary")
    BuildRegistry
End Sub

' @desc 레지스트리 전체를 새로 빌드함 (테스트/재적재용 강제 리셋)
Public Sub RebuildRegistry()
    Set m_reg = Nothing
    EnsureRegistry
End Sub

' @desc 등록된 모든 명령어 정식명 배열 반환 (역검증 순회용)
Public Function AllCommands() As Variant
    EnsureRegistry
    AllCommands = m_reg.keys
End Function

' @desc 카테고리 받아 담당 ICommand 핸들러 인스턴스 반환 (1단 디스패치)
Public Function handler(ByVal cat As eCategory) As ICommand
    EnsureHandlers
    Set handler = m_handlers(cat)
End Function

' @desc 카테고리별 핸들러 인스턴스를 1회 생성 배선함 (SYSTEM/MEMORY 1:1, 총 5클래스)
Private Sub EnsureHandlers()
    If Not m_handlers Is Nothing Then Exit Sub
    Set m_handlers = CreateObject("Scripting.Dictionary")
    Set m_handlers(FLOW) = New clsCommandsFlow
    Set m_handlers(ACTION) = New clsCommandsAction
    Set m_handlers(VISION) = New clsCommandsVision
    Set m_handlers(Memory) = New clsCommandsMemory
    Set m_handlers(SYSTEM) = New clsCommandsSystem
End Sub

'==============================================================================
' 2. 한 줄 등록 헬퍼
'==============================================================================

' [Reg] New CmdDef + Init 후 Dictionary 등록
Private Sub Reg(ByVal canon As String, ByVal cat As eCategory, ByVal role As eBlockRole, _
                ByVal group As String, ByVal color As eColorGroup, ByVal scope As ePaintScope, _
                Optional ByVal sp1 As ParamSpec, _
                Optional ByVal sp2 As ParamSpec, _
                Optional ByVal sp3 As ParamSpec)
    Dim d As CmdDef
    Set d = New CmdDef
    d.Init canon, cat, role, group, color, scope, sp1, sp2, sp3
    Set m_reg(canon) = d
End Sub

' [P] New ParamSpec + Init 후 반환 (decode 슬롯 정의)
'   kind 생략 시 RAW 슬롯은 RK_TEXT, RESOLVE 슬롯은 RK_NONE (Init이 불변식 강제)
Private Function P(ByVal Meaning As String, ByVal dtype As eDataType, _
                   ByVal rmode As eResolveMode, ByVal Required As Boolean, _
                   Optional ByVal kind As eRawKind = RK_NONE) As ParamSpec
    Dim s As ParamSpec
    Set s = New ParamSpec
    Set P = s.Init(Meaning, dtype, rmode, Required, kind)
End Function

'==============================================================================
' 3. 레지스트리 적재 (현행 30개 + INC/DEC/MOVE + FOREACH_ROW + ELSE_IF = 36개)
'   초기값: 설계문서 부록 A / A.1 / A.3
'   RAW 슬롯 4종 (RawKind로 선언): 연산자(IF/ELSE_IF/WHILE P2), lvalue(SET/INC/DEC P1),
'   점프대상(GOTO/CALL P3), 일반 텍스트(LABEL P3). 그 외 RESOLVE.
'==============================================================================
Private Sub BuildRegistry()

    '--- [FLOW] 19개 ------------------------------------------------------------
    Reg "IF", FLOW, STARTER, "IF", clrFLOW, FULLROW, _
        P("좌변", DT_STR, RESOLVE, True), P("연산자", DT_STR, raw, True, RK_CONDITION), P("우변", DT_STR, RESOLVE, True)
    Reg "ELSE_IF", FLOW, MIDBLOCK, "IF", clrFLOW, FULLROW, _
        P("좌변", DT_STR, RESOLVE, True), P("연산자", DT_STR, raw, True, RK_CONDITION), P("우변", DT_STR, RESOLVE, True)
    Reg "ELSE", FLOW, MIDBLOCK, "IF", clrFLOW, FULLROW
    Reg "END_IF", FLOW, ENDER, "IF", clrFLOW, FULLROW
    Reg "WHILE", FLOW, STARTER, "WHILE", clrFLOW, FULLROW, _
        P("좌변", DT_STR, RESOLVE, True), P("연산자", DT_STR, raw, True, RK_CONDITION), P("우변", DT_STR, RESOLVE, True)
    Reg "END_WHILE", FLOW, ENDER, "WHILE", clrFLOW, FULLROW
    Reg "REPEAT", FLOW, STARTER, "REPEAT", clrFLOW, FULLROW, _
        P("반복횟수(0/빈값=무한)", DT_NUM, RESOLVE, False)
    Reg "END_REPEAT", FLOW, ENDER, "REPEAT", clrFLOW, FULLROW
    Reg "FOREACH_ROW", FLOW, STARTER, "FOREACH", clrFLOW, FULLROW, _
        P("워크북 이름/경로", DT_STR, RESOLVE, True), P("시트 또는 시트!주소", DT_STR, RESOLVE, True), P("순회모드(ALL_ROWS/VISIBLE_ROWS)", DT_STR, RESOLVE, False)
    Reg "END_FOREACH", FLOW, ENDER, "FOREACH", clrFLOW, FULLROW
    Reg "SELECT", FLOW, STARTER, "SELECT", clrFLOW, FULLROW, _
        P("대상값", DT_STR, RESOLVE, True)
    Reg "CASE", FLOW, MIDBLOCK, "SELECT", clrFLOW, FULLROW, _
        P("비교값", DT_STR, RESOLVE, False)
    Reg "END_SELECT", FLOW, ENDER, "SELECT", clrFLOW, FULLROW
    Reg "BREAK", FLOW, NOBLOCK, "", clrFLOW, FULLROW
    Reg "CONTINUE", FLOW, NOBLOCK, "", clrFLOW, FULLROW
    ' GOTO/CALL 대상은 정적 라벨만 허용(숫자/토큰 거부) - P3를 raw로 고정해
    '   정적 검증(정적 라벨 계약)과 런타임 디코딩을 일치시킨다
    Reg "GOTO", FLOW, NOBLOCK, "", clrFLOW, FULLROW, _
        sp3:=P("라벨", DT_STR, raw, True, RK_JUMP_TARGET)
    Reg "CALL", FLOW, NOBLOCK, "", clrFLOW, FULLROW, _
        sp3:=P("라벨", DT_STR, raw, True, RK_JUMP_TARGET)
    Reg "RETURN", FLOW, NOBLOCK, "", clrFLOW, FULLROW
    Reg "LABEL", FLOW, NOBLOCK, "", clrFLOW, FULLROW, _
        sp3:=P("라벨명", DT_STR, raw, True)

    '--- [ACTION] 6개 -----------------------------------------------------------
    Reg "MOVE", ACTION, NOBLOCK, "", clrACTION, cmdCell, _
        P("x", DT_NUM, RESOLVE, True), P("y", DT_NUM, RESOLVE, True), P("사후대기ms", DT_NUM, RESOLVE, False)
    Reg "CLICK", ACTION, NOBLOCK, "", clrACTION, cmdCell, _
        P("x", DT_NUM, RESOLVE, True), P("y", DT_NUM, RESOLVE, True), P("사후대기ms", DT_NUM, RESOLVE, False)
    Reg "RCLICK", ACTION, NOBLOCK, "", clrACTION, cmdCell, _
        P("x", DT_NUM, RESOLVE, True), P("y", DT_NUM, RESOLVE, True), P("사후대기ms", DT_NUM, RESOLVE, False)
    Reg "KEY", ACTION, NOBLOCK, "", clrACTION, cmdCell, _
        P("키코드/이름", DT_STR, RESOLVE, True), Nothing, P("사후대기ms", DT_NUM, RESOLVE, False)
    Reg "COMBO", ACTION, NOBLOCK, "", clrACTION, cmdCell, _
        P("key1", DT_STR, RESOLVE, True), P("key2", DT_STR, RESOLVE, True), P("사후대기ms", DT_NUM, RESOLVE, False)
    Reg "TYPE", ACTION, NOBLOCK, "", clrACTION, cmdCell, _
        P("딜레이ms", DT_NUM, RESOLVE, False), Nothing, P("텍스트", DT_STR, RESOLVE, True)

    '--- [MEMORY] lvalue-writer 3개 (SET 현행 + INC/DEC 신규) --------------------
    '   P1은 RAW(좌변=주소). ColorGroup은 현행 SET을 따라 clrACTION(연청).
    Reg "SET", Memory, NOBLOCK, "", clrACTION, cmdCell, _
        P("좌변/주소", DT_STR, raw, True, RK_LVALUE), Nothing, P("값", DT_STR, RESOLVE, True)
    Reg "INC", Memory, NOBLOCK, "", clrACTION, cmdCell, _
        P("좌변/주소", DT_STR, raw, True, RK_LVALUE), Nothing, P("증분(기본1)", DT_NUM, RESOLVE, False)
    Reg "DEC", Memory, NOBLOCK, "", clrACTION, cmdCell, _
        P("좌변/주소", DT_STR, raw, True, RK_LVALUE), Nothing, P("감분(기본1)", DT_NUM, RESOLVE, False)

    '--- [SYSTEM] 4개 (FOCUS/SLEEP는 clrACTION, NOTIFY/EXIT는 clrSYSTEM) ---------
    Reg "FOCUS", SYSTEM, NOBLOCK, "", clrACTION, cmdCell, _
        sp3:=P("창제목", DT_STR, RESOLVE, True)
    Reg "SLEEP", SYSTEM, NOBLOCK, "", clrACTION, cmdCell, _
        sp3:=P("대기ms", DT_NUM, RESOLVE, True)
    Reg "NOTIFY", SYSTEM, NOBLOCK, "", clrSYSTEM, FULLROW, _
        P("표시시간초(빈값=2, 0=확인대기)", DT_NUM, RESOLVE, False), Nothing, _
        P("메시지", DT_STR, RESOLVE, True)
    Reg "EXIT", SYSTEM, NOBLOCK, "", clrSYSTEM, FULLROW

    '--- [VISION] 4개 -----------------------------------------------------------
    Reg "WAIT_APPEAR", VISION, NOBLOCK, "", clrVISION, cmdCell, _
        P("x", DT_NUM, RESOLVE, True), P("y", DT_NUM, RESOLVE, True), P("hexColor", DT_STR, RESOLVE, True)
    Reg "WAIT_VANISH", VISION, NOBLOCK, "", clrVISION, cmdCell, _
        P("x", DT_NUM, RESOLVE, True), P("y", DT_NUM, RESOLVE, True), P("hexColor", DT_STR, RESOLVE, True)
    Reg "CLICK_SYNC", VISION, NOBLOCK, "", clrVISION, cmdCell, _
        P("x", DT_NUM, RESOLVE, True), P("y", DT_NUM, RESOLVE, True)
    Reg "ASSERT", VISION, NOBLOCK, "", clrVISION, cmdCell, _
        P("x", DT_NUM, RESOLVE, True), P("y", DT_NUM, RESOLVE, True), P("조건(!=금지색)", DT_STR, RESOLVE, True)

End Sub
