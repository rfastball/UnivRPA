VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

' ==============================================================================
' Class Name: clsTableSnapshot
' Description: 엑셀 시트 데이터를 메모리화하여 관리하는 고속 데이터 컨테이너
' Features:
'   읽기 전용 2D 스냅샷과 헤더 이름 기반 셀 조회.
' ==============================================================================

' [저장소]
Private m_vData As Variant          ' 전체 데이터 (2D Array, 1-based or 0-based)
Private m_dicHeaders As Object      ' 컬럼명 -> 열 번호 매핑 (Key: String, Value: Long)

' [메타 데이터]
Private m_lRowCount As Long
Private m_lColCount As Long
Private m_lHeaderRowIndex As Long   ' 헤더가 위치한 행 번호 (기본 1)

' [초기화]
Private Sub Class_Initialize()
    Set m_dicHeaders = CreateObject("Scripting.Dictionary")
    m_lHeaderRowIndex = 1
End Sub

Private Sub Class_Terminate()
    Set m_dicHeaders = Nothing
    m_vData = Empty
End Sub

' ==============================================================================
' [1] 데이터 로드 (Load)
' ==============================================================================

' ==============================================================================
' [1-1] 시트 데이터 로드 (Smart Range Detection 적용)
' 사용법 1: t.LoadFromSheet Sheets("Data")        -> 시트 전체 데이터 자동 탐색 (추천)
' 사용법 2: t.LoadFromSheet Range("A1")           -> A1부터 끝까지 자동 탐색 (추천)
' 사용법 3: t.LoadFromSheet Range("A1:F500")      -> 지정된 범위만 강제 로드
' ==============================================================================
Public Sub LoadFromSheet(ByVal vSource As Variant, Optional ByVal visibleOnly As Boolean = False)
    Dim rngTarget As Range
    
    ' 1. 입력 타입 확인 및 범위 확정
    If TypeOf vSource Is Worksheet Then
        ' [Case A] 시트 객체가 넘어온 경우 -> 데이터 있는 전체 범위 자동 탐색
        If visibleOnly Then Set rngTarget = GetAutoFilterRange(vSource)
        If rngTarget Is Nothing Then Set rngTarget = GetRealUsedRange(vSource)
        
    ElseIf TypeOf vSource Is Range Then
        ' [Case B] Range 객체가 넘어온 경우
        If vSource.Cells.count = 1 Then
            ' [Case B-1] 단일 셀(예: A1)만 넘김 -> 여기서부터 끝까지 자동 탐색
            Set rngTarget = GetRealUsedRange(vSource.Worksheet, vSource)
        Else
            ' [Case B-2] 구체적 범위(예: A1:E100)를 넘김
            Set rngTarget = ClampToUsedRange(vSource)
        End If
    Else
        Err.Raise 9999, , "LoadFromSheet: Worksheet 또는 Range 객체만 지원합니다."
    End If
    
    ' 2. 데이터 유효성 검사
    If rngTarget Is Nothing Then
        m_lRowCount = 0
        Exit Sub
    End If
    
    Dim vData As Variant
    vData = SnapshotVisibleText(rngTarget, visibleOnly)
    
    ' 4. 배열 로드 실행
    LoadFromArray vData
End Sub

Private Function GetAutoFilterRange(ByVal ws As Worksheet) As Range
    On Error GoTo SafeExit
    If ws.AutoFilterMode Then Set GetAutoFilterRange = ws.AutoFilter.Range
SafeExit:
End Function

Private Function ClampToUsedRange(ByVal rngSource As Range) As Range
    Dim rngUsed As Range
    Set rngUsed = rngSource.Worksheet.UsedRange

    If rngUsed Is Nothing Then
        Set ClampToUsedRange = rngSource
        Exit Function
    End If

    Dim rngIntersect As Range
    Set rngIntersect = Intersect(rngSource, rngUsed)
    If rngIntersect Is Nothing Then
        Set ClampToUsedRange = rngSource.Cells(1, 1)
    Else
        Set ClampToUsedRange = rngIntersect
    End If
End Function

' [E] 셀 단위 COM 왕복 축소: 값은 .Value 일괄 스냅샷(1회)에서 읽고, 셀 단위 COM은
'   표시 텍스트(.Text)와 드문 오버플로 보정(.NumberFormat)만 남긴다. 표시 텍스트가
'   진실이라는 계약(날짜/서식 열)은 유지 - .Value2 전면 전환은 서식 열 출력이 바뀜.
Private Function SnapshotVisibleText(ByVal rngTarget As Range, ByVal visibleOnly As Boolean) As Variant
    Dim result() As Variant
    Dim R As Long, c As Long
    Dim targetRow As Long
    Dim includedRows As Long
    Dim cell As Range
    Dim textValue As String

    Dim rowCount As Long, colCount As Long
    rowCount = rngTarget.rows.count
    colCount = rngTarget.Columns.count

    ' 값 일괄 스냅샷 (.Value는 Date/Currency 타입 보존 - 폴백 포맷 의미 유지)
    Dim vValues As Variant
    If rngTarget.Cells.count = 1 Then
        ReDim vValues(1 To 1, 1 To 1)
        vValues(1, 1) = rngTarget.value
    Else
        vValues = rngTarget.value
    End If

    ' 포함 행 판정 1회 캐시 (visibleOnly의 Hidden/Height COM 조회가 2패스 중복되지 않게)
    Dim rowIncluded() As Boolean
    ReDim rowIncluded(1 To rowCount)
    rowIncluded(1) = True
    includedRows = 1
    For R = 2 To rowCount
        If visibleOnly Then
            rowIncluded(R) = IsWorksheetRowVisible(rngTarget.rows(R))
        Else
            rowIncluded(R) = True
        End If
        If rowIncluded(R) Then includedRows = includedRows + 1
    Next R

    ReDim result(1 To includedRows, 1 To colCount)

    For R = 1 To rowCount
        If Not rowIncluded(R) Then GoTo NextSourceRow

        targetRow = targetRow + 1
        For c = 1 To colCount
            If R = 1 Then
                textValue = HeaderText(vValues(1, c))
            Else
                Set cell = rngTarget.Cells(R, c)
                textValue = cell.text
                If IsOverflowText(textValue) Or (Len(textValue) = 0 And CellHasValue(vValues(R, c))) Then
                    textValue = modFilterUtils.GetSafeFormattedText(vValues(R, c), cell.NumberFormat)
                End If
            End If
            result(targetRow, c) = textValue
        Next c
NextSourceRow:
    Next R

    SnapshotVisibleText = result
End Function

Private Function IsWorksheetRowVisible(ByVal rngRow As Range) As Boolean
    On Error GoTo SafeExit

    If rngRow.EntireRow.Hidden Then Exit Function
    If rngRow.Height <= 0 Then Exit Function

    IsWorksheetRowVisible = True
SafeExit:
End Function

Private Function HeaderText(ByVal vVal As Variant) As String
    If IsError(vVal) Or IsNull(vVal) Or IsEmpty(vVal) Then Exit Function
    HeaderText = CStr(vVal)
End Function

Private Function CellHasValue(ByVal vVal As Variant) As Boolean
    If IsError(vVal) Or IsNull(vVal) Or IsEmpty(vVal) Then Exit Function
    CellHasValue = (Len(CStr(vVal)) > 0)
End Function

Private Function IsOverflowText(ByVal textValue As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(textValue) = 0 Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch <> "#" Then Exit Function
    Next i

    IsOverflowText = True
End Function

' ==============================================================================
' [Best Practice] Find 메서드를 이용한 고속 범위 탐색
' ==============================================================================
Private Function GetRealUsedRange(ByVal ws As Worksheet, Optional ByVal rngStart As Range = Nothing) As Range
    Dim lLastRow As Long, lLastCol As Long
    Dim lStartRow As Long, lStartCol As Long
    Dim rngFindRow As Range, rngFindCol As Range
    
    ' 1. 시작점 설정 (기본 A1)
    If rngStart Is Nothing Then
        lStartRow = 1: lStartCol = 1
    Else
        lStartRow = rngStart.Row: lStartCol = rngStart.Column
    End If
    
    ' 2. 데이터가 있는 '진짜' 마지막 행 찾기 (뒤에서부터 검색)
    ' * xlFormulas 옵션: 숨겨진 수식이나 값 모두 감지
    Set rngFindRow = ws.Cells.Find(What:="*", _
                                   After:=ws.Cells(1, 1), _
                                   LookIn:=xlFormulas, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlPrevious)
                                   
    ' 3. 데이터가 있는 '진짜' 마지막 열 찾기
    Set rngFindCol = ws.Cells.Find(What:="*", _
                                   After:=ws.Cells(1, 1), _
                                   LookIn:=xlFormulas, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   SearchDirection:=xlPrevious)
    
    ' 4. 데이터가 아예 없는 빈 시트일 경우 방어
    If rngFindRow Is Nothing Or rngFindCol Is Nothing Then
        ' 데이터 없음 -> 시작 셀 하나만 반환 (에러 방지)
        Set GetRealUsedRange = ws.Cells(lStartRow, lStartCol)
        Exit Function
    End If
    
    lLastRow = rngFindRow.Row
    lLastCol = rngFindCol.Column
    
    ' 시작점보다 데이터가 앞쪽에 있으면 보정
    If lLastRow < lStartRow Then lLastRow = lStartRow
    If lLastCol < lStartCol Then lLastCol = lStartCol
    
    ' 5. 사각형 범위 확정 (중간에 구멍이 뚫려도 직사각형으로 잡음)
    Set GetRealUsedRange = ws.Range(ws.Cells(lStartRow, lStartCol), ws.Cells(lLastRow, lLastCol))
End Function

' 1-2. 배열에서 직접 로드 (데이터 정규화 포함 버전)
Public Sub LoadFromArray(ByVal vArr As Variant, Optional ByVal bHasHeader As Boolean = True)
    Dim i As Long, j As Long
    Dim sColName As String
    
    Clear ' 기존 데이터 초기화
    
    m_vData = vArr
    m_lRowCount = UBound(m_vData, 1)
    m_lColCount = UBound(m_vData, 2)
    
    ' [전처리] 전체 데이터 영역 초고속 트리밍
    ' * 문자열 타입(vbString)인 경우에만 Trim$ 수행하여 속도 최적화
    For i = LBound(m_vData, 1) To UBound(m_vData, 1)
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            If VarType(m_vData(i, j)) = vbString Then
                m_vData(i, j) = Trim$(m_vData(i, j))
            End If
        Next j
    Next i
    
    ' 헤더 매핑 생성
    If bHasHeader Then
        Dim lHeadRow As Long: lHeadRow = LBound(m_vData, 1)
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            ' 헤더 자체도 다시 한번 Trim (이미 위에서 처리되었지만 명시적 유지)
            sColName = CStr(m_vData(lHeadRow, j))
            If Len(sColName) > 0 Then
                ' 중복 헤더는 조용히 덮어쓰지 않는다: 스크립트가 "어느 열을 보는지"
                '   잃으면 자동화가 위험하므로 자동 suffix 대신 하드 에러
                If m_dicHeaders.exists(UCase(sColName)) Then
                    Err.Raise 5, "clsTableSnapshot.LoadFromArray", "중복 헤더: " & sColName
                End If
                m_dicHeaders(UCase(sColName)) = j
            End If
        Next j
        m_lHeaderRowIndex = lHeadRow
    Else
        ' 헤더가 없는 경우
        For j = LBound(m_vData, 2) To UBound(m_vData, 2)
            m_dicHeaders("COLUMN" & j) = j
        Next j
        m_lHeaderRowIndex = LBound(m_vData, 1) - 1
    End If
End Sub



' 초기화
Public Sub Clear()
    m_vData = Empty
    m_dicHeaders.RemoveAll
    m_lRowCount = 0
    m_lColCount = 0
End Sub

' ==============================================================================
' [2] 데이터 조회 (Select/Get)
' ==============================================================================

' 2-1. 특정 셀 데이터 가져오기 (열 이름 이용)
Public Function GetValue(ByVal lRow As Long, ByVal sColName As String) As Variant
    Dim lCol As Long
    sColName = UCase(Trim(sColName))
    
    If m_dicHeaders.exists(sColName) Then
        lCol = m_dicHeaders(sColName)
        ' 배열 범위 방어
        If lRow >= LBound(m_vData, 1) And lRow <= m_lRowCount Then
            GetValue = m_vData(lRow, lCol)
        End If
    End If
End Function

' ==============================================================================
' [4] 속성 (Properties)
' ==============================================================================

Public Property Get RowCount() As Long
    RowCount = m_lRowCount
End Property

Public Property Get FirstDataRow() As Long
    FirstDataRow = m_lHeaderRowIndex + 1
End Property

Public Property Get ColCount() As Long
    ColCount = m_lColCount
End Property

Public Property Get headers() As Object
    Set headers = m_dicHeaders
End Property
