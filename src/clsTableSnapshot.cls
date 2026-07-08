VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==============================================================================
' Class Name: clsTableSnapshot
' Description: 엑셀 시트 데이터를 메모리화하여 관리하는 고속 데이터 컨테이너
' Features:
'   1. Lazy Indexing: 검색이 필요한 시점에 인덱스 생성 (메모리 절약)
'   2. Row-ID Mapping: 중복 데이터 허용 (Key -> Collection of RowIDs)
'   3. Direct Access: 2D 배열 기반의 초고속 데이터 접근
' ==============================================================================

' [저장소]
Private m_vData As Variant          ' 전체 데이터 (2D Array, 1-based or 0-based)
Private m_dicHeaders As Object      ' 컬럼명 -> 열 번호 매핑 (Key: String, Value: Long)
Private m_dicIndices As Object      ' 생성된 인덱스 캐시 (Key: ColName, Value: Dictionary)

' [메타 데이터]
Private m_lRowCount As Long
Private m_lColCount As Long
Private m_lHeaderRowIndex As Long   ' 헤더가 위치한 행 번호 (기본 1)

' [초기화]
Private Sub Class_Initialize()
    Set m_dicHeaders = CreateObject("Scripting.Dictionary")
    Set m_dicIndices = CreateObject("Scripting.Dictionary")
    m_lHeaderRowIndex = 1
End Sub

Private Sub Class_Terminate()
    Set m_dicHeaders = Nothing
    Set m_dicIndices = Nothing
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

Private Function SnapshotVisibleText(ByVal rngTarget As Range, ByVal visibleOnly As Boolean) As Variant
    Dim result() As Variant
    Dim R As Long, c As Long
    Dim targetRow As Long
    Dim includedRows As Long
    Dim cell As Range
    Dim textValue As String

    includedRows = 1
    For R = 2 To rngTarget.rows.count
        If IsSnapshotRowIncluded(R, rngTarget.rows(R), visibleOnly) Then includedRows = includedRows + 1
    Next R

    ReDim result(1 To includedRows, 1 To rngTarget.Columns.count)

    For R = 1 To rngTarget.rows.count
        If Not IsSnapshotRowIncluded(R, rngTarget.rows(R), visibleOnly) Then GoTo NextSourceRow

        targetRow = targetRow + 1
        For c = 1 To rngTarget.Columns.count
            Set cell = rngTarget.Cells(R, c)
            If R = 1 Then
                textValue = HeaderText(cell)
            Else
                textValue = cell.text
                If IsOverflowText(textValue) Or (Len(textValue) = 0 And CellHasValue(cell)) Then
                    textValue = GetSafeFormattedText(cell.value, cell.NumberFormat)
                End If
            End If
            result(targetRow, c) = textValue
        Next c
NextSourceRow:
    Next R

    SnapshotVisibleText = result
End Function

Private Function IsSnapshotRowIncluded(ByVal sourceRowIndex As Long, ByVal rngRow As Range, ByVal visibleOnly As Boolean) As Boolean
    If sourceRowIndex = 1 Then
        IsSnapshotRowIncluded = True
        Exit Function
    End If

    If Not visibleOnly Then
        IsSnapshotRowIncluded = True
    Else
        IsSnapshotRowIncluded = IsWorksheetRowVisible(rngRow)
    End If
End Function

Private Function IsWorksheetRowVisible(ByVal rngRow As Range) As Boolean
    On Error GoTo SafeExit

    If rngRow.EntireRow.Hidden Then Exit Function
    If rngRow.Height <= 0 Then Exit Function

    IsWorksheetRowVisible = True
SafeExit:
End Function

Private Function HeaderText(ByVal cell As Range) As String
    If IsError(cell.value) Or IsNull(cell.value) Or IsEmpty(cell.value) Then Exit Function
    HeaderText = CStr(cell.value)
End Function

Private Function CellHasValue(ByVal cell As Range) As Boolean
    If IsError(cell.value) Or IsNull(cell.value) Or IsEmpty(cell.value) Then Exit Function
    CellHasValue = (Len(CStr(cell.value)) > 0)
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

Private Function GetSafeFormattedText(ByVal vVal As Variant, ByVal sFmt As String) As String
    Dim sResult As String
    Dim bSuccess As Boolean

    If IsError(vVal) Then
        GetSafeFormattedText = CStr(vVal)
        Exit Function
    End If
    If IsEmpty(vVal) Or IsNull(vVal) Then
        GetSafeFormattedText = ""
        Exit Function
    End If

    If sFmt <> "" Then
        On Error Resume Next
        sResult = Application.WorksheetFunction.text(vVal, sFmt)
        bSuccess = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If

    If Not bSuccess And sFmt <> "" Then
        On Error Resume Next
        sResult = Format(vVal, sFmt)
        bSuccess = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If

    If Not bSuccess Then
        If IsDate(vVal) Then
            sResult = Format(vVal, "yyyy-mm-dd")
        Else
            sResult = CStr(vVal)
        End If
    End If

    GetSafeFormattedText = sResult
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
    m_dicIndices.RemoveAll
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

' 2-2. 특정 행의 전체 레코드 가져오기 (배열 반환)
Public Function GetRecord(ByVal lRow As Long) As Variant
    Dim vRow() As Variant
    Dim j As Long, lStartCol As Long, lEndCol As Long
    
    If lRow < LBound(m_vData, 1) Or lRow > m_lRowCount Then Exit Function
    
    lStartCol = LBound(m_vData, 2)
    lEndCol = UBound(m_vData, 2)
    
    ReDim vRow(lStartCol To lEndCol)
    
    For j = lStartCol To lEndCol
        vRow(j) = m_vData(lRow, j)
    Next j
    
    GetRecord = vRow
End Function

' ==============================================================================
' [3] 인덱싱 및 검색 (Core Engine)
' ==============================================================================

' 3-1a. 인덱스 생성 및 획득 (By 컬럼명)
Public Function GetIndex(ByVal sColName As String) As Object
    Dim lCol As Long
    Dim dicResult As Object
    
    sColName = UCase(Trim(sColName))
    
    ' 1. 캐시 확인 (키: 컬럼명)
    If m_dicIndices.exists(sColName) Then
        Set GetIndex = m_dicIndices(sColName)
        Exit Function
    End If
    
    ' 2. 컬럼 인덱스 찾기
    If Not m_dicHeaders.exists(sColName) Then
        Exit Function ' 컬럼 없음
    End If
    lCol = m_dicHeaders(sColName)
    
    ' 3. 숫자 인덱스 기반으로 실제 인덱스 생성/조회
    Set dicResult = GetIndexByIndex(lCol) ' This will build and cache it for lCol if needed
    
    ' 4. 컬럼명 기준으로도 캐시 등록
    If Not dicResult Is Nothing Then
        Set m_dicIndices(sColName) = dicResult
    End If
    Set GetIndex = dicResult
End Function

' 3-1b. 인덱스 생성 및 획득 (By 숫자 인덱스) - v2.5.1 최적화
Public Function GetIndexByIndex(ByVal lCol As Long) As Object
    Dim i As Long
    Dim vVal As Variant
    Dim sKey As String
    Dim dicIdx As Object
    Dim colRows As Collection

    ' 1. 캐시 확인 (키: 숫자 인덱스)
    If m_dicIndices.exists(lCol) Then
        Set GetIndexByIndex = m_dicIndices(lCol)
        Exit Function
    End If

    ' 2. 컬럼 유효성 검사 (배열 범위)
    If lCol < LBound(m_vData, 2) Or lCol > UBound(m_vData, 2) Then
        Exit Function
    End If

    ' 3. 인덱스 빌드 시작
    Set dicIdx = CreateObject("Scripting.Dictionary")
    dicIdx.CompareMode = vbTextCompare ' 대소문자 무시

    For i = (m_lHeaderRowIndex + 1) To m_lRowCount
        vVal = m_vData(i, lCol)
        
        If Not IsError(vVal) Then
            If IsEmpty(vVal) Then sKey = "" Else sKey = CStr(vVal)
            
            If Not dicIdx.exists(sKey) Then
                Set colRows = New Collection
                dicIdx.Add sKey, colRows
            Else
                Set colRows = dicIdx(sKey)
            End If
            colRows.Add i
        End If
    Next i

    ' 4. 캐시 등록 및 반환
    Set m_dicIndices(lCol) = dicIdx
    Set GetIndexByIndex = dicIdx
End Function


' 3-2a. 특정 조건에 맞는 행 번호들 찾기 (By 컬럼명)
Public Function SelectRows(ByVal sColName As String, ByVal vValue As Variant) As Collection
    Dim dicIdx As Object
    Dim sKey As String
    
    Set dicIdx = GetIndex(sColName)
    If dicIdx Is Nothing Then Set SelectRows = New Collection: Exit Function
    
    If IsEmpty(vValue) Then sKey = "" Else sKey = CStr(vValue)
    
    ' CompareMode 설정으로 인해 대소문자 무시 검색
    If dicIdx.exists(sKey) Then
        Set SelectRows = dicIdx(sKey)
    Else
        Set SelectRows = New Collection
    End If
End Function

' 3-2b. 특정 조건에 맞는 행 번호들 찾기 (By 숫자 인덱스) - v2.5.1 최적화
Public Function SelectRowsByIndex(ByVal lCol As Long, ByVal vValue As Variant) As Collection
    Dim dicIdx As Object
    Dim sKey As String

    Set dicIdx = GetIndexByIndex(lCol)
    If dicIdx Is Nothing Then Set SelectRowsByIndex = New Collection: Exit Function

    If IsEmpty(vValue) Then sKey = "" Else sKey = CStr(vValue)
    
    If dicIdx.exists(sKey) Then
        Set SelectRowsByIndex = dicIdx(sKey)
    Else
        Set SelectRowsByIndex = New Collection
    End If
End Function


' 3-3a. 특정 컬럼의 고유값 목록 추출 (By 컬럼명)
Public Function GetUniqueKeys(ByVal sColName As String) As Variant
    Dim dicIdx As Object
    Set dicIdx = GetIndex(sColName)
    
    If dicIdx Is Nothing Then
        GetUniqueKeys = Array()
    Else
        GetUniqueKeys = dicIdx.keys
    End If
End Function

' 3-3b. 특정 컬럼의 고유값 목록 추출 (By 숫자 인덱스) - v2.5.1 최적화
Public Function GetUniqueKeysByIndex(ByVal lCol As Long) As Variant
    Dim dicIdx As Object
    Set dicIdx = GetIndexByIndex(lCol)
    
    If dicIdx Is Nothing Then
        GetUniqueKeysByIndex = Array()
    Else
        GetUniqueKeysByIndex = dicIdx.keys
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

Public Property Get RawArray() As Variant
    RawArray = m_vData
End Property
