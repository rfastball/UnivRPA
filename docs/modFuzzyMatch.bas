Attribute VB_Name = "modFuzzyMatch"
' @desc Levenshtein Distance 기반 문자열 퍼지 매칭 함수 담당
' ===========================================================================
' 모듈   : modFuzzyMatch
' 목적   : 문자열 퍼지 매칭 (Levenshtein Distance 기반)
' 의존성 : 없음 (순수 VBA 내장 함수만 사용)
' Public API:
'   Get_Distance    - 두 문자열 간 Levenshtein 편집 거리 반환
'   Get_Similarity  - 두 문자열 간 유사도(0.0 ~ 1.0) 반환
'   Find_Best_Match - 후보 목록에서 가장 유사한 문자열 반환
' ===========================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Private Helper: Min3
' 목적 : 3개의 Long 값 중 최솟값 반환 (레벤슈타인 내부 루프 전용)
' ---------------------------------------------------------------------------
Private Function Min3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Dim lMin As Long
    lMin = a
    If b < lMin Then lMin = b
    If c < lMin Then lMin = c
    Min3 = lMin
End Function

' ---------------------------------------------------------------------------
' Public Function: Get_Distance
' 목적 : 두 문자열 간 Levenshtein 편집 거리 반환
' 파라미터:
'   str1       - 비교할 첫 번째 문자열
'   str2       - 비교할 두 번째 문자열
'   ignoreCase - True(기본값)이면 대소문자 무시
' 반환 : 편집 거리 (Long)
' ---------------------------------------------------------------------------
' @desc 두 문자열을 받아 Levenshtein 편집 거리를 Long으로 반환
Public Function Get_Distance( _
    ByVal str1 As String, _
    ByVal str2 As String, _
    Optional ByVal ignoreCase As Boolean = True) As Long

    Dim s1 As String
    Dim s2 As String
    Dim len1 As Long
    Dim len2 As Long
    Dim i As Long
    Dim j As Long
    Dim cost As Long
    Dim d() As Long

    ' 진입 시 Trim 처리
    s1 = Trim$(str1)
    s2 = Trim$(str2)

    ' ignoreCase 적용
    If ignoreCase Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    len1 = Len(s1)
    len2 = Len(s2)

    ' Early Exit
    If len1 = 0 And len2 = 0 Then
        Get_Distance = 0
        Exit Function
    End If
    If len1 = 0 Then
        Get_Distance = len2
        Exit Function
    End If
    If len2 = 0 Then
        Get_Distance = len1
        Exit Function
    End If
    If s1 = s2 Then
        Get_Distance = 0
        Exit Function
    End If

    ' Levenshtein DP 배열 동적 할당
    ReDim d(0 To len1, 0 To len2) As Long

    ' 초기화
    For i = 0 To len1
        d(i, 0) = i
    Next i
    For j = 0 To len2
        d(0, j) = j
    Next j

    ' 2중 루프
    For i = 1 To len1
        For j = 1 To len2
            cost = IIf(Mid$(s1, i, 1) = Mid$(s2, j, 1), 0, 1)
            d(i, j) = Min3(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i

    Get_Distance = d(len1, len2)
End Function

' ---------------------------------------------------------------------------
' Public Function: Get_Similarity
' 목적 : 두 문자열 간 유사도 반환 (0.0 = 완전 다름, 1.0 = 완전 일치)
' 파라미터:
'   str1       - 비교할 첫 번째 문자열
'   str2       - 비교할 두 번째 문자열
'   ignoreCase - True(기본값)이면 대소문자 무시
' 반환 : 유사도 (Double, 0.0 ~ 1.0)
' ---------------------------------------------------------------------------
' @desc 두 문자열을 받아 유사도 0.0~1.0을 Double로 반환
Public Function Get_Similarity( _
    ByVal str1 As String, _
    ByVal str2 As String, _
    Optional ByVal ignoreCase As Boolean = True) As Double

    Dim s1 As String
    Dim s2 As String
    Dim maxLen As Long

    ' 진입 시 Trim 처리
    s1 = Trim$(str1)
    s2 = Trim$(str2)

    ' Early Exit: 둘 다 빈 문자열 = 완전 일치
    If Len(s1) = 0 And Len(s2) = 0 Then
        Get_Similarity = 1#
        Exit Function
    End If

    maxLen = Len(s1)
    If Len(s2) > maxLen Then maxLen = Len(s2)

    If maxLen = 0 Then
        Get_Similarity = 1#
        Exit Function
    End If

    Get_Similarity = 1# - (CDbl(Get_Distance(s1, s2, ignoreCase)) / CDbl(maxLen))
End Function

' ---------------------------------------------------------------------------
' Public Function: Find_Best_Match
' 목적 : 후보 목록에서 targetStr과 가장 유사한 문자열 반환
' 파라미터:
'   targetStr  - 검색 기준 문자열
'   candidates - 후보 목록 (배열 Variant() 또는 Collection)
'   threshold  - 최소 유사도 기준 (기본값 0.75)
'   ignoreCase - True(기본값)이면 대소문자 무시
' 반환 : 가장 유사한 후보 문자열 (매칭 실패 시 "" 반환)
' ---------------------------------------------------------------------------
' @desc targetStr과 가장 유사한 후보를 반환, 임계값 미달 시 "" 반환
Public Function Find_Best_Match( _
    ByVal targetStr As String, _
    ByVal candidates As Variant, _
    Optional ByVal threshold As Double = 0.75, _
    Optional ByVal ignoreCase As Boolean = True) As String

    Dim sTarget As String
    Dim sCandidate As String
    Dim sTypeName As String
    Dim vItem As Variant
    Dim dScore As Double
    Dim dBestScore As Double
    Dim sBestMatch As String
    Dim i As Long

    ' 진입 시 targetStr Trim 처리
    sTarget = Trim$(targetStr)

    ' Early Exit: 빈 문자열
    If Len(sTarget) = 0 Then
        Find_Best_Match = ""
        Exit Function
    End If

    dBestScore = -1#
    sBestMatch = ""

    sTypeName = TypeName(candidates)

    ' candidates 타입 분기
    Select Case True

        ' a) 배열(Variant 배열)
        Case sTypeName = "Variant()" Or sTypeName = "String()" Or _
             sTypeName = "Long()" Or sTypeName = "Integer()" Or _
             (Left$(sTypeName, 1) <> "" And Right$(sTypeName, 2) = "()")

            For i = LBound(candidates) To UBound(candidates)
                sCandidate = CStr(candidates(i))
                dScore = Get_Similarity(sTarget, sCandidate, ignoreCase)
                If dScore > dBestScore And dScore >= threshold Then
                    dBestScore = dScore
                    sBestMatch = sCandidate
                End If
            Next i

        ' b) Collection
        Case sTypeName = "Collection"
            For Each vItem In candidates
                sCandidate = CStr(vItem)
                dScore = Get_Similarity(sTarget, sCandidate, ignoreCase)
                If dScore > dBestScore And dScore >= threshold Then
                    dBestScore = dScore
                    sBestMatch = sCandidate
                End If
            Next vItem

        ' c) 그 외 → "" 반환
        Case Else
            Find_Best_Match = ""
            Exit Function

    End Select

    Find_Best_Match = sBestMatch
End Function
