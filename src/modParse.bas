Attribute VB_Name = "modParse"
' @desc 스크립트 숫자 문자열을 엄격 문법으로 파싱해 반환 담당
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

'==============================================================================
' [Utility] 무상태 숫자 파서 (레이어 중립, @depends 없음)
' 허용 문법: [양끝공백] [+|-] 숫자+ [ "." 숫자+ ]  (소수점은 "." 고정, 로케일 독립)
' 거부: 빈값, "1,000", "$5", "1e3", "&HFF", "12abc", 내부 공백
' 검증 게이트(ValidateDecodedSlots)/핸들러 소비/조건 비교(modCondition)가
' 모두 이 파서를 공유한다 - IsNumeric과 Val의 판정 불일치 제거가 목적.
'==============================================================================

' @desc 숫자 문자열 받아 Double 파싱 성공 여부 반환, 성공 시 result에 기록
Public Function TryParseDoubleStrict(ByVal text As String, ByRef result As Double) As Boolean
    Dim s As String
    s = Trim$(text)
    If Not IsStrictNumberSyntax(s) Then Exit Function
    result = val(s)
    TryParseDoubleStrict = True
End Function

' @desc 숫자 문자열 받아 Long 파싱 성공 여부 반환, 소수는 CLng 반올림
Public Function TryParseLongStrict(ByVal text As String, ByRef result As Long) As Boolean
    Dim d As Double
    If Not TryParseDoubleStrict(text, d) Then Exit Function
    If d < -2147483648# Or d > 2147483647# Then Exit Function
    ' CLng 반올림(기존 CLng(Val(...)) 소비처의 "1.5" -> 2 동작 보존)
    result = CLng(d)
    TryParseLongStrict = True
End Function

' @desc 숫자 문자열 받아 Long 반환, 파싱 실패 시 기본값 반환
Public Function ParseLongOrDefault(ByVal text As String, ByVal defaultValue As Long) As Long
    Dim n As Long
    If TryParseLongStrict(text, n) Then
        ParseLongOrDefault = n
    Else
        ParseLongOrDefault = defaultValue
    End If
End Function

' @desc 게이트 통과 보장된 숫자 문자열 받아 Long 반환, 실패 시 형식오류(13) 발생
Public Function ParseLongStrict(ByVal text As String) As Long
    Dim n As Long
    If Not TryParseLongStrict(text, n) Then
        Err.Raise 13, "modParse.ParseLongStrict", "숫자 아님: '" & text & "'"
    End If
    ParseLongStrict = n
End Function

Private Function IsStrictNumberSyntax(ByVal s As String) As Boolean
    Dim total As Long
    total = Len(s)
    If total = 0 Then Exit Function

    Dim idx As Long
    idx = 1
    Select Case Mid$(s, 1, 1)
        Case "+", "-"
            idx = 2
    End Select

    Dim digitCount As Long
    Dim dotSeen As Boolean
    Do While idx <= total
        Select Case Mid$(s, idx, 1)
            Case "0" To "9"
                digitCount = digitCount + 1
            Case "."
                If dotSeen Then Exit Function
                dotSeen = True
            Case Else
                Exit Function
        End Select
        idx = idx + 1
    Loop

    IsStrictNumberSyntax = (digitCount > 0)
End Function

' @desc 쓰기 토큰 받아 MEM/CLIP 종류와 키 반환, 빈 이름/읽기전용이면 False 반환
Public Function TryParseLValue(ByVal text As String, ByRef kind As String, ByRef key As String) As Boolean
    kind = ""
    key = ""
    text = Trim$(text)
    If Len(text) < 2 Then Exit Function
    If Left$(text, 1) <> "{" Or Right$(text, 1) <> "}" Then Exit Function

    Dim inner As String
    inner = Trim$(Mid$(text, 2, Len(text) - 2))
    If Len(inner) = 0 Then Exit Function
    If UCase$(inner) = "$CLIP" Then
        kind = "CLIP"
    ElseIf Left$(inner, 1) = "$" Then
        Exit Function
    Else
        kind = "MEM"
        key = inner
    End If
    TryParseLValue = True
End Function
