Attribute VB_Name = "modTokenResolver"
' @desc 파라미터 토큰({name}/{$...})을 스캔해 값으로 치환 반환 담당
' @depends modContext, modWinAPI, modParse, modStringDistance
Option Explicit

'==============================================================================
' [Layer 1] modTokenResolver: 파라미터 토큰 해석기 (modEngine에서 추출 - D4)
'==============================================================================
' 역할: RESOLVE 슬롯의 {...} 스팬을 깊이 인식으로 스캔해 값 치환만 수행한다.
'   시스템 토큰($BLANK/$CLIP/$LOOP/$ROW./$PIXEL: 등)과 사용자 메모리 {name}을 분류.
'   *? 와일드카드는 건드리지 않으며, 미정의 변수는 추천과 함께 에러를 던진다.
'==============================================================================

' [해석기 v2 토크나이저] 중첩 중괄호 지원: {$PIXEL:{cx},{cy}}
' @desc 원본 문자열 받아 토큰 치환된 문자열 반환, 짝 없는 중괄호는 리터럴 유지
Public Function ResolveParam(ByVal rawVal As String) As String
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
    Err.Raise Err.Number, "modTokenResolver.ResolveParam", Err.Description
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
        Case "$RTOTAL"
            If modContext.GetActiveForeachTotalRows() > 0 Then
                ResolveToken = CStr(modContext.GetActiveForeachTotalRows())
            Else
                ResolveToken = "{" & inner & "}"
            End If
            Exit Function
        Case "$RIDX"
            If modContext.GetActiveForeachOrdinal() > 0 Then
                ResolveToken = CStr(modContext.GetActiveForeachOrdinal())
            Else
                ResolveToken = "{" & inner & "}"
            End If
            Exit Function
    End Select

    If Left$(up, 5) = "$ROW." Then
        ResolveToken = ResolveRowToken(Mid$(key, 6), inner)
        Exit Function
    End If

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
    Dim value As String
    If modContext.TryGetVar(key, value) Then
        ResolveToken = value
    Else
        Err.Raise 5, "modTokenResolver.ResolveToken", UndefinedVarMessage(key, modContext.DefinedVarNames())
    End If

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modTokenResolver.ResolveToken", Err.Description
    Resume CleanUp
End Function

Private Function UndefinedVarMessage(ByVal key As String, ByVal candidates As Variant) As String
    Dim msg As String
    msg = "Undefined variable: {" & key & "}"

    Dim suggestion As String
    suggestion = modStringDistance.ClosestName(key, candidates)
    If Len(suggestion) > 0 Then msg = msg & ". Did you mean {" & suggestion & "}?"

    UndefinedVarMessage = msg
End Function

Private Function UndefinedColumnMessage(ByVal colName As String, ByVal candidates As Variant) As String
    Dim msg As String
    msg = "Undefined column: {$ROW." & colName & "}"

    Dim suggestion As String
    suggestion = modStringDistance.ClosestName(UCase$(Trim$(colName)), candidates)
    If Len(suggestion) > 0 Then msg = msg & ". Did you mean {$ROW." & suggestion & "}?"

    UndefinedColumnMessage = msg
End Function

Private Function ResolveRowToken(ByVal colName As String, ByVal originalInner As String) As String
    On Error GoTo ErrHandler

    Dim tbl As Object
    Set tbl = modContext.ActiveForeachTable()
    If tbl Is Nothing Or modContext.ActiveForeachRow() <= 0 Then
        ResolveRowToken = "{" & originalInner & "}"
        Exit Function
    End If

    ' [미지 컬럼 = 시끄러운 에러] 오타난 헤더가 조용히 빈 값이 되는 것을 막는다.
    '   {name}(미정의 변수)과 대칭 - 존재하는 컬럼의 빈 셀만 ""로 남긴다.
    Dim colKey As String: colKey = UCase$(Trim$(colName))
    If Not tbl.headers.exists(colKey) Then
        Err.Raise 5, "modTokenResolver.ResolveRowToken", UndefinedColumnMessage(colName, tbl.headers.keys)
    End If

    Dim value As Variant
    value = tbl.GetValue(modContext.ActiveForeachRow(), colName)
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        ResolveRowToken = ""
    Else
        ResolveRowToken = CStr(value)
    End If

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modTokenResolver.ResolveRowToken", Err.Description
    Resume CleanUp
End Function

' [Helper] $PIXEL 좌표 -> 색상 헥스 (coordText 는 이미 해석된 "x,y")
Private Function ResolvePixelToken(ByVal coordText As String) As String
    On Error GoTo ErrHandler
    Dim coords() As String
    coords = Split(coordText, ",")
    If UBound(coords) = 1 Then
        ResolvePixelToken = modWinAPI.GetColorHexAt(modParse.ParseLongOrDefault(coords(0), 0), modParse.ParseLongOrDefault(coords(1), 0))
    Else
        ResolvePixelToken = "{$PIXEL:" & coordText & "}"  ' 형식 불일치 시 원본
    End If

CleanUp:
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "modTokenResolver.ResolvePixelToken", Err.Description
    Resume CleanUp
End Function
