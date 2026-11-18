VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 실행할 명령과 파라미터 원문을 보관하고 구문 검증 담당
' @depends CmdDef, ParamSpec, modCmdRegistry, modCondition
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

Public RuntimeLine As Long
Public sheetRow As Long
Public rawCmd As String
Public cmd As String
Public p1 As String
Public p2 As String
Public p3 As String
Public def As CmdDef

Public Sub Init(ByVal runtimeLineValue As Long, ByVal sheetRowValue As Long, _
                ByVal rawCmdValue As String, ByVal cmdValue As String, _
                ByVal p1Value As String, ByVal p2Value As String, ByVal p3Value As String, _
                ByVal CmdDef As CmdDef)
    RuntimeLine = runtimeLineValue
    sheetRow = sheetRowValue
    rawCmd = rawCmdValue
    cmd = cmdValue
    p1 = p1Value
    p2 = p2Value
    p3 = p3Value
    Set def = CmdDef
End Sub

' @desc 명령 등록ㆍ슬롯 사용ㆍ필수 원문ㆍ조건 연산자를 검사해 첫 오류를 반환함
Public Function RawSyntaxError() As String
    If def Is Nothing Then
        RawSyntaxError = "알 수 없는 명령어: " & cmd
        Exit Function
    End If

    RawSyntaxError = SlotSyntaxError(1, def.p1, p1)
    If Len(RawSyntaxError) = 0 Then RawSyntaxError = SlotSyntaxError(2, def.p2, p2)
    If Len(RawSyntaxError) = 0 Then RawSyntaxError = SlotSyntaxError(3, def.p3, p3)
End Function

Private Function SlotSyntaxError(ByVal slot As Long, ByVal spec As ParamSpec, ByVal value As String) As String
    If spec Is Nothing Then
        If Len(Trim$(value)) > 0 Then SlotSyntaxError = cmd & " does not use P" & slot & "."
    ElseIf spec.Required And Len(Trim$(value)) = 0 Then
        SlotSyntaxError = cmd & " P" & slot & "(" & spec.Meaning & ") 필수값 누락"
    ElseIf spec.RawKind = RK_CONDITION Then
        If Not modCondition.IsConditionOperator(value) Then SlotSyntaxError = "Invalid condition operator: " & value
    End If
End Function
