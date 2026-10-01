VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
