VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RuntimeLine As Long
Public SheetRow As Long
Public RawCmd As String
Public cmd As String
Public P1 As String
Public P2 As String
Public P3 As String
Public def As CmdDef

Public Sub Init(ByVal runtimeLineValue As Long, ByVal sheetRowValue As Long, _
                ByVal rawCmdValue As String, ByVal cmdValue As String, _
                ByVal p1Value As String, ByVal p2Value As String, ByVal p3Value As String, _
                ByVal cmdDef As CmdDef)
    RuntimeLine = runtimeLineValue
    SheetRow = sheetRowValue
    RawCmd = rawCmdValue
    cmd = cmdValue
    P1 = p1Value
    P2 = p2Value
    P3 = p3Value
    Set def = cmdDef
End Sub
