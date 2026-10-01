VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSelectBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' SPDX-License-Identifier: MIT
' Copyright (c) 2026 JM - see LICENSE
Option Explicit

Private m_selectLine As Long
Private m_endSelectLine As Long
Private m_caseLines As Collection

Private Sub Class_Initialize()
    Set m_caseLines = New Collection
End Sub

Public Property Get SelectLine() As Long
    SelectLine = m_selectLine
End Property

Public Property Let SelectLine(ByVal value As Long)
    m_selectLine = value
End Property

Public Property Get EndSelectLine() As Long
    EndSelectLine = m_endSelectLine
End Property

Public Property Let EndSelectLine(ByVal value As Long)
    m_endSelectLine = value
End Property

Public Property Get CaseCount() As Long
    CaseCount = m_caseLines.count
End Property

Public Function CaseLineAt(ByVal index As Long) As Long
    If index < 1 Or index > m_caseLines.count Then Err.Raise 9, "clsSelectBlock.CaseLineAt", "CASE index out of range."
    CaseLineAt = CLng(m_caseLines.item(index))
End Function

Public Sub AddCaseLine(ByVal lineNo As Long)
    m_caseLines.Add lineNo
End Sub
