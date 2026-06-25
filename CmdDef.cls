VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 명령어 1개의 메타데이터 레코드(식별/구조/서식/파라미터) 보유
' @depends modCmdRegistry, ParamSpec
Option Explicit

' [식별]
Public Canonical As String      ' 명령어 정식명 (레지스트리 키)
Public IsPattern As Boolean     ' 패턴/별칭 여부 (예약 슬롯)

' [구조]
Public Category As eCategory    ' FLOW/ACTION/VISION/MEMORY/SYSTEM (1단 디스패치 키)
Public BlockRole As eBlockRole  ' STARTER/ENDER/MIDBLOCK/NOBLOCK
Public BlockGroup As String     ' IF/WHILE/REPEAT/SELECT/"" (검증·필터 용도)

' [서식] Category와 독립 축
Public ColorGroup As eColorGroup  ' clrFLOW/clrACTION/clrVISION/clrSYSTEM
Public PaintScope As ePaintScope  ' FULLROW / CMDCELL

' [파라미터] 각각 ParamSpec (없으면 Nothing)
Public P1 As ParamSpec
Public P2 As ParamSpec
Public P3 As ParamSpec

' @desc 식별/구조/서식/파라미터 일괄 받아 레코드에 세팅 후 Me 반환 (생성자 대용 팩토리)
Public Function Init(ByVal vCanon As String, ByVal vCat As eCategory, ByVal vRole As eBlockRole, _
                     ByVal vGroup As String, ByVal vColor As eColorGroup, ByVal vScope As ePaintScope, _
                     Optional ByVal sp1 As ParamSpec, _
                     Optional ByVal sp2 As ParamSpec, _
                     Optional ByVal sp3 As ParamSpec) As CmdDef
    Canonical = vCanon
    Category = vCat
    BlockRole = vRole
    BlockGroup = vGroup
    ColorGroup = vColor
    PaintScope = vScope
    Set P1 = sp1
    Set P2 = sp2
    Set P3 = sp3
    Set Init = Me
End Function

' [편의] 슬롯 번호(1/2/3)로 ParamSpec 반환 (없으면 Nothing)
Public Function Param(ByVal idx As Long) As ParamSpec
    Select Case idx
        Case 1: Set Param = P1
        Case 2: Set Param = P2
        Case 3: Set Param = P3
    End Select
End Function
