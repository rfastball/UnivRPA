VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 명령어 파라미터 슬롯 1개의 decode 스펙(의미/타입/해석모드/필수) 보유
' @depends modCmdRegistry
Option Explicit

' [decode 스펙] CmdDef가 P1/P2/P3로 3개 보유
Public Meaning As String        ' 사람용 의미 힌트 (예: "좌변", "연산자")
Public DataType As eDataType    ' DT_STR / DT_NUM
Public ResolveMode As eResolveMode  ' RESOLVE(토큰 치환) / RAW(원본 유지)
Public Required As Boolean      ' 필수 슬롯 여부

' @desc 4속성 받아 슬롯에 일괄 세팅 후 Me 반환 (생성자 대용 팩토리)
Public Function Init(ByVal vMeaning As String, ByVal vType As eDataType, _
                     ByVal vMode As eResolveMode, ByVal vRequired As Boolean) As ParamSpec
    Meaning = vMeaning
    DataType = vType
    ResolveMode = vMode
    Required = vRequired
    Set Init = Me
End Function

' [편의] RAW 슬롯 여부 (lvalue/연산자만 True)
Public Property Get IsRaw() As Boolean
    IsRaw = (ResolveMode = RAW)
End Property
