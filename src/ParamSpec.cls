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
Public RawKind As eRawKind      ' RAW 슬롯 시맨틱 (C2). RESOLVE 슬롯은 항상 RK_NONE

' @desc 슬롯 속성 받아 일괄 세팅 후 Me 반환 (생성자 대용 팩토리)
'   불변식: RawKind는 RAW 슬롯에서만 유효 - RAW인데 생략하면 RK_TEXT,
'   RESOLVE면 지정값과 무관하게 RK_NONE으로 강제한다.
Public Function Init(ByVal vMeaning As String, ByVal vType As eDataType, _
                     ByVal vMode As eResolveMode, ByVal vRequired As Boolean, _
                     Optional ByVal vRawKind As eRawKind = RK_NONE) As ParamSpec
    Meaning = vMeaning
    DataType = vType
    ResolveMode = vMode
    Required = vRequired
    If vMode = raw Then
        If vRawKind = RK_NONE Then RawKind = RK_TEXT Else RawKind = vRawKind
    Else
        RawKind = RK_NONE
    End If
    Set Init = Me
End Function

' [편의] RAW 슬롯 여부 (RawKind <> RK_NONE 과 동치)
Public Property Get IsRaw() As Boolean
    IsRaw = (ResolveMode = raw)
End Property
