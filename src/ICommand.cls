VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @desc 카테고리별 명령어 핸들러의 공통 디스패치 인터페이스 정의
Option Explicit

' 2단 디스패치 공통 약속.
'   cmd  = 정규화된 명령어 정식명
'   R    = 현재 줄 번호
'   args = ParamSpec.ResolveMode에 따라 디스패처가 미리 디코드한 슬롯(1..3)
'          RAW 슬롯은 원본, RESOLVE 슬롯은 토큰 치환본
'   반환 = CMD_NEXT_LINE(0) / 양수(점프 줄) / CMD_HALT(-2) / CMD_ERROR(-1)
Public Function Execute(ByVal cmd As String, ByVal R As Long, ByRef args() As String) As Long
End Function
