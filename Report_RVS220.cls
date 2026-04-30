VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RVS220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　紹介制限ヤード一覧出力
'   プログラムＩＤ　：　RVS210
'   作　成　日　　　：  2007/05/01
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.0
'**********************************************
Option Compare Database
Option Explicit

Private Const P_MAX_ROW_COUNT As Integer = 20

Private intTotalRowCount As Integer    ' 1ページのトータル行数

Private intRowCount      As Integer    ' 行数カウント変数
Private intMainYardCount As Integer    ' 解約ヤード毎の件数

' ***********************************************
' ページヘッダーセクション_Print
' ***********************************************
Private Sub ページヘッダーセクション_Print(Cancel As Integer, PrintCount As Integer)
    intTotalRowCount = 0
End Sub

' ***********************************************
' グループヘッダー0_Format
' ***********************************************
Private Sub グループヘッダー0_Format(Cancel As Integer, FormatCount As Integer)

'    intMainYardCount = 0
'    intRowCount = 0
'
'    intMainYardCount = DCount("解約ヤードコード", "RVS210_W01", "解約ヤードコード = '" & Me![ダミー解約ヤードコード] & "' ")
End Sub

' ***********************************************
' グループヘッダー1_Format
' ***********************************************
Private Sub グループヘッダー1_Format(Cancel As Integer, FormatCount As Integer)

    intMainYardCount = 0
    intRowCount = 0

    intMainYardCount = DCount("解約ヤードコード", "RVS220_W01", "解約ヤードコード = '" & Me![ダミー解約ヤードコード] & "' " & _
                              "AND 紹介制限期限 = '" & Me![ダミー紹介制限期限] & "' ")

End Sub

' ***********************************************
' 詳細_Print
' ***********************************************
Private Sub 詳細_Print(Cancel As Integer, PrintCount As Integer)

    intTotalRowCount = intTotalRowCount + 1
    intRowCount = intRowCount + 1

    If (intRowCount = intMainYardCount Or intTotalRowCount = 20) Then
        Me![line011].Visible = True
    Else
        Me![line011].Visible = False
    End If
End Sub




