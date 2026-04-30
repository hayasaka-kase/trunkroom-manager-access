VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS810_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード一覧サブ
'   プログラムＩＤ　：　FVS810_S01
'   作　成　日　　　：  2003/06/09
'   作　成　者　　　：  Eagle Soft 田島
'**********************************************
'修正履歴
'   修　正　日　　　：  2004/11/06
'   修　正　者　　　：  N.MIURA
'   修　正　内　容　：  所在地住所３追加
'**********************************************


Option Compare Database
Option Explicit

Const CONTAINER_LIST = "FVS820"

Private Sub cmd_詳細_Click()
   
    Dim stDocName As String
    Dim stLinkCriteria As String

    doCmd.OpenForm CONTAINER_LIST, , , , , , Me.txt_ヤードコード

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Call Parent.Form_KeyDown(KeyCode, Shift)
End Sub

