VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS820_S02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約入力サブ-下段サイズ情報
'   プログラムＩＤ　：　FVS820_S02
'   作　成　日　　　：  2003/06/09
'   作　成　者　　　：  Eagle Soft 田島
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'**********************************************


Option Compare Database
Option Explicit
    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 
    Call Parent.Form_KeyDown(KeyCode, Shift)
    

End Sub
