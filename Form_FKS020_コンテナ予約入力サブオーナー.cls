VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS020_コンテナ予約入力サブオーナー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　利用状況オーナー
'   プログラムＩＤ　：　FKS020_コンテナ予約入力サブオーナー
'   作　成　日　　　：  2019/11/20
'   作　成　者　　　：  Y.WADA
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  yyyy/mm/dd
'   修　正　者　　　：  x.xxxx
'   修　正　内　容　：
'   Ver             ：  0.x
'**********************************************

Option Compare Database
Option Explicit

'==============================================================================*
'
'        MODULE_NAME      :詳細ボタン
'        MODULE_ID        :cmd_syosai_Click
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_syosai_Click()
    Dim strArg As String
    
    On Error GoTo Exception
    
    strArg = ""
    strArg = strArg & "ｦ"
    strArg = strArg & vbTab & Nz(Me![HKEIT_SHIRC], "")
    strArg = strArg & vbTab & "1"   '参照モード(0:更新あり、1:参照、2:一部更新（２．契約者住所･連絡先、３．その他住所･連絡先のみ更新可）)
    strArg = strArg & vbTab         'EOF

    '（加瀬システム）仕入先マスタを開く
    Call MSZZ080_SysFormOpen("MFAM00C", strArg)

Exit Sub

Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description

End Sub
'****************************  ended or program ********************************
