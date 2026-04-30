VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS450_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　価格テンプレート設定ヤード一覧
'   プログラムＩＤ　：　FKS450
'   作　成　日　　　：  2017/11/01
'   作　成　者　　　：  イーグルソフト Y
'   Ver             ：  1.0
'**********************************************
'修正履歴
'   更　新　日　　　：  2017/12/01
'   更　新　者　　　：  イーグルソフト Y
'   修　正　内　容　：　キャンペーン一括設定対応
'   Ver             ：  1.1
'**********************************************
'修正履歴
'   更　新　日　　　：  2018/08/11
'   更　新　者　　　：  イーグルソフト Y
'   修　正　内　容　：　事務手数料対応
'   Ver             ：  1.3
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'
'        MODULE_NAME      :cmb_Kakaku_New_Setei_AfterUpdate
'        機能             :cmb_Kakaku_New_Setei_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kakaku_New_Setei_AfterUpdate()
    Refresh
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Waribiki_New_Setei_AfterUpdate
'        機能             :cmcmb_Waribiki_New_Setei_AfterUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Waribiki_New_Setei_AfterUpdate()
    Refresh
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'F1=「終了」 F2=「検索」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Parent.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Parent.cmd_検索.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

