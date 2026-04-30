VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS021_S03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　空状況検索
'   プログラムＩＤ　：　FVS021_S03
'   作　成　日　　　：  2005/07/28
'   作　成　者　　　：  K.ISHIZAKA
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  20050829
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.1
'   修　正　内　容　：  表示順変更
'**********************************************
'   修　正　日　　　：  2005/11/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.2
'   修　正　内　容　：  デザイン項目追加と色変更
'**********************************************
'   修　正　日　　　：  2005/12/05
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.3
'   修　正　内　容　：  自社理由の対応
'**********************************************
'   修　正　日　　　：  2006/08/31
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.9
'   修　正　内　容　：  ・ヤード非表示日を過ぎたヤードは選択リストから外す
'                       ・営業終了日を過ぎたヤードは検索対象外
'                       ・営業開始日を過ぎていなければ一覧上に表示する
'                       ・営業終了日が入力されていれば一覧上に表示する
'                       ・地主解約のチェックを外す
'                       ・コンテナ種別=0の条件をはずす
'                       ・備考欄に部屋と顧客の直近修繕を出力
'                       ・修繕入力に遷移可能にする
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   イベント関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : 予約受付ボタン押下時処理
'       MODULE_ID       : cmdUke_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUke_Click()
    Call Me.Parent.ShowFVS500(Me.txt_YOUKT_NO, Nz(Me.txt_YOUKT_UKNO))
End Sub

'>>> DELETE START 2006/08/31 K.ISHIZAKA >>>
''==============================================================================*
''
''       MODULE_NAME     : カレントレコード移動時処理
''       MODULE_ID       : Form_Current
''       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub Form_Current()
'    Dim bYoyaku             As Boolean
'
''    bYoyaku = (Nz(Me.txt_YOUKT_UKNO, "") <> "")                                'DELETE 2005/12/05 K.ISHIAKA
'    bYoyaku = (Nz(Me.txt_YOUKT_YUKBN, "20") <> "20")                            'INSERT 2005/12/05 K.ISHIAKA
'    With Me.txt_YOUKT_RENKBN
'        .Enabled = Not bYoyaku
'        .Locked = bYoyaku
'    End With
'    With Me.txt_YOUKT_BIKO
'        .Enabled = Not bYoyaku
'        .Locked = bYoyaku
'    End With
'End Sub
'<<< DELETE END   2006/08/31 K.ISHIZAKA <<<

'==============================================================================*
'
'       MODULE_NAME     : フォームキー押下時処理
'       MODULE_ID       : Form_KeyDown
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Me.Parent.KeyDown(KeyCode, Shift)
End Sub

'>>> DELETE START 2006/08/31 K.ISHIZAKA >>>
''==============================================================================*
''
''       MODULE_NAME     : 内容変更時処理
''       MODULE_ID       : txt_YOUKT_BIKO_Change
''       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_YOUKT_BIKO_Change()
'    Me.txt_CHANGE = 1
'    Call Me.Parent.SetChangeFlag
'End Sub
'
''==============================================================================*
''
''       MODULE_NAME     : 内容変更時処理
''       MODULE_ID       : txt_YOUKT_RENKBN_Change
''       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_YOUKT_RENKBN_Change()
'    Me.txt_CHANGE = 1
'    Call Me.Parent.SetChangeFlag
'End Sub
'<<< DELETE END   2006/08/31 K.ISHIZAKA <<<

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_BIKO_DblClick
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_BIKO_DblClick(Cancel As Integer)
    Call Form_KeyDown(CmReprMod.P_呼出_部屋, 0)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_NAME_DblClick
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_NAME_DblClick(Cancel As Integer)
    Call Form_KeyDown(CmReprMod.P_呼出_顧客, 0)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_NO_DblClick
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_NO_DblClick(Cancel As Integer)
    Call Form_KeyDown(CmReprMod.P_呼出_部屋, 0)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_RENKBN_DblClick
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_RENKBN_DblClick(Cancel As Integer)
    Call Form_KeyDown(CmReprMod.P_呼出_顧客, 0)
End Sub

'****************************  ended or program ********************************
