VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : コンテナ管理サブシステム
'
'        PROGRAM_NAME    : メニュー
'        PROGRAM_ID      : FKS000
'        PROGRAM_KBN     : FORM
'        Ver             : 2.0
'
'        CREATE          :
'        CERATER         :
'
'        UPDATE          : 2003/10/28
'        UPDATER         : N.MIURA
'        Ver             : 3.0
'                        :
'
'        UPDATE          : 2004/08/06
'        UPDATER         : N.MIURA
'        Ver             : 3.1
'                        : 代引きリスト追加
'
'        UPDATE          : 2004/11/22
'        UPDATER         : N.MIURA
'        Ver             : 3.2
'                        : フォーム名変更"FKS160_解約一覧表"
'
'        UPDATE          : 2004/11/24
'        UPDATER         : N.MIURA
'        Ver             : 3.3
'                        : フォーム名変更
'
'        UPDATE          : 2004/12/08
'        UPDATER         : N.MIURA
'        Ver             : 3.4
'                        : 解約入力機能追加
'
'        UPDATE          : 2005/01/29
'        UPDATER         : N.MIURA
'        Ver             : 3.5
'                        : 予約データ出力
'
'        UPDATE          : 2005/07/01
'        UPDATER         : c.iguchi
'        Ver             : 3.6
'                        : 画面レイアウト変更
'
'        UPDATE          : 2005/07/13
'        UPDATER         : c.iguchi
'        Ver             : 3.7
'                        : FKS041からFVS600変更
'
'        UPDATE          : 2005/07/30
'        UPDATER         : K.ISHIZAKA
'        Ver             : 3.8
'                        : 空き検索新版対応
'
'        UPDATE          : 2005/08/02
'        UPDATER         : c.iguchi
'        Ver             : 3.9
'                        : メニュー画面変更
'                        : ボタン追加
'
'        UPDATE          : 2005/10/07
'        UPDATER         : c.iguchi
'        Ver             : 4.0
'                        : ボタン追加
'
'        UPDATE          : 2006/03/18
'        UPDATER         : tajima
'        Ver             : 4.1
'                        : 部門コードによる契約入力制御追加
'
'        UPDATE          : 2006/06/24
'        UPDATER         : tajima
'        Ver             : 4.2
'                        : 修繕メニュー追加
'
'        UPDATE          : 2007/03/01
'        UPDATER         : tajima
'        Ver             : 4.3
'                        : メニュー内容を若干変更
'
'        UPDATE          : 2007/05/01
'        UPDATER         : tajima
'        Ver             : 4.4
'                        : ヤード解約業務ボタンをメニュー内容に追加
'
'        UPDATE          : 2007/07/14
'        UPDATER         : tajima
'        Ver             : 4.5
'                        : 全メニューのボタンを権限による制御とした
'
'        UPDATE          : 2008/06/03
'        UPDATER         : N.Miura
'        Ver             : 4.6
'                        : 不要機能の削除
'                          FKS390(払い込みﾃﾞｰﾀ作成)
'
'        UPDATE          : 2009/06/17
'        UPDATER         : N.Miura
'        Ver             : 4.7
'                        : 不要機能の削除
'                          01.LF0038(売上リスト)
'                          02.FKS060(メンテ入力仮)
'                          03.FKS170(請求明細一覧表)
'                          04.LF0041(契約日報リスト)
'                          05.FKS160(解約一覧表)
'                          06.FKS150(更新一覧表)
'                          07.LF0043(契約一覧表２)
'                          08.LF0049(修繕データ出力)
'                          09.LF0016(宛名印刷”満了日”)
'                          10.FKS190(預り保証金一覧表)
'                          11.FKS010(地主契約入力)
'                          12.FKS100(地主台帳)
'                          13.FKS110(地主契約更新一覧表)
'                          14.FKS008(損益管理メニュー)
'                          15.FKS009(分析帳票メニュー)
'                          16.FKS00R(修繕管理メニュー)
'
'        UPDATE          : 2009/06/17
'        UPDATER         : N.Miura
'        Ver             : 4.8
'                          メニュー追加
'                          01.ヤード修繕
'                          02.コンテナ（部屋）修繕
'                          03.顧客問い合わせ
'                          04.修理依頼書出力
'
'        UPDATE          : 2009/06/18
'        UPDATER         : N.Miura
'        Ver             : 4.9
'                          メニュー追加
'                          01.契約状況一覧表
'                          02.分析表
'                          03.坪単価一覧表
'                          04.稼動明細一覧表
'                          05.予約集計表
'                          06.契約・解約集計表
'
'        UPDATE          : 2009/06/18
'        UPDATER         : N.Miura
'        Ver             : 5.0
'                        : 01.FKS00M(マスタ保守)の削除
'                          02.FKS430(顧客マスタ保守)の追加
'                          03.FKS440(ヤードマスタ保守)の追加
'                          04.FKS600(セキュリティカード保守)の追加
'                          05.FKS800(ボックスマスタ保守)の追加
'                          06.FKS700(割引マスタ保守)の追加
'                          07.FKS720(ヤード別割引設定)の追加
'                          08.FKS480(顧客リスト)の追加
'                          09.FKS490(ヤードリスト)の追加
'                          10.FKS540(近隣ヤードリスト)の追加
'                          11.FKS820(ボックスリスト)の追加
'
'        UPDATE          : 2009/10/26
'        UPDATER         : M.Honda
'        Ver             : 5.1
'                        : 01.FVS100(予約・契約・解約集計表)の追加
'
'        UPDATE          : 2009/11/05
'        UPDATER         : M.Honda
'        Ver             : 5.2
'                        : 01.FVS100(予約・契約・解約集計表)の名称を変更
'                        : 予約・契約・解約集計表 ⇒ 集計件数データ出力
'
'        UPDATE          : 2011/03/10
'        UPDATER         : M.Honda
'        Ver             : 5.3
'                        : 01.FVS300（毎月回収方法進捗管理）追加
'
'        UPDATE          : 2011/07/07
'        UPDATER         : M.RYU
'        Ver             : 5.4
'                        : 宛先印刷画面ID変更
'
'
'        UPDATE          : 2013/04/16
'        UPDATER         : M.HONDA
'        Ver             : 5.5
'                        : 起動時のログを取得
'
'        UPDATE          : 2013/09/30
'        UPDATER         : K.KINEBUCHI
'        Ver             : 5.6
'                        : メニューより「業務終了」ボタンを押した際、確認メッセージは
'                        　表示しない
'
'        UPDATE          : 2014/03/03
'        UPDATER         : M.HONDA
'        Ver             : 5.7
'                        : ヤードコンテナ一覧ボタンを追加
'
'        UPDATE          : 2017/11/05
'        UPDATER         : K.SATO
'        Ver             : 5.8
'                        : 部屋一括修繕ボタンを追加
'
'        UPDATE          : 2017/11/01
'        UPDATER         : H.TAJIMA
'        Ver             : 5.9
'                        : テンプレート一括設定ボタンを追加
'                        : 価格テンプレート保守ボタンを追加
'
'        UPDATE          : 2019/03/25
'        UPDATER         : M.HONDA
'        Ver             : 6.0
'                        : 部門 : Hのみ利用していない機能を非表示へ
'
'        UPDATE          : 2020/05/23
'        UPDATER         : N.IMAI
'        Ver             : 6.1
'                        : コンテナ在庫表ボタンを追加
'
'        UPDATE          : 2026/01/23
'        UPDATER         : M.HONDA
'        Ver             : 6.3
'                        : 毎月回収方法進捗管理を表示へ変更
'

'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS000"
Private strCONT_BUMOC As String 'ADD 2006/03/18 Tajima
' 2007/07/14 ADD tajima
'ADO定数（アクティブXを参照設定しないためローカルで定義する）
Private Const adCmdText = 1
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1
Private Const adOpenDynamic = 2
Private Const adLockPessimistic = 2
Private Const adCmdStoredProc = 4

'==============================================================================*
'
'        MODULE_NAME      :ヤードコンテナ一覧
'        MODULE_ID        :cmdFKS830_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS830_Click()
    Call MSHS010_M00("FKS830", "M", strCONT_BUMOC)
    doCmd.OpenForm "FKS830"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :予約・契約・解約集計表
'        MODULE_ID        :cmdFVS100_Click
'        CREATE_DATE      :2009/10/26 M.Honda
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS100_Click()
    
    Call MSHS010_M00("FVS100", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS100"

End Sub

'==============================================================================*
'
'        MODULE_NAME      :毎月回収方法進捗管理
'        MODULE_ID        :cmdFVS300_Click
'        CREATE_DATE      :2011/03/10 M.HONDA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS300_Click()
    Call MSHS010_M00("FVS300", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS300"
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームオープン
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    
    Call MSHS010_M00(PROG_ID, "1")

    gstMenubar = "ON"                                                           'INSERT 20030815 N.MIURA START
    'gstMenubar = "OFF"                                                         'INSERT 20030815 N.MIURA START
    Call CmSetMenuBar
    
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")      'ADD 2006/03/18 Tajima
    Select Case strCONT_BUMOC                                                   'REP 2006/03/18 Tajima
'    Case "0"                                                                   'DELETE 20090617 N.MIURA
'         Me.TX_SYSTN = "看板管理システム"                                      'DELETE 20090617 N.MIURA
    Case "1"
         Me.TX_SYSTN = "オフィス管理システム"
         Me.TX_SYSTN.BackColor = RGB(255, 165, 0)
         Me.ラベル425.BackColor = RGB(255, 165, 0)
         Me.ラベル427.BackColor = RGB(255, 165, 0)
         Me.ラベル319.BackColor = RGB(255, 165, 0)
         Me.ラベル435.BackColor = RGB(255, 165, 0)
         Me.ラベル332.BackColor = RGB(255, 165, 0)
         Me.ラベル442.BackColor = RGB(255, 165, 0)
         Me.ラベル395.BackColor = RGB(255, 165, 0)
         Me.ボックス308.BackColor = RGB(128, 0, 128)
         Me.ボックス432.BackColor = RGB(128, 0, 128)
         Me.ボックス316.BackColor = RGB(128, 0, 128)
         Me.ボックス411.BackColor = RGB(128, 0, 128)
         Me.ボックス331.BackColor = RGB(128, 0, 128)
         Me.ボックス441.BackColor = RGB(128, 0, 128)
         Me.ボックス394.BackColor = RGB(128, 0, 128)
         Me.フォームヘッダー.BackColor = RGB(128, 0, 128)
    Case "8"
         Me.TX_SYSTN = "トランク管理システム"
    Case "H"
         Me.TX_SYSTN = "コンテナ管理システム"
         Me.TX_SYSTN.BackColor = RGB(255, 0, 204)
         Me.ラベル425.BackColor = RGB(255, 0, 204)
         Me.ラベル427.BackColor = RGB(255, 0, 204)
         Me.ラベル319.BackColor = RGB(255, 0, 204)
         Me.ラベル435.BackColor = RGB(255, 0, 204)
         Me.ラベル332.BackColor = RGB(255, 0, 204)
         Me.ラベル442.BackColor = RGB(255, 0, 204)
         Me.ラベル395.BackColor = RGB(255, 0, 204)
         Me.ボックス308.BackColor = RGB(137, 189, 222)
         Me.ボックス432.BackColor = RGB(137, 189, 222)
         Me.ボックス316.BackColor = RGB(137, 189, 222)
         Me.ボックス411.BackColor = RGB(137, 189, 222)
         Me.ボックス331.BackColor = RGB(137, 189, 222)
         Me.ボックス441.BackColor = RGB(137, 189, 222)
         Me.ボックス394.BackColor = RGB(137, 189, 222)
         Me.フォームヘッダー.BackColor = RGB(137, 189, 222)
         
         
         
    Case "T"
         Me.TX_SYSTN = "横浜鋼管管理システム"
    Case Else
         Me.TX_SYSTN = "管理システム"                                           'REP 2006/03/24 Tajima 元は"表示エラー"
    End Select                                                                  'INSERT 20030815 N.MIURA END
    
    Call sub権限によるボタン制御  '2007/07/14 add tajima
    
    doCmd.Maximize
'''''    Call CmWindowSize
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS410_Click
'        CREATE_DATE      :2006/03/24 Tajima
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS410_Click()
    Call MSHS010_M00("FVS410", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS410"
End Sub

'==============================================================================*
'       2005/10/07  iguchi  ADD
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS530_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS530_Click()
    Call MSHS010_M00("FVS530", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS530"
End Sub

'==============================================================================*
'       2005/11/07  iguchi  ADD
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS540_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS540_Click()
    Call MSHS010_M00("FVS540", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS540"
End Sub

'==============================================================================*
'       2005/08/02   iguchi START
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS700_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS700_Click()
    Call MSHS010_M00("FVS700", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS700"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS710_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS710_Click()
    Call MSHS010_M00("FVS710", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS710"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS750_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS750_Click()
    Call MSHS010_M00("FVS750", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS750"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS760_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS760_Click()
    Call MSHS010_M00("FVS760", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS760"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFVS770_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFVS770_Click()
    Call MSHS010_M00("FVS770", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS770"
End Sub   '2005/08/02   iguchi     END
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd140_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd140_Click()
    
    'doCmd.OpenForm "FKS140_契約一覧表"                                        'DELETE 20041124 N.MIURA
    Call MSHS010_M00("FKS140", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS140"                                                    'INSERT 20041124 N.MIURA

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd150_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmd150_Click()
'    'doCmd.OpenForm "FKS150_更新一覧表"                                         'DELETE 20041124 N.MIURA
'    doCmd.OpenForm "FKS150"                                                     'INSERT 20041124 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS020_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS020_Click()
    
    Call MSHS010_M00("FKS020", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS020_コンテナ予約入力"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS040_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS040_Click()

'    If strCONT_BUMOC = "0" Then                                                'DELETE 20090617 N.MIURA
'      doCmd.OpenForm "FVS602"    'カンバン用の契約入力※旧バージョン           'DELETE 20090617 N.MIURA
'    Else                                                                       'DELETE 20090617 N.MIURA
    Call MSHS010_M00("FVS600", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS600"    '受付対応用の契約入力
'    End If                                                                     'DELETE 20090617 N.MIURA

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS060_Click
'        CREATE_DATE      :2004/12/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS060_Click()
'    doCmd.OpenForm "FKS060"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS090_Click
'        CREATE_DATE      :2005/05/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS090_Click()

    Call MSHS010_M00("FVS500", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS500"
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS025_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS025_Click()
    
'    doCmd.OpenForm "FVS000"                                                    'DELETE 20050730 K.ISHIZAKA

    Call MSHS010_M00("FVS011", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA

    doCmd.OpenForm "FVS011"                                                     'INSERT 20050730 K.ISHIZAKA

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS026_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmdFKS026_Click()                                                  'DELETE 20050730 K.ISHIZAKA
'                                                                               'DELETE 20050730 K.ISHIZAKA
'    doCmd.OpenForm "FVS800"                                                    'DELETE 20050730 K.ISHIZAKA
'                                                                               'DELETE 20050730 K.ISHIZAKA
'End Sub                                                                        'DELETE 20050730 K.ISHIZAKA
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS050_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS050_Click()
    
    'doCmd.OpenForm "FKS050_請求データ修正"
    Call MSHS010_M00("FKS051", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS051"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS010_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS010_Click()
'    doCmd.OpenForm "FKS010_地主契約入力"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS100_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS100_Click()
'    'doCmd.OpenForm "FKS100_地主台帳"                                           'DELETE 20041124 N.MIURA
'    doCmd.OpenForm "FKS100"                                                     'INSERT 20041124 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS110_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS110_Click()
'    'doCmd.OpenForm "FKS110_地主契約更新一覧表"                                 'DELETE 20041124 N.MIURA
'    doCmd.OpenForm "FKS110"                                                     'INSERT 20041124 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS140_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS140_Click()
    
    'doCmd.OpenForm "FKS140_契約一覧表"                                         'DELETE 20041124 N.MIURA
    doCmd.OpenForm "FKS140"                                                     'INSERT 20041124 N.MIURA

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS141_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS141_Click()
'    doCmd.OpenForm "LF0043"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS142_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS142_Click()
'    doCmd.OpenForm "LF0041"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS150_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS150_Click()
'    'doCmd.OpenForm "FKS150_更新一覧表"                                         'DELETE 20041124 N.MIURA
'    doCmd.OpenForm "FKS150"                                                     'INSERT 20041124 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS160_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS160_Click()
'    'doCmd.OpenForm "FKS160_解約一覧表"                                         'DELETE 20041122 N.MIURA
'    doCmd.OpenForm "FKS160"                                                     'INSERT 20041122 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS170_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS170_Click()
'    'doCmd.OpenForm "FKS170_請求明細一覧表"                                     'DELETE 20041124 N.MIURA
'    doCmd.OpenForm "FKS170"                                                     'INSERT 20041124 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS171_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS171_Click()
'    doCmd.OpenForm "LF0038"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS190_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS190_Click()
'    'doCmd.OpenForm "FKS190_預り保証金一覧表"                                   'DELETE 20041124 N.MIURA
'    doCmd.OpenForm "FKS190"                                                     'INSERT 20041124 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMDFKS240_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CMDFKS240_Click()
    
    'doCmd.OpenForm "FKS240_請求明細書"                                         'DELETE 20041124 N.MIURA
    Call MSHS010_M00("FKS240", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS240"                                                     'INSERT 20041124 N.MIURA

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS390_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20080603 N.MIURA
'Private Sub cmdFKS390_Click()
'    'doCmd.OpenForm "FKS390_払込データ作成"                                     'DELETE 20041124 N.MIURA
'    doCmd.OpenForm "FKS390"                                                     'INSERT 20041124 N.MIURA
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS610_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS610_Click()
    
    Call MSHS010_M00("LF0015", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0015"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS611_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS611_Click()
'    doCmd.OpenForm "LF0016"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS612_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS612_Click()
    
'    doCmd.OpenForm "MFLY010"       'DELETE 2011/07/07 M.RYU
    Call MSHS010_M00("LF0052", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0052"         'INSERT 2011/07/07 M.RYU

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS612_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdLF0050_Click()
    
    Call MSHS010_M00("LF0050", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0050"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdLF0049_Click
'        CREATE_DATE      :2005/01/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdLF0049_Click()
'    doCmd.OpenForm "LF0049"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdLFVS510_Click
'        CREATE_DATE      :2005/05/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdLFVS510_Click()
    Call MSHS010_M00("FVS510", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS510"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS008_Click
'        CREATE_DATE      :2003/10/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS008_Click()
'    doCmd.OpenForm "FKS008"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS009_Click
'        CREATE_DATE      :2003/10/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmdFKS009_Click()
'    doCmd.OpenForm "FKS009"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS00M_Click
'        CREATE_DATE      :2003/10/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090618 N.MIURA
'Private Sub cmdFKS00M_Click()
'    doCmd.OpenForm "FKS00M"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS00R_Click
'        CREATE_DATE      :2006/06/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Private Sub cmdFKS00R_Click()
'    doCmd.OpenForm "FKS00R"
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_ヤード解約業務_Click
'        CREATE_DATE      :2007/05/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_ヤード解約業務_Click()
    Call MSHS010_M00("FVS210", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS210"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ヤード修繕
'        MODULE_ID        :cmd_ヤード修繕_Click
'        CREATE_DATE      :2009/06/17 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_ヤード修繕_Click()
    Call MSHS010_M00("FVS901", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS901"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :部屋修繕
'        MODULE_ID        :cmd_部屋修繕_Click
'        CREATE_DATE      :2009/06/17 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_部屋修繕_Click()
    Call MSHS010_M00("FVS903", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS903"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客問合せ
'        MODULE_ID        :cmd_顧客問合せ_Click
'        CREATE_DATE      :2009/06/17 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_顧客問合せ_Click()
    Call MSHS010_M00("FVS902", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS902"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :メンテナンス依頼書出力
'        MODULE_ID        :cmd_メンテナンス依頼書出力_Click
'        CREATE_DATE      :2009/06/17 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_メンテナンス依頼書出力_Click()
    Call MSHS010_M00("FVS920", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FVS920"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :契約状況一覧表
'        MODULE_ID        :cmd600_Click
'        CREATE_DATE      :2009/06/18 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd600_Click()
    Call MSHS010_M00("LF0000", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0000"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :分析表
'        MODULE_ID        :cmd602_Click
'        CREATE_DATE      :2009/06/18 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd602_Click()
    Call MSHS010_M00("LF0003", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0003"
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :坪単価一覧表
'        MODULE_ID        :cmd604_Click
'        CREATE_DATE      :2009/06/18 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd604_Click()
    
    Call MSHS010_M00("LF0005", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0005"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :稼動明細一覧表
'        MODULE_ID        :cmd605_Click
'        CREATE_DATE      :2009/06/18 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd605_Click()
    
    Call MSHS010_M00("LF0007", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0007"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :契約・解約集計表
'        MODULE_ID        :cmd607_Click
'        CREATE_DATE      :2009/06/18 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd607_Click()
    
    Call MSHS010_M00("LF0001", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0001"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :予約集計表
'        MODULE_ID        :cmd_予約集計表_Click
'        CREATE_DATE      :2009/06/18 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_予約集計表_Click()
    
    Call MSHS010_M00("LF0008", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "LF0008"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :顧客マスタ保守
'        MODULE_ID        :cmdFKS430_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS430_Click()
    Call MSHS010_M00("FKS430", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS430"                                                     'INSERT 20041106 N.MIURA

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタ保守
'        MODULE_ID        :cmdFKS440_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS440_Click()
    Call MSHS010_M00("FKS440", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS440_ヤードマスタ保守"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :セキュリティカードマスタメンテ
'        MODULE_ID        :cmdFKS600_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS600_Click()
    Call MSHS010_M00("FKS600", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS600"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ボックスマスタメンテ
'        MODULE_ID        :cmdFKS800_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS800_Click()
    Call MSHS010_M00("FKS800", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS800"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :割引マスタメンテナンス
'        MODULE_ID        :cmdFKS700_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS700_Click()
    Call MSHS010_M00("FKS700", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS700"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ヤード別割引設定
'        MODULE_ID        :cmdFKS720_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
Private Sub cmdFKS720_Click()
    Call MSHS010_M00("FKS720", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS720"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS480_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS480_Click()
    Call MSHS010_M00("FKS480", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS480"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmdFKS490_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS490_Click()
    Call MSHS010_M00("FKS490", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS490"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :近隣ヤードマスタ
'        MODULE_ID        :cmdFKS540_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS540_Click()
    Call MSHS010_M00("FKS540", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS540"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ボックスマスタリスト
'        MODULE_ID        :cmdFKS820_Click
'        CREATE_DATE      :2009/06/18 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS820_Click()
    Call MSHS010_M00("FKS820", "M", strCONT_BUMOC)                                'INSERT 20130416 M.HONDA
    doCmd.OpenForm "FKS820"

End Sub

'==============================================================================*
'
'        MODULE_NAME      :テンプレート一括設定
'        MODULE_ID        :cmdFKS450_Click
'        CREATE_DATE      :2017/11/01 tajima
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS450_Click()
    Call MSHS010_M00("FKS450", "M", strCONT_BUMOC)
    doCmd.OpenForm "FKS450"

End Sub

'==============================================================================*
'
'        MODULE_NAME      :価格テンプレートマスタメンテ
'        MODULE_ID        :cmdFKS460_Click
'        CREATE_DATE      :2017/11/09   ysuzuki
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdFKS460_Click()
    Call MSHS010_M00("FKS460", "M", strCONT_BUMOC)
    doCmd.OpenForm "FKS460"

End Sub

'==============================================================================*
'
'        MODULE_NAME      :業務終了
'        MODULE_ID        :cmdSysEnd_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdSysEnd_Click()

    'Call CmSysEnd("コンテナ管理システム")                                      'DELETE 20041124 N.MIURA
'    Call CmSysEnd("管理システム")                                               'INSERT 20041124 N.MIURA   'DELETE 2013/09/30 K.KINEBUCHI

    doCmd.Quit acQuitSaveAll                                                    'INSERT 2013/09/30 K.KINEBUCHI
  

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2004/12/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    Call MSHS010_M00(PROG_ID, "2")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    
    doCmd.Maximize
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            KeyCode = 0
        Case Else
    End Select

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :sub権限によるボタン制御
'        CREATE_DATE      :2007/07/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub権限によるボタン制御()

    Dim isException     As Boolean
    Dim adoDbConnection As Object       'ADODB.Connection
    Dim rsData          As Object       'ADODB.Recordset
    Dim sqlText         As String
    Dim userId          As String
    
    isException = False
On Error GoTo Exception

    ' 画面の全ボタンを無効化
'    Call setAllButtonUse(False)                                               'DELETE 20090617 N.MIURA

    'ユーザ情報取得
    userId = MSZZ000.LsGetUserName()

    'ADO接続Object生成
    Set adoDbConnection = MSZZ025.ADODB_Connection()

    sqlText = "SELECT * "
    sqlText = sqlText & " FROM ONLE_TABL "
    sqlText = sqlText & " WHERE ONLET_USERB = '" & userId & "' "
    sqlText = sqlText & "  AND ONLET_ACCEI <> '9' "

    Set rsData = MSZZ025.ADODB_Recordset(sqlText, adoDbConnection, adoReadOnly)
    
Dim wk As String
    
    Do Until rsData.EOF
        ' ボタンの使用可否を設定する
wk = wk & " " & Nz(rsData.Fields("ONLET_PROGB"), "")
        Call setButtonUse(Nz(rsData.Fields("ONLET_PROGB"), ""))
        rsData.MoveNext
    Loop
Debug.Print wk

    '2019/03/25 M.HONDA INS
    '部門によるメニュー制御
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    If strCONT_BUMOC = "H" Then
        'Me.cmdLFVS510.Visible = False       '予約受付リスト
       ' Me.cmdFVS300.Visible = False        '毎月回収方法進捗管理
        Me.cmdFVS750.Visible = False        '延長リスト
        Me.cmd600.Visible = False           '契約状況一覧表
        Me.cmd602.Visible = False           '分析表
        Me.cmd604.Visible = False           '坪単価一覧表
        Me.cmd_予約集計表.Visible = False   '予約集計表
        Me.cmd607.Visible = False           '契約・解約集計表
        Me.cmdFVS100.Visible = False        '集計件数データ出力
        Me.cmd_ヤード修繕.Visible = False   'ヤード修繕
        Me.cmd_顧客問合せ.Visible = False   '顧客問合せ
        Me.cmdFKS480.Visible = False        '顧客リスト
        Me.cmdFKS540.Visible = False        '近隣ヤードリスト
        Me.cmdFKS820.Visible = False        'ボックスリスト
    Else
        Me.cmdLF0060.Visible = False        'コンテナ在庫表 INSERT 2020/05/23 N.IMAI
    End If

    
    GoTo Finally
Exception:
    isException = True
Finally:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing
    If isException Then
        Call Err.Raise(Err.Number, "sub権限によるボタン制御" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :ボタンの使用可否設定
'        CREATE_DATE      :2007/07/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Sub setButtonUse(ByVal programID As String)

    Select Case programID
    '★予約処理系
        '空き検索
        Case "FVS011"
              Me.cmdFKS025.Enabled = True
        '予約入力
        Case "FVS500"
              Me.cmdFKS090.Enabled = True
        '予約受付リスト
        Case "FVS510"
              Me.cmdLFVS510.Enabled = True
        '予約一覧
        Case "FVS540"
              Me.cmdFVS540.Enabled = True
        'ご紹介連絡表
        Case "FVS530"
              Me.cmdFVS530.Enabled = True
        '送り状リスト作成
        Case "FVS410"
              Me.cmdFVS410.Enabled = True
        'メンテ入力（仮）
'        Case "FKS060"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS060.Enabled = True                                     'DELETE 20090617 N.MIURA
    '★契約処理系
        '利用状況
        Case "FKS020"
              Me.cmdFKS020.Enabled = True
        '契約入力
        Case "FVS600"
              Me.cmdFKS040.Enabled = True
        '契約一覧表
        Case "FKS140"
              Me.cmd140.Enabled = True
        '契約一覧表２
'        Case "LF0043"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS141.Enabled = True                                     'DELETE 20090617 N.MIURA
        '更新一覧表
'        Case "FKS150"                                                         'DELETE 20090617 N.MIURA
'              Me.cmd150.Enabled = True                                        'DELETE 20090617 N.MIURA
        '解約一覧表
'        Case "FKS160"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS160.Enabled = True                                     'DELETE 20090617 N.MIURA
        '契約日報リスト
'        Case "LF0041"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS142.Enabled = True                                     'DELETE 20090617 N.MIURA
    '★解約処理系
        '解約入力
        Case "FVS700"
              Me.cmdFVS700.Enabled = True
        '解約一覧入力
        Case "FVS710"
              Me.cmdFVS710.Enabled = True
        '解約メンテナンス表
        Case "FVS760"
              Me.cmdFVS760.Enabled = True
        '延長リスト
        Case "FVS750"
              Me.cmdFVS750.Enabled = True
        '精算依頼リスト
        Case "FVS770"
              Me.cmdFVS770.Enabled = True
        'ヤード解約業務
        Case "FVS210"
              Me.cmd_ヤード解約業務.Enabled = True
    '★作表処理系
'                                                                              'INSERT 20090618 N.MIURA START
       Case "LF0000": Me.cmd600.Enabled = True         '契約状況一覧表
       Case "LF0003": Me.cmd602.Enabled = True         '分析表
       Case "LF0005": Me.cmd604.Enabled = True         '坪単価一覧表
       Case "LF0007": Me.cmd605.Enabled = True         '稼動明細一覧表
       Case "LF0008": Me.cmd_予約集計表.Enabled = True '予約集計表
       Case "LF0001": Me.cmd607.Enabled = True         '契約・解約集計表

       Case "LF0060": Me.cmdLF0060.Enabled = True      'コンテナ在庫表         'INSERT 2020/05/23 N.IMAI
'                                                                              'INSERT 20090618 N.MIURA ENDED
       '' 20091026 M.Honda START
        Case "FVS100"
              Me.cmdFVS100.Enabled = True              '' 予約・契約・解約集計表
       '' 20091026 M.Honda END
        '預り保証金一覧表
'        Case "FKS190"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS190.Enabled = True                                     'DELETE 20090617 N.MIURA
        '宛名印刷（ヤード）
        Case "LF0015"
              Me.cmdFKS610.Enabled = True
        '宛名印刷（満了日）
'        Case "LF0016"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS611.Enabled = True                                     'DELETE 20090617 N.MIURA
        '宛名印刷
        Case "MFLY010"
              Me.cmdFKS612.Enabled = True
        '契約データ出力
        Case "LF0050"
              Me.cmdLF0050.Enabled = True
        '修繕データ出力
'        Case "LF0049"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdLF0049.Enabled = True                                     'DELETE 20090617 N.MIURA
    '★請求入金処理系
        '請求データ修正入力
        Case "FKS051"
              Me.cmdFKS050.Enabled = True
        '払込表データ作成
        'Case "FKS390"                                                         'DELETE 20080603 N.MIURA
        '      Me.cmdFKS390.Enabled = True                                     'DELETE 20080603 N.MIURA
        '請求明細一覧表
'        Case "FKS170"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS170.Enabled = True                                     'DELETE 20090617 N.MIURA
        '売上リスト
'        Case "LF0038"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS171.Enabled = True                                     'DELETE 20090617 N.MIURA
        '請 求 書
        Case "FKS240"
              Me.cmdFKS240.Enabled = True
    '★地主契約
        '地 主 契 約 入 力
'        Case "FKS010"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS010.Enabled = True                                     'DELETE 20090617 N.MIURA
        '地  主  台  帳
'        Case "FKS100"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS100.Enabled = True                                     'DELETE 20090617 N.MIURA
        '地主契約更新一覧表
'        Case "FKS110"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS110.Enabled = True                                     'DELETE 20090617 N.MIURA
    '★修繕処理                                                                'INSERT 20090617 N.MIURA START
       Case "FVS901": Me.cmd_ヤード修繕.Enabled = True             'ヤード修繕
       Case "FVS903": Me.cmd_部屋修繕.Enabled = True               'コンテナ（部屋）修繕
       Case "FVS902": Me.cmd_顧客問合せ.Enabled = True             '顧客問合
       Case "FVS920": Me.cmd_メンテナンス依頼書出力.Enabled = True '修理依頼書出力
       Case "FVS904"
            Me.cmd_ヤード一括修繕.Enabled = True                   'ヤード一括修繕　INSERT 2017-11-02 add k.sato
'                                                                              'INSERT 20090617 N.MIURA ENDED
    '★その他メニュー
        '分析帳票
'        Case "FKS009"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS009.Enabled = True                                     'DELETE 20090617 N.MIURA
        '損益管理
'        Case "FKS008"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS008.Enabled = True                                     'DELETE 20090617 N.MIURA
        'マスタ保守
'        Case "FKS00M"                                                         'DELETE 20090618 N.MIURA
'              Me.cmdFKS00M.Enabled = True                                     'DELETE 20090618 N.MIURA
        '修繕管理
'        Case "FKS00R"                                                         'DELETE 20090617 N.MIURA
'              Me.cmdFKS00R.Enabled = True                                     'DELETE 20090617 N.MIURA
    '★マスタ保守                                                              'INSERT 20090618 N.MIURA START
       Case "FKS430": Me.cmdFKS430.Enabled = True '顧客マスタ保守
       Case "FKS440": Me.cmdFKS440.Enabled = True 'ヤードマスタ保守
       Case "FKS600": Me.cmdFKS600.Enabled = True 'セキュリティカード保守
       Case "FKS800": Me.cmdFKS800.Enabled = True 'ボックスマスタ保守
       Case "FKS700": Me.cmdFKS700.Enabled = True '割引マスタ保守
       Case "FKS720": Me.cmdFKS720.Enabled = True 'ヤード別割引設定
       Case "FKS480": Me.cmdFKS480.Enabled = True '顧客リスト
       Case "FKS490": Me.cmdFKS490.Enabled = True 'ヤードリスト
       Case "FKS540": Me.cmdFKS540.Enabled = True '近隣ヤードリスト
       Case "FKS820": Me.cmdFKS820.Enabled = True 'ボックスリスト
'                                                                              'INSERT 20090618 N.MIURA ENDED
       '毎月回収方法進捗管理
       Case "FVS300": Me.cmdFVS300.Enabled = True                              'INSERT 20110310 M.HONDA

       Case "FKS830": Me.cmdFKS830.Enabled = True                              'INSERT 2014/03/03 INS M.HONDA

       Case "FKS450": Me.cmdFKS450.Enabled = True                              'INSERT 2017/11/01 INS tajima
       
       Case "FKS460": Me.cmdFKS460.Enabled = True                              'INSERT 2017/11/09 INS ysuzuki

    End Select
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :全ボタンの使用可否を制御
'        CREATE_DATE      :2007/07/14
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20090617 N.MIURA
'Sub setAllButtonUse(ByVal isEnabled As Boolean)
'      '空き検索
'      Me.cmdFKS025.Enabled = isEnabled
'      '予約入力
'      Me.cmdFKS090.Enabled = isEnabled
'      '予約受付リスト
'      Me.cmdLFVS510.Enabled = isEnabled
'      '予約一覧
'      Me.cmdFVS540.Enabled = isEnabled
'      'ご紹介連絡表
'      Me.cmdFVS530.Enabled = isEnabled
'      '送り状リスト作成
'      Me.cmdFVS410.Enabled = isEnabled
'      'メンテ入力（仮）これは特に使用しない
'      'Me.cmdFKS060.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '利用状況
'      Me.cmdFKS020.Enabled = isEnabled
'      '契約入力
'      Me.cmdFKS040.Enabled = isEnabled
'      '契約一覧表１
'      Me.cmd140.Enabled = isEnabled
'      '契約一覧表２
''      Me.cmdFKS141.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '更新一覧表
''      Me.cmd150.Enabled = isEnabled                                           'DELETE 20090617 N.MIURA
'      '解約一覧表
''      Me.cmdFKS160.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '契約日報リスト
''      Me.cmdFKS142.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '解約入力
'      Me.cmdFVS700.Enabled = isEnabled
'      '解約一覧入力
'      Me.cmdFVS710.Enabled = isEnabled
'      '解約メンテナンス表
'      Me.cmdFVS760.Enabled = isEnabled
'      '延長リスト
'      Me.cmdFVS750.Enabled = isEnabled
'      '精算依頼リスト
'      Me.cmdFVS770.Enabled = isEnabled
'      'ヤード解約業務
'      Me.cmd_ヤード解約業務.Enabled = isEnabled
'      '預り保証金一覧表
''      Me.cmdFKS190.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '宛名印刷（ヤード）
'      Me.cmdFKS610.Enabled = isEnabled
'      '宛名印刷（満了日）
''      Me.cmdFKS611.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '宛名印刷
'      Me.cmdFKS612.Enabled = isEnabled
'      '契約データ出力
'      Me.cmdLF0050.Enabled = isEnabled
'      '修繕データ出力
''      Me.cmdLF0049.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '請求データ修正入力
'      Me.cmdFKS050.Enabled = isEnabled
'      '払込表データ作成
'      'Me.cmdFKS390.Enabled = isEnabled                                        'DELETE 20080603 N.MIRUA
'      '請求明細一覧表
''      Me.cmdFKS170.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '売上リスト
''      Me.cmdFKS171.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '請 求 書
'      Me.cmdFKS240.Enabled = isEnabled
'      '地 主 契 約 入 力
''      Me.cmdFKS010.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '地  主  台  帳
''      Me.cmdFKS100.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '地主契約更新一覧表
''      Me.cmdFKS110.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '分析帳票
''      Me.cmdFKS009.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      '損益管理
''      Me.cmdFKS008.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'      'マスタ保守
'      Me.cmdFKS00M.Enabled = isEnabled
'      '修繕管理
''      Me.cmdFKS00R.Enabled = isEnabled                                        'DELETE 20090617 N.MIURA
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :部屋一括修繕
'        MODULE_ID        :cmd_部屋一括修繕_Click
'        CREATE_DATE      :2017-11-05 k.sato
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2017-11-05 add k.sato
Private Sub cmd_ヤード一括修繕_Click()
    Call MSHS010_M00("FVS904", "M", strCONT_BUMOC)
    doCmd.OpenForm "FVS904"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :コンテナ在庫表
'        MODULE_ID        :cmdLF0060_Click
'        CREATE_DATE      :2020/05/23 N.IMAI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdLF0060_Click()
    
    Call MSHS010_M00("LF0060", "M", strCONT_BUMOC)
    doCmd.OpenForm "LF0060"

End Sub
'****************************  ended or program ********************************
