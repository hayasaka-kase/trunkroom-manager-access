VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FTG021_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : コンテナ管理システム
'       SUB_SYSTEM_NAME : 統合検索
'
'       PROGRAM_NAME    : 統合空詳細 上段一覧
'       PROGRAM_ID      : FTG021_S01
'       PROGRAM_KBN     : FROM
'
'       CREATE          : 2011/02/12
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/03/10
'       UPDATER         : K.ISHZIAKA
'       Ver             : 0.2
'                       : 背景色設定
'                       : 次画面を開く関数に引数追加
'
'       UPDATE          : 2013/02/21
'       UPDATER         : M.HONDA
'       Ver             : 0.3
'                       : 画面項目の並び順を変更
'                       :
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   イベント関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : 予約受付ボタン押下時処理
'       MODULE_ID       : cmdUke_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUke_Click()
    On Error GoTo ErrorHandler
'    Call Parent.ShowFVS500(Format(txt_YOUKT_NO.VALUE, "000000"), Nz(Me.txt_YOUKT_UKNO)) 'DELETE 2011/03/10 K.ISHIZAKA
    Call Parent.ShowFVS500(Format(txt_YOUKT_NO.VALUE, "000000"), Nz(Me.txt_YOUKT_UKNO), _
        txt_CNTA_SIZE.VALUE, "0")                                               'INSERT 2011/03/10 K.ISHIZAKA
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmdUke_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームキー押下時処理
'       MODULE_ID       : Form_KeyDown
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    Call Parent.KeyDown(KeyCode, Shift)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Form_KeyDown", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームロード時処理
'       MODULE_ID       : Form_Load
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Call setUsageBackColor(txt_CNTA_USAGE, _
        txt_YOUKT_YUKBN_NAME, txt_YOUTK_COUNT, txt_TOTAL, txt_YOUKT_TKDATE, _
        txt_PRIC_PRICA, txt_CNTA_PRICE_DIFF, txt_PRIC_FZAPPI, txt_PRIC_EZAPPI, _
        txt_CNTA_HEIGHT, txt_CNTA_WIDTH, txt_CNTA_DEPTH)                        'INSERT 2011/03/10 K.ISHIZAKA
    ' 営業開始日に達してなければ部屋指定の予約ボタンは押せないようにする
    If Nz(Parent.txt_YARD_BEGIN_DAY.VALUE) > Format$(Now, "yyyy/mm/dd") Then
        cmdUke.Enabled = False
    End If
    Me.Filter = "CNTA_STEP=0"
    Me.FilterOn = True
    Me.OrderBy = "CNTA_NO"
    Me.OrderByOn = True
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Form_Load", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_BIKO_DblClick
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_BIKO_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    Call Form_KeyDown(CmReprMod.P_呼出_部屋, 0)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_YOUKT_BIKO_DblClick", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_NAME_DblClick
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_NAME_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    Call Form_KeyDown(CmReprMod.P_呼出_顧客, 0)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_YOUKT_NAME_DblClick", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_NO_DblClick
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_NO_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    Call Form_KeyDown(CmReprMod.P_呼出_部屋, 0)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_YOUKT_NO_DblClick", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YOUKT_RENKBN_DblClick
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YOUKT_RENKBN_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    Call Form_KeyDown(CmReprMod.P_呼出_顧客, 0)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_YOUKT_RENKBN_DblClick", True)
End Sub

'****************************  ended of program ********************************
