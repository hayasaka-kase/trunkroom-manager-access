VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS540_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約一覧サブ(一覧)
'   プログラムＩＤ　：　FVS540_S01
'   作　成　日　　　：  2005/12/20
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/1/24
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.1
'   修　正　内　容　：  契約入力画面に遷移するボタン追加
'
'   修　正　日　　　：  2006/03/09　2006/04/06
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.2
'   修　正　内　容　：  1.契約入力画面への遷移を受付入力画面への遷移に変更(3/9)
'                  :   2.鍵配送状況を表示（デザインのみ変更)(4/6)
'
'   修　正　日　　　：  2007/08/23
'   修　正　者　　　：  tajima
'   Ver             ：  0.3
'   修　正　内　容　： 「受付種別」「申込書送付方」「受付番号」追加（デザインのみ変更）
'
'   修　正　日　　　：  2015/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  1.3
'   修　正　内　容　：  項目に身分証明書確認済みを追加
'
'   修　正　日　　　：  2022/08/22
'   修　正　者　　　：  N.IMAI
'   Ver             ：  1.6
'   修　正　内　容　：  検索条件追加：初期費用請求方法、入金確認日、メール発行日
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Const P_ご紹介画面          As String = "FVS550"
Private Const P_予約受付画面        As String = "FVS500"

'06.1.24 SHIBAZAKI 追加
'2006/03/09 SHIBAZAKI 削除
'Private Const P_契約入力画面             As String = "FVS600"
'Private Const P_契約入力パラメータ       As String = "9"
'2006/03/09 SHIBAZAKI 追加
Private Const P_受付入力画面             As String = "FVS400"
Private Const P_BOOKING_STATUS_ACCEPT   As Integer = 10 '予約受付状態区分「受付」

'ご紹介区分
Private Const P_INTRO_STATUS_KEEP           As String = "1" 'ご紹介区分「取り置きした」
Private Const P_INTRO_STATUS_CONTRACT       As String = "2" 'ご紹介区分「契約受付した」
Private Const P_INTRO_STATUS_NOT_INTRO      As String = "9" 'ご紹介区分「再度ご紹介不要」

'==============================================================================*
'
'        MODULE_NAME      :cmd_Gosyoukai_Click
'        機能             :ご紹介ボタンクリックイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Gosyoukai_Click()

    '画面項目無効化
    '06.1.24 SHIBAZAKI 変更
    'Call Parent.FormEnable(False)
    'cmd_Yoyaku.Enabled = False
    Call ScreenChangeStandby(False, False, False)
    
    '予約受付表示
    'パラメータはプログラムID＋受付番号
    doCmd.OpenForm P_ご紹介画面, , , , , , Parent.NAME & "," & txt_Ukno

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Uketsuke_Click
'        機能             :契約入力ボタンクリックイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Uketsuke_Click()

    If Val(Nz(txt_Ukkbn, 0)) <> P_BOOKING_STATUS_ACCEPT Then
        MsgBox "予約状態が「受付」でなければ、受付入力画面に移動できません。"
        Exit Sub
    End If
    
    '画面項目無効化
    Call ScreenChangeStandby(False, False, False)
    
'2006/03/09 契約画面じゃなく受付画面に遷移する
'    '予約受付表示
'    'パラメータは区分＋受付番号
'    doCmd.OpenForm P_契約入力画面, , , , , , P_契約入力パラメータ & txt_Ukno
    'パラメータは予約番号＋受付番号＋移動元契約番号
    doCmd.OpenForm P_受付入力画面, , , , , , Parent.NAME & "," & txt_Ukno & "," & txt_RcptNo & "," & txt_MotoAcptNo

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Yoyaku_Click
'        機能             :予約受付ボタンクリックイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_YOYAKU_Click()

    '画面項目無効化
    '06.1.24 SHIBAZAKI 変更
'    Call Parent.FormEnable(False)
'    cmd_Gosyoukai.Enabled = False
    Call ScreenChangeStandby(False, False, False)
    
    '予約受付表示
    'パラメータはヤードコード＋コンテナ番号（NULLの場合は空白６文字）＋受付番号＋プログラムID
    doCmd.OpenForm P_予約受付画面, , , , , , _
        Format(txt_YCODE, "000000") + _
        IIf(Nz(txt_Cno) = "", "      ", Format(txt_Cno, "000000")) + _
        txt_Ukno + _
        Parent.NAME

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
Private Sub ScreenChangeStandby(blnGosyoukaiEnable As Boolean, _
                                blnYoyakuEnable As Boolean, _
                                blnKeiyakuEnable As Boolean)

    '画面項目無効化
    txt_Focus.Visible = True
    txt_Focus.SetFocus
    Call Parent.FormEnable(False)
    cmd_Gosyoukai.Enabled = blnGosyoukaiEnable
    cmd_YOYAKU.Enabled = blnYoyakuEnable
    cmd_Uketsuke.Enabled = blnKeiyakuEnable

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
