VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS550_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　取置き履歴(一覧)
'   プログラムＩＤ　：　FVS550_S01
'   作　成　日　　　：  2005/12/20
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/03/09
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.1
'   修　正　内　容　：  1.受付時に受付入力画面に遷移する
'
'   修　正　日　　　：  2006/09/11
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.2
'   修　正　内　容　：  1.営業期間を追加（デザインのみ変更）
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

'ご紹介区分
Private Const P_INTRO_STATUS_KEEP           As String = "1" 'ご紹介区分「取り置きした」
Private Const P_INTRO_STATUS_ACCEPT         As String = "2" 'ご紹介区分「受付した」
Private Const P_INTRO_STATUS_CANCEL         As String = "6" 'ご紹介区分「キャンセルした」
Private Const P_INTRO_STATUS_NOT_INTRO      As String = "9" 'ご紹介区分「再度ご紹介不要」

'==============================================================================*
'
'        MODULE_NAME      :cmd_Fuyou_Click
'        機能             :紹介不要ボタンクリックイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Fuyou_Click()

    If txt_Introkbn = P_INTRO_STATUS_NOT_INTRO Then
        MsgBox "既に紹介不要です。"
        Exit Sub
    End If
    
    If txt_Introkbn <> P_INTRO_STATUS_KEEP And _
       txt_Introkbn <> P_INTRO_STATUS_CANCEL _
    Then
        MsgBox "「取置き」または「キャンセル」ではなければ、「紹介不要」には出来ません。"
        Exit Sub
    End If
    
    '紹介区分を「再度ご紹介不要」に更新
    If Parent.UpdateProc(Val(txt_IntroNo), P_INTRO_STATUS_NOT_INTRO, txt_YCODE, txt_Cno, txt_Introkbn) Then
        Call Parent.AllReload
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_KeiUke_Click
'        機能             :契約受付ボタンクリックイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_KeiUke_Click()

    If txt_Introkbn <> P_INTRO_STATUS_KEEP Then
        MsgBox "「取置き」ではなければ、「受付」には出来ません。"
        Exit Sub
    End If

    '紹介区分を「契約受付した」に更新　＆　予約受付状態を「契約受付」に更新
    If Parent.UpdateProc(Val(txt_IntroNo), P_INTRO_STATUS_ACCEPT, txt_YCODE, txt_Cno, txt_Introkbn) Then
        '2006/03/09 SHIBAZAKI 受付画面に遷移する
        doCmd.OpenForm "FVS400", , , , , , Parent.strFormFrom & "@," & txt_Ukno & ",,"
        Parent.blnReloadSkip = True '呼び出しもとのリフレッシュは行わない
        doCmd.Close acForm, Parent.FormName
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Torioki_Click
'        機能             :取り置きボタンクリックイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Torioki_Click()

    If txt_KeepFlag = "0" Then
        MsgBox "紹介待ちデータではありません。"
        Exit Sub
    End If
    
    '紹介区分を「取り置きした」に更新　＆　予約受付状態を「予約取置き」に更新
    If Parent.UpdateProc(Val(txt_IntroNo), P_INTRO_STATUS_KEEP, txt_YCODE, txt_Cno, txt_Introkbn) Then
        Call Parent.AllReload
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SubFormButtonEnable
'        機能             :ボタンの有効／無効を設定する
'        IN               :True=有効にする False=無効にする
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SubFormButtonEnable(blnTorioki As Boolean, blnKeiUke As Boolean, blnFuyou As Boolean)

    cmd_Torioki.Enabled = blnTorioki
    cmd_KeiUke.Enabled = blnKeiUke
    cmd_Fuyou.Enabled = blnFuyou
    
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
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

