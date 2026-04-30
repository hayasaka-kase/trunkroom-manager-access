VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS710"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細）
'   プログラムＩＤ　：　FKS710
'   作　成　日　　　：  2008/07/28
'   作　成　者　　　：  T.Suzuki
'   Ver             ：  0.0
'**********************************************
'   修　正　日　　　：  2008/09/05
'   修　正　者　　　：  K.KINEBUCHI
'   内　　　容　　　：  割引マスタに割引番号を発番する際、割引マスタの件数＋ｎ
'   　　　　　　　　　　で発番していた箇所を、割引番号の最大値＋ｎで発番するよ
'   　　　　　　　　　　うに修正した。
'   　　　　　　　　　　※割引マスタ一覧で削除を行った場合、割引きマスタの件数
'   　　　　　　　　　　　で加算すると、割引番号が重複しエラーとなるため。
'   Ver            ：  0.1
'**********************************************
'   修　正　日　　　：  2008/11/13
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  次画面へ文言を引き数で渡す際に
'   　　　　　　　　　　区切り文字がカンマの場合、不具合があるので
'   　　　　　　　　　　区切り文字をタブに変更
'   Ver            ：  0.2
'**********************************************
'   修　正　日　　　：  2008/11/14
'   修　正　者　　　：  K.KINEBUCHI
'   内　　　容　　　：  サイズの大小比較チェックの不具合を修正する
'   Ver            ：  0.3
'**********************************************
'   修　正　日　　　：  2009/01/14
'   修　正　者　　　：  hirano
'   内　　　容　　　：  opt_割引形式の「n%引き」を非表示にする
'   Ver            ：  0.4
'**********************************************
'   修　正　日　　　：  2009/06/30
'   修　正　者　　　：  hirano
'   内　　　容　　　：  割引マスタ　KASE_DBへ移行に伴う変更
'   Ver            ：  0.5
'**********************************************
'   修　正　日　　　：  2011/08/29
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  割引マスタ　事務手数料割引対応
'   Ver            ：  0.6
'**********************************************
'   修　正　日　　　：  2012/04/05
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  割引マスタ　利用期間制限なし対応
'   Ver            ：  0.7
'**********************************************
'   修　正　日　　　：  2013/02/13
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  割引マスタ　%割引を使用可能へ
'   Ver            ：  0.8
'**********************************************

Option Compare Database
Option Explicit

Private Const PROG_ID = "FKS710"

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_ROOM_WORK_TABLE       As String = "WK_FKS700"

Private Const P_KAKUNIN               As String = "確認"
Private Const P_ERROR                 As String = "エラー"

Private Const MODE_INSERT             As Integer = 1          ' 登録モード
Private Const MODE_UPDATE             As Integer = 2          ' 修正モード

'コントロールの種類
Private Const P_CONTROL_数値          As Integer = 1          ' 数値
Private Const P_CONTROL_テキスト      As Integer = 3          ' テキスト
Private Const P_CONTROL_金額          As Integer = 6          ' 金額

Private Const P_NEBIKI_KIKAN_LENGTH   As Integer = 2          ' 割引期間
Private Const P_NEBIKI_VALUE_LENGTH   As Integer = 9          ' 割引値
Private Const P_NEBIKI_VALUE_LENGTH2  As Integer = 2          ' %
Private Const P_UKETSUKE_KIND_LENGTH  As Integer = 2          ' 受付種別
Private Const P_SIZE_LENGTH           As Integer = 5          ' サイズ
Private Const P_STEP_LENGTH           As Integer = 1          ' 段
Private Const P_YOTO_LENGTH           As Integer = 2          ' 用途
Private Const P_RIYO_KIKAN_LENGTH     As Integer = 1          ' 利用期間
Private Const P_MONGON_LENGTH         As Integer = 42         ' お客様提示文言

Private Const P_SIZE_MAX_VALUE        As Single = 999.99      ' サイズ

Private Const P_MINUS                 As Integer = 45         ' マイナス符号

Private Const P_ALL_NINE              As String = "9999999999"

'不足桁０埋め用
Private Const P_ZERO_SUPPRESS         As String = "0000000000"


Private Const P_FMT_STANDARD          As String = "Standard"  ' 書式：標準
Private Const P_FMT_FIXED             As String = "Fixed"     ' 書式：固定
Private pobjDB                        As Object

Private pstrNEBIKI_NO                 As String               ' 割引番号
Private pstrBumonCode                 As String               ' 部門コード
Private pstrProcMode                  As String               ' 処理モード

Private pboolFrmActive                As Boolean
Private pstrParameter()               As String               ' 初期表示データ

Private var変更日付                   As Variant              '日付入力チェック用
'==============================================================================*
'
'        MODULE_NAME      :chk_サイズ指定_AfterUpdate
'        MODULE_ID        :chk_サイズ指定_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_サイズ指定_AfterUpdate()

    Select Case Me.chk_サイズ指定.VALUE
        Case "-1"
            ' *** チェックあり ***
            Me.txt_サイズFrom.Enabled = True
            Me.txt_サイズTo.Enabled = True

        Case "0"
            ' *** チェックなし ***
            Me.txt_サイズFrom.Enabled = False
            Me.txt_サイズTo.Enabled = False
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_段指定_AfterUpdate
'        MODULE_ID        :chk_段指定_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_段指定_AfterUpdate()

    Select Case Me.chk_段指定.VALUE
        Case "-1"
            ' *** チェックあり ***
            Me.cmb_段.Enabled = True

        Case "0"
            ' *** チェックなし ***
            Me.cmb_段.Enabled = False
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_用途指定_AfterUpdate
'        MODULE_ID        :chk_用途指定_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_用途指定_AfterUpdate()

    Select Case Me.chk_用途指定.VALUE
        Case "-1"
            ' *** チェックあり ***
            Me.cmb_用途.Enabled = True

        Case "0"
            ' *** チェックなし ***
            Me.cmb_用途.Enabled = False
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_AfterUpdate
'        MODULE_ID        :cmb_受付種別_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_AfterUpdate()

    Me.txt_受付種別 = fncAfterUpdateCommon(Me.cmb_受付種別, P_UKETSUKE_KIND_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_BeforeUpdate
'        MODULE_ID        :cmb_受付種別_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_受付種別, P_UKETSUKE_KIND_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_KeyPress
'        MODULE_ID        :cmb_受付種別_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_受付種別, KeyAscii, P_UKETSUKE_KIND_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_段_AfterUpdate
'        MODULE_ID        :cmb_段_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_段_AfterUpdate()

    Me.txt_段 = fncAfterUpdateCommon(Me.cmb_段, P_STEP_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_段_BeforeUpdate
'        MODULE_ID        :cmb_段_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_段_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_段, P_STEP_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_段_KeyPress
'        MODULE_ID        :cmb_段_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_段_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_段, KeyAscii, P_STEP_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_割引期間_AfterUpdate
'        MODULE_ID        :cmb_割引期間_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_割引期間_AfterUpdate()

    Me.txt_割引期間 = fncAfterUpdateCommon(Me.cmb_割引期間, P_NEBIKI_KIKAN_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_割引期間_BeforeUpdate
'        MODULE_ID        :cmb_割引期間_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_割引期間_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_割引期間, P_NEBIKI_KIKAN_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_割引期間_KeyPress
'        MODULE_ID        :cmb_割引期間_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_割引期間_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_割引期間, KeyAscii, P_NEBIKI_KIKAN_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_AfterUpdate
'        MODULE_ID        :cmb_用途_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_AfterUpdate()

    Me.txt_用途 = fncAfterUpdateCommon(Me.cmb_用途, P_YOTO_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_BeforeUpdate
'        MODULE_ID        :cmb_用途_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_用途, P_YOTO_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_KeyPress
'        MODULE_ID        :cmb_用途_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_用途, KeyAscii, P_YOTO_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_利用期間_AfterUpdate
'        MODULE_ID        :cmb_利用期間_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用期間_AfterUpdate()

    Me.txt_利用期間 = fncAfterUpdateCommon(Me.cmb_利用期間, P_RIYO_KIKAN_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_利用期間_BeforeUpdate
'        MODULE_ID        :cmb_利用期間_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用期間_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_利用期間, P_RIYO_KIKAN_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_利用期間_KeyPress
'        MODULE_ID        :cmb_利用期間_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用期間_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_利用期間, KeyAscii, P_RIYO_KIKAN_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_文言自動生成_Click
'        MODULE_ID        :cmd_文言自動生成_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_文言自動生成_Click()
    
    On Error GoTo Exception
    
    ' 入力チェック
    If fncInputCheck(False) = True Then
        ' 入力に問題がなければ文言自動生成
        Me.txt_文言自動生成.VALUE = fncCreateAutoText()
        ' さらに「お客様提示文言」が未入力なら自動生成を設定する
        If Nz(Me.txt_お客様提示文言.VALUE, "") = "" Then
            Me.txt_お客様提示文言.VALUE = Me.txt_文言自動生成.VALUE
        End If
    End If
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmd_保存_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opg_割引形式_AfterUpdate
'        MODULE_ID        :opg_割引形式_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_割引形式_AfterUpdate()

    ' 割引値をクリア
    Me.txt_割引値.VALUE = ""
    '----↓↓↓↓----20110829--M.HONDA--Ins-------↓↓↓↓---<s>
    Me.cmd_文言自動生成.Enabled = True
    Me.txt_お客様提示文言.Enabled = True
    Me.txt_文言自動生成 = ""
    Me.txt_お客様提示文言 = ""
    '----↑↑↑↑----20110829--M.HONDA--Ins-------↑↑↑↑---<e>

    ' 割引種別ガイダンスリセット
    Select Case Me.opg_割引形式.VALUE
        Case "1"
            ' *** n円指定 ***
            Me.lbl_基本料金.Caption = "基本料金を"
            Me.lbl_円.Caption = "円とする"

        Case "2"
            ' *** n円割引 ***
            Me.lbl_基本料金.Caption = "基本料金から"
            Me.lbl_円.Caption = "円引きとする"

        Case "3"
            ' *** ｎ％割引 ***
            Me.lbl_基本料金.Caption = "基本料金から"
            Me.lbl_円.Caption = "％引きとする"

        '----↓↓↓↓----20110829--M.HONDA--Ins-------↓↓↓↓---<s>
        Case "10"
            ' *** 事務手数料割引 ***
            Me.lbl_基本料金.Caption = "事務手数料を"
            Me.lbl_円.Caption = "円とする"
            Me.txt_割引値.VALUE = 0
            Me.cmd_文言自動生成.Enabled = False
            Me.txt_お客様提示文言.Enabled = False
            Me.txt_文言自動生成 = "事務手数料割引"
            Me.txt_お客様提示文言 = "事務手数料割引"
        '----↑↑↑↑----20110829--M.HONDA--Ins-------↑↑↑↑---<e>

        Case Else
            ' *** 未選択 ***
            Me.lbl_基本料金.Caption = "基本料金を"
            Me.lbl_円.Caption = "円とする"
    End Select
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_お客様提示文言_AfterUpdate
'        MODULE_ID        :txt_お客様提示文言_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_お客様提示文言_AfterUpdate()

    Call fncAfterUpdateCommon(txt_お客様提示文言, P_MONGON_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_お客様提示文言_BeforeUpdate
'        MODULE_ID        :txt_お客様提示文言_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_お客様提示文言_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txt_お客様提示文言, P_MONGON_LENGTH, P_CONTROL_テキスト, Cancel, _
                               "お客様提示文言は半角42桁分までです。" & vbCrLf & "内容を修正して下さい。")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズFrom_AfterUpdate
'        MODULE_ID        :txt_サイズFrom_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズFrom_AfterUpdate()

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズFrom_BeforeUpdate
'        MODULE_ID        :txt_サイズFrom_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズFrom_BeforeUpdate(Cancel As Integer)

    If fncChkTextInput2(Me.txt_サイズFrom, P_SIZE_MAX_VALUE, P_SIZE_LENGTH) = False Then
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズFrom_KeyPress
'        MODULE_ID        :txt_サイズFrom_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズFrom_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    ' マイナス符号が入力された場合、入力を無効にする。
    If KeyAscii = P_MINUS Then KeyAscii = 0: Exit Sub

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.txt_サイズFrom, 3, 2, KeyAscii)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズTo_AfterUpdate
'        MODULE_ID        :txt_サイズTo_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズTo_AfterUpdate()

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズTo_BeforeUpdate
'        MODULE_ID        :txt_サイズTo_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズTo_BeforeUpdate(Cancel As Integer)

    If fncChkTextInput2(Me.txt_サイズTo, P_SIZE_MAX_VALUE, P_SIZE_LENGTH) = False Then
        MsgBox "桁数オーバーしています｡", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズTo_KeyPress
'        MODULE_ID        :txt_サイズTo_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズTo_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    ' マイナス符号が入力された場合、入力を無効にする。
    If KeyAscii = P_MINUS Then KeyAscii = 0: Exit Sub

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.txt_サイズTo, 3, 2, KeyAscii)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_割引値_AfterUpdate
'        MODULE_ID        :txt_割引値_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_割引値_AfterUpdate()

    Dim intLength As Integer

    If opg_割引形式.VALUE = "1" Or opg_割引形式 = "2" Then
        intLength = P_NEBIKI_VALUE_LENGTH
    ElseIf opg_割引形式.VALUE = "1" Or opg_割引形式 = "3" Then
        intLength = P_NEBIKI_VALUE_LENGTH2
    Else
        intLength = P_NEBIKI_VALUE_LENGTH
    End If
    Me.txt_割引値 = fncAfterUpdateCommon(Me.txt_割引値, intLength, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_割引値_BeforeUpdate
'        MODULE_ID        :txt_割引値_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_割引値_BeforeUpdate(Cancel As Integer)

    Dim intLength As Integer

    If opg_割引形式.VALUE = "1" Or opg_割引形式 = "2" Then
        intLength = P_NEBIKI_VALUE_LENGTH
    ElseIf opg_割引形式.VALUE = "1" Or opg_割引形式 = "3" Then
        intLength = P_NEBIKI_VALUE_LENGTH2
    Else
        intLength = P_NEBIKI_VALUE_LENGTH
    End If

    Call fncBeforeUpdateCommon(Me.txt_割引値, intLength, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_割引値_GotFocus
'        MODULE_ID        :txt_割引値_GotFocus
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_割引値_GotFocus()

    Me.txt_割引値.Format = P_FMT_FIXED  ' 書式：固定(カンマ編集なし)
    Me.txt_割引値.VALUE = Replace(Me.txt_割引値, ",", "")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_割引値_LostFocus
'        MODULE_ID        :txt_割引値_LostFocus
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_割引値_LostFocus()

    Me.txt_割引値.Format = P_FMT_STANDARD  ' 書式：標準(カンマ編集あり)
    
    '----↓↓↓↓----20110829--M.HONDA--Ins-------↓↓↓↓---<s>
'    If opg_割引形式.VALUE = "1" Then
    If opg_割引形式.VALUE = "1" Or opg_割引形式.VALUE = "10" Then
    '----↑↑↑↑----20110829--M.HONDA--Ins-------↑↑↑↑---<e>
        Me.txt_割引値.VALUE = Format(Me.txt_割引値, "#,##0")
    Else
        Me.txt_割引値.VALUE = Format(Me.txt_割引値, "#,###")
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_割引値_KeyPress
'        MODULE_ID        :txt_割引値_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_割引値_KeyPress(KeyAscii As Integer)

    Dim intLength As Integer

    If opg_割引形式.VALUE = "1" Or opg_割引形式 = "2" Then
        intLength = P_NEBIKI_VALUE_LENGTH
    ElseIf opg_割引形式.VALUE = "1" Or opg_割引形式 = "3" Then
        intLength = P_NEBIKI_VALUE_LENGTH2
    Else
        intLength = P_NEBIKI_VALUE_LENGTH
    End If

    Call CmKeyPress_For_Num1(Me.txt_割引値, KeyAscii, intLength)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_Click
'        MODULE_ID        :cmd_保存_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()

    Dim objAdoDbConnection As Object
    Dim isExecuteFlg       As Boolean
    Dim isException        As Boolean

    isException = False

    On Error GoTo Exception
    '2009/06/30 MOD <S> 割引マスタ　KASE_DBへ移行
    'Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set objAdoDbConnection = MSZZ025.ADODB_Connection()
    '2009/06/30 MOD <E>

    ' 入力チェック
    If fncInputCheck() = True Then
        If is同一割引条件Check(objAdoDbConnection) = True Then

            ' 自動文言が設定されていなければ設定しておく・・・
            If Nz(Me.txt_文言自動生成.VALUE, "") = "" Then
                Me.txt_文言自動生成.VALUE = fncCreateAutoText()
            End If

            ' 実際の更新開始
            If (pstrProcMode = "0") Then
                ' *** 新規登録 ***
                If MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
                    ' 新規
                    isExecuteFlg = fncHozon(MODE_INSERT, objAdoDbConnection)
                    Call setDefaultWindow
                End If
    
            ElseIf (pstrProcMode = "1") Then
                ' *** 修正 ***
                If MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
                    isExecuteFlg = fncHozon(MODE_UPDATE, objAdoDbConnection)
                End If
            End If
            
        End If
    End If

    If isExecuteFlg Then
        ' 一覧再描画
        Call Form_FKS700.createNebikiData
        Call Form_FKS700.Requery
            
        ' 一覧に戻る(修正モードのとき)
        Call CloseWindow
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objAdoDbConnection Is Nothing Then
        objAdoDbConnection.Close
        Set objAdoDbConnection = Nothing
    End If

    If isException Then Call MSZZ024_M00("cmd_保存_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_戻る_Click
'        MODULE_ID        :cmd_戻る_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_戻る_Click()

    Dim isException As Boolean

    isException = False

    On Error GoTo Exception

    Call CloseWindow

    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_戻る_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Close
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Dim isException As Boolean

    isException = False

    On Error GoTo Exception

    ' コネクションクローズ
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Close", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Dim i              As Integer
    Dim isException    As Boolean

    isException = False

    On Error GoTo Exception

    pboolFrmActive = False

    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb

    ' パラメータ設定
    'pstrParameter = Split(Nz(Me.OpenArgs), ",")
    pstrParameter = Split(Nz(Me.OpenArgs), vbTab)                               'I:2008.11.13 N.Imai

    ' 初期表示項目設定処理
    Call setDefaultWindow

    pboolFrmActive = True

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :fncInputCheck
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck(Optional aお客提示文言チェック可否 As Boolean = True) As Boolean
    
    Dim strSQL        As String
    Dim rsData        As Object
    Dim dblSizeFrom   As Double                                 'INSERT 2008/11/14 K.KINEBUCHI
    Dim dblSizeTo     As Double                                 'INSERT 2008/11/14 K.KINEBUCHI

    On Error GoTo Exception
    
    ' 初期値設定
    fncInputCheck = False

    ' 必須入力チェック

    ' 割引期間
    If Trim(Nz(Me.cmb_割引期間.VALUE)) = "" Then
        MsgBox "必須入力または必須選択項目です。", vbInformation, P_KAKUNIN
        Me.cmb_割引期間.SetFocus
        Exit Function
    End If

    '----↓↓↓↓----20110829--M.HONDA--Ins-------↓↓↓↓---<s>
    ' 割引期間
    If Me.opg_割引形式.VALUE = 10 Then
        If Trim(Nz(Me.cmb_割引期間.VALUE)) <> "01" Then
            MsgBox "事務手数料割引は初月を選択してください。", vbInformation, P_KAKUNIN
            Me.cmb_割引期間.SetFocus
            Exit Function
        End If
    End If
    '----↑↑↑↑----20110829--M.HONDA--Ins-------↑↑↑↑---<e>
    
    '----↓↓↓↓----20120425--M.HONDA--Ins-------↓↓↓↓---<s>
    ' 割引期間
    If Me.opg_割引形式.VALUE = 10 Then
        If Trim(Nz(Me.cmb_利用期間.VALUE)) = "5" Then
            MsgBox "事務手数料割引は利用期間制限なしは選択できません。", vbInformation, P_KAKUNIN
            Me.cmb_割引期間.SetFocus
            Exit Function
        End If
    End If
    '----↑↑↑↑----20120425--M.HONDA--Ins-------↑↑↑↑---<e>
    
    
    

    ' 割引形式
    If Me.opg_割引形式.VALUE = 0 Then
        MsgBox "必須入力または必須選択項目です。", vbInformation, P_KAKUNIN
        Me.opg_割引形式.SetFocus
        Exit Function
    End If
    
    ' 割引値
    If Trim(Nz(Me.txt_割引値.VALUE)) = "" Then
        MsgBox "必須入力または必須選択項目です。", vbInformation, P_KAKUNIN
        Me.txt_割引値.SetFocus
        Exit Function
    End If

    ' 受付種別
    If Trim(Nz(Me.cmb_受付種別.VALUE)) = "" Then
        MsgBox "必須入力または必須選択項目です。", vbInformation, P_KAKUNIN
        Me.cmb_受付種別.SetFocus
        Exit Function
    End If

    ' サイズ指定
    If Me.chk_サイズ指定.VALUE = "-1" Then
        ' チェックあり
        If (Trim(Nz(Me.txt_サイズFrom.VALUE)) = "" And Trim(Nz(Me.txt_サイズTo.VALUE)) = "") Then
            MsgBox "必須設定としています。" & vbCrLf & "入力するか不要ならチェックを外して下さい。", vbInformation, P_KAKUNIN
            Me.txt_サイズFrom.SetFocus
            Exit Function

        ElseIf (Trim(Nz(Me.txt_サイズFrom.VALUE)) <> "" And Trim(Nz(Me.txt_サイズTo.VALUE)) = "") Then
            ' Toが未入力の場合
            ' 入力値を未入力値にコピーする
            Me.txt_サイズTo.VALUE = Me.txt_サイズFrom.VALUE

        ElseIf (Trim(Nz(Me.txt_サイズFrom.VALUE)) = "" And Trim(Nz(Me.txt_サイズTo.VALUE)) <> "") Then
            ' Fromが未入力の場合
            ' 入力値を未入力値にコピーする
            Me.txt_サイズFrom.VALUE = Me.txt_サイズTo.VALUE
        End If

        ' From～To入力矛盾チェック
'        If Me.txt_サイズFrom.VALUE > Me.txt_サイズTo.VALUE Then                         'DELETE 2008/11/14 K.KINEBUCHI
        dblSizeFrom = CDbl(Trim(Nz(Me.txt_サイズFrom.VALUE, 0)))                         'INSERT 2008/11/14 K.KINEBUCHI
        dblSizeTo = CDbl(Trim(Nz(Me.txt_サイズTo.VALUE, 0)))                             'INSERT 2008/11/14 K.KINEBUCHI
        If dblSizeFrom > dblSizeTo Then                                                 'INSERT 2008/11/14 K.KINEBUCHI
            ' 大小チェックエラー
            ' サイズの範囲指定が矛盾しています。 \n From <= To となるように入力して下さい。
            MsgBox "サイズの範囲指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbInformation, P_ERROR
            Me.txt_サイズFrom.SetFocus
            Exit Function
        End If

    End If

    ' 段指定
    If Me.chk_段指定.VALUE = "-1" Then
        ' チェックあり
        If Trim(Nz(Me.cmb_段.VALUE)) = "" Then
            MsgBox "必須設定としています。" & vbCrLf & "入力するか不要ならチェックを外して下さい。", vbInformation, P_KAKUNIN
            Me.cmb_段.SetFocus
            Exit Function
        End If
    End If

    ' 用途指定
    If Me.chk_用途指定.VALUE = "-1" Then
        ' チェックあり
        If Trim(Nz(Me.cmb_用途.VALUE)) = "" Then
            MsgBox "必須設定としています。" & vbCrLf & "入力するか不要ならチェックを外して下さい。", vbInformation, P_KAKUNIN
            Me.cmb_用途.SetFocus
            Exit Function
        End If
    End If

    ' 利用期間
    If Trim(Nz(Me.cmb_利用期間.VALUE)) = "" Then
        MsgBox "必須入力または必須選択項目です。", vbInformation, P_KAKUNIN
        Me.cmb_利用期間.SetFocus
        Exit Function
    End If

    ' お客様提示文言
    If aお客提示文言チェック可否 = True Then
        If Trim(Nz(Me.txt_お客様提示文言.VALUE)) = "" Then
            MsgBox "必須入力または必須選択項目です。", vbInformation, P_KAKUNIN
            Me.txt_お客様提示文言.SetFocus
            Exit Function
        End If
        
        Dim dummy As Integer
        If False = fncBeforeUpdateCommon(txt_お客様提示文言, P_MONGON_LENGTH, P_CONTROL_テキスト, dummy, _
                "お客様提示文言は半角42桁分までです。" & vbCrLf & "内容を修正して下さい。") Then
            Me.txt_お客様提示文言.SetFocus
            Exit Function
        End If
    End If

    fncInputCheck = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncInputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :同一条件割引存在チェック
'        MODULE_ID        :is同一割引条件Check
'        NOTE             :objAdoDbConnection コネクション
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function is同一割引条件Check(objAdoDbConnection As Object) As Boolean
    
    Dim strSQL        As String
    Dim txtエラー文言 As String
    Dim rsData        As Object

    On Error GoTo Exception
    
    ' 初期値設定
    is同一割引条件Check = False
    txtエラー文言 = ""

    ' 同一内容データ存在チェック
    strSQL = ""
    strSQL = "SELECT DCNTM_NO FROM DCNT_MAST "
    strSQL = strSQL & " WHERE DCNTM_BUMOC      = '" & pstrBumonCode & "' "                    ' 部門コード
    strSQL = strSQL & "   AND DCNTM_KIND       = '1' "                                        ' 割引種別("1"固定)
    strSQL = strSQL & "   AND DCNTM_TYPE       = '" & Me.opg_割引形式.VALUE & "' "            ' 割引形式
    strSQL = strSQL & "   AND DCNTM_VALUE      = " & Replace(Me.txt_割引値.VALUE, ",", "")    ' 割引値
    strSQL = strSQL & "   AND DCNTM_PERIOD     = '" & Me.cmb_割引期間.VALUE & "' "            ' 割引期間
    
    If Me.chk_サイズ指定.VALUE = "-1" Then
        strSQL = strSQL & "   AND DCNTM_SIZE_FROM  = " & Me.txt_サイズFrom.VALUE              ' サイズFrom
        strSQL = strSQL & "   AND DCNTM_SIZE_TO    = " & Me.txt_サイズTo.VALUE                ' サイズTo
    Else
        strSQL = strSQL & "   AND DCNTM_SIZE_FROM  IS NULL "                                  ' サイズFrom
        strSQL = strSQL & "   AND DCNTM_SIZE_TO    IS NULL "                                  ' サイズTo
    End If
    
    If Me.chk_段指定.VALUE = "-1" Then
        strSQL = strSQL & "   AND DCNTM_FLOOR      = '" & Me.cmb_段.VALUE & "' "              ' 適用対象階層
    Else
        strSQL = strSQL & "   AND DCNTM_FLOOR      IS NULL "                                  ' 適用対象階層
    End If
    
    If Me.chk_用途指定.VALUE = "-1" Then
        strSQL = strSQL & "   AND DCNTM_USAGE      = '" & Me.cmb_用途.VALUE & "' "            ' 適用対象用途
    Else
        strSQL = strSQL & "   AND DCNTM_USAGE      IS NULL "                                  ' 適用対象用途
    End If
    
    strSQL = strSQL & "   AND DCNTM_GENKBN     = '" & Me.cmb_受付種別.VALUE & "' "            ' 適用発生区分
    strSQL = strSQL & "   AND DCNTM_USE_PERIOD = '" & Me.cmb_利用期間.VALUE & "' "            ' 適用利用月数
    
    ' 修正モードならば自分自身以外でチェック
    If pstrProcMode = "1" Then
        strSQL = strSQL & " AND DCNTM_NO <> '" & pstrNEBIKI_NO & "' "
    End If
    
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    
    
    If Not rsData.EOF Then
        ' 割引番号設定
        txtエラー文言 = "既に割引番号" & rsData.Fields("DCNTM_NO") & "で同一条件の割引情報が存在します。" & vbCrLf & _
                        "条件を変えるか、既存割引" & rsData.Fields("DCNTM_NO") & "を確認して下さい。"
        is同一割引条件Check = False
    Else
        is同一割引条件Check = True
    End If
            
    rsData.Close
    Set rsData = Nothing
            
    ' 最後にメッセージが設定されていれば、それを出力する
    If txtエラー文言 <> "" Then
        MsgBox txtエラー文言, vbInformation, P_ERROR
    End If
            
    Exit Function

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If

    Call Err.Raise(Err.Number, "is同一割引条件Check" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :自動文言生成
'        MODULE_ID        :fncCreateAutoText
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateAutoText() As String

    Dim strAutoText As String

    On Error GoTo Exception

    fncCreateAutoText = ""

    ' 受付種別
    Select Case Me.cmb_受付種別.VALUE
        Case "00"
            ' *** 電話 ***
            strAutoText = "電話から申込した方は"

        Case "01"
            ' *** ネット ***
            strAutoText = "ネットから申込した方は"

        Case "99"
            ' *** 区別無し ***
            ' 文言無し
    End Select

    ' サイズ指定
    If Me.chk_サイズ指定.VALUE = "-1" Then
        ' チェックあり
        If (Me.txt_サイズFrom.VALUE = Me.txt_サイズTo.VALUE) Then
            ' 同値
            strAutoText = strAutoText & Me.txt_サイズFrom.VALUE & "帖なら"
        Else
            ' 違値
            strAutoText = strAutoText & Me.txt_サイズFrom.VALUE & "～" & Me.txt_サイズTo.VALUE & "帖なら"
        End If

    Else
        ' 文言無し
    End If

    ' 段指定
    If Me.chk_段指定.VALUE = "-1" Then
        strAutoText = strAutoText & Me.txt_段.VALUE & "なら"
    Else
        ' 文言無し
    End If

    ' 用途
    If Me.chk_用途指定.VALUE = "-1" Then
        strAutoText = strAutoText & Me.txt_用途.VALUE & "なら"
    Else
        ' 文言無し
    End If

    ' 割引期間
    If Me.cmb_割引期間.VALUE = "00" Then
        strAutoText = strAutoText & "基本料金"
    Else
        strAutoText = strAutoText & Me.txt_割引期間.VALUE
    End If

    ' 割引形式
    Select Case Me.opg_割引形式.VALUE
        Case "1"
            ' *** 円指定 ***
            strAutoText = strAutoText & Replace(Me.txt_割引値.VALUE, ",", "") & "円"

        Case "2"
            ' *** n円割引 ***
            strAutoText = strAutoText & Replace(Me.txt_割引値.VALUE, ",", "") & "円引"

        Case "3"
            ' *** n%引き ***
            strAutoText = strAutoText & Replace(Me.txt_割引値.VALUE, ",", "") & "%引"
    End Select

    ' 利用期間
    If Me.cmb_利用期間.VALUE = "2" Then
        ' 3ヶ月以上
        ' 文言無し
    Else
        strAutoText = strAutoText & "※但し" & Me.txt_利用期間.VALUE & "ご利用の方のみ"
    End If


    fncCreateAutoText = strAutoText

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncCreateAutoText" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :必須チェックコントロール
'        MODULE_ID        :fncRequireControl
'        NOTE             :引数からの値を空データかチェックする
'                          objControl    フォーカスをセットするコントロール
'                          strValue      必須チェックする元
'                          strMsgText    エラーメッセージで使用
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRequireControl( _
                    objControl As Object, _
                    strValue As String, _
                    strMsgText As String, _
                    intLength As Integer) As Boolean

    On Error GoTo Exception

    fncRequireControl = False

    If (strValue = "") Then
        MsgBox strMsgText & "は必須入力です。", vbInformation, "確認"
        objControl.SetFocus
        Exit Function
    End If

    If (CmLenb(strValue) > intLength) Then
        MsgBox strMsgText & "は桁数オーバーです。", vbInformation, "確認"
        objControl.SetFocus
        Exit Function
    End If

    fncRequireControl = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncRequireControl" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      : 割引番号自動発番処理
'        MODULE_ID        : fncGetAutoNebikiNo
'        PARAMETER        : 引数なし
'        戻り値　　　　　 : 発番した割引番号
'        CREATE_DATE      : 2008/07/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetAutoNebikiNo() As String

    Dim objAdoDbConnection   As Object
    Dim strSQL               As String
    Dim rsData               As Object
    Dim intAddValue          As Integer

    On Error GoTo Exception

    ' 初期値設定
    intAddValue = 1

    ' 割引マスタの割引番号の最大値＋ｎを初期割引番号とする
    strSQL = ""
'    strSQL = strSQL & "SELECT COUNT(*) AS Cnt FROM DCNT_MAST "                         'DELETE 2008/09/05 杵渕孝一
    '2009/06/30 MOD <S> 割引マスタ KASE_DB移行
    'strSQL = strSQL & "SELECT MAX(DCNTM_NO) AS Cnt FROM DCNT_MAST "
    'Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    strSQL = strSQL & "SELECT MAX(DCNTM_NO) AS Cnt FROM DCNT_MAST WHERE DCNTM_BUMOC = '" & pstrBumonCode & "' "
    Set objAdoDbConnection = MSZZ025.ADODB_Connection()
    '2009/06/30 MOD <E>
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

    If Nz(rsData.Fields("Cnt"), "") = "" Then
        fncGetAutoNebikiNo = Format(1, "000001")
    Else
'        fncGetAutoNebikiNo = Format(rsData.Fields("Cnt") + intAddValue, "000000")      'DELETE 2008/09/05 杵渕孝一
        fncGetAutoNebikiNo = Format(CLng(rsData.Fields("Cnt")) + intAddValue, "000000") 'INSERT 2008/09/05 杵渕孝一
    End If

    rsData.Close
    objAdoDbConnection.Close
    Set rsData = Nothing
    Set objAdoDbConnection = Nothing

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If

    If Err <> 0 Then
        If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
        Call Err.Raise(Err.Number, "fncGetAutoNebikiNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :一覧に戻る処理
'        MODULE_ID        :CloseWindow
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CloseWindow()

    On Error GoTo Exception

    doCmd.Close acForm, Me.FormName

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "CloseWindow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォーム初期表示処理
'        MODULE_ID        :setDefaultWindow
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setDefaultWindow()

    Dim bMoveYard          As Boolean
    Dim intLoopNo          As Integer
    Dim intKEIYAKU_CNT     As Integer
    Dim objAdoDbConnection As Object

    On Error GoTo Exception

    pstrProcMode = pstrParameter(13)   ' １３番目は処理モード

    ' (0)割引番号 - (1)割引期間 - (2)割引形式 - (3)割引値 - (4)受付種別 - (5)サイズFrom - (6)サイズTo(5) - (7)段 - (8)用途 -
    ' (9)利用期間 - (10)文言自動生成 - (11)お客様提示文言 - (12)タイトル - (13)処理モード

    ' 割引番号設定
    pstrNEBIKI_NO = pstrParameter(0)
    If Nz(pstrNEBIKI_NO, "") <> "" Then
        Me.lbl_割引番号.Caption = "割引番号:" & pstrNEBIKI_NO
    Else
        Me.lbl_割引番号.Caption = ""
    End If

    ' 初期画面設定
    ' (1)割引期間
    Me.cmb_割引期間.VALUE = pstrParameter(1)
    Call cmb_割引期間_AfterUpdate

    ' (2)割引形式
    Me.opg_割引形式.VALUE = pstrParameter(2)
    If Me.opg_割引形式.VALUE = "" Then
        Me.opg_割引形式.VALUE = "0"
    End If
    Call opg_割引形式_AfterUpdate

    ' (3)割引値
    Me.txt_割引値.VALUE = pstrParameter(3)
    Call txt_割引値_LostFocus

    ' (4)受付種別
    Me.cmb_受付種別.VALUE = pstrParameter(4)
    Call cmb_受付種別_AfterUpdate

    ' (5)サイズFrom
    Me.txt_サイズFrom.VALUE = pstrParameter(5)

    ' (6)サイズTo
    Me.txt_サイズTo.VALUE = pstrParameter(6)

    If Trim(Nz(Me.txt_サイズFrom.VALUE, "")) = "" And Trim(Nz(Me.txt_サイズFrom.VALUE, "")) = "" Then
        Me.chk_サイズ指定.VALUE = "0"
    Else
        Me.chk_サイズ指定.VALUE = "-1"
    End If
    Call chk_サイズ指定_AfterUpdate

    ' (7)段
    Me.cmb_段 = pstrParameter(7)
    Call cmb_段_AfterUpdate

    If Trim(Nz(Me.txt_段.VALUE)) = "" Then
        Me.chk_段指定.VALUE = "0"
    Else
        Me.chk_段指定.VALUE = "-1"
    End If
    Call chk_段指定_AfterUpdate

    ' (8)用途
    Me.cmb_用途 = pstrParameter(8)
    Call cmb_用途_AfterUpdate

    If Trim(Nz(Me.txt_用途.VALUE)) = "" Then
        Me.chk_用途指定.VALUE = "0"
    Else
        Me.chk_用途指定.VALUE = "-1"
    End If
    Call chk_用途指定_AfterUpdate

    ' (9)利用期間
    If pstrProcMode = "0" Then
        Me.cmb_利用期間.VALUE = "2"  ' ３ヶ月以上
    Else
        Me.cmb_利用期間.VALUE = pstrParameter(9)
    End If
    Call cmb_利用期間_AfterUpdate

    ' (10)自動文言
    Me.txt_文言自動生成.VALUE = pstrParameter(10)

    ' (11)手動文言
    Me.txt_お客様提示文言.VALUE = pstrParameter(11)

    Select Case pstrProcMode
        Case "0"
            ' *** 新規 ***
            lbl_契約数.Caption = "0件"

        Case "1"
            ' *** 修正 ***
            Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
            intKEIYAKU_CNT = Form_FKS700.fncGetKeiyaku_SU(pstrNEBIKI_NO, objAdoDbConnection)
            lbl_契約数.Caption = intKEIYAKU_CNT & "件"

            objAdoDbConnection.Close
            Set objAdoDbConnection = Nothing

            If intKEIYAKU_CNT > 0 Then
                ' 該当する割引を使用した契約が存在する場合
                ' ⇒ コントロールの制御
                Me.cmb_割引期間.Enabled = False
'                Me.cmb_割引期間.Locked = True
        
                Me.opg_割引形式.Enabled = False
'                Me.opg_割引形式.Locked = True
        
                Me.txt_割引値.Enabled = False
'                Me.txt_割引値.Locked = True
        
                Me.cmb_受付種別.Enabled = False
'                Me.cmb_受付種別.Locked = True
        
                Me.chk_サイズ指定.Enabled = False
'                Me.chk_サイズ指定.Locked = True
        
                Me.txt_サイズFrom.Enabled = False
'                Me.txt_サイズFrom.Locked = True
        
                Me.txt_サイズTo.Enabled = False
'                Me.txt_サイズTo.Locked = True
        
                Me.chk_段指定.Enabled = False
'                Me.chk_段指定.Locked = True
        
                Me.cmb_段.Enabled = False
'                Me.cmb_段.Locked = True
        
                Me.chk_用途指定.Enabled = False
'                Me.chk_用途指定.Locked = True
        
                Me.cmb_用途.Enabled = False
'                Me.cmb_用途.Locked = True
        
                Me.cmb_利用期間.Enabled = False
'                Me.cmb_利用期間.Locked = True

                Me.txt_お客様提示文言.SetFocus
            End If

    End Select

    ' (12)タイトル
    Me.lbl_タイトル.Caption = pstrParameter(12)

    Exit Sub

Exception:
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "setDefaultWindow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :更新処理
'        MODULE_ID        :fncHozon
'        NOTE             :
'        CREATE_DATE      :2008/07/28
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncHozon(iMode As Integer, objAdoDbConnection As Object) As Boolean

    On Error GoTo Exception

    fncHozon = False

    ' トランザクション開始
    objAdoDbConnection.BeginTrans

    ' DB更新処理
    Select Case iMode
        Case MODE_INSERT
            ' *** 登録 ***
            fncHozon = INS_RTN(objAdoDbConnection)

        Case MODE_UPDATE
            ' *** 修正 ***
            fncHozon = UPD_RTN(objAdoDbConnection)
    End Select

Exception:
    If fncHozon = True Then
        ' コミット
        objAdoDbConnection.CommitTrans
    Else
        ' ロールバック
        objAdoDbConnection.RollbackTrans
    End If

    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncHozon" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      : 新規追加処理
'        MODULE_ID        : INS_RTN
'        PARAMETER        : objAdoDbConnection コネクション
'        戻り値           : True=正常 False=異常
'        CREATE_DATE      : 2008/07/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function INS_RTN(objAdoDbConnection As Object) As Boolean
  
    Dim strSQL             As String
    Dim rsData             As Object
    Dim strDate            As String
    Dim strTime            As String
    Dim strNEBIKI_NO       As String

    On Error GoTo Exception

    INS_RTN = False

    '初期値設定
    strDate = Format(DATE, "yyyymmdd")  ' システム日付
    strTime = Format(Now, "hhmmss")     ' システム時刻

    ' 割引マスタ登録
    Set rsData = MSZZ025.ADODB_Recordset("DCNT_MAST", objAdoDbConnection, adoAppendOnly)

    rsData.AddNew

    strNEBIKI_NO = fncGetAutoNebikiNo
    rsData.Fields("DCNTM_NO") = strNEBIKI_NO                                                                 ' 割引番号
    rsData.Fields("DCNTM_BUMOC") = pstrBumonCode                                                             ' 部門コード
    rsData.Fields("DCNTM_KIND") = "1"                                                                        ' 割引種別("1"固定)
    rsData.Fields("DCNTM_TYPE") = Me.opg_割引形式.VALUE                                                      ' 割引形式
    rsData.Fields("DCNTM_VALUE") = Me.txt_割引値.VALUE                                                       ' 割引値
    rsData.Fields("DCNTM_PERIOD") = Me.cmb_割引期間.VALUE                                                    ' 割引期間

    If Me.chk_サイズ指定.VALUE = "-1" Then
        ' チェックあり
        rsData.Fields("DCNTM_SIZE_FROM") = Me.txt_サイズFrom.VALUE                                           ' 適用対象サイズFrom
        rsData.Fields("DCNTM_SIZE_TO") = Me.txt_サイズTo.VALUE                                               ' 適用対象サイズTo
    Else
        ' チェックなし
        rsData.Fields("DCNTM_SIZE_FROM") = Null                                                              ' 適用対象サイズFrom
        rsData.Fields("DCNTM_SIZE_TO") = Null                                                                ' 適用対象サイズTo
    End If

    If Me.chk_段指定.VALUE = "-1" Then
        ' チェックあり
        rsData.Fields("DCNTM_FLOOR") = Me.cmb_段.VALUE                                                       ' 適用対象階層
    Else
        ' チェックなし
        rsData.Fields("DCNTM_FLOOR") = Null                                                                  ' 適用対象階層
    End If

    If Me.chk_用途指定.VALUE = "-1" Then
        ' チェックあり
        rsData.Fields("DCNTM_USAGE") = Me.cmb_用途.VALUE                                                     ' 適用対象用途
    Else
        ' チェックなし
        rsData.Fields("DCNTM_USAGE") = Null                                                                  ' 適用対象用途
    End If

    rsData.Fields("DCNTM_GENKBN") = IIf(Me.cmb_受付種別.VALUE = "", Null, Me.cmb_受付種別.VALUE)             ' 適用発生区分
    rsData.Fields("DCNTM_USE_PERIOD") = Me.cmb_利用期間.VALUE                                                ' 適用利用月数
    rsData.Fields("DCNTM_ATEXT") = IIf(Me.txt_文言自動生成.VALUE = "", "", Me.txt_文言自動生成.VALUE)        ' 自動文言
    rsData.Fields("DCNTM_MTEXT") = IIf(Me.txt_お客様提示文言.VALUE = "", "", Me.txt_お客様提示文言.VALUE)    ' 手動文言
    rsData.Fields("DCNTM_INSED") = strDate                                                                   ' 作成日付
    rsData.Fields("DCNTM_INSEJ") = strTime                                                                   ' 作成時刻
    rsData.Fields("DCNTM_INSPB") = PROG_ID                                                                   ' 作成プログラムID
    rsData.Fields("DCNTM_INSUB") = Form_FKS700.gstrUserName                                                  ' 作成ユーザーID

    rsData.UPDATE

    rsData.Close
    Set rsData = Nothing

    ' メッセージ出力
    MsgBox "割引番号" & strNEBIKI_NO & "で新規登録しました｡" & vbCrLf & "一覧に戻ります。", vbInformation, P_KAKUNIN

    INS_RTN = True

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If

    If Err <> 0 Then
        Call Err.Raise(Err.Number, "INS_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      : 修正処理
'        MODULE_ID        : UPD_RTN
'        PARAMETER        : objAdoDbConnection コネクション
'        戻り値　　　　　 : True=正常 False=異常
'        CREATE_DATE      : 2008/07/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function UPD_RTN(objAdoDbConnection As Object) As Boolean

    Dim strSQL               As String
    Dim rsData               As Object
    Dim strDate              As String
    Dim strTime              As String

    On Error GoTo Exception

    UPD_RTN = False

    '初期値設定
    strDate = Format(DATE, "yyyymmdd")
    strTime = Format(Now, "hhmmss")

    ' 割引マスタ更新
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM DCNT_MAST "
    '2009/06/30 MOD <S> hirano 割引マスタ KASE_DB 移行
    'strSQL = strSQL & "WHERE DCNTM_NO = '" & pstrNEBIKI_NO & "' "
    strSQL = strSQL & "WHERE DCNTM_NO = '" & pstrNEBIKI_NO & "' AND DCNTM_BUMOC = '" & pstrBumonCode & "' "
    '2009/06/30 MOD <E>
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

    If Not rsData.EOF Then
        rsData.Fields("DCNTM_TYPE") = Me.opg_割引形式.VALUE                                                   ' 割引形式
        rsData.Fields("DCNTM_VALUE") = Me.txt_割引値.VALUE                                                    ' 割引値
        rsData.Fields("DCNTM_PERIOD") = Me.cmb_割引期間.VALUE                                                 ' 割引期間

        If Me.chk_サイズ指定.VALUE = "-1" Then
            ' *** チェックあり ***
            rsData.Fields("DCNTM_SIZE_FROM") = Me.txt_サイズFrom.VALUE                                        ' 適用対象サイズFrom
            rsData.Fields("DCNTM_SIZE_TO") = Me.txt_サイズTo.VALUE                                            ' 適用対象サイズTo
        Else
            ' *** チェックなし ***
            rsData.Fields("DCNTM_SIZE_FROM") = Null                                                           ' 適用対象サイズFrom
            rsData.Fields("DCNTM_SIZE_TO") = Null                                                             ' 適用対象サイズTo
        End If

        If Me.chk_段指定.VALUE = -1 Then
            ' *** チェックあり ***
            rsData.Fields("DCNTM_FLOOR") = Me.cmb_段.VALUE                                                    ' 適用対象階層
        Else
            ' *** チェックなし ***
            rsData.Fields("DCNTM_FLOOR") = Null                                                               ' 適用対象階層
        End If

        If Me.chk_用途指定.VALUE = -1 Then
            ' *** チェックあり ***
            rsData.Fields("DCNTM_USAGE") = Me.cmb_用途.VALUE                                                  ' 適用対象用途
        Else
            ' *** チェックなし ***
            rsData.Fields("DCNTM_USAGE") = Null                                                               ' 適用対象用途
        End If

        rsData.Fields("DCNTM_GENKBN") = Me.cmb_受付種別.VALUE                                                 ' 適用発生区分
        rsData.Fields("DCNTM_USE_PERIOD") = Me.cmb_利用期間.VALUE                                             ' 適用利用月数
        rsData.Fields("DCNTM_ATEXT") = IIf(Me.txt_文言自動生成.VALUE = "", "", Me.txt_文言自動生成.VALUE)     ' 自動文言
        rsData.Fields("DCNTM_MTEXT") = IIf(Me.txt_お客様提示文言.VALUE = "", "", Me.txt_お客様提示文言.VALUE) ' 手動文言
        rsData.Fields("DCNTM_UPDAD") = strDate                                                                ' 更新日付
        rsData.Fields("DCNTM_UPDAJ") = strTime                                                                ' 更新時刻
        rsData.Fields("DCNTM_UPDPB") = PROG_ID                                                                ' 更新プログラムID
        rsData.Fields("DCNTM_UPDUB") = Form_FKS700.gstrUserName                                               ' 更新ユーザーID

        rsData.UPDATE
    End If

    rsData.Close
    Set rsData = Nothing

    ' メッセージ出力
    MsgBox "割引番号" & pstrNEBIKI_NO & "を修正しました。" & vbCrLf & "一覧に戻ります。", vbInformation, P_KAKUNIN

    ' 正常終了
    UPD_RTN = True

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If

    If Err <> 0 Then
        Call Err.Raise(Err.Number, "UPD_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncBeforeUpdateCommon
'        機能             :更新前イベント時の処理
'                         :入力チェックを行う
'        IN               :①対象コントロール
'                         :②入力最大桁数
'                         :③定数値「コントロールの種類」参照
'                         :④更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncBeforeUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer, ByRef Cancel As Integer, Optional strErrMsg As String) As Boolean

    Dim strTemp     As String
    Dim blnRet      As Boolean

    If Left(objTarget.NAME, 3) = "cmb" Then
        'コンボボックス入力チェック
        fncBeforeUpdateCommon = fncChkComboInput(objTarget, intLength, Cancel)
    Else
        '入力桁数チェック
        fncBeforeUpdateCommon = fncChkTextInput(objTarget, intLength, Cancel, strErrMsg)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncAfterUpdateCommon
'        機能             :更新後イベント時の処理
'                         :数値項目は０埋めまたは０消しした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'                         :③定数値「コントロールの種類」参照
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :年月は"yyyy/mm"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncAfterUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer) As String

    fncAfterUpdateCommon = ""

    '体裁を整える
    Select Case intKind
    Case P_CONTROL_数値 '数値項目は０埋め
        objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    Case Else
        fncAfterUpdateCommon = Nz(objTarget)
    End Select

    ' コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        If Nz(objTarget) <> "" And _
           objTarget.ListCount <> 0 _
        Then
            fncAfterUpdateCommon = Nz(objTarget.Column(1))
        End If
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDateCheck
'        機能             :日付入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    fncDateCheck = True
    
    var変更日付 = CmDateChecker(txtTarget)
    If IsNull(var変更日付) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkComboInput
'        機能             :コンボボックス入力チェック
'        IN               :第一引数：対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :0=正常
'                          1=入力規則に反している
'                          2=入力値がリストに存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer, ByRef Cancel As Integer) As Integer

    Dim strCode     As String
    Dim intRet      As Integer
    Dim intCount    As Integer
    Dim strValue    As String
    Dim blnFound    As Boolean
    
    intRet = 0
    
    strCode = Nz(cmbCheck, "")
    
    ' 空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkComboInput_End
    End If
    
    If Not fncNumericCheck(cmbCheck, intMaxLength, Cancel) Then
        intRet = 1
        GoTo ChkComboInput_End
    End If
    
    ' 入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        blnFound = False
        ' 入力値を０サプレス
        strValue = Format(cmbCheck, Mid(P_ZERO_SUPPRESS, 1, intMaxLength))
        ' コンボボックス内にコードが存在してる？
        For intCount = 0 To cmbCheck.ListCount - 1
            ' コード値比較
            If cmbCheck.Column(0, intCount) = strValue Then
                blnFound = True
                Exit For
            End If
        Next
        
        If Not blnFound Then
            MsgBox "該当するコードがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
            intRet = 2
            GoTo ChkComboInput_End
        End If
    End If

ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If

    fncChkComboInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkTextInput
'        機能             :テキストボックス入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTextInput(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer, Optional strErrMsg As String) As Boolean

    If Nz(txtCheck) = "" Then
        fncChkTextInput = True
        Exit Function
    End If

    If CmLenb(txtCheck) > intMaxLength Then
        MsgBox strErrMsg, vbInformation, P_ERROR
'        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        fncChkTextInput = False
    Else
        fncChkTextInput = True
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkSingleByte
'        機能             :半角文字入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkSingleByte(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean

    fncChkSingleByte = False

    If Nz(txtCheck) = "" Then
        fncChkSingleByte = True
        Exit Function
    End If

    ' 桁数チェック
    If Not fncChkTextInput(txtCheck, intMaxLength, Cancel) Then
        Exit Function
    End If

    ' 半角文字チェック
    If CmChkCher(Nz(txtCheck.VALUE), "A") Then
        fncChkSingleByte = True
    Else
        MsgBox "半角英数字で入力してください。", vbExclamation, P_KAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :小数入力項目の入力チェック
'        MODULE_ID        :fncChkTextInput2
'        Parameter        :ctlTxt      = 対象となるテキストボックス
'                          varMaxValue = 最大値
'                          intLength   = 長さ
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTextInput2(ByRef ctlTxt As TextBox, _
                                 ByVal varMaxValue As Variant, ByVal intLength As Integer, _
                                 Optional ByVal intKbn As Integer = 0) As Boolean

    fncChkTextInput2 = False

    If IsNull(ctlTxt) = True Then
        fncChkTextInput2 = True: Exit Function
    End If

    Select Case intKbn
        Case 0
            If (ctlTxt > varMaxValue Or _
                Len(Replace(Replace(ctlTxt, ".", ""), "-", "")) > intLength) Then
                Exit Function
            End If

        Case 1
            ' *** フォームのエラーイベントでのエラートラップのとき
            If (ctlTxt > varMaxValue Or _
                Len(Replace(Replace(ctlTxt.Text, ".", ""), "-", "")) > intLength) Then
                Exit Function
            End If
    End Select

    fncChkTextInput2 = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncNumericCheck
'        機能             :数値入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncNumericCheck(objTarget As Object, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean

    Dim strNumber       As String
    Dim lngCompare      As Long

    fncNumericCheck = False

    strNumber = Nz(objTarget, "")
    ' 空白なら入力チェックしない
    If strNumber = "" Then
        fncNumericCheck = True
        Exit Function
    End If

    lngCompare = Val(Mid(P_ALL_NINE, 1, intMaxLength))
    If Val(Nz(objTarget, 0)) > lngCompare Then
        ' 桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        Exit Function
    End If

    fncNumericCheck = True

End Function
