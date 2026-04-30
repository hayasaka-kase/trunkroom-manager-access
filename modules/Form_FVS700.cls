VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　解約入力
'   プログラムＩＤ　：　FVS700/FVS701
'                  :  ※FVS701はFVS700のコピーです
'   作　成　日　　　：  2005/07/11
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：2005/10/02
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：1.解約日入力不可時に3ヶ月前エラーでSetFocusしようとしてシステムエラー
'                       入力チェック時に、Enabledプロパティをチェックするよう修正
'                   : 2.画面に「予約名」「予約番号」追加
'                       コンテナマスタから取得し、保存時にコンテナマスタを更新する
'   Ver             ：  0.1
'
'   修　正　日　　　：2005/12/08
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：1.完了で保存時に、紹介待ち情報作成
'                   : 2.完了から他のものに変更する際に、紹介待ち情報削除
'   Ver             ：  0.2
'
'   修　正　日　　　：2006/2/3
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：1.コンテナ番号指定時に最新の契約を取得する条件を受注契約番号から初回契約日に変更
'                   : 2.コンボボックスの入力時に、先頭の０を入れなくても可にする。
'   Ver             ：  0.3
'
'   修　正　日　　　：2006/3/4
'   修　正　者      ：イーグルソフト 田島
'   修　正　内　容　：1.サイズの表示を追加
'                   : 2.キャンペーン適用満了日チェック対応
'   Ver             ：  0.4
'
'   修　正　日　　　：2006/5/12
'   修　正　者      ：イーグルソフト 田島
'   修　正　内　容　：1.鍵メンテ対応
'                   : 2.セキュリティカード対応
'   Ver             ：  0.5
'
'   修　正　日　　　：2006/6/28
'   修　正　者      ：イーグルソフト 田島
'   修　正　内　容　：1.修繕呼出対応
'   Ver             ：  0.6
'
'   修　正　日　　　：2006/09/19
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：1.ヤード非表示日を過ぎたヤードは選択リストから外す
'                   : 2.営業終了日、ﾔｰﾄﾞ解約予定日追加
'                   : 3.解約日の入力はヤード解約予定日までとする
'                   : 4.予約客先名を削除
'                   : 5.電話番号を削除
'                   : 6.コンテナの直近修繕を表示
'                   : 7.顧客クレームを表示
'   Ver             ：  0.7
'
'   修　正　日      ：2008/03/23
'   修　正　者      ：tajima
'   修　正　内　容  ：解約区分を「未使用キャンセル」にすれば自動的に解約日を
'                     契約開始日の前月末日とする
'   Ver             ：  0.8
'
'   修　正　日      ：2011/09/22
'   修　正　者      ：M.RYU
'   修　正　内　容  ：受付区分=73の場合、"紹介者キャンペー"を表示
'   Ver             ：0.9
'
'   修　正　日      ：2012/01/12
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：解約予定・解約延長の際に解約日の翌月５日を清算依頼日にセットする。
'   Ver             ：1.0
'
'   修　正　日      ：2012/07/04
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：ログ出力対応
'   Ver             ：1.1
'
'   修　正　日      ：2013/06/03
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：終了メッセージを表示させないように修正
'   Ver             ：1.2
'
'   修　正　日      ：2014/04/11
'   修　正　者      ：MIYAMOTO
'   修　正　内　容  ：解約予定に変更した際、新解除番号を発番・登録する
'   Ver             ：1.3
'
'   修　正　日      ：2014/06/17
'   修　正　者      ：MIYAMOTO
'   修　正　内　容  ：地方ヤード以外のコンテナ鍵区分が自動鍵変換対象だった場合、解約予定にした時に自動で自動鍵に変更する
'   Ver             ：1.4
'
'   修　正　日      ：2014/07/04
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：解約延長へ変更した際に、鍵区分を元に戻す処理をコメントアウト
'   Ver             ：1.5
'
'   修　正　日      ：2014/08/02
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：地方ヤードの際には、自動鍵解除番号発番トランへの登録は行わないように修正。
'   Ver             ：1.6
'
'   修　正　日      ：2016/05/20
'   修　正　者      ：MIYAMOTO
'   修　正　内　容  ：地方ヤードの鍵区分を考慮した値取得、条件に変更
'   Ver             ：1.7
'
'   修　正　日      ：2017/05/06
'   修　正　者      ：K.SATO
'   修　正　内　容  ：契約＝終了（P_STATUS_COMPLETE）時にコンテナマスタの用途のコンテナ（バイク）「３３」をコンテナ「０」にする
'   Ver             ：1.8
'
'   修　正　日      ：2018/08/29
'   修　正　者      ：N.IMAI
'   修　正　内　容  ：契約＝終了（P_STATUS_COMPLETE）だけを条件としてコンテナマスタの用途のコンテナ（バイク）「３３」をコンテナ「０」にする
'   Ver             ：1.9
'
'   修　正　日      ：2021/06/02
'   修　正　者      ：EGL
'   修　正　内　容  ：QR番号対応
'   Ver             ：2.0
'
'   修　正　日      ：2022/04/28
'   修　正　者      ：N.IMAI
'   修　正　内　容  ：ヤード保守利用可能者の場合、解除番号のマスクをしない
'   Ver             ：2.1
'
'   修　正　日      ：2023/02/09
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：スタータスに自動解約不可を追加
'   Ver             ：2.2
'
'   修　正　日      ：2025/11/20
'   修　正　者      ：K.KINEBUCHI
'   修　正　内　容  ：「メンテメモ」の下に文言追加
'                     「 ※解約メンテナンス表の「メンテメモ」に出力される文字数は、漢字文字で先頭から18文字です」
'   Ver             ：2.3
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'接続ＤＢオブジェクト
Private objDb           As DAO.Database
Private pdbAccess       As Database 'ワークテーブルのため 2006/05/27 tajima

Private Const PROG_ID = "FVS700"

'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'処理モード管理変数
Private intProcMode     As Integer

'編集モードフラグ　入力値が変更された場合ONになる
Private blnChangeMode As Boolean

'フォームのキャプションを保存する
Private strFormCaption As String

'True=前画面から契約番号を受け取っている
Private blnContractMode As Boolean

'状態区分前回入力値
Private strLastStatus As String

'解約日前回入力値
Private strLastCancelDate As String

'既存データ解約日
Private strExistingCancelDate   As String

'既存データ契約状態
Private strExistingStatus   As String

'既存データメンテ依頼日
Private strMaintenanceRequestDate   As String

'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant

'パラメータでもらった呼び出し元画面名
Private strFormFrom     As String

'データを更新したらTrueになる
Private blnUpdate       As Boolean

'表示データよりも新しい契約データが存在する場合Trueになる
Private blnNewContract  As Boolean

'キャンペーン情報（キャッシュ) 2006/03/04 Tajima ADD
Private pstrCampaignInfomationText  As String                 ' キャンペーン適用情報文言
Private pstrCampaignWarningText     As String                 ' キャンペーン解約警告文言

'鍵＆セキュリティカード関連の追加項目 2006/05/13 tajima ADD
Private pstr元鍵区分 As String

Private pstr現解除番号  As String                               'INSERT 2014/04/11 MIYAMOTO

Private pvONLET_ACCEI   As String                               'INSERT 2021/06/02 EGL
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Const P_PROG_ID             As String = "FVS700"
Private Const P_PROG_ID_REFERENCE   As String = "FVS701"
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

'値が変更された際にタイトル表示する
Private Const P_EDIT_MODE           As String = "(修正中)"
'参照モード(FVS701)として起動されたされた際にタイトル表示する
Private Const P_REFERENCE_MODE      As String = "(参照モード)"

'背景色
Private Const P_BACKCOLOR_YELLOW    As Long = 14680063      '黄
Private Const P_BACKCOLOR_WHITE     As Long = 16777215      '白
Private Const P_BACKCOLOR_BLUE      As Long = 8388608       '青
Private Const P_BACKCOLOR_GLAY      As Long = 12632256      '灰

'契約状態
Private Const P_STATUS_UNSETTING    As Integer = 0      '契約状態「未設定」
Private Const P_STATUS_CONTRACT     As String = "01"    '契約状態「契約」
Private Const P_STATUS_CANCEL       As String = "02"    '契約状態「解約予定」
Private Const P_STATUS_EXTEND       As String = "03"    '契約状態「解約延長」
Private Const P_STATUS_MAINTENANCE  As String = "04"    '契約状態「メンテ中」
Private Const P_STATUS_FUKA         As String = "05"    '契約状態「自動完了不可」

Private Const P_STATUS_COMPLETE     As String = "09"    '契約状態「完了」

'予約受付状態区分
Private Const P_ACCEPT_STATUS_DECISION  As Integer = 1      '予約受付状態区分「予約確定」
Private Const P_ACCEPT_STATUS_KEEP      As Integer = 2      '予約受付状態区分「予約取置」
Private Const P_ACCEPT_STATUS_CONTRACT  As Integer = 10     '予約受付状態区分「契約受付」

'解約区分
Private Const P_CANCEL_NORMAL       As String = "01"    '解約区分「通常」

'処理モード
Private Const P_ヤード選択前    As Integer = 1      'ヤードが未選択
Private Const P_検索前          As Integer = 2      'ヤードが選択され、コンテナ番号が未選択
Private Const P_表示中          As Integer = 3      '検索結果表示中
Private Const P_検索後          As Integer = 4      '検索データ表示後
Private Const P_参照モード      As Integer = 5      '参照モード

'桁数
Private Const P_DIVISION_LENGTH     As Integer = 2      '区分
Private Const P_PERSON_LENGTH       As Integer = 3      '担当コード
Private Const P_CODE_LENGTH         As Integer = 6      'ヤードコード／コンテナ番号
Private Const P_DATE_LENGTH         As Integer = 10      '日付
Private Const P_YM_LENGTH           As Integer = 7       '年月
Private Const P_MAINTENANCE_MEMO_LENGTH     As Integer = 120       'メンテメモ
Private Const P_NOTE_LENGTH         As Integer = 60       '精算備考

'20051002 SHIBAZAKI 予約名/番号追加
Private Const P_BOOKING_NAME_LENGTH     As Integer = 36       '予約名
Private Const P_BOOKING_NO_LENGTH       As Integer = 15       '予約番号

Private Const P_ZERO_SUPPRESS       As String = "0000000000"    '不足桁０埋め用

'契約トランの検索方法
Private Const P_契約番号指定 As Integer = 0     'パラメータで契約番号を受け取った場合
Private Const P_コンテナ指定 As Integer = 1     'ヤードコード／コンテナ番号を選択した場合

'呼び出し元画面名
Private Const P_契約入力        As String = "FVS600"
Private Const P_契約入力_参照モード        As String = "FVS601"
Private Const P_解約一覧        As String = "FVS720"

'2005/12/8 SHIBAZAKI 追加↓
Private Const P_UKNO_HEAD           As String = "@" '紹介トラン作成時の契約番号の頭につける文字
Private Const P_CNTA_USE_OK         As Integer = 1  'コンテナ利用可否「可」
Private Const P_CNTA_USE_NO         As Integer = 0  'コンテナ利用可否「否」
'2005/12/8 SHIBAZAKI 追加↑

Private Const P_UKKBN_73            As Integer = 73 '受付区分（長期契約者紹介） 'INSERT 2011/09/22 M.RYU
Private P_UKKBN_73_TXT              As String   'INSERT 2011/09/22 M.RYU

'↓ INSERT 2014/06/17 MIYAMOTO
Private Const P_AUTO_KAGI_KBN           As Integer = 1  '自動鍵の鍵区分
Private Const P_TIHO_KBN                As Integer = -1 '地方ヤード
Private Const P_AUTO_KAGI_HENKO_RIYU    As Integer = 3  '鍵変更理由
'↑ INSERT 2014/06/17 MIYAMOTO

'2006/05/25 tajima Add
Private Const P_WORK_TABLE          As String = "FVS700_W01"
Private Const P_SUB_FORM            As String = "FVS700_S01"

'用途区分
Private Const P_USAGE_KONTENA       As Integer = 0     'コンテナ INS 2017/05/06 K.SATO
Private Const P_USAGE_KONTENABIKE   As Integer = 33    'コンテナ（バイク） INS 2017/05/06 K.SATO
'

'==============================================================================*
'
'        MODULE_NAME      :chk_AccountFlag_AfterUpdate
'        機能             :精算済み AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_AccountFlag_AfterUpdate()

    Call AfterUpdateCommon(chk_AccountFlag, 0, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelAcceptPerson_AfterUpdate
'        機能             :解約受付担当 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelAcceptPerson_AfterUpdate()

    txt_CancelAcceptPerson = AfterUpdateCommon(cmb_CancelAcceptPerson, P_PERSON_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelAcceptPerson_BeforeUpdate
'        機能             :解約受付担当 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelAcceptPerson_BeforeUpdate(Cancel As Integer)

    On Error GoTo cmb_CancelAcceptPerson_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_CancelAcceptPerson, P_PERSON_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する担当者マスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo cmb_CancelAcceptPerson_BeforeUpdate_End
  
cmb_CancelAcceptPerson_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_CancelAcceptPerson_BeforeUpdate_End:
    Exit Sub

End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelAcceptPerson_KeyPress
'        機能             :解約受付担当 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelAcceptPerson_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_CancelAcceptPerson, KeyAscii, P_PERSON_LENGTH)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelDivision_AfterUpdate
'        機能             :解約区分 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelDivision_AfterUpdate()

    txt_CancelDivision = AfterUpdateCommon(cmb_CancelDivision, P_DIVISION_LENGTH, 1)

    '▼2008/03/23 add tajima
    '  未使用キャンセル且つ契約開始日設定済み、解約日付未入力ならば
    '  解約日を契約開始日前月末日とする
    Dim strTargetDate  As String
    If Me.cmb_CancelDivision.VALUE = "03" And _
        Nz(Me.txt_StartDate.VALUE, "") <> "" And _
        Nz(Me.txt_CancelDate.VALUE, "") = "" Then
        strTargetDate = Me.txt_StartDate.VALUE
        strTargetDate = DateSerial(Year(strTargetDate), _
                                   Month(strTargetDate), _
                                    1) - 1
        Me.txt_CancelDate = Format(strTargetDate, "YYYY/MM/DD")
        '解約受付日の初期設定
        If Nz(Me.txt_CancelAcceptDate, "") = "" Then
            Me.txt_CancelAcceptDate = Format(Now, "YYYY/MM/DD")
        End If
    End If
    '▲2008/03/23 add tajima

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelDivision_BeforeUpdate
'        機能             :解約区分 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelDivision_BeforeUpdate(Cancel As Integer)

    On Error GoTo cmb_CancelDivision_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_CancelDivision, P_DIVISION_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する解約区分がありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo cmb_CancelDivision_BeforeUpdate_End
  
cmb_CancelDivision_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_CancelDivision_BeforeUpdate_End:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelDivision_KeyPress
'        機能             :解約区分 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelDivision_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_CancelDivision, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelReason_AfterUpdate
'        機能             :解約理由 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelReason_AfterUpdate()

    txt_CancelReason = AfterUpdateCommon(cmb_CancelReason, P_DIVISION_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelReason_BeforeUpdate
'        機能             :解約理由 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelReason_BeforeUpdate(Cancel As Integer)

    On Error GoTo cmb_CancelReason_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_CancelReason, P_DIVISION_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する解約理由がありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo cmb_CancelReason_BeforeUpdate_End
  
cmb_CancelReason_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_CancelReason_BeforeUpdate_End:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelReason_KeyPress
'        機能             :解約理由 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CancelReason_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_CancelReason, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ContractState_AfterUpdate
'        機能             :状態区分 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ContractState_AfterUpdate()

    txt_ContractState = AfterUpdateCommon(cmb_ContractState, P_DIVISION_LENGTH, 1)

    If intProcMode <> P_表示中 Then
        Call SelectStatus
    End If
    
'↓ INSERT 2014/06/17 MIYAMOTO
    If Nz(txt_変換鍵区分) <> "" Then
        '変換鍵区分の自動鍵区分を取得
'↓ UPDATE 2016/05/20 MIYAMOTO
'        txt_変換鍵自動鍵区分 = DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txt_変換鍵区分)
        txt_変換鍵自動鍵区分 = MSZZ070_KaiKey(txt_YardTihoKbn, DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txt_変換鍵区分))
'↑ UPDATE 2016/05/20 MIYAMOTO
    End If
'↑ INSERT 2014/06/17 MIYAMOTO
    
'↓ INSERT 2014/04/11 MIYAMOTO
    Dim strBUMOC        As String
    '解約予定に変更した時、新しい解除番号を発番し表示する
    If cmb_ContractState = P_STATUS_CANCEL Then
        If Nz(txt_自動鍵区分) = "1" Then
            strBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
            txt_新解除番号 = MSZZ070_M00(strBUMOC, cmb_Yard, cmb_CntaNo, pstr元鍵区分)
'↓ 2021/06/02 insert EGL
            If pstr元鍵区分 = 16 Then
                txt_新解除番号 = Trim(txt_新解除番号 & Space(1) & MSZZ070_M01(strBUMOC, pstr元鍵区分))
            End If
'↑ 2021/06/02 insert EGL

'↓ INSERT 2014/06/17 MIYAMOTO
        '元データが自動鍵以外の場合
        Else
'↓ UPDATE 2016/05/20 MIYAMOTO
'            '地方ヤード以外の場合、且つ変換鍵でなければ
'            If Nz(txt_YardTihoKbn) <> P_TIHO_KBN And Nz(txt_変換鍵区分) <> "-1" Then
            '変換鍵の場合
            If Nz(txt_変換鍵区分) <> "0" Then
'↑ UPDATE 2016/05/20 MIYAMOTO
                '画面の新(変更後)鍵区分を自動鍵の鍵区分に変更
                cmb_NewKeyDivision = txt_変換鍵区分
                Call cmb_NewKeyDivision_AfterUpdate

                '更に変換先の鍵区分が自動鍵
                If Nz(txt_変換鍵自動鍵区分) = P_AUTO_KAGI_KBN Then
                    '自動鍵の解除番号発番
                    strBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
                    txt_新解除番号 = MSZZ070_M00(strBUMOC, cmb_Yard, cmb_CntaNo, txt_変換鍵区分)
'↓ 2021/06/02 insert EGL
                    If txt_変換鍵区分 = 16 Then
                        txt_新解除番号 = txt_新解除番号 & Space(1) & MSZZ070_M01(strBUMOC, txt_変換鍵区分)
                    End If
'↑ 2021/06/02 insert EGL
                End If
                
                '鍵変更理由の変更(変更理由コードに"03"をセット)
                cmb_鍵変更理由 = P_AUTO_KAGI_HENKO_RIYU
                Call cmb_鍵変更理由_AfterUpdate
            End If
'↑ INSERT 2014/06/17 MIYAMOTO
        End If
'↓ INSERT 2014/06/17 MIYAMOTO
    '契約、解約延長に変更した場合
    ElseIf cmb_ContractState = P_STATUS_CONTRACT Or _
           cmb_ContractState = P_STATUS_EXTEND Then
        '元データが自動鍵以外の場合
        If Nz(txt_自動鍵区分) <> "1" Then
'↓ UPDATE 2016/05/20 MIYAMOTO
'            '地方ヤード以外且つ、変換鍵でなければ
'            If Nz(txt_YardTihoKbn) <> P_TIHO_KBN And Nz(txt_変換鍵区分) <> "-1" Then
            '変換鍵の場合
            If Nz(txt_変換鍵区分) <> "0" Then
'↑ UPDATE 2016/05/20 MIYAMOTO
                '画面の新(変更後)鍵区分を元データの鍵区分に変更
                cmb_NewKeyDivision = pstr元鍵区分
                Call cmb_NewKeyDivision_AfterUpdate

                '解除番号のクリア
                txt_新解除番号 = ""

                '鍵変更理由のクリア
                cmb_鍵変更理由 = ""
                Call cmb_鍵変更理由_AfterUpdate
            End If
        End If
'↑ INSERT 2014/06/17 MIYAMOTO
    End If
'↑ INSERT 2014/04/11 MIYAMOTO
    
    '入力値を保存
    strLastStatus = cmb_ContractState

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ContractState_BeforeUpdate
'        機能             :状態区分 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ContractState_BeforeUpdate(Cancel As Integer)

    On Error GoTo cmb_ContractState_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_ContractState, P_DIVISION_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する状態区分がありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    '解約日が未入力の場合、「メンテ中」「完了」には出来ない
    If Nz(txt_CancelDate) = "" And _
      (cmb_ContractState = P_STATUS_COMPLETE Or cmb_ContractState = P_STATUS_MAINTENANCE) _
    Then
        MsgBox "解約日が未入力の場合、「メンテ中」「完了」に変更できません。", vbInformation, P_KAKUNIN
        SendKeys "{ESC}"
        Cancel = True
    End If
    
    GoTo cmb_ContractState_BeforeUpdate_End
  
cmb_ContractState_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_ContractState_BeforeUpdate_End:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ComboToName
'        機能             :コンボボックスで選択されたコードに対する名称を返却する
'        IN               :対象コンボボックス
'        OUT              :名称
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ComboToName(cmbTarget As ComboBox) As String

    Dim strName     As String

    strName = ""
    
    If Not IsNull(cmbTarget.Column(1)) Then
        strName = cmbTarget.Column(1)
    End If
    
    ComboToName = strName
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_ContractState_KeyPress
'        機能             :状態区分 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ContractState_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_ContractState, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MaintenancePerson_AfterUpdate
'        機能             ：メンテ担当 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MaintenancePerson_AfterUpdate()

    txt_MaintenancePerson = AfterUpdateCommon(cmb_MaintenancePerson, P_PERSON_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MaintenancePerson_BeforeUpdate
'        機能             ：メンテ担当 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MaintenancePerson_BeforeUpdate(Cancel As Integer)

    On Error GoTo cmb_MaintenancePerson_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_MaintenancePerson, P_PERSON_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する担当者マスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo cmb_MaintenancePerson_BeforeUpdate_End
  
cmb_MaintenancePerson_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_MaintenancePerson_BeforeUpdate_End:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MaintenancePerson_KeyPress
'        機能             ：メンテ担当 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MaintenancePerson_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_MaintenancePerson, KeyAscii, P_PERSON_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_NewKeyDivision_AfterUpdate
'        機能             ：新鍵区分 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_NewKeyDivision_AfterUpdate()

    txt_NewKeyDivision = AfterUpdateCommon(cmb_NewKeyDivision, P_DIVISION_LENGTH, 1)

'↓ INSERT 2016/05/20 MIYAMOTO
    If MSZZ070_KaiKey(txt_YardTihoKbn, cmb_NewKeyDivision.Column(2)) = "2" Then
        txt_新解除番号.TabStop = True
        txt_新解除番号.Enabled = True
    Else
        txt_新解除番号.TabStop = False
        txt_新解除番号.Enabled = False
    End If
'↑ INSERT 2016/05/20 MIYAMOTO

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_NewKeyDivision_BeforeUpdate
'        機能             ：新鍵区分 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_NewKeyDivision_BeforeUpdate(Cancel As Integer)

    On Error GoTo cmb_NewKeyDivision_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_NewKeyDivision, P_DIVISION_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する鍵区分がありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
'↓ INSERT 2014/06/17 MIYAMOTO
    '鍵情報メンテの「変更する」チェックがONの場合
    If chk_鍵変更.VALUE = -1 Then
        Dim strNameValFrom      As String
'↓ UPDATE 2016/05/20 MIYAMOTO
'        strNameValFrom = Nz(cmb_NewKeyDivision.Column(2))
'        If strNameValFrom = "1" Then
'            MsgBox "解除番号を自動発番させる鍵区分には手動で変更できません。", vbInformation, P_KAKUNIN
        strNameValFrom = MSZZ070_KaiKey(txt_YardTihoKbn, Nz(cmb_NewKeyDivision.Column(2)))
        If strNameValFrom <> "0" Then
            MsgBox "解除番号を発番させる鍵区分には手動で変更できません。", vbInformation, P_KAKUNIN
'↑ UPDATE 2016/05/20 MIYAMOTO
            SendKeys "{ESC}"
            Cancel = True
        Else
            txt_新解除番号 = ""
        End If
    End If
'↑ INSERT 2014/06/17 MIYAMOTO
    
    GoTo cmb_NewKeyDivision_BeforeUpdate_End
  
cmb_NewKeyDivision_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_NewKeyDivision_BeforeUpdate_End:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_NewKeyDivision_KeyPress
'        機能             ：新鍵区分 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_NewKeyDivision_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_NewKeyDivision, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_鍵変更担当_AfterUpdate
'        機能             :鍵変更担当 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵変更担当_AfterUpdate()

    txt_鍵変更担当 = AfterUpdateCommon(cmb_鍵変更担当, P_PERSON_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_鍵変更担当_BeforeUpdate
'        機能             :鍵変更担当 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵変更担当_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_鍵変更担当, P_PERSON_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する担当者マスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo Finary
  
Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Finary:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CancelAcceptPerson_KeyPress
'        機能             :解約受付担当 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵変更担当_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_鍵変更担当, KeyAscii, P_PERSON_LENGTH)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_鍵変更理由_AfterUpdate
'        機能             :鍵変更理由 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵変更理由_AfterUpdate()

    txt_鍵変更理由 = AfterUpdateCommon(cmb_鍵変更理由, P_DIVISION_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_鍵変更理由_BeforeUpdate
'        機能             :鍵変更理由 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵変更理由_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_鍵変更理由, P_DIVISION_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する変更理由がありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo Finary
  
Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Finary:
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_鍵変更理由_KeyPress
'        機能             :鍵変更理由 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵変更理由_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_鍵変更理由, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_キャンペーン内容_Click
'        機能             ：cmd_キャンペーン内容 Click ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_キャンペーン内容_Click()
  
  If pstrCampaignInfomationText <> "" Then
      MsgBox pstrCampaignInfomationText, vbOKOnly + vbInformation, "この契約で適用したキャンペーンについて"
  End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_下画面_Clickイベント
'        MODULE_ID        :cmd_下画面_Click
'        CREATE_DATE      :2006/05/29
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_下画面_Click()

    GoToPage 1, , Me.Detail.HEIGHT - Me.フォームヘッダー.HEIGHT '- Me.フォームフッター.Height

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_上画面_Clickイベント
'        MODULE_ID        :cmd_上画面_Click
'        CREATE_DATE      :2006/05/29
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_上画面_Click()
    
    GoToPage 1, , 0

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()

    doCmd.Maximize

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form Openイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Title.BackColor = RGB(255, 0, 204)
    End Select
    
    Call MSHS010_M00(PROG_ID, "1")
    
    doCmd.Maximize
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form Loadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim lngRc As Long
    
    'Me.Caption = Me.NAME & Me.Caption                                          'DELETE 2022/04/28 N.IMAI
    
    blnUpdate = False
    
    Dim strOpenArg  As String
    
    strFormFrom = ""
    txt_ContractNo = ""
    P_UKKBN_73_TXT = DLookup("NAME_NAME", "dbo_NAME_MAST", "NAME_ID='070' AND NAME_CODE=" & P_UKKBN_73)     'INSERT 2011/09/22 M.RYU
    
    strOpenArg = Nz(Me.OpenArgs, "")
    If strOpenArg <> "" Then
        lngRc = InStr(strOpenArg, ",")
        If lngRc = 0 Then
            txt_ContractNo = strOpenArg
        Else
            txt_ContractNo = Left(strOpenArg, lngRc - 1)
            strFormFrom = Mid(strOpenArg, lngRc + 1)
        End If
    End If
    
    '前画面からは受注契約番号が来る
    If txt_ContractNo = "" Then
        blnContractMode = False
    Else
        blnContractMode = True
    End If
    
    blnChangeMode = False
    strFormCaption = Me.Caption

'↓ 2021/06/02 insert EGL
    pvONLET_ACCEI = fncGetKaiKeyAdmin()
    If pvONLET_ACCEI = 0 Then
        'マスク
        txt_解除番号.InputMask = "Password"
        txt_新解除番号.InputMask = "Password"
    Else
        txt_解除番号.InputMask = ""
        txt_新解除番号.InputMask = ""
    End If
'↑ 2021/06/02 insert EGL
    
    'SQL-Server接続
    Dim strBUMOC As String
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBUMOC = "" Then
        'テーブル[CONT_MAST]の設定不正
        Exit Sub
    End If
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBUMOC & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If
    '接続文字列取得
    strConnect = MSZZ007_M00(strBUMOC)
    If strConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If

    ' 2006/05/27 ADD tajima Start
    ' Access側ワーククリア
    If Not ClearFVS700_W01(P_WORK_TABLE) Then
        'doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    ' 2006/05/27 ADD tajima End

    If blnContractMode Then
        'パラメータがある場合、契約情報を検索する
        Call GetCargData(P_契約番号指定)
        cmb_Yard.Enabled = False
        cmb_CntaNo.Enabled = False
        'FVS701として起動された場合、参照モードにする
        If Me.NAME = P_PROG_ID_REFERENCE Then
            Call ModeChange(P_参照モード)
        End If
    Else
        'パラメータが無い場合
        'ヤードコードだけ選択できるようにする
        Call ModeChange(P_ヤード選択前)
    End If
   
End Sub

'==============================================================================*
'
'        MODULE_NAME      :解除番号入力権限取得
'        MODULE_ID        :fncGetKaiKeyAdmin
'        CREATE_DATE      :2021/06/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKaiKeyAdmin() As String

    Dim isException     As Boolean
    Dim strSQL          As String
    Dim ojDyn           As Recordset
    Dim vCode(0 To 1)   As Variant
    Dim userId          As String
    Dim dbKASE_DB       As Database
    
    isException = False
    On Error GoTo Exception
    
    ' 加瀬DB接続
    Set dbKASE_DB = Workspaces(0).OpenDatabase(Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'")), dbDriverNoPrompt, False, MSZZ007_M00())
    
    'ユーザ情報取得
    userId = MSZZ000.LsGetUserName()
    
    strSQL = ""
    strSQL = "SELECT ONLET_ACCEI "
    strSQL = strSQL & "FROM ONLE_TABL "
    'strSQL = strSQL & "WHERE ONLET_PROGB = 'FVS400' "                              'DELETE 2022/04/28 N.IMAI
    strSQL = strSQL & "WHERE (ONLET_PROGB = 'FVS400' OR ONLET_PROGB = 'FKS440') "   'INSERT 2022/04/28 N.IMAI
    strSQL = strSQL & "AND ONLET_USERB = '" & userId & "'; "
    Set ojDyn = dbKASE_DB.OpenRecordset(strSQL, dbOpenDynaset)
    
    If Not ojDyn.EOF Then
        fncGetKaiKeyAdmin = ojDyn!ONLET_ACCEI
    Else
        fncGetKaiKeyAdmin = "0"
    End If
    
    GoTo Finally
Exception:
    isException = True
Finally:
    If Not ojDyn Is Nothing Then ojDyn.Close: Set ojDyn = Nothing
    If Not dbKASE_DB Is Nothing Then dbKASE_DB.Close: Set dbKASE_DB = Nothing
    If isException Then
        Call Err.Raise(Err.Number, "fncGetKaiKeyAdmin" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

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

    'F1=「終了」F2=「保存」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_End.SetFocus
        Case vbKeyF2
            '20051004 SHIBAZAKI Enabledプロパティをチェックする
            If Me.cmd_Enter.Enabled Then
                Me.cmd_Enter.SetFocus
            End If
        
        ' ▼2006/06/24 Add tajima
        Case vbKeyF8, vbKeyF9, vbKeyF12
            If Me.Form.NAME = P_PROG_ID Then
                 Call CmReprMod.OpenFVS90n(KeyCode, _
                                           Me.cmb_Yard, _
                                           Me.cmb_CntaNo, _
                                           Me.txt_UserCode, _
                                           Me.txt_ContractNo _
                                           )
             
             ElseIf CmReprMod.P_呼出_ヤード Or CmReprMod.P_呼出_顧客 Or CmReprMod.P_呼出_部屋 Then
                 MsgBox "参照モードでは修繕関連の呼出はできません。" & vbCr & "利用状況にて呼出をしてください。", vbInformation, P_KAKUNIN
             End If
             KeyCode = 0 '2006/10/12 add tajima
       ' ▲2006/06/24 Add tajima
        
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Error
'        機能             :Form_Errorイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
        Case 2169       '「レコードは保存されません。」という時
            DataErr = 0
            Response = 0
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Close
'        機能             :Form_Closeイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Call MSHS010_M00(PROG_ID, "2")
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Form_Unload
'        機能             :Form_Unloadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    ' 2006/05/27 ADD tajima Start
    'Access側ワークテーブルクリア
    If Not ClearFVS700_W01(P_WORK_TABLE) Then
        'doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    ' 2006/05/27 ADD tajima End
    
    Call ParentReload
    
    Exit Sub
    
Err_Form_Unload:
    MsgBox Error$
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ParentReload
'        機能             :呼び出し元更新処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ParentReload()

    Dim intReloadParam      As Integer
    
    If blnUpdate Then
        intReloadParam = 1
    Else
        intReloadParam = 0
    End If
    
    Dim objForm     As Form
    
    If strFormFrom <> "" And strFormFrom <> P_契約入力_参照モード Then
        Set objForm = Forms(strFormFrom)
        If Not objForm Is Nothing Then
            Call objForm.subReload(intReloadParam)
        End If
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Enter_Click
'        機能             :保存ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Enter_Click()

    On Error GoTo cmd_Enter_Click_On_Err

    Dim lngRc As Long

    '必須チェック
    If Not RequireCheck Then
        Exit Sub
    End If

    '入力妥当性チェック
    If Not ProprietyCheck Then
        Exit Sub
    End If

    '完了から他の状態区分に変更する場合、それが可能かをチェック
    If cmb_ContractState <> P_STATUS_COMPLETE And strExistingStatus = P_STATUS_COMPLETE Then
        If Not CompleteToOtherCheck Then
            Exit Sub
        End If
    End If
    
    If Nz(Me.txt_Ukkbn, "") = P_UKKBN_73_TXT Then       'INSERT 2011/09/22 M.RYU
        '長期契約者紹介の顧客のとき、メッセージを表示
        lngRc = MsgBox(P_UKKBN_73_TXT & "の顧客です。保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)    'INSERT 2011/09/22 M.RYU
    Else
        lngRc = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
    End If
    If lngRc = vbNo Then GoTo cmd_Enter_Click_End

    If DB_Access() Then
        blnUpdate = True

        If blnContractMode Then
            '契約番号指定モード時は保存終了時に画面を閉じる
            doCmd.Close acForm, Me.FormName
        Else
            Call cmd_上画面_Click                   '2006/06/01 iguchi ADD
            cmb_Yard.SetFocus

            Call ClearFVS700_W01(P_WORK_TABLE)      ' Access側ワーククリア

            Call ModeChange(P_ヤード選択前)
        End If
    End If

    GoTo cmd_Enter_Click_End

cmd_Enter_Click_On_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmd_Enter_Click_End:

End Sub

'==============================================================================*
'
'        MODULE_NAME      :RequireCheck
'        機能             :必須入力チェック
'        IN               :
'        OUT              :True=必須入力が全て入力されている False=されていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function RequireCheck() As Boolean

    RequireCheck = False
    
    If Not RequireControl(cmb_ContractState, "契約区分", True) Then Exit Function
    
    ' 2006/05/13 Add Tajima
    ' 鍵変更時の必須チェック
    If Not RequireControl(cmb_NewKeyDivision, "新鍵区分", False) Then Exit Function
    If Not RequireControl(txt_新貸出本数, "新貸出本数", False) Then Exit Function
    If Not RequireControl(txt_鍵変更日, "鍵変更日", False) Then Exit Function
    If Not RequireControl(cmb_鍵変更担当, "鍵変更担当", False) Then Exit Function
    If Not RequireControl(cmb_鍵変更理由, "鍵変更理由", False) Then Exit Function
    ' 2006/05/13 Add Tajima ここまで
    
    '状態区分「契約」または、新契約ありの「完了」（完了※）の場合は必須項目無し
    If cmb_ContractState = P_STATUS_CONTRACT Or _
       (cmb_ContractState = P_STATUS_COMPLETE And blnNewContract) _
    Then
        RequireCheck = True
        Exit Function
    End If
    If cmb_ContractState = P_STATUS_EXTEND Then
        '解約延長時は解約予定日は必須
        If Not RequireControl(txt_CancelExpectedDate, "前回解約日", True) Then Exit Function
    End If
    If cmb_ContractState = P_STATUS_CANCEL Then
        '解約時は解約日は必須
        If Not RequireControl(txt_CancelDate, "解約日", True) Then Exit Function
    End If
    
    If Not RequireControl(txt_CancelAcceptDate, "解約受付日", True) Then Exit Function
    
    If Not RequireControl(cmb_CancelDivision, "解約区分", True) Then Exit Function
    
    If Not RequireControl(cmb_CancelReason, "解約理由", True) Then Exit Function
    
    If Not RequireControl(cmb_CancelAcceptPerson, "解約受付担当", True) Then Exit Function

    'メンテ中はメンテ依頼日は必須
    If cmb_ContractState = P_STATUS_MAINTENANCE Then
        If Not RequireControl(txt_MaintenanceRequestDate, "メンテ依頼日", False) Then Exit Function
    End If
    
    '完了時は、読み込み時にメンテナンス依頼日が格納されている場合、
    'メンテ依頼日、メンテ完了日、メンテ担当は必須
    If cmb_ContractState = P_STATUS_COMPLETE And Not strMaintenanceRequestDate = "" Then
        If Not RequireControl(txt_MaintenanceRequestDate, "メンテ依頼日", False) Then Exit Function
        If Not RequireControl(txt_MaintenanceEndDate, "メンテ完了日", False) Then Exit Function
        If Not RequireControl(cmb_MaintenancePerson, "メンテ担当", False) Then Exit Function
    End If
    
    RequireCheck = True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :RequireControl
'        機能             :必須チェック
'        IN               :①対象コントロール　②エラーメッセージに出す文言
'                         :③上下画面切り替え
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function RequireControl(objTarget As Object, strMsgText As String, a上画面切替 As Boolean) As Boolean

    RequireControl = True
    
    '20051004 SHIBAZAKI Enabledプロパティをチェックする
    If objTarget.Enabled Then
        If Nz(objTarget.VALUE) = "" Then
            MsgBox strMsgText & "は必須入力です。", vbInformation, P_KAKUNIN
            '2006/06/01 iguchi start
            If a上画面切替 = True Then
                Call cmd_上画面_Click
            Else
                Call cmd_下画面_Click
            End If
            '2006/06/01 iguchi end
            objTarget.SetFocus
            RequireControl = False
        End If
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :ProprietyCheck
'        機能             :入力妥当性チェック
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ProprietyCheck() As Boolean
    
    ProprietyCheck = False
    
    '20051004 SHIBAZAKI Enabledプロパティをチェックする
    If txt_CancelDate.Enabled And Nz(txt_CancelDate) <> "" Then
        '過去３ヶ月以内であること
        If StrComp(txt_CancelDate, Format(DateAdd("m", C_KAIYAKU_KIKAN_FROM + 1, Now()), "yyyy/mm") & "/01") < 0 Then
            Call cmd_上画面_Click                   '2006/06/01 iguchi ADD
            Me.txt_CancelDate.SetFocus
            MsgBox "解約日付に「" & Format(DateAdd("m", C_KAIYAKU_KIKAN_FROM + 1, Now()), "yyyy/mm") & "/01」より前は指定できません。", _
                   vbInformation, _
                   P_KAKUNIN
            Exit Function
        End If

        '未来20ヶ月以内であること
        If StrComp(txt_CancelDate, Format(DateAdd("m", C_KAIYAKU_KIKAN_TO, Now()), "yyyy/mm") & "/01") >= 0 Then
            Call cmd_上画面_Click                   '2006/06/01 iguchi ADD
            Me.txt_CancelDate.SetFocus
            MsgBox "解約日付に「" & Format(DateAdd("m", C_KAIYAKU_KIKAN_TO, Now()), "yyyy/mm") & "/01」以降は指定できません。", _
                   vbInformation, _
                   P_KAKUNIN
            Exit Function
        End If
        
        '2006/09/19 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↓
        If txt_YardEndDay <> "" Then
            If DateDiff("d", txt_YardEndDay, txt_CancelDate) > 0 Then
                Call cmd_上画面_Click
                Me.txt_CancelDate.SetFocus
                MsgBox "解約日はヤード解約予定日 [" & txt_YardEndDay & "] 以前で無ければいけません。", vbInformation, P_KAKUNIN
                Exit Function
            End If
        End If
        '2006/09/19 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↑
    End If
    
    Select Case cmb_ContractState
    Case P_STATUS_CANCEL
        '20051004 SHIBAZAKI Enabledプロパティをチェックする
        If txt_CancelDate.Enabled Then
            If DateDiff("d", txt_CancelDate, txt_CancelExpectedDate) <> 0 Then
                MsgBox "解約日と前回解約日が異なります。", vbInformation, P_KAKUNIN
                Call cmd_上画面_Click               '2006/06/01 iguchi ADD
                txt_CancelDate.SetFocus
                Exit Function
            End If
        End If
    Case P_STATUS_EXTEND
        '20051004 SHIBAZAKI Enabledプロパティをチェックする
        If txt_CancelExpectedDate.Enabled Then
            If Nz(txt_CancelDate) <> "" Then
                If DateDiff("d", txt_CancelDate, txt_CancelExpectedDate) >= 0 Then
                    MsgBox "前回解約日は解約日より過去の日付を入力してください。", vbInformation, P_KAKUNIN
                    Call cmd_上画面_Click           '2006/06/01 iguchi ADD
                    txt_CancelExpectedDate.SetFocus
                    Exit Function
                End If
            End If
        End If
'    Case P_STATUS_MAINTENANCE, P_STATUS_COMPLETE
'        If Nz(txt_CancelDate) <> "" And Nz(txt_CancelExpectedDate) <> "" Then
'            If DateDiff("d", txt_CancelDate, txt_CancelExpectedDate) > 0 Then
'                MsgBox "解約日は前回解約日以降の日付を入力してください。", vbInformation, P_KAKUNIN
'                txt_CancelDate.SetFocus
'                Exit Function
'            End If
'        End If
    End Select
    
    ProprietyCheck = True
    
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :CompleteToOtherCheck
'        機能             :状態区分「完了」から他のものに変更可能かチェックする
'        IN               :
'        OUT              :True=変更可能 False=変更不可
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CompleteToOtherCheck() As Boolean
On Error GoTo CompleteToOtherCheck_Err

    '2005/12/9 SHIBAZAKI チェックモジュール共通関数化
    Dim strMsg      As String
    
    CompleteToOtherCheck = False
    
    If Not cmb_ContractState.Enabled Then
        CompleteToOtherCheck = True
        Exit Function
    End If
    
    Set objDb = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    If CheckLending(objDb, pcintKAIMODO_CHECK, Val(cmb_Yard), Val(cmb_CntaNo), strMsg) <> pcintCHECK_OK Then
        strMsg = "「完了」から「" & txt_ContractState & "」に変更できません。" & vbCrLf & strMsg
        Call cmd_上画面_Click                           '2006/06/01 iguchi ADD
        cmb_ContractState.SetFocus
        GoTo CompleteToOtherCheck_Err
    End If
    
    CompleteToOtherCheck = True
    
    GoTo CompleteToOtherCheck_End

CompleteToOtherCheck_Err:
    If strMsg = "" Then
        strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    End If
    Err.Clear

CompleteToOtherCheck_End:
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    
    If strMsg <> "" Then
        MsgBox strMsg, vbCritical, P_ERROR
    End If
    
    doCmd.SetWarnings True

'On Error GoTo CompleteToOtherCheck_Err
'
'    CompleteToOtherCheck = False
'
'    doCmd.SetWarnings False
'
'    Set objDb = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
'
'    '20051004 SHIBAZAKI Enabledプロパティをチェックする
'    If Not cmb_ContractState.Enabled Then
'        CompleteToOtherCheck = True
'        GoTo CompleteToOtherCheck_End
'    End If
'
'    If Not PossibleStatusChange() Then
'        MsgBox "既に別の契約が行われているため、「完了」から「" & txt_ContractState & "」に変更できません。", vbInformation, P_KAKUNIN
'        cmb_ContractState.SetFocus
'        GoTo CompleteToOtherCheck_End
'    End If
'
'    If Not PossibleKeep() Then
'        MsgBox "既に予約取置、契約受付された予約があるため、「完了」から「" & txt_ContractState & "」に変更できません。", vbInformation, P_KAKUNIN
'        cmb_ContractState.SetFocus
'        GoTo CompleteToOtherCheck_End
'    End If
'
'    CompleteToOtherCheck = True
'
'    GoTo CompleteToOtherCheck_End
'

End Function

'2005/12/9 SHIBAZAKI チェックモジュール共通関数化
''==============================================================================*
''
''        MODULE_NAME      :PossibleStatusChange
''        機能             :完了⇒完了以外の状態区分変更が可能かチェックする
''        IN               :
''        OUT              :True=可 False=不可
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function PossibleStatusChange() As Boolean
'
'On Error GoTo PossibleStatusChange_Err
'
'    Dim strSql      As String
'    Dim rsCarg      As Recordset
'
'    PossibleStatusChange = True
'
'    '同ヤード、同コンテナで「完了」以外の状態区分が存在する場合、不可とする
'    strSql = strSql & " SELECT * FROM CARG_FILE "
'    strSql = strSql & "  WHERE CARG_YCODE = " & cmb_Yard
'    strSql = strSql & "    AND CARG_NO = " & cmb_CntaNo
'    strSql = strSql & "    AND CARG_AGRE <> " & P_STATUS_COMPLETE
'
'    Set rsCarg = objDb.OpenRecordset(strSql, dbOpenSnapshot)
'
'    If Not rsCarg.EOF Then
'        PossibleStatusChange = False
'    End If
'
'    GoTo PossibleStatusChange_End
'
'PossibleStatusChange_Err:
'    MsgBox "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr _
'            & Err.Description, vbCritical + vbOKOnly, P_ERROR
'    Err.Clear
'
'PossibleStatusChange_End:
'    If Not rsCarg Is Nothing Then rsCarg.Close: Set rsCarg = Nothing
'
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :cmd_Kesi_Click
'        機能             :取消ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kesi_Click()
    
    Dim lngRc As Long
    
    '確認ﾒｯｾｰｼﾞ
    If blnChangeMode Then
        lngRc = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
        If lngRc = vbNo Then Exit Sub
    End If
    
    '画面を初期表示の状態に戻す
    ' 2006/05/27 ADD tajima Start
    ' Access側ワーククリア
    If Not ClearFVS700_W01(P_WORK_TABLE) Then
        'doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    ' 2006/05/27 ADD tajima End
    
    If blnContractMode Then
        '契約番号指定時は検索直後の状態に戻す
        Call GetCargData(P_契約番号指定)
    Else
        'ヤードコード選択待ちの状態に戻す
        Call ModeChange(P_ヤード選択前)
    End If
        Call cmd_上画面_Click           '2006/06/01 iguchi ADD
 
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_End_Click
'        機能             :終了ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()
    Dim strMsg As String
    Dim varDefBtn As Variant
    
    On Error Resume Next
    '' 2013/06/03 M.HONDA START
'    '2006/03/18 ADD tajima Start
'    If blnChangeMode = True Then
'      strMsg = "修正中ですが保存せずに終了しますか？"
'      varDefBtn = vbQuestion + vbYesNo + vbDefaultButton2
'    Else
'      strMsg = "終了します。よろしいですか？"
'      varDefBtn = vbQuestion + vbYesNo + vbDefaultButton1
'    End If
'    '2006/03/18 ADD tajima End
'
'    '終了メッセージを表示
'    If MsgBox(strMsg, varDefBtn, P_KAKUNIN) = vbYes Then
'        doCmd.Close acForm, Me.FormName
'    End If
    '' 2013/06/03 M.HONDA START
        doCmd.Close acForm, Me.FormName
End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_BeforeUpdate
'        機能             :ヤード番号 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_BeforeUpdate(Cancel As Integer)

On Error GoTo cmb_Yard_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_Yard, P_CODE_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo cmb_Yard_BeforeUpdate_End
  
cmb_Yard_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Yard_BeforeUpdate_End:
    Exit Sub
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_AfterUpdate
'        機能             :ヤード番号 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_AfterUpdate()

    On Error GoTo cmb_Yard_AfterUpdate_Err

    '体裁を整える
    Me.cmb_Yard = Format(Me.cmb_Yard, Mid(P_ZERO_SUPPRESS, 1, P_CODE_LENGTH))
    
    'ヤード名／住所取得
    If Not GetYardData() Then
        MsgBox "該当するヤードマスタがありません。", vbInformation, P_KAKUNIN
    End If
    
    'ヤードマスタが検索できた場合、モードを検索前（コンテナ番号指定待ち）にする
    'ヤードマスタが検索できない場合、（ヤードコードが空白等）モードをヤード選択前（ヤードコード指定待ち）
    If Nz(txt_Yard) = "" Then
        Call ModeChange(P_ヤード選択前)
    Else
        cmb_CntaNo.Requery
        Call ModeChange(P_検索前)
    End If
    
    GoTo cmb_Yard_AfterUpdate_End
    
cmb_Yard_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
cmb_Yard_AfterUpdate_End:
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_KeyPress
'        機能             :ヤード番号 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Yard, KeyAscii, P_CODE_LENGTH)
End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_CntaNo_AfterUpdate
'        機能             :コンテナ番号 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CntaNo_AfterUpdate()

    On Error GoTo cmb_CntaNo_AfterUpdate_Err

    '体裁を整える
    Me.cmb_CntaNo = Format(Me.cmb_CntaNo, Mid(P_ZERO_SUPPRESS, 1, P_CODE_LENGTH))
    
    'コンテナ段／種別取得
    If Not GetCntaData() Then
        MsgBox "該当するコンテナマスタがありません。", vbInformation, P_KAKUNIN
    End If
    
    'コンテナマスタが検索できない場合、モードを検索前（コンテナ番号指定待ち）にする
    If Nz(txt_STEP) = "" Then
        Call ModeChange(P_検索前)
    Else
        'そうでなければ、契約トランをコンテナ情報を条件に検索する
        If Not GetCargData(P_コンテナ指定) Then
            Call ModeChange(P_検索前)
        End If
    End If
    
    GoTo cmb_CntaNo_AfterUpdate_End
    
cmb_CntaNo_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
cmb_CntaNo_AfterUpdate_End:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CntaNo_BeforeUpdate
'        機能             :コンテナ番号 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CntaNo_BeforeUpdate(Cancel As Integer)

On Error GoTo cmb_CntaNo_BeforeUpdate_Err

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_CntaNo, P_CODE_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当するコンテナマスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If
    
    GoTo cmb_CntaNo_BeforeUpdate_End
  
cmb_CntaNo_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_CntaNo_BeforeUpdate_End:
    Exit Sub

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_No_KeyPress
'        機能             :ｺﾝﾃﾅ番号KeyPressｲﾍﾞﾝﾄ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CntaNo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Num1(Me.cmb_CntaNo, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :GetYardData
'        機能             :ヤードマスタ問い合わせ
'        IN               :
'        OUT              :取得結果 True:取得できた False:取得できなかった
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetYardData() As Boolean
    
    Dim strSQL  As String
    Dim rsYard  As Recordset
    
On Error GoTo GetYardData_Err
  
    doCmd.Hourglass True

    GetYardData = False
  
    '空白なら名称／住所も空白にする
    If Me.cmb_Yard = "" Then
        GetYardData = True
        txt_Yard = ""
        txt_Yadd = ""
        
        '2006/09/19 SHIBAZAKI ヤード営業期間追加と解約予定日追加
        txt_YardKikanFrom = ""
        txt_YardKikanTo = ""
        txt_YardEndDay = ""
        txt_YardTihoKbn = ""                'INSERT 2014/06/17 MIYAMOTO
        
        GoTo GetYardData_End
    End If
      
    doCmd.SetWarnings False
      
    strSQL = MakeYardSql()
    'データ取得
    Set objDb = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    Set rsYard = objDb.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If Not (rsYard.EOF) Then
        With Me
            .txt_Yard = rsYard.Fields("YARD_NAME")
            .txt_Yadd = rsYard.Fields("YARD_ADDR_1") & rsYard.Fields("YARD_ADDR_2") & rsYard.Fields("YARD_ADDR_3")
            
            '2006/09/19 SHIBAZAKI ヤード営業期間追加と解約予定日追加
            .txt_YardKikanFrom = Nz(rsYard.Fields("YARD_BEGIN_DAY"), "")
            .txt_YardKikanTo = Nz(rsYard.Fields("YARD_RENTEND_DAY"), "")
            .txt_YardEndDay = Nz(rsYard.Fields("YARD_END_DAY"), "")
            .txt_YardTihoKbn = Nz(rsYard.Fields("YARD_TIHO_KBN"), "")       'INSERT 2014/06/17 MIYAMOTO
       End With
    End If
    rsYard.Close
    Set rsYard = Nothing
    objDb.Close
    Set objDb = Nothing
    
    GetYardData = True
    GoTo GetYardData_End

GetYardData_Err:
    If Not rsYard Is Nothing Then rsYard.Close: Set rsYard = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear

GetYardData_End:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeYardSql
'        機能             :ヤードマスタ問い合わせSQL作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeYardSql() As String
    
    Dim strSQL  As String
    
    'データ取得
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "YARD_NAME,"
    strSQL = strSQL & "YARD_ADDR_1,"
    strSQL = strSQL & "YARD_ADDR_2, "
    strSQL = strSQL & "YARD_ADDR_3 "
    
    '2006/09/19 SHIBAZAKI ヤード営業期間とヤード解約予定日追加
    strSQL = strSQL & ","
    strSQL = strSQL & "YARD_BEGIN_DAY, "
    strSQL = strSQL & "YARD_RENTEND_DAY, "
    strSQL = strSQL & "YARD_END_DAY "
    strSQL = strSQL & ", YARD_TIHO_KBN "        'INSERT 2014/06/17 MIYAMOTO
    
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "YARD_MAST "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "YARD_CODE =" & Val(cmb_Yard.VALUE)
    MakeYardSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :GetCntaData
'        機能             :コンテナマスタ問い合わせ
'        IN               :
'        OUT              :取得結果 True:取得できた False:取得できなかった
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetCntaData() As Boolean
    
    Dim strSQL  As String
    Dim vRet    As Variant
    Dim rsCnta  As Recordset
    
On Error GoTo GetCntaData_Err
  
    doCmd.Hourglass True
    GetCntaData = False
  
    '空白ならば段／種別を空白にする
    If Nz(Me.cmb_Yard) = "" Or Nz(Me.cmb_CntaNo) = "" Then
        GetCntaData = True
        txt_STEP = ""
        txt_StepName = ""
        txt_CntaKind = ""
        txt_CntaKindName = ""
        GoTo GetCntaData_End
    End If
      
    doCmd.SetWarnings False
    
    strSQL = MakeCntaSql()

    'データ取得
    Set objDb = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    Set rsCnta = objDb.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If Not (rsCnta.EOF) Then
        With Me
            .txt_STEP = rsCnta.Fields("CNTA_STEP")
            .txt_StepName = rsCnta.Fields("STEP_NAME")
            .txt_CntaKind = rsCnta.Fields("CNTA_SYU")
            .txt_CntaKindName = rsCnta.Fields("KIND_NAME")
            .txt_SIZE = Format(rsCnta.Fields("CNTA_SIZE"), "0.00")  'ADD 2006/03/04 tajima
       End With
    End If
    rsCnta.Close
    Set rsCnta = Nothing
    objDb.Close
    Set objDb = Nothing
    
    GetCntaData = True
    GoTo GetCntaData_End

GetCntaData_Err:
    If Not rsCnta Is Nothing Then rsCnta.Close: Set rsCnta = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear

GetCntaData_End:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeCntaSql
'        機能             :コンテナマスタマスタ問い合わせSQL作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeCntaSql() As String
    
    Dim strSQL As String

    strSQL = strSQL & " SELECT CNTA_MAST.CNTA_STEP, "
    strSQL = strSQL & "        STEPNAME.NAME_NAME AS STEP_NAME, "
    strSQL = strSQL & "        CNTA_MAST.CNTA_SYU, "
    strSQL = strSQL & "        SYUNAME.NAME_NAME AS KIND_NAME "
    strSQL = strSQL & "       ,CNTA_MAST.CNTA_SIZE "            'ADD 2006/03/04 tajima
    strSQL = strSQL & "   FROM CNTA_MAST "
    strSQL = strSQL & "         INNER JOIN NAME_MAST STEPNAME"
    strSQL = strSQL & "            ON STEPNAME.NAME_CODE = CNTA_MAST.CNTA_STEP "
    strSQL = strSQL & "         INNER JOIN NAME_MAST SYUNAME "
    strSQL = strSQL & "            ON SYUNAME.NAME_CODE = CNTA_MAST.CNTA_SYU "
    strSQL = strSQL & "  WHERE CNTA_MAST.CNTA_CODE = " & Val(cmb_Yard.VALUE)
    strSQL = strSQL & "    AND CNTA_MAST.CNTA_NO = " & Val(cmb_CntaNo.VALUE)
    strSQL = strSQL & "    AND STEPNAME.NAME_ID = '090' "
    strSQL = strSQL & "    AND SYUNAME.NAME_ID = '040' "
    
    MakeCntaSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :ChkComboInput
'        機能             :コンボボックス入力チェック
'        IN               :対象コンボボックス
'        OUT              :0=正常
'                          1=入力規則に反している
'                          2=入力値がリストに存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer) As Integer

    Dim strCode As String
    Dim intRet  As Integer
    Dim intCount    As Integer  '06.02.3 SHIBAZAKI 追加
    Dim strValue    As String   '06.02.3 SHIBAZAKI 追加
    Dim blnFound    As Boolean  '06.02.3 SHIBAZAKI 追加
    
    intRet = 0
    
    strCode = Nz(cmbCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkComboInput_End
    End If
    
    '入力チェック
    Select Case CmNumCheck(strCode, intMaxLength, 0)
        Case 3
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            intRet = 1
            GoTo ChkComboInput_End
        Case 2
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            SendKeys "{ESC}"
            intRet = 1
            GoTo ChkComboInput_End
    End Select
    
    '入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        '06.2.3 SHIBAZAKI 追加↓
        blnFound = False
        '入力値を０サプレス
        strValue = Format(cmbCheck, Mid(P_ZERO_SUPPRESS, 1, intMaxLength))
        'コンボボックス内にコードが存在してる？
        For intCount = 0 To cmbCheck.ListCount - 1
            'コード値比較
            If cmbCheck.Column(0, intCount) = strValue Then
                blnFound = True
                Exit For
            End If
        Next
        '06.2.3 SHIBAZAKI 追加↑
        
        '06.2.3 SHIBAZAKI 変更
        'If cmbCheck.ListIndex = -1 Then
        If Not blnFound Then
            intRet = 2
            GoTo ChkComboInput_End
        End If
    End If

ChkComboInput_End:
    ChkComboInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :ChkTextInput
'        機能             :コンボボックス入力チェック
'        IN               :対象コンボボックス
'        OUT              :
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function chkTextInput(txtCheck As TextBox, intMaxLength As Integer) As Integer

    Dim strCode As String
    Dim intRet  As Integer
    
    intRet = 0
    
    strCode = Nz(txtCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkTextInput_End
    End If
    
    '入力チェック
    intRet = CmBCheck_Char(strCode, intMaxLength)
    Select Case intRet
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo ChkTextInput_End

        Case 9
            MsgBox "備考を入力してください。", vbInformation, P_ERROR
            SendKeys "{ESC}"
            GoTo ChkTextInput_End
    End Select
    
ChkTextInput_End:
    chkTextInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :GetCargData
'        機能             :コンテナ契約ファイル問い合わせ
'        IN               :iMode=0(契約番号指定)/1(コンテナ指定)
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetCargData(ByVal iMode As Integer) As Boolean
    
    Dim strSQL          As String
    Dim rsCarg          As Recordset
    Dim strMntCompDate  As String
    Dim lngUcode        As Long
    
On Error GoTo GetCargData_Err
    
    doCmd.Hourglass True
    GetCargData = False
  
    'コンテナ指定でヤードコード／コンテナ番号が指定されていない
    'または、契約番号指定で契約番号が指定されていない
    If (iMode = P_コンテナ指定 And (Nz(Me.cmb_Yard) = "" Or Nz(Me.cmb_CntaNo) = "")) Or _
       (iMode = P_契約番号指定 And Nz(Me.txt_ContractNo) = "") _
    Then
        GetCargData = True
        GoTo GetCargData_End
    End If
      
    doCmd.SetWarnings False
    
    strSQL = MakeCargSql(iMode)
  
    'データ取得
    Set objDb = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    Set rsCarg = objDb.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsCarg.EOF Then
        MsgBox "該当するコンテナ契約ファイルが存在しません。", vbInformation, P_ERROR
        GetCargData = False
    Else
        '画面表示
        Call ModeChange(P_表示中)
        Call DispCargData(rsCarg, iMode)
        blnNewContract = ExistNewContract(rsCarg)
        '件数取得
        strMntCompDate = Nz(rsCarg.Fields("CARG_MNT_CMPDATE"))
        lngUcode = Nz(rsCarg.Fields("CARG_UCODE"), 0)
        'キャンペーン文言取得
        Call GetCampaignTexts(rsCarg) ' 2006/03/04 ADD tajima
        Dim ACPTNO As String                '2006/05/26 ADD tajima
        ACPTNO = rsCarg.Fields("CARG_ACPTNO")  '2006/05/26 ADD tajima
        rsCarg.Close
        Set rsCarg = Nothing
        Call GetCargCount(strMntCompDate, lngUcode)
        ' サブフォームの設定 start 2006/05/26 Add tajima
        Call MakeFVS700_W01(objDb, CLng(Me.cmb_Yard), ACPTNO)
        SubForm.Requery
        ' end 2006/05/26 Add tajima
        '状態区分による画面制御
        Call SelectStatus
        '処理モード変更
        Call ModeChange(P_検索後)
        blnChangeMode = False
        Me.Caption = strFormCaption
        GetCargData = True
    End If
    
    GoTo GetCargData_End

GetCargData_Err:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear

GetCargData_End:
    If Not rsCarg Is Nothing Then rsCarg.Close: Set rsCarg = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    doCmd.Hourglass False
    doCmd.SetWarnings True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeCargSql
'        機能             :コンテナ契約ファイル問い合わせSQL作成
'        IN               :iMode=0(契約番号指定)/1(コンテナ指定)
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeCargSql(iMode As Integer) As String

    Dim strSQL As String
    
    strSQL = " SELECT "
    If iMode = P_契約番号指定 Then
        '契約番号指定の場合、コンテナの情報も検索する
        strSQL = strSQL & " CARG_YCODE,  "
        strSQL = strSQL & " YARD_NAME, "
        strSQL = strSQL & " YARD_ADDR_1+ "
        strSQL = strSQL & " YARD_ADDR_2+ "
        strSQL = strSQL & " YARD_ADDR_3 AS YARD_ADDR, "
        strSQL = strSQL & " CARG_NO,  "
        strSQL = strSQL & " CNTA_STEP, "
        strSQL = strSQL & " STEPNAME.NAME_NAME AS STEP_NAME, "
        strSQL = strSQL & " CNTA_SYU, "
        strSQL = strSQL & " SYUNAME.NAME_NAME AS KIND_NAME, "
        strSQL = strSQL & " CNTA_SIZE, "      'ADD 2006/03/04 tajima
        '2006/09/19 SHIBAZAKI ヤード営業期間追加と解約予定日追加
        strSQL = strSQL & " YARD_BEGIN_DAY, "
        strSQL = strSQL & " YARD_RENTEND_DAY, "
        strSQL = strSQL & " YARD_END_DAY, "
        strSQL = strSQL & " YARD_TIHO_KBN, "        'INSERT 2014/06/17 MIYAMOTO
    Else
        strSQL = strSQL & " TOP 1 "
    End If
    
    strSQL = strSQL & MakeCommonSelect() & MakeFrom(iMode) & MakeWhere(iMode)

    If iMode = P_コンテナ指定 Then
        'コンテナ指定の場合、初回契約日で降順ソートする
        '2006.2.3 SHIBAZAKI 変更
        'strSQL = strSQL & " ORDER BY CARG_ACPTNO DESC "
        strSQL = strSQL & " ORDER BY CARG_FSDATE DESC "
    End If
    
    MakeCargSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeCommonSelect
'        機能             :コンテナ契約ファイル問い合わせSQL作成
'                          (SELECT句共通部分)
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeCommonSelect() As String

    Dim strSQL As String
    
    strSQL = strSQL & " CARG_ACPTNO,  "
    strSQL = strSQL & " CARG_UCODE, "
    strSQL = strSQL & " USER_NAME, "
    strSQL = strSQL & " USER_KANA, "
    strSQL = strSQL & " CARG_AGRE, "
    strSQL = strSQL & " CARG_KAIUKEDATE, "
    strSQL = strSQL & " CARG_KAITANTO, "
    strSQL = strSQL & " CARG_KAICD, "
    strSQL = strSQL & " CARG_KAIRIYUCD, "
    strSQL = strSQL & " CARG_KAIJB, "
    strSQL = strSQL & " CARG_KYDATE, "
    strSQL = strSQL & " CARG_FSDATE, "
    strSQL = strSQL & " CARG_STDATE, "
    strSQL = strSQL & " CARG_EDDATE, "
    strSQL = strSQL & " CARG_CYDATE, "
    strSQL = strSQL & " CARG_KAGIICD, "
    strSQL = strSQL & " KAGIICD_NAME.NAME_NAME AS KAGIICD_NAME,"
    strSQL = strSQL & " KAGIICD_NAME.NAME_VALUE_FROM AS AUTO_KAGI_KBN,"     'INSERT 2014/04/11 MIYAMOTO
    strSQL = strSQL & " KAGIICD_NAME.NAME_VALUE_TO AS CHANGE_KAGI_KBN,"     'INSERT 2014/06/17 MIYAMOTO
    strSQL = strSQL & " RCPT_TRAN.RCPT_KAINO,"                              'INSERT 2014/04/11 MIYAMOTO
    strSQL = strSQL & " RCPT_TRAN.RCPT_BUIL_KAGI_NO,"                       'INSERT 2021/06/02 EGL
    strSQL = strSQL & " CARG_KEY_RETDATE, "
    strSQL = strSQL & " CARG_SPAREKEY, "
    strSQL = strSQL & " CARG_MNT_REQDATE, "
    strSQL = strSQL & " CARG_MNT_CMPDATE, "
    strSQL = strSQL & " CARG_MNT_TANTO, "
    strSQL = strSQL & " CARG_NIRCD, "
    strSQL = strSQL & " CARG_SYOCD, "
    strSQL = strSQL & " CARG_HOSCD, "
    strSQL = strSQL & " CARG_MNT_MEMO, "
    strSQL = strSQL & " CARG_LIQ_REQDATE, "
    strSQL = strSQL & " CARG_LIQ_NOTE, "
    strSQL = strSQL & " CARG_LIQ_FLG, "
    strSQL = strSQL & " CARG_CONTNA, "

    '20051002 SHIBAZAKI コンテナマスタから予約名と予約番号を取得する
    strSQL = strSQL & " CNTA_YNAME, "
    strSQL = strSQL & " CNTA_YTEL "

    '2006/03/04 tajima 受付対応、予約番号とキャンペーン適用満了日を取得する
    strSQL = strSQL & " ,CARG_UKNO "
    strSQL = strSQL & " ,CARG_CMP_EXDATE "
    strSQL = strSQL & " ,RCPT_TRAN.RCPT_SEV1N "
    strSQL = strSQL & " ,RCPT_TRAN.RCPT_SEV2N "
    strSQL = strSQL & " ,RCPT_TRAN.RCPT_SEV3N "
    strSQL = strSQL & " ,RCPT_TRAN.RCPT_ENDEN "
    
    '2006/05/13 tajima 鍵管理対応
    strSQL = strSQL & " ,CARG_KEY_LENTNUM "               'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CARG_KEY_RETNUM "                'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CARG_SPAREKEY_NOTE "             'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CARG_KEY_CHGDATE "               'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CARG_KEY_CHGTANTO "              'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CARG_KEY_CHGRIYUCD "             'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CARG_KEY_CHGNOTE "               'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CNTA_MAST.CNTA_KEY_NUM "         'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CNTA_MAST.CNTA_KEY_TYPE "        'ADD 2006/05/13 tajima
    strSQL = strSQL & " ,CNTA_MAST.CNTA_KEY_NO "          'ADD 2006/05/13 tajima
    
    strSQL = strSQL & " ,YOUK_TRAN.YOUKT_UKKBN  "  '受付区分   'INSERT 2011/09/22 M.RYU
    
    '↓ INSERT 2014/04/11 MIYAMOTO
    strSQL = strSQL & " ,( " & vbCrLf
    strSQL = strSQL & "    SELECT  AKKNT.AKKNT_KAINO " & vbCrLf
    strSQL = strSQL & "    FROM    ( " & vbCrLf
    strSQL = strSQL & "              SELECT  AKKNT_YARD " & vbCrLf
    strSQL = strSQL & "              ,       AKKNT_CTNO " & vbCrLf
    strSQL = strSQL & "              ,       AKKNT_KTYPE " & vbCrLf
    strSQL = strSQL & "              ,       AKKNT_HATUD " & vbCrLf
'    strSql = strSql & "              ,       ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE ORDER BY AKKNT_HATUD DESC) AS SEQ " & vbCrLf'UPDATE 2021/06/02 EGL
    strSQL = strSQL & "              ,       ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE, AKKNT_STYPE ORDER BY AKKNT_HATUD DESC) AS SEQ " & vbCrLf 'UPDATE 2021/06/02 EGL
    strSQL = strSQL & "              ,       AKKNT_KAINO " & vbCrLf
    strSQL = strSQL & "              FROM    AKKN_TRAN " & vbCrLf
'↓ UPDATE 2014/06/17 MIYAMOTO
'    strSQL = strSQL & "              WHERE   AKKNT_YARD = CARG_FILE.CARG_YCODE " & vbCrLf
'    strSQL = strSQL & "              AND     AKKNT_CTNO = CARG_FILE.CARG_NO " & vbCrLf
'    strSQL = strSQL & "              AND     AKKNT_KTYPE = CARG_FILE.CARG_KAGIICD " & vbCrLf
    strSQL = strSQL & "              WHERE   AKKNT_YARD = CNTA_MAST.CNTA_CODE " & vbCrLf
    strSQL = strSQL & "              AND     AKKNT_CTNO = CNTA_MAST.CNTA_NO " & vbCrLf
    strSQL = strSQL & "              AND     AKKNT_KTYPE = CNTA_MAST.CNTA_KEY_TYPE " & vbCrLf
    strSQL = strSQL & "              AND     AKKNT_STYPE = '4'" 'INSERT 2021/06/02 EGL
'↑ UPDATE 2014/06/17 MIYAMOTO
    strSQL = strSQL & "            ) AKKNT " & vbCrLf
    strSQL = strSQL & "    WHERE   AKKNT.SEQ = 1 " & vbCrLf
    strSQL = strSQL & "  ) AS AKKNT_KAINO " & vbCrLf
    '↑ INSERT 2014/04/11 MIYAMOTO
    
'↓ INSERT 2021/06/02 EGL
    strSQL = strSQL & " ,( " & vbCrLf
    strSQL = strSQL & "    SELECT  AKKNT.AKKNT_KAINO " & vbCrLf
    strSQL = strSQL & "    FROM    ( " & vbCrLf
    strSQL = strSQL & "              SELECT  AKKNT_YARD " & vbCrLf
    strSQL = strSQL & "              ,       AKKNT_CTNO " & vbCrLf
    strSQL = strSQL & "              ,       AKKNT_KTYPE " & vbCrLf
    strSQL = strSQL & "              ,       AKKNT_HATUD " & vbCrLf
    strSQL = strSQL & "              ,       ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE, AKKNT_STYPE ORDER BY AKKNT_HATUD DESC) AS SEQ " & vbCrLf
    strSQL = strSQL & "              ,       AKKNT_KAINO " & vbCrLf
    strSQL = strSQL & "              FROM    AKKN_TRAN " & vbCrLf
    strSQL = strSQL & "              WHERE   AKKNT_YARD = CNTA_MAST.CNTA_CODE " & vbCrLf
    strSQL = strSQL & "              AND     AKKNT_CTNO = CNTA_MAST.CNTA_NO " & vbCrLf
    strSQL = strSQL & "              AND     AKKNT_KTYPE = CNTA_MAST.CNTA_KEY_TYPE " & vbCrLf
    strSQL = strSQL & "              AND     AKKNT_STYPE = '6'"
    strSQL = strSQL & "            ) AKKNT " & vbCrLf
    strSQL = strSQL & "    WHERE   AKKNT.SEQ = 1 " & vbCrLf
    strSQL = strSQL & "  ) AS AKKNT_KAINO_QR " & vbCrLf
'↑ INSERT 2021/06/02 EGL
    
    MakeCommonSelect = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeFrom
'        機能             :コンテナ契約ファイル問い合わせSQL作成
'                          (FROM句)
'        IN               :iMode=0(契約番号指定)/1(コンテナ指定)
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeFrom(iMode As Integer) As String

    Dim strSQL As String
    
    '20051002 SHIBAZAKI コンテナマスタから予約名と予約番号を取得する
    strSQL = strSQL & " FROM  "
    strSQL = strSQL & " CARG_FILE "
    strSQL = strSQL & " INNER JOIN CNTA_MAST "
    strSQL = strSQL & "    ON CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE "
    strSQL = strSQL & "   AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO "
    
    If iMode = P_契約番号指定 Then
        '契約番号指定の場合、コンテナの情報も検索する
        strSQL = strSQL & " INNER JOIN YARD_MAST "
        strSQL = strSQL & "    ON YARD_MAST.YARD_CODE = CARG_FILE.CARG_YCODE "
        strSQL = strSQL & " INNER JOIN NAME_MAST AS STEPNAME "
        strSQL = strSQL & "    ON STEPNAME.NAME_CODE = CNTA_MAST.CNTA_STEP "
        strSQL = strSQL & " INNER JOIN NAME_MAST AS SYUNAME "
        strSQL = strSQL & "    ON SYUNAME.NAME_CODE = CNTA_MAST.CNTA_SYU "
    End If
        
    strSQL = strSQL & " INNER JOIN USER_MAST "
    strSQL = strSQL & "    ON USER_MAST.USER_CODE = CARG_FILE.CARG_UCODE "
    strSQL = strSQL & " LEFT OUTER JOIN NAME_MAST AS KAGIICD_NAME "
    strSQL = strSQL & "    ON KAGIICD_NAME.NAME_CODE = CAST( CARG_FILE.CARG_KAGIICD AS INT )"
    strSQL = strSQL & "   AND KAGIICD_NAME.NAME_ID = '060' "
    
    'Add 20060304 tajima キャンペーン情報の取得
    strSQL = strSQL & " LEFT OUTER JOIN RCPT_TRAN "
    strSQL = strSQL & "    ON CARG_FILE.CARG_UKNO = RCPT_TRAN.RCPT_NO       "

    'INSERT 2011/09/22 M.RYU 受付区分を取得
    strSQL = strSQL & " LEFT OUTER JOIN YOUK_TRAN "
    strSQL = strSQL & "    ON CARG_FILE.CARG_UKNO = YOUK_TRAN.YOUKT_UKNO    "
    
    MakeFrom = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeWhere
'        機能             :コンテナ契約ファイル問い合わせSQL作成
'                          (WHERE句)
'        IN               :iMode=0(契約番号指定)/1(コンテナ指定)
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeWhere(iMode As Integer) As String

    Dim strSQL As String
    
    If iMode = P_契約番号指定 Then
        strSQL = strSQL & " WHERE CARG_ACPTNO = '" & txt_ContractNo & "' "
        strSQL = strSQL & "   AND STEPNAME.NAME_ID = '090' "
        strSQL = strSQL & "   AND SYUNAME.NAME_ID = '040' "
    Else
        strSQL = strSQL & " WHERE CARG_YCODE = " & Val(cmb_Yard.VALUE)
        strSQL = strSQL & "   AND CARG_NO = " & Val(cmb_CntaNo.VALUE)
    End If
    
    MakeWhere = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :DispCargData
'        機能             :契約トラン画面表示
'        IN               :rsCarg(検索結果レコードセット)
'                          iMode=0(契約番号指定)/1(コンテナ指定)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DispCargData(rsCarg As Recordset, iMode As Integer)

    With Me
        If iMode = P_契約番号指定 Then
            '契約番号指定の場合、コンテナの情報も表示する
            .cmb_Yard = Format(rsCarg.Fields("CARG_YCODE"), Mid(P_ZERO_SUPPRESS, 1, P_CODE_LENGTH))
            .txt_Yard = rsCarg.Fields("YARD_NAME")
            .txt_Yadd = rsCarg.Fields("YARD_ADDR")
            .cmb_CntaNo = Format(rsCarg.Fields("CARG_NO"), Mid(P_ZERO_SUPPRESS, 1, P_CODE_LENGTH))
            .txt_STEP = rsCarg.Fields("CNTA_STEP")
            .txt_StepName = rsCarg.Fields("STEP_NAME")
            .txt_CntaKind = rsCarg.Fields("CNTA_SYU")
            .txt_CntaKindName = rsCarg.Fields("KIND_NAME")
            .txt_SIZE = Format(rsCarg.Fields("CNTA_SIZE"), "0.00")  'ADD 2006/03/04 tajima
            '2006/09/19 SHIBAZAKI ヤード営業期間追加と解約予定日追加
            .txt_YardKikanFrom = Nz(rsCarg.Fields("YARD_BEGIN_DAY"), "")
            .txt_YardKikanTo = Nz(rsCarg.Fields("YARD_RENTEND_DAY"), "")
            .txt_YardEndDay = Nz(rsCarg.Fields("YARD_END_DAY"), "")
            .txt_YardTihoKbn = Nz(rsCarg.Fields("YARD_TIHO_KBN"), "")       'INSERT 2014/06/17 MIYAMOTO
        End If
        
        Call DispContract(rsCarg)
        Call DispMaintenance(rsCarg)
        Call DispAccount(rsCarg)
        
        '2006/09/19 SHIBAZAKI　コンテナの直近修繕と顧客クレームを表示
        Call DispReprInfo
       
    End With

End Sub

'2006/09/19 SHIBAZAKI　コンテナの直近修繕と顧客クレームを表示
'==============================================================================*
'
'        MODULE_NAME      :DispMntInfo
'        機能             :修繕情報表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DispReprInfo()
On Error GoTo ErrRtn

    Dim strKagtosSqlserver      As String
    Dim strKagtosConnect        As String
    Dim dbKagtos                As Database
    Dim strBUMOC                As String
    Dim blnError                As Boolean
    
    blnError = False
    
    '加瀬ＤＢに接続
    strKagtosSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    strKagtosConnect = MSZZ007_M00()
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    Set dbKagtos = Workspaces(0).OpenDatabase(strKagtosSqlserver, dbDriverNoPrompt, False, strKagtosConnect)

    txt_CntaRepr = GetRoomLatestReprText(dbKagtos, strBUMOC, cmb_Yard, cmb_CntaNo)
    txt_UserRepr = GetUserLatestKotoText(dbKagtos, strBUMOC, txt_UserCode)
    
    dbKagtos.Close
    Set dbKagtos = Nothing
    
    GoTo ExitRtn

ErrRtn:
    blnError = True
    
ExitRtn:
    If Not dbKagtos Is Nothing Then dbKagtos.Close: Set dbKagtos = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DispContract
'        機能             :契約情報画面表示
'        IN               :rsCarg(検索結果レコードセット)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DispContract(rsCarg As Recordset)

    With Me
        .txt_ContractNo = rsCarg.Fields("CARG_ACPTNO")
        
        .txt_UserCode = rsCarg.Fields("CARG_UCODE")
        .txt_UserName = rsCarg.Fields("USER_NAME")
        .txt_UserKana = rsCarg.Fields("USER_KANA")
         
         '受付区分  'INSERT 2011/09/22 M.RYU
        .txt_Ukkbn = IIf(Nz(rsCarg.Fields("YOUKT_UKKBN"), "0") = P_UKKBN_73, P_UKKBN_73_TXT, "")
        
        '「未設定」の場合、「契約」に変える
        If rsCarg.Fields("CARG_AGRE") = P_STATUS_UNSETTING Then
            .cmb_ContractState = P_STATUS_CONTRACT
        Else
            .cmb_ContractState = rsCarg.Fields("CARG_AGRE")
        End If
        strExistingStatus = Format(.cmb_ContractState, "00")
        Call cmb_ContractState_AfterUpdate
        
        .txt_CancelAcceptDate = rsCarg.Fields("CARG_KAIUKEDATE")
        .cmb_CancelAcceptPerson = rsCarg.Fields("CARG_KAITANTO")
        Call cmb_CancelAcceptPerson_AfterUpdate
        
        .cmb_CancelDivision = rsCarg.Fields("CARG_KAICD")
        Call cmb_CancelDivision_AfterUpdate
        
        .cmb_CancelReason = rsCarg.Fields("CARG_KAIRIYUCD")
        Call cmb_CancelReason_AfterUpdate
        
        .txt_CancelDate = rsCarg.Fields("CARG_KYDATE")
        '状態区分「延長」選択時に解約予定年月に設定する
        strExistingCancelDate = Nz(rsCarg.Fields("CARG_KYDATE"))
        strLastCancelDate = Nz(rsCarg.Fields("CARG_KYDATE"))
        
        .txt_ContractDate = rsCarg.Fields("CARG_FSDATE")
        .txt_StartDate = rsCarg.Fields("CARG_STDATE")
        .txt_EndDate = rsCarg.Fields("CARG_EDDATE")
        .txt_CancelExpectedDate = rsCarg.Fields("CARG_CYDATE")
        
        txt_KeyDivision = rsCarg.Fields("KAGIICD_NAME")
        .txt_KeyReturnDate = rsCarg.Fields("CARG_KEY_RETDATE")
        
        .opg_SpareKey.VALUE = rsCarg.Fields("CARG_SPAREKEY")
        .chk_AccountFlag.VALUE = rsCarg.Fields("CARG_LIQ_FLG")

        '2006/09/19 SHIBAZAKI 予約名と予約番号削除
        '20051002 SHIBAZAKI 予約名と予約番号追加
'        .txt_BookingName = rsCarg.Fields("CNTA_YNAME")
'        .txt_BookingNo = rsCarg.Fields("CNTA_YTEL")

'        .cmb_NewKeyDivision = rsCarg.Fields("CARG_NEW_KEYCD")
'        Call cmb_NewKeyDivision_AfterUpdate
'        .txt_NewKeyNo = rsCarg.Fields("CARG_NEW_KEYNO")
      
        '2006/03/04 tajima キャンペーン適用満了日を追加
        txt_CampaignExpiration = rsCarg.Fields("CARG_CMP_EXDATE")
         ' キャンペーン適用満了日が入っていればボタン押下可　'I Love Japanese
        If Nz(txt_CampaignExpiration, "") <> "" Then
          cmd_キャンペーン内容.Enabled = True
        Else
          cmd_キャンペーン内容.Enabled = False
        End If

        '2006/05/13 tajima 鍵管理対応 Start
        .txt_返却本数 = rsCarg.Fields("CARG_KEY_RETNUM")
        .txt_スペアキー備考 = rsCarg.Fields("CARG_SPAREKEY_NOTE")
        .txt_貸出本数 = rsCarg.Fields("CARG_KEY_LENTNUM")
        .txt_貸出本数2 = rsCarg.Fields("CARG_KEY_LENTNUM")
        .txt_鍵番号 = rsCarg.Fields("CARG_KAIJB")
'↓ UPDATE 2016/05/20 MIYAMOTO
'        .txt_自動鍵区分 = rsCarg.Fields("AUTO_KAGI_KBN")            'INSERT 2014/04/11 MIYAMOTO
        .txt_自動鍵区分 = MSZZ070_KaiKey(txt_YardTihoKbn, rsCarg.Fields("AUTO_KAGI_KBN"))
'↑ UPDATE 2016/05/20 MIYAMOTO

'↓ UPDATE 2021/06/02 EGL        .txt_解除番号 = rsCarg.Fields("RCPT_KAINO")                 'INSERT 2014/04/11 MIYAMOTO
        .txt_解除番号 = Trim(rsCarg.Fields("RCPT_KAINO") & Space(1) & rsCarg.Fields("RCPT_BUIL_KAGI_NO"))
'↑ UPDATE 2021/06/02 EGL

'↓ UPDATE 2021/06/02 EGL       pstr現解除番号 = Nz(rsCarg.Fields("AKKNT_KAINO"), "")       'INSERT 2014/04/11 MIYAMOTO
       pstr現解除番号 = Trim(Nz(rsCarg.Fields("AKKNT_KAINO"), "") & Space(1) & Nz(rsCarg.Fields("AKKNT_KAINO_QR"), ""))
'↑ UPDATE 2021/06/02 EGL
        
        pstr元鍵区分 = Nz(rsCarg.Fields("CARG_KAGIICD"), "")
'↓ UPDATE 2016/05/20 MIYAMOTO
'        .txt_変換鍵区分 = Nz(rsCarg.Fields("CHANGE_KAGI_KBN"), "")  'INSERT 2014/06/17 MIYAMOTO
        .txt_変換鍵区分 = MSZZ070_KaiKey(txt_YardTihoKbn, Nz(rsCarg.Fields("CHANGE_KAGI_KBN"), ""))
'↑ UPDATE 2016/05/20 MIYAMOTO
        '2006/05/13 tajima 鍵管理対応 End

    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DispMaintenance
'        機能             :メンテナンス情報画面表示
'        IN               :rsCarg(検索結果レコードセット)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DispMaintenance(rsCarg As Recordset)

    With Me
        .txt_MaintenanceRequestDate = rsCarg.Fields("CARG_MNT_REQDATE")
        strMaintenanceRequestDate = Nz(rsCarg.Fields("CARG_MNT_REQDATE"))
        .txt_MaintenanceEndDate = rsCarg.Fields("CARG_MNT_CMPDATE")
        .cmb_MaintenancePerson = rsCarg.Fields("CARG_MNT_TANTO")
        Call cmb_MaintenancePerson_AfterUpdate
        
        .opg_BaggageDivision.VALUE = rsCarg.Fields("CARG_NIRCD")
        .opg_DisposeDivision.VALUE = rsCarg.Fields("CARG_SYOCD")
        .opg_RepairDivision.VALUE = rsCarg.Fields("CARG_HOSCD")
        .txt_MaintenanceMemo = rsCarg.Fields("CARG_MNT_MEMO")
        
        '2006/05/13 tajima 鍵管理対応 Start
        ' 新鍵（コンテナマスタの鍵）
        .txt_NewKeyNo = Nz(rsCarg.Fields("CNTA_KEY_NO"))
        .txt_新貸出本数 = Nz(rsCarg.Fields("CNTA_KEY_NUM"))
        .cmb_NewKeyDivision = rsCarg.Fields("CNTA_KEY_TYPE")
        Call cmb_NewKeyDivision_AfterUpdate
        '↓ INSERT 2014/04/11 MIYAMOTO
        If rsCarg.Fields("CARG_AGRE") <> 1 Then
            'UPDATE 2021/06/02 EGL .txt_新解除番号 = rsCarg.Fields("AKKNT_KAINO")
            .txt_新解除番号 = Trim(rsCarg.Fields("AKKNT_KAINO") & Space(1) & rsCarg.Fields("AKKNT_KAINO_QR"))
        End If
        '↑ INSERT 2014/04/11 MIYAMOTO
        ' 新鍵付け替え時の情報
        .txt_鍵変更日 = Nz(rsCarg.Fields("CARG_KEY_CHGDATE"))
        .cmb_鍵変更担当 = rsCarg.Fields("CARG_KEY_CHGTANTO")
        Call cmb_鍵変更担当_AfterUpdate
        .cmb_鍵変更理由 = rsCarg.Fields("CARG_KEY_CHGRIYUCD")
        Call cmb_鍵変更理由_AfterUpdate
        .txt_鍵変更備考 = rsCarg.Fields("CARG_KEY_CHGNOTE")
        .chk_鍵変更 = 0
        Call chk_鍵変更_AfterUpdate
        '2006/05/13 tajima 鍵管理対応 End
        
    End With

End Sub

'==============================================================================*
'
'        MODULE_NAME      :DispAccount
'        機能             :精算情報画面表示
'        IN               :rsCarg(検索結果レコードセット)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DispAccount(rsCarg As Recordset)

    Me.txt_AccountRequestDate = rsCarg.Fields("CARG_LIQ_REQDATE")
    Me.txt_AccountNote = rsCarg.Fields("CARG_LIQ_NOTE")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :GetCargCount
'        機能             :件数情報取得／画面表示
'        IN               :rsCarg(検索結果レコードセット)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetCargCount(strMntCompDate As String, lngUcode As Long) As Boolean
    
    Dim strSQL      As String
    Dim rsCount     As Recordset
    Dim strCount    As String
    
On Error GoTo GetCargCount_Err
  
    GetCargCount = False
  
    '対象データにメンテ完了日と顧客コードが格納されていない
    If strMntCompDate = "" Or lngUcode = 0 Then
        txt_Count = "0/0"
        GetCargCount = True
        GoTo GetCargCount_End
    End If
      
    doCmd.SetWarnings False
    
    '件数取得
    strSQL = MakeCargCountSql(Format(strMntCompDate, "yyyy/mm"), lngUcode)
    
  
    Set rsCount = objDb.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsCount.EOF Then
        MsgBox "コンテナ契約ファイル件数取得時にエラーが発生しました。", vbInformation, P_ERROR
    Else
        '最初のレコードに同月内メンテ件数が格納される
        strCount = Nz(rsCount.Fields("CARG_COUNT"), "0")
        rsCount.MoveNext
        If rsCount.EOF Then
            MsgBox "コンテナ契約ファイル件数取得時にエラーが発生しました。", vbInformation, P_ERROR
        End If
        '二番目のレコードに解約件数が格納される
        strCount = strCount & "/" & Nz(rsCount.Fields("CARG_COUNT"), "0")
        '画面表示
        txt_Count = strCount
    End If
    rsCount.Close
    Set rsCount = Nothing
    
    GetCargCount = True
    GoTo GetCargCount_End

GetCargCount_Err:
    If Not rsCount Is Nothing Then rsCount.Close: Set rsCount = Nothing
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear

GetCargCount_End:
    doCmd.SetWarnings True

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeCargCountSql
'        機能             :件数情報取得SQL作成
'        IN               :SQL文
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeCargCountSql(strMntCmpDate As String, lngUcode As Long) As String
    
    Dim strSQL  As String
    
    'データ取得
    strSQL = strSQL & " SELECT Count(*) AS CARG_COUNT "
    strSQL = strSQL & "   FROM CARG_FILE "
    strSQL = strSQL & "  WHERE YEAR(CARG_MNT_CMPDATE) = " & Format(strMntCmpDate, "yyyy") & " "
    strSQL = strSQL & "    AND MONTH(CARG_MNT_CMPDATE) = " & Format(strMntCmpDate, "m") & " "
    strSQL = strSQL & "    AND CARG_AGRE = " & P_STATUS_COMPLETE
    strSQL = strSQL & "    AND CARG_UCODE = " & lngUcode
    
    strSQL = strSQL & " UNION ALL "
    
    strSQL = strSQL & " SELECT Count(*) AS CARG_COUNT "
    strSQL = strSQL & "   FROM CARG_FILE "
    strSQL = strSQL & "  WHERE YEAR(CARG_KYDATE) = " & Format(strMntCmpDate, "yyyy") & " "
    strSQL = strSQL & "    AND MONTH(CARG_KYDATE) = " & Format(strMntCmpDate, "m") & " "
    strSQL = strSQL & "    AND CARG_UCODE = " & lngUcode

    MakeCargCountSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :ExistNewContract
'        機能             :表示データよりも新しい契約が存在するかチェック
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ExistNewContract(rsCarg As Recordset) As Boolean
On Error GoTo ExistNewContract_Err

    Dim strSQL          As String
    Dim rsNewContract   As Recordset

    ExistNewContract = False
    
    If rsCarg.Fields("CARG_CONTNA") = 0 Or _
       Nz(rsCarg.Fields("CARG_KYDATE")) = "" Or _
       strFormFrom <> P_解約一覧 Then
        Exit Function
    End If
    
    strSQL = MakeExistNewContractSql(rsCarg)
    
    Set rsNewContract = objDb.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If Not rsNewContract.EOF Then
        ExistNewContract = True
    End If
    rsNewContract.Close
    Set rsNewContract = Nothing
    
    GoTo ExistNewContract_End

ExistNewContract_Err:
    If Not rsNewContract Is Nothing Then rsNewContract.Close: Set rsNewContract = Nothing
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear

ExistNewContract_End:
    doCmd.SetWarnings True
    
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeExistNewContractSql
'        機能             :表示データよりも新しい契約が存在するかチェックSQL作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeExistNewContractSql(rsCarg As Recordset) As String

    Dim strSQL      As String
    
    '受注契約番号が自分よりも新しければ、それが新しい契約となる
    strSQL = strSQL & " SELECT * FROM CARG_FILE "
    strSQL = strSQL & "  WHERE CARG_YCODE = " & rsCarg.Fields("CARG_YCODE")
    strSQL = strSQL & "    AND CARG_NO = " & rsCarg.Fields("CARG_NO")
    strSQL = strSQL & "    AND CARG_ACPTNO > '" & rsCarg.Fields("CARG_ACPTNO") & "'"
    
    MakeExistNewContractSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :FormClear
'        機能             :画面クリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FormClear()

    Call ClearContainer(True)   'ヤードコードもクリア
    Call ClearExcludingContainer

End Sub


'==============================================================================*
'
'        MODULE_NAME      :ClearExcludingContainer
'        機能             :画面クリア(コンテナ情報以外)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClearExcludingContainer()
    
    Call ClearContract
    Call ClearMaintenance
    Call ClearAccount

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ClearContainer
'        機能             :画面クリア(コンテナ情報)
'        IN               :True:ヤードコードをクリアする False:ヤードコードはクリアしない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClearContainer(blnYardClr As Boolean)

    If blnYardClr Then
        cmb_Yard = ""
        txt_Yard = ""
        txt_Yadd = ""
    End If
    
    cmb_CntaNo = ""
    txt_STEP = ""
    txt_StepName = ""
    txt_CntaKind = ""
    txt_CntaKindName = ""
    txt_SIZE = ""         'ADD 2006/03/04 tajima

    '2006/09/19 SHIBAZAKI コンテナ／顧客の直近修繕表示
    txt_CntaRepr = ""

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ClearContract
'        機能             :画面クリア(契約情報)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClearContract()

    txt_ContractNo = ""
    txt_UserCode = ""
    txt_UserName = ""
    txt_UserKana = ""
    txt_Ukkbn = ""      'INSERT 2011/09/22 M.RYU
    
    cmb_ContractState = ""
    txt_ContractState = ""
    
    txt_CancelAcceptDate = ""
    cmb_CancelAcceptPerson = ""
    txt_CancelAcceptPerson = ""
    
    cmb_CancelDivision = ""
    txt_CancelDivision = ""
    
    cmb_CancelReason = ""
    txt_CancelReason = ""
    
    txt_CancelDate = ""
    txt_CancelExpectedDate = ""
    txt_ContractDate = ""
    txt_StartDate = ""
    txt_EndDate = ""
    
    txt_KeyDivision = ""
    txt_KeyReturnDate = ""
    
    opg_SpareKey.VALUE = Null
    
    '2006/09/19 SHIBAZAKI コンテナ／顧客の直近修繕表示
    txt_UserRepr = ""
    
    '2006/09/19 SHIBAZAKI 予約名／予約番号削除　コンテナ／顧客の直近修繕表示
    '20051004 SHIBAZAKI 項目追加
'    txt_BookingName = ""
'    txt_BookingNo = ""
   
    txt_Count = ""
    
    'ADD 2006/03/04 tajima キャンペーン適用満了日項目の追加
    txt_CampaignExpiration = ""
    cmd_キャンペーン内容.Enabled = False
    pstrCampaignInfomationText = ""
    pstrCampaignWarningText = ""

    'ADD 2006/05/13 tajima 鍵関係
    txt_貸出本数 = ""
    txt_返却本数 = ""
    txt_スペアキー備考 = ""
    txt_鍵番号 = ""
    txt_貸出本数2 = ""
    txt_自動鍵区分 = ""             'INSERT 2014/04/11 MIYAMOTO
    txt_解除番号 = ""               'INSERT 2014/04/11 MIYAMOTO
    txt_変換鍵区分 = ""             'INSERT 2014/06/17 MIYAMOTO
    txt_変換鍵自動鍵区分 = ""       'INSERT 2014/06/17 MIYAMOTO


End Sub

'==============================================================================*
'
'        MODULE_NAME      :ClearMaintenance
'        機能             :画面クリア(メンテナンス情報)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClearMaintenance()
    
    txt_MaintenanceRequestDate = ""
    txt_MaintenanceEndDate = ""
    cmb_MaintenancePerson = ""
    txt_MaintenancePerson = ""
    
    opg_BaggageDivision.VALUE = Null
    opg_DisposeDivision.VALUE = Null
    opg_RepairDivision.VALUE = Null
    txt_MaintenanceMemo = ""

    ' 2006/05/19 tajima Add
    cmb_NewKeyDivision = ""
    txt_NewKeyDivision = ""
    txt_NewKeyNo = ""
    txt_新貸出本数 = ""
    txt_新解除番号 = ""                 'INSERT 2014/04/11 MIYAMOTO
    txt_鍵変更日 = ""
    cmb_鍵変更担当 = ""
    txt_鍵変更担当 = ""
    cmb_鍵変更理由 = ""
    txt_鍵変更理由 = ""
    txt_鍵変更備考 = ""
    chk_鍵変更 = 0
    Call chk_鍵変更_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ClearAccount
'        機能             :画面クリア(精算情報)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClearAccount()
    
    txt_AccountRequestDate = ""
    txt_AccountNote = ""
    chk_AccountFlag.VALUE = 0

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ModeChange
'        機能             :画面モード変更
'        IN               :画面モード
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ModeChange(iMode As Integer)

    Select Case iMode
    Case P_ヤード選択前
        'ヤードコード指定待ちの状態
        '画面クリア／ヤードコード以外使用不可
        Call FormClear
        
        Call EnableContainer(False)
        Call EnableContract(False)
        Call EnableMaintenance(False)
        Call EnableAccount(False)
        Call EnableButton(False)
        cmb_Yard.SetFocus
    Case P_検索前
        'ヤードコードが選択され、コンテナ番号指定待ちの状態
        'コンテナ番号以下クリア／ヤードコードおよびコンテナ番号以外使用不可
        Call ClearContainer(False)
        Call ClearExcludingContainer
        
        Call EnableContainer(True)
        Call EnableContract(False)
        Call EnableMaintenance(False)
        Call EnableAccount(False)
        Call EnableButton(False)
        cmb_CntaNo.SetFocus
    Case P_検索後
        'コンテナ契約トラン検索／データ画面表示が終わった状態
        '全て使用可
'        Call EnableContainer(True)
'        Call EnableContract(True)
'        Call EnableMaintenance(True)
'        Call EnableAccount(True)
        Call EnableButton(True)
'        cmb_ContractState.SetFocus
    Case P_参照モード
        '入力テキストボックスを黄色/使用不可/編集ロックにしてコンボを使用不可にする
        'ボタンは終了のみ使用可にする
        cmd_End.SetFocus
        Call EnableContract(False)
        Call EnableMaintenance(False)
        Call EnableAccount(False)
        Call EnableButton(False)
        Call ReferenceMode
        SubForm.Enabled = False '2006/06/01 tajima Add
    End Select
    intProcMode = iMode
    
    blnChangeMode = False
    Me.Caption = strFormCaption
    
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EnableContainer
'        機能             :コントロール使用可不可(コンテナ情報)
'        IN               :True=使用可 False=使用不可
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub EnableContainer(blnEnable As Boolean)

    If Not blnContractMode Then
        cmb_CntaNo.Enabled = blnEnable
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EnableContract
'        機能             :コントロール使用可不可(契約情報)
'        IN               :第一引数：True=使用可 False=使用不可
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub EnableContract(blnEnable As Boolean, Optional blnIgnoreState As Boolean = False)

    If Not blnIgnoreState Then
        cmb_ContractState.Enabled = blnEnable
    End If
    
    txt_CancelAcceptDate.Enabled = blnEnable
    cmb_CancelAcceptPerson.Enabled = blnEnable
    
    cmb_CancelDivision.Enabled = blnEnable
    
    cmb_CancelReason.Enabled = blnEnable
    
    txt_CancelDate.Enabled = blnEnable
    txt_CancelExpectedDate.Enabled = blnEnable
    
    txt_KeyReturnDate.Enabled = blnEnable
    
    opg_SpareKey.Enabled = blnEnable
    
'    '2006/09/19 SHIBAZAKI スペアキー備考を制御
'    txt_スペアキー備考.Enabled = blnEnable
    
    '2006/09/19 SHIBAZAKI 予約名／予約番号削除
    '20051004 SHIBAZAKI 項目追加
'    txt_BookingName.Enabled = blnEnable
'    txt_BookingNo.Enabled = blnEnable
    
    '2006/05/13 tajima 鍵対応
    txt_返却本数.Enabled = blnEnable
   
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EnableMaintenance
'        機能             :コントロール使用可不可(メンテナンス情報)
'        IN               :True=使用可 False=使用不可
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub EnableMaintenance(blnEnable As Boolean)

    txt_MaintenanceRequestDate.Enabled = blnEnable
    txt_MaintenanceEndDate.Enabled = blnEnable
    cmb_MaintenancePerson.Enabled = blnEnable
    
    opg_BaggageDivision.Enabled = blnEnable
    opg_DisposeDivision.Enabled = blnEnable
    opg_RepairDivision.Enabled = blnEnable
    txt_MaintenanceMemo.Enabled = blnEnable

    '2006/05/13 tajima add
    chk_鍵変更.Enabled = blnEnable
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EnableAccount
'        機能             :コントロール使用可不可(精算情報)
'        IN               :True=使用可 False=使用不可
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub EnableAccount(blnEnable As Boolean)

    txt_AccountRequestDate.Enabled = blnEnable
    txt_AccountNote.Enabled = blnEnable
    chk_AccountFlag.Enabled = blnEnable

End Sub

'==============================================================================*
'
'        MODULE_NAME      :EnableButton
'        機能             :コントロール使用可不可(ボタン)
'        IN               :True=使用可 False=使用不可
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub EnableButton(blnEnable As Boolean)

    cmd_Enter.Enabled = blnEnable

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ReferenceMode
'        機能             :参照モード設定
'        IN               :省略可　True=参照モードに設定(規定値) False=完了※に設定
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ReferenceMode(Optional blnReference As Boolean = True)

    '入力コントロールを黄色化
    Call YellowControl(txt_CancelAcceptDate)
    Call YellowControl(txt_CancelDate)
    Call YellowControl(txt_CancelExpectedDate)
    Call YellowControl(txt_KeyReturnDate)
    Call YellowControl(opg_SpareKey)

    '2006/09/19 SHIBAZAKI 予約名／予約番号削除
    '20051004 SHIBAZAKI 項目追加
'    Call YellowControl(txt_BookingName)
'    Call YellowControl(txt_BookingNo)

    Call YellowControl(txt_MaintenanceRequestDate)
    Call YellowControl(txt_MaintenanceEndDate)
    Call YellowControl(opg_BaggageDivision)
    Call YellowControl(opg_DisposeDivision)
    Call YellowControl(opg_RepairDivision)
    Call YellowControl(txt_MaintenanceMemo)
    
    If blnReference Then
        Call YellowControl(txt_AccountRequestDate)
        Call YellowControl(txt_AccountNote)
        
        '取り消しボタン使用不可
        cmd_Kesi.Enabled = False
        
        '画面タイトルに”（参照モード）”追加
        Me.Caption = Me.Caption & P_REFERENCE_MODE
        strFormCaption = Me.Caption
    End If
    
    ' 2006/05/13 Tajima ADD
    Call YellowControl(txt_スペアキー備考)
    Call YellowControl(txt_返却本数)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :YellowControl
'        機能             :コントロール黄色化
'        IN               :第一引数：対象コントロール
'                         :第二引数：省略可 True=黄色化（規定値） False=白化
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub YellowControl(objTarget As Object, Optional blnYellow As Boolean = True)

    If blnYellow Then
        '背景色を黄色
        objTarget.BackColor = P_BACKCOLOR_YELLOW
    Else
        '背景色を白
        objTarget.BackColor = P_BACKCOLOR_WHITE
    End If
    
    '編集ロック
    objTarget.Enabled = Not blnYellow
    objTarget.Locked = blnYellow

End Sub

'==============================================================================*
'
'        MODULE_NAME      :DB_Access
'        機能             :DBAccess処理
'        IN               :
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DB_Access() As Boolean
    
On Error GoTo DB_Access_On_Err
    
    doCmd.Hourglass True
    DB_Access = False
    doCmd.SetWarnings False

    'DBｱｸｾｽ処理
    If DB_AccessMain() = False Then GoTo DB_Access_End
    
    DB_Access = True
    
    GoTo DB_Access_End
    
DB_Access_On_Err:
    MsgBox "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr _
            & Err.Description, vbCritical + vbOKOnly, P_ERROR
    Err.Clear

DB_Access_End:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB_AccessMain
'        機能             :データ更新処理
'        IN               :
'        OUT              :結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DB_AccessMain() As Boolean

On Error GoTo DB_AccessMain_Err

    Dim strSQL      As String
    Dim strUkeNo    As String
    Dim strUser     As String
    Dim strwork     As String   '顧客カナ・代表者カナ対応
    Dim errLoop     As Error
    Dim rsCarg      As DAO.Recordset
    Dim objWs       As Workspace
    Dim strMsg      As String
    Dim intCntaUse  As Integer
    
    DB_AccessMain = False
    
    strMsg = ""
    
    strSQL = strSQL & "SELECT * FROM CARG_FILE "
    strSQL = strSQL & " WHERE CARG_ACPTNO = '" & txt_ContractNo & "' "
    
    Set objDb = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    '20051002 SHIBAZAKI トランザクション処理追加
    Set objWs = DBEngine.Workspaces(0)
    objWs.BeginTrans
     
    Set rsCarg = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rsCarg.EOF Then
        strMsg = "コンテナ契約ファイル更新情報読み込み中にエラーが発生しました"
        GoTo DB_AccessMain_Err
    End If
    
    With rsCarg
        .Edit
        Call SetCargFile(rsCarg)
        .UPDATE
    End With
    
    rsCarg.Close
    Set rsCarg = Nothing
    
    '状態区分「完了」で保存する場合、紹介待ち情報を作成する
    '契約番号の先頭に"@"をつけて予約番号として共通関数をCALLする。
    '完了から他のものに変更する場合は、コンテナの利用可否を「可」にする。
    
    '2005/12/8 SHIBAZAKI 追加↓
    If cmb_ContractState = P_STATUS_COMPLETE And strExistingStatus <> P_STATUS_COMPLETE Then
        '紹介待ち作成
        intCntaUse = 1
    ElseIf cmb_ContractState <> P_STATUS_COMPLETE And strExistingStatus = P_STATUS_COMPLETE Then
        '紹介待ち削除
        intCntaUse = 2
    Else
        '紹介待ちは操作しない
        intCntaUse = 0
    End If
    '2005/12/8 SHIBAZAKI 追加↑
    
    'TODO 一時機能カット
    '状態区分「完了」で保存する場合、予約自動取置き処理を行う
'    If cmb_ContractState = P_STATUS_COMPLETE And strExistingStatus <> P_STATUS_COMPLETE Then
'        Call AutoKeep
'    End If
    
    '20051004 SHIBAZAKI コンテナマスタの予約名と予約番号を更新する &鍵もね by Taddy
    If Not SetCntaMast(intCntaUse, strMsg) Then
        GoTo DB_AccessMain_Err
    End If
    
    '2006/05/27 セキュリティカードの更新をする
    If SetSecurityCardMast() = False Then
        GoTo DB_AccessMain_Err
    End If
       
    '↓INSERT 2014/04/11 MIYAMOTO
'↓ UPDATE 2014/06/17 MIYAMOTO
'    '状態区分「解約予定」で鍵区分が自動鍵、且つ現解除番号と新解除番号が異なった状態で保存の場合、発番された新解除番号の登録を行う
'    If cmb_ContractState = P_STATUS_CANCEL And _
'       txt_自動鍵区分 = "1" And _
'       pstr現解除番号 <> txt_新解除番号 Then
'        If Not SetAkknTran() Then
'            GoTo DB_AccessMain_Err
'        End If
'    End If
    '状態区分「解約予定」の場合
    'If cmb_ContractState = P_STATUS_CANCEL Then    '2014/08/02 M.HONDA 地方ヤードの場合には処理しない UPD
    If cmb_ContractState = P_STATUS_CANCEL And Nz(txt_YardTihoKbn) <> P_TIHO_KBN Then     '2014/08/02 M.HONDA 地方ヤードの場合には処理しない UPD
    
        '元データが自動鍵の場合、現解除番号と新解除番号が異なった時、発番された新解除番号で自動鍵解除番号発番トランの登録を行う
        If txt_自動鍵区分 = "1" And _
           pstr現解除番号 <> txt_新解除番号 Then
            If Not SetAkknTran() Then
                GoTo DB_AccessMain_Err
            End If
        '元データが自動鍵以外の場合、変換された鍵区分が自動鍵の時、発番された新解除番号で自動鍵解除番号発番トランの登録を行う
        Else
            If Nz(txt_変換鍵自動鍵区分) = "1" Then
                If Not SetAkknTran() Then
                    GoTo DB_AccessMain_Err
                End If
            End If
        End If
    End If
'↑ UPDATE 2014/06/17 MIYAMOTO
    '↑INSERT 2014/04/11 MIYAMOTO
    
'↓ INSERT 2016/05/20 MIYAMOTO
    If chk_鍵変更.VALUE = -1 Then
        '状態区分が「メンテ中」又は「完了」の場合
        If cmb_ContractState = P_STATUS_MAINTENANCE Or cmb_ContractState = P_STATUS_COMPLETE Then
            '鍵区分が手動解除番号設定鍵で新解除番号が入力・変更されていた場合、入力された新解除番号で自動鍵解除番号発番トランの登録を行う
            If MSZZ070_KaiKey(txt_YardTihoKbn, cmb_NewKeyDivision.Column(2)) = "2" And _
               txt_新解除番号 <> "" And _
               pstr現解除番号 <> txt_新解除番号 Then
                If Not SetAkknTran() Then
                    GoTo DB_AccessMain_Err
                End If
            End If
        End If
    End If
'↑ INSERT 2016/05/20 MIYAMOTO
    
    '20051002 SHIBAZAKI トランザクション処理追加
    objWs.CommitTrans
    
    ' cmb_ContractState(現在の入力)が 09 かつ、
    ' strExistingStatus(DBから読み込んだ時の値)が 09 以外の場合のみ実行
    If cmb_ContractState = P_STATUS_COMPLETE And strExistingStatus <> P_STATUS_COMPLETE Then
        Call CallRoomReleaseAPI
    End If
    
    objDb.Close
    Set objDb = Nothing
    
    doCmd.SetWarnings False
    DB_AccessMain = True
    
    GoTo DB_AccessMain_End
    
DB_AccessMain_Err:
    '20051002 SHIBAZAKI トランザクション処理追加
    If Not objWs Is Nothing Then
        objWs.Rollback
    End If
    
    If strMsg = "" Then
        strMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr _
                & Err.Description
    End If
    
    Err.Clear

DB_AccessMain_End:
    If Not rsCarg Is Nothing Then rsCarg.Close: Set rsCarg = Nothing
    '20051002 SHIBAZAKI トランザクション処理追加
    If Not objWs Is Nothing Then Set objWs = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    
    If strMsg <> "" Then
        MsgBox strMsg, vbCritical + vbOKOnly, P_ERROR
    End If
    
End Function

'↓INSERT 2014/04/11 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :SetAkknTran
'        機能             :自動鍵解除番号発番トラン登録処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SetAkknTran() As Boolean
    Dim strSQL          As String
    Dim isException     As Boolean
    Dim rsAkkn          As Object           ' 自動鍵解除番号発番トランレコードセット
'↓ 2021/06/02 insert EGL
    Dim strBUMOC        As String           ' 部門コード
    Dim dbSQLServer     As Object
'↑ 2021/06/02 insert EGL
On Error GoTo SetAkknTran_Err

    SetAkknTran = False
    isException = False
    
'↓ 2021/06/02 insert EGL
    strBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    'SQLサーバー接続
    Set dbSQLServer = ADODB_Connection(strBUMOC)
'↑ 2021/06/02 insert EGL

    strSQL = ""
    strSQL = strSQL & "SELECT      *" & vbCrLf
    strSQL = strSQL & "FROM        AKKN_TRAN" & vbCrLf
    strSQL = strSQL & "WHERE       AKKNT_YARD = " & Val(cmb_Yard.VALUE) & vbCrLf
    strSQL = strSQL & "AND         AKKNT_CTNO = " & Val(cmb_CntaNo.VALUE) & vbCrLf
    strSQL = strSQL & "AND         AKKNT_KTYPE = '" & zeroByteToNull(cmb_NewKeyDivision) & "'" & vbCrLf
    strSQL = strSQL & "AND         AKKNT_STYPE = '4'" & vbCrLf 'INSERT 2021/06/04 EGL

'UPDATE 2021/06/02 EGL    Set rsAkkn = objDb.OpenRecordset(strSql, dbOpenDynaset)
    Set rsAkkn = ADODB_Recordset(strSQL, dbSQLServer, adoAppendOnly) '2021/06/02 insert EGL

    '自動鍵解除番号発番トランにデータ登録
    With rsAkkn
        .AddNew

        .Fields("AKKNT_YARD") = Val(cmb_Yard.VALUE)                                         ' ヤードコード
        .Fields("AKKNT_CTNO") = Val(cmb_CntaNo.VALUE)                                       ' コンテナ番号
        .Fields("AKKNT_KTYPE") = zeroByteToNull(cmb_NewKeyDivision)                         ' 鍵区分
        .Fields("AKKNT_HATUD") = Format$(DATE, "YYYYMMDD") & Format$(time, "HHMMSS")        ' 解除番号発番日時
'UPDATE 2021/06/02 EGL        .Fields("AKKNT_KAINO") = Nz(txt_新解除番号)                                         ' 解除番号
        .Fields("AKKNT_KAINO") = Nz(Mid(txt_新解除番号, 1, 4))                              ' 解除番号
'INSERT 2021/06/04 EGL
        .Fields("AKKNT_STYPE") = "4"
        .Fields("AKKNT_INSED") = Format$(DATE, "YYYYMMDD")                                  ' 作成日付
        .Fields("AKKNT_INSEJ") = Format$(time, "HHMMSS")                                    ' 作成時刻
        .Fields("AKKNT_INSPB") = PROG_ID                                                    ' 作成プログラムID
        .Fields("AKKNT_INSUB") = LsGetUserName()                                            ' 作成ユーザーID

        .UPDATE
    End With
    
'↓ INSERT 2021/06/04 EGL
    If cmb_NewKeyDivision = 16 And Len(Mid(txt_新解除番号, 6, 6)) = 6 Then
        strSQL = ""
        strSQL = strSQL & "SELECT      *" & vbCrLf
        strSQL = strSQL & "FROM        AKKN_TRAN" & vbCrLf
        strSQL = strSQL & "WHERE       AKKNT_YARD = " & Val(cmb_Yard.VALUE) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_CTNO = " & Val(cmb_CntaNo.VALUE) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_KTYPE = '" & zeroByteToNull(cmb_NewKeyDivision) & "'" & vbCrLf
        strSQL = strSQL & "AND         AKKNT_STYPE = '4'" & vbCrLf 'INSERT 2021/06/04 EGL
        Set rsAkkn = ADODB_Recordset(strSQL, dbSQLServer, adoAppendOnly) '2021/06/02 insert EGL
        '自動鍵解除番号発番トランにデータ登録
        With rsAkkn
            .AddNew
            .Fields("AKKNT_YARD") = Val(cmb_Yard.VALUE)                                         ' ヤードコード
            .Fields("AKKNT_CTNO") = Val(cmb_CntaNo.VALUE)                                       ' コンテナ番号
            .Fields("AKKNT_KTYPE") = zeroByteToNull(cmb_NewKeyDivision)                         ' 鍵区分
            .Fields("AKKNT_HATUD") = Format$(DATE, "YYYYMMDD") & Format$(time, "HHMMSS")        ' 解除番号発番日時
            .Fields("AKKNT_KAINO") = Nz(Mid(txt_新解除番号, 6, 6))                              ' 解除番号
            .Fields("AKKNT_STYPE") = "6"
            .Fields("AKKNT_INSED") = Format$(DATE, "YYYYMMDD")                                  ' 作成日付
            .Fields("AKKNT_INSEJ") = Format$(time, "HHMMSS")                                    ' 作成時刻
            .Fields("AKKNT_INSPB") = PROG_ID                                                    ' 作成プログラムID
            .Fields("AKKNT_INSUB") = LsGetUserName()                                            ' 作成ユーザーID
            .UPDATE
        End With
    End If
'↑ INSERT 2021/06/04 EGL

    SetAkknTran = True

    GoTo SetAkknTran_End

SetAkknTran_Err:
    isException = True

SetAkknTran_End:
    If Not rsAkkn Is Nothing Then rsAkkn.Close: Set rsAkkn = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "SetAkknTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'↑INSERT 2014/04/11 MIYAMOTO

''==============================================================================*
''
''        MODULE_NAME      :AutoKeep
''        機能             :予約自動取置き
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function AutoKeep() As Boolean
'
'On Error GoTo AutoKeep_Err
'
'    Dim strSql      As String
'    Dim rsYouk      As Recordset
'    Dim intToriDd   As Integer
'
'    AutoKeep = True
'
'    '取置き日数をコントロールマスタから取得
'    intToriDd = Nz(DLookup("CONT_TORI_DD", "dbo_CONT_MAST", "CONT_KEY=1"))
'
'    '予約受付トランから同コンテナの「予約確定」データを取得
'    strSql = ""
'    strSql = strSql & " SELECT TOP 1 * FROM YOUK_TRAN "
'    strSql = strSql & "  WHERE YOUKT_YCODE = " & cmb_Yard
'    strSql = strSql & "    AND YOUKT_NO = " & cmb_CntaNo
'    strSql = strSql & "    AND YOUKT_YUKBN = " & P_ACCEPT_STATUS_DECISION
'    strSql = strSql & "  ORDER BY YOUKT_UKDATE, YOUKT_UKNO "
'
'    Set rsYouk = objDb.OpenRecordset(strSql, dbOpenDynaset)
'
'    If Not rsYouk.EOF Then
'        With rsYouk
'            .Edit
'            '予約受付状態区分
'            .Fields("YOUKT_YUKBN") = P_ACCEPT_STATUS_KEEP
'            '自動区分
'            .Fields("YOUKT_AUTOKBN") = P_ACCEPT_STATUS_KEEP
'            '契約受付日
'            .Fields("YOUKT_KUDATE") = DATE
'            '取置き期限日
'            .Fields("YOUKT_TKDATE") = DateAdd("d", intToriDd, DATE)
'            .Fields("YOUKT_UPDAD") = Format$(DATE, "yyyymmdd")          '更新日付(var/Null=Y)
'            .Fields("YOUKT_UPDAJ") = Format$(Time, "hhmmss")            '更新時刻(var/Null=Y)
'            .Fields("YOUKT_UPDPB") = Me.NAME                            '更新プログラムID(var/Null=Y)
'            .Fields("YOUKT_UPDUB") = LsGetUserName()                    '更新ユーザーID(var/Null=Y)
'            .Update
'        End With
'    End If
'
'    GoTo AutoKeep_End
'
'AutoKeep_Err:
'    MsgBox "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr _
'            & Err.Description, vbCritical + vbOKOnly, P_ERROR
'
'AutoKeep_End:
'    If Not rsYouk Is Nothing Then rsYouk.Close: Set rsYouk = Nothing
'
'End Function
'
''==============================================================================*
''
''        MODULE_NAME      :PossibleKeep
''        機能             :自動取置き可否チェック
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function PossibleKeep() As Boolean
'
'On Error GoTo PossibleKeep_Err
'
'    Dim strSql      As String
'    Dim rsYouk      As Recordset
'
'    PossibleKeep = True
'
'    strSql = strSql & " SELECT * FROM YOUK_TRAN "
'    strSql = strSql & "  WHERE YOUKT_YCODE = " & cmb_Yard
'    strSql = strSql & "    AND YOUKT_NO = " & cmb_CntaNo
'    strSql = strSql & "    AND ( YOUKT_YUKBN = " & P_ACCEPT_STATUS_KEEP
'    strSql = strSql & "       OR YOUKT_YUKBN = " & P_ACCEPT_STATUS_CONTRACT & " ) "
'
'    Set rsYouk = objDb.OpenRecordset(strSql, dbOpenDynaset)
'
'    If Not rsYouk.EOF Then
'        PossibleKeep = False
'    End If
'
'    GoTo PossibleKeep_End
'
'PossibleKeep_Err:
'    MsgBox "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr _
'            & Err.Description, vbCritical + vbOKOnly, P_ERROR
'
'PossibleKeep_End:
'    If Not rsYouk Is Nothing Then rsYouk.Close: Set rsYouk = Nothing
'
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :SetCargFile
'        機能             :データ設定処理
'        IN               :更新対象レコードセット
'        OUT              :既存データ状態区分
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetCargFile(rsCarg As Recordset)

    With rsCarg
    
        .Fields("CARG_AGRE") = cmb_ContractState
        
        '契約状態が「完了」の場合、コンテナ状態を－１にする
        If cmb_ContractState = P_STATUS_COMPLETE Then
            .Fields("CARG_CONTNA") = -1
        Else
            .Fields("CARG_CONTNA") = 0
        End If
        
        .Fields("CARG_KAIUKEDATE") = zeroByteToNull(txt_CancelAcceptDate)
        .Fields("CARG_KAITANTO") = zeroByteToNull(cmb_CancelAcceptPerson)
        .Fields("CARG_KAICD") = zeroByteToNull(cmb_CancelDivision)
        .Fields("CARG_KAIRIYUCD") = zeroByteToNull(cmb_CancelReason)
        .Fields("CARG_KYDATE") = zeroByteToNull(txt_CancelDate)
        .Fields("CARG_KEY_RETDATE") = zeroByteToNull(txt_KeyReturnDate)
        .Fields("CARG_SPAREKEY") = opg_SpareKey.VALUE
        .Fields("CARG_CYDATE") = zeroByteToNull(txt_CancelExpectedDate)
        
        .Fields("CARG_MNT_REQDATE") = zeroByteToNull(txt_MaintenanceRequestDate)
        .Fields("CARG_MNT_CMPDATE") = zeroByteToNull(txt_MaintenanceEndDate)
        .Fields("CARG_MNT_TANTO") = Nz(cmb_MaintenancePerson, Null)
        
        .Fields("CARG_NIRCD") = opg_BaggageDivision.VALUE
        .Fields("CARG_SYOCD") = opg_DisposeDivision.VALUE
        .Fields("CARG_HOSCD") = opg_RepairDivision.VALUE
        .Fields("CARG_MNT_MEMO") = Nz(txt_MaintenanceMemo, Null)
    
        .Fields("CARG_LIQ_REQDATE") = zeroByteToNull(txt_AccountRequestDate)
        .Fields("CARG_LIQ_NOTE") = Nz(txt_AccountNote, Null)
        
        .Fields("CARG_LIQ_FLG") = chk_AccountFlag.VALUE
        
        .Fields("CARG_UPDATE") = Now
        
        ' 2006/05/13 Add tajima Start
        .Fields("CARG_SPAREKEY_NOTE") = Nz(Trim(Me.txt_スペアキー備考), "")
        'スペアキー有無しを指定してなければスペアキー備考入力有無によって有無を設定しちゃう
        If IsNull(.Fields("CARG_SPAREKEY")) Then
            If .Fields("CARG_SPAREKEY_NOTE") <> "" Then
                .Fields("CARG_SPAREKEY") = 0  'ｽﾍﾟｱｷｰ備考に何か入力があればｽﾍﾟｱｷｰ無し
            Else
                .Fields("CARG_SPAREKEY") = 1 'ｽﾍﾟｱｷｰ備考に何も入力なければｽﾍﾟｱｷｰ有り
            End If
        End If
        .Fields("CARG_KEY_RETNUM") = zeroByteToNull(txt_返却本数)
        .Fields("CARG_KEY_CHGDATE") = zeroByteToNull(txt_鍵変更日)
        .Fields("CARG_KEY_CHGTANTO") = zeroByteToNull(cmb_鍵変更担当)
        .Fields("CARG_KEY_CHGRIYUCD") = zeroByteToNull(cmb_鍵変更理由)
        .Fields("CARG_KEY_CHGNOTE") = zeroByteToNull(txt_鍵変更備考)
        ' 2006/05/13 Add tajima End
        
        .Fields("CARG_UPDAD") = Format$(DATE, "yyyymmdd")          '更新日付(var/Null=Y)
        .Fields("CARG_UPDAJ") = Format$(time, "hhmmss")            '更新時刻(var/Null=Y)
        .Fields("CARG_UPDPB") = Me.NAME                            '更新プログラムID(var/Null=Y)
        .Fields("CARG_UPDUB") = LsGetUserName()                    '更新ユーザーID(var/Null=Y)
    End With
    
End Sub

'20051004 SHIBAZAKI コンテナマスタの予約名と予約番号を更新する
'==============================================================================*
'
'        MODULE_NAME      :SetCargFile
'        機能             :データ設定処理
'        IN               :更新対象レコードセット
'        OUT              :既存データ状態区分
'
'==============================================================================*
'        UPDATE_DATE      :2017/05/06 K.SATO
'                          intCntaUse=1（完了）のときに用途＝コンテナバイクならコンテナにする
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2005/12/8 SHIBAZAKI パラメータ追加
'Private Function SetCntaMast() As Boolean
Private Function SetCntaMast(intCntaUse As Integer, ByRef strMsg As String) As Boolean

On Error GoTo SetCntaMast_Err

    Dim strSQL      As String
    Dim rsCnta      As DAO.Recordset
    
    SetCntaMast = False
    
    strSQL = strSQL & "SELECT * FROM CNTA_MAST "
    strSQL = strSQL & " WHERE CNTA_CODE = " & Val(cmb_Yard.VALUE)
    strSQL = strSQL & "   AND CNTA_NO = " & Val(cmb_CntaNo.VALUE)
     
    Set rsCnta = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rsCnta.EOF Then
        strMsg = "コンテナマスタ更新情報読み込み中にエラーが発生しました"
        GoTo SetCntaMast_End
    End If
    
    With rsCnta
        .Edit
        
        '2006/09/19 SHIBAZAKI 予約名／予約番号削除
'        .Fields("CNTA_YNAME") = zeroByteToNull(txt_BookingName)
'        .Fields("CNTA_YTEL") = zeroByteToNull(txt_BookingNo)
        
        '2005/12/8 SHIBAZAKI 追加↓
        '利用可否を更新するか？
        If intCntaUse = 1 Then
            '利用可否を「否」に更新
            .Fields("CNTA_USE") = P_CNTA_USE_NO
            .Fields("CNTA_UKNO") = P_UKNO_HEAD & txt_ContractNo
            
'DELETE 2018/08/29 N.IMAI Start
'            'INS 2017/05//6 K.SATO Start
'            If .Fields("CNTA_USAGE") = P_USAGE_KONTENABIKE Then
'                .Fields("CNTA_USAGE") = P_USAGE_KONTENA
'            End If
'            'INS 2017/05//6 K.SATO End
'DELETE 2018/08/29 N.IMAI End
        
        ElseIf intCntaUse = 2 Then
            '利用可否を「可」に更新
            .Fields("CNTA_USE") = P_CNTA_USE_OK
            .Fields("CNTA_UKNO") = Null
        End If
        '2005/12/8 SHIBAZAKI 追加↑
        
        'INSERT 2018/08/29 N.IMAI Start
        If cmb_ContractState = P_STATUS_COMPLETE Then
            If .Fields("CNTA_USAGE") = P_USAGE_KONTENABIKE Then
                .Fields("CNTA_USAGE") = P_USAGE_KONTENA
            End If
        End If
        'INSERT 2018/08/29 N.IMAI End
        
        .Fields("CNTA_UPDATE") = DATE
        
        ' 2006/05/13 Add tajima Start
        ' 鍵を変更するならば新項目を設定する
'↓ UPDATE 2014/06/17 MIYAMOTO
'        If chk_鍵変更.VALUE = -1 Then
        If Not ((cmb_ContractState = P_STATUS_MAINTENANCE Or cmb_ContractState = P_STATUS_COMPLETE) And (chk_鍵変更.VALUE <> -1)) Then
'↑ UPDATE 2014/06/17 MIYAMOTO
          .Fields("CNTA_KEY_TYPE") = zeroByteToNull(cmb_NewKeyDivision)
          .Fields("CNTA_KEY_NO") = zeroByteToNull(txt_NewKeyNo)
          .Fields("CNTA_KEY_NUM") = Nz(txt_新貸出本数, 0)
        End If                                     'DELETE 2014/06/17 MIYAMOTO
        ' 2006/05/13 Add tajima End
        .UPDATE
    End With
    
    SetCntaMast = True
    
    GoTo SetCntaMast_End

SetCntaMast_Err:
    strMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description
    Err.Clear

SetCntaMast_End:
    If Not rsCnta Is Nothing Then rsCnta.Close: Set rsCnta = Nothing
    
End Function

Private Function zeroByteToNull(varInput As Variant) As Variant

    zeroByteToNull = varInput
    
    If varInput = "" Then
        zeroByteToNull = Null
    End If
    
End Function
    
'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function dateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    dateCheck = True
    
    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        dateCheck = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :DateCheckYM
'        機能             :年月入力チェック＆フォーマット変換
'        IN               :(1)対象テキスト
'                          (2)True=年4桁に変換　False=年2桁に変換
'        OUT              :パラメータで指定された形式の年月。値が不正の場合NULLが返る
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DateCheckYM(txtTarget As Variant, blnLongDate As Boolean) As Variant
    
    '正しい入力方は下記のみ
    '(1)年月のスラッシュ区切り
    '   年月の桁は問わない
    '(2)yyyymm
    '(3)yymm
    
    DateCheckYM = ""
    
    If Nz(txtTarget) = "" Then
        Exit Function
    End If

    Dim varDate As Variant
    
    varDate = GetYyMm(txtTarget)
    If IsNull(varDate) Then
        GoTo DateCheckYM_InputError
    End If
    
    '入力値をyy年mm月01日の形に合わせる
    varDate = varDate & "01日"
    
    'VB の関数で日付と判断できる場合はOK
    If IsDate(varDate) Then
        If blnLongDate Then
            DateCheckYM = Format(varDate, "yyyy/mm")
        Else
            DateCheckYM = Format(varDate, "yy/mm")
        End If
        Exit Function
    End If
    
DateCheckYM_InputError:
    MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
    DateCheckYM = Null

End Function

'==============================================================================*
'
'        MODULE_NAME      :GetYyMm
'        機能             :年月取得　yy年mm月に変換する
'        IN               :対象テキスト
'        OUT              :値が不正の場合NULLが返る
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetYyMm(txtTarget As Variant) As Variant

    If Nz(txtTarget) = "" Then
        GetYyMm = ""
        Exit Function
    End If
    
    Dim lngInStr As Long
    Dim lngYearLen As Long
    Dim lngMonthLen As Long
    
    'スラッシュを探す
    lngInStr = InStr(txtTarget, "/")
    If lngInStr = 0 And Len(txtTarget) = 4 Then
        'スラ無し４桁　ｙｙｍｍと入力されたと判定
        lngYearLen = 2
        lngMonthLen = 3
    ElseIf lngInStr = 0 And Len(txtTarget) = 6 Then
        'スラ無し６桁　ｙｙｙｙｍｍと入力されたと判定
        lngYearLen = 4
        lngMonthLen = 5
    ElseIf lngInStr <> 0 Then
        'スラあり　スラ前を年、スラ後を月と判定
        lngYearLen = lngInStr - 1
        lngMonthLen = lngInStr + 1
    Else
        GoTo GetYyMm_InputError
    End If
    
    '入力値をyy年mm月の形に合わせる
    GetYyMm = Format(Left(txtTarget, lngYearLen), "00") & "年" _
              & Format(Mid(txtTarget, lngMonthLen), "00") & "月"
    Exit Function
              
GetYyMm_InputError:
    GetYyMm = Null
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :opg_BaggageDivision_AfterUpdate
'        機能             :荷入区分 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_BaggageDivision_AfterUpdate()
    
    Call AfterUpdateCommon(opg_BaggageDivision, 0, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :opg_DisposeDivision_AfterUpdate
'        機能             :処分区分 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_DisposeDivision_AfterUpdate()
    
    Call AfterUpdateCommon(opg_DisposeDivision, 0, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :opg_RepairDivision_AfterUpdate
'        機能             :補修区分 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_RepairDivision_AfterUpdate()
    
    Call AfterUpdateCommon(opg_RepairDivision, 0, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :opg_SpareKey_AfterUpdate
'        機能             :スペアキー AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_SpareKey_AfterUpdate()
    
    Call AfterUpdateCommon(opg_SpareKey, 0, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountNote_AfterUpdate
'        機能             :精算備考 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountNote_AfterUpdate()
    
    Call AfterUpdateCommon(txt_AccountNote, 0, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountNote_BeforeUpdate
'        機能             :精算備考 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountNote_BeforeUpdate(Cancel As Integer)

    If chkTextInput(txt_AccountNote, P_NOTE_LENGTH) <> 0 Then
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountNote_KeyPress
'        機能             :精算備考 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountNote_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char2(txt_AccountNote, KeyAscii, P_NOTE_LENGTH)

End Sub

'2006/09/19 SHIBAZAKI 予約名削除↓
'20051002 SHIBAZAKI 予約名追加
''==============================================================================*
''
''        MODULE_NAME      :txt_BookingName_AfterUpdate
''        機能             :予約名 AfterUpdate
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_BookingName_AfterUpdate()
'
'    Call AfterUpdateCommon(txt_BookingName, 0, 3)
'
'End Sub
'
''20051002 SHIBAZAKI 予約名追加
''==============================================================================*
''
''        MODULE_NAME      :txt_BookingName_BeforeUpdate
''        機能             :予約名 BeforeUpdate
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_BookingName_BeforeUpdate(Cancel As Integer)
'
'    If chkTextInput(txt_BookingName, P_BOOKING_NAME_LENGTH) <> 0 Then
'        Cancel = True
'    End If
'
'End Sub
'
''20051002 SHIBAZAKI 予約名追加
''==============================================================================*
''
''        MODULE_NAME      :txt_BookingName_KeyPress
''        機能             :予約名 KeyPress
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_BookingName_KeyPress(KeyAscii As Integer)
'
'    Call CmKeyPress_For_Char(txt_BookingName, KeyAscii, P_BOOKING_NAME_LENGTH)
'
'End Sub
'2006/09/19 SHIBAZAKI 予約名削除↑

'2006/09/19 SHIBAZAKI 予約番号削除↓
'20051002 SHIBAZAKI 予約番号追加
''==============================================================================*
''
''        MODULE_NAME      :txt_BookingNo_AfterUpdate
''        機能             :予約番号 AfterUpdate
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_BookingNo_AfterUpdate()
'
'    Call AfterUpdateCommon(txt_BookingNo, 0, 3)
'
'End Sub
'
''20051002 SHIBAZAKI 予約番号追加
''==============================================================================*
''
''        MODULE_NAME      :txt_BookingNo_BeforeUpdate
''        機能             :予約番号 BeforeUpdate
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_BookingNo_BeforeUpdate(Cancel As Integer)
'
'    If chkTextInput(txt_BookingNo, P_BOOKING_NO_LENGTH) <> 0 Then
'        Cancel = True
'    End If
'
'End Sub
'
''20051002 SHIBAZAKI 予約番号追加
''==============================================================================*
''
''        MODULE_NAME      :txt_BookingNo_KeyPress
''        機能             :予約番号 KeyPress
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_BookingNo_KeyPress(KeyAscii As Integer)
'
'    Call CmKeyPress_For_Char(txt_BookingNo, KeyAscii, P_BOOKING_NO_LENGTH)
'
'End Sub
'2006/09/19 SHIBAZAKI 予約番号削除↑

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelAcceptDate_AfterUpdate
'        機能             :解約受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelAcceptDate_AfterUpdate()

    txt_CancelAcceptDate = AfterUpdateCommon(txt_CancelAcceptDate, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelAcceptDate_BeforeUpdate
'        機能             :解約受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelAcceptDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_CancelAcceptDate, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelAcceptDate_GotFocus
'        機能             :解約受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelAcceptDate_GotFocus()

    '体裁を整える
    txt_CancelAcceptDate = Format(txt_CancelAcceptDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelAcceptDate_KeyPress
'        機能             :解約受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelAcceptDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_CancelAcceptDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelAcceptDate_LostFocus
'        機能             :解約受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelAcceptDate_LostFocus()

    '体裁を整える
    txt_CancelAcceptDate = Format(Me.txt_CancelAcceptDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDate_AfterUpdate
'        機能             :解約日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDate_AfterUpdate()
    
    txt_CancelDate = AfterUpdateCommon(txt_CancelDate, P_DATE_LENGTH, 2)
    
    If (cmb_ContractState = P_STATUS_CANCEL) Then
        txt_CancelExpectedDate = txt_CancelDate
    End If
    
    If (cmb_ContractState = P_STATUS_EXTEND And Nz(txt_CancelDate) <> "" And strExistingCancelDate <> "" And strLastCancelDate <> txt_CancelDate) Then
        txt_CancelExpectedDate = strExistingCancelDate
    End If
    
    strLastCancelDate = txt_CancelDate

    '2012/01/12 HONDA 解約予定／解約延長の場合には解約予定日の清算依頼日をセット
    If cmb_ContractState = P_STATUS_CANCEL Or cmb_ContractState = P_STATUS_EXTEND Then
        txt_AccountRequestDate = CDate(DateAdd("m", 1, Format(txt_CancelDate, "yyyy/mm") + "/05")) ' 清算依頼日
    End If
    '2012/01/12 HONDA 解約予定／解約延長の場合には解約予定日の清算依頼日をセット


End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDate_BeforeUpdate
'        機能             :解約日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    If dateCheck(txt_CancelDate, Cancel) = False Then
      Exit Sub
    End If

    '2006/09/19 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↓
    If txt_YardEndDay <> "" Then
        If DateDiff("d", txt_YardEndDay, VarChangeDate) > 0 Then
            MsgBox "解約日はヤード解約予定日 [" & txt_YardEndDay & "] 以前で無ければいけません。", vbInformation, P_KAKUNIN
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If
    '2006/09/19 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↑
    
    
    'ADD 2006/03/04 tajima キャンペーン適用満了日とのチェック
    If Nz(Me.txt_CampaignExpiration, "") <> "" Then
      If CDate(VarChangeDate) <= CDate(Me.txt_CampaignExpiration) Then
        Dim ret As Integer
        ret = MsgBox(pstrCampaignWarningText, vbYesNo + vbExclamation + vbDefaultButton2, "この契約で適用したキャンペーンについて")
        If ret = vbNo Then
          Cancel = True
          SendKeys "{ESC}"
          Exit Sub
        End If
      End If
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDate_GotFocus
'        機能             :解約日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDate_GotFocus()

    '体裁を整える
    txt_CancelDate = Format(txt_CancelDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDate_KeyPress
'        機能             :解約日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_CancelDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDate_LostFocus
'        機能             :解約日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDate_LostFocus()

    '体裁を整える
    txt_CancelDate = Format(Me.txt_CancelDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelExpectedDate_AfterUpdate
'        機能             :解約予定年月 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelExpectedDate_AfterUpdate()
    
    txt_CancelExpectedDate = AfterUpdateCommon(txt_CancelExpectedDate, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelExpectedDate_BeforeUpdate
'        機能             :解約予定年月 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelExpectedDate_BeforeUpdate(Cancel As Integer)

    '日付入力チェック
    Call dateCheck(txt_CancelExpectedDate, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelExpectedDate_GotFocus
'        機能             :解約予定年月 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelExpectedDate_GotFocus()

    '体裁を整える
    txt_CancelExpectedDate = Format(txt_CancelExpectedDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelExpectedDate_KeyPress
'        機能             :解約予定年月 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelExpectedDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_CancelExpectedDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelExpectedDate_LostFocus
'        機能             :解約予定年月 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelExpectedDate_LostFocus()

    '体裁を整える
    txt_CancelExpectedDate = Format(Me.txt_CancelExpectedDate, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_KeyReturnDate_AfterUpdate
'        機能             :鍵返却日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyReturnDate_AfterUpdate()
    
    txt_KeyReturnDate = AfterUpdateCommon(txt_KeyReturnDate, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeyReturnDate_BeforeUpdate
'        機能             :鍵返却日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyReturnDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    If dateCheck(txt_KeyReturnDate, Cancel) = False Then
      Exit Sub
    End If
    
    If Nz(VarChangeDate, "") <> "" And Nz(Me.txt_StartDate, "") <> "" Then
      If CDate(VarChangeDate) < CDate(Me.txt_StartDate) Then
        Dim ret As Integer
        ret = MsgBox("返却日が契約日以前です。宜しいのですか？", _
                      vbYesNo + vbExclamation + vbDefaultButton2, "入力矛盾警告")
        If ret = vbNo Then
          Cancel = True
          SendKeys "{ESC}"
          Exit Sub
        End If
      End If
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeyReturnDate_GotFocus
'        機能             :鍵返却日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyReturnDate_GotFocus()

    '体裁を整える
    txt_KeyReturnDate = Format(txt_KeyReturnDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeyReturnDate_KeyPress
'        機能             :鍵返却日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyReturnDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_KeyReturnDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeyReturnDate_LostFocus
'        機能             :鍵返却日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyReturnDate_LostFocus()

    '体裁を整える
    txt_KeyReturnDate = Format(Me.txt_KeyReturnDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceMemo_AfterUpdate
'        機能             :メンテメモ AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceMemo_AfterUpdate()
    
    Call AfterUpdateCommon(txt_MaintenanceMemo, 0, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceMemo_BeforeUpdate
'        機能             :メンテメモ BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceMemo_BeforeUpdate(Cancel As Integer)

    If chkTextInput(txt_MaintenanceMemo, P_MAINTENANCE_MEMO_LENGTH) <> 0 Then
        Cancel = True
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceMemo_KeyPress
'        機能             :メンテメモ KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceMemo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char2(txt_MaintenanceMemo, KeyAscii, P_MAINTENANCE_MEMO_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_AfterUpdate
'        機能             :メンテ依頼日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_AfterUpdate()
    
    txt_MaintenanceRequestDate = AfterUpdateCommon(txt_MaintenanceRequestDate, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_BeforeUpdate
'        機能             :メンテ依頼日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_MaintenanceRequestDate, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_GotFocus
'        機能             :メンテ依頼日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_GotFocus()

    '体裁を整える
    txt_MaintenanceRequestDate = Format(txt_MaintenanceRequestDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_KeyPress
'        機能             :メンテ依頼日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_MaintenanceRequestDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_LostFocus
'        機能             :メンテ依頼日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_LostFocus()

    '体裁を整える
    txt_MaintenanceRequestDate = Format(Me.txt_MaintenanceRequestDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_AfterUpdate
'        機能             :メンテ完了日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_AfterUpdate()
    
    txt_MaintenanceEndDate = AfterUpdateCommon(txt_MaintenanceEndDate, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_BeforeUpdate
'        機能             :メンテ完了日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_BeforeUpdate(Cancel As Integer)

    '日付入力チェック
    Call dateCheck(txt_MaintenanceEndDate, Cancel)
    '2006/05/31 ADD 井口 START↓
    If txt_MaintenanceEndDate < txt_MaintenanceRequestDate Then
        MsgBox "メンテ依頼日より前の日付は入力しないでください", vbInformation, P_KAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If
    '2006/05/31 ADD 井口   END↑


End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_GotFocus
'        機能             :メンテ完了日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_GotFocus()

    '体裁を整える
    txt_MaintenanceEndDate = Format(txt_MaintenanceEndDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_KeyPress
'        機能             :メンテ完了日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_MaintenanceEndDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_LostFocus
'        機能             :メンテ完了日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_LostFocus()

    '体裁を整える
    txt_MaintenanceEndDate = Format(Me.txt_MaintenanceEndDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_AfterUpdate
'        機能             :精算依頼日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_AfterUpdate()

    txt_AccountRequestDate = AfterUpdateCommon(txt_AccountRequestDate, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_BeforeUpdate
'        機能             :メンテ完了日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_AccountRequestDate, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_GotFocus
'        機能             :メンテ完了日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_GotFocus()

    '体裁を整える
    txt_AccountRequestDate = Format(txt_AccountRequestDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_KeyPress
'        機能             :メンテ完了日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_AccountRequestDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_LostFocus
'        機能             :メンテ完了日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_LostFocus()

    '体裁を整える
    txt_AccountRequestDate = Format(Me.txt_AccountRequestDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectStatus
'        機能             :ステータス変更時の画面制御
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectStatus()

    Select Case cmb_ContractState
    Case P_STATUS_CONTRACT
        Call SelectContract
    Case P_STATUS_CANCEL
        Call SelectCacnel
    Case P_STATUS_FUKA
        Call SelectCacnel
     
    Case P_STATUS_EXTEND
        Call SelectExtend
    Case P_STATUS_MAINTENANCE
        Call SelectMaintenance
    Case P_STATUS_COMPLETE
        If blnNewContract Then
            Call SelectCompleteNewContract
        Else
            Call SelectComplete
        End If
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectContract
'        機能             :ステータス=「契約」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectContract()

    txt_CancelAcceptDate = ""
    cmb_CancelAcceptPerson = ""
    txt_CancelAcceptPerson = ""

    cmb_CancelDivision = ""
    txt_CancelDivision = ""
    
    cmb_CancelReason = ""
    txt_CancelReason = ""
    
    txt_CancelDate = ""
    txt_CancelExpectedDate = ""
    
    'txt_KeyDivision = ""   2006/05/12 Del Tajima
    txt_KeyReturnDate = ""
    
    opg_SpareKey.VALUE = Null

    txt_MaintenanceRequestDate = ""
    txt_MaintenanceEndDate = ""
    cmb_MaintenancePerson = ""
    txt_MaintenancePerson = ""
    
    opg_BaggageDivision.VALUE = Null
    opg_DisposeDivision.VALUE = Null
    opg_RepairDivision.VALUE = Null

    txt_AccountRequestDate = ""
    
    Call YellowControl(txt_CancelDate, False)
    Call YellowControl(txt_CancelExpectedDate, False)
    
    Call EnableContract(False, True)
    cmb_ContractState.Enabled = True
    
    '2006/09/19 SHIBAZAKI 予約名/予約番号削除
    '20051004 SHIBAZAKI 項目追加
'    txt_BookingName.Enabled = True
'    txt_BookingNo.Enabled = True
    
    Call EnableMaintenance(False)
    Call EnableAccount(True)
    
    txt_MaintenanceMemo.Enabled = True
    txt_AccountRequestDate.Enabled = False
    
    '2006/05/13 tajima 項目追加
    txt_返却本数 = ""
'    txt_スペアキー備考.Enabled = True
    chk_鍵変更.VALUE = 0
    chk_鍵変更.Enabled = False
    chk_鍵変更_AfterUpdate
    
End Sub


'==============================================================================*
'
'        MODULE_NAME      :SelectCacnel
'        機能             :ステータス=「解約予定」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectCacnel()

    Call EnableContract(True)
    Call EnableMaintenance(False)
    Call EnableAccount(True)
    Call YellowControl(txt_CancelDate, False)
    Call YellowControl(txt_CancelExpectedDate)
    
    If intProcMode <> P_表示中 Then
        txt_CancelExpectedDate = txt_CancelDate
    End If
    
    txt_MaintenanceRequestDate = ""
    txt_MaintenanceEndDate = ""
    cmb_MaintenancePerson = ""
    txt_MaintenancePerson = ""
    opg_RepairDivision.VALUE = Null
    
    txt_MaintenanceMemo.Enabled = True
    opg_BaggageDivision.Enabled = True
    opg_DisposeDivision.Enabled = True
    
    chk_鍵変更.VALUE = 0
    chk_鍵変更.Enabled = False
    chk_鍵変更_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectExtend
'        機能             :ステータス=「解約延長」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectExtend()
    
    Call SelectCacnel

    If intProcMode <> P_表示中 Then
        '状態区分を「延長」に変更したら、解約予定日に既登録データの解約日を設定する
        If cmb_ContractState = P_STATUS_EXTEND And strLastStatus <> P_STATUS_EXTEND Then
            txt_CancelExpectedDate = Format(strExistingCancelDate, "yyyy/mm/dd")
        End If
    End If
    
    '解約日/解約予定日白化
    Call YellowControl(txt_CancelDate, False)
    Call YellowControl(txt_CancelExpectedDate, False)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectMaintenance
'        機能             :ステータス=「メンテ中」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectMaintenance()

    Call SelectComplete

    txt_MaintenanceEndDate = ""
    txt_MaintenanceEndDate.Enabled = False
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectComplete
'        機能             :ステータス=「完了」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectComplete()

    Call EnableContract(True)
    Call EnableMaintenance(True)
    Call EnableAccount(True)
    
    Call YellowControl(txt_CancelDate)
    Call YellowControl(txt_CancelExpectedDate)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectCompleteNewContract
'        機能             :ステータス=「完了」＆表示データより新しい契約がある
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectCompleteNewContract()

    cmd_End.SetFocus
    Call EnableContract(False)
    Call EnableMaintenance(False)
    Call EnableAccount(True)
    Call ReferenceMode(False)
    
    '2006/09/19 SHIBAZAKI 予約名/予約番号削除
    '20051004 SHIBAZAKI 項目追加
'    txt_BookingName.Enabled = True
'    txt_BookingNo.Enabled = True
'    Call YellowControl(txt_BookingName, False)
'    Call YellowControl(txt_BookingNo, False)
    
'    txt_BookingName.SetFocus
    cmd_下画面.SetFocus
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :AfterUpdateCommon
'        機能             :更新後イベント時の処理
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'                         :③1=数値項目 2=日付項目 3=テキスト項目
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function AfterUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer) As String

    AfterUpdateCommon = ""
    
    '体裁を整える
    Select Case intKind
    Case 1  '数値項目は０埋め
        objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    Case 2
        AfterUpdateCommon = Format(VarChangeDate, "yyyy/mm/dd")
    Case Else
        AfterUpdateCommon = Nz(objTarget)
    End Select
    
    '変更された場合（初回だけ）タイトルにその旨を表示する
    If Not blnChangeMode Then
        blnChangeMode = True
        Me.Caption = Me.Caption & P_EDIT_MODE
    End If
    
    'コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        AfterUpdateCommon = Nz(objTarget.Column(1, objTarget.ListIndex))
    End If

End Function


'==============================================================================*
'
'        MODULE_NAME      :文字項目のチェックを一括でする。
'        MODULE_ID        :CmBCheck_Char
'        CREATE_DATE      :
'
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

    On Error GoTo err_rtn

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
       CmBCheck_Char = 2
       GoTo Exit_rtn
    End If

    CmBCheck_Char = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmBCheck_Char = 9
    GoTo Exit_rtn
End Function


'==============================================================================*
'
'        MODULE_NAME      :GetCampaignTexts
'        機能             :キャンペーン文言の取得
'        IN               :rsCarg(検索結果レコードセット)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub GetCampaignTexts(rsCarg As Recordset)

    pstrCampaignInfomationText = ""
    pstrCampaignWarningText = ""

    Dim strキャンペ適満日 As String
    
    strキャンペ適満日 = Format(rsCarg.Fields("CARG_CMP_EXDATE"), "yyyy/mm/dd")
    
    'キャンペーン適用満了日に値が入っていればキャンペーン文言を生成する
    If Nz(strキャンペ適満日, "") <> "" Then
        pstrCampaignInfomationText = CmKomsMod.GetCampaignInfomationText( _
                                              Nz(rsCarg.Fields("RCPT_SEV1N"), ""), _
                                              Nz(rsCarg.Fields("RCPT_SEV2N"), ""), _
                                              Nz(rsCarg.Fields("RCPT_SEV3N"), ""), _
                                              Nz(rsCarg.Fields("RCPT_ENDEN"), ""), _
                                              strキャンペ適満日 _
                                              )
        pstrCampaignWarningText = CmKomsMod.GetCampaignWarningText( _
                                              Nz(rsCarg.Fields("RCPT_SEV1N"), ""), _
                                              Nz(rsCarg.Fields("RCPT_SEV2N"), ""), _
                                              Nz(rsCarg.Fields("RCPT_SEV3N"), ""), _
                                              Nz(rsCarg.Fields("RCPT_ENDEN"), ""), _
                                              strキャンペ適満日 _
                                              )
    End If

End Sub
'*********************** これ以降は 0.5 2006/05/13 Add *************************
'==============================================================================*
'
'        MODULE_NAME      :EnableKeyMaintenance
'        機能             :コントロール使用可不可(鍵メンテ)
'        IN               :True=使用可 False=使用不可
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub EnableKeyMaintenance(blnEnable As Boolean)

'↓ UPDATE 2014/04/11 MIYAMOTO
'    cmb_NewKeyDivision.Enabled = blnEnable
    If Nz(txt_自動鍵区分) = "1" Then
        cmb_NewKeyDivision.Enabled = False
    Else
        cmb_NewKeyDivision.Enabled = blnEnable
    End If
'↑ UPDATE 2014/04/11 MIYAMOTO
    txt_NewKeyNo.Enabled = blnEnable
    txt_新貸出本数.Enabled = blnEnable
'    txt_新解除番号.Enabled = blnEnable          'INSERT 2014/04/11 MIYAMOTO → DELETE 2016/05/20 MIYAMOTO
    txt_鍵変更日.Enabled = blnEnable
    cmb_鍵変更担当.Enabled = blnEnable
    cmb_鍵変更理由.Enabled = blnEnable
    txt_鍵変更備考.Enabled = blnEnable

End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_鍵変更_AfterUpdate
'        機能             :鍵変更 AfterUpdate ｲﾍﾞﾝﾄ処理
'        NOTE             :変更情報入力可否の制御を行う
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_鍵変更_AfterUpdate()

    If chk_鍵変更.VALUE = 0 Then
    
      EnableKeyMaintenance (False)
      
    ' 無効なら表示項目は灰モード
      txt_NewKeyDivision.BackColor = P_BACKCOLOR_GLAY
      txt_鍵変更担当.BackColor = P_BACKCOLOR_GLAY
      txt_鍵変更理由.BackColor = P_BACKCOLOR_GLAY
      txt_新解除番号.BackColor = P_BACKCOLOR_GLAY           'INSERT 2014/04/11 MIYAMOTO
        
    Else
      EnableKeyMaintenance (True)
    ' 有効リンなら表示項目は黄モード
      txt_NewKeyDivision.BackColor = P_BACKCOLOR_YELLOW
      txt_鍵変更担当.BackColor = P_BACKCOLOR_YELLOW
      txt_鍵変更理由.BackColor = P_BACKCOLOR_YELLOW
      txt_新解除番号.BackColor = P_BACKCOLOR_YELLOW         'INSERT 2014/04/11 MIYAMOTO
    
    ' 入力項目に初期値を与える
      If Nz(cmb_NewKeyDivision, "") = "" Then
        cmb_NewKeyDivision = pstr元鍵区分
        Call cmb_NewKeyDivision_AfterUpdate
      End If
      
      If Nz(txt_NewKeyNo, "") = "" Then
        txt_NewKeyNo = txt_鍵番号
      End If
      
      If Nz(txt_新貸出本数, "") = "" Then
        txt_新貸出本数 = txt_貸出本数2
      End If
      
'↓ INSERT 2014/04/11 MIYAMOTO
      If Nz(txt_新解除番号, "") = "" Then
        txt_新解除番号 = txt_解除番号
      End If
'↑ INSERT 2014/04/11 MIYAMOTO
      
      If Nz(txt_鍵変更日, "") = "" Then
        txt_鍵変更日 = Format(Now, "yyyy/mm/dd")
      End If
    
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_NewKeyNo_AfterUpdate
'        機能             :新鍵番号 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_NewKeyNo_AfterUpdate()
    
    Call AfterUpdateCommon(txt_NewKeyNo, 0, 3)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_NewKeyNo_BeforeUpdate
'        機能             :新鍵番号 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_NewKeyNo_BeforeUpdate(Cancel As Integer)
    
    If chkTextInput(txt_NewKeyNo, 10) <> 0 Then
        Cancel = True
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_NewKeyNo_KeyPress
'        機能             :新鍵番号 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_NewKeyNo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char2(txt_NewKeyNo, KeyAscii, 10)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更日_AfterUpdate
'        機能             :鍵変更日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更日_AfterUpdate()

    txt_鍵変更日 = AfterUpdateCommon(txt_鍵変更日, P_DATE_LENGTH, 2)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更日_BeforeUpdate
'        機能             :鍵変更日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更日_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_鍵変更日, Cancel)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更日_GotFocus
'        機能             :鍵変更 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更日_GotFocus()

    '体裁を整える
    txt_鍵変更日 = Format(txt_鍵変更日, "yy/mm/dd")
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更日_KeyPress
'        機能             :鍵変更日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更日_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(txt_鍵変更日, KeyAscii, P_DATE_LENGTH)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更日_LostFocus
'        機能             :鍵変更日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更日_LostFocus()

    '体裁を整える
    txt_鍵変更日 = Format(txt_鍵変更日, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更備考_AfterUpdate
'        機能             :鍵変更備考 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更備考_AfterUpdate()
    
    Call AfterUpdateCommon(txt_鍵変更備考, 0, 3)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更備考_BeforeUpdate
'        機能             :鍵変更備考 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更備考_BeforeUpdate(Cancel As Integer)

    If chkTextInput(txt_鍵変更備考, 60) <> 0 Then
        Cancel = True
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_鍵変更備考_KeyPress
'        機能             :鍵変更備考 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_鍵変更備考_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char2(txt_鍵変更備考, KeyAscii, 60)

End Sub

'↓ INSERT 2016/05/20 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :txt_新解除番号_AfterUpdate
'        機能             :新解除番号 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_新解除番号_AfterUpdate()

    If Len(txt_新解除番号) > 0 Then
        If Not (LenB(StrConv(txt_新解除番号, vbFromUnicode)) = 4 And IsNumeric(txt_新解除番号)) = True Then
            MsgBox "半角数字4桁以内で入力してください。", vbExclamation
            txt_新解除番号 = ""
        End If
    End If

End Sub
'↑ INSERT 2016/05/20 MIYAMOTO

'↓ INSERT 2016/05/20 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :txt_新解除番号_GotFocus
'        機能             :新解除番号 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_新解除番号_GotFocus()

    '「変更する」がチェックされていて、鍵区分が変更可能の状態の時、
    '鍵区分が手動解除番号入力対象の場合、入力可能にする
    If chk_鍵変更.VALUE <> 0 Then
        If MSZZ070_KaiKey(txt_YardTihoKbn, cmb_NewKeyDivision.Column(2)) = "2" Then
            txt_新解除番号.Locked = False
            txt_新解除番号.InputMask = ""
        End If
    End If

End Sub
'↑ INSERT 2016/05/20 MIYAMOTO

'↓ INSERT 2016/05/20 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :txt_新解除番号_LostFocus
'        機能             :新解除番号 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_新解除番号_LostFocus()

    '「変更する」がチェックされていて、鍵区分が変更可能の状態の時、
    '鍵区分が手動解除番号入力対象の場合、入力可能にする
    If chk_鍵変更.VALUE <> 0 Then
        If MSZZ070_KaiKey(txt_YardTihoKbn, cmb_NewKeyDivision.Column(2)) = "2" Then
            txt_新解除番号.Locked = True
            If pvONLET_ACCEI = 0 Then   'INSERT 2021/06/02 EGL
                txt_新解除番号.InputMask = "Password"
            End If
        End If
    End If

End Sub
'↑ INSERT 2016/05/20 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :txt_新貸出本数_AfterUpdate
'        機能             :新貸出本数 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_新貸出本数_AfterUpdate()

    Call AfterUpdateCommon(txt_新貸出本数, 0, 3)
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_新貸出本数_BeforeUpdate
'        機能             :新貸出本数 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_新貸出本数_BeforeUpdate(Cancel As Integer)

    If chkTextInput(txt_新貸出本数, 2) <> 0 Then
        Cancel = True
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_新貸出本数_KeyPress
'        機能             :新貸出本数 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_新貸出本数_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txt_新貸出本数, KeyAscii, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_返却本数_AfterUpdate
'        機能             :返却本数 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却本数_AfterUpdate()

    Call AfterUpdateCommon(txt_返却本数, 0, 3)
        
    If Nz(txt_返却本数, "") <> "" Then
      txt_返却本数 = CInt(txt_返却本数) 'ゼロサプレス解除
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_返却本数_BeforeUpdate
'        機能             :新貸出本数 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却本数_BeforeUpdate(Cancel As Integer)

    If chkTextInput(txt_返却本数, 2) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    
    '貸出本数との矛盾チェック
    If Nz(Me.txt_返却本数, "") <> "" Then
      If CInt(Nz(Me.txt_貸出本数, 0)) < CInt(txt_返却本数) Then
        Dim ret As Integer
        ret = MsgBox("貸出した本数以上です。宜しいのですか？", _
                      vbYesNo + vbExclamation + vbDefaultButton2, "入力矛盾警告")
        If ret = vbNo Then
          Cancel = True
          SendKeys "{ESC}"
          Exit Sub
        End If
      End If
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_返却本数_KeyPress
'        機能             :返却本数 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却本数_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Num1(txt_返却本数, KeyAscii, 2)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_スペアキー備考_AfterUpdate
'        機能             :スペアキー備考 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_スペアキー備考_AfterUpdate()

    Call AfterUpdateCommon(txt_スペアキー備考, 0, 3)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_スペアキー備考_BeforeUpdate
'        機能             :スペアキー備考 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_スペアキー備考_BeforeUpdate(Cancel As Integer)
    
    If chkTextInput(txt_スペアキー備考, 60) <> 0 Then
        Cancel = True
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_スペアキー備考_KeyPress
'        機能             :スペアキー備考 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_スペアキー備考_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char2(txt_スペアキー備考, KeyAscii, 60)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :MakeFVS700_W01
'        機能             :ワークテーブル(FVS700_W01)作成
'        IN               :
'        OUT              :
'        NOTE             :サブフォーム(FVS700S01)とリンクするテーブル
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeFVS700_W01(dbSQLServer As DAO.Database, _
                                ByVal aヤードコード As Long, _
                                ByVal a契約番号 As String) As Boolean
    
    Dim rsSqlServer As Recordset
    Dim rsAccess    As Recordset
    Dim isException As Boolean

    isException = False
    MakeFVS700_W01 = False
    On Error GoTo Exception

    ' Access接続 ※Privateグローバルのオブジェクト
    Set pdbAccess = CurrentDb

    'ワークテーブルクリア
    If Not ClearFVS700_W01(P_WORK_TABLE) Then
        GoTo Exception
    End If

    'ワークテーブルを開く
    Set rsAccess = pdbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)

    'ワークテーブルの元ネタ取得
    ' 現在この契約で使用中のカードを取得するSQL生成
    Dim strSQL As String
    strSQL = "SELECT "
    strSQL = strSQL & "   SCRDM_CARD_NO, SCRDM_USAGE"
    strSQL = strSQL & "   ,USAGE_NAME.NAME_NAME AS USAGE_NAME"
    strSQL = strSQL & "   ,SCRDM_LENDING_DAY"
    strSQL = strSQL & "   ,SCRDM_STOP_DAY"
    strSQL = strSQL & "   ,SCRDM_NOTE"
    strSQL = strSQL & " FROM SCRD_MAST"
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST AS USAGE_NAME"
    strSQL = strSQL & "   ON USAGE_NAME.NAME_CODE = SCRD_MAST.SCRDM_USAGE"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "   SCRDM_YCODE =" & aヤードコード
    strSQL = strSQL & "   AND SCRDM_ACPTNO ='" & a契約番号 & "'"
    strSQL = strSQL & "   AND USAGE_NAME.NAME_ID = '099'"
    strSQL = strSQL & " ORDER BY SCRDM_CARD_NO"
    ' 現在この契約で使用中のカードを取得するSQL生成
    Set rsSqlServer = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    'ワークテーブルに元ネタを元にレコードを追記していく
    Do While Not rsSqlServer.EOF
        
        rsAccess.AddNew
        ' キーとなるカード番号
        rsAccess.Fields("CARD_NO") = rsSqlServer.Fields("SCRDM_CARD_NO")
        ' カード番号/用途名
        rsAccess.Fields("CARD_NAME") = rsSqlServer.Fields("SCRDM_CARD_NO") & "/" & _
                                        Nz(rsSqlServer.Fields("USAGE_NAME"), "")
        ' 貸出日(mm/dd)
        rsAccess.Fields("LENDING_DAY") = rsSqlServer.Fields("SCRDM_LENDING_DAY")
        ' 備考
        rsAccess.Fields("NOTE") = rsSqlServer.Fields("SCRDM_NOTE")
        ' 停止日が入っていれば チェック済み、なければ未チェックとする
        If Nz(rsSqlServer.Fields("SCRDM_STOP_DAY"), "") <> "" Then
            rsAccess.Fields("STOP_CHK") = -1
        Else
            rsAccess.Fields("STOP_CHK") = 0
        End If
        ' 返却日は未チェックで ※返却済みはこのフォームには出てこない
        rsAccess.Fields("RETURN_CHK") = 0
        rsAccess.UPDATE
        
        rsSqlServer.MoveNext
    Loop
    
    If rsSqlServer.RecordCount = 0 Then
        lbl_カード貸出枚数.Caption = "貸出無し"
    Else
        lbl_カード貸出枚数.Caption = rsSqlServer.RecordCount & "枚貸出"
    End If
    
    MakeFVS700_W01 = True
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not pdbAccess Is Nothing Then pdbAccess.Close: Set pdbAccess = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "MakeFVS700_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :ClearFVS700_W01
'        機能             :ワークテーブルクリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ClearFVS700_W01(strTable As String) As Boolean

    Dim tdfNew      As TableDef
    Dim isException As Boolean
    Dim isConnect   As Boolean
    
    On Error GoTo Exception
    
    ClearFVS700_W01 = False
    isException = False
    isConnect = False
    
    If pdbAccess Is Nothing Then
        Set pdbAccess = CurrentDb
        isConnect = True
    End If
    
    SubForm.SourceObject = ""
    
    'ワークテーブル削除
    If TableExist(pdbAccess, strTable) Then
        Call pdbAccess.TableDefs.Delete(strTable)
    End If

    'ワークテーブル作成
    Set tdfNew = pdbAccess.CreateTableDef(strTable)
    Call FieldAppendFVS700_W01(tdfNew)
    Call pdbAccess.TableDefs.Append(tdfNew)
    
    lbl_カード貸出枚数.Caption = "貸出無し"
    
    ClearFVS700_W01 = True
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isConnect = True Then
        If Not pdbAccess Is Nothing Then pdbAccess.Close: Set pdbAccess = Nothing
    End If
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    
    SubForm.SourceObject = P_SUB_FORM
    
    If isException = True Then
        Call Err.Raise(Err.Number, "ClearFVS700_W01" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :FieldAppendFVS700_W01
'        機能             :ワークテーブル列作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FieldAppendFVS700_W01(tdfNew As TableDef)

    With tdfNew
 
        Call .Fields.Append(.CreateField("CARD_NO", DataTypeEnum.dbText, 11))
        Call .Fields.Append(.CreateField("CARD_NAME", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("LENDING_DAY", DataTypeEnum.dbDate))
        Call .Fields.Append(.CreateField("RETURN_CHK", DataTypeEnum.dbInteger))
        Call .Fields.Append(.CreateField("STOP_CHK", DataTypeEnum.dbInteger))
        Call .Fields.Append(.CreateField("NOTE", DataTypeEnum.dbText, 60))
        
    End With
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :TableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :テーブル名
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function TableExist(dbAccess As Database, strTableName As String) As Boolean

    Dim tdf As TableDef
    
    TableExist = False
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            TableExist = True
            Exit For
        End If
    Next tdf

End Function
'==============================================================================*
'
'        MODULE_NAME      :SetSecurityCardMast
'        機能             :セキュリティカードマスタに対する更新
'        IN               :
'        OUT              :
'        NOTE             :ワークテーブル(FVS700_W01)を元に更新を行う
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SetSecurityCardMast() As Boolean

    Dim strSQL        As String
    Dim isException   As Boolean
    Dim cardRecord    As Recordset '書込先のカードマスタのレコードセット
    Dim workRecord    As Recordset '読込先のワークテーブルのレコードセット
    Dim isChange      As Boolean   '書込先への変更可否フラグ
    Dim strUserName   As String    '更新ユーザID
    Dim strDate       As String    '更新日付
    Dim strTime       As String    '更新時刻
    
    On Error GoTo Exception
    
    SetSecurityCardMast = False
    
    ' 更新情報を取得しておく
    strUserName = LsGetUserName()
    strDate = Format$(DATE, "yyyymmdd")
    strTime = Format$(time, "hhmmss")
     ' Access接続 ※Privateグローバルのオブジェクト
    Set pdbAccess = CurrentDb
    
    ' ワークテーブルの内容をカードマスタに反映させる
    strSQL = " SELECT * FROM " & P_WORK_TABLE
    Set workRecord = pdbAccess.OpenRecordset(strSQL, dbReadOnly)
    
    With workRecord
        Do Until .EOF
            '更新対象のカードを開く
            strSQL = " SELECT * FROM SCRD_MAST"
            strSQL = strSQL & " WHERE SCRDM_YCODE = " & Val(cmb_Yard.VALUE)
            strSQL = strSQL & "   AND SCRDM_CARD_NO = '" & .Fields("CARD_NO") & "'"
            Set cardRecord = objDb.OpenRecordset(strSQL, dbOpenDynaset)
            
            If Not cardRecord.EOF Then
                 
                 isChange = False
                 cardRecord.Edit
                 
                 ' 備考の設定
                 If Nz(cardRecord.Fields("SCRDM_NOTE"), "") <> Nz(.Fields("NOTE"), "") Then
                     cardRecord.Fields("SCRDM_NOTE") = .Fields("NOTE")
                     isChange = True
                 End If
                 
                 ' 返却日カードの設定
                 If .Fields("RETURN_CHK") = -1 Then
                    Call CmKomsMod.InitializSCards(cardRecord)   '返却カード手順は共通化
                    isChange = True
                 End If
                 
                 ' 停止日の設定
                 If .Fields("STOP_CHK") = -1 Then
                     ' 停止チェックに変えたならば停止日を設定
                     If Nz(cardRecord.Fields("SCRDM_STOP_DAY"), "") = "" Then
                         cardRecord.Fields("SCRDM_STOP_DAY") = DATE
                         isChange = True
                     End If
                 Else
                     ' 停止チェックを外して変えたならば停止日をクリア
                     If Nz(cardRecord.Fields("SCRDM_STOP_DAY"), "") <> "" Then
                         cardRecord.Fields("SCRDM_STOP_DAY") = Null
                         isChange = True
                     End If
                 End If
          
                 '書込先に変更がおきていれば更新記録を付けてアップデート実施
                 If isChange = True Then
                     cardRecord.Fields("SCRDM_UPDAD") = strDate
                     cardRecord.Fields("SCRDM_UPDAJ") = strTime
                     cardRecord.Fields("SCRDM_UPDPB") = Me.NAME
                     cardRecord.Fields("SCRDM_UPDUB") = strUserName
                     cardRecord.UPDATE
                 End If
            
            End If  'If Not cardRecord.EOF Then
    
            ' ワークテーブルの次レコード呼出
            .MoveNext
        Loop
    End With
    SetSecurityCardMast = True
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not cardRecord Is Nothing Then cardRecord.Close: Set cardRecord = Nothing
    If Not workRecord Is Nothing Then workRecord.Close: Set workRecord = Nothing
    If Not pdbAccess Is Nothing Then pdbAccess.Close: Set pdbAccess = Nothing
    
    If isException = True Then
        Call Err.Raise(Err.Number, "SetSecurityCardMast" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :ChangeSubForm
'        機能             :サブフォームが呼ぶ変更通知
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub ChangeSubForm()

    '変更された場合（初回だけ）タイトルにその旨を表示する
    If Not blnChangeMode Then
        blnChangeMode = True
        Me.Caption = Me.Caption & P_EDIT_MODE
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :SubReloadfromFVS90n
'        機能             :FVS90n系が呼ぶ、修繕関係の再読込処理
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SubReloadfromFVS90n()
    Call DispReprInfo
End Sub
'****************************  ended of program ********************************


'==============================================================================*
'
'        MODULE_NAME      : CallRoomReleaseAPI
'        機能             : 部屋解除API（顔認証データ等の削除）を呼び出す
'        IN               : なし（画面のcmb_Yard, cmb_CntaNoを利用）
'        OUT              : True=成功(204,404), False=失敗
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CallRoomReleaseAPI() As Boolean
    On Error GoTo Exception

    Dim objRequest As Object
    Dim strUrl     As String
    Dim strApiKey  As String
    Dim strJson    As String
    Dim intStatus  As Integer
    Dim StrYard    As String
    Dim strRoom    As String

    CallRoomReleaseAPI = False

    ' --------------------------------------------------------
    ' テスト環境用
    ' --------------------------------------------------------
    strUrl = "https://dev-facekey-admin.kase3535.com//api/v1/external/system/rooms"
    strApiKey = "ff246566d3981d6d571f0011cbc1668f15172d6e73d6a65b827e681b6c6de492"

    ' 画面の値をそのまま取得
    StrYard = Nz(Me.cmb_Yard.VALUE, "")
    strRoom = Nz(Me.cmb_CntaNo.VALUE, "")

    ' JSONボディの作成
    strJson = "{""yard_code"":""" & StrYard & """, ""room_code"":""" & strRoom & """}"

    ' XMLHTTPオブジェクトの作成 (遅延バインディング)
    Set objRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' APIリクエストの設定
    objRequest.Open "DELETE", strUrl, False
    objRequest.setRequestHeader "Content-Type", "application/json"
    objRequest.setRequestHeader "X-API-Key", strApiKey

    ' リクエスト送信
    objRequest.send strJson

    ' ステータスコードの判定
    intStatus = objRequest.Status
    Select Case intStatus
        Case 204
            CallRoomReleaseAPI = True
        Case 404
            ' 顔認証未導入の部屋などの場合は404が返るため、エラーとせず無視する
            CallRoomReleaseAPI = True
        Case 401
            MsgBox "【API連携エラー】" & vbCrLf & "APIキーが未指定、または無効です。(401)", vbExclamation, P_ERROR
        Case 403
            MsgBox "【API連携エラー】" & vbCrLf & "IP制限によりアクセスが拒否されました。(403)", vbExclamation, P_ERROR
        Case 422
            MsgBox "【API連携エラー】" & vbCrLf & "リクエストの形式が正しくありません。(422)", vbExclamation, P_ERROR
        Case 500
            MsgBox "【API連携エラー】" & vbCrLf & "サーバー内部エラーが発生しました。(500)", vbExclamation, P_ERROR
        Case Else
            MsgBox "【API連携エラー】" & vbCrLf & "予期せぬエラーが発生しました。(HTTP " & intStatus & ")", vbExclamation, P_ERROR
    End Select

    GoTo Finally

Exception:
    MsgBox "APIの呼び出し中にシステムエラーが発生しました。" & vbCrLf & _
           "ｴﾗｰ番号: " & Err.Number & vbCrLf & Err.Description, vbCritical, P_ERROR

Finally:
    If Not objRequest Is Nothing Then Set objRequest = Nothing
End Function

