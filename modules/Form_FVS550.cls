VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS550"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　取置き履歴
'   プログラムＩＤ　：　FVS550
'   作　成　日　　　：  2005/12/20
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/1/24
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.1
'   修　正　内　容　：  ヘッダに移動理由追加
'
'   修　正　日　　　：  2006/03/09
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.2
'   修　正　内　容　：  1.受付時に受付入力画面に遷移する
'                  :   2.再取置き時に受付データを削除する
'
'   修　正　日　　　：  2006/09/11
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.3
'   修　正　内　容　：  1.受付としたときに希望開始日に受付日＋サービス期間を設定する
'                  :   2.受付日＋サービス期間が営業期間外ならばエラーにする
'                  :   3.サービス期間をコントロールマスタから取得
'
'   修　正　日　　　：  2007/04/16
'   修　正　者　　　：  tajima
'   Ver             ：  0.4
'   修　正　内　容　：  近隣ヤード紹介対応
'                   ：  ①希望ヤードを表示する※暫定
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
Private dbSQLServer     As Database
Private dbAccess        As Database

'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'部門コード
Private strBumonCode    As String

'呼び出し元画面ID
Public strFormFrom     As String

'データを更新したらTrueになる
Private blnUpdate       As Boolean

'2006/03/09 SHIBAZAKI 追加
Public blnReloadSkip    As Boolean  'Trueならば終了時に呼び出しもとのリフレッシュは行わない

'2006/09/11 SHIBAZAKI サービス期間をコントロールマスタから取得
Private intServiceDay   As Integer

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

Private Const P_WORK_TABLE          As String = "FVS550_W01"

'取置き一覧サブフォームレコードソース
Private Const P_SUBFORM_RECORD_SOURCE   As String = _
                            " SELECT * FROM " & P_WORK_TABLE & _
                            "  ORDER BY INTRONO DESC; "
'予約受付状態区分
Private Const P_BOOKING_STATUS_AUTO_INTRO       As Integer = 0  '予約受付状態区分「自動紹介」
Private Const P_BOOKING_STATUS_JISYA            As Integer = 1  '予約受付状態区分「自社理由」
Private Const P_BOOKING_STATUS_KEEP             As Integer = 2  '予約受付状態区分「予約取置」
Private Const P_BOOKING_STATUS_ACCEPT           As Integer = 10 '予約受付状態区分「受付」
Private Const P_BOOKING_STATUS_CONTRACT         As Integer = 20 '予約受付状態区分「契約」
Private Const P_BOOKING_STATUS_BOOKING_CANCEL   As Integer = 51 '予約受付状態区分「予約キャンセル」
Private Const P_BOOKING_STATUS_KEEP_CANCEL      As Integer = 52 '予約受付状態区分「取置キャンセル」
Private Const P_BOOKING_STATUS_ACCEPT_CANCEL    As Integer = 53 '予約受付状態区分「受付キャンセル」

'ご紹介区分
Private Const P_INTRO_STATUS_KEEP           As String = "1" 'ご紹介区分「取り置きした」
Private Const P_INTRO_STATUS_ACCEPT         As String = "2" 'ご紹介区分「受付した」
Private Const P_INTRO_STATUS_CANCEL         As String = "6" 'ご紹介区分「キャンセルした」
Private Const P_INTRO_STATUS_NOT_INTRO      As String = "9" 'ご紹介区分「再度ご紹介不要」

Private Const P_パラメータ数                        As Integer = 2
'前画面からのパラメータのインデックス 上から順にカンマ区切りで渡してもらう
Private Const P_パラメータ_呼び出し元画面            As Integer = 0
Private Const P_パラメータ_予約受付番号              As Integer = 1

'再取置き可能な場合に表示する文字列
Private Const P_可能文字列                  As String = "可能"

'印刷済みの場合に表示する文字列
Private Const P_印刷済み文字列              As String = "済"

'コンテナ利用可否
Private Const P_CNTA_USE_OK                 As Integer = 1  '「可」
Private Const P_CNTA_USE_NO                 As Integer = 0  '「否」

'自動区分
Private Const P_AUTOKBN_無し                As Integer = 0

'コントロールマスタのキー
Private Const P_CONT_KEY                    As Integer = 1

'保存時確認メッセージ
Private Const P_SAVE_CONFIRM_再取置き       As String = "再取置きします。"
Private Const P_SAVE_CONFIRM_受付           As String = "受付します。"
Private Const P_SAVE_CONFIRM_紹介不要       As String = "再度ご紹介不要にします。"

'06.1.24 SHIBAZAKI 追加
Private Const P_MOVEKBN_無し        As Integer = 0

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        機能             :終了ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next
    '終了メッセージを表示
    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton1, P_KAKUNIN) = vbYes Then
        doCmd.Close acForm, Me.FormName
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_照会可能印刷_Click
'        機能             :照会可能印刷ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_照会可能印刷_Click()
    
    If pfncRptPrintFVS530(pintCstOther, pintCstPreview, pintCstViewOnly, , SubHeader.Form.txt_Ukno) Then
        Call AllReload
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_新規印刷_Click
'        機能             :新規印刷ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_新規印刷_Click()
    
    If pfncRptPrintFVS530(pintCstOther, pintCstPreview, pintCstViewNew, , SubHeader.Form.txt_Ukno) Then
        Call AllReload
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_全て印刷_Click
'        機能             :全て印刷ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_全て印刷_Click()
    
    If pfncRptPrintFVS530(pintCstOther, pintCstPreview, pintCstViewAll, , SubHeader.Form.txt_Ukno) Then
        Call AllReload
    End If

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
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'F1=「終了」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

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
    
    Dim blnRet      As Boolean
    Dim strTemp     As String
    Dim strParam()  As String
    
    blnUpdate = False
    
    '2006/03/09 SHIBAZAKI 追加
    blnReloadSkip = False
    
    'パラメータ取得
    strTemp = Nz(Me.OpenArgs)
    
    strParam = Split(strTemp, ",")
    If UBound(strParam) <> P_パラメータ数 - 1 Then
        MsgBox "パラメータが正しくありません"
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    
    strFormFrom = strParam(P_パラメータ_呼び出し元画面)
    SubHeader.Form.txt_Ukno = strParam(P_パラメータ_予約受付番号)
    
    'SQL-Server接続準備
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Exit Sub
    End If
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumonCode & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If
    '接続文字列取得
    strConnect = MSZZ007_M00(strBumonCode)
    If strConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If
    
    '2006/09/11 SHIBAZAKI サービス期間をコントロールマスタから取得
    intServiceDay = Nz(DLookup("CONT_SERVICE_DD", "dbo_CONT_MAST", "CONT_KEY=1"), 0)

    '画面更新
    If Not AllReload() Then
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :AllReload
'        機能             :画面表示更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function AllReload() As Boolean

    AllReload = False
    
    'ヘッダ情報取得
    If Not DispHeader() Then
        Exit Function
    End If
    
    'サブフォームのボタン使用可不可設定
    Select Case Val(SubHeader.Form.txt_Yukbn)
    Case P_BOOKING_STATUS_AUTO_INTRO, P_BOOKING_STATUS_KEEP
        '「自動紹介」「取置き」＝すべて使用可
        Call SubForm.Form.SubFormButtonEnable(True, True, True) '左から、再取、受付、不要
        SubForm.SetFocus
    Case P_BOOKING_STATUS_JISYA
        '「自社理由」＝受付ボタンのみ使用可
        cmd_終了.SetFocus
        Call Me.SubForm.Form.SubFormButtonEnable(False, True, False)
        SubForm.SetFocus
    Case P_BOOKING_STATUS_BOOKING_CANCEL, P_BOOKING_STATUS_KEEP_CANCEL, P_BOOKING_STATUS_ACCEPT_CANCEL
        '「予約キャンセル」「取置きキャンセル」「受付キャンセル」
        '　　＝再取置きボタンのみ使用可
        cmd_終了.SetFocus
        Call Me.SubForm.Form.SubFormButtonEnable(True, False, False)
        SubForm.SetFocus
    Case Else
        '「受付」「契約」
        '　　＝すべて使用不可
        cmd_終了.SetFocus
        Call Me.SubForm.Form.SubFormButtonEnable(False, False, False)
        SubForm.SetFocus
    End Select
    
    'ワークテーブル作成
    If Not MakeWork() Then
        Exit Function
    End If
    
    SubForm.Requery

    AllReload = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :DispHeader
'        機能             :ヘッダ表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DispHeader() As Boolean
On Error GoTo err_rtn
    
    Dim strSQL          As String
    Dim rsHeader        As Recordset
    Dim strMsg          As String
    
    strMsg = ""

    DispHeader = False
    
    doCmd.Hourglass True

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    'ヘッダ情報取得
    strSQL = MakeHeaderSql()
    Set rsHeader = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsHeader.EOF Then
        strMsg = "ヘッダ情報取得時にエラーが発生しました。"
        GoTo Exit_rtn
    End If
    
    '表示
    Call SetHeaderForm(rsHeader)

    '担当者名表示
    If Not DispTantoName() Then
        strMsg = "担当者名取得時にエラーが発生しました。"
        GoTo Exit_rtn
    End If
    
    DispHeader = True
    
    GoTo Exit_rtn
    
err_rtn:
    strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    If Not rsHeader Is Nothing Then rsHeader.Close: Set rsHeader = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    
    If strMsg <> "" Then
        MsgBox strMsg, vbCritical, P_ERROR
    End If
    
    doCmd.Hourglass False
    
    Exit Function

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeHeaderSql
'        機能             :ヘッダー取得SQL作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeHeaderSql() As String
    
    Dim strSQL      As String
    
    strSQL = strSQL & " SELECT YOUKT_UKDATE, "
    
    '2006/09/11 SHIBAZAKI 契約受付日を検索
    strSQL = strSQL & "        YOUKT_KUDATE, "
    
    strSQL = strSQL & "        YOUKT_UKTANTO, "
    strSQL = strSQL & "        YOUKT_YUKBN, "
    strSQL = strSQL & "        YUKBN_NAME.NAME_NAME AS YUKBNNAME, "
    strSQL = strSQL & "        YOUKT_UCODE, "
    strSQL = strSQL & "        YOUKT_NAME, "
    strSQL = strSQL & "        YOUKT_KANA, "
    strSQL = strSQL & "        YOUKT_TANM, "
    strSQL = strSQL & "        YOUKT_TAKA, "
    strSQL = strSQL & "        YOUKT_RENKBN, "
    strSQL = strSQL & "        RENKBN_NAME.NAME_NAME AS RENKBNNAME, "
    strSQL = strSQL & "        YOUKT_TEL, "
    strSQL = strSQL & "        YOUKT_KEITAI, "
    strSQL = strSQL & "        YOUKT_FAX, "
    strSQL = strSQL & "        YOUKT_USAGE, "
    strSQL = strSQL & "        USAGE_NAME.NAME_NAME AS USAGENAME, "
    strSQL = strSQL & "        CASE WHEN YOUKT_NEARKBN = -1 "
    strSQL = strSQL & "        THEN YARD_NAME + 'の近隣ヤードも紹介対象' "
    strSQL = strSQL & "        ELSE '' "
    strSQL = strSQL & "        END AS NEARKBN, "
    strSQL = strSQL & "        YOUKT_SIZE_FROM, "
    strSQL = strSQL & "        YOUKT_SIZE_TO, "
    strSQL = strSQL & "        YOUKT_STEP, "
    strSQL = strSQL & "        STEP_NAME.NAME_NAME AS STEPNAME, "
    strSQL = strSQL & "        YOUKT_PRICFROM, "
    strSQL = strSQL & "        YOUKT_PRICTO "
    
    '06.1.24 SHIBAZAKI 追加
    strSQL = strSQL & "       ,CASE WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_無し
    strSQL = strSQL & "            THEN NULL "
    strSQL = strSQL & "            ELSE MOVEKBN_NAME.NAME_NAME "
    strSQL = strSQL & "        END AS MOVEKBNNAME "
    
    '▼2007/04/16 tajima add  元々の希望ヤードを取得
    strSQL = strSQL & "        ,YOUKT_REQ_YCODE,YARD_NAME "
    '▲2007/04/16 tajima add
    
    strSQL = strSQL & "   FROM YOUK_TRAN "
    strSQL = strSQL & "        INNER JOIN NAME_MAST YUKBN_NAME "
    strSQL = strSQL & "           ON YUKBN_NAME.NAME_ID = '080' "
    strSQL = strSQL & "          AND YUKBN_NAME.NAME_CODE = YOUKT_YUKBN "
    
    strSQL = strSQL & "        INNER JOIN NAME_MAST RENKBN_NAME "
    strSQL = strSQL & "           ON RENKBN_NAME.NAME_ID = '085' "
    strSQL = strSQL & "          AND RENKBN_NAME.NAME_CODE = YOUKT_RENKBN "
    
    strSQL = strSQL & "         LEFT OUTER JOIN NAME_MAST USAGE_NAME "
    strSQL = strSQL & "           ON USAGE_NAME.NAME_ID = '086' "
    strSQL = strSQL & "          AND USAGE_NAME.NAME_CODE = YOUKT_USAGE "
    
    strSQL = strSQL & "         LEFT OUTER JOIN NAME_MAST STEP_NAME "
    strSQL = strSQL & "           ON STEP_NAME.NAME_ID = '090' "
    strSQL = strSQL & "          AND STEP_NAME.NAME_CODE = YOUKT_STEP "
    
    '06.1.24 SHIBAZAKI 追加
    strSQL = strSQL & "         LEFT  OUTER JOIN NAME_MAST MOVEKBN_NAME "
    strSQL = strSQL & "           ON MOVEKBN_NAME.NAME_ID = '098' "
    strSQL = strSQL & "          AND MOVEKBN_NAME.NAME_CODE = YOUKT_MOVEKBN "
    
    '▼2007/04/16 tajima add  元々の希望ヤードを取得
    strSQL = strSQL & "        LEFT OUTER JOIN YARD_MAST ON YARD_CODE = YOUKT_REQ_YCODE "
    '▲2007/04/16 tajima add
    
    strSQL = strSQL & "  WHERE YOUKT_UKNO = '" & SubHeader.Form.txt_Ukno & "' "

    MakeHeaderSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :SetHeaderForm
'        機能             :ヘッダ情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetHeaderForm(rsHeader As Recordset)

    With SubHeader.Form
        .txt_Ukdate = Format(rsHeader.Fields("YOUKT_UKDATE"), "yyyy/mm/dd")
        .txt_TantoCode = rsHeader.Fields("YOUKT_UKTANTO")
        .txt_Yukbn = rsHeader.Fields("YOUKT_YUKBN")

        '06.1.24 SHIBAZAKI 変更
'        .txt_YukbnName = rsHeader.Fields("YUKBNNAME")
        .txt_YukbnName = rsHeader.Fields("YUKBNNAME") & _
                         IIf(Nz(rsHeader.Fields("MOVEKBNNAME"), "") = "", _
                         "", _
                         "(" & rsHeader.Fields("MOVEKBNNAME") & ")")
        
        .txt_UserCode = rsHeader.Fields("YOUKT_UCODE")
        .txt_UserName = rsHeader.Fields("YOUKT_NAME")
        .txt_UserKana = rsHeader.Fields("YOUKT_KANA")
        .txt_Tanm = rsHeader.Fields("YOUKT_TANM")
        .txt_Taka = rsHeader.Fields("YOUKT_TAKA")
        .txt_Renkbn = rsHeader.Fields("YOUKT_RENKBN")
        .txt_RenkbnName = rsHeader.Fields("RENKBNNAME")
        .txt_Tel = rsHeader.Fields("YOUKT_TEL")
        .txt_Keitai = rsHeader.Fields("YOUKT_KEITAI")
        .txt_Fax = rsHeader.Fields("YOUKT_FAX")
        .txt_Usage = rsHeader.Fields("YOUKT_USAGE")
        .txt_UsageName = rsHeader.Fields("USAGENAME")
        .lbl_NearYard.Caption = rsHeader.Fields("NEARKBN")
        .txt_SizeFrom = rsHeader.Fields("YOUKT_SIZE_FROM")
        .txt_SizeTo = rsHeader.Fields("YOUKT_SIZE_TO")
        .txt_STEP = rsHeader.Fields("YOUKT_STEP")
        .txt_StepName = rsHeader.Fields("STEPNAME")
        .txt_PricFrom = rsHeader.Fields("YOUKT_PRICFROM")
        .txt_PricTo = rsHeader.Fields("YOUKT_PRICTO")
        
        '2006/09/11 SHIBAZAKI 契約受付日を検索
        .txt_Kudate = rsHeader.Fields("YOUKT_KUDATE")
    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DispTantoName
'        機能             :担当者名画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DispTantoName() As Boolean

    Dim strKaseServer       As String
    Dim strKaseConnect      As String
    Dim dbKase              As Database
    Dim rsTanto             As Recordset
    
    DispTantoName = False
    
    '加瀬DB接続
    'SQL-Server接続
    strKaseServer = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    If strKaseServer = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Function
    End If
    '接続文字列取得
    strKaseConnect = MSZZ007_M00()
    If strKaseConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Function
    End If
    
    '加瀬DB接続
    Set dbKase = Workspaces(0).OpenDatabase(strKaseServer, dbDriverNoPrompt, False, strKaseConnect)
    
    '担当マスタ検索
    Set rsTanto = dbKase.OpenRecordset(MakeTantoSql(), dbOpenSnapshot, dbSQLPassThrough)
    
    If rsTanto.EOF Then
        Exit Function
    End If
    
    '担当名表示
    SubHeader.Form.txt_TantoName = rsTanto.Fields("TANTM_TANTN")
    
    If Not rsTanto Is Nothing Then rsTanto.Close: Set rsTanto = Nothing
    If Not dbKase Is Nothing Then dbKase.Close: Set dbKase = Nothing

    DispTantoName = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeTantoSql
'        機能             :担当者名取得SQL作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeTantoSql() As String

    Dim strSQL      As String
    
    strSQL = strSQL & " SELECT TANTM_TANTN "
    strSQL = strSQL & "   FROM TANT_MAST "
    strSQL = strSQL & "  WHERE TANTM_TANTC = '" & SubHeader.Form.txt_TantoCode & "' "
    strSQL = strSQL & "    AND TANTM_BUMOC = '" & strBumonCode & "' "
    
    MakeTantoSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeWork
'        機能             :ワークテーブル作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeWork() As Boolean
On Error GoTo err_rtn
    'FVS550_W01：サブフォームとリンクするテーブル
    
    Dim rsSqlServer     As Recordset
    Dim rsAccess        As Recordset
    Dim lngRecordCount  As Long
    Dim strMsg          As String
    
    MakeWork = False
    
    doCmd.Hourglass True
    
    strMsg = ""
    
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    'ワークテーブルの元ネタ取得
    Set rsSqlServer = dbSQLServer.OpenRecordset(MakeIntrSql(), dbOpenSnapshot, dbSQLPassThrough)
    
    'ACCESS接続
    Set dbAccess = CurrentDb

    'Access側ワークテーブル削除
    Call ClearWork(P_WORK_TABLE)

    'ワークテーブルレコード追加
    Set rsAccess = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    lngRecordCount = 0
    
    Call AddWork(rsSqlServer, rsAccess)
    
    If rsSqlServer.RecordCount = 0 Then
        strMsg = "取置き履歴が存在しません。"
        GoTo Exit_rtn
    End If
    
    MakeWork = True
    
    GoTo Exit_rtn
    
err_rtn:
    strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    
    doCmd.Hourglass False
    
    If strMsg <> "" Then
        MsgBox strMsg, vbCritical, P_ERROR
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :ClearWork
'        機能             :ワークテーブルクリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClearWork(strTable As String)

    Dim tdfNew      As TableDef
    
    SubForm.Form.RecordSource = ""
    
    'ワークテーブル削除
    If TableExist(strTable) Then
        Call dbAccess.TableDefs.Delete(strTable)
    End If
    
    'ワークテーブル作成
    Set tdfNew = dbAccess.CreateTableDef(strTable)
    Call FieldAppend(tdfNew)
    Call dbAccess.TableDefs.Append(tdfNew)
    
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    
    SubForm.Form.RecordSource = P_SUBFORM_RECORD_SOURCE
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :FieldAppend
'        機能             :ワークテーブル列作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew

        Call .Fields.Append(.CreateField("UKNO", DataTypeEnum.dbText, 11))
        Call .Fields.Append(.CreateField("INTRONO", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("YCODE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("CNO", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("NAME", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("ADDR_1", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("ADDR_2", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("ADDR_3", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("SIZE", DataTypeEnum.dbDouble))
        Call .Fields.Append(.CreateField("RENTKG", DataTypeEnum.dbText, 15))
        Call .Fields.Append(.CreateField("INSED", DataTypeEnum.dbText, 8))
        Call .Fields.Append(.CreateField("STATUS", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("PRINTED", DataTypeEnum.dbText, 2))
        Call .Fields.Append(.CreateField("CAMPAIGN1", DataTypeEnum.dbText, 60))
        Call .Fields.Append(.CreateField("CAMPAIGN2", DataTypeEnum.dbText, 60))
        Call .Fields.Append(.CreateField("CAMPAIGN3", DataTypeEnum.dbText, 60))
        Call .Fields.Append(.CreateField("CAMPAIGN4", DataTypeEnum.dbText, 60))
        Call .Fields.Append(.CreateField("INTROKBN", DataTypeEnum.dbText, 1))
        Call .Fields.Append(.CreateField("ZAPPI_NAME", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("ZAPPI_GAKU", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("POSSIBLE_KEEP", DataTypeEnum.dbText, 4))
        Call .Fields.Append(.CreateField("POSSIBLE_KEEP_FLAG", DataTypeEnum.dbText, 1))
        Call .Fields.Append(.CreateField("REASON", DataTypeEnum.dbText, 60))
        
        '2006/09/11 SHIBAZAKI 営業期間追加
        Call .Fields.Append(.CreateField("YARD_BEGIN_DAY", DataTypeEnum.dbText, 10))
        Call .Fields.Append(.CreateField("YARD_RENTEND_DAY", DataTypeEnum.dbText, 10))
   
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
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
Private Function TableExist(strTableName As String) As Boolean

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
'        MODULE_NAME      :MakeIntrSql
'        機能             :紹介トラン検索SQL文作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeIntrSql() As String

    Dim strSQL      As String
    
    strSQL = strSQL & " SELECT INTRT_UKNO, "
    strSQL = strSQL & "        INTRT_INTRONO, "
    strSQL = strSQL & "        INTRT_YCODE, "
    strSQL = strSQL & "        INTRT_NO, "
    strSQL = strSQL & "        YARD_NAME, "
    strSQL = strSQL & "        YARD_ADDR_1, "
    strSQL = strSQL & "        YARD_ADDR_2, "
    strSQL = strSQL & "        YARD_ADDR_3, "
    
    '2006/09/11 SHIBAZAKI 営業期間追加
    strSQL = strSQL & "        YARD_BEGIN_DAY, "
    strSQL = strSQL & "        YARD_RENTEND_DAY, "
    
    strSQL = strSQL & "        CNTA_SIZE, "
    strSQL = strSQL & "        PRIC_PRICE + ISNULL(CNTA_PRICE_DIFF,0) + ISNULL(PRIC_EZAPPI,0) AS CHRGE, " '賃料＋差額＋毎月雑費
    strSQL = strSQL & "        CNTA_REASON, "
    strSQL = strSQL & "        CNTA_UKNO, "
    strSQL = strSQL & "        INTRT_INSED, "
    strSQL = strSQL & "        INTRT_INSEJ, "
    strSQL = strSQL & "        INTRT_INTROKBN, "
    strSQL = strSQL & "        INTROKBN.NAME_NAME AS INTROKBN_NAME, "
    strSQL = strSQL & "        CASE WHEN INTRT_FOUTD IS NOT NULL "      '印刷日が入ってる場合は”済”
    strSQL = strSQL & "        THEN '" & P_印刷済み文字列 & "' "
    strSQL = strSQL & "        ELSE '' "
    strSQL = strSQL & "        END AS PRINTED, "
    strSQL = strSQL & "        YARD_SEV1N, "
    strSQL = strSQL & "        YARD_SEV2N, "
    strSQL = strSQL & "        YARD_SEV3N, "
    strSQL = strSQL & "        YARD_ENDEN, "
    strSQL = strSQL & "        ZAPPI.NAME_NAME AS EZAPPI_NAME, "
    strSQL = strSQL & "        PRIC_EZAPPI, "
    strSQL = strSQL & "        INTRT_UPDAD, "
    strSQL = strSQL & "        INTRT_UPDAJ "
    
    strSQL = strSQL & "   FROM INTR_TRAN "
    strSQL = strSQL & "        INNER JOIN CNTA_MAST "
    strSQL = strSQL & "           ON CNTA_CODE = INTRT_YCODE "
    strSQL = strSQL & "          AND CNTA_NO = INTRT_NO "
    
    strSQL = strSQL & "        INNER JOIN YARD_MAST "
    strSQL = strSQL & "           ON YARD_CODE = INTRT_YCODE "
    
    strSQL = strSQL & "        INNER JOIN NAME_MAST INTROKBN "
    strSQL = strSQL & "           ON INTROKBN.NAME_CODE = INTRT_INTROKBN "
    strSQL = strSQL & "          AND INTROKBN.NAME_ID = '097' "

    strSQL = strSQL & "         LEFT OUTER JOIN PRIC_TABL "
    strSQL = strSQL & "           ON PRIC_YCODE = CNTA_CODE "
    strSQL = strSQL & "          AND PRIC_USAGE = CNTA_USAGE "
    strSQL = strSQL & "          AND PRIC_SIZE = CNTA_SIZE "
    strSQL = strSQL & "          AND PRIC_STEP = CNTA_STEP "
    
    strSQL = strSQL & "         LEFT OUTER JOIN NAME_MAST ZAPPI "
    strSQL = strSQL & "           ON ZAPPI.NAME_CODE = PRIC_EZAPPI_CODE "
    strSQL = strSQL & "          AND ZAPPI.NAME_ID = '096' "
    
    strSQL = strSQL & "  WHERE INTRT_UKNO = '" & SubHeader.Form.txt_Ukno & "' "
    
    strSQL = strSQL & "  ORDER BY INTRT_INTRONO DESC "
    
    MakeIntrSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :AddWork
'        機能             :作業テーブル追加
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub AddWork(rsSource As Recordset, rsDestination As Recordset)

    Dim strPossibleKeep()   As String
    Dim lngKeepCount        As Long
    
    lngKeepCount = -1
    strPossibleKeep = Split("")
    
    Do While Not rsSource.EOF
        With rsDestination
            .AddNew
            
            .Fields("UKNO") = rsSource.Fields("INTRT_UKNO")
            .Fields("INTRONO") = rsSource.Fields("INTRT_INTRONO")
            .Fields("YCODE") = rsSource.Fields("INTRT_YCODE")
            .Fields("CNO") = rsSource.Fields("INTRT_NO")
            .Fields("NAME") = rsSource.Fields("YARD_NAME")
            .Fields("ADDR_1") = rsSource.Fields("YARD_ADDR_1")
            .Fields("ADDR_2") = rsSource.Fields("YARD_ADDR_2")
            .Fields("ADDR_3") = rsSource.Fields("YARD_ADDR_3")
            .Fields("SIZE") = rsSource.Fields("CNTA_SIZE")
            .Fields("RENTKG") = rsSource.Fields("CHRGE")
            .Fields("INSED") = rsSource.Fields("INTRT_INSED")
            .Fields("STATUS") = rsSource.Fields("INTROKBN_NAME")
            .Fields("PRINTED") = rsSource.Fields("PRINTED")
            .Fields("CAMPAIGN1") = rsSource.Fields("YARD_SEV1N")
            .Fields("CAMPAIGN2") = rsSource.Fields("YARD_SEV2N")
            .Fields("CAMPAIGN3") = rsSource.Fields("YARD_SEV3N")
            .Fields("CAMPAIGN4") = rsSource.Fields("YARD_ENDEN")
            .Fields("INTROKBN") = rsSource.Fields("INTRT_INTROKBN")
            .Fields("ZAPPI_NAME") = rsSource.Fields("EZAPPI_NAME")
            .Fields("ZAPPI_GAKU") = rsSource.Fields("PRIC_EZAPPI")
            .Fields("REASON") = rsSource.Fields("CNTA_REASON")
            
            '2006/09/11 SHIBAZAKI 営業期間追加
            .Fields("YARD_BEGIN_DAY") = rsSource.Fields("YARD_BEGIN_DAY")
            .Fields("YARD_RENTEND_DAY") = rsSource.Fields("YARD_RENTEND_DAY")
            
            '再取置き可能か？
            If PossibleKeep(rsSource, strPossibleKeep) Then
                '可能
                .Fields("POSSIBLE_KEEP") = P_可能文字列
                .Fields("POSSIBLE_KEEP_FLAG") = "1"
                lngKeepCount = lngKeepCount + 1
                ReDim Preserve strPossibleKeep(0 To lngKeepCount)
                strPossibleKeep(lngKeepCount) = .Fields("YCODE") & .Fields("CNO")
            Else
                '不可
                .Fields("POSSIBLE_KEEP") = ""
                .Fields("POSSIBLE_KEEP_FLAG") = "0"
            End If
            
            .UPDATE
        End With
        
        rsSource.MoveNext
    Loop
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :PossibleKeep
'        機能             :再取置き可不可チェック
'        IN               :
'        OUT              :True=可 False=不可
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function PossibleKeep(rsIntro As Recordset, _
                              strPossibleKeep() As String) As Boolean

    Dim strTimeStampD       As String
    Dim strTimeStampJ       As String
    Dim lngKeepCount        As Long
    
    PossibleKeep = False
    
    '再取置きボタンが使用不可ならば、常に再取置き不可
    If Not SubForm.Form.cmd_Torioki.Enabled Then
        Exit Function
    End If
    
    With rsIntro
        
        '紹介区分が「キャンセル」「紹介不要」、
        'かつ紹介待ち予約番号が現在の予約番号の場合に、再取置き可能とする
        If (.Fields("INTRT_INTROKBN") = P_INTRO_STATUS_CANCEL Or _
            .Fields("INTRT_INTROKBN") = P_INTRO_STATUS_NOT_INTRO) And _
           .Fields("CNTA_UKNO") = SubHeader.Form.txt_Ukno _
        Then
            PossibleKeep = True
            '既に同じコンテナが再取置き可能とされていた場合は、再取置き不可。
            For lngKeepCount = 0 To UBound(strPossibleKeep)
                If (.Fields("INTRT_YCODE") & .Fields("INTRT_NO")) = strPossibleKeep(lngKeepCount) Then
                    PossibleKeep = False
                    Exit For
                End If
            Next
        End If
    End With
    
End Function

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
    
    Dim objForm         As Object
    
    'ACCESS接続
    Set dbAccess = CurrentDb

    'Access側ワークテーブル削除
    Call ClearWork(P_WORK_TABLE)
    
    dbAccess.Close
    Set dbAccess = Nothing
    
    '2006/03/09 SHIBAZAKI 追加
    If blnReloadSkip Then
        Exit Sub
    End If
    
    Set objForm = Forms(strFormFrom)
    If Not objForm Is Nothing Then
        Call objForm.subReload(IIf(blnUpdate, 1, 0))
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :UpdateProc
'        機能             :更新処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UpdateProc(lngIntroNo As Long, _
                           strKbn As String, _
                           strYcode As String, _
                           strCno As String, _
                           strNowKbn As String) As Boolean
On Error GoTo err_rtn

    Dim objWs               As Workspace
    Dim strMsg              As String
    
    '2006/09/11 SHIBAZAKI 受付日＋８日後が営業期間外ならばエラーにする
    Dim strBeginDay         As String
    Dim strEndDay           As String
    Dim strUkDate           As String
    
    UpdateProc = False
    
    Select Case strKbn
    Case P_INTRO_STATUS_KEEP
        strMsg = P_SAVE_CONFIRM_再取置き
    Case P_INTRO_STATUS_ACCEPT
        '2006/09/11 SHIBAZAKI 受付日＋サービス期間が営業期間外ならばエラーにする↓
        strBeginDay = IIf(Nz(SubForm.Form.txt_BeginDay, "") = "", "0001/01/01", SubForm.Form.txt_BeginDay)
        strEndDay = IIf(Nz(SubForm.Form.txt_EndDay, "") = "", "9999/12/31", SubForm.Form.txt_EndDay)
        
        strUkDate = Format(DateAdd("d", intServiceDay, DATE), "yyyy/mm/dd")
        
        If DateDiff("d", strBeginDay, strUkDate) < 0 Or _
           DateDiff("d", strEndDay, strUkDate) > 0 _
        Then
            MsgBox "本日から" & intServiceDay & "日後[" & strUkDate & "]を希望開始日に設定しますが、その日付が営業期間外です。" & vbCrLf & _
                   "営業期間[" & Nz(SubForm.Form.txt_BeginDay, "") & "～" & Nz(SubForm.Form.txt_EndDay, "") & "]" & vbCrLf & _
                   "予約入力で受付にしてください。", vbInformation, P_ERROR
            Exit Function
        End If
        '2006/09/11 SHIBAZAKI 受付日＋サービス期間が営業期間外ならばエラーにする↑
        
        strMsg = P_SAVE_CONFIRM_受付
    Case P_INTRO_STATUS_NOT_INTRO
        strMsg = P_SAVE_CONFIRM_紹介不要
    End Select
    If MsgBox(strMsg & "よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
        Exit Function
    End If
    
    strMsg = ""
    
    doCmd.Hourglass True
    
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    'チェック
    If Not CheckProc(strKbn, strYcode, strCno, strMsg) Then
        GoTo err_rtn
    End If

    Set objWs = DBEngine.Workspaces(0)
    objWs.BeginTrans

    '御紹介トラン更新
    If Not UpdateIntr(lngIntroNo, strKbn, strYcode, strCno, strNowKbn, strMsg) Then
        GoTo err_rtn
    End If
    
    '紹介不要以外の場合、予約トランの予約状態を更新
    If strKbn <> P_INTRO_STATUS_NOT_INTRO Then
        '予約受付トラン更新
        If Not UpdateYouk(strKbn, strYcode, strCno, strMsg) Then
            GoTo err_rtn
        End If
    End If
    
    '2006/03/09 SHIBAZAKI 追加
    '再取置きの場合、存在すれば受付トランのデータを削除する。
    If strKbn = P_INTRO_STATUS_KEEP Then
        If Not DeleteRcpt(strMsg) Then
            GoTo err_rtn
        End If
    End If
    
    objWs.CommitTrans

    blnUpdate = True
    UpdateProc = True
    GoTo Exit_rtn
    
err_rtn:
    If Not objWs Is Nothing Then
        objWs.Rollback
    End If
    
    If strMsg = "" Then
        strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    End If
    
    Err.Clear

Exit_rtn:
    If Not objWs Is Nothing Then Set objWs = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    
    If strMsg <> "" Then
        MsgBox strMsg, vbCritical, P_ERROR
    End If
    
    doCmd.Hourglass False

End Function

'==============================================================================*
'
'        MODULE_NAME      :CheckProc
'        機能             :チェック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CheckProc(strKbn As String, _
                           strYcode As String, _
                           strCno As String, _
                           ByRef strMsg As String) As Boolean

    CheckProc = False
    
    '紹介不要の場合はチェックしない。
    If strKbn = P_INTRO_STATUS_NOT_INTRO Then
        CheckProc = True
        Exit Function
    End If
    
    If strKbn = P_INTRO_STATUS_ACCEPT Then
        '受付に変更するためのチェック
        If CheckLending(dbSQLServer, _
                        pcintYOYAKU_CHECK, _
                        Val(strYcode), _
                        Val(strCno), _
                        strMsg, _
                        SubHeader.Form.txt_Ukno) <> pcintCHECK_OK _
        Then
            '受付のためのチェックならば、チェックがNGならば不可
            Exit Function
        End If
    Else
        '再取置きするためのチェック
        If CheckLending(dbSQLServer, _
                        pcintSHOUMATI_UNITCHECK, _
                        Val(strYcode), _
                        Val(strCno), _
                        strMsg) <> pcintSHOUKAI_WAIT _
        Then
            '再取置きのためのチェックならば、紹介待ちになっていなければ不可。
            strMsg = "紹介待ちデータではありません。"
            Exit Function
        End If
    End If
    
    CheckProc = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :UpdateIntr
'        機能             :ご紹介区分更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function UpdateIntr(lngIntroNo As Long, _
                            strKbn As String, _
                            strYcode As String, _
                            strCno As String, _
                            strNowKbn As String, _
                            ByRef strMsg As String) As Boolean
    
    UpdateIntr = False
    
    Select Case strKbn
    Case P_INTRO_STATUS_NOT_INTRO:
        '紹介不要
        If strNowKbn = P_INTRO_STATUS_KEEP Then
            '取置きから紹介不要に変更
            '  ⇒ 紹介区分＝「紹介不要」　コンテナ指定で紹介待ち作成
            If Not IntroWaitSingle(dbSQLServer, _
                                   SubHeader.Form.txt_Ukno, _
                                   strKbn, _
                                   strMsg, _
                                   Val(strYcode), _
                                   Val(strCno)) _
            Then
                Exit Function
            End If
        Else
            'キャンセルから紹介不要に変更
            '⇒自分自身の紹介区分を変えるだけ
            If Not CmnUpdateIntrTran(dbSQLServer, _
                                     SubHeader.Form.txt_Ukno, _
                                     strNowKbn, _
                                     strKbn, _
                                     strMsg, _
                                     lngIntroNo) _
            Then
                Exit Function
            End If
        End If
    Case P_INTRO_STATUS_KEEP:
        '取置き
        '  ⇒ 紹介区分＝「取置き」
        If Not CmnUpdateIntrTran(dbSQLServer, _
                                 SubHeader.Form.txt_Ukno, _
                                 strNowKbn, _
                                 strKbn, _
                                 strMsg, _
                                 lngIntroNo) _
        Then
            Exit Function
        End If
        '⇒紹介待ち状態を消す
        If Not CmnUpdateCntaMast(dbSQLServer, Val(strYcode), Val(strCno), strMsg) Then
            Exit Function
        End If
    Case P_INTRO_STATUS_ACCEPT:
        '受付
        '  ⇒ 紹介区分＝「受付」
        '  ⇒ 選ばれたコンテナ以外を紹介区分＝「キャンセル」および紹介待ち作成
        If Not IntroChgAndIntroWait(dbSQLServer, _
                                    SubHeader.Form.txt_Ukno, _
                                    strKbn, _
                                    strMsg, _
                                    Val(strYcode), _
                                    Val(strCno)) _
        Then
            Exit Function
        End If
    End Select

    UpdateIntr = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :UpdateYouk
'        機能             :予約受付更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function UpdateYouk(strKbn As String, strYcode As String, strCno As String, ByRef strMsg As String) As Boolean
On Error GoTo err_rtn

    Dim rsYouk              As Recordset
    
    UpdateYouk = False
    
    '更新対象取得
    Set rsYouk = dbSQLServer.OpenRecordset(GetYoukUpdate(), dbOpenDynaset)
    
    If rsYouk.EOF Then
        strMsg = "予約受付トラン更新情報読み込み中にエラーが発生しました"
        GoTo Exit_rtn
    End If
    
    '予約トラン更新
    With rsYouk
        .Edit
        
        If strKbn = P_INTRO_STATUS_ACCEPT Then
            '契約受付
            .Fields("YOUKT_YUKBN") = P_BOOKING_STATUS_ACCEPT
            
            'コンテナ番号を受付したコンテナ番号に変更する
            .Fields("YOUKT_NO") = strCno
            
            '受付期限日
            .Fields("YOUKT_TKDATE") = DateAdd("d", GetKeiukeTld(), DATE)
            
            '自動区分を「無し」にする
            .Fields("YOUKT_AUTOKBN") = P_AUTOKBN_無し
        
            '2006/09/11 SHIBAZAKI 受付としたときに希望開始日に受付日＋サービス期間を設定する
            .Fields("YOUKT_KKDATE") = DateAdd("d", intServiceDay, DATE)
        Else
            '予約取置き
            .Fields("YOUKT_YUKBN") = P_BOOKING_STATUS_KEEP
            
            '予約取消日をクリアする
            .Fields("YOUKT_YTDATE") = Null
            
            'コンテナ番号をクリアする
            .Fields("YOUKT_NO") = Null
            
            '2006/03/09 SHIBAZAKI 追加
            '受付番号をクリア
            .Fields("YOUKT_RCPT_NO") = Null

        End If
        
        .Fields("YOUKT_YCODE") = strYcode
        .Fields("YOUKT_KUDATE") = DATE
        .Fields("YOUKT_UPDAD") = Format$(DATE, "yyyymmdd")          '更新日付(var/Null=Y)
        .Fields("YOUKT_UPDAJ") = Format$(time, "hhmmss")            '更新時刻(var/Null=Y)
        .Fields("YOUKT_UPDPB") = Me.NAME                            '更新プログラムID(var/Null=Y)
        .Fields("YOUKT_UPDUB") = LsGetUserName()                    '更新ユーザーID(var/Null=Y)
        
        .UPDATE
    End With
    
    UpdateYouk = True

    GoTo Exit_rtn

err_rtn:
    strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    If Not rsYouk Is Nothing Then rsYouk.Close: Set rsYouk = Nothing

End Function

'==============================================================================*
'
'        MODULE_NAME      :GetYoukUpdate
'        機能             :予約受付更新情報取得SQL作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetYoukUpdate() As String

    Dim strSQL      As String
    
    strSQL = strSQL & "SELECT * FROM YOUK_TRAN "
    strSQL = strSQL & "  WHERE YOUKT_UKNO = '" & SubHeader.Form.txt_Ukno & "' "
    
    GetYoukUpdate = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :GetYoukUpdate
'        機能             :受付期限日数取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetKeiukeTld() As Integer

    Dim strSQL      As String
    Dim rsCont      As Recordset
    
    strSQL = strSQL & "SELECT CONT_KEIUKE_TLD FROM CONT_MAST "
    strSQL = strSQL & "  WHERE CONT_KEY = " & P_CONT_KEY

    Set rsCont = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    GetKeiukeTld = rsCont.Fields("CONT_KEIUKE_TLD")

    rsCont.Close
    Set rsCont = Nothing

End Function

'2006/03/09 SHIBAZAKI 追加
'==============================================================================*
'
'        MODULE_NAME      :DeleteRcpt
'        機能             :受付期限日数取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DeleteRcpt(ByRef strMsg As String) As Boolean
On Error GoTo err_rtn

    Dim rsRcpt              As Recordset
    Dim strSQL              As String
    
    DeleteRcpt = False
    
    '更新対象取得
    strSQL = "SELECT * FROM RCPT_TRAN WHERE RCPT_NO = '" & SubHeader.Form.txt_Ukno & "' "
    Set rsRcpt = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    '受付トラン削除
    With rsRcpt
        If Not .EOF Then
            .Delete
        End If
        .Close
    End With
    Set rsRcpt = Nothing
    
    DeleteRcpt = True

    GoTo Exit_rtn

err_rtn:
    strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    If Not rsRcpt Is Nothing Then rsRcpt.Close: Set rsRcpt = Nothing

End Function

