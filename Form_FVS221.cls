VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS221"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　解約ヤード詳細
'   プログラムＩＤ　：　FVS221
'   作　成　日　　　：  2009/04/30
'   作　成　者　　　：  KIHARA
'   Ver             ：  0.0
'   備考            ：  FVS550 取置き履歴を参考に作成
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   Ver             ：
'   修　正　内　容　：
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
Public strFormFrom      As String

'データを更新したらTrueになる
Private blnUpdate       As Boolean

Public blnReloadSkip    As Boolean  'Trueならば終了時に呼び出しもとのリフレッシュは行わない

'サービス期間をコントロールマスタから取得
Private intServiceDay   As Integer

'ヤードコード
Private strYcode        As String
'コンテナ番号
Private strCno          As String
'紹介先ヤードコード
Private strIntrYcd      As String
'紹介先コンテナ番号
Private strIntrCno      As String
'月額使用料
Private strRENTKG       As String

'▼ヤード解約情報保存時に使用▼
'Private WRI_ObjWS       As Workspace
'Private WRI_objDb       As Database

''回答区分
'Private strKaitoKbn     As String
''ヤード営業開始日
'Private strYBeginDay    As String
''サービス期間日数
'Private strCONTServDD   As Date

''移動先受付情報
'Private IDOSAKI_RCPT        As FVS221Type_RCPT
'Private Type FVS221Type_RCPT
'    RCPT_NO                 As String   '受付番号
'    RCPT_UCODE              As String   '顧客コード
'    RCPT_BUTU               As String   '主な収納物
'    RCPT_HOSYICD            As String   '保証区分
'    RCPT_HOSYO_CD           As String   '保証会社コード
'    RCPT_SEV_EXPIRE         As String   'キャンペーン適用満了年月
'    RCPT_SEV_EXPIRE_MONTH   As String   'キャンペーン適用月数
'    RCPT_SEV1N              As String   'サービス内容１
'    RCPT_SEV2N              As String   'サービス内容２
'    RCPT_SEV3N              As String   'サービス内容３
'    RCPT_ENDEN              As String   'サービス期間
'    RCPT_MITU_CD            As String   '見積もり請求区分
'    RCPT_SEIKYU_CD          As String   '初期費用請求方法
'    RCPT_SECUKG             As String   '保証金
'    RCPT_KAGIICD            As String   '鍵区分
'    RCPT_KAGI_NO            As String   '鍵番号
'    RCPT_KAGI_LENDNUM       As String   '鍵貸し出し本数
'    RCPT_SEND_NAME          As String   '鍵送付先名
'    RCPT_SEND_HONORIFIC     As String   '鍵送付先敬称
'    RCPT_SEND_YUBINO        As String   '鍵送付先郵便番号
'    RCPT_SEND_ADR_1         As String   '鍵送付先住所１
'    RCPT_SEND_ADR_2         As String   '鍵送付先住所２
'    RCPT_SEND_ADR_3         As String   '鍵送付先住所３
'    RCPT_SEND_TEL           As String   '鍵送付先電話番号
'    RCPT_CARG_ACPTNO        As String   '受注契約番号
'End Type
'▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

Private Const P_WORK_TABLE          As String = "FVS221_W01"

'取置き一覧サブフォームレコードソース
Private Const P_SUBFORM_RECORD_SOURCE   As String = _
                            " SELECT * FROM " & P_WORK_TABLE & _
                            "  ORDER BY INTRONO DESC; "

'ご紹介区分
Private Const P_INTRO_STATUS_CANCEL         As String = "6" 'ご紹介区分「キャンセルした」
Private Const P_INTRO_STATUS_NOT_INTRO      As String = "9" 'ご紹介区分「再度ご紹介不要」

Private Const P_パラメータ数                        As Integer = 5

'前画面からのパラメータのインデックス 上から順にカンマ区切りで渡してもらう
Private Const P_パラメータ_呼び出し元画面             As Integer = 0
Private Const P_パラメータ_ヤードコード               As Integer = 1
Private Const P_パラメータ_コンテナ番号               As Integer = 2
Private Const P_パラメータ_紹介先情報                 As Integer = 3
Private Const P_パラメータ_月額使用料                 As Integer = 4

'再取置き可能な場合に表示する文字列
Private Const P_可能文字列                  As String = "可能"

'印刷済みの場合に表示する文字列
Private Const P_印刷済み文字列              As String = "済"

'コントロールマスタのキー
Private Const P_CONT_KEY                    As Integer = 1

'▼ヤード解約情報保存時に使用▼
'保存時確認メッセージ
'Private Const strmsg保存確認                As String = "保存します。よろしいですか？"

Private Const pcintKONT_DB          As Integer = 0  ' コンテナDB
Private Const pcintKASE_DB          As Integer = 1  ' 加瀬DB

Private pobjKONT_DB                 As Database     ' コンテナDB
Private pobjKASE_DB                 As Database     ' 加瀬DB

'▼ヤード解約情報保存時に使用▼
'更新時エラーメッセージ
'Private Const strmsg更新エラー      As String = "更新中にエラーが発生しました。（処理は中断されました。）

'画面ID
Private Const PROG_ID = "FVS221"

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        機能             :終了ボタンクリック処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next
    '終了メッセージを表示
    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton1, P_KAKUNIN) = vbYes Then
        doCmd.Close acForm, Me.FormName
    End If
    
Exit_rtn:
    Exit Sub
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activateイベント
'        CREATE_DATE      :2009/04/30
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
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    KeyCode = 0
    Select Case KeyCode
        Case vbKeyF1    'F1=「終了」
            Me.cmd_終了.SetFocus
'▼ヤード解約情報保存時に使用▼
'        Case vbKeyF2    'F2=「保存」
'            Me.cmd_保存.SetFocus
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form Loadイベント
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim blnRet      As Boolean
    Dim strTemp     As String
    Dim strParam()  As String
    Dim strTemp2()  As String
    
    blnUpdate = False
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
    strYcode = strParam(P_パラメータ_ヤードコード)
    strCno = strParam(P_パラメータ_コンテナ番号)
    
    strTemp2() = Split(strParam(P_パラメータ_紹介先情報), "-")
    strIntrYcd = strTemp2(0) '紹介先ヤードコード
    strIntrCno = strTemp2(1) '紹介先コンテナ番号
    
    strRENTKG = strParam(P_パラメータ_月額使用料)
    
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
    
    'サービス期間をコントロールマスタから取得
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
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function AllReload() As Boolean

    AllReload = False
    
    'ヘッダ情報取得
    If Not DispHeader() Then
        Exit Function
    End If
    
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
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DispHeader() As Boolean
On Error GoTo err_rtn
    
    Dim rsHeader       As Recordset     '契約・顧客情報
'    Dim rsHeader2       As Recordset   'ヤード解約回答情報　'▼ヤード解約情報保存時に使用▼

    Dim strMsg         As String
    strMsg = ""

    DispHeader = False
    doCmd.Hourglass True

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    'ヘッダ情報取得
    Set rsHeader = dbSQLServer.OpenRecordset(MakeHeaderSql(), dbOpenSnapshot, dbSQLPassThrough)
'    Set rsHeader2 = dbSqlServer.OpenRecordset(MakeHeaderSql2(rsHeader), dbOpenSnapshot, dbSQLPassThrough)　'▼ヤード解約情報保存時に使用▼

    If rsHeader.EOF Then
        strMsg = "ヘッダ契約情報･顧客情報取得時にエラーが発生しました。"
        GoTo Exit_rtn
    End If
    
'▼ヤード解約情報保存時に使用▼
'    '値をセット
'    IDOSAKI_RCPT.RCPT_UCODE = rsHeader.Fields("CARG_UCODE")
   
    '表示
    Call SetHeaderForm(rsHeader)
    
'▼ヤード解約情報保存時に使用▼
'    If Not rsHeader2.EOF Then
'        Call SetHeaderForm2(rsHeader2)
'    End If

    DispHeader = True
    
    GoTo Exit_rtn
    
err_rtn:
    strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    If Not rsHeader Is Nothing Then rsHeader.Close: Set rsHeader = Nothing
'    If Not rsHeader2 Is Nothing Then rsHeader2.Close: Set rsHeader2 = Nothing　'▼ヤード解約情報保存時に使用▼
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
'        機能             :ヘッダー取得SQL作成（契約情報・顧客情報）
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeHeaderSql() As String
    
    Dim strSQL      As String
    
    strSQL = strSQL & " SELECT USER_NAME, "
    strSQL = strSQL & "        USER_KANA, "
    strSQL = strSQL & "        USER_TANM, "
    strSQL = strSQL & "        USER_TAKA, "
    strSQL = strSQL & "        USER_YUBINO, "
    strSQL = strSQL & "        USER_ADR_1, "
    strSQL = strSQL & "        USER_ADR_2, "
    strSQL = strSQL & "        USER_ADR_3, "
    strSQL = strSQL & "        USER_TEL, "
    strSQL = strSQL & "        USER_KEITAI, "
    strSQL = strSQL & "        USER_FAX, "
    strSQL = strSQL & "        USER_MAIL, "
    strSQL = strSQL & "        USER_KKBN, "
    strSQL = strSQL & "        USER_SKBN, "
    strSQL = strSQL & "        USER_SDAY, "
    strSQL = strSQL & "        USER_BIKO, "
    strSQL = strSQL & "        CARG_ACPTNO, "
    strSQL = strSQL & "        CARG_FSDATE, "
    strSQL = strSQL & "        CARG_STDATE, "
    strSQL = strSQL & "        CARG_EDDATE, "
    strSQL = strSQL & "        CARG_KYDATE, "
    strSQL = strSQL & "        CARG_KAIUKEDATE, "
    strSQL = strSQL & "        CARG_KAICD, "
    strSQL = strSQL & "        CARG_HOSYICD, "
    strSQL = strSQL & "        CARG_HOSYO_CD, "
    strSQL = strSQL & "        CARG_SECUKG, "
    strSQL = strSQL & "        CARG_HOSYD, "
    strSQL = strSQL & "        CARG_RENTKG, "
    strSQL = strSQL & "        CARG_SYOZEI, "
'    strSQL = strSQL & "        CARG_UCODE, "　'▼ヤード解約情報保存時に使用▼
    strSQL = strSQL & "        CARG_BIKO "
    
    strSQL = strSQL & "   FROM CARG_FILE "
    strSQL = strSQL & "        INNER JOIN USER_MAST "
    strSQL = strSQL & "           ON USER_CODE = CARG_UCODE "
    
    strSQL = strSQL & "         LEFT OUTER JOIN YARD_MAST "
    strSQL = strSQL & "           ON YARD_CODE = CARG_YCODE "
    
    strSQL = strSQL & "         LEFT OUTER JOIN RCPT_TRAN "
    strSQL = strSQL & "           ON RCPT_YCODE = CARG_YCODE "
    strSQL = strSQL & "          AND RCPT_CNO = CARG_NO "
        
    strSQL = strSQL & "  WHERE CARG_YCODE = '" & strYcode & "' "
    strSQL = strSQL & "    AND CARG_NO = '" & strCno & "'  "

    MakeHeaderSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeHeaderSql2
'        機能             :ヘッダー取得SQL作成（ヤード解約情報）
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function MakeHeaderSql2(rsHeader As Recordset) As String
'
'    Dim strSQL    As String
'    Dim strAcptNo As String         '受注契約番号
'
'    strAcptNo = rsHeader.Fields("CARG_ACPTNO")
'
'    strSQL = strSQL & " SELECT CANYT_KBN, "
'    strSQL = strSQL & "        CANYT_YE_DATE, "
'    strSQL = strSQL & "        CANYT_MOVEKBN, "
'    strSQL = strSQL & "        CANYT_NOTE "
'
'    strSQL = strSQL & "   FROM CANY_TRAN "
'
'    strSQL = strSQL & "  WHERE CANYT_ACPTNO = '" & strAcptNo & "' "
'
'    MakeHeaderSql2 = strSQL
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :SetHeaderForm
'        機能             :ヘッダ情報画面表示
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetHeaderForm(rsHeader As Recordset)
    Dim vRet     As Variant

    With SubHeader.Form
     '契約情報タブ
        .txt_受注契約番号 = rsHeader.Fields("CARG_ACPTNO")
        .txt_初回契約日 = Format(rsHeader.Fields("CARG_FSDATE"), "YYYY/MM/DD")
        .txt_契約開始日 = Format(rsHeader.Fields("CARG_STDATE"), "YYYY/MM/DD")
        .txt_契約満了日 = Format(rsHeader.Fields("CARG_EDDATE"), "YYYY/MM/DD")
        .txt_解約日付 = Format(rsHeader.Fields("CARG_KYDATE"), "YYYY/MM/DD")
'▼ヤード解約情報保存時に使用▼
'        ' ダミーとして解約日付のタグにコンテナ契約ファイル.解約日付を設定
'        .txt_解約日付.Tag = Format(rsHeader.Fields("CARG_KYDATE"), "YYYY/MM/DD")
        .cmb_解約区分 = rsHeader.Fields("CARG_KAICD")
        .cmb_保証区分 = Format(rsHeader.Fields("CARG_HOSYICD"), "0")
        .cmb_保証会社コード = Format(rsHeader.Fields("CARG_HOSYO_CD"), "0")
        .txt_預り保証金 = rsHeader.Fields("CARG_SECUKG")
        .txt_承認日 = Format(rsHeader.Fields("CARG_HOSYD"), "YYYY/MM/DD")
        .txt_月額料金 = rsHeader.Fields("CARG_RENTKG")
        .txt_消費税 = rsHeader.Fields("CARG_SYOZEI")
        .txt_合計 = rsHeader.Fields("CARG_RENTKG") + rsHeader.Fields("CARG_SYOZEI")
        .txt_契約備考 = rsHeader.Fields("CARG_BIKO")
        '解約受付日取得
        If Nz(rsHeader.Fields("CARG_KAIUKEDATE"), "") = "" Then
            .txt_解約受付日 = Null
        Else
            .txt_解約受付日 = Format(rsHeader.Fields("CARG_KAIUKEDATE"), "YYYY/MM/DD")
        End If
            
        '解約区分名取得
        Call cmb_解約区分_AfterUpdate

        '保証区分名取得
        Call cmb_保証区分_AfterUpdate
        
        '保証会社名取得
        Call cmb_保証会社コード_AfterUpdate

    '顧客情報タブ
        .txt_顧客名 = rsHeader.Fields("USER_NAME")
        .txt_顧客名カナ = rsHeader.Fields("USER_KANA")
        .txt_代表者名 = rsHeader.Fields("USER_TANM")
        .txt_代表者名カナ = rsHeader.Fields("USER_TAKA")
        .txt_郵便番号 = rsHeader.Fields("USER_YUBINO")
        .txt_住所1 = rsHeader.Fields("USER_ADR_1")
        .txt_住所2 = rsHeader.Fields("USER_ADR_2")
        .txt_住所3 = rsHeader.Fields("USER_ADR_3")
        .txt_電話番号 = rsHeader.Fields("USER_TEL")
        .txt_携帯番号 = rsHeader.Fields("USER_KEITAI")
        .txt_FAX番号 = rsHeader.Fields("USER_FAX")
        .txt_メールアドレス = rsHeader.Fields("USER_MAIL")
        .cmb_顧客区分 = rsHeader.Fields("USER_KKBN")
        .cmb_集金区分 = rsHeader.Fields("USER_SKBN")
        .txt_締日 = rsHeader.Fields("USER_SDAY")
        .txt_顧客備考 = rsHeader.Fields("USER_BIKO")
        
        '顧客区分名取得
        vRet = fncSelectRyak("020", .cmb_顧客区分)
        .txt_顧客区分名 = vRet

        '集金区分名取得
        vRet = fncSelectRyak("030", .cmb_集金区分)
        .txt_集金区分名 = vRet
        
    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SetHeaderForm2
'        機能             :ヘッダ情報画面表示
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub SetHeaderForm2(rsHeader2 As Recordset)
'
'    With SubHeader.Form
'     '解約通知回答情報
'        '回答区分
'        .cmb_Kaito = rsHeader2.Fields("CANYT_KBN")
'        Call cmb_Kaito_AfterUpdate      '回答区分名取得
'
'        '移動区分
'        If Nz(rsHeader2.Fields("CANYT_MOVEKBN"), "") = "" Then
'            .cmb_MoveKbn = Null
'        Else
'        .cmb_MoveKbn = Format(rsHeader2.Fields("CANYT_MOVEKBN"), "00")
'        Call cmb_MoveKbn_AfterUpdate    '移動区分名取得
'        End If
'
'        '解約日
'        If Nz(rsHeader2.Fields("CANYT_YE_DATE"), "") = "" Then
'            .txt_解約日 = Null
'        Else
'            .txt_解約日 = Format(rsHeader2.Fields("CANYT_YE_DATE"), "YYYY/MM/DD")
'        End If
'
'        '備考
'        If Nz(rsHeader2.Fields("CANYT_NOTE"), "") = "" Then
'            .txt_備考 = Null
'        Else
'            .txt_備考 = rsHeader2.Fields("CANYT_NOTE")
'        End If
'    End With
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Kaito_AfterUpdate
'        機能             :cmb_Kaito_更新後処理
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_Kaito_AfterUpdate()
'
''パラメータ     0   使用可
''               1   使用不可
'    With SubHeader.Form
'        Select Case .cmb_Kaito.Column(0)
'            Case 0, 2 '未回答、移動
'                Call CmLockControl(.txt_解約日, 0)
'                Call CmLockControl(.cmb_MoveKbn, 0)
'            Case 1  '解約
'                Call CmLockControl(.txt_解約日, 0)
'                Call CmLockControl(.cmb_MoveKbn, 1)
'                Call CmLockControl(.txt_MoveKbn, 1)
'            Case 3  '他希望
'                Call CmLockControl(.txt_解約日, 1)
'                Call CmLockControl(.cmb_MoveKbn, 1)
'                Call CmLockControl(.txt_MoveKbn, 1)
'        End Select
'    End With
'End Sub

'==============================================================================*
'
'        MODULE_ID        :CmLockControl
'        機能             :入力区分によるコントロールのロックを設定
'        CREATE_DATE      :2009/04/30
'
'引数 ０  使用可
'     １  使用不可
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CmLockControl(ByVal ContNM As Object, ByVal iMode As Integer)
'
'    Dim objControl As Object
'
'    Set objControl = ContNM
'
'    If iMode = 0 Then
'        objControl.Enabled = True
'        objControl.Locked = False
'    Else
'        objControl.Enabled = False
'        objControl.Locked = True
'    End If
'
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_BeforeUpdate
'        機能             :cmb_MoveKbn_更新前処理
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_MoveKbn_BeforeUpdate(Cancel As Integer)
'
'    Dim varCode As Variant
'
'    If ControlCrear(Me.cmb_MoveKbn, Me.txt_MoveKbn) = False Then Exit Sub
'
'    varCode = Me.cmb_MoveKbn
'
'    '入力チェック
'    Select Case CmbCheck_Num(varCode, 2)
'        Case 0
'
'        Case 2
'            '桁数ｴﾗｰ
'            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
'            GoTo Exception
'
'        Case 9
'            MsgBox "コードを入力してください。", vbInformation, P_ERROR
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exception
'    End Select
'
'    If Chk_NameBefore(Me.cmb_MoveKbn, "084") = "" Then
'       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
'       GoTo Exception
'    End If
'
'    Exit Sub
'
'Exception:
'    Cancel = True
'    SendKeys "{ESC}"
'
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_AfterUpdate
'        機能        　　 :cmb_MoveKbn更新後処理
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_MoveKbn_AfterUpdate()
'
'    On Error GoTo Exception
'
'    'ファイルへの存在チェック
'    If Chk_NameAfter(Me.cmb_MoveKbn, Me.txt_MoveKbn, "084") = False Then Exit Sub
'
'    Exit Sub
'
'Exception:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'
'End Sub
'
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_KeyPress
'        機能             :cmb_MoveKbn_キー入力時
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_MoveKbn_KeyPress(KeyAscii As Integer)
'    Call CmKeyPress_For_Num1(Me.cmb_MoveKbn, KeyAscii, 2)
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :MakeWork
'        機能             :ワークテーブル作成
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeWork() As Boolean
On Error GoTo err_rtn
    'FVS221_W01：サブフォームとリンクするテーブル
    
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
        strMsg = "紹介履歴が存在しません。"
        GoTo Exit_rtn
    End If
    
    MakeWork = True
    
    GoTo Exit_rtn
    
err_rtn:
    strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
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
'        CREATE_DATE      :2009/04/30
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
'        CREATE_DATE      :2009/04/30
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
'        CREATE_DATE      :2009/04/30
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
'        CREATE_DATE      :2009/04/30
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
    
    strSQL = strSQL & "  WHERE INTRT_YCODE = '" & strIntrYcd & "' "
    strSQL = strSQL & "  AND   INTRT_NO = '" & strIntrCno & "' "
    
    strSQL = strSQL & "  ORDER BY INTRT_INTRONO DESC "
    
    MakeIntrSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :AddWork
'        機能             :作業テーブル追加
'        CREATE_DATE      :2009/04/30
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
'        CREATE_DATE      :2009/04/30
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
'        CREATE_DATE      :2009/04/30
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
'        MODULE_NAME      :cmb_解約区分_AfterUpdate
'        機能             :cmb_解約区分更新処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約区分_AfterUpdate()

    Dim vCode       As Variant
    Dim strName     As String

    vCode = Nz(SubHeader.Form.cmb_解約区分, "")
    If vCode = "" Then
        SubHeader.Form.txt_解約区分 = ""
        Exit Sub
    End If
    SubHeader.Form.cmb_解約区分 = Format(vCode, "00")

    strName = fncSelectName("089", SubHeader.Form.cmb_解約区分)

    SubHeader.Form.txt_解約区分 = strName

    '  未使用キャンセル且つ契約開始日設定済み、解約日付未入力ならば
    '  解約日を契約開始日前月末日とする
    Dim strTargetDate  As String
    If SubHeader.Form.cmb_解約区分.VALUE = "03" And _
        Nz(SubHeader.Form.txt_契約開始日.VALUE, "") <> "" And _
        Nz(SubHeader.Form.txt_解約日付.VALUE, "") = "" Then
        
        strTargetDate = SubHeader.Form.txt_契約開始日.VALUE
        strTargetDate = DateSerial(Year(strTargetDate), _
                                   Month(strTargetDate), _
                                    1) - 1
        SubHeader.Form.txt_解約日付 = Format(strTargetDate, "YYYY/MM/DD")
        '解約受付日の初期設定
        If Nz(SubHeader.Form.txt_解約受付日, "") = "" Then
            SubHeader.Form.txt_解約受付日 = Format(Now, "YYYY/MM/DD")
        End If
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncSelectName
'        機能             :名称取得
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectName(stid As String, Optional ByVal ContNM As Object, _
                               Optional ByRef strCode As String) As String

    Dim objControl  As Object
    Dim vCode       As Variant
    Dim strID       As String
    Dim strWhere    As String
    Dim strName     As String

    fncSelectName = ""
    Set objControl = ContNM

    If Nz(strCode, "") = "" Then
        vCode = Nz(objControl, "")
    Else
        vCode = Nz(strCode, "")
    End If

    If vCode = "" Then Exit Function

    strID = stid

    ' 名称マスタより名称を取得
    strWhere = ""
    strWhere = strWhere & " NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & (vCode)
    strName = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")

    fncSelectName = strName
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_保証区分_AfterUpdate
'        機能             :cmb_保証区分更新処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_AfterUpdate()

    Dim strRyak     As String
    Dim vCode       As Variant

    vCode = Nz(SubHeader.Form.cmb_保証区分, "")

    If vCode = "" Then
        SubHeader.Form.txt_保証区分 = ""
        Exit Sub
    End If

    strRyak = fncSelectRyak("200", SubHeader.Form.cmb_保証区分)

    SubHeader.Form.txt_保証区分 = strRyak

    SubHeader.Form.lbl_預り保証金 = SubHeader.Form.cmb_保証区分.Column(3, SubHeader.Form.cmb_保証区分.ListIndex)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncSelectRyak
'        機能　　　       :保証区分、鍵区分の値を取得
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectRyak(stid As String, Optional ByVal ContNM As Object, _
                               Optional ByRef strCode As String) As String

    Dim objControl  As Object
    Dim vCode       As Variant
    Dim strID       As String
    Dim strWhere    As String
    Dim strRyak     As String

    fncSelectRyak = ""
    Set objControl = ContNM

    If Nz(strCode, "") = "" Then
        vCode = Nz(objControl, "")
    Else
        vCode = Nz(strCode, "")
    End If
    If vCode = "" Then Exit Function

    strID = stid

    ' 名称マスタより略称を取得
    strWhere = ""
    strWhere = strWhere & " NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & (vCode)
    strRyak = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWhere), "")

    fncSelectRyak = strRyak
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_保証会社コード_AfterUpdate
'        機能             :cmb_保証会社コード_更新処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証会社コード_AfterUpdate()
On Error GoTo ErrRtn

    Dim vCode    As Variant     '保証会社コード
    Dim strSQL      As String
    Dim objRs       As Recordset
    Dim blnError    As Boolean
    
    blnError = False
    vCode = Nz(SubHeader.Form.cmb_保証会社コード, "")

    If vCode = "" Then
        SubHeader.Form.txt_保証会社コード = ""
        Exit Sub
    End If

    '加瀬DB接続
    If fncConnectDB(pcintKASE_DB) = False Then Exit Sub

    strSQL = ""
    strSQL = strSQL & "SELECT TOP 1 SHIRM_SHIRN" & Chr(13)
    strSQL = strSQL & "  FROM SHIR_MAST" & Chr(13)
    strSQL = strSQL & " WHERE SHIRM_SHIRC = '" & vCode & "'"
    If SubHeader.Form.cmb_保証区分 = 1 Then
        strSQL = strSQL & Chr(13) & "   AND SHIRM_BUMOC = '" & strBumonCode & "'"
    End If
    
    Set objRs = pobjKASE_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
    
    If Not (objRs.BOF And objRs.EOF) Then
        SubHeader.Form.txt_保証会社コード = objRs.Fields("SHIRM_SHIRN")
    Else
        SubHeader.Form.txt_保証会社コード = ""
    End If
    
    objRs.Close
    Set objRs = Nothing

   '加瀬DB切断
    If fncDisConnectDB(pcintKASE_DB) = False Then GoTo ErrRtn
    
    GoTo ExitRtn
       
ErrRtn:
    blnError = True
    
ExitRtn:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKASE_DB Is Nothing Then pobjKASE_DB.Close: Set pobjKASE_DB = Nothing
   
    If blnError Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
   
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_Clickイベント
'        MODULE_ID        :cmd_保存_Click
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmd_保存_Click()
'
'    Dim intAns      As Integer
'    Dim intRtn      As Integer
'
'    '入力中のレコードを確定する。
'    On Error Resume Next
'    doCmd.RunCommand acCmdSaveRecord
'
'    On Error GoTo err_rtn
'
'    '確認メッセージ
'    intAns = MsgBox(strmsg保存確認, vbQuestion + vbYesNo, P_KAKUNIN)
'    If intAns = vbNo Then GoTo Exit_rtn
'
'    '回答区分をセット
'    strKaitoKbn = SubHeader.Form.cmb_Kaito.Column(0)
'    'DB更新
'    If WRI_RTN = True Then
'        doCmd.Echo False
'        Call CLR_RTN  ' 画面の初期化
'        SubHeader.Form.cmb_MoveKbn.Requery
'        doCmd.Echo True
'    End If
'
'Exit_rtn:
'    Exit Sub
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'    GoTo Exit_rtn
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :ＤＢの更新処理
'        MODULE_ID        :WRI_RTN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function WRI_RTN() As Boolean
'
'    Dim strErrFncName As String
'
'    On Error GoTo Update_Err
'
'    WRI_RTN = False
'
'    doCmd.Hourglass True
'    doCmd.SetWarnings False
'
'    strErrFncName = ""
'
'    Set WRI_ObjWS = DBEngine(0)
'    Set WRI_objDb = CurrentDb()
'
'    'トランザクション開始
'    WRI_ObjWS.BeginTrans
'
'    ' コンテナDB接続
'    If fncConnectDB(pcintKONT_DB) = False Then Exit Function
'
'
'    If strKaitoKbn = 1 Or 2 Then       '回答区分：解約、移動
'        ' コンテナ契約トラン更新
'        strErrFncName = "WRI_CARG_FILE"
'        If WRI_CARG_FILE = False Then
'            GoTo Update_Err
'        End If
'    End If
'
'    If strKaitoKbn = 1 Or 2 Or 3 Then  '回答区分：解約、移動、他希望
'        ' 予約受付トラン更新
'        strErrFncName = "WRI_YOUK_TRAN"
'        If WRI_YOUK_TRAN = False Then
'            GoTo Update_Err
'        End If
'
'        ' 予約ご紹介トラン更新
'        strErrFncName = "WRI_INTR_TRAN"
'        If WRI_INTR_TRAN = False Then
'            GoTo Update_Err
'        End If
'    End If
'
'    If strKaitoKbn = 2 Then            '回答区分：移動
'        '★移動先予約SQLで取得
'        strErrFncName = "GET_YOUK_TRAN"
'        If GET_YOUK_TRAN = False Then
'            GoTo Update_Err
'        End If
'        '●移動元受付SQL
'        strErrFncName = "GET_RCPT_TRAN"
'        If GET_RCPT_TRAN = False Then
'            GoTo Update_Err
'        End If
'        '▲コンテナマスタから取得
'        strErrFncName = "GET_CNTA_MAST"
'        If GET_CNTA_MAST = False Then
'            GoTo Update_Err
'        End If
'        'ヤードマスタ､コントロールマスタからヤード営業開始日を取得
'        strErrFncName = "GetYardBeginDay"
'        If GetYardBeginDay = False Then
'            GoTo Update_Err
'        End If
'        '■セキュリティカードマスタから取得
'        strErrFncName = "GET_SCRD_MAST"
'        If GET_SCRD_MAST = False Then
'            GoTo Update_Err
'        End If
'        '♪受注契約番号を取得、コンテナマスタ更新
'        strErrFncName = "fncAcptNewNo"
'        If fncAcptNewNo(strSqlserver, strAcptNo) = False Then
'            GoTo Update_Err
'        End If
'
'        ' 受付トラン登録
'        strErrFncName = "WRI_RCPT_TRAN"
'        If WRI_RCPT_TRAN = False Then
'            GoTo Update_Err
'        End If
'    End If
'
'    ' ヤード解約トラン登録              '回答区分：すべて
'    strErrFncName = "WRI_CANY_TRAN"
'    If WRI_CANY_TRAN = False Then
'        GoTo Update_Err
'    End If
'
'    'コミット完了
'    WRI_ObjWS.CommitTrans
'    WRI_RTN = True
'
'Exit_rtn:
'    doCmd.Hourglass False
'    doCmd.SetWarnings True
'    Set WRI_objDb = Nothing
'    Set WRI_ObjWS = Nothing
'
'
'    If Not pobjKONT_DB Is Nothing Then
'        ' コンテナDB切断
'        If fncDisConnectDB(pcintKONT_DB) = False Then Exit Function
'    End If
'
'    Exit Function
'
'Update_Err:
'    WRI_ObjWS.Rollback
'    MsgBox strmsg更新エラー & vbCrLf & _
'           Err.Number & Space(1) & Err.Description & vbCrLf & _
'           "(" & strErrFncName & ")", vbCritical, P_KAKUNIN
'    Err.Clear
'    GoTo Exit_rtn
'End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        Parameter        :intDBKbn
'                           　pcintKONT_DB：コンテナDB
'                             pcintKASE_DB：加瀬DB
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectDB(ByVal intDBKbn As Integer) As Boolean

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo fncConnectDB_Err

    fncConnectDB = False

    Select Case intDBKbn
        Case pcintKONT_DB  ' *** コンテナDB ***
            ' -----------------------------------------------------------------------------------------------------------
            ' コンテナDB接続
            strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumonCode & "'"), "")

            If strDataSource = "" Then
                MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
                Set pobjKONT_DB = Nothing
                Exit Function
            Else
                Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumonCode))
            End If
            ' -----------------------------------------------------------------------------------------------------------

        Case pcintKASE_DB  ' *** 加瀬DB ***
            ' -----------------------------------------------------------------------------------------------------------
            ' 加瀬DB接続
            strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))

            If strDataSource = "" Then
                MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
                Set pobjKASE_DB = Nothing
                Exit Function
            End If

            Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
            ' -----------------------------------------------------------------------------------------------------------
    End Select

    fncConnectDB = True

    Exit Function

fncConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :fncDisConnectDB
'        Parameter        :intDBKbn
'                             0：コンテナDB
'                             1：加瀬DB
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDisConnectDB(ByVal intDBKbn As Integer) As Boolean

    On Error GoTo fncDisConnectDB_Err

    fncDisConnectDB = False

    Select Case intDBKbn
        Case 0  ' *** コンテナDB ***
            pobjKONT_DB.Close
            Set pobjKONT_DB = Nothing
    
        Case 1  ' *** 加瀬DB ***
            pobjKASE_DB.Close
            Set pobjKASE_DB = Nothing
    End Select

    fncDisConnectDB = True

    Exit Function

fncDisConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :画面初期化
'        MODULE_ID        :CLR_RTN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CLR_RTN()
'
'    '画面上のテキストボックスをクリアする
'    With SubHeader.Form
'        .cmb_Kaito = ""
'        .txt_解約日 = ""
'        .cmb_MoveKbn = ""
'        .txt_MoveKbn = ""
'        .txt_備考 = ""
'        .txt_解約日付 = .txt_解約日付.Tag   'コンテナ契約トランの解約日に戻す
'    End With
'
'End Sub
'
'==============================================================================*
'
'        MODULE_NAME      : ヤード解約番号自動発番処理
'        MODULE_ID        : fncGetAutoCanytNo
'        PARAMETER        : 引数なし
'        戻り値　　　　　 : 発番した割引番号
'        CREATE_DATE      : 2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncGetAutoCanytNo() As String
'
'    Dim objAdoDbConnection   As Object
'    Dim strSQL               As String
'    Dim rsData               As Object
'    Dim intAddValue          As Integer
'
'    On Error GoTo Exception
'
'    ' 初期値設定
'    intAddValue = 1
'
'    ' ヤード解約トランのヤード解約番号の最大値＋ｎを初期解約番号とする
'    strSQL = ""
'    strSQL = strSQL & "SELECT MAX(CANYT_NO) AS Cnt FROM CANY_TRAN "
'    Set rsData = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    If Nz(rsData.Fields("Cnt"), "") = "" Then
'        fncGetAutoCanytNo = "000001"
'    Else
'        fncGetAutoCanytNo = Format(CLng(rsData.Fields("Cnt")) + intAddValue, "000000")
'    End If
'
'    rsData.Close
'    Set rsData = Nothing
'
'Exception:
'    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
'    Exit Function
'
'End Function
'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約トラン更新
'        MODULE_ID        :WRI_CARG_FILE
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function WRI_CARG_FILE() As Boolean
'
'    Dim strSQL      As String
'    Dim objRsCARG   As Recordset
'
'    WRI_CARG_FILE = False
'
'    strSQL = ""
'    strSQL = "SELECT * FROM CARG_FILE "
'    strSQL = strSQL & "WHERE CARG_YCODE = " & strYcode
'    strSQL = strSQL & "  AND CARG_NO = " & strCno
'
'    Set objRsCARG = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    ' コンテナ契約トラン更新
'    If fncUPDCARG_TRAN(objRsCARG) = False Then GoTo WRI_CARG_FILE_Exit
'    WRI_CARG_FILE = True
'
'WRI_CARG_FILE_Exit:
'    If Not objRsCARG Is Nothing Then objRsCARG.Close: Set objRsCARG = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約トラン更新処理
'        MODULE_ID        :fncUPDCARG_TRAN
'        Parameter        :objRsCARG   = コンテナ契約トランレコードセット
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncUPDCARG_TRAN(ByRef objRsCARG As Recordset) As Boolean
'
'    fncUPDCARG_TRAN = False
'
'    Dim strYeDate As String
'
'    strYeDate = SubHeader.Form.txt_解約日
'
'    ' コンテナ契約トランを更新
'    With objRsCARG
'            .Edit
'            .Fields("CARG_AGRE") = 2                                '契約状態　02（解約予定）
'            .Fields("CARG_CYDATE") = strYeDate                      '解約予定日
'            .Fields("CARG_KYDATE") = strYeDate                      '解約日付
'            .Fields("CARG_KAIUKEDATE") = DATE                       '解約受付日
'            .Fields("CARG_KAITANTO") = "092"                        '解約担当者コード 092（固定コード）
'        Select Case strKaitoKbn
'            Case 1 '回答区分：解約
'                .Fields("CARG_KAICD") = 1                           '解約区分 1（通常）
'            Case 2 '回答区分：移動
'                .Fields("CARG_KAICD") = 4                           '解約区分 4（移動）
'                .Fields("CARG_KAIRIYUCD") = 7                       '解約理由区分コード 7（加瀬別ＢＯＸ移動）
'        End Select
'            .Fields("CARG_UPDAD") = Format$(DATE, "YYYYMMDD")       '更新日付
'            .Fields("CARG_UPDAJ") = Format$(Time, "hhmmss")         '更新時刻
'            .Fields("CARG_UPDPB") = PROG_ID                         '更新プログラムID
'            .Fields("CARG_UPDUB") = MSZZ000.LsGetUserName           '更新ユーザーID
'            .UPDATE
'    End With
'
'    fncUPDCARG_TRAN = True
'
'fncUPDCARG_TRAN_Exit:
'    Exit Function
'End Function

'==============================================================================*
'
'        MODULE_NAME      :予約受付トラン更新
'        MODULE_ID        :WRI_YOUK_TRAN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function WRI_YOUK_TRAN() As Boolean
'    Dim strSQL      As String
'    Dim objRsYOUK   As Recordset
'
'    WRI_YOUK_TRAN = False
'
'    strSQL = ""
'    strSQL = "SELECT * FROM YOUK_TRAN "
'    strSQL = strSQL & "WHERE YOUKT_YCODE = " & strIntrYcd
'    strSQL = strSQL & "AND YOUKT_NO = " & strIntrCno
'
'    Set objRsYOUK = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    ' 予約受付トランを更新
'    If fncUPDYOUK_TRAN(objRsYOUK) = False Then GoTo WRI_YOUK_TRAN_Exit
'
'    WRI_YOUK_TRAN = True
'
'WRI_YOUK_TRAN_Exit:
'    If Not objRsYOUK Is Nothing Then objRsYOUK.Close: Set objRsYOUK = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :予約受付トラン更新処理
'        MODULE_ID        :fncUPDYOUK_TRAN
'        Parameter        :objRsYOUK   = 予約受付トランレコードセット
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncUPDYOUK_TRAN(ByRef objRsYOUK As Recordset) As Boolean
'
'    fncUPDYOUK_TRAN = False
'
'    ' 予約受付トランを更新
'    With objRsYOUK
'            .Edit
'        Select Case strKaitoKbn
'            Case 1, 3   '回答区分：解約､他希望
'                .Fields("YOUKT_YUKBN") = 52                     ' 予約受付状態区分　52(取置きキャンセル)
'            Case 2      '回答区分：移動
'                .Fields("YOUKT_YUKBN") = 10                     ' 予約受付状態区分　10(受付)
'        End Select
'            .Fields("YOUKT_UPDAD") = Format$(DATE, "YYYYMMDD")  ' 更新日付
'            .Fields("YOUKT_UPDAJ") = Format$(Time, "hhmmss")    ' 更新時刻
'            .Fields("YOUKT_UPDPB") = PROG_ID                    ' 更新プログラムID
'            .Fields("YOUKT_UPDUB") = MSZZ000.LsGetUserName      ' 更新ユーザーID
'            .UPDATE
'    End With
'
'    fncUPDYOUK_TRAN = True
'
'fncUPDYOUK_TRAN_Exit:
'    Exit Function
'End Function

'==============================================================================*
'
'        MODULE_NAME      :予約ご紹介トラン更新
'        MODULE_ID        :WRI_INTR_TRAN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function WRI_INTR_TRAN() As Boolean
'    Dim strSQL      As String
'    Dim objRsINTR   As Recordset
'
'    WRI_INTR_TRAN = False
'
'    strSQL = ""
'    strSQL = "SELECT * FROM INTR_TRAN "
'    strSQL = strSQL & "WHERE INTRT_YCODE = " & strIntrYcd
'    strSQL = strSQL & "AND INTRT_NO = " & strIntrCno
'
'    Set objRsINTR = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    ' 予約ご紹介トラン更新
'    If fncUPDINTR_TRAN(objRsINTR) = False Then GoTo WRI_INTR_TRAN_Exit
'
'    WRI_INTR_TRAN = True
'
'WRI_INTR_TRAN_Exit:
'    If Not objRsINTR Is Nothing Then objRsINTR.Close: Set objRsINTR = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :予約ご紹介トラン更新処理
'        MODULE_ID        :fncUPDINTR_TRAN
'        Parameter        :objRsINTR   = 受付ご紹介トランレコードセット
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncUPDINTR_TRAN(ByRef objRsINTR As Recordset) As Boolean
'
'    fncUPDINTR_TRAN = False
'
'    ' 予約受付トランを更新
'    With objRsINTR
'            .Edit
'        Select Case strKaitoKbn
'            Case 1, 3   '回答区分：解約､他希望
'                .Fields("INTRT_INTROKBN") = 6                   ' ご紹介区分　6(キャンセル)
'            Case 2      '回答区分：移動
'                .Fields("INTRT_INTROKBN") = 2                   ' ご紹介区分　2(受付)
'        End Select
'            .Fields("INTRT_UPDAD") = Format$(DATE, "YYYYMMDD")  ' 更新日付
'            .Fields("INTRT_UPDAJ") = Format$(Time, "hhmmss")    ' 更新時刻
'            .Fields("INTRT_UPDPB") = PROG_ID                    ' 更新プログラムID
'            .Fields("INTRT_UPDUB") = MSZZ000.LsGetUserName      ' 更新ユーザーID
'            .UPDATE
'    End With
'
'    fncUPDINTR_TRAN = True
'
'fncUPDINTR_TRAN_Exit:
'    Exit Function
'End Function

'==============================================================================*
'
'        MODULE_NAME      :予約受付トラン項目取得
'        MODULE_ID        :GET_YOUK_TRAN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function GET_YOUK_TRAN() As Boolean
'    Dim strSQL      As String
'    Dim objRsYOUK   As Recordset
'
'    GET_YOUK_TRAN = False
'
'    strSQL = ""
'    strSQL = "SELECT YOUKT_UKNO, "
'    strSQL = strSQL & "YOUKT_BUTU "
'    strSQL = strSQL & "FROM YOUK_TRAN "
'    strSQL = strSQL & "WHERE YOUKT_YCODE = " & strIntrYcd
'    strSQL = strSQL & "AND YOUKT_NO = " & strIntrCno
'
'    Set objRsYOUK = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    '項目をセット
'    With IDOSAKI_RCPT
'        .RCPT_NO = objRsYOUK.Fields("YOUKT_UKNO")   '予約番号
'        .RCPT_BUTU = objRsYOUK.Fields("YOUKT_BUTU") '予定収納物
'    End With
'
'    GET_YOUK_TRAN = True
'
'GET_YOUK_TRAN_Exit:
'    If Not objRsYOUK Is Nothing Then objRsYOUK.Close: Set objRsYOUK = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :受付トラン項目取得
'        MODULE_ID        :GET_RCPT_TRAN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function GET_RCPT_TRAN() As Boolean
'    Dim strSQL      As String
'    Dim objRsRCPT   As Recordset
'
'    GET_RCPT_TRAN = False
'
'    strSQL = ""
'    strSQL = "SELECT RCPT_HOSYICD, "
'    strSQL = strSQL & "RCPT_HOSYO_CD, "
'    strSQL = strSQL & "RCPT_SEV_EXPIRE, "
'    strSQL = strSQL & "RCPT_SEV_EXPIRE_MONTH, "
'    strSQL = strSQL & "RCPT_SEV1N, "
'    strSQL = strSQL & "RCPT_SEV2N, "
'    strSQL = strSQL & "RCPT_SEV3N, "
'    strSQL = strSQL & "RCPT_ENDEN, "
'    strSQL = strSQL & "RCPT_MITU_CD, "
'    strSQL = strSQL & "RCPT_SEIKYU_CD, "
'    strSQL = strSQL & "RCPT_SECUKG, "
'    strSQL = strSQL & "RCPT_SEND_NAME, "
'    strSQL = strSQL & "RCPT_SEND_HONORIFIC, "
'    strSQL = strSQL & "RCPT_SEND_YUBINO, "
'    strSQL = strSQL & "RCPT_SEND_ADR_1, "
'    strSQL = strSQL & "RCPT_SEND_ADR_2, "
'    strSQL = strSQL & "RCPT_SEND_ADR_3, "
'    strSQL = strSQL & "RCPT_SEND_TEL "
'    strSQL = strSQL & "FROM RCPT_TRAN "
'    strSQL = strSQL & "WHERE RCPT_YCODE = " & strYcode
'    strSQL = strSQL & " AND RCPT_CNO = " & strCno
'
'    Set objRsRCPT = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    '項目をセット
'    With IDOSAKI_RCPT
'        .RCPT_HOSYICD = objRsRCPT.Fields("RCPT_HOSYICD")                    '保証区分
'        .RCPT_HOSYO_CD = Nz(objRsRCPT.Fields("RCPT_HOSYO_CD"), "")          '保証会社コード
'        .RCPT_SEV_EXPIRE = objRsRCPT.Fields("RCPT_SEV_EXPIRE")              'キャンペーン適用満了年月
'        .RCPT_SEV_EXPIRE_MONTH = objRsRCPT.Fields("RCPT_SEV_EXPIRE_MONTH")  'キャンペーン適用月数
'        .RCPT_SEV1N = objRsRCPT.Fields("RCPT_SEV1N")                        'サービス内容１
'        .RCPT_SEV2N = objRsRCPT.Fields("RCPT_SEV2N")                        'サービス内容２
'        .RCPT_SEV3N = objRsRCPT.Fields("RCPT_SEV3N")                        'サービス内容３
'        .RCPT_ENDEN = objRsRCPT.Fields("RCPT_ENDEN")                        'サービス期間
'        .RCPT_MITU_CD = Nz(objRsRCPT.Fields("RCPT_MITU_CD"), "")            '見積もり請求区分
'        .RCPT_SEIKYU_CD = objRsRCPT.Fields("RCPT_SEIKYU_CD")                '初期費用請求方法
'        .RCPT_SECUKG = objRsRCPT.Fields("RCPT_SECUKG")                      '保証金
'        .RCPT_SEND_NAME = objRsRCPT.Fields("RCPT_SEND_NAME")                '鍵送付先名
'        .RCPT_SEND_HONORIFIC = objRsRCPT.Fields("RCPT_SEND_HONORIFIC")      '鍵送付先敬称
'        .RCPT_SEND_YUBINO = objRsRCPT.Fields("RCPT_SEND_YUBINO")            '鍵送付先郵便番号
'        .RCPT_SEND_ADR_1 = objRsRCPT.Fields("RCPT_SEND_ADR_1")              '鍵送付先住所１
'        .RCPT_SEND_ADR_2 = objRsRCPT.Fields("RCPT_SEND_ADR_2")              '鍵送付先住所２
'        .RCPT_SEND_ADR_3 = objRsRCPT.Fields("RCPT_SEND_ADR_3")              '鍵送付先住所３
'        .RCPT_SEND_TEL = objRsRCPT.Fields("RCPT_SEND_TEL")                  '鍵送付先電話番号
'    End With
'
'    GET_RCPT_TRAN = True
'
'GET_YOUK_TRAN_Exit:
'    If Not objRsRCPT Is Nothing Then objRsRCPT.Close: Set objRsRCPT = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ項目取得
'        MODULE_ID        :GET_CNTA_MAST
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function GET_CNTA_MAST() As Boolean
'    Dim strSQL      As String
'    Dim objRsCNTA   As Recordset
'
'    GET_CNTA_MAST = False
'
'    strSQL = ""
'    strSQL = "SELECT CNTA_KEY_TYPE, "
'    strSQL = strSQL & "CNTA_KEY_NO, "
'    strSQL = strSQL & "CNTA_KEY_NUM "
'    strSQL = strSQL & "FROM CNTA_MAST "
'    strSQL = strSQL & "WHERE CNTA_CODE = " & strYcode
'    strSQL = strSQL & " AND CNTA_NO = " & strCno
'
'    Set objRsCNTA = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    '項目をセット
'    With IDOSAKI_RCPT
'        .RCPT_KAGIICD = objRsCNTA.Fields("CNTA_KEY_TYPE")               '鍵区分
'        .RCPT_KAGI_NO = Nz(objRsCNTA.Fields("CNTA_KEY_NO"), "")         '鍵番号
'        .RCPT_KAGI_LENDNUM = Nz(objRsCNTA.Fields("CNTA_KEY_NUM"), "")   '鍵貸し出し本数
'    End With
'
'    GET_CNTA_MAST = True
'
'GET_CNTA_MAST_Exit:
'    If Not objRsCNTA Is Nothing Then objRsCNTA.Close: Set objRsCNTA = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤード営業開始日取得
'        MODULE_ID        :GetYardBeginDay
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function GetYardBeginDay() As Boolean
'    Dim strSQL      As String
'    Dim objRs       As Recordset
'
'    GetYardBeginDay = False
'
'    strSQL = ""
'    strSQL = "SELECT YARD_BEGIN_DAY, "
'    strSQL = strSQL & "CONT_SERVICE_DD "
'    strSQL = strSQL & "FROM YARD_MAST, "
'    strSQL = strSQL & "CONT_MAST "
'    strSQL = strSQL & "WHERE YARD_CODE = " & strYcode
'
'    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    '項目をセット
'    strYBeginDay = objRs.Fields("YARD_BEGIN_DAY")
'    strCONTServDD = objRs.Fields("CONT_SERVICE_DD")
'
'    GetYardBeginDay = True
'
'GetYardBeginDay_Exit:
'    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'       MODULE_NAME     : 希望開始日/取置期限日の算出
'       MODULE_ID       : DateAddColumn
'       IN              : dUKDATE               受付日(I)
'       OUT             : 希望開始日/取置期限日(Date)
'       CREATE_DATE     : 2009/04/30
'       MEMO            : MNR210　Insert_RCPT_TRANを参考に作成
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function DateAddColumn(ByVal dUKDATE As Date) As Date
'
'    Dim dResult             As Date
'    On Error GoTo ErrorHandler
'
'    dResult = DateAdd("d", strCONTServDD, dUKDATE)
'    If dResult < strYBeginDay Then
'        '求めた希望開始日がヤードの営業開始日以前ならば営業開始日とする
'        dResult = strYBeginDay
'    End If
'    DateAddColumn = dResult
'Exit Function
'
'ErrorHandler:
'    Call Err.Raise(Err.Number, "DateAddColumn" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDailyRateKOMS
'        機能             :日割り金額を算出する（ＫＯＭＳバージョン）
'        IN               :①起算日 As String (yyyymmdd形式)
'                         :②金額   As Long
'                         :③日割り日数   As Integer 出力用
'        OUT              :日割り金額 As Long
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncDailyRateKOMS(ByVal dKisanDate As Date, ByVal lngKingaku As Long, ByRef intDays As Integer) As Long
'    Dim intLastOfMonth      As Integer
'    Dim dblTemp             As Double
'    On Error GoTo ErrorHandler
'
'    '起算日が一日の場合日割りなし
'    If Day(dKisanDate) = 1 Then
'        intDays = 0
'        fncDailyRateKOMS = lngKingaku
'    Else
'        '起算日の月の最終日
'        intLastOfMonth = Day(DateSerial(Year(dKisanDate), Month(dKisanDate) + 1, 0))
'
'        '日割り日数＝終算日－起算日＋１日（起算日も含めるため）
'        intDays = intLastOfMonth - Day(dKisanDate) + 1
'
'        '日割り日数÷起算日の月の最終日×金額  ※１０の位以下切捨て
'        dblTemp = intDays / intLastOfMonth * lngKingaku
'        fncDailyRateKOMS = Fix(dblTemp / 100) * 100
'    End If
'Exit Function
'
'ErrorHandler:                   '↓自分の関数名
'    Call Err.Raise(Err.Number, "fncDailyRateKOMS" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Function
'==============================================================================*
'
'        MODULE_NAME      :セキュリティカードマスタ項目取得
'        MODULE_ID        :GET_SCRD_MAST
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function GET_SCRD_MAST() As Boolean
'    Dim strSQL      As String
'    Dim objRsSCRD   As Recordset
'
'    GET_SCRD_MAST = False
'
'    strSQL = ""
'    strSQL = "SELECT ???? "
'    strSQL = strSQL & "FROM SCRD_MAST "
'    strSQL = strSQL & "WHERE SCRDM_YCODE = " & strYcode
'    strSQL = strSQL & "AND SCRDM_ROOM_NO = " & strCno
'
'    Set objRsSCRD = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    '項目をセット
'    With FVS221Type_RCPT.RCPT_SECURITY_CARD = objRsCNTA.Fields("RCPT_SECURITY_CARD") 'セキュリティーカード
'
'    GET_SCRD_MAST = True
'
'GET_SCRD_MAST_Exit:
'    If Not objRsSCRD Is Nothing Then objRsSCRD.Close: Set objRsSCRD = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :受付トラン登録
'        MODULE_ID        :WRI_RCPT_TRAN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function WRI_RCPT_TRAN() As Boolean
'    Dim strSQL      As String
'    Dim objRsRCPT   As Recordset
'
'    WRI_RCPT_TRAN = False
'
'    strSQL = ""
'    strSQL = "SELECT * FROM RCPT_TRAN "
'
'    Set objRsRCPT = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    ' 受付トラン登録
'    If fncINSRCPT_TRAN(objRsRCPT) = False Then GoTo WRI_RCPT_TRAN_Exit
'
'    WRI_RCPT_TRAN = True
'
'WRI_RCPT_TRAN_Exit:
'    If Not objRsRCPT Is Nothing Then objRsRCPT.Close: Set objRsRCPT = Nothing
'    Exit Function
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :受付トラン登録処理
'        MODULE_ID        :fncINSRCPT_TRAN
'        IN               :objRsRCPT   = 受付トランレコードセット
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncINSRCPT_TRAN(ByRef objRsRCPT As Recordset) As Boolean
'    Dim strReturnMessage    As String
'    Dim strYeDate           As String
'    Dim strUkeNote          As String
'    Dim intDummy            As Integer
'
'    fncINSRCPT_TRAN = False
'
'    'SQLサーバー接続
'    Set dbSqlServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
'
'    ' 受付トランを登録
'    With objRsRCPT
'            .AddNew
'            .Fields("RCPT_NO") = IDOSAKI_RCPT.RCPT_NO                                                                   '受付番号                   ★移動先予約SQLで取得
'            .Fields("RCPT_UKTANTO") = "092"                                                                             '受付担当（固定）
'            .Fields("RCPT_UKDATE") = DATE                                                                               '受付日
'            .Fields("RCPT_CYDATE") = IIf(SubHeader.Form.txt_解約日 = "", Null, SubHeader.Form.txt_解約日)               '解約予定日
'            .Fields("RCPT_UKNOTE") = strYcode & "-" & strCno & "から自社理由移動"                                       '受付備考
'            .Fields("RCPT_UCODE") = IDOSAKI_RCPT.RCPT_UCODE                                                             '顧客コード
'            .Fields("RCPT_SEND_CD") = 0                                                                                 '書類送付方法   0（固定）
'            .Fields("RCPT_BUTU") = IDOSAKI_RCPT.RCPT_BUTU                                                               '主な収納物                 ★移動先予約SQLで取得
'            .Fields("RCPT_YCODE") = strIntrYcd                                                                          'ヤードコード
'            .Fields("RCPT_CNO") = strIntrCno                                                                            'コンテナ番号
'            .Fields("RCPT_HOSYICD") = IDOSAKI_RCPT.RCPT_HOSYICD                                                         '保証区分                   ●移動元受付SQLで取得
'            .Fields("RCPT_HOSYO_CD") = Nz(IDOSAKI_RCPT.RCPT_HOSYO_CD, "")                                               '保証会社コード             ●移動元受付SQLで取得
'            .Fields("RCPT_SEV_EXPIRE") = Nz(IDOSAKI_RCPT.RCPT_SEV_EXPIRE, "")                                           'キャンペーン適用満了年月   ●移動元受付SQLで取得
'            .Fields("RCPT_SEV_EXPIRE_MONTH") = IDOSAKI_RCPT.RCPT_SEV_EXPIRE_MONTH                                       'キャンペーン適用月数       ●移動元受付SQLで取得
'            .Fields("RCPT_SEV1N") = IDOSAKI_RCPT.RCPT_SEV1N                                                             'サービス内容１             ●移動元受付SQLで取得
'            .Fields("RCPT_SEV2N") = IDOSAKI_RCPT.RCPT_SEV2N                                                             'サービス内容２             ●移動元受付SQLで取得
'            .Fields("RCPT_SEV3N") = IDOSAKI_RCPT.RCPT_SEV3N                                                             'サービス内容３             ●移動元受付SQLで取得
'            .Fields("RCPT_ENDEN") = IDOSAKI_RCPT.RCPT_ENDEN                                                             'サービス期間               ●移動元受付SQLで取得
'            .Fields("RCPT_MITU_CD") = IIf(IDOSAKI_RCPT.RCPT_MITU_CD = "", Null, IDOSAKI_RCPT.RCPT_MITU_CD)              '見積もり請求区分           ●移動元受付SQLで取得
'            .Fields("RCPT_SEIKYU_CD") = IDOSAKI_RCPT.RCPT_SEIKYU_CD                                                     '初期費用請求方法           ●移動元受付SQLで取得
'            .Fields("RCPT_STDATE") = DateAddColumn(.Fields("RCPT_UKDATE"))                                              '契約開始日                 ※受付日＋CONT_SERVICE_DD
'            '当月内解約の場合、起算日を１日にする                                                                       '起算日                     ※契約開始日　の１日
'            If Format(.Fields("RCPT_STDATE"), "yyyymm") = Format(.Fields("RCPT_CYDATE"), "yyyymm") Then
'                .Fields("RCPT_KISAN_DATE") = DateSerial(Year(.Fields("RCPT_STDATE")), Month(.Fields("RCPT_STDATE")), 1)
'            Else
'                .Fields("RCPT_KISAN_DATE") = .Fields("RCPT_STDATE")
'            End If
'            .Fields("RCPT_DC_ENABLE") = 0                                                                               '値引有効可否 0（固定）
'            .Fields("RCPT_RENTKG") = strRentkg                                                                          '月額使用料
'            .Fields("RCPT_FIRSTKG") = fncDailyRateKOMS(.Fields("RCPT_KISAN_DATE"), .Fields("RCPT_RENTKG"), intDummy)    '初回使用料                 ※起算日を元に日割り計算
'            .Fields("RCPT_SECUKG") = IDOSAKI_RCPT.RCPT_SECUKG                                                           '保証金                     ●移動元受付SQLで取得
'            .Fields("RCPT_SECUKG_FLG") = 1                                                                              '保証金修正フラグ  1（固定）
'            .Fields("RCPT_SUMIKG") = 0                                                                                  '未使用移動元入金済み金額 0（固定）
'            .Fields("RCPT_KAGIICD") = IDOSAKI_RCPT.RCPT_KAGIICD                                                         '鍵区分                     ▲１の情報でコンテナマスタから取得
'            .Fields("RCPT_KAGI_NO") = Nz(IDOSAKI_RCPT.RCPT_KAGI_NO, "")                                                 '鍵番号                     ▲１の情報でコンテナマスタから取得
''             .Fields("RCPT_SECURITY_CARD") =  設定必要                                                                 'セキュリティーカード       ※CmKomsModモジュールのGetRoomAllotSCardsをCallする
'             .Fields("RCPT_KAGI_LENDNUM") = IIf(IDOSAKI_RCPT.RCPT_KAGI_LENDNUM = "", Null, IDOSAKI_RCPT.RCPT_KAGI_LENDNUM) '鍵貸し出し本数          ▲１の情報でコンテナマスタから取得
'             .Fields("RCPT_SEND_NAME") = IDOSAKI_RCPT.RCPT_SEND_NAME                                                    '鍵送付先名                 ●移動元受付SQLで取得
'             .Fields("RCPT_SEND_HONORIFIC") = IDOSAKI_RCPT.RCPT_SEND_HONORIFIC                                          '鍵送付先敬称               ●移動元受付SQLで取得
'             .Fields("RCPT_SEND_YUBINO") = IDOSAKI_RCPT.RCPT_SEND_YUBINO                                                '鍵送付先郵便番号           ●移動元受付SQLで取得
'             .Fields("RCPT_SEND_ADR_1") = IDOSAKI_RCPT.RCPT_SEND_ADR_1                                                  '鍵送付先住所１             ●移動元受付SQLで取得
'             .Fields("RCPT_SEND_ADR_2") = IDOSAKI_RCPT.RCPT_SEND_ADR_2                                                  '鍵送付先住所２             ●移動元受付SQLで取得
'             .Fields("RCPT_SEND_ADR_3") = IDOSAKI_RCPT.RCPT_SEND_ADR_3                                                  '鍵送付先住所３             ●移動元受付SQLで取得
'             .Fields("RCPT_SEND_TEL") = IDOSAKI_RCPT.RCPT_SEND_TEL                                                      '鍵送付先電話番号           ●移動元受付SQLで取得
'             .Fields("RCPT_CARG_ACPTNO") = fncNewNo(.Fields("RCPT_STDATE"))                                             '受注契約番号               ※発番後、コントロールマスタにも登録？
'            .Fields("RCPT_FWARIBIKI") = 0                                                                               '初月割引額  0（固定）
'            .Fields("RCPT_EWARIBIKI") = 0                                                                               '毎月割引額  0（固定）
'            .Fields("RCPT_INSED") = Format$(DATE, "YYYYMMDD")                                                           '作成日付
'            .Fields("RCPT_INSEJ") = Format$(Time, "hhmmss")                                                             '作成時刻
'            .Fields("RCPT_INSPB") = PROG_ID                                                                             '作成プログラムID
'            .Fields("RCPT_INSUB") = MSZZ000.LsGetUserName                                                               '作成ユーザーID
'            .UPDATE
'    End With
'
'    fncINSRCPT_TRAN = True
'
'fncINSRCPT_TRAN_Exit:
'    If Not dbSqlServer Is Nothing Then dbSqlServer.Close: Set dbSqlServer = Nothing
'    Exit Function
'End Function
'
'==============================================================================*
'
'        MODULE_NAME      :ヤード解約トラン更新
'        MODULE_ID        :WRI_CANY_TRAN
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function WRI_CANY_TRAN() As Boolean
'    Dim strSQL      As String
'    Dim objRsCANY   As Recordset
'
'    WRI_CANY_TRAN = False
'
'    strSQL = ""
'    strSQL = "SELECT * FROM CANY_TRAN "
'
'    Set objRsCANY = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
'    ' ヤード解約トラン登録
'    If fncINSCANY_TRAN(objRsCANY) = False Then GoTo WRI_CANY_TRAN_Exit
'
'    WRI_CANY_TRAN = True
'
'WRI_CANY_TRAN_Exit:
'    If Not objRsCANY Is Nothing Then objRsCANY.Close: Set objRsCANY = Nothing
'    Exit Function
'
'End Function
'
'==============================================================================*
'
'        MODULE_NAME      :ヤード解約トラン登録処理
'        MODULE_ID        :fncINSCANY_TRAN
'        Parameter        :objRsCANY   = ヤード解約トランレコードセット
'        CREATE_DATE      :2009/04/30
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncINSCANY_TRAN(ByRef objRsCANY As Recordset) As Boolean
'
'    fncINSCANY_TRAN = False
'
'    Dim strAcptNo As String         '受注契約番号
'    Dim strYeDate As String         'ヤード解約予定日
'    Dim strKaikbn As String         '区分
'    Dim strIdokbn As String         '移動方法
'    Dim strNote As String           '備考
'
'    ' 部門コード取得
'    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
'    If strBumonCode = "" Then
'        Exit Function
'    End If
'
'    With SubHeader.Form
'        strAcptNo = .txt_受注契約番号
'        strYeDate = Nz(.txt_解約日, "")
'        strKaikbn = Nz(.cmb_Kaito.Column(0), "")    '区分
'        strIdokbn = Nz(.cmb_MoveKbn, "")            '移動方法
'        strNote = Nz(.txt_備考, "")
'    End With
'
'    ' ヤード解約トランにデータ登録
'    With objRsCANY
'            .AddNew
'            .Fields("CANYT_ACPTNO") = strAcptNo                     '受注受付番号
'            .Fields("CANYT_BUMOC") = strBumonCode                   '部門コード
'            .Fields("CANYT_YCODE") = strYcode                       'ヤードコード
'            .Fields("CANYT_CNTANO") = strCno                        'コンテナ番号
'            If strYeDate <> "" Then
'                .Fields("CANYT_YE_DATE") = strYeDate                'ヤード解約予定日
'            End If
'            If strKaikbn <> "" Then
'                .Fields("CANYT_KBN") = strKaikbn                    '区分
'            End If
'            If strIdokbn <> "" Then
'                .Fields("CANYT_MOVEKBN") = strIdokbn                '移動方法
'            End If
'            If strNote <> "" Then
'                .Fields("CANYT_NOTE") = strNote                     '備考
'            End If
'            .Fields("CANYT_INSED") = Format$(DATE, "yyyymmdd")      '作成日付
'            .Fields("CANYT_INSEJ") = Format$(Time, "hhmmss")        '作成時刻
'            .Fields("CANYT_INSPB") = PROG_ID                        '作成プログラムID
'            .Fields("CANYT_INSUB") = MSZZ000.LsGetUserName          '作成ユーザID
'            .UPDATE
'    End With
'
'    fncINSCANY_TRAN = True
'
'fncINSCANY_TRAN_Exit:
'    Exit Function
'End Function
'
'==============================================================================*
'
'        MODULE_NAME      :fncNewNo
'        機能             :受注契約番号発番
'        IN               :dKaisi   契約開始日
'        OUT              :契約番号(String)
'        CREATE_DATE      :2009/04/30
'        MEMO             :FVS400 fncNewNoを参考に作成
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncNewNo(ByVal dKaisi As Date) As String
'On Error GoTo ErrorHandler
'
'    Dim rsData      As Recordset
'    Dim lngKyno     As Long
'    Dim blnError    As Boolean
'    Dim blnUpdate   As Boolean
'    Dim strAcptNo   As String
'    blnError = False
'
'    fncNewNo = False
'
'    ' コントロールマスタより取得
'    Set rsData = pobjKONT_DB.OpenRecordset(fncMakeContMastSql(), dbOpenDynaset)
'
'    If rsData.EOF Then
'        Call MSZZ024_M10("OpenRecordset", "コントロールマスタを参照できませんでした。")
'    End If
'
'    blnUpdate = False
'    With rsData
'            '契約NO
'            lngKyno = Nz(.Fields("CONT_KYNO"), 0) + 1
'            strAcptNo = fncMakeAcptNo(lngKyno, dKaisi)
'            .Edit
'            .Fields("CONT_KYNO") = lngKyno         ' 契約No
'            .Fields("CONT_UPDATE") = Now           ' 更新日
'            .UPDATE
'        .Close
'    End With
'
'    Set rsData = Nothing
'
'    fncNewNo = strAcptNo
'
'    GoTo ExitRtn
'
'ErrorHandler:
'    blnError = True
'
'ExitRtn:
'    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
'    If blnError Then
'        Call Err.Raise(Err.Number, "fncNewNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
'
'End Function
'
'==============================================================================*
'
'        MODULE_NAME      :fncMakeAcptNo
'        機能             :受注契約番号作成
'        IN               :lngKyno　    契約番号
'        OUT              :受注契約番号
'        CREATE_DATE      :2009/04/30
'        MEMO             :FVS400 fncMakeAcptNoを参考に作成
'
'▼ヤード解約情報保存時に使用▼
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncMakeAcptNo(lngKyno As Long, ByVal dKaisi As Date) As String
'
'    Dim strYY       As String
'    Dim strMM       As String
'
'    '年取得
'    strYY = ""
'    Select Case Year(dKaisi)
'        Case "1998"
'            strYY = "&&"
'        Case "1999"
'            strYY = "**"
'        Case Else
'            strYY = Mid(Format(dKaisi, "yyyy"), 3)
'    End Select
'
'    '月取得
'    strMM = Hex(Month(dKaisi))
'
'    ' 戻り値｢受注契約番号｣セット
'    fncMakeAcptNo = strBumonCode + strYY & strMM & Trim(Format$(lngKyno, "000000"))
'
'End Function
'
'==============================================================================*
'
''        MODULE_NAME      :コントロールマスタ検索SQL作成
''        MODULE_ID        :fncMakeContMastSql
''        CREATE_DATE      :2009/04/30
''        MEMO             :FVS400 fncMakeContMastSqlを参考に作成
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncMakeContMastSql() As String
'
'    Dim strSQL      As String
'
'    strSQL = "SELECT * "
'    strSQL = strSQL & "  FROM CONT_MAST "
'    strSQL = strSQL & " WHERE CONT_KEY = " & P_CONT_KEY
'
'    fncMakeContMastSql = strSQL
'
'End Function
