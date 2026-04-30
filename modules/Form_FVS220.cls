VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード解約一覧サブ
'   プログラムＩＤ　：　FVS220_S01
'   作　成　日　　　：  2009/04/30
'   作　成　者　　　：
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2011/09/30
'   修　正　者　　　：  M.RYU
'   修　正　内　容　：  返送日を設定、FVS220画面に返送日を入力しヘルパーに表示
'                       自動は「書類作成日+21日」
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：  2011/11/15
'   修　正　者　　　：  M.RYU
'   修　正　内　容　：  紹介制限ヤードのみ対象のチェックボックスを追加
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：  2012/01/10
'   修　正　者　　　：  M.RYU
'   修　正　内　容　：  紹介ボタン動作修正⇒紹介sql修正
'   Ver             ：  0.3
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：  2012/01/12
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  紹介時に段も見るように修正
'   Ver             ：  0.4
'**********************************************

Option Compare Database
Option Explicit
'==============================================================================*
'   Private定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FVS220"

Private Const P_MSG_ERR             As String = "エラー"
Private Const P_MSG_KAKUNIN         As String = "確認"
Private Const P_パラメータ数        As Integer = 4

Private Const pcintKONT_DB          As Integer = 0  ' コンテナDB
Private Const pcintKASE_DB          As Integer = 1  ' 加瀬DB

Private Const pcintSHOKAI           As Integer = 0  ' 紹介
Private Const pcintPreView          As Integer = 1  ' ﾌﾟﾚﾋﾞｭｰ
Private Const pcintPrint            As Integer = 2  ' 印刷
'ヤード解約回答
Private Const pcintNoANS            As Integer = 0  ' 未回答
Private Const pcintCancel           As Integer = 1  ' 解約
Private Const pcintMove             As Integer = 2  ' 移動
Private Const pcintOther            As Integer = 3  ' 他希望

Private Const strMsg_err            As String = "エラー"

Private Const strYOUK_UKTANTO       As String = "887"       'ヤード解約時の移動担当は887とする
Private Const strYOUK_UKKBN         As String = "99"
Private Const P_INTROKBN_01         As String = "1"          '取置

Private Const P_該当無              As String = "該当無"
'ステータス
Private Const P_STS未紹介           As String = "未紹介"
Private Const P_STS取置             As String = "紹介中（取置）"
Private Const P_STS受付             As String = "受付"
Private Const P_STS契約             As String = "決定 (契約)"
Private Const P_STS他希望           As String = "他希望(取置キャンセル)"


'識別(受付番号採番用)
Private Const P_UKE_U               As String = "U"

Private Const P_WORK_TABLE          As String = "FVS220_W01"
'やーd解約一覧サブフォームレコードソース
Private Const P_SUBFORM_RECORD_SOURCE   As String = _
                            " SELECT * FROM " & P_WORK_TABLE & _
                            "  ORDER BY FVS220_W01.FVS220_MULT_CNT DESC , FVS220_W01.FVS220_MOVE_CNT DESC , FVS220_W01.FVS220_OLD_KDATE, CLng([FVS220_CNTANO]); "

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private strSqlserver                As String   '接続ＤＢ名
Private strConnect                  As String       '接続文字列
Private strBumonCode                As String       '部門コード
Private strYardEigyoEndDate         As String       'ヤード営業終了日
Private pobjKONT_DB                 As Database     ' コンテナDB
Private pobjKASE_DB                 As Database     ' 加瀬DB

Private pstrWhere                   As String        ' 検索条件
'呼び出し元画面ID
Public strFormFrom     As String
'解約紹介情報
Private Type FVS220Type_KAIYAKUSHOKAI
    FVS220_CUSTOMER         As String   '顧客ID
    FVS220_CUSTOMER_NAME    As String   '顧客名
    FVS220_USAGE_CD         As String   '用途コード
    FVS220_STEP_CD          As String   '段コード
    FVS220_SIZE             As String   'サイズ
    FVS220_RENTAL           As String   '月額使用料
    FVS220_YCODE            As String   'ヤードコード
    FVS220_CNTANO           As String   'コンテナNO
    FVS220_MOVE_YCODE       As String   '移動先ヤードコード
    FVS220_MOVE_CNTANO      As String   '移動先コンテナNO
    FVS220_MOVE_KIRO        As String   '移動先キロ
    FVS220_MOVE_STEP_CD     As String   '移動先段コード
    FVS220_MOVE_STEP        As String   '移動先段
    FVS220_MOVE_SIZE        As String   '移動先サイズ
    FVS220_MOVE_RENTAL      As String   '移動先月額使用料
    FVS220_ACPTNO           As String   '移動元契約NO
    FVS220_UKNO             As String   '移動元予約番号
End Type

'顧客情報
Private Type FVS220Type_KOKYAKU
    USER_KANA               As String     '顧客名カナ
    USER_YUBINO             As String     '郵便番号
    USER_ADR_1              As String     '住所１
    USER_ADR_2              As String     '住所２
    USER_ADR_3              As String     '住所３
    USER_TANM               As String     '代表社名
    USER_TAKA               As String     '代表者カナ
    USER_TEL                As String     '電話番号
    USER_FAX                As String     'FAX
    USER_KEITAI             As String     '携帯番号
    USER_MAIL               As String     'メール
    USER_KKBN               As String     '顧客区分
End Type
    
'顧客情報
Private Type FVS220Type_MOTOCARG
    RCPT_BUTU               As String       '収納物
    YOUKT_RENJFROM          As String       '連絡時間FROM
    YOUKT_RENJTO            As String       '連絡時間TO
    YOUKT_MAIL              As String       'メールアドレス
End Type
    
'==============================================================================*
'
'        MODULE_NAME      :chk_全チェック_Click
'        機能             :チェックボックスクリック
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_全チェック_Click()
    On Error Resume Next
    fncChkON (chk_全チェック.VALUE)
    '再読込
    Me.FVS220_S01.Requery
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Ucode_AfterUpdate
'        機能             :顧客選択後処理
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ucode_AfterUpdate()
    If Nz(cmb_Ucode, "") = "" Then
        txt_Uname = ""
    Else
        txt_Uname = cmb_Ucode.Column(1, cmb_Ucode.ListIndex)
    End If
End Sub

''==============================================================================*
''
''        MODULE_NAME      :cmd_プレビュー_Click
''        機能             :プレビュ処理
''        IN               :
''        OUT              :
''        作成日           :2009/04/30  新規作成
''        更新日           :0000/00/00
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmd_プレビュー_Click()
'    Dim intCnt As Integer
'    Dim wSHOKAI() As FVS220Type_KAIYAKUSHOKAI
'    On Error GoTo Exception
'    '選択チェック
'    If funcSelectChk(pcintPreView, intCnt, wSHOKAI) = True Then
'        '未選択であればエラーメッセージ
'        If intCnt = 0 Then MsgBox "印刷対象を指定してください", vbOKOnly, P_MSG_KAKUNIN: Exit Sub
'    Else
'        Exit Sub
'    End If
'    ' 確認メッセージ表示
'    If MsgBox("印刷します。よろしいですか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
'        Exit Sub
'    End If
'    '印刷処理
'    If fncPrintPreview(intCnt, wSHOKAI) = True Then
'        Me.FVS220_S01.Requery
'    End If
'    Exit Sub
'
'Exception:
'    Call MSZZ024_M00("cmd_印刷_Click", True)
'
'
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_印刷_Click
'        機能             :解約通知書印刷
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_印刷_Click()
    Dim intCnt As Integer
    Dim wSHOKAI() As FVS220Type_KAIYAKUSHOKAI
    On Error GoTo Exception
    '選択チェック
    If funcSelectChk(pcintPrint, intCnt, wSHOKAI) = True Then
        '未選択であればエラーメッセージ
        If intCnt = 0 Then MsgBox "印刷対象を指定してください", vbOKOnly, P_MSG_KAKUNIN: Exit Sub
    Else
        Exit Sub
    End If
    ' 確認メッセージ表示
    If MsgBox("印刷します。よろしいですか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    '印刷処理
    If fncPrintHelper(intCnt, wSHOKAI) = True Then
        Me.FVS220_S01.Requery
    End If
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_印刷_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_検索_Click
'        機能             :検索処理
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索_Click()
    Dim strSQL As String
    Dim strWhere As String
    '全チェッククリア
    chk_全チェック.VALUE = False
    fncChkON (False)
    
    strWhere = ""
    '条件指定
    If Me.txt_Kana <> "" Then
        '顧客カナ
        strWhere = "WHERE FVS220_CUSTOMER_KANA LIKE '*" & Me.txt_Kana & "*'"
    End If
    If Me.cmb_Ucode <> "" Then
        '顧客コード
        If strWhere = "" Then
            strWhere = "WHERE FVS220_CUSTOMER = '" & Me.cmb_Ucode & "'"
        Else
            strWhere = "AND FVS220_CUSTOMER = '" & Me.cmb_Ucode & "'"
        End If
    End If
    If Me.cmb_YukbnFrom <> "" Then
        '状況
        If strWhere = "" Then
            strWhere = "WHERE FVS220_SITUATION = '" & Me.cmb_YukbnFrom & "'"
        Else
            strWhere = "AND FVS220_SITUATION = '" & Me.cmb_YukbnFrom & "'"
        End If
    End If
    'strSQL = "SELECT FVS220_W01.FVS220_CUSTOMER, FVS220_W01.FVS220_CUSTOMER_NAME, FVS220_W01.FVS220_USAGE, FVS220_W01.FVS220_STEP, FVS220_W01.FVS220_SIZE, FVS220_W01.FVS220_RENTAL, FVS220_W01.FVS220_SITUATION, FVS220_W01.FVS220_MOVE_YCODE, FVS220_W01.FVS220_MOVE_CNTANO, FVS220_W01.FVS220_LUGGAGE, FVS220_W01.FVS220_CHECK, FVS220_W01.FVS220_YCODE, FVS220_W01.FVS220_CNTANO FROM FVS220_W01 "
    strSQL = "SELECT FVS220_W01.* FROM FVS220_W01 "
    strSQL = strSQL & strWhere
    strSQL = strSQL & "ORDER BY FVS220_W01.FVS220_MULT_CNT DESC , FVS220_W01.FVS220_MOVE_CNT DESC , FVS220_W01.FVS220_OLD_KDATE, CLng([FVS220_CNTANO]);"
    Me.FVS220_S01.Form.RecordSource = strSQL
    Me.FVS220_S01.Requery
    pstrWhere = strWhere
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        機能             :終了
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
  On Error GoTo Exception
  
    ' 確認メッセージ表示
    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If

    ' 本画面を閉じる
    doCmd.Close acForm, Me.FormName
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_終了_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_紹介_Click
'        機能             :紹介処理
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_紹介_Click()
    Dim intCnt As Integer
    Dim wSHOKAI() As FVS220Type_KAIYAKUSHOKAI
    On Error GoTo Exception
    '入力チェック
    If Me.txt_Syokaihani = "" Then MsgBox "紹介範囲を指定してください", vbOKOnly, P_MSG_KAKUNIN: Exit Sub
    '選択チェック
    If funcSelectChk(pcintSHOKAI, intCnt, wSHOKAI) = True Then
        '未選択であればエラーメッセージ
        If intCnt = 0 Then MsgBox "紹介対象を指定してください", vbOKOnly, P_MSG_KAKUNIN: Exit Sub
    Else
        Exit Sub
    End If
        
    ' 確認メッセージ表示
    If MsgBox("紹介します。よろしいですか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    '紹介処理
    If fncSyokai(intCnt, wSHOKAI) = True Then
        Me.FVS220_S01.Requery
    End If
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_紹介_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activateイベント
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
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
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode, Shift)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form Loadイベント
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    Dim strTemp     As String
    Dim strParam()  As String
    'パラメータ取得
    strTemp = Nz(Me.OpenArgs)
    
    strParam = Split(strTemp, ",")
    If UBound(strParam) <> P_パラメータ数 - 1 Then
        MsgBox "パラメータが正しくありません"
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    
    strFormFrom = strParam(0)
    Me.txt_ヤードコード = strParam(1) 'ヤードコード
    Me.txt_ヤード名 = strParam(2) 'ヤード名
    strYardEigyoEndDate = strParam(3) 'ヤード解約予定日
    ' 部門コード取得
    '接続文字列取得
    Call fncGetConnectString
    
    'pstrBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    Me.txt_Syokaihani = 2
    pstrWhere = ""      '検索条件初期化
    
    txt_Kana.SetFocus
    'フォームの初期化処理
    Call DEL_RTN       ' ワーク削除 & 作成
    'ワーク作成
    Call Make_JetWork
    Me.FVS220_S01.Requery
    
    Me.txt_Hensobi = Format(DateAdd("d", 21, Now), "yyyy/mm/dd")   '書類作成日+26日   'INSERT 2011/09/30 M.RYU
    Me.chk_Syokaitaisyo.VALUE = False
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form Openイベント
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    Me.FVS220_S01.Requery

End Sub
'==============================================================================*
'
'        MODULE_NAME      :fncGetConnectString
'        機能             :接続文字列取得
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetConnectString() As Boolean
    On Error GoTo ErrorHandler
    
    fncGetConnectString = False
    
    '部門コード取得
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("DLookup(""CONT_BUMOC"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", "テーブル[CONT_MAST]の設定不正。")
    End If
    
    'SQL-Server名
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumonCode & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("DLookup(""SETUT_SETUN"", ""SETU_TABL"", ""SETUT_SETUB='ODBC_DATA_SOURCE_NAME_" & strBumonCode & "')", "テーブル[SETU_TABL]の設定不正。")
    End If
    
    '接続文字列取得
    strConnect = MSZZ007_M00(strBumonCode)
    If strConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "テーブル[SETU_TABL]の設定不正。")
    End If

    fncGetConnectString = True
    
    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetConnectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :DEL_RTN
'        機能             :ワーク削除
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DEL_RTN()

    Dim objDb   As Database
    Dim strSQL  As String

    On Error Resume Next

    Set objDb = CurrentDb()

    ' 解約ヤード・契約者状況一覧
    strSQL = "DELETE FROM FVS220_W01"
    objDb.Execute strSQL, dbFailOnError


    objDb.Close
    Set objDb = Nothing
End Sub

''==============================================================================*
''
''        MODULE_NAME      :DB接続
''        MODULE_ID        :fncConnectDB
''        Parameter        :intDBKbn
''                          pcintKONT_DB：コンテナDB
''                          pcintKASE_DB：加瀬DB
''        作成日           :2009/04/30  新規作成
''        更新日           :0000/00/00
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncConnectDB(ByVal intDBKbn As Integer) As Boolean
'
''    Dim strBumoCD     As String  ' 2005/11/08 DEL T.SUZUKI
'
'    Dim strConnect    As String
'    Dim strDataSource As String
'
'    On Error GoTo fncConnectDB_Err
'
'    fncConnectDB = False
'
'    Select Case intDBKbn
'        Case pcintKONT_DB  ' *** コンテナDB ***
'            ' -----------------------------------------------------------------------------------------------------------
'            ' コンテナDB接続
'            strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumonCode & "'"), "")
'
'            If strDataSource = "" Then
'                MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
'                Set pobjKONT_DB = Nothing
'                Exit Function
'            Else
'                Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumonCode))
'            End If
'            ' -----------------------------------------------------------------------------------------------------------
'
'        Case pcintKASE_DB  ' *** 加瀬DB ***
'            ' -----------------------------------------------------------------------------------------------------------
'            ' 加瀬DB接続
'            strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
'
'            If strDataSource = "" Then
'                MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
'                Set pobjKASE_DB = Nothing
'                Exit Function
'            End If
'
'            Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
'            ' -----------------------------------------------------------------------------------------------------------
'    End Select
'
'    fncConnectDB = True
'
'    Exit Function
'
'fncConnectDB_Err:
'    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
'    Err.Clear
'End Function

''==============================================================================*
''
''        MODULE_NAME      :DB切断
''        MODULE_ID        :fncDisConnectDB
''        Parameter        :intDBKbn
''                                   0：コンテナDB
''                                   1：加瀬DB
''        作成日           :2009/04/30  新規作成
''        更新日           :0000/00/00
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncDisConnectDB(ByVal intDBKbn As Integer) As Boolean
'
'    On Error GoTo fncDisConnectDB_Err
'
'    fncDisConnectDB = False
'
'    Select Case intDBKbn
'        Case 0  ' *** コンテナDB ***
'            pobjKONT_DB.Close
'            Set pobjKONT_DB = Nothing
'
'        Case 1  ' *** 加瀬DB ***
'            pobjKASE_DB.Close
'            Set pobjKASE_DB = Nothing
'    End Select
'
'    fncDisConnectDB = True
'
'    Exit Function
'
'fncDisConnectDB_Err:
'    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
'    Err.Clear
'End Function

'==============================================================================*
'
'        MODULE_NAME      :レコード選択チェックON/OFF
'        MODULE_ID        :fncChkON
'        戻り値           :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'==============================================================================*
'
'        MODULE_NAME      :fncChkON
'        機能             :レコード選択チェックON/OFF
'        IN               :フラグ　ON/OFF
'        OUT              :fncChkON,Output,Boolean,True 正常,False エラー
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkON(ByVal intFlag As Integer) As Boolean
    Dim objDb   As Database
    Dim strSQL  As String
    Dim strFlag As String
    On Error GoTo fncChkON_Err

    fncChkON = False
    Set objDb = CurrentDb()
    If intFlag Then
        strFlag = "-1"  'ON
        ' 解約ヤード・契約者状況一覧　表示分のみチェック
        strSQL = "update FVS220_W01 SET FVS220_CHECK =" & strFlag & " " & pstrWhere
    Else
        strFlag = "0"   'OFF
        ' 解約ヤード・契約者状況一覧
        strSQL = "update FVS220_W01 SET FVS220_CHECK =" & strFlag
    End If
    objDb.Execute strSQL, dbFailOnError
    objDb.Close
    Set objDb = Nothing
    fncChkON = True

    Exit Function

fncChkON_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :funcSelectChk
'        機能             :レコード選択チェック件数
'        IN               :pMode    モード（印刷/プレビュー）
'                          pintCnt  レコード選択件数
'                          pSHOKAI  紹介データ
'        OUT              :funcSelectChk,Output,Boolean,True 正常,False エラー
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function funcSelectChk(pMode As Integer, _
                               pintCnt As Integer, _
                               pSHOKAI() As FVS220Type_KAIYAKUSHOKAI) As Boolean
    Dim objRs        As Recordset
    Dim objSubForm   As SubForm
    Dim intCnt       As Integer
    On Error GoTo funcSelectChk_Err
    '初期化
    intCnt = 0
    pintCnt = 0
    funcSelectChk = False
    Erase pSHOKAI                   '初期化
    Set objSubForm = Me.FVS220_S01
    
    doCmd.GoToRecord , , acFirst    '先頭へ移動
    'objSubForm.Requery
    Set objRs = objSubForm.Form.RecordsetClone
    objRs.MoveLast
    objRs.MoveFirst
    Do Until objRs.EOF
        If Nz(objRs.Fields("FVS220_CHECK"), False) = True Then
            Select Case pMode
            Case pcintSHOKAI    '紹介
                If Nz(objRs.Fields("FVS220_SITUATION"), "") <> P_STS未紹介 Then
                    MsgBox "既に紹介済みです｡ " & vbCrLf & "(" & Nz(objRs.Fields("FVS220_CUSTOMER"), "") & "-" & Nz(objRs.Fields("FVS220_CUSTOMER_NAME"), "") & ")", vbOKOnly, P_MSG_KAKUNIN
                    funcSelectChk = False
                    GoTo funcSelectChk_Exit
                End If
            Case pcintPreView, pcintPrint   '印刷､プレビュー
                If Nz(objRs.Fields("FVS220_SITUATION"), "") = P_STS未紹介 Then
                    MsgBox "未紹介です｡ 紹介を行ってください｡" & vbCrLf & "(" & Nz(objRs.Fields("FVS220_CUSTOMER"), "") & "-" & Nz(objRs.Fields("FVS220_CUSTOMER_NAME"), "") & ")", vbOKOnly, P_MSG_KAKUNIN
                    funcSelectChk = False
                    GoTo funcSelectChk_Exit
                End If
            End Select
            '選択された場合
            ReDim Preserve pSHOKAI(intCnt)
            pSHOKAI(intCnt).FVS220_CUSTOMER = Nz(objRs.Fields("FVS220_CUSTOMER"), "")
            pSHOKAI(intCnt).FVS220_CUSTOMER_NAME = Nz(objRs.Fields("FVS220_CUSTOMER_NAME"), "")
            pSHOKAI(intCnt).FVS220_USAGE_CD = Nz(objRs.Fields("FVS220_USAGE_CD"), "")
            pSHOKAI(intCnt).FVS220_STEP_CD = Nz(objRs.Fields("FVS220_STEP_CD"), "")
            pSHOKAI(intCnt).FVS220_SIZE = Nz(objRs.Fields("FVS220_SIZE"), "")
            pSHOKAI(intCnt).FVS220_RENTAL = Nz(objRs.Fields("FVS220_RENTAL"), "")
            pSHOKAI(intCnt).FVS220_CNTANO = Nz(objRs.Fields("FVS220_CNTANO"), "")
            pSHOKAI(intCnt).FVS220_YCODE = Nz(objRs.Fields("FVS220_YCODE"), "")
            pSHOKAI(intCnt).FVS220_MOVE_YCODE = Nz(objRs.Fields("FVS220_MOVE_YCODE"), "")
            pSHOKAI(intCnt).FVS220_MOVE_CNTANO = Nz(objRs.Fields("FVS220_MOVE_CNTANO"), "")
            pSHOKAI(intCnt).FVS220_ACPTNO = Nz(objRs.Fields("FVS220_ACPTNO"), "")
            pSHOKAI(intCnt).FVS220_UKNO = Nz(objRs.Fields("FVS220_UKNO"), "")
            intCnt = intCnt + 1
        End If
        objRs.MoveNext
    Loop
    objRs.Close
    Set objRs = Nothing

    Set objSubForm = Nothing
    pintCnt = intCnt
    funcSelectChk = True

funcSelectChk_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objSubForm Is Nothing Then Set objSubForm = Nothing
    Exit Function

funcSelectChk_Err:
    pintCnt = 0
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo funcSelectChk_Exit

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncSyokai
'        機能             :ヤード紹介処理
'        IN               :pintCnt  件数
'                          pSHOKAI()  紹介対象情報
'        OUT              :fncSyokai,Output,Boolean,True 正常,False エラー
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSyokai(pintCnt As Integer, _
                           pSHOKAI() As FVS220Type_KAIYAKUSHOKAI) As Boolean
    Dim strSQL          As String
    Dim dbSQLServer     As Database
    Dim objWs           As Workspace
    Dim objDb           As Database
    Dim strMsg          As String
    Dim vCode           As Variant  'ヤードコード
    Dim vKiro           As Variant  'キロ
    Dim dNowDate        As Date
    Dim intCnt          As Integer
    Dim intCnt2         As Integer
    Dim intSyokaiCnt    As Integer  '紹介件数
    Dim strWKSize       As String
    Dim strUser         As String
    Dim strUkeNo        As String
    Dim strErrMsg       As String
    Dim objRsKONT       As Recordset
    Dim objRsWk         As Recordset
    Dim objRsYOUK       As Recordset
    Dim wkUSER As FVS220Type_KOKYAKU    '顧客情報
    Dim wkMOTOCarg  As FVS220Type_MOTOCARG  '元契約情報
    Dim bDoblChk        As Boolean      '重複紹介チェックフラグ
    Dim wkDoblChk()     As String       '重複紹介チェック用
    On Error GoTo Err_fncSyokai
    Screen.MousePointer = 11            '砂時計
    fncSyokai = False
    strWKSize = ""                      'サイズ比較
    intSyokaiCnt = 0                    '紹介件数
    dNowDate = Now                      'システム日付
    '作成UserID取得
    strUser = LsGetUserName()
    
    vCode = Val(Nz(Me.txt_ヤードコード, ""))
    vKiro = Val(Nz(Me.txt_Syokaihani, "0"))
    If vCode = 0 Then Exit Function
    
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    'トランザクション開始
    Set objWs = DBEngine.Workspaces(0)
    
    ' コンテナDB接続
    'If fncConnectDB(pcintKONT_DB) = False Then Exit Function
    '対象ヤード、コンテナ検索
    strSQL = fncMakeFreeSearchSql(vCode, vKiro)
    Set objRsKONT = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)  ' 2005/11/08 ADD T.SUZUKI
    For intCnt = 0 To pintCnt - 1
        If strWKSize <> "" Then
            If strWKSize <> pSHOKAI(intCnt).FVS220_SIZE Then
                'サイズが異なれば頭から検索
                objRsKONT.MoveFirst
            End If
        End If
        Do Until objRsKONT.EOF
            'Debug.Print objRsKONT.Fields("NYAR_NCODE") & "-" & objRsKONT.Fields("CNTA_SIZE") & "-" & pSHOKAI(intCnt).FVS220_SIZE
            '' 20120113 M.HONDA Start
            '' If Val(objRsKONT.Fields("CNTA_SIZE")) = Val(pSHOKAI(intCnt).FVS220_SIZE) Then
            If Val(objRsKONT.Fields("CNTA_SIZE")) = Val(pSHOKAI(intCnt).FVS220_SIZE) And Val(objRsKONT.Fields("CNTA_STEP")) = Val(pSHOKAI(intCnt).FVS220_STEP_CD) Then
            '' 20120113 M.HONDA END
                '重複チェック
                bDoblChk = False
                If intSyokaiCnt > 0 Then
                    For intCnt2 = 0 To UBound(wkDoblChk)
                        If wkDoblChk(intCnt2) = objRsKONT.Fields("NYAR_NCODE") & "-" & objRsKONT.Fields("CNTA_NO") Then
                            bDoblChk = True
                            Exit For
                        End If
                    Next
                End If
                If bDoblChk = False Then
                    '対象データセット
                    strWKSize = pSHOKAI(intCnt).FVS220_SIZE
                    pSHOKAI(intCnt).FVS220_MOVE_YCODE = objRsKONT.Fields("NYAR_NCODE")
                    pSHOKAI(intCnt).FVS220_MOVE_CNTANO = objRsKONT.Fields("CNTA_NO")
                    pSHOKAI(intCnt).FVS220_MOVE_KIRO = Nz(objRsKONT.Fields("NYAR_KIRO"), 999999)
                    pSHOKAI(intCnt).FVS220_MOVE_STEP_CD = Nz(objRsKONT.Fields("CNTA_STEP"), "")
                    pSHOKAI(intCnt).FVS220_MOVE_STEP = Nz(objRsKONT.Fields("CNTA_STEP_NAME"), "")
                    pSHOKAI(intCnt).FVS220_MOVE_SIZE = Nz(objRsKONT.Fields("CNTA_SIZE"), 0)
                    pSHOKAI(intCnt).FVS220_MOVE_RENTAL = Nz(objRsKONT.Fields("PRIC_PRICE"), 0)
                    '選んだデータはヤードコンテナは選ばせない処理が必要。
                    ReDim Preserve wkDoblChk(intSyokaiCnt)
                    '重複チェック用キーセット
                    wkDoblChk(intSyokaiCnt) = objRsKONT.Fields("NYAR_NCODE") & "-" & objRsKONT.Fields("CNTA_NO")
                    intSyokaiCnt = intSyokaiCnt + 1
                    objRsKONT.MoveNext
                    Exit Do
                End If
            End If
            objRsKONT.MoveNext
        Loop
        If objRsKONT.EOF Then
            If pSHOKAI(intCnt).FVS220_MOVE_YCODE = "" Then
                pSHOKAI(intCnt).FVS220_MOVE_YCODE = P_該当無
                If objRsKONT.RecordCount <> 0 Then objRsKONT.MoveFirst
            End If
        End If
    Next
    objRsKONT.Close
    Set objRsKONT = Nothing
    '予約取置きデータ作成
    For intCnt = 0 To pintCnt - 1
        objWs.BeginTrans
        If pSHOKAI(intCnt).FVS220_MOVE_YCODE <> P_該当無 Then
            strSQL = "SELECT * FROM YOUK_TRAN "
            Set objRsYOUK = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
            With objRsYOUK
                '受付番号採番処理
                If fncUkeNo_Numbering(strUkeNo, dbSQLServer) = False Then GoTo Exit_fncSyokai
                '顧客情報取得
                If fncGet_UserInfo(dbSQLServer, pSHOKAI(intCnt).FVS220_CUSTOMER, wkUSER) = False Then GoTo Exit_fncSyokai
                '元契約情報取得
                If fncGet_FromCRAGInfo(dbSQLServer, pSHOKAI(intCnt).FVS220_UKNO, wkMOTOCarg) = False Then GoTo Exit_fncSyokai
                .AddNew
                .Fields("YOUKT_UKNO") = strUkeNo                                '受付番号(var)
                .Fields("YOUKT_UKDATE") = dNowDate                              '受付日(dat)
                .Fields("YOUKT_UKTANTO") = strYOUK_UKTANTO                      '受付担当者コード(var)
                .Fields("YOUKT_UKKBN") = strYOUK_UKKBN                          '受付区分(num)
                .Fields("YOUKT_YCODE") = pSHOKAI(intCnt).FVS220_MOVE_YCODE      'ﾔｰﾄﾞｺｰﾄﾞ(num)
                .Fields("YOUKT_NEARKBN") = 0                                    '近隣ﾔｰﾄﾞ(num)※ﾌｪｰｽﾞ1非対応項目
                .Fields("YOUKT_NO") = pSHOKAI(intCnt).FVS220_MOVE_CNTANO        'ｺﾝﾃﾅ番号
                .Fields("YOUKT_YUKBN") = 2                                      '予約受付状態区分
                .Fields("YOUKT_KUDATE") = dNowDate                              '契約受付日
                .Fields("YOUKT_TKDATE") = strYardEigyoEndDate                   '取り置き期限日
                .Fields("YOUKT_BUTU") = wkMOTOCarg.RCPT_BUTU                    '予定収納物(var/Null=Y)
    '            .Fields("YOUKT_BIKO") = Nz(Me.txt_Biko, Null)                   '備考(var/Null=Y)
                .Fields("YOUKT_KIBOUSU") = 1                                     '希望本数(num/Null=Y)
                .Fields("YOUKT_UCODE") = pSHOKAI(intCnt).FVS220_CUSTOMER         '顧客ｺｰﾄﾞ(num/Null=Y)
                .Fields("YOUKT_NAME") = pSHOKAI(intCnt).FVS220_CUSTOMER_NAME     '顧客名(var)
                .Fields("YOUKT_KANA") = wkUSER.USER_KANA                         '顧客名ｶﾅ(var)
                .Fields("YOUKT_TANM") = wkUSER.USER_TANM                         '代表者名(var/Null=Y)
                .Fields("YOUKT_TAKA") = wkUSER.USER_TAKA                         '代表者名ｶﾅ(var/Null=Y)
                .Fields("YOUKT_YUBINO") = wkUSER.USER_YUBINO                     '郵便番号(var/Null=Y)
                .Fields("YOUKT_ADR_1") = wkUSER.USER_ADR_1                       '住所1(var/Null=Y)
                .Fields("YOUKT_ADR_2") = wkUSER.USER_ADR_2                       '住所2(var/Null=Y)
                .Fields("YOUKT_ADR_3") = wkUSER.USER_ADR_3                       '住所3(var/Null=Y)
                .Fields("YOUKT_TEL") = wkUSER.USER_TEL                           '電話番号(var/Null=Y)
                .Fields("YOUKT_KEITAI") = wkUSER.USER_KEITAI                    '携帯番号(var/Null=Y)
                .Fields("YOUKT_FAX") = wkUSER.USER_FAX                          'ＦＡＸ(var/Null=Y)
                .Fields("YOUKT_KKBN") = wkUSER.USER_KKBN                        '顧客区分(num)
                .Fields("YOUKT_RENKBN") = 0 'Val(Me.cmb_Renkbn)                 '連絡区分(num)
                '.Fields("YOUKT_RENJFROM") = wkMOTOCarg.YOUKT_RENJFROM           '連絡時間FROM(var/Null=Y)
                '.Fields("YOUKT_RENJTO") = wkMOTOCarg.YOUKT_RENJTO               '連絡時間TO(var/Null=Y)
                .Fields("YOUKT_MAIL") = wkUSER.USER_MAIL                        'メールアドレス(var/Null=Y) 顧客から取得
                .Fields("YOUKT_AUTOKBN") = 0                                    '自動区分
                .Fields("YOUKT_MOVEKBN") = 1                                    '移動区分
                .Fields("YOUKT_MOTO_ACPTNO") = pSHOKAI(intCnt).FVS220_ACPTNO    '元受注契約番号
                .Fields("YOUKT_REQ_YCODE") = pSHOKAI(intCnt).FVS220_MOVE_YCODE  '希望ヤードコード
                .Fields("YOUKT_MOVE_TANTO") = 0                                 '移動担当
                .Fields("YOUKT_GENKBN") = 0                                     '発生区分
                .Fields("YOUKT_ORG_YUKBN") = 0                                  '登録時予約状態
                .Fields("YOUKT_SEND_CD") = 0                                    '書類送付方法
                .Fields("YOUKT_USE_PERIOD") = 0                                 '利用予定期間
                .Fields("YOUKT_INSED") = Format$(DATE, "yyyymmdd")              '作成日付(var)
                .Fields("YOUKT_INSEJ") = Format$(time, "hhmmss")                '作成時刻(var)
                .Fields("YOUKT_INSPB") = PROG_ID                                '作成プログラムID(var)
                .Fields("YOUKT_INSUB") = strUser                                '作成ユーザーID(var)
                .UPDATE
                '紹介トラン
                If CmnInsertIntrTran(dbSQLServer, _
                                     strUkeNo, _
                                     CLng(pSHOKAI(intCnt).FVS220_MOVE_YCODE), _
                                     CLng(pSHOKAI(intCnt).FVS220_MOVE_CNTANO), _
                                     P_INTROKBN_01, strErrMsg) = False Then
                    Call Err.Raise(Err.Number, "fncUpdINTR_TRAN[" & strErrMsg & "]" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
                    GoTo Exit_fncSyokai
                End If
                .Close
            End With
            Set objRsYOUK = Nothing
        End If
        '対象ヤード、コンテナをローカルテーブルへ更新
        ' カレントDB接続
        Set objDb = CurrentDb()
        strSQL = "SELECT * FROM FVS220_W01 "
        strSQL = strSQL & "where FVS220_W01.FVS220_YCODE = """ & pSHOKAI(intCnt).FVS220_YCODE & """"
        strSQL = strSQL & " And FVS220_W01.FVS220_CNTANO = """ & pSHOKAI(intCnt).FVS220_CNTANO & """"
        Set objRsWk = objDb.OpenRecordset(strSQL)
        If Not objRsWk.EOF Then
            With objRsWk
                '該当レコード
                .Edit
                If pSHOKAI(intCnt).FVS220_MOVE_YCODE = P_該当無 Then
                    .Fields("FVS220_MOVE_YCODE") = pSHOKAI(intCnt).FVS220_MOVE_YCODE
                    .Fields("FVS220_MOVE_CNTANO") = ""
                    .Fields("FVS220_MOVE_SIZE") = ""
                    .Fields("FVS220_MOVE_RENTAL") = ""
                Else
                    .Fields("FVS220_SITUATION") = P_STS取置
                    .Fields("FVS220_MOVE_YCODE") = Format$(pSHOKAI(intCnt).FVS220_MOVE_YCODE, "000000")
                    .Fields("FVS220_MOVE_CNTANO") = Format$(pSHOKAI(intCnt).FVS220_MOVE_CNTANO, "000000")
                    .Fields("FVS220_MOVE_KYORI") = Format$(pSHOKAI(intCnt).FVS220_MOVE_KIRO, "0.00")
                    .Fields("FVS220_MOVE_STEP_CD") = pSHOKAI(intCnt).FVS220_MOVE_STEP_CD
                    .Fields("FVS220_MOVE_STEP") = pSHOKAI(intCnt).FVS220_MOVE_STEP
                    .Fields("FVS220_MOVE_SIZE") = Format$(pSHOKAI(intCnt).FVS220_MOVE_SIZE, "0.00")
                    '段が変更されたら、移動先の価格を表示。同じなら前の価格を維持させる
                    If .Fields("FVS220_STEP_CD") = pSHOKAI(intCnt).FVS220_MOVE_STEP_CD Then
                        .Fields("FVS220_MOVE_RENTAL") = .Fields("FVS220_RENTAL")
                    Else
                        .Fields("FVS220_MOVE_RENTAL") = Format$(pSHOKAI(intCnt).FVS220_MOVE_RENTAL, "###,###,##0")
                    End If
                End If
                .UPDATE
            End With
        End If
        objRsWk.Close
        Set objRsWk = Nothing
        objDb.Close
        Set objDb = Nothing
        'コミット
        objWs.CommitTrans
    Next
        
    fncSyokai = True

Exit_fncSyokai:
    If Not objRsYOUK Is Nothing Then objRsYOUK.Close: Set objRsYOUK = Nothing
    If Not objRsWk Is Nothing Then objRsWk.Close: Set objRsWk = Nothing
    If Not objRsKONT Is Nothing Then objRsKONT.Close: Set objRsKONT = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    If Not objWs Is Nothing Then Set objWs = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    Screen.MousePointer = 0
    Exit Function

Err_fncSyokai:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    'ロールバック
    If Not objWs Is Nothing Then
        objWs.Rollback
    End If
    Resume Exit_fncSyokai
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeFreeSearchSql
'        機能             :紹介検索SQL作成
'        IN               :
'        OUT              :SQL文
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeFreeSearchSql(pvCode As Variant, pvKiro As Variant) As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT NYAR_YCODE,ISNULL(NYAR_KIRO,999) as NYAR_KIRO ,NYAR_NCODE,CNTA_NO,CNTA_SYU, CNTA_STEP, STEP_NAME.NAME_NAME CNTA_STEP_NAME, CNTA_USAGE, "
    strSQL = strSQL & "       CASE CNTA_USAGE WHEN 0 THEN 10 WHEN 1 THEN 10 WHEN 3 THEN 30 WHEN 31 THEN"
    strSQL = strSQL & "        30 WHEN 32 THEN 30 ELSE NULL END AS CNTA_USAGE2, CNTA_SIZE, PRIC_PRICE "
    strSQL = strSQL & "  FROM NYAR_MAST"
    strSQL = strSQL & "       INNER JOIN CNTA_MAST ON (CNTA_CODE = NYAR_NCODE)"
    strSQL = strSQL & "        LEFT OUTER JOIN PRIC_TABL ON (PRIC_YCODE = CNTA_CODE AND PRIC_USAGE = CNTA_USAGE AND PRIC_SIZE = CNTA_SIZE AND  PRIC_STEP = CNTA_STEP)"
    strSQL = strSQL & "        LEFT OUTER JOIN NAME_MAST STEP_NAME ON STEP_NAME.NAME_ID = '090' AND STEP_NAME.NAME_CODE = CNTA_MAST.CNTA_STEP"
    strSQL = strSQL & " WHERE NYAR_YCODE = " & pvCode
    strSQL = strSQL & "   AND NYAR_NCODE <> " & pvCode
    
    'INSERT 2011/11/15 M.RYU----'紹介制限対象のみ----
    If Me.chk_Syokaitaisyo.VALUE = True Then
        strSQL = strSQL & " AND NYAR_NCODE IN (SELECT YARD_CODE  FROM YARD_MAST WHERE YARD_INLIMIT_YCODE= " & pvCode & ") "
    End If
    '------------------
    
    strSQL = strSQL & "   AND isnull(NYAR_KIRO,0)  <= " & pvKiro
    strSQL = strSQL & "   AND CNTA_USE = 1 "
    strSQL = strSQL & "   AND CNTA_USAGE IS NOT NULL "
    strSQL = strSQL & "   AND CNTA_SIZE IS NOT NULL "
    strSQL = strSQL & "   AND EXISTS"
    strSQL = strSQL & "           (SELECT *"
    strSQL = strSQL & "           FROM  YARD_MAST"
    strSQL = strSQL & "           WHERE YARD_CODE = CNTA_CODE "
    strSQL = strSQL & "             AND ISNULL(YARD_RENTEND_DAY, '9999/12/31') > GETDATE()"
    strSQL = strSQL & "             AND ISNULL(YARD_BEGIN_DAY, '1900/01/01')   <= GETDATE()"
'    strSQL = strSQL & "             AND ISNULL(YARD_INLIMIT_DAY, '1900/01/01') < GETDATE()) "  'DELETE 2012/01/10 M.RYU
    strSQL = strSQL & "             AND ISNULL(YARD_INLIMIT_DAY, '1900/01/01') >= GETDATE()) "  'INSERT 2012/01/10 M.RYU
    strSQL = strSQL & "   AND NOT EXISTS"
    strSQL = strSQL & "           (SELECT *"
    strSQL = strSQL & "           FROM  CARG_FILE"
    strSQL = strSQL & "           WHERE CARG_YCODE = CNTA_CODE "
    strSQL = strSQL & "             AND CARG_NO = CNTA_NO  "
    strSQL = strSQL & "             AND CARG_CONTNA = 0)"
    strSQL = strSQL & "   AND NOT EXISTS"
    strSQL = strSQL & "           (SELECT *"
    strSQL = strSQL & "           FROM  INTR_TRAN"
    strSQL = strSQL & "           WHERE INTRT_YCODE = CNTA_CODE "
    strSQL = strSQL & "             AND INTRT_NO = CNTA_NO "
    strSQL = strSQL & "             AND INTRT_INTROKBN IN (1, 2))"
    strSQL = strSQL & "ORDER BY ISNULL(NYAR_KIRO,999),NYAR_NCODE,CNTA_NO"
    fncMakeFreeSearchSql = strSQL
End Function
'==============================================================================*
'
'        MODULE_NAME      :UkeNo_Numbering
'        機能             :受付番号採番処理
'        IN               :strNo = 識別(1)+部門ｺｰﾄﾞ(1)+受付年(2)+受付月(1)+ｼｰｹﾝｽ(6)[11桁]
'        OUT              :取得結果
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUkeNo_Numbering(ByRef strNO As String, dbSQLServer As Database) As Boolean

    Dim strUkeNo    As String
    Dim strSEQ      As String
    Dim strYY       As String
    Dim strMM       As String
    Dim strSQL      As String
    Dim lngUke      As String
    Dim rsCon       As Recordset

    fncUkeNo_Numbering = False

    '識別
    strUkeNo = P_UKE_U

    '部門ｺｰﾄﾞ(ｺﾝﾄﾛｰﾙﾏｽﾀにある部門ｺｰﾄﾞを使用する)
    strSQL = "SELECT CONT_BUMOC, CONT_YOYA_NO, CONT_UPDATE FROM CONT_MAST "                 ' 2005/10/06 ADD T.SUZUKI
    Set rsCon = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)                            ' 2005/10/06 ADD T.SUZUKI
'    strSQL = "SELECT CONT_BUMOC,CONT_YOYA_NO,CONT_UPDATE FROM dbo_CONT_MAST "              ' 2005/10/06 DEL T.SUZUKI
'    Set rsCon = objDb.OpenRecordset(strSQL, dbOpenDynaset)                                 ' 2005/10/06 DEL T.SUZUKI

    If rsCon.EOF Then
        rsCon.Close
        Set rsCon = Nothing
        Exit Function
    Else
        With rsCon
            .Edit
            .Fields("CONT_YOYA_NO") = .Fields("CONT_YOYA_NO") + 1   '受付ｼｰｹﾝｽ№CountUp
            .Fields("CONT_UPDATE") = Now()                          '更新日
            strSEQ = CStr(Format(.Fields("CONT_YOYA_NO"), "000000")) '受付№保持
            strUkeNo = strUkeNo & .Fields("CONT_BUMOC")              '部門ｺｰﾄﾞｾｯﾄ
            .UPDATE
            .Close
        End With
        Set rsCon = Nothing
    End If

    '受付年
    Select Case Format(DATE, "yyyy")
        Case "1998": strYY = "&&"
        Case "1999": strYY = "**"
        Case Else
            If Mid$(Trim$(Format$(DATE, "yyyy")), 1, 2) = "20" Then
                strYY = Mid$(Trim$(Format$(DATE, "yyyy")), 3)
            End If
    End Select
    strUkeNo = strUkeNo & strYY

    '受付月
    Select Case Format(DATE, "m")
        Case "10": strMM = "A"
        Case "11": strMM = "B"
        Case "12": strMM = "C"
        Case Else
            strMM = Trim$(Format$(DATE, "m"))
    End Select
    strUkeNo = strUkeNo & strMM

    'ｼｰｹﾝｽ№ｾｯﾄ
    strUkeNo = strUkeNo & strSEQ

    strNO = strUkeNo

    GoTo UkeNo_Numbering_End

UkeNo_Numbering_Err:
    Call Err.Raise(Err.Number, "UkeNo_Numbering" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)  ' 2006/04/07 DEL T.SUZUKI
    Exit Function

UkeNo_Numbering_End:
    fncUkeNo_Numbering = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGet_UserInfo
'        機能             :顧客情報取得処理
'        IN               :strUCODE  顧客コード
'        OUT              :取得結果
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGet_UserInfo(dbSQLServer As Database, ByRef strUCODE As String, pUSER As FVS220Type_KOKYAKU) As Boolean
    Dim strSQL    As String
    Dim objrsKok           As Recordset
    On Error GoTo fncGet_UserInfo_Err
    
    strSQL = ""
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "USER_MAST "               ' 205/10/06 ADD T.SUZUKI
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "USER_CODE = '" & strUCODE & "'"
    strSQL = strSQL & ";"
    
    Set objrsKok = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If Not (objrsKok.BOF And objrsKok.EOF) Then
        With pUSER
            .USER_KANA = Nz(objrsKok.Fields("USER_KANA"), "")       '顧客名カナ
            .USER_YUBINO = Nz(objrsKok.Fields("USER_YUBINO"), "")    '郵便番号
            .USER_ADR_1 = Nz(objrsKok.Fields("USER_ADR_1"), "")      '住所１
            .USER_ADR_2 = Nz(objrsKok.Fields("USER_ADR_2"), "")      '住所２
            .USER_ADR_3 = Nz(objrsKok.Fields("USER_ADR_3"), "")      '住所３
            .USER_TANM = Nz(objrsKok.Fields("USER_TANM"), "")       '代表社名
            .USER_TAKA = Nz(objrsKok.Fields("USER_TAKA"), "")       '代表者カナ
            .USER_TEL = Nz(objrsKok.Fields("USER_TEL"), "")         '電話番号
            .USER_KEITAI = Nz(objrsKok.Fields("USER_KEITAI"), "")   '携帯番号
            .USER_FAX = Nz(objrsKok.Fields("USER_FAX"), "")         'FAX番号
            .USER_MAIL = Nz(objrsKok.Fields("USER_MAIL"), "")       'メール
            .USER_KKBN = Nz(objrsKok.Fields("USER_KKBN"), "")      '顧客区分
        End With
    End If
    objrsKok.Close
    Set objrsKok = Nothing
    fncGet_UserInfo = True
    Exit Function
fncGet_UserInfo_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    If Not objrsKok Is Null Then objrsKok.Close: Set objrsKok = Nothing
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGet_FromCRAGInfo
'        機能             :元契約情報取得処理
'        IN               :strUCODE  顧客コード
'        OUT              :取得結果
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGet_FromCRAGInfo(dbSQLServer As Database, ByRef strUkno As String, pCARG As FVS220Type_MOTOCARG) As Boolean
    Dim strSQL     As String
    Dim objRsCarg  As Recordset
    On Error GoTo fncGet_FromCRAGInfo_Err
    
    strSQL = ""
    strSQL = strSQL & "SELECT RCPT_BUTU , YOUKT_RENJFROM,YOUKT_RENJTO,YOUKT_MAIL "
    strSQL = strSQL & "FROM RCPT_TRAN "
    strSQL = strSQL & "INNER JOIN YOUK_TRAN ON YOUKT_UKNO = RCPT_NO "
    strSQL = strSQL & "where RCPT_NO = '" & strUkno & "' "
    strSQL = strSQL & ";"
    Set objRsCarg = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If Not (objRsCarg.BOF And objRsCarg.EOF) Then
        With pCARG
            .RCPT_BUTU = Nz(objRsCarg.Fields("RCPT_BUTU"), "")              '収納物
            .YOUKT_RENJFROM = Nz(objRsCarg.Fields("YOUKT_RENJFROM"), "")    '連絡時間FROM
            .YOUKT_RENJTO = Nz(objRsCarg.Fields("YOUKT_RENJTO"), "")        '連絡時間TO
            .YOUKT_MAIL = Nz(objRsCarg.Fields("YOUKT_MAIL"), "")            'メールアドレス
        End With
    End If
    objRsCarg.Close
    Set objRsCarg = Nothing
    fncGet_FromCRAGInfo = True
    Exit Function
fncGet_FromCRAGInfo_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    If Not objRsCarg Is Null Then objRsCarg.Close: Set objRsCarg = Nothing
End Function

'==============================================================================*
'
'        MODULE_NAME      :Make_JetWork
'        機能             :サブフォーム用のＷＯＲＫ作成
'        IN               :
'        OUT              :Make_JetWork,Output,Boolean,True 正常,False エラー
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_JetWork() As Boolean

    Dim strSQL          As String
    Dim objDb           As Database
    Dim objAdoDbConnection   As Object
    Dim strMsg          As String
    Dim vCode           As Variant
    Dim strKCODE        As String       '顧客コード
    Dim lngMltCnt       As Long         '複数契約件数
    Dim lngMoveCnt      As Long         '移動回数
    Dim dtKDate         As Date          '最古契約日
    Dim objRsKONT       As Object    ' コンテナレコードセット
    Dim objRsWk         As Recordset    ' MDBレコードセット
    Dim strCARG_ACPTNO  As String       '前レコード比較
    Dim strYOUKT_UKNO   As String       '前レコード比較
    
    On Error GoTo Err_Make_JetWork

    Make_JetWork = False
    strCARG_ACPTNO = ""
    strYOUKT_UKNO = ""
    vCode = Val(Nz(Me.txt_ヤードコード, ""))

'    If vCode = "" Then Exit Function
    If vCode = 0 Then Exit Function
    ' カレントDB接続
    Set objDb = CurrentDb()
    'Access側ワークテーブル削除
    Call ClearWork(objDb, P_WORK_TABLE)
    
    
    ' コンテナDB接続
    'If fncConnectDB(pcintKONT_DB) = False Then Exit Function
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(strBumonCode)

    '契約
    strSQL = fncMakeYARDListSql
    Set objRsKONT = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
    'Set objRsKONT = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
'    Set objRsKONT = objDb.OpenRecordset(strSQL, dbReadOnly)

    strSQL = ""
    strSQL = "SELECT * FROM FVS220_W01"
    Set objRsWk = objDb.OpenRecordset(strSQL)

    Do Until objRsKONT.EOF
        With objRsWk
            If strCARG_ACPTNO = Nz(objRsKONT.Fields("CARG_ACPTNO"), "") Then
                '同じ契約であれば作成しない
                GoTo next_rec
            Else
                .AddNew
            End If
            strCARG_ACPTNO = Nz(objRsKONT.Fields("CARG_ACPTNO"), "")
            strYOUKT_UKNO = Nz(objRsKONT.Fields("YOUKT_UKNO"), "")
            strKCODE = Format$(Nz(objRsKONT.Fields("CARG_UCODE"), 0), "000000")
            'ヤード解約ソート情報取得
            If fncGet_OrderInfo(objAdoDbConnection, strKCODE, _
                               lngMltCnt, _
                               lngMoveCnt, _
                               dtKDate) = False Then GoTo Err_Make_JetWork
            .Fields("FVS220_CUSTOMER") = Format$(Nz(objRsKONT.Fields("CARG_UCODE"), 0), "000000")           ' 顧客コード
            .Fields("FVS220_CUSTOMER_NAME") = objRsKONT.Fields("USER_NAME")                                 ' 顧客名
            .Fields("FVS220_CUSTOMER_KANA") = objRsKONT.Fields("USER_KANA")                                 ' 顧客名　カナ
            .Fields("FVS220_USAGE_CD") = objRsKONT.Fields("CNTA_USAGE")                                     ' 用途コード
            .Fields("FVS220_USAGE") = objRsKONT.Fields("YOTO_NAME")                                         ' 用途区分
            .Fields("FVS220_STEP_CD") = objRsKONT.Fields("CNTA_STEP")                                       ' 段コード
            .Fields("FVS220_STEP") = objRsKONT.Fields("STEP_NAME1")                                         ' 段
            .Fields("FVS220_SIZE") = Format$(Nz(objRsKONT.Fields("CNTA_SIZE1"), 0), "0.00")                 ' サイズ
            .Fields("FVS220_RENTAL") = Format$(Nz(objRsKONT.Fields("KINGAKU1"), 0), "###,###,##0")          ' 月額使用料
            .Fields("FVS220_KAIYAKU_DATE") = objRsKONT.Fields("CARG_KYDATE")                                ' 解約日
            .Fields("FVS220_STATUS") = Nz(objRsKONT.Fields("KYK_NAME"), "")                                 ' 契約状態
            '■Select Case Nz(objRsKONT.Fields("CANYT_KBN"), "")
            '■Case pcintCancel, pcintMove, pcintOther
            '■    .Fields("FVS220_SITUATION") = Nz(objRsKONT.Fields("YKK_NAME"), "")
            '■Case Else
                Select Case Nz(objRsKONT.Fields("YOUKT_YUKBN"), "")
                Case ""
                    .Fields("FVS220_SITUATION") = P_STS未紹介
                Case "2"
                    .Fields("FVS220_SITUATION") = P_STS取置
                Case "10"
                    .Fields("FVS220_SITUATION") = P_STS受付
                Case "20"
                    .Fields("FVS220_SITUATION") = P_STS契約
                Case "52"
                    '解約の場合のも考慮が必要■
                    .Fields("FVS220_SITUATION") = P_STS他希望
                Case "53"
                    .Fields("FVS220_SITUATION") = P_STS他希望
                End Select
            '■End Select
            .Fields("FVS220_MOVE_YCODE") = IIf(Nz(objRsKONT.Fields("YOUKT_YCODE"), "") = "", "", Format$(objRsKONT.Fields("YOUKT_YCODE"), "000000"))                  '移動先ヤード
            .Fields("FVS220_MOVE_CNTANO") = IIf(Nz(objRsKONT.Fields("YOUKT_MOVE_CNTANO"), "") = "", "", Format$(objRsKONT.Fields("YOUKT_MOVE_CNTANO"), "000000"))     '移動先コンテナ
            .Fields("FVS220_MOVE_KYORI") = IIf(Nz(objRsKONT.Fields("NYAR_KIRO"), "") = "", "", Format$(objRsKONT.Fields("NYAR_KIRO"), "0.00"))                        '移動先距離
            .Fields("FVS220_MOVE_STEP_CD") = Nz(objRsKONT.Fields("MOVE_STEP_CD"), "")                                                                                 '移動先段コード
            .Fields("FVS220_MOVE_STEP") = Nz(objRsKONT.Fields("MOVE_STEP_NAME"), "")                                                                                  '移動先段名
            .Fields("FVS220_MOVE_SIZE") = IIf(Nz(objRsKONT.Fields("MOVE_CNTA_SIZE"), "") = "", "", Format$(objRsKONT.Fields("MOVE_CNTA_SIZE"), "0.00"))               '移動先サイズ
            If objRsKONT.Fields("CNTA_STEP") = Nz(objRsKONT.Fields("MOVE_STEP_CD"), "") Then
                '段が同じ
                .Fields("FVS220_MOVE_RENTAL") = Format$(Nz(objRsKONT.Fields("KINGAKU1"), 0), "###,###,##0")          ' 月額使用料
            Else
                '段が異なる
                .Fields("FVS220_MOVE_RENTAL") = IIf(Nz(objRsKONT.Fields("PRIC_PRICE"), 0) = 0, "", Format$(objRsKONT.Fields("PRIC_PRICE"), "###,###,##0"))            '移動先コンテナ
            End If
            Select Case Nz(objRsKONT.Fields("YOUKT_YUKBN"), "")
            Case "2"    '取置
                .Fields("FVS220_LUGGAGE") = ""
            Case Else
                .Fields("FVS220_LUGGAGE") = Nz(objRsKONT.Fields("MOVETANTO_NAME"), "")
            End Select
            .Fields("FVS220_CHECK") = False
            .Fields("FVS220_CNTANO") = objRsKONT.Fields("CARG_NO")                                     ' 契約状態
            .Fields("FVS220_YCODE") = Me.txt_ヤードコード ''objRsKONT.Fields("CNTA_USAGE")             ' ヤードコード
            .Fields("FVS220_UKNO") = objRsKONT.Fields("CARG_UKNO")                                     ' 予約番号
            .Fields("FVS220_ACPTNO") = objRsKONT.Fields("CARG_ACPTNO")                                 ' 契約番号
            .Fields("FVS220_MULT_CNT") = lngMltCnt                                  ' 複数コンテナ契約件数
            .Fields("FVS220_MOVE_CNT") = lngMoveCnt                                 ' 移動回数
            .Fields("FVS220_OLD_KDATE") = dtKDate                                   ' 契約開始日
            .UPDATE
        End With
next_rec:
        objRsKONT.MoveNext
    Loop
    objRsKONT.Close
    Set objRsKONT = Nothing

    objRsWk.Close
    Set objRsWk = Nothing
    
    objDb.Close
    Set objDb = Nothing

    ' コンテナDB切断
    'If fncDisConnectDB(pcintKONT_DB) = False Then Exit Function
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    Make_JetWork = True

Exit_Make_JetWork:
    If Not objRsWk Is Nothing Then objRsWk.Close: Set objRsWk = Nothing
    If Not objRsKONT Is Nothing Then objRsKONT.Close: Set objRsKONT = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    'If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    Exit Function

Err_Make_JetWork:
    
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_JetWork
End Function

'==============================================================================*
'
'        MODULE_NAME      :ClearWork
'        機能             :ワークテーブルクリア
'        IN               :
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ClearWork(pobjDB As Database, strTable As String)

    Dim tdfNew      As TableDef
    
    FVS220_S01.Form.RecordSource = ""
    
    'ワークテーブル削除
    If TableExist(strTable, pobjDB) Then
        Call pobjDB.TableDefs.Delete(strTable)
    End If
    
    'ワークテーブル作成
    Set tdfNew = pobjDB.CreateTableDef(strTable)
    Call FieldAppend(tdfNew)
    Call pobjDB.TableDefs.Append(tdfNew)
    
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    
    FVS220_S01.Form.RecordSource = P_SUBFORM_RECORD_SOURCE
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :TableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :テーブル名
'        OUT              :True=存在する False=存在しない
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function TableExist(strTableName As String, pobjDB As Database) As Boolean

    Dim tdf As TableDef
    
    TableExist = False
    
    For Each tdf In pobjDB.TableDefs
        If tdf.NAME = strTableName Then
            TableExist = True
            Exit For
        End If
    Next tdf

End Function
'==============================================================================*
'
'        MODULE_NAME      :FieldAppend
'        機能             :ワークテーブル列作成
'        IN               :TableDef
'        OUT              :
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew

        Call .Fields.Append(.CreateField("FVS220_CUSTOMER", DataTypeEnum.dbText, 10))       '顧客ID
        Call .Fields.Append(.CreateField("FVS220_CUSTOMER_NAME", DataTypeEnum.dbText, 50))  '顧客名
        Call .Fields.Append(.CreateField("FVS220_CUSTOMER_KANA", DataTypeEnum.dbText, 50))  '顧客名_カナ
        Call .Fields.Append(.CreateField("FVS220_CNTANO", DataTypeEnum.dbText, 10))         'コンテナNO
        Call .Fields.Append(.CreateField("FVS220_YCODE", DataTypeEnum.dbText, 10))          'ヤードコード
        Call .Fields.Append(.CreateField("FVS220_UKNO", DataTypeEnum.dbText, 20))           '予約番号
        Call .Fields.Append(.CreateField("FVS220_ACPTNO", DataTypeEnum.dbText, 20))         '契約番号
        Call .Fields.Append(.CreateField("FVS220_USAGE_CD", DataTypeEnum.dbText, 5))        '用途コード
        Call .Fields.Append(.CreateField("FVS220_USAGE", DataTypeEnum.dbText, 50))          '用途
        Call .Fields.Append(.CreateField("FVS220_STEP_CD", DataTypeEnum.dbText, 5))         '段コード
        Call .Fields.Append(.CreateField("FVS220_STEP", DataTypeEnum.dbText, 10))           '段
        Call .Fields.Append(.CreateField("FVS220_SIZE", DataTypeEnum.dbText, 20))           'サイズ
        Call .Fields.Append(.CreateField("FVS220_RENTAL", DataTypeEnum.dbText, 50))         '月額使用料
        Call .Fields.Append(.CreateField("FVS220_KAIYAKU_DATE", DataTypeEnum.dbText, 20))   '解約予定日
        Call .Fields.Append(.CreateField("FVS220_STATUS", DataTypeEnum.dbText, 50))         '契約状態
        Call .Fields.Append(.CreateField("FVS220_SITUATION", DataTypeEnum.dbText, 50))      '状況
        Call .Fields.Append(.CreateField("FVS220_MOVE_YCODE", DataTypeEnum.dbText, 10))     '移動候補／移動先(ヤードコード)
        Call .Fields.Append(.CreateField("FVS220_MOVE_CNTANO", DataTypeEnum.dbText, 10))    '移動候補／移動先(コンテナＮＯ)
        Call .Fields.Append(.CreateField("FVS220_MOVE_KYORI", DataTypeEnum.dbText, 20))     '移動候補／距離
        Call .Fields.Append(.CreateField("FVS220_MOVE_STEP_CD", DataTypeEnum.dbText, 5))    '移動候補／段コード
        Call .Fields.Append(.CreateField("FVS220_MOVE_STEP", DataTypeEnum.dbText, 10))      '移動候補／段
        Call .Fields.Append(.CreateField("FVS220_MOVE_SIZE", DataTypeEnum.dbText, 20))      '移動候補／サイズ
        Call .Fields.Append(.CreateField("FVS220_MOVE_RENTAL", DataTypeEnum.dbText, 50))    '移動候補／月額使用料
        Call .Fields.Append(.CreateField("FVS220_LUGGAGE", DataTypeEnum.dbText, 50))        '荷物移動
        Call .Fields.Append(.CreateField("FVS220_CHECK", DataTypeEnum.dbBoolean))           '紹介チェック
        Call .Fields.Append(.CreateField("FVS220_MULT_CNT", DataTypeEnum.dbLong))           '複数契約件数
        Call .Fields.Append(.CreateField("FVS220_MOVE_CNT", DataTypeEnum.dbLong))           '移動回数
        Call .Fields.Append(.CreateField("FVS220_OLD_KDATE", DataTypeEnum.dbDate))          '最古契約日
        
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncMakeYARDListSql
'        機能             :ヤード解約一覧
'        IN               :
'        OUT              :SQL文
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeYARDListSql() As String
    Dim strSQL      As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT  CARG_STDATE "
    strSQL = strSQL & "       ,USER_TEL "
    strSQL = strSQL & "       ,USER_KEITAI "
    strSQL = strSQL & "       ,USER_FAX "
    strSQL = strSQL & "       ,CARG_UCODE "
    strSQL = strSQL & "       ,USER_NAME "
    strSQL = strSQL & "       ,USER_KANA "
    strSQL = strSQL & "       ,CARG_NO "
    strSQL = strSQL & "       ,CM1.CNTA_USAGE "
    strSQL = strSQL & "       ,YOTO_NAME.NAME_NAME YOTO_NAME "
    strSQL = strSQL & "       ,CM1.CNTA_SIZE CNTA_SIZE1 "
    strSQL = strSQL & "       ,CM1.CNTA_STEP"
    strSQL = strSQL & "       ,STEP_NAME1.NAME_NAME STEP_NAME1 "
    strSQL = strSQL & "       ,ISNULL(CARG_RENTKG,0) + ISNULL(CARG_SYOZEI,0) KINGAKU1 "
'    strSQL = strSQL & "       ,HOSYICD_NAME1.NAME_NAME HOSYICD_NAME1 "
    strSQL = strSQL & "       ,CASE WHEN CARG_HOSYICD IN(3, 4, 6) THEN 0 ELSE CARG_SECUKG END CARG_SECUKG "
    strSQL = strSQL & "       ,KKBN_NAME.NAME_NAME KKBN_NAME "
    strSQL = strSQL & "       ,SKBN_NAME.NAME_NAME SKBN_NAME "
    strSQL = strSQL & "       ,KYK_NAME.NAME_NAME KYK_NAME "
'    strSQL = strSQL & "       ,MISY.TANTM_TANTN    TANTM_TANTN       "
    strSQL = strSQL & "       ,CARG_KEY_RETDATE "
    strSQL = strSQL & "       ,CONVERT(VARCHAR,YOUKT_YCODE) YOUKT_YCODE "
    strSQL = strSQL & "       ,YARD_NAME "
    strSQL = strSQL & "       ,YOUKT_NO "
    strSQL = strSQL & "       ,CM2.CNTA_SIZE CNTA_SIZE2 "
    strSQL = strSQL & "       ,STEP_NAME2.NAME_NAME STEP_NAME2 "
    strSQL = strSQL & "       ,ISNULL(RCPT_RENTKG,0) + ISNULL(RCPT_EZAPPI,0) + ISNULL(RCPT_ADD_EZAPPI1,0) + ISNULL(RCPT_ADD_EZAPPI2,0) KINGAKU2 "
    strSQL = strSQL & "       ,HOSYICD_NAME2.NAME_NAME HOSYICD_NAME2 "
    strSQL = strSQL & "       ,MOVETANTO_NAME.NAME_NAME MOVETANTO_NAME "
    strSQL = strSQL & "       ,CASE WHEN RCPT_HOSYICD IN(3, 4, 6) THEN 0 ELSE RCPT_SECUKG END RCPT_SECUKG "
    strSQL = strSQL & "       ,CARG_ACPTNO "
    strSQL = strSQL & "       ,CARG_UKNO "
    strSQL = strSQL & "       ,YOUKT_YUKBN "            '移動先予約状態
    strSQL = strSQL & "       ,YOUK_TRAN.YOUKT_YCODE YOUKT_MOVE_YCODE "  '移動先ヤード
    strSQL = strSQL & "       ,YOUK_TRAN.YOUKT_NO YOUKT_MOVE_CNTANO "     '移動先コンテナNO
    strSQL = strSQL & "       ,CM2.CNTA_SIZE MOVE_CNTA_SIZE "
    strSQL = strSQL & "       ,CM2.CNTA_STEP MOVE_STEP_CD "
    strSQL = strSQL & "       ,STEP_NAME2.NAME_NAME MOVE_STEP_NAME "
    strSQL = strSQL & "       ,PRIC_PRICE "
    strSQL = strSQL & "       ,CARG_KYDATE "
    strSQL = strSQL & "       ,NYAR_KIRO "
    '■strSQL = strSQL & "       ,CANYT_KBN "
    '■strSQL = strSQL & "       ,YKK_NAME.NAME_NAME YKK_NAME "
    strSQL = strSQL & "       ,YOUKT_UKNO "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "       CARG_FILE "
    strSQL = strSQL & "       INNER JOIN CNTA_MAST CM1 ON "
    strSQL = strSQL & "           CM1.CNTA_CODE = CARG_YCODE AND CM1.CNTA_NO = CARG_NO "
    strSQL = strSQL & "       INNER JOIN USER_MAST ON "
    strSQL = strSQL & "           USER_CODE = CARG_UCODE "
'    strSQL = strSQL & "       LEFT OUTER JOIN  OPENROWSET('SQLOLEDB','KASE-SERVER';'SA';'','SELECT TANTM_TANTN, MISYT_BUMOC, MISYT_KOKYC  FROM TEST_KASE_DB.dbo.MISY_TRAN INNER JOIN TEST_KASE_DB.dbo.TANT_MAST ON MISYT_MISTC = TANTM_TANTC AND TANTM_BUMOC = ''L'' ')  MISY            ON REPLACE(STR(CARG_UCODE,6),' ','0') = MISYT_KOKYC AND MISYT_BUMOC = 'H'        "
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME1 ON "
    strSQL = strSQL & "           STEP_NAME1.NAME_ID = '090' AND STEP_NAME1.NAME_CODE = CM1.CNTA_STEP "
'    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST HOSYICD_NAME1 ON "
'    strSQL = strSQL & "           HOSYICD_NAME1.NAME_ID = '200' AND HOSYICD_NAME1.NAME_CODE = CARG_HOSYICD "
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST KKBN_NAME ON "
    strSQL = strSQL & "           KKBN_NAME.NAME_ID = '020' AND KKBN_NAME.NAME_CODE = USER_KKBN "
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST YOTO_NAME ON "
    strSQL = strSQL & "           YOTO_NAME.NAME_ID = '086' AND YOTO_NAME.NAME_CODE = CM1.CNTA_USAGE "
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST SKBN_NAME ON "
    strSQL = strSQL & "           SKBN_NAME.NAME_ID = '030' AND SKBN_NAME.NAME_CODE = USER_SKBN "
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST KYK_NAME ON "
    strSQL = strSQL & "           KYK_NAME.NAME_ID = '088' AND KYK_NAME.NAME_CODE = CARG_AGRE "
    '移動先予約情報
    strSQL = strSQL & "       LEFT OUTER JOIN (YOUK_TRAN "
    strSQL = strSQL & "       INNER JOIN YARD_MAST ON "
    strSQL = strSQL & "           YOUKT_YCODE = YARD_CODE "
    strSQL = strSQL & "           AND YOUKT_YUKBN IN( 2,10,20,52,53) "             '取置 受付 契約 取置キャンセル 受付キャンセル
    strSQL = strSQL & "       INNER JOIN CNTA_MAST CM2 ON "
    strSQL = strSQL & "           YOUKT_YCODE = CM2.CNTA_CODE AND YOUKT_NO = CM2.CNTA_NO "
    strSQL = strSQL & "       LEFT OUTER JOIN RCPT_TRAN ON "
    strSQL = strSQL & "           YOUKT_UKNO  = RCPT_NO "
    strSQL = strSQL & "       LEFT OUTER JOIN PRIC_TABL ON (PRIC_YCODE = CM2.CNTA_CODE AND PRIC_USAGE = CM2.CNTA_USAGE AND PRIC_SIZE = CM2.CNTA_SIZE AND  PRIC_STEP = CM2.CNTA_STEP) "
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME2 ON "
    strSQL = strSQL & "           STEP_NAME2.NAME_ID = '090' AND STEP_NAME2.NAME_CODE = CM2.CNTA_STEP "
    strSQL = strSQL & "       LEFT  OUTER JOIN NAME_MAST HOSYICD_NAME2 ON "
    strSQL = strSQL & "           HOSYICD_NAME2.NAME_ID = '200' AND HOSYICD_NAME2.NAME_CODE = RCPT_HOSYICD "
    strSQL = strSQL & "       LEFT  OUTER JOIN NAME_MAST MOVETANTO_NAME ON "
    strSQL = strSQL & "           MOVETANTO_NAME.NAME_ID = '084' AND MOVETANTO_NAME.NAME_CODE = YOUKT_MOVE_TANTO "
    strSQL = strSQL & "       ) ON YOUKT_MOTO_ACPTNO = CARG_ACPTNO "
    strSQL = strSQL & "       LEFT OUTER JOIN NYAR_MAST ON CARG_YCODE = NYAR_YCODE AND YOUKT_YCODE = NYAR_NCODE "
    '■strSQL = strSQL & "       LEFT OUTER JOIN (CANY_TRAN"
    '■strSQL = strSQL & "                    LEFT OUTER JOIN NAME_MAST YKK_NAME ON  YKK_NAME.NAME_ID = '161' AND YKK_NAME.NAME_CODE = CANYT_KBN) ON CANYT_ACPTNO = CARG_ACPTNO "

    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "      CARG_YCODE  = '" & Me.txt_ヤードコード & "' "
    strSQL = strSQL & "  AND CARG_AGRE   <> 9 "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "      USER_KANA ,YOUKT_UKNO DESC "
    
    fncMakeYARDListSql = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGet_OrderInfo
'        機能             :ヤード解約一覧 ソート順情報取得
'        IN               :objConnection:ｺﾝﾃﾅDB
'                          pstrCUSTOMER :顧客コード
'        OUT              :plngMltCnt   :複数ｺﾝﾃﾅ契約件数
'                          plngMoveCnt  :移動回数
'                          pdtKDATE     :契約開始日
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGet_OrderInfo(objConnection As Object, _
                                  ByRef pstrCUSTOMER As String, _
                                  plngMltCnt As Long, _
                                  plngMoveCnt As Long, _
                                  pdtKDATE As Date) As Boolean
    Dim strSQL       As String
    Dim rsWork       As Object
    fncGet_OrderInfo = False
On Error GoTo ErrorHandle
    '複数コンテナ契約件数
    strSQL = "SELECT count(*) as MLTCNT FROM CARG_FILE "
    strSQL = strSQL & "WHERE CARG_UCODE = '" & pstrCUSTOMER & "' and CARG_AGRE < 9 "
    Set rsWork = MSZZ025.ADODB_Recordset(strSQL, objConnection)
    If rsWork.EOF Then
        plngMltCnt = 0
    Else
        plngMltCnt = Nz(rsWork.Fields("MLTCNT"), 0)
    End If
    rsWork.Close
    Set rsWork = Nothing
    '移動回数
    strSQL = "select count(*) as MOVECNT FROM CARG_FILE "
    strSQL = strSQL & "where CARG_UCODE = '" & pstrCUSTOMER & "' and CARG_KAICD = 4 and CARG_AGRE = 9 "
    Set rsWork = MSZZ025.ADODB_Recordset(strSQL, objConnection)
    If rsWork.EOF Then
        plngMoveCnt = 0
    Else
        plngMoveCnt = Nz(rsWork.Fields("MOVECNT"), 0)
    End If
    rsWork.Close
    Set rsWork = Nothing
    '古い顧客検索(契約開始日の最も古い日付)取得
    strSQL = "select MIN(CARG_STDATE) as STDATE FROM CARG_FILE "
    strSQL = strSQL & "where CARG_UCODE = '" & pstrCUSTOMER & "'  "
    Set rsWork = MSZZ025.ADODB_Recordset(strSQL, objConnection)
    If rsWork.EOF Then
        pdtKDATE = Now
    Else
        pdtKDATE = Nz(rsWork.Fields("STDATE"), Now)
    End If
    rsWork.Close
    Set rsWork = Nothing
    fncGet_OrderInfo = True
    Exit Function
ErrorHandle:
    Call Err.Raise(Err.Number, "fncGetOrderInfo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    If Not rsWork Is Nothing Then rsWork.Close: Set rsWork = Nothing              ' 2005/10/24 ADD T.SUZUKI
End Function


''==============================================================================*
''
''        MODULE_NAME      :解約通知書プレビュー処理･･･NG
''        MODULE_ID        :fncPrintPreview
''        CREATE_DATE      :2009/04/03
''        戻り値           :fncPrintPreview,Output,Boolean,True 正常,False エラー
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncPrintPreview(pintCnt As Integer, pSHOKAI() As FVS220Type_KAIYAKUSHOKAI) As Boolean
'    Dim intCnt          As Integer
'    Dim intxlCount          As Integer 'ヤード解約ヘルパーのカウンタ
'    Dim intxl2Count         As Integer '地図,配置のカウンタ
'    Dim intCount            As Integer 'カウンタ
'    Dim strPath             As String 'ヤード解約ヘルパー
'    Dim strPath2            As String '地図,配置
'    Dim xlApp               As Object
'    Dim xlBook              As Object 'ヤード解約ヘルパー
'    Dim xlBook2             As Object 'ヤード解約ヘルパー2件目以降
'    Dim stCARG              As MSZZ099.Type_CARG_INF
'    Dim stCARG2             As MSZZ099.Type_CARG_INF2
'    Dim stCONT              As MSZZ099.Type_CONT_MAST
'    Dim varPrintSeets       As Variant
'    'ヘルパー情報取得
'    If MSZZ099.GetHelperFileInfo(strPath, strPath2) = False Then
'        GoTo Err_fncPrintHelper
'    End If
'
'    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
'    '2件目以降作成先ファイル
''    If Dir(strPath & "cp解約ヘルパー.xls") <> "" Then
''        Kill strPath & "cp解約ヘルパー.xls"
''    End If
''    FileCopy strPath & "解約ヘルパー.xls", strPath & "cp解約ヘルパー.xls"
'
'    ' Excelオブジェクトを生成する
'    Set xlApp = CreateObject("Excel.Application")
'    On Error GoTo ErrorHandler2
'    Set xlBook = xlApp.Workbooks.Open(strPath & "解約ヘルパー.xls")                              'コピー元
'    Set xlBook2 = xlApp.Workbooks.Add
'
'    For intCnt = 0 To pintCnt - 1
'        'Call subPrintHelperKyk(pSHOKAI(intCnt).FVS220_YCODE, pSHOKAI(intCnt).FVS220_CNTANO, pSHOKAI(intCnt).FVS220_MOVE_YCODE, pSHOKAI(intCnt).FVS220_MOVE_CNTANO)
'        '解約データ取得
'        Call MSZZ099.GetHelperDataKyk(strBumonCode, _
'                                      pSHOKAI(intCnt).FVS220_YCODE, _
'                                      pSHOKAI(intCnt).FVS220_CNTANO, _
'                                      pSHOKAI(intCnt).FVS220_MOVE_YCODE, _
'                                      pSHOKAI(intCnt).FVS220_MOVE_CNTANO, _
'                                      stCONT, _
'                                      stCARG, _
'                                      stCARG2)
''        If intCnt = 0 Then
'            '基本入力シートに値を設定
'            Call MSZZ099.SetBaseSheetKyk(strBumonCode, _
'                                         xlBook.Worksheets("基本入力"), _
'                                         stCARG, _
'                                         stCONT, _
'                                         stCARG2)
'            '出力対象のシートを取得する
'            intxlCount = xlBook.Sheets.Count
'            ReDim varPrintSeets(intxlCount - 2)
'            For intCount = 1 To intxlCount - 1
'                varPrintSeets(intCount - 1) = xlBook.Worksheets(intCount + 1).NAME
'            Next
'            xlBook.Activate
'            xlBook.Sheets(varPrintSeets).Select
'            intxlCount = xlBook2.Sheets.Count
'            xlBook.Sheets(varPrintSeets).Copy After:=xlBook2.Sheets(intxlCount)
''
''        Else
''            '基本入力シートに値を設定
''            Call MSZZ099.SetBaseSheetKyk(strBumonCode, _
''                                         xlBook2.Worksheets("基本入力"), _
''                                         stCARG, _
''                                         stCONT, _
''                                         stCARG2)
''            '出力対象のシートを取得する
''            intxlCount = xlBook2.Sheets.Count
''            ReDim varPrintSeets(intxlCount - 2)
''            For intCount = 1 To intxlCount - 1
''                varPrintSeets(intCount - 1) = xlBook2.Worksheets(intCount + 1).NAME
''            Next
''            xlBook2.Activate
''            xlBook2.Sheets(varPrintSeets).Select
''            intxlCount = xlBook.Sheets.Count
''            xlBook2.Sheets(varPrintSeets).Copy After:=xlBook.Sheets(intxlCount)
'            'For intCount = 1 To intxlCount - 1
'            '    xlBook2.Sheets(varPrintSeets(intCount - 1)).Select
'            '    'varPrintSeets(intCount - 1) = xlBook2.Worksheets(intCount + 1).NAME
'            '    intxlCount = xlBook.Sheets.Count
'            '    xlBook2.Sheets(varPrintSeets(intCount - 1)).Copy After:=xlBook.Sheets(intxlCount)
'            'Next
'            'xlBook2.Close False
'            '
''        End If
'
'    Next
'
'    '出力対象のシートを取得する
'    intxlCount = xlBook2.Sheets.Count
'    ReDim varPrintSeets(intxlCount - 2)
'    For intCount = 1 To intxlCount - 1
'        varPrintSeets(intCount - 1) = xlBook2.Worksheets(intCount + 1).NAME
'    Next
'    '印刷プレビュー表示
'    xlApp.Visible = True
'    xlBook2.Activate
'    Call xlBook2.Sheets(varPrintSeets).PrintPreview
'
'    xlBook2.Close False
'    'EXCELファイルを閉じる
'    xlBook.Close False
'    On Error GoTo ErrorHandler1
'    'EXCEL終了
'    xlApp.DisplayAlerts = False
'    xlApp.Quit
'
'    fncPrintPreview = True
'Exit_fncPrintPreview:
'    Exit Function
'ErrorHandler2:
'    'Resume
'    xlBook.Close False
'ErrorHandler1:
'    xlApp.Visible = True
'Err_fncPrintHelper:
'    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
'    Resume Exit_fncPrintPreview
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :fncPrintHelper
'        機能             :解約通知書印刷処理
'        IN               :
'        OUT              :fncPrintHelper,Output,Boolean,True 正常,False エラー
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncPrintHelper(pintCnt As Integer, _
                                pSHOKAI() As FVS220Type_KAIYAKUSHOKAI) As Boolean
    Dim intCnt          As Integer
        
    For intCnt = 0 To pintCnt - 1
        If funcPrintHelperKyk(pSHOKAI(intCnt).FVS220_YCODE, pSHOKAI(intCnt).FVS220_CNTANO, pSHOKAI(intCnt).FVS220_MOVE_YCODE, pSHOKAI(intCnt).FVS220_MOVE_CNTANO) = False Then GoTo Err_fncPrintHelper
    Next
    fncPrintHelper = True
Exit_fncPrintHelper:
    Exit Function

Err_fncPrintHelper:
    Call MSZZ024_M00("fncPrintHelper", True)
    'Err.Clear
    'MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_fncPrintHelper

End Function

'==============================================================================*
'
'        MODULE_NAME      :funcPrintHelperKyk
'        機能             :ヘルパー連動印刷
'        PARAM            :strYardc            ヤードコード(I)
'                         :strCARG_NO          コンテナ番号(I)
'                         :strYardcido         移動先ヤードコード(I)
'                         :strCARG_Noido       移動先コンテナ番号(I)
'                         :intPrevew       　　プレビューフラグ(I)
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function funcPrintHelperKyk(ByVal strYardc As String, _
                             ByVal strCARG_NO As String, _
                             ByVal strYardcido As String, _
                             ByVal strCARG_Noido As String, _
                             Optional intPrevew As Integer = 0) As Boolean
On Error GoTo ErrorHandler

    Dim strHelperFileName   As String
    Dim varPrintSeets       As Variant
    Dim blnError            As Boolean
    funcPrintHelperKyk = False
    blnError = False
    doCmd.Hourglass True
    If strBumonCode = "" Then strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    
    Call MSZZ049.HelperPrintPreviewKyk(strBumonCode, strYardc, strCARG_NO, strYardcido, strCARG_Noido, intPrevew)

    doCmd.Hourglass False
    blnError = True
    funcPrintHelperKyk = True
    Exit Function
        
ErrorHandler:
    
    doCmd.Hourglass False
    If Not blnError Then
        Call Err.Raise(Err.Number, "funcPrintHelperKyk" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function


'==============================================================================*
'   外部関数
'==============================================================================*
'==============================================================================*
'
'        MODULE_NAME      : キー押下処理
'        MODULE_ID        : KeyDown
'        PARAM            : KeyCode           キーコード
'                         : Shift             シフトキー状態
'        作成日           :2009/04/30  新規作成
'        更新日           :0000/00/00
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    'F1:戻る
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Me.cmd_終了.SetFocus
        'Call cmd_終了_Click
    Case vbKeyF2
        KeyCode = 0
        Me.cmd_検索.SetFocus
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select
End Sub
'****************************  ended or program ********************************

