VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　送り状リスト作成
'   プログラムＩＤ　：　FVS410
'   作　成　日　　　：  2006/02/21
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'
'   修　正　日　　　：2006/05/12
'   修　正　者　　　：イーグルソフト 柴崎
'   修　正　内　容　：1.次の項目をコントロールマスタに追加。そこから取得するよう変更。
'   　　　　　　　　：・依頼主名
'   　　　　　　　　：・依頼主郵便番号
'   　　　　　　　　：・依頼主住所１～３
'   　　　　　　　　：・依頼主電話番号
'   Ver 　　　　　　：0.1
'
'   修　正　日　　　：2009/01/25
'   修　正　者　　　：peag
'   修　正　内　容　：1.自動割引対応
'   Ver 　　　　　　：0.2
'
'   修　正　日　　　：2013/06/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：終了メッセージを表示しないように修正
'   Ver 　　　　　　：0.3
'
'   修　正　日　　　：2015/04/28
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：年払いの際には月額使用料 －１ヶ月の料金とする。
'   Ver 　　　　　　：0.4
'
'   修　正　日　　　：2017/06/22
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：配送指定時間を変更
'   Ver 　　　　　　：0.5
'
'   修　正　日　　　：2019/12/16
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：0.3で行った年払いの対応を廃止
'   Ver 　　　　　　：0.6
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'部門コード
Private strBumonCode    As String

'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

'桁数
Private Const P_DATE_LENGTH         As Integer = 10      '日付

'紹介請求方法オプションValue値
Private Const P_OPT_SEIKYU_CD_代引き        As Integer = 1
Private Const P_OPT_SEIKYU_CD_元払い        As Integer = 2

'既出力データオプションValue値
Private Const P_OPT_ALREADY_含めない        As Integer = 1
Private Const P_OPT_ALREADY_含める          As Integer = 2

'送り状種類
Private Const P_SENDKIND_発払い         As String = "0"
Private Const P_SENDKIND_コレクト       As String = "2"

'配達指定時間帯
Private Const P_TIMEZONE_指定なし           As String = ""
Private Const P_TIMEZONE_午前中             As String = "0812"
Private Const P_TIMEZONE_1214               As String = "1214"
Private Const P_TIMEZONE_1416               As String = "1416"
Private Const P_TIMEZONE_1618               As String = "1618"
Private Const P_TIMEZONE_1820               As String = "1820"
'Private Const P_TIMEZONE_2021               As String = "2021"
Private Const P_TIMEZONE_2021               As String = "1921"


'作成ファイル名先頭部分
Private Const P_FILENAME_HEAD               As String = "送状B2リスト"
'---- コード値 ----

'初期費用請求方法
Private Const P_SEIKYUCD_代引き         As Integer = 2

'配達指定時間帯
Private Const P_SPECIFY_TIME_指定無し   As Integer = 1
Private Const P_SPECIFY_TIME_午前中     As Integer = 2
Private Const P_SPECIFY_TIME_1214       As Integer = 3
Private Const P_SPECIFY_TIME_1416       As Integer = 4
Private Const P_SPECIFY_TIME_1618       As Integer = 5
Private Const P_SPECIFY_TIME_1820       As Integer = 6
Private Const P_SPECIFY_TIME_2021       As Integer = 7

'ADO定数（アクティブXを参照設定しないためローカルで定義する）
Private Const adCmdText = 1
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1
Private Const adOpenDynamic = 2
Private Const adLockPessimistic = 2
Private Const adCmdStoredProc = 4

'==============================================================================*
'
'        MODULE_NAME      :cmdCreate_Click
'        機能             :cmdCreate_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdCreate_Click()
On Error GoTo ErrorHandler
    
    If Nz(txtKeySendDate) = "" Then
        MsgBox "鍵送付日は必須入力です。", vbCritical, P_ERROR
        txtKeySendDate.SetFocus
        Exit Sub
    End If
    
    If MsgBox("送り状リストを作成します。よろしいですか？", vbYesNo + vbQuestion, P_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    
    Call fncMakeSendList
    
    Exit Sub
ErrorHandler:
    Call MSZZ024_M00("cmdCreate_Click", True)
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncMakeSendList
'        機能             :送り状リスト作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeSendList() As Boolean
On Error GoTo ErrorHandler

    Dim dbSQLServer         As Object
    Dim rsData              As Object
    Dim blnError            As Boolean
    Dim blnIsTransaction    As Boolean
    Dim strKaseInfo()       As String
    Dim strUserID           As String
    Dim lngCount            As Long
    Dim strFilePath         As String
    Dim strBumonCode        As String
    
    blnIsTransaction = False
    blnError = False
    
    fncMakeSendList = False
    doCmd.Hourglass True
    
    'SQLサーバー接続
    '部門コード取得
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("DLookup(""CONT_BUMOC"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", "テーブル[CONT_MAST]の設定不正。")
    End If
    
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 0
    Call dbSQLServer.Open(MSZZ007_M10(strBumonCode))
    
    '加瀬固有情報取得
    Call subGetKaseInfo(dbSQLServer, strKaseInfo)
    
    'トランザクション開始
    dbSQLServer.BeginTrans
    blnIsTransaction = True
    
    'データ取得
    Set rsData = CreateObject("ADODB.Recordset")
    Call rsData.Open(fncMakeSql(), dbSQLServer, adOpenDynamic, adLockPessimistic)
    
    'カギ送付日＋現在時間で作る
    strFilePath = txtSavePath & P_FILENAME_HEAD & Format(txtKeySendDate, "yyyymmdd") & Format(Now, "hhmmss") & ".csv"
    Open strFilePath For Output As #1
    
    'データ出力
    strUserID = LsGetUserName()
    lngCount = 0
    While Not rsData.EOF
        lngCount = lngCount + 1
        '送り状Excelシート値設定
        Call subSetCsv(rsData, strKaseInfo)
        
        'データ更新
        rsData.Fields("RCPT_SLIP_NO") = "-"
        rsData.Fields("RCPT_UPDAD") = Format(DATE, "yyyymmdd")
        rsData.Fields("RCPT_UPDAJ") = Format(Now, "hhmmss")
        rsData.Fields("RCPT_UPDPB") = Me.NAME
        rsData.Fields("RCPT_UPDUB") = strUserID
        rsData.UPDATE
        rsData.MoveNext
    Wend
    
    Close
    
    'コミット
    dbSQLServer.CommitTrans
    blnIsTransaction = False

    rsData.Close
    Set rsData = Nothing
    dbSQLServer.Close
    Set dbSQLServer = Nothing
    
    If lngCount = 0 Then
        Call subKillFile(strFilePath)
        MsgBox "該当データが存在しません", , P_KAKUNIN
    Else
        MsgBox "送り状リストを作成しました。" & vbCrLf & strFilePath & vbCrLf & Nz(lngCount, 0) & " 件", , "正常終了"
    End If

    fncMakeSendList = True
    GoTo ExitRtn
    
ErrorHandler:
    If blnIsTransaction Then dbSQLServer.RollbackTrans
    Call subKillFile(strFilePath)
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    Close

    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "fncMakeSendList" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeSql
'        機能             :データ検索SQL作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeSql() As String

    Dim strSQL      As String
    
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM RCPT_TRAN "
    strSQL = strSQL & "  WHERE RCPT_KEY_SEND_DATE = '" & txtKeySendDate & "' "
    
    '請求方法の条件が「代引き」の場合、初回請求方法＝「代引き」
    '請求方法の条件が「元払い」の場合、初回請求方法≠「代引き」
    strSQL = strSQL & "    AND RCPT_SEIKYU_CD "
    strSQL = strSQL & IIf(fraSeikyuCd.VALUE = P_OPT_SEIKYU_CD_代引き, " = ", " <> ")
    strSQL = strSQL & P_SEIKYUCD_代引き
    
    '既出力データの条件　「含める」の場合は無条件　「含めない」の場合は伝票番号が空のもののみ検索
    If fraAlready.VALUE = P_OPT_ALREADY_含めない Then
        strSQL = strSQL & "    AND ( RCPT_SLIP_NO IS NULL "
        strSQL = strSQL & "       OR RCPT_SLIP_NO = '' ) "
    End If
    
    fncMakeSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subKillFile
'        機能             :ファイル削除
'        IN               :対象ファイル名
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKillFile(strFilename As String)

    If strFilename = "" Then
        Exit Sub
    End If
    
    If Dir(strFilename) <> "" Then
        Kill strFilename
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subGetKaseInfo
'        機能             :加瀬固有情報取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetKaseInfo(objDb As Object, ByRef strKaseInfo() As String)
On Error GoTo ErrorHandler

    Dim blnError            As Boolean
    Dim rsContMast          As Object
    Dim strSQL              As String
    
    blnError = False
    
    ReDim strKaseInfo(0 To 10)
    
    'コントロールマスタデータ取得
    Set rsContMast = CreateObject("ADODB.Recordset")
    strSQL = " SELECT * FROM CONT_MAST "
    strSQL = strSQL & " WHERE CONT_KEY = '1' "
    Call rsContMast.Open(strSQL, objDb, adOpenForwardOnly, adLockReadOnly)
    
    If rsContMast.EOF Then
        Call MSZZ024_M10("CONT_MAST", "コントロールマスタのデータを取得できませんでした。")
    End If
    
    With rsContMast
        strKaseInfo(0) = Nz(.Fields("CONT_CLIENT_CODE"))
        
        '2006/05/12 SHIBAZAKI コントロールマスタ項目変更
        strKaseInfo(1) = Nz(.Fields("CONT_CLIENT_TEL"))
        strKaseInfo(2) = Nz(.Fields("CONT_CLIENT_NAME"))
        strKaseInfo(3) = Replace(Nz(.Fields("CONT_CLIENT_YUBINO")), "-", "")
        strKaseInfo(4) = Nz(.Fields("CONT_CLIENT_ADDR_1")) & Nz(.Fields("CONT_CLIENT_ADDR_2"))
        If Nz(.Fields("CONT_CLIENT_ADDR_3")) <> "" Then
            strKaseInfo(4) = strKaseInfo(4) & "　" & .Fields("CONT_CLIENT_ADDR_3")
        End If
'        strKaseInfo(1) = Nz(.Fields("CONT_TEL_NO"))
'        strKaseInfo(2) = Nz(.Fields("CONT_KAISYA"))
'        strKaseInfo(3) = Replace(Nz(.Fields("CONT_YUBINO")), "-", "")
'        strKaseInfo(4) = Nz(.Fields("CONT_ADDR_1")) & Nz(.Fields("CONT_ADDR_2"))
        
        strKaseInfo(5) = Nz(.Fields("CONT_CLIENT_KANA"))
        strKaseInfo(6) = Nz(.Fields("CONT_DEMAND_USER_CODE"))
        strKaseInfo(7) = Nz(.Fields("CONT_CARRIAGE_CODE"))
        strKaseInfo(8) = Nz(.Fields("CONT_GOODS_CODE"))
        strKaseInfo(9) = Nz(.Fields("CONT_GOODS"))
        strKaseInfo(10) = Nz(.Fields("CONT_KIJI"))
        
        .Close
    End With
    
    Set rsContMast = Nothing
    
    Exit Sub
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsContMast Is Nothing Then rsContMast.Close: Set rsContMast = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "subGetKaseInfo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetExcelRange
'        機能             :送り状リスト値設定
'        IN               :rsData 検索データ
'                         :objSeet 設定対象
'                         :strKaseInfo 加瀬固有情報
'                         :lngRow 設定行
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetCsv(rsData As Object, strKaseInfo() As String)

    Dim strTemp     As String
    
    strTemp = ""
    With rsData
        '出荷予定日 = カギ送付日
        strTemp = strTemp & Format(.Fields("RCPT_KEY_SEND_DATE"), "yyyy/mm/dd") & ","
        'お客様管理番号＝受付番号
        strTemp = strTemp & .Fields("RCPT_NO") & ","

        '送り状種類＝初回請求区分
        If .Fields("RCPT_SEIKYU_CD") = P_SEIKYUCD_代引き Then
            '代引き＝コレクト
            strTemp = strTemp & P_SENDKIND_コレクト & ","
        Else
            '代引き以外＝発払い
            strTemp = strTemp & P_SENDKIND_発払い & ","
        End If
        
        '温度区分=空白
        strTemp = strTemp & ","

        '送付先情報設定
        strTemp = strTemp & fncSetSendInfo(rsData)

        '加瀬固有情報設定
        strTemp = strTemp & fncSetKaseInfo(strKaseInfo)

        '配達指定日＝カギ到着予定日
        strTemp = strTemp & Format(.Fields("RCPT_ARRIVAL_EXPECT"), "yyyy/mm/dd") & ","

        '配達指定時間帯＝カギ配達指定時間帯
        strTemp = strTemp & fncGetTimeZone(Nz(.Fields("RCPT_SPECIFY_TIME"))) & ","
        
        'ｺﾚｸﾄ代金引換額＝初回請求金額
        strTemp = strTemp & fncGetPay(rsData) & ","
        
        '１０個分空白
        strTemp = strTemp & ",,,,,,,,,,"

        '営業所留め置き & 最後二つ空白
        strTemp = strTemp & "0,,"
    End With

    Print #1, strTemp

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncGetPay
'        機能             :初回請求金額計算
'        IN               :rsData 検索データ
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetPay(rsData As Object) As Double

    Dim intPayMonth As Integer
    Dim dblPay      As Double

    With rsData
        '前払い月数
        intPayMonth = DateDiff("m", _
                               Format(.Fields("RCPT_KISAN_DATE"), "yyyy/mm/dd"), _
                               Left(.Fields("RCPT_SEIKYU_KIKAN"), 4) & "/" & Right(.Fields("RCPT_SEIKYU_KIKAN"), 2) & "/01")
        '月額使用料＋毎月雑費
        dblPay = Nz(.Fields("RCPT_RENTKG"), 0) _
               + Nz(.Fields("RCPT_EZAPPI"), 0) _
               + Nz(.Fields("RCPT_ADD_EZAPPI1"), 0) _
               + Nz(.Fields("RCPT_ADD_EZAPPI2"), 0)
        
        '前払い月数を掛ける
        '年払いの際には、前払い月数 -1とする。
        '2015/04/28 M.HONDA INS
        'DELETE 2019/12/16 N.IMAI If intPayMonth = 11 Then
        'DELETE 2019/12/16 N.IMAI     dblPay = dblPay * (intPayMonth - 1)
        'DELETE 2019/12/16 N.IMAI Else
            dblPay = dblPay * intPayMonth
        'DELETE 2019/12/16 N.IMAI End If
        '2015/04/28 M.HONDA INS

        'プラス初回使用料
        dblPay = dblPay + Nz(.Fields("RCPT_FIRSTKG"), 0)
        
        'プラス日割り毎月雑費
        dblPay = dblPay + Nz(.Fields("RCPT_EZAPPI_DAILY"), 0) _
                        + Nz(.Fields("RCPT_ADD_EZAPPI_DAILY1"), 0) _
                        + Nz(.Fields("RCPT_ADD_EZAPPI_DAILY2"), 0)
        
        'プラス初回のみ雑費
        dblPay = dblPay + Nz(.Fields("RCPT_FZAPPI"), 0) _
                        + Nz(.Fields("RCPT_ADD_FZAPPI1"), 0) _
                        + Nz(.Fields("RCPT_ADD_FZAPPI2"), 0)
        
        'プラス保証金
        dblPay = dblPay + Nz(.Fields("RCPT_SECUKG"), 0)
        
        'マイナス支払済み金額
        dblPay = dblPay - Nz(.Fields("RCPT_SUMIKG"), 0)
    
        '割引額反映（マイナス値なのでそのまま足し込み 2009/01/25
        dblPay = dblPay + Nz(.Fields("RCPT_FWARIBIKI"), 0) + Nz(.Fields("RCPT_EWARIBIKI"), 0)
    End With
    
    fncGetPay = dblPay
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subSetSendInfo
'        機能             :送付先情報設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetSendInfo(rsData As Object) As String

    Dim strTemp     As String
    Dim strAdd      As String
    
    strTemp = ""
    With rsData
        'お届け先コード＝顧客コード
        strTemp = strTemp & Format(.Fields("RCPT_UCODE"), "000000") & ","
        'お届け先電話番号＝カギ送付先電話番号
        strTemp = strTemp & .Fields("RCPT_SEND_TEL") & ","
        
        'お届け先電話番号枝番＝空白
        strTemp = strTemp & ","
        
        'お届け先名＝カギ送付先名
        strTemp = strTemp & .Fields("RCPT_SEND_NAME") & ","
        'お届け先郵便番号＝カギ送付先郵便番号
        strTemp = strTemp & Replace(.Fields("RCPT_SEND_YUBINO"), "-", "") & ","
        
        'お届け先住所＝カギ送付先住所
        strAdd = .Fields("RCPT_SEND_ADR_1") & .Fields("RCPT_SEND_ADR_2")
        If Nz(.Fields("RCPT_SEND_ADR_3")) <> "" Then
            '建物名（住所３）は全角スペースで区切る
            strAdd = strAdd & "　" & .Fields("RCPT_SEND_ADR_3")
        End If
        strTemp = strTemp & strAdd & ","
        
        '会社・部門名1 & 会社・部門名2 & お届け先略称ｶﾅ＝空白
        strTemp = strTemp & ",,,"
        
        '敬称＝送付先敬称
        cmbSendHonorific = Format(Nz(.Fields("RCPT_SEND_HONORIFIC"), 0), "00")
        If cmbSendHonorific <> -1 Then
            strTemp = strTemp & cmbSendHonorific.Column(1, cmbSendHonorific.ListIndex)
        End If
        strTemp = strTemp & ","
    End With
    
    fncSetSendInfo = strTemp

End Function

'==============================================================================*
'
'        MODULE_NAME      :subSetKaseInfo
'        機能             :加瀬固有情報設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetKaseInfo(strKaseInfo() As String) As String

    Dim strTemp     As String
    
    strTemp = ""
    
    'ご依頼主コード
    strTemp = strTemp & strKaseInfo(0) & ","
    'ご依頼主電話番号
    strTemp = strTemp & strKaseInfo(1) & ","
    
    'ご依頼主電話番号枝番＝空白
    strTemp = strTemp & ","
    
    'ご依頼主名
    strTemp = strTemp & strKaseInfo(2) & ","
    'ご依頼主郵便番号
    strTemp = strTemp & strKaseInfo(3) & ","
    
    'ご依頼主住所
    '2006/05/12 SHIBAZAKI コントロールマスタ項目変更
'    strTemp = strTemp & "神奈川県" & strKaseInfo(4) & ","
    strTemp = strTemp & strKaseInfo(4) & ","
    'ご依頼主省略カナ
    strTemp = strTemp & strKaseInfo(5) & ","
    
    '請求先顧客コード
    strTemp = strTemp & strKaseInfo(6) & ","
    
    '請求先分類ｺｰﾄﾞ＝空白
    strTemp = strTemp & ","
    
    '運賃管理番号
    strTemp = strTemp & strKaseInfo(7) & ","
    
    '品名コード１
    strTemp = strTemp & strKaseInfo(8) & ","
    '品名１
    strTemp = strTemp & strKaseInfo(9) & ","
    
    '品名ｺｰﾄﾞ2 & 品名2 & 荷扱い1 & 荷扱い2＝空白
    strTemp = strTemp & ",,,,"
    
    '記事
    strTemp = strTemp & strKaseInfo(10) & ","

    fncSetKaseInfo = strTemp
    
End Function


'==============================================================================*
'
'        MODULE_NAME      :fncGetTimeZone
'        機能             :配達指定時間取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetTimeZone(strSpecifyTime As String) As String

    Dim strTemp     As String
    
    Select Case strSpecifyTime
    Case P_SPECIFY_TIME_指定無し
        strTemp = P_TIMEZONE_指定なし
    Case P_SPECIFY_TIME_午前中
        strTemp = P_TIMEZONE_午前中
    Case P_SPECIFY_TIME_1214
        strTemp = P_TIMEZONE_1214
    Case P_SPECIFY_TIME_1416
        strTemp = P_TIMEZONE_1416
    Case P_SPECIFY_TIME_1618
        strTemp = P_TIMEZONE_1618
    Case P_SPECIFY_TIME_1820
        strTemp = P_TIMEZONE_1820
    Case P_SPECIFY_TIME_2021
        strTemp = P_TIMEZONE_2021
    Case Else
        strTemp = ""
    End Select
    
    fncGetTimeZone = strTemp

End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activate イベント
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
'        機能             :Form_KeyDown イベント
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
            Me.cmdEnd.SetFocus
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdEnd_Click
'        機能             :cmdEnd_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdEnd_Click()
    
    '' 2013/06/03 M.HONDA START
'    If MsgBox("終了します。よろしいですか？", vbYesNo + vbQuestion, P_KAKUNIN) = vbNo Then
'        Exit Sub
'    End If
'    On Error Resume Next
    '' 2013/06/03 M.HONDA END
    doCmd.Close acForm, Me.FormName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form_Open イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
On Error GoTo ErrorHandler
     Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)

    End Select
    'ADO存在チェック
    If Not MSZZ013_M00("ADODB.Connection") Or Not MSZZ013_M00("ADODB.Recordset") Then
        Call MSZZ024_M10("MSZZ013_M00", "「ActiveX Data Objects」がインストールされてません")
    End If
    
    'ファイル保存先取得
    txtSavePath = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""FVS410"" AND INTIF_RECFB = ""SAVE_PATH""")
    If Nz(txtSavePath) = "" Then
        Call MSZZ024_M10("DLookup(""INTIF_RECDB"", ""INTI_FILE"", ""INTIF_PROGB=""FVS410"" AND INTIF_RECFB = ""SAVE_PATH"")", _
                         "送り状リストの保存先を取得できませんでした。")
    End If
    If Right(txtSavePath, 1) <> "\" Then
        txtSavePath = txtSavePath & "\"
    End If
    
    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("Form_Open", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeySendDate_AfterUpdate
'        機能             :txtKeySendDate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeySendDate_AfterUpdate()

    txtKeySendDate = Format(VarChangeDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeySendDate_BeforeUpdate
'        機能             :txtKeySendDate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeySendDate_BeforeUpdate(Cancel As Integer)

    Call fncDateCheck(txtKeySendDate, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeySendDate_GotFocus
'        機能             :txtKeySendDate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeySendDate_GotFocus()

    txtKeySendDate = Format(txtKeySendDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeySendDate_KeyPress
'        機能             :txtKeySendDate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeySendDate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtKeySendDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeySendDate_LostFocus
'        機能             :txtKeySendDate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeySendDate_LostFocus()

    txtKeySendDate = Format(txtKeySendDate, "yyyy/mm/dd")

End Sub

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
    
    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

End Function
