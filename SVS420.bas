Attribute VB_Name = "SVS420"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　B2B配送データ取込
'   プログラムＩＤ　：　SVS420
'   作　成　日　　　：  2006/02/27
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
Option Compare Database
Option Explicit

#Const SEND_MAIL = True
#Const DEL_FILE = True

Private dtProcDate                      As Date
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'取込エラー番号
Private Const P_ERROR_受付番号空白          As Integer = 1
Private Const P_ERROR_受付番号該当無し      As Integer = 2
Private Const P_ERROR_既に取込済み          As Integer = 3
Private Const P_ERROR_部門コード不一致      As Integer = 4

'取込エラーログ格納先　取込ファイル格納フォルダ内の下記フォルダに格納する
Private Const P_TORIKOMI_ERROR_LOG_DIR      As String = "ErrorLog"

'取込データ配列要素番号
Private Const P_PROCDATA_顧客コード         As Integer = 0
Private Const P_PROCDATA_顧客名             As Integer = 1
Private Const P_PROCDATA_受付番号           As Integer = 2
Private Const P_PROCDATA_伝票番号           As Integer = 3
Private Const P_PROCDATA_配送状況           As Integer = 4
Private Const P_PROCDATA_日付               As Integer = 5

'配送状況
Private Const P_DERIVERYSTATUS_配達完了     As String = "配達完了"

'プログラムID
Private Const P_PROGRAM_ID                  As String = "SVS420"

'CAMP無視のテスト
Sub a00TEST_SVS420()

    If SVS420_M00 Then
        MsgBox "ok"
    Else
        MsgBox "error"
    End If
    
End Sub

'==============================================================================*
'
'       MODULE_NAME     : B2B配送データ取込
'       MODULE_ID       : SVS420_M00
'       CREATE_DATE     :
'       PARAM           :
'       RETURN          : 正常(True) エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function SVS420_M00() As Boolean
On Error GoTo ErrorHandler
    
    Dim strBumonArr()       As String
    Dim varFilePath         As Variant
    Dim strTemp             As String
    Dim strFilename()       As String
    Dim lngCount            As Long
    Dim strErrLogFile       As String
    Dim strMailSubject      As String
    Dim strMailMessage      As String
    Dim strMailAttach       As String
    
    SVS420_M00 = False
    
    '開始ログ
    Call MSZZ003_M00(P_PROGRAM_ID, "0", "")
    
    dtProcDate = Now
    
    '対象部門コード取得
    strBumonArr = fncGetBumonArr()
    
    '取込ファイル格納パス取得
    varFilePath = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""" & P_PROGRAM_ID & """ AND INTIF_RECFB = ""FILE_PATH""")
    If Nz(varFilePath) = "" Then
        Call MSZZ024_M10("DlookUp", "取込ファイルの格納場所を取得できませんでした。")
    End If
    If Right(varFilePath, 1) <> "\" Then
        varFilePath = varFilePath & "\"
    End If
    
    '取込ファイル格納パス配下のＣＳＶファイル名を配列に格納
    ReDim strFilename(0)
    strTemp = Dir(varFilePath & "*.csv")
    lngCount = 0
    While strTemp <> ""
        lngCount = lngCount + 1
        ReDim Preserve strFilename(lngCount)
        strFilename(lngCount - 1) = strTemp
        strTemp = Dir()
    Wend
    
    '主処理
    If fncMain(strBumonArr, Nz(varFilePath), strFilename) > 0 Then
        '取込エラーメール用件名＆メッセージ
        strMailSubject = "【取込エラー】"
        strMailMessage = "取込エラーが発生しました。" & vbCrLf & _
                         "実行ログを添付します。" & vbCrLf & _
                         "取込エラーの詳細は、下記フォルダ内のログファイルを参照してください。" & vbCrLf & _
                         varFilePath & P_TORIKOMI_ERROR_LOG_DIR
    Else
        '正常終了メール用件名＆メッセージ
        strMailSubject = "【正常終了】"
        strMailMessage = "正常終了しました。" & vbCrLf & "実行ログを添付します。"
    End If
    strMailAttach = ""  '追加添付ファイル無し
    
    '正常終了ログ
    Call MSZZ003_M00(P_PROGRAM_ID, "1", "")
    
    SVS420_M00 = True
    
    GoTo ExitRtn

ErrorHandler:
    'エラーログ出力
    strErrLogFile = MSZZ024_M00("SVS420_M00", False)
    'エラー終了ログ出力
    Call MSZZ003_M00(P_PROGRAM_ID, "9", "【不正終了】エラーログ[" & strErrLogFile & "]")
    Err.Clear
    
    '不正終了メール用件名＆メッセージ
    strMailSubject = "【不正終了】"
    strMailMessage = "不正終了しました。" & vbCrLf & "実行ログとエラーログを添付します。"
    strMailAttach = strErrLogFile   'エラーログを添付

ExitRtn:
#If SEND_MAIL Then
    'メール配信
    Call MSZZ021_M00(P_PROGRAM_ID, strMailSubject, strMailMessage, strMailAttach)
#End If

End Function

'==============================================================================*
'
'       MODULE_NAME     : 部門コード取得
'       MODULE_ID       : fncGetBumonArr
'       CREATE_DATE     :
'       PARAM           :
'       RETURN          : 部門コードの配列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBumonArr() As String()

On Error GoTo ErrorHandler

    Dim strSqlserver    As String
    Dim strConnect      As String
    Dim dbSQLServer     As DAO.Database
    Dim rsData          As DAO.Recordset
    Dim blnError        As Boolean
    Dim strBumonArr()   As String
    Dim lngCount        As Long
    Dim strSQL          As String
    
    'SQL-Server名
    strSqlserver = fncGetSqlServer()

    '接続文字列取得
    strConnect = fncGetConnectString()

    'SQLサーバー接続(加瀬DB)
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    'プログラムパラメータテーブル検索
    strSQL = " SELECT PGPAT_PARAN " & _
             "   FROM PGPA_TABL " & _
             "  WHERE PGPAT_PGP2B = '" & P_PROGRAM_ID & "'"
    Set rsData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If rsData.EOF Then
        Call MSZZ024_M10(strSQL, "部門コードを取得できませんでした。")
    End If

    With rsData
        lngCount = 0
        While Not .EOF
            If Nz(.Fields("PGPAT_PARAN")) = "" Then
                Call MSZZ024_M10(strSQL, "部門コードの設定が正しくありません。")
            End If
            
            '配列に格納
            lngCount = lngCount + 1
            ReDim Preserve strBumonArr(lngCount)
            strBumonArr(lngCount - 1) = Nz(.Fields("PGPAT_PARAN"))
            
            .MoveNext
        Wend
        .Close
    End With

    Set rsData = Nothing

    'データベース切断
    dbSQLServer.Close
    Set dbSQLServer = Nothing

    fncGetBumonArr = strBumonArr

    Exit Function

ErrorHandler:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    Call Err.Raise(Err.Number, "fncGetBumonArr" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'       MODULE_NAME     : 主処理
'       MODULE_ID       : fncMain
'       CREATE_DATE     :
'       PARAM           : strBumonCode  部門コードの配列
'                       : strFilePath   取込ファイルが格納されたパス
'                       : strFileName   取込ファイル名の配列
'       RETURN          : 正常(0) エラー(-1) 取込エラー(1)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMain(strBumonCode() As String, _
                         strFilePath As String, _
                         strFilename() As String) As Integer
On Error GoTo ErrorHandler

    Dim objWs           As Workspace
    Dim lngCount        As Long
    Dim dbSQLServer()   As DAO.Database
    Dim blnTorikomiErr  As Boolean
    
    fncMain = -1
    
    ReDim dbSQLServer(UBound(strBumonCode))
    
    '対象部門全てのデータベースに接続する
    Call fncConnectDatabase(strBumonCode, dbSQLServer)
    
    'トランザクション開始
    Set objWs = DBEngine.Workspaces(0)
    objWs.BeginTrans
    
    '全ファイル処理
    blnTorikomiErr = False
    For lngCount = 0 To UBound(strFilename) - 1
        'ファイルを読み込んで受付トランを更新する
        If fncProcTorikomiFile(dbSQLServer, strBumonCode, strFilePath, strFilename(lngCount)) > 0 Then
            '取込エラーが発生している
            blnTorikomiErr = True
        End If
    Next

    'コミット
    objWs.CommitTrans
    Set objWs = Nothing
    
    'データベース切断
    Call fncDisconnectDatabase(dbSQLServer)

#If DEL_FILE Then
    'ファイルバックアップ
    For lngCount = 0 To UBound(strFilename) - 1
        Call subFileBackup(strFilePath, strFilename(lngCount))
    Next
#End If

    '取込エラー時は１を返却
    fncMain = IIf(blnTorikomiErr, 1, 0)

    Exit Function
    
ErrorHandler:
    'ロールバック
    If Not objWs Is Nothing Then
        objWs.Rollback
        Set objWs = Nothing
    End If
    
    'データベース切断
    Call fncDisconnectDatabase(dbSQLServer)
    
    Call Err.Raise(Err.Number, "fncMain" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'       MODULE_NAME     : 取込ファイルを処理
'       MODULE_ID       : fncProcTorikomiFile
'       CREATE_DATE     :
'       PARAM           : objDb                 DAO.Databaseの配列
'                       : strBumonCode          部門コードの配列
'                       : strFilePath           取込ファイルが格納されたパス
'                       : strFileName           取込ファイル名
'       RETURN          : 正常(0) エラー(-1) 取込エラー(1以上)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncProcTorikomiFile(objDb() As Database, _
                                     strBumonCode() As String, _
                                     strFilePath As String, _
                                     strFilename As String) As Integer
On Error GoTo ErrorHandler
    
    Dim lngIndex            As Long
    Dim strRead             As String
    Dim strProcData()       As String
    Dim strErrLogFile()     As String   '取込エラーログファイル名（部門ごと）
    Dim strComErrLogFile    As String   '取込エラーログファイル名（部門共通）
    Dim strUserID           As String

    Dim intError            As Integer

    Dim lngReadCnt          As Long     '読み込み件数
    Dim lngBumonUnmatchCnt  As Long     '部門コード不一致件数
    Dim lngRcptNoBlankCnt   As Long     '受付番号空白件数
    Dim lngErrCnt()         As Long     'エラー件数(部門ごと)
    Dim lngProcCnt()        As Long     '取込件数(部門ごと)
    Dim lngBumonSameCnt()   As Long     '部門コード一致件数(部門ごと)
    Dim lngErrSum           As Long     'エラー件数合計
    Dim lngProcSum          As Long     '取込件数合計

    fncProcTorikomiFile = -1

    Call MSZZ003_M00(P_PROGRAM_ID, "8", "取込ファイル = [" & strFilePath & strFilename & "]")

    '取込ファイルオープン
    Open strFilePath & strFilename For Input As #1

    '件数カウンタをクリア/取込エラーログファイル名を初期化
    lngReadCnt = 0
    lngBumonUnmatchCnt = 0
    lngRcptNoBlankCnt = 0
    lngErrSum = 0
    lngProcSum = 0
    strComErrLogFile = ""
    
    ReDim lngErrCnt(UBound(strBumonCode))
    ReDim lngProcCnt(UBound(strBumonCode))
    ReDim lngBumonSameCnt(UBound(strBumonCode))
    ReDim strErrLogFile(UBound(strBumonCode))
    For lngIndex = 0 To UBound(strBumonCode) - 1
        lngErrCnt(lngIndex) = 0
        lngProcCnt(lngIndex) = 0
        lngBumonSameCnt(lngIndex) = 0
        strErrLogFile(lngIndex) = ""
    Next
    
    strUserID = LsGetUserName()

    While Not EOF(1)
        'ファイル読み込み
        Line Input #1, strRead

        '取り込みデータ(CSV)を配列化
        strProcData = Split(strRead, ",")
        '最後尾に要素追加
        '取込処理でエラーログ用の値を入れるためのテンポラリとする
        ReDim Preserve strProcData(UBound(strProcData) + 1)
        strProcData(UBound(strProcData)) = ""

        lngReadCnt = lngReadCnt + 1     '読み込み件数

        If Trim(strProcData(P_PROCDATA_受付番号)) = "お客様管理番号" Then
            '受付番号の位置に格納されている文字列が"お客様管理番号"の場合は見出し行なので読み飛ばす。
            lngReadCnt = lngReadCnt - 1
        ElseIf Trim(strProcData(P_PROCDATA_受付番号)) = "" Then
            '受付番号が格納されていない場合はエラーログに出力する(部門共通)
            Call subErrorLog(strFilePath, strFilename, strComErrLogFile, strProcData, "_", strRead, P_ERROR_受付番号空白)
            lngRcptNoBlankCnt = lngRcptNoBlankCnt + 1
        Else
            '受付番号2桁目が一致する部門コードを配列内から探す
            lngIndex = fncGetIndex(strBumonCode, Mid(Trim(strProcData(P_PROCDATA_受付番号)), 2, 1))
            If lngIndex = -1 Then
                '部門コードが一致しない場合はエラーログに出力する(部門共通)
                Call subErrorLog(strFilePath, strFilename, strComErrLogFile, strProcData, "_", strRead, P_ERROR_部門コード不一致)
                lngBumonUnmatchCnt = lngBumonUnmatchCnt + 1
            Else
                '受付番号2桁目が部門コードと一致するレコードを処理する
                lngBumonSameCnt(lngIndex) = lngBumonSameCnt(lngIndex) + 1     '部門コード一致件数
                intError = 0
                '受付トラン更新
                Call fncUpdateRcpt(objDb(lngIndex), strProcData, intError, strUserID)
    
                If intError = 0 Then
                    '正常取込
                    lngProcCnt(lngIndex) = lngProcCnt(lngIndex) + 1
                    lngProcSum = lngProcSum + 1
                Else
                    '取込エラーが発生(部門ごと)
                    Call subErrorLog(strFilePath, strFilename, strErrLogFile(lngIndex), strProcData, strBumonCode(lngIndex), strRead, intError)
                    lngErrCnt(lngIndex) = lngErrCnt(lngIndex) + 1
                    lngErrSum = lngErrSum + 1
                End If
            End If
        End If
    Wend

    Close

    '結果ログ
    Call MSZZ003_M00(P_PROGRAM_ID, "8", _
        "  処理結果 読み込み=[" & lngReadCnt & "]件 " & _
        "全取込=[" & lngProcSum & "]件 " & _
        "全取込エラー=[" & lngErrSum + lngBumonUnmatchCnt + lngRcptNoBlankCnt & "]件")
    
    Call MSZZ003_M00(P_PROGRAM_ID, "8", _
        "  部門コード不一致=[" & lngBumonUnmatchCnt & "]件 " & _
        "受付番号空白=[" & lngRcptNoBlankCnt & "]件 ")
    
    '部門ごとの内訳ログ
    For lngIndex = 0 To UBound(strBumonCode) - 1
        Call MSZZ003_M00(P_PROGRAM_ID, "8", _
            "  部門[" & strBumonCode(lngIndex) & "]内訳 " & _
            "対象=[" & lngBumonSameCnt(lngIndex) & "]件 " & _
            "取込=[" & lngProcCnt(lngIndex) & "]件 " & _
            "取込エラー=[" & lngErrCnt(lngIndex) & "]件")
    Next
    

    '取込エラー時は１以上の値を返却
    fncProcTorikomiFile = lngErrSum + lngBumonUnmatchCnt + lngRcptNoBlankCnt

    Exit Function

ErrorHandler:
    Close
    Call Err.Raise(Err.Number, "fncProcTorikomiFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'       MODULE_NAME     : 受付トラン更新
'       MODULE_ID       : fncUpdateRcpt
'       CREATE_DATE     :
'       PARAM           : objDb                 DAO.Database
'                       : strProcData           取込データ
'                       : intError              取込エラー時にエラー番号格納
'                       : strUserId             更新ユーザーID
'                       : strMessage            エラー発生時にエラーメッセージを返却
'       RETURN          : 正常(True) エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateRcpt(objDb As Database, _
                               strProcData() As String, _
                               ByRef intError As Integer, _
                               strUserID As String) As Boolean
On Error GoTo ErrorHandler

    Dim blnError    As Boolean
    Dim rsData      As Recordset
    Dim blnAbout    As Boolean
    
    fncUpdateRcpt = False
    
    blnAbout = False
    
    '受付トランを受付番号で検索
    Set rsData = objDb.OpenRecordset(fncMakeRcptSql(Trim(strProcData(P_PROCDATA_受付番号))), dbOpenDynaset)

    If rsData.EOF Then
        '取込エラー　該当データ無し
        intError = P_ERROR_受付番号該当無し
        fncUpdateRcpt = True
        GoTo ExitRtn
    End If

    If Nz(rsData.Fields("RCPT_SLIP_NO")) <> "" And Nz(rsData.Fields("RCPT_SLIP_NO")) <> "-" Then
        '取込エラー　伝票番号取込済み
        intError = P_ERROR_既に取込済み
        strProcData(UBound(strProcData)) = Nz(rsData.Fields("RCPT_SLIP_NO"))
        fncUpdateRcpt = True
        GoTo ExitRtn
    End If
    
    '更新
    With rsData
        .Edit
        .Fields("RCPT_SLIP_NO") = Trim(strProcData(P_PROCDATA_伝票番号))
        .Fields("RCPT_DERIVERY_STATUS") = Trim(strProcData(P_PROCDATA_配送状況))
        If .Fields("RCPT_DERIVERY_STATUS") = P_DERIVERYSTATUS_配達完了 Then
            '配送状況が「配達完了」ならば到着日を更新する
            .Fields("RCPT_ARRIVAL_DATE") = Format(DATE, "yyyy/") & Trim(strProcData(P_PROCDATA_日付))
            .Fields("RCPT_PAYDATE") = .Fields("RCPT_ARRIVAL_DATE")
        End If
        .Fields("RCPT_UPDAD") = Format(DATE, "yyyymmdd")
        .Fields("RCPT_UPDAJ") = Format(Now, "hhmmss")
        .Fields("RCPT_UPDPB") = P_PROGRAM_ID
        .Fields("RCPT_UPDUB") = strUserID
        .UPDATE
        .Close
    End With

    Set rsData = Nothing
    
    fncUpdateRcpt = True
    GoTo ExitRtn

ErrorHandler:
    blnAbout = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing

    If blnAbout Then
        Call Err.Raise(Err.Number, "fncUpdateRcpt" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

Private Function fncMakeRcptSql(strRcptNo As String) As String

    Dim strSQL      As String
    
    strSQL = strSQL & " SELECT RCPT_SLIP_NO, "
    strSQL = strSQL & "        RCPT_DERIVERY_STATUS, "
    strSQL = strSQL & "        RCPT_ARRIVAL_DATE, "
    strSQL = strSQL & "        RCPT_PAYDATE, "
    strSQL = strSQL & "        RCPT_UPDAD, "
    strSQL = strSQL & "        RCPT_UPDAJ, "
    strSQL = strSQL & "        RCPT_UPDPB, "
    strSQL = strSQL & "        RCPT_UPDUB "
    strSQL = strSQL & "   FROM RCPT_TRAN "
    strSQL = strSQL & "  WHERE RCPT_NO = '" & strRcptNo & "'"
    
    fncMakeRcptSql = strSQL
    
End Function
'==============================================================================*
'
'       MODULE_NAME     : 取込エラーログ出力
'       MODULE_ID       : subErrorLog
'       CREATE_DATE     :
'       PARAM           : strFilePath           取込ファイルパス
'                       : strFileName           取込ファイル名
'                       : strErrLogFile         取込エラーログファイル名
'                       : strProcData           取込データ
'                       : strBumonCode          部門コード
'                       : strPlaneData          取込データ(ファイルから読み込んだ文字列)
'                       : intErrKind            取込エラー番号
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subErrorLog(strFilePath As String, _
                        strFilename As String, _
                        ByRef strErrLogFile As String, _
                        strProcData() As String, _
                        strBumonCode As String, _
                        strPlaneData As String, _
                        intErrKind As Integer)

    Dim strErrLogPath           As String
    Dim blnFirstError           As Boolean
    Dim strTemp                 As String
    
    '初回はエラーログファイル名が格納されていないので、新規ファイル作成
    If strErrLogFile = "" Then
        strErrLogPath = strFilePath & P_TORIKOMI_ERROR_LOG_DIR
        
        If Dir(strErrLogPath, vbDirectory) = "" Then
            'フォルダ作成
            Call MkDir(strErrLogPath)
        End If
        
        strErrLogFile = strErrLogPath & "\" & "取込Err" & Format(dtProcDate, "yyyymmddhhmmss_") & strBumonCode & "(" & strFilename & ").log"
        blnFirstError = True
    Else
        blnFirstError = False
    End If

    Open strErrLogFile For Append As #2
    
    '初回のみヘッダ部分を出力する
    If blnFirstError Then
        Print #2, "===================================================================="
        Print #2, " 部門コード=[" & strBumonCode & "]"
        Print #2, " 取込ファイル=[" & strFilePath & strFilename & "]"
        Print #2, "===================================================================="
        Print #2, "--------------------------------------------------------------"
    End If
    
    'エラー番号ごとのメッセージ
    Select Case intErrKind
    Case P_ERROR_受付番号空白
        strTemp = "≪受付番号が空白≫"
    Case P_ERROR_受付番号該当無し
        strTemp = "≪受付番号が受付データに該当無し≫"
    Case P_ERROR_既に取込済み
        strTemp = "≪既に伝票番号が取り込まれている≫"
    Case P_ERROR_部門コード不一致
        strTemp = "≪部門コードが一致しない≫"
    End Select
    
    Print #2, strTemp
    
    Print #2, "  <読み込みデータ> "
    Print #2, "    顧客コード=[" & strProcData(P_PROCDATA_顧客コード) & "]"
    Print #2, "    顧客名=[" & strProcData(P_PROCDATA_顧客名) & "]"
    Print #2, "    受付番号=[" & strProcData(P_PROCDATA_受付番号) & "]"
    Print #2, "    伝票番号=[" & strProcData(P_PROCDATA_伝票番号) & "]"
    Print #2, "    全て=[" & strPlaneData & "]"
    
    If intErrKind = P_ERROR_既に取込済み Then
        '取り込みデータ配列の最後尾に受付データの伝票番号が格納されている
        Print #2, ""
        Print #2, "  <受付データ> "
        Print #2, "    伝票番号=[" & strProcData(UBound(strProcData)) & "]"
    End If
    
    Print #2, "--------------------------------------------------------------"
    
    Close #2
    
End Sub

'==============================================================================*
'
'   ファイルバックアップ
'
'   引数：
'       strFileName：ファイルのパス
'       strFileName：ファイル名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFileBackup(strFilePath As String, strFilename As String)

    Dim strBackupFileName       As String
    Dim strBackupPath           As String

    strBackupPath = strFilePath & BACKUP_PATH
    
    'ファイルコピー
    If Dir(strBackupPath, vbDirectory) = "" Then
        'フォルダ作成
        Call MkDir(strBackupPath)
    End If
    
    strBackupFileName = strBackupPath & "\" & Format(Now, "yyyymmddhhnnss") & strFilename
    Call FileCopy(strFilePath & strFilename, strBackupFileName)

    '元ファイル削除
    Call Kill(strFilePath & strFilename)

End Sub


Private Function fncConnectDatabase(strBumonCode() As String, ByRef dbSQLServer() As DAO.Database)
On Error GoTo ErrorHandler
    
    Dim lngCount        As Long
    Dim strSqlserver    As String
    Dim strConnect      As String

    '対象部門全てのデータベースに接続する
    For lngCount = 0 To UBound(strBumonCode) - 1
        'SQL-Server名
        strSqlserver = fncGetSqlServer(strBumonCode(lngCount))
        
        '接続文字列取得
        strConnect = fncGetConnectString(strBumonCode(lngCount))
        
        'SQLサーバー接続
        Set dbSQLServer(lngCount) = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    Next

    Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "fncConnectDatabase" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

Private Function fncDisconnectDatabase(ByRef dbSQLServer() As DAO.Database)
On Error GoTo ErrorHandler
    
    Dim lngCount        As Long
    
    'データベース切断
    For lngCount = 0 To UBound(dbSQLServer) - 1
        If Not dbSQLServer(lngCount) Is Nothing Then
            dbSQLServer(lngCount).Close
            Set dbSQLServer(lngCount) = Nothing
        End If
    Next

    Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "fncDisconnectDatabase" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

Private Function fncGetIndex(strArray() As String, strSearch As String) As Integer
    
    Dim lngCount        As Integer
    
    
    For lngCount = 0 To UBound(strArray) - 1
        If strArray(lngCount) = strSearch Then
            Exit For
        End If
    Next
    
    If lngCount >= UBound(strArray) Then
        fncGetIndex = -1
    Else
        fncGetIndex = lngCount
    End If
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : SQLサーバー名取得
'       MODULE_ID       : fncGetBumonArr
'       CREATE_DATE     :
'       PARAM           : strBumonCode          部門コード(省略可)
'       RETURN          : SQLサーバー名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetSqlServer(Optional strBumonCode As String = "") As String
On Error GoTo ErrorHandler

    Dim strSqlserver    As String
    Dim strParam        As String
    
    strParam = "ODBC_DATA_SOURCE_NAME"
    If strBumonCode <> "" Then
        strParam = strParam & "_" & strBumonCode
    End If

    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & strParam & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("DlookUp", "SQL-Server名の設定不正。")
    End If

    fncGetSqlServer = strSqlserver
    
    Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetSqlServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function


'==============================================================================*
'
'       MODULE_NAME     : 接続文字列取得
'       MODULE_ID       : fncGetConnectString
'       CREATE_DATE     :
'       PARAM           : strBumonCode          部門コード(省略可)
'       RETURN          : 接続文字列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetConnectString(Optional strBumonCode As String = "") As String
On Error GoTo ErrorHandler

    Dim strConnectString    As String
    
    strConnectString = MSZZ007_M00(strBumonCode)
    If strConnectString = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "テーブル[SETU_TABL]の設定不正。")
    End If

    fncGetConnectString = strConnectString
    
    Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetConnectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

