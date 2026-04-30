Attribute VB_Name = "Print200"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　紹介制限ヤード状況出力
'   プログラムＩＤ　：　Print200
'   作　成　日　　　：  2007/02/17
'   作　成　者　　　：  イーグルソフト 鈴木
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :
'   UPDATER         :
'   Ver             :
'   変更内容        :
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'処理モード
Public Const P901_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P901_MODE_EXCEL                As Integer = 2  'Excelに出力
Public Const P901_MODE_PRINT                As Integer = 3  'プレビューを表示しないで印刷

Private Const P_SQL値_上段                  As Integer = 0
Private Const P_SQL値_下段                  As Integer = 1

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RVS200_W01"

'レポート名
Private Const P_REPORT                      As String = "RVS200"

Private pstrBumonCd                         As String        ' 部門コード
Private pstrBumonNm                         As String        ' 部門名

Sub a00Test_fncPrintYardInlimit()

    If Not fncPrintYardInlimit(P901_MODE_PREVIEW, "12") Then
        MsgBox "False"
    End If

'    If Not fncPrintYardInlimit(P901_MODE_EXCEL, "12") Then
'        MsgBox "False"
'    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ヤードメンテナンス依頼書出力
'       MODULE_ID       : fncPrintYardInlimit
'       CREATE_DATE     : 2007/02/17
'                       :
'       PARAM           : intMode       - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'                       : strYardCode   - ヤードコード'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncPrintYardInlimit(intMode As Integer, _
                                    strYardCode As String) As Boolean
On Error GoTo ErrorHandler

    Dim dbSQLServer     As Database
    Dim rsGetData       As Recordset
    Dim blnError        As Boolean

    blnError = False

    fncPrintYardInlimit = False

    'DB接続
    Call subConnectServer(dbSQLServer)

    '部門コード／部門名取得
    Call subGetBumonName

    'データ検索
    If Not fncGetData(dbSQLServer, rsGetData, strYardCode) Then
        '該当データ無し
        GoTo ExitRtn
    End If

    'ワークテーブル作成
    Call subMakeWork(rsGetData, intMode)

    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    '出力
    Select Case intMode
    Case P901_MODE_PREVIEW:
        'レポートプレビュー
        doCmd.OpenReport P_REPORT, acViewPreview
    Case P901_MODE_EXCEL:
        'EXCELファイル出力
        On Error Resume Next
        doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, , True
        On Error GoTo ErrorHandler
    Case P901_MODE_PRINT:
        'レポート印刷
        On Error Resume Next
        doCmd.OpenReport P_REPORT
        On Error GoTo ErrorHandler
    End Select

    fncPrintYardInlimit = True

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, "fncPrintYardInlimit" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :subClearWork
'        機能             :ワークテーブルクリア
'        IN               :dbAccess     - ACCESSデータベースオブジェクト(省略可)
'                         :strTableName - テーブル名(省略可)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub psubClearWork(Optional dbAccess As Database = Null, _
                         Optional strTable As String = P_WORK_TABLE)

On Error GoTo ErrorHandler

    Dim tdfNew      As TableDef
    Dim blnError    As Boolean
    Dim blnConnect  As Boolean

    blnError = False

    'データベースに未接続ならばCurrentDbに接続する
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
        blnConnect = True
    Else
        blnConnect = False
    End If

    'ワークテーブル削除
    If fncTableExist(dbAccess, strTable) Then
        Call dbAccess.TableDefs.Delete(strTable)
    End If

    'ワークテーブル作成
    Set tdfNew = dbAccess.CreateTableDef(strTable)
    Call subFieldAppend(tdfNew)
    Call dbAccess.TableDefs.Append(tdfNew)

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    If blnConnect And Not dbAccess Is Nothing Then
        dbAccess.Close
        Set dbAccess = Nothing
    End If

    If blnError Then
        Call Err.Raise(Err.Number, "psubClearWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : データ検索
'       MODULE_ID       : fncGetData
'       CREATE_DATE     : 2007/02/17
'                       :
'       PARAM           : dbSqlServer - KOMSに接続したデータベースオブジェクト
'                       : rsGetData   - 検索結果を格納するレコードセット
'                       : strYardCode - ヤードコード
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(dbSQLServer As Database, _
                            ByRef rsGetData As Recordset, _
                            strYardCode As String) As Boolean

On Error GoTo ErrorHandler
    
    Dim strSQL      As String

    fncGetData = False

    'SQL文作成
    strSQL = fncMakeGetDataSql(strYardCode)

    '検索
    Set rsGetData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    'データが存在しない場合Falseを返却
    fncGetData = Not rsGetData.EOF

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : SQL文作成
'       MODULE_ID       : fncMakeGetDataSql
'       CREATE_DATE     : 2007/02/17
'                       :
'                       : strYardCode - ヤードコード
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(strYardCode As String) As String

    Dim strSQL              As String

    strSQL = " SELECT"
    strSQL = strSQL & "     NYAR_YCODE " & Chr(13)
    strSQL = strSQL & "    ,NYAR_NCODE " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_NAME AS YARD_NAME01" & Chr(13)
    strSQL = strSQL & "    ,ROUND(NYAR_KIRO,2) AS NYAR_KIRO " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_ADDR_1 " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_ADDR_2 " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_ADDR_3 " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_RENTEND_DAY " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_END_DAY " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_INLIMIT_DAY " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_INLIMIT_YCODE " & Chr(13)
    strSQL = strSQL & "    ,INLIMIT_YARD.YARD_NAME AS YARD_NAME02 " & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_上段, 0.1, 1) & " AS UP_01_10_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_下段, 0.1, 1) & " AS DW_01_10_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_上段, 1.1, 2) & " AS UP_11_20_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_下段, 1.1, 2) & " AS DW_11_20_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_上段, 2.1, 3) & " AS UP_21_30_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_下段, 2.1, 3) & " AS DW_21_30_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_上段, 3.1, 4) & " AS UP_31_40_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_下段, 3.1, 4) & " AS DW_31_40_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_上段, 4.1, 8) & " AS UP_41_80_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_下段, 4.1, 8) & " AS DW_41_80_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_上段, 8.1, 99.9) & " AS UP_81_999_CNT" & Chr(13)
    strSQL = strSQL & "    ," & get段帖毎利用設置数Sql(P_SQL値_下段, 8.1, 99.9) & " AS DW_81_999_CNT" & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "    NYAR_MAST " & Chr(13)
    strSQL = strSQL & "    INNER JOIN CNTA_MAST ON " & Chr(13)
    strSQL = strSQL & "        CNTA_CODE = NYAR_NCODE " & Chr(13)
    strSQL = strSQL & "    INNER JOIN YARD_MAST ON " & Chr(13)
    strSQL = strSQL & "        YARD_MAST.YARD_CODE = NYAR_NCODE  " & Chr(13)
    strSQL = strSQL & "    LEFT OUTER JOIN YARD_MAST INLIMIT_YARD ON " & Chr(13)
    strSQL = strSQL & "        INLIMIT_YARD.YARD_CODE = YARD_MAST.YARD_INLIMIT_YCODE " & Chr(13)
    strSQL = strSQL & "    LEFT OUTER JOIN CARG_FILE ON " & Chr(13)
    strSQL = strSQL & "        CARG_YCODE = CNTA_CODE AND  CARG_NO = CNTA_NO AND  CARG_AGRE <> 9" & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "    NYAR_YCODE = '" & strYardCode & "' " & Chr(13)
    strSQL = strSQL & "    AND CNTA_USE <> 9 " & Chr(13)
    strSQL = strSQL & "    AND ISNULL(YARD_MAST.YARD_NONDISP_DAY,'9999/12/31') > GETDATE() " & Chr(13)
    strSQL = strSQL & "GROUP BY " & Chr(13)
    strSQL = strSQL & "    NYAR_YCODE "
    strSQL = strSQL & "   ,NYAR_NCODE "
    strSQL = strSQL & "   ,YARD_MAST.YARD_NAME "
    strSQL = strSQL & "   ,NYAR_KIRO  "
    strSQL = strSQL & "   ,YARD_MAST.YARD_ADDR_1 "
    strSQL = strSQL & "   ,YARD_MAST.YARD_ADDR_2 "
    strSQL = strSQL & "   ,YARD_MAST.YARD_ADDR_3 "
    strSQL = strSQL & "   ,YARD_MAST.YARD_RENTEND_DAY  "
    strSQL = strSQL & "   ,YARD_MAST.YARD_END_DAY  "
    strSQL = strSQL & "   ,YARD_MAST.YARD_INLIMIT_DAY  "
    strSQL = strSQL & "   ,YARD_MAST.YARD_INLIMIT_YCODE   "
    strSQL = strSQL & "   ,INLIMIT_YARD.YARD_NAME  "
    strSQL = strSQL & "ORDER BY " & Chr(13)
    strSQL = strSQL & "    NYAR_KIRO"

    fncMakeGetDataSql = strSQL

End Function

Private Function get段帖毎利用設置数Sql(ByVal a段区分 As Integer, _
                                        ByVal aSizeFrom As Double, _
                                        ByVal aSizeTo As Double _
                                        ) As String
    Dim sqlText As String
    
    sqlText = " CONVERT( VARCHAR,"
    sqlText = sqlText & " COUNT(CASE WHEN CNTA_STEP = " & a段区分
    sqlText = sqlText & "  AND " & aSizeFrom & " <= CNTA_SIZE AND CNTA_SIZE <= " & aSizeTo
    sqlText = sqlText & "  THEN CARG_YCODE ELSE NULL END) "
    sqlText = sqlText & ") + ' (' +  STR(CONVERT( VARCHAR, "
    sqlText = sqlText & " COUNT(CASE WHEN CNTA_STEP = " & a段区分
    sqlText = sqlText & "  AND " & aSizeFrom & " <= CNTA_SIZE AND CNTA_SIZE <= " & aSizeTo
    sqlText = sqlText & "  THEN CNTA_CODE  ELSE NULL END) "
    sqlText = sqlText & "),2) + ')' "
    
    get段帖毎利用設置数Sql = sqlText
                                        
End Function
'==============================================================================*
'
'       MODULE_NAME     : SQL文作成
'       MODULE_ID       : fncMakeGetDataSql
'       CREATE_DATE     : 2007/02/17
'                       :
'                       : strYardCode - ヤードコード
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql2(strYardCode As String) As String

    Dim strSQL              As String
    Dim strOpenRowSetSql    As String

    strSQL = " SELECT"
    strSQL = strSQL & "     NYAR_YCODE " & Chr(13)
    strSQL = strSQL & "    ,NYAR_NCODE " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_NAME AS YARD_NAME01" & Chr(13)
    strSQL = strSQL & "    ,NYAR_KIRO " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_ADDR_1 " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_ADDR_2 " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_ADDR_3 " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_RENTEND_DAY " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_END_DAY " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_INLIMIT_DAY " & Chr(13)
    strSQL = strSQL & "    ,YARD_MAST.YARD_INLIMIT_YCODE " & Chr(13)
    strSQL = strSQL & "    ,INLIMIT_YARD.YARD_NAME AS YARD_NAME02 " & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 0" & Chr(13)
    strSQL = strSQL & "        AND CNTA_SIZE between 0.1 AND 1.0" & Chr(13)
    strSQL = strSQL & "     ) UP_01_10_CNT" & Chr(13)
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 1" & Chr(13)
    strSQL = strSQL & "        AND CNTA_SIZE between 0.1 AND 1.0" & Chr(13)
    strSQL = strSQL & "     ) DW_01_10_CNT" & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 0" & Chr(13)
    strSQL = strSQL & "        AND CNTA_SIZE between 1.1 AND 2.0" & Chr(13)
    strSQL = strSQL & "     ) UP_11_20_CNT" & Chr(13)
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 1" & Chr(13)
    strSQL = strSQL & "        and CNTA_SIZE between 1.1 and 2.0" & Chr(13)
    strSQL = strSQL & "     ) DW_11_20_CNT" & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 0" & Chr(13)
    strSQL = strSQL & "        AND CNTA_SIZE between 2.1 AND 3.0" & Chr(13)
    strSQL = strSQL & "     ) UP_21_30_CNT" & Chr(13)
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 1" & Chr(13)
    strSQL = strSQL & "        and CNTA_SIZE between 2.1 and 3.0" & Chr(13)
    strSQL = strSQL & "     ) DW_21_30_CNT" & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 0" & Chr(13)
    strSQL = strSQL & "        AND CNTA_SIZE between 3.1 AND 4.0" & Chr(13)
    strSQL = strSQL & "     ) UP_31_40_CNT" & Chr(13)
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 1" & Chr(13)
    strSQL = strSQL & "        and CNTA_SIZE between 3.1 and 4.0" & Chr(13)
    strSQL = strSQL & "     ) DW_31_40_CNT" & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 0" & Chr(13)
    strSQL = strSQL & "        AND CNTA_SIZE between 4.1 AND 8.0" & Chr(13)
    strSQL = strSQL & "     ) UP_41_80_CNT" & Chr(13)
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 1" & Chr(13)
    strSQL = strSQL & "        and CNTA_SIZE between 4.1 and 8.0" & Chr(13)
    strSQL = strSQL & "     ) DW_41_80_CNT" & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 0" & Chr(13)
    strSQL = strSQL & "        AND CNTA_SIZE between 8.1 AND 99.9" & Chr(13)
    strSQL = strSQL & "     ) UP_81_999_CNT" & Chr(13)
    strSQL = strSQL & "    ,(SELECT" & Chr(13)
    strSQL = strSQL & "        CONVERT(VARCHAR,count(CARG_YCODE)) + '(' + STR(CONVERT(VARCHAR,count(CNTA_CODE)),2) + ')' CNTA_COUNT" & Chr(13)
    strSQL = strSQL & "      FROM" & Chr(13)
    strSQL = strSQL & "            CNTA_MAST" & Chr(13)
    strSQL = strSQL & "            LEFT OUTER JOIN CARG_FILE ON" & Chr(13)
    strSQL = strSQL & "                CNTA_MAST.CNTA_CODE = CARG_FILE.CARG_YCODE" & Chr(13)
    strSQL = strSQL & "            AND CNTA_MAST.CNTA_NO = CARG_FILE.CARG_NO" & Chr(13)
    strSQL = strSQL & "            AND 9 <> CARG_FILE.CARG_AGRE" & Chr(13)
    strSQL = strSQL & "      WHERE" & Chr(13)
    strSQL = strSQL & "            CNTA_CODE = NYAR_NCODE" & Chr(13)
    strSQL = strSQL & "        AND CNTA_USE <> 9" & Chr(13)
    strSQL = strSQL & "        AND CNTA_STEP = 1" & Chr(13)
    strSQL = strSQL & "        and CNTA_SIZE between 8.1 and 99.9" & Chr(13)
    strSQL = strSQL & "     ) DW_81_999_CNT" & Chr(13)
    ' ------------------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "    NYAR_MAST " & Chr(13)
    strSQL = strSQL & "    INNER JOIN YARD_MAST ON " & Chr(13)
    strSQL = strSQL & "        YARD_MAST.YARD_CODE = NYAR_NCODE " & Chr(13)
    strSQL = strSQL & "    AND ISNULL(YARD_MAST.YARD_NONDISP_DAY,'9999/12/31') > GETDATE() " & Chr(13)
    strSQL = strSQL & "    LEFT OUTER JOIN YARD_MAST AS INLIMIT_YARD ON " & Chr(13)
    strSQL = strSQL & "        INLIMIT_YARD.YARD_CODE = YARD_MAST.YARD_INLIMIT_YCODE " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "    NYAR_YCODE = '" & strYardCode & "' " & Chr(13)
    strSQL = strSQL & "ORDER BY " & Chr(13)
    strSQL = strSQL & "    NYAR_KIRO"

    fncMakeGetDataSql2 = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncTableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :dbAccess     - ACCESSデータベースオブジェクト
'                         :strTableName - テーブル名
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTableExist(dbAccess As Database, strTableName As String) As Boolean

    Dim tdf As TableDef
    
    fncTableExist = False
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            fncTableExist = True
            Exit For
        End If
    Next tdf

End Function

'==============================================================================*
'
'        MODULE_NAME      :subFieldAppend
'        機能             :ワークテーブル列作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew
 
        Call .Fields.Append(.CreateField("部門コード", DataTypeEnum.dbText, 36))        '部門コード
        Call .Fields.Append(.CreateField("部門名", DataTypeEnum.dbText, 36))            '部門名
        Call .Fields.Append(.CreateField("区分", DataTypeEnum.dbText, 14))              '区分
        Call .Fields.Append(.CreateField("メインヤードコード", DataTypeEnum.dbText, 6)) 'ヤードコード(メイン)
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))       'ヤードコード
        Call .Fields.Append(.CreateField("ヤード名", DataTypeEnum.dbText, 36))          'ヤード名
        Call .Fields.Append(.CreateField("ヤード住所", DataTypeEnum.dbText, 200))       'ヤード住所
        Call .Fields.Append(.CreateField("距離", DataTypeEnum.dbText, 20))              '距離
        Call .Fields.Append(.CreateField("上段01", DataTypeEnum.dbText, 20))            '上段01
        Call .Fields.Append(.CreateField("下段01", DataTypeEnum.dbText, 20))            '下段01
        Call .Fields.Append(.CreateField("上段02", DataTypeEnum.dbText, 20))            '上段02
        Call .Fields.Append(.CreateField("下段02", DataTypeEnum.dbText, 20))            '下段02
        Call .Fields.Append(.CreateField("上段03", DataTypeEnum.dbText, 20))            '上段03
        Call .Fields.Append(.CreateField("下段03", DataTypeEnum.dbText, 20))            '下段03
        Call .Fields.Append(.CreateField("上段04", DataTypeEnum.dbText, 20))            '上段04
        Call .Fields.Append(.CreateField("下段04", DataTypeEnum.dbText, 20))            '下段04
        Call .Fields.Append(.CreateField("上段05", DataTypeEnum.dbText, 20))            '上段05
        Call .Fields.Append(.CreateField("下段05", DataTypeEnum.dbText, 20))            '下段05
        Call .Fields.Append(.CreateField("上段06", DataTypeEnum.dbText, 20))            '上段06
        Call .Fields.Append(.CreateField("下段06", DataTypeEnum.dbText, 20))            '下段06
        Call .Fields.Append(.CreateField("営業終了日", DataTypeEnum.dbText, 10))        '営業終了日
        Call .Fields.Append(.CreateField("ヤード解約予定", DataTypeEnum.dbText, 10))    'ヤード解約予定
        Call .Fields.Append(.CreateField("制限", DataTypeEnum.dbText, 2))               '制限
        Call .Fields.Append(.CreateField("備考", DataTypeEnum.dbText, 255))             '備考

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subMakeWork
'        機能             :ワークテーブルデータ追加
'        IN               :rsSource    - 検索結果が格納されたレコードセット
'                         :intMode     - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subMakeWork(rsSource As Recordset, intMode As Integer)

On Error GoTo ErrorHandler

    Dim dbAccess        As Database
    Dim rsDestination   As Recordset
    Dim blnError        As Boolean
    Dim intLoopCount    As Integer

    blnError = False
    intLoopCount = 0

    Set dbAccess = CurrentDb

    'ワークテーブルクリア
    Call psubClearWork(dbAccess, P_WORK_TABLE)

    'ワークテーブルのレコードセット
    Set rsDestination = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)

    'データ追加
    While Not rsSource.EOF
        intLoopCount = intLoopCount + 1
        Call subAddNew(rsSource, rsDestination, intMode, intLoopCount)
        rsSource.MoveNext
    Wend

    GoTo EndRtn

ErrorHandler:
    blnError = True

EndRtn:
    If Not rsDestination Is Nothing Then rsDestination.Close: Set rsDestination = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, "subMakeWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subAddNew
'        機能             :ワークテーブルAddNew
'        IN               :rsSource      - 検索結果が格納されたレコードセット
'                         :rsDestination - ワークテーブルのレコードセット
'                         :intMode     - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subAddNew(rsSource As Recordset, rsDestination As Recordset, intMode As Integer, intLoopCount As Integer)

    Dim strTemp     As String

    With rsSource
        rsDestination.AddNew

        rsDestination.Fields("部門コード") = pstrBumonCd                                                   '部門コード
        rsDestination.Fields("部門名") = pstrBumonNm                                                       '部門名

        If intLoopCount = 1 Then
            rsDestination.Fields("区分") = "1"                                                             'データ区分＝'1'⇒解約ヤード

            rsDestination.Fields("営業終了日") = .Fields("YARD_RENTEND_DAY")                               '営業終了日
            rsDestination.Fields("ヤード解約予定") = .Fields("YARD_END_DAY")                               'ヤード解約予定
            
        Else
            rsDestination.Fields("区分") = "2"                                                             'データ区分＝'1'⇒近隣ヤード情報

            If Nz(.Fields("YARD_INLIMIT_DAY"), "") = "" Then
                rsDestination.Fields("備考") = ""                                                              '備考
            Else
                rsDestination.Fields("備考") = .Fields("YARD_INLIMIT_DAY") & "まで" & Chr(13) & Chr(10) & _
                                               Format(.Fields("YARD_INLIMIT_YCODE"), "000000") & Space(1) & _
                                               .Fields("YARD_NAME02") & "にて停止"                             '備考
            End If
        End If

        rsDestination.Fields("メインヤードコード") = Format(.Fields("NYAR_YCODE"), "000000")               'ヤードコード＝近隣ヤードマスタ．ヤードコード
        rsDestination.Fields("ヤードコード") = Format(.Fields("NYAR_NCODE"), "000000")                     'ヤードコード＝近隣ヤードマスタ．近隣ヤードコード
        rsDestination.Fields("ヤード名") = .Fields("YARD_NAME01")                                          'ヤードマスタ.ヤード名
        rsDestination.Fields("ヤード住所") = .Fields("YARD_ADDR_1") & .Fields("YARD_ADDR_2") & .Fields("YARD_ADDR_3")   'ヤード住所＝ヤードマスタ．ヤード住所１ & ヤードマスタ．ヤード住所２ & ヤードマスタ．ヤード住所３
        rsDestination.Fields("距離") = .Fields("NYAR_KIRO")                                                '近隣ヤードマスタ.距離

        rsDestination.Fields("上段01") = .Fields("UP_01_10_CNT")                                           '
        rsDestination.Fields("下段01") = .Fields("DW_01_10_CNT")                                           '
        rsDestination.Fields("上段02") = .Fields("UP_11_20_CNT")                                           '
        rsDestination.Fields("下段02") = .Fields("DW_11_20_CNT")                                           '
        rsDestination.Fields("上段03") = .Fields("UP_21_30_CNT")                                           '
        rsDestination.Fields("下段03") = .Fields("DW_21_30_CNT")                                           '
        rsDestination.Fields("上段04") = .Fields("UP_31_40_CNT")                                           '
        rsDestination.Fields("下段04") = .Fields("DW_31_40_CNT")                                           '
        rsDestination.Fields("上段05") = .Fields("UP_41_80_CNT")                                           '
        rsDestination.Fields("下段05") = .Fields("DW_41_80_CNT")                                           '
        rsDestination.Fields("上段06") = .Fields("UP_81_999_CNT")                                          '
        rsDestination.Fields("下段06") = .Fields("DW_81_999_CNT")                                          '

        If Nz(.Fields("YARD_INLIMIT_DAY"), "") = "" Then
            rsDestination.Fields("制限") = "無"                                                            '制限＝「無」
        Else
            rsDestination.Fields("制限") = "有"                                                            '制限＝「有」
        End If

        rsDestination.UPDATE
    End With

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncStrToDate
'        機能             :YYYYMMDD文字列を日付型に変換
'        IN               :YYYYMMDD文字列
'        OUT              :日付型に変換した結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncStrToDate(strYyyyMmDd As Variant) As Variant

    If Nz(strYyyyMmDd) = "" Then
        '空ならばNULLを返却
        fncStrToDate = Null
    Else
        fncStrToDate = DateSerial(Left(strYyyyMmDd, 4), Mid(strYyyyMmDd, 5, 2), Right(strYyyyMmDd, 2))
    End If

End Function

'==============================================================================*
'
'       MODULE_NAME     : KOMSデータベース接続
'       MODULE_ID       : subConnectServer
'       CREATE_DATE     :
'       PARAM           : データベースオブジェクト
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectServer(ByRef dbSQLServer As Database)

On Error GoTo ErrorHandler

    Dim strSqlserver    As String
    Dim strConnect      As String
    Dim strBUMOC        As String

    '部門コード
    strBUMOC = fncGetBumonCode()

    'SQL-Server名
    strSqlserver = fncGetSqlServer(strBUMOC)

    '接続文字列取得
    strConnect = fncGetConnectString(strBUMOC)

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subConnectServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 部門コード取得
'       MODULE_ID       : fncGetBumonCode
'       CREATE_DATE     :
'       PARAM           :
'       RETURN          : 部門コード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBumonCode() As String

On Error GoTo ErrorHandler

    Dim strBumonCode        As String

    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    fncGetBumonCode = strBumonCode

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetBumonCode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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
        Call MSZZ024_M10("MSZZ007_M00", "接続文字列の設定不正。")
    End If

    fncGetConnectString = strConnectString

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetConnectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :subGetBumonName
'        CREATE_DATE      :2007/02/17
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetBumonName()

    Dim strSQL  As String
    Dim objRs   As Recordset
    Dim objDb   As Database

    On Error GoTo ErrorHandler

    Set objDb = CurrentDb

    strSQL = "SELECT "
    strSQL = strSQL & "BUMOM_BUMOC, "  ' 部門コード
    strSQL = strSQL & "BUMOM_BUMON "   ' 部門名称
    strSQL = strSQL & "FROM dbo_CONT_MAST INNER JOIN BUMO_MAST ON "
    strSQL = strSQL & "dbo_CONT_MAST.CONT_BUMOC = BUMO_MAST.BUMOM_BUMOC "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            pstrBumonCd = .Fields("BUMOM_BUMOC")
            pstrBumonNm = .Fields("BUMOM_BUMON")
        Else
            pstrBumonCd = Null
            pstrBumonNm = Null
        End If
    End With

subGetBumonName_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subGetBumonName" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    GoTo subGetBumonName_Exit
End Sub
