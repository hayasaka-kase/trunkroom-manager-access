Attribute VB_Name = "Print231"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：  一時使用契約変更承諾書出力
'   プログラムＩＤ　：　Print231
'   作　成　日　　　：  2010/02/03
'   作　成　者　　　：  M.RYU
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :   2010/02/04
'   UPDATER         :   M.Honda
'   Ver             :   0.1
'   変更内容        :   データの取得条件を変更
'
'   UPDATE          :   2010/02/10
'   UPDATER         :   M.Honda
'   Ver             :   0.2
'   変更内容        :   移動区分の条件を削除
'
'   UPDATE          :   2010/02/10
'   UPDATER         :   M.RYU
'   Ver             :   0.3
'   変更内容        :   変更前月額使用料を契約月額に修正
'
'   UPDATE          :   2010/07/03
'   UPDATER         :   M.RYU
'   Ver             :   0.4
'   変更内容        :   変更前契約№・変更後契約№を追加
'
'   UPDATE          :   2013/03/27
'   UPDATER         :   M.HONDA
'   Ver             :   0.5
'   変更内容        :   変更前他金額を表示するように修正
'
'   UPDATE          :   2015/08/05
'   UPDATER         :   M.HONDA
'   Ver             :   0.6
'   変更内容        :   金額の取得方法を修正
'
'   UPDATE          :   2018/09/21
'   UPDATER         :   EGL
'   Ver             :   0.7
'   変更内容        :   階数表示(トランクルーム時)の対応
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'処理モード
Public Const P231_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P231_MODE_EXCEL                As Integer = 2  'Excelに出力

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RKS231_W01"

'レポート名
Private Const P_REPORT                      As String = "RKS231"

'==============================================================================*
'
'       MODULE_NAME     : 一時使用契約変更承諾書出力
'       MODULE_ID       : PrintUserMoveList
'       CREATE_DATE     : 2010/02/03
'                       :
'       PARAM           : intMode          - 1=印刷プレビュー 2=Excel出力
'                       : strYardCodeFrom  - ヤードコードFrom（省略可）
'                       : strYardCodeTo    - ヤードコードTo（省略可）
'                       : strUserCodeFrom  - ユーザーコードFrom（省略可）
'                       : strUserCodeTo    - ユーザーコードTo（省略可）
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function PrintUserMoveList(intMode As Integer, _
                                  Optional strYardCodeFrom As String = "", _
                                  Optional strYardCodeTo As String = "", _
                                  Optional strUserCodeFrom As String = "", _
                                  Optional strUserCodeTo As String = "" _
                                  ) As Boolean

    Dim rsGetData       As Object
    Dim blnError        As Boolean
    Dim adoDbConnection As Object

On Error GoTo ErrorHandler

    blnError = False
    PrintUserMoveList = False

    'DB接続
    Call subConnectServer(adoDbConnection)

    'データ検索
    If Not fncGetData(adoDbConnection, rsGetData, strYardCodeFrom, strYardCodeTo, strUserCodeFrom, strUserCodeTo) Then
        '該当データ無し
        GoTo ExitRtn
    End If

    'ワークテーブル作成
    Call subMakeWork(adoDbConnection, rsGetData, intMode)

    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing

    '出力
    Select Case intMode
        Case P231_MODE_PREVIEW:
            'レポートプレビュー
            doCmd.OpenReport P_REPORT, acViewPreview

        Case P231_MODE_EXCEL:
            'EXCELファイル出力
            On Error Resume Next
            doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, , True
            On Error GoTo ErrorHandler
    End Select

    PrintUserMoveList = True

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "PrintUserMoveList" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : データ検索
'       MODULE_ID       : fncGetData
'       CREATE_DATE     : 2010/02/03
'                       :
'       PARAM           : aConnection      - データベースオブジェクト
'                       : rsGetData        - 検索結果を格納するレコードセット
'                       : strYardCodeFrom  - ヤードコードFrom
'                       : strYardCodeTo    - ヤードコードTo
'                       : strUserCodeFrom  - ユーザーコードFrom
'                       : strUserCodeTo    - ユーザーコードTo
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(aConnection As Object, _
                            ByRef rsGetData As Object, _
                            strYardCodeFrom As String, _
                            strYardCodeTo As String, _
                            strUserCodeFrom As String, _
                            strUserCodeTo As String _
                            ) As Boolean

    Dim strSQL      As String
    Dim rsData      As Object

On Error GoTo ErrorHandler

    fncGetData = False

    'メインSQL文作成
    strSQL = fncMakeGetDataSql(strYardCodeFrom, strYardCodeTo, strUserCodeFrom, strUserCodeTo)

    ' レコードセット作成
    Set rsGetData = MSZZ025.ADODB_Recordset(strSQL, aConnection)

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
'       CREATE_DATE     : 2010/02/03
'                       :
'       PARAM           : strYardCodeFrom  - ヤードコードFrom
'                       : strYardCodeTo    - ヤードコードTo
'                       : strUserCodeFrom  - ユーザーコードFrom
'                       : strUserCodeTo    - ユーザーコードTo
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(strYardCodeFrom As String, _
                            strYardCodeTo As String, _
                            strUserCodeFrom As String, _
                            strUserCodeTo As String _
                                  ) As String

    'KASE_DB名前を取得
    Dim strKASEDBN As String
    strKASEDBN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME'")
    strKASEDBN = strKASEDBN & ".dbo."

    Dim strSQL              As String
    strSQL = " SELECT * " & Chr(13)
    strSQL = strSQL & " FROM " & Chr(13)

    ' 変更前データ取得SQL
    ' 【SELECT句】
    strSQL = strSQL & " ( " & Chr(13)
    strSQL = strSQL & " SELECT  CNTA_MAST.CNTA_USAGE  AS CNTA_USAGE_B   " & Chr(13)
    strSQL = strSQL & "        ,NAME_MAST.NAME_NAME   AS NAME_NAME_B    " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_YCODE  AS CARG_YCODE_B   " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_NAME   AS YARD_NAME_B    " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_NO     AS CARG_NO_B      " & Chr(13)
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_STEP   AS CNTA_STEP_B    " & Chr(13)
    
' ▼ 2018/09/21 egl add
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_FLOOR  AS CNTA_FLOOR_B  " & Chr(13)
' ▲ 2018/09/21 egl add
    
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_SIZE   AS CNTA_SIZE_B    " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_1 AS YARD_ADDR_1_B " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_2 AS YARD_ADDR_2_B " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_ACPTNO AS CARG_ACPTNO_B " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_UCODE  AS CARG_UCODE_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_NAME   AS USER_NAME_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_TANM   AS USER_TANM_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_YUBINO AS USER_YUBINO_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_ADR_1  AS USER_ADR_1_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_ADR_2  AS USER_ADR_2_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_ADR_3  AS USER_ADR_3_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_TEL    AS USER_TEL_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_FAX    AS USER_FAX_B " & Chr(13)
    strSQL = strSQL & "        ,USER_MAST.USER_KEITAI AS USER_KEITAI_B " & Chr(13)
    '2015/08/05 N.HONDA UPD
    '' 2013/03/27 M.HONDA START
    'strSQL = strSQL & "        ,RCPT_TRAN.RCPT_RENTKG + RCPT_EWARIBIKI AS RCPT_RENTKG_B " & Chr(13)
    '' strSQL = strSQL & "        ,RCPT_TRAN.RCPT_RENTKG - RCPT_EWARIBIKI AS RCPT_RENTKG_B " & Chr(13)
    '' 2013/03/27 M.HONDA END
    strSQL = strSQL & "        , (CASE "
    strSQL = strSQL & "            WHEN DCRAT_TO = '99999999'  THEN RCPT_TRAN.RCPT_RENTKG + RCPT_EWARIBIKI"
    strSQL = strSQL & "               Else RCPT_TRAN.RCPT_RENTKG END) AS RCPT_RENTKG_B " & Chr(13)
    '2015/08/05 N.HONDA UPD
    
    strSQL = strSQL & "        ,CARG_FILE.CARG_RENTKG AS CARG_RENTKG_B " & Chr(13)
    strSQL = strSQL & "        ,CARG_FILE.CARG_SYOZEI AS CARG_SYOZEI_B " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_EZAPPI AS RCPT_EZAPPI_B " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_ADD_EZAPPI1 AS RCPT_ADD_EZAPPI1_B " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_ADD_EZAPPI2 AS RCPT_ADD_EZAPPI2_B " & Chr(13)
    strSQL = strSQL & " FROM (((((( CARG_FILE INNER JOIN YARD_MAST ON CARG_FILE.CARG_YCODE = YARD_MAST.YARD_CODE ) " & Chr(13)  ' ヤードマスタ
    strSQL = strSQL & "   INNER JOIN CNTA_MAST ON ( CARG_FILE.CARG_YCODE = CNTA_MAST.CNTA_CODE )    " & Chr(13)
    strSQL = strSQL & "                       AND ( CARG_FILE.CARG_NO    = CNTA_MAST.CNTA_NO ) )    " & Chr(13)                  ' コンテナマスタ
    strSQL = strSQL & "   INNER JOIN USER_MAST ON CARG_FILE.CARG_UCODE   = USER_MAST.USER_CODE )    " & Chr(13)                  ' ユーザーマスタ
    strSQL = strSQL & "    LEFT JOIN RCPT_TRAN ON CARG_FILE.CARG_UKNO    = RCPT_TRAN.RCPT_NO   )    " & Chr(13)                  ' RCPT_TRAN
    strSQL = strSQL & "   INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE   = NAME_MAST.NAME_CODE )    " & Chr(13)                  ' NAME_MAST
    
    '2015/08/05 M.HONDA INS
    strSQL = strSQL & "   LEFT JOIN DCRA_TRAN ON DCRA_TRAN.DCRAT_ACPTNO = CARG_FILE.CARG_ACPTNO AND  " & Chr(13)
    strSQL = strSQL & "        DCRAT_SEIKYU_KBN = '1')                                               " & Chr(13)
    '2015/08/05 M.HONDA INS



    '' 2010/02/04 Ver0.1 start
    '' strSQL = strSQL & " WHERE CARG_FILE.CARG_AGRE <> 9 AND NAME_MAST.NAME_ID = '086' " & Chr(13)
    strSQL = strSQL & " WHERE NAME_MAST.NAME_ID = '086' " & Chr(13)
    '' strSQL = strSQL & "   AND ISNULL(CARG_FILE.CARG_KYDATE,'1900/12/31') >= GETDATE()   " & Chr(13) '解約日>=当日
    '' 2010/02/04 Ver0.1 end
    'Where句作成
    'ヤードコードの範囲条件
    strSQL = strSQL & fncMakeBetween("CARG_FILE.CARG_YCODE", strYardCodeFrom, strYardCodeTo)
    'ユーザーコードの条件指定
    strSQL = strSQL & fncMakeBetween("CARG_FILE.CARG_UCODE", strUserCodeFrom, strUserCodeTo)
    strSQL = strSQL & " ) Before " & Chr(13)  ' 変更前

    ' 変更後データ取得SQL
    ' 【SELECT句】
    strSQL = strSQL & " , " & Chr(13)
    strSQL = strSQL & " ( " & Chr(13)
    strSQL = strSQL & " SELECT  CNTA_MAST.CNTA_USAGE  AS CNTA_USAGE_A   " & Chr(13)
    strSQL = strSQL & "        ,NAME_MAST.NAME_NAME   AS NAME_NAME_A    " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_YCODE  AS RCPT_YCODE_A   " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_NAME   AS YARD_NAME_A    " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_CNO    AS RCPT_CNO_A     " & Chr(13)
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_STEP   AS CNTA_STEP_A    " & Chr(13)
    
' ▼ 2018/09/21 egl add
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_FLOOR  AS CNTA_FLOOR_A  " & Chr(13)
' ▲ 2018/09/21 egl add
    
    strSQL = strSQL & "        ,CNTA_MAST.CNTA_SIZE   AS CNTA_SIZE_A    " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_1 AS YARD_ADDR_1_A  " & Chr(13)
    strSQL = strSQL & "        ,YARD_MAST.YARD_ADDR_2 AS YARD_ADDR_2_A  " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_RENTKG AS RCPT_RENTKG_A  " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_EZAPPI AS RCPT_EZAPPI_A  " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_ADD_EZAPPI1  AS RCPT_ADD_EZAPPI1_A  " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_ADD_EZAPPI2  AS RCPT_ADD_EZAPPI2_A  " & Chr(13)
    strSQL = strSQL & "        ,YOUK_TRAN.YOUKT_MOTO_ACPTNO AS YOUKT_MOTO_ACPTNO_A " & Chr(13)
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_KISAN_DATE AS RCPT_KISAN_DATE_A     " & Chr(13) '契約開始日⇒起算日
    strSQL = strSQL & "        ,YOUK_TRAN.YOUKT_MOVEKBN AS YOUKT_MOVEKBN_A         " & Chr(13) '移動区分⇒自社理由移動='1'
    strSQL = strSQL & "        ,RCPT_TRAN.RCPT_CARG_ACPTNO AS CARG_ACPTNO_A        " & Chr(13)  '--20100703--ryu--add--
    strSQL = strSQL & "   FROM ((((( RCPT_TRAN INNER JOIN YARD_MAST ON RCPT_TRAN.RCPT_YCODE = YARD_MAST.YARD_CODE ) " & Chr(13)  ' ヤードマスタ
    strSQL = strSQL & "   INNER JOIN CNTA_MAST ON ( RCPT_TRAN.RCPT_YCODE = CNTA_MAST.CNTA_CODE )    " & Chr(13)
    strSQL = strSQL & "                       AND ( RCPT_TRAN.RCPT_CNO   = CNTA_MAST.CNTA_NO ) )    " & Chr(13)                  ' コンテナマスタ
    strSQL = strSQL & "   INNER JOIN USER_MAST ON RCPT_TRAN.RCPT_UCODE   = USER_MAST.USER_CODE )    " & Chr(13)                  ' ユーザーマスタ
    strSQL = strSQL & "   INNER JOIN YOUK_TRAN ON RCPT_TRAN.RCPT_NO      = YOUK_TRAN.YOUKT_UKNO)    " & Chr(13)                  ' 予約受付トラン
    strSQL = strSQL & "   INNER JOIN NAME_MAST ON CNTA_MAST.CNTA_USAGE   = NAME_MAST.NAME_CODE )    " & Chr(13)                  ' NAME_MAST
    strSQL = strSQL & "   LEFT JOIN CARG_FILE ON RCPT_TRAN.RCPT_NO = CARG_FILE.CARG_UKNO"
    strSQL = strSQL & " WHERE NAME_MAST.NAME_ID = '086' " & Chr(13)
    strSQL = strSQL & " ) After " & Chr(13)  ' 変更後
    '' 2010/02/10 Ver0.2 Honda START
    '' strSQL = strSQL & " WHERE Before.CARG_ACPTNO_B = After.YOUKT_MOTO_ACPTNO_A AND After.YOUKT_MOVEKBN_A = '1'"  ' 変更前と変更後の結合条件
    strSQL = strSQL & " WHERE Before.CARG_ACPTNO_B = After.YOUKT_MOTO_ACPTNO_A"  ' 変更前と変更後の結合条件
    '' 2010/02/10 Ver0.2 Honda END

    fncMakeGetDataSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeBetween
'        機能             :範囲条件作成
'        IN               :第一引数－対象テーブルのカラム名
'                         :第二引数－範囲条件値FROM
'                         :第三引数－範囲条件値TO
'        OUT              :条件文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeBetween(strColName As String, strfrom As String, strTo As String) As String

    Dim strTemp     As String
    
    strTemp = ""
    
    If strfrom <> "" And strTo <> "" Then
        '共に空白ではない場合、
        If strfrom = strTo Then
            'FROMとTOが同一の場合、一致条件
            strTemp = " AND " & strColName & " = '" & strfrom & "' "
        Else
            'FROMとTOが異なる場合、BETWEEN条件
            strTemp = " AND " & strColName & " BETWEEN '" & strfrom & "' AND '" & strTo & "' "
        End If
    ElseIf strfrom <> "" Then
        'FROMのみの場合、それ以上であることが条件
        strTemp = " AND " & strColName & " >= '" & strfrom & "' "
    ElseIf strTo <> "" Then
        'TOのみの場合、それ以下であることが条件
        strTemp = " AND " & strColName & " <= '" & strTo & "' "
    End If
    
    fncMakeBetween = strTemp

End Function

'==============================================================================*
'
'        MODULE_NAME      :subMakeWork
'        機能             :ワークテーブルデータ追加
'        CREATE_DATE      : 2010/02/03
'        IN               :rsSource    - 検索結果が格納されたレコードセット
'                         :intMode     - 1=印刷プレビュー 2=Excel出力
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subMakeWork(aConnection As Object, rsSource As Object, intMode As Integer)

    Dim strSQL          As String
    Dim dbAccess        As Database
    Dim rsDestination   As Recordset
    Dim blnError        As Boolean
    Dim strSize         As String
    Dim strBUMOC        As String

On Error GoTo ErrorHandler

    blnError = False
    
' ▼ 2018/09/21 egl add
    strBUMOC = fncGetBumonCode()
' ▲ 2018/09/21 egl add

    Set dbAccess = CurrentDb

    'ワークテーブルクリア
    Call subClearWork(dbAccess, P_WORK_TABLE)

    'ワークテーブルのレコードセット
    Set rsDestination = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)

    'データ追加
    With rsSource
        While Not .EOF

            ' AddNew
            rsDestination.AddNew

            ' *** 変更前 *** (16項目)
            rsDestination.Fields("変更前商品") = .Fields("NAME_NAME_B")
            rsDestination.Fields("変更前物件") = Format(.Fields("CARG_YCODE_B"), "000000")   ' ヤードコード
            rsDestination.Fields("変更前物件名称") = .Fields("YARD_NAME_B")
            rsDestination.Fields("変更前使用物件") = Format(.Fields("CARG_NO_B"), "000000")  ' コンテナ番号
            rsDestination.Fields("変更前物件住所") = .Fields("YARD_ADDR_1_B") & .Fields("YARD_ADDR_2_B")

'-----------20100220----M.Ryu-----del-------<s>
'            If Nz(.Fields("RCPT_RENTKG_B"), "") <> "" Then
'                rsDestination.Fields("変更前月額使用料・賃料") = "\" & Format(.Fields("RCPT_RENTKG_B"), "#,##0")
'                rsDestination.Fields("変更前他月額料") = "\" & Format(Nz(.Fields("RCPT_EZAPPI_B"), 0) + Nz(.Fields("RCPT_ADD_EZAPPI1_B"), 0) + Nz(.Fields("RCPT_ADD_EZAPPI2_B"), 0), "#,##0")
'            Else
            '' 20130327 M.HONDA START
            '' 変更前他月額料を表示するように修正
            rsDestination.Fields("変更前月額使用料・賃料") = "\" & Format(Nz(.Fields("RCPT_RENTKG_B"), 0), "#,##0")
            rsDestination.Fields("変更前他月額料") = "\" & Format(Nz(.Fields("RCPT_EZAPPI_B"), 0), "#,##0")

'            rsDestination.Fields("変更前月額使用料・賃料") = "\" & Format(Nz(.Fields("CARG_RENTKG_B"), 0) + Nz(.Fields("CARG_SYOZEI_B"), 0), "#,##0")
'            rsDestination.Fields("変更前他月額料") = "\0"
'            End If
            '' 20130327 M.HONDA END

'-----------20100220----M.Ryu-----del-------<e>

            rsDestination.Fields("変更前顧客No") = Format(.Fields("CARG_UCODE_B"), "000000")  ' 顧客コード
            rsDestination.Fields("変更前契約者氏名") = .Fields("USER_NAME_B")
            rsDestination.Fields("変更前法人代表者名") = .Fields("USER_TANM_B")
            rsDestination.Fields("変更前郵便番号") = .Fields("USER_YUBINO_B")
            rsDestination.Fields("変更前住所1") = .Fields("USER_ADR_1_B") & .Fields("USER_ADR_2_B")
            rsDestination.Fields("変更前住所2") = .Fields("USER_ADR_3_B")
            rsDestination.Fields("変更前電話番号") = .Fields("USER_TEL_B")
            rsDestination.Fields("変更前FAX番号") = .Fields("USER_FAX_B")
            rsDestination.Fields("変更前携帯番号") = .Fields("USER_KEITAI_B")
            
            ' 段区分名のみ表示
' ▼ 2018/09/21 egl dell
'            If Nz(.Fields("CNTA_STEP_B"), 0) = 0 Then
'               strSize = "上段"
'            Else
'               strSize = "下段"
'            End If
' ▲ 2018/09/21 egl dell
' ▼ 2018/09/21 egl add
            If strBUMOC = "H" Then
                'コンテナ
                If Nz(.Fields("CNTA_STEP_B"), 0) = 0 Then
                   strSize = "上段"
                Else
                   strSize = "下段"
                End If
            Else
                'コンテナ以外(階数を表示)
                strSize = MSZZ039.fncGetKaiName(Nz(.Fields("CNTA_FLOOR_B"), 0))
            End If
' ▲ 2018/09/21 egl add
            
            ' コンテナマスタ.実帖をセット
            strSize = strSize & "    " & Format(.Fields("CNTA_SIZE_B"), "0.00") & "帖"
            rsDestination.Fields("変更前サイズ") = strSize
            
            rsDestination.Fields("変更前契約№") = .Fields("CARG_ACPTNO_B") '--20100630--add-
      
            ' *** 変更後 *** (8項目)
            rsDestination.Fields("変更後商品") = .Fields("NAME_NAME_A")
            rsDestination.Fields("変更後物件") = Format(.Fields("RCPT_YCODE_A"), "000000")    ' ヤードコード
            rsDestination.Fields("変更後物件名称") = .Fields("YARD_NAME_A")
            rsDestination.Fields("変更後使用物件") = Format(.Fields("RCPT_CNO_A"), "000000")  ' コンテナ番号
            rsDestination.Fields("変更後物件住所") = .Fields("YARD_ADDR_1_A") & .Fields("YARD_ADDR_2_A")
            rsDestination.Fields("変更後月額使用料・賃料") = "\" & Format(Nz(.Fields("RCPT_RENTKG_A"), 0), "#,##0")
            rsDestination.Fields("変更後他月額料") = "\" & Format(Nz(.Fields("RCPT_EZAPPI_A"), 0) + Nz(.Fields("RCPT_ADD_EZAPPI1_A"), 0) + Nz(.Fields("RCPT_ADD_EZAPPI2_A"), 0), "#,##0")
            rsDestination.Fields("変更後起算日") = .Fields("RCPT_KISAN_DATE_A")
            
            ' 段区分名のみ表示
' ▼ 2018/09/21 egl dell
'            If Nz(.Fields("CNTA_STEP_A"), 0) = 0 Then
'               strSize = "上段"
'            Else
'               strSize = "下段"
'            End If
' ▲ 2018/09/21 egl dell
' ▼ 2018/09/21 egl add
            If strBUMOC = "H" Then
                'コンテナ
                If Nz(.Fields("CNTA_STEP_A"), 0) = 0 Then
                   strSize = "上段"
                Else
                   strSize = "下段"
                End If
            Else
                'コンテナ以外(階数を表示)
                strSize = MSZZ039.fncGetKaiName(Nz(.Fields("CNTA_FLOOR_A"), 0))
            End If
' ▲ 2018/09/21 egl add

            ' コンテナマスタ.実帖をセット
            strSize = strSize & "    " & Format(.Fields("CNTA_SIZE_A"), "0.00") & "帖"
            rsDestination.Fields("変更後サイズ") = strSize
            
            rsDestination.Fields("変更後契約№") = .Fields("CARG_ACPTNO_A") '--20100630--add-
            
            
            If .Fields("YOUKT_MOVEKBN_A") = 2 Then
                rsDestination.Fields("注意文言") = "尚、変更月の前月末までに移動を終わらせてください。終わらない場合には一ヶ月分の賃料を頂きます。"   '2016/09/17 M.HONDA INS
            End If
                       
            
            
            ' Update
            rsDestination.UPDATE

            .MoveNext
        Wend
    End With

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
'        MODULE_NAME      :subClearWork
'        機能             :ワークテーブルクリア
'        IN               :dbAccess     - ACCESSデータベースオブジェクト(省略可)
'                         :strTableName - テーブル名(省略可)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subClearWork(Optional dbAccess As Database = Null, _
                         Optional strTable As String = P_WORK_TABLE)

    Dim tdfNew      As TableDef
    Dim blnError    As Boolean
    Dim blnConnect  As Boolean

On Error GoTo ErrorHandler

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
        Call Err.Raise(Err.Number, "subClearWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

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
'        CREATE_DATE      : 2010/02/03
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew
 
        ' +++ 変更前 +++
        Call .Fields.Append(.CreateField("変更前商品", DataTypeEnum.dbText, 20))              '変更前商品
        Call .Fields.Append(.CreateField("変更前物件", DataTypeEnum.dbText, 6))               '変更前物件(ヤードコード)
        Call .Fields.Append(.CreateField("変更前物件名称", DataTypeEnum.dbText, 36))          '変更前物件名称(ヤード名)
        Call .Fields.Append(.CreateField("変更前使用物件", DataTypeEnum.dbText, 6))           '変更前使用物件(コンテナ番号)
        Call .Fields.Append(.CreateField("変更前物件住所", DataTypeEnum.dbText, 72))          '変更前物件住所1
        Call .Fields.Append(.CreateField("変更前月額使用料・賃料", DataTypeEnum.dbText, 50))  '変更前月額使用料・賃料
        Call .Fields.Append(.CreateField("変更前他月額料", DataTypeEnum.dbText, 50))          '変更前他月額料
        Call .Fields.Append(.CreateField("変更前顧客No", DataTypeEnum.dbText, 6))             '変更前顧客No
        Call .Fields.Append(.CreateField("変更前契約者氏名", DataTypeEnum.dbText, 36))        '変更前契約者氏名
        Call .Fields.Append(.CreateField("変更前法人代表者名", DataTypeEnum.dbText, 18))      '変更前法人代表者名
        Call .Fields.Append(.CreateField("変更前郵便番号", DataTypeEnum.dbText, 10))          '変更前郵便番号
        Call .Fields.Append(.CreateField("変更前住所1", DataTypeEnum.dbText, 72))             '変更前住所1
        Call .Fields.Append(.CreateField("変更前住所2", DataTypeEnum.dbText, 36))             '変更前住所2
        Call .Fields.Append(.CreateField("変更前電話番号", DataTypeEnum.dbText, 15))          '変更前電話番号
        Call .Fields.Append(.CreateField("変更前FAX番号", DataTypeEnum.dbText, 15))           '変更前FAX番号
        Call .Fields.Append(.CreateField("変更前携帯番号", DataTypeEnum.dbText, 15))          '変更前携帯番号
        Call .Fields.Append(.CreateField("変更前サイズ", DataTypeEnum.dbText, 50))            '変更前サイズ
        Call .Fields.Append(.CreateField("変更前契約№", DataTypeEnum.dbText, 20))            '変更前契約№--20100630--ryu--add--
        

        ' +++ 変更後 +++
        Call .Fields.Append(.CreateField("変更後商品", DataTypeEnum.dbText, 20))              '変更後商品
        Call .Fields.Append(.CreateField("変更後物件", DataTypeEnum.dbText, 6))               '変更後物件(ヤードコード)
        Call .Fields.Append(.CreateField("変更後物件名称", DataTypeEnum.dbText, 36))          '変更後物件名称(ヤード名)
        Call .Fields.Append(.CreateField("変更後使用物件", DataTypeEnum.dbText, 6))           '変更後使用物件(コンテナ番号)
        Call .Fields.Append(.CreateField("変更後物件住所", DataTypeEnum.dbText, 36))          '変更後物件住所
        Call .Fields.Append(.CreateField("変更後月額使用料・賃料", DataTypeEnum.dbText, 50))  '変更後月額使用料・賃料
        Call .Fields.Append(.CreateField("変更後他月額料", DataTypeEnum.dbText, 50))          '変更後他月額料
        Call .Fields.Append(.CreateField("変更後起算日", DataTypeEnum.dbText, 50))            '起算日
        Call .Fields.Append(.CreateField("変更後サイズ", DataTypeEnum.dbText, 50))            '変更後サイズ
        Call .Fields.Append(.CreateField("変更後契約№", DataTypeEnum.dbText, 20))            '変更後契約№--20100630--ryu--add--
        
        Call .Fields.Append(.CreateField("注意文言", DataTypeEnum.dbText, 100))               '注意文言 2016/09/17 M.HONDA INS
        
        
        
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With

End Sub

'==============================================================================*
'
'       MODULE_NAME     : データベース接続
'       MODULE_ID       : subConnectServer
'       CREATE_DATE     : 2010/02/03
'       PARAM           : データベースオブジェクト
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectServer(adoDbConnection As Object)

'    Dim adoDbConnection As Object       'ADODB.Connection
    Dim strBUMOC        As String

On Error GoTo ErrorHandler

    '部門コード取得
    strBUMOC = fncGetBumonCode()

    'ADO接続Object生成
    Set adoDbConnection = MSZZ025.ADODB_Connection(strBUMOC)

    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subConnectServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 部門コード取得
'       MODULE_ID       : fncGetBumonCode
'       CREATE_DATE     : 2010/02/03
'       PARAM           :
'       RETURN          : 部門コード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBumonCode() As String

    Dim strBumonCode        As String

On Error GoTo ErrorHandler

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
