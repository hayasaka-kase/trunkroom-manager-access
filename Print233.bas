Attribute VB_Name = "Print233"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：  委託料加瀬負担
'   プログラムＩＤ　：　Print233
'   作　成　日　　　：  2013/02/21
'   作　成　者　　　：  M.HONDA
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :   2013/07/16
'   UPDATER         :   M.HONDA
'   Ver             :   0.1
'   変更内容        :   初回請求方法を追加
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'処理モード
Public Const P233_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P233_MODE_EXCEL                As Integer = 2  'Excelに出力

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RKS233_W01"

'レポート名
Private Const P_REPORT                      As String = "RKS233"

'==============================================================================*
'
'       MODULE_NAME     : 委託料加瀬負担一覧出力
'       MODULE_ID       : PrintUserMoveList
'       CREATE_DATE     : 2013/05/01
'                       :
'       PARAM           : intMode          - 1=印刷プレビュー 2=Excel出力
'                       : str日付From      -
'                       : str日付To        -
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function PrintUserMoveList(intMode As Integer, str日付From As String, str日付To As String) As Boolean

    Dim rsGetData       As Object
    Dim blnError        As Boolean
    Dim adoDbConnection As Object

On Error GoTo ErrorHandler

    blnError = False
    PrintUserMoveList = False

    'DB接続
    Call subConnectServer(adoDbConnection)

    'データ検索
    If Not fncGetData(adoDbConnection, rsGetData, str日付From, str日付To) Then
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
'       CREATE_DATE     :
'                       :
'       PARAM           : aConnection      - データベースオブジェクト
'                       : rsGetData        - 検索結果を格納するレコードセット
'                       : str日付From      -
'                       : str日付To        -
'
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(aConnection As Object, ByRef rsGetData As Object, str日付From As String, str日付To As String) As Boolean

    Dim strSQL      As String
    Dim rsData      As Object

On Error GoTo ErrorHandler

    fncGetData = False

    'メインSQL文作成
    strSQL = fncMakeGetDataSql(str日付From, str日付To)

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
'       CREATE_DATE     :
'                       :
'       PARAM           : str売上年月      - 売上年月
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(str日付From As String, str日付To As String) As String

    Dim strSQL              As String
    Dim strBumonCode        As String
    
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    
    strSQL = "          SELECT CARG_UCODE, "
    strSQL = strSQL & "        USER_NAME, "
    strSQL = strSQL & "        CARG_YCODE, "
    strSQL = strSQL & "        YARD_NAME, "
    strSQL = strSQL & "        CARG_NO, "
    strSQL = strSQL & "        CARG_STDATE, "
    strSQL = strSQL & "        CARG_HOSYICD, "
    strSQL = strSQL & "        CARG_HOSYO_CD, "
    strSQL = strSQL & "        SHIRM_SHIRN, "
    strSQL = strSQL & "        CARG_RENTKG + CARG_SYOZEI as CARG_SUM , "
    strSQL = strSQL & "        RCPT_SECUKG_KASE, "
    strSQL = strSQL & "        BUMOM_BUMOC, "
    strSQL = strSQL & "        BUMOM_BUMON, "
    strSQL = strSQL & "        RCPT_SEIKYU_CD, "
    strSQL = strSQL & "        NAME_NAME "     '' 2013/07/06 M.HONDA INS
    strSQL = strSQL & "   FROM CARG_FILE "
    strSQL = strSQL & "        INNER JOIN RCPT_TRAN ON "
    strSQL = strSQL & "        CARG_ACPTNO = RCPT_CARG_ACPTNO "
    strSQL = strSQL & "        AND ISNULL(RCPT_SECUKG_KASE,0) <> 0 "
    '' 2013/07/06 M.HONDA INS
    strSQL = strSQL & "        LEFT JOIN NAME_MAST ON "
    strSQL = strSQL & "        NAME_ID = '014' AND "
    strSQL = strSQL & "        NAME_CODE = RCPT_SEIKYU_CD "
    '' 2013/07/06 M.HONDA INS
    strSQL = strSQL & "        INNER JOIN USER_MAST ON "
    strSQL = strSQL & "        CARG_UCODE = USER_CODE "
    strSQL = strSQL & "        INNER JOIN YARD_MAST ON"
    strSQL = strSQL & "        CARG_YCODE = YARD_CODE"
    strSQL = strSQL & "        INNER JOIN KASE_DB.dbo.BUMO_MAST ON"
    strSQL = strSQL & "        BUMOM_BUMOC = '" & strBumonCode & "'"
    strSQL = strSQL & "        INNER JOIN KASE_DB.dbo.SHIR_MAST ON "
    strSQL = strSQL & "        SHIRM_BUMOC = '" & strBumonCode & "' AND  "
    strSQL = strSQL & "        SHIRM_SHIRC = CARG_HOSYO_CD "
    strSQL = strSQL & "  WHERE  CARG_STDATE BETWEEN '" & str日付From & "' AND '" & str日付To & "' "
    strSQL = strSQL & "   ORDER BY CARG_UCODE "
    
    fncMakeGetDataSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :subMakeWork
'        機能             :ワークテーブルデータ追加
'        CREATE_DATE      :
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

On Error GoTo ErrorHandler

    blnError = False

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
            rsDestination.Fields("顧客コード") = .Fields("CARG_UCODE")
            rsDestination.Fields("顧客名称") = .Fields("USER_NAME")
            rsDestination.Fields("ヤードコード") = .Fields("CARG_YCODE")
            rsDestination.Fields("ヤード名称") = .Fields("YARD_NAME")
            rsDestination.Fields("部屋番号") = .Fields("CARG_NO")
            rsDestination.Fields("賃料発生日") = .Fields("CARG_STDATE")
            rsDestination.Fields("保証会社コード") = .Fields("CARG_HOSYO_CD")
            rsDestination.Fields("保証会社名称") = .Fields("SHIRM_SHIRN")
            rsDestination.Fields("初回支払方法") = .Fields("NAME_NAME")                 '' 2013/07/06 M.HONDA INS
            rsDestination.Fields("月額使用料") = .Fields("CARG_SUM")
            rsDestination.Fields("加瀬負担委託料") = .Fields("RCPT_SECUKG_KASE")
            rsDestination.Fields("部門コード") = .Fields("BUMOM_BUMOC")
            rsDestination.Fields("部門名称") = .Fields("BUMOM_BUMON")
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
 
        Call .Fields.Append(.CreateField("顧客コード", DataTypeEnum.dbText, 6))
        Call .Fields.Append(.CreateField("顧客名称", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))
        Call .Fields.Append(.CreateField("ヤード名称", DataTypeEnum.dbText, 60))
        Call .Fields.Append(.CreateField("部屋番号", DataTypeEnum.dbText, 6))
        Call .Fields.Append(.CreateField("賃料発生日", DataTypeEnum.dbText, 10))
        Call .Fields.Append(.CreateField("保証会社コード", DataTypeEnum.dbText, 6))
        Call .Fields.Append(.CreateField("保証会社名称", DataTypeEnum.dbText, 50))
        Call .Fields.Append(.CreateField("初回支払方法", DataTypeEnum.dbText, 50))    '' 2013/07/06 M.HONDA INS
        Call .Fields.Append(.CreateField("月額使用料", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("加瀬負担委託料", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("部門コード", DataTypeEnum.dbText, 50))
        Call .Fields.Append(.CreateField("部門名称", DataTypeEnum.dbText, 50))
        
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
