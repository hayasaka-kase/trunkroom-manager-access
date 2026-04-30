Attribute VB_Name = "Print234"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：  配送サービス一覧
'   プログラムＩＤ　：　Print234
'   作　成　日　　　：  2015/01/08
'   作　成　者　　　：  M.HONDA
'   Ver             ：  0.0
'   備考            ：
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'処理モード
Public Const P234_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P234_MODE_EXCEL                As Integer = 2  'Excelに出力

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RKS234_W01"

'レポート名
Private Const P_REPORT                      As String = "RKS234"

'==============================================================================*
'
'       MODULE_NAME     : 配送サービス一覧
'       MODULE_ID       : PrintDlveList
'       CREATE_DATE     :
'                       :
'       PARAM           : intMode          - 1=印刷プレビュー 2=Excel出力
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function PrintDlveList(intMode As Integer, str日付From As String, str日付To As String, strCODEFROM As String, strCODETO As String, str顧客FROM As String, str顧客TO As String) As Boolean

    Dim rsGetData       As Object
    Dim blnError        As Boolean
    Dim adoDbConnection As Object

On Error GoTo ErrorHandler

    blnError = False
    PrintDlveList = False

    'DB接続
    Call subConnectServer(adoDbConnection)

    'データ検索
    If Not fncGetData(adoDbConnection, rsGetData, str日付From, str日付To, strCODEFROM, strCODETO, str顧客FROM, str顧客TO) Then
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

        Case P234_MODE_EXCEL:
            'EXCELファイル出力
            On Error Resume Next
            doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, , True
            On Error GoTo ErrorHandler
    End Select

    PrintDlveList = True

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
Private Function fncGetData(aConnection As Object, ByRef rsGetData As Object, str日付From As String, str日付To As String, strCODEFROM As String, strCODETO As String, str顧客FROM As String, str顧客TO As String) As Boolean

    Dim strSQL      As String
    Dim rsData      As Object

On Error GoTo ErrorHandler

    fncGetData = False

    'メインSQL文作成
    strSQL = fncMakeGetDataSql(str日付From, str日付To, strCODEFROM, strCODETO, str顧客FROM, str顧客TO)

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
Private Function fncMakeGetDataSql(str日付From As String, str日付To As String, strCODEFROM As String, strCODETO As String, str顧客FROM As String, str顧客TO As String) As String

    Dim strSQL              As String
    
    strSQL = "          SELECT RCPT_NO, "
    strSQL = strSQL & "        RCPT_UCODE, "
    strSQL = strSQL & "        RCPT_UKDATE, "
    strSQL = strSQL & "        USER_NAME, "
    strSQL = strSQL & "        USER_YUBINO, "
    strSQL = strSQL & "        USER_ADR_1, "
    strSQL = strSQL & "        USER_ADR_2, "
    strSQL = strSQL & "        USER_ADR_3, "
    strSQL = strSQL & "        USER_TEL, "
    strSQL = strSQL & "        USER_KEITAI, "
    strSQL = strSQL & "        DLVRT_OUT_ADR, "
    strSQL = strSQL & "        DLVRT_IN_YARD_CD, "
    strSQL = strSQL & "        YARD_YUBINO, "
    strSQL = strSQL & "        YARD_ADDR_1, "
    strSQL = strSQL & "        YARD_ADDR_2, "
    strSQL = strSQL & "        RCPT_CNO, "
    strSQL = strSQL & "        DLVRT_KIBO_DATE1, "
    strSQL = strSQL & "        TIMECD1.NAME_NAME AS TIMECD1, "
    strSQL = strSQL & "        DLVRT_KIBO_DATE2, "
    strSQL = strSQL & "        TIMECD2.NAME_NAME AS TIMECD2, "
    strSQL = strSQL & "        DLVRT_KIBO_DATE3, "
    strSQL = strSQL & "        TIMECD3.NAME_NAME AS TIMECD3 "
    strSQL = strSQL & "FROM RCPT_TRAN "
    strSQL = strSQL & "   INNER JOIN DLVR_TRAN ON "
    strSQL = strSQL & "        RCPT_NO = DLVRT_NO "
    strSQL = strSQL & "   INNER JOIN YARD_MAST ON "
    strSQL = strSQL & "        RCPT_YCODE = YARD_CODE "
    strSQL = strSQL & fncMakeBetween("YARD_CODE", strCODEFROM, strCODETO)
    strSQL = strSQL & "   INNER JOIN USER_MAST ON "
    strSQL = strSQL & "        RCPT_UCODE = USER_CODE "
    strSQL = strSQL & fncMakeBetween("USER_CODE", str顧客FROM, str顧客TO)
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST TIMECD1 "
    strSQL = strSQL & "     ON TIMECD1.NAME_ID = '016' "
    strSQL = strSQL & "    AND TIMECD1.NAME_CODE = DLVRT_KIBO_TIMECD1 "
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST TIMECD2 "
    strSQL = strSQL & "     ON TIMECD2.NAME_ID = '016' "
    strSQL = strSQL & "    AND TIMECD2.NAME_CODE = DLVRT_KIBO_TIMECD2 "
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST TIMECD3 "
    strSQL = strSQL & "     ON TIMECD3.NAME_ID = '016' "
    strSQL = strSQL & "    AND TIMECD3.NAME_CODE = DLVRT_KIBO_TIMECD3 "
    strSQL = strSQL & "WHERE 1 = 1 "
    strSQL = strSQL & fncMakeBetween("RCPT_UKDATE", str日付From, str日付To)
    
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
            rsDestination.Fields("受付番号") = .Fields("RCPT_NO")
            rsDestination.Fields("顧客番号") = .Fields("RCPT_UCODE")
            rsDestination.Fields("受付日") = .Fields("RCPT_UKDATE")
            rsDestination.Fields("顧客名") = .Fields("USER_NAME")
            rsDestination.Fields("郵便番号") = .Fields("USER_YUBINO")
            rsDestination.Fields("住所1") = .Fields("USER_ADR_1")
            rsDestination.Fields("住所2") = .Fields("USER_ADR_2")
            rsDestination.Fields("住所3") = .Fields("USER_ADR_3")
            rsDestination.Fields("電話番号") = .Fields("USER_TEL")
            rsDestination.Fields("携帯番号") = .Fields("USER_KEITAI")
            rsDestination.Fields("搬出地住所") = .Fields("DLVRT_OUT_ADR")
            rsDestination.Fields("搬出入地ヤード") = .Fields("DLVRT_IN_YARD_CD")
            rsDestination.Fields("搬出入地1") = .Fields("YARD_ADDR_1")
            rsDestination.Fields("搬出入地2") = .Fields("YARD_ADDR_2")
            rsDestination.Fields("搬出入地部屋") = .Fields("RCPT_CNO")
            rsDestination.Fields("配送希望日1") = .Fields("DLVRT_KIBO_DATE1")
            rsDestination.Fields("配送希望時間1") = .Fields("TIMECD1")
            rsDestination.Fields("配送希望日2") = .Fields("DLVRT_KIBO_DATE2")
            rsDestination.Fields("配送希望時間2") = .Fields("TIMECD2")
            rsDestination.Fields("配送希望日3") = .Fields("DLVRT_KIBO_DATE3")
            rsDestination.Fields("配送希望時間3") = .Fields("TIMECD3")
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
'        CREATE_DATE      :
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew
 
    Call .Fields.Append(.CreateField("受付番号", DataTypeEnum.dbText, 11))
    Call .Fields.Append(.CreateField("受付日", DataTypeEnum.dbText, 20))
    Call .Fields.Append(.CreateField("顧客番号", DataTypeEnum.dbText, 6))
    Call .Fields.Append(.CreateField("顧客名", DataTypeEnum.dbText, 36))
    Call .Fields.Append(.CreateField("郵便番号", DataTypeEnum.dbText, 10))
    Call .Fields.Append(.CreateField("住所1", DataTypeEnum.dbText, 36))
    Call .Fields.Append(.CreateField("住所2", DataTypeEnum.dbText, 36))
    Call .Fields.Append(.CreateField("住所3", DataTypeEnum.dbText, 36))
    Call .Fields.Append(.CreateField("電話番号", DataTypeEnum.dbText, 15))
    Call .Fields.Append(.CreateField("携帯番号", DataTypeEnum.dbText, 15))
    Call .Fields.Append(.CreateField("搬出地住所", DataTypeEnum.dbText, 108))
    Call .Fields.Append(.CreateField("搬出入地ヤード", DataTypeEnum.dbText, 6))
    'Call .Fields.Append(.CreateField("搬出入地郵便番号", DataTypeEnum.dbText, 10))
    Call .Fields.Append(.CreateField("搬出入地1", DataTypeEnum.dbText, 36))
    Call .Fields.Append(.CreateField("搬出入地2", DataTypeEnum.dbText, 36))
    Call .Fields.Append(.CreateField("搬出入地部屋", DataTypeEnum.dbText, 6))
    Call .Fields.Append(.CreateField("配送希望日1", DataTypeEnum.dbText, 8))
    Call .Fields.Append(.CreateField("配送希望時間1", DataTypeEnum.dbText, 20))
    Call .Fields.Append(.CreateField("配送希望日2", DataTypeEnum.dbText, 8))
    Call .Fields.Append(.CreateField("配送希望時間2", DataTypeEnum.dbText, 20))
    Call .Fields.Append(.CreateField("配送希望日3", DataTypeEnum.dbText, 8))
    Call .Fields.Append(.CreateField("配送希望時間3", DataTypeEnum.dbText, 20))
    
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


