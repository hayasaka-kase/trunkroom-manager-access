Attribute VB_Name = "Print232"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：  補償加入費一覧
'   プログラムＩＤ　：　Print232
'   作　成　日　　　：  2013/02/21
'   作　成　者　　　：  M.HONDA
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :   2014/12/03
'   UPDATER         :   M.Honda
'   Ver             :   0.1
'   変更内容        :   保証加入費の金額を表示
'
'   UPDATE          :   2017/08/01
'   UPDATER         :   M.Honda
'   Ver             :   0.2
'   変更内容        :   エクセル出力時のコマンドを変更
'
'   UPDATE          :   2018/02/01
'   UPDATER         :   M.Honda
'   Ver             :   0.3
'   変更内容        :   契約日を追加
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'処理モード
Public Const P232_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P232_MODE_EXCEL                As Integer = 2  'Excelに出力

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RKS232_W01"

'レポート名
Private Const P_REPORT                      As String = "RKS232"

'==============================================================================*
'
'       MODULE_NAME     : 補償加入費一覧出力
'       MODULE_ID       : PrintUserMoveList
'       CREATE_DATE     : 2010/02/03
'                       :
'       PARAM           : intMode          - 1=印刷プレビュー 2=Excel出力
'                       : str売上年月      - 売上年月
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function PrintUserMoveList(intMode As Integer, str売上年月 As String) As Boolean

    Dim rsGetData       As Object
    Dim blnError        As Boolean
    Dim adoDbConnection As Object

On Error GoTo ErrorHandler

    blnError = False
    PrintUserMoveList = False

    'DB接続
    Call subConnectServer(adoDbConnection)

    'データ検索
    If Not fncGetData(adoDbConnection, rsGetData, str売上年月) Then
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
            '2017/08/01 M.HONDA UPD
            'doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, , True
            doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLSX, , True
            '2017/08/01 M.HONDA UPD
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
'                       : str売上年月      - 売上年月
'
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(aConnection As Object, ByRef rsGetData As Object, str売上年月 As String) As Boolean

    Dim strSQL      As String
    Dim rsData      As Object

On Error GoTo ErrorHandler

    fncGetData = False

    'メインSQL文作成
    strSQL = fncMakeGetDataSql(str売上年月)

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
Private Function fncMakeGetDataSql(str売上年月 As String) As String

    Dim strSQL              As String
    Dim strBumonCode        As String
    
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    
    strSQL = " select URIAT_SYO1C, "
    strSQL = strSQL & "        TINHM_TINTN, "
    strSQL = strSQL & "        URIAT_NUMBC, "
    strSQL = strSQL & "        URIAT_KOKYC, "
    strSQL = strSQL & "        KOKYM_KOKYN, "
    strSQL = strSQL & "        KOKYM_YUBIB, "
    strSQL = strSQL & "        KOKYM_ADR1N, "
    strSQL = strSQL & "        KOKYM_ADR2N, "
    strSQL = strSQL & "        KOKYM_ADR3N, "
    strSQL = strSQL & "        KEIYT_KAISD, "
    strSQL = strSQL & "        KEIYT_MOUSD, " '2018/02/01 M.HONDA
    strSQL = strSQL & "        JUKUT_KINGA, " '2014/12/03 M.HODNA INS
    strSQL = strSQL & "        TINDM_SENSQ  " '2019/11/07 M.HODNA INS
    strSQL = strSQL & " from URIA_TRAN "
    strSQL = strSQL & "    INNER JOIN JUKU_TRAN "
    strSQL = strSQL & "        ON     URIAT_BUMOC = JUKUT_BUMOC "
    strSQL = strSQL & "        and    URIAT_KEIYB = JUKUT_KEIYB "
    strSQL = strSQL & "        and    URIAT_URIAI = JUKUT_SEIKI "
    strSQL = strSQL & "        and    JUKUT_ZAPII = '68' "
    strSQL = strSQL & "    INNER JOIN KOKY_MAST "
    strSQL = strSQL & "        ON    URIAT_BUMOC = KOKYM_BUMOC "
    strSQL = strSQL & "        and   URIAT_KOKYC = KOKYM_KOKYC "
    strSQL = strSQL & "    INNER JOIN TINH_MAST "
    strSQL = strSQL & "        ON     URIAT_BUMOC = TINHM_BUMOC "
    strSQL = strSQL & "        and    URIAT_SYO1C = TINHM_TINTC "
    
    strSQL = strSQL & "        INNER Join "
    strSQL = strSQL & "        TIND_MAST ON "
    strSQL = strSQL & "        TINDM_BUMOC = TINHM_BUMOC AND "
    strSQL = strSQL & "        TINDM_TINTC = TINHM_TINTC AND "
    strSQL = strSQL & "        TINDM_NUMBC = URIAT_NUMBC "
    
    strSQL = strSQL & "    INNER JOIN KEIY_TRAN "
    strSQL = strSQL & "        ON     URIAT_KEIYB = KEIYT_KEIYB "
    strSQL = strSQL & " where  URIAT_URYMD = '" & str売上年月 & "'"
    strSQL = strSQL & " and    URIAT_BUMOC = '" & strBumonCode & "'"
    strSQL = strSQL & " order by URIAT_SYO1C, URIAT_NUMBC "

    fncMakeGetDataSql = strSQL

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
            rsDestination.Fields("ヤードコード") = .Fields("URIAT_SYO1C")
            rsDestination.Fields("ヤード名称") = .Fields("TINHM_TINTN")
            rsDestination.Fields("部屋番号") = .Fields("URIAT_NUMBC")
            rsDestination.Fields("顧客コード") = .Fields("URIAT_KOKYC")
            rsDestination.Fields("顧客名") = .Fields("KOKYM_KOKYN")
            rsDestination.Fields("郵便番号") = .Fields("KOKYM_YUBIB")
            rsDestination.Fields("住所1") = .Fields("KOKYM_ADR1N")
            rsDestination.Fields("住所2") = .Fields("KOKYM_ADR2N")
            rsDestination.Fields("住所3") = .Fields("KOKYM_ADR3N")
            rsDestination.Fields("申込日") = .Fields("KEIYT_MOUSD")
            rsDestination.Fields("賃料発生日") = .Fields("KEIYT_KAISD")   '2018/02/01 M.HODNA INS
            rsDestination.Fields("保証加入費") = .Fields("JUKUT_KINGA")   '2014/12/03 M.HODNA INS
            rsDestination.Fields("サイズ") = .Fields("TINDM_SENSQ")   '2014/12/03 M.HODNA INS
            
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
 
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))             'ヤードコード
        Call .Fields.Append(.CreateField("ヤード名称", DataTypeEnum.dbText, 60))              'ヤード名称
        Call .Fields.Append(.CreateField("部屋番号", DataTypeEnum.dbText, 6))                 '部屋番号
        Call .Fields.Append(.CreateField("顧客コード", DataTypeEnum.dbText, 6))               '顧客コード
        Call .Fields.Append(.CreateField("顧客名", DataTypeEnum.dbText, 50))                  '顧客名
        Call .Fields.Append(.CreateField("郵便番号", DataTypeEnum.dbText, 10))                '郵便番号
        Call .Fields.Append(.CreateField("住所1", DataTypeEnum.dbText, 36))                   '住所1
        Call .Fields.Append(.CreateField("住所2", DataTypeEnum.dbText, 36))                   '住所2
        Call .Fields.Append(.CreateField("住所3", DataTypeEnum.dbText, 36))                   '住所3
        Call .Fields.Append(.CreateField("申込日", DataTypeEnum.dbText, 8))               '契約日
        Call .Fields.Append(.CreateField("賃料発生日", DataTypeEnum.dbText, 8))               '賃料発生日
        Call .Fields.Append(.CreateField("保証加入費", DataTypeEnum.dbLong))               '保証加入費
        Call .Fields.Append(.CreateField("サイズ", DataTypeEnum.dbText, 8))              '保証加入費
        
        
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

On Error GoTo ErrorHandler

    'ADO接続Object生成
    'KASEDBへ接続
    Set adoDbConnection = MSZZ025.ADODB_Connection()

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
