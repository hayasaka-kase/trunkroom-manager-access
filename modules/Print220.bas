Attribute VB_Name = "Print220"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　紹介制限ヤード一覧出力
'   プログラムＩＤ　：　Print220
'   作　成　日　　　：  2007/05/01
'   作　成　者　　　：  イーグルソフト 鈴木
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :
'   UPDATER            :
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
Public Const P220_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P220_MODE_EXCEL                As Integer = 2  'Excelに出力
Public Const P220_MODE_PRINT                As Integer = 3  'プレビューを表示しないで印刷

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RVS220_W01"

'レポート名
Private Const P_REPORT                      As String = "RVS220"

Private pstrBumonCd                         As String        ' 部門コード
Private pstrBumonNm                         As String        ' 部門名

'==============================================================================*
' デバック用
'==============================================================================*
Sub a00Test_fncPrintYardListInlimit()

'    If Not fncPrintYardListInlimit(P220_MODE_PREVIEW, "12", "12") Then
    If Not fncPrintYardListInlimit(P220_MODE_PREVIEW, "12", "") Then
        MsgBox "False"
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 紹介制限ヤード一覧出力
'       MODULE_ID       : fncPrintYardListInlimit
'       CREATE_DATE     : 2007/05/05
'                       :
'       PARAM           : intMode       - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'                       : strYardCodeF  - ヤードコードFrom
'                         strYardCodeT  - ヤードコードTo
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncPrintYardListInlimit(intMode As Integer, _
                                        strYardCodeF As String, _
                                        strYardCodeT As String) As Boolean
                                        
On Error GoTo ErrorHandler

    Dim dbSQLServer     As Database
    Dim rsGetData       As Recordset
    Dim blnError        As Boolean

    blnError = False

    fncPrintYardListInlimit = False

    'DB接続
    Call subConnectServer(dbSQLServer)

    '部門コード／部門名取得
    Call subGetBumonName

    'データ検索
    If Not fncGetData(dbSQLServer, rsGetData, strYardCodeF, strYardCodeT) Then
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
    Case P220_MODE_PREVIEW:
        'レポートプレビュー
        doCmd.OpenReport P_REPORT, acViewPreview
    Case P220_MODE_EXCEL:
        'EXCELファイル出力
        On Error Resume Next
        doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, , True
        On Error GoTo ErrorHandler
    Case P220_MODE_PRINT:
        'レポート印刷
        On Error Resume Next
        doCmd.OpenReport P_REPORT
        On Error GoTo ErrorHandler
    End Select

    fncPrintYardListInlimit = True

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
'       CREATE_DATE     : 2007/05/05
'                       :
'       PARAM           : dbSqlServer  - KOMSに接続したデータベースオブジェクト
'                       : rsGetData    - 検索結果を格納するレコードセット
'                       : strYardCodeF - ヤードコードFrom
'                       : strYardCodeT - ヤードコードTo
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(dbSQLServer As Database, _
                            ByRef rsGetData As Recordset, _
                            strYardCodeF As String, _
                            strYardCodeT As String) As Boolean

On Error GoTo ErrorHandler
    
    Dim strSQL      As String

    fncGetData = False

    'SQL文作成
    strSQL = fncMakeGetDataSql(strYardCodeF, strYardCodeT)

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
'       CREATE_DATE     : 2007/05/05
'                       :
'                       : strYardCodeF - ヤードコード
'                       : strYardCodeT - ヤードコード
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(strYardCodeF As String, strYardCodeT As String) As String

    Dim strSQL      As String
    Dim strWhere    As String

    strSQL = " SELECT  YARD_MAST.YARD_CODE "                           ' -- 制限されているヤードコード
    strSQL = strSQL & "       ,YARD_MAST.YARD_NAME YARD_NAME "         ' -- 制限されているヤード名
    strSQL = strSQL & "       ,CNTA_CNT "                              ' -- 設置数
    strSQL = strSQL & "       ,CNTA_CNT - KEIYAKU_CNT SPACE "          ' -- 空数
    strSQL = strSQL & "       ,COUNT(IDO_CNT.YARD_CODE) IDO_CNT "      ' -- 解約ヤードからの移動契約数
    strSQL = strSQL & "       ,IDO_TOTAL.IDO_TOTAL "                   ' -- 移動契約総数
    strSQL = strSQL & "       ,YARD_MAST.YARD_INLIMIT_DAY "            ' -- 紹介制限期限
    strSQL = strSQL & "       ,YARD_MAST.YARD_INLIMIT_YCODE "          ' -- 制限をしている解約ヤードコード
    strSQL = strSQL & "       ,KAIYAKU_YARD.YARD_NAME YARD_NAME_K "    ' -- 制限をしている解約ヤード名
    strSQL = strSQL & "       ,CARG_CNT "                              ' -- 残契約数
    strSQL = strSQL & "  FROM YARD_MAST "                              ' -- ヤードマスタ
    ' -------------------------------------------------------------------------------------------------------------
    ' コンテナマスタ(設置数取得)
    strSQL = strSQL & "       INNER JOIN "
    strSQL = strSQL & "       ( "
    strSQL = strSQL & "         SELECT  CNTA_CODE "
    strSQL = strSQL & "                ,COUNT(*) CNTA_CNT "
    strSQL = strSQL & "           FROM CNTA_MAST "
    strSQL = strSQL & "          WHERE CNTA_USE <> '9' "
    strSQL = strSQL & "          GROUP BY CNTA_CODE "
    strSQL = strSQL & "         ) CNTA_MAST "                                               ' -- コンテナマスタ
    strSQL = strSQL & "       ON YARD_MAST.YARD_CODE = CNTA_CODE "
    ' -------------------------------------------------------------------------------------------------------------
    ' -- コンテナ契約ファイル(空数取得のための契約数取得)
    strSQL = strSQL & "       LEFT OUTER JOIN "
    strSQL = strSQL & "       ( "
    strSQL = strSQL & "         SELECT  CARG_YCODE "
    strSQL = strSQL & "                ,COUNT(*) KEIYAKU_CNT "
    strSQL = strSQL & "           FROM CARG_FILE "
    strSQL = strSQL & "          WHERE CARG_AGRE <> '9' "
    strSQL = strSQL & "          GROUP BY CARG_YCODE "
    strSQL = strSQL & "        ) KEIYAKU "
    strSQL = strSQL & "       ON YARD_MAST.YARD_CODE = KEIYAKU.CARG_YCODE "
    ' -------------------------------------------------------------------------------------------------------------
    ' -- 解約ヤードマスタ(解約ヤード名取得)
    strSQL = strSQL & "       INNER JOIN YARD_MAST KAIYAKU_YARD "
    strSQL = strSQL & "       ON YARD_MAST.YARD_INLIMIT_YCODE = KAIYAKU_YARD.YARD_CODE "    ' -- 解約ヤード名称取得用のヤードマスタ
    ' -------------------------------------------------------------------------------------------------------------
    ' -- 解約ヤードからの移動契約数(解約ヤードコードとイコールの数)
    ' -- メインSQLとは制限されているヤードコード、制限しているヤードコードとを外部結合すればOK！！！
    strSQL = strSQL & "       LEFT OUTER JOIN "
    strSQL = strSQL & "       ( "
    strSQL = strSQL & "         SELECT  CARG_FILE.CARG_YCODE   YARD_CODE "                   ' -- 制限されているヤードコード
    strSQL = strSQL & "                ,CARG_FILE2.CARG_YCODE YARD_CODE_K "                  ' -- 制限しているヤードコード
    strSQL = strSQL & "           FROM  CARG_FILE "                                          ' -- コンテナ契約ファイル(制限されているヤード)
    strSQL = strSQL & "                ,YOUK_TRAN "                                          ' -- 予約受付トラン
    strSQL = strSQL & "                ,CARG_FILE CARG_FILE2 "                               ' -- コンテナ契約ファイル(制限しているヤード)
    strSQL = strSQL & "          WHERE YOUKT_MOTO_ACPTNO = CARG_FILE2.CARG_ACPTNO "
    strSQL = strSQL & "            AND CARG_FILE.CARG_UKNO = YOUKT_UKNO "
    strSQL = strSQL & "        ) IDO_CNT "
    strSQL = strSQL & "       ON  YARD_MAST.YARD_CODE = IDO_CNT.YARD_CODE "
    strSQL = strSQL & "       AND YARD_MAST.YARD_INLIMIT_YCODE = IDO_CNT.YARD_CODE_K "
    ' -------------------------------------------------------------------------------------------------------------
    ' -- コンテナ契約ファイル(移動契約総数)
    ' -- メインSQLとは制限をしている解約ヤードコードと外部結合をすればOK！！！
    strSQL = strSQL & "       LEFT OUTER JOIN "
    strSQL = strSQL & "       ( "
    strSQL = strSQL & "         SELECT  CARG_FILE2.CARG_YCODE "
    strSQL = strSQL & "                ,COUNT(*) IDO_TOTAL "
    strSQL = strSQL & "           FROM  CARG_FILE "                                          ' -- コンテナ契約ファイル(制限しているヤード)
    strSQL = strSQL & "                ,YOUK_TRAN "                                          ' -- 予約受付トラン
    strSQL = strSQL & "                ,CARG_FILE CARG_FILE2 "                               ' -- コンテナ契約ファイル(制限されているヤード)
    strSQL = strSQL & "          WHERE YOUKT_MOTO_ACPTNO = CARG_FILE2.CARG_ACPTNO"
    strSQL = strSQL & "            AND CARG_FILE.CARG_UKNO = YOUKT_UKNO "
    strSQL = strSQL & "          GROUP BY CARG_FILE2.CARG_YCODE "
    strSQL = strSQL & "        ) IDO_TOTAL "
    strSQL = strSQL & "       ON YARD_MAST.YARD_INLIMIT_YCODE = IDO_TOTAL.CARG_YCODE "
    ' -------------------------------------------------------------------------------------------------------------
    ' -- コンテナ契約ファイル(残契約数)
    strSQL = strSQL & "       LEFT OUTER JOIN "
    strSQL = strSQL & "       ( "
    strSQL = strSQL & "         SELECT  CARG_YCODE "
    strSQL = strSQL & "                ,COUNT(*) CARG_CNT "
    strSQL = strSQL & "           FROM CARG_FILE "
    strSQL = strSQL & "          WHERE CARG_AGRE <> '9' "
    strSQL = strSQL & "          GROUP BY CARG_YCODE "
    strSQL = strSQL & "        ) CARG_FILE "
    strSQL = strSQL & "       ON YARD_MAST.YARD_INLIMIT_YCODE = CARG_FILE.CARG_YCODE "
'TODO たぶん× このコードだと移動先ヤードの契約数を取得   strSQL = strSQL & "       ON YARD_MAST.YARD_CODE = CARG_FILE.CARG_YCODE "
    ' -------------------------------------------------------------------------------------------------------------
    strSQL = strSQL & " WHERE CONVERT(DATETIME, ISNULL(YARD_MAST.YARD_INLIMIT_DAY,'1900/01/01')) >= GETDATE() "

    ' ヤード範囲のWhere句生成 TODO:ちょっと保留ね
    'strWhere = fncMakeSqlWhere(strYardCodeF, strYardCodeT)

    If Nz(strWhere, "") <> "" Then
        strSQL = strSQL & strWhere
    End If

    strSQL = strSQL & " GROUP BY  YARD_MAST.YARD_CODE "                                     ' -- 制限されているヤードコード
    strSQL = strSQL & "          ,YARD_MAST.YARD_NAME "                                     ' -- 制限されているヤード名
    strSQL = strSQL & "          ,CNTA_CNT "                                                ' -- 設置数
    strSQL = strSQL & "          ,CNTA_CNT - KEIYAKU_CNT "                                  ' -- 空数
    strSQL = strSQL & "          ,IDO_TOTAL.IDO_TOTAL "                                     ' -- 移動契約総数
    strSQL = strSQL & "          ,YARD_MAST.YARD_INLIMIT_DAY "                              ' -- 紹介制限期限
    strSQL = strSQL & "          ,YARD_MAST.YARD_INLIMIT_YCODE "                            ' -- 制限をしている解約ヤードコード
    strSQL = strSQL & "          ,KAIYAKU_YARD.YARD_NAME "                                  ' -- 制限をしている解約ヤード名
    strSQL = strSQL & "          ,CARG_CNT "                                                ' -- 残契約数
    strSQL = strSQL & " ORDER BY  YARD_MAST.YARD_INLIMIT_YCODE "                            ' -- 制限をしている解約ヤードコード
    strSQL = strSQL & "          ,YARD_MAST.YARD_CODE "                                     ' -- ヤードコード

    fncMakeGetDataSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      : fncMakeSqlWhere
'        機能             : Where文作成
'        OUT              : where文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeSqlWhere(strYardCodeF As String, strYardCodeT As String) As String

    Dim strWhere As String

    ' ヤードコード
    If (Nz(strYardCodeF, "") <> "" And Nz(strYardCodeT, "") <> "") Then
        ' ヤードコード(From)：入力
        ' ヤードコード(To)  ：入力
        strWhere = " AND YARD_MAST.YARD_CODE BETWEEN '" & strYardCodeF & "' AND '" & strYardCodeT & "' "

    ElseIf (Nz(strYardCodeF, "") <> "" And Nz(strYardCodeT, "") = "") Then
        ' ヤードコード(From)：入力
        ' ヤードコード(To)  ：未入力
        strWhere = " AND YARD_MAST.YARD_CODE >= '" & strYardCodeF & "' "

    ElseIf (Nz(strYardCodeF, "") = "" And Nz(strYardCodeT, "") <> "") Then
        ' ヤードコード(From)：未入力
        ' ヤードコード(To)  ：入力
        strWhere = " AND YARD_MAST.YARD_CODE <= '" & strYardCodeT & "' "
    End If

    fncMakeSqlWhere = strWhere

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
        Call .Fields.Append(.CreateField("部門コード", DataTypeEnum.dbText, 36))                '部門コード
        Call .Fields.Append(.CreateField("部門名", DataTypeEnum.dbText, 36))                    '部門名
        Call .Fields.Append(.CreateField("区分", DataTypeEnum.dbText, 14))                      '区分
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))                'ヤードコード
        Call .Fields.Append(.CreateField("ヤード名", DataTypeEnum.dbText, 36))                   'ヤード名
        Call .Fields.Append(.CreateField("設置数", DataTypeEnum.dbInteger))                     '設置数
        Call .Fields.Append(.CreateField("空数", DataTypeEnum.dbInteger))                       '空数
        Call .Fields.Append(.CreateField("解約ヤードからの移動契約数", DataTypeEnum.dbInteger))   '解約ヤードからの移動契約数
        Call .Fields.Append(.CreateField("移動契約総数", DataTypeEnum.dbInteger))                '移動契約総数
        Call .Fields.Append(.CreateField("紹介制限期限", DataTypeEnum.dbText, 20))               '紹介制限期限
        Call .Fields.Append(.CreateField("解約ヤードコード", DataTypeEnum.dbText, 6))            '解約ヤードコード
        Call .Fields.Append(.CreateField("解約ヤード名", DataTypeEnum.dbText, 36))               '解約ヤード名
        Call .Fields.Append(.CreateField("残契約数", DataTypeEnum.dbInteger))                    '残契約数

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
        Call subAddNew(rsSource, rsDestination)
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
Private Sub subAddNew(rsSource As Recordset, rsDestination As Recordset)

    Dim strTemp     As String

    With rsSource
        rsDestination.AddNew

        rsDestination.Fields("部門コード") = pstrBumonCd                                           '部門コード
        rsDestination.Fields("部門名") = pstrBumonNm                                               '部門名
        rsDestination.Fields("区分") = "2"                                                         '区分
        rsDestination.Fields("ヤードコード") = Format(.Fields("YARD_CODE"), "000000")              '制限されているヤードコード
        rsDestination.Fields("ヤード名") = .Fields("YARD_NAME")                                    '制限されているヤード名
        rsDestination.Fields("設置数") = Nz(.Fields("CNTA_CNT"), 0)                                '設置数
        rsDestination.Fields("空数") = Nz(.Fields("SPACE"), 0)                                     '空数
        rsDestination.Fields("解約ヤードからの移動契約数") = Nz(.Fields("IDO_CNT"), 0)               '解約ヤードからの移動契約数
        rsDestination.Fields("移動契約総数") = Nz(.Fields("IDO_TOTAL"), 0)                          '移動契約総数
        rsDestination.Fields("紹介制限期限") = .Fields("YARD_INLIMIT_DAY") & "まで"                 '紹介制限期限
        rsDestination.Fields("解約ヤードコード") = Format(.Fields("YARD_INLIMIT_YCODE"), "000000")  '解約ヤードコード
        rsDestination.Fields("解約ヤード名") = .Fields("YARD_NAME_K")                               '解約ヤード名
        rsDestination.Fields("残契約数") = Nz(.Fields("CARG_CNT"), 0)                               '残契約数

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
'       CREATE_DATE     : 2007/05/05
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
'       CREATE_DATE     : 2007/05/05
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
'       CREATE_DATE     : 2007/05/05
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
'       CREATE_DATE     : 2007/05/05
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
'        MODULE_NAME      : 部門コード/部門名称の取得
'        MODULE_ID        : subGetBumonName
'        CREATE_DATE      : 2007/05/05
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetBumonName()

On Error GoTo ErrorHandler

    Dim strSQL  As String
    Dim objRs   As Recordset
    Dim objDb   As Database

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
