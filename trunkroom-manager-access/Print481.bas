Attribute VB_Name = "Print481"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　契約リスト出力
'   プログラムＩＤ　：　Print481
'   作　成　日　　　：  2007/06/22
'   作　成　者　　　：  イーグルソフト 鈴木
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :   2009/12/08
'   UPDATER         :   M.RYU
'   Ver             :   0.1
'   変更内容        :   ①RKS481_W01　⇒　ワークテーブルを作成するとき、保証会社コードや名称を追加
'                       ②RKS481　⇒　レポート出力のとき、保証会社名欄を追加
'
'   UPDATE          :   2011/03/02
'   UPDATER         :   M.RYU
'   Ver             :   0.2
'   変更内容        :　 検索条件に保証会社ｺｰﾄﾞを追加
'
'   UPDATE          :   2025/12/16
'   UPDATER         :   M.HONDA
'   Ver             :   0.3
'   変更内容        :　 担当者カナをエクセル表示に追加
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'処理モード
Public Const P481_MODE_PREVIEW              As Integer = 1  '印刷プレビューを表示
Public Const P481_MODE_EXCEL                As Integer = 2  'Excelに出力
Public Const P481_MODE_PRINT                As Integer = 3  'プレビューを表示しないで印刷

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RKS481_W01"

'レポート名
Private Const P_REPORT                      As String = "RKS481"

'***************************************
' テストプロ
'***************************************
Sub a00Test_fncPrintMaintenanceRequest()

    If Not PrintUserContractList(P481_MODE_PREVIEW, _
                                 "", "", _
                                 "", "", _
                                 "", "", _
                                 "", "", _
                                 "", "", _
                                 "", "", "") Then
        MsgBox "False"
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 契約リスト出力
'       MODULE_ID       : PrintUserContractList
'       CREATE_DATE     : 2007/06/22
'                       :
'       PARAM           : intMode        - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'                       : strKeiyakuDayF - 初回契約日From
'                       : strKeiyakuDayT - 初回契約日To
'                       : strYardCodeF   - ヤードコードFrom
'                       : strYardCodeT   - ヤードコードTo
'                       : strUserCodeF   - 顧客コードFrom
'                       : strUserCodeF   - 顧客コードTo
'                       : strChangeKbnF  - 変更内容From
'                       : strChangeKbnT  - 変更内容To
'                       : strChangeDayF  - 変更日付From
'                       : strChangeDayT  - 変更日付To
'                       : strHosyoCodeF  - 保証区分From
'                       : strHosyoCodeT  - 保証区分To
'                       :
'       NOTE            :
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function PrintUserContractList(intMode As Integer, _
                                      strKeiyakuDayF As String, _
                                      strKeiyakuDayT As String, _
                                      strYardCodeF As String, _
                                      strYardCodeT As String, _
                                      strUserCodeF As String, _
                                      strUserCodeT As String, _
                                      strChangeKbnF As String, _
                                      strChangeKbnT As String, _
                                      strChangeDayF As String, _
                                      strChangeDayT As String, _
                                      strHosyoCodeF As String, _
                                      strHosyoCodeT As String, _
                                      strHosyoCampc As String) As Boolean      'UPDATE 2011/03/02 M.RYU '保証会社ｺｰﾄﾞを追加

    Dim dbSQLServer     As Database
    Dim rsGetData       As Recordset
    Dim blnError        As Boolean

    On Error GoTo ErrorHandler

    blnError = False

    PrintUserContractList = False

    ' DB接続
    Call subConnectServer(dbSQLServer)

    ' データ検索
    If Not fncGetData(dbSQLServer, rsGetData, strKeiyakuDayF, strKeiyakuDayT, strYardCodeF, strYardCodeT, _
                      strUserCodeF, strUserCodeT, strChangeKbnF, strChangeKbnT, _
                      strChangeDayF, strChangeDayT, strHosyoCodeF, strHosyoCodeT, _
                      strHosyoCampc) Then       'UPDATE 2011/03/02 M.RYU '保証会社ｺｰﾄﾞを追加
        ' 該当データ無し
        GoTo ExitRtn
    End If

    ' ワークテーブル作成
    Call subMakeWork(rsGetData, intMode)

    ' DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    ' 出力
    Select Case intMode
        Case P481_MODE_PREVIEW:
            ' レポートプレビュー
            doCmd.OpenReport P_REPORT, acViewPreview

        Case P481_MODE_EXCEL:
            ' EXCELファイル出力
            On Error Resume Next
            doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, , True
            On Error GoTo ErrorHandler

        Case P481_MODE_PRINT:
            ' レポート印刷
            On Error Resume Next
            doCmd.OpenReport P_REPORT
            On Error GoTo ErrorHandler
    End Select

    PrintUserContractList = True

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "PrintUserContractList" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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
'       MODULE_NAME     : データ検索
'       MODULE_ID       : fncGetData
'       CREATE_DATE     : 2007/06/22
'                       :
'       PARAM           : dbSqlServer    - KOMSに接続したデータベースオブジェクト
'                       : rsGetData      - 検索結果を格納するレコードセット
'                       : strKeiyakuDayF - 初回契約日From
'                       : strKeiyakuDayT - 初回契約日To
'                       : strYardCodeF   - ヤードコードFrom
'                       : strYardCodeT   - ヤードコードTo
'                       : strUserCodeF   - 顧客コードFrom
'                       : strUserCodeT   - 顧客コードTo
'                       : strChangeKbnF  - 変更内容From
'                       : strChangeKbnT  - 変更内容To
'                       : strChangeDayF  - 変更日付From
'                       : strChangeDayT  - 変更日付To
'                       : strHosyoCodeF  - 保証区分From
'                       : strHosyoCodeT  - 保証区分To
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(dbSQLServer As Database, _
                            rsGetData As Recordset, _
                            strKeiyakuDayF As String, _
                            strKeiyakuDayT As String, _
                            strYardCodeF As String, _
                            strYardCodeT As String, _
                            strUserCodeF As String, _
                            strUserCodeT As String, _
                            strChangeKbnF As String, _
                            strChangeKbnT As String, _
                            strChangeDayF As String, _
                            strChangeDayT As String, _
                            strHosyoCodeF As String, _
                            strHosyoCodeT As String, _
                            strHosyoCampc As String) As Boolean      'UPDATE 2011/03/02 M.RYU '保証会社ｺｰﾄﾞを追加

    Dim strSQL      As String

    On Error GoTo ErrorHandler

    fncGetData = False

    'SQL文作成
    strSQL = fncMakeGetDataSql(strKeiyakuDayF, strKeiyakuDayT, strYardCodeF, strYardCodeT, strUserCodeF, strUserCodeT, _
                               strChangeKbnF, strChangeKbnT, strChangeDayF, strChangeDayT, strHosyoCodeF, strHosyoCodeT, _
                               strHosyoCampc)        'UPDATE 2011/03/02 M.RYU '保証会社ｺｰﾄﾞを追加

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
'       CREATE_DATE     : 2007/06/22
'                       :
'       PARAM           : strKeiyakuDayF - 初回契約日From
'                       : strKeiyakuDayT - 初回契約日To
'                       : strYardCodeF   - ヤードコードFrom
'                       : strYardCodeT   - ヤードコードTo
'                       : strUserCodeF   - 顧客コードFrom
'                       : strUserCodeT   - 顧客コードTo
'                       : strChangeKbnF  - 変更内容From
'                       : strChangeKbnT  - 変更内容To
'                       : strChangeDayF  - 変更日付From
'                       : strChangeDayT  - 変更日付To
'                       : strHosyoCodeF  - 保証区分From
'                       : strHosyoCodeT  - 保証区分To
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(strKeiyakuDayF As String, _
                                   strKeiyakuDayT As String, _
                                   strYardCodeF As String, _
                                   strYardCodeT As String, _
                                   strUserCodeF As String, _
                                   strUserCodeT As String, _
                                   strChangeKbnF As String, _
                                   strChangeKbnT As String, _
                                   strChangeDayF As String, _
                                   strChangeDayT As String, _
                                   strHosyoCodeF As String, _
                                   strHosyoCodeT As String, _
                                   strHosyoCampc As String) As String      'UPDATE 2011/03/02 M.RYU '保証会社ｺｰﾄﾞを追加

    '----20091110----M.RYU----add--------<s>--'KASE_DB名前を取得
    Dim strKASEDBN As String
    strKASEDBN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME'")
    strKASEDBN = strKASEDBN & ".dbo."
    '----20091110----M.RYU----add--------<e>
    
    Dim strSQL              As String

    strSQL = " SELECT    USER_MAST.USER_CODE     " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_CHG_RKBN " & Chr(13)
    strSQL = strSQL & " ,NAME1.NAME_NAME         CHG_RKBN_NAME " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_NAME " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_KANA " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_TANM " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_TAKA " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_YUBINO " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_KKBN " & Chr(13)
    strSQL = strSQL & " ,NAME2.NAME_NAME         KKBN_NAME" & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_ADR_1 " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_ADR_2 " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_ADR_3 " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_TEL " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_FAX " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_KEITAI " & Chr(13)
    '2016/11/01 M.HONDA INS
    strSQL = strSQL & " ,USER_MAST.USER_RNAME  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RKANA  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RPOST  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RADR_1  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RADR_2  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RADR_3  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RTEL  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RKEITAI  " & Chr(13)
    strSQL = strSQL & " ,USER_MAST.USER_RFAX  " & Chr(13)
    '2016/11/01 M.HONDA INS
    strSQL = strSQL & " ,CARG_FILE.CARG_ACPTNO " & Chr(13)
    strSQL = strSQL & " ,CARG_FILE.CARG_HOSYB " & Chr(13)
    strSQL = strSQL & " ,CARG_FILE.CARG_YCODE " & Chr(13)
    strSQL = strSQL & " ,YARD_MAST.YARD_NAME " & Chr(13)
    strSQL = strSQL & " ,CARG_FILE.CARG_NO " & Chr(13)
    strSQL = strSQL & " ,CARG_FILE.CARG_HOKAI " & Chr(13)
    strSQL = strSQL & " ,NAME3.NAME_RYAK HOKAI_NAME " & Chr(13)
    
    '----20091110----M.RYU----add--------<s>
    strSQL = strSQL & " ,CARG_FILE.CARG_HOSYO_CD " & Chr(13)
    strSQL = strSQL & " ,NAME4.SHIRM_SHIRN       " & Chr(13)
    '----20091110----M.RYU----add--------<e>
    
    strSQL = strSQL & " FROM   USER_MAST " & Chr(13)                                                                 ' ユーザーマスタ
    strSQL = strSQL & "       ,CARG_FILE " & Chr(13)                                                                 ' コンテナ契約ファイル
    strSQL = strSQL & "       ,YARD_MAST " & Chr(13)                                                                 ' ヤードマスタ"
    strSQL = strSQL & "       ,(SELECT NAME_CODE, NAME_NAME FROM NAME_MAST WHERE NAME_ID = '248') NAME1 " & Chr(13)  ' 変更内容
    strSQL = strSQL & "       ,(SELECT NAME_CODE, NAME_NAME FROM NAME_MAST WHERE NAME_ID = '020') NAME2 " & Chr(13)  ' 区分名
    strSQL = strSQL & "       ,(SELECT NAME_CODE, NAME_RYAK FROM NAME_MAST WHERE NAME_ID = '200') NAME3 " & Chr(13)  ' 保証区分名
    
    '----20091110----M.RYU----add--------<s>
    strSQL = strSQL & "     ,(SELECT SHIR_MAST.SHIRM_SHIRC ,                                                                      " & Chr(13)
    strSQL = strSQL & "         (CASE WHEN SHIR_MAST.SHIRM_SHIRN LIKE '%クレデンス%' THEN '株式会社クレデンス'                    " & Chr(13)
    strSQL = strSQL & "         ELSE CASE WHEN SHIR_MAST.SHIRM_SHIRN LIKE '%パルマ%' THEN '株式会社パルマ'                        " & Chr(13)
    strSQL = strSQL & "         ELSE CASE WHEN SHIR_MAST.SHIRM_SHIRN LIKE '%アールエムトラス%' THEN 'アールエムトラスト株式会社'  " & Chr(13)
    strSQL = strSQL & "         ELSE CASE WHEN SHIR_MAST.SHIRM_SHIRN LIKE '%リプラス%' THEN '株式会社リプラス'                    " & Chr(13)
    strSQL = strSQL & "         ELSE  SHIR_MAST.SHIRM_SHIRN END END END END) AS SHIRM_SHIRN                                       " & Chr(13)
    strSQL = strSQL & "     FROM " & strKASEDBN & "SHIR_MAST                                                                      " & Chr(13)
    strSQL = strSQL & "     INNER JOIN CONT_MAST ON CONT_MAST.CONT_BUMOC = SHIR_MAST.SHIRM_BUMOC                                  " & Chr(13)
    strSQL = strSQL & "     WHERE SHIR_MAST.SHIRM_SHIRI IN ('8','9') ) AS NAME4                                                   " & Chr(13)
    '----20091110----M.RYU----add--------<e>
    
    strSQL = strSQL & " WHERE USER_MAST.USER_CHG_RKBN = NAME1.NAME_CODE     " & Chr(13)                                  ' 変更内容
    strSQL = strSQL & "   AND USER_MAST.USER_KKBN     = NAME2.NAME_CODE     " & Chr(13)                                  ' 区分名
    strSQL = strSQL & "   AND CARG_FILE.CARG_HOSYICD  = NAME3.NAME_CODE     " & Chr(13)                                  ' 保証区分名
    
    '----20091110----M.RYU----add--------
    strSQL = strSQL & "   AND CARG_FILE.CARG_HOSYO_CD = NAME4.SHIRM_SHIRC   " & Chr(13)                              ' 保証会社名
    
    strSQL = strSQL & "   AND CARG_FILE.CARG_AGRE     <> 9 " & Chr(13)                                               ' 契約状態
    strSQL = strSQL & "   AND CARG_FILE.CARG_YCODE    = YARD_MAST.YARD_CODE " & Chr(13)                              ' ヤードコード
    strSQL = strSQL & "   AND USER_MAST.USER_CODE     = CARG_FILE.CARG_UCODE " & Chr(13)                             ' ユーザーマスタ VS コンテナ契約ファイル

    ' From～To条件設定
    strSQL = strSQL & fncMakeBetween(strKeiyakuDayF, strKeiyakuDayT, _
                                     strYardCodeF, strYardCodeT, _
                                     strUserCodeF, strUserCodeT, _
                                     strChangeKbnF, strChangeKbnT, _
                                     strChangeDayF, strChangeDayT, _
                                     strHosyoCodeF, strHosyoCodeT)
    If strHosyoCampc <> "" Then
        strSQL = strSQL & "   AND CARG_FILE.CARG_HOSYO_CD     = '" & strHosyoCampc & "' " & Chr(13)     '保証会社   'INSERT 2011/03/02 M.RYU '保証会社ｺｰﾄﾞを追加                                              ' 保証会社
    End If

    ' *** ソート順 ***
    strSQL = strSQL & " ORDER BY  USER_MAST.USER_CODE " & Chr(13)
    strSQL = strSQL & "         ,CARG_FILE.CARG_YCODE " & Chr(13)
    strSQL = strSQL & "         ,CARG_FILE.CARG_NO "

    fncMakeGetDataSql = strSQL

End Function

'==============================================================================*
'
'       MODULE_NAME     :範囲条件作成
'       MODULE_ID       :fncMakeBetween
'       PARAM           : strKeiyakuDayF - 初回契約日From
'                       : strKeiyakuDayT - 初回契約日To
'                       : strYardCodeF   - ヤードコードFrom
'                       : strYardCodeT   - ヤードコードTo
'                       : strUserCodeF   - 顧客コードFrom
'                       : strUserCodeT   - 顧客コードTo
'                       : strChangeKbnF  - 変更内容From
'                       : strChangeKbnT  - 変更内容To
'                       : strChangeDayF  - 変更日付From
'                       : strChangeDayT  - 変更日付To
'                       : strHosyoCodeF  - 保証区分From
'                       : strHosyoCodeT  - 保証区分To
'        OUT            :条件文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeBetween(strKeiyakuDayF As String, _
                                strKeiyakuDayT As String, _
                                strYardCodeF As String, _
                                strYardCodeT As String, _
                                strUserCodeF As String, _
                                strUserCodeT As String, _
                                strChangeKbnF As String, _
                                strChangeKbnT As String, _
                                strChangeDayF As String, _
                                strChangeDayT As String, _
                                strHosyoCodeF As String, _
                                strHosyoCodeT As String) As String

    Dim strTemp     As String

    strTemp = ""

    ' *** 初回契約日 ***
    If strKeiyakuDayF <> "" And strKeiyakuDayT <> "" Then

        ' 共に空白ではない場合
        If strKeiyakuDayF = strKeiyakuDayT Then
            '　FROMとTOが同一の場合、一致条件
            strTemp = " AND CARG_FSDATE = '" & strKeiyakuDayF & "'"
        Else
            ' FROMとTOが異なる場合、BETWEEN条件
            strTemp = " AND CARG_FSDATE BETWEEN '" & strKeiyakuDayF & "' AND '" & strKeiyakuDayT & "' "
        End If
        
    ElseIf strKeiyakuDayF <> "" Then
        ' FROMのみの場合、それ以上であることが条件
        strTemp = " AND CARG_FSDATE >= '" & strKeiyakuDayF & "' "

    ElseIf strKeiyakuDayT <> "" Then
        ' TOのみの場合、それ以下であることが条件
            strTemp = " AND CARG_FSDATE <= '" & strKeiyakuDayT & "' "
    End If

    ' *** ヤードコード ***
    If strYardCodeF <> "" And strYardCodeT <> "" Then

        ' 共に空白ではない場合
        If strYardCodeF = strYardCodeT Then
            '　FROMとTOが同一の場合、一致条件
            strTemp = strTemp & " AND CARG_YCODE = '" & strYardCodeF & "'"
        Else
            ' FROMとTOが異なる場合、BETWEEN条件
            strTemp = strTemp & " AND CARG_YCODE BETWEEN '" & strYardCodeF & "' AND '" & strYardCodeT & "' "
        End If
        
    ElseIf strYardCodeF <> "" Then
        ' FROMのみの場合、それ以上であることが条件
        strTemp = strTemp & " AND CARG_YCODE >= '" & strYardCodeF & "' "

    ElseIf strYardCodeT <> "" Then
        ' TOのみの場合、それ以下であることが条件
        strTemp = strTemp & " AND CARG_YCODE <= '" & strYardCodeT & "' "
    End If

    ' *** 顧客コード ***
    If strUserCodeF <> "" And strUserCodeT <> "" Then

        ' 共に空白ではない場合
        If strUserCodeF = strUserCodeT Then
            '　FROMとTOが同一の場合、一致条件
            strTemp = strTemp & " AND USER_CODE = '" & strUserCodeF & "'"
        Else
            ' FROMとTOが異なる場合、BETWEEN条件
            strTemp = strTemp & " AND USER_CODE BETWEEN '" & strUserCodeF & "' AND '" & strUserCodeT & "' "
        End If

    ElseIf strUserCodeF <> "" Then
        ' FROMのみの場合、それ以上であることが条件
        strTemp = strTemp & " AND USER_CODE >= '" & strUserCodeF & "' "

    ElseIf strUserCodeT <> "" Then
        ' TOのみの場合、それ以下であることが条件
        strTemp = strTemp & " AND USER_CODE <= '" & strUserCodeT & "' "
    End If

    ' *** 変更内容 ***
    If strChangeKbnF <> "" And strChangeKbnT <> "" Then

        ' 共に空白ではない場合
        If strChangeKbnF = strChangeKbnT Then
            '　FROMとTOが同一の場合、一致条件
            strTemp = strTemp & " AND USER_CHG_RKBN = '" & strChangeKbnF & "'"
        Else
            ' FROMとTOが異なる場合、BETWEEN条件
            strTemp = strTemp & " AND USER_CHG_RKBN BETWEEN '" & strChangeKbnF & "' AND '" & strChangeKbnT & "' "
        End If

    ElseIf strChangeKbnF <> "" Then
        ' FROMのみの場合、それ以上であることが条件
        strTemp = strTemp & " AND USER_CHG_RKBN >= '" & strChangeKbnF & "' "

    ElseIf strChangeKbnT <> "" Then
        ' TOのみの場合、それ以下であることが条件
        strTemp = strTemp & " AND USER_CHG_RKBN <= '" & strChangeKbnT & "' "
    End If

    ' *** 変更日付 ***
    If strChangeDayF <> "" And strChangeDayT <> "" Then

        ' 共に空白ではない場合
        If strChangeDayF = strChangeDayT Then
            '　FROMとTOが同一の場合、一致条件
            strTemp = strTemp & " AND USER_CHG_DATE = '" & strChangeDayF & "'"
        Else
            ' FROMとTOが異なる場合、BETWEEN条件
            strTemp = strTemp & " AND USER_CHG_DATE BETWEEN '" & strChangeDayF & "' AND '" & strChangeDayT & "' "
        End If
        
    ElseIf strChangeDayF <> "" Then
        ' FROMのみの場合、それ以上であることが条件
        strTemp = strTemp & " AND USER_CHG_DATE >= '" & strChangeDayF & "' "

    ElseIf strChangeDayT <> "" Then
        ' TOのみの場合、それ以下であることが条件
        strTemp = strTemp & " AND USER_CHG_DATE <= '" & strChangeDayT & "' "
    End If

    ' *** 保証区分 ***
    If strHosyoCodeF <> "" And strHosyoCodeT <> "" Then

        ' 共に空白ではない場合
        If strHosyoCodeF = strHosyoCodeT Then
            '　FROMとTOが同一の場合、一致条件
            strTemp = strTemp & " AND CARG_HOSYICD = " & strHosyoCodeF
        Else
            ' FROMとTOが異なる場合、BETWEEN条件
            strTemp = strTemp & " AND CARG_HOSYICD BETWEEN " & strHosyoCodeF & " AND " & strHosyoCodeT
        End If
        
    ElseIf strHosyoCodeF <> "" Then
        ' FROMのみの場合、それ以上であることが条件
        strTemp = strTemp & " AND CARG_HOSYICD >= " & strHosyoCodeF

    ElseIf strHosyoCodeT <> "" Then
        ' TOのみの場合、それ以下であることが条件
        strTemp = strTemp & " AND CARG_HOSYICD <= " & strHosyoCodeT
    End If

    fncMakeBetween = strTemp

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
  
        Call .Fields.Append(.CreateField("顧客番号", DataTypeEnum.dbText, 6))               '顧客番号
        Call .Fields.Append(.CreateField("変更区分", DataTypeEnum.dbInteger, 2))            '変更区分
        Call .Fields.Append(.CreateField("変更内容", DataTypeEnum.dbText, 20))              '変更内容
        Call .Fields.Append(.CreateField("契約者", DataTypeEnum.dbText, 36))                '契約者
        Call .Fields.Append(.CreateField("契約者カナ", DataTypeEnum.dbText, 36))                '契約者
        
        Call .Fields.Append(.CreateField("代表・担当", DataTypeEnum.dbText, 18))            '代表・担当
        Call .Fields.Append(.CreateField("代表・担当カナ", DataTypeEnum.dbText, 36))            '代表・担当
        
        Call .Fields.Append(.CreateField("郵便番号", DataTypeEnum.dbText, 10))              '郵便番号
        Call .Fields.Append(.CreateField("区分", DataTypeEnum.dbInteger, 2))                '区分
        Call .Fields.Append(.CreateField("区分名", DataTypeEnum.dbText, 20))                '区分名
        Call .Fields.Append(.CreateField("住所1", DataTypeEnum.dbText, 120))                '住所
        Call .Fields.Append(.CreateField("住所2", DataTypeEnum.dbText, 120))                '住所
        Call .Fields.Append(.CreateField("TEL", DataTypeEnum.dbText, 15))                   'TEL
        Call .Fields.Append(.CreateField("FAX", DataTypeEnum.dbText, 15))                   'FAX
        Call .Fields.Append(.CreateField("携帯", DataTypeEnum.dbText, 15))                  '携帯
        
        
        Call .Fields.Append(.CreateField("連絡先担当者名", DataTypeEnum.dbText, 120))                '区分
        Call .Fields.Append(.CreateField("連絡先担当者名カナ", DataTypeEnum.dbText, 120))            '20251216 M.HONDA INS
        
        
        Call .Fields.Append(.CreateField("連絡先郵便番号", DataTypeEnum.dbText, 120))                '区分名
        Call .Fields.Append(.CreateField("連絡先住所１", DataTypeEnum.dbText, 120))                '住所
        Call .Fields.Append(.CreateField("連絡先住所２", DataTypeEnum.dbText, 120))                '住所
        Call .Fields.Append(.CreateField("連絡先住所３", DataTypeEnum.dbText, 120))                   'TEL
        Call .Fields.Append(.CreateField("連絡先電話番号", DataTypeEnum.dbText, 120))                   'FAX
        Call .Fields.Append(.CreateField("連絡先携帯番号", DataTypeEnum.dbText, 120))                  '携帯
        Call .Fields.Append(.CreateField("連絡先ＦＡＸ", DataTypeEnum.dbText, 120))                  '携帯
       
        
        
        
        
        
        Call .Fields.Append(.CreateField("明細順", DataTypeEnum.dbInteger, 5))              '明細順
        Call .Fields.Append(.CreateField("契約番号", DataTypeEnum.dbText, 10))              '契約番号
        Call .Fields.Append(.CreateField("保証No", DataTypeEnum.dbText, 20))                '保証No
        
        '----20091110----M.RYU----add-----<s>
        Call .Fields.Append(.CreateField("保証会社ｺｰﾄﾞ", DataTypeEnum.dbText, 6))            '保証区分
        Call .Fields.Append(.CreateField("保証会社名", DataTypeEnum.dbText, 50))             '保証区分名
        '----20091110----M.RYU----add-----<e>
        
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))            'ヤードコード
        Call .Fields.Append(.CreateField("ヤード名", DataTypeEnum.dbText, 36))               'ヤード名
        Call .Fields.Append(.CreateField("使用物件", DataTypeEnum.dbText, 6))                '使用物件(コンテナNo)
        Call .Fields.Append(.CreateField("保証区分", DataTypeEnum.dbText, 1))                '保証区分
        Call .Fields.Append(.CreateField("保証区分名", DataTypeEnum.dbText, 20))             '保証区分名

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

    Dim dbAccess        As Database
    Dim rsDestination   As Recordset
    Dim blnError        As Boolean

    Dim intLoopCount    As Integer
    Dim strUserCodeBK   As String

    On Error GoTo ErrorHandler

    blnError = False

    Set dbAccess = CurrentDb

    ' ワークテーブルクリア
    Call subClearWork(dbAccess, P_WORK_TABLE)

    ' ワークテーブルのレコードセット
    Set rsDestination = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)

    intLoopCount = 0
    strUserCodeBK = ""

    With rsSource
        While Not .EOF

            If strUserCodeBK <> .Fields("USER_CODE") Then
                intLoopCount = 0
            End If

            intLoopCount = intLoopCount + 1

            rsDestination.AddNew

            rsDestination.Fields("顧客番号") = Format(.Fields("USER_CODE"), "000000")
            rsDestination.Fields("変更区分") = .Fields("USER_CHG_RKBN")
            rsDestination.Fields("変更内容") = .Fields("CHG_RKBN_NAME")
            rsDestination.Fields("契約者") = .Fields("USER_NAME")
            rsDestination.Fields("契約者カナ") = .Fields("USER_KANA")
            rsDestination.Fields("代表・担当") = .Fields("USER_TANM")
            rsDestination.Fields("代表・担当カナ") = .Fields("USER_TAKA")
            
            rsDestination.Fields("郵便番号") = .Fields("USER_YUBINO")
            rsDestination.Fields("区分") = .Fields("USER_KKBN")
            rsDestination.Fields("区分名") = .Fields("KKBN_NAME")
            rsDestination.Fields("住所1") = .Fields("USER_ADR_1") & .Fields("USER_ADR_2")
            rsDestination.Fields("住所2") = .Fields("USER_ADR_3")
            rsDestination.Fields("TEL") = .Fields("USER_TEL")
            rsDestination.Fields("FAX") = .Fields("USER_FAX")
            rsDestination.Fields("携帯") = .Fields("USER_KEITAI")
            
            
            
            rsDestination.Fields("連絡先担当者名") = .Fields("USER_RNAME")
            rsDestination.Fields("連絡先担当者名カナ") = .Fields("USER_RKANA")  '20251216 M.HONDA INS
            rsDestination.Fields("連絡先郵便番号") = .Fields("USER_RPOST")
            rsDestination.Fields("連絡先住所１") = .Fields("USER_RADR_1")
            rsDestination.Fields("連絡先住所２") = .Fields("USER_RADR_2")
            rsDestination.Fields("連絡先住所３") = .Fields("USER_RADR_3")
            rsDestination.Fields("連絡先電話番号") = .Fields("USER_RTEL")
            rsDestination.Fields("連絡先携帯番号") = .Fields("USER_RKEITAI")
            rsDestination.Fields("連絡先ＦＡＸ") = .Fields("USER_RFAX")
      
            
            
            
            
            
            rsDestination.Fields("明細順") = intLoopCount
            rsDestination.Fields("契約番号") = .Fields("CARG_ACPTNO")
            rsDestination.Fields("保証No") = .Fields("CARG_HOSYB")
            
            '----20091110----M.RYU----add----------<s>
            rsDestination.Fields("保証会社ｺｰﾄﾞ") = Format(.Fields("CARG_HOSYO_CD"), "000000")
            rsDestination.Fields("保証会社名") = .Fields("SHIRM_SHIRN")
            '----20091110----M.RYU----add----------<e>
            
            rsDestination.Fields("ヤードコード") = Format(.Fields("CARG_YCODE"), "000000")
            rsDestination.Fields("ヤード名") = .Fields("YARD_NAME")
            rsDestination.Fields("使用物件") = Format(.Fields("CARG_NO"), "000000")
            rsDestination.Fields("保証区分") = .Fields("CARG_HOKAI")
            rsDestination.Fields("保証区分名") = .Fields("HOKAI_NAME")

            ' 顧客コード退避
            strUserCodeBK = .Fields("USER_CODE")

            ' 更新要の場合の処理
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
'       MODULE_NAME     : KOMSデータベース接続
'       MODULE_ID       : subConnectServer
'       CREATE_DATE     :
'       PARAM           : データベースオブジェクト
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectServer(ByRef dbSQLServer As Database)

    Dim strSqlserver    As String
    Dim strConnect      As String
    Dim strBUMOC        As String
    
    On Error GoTo ErrorHandler

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

    Dim strSqlserver    As String
    Dim strParam        As String

    On Error GoTo ErrorHandler

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

    Dim strConnectString    As String

    On Error GoTo ErrorHandler

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
