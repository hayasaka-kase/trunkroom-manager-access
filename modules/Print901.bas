Attribute VB_Name = "Print901"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤードメンテナンス依頼書出力
'   プログラムＩＤ　：　Print901
'   作　成　日　　　：  2007/02/12
'   作　成　者　　　：  イーグルソフト 鈴木
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :   2011/07/30
'   UPDATER         :   M.RYU
'   Ver             :   0.1
'   変更内容        :   検索条件を追加：修理区分・完了日
'
'   UPDATE          :   2021/09/17
'   UPDATER         :   N.IMAI
'   Ver             :   0.2
'   変更内容        :   「入力区分」追加
'                       抽出条件に「メンテ依頼」を追加
'
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

'ワークテーブル名
Private Const P_WORK_TABLE                  As String = "RVS901_W01"

'レポート名
Private Const P_REPORT                      As String = "RVS901"

'名称ID
Private Const P_NAME_修理区分               As String = "254"

'コード値
Private Const P_INPTI_ヤード                As String = "01"     '入力対象　ヤード
Private Const P_TYPEC_修理                  As String = "02"     '入力区分　修理
Private Const P_TYPEC_メンテ依頼            As String = "07"     '入力区分　メンテ依頼 I:2021/09/17

Sub a00Test_fncPrintYardMaintenanceRequest()

    If Not fncPrintYardMaintenanceRequest(P901_MODE_PREVIEW, "", "000001", "999999", True) Then
        MsgBox "False"
    End If

'    If Not fncPrintYardMaintenanceRequest(P901_MODE_EXCEL, "", "000001", "999999", True) Then
'        MsgBox "False"
'    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ヤードメンテナンス依頼書出力
'       MODULE_ID       : fncPrintMaintenanceRequest
'       CREATE_DATE     : 2007/02/12
'                       :
'       PARAM           : intMode     - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'                       : strUniqec   - 修繕トランユニークコード（省略可）
'                       : strYardFrom - ヤードコード範囲ＦＲＯＭ（省略可）
'                       : strYardTo   - ヤードコード範囲ＴＯ（省略可）
'                       : blnComplete - True=完了済み含む False=完了済みは省く（省略時False）
'                       :
'       NOTE            : 1.ユニークコードが指定された場合、以下は全て無視する
'                       : 2.ヤードコードFROMとTOが異なる場合、BETWEEN条件で検索
'                       : 3.ヤードコードFROMのみの場合、それ以上であることを条件に検索
'                       :   ただし、部屋番号が指定されている場合は、ヤードFROMと一致条件で、
'                       :   かつ部屋番号が一致する条件になる
'                       : 4.ヤードコードTOのみの場合、それ以下であることを条件に検索
'                       : 5.ヤードコードが共に指定なしの場合、ヤードコードは検索条件としない
'                       : 6.入力対照=コンテナ かつ 入力区分=修理 のデータのみ検索
'                       : 7.ヤード表示期間のヤードのみ検索対象
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'DELETE 2011/07/30 M.RYU
'Public Function fncPrintYardMaintenanceRequest(intMode As Integer, _
'                                               Optional strUniqec As String = "", _
'                                               Optional strYardFrom As String = "", _
'                                               Optional strYardTo As String = "", _
'                                               Optional blnComplete As Boolean = False _
'                                              ) As Boolean
'INSERT 2011/07/30 M.RYU
Public Function fncPrintYardMaintenanceRequest(intMode As Integer, _
                                               Optional strUniqec As String = "", _
                                               Optional strYardFrom As String = "", _
                                               Optional strYardTo As String = "", _
                                               Optional blnComplete As Boolean = False, _
                                               Optional strReprFrom As String = "", _
                                               Optional strReprTo As String = "", _
                                               Optional strCompdFrom As String = "", _
                                               Optional strCompdTo As String = "") As Boolean
On Error GoTo ErrorHandler

    Dim dbSQLServer     As Database
    Dim rsGetData       As Recordset
    Dim blnError        As Boolean

    blnError = False

    fncPrintYardMaintenanceRequest = False

    'DB接続
    Call subConnectServer(dbSQLServer)

    'データ検索
'    DELETE 2011/07/30 M.RYU
'    If Not fncGetData(dbSqlServer, rsGetData, strUniqec, strYardFrom, strYardTo, blnComplete) Then
    
    'INSERT 2011/07/30 M.RYU
    If Not fncGetData(dbSQLServer, rsGetData, strUniqec, strYardFrom, _
            strYardTo, blnComplete, strReprFrom, strReprTo, strCompdFrom, strCompdTo) Then
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

    fncPrintYardMaintenanceRequest = True

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, "fncPrintYardMaintenanceRequest" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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
'       CREATE_DATE     : 2007/02/12
'                       :
'       PARAM           : dbSqlServer - KOMSに接続したデータベースオブジェクト
'                       : rsGetData   - 検索結果を格納するレコードセット
'                       : strUniqec   - 修繕トランユニークコード
'                       : strYardFrom - ヤードコード範囲ＦＲＯＭ
'                       : strYardTo   - ヤードコード範囲ＴＯ
'                       : blnComplete - True=完了済み含む False=完了済みは省く
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'DELETE 2011/07/30 M.RYU
'Private Function fncGetData(dbSqlServer As Database, _
'                            ByRef rsGetData As Recordset, _
'                            strUniqec As String, _
'                            strYardFrom As String, _
'                            strYardTo As String, _
'                            blnComplete As Boolean _
'                            ) As Boolean

'INSERT 2011/07/30 M.RYU
Private Function fncGetData(dbSQLServer As Database, _
                            ByRef rsGetData As Recordset, _
                            strUniqec As String, _
                            strYardFrom As String, _
                            strYardTo As String, _
                            blnComplete As Boolean, _
                            strReprFrom As String, _
                            strReprTo As String, _
                            strCompdFrom As String, _
                            strCompdTo As String) As Boolean
On Error GoTo ErrorHandler

    Dim strSQL      As String
    
    fncGetData = False
    
    'SQL文作成
'    strSQL = fncMakeGetDataSql(strUniqec, strYardFrom, strYardTo, blnComplete) 'DELETE 2011/07/30 M.RYU
    strSQL = fncMakeGetDataSql(strUniqec, strYardFrom, strYardTo, blnComplete, _
                             strReprFrom, strReprTo, strCompdFrom, strCompdTo)  'INSERT 2011/07/30 M.RYU
    
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
'       CREATE_DATE     : 2006/07/27
'                       :
'       PARAM           : strUniqec   - 修繕トランユニークコード
'                       : strYardFrom - ヤードコード範囲ＦＲＯＭ
'                       : strYardTo   - ヤードコード範囲ＴＯ
'                       : blnComplete - True=完了済み含む False=完了済みは省く
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'DELETE 2011/07/30 M.RYU
'Private Function fncMakeGetDataSql(strUniqec As String, _
'                                   strYardFrom As String, _
'                                   strYardTo As String, _
'                                   blnComplete As Boolean _
'                                   ) As String
'INSERT 2011/07/30 M.RYU
Private Function fncMakeGetDataSql(strUniqec As String, _
                                   strYardFrom As String, _
                                   strYardTo As String, _
                                   blnComplete As Boolean, _
                                   strReprFrom As String, _
                                   strReprTo As String, _
                                   strCompdFrom As String, _
                                   strCompdTo As String) As String

    Dim strSQL              As String
    Dim strOpenRowSetSql    As String

    strSQL = strSQL & " SELECT KAGTOS.BUMOM_BUMON, "     ' 部門マスタ.部門名
    strSQL = strSQL & "        KAGTOS.REPRT_GENTD, "     ' 部門マスタ.発生日
    strSQL = strSQL & "        YARD_CODE, "              ' ヤードマスタ.ヤードコード
    strSQL = strSQL & "        YARD_NAME, "              ' ヤードマスタ.ヤード名
    strSQL = strSQL & "        YARD_ADDR_1, "            ' ヤードマスタ.ヤード住所１
    strSQL = strSQL & "        YARD_ADDR_2, "            ' ヤードマスタ.ヤード住所２
    strSQL = strSQL & "        YARD_ADDR_3, "            ' ヤードマスタ.ヤード住所３
    strSQL = strSQL & "        YARD_BEGIN_DAY, "         ' ヤードマスタ.営業開始日
    strSQL = strSQL & "        YARD_END_DAY, "           ' ヤードマスタ.営業終了日
    strSQL = strSQL & "        YARD_RENTEND_DAY, "       ' ヤードマスタ.ヤード解約予定
    strSQL = strSQL & "        REPR_TYPEC_NAME, "        ' 修繕トラン.入力区分名 INSERT 2021/09/17 N.IMAI
    strSQL = strSQL & "        REPR_CODE_NAME, "         ' 修繕トラン.修理区分名
    
    strSQL = strSQL & "        KAGTOS.TINHM_BUKHC, "
    
    strSQL = strSQL & "        KAGTOS.REPRT_CNT1N, "     ' 修繕トラン.内容１
    strSQL = strSQL & "        KAGTOS.REPRT_CNT2N, "     ' 修繕トラン.内容２
    strSQL = strSQL & "        KAGTOS.REPRT_CNT1N1, "    ' 修繕トラン.内容１サブ1
    strSQL = strSQL & "        KAGTOS.REPRT_CNT1N2, "    ' 修繕トラン.内容１サブ2
    strSQL = strSQL & "        KAGTOS.REPRT_CNT2N1, "    ' 修繕トラン.内容２サブ1
    strSQL = strSQL & "        KAGTOS.REPRT_CNT2N2, "    ' 修繕トラン.内容２サブ2
    strSQL = strSQL & "        REPR_TANTO_NAME, "        ' 修繕トラン.担当者名
    strSQL = strSQL & "        KAGTOS.REPRT_COMPD, "     ' 修繕トラン.完了日
    strSQL = strSQL & "        KAGTOS.REPRT_GYOUN, "     ' 修繕トラン.業者名
    strSQL = strSQL & "        KAGTOS.REPRT_KINGA "      ' 修繕トラン.金額

    '【ヤードマスタ】
    strSQL = strSQL & "   FROM YARD_MAST "

    'OpenRowset SQL文作成
'    strOpenRowSetSql = fncMakeOpenRowsetSql(strUniqec, strYardFrom, strYardTo, blnComplete)            'DELETE 2011/07/30 M.RYU
    strOpenRowSetSql = fncMakeOpenRowsetSql(strUniqec, strYardFrom, strYardTo, blnComplete, _
                                          strReprFrom, strReprTo, strCompdFrom, strCompdTo)              'INSERT 2011/07/30 M.RYU

    'オープンロウセットＳＱＬ変換
    strSQL = strSQL & "        INNER JOIN "
    strSQL = strSQL & fncOpenRowsetString(strOpenRowSetSql) & " KAGTOS "

    strSQL = strSQL & "           ON YARD_CODE = CONVERT(NUMERIC, KAGTOS.REPRT_YARDC) "

    'ユニークキーが指定されていない場合
    If strUniqec = "" Then
        'ヤード非表示日を過ぎている場合は対象外
        strSQL = strSQL & "  WHERE YARD_NONDISP_DAY > GETDATE() "
        strSQL = strSQL & "     OR YARD_NONDISP_DAY IS NULL "

        ' 並び換え
        strSQL = strSQL & "  ORDER BY YARD_CODE,KAGTOS.REPRT_GENTD "
    End If

    fncMakeGetDataSql = strSQL

End Function

'==============================================================================*
'
'       MODULE_NAME     : OpenRowset SQL文作成
'       MODULE_ID       : fncMakeOpenRowsetSql
'       CREATE_DATE     : 2006/07/27
'                       :
'       PARAM           : strUniqec   - 修繕トランユニークコード
'                       : strYardFrom - ヤードコード範囲ＦＲＯＭ
'                       : strYardTo   - ヤードコード範囲ＴＯ
'                       : blnComplete - True=完了済み含む False=完了済みは省く
'                       :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'DELETE 2011/07/30 M.RYU
'Private Function fncMakeOpenRowsetSql(strUniqec As String, _
'                                      strYardFrom As String, _
'                                      strYardTo As String, _
'                                      blnComplete As Boolean _
'                                      ) As String

'INSERT 2011/07/30 M.RYU
Private Function fncMakeOpenRowsetSql(strUniqec As String, _
                                      strYardFrom As String, _
                                      strYardTo As String, _
                                      blnComplete As Boolean, _
                                      strReprFrom As String, _
                                      strReprTo As String, _
                                      strCompdFrom As String, _
                                      strCompdTo As String) As String
                                      
    Dim strSQL              As String

    strSQL = strSQL & " SELECT REPRT_YARDC, "
    strSQL = strSQL & "        TINHM_BUKHC, "
    strSQL = strSQL & "        REPRT_GENTD, "
    strSQL = strSQL & "        REPRT_CNT1N, "
    strSQL = strSQL & "        REPRT_CNT2N, "
    strSQL = strSQL & "        REPRT_CNT1N1, "
    strSQL = strSQL & "        REPRT_CNT1N2, "
    strSQL = strSQL & "        REPRT_CNT2N1, "
    strSQL = strSQL & "        REPRT_CNT2N2, "
    strSQL = strSQL & "        TYPEC_NAME.CODET_NAMEN AS REPR_TYPEC_NAME, "     'INSERT 2021/09/17 N.IMAI
    strSQL = strSQL & "        TANTM_NAME.TANTM_TANTN AS REPR_TANTO_NAME, "
    strSQL = strSQL & "        REPRT_COMPD, "
    strSQL = strSQL & "        REPRT_GYOUN, "
    strSQL = strSQL & "        BUMOM_BUMON, "
    strSQL = strSQL & "        REPRC_NAME.CODET_NAMEN AS REPR_CODE_NAME, "
    strSQL = strSQL & "        REPRT_KINGA "

    '【修繕トラン】
    strSQL = strSQL & "   FROM REPR_TRAN AS REPR_TRAN "

    '【部門マスタ】
    strSQL = strSQL & "        INNER JOIN BUMO_MAST "
    strSQL = strSQL & "           ON BUMOM_BUMOC = REPRT_BUMOC "

    strSQL = strSQL & "        INNER JOIN TINH_MAST "
    strSQL = strSQL & "           ON TINHM_BUMOC = REPRT_BUMOC AND "
    strSQL = strSQL & "              TINHM_TINTC = REPRT_YARDC "


    '【コードマスタ】
    strSQL = strSQL & "        LEFT OUTER JOIN CODE_TABL AS REPRC_NAME "
    strSQL = strSQL & "           ON REPRC_NAME.CODET_SIKBC = '" & P_NAME_修理区分 & "' "
    strSQL = strSQL & "          AND REPRC_NAME.CODET_CODEC = CONVERT(FLOAT, REPRT_REPRC) "

    '【担当者マスタ】
    strSQL = strSQL & "        LEFT OUTER JOIN TANT_MAST AS TANTM_NAME "
    strSQL = strSQL & "           ON TANTM_BUMOC = REPRT_BUMOC "
    strSQL = strSQL & "          AND TANTM_TANTC = REPRT_TANTC "

    'INSERT 2021/09/17 N.IMAI Start
    strSQL = strSQL & "        LEFT OUTER JOIN CODE_TABL AS TYPEC_NAME "
    strSQL = strSQL & "           ON TYPEC_NAME.CODET_SIKBC = '250' "
    strSQL = strSQL & "          AND TYPEC_NAME.CODET_CODEC = CONVERT(FLOAT, REPRT_TYPEC) "
    'INSERT 2021/09/17 N.IMAI End

    If strUniqec = "" Then
        '入力対象＝コンテナ
        strSQL = strSQL & " WHERE REPRT_INPTI = '" & P_INPTI_ヤード & "' "

        '入力区分＝修理
        'strSql = strSql & "   AND REPRT_TYPEC = '" & P_TYPEC_修理 & "' "                                   'DELETE 2021/09/17 N.IMAI
        strSQL = strSQL & "   AND REPRT_TYPEC IN ('" & P_TYPEC_修理 & "','" & P_TYPEC_メンテ依頼 & "') "    'INSERT 2021/09/17 N.IMAI

        'ヤードの範囲条件
        strSQL = strSQL & fncMakeBetween("REPRT_YARDC", strYardFrom, strYardTo)

        If Not blnComplete Then
            '完了済みは除く　完了日がNULLのデータのみが対象
            strSQL = strSQL & " AND REPRT_COMPD IS NULL "
        End If
        
        '修理区分REPRT_REPRC    'INSERT 2011/07/30 M.RYU
        If strReprFrom <> "" Then strSQL = strSQL & " AND ISNULL(REPRT_REPRC,99) >= " & strReprFrom & " "
        If strReprTo <> "" Then strSQL = strSQL & " AND ISNULL(REPRT_REPRC,99) <= " & strReprTo & " "
        
        '完了日REPRT_COMPD      'INSERT 2011/07/30 M.RYU
        If strCompdFrom <> "" Then strSQL = strSQL & " AND ISNULL(REPRT_COMPD,99999999) >= " & strCompdFrom & " "
        If strCompdTo <> "" Then strSQL = strSQL & " AND ISNULL(REPRT_COMPD,99999999) <= " & strCompdTo & " "
        
    Else
        'ユニークコードが指定されている場合、それだけを一致条件
        strSQL = strSQL & " WHERE REPRT_UNIQEC='" & strUniqec & "' "
    End If

    fncMakeOpenRowsetSql = strSQL
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
'       MODULE_NAME     : オープンロウセットＳＱＬ変換
'       MODULE_ID       : fncOpenRowsetString
'       CREATE_DATE     :
'       PARAM           : strSQL                SQL文
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncOpenRowsetString(ByVal strSQL As String) As String

    Dim strSvr              As String
    Dim strUid              As String
    Dim strPwd              As String
    Dim strDBN              As String
    Dim iSt                 As Long
    Dim iEd                 As Long
    Dim strNew              As String

    On Error GoTo ErrorHandler

    '加瀬DBの接続文字列
    strSvr = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_SERVER_NAME'")
    strUid = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_USER_ID'")
    strPwd = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_PASSWORD'")
    strDBN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME'")

    If strPwd = "#NULL#" Then
        strPwd = ""
    End If
    
    'シングルクォートを２連にする
    iSt = 1
    iEd = InStr(iSt, strSQL, "'")
    While iEd > 0
        strNew = strNew & Mid(strSQL, iSt, iEd - iSt + 1) & "'"
        iSt = iEd + 1
        iEd = InStr(iSt, strSQL, "'")
    Wend
    strNew = strNew & Mid(strSQL, iSt)
    
    'FROMテーブル名を[データベース名].[オーナー].[テーブル名]にする
    iSt = InStr(1, strNew, " FROM ") + 5
    strNew = Left(strNew, iSt) & strDBN & ".dbo." & LTrim(Mid(strNew, iSt + 1))
    
    'JOINテーブル名を[データベース名].[オーナー].[テーブル名]にする
    iSt = 1
    Do
        iSt = InStr(iSt, strNew, " JOIN ")
        If iSt = 0 Then
            Exit Do
        End If
        iSt = iSt + 5
        strNew = Left(strNew, iSt) & strDBN & ".dbo." & LTrim(Mid(strNew, iSt + 1))
    Loop

    fncOpenRowsetString = " OPENROWSET('SQLOLEDB','" & strSvr & "';'" & strUid & "';'" & strPwd & "','" & strNew & "') "
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncOpenRowsetString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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
 
        Call .Fields.Append(.CreateField("部門名", DataTypeEnum.dbText, 36))            '部門名
        Call .Fields.Append(.CreateField("依頼日", DataTypeEnum.dbText, 14))            '依頼日
        Call .Fields.Append(.CreateField("入力担当者", DataTypeEnum.dbText, 36))        '入力担当者
        Call .Fields.Append(.CreateField("入力区分", DataTypeEnum.dbText, 100))         '入力区分   I:2021/09/17 N.IMAI

        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))       'ヤードコード
        Call .Fields.Append(.CreateField("借上物件コード", DataTypeEnum.dbText, 6))       '借上げコード
        
        Call .Fields.Append(.CreateField("ヤード名", DataTypeEnum.dbText, 36))          'ヤード名
        Call .Fields.Append(.CreateField("ヤード住所", DataTypeEnum.dbText, 200))       'ヤード住所

        Call .Fields.Append(.CreateField("営業開始日", DataTypeEnum.dbText, 10))        '営業開始日
        Call .Fields.Append(.CreateField("営業終了日", DataTypeEnum.dbText, 10))        '営業終了日
        Call .Fields.Append(.CreateField("ヤード解約予定", DataTypeEnum.dbText, 10))    'ヤード解約予定

        Call .Fields.Append(.CreateField("修理区分", DataTypeEnum.dbText, 36))          '修理区分

        Call .Fields.Append(.CreateField("依頼内容", DataTypeEnum.dbText, 80))          '依頼内容
        Call .Fields.Append(.CreateField("依頼内容サブ１", DataTypeEnum.dbText, 80))    '依頼内容サブ１
        Call .Fields.Append(.CreateField("依頼内容サブ２", DataTypeEnum.dbText, 80))    '依頼内容サブ２

        Call .Fields.Append(.CreateField("作業報告", DataTypeEnum.dbText, 80))          '作業報告
        Call .Fields.Append(.CreateField("作業報告サブ１", DataTypeEnum.dbText, 80))    '作業報告１
        Call .Fields.Append(.CreateField("作業報告サブ２", DataTypeEnum.dbText, 80))    '作業報告２

        Call .Fields.Append(.CreateField("作業完了日", DataTypeEnum.dbText, 10))        '作業完了日
        Call .Fields.Append(.CreateField("作業者", DataTypeEnum.dbText, 60))            '作業者

        Call .Fields.Append(.CreateField("金額", DataTypeEnum.dbText, 50))              '金額

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
    
    blnError = False
    
    Set dbAccess = CurrentDb
    
    'ワークテーブルクリア
    Call psubClearWork(dbAccess, P_WORK_TABLE)
    
    'ワークテーブルのレコードセット
    Set rsDestination = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    
    'データ追加
    While Not rsSource.EOF
        Call subAddNew(rsSource, rsDestination, intMode)
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
Private Sub subAddNew(rsSource As Recordset, rsDestination As Recordset, intMode As Integer)

    Dim strTemp     As String

    With rsSource
        rsDestination.AddNew

        rsDestination.Fields("部門名") = .Fields("BUMOM_BUMON")                                                    '部門名＝部門マスタ．部門名

        '印刷時とExcel出力時でフォーマットが異なる
        If intMode = P901_MODE_EXCEL Then
            strTemp = Format(fncStrToDate(.Fields("REPRT_GENTD")), "yyyy/mm/dd")
        Else
            strTemp = Format(fncStrToDate(.Fields("REPRT_GENTD")), "yyyy年mm月dd日")
        End If

        rsDestination.Fields("依頼日") = strTemp                                                                   '依頼日＝修繕トラン．発生日
        rsDestination.Fields("入力担当者") = .Fields("REPR_TANTO_NAME")                                            '

        rsDestination.Fields("ヤードコード") = Format(.Fields("YARD_CODE"), "000000")                              'ヤードコード＝ヤードマスタ．ヤードコード
        rsDestination.Fields("ヤード名") = .Fields("YARD_NAME")                                                    'ヤード名＝ヤードマスタ．ヤード名
        rsDestination.Fields("ヤード住所") = .Fields("YARD_ADDR_1") & .Fields("YARD_ADDR_2") & .Fields("YARD_ADDR_3")   'ヤード住所＝ヤードマスタ．ヤード住所１ & ヤードマスタ．ヤード住所２ & ヤードマスタ．ヤード住所３
        rsDestination.Fields("営業開始日") = .Fields("YARD_BEGIN_DAY")                                             '営業開始日＝ヤードマスタ．営業開始日
        rsDestination.Fields("営業終了日") = .Fields("YARD_RENTEND_DAY")                                            '営業開始日＝ヤードマスタ．営業終了日
        rsDestination.Fields("ヤード解約予定") = .Fields("YARD_END_DAY")                                            '営業開始日＝ヤードマスタ．ヤード解約予定

        rsDestination.Fields("修理区分") = .Fields("REPR_CODE_NAME")                                               '修理区分＝修繕トラン．修理区分の名称

        rsDestination.Fields("依頼内容") = .Fields("REPRT_CNT1N")                                                  '依頼内容＝修繕トラン．内容１
        rsDestination.Fields("依頼内容サブ１") = .Fields("REPRT_CNT1N1")                                           '依頼内容＝修繕トラン．内容１サブ1
        rsDestination.Fields("依頼内容サブ２") = .Fields("REPRT_CNT1N2")                                           '依頼内容＝修繕トラン．内容１サブ2

        rsDestination.Fields("作業報告") = .Fields("REPRT_CNT2N")                                                  '作業報告＝修繕トラン．内容２
        rsDestination.Fields("作業報告サブ１") = .Fields("REPRT_CNT2N1")                                           '作業報告＝修繕トラン．内容２サブ1
        rsDestination.Fields("作業報告サブ２") = .Fields("REPRT_CNT2N2")                                           '作業報告＝修繕トラン．内容２サブ2

        rsDestination.Fields("作業完了日") = Format(fncStrToDate(.Fields("REPRT_COMPD")), "yyyy/mm/dd")            '作業完了日＝修繕トラン．完了日
        rsDestination.Fields("作業者") = .Fields("REPRT_GYOUN")                                                    '作業者＝修繕トラン．業者名

        rsDestination.Fields("金額") = Format(.Fields("REPRT_KINGA"), "#,###,###")                                 '金額＝修繕トラン．金額
        
        rsDestination.Fields("入力区分") = .Fields("REPR_TYPEC_NAME")                   'INSERT 2021/09/17 N.IMAI
        
        rsDestination.Fields("借上物件コード") = .Fields("TINHM_BUKHC")
        
        
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
