Attribute VB_Name = "Print210"
 '****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード解約業務
'   プログラムＩＤ　：　Print210
'   作　成　日　　　：  2007/04/09
'   作　成　者　　　：  イーグルソフト 上野
'   Ver             ：  0.0
'   備考            ：
'
'   UPDATE          :   2007/07/16
'   UPDATER         :   tajima
'   Ver             :   0.1
'   変更内容        :   解約ヤード契約者一覧の取得SQLバグ修正
'                   :   代替候補一覧で契約者が53人以上だとエラーになるバグ修正
'
'   UPDATE          :   2012/01/24
'   UPDATER         :   M.RYU
'   Ver             :   0.2
'   変更内容        :   出力フォーマット修正　使用用途列を追加
'
'   UPDATE          :   2018/02/17
'   UPDATER         :   k.sato
'   Ver             :   0.3
'   変更内容        :   解約ヤード振分け表追加
'
'   UPDATE          :   2018/06/17
'   UPDATER         :   k.sato
'   Ver             :   0.4
'   変更内容        :   解約ヤード振分け表の条件に営業日有無を追加
'                       解約ヤード振分け表の条件に取り置き除外を追加
'                       解約ヤード振分け表のテンプレートファイル設定読み込みキー情報を変更
'                       解約ヤード振分け表のテンプレートファイル設定読み込みのエラーチェック追加
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'処理モード
Public Const P210_MODE_PREVIEW              As Integer = 1              '印刷プレビューを表示
Public Const P210_MODE_EXCEL                As Integer = 2              'Excelに出力
Public Const P210_MODE_PRINT                As Integer = 3              'プレビューを表示しないで印刷

'帳票区分
Public Const P210_帳票A                     As String = "1"             '解約ヤード契約者一覧
Public Const P210_帳票B                     As String = "2"             '
Public Const P210_帳票C                     As String = "3"             '振分表 INSERT 2018/02/17 add

Private Const P_WORK_TABLE                  As String = "FVS210_W01"    'ワークテーブル名
Private Const P_REPORT                      As String = "FVS210"        'レポート名

Private pstrPrintKbn                        As String                   '帳票区分

'Excel 定数
Public Const xlDown As Long = -4121
Public Const xlNone As Long = -4142
Public Const xlEdgeLeft As Long = 7
Public Const xlContinuous As Long = 1
Public Const xlThin As Long = 2
Public Const xlAutomatic As Long = -4105
Public Const xlEdgeTop As Long = 8
Public Const xlEdgeBottom As Long = 9
Public Const xlEdgeRight As Long = 10
Public Const xlInsideVertical As Long = 11
Public Const xlSolid As Long = 1
Public Const xlDiagonalDown As Long = 5
Public Const xlDiagonalUp As Long = 6
Public Const xlCenter As Long = -4108
Public Const xlBottom As Long = -4107
Public Const xlContext As Long = -5002
Public Const xlRight As Long = -4152
Public Const xlInsideHorizontal As Long = 12
Public Const xlToRight As Long = -4161
Public Const xlPasteFormulas As Long = -4122         'INSERT 2018-02-17 add start
Public Const xlHairline As Long = 1                  'INSERT 2018-02-17 add start

'INSERT 2018-02-17 add start
Public Const par_FVS210_INTIF_RECFB_1 As String = "FVS210_PATH"
Public Const par_FVS210_INTIF_RECFB_2 As String = "FVS210_FILE"

#Const test = False

' 構造体(SIZE)
Private Type Type_row_col
    row       As Long             ' ROW
    col       As Long             ' COL
End Type
Private kinrin_start As Type_row_col

' 構造体(SIZE)
Private Type Type_SIZE
    STEP        As String             ' STEP
    SIZE        As String             ' SIZE
    AGREE       As String             ' 契約数
    FREE        As String             ' 空数
End Type
Private pSIZELists() As Type_SIZE
'
Private Type Type_NYAR
    YCODE       As String             ' YAER_CODE
    YNAME       As String             ' YARD_NAME
    KIRO        As String             ' NYAR_KIRO
    END_DATE    As String             '
End Type
Private pNYARLists() As Type_NYAR
Private pNYARListsJoin As String
Private pSIZEListsUpRows As Integer
Private pSIZEListsDownRows As Integer

Public Const NYARMAX As Integer = 17  '近隣ヤードは１７件まで出力
'INSERT 2018-02-17 add end

'        ReDim pNYARLists(objConRs.RecordCount - 1)
'        ReDim pSIZELists(objConRs.RecordCount - 1)

'==============================================================================*
'
'       MODULE_NAME     : ヤード解約業務帳票出力
'       MODULE_ID       : fncPrintYardInlimit
'       CREATE_DATE     : 2007/04/09
'                       :
'       PARAM           : intMode       - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'                       : strBumonCd    - 部門コード
'                       : strYardCode   - ヤードコード
'                       : strYardName   - ヤード名
'                       : intOutSbt     - 1=帳票A  2=帳票B （定数宣言あり）
'       RETURN          : 正常(True) 該当データ無し(False)
'                       : 不正終了時は例外を発生。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncPrintYardInlimit(intMode As Integer, _
                                    strBumonCd As String, _
                                    strYardCode As String, _
                                    strYardName As String, _
                                    intOutSbt As Integer _
                                    ) As Boolean
On Error GoTo ErrorHandler

    Dim dbSQLServer     As Database
    Dim rsGetData       As Recordset
    Dim blnError        As Boolean
    Dim blnRet          As Boolean
    Dim strFilename     As String
    Dim strSQL          As String
    
    Dim xlsfilepath     As String 'INSERT 2018-06-23
    Dim xlsfilename     As String 'INSERT 2018-06-23

    blnError = False

    fncPrintYardInlimit = False
    
    xlsfilepath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & P_REPORT & "' AND INTIF_RECFB='" & par_FVS210_INTIF_RECFB_1 & "'"), "")
    xlsfilename = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & P_REPORT & "' AND INTIF_RECFB='" & par_FVS210_INTIF_RECFB_2 & "'"), "")
    
    If xlsfilepath & xlsfilename = "" Then
        MsgBox "INTI_FILE読み込みエラー" & vbCrLf & "INTIF_PROGB=" & par_FVS210_INTIF_RECFB_1 & vbCrLf & "INTIF_PROGB=" & par_FVS210_INTIF_RECFB_2
        fncPrintYardInlimit = True
        Exit Function
    End If
    
    '帳票区分セット
    pstrPrintKbn = intOutSbt

    'DB接続
    Call subConnectServer(dbSQLServer, strBumonCd)


#If test Then

    MsgBox "testのためスキップ"

#Else

    'データ検索
    If Not fncGetData(dbSQLServer, strBumonCd, rsGetData, strYardCode) Then
        '該当データ無し
        GoTo ExitRtn
    End If
    
    'ワークテーブル作成
    Call subMakeWork(rsGetData, intMode)
    
    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing

    If pstrPrintKbn = P210_帳票B Then
        '解約ヤード契約一覧取得
        strSQL = fncMakeGetDataSqlB2(strYardCode)
        Set rsGetData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
        If rsGetData.EOF Then
            '該当データ無し
            GoTo ExitRtn
        End If
    End If

#End If

    '保存ﾀﾞｲｱﾛｸﾞ表示
    strFilename = fncGetFileName(strYardCode)
    If strFilename = "" Then
        'ｷｬﾝｾﾙ時は正常終了
        fncPrintYardInlimit = True
        GoTo ExitRtn
    End If

    'EXCELファイル出力
    If pstrPrintKbn = P210_帳票A Then
        '帳票A作成
        blnRet = fncExcelFormatSaveA(strFilename, strYardCode, strYardName)
    ElseIf pstrPrintKbn = P210_帳票B Then
        '帳票B作成
        blnRet = fncExcelFormatSaveB(strFilename, strYardCode, strYardName, rsGetData)
    ElseIf pstrPrintKbn = P210_帳票C Then
        '帳票B作成
        blnRet = fncExcelFormatSaveC(strFilename, strYardCode, strYardName)
    End If
    If blnRet = False Then
        MsgBox "ファイルの自動変換に失敗しました。"
        On Error Resume Next
        doCmd.OutputTo acOutputTable, P_WORK_TABLE, "MicrosoftExcel(*.xls)", "", True, ""
        On Error GoTo ErrorHandler
    End If
    
    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

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
'       MODULE_NAME     : ワークテーブルクリア
'       MODULE_ID       : psubClearWork
'       CREATE_DATE     : 2007/04/09
'       PARAM           : dbAccess     - ACCESSデータベースオブジェクト(省略可)
'                       : rsRecord     - 検索結果が格納されたレコードセット
'                       : strTableName - テーブル名(省略可)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub psubClearWork(Optional dbAccess As Database = Null, Optional rsRecord As Recordset, _
                         Optional strTable As String = P_WORK_TABLE)

On Error GoTo ErrorHandler

    Dim tdfNew      As TableDef
    Dim rsClone     As Recordset
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
    If pstrPrintKbn = P210_帳票A Then
        Call subFieldAppendA(tdfNew)
    ElseIf pstrPrintKbn = P210_帳票B Then
        Call subFieldAppendB(tdfNew)
    ElseIf pstrPrintKbn = P210_帳票C Then            'INSERT 2018-02-17 add
        Call subFieldAppendC(tdfNew)                 'INSERT 2018-02-17 add
    End If
    Call dbAccess.TableDefs.Append(tdfNew)

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    If Not rsClone Is Nothing Then rsClone.Close: Set rsClone = Nothing
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
'       CREATE_DATE     : 2007/04/09
'                       :
'       PARAM           : dbSqlServer - KONTに接続したデータベースオブジェクト
'                       : strBumonCd  - 部門CD
'                       : rsGetData   - 検索結果を格納するレコードセット
'                       : strYardCode - ヤードコード
'                       :
'       RETURN          : 正常(True) 該当データ無し(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData(dbSQLServer As Database, _
                            ByVal strBumonCd As String, _
                            ByRef rsGetData As Recordset, _
                            ByVal strYardCode As String _
                            ) As Boolean

On Error GoTo ErrorHandler
    
    Dim strSQL      As String
    Dim blnError    As Boolean
    
    blnError = False
    fncGetData = False

    'SQL文作成
    If pstrPrintKbn = P210_帳票A Then
        strSQL = fncMakeGetDataSqlA(strBumonCd, strYardCode)
    ElseIf pstrPrintKbn = P210_帳票B Then
        strSQL = fncMakeGetDataSqlB1(strYardCode)
    'INSERT 2018-02-17 add start
    ElseIf pstrPrintKbn = P210_帳票C Then
        strSQL = fncMakeGetDataSqlC(strBumonCd, strYardCode)
    'INSERT 2018-02-17 add end
    End If

    '検索
    Set rsGetData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    'データが存在しない場合Falseを返却
    fncGetData = Not rsGetData.EOF

    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncGetData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : SQL文作成(帳票A)
'       MODULE_ID       : fncMakeGetDataSqlA
'       CREATE_DATE     : 2007/04/09
'       PARAM           : strBumonCd  - 部門CD
'                       : strYardCode - ヤードコード
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSqlA(strBumonCd As String, _
                                    strYardCode As String _
                                    ) As String

On Error GoTo ErrorHandler
    
    Dim strSQL              As String
    Dim strOpenRowSetSql    As String
    Dim blnError            As Boolean
    
    blnError = False

    strSQL = "SELECT "
    strSQL = strSQL & " CARG_STDATE " & Chr(13)
    strSQL = strSQL & "       ,USER_TEL " & Chr(13)
    strSQL = strSQL & "       ,USER_KEITAI " & Chr(13)
    strSQL = strSQL & "       ,USER_FAX " & Chr(13)
    strSQL = strSQL & "       ,CARG_UCODE " & Chr(13)
    strSQL = strSQL & "       ,USER_NAME " & Chr(13)
    strSQL = strSQL & "       ,CARG_NO " & Chr(13)
    strSQL = strSQL & "       ,CM1.CNTA_SIZE CNTA_SIZE1 " & Chr(13)
    strSQL = strSQL & "       ,STEP_NAME1.NAME_NAME STEP_NAME1  " & Chr(13)
    strSQL = strSQL & "       ,USAGE_NAME1.NAME_NAME USAGE_NAME1 " & Chr(13)      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "       ,ISNULL(CARG_RENTKG,0) + ISNULL(CARG_SYOZEI,0) KINGAKU1 " & Chr(13)
    strSQL = strSQL & "       ,HOSYICD_NAME1.NAME_NAME HOSYICD_NAME1 " & Chr(13)
    strSQL = strSQL & "       ,CASE WHEN CARG_HOSYICD IN(3, 4, 6) THEN 0 ELSE CARG_SECUKG END CARG_SECUKG " & Chr(13)
    strSQL = strSQL & "       ,KKBN_NAME.NAME_NAME KKBN_NAME " & Chr(13)
    strSQL = strSQL & "       ,SKBN_NAME.NAME_NAME SKBN_NAME " & Chr(13)
    strSQL = strSQL & "       ,MISY.TANTM_TANTN    TANTM_TANTN"
    strSQL = strSQL & "       ,CARG_KEY_RETDATE " & Chr(13)
    strSQL = strSQL & "       ,CONVERT(VARCHAR,YOUKT_YCODE) YOUKT_YCODE " & Chr(13)
    strSQL = strSQL & "       ,YARD_NAME " & Chr(13)
    strSQL = strSQL & "       ,YOUKT_NO " & Chr(13)
    strSQL = strSQL & "       ,CM2.CNTA_SIZE CNTA_SIZE2 " & Chr(13)
    strSQL = strSQL & "       ,STEP_NAME2.NAME_NAME STEP_NAME2 " & Chr(13)
    strSQL = strSQL & "       ,USAGE_NAME2.NAME_NAME USAGE_NAME2 " & Chr(13)      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "       ,ISNULL(RCPT_RENTKG,0) + ISNULL(RCPT_EZAPPI,0) + ISNULL(RCPT_ADD_EZAPPI1,0) + ISNULL(RCPT_ADD_EZAPPI2,0) KINGAKU2 " & Chr(13)
    strSQL = strSQL & "       ,HOSYICD_NAME2.NAME_NAME HOSYICD_NAME2 " & Chr(13)
    strSQL = strSQL & "       ,MOVETANTO_NAME.NAME_NAME MOVETANTO_NAME " & Chr(13)
    strSQL = strSQL & "       ,CASE WHEN RCPT_HOSYICD IN(3, 4, 6) THEN 0 ELSE RCPT_SECUKG END RCPT_SECUKG " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "       CARG_FILE " & Chr(13)
    strSQL = strSQL & "       INNER JOIN CNTA_MAST CM1 ON " & Chr(13)
    strSQL = strSQL & "           CM1.CNTA_CODE = CARG_YCODE AND CM1.CNTA_NO = CARG_NO " & Chr(13)
    strSQL = strSQL & "       INNER JOIN USER_MAST ON " & Chr(13)
    strSQL = strSQL & "           USER_CODE = CARG_UCODE " & Chr(13)
    
    'OpenRowset SQL文作成
    strOpenRowSetSql = fncMakeOpenRowsetSql()

    'OpenRowset SQL変換
    strSQL = strSQL & "       LEFT OUTER JOIN "
    strSQL = strSQL & fncOpenRowsetString(strOpenRowSetSql) & " MISY "
    strSQL = strSQL & "           ON REPLACE(STR(CARG_UCODE,6),' ','0') = MISYT_KOKYC AND MISYT_BUMOC = '" & strBumonCd & "' "
    
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME1 ON " & Chr(13)
    strSQL = strSQL & "           STEP_NAME1.NAME_ID = '090' AND STEP_NAME1.NAME_CODE = CM1.CNTA_STEP " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST HOSYICD_NAME1 ON " & Chr(13)
    strSQL = strSQL & "           HOSYICD_NAME1.NAME_ID = '200' AND HOSYICD_NAME1.NAME_CODE = CARG_HOSYICD " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST KKBN_NAME ON " & Chr(13)
    strSQL = strSQL & "           KKBN_NAME.NAME_ID = '020' AND KKBN_NAME.NAME_CODE = USER_KKBN " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST SKBN_NAME ON " & Chr(13)
    strSQL = strSQL & "           SKBN_NAME.NAME_ID = '030' AND SKBN_NAME.NAME_CODE = USER_SKBN " & Chr(13)
    
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST USAGE_NAME1 ON " & Chr(13)                                      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "           USAGE_NAME1.NAME_ID = '086' AND USAGE_NAME1.NAME_CODE = CM1.CNTA_USAGE " & Chr(13)    'INSERT 2012/01/24 M.RYU
        
    strSQL = strSQL & "       LEFT OUTER JOIN (YOUK_TRAN " & Chr(13)
    strSQL = strSQL & "       INNER JOIN YARD_MAST ON " & Chr(13)
    strSQL = strSQL & "           YOUKT_YCODE = YARD_CODE " & Chr(13)
    strSQL = strSQL & "           AND YOUKT_YUKBN IN(10,20) " & Chr(13) '2007/07/16 add tajima 受付&契約中の予約トランを対象に
    strSQL = strSQL & "       INNER JOIN CNTA_MAST CM2 ON " & Chr(13)
    strSQL = strSQL & "           YOUKT_YCODE = CM2.CNTA_CODE AND YOUKT_NO = CM2.CNTA_NO " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN RCPT_TRAN ON " & Chr(13)
    strSQL = strSQL & "           YOUKT_UKNO  = RCPT_NO " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME2 ON " & Chr(13)
    strSQL = strSQL & "           STEP_NAME2.NAME_ID = '090' AND STEP_NAME2.NAME_CODE = CM2.CNTA_STEP " & Chr(13)
    strSQL = strSQL & "       LEFT  OUTER JOIN NAME_MAST HOSYICD_NAME2 ON " & Chr(13)
    strSQL = strSQL & "           HOSYICD_NAME2.NAME_ID = '200' AND HOSYICD_NAME2.NAME_CODE = RCPT_HOSYICD " & Chr(13)
    strSQL = strSQL & "       LEFT  OUTER JOIN NAME_MAST MOVETANTO_NAME ON " & Chr(13)
    strSQL = strSQL & "           MOVETANTO_NAME.NAME_ID = '084' AND MOVETANTO_NAME.NAME_CODE = YOUKT_MOVE_TANTO " & Chr(13)
        
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST USAGE_NAME2 ON " & Chr(13)                                      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "           USAGE_NAME2.NAME_ID = '086' AND USAGE_NAME2.NAME_CODE = CM2.CNTA_USAGE " & Chr(13)    'INSERT 2012/01/24 M.RYU
    
    strSQL = strSQL & "       ) ON YOUKT_MOTO_ACPTNO = CARG_ACPTNO " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "      CARG_YCODE  = '" & strYardCode & "' " & Chr(13)
    strSQL = strSQL & "  AND CARG_AGRE   <> 9 " & Chr(13)
'    strSQL = strSQL & "   AND ISNULL(CARG_KYDATE,'9999/12/31') > '" & a営業終了日 & "' " & Chr(13)
    strSQL = strSQL & "ORDER BY " & Chr(13)
    strSQL = strSQL & "      USER_KANA "

    fncMakeGetDataSqlA = strSQL

    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncMakeGetDataSqlA" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : SQL文作成(帳票B -1)
'       MODULE_ID       : fncMakeGetDataSqlB1
'       CREATE_DATE     : 2007/04/09
'                       : strYardCode - ヤードコード
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSqlB1(strYardCode As String) As String

On Error GoTo ErrorHandler
    
    Dim strSQL              As String
    Dim blnError            As Boolean
    
    blnError = False

    strSQL = "SELECT "
    strSQL = strSQL & " YARD_CODE " & Chr(13)
    strSQL = strSQL & "       ,YARD_NAME " & Chr(13)
    strSQL = strSQL & "       ,CNTA_NO " & Chr(13)
    strSQL = strSQL & "       ,USAGE_NAME.NAME_NAME USAGE_NAME " & Chr(13)
    strSQL = strSQL & "       ,STEP_NAME.NAME_NAME STEP_NAME " & Chr(13)
    strSQL = strSQL & "       ,CNTA_SIZE " & Chr(13)
    strSQL = strSQL & "       ,PRIC_PRICE " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "       YARD_MAST" & Chr(13)
    strSQL = strSQL & "       INNER JOIN (" & Chr(13)
    strSQL = strSQL & "       CNTA_MAST  " & Chr(13)
    strSQL = strSQL & "       INNER JOIN PRIC_TABL ON " & Chr(13)
    strSQL = strSQL & "           PRIC_YCODE = CNTA_CODE AND PRIC_USAGE = CNTA_USAGE AND " & Chr(13)
    strSQL = strSQL & "           PRIC_SIZE = CNTA_SIZE AND CNTA_STEP = PRIC_STEP " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST USAGE_NAME ON " & Chr(13)
    strSQL = strSQL & "           USAGE_NAME.NAME_ID = '086' AND USAGE_NAME.NAME_CODE = CNTA_USAGE " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME ON " & Chr(13)
    strSQL = strSQL & "           STEP_NAME.NAME_ID = '090' AND STEP_NAME.NAME_CODE = CNTA_STEP " & Chr(13)
    strSQL = strSQL & "       ) ON CNTA_CODE = YARD_CODE " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "       YARD_INLIMIT_YCODE  = '" & strYardCode & "' " & Chr(13)
    strSQL = strSQL & "  AND  ISNULL(YARD_INLIMIT_DAY, '1900/01/01') >= '" & Format$(DATE, "YYYY/MM/DD") & "'" & Chr(13)
    
    strSQL = strSQL & "  AND  NOT EXISTS ( " & Chr(13)
    strSQL = strSQL & "       SELECT * FROM CARG_FILE  " & Chr(13)
    strSQL = strSQL & "       WHERE CARG_YCODE = CNTA_CODE " & Chr(13)
    strSQL = strSQL & "       AND   CARG_NO    = CNTA_NO " & Chr(13)
    strSQL = strSQL & "       AND   CARG_AGRE <> '9' ) " & Chr(13)
    strSQL = strSQL & "  AND  NOT EXISTS ( " & Chr(13)
    strSQL = strSQL & "       SELECT * FROM INTR_TRAN  " & Chr(13)
    strSQL = strSQL & "       WHERE INTRT_YCODE = CNTA_CODE " & Chr(13)
    strSQL = strSQL & "       AND   INTRT_NO    = CNTA_NO " & Chr(13)
    strSQL = strSQL & "       AND   INTRT_INTROKBN IN('1','2') ) " & Chr(13)
    strSQL = strSQL & "ORDER BY " & Chr(13)
    strSQL = strSQL & "       PRIC_PRICE, CNTA_SIZE, CNTA_STEP "

    fncMakeGetDataSqlB1 = strSQL

    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncMakeGetDataSqlB1" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : SQL文作成(帳票B -2)
'       MODULE_ID       : fncMakeGetDataSqlB2
'       CREATE_DATE     : 2007/04/09
'                       : strYardCode - ヤードコード
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSqlB2(strYardCode As String) As String

On Error GoTo ErrorHandler
    
    Dim strSQL              As String
    Dim blnError            As Boolean
    
    blnError = False

    strSQL = "SELECT "
    strSQL = strSQL & " CARG_NO " & Chr(13)
    strSQL = strSQL & "       ,USAGE_NAME.NAME_NAME USAGE_NAME " & Chr(13)
    strSQL = strSQL & "       ,USER_CODE " & Chr(13)
    strSQL = strSQL & "       ,USER_NAME " & Chr(13)
    strSQL = strSQL & "       ,CNTA_SIZE " & Chr(13)
    strSQL = strSQL & "       ,STEP_NAME.NAME_NAME STEP_NAME " & Chr(13)
    strSQL = strSQL & "       ,ISNULL(CARG_RENTKG,0) + ISNULL(CARG_SYOZEI,0) KINGAKU " & Chr(13)
    strSQL = strSQL & "       ,YOUKT_YCODE " & Chr(13)
    strSQL = strSQL & "       ,YOUKT_NO " & Chr(13)
    strSQL = strSQL & "       ,YOUKT_YUKBN " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "       CARG_FILE " & Chr(13)
    strSQL = strSQL & "       INNER JOIN CNTA_MAST ON " & Chr(13)
    strSQL = strSQL & "           CNTA_CODE = CARG_YCODE AND CNTA_NO = CARG_NO " & Chr(13)
    strSQL = strSQL & "       INNER JOIN USER_MAST ON " & Chr(13)
    strSQL = strSQL & "           USER_CODE = CARG_UCODE " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST USAGE_NAME ON " & Chr(13)
    strSQL = strSQL & "           USAGE_NAME.NAME_ID = '086' AND USAGE_NAME.NAME_CODE = CNTA_USAGE " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME ON " & Chr(13)
    strSQL = strSQL & "           STEP_NAME.NAME_ID = '090' AND STEP_NAME.NAME_CODE = CNTA_STEP " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN YOUK_TRAN ON " & Chr(13)
    strSQL = strSQL & "           CARG_ACPTNO = YOUKT_MOTO_ACPTNO AND YOUKT_YUKBN <> 53 " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "       CARG_YCODE  = '" & strYardCode & "' " & Chr(13)
    strSQL = strSQL & "  AND  CARG_AGRE   <> 9 " & Chr(13)
    strSQL = strSQL & "ORDER BY " & Chr(13)
    strSQL = strSQL & "       USER_KANA "

    fncMakeGetDataSqlB2 = strSQL

    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncMakeGetDataSqlB2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : SQL文作成(帳票C)
'       MODULE_ID       : fncMakeGetDataSqlC
'       CREATE_DATE     : 2018/02/17
'       PARAM           : strBumonCd  - 部門CD
'                       : strYardCode - ヤードコード
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSqlC(strBumonCd As String, _
                                    strYardCode As String _
                                    ) As String

On Error GoTo ErrorHandler
    
    Dim strSQL              As String
    Dim strOpenRowSetSql    As String
    Dim blnError            As Boolean
    
    blnError = False

    strSQL = "SELECT "
    strSQL = strSQL & " CARG_STDATE " & Chr(13)
    strSQL = strSQL & "       ,USER_TEL " & Chr(13)
    strSQL = strSQL & "       ,USER_KEITAI " & Chr(13)
    strSQL = strSQL & "       ,USER_FAX " & Chr(13)
    'UPDATE 2018-06-21
    'strSQL = strSQL & "       ,CARG_UCODE " & Chr(13)
    'strSQL = strSQL & "       ,USER_NAME " & Chr(13)
    strSQL = strSQL & "         ,CASE "
    strSQL = strSQL & "            WHEN CARG_UCODE is null and yt2.youkt_yukbn = 2 THEN 999999 "
    strSQL = strSQL & "            Else CARG_UCODE "
    strSQL = strSQL & "          END CARG_UCODE "
    strSQL = strSQL & "         ,CASE "
    strSQL = strSQL & "            WHEN USER_NAME is null and yt2.youkt_yukbn = 2 THEN yt2.youkt_name "
    strSQL = strSQL & "            Else USER_NAME "
    strSQL = strSQL & "          END USER_NAME "
    'UPDATE 2018-06-21
    strSQL = strSQL & "       ,CARG_NO " & Chr(13)
    strSQL = strSQL & "       ,CM1.CNTA_SIZE CNTA_SIZE1 " & Chr(13)
    strSQL = strSQL & "       ,STEP_NAME1.NAME_NAME STEP_NAME1  " & Chr(13)
    strSQL = strSQL & "       ,USAGE_NAME1.NAME_NAME USAGE_NAME1 " & Chr(13)      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "       ,ISNULL(CARG_RENTKG,0) + ISNULL(CARG_SYOZEI,0) KINGAKU1 " & Chr(13)
    strSQL = strSQL & "       ,HOSYICD_NAME1.NAME_NAME HOSYICD_NAME1 " & Chr(13)
    strSQL = strSQL & "       ,CASE WHEN CARG_HOSYICD IN(3, 4, 6) THEN 0 ELSE CARG_SECUKG END CARG_SECUKG " & Chr(13)
    strSQL = strSQL & "       ,KKBN_NAME.NAME_NAME KKBN_NAME " & Chr(13)
    strSQL = strSQL & "       ,SKBN_NAME.NAME_NAME SKBN_NAME " & Chr(13)
    strSQL = strSQL & "       ,MISY.TANTM_TANTN    TANTM_TANTN"
    strSQL = strSQL & "       ,CARG_KEY_RETDATE " & Chr(13)
    strSQL = strSQL & "       ,CONVERT(VARCHAR, yt1.YOUKT_YCODE) YOUKT_YCODE " & Chr(13)
    strSQL = strSQL & "       ,YARD_MAST.YARD_NAME " & Chr(13)
    strSQL = strSQL & "       ,yt1.YOUKT_NO " & Chr(13)
    strSQL = strSQL & "       ,CM2.CNTA_SIZE CNTA_SIZE2 " & Chr(13)
    strSQL = strSQL & "       ,STEP_NAME2.NAME_NAME STEP_NAME2 " & Chr(13)
    strSQL = strSQL & "       ,USAGE_NAME2.NAME_NAME USAGE_NAME2 " & Chr(13)      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "       ,ISNULL(RCPT_RENTKG,0) + ISNULL(RCPT_EZAPPI,0) + ISNULL(RCPT_ADD_EZAPPI1,0) + ISNULL(RCPT_ADD_EZAPPI2,0) KINGAKU2 " & Chr(13)
    strSQL = strSQL & "       ,HOSYICD_NAME2.NAME_NAME HOSYICD_NAME2 " & Chr(13)
    strSQL = strSQL & "       ,MOVETANTO_NAME.NAME_NAME MOVETANTO_NAME " & Chr(13)
    strSQL = strSQL & "       ,CASE WHEN RCPT_HOSYICD IN(3, 4, 6) THEN 0 ELSE RCPT_SECUKG END RCPT_SECUKG " & Chr(13)
    strSQL = strSQL & "       ,CM1.CNTA_CODE " & Chr(13)
    strSQL = strSQL & "       ,CM1.CNTA_NO " & Chr(13)
    strSQL = strSQL & "       ,CM1.CNTA_USE " & Chr(13)
    strSQL = strSQL & "       ,NYAR_KIRO " & Chr(13)
    strSQL = strSQL & "       ,Y1.YARD_NAME AS YNAME " & Chr(13)
    strSQL = strSQL & "       ,Y1.YARD_RENTEND_DAY  " & Chr(13)
    strSQL = strSQL & "FROM " & Chr(13)
    strSQL = strSQL & "       NYAR_MAST  " & Chr(13)
    strSQL = strSQL & "       INNER JOIN CNTA_MAST CM1 ON " & Chr(13)
    strSQL = strSQL & "           CM1.CNTA_CODE = NYAR_NCODE " & Chr(13)
    strSQL = strSQL & "       INNER JOIN YARD_MAST Y1 ON " & Chr(13)
    strSQL = strSQL & "           CM1.CNTA_CODE = Y1.YARD_CODE " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN CARG_FILE ON " & Chr(13)
    strSQL = strSQL & "           CM1.CNTA_CODE = CARG_YCODE AND CM1.CNTA_NO = CARG_NO " & Chr(13)
    strSQL = strSQL & "           AND CARG_AGRE   <> 9 " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN USER_MAST ON " & Chr(13)
    strSQL = strSQL & "           USER_CODE = CARG_UCODE " & Chr(13)
    
    'INSERT 2018-06-21 start
    strSQL = strSQL & "       LEFT OUTER JOIN youk_tran yt2 ON"
    strSQL = strSQL & "            CM1.CNTA_CODE = yt2.youkt_YCODE And CM1.CNTA_NO = yt2.youkt_no"
    strSQL = strSQL & "            and yt2.youkt_yukbn = '2'"
    'INSERT 2018-06-21 end
    
    'OpenRowset SQL文作成
    strOpenRowSetSql = fncMakeOpenRowsetSql()

    'OpenRowset SQL変換
    strSQL = strSQL & "       LEFT OUTER JOIN "
    strSQL = strSQL & fncOpenRowsetString(strOpenRowSetSql) & " MISY "
    strSQL = strSQL & "           ON REPLACE(STR(CARG_UCODE,6),' ','0') = MISYT_KOKYC AND MISYT_BUMOC = '" & strBumonCd & "' "
    
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME1 ON " & Chr(13)
    strSQL = strSQL & "           STEP_NAME1.NAME_ID = '090' AND STEP_NAME1.NAME_CODE = CM1.CNTA_STEP " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST HOSYICD_NAME1 ON " & Chr(13)
    strSQL = strSQL & "           HOSYICD_NAME1.NAME_ID = '200' AND HOSYICD_NAME1.NAME_CODE = CARG_HOSYICD " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST KKBN_NAME ON " & Chr(13)
    strSQL = strSQL & "           KKBN_NAME.NAME_ID = '020' AND KKBN_NAME.NAME_CODE = USER_KKBN " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST SKBN_NAME ON " & Chr(13)
    strSQL = strSQL & "           SKBN_NAME.NAME_ID = '030' AND SKBN_NAME.NAME_CODE = USER_SKBN " & Chr(13)
    
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST USAGE_NAME1 ON " & Chr(13)                                      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "           USAGE_NAME1.NAME_ID = '086' AND USAGE_NAME1.NAME_CODE = CM1.CNTA_USAGE " & Chr(13)    'INSERT 2012/01/24 M.RYU
        
    strSQL = strSQL & "       LEFT OUTER JOIN (YOUK_TRAN yt1 " & Chr(13)
    strSQL = strSQL & "       INNER JOIN YARD_MAST ON " & Chr(13)
    strSQL = strSQL & "           yt1.YOUKT_YCODE = YARD_CODE " & Chr(13)
    strSQL = strSQL & "           AND yt1.YOUKT_YUKBN IN(2,10,20) " & Chr(13) '2007/07/16 add tajima 受付&契約中の予約トランを対象に
    strSQL = strSQL & "       INNER JOIN CNTA_MAST CM2 ON " & Chr(13)
    strSQL = strSQL & "           yt1.YOUKT_YCODE = CM2.CNTA_CODE AND yt1.YOUKT_NO = CM2.CNTA_NO " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN RCPT_TRAN ON " & Chr(13)
    strSQL = strSQL & "           yt1.YOUKT_UKNO  = RCPT_NO " & Chr(13)
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST STEP_NAME2 ON " & Chr(13)
    strSQL = strSQL & "           STEP_NAME2.NAME_ID = '090' AND STEP_NAME2.NAME_CODE = CM2.CNTA_STEP " & Chr(13)
    strSQL = strSQL & "       LEFT  OUTER JOIN NAME_MAST HOSYICD_NAME2 ON " & Chr(13)
    strSQL = strSQL & "           HOSYICD_NAME2.NAME_ID = '200' AND HOSYICD_NAME2.NAME_CODE = RCPT_HOSYICD " & Chr(13)
    strSQL = strSQL & "       LEFT  OUTER JOIN NAME_MAST MOVETANTO_NAME ON " & Chr(13)
    strSQL = strSQL & "           MOVETANTO_NAME.NAME_ID = '084' AND MOVETANTO_NAME.NAME_CODE = yt1.YOUKT_MOVE_TANTO " & Chr(13)
        
    strSQL = strSQL & "       LEFT OUTER JOIN NAME_MAST USAGE_NAME2 ON " & Chr(13)                                      'INSERT 2012/01/24 M.RYU
    strSQL = strSQL & "           USAGE_NAME2.NAME_ID = '086' AND USAGE_NAME2.NAME_CODE = CM2.CNTA_USAGE " & Chr(13)    'INSERT 2012/01/24 M.RYU
    
    strSQL = strSQL & "       ) ON yt1.YOUKT_MOTO_ACPTNO = CARG_ACPTNO " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "      NYAR_YCODE  = '" & strYardCode & "' " & Chr(13)
'    strSQL = strSQL & "  AND CARG_AGRE   <> 9 " & Chr(13)
'    strSQL = strSQL & "   AND ISNULL(CARG_KYDATE,'9999/12/31') > '" & a営業終了日 & "' " & Chr(13)
    
    strSQL = strSQL & "ORDER BY " & Chr(13)
    strSQL = strSQL & "      USER_KANA "

    fncMakeGetDataSqlC = strSQL

    'Call outlog(strSQL)
    'Call Err.Raise(Err.Number, "fncMakeGetDataSqlC" & vbRightAllow & Err.Source, strSQL, Err.HelpFile, Err.HelpContext)

    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncMakeGetDataSqlC" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : ACCESSテーブル存在チェック
'       MODULE_ID       : fncTableExist
'       CREATE_DATE     : 2007/04/09
'       PARAM           : dbAccess     - ACCESSデータベースオブジェクト
'                       : strTableName - テーブル名
'       RETURN          : True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTableExist(dbAccess As Database, strTableName As String) As Boolean

On Error GoTo ErrorHandler

    Dim tdf         As TableDef
    Dim blnError    As Boolean

    blnError = False

    fncTableExist = False
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            fncTableExist = True
            Exit For
        End If
    Next tdf
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not tdf Is Nothing Then Set tdf = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncTableExist" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブル列作成(帳票A)
'       MODULE_ID       : subFieldAppendA
'       CREATE_DATE     : 2007/04/09
'       PARAM           : tdNew            - TableDefオブジェクト
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFieldAppendA(tdfNew As TableDef)

On Error GoTo ErrorHandler

    Dim intCount    As Integer
    Dim blnError    As Boolean
    
    blnError = False

    With tdfNew
    
        Call .Fields.Append(.CreateField("備考", DataTypeEnum.dbText, 20))               '備考
        Call .Fields.Append(.CreateField("１報", DataTypeEnum.dbText, 20))               '１報
        Call .Fields.Append(.CreateField("２報", DataTypeEnum.dbText, 20))               '２報
        Call .Fields.Append(.CreateField("返事１", DataTypeEnum.dbText, 20))             '返事１
        Call .Fields.Append(.CreateField("最終通告", DataTypeEnum.dbText, 20))           '最終通告
        Call .Fields.Append(.CreateField("返事２", DataTypeEnum.dbText, 20))             '返事２
        Call .Fields.Append(.CreateField("契約日", DataTypeEnum.dbText, 10))             '契約日
        Call .Fields.Append(.CreateField("連絡先", DataTypeEnum.dbText, 15))             '連絡先
        Call .Fields.Append(.CreateField("顧客コード", DataTypeEnum.dbText, 6))          '顧客コード
        Call .Fields.Append(.CreateField("顧客名", DataTypeEnum.dbText, 36))             '顧客名
        Call .Fields.Append(.CreateField("コンテナ番号", DataTypeEnum.dbText, 6))        'コンテナ番号
        Call .Fields.Append(.CreateField("実帖", DataTypeEnum.dbText, 8))                '実帖
        Call .Fields.Append(.CreateField("段", DataTypeEnum.dbText, 20))                 '段
        Call .Fields.Append(.CreateField("用途", DataTypeEnum.dbText, 30))               '用途          'INSERT 2012/01/24 M.RYU
        Call .Fields.Append(.CreateField("使用料", DataTypeEnum.dbText, 12))             '使用料
        Call .Fields.Append(.CreateField("保証区分", DataTypeEnum.dbText, 20))           '保証区分
        Call .Fields.Append(.CreateField("保証金", DataTypeEnum.dbText, 10))             '保証金
        Call .Fields.Append(.CreateField("契約", DataTypeEnum.dbText, 20))               '契約
        Call .Fields.Append(.CreateField("支払い", DataTypeEnum.dbText, 20))             '支払い
        
        Call .Fields.Append(.CreateField("未収担当", DataTypeEnum.dbText, 20))           '未収担当
        Call .Fields.Append(.CreateField("サービス", DataTypeEnum.dbText, 20))           'サービス
        Call .Fields.Append(.CreateField("移動", DataTypeEnum.dbText, 20))               '移動
        Call .Fields.Append(.CreateField("移動先ヤード", DataTypeEnum.dbText, 42))       '移動先ヤード
        Call .Fields.Append(.CreateField("移動先コンテナ番号", DataTypeEnum.dbText, 6))  '移動先コンテナ番号
        Call .Fields.Append(.CreateField("移動先実帖", DataTypeEnum.dbText, 8))          '移動先実帖
        Call .Fields.Append(.CreateField("移動先段", DataTypeEnum.dbText, 20))           '移動先段
        Call .Fields.Append(.CreateField("移動先用途", DataTypeEnum.dbText, 30))         '移動先用途    'INSERT 2012/01/24 M.RYU
        Call .Fields.Append(.CreateField("移動先使用料", DataTypeEnum.dbText, 20))       '移動先使用料
        Call .Fields.Append(.CreateField("移動先保証区分", DataTypeEnum.dbText, 20))     '移動先保証区分
        Call .Fields.Append(.CreateField("移動先保証金", DataTypeEnum.dbText, 10))       '移動先保証金
        
        Call .Fields.Append(.CreateField("鍵種類", DataTypeEnum.dbText, 20))             '鍵種類
        Call .Fields.Append(.CreateField("客鍵預", DataTypeEnum.dbText, 20))             '客鍵預
        Call .Fields.Append(.CreateField("客鍵返", DataTypeEnum.dbText, 20))             '客鍵返
        Call .Fields.Append(.CreateField("移動２", DataTypeEnum.dbText, 20))             '移動2
        
        Call .Fields.Append(.CreateField("撮影日", DataTypeEnum.dbText, 20))             '撮影日
        Call .Fields.Append(.CreateField("撮影立会", DataTypeEnum.dbText, 20))           '撮影立会
        Call .Fields.Append(.CreateField("撮影日連絡", DataTypeEnum.dbText, 20))         '撮影日連絡
        Call .Fields.Append(.CreateField("移動日", DataTypeEnum.dbText, 20))             '移動日
        Call .Fields.Append(.CreateField("移動立会", DataTypeEnum.dbText, 20))           '移動立会
        Call .Fields.Append(.CreateField("移動日連絡", DataTypeEnum.dbText, 20))         '移動日連絡
        Call .Fields.Append(.CreateField("新契約発送", DataTypeEnum.dbText, 20))         '新契約発送
        Call .Fields.Append(.CreateField("新鍵発送", DataTypeEnum.dbText, 20))           '新鍵発送
        Call .Fields.Append(.CreateField("新鍵本数", DataTypeEnum.dbText, 20))           '新鍵本数
        Call .Fields.Append(.CreateField("新契約戻", DataTypeEnum.dbText, 20))           '新契約戻
        Call .Fields.Append(.CreateField("旧鍵戻", DataTypeEnum.dbText, 10))             '旧鍵戻
        Call .Fields.Append(.CreateField("備考２", DataTypeEnum.dbText, 20))             '備考
        Call .Fields.Append(.CreateField("完了", DataTypeEnum.dbText, 20))               '完了
        
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subFieldAppendA" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブル列作成(帳票B)
'       MODULE_ID       : subFieldAppendB
'       CREATE_DATE     : 2007/04/09
'       PARAM           : tdNew            - TableDefオブジェクト
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFieldAppendB(tdfNew As TableDef)

On Error GoTo ErrorHandler

    Dim intCount    As Integer
    Dim blnError    As Boolean

    blnError = False

    With tdfNew

        Call .Fields.Append(.CreateField("ヤード", DataTypeEnum.dbText, 42))             'ヤード
        Call .Fields.Append(.CreateField("空コンテナ", DataTypeEnum.dbText, 6))          'コンテナ
        Call .Fields.Append(.CreateField("用途", DataTypeEnum.dbText, 20))               '用途
        Call .Fields.Append(.CreateField("段", DataTypeEnum.dbText, 20))                 '段
        Call .Fields.Append(.CreateField("帖", DataTypeEnum.dbText, 20))                 '帖
        Call .Fields.Append(.CreateField("金額", DataTypeEnum.dbText, 20))               '金額

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subFieldAppendB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブル列作成(帳票C)
'       MODULE_ID       : subFieldAppendC
'       CREATE_DATE     : 2018/02/17
'       PARAM           : tdNew            - TableDefオブジェクト
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFieldAppendC(tdfNew As TableDef)

On Error GoTo ErrorHandler

    Dim intCount    As Integer
    Dim blnError    As Boolean

    blnError = False

    With tdfNew

    
        Call .Fields.Append(.CreateField("備考", DataTypeEnum.dbText, 20))               '備考
        Call .Fields.Append(.CreateField("１報", DataTypeEnum.dbText, 20))               '１報
        Call .Fields.Append(.CreateField("２報", DataTypeEnum.dbText, 20))               '２報
        Call .Fields.Append(.CreateField("返事１", DataTypeEnum.dbText, 20))             '返事１
        Call .Fields.Append(.CreateField("最終通告", DataTypeEnum.dbText, 20))           '最終通告
        Call .Fields.Append(.CreateField("返事２", DataTypeEnum.dbText, 20))             '返事２
        Call .Fields.Append(.CreateField("契約日", DataTypeEnum.dbText, 10))             '契約日
        Call .Fields.Append(.CreateField("連絡先", DataTypeEnum.dbText, 15))             '連絡先
        Call .Fields.Append(.CreateField("顧客コード", DataTypeEnum.dbText, 6))          '顧客コード
        Call .Fields.Append(.CreateField("顧客名", DataTypeEnum.dbText, 36))             '顧客名
        Call .Fields.Append(.CreateField("コンテナ番号", DataTypeEnum.dbText, 6))        'コンテナ番号
        Call .Fields.Append(.CreateField("実帖", DataTypeEnum.dbText, 8))                '実帖
        Call .Fields.Append(.CreateField("段", DataTypeEnum.dbText, 20))                 '段
        Call .Fields.Append(.CreateField("用途", DataTypeEnum.dbText, 30))               '用途          'INSERT 2012/01/24 M.RYU
        Call .Fields.Append(.CreateField("使用料", DataTypeEnum.dbText, 12))             '使用料
        Call .Fields.Append(.CreateField("保証区分", DataTypeEnum.dbText, 20))           '保証区分
        Call .Fields.Append(.CreateField("保証金", DataTypeEnum.dbText, 10))             '保証金
        Call .Fields.Append(.CreateField("契約", DataTypeEnum.dbText, 20))               '契約
        Call .Fields.Append(.CreateField("支払い", DataTypeEnum.dbText, 20))             '支払い
        
        Call .Fields.Append(.CreateField("未収担当", DataTypeEnum.dbText, 20))           '未収担当
        Call .Fields.Append(.CreateField("サービス", DataTypeEnum.dbText, 20))           'サービス
        Call .Fields.Append(.CreateField("移動", DataTypeEnum.dbText, 20))               '移動
        Call .Fields.Append(.CreateField("移動先ヤード", DataTypeEnum.dbText, 42))       '移動先ヤード
        Call .Fields.Append(.CreateField("移動先コンテナ番号", DataTypeEnum.dbText, 6))  '移動先コンテナ番号
        Call .Fields.Append(.CreateField("移動先実帖", DataTypeEnum.dbText, 8))          '移動先実帖
        Call .Fields.Append(.CreateField("移動先段", DataTypeEnum.dbText, 20))           '移動先段
        Call .Fields.Append(.CreateField("移動先用途", DataTypeEnum.dbText, 30))         '移動先用途    'INSERT 2012/01/24 M.RYU
        Call .Fields.Append(.CreateField("移動先使用料", DataTypeEnum.dbText, 20))       '移動先使用料
        Call .Fields.Append(.CreateField("移動先保証区分", DataTypeEnum.dbText, 20))     '移動先保証区分
        Call .Fields.Append(.CreateField("移動先保証金", DataTypeEnum.dbText, 10))       '移動先保証金
        
        Call .Fields.Append(.CreateField("鍵種類", DataTypeEnum.dbText, 20))             '鍵種類
        Call .Fields.Append(.CreateField("客鍵預", DataTypeEnum.dbText, 20))             '客鍵預
        Call .Fields.Append(.CreateField("客鍵返", DataTypeEnum.dbText, 20))             '客鍵返
        Call .Fields.Append(.CreateField("移動２", DataTypeEnum.dbText, 20))             '移動2
        
        Call .Fields.Append(.CreateField("撮影日", DataTypeEnum.dbText, 20))             '撮影日
        Call .Fields.Append(.CreateField("撮影立会", DataTypeEnum.dbText, 20))           '撮影立会
        Call .Fields.Append(.CreateField("撮影日連絡", DataTypeEnum.dbText, 20))         '撮影日連絡
        Call .Fields.Append(.CreateField("移動日", DataTypeEnum.dbText, 20))             '移動日
        Call .Fields.Append(.CreateField("移動立会", DataTypeEnum.dbText, 20))           '移動立会
        Call .Fields.Append(.CreateField("移動日連絡", DataTypeEnum.dbText, 20))         '移動日連絡
        Call .Fields.Append(.CreateField("新契約発送", DataTypeEnum.dbText, 20))         '新契約発送
        Call .Fields.Append(.CreateField("新鍵発送", DataTypeEnum.dbText, 20))           '新鍵発送
        Call .Fields.Append(.CreateField("新鍵本数", DataTypeEnum.dbText, 20))           '新鍵本数
        Call .Fields.Append(.CreateField("新契約戻", DataTypeEnum.dbText, 20))           '新契約戻
        Call .Fields.Append(.CreateField("旧鍵戻", DataTypeEnum.dbText, 10))             '旧鍵戻
        Call .Fields.Append(.CreateField("備考２", DataTypeEnum.dbText, 20))             '備考
        Call .Fields.Append(.CreateField("完了", DataTypeEnum.dbText, 20))               '完了
        
        Call .Fields.Append(.CreateField("ヤードコード", DataTypeEnum.dbText, 6))        'ヤードコード
        Call .Fields.Append(.CreateField("ヤード名", DataTypeEnum.dbText, 60))           'ヤード名
        Call .Fields.Append(.CreateField("距離", DataTypeEnum.dbText, 20))               '距離
        Call .Fields.Append(.CreateField("営業終了日", DataTypeEnum.dbText, 20))         '営業終了日
        Call .Fields.Append(.CreateField("修理", DataTypeEnum.dbText, 20))               '修理

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subFieldAppendB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルデータ追加
'       MODULE_ID       : subMakeWork
'       CREATE_DATE     : 2007/04/09
'       PARAM           : rsSource    - 検索結果が格納されたレコードセット
'                       : intMode     - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
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
    Call psubClearWork(dbAccess, rsSource, P_WORK_TABLE)

    'ワークテーブルのレコードセット
    Set rsDestination = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)

    'データ追加
    While Not rsSource.EOF
        intLoopCount = intLoopCount + 1
        If pstrPrintKbn = P210_帳票A Then
            Call subAddNewA(rsSource, rsDestination)
        ElseIf pstrPrintKbn = P210_帳票B Then
            Call subAddNewB(rsSource, rsDestination)
        'INSERT 2018-02-17 add start
        ElseIf pstrPrintKbn = P210_帳票C Then
            Call subAddNewC(rsSource, rsDestination)
        End If
        'INSERT 2018-02-17 add end
        rsSource.MoveNext
    Wend

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsDestination Is Nothing Then rsDestination.Close: Set rsDestination = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "subMakeWork" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルAddNew(帳票A)
'       MODULE_ID       : subAddNewA
'       CREATE_DATE     : 2007/04/09
'       PARAM           : rsSource      - 検索結果が格納されたレコードセット
'                       : rsDestination - ワークテーブルのレコードセット
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subAddNewA(rsSource As Recordset, rsDestination As Recordset)

On Error GoTo ErrorHandler
    
    Dim blnError    As Boolean
    
    blnError = False

    With rsSource
    
        rsDestination.AddNew
            
        rsDestination.Fields("契約日") = Format$(.Fields("CARG_STDATE"), "YYYY/MM/DD")                       '契約日
        If Nz(.Fields("USER_TEL")) <> "" Then
            rsDestination.Fields("連絡先") = .Fields("USER_TEL")                                             '連絡先(電話番号)
        ElseIf Nz(.Fields("USER_KEITAI")) <> "" Then
            rsDestination.Fields("連絡先") = .Fields("USER_KEITAI")                                          '連絡先(携帯電話)
        ElseIf Nz(.Fields("USER_FAX")) <> "" Then
            rsDestination.Fields("連絡先") = .Fields("USER_FAX")                                             '連絡先(FAX)
        End If
        rsDestination.Fields("顧客コード") = Format$(.Fields("CARG_UCODE"), "000000")                        '顧客コード
        rsDestination.Fields("顧客名") = .Fields("USER_NAME")                                                '顧客名
        rsDestination.Fields("コンテナ番号") = Format$(.Fields("CARG_NO"), "000000")                         'コンテナ番号
        rsDestination.Fields("実帖") = .Fields("CNTA_SIZE1")                                                 '実帖
        rsDestination.Fields("段") = .Fields("STEP_NAME1")                                                   '段
        rsDestination.Fields("用途") = .Fields("USAGE_NAME1")                                                '用途          'INSERT 2012/01/24 M.RYU
        rsDestination.Fields("使用料") = Format(.Fields("KINGAKU1"), "#,###,###,##0")                        '使用料
        rsDestination.Fields("保証区分") = .Fields("HOSYICD_NAME1")                                          '保証区分
        rsDestination.Fields("保証金") = Format(.Fields("CARG_SECUKG"), "#,###,###,##0")                     '保証金
        rsDestination.Fields("契約") = .Fields("KKBN_NAME")                                                  '契約
        rsDestination.Fields("支払い") = .Fields("SKBN_NAME")                                                '支払い
        rsDestination.Fields("未収担当") = .Fields("TANTM_TANTN")                                            '未収担当
        
        If Nz(.Fields("YOUKT_YCODE")) <> "" Then
            rsDestination.Fields("移動先ヤード") = Format$(CStr(.Fields("YOUKT_YCODE")), "000000") + _
                                                    " " + .Fields("YARD_NAME")                               '移動先ヤード
            rsDestination.Fields("移動先コンテナ番号") = Format$(.Fields("YOUKT_NO"), "000000")              '移動先コンテナ番号
            rsDestination.Fields("移動先実帖") = .Fields("CNTA_SIZE2")                                       '移動先実帖
            rsDestination.Fields("移動先段") = .Fields("STEP_NAME2")                                         '移動先段
            rsDestination.Fields("移動先用途") = .Fields("USAGE_NAME2")                                      '移動先用途    'INSERT 2012/01/24 M.RYU
            rsDestination.Fields("移動先使用料") = Format(.Fields("KINGAKU2"), "#,###,###,##0")              '移動先使用料
            rsDestination.Fields("移動先保証区分") = .Fields("HOSYICD_NAME2")                                '移動先保証区分
            rsDestination.Fields("移動先保証金") = Format(.Fields("RCPT_SECUKG"), "#,###,###,##0")           '移動先保証金
            rsDestination.Fields("移動２") = .Fields("MOVETANTO_NAME")                                       '移動２
        End If
        
        If Nz(.Fields("CARG_KEY_RETDATE")) <> "" Then
            rsDestination.Fields("旧鍵戻") = Format$(.Fields("CARG_KEY_RETDATE"), "YYYY/MM/DD")              '旧鍵戻
        End If
                    
        rsDestination.UPDATE
    
    End With
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subAddNewA" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルAddNew(帳票B)
'       MODULE_ID       : subAddNewB
'       CREATE_DATE     : 2007/04/09
'       PARAM           : rsSource      - 検索結果が格納されたレコードセット
'                       : rsDestination - ワークテーブルのレコードセット
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subAddNewB(rsSource As Recordset, rsDestination As Recordset)

On Error GoTo ErrorHandler
    
    Dim blnError    As Boolean
    
    blnError = False

    With rsSource
    
        rsDestination.AddNew
        
        rsDestination.Fields("ヤード") = Format$(.Fields("YARD_CODE"), "000000") + " " + .Fields("YARD_NAME") 'ヤード
        rsDestination.Fields("空コンテナ") = Format$(.Fields("CNTA_NO"), "000000")                            '空コンテナ
        rsDestination.Fields("用途") = .Fields("USAGE_NAME")                                                  '用途
        rsDestination.Fields("段") = .Fields("STEP_NAME")                                                     '段
        rsDestination.Fields("帖") = .Fields("CNTA_SIZE")                                                     '帖
        rsDestination.Fields("金額") = Format(.Fields("PRIC_PRICE"), "#,###,###,##0")                         '金額
                
        rsDestination.UPDATE
    
    End With
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subAddNewB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルAddNew(帳票C)
'       MODULE_ID       : subAddNewC
'       CREATE_DATE     : 2018/02/17
'       PARAM           : rsSource      - 検索結果が格納されたレコードセット
'                       : rsDestination - ワークテーブルのレコードセット
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subAddNewC(rsSource As Recordset, rsDestination As Recordset)

On Error GoTo ErrorHandler
    
    Dim blnError    As Boolean
    
    blnError = False

    With rsSource
    
    
        'Debug.Print Format$(.Fields("CNTA_CODE"), "000000") + " " + Format$(.Fields("CNTA_NO"), "000000")
    
        rsDestination.AddNew
        
        rsDestination.Fields("契約日") = Format$(.Fields("CARG_STDATE"), "YYYY/MM/DD")                       '契約日
        If Nz(.Fields("USER_TEL")) <> "" Then
            rsDestination.Fields("連絡先") = .Fields("USER_TEL")                                             '連絡先(電話番号)
        ElseIf Nz(.Fields("USER_KEITAI")) <> "" Then
            rsDestination.Fields("連絡先") = .Fields("USER_KEITAI")                                          '連絡先(携帯電話)
        ElseIf Nz(.Fields("USER_FAX")) <> "" Then
            rsDestination.Fields("連絡先") = .Fields("USER_FAX")                                             '連絡先(FAX)
        End If
        rsDestination.Fields("顧客コード") = Format$(.Fields("CARG_UCODE"), "000000")                        '顧客コード
        
#If test Then
        rsDestination.Fields("顧客名") = ""                                                                  '顧客名
#Else
        rsDestination.Fields("顧客名") = .Fields("USER_NAME")                                                '顧客名
#End If
        
        'rsDestination.Fields("コンテナ番号") = Format$(.Fields("CARG_NO"), "000000")                         'コンテナ番号
        rsDestination.Fields("コンテナ番号") = Format$(.Fields("CNTA_NO"), "000000")                         'コンテナ番号
        rsDestination.Fields("実帖") = .Fields("CNTA_SIZE1")                                                 '実帖
        rsDestination.Fields("段") = .Fields("STEP_NAME1")                                                   '段
        rsDestination.Fields("用途") = .Fields("USAGE_NAME1")                                                '用途          'INSERT 2012/01/24 M.RYU
        rsDestination.Fields("使用料") = Format(.Fields("KINGAKU1"), "#,###,###,##0")                        '使用料
        rsDestination.Fields("保証区分") = .Fields("HOSYICD_NAME1")                                          '保証区分
        rsDestination.Fields("保証金") = Format(.Fields("CARG_SECUKG"), "#,###,###,##0")                     '保証金
        rsDestination.Fields("契約") = .Fields("KKBN_NAME")                                                  '契約
        rsDestination.Fields("支払い") = .Fields("SKBN_NAME")                                                '支払い
        rsDestination.Fields("未収担当") = .Fields("TANTM_TANTN")                                            '未収担当
        
        If Nz(.Fields("YOUKT_YCODE")) <> "" Then
            rsDestination.Fields("移動先ヤード") = Format$(CStr(.Fields("YOUKT_YCODE")), "000000") + _
                                                    " " + .Fields("YARD_NAME")                               '移動先ヤード
            rsDestination.Fields("移動先コンテナ番号") = Format$(.Fields("YOUKT_NO"), "000000")              '移動先コンテナ番号
            rsDestination.Fields("移動先実帖") = .Fields("CNTA_SIZE2")                                       '移動先実帖
            rsDestination.Fields("移動先段") = .Fields("STEP_NAME2")                                         '移動先段
            rsDestination.Fields("移動先用途") = .Fields("USAGE_NAME2")                                      '移動先用途    'INSERT 2012/01/24 M.RYU
            rsDestination.Fields("移動先使用料") = Format(.Fields("KINGAKU2"), "#,###,###,##0")              '移動先使用料
            rsDestination.Fields("移動先保証区分") = .Fields("HOSYICD_NAME2")                                '移動先保証区分
            rsDestination.Fields("移動先保証金") = Format(.Fields("RCPT_SECUKG"), "#,###,###,##0")           '移動先保証金
            rsDestination.Fields("移動２") = .Fields("MOVETANTO_NAME")                                       '移動２
        End If
        
        If Nz(.Fields("CARG_KEY_RETDATE")) <> "" Then
            rsDestination.Fields("旧鍵戻") = Format$(.Fields("CARG_KEY_RETDATE"), "YYYY/MM/DD")              '旧鍵戻
        End If
        
        rsDestination.Fields("ヤードコード") = Format$(.Fields("CNTA_CODE"), "000000")                       'ヤード
        rsDestination.Fields("ヤード名") = .Fields("YNAME")                                                  'ヤード名
        rsDestination.Fields("距離") = .Fields("NYAR_KIRO")                                                  '距離
        rsDestination.Fields("営業終了日") = .Fields("YARD_RENTEND_DAY")                                     '営業終了日
        rsDestination.Fields("修理") = Nz(.Fields("CNTA_USE"), "")                                           '営業終了日
                
        rsDestination.UPDATE
    
    End With
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subAddNewB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub
'==============================================================================*
'
'       MODULE_NAME     : KOMSデータベース接続
'       MODULE_ID       : subConnectServer
'       CREATE_DATE     :
'       PARAM           : dbSqlServer    -データベースオブジェクト
'                       : strBumonCd     -部門コード
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectServer(ByRef dbSQLServer As Database, ByVal strBumonCd As String)

On Error GoTo ErrorHandler

    Dim strSqlserver    As String
    Dim strConnect      As String
    Dim blnError        As Boolean
    
    blnError = False

    'SQL-Server名
    strSqlserver = fncGetSqlServer(strBumonCd)

    '接続文字列取得
    strConnect = fncGetConnectString(strBumonCd)

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "subConnectServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

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
    Dim blnError        As Boolean
    
    blnError = False

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

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncGetSqlServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
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
    Dim blnError        As Boolean
    
    blnError = False

    strConnectString = MSZZ007_M00(strBumonCode)
    If strConnectString = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "接続文字列の設定不正。")
    End If

    fncGetConnectString = strConnectString

    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncGetConnectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      : エクセル自動書式(帳票A)
'        MODULE_ID        : fncExcelFormatSaveA
'        CREATE_DATE      : 2007/04/09
'        PARAM            : strFileName      - 出力ﾌｧｲﾙﾊﾟｽ&ﾌｧｲﾙ名
'                         : strYardCode      - ヤードコード
'                         : strYardName      - ヤード名
'        RETURN           : 結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncExcelFormatSaveA(ByVal strFilename As String, _
                                        ByVal strYardCode As String, ByVal strYardName As String) As Boolean
    
On Error GoTo ErrorHandler
    
    Dim xlApp       As Object
    Dim xlBook      As Object
    Dim blnError    As Boolean
    Dim strRow      As String
    Dim strTitle    As String
               
    blnError = False
               
    'ﾌｧｲﾙ存在ﾁｪｯｸ
    If Dir(strFilename) <> "" Then
        '存在時は削除
        Kill strFilename
    End If
    
    'Excel出力
    doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, strFilename, False, ""
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(strFilename)
    
    Call setLineStyle(strYardCode, strYardName, xlBook, strRow, strTitle)   'INSERT 2012/01/24 M.RYU
    
    '----↓↓↓----DELETE 2012/01/24 M.RYU------------
'    '帳票A 編集
'    With xlBook.ActiveSheet
'
'        '行挿入
'        .Rows("1:2").Insert Shift:=xlDown
'
'        '見出し編集
'        .Range("A2").FormulaR1C1 = "対応状況"
'        .Range("G2").FormulaR1C1 = "現契約情報"
'        .Range("D3").FormulaR1C1 = "返事"
'        .Range("F3").FormulaR1C1 = "返事"
'
'        '---↓↓----UPDATE 2012/01/24 M.RYU-----------<S>
''        .Range("U2").FormulaR1C1 = "移動先情報"
''        .Range("AM2").FormulaR1C1 = "新(移動先)契約書類の状況"
''        .Range("AG2").FormulaR1C1 = "加瀬移動時の連絡状況"
''        .Range("W3").FormulaR1C1 = "コンテナ番号"
''        .Range("X3").FormulaR1C1 = "実帖"
''        .Range("Y3").FormulaR1C1 = "段"
''        .Range("Z3").FormulaR1C1 = "使用料"
''        .Range("AA3").FormulaR1C1 = "保障区分"
''        .Range("AB3").FormulaR1C1 = "保証金"
''        .Range("AF3").FormulaR1C1 = "移動"
'
'        .Range("V2").FormulaR1C1 = "移動先情報"
'        .Range("AI2").FormulaR1C1 = "加瀬移動時の連絡状況"
'        .Range("AO2").FormulaR1C1 = "新(移動先)契約書類の状況"
'        .Range("X3").FormulaR1C1 = "コンテナ番号"
'        .Range("Y3").FormulaR1C1 = "実帖"
'        .Range("Z3").FormulaR1C1 = "段"
'        .Range("AA3").FormulaR1C1 = "用途"
'        .Range("AB3").FormulaR1C1 = "使用料"
'        .Range("AC3").FormulaR1C1 = "保障区分"
'        .Range("AD3").FormulaR1C1 = "保証金"
'        .Range("AH3").FormulaR1C1 = "移動"
'        .Range("AU2").FormulaR1C1 = "備考"
'        '---↑↑----UPDATE 2012/01/24 M.RYU-----------<E>
'
'        '[対応状況]編集（罫線、色）
'        .Range("A2:F2").Borders(xlDiagonalDown).LineStyle = xlNone
'        .Range("A2:F2").Borders(xlDiagonalUp).LineStyle = xlNone
'        With .Range("A2:F2").Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:F2").Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:F2").Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:F2").Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        .Range("A2:F2").Borders(xlInsideVertical).LineStyle = xlNone
'        With .Range("A2:F2").Interior
'            .ColorIndex = 15
'            .Pattern = xlSolid
'        End With
'        .Range("A2:F2").Merge
'
'        '[現契約情報]編集（罫線、色）
'        .Range("G2:T2").Borders(xlDiagonalDown).LineStyle = xlNone
'        .Range("G2:T2").Borders(xlDiagonalUp).LineStyle = xlNone
'        With .Range("G2:T2").Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("G2:T2").Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("G2:T2").Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("G2:T2").Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        .Range("G2:T2").Borders(xlInsideVertical).LineStyle = xlNone
'        With .Range("G2:T2").Interior
'            .ColorIndex = 15
'            .Pattern = xlSolid
'        End With
'        .Range("G2:T2").Merge
'
'        '[移動先情報]編集（罫線、色）
'        .Range("U2:AF2").Borders(xlDiagonalDown).LineStyle = xlNone
'        .Range("U2:AF2").Borders(xlDiagonalUp).LineStyle = xlNone
'        With .Range("U2:AF2").Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("U2:AF2").Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("U2:AF2").Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("U2:AF2").Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        .Range("U2:AF2").Borders(xlInsideVertical).LineStyle = xlNone
'        With .Range("U2:AF2").Interior
'            .ColorIndex = 15
'            .Pattern = xlSolid
'        End With
'        .Range("U2:AF2").Merge
'
'        '[加瀬移動時の連絡状況]編集（罫線、色）
'        .Range("AG2:AL2").Borders(xlDiagonalDown).LineStyle = xlNone
'        .Range("AF2:AL2").Borders(xlDiagonalUp).LineStyle = xlNone
'        With .Range("AG2:AL2").Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AG2:AL2").Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AG2:AL2").Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AG2:AL2").Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        .Range("AG2:AL2").Borders(xlInsideVertical).LineStyle = xlNone
'        With .Range("AG2:AL2").Interior
'            .ColorIndex = 15
'            .Pattern = xlSolid
'        End With
'        .Range("AG2:AL2").Merge
'
'        '[ ]編集（罫線、色）
'        .Range("AM2:AQ2").Borders(xlDiagonalDown).LineStyle = xlNone
'        .Range("AM2:AQ2").Borders(xlDiagonalUp).LineStyle = xlNone
'        With .Range("AM2:AQ2").Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AM2:AQ2").Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AM2:AQ2").Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AM2:AQ2").Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        .Range("AM2:AQ2").Borders(xlInsideVertical).LineStyle = xlNone
'        With .Range("AM2:AQ2").Interior
'            .ColorIndex = 15
'            .Pattern = xlSolid
'        End With
'        .Range("AM2:AQ2").Merge
'
'        '[備考]編集（罫線、色）
'        With .Range("AR2:AR3")
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        .Range("AR2:AR3").Merge
'        .Range("AR2:AR3").Borders(xlDiagonalDown).LineStyle = xlNone
'        .Range("AR2:AR3").Borders(xlDiagonalUp).LineStyle = xlNone
'        With .Range("AR2:AR3").Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AR2:AR3").Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AR2:AR3").Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AR2:AR3").Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        .Range("AR2:AR3").Borders(xlInsideHorizontal).LineStyle = xlNone
'
'        '[完了]編集（罫線、色）
'        With .Range("AS2:AS3")
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        .Range("AS2:AS3").Merge
'        .Range("AS2:AS3").Borders(xlDiagonalDown).LineStyle = xlNone
'        .Range("AS2:AS3").Borders(xlDiagonalUp).LineStyle = xlNone
'        With .Range("AS2:AS3").Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AS2:AS3").Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AS2:AS3").Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("AS2:AS3").Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        .Range("AS2:AS3").Borders(xlInsideHorizontal).LineStyle = xlNone
'        .Range("AR2").FormulaR1C1 = "備考"
'
'        'データ部編集(罫線)
'        strRow = .Range("G3").End(xlDown).Row
'        With .Range("A2:AS" + strRow).Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:AS" + strRow).Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:AS" + strRow).Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:AS" + strRow).Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:AS" + strRow).Borders(xlInsideVertical)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'        With .Range("A2:AS" + strRow).Borders(xlInsideHorizontal)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'        End With
'
'        '数値項目右詰め
'        .Range("L4:L" + strRow).HorizontalAlignment = xlRight
'        .Range("P4:P" + strRow).HorizontalAlignment = xlRight
'        .Range("N4:N" + strRow).HorizontalAlignment = xlRight
'        .Range("X4:X" + strRow).HorizontalAlignment = xlRight
'        .Range("Z4:Z" + strRow).HorizontalAlignment = xlRight
'        .Range("AB4:AB" + strRow).HorizontalAlignment = xlRight
'
'
'        '日付項目書式設定
'        .Range("B4:F" + strRow).NumberFormatLocal = "m/d;@"
'        .Range("AG4:AG" + strRow).NumberFormatLocal = "m/d;@"
'        .Range("AI4:AJ" + strRow).NumberFormatLocal = "m/d;@"
'        .Range("AL4:AM" + strRow).NumberFormatLocal = "m/d;@"
'        .Range("AP4:AQ" + strRow).NumberFormatLocal = "m/d;@"
'
'        .Cells.Select
'        .Cells.EntireColumn.AutoFit
'        .Range("L4").Select
'
'        strTitle = strYardCode & " " & strYardName & "　　作成日時：" & Format(Now, "yyyy/mm/dd hh:mm")
'        .Range("A1").FormulaR1C1 = strTitle
'
        xlApp.ActiveWindow.FreezePanes = True
'
'    End With

    xlBook.Save
    xlApp.WindowState = -4137
    xlApp.Visible = True
    fncExcelFormatSaveA = True
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not xlBook Is Nothing Then Set xlBook = Nothing
    If Not xlApp Is Nothing Then Set xlApp = Nothing
    If blnError Then
        fncExcelFormatSaveA = False
        Call Err.Raise(Err.Number, "fncExcelFormatSaveA" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
        
End Function

'==============================================================================*
'
'        MODULE_NAME      : エクセル自動書式(帳票B)
'        MODULE_ID        : fncExcelFormatSaveB
'        CREATE_DATE      : 2007/04/09
'        PARAM            : strFileName      - 出力ﾌｧｲﾙﾊﾟｽ&ﾌｧｲﾙ名
'                         : strYardCode      - ヤードコード
'                         : strYardName      - ヤード名
'                         : rsRecord         - 解約ヤード契約一覧情報
'        RETURN           : 結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncExcelFormatSaveB(ByVal strFilename As String, ByVal strYardCode As String, _
                                     ByVal strYardName As String, ByRef rsRecord As Recordset) As Boolean
    
On Error GoTo ErrorHandler
    
    Dim xlApp       As Object
    Dim xlBook      As Object
    Dim blnError    As Boolean
    Dim strTemp     As String
    Dim strBikou    As String
    Dim strRow      As String
    Dim strCol      As String
    Dim intCol      As Integer
    Dim intIx       As Integer
               
    blnError = False
               
    'ﾌｧｲﾙ存在ﾁｪｯｸ
    If Dir(strFilename) <> "" Then
        '存在時は削除
        Kill strFilename
    End If
    
    'Excel出力
    doCmd.OutputTo acOutputTable, P_WORK_TABLE, acFormatXLS, strFilename, False, ""
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(strFilename)

    '帳票B 編集
    With xlBook.ActiveSheet

        '行挿入
        .Rows("1:7").insert Shift:=xlDown

        '見出し編集
        .Range("F1").FormulaR1C1 = "契約コンテナ"
        .Range("F2").FormulaR1C1 = "用途"
        .Range("F3").FormulaR1C1 = "顧客コード"
        .Range("F4").FormulaR1C1 = "備考"
        .Range("F5").FormulaR1C1 = "顧客名"
        .Range("F6").FormulaR1C1 = "段"
        .Range("F7").FormulaR1C1 = "帖"
        .Range("F8").FormulaR1C1 = "賃料＼使用料"
        
        .Range("A1").FormulaR1C1 = Format(DATE, "YYYY/MM/DD") + "現在"
        .Range("A2").FormulaR1C1 = strYardCode + " " + strYardName + " の代替候補"
        .Range("A2").Font.Bold = True
    
        '解約ヤードの契約一覧作成
        intIx = 7                                                                      '初期列
        While Not rsRecord.EOF
            strCol = fncConvertToLetter(intIx)
            .Range(strCol + "1").NumberFormatLocal = "@"                                           '書式設定(文字列)
            .Range(strCol + "1").FormulaR1C1 = Format$(CStr(rsRecord("CARG_NO")), "000000")        '契約コンテナ
            .Range(strCol + "2").FormulaR1C1 = rsRecord("USAGE_NAME")                              '用途
            .Range(strCol + "3").NumberFormatLocal = "@"                                           '書式設定(文字列)
            .Range(strCol + "3").FormulaR1C1 = Format$(rsRecord("USER_CODE"), "000000")            '顧客コード
            strBikou = ""
            If Nz(rsRecord("YOUKT_YUKBN")) <> "" Then
                Select Case rsRecord("YOUKT_YUKBN")
                    Case 2
                        strBikou = "取置中" & Chr(13) & "" & Chr(10)
                    Case 10
                        strBikou = "受付中" & Chr(13) & "" & Chr(10)
                    Case 20
                        strBikou = "受付完" & Chr(13) & "" & Chr(10)
                End Select
                strBikou = strBikou & Format$(rsRecord("YOUKT_YCODE"), "000000") & "-" & Format$(rsRecord("YOUKT_NO"), "000000")
            End If
            .Range(strCol + "4").WrapText = True
            .Range(strCol + "4").ColumnWidth = 20
            .Range(strCol + "4").FormulaR1C1 = strBikou                                            '備考
            .Range(strCol + "5").FormulaR1C1 = rsRecord("USER_NAME")                               '顧客名
            .Range(strCol + "6").FormulaR1C1 = rsRecord("STEP_NAME")                               '段
            .Range(strCol + "7").FormulaR1C1 = Format(rsRecord("CNTA_SIZE"), "###.00")             '帖
            .Range(strCol + "8").FormulaR1C1 = Format(rsRecord("KINGAKU"), "#,###,###,##0")        '金額
            intIx = intIx + 1
            rsRecord.MoveNext
        Wend

        '罫線描画
        intCol = .Range("F1").End(xlToRight).Column
        strCol = .Range("F1").End(xlToRight).ADDRESS
        strCol = Mid(strCol, InStr(strCol, "$") + 1, InStr(2, strCol, "$") - 2)

        With .Range("F1:" + strCol + "8").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("F1:" + strCol + "8").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("F1:" + strCol + "8").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("F1:" + strCol + "8").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("F1:" + strCol + "8").Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("F1:" + strCol + "8").Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        With .Range("F1:F8").Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With

        strRow = .Range("A8").End(xlDown).row
        With .Range("A8:" + strCol + strRow).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A8:" + strCol + strRow).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A8:" + strCol + strRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A8:" + strCol + strRow).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A8:" + strCol + strRow).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A8:" + strCol + strRow).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        '金額チェック
        For intIx = 9 To CInt(strRow) Step 1
            If fncCheckKingaku(xlBook.ActiveSheet, intIx, intCol) Then
                .Range("F" + CStr(intIx)).Font.ColorIndex = 3
            End If
        Next

        '数値項目右詰め
        .Range("E9:E" + strRow).HorizontalAlignment = xlRight
        .Range("F9:F" + strRow).HorizontalAlignment = xlRight
        
        'ウィンドウ枠固定
        .Cells.Select
        .Cells.Font.SIZE = 9
        .Cells.EntireColumn.AutoFit
        .Range("G9").Select
        
        .Range("A4").FormulaR1C1 = "◎：段が同じで、帖が同じか広い、かつ、金額が同じか安い"
        .Range("A5").FormulaR1C1 = "○：帖・段が同じで、金額が高い"
        .Range("A6").FormulaR1C1 = "△：帖が同じ"
        
        xlApp.ActiveWindow.FreezePanes = True
        .Range("A8:F8").AutoFilter

    End With

    xlBook.Save
    xlApp.WindowState = -4137
    xlApp.Visible = True
    fncExcelFormatSaveB = True
    
    
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not xlBook Is Nothing Then Set xlBook = Nothing
    If Not xlApp Is Nothing Then Set xlApp = Nothing
    If blnError Then
        fncExcelFormatSaveB = False
        Call Err.Raise(Err.Number, "fncExcelFormatSaveB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
        
End Function


'==============================================================================*
'
'        MODULE_NAME      : エクセル自動書式(帳票C)
'        MODULE_ID        : fncExcelFormatSaveC
'        CREATE_DATE      : 2018/02/17
'        PARAM            : strFileName      - 出力ﾌｧｲﾙﾊﾟｽ&ﾌｧｲﾙ名
'                         : strYardCode      - ヤードコード
'                         : strYardName      - ヤード名
'        RETURN           : 結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncExcelFormatSaveC(ByVal strFilename As String, _
                                        ByVal strYardCode As String, ByVal strYardName As String) As Boolean
    
On Error GoTo ErrorHandler
    
    Dim xlApp       As Object
    Dim xlBook      As Object
    Dim blnError    As Boolean
    Dim strRow      As String
    Dim strTitle    As String
               
    Dim dbAccess        As Database
    Dim rs              As Recordset
    Dim i               As Integer
    Dim row             As Integer
    Dim col             As Integer
    Dim intLoopCount    As Integer
    Dim strSQL          As String
    Dim wk              As String
    Dim tmp_yard_index  As Integer
    Dim tmp_size_index  As Integer
    
    Dim rc_title As Type_row_col
    Dim rc_date As Type_row_col
    Dim rc_username As Type_row_col
    Dim rc_kinrin_ycode As Type_row_col
    Dim rc_kinrin_yname As Type_row_col
    Dim rc_kinrin_kiro As Type_row_col
    Dim rc_kinrin_step As Type_row_col
    Dim rc_kinrin_size As Type_row_col
    Dim rc_kinrin_end As Type_row_col
    Dim rc_kinrin_data As Type_row_col
    Dim rc_kinrin_sum As Type_row_col
    Dim ilen            As Integer
   
    Dim xlsfilepath As String
    Dim xlsfilename As String
        
    blnError = False
    intLoopCount = 0

    Call subMakeQueryWorkC(strYardCode, 0)
    
    Set dbAccess = CurrentDb

    rc_title.row = 2
    rc_title.col = 12
    rc_date.row = 2
    rc_date.col = 27
    rc_username.row = 2
    rc_username.col = 30

    rc_kinrin_ycode.row = 4
    rc_kinrin_ycode.col = 6
    rc_kinrin_yname.row = 5
    rc_kinrin_yname.col = 6
    rc_kinrin_end.row = 6
    rc_kinrin_end.col = 6
    rc_kinrin_kiro.row = 7
    rc_kinrin_kiro.col = 6
    rc_kinrin_step.row = 8
    rc_kinrin_step.col = 2
    rc_kinrin_size.row = 8
    rc_kinrin_size.col = 3
    rc_kinrin_data.row = 8
    rc_kinrin_data.col = 6
    rc_kinrin_sum.row = 8
    rc_kinrin_sum.col = 60
     
    blnError = False

    xlsfilepath = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & P_REPORT & "' AND INTIF_RECFB='" & par_FVS210_INTIF_RECFB_1 & "'")
    xlsfilename = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & P_REPORT & "' AND INTIF_RECFB='" & par_FVS210_INTIF_RECFB_2 & "'")
    
    'ﾌｧｲﾙ存在ﾁｪｯｸ
    If Dir(strFilename) <> "" Then
        '存在時は削除
        Kill strFilename
    End If
    
    FileCopy xlsfilepath & "\" & xlsfilename, strFilename
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(strFilename)
    
    With xlBook.ActiveSheet
    
        .Rows(1).Select
        .Rows(1).ClearContents
        .Rows(1).Select
        .Rows(2).ClearContents
        .Rows(4).Select
        .Rows(4).ClearContents

        .Cells(rc_title.row, rc_title.col).VALUE = strYardName & " 振分表"
        .Cells(rc_date.row, rc_date.col).VALUE = Format(DATE, "yyyy/mm/dd") & "作成"
       
        intLoopCount = 0
        ilen = 0
        For col = 0 To UBound(pNYARLists)
        
            'Debug.Print .Cells(rc_kinrin_yname.row, (col * 3) + rc_kinrin_ycode.col).VALUE
        
            .Cells(rc_kinrin_ycode.row, (col * 3) + rc_kinrin_ycode.col).VALUE = pNYARLists(intLoopCount).YCODE
            .Cells(rc_kinrin_yname.row, (col * 3) + rc_kinrin_ycode.col).VALUE = pNYARLists(intLoopCount).YNAME '& "ああああああああ"
            .Cells(rc_kinrin_kiro.row, (col * 3) + rc_kinrin_ycode.col).VALUE = pNYARLists(intLoopCount).KIRO
            .Cells(rc_kinrin_end.row, (col * 3) + rc_kinrin_end.col).VALUE = pNYARLists(intLoopCount).END_DATE
            
            'Debug.Print intLoopCount & " " & pNYARLists(intLoopCount).YNAME
        
            If (ilen < LenB(.Cells(rc_kinrin_yname.row, (col * 3) + rc_kinrin_ycode.col).VALUE)) Then
                ilen = LenB(.Cells(rc_kinrin_yname.row, (col * 3) + rc_kinrin_ycode.col).VALUE)
            End If
            
            intLoopCount = intLoopCount + 1
        Next
    
        '高さ調整（９バイトまではデフォルト１８ポイント）
        .Rows(rc_kinrin_yname.row).RowHeight = (ilen / 9) * 18
    
        intLoopCount = 0
        For row = rc_kinrin_size.row To UBound(pSIZELists) + rc_kinrin_size.row
            .Cells(row, rc_kinrin_size.col).VALUE = pSIZELists(intLoopCount).SIZE
            .Cells(row, rc_kinrin_sum.col).Formula = "=SUM(I" & row & ":BG" & row & ")-F" & row
            intLoopCount = intLoopCount + 1
        Next
    
        .Cells(rc_kinrin_step.row, rc_kinrin_step.col).VALUE = "上"
        .Cells(rc_kinrin_step.row + pSIZEListsUpRows, rc_kinrin_step.col).VALUE = "下"
    
   
        'ユーザー数取得
        strSQL = ""
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " ヤードコード "
        strSQL = strSQL & " ,段 "
        strSQL = strSQL & " ,実帖 "
        strSQL = strSQL & " ,count(*) as ユーザー数 "
        strSQL = strSQL & " FROM " & P_WORK_TABLE
        strSQL = strSQL & " WHERE 1=1"
        strSQL = strSQL & " AND ヤードコード = '" & strYardCode & "' "
        strSQL = strSQL & " AND 顧客コード <> """" "
        strSQL = strSQL & " GROUP BY "
        strSQL = strSQL & " ヤードコード "
        strSQL = strSQL & " ,段 "
        strSQL = strSQL & " ,実帖 "
        Set rs = dbAccess.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)
           
        'データセット
        intLoopCount = 0
        Do While Not rs.EOF
                
            tmp_yard_index = 0
            tmp_size_index = 0
            
            For i = 0 To UBound(pSIZELists)
                If pSIZELists(i).STEP = rs.Fields("段") And pSIZELists(i).SIZE = rs.Fields("実帖") Then
                    tmp_size_index = i
                    Exit For
                End If
            Next
            
            'ユーザー数
            .Cells(rc_kinrin_data.row + tmp_size_index, rc_kinrin_data.col).VALUE = rs.Fields("ユーザー数")
                
            intLoopCount = intLoopCount + 1
            
            rs.MoveNext
        Loop
                   
        '空き数取得
        strSQL = ""
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " ヤードコード "
        strSQL = strSQL & " ,段 "
        strSQL = strSQL & " ,実帖 "
        strSQL = strSQL & " ,count(*) as 空き数 "
        strSQL = strSQL & " FROM " & P_WORK_TABLE
        strSQL = strSQL & " WHERE 1=1"
        strSQL = strSQL & " AND ヤードコード in (" & pNYARListsJoin & ") "
        strSQL = strSQL & " AND 顧客コード = """" "
        strSQL = strSQL & " AND 修理 <> ""0"" "
        strSQL = strSQL & " AND 修理 <> ""9"" "
        strSQL = strSQL & " GROUP BY "
        strSQL = strSQL & " ヤードコード "
        strSQL = strSQL & " ,段 "
        strSQL = strSQL & " ,実帖 "
        Set rs = dbAccess.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)
                   
        intLoopCount = 0
        Do While Not rs.EOF
            'pNYARLists(intLoopCount).KIRO = Nz(rs.Fields("距離"), "")
                
            tmp_yard_index = 0
            tmp_size_index = 0
            
            For i = 0 To UBound(pNYARLists)
                If pNYARLists(i).YCODE = rs.Fields("ヤードコード") Then
                    tmp_yard_index = i
                    Exit For
                End If
            Next
            For i = 0 To UBound(pSIZELists)
                If pSIZELists(i).STEP = rs.Fields("段") And pSIZELists(i).SIZE = rs.Fields("実帖") Then
                    tmp_size_index = i
                    Exit For
                End If
            Next
            
            '空き数
            .Cells(rc_kinrin_data.row + tmp_size_index, rc_kinrin_data.col + (tmp_yard_index * 3)).VALUE = rs.Fields("空き数")
            
            'Debug.Print tmp_size_index & " " & tmp_yard_index & " " & rc_kinrin_data.row + tmp_size_index & " " & rc_kinrin_data.col + (tmp_yard_index * 3)
            
                
            intLoopCount = intLoopCount + 1
            
            rs.MoveNext
        Loop
    
        .Rows(rc_kinrin_data.row).Select
        .Range("A" & rc_kinrin_data.row).Activate
        .Rows(rc_kinrin_data.row).Copy
        .Rows(rc_kinrin_data.row + 1 & ":" & rc_kinrin_data.row + UBound(pSIZELists)).Select
        .Range("A" & rc_kinrin_data.row).Activate
        .Rows(rc_kinrin_data.row + 1 & ":" & rc_kinrin_data.row + UBound(pSIZELists)).PasteSpecial Paste:=&HFFFFEFE6, Operation:=&HFFFFEFD2, SkipBlanks:=False, Transpose:=False
        
        .Rows(rc_kinrin_data.row).Select
        .Rows(rc_kinrin_data.row).Copy
        .Rows(rc_kinrin_data.row + 1 & ":" & rc_kinrin_data.row + UBound(pSIZELists)).Select
        .Range("A" & rc_kinrin_data.row).Activate
        .Rows(rc_kinrin_data.row + 1 & ":" & rc_kinrin_data.row + UBound(pSIZELists)).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Application.CutCopyMode = False
        
        With .Range("B" & rc_kinrin_data.row & ":B" & rc_kinrin_data.row + pSIZEListsUpRows - 1)
            .Select
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
            
        With .Range("C" & rc_kinrin_data.row & ":BJ" & rc_kinrin_data.row + pSIZEListsUpRows - 1)
            '中罫線を点線
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            
        End With
        
        With .Range("B" & rc_kinrin_data.row + pSIZEListsUpRows & ":B" & rc_kinrin_data.row + pSIZEListsUpRows + pSIZEListsDownRows - 1)
            .Select
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        
        End With
        
        With .Range("C" & rc_kinrin_data.row + pSIZEListsUpRows & ":BJ" & rc_kinrin_data.row + pSIZEListsUpRows + pSIZEListsDownRows - 1)
            
            '中罫線を点線
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlHairline
            End With
            
        End With
        
        .Range("A1").Activate
        
    End With
    
    '行固定
    'xlApp.ActiveWindow.FreezePanes = True

    xlBook.Save
    xlApp.WindowState = -4137
    xlApp.Visible = True
    fncExcelFormatSaveC = True
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

    If Not xlBook Is Nothing Then
        xlBook.Close (False)
     End If

ExitRtn:
    If Not xlBook Is Nothing Then
        Set xlBook = Nothing
     End If
    If Not xlApp Is Nothing Then Set xlApp = Nothing
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If blnError Then
        fncExcelFormatSaveC = False
        Call Err.Raise(Err.Number, "fncExcelFormatSaveC" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
        
End Function

'==============================================================================*
'
'        MODULE_NAME      : 金額チェック
'        MODULE_ID        : fncCheckKingaku
'        PARAMETER        : xlSheet                         - ワークシートオブジェクト
'                         : intCulIx                        - 比較元行番号
'                         : intCol                          - 最終列番号
'        CREATE_DATE      : 2007/04/09
'        UPDATE_DATE      :
'        NOTE             : 戻り値としてチェック結果(True/False)を返す。
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckKingaku(xlSheet As Object, ByVal intCurIx As Integer, ByVal intCol As Integer) As Boolean

On Error GoTo ErrorHandler

    Dim blnError    As Boolean
    Dim dbl帖       As Double               '比較元-帖
    Dim str段       As String               '比較元-段
    Dim lng金額     As Long                 '比較元-金額
    Dim intIx       As Integer
    Dim strCol      As String
    
    fncCheckKingaku = False
    
    With xlSheet
        '比較元情報取得
        dbl帖 = CDbl(.Range("E" + CStr(intCurIx)))
        str段 = .Range("D" + CStr(intCurIx))
        lng金額 = CLng(.Range("F" + CStr(intCurIx)))
        
        '列数分比較を行う
        For intIx = 7 To intCol Step 1
            '列番号をｱﾙﾌｧﾍﾞｯﾄ文字へ変換
            strCol = fncConvertToLetter(intIx)
            
            If dbl帖 = CDbl(.Range(strCol + "7")) Then
                '帖が同じ場合は「△」
                .Range(strCol + CStr(intCurIx)).FormulaR1C1 = "△"
'''                With .Range(strCol + CStr(intCurIx)).Interior
'''                    .ColorIndex = 8
'''                    .Pattern = xlSolid
'''                End With
            
                If str段 = .Range(strCol + "6") Then
                    If lng金額 > CLng(.Range(strCol + "8")) Then
                        '帖・段が同じで金額が高い場合は「○」
                        .Range(strCol + CStr(intCurIx)).FormulaR1C1 = "○"
'''                        With .Range(strCol + CStr(intCurIx)).Interior
'''                            .ColorIndex = 4
'''                            .Pattern = xlSolid
'''                        End With
                        '同じ帖・段で金額が高い場合は赤字
                        fncCheckKingaku = True
                    End If
                End If
            End If
            
            If dbl帖 >= CDbl(.Range(strCol + "7")) Then
                If str段 = .Range(strCol + "6") Then
                    If lng金額 <= CLng(.Range(strCol + "8")) Then
                        '段が同じで、帖が同じか広い、且つ、金額が同じか安い場合は「◎」
                        .Range(strCol + CStr(intCurIx)).FormulaR1C1 = "◎"
'''                        With .Range(strCol + CStr(intCurIx)).Interior
'''                            .ColorIndex = 6
'''                            .Pattern = xlSolid
'''                        End With
                    End If
                End If
            End If
        Next
    End With
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        fncCheckKingaku = False
        Call Err.Raise(Err.Number, "fncCheckKingaku" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
        
End Function

'==============================================================================*
'
'        MODULE_NAME      : 列番号変換
'        MODULE_ID        : fncConvertToLetter
'        PARAMETER        : intCol             -列番号
'        CREATE_DATE      : 2007/04/09
'        UPDATE_DATE      :
'        NOTE             : 戻り値として列値(ｱﾙﾌｧﾍﾞｯﾄ)を返す。
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Function fncConvertToLetter(intCol As Integer) As String

On Error GoTo ErrorHandler

    Dim blnError    As Boolean
    Dim intAlpha As Integer
    Dim intRemainder As Integer
   
    ' intAlpha = Int(intCol / 27)      2007/07/16 修正前
    intAlpha = Int((intCol - 1) / 26) '2007/07/16 修正後

    intRemainder = intCol - (intAlpha * 26)
    If intAlpha > 0 Then
      fncConvertToLetter = Chr(intAlpha + 64)
    End If
    If intRemainder > 0 Then
       fncConvertToLetter = fncConvertToLetter & Chr(intRemainder + 64)
    End If
   
    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        fncConvertToLetter = False
        Call Err.Raise(Err.Number, "fncConvertToLetter" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :ﾌｧｲﾙ保存ﾀﾞｲｱﾛｸﾞ表示
'        MODULE_ID        :fncGetFileName
'        PARAMETER        :strYardCode   - ヤードコード
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'        NOTE             :戻り値としてファイルパスとファイル名を返す。
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetFileName(ByVal strYardCode As String) As String

On Error GoTo ErrorHandler

    Dim returnValue As Integer
    Dim strFilePath As String           '初期表示ﾌｫﾙﾀﾞ
    Dim strTitle    As String           'ﾀﾞｲｱﾛｸﾞﾀｲﾄﾙ
    Dim strFilter   As String           'ﾌｨﾙﾀ(ﾌｧｲﾙ種別)
    Dim blnError    As Boolean
    
    blnError = False
        
    strFilter = "Excelﾌｧｲﾙ (*.xls)|*.xls"
    strFilePath = strYardCode & "-" & Format$(Now, "yyyymmddhhmm") & ".xls"
    
    If pstrPrintKbn = P210_帳票C Then
        strFilePath = "ヤード振分け表-" & strFilePath
    End If
    
    WizHook.key = 51488399 'WIZHOOK有効
    returnValue = WizHook.getFileName( _
                    0, "", strTitle, "", strFilePath, "", _
                    strFilter, _
                    0, 0, &H1, 0 _
                    )
    WizHook.key = 0 ' WizHook 無効
    
    If returnValue = 0 Then
        fncGetFileName = strFilePath
    Else
        fncGetFileName = ""
    End If
   
    GoTo ExitRtn

ErrorHandler:
    blnError = True
ExitRtn:
    If blnError Then
        Call Err.Raise(Err.Number, "fncGetFileName" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : OpenRowset SQL文作成
'       MODULE_ID       : fncMakeOpenRowsetSql
'       CREATE_DATE     : 2007/04/25
'       PARAM           :
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeOpenRowsetSql() As String

    Dim strSQL              As String

    strSQL = strSQL & "SELECT TANTM_TANTN, "
    strSQL = strSQL & "MISYT_BUMOC, "
    strSQL = strSQL & "MISYT_KOKYC  "

    '【未収トラン】
    strSQL = strSQL & "FROM MISY_TRAN "

    '【担当者マスタ】
    strSQL = strSQL & "INNER JOIN TANT_MAST "
    strSQL = strSQL & "ON MISYT_MISTC = TANTM_TANTC "
    strSQL = strSQL & "AND TANTM_BUMOC = 'L' "
    
    fncMakeOpenRowsetSql = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : OpenRowset SQL変換
'       MODULE_ID       : fncOpenRowsetString
'       CREATE_DATE     : 2007/04/25
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
'       MODULE_NAME     : EXCEL線
'       MODULE_ID       : setLineStyle
'       CREATE_DATE     : 2012/01/24 M.RYU
'       PARAM           :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setLineStyle(ByVal strYardCode As String, ByVal strYardName As String, _
                ByRef xlBook As Object, ByRef strRow As String, ByRef strTitle As String)
   '帳票A 編集
    With xlBook.ActiveSheet
        
        '行挿入
        .Rows("1:2").insert Shift:=xlDown
        
        '見出し編集
        .Range("A2").FormulaR1C1 = "対応状況"
        .Range("G2").FormulaR1C1 = "現契約情報"
        .Range("D3").FormulaR1C1 = "返事"
        .Range("F3").FormulaR1C1 = "返事"
                        
        .Range("V2").FormulaR1C1 = "移動先情報"
        .Range("AI2").FormulaR1C1 = "加瀬移動時の連絡状況"
        .Range("AO2").FormulaR1C1 = "新(移動先)契約書類の状況"
        .Range("X3").FormulaR1C1 = "コンテナ番号"
        .Range("Y3").FormulaR1C1 = "実帖"
        .Range("Z3").FormulaR1C1 = "段"
        .Range("AA3").FormulaR1C1 = "用途"
        .Range("AB3").FormulaR1C1 = "使用料"
        .Range("AC3").FormulaR1C1 = "保障区分"
        .Range("AD3").FormulaR1C1 = "保証金"
        .Range("AH3").FormulaR1C1 = "移動"
                
        '[対応状況]編集（罫線、色）
        .Range("A2:F2").Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("A2:F2").Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("A2:F2").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:F2").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:F2").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:F2").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("A2:F2").Borders(xlInsideVertical).LineStyle = xlNone
        With .Range("A2:F2").Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        .Range("A2:F2").Merge
        
        '[現契約情報]編集（罫線、色）
        .Range("G2:U2").Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("G2:U2").Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("G2:U2").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("G2:U2").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("G2:U2").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("G2:U2").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("G2:U2").Borders(xlInsideVertical).LineStyle = xlNone
        With .Range("G2:U2").Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        .Range("G2:U2").Merge
        
        '[移動先情報]編集（罫線、色）
        .Range("V2:AH2").Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("V2:AH2").Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("V2:AH2").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("V2:AH2").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("V2:AH2").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("V2:AH2").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("V2:AH2").Borders(xlInsideVertical).LineStyle = xlNone
        With .Range("V2:AH2").Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        .Range("V2:AH2").Merge
        
        '[加瀬移動時の連絡状況]編集（罫線、色）
        .Range("AI2:AN2").Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("AF2:AL2").Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("AI2:AN2").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AI2:AN2").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AI2:AN2").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AI2:AN2").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("AI2:AN2").Borders(xlInsideVertical).LineStyle = xlNone
        With .Range("AI2:AN2").Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        .Range("AI2:AN2").Merge
    
        '[ ]編集（罫線、色）
        .Range("AO2:AS2").Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("AO2:AS2").Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("AO2:AS2").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AO2:AS2").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AO2:AS2").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AO2:AS2").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("AO2:AS2").Borders(xlInsideVertical).LineStyle = xlNone
        With .Range("AO2:AS2").Interior
            .ColorIndex = 15
            .Pattern = xlSolid
        End With
        .Range("AO2:AS2").Merge
            
        '[備考]編集（罫線、色）
        With .Range("AT2:AT3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Range("AT2:AT3").Merge
        .Range("AT2:AT3").Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("AT2:AT3").Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("AT2:AT3").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AT2:AT3").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AT2:AT3").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AT2:AT3").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("AT2:AT3").Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("AT2").FormulaR1C1 = "備考"
        
        '[完了]編集（罫線、色）
        With .Range("AU2:AU3")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Range("AU2:AU3").Merge
        .Range("AU2:AU3").Borders(xlDiagonalDown).LineStyle = xlNone
        .Range("AU2:AU3").Borders(xlDiagonalUp).LineStyle = xlNone
        With .Range("AU2:AU3").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AU2:AU3").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AU2:AU3").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("AU2:AU3").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("AU2:AU3").Borders(xlInsideHorizontal).LineStyle = xlNone
        
            
        'データ部編集(罫線)
        strRow = .Range("G3").End(xlDown).row
        With .Range("A2:AU" + strRow).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:AU" + strRow).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:AU" + strRow).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:AU" + strRow).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:AU" + strRow).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range("A2:AU" + strRow).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
                
        '数値項目右詰め
        .Range("L4:L" + strRow).HorizontalAlignment = xlRight
        .Range("Q4:Q" + strRow).HorizontalAlignment = xlRight
        .Range("O4:O" + strRow).HorizontalAlignment = xlRight
        .Range("Y4:Y" + strRow).HorizontalAlignment = xlRight
        .Range("AB4:AB" + strRow).HorizontalAlignment = xlRight
        .Range("AD4:AD" + strRow).HorizontalAlignment = xlRight
        
        
        '日付項目書式設定
        .Range("B4:F" + strRow).NumberFormatLocal = "m/d;@"
        .Range("AI4:AI" + strRow).NumberFormatLocal = "m/d;@"
        .Range("AK4:AL" + strRow).NumberFormatLocal = "m/d;@"
        .Range("AN4:AO" + strRow).NumberFormatLocal = "m/d;@"
        .Range("AR4:AS" + strRow).NumberFormatLocal = "m/d;@"
        
        .Cells.Select
        .Cells.EntireColumn.AutoFit
        .Range("L4").Select
        
        strTitle = strYardCode & " " & strYardName & "　　作成日時：" & Format(Now, "yyyy/mm/dd hh:mm")
        .Range("A1").FormulaR1C1 = strTitle
        
    End With


End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブル検索クエリ作成
'       MODULE_ID       : subMakeQueryWorkC
'       CREATE_DATE     : 2018/02/17
'       PARAM           : intMode     - 1=印刷プレビュー 2=Excel出力 3=印刷（定数宣言あり）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subMakeQueryWorkC(ByVal strYardCode As String, intMode As Integer)

On Error GoTo ErrorHandler

    Dim dbAccess        As Database
    Dim rs              As Recordset
    Dim blnError        As Boolean
    Dim intLoopCount    As Integer
    Dim strSQL          As String
    Dim wk              As String
    Dim wkDate          As String
    
    blnError = False
    intLoopCount = 0

    Set dbAccess = CurrentDb

    '近隣ヤード取得
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " ヤードコード "
    strSQL = strSQL & " ,ヤード名 "
    strSQL = strSQL & " ,距離 "
    strSQL = strSQL & " ,営業終了日 "
    strSQL = strSQL & " FROM " & P_WORK_TABLE
    
    strSQL = strSQL & " GROUP BY "
    strSQL = strSQL & " ヤードコード "
    strSQL = strSQL & " ,ヤード名 "
    strSQL = strSQL & " ,距離 "
    strSQL = strSQL & " ,営業終了日 "
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " CDbl(nz(距離,0)) "
    Set rs = dbAccess.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

    'データセット
    ReDim pNYARLists(NYARMAX)
    intLoopCount = 0
    pNYARListsJoin = ""
    Do While Not rs.EOF
    
        wkDate = Format$(rs.Fields("営業終了日"), "yyyymmdd")
        If Trim(wkDate) = "" Then
            wkDate = "21001231"
        End If
        'UPDATE 2018-06-17 start
        'If intLoopCount = 0 Or (intLoopCount <> 0 And wkdate > Format(pNYARLists(0).END_DATE, "yyyymmdd")) Then
        If intLoopCount = 0 Or _
           ((intLoopCount <> 0 And wkDate > Format(pNYARLists(0).END_DATE, "yyyymmdd")) And _
           (Form_FVS210.chk_営業日.VALUE = True Or (Form_FVS210.chk_営業日.VALUE = False And wkDate = "21001231"))) Then
        'UPDATE 2018-06-17 end
    
            pNYARLists(intLoopCount).YCODE = rs.Fields("ヤードコード")
            pNYARLists(intLoopCount).YNAME = rs.Fields("ヤード名")
            pNYARLists(intLoopCount).KIRO = Nz(rs.Fields("距離"), "")
            pNYARLists(intLoopCount).END_DATE = Nz(rs.Fields("営業終了日"), "")
            
            If rs.Fields("ヤードコード") <> strYardCode Then
                If pNYARListsJoin = "" Then
                    pNYARListsJoin = """" & rs.Fields("ヤードコード") & """"
                Else
                    pNYARListsJoin = pNYARListsJoin & ",""" & rs.Fields("ヤードコード") & """"
                End If
            End If
        
            intLoopCount = intLoopCount + 1
        
        End If
        
        If intLoopCount = NYARMAX + 1 Then
            Exit Do
        End If
        rs.MoveNext
    Loop
    'Debug.Print "intLoopCount=" & intLoopCount

    'サイズ取得
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & "  段 "
    strSQL = strSQL & " ,実帖 "
    strSQL = strSQL & " FROM " & P_WORK_TABLE
    strSQL = strSQL & " WHERE 実帖 <> ""0"" "
    strSQL = strSQL & " GROUP BY "
    strSQL = strSQL & "  段 "
    strSQL = strSQL & " ,実帖 "
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " 段 desc "
    Set rs = dbAccess.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)

    'データ追加
    pSIZEListsUpRows = 0
    pSIZEListsDownRows = 0
    ReDim pSIZELists(rs.RecordCount - 1)
    intLoopCount = 0
    Do While Not rs.EOF
        
        pSIZELists(intLoopCount).STEP = rs.Fields("段")
        pSIZELists(intLoopCount).SIZE = rs.Fields("実帖")
        
        If rs.Fields("段") = "２段" Then
            pSIZEListsUpRows = pSIZEListsUpRows + 1
        Else
            pSIZEListsDownRows = pSIZEListsDownRows + 1
        End If

        
        intLoopCount = intLoopCount + 1
        
        rs.MoveNext
    Loop


    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "subMakeQueryWorkC" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'
Sub outlog(strMessage As String)
    Dim Locator             As Object
    Dim SERVICE             As Object
    Dim Wmi                 As Object
    Dim strPath             As String
    Dim strFile             As String
    Dim iFlno               As Integer
    
    On Error Resume Next
    'strPath = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'MSZZ024'")
    If strPath = "" Then
        strPath = Application.CurrentDb.NAME
        strPath = Left(strPath, Len(strPath) - Len(Dir(strPath)))
        strPath = strPath & "ErrorLog\"
        If Dir(strPath, vbDirectory) = "" Then
            Call MkDir(strPath)
        End If
    End If
    strFile = Format(Now, "yyyymmddhhnnss")
    
    Set Wmi = Nothing
    Set Locator = Nothing
    Set SERVICE = Nothing

    iFlno = FreeFile()
    Open strPath & strFile & ".log" For Append As #iFlno
    Print #iFlno, strMessage
    Close #iFlno
End Sub
