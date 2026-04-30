Attribute VB_Name = "LS0060A"
'****************************  strat of program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : コンテナ管理システム
'        SUB_SYSTEM_NAME : 帳票
'
'        PROGRAM_NAME    : コンテナ在庫一覧表(データ更新)
'        PROGRAM_ID      : LS0060A
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2019/11/20
'        CERATER         : Y.WADA
'        Ver             : 0.0
'
'        UPDATE          : 2020/02/15
'        UPDATER         : Y.WADA
'        Ver             : 0.1
'                        : 在庫数でCBYS_RECDを併せて参照するように変更
'
'        UPDATE          : 2020/06/04
'        UPDATER         : Y.WADA
'        Ver             : 0.2
'                        : 枝番00の条件を外す
'
'        UPDATE          : 2020/06/09
'        UPDATER         : Y.WADA
'        Ver             : 0.3
'                        : 仕入商品区分の判断方法変更
'                          タイトル／シート名変更
'                          　自社ISOコンテナ　 →　スタンダードコンテナ（ISO）
'                          　自社JISコンテナ　 →　建築確認コンテナ（JIS）
'                          　自社コンテナ以外　→　その他コンテナなど
'                          各シート共通
'                              ②出港数・・・800017：直送輸入済みコンテナ分を含む
'                              ④海外からの入庫数　※②と同様・営業ヤードへ直送分は除く・・・800017：直送輸入済みコンテナ分を含まない
'                              ⑦営業ヤードへの出庫数・・・800017：直送輸入済みコンテナ分を含まない
'                              ⑨梶山ヤードからの入庫数※⑦と同様・港から営業ヤードへ直送分含む・・・800017：直送輸入済みコンテナ分を含む
'                          自社コンテナ以外
'                              サイズ、ドア数を追加
'                              ※バイクボックスにサイズとドア数がありそれぞれ分かれているので追加してください。台車やｶﾞｰﾄﾞﾏﾝﾎﾞｯｸｽはサイズなどないので空白でかまいません。
'
'        UPDATE          : 2020/06/12
'        UPDATER         : Y.WADA
'        Ver             : 0.4
'                        : クエリ修正
'
'        UPDATE          : 2020/06/18
'        UPDATER         : Y.WADA
'        Ver             : 0.5
'                        : 直送分で、月跨ぎの移動（例：５月直送→６月に営業ヤードへ設置）の対応
'
'        UPDATE          : 2020/07/07
'        UPDATER         : Y.WADA
'        Ver             : 0.6
'                        : 項目追加：港保管、不足在庫
'
'        UPDATE          : 2020/07/22
'        UPDATER         : Y.WADA
'        Ver             : 0.6
'                        : 港保管のクエリ修正
'
'        UPDATE          : 2022/09/30
'        UPDATER         : K.KINEBUCHI
'        Ver             : 0.8
'                        : １）港保管クエリ修正
'                          ２）中古購入コンテナ対応
'
'        UPDATE          : 2022/10/26
'        UPDATER         : N.IMAI
'        Ver             : 1.0
'                        : 在庫総合計の作成、出力を追加
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_PROG_ID     As String = "LS0060A"

Private strProg             As String
Private strUser             As String
Private strDate             As String
Private strTime             As String

'==============================================================================*
'
'       MODULE_NAME     : 対象の設置日取得
'       MODULE_ID       : LS0060A_SETDATE_YM
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       RETURN          : 設置日（YYYYMM）・・・締日を考慮した当月
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function LS0060A_SETDATE_YM() As String

    On Error GoTo ErrorHandler

    Dim strClDay            As String
    Dim strYMD              As String
    
    LS0060A_SETDATE_YM = ""
    strYMD = Format(Now, "yyyymmdd")
    
    '締日取得
    strClDay = Right("00" & Nz(DLookup("PGPAT_PARAN", "PGPA_TABL", "PGPAT_PGP1B = 'LF0060' AND PGPAT_PGP2B =" & "'CLDAY'")), 2)
    
    '設置日（YYYYMM）
    If Right(strYMD, 2) <= strClDay Then
        '締日以前は、前月
        LS0060A_SETDATE_YM = Format(DateAdd("m", -1, CDate(Format(strYMD, "####/##/##"))), "yyyymm")
    Else
        '締日より後は、当月
        LS0060A_SETDATE_YM = Left(strYMD, 6)
    End If

Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "LS0060A_SETDATE_YM" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function


'==============================================================================*
'
'       MODULE_NAME     : 主処理
'       MODULE_ID       : LS0060A_M00
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       PARAM           : strSETDATE_YM         設置日（YYYYMM）、省略時は締日を考慮した当月(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function LS0060A_M00(Optional ByVal strSETDATE_YM As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim strBUMOC            As String
    Dim objCon              As Object
    Dim blRet               As Boolean
    Dim strErrMsg           As String
    Dim lngCnt              As Long
    
    LS0060A_M00 = False
    
    Call MSZZ003_M00(P_PROG_ID, "0", "LS0060A_M00")
    Call MSZZ003_M00(P_PROG_ID, "8", "====================================================")

    
    strProg = Left(P_PROG_ID, 11)
    strUser = Left(LsGetUserName(), 8)
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")

    If strSETDATE_YM = "" Then
        strSETDATE_YM = LS0060A_SETDATE_YM
    End If

    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))

    Set objCon = ADODB_Connection(strBUMOC)
    On Error GoTo ErrorHandler1
    
    'トランザクション開始
    objCon.BeginTrans
    On Error GoTo ErrorHandler2
    
    '在庫集計トラン・対象月を上書きのため削除
    
    lngCnt = ADODB_Execute(SqlDeleteZAIK_SUMM(strSETDATE_YM), objCon)
    Call MSZZ003_M00(P_PROG_ID, "8", "在庫集計トラン削除件数  = " & Format(lngCnt))

    lngCnt = ADODB_Execute(SqlInsertZAIK_SUMM(strSETDATE_YM), objCon)
    Call MSZZ003_M00(P_PROG_ID, "8", "在庫集計トラン登録件数  = " & Format(lngCnt))
    
    'INSERT 2022/10/26 N.IMAI Start
    If fncPLS0060(objCon, strSETDATE_YM) Then
        Call MSZZ003_M00(P_PROG_ID, "8", "在庫総合計トラン登録")
    End If
    'INSERT 2022/10/26 N.IMAI End
    
    'トランザクション終了（Commit)
    objCon.CommitTrans
    Call MSZZ003_M00(P_PROG_ID, "8", "Commit")
'    objCon.RollbackTrans  '★debug
'    Call MSZZ003_M00(P_PROG_ID, "8", "Rollback(★debug)")  '★debug
    
    On Error GoTo ErrorHandler1
        
    objCon.Close
    On Error GoTo ErrorHandler
            
    Call MSZZ003_M00(P_PROG_ID, "8", "====================================================")
    Call MSZZ003_M00(P_PROG_ID, "1", "LS0060A_M00")
    
    LS0060A_M00 = True
Exit Function

ErrorHandler2:
    objCon.RollbackTrans
'    Call MSZZ003_M00(P_PROG_ID, "8", "Rollback")
ErrorHandler1:
    objCon.Close
ErrorHandler:
    strErrMsg = Err.Description
    Call MSZZ024_M00("LS0060A_M00", False)
    Call MSZZ003_M00(P_PROG_ID, "9", strErrMsg)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 在庫集計トラン削除
'       MODULE_ID       : SqlDeleteZAIK_SUMM
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       PARAM           : strSETDATE_YM         設置日（YYYYMM）(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SqlDeleteZAIK_SUMM(strSETDATE_YM As String) As String
    On Error GoTo ErrorHandler
    
    Dim strSQL              As String
    
    strSQL = ""
    strSQL = strSQL & vbCrLf & "DECLARE @SETDATE_YM CHAR(6) = '" & strSETDATE_YM & "';    --設置日（年月）"
    strSQL = strSQL & vbCrLf & "DELETE"
    strSQL = strSQL & vbCrLf & "FROM    ZAIK_SUMM"
    strSQL = strSQL & vbCrLf & "WHERE   ZAIKS_ZAISD = @SETDATE_YM"
    strSQL = strSQL & vbCrLf & ";"
    
    SqlDeleteZAIK_SUMM = strSQL
Exit Function
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "SqlDeleteZAIK_SUMM" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 在庫集計トラン登録
'       MODULE_ID       : SqlInsertZAIK_SUMM
'       CREATE_DATE     : 2019/11/20            Y.WADA
'       PARAM           : strSETDATE_YM         設置日（YYYYMM）(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SqlInsertZAIK_SUMM(strSETDATE_YM As String) As String
    On Error GoTo ErrorHandler
    
    Dim strSQL              As String
    Dim strDataSource       As String
    
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    
'DELETE 2020/07/07 Y.WADA Start
'    strSQL = ""
'    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSED       varchar(8)  =   '" & strDate & "';"
'    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSEJ       varchar(6)  =   '" & strTime & "';"
'    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSPB       varchar(11) =   '" & strProg & "';"
'    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSUB       varchar(8)  =   '" & strUser & "';"
'    strSQL = strSQL & vbCrLf & "DECLARE @SETDATE_YM CHAR(6) = '" & strSETDATE_YM & "';    --設置日（年月）"
'    strSQL = strSQL & vbCrLf & "DECLARE @H_BUMOC    CHAR(1) = 'ﾗ';         --発注・部門コード"
'    strSQL = strSQL & vbCrLf & "DECLARE @S_BUMOC    CHAR(1) = 'H';         --出港・部門コード;"
'    strSQL = strSQL & vbCrLf & "--------------------------------"
'    strSQL = strSQL & vbCrLf & "-- 在庫集計トランの登録"
'
'    'INSERT 2020/02/15 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "    WITH CBYS_TRAN2_tmp1 AS"
'    strSQL = strSQL & vbCrLf & "    ("
'    'INSERT 2020/06/18 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "        SELECT"
'    strSQL = strSQL & vbCrLf & "            *"
'    strSQL = strSQL & vbCrLf & "            ,   CASE"
'    strSQL = strSQL & vbCrLf & "                WHEN"
'    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800012')   --800012:輸入済みコンテナ"
'    strSQL = strSQL & vbCrLf & "                    THEN"
'    strSQL = strSQL & vbCrLf & "                        '②海外・出港数'"
'    strSQL = strSQL & vbCrLf & "                WHEN"
'    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800017')   --800017：直送輸入済みコンテナ"
'    strSQL = strSQL & vbCrLf & "                    THEN"
'    strSQL = strSQL & vbCrLf & "                        '②海外・出港数（直送）'"
'    strSQL = strSQL & vbCrLf & "                WHEN"
'    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800007','800008','800009','800013')    --800007:廃棄ヤード　800008:撤去済みヤード　800009:解約ヤード　800013:入替コンテナ"
'    strSQL = strSQL & vbCrLf & "                    THEN"
'    strSQL = strSQL & vbCrLf & "                        '⑤梶山・撤去数'"
'    strSQL = strSQL & vbCrLf & "                WHEN"
'    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800010')   --800010:販売ヤード"
'    strSQL = strSQL & vbCrLf & "                    THEN"
'    strSQL = strSQL & vbCrLf & "                        '⑥梶山・売却数'"
'    strSQL = strSQL & vbCrLf & "                ELSE"
'    strSQL = strSQL & vbCrLf & "                        --'⑦梶山・営業ヤードへの出庫数'"
'    strSQL = strSQL & vbCrLf & "                        '⑨梶山ヤードからの入庫数'"
'    strSQL = strSQL & vbCrLf & "                END AS [列名]"
'    strSQL = strSQL & vbCrLf & "        FROM ("
'    'INSERT 2020/06/18 Y.WADA End
'    strSQL = strSQL & vbCrLf & "        SELECT"
'    strSQL = strSQL & vbCrLf & "        1 as tbl"
'    strSQL = strSQL & vbCrLf & "            ,CBYST_BOXNO"
'    strSQL = strSQL & vbCrLf & "            ,CBYST_EDANO"
'    strSQL = strSQL & vbCrLf & "            ,CBYST_SYUBETSU"
'    strSQL = strSQL & vbCrLf & "            ,CBYST_YCODE"
'    strSQL = strSQL & vbCrLf & "        --    ,CBYST_STATUS"
'    strSQL = strSQL & vbCrLf & "            ,CBYST_SETDATE"
'    strSQL = strSQL & vbCrLf & "            ,CBYST_SETFLOOR"
'    strSQL = strSQL & vbCrLf & "            ,CBYST_BIKON"
'    strSQL = strSQL & vbCrLf & "        FROM"
'    strSQL = strSQL & vbCrLf & "            CBYS_TRAN"
'    strSQL = strSQL & vbCrLf & "        WHERE"
'    strSQL = strSQL & vbCrLf & "            LEFT(CBYST_SETDATE, 6) = @SETDATE_YM"
'    strSQL = strSQL & vbCrLf & "        UNION ALL"
'    strSQL = strSQL & vbCrLf & "        SELECT"
'    strSQL = strSQL & vbCrLf & "        2 as tbl"
'    strSQL = strSQL & vbCrLf & "            ,CBYSR_BOXNO"
'    strSQL = strSQL & vbCrLf & "            ,CBYSR_EDANO"
'    strSQL = strSQL & vbCrLf & "            ,CBYSR_SYUBETSU"
'    strSQL = strSQL & vbCrLf & "            ,CBYSR_YCODE"
'    strSQL = strSQL & vbCrLf & "            ,CBYSR_SETDATE"
'    strSQL = strSQL & vbCrLf & "            ,CBYSR_SETFLOOR"
'    strSQL = strSQL & vbCrLf & "            ,CBYSR_BIKON"
'    strSQL = strSQL & vbCrLf & "        FROM"
'    strSQL = strSQL & vbCrLf & "            CBYS_RECD"
'    strSQL = strSQL & vbCrLf & "        WHERE"
'    strSQL = strSQL & vbCrLf & "            LEFT(CBYSR_SETDATE, 6) = @SETDATE_YM"
'    'INSERT 2020/06/18 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "    ) a"
'    'INSERT 2020/06/18 Y.WADA End
'    strSQL = strSQL & vbCrLf & "    )"
'    strSQL = strSQL & vbCrLf & "    , CBYS_TRAN2_tmp2 AS"
'    strSQL = strSQL & vbCrLf & "    ("
''DELETE 2020/06/18 Y.WADA Start
''    strSQL = strSQL & vbCrLf & "        SELECT"
''    strSQL = strSQL & vbCrLf & "            ROW_NUMBER()"
''    strSQL = strSQL & vbCrLf & "                OVER "
''    strSQL = strSQL & vbCrLf & "                ("
''    strSQL = strSQL & vbCrLf & "                    PARTITION BY"
''    strSQL = strSQL & vbCrLf & "                        CASE"
''    strSQL = strSQL & vbCrLf & "                        WHEN"
''    strSQL = strSQL & vbCrLf & "                                CBYST_YCODE IN ('800012')   --800012:輸入済みコンテナ"
''    strSQL = strSQL & vbCrLf & "                            THEN"
''    strSQL = strSQL & vbCrLf & "                                '②海外・出港数'"
''    'INSERT 2020/06/09 Y.WADA Start
''    strSQL = strSQL & vbCrLf & "                        WHEN"
''    strSQL = strSQL & vbCrLf & "                                CBYST_YCODE IN ('800017')   --800017：直送輸入済みコンテナ"
''    strSQL = strSQL & vbCrLf & "                            THEN"
''    strSQL = strSQL & vbCrLf & "                                '②海外・出港数（直送）'"
''    'INSERT 2020/06/09 Y.WADA End
''    strSQL = strSQL & vbCrLf & "                        WHEN"
''    strSQL = strSQL & vbCrLf & "                                CBYST_YCODE IN ('800007','800008','800009','800013')    --800007:廃棄ヤード　800008:撤去済みヤード　800009:解約ヤード　800013:入替コンテナ"
''    strSQL = strSQL & vbCrLf & "                            THEN"
''    strSQL = strSQL & vbCrLf & "                                '⑤梶山・撤去数'"
''    strSQL = strSQL & vbCrLf & "                        WHEN"
''    strSQL = strSQL & vbCrLf & "                                CBYST_YCODE IN ('800010')   --800010:販売ヤード"
''    strSQL = strSQL & vbCrLf & "                            THEN"
''    strSQL = strSQL & vbCrLf & "                                '⑥梶山・売却数'"
''    strSQL = strSQL & vbCrLf & "                        ELSE"
''    strSQL = strSQL & vbCrLf & "                                '⑦梶山・営業ヤードへの出庫数'"
''    strSQL = strSQL & vbCrLf & "                        END"
''    strSQL = strSQL & vbCrLf & "                    ,   CBYST_BOXNO"
''    strSQL = strSQL & vbCrLf & "                    ,   CBYST_EDANO"
''    strSQL = strSQL & vbCrLf & "                    ORDER BY CBYST_SETDATE DESC, tbl"
''    strSQL = strSQL & vbCrLf & "                ) AS rno"
''DELETE 2020/06/18 Y.WADA End
'    'INSERT 2020/06/18 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "        SELECT"
'    strSQL = strSQL & vbCrLf & "            ROW_NUMBER()"
'    strSQL = strSQL & vbCrLf & "                OVER "
'    strSQL = strSQL & vbCrLf & "                ("
'    strSQL = strSQL & vbCrLf & "                    PARTITION BY"
'    strSQL = strSQL & vbCrLf & "                        [列名]"
'    strSQL = strSQL & vbCrLf & "                    ,   CBYST_BOXNO"
'    strSQL = strSQL & vbCrLf & "                    ,   CBYST_EDANO"
'    strSQL = strSQL & vbCrLf & "                    ORDER BY CBYST_SETDATE DESC, tbl"
'    strSQL = strSQL & vbCrLf & "                ) AS rno"
'    'INSERT 2020/06/18 Y.WADA End
'    strSQL = strSQL & vbCrLf & "            , *"
'    strSQL = strSQL & vbCrLf & "        FROM"
'    strSQL = strSQL & vbCrLf & "            CBYS_TRAN2_tmp1"
'    'INSERT 2020/06/18 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "        WHERE"
'    strSQL = strSQL & vbCrLf & "            CBYST_YCODE NOT IN ('800011', '800014') --800011:販売済みコンテナ、800014:梶山コンテナ置き場"
'    'INSERT 2020/06/18 Y.WADA End
'    strSQL = strSQL & vbCrLf & "    )"
''    strSQL = strSQL & vbCrLf & "    , CBYS_TRAN2 AS"           'DELETE 2020/06/18 Y.WADA
'    strSQL = strSQL & vbCrLf & "    , CBYS_TRAN2_tmp3 AS"       'INSERT 2020/06/18 Y.WADA
'    strSQL = strSQL & vbCrLf & "    ("
'    strSQL = strSQL & vbCrLf & "        SELECT * FROM CBYS_TRAN2_tmp2"
'    strSQL = strSQL & vbCrLf & "        WHERE rno = 1"
'    strSQL = strSQL & vbCrLf & "    )"
'    'INSERT 2020/06/18 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "    , CBYS_TRAN2 AS"
'    strSQL = strSQL & vbCrLf & "    ("
'    strSQL = strSQL & vbCrLf & "        SELECT * FROM CBYS_TRAN2_tmp3"
'    strSQL = strSQL & vbCrLf & "        UNION ALL"
'    strSQL = strSQL & vbCrLf & "        SELECT"
'    strSQL = strSQL & vbCrLf & "            rno"
'    strSQL = strSQL & vbCrLf & "        ,   tbl"
'    strSQL = strSQL & vbCrLf & "        ,   CBYST_BOXNO"
'    strSQL = strSQL & vbCrLf & "        ,   CBYST_EDANO"
'    strSQL = strSQL & vbCrLf & "        ,   CBYST_SYUBETSU"
'    strSQL = strSQL & vbCrLf & "        ,   CBYST_YCODE"
'    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETDATE"
'    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETFLOOR"
'    strSQL = strSQL & vbCrLf & "        ,   CBYST_BIKON"
'    strSQL = strSQL & vbCrLf & "        ,   '⑦梶山・営業ヤードへの出庫数' AS [列名]"
'    strSQL = strSQL & vbCrLf & "        FROM"
'    strSQL = strSQL & vbCrLf & "            CBYS_TRAN2_tmp3"
'    strSQL = strSQL & vbCrLf & "        WHERE"
'    strSQL = strSQL & vbCrLf & "            [列名]  =   '⑨梶山ヤードからの入庫数'"
'    strSQL = strSQL & vbCrLf & "        AND NOT EXISTS  --直送分を除外"
'    strSQL = strSQL & vbCrLf & "            ("
'    strSQL = strSQL & vbCrLf & "                SELECT *"
'    strSQL = strSQL & vbCrLf & "                FROM"
'    strSQL = strSQL & vbCrLf & "                    ("
'    strSQL = strSQL & vbCrLf & "                        SELECT  TOP 1"
'    strSQL = strSQL & vbCrLf & "                            CBYSR_YCODE"
'    strSQL = strSQL & vbCrLf & "                        FROM"
'    strSQL = strSQL & vbCrLf & "                            CBYS_RECD"
'    strSQL = strSQL & vbCrLf & "                        WHERE"
'    strSQL = strSQL & vbCrLf & "                            CBYSR_SETDATE   <=  CBYST_SETDATE"
'    strSQL = strSQL & vbCrLf & "                        AND CBYSR_BOXNO     =   CBYST_BOXNO"
'    strSQL = strSQL & vbCrLf & "                        AND CBYSR_EDANO     =   CBYST_EDANO"
'    strSQL = strSQL & vbCrLf & "                        AND CBYSR_YCODE     <>  CBYST_YCODE"
'    strSQL = strSQL & vbCrLf & "                        ORDER BY CBYSR_SETDATE DESC"
'    strSQL = strSQL & vbCrLf & "                    ) a"
'    strSQL = strSQL & vbCrLf & "                WHERE"
'    strSQL = strSQL & vbCrLf & "                    CBYSR_YCODE = '800017'  --800017：直送輸入済みコンテナ"
'    strSQL = strSQL & vbCrLf & "            )"
'    strSQL = strSQL & vbCrLf & "    )"
'    'INSERT 2020/06/18 Y.WADA End
'    'INSERT 2020/02/15 Y.WADA End
'    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
'    strSQL = strSQL & vbCrLf & "--①海外発注"
'    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
''    strSQL = strSQL & vbCrLf & "WITH   [海外発注1] AS"  'DELETE 2020/02/15 Y.WADA
'    strSQL = strSQL & vbCrLf & ",   [海外発注1] AS"  'INSERT 2020/02/15 Y.WADA
'    strSQL = strSQL & vbCrLf & "("
'    strSQL = strSQL & vbCrLf & "    SELECT"
'    strSQL = strSQL & vbCrLf & "        '①発注数'  [列名]"
'    strSQL = strSQL & vbCrLf & "    ,   LEFT(MDOCT_TYUMD, 6)    AS [年月]"
'    strSQL = strSQL & vbCrLf & "    ,   CASE"
'    strSQL = strSQL & vbCrLf & "        WHEN"
'    strSQL = strSQL & vbCrLf & "                ISNULL(HANBM_DOORQ,0) != 0      --ドア数"
'    strSQL = strSQL & vbCrLf & "            AND HANBM_SYUBI = 2                 --1:新型（JIS）　2:旧型（ISO）"
'    strSQL = strSQL & vbCrLf & "            THEN"
'    strSQL = strSQL & vbCrLf & "                1   --ISOコンテナ"
'    strSQL = strSQL & vbCrLf & "        WHEN"
'    strSQL = strSQL & vbCrLf & "                ISNULL(HANBM_DOORQ,0) != 0      --ドア数"
'    strSQL = strSQL & vbCrLf & "            AND HANBM_SYUBI = 1                 --1:新型（JIS）　2:旧型（ISO）"
'    strSQL = strSQL & vbCrLf & "            THEN"
'    strSQL = strSQL & vbCrLf & "                2   --JISコンテナ"
'    strSQL = strSQL & vbCrLf & "        WHEN"
'    strSQL = strSQL & vbCrLf & "                HANBM_SYOHN LIKE '%バイク%'"
'    strSQL = strSQL & vbCrLf & "            THEN"
'    strSQL = strSQL & vbCrLf & "                3   --バイクボックス"
'    strSQL = strSQL & vbCrLf & "        WHEN"
'    strSQL = strSQL & vbCrLf & "                    HANBM_SYOHN LIKE '%台車%'"
'    strSQL = strSQL & vbCrLf & "            THEN"
'    strSQL = strSQL & vbCrLf & "                4   --台車"
'    strSQL = strSQL & vbCrLf & "        WHEN"
'    strSQL = strSQL & vbCrLf & "                    HANBM_SYOHN LIKE '%ガードマン%'"
'    strSQL = strSQL & vbCrLf & "            THEN"
'    strSQL = strSQL & vbCrLf & "                5   --ガードマンボックス"
'    strSQL = strSQL & vbCrLf & "        ELSE"
'    strSQL = strSQL & vbCrLf & "                99  --その他"
'    strSQL = strSQL & vbCrLf & "        END AS [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "    ,   HANBM_SYOHN         AS [商品名]     --商品名"
'    strSQL = strSQL & vbCrLf & "    ,   HANBM_SIZEI         AS [サイズ]"
'    strSQL = strSQL & vbCrLf & "    ,   HANBM_DOORQ         AS [ドア数]"
'    strSQL = strSQL & vbCrLf & "    ,   HJUKT_HANBQ         AS [本数]"
'    strSQL = strSQL & vbCrLf & "    ,   HJUKT_USEDI         AS [新品・中古区分]     --0:新品 1:中古"
'    strSQL = strSQL & vbCrLf & "    "
'    strSQL = strSQL & vbCrLf & "    FROM"
'    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.KEIY_TRAN"
'    strSQL = strSQL & vbCrLf & "    INNER JOIN"
'    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.HJUK_TRAN"
'    strSQL = strSQL & vbCrLf & "    ON"
'    strSQL = strSQL & vbCrLf & "        HJUKT_KEIYB = KEIYT_KEIYB"
'    strSQL = strSQL & vbCrLf & "    INNER JOIN"
'    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.MDOC_TRAN"
'    strSQL = strSQL & vbCrLf & "    ON"
'    strSQL = strSQL & vbCrLf & "        MDOCT_KEIYB = KEIYT_KEIYB"
'    strSQL = strSQL & vbCrLf & "    AND"
'    strSQL = strSQL & vbCrLf & "        --MDOCT_TYUMD BETWEEN　@SETDATE_S AND @SETDATE_E"
'    strSQL = strSQL & vbCrLf & "        LEFT(MDOCT_TYUMD, 6) = @SETDATE_YM"
'    strSQL = strSQL & vbCrLf & "    AND"
'    strSQL = strSQL & vbCrLf & "        MDOCT_TNAMN LIKE '%自社%'"
'    strSQL = strSQL & vbCrLf & "    INNER JOIN"
'    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.HANB_MAST"
'    strSQL = strSQL & vbCrLf & "    ON"
'    strSQL = strSQL & vbCrLf & "        HANBM_SYOHC = HJUKT_SYOHC"
'    strSQL = strSQL & vbCrLf & "    AND"
'    strSQL = strSQL & vbCrLf & "        HANBM_SYOHI = 1                 --1:仕入商品　2:不可商品"
'    strSQL = strSQL & vbCrLf & "    AND "
'    strSQL = strSQL & vbCrLf & "        ISNULL(HANBM_SYUTI,0) = 0       --トランク用部品種別"
'    strSQL = strSQL & vbCrLf & "    WHERE"
'    strSQL = strSQL & vbCrLf & "        KEIYT_BUMOC = @H_BUMOC"
'    strSQL = strSQL & vbCrLf & ")"
''    strSQL = strSQL & vbCrLf & "--, [海外発注2] AS"
''    strSQL = strSQL & vbCrLf & "--("
''    strSQL = strSQL & vbCrLf & "--    SELECT "
''    strSQL = strSQL & vbCrLf & "--        [列名]"
''    strSQL = strSQL & vbCrLf & "--    ,   [年月]"
''    strSQL = strSQL & vbCrLf & "--    ,   [仕入商品区分]"
''    strSQL = strSQL & vbCrLf & "--    ,   NAME_VALUE_FROM AS [在庫種類]"
''    strSQL = strSQL & vbCrLf & "--    ,   NAME_NAME  AS [商品名]"
''    strSQL = strSQL & vbCrLf & "--    ,   IIF(NAME_VALUE_FROM = 3, NULL, [サイズ]) AS [サイズ] "
''    strSQL = strSQL & vbCrLf & "--    ,   IIF(NAME_VALUE_FROM = 3, NULL, [ドア数]) AS [ドア数] "
''    strSQL = strSQL & vbCrLf & "--    ,   [本数]"
''    strSQL = strSQL & vbCrLf & "--    ,   [新品・中古区分]"
''    strSQL = strSQL & vbCrLf & "--    FROM [海外発注1]"
''    strSQL = strSQL & vbCrLf & "--    LEFT JOIN NAME_MAST"
''    strSQL = strSQL & vbCrLf & "--        ON  NAME_ID =451"
''    strSQL = strSQL & vbCrLf & "--        AND NAME_CODE = [仕入商品区分]"
''    strSQL = strSQL & vbCrLf & "--)"
'    strSQL = strSQL & vbCrLf & ", [海外発注] AS"
'    strSQL = strSQL & vbCrLf & "("
'    strSQL = strSQL & vbCrLf & "    SELECT "
'    strSQL = strSQL & vbCrLf & "        [列名]"
'    strSQL = strSQL & vbCrLf & "    ,   [年月]"
'    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
'    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
'    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
'    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
'    strSQL = strSQL & vbCrLf & "    ,   SUM([本数]) [本数]"
'    strSQL = strSQL & vbCrLf & "    FROM [海外発注1]"
'    strSQL = strSQL & vbCrLf & "    GROUP BY "
'    strSQL = strSQL & vbCrLf & "        [列名]"
'    strSQL = strSQL & vbCrLf & "    ,   [年月]"
'    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
'    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
'    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
'    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
'    strSQL = strSQL & vbCrLf & ")"
'    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
'    strSQL = strSQL & vbCrLf & "--②海外・出港数"
'    strSQL = strSQL & vbCrLf & "--⑤梶山・撤去数"
'    strSQL = strSQL & vbCrLf & "--⑥梶山・売却数"
'    strSQL = strSQL & vbCrLf & "--⑦梶山・営業ヤードへの出庫数"
'    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
'    strSQL = strSQL & vbCrLf & ", [出港1] AS"
'    strSQL = strSQL & vbCrLf & "("
'    strSQL = strSQL & vbCrLf & "    SELECT"
'    'DELETE 2020/06/18 Y.WADA Start
'    '    strSQL = strSQL & vbCrLf & "        CASE"
'    '    strSQL = strSQL & vbCrLf & "        WHEN"
'    '    strSQL = strSQL & vbCrLf & "                CBYST_YCODE IN ('800012')   --800012:輸入済みコンテナ"
'    '    strSQL = strSQL & vbCrLf & "            THEN"
'    '    strSQL = strSQL & vbCrLf & "                '②海外・出港数'"
'    '    'INSERT 2020/06/09 Y.WADA Start
'    '    strSQL = strSQL & vbCrLf & "        WHEN"
'    '    strSQL = strSQL & vbCrLf & "                CBYST_YCODE IN ('800017')   --800017：直送輸入済みコンテナ"
'    '    strSQL = strSQL & vbCrLf & "            THEN"
'    '    strSQL = strSQL & vbCrLf & "                '②海外・出港数（直送）'"
'    '    'INSERT 2020/06/09 Y.WADA End
'    '    strSQL = strSQL & vbCrLf & "        WHEN"
'    '    strSQL = strSQL & vbCrLf & "                CBYST_YCODE IN ('800007','800008','800009','800013')    --800007:廃棄ヤード　800008:撤去済みヤード　800009:解約ヤード　800013:入替コンテナ"
'    '    strSQL = strSQL & vbCrLf & "            THEN"
'    '    strSQL = strSQL & vbCrLf & "                '⑤梶山・撤去数'"
'    '    strSQL = strSQL & vbCrLf & "        WHEN"
'    '    strSQL = strSQL & vbCrLf & "                CBYST_YCODE IN ('800010')   --800010:販売ヤード"
'    '    strSQL = strSQL & vbCrLf & "            THEN"
'    '    strSQL = strSQL & vbCrLf & "                '⑥梶山・売却数'"
'    '    strSQL = strSQL & vbCrLf & "        ELSE"
'    '    strSQL = strSQL & vbCrLf & "                '⑦梶山・営業ヤードへの出庫数'"
'    '    strSQL = strSQL & vbCrLf & "        END AS [列名]"
'    'DELETE 2020/06/18 Y.WADA End
'    strSQL = strSQL & vbCrLf & "        [列名]" 'INSERT 2020/06/18 Y.WADA
'    strSQL = strSQL & vbCrLf & "        ,   LEFT(CBYST_SETDATE, 6)  AS [年月]"
'    strSQL = strSQL & vbCrLf & "        ,   CASE"
'    strSQL = strSQL & vbCrLf & "            WHEN"
'    strSQL = strSQL & vbCrLf & "                    CBOXM_CTBNR = 1     --1:海上コンテナ　99:--"
'    strSQL = strSQL & vbCrLf & "                AND CBOXM_CTYTO = 1     --1:荷物収納　2:バイク収納"
'    strSQL = strSQL & vbCrLf & "                AND CBOXM_TYPE = 2      --1:新型（JIS）　2:旧型（ISO）"
'    strSQL = strSQL & vbCrLf & "                THEN"
'    strSQL = strSQL & vbCrLf & "                    1   --ISOコンテナ"
'    strSQL = strSQL & vbCrLf & "            WHEN"
'    strSQL = strSQL & vbCrLf & "                    CBOXM_CTBNR = 1     --1:海上コンテナ　99:--"
'    strSQL = strSQL & vbCrLf & "                AND CBOXM_CTYTO = 1     --1:荷物収納　2:バイク収納"
'    strSQL = strSQL & vbCrLf & "                AND CBOXM_TYPE = 1      --1:新型（JIS）　2:旧型（ISO）"
'    strSQL = strSQL & vbCrLf & "                THEN"
'    strSQL = strSQL & vbCrLf & "                    2   --JISコンテナ"
'    strSQL = strSQL & vbCrLf & "            WHEN"
'    strSQL = strSQL & vbCrLf & "                CBOXM_CTYTO = 2         --1:荷物収納　2:バイク収納"
'    strSQL = strSQL & vbCrLf & "                THEN"
'    strSQL = strSQL & vbCrLf & "                    3   --バイクボックス"
''DELETE 2020/06/09 Y.WADA Start
''    strSQL = strSQL & vbCrLf & "            WHEN"
''    strSQL = strSQL & vbCrLf & "                        CBYST_BIKON LIKE '%台車%'"
''    strSQL = strSQL & vbCrLf & "                THEN"
''    strSQL = strSQL & vbCrLf & "                    4   --台車"
''    strSQL = strSQL & vbCrLf & "            WHEN"
''    strSQL = strSQL & vbCrLf & "                        CBYST_BIKON LIKE '%ガードマン%'"
''    strSQL = strSQL & vbCrLf & "                THEN"
''    strSQL = strSQL & vbCrLf & "                    5   --ガードマンボックス"
''DELETE 2020/06/09 Y.WADA End
'    'INSERT 2020/06/09 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "            WHEN"
'    strSQL = strSQL & vbCrLf & "                        CBOXM_CTYTO = 30 --30:台車"
'    strSQL = strSQL & vbCrLf & "                THEN"
'    strSQL = strSQL & vbCrLf & "                    4   --台車"
'    strSQL = strSQL & vbCrLf & "            WHEN"
'    strSQL = strSQL & vbCrLf & "                        CBOXM_CTYTO = 20 --20:ガードマンボックス"
'    strSQL = strSQL & vbCrLf & "                THEN"
'    strSQL = strSQL & vbCrLf & "                    5   --ガードマンボックス"
'    strSQL = strSQL & vbCrLf & "            WHEN"
'    strSQL = strSQL & vbCrLf & "                    CBOXM_CTYTO = 1     --1:荷物収納"
'    strSQL = strSQL & vbCrLf & "                AND CBOXM_CTBNR = 40    --コンテナ分類(267)、40:市販物置"
'    strSQL = strSQL & vbCrLf & "                THEN"
'    strSQL = strSQL & vbCrLf & "                    6  --市販物置"
'    'INSERT 2020/06/09 Y.WADA End
'    strSQL = strSQL & vbCrLf & "            ELSE"
'    strSQL = strSQL & vbCrLf & "                    99  --その他"
'    strSQL = strSQL & vbCrLf & "            END AS [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "        ,   CBOXM_SIZE          AS [サイズ]"
'    strSQL = strSQL & vbCrLf & "        ,   CBOXM_DOOR          AS [ドア数]"
'    strSQL = strSQL & vbCrLf & "        ,   1                   AS [本数]"
'    strSQL = strSQL & vbCrLf & "        ,   CBOXM_CTBNR         AS [コンテナ分類]"
'    strSQL = strSQL & vbCrLf & "    FROM"
''    strSQL = strSQL & vbCrLf & "        CBYS_TRAN" 'DELETE 2020/02/15 Y.WADA
'    strSQL = strSQL & vbCrLf & "        CBYS_TRAN2" 'INSERT 2020/02/15 Y.WADA
'    strSQL = strSQL & vbCrLf & "    INNER JOIN"
'    strSQL = strSQL & vbCrLf & "        CBOX_MAST"
'    strSQL = strSQL & vbCrLf & "    ON"
'    strSQL = strSQL & vbCrLf & "        CBOXM_BOXNO = CBYST_BOXNO"
'    strSQL = strSQL & vbCrLf & "    AND"
'    strSQL = strSQL & vbCrLf & "        CBOXM_EDANO = CBYST_EDANO"
'    strSQL = strSQL & vbCrLf & "    AND"
'    strSQL = strSQL & vbCrLf & "        CBOXM_KRKBN = 1         --1:自社"
'    strSQL = strSQL & vbCrLf & "    INNER JOIN"
'    strSQL = strSQL & vbCrLf & "        YARD_MAST"
'    strSQL = strSQL & vbCrLf & "    ON"
'    strSQL = strSQL & vbCrLf & "        YARD_CODE = CBYST_YCODE"
'    strSQL = strSQL & vbCrLf & "    AND"
'    strSQL = strSQL & vbCrLf & "        LEFT(ISNULL(CONVERT(varchar(8),YARD_END_DAY,112),'99991231'),6) > LEFT(CBYST_SETDATE, 6)"
'    strSQL = strSQL & vbCrLf & "    INNER JOIN"
'    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.TINH_MAST"
'    strSQL = strSQL & vbCrLf & "    ON"
'    strSQL = strSQL & vbCrLf & "        TINHM_BUMOC = @S_BUMOC"
'    strSQL = strSQL & vbCrLf & "    AND"
'    strSQL = strSQL & vbCrLf & "        TINHM_TINTC = CBYST_YCODE"
''DELETE 2020/06/18 Y.WADA Start
''    strSQL = strSQL & vbCrLf & "    WHERE"
'''DELETE 2020/06/04 Y.WADA Start
'''    strSQL = strSQL & vbCrLf & "        CBYST_EDANO = '00'  --枝番"
'''    strSQL = strSQL & vbCrLf & "    AND"
'''DELETE 2020/06/04 Y.WADA End
''    strSQL = strSQL & vbCrLf & "        --CBYST_SETDATE BETWEEN @SETDATE_S AND @SETDATE_E"
''    strSQL = strSQL & vbCrLf & "        LEFT(CBYST_SETDATE, 6) = @SETDATE_YM"
''    strSQL = strSQL & vbCrLf & "    AND"
''    strSQL = strSQL & vbCrLf & "        CBYST_YCODE NOT IN ('800011', '800014') --800011:販売済みコンテナ、800014:梶山コンテナ置き場"
''DELETE 2020/06/18 Y.WADA End
'    strSQL = strSQL & vbCrLf & ")"
''    strSQL = strSQL & vbCrLf & "--, [出港2] AS"
''    strSQL = strSQL & vbCrLf & "--("
''    strSQL = strSQL & vbCrLf & "--    SELECT "
''    strSQL = strSQL & vbCrLf & "--        [列名]"
''    strSQL = strSQL & vbCrLf & "--    ,   [年月]"
''    strSQL = strSQL & vbCrLf & "--    ,   [仕入商品区分]"
''    strSQL = strSQL & vbCrLf & "--    ,   NAME_VALUE_FROM AS [在庫種類]"
''    strSQL = strSQL & vbCrLf & "--    ,   NAME_NAME  AS [商品名]"
''    strSQL = strSQL & vbCrLf & "--    ,   IIF(NAME_VALUE_FROM = 3, NULL, [サイズ]) AS [サイズ] "
''    strSQL = strSQL & vbCrLf & "--    ,   IIF(NAME_VALUE_FROM = 3, NULL, [ドア数]) AS [ドア数] "
''    strSQL = strSQL & vbCrLf & "--    ,   [本数]"
''    strSQL = strSQL & vbCrLf & "--    FROM [出港1]"
''    strSQL = strSQL & vbCrLf & "--    LEFT JOIN NAME_MAST"
''    strSQL = strSQL & vbCrLf & "--        ON  NAME_ID =451"
''    strSQL = strSQL & vbCrLf & "--        AND NAME_CODE = [仕入商品区分]"
''    strSQL = strSQL & vbCrLf & "--)"
'    strSQL = strSQL & vbCrLf & ", [出港] AS"
'    strSQL = strSQL & vbCrLf & "("
'    strSQL = strSQL & vbCrLf & "    SELECT "
'    strSQL = strSQL & vbCrLf & "        [列名]"
'    strSQL = strSQL & vbCrLf & "    ,   [年月]"
'    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
'    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
'    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
'    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
'    strSQL = strSQL & vbCrLf & "    ,   SUM([本数]) [本数]"
'    strSQL = strSQL & vbCrLf & "    FROM [出港1]"
'    strSQL = strSQL & vbCrLf & "    GROUP BY "
'    strSQL = strSQL & vbCrLf & "        [列名]"
'    strSQL = strSQL & vbCrLf & "    ,   [年月]"
'    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
'    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
'    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
'    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
'    strSQL = strSQL & vbCrLf & ")"
'    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
'    strSQL = strSQL & vbCrLf & "--(A)海外・前月在庫数"
'    strSQL = strSQL & vbCrLf & "--(B)梶山・前月在庫数"
'    strSQL = strSQL & vbCrLf & "--(C)営業・前月在庫数"
'    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
'    strSQL = strSQL & vbCrLf & ", [在庫] AS"
'    strSQL = strSQL & vbCrLf & "("
'    strSQL = strSQL & vbCrLf & "    SELECT"
'    strSQL = strSQL & vbCrLf & "        [列名]"
'    strSQL = strSQL & vbCrLf & "    ,   [年月]"
'    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "    ,   [サイズ]"
'    strSQL = strSQL & vbCrLf & "    ,   [ドア数]"
'    strSQL = strSQL & vbCrLf & "    ,   [本数]"
'    strSQL = strSQL & vbCrLf & "    FROM"
'    strSQL = strSQL & vbCrLf & "        ("
'    strSQL = strSQL & vbCrLf & "            SELECT"
'    strSQL = strSQL & vbCrLf & "                @SETDATE_YM     AS [年月]"
'    strSQL = strSQL & vbCrLf & "            ,   ZAIKS_SISYI     AS [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & "            ,   ZAIKS_SIZE      AS [サイズ]"
'    strSQL = strSQL & vbCrLf & "            ,   ZAIKS_DOOR      AS [ドア数]"
'    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZKAIQ, 0) + ISNULL(ZAIKS_HACYQ, 0) - ISNULL(ZAIKS_SYUKQ, 0) AS NUMERIC)    AS [(A)海外・前月在庫数]"
''DELETE 2020/06/09 Y.WADA Start
''    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZKAJQ, 0) + ISNULL(ZAIKS_SYUKQ, 0) + ISNULL(ZAIKS_TEKYQ, 0) - ISNULL(ZAIKS_BAIKQ, 0) - ISNULL(ZAIKS_ESYKQ, 0) AS NUMERIC)  AS [(B)梶山・前月在庫数]"
''    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZEIGQ, 0) + ISNULL(ZAIKS_ESYKQ, 0) - ISNULL(ZAIKS_TEKYQ, 0) AS NUMERIC)     AS [(C)営業・前月在庫数]"
''DELETE 2020/06/09 Y.WADA End
'    'INSERT 2020/06/09 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZKAJQ, 0) + ISNULL(ZAIKS_KAINQ, 0) + ISNULL(ZAIKS_TEKYQ, 0) - ISNULL(ZAIKS_BAIKQ, 0) - ISNULL(ZAIKS_ESYKQ, 0) AS NUMERIC)  AS [(B)梶山・前月在庫数]"
'    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZEIGQ, 0) + ISNULL(ZAIKS_KAJNQ, 0) - ISNULL(ZAIKS_TEKYQ, 0) AS NUMERIC)     AS [(C)営業・前月在庫数]"
'    'INSERT 2020/06/09 Y.WADA End
'    strSQL = strSQL & vbCrLf & "            FROM    ZAIK_SUMM"
'    strSQL = strSQL & vbCrLf & "            WHERE   ZAIKS_ZAISD             =   LEFT(CONVERT(VARCHAR, DATEADD(m, -1, CONVERT(DATETIME, @SETDATE_YM + '01')), 112), 6)"
'    strSQL = strSQL & vbCrLf & "        ) a"
'    strSQL = strSQL & vbCrLf & "    UNPIVOT"
'    strSQL = strSQL & vbCrLf & "        ("
'    strSQL = strSQL & vbCrLf & "            [本数] FOR [列名] IN ([(A)海外・前月在庫数], [(B)梶山・前月在庫数], [(C)営業・前月在庫数])"
'    strSQL = strSQL & vbCrLf & "        ) b"
'    strSQL = strSQL & vbCrLf & ")"
'    strSQL = strSQL & vbCrLf & ",   [入出在] AS"
'    strSQL = strSQL & vbCrLf & "("
'    strSQL = strSQL & vbCrLf & "                select * from [海外発注]"
'    strSQL = strSQL & vbCrLf & "    union all   select * from [出港]"
'    strSQL = strSQL & vbCrLf & "    union all   select * from [在庫]"
'    strSQL = strSQL & vbCrLf & ")"
'    strSQL = strSQL & vbCrLf & "INSERT"
'    strSQL = strSQL & vbCrLf & "INTO"
'    strSQL = strSQL & vbCrLf & "    ZAIK_SUMM"
'    strSQL = strSQL & vbCrLf & "("
'    strSQL = strSQL & vbCrLf & "    ZAIKS_ZAISD -- 在庫集計年月"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_SISYI -- 仕入商品区分"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_SIZE  -- サイズ"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_DOOR  -- ドア数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZKAIQ -- 海外前月在庫数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZKAJQ -- 梶山在庫前月在庫数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZEIGQ -- 営業ヤード前月在庫数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_HACYQ -- 発注数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_SYUKQ -- 出港数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_KAINQ -- 海外からの入庫数"        'INSERT 2020/06/09 Y.WADA
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_TEKYQ -- 営業ヤードからの撤去数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_BAIKQ -- 売却数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_ESYKQ -- 営業ヤードへの出庫数"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_KAJNQ -- 梶山ヤードからの入庫数"  'INSERT 2020/06/09 Y.WADA
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSED -- 作成日付"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSEJ -- 作成時刻"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSPB -- 作成プログラムＩＤ"
'    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSUB -- 作成ユーザーＩＤ"
'    strSQL = strSQL & vbCrLf & ")"
'    strSQL = strSQL & vbCrLf & "SELECT"
'    strSQL = strSQL & vbCrLf & "    [年月]"
'    strSQL = strSQL & vbCrLf & ",   [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & ",   [サイズ]"
'    strSQL = strSQL & vbCrLf & ",   [ドア数]"
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(A)海外・前月在庫数', [本数], NULL)) AS ZKAIQ"
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(B)梶山・前月在庫数', [本数], NULL)) AS ZKAJQ"
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(C)営業・前月在庫数', [本数], NULL)) AS ZEIGQ"
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '①発注数', [本数], NULL)) AS HACYQ"
' '   strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '②海外・出港数', [本数], NULL)) AS SYUKQ"    'DELETE 2020/06/09 Y.WADA
'    'INSERT 2020/06/09 Y.WADA Start
''    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] IN ('②海外・出港数', '②海外・出港数（直行）'), [本数], NULL)) AS SYUKQ --②"  'DELETE 2020/06/12 Y.WADA
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] IN ('②海外・出港数', '②海外・出港数（直送）'), [本数], NULL)) AS SYUKQ --②"
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '②海外・出港数', [本数], NULL)) AS KAINQ  --④"
'    'INSERT 2020/06/09 Y.WADA End
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑤梶山・撤去数', [本数], NULL)) AS TEKYQ"
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑥梶山・売却数', [本数], NULL)) AS BAIKQ"
''    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑦梶山・営業ヤードへの出庫数', [本数], NULL)) AS ESYKQ"
''    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] IN ('⑦梶山・営業ヤードへの出庫数', '②海外・出港数（直行）'), [本数], NULL)) AS KAJNQ   --⑨"   'INSERT 2020/06/09 Y.WADA  'DELETE 2020/06/12 Y.WADA
''DELETE 2020/06/18 Y.WADA Start
''    strSQL = strSQL & vbCrLf & ",   NULLIF(ISNULL(SUM(IIF([列名] = '⑦梶山・営業ヤードへの出庫数', [本数], NULL)),0) - ISNULL(SUM(IIF([列名] = '②海外・出港数（直送）', [本数], NULL)), 0),0) AS ESYKQ   --⑦" 'INSERT 2020/06/12 Y.WADA
''    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑦梶山・営業ヤードへの出庫数', [本数], NULL)) AS KAJNQ   --⑨"    'INSERT 2020/06/12 Y.WADA
''DELETE 2020/06/18 Y.WADA End
'    'INSERT 2020/06/18 Y.WADA Start
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑦梶山・営業ヤードへの出庫数',    [本数], NULL)) AS ESYKQ   --⑦"
'    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑨梶山ヤードからの入庫数',        [本数], NULL)) AS KAJNQ   --⑨"
'    'INSERT 2020/06/18 Y.WADA End
'    strSQL = strSQL & vbCrLf & ",   @wk_INSED   AS INSED"
'    strSQL = strSQL & vbCrLf & ",   @wk_INSEJ   AS INSEJ"
'    strSQL = strSQL & vbCrLf & ",   @wk_INSPB   AS INSPB"
'    strSQL = strSQL & vbCrLf & ",   @wk_INSUB   AS INSUB"
'    strSQL = strSQL & vbCrLf & "FROM"
'    strSQL = strSQL & vbCrLf & "    [入出在]"
'    strSQL = strSQL & vbCrLf & "GROUP BY"
'    strSQL = strSQL & vbCrLf & "    [年月]"
'    strSQL = strSQL & vbCrLf & ",   [仕入商品区分]"
'    strSQL = strSQL & vbCrLf & ",   [サイズ]"
'    strSQL = strSQL & vbCrLf & ",   [ドア数]"
'    strSQL = strSQL & vbCrLf & ";"
'    strSQL = strSQL & vbCrLf & "★"
'DELETE 2020/07/07 Y.WADA End
    'INSERT 2020/07/07 Y.WADA Start
    strSQL = ""
    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSED       varchar(8)  =   '" & strDate & "';"
    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSEJ       varchar(6)  =   '" & strTime & "';"
    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSPB       varchar(11) =   '" & strProg & "';"
    strSQL = strSQL & vbCrLf & "DECLARE @wk_INSUB       varchar(8)  =   '" & strUser & "';"
    strSQL = strSQL & vbCrLf & "DECLARE @SETDATE_YM CHAR(6) = '" & strSETDATE_YM & "';    --設置日（年月）"
    strSQL = strSQL & vbCrLf & "DECLARE @H_BUMOC    CHAR(1) = 'ﾗ';         --発注・部門コード"
    strSQL = strSQL & vbCrLf & "DECLARE @S_BUMOC    CHAR(1) = 'H';         --出港・部門コード;"
    strSQL = strSQL & vbCrLf & "--------------------------------"
    strSQL = strSQL & vbCrLf & "-- 在庫集計トランの登録"
    strSQL = strSQL & vbCrLf & "    WITH CBYS_TRAN2_tmp1 AS"
    strSQL = strSQL & vbCrLf & "    ("
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "            *"
    strSQL = strSQL & vbCrLf & "            ,   CASE"
    strSQL = strSQL & vbCrLf & "                WHEN"
    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800012')   --800012:輸入済みコンテナ"
    strSQL = strSQL & vbCrLf & "                    THEN"
    strSQL = strSQL & vbCrLf & "                        '②海外・出港数'"
    strSQL = strSQL & vbCrLf & "                WHEN"
    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800017')   --800017：直送輸入済みコンテナ"
    strSQL = strSQL & vbCrLf & "                    THEN"
    strSQL = strSQL & vbCrLf & "                        '②海外・出港数（直送）＋⑮港保管'"
    strSQL = strSQL & vbCrLf & "                WHEN"
    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800007','800008','800009','800010','800013')    --800007:廃棄ヤード　800008:撤去済みヤード　800009:解約ヤード　800010:販売ヤード　800013:入替コンテナ"
    strSQL = strSQL & vbCrLf & "                    THEN"
    strSQL = strSQL & vbCrLf & "                        '⑤梶山・撤去数'"
    strSQL = strSQL & vbCrLf & "                WHEN"
    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800011')   --800011:販売済ヤード"
    strSQL = strSQL & vbCrLf & "                    THEN"
    strSQL = strSQL & vbCrLf & "                        '⑥梶山・売却数'"
    strSQL = strSQL & vbCrLf & "                WHEN"
    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800018')   --800018:コンテナ不要在庫"
    strSQL = strSQL & vbCrLf & "                    THEN"
    strSQL = strSQL & vbCrLf & "                        '⑯不足在庫'"
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & "                WHEN"
    strSQL = strSQL & vbCrLf & "                        CBYST_YCODE IN ('800019')   --800019:中古購入ヤード"
    strSQL = strSQL & vbCrLf & "                    THEN"
    strSQL = strSQL & vbCrLf & "                        '21中古購入・出庫数'"
'INSERT 2022/09/30 K.KINEBUCHI end
    strSQL = strSQL & vbCrLf & "                ELSE"
    strSQL = strSQL & vbCrLf & "                        --'⑦梶山・営業ヤードへの出庫数'"
    strSQL = strSQL & vbCrLf & "                        '⑨梶山ヤードからの入庫数'"
    strSQL = strSQL & vbCrLf & "                END AS [列名]"
    strSQL = strSQL & vbCrLf & "        FROM ("
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "        1 as tbl"
    strSQL = strSQL & vbCrLf & "            ,CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "            ,CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "            ,CBYST_SYUBETSU"
    strSQL = strSQL & vbCrLf & "            ,CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "        --    ,CBYST_STATUS"
    strSQL = strSQL & vbCrLf & "            ,CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "            ,CBYST_SETFLOOR"
    strSQL = strSQL & vbCrLf & "            ,CBYST_BIKON"
    strSQL = strSQL & vbCrLf & "        FROM"
    strSQL = strSQL & vbCrLf & "            CBYS_TRAN"
    strSQL = strSQL & vbCrLf & "        WHERE"
    strSQL = strSQL & vbCrLf & "            LEFT(CBYST_SETDATE, 6) = @SETDATE_YM"
    strSQL = strSQL & vbCrLf & "        UNION ALL"
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "        2 as tbl"
    strSQL = strSQL & vbCrLf & "            ,CBYSR_BOXNO"
    strSQL = strSQL & vbCrLf & "            ,CBYSR_EDANO"
    strSQL = strSQL & vbCrLf & "            ,CBYSR_SYUBETSU"
    strSQL = strSQL & vbCrLf & "            ,CBYSR_YCODE"
    strSQL = strSQL & vbCrLf & "            ,CBYSR_SETDATE"
    strSQL = strSQL & vbCrLf & "            ,CBYSR_SETFLOOR"
    strSQL = strSQL & vbCrLf & "            ,CBYSR_BIKON"
    strSQL = strSQL & vbCrLf & "        FROM"
    strSQL = strSQL & vbCrLf & "            CBYS_RECD"
    strSQL = strSQL & vbCrLf & "        WHERE"
    strSQL = strSQL & vbCrLf & "            LEFT(CBYSR_SETDATE, 6) = @SETDATE_YM"
    strSQL = strSQL & vbCrLf & "    ) a"
    strSQL = strSQL & vbCrLf & "    )"
    strSQL = strSQL & vbCrLf & "    , CBYS_TRAN2_tmp2 AS"
    strSQL = strSQL & vbCrLf & "    ("
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "            ROW_NUMBER()"
    strSQL = strSQL & vbCrLf & "                OVER "
    strSQL = strSQL & vbCrLf & "                ("
    strSQL = strSQL & vbCrLf & "                    PARTITION BY"
    strSQL = strSQL & vbCrLf & "                        [列名]"
    strSQL = strSQL & vbCrLf & "                    ,   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "                    ,   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "                    ORDER BY CBYST_SETDATE DESC, tbl"
    strSQL = strSQL & vbCrLf & "                ) AS rno"
    strSQL = strSQL & vbCrLf & "            , *"
    strSQL = strSQL & vbCrLf & "        FROM"
    strSQL = strSQL & vbCrLf & "            CBYS_TRAN2_tmp1"
    strSQL = strSQL & vbCrLf & "        WHERE"
    strSQL = strSQL & vbCrLf & "            CBYST_YCODE NOT IN ('800014') --800014:梶山コンテナ置き場"
    strSQL = strSQL & vbCrLf & "    )"
    strSQL = strSQL & vbCrLf & "    , CBYS_TRAN2_tmp3 AS"
    strSQL = strSQL & vbCrLf & "    ("
    strSQL = strSQL & vbCrLf & "        SELECT * FROM CBYS_TRAN2_tmp2"
    strSQL = strSQL & vbCrLf & "        WHERE rno = 1"
    strSQL = strSQL & vbCrLf & "    )"
    strSQL = strSQL & vbCrLf & "    , CBYS_TRAN2 AS"
    strSQL = strSQL & vbCrLf & "    ("
    strSQL = strSQL & vbCrLf & "        SELECT * FROM CBYS_TRAN2_tmp3"
    strSQL = strSQL & vbCrLf & "        UNION ALL"
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "            rno"
    strSQL = strSQL & vbCrLf & "        ,   tbl"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SYUBETSU"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETFLOOR"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BIKON"
    strSQL = strSQL & vbCrLf & "        ,   IIF"
    strSQL = strSQL & vbCrLf & "                ("
    strSQL = strSQL & vbCrLf & "                    ("
    strSQL = strSQL & vbCrLf & "                        --直送なら1"
    strSQL = strSQL & vbCrLf & "                        SELECT COUNT(*)"
    strSQL = strSQL & vbCrLf & "                        FROM"
    strSQL = strSQL & vbCrLf & "                            ("
    strSQL = strSQL & vbCrLf & "                                SELECT  TOP 1"
    strSQL = strSQL & vbCrLf & "                                    CBYSR_YCODE"
    strSQL = strSQL & vbCrLf & "                                FROM"
    strSQL = strSQL & vbCrLf & "                                    CBYS_RECD"
    strSQL = strSQL & vbCrLf & "                                WHERE"
    strSQL = strSQL & vbCrLf & "                                    CBYSR_SETDATE   <=  CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "                                AND CBYSR_BOXNO     =   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "                                AND CBYSR_EDANO     =   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "                                AND CBYSR_YCODE     <>  CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "                                ORDER BY CBYSR_SETDATE DESC"
    strSQL = strSQL & vbCrLf & "                            ) a"
    strSQL = strSQL & vbCrLf & "                        WHERE"
    strSQL = strSQL & vbCrLf & "                            CBYSR_YCODE = '800017'  --800017：直送輸入済みコンテナ"
    strSQL = strSQL & vbCrLf & "                    ) > 0"
'    strSql = strSql & vbCrLf & "                ,   '②海外・出港数（直送）'"      'DELETE 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & "                ,   '⑱港保管出庫数'"               'INSERT 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & "                ,   '⑦梶山・営業ヤードへの出庫数'"
    strSQL = strSQL & vbCrLf & "                ) AS [列名]"
    strSQL = strSQL & vbCrLf & "        FROM"
    strSQL = strSQL & vbCrLf & "            CBYS_TRAN2_tmp3"
    strSQL = strSQL & vbCrLf & "        WHERE"
    strSQL = strSQL & vbCrLf & "            [列名]  =   '⑨梶山ヤードからの入庫数'"
    strSQL = strSQL & vbCrLf & "        UNION ALL"
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "            rno"
    strSQL = strSQL & vbCrLf & "        ,   tbl"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SYUBETSU"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETFLOOR"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BIKON"
    strSQL = strSQL & vbCrLf & "        ,   '⑮港保管' AS [列名]"
    strSQL = strSQL & vbCrLf & "        FROM"
    strSQL = strSQL & vbCrLf & "            CBYS_TRAN2_tmp3"
    strSQL = strSQL & vbCrLf & "        WHERE"
    strSQL = strSQL & vbCrLf & "            [列名]  =   '②海外・出港数（直送）＋⑮港保管'"

'DELETE 2020/07/22 Y.WADA Start
'    strSQL = strSQL & vbCrLf & "        AND EXISTS --現在地が自分"
'    strSQL = strSQL & vbCrLf & "            ("
'    strSQL = strSQL & vbCrLf & "                SELECT *"
'    strSQL = strSQL & vbCrLf & "                FROM"
'    strSQL = strSQL & vbCrLf & "                    ("
'    strSQL = strSQL & vbCrLf & "                        SELECT  TOP 1"
'    strSQL = strSQL & vbCrLf & "                            CBYSR_YCODE"
'    strSQL = strSQL & vbCrLf & "                        FROM"
'    strSQL = strSQL & vbCrLf & "                            CBYS_RECD"
'    strSQL = strSQL & vbCrLf & "                        WHERE"
'    strSQL = strSQL & vbCrLf & "                            CBYSR_SETDATE   <=  CBYST_SETDATE"
'    strSQL = strSQL & vbCrLf & "                        AND CBYSR_BOXNO     =   CBYST_BOXNO"
'    strSQL = strSQL & vbCrLf & "                        AND CBYSR_EDANO     =   CBYST_EDANO"
'    strSQL = strSQL & vbCrLf & "                        --AND CBYSR_YCODE     <>  CBYST_YCODE"
'    strSQL = strSQL & vbCrLf & "                        ORDER BY CBYSR_SETDATE DESC"
'    strSQL = strSQL & vbCrLf & "                    ) a"
'    strSQL = strSQL & vbCrLf & "                WHERE"
'    strSQL = strSQL & vbCrLf & "                    CBYSR_YCODE = CBYST_YCODE"
'    strSQL = strSQL & vbCrLf & "            )"
'DELETE 2020/07/22 Y.WADA End
    
    'INSERT 2020/07/22 Y.WADA Start
    strSQL = strSQL & vbCrLf & "        AND EXISTS --現在地が自分"
    strSQL = strSQL & vbCrLf & "            ("
    strSQL = strSQL & vbCrLf & "                SELECT *"
    strSQL = strSQL & vbCrLf & "                FROM"
    strSQL = strSQL & vbCrLf & "                    ("
    strSQL = strSQL & vbCrLf & "                        SELECT  TOP 1"
    strSQL = strSQL & vbCrLf & "                            YCODE"
    strSQL = strSQL & vbCrLf & "                        FROM"
    strSQL = strSQL & vbCrLf & "                            ("
    strSQL = strSQL & vbCrLf & "                               SELECT"
    strSQL = strSQL & vbCrLf & "                                1 as tbl"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_BOXNO    AS BOXNO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_EDANO    AS EDANO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SYUBETSU AS SYUBETSU"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_YCODE    AS YCODE"
    strSQL = strSQL & vbCrLf & "                                --    ,CBYST_STATUS"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SETDATE  AS SETDATE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SETFLOOR AS SETFLOOR"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_BIKON    AS BIKON"
    strSQL = strSQL & vbCrLf & "                                FROM"
    strSQL = strSQL & vbCrLf & "                                    CBYS_TRAN"
    strSQL = strSQL & vbCrLf & "                                UNION ALL"
    strSQL = strSQL & vbCrLf & "                                SELECT"
    strSQL = strSQL & vbCrLf & "                                2 as tbl"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_BOXNO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_EDANO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SYUBETSU"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_YCODE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SETDATE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SETFLOOR"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_BIKON"
    strSQL = strSQL & vbCrLf & "                                FROM"
    strSQL = strSQL & vbCrLf & "                                    CBYS_RECD"
    strSQL = strSQL & vbCrLf & "                            ) tbl2"
    strSQL = strSQL & vbCrLf & "                        WHERE"
    strSQL = strSQL & vbCrLf & "                            SETDATE   <=  CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "                        AND BOXNO     =   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "                        AND EDANO     =   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "                        --AND CBYSR_YCODE     <>  CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "                        ORDER BY SETDATE DESC"
    strSQL = strSQL & vbCrLf & "                    ) a"
    strSQL = strSQL & vbCrLf & "                WHERE"
    strSQL = strSQL & vbCrLf & "                    YCODE = CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "            )"
    'INSERT 2020/07/22 Y.WADA End
    
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & "        UNION ALL"
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "            rno"
    strSQL = strSQL & vbCrLf & "        ,   tbl"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SYUBETSU"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETFLOOR"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BIKON"
    strSQL = strSQL & vbCrLf & "        ,   '⑰港保管入庫数' AS [列名]"
    strSQL = strSQL & vbCrLf & "        FROM"
    strSQL = strSQL & vbCrLf & "            CBYS_TRAN2_tmp3"
    strSQL = strSQL & vbCrLf & "        WHERE"
    strSQL = strSQL & vbCrLf & "            [列名]  =   '②海外・出港数（直送）＋⑮港保管'"
    strSQL = strSQL & vbCrLf & "        AND EXISTS --現在地が自分"
    strSQL = strSQL & vbCrLf & "            ("
    strSQL = strSQL & vbCrLf & "                SELECT *"
    strSQL = strSQL & vbCrLf & "                FROM"
    strSQL = strSQL & vbCrLf & "                    ("
    strSQL = strSQL & vbCrLf & "                        SELECT  TOP 1"
    strSQL = strSQL & vbCrLf & "                            YCODE"
    strSQL = strSQL & vbCrLf & "                        FROM"
    strSQL = strSQL & vbCrLf & "                            ("
    strSQL = strSQL & vbCrLf & "                               SELECT"
    strSQL = strSQL & vbCrLf & "                                1 as tbl"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_BOXNO    AS BOXNO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_EDANO    AS EDANO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SYUBETSU AS SYUBETSU"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_YCODE    AS YCODE"
    strSQL = strSQL & vbCrLf & "                                --    ,CBYST_STATUS"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SETDATE  AS SETDATE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SETFLOOR AS SETFLOOR"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_BIKON    AS BIKON"
    strSQL = strSQL & vbCrLf & "                                FROM"
    strSQL = strSQL & vbCrLf & "                                    CBYS_TRAN"
    strSQL = strSQL & vbCrLf & "                                UNION ALL"
    strSQL = strSQL & vbCrLf & "                                SELECT"
    strSQL = strSQL & vbCrLf & "                                2 as tbl"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_BOXNO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_EDANO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SYUBETSU"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_YCODE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SETDATE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SETFLOOR"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_BIKON"
    strSQL = strSQL & vbCrLf & "                                FROM"
    strSQL = strSQL & vbCrLf & "                                    CBYS_RECD"
    strSQL = strSQL & vbCrLf & "                            ) tbl2"
    strSQL = strSQL & vbCrLf & "                        WHERE"
    strSQL = strSQL & vbCrLf & "                            SETDATE   <=  CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "                        AND BOXNO     =   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "                        AND EDANO     =   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "                        --AND CBYSR_YCODE     <>  CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "                        ORDER BY SETDATE DESC"
    strSQL = strSQL & vbCrLf & "                    ) a"
    strSQL = strSQL & vbCrLf & "                WHERE"
    strSQL = strSQL & vbCrLf & "                    YCODE = CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "            )"
    strSQL = strSQL & vbCrLf & "        UNION ALL"
    strSQL = strSQL & vbCrLf & "        SELECT"
    strSQL = strSQL & vbCrLf & "            rno"
    strSQL = strSQL & vbCrLf & "        ,   tbl"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SYUBETSU"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_SETFLOOR"
    strSQL = strSQL & vbCrLf & "        ,   CBYST_BIKON"
    strSQL = strSQL & vbCrLf & "        ,   '21中古購入から梶山ヤードへ出庫' AS [列名]"
    strSQL = strSQL & vbCrLf & "        FROM"
    strSQL = strSQL & vbCrLf & "            CBYS_TRAN2_tmp3"
    strSQL = strSQL & vbCrLf & "        WHERE"
    strSQL = strSQL & vbCrLf & "            [列名]  =   '21中古購入・出庫数'"
    strSQL = strSQL & vbCrLf & "        AND EXISTS --現在地が自分"
    strSQL = strSQL & vbCrLf & "            ("
    strSQL = strSQL & vbCrLf & "                SELECT *"
    strSQL = strSQL & vbCrLf & "                FROM"
    strSQL = strSQL & vbCrLf & "                    ("
    strSQL = strSQL & vbCrLf & "                        SELECT  TOP 1"
    strSQL = strSQL & vbCrLf & "                            YCODE"
    strSQL = strSQL & vbCrLf & "                        FROM"
    strSQL = strSQL & vbCrLf & "                            ("
    strSQL = strSQL & vbCrLf & "                               SELECT"
    strSQL = strSQL & vbCrLf & "                                1 as tbl"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_BOXNO    AS BOXNO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_EDANO    AS EDANO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SYUBETSU AS SYUBETSU"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_YCODE    AS YCODE"
    strSQL = strSQL & vbCrLf & "                                --    ,CBYST_STATUS"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SETDATE  AS SETDATE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_SETFLOOR AS SETFLOOR"
    strSQL = strSQL & vbCrLf & "                                    ,CBYST_BIKON    AS BIKON"
    strSQL = strSQL & vbCrLf & "                                FROM"
    strSQL = strSQL & vbCrLf & "                                    CBYS_TRAN"
    strSQL = strSQL & vbCrLf & "                                UNION ALL"
    strSQL = strSQL & vbCrLf & "                                SELECT"
    strSQL = strSQL & vbCrLf & "                                2 as tbl"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_BOXNO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_EDANO"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SYUBETSU"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_YCODE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SETDATE"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_SETFLOOR"
    strSQL = strSQL & vbCrLf & "                                    ,CBYSR_BIKON"
    strSQL = strSQL & vbCrLf & "                                FROM"
    strSQL = strSQL & vbCrLf & "                                    CBYS_RECD"
    strSQL = strSQL & vbCrLf & "                            ) tbl2"
    strSQL = strSQL & vbCrLf & "                        WHERE"
    strSQL = strSQL & vbCrLf & "                            SETDATE   <=  CBYST_SETDATE"
    strSQL = strSQL & vbCrLf & "                        AND BOXNO     =   CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "                        AND EDANO     =   CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "                        --AND CBYSR_YCODE     <>  CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "                        ORDER BY SETDATE DESC"
    strSQL = strSQL & vbCrLf & "                    ) a"
    strSQL = strSQL & vbCrLf & "                WHERE"
    strSQL = strSQL & vbCrLf & "                    YCODE = CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "            )"
'INSERT 2022/09/30 K.KINEBUCHI end
    
    strSQL = strSQL & vbCrLf & "    )"
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & "--①海外発注"
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & ",   [海外発注1] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT"
    strSQL = strSQL & vbCrLf & "        '①発注数'  [列名]"
    strSQL = strSQL & vbCrLf & "    ,   LEFT(MDOCT_TYUMD, 6)    AS [年月]"
    strSQL = strSQL & vbCrLf & "    ,   CASE"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                ISNULL(HANBM_DOORQ,0) != 0      --ドア数"
    strSQL = strSQL & vbCrLf & "            AND HANBM_SYUBI = 2                 --1:新型（JIS）　2:旧型（ISO）"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                1   --ISOコンテナ"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                ISNULL(HANBM_DOORQ,0) != 0      --ドア数"
    strSQL = strSQL & vbCrLf & "            AND HANBM_SYUBI = 1                 --1:新型（JIS）　2:旧型（ISO）"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                2   --JISコンテナ"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                HANBM_SYOHN LIKE '%バイク%'"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                3   --バイクボックス"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                    HANBM_SYOHN LIKE '%台車%'"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                4   --台車"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                    HANBM_SYOHN LIKE '%ガードマン%'"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                5   --ガードマンボックス"
    strSQL = strSQL & vbCrLf & "        ELSE"
    strSQL = strSQL & vbCrLf & "                99  --その他"
    strSQL = strSQL & vbCrLf & "        END AS [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    ,   HANBM_SYOHN         AS [商品名]     --商品名"
    strSQL = strSQL & vbCrLf & "    ,   HANBM_SIZEI         AS [サイズ]"
    strSQL = strSQL & vbCrLf & "    ,   HANBM_DOORQ         AS [ドア数]"
    strSQL = strSQL & vbCrLf & "    ,   HJUKT_HANBQ         AS [本数]"
    strSQL = strSQL & vbCrLf & "    ,   HJUKT_USEDI         AS [新品・中古区分]     --0:新品 1:中古"
    strSQL = strSQL & vbCrLf & "    "
    strSQL = strSQL & vbCrLf & "    FROM"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.KEIY_TRAN"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.HJUK_TRAN"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        HJUKT_KEIYB = KEIYT_KEIYB"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.MDOC_TRAN"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        MDOCT_KEIYB = KEIYT_KEIYB"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        --MDOCT_TYUMD BETWEEN　@SETDATE_S AND @SETDATE_E"
    strSQL = strSQL & vbCrLf & "        LEFT(MDOCT_TYUMD, 6) = @SETDATE_YM"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        MDOCT_TNAMN LIKE '%自社%'"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.HANB_MAST"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        HANBM_SYOHC = HJUKT_SYOHC"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        HANBM_SYOHI = 1                 --1:仕入商品　2:不可商品"
    strSQL = strSQL & vbCrLf & "    AND "
    strSQL = strSQL & vbCrLf & "        ISNULL(HANBM_SYUTI,0) = 0       --トランク用部品種別"
    strSQL = strSQL & vbCrLf & "    WHERE"
    strSQL = strSQL & vbCrLf & "        KEIYT_BUMOC = @H_BUMOC"
    strSQL = strSQL & vbCrLf & ")"
    strSQL = strSQL & vbCrLf & ", [海外発注] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT "
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "    ,   [年月]"
    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
    strSQL = strSQL & vbCrLf & "    ,   SUM([本数]) [本数]"
    strSQL = strSQL & vbCrLf & "    FROM [海外発注1]"
    strSQL = strSQL & vbCrLf & "    GROUP BY "
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "    ,   [年月]"
    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
    strSQL = strSQL & vbCrLf & ")"
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & "--②海外・出港数"
    strSQL = strSQL & vbCrLf & "--⑤梶山・撤去数"
    strSQL = strSQL & vbCrLf & "--⑥梶山・売却数"
    strSQL = strSQL & vbCrLf & "--⑦梶山・営業ヤードへの出庫数"
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & ", [出港1] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT"
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "        ,   LEFT(CBYST_SETDATE, 6)  AS [年月]"
    strSQL = strSQL & vbCrLf & "        ,   CASE"
    strSQL = strSQL & vbCrLf & "            WHEN"
    strSQL = strSQL & vbCrLf & "                    CBOXM_CTBNR = 1     --1:海上コンテナ　99:--"
    strSQL = strSQL & vbCrLf & "                AND CBOXM_CTYTO = 1     --1:荷物収納　2:バイク収納"
    strSQL = strSQL & vbCrLf & "                AND CBOXM_TYPE = 2      --1:新型（JIS）　2:旧型（ISO）"
    strSQL = strSQL & vbCrLf & "                THEN"
    strSQL = strSQL & vbCrLf & "                    1   --ISOコンテナ"
    strSQL = strSQL & vbCrLf & "            WHEN"
    strSQL = strSQL & vbCrLf & "                    CBOXM_CTBNR = 1     --1:海上コンテナ　99:--"
    strSQL = strSQL & vbCrLf & "                AND CBOXM_CTYTO = 1     --1:荷物収納　2:バイク収納"
    strSQL = strSQL & vbCrLf & "                AND CBOXM_TYPE = 1      --1:新型（JIS）　2:旧型（ISO）"
    strSQL = strSQL & vbCrLf & "                THEN"
    strSQL = strSQL & vbCrLf & "                    2   --JISコンテナ"
    strSQL = strSQL & vbCrLf & "            WHEN"
    strSQL = strSQL & vbCrLf & "                CBOXM_CTYTO = 2         --1:荷物収納　2:バイク収納"
    strSQL = strSQL & vbCrLf & "                THEN"
    strSQL = strSQL & vbCrLf & "                    3   --バイクボックス"
    strSQL = strSQL & vbCrLf & "            WHEN"
    strSQL = strSQL & vbCrLf & "                        CBOXM_CTYTO = 30 --30:台車"
    strSQL = strSQL & vbCrLf & "                THEN"
    strSQL = strSQL & vbCrLf & "                    4   --台車"
    strSQL = strSQL & vbCrLf & "            WHEN"
    strSQL = strSQL & vbCrLf & "                        CBOXM_CTYTO = 20 --20:ガードマンボックス"
    strSQL = strSQL & vbCrLf & "                THEN"
    strSQL = strSQL & vbCrLf & "                    5   --ガードマンボックス"
    strSQL = strSQL & vbCrLf & "            WHEN"
    strSQL = strSQL & vbCrLf & "                    CBOXM_CTYTO = 1     --1:荷物収納"
    strSQL = strSQL & vbCrLf & "                AND CBOXM_CTBNR = 40    --コンテナ分類(267)、40:市販物置"
    strSQL = strSQL & vbCrLf & "                THEN"
    strSQL = strSQL & vbCrLf & "                    6  --市販物置"
    strSQL = strSQL & vbCrLf & "            ELSE"
    strSQL = strSQL & vbCrLf & "                    99  --その他"
    strSQL = strSQL & vbCrLf & "            END AS [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "        ,   CBOXM_SIZE          AS [サイズ]"
    strSQL = strSQL & vbCrLf & "        ,   CBOXM_DOOR          AS [ドア数]"
    strSQL = strSQL & vbCrLf & "        ,   1                   AS [本数]"
    strSQL = strSQL & vbCrLf & "        ,   CBOXM_CTBNR         AS [コンテナ分類]"
    strSQL = strSQL & vbCrLf & "    FROM"
    strSQL = strSQL & vbCrLf & "        CBYS_TRAN2"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        CBOX_MAST"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        CBOXM_BOXNO = CBYST_BOXNO"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        CBOXM_EDANO = CBYST_EDANO"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        CBOXM_KRKBN = 1         --1:自社"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        YARD_MAST"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        YARD_CODE = CBYST_YCODE"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        LEFT(ISNULL(CONVERT(varchar(8),YARD_END_DAY,112),'99991231'),6) > LEFT(CBYST_SETDATE, 6)"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.TINH_MAST"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        TINHM_BUMOC = @S_BUMOC"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        TINHM_TINTC = CBYST_YCODE"
    strSQL = strSQL & vbCrLf & ")"
    strSQL = strSQL & vbCrLf & ", [出港] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT "
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "    ,   [年月]"
    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
    strSQL = strSQL & vbCrLf & "    ,   SUM([本数]) [本数]"
    strSQL = strSQL & vbCrLf & "    FROM [出港1]"
    strSQL = strSQL & vbCrLf & "    GROUP BY "
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "    ,   [年月]"
    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    --,   [在庫種類]"
    strSQL = strSQL & vbCrLf & "    --,   [商品名]"
    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
    strSQL = strSQL & vbCrLf & ")"
    
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & "--20中古購入コンテナ発注"
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & ",   [中古購入発注1] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT"
    strSQL = strSQL & vbCrLf & "        '20中古購入から自社へ発注'  [列名]"
    strSQL = strSQL & vbCrLf & "    ,   LEFT(MDOCT_TYUMD, 6)    AS [年月]"
    strSQL = strSQL & vbCrLf & "    ,   CASE"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                ISNULL(HANBM_DOORQ,0) != 0      --ドア数"
    strSQL = strSQL & vbCrLf & "            AND HANBM_SYUBI = 2                 --1:新型（JIS）　2:旧型（ISO）"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                1   --ISOコンテナ"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                ISNULL(HANBM_DOORQ,0) != 0      --ドア数"
    strSQL = strSQL & vbCrLf & "            AND HANBM_SYUBI = 1                 --1:新型（JIS）　2:旧型（ISO）"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                2   --JISコンテナ"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                HANBM_SYOHN LIKE '%バイク%'"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                3   --バイクボックス"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                    HANBM_SYOHN LIKE '%台車%'"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                4   --台車"
    strSQL = strSQL & vbCrLf & "        WHEN"
    strSQL = strSQL & vbCrLf & "                    HANBM_SYOHN LIKE '%ガードマン%'"
    strSQL = strSQL & vbCrLf & "            THEN"
    strSQL = strSQL & vbCrLf & "                5   --ガードマンボックス"
    strSQL = strSQL & vbCrLf & "        ELSE"
    strSQL = strSQL & vbCrLf & "                99  --その他"
    strSQL = strSQL & vbCrLf & "        END AS [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    ,   HANBM_SYOHN         AS [商品名]     --商品名"
    strSQL = strSQL & vbCrLf & "    ,   HANBM_SIZEI         AS [サイズ]"
    strSQL = strSQL & vbCrLf & "    ,   HANBM_DOORQ         AS [ドア数]"
    strSQL = strSQL & vbCrLf & "    ,   HJUKT_HANBQ         AS [本数]"
    strSQL = strSQL & vbCrLf & "    ,   HJUKT_USEDI         AS [新品・中古区分]     --0:新品 1:中古"
    strSQL = strSQL & vbCrLf & "    "
    strSQL = strSQL & vbCrLf & "    FROM"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.KEIY_TRAN"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.HJUK_TRAN"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        HJUKT_KEIYB = KEIYT_KEIYB"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        HJUKT_USEDI = 1                             --中古のみ"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.MDOC_TRAN"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        MDOCT_KEIYB = KEIYT_KEIYB"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        LEFT(MDOCT_TYUMD, 6) = @SETDATE_YM"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        MDOCT_TNAMN LIKE '%購入%'"
    strSQL = strSQL & vbCrLf & "    INNER JOIN"
    strSQL = strSQL & vbCrLf & "        " & strDataSource & ".DBO.HANB_MAST"
    strSQL = strSQL & vbCrLf & "    ON"
    strSQL = strSQL & vbCrLf & "        HANBM_SYOHC = HJUKT_SYOHC"
    strSQL = strSQL & vbCrLf & "    AND"
    strSQL = strSQL & vbCrLf & "        HANBM_SYOHI = 1                 --1:仕入商品　2:不可商品"
    strSQL = strSQL & vbCrLf & "    AND "
    strSQL = strSQL & vbCrLf & "        ISNULL(HANBM_SYUTI,0) = 0       --トランク用部品種別"
    strSQL = strSQL & vbCrLf & "    WHERE"
    strSQL = strSQL & vbCrLf & "        KEIYT_BUMOC = @S_BUMOC"
    strSQL = strSQL & vbCrLf & ")"
    strSQL = strSQL & vbCrLf & ", [中古購入発注] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT "
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "    ,   [年月]"
    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
    strSQL = strSQL & vbCrLf & "    ,   SUM([本数]) [本数]"
    strSQL = strSQL & vbCrLf & "    FROM [中古購入発注1]"
    strSQL = strSQL & vbCrLf & "    GROUP BY "
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "    ,   [年月]"
    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    ,   [サイズ] "
    strSQL = strSQL & vbCrLf & "    ,   [ドア数] "
    strSQL = strSQL & vbCrLf & ")"
'INSERT 2022/09/30 K.KINEBUCHI end
    
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & "--(A)海外・前月在庫数"
    strSQL = strSQL & vbCrLf & "--(B)梶山・前月在庫数"
    strSQL = strSQL & vbCrLf & "--(C)営業・前月在庫数"
    strSQL = strSQL & vbCrLf & "------------------------------------------------------"
    strSQL = strSQL & vbCrLf & ", [在庫] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT"
    strSQL = strSQL & vbCrLf & "        [列名]"
    strSQL = strSQL & vbCrLf & "    ,   [年月]"
    strSQL = strSQL & vbCrLf & "    ,   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "    ,   [サイズ]"
    strSQL = strSQL & vbCrLf & "    ,   [ドア数]"
    strSQL = strSQL & vbCrLf & "    ,   [本数]"
    strSQL = strSQL & vbCrLf & "    FROM"
    strSQL = strSQL & vbCrLf & "        ("
    strSQL = strSQL & vbCrLf & "            SELECT"
    strSQL = strSQL & vbCrLf & "                @SETDATE_YM     AS [年月]"
    strSQL = strSQL & vbCrLf & "            ,   ZAIKS_SISYI     AS [仕入商品区分]"
    strSQL = strSQL & vbCrLf & "            ,   ZAIKS_SIZE      AS [サイズ]"
    strSQL = strSQL & vbCrLf & "            ,   ZAIKS_DOOR      AS [ドア数]"
'    strSql = strSql & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZKAIQ, 0) + ISNULL(ZAIKS_HACYQ, 0) - ISNULL(ZAIKS_SYUKQ, 0) AS NUMERIC)    AS [(A)海外・前月在庫数]"                         'DELETE 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZKAIQ, 0) + ISNULL(ZAIKS_HACYQ, 0) - ISNULL(ZAIKS_SYUKQ, 0) - ISNULL(ZAIKS_MIHOQ, 0) AS NUMERIC)    AS [(A)海外・前月在庫数]" 'INSERT 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZKAJQ, 0) + ISNULL(ZAIKS_KAINQ, 0) + ISNULL(ZAIKS_TEKYQ, 0) - ISNULL(ZAIKS_BAIKQ, 0) - ISNULL(ZAIKS_ESYKQ, 0) AS NUMERIC)  AS [(B)梶山・前月在庫数]"
    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZEIGQ, 0) + ISNULL(ZAIKS_KAJNQ, 0) - ISNULL(ZAIKS_TEKYQ, 0) AS NUMERIC)     AS [(C)営業・前月在庫数]"
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZMIHQ, 0) + ISNULL(ZAIKS_MHNYQ, 0) - ISNULL(ZAIKS_MHSYQ, 0) AS NUMERIC)     AS [(E)港保管・前月在庫数]"
    strSQL = strSQL & vbCrLf & "            ,   CAST(ISNULL(ZAIKS_ZTASQ, 0) + ISNULL(ZAIKS_ZTHAQ, 0) - ISNULL(ZAIKS_ZTSKQ, 0) - ISNULL(ZAIKS_ZTSEQ, 0) AS NUMERIC)    AS [(F)中古購入・前月在庫数]"
'INSERT 2022/09/30 K.KINEBUCHI end
    strSQL = strSQL & vbCrLf & "            FROM    ZAIK_SUMM"
    strSQL = strSQL & vbCrLf & "            WHERE   ZAIKS_ZAISD             =   LEFT(CONVERT(VARCHAR, DATEADD(m, -1, CONVERT(DATETIME, @SETDATE_YM + '01')), 112), 6)"
    strSQL = strSQL & vbCrLf & "        ) a"
    strSQL = strSQL & vbCrLf & "    UNPIVOT"
    strSQL = strSQL & vbCrLf & "        ("
'    strSql = strSql & vbCrLf & "            [本数] FOR [列名] IN ([(A)海外・前月在庫数], [(B)梶山・前月在庫数], [(C)営業・前月在庫数])"                                                        'DELETE 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & "            [本数] FOR [列名] IN ([(A)海外・前月在庫数], [(B)梶山・前月在庫数], [(C)営業・前月在庫数], [(E)港保管・前月在庫数], [(F)中古購入・前月在庫数])"     'INSERT 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & "        ) b"
    strSQL = strSQL & vbCrLf & ")"
    strSQL = strSQL & vbCrLf & ",   [入出在] AS"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "                select * from [海外発注]"
    strSQL = strSQL & vbCrLf & "    union all   select * from [出港]"
    strSQL = strSQL & vbCrLf & "    union all   select * from [中古購入発注]"                                                                                                                   'INSERT 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & "    union all   select * from [在庫]"
    strSQL = strSQL & vbCrLf & ")"
    strSQL = strSQL & vbCrLf & "INSERT"
    strSQL = strSQL & vbCrLf & "INTO"
    strSQL = strSQL & vbCrLf & "    ZAIK_SUMM"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    ZAIKS_ZAISD -- 在庫集計年月"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_SISYI -- 仕入商品区分"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_SIZE  -- サイズ"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_DOOR  -- ドア数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZKAIQ -- 海外前月在庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZKAJQ -- 梶山在庫前月在庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZEIGQ -- 営業ヤード前月在庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_HACYQ -- 発注数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_SYUKQ -- 出港数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_KAINQ -- 海外からの入庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_TEKYQ -- 営業ヤードからの撤去数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_BAIKQ -- 売却数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ESYKQ -- 営業ヤードへの出庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_KAJNQ -- 梶山ヤードからの入庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_MIHOQ -- 港保管"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_FUSZQ -- 不足在庫"
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZMIHQ  -- 港保管前月在庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_MHNYQ  -- 港保管入庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_MHSYQ  -- 港保管出庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTASQ  -- 中古購入前月在庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTHAQ  -- 中古購入から自社へ発注数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTSKQ  -- 中古購入から梶山ヤードへ出庫数"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_ZTSEQ  -- 中古購入から営業ヤードへ出庫数"             '現在直送は無い
'INSERT 2022/09/30 K.KINEBUCHI end
    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSED -- 作成日付"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSEJ -- 作成時刻"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSPB -- 作成プログラムＩＤ"
    strSQL = strSQL & vbCrLf & ",   ZAIKS_INSUB -- 作成ユーザーＩＤ"
    strSQL = strSQL & vbCrLf & ")"
    strSQL = strSQL & vbCrLf & "SELECT"
    strSQL = strSQL & vbCrLf & "    [年月]"
    strSQL = strSQL & vbCrLf & ",   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & ",   [サイズ]"
    strSQL = strSQL & vbCrLf & ",   [ドア数]"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(A)海外・前月在庫数', [本数], NULL))          AS ZKAIQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(B)梶山・前月在庫数', [本数], NULL))          AS ZKAJQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(C)営業・前月在庫数', [本数], NULL))          AS ZEIGQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '①発注数', [本数], NULL))                     AS HACYQ"
'    strSql = strSql & vbCrLf & ",   SUM(IIF([列名] IN ('②海外・出港数', '②海外・出港数（直送）'), [本数], NULL)) AS SYUKQ --②"      'DELETE 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '②海外・出港数', [本数], NULL))               AS SYUKQ  --②"                     'INSERT 2022/09/30 K.KINEBUCHI
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '②海外・出港数', [本数], NULL))               AS KAINQ  --④"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑤梶山・撤去数', [本数], NULL))               AS TEKYQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑥梶山・売却数', [本数], NULL))               AS BAIKQ"
'DELETE 2022/09/30 K.KINEBUCHI start
'    strSql = strSql & vbCrLf & ",   SUM(IIF([列名] = '⑦梶山・営業ヤードへの出庫数', [本数], NULL)) AS ESYKQ"
'    strSql = strSql & vbCrLf & ",   SUM(IIF([列名] = '⑨梶山ヤードからの入庫数', [本数], NULL))     AS KAJNQ"
'DELETE 2022/09/30 K.KINEBUCHI end
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] IN ('⑦梶山・営業ヤードへの出庫数', '21中古購入から梶山ヤードへ出庫'), [本数], NULL)) AS ESYKQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] IN ('⑨梶山ヤードからの入庫数', '21中古購入から梶山ヤードへ出庫'), [本数], NULL))     AS KAJNQ"
'INSERT 2022/09/30 K.KINEBUCHI end
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑮港保管', [本数], NULL))                     AS MIHOQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑯不足在庫', [本数], NULL))                   AS FUSZQ"
'INSERT 2022/09/30 K.KINEBUCHI start
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(E)港保管・前月在庫数', [本数], NULL))        AS ZMIHQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑰港保管入庫数', [本数], NULL))               AS MHNYQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '⑱港保管出庫数', [本数], NULL))               AS MFSYQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '(F)中古購入・前月在庫数', [本数], NULL))      AS ZTASQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '20中古購入から自社へ発注', [本数], NULL))     AS ZTAHQ"
    strSQL = strSQL & vbCrLf & ",   SUM(IIF([列名] = '21中古購入から梶山ヤードへ出庫', [本数], NULL))   AS ZTSKQ"
    strSQL = strSQL & vbCrLf & ",   NULL                                                            AS ZTSEQ"                           '現在直送は無い
'INSERT 2022/09/30 K.KINEBUCHI end
    strSQL = strSQL & vbCrLf & ",   @wk_INSED   AS INSED"
    strSQL = strSQL & vbCrLf & ",   @wk_INSEJ   AS INSEJ"
    strSQL = strSQL & vbCrLf & ",   @wk_INSPB   AS INSPB"
    strSQL = strSQL & vbCrLf & ",   @wk_INSUB   AS INSUB"
    strSQL = strSQL & vbCrLf & "FROM"
    strSQL = strSQL & vbCrLf & "    [入出在]"
    strSQL = strSQL & vbCrLf & "GROUP BY"
    strSQL = strSQL & vbCrLf & "    [年月]"
    strSQL = strSQL & vbCrLf & ",   [仕入商品区分]"
    strSQL = strSQL & vbCrLf & ",   [サイズ]"
    strSQL = strSQL & vbCrLf & ",   [ドア数]"
    strSQL = strSQL & vbCrLf & ";"
    'INSERT 2020/07/07 Y.WADA End
    
    SqlInsertZAIK_SUMM = strSQL
Exit Function
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "SqlInsertZAIK_SUMM" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 在庫総合計トラン作成
'       MODULE_ID       : fncPLS0060
'       CREATE_DATE     : 2022/10/26            N.IMAI
'       PARAMETER       : objCon        DBコネクション(I)
'                       : strSETDATE_YM 対象年月(YYYYMM)(I)
'       RETURN          : 正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncPLS0060(objCon As Object, strSETDATE_YM As String) As Boolean
    Dim objCmd      As Object
   
    Dim ret As String
    Dim msg As String
       
    fncPLS0060 = False
   
    On Error GoTo ErrorHandler
   
    Set objCmd = CreateObject("ADODB.Command")
    
    'プロシージャ実行
    With objCmd
        .CommandTimeout = 1800
        .CommandType = 4        '4:adCmdStoredProc
        .CommandText = "PLS0060"
        Set .ActiveConnection = objCon
        .Parameters.Refresh
        .Parameters("@SETDATE_YM").VALUE = strSETDATE_YM
        .Parameters("@INSED").VALUE = strDate
        .Parameters("@INSEJ").VALUE = strTime
        .Parameters("@INSPB").VALUE = strProg
        .Parameters("@INSUB").VALUE = strUser
        .Execute
    End With
        
    Set objCmd = Nothing
    
    fncPLS0060 = True

ErrorHandler:
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncPLS0060" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
'****************************  ended of program ********************************
