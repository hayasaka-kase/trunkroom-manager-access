Attribute VB_Name = "MSZZ045"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 自動割引制御
'
'        PROGRAM_NAME    : 自動割引制御
'        PROGRAM_ID      : MSZZ045
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2008/10/27
'        CERATER         : N.IMAI
'        Ver             : 0.0
'
'        UPDATE          : 2008/12/03
'        UPDATER         : N.IMAI
'        Ver             : 0.1
'                        : 解約まで、指定月まで適用の混在に対応
'
'        UPDATE          : 2008/12/20
'        UPDATER         : N.IMAI
'        Ver             : 0.2
'                        : RCPT_NO→RCPT_CARG_ACPTNOに変更
'
'        UPDATE          : 2008/12/20
'        UPDATER         : hirano
'        Ver             : 0.2
'                        : 画面情報にて割引月額料金取得 Function MSZZ045_fncNebikiMonCalc追加
'
'        UPDATE          : 2008/1/14
'        UPDATER         : hirano
'        Ver             : 0.3
'                        : 値引きデータ 上限チェック不具合修正
'
'        UPDATE          : 2009/1/25
'        UPDATER         : tajima
'        Ver             : 0.4
'                        : 変数型指定ミス対応
'
'        UPDATE          : 2009/02/03
'        UPDATER         : hirano
'        Ver             : 0.5
'                        : つきまたがり表示不具合解消
'
'        UPDATE          : 2009/06/30
'        UPDATER         : hirano
'        Ver             : 0.6
'                        : 割引マスタ KASE_DBへ移動に伴う変更
'
'        UPDATE          : 2009/08/17
'        UPDATER         : hirano
'        Ver             : 0.7
'                        : 割引優先順位の判定方法変更。金額指定は、優先し他割引の開始日をずらす。
'
'        UPDATE          : 2011/03/12
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.8
'                        : １．「コンテナ（バイク）」のネット予約で、コンテナマスタの用途が「コンテナ」でも「コンテナ（バイク）」の割引を適用
'                        : ２．「コンテナ」希望で「バイク」が割り当てられたネット予約で、「バイク（コンテナ）」の割引を適用
'
'        UPDATE          : 2011/04/06
'        UPDATER         : M.HONDA
'        Ver             : 0.9
'                        : １．C_USAGE_の定数を文字列に変更
'
'        UPDATE          : 2011/08/25
'        UPDATER         : M.RYU
'        Ver             : 1.0
'                        : キャンペーン情報の取得は受付日で検索
'
'        UPDATE          : 2011/08/29
'        UPDATER         : M.HONDA
'        Ver             : 1.1
'                        : 事務手数料割引対応
'
'        UPDATE          : 2011/09/5
'        UPDATER         : M.HONDA
'        Ver             : 1.2
'                        : 事務手数料割引対応
'
'        UPDATE          : 2012/04/05
'        UPDATER         : M.HONDA
'        Ver             : 1.3
'                        : 利用期間 = 制限なし対応
'
'        UPDATE          : 2013/10/29
'        UPDATER         : M.HONDA
'        Ver             : 1.4
'                        : 手動ﾄﾘﾌﾟﾙｷｬﾝﾍﾟｰﾝ対応
'
'        UPDATE          : 2016/12/14
'        UPDATER         : M.HONDA
'        Ver             : 1.5
'                        : 日割り計算の10円単位切り捨てを廃止
'
'        UPDATE          : 2024/11/27
'        UPDATER         : M.HONDA
'        Ver             : 1.6
'                        : 値引き文言取得時がNULLの場合エラーとなる問題に対応
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'割引情報
Private Type MSZZ045Type_RLDN_TRAN_INF
    RLDNT_BUMOC         As String
    RLDNT_YCODE         As String
    RLDNT_NO            As String
    RLDNT_ENABLE        As String
    RLDNT_FROM          As String
    RLDNT_TO            As String
    RLDNT_ORDER         As Integer
    RLDNT_TEXT          As String
    RLDNT_NOTE          As String
    DCNTM_TYPE          As Integer
    DCNTM_PERIOD        As Integer
    DCNTM_VALUE         As Long     '2009/1/25
    VALUE               As Long     '割引額（初回分）
    VALUE_TRUE          As Long     '割引額
    組合せインデックス  As Integer  '1:期限あり 2:期限なし
    組合せ              As Boolean
    有効フラグ          As Boolean
    DCNTM_USE_PERIOD    As Integer  '予定期間
    DCRAT_SEV1N         As String   '2015/09/30 M.HONDA INS
    DCRAT_SEV2N         As String   '2015/09/30 M.HONDA INS
    DCRAT_SEV3N         As String   '2015/09/30 M.HONDA INS
    DCRAT_ENDEN         As String   '2015/09/30 M.HONDA INS
    DCRAT_EXMONTH       As Integer  '2015/10/14 M.HONDA INS
    
End Type
Public Type MSZZ045割引情報
    件数                As Integer
    TRAN_DATA()         As MSZZ045Type_RLDN_TRAN_INF
    REQUEST_USAGE       As String                       'ネット予約希望用途     'INSERT 2011/03/12 SHIBAZAKI
End Type
'割付情報
Private Type MSZZ045Type_DCRA_TRAN_INF
    DCRAT_ACPTNO        As String
    DCRAT_DCNT_NO       As String
    DCRAT_ENABLE        As String
    DCRAT_FROM          As String
    DCRAT_TO            As String
    DCRAT_PRICE         As Long
    DCRAT_TEXT          As String
    DCRAT_IYAKU_SEIKYU  As String
    DCRAT_SEIKYU_KBN    As String
    DCNTM_PERIOD        As String
    DCNTM_TYPE          As String
    VALUE               As Long     '割引額
    組合せインデックス  As Integer  '1:期限あり 2:期限なし
    組合せ              As Boolean
End Type
Public Type MSZZ045割付情報
    年月                As String
    件数                As Integer
    設定金額            As Long
    割引合計            As Long
    TRAN_DATA()         As MSZZ045Type_DCRA_TRAN_INF
End Type

'↓INSERT 2011/04/06 M.HONDA
Private Const C_USAGE_コンテナ = "0"
Private Const C_USAGE_バイク = "3"
Private Const C_USAGE_コンテナバイク = "33"
Private Const C_USAGE_バイクコンテナ = "39"
'↓INSERT 2011/03/12 SHIBAZAKI
'Private Const C_USAGE_コンテナ = 0
'Private Const C_USAGE_バイク = 3
'Private Const C_USAGE_コンテナバイク = 33
'Private Const C_USAGE_バイクコンテナ = 39
'↑INSERT 2011/03/12 SHIBAZAKI
'↑INSERT 2011/04/06 M.HONDA

'==============================================================================*
'
'       MODULE_NAME     : 割引料金取得（DB更新前・仮計算用）
'       MODULE_ID       : MSZZ045_fncGetNebikiData
'       CREATE_DATE     : 2008/10/10            N.IMAI
'       PARAM           : aMSZZ045割引情報      割引情報構造体
'                       : strYard               ヤードコード
'                       : strNo                 コンテナ番号
'                       : lngPrice              元値（初回分金額）
'                       : lngPriceTrue          元値
'                       : strRcptUkDate         受付日(yyyymmdd)
'                       : strKisanDate          起算日(yyyymmdd)
'                       : strDCStDate           割引適用開始月(yyyymm)
'                       : strBumonCode          部門コード
'                       : lngDCMax              上限値
'                       : intUsePeriod          利用予定期間
'                       : blnネット予約         True:ネット予約
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'UPDATE 2011/08/25 M.RYU 'strRcptUkDate受付日を追加
'Public Function MSZZ045_fncGetNebikiData(ByRef aMSZZ045割引情報 As MSZZ045割引情報, strYard As String, strNO As String, lngPrice As Long, lngPriceTrue As Long, strKisanDate As String, strDCStDate As String, strBumonCode As String, lngDCMax As Long, intUsePeriod As Integer, blnネット予約 As Boolean) As Boolean
Public Function MSZZ045_fncGetNebikiData(ByRef aMSZZ045割引情報 As MSZZ045割引情報, StrYard As String, _
                      strNO As String, lngPrice As Long, lngPriceTrue As Long, strRcptUkDate As String, _
                      strKisanDate As String, strDCStDate As String, strBumonCode As String, lngDCMax As Long, _
                      intUsePeriod As Integer, blnネット予約 As Boolean, _
                      Optional chkForcing As Integer = 0) As Boolean

    Dim i                   As Integer
    Dim j                   As Integer
    Dim objAdoDbConnection  As Object
    Dim rsCNTA_MAST         As Object
    Dim rsNebikiData        As Object
    Dim strSQL              As String
    Dim lngYARD             As Long
    Dim lngNo               As Long
    Dim iIndex              As Integer
    Dim blnFLG              As Boolean
    Dim intCNTA_STEP        As Integer
    Dim dblCNTA_SIZE        As Double
    Dim intCNTA_USAGE       As Integer
    Dim lngKingaku          As Long
    Dim lngKingakuTrue      As Long
    Dim strfrom             As String
    Dim strTo               As String
    Dim intLastOfMonth      As Integer
    Dim wMSZZ045割引情報    As MSZZ045割引情報
    Dim wDate               As Date
    Dim strKaseSqlSvr       As String   'INS 2009/06/30 hirano

    On Error GoTo Exception
    
    '初期値設定
    MSZZ045_fncGetNebikiData = False
    wMSZZ045割引情報.件数 = 0
    ReDim wMSZZ045割引情報.TRAN_DATA(0)
    
    '引数チェック
    strRcptUkDate = Replace(strRcptUkDate, "/", "")     'INSERT 2011/08/25 M.RYU
    strKisanDate = Replace(strKisanDate, "/", "")
    strDCStDate = Replace(strDCStDate, "/", "")
    If StrYard = "" Or strNO = "" Or Len(strKisanDate) <> 8 Or Len(strDCStDate) <> 6 Then
        Exit Function
    End If
    '2009/06/30 INS <S> 割引マスタ KASE_DBへ移動
    'SQL-Server名
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    '2009/06/30 INS <E>
    '数値に変換
    lngYARD = CLng(StrYard)
    lngNo = CLng(strNO)

    'コンテナマスタを読込み、コンテナ情報（段区分、実帖、用途）取得
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM CNTA_MAST "
    strSQL = strSQL + "WHERE CNTA_CODE           = " & CLng(StrYard)
    strSQL = strSQL + "  AND CNTA_NO             = " & CLng(strNO)
    strSQL = strSQL + "  AND CNTA_NEBIKI_DISABLE = 0"
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(strBumonCode)
    Set rsCNTA_MAST = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
    If rsCNTA_MAST.EOF = False Then
        'ワークセット
        intCNTA_STEP = rsCNTA_MAST.Fields("CNTA_STEP")
        dblCNTA_SIZE = rsCNTA_MAST.Fields("CNTA_SIZE")
        intCNTA_USAGE = rsCNTA_MAST.Fields("CNTA_USAGE")
        
        '↓INSERT 2011/03/12 SHIBAZAKI
        If blnネット予約 Then
            If aMSZZ045割引情報.REQUEST_USAGE = C_USAGE_コンテナバイク Then
                'コンテナ（バイク）を希望=割引を適用する用途は「コンテナ（バイク）」
                intCNTA_USAGE = C_USAGE_コンテナバイク
                
            ElseIf aMSZZ045割引情報.REQUEST_USAGE = C_USAGE_コンテナ And rsCNTA_MAST.Fields("CNTA_USAGE") = C_USAGE_バイク Then
                'コンテナ希望だけどバイクボックスを受け付け=割引を適用する用途は「バイク（コンテナ）」
                intCNTA_USAGE = C_USAGE_バイクコンテナ
            End If
        End If
        '↑INSERT 2011/03/12 SHIBAZAKI
        
        '↓INSERT 2013/10/24 M.HONDA
        '強制ﾄﾘﾌﾟﾙにﾁｪｯｸが入っている場合には、強制的にﾄﾘﾌﾟﾙｷｬﾝﾍﾟｰﾝの割引を適用する。
        '割引No00009が部門共通でﾄﾘﾌﾟﾙｷｬﾝﾍﾟｰﾝの割引
        '各ﾔｰﾄﾞには割当てていないのでSQLで強制的に取得
        If chkForcing = -1 Then
            strSQL = ""
            strSQL = strSQL + "SELECT *,"
            strSQL = strSQL + "'" & strBumonCode & "' as RLDNT_BUMOC, "
            strSQL = strSQL + "'" & StrYard & "' as RLDNT_YCODE, "
            strSQL = strSQL + "'000009' as RLDNT_NO, "
            strSQL = strSQL + "'1' as RLDNT_ENABLE, "
            strSQL = strSQL + "'1' as RLDNT_ORDER, "
            strSQL = strSQL + "'キャンペーン割引・初月・翌月０円' as RLDNT_TEXT,"
            
            strSQL = strSQL + "'初月・翌月無料' as YARD_SEV1N, "
            strSQL = strSQL + "'事務手数料無料' as YARD_SEV2N, "
            strSQL = strSQL + "'' as YARD_SEV3N, "
            strSQL = strSQL + "'' as YARD_ENDEN, "
            strSQL = strSQL + "'3' as YARD_SEV_EXMONTH, "
            strSQL = strSQL + "'' as RLDNT_NOTE"
            strSQL = strSQL + "  FROM " & strKaseSqlSvr & ".dbo.DCNT_MAST "
            strSQL = strSQL + " WHERE DCNTM_NO =  '000009' "
            strSQL = strSQL + " AND   DCNTM_BUMOC =  '" & strBumonCode & "'"
            strSQL = strSQL + " ORDER BY DCNTM_USE_PERIOD DESC "
        Else
            '該当ヤードが使用できるレンタル物件割引データを読込みワークに保存
            strSQL = ""
            strSQL = strSQL + "SELECT * "
            strSQL = strSQL + "  FROM RLDN_TRAN "
            '2009/06/30 MOD <S> 割引マスタ KASE_DBへ移動
            'strSQL = strSQL + "      ,DCNT_MAST "
            strSQL = strSQL + "      ," & strKaseSqlSvr & ".dbo.DCNT_MAST "
            '2009/06/30 MOD <E>
            '2015/09/30 M.HONDA INS
            strSQL = strSQL + ",YARD_MAST "
            '2015/09/30 M.HONDA INS
            strSQL = strSQL + " WHERE RLDN_TRAN.RLDNT_BUMOC  =  '" & strBumonCode & "'"
            strSQL = strSQL + "   AND RLDN_TRAN.RLDNT_YCODE  =  '" & StrYard & "'"
    '        strSQL = strSQL + "   AND RLDN_TRAN.RLDNT_FROM   <= '" & strKisanDate & "'"    'DELETE 2011/08/25 M.RYU
    '        strSQL = strSQL + "   AND RLDN_TRAN.RLDNT_TO     >= '" & strKisanDate & "'"    'DELETE 2011/08/25 M.RYU
            strSQL = strSQL + "   AND RLDN_TRAN.RLDNT_FROM   <= " & strRcptUkDate           'INSERT 2011/08/25 M.RYU
            strSQL = strSQL + "   AND RLDN_TRAN.RLDNT_TO     >= " & strRcptUkDate           'INSERT 2011/08/25 M.RYU
            strSQL = strSQL + "   AND RLDN_TRAN.RLDNT_ENABLE =  '1'"
            '2009/02/03 DEL <S> hirano 過去分は、そのまま表示
            'strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_TO     >= '" & Format$(Now, "yyyymmdd") & "'"
            '2009/02/03 DEL <E>
            '2009/06/30 INS <S> 割引マスタ KASE_DBへ移動
            strSQL = strSQL + "   AND DCNT_MAST.DCNTM_BUMOC  =  RLDN_TRAN.RLDNT_BUMOC "
            '2009/06/30 INS <E>
            strSQL = strSQL + "   AND DCNT_MAST.DCNTM_NO     =  RLDN_TRAN.RLDNT_NO "
            '----↓↓↓↓----20110905--M.HONDA--Ins-------↓↓↓↓---<s>
            '事務手数料割引は含めない。
            strSQL = strSQL + "   AND DCNT_MAST.DCNTM_TYPE  <> 10 "
            '----↑↑↑↑----20110505--M.HONDA--Ins-------↑↑↑↑---<e>
            strSQL = strSQL + "   AND YARD_CODE = RLDNT_YCODE "
            '----↓↓↓↓----20120329--M.HONDA--Ins-------↓↓↓↓---<s>
            '' strSQL = strSQL + " ORDER BY RLDN_TRAN.RLDNT_ORDER"
            strSQL = strSQL + " ORDER BY DCNTM_USE_PERIOD DESC, RLDN_TRAN.RLDNT_ORDER "
            '----↑↑↑↑----20120329--M.HONDA--Ins-------↑↑↑↑---<e>
        End If
        '↑INSERT 2013/10/24 M.HONDA
        
        Set rsNebikiData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
        Do Until rsNebikiData.EOF
            '初期値設定
            iIndex = 0
            lngKingaku = 0
            lngKingakuTrue = 0
            '開始日を算出
            If Left$(strKisanDate, 6) = strDCStDate Then
                '起算日（年月）＝適用開始月の場合、起算日をそのままセット
                strfrom = strKisanDate
            Else
                '起算日（年月）<> 適用開始月の場合、適用開始月の月初をセット
                strfrom = strDCStDate & "01"
            End If
            Do Until rsNebikiData.EOF
                '規格外のデータは無視する
                    blnFLG = True
                '段
                If Not IsNull(rsNebikiData.Fields("DCNTM_FLOOR")) Then
                    If rsNebikiData.Fields("DCNTM_FLOOR") <> intCNTA_STEP Then
                        blnFLG = False
                    End If
                End If
                'サイズ
                If Not IsNull(rsNebikiData.Fields("DCNTM_SIZE_FROM")) Then
                    If rsNebikiData.Fields("DCNTM_SIZE_FROM") > dblCNTA_SIZE Then
                        blnFLG = False
                    End If
                End If
                If Not IsNull(rsNebikiData.Fields("DCNTM_SIZE_TO")) Then
                    If rsNebikiData.Fields("DCNTM_SIZE_TO") < dblCNTA_SIZE Then
                        blnFLG = False
                    End If
                End If
                '使用区分
                If Not IsNull(rsNebikiData.Fields("DCNTM_USAGE")) Then
                    If rsNebikiData.Fields("DCNTM_USAGE") <> intCNTA_USAGE Then
                        blnFLG = False
                    End If
                End If
                
                '利用期間
                '----↓↓↓↓----20120329--M.HONDA--Ins-------↓↓↓↓---<s>
                '' 制限なしの割引データの場合が該当する場合には、利用期間チェックは行わない。
                If rsNebikiData.Fields("DCNTM_USE_PERIOD") <> 5 Then
                '----↑↑↑↑----20120329--M.HONDA--Ins-------↑↑↑↑---<e>
                    If intUsePeriod <> 0 Then
                        If rsNebikiData.Fields("DCNTM_USE_PERIOD") > intUsePeriod Then
                            blnFLG = False
                        End If
                    End If
                '----↓↓↓↓----20120329--M.HONDA--Ins-------↓↓↓↓---<s>
                End If
                '----↑↑↑↑----20120329--M.HONDA--Ins-------↑↑↑↑---<e>
                
                '種別チェック
                If rsNebikiData.Fields("DCNTM_GENKBN") <> 99 Then
                    If blnネット予約 Then
                        If rsNebikiData.Fields("DCNTM_GENKBN") <> 1 Then
                            blnFLG = False
                        End If
                    Else
                        If rsNebikiData.Fields("DCNTM_GENKBN") <> 0 Then
                            blnFLG = False
                        End If
                    End If
                End If
                '組合せチェック
                If iIndex > 0 And blnFLG = True Then
                    blnFLG = fncb組合せチェック(wMSZZ045割引情報, rsNebikiData.Fields("DCNTM_PERIOD"), rsNebikiData.Fields("DCNTM_TYPE"), rsNebikiData.Fields("DCNTM_VALUE"))
                End If
                '割引上限チェック
                lngKingaku = 0
                lngKingakuTrue = 0
                Select Case Nz(rsNebikiData.Fields("DCNTM_TYPE"))
                    '円指定
                    Case 1
                        lngKingaku = -(lngPrice - rsNebikiData.Fields("DCNTM_VALUE"))
                        lngKingakuTrue = -(lngPriceTrue - rsNebikiData.Fields("DCNTM_VALUE"))
                    '円割引
                    Case 2
                        lngKingaku = lngKingaku - rsNebikiData.Fields("DCNTM_VALUE")
                        lngKingakuTrue = lngKingakuTrue - rsNebikiData.Fields("DCNTM_VALUE")
                    '％割引（１０の位以下切捨て）
                    Case 3
                        '2016/12/14 M.HONDA UPD
                        'lngKingaku = lngKingaku - (Fix(lngPrice * (rsNebikiData.Fields("DCNTM_VALUE") / 10000)) * 100)
                        'lngKingakuTrue = lngKingakuTrue - (Fix(lngPriceTrue * (rsNebikiData.Fields("DCNTM_VALUE") / 10000)) * 100)
                        '10の位の切り捨てを辞める。
                        lngKingaku = lngKingaku - (Fix(lngPrice * (rsNebikiData.Fields("DCNTM_VALUE") / 100)))
                        lngKingakuTrue = lngKingakuTrue - (Fix(lngPriceTrue * (rsNebikiData.Fields("DCNTM_VALUE") / 100)))
                        '2016/12/14 M.HONDA UPD
                    '----↓↓↓↓----2011/08/29--M.HONDA--Ins-------↓↓↓↓---<s>
                    Case 10
                        '' 事務手数料割引は割引割付を作らない
                        blnFLG = False
                    '----↑↑↑↑----2011/08/29--M.HONDA--Ins-------↑↑↑↑---<e>
                End Select
                If Nz(rsNebikiData.Fields("DCNTM_TYPE")) = 2 Then
                    If Left$(strKisanDate, 6) = strDCStDate And Mid(strKisanDate, 7, 2) <> 1 Then
                        '日割り計算（１日以外）
                        '日割り日数÷起算日の月の最終日の日にち×月額(１０の位以下切捨て)
                        '起算日の月の最終日の日にち
                        intLastOfMonth = Day(DateAdd("d", -1, (DateAdd("M", 1, DateSerial(Left$(strDCStDate, 4), Mid$(strDCStDate, 5, 2), 1)))))
                        '▼2009/1/14 hirano MOD 日数計算不具合修正　カッコの位置が異なる
                        'lngKingaku = Fix(((intLastOfMonth - Int(Right$(strKisanDate, 2) + 1)) / intLastOfMonth * lngKingaku) / 100) * 100
                        '2016/12/14 M.HONDA UPD
                        '10単位の切り捨てを廃止
                        'lngKingaku = Fix(((intLastOfMonth - Int(Right$(strKisanDate, 2)) + 1) / intLastOfMonth * lngKingaku) / 100) * 100
                        lngKingaku = Fix(((intLastOfMonth - Int(Right$(strKisanDate, 2)) + 1) / intLastOfMonth * lngKingaku))
                        '▲2009/1/14 hirano MOD
                        '2016/12/14 M.HONDA UPD
                    End If
                End If
                'ワークセット
                If blnFLG = True Then
                    ReDim Preserve wMSZZ045割引情報.TRAN_DATA(iIndex)
                    '終了日を算出
                    If rsNebikiData.Fields("DCNTM_PERIOD") = 0 Then
                        strTo = "99999999"
                    Else
                        strTo = Format$(DateAdd("d", -1, DateAdd("M", Nz(rsNebikiData.Fields("DCNTM_PERIOD")), CDate(Left$(strfrom, 4) & "/" & Mid$(strfrom, 5, 2) & "/01"))), "yyyyMMdd")
                    End If
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_BUMOC = rsNebikiData.Fields("RLDNT_BUMOC")
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_YCODE = rsNebikiData.Fields("RLDNT_YCODE")
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_NO = rsNebikiData.Fields("RLDNT_NO")
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_ENABLE = rsNebikiData.Fields("RLDNT_ENABLE")
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_FROM = strfrom
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_TO = strTo
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_ORDER = rsNebikiData.Fields("RLDNT_ORDER")
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_TEXT = rsNebikiData.Fields("RLDNT_TEXT")
                    wMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_NOTE = Nz(rsNebikiData.Fields("RLDNT_NOTE"))
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCNTM_TYPE = Nz(rsNebikiData.Fields("DCNTM_TYPE"))
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCNTM_PERIOD = Nz(rsNebikiData.Fields("DCNTM_PERIOD"))
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCNTM_VALUE = Nz(rsNebikiData.Fields("DCNTM_VALUE"))
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCNTM_USE_PERIOD = rsNebikiData.Fields("DCNTM_USE_PERIOD")
'                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_SEV1N = rsNebikiData.Fields("YARD_SEV1N") '2015/09/30 M.HONDA INS
'                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_SEV2N = rsNebikiData.Fields("YARD_SEV2N") '2015/09/30 M.HONDA INS
'                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_SEV3N = rsNebikiData.Fields("YARD_SEV3N") '2015/09/30 M.HONDA INS
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_SEV1N = Nz(rsNebikiData.Fields("YARD_SEV1N"))  'INSERT 2024/11/27 N.IMAI
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_SEV2N = Nz(rsNebikiData.Fields("YARD_SEV2N"))  'INSERT 2024/11/27 N.IMAI
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_SEV3N = Nz(rsNebikiData.Fields("YARD_SEV3N"))  'INSERT 2024/11/27 N.IMAI
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_ENDEN = Nz(rsNebikiData.Fields("YARD_ENDEN"))  'INSERT 2024/11/27 N.IMAI
                    wMSZZ045割引情報.TRAN_DATA(iIndex).DCRAT_EXMONTH = rsNebikiData.Fields("YARD_SEV_EXMONTH") '2015/10/14 M.HONDA INS
                    
                    
                    '金額算出
                    If Left$(strKisanDate, 6) <> strDCStDate Then
                        wMSZZ045割引情報.TRAN_DATA(iIndex).VALUE = lngKingakuTrue
                    Else
                        wMSZZ045割引情報.TRAN_DATA(iIndex).VALUE = lngKingaku
                    End If
                    wMSZZ045割引情報.TRAN_DATA(iIndex).VALUE_TRUE = lngKingakuTrue
                    If rsNebikiData.Fields("DCNTM_PERIOD") = 0 Then
                        '解約まで適用
                        wMSZZ045割引情報.TRAN_DATA(iIndex).組合せインデックス = 1
                        If rsNebikiData.Fields("DCNTM_TYPE") = 1 Then
                            wMSZZ045割引情報.TRAN_DATA(iIndex).組合せ = False
                        Else
                            wMSZZ045割引情報.TRAN_DATA(iIndex).組合せ = True
                        End If
                    Else
                        '指定月まで適用
                        wMSZZ045割引情報.TRAN_DATA(iIndex).組合せインデックス = 2
                        If rsNebikiData.Fields("DCNTM_TYPE") = 1 Then
                            wMSZZ045割引情報.TRAN_DATA(iIndex).組合せ = False
                        Else
                            wMSZZ045割引情報.TRAN_DATA(iIndex).組合せ = True
                        End If
                    End If
                    wMSZZ045割引情報.TRAN_DATA(iIndex).有効フラグ = True
                    iIndex = iIndex + 1
                    wMSZZ045割引情報.件数 = iIndex                              'I:2008/12/02 N.IMAI
                End If
                rsNebikiData.MoveNext
            Loop
            'D:2008/12/02 N.IMAI aMSZZ045割引情報.件数 = iIndex
        Loop
    End If
    '矛盾チェック
    '円指定チェック
    For i = 0 To wMSZZ045割引情報.件数 - 1
        '解約までの円指定がある場合、他の円引き・％引きは無効にする
        If wMSZZ045割引情報.TRAN_DATA(i).DCNTM_TYPE = 1 _
        And wMSZZ045割引情報.TRAN_DATA(i).DCNTM_PERIOD = 0 Then
            For j = i + 1 To wMSZZ045割引情報.件数 - 1
                If wMSZZ045割引情報.TRAN_DATA(j).DCNTM_TYPE <> 1 Then
                    wMSZZ045割引情報.TRAN_DATA(j).有効フラグ = False
                End If
            Next j
        End If
    Next i
    '期間矛盾チェック
    For i = 0 To wMSZZ045割引情報.件数 - 1
        If wMSZZ045割引情報.TRAN_DATA(i).有効フラグ = True Then
            '円指定がある場合、それ以外のFrom,Toをずらす
            If wMSZZ045割引情報.TRAN_DATA(i).DCNTM_TYPE = 1 Then
                If wMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO <> 99999999 Then
                    wDate = CDate(Left$(wMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO, 4) _
                    & "/" & Mid$(wMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO, 5, 2) _
                    & "/" & Mid$(wMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO, 7, 2))
                    '2009/08/17 MOD hirano <S> 期間をずらす範囲は、1件目から、円指定なので、割引開始日が同じモノをずらす
                    'For j = i + 1 To wMSZZ045割引情報.件数 - 1
                    '    wMSZZ045割引情報.TRAN_DATA(j).RLDNT_FROM = Format$(DateAdd("d", 1, wDate), "yyyyMMdd")
                    '    If wMSZZ045割引情報.TRAN_DATA(j).RLDNT_TO <> 99999999 Then
                    '        wMSZZ045割引情報.TRAN_DATA(j).RLDNT_TO = Format$(DateAdd("d", -1, DateAdd("M", wMSZZ045割引情報.TRAN_DATA(j).DCNTM_PERIOD, wDate)), "yyyyMMdd")
                    '    End If
                    '    '初月ではなくなるので、値引金額も初月分ではなくなる
                    '    wMSZZ045割引情報.TRAN_DATA(j).VALUE = wMSZZ045割引情報.TRAN_DATA(j).VALUE_TRUE
                    'Next j
                    For j = 0 To wMSZZ045割引情報.件数 - 1
                        '該当レコード以外の開始をずらす
                        If wMSZZ045割引情報.TRAN_DATA(j).有効フラグ = True And _
                         j <> i And _
                         wMSZZ045割引情報.TRAN_DATA(j).RLDNT_FROM = wMSZZ045割引情報.TRAN_DATA(i).RLDNT_FROM Then
                            wMSZZ045割引情報.TRAN_DATA(j).RLDNT_FROM = Format$(DateAdd("d", 1, wDate), "yyyyMMdd")
                            If wMSZZ045割引情報.TRAN_DATA(j).RLDNT_TO <> 99999999 Then
                                wMSZZ045割引情報.TRAN_DATA(j).RLDNT_TO = Format$(DateAdd("d", -1, DateAdd("M", wMSZZ045割引情報.TRAN_DATA(j).DCNTM_PERIOD, wDate)), "yyyyMMdd")
                            End If
                            '初月ではなくなるので、値引金額も初月分ではなくなる
                            wMSZZ045割引情報.TRAN_DATA(j).VALUE = wMSZZ045割引情報.TRAN_DATA(j).VALUE_TRUE
                        End If
                    Next j
                    '2009/08/17 MOD hirano <E>
                Else
                    For j = i + 1 To wMSZZ045割引情報.件数 - 1
                        '解約まで円指定があった場合、それ以降の円指定は無効にする
                        If wMSZZ045割引情報.TRAN_DATA(j).DCNTM_TYPE = 1 Then
                            wMSZZ045割引情報.TRAN_DATA(j).有効フラグ = False
                        End If
                    Next j
                End If
            End If
        End If
    Next i
    '上限チェック
    lngKingaku = 0
    For i = 0 To wMSZZ045割引情報.件数 - 1
        If wMSZZ045割引情報.TRAN_DATA(i).有効フラグ = True _
        And wMSZZ045割引情報.TRAN_DATA(i).DCNTM_TYPE <> 1 Then
            '2008/12/22 lngKingaku = lngKingaku + wMSZZ045割引情報.TRAN_DATA(i).VALUE
            lngKingaku = lngKingaku + wMSZZ045割引情報.TRAN_DATA(i).VALUE_TRUE  'I:2008/12/22
            '▼2009/1/14 hirano MOD 有効対象のチェック修正
            'If -lngDCMax > lngKingaku Or lngPrice + lngKingaku < 0 Then
            If -lngDCMax > lngKingaku Or lngPriceTrue + lngKingaku < 0 Then
            '▲2009/1/14 hirano MOD
                wMSZZ045割引情報.TRAN_DATA(i).有効フラグ = False
                'D:2008/12/22 lngKingaku = lngKingaku - wMSZZ045割引情報.TRAN_DATA(i).VALUE
                lngKingaku = lngKingaku - wMSZZ045割引情報.TRAN_DATA(i).VALUE_TRUE  'I:2008/12/22
            End If
        End If
    Next i
    '有効なものだけ返す
        iIndex = 0
    aMSZZ045割引情報.件数 = 0
    For i = 0 To wMSZZ045割引情報.件数 - 1
        If wMSZZ045割引情報.TRAN_DATA(i).有効フラグ = True Then
            ReDim Preserve aMSZZ045割引情報.TRAN_DATA(iIndex)
            aMSZZ045割引情報.TRAN_DATA(iIndex) = wMSZZ045割引情報.TRAN_DATA(i)
            aMSZZ045割引情報.件数 = aMSZZ045割引情報.件数 + 1
            iIndex = iIndex + 1
        End If
    Next i
    MSZZ045_fncGetNebikiData = True
    
Exception:

    If Not rsCNTA_MAST Is Nothing Then: rsCNTA_MAST.Close: Set rsCNTA_MAST = Nothing
    If Not rsNebikiData Is Nothing Then: rsNebikiData.Close: Set rsNebikiData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    If Err <> 0 Then
        MSZZ045_fncGetNebikiData = False
        Call Err.Raise(Err.Number, "MSZZ045_fncGetNebikiData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'       MODULE_NAME     : 組合せチェック
'       MODULE_ID       : fncb組合せチェック
'       CREATE_DATE     : 2008/10/10            N.IMAI
'       PARAM           : aMSZZ045割引情報      割引情報構造体
'                       : intPeriod             割引期間
'                       : intType               割引形式
'                       : lngValue              割引値
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncb組合せチェック(ByRef aMSZZ045割引情報 As MSZZ045割引情報, intPeriod As Integer, intType As Integer, lngValue As Long) As Boolean

    Dim i           As Integer
    Dim iIndex      As Integer
    Dim blnFLG      As Boolean

    '初期値設定
    fncb組合せチェック = True
    
    If aMSZZ045割引情報.件数 > 0 Then                                           'I:2008/12/02
        '今回の組合せ
        If intPeriod = 0 Then
            '解約まで適用
            iIndex = 1
            If intType = 1 Then
                blnFLG = False
            Else
                blnFLG = True
            End If
        Else
            '指定月まで適用
            iIndex = 2
            If intType = 1 Then
                blnFLG = False
            Else
                blnFLG = True
            End If
        End If
        '過去の割引と組合せ可能かチェック
        For i = 0 To UBound(aMSZZ045割引情報.TRAN_DATA)
        
            '----↓↓↓↓----2012/03/29--M.HONDA--Ins-------↓↓↓↓---<s>
            ' 利用期間 : 無期限割が適用されていれば、他の割引は適用不可
            If aMSZZ045割引情報.TRAN_DATA(i).DCNTM_USE_PERIOD = 5 Then
                fncb組合せチェック = False
                Exit For
            End If
            '----↑↑↑↑----2012/03/29--M.HONDA--Ins-------↑↑↑↑---<e>
        
            '既に○×が異なるデータが存在する場合、適用できない
            If aMSZZ045割引情報.TRAN_DATA(i).組合せインデックス = iIndex _
            And aMSZZ045割引情報.TRAN_DATA(i).組合せ <> blnFLG Then
                fncb組合せチェック = False
                Exit For
            End If
            '既に×が存在する場合、適用できない
            If aMSZZ045割引情報.TRAN_DATA(i).組合せインデックス = iIndex _
            And aMSZZ045割引情報.TRAN_DATA(i).組合せ = False Then
                fncb組合せチェック = False
                Exit For
            End If
        Next i
    End If

End Function

'==============================================================================*
'
'       MODULE_NAME     : 適用割引割付取得（本計算用）
'       MODULE_ID       : MSZZ045_fncGetDCRA_TRAN
'       CREATE_DATE     : 2008/10/10            N.IMAI
'       PARAM           : aMSZZ045割付情報      適用割引割付情報構造体
'                       : strAcptNo             受注契約番号
'                       : strBumonCode          部門コード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ045_fncGetDCRA_TRAN(ByRef aMSZZ045割引情報 As MSZZ045割引情報, strACPTNO As String, strBumonCode As String) As Boolean

    Dim objAdoDbConnection  As Object
    Dim rsData              As Object
    Dim strSQL              As String
    Dim iIndex              As Integer
    Dim strDcntNo           As String                                           'I:2008/12/22
    Dim strKaseSqlSvr       As String                                           'INS 2009/06/30 hirano

    On Error GoTo Exception
    
    '初期値設定
    MSZZ045_fncGetDCRA_TRAN = False
    strDcntNo = ""                                                              'I:2008/12/22
    
    '引数チェック
    If strACPTNO = "" Then
        MSZZ045_fncGetDCRA_TRAN = False
        Exit Function
    End If

    '2009/06/30 INS <S> 割引マスタ KASE_DBへ移動
    'SQL-Server名
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    '2009/06/30 INS <E>

    '適用割引割付情報取得
    strSQL = ""
    strSQL = strSQL + "SELECT * "
    strSQL = strSQL + "  FROM DCRA_TRAN "
    '2009/06/30 MOD <S> 割引マスタKASE_DBへ移行
    'strSQL = strSQL + "      ,DCNT_MAST "
    strSQL = strSQL + "      ," & strKaseSqlSvr & ".dbo.DCNT_MAST "
    '2009/06/30 MOD <E>
    strSQL = strSQL + " WHERE DCRAT_ACPTNO     = '" & strACPTNO & "' "
    strSQL = strSQL + "   AND DCNTM_NO         = DCRA_TRAN.DCRAT_DCNT_NO "
    '2009/06/30 INS <S> 割引マスタKASE_DBへ移行
    strSQL = strSQL + "   AND DCNTM_BUMOC     = '" & strBumonCode & "' "
    '2009/06/30 INS <E>
    strSQL = strSQL + " ORDER BY DCNTM_NO, DCRAT_FROM "
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(strBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
    '初期値設定
    iIndex = 0
    aMSZZ045割引情報.件数 = 0
    If rsData.EOF = False Then
        Do Until rsData.EOF
            'ワークセット
            '2008/12/22 Start
            If strDcntNo <> rsData.Fields("DCRAT_DCNT_NO") Then
                ReDim Preserve aMSZZ045割引情報.TRAN_DATA(iIndex)
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_BUMOC = strBumonCode
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_YCODE = ""
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_NO = rsData.Fields("DCRAT_DCNT_NO")
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_ENABLE = rsData.Fields("DCRAT_ENABLE")
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_FROM = rsData.Fields("DCRAT_FROM")
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_TO = rsData.Fields("DCRAT_TO")
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_ORDER = 0
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_TEXT = rsData.Fields("DCRAT_TEXT")
                aMSZZ045割引情報.TRAN_DATA(iIndex).RLDNT_NOTE = Nz(rsData.Fields("DCNTM_MTEXT"))
                aMSZZ045割引情報.TRAN_DATA(iIndex).DCNTM_TYPE = Nz(rsData.Fields("DCNTM_TYPE"))
                aMSZZ045割引情報.TRAN_DATA(iIndex).DCNTM_PERIOD = Nz(rsData.Fields("DCNTM_PERIOD"))
                aMSZZ045割引情報.TRAN_DATA(iIndex).DCNTM_VALUE = Nz(rsData.Fields("DCNTM_VALUE"))
                aMSZZ045割引情報.TRAN_DATA(iIndex).VALUE = Nz(rsData.Fields("DCRAT_PRICE"))
                aMSZZ045割引情報.TRAN_DATA(iIndex).VALUE_TRUE = Nz(rsData.Fields("DCRAT_PRICE"))
                aMSZZ045割引情報.TRAN_DATA(iIndex).有効フラグ = True
                iIndex = iIndex + 1
                aMSZZ045割引情報.件数 = iIndex
                strDcntNo = rsData.Fields("DCRAT_DCNT_NO")
            Else
                aMSZZ045割引情報.TRAN_DATA(iIndex - 1).RLDNT_TO = rsData.Fields("DCRAT_TO")
                aMSZZ045割引情報.TRAN_DATA(iIndex - 1).VALUE_TRUE = Nz(rsData.Fields("DCRAT_PRICE"))
            End If
            '2008/12/22 End
            rsData.MoveNext
        Loop
    End If
    MSZZ045_fncGetDCRA_TRAN = True
    
Exception:

    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "MSZZ045_fncGetDCRA_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'       MODULE_NAME     : 画面情報にて仮の割引料金取得（FVS400のみ使用？）
'       MODULE_ID       : MSZZ045_fncNebikiCalc
'       CREATE_DATE     : 2008/11/16            N.IMAI
'       PARAM           : aMSZZ045割付情報      適用割引割付情報構造体
'                       : aMSZZ045割引情報      割引情報構造体
'                       : strNowDateYM          該当年月（yyyymm）
'                       : strStartYM            適用開始年月（yyyymm）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ045_fncNebikiCalc(ByRef aMSZZ045割付情報 As MSZZ045割付情報, aMSZZ045割引情報 As MSZZ045割引情報, strNowDateYM As String, strStartYM As String) As Boolean

    Dim i                   As Integer
    Dim iIndex              As Integer

    On Error GoTo Exception
    
    '初期値設定
    MSZZ045_fncNebikiCalc = False
    aMSZZ045割付情報.件数 = 0
    aMSZZ045割付情報.設定金額 = 0
    aMSZZ045割付情報.割引合計 = 0
    aMSZZ045割付情報.年月 = Left$(strNowDateYM, 6)
    iIndex = 0
    
    For i = 0 To aMSZZ045割引情報.件数 - 1
'20081216 Start
'        If Left$(aMSZZ045割引情報.TRAN_DATA(i).RLDNT_FROM, 6) <= strNowDateYM _
'        And Left$(aMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO, 6) >= strNowDateYM Then
        If Left$(aMSZZ045割引情報.TRAN_DATA(i).RLDNT_FROM, 6) <= Left$(strNowDateYM, 6) _
        And Left$(aMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO, 6) >= Left$(strNowDateYM, 6) Then
'20081216 End
            ReDim Preserve aMSZZ045割付情報.TRAN_DATA(iIndex)
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_ACPTNO = ""
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_DCNT_NO = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_NO
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_ENABLE = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_ENABLE
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_FROM = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_FROM
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_TO = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_PRICE = aMSZZ045割引情報.TRAN_DATA(i).VALUE
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_TEXT = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_TEXT
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_IYAKU_SEIKYU = ""
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_SEIKYU_KBN = ""
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCNTM_PERIOD = aMSZZ045割引情報.TRAN_DATA(i).DCNTM_PERIOD
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCNTM_TYPE = aMSZZ045割引情報.TRAN_DATA(i).DCNTM_TYPE
            If Left$(strNowDateYM, 6) = strStartYM Then
                aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE = aMSZZ045割引情報.TRAN_DATA(i).VALUE
            Else
                aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE = aMSZZ045割引情報.TRAN_DATA(i).VALUE_TRUE
            End If
            If aMSZZ045割引情報.TRAN_DATA(i).DCNTM_TYPE = 1 Then
                aMSZZ045割付情報.設定金額 = aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE
                aMSZZ045割付情報.割引合計 = aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE
            Else
                aMSZZ045割付情報.割引合計 = aMSZZ045割付情報.割引合計 + aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE
            End If
            aMSZZ045割付情報.件数 = aMSZZ045割付情報.件数 + 1
            iIndex = iIndex + 1
        End If
    Next i
    MSZZ045_fncNebikiCalc = True
    
Exception:

    If Err <> 0 Then
        MSZZ045_fncNebikiCalc = False
        Call Err.Raise(Err.Number, "MSZZ045_fncNebikiCalc" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'==============================================================================*
'
'       MODULE_NAME     : 画面情報にて割引月額料金取得（FVS400のみ使用）
'       MODULE_ID       : MSZZ045_fncNebikiMonCalc
'       CREATE_DATE     : 2008/12/20            hirano
'       PARAM           : aMSZZ045割付情報      適用割引割付情報構造体
'                       : aMSZZ045割引情報      割引情報構造体
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ045_fncNebikiMonCalc(ByRef aMSZZ045割付情報 As MSZZ045割付情報, aMSZZ045割引情報 As MSZZ045割引情報) As Boolean

    Dim i                   As Integer
    Dim iIndex              As Integer

    On Error GoTo Exception
    
    '初期値設定
    MSZZ045_fncNebikiMonCalc = False
    aMSZZ045割付情報.件数 = 0
    aMSZZ045割付情報.設定金額 = 0
    aMSZZ045割付情報.割引合計 = 0
    iIndex = 0
    For i = 0 To aMSZZ045割引情報.件数 - 1
        '有効可否:1_有効,有効期間To:99999999 のデータ（毎月値引き）
        If aMSZZ045割引情報.TRAN_DATA(i).RLDNT_ENABLE = "1" _
        And aMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO = "99999999" Then
            ReDim Preserve aMSZZ045割付情報.TRAN_DATA(iIndex)
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_ACPTNO = ""
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_DCNT_NO = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_NO
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_ENABLE = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_ENABLE
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_FROM = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_FROM
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_TO = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_TO
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_PRICE = aMSZZ045割引情報.TRAN_DATA(i).VALUE
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_TEXT = aMSZZ045割引情報.TRAN_DATA(i).RLDNT_TEXT
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_IYAKU_SEIKYU = ""
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCRAT_SEIKYU_KBN = ""
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCNTM_PERIOD = aMSZZ045割引情報.TRAN_DATA(i).DCNTM_PERIOD
            aMSZZ045割付情報.TRAN_DATA(iIndex).DCNTM_TYPE = aMSZZ045割引情報.TRAN_DATA(i).DCNTM_TYPE
            '月額をセットする
            aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE = aMSZZ045割引情報.TRAN_DATA(i).VALUE_TRUE
            If aMSZZ045割引情報.TRAN_DATA(i).DCNTM_TYPE = 1 Then
                aMSZZ045割付情報.設定金額 = aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE
                aMSZZ045割付情報.割引合計 = aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE
            Else
                aMSZZ045割付情報.割引合計 = aMSZZ045割付情報.割引合計 + aMSZZ045割付情報.TRAN_DATA(iIndex).VALUE
            End If
            aMSZZ045割付情報.件数 = aMSZZ045割付情報.件数 + 1
            iIndex = iIndex + 1
        End If
    Next i
    MSZZ045_fncNebikiMonCalc = True
Exception:

    If Err <> 0 Then
        MSZZ045_fncNebikiMonCalc = False
        Call Err.Raise(Err.Number, "MSZZ045_fncNebikiCalc" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'****************************  ended or program ********************************
