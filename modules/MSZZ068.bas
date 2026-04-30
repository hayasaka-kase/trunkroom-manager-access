Attribute VB_Name = "MSZZ068"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 雑費自動割引
'
'        PROGRAM_NAME    : 雑費自動割引
'        PROGRAM_ID      : MSZZ068
'        PROGRAM_KBN     : MODULE
'        NOTE            : 月額使用料以外（雑費＆保証委託料）に対する
'                          自動割引制御を行うモジュール
'
'        CREATE          : 2011/06/11
'        CERATER         : tajima
'        Ver             : 0.0
'
'        CREATE          : 2011/08/17
'        CERATER         : M.HONDA
'        Ver             : 0.1
'                          WEB口座振替対応
'
'        CREATE          : 2014/08/25
'        CERATER         : M.HONDA
'        Ver             : 0.2
'                          ﾊﾟﾙﾏWEB24対応
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const CST_割引_銀行引落 = "1"
Private Const CST_割引_クレジットカード = "2"
Private Const CST_MSZZ068_保証委託割引番号 = "MSZZ0068_01"
Private Const CST_MODULE_ID = "MSZZ0068"
'PGPA_TABLE 取得キーワード
Private Const CST_摘要支払方法 = "TEKIYO_HOU"
Private Const CST_割引番号 = "DCNT_NO"
Private Const CST_摘要保証会社 = "TEKIYO_HCD"
Private Const CST_顧客割引率 = "KYAKU_RATE"
Private Const CST_実割引率 = "JITSU_RATE"
Private Const CST_適用文言 = "MONGON"
Private Const CST_開始日 = "SDATE"
Private Const CST_終了日 = "EDATE"
Private Const CST_鍵文字位置 = 11
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'割付情報
Public Type MSZZ068Type_DCRA_TRAN_INF
    DCRAT_ACPTNO        As String
    DCRAT_DCNT_NO       As String
    DCRAT_ENABLE        As String
    DCRAT_FROM          As String
    DCRAT_TO            As String
    DCRAT_PRICE         As Long
    DCRAT_TEXT          As String
    DCRAT_IYAKU_SEIKYU  As String
    DCRAT_SEIKYU_KBN    As String
    DCRAT_JITSU_KEIHI   As Long
End Type
'==============================================================================*
'
'       MODULE_NAME     : 保証委託料割引料金の取得
'       MODULE_ID       : MSZZ068_getHoshoItakuWaribiki
'       CREATE_DATE     : 2011/06/11
'       NOTE            : INパラメータを元に保証委託料割引の可否判断と取得を行う
'       PARAM           : a毎月支払方法
'                       : a保証会社コード
'                       : a起算日               起算日(yyyymmdd)
'                       : a保証委託料
'                       : anOut割引情報構造体   摘要された割引情報構造体
'       RETURN          : TRUE...割引アリ、FALSE...割引ナシ
'                       : ※FALSEの場合は anOut割引情報構造体 の中身保証無し
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ068_getHoshoItakuWaribiki(ByVal a毎月支払方法 As String, _
                                              ByVal a保証会社コード As String, _
                                              ByVal a起算日 As String, _
                                              ByVal a保証料 As Long, _
                                              ByRef anOut割引情報構造体 As MSZZ068Type_DCRA_TRAN_INF _
                                             ) As Boolean
    '制御変数 & 計算保持値
    Dim is割引 As Boolean
    Dim isException As Boolean
    Dim toDateText As String
    '摘要情報
    Dim str摘要保証会社 As String
    Dim str開始日       As String
    Dim str終了日       As String
    Dim int顧客割引率   As Integer
    Dim int実割引率     As Integer
    Dim str摘要文言     As String
    Dim str割引番号     As String

    is割引 = False
    isException = False
    
    On Error GoTo Exception
    
    '保証委託料割引の条件情報取得
    '2014/08/25 M.HONDA UPD
    If True = fnc保証委託料情報取得(a毎月支払方法, _
                                    a保証会社コード, _
                                    str割引番号, _
                                    str摘要保証会社, _
                                    str開始日, _
                                    str終了日, _
                                    int顧客割引率, _
                                    int実割引率, _
                                    str摘要文言) Then
    'If True = fnc保証委託料情報取得(a毎月支払方法, _
                                    str割引番号, _
                                    str摘要保証会社, _
                                    str開始日, _
                                    str終了日, _
                                    int顧客割引率, _
                                    int実割引率, _
                                    str摘要文言) Then
    '2014/08/25 M.HONDA UPD
                                    
                                    
        ' 各条件での割引適用の評価
        If str開始日 <= a起算日 And a起算日 <= str終了日 Then
            If int顧客割引率 <= int実割引率 Then
                If str摘要文言 <> "" Then
                    is割引 = True
                    If str摘要保証会社 <> "" And str摘要保証会社 <> a保証会社コード Then
                        ' この摘要保証会社のキーワードレコードが無ければ
                        ' ここのロジックは通らないので保証会社関係無しという設定も出来ます
                        is割引 = False
                    End If
                End If
            End If
        End If
    End If
    
    '割引情報をセット
    If is割引 = True Then
        'anOut割引情報構造体.DCRAT_ACPTNO 受注契約番号は上位でセットすること
        anOut割引情報構造体.DCRAT_DCNT_NO = str割引番号
        anOut割引情報構造体.DCRAT_ENABLE = "1"
        anOut割引情報構造体.DCRAT_FROM = a起算日 '起算日をセット
        toDateText = Format$(DateAdd("d", -1, DateAdd("M", 1, CDate(Left$(a起算日, 4) & "/" & Mid$(a起算日, 5, 2) & "/01"))), "yyyyMMdd")
        anOut割引情報構造体.DCRAT_TO = toDateText '起算日の月の末日セット
        anOut割引情報構造体.DCRAT_PRICE = 0 - (a保証料 * (int顧客割引率 / 100))
        anOut割引情報構造体.DCRAT_TEXT = str摘要文言
        anOut割引情報構造体.DCRAT_IYAKU_SEIKYU = 0
        anOut割引情報構造体.DCRAT_SEIKYU_KBN = "2" '保証料割引を示すコード
        anOut割引情報構造体.DCRAT_JITSU_KEIHI = a保証料 - (a保証料 * (int実割引率 / 100))
    Else
        ' 各項目初期化
        anOut割引情報構造体.DCRAT_DCNT_NO = ""
        anOut割引情報構造体.DCRAT_ENABLE = ""
        anOut割引情報構造体.DCRAT_FROM = ""
        anOut割引情報構造体.DCRAT_TO = ""
        anOut割引情報構造体.DCRAT_PRICE = 0
        anOut割引情報構造体.DCRAT_TEXT = ""
        anOut割引情報構造体.DCRAT_IYAKU_SEIKYU = 0
        anOut割引情報構造体.DCRAT_SEIKYU_KBN = ""
        anOut割引情報構造体.DCRAT_JITSU_KEIHI = 0
    End If

    GoTo Finally

Exception:
    isException = True
    
Finally:
    MSZZ068_getHoshoItakuWaribiki = is割引
    If isException = True Then
        Call Err.Raise(Err.Number, "MSZZ068_getHoshoItakuWaribiki" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function


'==============================================================================*
'
'       MODULE_NAME     : 保証委託料情報取得
'       MODULE_ID       : fnc保証委託料情報取得
'       CREATE_DATE     : 2011/6/11 tajima
'       NOTE            : INパラメータに摘要する割引情報を取得する
'                       : ※将来的に割引元情報が変わってもここだけ修正すれば良い方向にしたい
'       PARAM           : a毎月支払方法
'                       : anOut割引番号
'                       : anOut摘要保証会社
'                       : anOut開始日
'                       : anOut終了日
'                       : anOut顧客割引率
'                       : anOut実割引率
'                       : anOut摘要文言
'       RETURN          : TRUE...該当アリ、FALSE...該当ナシ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fnc保証委託料情報取得(ByVal a毎月支払方法 As String, _
                                       ByVal a保証会社コード As String, _
                                                ByRef anOut割引番号 As String, _
                                                ByRef anOut摘要保証会社 As String, _
                                                ByRef anOut開始日 As String, _
                                                ByRef anOut終了日 As String, _
                                                ByRef anOut顧客割引率 As Integer, _
                                                ByRef anOut実割引率 As Integer, _
                                                ByRef anOut摘要文言 As String _
                                            ) As Boolean
    Dim isException As Boolean  'エラー発生検知フラグ
    Dim objCon      As Object
    Dim objRst      As Object
    Dim is取得結果  As Boolean
    Dim sqlText     As String
    Dim keyWord     As String
    Dim fieldData   As String
    
    is取得結果 = False
    isException = False
    
    On Error GoTo Exception
    
    '初期化
    anOut摘要保証会社 = ""
    
    ' SQL生成
    'sqlText = "SELECT * FROM PGPA_TABL WHERE PGPAT_PGP1B LIKE '" & CST_MODULE_ID & "%' AND PGPAT_PGP2B LIKE " _
                    & "'HO_WARI_" & a毎月支払方法 & "_%' AND ORDER BY 1 "
    
    sqlText = "SELECT "
    sqlText = sqlText & "PGPAT_PGP1B, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_TEKIYO_HOU' THEN PGPAT_PARAN ELSE NULL END) AS TEKIYO_HOU, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_DCNT_NO' THEN PGPAT_PARAN ELSE NULL END) AS DCNT_NO, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_TEKIYO_HCD' THEN PGPAT_PARAN ELSE NULL END) AS TEKIYO_HCD, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_KYAKU_RATE' THEN PGPAT_PARAN ELSE NULL END) AS KYAKU_RATE, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_JITSU_RATE' THEN PGPAT_PARAN ELSE NULL END) AS JITSU_RATE, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_MONGON' THEN PGPAT_PARAN ELSE NULL END) AS MONGON, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_SDATE' THEN PGPAT_PARAN ELSE NULL END) AS SDATE, "
    sqlText = sqlText & "max(CASE PGPAT_PGP2B WHEN 'HO_WARI_" & a毎月支払方法 & "_EDATE' THEN PGPAT_PARAN ELSE NULL END) AS EDATE "
    sqlText = sqlText & "FROM PGPA_TABL WHERE PGPAT_PGP1B LIKE '" & CST_MODULE_ID & "%' "
    sqlText = sqlText & "AND PGPAT_PGP2B LIKE   'HO_WARI_2_%' "
    sqlText = sqlText & "group by PGPAT_PGP1B "
    sqlText = sqlText & "ORDER BY 1 "
    
    ' SQL実行
    Set objCon = ADODB_Connection() 'KASE接続
    Set objRst = ADODB_Recordset(sqlText, objCon)
    If objRst.EOF Then
        'パラメータテーブルに登録なし
        fnc保証委託料情報取得 = False
        GoTo Finally
    End If
    
    ' SQL実行結果を各条件へ展開
    Do Until objRst.EOF
        'keyWord = Mid$(Nz(objRst.Fields("PGPAT_PGP2B")), CST_鍵文字位置) '2014/08/25 M.HONDA DEL
        'fieldData = Nz(objRst.Fields("PGPAT_PARAN"))
        '2014/08/25 M.HONDA START
        fieldData = Nz(objRst.Fields("TEKIYO_HCD"))
        ' 共通として項目が未設定ならば評価出来ないので摘要無しにする
        If fieldData = "" Then
            GoTo Finally
        End If
        
        '保証会社が合致しないときは処理しない
        If a保証会社コード = Nz(objRst.Fields("TEKIYO_HCD")) Then
            anOut割引番号 = Nz(objRst.Fields("DCNT_NO"))
            anOut摘要保証会社 = Nz(objRst.Fields("TEKIYO_HCD"))
            anOut開始日 = Nz(objRst.Fields("SDATE"))
            anOut終了日 = Nz(objRst.Fields("EDATE"))
            anOut顧客割引率 = Nz(objRst.Fields("KYAKU_RATE"))
            anOut実割引率 = Nz(objRst.Fields("JITSU_RATE"))
            anOut摘要文言 = Nz(objRst.Fields("MONGON"))
        End If
        
'        ' 取得した各項目の振り分け、その場で判断出来れば摘要有無の評価も行う
'        Select Case keyWord
'            Case CST_摘要支払方法
'                If fieldData <> a毎月支払方法 Then
'                    GoTo Finally
'                End If
'            Case CST_割引番号
'                anOut割引番号 = fieldData
'            Case CST_摘要保証会社
'                anOut摘要保証会社 = fieldData
'            Case CST_開始日
'                anOut開始日 = fieldData
'            Case CST_終了日
'                anOut終了日 = fieldData
'            Case CST_顧客割引率
'                anOut顧客割引率 = fieldData
'            Case CST_実割引率
'                anOut実割引率 = fieldData
'            Case CST_適用文言
'                anOut摘要文言 = fieldData
'        End Select
        '2014/08/25 M.HONDA END
        objRst.MoveNext
    Loop
    is取得結果 = True

    GoTo Finally

Exception:
    isException = True
    is取得結果 = False
    
Finally:
    fnc保証委託料情報取得 = is取得結果
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "fnc保証委託料情報取得" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : 保証委託料適用割引割付取得
'       MODULE_ID       : MSZZ045_getHoshoWaribikiDCRA_TRAN
'       CREATE_DATE     : 2011/06/18            tajima
'       PARAM           : a部門コード
'                       : a受注契約番号
'                       : anOut割引情報構造体   摘要された割引情報構造体
'       RETURN          : TRUE...該当アリ、FALSE...該当ナシ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ045_getHoshoWaribikiDCRA_TRAN(ByVal a部門コード As String, _
                                                  ByVal a受注契約番号 As String, _
                                                  ByRef anOut割引割付情報 As MSZZ068Type_DCRA_TRAN_INF _
                                                  ) As Boolean
    Dim objCon      As Object
    Dim is取得結果  As Boolean
    Dim isException As Boolean  'エラー発生検知フラグ

    isException = False
    
    On Error GoTo Exception
    
    Set objCon = ADODB_Connection(a部門コード) 'KASE接続

    is取得結果 = MSZZ045_getHoshoWaribikiDCRA_TRAN2(objCon, a受注契約番号, anOut割引割付情報)

    GoTo Finally

Exception:
    isException = True
    is取得結果 = False
    
Finally:
    MSZZ045_getHoshoWaribikiDCRA_TRAN = is取得結果
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "MSZZ045_getHoshoWaribikiDCRA_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : 保証委託料適用割引割付取得(DBコネクション受取版）
'       MODULE_ID       : MSZZ045_getHoshoWaribikiDCRA_TRAN2
'       CREATE_DATE     : 2011/06/18            tajima
'       PARAM           : aDBconection
'                       : a受注契約番号
'                       : anOut割引情報構造体   摘要された割引情報構造体
'       RETURN          : TRUE...該当アリ、FALSE...該当ナシ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ045_getHoshoWaribikiDCRA_TRAN2(aDBconection As Object, _
                                                  ByVal a受注契約番号 As String, _
                                                  ByRef anOut割引割付情報 As MSZZ068Type_DCRA_TRAN_INF _
                                                  ) As Boolean

    Dim isException As Boolean  'エラー発生検知フラグ
    Dim objRst      As Object
    Dim is取得結果  As Boolean
    Dim sqlText     As String

    isException = False
    
    On Error GoTo Exception
    
    ' SQL生成 !要注意 DCRAT_DCNT_NO LIKE 'ZH%' は暫定だからね！
    sqlText = "SELECT * FROM DCRA_TRAN WHERE DCRAT_ACPTNO ='" & a受注契約番号 & "' AND DCRAT_DCNT_NO LIKE 'ZH%'"
    
    Set objRst = ADODB_Recordset(sqlText, aDBconection)
    
    If objRst.EOF Then
        '摘要された保証料値引は無し
        is取得結果 = False
        GoTo Finally
    End If

    ' SQL実行結果を各条件へ展開
    anOut割引割付情報.DCRAT_ACPTNO = objRst.Fields("DCRAT_ACPTNO")
    anOut割引割付情報.DCRAT_DCNT_NO = objRst.Fields("DCRAT_DCNT_NO")
    anOut割引割付情報.DCRAT_ENABLE = objRst.Fields("DCRAT_ENABLE")
    anOut割引割付情報.DCRAT_FROM = objRst.Fields("DCRAT_FROM")
    anOut割引割付情報.DCRAT_TO = objRst.Fields("DCRAT_TO")
    anOut割引割付情報.DCRAT_PRICE = objRst.Fields("DCRAT_PRICE")
    anOut割引割付情報.DCRAT_TEXT = objRst.Fields("DCRAT_TEXT")
    anOut割引割付情報.DCRAT_IYAKU_SEIKYU = objRst.Fields("DCRAT_IYAKU_SEIKYU")
    anOut割引割付情報.DCRAT_SEIKYU_KBN = objRst.Fields("DCRAT_SEIKYU_KBN")
    anOut割引割付情報.DCRAT_JITSU_KEIHI = objRst.Fields("DCRAT_JITSU_KEIHI")
    
    is取得結果 = True
    
    GoTo Finally

Exception:
    isException = True
    is取得結果 = False
    
Finally:
    MSZZ045_getHoshoWaribikiDCRA_TRAN2 = is取得結果
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "MSZZ045_getHoshoWaribikiDCRA_TRAN2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'==============================================================================*
'
'       MODULE_NAME     : ネット予約での保証会社コード取得
'       MODULE_ID       : MSZZ045_getNetHoshoKaishaCd
'       CREATE_DATE     : 2011/06/18            tajima
'       PARAM           : aDBconection
'       RETURN          : 取得した保証会社コード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ045_getNetHoshoKaishaCd() As String

    Dim isException     As Boolean  'エラー発生検知フラグ
    Dim objCon          As Object
    Dim objRst          As Object
    Dim str保証会社CD   As String
    Dim sqlText         As String

    isException = False
    
    On Error GoTo Exception
    
        ' SQL生成
    sqlText = "SELECT * FROM PGPA_TABL WHERE PGPAT_PGP1B LIKE '" & CST_MODULE_ID & "%' AND PGPAT_PGP2B ='HO_WARI_2_TEKIYO_HCD' "
    
    ' SQL実行
    Set objCon = ADODB_Connection() 'KASE接続
    Set objRst = ADODB_Recordset(sqlText, objCon)
    If objRst.EOF Then
        'パラメータテーブルに登録なし
        str保証会社CD = ""
    Else
        str保証会社CD = Nz(objRst.Fields("PGPAT_PARAN"), "")
    End If

    GoTo Finally

Exception:
    isException = True
    str保証会社CD = ""
    
Finally:
    MSZZ045_getNetHoshoKaishaCd = str保証会社CD
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "MSZZ045_getNetHoshoKaishaCd" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'==============================================================================*
'
'       MODULE_NAME     : 支払方法による保証会社のコレクション取得
'       MODULE_ID       : MSZZ045_getShiharaiHoshoCollection
'       CREATE_DATE     : 2011/06/18  tajima
'       PARAM           : aYYYYMMDD  '摘要日付
'       RETURN          : コレクション
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ045_getShiharaiHoshoCollection(aYYYYMMDD As String) As Collection
    
    Dim isException     As Boolean  'エラー発生検知フラグ
    Dim objCon          As Object
    Dim colItem         As New Collection
    
    On Error GoTo Exception
    
    Set objCon = ADODB_Connection()
    
    Call set支払保証会社CollectionItem(objCon, aYYYYMMDD, "1", colItem)
    Call set支払保証会社CollectionItem(objCon, aYYYYMMDD, "2", colItem)
    Call set支払保証会社CollectionItem(objCon, aYYYYMMDD, "3", colItem)      '' 2011/08/17 M.HONDA INS
    
    Set MSZZ045_getShiharaiHoshoCollection = colItem
    GoTo Finally

Exception:
    isException = True
    
Finally:
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "MSZZ045_getShiharaiHoshoCollection" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
'==============================================================================*
'
'       MODULE_NAME     : 支払方法による保証会社のコレクション取得
'       MODULE_ID       : set支払保証会社CollectionItem
'       CREATE_DATE     : 2011/06/18  tajima
'       PARAM           : aYYYYMMDD  '摘要日付
'       RETURN          : コレクション
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub set支払保証会社CollectionItem(anObjCon As Object, aYYYYMMDD As String, aGetKey As String, ByRef aCollection As Collection)
    
    Dim isException     As Boolean  'エラー発生検知フラグ
    Dim objRst          As Object
    Dim strSQL          As String
    Dim keyWord         As String
    Dim fieldData       As String
    Dim str支払方法     As String
    Dim str保証会社     As String
    Dim str開始日       As String
    Dim str終了日       As String
    
    On Error GoTo Exception
    
    isException = False
        
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " PGPAT_PGP2B,"
    strSQL = strSQL & " PGPAT_PARAN "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " PGPA_TABL "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & "     PGPAT_PGP1B LIKE '" & CST_MODULE_ID & "%' "
    strSQL = strSQL & " AND PGPAT_PGP2B LIKE 'HO_WARI_" & aGetKey + "%' "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " PGPAT_PGP2B "
    
    Set objRst = ADODB_Recordset(strSQL, anObjCon)
    
    str支払方法 = ""
    str保証会社 = ""
    str開始日 = ""
    str終了日 = ""
    
    While Not objRst.EOF
        keyWord = Mid$(Nz(objRst.Fields("PGPAT_PGP2B")), CST_鍵文字位置)
        fieldData = Nz(objRst.Fields("PGPAT_PARAN"))
        Select Case keyWord
            Case CST_摘要支払方法
                str支払方法 = fieldData
            Case CST_摘要保証会社
                str保証会社 = fieldData
            Case CST_開始日
                str開始日 = fieldData
            Case CST_終了日
                str終了日 = fieldData
       End Select
       'データが揃えば解析＆コレクション化
        If str支払方法 <> "" And str保証会社 <> "" And str開始日 <> "" And str終了日 <> "" Then
            If str開始日 <= aYYYYMMDD And aYYYYMMDD <= str終了日 Then
                aCollection.Add str保証会社, str支払方法
                GoTo Finally
            End If
            str支払方法 = ""
            str保証会社 = ""
            str開始日 = ""
            str終了日 = ""
        End If
        objRst.MoveNext
    Wend
    
    GoTo Finally

Exception:
    isException = True
    
Finally:
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "get支払保証会社CollectionItem" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub
'****************************  ended or program ********************************
