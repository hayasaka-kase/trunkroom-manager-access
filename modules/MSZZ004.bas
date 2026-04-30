Attribute VB_Name = "MSZZ004"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : 消費税
'        PROGRAM_ID      : MSZZ004
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2003/04/16
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2004/04/22
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'                        : ﾚｺｰﾄﾞｾｯﾄへ変更
'
'        UPDATE          : 2004/12/24
'        UPDATER         : N.MIURA
'        Ver             : 0.2
'                        : 切り上げ値修正
'
'        UPDATE          : 2007/03/15
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.3
'                        : MSZZ004_M10 追加
'                        : MSZZ004_M20 追加
'
'        UPDATE          : 2009/05/19
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.4
'                        : MSZZ004_M30(端数処理) 追加
'
'        UPDATE          : 2010/04/14
'        UPDATER         : YAMA
'        Ver             : 0.5
'                        : MSZZ004_M40(消費税) 追加
'
'        UPDATE          : 2010/06/30
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.6
'                        : 内税計算では想定で算出してた値をそのまま解とする
'
'        UPDATE          : 2014/04/01
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.7
'                        : 税率を引数で受け取るバージョンを追加
'                        : MSZZ004_M11(消費税)
'                        : MSZZ004_M41(消費税)
'
'        UPDATE          : 2014/04/17
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.8
'                        : 端数区分を引数で受け取るバージョンを追加（管理委託で使用）
'                        : MSZZ004_M12
'                        : MSZZ004_M42
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "MSZZ004"
'
Private dbs             As Database                                             'INSERT 20040422 N.MIURA
Private strSQL          As String                                               'INSERT 20040422 N.MIURA
'
Private RST_CONF        As Recordset                                            'INSERT 20040422 N.MIURA
'
Private USER_ID         As String
Private MODL_ID         As String
'
Private WK_ERR          As String
'
Private WK_SYYMD        As String
Private WK_SYOLR        As Double
Private WK_SYNWR        As Double
Private WK_SYKBI        As String
'
'==============================================================================*
'   テスト用
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ004_TEST()
    Dim TEST_SYOHR As Double
    Dim TEST_SYKBI As Double
    
    Call MSZZ004_M00("200312", TEST_SYOHR, TEST_SYKBI)
    MsgBox (TEST_SYOHR)
    MsgBox (TEST_SYKBI)

End Function
'==============================================================================*
'
'        MODULE_NAME      :メイン
'        MODULE_ID        :MSZZ004
'        CREATE_DATE      :2003/02/10
'        PARAM            : strTAYMD            算出年月(I) yyyymm
'                         : strSYOHR            税率(O)
'                         : strSYKBI            端数処理用加算値(O) Double
'        RETURN           : 正常(True)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ004_M00(strTAYMD As String, strSYOHR As Double, strSYKBI As Double) As Boolean
    On Error GoTo MSZZ004_M00_Err
    
    MSZZ004_M00 = False
    
    Call MSZZ004_DBOPEN                                                         'INSERT 20040422 N.MIURA
    
    USER_ID = LsGetUserName()
    
    WK_ERR = 0
    
    'WK_SYYMD = DlookUp("CONFT_SYYMD", "CONF_TABL")                             'DELETE 20040422 N.MIURA START
    'WK_SYOLR = DlookUp("CONFT_SYOLR", "CONF_TABL")
    'WK_SYNWR = DlookUp("CONFT_SYNWR", "CONF_TABL")
    'WK_SYKBI = DlookUp("CONFT_SYKBI", "CONF_TABL")                             'DELETE 20040422 N.MIURA ENDED
    
    strSQL = ""                                                                 'INSERT 20040422 N.MIURA START
'    strSQL = strSQL & "SELECT "                                                'DELETE 20070317 K.ISHIZAKA START
'    strSQL = strSQL & "CONFT_SYYMD, "
'    strSQL = strSQL & "CONFT_SYOLR, "
'    strSQL = strSQL & "CONFT_SYNWR, "
'    strSQL = strSQL & "CONFT_SYKBI  "
'    strSQL = strSQL & "FROM "
'    strSQL = strSQL & "CONF_TABL "
'    strSQL = strSQL & "WHERE "
'    strSQL = strSQL & "CONFT_NUMBC = '1' "
'    strSQL = strSQL & "; "                                                     'DELETE 20070317 K.ISHIZAKA ENDED
    strSQL = SELECT_CONF_TABL()                                                 'INSERT 20070317 K.ISHIZAKA
    
    Set RST_CONF = dbs.OpenRecordset(strSQL, dbOpenDynaset, dbReadOnly)
    
    If Not RST_CONF.EOF Then
       WK_SYYMD = RST_CONF.Fields("CONFT_SYYMD")
       WK_SYOLR = RST_CONF.Fields("CONFT_SYOLR")
       WK_SYNWR = RST_CONF.Fields("CONFT_SYNWR")
       WK_SYKBI = RST_CONF.Fields("CONFT_SYKBI")
    Else
       WK_ERR = 1
    End If
    
    RST_CONF.Close
           
    Set RST_CONF = Nothing                                                      'INSERT 20040422 N.MIURA ENDED
    
    If strTAYMD >= WK_SYYMD Then
       strSYOHR = WK_SYNWR
    Else
       strSYOHR = WK_SYOLR
    End If
    
    Select Case WK_SYKBI
           Case 0 '四捨五入
                strSYKBI = 0.5
           Case 1 '切捨て
                strSYKBI = 0
           Case 2 '切上げ
                'strSYKBI = 0.6                                                 'DELETE 20041224 N.MIURA
                strSYKBI = 0.9                                                  'INSERT 20041224 N.MIURA
    End Select
    
    If WK_ERR = 0 Then
       MSZZ004_M00 = True
    End If
    
    Call MSZZ004_DBCLOS                                                         'INSERT 20040422 N.MIURA
    
MSZZ004_M00_Exit:
    Exit Function

MSZZ004_M00_Err:
    MsgBox Err.Description
    Resume MSZZ004_M00_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :MSZZ004_DBOPEN
'        CREATE_DATE      :2004/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MSZZ004_DBOPEN()
    On Error GoTo Err_MSZZ004_DBOPEN
    
    Set dbs = CurrentDb()
    
Exit_MSZZ004_DBOPEN:
    Exit Function

Err_MSZZ004_DBOPEN:
    MsgBox Err.Description
    Resume Exit_MSZZ004_DBOPEN
End Function
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :MSZZ004_DBCLOS
'        CREATE_DATE      :2004/04/22
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MSZZ004_DBCLOS()
    On Error GoTo Err_MSZZ004_DBCLOS
    
    dbs.Close
    
    Set dbs = Nothing

Exit_MSZZ004_DBCLOS:
    Exit Function

Err_MSZZ004_DBCLOS:
    MsgBox Err.Description
    Resume Exit_MSZZ004_DBCLOS
End Function

'==============================================================================*
'   テスト用
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_AA0()
    Dim i       As Long
    
    For i = 19 To 129
        Debug.Print "===" & Format(i) & "==="
        Call TEST_AA1(i)
        If i Mod 10 = 0 Then
            Stop
        End If
    Next
End Sub

'==============================================================================*
'   テスト用
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_AA1(ByVal lngKinga As Long)
    Dim lngPrice    As Long
    Dim lngTax      As Long
    On Error GoTo ErrorHandler
    
    Call MSZZ004_M10(lngKinga, "200703", "1", lngPrice, lngTax)
    Debug.Print "外税   Price=" & Format(lngPrice) & " Tax=" & Format(lngTax) & " 税込み=" & Format(lngPrice + lngTax)
    
    Call MSZZ004_M10(lngKinga, "200703", "2", lngPrice, lngTax)
    Debug.Print "内税   Price=" & Format(lngPrice) & " Tax=" & Format(lngTax)
    
'    Call MSZZ004_M10(lngKinga, "200703", "3", lngPrice, lngTax)
'    Debug.Print "非課税 Price=" & Format(lngPrice) & " Tax=" & Format(lngTax)
Exit Sub

ErrorHandler:
    Debug.Print Err.Description
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 課税区分によって金額から
'                           本体価格と消費税（指定年月での税率を用いる）を算出する
'        MODULE_ID        : MSZZ004_M10
'        CREATE_DATE      : 2007/03/15          K.ISHIZAKA
'        PARAM            : lngKingaku          金額(I)
'                         : strDate             算出年月(I) yyyymm
'                         : strKazeiKbn         課税区分(I) 1:外税 2:内税 3:非課税
'                         : lngPrice            本体価格(O) Long
'                         : lngTax              消費税(O)   Long
'                         : [bClr]              省略可：グローバル領域のクリア（CONF_TABL再読込）
'                                                       する(True)／しない(False) Default=False
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ004_M10(ByVal lngKingaku As Long, ByVal strDate As String, ByVal strKazeiKbn As String, ByRef lngPrice As Long, ByRef lngTax As Long, Optional bClr As Boolean = False)
    Dim dblRate             As Double
    Dim strRoundKbn         As String
    On Error GoTo ErrorHandler

    dblRate = MSZZ004_M20(strDate, strRoundKbn, bClr)
    Select Case strKazeiKbn
    Case "1"    '外税
        lngPrice = lngKingaku
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
    Case "2"    '内税
        lngTax = fncPriceInTax(lngKingaku, dblRate, strRoundKbn, lngPrice)
    Case "3"    '非課税
        lngPrice = lngKingaku
        lngTax = 0
    End Select
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M10" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 課税区分によって金額から
'                           本体価格と消費税（指定税率を用いる）を算出する
'        MODULE_ID        : MSZZ004_M11
'        CREATE_DATE      : 2014/04/01          K.ISHIZAKA
'        PARAM            : lngKingaku          金額(I)
'                         : dblRate             税率(I) Double 0.05 とか
'                         : strKazeiKbn         課税区分(I) 1:外税 2:内税 3:非課税
'                         : lngPrice            本体価格(O) Long
'                         : lngTax              消費税(O)   Long
'                         : [bClr]              省略可：グローバル領域のクリア（CONF_TABL再読込）
'                                                       する(True)／しない(False) Default=False
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ004_M11(ByVal lngKingaku As Long, ByVal dblRate As Double, ByVal strKazeiKbn As String, ByRef lngPrice As Long, ByRef lngTax As Long, Optional bClr As Boolean = False)
    Dim dblDummy            As Double
    Dim strRoundKbn         As String
    On Error GoTo ErrorHandler

    dblDummy = MSZZ004_M20("999912", strRoundKbn, bClr)
    Select Case strKazeiKbn
    Case "1"    '外税
        lngPrice = lngKingaku
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
    Case "2"    '内税
        lngTax = fncPriceInTax(lngKingaku, dblRate, strRoundKbn, lngPrice)
    Case "3"    '非課税
        lngPrice = lngKingaku
        lngTax = 0
    End Select
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M11" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 課税区分によって金額から
'                           本体価格と消費税（指定年月での税率を用いる）を算出する
'        MODULE_ID        : MSZZ004_M12
'        CREATE_DATE      : 2014/04/17          K.ISHIZAKA
'        PARAM            : lngKingaku          金額(I)
'                         : strDate             算出年月(I) yyyymm
'                         : strKazeiKbn         課税区分(I) 1:外税 2:内税 3:非課税
'                         : strRoundKbn         端数区分(I) 0:四捨五入 1:切捨て 2:切上げ
'                         : lngPrice            本体価格(O) Long
'                         : lngTax              消費税(O)   Long
'                         : [bClr]              省略可：グローバル領域のクリア（CONF_TABL再読込）
'                                                       する(True)／しない(False) Default=False
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ004_M12(ByVal lngKingaku As Long, ByVal strDate As String, ByVal strKazeiKbn As String, ByVal strRoundKbn As String, ByRef lngPrice As Long, ByRef lngTax As Long, Optional bClr As Boolean = False)
    Dim dblRate             As Double
    Dim strDummy            As String
    On Error GoTo ErrorHandler

    dblRate = MSZZ004_M20(strDate, strDummy, bClr)
    Select Case strKazeiKbn
    Case "1"    '外税
        lngPrice = lngKingaku
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
    Case "2"    '内税
        lngTax = fncPriceInTax(lngKingaku, dblRate, strRoundKbn, lngPrice)
    Case "3"    '非課税
        lngPrice = lngKingaku
        lngTax = 0
    End Select
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M12" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 外税での消費税計算
'        MODULE_ID        : fncPriceOutTax
'        CREATE_DATE      : 2007/03/15          K.ISHIZAKA
'        PARAM            : lngPrice            本体価格(I)
'                         : dblRate             税率(I)
'                         : strRoundKbn         端数区分(I) 0:四捨五入 1:切捨て 2:切上げ
'        RETURN           : 消費税(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncPriceOutTax(ByVal lngPrice As Long, ByVal dblRate As Double, ByVal strRoundKbn As String) As Long
    On Error GoTo ErrorHandler
    
    fncPriceOutTax = fncRound(lngPrice * dblRate, strRoundKbn)
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncPriceOutTax" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      : 内税での消費税計算
'        MODULE_ID        : fncPriceInTax
'        CREATE_DATE      : 2007/03/15          K.ISHIZAKA
'        PARAM            : lngPriceInTax       税込み金額(I)
'                         : dblRate             税率(I)
'                         : strRoundKbn         端数区分(I) 0:四捨五入 1:切捨て 2:切上げ
'                         : lngPrice            本体価格(O) Long
'        RETURN           : 消費税(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncPriceInTax(ByVal lngPriceInTax As Long, ByVal dblRate As Double, ByVal strRoundKbn As String, ByRef lngPrice As Long) As Long
    Dim lngTax              As Long
    Dim lngTax2             As Long
    On Error GoTo ErrorHandler

    '①想定消費税
    lngTax = fncRound(lngPriceInTax * dblRate / (1 + dblRate), strRoundKbn)
    '②想定本体価格
    lngPrice = lngPriceInTax - lngTax
'****                                                                           'DELETE START 2010/06/30 K.ISHIZAKA
'    '③想定本体価格から求めた消費税
'    lngTax2 = fncPriceOutTax(lngPrice, dblRate, strRoundKbn)
'    '[引数:税込み金額] ＜ （② ＋ ③）の場合は本体価格を小さくする
'    If Abs(lngTax) < Abs(lngTax2) Then
'        lngTax = lngTax2
'        lngPrice = lngPriceInTax - lngTax
'    End If
'****                                                                           'DELETE END   2010/06/30 K.ISHIZAKA
    fncPriceInTax = lngTax
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncPriceInTax" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      : 端数処理
'        MODULE_ID        : fncRound
'        CREATE_DATE      : 2007/03/15          K.ISHIZAKA
'        PARAM            : dblWork             金額(I)
'                         : strRoundKbn         端数区分(I) 0:四捨五入 1:切捨て 2:切上げ
'        RETURN           : 端数処理した金額(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRound(ByVal dblWork As Double, ByVal strRoundKbn As String) As Long
    Dim dblFix              As Double
    Dim dblSgn              As Double
    On Error GoTo ErrorHandler

    '絶対値
    dblSgn = Sgn(dblWork)
    dblWork = Abs(dblWork)
    '切捨て値
    dblFix = Fix(dblWork)
    Select Case strRoundKbn
    Case "0"    '四捨五入
        If dblWork >= (dblFix + 0.5) Then
            dblFix = dblFix + 1
        End If
    Case "1"    '切捨て
    Case "2"    '切上げ
        If dblWork > dblFix Then
            dblFix = dblFix + 1
        End If
    Case Else
        Call MSZZ024_M10("Round", "端数区分が不正です。")
    End Select
    '符号を戻して返す
    fncRound = CLng(dblFix * dblSgn)
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncRound" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      : 税率取得処理
'        MODULE_ID        : MSZZ004_M20
'        CREATE_DATE      : 2007/03/15          K.ISHIZAKA
'        PARAM            : strDate             算出年月(I) yyyymm
'                         : strRoundKbn         端数区分(O) 0:四捨五入 1:切捨て 2:切上げ
'                         : [bClr]              省略可：グローバル領域のクリア(I) する(True)／しない(False)
'        RETURN           : 税率(Double)        0.05 とか
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ004_M20(ByVal strDate As String, ByRef strRoundKbn As String, Optional bClr As Boolean = False) As Double
    Static gstrKirikaeDate  As String
    Static gdblZeiRituOld   As Double
    Static gdblZeiRituNew   As Double
    Static gstrRoundKubun   As String
    On Error GoTo ErrorHandler
    
    If (gstrKirikaeDate = "") Or bClr Then
        Dim objCon          As Object
        Dim objRst          As Object
        Dim strSQL          As String
        
        strSQL = SELECT_CONF_TABL()
        Set objCon = ADODB_Connection()
        On Error GoTo ErrorHandler2
        Set objRst = ADODB_Recordset(strSQL, objCon)
        On Error GoTo ErrorHandler3
        With objRst
            If .EOF Then
                Call MSZZ024_M10(strSQL, "[CONF_TABL]の設定不足です。")
            End If
            gstrKirikaeDate = .Fields("CONFT_SYYMD")
            gdblZeiRituOld = .Fields("CONFT_SYOLR")
            gdblZeiRituNew = .Fields("CONFT_SYNWR")
            gstrRoundKubun = .Fields("CONFT_SYKBI")
            .Close
        End With
        On Error GoTo ErrorHandler2
        objCon.Close
        On Error GoTo ErrorHandler
    End If
    strRoundKbn = gstrRoundKubun
    MSZZ004_M20 = IIf(Left(strDate, 6) < gstrKirikaeDate, gdblZeiRituOld, gdblZeiRituNew)
Exit Function

ErrorHandler3:
    objRst.Close
ErrorHandler2:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M20" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      : CONF_TABL取得ＳＱＬ
'        MODULE_ID        : SELECT_CONF_TABL
'        CREATE_DATE      : 2007/03/15          K.ISHIZAKA
'        RETURN           : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SELECT_CONF_TABL() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CONFT_SYYMD,"
    strSQL = strSQL & " CONFT_SYOLR,"
    strSQL = strSQL & " CONFT_SYNWR,"
    strSQL = strSQL & " CONFT_SYKBI "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " CONF_TABL "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CONFT_NUMBC = '1'"
    strSQL = strSQL & ";"
    
    SELECT_CONF_TABL = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      : 端数処理
'        MODULE_ID        : MSZZ004_M30
'        CREATE_DATE      : 2009/05/19          S.SHIBAZAKI
'        PARAM            : dblWork             金額(I)
'                         : strRoundKbn         端数区分(I) 0:四捨五入 1:切捨て 2:切上げ
'        RETURN           : 端数処理した金額(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ004_M30(ByVal dblWork As Double, ByVal strRoundKbn As String) As Long
    On Error GoTo ErrorHandler
    
    MSZZ004_M30 = fncRound(dblWork, strRoundKbn)
    
    Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M30" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      : 消費税計算
'                           本体価格から消費税と合計を算出
'        MODULE_ID        : MSZZ004_M40
'        CREATE_DATE      : 2007/04/10          YAMA
'        PARAM            : lngKingaku          金額(I)
'                         : strDate             算出年月(I) yyyymm
'                         : strKazeiKbn         課税区分(I) 1:外税 2:内税 3:非課税
'                         : lngTotalPrice       合計価格(O) Long
'                         : lngTax              消費税(O)   Long
'                         : [bClr]              省略可：グローバル領域のクリア（CONF_TABL再読込）
'                                                       する(True)／しない(False) Default=False
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ004_M40(ByVal lngKingaku As Long, ByVal strDate As String, ByVal strKazeiKbn As String, ByRef lngTotalPrice As Long, ByRef lngTax As Long, Optional bClr As Boolean = False)
    Dim dblRate             As Double
    Dim strRoundKbn         As String
    On Error GoTo ErrorHandler
    
    dblRate = MSZZ004_M20(strDate, strRoundKbn, bClr)
    Select Case strKazeiKbn
    Case "1"    '外税
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
        lngTotalPrice = lngKingaku
    Case "2"    '内税
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
        lngTotalPrice = lngKingaku + lngTax
    Case "3"    '非課税
        lngTotalPrice = lngKingaku
        lngTax = 0
    End Select
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M40" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 消費税計算
'                           本体価格から消費税と合計を算出
'        MODULE_ID        : MSZZ004_M41
'        CREATE_DATE      : 2014/04/01          K.ISHIZAKA
'        PARAM            : lngKingaku          金額(I)
'                         : dblRate             税率(I) Double 0.05 とか
'                         : strKazeiKbn         課税区分(I) 1:外税 2:内税 3:非課税
'                         : lngTotalPrice       合計価格(O) Long
'                         : lngTax              消費税(O)   Long
'                         : [bClr]              省略可：グローバル領域のクリア（CONF_TABL再読込）
'                                                       する(True)／しない(False) Default=False
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ004_M41(ByVal lngKingaku As Long, ByVal dblRate As Double, ByVal strKazeiKbn As String, ByRef lngTotalPrice As Long, ByRef lngTax As Long, Optional bClr As Boolean = False)
    Dim dblDummy            As Double
    Dim strRoundKbn         As String
    On Error GoTo ErrorHandler
    
    dblDummy = MSZZ004_M20("999912", strRoundKbn, bClr)
    Select Case strKazeiKbn
    Case "1"    '外税
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
        lngTotalPrice = lngKingaku
    Case "2"    '内税
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
        lngTotalPrice = lngKingaku + lngTax
    Case "3"    '非課税
        lngTotalPrice = lngKingaku
        lngTax = 0
    End Select
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M41" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 消費税計算
'                           本体価格から消費税と合計を算出
'        MODULE_ID        : MSZZ004_M42
'        CREATE_DATE      : 2014/04/17          K.ISHIZAKA
'        PARAM            : lngKingaku          金額(I)
'                         : strDate             算出年月(I) yyyymm
'                         : strKazeiKbn         課税区分(I) 1:外税 2:内税 3:非課税
'                         : strRoundKbn         端数区分(I) 0:四捨五入 1:切捨て 2:切上げ
'                         : lngTotalPrice       合計価格(O) Long
'                         : lngTax              消費税(O)   Long
'                         : [bClr]              省略可：グローバル領域のクリア（CONF_TABL再読込）
'                                                       する(True)／しない(False) Default=False
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ004_M42(ByVal lngKingaku As Long, ByVal strDate As String, ByVal strKazeiKbn As String, ByVal strRoundKbn As String, ByRef lngTotalPrice As Long, ByRef lngTax As Long, Optional bClr As Boolean = False)
    Dim dblRate             As Double
    Dim strDummy            As String
    On Error GoTo ErrorHandler
    
    dblRate = MSZZ004_M20(strDate, strDummy, bClr)
    Select Case strKazeiKbn
    Case "1"    '外税
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
        lngTotalPrice = lngKingaku
    Case "2"    '内税
        lngTax = fncPriceOutTax(lngKingaku, dblRate, strRoundKbn)
        lngTotalPrice = lngKingaku + lngTax
    Case "3"    '非課税
        lngTotalPrice = lngKingaku
        lngTax = 0
    End Select
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ004_M42" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended or program ********************************
