Attribute VB_Name = "MNR100"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : インターネット予約
'
'        PROGRAM_NAME    : インターネット予約用パス管理
'        PROGRAM_ID      : MNR100
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/08/14
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          : 2007/09/01
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.1
'                        : ヤード条件は営業終了日まで
'                          コンテナ条件は営業終了日が設定されていないに修正
'
'        UPDATE          : 2007/09/19
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.2
'                        : ネット割引サービスは３ヶ月以上ご利用の場合のみ
'
'        UPDATE          : 2008/10/06
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.3
'                        : ２０００円割引の停止
'                          ※もとに戻す場合があってもVer0.4は適用すること
'                         （解約日とサービス適用期間とは別で考えないといけない！！）
'
'        UPDATE          : 2008/10/14
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.4
'                        : 解約日設定時にサービス適用期間を使用しない
'
'        UPDATE          : 2009/04/18
'        UPDATER         : tajima
'        Ver             : 0.5
'                        : ２０００円割引を期間指定で制御出来るようにする
'
'        UPDATE          : 2009/06/30
'        UPDATER         : hirano
'        Ver             : 0.6
'                        : 事務手数料割引対応
'
'        UPDATE          : 2009/12/02
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.7
'                        : 料金差額の設定されていないコンテナだけを対象とする
'
'        UPDATE          : 2010/09/09
'        UPDATER         : M.RYU
'        Ver             : 0.8
'                        : 用途区分=99のデータを除外する
'
'        UPDATE          : 2010/09/17
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.9
'                        : Ver0.8の金額チェックはなし
'
'        UPDATE          : 2012/07/18
'        UPDATER         : M.HONDA
'        Ver             : 1.0
'                        : netforest -> ﾗﾋﾞｯﾄｻｲﾄにｻｰﾊﾞｰ変更に伴いﾊﾟｽを変更
'
'        UPDATE          : 2013/01/11
'        UPDATER         : M.HONDA
'        Ver             : 1.1
'                        : コンテナ取得条件に柱があるものを除く条件を追加
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const MODULE_ID     As String = "MNR100"
Public Const C_TAB_INDEX    As Long = 4

'▼2009/04/18 ２０００円割引復活
'       ２０００円割引の停止                                                    'DELETE START 2008/10/06 K.ISHIZAKA
Public Const C_YARD_SEV1N   As String = "ネット割引サービス：使用料から２，０００円引き"
Public Const C_YARD_SEV2N   As String = "他のキャンペーン・サービスとは併用できません"
Public Const C_YARD_SEV3N   As String = "※３ヶ月以上ご利用の場合のみ"          'INSERT 2007/09/19 K.ISHIZAKA
Public Const C_YARD_SEV_EXMONTH  As Long = 3 'キャンペーン適用満了月数          'INSERT 2007/09/19 K.ISHIZAKA
'                                                                               'DELETE END   2008/10/06 K.ISHIZAKA
'▲2009/04/18
Public Const C_USE_PERIOD   As String = "1" 'ご利用期間：２ヶ月以下の区分       'INSERT 2007/09/19 K.ISHIZAKA
Public Const C_USE_MONTH    As Long = 2     'ご利用期間：２ヶ月以下の月数       'INSERT 2008/10/14 K.ISHIZAKA

'▼2009/04/18 ２０００円割引復活
'       ２０００円割引の停止                                                    'DELETE START 2008/10/06 K.ISHIZAKA
Public Const C_RCPT_ADD_EZAPPI_CODE1    As Long = 57
Public Const C_RCPT_ADD_EZAPPI1         As Long = -2000
'                                                                               'DELETE END   2008/10/06 K.ISHIZAKA
Private P_2kSEV_OUT_FROM    As String       '2000円割引提示開始日
Private P_2kSEV_OUT_TO      As String       '2000円割引提示終了日
Private P_2kSEV_GET_FROM    As String       '2000円割引適用開始日
Private P_2kSEV_GET_TO      As String       '2000円割引適用終了日
'▲2009/04/18
'2009/06/30 INS <S> hirano
Public P_GET_OFFICE_FEE As String
Private P_GET_OFFICE_FEE_FROM As String
Private P_GET_OFFICE_FEE_TO   As String
'2009/06/30 INS <E> hirano
Private P_LOCAL_ROOT        As String       'インターネット予約用の社内ベースフォルダー
Private P_REMOTE_ROOT       As String       'インターネット予約用の公開ベースフォルダー

Public Enum NetReceptType
    LocalPath
    RemotePath
End Enum

Public Enum NetReceptPath
    P_ADDR
    P_LINE
    P_YARD
    P_RSA_PUBLIC
    P_RSA_PRIVATE
    P_YOUK
    P_PDF
    P_MAIL
    P_DLVR
End Enum

'==============================================================================*
'
'       MODULE_NAME     : インターネット用パス取得
'       MODULE_ID       : GetNetReceptPath
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : eType                 Local / Remote(I)
'                       : ePath                 ADDR (I)
'       RETURN          : XMLタグ(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetNetReceptPath(ByVal eType As NetReceptType, ByVal ePath As NetReceptPath) As String
    Dim strPath         As String
    On Error GoTo ErrorHandler
        
    Call Select_INTI_FILE
    Select Case ePath
    Case P_ADDR
        '' 2012/07/08 M.HONDA
        strPath = "html\rent\data\map\addr\"
        '' strPath = "htdocs\rent\data\map\addr\"
    Case P_LINE
        '' 2012/07/08 M.HONDA
        strPath = "html\rent\data\map\line\"
        '' strPath = "htdocs\rent\data\map\line\"
    Case P_YARD
        '' 2012/07/08 M.HONDA
        strPath = "html\rent\data\YARD_MAST\"
        '' strPath = "htdocs\rent\data\YARD_MAST\"
        
'旧テスト環境
    Case P_RSA_PUBLIC
        strPath = "data\rent\PUBLIC_KEY\"
    Case P_RSA_PRIVATE
        strPath = "data\rent\PRIVATE_KEY\"
    Case P_YOUK
        strPath = "data\rent\YOUK_TRAN\"
    Case P_PDF
        strPath = "data\rent\PDF\"
    Case P_MAIL
        strPath = "data\rent\mail\"
    Case P_DLVR
        strPath = "data\rent\DLVR\"
    End Select
        
'    Case P_RSA_PUBLIC
'        strPath = "www.kase3535.com\data\rent\PUBLIC_KEY\"
'    Case P_RSA_PRIVATE
'        strPath = "www.kase3535.com\data\rent\PRIVATE_KEY\"
'    Case P_YOUK
'        strPath = "www.kase3535.com\data\rent\YOUK_TRAN\"
'    Case P_PDF
'        strPath = "www.kase3535.com\data\rent\PDF\"
'    Case P_MAIL
'        strPath = "www.kase3535.com\data\rent\mail\"
'    Case P_DLVR
'        strPath = "www.kase3535.com\data\rent\DLVR\"
'    End Select

    If eType = LocalPath Then
        strPath = P_LOCAL_ROOT & strPath
        Call MkDirEx(strPath)
    Else
        strPath = P_REMOTE_ROOT & Replace(strPath, "\", "/")
    End If
    GetNetReceptPath = strPath
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetNetReceptPath" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ベースパス取得
'       MODULE_ID       : Select_INTI_FILE
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Select_INTI_FILE()
    On Error GoTo ErrorHandler
    
    P_LOCAL_ROOT = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='LOCAL'"))
    If P_LOCAL_ROOT = "" Then
        Call MSZZ024_M10("INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECDB='LOCAL'")
    End If
    P_REMOTE_ROOT = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='REMOTE'"))
    If P_REMOTE_ROOT = "" Then
        Call MSZZ024_M10("INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECDB='REMOTE'")
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_INTI_FILE" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : インターネット用パス変換
'       MODULE_ID       : ChangeNetReceptPath
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strPath               パス(I)
'       RETURN          : 変換パス(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ChangeNetReceptPath(ByVal strPath As String) As String
    On Error GoTo ErrorHandler
    
    If InStr(1, strPath, P_LOCAL_ROOT) > 0 Then
        strPath = Replace(strPath, P_LOCAL_ROOT, P_REMOTE_ROOT)
        strPath = Replace(strPath, "\", "/")
    ElseIf InStr(1, strPath, P_REMOTE_ROOT) > 0 Then
        strPath = Replace(strPath, P_REMOTE_ROOT, P_LOCAL_ROOT)
        strPath = Replace(strPath, "/", "\")
    End If
    ChangeNetReceptPath = strPath
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ChangeNetReceptPath" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : フォルダー作成
'       MODULE_ID       : MkDirEx
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strPath               パス(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MkDirEx(ByVal strPath As String)
    Dim strSplit()          As String
    On Error GoTo ErrorHandler

    If Dir(strPath, vbDirectory) = "" Then
        strSplit = Split(strPath, "\")
        ReDim Preserve strSplit(UBound(strSplit) - IIf(strSplit(UBound(strSplit)) = "", 2, 1))
        Call MkDirEx(Join(strSplit, "\"))
        MkDir strPath
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MkDirEx" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ノードタグ出力
'       MODULE_ID       : outputNodes
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objRst                レコードセット(I)
'                       : file                  ファイルID(I)
'                       : [i]                   インデックス(I)
'                                               省略時：先頭項目を親ノードとし以降を子ノードにする
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub outputNodes(objRst As Object, ByVal file As Integer, Optional i As Long = 0)
    On Error GoTo ErrorHandler
    
    If i > 1 Then
        Call outputTags(objRst, file, i)
    Else
        Dim strSave     As String
        
        strSave = objRst.Fields(i).VALUE
        With objRst.Fields(i)
            UTF8Write file, Space(C_TAB_INDEX * (i + 1)) & "<" & .NAME & ">"
            UTF8Write file, Space(C_TAB_INDEX * (i + 2)) & "<name>" & Format(.VALUE) & "</name>"
        End With
        While strSave = objRst.Fields(i).VALUE
            Call outputNodes(objRst, file, i + 1)
            If Not objRst.EOF Then
                objRst.MoveNext
            End If
            If objRst.EOF Then
                UTF8Write file, Space(C_TAB_INDEX * (i + 1)) & "</" & objRst.Fields(i).NAME & ">"
                Exit Sub
            End If
        Wend
        UTF8Write file, Space(C_TAB_INDEX * (i + 1)) & "</" & objRst.Fields(i).NAME & ">"
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "outputNodes" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : タグ出力
'       MODULE_ID       : outputTags
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objRst                レコードセット(I)
'                       : file                  ファイルID(I)
'                       : i                     インデックス(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub outputTags(objRst As Object, ByVal file As Integer, ByVal i As Long)
    Dim j                   As Long
    On Error GoTo ErrorHandler
    
    For j = i To objRst.Fields.Count - 1
        With objRst.Fields(j)
            If (.Type = adChar) Or (.Type = adVarChar) Then
                If Nz(.VALUE, "") <> "" Then
                    UTF8Write file, Space(C_TAB_INDEX * (i + 1)) & "<" & .NAME & ">" & .VALUE & "</" & .NAME & ">"
                End If
            Else
                If Not IsNull(.VALUE) Then
                    UTF8Write file, Space(C_TAB_INDEX * (i + 1)) & "<" & .NAME & ">" & Format(.VALUE) & "</" & .NAME & ">"
                End If
            End If
        End With
    Next
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "outputTags" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 都道府県名取得
'       MODULE_ID       : getPrefecture
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strNO                 都道府県番号(I)
'       RETURN          : 都道府県名(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function getPrefecture(ByVal strNO As String) As String
    Dim objConn             As Object
    Dim objRst              As Object
    On Error GoTo ErrorHandler

    Set objConn = ADODB_Connection("MAP")
    On Error GoTo ErrorHandler1

    Set objRst = ADODB_Recordset(SelectTodo(strNO), objConn)
    On Error GoTo ErrorHandler2
    With objRst
        If Not .EOF Then
            getPrefecture = Trim(.Fields(0).VALUE)
        Else
            getPrefecture = ""
        End If
        .Close
        On Error GoTo ErrorHandler1
    End With

    objConn.Close
    On Error GoTo ErrorHandler
Exit Function

ErrorHandler2:
    objRst.Close
ErrorHandler1:
    objConn.Close
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "getPrefecture" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 都道府県名取得
'       MODULE_ID       : SelectTodo
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strNO                 都道府県番号(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SelectTodo(ByVal strNO As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " TODOM_TODON AS prefecture "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " TODO_MAST "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " TODOM_TODOI = '" & strNO & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " TODOM_SEIBY IS NOT NULL "

    SelectTodo = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードマスタ取得条件
'       MODULE_ID       : NetReceptWhereYardMast
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function NetReceptWhereYardMast(Optional bWhere As Boolean = True) As String
    Dim strSQL              As String

    If bWhere Then
        strSQL = strSQL & "WHERE"
        strSQL = strSQL & " CONT_KEY='1' "
    End If
    strSQL = strSQL & "AND"
    strSQL = strSQL & " YARD_IDO IS NOT NULL "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " YARD_KEIDO IS NOT NULL "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " ISNULL(YARD_BEGIN_DAY, '1900/01/01') <= GETDATE() "
    strSQL = strSQL & "AND"
    If bWhere Then                                                              'INSERT 2007/09/01 K.ISHIZAKA
        strSQL = strSQL & " ISNULL(YARD_RENTEND_DAY,'9999/12/31') >= GETDATE() " 'INSERT 2007/09/01 K.ISHIZAKA
    Else                                                                        'INSERT 2007/09/01 K.ISHIZAKA
        strSQL = strSQL & " YARD_RENTEND_DAY IS NULL "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " ISNULL(YARD_INLIMIT_DAY,'1900/01/01') < GETDATE() -1 "
    End If                                                                      'INSERT 2007/09/01 K.ISHIZAKA

    NetReceptWhereYardMast = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナマスタ取得条件
'       MODULE_ID       : NetReceptWhereCntaMast
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function NetReceptWhereCntaMast() As String
    Dim strSQL              As String

    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CONT_KEY='1' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USE = 1 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE IS NOT NULL "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_SIZE  IS NOT NULL "
    
    '--↓↓---2010/09/09-----ryu-----add----------<s>--'用途区分の99のデータを外す
    strSQL = strSQL & " AND CNTA_USAGE != 99 "
'    strSQL = strSQL & " AND PRIC_PRICE IS NOT NULL "                           'DELETE 2010/09/17 K.ISHIZAKA
    '--↑↑---2010/09/09-----ryu-----add----------<e>
    
    '' --　2013/01/11 M.HONDA START
    '--料金差額の設定されていないコンテナだけを対象とする
    strSQL = strSQL & "AND"                                                     'INSERT 2009/12/02 K.ISHIZAKA
    'strSQL = strSQL & " ISNULL(CNTA_PRICE_DIFF,0) = 0 "                         'INSERT 2009/12/02 K.ISHIZAKA
    strSQL = strSQL & " ( ISNULL(CNTA_PRICE_DIFF,0) = 0 AND CNTA_REASON NOT LIKE '%柱%' )"
    '' --　2013/01/11 M.HONDA END
    
    '--契約されていないこと
    strSQL = strSQL & "AND NOT EXISTS"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " CARG_FILE "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CARG_YCODE  = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_NO     = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_CONTNA = 0 "
    strSQL = strSQL & ")"
    '--取り置き、契約受付されていないこと
    strSQL = strSQL & "AND NOT EXISTS"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " INTR_TRAN "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " INTRT_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_NO    = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_INTROKBN IN(1, 2) "
    strSQL = strSQL & ")"
    '--ヤードが閉鎖されていないこと
    strSQL = strSQL & "AND EXISTS"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " YARD_MAST "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " YARD_CODE = CNTA_CODE "
    strSQL = strSQL & NetReceptWhereYardMast(False)
'    strSQL = strSQL & "AND"
'    strSQL = strSQL & " ISNULL(YARD_NETUSE_KBN,0) != 0 "
    strSQL = strSQL & ")"

    NetReceptWhereCntaMast = strSQL
End Function

'****************************  2009/04/18 add **********************************
'==============================================================================*
'
'       MODULE_NAME     : ２千円引きサービス提示期間の取得
'       MODULE_ID       : Get2kDiscountServicePreRange
'       CREATE_DATE     : 2009/04/18 byPEGA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Get2kDiscountServiceOutRange()
    On Error GoTo ErrorHandler
    
    P_2kSEV_OUT_FROM = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='2kSEV_OUT_FROM'"))
    If P_2kSEV_OUT_FROM = "" Then
        P_2kSEV_OUT_FROM = "0" '未設定の場合はサービスしないようにする
    End If
    P_2kSEV_OUT_TO = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='2kSEV_OUT_TO'"))
    If P_2kSEV_OUT_TO = "" Then
        P_2kSEV_OUT_TO = "0" '未設定の場合はサービスしないようにする
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Get2kDiscountServiceOutRange" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'       MODULE_NAME     : ２千円引きサービス適用期間の取得
'       MODULE_ID       : Get2kDiscountServiceGetRange
'       CREATE_DATE     : 2009/04/18 byPEGA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Get2kDiscountServiceGetRange()
    On Error GoTo ErrorHandler
    
    P_2kSEV_GET_FROM = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='2kSEV_GET_FROM'"))
    If P_2kSEV_GET_FROM = "" Then
        P_2kSEV_GET_FROM = "0" '未設定の場合はサービスしないようにする
    End If
    P_2kSEV_GET_TO = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='2kSEV_GET_TO'"))
    If P_2kSEV_GET_TO = "" Then
        P_2kSEV_GET_TO = "0" '未設定の場合はサービスしないようにする
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Get2kDiscountServiceGetRange" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'       MODULE_NAME     : ２千円引きサービスの提示可否確認
'       MODULE_ID       : Is2kDiscountServiceOut
'       CREATE_DATE     : 2009/04/18 byPEGA
'       RETURN          : true...サービス期間中、false...サービス期間外
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Is2kDiscountServiceOut() As Boolean
    On Error GoTo ErrorHandler
    
    Dim dateText As String
    
    Is2kDiscountServiceOut = False

    dateText = Format$(Now, "yyyymmddhhmmss")
    
    '設定した期間内か否か
    If P_2kSEV_OUT_FROM <= dateText And dateText <= P_2kSEV_OUT_TO Then
        Is2kDiscountServiceOut = True
    End If

Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Is2kDiscountServiceOut" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'       MODULE_NAME     : ２千円引きサービスの適用可否確認
'       MODULE_ID       : Is2kDiscountServiceGet
'       CREATE_DATE     : 2009/04/18 byPEGA
'       RETURN          : true...サービス期間中、false...サービス期間外
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Is2kDiscountServiceGet() As Boolean
    On Error GoTo ErrorHandler
    
    Dim dateText As String
    
    Is2kDiscountServiceGet = False

    dateText = Format$(Now, "yyyymmddhhmmss")
    
    '設定した期間内か否か
    If P_2kSEV_GET_FROM <= dateText And dateText <= P_2kSEV_GET_TO Then
        Is2kDiscountServiceGet = True
    End If

Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Is2kDiscountServiceGet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'       MODULE_NAME     : 事務手数料適用情報の取得
'       MODULE_ID       : GetOffiveFeeGetRange
'       CREATE_DATE     : 2009/06/30 by　hirano
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub GetOffiveFeeGetRange()
    On Error GoTo ErrorHandler
    
    P_GET_OFFICE_FEE = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='OFFICE_FEE'"))
    If P_GET_OFFICE_FEE = "" Then
        P_GET_OFFICE_FEE = "5000" '未設定の場合は事務手数料は5000円固定
    End If
    P_GET_OFFICE_FEE_FROM = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='OFFICE_FEE_FROM'"))
    If P_GET_OFFICE_FEE_FROM = "" Then
        P_GET_OFFICE_FEE_FROM = "0" '未設定の場合はサービスしないようにする
    End If
    P_GET_OFFICE_FEE_TO = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='OFFICE_FEE_TO'"))
    If P_GET_OFFICE_FEE_TO = "" Then
        P_GET_OFFICE_FEE_TO = "0" '未設定の場合はサービスしないようにする
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetOffiveFeeGetRange" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'       MODULE_NAME     : 事務手数料の適用可否確認
'       MODULE_ID       : IsOffiveFeeGet
'       CREATE_DATE     : 2009/06/30 by hirano
'       RETURN          : true...サービス期間中、false...サービス期間外
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function IsOffiveFeeGet() As Boolean
    On Error GoTo ErrorHandler
    
    Dim dateText As String
    
    IsOffiveFeeGet = False

    dateText = Format$(Now, "yyyymmddhhmmss")
    
    '設定した期間内か否か
    If P_GET_OFFICE_FEE_FROM <= dateText And dateText <= P_GET_OFFICE_FEE_TO Then
        IsOffiveFeeGet = True
    End If

Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Is2kDiscountServiceGet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended or program ********************************
