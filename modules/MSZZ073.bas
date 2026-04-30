Attribute VB_Name = "MSZZ073"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : ReCaRo8接続
'       PROGRAM_ID      : MSZZ073
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2017/01/26
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          : 2018/01/13
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : レンタルサーバー側の、2017/09/14「nginx」導入による対応
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID       As String = "MSZZ073"

Private Const C_BASE64_TABLE = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

'==============================================================================*
'   テスト
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub aaa()
    Dim strXml              As String
    Dim strResponseXml      As String
    Dim iniFile             As Collection
    Dim objDoc              As Object
    Dim objTbl              As Object
    Dim iv                  As String
    Dim key                 As String
    Dim objNd               As Object

'    strXml = ""
'    strXml = strXml & "<?xml version=""1.0"" encoding=""UTF-8""?>"
'    strXml = strXml & "<root>"
'    strXml = strXml & "<table ID=""LGIN_TABL"">"
'    strXml = strXml & "<where>"
'    strXml = strXml & "<LGINT_MAILB type=""varchar"">kiyoshi_ishizaka@kasegroup.co.jp</LGINT_MAILB>"
'    strXml = strXml & "</where>"
'    strXml = strXml & "</table>"
'    strXml = strXml & "<table ID=""USER_MAST"">"
'    strXml = strXml & "<where>"
'    strXml = strXml & "<USERM_MAILB type=""varchar"">kiyoshi_ishizaka@kasegroup.co.jp</USERM_MAILB>"
'    strXml = strXml & "</where>"
'    strXml = strXml & "</table>"
'    strXml = strXml & "</root>"
    
    Set objDoc = recaro8_createDocument()
    Set objTbl = recaro8_createTable(objDoc, "LGIN_TABL")
    Call recaro8_whereColumn(objTbl, "LGINT_MAILB", "kiyoshi_ishizaka@kasegroup.co.jp", "varchar")
    Set objTbl = recaro8_createTable(objDoc, "USER_MAST")
    Call recaro8_whereColumn(objTbl, "USERM_MAILB", "kiyoshi_ishizaka@kasegroup.co.jp", "varchar")
    
    Set iniFile = get_ini_file()
    If recaro8_table_ctrl(objDoc, strResponseXml, iniFile) Then
        Debug.Print "OK:"
        'Debug.Print strResponseXml
        Set objDoc = XmlCreateDocument(strResponseXml)
        
        iv = iniFile("INITIAL_VECTOR")
        key = objDoc.selectSingleNode("root/table[@ID=""LGIN_TABL""]/row/LGINT_TOKEN").Text
        Set objNd = objDoc.selectSingleNode("root/table[@ID=""USER_MAST""]/row/USERM_USERN")
        Debug.Print AesDecrypt(key, iv, objNd.Text)
    Else
        Debug.Print "NG:" & strResponseXml
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : Base64の乱数文字列を生成する
'       MODULE_ID       : random_pseudo_base64
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : intLength             長さ(I)
'       RETURN          : 文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function random_pseudo_base64(ByVal intLength As Integer) As String
    Dim strResult           As String
    Dim iMax                As Integer
    Dim i                   As Integer
    On Error GoTo ErrorHandler

    iMax = Len(C_BASE64_TABLE)
    For i = 1 To intLength
        Randomize
        strResult = Mid(C_BASE64_TABLE, Int((iMax * Rnd) + 1), 1) & strResult
    Next
    random_pseudo_base64 = strResult
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "random_pseudo_base64" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ノードのテキスト文字列を取得する
'       MODULE_ID       : getNodeText
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : objRow                行オブジェクト(I) MSXML2.IXMLDOMElement
'                       : strColName            カラム名(I)
'                       : [strQuotationMarks]   引用符(I)
'                       : [zeroByteToNull]      ラカ文字のときnullを返すかどうか(False:カラ文字／True:null)
'       RETURN          : テキスト文字列(Object) MSXML2.DOMDocument
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function getNodeText(objRow As Object, ByVal strColName As String, Optional ByVal strQuotationMarks As String = "", Optional ByVal zeroByteToNull As Boolean = False) As String
    Dim strValue            As String
    On Error GoTo ErrorHandler
    
    strValue = objRow.selectSingleNode(strColName).Text
    strValue = Replace(strValue, Chr(16), "")
    If zeroByteToNull And (Len(strValue) = 0) Then
        getNodeText = "null"
    Else
        getNodeText = strQuotationMarks & strValue & strQuotationMarks
    End If
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "getNodeText" & vbRightAllow & Err.Source, Err.Description & vbCrLf & "columnName:" & strColName, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ログインする
'       MODULE_ID       : recaro8_login
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : strUser               ユーザー(I)
'                       : strPass               パスワード(I)
'                       : [iniFile]             接続情報(I)
'       RETURN          : 接続許可文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function recaro8_login(ByVal strUser As String, ByVal strPass As String, Optional iniFile As Collection = Nothing) As String
    Dim objDoc              As Object
    Dim objTbl              As Object
    Dim objRow              As Object
    Dim strTemp             As String
    Dim strResponseXml      As String
    On Error GoTo ErrorHandler
    
    If iniFile Is Nothing Then
        Set iniFile = get_ini_file()
    End If
    
    'データ取得のコマンド生成
    Set objDoc = recaro8_createDocument()
    Set objTbl = recaro8_createTable(objDoc, "LGIN_TABL")
    Call recaro8_whereColumn(objTbl, "LGINT_MAILB", strUser, "varchar")
    'データ取得
    If Not recaro8_table_ctrl(objDoc, strResponseXml, iniFile) Then
        Call MSZZ024_M10("select LGIN_TABL", strResponseXml)
    End If

    '戻りのデータからパスワードを比較
    Set objDoc = XmlCreateDocument(strResponseXml)
    Set objRow = objDoc.selectSingleNode("root/table/row")
    If objRow Is Nothing Then
        Call MSZZ024_M10("recaro8_password_verify", "ユーザー・パスワードの不正でログインできませんでした。")
    End If
    If Not recaro8_password_verify(objRow, strPass) Then
        Call MSZZ024_M10("recaro8_password_verify", "パスワードの不正でログインできませんでした。")
    End If
    
    'ログインすることで他の接続と分ける
    strTemp = random_pseudo_base64(255)
    Set objDoc = recaro8_createDocument()
    Set objTbl = recaro8_createTable(objDoc, "LGIN_TABL", "UPDATE")
    Call recaro8_setColumn(objTbl, "LGINT_LASTD", Format(Now(), "yyyymmdd"), "char")
    Call recaro8_setColumn(objTbl, "LGINT_LASTJ", Format(Now(), "hhnnss"), "char")
    Call recaro8_setColumn(objTbl, "LGINT_TMPTC", strTemp, "varchar")
    Call recaro8_setColumn(objTbl, "LGINT_TMPTL", Format(DateAdd("h", 1, Now()), "yyyymmddhhnnss"), "char")
    Call recaro8_whereColumn(objTbl, "LGINT_MAILB", strUser, "varchar")

    If Not recaro8_table_ctrl(objDoc, strResponseXml, iniFile) Then
        Call MSZZ024_M10("update LGIN_TABL", strResponseXml)
    End If
    '成功したら一時キーを返す
    recaro8_login = strTemp
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_login" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : パスワードチェック
'       MODULE_ID       : recaro8_password_verify
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : objRow                行ノード
'                       : strPass               パスワード(I)
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function recaro8_password_verify(objRow As Object, ByVal strPass As String) As Boolean
    Dim objHash             As Object
    On Error GoTo ErrorHandler
    
    'Set objHash = CreateObject("System.Security.Cryptography.Rfc2898DeriveBytes")
    'recaro8_password_verify = objHash.Verify(strPass, objRow.selectSingleNode("LGINT_PASSN").Text)
    recaro8_password_verify = True
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_password_verify" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : PHP[SendTableCtrl]から移植
'       MODULE_ID       : recaro8_createDocument
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : [strUser]             ユーザー(I)
'                       : [strConnectKey]       接続許可文字列(I)
'       RETURN          : ドキュメントオブジェクト(Object) MSXML2.DOMDocument
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function recaro8_createDocument(Optional ByVal strUser As String = "", Optional ByVal strConnectKey As String = "") As Object
    Dim objDoc              As Object
    Dim objTbl              As Object
    On Error GoTo ErrorHandler
    
    Set objDoc = XmlCreateDocument("<?xml version=""1.0"" encoding=""UTF-8""?><root/>")
    If strConnectKey <> "" Then
        Set objTbl = recaro8_createTable(objDoc, "LGIN_TABL", "UPDATE")
        Call recaro8_whereColumn(objTbl, "LGINT_MAILB", strUser, "varchar")
        Call recaro8_whereColumn(objTbl, "LGINT_TMPTC", strConnectKey, "varchar")
        Call recaro8_whereColumn(objTbl, "LGINT_TMPTL", Format(Now(), "yyyymmddhhnnss"), "char", ">")
        Call recaro8_setColumn(objTbl, "LGINT_TMPTL", Format(DateAdd("h", 1, Now()), "yyyymmddhhnnss"), "char")
    End If
    Set recaro8_createDocument = objDoc
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_createDocument" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : PHP[SendTableCtrl]から移植
'       MODULE_ID       : recaro8_createTable
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : objRoot               ドキュメントオブジェクト(I) MSXML2.DOMDocument
'                       : strTblName            テーブル名(I)
'                       : [strActionName]       アクション文字列(I)
'                       : [lngLimit]            リミット(I)
'       RETURN          : テーブルオブジェクト(Object) MSXML2.IXMLDOMElement
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function recaro8_createTable(objDoc As Object, ByVal strTblName As String, Optional ByVal strActionName As String = "SELECT", Optional ByVal lngLimit = Null) As Object
    Dim objRoot             As Object 'IXMLDOMElement
    Dim objTbl              As Object 'IXMLDOMElement
    Dim objAttr             As Object 'IXMLDOMAttribute
    On Error GoTo ErrorHandler
    
    Set objRoot = objDoc.selectSingleNode("root")
    Set objTbl = objDoc.createElement("table")
    Call objRoot.appendChild(objTbl)
    
    Set objAttr = objDoc.createAttribute("ID")
    objAttr.nodeValue = strTblName
    Call objTbl.Attributes.setNamedItem(objAttr)
    
    Set objAttr = objDoc.createAttribute("Action")
    objAttr.nodeValue = strActionName
    Call objTbl.Attributes.setNamedItem(objAttr)
    
    If Not IsNull(lngLimit) Then
        Set objAttr = objDoc.createAttribute("limit")
        objAttr.nodeValue = lngLimit
        Call objTbl.Attributes.setNamedItem(objAttr)
    End If
    Set recaro8_createTable = objTbl
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_createTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : PHP[SendTableCtrl]から移植
'       MODULE_ID       : recaro8_setColumn
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : objTbl                テーブルオブジェクト(I) MSXML2.IXMLDOMElement
'                       : strColumnName         名前(I)
'                       : strColumnValue        値(I)
'                       : [strType]             型(I)
'       RETURN          : カラムオブジェクト(Object) MSXML2.IXMLDOMElement
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function recaro8_setColumn(objTbl As Object, ByVal strColumnName As String, ByVal strColumnValue As String, Optional ByVal strType As String = "auto") As Object
    Dim objDoc              As Object 'MSXML2.DOMDocument
    Dim objSet              As Object 'IXMLDOMElement
    Dim objCol              As Object 'IXMLDOMElement
    Dim objTxt              As Object 'IXMLDOMText
    Dim objAttr             As Object 'IXMLDOMAttribute
    On Error GoTo ErrorHandler
    
    Set objDoc = objTbl.ownerDocument
    Set objSet = objTbl.selectSingleNode("set")
    If objSet Is Nothing Then
        Set objSet = objDoc.createElement("set")
        Call objTbl.appendChild(objSet)
    End If
        
    Set objCol = objDoc.createElement(strColumnName)
    Call objSet.appendChild(objCol)
    
    Set objTxt = objDoc.createTextNode(strColumnValue)
    Call objCol.appendChild(objTxt)
    
    If strType <> "auto" Then
        Set objAttr = objDoc.createAttribute("type")
        objAttr.nodeValue = strType
        Call objCol.Attributes.setNamedItem(objAttr)
    End If
    Set recaro8_setColumn = objCol
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_setColumn" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : PHP[SendTableCtrl]から移植
'       MODULE_ID       : recaro8_whereColumn
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : objTbl                テーブルオブジェクト(I) MSXML2.IXMLDOMElement
'                       : strColumnName         名前(I)
'                       : strColumnValue        値(I)
'                       : [strType]             型(I)
'                       : [strCondition]        条件(I)
'       RETURN          : カラムオブジェクト(Object) MSXML2.IXMLDOMElement
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function recaro8_whereColumn(objTbl As Object, ByVal strColumnName As String, ByVal strColumnValue As String, Optional ByVal strType As String = "auto", Optional ByVal strCondition As String = "=") As Object
    Dim objDoc              As Object 'MSXML2.DOMDocument
    Dim objSet              As Object 'IXMLDOMElement
    Dim objCol              As Object 'IXMLDOMElement
    Dim objTxt              As Object 'IXMLDOMText
    Dim objAttr             As Object 'IXMLDOMAttribute
    On Error GoTo ErrorHandler
    
    Set objDoc = objTbl.ownerDocument
    Set objSet = objTbl.selectSingleNode("where")
    If objSet Is Nothing Then
        Set objSet = objDoc.createElement("where")
        Call objTbl.appendChild(objSet)
    End If
        
    Set objCol = objDoc.createElement(strColumnName)
    Call objSet.appendChild(objCol)
    
    Set objTxt = objDoc.createTextNode(strColumnValue)
    Call objCol.appendChild(objTxt)
    
    If strType <> "auto" Then
        Set objAttr = objDoc.createAttribute("type")
        objAttr.nodeValue = strType
        Call objCol.Attributes.setNamedItem(objAttr)
    End If
    
    If strCondition <> "=" Then
        Set objAttr = objDoc.createAttribute("condition")
        objAttr.nodeValue = strCondition
        Call objCol.Attributes.setNamedItem(objAttr)
    End If
    Set recaro8_whereColumn = objCol
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_whereColumn" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : データ取得
'       MODULE_ID       : recaro8_select_xml
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : strMethod             メソッド(I)
'                       : objDic                パラメータ(I)
'                       : strResponseXml        受信ＸＭＬデータ(O)
'                       : [iniFile]             接続情報等(I)
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function recaro8_select_xml(ByVal strMethod As String, objDic As Object, ByRef strResponseXml As String, Optional iniFile As Collection = Nothing) As Boolean
    Dim strXml              As String
    On Error GoTo ErrorHandler
    
    If iniFile Is Nothing Then
        Set iniFile = get_ini_file()
    End If
    strXml = textMethodParameter(strMethod, objDic)
    recaro8_select_xml = XmlHttpGet(iniFile("SelectXmlUrl"), iniFile("BA_USER"), iniFile("BA_PASSWORD"), iniFile("SHOP_ID"), strXml, strResponseXml)
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_select_xml" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : データ取得
'       MODULE_ID       : recaro8_select_xml
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : objDoc                送信ドキュメントオブジェクト(I)
'                       : strResponseXml        受信ＸＭＬデータ(O)
'                       : [iniFile]             接続情報等(I)
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function recaro8_table_ctrl(objDoc As Object, ByRef strResponseXml As String, Optional iniFile As Collection = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    If iniFile Is Nothing Then
        Set iniFile = get_ini_file()
    End If
    recaro8_table_ctrl = XmlHttpGet(iniFile("TableCtrlUrl"), iniFile("BA_USER"), iniFile("BA_PASSWORD"), iniFile("SHOP_ID"), objDoc.XML, strResponseXml)
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "recaro8_table_ctrl" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンフィグファイルの取得
'       MODULE_ID       : get_ini_file
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       RETURN          : コンフィグファイルの情報(Collection)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function get_ini_file() As Collection
    Dim strFile             As String
    On Error GoTo ErrorHandler
    
    strFile = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & PROG_ID & "'"))
    If strFile = "" Then
        Call MSZZ024_M10("DLookup", "INTI_FILEの設定不足です。")
    End If
    Set get_ini_file = parse_ini_file(strFile)
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "get_ini_file" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンフィグファイルのパース
'       MODULE_ID       : parse_ini_file
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       RETURN          : コンフィグファイルの情報(Collection)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function parse_ini_file(ByVal strInifile As String) As Collection
    Dim strResult           As String
    Dim strLine             As Variant
    Dim i                   As Integer
    Dim co                  As New Collection
    Dim strKeyVal           As Variant
    On Error GoTo ErrorHandler
    
    strResult = UTF8Read(strInifile)
    For Each strLine In Split(strResult, vbLf)
        If Trim(strLine) <> "" And Left(strLine, 1) <> ";" And (InStr(strLine, "=") > 0) Then
            strKeyVal = Split(Trim(strLine), "=")
            co.Add Replace(strKeyVal(1), """", ""), strKeyVal(0)
        End If
    Next
    Set parse_ini_file = co
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "parse_ini_file" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : パラメータを展開
'       MODULE_ID       : textMethodParameter
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : strMethod             メソッド(I)
'                       : objDic                パラメータ(I)
'       RETURN          : パラメータ(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function textMethodParameter(ByVal strMethod As String, objDic As Object) As String
    Dim arrKeys             As Variant
    Dim arrItem             As Variant
    Dim strParam            As String
    Dim i                   As Long
    On Error GoTo ErrorHandler

    strParam = "method=" & strMethod
    If objDic.Count > 0 Then
        arrKeys = objDic.Keys()
        arrItem = objDic.Items()
        For i = 0 To objDic.Count - 1
            strParam = strParam & "&" & arrKeys(i) & "=" & arrItem(i)
        Next
    End If
    textMethodParameter = strParam
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "textMethodParameter" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : HTTP通信
'       MODULE_ID       : XmlHttpGet
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : strUrl                接続ＵＲＬ(I)
'                       : strBaUser             ユーザー(I)
'                       : strBaPwd              パスワード(I)
'                       : strShopId             ショップＩＤ(I)
'                       : strXml                送信ＸＭＬデータ(I)
'                       : strResponseXml        受信ＸＭＬデータ(O)
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function XmlHttpGet(ByVal strUrl As String, _
    ByVal strBaUser As String, ByVal strBaPwd As String, ByVal strShopId As String, _
    ByVal strXml As Variant, _
    ByRef strResponseXml As String) As Boolean
    Dim objHttp             As Object
    Dim strResultStatus     As String
    On Error GoTo ErrorHandler
    
'    Set objHttp = CreateObject("Microsoft.XMLHTTP")
    Set objHttp = CreateObject("Msxml2.XMLHTTP.6.0")
    With objHttp
        .Open "POST", strUrl, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        If (strBaUser <> "") And (strBaPwd <> "") Then
            .setRequestHeader "Authorization", "Basic " & UnicodeToBase64String(strBaUser & ":" & strBaPwd)
        End If
        If strShopId <> "" Then
'            .setRequestHeader "SHOP_ID", strShopId                             'DELETE 2018/01/13 K.ISHIZAKA
            .setRequestHeader "SHOPID", strShopId                               'INSERT 2018/01/13 K.ISHIZAKA
        End If
        .send strXml
    
        strResponseXml = .responseText
        
        strResultStatus = Nz(.getResponseHeader("Result-status"))
        XmlHttpGet = (strResultStatus = "OK")
    End With
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "XmlHttpGet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : Unicode文字列をBase64文字列に変換
'       MODULE_ID       : ToBase64String
'       CREATE_DATE     : 2017/01/26            K.ISHIZAKA
'       PARAM           : strText               文字列(I)
'       RETURN          : Base64文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function UnicodeToBase64String(ByVal strText As String) As String
    Dim bytUtf8()           As Byte
    On Error GoTo ErrorHandler
    
    bytUtf8 = UTF8_GetBytes(strText)            'Unicode → UTF-8 変換
    strText = ToBase64String(bytUtf8)           'BASE64文字列にして返す
    
    UnicodeToBase64String = strText
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UnicodeToBase64String" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended of program ********************************

