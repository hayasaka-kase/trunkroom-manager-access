Attribute VB_Name = "MSZZ026"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : ＸＭＬＨＴＴＰ関数
'        PROGRAM_ID      : MSZZ026
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/02/17
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          : 2007/08/14
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.1
'                        : 作成時にXMLファイル名を指定できるようにするための
'
'==============================================================================*
Option Explicit

'==============================================================================*
'   API宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Function WideCharToMultiByte Lib "kernel32" _
    (ByVal CodePage As Long, _
     ByVal dwFlags As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long, _
     ByRef lpMultiByteStr As Any, _
     ByVal cchMultiByte As Long, _
     ByVal lpDefaultChar As String, _
     ByVal lpUsedDefaultChar As Long) As Long
     
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const CP_UTF8 = 65001

'==============================================================================*
'
'       MODULE_NAME     : UFT-8 変換
'       MODULE_ID       : Utf8Encode
'       CREATE_DATE     : 2007/02/17            K.ISHIZAKA
'       PARAM           : strUnicode            Unicode 文字列(I)
'       RETURN          : UTF-8 文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Utf8Encode(ByRef strUnicode As String) As String
    Dim lngUnicodeLength    As Long             'Unicode 文字列の数を格納する変数
    Dim lngBufferSize       As Long             'UTF-8 文字列を格納するバッファのサイズを格納する変数
    Dim lngResult           As Long             'WideCharToMultiByte の戻り値を格納する変数
    Dim bytUtf8()           As Byte
    Dim strUtf8             As String
    Dim i                   As Long
    On Error GoTo ErrorHandler

    lngUnicodeLength = Len(strUnicode)          '変換元 Unicode 文字列の数を得る
    If lngUnicodeLength = 0 Then Exit Function  '0 文字の場合関数を抜ける
    lngBufferSize = lngUnicodeLength * 3 + 1    'バッファサイズを Unicode 文字列数の 3 倍 + 1 バイトと決める
    ReDim bytUtf8(lngBufferSize)                'UTF-8 文字列を格納するバッファを確保
    lngResult = WideCharToMultiByte _
        (CP_UTF8, _
         0, _
         StrPtr(strUnicode), _
         lngUnicodeLength, _
         bytUtf8(0), _
         lngBufferSize, _
         vbNullString, _
         0)                                     'WideCharToMultiByte で Unicode → UTF-8 変換
    For i = 0 To lngResult - 1
        strUtf8 = strUtf8 & "%" & Hex(bytUtf8(i))
    Next
    Utf8Encode = strUtf8                        'UTF-8 文字列を返す
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "Utf8Encode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : HTTP要求
'       MODULE_ID       : XmlHttpGet
'       CREATE_DATE     : 2007/02/17            K.ISHIZAKA
'       PARAM           : strUrl                ＵＲＬ(I)
'       RETURN          : 回答文字列(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function XmlHttpGet(ByVal strUrl As String) As String
    Dim xmlHttp             As Object
    On Error GoTo ErrorHandler
    
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    Call xmlHttp.Open("GET", strUrl, False)
    Call xmlHttp.send(Null)
    XmlHttpGet = xmlHttp.responseText
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "XmlHttpGet" & vbRightAllow & Err.Source, Err.Description & strUrl, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : XMLドキュメント作成
'       MODULE_ID       : XmlCreateDocument
'       CREATE_DATE     : 2007/02/17            K.ISHIZAKA
'       PARAM           : [strXML]              XML文字列：省略可(I)
'       RETURN          : XMLドキュメント(Object)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function XmlCreateDocument(Optional strXml As String = "") As Object
    Dim objDoc              As Object   'MSXML2.DOMDocument
    On Error GoTo ErrorHandler
    
    Set objDoc = CreateObject("MSXML2.DOMDocument")
    If strXml <> "" Then
'        Call objDoc.loadXML(strXML)                                            'DELETE 2007/08/14 K.ISHIZAKA
        If Left(strXml, 1) = "<" Then                                           'INSERT START 2007/08/14 K.ISHIZAKA
            If objDoc.loadXML(strXml) = False Then
                Call MSZZ024_M10("MSXML2.DOMDocument", "Load Error!![" & strXml & "]")
            End If
        Else
            objDoc.async = False
            If objDoc.Load(strXml) = False Then
                Call MSZZ024_M10("MSXML2.DOMDocument", "Load Error!![" & strXml & "]")
            End If
        End If                                                                  'INSERT END   2007/08/14 K.ISHIZAKA
    End If
    Set XmlCreateDocument = objDoc
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "XmlCreateDocument" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ＸＭＬからコンボボックスリスト一覧を作成
'       MODULE_ID       : XmlNodesMakeRowSource
'       CREATE_DATE     : 2007/02/17            K.ISHIZAKA
'       PARAM           : objDoc                XMLドキュメント(I)
'                       : strPath               パス(I)
'                       : [strOutOfData]        対象外データ：省略可(I)
'       RETURN          : 一覧(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function XmlNodesMakeRowSource(objDoc As Object, ByVal strPath As String, Optional strOutOfData As String = "") As String
    Dim strNames            As String
    Dim objNd               As Object   'MSXML2.IXMLDOMNode
    On Error GoTo ErrorHandler
    
    For Each objNd In objDoc.selectNodes(strPath)
        If strOutOfData <> objNd.Text Then
            strNames = strNames & ";" & objNd.Text
        End If
    Next
    XmlNodesMakeRowSource = Mid(strNames, 2)
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "XmlNodesMakeRowSource" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ＸＭＬからコンボボックスリスト一覧を作成
'       MODULE_ID       : XmlSingleNodeMakeRowSource
'       CREATE_DATE     : 2007/02/17            K.ISHIZAKA
'       PARAM           : objDoc                XMLドキュメント(I)
'                       : strPath               パス(I)
'                       : [strChar]             区切り文字：省略可(I)
'       RETURN          : 一覧(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function XmlSingleNodeMakeRowSource(objDoc As Object, ByVal strPath As String, Optional strChar As String = " ") As String
    On Error GoTo ErrorHandler
    
    XmlSingleNodeMakeRowSource = Replace(objDoc.selectSingleNode(strPath).Text, strChar, ";")
Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "XmlSingleNodeMakeRowSource" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended or program ********************************
