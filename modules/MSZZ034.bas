Attribute VB_Name = "MSZZ034"
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : UTF8形式ファイル関数
'       PROGRAM_ID      : MSZZ034
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2007/08/14
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/02/17
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : 関数 UTF8Read を追加
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Const UTF8_XML_HEADER    As String = "<?xml version=""1.0"" encoding=""UTF-8""?>"

'==============================================================================*
'
'       MODULE_NAME     : UTF8形式でファイル保存
'       MODULE_ID       : UTF8Save
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strFullpathFilename   ファイル名(I)
'                       : strText               内容(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub UTF8Save(ByVal strFullpathFilename As String, ByVal strText As String)
    Dim intFilenumber       As Integer
    On Error GoTo ErrorHandler

    intFilenumber = UTF8Open(strFullpathFilename)
    On Error GoTo ErrorHandler1
    UTF8Write intFilenumber, strText
    UTF8Close intFilenumber
Exit Sub

ErrorHandler1:
    UTF8Close intFilenumber
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UTF8Save" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : UTF8形式でファイルオープン
'       MODULE_ID       : UTF8Open
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strFullpathFilename   ファイル名(I)
'       RETURN          : ファイルID(Integer)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UTF8Open(ByVal strFullpathFilename As String) As Integer
    Dim intFilenumber       As Integer
    On Error GoTo ErrorHandler
    
    If Dir(strFullpathFilename) <> "" Then
        Kill strFullpathFilename
    End If
    intFilenumber = FreeFile()
    Open strFullpathFilename For Binary Access Write As intFilenumber
    UTF8Open = intFilenumber
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UTF8Open" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : UTF8形式でファイル書き出し
'       MODULE_ID       : UTF8Write
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : intFilenumber         ファイルID(I)
'                       : strText               内容(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub UTF8Write(ByVal intFilenumber As Integer, ByVal strText As String)
    Dim bytText()           As Byte
    Dim i                   As Long
    On Error GoTo ErrorHandler
    
    bytText = UTF8_GetBytes(strText & vbCrLf)
    i = UBound(bytText)
    If bytText(i) = Asc(vbNullChar) Then
        ReDim Preserve bytText(i - 1)
    End If
    Put intFilenumber, , bytText
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UTF8Write" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : UTF8形式でファイルクローズ
'       MODULE_ID       : UTF8Close
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : intFilenumber         ファイルID(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub UTF8Close(ByVal intFilenumber As Integer)
    Close intFilenumber
End Sub

'==============================================================================*
'
'       MODULE_NAME     : UTF8形式ファイルの読み込み
'       MODULE_ID       : UTF8Read
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       PARAM           : strFileName           ファイル名(I)
'       RETURN          : ファイル内容(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UTF8Read(ByVal strFilename As String) As String
    Dim objUTF8             As Object
    Dim strBuff             As String
    On Error GoTo ErrorHandler
    
    Set objUTF8 = CreateObject("ADODB.Stream")
    With objUTF8
        .Type = 2                   '1:バイナリ、 2:テキスト
        .Charset = "UTF-8"
        .LineSeparator = -1         '-1:CRLF、 10:LF、 13:CR
        .Open
        On Error GoTo ErrorHandler1
        .LoadFromFile strFilename
        strBuff = .ReadText(-1)    '-1:All、 -2:Line
        .Close
    End With
    On Error GoTo ErrorHandler
    UTF8Read = strBuff
Exit Function

ErrorHandler1:
    objUTF8.Close
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "UTF8Read" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended of program ********************************
