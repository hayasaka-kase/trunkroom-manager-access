Attribute VB_Name = "MSZZ062"
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : EUC形式ファイル関数
'       PROGRAM_ID      : MSZZ062
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2011/01/24
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          :
'       UPDATER         :
'       Ver             :
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Const EUC_XML_HEADER    As String = "<?xml version=""1.0"" encoding=""EUC-JP""?>"

'==============================================================================*
'
'       MODULE_NAME     : EUC形式でファイル保存
'       MODULE_ID       : EUCSave
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strFullpathFilename   ファイル名(I)
'                       : strText               内容(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub EUCSave(ByVal strFullpathFilename As String, ByVal strText As String)
    Dim objEUC              As Object
    On Error GoTo ErrorHandler

    Set objEUC = EUCOpen()
    On Error GoTo ErrorHandler1
    EUCWrite objEUC, strText
    EUCClose objEUC, strFullpathFilename
Exit Sub

ErrorHandler1:
    EUCClose objEUC
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "EUCSave" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : EUC形式でファイルオープン
'       MODULE_ID       : EUCOpen
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       RETURN          : ストリームオブジェクト(Object)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function EUCOpen() As Object
    Dim objEUC              As Object
    On Error GoTo ErrorHandler
    
    Set objEUC = CreateObject("ADODB.Stream")
    With objEUC
        .Type = 2           '1:バイナリ、 2:テキスト
        .Charset = "euc-jp"
        .LineSeparator = 10     '-1:CRLF、 10:LF、 13:CR
        .Open
    End With
    Set EUCOpen = objEUC
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "EUCOpen" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : EUC形式でファイル書き出し
'       MODULE_ID       : EUCWrite
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : objEUC                ストリームオブジェクト(I)
'                       : strText               内容(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub EUCWrite(objEUC As Object, ByVal strText As String)
    On Error GoTo ErrorHandler
    
    objEUC.WriteText strText, 1         '0:改行なし、 1:あり
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "EUCWrite" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : EUC形式でファイルクローズ
'       MODULE_ID       : EUCClose
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : objEUC                ストリームオブジェクト(I)
'                       : [strFullpathFilename] 保存ファイル名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub EUCClose(objEUC As Object, Optional ByVal strFullpathFilename As String = "")
    If strFullpathFilename <> "" Then
        On Error GoTo ErrorHandler
        objEUC.SaveToFile strFullpathFilename, 2
    End If
    objEUC.Close
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "EUCClose" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : EUC形式ファイルの読み込み
'       MODULE_ID       : EUCRead
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       PARAM           : strFileName           ファイル名(I)
'       RETURN          : ファイル内容(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function EUCRead(ByVal strFilename As String) As String
    Dim objEUC             As Object
    Dim strBuff             As String
    On Error GoTo ErrorHandler
    
    Set objEUC = EUCOpen()
    On Error GoTo ErrorHandler1
    With objEUC
        .LoadFromFile strFilename
        strBuff = .ReadText(-1)    '-1:All、 -2:Line
        .Close
    End With
    On Error GoTo ErrorHandler
    EUCRead = strBuff
Exit Function

ErrorHandler1:
    objEUC.Close
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "EUCRead" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended of program ********************************
