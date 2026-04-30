VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySqlRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : WebServer MySQL レコードセットクラス
'       PROGRAM_ID      : MySqlRecordSet
'       PROGRAM_KBN     : Class
'
'       CREATE          : 2011/05/21
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2012/06/16
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : プロパティBOFの追加
'                       : フィールドのクラス化（ForEach や With や .Name 、.Value に対応）
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'ADODB.ErrorValueEnum の完全エラーコードと同じ
Private Enum ErrorValueEnum
    adErrNoCurrentRecord = -2146825267  'BOF と EOF のいずれかが True になっているか、またはカレントレコードが削除されています。要求された操作には、カレントレコードが必要です。
    adErrObjectClosed = -2146824584     'オブジェクトが閉じている場合は、操作は許可されません。
    adErrObjectOpen = -2146824583       'オブジェクトが開いている場合は、操作は許可されません。
End Enum

'ADODB.ErrorValueEnum のエラー内容と同じ
Private Const adErrMsgNoCurrentRecord   As String = "BOF と EOF のいずれかが True になっているか、またはカレントレコードが削除されています。要求された操作には、カレントレコードが必要です。"
Private Const adErrMsgObjectClosed      As String = "オブジェクトが閉じている場合は、操作は許可されません。"
Private Const adErrMsgObjectOpen        As String = "オブジェクトが開いている場合は、操作は許可されません。"

'ADODB.Stream の ReadText パラメータ    '-1:All、 -2:Line
Private Const adReadLine                As Long = -2
Private Const adReadAll                 As Long = -1

'拡張子
Private Const C_EXTENSION_CMD           As String = ".txt"
Private Const C_EXTENSION_LOG           As String = ".log"
Private Const C_EXTENSION_ERR           As String = ".err"

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private strPath                         As String           'リダイレクト用一時ファイル名
Private objStream                       As Object           'ファイルアクセスオブジェクト
Private varFields                       As Variant          '内部的にフィールド名を保持
Private objFields                       As Collection       '内部的にフィールド値を保持

Private lngBofPosition                  As Long

Private strMyCharset                    As String           '内部的に Charset を保持
Private lngMyLineSeparator              As Long             '内部的に LineSeparator を保持

'==============================================================================*
'
'       MODULE_NAME     : クラス初期処理
'       MODULE_ID       : Class_Initialize
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set objFields = New Collection
    Set objStream = Nothing
    strMyCharset = "UTF-8"                  'UTF-8、 euc-jp
    lngMyLineSeparator = 10                 '-1:CRLF、 10:LF、 13:CR
    strPath = Environ("temp") & "\MySqlCmd" & Format(Now, "yyyymmddhhnnss")
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "MySqlRecordSet.Class_Initialize" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : クラス終了処理
'       MODULE_ID       : Class_Terminate
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Class_Terminate()
    Me.Close_
    Set objFields = Nothing
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 文字セット（プロパティ）
'       MODULE_ID       : Charset
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : strCharset            文字セット(I) 例、"UTF-8"、"euc-jp" など
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Let Charset(ByVal strCharset As String)
    If Not objStream Is Nothing Then
        Call Err.Raise(adErrObjectOpen, "MySqlRecordSet.Charset", adErrMsgObjectOpen)
    End If
    strMyCharset = strCharset
End Property

'==============================================================================*
'
'       MODULE_NAME     : 文字セット（プロパティ）
'       MODULE_ID       : Charset
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       RETURN          : 文字セット(String)    例、"UTF-8"、"euc-jp" など
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get Charset() As String
    Charset = strMyCharset
End Property

'==============================================================================*
'
'       MODULE_NAME     : 行末文字コード（プロパティ）
'       MODULE_ID       : LineSeparator
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : lngLineSeparator      文字コード(I) 例、-1:CRLF、 10:LF、 13:CR など
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Let LineSeparator(ByVal lngLineSeparator As Long)
    If Not objStream Is Nothing Then
        Call Err.Raise(adErrObjectOpen, "MySqlRecordSet.LineSeparator", adErrMsgObjectOpen)
    End If
    lngMyLineSeparator = lngLineSeparator
End Property

'==============================================================================*
'
'       MODULE_NAME     : 行末文字コード（プロパティ）
'       MODULE_ID       : LineSeparator
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       RETURN          : 文字コード(Long)      例、-1:CRLF、 10:LF、 13:CR など
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get LineSeparator() As Long
    LineSeparator = lngMyLineSeparator
End Property

'==============================================================================*
'
'       MODULE_NAME     : ファイル初端（プロパティ）
'       MODULE_ID       : BOF
'       CREATE_DATE     : 2012/06/16            K.ISHIZAKA
'       RETURN          : 初端(True)／以外(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get BOF() As Boolean
    If objStream Is Nothing Then
        Call Err.Raise(adErrObjectClosed, "MySqlRecordSet.EOF", adErrMsgObjectClosed)
    End If
    BOF = (objStream.Position <= lngBofPosition) And (objFields.Count = 0)
End Property

'==============================================================================*
'
'       MODULE_NAME     : ファイル終端（プロパティ）
'       MODULE_ID       : EOF
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       RETURN          : 終端(True)／以外(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get EOF() As Boolean
    If objStream Is Nothing Then
        Call Err.Raise(adErrObjectClosed, "MySqlRecordSet.EOF", adErrMsgObjectClosed)
    End If
    EOF = objStream.EOS And (objFields.Count = 0)
End Property

'==============================================================================*
'
'       MODULE_NAME     : フィールド（プロパティ）
'       MODULE_ID       : Fields
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : varIndexOrName        指定なし
'                                               フィールド名(String)
'                                               フィールドインデックス(Long)
'       RETURN          : 引数指定なしのとき、コレクションオブジェクト(Collection)
'                         以外のとき、フィールドの内容(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Property Get Fields(Optional ByVal varIndexOrName As Variant) As Variant
    If objStream Is Nothing Then
        Call Err.Raise(adErrObjectClosed, "MySqlRecordSet.Fields", adErrMsgObjectClosed)
    End If
    If IsMissing(varIndexOrName) Then
        Set Fields = objFields
    ElseIf TypeName(varIndexOrName) = "String" Then
'        Fields = objFields(varIndexOrName)                                     'DELETE 2012/06/16 K.ISHIZAKA
        Set Fields = objFields(varIndexOrName)                                  'INSERT 2012/06/16 K.ISHIZAKA
    Else
        'VB の Collectoin は 1 オリジンなので +1 する
'        Fields = objFields(varIndexOrName + 1)                                 'DELETE 2012/06/16 K.ISHIZAKA
        Set Fields = objFields(varIndexOrName + 1)                              'INSERT 2012/06/16 K.ISHIZAKA
    End If
End Property

'==============================================================================*
'
'       MODULE_NAME     : フィールド名取得
'       MODULE_ID       : GetFieldName
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : lngIndex              フィールドインデックス(Long)
'       RETURN          : フィールド名(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetFieldName(ByVal lngIndex As Long) As String
    If objStream Is Nothing Then
        Call Err.Raise(adErrObjectClosed, "MySqlRecordSet.GetFieldName", adErrMsgObjectClosed)
    End If
    GetFieldName = varFields(lngIndex)
End Function

'==============================================================================*
'
'       MODULE_NAME     : フィールド名全て取得
'       MODULE_ID       : GetFieldNameAll
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       RETURN          : フィールド名配列(String())
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetFieldNameAll() As String()
    If objStream Is Nothing Then
        Call Err.Raise(adErrObjectClosed, "MySqlRecordSet.GetFieldNameAll", adErrMsgObjectClosed)
    End If
    GetFieldNameAll = varFields
End Function

'==============================================================================*
'
'       MODULE_NAME     : 先頭レコードへ移動
'       MODULE_ID       : MoveFirst
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MoveFirst()
    Dim strMsg              As String
    On Error GoTo ErrorHandler
    
    If objStream Is Nothing Then
        Call Err.Raise(adErrObjectClosed, "MySqlRecordSet.MoveFirst", adErrMsgObjectClosed)
    End If
    With objStream
        .Position = 0
        If Not .EOS Then
            strMsg = .ReadText(adReadLine)  '先頭１行目はフィールド名なので、読み捨て
        End If
    End With
    Me.MoveNext
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "MySqlRecordSet.MoveFirst" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 次レコードへ移動
'       MODULE_ID       : MoveNext
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MoveNext()
    Dim i                   As Long
    Dim varArray            As Variant
    Dim fld                 As MySqlField                                       'INSERT 2012/06/16 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    If objStream Is Nothing Then
        Call Err.Raise(adErrObjectClosed, "MySqlRecordSet.MoveNext", adErrMsgObjectClosed)
    End If
    If Me.EOF Then
        Call Err.Raise(adErrNoCurrentRecord, "MySqlRecordSet.MoveNext", adErrMsgNoCurrentRecord)
    End If
    For i = objFields.Count To 1 Step -1
        objFields.Remove i
    Next
    If Not objStream.EOS Then
        varArray = Split(objStream.ReadText(adReadLine), vbTab)
        For i = 0 To UBound(varArray)
            Set fld = New MySqlField                                            'INSERT 2012/06/16 K.ISHIZAKA
            fld.SetValue varArray(i), varFields(i)                              'INSERT 2012/06/16 K.ISHIZAKA
'            objFields.Add varArray(i), varFields(i)                            'DELETE 2012/06/16 K.ISHIZAKA
            objFields.Add fld, varFields(i)                                     'INSERT 2012/06/16 K.ISHIZAKA
        Next
    End If
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "MySqlRecordSet.MoveNext" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : オープン
'       MODULE_ID       : Open_
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : strPuttyCommand       PuttyLink のコマンド
'                       : strMySqlCommad        MySQL のコマンド
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Open_(ByVal strPuttyCommand As String, ByVal strMySqlCommad As String) As Long
    Dim lngResult           As Long
    Dim strMsg              As String
    On Error GoTo ErrorHandler

    If Not objStream Is Nothing Then
        Call Err.Raise(adErrObjectOpen, "MySqlRecordSet.Open_", adErrMsgObjectOpen)
    End If
    lngBofPosition = 0                                                          'INSERT 2012/06/16 K.ISHIZAKA
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Type = 2                               '1:バイナリ、 2:テキスト
        .Charset = strMyCharset
        .LineSeparator = lngMyLineSeparator     '-1:CRLF、 10:LF、 13:CR
        If strMyCharset = "UTF-8" Then
            'UTF-8のとき ADODB.Stream で書き込むと BOM が先頭についてしまうので、独自の UTF-8 で書き出す
            UTF8Save strPath & C_EXTENSION_CMD, strMySqlCommad
        Else
            .Open
            On Error GoTo ErrorHandler1
            .WriteText strMySqlCommad, 1        '0:改行なし、 1:あり
            .SaveToFile strPath & C_EXTENSION_CMD, 2
            .Close
            On Error GoTo ErrorHandler
        End If
    End With
    lngResult = ShellWScriptShell(strPuttyCommand & " -m " & strPath & C_EXTENSION_CMD & " 1> " & strPath & C_EXTENSION_LOG & " 2> " & strPath & C_EXTENSION_ERR)
    If lngResult <> 0 Then
        'エラーが発生していたら、StdErr の内容をリダイレクトしたファイルから
        '文字コード変換して読み込み、エラー発生させる
        With objStream
            .Open
            On Error GoTo ErrorHandler1
            .LoadFromFile strPath & C_EXTENSION_ERR
            strMsg = .ReadText(adReadAll)
            .Close
            On Error GoTo ErrorHandler
        End With
        Set objStream = Nothing
        Call Err.Raise(lngResult, "MySqlRecordSet.Open_", strMsg)
    End If
    '正常なら、StdOut の内容をリダイレクトしたファイルから
    '文字コード変換して先頭１行を読み込み、フィールド名として保持
    '２行目以降をデータとして扱う
    With objStream
        .Open
        .LoadFromFile strPath & C_EXTENSION_LOG
        If Not .EOS Then
            strMsg = .ReadText(adReadLine)
            varFields = Split(strMsg, vbTab)
            lngBofPosition = .Position                                          'INSERT 2012/06/16 K.ISHIZAKA
            If Not .EOS Then
                Me.MoveNext
            End If
        End If
    End With
    Open_ = lngResult
Exit Function

ErrorHandler1:
    objStream.Close
    Set objStream = Nothing
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "MySqlRecordSet.Open_" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : クローズ
'       MODULE_ID       : Close_
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Close_()
    Dim i                   As Long
    'この関数内で On Error を使用してはいけない！
    'また、この関数内でエラーが発生すると、本来のエラー内容が分からなくなるので
    '絶対に間違えたコードを書かないこと！
    
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
    If Not objFields Is Nothing Then
        For i = objFields.Count To 1 Step -1
            objFields.Remove i
        Next
    End If
    varFields = Empty
    If Dir(strPath & C_EXTENSION_CMD, vbNormal) <> "" Then
        Kill strPath & C_EXTENSION_CMD
    End If
    If Dir(strPath & C_EXTENSION_LOG, vbNormal) <> "" Then
        Kill strPath & C_EXTENSION_LOG
    End If
    If Dir(strPath & C_EXTENSION_ERR, vbNormal) <> "" Then
        Kill strPath & C_EXTENSION_ERR
    End If
End Sub

'****************************  ended of program ********************************
