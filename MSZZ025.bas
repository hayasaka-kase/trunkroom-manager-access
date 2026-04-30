Attribute VB_Name = "MSZZ025"
'****************************  start of program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : ＡＤＯ関数
'        PROGRAM_ID      : MSZZ025
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/02/14
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          : 2007/03/19
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.1
'                        : ADODB_RecordsetMode に AppendOnly を追加
'
'        UPDATE          : 2007/08/14
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.2
'                        : ADODB_ExecGetLong を追加
'
'        UPDATE          : 2008/12/08
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.3
'                        : ADODB_ExecGetVariant を追加
'
'        UPDATE          : 2010/06/12
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.4
'                        : ADODB_DropTable を追加
'                        : GetSourceTableName を追加
'
'        UPDATE          : 2011/10/31
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.5
'                        : ADODB_DropTable をエラーハンドル内で呼出しても良いようにする
'
'        UPDATE          : 2011/12/28
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.6
'                        : ADODB_CreateWorkTable_CurrentDb を追加
'
'        UPDATE          : 2012/01/19
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.7
'                        : ADODB_Execute をエラーハンドル内で呼出しても良いようにする
'
'        UPDATE          : 2012/03/02
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.8
'                        : ADODB_Reset_CurrentDb を追加
'                          ※ADODB_CreateWorkTable_CurrentDbで作ったワークテーブルの内容を入れ替えるため
'
'        UPDATE          : 2014/06/09
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.9
'                        : ADODB_DropTable で一時テーブルを削除できるようにする
'
'        UPDATE          : 2015/08/12
'        UPDATER         : K.ISHIZAKA
'        Ver             : 1.0
'                        : CreatWorkTable で adDBTimeStamp に対応する
'
'        UPDATE          : 2016/04/04
'        UPDATER         : K.ISHIZAKA
'        Ver             : 1.1
'                        : タイムアウト設定が指定されないとき、デフォルトに戻すのを止める
'
'        UPDATE          : 2019/09/05
'        UPDATER         : M.HONDA
'        Ver             : 1.1
'                        : adIntegerを追加
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const C_DEFAULT_TIMEOUT     As Long = 180

'AdoEnums.DataType
Public Enum AdoEnums_DataType
    adChar = 129
    adNumeric = 131
    adDBTimeStamp = 135
    adVarChar = 200
    adInteger = 3
End Enum

'AdoEnums.CursorType
Private Enum AdoEnums_CursorType
    adOpenForwardOnly = 0
    adOpenKeyset = 1
End Enum

'AdoEnums.LockType
Private Enum AdoEnums_LockType
    adLockReadOnly = 1
    adLockPessimistic = 2
'    adLockOptimistic = 3
End Enum

'AdoEnums.CommandType
Private Enum AdoEnums_CommandType
    adCmdText = 1
    adCmdTable = 2
End Enum

'AdoEnums.FieldAttribute
Public Enum AdoEnums_FieldAttribute
    adFldUnknownUpdatable = 8
End Enum

Public Enum ADODB_RecordsetMode
'    adoReadOnly = True                                                         'DELETE 2007/03/19 K.ISHIZAKA
'    adoReadWrite = False                                                       'DELETE 2007/03/19 K.ISHIZAKA
    adoReadOnly = 0                                                             'INSERT 2007/03/19 K.ISHIZAKA
    adoReadWrite = 1                                                            'INSERT 2007/03/19 K.ISHIZAKA
    adoAppendOnly = 2                                                           'INSERT 2007/03/19 K.ISHIZAKA
End Enum

'==============================================================================*
'
'       MODULE_NAME     : ADO接続を開く
'       MODULE_ID       : ADODB_Connection
'       CREATE_DATE     : 2007/02/14            K.ISHIZAKA
'       PARAM           : [strBumoc]            接続先ＤＢの部門コード：省略時KASE_DB(I)
'       RETURN          : コネクションオブジェクト(Object)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function ADODB_Connection(Optional strBUMOC As String = "") As Object
    Dim dbObj               As Object
    On Error GoTo ErrorHandler
    
    Set dbObj = CreateObject("ADODB.Connection")
    dbObj.CommandTimeout = C_DEFAULT_TIMEOUT
    Call dbObj.Open(MSZZ007_M10(strBUMOC))
    
    Set ADODB_Connection = dbObj
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "ADODB_Connection" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ADO接続でレコードセットを開く
'       MODULE_ID       : ADODB_Recordset
'       CREATE_DATE     : 2007/02/14            K.ISHIZAKA
'       PARAM           : strSQL                ＳＱＬ文字列(I)
'                       : objConnection         コネクションオブジェクト(I)
'                       : [bRecordsetMode]      モード：省略時読込専用(I)
'       RETURN          : レコードセットオブジェクト(Object)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function ADODB_Recordset(ByVal strSQL As String, objConnection As Object, Optional bRecordsetMode As ADODB_RecordsetMode = adoReadOnly) As Object 'DELETE 2016/04/04 K.ISHIZAKA
Public Function ADODB_Recordset(ByVal strSQL As String, objConnection As Object, Optional bRecordsetMode As ADODB_RecordsetMode = adoReadOnly, Optional lTimeout As Variant) As Object 'INSERT 2016/04/04 K.ISHIZAKA
    Dim rsObj               As Object
    On Error GoTo ErrorHandler
    
'    objConnection.CommandTimeout = C_DEFAULT_TIMEOUT                           'DELETE 2016/04/04 K.ISHIZAKA
    If Not IsMissing(lTimeout) Then                                             'INSERT 2016/04/04 K.ISHIZAKA
        objConnection.CommandTimeout = lTimeout                                 'INSERT 2016/04/04 K.ISHIZAKA
    End If                                                                      'INSERT 2016/04/04 K.ISHIZAKA
    Set rsObj = CreateObject("ADODB.Recordset")
    If bRecordsetMode = adoReadOnly Then
        rsObj.Open strSQL, objConnection, adOpenForwardOnly, adLockReadOnly, IIf(InStr(1, Trim(strSQL), " ") > 0, adCmdText, adCmdTable)
    Else
        If (bRecordsetMode = adoAppendOnly) And (InStr(1, Trim(strSQL), " ") <= 0) Then 'INSERT 2007/03/19 K.ISHIZAKA
            strSQL = "SELECT TOP 0 * FROM " & strSQL                            'INSERT 2007/03/19 K.ISHIZAKA
        End If                                                                  'INSERT 2007/03/19 K.ISHIZAKA
        rsObj.Open strSQL, objConnection, adOpenKeyset, adLockPessimistic, IIf(InStr(1, Trim(strSQL), " ") > 0, adCmdText, adCmdTable)
    End If
    
    Set ADODB_Recordset = rsObj
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "ADODB_Recordset" & vbRightAllow & Err.Source, Err.Description & vbCrLf & "SQL:" & vbCrLf & strSQL, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ADO接続でＳＱＬを実行する
'       MODULE_ID       : ADODB_Execute
'       CREATE_DATE     : 2007/02/14            K.ISHIZAKA
'       PARAM           : strSQL                ＳＱＬ文字列(I)
'                       : [objConnection]       コネクションオブジェクト：省略可(I)
'                                                   省略時 KASE_DB に対して実行する
'                       : [lTimeout]            タイムアウト値：省略可(I)
'       RETURN          : 対象件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function ADODB_Execute(ByVal strSQL As String, Optional objConnection As Object = Nothing, Optional lTimeout As Long = C_DEFAULT_TIMEOUT) As Long 'DELETE 2016/04/04 K.ISHIZAKA
Public Function ADODB_Execute(ByVal strSQL As String, Optional objConnection As Object = Nothing, Optional lTimeout As Variant) As Long 'INSERT 2016/04/04 K.ISHIZAKA
    Dim lngCnt              As Long
    Dim dbObj               As Object
    '↓DELETE 2012/01/19 SHIBAZAKI
'    On Error GoTo ErrorHandler
'    If objConnection Is Nothing Then
'        Set dbObj = ADODB_Connection()
'        On Error GoTo ErrorHandler2
'        dbObj.CommandTimeout = lTimeout
'        dbObj.Execute strSQL, lngCnt, adCmdText
'        dbObj.Close
'        On Error GoTo ErrorHandler
'    Else
'        objConnection.CommandTimeout = lTimeout
'        objConnection.Execute strSQL, lngCnt, adCmdText
'    End If
    '↑DELETE 2012/01/19 SHIBAZAKI
    
    '↓INSERT 2012/01/19 SHIBAZAKI
    Dim lngErrNumber        As Long
    Dim strErrSource        As String
    Dim strErrDescription   As String
    Dim strErrHelpFile      As String
    Dim lngErrHelpContext   As Long
    
    If Err.Number = 0 Then
        On Error GoTo ErrorHandler
        If objConnection Is Nothing Then
            Set dbObj = ADODB_Connection()
            On Error GoTo ErrorHandler2
            If Not IsMissing(lTimeout) Then                                     'INSERT 2016/04/04 K.ISHIZAKA
                dbObj.CommandTimeout = lTimeout
            Else                                                                'INSERT 2016/04/04 K.ISHIZAKA
                dbObj.CommandTimeout = C_DEFAULT_TIMEOUT                        'INSERT 2016/04/04 K.ISHIZAKA
            End If                                                              'INSERT 2016/04/04 K.ISHIZAKA
            dbObj.Execute strSQL, lngCnt, adCmdText
            dbObj.Close
            On Error GoTo ErrorHandler
        Else
            If Not IsMissing(lTimeout) Then                                     'INSERT 2016/04/04 K.ISHIZAKA
                objConnection.CommandTimeout = lTimeout
            End If                                                              'INSERT 2016/04/04 K.ISHIZAKA
            objConnection.Execute strSQL, lngCnt, adCmdText
        End If
    Else
        lngErrNumber = Err.Number
        strErrSource = Err.Source
        strErrDescription = Err.Description
        strErrHelpFile = Err.HelpFile
        lngErrHelpContext = Err.HelpContext
        On Error Resume Next
        If objConnection Is Nothing Then
            Set dbObj = ADODB_Connection()
            If Err.Number = 0 Then
                If Not IsMissing(lTimeout) Then                                 'INSERT 2016/04/04 K.ISHIZAKA
                    dbObj.CommandTimeout = lTimeout
                Else                                                            'INSERT 2016/04/04 K.ISHIZAKA
                    dbObj.CommandTimeout = C_DEFAULT_TIMEOUT                    'INSERT 2016/04/04 K.ISHIZAKA
                End If                                                          'INSERT 2016/04/04 K.ISHIZAKA
                dbObj.Execute strSQL, lngCnt, adCmdText
                dbObj.Close
            End If
        Else
            objConnection.CommandTimeout = lTimeout
            objConnection.Execute strSQL, lngCnt, adCmdText
        End If
        Call Err.Raise(lngErrNumber, strErrSource, strErrDescription, strErrHelpFile, lngErrHelpContext)
    End If
    '↑INSERT 2012/01/19 SHIBAZAKI
    
    ADODB_Execute = lngCnt
Exit Function

ErrorHandler2:
    dbObj.Close
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "ADODB_Execute" & vbRightAllow & Err.Source, Err.Description & vbCrLf & "SQL:" & vbCrLf & strSQL, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ADO接続でＳＱＬを実行し先頭フィールドをLong型で取得
'       MODULE_ID       : ADODB_ExecGetLong
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strSQL                ＳＱＬ文字列(I)
'                       : objConnection         コネクションオブジェクト(I)
'                       : [lngEofValue]         EOFだったときの戻り値(I)省略時:-1
'       RETURN          : 先頭フィールドの内容(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function ADODB_ExecGetLong(ByVal strSQL As String, objConnection As Object, Optional lngEofValue As Long = -1) As Long 'DELETE 2016/04/04 K.ISHIZAKA
Public Function ADODB_ExecGetLong(ByVal strSQL As String, objConnection As Object, Optional lngEofValue As Long = -1, Optional lTimeout As Variant) As Long 'INSERT 2016/04/04 K.ISHIZAKA
    Dim objRst              As Object
    On Error GoTo ErrorHandler

'    Set objRst = ADODB_Recordset(strSQL, objConnection)                        'DELETE 2016/04/04 K.ISHIZAKA
    Set objRst = ADODB_Recordset(strSQL, objConnection, adoReadOnly, lTimeout)  'DELETE 2016/04/04 K.ISHIZAKA
    On Error GoTo ErrorHandler1
    With objRst
        If .EOF Then
            ADODB_ExecGetLong = lngEofValue
        Else
            ADODB_ExecGetLong = CLng(Nz(.Fields(0).VALUE, 0))
        End If
        .Close
    End With
Exit Function

ErrorHandler1:
    objRst.Close
ErrorHandler:
    Call Err.Raise(Err.Number, "ADODB_ExecGetLong" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ADO接続でＳＱＬを実行し先頭フィールドをVariant型で取得
'       MODULE_ID       : ADODB_ExecGetVariant
'       CREATE_DATE     : 2008/12/08            S.SHIBAZAKI
'       PARAM           : strSQL                ＳＱＬ文字列(I)
'                       : objConnection         コネクションオブジェクト(I)
'                       : [varEofValue]         EOFだったときの戻り値(I)省略時:Null
'       RETURN          : 先頭フィールドの内容(Variant)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function ADODB_ExecGetVariant(ByVal strSQL As String, objConnection As Object, Optional varEofValue As Variant = Null) As Variant 'DELETE 2016/04/04 K.ISHIZAKA
Public Function ADODB_ExecGetVariant(ByVal strSQL As String, objConnection As Object, Optional varEofValue As Variant = Null, Optional lTimeout As Variant) As Variant 'INSERT 2016/04/04 K.ISHIZAKA
    Dim objRst              As Object
    On Error GoTo ErrorHandler

'    Set objRst = ADODB_Recordset(strSQL, objConnection)                        'DELETE 2016/04/04 K.ISHIZAKA
    Set objRst = ADODB_Recordset(strSQL, objConnection, adoReadOnly, lTimeout)  'DELETE 2016/04/04 K.ISHIZAKA
    On Error GoTo ErrorHandler1
    With objRst
        If .EOF Then
            ADODB_ExecGetVariant = varEofValue
        Else
            ADODB_ExecGetVariant = .Fields(0).VALUE
        End If
        .Close
    End With
Exit Function

ErrorHandler1:
    objRst.Close
ErrorHandler:
    Call Err.Raise(Err.Number, "ADODB_ExecGetVariant" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : テーブル削除
'       MODULE_ID       : DropTableSQL
'       CREATE_DATE     : 2010/06/12            K.ISHIZAKA
'       PARAM           : strTableName          テーブル名(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DropTableSQL(ByVal strTableName As String) As String
    Dim strSQL              As String

    If Left(strTableName, 1) = "#" Then                                         'INSERT 2014/06/09 K.ISHIZAKA
        strSQL = strSQL & "IF OBJECT_ID('tempdb.." & strTableName & "') IS NOT NULL" 'INSERT 2014/06/09 K.ISHIZAKA
    Else                                                                        'INSERT 2014/06/09 K.ISHIZAKA
        strSQL = strSQL & "IF EXISTS(SELECT * FROM SYSOBJECTS WHERE NAME = '" & GetSourceTableName(strTableName) & "')"
    End If                                                                      'INSERT 2014/06/09 K.ISHIZAKA
    strSQL = strSQL & vbCrLf                                                    'INSERT 2014/06/09 K.ISHIZAKA
    strSQL = strSQL & " DROP TABLE " & strTableName
    strSQL = strSQL & vbCrLf                                                    'INSERT 2014/06/09 K.ISHIZAKA
    strSQL = strSQL & ";"

    DropTableSQL = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : テーブル名の括弧をはずす  [WORK-TABLE]みたいになってるやつ
'       MODULE_ID       : GetSourceTableName
'       CREATE_DATE     : 2010/06/12            K.ISHIZAKA
'       PARAM           : strTableName          テーブル名(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetSourceTableName(ByVal strTableName As String) As String
    If (Left(strTableName, 1) = "[") And (Right(strTableName, 1) = "]") Then
        GetSourceTableName = Mid(strTableName, 2, Len(strTableName) - 2)
    Else
        GetSourceTableName = strTableName
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : ADO接続でテーブルを削除する
'       MODULE_ID       : ADODB_DropTable
'       CREATE_DATE     : 2010/06/12            K.ISHIZAKA
'       PARAM           : strTableName          テーブル名(I)
'                       : [objConnection]       コネクションオブジェクト：省略可(I)
'                                                   省略時 KASE_DB に対して実行する
'                       : [lTimeout]            タイムアウト値：省略可(I)
'       RETURN          : 対象件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function ADODB_DropTable(ByVal strTableName As String, Optional objConnection As Object = Nothing, Optional lTimeout As Long = C_DEFAULT_TIMEOUT) As Long 'DELETE 2016/04/04 K.ISHIZAKA
Public Function ADODB_DropTable(ByVal strTableName As String, Optional objConnection As Object = Nothing, Optional lTimeout As Variant) As Long 'INSERT 2016/04/04 K.ISHIZAKA
    Dim lngErrNumber        As Long
    Dim strErrSource        As String
    Dim strErrDescription   As String
    Dim strErrHelpFile      As String
    Dim lngErrHelpContext   As Long
    
    If Err.Number = 0 Then
        On Error GoTo ErrorHandler
        ADODB_DropTable = ADODB_Execute(DropTableSQL(strTableName), objConnection, lTimeout)
    Else
        lngErrNumber = Err.Number
        strErrSource = Err.Source
        strErrDescription = Err.Description
        strErrHelpFile = Err.HelpFile
        lngErrHelpContext = Err.HelpContext
        On Error Resume Next
        ADODB_DropTable = ADODB_Execute(DropTableSQL(strTableName), objConnection, lTimeout)
        Call Err.Raise(lngErrNumber, strErrSource, strErrDescription, strErrHelpFile, lngErrHelpContext)
    End If
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "ADODB_DropTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ＳＱＬの結果でワークテーブルをカレントDB(Access側)に作成する
'                         ＳＱＬで指定した列以外に[ROW_NUM]列が追加され読込順が格納される
'       MODULE_ID       : CreatWorkTable
'       CREATE_DATE     : 2011/12/28            K.ISHIZAKA
'       PARAM           : strTableName          テーブル名(I)
'                       : strSQL                実行ＳＱＬ(I)
'                       : [objConnection]       コネクションオブジェクト：省略可(I)
'                                                   省略時 KASE_DB に対して実行する
'       RETURN          : 対象件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function ADODB_CreateWorkTable_CurrentDb(ByVal strTableName As String, ByVal strSQL As String, Optional objConnection As Object = Nothing) As Long 'DELETE 2016/04/04 K.ISHIZAKA
Public Function ADODB_CreateWorkTable_CurrentDb(ByVal strTableName As String, ByVal strSQL As String, Optional objConnection As Object = Nothing, Optional lTimeout As Variant) As Long 'INSERT 2016/04/04 K.ISHIZAKA
    Dim lngCnt              As Long
    Dim objCon              As Object
    On Error GoTo ErrorHandler
    
    If objConnection Is Nothing Then
        Set objCon = ADODB_Connection()
        On Error GoTo ErrorHandler2
'        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objCon)       'DELETE 2012/03/02 K.ISHIZAKA
'        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objCon, True) 'DELETE 2016/04/04 K.ISHIZAKA 'INSERT 2012/03/02 K.ISHIZAKA
        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objCon, True, lTimeout) 'INSERT 2016/04/04 K.ISHIZAKA
        objCon.Close
        On Error GoTo ErrorHandler
    Else
'        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objConnection) 'DELETE 2012/03/02 K.ISHIZAKA
'        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objConnection, True) 'DELETE 2016/04/04 K.ISHIZAKA 'INSERT 2012/03/02 K.ISHIZAKA
        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objConnection, True, lTimeout) 'INSERT 2016/04/04 K.ISHIZAKA
    End If
    ADODB_CreateWorkTable_CurrentDb = lngCnt
Exit Function

ErrorHandler2:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ADODB_CreateWorkTable_CurrentDb" & vbRightAllow & Err.Source, Err.Description & vbCrLf & "SQL:" & vbCrLf & strSQL, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ＳＱＬの結果でカレントDB(Access側)のワークテーブル内容を入れ替える
'                         ＳＱＬで指定した列以外に[ROW_NUM]列が追加され読込順が格納される
'       MODULE_ID       : CreatWorkTable
'       CREATE_DATE     : 2012/03/02            K.ISHIZAKA
'       PARAM           : strTableName          テーブル名(I)
'                       : strSQL                実行ＳＱＬ(I)
'                       : [objConnection]       コネクションオブジェクト：省略可(I)
'                                                   省略時 KASE_DB に対して実行する
'       RETURN          : 対象件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function ADODB_Reset_CurrentDb(ByVal strTableName As String, ByVal strSQL As String, Optional objConnection As Object = Nothing) As Long 'DELETE 2016/04/04 K.ISHIZAKA
Public Function ADODB_Reset_CurrentDb(ByVal strTableName As String, ByVal strSQL As String, Optional objConnection As Object = Nothing, Optional lTimeout As Variant) As Long 'INSERT 2016/04/04 K.ISHIZAKA
    Dim lngCnt              As Long
    Dim objCon              As Object
    On Error GoTo ErrorHandler
    
    If objConnection Is Nothing Then
        Set objCon = ADODB_Connection()
        On Error GoTo ErrorHandler2
'        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objCon, False) 'DELETE 2016/04/04 K.ISHIZAKA
        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objCon, False, lTimeout) 'INSERT 2016/04/04 K.ISHIZAKA
        objCon.Close
        On Error GoTo ErrorHandler
    Else
'        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objConnection, False) 'DELETE 2016/04/04 K.ISHIZAKA
        lngCnt = CreateWorkTable_CurrentDb(strTableName, strSQL, objConnection, False, lTimeout) 'INSERT 2016/04/04 K.ISHIZAKA
    End If
    ADODB_Reset_CurrentDb = lngCnt
Exit Function

ErrorHandler2:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ADODB_Reset_CurrentDb" & vbRightAllow & Err.Source, Err.Description & vbCrLf & "SQL:" & vbCrLf & strSQL, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ＳＱＬの結果でワークテーブルをカレントDB(Access側)に作成する
'       MODULE_ID       : CreateWorkTable_CurrentDb
'       CREATE_DATE     : 2011/12/28            K.ISHIZAKA
'       PARAM           : strTableName          テーブル名(I)
'                       : strSQL                実行ＳＱＬ(I)
'                       : objCon                コネクションオブジェクト(I)
'       RETURN          : 対象件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function CreateWorkTable_CurrentDb(ByVal strTableName As String, ByVal strSQL As String, objCon As Object) As Long 'DELETE 2012/03/02 K.ISHIZAKA
'Private Function CreateWorkTable_CurrentDb(ByVal strTableName As String, ByVal strSQL As String, objCon As Object, ByVal blCreate As Boolean) As Long 'DELETE 2016/04/04 K.ISHIZAKA 'INSERT 2012/03/02 K.ISHIZAKA
Private Function CreateWorkTable_CurrentDb(ByVal strTableName As String, ByVal strSQL As String, objCon As Object, ByVal blCreate As Boolean, Optional lTimeout As Variant) As Long 'INSERT 2016/04/04 K.ISHIZAKA
    Dim objRst              As Object
    Dim objFld              As Object
    Dim rstWk               As Recordset
    Dim i                   As Long
    On Error GoTo ErrorHandler

'    Set objRst = ADODB_Recordset(strSQL, objCon)                               'DELETE 2016/04/04 K.ISHIZAKA
    Set objRst = ADODB_Recordset(strSQL, objCon, adoReadOnly, lTimeout)         'INSERT 2016/04/04 K.ISHIZAKA
    On Error GoTo ErrorHandler1
    If blCreate Then                                                            'INSERT 2012/03/02 K.ISHIZAKA
        Call CreatWorkTable(objRst, strTableName)
    Else                                                                        'INSERT 2012/03/02 K.ISHIZAKA
        CurrentDb.Execute "DELETE FROM " & strTableName                         'INSERT 2012/03/02 K.ISHIZAKA
    End If                                                                      'INSERT 2012/03/02 K.ISHIZAKA
    Set rstWk = CurrentDb.OpenRecordset(strTableName, dbOpenDynaset, dbAppendOnly)
    On Error GoTo ErrorHandler2
    i = 0
    While Not objRst.EOF
        rstWk.AddNew
        On Error GoTo ErrorHandler3
        For Each objFld In objRst.Fields
            With objFld
                rstWk.Fields(.NAME).VALUE = .VALUE
            End With
        Next
        rstWk.Fields("ROW_NUM").VALUE = i
        rstWk.UPDATE
        On Error GoTo ErrorHandler2
        objRst.MoveNext
        i = i + 1
    Wend
    rstWk.Close
    On Error GoTo ErrorHandler1
    objRst.Close
    On Error GoTo ErrorHandler
    CreateWorkTable_CurrentDb = i
Exit Function

ErrorHandler3:
    rstWk.CancelUpdate
ErrorHandler2:
    rstWk.Close
ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "CreateWorkTable_CurrentDb" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : レコードセットと同様のワークテーブルをカレントDB(Access側)に作成する
'       MODULE_ID       : CreatWorkTable
'       CREATE_DATE     : 2011/12/28            K.ISHIZAKA
'       PARAM           : objRst                レコードセット(I)
'                       : strTableName          テーブル名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CreatWorkTable(objRst As Object, ByVal strTableName As String)
    Dim fldNew              As Field
    Dim tdfNew              As TableDef
    Dim objFld              As Object
    On Error GoTo ErrorHandler

'    Call DropWorkTable(strTableName)                                           'DELETE 2012/03/02 K.ISHIZAKA
    Call DropWorkTable_CurrentDb(strTableName)                                  'INSERT 2012/03/02 K.ISHIZAKA
    Set tdfNew = CurrentDb().CreateTableDef(strTableName)
    For Each objFld In objRst.Fields
        With objFld
            If .Type = AdoEnums_DataType.adNumeric Then
                If .NumericScale > 0 Then
                    Set fldNew = tdfNew.CreateField(.NAME, dbDouble)
                Else
                    Set fldNew = tdfNew.CreateField(.NAME, dbLong)
                End If
            ElseIf .Type = AdoEnums_DataType.adDBTimeStamp Then                 'INSERT 2015/08/12 K.ISHIZAKA
                Set fldNew = tdfNew.CreateField(.NAME, dbDate)                  'INSERT 2015/08/12 K.ISHIZAKA
            ElseIf .Type = AdoEnums_DataType.adInteger Then
                Set fldNew = tdfNew.CreateField(.NAME, dbLong)                  'INS 2019/09/05 M.HONDA
            Else
                Set fldNew = tdfNew.CreateField(.NAME, dbText, .DefinedSize)
                fldNew.AllowZeroLength = True
            End If
        End With
        tdfNew.Fields.Append fldNew
    Next
    tdfNew.Fields.Append tdfNew.CreateField("ROW_NUM", dbLong)
    CurrentDb().TableDefs.Append tdfNew
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "CreatWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブル削除
'       MODULE_ID       : DropWorkTable
'       CREATE_DATE     : 2011/12/28            K.ISHIZAKA
'       PARAM           : strTableName          テーブル名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub DropWorkTable(ByVal strTableName As String)                        'DELETE 2012/03/02 K.ISHIZAKA
Public Sub DropWorkTable_CurrentDb(ByVal strTableName As String)                'INSERT 2012/03/02 K.ISHIZAKA
    Dim tdf                 As TableDef
    On Error GoTo ErrorHandler

    On Error Resume Next
    Set tdf = CurrentDb().TableDefs(strTableName)
    If Err.Number = 3265 Then
        Err.Clear                                                               'INSERT 2012/03/02 K.ISHIZAKA
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    CurrentDb().TableDefs.Delete strTableName
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "DropWorkTable_CurrentDb" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended of program ********************************


