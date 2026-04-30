Attribute VB_Name = "MSZZ002"
'****************************  strat of program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    :
'        PROGRAM_ID      : MSZZ002
'        PROGRAM_KBN     :
'
'        CREATE          : 2002/09/20
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2003/02/10
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'
'        UPDATE          : 2003/09/08
'        UPDATER         : N.MIURA
'        Ver             : 0.2
'
'        UPDATE          : 2005/01/07
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.3
'           パスワードを保存しておく属性を付加する
'
'        UPDATE          : 2005/02/05
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.4
'           正常の戻り値を０にする
'
'        UPDATE          : 2005/09/13
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.5
'           接続文字列の修正
'
'        UPDATE          : 2008/07/30
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.6
'           MSZZ002_M20のパラメータにユーザーＩＤとパスワードを追加
'
'       UPDATE          : 2011/05/30
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.7
'                       : ワークスペースを閉じると、他で開いている
'                         レコードセットなどにも影響を及ぼすので
'                         カレントＤＢに変更する
'
'       UPDATE          : 2011/10/31
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.8
'                       : MSZZ002_LinkOn MSZZ002_LinkOff を追加する
'                         テーブルリンクとリンク解除がエラーハンドラー内で呼び出されても良いようにする
'
'==============================================================================*
Option Explicit

'==============================================================================*
'
'       MODULE_NAME     : テーブルをリンクする
'       MODULE_ID       : MSZZ002_LinkOn
'       CREATE_DATE     : 2011/10/31            K.ISHIZAKA
'       PARAM           : strTableId            テーブルID(I)
'                       : [strBumoc]            接続部門コード(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ002_LinkOn(ByVal strTableId As String, Optional ByVal strBUMOC As String = "")
    Call TableLink(strTableId, True, strBUMOC)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : テーブルリンクを解除する
'       MODULE_ID       : MSZZ002_LinkOff
'       CREATE_DATE     : 2011/10/31            K.ISHIZAKA
'       PARAM           : strTableId            テーブルID(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ002_LinkOff(ByVal strTableId As String)
    Call TableLink(strTableId, False, "")
End Sub

'==============================================================================*
'
'       MODULE_NAME     : テーブルをリンクするor解除する
'       MODULE_ID       : TableLink
'       CREATE_DATE     : 2011/10/31            K.ISHIZAKA
'       PARAM           : strTableId            テーブルID(I)
'                       : blOnOff               オン(True)／オフ(False)
'                       : [strBumoc]            接続部門コード(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TableLink(ByVal strTableId As String, ByVal blOnOff As Boolean, Optional ByVal strBUMOC As String = "")
    Dim strDNS              As String
    Dim strSvr              As String
    Dim strDBN              As String
    Dim strUid              As String
    Dim strPwd              As String
    Dim strErrMessage       As String
    Dim strBumocProc        As String
    Dim lngErrNumber        As Long
    Dim strErrSource        As String
    Dim strErrDescription   As String
    Dim strErrHelpFile      As String
    Dim lngErrHelpContext   As Long
    
    'エラーハンドラーからの呼び出しを可能にする
    If Err.Number = 0 Then
        lngErrNumber = 0
        On Error GoTo ErrorHandler
    Else
        lngErrNumber = Err.Number
        strErrSource = Err.Source
        strErrDescription = Err.Description
        strErrHelpFile = Err.HelpFile
        lngErrHelpContext = Err.HelpContext
        On Error Resume Next
    End If
    
    strTableId = MSZZ025.GetSourceTableName(strTableId)
    'リンクをはずす
    If MSZZ002_M00(strTableId, strErrMessage) <> 0 Then
        If blOnOff Or (lngErrNumber <> 0) Then 'リンクするときとエラーハンドル中はエラーを無視する
            strErrMessage = ""
        Else
            Call MSZZ024_M10("MSZZ002_M00", strErrMessage)
        End If
    End If
    
    'リンクする
    If blOnOff Then
        strBumocProc = IIf(strBUMOC <> "", "_" & strBUMOC, "")
    
        strDNS = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATA_SOURCE_NAME" & strBumocProc & "'")
        strSvr = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_SERVER_NAME" & strBumocProc & "'")
        strDBN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME" & strBumocProc & "'")
        strUid = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_USER_ID" & strBumocProc & "'")
        strPwd = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_PASSWORD" & strBumocProc & "'")
    
        If MSZZ002_M20(strDNS, strSvr, strDBN, strUid, strPwd, "dbo." & strTableId, strTableId, strErrMessage) <> 0 Then
            If lngErrNumber = 0 Then
                Call MSZZ024_M10("MSZZ002_M20", strErrMessage)
            End If
        End If
    End If

    'エラー内容を戻して、エラーハンドラーに返す
    If lngErrNumber <> 0 Then
        Call Err.Raise(lngErrNumber, strErrSource, strErrDescription, strErrHelpFile, lngErrHelpContext)
    End If
Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "MSZZ002.TableLink" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :MSZZ002_M00
'        CREATE_DATE      :2002/09/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ002_M00(tdef_alias As String, error_message As String) As Integer
    On Error GoTo fpDAT_Error:
    
'    MSZZ002_M00 = False                                                        'DELETE START 2011/05/30 K.ISHIZAKA
'
'    Dim ws As Workspace, db As Database, tdef As TableDef
'
'    Set ws = DBEngine.Workspaces(0)
'    Set db = ws.Databases(0)                                                   'DELETE END   2011/05/30 K.ISHIZAKA
    Dim db                  As Database                                         'INSERT 2011/05/30 K.ISHIZAKA
    
    Set db = CurrentDb                                                          'INSERT 2011/05/30 K.ISHIZAKA
    db.TableDefs.Delete tdef_alias
    db.TableDefs.Refresh                                                        'INSERT 2011/05/30 K.ISHIZAKA
'    'db.Close
'    ws.Close                                                                   'DELETE 2011/05/30 K.ISHIZAKA
    
    'MSZZ002_M00 = True                                                         'DEL 20050205 K.ISHIZAKA
    MSZZ002_M00 = 0                                                             'ADD 20050205 K.ISHIZAKA

    Exit Function

fpDAT_Error:

    error_message = error_message & "/MSZZ002_M00: " & CStr(Err.Number) & " " & Err.Description
    MSZZ002_M00 = Err.Number

    Exit Function

End Function
'==============================================================================*
'
'        MODULE_NAME      :リンク
'        MODULE_ID        :fpAppendAttachTable
'        CREATE_DATE      :2002/09/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSZZ002_M10(database_name As String, tdef_name As String, tdef_alias As String, error_message As String) As Integer
    On Error GoTo fpAAT_Error:
    
'    MSZZ002_M10 = False                                                        'DELETE START 2011/05/30 K.ISHIZAKA
'
'    Dim ws As Workspace, db As Database, tdef As TableDef, i As Integer
'
'    Set ws = DBEngine.Workspaces(0)
'    Set db = ws.Databases(0)                                                   'DELETE END   2011/05/30 K.ISHIZAKA
    Dim tdef                As TableDef                                         'INSERT 2011/05/30 K.ISHIZAKA
    Dim db                  As Database                                         'INSERT 2011/05/30 K.ISHIZAKA
    
    Set db = CurrentDb                                                          'INSERT 2011/05/30 K.ISHIZAKA
    Set tdef = db.CreateTableDef(tdef_alias)
    tdef.Connect = ";DATABASE=" & database_name
    tdef.SourceTableName = tdef_name
    db.TableDefs.Append tdef
    db.TableDefs.Refresh
    'db.Close
'    ws.Close                                                                   'DELETE 2011/05/30 K.ISHIZAKA
    
    'MSZZ002_M10 = True                                                         'DEL 20050205 K.ISHIZAKA
    MSZZ002_M10 = 0                                                             'ADD 20050205 K.ISHIZAKA

    Exit Function

fpAAT_Error:

    error_message = error_message & "/MSZZ002_M10: " & CStr(Err.Number) & " " & Err.Description
    MSZZ002_M10 = Err.Number

    Exit Function

End Function
'==============================================================================*
'
'        MODULE_NAME      :リンク
'        MODULE_ID        :MSZZ002_M20
'        CREATE_DATE      :2003/09/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'UPDATE 2008/07/30 SHIBAZAKI login_idとpasswordを追加
Function MSZZ002_M20(dsn_name As String, _
                     server_name As String, _
                     databese_name As String, _
                     login_id As String, _
                     password As String, _
                     tdef_name As String, _
                     tdef_alias As String, _
                     error_message As String) As Integer
    On Error GoTo fpAAT4_Error:
    
    MSZZ002_M20 = False
    
'    Dim ws        As Workspace                                                 'DELETE 2011/05/30 K.ISHIZAKA
    Dim db        As Database
    Dim tdef      As TableDef
    Dim i         As Integer
    Dim strPARAN  As String
    Dim strPassword     As String                                               'INSERT 2008/07/30 SHIBAZAKI
    
    strPassword = IIf(password = "#NULL#", "", password)                        'INSERT 2008/07/30 SHIBAZAKI
    
'    Set ws = DBEngine.Workspaces(0)                                            'DELETE 2011/05/30 K.ISHIZAKA
'    Set db = ws.Databases(0)                                                   'DELETE 2011/05/30 K.ISHIZAKA
    Set db = CurrentDb                                                          'INSERT 2011/05/30 K.ISHIZAKA
    Set tdef = db.CreateTableDef(tdef_alias)
    
    strPARAN = ""
    strPARAN = strPARAN & "ODBC;"
    strPARAN = strPARAN & "DSN=" & dsn_name & ";"
    strPARAN = strPARAN & "SERVER="                                             '20050913 ADD K.ISHIZAKA
    strPARAN = strPARAN & server_name & ";"
'    strPARAN = strPARAN & "UID=SA;"                                            'DELETE 2008/07/30 SHIBAZAKI
'    strPARAN = strPARAN & "PWD=;"                                              'DELETE 2008/07/30 SHIBAZAKI
    strPARAN = strPARAN & "UID=" & login_id & ";"                               'INSERT 2008/07/30 SHIBAZAKI
    strPARAN = strPARAN & "PWD=" & strPassword & ";"                            'INSERT 2008/07/30 SHIBAZAKI
    strPARAN = strPARAN & "DATABASE=" & databese_name
    tdef.Connect = strPARAN
    
    tdef.SourceTableName = tdef_name
    
    tdef.Attributes = dbAttachSavePWD                                           '20050107 ADD K.ISHIZAKA
    
    db.TableDefs.Append tdef
    db.TableDefs.Refresh
    
    'db.Close
'    ws.Close                                                                   'DELETE 2011/05/30 K.ISHIZAKA
    
    'MSZZ002_M20 = True                                                         'DEL 20050205 K.ISHIZAKA
    MSZZ002_M20 = 0                                                             'ADD 20050205 K.ISHIZAKA

    Exit Function

fpAAT4_Error:

    error_message = error_message & "/MSZZ002_M20: " & CStr(Err.Number) & " " & Err.Description
    MSZZ002_M20 = Err.Number

    Exit Function

End Function

'****************************  ended of program ********************************
