Attribute VB_Name = "MSHS010"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    : 起動ログ出力
'        PROGRAM_ID      : MSHS010
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2002/03/28
'        CERATER         : N.MIURA
'        Ver             : 0.0
'
'        UPDATE          : 2002/10/22
'        UPDATER         : N.MIURA
'        Ver             : 0.1
'                        : ステータス設定
'
'        UPDATE          : 2003/08/21
'        UPDATER         : N.MIURA
'        Ver             : 0.2
'                        : レコードセットに変更
'
'        UPDATE          : 2004/02/03
'        UPDATER         : N.MIURA
'        Ver             : 0.3
'                        : ＳＳ７切換対応
'
'        UPDATE          : 2004/11/17
'        UPDATER         : N.MIURA
'        Ver             : 0.7
'                        : 不具合修正
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private USER_ID         As String
'
Private wspWorkspace    As Workspace 'ワークスペース
'Private DBS             As Database                                            'DELETE 20041117 N,MIURA
'Private strSql          As String                                              'DELETE 20041117 N,MIURA
'
'Private RST_ECLG        As Recordset                                           'DELETE 20040203 N.MIURA
'
'==============================================================================*
'
'        MODULE_NAME      :テスト
'        MODULE_ID        :TEST_MSHS010_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function TEST_MSHS010_M00()
    
    If MSHS010_M00("TEST", "1") = True Then
       MsgBox ("正常")
    Else
       MsgBox ("異常")
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :メイン
'        MODULE_ID        :MSHS010_M00
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function MSHS010_M00(MSHS010_PROG_ID As String, MSHS010_KBN As String, Optional Work As String = "") As Boolean
    On Error GoTo MSHS010_M00_Err
   
    MSHS010_M00 = False
    
    Call MSHS010_M10
    Call MSHS010_M20(MSHS010_PROG_ID, MSHS010_KBN, Work)
    
    MSHS010_M00 = True

MSHS010_M00_Exit:
    Exit Function

MSHS010_M00_Err:
    MsgBox Error$
    Resume MSHS010_M00_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :前処理
'        MODULE_ID        :MSHS010_M10
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MSHS010_M10()
    On Error GoTo MSHS010_M10_Err
    
    USER_ID = LsGetUserName()

MSHS010_M10_Exit:
    Exit Sub

MSHS010_M10_Err:
    MsgBox Err.Description
    Resume MSHS010_M10_Exit
End Sub
'==============================================================================*
'
'        MODULE_NAME      :主処理
'        MODULE_ID        :MSHS010_M20
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MSHS010_M20(PROG_ID As String, kbn As String, Optional Work As String = "")
    On Error GoTo MSHS010_M20_Err
    
    Dim strSS7              As String
    Dim dbs                 As Database
    Dim strSQL              As String
    
    strSS7 = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"), "")
    If strSS7 = "" Then
        MsgBox "中止します。" & vbCrLf & "テーブル[SETU_TABL]の設定不正です。", vbInformation
        Set dbs = Nothing
    Else
        Set dbs = Workspaces(0).OpenDatabase(strSS7, dbDriverNoPrompt, False, MSZZ007_M00())
    End If
    
    strSQL = ""
    strSQL = strSQL & "INSERT INTO "
    strSQL = strSQL & "ECLG_TABL "
    strSQL = strSQL & "( "
    strSQL = strSQL & "ECLGT_INSED, "
    strSQL = strSQL & "ECLGT_INSEJ, "
    strSQL = strSQL & "ECLGT_INSPB, "
    strSQL = strSQL & "ECLGT_INSUB, "
    strSQL = strSQL & "ECLGT_ECLGI, "
    strSQL = strSQL & "ECLGT_ETEXT  "
    strSQL = strSQL & ") "
    strSQL = strSQL & "VALUES "
    strSQL = strSQL & "( "
    strSQL = strSQL & "'" & Format(Now(), "yyyymmdd") & "',"
    strSQL = strSQL & "'" & Format(Now(), "hhmmss") & "',"
    strSQL = strSQL & "'" & PROG_ID & "',"
    strSQL = strSQL & "'" & USER_ID & "',"
    strSQL = strSQL & "'" & kbn & "', "
    strSQL = strSQL & "'" & Work & "' "
    strSQL = strSQL & ") "
    strSQL = strSQL & ";"
    dbs.Execute strSQL, dbSQLPassThrough
    
    dbs.Close
    
    Set dbs = Nothing
       
MSHS010_M20_Exit:
    Exit Sub

MSHS010_M20_Err:
    MsgBox Err.Description
    Resume MSHS010_M20_Exit
End Sub
'==============================================================================*
'
'        MODULE_NAME      :主処理
'        MODULE_ID        :MSHS010_M20
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub MSHS010_M20(PROG_ID As String, KBN As String)
'    On Error GoTo MSHS010_M20_Err
'    Set DBS = CurrentDb()
'    strSql = ""                                                                'INSERT 20040203 N.MIURA START
'    strSql = strSql & "INSERT INTO "
'    strSql = strSql & "ECLG_TABL "
'    strSql = strSql & "( "
'    strSql = strSql & "ECLGT_INSED, "
'    strSql = strSql & "ECLGT_INSEJ, "
'    strSql = strSql & "ECLGT_INSPB, "
'    strSql = strSql & "ECLGT_INSUB, "
'    strSql = strSql & "ECLGT_ECLGI  "
'    strSql = strSql & ") "
'    strSql = strSql & "VALUES "
'    strSql = strSql & "( "
'    strSql = strSql & "Format(Now(), 'yyyymmdd') , "
'    strSql = strSql & "Format(Now(), 'hhmmss')   , "
'    strSql = strSql & "'" & PROG_ID & "',"
'    strSql = strSql & "'" & USER_ID & "',"
'    strSql = strSql & "'" & KBN & "' "
'    strSql = strSql & ") "
'    strSql = strSql & ";"
'    DBS.Execute strSql, dbFailOnError                                           'INSERT 20040203 N.MIURA ENDED
'    'Set RST_ECLG = dbs.OpenRecordset("ECLG_TABL", dbOpenDynaset, dbAppendOnly) 'DELETE 20040203 N.MIURA START
'    'With RST_ECLG
'    '    .AddNew
'    '    .Fields("ECLGT_INSED").Value = Format(Now(), "yyyymmdd")
'    '    .Fields("ECLGT_INSEJ").Value = Format(Now(), "hhmmss")
'    '    .Fields("ECLGT_INSPB").Value = PROG_ID
'    '    .Fields("ECLGT_INSUB").Value = USER_ID
'    '    .Fields("ECLGT_ECLGI").Value = KBN
'    '   .Update
'    'End With
'    'RST_ECLG.CLOSE
'    'Set RST_ECLG = Nothing                                                     'DELETE 20040203 N.MIURA START
'    DBS.CLOSE
'    Set DBS = Nothing
'MSHS010_M20_Exit:
'    Exit Sub
'MSHS010_M20_Err:
'    MsgBox err.Description
'    Resume MSHS010_M20_Exit
'End Sub
'****************************  ended or program ********************************

