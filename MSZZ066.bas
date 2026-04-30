Attribute VB_Name = "MSZZ066"
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : WebServer MySQL 関数（SSH2.0対応）
'       PROGRAM_ID      : MSZZ066
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2011/05/21
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2012/06/16
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : MySql_ChangeUserLevel の引数に接続先を指定可能にする
'                       : MySqlRecordSetのフィールドのクラス化（ForEach や With や .Name 、.Value に対応）
'
'       UPDATE          : 2015/01/06
'       UPDATER         : M.HONDA
'       Ver             : 0.2
'                       : kshが利用できないためsh変更
'
'       UPDATE          : 2022/12/26
'       UPDATER         : N.IMAI
'       Ver             : 0.3
'                       : AWS対応
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID               As String = "MSZZ066"

'INTI_FILE の INTIF_RECFB（キーとなる文字列）
'Private Const C_LOCAL_PATH          As String = "LocalPath"
'Private Const C_REMOTE_PATH         As String = "RemotePath"
'Private Const C_MYSQL_DATABASE1     As String = "MysqlDataBase1"
'Private Const C_MYSQL_DATABASE2     As String = "MysqlDataBase2"
Private Const C_MYSQL_LOGIN_PATH    As String = "MysqlLoginPath"
'Private Const C_MYSQL_LOGIN_FILE    As String = "MysqlLoginFile"
Private Const C_MYSQL_LOGIN_USER    As String = "MysqlLoginUser"
Private Const C_MYSQL_LOGIN_PWD     As String = "MysqlLoginPassword"
Private Const C_MYSQL_SUPER_USER    As String = "MysqlSuperUser"
Private Const C_MYSQL_SUPER_PWD     As String = "MysqlSuperPassword"

'Web 上の MySQL にログインする権限
Public Enum MySql_UserLevel
    LoginUser               '通常の権限
    SuperUser               '
End Enum

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private strPuttyCommand      As String
Private strMySqlCommand      As String
Private colIntiFile          As Collection

'==============================================================================*
'   テスト
'==============================================================================*
Sub TEST_MySql()
    Dim strSQL              As String
    Dim objMySQL            As MySqlRecordSet
    Dim strFld              As String
    Dim i                   As Long
    On Error GoTo ErrorHandler
    
    Call MySql_ChangeUserLevel(LoginUser)
    strSQL = "SELECT * FROM TOGO_CTRL_TABL WHERE CTRLT_CODE = 'KISAN_PTN_1'"
    Set objMySQL = MySql_OpenRecordSet(strSQL)
    On Error GoTo ErrorHandler1
    While Not objMySQL.EOF
        For i = 0 To objMySQL.Fields.Count - 1
            strFld = objMySQL.Fields(i).NAME               'objMySQL.GetFieldName(i) でもOK
            Debug.Print strFld
            
            Debug.Print objMySQL.Fields(i).VALUE      'インデックスで取得  objMySQL.Fields(i) は使用できません
            Debug.Print objMySQL.Fields(strFld).VALUE 'フィールド名で取得  objMySQL.Fields(strFld) は使用できません
            
            With objMySQL.Fields(i)
                Debug.Print .NAME
                Debug.Print .VALUE
            End With
            With objMySQL.Fields(strFld)
                Debug.Print .NAME
                Debug.Print .VALUE
            End With
        Next
        Dim objFld As MySqlField
        For Each objFld In objMySQL.Fields
            With objFld
                Debug.Print .NAME
                Debug.Print .VALUE
            End With
        Next
        objMySQL.MoveNext   '次のレコード
    Wend
    
'    With objMySQL
'        If Not .BOF Then
'            .MoveFirst  '先頭レコードへ移動
'            While Not .EOF
'                Debug.Print "(" & .Fields("YARD_CODE").VALUE & "-" & .Fields("YARD_USAGE").VALUE & ")" & .Fields("YARD_NAME").VALUE  'フィールド名で
'                .MoveNext
'            Wend
'        End If
'        .Close_
'        On Error GoTo ErrorHandler
'    End With
    Set objMySQL = Nothing
Exit Sub

ErrorHandler1:
    objMySQL.Close_
ErrorHandler:          '↓自分の関数名
    MsgBox "(" & Err.Number & ")" & Err.Description, vbOKOnly + vbExclamation, "TEST_MySql"
End Sub

Sub TEST_MySql2()
    Dim strSQL              As String
    Dim objMySQL            As MySqlRecordSet
    Dim strFld              As String
    Dim i                   As Long
    On Error GoTo ErrorHandler
    
    strSQL = "SELECT TOGO_YARD_MAST.* FROM TOGO_YARD_MAST ORDER BY YARD_CODE LIMIT 10 "
    Set objMySQL = MySql_OpenRecordSet(strSQL)
    On Error GoTo ErrorHandler1
    While Not objMySQL.EOF
        For i = 0 To objMySQL.Fields.Count - 1
            strFld = objMySQL.GetFieldName(i)   'objMySQL.Fields(i).Name でもOK
            Debug.Print strFld
            Debug.Print objMySQL.Fields(i).VALUE      'インデックスで取得  objMySQL.Fields(i) は使用できません
            Debug.Print objMySQL.Fields(strFld).VALUE 'フィールド名で取得  objMySQL.Fields(strFld) は使用できません
        Next
        objMySQL.MoveNext   '次のレコード
    Wend
    
    With objMySQL
        If Not .BOF Then
            .MoveFirst  '先頭レコードへ移動
            While Not .EOF
                Debug.Print "(" & .Fields("YARD_CODE").VALUE & "-" & .Fields("YARD_USAGE").VALUE & ")" & .Fields("YARD_NAME").VALUE  'フィールド名で
                .MoveNext
            Wend
        End If
        .Close_
        On Error GoTo ErrorHandler
    End With
    Set objMySQL = Nothing
Exit Sub

ErrorHandler1:
    objMySQL.Close_
ErrorHandler:          '↓自分の関数名
    MsgBox "(" & Err.Number & ")" & Err.Description, vbOKOnly + vbExclamation, "TEST_MySql"
End Sub
'==============================================================================*
'
'       MODULE_NAME     : Web 上の MySQL にログインする権限の変更
'       MODULE_ID       : MySql_ChangeUserLevel
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : userLevel             権限(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub MySql_ChangeUserLevel(ByVal userLevel As MySql_UserLevel)           'DELETE 2012/06/16 K.ISHIZAKA
Public Sub MySql_ChangeUserLevel(ByVal userLevel As MySql_UserLevel, Optional ByVal strSETUB As String = "") 'INSERT 2012/06/16 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    If colIntiFile Is Nothing Then
        Set colIntiFile = MSZZD00_RECDB_Collection("MSTGP02")
    End If

    strPuttyCommand = Environ("ComSpec") & " /c "
    strPuttyCommand = strPuttyCommand & MSZZ064.getPuttyPath() & "plink" & MSZZ064.connectString(strSETUB)

    'strMySqlCommand = "/bin/ksh " & colIntiFile(C_MYSQL_LOGIN_PATH) & "mysqlrun.ksh "
    'strMySqlCommand = "/bin/sh " & colIntiFile(C_MYSQL_LOGIN_PATH) & "mysqlrun.ksh "   'DELETE 2022/12/26 N.IMAI
    strMySqlCommand = colIntiFile(C_MYSQL_LOGIN_PATH)                                   'INSERT 2022/12/26 N.IMAI
'以下mysqlrun.kshの中身 MSTGP02 にてアクティブデータベースを切り替えてるとこに追加する
'mysql -u $1 -p$2 -D KASE_DB1 << EOF
'$3
'\
'EOF
'DELETE 2022/12/26 N.IMAI Start
'    Select Case userLevel
'    Case MySql_UserLevel.LoginUser
'        strMySqlCommand = strMySqlCommand & " " & colIntiFile(C_MYSQL_LOGIN_USER) & " " & colIntiFile(C_MYSQL_LOGIN_PWD)
'    Case MySql_UserLevel.SuperUser
'        strMySqlCommand = strMySqlCommand & " " & colIntiFile(C_MYSQL_SUPER_USER) & " " & colIntiFile(C_MYSQL_SUPER_PWD)
'    End Select
'DELETE 2022/12/26 N.IMAI End
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "MySql_ChangeUserLevel" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : Web 上の MySQL でＳＱＬを実行しレコードセットを開く
'       MODULE_ID       : MySql_OpenRecordSet
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : strSQL()              ＳＱＬ文字列(I)
'       RETURN          : レコードセット(MySqlRecordSet)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MySql_OpenRecordSet(ParamArray strSQL() As Variant) As MySqlRecordSet
    Dim strCmd              As String
    Dim varSQL              As Variant
    Dim objMySQL            As New MySqlRecordSet
    On Error GoTo ErrorHandler
    
    strCmd = " """
    For Each varSQL In strSQL
        If InStr(varSQL, " * ") > 0 Then
'            Call MSZZ024_M10("MySql_OpenRecordSet", "全列指定のときはテーブル名を明確にしてください。")
        End If
        strCmd = strCmd & varSQL
        If Right(varSQL, 1) <> ";" Then
            strCmd = strCmd & ";"
        End If
    Next
    strCmd = strCmd & """"
    If colIntiFile Is Nothing Then
        Call MySql_ChangeUserLevel(LoginUser)
    End If
    Call objMySQL.Open_(strPuttyCommand, strMySqlCommand & strCmd)
    Set MySql_OpenRecordSet = objMySQL
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "MySql_OpenRecordSet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : Web 上の MySQL でＳＱＬを実行し先頭フィールドをLong型で取得
'       MODULE_ID       : MySql_ExecGetLong
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : strSQL                ＳＱＬ文字列(I)
'                       : [lngEofValue]         EOFだったときの戻り値(I)省略時:-1
'       RETURN          : 先頭フィールドの内容(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MySql_ExecGetLong(ByVal strSQL As String, Optional lngEofValue As Long = -1) As Long
    Dim objMySQL            As MySqlRecordSet
    On Error GoTo ErrorHandler

    Set objMySQL = MySql_OpenRecordSet(strSQL)
    On Error GoTo ErrorHandler1
    With objMySQL
        If .EOF Then
            MySql_ExecGetLong = lngEofValue
        Else
'            MySql_ExecGetLong = CLng(Nz(.Fields(0), 0))                        'DELETE 2012/06/16 K.ISHIZAKA
            MySql_ExecGetLong = CLng(Nz(.Fields(0).VALUE, 0))                   'INSERT 2012/06/16 K.ISHIZAKA
        End If
        .Close_
    End With
Exit Function

ErrorHandler1:
    objMySQL.Close_
ErrorHandler:
    Call Err.Raise(Err.Number, "MySql_ExecGetLong" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : Web 上の MySQL でＳＱＬを実行する
'       MODULE_ID       : MySql_Execute
'       CREATE_DATE     : 2011/05/21            K.ISHIZAKA
'       PARAM           : strSQL                ＳＱＬ文字列(I)
'       RETURN          : 対象件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MySql_Execute(ByVal strSQL As String) As Long
    Call MSZZ024_M10("MySql_Execute", "現在のバージョンではサポートされていません！")
End Function

'****************************  ended of program ********************************

