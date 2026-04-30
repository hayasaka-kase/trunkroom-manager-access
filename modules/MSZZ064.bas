Attribute VB_Name = "MSZZ064"
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬総合システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : SFTP、SSH（2.0対応）関数
'       PROGRAM_ID      : MSZZ064
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2011/02/17
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/02/21
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                         同期処理とし、戻り値はコマンドの戻り値にする
'
'       UPDATE          : 2011/05/21
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.2
'                         いくつかの関数をPrivateからPublicに変更する
'
'       UPDATE          : 2012/06/29
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.3
'                         パスの取得先
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID       As String = "MSZZ064"

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal PROCESS As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400&
Private Const STILL_ACTIVE = &H103&

'==============================================================================*
'   テスト
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_puttySftp()
    Dim strBatchFile As String
    Dim lngCnt As Long
    On Error GoTo ErrorHandler
    
'    strBatchFile = getPuttyPath() & "testbatch.txt"
'    lngCnt = puttySftp(strBatchFile)
'    Debug.Print lngCnt
    
    lngCnt = puttyLink("grep -e 'KASE_DB1' /home/kase3535/include/mysql.login.txt")
    'これは０が帰る
    Debug.Print lngCnt
    
    lngCnt = puttyLink("grep -e 'KASE_DB2' /home/kase3535/include/mysql.login.txt")
    'これは１が帰る
    Debug.Print lngCnt
    
    lngCnt = puttyLink("grep -e 'KASE_DB2' /home/kase3535/include/mysql.login3.txt")
    'これは２が帰る
    Debug.Print lngCnt
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("TEST_puttySftp", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : SFTP
'       MODULE_ID       : puttySftp
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       PARAM           : strBatchFile          バッチファイル(I)
'                       : [strSETUB]            接続部門(I)
'       RETURN          : コマンドの戻り値
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function puttySftp(ByVal strBatchFile As String, Optional ByVal strSETUB As String = "") As Long
    Dim strCommand          As String
    On Error GoTo ErrorHandler

    strCommand = getPuttyPath() & "psftp -b " & strBatchFile & connectString(strSETUB)
    puttySftp = ShellWScriptShell(strCommand)
'    puttySftp = ShellWinAPI(strCommand)
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "puttySftp" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : SSH
'       MODULE_ID       : puttyLink
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       PARAM           : strCommand            コマンド(I)
'                       : [strSETUB]            接続部門(I)
'       RETURN          : コマンドの戻り値
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function puttyLink(ByVal strCommand As String, Optional ByVal strSETUB As String = "") As Long
    On Error GoTo ErrorHandler

    strCommand = getPuttyPath() & "plink " & connectString(strSETUB) & " " & strCommand
    puttyLink = ShellWScriptShell(strCommand)
'    puttyLink = ShellWinAPI(strCommand)
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "puttyLink" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : PUTTYモジュールのパス取得
'       MODULE_ID       : getPuttyPath
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       RETURN          : PUTTYモジュールのパス
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function getPuttyPath() As String                                      'DELETE 2011/05/21 K.ISHIZAKA
Public Function getPuttyPath() As String                                        'INSERT 2011/05/21 K.ISHIZAKA
    Dim strPath             As String
    Const C_PUTTY_PATH      As String = "PuttyTool\"                            'INSERT 2012/06/29 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    strPath = Application.CurrentProject.path & "\" & C_PUTTY_PATH              'INSERT START 2012/06/29 K.ISHIZAKA
    If Dir(strPath, vbDirectory) <> "" Then
        getPuttyPath = strPath
        Exit Function
    End If                                                                      'INSERT END   2012/06/29 K.ISHIZAKA
    strPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & PROG_ID & "'"))
    If strPath = "" Then
        Call MSZZ024_M10("DLookup", "INTI_FILEの設定不足です。")
    End If
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    getPuttyPath = strPath
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "getPuttyPath" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : SSH接続文字列
'       MODULE_ID       : connectString
'       CREATE_DATE     : 2011/02/17            K.ISHIZAKA
'       PARAM           : strSETUB              接続部門(I)
'       RETURN          : SSH接続文字列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function connectString(ByVal strSETUB As String) As String             'DELETE 2011/05/21 K.ISHIZAKA
Public Function connectString(ByVal strSETUB As String) As String               'INSERT 2011/05/21 K.ISHIZAKA
    Dim strSvr              As String
    Dim strUid              As String
    Dim strPwd              As String
    On Error GoTo ErrorHandler
    
    If strSETUB <> "" Then
        strSETUB = "_" & strSETUB
    End If
    strSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'SFTP_SERVER_NAME" & strSETUB & "'"))
    strUid = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'SFTP_USER_ID" & strSETUB & "'"))
    strPwd = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'SFTP_PASSWORD" & strSETUB & "'"))
    If (strSvr = "") Or (strUid = "") Or (strPwd = "") Then
        Call MSZZ024_M10("DLookup", "SETU_TABLの設定不足です。")
    End If
    connectString = " -batch -pw " & strPwd & " " & strUid & "@" & strSvr
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "connectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : プロセスを実行し終了を待つ(WScript.Shell Version)
'       MODULE_ID       : ShellWScriptShell
'       CREATE_DATE     : 2011/02/26            K.ISHIZAKA
'       PARAM           : strCommand            コマンド(I)
'       RETURN          : プロセスの終了コード(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function ShellWScriptShell(ByVal strCommand As String) As Long         'DELETE 2011/05/21 K.ISHIZAKA
Public Function ShellWScriptShell(ByVal strCommand As String) As Long           'INSERT 2011/05/21 K.ISHIZAKA
    Dim objWss              As Object
    On Error GoTo ErrorHandler

    Set objWss = CreateObject("WScript.Shell")
    ShellWScriptShell = objWss.Run(strCommand, 0, True)
    Set objWss = Nothing
Exit Function

ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "ShellWScriptShell" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : プロセスを実行し終了を待つ(WinAPI Version)
'       MODULE_ID       : ShellWinAPI
'       CREATE_DATE     : 2011/02/26            K.ISHIZAKA
'       PARAM           : strCommand            コマンド(I)
'       RETURN          : プロセスの終了コード(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ShellWinAPI(ByVal strCommand As String) As Long
    Dim lngProcessID        As Long
    Dim lngHandle           As Long
    Dim lngExitCode         As Long
    On Error GoTo ErrorHandler
    
    lngProcessID = Shell(strCommand, vbHide)
    If lngProcessID = 0 Then
        Call MSZZ024_M10("Shell", "プロセスの実行に失敗しました。")
    End If
    lngHandle = OpenProcess(PROCESS_QUERY_INFORMATION, 1, lngProcessID)
    If lngHandle = 0 Then
        Call MSZZ024_M10("OpenProcess", "プロセスの取得に失敗しました。")
    End If
    On Error GoTo ErrorHandler1
    lngExitCode = STILL_ACTIVE
    While lngExitCode = STILL_ACTIVE
        If GetExitCodeProcess(lngHandle, lngExitCode) = 0 Then
            Call MSZZ024_M10("GetExitCodeProcess", "プロセスの取得に失敗しました。")
        End If
        DoEvents
    Wend
    Call CloseHandle(lngHandle)
    ShellWinAPI = lngExitCode
Exit Function

ErrorHandler1:
    Call CloseHandle(lngHandle)
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "ShellWinAPI" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended of program ********************************
