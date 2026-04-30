Attribute VB_Name = "MSZZ024"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : エラー情報出力
'        PROGRAM_ID      : MSZZ024
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2006/01/13
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          : 2006/03/03
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.1
'                        : 作成したファイル名を戻り値にする
'
'        UPDATE          : 2006/06/19
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.2
'                        : ファイルが指定パスに作成できない場合、自下に作成する
'                        : ファイルに書き込めなかったときは戻り値が空になる
'
'        UPDATE          : 2006/06/20
'        UPDATER         : K.ISHZIAKA
'        Ver             : 0.3
'                        : エラー判定の前にクリアするのを忘れていた
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Const vbRightAllow   As String = " → "

'==============================================================================*
'
'       MODULE_NAME     : 例外処理に対するエラー情報をサーバーにファイル出力する
'       MODULE_ID       : MSZZ024_M00
'       CREATE_DATE     : 2006/01/13
'       PARAM           : strFunctionName       関数名(I)
'                       : blMsgBox              メッセージ表示(True)/非表示(False)(I)
'       RETURN          : 作成したファイル名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub MSZZ024_M00(ByVal strFunctionName As String, ByVal blMsgBox As Boolean)                 'DELETE 2006/03/03 K.ISHIZAKA
Public Function MSZZ024_M00(ByVal strFunctionName As String, ByVal blMsgBox As Boolean) As String   'INSERT 2006/03/03 K.ISHIZAKA
    Dim Locator             As Object
    Dim SERVICE             As Object
    Dim Wmi                 As Object
    Dim strMessage          As String
    Dim strPath             As String
    Dim strFile             As String
    Dim iFlno               As Integer
    
    strMessage = strMessage & vbCrLf & "エラー情報" & vbCrLf
    strMessage = strMessage & "Err.Number=" & Err.Number & vbCrLf
    strMessage = strMessage & "Err.Source=" & strFunctionName & vbRightAllow & Err.Source & vbCrLf
    strMessage = strMessage & "Err.Description=" & Err.Description & vbCrLf
    If Application.CurrentObjectType = acForm Then
        strMessage = strMessage & "Err.CurrentObjectName=" & Application.CurrentObjectName & vbCrLf
    End If

    On Error Resume Next
    strPath = DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'MSZZ024'")
    If strPath = "" Then
        strPath = Application.CurrentDb.NAME
        strPath = Left(strPath, Len(strPath) - Len(Dir(strPath)))
        strPath = strPath & "ErrorLog\"
        If Dir(strPath, vbDirectory) = "" Then
            Call MkDir(strPath)
        End If
    End If
    strFile = Format(Now, "yyyymmddhhnnss")
    
    strMessage = strMessage & vbCrLf & "起動中情報" & vbCrLf
    For Each Wmi In Application.Forms
        strMessage = strMessage & "Form.Name=" & Wmi.NAME & vbCrLf
    Next
    For Each Wmi In Application.Reports
        strMessage = strMessage & "Reports.Name=" & Wmi.NAME & vbCrLf
    Next

    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Set SERVICE = Locator.ConnectServer

    strMessage = strMessage & vbCrLf & "マシン情報" & vbCrLf
    For Each Wmi In SERVICE.ExecQuery("Select * From Win32_OperatingSystem")
        strFile = strFile & Wmi.CSName
        strMessage = strMessage & "Computer.Name=" & Wmi.CSName & vbCrLf
        strMessage = strMessage & "Computer.Description=" & Wmi.Description & vbCrLf
    Next

    strMessage = strMessage & vbCrLf & "ログイン情報" & vbCrLf
    For Each Wmi In SERVICE.ExecQuery("Select * From Win32_Computersystem")
        strMessage = strMessage & "Login.UserName=" & Wmi.userName & vbCrLf
        strMessage = strMessage & "Login.Domain=" & Wmi.Domain & vbCrLf
    Next

    strMessage = strMessage & vbCrLf & "ネットワーク情報" & vbCrLf
    For Each Wmi In SERVICE.ExecQuery("Select * From Win32_NetworkAdapterConfiguration")
        If Wmi.IPEnabled Then
            strMessage = strMessage & "Nic.IPAddress=" & Wmi.IPAddress(0) & vbCrLf
            strMessage = strMessage & "Nic.IPSubnet=" & Wmi.IPSubnet(0) & vbCrLf
            strMessage = strMessage & "Nic.MACAddress=" & Wmi.MACAddress & vbCrLf
        End If
    Next

    Set Wmi = Nothing
    Set Locator = Nothing
    Set SERVICE = Nothing

    iFlno = FreeFile()
    Err.Clear                                                                   'INSERT 2006/06/20 K.ISHIZAKA
    Open strPath & strFile & ".log" For Append As #iFlno
'>>> INSERT START 2006/06/19 K.ISHIZAKA >>>
    If Err.Number <> 0 Then
        Err.Clear
        strPath = Application.CurrentDb.NAME
        strPath = Left(strPath, Len(strPath) - Len(Dir(strPath)))
        strPath = strPath & "ErrorLog\"
        If Dir(strPath, vbDirectory) = "" Then
            Call MkDir(strPath)
        End If
        Open strPath & strFile & ".log" For Append As #iFlno
    End If
    If Err.Number <> 0 Then
        If blMsgBox Then
            MsgBox "システムエラーが発生しました。" & vbCrLf & _
                    strMessage
        End If
        MSZZ024_M00 = ""
        Exit Function
    End If
'<<< INSERT END   2006/06/19 K.ISHIZAKA <<<
    Print #iFlno, strMessage
    Close #iFlno

    If blMsgBox Then
        MsgBox "システムエラーが発生しました。" & vbCrLf & _
               "次の番号を控えてシステム管理者に連絡してください。" & vbCrLf & _
               "[" & strFile & "]"
    End If
    
    MSZZ024_M00 = strPath & strFile & ".log"                                    'INSERT 2006/03/03 K.ISHIZAKA
End Function

'==============================================================================*
'
'       MODULE_NAME     : 任意にエラー終了させる
'       MODULE_ID       : MSZZ024_M10
'       CREATE_DATE     : 2006/01/13
'       PARAM           : strCommand            直前の命令文(I)
'                       : strErrMessage         エラーメッセージ(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ024_M10(ByVal strCommand As String, ByVal strErrMessage As String)
    Call Err.Raise(65535, strCommand, strErrMessage)
End Sub

'==============================================================================*
'       以下　コーディング例
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Command1_Click() As Boolean
    On Error GoTo ErrorHandler
    
    Call aaa
Exit Function
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Command1_Click", True)   '←親となる関数に対してだけ呼び出しを記述
End Function

'==============================================================================*
Private Sub aaa()
    On Error GoTo ErrorHandler
    
    Call bbb
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "aaa" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
Private Sub bbb()
    On Error GoTo ErrorHandler
    
    Call ccc
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "bbb" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
Private Sub ccc()
    On Error GoTo ErrorHandler
    Dim xx As Boolean
    
    If xx = False Then
            '↓自分でエラー終了したいとき
        Call MSZZ024_M10("OpenRecordset", "Timeoutしました。")
    End If
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ccc" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended or program ********************************
