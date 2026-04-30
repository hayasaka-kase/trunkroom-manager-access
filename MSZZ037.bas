Attribute VB_Name = "MSZZ037"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : FTPファイル転送関数
'        PROGRAM_ID      : MSZZ037
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/08/14
'        CERATER         : K.ISHZIAKA
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const C_FTP_PROVIDER = "basp21.FTP"

'==============================================================================*
'
'       MODULE_NAME     : ソケット作成
'       MODULE_ID       : FtpCreate
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : [bPASV]               パッシブモード(I)省略時(False)
'       RETURN          : FTPオブジェクト(Object)
'
'       SETU_TABLに設定しておけば接続まで行う
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function FtpCreate(Optional bPAVS As Boolean = False) As Object
    Dim objFtp              As Object
    Dim strSvr              As String
    Dim strUid              As String
    Dim strPwd              As String
    On Error GoTo ErrorHandler
    
    Set objFtp = CreateObject(C_FTP_PROVIDER)
    If objFtp Is Nothing Then
        Call MSZZ024_M10("FTP.PROVIDER", "システムの設定不足です。")
    End If
    
    strSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'FTP_SERVER_NAME'"))
    strUid = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'FTP_USER_ID'"))
    strPwd = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'FTP_PASSWORD'"))
    If (strSvr <> "") And (strUid <> "") And (strPwd <> "") Then
        FtpConnect objFtp, strSvr, strUid, strPwd, bPAVS
    End If
    
    Set FtpCreate = objFtp
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "FtpCreate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       SETU_TABLに設定しておけばFtpCreate時に自動で接続をしてくれるので呼び出す必要はない
'
'       MODULE_NAME     : 接続
'       MODULE_ID       : FtpConnect
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objFtp                FTPオブジェクト(I)
'                       : strSvr                サーバー(I)
'                       : strUid                ユーザー(I)
'                       : strPwd                パスワード(I)
'       PARAM           : [bPASV]               パッシブモード(I)省略時(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub FtpConnect(objFtp As Object, ByVal strSvr As String, ByVal strUid As String, ByVal strPwd As String, Optional bPAVS As Boolean = False)
    Dim lngResult           As Long
    Dim strErrMsg           As String
    On Error GoTo ErrorHandler
    
    lngResult = objFtp.Connect(strSvr, strUid, strPwd)
    If lngResult <> 0 Then
        Select Case lngResult
        Case -1
            strErrMsg = "Can't open sockt!"
        Case -2
            strErrMsg = "Timeout!"
        Case 1 To 5
            strErrMsg = objFtp.GetReply()
        Case Is > 10000
            strErrMsg = "Winsock Error!"
        Case Else
            strErrMsg = "unkown Error!"
        End Select
        Call Err.Raise(lngResult, "FTP.Connect", strErrMsg)
    End If
    If bPAVS Then
        lngResult = objFtp.Command("pasv")
        If lngResult <> 2 Then
            strErrMsg = objFtp.GetReply()
            objFtp.Close
            Call Err.Raise(-1, "FTP.pasv", strErrMsg)
        End If
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "FtpConnect" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 切断
'       MODULE_ID       : FtpClose
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objFtp                FTPオブジェクト(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub FtpClose(objFtp As Object)
    objFtp.Close
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ファイル送信
'       MODULE_ID       : FtpPut
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objFtp                FTPオブジェクト(I)
'                       : strLocal              送信元フルパスファイル名(I)
'                       : strText               送信先フォルダー名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub FtpPut(objFtp As Object, ByVal strLocal As String, ByVal strRemote As String)
    Dim strErrMsg           As String
    On Error GoTo ErrorHandler
    
    If objFtp.PutFile(strLocal, strRemote, 1) <> 1 Then
        strErrMsg = objFtp.GetReply()
        objFtp.Close
        Call Err.Raise(-1, "FTP.Put", strErrMsg)
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "FtpPut" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ファイル取得
'       MODULE_ID       : FtpGet
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objFtp                FTPオブジェクト(I)
'                       : strRemote             取得元フルパスファイル名(I)
'                       : strLocal              取得先フォルダー名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub FtpGet(objFtp As Object, ByVal strRemote As String, ByVal strLocal As String)
    Dim strErrMsg           As String
    On Error GoTo ErrorHandler
    
    If objFtp.GetFile(strRemote, strLocal, 1) < 0 Then
        strErrMsg = objFtp.GetReply()
        objFtp.Close
        Call Err.Raise(-1, "FTP.Get", strErrMsg)
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "FtpGet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ファイル削除
'       MODULE_ID       : FtpDelete
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objFtp                FTPオブジェクト(I)
'                       : strRemote             フルパスファイル名(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub FtpDelete(objFtp As Object, ByVal strRemote As String)
    Dim strErrMsg           As String
    On Error GoTo ErrorHandler
    
    If objFtp.DeleteFile(strRemote) <= 0 Then
        strErrMsg = objFtp.GetReply()
        objFtp.Close
        Call Err.Raise(-1, "FTP.Delete", strErrMsg)
    End If
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "FtpDelete" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ファイル移動（取得後に削除する）
'       MODULE_ID       : FtpMove
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objFtp                FTPオブジェクト(I)
'                       : strRemote             取得元フォルダー名(I)
'                       : strLocal              取得先フォルダー名(I)
'                       : strFiles()            ファイル名(O)
'       RETURN          : 件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function FtpMove(objFtp As Object, ByVal strRemote As String, ByVal strLocal As String, ByRef strFiles() As String) As Long
    Dim varResult           As Variant
    Dim strErrMsg           As String
    Dim strPath             As String
    Dim i                   As Long
    On Error GoTo ErrorHandler
    
    varResult = Split(strRemote, "/")
    varResult(UBound(varResult)) = ""
    strPath = Join(varResult, "/")
    varResult = objFtp.GetDir(strRemote, 0)
    If IsArray(varResult) Then
        ReDim strFiles(0 To UBound(varResult) - LBound(varResult))
        For i = 0 To UBound(varResult) - LBound(varResult)
            strFiles(i) = varResult(i + LBound(varResult))
            If Left(strFiles(i), Len(strPath)) = strPath Then
                strFiles(i) = Mid(strFiles(i), Len(strPath) + 1)
            End If
            FtpGet objFtp, strPath & strFiles(i), strLocal
            FtpDelete objFtp, strPath & strFiles(i)
        Next
    Else
        i = 0
    End If
    FtpMove = i
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "FtpMove" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'****************************  ended or program ********************************
