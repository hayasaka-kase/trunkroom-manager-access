Attribute VB_Name = "MSZZ021"
'****************************  strat or program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : メール送信
'       PROGRAM_ID      : MSZZ021
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2005/09/13
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          : 2005/11/04
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : マクロから呼びさせるようにFunctionに変更する
'
'       UPDATE          : 2005/11/10
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.2
'                       : １回だけリトライする
'
'       UPDATE          : 2005/11/24
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.3
'                       : 認証送信に変更する（認証方式 AUTH PLAIN）
'
'       UPDATE          : 2006/03/03
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.4
'                       : 任意のメッセージと任意の添付ファイルを追加可能にする
'
'       UPDATE          : 2006/03/13
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.5
'                       : システム用ＩＤに変更する
'
'       UPDATE          : 2006/08/14
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.6
'                       : ネットフォレストに対応する
'
'       UPDATE          : 2017/04/01
'       UPDATER         : K.SATO
'       Ver             : 0.7
'                       : ご紹介連絡表のメール送信部品を追加する
'
'       UPDATE          : 2018/09/15
'       UPDATER         : N.IMAI
'       Ver             : 0.8
'                       : ご紹介メールにBCCを追加
'
'**********************************************
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Const C_SMTP_SERVER = "post.kasegroup.co.jp"                           'DELETE 2007/08/14 K.ISHIZAKA
Private Const C_SMTP_SERVER = "v2200-054.mailsecure.jp"                                       'INSERT 2007/08/14 K.ISHIZAKA
'Private Const C_SEND_FROM = "システムから自動送信<kiyoshi_ishizaka@kasegroup.co.jp>"   'DELETE 2006/03/13 K.ISHIZAKA
Private Const C_SEND_FROM = "システムから自動送信<kase_system@kasegroup.co.jp>" 'INSERT 2006/03/13 K.ISHIZAKA
Private Const C_SEND_FROM_YYK = "yoyaku.annai@kasegroup.co.jp"                  'INSERT 2017/04/01 K.SATO
'Private Const C_USER = "kpcx000"                                               'DELETE 2006/03/13 K.ISHIZAKA 'INSERT 2005/11/24 K.ISHIZAKA
'Private Const C_USER = "kpcz044"                                               'DELETE 2007/08/14 K.ISHIZAKA 'INSERT 2006/03/13 K.ISHIZAKA
Private Const C_USER = "kase_system"                            'INSERT 2007/08/14 K.ISHIZAKA
'Private Const C_PASS = "kase"                                                  'DELETE 2007/08/14 K.ISHIZAKA 'INSERT 2005/11/24 K.ISHIZAKA
Private Const C_PASS = "kase1883"                                               'INSERT 2007/08/14 K.ISHIZAKA

'==============================================================================*
'       テスト用
'==============================================================================*
Sub TEST_SendMailAuthentication()
    Dim basp21              As Object
    Dim strErr              As String

    Set basp21 = CreateObject("basp21")
    If Not basp21 Is Nothing Then
'        strErr = SendMailAuthentication(basp21, "<zaka.ist.plus@hotmail.co.jp>", "配信テスト", "勝手に送るな!!", "")
        strErr = SendMailAuthentication2(basp21, C_SEND_FROM, C_USER, C_PASS, "<zaka.ist.plus@hotmail.co.jp>", "配信テスト", "勝手に送るな!!", "")
        Set basp21 = Nothing
    Else
        strErr = "『BASP21』がインストールされていません!" & vbCrLf & "メール配信できませんでした。"
    End If
    If strErr <> "" Then
        MsgBox strErr
    Else
        MsgBox "正常にメール配信しました。"
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : INIT_FILEの内容でメール送信する
'       MODULE_ID       : MSZZ021_M00
'       CREATE_DATE     : 2005/09/13
'       PARAM           : strProgb          プログラムＩＤ(I)
'                       : strAddSub         追加タイトル(I)
'                       : strAddMsg         追加メッセージ(I)
'                       : strAddFile        追加添付ファイル(I)
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub MSZZ021_M00(ByVal strProgb As String)                               'DELETE 2005/11/04 K.ISHIZAKA
'Public Function MSZZ021_M00(ByVal strProgb As String) As Boolean               'DELETE 2006/03/03 K.ISHIZAKA 'INSERT 2005/11/04 K.ISHIZAKA
Public Function MSZZ021_M00(ByVal strPROGB As String, Optional strAddSub As String = "", Optional strAddMsg As String = "", Optional strAddFile As String = "") As Boolean 'INSERT 2006/03/03 K.ISHIZAKA
    Dim basp21              As Object
    Dim strErr              As String
    Dim strSendTo           As String
    Dim strSubject          As String
    Dim strMessage          As String
    Dim strFiles            As String
    Dim strBCC              As String                                           'INSERT 2018/09/15 N.IMAI
    Dim rsInitFile          As Recordset
    On Error GoTo ErrorHandler                                                  'INSERT 2006/03/03 K.ISHIZAKA

    Set rsInitFile = CurrentDb().OpenRecordset(fncGetInitFile(strPROGB), dbOpenForwardOnly)
    With rsInitFile
        While Not .EOF
            Select Case UCase(fncGetRECFB(.Fields("INTIF_RECFB")))
            Case "SENDTO"
                strSendTo = strSendTo & vbTab & .Fields("INTIF_RECDB")
            Case "SUBJECT"
                strSubject = .Fields("INTIF_RECDB")
            Case "MESSAGE"
                strMessage = strMessage & .Fields("INTIF_RECDB") & vbCrLf
            Case "FILE"
                strFiles = strFiles & vbTab & .Fields("INTIF_RECDB")
            'INSERT 2018/09/15 N.IMAI Start
            Case "BCC"
                strBCC = Nz(strFiles & vbTab & .Fields("INTIF_RECDB"), "")
            'INSERT 2018/09/15 N.IMAI End
            End Select
            .MoveNext
        Wend
        .Close
    End With
'>>> INSERT START 2006/03/03 K.ISHIZAKA >>>
    If strAddSub <> "" Then
        strSubject = strSubject & strAddSub
    End If
    If strAddMsg <> "" Then
        strMessage = strMessage & strAddMsg & vbCrLf
    End If
    If strAddFile <> "" Then
        strFiles = strFiles & vbTab & strAddFile
    End If
'<<< INSERT END   2006/03/03 K.ISHIZAKA <<<
    Set basp21 = CreateObject("basp21")
    If Not basp21 Is Nothing Then                                               'INSERT 2005/11/04 K.ISHIZAKA
'        strErr = basp21.SendMail(C_SMTP_SERVER, Mid(strSendTo, 2), C_SEND_FROM, strSubject, strMessage, Mid(strFiles, 2))  'DELETE 2005/11/24 K.ISHIZAKA
'        strErr = SendMailAuthentication(basp21, Mid(strSendTo, 2), strSubject, strMessage, Mid(strFiles, 2))               'DELETE 2007/08/14 K.ISHIZAKA 'INSERT 2005/11/24 K.ISHIZAKA
'        strErr = SendMailAuthentication2(basp21, C_SEND_FROM, C_USER, C_PASS, Mid(strSendTo, 2), strSubject, strMessage, Mid(strFiles, 2))                   'DELETE 2018/09/15 N.IMAI 'INSERT 2007/08/14 K.ISHIZAKA
        strErr = SendMailAuthenticationBCC(basp21, C_SEND_FROM, C_USER, C_PASS, Mid(strSendTo, 2), strBCC, strSubject, strMessage, Mid(strFiles, 2))          'INSERT 2018/09/15 N.IMAI
        '１回だけリトライする
        If strErr <> "" Then                                                    'INSERT 2005/11/10 K.ISHIZAKA
'            strErr = basp21.SendMail(C_SMTP_SERVER, Mid(strSendTo, 2), C_SEND_FROM, strSubject, strMessage, Mid(strFiles, 2))  'DELETE 2005/11/24 K.ISHIZAKA 'INSERT 2005/11/10 K.ISHIZAKA
'            strErr = SendMailAuthentication(basp21, Mid(strSendTo, 2), strSubject, strMessage, Mid(strFiles, 2))               'DELETE 2007/08/14 K.ISHIZAKA 'INSERT 2005/11/24 K.ISHIZAKA
'            strErr = SendMailAuthentication2(basp21, C_SEND_FROM, C_USER, C_PASS, Mid(strSendTo, 2), strSubject, strMessage, Mid(strFiles, 2))  'INSERT 2007/08/14 K.ISHIZAKA
            strErr = SendMailAuthenticationBCC(basp21, C_SEND_FROM, C_USER, C_PASS, Mid(strSendTo, 2), strBCC, strSubject, strMessage, Mid(strFiles, 2))      'INSERT 2018/09/15 N.IMAI
        End If                                                                  'INSERT 2005/11/10 K.ISHIZAKA
        Set basp21 = Nothing
    Else                                                                        'INSERT 2005/11/04 K.ISHIZAKA
        strErr = "『BASP21』がインストールされていません!" & vbCrLf & "メール配信できませんでした。"    'INSERT 2005/11/04 K.ISHIZAKA
    End If                                                                      'INSERT 2005/11/04 K.ISHIZAKA
    If strErr <> "" Then
        MsgBox strErr
        MSZZ021_M00 = False                                                     'INSERT 2005/11/04 K.ISHIZAKA
    Else                                                                        'INSERT 2005/11/04 K.ISHIZAKA
        MSZZ021_M00 = True                                                      'INSERT 2005/11/04 K.ISHIZAKA
    End If
Exit Function                                                                   'INSERT 2006/03/03 K.ISHIZAKA
    
ErrorHandler:                   '↓自分の関数名                                 'INSERT 2006/03/03 K.ISHIZAKA
    Call Err.Raise(Err.Number, "MSZZ021_M00" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext) 'INSERT 2006/03/03 K.ISHIZAKA
End Function

'==============================================================================*
'
'       MODULE_NAME     : INIT_FILEの内容でメール送信する
'       MODULE_ID       : MSZZ021_M00_CUSTOM
'       CREATE_DATE     : 2017/03/26
'       PARAM           : strProgb          プログラムＩＤ(I)
'                       : strSendTo         宛先
'                       : strSendToCC       CC
'                       : strSubject        件名
'                       : strMessage        本文
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSZZ021_M00_CUSTOM(ByVal strPROGB As String, _
                                   Optional strSendTo As String = "", _
                                   Optional strSendToCC As String = "", _
                                   Optional strSubject As String = "", _
                                   Optional strMessage As String = "" _
                                   ) As Boolean
    Dim basp21              As Object
    Dim strErr              As String
    Dim strFiles            As String
    Dim strBCC              As String                                           'INSERT 2018/09/15 N.IMAI
    On Error GoTo ErrorHandler

    strBCC = Nz(DLookup("INTIF_RECDB", "INTI_FILE", " INTIF_PROGB = '" & strPROGB & "' AND INTIF_RECFB = 'BCC' "), "")  'INSERT 2018/09/15 N.IMAI
    Set basp21 = CreateObject("basp21")
    If Not basp21 Is Nothing Then
        'strErr = SendMailAuthentication2(basp21, C_SEND_FROM_YYK, C_USER, C_PASS, strSendTo, strSubject, strMessage, strFiles)                 'DELETE 2018/09/15 N.IMAI
        strErr = SendMailAuthenticationBCC(basp21, C_SEND_FROM_YYK, C_USER, C_PASS, strSendTo, strBCC, strSubject, strMessage, strFiles)        'INSERT 2018/09/15 N.IMAI
        '１回だけリトライする
        If strErr <> "" Then
            'strErr = SendMailAuthentication2(basp21, C_SEND_FROM_YYK, C_USER, C_PASS, strSendTo, strSubject, strMessage, strFiles)             'DELETE 2018/09/15 N.IMAI
            strErr = SendMailAuthenticationBCC(basp21, C_SEND_FROM_YYK, C_USER, C_PASS, strSendTo, strBCC, strSubject, strMessage, strFiles)    'INSERT 2018/09/15 N.IMAI
        End If
        Set basp21 = Nothing
    Else
        strErr = "『BASP21』がインストールされていません!" & vbCrLf & "メール配信できませんでした。"
    End If
    If strErr <> "" Then
        Call MSZZ003_M00(strPROGB, "8", " " & strErr)
        MSZZ021_M00_CUSTOM = False
    Else
        MSZZ021_M00_CUSTOM = True
    End If
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ021_M00_CUSTOM" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : INIT_FILE取得用ＳＱＬ作成
'       MODULE_ID       : fncGetInitFile
'       CREATE_DATE     : 2005/09/13
'       PARAM           : strProgb          プログラムＩＤ(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetInitFile(ByVal strPROGB As String) As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler                                                  'INSERT 2006/03/03 K.ISHIZAKA
    
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " INTIF_RECFB,"
    strSQL = strSQL & " INTIF_RECDB "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " INTI_FILE "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " INTIF_PROGB = '" & strPROGB & "' "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " INTIF_RECFB "
    fncGetInitFile = strSQL
Exit Function                                                                   'INSERT 2006/03/03 K.ISHIZAKA
    
ErrorHandler:                   '↓自分の関数名                                 'INSERT 2006/03/03 K.ISHIZAKA
    Call Err.Raise(Err.Number, "fncGetInitFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext) 'INSERT 2006/03/03 K.ISHIZAKA
End Function

'==============================================================================*
'
'       MODULE_NAME     : ピリオド以降切捨て
'       MODULE_ID       : fncGetRECFB
'       CREATE_DATE     : 2005/09/13
'       PARAM           : strProgb          プログラムＩＤ(I)
'       RETURN          : セッション名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetRECFB(ByVal strFields As String) As String
    Dim iPos                As Integer
    On Error GoTo ErrorHandler                                                  'INSERT 2006/03/03 K.ISHIZAKA
    
    iPos = InStr(strFields, ".")
    If iPos > 0 Then
        fncGetRECFB = Left(strFields, iPos - 1)
    Else
        fncGetRECFB = strFields
    End If
Exit Function                                                                   'INSERT 2006/03/03 K.ISHIZAKA
    
ErrorHandler:                   '↓自分の関数名                                 'INSERT 2006/03/03 K.ISHIZAKA
    Call Err.Raise(Err.Number, "fncGetRECFB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext) 'INSERT 2006/03/03 K.ISHIZAKA
End Function

'==============================================================================*
'
'       MODULE_NAME     : メール認証送信（認証方式 AUTH PLAIN）
'       MODULE_ID       : SendMailAuthentication
'       CREATE_DATE     : 2005/11/24
'       PARAM           : basp21            プログラムＩＤ(I)
'                       : strSendTo         送信先(I)
'                       : strSubject        件名(I)
'                       : strMessage        本文(I)
'                       : strFiles          添付ファイル(I)
'       RETURN          : エラーメッセージ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE START 2007/08/14 K.ISHIZAKA
'Private Function SendMailAuthentication(ByVal basp21 As Object, ByVal strSendTo As String, ByVal strSubject As String, ByVal strMessage As String, ByVal strFiles As String) As String
'    Dim outArray        As Variant
'    On Error GoTo ErrorHandler                                                  'INSERT 2006/03/03 K.ISHIZAKA
'
'    outArray = basp21.RcvMail(C_SMTP_SERVER, C_USER, C_PASS, "STAT", ".")
'    If IsArray(outArray) Then
'        SendMailAuthentication = basp21.SendMail(C_SMTP_SERVER, strSendTo, _
'            C_SEND_FROM & vbTab & C_USER & ":" & C_PASS & vbTab & "PLAIN", _
'            strSubject, strMessage, strFiles)
'    Else
'        SendMailAuthentication = outArray
'    End If
'Exit Function                                                                   'INSERT 2006/03/03 K.ISHIZAKA
'
'ErrorHandler:                   '↓自分の関数名                                 'INSERT 2006/03/03 K.ISHIZAKA
'    Call Err.Raise(Err.Number, "SendMailAuthentication" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext) 'INSERT 2006/03/03 K.ISHIZAKA
'End Function                                                                   'DELETE END   2007/08/14 K.ISHIZAKA

'==============================================================================*
'
'       MODULE_NAME     : メール認証送信（認証方式 AUTH PLAIN）
'       MODULE_ID       : SendMailAuthentication2
'       CREATE_DATE     : 2007/08/14
'       PARAM           : basp21            プログラムＩＤ(I)
'                       : strSendFrom       送信元(I)
'                       : strUser           ユーザー(I)
'                       : strPass           パスワード(I)
'                       : strSendTo         送信先(I)
'                       : strSubject        件名(I)
'                       : strMessage        本文(I)
'                       : strFiles          添付ファイル(I)
'       RETURN          : エラーメッセージ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SendMailAuthentication2(ByVal basp21 As Object, ByVal strSendFrom As String, ByVal strUser As String, ByVal strPass As String, ByVal strSendTo As String, ByVal strSubject As String, ByVal strMessage As String, ByVal strFiles As String) As String
    Dim outArray            As Variant
    On Error GoTo ErrorHandler
    
    outArray = basp21.RcvMail(C_SMTP_SERVER & vbTab & "110", strUser, strPass, "STAT", ".")
    If IsArray(outArray) Then
        SendMailAuthentication2 = basp21.sendMail(C_SMTP_SERVER & ":25", strSendTo, _
            strSendFrom & vbTab & strUser & ":" & strPass & vbTab & "PLAIN", _
            strSubject, strMessage, strFiles)
    Else
        SendMailAuthentication2 = outArray
    End If
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "SendMailAuthentication2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext) 'INSERT 2006/03/03 K.ISHIZAKA
End Function

'==============================================================================*
'
'       MODULE_NAME     : メール認証送信（認証方式 AUTH PLAIN）
'       MODULE_ID       : SendMailAuthenticationBCC
'       CREATE_DATE     : 2018/09/15    N.IMAI
'       PARAM           : basp21            プログラムＩＤ(I)
'                       : strSendFrom       送信元(I)
'                       : strUser           ユーザー(I)
'                       : strPass           パスワード(I)
'                       : strSendTo         送信先(I)
'                       : strSendToBCC      送信先BCC(I)
'                       : strSubject        件名(I)
'                       : strMessage        本文(I)
'                       : strFiles          添付ファイル(I)
'       RETURN          : エラーメッセージ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SendMailAuthenticationBCC(ByVal basp21 As Object, ByVal strSendFrom As String, ByVal strUser As String, ByVal strPass As String, ByVal strSendTo As String, ByVal strSendToBCC As String, ByVal strSubject As String, ByVal strMessage As String, ByVal strFiles As String) As String
    Dim outArray            As Variant
    Dim strSend             As String
    On Error GoTo ErrorHandler
    
    If Nz(strSendToBCC, "") = "" Then
        SendMailAuthenticationBCC = SendMailAuthentication2(basp21, strSendFrom, strUser, strPass, strSendTo, strSubject, strMessage, strFiles)
    Else
        strSend = strSendTo & vbTab & "bcc" & vbTab & strSendToBCC
        outArray = basp21.RcvMail(C_SMTP_SERVER & vbTab & "110", strUser, strPass, "STAT", ".")
        If IsArray(outArray) Then
            SendMailAuthenticationBCC = _
                basp21.sendMail(C_SMTP_SERVER & ":25", _
                                strSend, _
                                strSendFrom & vbTab _
                                & strUser & ":" & strPass & vbTab _
                                & "PLAIN", _
                                strSubject, _
                                strMessage, _
                                strFiles)
        Else
            SendMailAuthenticationBCC = outArray
        End If
    End If
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "SendMailAuthenticationBCC" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 全て指定してメール送信する
'       MODULE_ID       : MSZZ021_M10
'       CREATE_DATE     : 2007/08/14
'       PARAM           : strSendFrom       送信元(I)
'                       : strSendTo         送信先(I)
'                       : strSubject        タイトル(I)
'                       : strMessage        メッセージ(I)
'                       : strFiles          添付ファイル(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub MSZZ021_M10(ByVal strSendFrom As String, ByVal strSendTo As String, ByVal strSubject As String, ByVal strMessage As String, Optional strFiles As String = "")
    Dim basp21              As Object
    Dim strErr              As String
    Dim strUser             As String
    On Error GoTo ErrorHandler

    strUser = Split(strSendFrom, "<")(1)
    strUser = Split(strUser, ">")(0)
    Set basp21 = CreateObject("basp21")
    If Not basp21 Is Nothing Then
        strErr = SendMailAuthentication2(basp21, strSendFrom, strUser, C_PASS, strSendTo, strSubject, strMessage, strFiles)
        '１回だけリトライする
        If strErr <> "" Then
            strErr = SendMailAuthentication2(basp21, strSendFrom, strUser, C_PASS, strSendTo, strSubject, strMessage, strFiles)
        End If
        Set basp21 = Nothing
    Else
        strErr = "『BASP21』がインストールされていません!" & vbCrLf & "メール配信できませんでした。"
    End If
    If strErr <> "" Then
        Call MSZZ024_M10("SendMailAuthentication2", strErr)
    End If
Exit Sub
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "MSZZ021_M10" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended or program ********************************

