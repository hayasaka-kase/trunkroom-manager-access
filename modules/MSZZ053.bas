Attribute VB_Name = "MSZZ053"
'****************************  start of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : 加瀬システム
'       SUB_SYSTEM_NAME : 共通関数
'
'       PROGRAM_NAME    : EXCEL出力
'       PROGRAM_ID      : MSZZ053
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2010/06/01
'       CERATER         : S.SHIBAZAKI
'       Ver             : 1.0
'
'       UPDATE          : 2010/06/12
'       UPDATER         : K.ISHIZAKA
'       Ver             : 1.1
'                       : エクセルが開かれているときの処理追加
'                       : テーブル名の括弧をはずす  [WORK-TABLE]みたいになってるやつ
'
'       UPDATE          : 2010/11/22
'       UPDATER         : K.ISHIZAKA
'       Ver             : 1.2
'                       : 一時テーブル対応（"#"つきのTEMPテーブル）
'
'       UPDATE          : 2011/05/21
'       UPDATER         : K.ISHIZAKA
'       Ver             : 1.3
'                       : 引数追加「出力先パス」、「エクセル起動」
'
'       UPDATE          : 2011/10/31
'       UPDATER         : K.ISHIZAKA
'       Ver             : 1.4
'                       : エラー発生時にゴミが残らないように後処理をきちんとする
'
'       UPDATE          : 2017/12/25
'       UPDATER         : K.ISHIZAKA
'       Ver             : 1.5
'                       : 出力形式をxlsxに変更する
'
'       UPDATE          : 2019/03/16
'       UPDATER         : N.IMAI
'       Ver             : 1.6
'                       : xlsxでエラー時はxlsにする
'
'       UPDATE          : 2019/11/01
'       UPDATER         : N.IMAI
'       Ver             : 1.7
'                       : xlsxで出力する条件にEXCELのバージョンを見る処理を追加
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const DEFAULT_QUERY         As String = "MQZZ053"

'==============================================================================*
'   テスト
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub TEST_subLinkTempTableToExcel()
    Dim objCon              As Object
    Dim strTempTable        As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    strTempTable = "#WK_KAIK_TRAN"
    
    Set objCon = ADODB_Connection()
    On Error GoTo ErrorHandler1
    
    strSQL = "SELECT TOP 10 * INTO " & strTempTable & " FROM KAIK_TRAN WHERE KAIKT_SHYMD = '201011'"
    Call ADODB_Execute(strSQL, objCon)
    On Error GoTo ErrorHandler2
    
    strSQL = ""
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " " & Replace(strTempTable, "#", "") & ".KAIKT_SHYMD, "
    strSQL = strSQL & " " & strTempTable & ".KAIKT_KEIYB,"
'    strSQL = strSQL & " wk.KAIKT_BUNKB,"
    strSQL = strSQL & " KAIKT_BUMOC "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strTempTable ' & " wk "
    Call subLinkTempTableToExcel(objCon, strTempTable, strSQL)
    
ErrorHandler2:
    If Left(strTempTable, 1) <> "#" Then
        Call ADODB_DropTable(strTempTable, objCon)
    End If
    
ErrorHandler1:
    objCon.Close
    
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : テーブルリンク→クエリ作成→EXCEL出力の一連の処理
'                         テーブルがTempでも、そうでなくても可能とする
'       MODULE_ID       : subLinkTempTableToExcel
'       CREATE_DATE     : 2010/11/22            K.ISHIZAKA
'       PARAM           : objCon                コネクションオブジェクト(I)
'       PARAM           : strTableId            リンクするテーブルID(I)
'                       : strQuerySql           クエリSQL(I)
'                       : [strBumoc]            接続部門コード(I)省略可
'                       :                       KASE_DBに接続する場合は省略する
'                       : [strQueryId]          クエリID(I)省略可
'                       :                       省略時、呼出し元画面IDの二文字目を"Q"に変換
'                       :                       呼出し元画面IDが取得できない(画面から呼ばれてない)場合は"MQZZ053"
'                       : [strPath]             出力先パス(I)省略可
'                       : [blAutoStart]         エクセル起動(I)：起動する(True)規定値／起動しない(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub subLinkTempTableToExcel(objCon As Object, ByVal strTempTable As String, ByVal strQuerySql As String, Optional ByVal strBUMOC As String = "", Optional ByVal strQueryId As String = "") 'DELETE 2011/05/21 K.ISHIZAKA
Public Sub subLinkTempTableToExcel(objCon As Object, ByVal strTempTable As String, ByVal strQuerySql As String, _
    Optional ByVal strBUMOC As String = "", Optional ByVal strQueryId As String = "", _
    Optional ByVal strPath As String = "", Optional ByVal blAutoStart As Boolean = True) 'INSERT 2011/05/21 K.ISHIZAKA
    Dim strTableId          As String
    Dim strWorkTableId      As String
    On Error GoTo ErrorHandler

    strTableId = MSZZ025.GetSourceTableName(strTempTable)
    '一時テーブルのとき
    If Left(strTableId, 1) = "#" Then
        strWorkTableId = "[TEMP_" & LsGetComputerName() & "_" & Format(Now, "yyyymmddhhnnss") & "_" & Mid(strTableId, 2) & "]"
        If strTableId <> strTempTable Then
            strQuerySql = Replace(strQuerySql, strTempTable, strWorkTableId)
        Else
            strQuerySql = Replace(strQuerySql, strTableId & " ", strWorkTableId & " ")
            strQuerySql = Replace(strQuerySql, strTableId & ".", strWorkTableId & ".")
            If Right(strQuerySql, Len(strTableId)) = strTableId Then
                strQuerySql = Replace(strQuerySql, strTableId, strWorkTableId)
            Else
                strQuerySql = Replace(strQuerySql, strTableId & ";", strWorkTableId & ";")
            End If
            strTableId = Mid(strTableId, 2)
            strQuerySql = Replace(strQuerySql, strTableId & " ", strWorkTableId & " ")
            strQuerySql = Replace(strQuerySql, strTableId & ".", strWorkTableId & ".")
            If Right(strQuerySql, Len(strTableId)) = strTableId Then
                strQuerySql = Replace(strQuerySql, strTableId, strWorkTableId)
            Else
                strQuerySql = Replace(strQuerySql, strTableId & ";", strWorkTableId & ";")
            End If
        End If
        '実テーブルを仮作成する
        Call ADODB_Execute("SELECT * INTO " & strWorkTableId & " FROM " & strTempTable, objCon)
        On Error GoTo ErrorHandler1
'        Call MSZZ053.subLinkTableToExcel(strWorkTableId, strQuerySql, strBUMOC, strQueryId) 'DELETE 2011/05/21 K.ISHIZAKA
        Call MSZZ053.subLinkTableToExcel(strWorkTableId, strQuerySql, strBUMOC, strQueryId, strPath, blAutoStart) 'INSERT 2011/05/21 K.ISHIZAKA
        '実テーブルを削除する
        Call ADODB_DropTable(strWorkTableId, objCon)
        On Error GoTo ErrorHandler
    Else
'        Call MSZZ053.subLinkTableToExcel(strTempTable, strQuerySql, strBUMOC, strQueryId) 'DELETE 2011/05/21 K.ISHIZAKA
        Call MSZZ053.subLinkTableToExcel(strTempTable, strQuerySql, strBUMOC, strQueryId, strPath, blAutoStart) 'INSERT 2011/05/21 K.ISHIZAKA
    End If
Exit Sub

ErrorHandler1:
    Call ADODB_DropTable(strTableId, objCon)
ErrorHandler:
    Call Err.Raise(Err.Number, "subLinkTempTableToExcel" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : テーブルリンク→クエリ作成→EXCEL出力の一連の処理
'       MODULE_ID       : subLinkTableExcelOutput
'       CREATE_DATE     : 2010/06/01
'       PARAM           : strTableId            リンクするテーブルID
'                       : strQuerySql           クエリSQL
'                       : strBumoc              接続部門コード　省略可
'                       :                       KASE_DBに接続する場合は省略する
'                       : strQueryId            クエリID　省略可
'                       :                       省略時、呼出し元画面IDの二文字目を"Q"に変換
'                       :                       呼出し元画面IDが取得できない(画面から呼ばれてない)場合は"MQZZ053"
'                       : [strPath]             出力先パス(I)省略可
'                       : [blAutoStart]         エクセル起動(I)：起動する(True)規定値／起動しない(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub subLinkTableToExcel(strTableId As String, strQuerySql As String, Optional strBumoc As String = "", Optional strQueryId As String = "") 'DELETE 2010/06/12 K.ISHIZAKA
'Public Sub subLinkTableToExcel(ByVal strTableId As String, ByVal strQuerySql As String, Optional ByVal strBUMOC As String = "", Optional ByVal strQueryId As String = "") 'DELETE 2011/05/21 K.ISHIZAKA
Public Sub subLinkTableToExcel(ByVal strTableId As String, ByVal strQuerySql As String, _
    Optional ByVal strBUMOC As String = "", Optional ByVal strQueryId As String = "", _
    Optional ByVal strPath As String = "", Optional ByVal blAutoStart As Boolean = True) 'INSERT 2011/05/21 K.ISHIZAKA
    
'    Dim strDNS              As String                                          'DELETE START 2011/10/31 K.ISHIZAKA
'    Dim strSvr              As String
'    Dim strDBN              As String
'    Dim strUid              As String
'    Dim strPwd              As String
'    Dim strErrMessage       As String
'    Dim strBumocProc        As String                                          'DELETE END   2011/10/31 K.ISHIZAKA
    Dim strQueryIdProc      As String
    
    On Error GoTo ErrorHandler
    
'    strTableId = MSZZ025.GetSourceTableName(strTableId)                        'DELETE 2011/10/31 K.ISHIZAKA 'INSERT 2010/06/12 K.ISHIZAKA
    
    strQueryIdProc = strQueryId
    If strQueryIdProc = "" Then
        strQueryIdProc = fncGetQueryId()
    End If
    
'    strBumocProc = IIf(strBumoc <> "", "_" & strBumoc, "")                     'DELETE START 2011/10/31 K.ISHIZAKA
'
'    strDNS = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATA_SOURCE_NAME" & strBumocProc & "'")
'    strSvr = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_SERVER_NAME" & strBumocProc & "'")
'    strDBN = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME" & strBumocProc & "'")
'    strUid = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_USER_ID" & strBumocProc & "'")
'    strPwd = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_PASSWORD" & strBumocProc & "'")
'
'    'リンクをはずす
'    Call MSZZ002_M00(strTableId, strErrMessage)
'    strErrMessage = ""                                                         'DELETE END   2011/10/31 K.ISHIZAKA
    On Error Resume Next
    'クエリー削除
    Call CurrentDb().QueryDefs.Delete(strQueryIdProc)
    On Error GoTo ErrorHandler
    
    'リンクする
'    If MSZZ002_M20(strDNS, strSvr, strDBN, strUid, strPwd, "dbo." & strTableId, strTableId, strErrMessage) <> 0 Then 'DELETE 2011/10/31 K.ISHIZAKA
'        Call MSZZ024_M10("MSZZ002_M20", strErrMessage)                         'DELETE 2011/10/31 K.ISHIZAKA
'    End If                                                                     'DELETE 2011/10/31 K.ISHIZAKA
    Call MSZZ002_LinkOn(strTableId, strBUMOC)                                   'INSERT 2011/10/31 K.ISHIZAKA
    On Error GoTo ErrorHandler1                                                 'INSERT 2011/10/31 K.ISHIZAKA
    
    'クエリー作成
    Call CurrentDb().CreateQueryDef(strQueryIdProc, strQuerySql)
    On Error GoTo ErrorHandler2                                                 'INSERT 2011/10/31 K.ISHIZAKA
    
    'エクセル出力
'    Call subOutputToExcel(strQueryIdProc)                                      'DELETE 2011/05/21 K.ISHIZAKA
    Call subOutputToExcel(strQueryIdProc, strPath, blAutoStart)                 'INSERT 2011/05/21 K.ISHIZAKA
    
    'クエリー削除
    Call CurrentDb().QueryDefs.Delete(strQueryIdProc)
    On Error GoTo ErrorHandler1                                                 'INSERT 2011/10/31 K.ISHIZAKA
    
    'リンクをはずす
'    If MSZZ002_M00(strTableId, strErrMessage) <> 0 Then                        'DELETE 2011/10/31 K.ISHIZAKA
'        Call MSZZ024_M10("MSZZ002_M00", strErrMessage)                         'DELETE 2011/10/31 K.ISHIZAKA
'    End If                                                                     'DELETE 2011/10/31 K.ISHIZAKA
    Call MSZZ002_LinkOff(strTableId)                                            'INSERT 2011/10/31 K.ISHIZAKA
    On Error GoTo ErrorHandler                                                  'INSERT 2011/10/31 K.ISHIZAKA
Exit Sub
    
ErrorHandler2:                                                                  'INSERT 2011/10/31 K.ISHIZAKA
    Call CurrentDb().QueryDefs.Delete(strQueryIdProc)                           'INSERT 2011/10/31 K.ISHIZAKA
ErrorHandler1:                                                                  'INSERT 2011/10/31 K.ISHIZAKA
    Call MSZZ002_LinkOff(strTableId)                                            'INSERT 2011/10/31 K.ISHIZAKA
ErrorHandler:
    Call Err.Raise(Err.Number, "subLinkTableToExcel" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : EXCEL出力
'       MODULE_ID       : subOutputToExcel
'       CREATE_DATE     : 2010/06/01
'       PARAM           : strQueryId            クエリID
'                       : [strPath]             出力先パス(I)省略可
'                       : [blAutoStart]         エクセル起動(I)：起動する(True)規定値／起動しない(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub subOutputToExcel(strQueryId As String)                              'DELETE 2011/05/21 K.ISHIZAKA
Public Sub subOutputToExcel(ByVal strQueryId As String, _
    Optional ByVal strPath As String = "", Optional ByVal blAutoStart As Boolean = True) 'INSERT 2011/05/21 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    If (Not blAutoStart) And (strPath = "") Then                                'INSERT START 2011/05/21 K.ISHIZAKA
        'Excel起動しないときには、strPath の指定必須
        Call MSZZ024_M10("OutputTo", "Excel起動しないときには、strPath を指定してください。")
    End If
    If strPath <> "" Then
        'If LCase(Right(strPath, 4)) <> ".xls" Then 'パスにファイル名まで指定されているか？         'DELETE 2019/03/16 N.IMAI
        If InStr(1, LCase(strPath), ".xls") <= 0 Then    'パスにファイル名まで指定されているか？    'INSERT 2019/03/16 N.IMAI
            If Right(strPath, 1) <> "\" Then
                strPath = strPath & "\" & strQueryId & ".xls"
            Else
                strPath = strPath & strQueryId & ".xls"
            End If
        End If
    End If                                                                      'INSERT END   2011/05/21 K.ISHIZAKA
    'エクセル出力
'    doCmd.OutputTo acOutputQuery, strQueryId, "MicrosoftExcel(*.xls)", "", True, "" 'DELETE 2011/05/21 K.ISHIZAKA
'    doCmd.OutputTo acOutputQuery, strQueryId, "MicrosoftExcel(*.xls)", strPath, blAutoStart, "" 'DELETE 2017/12/25 K.ISHIZAKA 'INSERT 2011/05/21 K.ISHIZAKA
    On Error Resume Next
    Dim xlApp   As Object                                                       'INSERT 2019/11/01 N.IMAI
    Set xlApp = CreateObject("Excel.Application")                               'INSERT 2019/11/01 N.IMAI
    'INSERT 2019/03/16 N.IMAI Start
    'If InStr(1, LCase(strPath), ".xlsx") > 0 Then                              'DELETE 2019/11/01 N.IMAI
    If xlApp.Version > 14 Or InStr(1, LCase(strPath), ".xlsx") > 0 Then  'INSERT 2019/11/01 N.IMAI
        doCmd.OutputTo acOutputQuery, strQueryId, acFormatXLSX, strPath, blAutoStart, ""    'INSERT 2017/12/25 K.ISHIZAKA
    Else
        doCmd.OutputTo acOutputQuery, strQueryId, acFormatXLS, strPath, blAutoStart, ""
    End If
    'INSERT 2019/03/16 N.IMAI End
    If Err <> 0 And Err <> 2501 Then
        Err.Clear
        On Error GoTo ErrorHandler
        doCmd.OutputTo acOutputQuery, strQueryId, acFormatXLS, strPath, blAutoStart, ""
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not xlApp Is Nothing Then Set xlApp = Nothing                            'INSERT 2019/11/01 N.IMAI
    If Err.Number = 2501 Then
        'キャンセルボタンはエラーにしない
        Resume Next
    ElseIf Err.Number = 2302 Then                                               'INSERT START 2010/06/12 K.ISHIZAKA
        MsgBox "ファイルを閉じてから保存してください。", vbInformation, Screen.ActiveForm.Caption
        Resume                                                                  'INSERT END   2010/06/12 K.ISHIZAKA
    End If
    Call Err.Raise(Err.Number, "subOutputToExcel" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :クエリID取得
'        MODULE_ID        :fncGetQueryId
'        Return           :クエリID
'        CREATE_DATE      :2010/06/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetQueryId() As String
    On Error Resume Next

    Dim objCallFrom     As Object
    Dim objParent       As Object
    Dim strQueryId      As String
    
    '呼び出し元Object取得
    Set objCallFrom = Application.CodeContextObject
    
    If objCallFrom Is Nothing Then
        '呼び出し元が、フォームやレポートではない場合、（標準モジュールから呼び出されたりとか）
        'CodeContextObjectプロパティを参照できないので、この標準モジュールのIDを返却する。
        strQueryId = DEFAULT_QUERY
    Else
        '名前を返却
        '呼び出し元がサブフォームならば親フォームの名前を返却
        strQueryId = objCallFrom.NAME
        Do
            Set objParent = Nothing
            Set objParent = objCallFrom.Parent
            If objParent Is Nothing Then
                '親フォームが存在しない
                Exit Do
            Else
                '親フォームが存在する
                strQueryId = objParent.NAME
                Set objCallFrom = objParent
            End If
        Loop Until objParent Is Nothing
    End If
    
    '呼び出しもと画面IDの二文字目を"Q"にする。
    strQueryId = Left(strQueryId, 1) & "Q" & Mid(strQueryId, 3)
    fncGetQueryId = strQueryId
    
    If Not objCallFrom Is Nothing Then
        Set objCallFrom = Nothing
    End If
    If Not objParent Is Nothing Then
        Set objParent = Nothing
    End If

End Function

'****************************  ended of program ********************************
