VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : コンテナ管理システム
'
'        PROGRAM_NAME    : クレジットカード登録メール送信
'        PROGRAM_ID      : FVS300
'
'        CREATE          : 2011/02/23
'        CERATER         : m.honda
'        Ver             : 0.0
'
'        UPDATE          : 2011/07/02
'        UPDATER         : M.RYU
'        Ver             : V0.1
'                        : 検索条件追加　⇒　保証会社コード、保証会社区分
'
'        UPDATE          : 2011/08/30
'        UPDATER         : M.RYU
'        Ver             : V0.2
'                        : 検索条件を追加⇒カード支払開始月・初回カード支払対象抽出
'                        : 抽出SQLを修正
'
'        UPDATE          : 2011/09/28
'        UPDATER         : M.RYU
'        Ver             : V0.3
'                        : クレジット・口座振替の場合、メール送信ボタン有効
'
'        UPDATE          : 2011/10/03
'        UPDATER         : M.RYU
'        Ver             : V0.4
'                        : 口座振替メール送信ボタンを追加
'
'        UPDATE          : 2011/11/11
'        UPDATER         : M.RYU
'        Ver             : V0.5
'                        : メール送信ボタンの表示条件を外す
'
'        UPDATE          : 2012/01/05
'        UPDATER         : M.RYU
'        Ver             : V0.6
'                        : FVS300_W01に顧客カナ項目を追加
'
'        UPDATE          : 2012/07/07
'        UPDATER         : M.HONDA
'        Ver             : V0.7
'                        : サイズの小数点対応
'
'        UPDATE          : 2013/06/03
'        UPDATER         : M.HONDA
'        Ver             : V0.8
'                        : 終了メッセージを表示させないように修正
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 接続部門
Private strBUMOC        As String
Private strKASE         As String

'接続ＤＢオブジェクト
Private dbAccess        As Database

'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'更新権限取得
Public bAuthority      As Boolean



'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

Private Const P_WORK_TABLE          As String = "FVS300_W01"
Private Const P_SUB_FORM            As String = "FVS300_S01"

'==============================================================================*
'
'        MODULE_NAME      :cmd_excel_Click
'        機能             :検索ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()
On Error GoTo err_rtn
    'ワークテーブル作成
    If Not MakeWork(True) Then
        Exit Sub
    End If
    
    doCmd.OutputTo acQuery, "FVS300_Q01", "MicrosoftExcel(*.xls)", "", True, ""
    
    Exit Sub

err_rtn:
    Call MSZZ024_M00("MakeWork", True)
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_検索_Click
'        機能             :検索ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索_Click()
    
    Dim strYM As String
    '入力チェック
    With SubCondition.Form       'INSERT 2011/08/30 M.RYU---<S>
        If .chk_CardObj.VALUE = True Then
            strYM = Nz(.txt_CardReqFrom)
            If strYM = "" Then
                MsgBox "カード支払開始月を入力してください。", vbOKOnly, Me.NAME
                Exit Sub
            End If
            If Len(strYM) <> 6 Then
                MsgBox "カード支払開始月を6桁の「YYYYMM」で入力してください。", vbOKOnly, Me.NAME
                Exit Sub
            End If
            If Right(strYM, 2) > 12 Or Right(strYM, 2) < 1 Then
                MsgBox "カード支払開始月を正しく入力してください。", vbOKOnly, Me.NAME
                Exit Sub
            End If
        End If                  'INSERT 2011/08/30 M.RYU---<E>
    End With
    
    '入力値を検索条件にしてワークテーブル作成
    Call subFormReload

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        機能             :終了ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

    '' 2013/06/03 M.HONDA START
'    On Error Resume Next
'    '終了メッセージを表示
'    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton1, P_KAKUNIN) = vbYes Then
'        doCmd.Close acForm, Me.FormName
'    End If
    '' 2013/06/03 M.HONDA END
    
    doCmd.Close acForm, Me.FormName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()

    doCmd.Maximize

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'F1=「終了」 F2=「検索」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Me.cmd_検索.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form Loadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim blnRet          As Boolean
    Dim strTemp         As String
    Dim strParam()      As String
    Dim strKOMS         As String
    Dim adoDbConnection As Object
    
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    strKOMS = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME_" & strBUMOC & "'")

    If strKOMS = "" Then
        Call MSZZ024_M10("SelectSetuTabl", "SETU_TABLに該当部門(" & strBUMOC & ")のデータが存在しません。")
    End If

    'KASE接続先データベース名取得
    strKASE = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATABASE_NAME'")
    
    'ワーククリア
    If Not ClearWork(P_WORK_TABLE) Then
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If

   '更新権限取得
    bAuthority = isAuthority
    
    ' コンボボックスの初期値設定
    With SubCondition.Form
        .cmb_集金区分.VALUE = "07"
        .cmb_毎月回収区分.VALUE = "02"
        .txt_集金区分 = "コンビニ"
        .txt_毎月回収区分 = "クレジットカード"
        .cmd_Mail.Enabled = False
        .chk_CardObj.VALUE = False          'INSERT 2011/08/30 M.RYU
        .txt_CardReqFrom.Enabled = False    'INSERT 2011/08/30 M.RYU
    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SubFormReload
'        機能             :一覧再表示
'        IN               :True=該当データ無しメッセージを表示 False=表示しない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function subFormReload(Optional blnNotFoundMsg As Boolean = True) As Boolean

    subFormReload = False
    
    'ワークテーブル作成
    If Not MakeWork(blnNotFoundMsg) Then
        Exit Function
    End If
    
    SubForm.Requery
    
    subFormReload = True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :ClearWork
'        機能             :ワークテーブルクリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ClearWork(strTable As String) As Boolean
On Error GoTo err_rtn

    Dim tdfNew      As TableDef
    Dim blnConnect  As Boolean
    
    ClearWork = False
    blnConnect = False
    
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
        blnConnect = True
    End If
        
    SubForm.SourceObject = ""
    
    'ワークテーブル削除
    If TableExist(strTable) Then
        Call dbAccess.TableDefs.Delete(strTable)
    End If

    'ワークテーブル作成
    Set tdfNew = dbAccess.CreateTableDef(strTable)
    Call FieldAppend(tdfNew)
    Call dbAccess.TableDefs.Append(tdfNew)
    
    ClearWork = True
    
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    If blnConnect Then
        If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    End If
    
    If Not tdfNew Is Nothing Then Set tdfNew = Nothing
    
    SubForm.SourceObject = P_SUB_FORM
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :FieldAppend
'        機能             :ワークテーブル列作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FieldAppend(tdfNew As TableDef)

    Dim fldNew      As Field
    Dim intCount    As Integer

    With tdfNew
 
        Call .Fields.Append(.CreateField("SEND", DataTypeEnum.dbLong))          '' メール送信フラグ
        Call .Fields.Append(.CreateField("BUMOC", DataTypeEnum.dbText, 2))      '' 部門
        Call .Fields.Append(.CreateField("UCODE", DataTypeEnum.dbLong))         '' 顧客コード
        Call .Fields.Append(.CreateField("USER", DataTypeEnum.dbText, 36))      '' 顧客名称
        Call .Fields.Append(.CreateField("USERKANA", DataTypeEnum.dbText, 18))  '' 顧客カナ     'INSERT 20120105 M.RYU
        Call .Fields.Append(.CreateField("RENKBN", DataTypeEnum.dbText, 20))    '' 連絡区分
        Call .Fields.Append(.CreateField("TEL", DataTypeEnum.dbText, 15))       '' 電話番号
        Call .Fields.Append(.CreateField("MAIL", DataTypeEnum.dbText, 60))      '' メールアドレス
        Call .Fields.Append(.CreateField("YOUKTGENKBN", DataTypeEnum.dbLong))   '' 申込区分
        Call .Fields.Append(.CreateField("GENKBN", DataTypeEnum.dbText, 20))    '' 申込名称
        Call .Fields.Append(.CreateField("RCPTKAIHI", DataTypeEnum.dbLong))     '' 毎月支払方法区分
        Call .Fields.Append(.CreateField("KAIHI", DataTypeEnum.dbText, 20))     '' 毎月支払方法名称
        Call .Fields.Append(.CreateField("USERSKBN", DataTypeEnum.dbLong))      '' 集金区分
        Call .Fields.Append(.CreateField("SKBN", DataTypeEnum.dbText, 20))      '' 集金名称
        Call .Fields.Append(.CreateField("YOUKTYUKBN", DataTypeEnum.dbLong))    '' 契約区分
        Call .Fields.Append(.CreateField("YUKBN", DataTypeEnum.dbText, 20))     '' 契約名称
        Call .Fields.Append(.CreateField("YOUKTUKDATE", DataTypeEnum.dbDate))   '' 予約日
        Call .Fields.Append(.CreateField("RCPTUKDATE", DataTypeEnum.dbDate))    '' 受付日
        Call .Fields.Append(.CreateField("KEYSENDDATE", DataTypeEnum.dbDate))   '' 鍵送付日
        Call .Fields.Append(.CreateField("SENDD", DataTypeEnum.dbDate))         '' 申込書送信日
        Call .Fields.Append(.CreateField("CONTN", DataTypeEnum.dbText, 240))    '' コメント
        Call .Fields.Append(.CreateField("ESENDD", DataTypeEnum.dbDate))        '' 登録日
        Call .Fields.Append(.CreateField("ECONTN", DataTypeEnum.dbText, 240))   '' 登録日コメント
        Call .Fields.Append(.CreateField("YCODE", DataTypeEnum.dbLong))         '' ヤード
        Call .Fields.Append(.CreateField("NAME", DataTypeEnum.dbText, 36))      '' ヤード名称
        Call .Fields.Append(.CreateField("ADDR", DataTypeEnum.dbText, 72))      '' 住所
        Call .Fields.Append(.CreateField("CNO", DataTypeEnum.dbLong))           '' 部屋番号
        Call .Fields.Append(.CreateField("STEP", DataTypeEnum.dbText, 20))      '' 段
        Call .Fields.Append(.CreateField("USAGE", DataTypeEnum.dbText, 20))     '' 使用用途
        '' 2012/07/07 M.HONDA START
        '' Call .Fields.Append(.CreateField("SIZE", DataTypeEnum.dbDouble))      '' サイズ
        Call .Fields.Append(.CreateField("SIZE", DataTypeEnum.dbDouble))          '' サイズ
        '' 2012/07/07 M.HONDA END
        Call .Fields.Append(.CreateField("RENTKG", DataTypeEnum.dbLong))        '' 月額
    
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :MakeWork
'        機能             :ワークテーブル作成
'        IN               :True=該当データ無しメッセージを表示 False=表示しない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MakeWork(blnNotFoundMsg As Boolean) As Boolean
On Error GoTo err_rtn
    
    Dim objCon          As Object       'コネクション
    Dim rsSqlServer     As Object
    Dim rsAccess        As Recordset

    MakeWork = False
    
    doCmd.Hourglass True

    'SQLサーバー接続
    Set objCon = ADODB_Connection(strBUMOC)
    
    'ワークテーブルの元ネタ取得
    Set rsSqlServer = ADODB_Recordset(MakeMailSql(strKASE), objCon)

    'ACCESS接続
    Set dbAccess = CurrentDb
        
    'ワークテーブルクリア
    If Not ClearWork(P_WORK_TABLE) Then
        GoTo Exit_rtn
    End If

    'ワークテーブルレコード追加
    Set rsAccess = dbAccess.OpenRecordset(P_WORK_TABLE, dbOpenDynaset, dbAppendOnly)

    Dim strTemp As String

    Do While Not rsSqlServer.EOF
        Call AddWork(rsSqlServer, rsAccess)
        rsSqlServer.MoveNext
    Loop

    If rsSqlServer.RecordCount = 0 Then
        If blnNotFoundMsg Then
            MsgBox "該当データが存在しません。", vbInformation, P_KAKUNIN
        End If
        MakeWork = False
        GoTo Exit_rtn
    Else
        If bAuthority = True Then
'            If SubCondition.Form.cmb_毎月回収区分 = "02" Or SubCondition.Form.cmb_毎月回収区分 = "03" Then  'クレジット・口座振替 'INSERT 2011/09/28 M.RYU 'DELETE 2011/11/11 M.RYU
            SubCondition.Form.cmd_Mail.Enabled = True
            SubCondition.Form.cmd_Kouza.Enabled = True      'INSERT 2011/10/03 M.RYU
'            End If         'DELETE 2011/11/11 M.RYU
        End If
    End If
    
    
    MakeWork = True
    GoTo Exit_rtn
    
err_rtn:
    Call MSZZ024_M00("MakeWork", True)
    Err.Clear

Exit_rtn:
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    doCmd.Hourglass False
    Exit Function
End Function


'==============================================================================*
'
'        MODULE_NAME      :TableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :テーブル名
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function TableExist(strTableName As String) As Boolean

    Dim tdf As TableDef
    
    TableExist = False
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            TableExist = True
            Exit For
        End If
    Next tdf

End Function

'''==============================================================================*
'''
'''        MODULE_NAME      :MakeYoukSql
'''        機能             :予約トラン検索SQL文作成
'''        IN               :
'''        OUT              :SQL文
'''--------------------------DELETE 2011/08/30 M.RYU------------------------------
'''==============================================================================*
'''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
''Private Function MakeMailSql(strKASE As String) As String
''
''    Dim strSQL      As String
''    Dim strWHERE    As String
''
''    strSQL = "          SELECT CONT_MAST.CONT_BUMOC BUMOC,"
''    strSQL = strSQL & "        YOUKT_GENKBN , "
''    strSQL = strSQL & "        NAME1.NAME_NAME AS GENKBN, "
''    strSQL = strSQL & "        YOUKT_YUKBN, "
''    strSQL = strSQL & "        NAME4.NAME_NAME AS YUKBN, "
''    strSQL = strSQL & "        YOUKT_UKDATE, "
''    strSQL = strSQL & "        RCPT_UKDATE, "
''    strSQL = strSQL & "        YOUKT_YCODE, "
''    strSQL = strSQL & "        YOUKT_NO, "
''    strSQL = strSQL & "        CNTA_STEP, "
''    strSQL = strSQL & "        CNTA_SIZE, "
''    strSQL = strSQL & "        NAME7.NAME_NAME AS STEP, "
''    strSQL = strSQL & "        CNTA_USAGE, "
''    strSQL = strSQL & "        NAME6.NAME_NAME AS USAGE, "
''    strSQL = strSQL & "        YARD_NAME, "
''    strSQL = strSQL & "        YARD_YUBINO, "
''    strSQL = strSQL & "        YARD_ADDR_1 + YARD_ADDR_2 AS YARD_ADDR, "
''    strSQL = strSQL & "        RCPT_RENTKG, "
''    strSQL = strSQL & "        RCPT_KEY_SEND_DATE, "
''    strSQL = strSQL & "        YOUKT_UCODE, "
''    strSQL = strSQL & "        USER_NAME, "
''    strSQL = strSQL & "        USER_MAIL, "
''    strSQL = strSQL & "        USER_SKBN, "
''    strSQL = strSQL & "        NAME2.NAME_NAME AS SKBN, "
''    strSQL = strSQL & "        YOUKT_RENKBN, "
''    strSQL = strSQL & "        NAME3.NAME_NAME AS RENKBN, "
''    strSQL = strSQL & "        RCPT_KAIHI, "
''    strSQL = strSQL & "        Case USER_KEITAI "
''    strSQL = strSQL & "        WHEN ' ' THEN "
''    strSQL = strSQL & "           USER_TEL "
''    strSQL = strSQL & "        ELSE "
''    strSQL = strSQL & "           USER_KEITAI"
''    strSQL = strSQL & "        END TELNO, "
''    strSQL = strSQL & "        NAME5.NAME_NAME AS KAIHI, "
''    strSQL = strSQL & "        Case YOUKT_GENKBN"
''    strSQL = strSQL & "        WHEN 0 THEN"
''    strSQL = strSQL & "           convert(datetime,ISNULL(MAX(CONVERT(DATETIME,SKOTO.KOTOT_GENTD)),RCPT_KEY_SEND_DATE))"
''    strSQL = strSQL & "        WHEN 1 THEN"
''    strSQL = strSQL & "           CONVERT(DATETIME,ISNULL(MAX(CONVERT(DATETIME,SKOTO.KOTOT_GENTD)),YOUKT_UKDATE)) "
''    strSQL = strSQL & "        END SENDD, "
''    strSQL = strSQL & "        SKOTO.KOTOT_CONTN AS CONTN, "
''    strSQL = strSQL & "        MAX(CONVERT(DATETIME,EKOTO.KOTOT_GENTD)) AS ESENDD, "
''    strSQL = strSQL & "        EKOTO.KOTOT_CONTN AS ECONTN  "
''    strSQL = strSQL & " FROM "
''    strSQL = strSQL & "     YOUK_TRAN "
''
''    strSQL = strSQL & "     Left Join "
''    strSQL = strSQL & "     CARG_FILE ON "
''    strSQL = strSQL & "     YOUKT_RCPT_NO = CARG_UKNO "
''
''    strSQL = strSQL & "     Left Join "
''    strSQL = strSQL & "     " & strKASE & ".dbo.KOTO_TRAN SKOTO ON "
''    strSQL = strSQL & "     CONVERT(CHAR, YOUKT_UCODE) = SKOTO.KOTOT_KOKYC AND "
''    strSQL = strSQL & "     SKOTO.KOTOT_TYPEC = '09' "
''    strSQL = strSQL & "     Left Join "
''    strSQL = strSQL & "     CONT_MAST CONT1 ON "
''    strSQL = strSQL & "     SKOTO.KOTOT_BUMOC = CONT1.CONT_BUMOC "
''    strSQL = strSQL & "     Left Join "
''    strSQL = strSQL & "     " & strKASE & ".dbo.KOTO_TRAN EKOTO ON "
''    strSQL = strSQL & "     CONVERT(CHAR, YOUKT_UCODE) = EKOTO.KOTOT_KOKYC AND "
''    strSQL = strSQL & "     EKOTO.KOTOT_TYPEC = '10' "
''    strSQL = strSQL & "     Left Join "
''    strSQL = strSQL & "     CONT_MAST CONT2 ON "
''    strSQL = strSQL & "     EKOTO.KOTOT_BUMOC = CONT2.CONT_BUMOC, "
''    strSQL = strSQL & "     RCPT_TRAN, "
'''    strSQL = strSQL & "     CARG_FILE, "
''    strSQL = strSQL & "     YARD_MAST, "
''    strSQL = strSQL & "     USER_MAST, "
''    strSQL = strSQL & "     CNTA_MAST, "
''    strSQL = strSQL & "     NAME_MAST NAME1, "
''    strSQL = strSQL & "     NAME_MAST NAME2, "
''    strSQL = strSQL & "     NAME_MAST NAME3, "
''    strSQL = strSQL & "     NAME_MAST NAME4, "
''    strSQL = strSQL & "     NAME_MAST NAME5, "
''    strSQL = strSQL & "     NAME_MAST NAME6, "
''    strSQL = strSQL & "     NAME_MAST NAME7, "
''    strSQL = strSQL & "     CONT_MAST CONT_MAST "
''    strSQL = strSQL & " WHERE YOUKT_RCPT_NO = RCPT_NO "
''    strSQL = strSQL & "  AND  YOUKT_YCODE = YARD_CODE "
''    strSQL = strSQL & "  AND  YOUKT_UCODE = USER_CODE "
''    strSQL = strSQL & "  AND  YOUKT_YCODE = CNTA_CODE "
''    strSQL = strSQL & "  AND  YOUKT_NO    = CNTA_NO "
'''    strSQL = strSQL & "  AND  RCPT_CARG_ACPTNO = CARG_ACPTNO "
''    strSQL = strSQL & "  AND  CARG_KYDATE IS NULL "
''    strSQL = strSQL & "  AND  NAME1.NAME_ID = '083' "
''    strSQL = strSQL & "  AND  NAME1.NAME_CODE = YOUKT_GENKBN "
''    strSQL = strSQL & "  AND  NAME2.NAME_ID = '030' "
''    strSQL = strSQL & "  AND  NAME2.NAME_CODE = USER_SKBN "
''    strSQL = strSQL & "  AND  NAME3.NAME_ID = '085' "
''    strSQL = strSQL & "  AND  NAME3.NAME_CODE = YOUKT_RENKBN "
''    strSQL = strSQL & "  AND  NAME4.NAME_ID = '080' "
''    strSQL = strSQL & "  AND  NAME4.NAME_CODE = YOUKT_YUKBN "
''    strSQL = strSQL & "  AND  NAME5.NAME_ID = '017' "
''    strSQL = strSQL & "  AND  NAME5.NAME_CODE = ISNULL(RCPT_KAIHI,1) "
''    strSQL = strSQL & "  AND  NAME6.NAME_ID =  '086' "
''    strSQL = strSQL & "  AND  NAME6.NAME_CODE = CNTA_USAGE "
''    strSQL = strSQL & "  AND  NAME7.NAME_ID =  '090' "
''    strSQL = strSQL & "  AND  NAME7.NAME_CODE = CNTA_STEP "
''    strWHERE = MakeWhere()
''    strSQL = strSQL & strWHERE
''    strSQL = strSQL & " GROUP BY  YOUKT_GENKBN,"
''    strSQL = strSQL & "           NAME1.NAME_NAME,"
''    strSQL = strSQL & "           YOUKT_YUKBN,"
''    strSQL = strSQL & "           NAME4.NAME_NAME,"
''    strSQL = strSQL & "           YOUKT_UKDATE,"
''    strSQL = strSQL & "           RCPT_UKDATE,"
''    strSQL = strSQL & "           YOUKT_YCODE,"
''    strSQL = strSQL & "           YARD_NAME,"
''    strSQL = strSQL & "           YARD_YUBINO, "
''    strSQL = strSQL & "           CNTA_STEP, "
''    strSQL = strSQL & "           CNTA_SIZE, "
''    strSQL = strSQL & "           CNTA_USAGE, "
''    strSQL = strSQL & "           YARD_ADDR_1 + YARD_ADDR_2, "
''    strSQL = strSQL & "           YOUKT_NO,"
''    strSQL = strSQL & "           RCPT_RENTKG, "
''    strSQL = strSQL & "           RCPT_KEY_SEND_DATE,"
''    strSQL = strSQL & "           YOUKT_UCODE,"
''    strSQL = strSQL & "           USER_NAME,"
''    strSQL = strSQL & "           USER_MAIL,"
''    strSQL = strSQL & "           USER_TEL,"
''    strSQL = strSQL & "           USER_KEITAI,"
''    strSQL = strSQL & "           USER_FAX,"
''    strSQL = strSQL & "           USER_SKBN,"
''    strSQL = strSQL & "           NAME2.NAME_NAME, "
''    strSQL = strSQL & "           YOUKT_RENKBN, "
''    strSQL = strSQL & "           NAME3.NAME_NAME, "
''    strSQL = strSQL & "           RCPT_KAIHI, "
''    strSQL = strSQL & "           NAME5.NAME_NAME, "
''    strSQL = strSQL & "           NAME6.NAME_NAME, "
''    strSQL = strSQL & "           NAME7.NAME_NAME, "
''    strSQL = strSQL & "           SKOTO.KOTOT_CONTN, "
''    strSQL = strSQL & "           EKOTO.KOTOT_CONTN, "
''    strSQL = strSQL & "           CONT_MAST.CONT_BUMOC "
''    MakeMailSql = strSQL & "ORDER BY YOUKT_UKDATE,YOUKT_UCODE"
''
''End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeYoukSql
'        機能             :予約トラン検索SQL文作成
'        IN               :
'        OUT              :SQL文
'        CREAT            :2011/08/30 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeMailSql(strKASE As String) As String

    Dim strSQL      As String
    Dim strWhere    As String
    
    strWhere = MakeWhere()
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT DTBL.* FROM (   "
    strSQL = strSQL & " SELECT                  "
    strSQL = strSQL & "     CONT_MAST.CONT_BUMOC BUMOC, "
    strSQL = strSQL & "     YOUKT_GENKBN,       "
    strSQL = strSQL & "     (SELECT NAME_NAME FROM NAME_MAST WHERE NAME_ID = '083' AND NAME_CODE = YOUKT_GENKBN ) AS GENKBN,    "
    strSQL = strSQL & "     YOUKT_YUKBN,        "
    strSQL = strSQL & "     (SELECT NAME_NAME FROM NAME_MAST WHERE NAME_ID = '080' AND NAME_CODE = YOUKT_YUKBN ) AS YUKBN,      "
    strSQL = strSQL & "     YOUKT_UKDATE,       "
    strSQL = strSQL & "     RCPT_UKDATE,        "
    strSQL = strSQL & "     YOUKT_YCODE,        "
    strSQL = strSQL & "     YOUKT_NO,           "
    strSQL = strSQL & "     CNTA_STEP,          "
    strSQL = strSQL & "     CNTA_SIZE,          "
    strSQL = strSQL & "     (SELECT NAME_NAME FROM NAME_MAST WHERE NAME_ID = '090' AND NAME_CODE = CNTA_STEP ) AS STEP,     "
    strSQL = strSQL & "     CNTA_USAGE,         "
    strSQL = strSQL & "     (SELECT NAME_NAME FROM NAME_MAST WHERE NAME_ID = '086' AND NAME_CODE = CNTA_USAGE ) AS USAGE,   "
    strSQL = strSQL & "     YARD_NAME,          "
    strSQL = strSQL & "     YARD_YUBINO,        "
    strSQL = strSQL & "     YARD_ADDR_1 + YARD_ADDR_2 AS YARD_ADDR, "
    strSQL = strSQL & "     RCPT_RENTKG,        "
    strSQL = strSQL & "     RCPT_KEY_SEND_DATE, "
    strSQL = strSQL & "     RIGHT('000000' + CONVERT(varchar, YOUKT_UCODE), 6) AS YOUKT_UCODE,  "
    strSQL = strSQL & "     USER_NAME,  "
    strSQL = strSQL & "     USER_KANA,  "           'INSERT 20120105 M.RYU
    strSQL = strSQL & "     USER_MAIL,  "
    strSQL = strSQL & "     USER_SKBN,  "
    strSQL = strSQL & "     (SELECT NAME_NAME FROM NAME_MAST WHERE NAME_ID = '030' AND NAME_CODE = USER_SKBN ) AS SKBN,         "
    strSQL = strSQL & "     YOUKT_RENKBN,       "
    strSQL = strSQL & "     (SELECT NAME_NAME FROM NAME_MAST WHERE NAME_ID = '085' AND NAME_CODE = YOUKT_RENKBN ) AS RENKBN,    "
    strSQL = strSQL & "     RCPT_KAIHI,     "
    strSQL = strSQL & "     CASE USER_KEITAI   WHEN ' ' THEN USER_TEL ELSE   USER_KEITAI  END TELNO,    "
    strSQL = strSQL & "     (SELECT NAME_NAME FROM NAME_MAST WHERE NAME_ID = '017' AND NAME_CODE = ISNULL(RCPT_KAIHI,1) ) AS KAIHI, "
    strSQL = strSQL & "     CASE YOUKT_GENKBN  WHEN 0 THEN CONVERT(DATETIME,ISNULL(MAX(CONVERT(DATETIME,SKOTO.KOTOT_GENTD)),RCPT_KEY_SEND_DATE))        "
    strSQL = strSQL & "                        WHEN 1 THEN CONVERT(DATETIME,ISNULL(MAX(CONVERT(DATETIME,SKOTO.KOTOT_GENTD)),YOUKT_UKDATE))  END SENDD,  "
    strSQL = strSQL & "     SKOTO.KOTOT_CONTN AS CONTN,         "
    strSQL = strSQL & "     MAX(CONVERT(DATETIME,EKOTO.KOTOT_GENTD)) AS ESENDD, "
    strSQL = strSQL & "     EKOTO.KOTOT_CONTN AS ECONTN,        "
    strSQL = strSQL & "     MAX(RCPT_SEIKYU_KIKAN) AS SEIKIKAN  "
    strSQL = strSQL & " FROM YOUK_TRAN  "
    strSQL = strSQL & "     LEFT  JOIN CARG_FILE ON YOUKT_RCPT_NO = CARG_UKNO   "
    strSQL = strSQL & "     INNER JOIN RCPT_TRAN ON YOUKT_RCPT_NO = RCPT_NO "
    strSQL = strSQL & "     INNER JOIN YARD_MAST ON YOUKT_YCODE = YARD_CODE "
    strSQL = strSQL & "     INNER JOIN USER_MAST ON YOUKT_UCODE = USER_CODE "
    strSQL = strSQL & "     INNER JOIN CNTA_MAST ON YOUKT_YCODE = CNTA_CODE "
    strSQL = strSQL & "     INNER JOIN CONT_MAST ON YOUKT_NO = CNTA_NO  "
    strSQL = strSQL & "     LEFT JOIN " & strKASE & ".dbo.KOTO_TRAN SKOTO ON CONVERT(CHAR, YOUKT_UCODE) = SKOTO.KOTOT_KOKYC AND SKOTO.KOTOT_TYPEC = '09'    "
    strSQL = strSQL & "     LEFT JOIN " & strKASE & ".dbo.KOTO_TRAN EKOTO ON CONVERT(CHAR, YOUKT_UCODE) = EKOTO.KOTOT_KOKYC AND EKOTO.KOTOT_TYPEC = '10'    "
    strSQL = strSQL & " WHERE CARG_KYDATE IS NULL   "
    strSQL = strSQL & strWhere
    strSQL = strSQL & " GROUP BY    "
    strSQL = strSQL & "     CONT_MAST.CONT_BUMOC,    YOUKT_GENKBN, YOUKT_YUKBN  ,  "
    strSQL = strSQL & "     YOUKT_UKDATE, RCPT_UKDATE, YOUKT_YCODE ,   "
    strSQL = strSQL & "     YOUKT_NO,     CNTA_STEP,   CNTA_SIZE,  CNTA_USAGE   ,  "
    strSQL = strSQL & "     YARD_NAME,    YARD_YUBINO, YARD_ADDR_1 + YARD_ADDR_2,  "
    strSQL = strSQL & "     RCPT_RENTKG,  RCPT_KEY_SEND_DATE,      YOUKT_UCODE  ,  "
    strSQL = strSQL & "     USER_NAME,    USER_KANA,   USER_MAIL,  USER_SKBN, YOUKT_RENKBN , "
    strSQL = strSQL & "     RCPT_KAIHI,   SKOTO.KOTOT_CONTN ,      EKOTO.KOTOT_CONTN ,  "
    strSQL = strSQL & "     USER_KEITAI,  USER_TEL,    USER_KEITAI     "
    strSQL = strSQL & " )  DTBL "
    
    With SubCondition.Form
        If .chk_CardObj.VALUE = True Then
            strSQL = strSQL & " WHERE   "
            strSQL = strSQL & "  DTBL.YOUKT_UCODE NOT IN (  "
            strSQL = strSQL & "     SELECT DISTINCT KAIST_KOKYC FROM " & strKASE & ".dbo.KAIS_TRAN KAIST        "
            strSQL = strSQL & "      WHERE KAIST.KAIST_BUMOC = DTBL.BUMOC AND KAIST.KAIST_KAISI = '6' "
            strSQL = strSQL & "      AND KAIST_KAYMD <= " & .txt_CardReqFrom & ")"
            strSQL = strSQL & "   AND DTBL.SEIKIKAN < " & .txt_CardReqFrom & "   "  '請求期間
            strSQL = strSQL & "   AND DTBL.ESENDD   < " & .txt_CardReqFrom & "   "  'カード登録日
            
        End If
    End With
    
    strSQL = strSQL & " ORDER BY YOUKT_UKDATE , YOUKT_UCODE ;   "
    
    MakeMailSql = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeWhere
'        機能             :Where句作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeWhere() As String

    Dim strWhere        As String
    
    strWhere = ""
    
    With SubCondition.Form
        If .chk_CardObj.VALUE = True Then                                                   'INSERT 2011/08/30 M.RYU
            strWhere = strWhere & " AND USER_SKBN = '04'                        "           'INSERT 2011/08/30 M.RYU
            strWhere = strWhere & " AND ISNULL(YOUKT_YUKBN,0) IN ('10' , '20')  "           'INSERT 2011/08/30 M.RYU
        Else
            ' 集金区分
            If Nz(.cmb_集金区分) <> "" Then
                strWhere = strWhere & " AND USER_SKBN = '" & .cmb_集金区分 & "' "
            End If
            
            ' 毎月回収区分
            If Nz(.cmb_毎月回収区分) <> "" Then
                strWhere = strWhere & " AND ISNULL(RCPT_KAIHI,1) = '" & .cmb_毎月回収区分 & "' "
            End If
            
            ' 受付種別
            If Nz(.cmb_受付種別) <> "" Then
                strWhere = strWhere & " AND ISNULL(YOUKT_GENKBN,0) = '" & .cmb_受付種別 & "' "
            End If
            
            ' 状態
            If Nz(.cmb_YukbnFrom) <> "" Then
                strWhere = strWhere & " AND ISNULL(YOUKT_YUKBN,0) = '" & .cmb_YukbnFrom & "' "
            Else
                strWhere = strWhere & " AND ISNULL(YOUKT_YUKBN,0) IN ('10' , '20') "        'INSERT 2011/08/30 M.RYU
            End If
    
            '顧客コード
            If Nz(.cmb_Ucode) <> "" Then
                '顧客コードが入力されている場合はそれで一致条件。カナは無視。
                strWhere = strWhere & MakeEqual("YOUKT_UCODE", Nz(.cmb_Ucode))
            End If
            
            ' 予約日
            strWhere = strWhere & MakeBetween("YOUKT_KUDATE", Nz(.txt_UkdateFrom), Nz(.txt_UkdateTo))
            
            ' 鍵送付日
            strWhere = strWhere & MakeBetween("RCPT_KEY_SEND_DATE", Nz(.txt_KeySendDateFrom), Nz(.txt_KeySendDateTo))
            
            '保証会社
            If Nz(.cmb_HosyId) <> "" Then                                                       'INSERT 2011/07/02 M.RYU
                strWhere = strWhere & " AND CARG_FILE.CARG_HOSYICD = " & .cmb_HosyId            'INSERT 2011/07/02 M.RYU
            End If
            
            '保証会社
            If Nz(.cmb_HosyC) <> "" Then                                                        'INSERT 2011/07/02 M.RYU
                strWhere = strWhere & MakeEqual("CARG_FILE.CARG_HOSYO_CD", .cmb_HosyC)          'INSERT 2011/07/02 M.RYU
            End If
            
        End If
        
    End With
    
    MakeWhere = strWhere
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeEqual
'        機能             :一致条件作成
'        IN               :第一引数－対象テーブルのカラム名
'                         :第二引数－条件値
'        OUT              :条件文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeEqual(strColName As String, strCond As String) As String

    MakeEqual = ""
    
    If strCond <> "" Then
        '指定されたカラム名と”＝”で条件値をくっつける
        MakeEqual = " AND " & strColName & " = '" & strCond & "' "
    End If
            
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeBetween
'        機能             :範囲条件作成
'        IN               :第一引数－対象テーブルのカラム名
'                         :第二引数－範囲条件値FROM
'                         :第三引数－範囲条件値TO
'        OUT              :条件文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeBetween(strColName As String, strfrom As String, strTo As String) As String

    Dim strTemp     As String
    
    strTemp = ""
    
    If strfrom <> "" And strTo <> "" Then
        '共に空白ではない場合、
        If strfrom = strTo Then
            'FROMとTOが同一の場合、一致条件
            strTemp = MakeEqual(strColName, strfrom)
        Else
            'FROMとTOが異なる場合、BETWEEN条件
            strTemp = " AND " & strColName & " BETWEEN '" & strfrom & "' AND '" & strTo & "' "
        End If
    ElseIf strfrom <> "" Then
        'FROMのみの場合、それ以上であることが条件
        strTemp = " AND " & strColName & " >= '" & strfrom & "' "
    ElseIf strTo <> "" Then
        'TOのみの場合、それ以下であることが条件
        strTemp = " AND " & strColName & " <= '" & strTo & "' "
    End If
    
    MakeBetween = strTemp

End Function

'==============================================================================*
'
'        MODULE_NAME      :AddWork
'        機能             :作業テーブル追加
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub AddWork(rsSource As Object, rsDestination As Recordset)

    With rsDestination
        .AddNew
 
        .Fields("SEND") = 0                                                     '' 送信フラグ
        .Fields("BUMOC") = rsSource.Fields("BUMOC")                             '' 部門コード
        .Fields("UCODE") = rsSource.Fields("YOUKT_UCODE")                       '' 顧客コード
        .Fields("USER") = rsSource.Fields("USER_NAME")                          '' 顧客名
        .Fields("USERKANA") = rsSource.Fields("USER_KANA")                      '' 顧客カナ         'INSERT 20120105 M.RYU
        .Fields("RENKBN") = rsSource.Fields("RENKBN")                           '' 連絡区分
        .Fields("TEL") = rsSource.Fields("TELNO")                               '' 電話番号
        .Fields("MAIL") = rsSource.Fields("USER_MAIL")                          '' メールアドレス
        .Fields("YOUKTGENKBN") = rsSource.Fields("YOUKT_GENKBN")                '' 申込区分
        .Fields("GENKBN") = rsSource.Fields("GENKBN")                           '' 申込名称
        .Fields("RCPTKAIHI") = rsSource.Fields("RCPT_KAIHI")                    '' 毎月支払方法
        .Fields("KAIHI") = rsSource.Fields("KAIHI")                             '' 毎月支払方法名称
        .Fields("USERSKBN") = rsSource.Fields("USER_SKBN")                      '' 集金区分
        .Fields("SKBN") = rsSource.Fields("SKBN")                               '' 集金区分名称
        .Fields("YOUKTYUKBN") = rsSource.Fields("YOUKT_YUKBN")                  '' 契約区分
        .Fields("YUKBN") = rsSource.Fields("YUKBN")                             '' 契約名称
        .Fields("YOUKTUKDATE") = rsSource.Fields("YOUKT_UKDATE")                '' 予約日
        .Fields("RCPTUKDATE") = rsSource.Fields("RCPT_UKDATE")                  '' 受付日
        .Fields("KEYSENDDATE") = rsSource.Fields("RCPT_KEY_SEND_DATE")          '' 鍵送付日
        .Fields("SENDD") = rsSource.Fields("SENDD")                             '' 申込書送信日
        .Fields("CONTN") = rsSource.Fields("CONTN")                             '' コメント
        .Fields("ESENDD") = rsSource.Fields("ESENDD")                           '' 登録日
        .Fields("ECONTN") = rsSource.Fields("ECONTN")                           '' 登録日コメント
        .Fields("YCODE") = rsSource.Fields("YOUKT_YCODE")                       '' ヤードコード
        .Fields("NAME") = rsSource.Fields("YARD_NAME")                          '' ヤード名称
        .Fields("ADDR") = rsSource.Fields("YARD_ADDR")                          '' 住所
        .Fields("CNO") = rsSource.Fields("YOUKT_NO")                            '' 部屋番号
        .Fields("STEP") = rsSource.Fields("STEP")                               '' 段
        .Fields("USAGE") = rsSource.Fields("USAGE")                             '' 使用用途
        .Fields("SIZE") = rsSource.Fields("CNTA_SIZE")                          '' サイズ
        .Fields("RENTKG") = rsSource.Fields("RCPT_RENTKG")                      '' 月額
        .UPDATE

    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form_Openイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    doCmd.Maximize

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload
'        機能             :Form_Unloadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    'Access側ワークテーブル削除
    Call ClearWork(P_WORK_TABLE)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :subReload
'        CREATE_DATE      :2005/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subReload(Optional intFlg As Integer = 0)

    Call FormEnable(True)
    SubForm.Form.cmd_Gosyoukai.Enabled = True
    SubForm.Form.cmd_YOYAKU.Enabled = True
    SubForm.Form.cmd_Uketsuke.Enabled = True
    SubForm.Form.cmd_Gosyoukai.SetFocus
    SubForm.Form.txt_Focus.Visible = False


    If intFlg = 1 Then
        Call subFormReload(False)
        SubForm.Requery
    End If

End Sub


'==============================================================================*
'
'        MODULE_NAME      :項目使用可不可設定
'        MODULE_ID        :FormEnable
'        CREATE_DATE      :2005/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub FormEnable(blnEnable As Boolean)

    SubCondition.Enabled = blnEnable
    cmd_終了.Enabled = blnEnable
    cmd_検索.Enabled = blnEnable


End Sub
'==============================================================================*
'
'        MODULE_NAME      :MakeDateTimeBetween
'        機能             :日時を考慮した範囲条件作成
'        IN               :第一引数－対象テーブルのカラム名
'                         :第二引数－範囲条件値FROM
'                         :第三引数－範囲条件値TO
'        OUT              :条件文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeDateTimeBetween(strColName As String, strfrom As String, strTo As String) As String

    Dim strTemp     As String
    
    strTemp = ""
    
    If strfrom <> "" And strTo <> "" Then
            strTemp = " AND " & strColName & " BETWEEN '" & strfrom & " 00:00:00' AND '" & _
                                                              strTo & " 23:59:59' "
    ElseIf strfrom <> "" Then
        'FROMのみの場合、それ以上であることが条件
        strTemp = " AND " & strColName & " >= '" & strfrom & " 00:00:00' "
    ElseIf strTo <> "" Then
        'TOのみの場合、それ以下であることが条件
        strTemp = " AND " & strColName & " <= '" & strTo & " 23:59:59' "
    End If
    
    MakeDateTimeBetween = strTemp
End Function

'==============================================================================*
'
'        MODULE_NAME      :メール権限可否
'        MODULE_ID        :IsAuthority
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = ユーザID
'        戻り値           : True...更新可
'                         : False...更新不可
'        CREATE_DATE      :
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function isAuthority() As Boolean
                              
    Dim sqlText             As String
    Dim rsData              As Object
    Dim adoDbConnection     As Object
    Dim userId              As String
     
    isAuthority = False
    
    On Error GoTo Exception
    
    'ユーザ情報取得
    userId = MSZZ000.LsGetUserName()

  
    'ADO接続Object生成
    Set adoDbConnection = MSZZ025.ADODB_Connection()

    sqlText = "SELECT * "
    sqlText = sqlText & " FROM ONLE_TABL "
    sqlText = sqlText & " WHERE ONLET_USERB = '" & userId & "' "
    sqlText = sqlText & " AND   ONLET_PROGB = 'FVS300' "
    sqlText = sqlText & "  AND ONLET_ACCEI <> '9' "

    Set rsData = MSZZ025.ADODB_Recordset(sqlText, adoDbConnection, adoReadOnly)
    
    Do Until rsData.EOF
        ' ボタンの使用可否を設定する
        If (Nz(rsData.Fields("ONLET_ACCEI"), "")) = 1 Then
            isAuthority = True
        Else
            isAuthority = False
        End If
        rsData.MoveNext
    Loop
    
    adoDbConnection.Close
    Set adoDbConnection = Nothing
   
    Exit Function

Exception:
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing
    Call Err.Raise(Err.Number, "IsAuthority" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'****************************  ended of program ********************************


