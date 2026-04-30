VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS540"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約一覧
'   プログラムＩＤ　：　FVS540
'   作　成　日　　　：  2005/12/20
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/1/24
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.1
'   修　正　内　容　：  1.契約入力画面に遷移するボタン追加
'   修　正　内　容　：  2.検索条件に移動理由追加
'   修　正　内　容　：  3.一覧に移動区分追加
'   修　正　内　容　：  4.件数表示追加
'
'   修　正　日　　　：  2006/02/08
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.2
'   修　正　内　容　：  1.他画面から戻った際にデータが存在しない場合に検索件数が０にならない不具合修正
'
'   修　正　日　　　：  2006/03/09 2006/04/06
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.3
'   修　正　内　容　：  1.契約入力画面への遷移を受付入力画面への遷移に変更(3/9)
'                  ：  2.鍵配送状況を表示(4/6)
'                  ：  3.条件追加(4/6)
'                   :    ・鍵配送日
'                   :    ・鍵配送状況
'                   :    ・伝票番号
'
'   修　正　日　　　：  2006/06/07
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.4
'   修　正　内　容　：  1.他画面からのリフレッシュ時に該当データ無しメッセージは表示しない
'
'   修　正　日　　　：  2007/08/23
'   修　正　者　　　：  tajima
'   Ver             ：  0.5
'   修　正　内　容　：  ネット予約対応、入力条件及び一覧に項目追加
'
'   修　正　日　　　：  2012/07/04
'   修　正　者　　　：  M.HONDA
'   Ver             ：  0.6
'   修　正　内　容　：  ログ出力対応
'
'   修　正　日　　　：  2012/11/01
'   修　正　者　　　：  M.HONDA
'   Ver             ：  0.7
'   修　正　内　容　：  エクセル出力ボタンを追加
'
'   修　正　日　　　：  2013/06/03
'   修　正　者　　　：  M.HONDA
'   Ver             ：  0.8
'   修　正　内　容　：  終了メッセージを表示させないように修正。
'
'   修　正　日　　　：  2014/12/25
'   修　正　者　　　：  M.HONDA
'   Ver             ：  0.9
'   修　正　内　容　：  入金日入力データも表示
'
'   修　正　日　　　：  2015/01/08
'   修　正　者　　　：  M.HONDA
'   Ver             ：  1.0
'   修　正　内　容　：  運搬サービスの内容を追加
'
'   修　正　日　　　：  2015/02/03
'   修　正　者　　　：  M.HONDA
'   Ver             ：  1.1
'   修　正　内　容　：  身分証明書を追加
'
'   修　正　日　　　：  2015/05/01
'   修　正　者　　　：  M.HONDA
'   Ver             ：  1.2
'   修　正　内　容　：  起算日・集金区分・担当者名称・初回支払方法を追加
'
'   修　正　日　　　：  2015/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  1.3
'   修　正　内　容　：  条件に身分証明書を確認する必要のあるもののみ表示を追加
'                       項目に身分証明書確認済みを追加
'
'   修　正　日　　　：  2015/07/27
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  1.4
'   修　正　内　容　：  Ver1.3で実際のクエリの条件に入ってなかった
'
'   修　正　日　　　：  2018/03/27
'   修　正　者　　　：  N.IMAI
'   Ver             ：  1.5
'   修　正　内　容　：  Ver1.3で実際のクエリの条件に入ってなかった
'
'   修　正　日　　　：  2022/08/22
'   修　正　者　　　：  N.IMAI
'   Ver             ：  1.6
'   修　正　内　容　：  検索条件追加：初期費用請求方法、入金確認日、メール発行日
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'接続ＤＢオブジェクト
Private dbSQLServer     As Database
Private dbAccess        As Database

'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FVS540"

Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

Private Const P_WORK_TABLE          As String = "FVS540_W01"
Private Const P_SUB_FORM            As String = "FVS540_S01"

Private Const P_予約状態_契約        As Integer = 20

Private Const P_自動条件_全て       As String = "1"
Private Const P_自動条件_自動のみ   As String = "2"
Private Const P_自動条件_自動除く   As String = "3"

Private Const P_自動条件名_全て     As String = "全て"
Private Const P_自動条件名_自動のみ As String = "自動のみ"
Private Const P_自動条件名_自動除く As String = "自動除く"

Private Const P_紹介区分_取置き     As String = "1"

Private Const P_連絡区分_電話       As String = "0"
Private Const P_連絡区分_携帯       As String = "1"
Private Const P_連絡区分_FAX        As String = "2"

'06.1.24 SHIBAZAKI 追加
Private Const P_LABEL_件数表示      As String = "該当予約 : @件"
Private Const P_MOVEKBN_無し        As Integer = 0

'配送状況
Private Const P_DERIVERYSTATUS_配達完了     As String = "配達完了"

'==============================================================================*
'
'        MODULE_NAME      :cmd_Excel_Click
'        機能             :Excelボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    'ワークテーブル作成
    If Not MakeWork(False) Then
        Exit Sub
    End If
    
    doCmd.OutputTo acQuery, "FVS540_Q01", "MicrosoftExcel(*.xls)", "", True, ""

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
    On Error Resume Next
    '' 2013/06/03 M.HONDA START
    '終了メッセージを表示
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
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    
    Call MSHS010_M00(PROG_ID, "2")
    
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
    Dim strBumonCode    As String
    
    'SQL-Server接続
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumonCode & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    '接続文字列取得
    strConnect = MSZZ007_M00(strBumonCode)
    If strConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    
    With SubCondition.Form.cmb_AUTO
        Call .AddItem(P_自動条件_全て + ";" + P_自動条件名_全て)
        Call .AddItem(P_自動条件_自動のみ + ";" + P_自動条件名_自動のみ)
        Call .AddItem(P_自動条件_自動除く + ";" + P_自動条件名_自動除く)
    End With
    
    'ワーククリア
    If Not ClearWork(P_WORK_TABLE) Then
        doCmd.Close acForm, Me.FormName
        Exit Sub
    End If
    
    
    '06.1.24 SHIBAZAKI 追加
    '件数表示
    Call DispCount(0)
    
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
'2006/06/06 SHIBAZAKI 他画面からのリフレッシュ時に該当データ無しメッセージは表示しない
'Private Function SubFormReload() As Boolean
Private Function subFormReload(Optional blnNotFoundMsg As Boolean = True) As Boolean

    subFormReload = False
    
    '06.1.24 SHIBAZAKI 追加
    '件数表示
    Call DispCount(0)
    
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
 
        Call .Fields.Append(.CreateField("UKNO", DataTypeEnum.dbText, 11))
        Call .Fields.Append(.CreateField("YCODE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("CNO", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("NAME", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("SIZE_FROM", DataTypeEnum.dbDouble))
        Call .Fields.Append(.CreateField("SIZE_TO", DataTypeEnum.dbDouble))
        Call .Fields.Append(.CreateField("STEP", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("UKDATE", DataTypeEnum.dbDate))
        Call .Fields.Append(.CreateField("UCODE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("USER", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("TEL", DataTypeEnum.dbText, 15))
        Call .Fields.Append(.CreateField("RENKUBUN", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("BIKO", DataTypeEnum.dbText, 60))
        Call .Fields.Append(.CreateField("TANTC", DataTypeEnum.dbText, 3))
        Call .Fields.Append(.CreateField("TANTN", DataTypeEnum.dbText, 18))
        Call .Fields.Append(.CreateField("UKEKUBUN", DataTypeEnum.dbText, 80))
        Call .Fields.Append(.CreateField("AUTOKUBUN", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("TKDATE", DataTypeEnum.dbDate))
        Call .Fields.Append(.CreateField("TKCOUNT", DataTypeEnum.dbLong))
        '06.1.24 SHIBAZAKI 追加
        Call .Fields.Append(.CreateField("UKKBNCODE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("MOVEKBNCODE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("MOVEKUBUN", DataTypeEnum.dbText, 20))
        '2006/03/09 SHIBAZAKI 追加
        Call .Fields.Append(.CreateField("MOTOACPTNO", DataTypeEnum.dbText, 10))
        Call .Fields.Append(.CreateField("RCPTNO", DataTypeEnum.dbText, 11))
        '2007/08/23 tajimaI 追加
        Call .Fields.Append(.CreateField("SENDNAME", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("GENNAME", DataTypeEnum.dbText, 20))
        '2012/11/01 M.HONDA ADD
        Call .Fields.Append(.CreateField("YTDATE", DataTypeEnum.dbDate))
        Call .Fields.Append(.CreateField("USAGENAME", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("DLVRT_NO", DataTypeEnum.dbText, 11))       '20150107 M.HONDA INS
        Call .Fields.Append(.CreateField("SHOUMEI_UP_UM", DataTypeEnum.dbText, 20))  '20150203 M.HONDA INS
        Call .Fields.Append(.CreateField("SHOUMEI_KDATE", DataTypeEnum.dbDate)) 'INSERT 2015/07/17 K.ISHIZAKA
        '2015/05/01 M.HONDA INS
        Call .Fields.Append(.CreateField("KISAN", DataTypeEnum.dbDate))
        Call .Fields.Append(.CreateField("SEIKYU", DataTypeEnum.dbText, 30))
        Call .Fields.Append(.CreateField("SKBN", DataTypeEnum.dbText, 30))
        '2015/05/01 M.HONDA INS
        
        Call .Fields.Append(.CreateField("RCPT_SLOPE", DataTypeEnum.dbText, 10))    'INSERT 2022/08/22 N.IMAI
        
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
'2006/06/06 SHIBAZAKI 他画面からのリフレッシュ時に該当データ無しメッセージは表示しない
'Private Function MakeWork() As Boolean
Private Function MakeWork(blnNotFoundMsg As Boolean) As Boolean
On Error GoTo err_rtn
    'FVS540_W01：サブフォームとリンクするテーブル
    
    Dim rsSqlServer     As Recordset
    Dim rsAccess        As Recordset

    MakeWork = False
    
    doCmd.Hourglass True

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    'ワークテーブルの元ネタ取得
    Set rsSqlServer = dbSQLServer.OpenRecordset(MakeYoukSql(), dbOpenSnapshot, dbSQLPassThrough)

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
        '2006/06/06 SHIBAZAKI 他画面からのリフレッシュ時に該当データ無しメッセージは表示しない
        If blnNotFoundMsg Then
            MsgBox "該当データが存在しません。", vbInformation, P_KAKUNIN
        End If
        MakeWork = False
        GoTo Exit_rtn
    End If
    
    '06.1.24 SHIBAZAKI 追加
    '件数表示
    Call DispCount(rsSqlServer.RecordCount)
    
    MakeWork = True
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    
    doCmd.Hourglass False
    
    Exit Function


End Function

'06.1.24 SHIBAZAKI 追加
'==============================================================================*
'
'        MODULE_NAME      :DispCount
'        機能             :件数表示
'        IN               :件数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DispCount(lngCount As Long)

    SubCondition.Form.lbl_Count.Caption = Replace(P_LABEL_件数表示, "@", Format(lngCount, "#,##0"))

End Sub

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

'==============================================================================*
'
'        MODULE_NAME      :MakeYoukSql
'        機能             :予約トラン検索SQL文作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeYoukSql() As String

    Dim strSQL       As String
    Dim strWhere     As String
    Dim strBumonCode As String
    Dim strSqlserver As String
    
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        doCmd.Close acForm, Me.FormName
        Exit Function
    End If
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'"))
    
    strSQL = strSQL & " SELECT YOUKT_UKNO, "
    strSQL = strSQL & "        YOUKT_YCODE, "
    strSQL = strSQL & "        YOUKT_NO, "
    strSQL = strSQL & "        YARD_NAME, "
    strSQL = strSQL & "        YOUKT_SIZE_FROM, "
    strSQL = strSQL & "        YOUKT_SIZE_TO, "
    strSQL = strSQL & "        STEP_NAME.NAME_NAME AS STEPNAME, "
    strSQL = strSQL & "        YOUKT_UKDATE, "
    strSQL = strSQL & "        YOUKT_UCODE, "
    strSQL = strSQL & "        YOUKT_NAME, "
    strSQL = strSQL & "        RENKBN_NAME.NAME_NAME AS RENKBNNAME, "
    '' 2012/11/01 M.HONDA START
    strSQL = strSQL & "        YOUKT_YTDATE, "
    strSQL = strSQL & "        YOUKT_USAGE, "
    strSQL = strSQL & "        USAGE_NAME.NAME_NAME AS USAGENAME, "
    '' 2012/11/01 M.HONDA END
    '連絡区分によって表示する電話番号が変わる
    strSQL = strSQL & "        CASE YOUKT_RENKBN "
    strSQL = strSQL & "            WHEN " & P_連絡区分_電話 & " THEN YOUKT_TEL "
    strSQL = strSQL & "            WHEN " & P_連絡区分_携帯 & " THEN YOUKT_KEITAI "
    strSQL = strSQL & "            WHEN " & P_連絡区分_FAX & " THEN YOUKT_FAX "
    strSQL = strSQL & "        END AS TELNO, "
    strSQL = strSQL & "        YOUKT_BIKO, "
    strSQL = strSQL & "        YOUKT_UKTANTO, "
    strSQL = strSQL & "        YUKBN_NAME.NAME_NAME AS YUKBNNAME, "
    strSQL = strSQL & "        AUTOKBN_NAME.NAME_NAME AS AUTOKBNNAME, "
    strSQL = strSQL & "        YOUKT_TKDATE, "
    '2007/08/23 tajima add
    strSQL = strSQL & "        SEND_NAME.NAME_NAME AS SENDNAME, "
    strSQL = strSQL & "        GENKBN_NAME.NAME_NAME AS GENKBNNAME, "
    '06.1.24 SHIBAZAKI 追加
    strSQL = strSQL & "        YOUKT_YUKBN, "
    strSQL = strSQL & "        YOUKT_MOVEKBN, "
    strSQL = strSQL & "        CASE WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_無し
    strSQL = strSQL & "            THEN NULL "
    strSQL = strSQL & "            ELSE MOVEKBN_NAME.NAME_NAME "
    strSQL = strSQL & "        END AS MOVEKBNNAME, "
    '2006/03/09 SHIBAZAKI 追加
    strSQL = strSQL & "        YOUKT_MOTO_ACPTNO, "
    strSQL = strSQL & "        YOUKT_RCPT_NO, "
    '2006/04/06 SHIBAZAKI 追加
    strSQL = strSQL & "        RCPT_DERIVERY_STATUS, "
    '20150108 M.HONDA INS
    strSQL = strSQL & "        (CASE"
    strSQL = strSQL & "                WHEN DLVRT_NO is null THEN ''"
    'strSQL = strSQL & "                Else '運搬サービスあり' END) AS DLVRT_NO, " 'DELETE 2022/08/22 N.IMAI
    strSQL = strSQL & "                Else '運搬あり' END) AS DLVRT_NO, "          'INSERT 2022/08/22 N.IMAI
    '20150204 M.HONDA INS
    strSQL = strSQL & "        (CASE WHEN RCPT_SHOUMEI_UP_UM = 1  THEN '身分書あり' ELSE '' END) AS SHOUMEI_UP_UM, "
    strSQL = strSQL & "        RCPT_SHOUMEI_KDATE AS SHOUMEI_KDATE, "           'INSERT 2015/07/17 K.ISHIZAKA
    '取置き件数
    strSQL = strSQL & "        ( SELECT COUNT(*) FROM INTR_TRAN "
    strSQL = strSQL & "           WHERE INTRT_UKNO = YOUKT_UKNO AND INTRT_INTROKBN = '" & P_紹介区分_取置き & "' ) AS TKCOUNT, "
    '2015/05/01 M.HONDA INS
    strSQL = strSQL & "        RCPT_KISAN_DATE, "
    strSQL = strSQL & "        RCPT_SEIKYU_CD, "
    strSQL = strSQL & "        SEIKYU_NAME.NAME_NAME AS SEIKYU, "
    strSQL = strSQL & "        TANTM_TANTN, "
    strSQL = strSQL & "        USER_SKBN, "
    strSQL = strSQL & "        SKBN_NAME.NAME_NAME AS SKBN "
    '2015/05/01 M.HONDA INS
    strSQL = strSQL & "       ,CASE WHEN RCPT_SLOPE = 1 THEN '貸出' ELSE '' END AS RCPT_SLOPE"  'INSERT 2022/08/22 N.IMAI
    strSQL = strSQL & "   FROM YOUK_TRAN "
    strSQL = strSQL & "        INNER JOIN YARD_MAST "
    strSQL = strSQL & "           ON YARD_CODE = YOUKT_YCODE "
    
    strSQL = strSQL & "        INNER JOIN NAME_MAST YUKBN_NAME "
    strSQL = strSQL & "           ON YUKBN_NAME.NAME_ID = '080' "
    strSQL = strSQL & "          AND YUKBN_NAME.NAME_CODE = YOUKT_YUKBN "
    
    strSQL = strSQL & "        INNER JOIN NAME_MAST RENKBN_NAME "
    strSQL = strSQL & "           ON RENKBN_NAME.NAME_ID = '085' "
    strSQL = strSQL & "          AND RENKBN_NAME.NAME_CODE = YOUKT_RENKBN "
    
    strSQL = strSQL & "        LEFT  OUTER JOIN NAME_MAST STEP_NAME "
    strSQL = strSQL & "           ON STEP_NAME.NAME_ID = '090' "
    strSQL = strSQL & "          AND STEP_NAME.NAME_CODE = YOUKT_STEP "
    
    strSQL = strSQL & "        LEFT  OUTER JOIN NAME_MAST AUTOKBN_NAME "
    strSQL = strSQL & "           ON AUTOKBN_NAME.NAME_ID = '095' "
    strSQL = strSQL & "          AND AUTOKBN_NAME.NAME_CODE = YOUKT_AUTOKBN "
    
    '06.1.24 SHIBAZAKI 追加
    strSQL = strSQL & "        LEFT  OUTER JOIN NAME_MAST MOVEKBN_NAME "
    strSQL = strSQL & "           ON MOVEKBN_NAME.NAME_ID = '098' "
    strSQL = strSQL & "          AND MOVEKBN_NAME.NAME_CODE = YOUKT_MOVEKBN "
    
    '2006/04/06 SHIBAZAKI 追加
    strSQL = strSQL & "        LEFT  OUTER JOIN RCPT_TRAN "
    strSQL = strSQL & "           ON RCPT_NO = YOUKT_RCPT_NO "
    
    '▼ 2007/08/23 tajima add
    strSQL = strSQL & "        LEFT  OUTER JOIN NAME_MAST SEND_NAME "
    strSQL = strSQL & "           ON SEND_NAME.NAME_ID = '012' "
    strSQL = strSQL & "          AND SEND_NAME.NAME_CODE = YOUKT_SEND_CD "
    
    strSQL = strSQL & "        LEFT  OUTER JOIN NAME_MAST GENKBN_NAME "
    strSQL = strSQL & "           ON GENKBN_NAME.NAME_ID = '083' "
    strSQL = strSQL & "          AND GENKBN_NAME.NAME_CODE = YOUKT_GENKBN "
    '▲ 2007/08/23 tajima add
    
    '' 2012/11/01 M.HONDA START
    strSQL = strSQL & "        LEFT  OUTER JOIN "
    strSQL = strSQL & "            NAME_MAST USAGE_NAME ON "
    strSQL = strSQL & "            USAGE_NAME.NAME_ID = '086' AND "
    strSQL = strSQL & "            USAGE_NAME.NAME_CODE = YOUKT_USAGE "
    '' 2012/11/01 M.HONDA END
    
    '' 2015/01/08 M.HONDA START
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & "            DLVR_TRAN ON "
    strSQL = strSQL & "            RCPT_NO = DLVRT_NO "
    '' 2015/01/08 M.HONDA END
    
    ''2015/05/01 M.HONDA INS
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & strSqlserver & ".dbo.TANT_MAST ON "
    strSQL = strSQL & "        TANTM_BUMOC = '" & strBumonCode & "' AND "
    strSQL = strSQL & "        TANTM_TANTC = YOUKT_UKTANTO "
    strSQL = strSQL & "        Left Join "
    strSQL = strSQL & "        USER_MAST ON "
    strSQL = strSQL & "        USER_CODE = YOUKT_UCODE "
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & "        NAME_MAST SEIKYU_NAME ON "
    strSQL = strSQL & "        SEIKYU_NAME.NAME_ID = '014' AND "
    strSQL = strSQL & "        SEIKYU_NAME.NAME_CODE = RCPT_SEIKYU_CD "
    strSQL = strSQL & "        LEFT OUTER JOIN "
    strSQL = strSQL & "        NAME_MAST SKBN_NAME ON "
    strSQL = strSQL & "        SKBN_NAME.NAME_ID = '030' AND "
    strSQL = strSQL & "        SKBN_NAME.NAME_CODE = USER_SKBN "
    ''2015/05/01 M.HONDA INS
    
    strSQL = strSQL & "  WHERE YOUKT_YUKBN <> " & P_予約状態_契約
    
    strWhere = MakeWhere()
    
    MakeYoukSql = strSQL & strWhere
    
    '' シングルを強制的に置き換える。
    Call MSHS010_M00(PROG_ID, "3", Replace(strWhere, "'", ""))
    
    
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
        strWhere = strWhere & MakeEqual("YOUKT_UKTANTO", Nz(.cmb_Tanto))                            '受付担当者
        strWhere = strWhere & MakeEqual("YOUKT_YCODE", Nz(.cmb_Yard))                               'ヤードコード
        strWhere = strWhere & MakeBetween("YOUKT_YUKBN", Nz(.cmb_YukbnFrom), Nz(.cmb_YukbnTo))      '予約受付区分
        strWhere = strWhere & MakeDateTimeBetween("YOUKT_UKDATE", Nz(.txt_UkdateFrom), Nz(.txt_UkdateTo)) '2007/08/23 fix tajima 予約日は日時を考慮する
        strWhere = strWhere & MakeBetween("YOUKT_KUDATE", Nz(.txt_KudateFrom), Nz(.txt_KudateTo))   '契約受付日
        strWhere = strWhere & MakeBetween("YOUKT_YTDATE", Nz(.txt_YtdateFrom), Nz(.txt_YtdateTo))   '予約取消日
        strWhere = strWhere & MakeBetween("YOUKT_KKDATE", Nz(.txt_KkdateFrom), Nz(.txt_KkdateTo))   '希望開始日
        strWhere = strWhere & MakeBetween("YOUKT_TKDATE", Nz(.txt_TkdateFrom), Nz(.txt_TkdateTo))   '取置き期限日
        strWhere = strWhere & MakeBetween("YOUKT_TADATE", Nz(.txt_TadateFrom), Nz(.txt_TadateTo))   '問合せ日
        
        strWhere = strWhere & MakeBetween("RCPT_PAYDATE", Nz(.txt_PaydateFrom), Nz(.txt_PaydateTo))             '入金確認日     INSERT 2022/08/22 N.IMAI
        strWhere = strWhere & MakeBetween("RCPT_YAMATO_DATE", Nz(.txt_YamatodateFrom), Nz(.txt_YamatodateTo))   'メール配信日   INSERT 2022/08/22 N.IMAI
        
        '2006/04/06 SHIBAZAKI 追加↓
        strWhere = strWhere & MakeBetween("RCPT_KEY_SEND_DATE", Nz(.txt_KeySendDateFrom), Nz(.txt_KeySendDateTo))   '鍵送付日
        strWhere = strWhere & MakeEqual("RCPT_SLIP_NO", Nz(.txt_SlipNo))                            '伝票番号
        '鍵配達完了のみ表示
        If .chkOnlyDeriveryComplete.VALUE = -1 Then
            'strWhere = strWhere & "(" & MakeEqual("RCPT_DERIVERY_STATUS", P_DERIVERYSTATUS_配達完了)
            strWhere = strWhere & " AND ( RCPT_DERIVERY_STATUS = '" & P_DERIVERYSTATUS_配達完了 & "'"
            strWhere = strWhere & " OR RCPT_PAYDATE IS NOT NULL " & " ) "         '2014/12/25 M.HONDA INS
        End If
        '2006/04/06 SHIBAZAKI 追加↑
        
        '身分証明書を確認する必要のあるもののみ表示
        If .chkOnlyShoumeiYouKakunin.VALUE = -1 Then                            'INSERT START 2015/07/27 K.ISHIZAKA
            strWhere = strWhere & " AND RCPT_SHOUMEI_KDATE IS NULL "
        End If                                                                  'INSERT END   2015/07/27 K.ISHIZAKA
        
        '06.1.24 SHIBAZAKI 追加
        strWhere = strWhere & MakeBetween("YOUKT_MOVEKBN", Nz(.cmb_MoveFrom), Nz(.cmb_MoveTo))      '移動理由
        
        '電話番号
        If Nz(.txt_Tel) <> "" Then
            '入力値が電話番号、FAX番号、携帯番号のいずれかに一致する条件
            strWhere = strWhere & " AND "
            strWhere = strWhere & " ( YOUKT_TEL = '" & .txt_Tel & "' OR "
            strWhere = strWhere & "   YOUKT_FAX = '" & .txt_Tel & "' OR "
            strWhere = strWhere & "   YOUKT_KEITAI = '" & .txt_Tel & "' ) "
        End If
        
        '顧客かな／顧客コード
        If Nz(.cmb_Ucode) <> "" Then
            '顧客コードが入力されている場合はそれで一致条件。カナは無視。
            strWhere = strWhere & MakeEqual("YOUKT_UCODE", Nz(.cmb_Ucode))
        ElseIf Nz(.txt_Kana) <> "" Then
            'カナで部分一致
            strWhere = strWhere & " AND YOUKT_KANA LIKE '%" & .txt_Kana & "%' "
        End If
    
        '自動区分
        Dim strTemp     As String
        If Nz(.cmb_AUTO) = P_自動条件_自動のみ Then
            strTemp = " AND ( YOUKT_AUTOKBN IS NOT NULL "
            strTemp = strTemp & " AND YOUKT_AUTOKBN != 0 ) "
        ElseIf Nz(.cmb_AUTO) = P_自動条件_自動除く Then
            strTemp = " AND ( YOUKT_AUTOKBN IS NULL "
            strTemp = strTemp & " OR YOUKT_AUTOKBN = 0 ) "
        End If
        If strTemp <> "" Then
            strWhere = strWhere & strTemp
        End If
        
    '▼ 2007/08/23 tajima add
        '受付種別
        If Nz(.cmb_受付種別) <> "" Then
            strWhere = strWhere & " AND ISNULL(YOUKT_GENKBN,0) = '" & .cmb_受付種別 & "' "
        End If
        
        '申込書送付方
        If Nz(.cmb_申込書送付方) <> "" Then
            strWhere = strWhere & MakeEqual("YOUKT_SEND_CD", Nz(.cmb_申込書送付方))
        End If
        
        '受付番号
        If Nz(.txt_受付番号) <> "" Then
            '前方一致
            strWhere = strWhere & " AND YOUKT_UKNO LIKE '" & .txt_受付番号 & "%' "
        End If
    '▲ 2007/08/23 tajima add
        
        If Nz(.cmbSeikyuCd) <> "" Then
            strWhere = strWhere & " AND RCPT_SEIKYU_CD = '" & .cmbSeikyuCd & "' "
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
Private Sub AddWork(rsSource As Recordset, rsDestination As Recordset)

    With rsDestination
        .AddNew
        
        .Fields("UKNO") = rsSource.Fields("YOUKT_UKNO")
        .Fields("YCODE") = rsSource.Fields("YOUKT_YCODE")
        .Fields("CNO") = rsSource.Fields("YOUKT_NO")
        .Fields("NAME") = rsSource.Fields("YARD_NAME")
        .Fields("SIZE_FROM") = rsSource.Fields("YOUKT_SIZE_FROM")
        .Fields("SIZE_TO") = rsSource.Fields("YOUKT_SIZE_TO")
        .Fields("STEP") = rsSource.Fields("STEPNAME")
        .Fields("UKDATE") = rsSource.Fields("YOUKT_UKDATE")
        .Fields("UCODE") = rsSource.Fields("YOUKT_UCODE")
        .Fields("USER") = rsSource.Fields("YOUKT_NAME")
        .Fields("TEL") = rsSource.Fields("TELNO")
        .Fields("RENKUBUN") = rsSource.Fields("RENKBNNAME")
        .Fields("BIKO") = rsSource.Fields("YOUKT_BIKO")
        .Fields("TANTC") = rsSource.Fields("YOUKT_UKTANTO")

        .Fields("UKEKUBUN") = rsSource.Fields("YUKBNNAME")
        '2006/04/06 SHIBAZAKI 追加
        If Nz(rsSource.Fields("MOVEKBNNAME"), "") <> "" Then
            .Fields("UKEKUBUN") = .Fields("UKEKUBUN") & "(" & rsSource.Fields("MOVEKBNNAME") & ")"
        End If
        'If Nz(rsSource.Fields("RCPT_DERIVERY_STATUS"), "") <> "" Then          'UPDATE 2018/03/27 N.IMAI
        If Nz(rsSource.Fields("RCPT_DERIVERY_STATUS"), "") <> "" _
        And Nz(rsSource.Fields("RCPT_SEIKYU_CD"), 0) <> 1 _
        And Nz(rsSource.Fields("RCPT_SEIKYU_CD"), 0) <> 3 Then
            .Fields("UKEKUBUN") = .Fields("UKEKUBUN") & "/" & rsSource.Fields("RCPT_DERIVERY_STATUS")
        End If
        
        .Fields("AUTOKUBUN") = rsSource.Fields("AUTOKBNNAME")
        .Fields("TKDATE") = rsSource.Fields("YOUKT_TKDATE")
        .Fields("TKCOUNT") = rsSource.Fields("TKCOUNT")
        
        '06.1.24 SHIBAZAKI 追加
        .Fields("UKKBNCODE") = rsSource.Fields("YOUKT_YUKBN")
        .Fields("MOVEKBNCODE") = rsSource.Fields("YOUKT_MOVEKBN")
        .Fields("MOVEKUBUN") = rsSource.Fields("MOVEKBNNAME")
        
        '2006/03/09 SHIBAZAKI 追加
        .Fields("MOTOACPTNO") = rsSource.Fields("YOUKT_MOTO_ACPTNO")
        .Fields("RCPTNO") = rsSource.Fields("YOUKT_RCPT_NO")
        
        '2007/08/23 tajima 追加
        .Fields("SENDNAME") = rsSource.Fields("SENDNAME")
        .Fields("GENNAME") = rsSource.Fields("GENKBNNAME")
        
        '2012/11/01 M.HONDA ADD
        .Fields("YTDATE") = rsSource.Fields("YOUKT_YTDATE")
        .Fields("USAGENAME") = rsSource.Fields("USAGENAME")
        .Fields("DLVRT_NO") = rsSource.Fields("DLVRT_NO")               '20150108 M.HONDA INS
        .Fields("SHOUMEI_UP_UM") = rsSource.Fields("SHOUMEI_UP_UM")     '20150205 M.HONDA INS
        .Fields("SHOUMEI_KDATE") = rsSource.Fields("SHOUMEI_KDATE")             'INSERT 2015/07/17 K.ISHIZAKA
        
        .Fields("KISAN") = rsSource.Fields("RCPT_KISAN_DATE")           '20150501 M.HONDA INS
        .Fields("SEIKYU") = rsSource.Fields("SEIKYU")                   '20150501 M.HONDA INS
        .Fields("SKBN") = rsSource.Fields("SKBN")                       '20150501 M.HONDA INS
        .Fields("TANTN") = rsSource.Fields("TANTM_TANTN")               '20150501 M.HONDA INS
        .Fields("RCPT_SLOPE") = rsSource.Fields("RCPT_SLOPE")           'INSERT 2022/08/22 N.IMAI
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
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
    End Select
    doCmd.Maximize

    Call MSHS010_M00(PROG_ID, "1")

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
        '2006/02/08 SHIBAZAKI 変更
        'Call MakeWork
        '2006/06/06 SHIBAZAKI 他画面からのリフレッシュ時に該当データ無しメッセージは表示しない
'        Call SubFormReload
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
'★★★★★★★★★★★★★★ 2007/08/23 add tajima ★★★★★★★★★★★★★★
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
