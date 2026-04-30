VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FTG011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : コンテナ管理システム
'       SUB_SYSTEM_NAME : 統合検索
'
'       PROGRAM_NAME    : 統合空状況 メイン
'       PROGRAM_ID      : FTG011
'       PROGRAM_KBN     : FROM
'
'       CREATE          : 2011/02/12
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/03/08
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : 価格、割引の結合修正
'
'       UPDATE          : 2011/03/10
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.2
'                       : サイズ指定時に並び順がおかしくなるのを修正
'                       : レンタル用途の関連付け検索のオフ機能が不足していた
'                       : 検索時に詳細設定が開いていたら閉じる
'
'       UPDATE          : 2011/03/16
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.3
'                       : レンタル用途の関連付け検索を削除
'                         サブフォームFTG011_S03を削除
'                       : 用途39:バイク(コンテナ)の対応
'
'       UPDATE          : 2011/07/27
'       UPDATER         : M.HONDA
'       Ver             : 0.4
'                       : 1.貸禁ヤードを表示されないように修正
'                         2.キャンペーンマークが正しく表示されるように修正
'                       : 3.用途バイクで絞った場合に正しく表示されるように修正
'                       : 4.距離の検索が正しく表示されるように修正
'                       : 5.空きがなくても解約があれば予約画面へ進めるように修正
'                       : 6.サブフォームの項目並び順を調整
'                       : 7.必要のない用途を削除
'
'       UPDATE          : 2011/11/24
'       UPDATER         : M.RYU
'       Ver             : 0.5
'                       : SQLエラーを修正
'
'       UPDATE          : 2012/06/16
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.6
'                       : WebDBとの切替(INTI_FILE)に対応させる
'                         MySQLとSQL-Serverにあう書き方にする
'
'       UPDATE          : 2012/10/11
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.7
'                       : 地図表示の高速化対応
'
'       UPDATE          : 2013/05/09
'       UPDATER         : M.HONDA
'       Ver             : 0.7
'                       : 起動LOG対応
'
'       UPDATE          : 2013/10/01
'       UPDATER         : K.KINEBUCHI
'       Ver             : 0.8
'                       : 終了ボタン押下時、メッセージ出力はしない。
'
'       UPDATE          : 2015/02/05
'       UPDATER         : M.HONDA
'       Ver             : 0.9
'                       : 文字コードをセット
'==============================================================================*
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'統合ＤＢ接続部門コード
Private Const C_TOGO_CONNENCT_BUMOC As String = "TOGO"

Private Const C_CURRENT_WORK_TABLE  As String = "WK_FTG011_TEMP"
'Private Const C_SERVER_TEMP_TABLE   As String = "#WK_FTG011_TEMP"              'DELETE 2012/06/16 K.ISHIZAKA

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public strSizeWhere                 As String                                   'INSERT 2011/03/10 K.ISHIZAKA

Private C_SERVER_TEMP_TABLE         As String                                   'INSERT 2012/06/16 K.ISHIZAKA
Private C_SERVER_TEMP_TABLE2        As String                                   'INSERT 2012/06/16 K.ISHIZAKA
Private C_SERVER_TEMP_TABLE3        As String                                   'INSERT 2012/06/16 K.ISHIZAKA
Private blWebDB                     As Boolean                                  'INSERT 2012/06/16 K.ISHIZAKA

'==============================================================================*
'   イベント関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : クリックされたとき
'       MODULE_ID       : chk_NearYard_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_NearYard_Click()
    On Error GoTo ErrorHandler
    txt_Km.Enabled = chk_NearYard.VALUE
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("chk_NearYard_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キーが押されたとき
'       MODULE_ID       : chk_NearYard_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_NearYard_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    
    Select Case KeyAscii
    Case 12288, vbKeySpace
        chk_NearYard.VALUE = Not chk_NearYard.VALUE
    End Select
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("chk_NearYard_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新後処理
'       MODULE_ID       : cmb_Usage_AfterUpdate
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Usage_AfterUpdate()
    On Error GoTo ErrorHandler
    Call ComboBox_AfterUpdate(cmb_Usage, txt_Usage.NAME)
    With cmb_Usage                                                              'INSERT START 2011/03/10 K.ISHIZAKA
        If .ListIndex >= 0 Then
            .VALUE = Format(.Text, .Format)
        End If
    End With                                                                    'INSERT END   2011/03/10 K.ISHIZAKA
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_Usage_AfterUpdate", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : cmb_Usage_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Usage_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    Call CmKeyPress_For_Num1(cmb_Usage, KeyAscii, 2)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_Usage_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : リストにない区分を選択したとき
'       MODULE_ID       : cmb_Usage_NotInList
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Usage_NotInList(NewData As String, Response As Integer)
    On Error GoTo ErrorHandler
    Call ComboBox_NotInList(cmb_Usage, NewData, Response)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_Usage_NotInList", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ヤードコードが選択されたとき
'       MODULE_ID       : cmb_YardCode_AfterUpdate
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardCode_AfterUpdate()
    On Error GoTo ErrorHandler
    
    With cmb_YardCode
        If .ListIndex >= 0 Then
            .VALUE = Format(.Text, .Format)
            opg_Yard.VALUE = opt_Yard_Name.OptionValue
            txt_YardAddr.VALUE = .Column(1, .ListIndex)
            If sub_FTG011_S02.Visible Then                                      'INSERT START 2011/03/10 K.ISHIZAKA
                Call cmd_Syousai_Click
            End If                                                              'INSERT END   2011/03/10 K.ISHIZAKA
            Call GetData(.VALUE, 100)
        End If
    End With
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_YardCode_AfterUpdate", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 終了ボタンが押されたとき
'       MODULE_ID       : cmd_EXIT_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_EXIT_Click()
    On Error GoTo ErrorHandler
    
'    Call CmSysEnd("統合検索システム")
    doCmd.Quit acQuitSaveAll                                                  'INSERT 2013/10/01 K.KINEBUCHI
    
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmd_EXIT_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 地図ボタンが押されたとき
'       MODULE_ID       : cmd_MAP_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_MAP_Click()
    On Error GoTo ErrorHandler
    Call MapOpen
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmd_MAP_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 検索ボタンが押されたとき
'       MODULE_ID       : cmd_Serach_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Serach_Click()
    Dim strMaxList          As String
    On Error GoTo ErrorHandler
    
    If Not inputCheck() Then
        Exit Sub
    End If
    cmb_YardCode.VALUE = ""                                                     'INSERT START 2011/03/10 K.ISHIZAKA
    If sub_FTG011_S02.Visible Then
        Call cmd_Syousai_Click
    End If                                                                      'INSERT END   2011/03/10 K.ISHIZAKA
    strMaxList = Trim(Nz(sub_FTG011_S02.Form.txt_MaxListCount.VALUE))
    Call GetData("", IIf(IsNumeric(strMaxList), CLng(strMaxList), 100))
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmd_Serach_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 詳細ボタン
'       MODULE_ID       : cmd_Syousai_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Syousai_Click()
    On Error GoTo ErrorHandler
    
    With sub_FTG011_S02
        If .Visible Then
            .Visible = False
            cmd_Syousai.Caption = "▼　詳細設定を開く　"
            frmJoken.HEIGHT = .Top - frmJoken.Top
        Else
            .Visible = True
            cmd_Syousai.Caption = "▲　詳細設定を閉じる"
            frmJoken.HEIGHT = .Top + .HEIGHT
            .Form.txt_FuzzyChar.VALUE = getFuzzyChar()
        End If
    End With
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmd_Syousai_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームアクティブ時処理
'       MODULE_ID       : Form_Activate
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    On Error Resume Next
    doCmd.Maximize
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームクローズ時処理
'       MODULE_ID       : Form_Close
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    On Error GoTo ErrorHandler
    Call MSHS010_M00("FTG011", "2")                                             ' 2013/05/09 M.HONDA INS
    Call dropWorkTable
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Form_Close", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームキー押下時処理
'       MODULE_ID       : Form_KeyDown
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    Call KeyDown(KeyCode, Shift)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Form_KeyDown", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームロード
'       MODULE_ID       : Form_Load
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Call MSHS010_M00("FTG011", "1")                                             ' 2013/05/09 M.HONDA INS
    blWebDB = CBool(Nz(DLookup("INTIF_RECFB", "INTI_FILE", "INTIF_PROGB = '" & Me.NAME & "'"), "False"))
    If blWebDB Then
        C_SERVER_TEMP_TABLE = C_CURRENT_WORK_TABLE
        C_SERVER_TEMP_TABLE2 = C_SERVER_TEMP_TABLE & "_2"
        C_SERVER_TEMP_TABLE3 = C_SERVER_TEMP_TABLE & "_3"
    Else
        C_SERVER_TEMP_TABLE = "#" & C_CURRENT_WORK_TABLE
        C_SERVER_TEMP_TABLE2 = C_SERVER_TEMP_TABLE
        C_SERVER_TEMP_TABLE3 = C_SERVER_TEMP_TABLE
    End If
                                                                                'INSERT END   2012/06/16 K.ISHIZAKA
    sub_FTG011_S02.Form.txt_FuzzyChar.VALUE = getFuzzyChar()
    Call setUsageBackColor(cmb_Usage, txt_Usage)                                'INSERT 2011/03/10 K.ISHIZAKA
    Call GetData("", 0)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Form_Load", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : txt_Km_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Km_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    Call CmKeyPress_For_Dec(txt_Km, KeyAscii, 6, 2)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_Km_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : txt_SizeFrom_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SizeFrom_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    Call CmKeyPress_For_Dec(txt_SizeFrom, KeyAscii, 6, 2)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_SizeFrom_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : txt_SizeTo_KeyPress
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyAscii          アスキーコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SizeTo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    Call CmKeyPress_For_Dec(txt_SizeTo, KeyAscii, 6, 2)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_SizeTo_KeyPress", True)
End Sub

'==============================================================================*
'   内部関数（イベントでない）
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'==============================================================================*
'
'       MODULE_NAME     : データ取得
'       MODULE_ID       : getData
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strYcode          ヤードコード(I)
'                       : lngMaxList        最大取得件数(I)
'       RETURN          : 件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetData(ByVal strYcode As String, ByVal lngMaxList As Long) As Long
    Dim objCon              As Object
    Dim objRst              As Object
    Dim objFld              As Object
    Dim rstWk               As Recordset
    Dim strKaseDb           As String
    Dim strUsage            As String                                           'INSERT 2011/03/16 K.ISHIZAKA
    Dim lngCnt              As Long
    Dim strSQL1             As String                                           'INSERT 2012/06/16 K.ISHIZAKA
    Dim strSQL2             As String                                           'INSERT 2012/06/16 K.ISHIZAKA
    Dim strSQL3             As String                                           'INSERT 2012/06/16 K.ISHIZAKA
    Dim strSQL4             As String                                           'INSERT 2012/06/16 K.ISHIZAKA
    Dim strSQL5             As String                                           'INSERT 2012/06/16 K.ISHIZAKA
    Dim strSQL0             As String          'INS 2015/02/05 M.HONDA
    On Error GoTo ErrorHandler

    strKaseDb = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'"), "")
    If strKaseDb = "" Then
        Call MSZZ024_M10("DLookup", "テーブル[SETU_TABL]の設定不正です。")
    End If
                                                                                'DELETE START 2012/06/16 K.ISHIZAKA
'    Set objCon = ADODB_Connection(C_TOGO_CONNENCT_BUMOC)
'    On Error GoTo ErrorHandler1
'    '条件に一致するヤードデータ
'    lngCnt = ADODB_Execute(createTempTable(strKaseDb, strYcode), objCon)
'    '近隣を対象とするとき
'    If chk_NearYard.VALUE <> 0 Then
'        lngCnt = ADODB_Execute(deleteTempTable(), objCon)
'    End If
'    If strYcode = "" Then                                                       'INSERT START 2011/03/16 K.ISHIZAKA
'        strUsage = Nz(cmb_Usage.VALUE)
'    End If                                                                      'INSERT END   2011/03/16 K.ISHIZAKA
''    Set objRst = ADODB_Recordset(selectTempTable(strKaseDb, lngMaxList, strYcode, "", "", ""), objCon) 'DELETE 2011/03/16 K.ISHIZAKA
'    Set objRst = ADODB_Recordset(selectTempTable(strKaseDb, lngMaxList, strYcode, strUsage, "", ""), objCon) 'INSERT 2011/03/16 K.ISHIZAKA
                                                                                'DELETE END   2012/06/16 K.ISHIZAKA
                                                                                
    'INS 2015/02/05 M.HONDA
    '文字コードを強制的にセット
    strSQL0 = "set character_set_client=utf8; set character_set_connection=utf8; set character_set_results=utf8;"
                                                                                'INSERT START 2012/06/16 K.ISHIZAKA
    '条件に一致するヤードデータ
    strSQL1 = createTempTable(strKaseDb, strYcode)
    strSQL2 = copyTempTable(C_SERVER_TEMP_TABLE2)
    '近隣を対象とするとき
    If chk_NearYard.VALUE <> 0 Then
        strSQL3 = copyTempTable(C_SERVER_TEMP_TABLE3)
        strSQL4 = deleteTempTable()
    End If
    If strYcode = "" Then
        strUsage = Nz(cmb_Usage.VALUE)
    End If
    strSQL5 = selectTempTable(strKaseDb, lngMaxList, strYcode, strUsage, "", "")
    
    If blWebDB Then
        Call MySql_ChangeUserLevel(LoginUser)
        Set objRst = MySql_OpenRecordSet(strSQL0, strSQL1, strSQL2, strSQL3, strSQL4, strSQL5)
    Else
        Set objCon = ADODB_Connection(C_TOGO_CONNENCT_BUMOC)
        On Error GoTo ErrorHandler1
        lngCnt = ADODB_Execute(strSQL1, objCon)
        If strSQL4 <> "" Then
            lngCnt = ADODB_Execute(strSQL4, objCon)
        End If
        Set objRst = ADODB_Recordset(strSQL5, objCon)
    End If
                                                                                'INSERT END   2012/06/16 K.ISHIZAKA
    On Error GoTo ErrorHandler2
    Call createWorkTable(objRst)
    Set rstWk = CurrentDb.OpenRecordset(C_CURRENT_WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    On Error GoTo ErrorHandler3
    lngCnt = 0
    While Not objRst.EOF
        rstWk.AddNew
        On Error GoTo ErrorHandler4
        rstWk.Fields("NO").VALUE = lngCnt + 1
        
        For Each objFld In objRst.Fields
            With objFld
                rstWk.Fields(.NAME).VALUE = .VALUE
            End With
        Next
        rstWk.UPDATE
        On Error GoTo ErrorHandler3
        lngCnt = lngCnt + 1
        objRst.MoveNext
    Wend
    rstWk.Close
    On Error GoTo ErrorHandler2
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        objRst.Close_
    Else                                                                        'INSERT END   2012/06/16 K.ISHIZAKA
        objRst.Close
        On Error GoTo ErrorHandler1
        objCon.Close
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
    On Error GoTo ErrorHandler
    sub_FTG011_S01.Form.Requery
    GetData = lngCnt
Exit Function

ErrorHandler4:
    rstWk.CancelUpdate
ErrorHandler3:
    rstWk.Close
ErrorHandler2:
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        objRst.Close_
    Else                                                                        'INSERT END   2012/06/16 K.ISHIZAKA
        objRst.Close
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
ErrorHandler1:
    If Not blWebDB Then                                                         'INSERT 2012/06/16 K.ISHIZAKA
        objCon.Close
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Get_DATA" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブル作成
'       MODULE_ID       : createWorkTable
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : objRst            レコードセット(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub createWorkTable(objRst As Object)
    Dim fldNew              As Field
    Dim tdfNew              As TableDef
    Dim objFld              As Object
    On Error GoTo ErrorHandler

    Call dropWorkTable
    Set tdfNew = CurrentDb().CreateTableDef(C_CURRENT_WORK_TABLE)
    With tdfNew
        .Fields.Append .CreateField("NO", dbLong)
                                                                                'INSERT START 2012/06/16 K.ISHIZAKA
        Call .Fields.Append(.CreateField("YARD_CODE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("YARD_USAGE", DataTypeEnum.dbLong))
'        Call .Fields.Append(.CreateField("CODET_NAMEN", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("CODET_KANAF", DataTypeEnum.dbText, 20))
        Call .Fields.Append(.CreateField("YARD_NAME", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("YARD_ADDR", DataTypeEnum.dbText, 108))
        Call .Fields.Append(.CreateField("YARD_NOTE", DataTypeEnum.dbText, 36))
        Call .Fields.Append(.CreateField("YARD_BIKON", DataTypeEnum.dbText, 60))
        Call .Fields.Append(.CreateField("YARD_IDO", DataTypeEnum.dbDouble))
        Call .Fields.Append(.CreateField("YARD_KEIDO", DataTypeEnum.dbDouble))
        Call .Fields.Append(.CreateField("YARD_BUMOC", DataTypeEnum.dbText, 1))
        Call .Fields.Append(.CreateField("YARD_BEGIN_DAY2", DataTypeEnum.dbText, 7))
        Call .Fields.Append(.CreateField("YARD_RENTEND_DAY2", DataTypeEnum.dbText, 7))
        Call .Fields.Append(.CreateField("SERVICE1", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("MIN_PRICE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("YARD_PRICE", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("NYAR_KIRO", DataTypeEnum.dbDouble))
        Call .Fields.Append(.CreateField("YARD_MIKAKUTEI", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("UP_AKI", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("DW_AKI", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("KAIYAKU", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("TORIOKI", DataTypeEnum.dbLong))
        Call .Fields.Append(.CreateField("SUM_AKI", DataTypeEnum.dbLong))
                                                                                'INSERT END   2012/06/16 K.ISHIZAKA
    End With
'    For Each objFld In objRst.Fields                                           'DELETE START 2012/06/16 K.ISHIZAKA
'        With objFld
'            If .Type = AdoEnums_DataType.adNumeric Then
'                If .NumericScale > 0 Then
'                    Set fldNew = tdfNew.CreateField(.Name, dbDouble)
'                Else
'                    Set fldNew = tdfNew.CreateField(.Name, dbLong)
'                End If
'            ElseIf .Type = AdoEnums_DataType.adDBTimeStamp Then
'                Set fldNew = tdfNew.CreateField(.Name, dbDate)
'            Else
'                Set fldNew = tdfNew.CreateField(.Name, dbText, .DefinedSize)
'                fldNew.AllowZeroLength = True
'            End If
'        End With
'        tdfNew.Fields.Append fldNew
'    Next                                                                       'DELETE END   2012/06/16 K.ISHIZAKA
    For Each objFld In tdfNew.Fields                                            'INSERT START 2012/06/16 K.ISHIZAKA
        If objFld.Type = dbText Then
            objFld.AllowZeroLength = True
        End If
    Next                                                                        'INSERT END   2012/06/16 K.ISHIZAKA
    CurrentDb().TableDefs.Append tdfNew
    sub_FTG011_S01.SourceObject = "FTG011_S01"
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "createWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルの削除
'       MODULE_ID       : dropWorkTable
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub dropWorkTable()
    On Error Resume Next
    sub_FTG011_S01.SourceObject = ""
    CurrentDb.TableDefs.Delete C_CURRENT_WORK_TABLE
    Select Case Err.Number
    Case 0
    Case 3265
        Err.Clear
    Case Else
        Call Err.Raise(Err.Number, "dropWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End Select
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用テーブル作成ＳＱＬ作成
'       MODULE_ID       : createTempTable
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strKaseDb         KASEデータベース名(I)
'                       : strYcode1         ヤードコード１(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createTempTable(ByVal strKaseDb As String, ByVal strYcode1 As String) As String
    Dim strSQL              As String
    Dim strNextSQL          As String
    On Error GoTo ErrorHandler
                                                                                'CHANGE START 2011/03/10 K.ISHIZAKA
                                                                                '見づらいのでコード綺麗にしました
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        strSQL = strSQL & "CREATE TEMPORARY TABLE " & C_SERVER_TEMP_TABLE & " "
    End If                                                                      'INSERT END   2012/06/16 K.ISHIZAKA
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " NYAR_NCODE,"
    strSQL = strSQL & " NYAR_NSAGE,"
    strSQL = strSQL & " MIN(ISNULL(NYAR_KIRO, 0)) AS NYAR_KIRO "
    If Not blWebDB Then                                                         'INSERT 2012/06/16 K.ISHIZAKA
        strSQL = strSQL & "INTO " & C_SERVER_TEMP_TABLE & " "
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " TOGO_NYAR_MAST "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_YARD_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " YARD_CODE = NYAR_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " YARD_USAGE = NYAR_USAGE "
    
    If strYcode1 <> "" Then
        'コンボボックスでヤードコードが選択されたとき
        strSQL = strSQL & "WHERE YARD_CODE = " & strYcode1 & " "
        strNextSQL = "AND"  '' INS 2011/07/27 M.HONDA
    Else
        'ヤード名称、住所、付近メモを指定したとき
        If Trim(Nz(txt_YardAddr.VALUE)) <> "" Then
            Select Case opg_Yard.VALUE
            Case opt_Yard_Name.OptionValue
                If sub_FTG011_S02.Form.chk_Fuzzy.VALUE <> 0 Then
                    strSQL = strSQL & "WHERE YARD_FUZZY_NAME LIKE '%" & changeFuzzyString(Trim(txt_YardAddr.VALUE)) & "%' "
                Else
                    strSQL = strSQL & "WHERE YARD_NAME LIKE '%" & Trim(txt_YardAddr.VALUE) & "%' "
                End If
            Case opt_Yard_Addr.OptionValue
                If sub_FTG011_S02.Form.chk_Fuzzy.VALUE <> 0 Then
                    strSQL = strSQL & "WHERE YARD_FUZZY_ADR1 LIKE '%" & changeFuzzyString(Trim(txt_YardAddr.VALUE)) & "%' "
                Else
                    strSQL = strSQL & "WHERE YARD_ADDR_1 LIKE '%" & Trim(txt_YardAddr.VALUE) & "%' "
                End If
            Case opt_Yard_Note.OptionValue
                strSQL = strSQL & "WHERE YARD_NOTE LIKE '%" & Trim(txt_YardAddr.VALUE) & "%' "
            End Select
            strNextSQL = "AND"
        Else
            strNextSQL = "WHERE"
        End If
    End If
    
    '用途指定したとき
    If Nz(cmb_Usage.VALUE, "") <> "" Then
        Select Case Val(Nz(cmb_Usage.VALUE))
        Case 10
'                If sub_FTG011_S02.Form.chk_UsagMast.VALUE <> 0 Then            'DELETE START 2011/03/16 K.ISHIZAKA
'                    strSQL = strSQL & strNextSQL & " YARD_USAGE IN(0,1,3) "
'                Else                                                           'DELETE END   2011/03/16 K.ISHIZAKA
'               2011/07/27 M.HONDA START
'               strSQL = strSQL & strNextSQL & " NYAR_USAGE IN(0,1,3) "
'               strSQL = strSQL & "AND NYAR_NSAGE IN(0,1,3) "
                strSQL = strSQL & strNextSQL & " NYAR_NSAGE IN(0,1,3) "
'               2011/07/27 M.HONDA END
                
'                End If                                                         'DELETE 2011/03/16 K.ISHIZAKA
        Case 30
'                If sub_FTG011_S02.Form.chk_UsagMast.VALUE <> 0 Then            'DELETE START 2011/03/16 K.ISHIZAKA
'                    strSQL = strSQL & strNextSQL & " YARD_USAGE IN(3,31,32,33) "
'                Else                                                           'DELETE END   2011/03/16 K.ISHIZAKA
'                End If                                                         'DELETE 2011/03/16 K.ISHIZAKA
'               2011/07/27 M.HONDA START
'                strSQL = strSQL & strNextSQL & " NYAR_USAGE IN(3,31,32,33) "
'                strSQL = strSQL & "AND NYAR_NSAGE IN(3,31,32,33) "
                strSQL = strSQL & strNextSQL & " NYAR_NSAGE IN(3,31,32,33) "
'               2011/07/27 M.HONDA END

        Case Else
'                If sub_FTG011_S02.Form.chk_UsagMast.VALUE <> 0 Then            'DELETE START 2011/03/16 K.ISHIZAKA
'                    strSQL = strSQL & strNextSQL & " YARD_USAGE IN("
'                    strSQL = strSQL & "SELECT USAG_NSAGE FROM TOGO_USAG_MAST "
'                    strSQL = strSQL & "WHERE USAG_USAGE = " & cmb_Usage.VALUE & " "
'                    strSQL = strSQL & ") "
'                Else                                                           'DELETE END   2011/03/16 K.ISHIZAKA
'               2011/07/27 M.HONDA START
'                strSQL = strSQL & strNextSQL & " YARD_USAGE = " & cmb_Usage.VALUE & " "
'                strSQL = strSQL & "AND NYAR_NSAGE = NYAR_USAGE "
                strSQL = strSQL & strNextSQL & " NYAR_NSAGE = " & cmb_Usage.VALUE & " "
'               2011/07/27 M.HONDA END
'                End If                                                         'DELETE 2011/03/16 K.ISHIZAKA
        End Select
    End If
    
    If chk_NearYard.VALUE = 0 Then
        strSQL = strSQL & "AND NYAR_YCODE = NYAR_NCODE "    '近隣を含めない
    '' 2011/07/27 M.HONDA START
    Else
        If txt_Km.VALUE <> "" Then
            strSQL = strSQL & "AND ISNULL(NYAR_KIRO, 0) <= " & txt_Km.VALUE
        End If
    '' 2011/07/27 M.HONDA END
    End If
    
    strSQL = strSQL & " GROUP BY"
    strSQL = strSQL & " NYAR_NCODE,"
    strSQL = strSQL & " NYAR_NSAGE "
                                                                                'CHANGE END   2011/03/10 K.ISHIZAKA
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        strSQL = Replace(strSQL, "ISNULL(", "IFNULL(", 1, -1, vbTextCompare)
    End If                                                                      'INSERT END   2012/06/16 K.ISHIZAKA
    createTempTable = strSQL
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "createTempTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用テーブル削除ＳＱＬ作成
'       MODULE_ID       : deleteTempTable
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function deleteTempTable() As String
    Dim strSQL              As String

    strSQL = strSQL & "DELETE"
    strSQL = strSQL & " FROM"                                                   'INSERT 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & " " & C_SERVER_TEMP_TABLE & " "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " EXISTS "
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
'    strSQL = strSQL & " * "                                                    'DELETE 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & " 1 "                                                     'INSERT 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "  DISTINCT"
    strSQL = strSQL & "  wk1.NYAR_NCODE AS WORK_YCODE,"
    strSQL = strSQL & "  wk2.NYAR_NCODE AS WORK_NCODE"
    strSQL = strSQL & " FROM"
'    strSQL = strSQL & "  " & C_SERVER_TEMP_TABLE & " wk1,"                     'DELETE 2012/06/16 K.ISHIZAKA
'    strSQL = strSQL & "  " & C_SERVER_TEMP_TABLE & " wk2"                      'DELETE 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & "  " & C_SERVER_TEMP_TABLE2 & " wk1,"                     'INSERT 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & "  " & C_SERVER_TEMP_TABLE3 & " wk2"                      'INSERT 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  wk1.NYAR_KIRO = wk2.NYAR_KIRO"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  wk1.NYAR_NCODE != wk2.NYAR_NCODE"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  wk1.NYAR_KIRO = 0"
    strSQL = strSQL & ") wk "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " TOGO_NYAR_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " NYAR_YCODE = WORK_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " NYAR_NCODE = WORK_NCODE "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " ISNULL(NYAR_KIRO,999.9999) >= 1.0 "
    strSQL = strSQL & ")"
    strSQL = strSQL & "AND"
    strSQL = strSQL & " NYAR_KIRO > 0.0 "

    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        strSQL = Replace(strSQL, "ISNULL(", "IFNULL(", 1, -1, vbTextCompare)
    End If                                                                      'INSERT END   2012/06/16 K.ISHIZAKA
    deleteTempTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用テーブル取得ＳＱＬ作成
'       MODULE_ID       : selectTempTable
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strKaseDb         KASEデータベース名(I)
'                       : lngMaxList        最大取得件数(I)
'                       : strYcode1         ヤードコード１(I)
'                       : strUsage1         用途１(I)
'                       : strYcode2         ヤードコード２(I)
'                       : strUsage2         用途２(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function selectTempTable(ByVal strKaseDb As String, ByVal lngMaxList As Long, ByVal strYcode1 As String, ByVal strUsage1 As String, ByVal strYcode2 As String, ByVal strUsage2 As String) As String
    Dim strSQL              As String
    Dim strWhere            As String       'INSERT 2011/11/24 M.RYU
    
    strSQL = ""                             'INSERT 2011/11/24 M.RYU
    strWhere = ""                           'INSERT 2011/11/24 M.RYU
    
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        strSQL = "SELECT"
    Else                                                                        'INSERT END   2012/06/16 K.ISHIZAKA
        strSQL = "SELECT TOP " & Format(lngMaxList)
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
    strSQL = strSQL & " YARD_CODE,"
    If (strUsage1 = "10") Or (strUsage1 = "39") Or (strUsage2 = "39") Then      'INSERT START 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " CASE YARD_USAGE WHEN 3 THEN 39 ELSE YARD_USAGE END AS"
    End If                                                                      'INSERT END   2011/03/16 K.ISHIZAKA
    strSQL = strSQL & " YARD_USAGE,"
    
    If strKaseDb <> "" Then
        '検索ボタンのとき、と、コンボボックスでヤードコードが選択されたとき
        '2011/07/27 M.HONDA START
        'strSQL = strSQL & " CODET_NAMEN,"
        strSQL = strSQL & " CODET_KANAF,"
        '2011/07/27 M.HONDA END
        strSQL = strSQL & " YARD_NAME,"
        If blWebDB Then                                                         'INSERT START 2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " CONCAT(YARD_ADDR_1, ISNULL(YARD_ADDR_2,''), ISNULL(YARD_ADDR_3,'')) AS YARD_ADDR,"
        Else                                                                    'INSERT END   2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " CAST(YARD_ADDR_1 + ISNULL(YARD_ADDR_2,'') + ISNULL(YARD_ADDR_3,'') AS VARCHAR(108)) AS YARD_ADDR,"
        End If                                                                  'INSERT 2012/06/16 K.ISHIZAKA
        strSQL = strSQL & " YARD_NOTE,"
        strSQL = strSQL & " YARD_BIKON,"
        strSQL = strSQL & " YARD_IDO,"
        strSQL = strSQL & " YARD_KEIDO,"
        strSQL = strSQL & " YARD_BUMOC,"
    
        If blWebDB Then                                                         'INSERT START 2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " CONCAT(DATE_FORMAT(YARD_BEGIN_DAY, '%y/%m'), '～') AS YARD_BEGIN_DAY2,"
            strSQL = strSQL & " CONCAT('～', DATE_FORMAT(YARD_RENTEND_DAY, '%y/%m')) AS YARD_RENTEND_DAY2,"
            strSQL = strSQL & " CAST(CASE WHEN COUNT(CASE WHEN SPAC_AKI = 0 THEN NULL ELSE RLDNT_NO END) = 0 THEN 0 ELSE 1 END AS SIGNED) AS SERVICE1,"
        Else                                                                    'INSERT END   2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " CAST(CONVERT(varchar(5), YARD_BEGIN_DAY, 11) + '～' AS VARCHAR(7)) AS YARD_BEGIN_DAY2,"
            strSQL = strSQL & " CAST('～' + CONVERT(varchar(5), YARD_RENTEND_DAY, 11) AS VARCHAR(7)) AS YARD_RENTEND_DAY2,"
            strSQL = strSQL & " CAST(CASE WHEN COUNT(CASE WHEN SPAC_AKI = 0 THEN NULL ELSE RLDNT_NO END) = 0 THEN 0 ELSE 1 END AS NUMERIC(1)) AS SERVICE1,"
        End If                                                                  'INSERT 2012/06/16 K.ISHIZAKA

        strSQL = strSQL & " CAST(MIN(PRIC_PRICE) AS NUMERIC(9)) AS MIN_PRICE,"
        strSQL = strSQL & " CAST("
        strSQL = strSQL & " ("
        strSQL = strSQL & "  SELECT"
        strSQL = strSQL & "   ISNULL(MIN(PRIC_PRICE),999999999)"
        strSQL = strSQL & "  FROM"
'        strSQL = strSQL & " " & C_SERVER_TEMP_TABLE                            'DELETE 2012/06/16 K.ISHIZAKA 'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & " " & C_SERVER_TEMP_TABLE2                            'INSERT 2012/06/16 K.ISHIZAKA
        strSQL = strSQL & "  INNER JOIN"                                        'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "   TOGO_SPAC_TABL"
        strSQL = strSQL & "  ON"                                                'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "   SPAC_YCODE = NYAR_NCODE"                          'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "  AND"                                               'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "   SPAC_USAGE = NYAR_NSAGE"                          'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "  INNER JOIN"
        strSQL = strSQL & "   TOGO_PRIC_TABL"
        strSQL = strSQL & "  ON"
        strSQL = strSQL & "   PRIC_YCODE = SPAC_YCODE"
        strSQL = strSQL & "  AND"
        strSQL = strSQL & "   PRIC_USAGE = SPAC_USAGE"
        strSQL = strSQL & "  AND"
        strSQL = strSQL & "   PRIC_SIZE = SPAC_SIZE"
        strSQL = strSQL & "  AND"
        strSQL = strSQL & "   PRIC_STEP = SPAC_STEP"
        strSQL = strSQL & "  AND"
        strSQL = strSQL & "   SPAC_AKI > 0"
        strSQL = strSQL & "  WHERE"
        strSQL = strSQL & "   PRIC_YCODE = YARD_CODE"
        
        strSizeWhere = ""                                                       'INSERT START 2011/03/10 K.ISHIZAKA
        If strYcode1 = "" Then
            'サイズ指定したとき
            If (Trim(Nz(txt_SizeFrom.VALUE)) <> "") And (Trim(Nz(txt_SizeTo.VALUE)) <> "") Then
                If Trim(Nz(txt_SizeFrom.VALUE)) = Trim(Nz(txt_SizeTo.VALUE)) Then
                    strSizeWhere = "= " & txt_SizeFrom.VALUE & " "
                Else
                    strSizeWhere = "BETWEEN " & txt_SizeFrom.VALUE & " AND " & txt_SizeTo.VALUE & " "
                End If
            ElseIf Trim(Nz(txt_SizeFrom.VALUE)) <> "" Then
                strSizeWhere = ">= " & txt_SizeFrom.VALUE & " "
            ElseIf Trim(Nz(txt_SizeTo.VALUE)) <> "" Then
                strSizeWhere = "<= " & txt_SizeTo.VALUE & " "
            End If
            If strSizeWhere <> "" Then
                strSQL = strSQL & " AND SPAC_SIZE " & strSizeWhere
            End If
        End If                                                                  'INSERT END   2011/03/10 K.ISHIZAKA
        
        strSQL = strSQL & " ) AS NUMERIC(9)) AS YARD_PRICE,"
        strSQL = strSQL & " MIN(NYAR_KIRO) AS NYAR_KIRO,"
    End If
    
    'strSQL = strSQL & " CAST(SUM(YARD_MIKAKUTEI) AS NUMERIC(9)) AS YARD_MIKAKUTEI,"
    strSQL = strSQL & " MAX(YARD_MIKAKUTEI) AS YARD_MIKAKUTEI,"
    strSQL = strSQL & " CAST(SUM(CASE WHEN SPAC_STEP = 0 THEN SPAC_AKI ELSE 0 END) AS NUMERIC(9)) AS UP_AKI,"
    strSQL = strSQL & " CAST(SUM(CASE WHEN SPAC_STEP = 1 THEN SPAC_AKI ELSE 0 END) AS NUMERIC(9)) AS DW_AKI,"
    strSQL = strSQL & " CAST(SUM(SPAC_KAI) AS NUMERIC(9)) AS KAIYAKU,"
    strSQL = strSQL & " CAST(SUM(SPAC_TORI) AS NUMERIC(9)) AS TORIOKI,"
    strSQL = strSQL & " CAST(SUM(SPAC_AKI) AS NUMERIC(9)) AS SUM_AKI "

    strSQL = strSQL & "FROM"
    If strKaseDb <> "" Then
        strSQL = strSQL & " " & C_SERVER_TEMP_TABLE & " "
        strSQL = strSQL & "INNER JOIN"
    End If
    strSQL = strSQL & " TOGO_YARD_MAST "
    If strKaseDb <> "" Then
        strSQL = strSQL & "ON"
        strSQL = strSQL & " YARD_CODE = NYAR_NCODE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " YARD_USAGE = NYAR_NSAGE "
        '' 2011/07/27 M.HONDA START
        strSQL = strSQL & "AND"
        If blWebDB Then                                                         'INSERT START 2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " ISNULL(YARD_INLIMIT_DAY,  '1900/01/01') < DATE_SUB(CURDATE(), INTERVAL 1 DAY)  "
        Else                                                                    'INSERT END   2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " ISNULL(YARD_INLIMIT_DAY,  '1900/01/01') < GETDATE() -1  "
        End If                                                                  'INSERT 2012/06/16 K.ISHIZAKA
        '' 2011/07/27 M.HONDA END
    End If
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_SPAC_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " SPAC_YCODE = YARD_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_USAGE = YARD_USAGE "
    
    If strKaseDb <> "" Then
'        strSQL = strSQL & "INNER JOIN"                                         'DELETE 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "LEFT JOIN"                                           'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & " TOGO_PRIC_TABL "
        strSQL = strSQL & "ON"
        strSQL = strSQL & " PRIC_YCODE = SPAC_YCODE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC_USAGE = SPAC_USAGE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC_SIZE = SPAC_SIZE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC_STEP = SPAC_STEP "
        strSQL = strSQL & "AND"                                                 'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & " SPAC_AKI > 0 "                                      'INSERT 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "LEFT JOIN"
        
        strSQL = strSQL & "("                                                   'INSERT START 2011/03/08 K.ISHIZAKA
        strSQL = strSQL & "SELECT"
        strSQL = strSQL & " RLDNT_YCODE,"
        strSQL = strSQL & " RLDNT_USAGE,"
        strSQL = strSQL & " RLDNT_SIZE,"
        strSQL = strSQL & " RLDNT_STEP,"
        strSQL = strSQL & " 1 AS RLDNT_NO "
        strSQL = strSQL & "FROM"                                                'INSERT END   2011/03/08 K.ISHIZAKA
        strSQL = strSQL & " TOGO_RLDN_TRAN "
        strSQL = strSQL & "WHERE"                                               'INSERT START 2011/03/08 K.ISHIZAKA
        If blWebDB Then                                                         'INSERT START 2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " DATE_FORMAT(NOW(), '%Y%m%d') BETWEEN RLDNT_FROM AND RLDNT_TO "
        Else                                                                    'INSERT END   2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " CONVERT(varchar(8), GETDATE(), 112) BETWEEN RLDNT_FROM AND RLDNT_TO "
        End If                                                                  'INSERT 2012/06/16 K.ISHIZAKA
        strSQL = strSQL & " AND RLDNT_ENABLE = 1 "                              'INSERT 2011/07/27 M.HONDA
        strSQL = strSQL & "GROUP BY"
        strSQL = strSQL & " RLDNT_YCODE,"
        strSQL = strSQL & " RLDNT_USAGE,"
        strSQL = strSQL & " RLDNT_SIZE,"
        strSQL = strSQL & " RLDNT_STEP"
        strSQL = strSQL & ") RLDN_TRAN "                                        'INSERT END   2011/03/08 K.ISHIZAKA
        
        strSQL = strSQL & "ON"
        strSQL = strSQL & " RLDNT_YCODE = SPAC_YCODE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " RLDNT_USAGE = SPAC_USAGE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " RLDNT_SIZE = SPAC_SIZE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " RLDNT_STEP = SPAC_STEP "
        strSQL = strSQL & "LEFT JOIN"
        If blWebDB Then                                                         'INSERT START 2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " CODE_TABL "
        Else                                                                    'INSERT END   2012/06/16 K.ISHIZAKA
            strSQL = strSQL & " " & strKaseDb & ".dbo.CODE_TABL "
        End If                                                                  'INSERT 2012/06/16 K.ISHIZAKA
        strSQL = strSQL & "ON"
        strSQL = strSQL & " CODET_SIKBC = '086' "
        strSQL = strSQL & "AND"
        If strUsage1 = "10" Then                                                'INSERT START 2011/03/16 K.ISHIZAKA
            strSQL = strSQL & " CODET_CODEC = CASE YARD_USAGE WHEN 3 THEN 39 ELSE YARD_USAGE END "
        Else                                                                    'INSERT END   2011/03/16 K.ISHIZAKA
            strSQL = strSQL & " CODET_CODEC = YARD_USAGE "
        End If                                                                  'INSERT 2011/03/16 K.ISHIZAKA

        If strYcode1 = "" Then
            'サイズ指定したとき
'                                                                               'DELETE START 2011/03/10 K.ISHIZAKA
'            If (Trim(Nz(txt_SizeFrom.VALUE)) <> "") And (Trim(Nz(txt_SizeTo.VALUE)) <> "") Then
'                If Trim(Nz(txt_SizeFrom.VALUE)) = Trim(Nz(txt_SizeTo.VALUE)) Then
'                    strSQL = strSQL & "WHERE SPAC_SIZE = " & txt_SizeFrom.VALUE & " "
'                Else
'                    strSQL = strSQL & "WHERE SPAC_SIZE BETWEEN " & txt_SizeFrom.VALUE & " AND " & txt_SizeTo.VALUE & " "
'                End If
'            ElseIf Trim(Nz(txt_SizeFrom.VALUE)) <> "" Then
'                strSQL = strSQL & "WHERE SPAC_SIZE >= " & txt_SizeFrom.VALUE & " "
'            ElseIf Trim(Nz(txt_SizeTo.VALUE)) <> "" Then
'                strSQL = strSQL & "WHERE SPAC_SIZE <= " & txt_SizeTo.VALUE & " "
'            End If                                                             'DELETE END   2011/03/10 K.ISHIZAKA
            If strSizeWhere <> "" Then                                          'INSERT START 2011/03/10 K.ISHIZAKA
                strSQL = strSQL & "WHERE SPAC_SIZE " & strSizeWhere
            End If                                                              'INSERT END   2011/03/10 K.ISHIZAKA
        End If
    Else
    
        '---↓↓↓↓----------UPDATE 2011/11/24 M.RYU----------↓↓↓↓-----------<S>
'        strSQL = strSQL & "WHERE (YARD_CODE = " & strYcode1 & " "
'        If strUsage1 <> "" Then
'            If strUsage1 = "39" Then                                            'INSERT START 2011/03/16 K.ISHIZAKA
'                strUsage1 = "3"
'            End If                                                              'INSERT END   2011/03/16 K.ISHIZAKA
'            strSQL = strSQL & " AND YARD_USAGE = " & strUsage1 & " "
'        End If
'        strSQL = strSQL & ")"
'        If strYcode2 <> "" Then
'            strSQL = strSQL & "OR (YARD_CODE = " & strYcode2 & " "
'            If strUsage2 <> "" Then
'                If strUsage2 = "39" Then                                        'INSERT START 2011/03/16 K.ISHIZAKA
'                    strUsage2 = "3"
'                End If                                                          'INSERT END   2011/03/16 K.ISHIZAKA
'                strSQL = strSQL & " AND YARD_USAGE = " & strUsage2 & " "
'            End If
'            strSQL = strSQL & " ) "
'        End If
        
        strWhere = ""
        If Trim(Nz(strYcode1)) <> "" Then
            strWhere = strWhere & " (YARD_CODE = " & strYcode1 & " "
            If strUsage1 <> "" Then
                If strUsage1 = "39" Then strUsage1 = "3"
                strWhere = strWhere & " AND YARD_USAGE = " & strUsage1 & " "
            End If
            strWhere = strWhere & ")"
            
            If strYcode2 <> "" Then strWhere = strWhere & " OR "
            
        End If
        
        If strYcode2 <> "" Then
            strWhere = strWhere & " (YARD_CODE = " & strYcode2 & " "
            If strUsage2 <> "" Then
                If strUsage2 = "39" Then strUsage2 = "3"
                strWhere = strWhere & " AND YARD_USAGE = " & strUsage2 & " "
            End If
            strWhere = strWhere & " ) "
            
        End If
        
        If Trim(strWhere) <> "" Then strSQL = strSQL & " WHERE " & strWhere
        
        '---↑↑↑↑----------UPDATE 2011/11/24 M.RYU----------↑↑↑↑-----------<E>
                
    End If
    
    
    strSQL = strSQL & " GROUP BY"
    strSQL = strSQL & "  YARD_CODE,"
    strSQL = strSQL & "  YARD_USAGE "
    If strKaseDb <> "" Then
        strSQL = strSQL & ","
        '2011/07/27 M.HONDA START
        'strSQL = strSQL & " CODET_NAMEN,"
        strSQL = strSQL & " CODET_KANAF,"
        '2011/07/27 M.HONDA END
        strSQL = strSQL & " YARD_NAME,"
        strSQL = strSQL & " YARD_ADDR_1,YARD_ADDR_2,YARD_ADDR_3,"
        strSQL = strSQL & " YARD_NOTE,"
        strSQL = strSQL & " YARD_BIKON,"
        strSQL = strSQL & " YARD_IDO,"
        strSQL = strSQL & " YARD_KEIDO,"
        strSQL = strSQL & " YARD_BUMOC,"
        strSQL = strSQL & " YARD_BEGIN_DAY,"
        strSQL = strSQL & " YARD_RENTEND_DAY "
        
        strSQL = strSQL & "HAVING"
        strSQL = strSQL & " MIN(NYAR_KIRO) = 0 "
        strSQL = strSQL & "OR"
        strSQL = strSQL & " MIN(PRIC_PRICE) IS NOT NULL "
        strSQL = strSQL & "ORDER BY"
        strSQL = strSQL & " CASE WHEN MIN(NYAR_KIRO) = 0 THEN 0 ELSE 1 END,"
        If sub_FTG011_S02.Form.chk_Sort_Service.VALUE <> 0 Then
            strSQL = strSQL & " SERVICE1 DESC," 'キャンペーン優先
        End If
        If opg_Sort.VALUE = opt_Sort_Km.OptionValue Then
            strSQL = strSQL & " ISNULL(MIN(NYAR_KIRO),999.9999),"
        End If
        If sub_FTG011_S02.Form.chk_Sort_YardCode.VALUE <> 0 Then
            strSQL = strSQL & " YARD_PRICE,"
            strSQL = strSQL & " YARD_CODE,"
            strSQL = strSQL & " ISNULL(MIN(PRIC_PRICE),999999999),"
            strSQL = strSQL & " YARD_USAGE,"
        Else
            strSQL = strSQL & " ISNULL(MIN(PRIC_PRICE),999999999),"
            strSQL = strSQL & " YARD_USAGE,"
            strSQL = strSQL & " YARD_CODE,"
        End If
        strSQL = strSQL & " SUM_AKI DESC "
    Else
        strSQL = strSQL & "ORDER BY"
        strSQL = strSQL & " YARD_CODE,"
        strSQL = strSQL & " YARD_USAGE "
    End If

    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        strSQL = Replace(strSQL, "ISNULL(", "IFNULL(", 1, -1, vbTextCompare)
        strSQL = Replace(strSQL, "NUMERIC(9)", "SIGNED", 1, -1, vbTextCompare)
        strSQL = strSQL & " LIMIT " & Format(lngMaxList)
    End If                                                                      'INSERT END   2012/06/16 K.ISHIZAKA
    selectTempTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 入力チェック
'       MODULE_ID       : inputCheck
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function inputCheck() As Boolean
    On Error GoTo ErrorHandler
    
    inputCheck = False
    If Not checkIsNumeric(txt_Km, 2) Then
        Exit Function
    End If
    If Not inputCheck_CNTA_SIZE(txt_SizeFrom) Then
        Exit Function
    End If
    If Not inputCheck_CNTA_SIZE(txt_SizeTo) Then
        Exit Function
    End If
    inputCheck = True
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "inputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナサイズ入力チェック
'       MODULE_ID       : inputCheck_CNTA_SIZE
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : ctrl              コンテナサイズ
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function inputCheck_CNTA_SIZE(ctrl As TextBox) As Boolean
    On Error GoTo ErrorHandler
    
    If Not checkIsNumeric(txt_SizeTo, 6, ".", 2) Then
        inputCheck_CNTA_SIZE = False
        Exit Function
    End If
    If Trim(Nz(ctrl.VALUE)) <> "" Then
        If (Val(ctrl.VALUE) <= 0) Or (Val(ctrl.VALUE) >= 1000) Then
            inputCheck_CNTA_SIZE = checkMsgBox(ctrl, "0.01～999.99の範囲で入力してください。")
            Exit Function
        End If
    End If
    inputCheck_CNTA_SIZE = True
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "inputCheck_CNTA_SIZE" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルデータ取得ＳＱＬ作成
'       MODULE_ID       : selectWorkTable
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strYcode1         ヤードコード１(I)
'                       : strUsage1         用途１(I)
'                       : strYcode2         ヤードコード２(I)
'                       : strUsage2         用途２(I)
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function selectWorkTable(ByVal strYcode1 As String, ByVal strUsage1 As String, ByVal strYcode2 As String, ByVal strUsage2 As String) As String
    Dim strSQL              As String
    
    strSQL = "SELECT * FROM " & C_CURRENT_WORK_TABLE & " "
    strSQL = strSQL & "WHERE (YARD_CODE = " & strYcode1 & " "
    If strUsage1 <> "" Then
        strSQL = strSQL & " AND YARD_USAGE = " & strUsage1 & " "
    End If
    strSQL = strSQL & ")"
    If strYcode2 <> "" Then
        strSQL = strSQL & "OR (YARD_CODE = " & strYcode2 & " "
        If strUsage2 <> "" Then
            strSQL = strSQL & " AND YARD_USAGE = " & strUsage2 & " "
        End If
        strSQL = strSQL & ")"
    End If
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " YARD_USAGE "
    
    selectWorkTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブル更新
'       MODULE_ID       : updateWorkTable
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : rstWk             レコードセット(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub updateWorkTable(rstWk As Recordset)
    On Error GoTo ErrorHandler

    With rstWk
        .Edit
        On Error GoTo ErrorHandler1
        .Fields("YARD_MIKAKUTEI").VALUE = 0
        .Fields("UP_AKI").VALUE = 0
        .Fields("DW_AKI").VALUE = 0
        .Fields("KAIYAKU").VALUE = 0
        .Fields("TORIOKI").VALUE = 0
        .Fields("SUM_AKI").VALUE = 0
        .UPDATE
        On Error GoTo ErrorHandler
    End With
Exit Sub

ErrorHandler1:
    rstWk.CancelUpdate
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "updateWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'   外部関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : キー押下処理
'       MODULE_ID       : KeyDown
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        cmd_EXIT.SetFocus
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 一覧データ検索処理
'       MODULE_ID       : Reload
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strYcode1         ヤードコード１(I)
'                       : strUsage1         用途１(I)
'                       : strYcode2         ヤードコード２(I)
'                       : strUsage2         用途２(I)
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Reload(ByVal strYcode1 As String, ByVal strUsage1 As String, ByVal strYcode2 As String, ByVal strUsage2 As String) As Boolean
    Dim objCon              As Object
    Dim objRst              As Object
    Dim objFld              As Object
    Dim rstWk               As Recordset
    Dim lngCount            As Long
    Dim lngSelTop           As Long
    Dim lngPageTop          As Long
    On Error GoTo ErrorHandler

    With sub_FTG011_S01.Form
        lngCount = .Count
        lngSelTop = .SelTop
        lngPageTop = .SelTop - CInt(.CurrentSectionTop / .詳細.HEIGHT)
    End With
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        Set objRst = MySql_OpenRecordSet(selectTempTable("", 999, strYcode1, strUsage1, strYcode2, strUsage2))
    Else                                                                        'INSERT END   2012/06/16 K.ISHIZAKA
        Set objCon = ADODB_Connection(C_TOGO_CONNENCT_BUMOC)
        On Error GoTo ErrorHandler1
        Set objRst = ADODB_Recordset(selectTempTable("", 999, strYcode1, strUsage1, strYcode2, strUsage2), objCon)
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
    On Error GoTo ErrorHandler2
    Set rstWk = CurrentDb.OpenRecordset(selectWorkTable(strYcode1, strUsage1, strYcode2, strUsage2), dbOpenDynaset)
    On Error GoTo ErrorHandler3
    While (Not objRst.EOF) And (Not rstWk.EOF)
        Select Case StrComp(rstWk.Fields("YARD_CODE").VALUE, objRst.Fields("YARD_CODE").VALUE)
        Case -1
            Call updateWorkTable(rstWk)
            rstWk.MoveNext
        Case 0
            Select Case StrComp(rstWk.Fields("YARD_USAGE").VALUE, objRst.Fields("YARD_USAGE").VALUE)
            Case -1
                Call updateWorkTable(rstWk)
                rstWk.MoveNext
            Case 0
                rstWk.Edit
                On Error GoTo ErrorHandler4
                For Each objFld In objRst.Fields
                    With objFld
                        rstWk.Fields(.NAME).VALUE = .VALUE
                    End With
                Next
                rstWk.UPDATE
                On Error GoTo ErrorHandler3
                rstWk.MoveNext
                objRst.MoveNext
            Case 1
                objRst.MoveNext
            End Select
        Case 1
            objRst.MoveNext
        End Select
    Wend
    While Not rstWk.EOF
        Call updateWorkTable(rstWk)
        rstWk.MoveNext
    Wend
    rstWk.Close
    On Error GoTo ErrorHandler2
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        objRst.Close_
    Else                                                                        'INSERT END   2012/06/16 K.ISHIZAKA
        objRst.Close
        On Error GoTo ErrorHandler1
        objCon.Close
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
    On Error GoTo ErrorHandler
    With sub_FTG011_S01.Form
        .Requery
        If .Count > 1 Then
            Call SetTopRow(sub_FTG011_S01.Form, lngPageTop + 1)
            If lngCount = .Count Then
                .SelTop = lngSelTop
            Else
                .SelTop = lngSelTop - 1
            End If
        End If
    End With
    Reload = True
Exit Function

ErrorHandler4:
    rstWk.CancelUpdate
ErrorHandler3:
    rstWk.Close
ErrorHandler2:
    If blWebDB Then                                                             'INSERT START 2012/06/16 K.ISHIZAKA
        objRst.Close_
    Else                                                                        'INSERT END   2012/06/16 K.ISHIZAKA
        objRst.Close
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
ErrorHandler1:
    If Not blWebDB Then                                                         'INSERT 2012/06/16 K.ISHIZAKA
        objCon.Close
    End If                                                                      'INSERT 2012/06/16 K.ISHIZAKA
ErrorHandler:
    Call MSZZ024_M00("Reload", False)
    Reload = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : 地図表示
'       MODULE_ID       : MapOpen
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MapOpen()
    Dim rsAccess            As Recordset
    Dim IEApp               As Object
    Dim Elmnt               As Object
    Dim strMark             As String
    On Error GoTo Err_Handler

'    Set rsAccess = CurrentDb().OpenRecordset("SELECT * FROM " & C_CURRENT_WORK_TABLE & " ORDER BY NO DESC", dbOpenForwardOnly) 'DELETE 2012/10/11 K.ISHIZAKA
    Set rsAccess = CurrentDb().OpenRecordset("SELECT * FROM " & C_CURRENT_WORK_TABLE & " ORDER BY 1 DESC", dbOpenForwardOnly) 'INSERT 2012/10/11 K.ISHIZAKA
    With rsAccess
        If .EOF Then
            .Close
            Exit Sub
        End If
        Call MNN020.WebGoogleOpen(IEApp, True)
        While Not .EOF
            If Nz(.Fields("SERVICE1")) = 1 Then
                strMark = "r"
            ElseIf .Fields("SUM_AKI") >= 10 Then
                strMark = "b"
            Else
                strMark = "g"
            End If
            If .Fields("NO") < 25 Then
                strMark = strMark & Format(.Fields("NO"))
            End If
            If (Nz(.Fields("YARD_IDO")) > 0) And (Nz(.Fields("YARD_KEIDO")) > 0) Then
'                Call MNN020.WebGoogleMarker(IEApp, .Fields("YARD_IDO"), .Fields("YARD_KEIDO"), _
                    strMark, .Fields("YARD_CODE"), .Fields("YARD_NAME"), _
                    Nz(.Fields("YARD_ADDR")), Nz(.Fields("YARD_NOTE")))         'DELETE 2012/10/11 K.ISHIZAKA
                Call MNN020.WebGoogleMarker(IEApp, .Fields("YARD_IDO"), .Fields("YARD_KEIDO"), _
                    strMark, .Fields("YARD_CODE"), .Fields("YARD_NAME"), _
                    Nz(.Fields("YARD_ADDR")), Nz(.Fields("YARD_NOTE")), (.Fields("NO") = 1)) 'INSERT 2012/10/11 K.ISHIZAKA
            End If
            .MoveNext
        Wend
        .Close
    End With
    Set Elmnt = IEApp.Document.getElementById("divHelp")
    Elmnt.Style.Display = ""
Exit Sub

Err_Handler:
    rsAccess.Close
    MsgBox "MapOpen:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用テーブルをコピーして作成ＳＱＬ作成
'       MODULE_ID       : copyTempTable
'       CREATE_DATE     : 2012/06/16        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function copyTempTable(ByVal strTable As String) As String
    Dim strSQL              As String
    
    strSQL = strSQL & "CREATE TEMPORARY TABLE " & strTable
    strSQL = strSQL & " SELECT " & C_SERVER_TEMP_TABLE & ".* "
    strSQL = strSQL & " FROM " & C_SERVER_TEMP_TABLE

    copyTempTable = strSQL
End Function

'****************************  ended of program ********************************
