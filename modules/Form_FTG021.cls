VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FTG021"
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
'       PROGRAM_NAME    : 統合空詳細 メイン
'       PROGRAM_ID      : FTG021
'       PROGRAM_KBN     : FROM
'
'       CREATE          : 2011/02/12
'       CERATER         : K.ISHZIAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/03/08
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.1
'                       : CONT_EXCEL_PATHのパス修正
'
'       UPDATE          : 2011/03/10
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.2
'                       : サイズ条件を追加
'                       : コンテナ(バイク)のときボックスが表示されていなかった件を修正
'                       : コンテナ(バイク)のときヤードへの予約を不可とする
'                       : コンテナ(バイク)のとき予約したらコンテナマスタの用途を更新する
'                         また、価格テーブルが存在しなければコンテナの内容で作成する
'
'       UPDATE          : 2011/03/16
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.3
'                       : 用途39:バイク(コンテナ)の対応
'
'       UPDATE          : 2011/04/06
'       UPDATER         : K.ISHIZAKA
'       Ver             : 0.4
'                       : 自社理由移動まで表示されてしまっていたのを修正
'                         一部インデックスが有効に使われていなかったので型変換させる
'
'       UPDATE          : 2011/07/27
'       UPDATER         : M.HONDA
'       Ver             : 0.5
'                       : WEBページ表示の拡張子をhtm -> htmlへ修正
'
'       UPDATE          : 2012/01/23
'       UPDATER         : M.HONDA
'       Ver             : 0.6
'                       : 修繕・顧客コメントを表示するように修正。
'
'       UPDATE          : 2012/06/13
'       UPDATER         : M.HONDA
'       Ver             : 0.7
'                       : ｺﾝﾃﾅ(ﾊﾞｲｸ)を選択時に初回雑費 = 95をセット
'
'       UPDATE          : 2012/12/27
'       UPDATER         : M.HONDA
'       Ver             : 0.8
'                       : ﾊﾞｲｸ系の用途の場合には、毎月雑費をコピーしない。
'
'       UPDATE          : 2013/05/09
'       UPDATER         : M.HONDA
'       Ver             : 0.9
'                       : 起動LOG対応
'
'       UPDATE          : 2013/09/24
'       UPDATER         : M.HONDA
'       Ver             : 1.0
'                       : バイクの場合には雑費を0円として処理
'
'       UPDATE          : 2013/11/27
'       UPDATER         : K.ISHIZAKA
'       Ver             : 1.1
'                       : 引数、予約受付番号を受取り、予約画面に引き渡す
'
'       UPDATE          : 2014/03/14
'       UPDATER         : K.ISHIZAKA
'       Ver             : 1.2
'                       : 一覧対象には、同じコンテナの契約更新がある場合、除外する
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'統合ＤＢ接続部門コード
Private Const C_TOGO_CONNENCT_BUMOC As String = "TOGO"

Private Const C_CURRENT_WORK_TABLE  As String = "WK_FTG021_TEMP"

Private Const C_YARD_CODE       As Integer = 0  'ヤードコード
Private Const C_CNTA_USAGE      As Integer = 1  'レンタル用途
Private Const C_CONT_BUMOC      As Integer = 2  '部門コード
Private Const C_CONT_SIZE       As Integer = 3  'サイズ条件                     'INSERT 2011/03/10 K.ISHIZAKA
Private Const C_YOUKT_UKNO      As Integer = 4  '予約受付番号                   'INSERT 2013/11/27 K.ISHIZAKA
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private strOpenArgs()       As String       '引数配列

'==============================================================================*
'   イベント関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : 終了ボタンが押されたとき
'       MODULE_ID       : cmd_EXIT_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_EXIT_Click()
    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 配置ボタンが押されたとき
'       MODULE_ID       : cmd_HAITIZU_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_HAITIZU_Click()
    Dim strExcel            As String
    On Error GoTo ErrorHandler

    strExcel = Dir(txt_CONT_EXCEL_PATH.VALUE & Format(strOpenArgs(C_YARD_CODE), "000000") & "*.XLS")
    If strExcel <> "" Then
        Call Y_Exec(Me.hwnd, txt_CONT_EXCEL_PATH.VALUE & strExcel, "", "", SW_SHOWNORMAL)
    Else
        MsgBox "ｺﾝﾃﾅ配置EXCELが存在しません。", vbOKOnly + vbInformation, Me.Caption
    End If
Exit Sub

ErrorHandler:          '↓自分の関数名
    If Err.Number = 52 Then
        Err.Clear
        MsgBox "ｺﾝﾃﾅ配置EXCELが存在しません。", vbOKOnly + vbInformation, Me.Caption
    Else
        Call MSZZ024_M00("cmd_HAITIZU_Click", True)
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : WEBボタンが押されたとき
'       MODULE_ID       : cmd_WEB_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_WEB_Click()
    On Error GoTo ErrorHandler
    Call MSZZ009_M00(cmd_WEB, txt_SubAddress, txt_Address)
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmd_WEB_Click", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 予約受付ボタン押下時処理
'       MODULE_ID       : cmd_YOYAKU_Click
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_YOYAKU_Click()
    On Error GoTo ErrorHandler
'    Call ShowFVS500("", "")                                                    'DELETE 2011/03/10 K.ISHIZAKA
    Call ShowFVS500("", "", "", "")                                             'INSERT 2011/03/10 K.ISHIZAKA
Exit Sub

ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmd_YOYAKU_Click", True)
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
    Call MSHS010_M00("FTG021", "2")                                             ' 2013/05/09 M.HONDA INS
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
'       MODULE_NAME     : フォームロード時処理
'       MODULE_ID       : Form_Load
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    On Error GoTo ErrorHandler
    
    Call MSHS010_M00("FTG021", "1")                                             ' 2013/05/09 M.HONDA INS
    
    '引数は [ヤードコード],[レンタル用途],[部門コード]
    If Nz(Me.OpenArgs) = "" Then
        strOpenArgs = Split(InputBox("[ヤードコード],[レンタル用途],[部門コード]を指定してください。", Me.Caption, "400308,0,H"), ",")
    Else
        strOpenArgs = Split(Me.OpenArgs, ",")
    End If
    txt_YARD_CODE.VALUE = Format(strOpenArgs(C_YARD_CODE), "000000")
'    cmd_YOYAKU.Enabled = (strOpenArgs(C_CNTA_USAGE) <> "33")                   'DELETE 2011/03/16 K.ISHIZAKA 'INSERT 2011/03/10 K.ISHIZAKA
    cmd_YOYAKU.Enabled = ((strOpenArgs(C_CNTA_USAGE) <> "33") And (strOpenArgs(C_CNTA_USAGE) <> "39")) 'INSERT 2011/03/16 K.ISHIZAKA
    Call GetData
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("Form_Load", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YARD_CODE_DblClick
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YARD_CODE_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    Call Form_KeyDown(CmReprMod.P_呼出_ヤード, 0)
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("txt_YARD_CODE_DblClick", True)
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
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetData() As Long
    Dim objCon              As Object
    Dim objRst              As Object
    Dim objFld              As Object
    Dim rstWk               As Recordset
    Dim strKaseDb           As String
    Dim strKomsDb           As String
    Dim lngCnt              As Long
    Dim dbKase              As Object
    Dim rsKase              As Object
    
    On Error GoTo ErrorHandler

    strKaseDb = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'"), "")
    If strKaseDb = "" Then
        Call MSZZ024_M10("DLookup", "テーブル[SETU_TABL]の設定不正です。")
    End If
    strKomsDb = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME_" & strOpenArgs(C_CONT_BUMOC) & "'"), "")
    If strKomsDb = "" Then
        Call MSZZ024_M10("DLookup", "テーブル[SETU_TABL]の設定不正です。")
    End If
    
    Set objCon = ADODB_Connection(C_TOGO_CONNENCT_BUMOC)
    On Error GoTo ErrorHandler1
    Set objRst = ADODB_Recordset(Select_YARD_MAST(strKomsDb), objCon)
    On Error GoTo ErrorHandler2
    With objRst
        If .EOF Then
            Call MSZZ024_M10("select", "ヤードコード[" & strOpenArgs(C_YARD_CODE) & "]レンタル用途[" & strOpenArgs(C_CNTA_USAGE) & "]のデータが存在しません。")
        End If
        txt_YARD_NAME.VALUE = .Fields("YARD_NAME").VALUE
        txt_YARD_ADDR.VALUE = .Fields("YARD_ADDR").VALUE
        txt_YARD_NOTE.VALUE = .Fields("YARD_NOTE").VALUE
        txt_YARD_BIKON.VALUE = .Fields("YARD_BIKON").VALUE
        txt_YARD_SEV1N.VALUE = .Fields("YARD_SEV1N").VALUE
        txt_YARD_SEV2N.VALUE = .Fields("YARD_SEV2N").VALUE
        txt_YARD_SEV3N.VALUE = .Fields("YARD_SEV3N").VALUE
        txt_YARD_ENDEN.VALUE = .Fields("YARD_ENDEN").VALUE
        txt_YARD_MIKAKUTEI.VALUE = .Fields("YARD_MIKAKUTEI").VALUE
        txt_YARD_BEGIN_DAY.VALUE = .Fields("YARD_BEGIN_DAY").VALUE
        txt_YARD_RENTEND_DAY.VALUE = .Fields("YARD_RENTEND_DAY").VALUE
        
        If Nz(.Fields("CONT_DISKBACK").VALUE) = "" Then
            cmd_WEB.Enabled = False
        Else
            cmd_WEB.Enabled = True
            Select Case strOpenArgs(C_CONT_BUMOC)
            Case "1":  txt_Address = .Fields("CONT_DISKBACK").VALUE & Format(strOpenArgs(C_YARD_CODE), "000000") & ".php"
            '' 2011/07/27 M.HODNA START
            '' Case "8":  txt_Address = .Fields("CONT_DISKBACK").VALUE & strOpenArgs(C_YARD_CODE) & ".htm"
            '' Case Else: txt_Address = .Fields("CONT_DISKBACK").VALUE & Format(strOpenArgs(C_YARD_CODE), "000000") & ".htm"
            '' 2011/07/27 M.HODNA END
            Case "8":  txt_Address = .Fields("CONT_DISKBACK").VALUE & strOpenArgs(C_YARD_CODE) & ".html"
            Case Else: txt_Address = .Fields("CONT_DISKBACK").VALUE & Format(strOpenArgs(C_YARD_CODE), "000000") & ".html"
            
            
            End Select
        End If
        txt_CONT_EXCEL_PATH.VALUE = Nz(.Fields("CONT_EXCEL_PATH").VALUE)
        With txt_CONT_EXCEL_PATH                                                'INSERT START 2011/03/08 K.ISHIZAKA
            If (.VALUE <> "") And (Right(.VALUE, 1) <> "\") Then
                .VALUE = .VALUE & "\"
            End If
        End With                                                                'INSERT END   2011/03/08 K.ISHIZAKA
        cmd_HAITIZU.Enabled = (Nz(.Fields("CONT_EXCEL_PATH").VALUE) <> "")
        
        .Close
    End With
    On Error GoTo ErrorHandler1
    
    Set objRst = ADODB_Recordset(select_CNTA_MAST(strKaseDb, strKomsDb), objCon)
    On Error GoTo ErrorHandler2
    Call createWorkTable(objRst)
    Set rstWk = CurrentDb.OpenRecordset(C_CURRENT_WORK_TABLE, dbOpenDynaset, dbAppendOnly)
    
    '' 20120130 M.HONDA START
    Set dbKase = CreateObject("ADODB.Connection")
    dbKase.CommandTimeout = 180
    Call dbKase.Open(MSZZ007_M10())
    Set rsKase = CreateObject("ADODB.Recordset")
    rsKase.Open fncSelectREPR_TRAN(), dbKase, adOpenForwardOnly, adLockReadOnly
    '' 20120130 M.HONDA END
    
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
        
        '' 20120130 M.HONDA START
        Do While Not rsKase.EOF
            Select Case StrComp(rsKase.Fields("REPRT_ROOMC"), Format(objRst.Fields("CNTA_NO"), "000000"))
            Case 0
                rstWk.Fields("YOUKT_BIKO") = rsKase.Fields("REPRT_CNT1N")
                rsKase.MoveNext
                Exit Do
            Case 1
                Exit Do
            End Select
            rsKase.MoveNext
        Loop
        '顧客修繕を呼び出す
        If rstWk.Fields("YOUKT_YUKBN") = "20" Then
            rstWk.Fields("YOUKT_RENKBN") = fncSelectKOTOT_CONTN(dbKase, objRst.Fields("YOUKT_UCODE"))
        End If
        '' 20120130 M.HONDA END
        
        rstWk.UPDATE
        On Error GoTo ErrorHandler3
        lngCnt = lngCnt + 1
        objRst.MoveNext
    Wend
    rstWk.Close
    On Error GoTo ErrorHandler2
    objRst.Close
    On Error GoTo ErrorHandler1
    objCon.Close
    On Error GoTo ErrorHandler
    sub_UP.Form.Requery
    sub_DW.Form.Requery
    On Error Resume Next
    If cmd_YOYAKU.Enabled Then                                                  'INSERT 2011/03/10 K.ISHIZAKA
        cmd_YOYAKU.SetFocus
    Else                                                                        'INSERT START 2011/03/10 K.ISHIZAKA
        sub_DW.SetFocus
    End If                                                                      'INSERT END   2011/03/10 K.ISHIZAKA
    
    GetData = lngCnt
Exit Function

ErrorHandler4:
    rstWk.CancelUpdate
ErrorHandler3:
    rstWk.Close
ErrorHandler2:
    objRst.Close
ErrorHandler1:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "getDATA" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'       MODULE_NAME     : コンテナ修繕直近取得処理
'       MODULE_ID       : fncSelectREPR_TRAN
'       CREATE_DATE     :
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectREPR_TRAN() As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " REPRT_ROOMC,"
    strSQL = strSQL & " SUBSTRING(REPRT_GENTD, 3, 2) + '/' +"
    strSQL = strSQL & " SUBSTRING(REPRT_GENTD, 5, 2) + ' ' +"
    strSQL = strSQL & " CASE WHEN REPRT_TYPEC = '02'"
    strSQL = strSQL & " THEN"
    strSQL = strSQL & " ("
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "  CODET_NAMEN"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  CODE_TABL"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  CODET_SIKBC = '251'"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  CODET_CODEC = CAST(REPRT_REPRC AS NUMERIC)"
    strSQL = strSQL & " )"
    strSQL = strSQL & " ELSE "
    strSQL = strSQL & "  ISNULL(REPRT_CNT1N,'')"
    strSQL = strSQL & " END AS REPRT_CNT1N "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " REPR_TRAN oya "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " REPRT_BUMOC = '" & strOpenArgs(C_CONT_BUMOC) & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " REPRT_YARDC = '" & Format(strOpenArgs(C_YARD_CODE), "000000") & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " REPRT_INPTI = '03' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " REPRT_UNIQEC ="
    strSQL = strSQL & " ("
    strSQL = strSQL & " SELECT TOP 1"
    strSQL = strSQL & "  REPRT_UNIQEC"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  REPR_TRAN ko"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  ko.REPRT_BUMOC = oya.REPRT_BUMOC"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ko.REPRT_YARDC = oya.REPRT_YARDC"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ko.REPRT_ROOMC = oya.REPRT_ROOMC"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ko.REPRT_INPTI = oya.REPRT_INPTI"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ko.REPRT_TYPEC != '02' "
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & "  REPRT_GENTD DESC,"
    strSQL = strSQL & "  REPRT_INSED DESC,"
    strSQL = strSQL & "  REPRT_INSEJ DESC"
    strSQL = strSQL & " ) "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " REPRT_ROOMC"

    fncSelectREPR_TRAN = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 顧客修繕直近取得更新処理
'       MODULE_ID       : fncSelectKOTOT_CONTN
'       CREATE_DATE     :
'       PARAM           : dbSqlServer       データベース(I)
'                       : strKOTOT_KOKYC    顧客コード(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectKOTOT_CONTN(dbSQLServer As Object, ByVal strKOTOT_KOKYC As String) As String
    Dim rsObject            As Object
    
    Set rsObject = CreateObject("ADODB.Recordset")
    With rsObject
        .Open fncSelectKOTO_TRAN(Format(strKOTOT_KOKYC, "000000")), dbSQLServer, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then
            fncSelectKOTOT_CONTN = .Fields(0)
        End If
        .Close
    End With
End Function

'==============================================================================*
'
'       MODULE_NAME     : 顧客修繕直近取得処理
'       MODULE_ID       : fncSelectKOTO_TRAN
'       CREATE_DATE     :
'       PARAM           : strKOTOT_KOKYC    顧客コード(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectKOTO_TRAN(ByVal strKOTOT_KOKYC As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT TOP 1"
    strSQL = strSQL & " SUBSTRING(KOTOT_GENTD, 3, 2) + '/' +"
    strSQL = strSQL & " SUBSTRING(KOTOT_GENTD, 5, 2) + ' ' +"
    strSQL = strSQL & " ISNULL(KOTOT_CONTN, '') "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " KOTO_TRAN "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " KOTOT_BUMOC = '" & strOpenArgs(C_CONT_BUMOC) & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " KOTOT_KOKYC = '" & strKOTOT_KOKYC & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " KOTOT_DISPI = 1 "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " KOTOT_GENTD DESC,"
    strSQL = strSQL & " KOTOT_INSED DESC,"
    strSQL = strSQL & " KOTOT_INSEJ DESC "

    fncSelectKOTO_TRAN = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードマスタ取得
'       MODULE_ID       : select_YARD_MAST
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strKomsDb         ＫＯＭＳデータベース名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Select_YARD_MAST(ByVal strKomsDb As String) As String
    Dim strSQL              As String
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_ADDR_1 + ISNULL(YARD_ADDR_2,'') + ISNULL(YARD_ADDR_3,'') AS YARD_ADDR,"
    strSQL = strSQL & " YARD_NOTE,"
    strSQL = strSQL & " YARD_BIKON,"
    strSQL = strSQL & " YARD_SEV1N,"
    strSQL = strSQL & " YARD_SEV2N,"
    strSQL = strSQL & " YARD_SEV3N,"
    strSQL = strSQL & " YARD_ENDEN,"
    strSQL = strSQL & " YARD_MIKAKUTEI,"
    strSQL = strSQL & " CONVERT(VARCHAR(10), YARD_BEGIN_DAY, 111) AS YARD_BEGIN_DAY,"
    strSQL = strSQL & " CONVERT(VARCHAR(10), YARD_RENTEND_DAY, 111) AS YARD_RENTEND_DAY,"
    strSQL = strSQL & " CONT_DISKBACK,"
    strSQL = strSQL & " CONT_EXCEL_PATH "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CONT_MAST,"
    strSQL = strSQL & " TOGO_YARD_MAST "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " YARD_CODE = " & strOpenArgs(C_YARD_CODE) & " "
    strSQL = strSQL & "AND"
    If strOpenArgs(C_CNTA_USAGE) = "39" Then                                    'INSERT START 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " YARD_USAGE = 3 "
    Else                                                                        'INSERT END   2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " YARD_USAGE = " & strOpenArgs(C_CNTA_USAGE) & " "
    End If                                                                      'INSERT 2011/03/16 K.ISHIZAKA

    Select_YARD_MAST = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナマスタ取得
'       MODULE_ID       : select_CNTA_MAST
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strKaseDb         ＫＡＳＥデータベース名(I)
'                       : strKomsDb         ＫＯＭＳデータベース名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function select_CNTA_MAST(ByVal strKaseDb As String, ByVal strKomsDb As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_STEP,"
    strSQL = strSQL & " CNTA_NO,"
    strSQL = strSQL & " CNTA_SIZE,"
'    If strOpenArgs(C_CNTA_USAGE) = "33" Then                                   'DELETE 2011/03/16 K.ISHIZAKA 'INSERT START 2011/03/10 K.ISHIZAKA
    If (strOpenArgs(C_CNTA_USAGE) = "33") Or (strOpenArgs(C_CNTA_USAGE) = "39") Then 'INSERT 2011/03/16 K.ISHIZAKA
'        strSQL = strSQL & " CAST('33' AS VARCHAR(2)) AS CNTA_USAGE,"           'DELETE 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " CAST('" & strOpenArgs(C_CNTA_USAGE) & "' AS VARCHAR(2)) AS CNTA_USAGE," 'INSERT 2011/03/16 K.ISHIZAKA
    Else
        strSQL = strSQL & " CAST(RIGHT('00' + CAST(CNTA_USAGE AS VARCHAR(2)), 2) AS VARCHAR(2)) AS CNTA_USAGE,"
    End If                                                                      'INSERT END   2011/03/10 K.ISHIZAKA
    strSQL = strSQL & " NAME086.NAME_NAME AS CNTA_SYU_NAME,"
    
    strSQL = strSQL & " CNTA_PRICE_DIFF,"
    strSQL = strSQL & " CNTA_REASON,"
    strSQL = strSQL & " CNTA_HEIGHT,"
    strSQL = strSQL & " CNTA_WIDTH,"
    strSQL = strSQL & " CNTA_DEPTH,"
    
'    If strOpenArgs(C_CNTA_USAGE) = "33" Then                                   'DELETE 2011/03/16 K.ISHIZAKA 'INSERT START 2011/03/10 K.ISHIZAKA
    If (strOpenArgs(C_CNTA_USAGE) = "33") Or (strOpenArgs(C_CNTA_USAGE) = "39") Then 'INSERT 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " ISNULL(PRIC33.PRIC_PRICE,  PRIC00.PRIC_PRICE)  AS PRIC_PRICE,"
        '' 2013/09/24 M.HONDA UPD
        '' strSQL = strSQL & " ISNULL(PRIC33.PRIC_FZAPPI, PRIC00.PRIC_FZAPPI) AS PRIC_FZAPPI,"
        '' strSQL = strSQL & " ISNULL(PRIC33.PRIC_EZAPPI, PRIC00.PRIC_EZAPPI) AS PRIC_EZAPPI,"
        strSQL = strSQL & " ISNULL(PRIC33.PRIC_FZAPPI, 0) AS PRIC_FZAPPI,"
        strSQL = strSQL & " ISNULL(PRIC33.PRIC_EZAPPI, 0) AS PRIC_EZAPPI,"
        '' 2013/09/24 M.HONDA UPD
    Else                                                                        'INSERT END   2011/03/10 K.ISHIZAKA
        strSQL = strSQL & " PRIC_PRICE,"
        strSQL = strSQL & " PRIC_FZAPPI,"
        strSQL = strSQL & " PRIC_EZAPPI,"
    End If                                                                      'INSERT 2011/03/10 K.ISHIZAKA

    strSQL = strSQL & " NAME080.NAME_CODE AS YOUKT_YUKBN,"
    strSQL = strSQL & " NAME080.NAME_NAME AS YOUKT_YUKBN_NAME,"

    strSQL = strSQL & " YOUKT_UKNO,"
    strSQL = strSQL & " ISNULL(CARG_KYDATE, YOUKT_TKDATE) AS YOUKT_TKDATE,"
    strSQL = strSQL & " ISNULL(CARG_UCODE, YOUKT_UCODE) AS YOUKT_UCODE,"
    strSQL = strSQL & " ISNULL(USER_NAME, YOUKT_NAME) AS YOUKT_NAME,"
    strSQL = strSQL & " CASE NAME080.NAME_CODE"
    strSQL = strSQL & " WHEN 20 THEN"
    strSQL = strSQL & " ("  'この括弧内は KASE DB
    strSQL = strSQL & " SELECT TOP 1"
    strSQL = strSQL & "  SUBSTRING(KOTOT_GENTD, 3, 2) + '/' + SUBSTRING(KOTOT_GENTD, 5, 2) + ' ' +"
    strSQL = strSQL & "  ISNULL(KOTOT_CONTN, '')"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  " & strKaseDb & ".dbo.KOTO_TRAN "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  KOTOT_BUMOC = '" & strOpenArgs(C_CONT_BUMOC) & "'"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  KOTOT_KOKYC = RIGHT('000000' + CAST(CARG_UCODE AS VARCHAR(6)), 6)"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  KOTOT_DISPI = 1"
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & "  KOTOT_GENTD DESC,"
    strSQL = strSQL & "  KOTOT_INSED DESC,"
    strSQL = strSQL & "  KOTOT_INSEJ DESC"
    strSQL = strSQL & " )"
    strSQL = strSQL & " WHEN 2 THEN"
    strSQL = strSQL & "  NAME085.NAME_NAME + ':' +"
    strSQL = strSQL & "  CASE YOUKT_RENKBN"
    strSQL = strSQL & "  WHEN  0 THEN"
    strSQL = strSQL & "   YOUKT_TEL"
    strSQL = strSQL & "  WHEN  1 THEN"
    strSQL = strSQL & "   YOUKT_KEITAI"
    strSQL = strSQL & "  WHEN  2 THEN"
    strSQL = strSQL & "   YOUKT_FAX"
    strSQL = strSQL & "  WHEN 10 THEN"
    strSQL = strSQL & "   YOUKT_MAIL"
    strSQL = strSQL & "  ELSE"
    strSQL = strSQL & "   NULL"
    strSQL = strSQL & "  END"
    strSQL = strSQL & " ELSE"
    strSQL = strSQL & "  NULL"
    strSQL = strSQL & " END AS YOUKT_RENKBN,"

    strSQL = strSQL & "(" 'この括弧内は KASE DB
    strSQL = strSQL & " SELECT TOP 1"
    strSQL = strSQL & "  SUBSTRING(REPRT_GENTD, 3, 2) + '/' + SUBSTRING(REPRT_GENTD, 5, 2) + ' ' +"
    strSQL = strSQL & "  CASE WHEN REPRT_TYPEC = '02'"
    strSQL = strSQL & "  THEN"
    strSQL = strSQL & "  ("
    strSQL = strSQL & "   SELECT"
    strSQL = strSQL & "    CODET_NAMEN"
    strSQL = strSQL & "   FROM"
    strSQL = strSQL & "    " & strKaseDb & ".dbo.CODE_TABL"
    strSQL = strSQL & "   WHERE"
    strSQL = strSQL & "    CODET_SIKBC = '251'"
    strSQL = strSQL & "   AND"
    strSQL = strSQL & "    CODET_CODEC = CAST(REPRT_REPRC AS NUMERIC)"
    strSQL = strSQL & "  )"
    strSQL = strSQL & "  ELSE"
    strSQL = strSQL & "   ISNULL(REPRT_CNT1N,'')"
    strSQL = strSQL & "  END"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  " & strKaseDb & ".dbo.REPR_TRAN"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  REPRT_BUMOC = '" & strOpenArgs(C_CONT_BUMOC) & "'"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  REPRT_YARDC = CNTA_CODE"
    strSQL = strSQL & " AND"
'    strSQL = strSQL & "  REPRT_ROOMC = CNTA_NO"                                'DELETE 2011/04/06 K.ISHIZAKA
    strSQL = strSQL & "  REPRT_ROOMC = CAST(CNTA_NO AS CHAR(6))"                'INSERT 2011/04/06 K.ISHIZAKA
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  REPRT_INPTI = '03'"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  REPRT_TYPEC != '02' "
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & "  REPRT_GENTD DESC,"
    strSQL = strSQL & "  REPRT_INSED DESC,"
    strSQL = strSQL & "  REPRT_INSEJ DESC"
    strSQL = strSQL & ") AS YOUKT_BIKO "

    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CNTA_MAST "
'    If strOpenArgs(C_CNTA_USAGE) = "33" Then                                   'DELETE 2011/03/16 K.ISHIZAKA 'INSERT START 2011/03/10 K.ISHIZAKA
    If (strOpenArgs(C_CNTA_USAGE) = "33") Or (strOpenArgs(C_CNTA_USAGE) = "39") Then 'INSERT 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & "LEFT JOIN"
        strSQL = strSQL & " " & strKomsDb & ".dbo.PRIC_TABL PRIC00 "
        strSQL = strSQL & "ON"
        strSQL = strSQL & " PRIC00.PRIC_YCODE = CNTA_CODE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC00.PRIC_STEP  = CNTA_STEP "
        strSQL = strSQL & "AND"
'        strSQL = strSQL & " PRIC00.PRIC_USAGE = 0 "                            'DELETE 2011/03/16 K.ISHIZAKA
        Select Case strOpenArgs(C_CNTA_USAGE)                                   'INSERT START 2011/03/16 K.ISHIZAKA
        Case "33"
            strSQL = strSQL & " PRIC00.PRIC_USAGE = 0 "
        Case "39"
            strSQL = strSQL & " PRIC00.PRIC_USAGE = 3 "
        Case Else
            strSQL = strSQL & " PRIC00.PRIC_USAGE is null "
        End Select                                                              'INSERT END   2011/03/16 K.ISHIZAKA
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC00.PRIC_SIZE  = CNTA_SIZE "
        strSQL = strSQL & "LEFT JOIN"
        strSQL = strSQL & " " & strKomsDb & ".dbo.PRIC_TABL PRIC33 "
        strSQL = strSQL & "ON"
        strSQL = strSQL & " PRIC33.PRIC_YCODE = CNTA_CODE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC33.PRIC_STEP  = CNTA_STEP "
        strSQL = strSQL & "AND"
'        strSQL = strSQL & " PRIC33.PRIC_USAGE = 33 "                           'DELETE 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " PRIC33.PRIC_USAGE = " & strOpenArgs(C_CNTA_USAGE) & " " 'INSERT 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC33.PRIC_SIZE  = CNTA_SIZE "
    Else                                                                        'INSERT END   2011/03/10 K.ISHIZAKA
        strSQL = strSQL & "LEFT JOIN"
        strSQL = strSQL & " " & strKomsDb & ".dbo.PRIC_TABL "
        strSQL = strSQL & "ON"
        strSQL = strSQL & " PRIC_YCODE = CNTA_CODE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC_STEP  = CNTA_STEP "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC_USAGE = CNTA_USAGE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " PRIC_SIZE  = CNTA_SIZE "
    End If                                                                      'INSERT 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.NAME_MAST AS NAME086 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " NAME086.NAME_ID = '086' "
    strSQL = strSQL & "AND"
'    If strOpenArgs(C_CNTA_USAGE) = "33" Then                                   'DELETE 2011/03/16 K.ISHIZAKA 'INSERT START 2011/03/10 K.ISHIZAKA
    If (strOpenArgs(C_CNTA_USAGE) = "33") Or (strOpenArgs(C_CNTA_USAGE) = "39") Then 'INSERT 2011/03/16 K.ISHIZAKA
'        strSQL = strSQL & " NAME086.NAME_CODE = 33 "                           'DELETE 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " NAME086.NAME_CODE = " & strOpenArgs(C_CNTA_USAGE) & " " 'INSERT 2011/03/16 K.ISHIZAKA
    Else                                                                        'INSERT END   2011/03/10 K.ISHIZAKA
        strSQL = strSQL & " NAME086.NAME_CODE = CNTA_USAGE "
    End If                                                                      'INSERT 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CARG_FILE "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CARG_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_NO = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_CONTNA = 0 "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.USER_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " USER_CODE = CARG_UCODE "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.INTR_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " INTRT_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_NO = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_INTROKBN = 1 "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.YOUK_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " YOUKT_UKNO = INTRT_UKNO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " YOUKT_YUKBN = 2 "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.NAME_MAST AS NAME085 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " NAME085.NAME_ID = '085' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " NAME085.NAME_CODE = YOUKT_RENKBN "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.NAME_MAST AS NAME080 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " NAME080.NAME_ID = '080' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " NAME080.NAME_CODE = CASE"
    strSQL = strSQL & " WHEN YOUKT_UKNO IS NOT NULL THEN 2"
    strSQL = strSQL & " WHEN CARG_KYDATE IS NOT NULL THEN 20"
    strSQL = strSQL & " ELSE NULL END "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CNTA_USE = 1 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & "("
    strSQL = strSQL & " CARG_CONTNA IS NULL "
    strSQL = strSQL & "OR"
    strSQL = strSQL & " ("                                                      'INSERT 2014/03/14 K.ISHIZAKA
    strSQL = strSQL & " CARG_KYDATE IS NOT NULL "
    strSQL = strSQL & "  AND NOT EXISTS"                                        'INSERT START 2014/03/14 K.ISHIZAKA
    strSQL = strSQL & "  ("                                                     '同じコンテナの契約更新があったらダメよ
    strSQL = strSQL & "  SELECT * "
    strSQL = strSQL & "  FROM " & strKomsDb & ".dbo.YOUK_TRAN youk "
    strSQL = strSQL & "  INNER JOIN"
    strSQL = strSQL & "   " & strKomsDb & ".dbo.CARG_FILE carg "
    strSQL = strSQL & "  ON"
    strSQL = strSQL & "   carg.CARG_UKNO = youk.YOUKT_UKNO "
    strSQL = strSQL & "  WHERE "
    strSQL = strSQL & "   youk.YOUKT_MOTO_ACPTNO = " & strKomsDb & ".dbo.CARG_FILE.CARG_ACPTNO "
    strSQL = strSQL & "  AND"
    strSQL = strSQL & "   youk.YOUKT_YUKBN = 20 "
    strSQL = strSQL & "  AND"
    strSQL = strSQL & "   carg.CARG_YCODE = CNTA_CODE "
    strSQL = strSQL & "  AND"
    strSQL = strSQL & "   carg.CARG_NO = CNTA_NO "
    strSQL = strSQL & "  AND"
    strSQL = strSQL & "   carg.CARG_CONTNA = 0 "
    strSQL = strSQL & "  )"
    strSQL = strSQL & " )"                                                      'INSERT END   2014/03/14 K.ISHIZAKA
    strSQL = strSQL & ")"
    strSQL = strSQL & "AND NOT EXISTS"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM " & strKomsDb & ".dbo.INTR_TRAN "
    strSQL = strSQL & "INNER JOIN " & strKomsDb & ".dbo.YOUK_TRAN "             'INSERT 2011/04/06 K.ISHIZAKA
    strSQL = strSQL & "ON YOUKT_UKNO = INTRT_UKNO "                             'INSERT 2011/04/06 K.ISHIZAKA
    strSQL = strSQL & "WHERE INTRT_YCODE    = CNTA_CODE "
    strSQL = strSQL & "AND   INTRT_NO       = CNTA_NO "
'    strSQL = strSQL & "AND   INTRT_INTROKBN = 2 "                              'DELETE 2011/04/06 K.ISHIZAKA
    strSQL = strSQL & "AND ( INTRT_INTROKBN = 2 OR YOUKT_YUKBN = 1 )"           'INSERT 2011/04/06 K.ISHIZAKA
    strSQL = strSQL & ")"
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_CODE = " & strOpenArgs(C_YARD_CODE) & " "
    strSQL = strSQL & "AND"
    If strOpenArgs(C_CNTA_USAGE) = "33" Then                                    'INSERT START 2011/03/10 K.ISHIZAKA
        strSQL = strSQL & " CNTA_USAGE = 0 "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " CNTA_BIKE_OK = 1 "
    ElseIf strOpenArgs(C_CNTA_USAGE) = "39" Then                                'INSERT 2011/03/16 K.ISHIZAKA
        strSQL = strSQL & " CNTA_USAGE = 3 "                                    'INSERT 2011/03/16 K.ISHIZAKA
    Else                                                                        'INSERT END   2011/03/10 K.ISHIZAKA
        strSQL = strSQL & " CNTA_USAGE = " & strOpenArgs(C_CNTA_USAGE) & " "
    End If                                                                      'INSERT 2011/03/10 K.ISHIZAKA
    If strOpenArgs(C_CONT_SIZE) <> "" Then                                      'INSERT START 2011/03/10 K.ISHIZAKA
        strSQL = strSQL & "AND"
        strSQL = strSQL & " CNTA_SIZE " & strOpenArgs(C_CONT_SIZE) & " "
    End If                                                                      'INSERT END   2011/03/10 K.ISHIZAKA
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_NO,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP "

    select_CNTA_MAST = strSQL
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
    End With
    For Each objFld In objRst.Fields
        With objFld
            If .Type = AdoEnums_DataType.adNumeric Then
                If .NumericScale > 0 Then
                    Set fldNew = tdfNew.CreateField(.NAME, dbDouble)
                Else
                    Set fldNew = tdfNew.CreateField(.NAME, dbLong)
                End If
            ElseIf .Type = AdoEnums_DataType.adDBTimeStamp Then
                Set fldNew = tdfNew.CreateField(.NAME, dbDate)
            Else
                Set fldNew = tdfNew.CreateField(.NAME, dbText, .DefinedSize)
                fldNew.AllowZeroLength = True
            End If
        End With
        tdfNew.Fields.Append fldNew
    Next
    CurrentDb().TableDefs.Append tdfNew
    sub_UP.SourceObject = "FTG021_S01"
    sub_DW.SourceObject = "FTG021_S02"
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
    sub_UP.SourceObject = ""
    sub_DW.SourceObject = ""
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
'       MODULE_NAME     : 価格テーブル登録
'       MODULE_ID       : insertPricTabl
'       CREATE_DATE     : 2011/03/10        K.ISHIZAKA
'       PARAM           : strSize           サイズ(I)
'                       : strStep           上下段(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function insertPricTabl(ByVal strSize As String, ByVal strStep As String) As String
    Dim strSQL              As String
    
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & " PRIC_TABL "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " PRIC_YCODE,"
'    strSQL = strSQL & " 33 AS PRIC_USAGE,"                                     'DELETE 2011/03/16 K.ISHIZAKA
    strSQL = strSQL & " " & strOpenArgs(C_CNTA_USAGE) & " AS PRIC_USAGE,"       'INSERT 2011/03/16 K.ISHIZAKA
    strSQL = strSQL & " PRIC_SIZE,"
    strSQL = strSQL & " PRIC_STEP,"
    strSQL = strSQL & " PRIC_PRICE,"
    
    '' INS 2012/06/13 M.HONDA START
    If strOpenArgs(C_CNTA_USAGE) = 33 Then
        strSQL = strSQL & " 95 AS PRIC_FZAPPI_CODE,"
    Else
        strSQL = strSQL & " PRIC_FZAPPI_CODE,"
    End If
    '' INS 2012/06/13 M.HONDA END
    strSQL = strSQL & " PRIC_FZAPPI,"
    
    '' 2012/12/27 M.HONDA INS
    If strOpenArgs(C_CNTA_USAGE) = 3 Or strOpenArgs(C_CNTA_USAGE) = 7 Or strOpenArgs(C_CNTA_USAGE) = 30 Or strOpenArgs(C_CNTA_USAGE) = 31 _
       Or strOpenArgs(C_CNTA_USAGE) = 32 Or strOpenArgs(C_CNTA_USAGE) = 33 Then
        strSQL = strSQL & " 0,"
        strSQL = strSQL & " 0,"
    Else
        strSQL = strSQL & " PRIC_EZAPPI_CODE,"
        strSQL = strSQL & " PRIC_EZAPPI,"
    End If
    '' 2012/12/27 M.HONDA INS
    
    strSQL = strSQL & " PRIC_NEBIKI_DISABLE,"
    strSQL = strSQL & " '" & Format(Now, "yyyymmdd") & "' AS PRIC_INSED,"
    strSQL = strSQL & " '" & Format(Now, "hhnnss") & "' AS PRIC_INSEJ,"
    strSQL = strSQL & " '" & Me.NAME & "' AS PRIC_INSPB,"
    strSQL = strSQL & " '" & LsGetUserName() & "' AS PRIC_INSUB,"
    strSQL = strSQL & " NULL AS PRIC_UPDAD,"
    strSQL = strSQL & " NULL AS PRIC_UPDAT,"
    strSQL = strSQL & " NULL AS PRIC_UPDPB,"
    strSQL = strSQL & " NULL AS PRIC_UPDUB "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " PRIC_TABL par "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " par.PRIC_YCODE = " & strOpenArgs(C_YARD_CODE) & " "
    strSQL = strSQL & "AND"
'    strSQL = strSQL & " par.PRIC_USAGE = 0 "                                   'DELETE 2011/03/16 K.ISHIZAKA
    Select Case strOpenArgs(C_CNTA_USAGE)                                       'INSERT START 2011/03/16 K.ISHIZAKA
    Case "33"
        strSQL = strSQL & " par.PRIC_USAGE = 0 "
    Case "39"
        strSQL = strSQL & " par.PRIC_USAGE = 3 "
    Case Else
        strSQL = strSQL & " par.PRIC_USAGE is null "
    End Select                                                                  'INSERT END   2011/03/16 K.ISHIZAKA
    strSQL = strSQL & "AND"
    strSQL = strSQL & " par.PRIC_SIZE = " & strSize & " "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " par.PRIC_STEP = " & strStep & " "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " NOT EXISTS"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " PRIC_TABL ch "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " ch.PRIC_YCODE = par.PRIC_YCODE "
    strSQL = strSQL & "AND"
'    strSQL = strSQL & " ch.PRIC_USAGE = 33 "                                   'DELETE 2011/03/16 K.ISHIZAKA
    strSQL = strSQL & " ch.PRIC_USAGE = " & strOpenArgs(C_CNTA_USAGE) & " "     'INSERT 2011/03/16 K.ISHIZAKA
    strSQL = strSQL & "AND"
    strSQL = strSQL & " ch.PRIC_SIZE = par.PRIC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " ch.PRIC_STEP = par.PRIC_STEP "
    strSQL = strSQL & ")"

    insertPricTabl = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナマスタ更新
'       MODULE_ID       : updateCntaMast
'       CREATE_DATE     : 2011/03/10        K.ISHIZAKA
'       PARAM           : strNo             コンテナ番号(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function updateCntaMast(ByVal strNO As String) As String
    Dim strSQL              As String
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & " CNTA_MAST "
    strSQL = strSQL & "SET"
'    strSQL = strSQL & " CNTA_USAGE = 33,"                                     'DELETE 2011/03/16 K.ISHIZAKA
    strSQL = strSQL & " CNTA_USAGE = " & strOpenArgs(C_CNTA_USAGE) & ","       'INSERT 2011/03/16 K.ISHIZAKA
    strSQL = strSQL & " CNTA_UPDATE = GETDATE() "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CNTA_CODE = " & strOpenArgs(C_YARD_CODE) & " "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_NO = " & strNO & " "

    updateCntaMast = strSQL
End Function
    
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
    
    'F1:戻る
    'F2:保存
    'F3:取消
    'F5:段切り替え
    'F8:配置
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        cmd_EXIT.SetFocus
    Case vbKeyF5
        If Me.ActiveControl.Properties("NAME") = sub_UP.NAME Then
            With sub_DW
                .SetFocus
                .Form!cmdUke.SetFocus
            End With
        Else
            With sub_UP
                .SetFocus
                .Form!cmdUke.SetFocus
            End With
        End If
    Case vbKeyF8
        Call CmReprMod.OpenFVS90n(KeyCode, Format(strOpenArgs(C_YARD_CODE), "000000"))
        KeyCode = 0
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        Select Case Me.ActiveControl.Properties("NAME")
        Case sub_UP.NAME, sub_DW.NAME
            Call CmReprMod.OpenFVS90n(KeyCode, Format(strOpenArgs(C_YARD_CODE), "000000"), Me.ActiveControl.Form!txt_YOUKT_NO, Me.ActiveControl.Form!txt_YOUKT_UCODE)
        End Select
        KeyCode = 0
    End Select
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 詳細データ検索処理
'       MODULE_ID       : Reload
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strYard1          ヤードコード１
'                       : strYard2          ヤードコード２
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Reload(ByVal strYcode1 As String, ByVal strUsage1 As String, ByVal strYcode2 As String, ByVal strUsage2 As String) As Boolean
    If Val(strOpenArgs(C_YARD_CODE)) = Val(strYcode1) Then
        If (strUsage1 = "") Or (Val(strUsage1) = Val(strOpenArgs(C_CNTA_USAGE))) Then
            Call GetData
        End If
    ElseIf Val(strOpenArgs(C_YARD_CODE)) = Val(strYcode2) Then
        If (strUsage2 = "") Or (Val(strUsage2) = Val(strOpenArgs(C_CNTA_USAGE))) Then
            Call GetData
        End If
    End If
    Reload = True
End Function

'==============================================================================*
'
'       MODULE_NAME     : 予約受付画面表示処理
'       MODULE_ID       : ShowFVS500
'       CREATE_DATE     : 2011/02/12        K.ISHIZAKA
'       PARAM           : strYOUKT_NO       コンテナ番号
'                       : strYOUKT_UKNO     予約受付番号
'                       : strSize           サイズ(I)
'                       : strStep           上下段(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub ShowFVS500(ByVal strYOUKT_NO As String, ByVal strYOUKT_UKNO As String) 'DELETE 2011/03/10 K.ISHIZAKA
Public Sub ShowFVS500(ByVal strYOUKT_NO As String, ByVal strYOUKT_UKNO As String, _
    ByVal strSize As String, ByVal strStep As String)                           'INSERT 2011/03/10 K.ISHIZAKA
    Dim objCon              As Object                                           'INSERT 2011/03/10 K.ISHIZAKA
    Dim lngCnt              As Long                                             'INSERT 2011/03/10 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    '予約番号が画面起動時の引数でわたってきたとき                               'INSERT START 2013/11/27 K.ISHIZAKA
    If UBound(strOpenArgs) >= C_YOUKT_UKNO Then
        '画面起動時の引数の予約番号がカラでない（指定されている）とき
        If strOpenArgs(C_YOUKT_UKNO) <> "" Then
            '関数に渡された予約番号が指定されていて、画面起動時の引数の予約番号と違うときはエラー
            If (strYOUKT_UKNO <> "") And (strYOUKT_UKNO <> strOpenArgs(C_YOUKT_UKNO)) Then
                MsgBox "その部屋は指定できません。", vbOKOnly + vbInformation, Me.Caption
                Exit Sub
            End If
            strYOUKT_UKNO = strOpenArgs(C_YOUKT_UKNO)
        End If
    End If                                                                      'INSERT END   2013/11/27 K.ISHIZAKA
    If fncFormIsLoadedMsg("FVS400", Me.Caption) Then
        Exit Sub
    End If
    If fncFormIsLoadedMsg("FVS550", Me.Caption) Then
        Exit Sub
    End If
    If fncFormIsLoadedMsg("FVS500", Me.Caption) Then
        Exit Sub
    End If
    If DLookup("CONT_BUMOC", "dbo_CONT_MAST") <> strOpenArgs(C_CONT_BUMOC) Then
        Call MSHS012_M10(strOpenArgs(C_CONT_BUMOC))
    End If
    If strYOUKT_NO <> "" Then
'        If strOpenArgs(C_CNTA_USAGE) = "33" Then                               'DELETE 2011/03/16 K.ISHIZAKA 'INSERT START 2011/03/10 K.ISHIZAKA
        If (strOpenArgs(C_CNTA_USAGE) = "33") Or (strOpenArgs(C_CNTA_USAGE) = "39") Then 'INSERT 2011/03/16 K.ISHIZAKA
            Set objCon = ADODB_Connection(strOpenArgs(C_CONT_BUMOC))
            On Error GoTo ErrorHandler1
            lngCnt = ADODB_Execute(insertPricTabl(strSize, strStep), objCon)
            lngCnt = ADODB_Execute(updateCntaMast(strYOUKT_NO), objCon)
            objCon.Close
            On Error GoTo ErrorHandler
        End If                                                                  'INSERT END   2011/03/10 K.ISHIZAKA
        strYOUKT_NO = Format(Val(strYOUKT_NO), "000000")
    End If
    doCmd.OpenForm "FVS500", , , , , , Format(strOpenArgs(C_YARD_CODE), "000000") & strYOUKT_NO & strYOUKT_UKNO & "," & strOpenArgs(C_CNTA_USAGE)
Exit Sub

ErrorHandler1:                                                                  'INSERT 2011/03/10 K.ISHIZAKA
    objCon.Close                                                                'INSERT 2011/03/10 K.ISHIZAKA
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "ShowFVS500" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'****************************  ended of program ********************************
