VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800
'   作　成　日　　　：  2008/06/02
'   作　成　者　　　：  iizuka
'   Ver             ：  0.0
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800
'   作　成　日　　　：  2013/06/03
'   作　成　者　　　：  M.HONDA
'   Ver             ：  0.1
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800
'   作　成　日　　　：  2013/11/22
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.2
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800
'   作　成　日　　　：  2014/01/17
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.3
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800
'   作　成　日　　　：  2016/12/13
'   作　成　者　　　：  M.HONDA
'   Ver             ：  0.4
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800
'   作　成　日　　　：  2018/03/16
'   作　成　者　　　：  M.HONDA
'   Ver             ：  0.5
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800
'   作　成　日　　　：  2023/12/08
'   作　成　者　　　：  N.IMAI
'   Ver             ：  0.7
'**********************************************

Option Compare Database
Option Explicit

Private Const PROG_ID = "FKS800"
Private Const P_ROOM_WORK_TABLE         As String = "WK_FKS800"
Private Const P_BOX_LIST_WORK_TABLE     As String = "WK_FKS800_S01_TEMP"
Private Const P_COMBO_BOX_WORK_TABLE    As String = "WK_FKS810_CMBD"
'↓INSERT 2013/11/22 MIYAMOTO
Private Const P_ORNER_INFO_TABLE        As String = "WK_FKS800_ORNER_INFO"
Private Const P_EXCEL_WORK_TABLE        As String = "WK_FKS800_EXCEL"       'ワークテーブル名
Private Const P_QUERY_NAME              As String = "FKS800_Q01"            'クエリ名
Private pobjTableDef                    As DAO.TableDef                     'ワークテーブル
Private pobjQueryDef                    As DAO.QueryDef                     'クエリ
'↑INSERT 2013/11/22 MIYAMOTO

Private pobjDB                       As Object

Private pstrBumonCode                As String       ' 部門コード
Private anStrParameter               As String

Public gstrUserName                  As String
'↓INSERT 2013/11/22 MIYAMOTO
Public gstrAutoBoxNoKaHi             As Integer
'↑INSERT 2013/11/22 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :cmb_状況_AfterUpdate
'        MODULE_ID        :cmb_状況_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_状況_AfterUpdate()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception

    If (Not Me.cmb_状況.ListIndex = -1) Then
        '検索
        Call createBoxData
    End If
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_状況_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_状況_KeyDown
'        MODULE_ID        :cmb_状況_KeyDown
'        NOTE             :
'        CREATE_DATE      :2008/06/23
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_状況_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '入力を許可しない
    KeyCode = 0

End Sub

'↓INSERT 2013/11/22 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :cmd_EXCEL_Click
'        MODULE_ID        :cmd_EXCEL_Click
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    Dim isException     As Boolean
    isException = False

On Error GoTo Exception

    If vbYes = MsgBox("印刷内容をEXCELに出力します。よろしいですか？", vbQuestion + vbYesNo, "確認") Then
        Call subOutputExcel
        'Excel起動
        doCmd.OutputTo acQuery, P_QUERY_NAME, "MicrosoftExcel(*.xls)", "", True, ""
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_EXCEL_Click", True)

End Sub
'↑INSERT 2013/11/22 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :cmd_ボックス検索_Click
'        MODULE_ID        :cmd_ボックス検索_Click
'        NOTE             :
'        CREATE_DATE      :
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_ボックス検索_Click()

    Dim objAdoDbConnection   As Object
    Dim rsData               As Object
    Dim rsBoxData            As Recordset
    Dim sSQL                 As String


    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    
    'ワークテーブルの存在チェック
    sSQL = ""
    sSQL = "SELECT "
    sSQL = sSQL & "CBYST_YCODE "
    sSQL = sSQL & "FROM  CBYS_TRAN "
    sSQL = sSQL & "WHERE CBYST_BOXNO = '" & txt_box & "'"
    Set rsData = MSZZ025.ADODB_Recordset(sSQL, objAdoDbConnection)

    While Not rsData.EOF
        MsgBox "ヤードコード" & " : " & (rsData.Fields("CBYST_YCODE"))
        rsData.MoveNext
    Wend

    rsData.Close
    Set rsData = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    Exit Sub

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsBoxData Is Nothing Then: rsBoxData.Close: Set rsBoxData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "getExcelData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_検索L_Click
'        MODULE_ID        :cmd_検索L_Click
'        NOTE             :
'        CREATE_DATE      :2016/12/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索_Click()

    Dim isException As Boolean
    isException = False

On Error GoTo Exception

    '検索
    Call createBoxData
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_検索_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :終了ボタン押下時処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    '' 2013/06/03 M.HONDA START
'    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, "確認") = vbYes Then
'        doCmd.Close acForm, Me.FormName
'    End If
    '' 2013/06/03 M.HONDA END
    doCmd.Close acForm, Me.FormName
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_終了_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤード_AfterUpdate
'        MODULE_ID        :cmb_ヤード_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤード_AfterUpdate()

    Dim isException As Boolean
'↓INSERT 2013/11/22 MIYAMOTO
    Dim strTINTI    As String
    Dim strKentiku  As String
    Dim strSQL      As String
'↑INSERT 2013/11/22 MIYAMOTO
    isException = False

On Error GoTo Exception

    If Me.cmb_ヤード.ListIndex = -1 Then
        Me.lbl_ヤード名.Caption = ""
        Me.lbl_ヤード所在地.Caption = ""
        Me.lbl_営業開始日.Caption = ""
        Me.lbl_営業終了日.Caption = ""
        Me.lbl_ヤード解約予定日.Caption = ""
'↓INSERT 2013/11/22 MIYAMOTO
        Me.lbl_借上物件NO.Caption = ""
        Me.lbl_賃貸種別.Caption = ""
        Me.lbl_建築確認.Caption = ""
        Me.txt_借上物件NO.VALUE = ""
        Me.オーナー情報.Requery
'↑INSERT 2013/11/22 MIYAMOTO
    Else
        Me.lbl_ヤード名.Caption = Me.cmb_ヤード.Column(1, Me.cmb_ヤード.ListIndex)
        Me.lbl_ヤード所在地.Caption = Me.cmb_ヤード.Column(2, Me.cmb_ヤード.ListIndex)
        Me.lbl_営業開始日.Caption = Me.cmb_ヤード.Column(3, Me.cmb_ヤード.ListIndex)
        Me.lbl_営業終了日.Caption = Me.cmb_ヤード.Column(4, Me.cmb_ヤード.ListIndex)
        Me.lbl_ヤード解約予定日.Caption = Me.cmb_ヤード.Column(5, Me.cmb_ヤード.ListIndex)
'↓INSERT 2013/11/22 MIYAMOTO
        Me.lbl_借上物件NO.Caption = IIf(IsNull(Me.cmb_ヤード.Column(6, Me.cmb_ヤード.ListIndex)), "", Me.cmb_ヤード.Column(6, Me.cmb_ヤード.ListIndex))
        strTINTI = IIf(IsNull(Me.cmb_ヤード.Column(7, Me.cmb_ヤード.ListIndex)), "", Me.cmb_ヤード.Column(7, Me.cmb_ヤード.ListIndex))
        If strTINTI = "" Then
            Me.lbl_賃貸種別.Caption = ""
        Else
            strSQL = "NAME_ID='207' AND NAME_CODE=" & strTINTI
            Me.lbl_賃貸種別.Caption = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL))
        End If
        strKentiku = IIf(IsNull(Me.cmb_ヤード.Column(8, Me.cmb_ヤード.ListIndex)), "", Me.cmb_ヤード.Column(8, Me.cmb_ヤード.ListIndex))
        If strKentiku = "" Then
            Me.lbl_建築確認.Caption = ""
        Else
            strSQL = "NAME_ID='041' AND NAME_CODE=" & strKentiku
            Me.lbl_建築確認.Caption = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL))
        End If
        Me.txt_借上物件NO.VALUE = Me.lbl_借上物件NO.Caption
        Me.オーナー情報.Requery
'↑INSERT 2013/11/22 MIYAMOTO
    End If
    
    Me.cmb_ヤード = Format(Nz(Me.cmb_ヤード), "000000")
    Me.cmd_新規追加.Enabled = (Me.cmb_ヤード <> "")
'↓INSERT 2013/11/22 MIYAMOTO
    Me.cmd_配置図.Enabled = (Me.cmb_ヤード <> "")
    Me.cmd_Excel.Enabled = (Me.cmb_ヤード <> "")
    Me.chk_自動発番可否.Enabled = (Me.cmb_ヤード <> "")
'↑INSERT 2013/11/22 MIYAMOTO
    
    Me.cmd_検索.Enabled = True   '2016/12/13 M.HONDA INS
    
    '検索
    Call createBoxData
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_ヤード_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_新規追加_Click
'        MODULE_ID        :cmd_新規追加_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_新規追加_Click()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    Call deleteWorkTable(False, True)

'↓INSERT 2013/11/22 MIYAMOTO
    gstrAutoBoxNoKaHi = Me.chk_自動発番可否.VALUE
'↑INSERT 2013/11/22 MIYAMOTO

    anStrParameter = Me.cmb_ヤード & "," & Me.lbl_ヤード名.Caption & "," & Me.lbl_営業開始日.Caption & "," & Me.lbl_営業終了日.Caption & "," & Me.lbl_ヤード解約予定日.Caption & ",,,,,,,,,,,,,,,,,,新規登録,0"
    doCmd.OpenForm "FKS810", , , , , acDialog, anStrParameter
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_新規追加_Click", True)
    
End Sub

'↓INSERT 2013/11/22 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :cmd_配置図_Click
'        MODULE_ID        :cmd_配置図_Click
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_配置図_Click()
    
    Dim valRet          As Variant
    Dim strExcel        As String
    Dim isException     As Boolean
    isException = False

On Error GoTo Exception

    Call MSHS010_M00(PROG_ID, "3", Me.cmb_ヤード)

    'Excel情報取得
    '[コントロールマスタ].Excel情報パス 存在チェック
    valRet = Nz(DLookup("CONT_EXCEL_PATH", "dbo_CONT_MAST", ""), "")

    If Right(valRet, 1) = "\" Then
        valRet = valRet
    Else
        valRet = valRet & "\"
    End If

    '[コントロールマスタ].EXCEL情報パス のディレクトリにExcelが存在するか検索する。
    'EXCELのファイル名は "(ヤードコード).XLS"とする。
    strExcel = Dir(valRet & Format(Me.cmb_ヤード, "000000") & "*.XLS")

    If strExcel <> "" Then
        Call Y_Exec(Me.hwnd, valRet & strExcel, "", "", SW_SHOWNORMAL)
        Exit Sub
    Else
        MsgBox "ｺﾝﾃﾅ配置EXCELが存在しません。", vbOK, "確認"
        Exit Sub
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_配置図_Click", True)

End Sub
'↑INSERT 2013/11/22 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    
    doCmd.Maximize
    doCmd.Requery

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Close
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    'テーブル削除
    Call dropWorkTable(1)
    Call dropWorkTable(2)
    Call dropWorkTable(3)
    
'↓INSERT 2013/11/22 MIYAMOTO
    Call dropWorkTable(4)
    'EXCEL出力用ワークテーブル、クエリー削除
    Call subDropWorkTable
    Call subDropQuery
'↑INSERT 2013/11/22 MIYAMOTO
    
    'コネクションクローズ
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Close", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb
    
    'ワークテーブルあったら削除
    If fncTableExist(P_BOX_LIST_WORK_TABLE) Then
        Call dropWorkTable(1)
    End If
    If fncTableExist(P_ROOM_WORK_TABLE) Then
        Call dropWorkTable(2)
    End If
    If fncTableExist(P_COMBO_BOX_WORK_TABLE) Then
        Call dropWorkTable(3)
    End If
'↓INSERT 2013/11/22 MIYAMOTO
    If fncTableExist(P_ORNER_INFO_TABLE) Then
        Call dropWorkTable(4)
    End If
'↑INSERT 2013/11/22 MIYAMOTO
    'ワークテーブル作成
    Call createWorkTable(1)
    Call createWorkTable(2)
    Call createWorkTable(3)
'↓INSERT 2013/11/22 MIYAMOTO
    Call createWorkTable(4)
'↑INSERT 2013/11/22 MIYAMOTO
    
'↓INSERT 2013/11/22 MIYAMOTO
    'EXCEL出力用ワークテーブル、クエリー削除
    If fncTableExist(P_EXCEL_WORK_TABLE) Then
        Call subDropWorkTable
    End If
    Call subDropQuery
'↑INSERT 2013/11/22 MIYAMOTO
    
    gstrUserName = MSZZ000.LsGetUserName
    
    Me.cmd_新規追加.Enabled = False
'↓INSERT 2013/11/22 MIYAMOTO
    Me.cmd_配置図.Enabled = False
    Me.cmd_Excel.Enabled = False
    Me.chk_自動発番可否.VALUE = 0
    Me.chk_自動発番可否.Enabled = False
'↑INSERT 2013/11/22 MIYAMOTO
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ACCESSテーブル存在チェック
'        MODULE_ID        :fncTableExist
'        NOTE             :IN strTableName = テーブル名
'        RETURN           :True=存在する False=存在しない
'        CREATE_DATE      :2008/06/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTableExist(strTableName As String) As Boolean

    Dim tdf As TableDef

    fncTableExist = False

    For Each tdf In pobjDB.TableDefs
        If tdf.NAME = strTableName Then
            fncTableExist = True
            Exit For
        End If
    Next tdf
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル削除
'        MODULE_ID        :dropWorkTable
'        NOTE             :anintSqlFlg   1→一覧画面用ワークテーブル
'                         :              2→ダイアログ（部屋割付一覧）用ワークテーブル
'                         :              3→ダイアログ（部屋割付選択コンボ）用ワークテーブル
'                         :補足※ダイアログのRecordSource代入（２つ）は
'                         :      ダイアログの方で代入する。ここでは行わない。
'        CREATE_DATE      :2008/06/24
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub dropWorkTable(anintSqlFlg As Integer)
    
    On Error GoTo Exception

    If (anintSqlFlg = 1) Then
        Me.ボックス一覧.Form.RecordSource = ""
        Me.ボックス一覧.Form.Requery
        pobjDB.Execute "DROP TABLE " & P_BOX_LIST_WORK_TABLE & " ;", dbFailOnError
    ElseIf (anintSqlFlg = 2) Then
        pobjDB.Execute "DROP TABLE " & P_ROOM_WORK_TABLE & " ;", dbFailOnError
    ElseIf (anintSqlFlg = 3) Then
        pobjDB.Execute "DROP TABLE " & P_COMBO_BOX_WORK_TABLE & " ;", dbFailOnError
'↓INSERT 2013/11/22 MIYAMOTO
    ElseIf (anintSqlFlg = 4) Then
        Me.オーナー情報.Form.RecordSource = ""
        Me.オーナー情報.Form.Requery
        pobjDB.Execute "DROP TABLE " & P_ORNER_INFO_TABLE & " ;", dbFailOnError
'↑INSERT 2013/11/22 MIYAMOTO
    End If

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "dropWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル作成
'        MODULE_ID        :createWorkTable
'        NOTE             :anintSqlFlg   1→一覧画面用ワークテーブル
'                         :              2→ダイアログ（部屋割付一覧）用ワークテーブル
'                         :              3→ダイアログ（部屋割付選択コンボ）用ワークテーブル
'                         :補足※ダイアログのRecordSource代入（２つ）は
'                         :      ダイアログの方で代入する。ここでは行わない。
'        CREATE_DATE      :2008/06/24
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub createWorkTable(anintSqlFlg As Integer)

    Dim strSQL As String
    
    On Error GoTo Exception
    
    If (anintSqlFlg = 1) Then
        '一覧画面用ワークテーブル
        strSQL = ""
        strSQL = strSQL & "CREATE TABLE " & P_BOX_LIST_WORK_TABLE & " ("
'↓UPDATE 2013/11/22 MIYAMOTO
'        strSQL = strSQL & " BOXNO                VARCHAR(16) "
'↓UPDATE 2014/01/17 MIYAMOTO
'        strSQL = strSQL & " BOXNO                VARCHAR(19) "
        strSQL = strSQL & " BOXNO                VARCHAR(16) "
        strSQL = strSQL & ",EDANO                VARCHAR(2) "
'↑UPDATE 2014/01/17 MIYAMOTO
'↑UPDATE 2013/11/22 MIYAMOTO
        strSQL = strSQL & ",TYPE                 VARCHAR(20) "
        strSQL = strSQL & ",TYPE_VARUE           VARCHAR(5) "
        strSQL = strSQL & ",INTYPE               VARCHAR(20) "
        strSQL = strSQL & ",INTYPE_VARUE         VARCHAR(5) "
        strSQL = strSQL & ",STATUS               VARCHAR(50) "
        strSQL = strSQL & ",SIZE                 VARCHAR(20) "
        strSQL = strSQL & ",SIZE_VARUE           VARCHAR(20) "
        strSQL = strSQL & ",DOOR                 VARCHAR(50) "
        strSQL = strSQL & ",DOOR_VARUE           VARCHAR(5) "
        strSQL = strSQL & ",DOORPSTN             VARCHAR(20) "
        strSQL = strSQL & ",DOORPSTN_VARUE       VARCHAR(5) "
        strSQL = strSQL & ",DOORTYPE             VARCHAR(20) "
        strSQL = strSQL & ",DOORTYPE_VARUE       VARCHAR(5) "
        strSQL = strSQL & ",STOCKDATE            VARCHAR(20) "
        strSQL = strSQL & ",SETDATE              VARCHAR(8) "
        strSQL = strSQL & ",SETFLOOR             VARCHAR(20) "
        strSQL = strSQL & ",SETFLOOR_VARUE       VARCHAR(5) "
        strSQL = strSQL & ",CNO                  VARCHAR(255) "
        strSQL = strSQL & ",BIKON                VARCHAR(60) "
'↓INSERT 2013/11/22 MIYAMOTO
        strSQL = strSQL & ",BKTSU                VARCHAR(2) "
        strSQL = strSQL & ",KNKBN                VARCHAR(20) "
        strSQL = strSQL & ",KNKBN_VALUE          VARCHAR(1) "
        strSQL = strSQL & ",CTBNR                VARCHAR(20) "
        strSQL = strSQL & ",CTBNR_VALUE          VARCHAR(2) "
        strSQL = strSQL & ",CTYTO                VARCHAR(20) "
        strSQL = strSQL & ",CTYTO_VALUE          VARCHAR(2) "
        strSQL = strSQL & ",KNKAK                VARCHAR(20) "
        strSQL = strSQL & ",KNKAK_VALUE          VARCHAR(1) "
        strSQL = strSQL & ",SHSHU                VARCHAR(20) "
        strSQL = strSQL & ",SHSHU_VALUE          VARCHAR(2) "
        strSQL = strSQL & ",KRKBN                VARCHAR(20) "
        strSQL = strSQL & ",KRKBN_VALUE          VARCHAR(2) "
        strSQL = strSQL & ",OLDBN                VARCHAR(16) "
'↑INSERT 2013/11/22 MIYAMOTO
        strSQL = strSQL & ")"
        pobjDB.Execute strSQL, dbFailOnError
        Me.ボックス一覧.Form.RecordSource = "SELECT * FROM " & P_BOX_LIST_WORK_TABLE & " ORDER BY BOXNO,SETDATE;"
        Me.ボックス一覧.Form.Requery
    ElseIf (anintSqlFlg = 2) Then
        'ダイアログ（部屋割付一覧）用ワークテーブル
        strSQL = ""
        strSQL = strSQL & "CREATE TABLE " & P_ROOM_WORK_TABLE & " ("
        strSQL = strSQL & " TEXTDATA     VARCHAR(100) "
        strSQL = strSQL & ",CNO          INTEGER "
        strSQL = strSQL & ",SIZE         VARCHAR(20) "
        strSQL = strSQL & ")"
        pobjDB.Execute strSQL, dbFailOnError
    ElseIf (anintSqlFlg = 3) Then
        'ダイアログ（部屋割付選択コンボ）用ワークテーブル
        strSQL = ""
        strSQL = strSQL & "CREATE TABLE " & P_COMBO_BOX_WORK_TABLE & " ("
        strSQL = strSQL & " TEXTDATA     VARCHAR(100) "
        strSQL = strSQL & ",CNO          INTEGER "
        strSQL = strSQL & ",SIZE         VARCHAR(20) "
        strSQL = strSQL & ")"
        pobjDB.Execute strSQL, dbFailOnError
'↓INSERT 2013/11/22 MIYAMOTO
    ElseIf (anintSqlFlg = 4) Then
        'オーナー情報表示用ワークテーブル
        strSQL = ""
        strSQL = strSQL & "CREATE TABLE " & P_ORNER_INFO_TABLE & " ("
        strSQL = strSQL & " BUMOC                VARCHAR(1) "
        strSQL = strSQL & ",BUKHC                VARCHAR(6) "
        strSQL = strSQL & ",SHIRC                VARCHAR(6) "
        strSQL = strSQL & ",SHIRN                VARCHAR(25) "
        strSQL = strSQL & ",KEIYD                VARCHAR(8) "
        strSQL = strSQL & ",KAIYD                VARCHAR(8) "
        strSQL = strSQL & ",KANRI                VARCHAR(20) "
        strSQL = strSQL & ",KANRI_VALUE          VARCHAR(1) "
        strSQL = strSQL & ")"
        pobjDB.Execute strSQL, dbFailOnError
        Me.オーナー情報.Form.RecordSource = "SELECT * FROM " & P_ORNER_INFO_TABLE & " ORDER BY SHIRC,KEIYD;"
        Me.オーナー情報.Form.Requery
'↑INSERT 2013/11/22 MIYAMOTO
    End If
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "createWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルデータ削除
'        MODULE_ID        :deleteWorkTable
'        NOTE             :anboolDelBoxList   True→一覧表示用ワーク削除
'                         :anboolDelRoom      True→詳細の部屋割付用ワーク削除
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub deleteWorkTable(anboolDelBoxList As Boolean, anboolDelRoom As Boolean)
    
    Dim strSQL              As String
    
    On Error GoTo Exception
    
    If (anboolDelBoxList) Then
        Call dropWorkTable(1)
        Call createWorkTable(1)
'↓INSERT 2013/11/22 MIYAMOTO
        Call dropWorkTable(4)
        Call createWorkTable(4)
'↑INSERT 2013/11/22 MIYAMOTO
    End If
    If (anboolDelRoom) Then
        Call dropWorkTable(2)
        Call createWorkTable(2)
    End If
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "deleteWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧作成
'        MODULE_ID        :createBoxData
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub createBoxData()

    Dim boolBeforeCheck As Boolean
    Dim strZyoukyou     As String
    Dim strSQL          As String

    On Error GoTo Exception
    
    boolBeforeCheck = ((Me.cmb_ヤード.ListIndex <> -1) And (Me.cmb_状況.ListIndex <> -1))
    
    If boolBeforeCheck Then
        
        'ワークテーブルのデータ削除
        Call deleteWorkTable(True, False)
        
        '状況によって検索する
        strZyoukyou = Nz(Me.cmb_状況.Column(1, Me.cmb_状況.ListIndex))
        
        If strZyoukyou = "0" Then
            Call getSelectBoxData(createBoxDataSQL09, "設置中")
'+++++++++++++++++++++++++++++++++++++++++++++++
'                ●createBoxDataSQL06とcreateBoxDataSQL07は
'                ●次フェーズで対応する。
'            Call getSelectBoxData(createBoxDataSQL06)
'            Call getSelectBoxData(createBoxDataSQL07)
'+++++++++++++++++++++++++++++++++++++++++++++++
        Else
            If strZyoukyou = "1" Then
                '割付済
                Call getSelectBoxData(createBoxDataSQL01, "割付済")
                Call getSelectBoxData(createBoxDataSQL03, "割付済（親）")
                Call getSelectBoxData(createBoxDataSQL04, "割付済（子）")
            ElseIf strZyoukyou = "2" Then
                '未割付
                Call getSelectBoxData(createBoxDataSQL02, "未割付")
            ElseIf strZyoukyou = "3" Then
                '撤去
                Call getSelectBoxData(createBoxDataSQL05, "撤去")
            Else
                MsgBox "状況の値に問題があります。"
            End If
        End If
        
        Me.ボックス一覧.Form.Requery
'↓INSERT 2013/11/22 MIYAMOTO
        Call getOrnerInfo
        Me.オーナー情報.Form.Requery
'↑INSERT 2013/11/22 MIYAMOTO
        
    Else
        
    End If
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "CreateBoxData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧表時用ワークテーブル更新
'        MODULE_ID        :getSelectBoxData
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getSelectBoxData(strSQL As String, anStrStatus As String)
    
    Dim objAdoDbConnection   As Object
    Dim rsData               As Object       'ADODB.Recordset
    Dim rsBoxData            As Recordset
    Dim sSQL                 As String
    
    
    On Error GoTo Exception
    
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)

    While Not rsData.EOF
        sSQL = ""
        sSQL = sSQL & "SELECT * FROM " & P_BOX_LIST_WORK_TABLE
'↓UPDATE 2013/11/22 MIYAMOTO
'        sSQL = sSQL & " WHERE BOXNO   = '" & rsData.Fields("BOXNO") & "'"
'↓UPDATE 2014/01/17 MIYAMOTO
'        sSQL = sSQL & " WHERE BOXNO   = '" & rsData.Fields("BOXNO") & " " & rsData.Fields("EDANO") & "'"
        sSQL = sSQL & " WHERE BOXNO   = '" & rsData.Fields("BOXNO") & "'"
        sSQL = sSQL & " AND   EDANO   = '" & rsData.Fields("EDANO") & "'"
'↑UPDATE 2014/01/17 MIYAMOTO
'↑UPDATE 2013/11/22 MIYAMOTO
        sSQL = sSQL & " AND   SETDATE = '" & rsData.Fields("SETDATE") & "'"
        Set rsBoxData = pobjDB.OpenRecordset(sSQL, dbOpenDynaset)
        If rsBoxData.EOF Then
            rsBoxData.AddNew
'↓UPDATE 2013/11/22 MIYAMOTO
'            rsBoxData.Fields("BOXNO") = rsData.Fields("BOXNO")
'↓UPDATE 2014/01/17 MIYAMOTO
'            rsBoxData.Fields("BOXNO") = rsData.Fields("BOXNO") & " " & rsData.Fields("EDANO")
            rsBoxData.Fields("BOXNO") = rsData.Fields("BOXNO")
            rsBoxData.Fields("EDANO") = rsData.Fields("EDANO")
'↑UPDATE 2014/01/17 MIYAMOTO
'↑UPDATE 2013/11/22 MIYAMOTO
            rsBoxData.Fields("TYPE") = rsData.Fields("TYPE")
            rsBoxData.Fields("TYPE_VARUE") = rsData.Fields("TYPE_VARUE")
            rsBoxData.Fields("INTYPE") = rsData.Fields("INTYPE")
            rsBoxData.Fields("INTYPE_VARUE") = rsData.Fields("INTYPE_VARUE")
            rsBoxData.Fields("SIZE") = rsData.Fields("SIZE")
            rsBoxData.Fields("SIZE_VARUE") = rsData.Fields("SIZE_VARUE")
            rsBoxData.Fields("DOOR") = rsData.Fields("DOOR") & "," & rsData.Fields("DOORPSTN") & "," & rsData.Fields("DOORTYPE")
            rsBoxData.Fields("DOOR_VARUE") = rsData.Fields("DOOR_VARUE")
            rsBoxData.Fields("DOORPSTN_VARUE") = rsData.Fields("DOORPSTN_VARUE")
            rsBoxData.Fields("DOORPSTN") = rsData.Fields("DOORPSTN")
            rsBoxData.Fields("DOORTYPE_VARUE") = rsData.Fields("DOORTYPE_VARUE")
            rsBoxData.Fields("DOORTYPE") = rsData.Fields("DOORTYPE")
            rsBoxData.Fields("STOCKDATE") = rsData.Fields("STOCKDATE")
            rsBoxData.Fields("STATUS") = anStrStatus
'↓INSERT 2013/11/22 MIYAMOTO
            rsBoxData.Fields("BKTSU") = rsData.Fields("BKTSU")
            rsBoxData.Fields("KNKBN") = rsData.Fields("KNKBN")
            rsBoxData.Fields("KNKBN_VALUE") = rsData.Fields("KNKBN_VALUE")
            rsBoxData.Fields("CTBNR") = rsData.Fields("CTBNR")
            rsBoxData.Fields("CTBNR_VALUE") = rsData.Fields("CTBNR_VALUE")
            rsBoxData.Fields("CTYTO") = rsData.Fields("CTYTO")
            rsBoxData.Fields("CTYTO_VALUE") = rsData.Fields("CTYTO_VALUE")
            rsBoxData.Fields("KNKAK") = rsData.Fields("KNKAK")
            rsBoxData.Fields("KNKAK_VALUE") = rsData.Fields("KNKAK_VALUE")
            rsBoxData.Fields("SHSHU") = rsData.Fields("SHSHU")
            rsBoxData.Fields("SHSHU_VALUE") = rsData.Fields("SHSHU_VALUE")
            rsBoxData.Fields("KRKBN") = rsData.Fields("KRKBN")
            rsBoxData.Fields("KRKBN_VALUE") = rsData.Fields("KRKBN_VALUE")
            rsBoxData.Fields("OLDBN") = rsData.Fields("OLDBN")
'↑INSERT 2013/11/22 MIYAMOTO
            If rsData.Fields("CNO") = "" Then
                rsBoxData.Fields("CNO") = "（未割付）"
            Else
                rsBoxData.Fields("CNO") = Format(Nz(rsData.Fields("CNO")), "000000")
            End If
            rsBoxData.Fields("SETDATE") = rsData.Fields("SETDATE")
            rsBoxData.Fields("SETFLOOR") = rsData.Fields("SETFLOOR")
            rsBoxData.Fields("SETFLOOR_VARUE") = rsData.Fields("SETFLOOR_VARUE")
            rsBoxData.Fields("BIKON") = rsData.Fields("BIKON")
        Else
            rsBoxData.Edit
            rsBoxData.Fields("CNO") = rsBoxData.Fields("CNO") & "," & Format(Nz(rsData.Fields("CNO")), "000000")
        End If
        rsBoxData.UPDATE
        rsBoxData.Close
        Set rsBoxData = Nothing
        rsData.MoveNext
    Wend

    rsData.Close
    Set rsData = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing
    
    Exit Function

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsBoxData Is Nothing Then: rsBoxData.Close: Set rsBoxData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "getBoxSelectData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :割付済（親子関係なし）
'        MODULE_ID        :createBoxDataSQL01
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL01() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CNTA_MAST.CNTA_NO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "      ,CBRA_TRAN,CNTA_MAST "
    strSQL = strSQL & " where NOT EXISTS (select * from CBYS_TRAN CBYS01 where CBYS01.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBYS01.CBYST_STATUS = 0 and CBYS01.CBYST_SYUBETSU = 1) "
    strSQL = strSQL & "   and CBRA_TRAN.CBRAT_CNO = CNTA_MAST.CNTA_NO "
    strSQL = strSQL & "   and CBRA_TRAN.CBRAT_YCODE = CNTA_MAST.CNTA_CODE "
    strSQL = strSQL & "   and CBOX_MAST.CBOXM_BOXNO = CBRA_TRAN.CBRAT_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBOX_MAST.CBOXM_EDANO = CBRA_TRAN.CBRAT_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL01 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :未割付
'        MODULE_ID        :createBoxDataSQL02
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL02() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NULL as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " where NOT EXISTS (select * from CBRA_TRAN where CBRA_TRAN.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO) "
'    strSQL = strSQL & "   and NOT EXISTS (select * from CBYS_TRAN CBYS01 where CBYS01.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBYS01.CBYST_STATUS = 0 and CBYS01.CBYST_SYUBETSU = 1) "
    strSQL = strSQL & " where NOT EXISTS (select * from CBRA_TRAN where CBRA_TRAN.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBRA_TRAN.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO) "
    strSQL = strSQL & "   and NOT EXISTS (select * from CBYS_TRAN CBYS01 where CBYS01.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBYS01.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO and CBYS01.CBYST_STATUS = 0 and CBYS01.CBYST_SYUBETSU = 1) "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL02 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :割付済（親）
'        MODULE_ID        :createBoxDataSQL03
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL03() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and w01.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "      ,(select CBYST_BOXNO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as w02 "
    strSQL = strSQL & "      ,(select CBYST_BOXNO,CBYST_EDANO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as w02 "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " where w02.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and w02.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL03 = strSQL
    
End Function


'==============================================================================*
'
'        MODULE_NAME      :割付済（子）
'        MODULE_ID        :createBoxDataSQL04
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL04() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBRA_W.CNTA_NO as CNO "
    strSQL = strSQL & "     ,CBYS_W.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,CBYS_W.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_W.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_W.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from  CBOX_MAST OYA_BOX,CBYS_TRAN OYA_CBYS,CBYS_TRAN "
    strSQL = strSQL & "      ,CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN  "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "            (SELECT CBYS_TRAN.CBYST_BOXNO,CBYS_TRAN.CBYST_SETDATE,CBYS_TRAN.CBYST_SETFLOOR, "
    strSQL = strSQL & "            (SELECT CBYS_TRAN.CBYST_BOXNO,CBYS_TRAN.CBYST_EDANO,CBYS_TRAN.CBYST_SETDATE,CBYS_TRAN.CBYST_SETFLOOR, "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "                    NAME_MAST7.NAME_RYAK,CBYS_TRAN.CBYST_BIKON "
    strSQL = strSQL & "               FROM CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7  "
    strSQL = strSQL & "                        ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR "
    strSQL = strSQL & "                       AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "              WHERE CBYS_TRAN.CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as CBYS_W  "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "       ON CBYS_W.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "       ON CBYS_W.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBYS_W.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN  "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "            (SELECT CBRA_TRAN.CBRAT_BOXNO,CNTA_MAST.CNTA_NO "
    strSQL = strSQL & "            (SELECT CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CNTA_MAST.CNTA_NO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "               FROM CBRA_TRAN,CNTA_MAST "
    strSQL = strSQL & "              WHERE CBRA_TRAN.CBRAT_CNO = CNTA_MAST.CNTA_NO "
    strSQL = strSQL & "                AND CBRA_TRAN.CBRAT_YCODE = CNTA_MAST.CNTA_CODE) as CBRA_W  "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "       ON CBOX_MAST.CBOXM_BOXNO = CBRA_W.CBRAT_BOXNO "
    strSQL = strSQL & "       ON CBOX_MAST.CBOXM_BOXNO = CBRA_W.CBRAT_BOXNO and CBOX_MAST.CBOXM_EDANO = CBRA_W.CBRAT_EDANO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "       "
    strSQL = strSQL & " where CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 1 "
    strSQL = strSQL & "   and OYA_BOX.CBOXM_BOXNO = CBYS_TRAN.CBYST_YCODE "
    strSQL = strSQL & "   and OYA_CBYS.CBYST_BOXNO = OYA_BOX.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and OYA_CBYS.CBYST_EDANO = OYA_BOX.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and OYA_CBYS.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and OYA_CBYS.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (OYA_CBYS.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL04 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :撤去
'        MODULE_ID        :createBoxDataSQL05
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL05() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBRA_RECD.CBRAR_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_RECD.CBYSR_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_RECD.CBYSR_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_RECD.CBYSR_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "      ,CBYS_RECD LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_RECD.CBYSR_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "          LEFT OUTER JOIN CBRA_RECD ON CBYS_RECD.CBYSR_YCODE = CBRA_RECD.CBRAR_YCODE AND CBYS_RECD.CBYSR_BOXNO = CBRA_RECD.CBRAR_BOXNO AND CBYS_RECD.CBYSR_SETDATE = CBRA_RECD.CBRAR_SETDATE "
    strSQL = strSQL & "          LEFT OUTER JOIN CBRA_RECD ON CBYS_RECD.CBYSR_YCODE = CBRA_RECD.CBRAR_YCODE AND CBYS_RECD.CBYSR_BOXNO = CBRA_RECD.CBRAR_BOXNO AND CBYS_RECD.CBYSR_EDANO = CBRA_RECD.CBRAR_EDANO AND CBYS_RECD.CBYSR_SETDATE = CBRA_RECD.CBRAR_SETDATE "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " where CBYS_RECD.CBYSR_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_RECD.CBYSR_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_RECD.CBYSR_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_RECD.CBYSR_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL05 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :移動する予定
'        MODULE_ID        :createBoxDataSQL06
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL06() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and w01.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "      ,(select CBYST_BOXNO from CBYS_TRAN where CBYST_STATUS = 1 and CBYS_TRAN.CBYST_SYUBETSU = 0) as w02 "
    strSQL = strSQL & "      ,(select CBYST_BOXNO,CBYST_EDANO from CBYS_TRAN where CBYST_STATUS = 1 and CBYS_TRAN.CBYST_SYUBETSU = 0) as w02 "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " where w02.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and w02.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL06 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :移動してくる予定
'        MODULE_ID        :createBoxDataSQL07
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL07() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and w01.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "      ,(select CBYST_BOXNO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 0) as w02 "
    strSQL = strSQL & "      ,(select CBYST_BOXNO,CBYST_EDANO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 0) as w02 "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " where w02.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and w02.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 1 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL07 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :設置中
'        MODULE_ID        :createBoxDataSQL09
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createBoxDataSQL09() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
'↑INSERT 2013/11/22 MIYAMOTO
'↓UPDATE 2013/11/22 MIYAMOTO
'    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'    strSQL = strSQL & "          LEFT OUTER JOIN (select CBYST_BOXNO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as w02 on w02.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and w01.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBYST_BOXNO,CBYST_EDANO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as w02 on w02.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and w02.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & " where CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
'↓INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
'↑INSERT 2013/11/22 MIYAMOTO
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
'↓UPDATE 2013/11/22 MIYAMOTO

'↓INSERT 2016/12/13 M.HONDA INS
    If Me.txt_検索 <> "" Then
        strSQL = strSQL & "   AND ( CBOX_MAST.CBOXM_BOXNO like '%" & txt_検索 & "%'"
        strSQL = strSQL & "   or  CBYST_BIKON like '%" & txt_検索 & "%')"
    End If
'↑INSERT 2016/12/13 M.HONDA INS

'    strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
'↑UPDATE 2013/11/22 MIYAMOTO
    strSQL = strSQL & " ; "
    
    createBoxDataSQL09 = strSQL
    
End Function

'↓INSERT 2013/11/22 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :オーナー情報表示用ワークテーブル更新
'        MODULE_ID        :getOrnerInfo
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getOrnerInfo()
    
    Dim objAdoDbConnection   As Object
    Dim rsData               As Object
    Dim rsOrnerData          As Recordset
    Dim strDate              As String
    Dim strSQL               As String
    Dim sSQL                 As String

On Error GoTo Exception

    '初期値設定
    strDate = Format(DATE, "yyyymmdd")

    strSQL = ""
'    strSQL = strSQL & "SELECT HKEI_TRAN.HKEIT_BUMOC AS BUMOC, "
'    strSQL = strSQL & "HKEI_TRAN.HKEIT_BUKHC AS BUKHC, "
'    strSQL = strSQL & "HKEI_TRAN.HKEIT_SHIRC AS SHIRC, "
'    strSQL = strSQL & "SHIR_MAST.SHIRM_SHIRN AS SHIRN, "
'    strSQL = strSQL & "HKEI_TRAN.HKEIT_KEIYD AS KEIYD, "
'    strSQL = strSQL & "HKEI_TRAN.HKEIT_KAIYD AS KAIYD, "
'    strSQL = strSQL & "CODE_TABL.CODET_NAMEN AS KANRI, "
'    strSQL = strSQL & "HKEI_TRAN.HKEIT_KANRI AS KANRI_VALUE "
'    strSQL = strSQL & "FROM (HKEI_TRAN INNER JOIN SHIR_MAST ON HKEI_TRAN.HKEIT_SHIRC = SHIR_MAST.SHIRM_SHIRC) INNER JOIN CODE_TABL ON HKEI_TRAN.HKEIT_KANRI = CODE_TABL.CODET_CODEC "
'    strSQL = strSQL & "WHERE (((HKEI_TRAN.HKEIT_KAIYD) Is Null Or (HKEI_TRAN.HKEIT_KAIYD)=' ' Or (HKEI_TRAN.HKEIT_KAIYD)>='" & strDate & "') AND ((HKEI_TRAN.HKEIT_BUMOC)='H') AND ((HKEI_TRAN.HKEIT_BUKHC)='" & Me.txt_借上物件NO & "') AND ((SHIR_MAST.SHIRM_BUMOC)='ｦ') AND ((CODE_TABL.CODET_SIKBC)='221')) "
'    strSQL = strSQL & "ORDER BY HKEI_TRAN.HKEIT_SHIRC"
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    HKEI_TRAN.HKEIT_BUMOC AS BUMOC, "
    strSQL = strSQL & "    HKEI_TRAN.HKEIT_BUKHC AS BUKHC, "
    strSQL = strSQL & "    HKEI_TRAN.HKEIT_SHIRC AS SHIRC, "
    strSQL = strSQL & "    SHIR_MAST.SHIRM_SHIRN AS SHIRN, "
    strSQL = strSQL & "    HKEI_TRAN.HKEIT_KEIYD AS KEIYD, "
    strSQL = strSQL & "    HKEI_TRAN.HKEIT_KAIYD AS KAIYD, "
    strSQL = strSQL & "    CODE_TABL.CODET_NAMEN AS KANRI, "
    strSQL = strSQL & "    HKEI_TRAN.HKEIT_KANRI AS KANRI_VALUE "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    HKEI_TRAN "
    strSQL = strSQL & "        INNER JOIN SHIR_MAST ON HKEI_TRAN.HKEIT_SHIRC = SHIR_MAST.SHIRM_SHIRC AND SHIR_MAST.SHIRM_BUMOC='ｦ' "
    strSQL = strSQL & "        INNER JOIN CODE_TABL ON HKEI_TRAN.HKEIT_KANRI = CODE_TABL.CODET_CODEC AND CODE_TABL.CODET_SIKBC='221' "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    (HKEI_TRAN.HKEIT_KAIYD Is Null Or HKEI_TRAN.HKEIT_KAIYD=' ' Or HKEI_TRAN.HKEIT_KAIYD>='" & strDate & "') "
    strSQL = strSQL & "AND (HKEI_TRAN.HKEIT_BUKHC='" & Me.txt_借上物件NO & "') "
    strSQL = strSQL & "AND (HKEI_TRAN.HKEIT_KANRI!='8') "
    strSQL = strSQL & "ORDER BY "
    strSQL = strSQL & "    HKEI_TRAN.HKEIT_SHIRC "
    Set objAdoDbConnection = MSZZ025.ADODB_Connection()
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

    While Not rsData.EOF
        'ワークテーブルの存在チェック
        sSQL = ""
        sSQL = sSQL & "SELECT * FROM " & P_ORNER_INFO_TABLE
        sSQL = sSQL & " WHERE BUMOC = '" & rsData.Fields("BUMOC") & "'"
        sSQL = sSQL & " AND BUKHC = '" & rsData.Fields("BUKHC") & "'"
        Set rsOrnerData = pobjDB.OpenRecordset(sSQL, dbOpenDynaset)
        With rsOrnerData
            If .EOF Then
                .AddNew
                .Fields("BUMOC") = rsData.Fields("BUMOC")
                .Fields("BUKHC") = rsData.Fields("BUKHC")
                .Fields("SHIRC") = rsData.Fields("SHIRC")
                .Fields("SHIRN") = rsData.Fields("SHIRN")
                .Fields("KEIYD") = rsData.Fields("KEIYD")
                .Fields("KAIYD") = rsData.Fields("KAIYD")
'                .Fields("入庫日") = Left(rsData.Fields("STOCKDATE"), 4) & "/" & Mid(rsData.Fields("STOCKDATE"), 5, 2) & "/" & _
'                                    Right(rsData.Fields("STOCKDATE"), 2)
                .Fields("KANRI") = rsData.Fields("KANRI")
                .Fields("KANRI_VALUE") = rsData.Fields("KANRI_VALUE")
                .UPDATE
            End If
            .Close
            Set rsOrnerData = Nothing
            rsData.MoveNext
        End With
    Wend

    rsData.Close
    Set rsData = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    Exit Function

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsOrnerData Is Nothing Then: rsOrnerData.Close: Set rsOrnerData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "getOrnerInfo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :EXCELに出力するデータをワークテーブルへ作成
'        MODULE_ID        :subOutputExcel
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subOutputExcel()

    Dim objAdoDbConnection  As Object
    Dim strZyoukyou         As String

On Error GoTo Exception

    'コネクション接続
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)

    'クエリ作成
    Call subCreateQuery

    'ワークテーブル削除
    Call subDropWorkTable

    'ワークテーブル作成
    Set pobjTableDef = pobjDB.CreateTableDef(P_EXCEL_WORK_TABLE)
    Call FieldAppend
    pobjDB.TableDefs.Append pobjTableDef

    '状況によって検索する
    strZyoukyou = Nz(Me.cmb_状況.Column(1, Me.cmb_状況.ListIndex))

    'ワークテーブル更新
    If strZyoukyou = "0" Then
        Call getExcelData(createExcelDataSQL09, "設置中")
    ElseIf strZyoukyou = "1" Then
        '割付済
        Call getExcelData(createExcelDataSQL01, "割付済")
        Call getExcelData(createExcelDataSQL03, "割付済（親）")
        Call getExcelData(createExcelDataSQL04, "割付済（子）")
    ElseIf strZyoukyou = "2" Then
        '未割付
        Call getExcelData(createExcelDataSQL02, "未割付")
    ElseIf strZyoukyou = "3" Then
        '撤去
        Call getExcelData(createExcelDataSQL05, "撤去")
    Else
        MsgBox "状況の値に問題があります。"
    End If

    Exit Sub

Exception:
    If Not objAdoDbConnection Is Nothing Then
        'コネクション解放
        objAdoDbConnection.Close
        Set objAdoDbConnection = Nothing
    End If
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "subOutputExcel" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :クエリ作成
'        MODULE_ID        :subCreateQuery
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreateQuery()

On Error GoTo Exception

    Dim strSQL As String

    Set pobjQueryDef = pobjDB.CreateQueryDef(P_QUERY_NAME)
    'count列以外をselectする
    strSQL = ""
    strSQL = strSQL & "SELECT ヤードコード, ヤード名,状況,コンテナ番号,枝番,分割数"
    strSQL = strSQL & ",所有,購入,入庫日,設置用途,分類,サイズ,タイプ"
    strSQL = strSQL & ",内装,ドア,内通路,設置日,階層,割付部屋番号,備考"
    strSQL = strSQL & " FROM " & P_EXCEL_WORK_TABLE
    strSQL = strSQL & " ORDER BY ヤードコード,コンテナ番号,枝番 "
    pobjQueryDef.sql = strSQL

    strSQL = ""
    strSQL = strSQL & "SELECT "
    
    
Exception:
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :クエリ削除
'        MODULE_ID        :subDropQuery
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDropQuery()

On Error GoTo Exception

    doCmd.DeleteObject acQuery, P_QUERY_NAME

Exception:
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル削除
'        MODULE_ID        :subDropWorkTable
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDropWorkTable()

On Error GoTo Exception

    doCmd.DeleteObject acTable, P_EXCEL_WORK_TABLE

Exception:
    Err.Clear
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用ワークテーブル列作成
'        MODULE_ID        :FieldAppend
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :

'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FieldAppend()

    Dim intCount    As Integer

    With pobjTableDef
        Call .Fields.Append(.CreateField("ヤードコード", dbText, 36))
        Call .Fields.Append(.CreateField("ヤード名", dbText, 50))
        Call .Fields.Append(.CreateField("コンテナ番号", dbText, 16))
        Call .Fields.Append(.CreateField("枝番", dbText, 2))
        Call .Fields.Append(.CreateField("タイプ", dbText, 6))
        Call .Fields.Append(.CreateField("内装", dbText, 6))
        Call .Fields.Append(.CreateField("サイズ", dbText, 6))
        Call .Fields.Append(.CreateField("ドア", dbText, 50))
        Call .Fields.Append(.CreateField("入庫日", dbText, 10))
        Call .Fields.Append(.CreateField("所有", dbText, 10))
        Call .Fields.Append(.CreateField("購入", dbText, 10))
        Call .Fields.Append(.CreateField("分類", dbText, 10))
        Call .Fields.Append(.CreateField("設置用途", dbText, 10))
        Call .Fields.Append(.CreateField("建築確認", dbText, 10))
        Call .Fields.Append(.CreateField("内通路", dbText, 10))
        Call .Fields.Append(.CreateField("分割数", dbText, 2))
        Call .Fields.Append(.CreateField("設置日temp", dbText, 8))  '設置日(yyyymmdd）
        Call .Fields.Append(.CreateField("設置日", dbText, 10))     '設置日(yyyy/mm/dd)
        Call .Fields.Append(.CreateField("階層", dbText, 6))
        Call .Fields.Append(.CreateField("状況", dbText, 50))
        Call .Fields.Append(.CreateField("割付部屋番号", dbText, 50))
        Call .Fields.Append(.CreateField("備考", dbText, 50))
        Call .Fields.Append(.CreateField("count", dbLong))  'カウンタ（割付部屋番号数をカウントする）
        
        '空文字列の登録を許可する
        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount

    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用ワークテーブル更新
'        MODULE_ID        :getExcelData
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getExcelData(strSQL As String, anStrStatus As String)
    
    Dim objAdoDbConnection   As Object
    Dim rsData               As Object
    Dim rsBoxData            As Recordset
    Dim sSQL                 As String

On Error GoTo Exception

    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)

    While Not rsData.EOF
        'ワークテーブルの存在チェック
        sSQL = ""
        sSQL = sSQL & "SELECT * FROM " & P_EXCEL_WORK_TABLE
        sSQL = sSQL & " WHERE コンテナ番号 = '" & rsData.Fields("BOXNO") & "'"
        sSQL = sSQL & " AND 枝番 = '" & rsData.Fields("EDANO") & "'"
        sSQL = sSQL & " AND 設置日temp = '" & rsData.Fields("SETDATE") & "'"   '設置日temp(テキスト型）
        Set rsBoxData = pobjDB.OpenRecordset(sSQL, dbOpenDynaset)
        With rsBoxData
            If .EOF Then
                .AddNew
                .Fields("count") = 1
                .Fields("ヤードコード") = Format(Nz(rsData.Fields("YARD_CODE")), "000000")
                .Fields("ヤード名") = rsData.Fields("YARD_NAME")
                .Fields("コンテナ番号") = rsData.Fields("BOXNO")
                .Fields("枝番") = rsData.Fields("EDANO")
                .Fields("タイプ") = rsData.Fields("TYPE")
                .Fields("内装") = rsData.Fields("INTYPE")
                .Fields("サイズ") = rsData.Fields("SIZE")
                .Fields("ドア") = rsData.Fields("DOOR") & "," & rsData.Fields("DOORPSTN") & "," & rsData.Fields("DOORTYPE")
                .Fields("入庫日") = Left(rsData.Fields("STOCKDATE"), 4) & "/" & Mid(rsData.Fields("STOCKDATE"), 5, 2) & "/" & _
                                    Right(rsData.Fields("STOCKDATE"), 2)
                .Fields("所有") = rsData.Fields("KRKBN")
                .Fields("購入") = rsData.Fields("KNKBN")
                .Fields("分類") = rsData.Fields("CTBNR")
                .Fields("設置用途") = rsData.Fields("CTYTO")
                .Fields("建築確認") = rsData.Fields("KNKAK")
                .Fields("内通路") = rsData.Fields("SHSHU")
                .Fields("分割数") = rsData.Fields("BKTSU")
                .Fields("状況") = anStrStatus
                If Nz(rsData.Fields("CNO"), "") = "" Then
                    .Fields("割付部屋番号") = "（未割付）"
                Else
                    .Fields("割付部屋番号") = Format(Nz(rsData.Fields("CNO")), "000000")
                End If
                .Fields("設置日temp") = rsData.Fields("SETDATE")  '設置日temp(yyyymmdd)
                .Fields("設置日") = Left(rsData.Fields("SETDATE"), 4) & "/" & Mid(rsData.Fields("SETDATE"), 5, 2) & "/" & _
                                    Right(rsData.Fields("SETDATE"), 2)
                .Fields("階層") = rsData.Fields("SETFLOOR")
                .Fields("備考") = rsData.Fields("BIKON")
                .UPDATE
            Else
                Select Case .Fields("count")  'カウンタ
                Case Is < 5
                    '4つ目までの割付部屋番号は、カンマ区切りで列挙する
                    .Edit
                    .Fields("割付部屋番号") = .Fields("割付部屋番号") & "," & Format(Nz(rsData.Fields("CNO")), "000000")
                    .Fields("count") = .Fields("count") + 1
                    .UPDATE
                Case 5
                    '5つ目の割付部屋番号は、カンマ区切りで列挙し、その番号の後に「他」と入力する
                    .Edit
                    .Fields("割付部屋番号") = .Fields("割付部屋番号") & "他"
                    .Fields("count") = .Fields("count") + 1
                    .UPDATE
                End Select
            End If
            .Close
            Set rsBoxData = Nothing
            rsData.MoveNext
        End With
    Wend

    rsData.Close
    Set rsData = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    Exit Function

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsBoxData Is Nothing Then: rsBoxData.Close: Set rsBoxData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "getExcelData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用(割付済（親子関係なし）)
'        MODULE_ID        :createExcelDataSQL01
'        NOTE             :createBoxDataSQL01を参考に作成
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createExcelDataSQL01() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      YARD_MAST.YARD_CODE as YARD_CODE "
    strSQL = strSQL & "     ,YARD_MAST.YARD_NAME as YARD_NAME "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BOXNO as BOXNO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
    strSQL = strSQL & "     ,CNTA_MAST.CNTA_NO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "      ,CBRA_TRAN,CNTA_MAST "
    strSQL = strSQL & "      ,YARD_MAST "
    strSQL = strSQL & " where NOT EXISTS (select * from CBYS_TRAN CBYS01 where CBYS01.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBYS01.CBYST_STATUS = 0 and CBYS01.CBYST_SYUBETSU = 1) "
    strSQL = strSQL & "   and CBRA_TRAN.CBRAT_CNO = CNTA_MAST.CNTA_NO "
    strSQL = strSQL & "   and CBRA_TRAN.CBRAT_YCODE = CNTA_MAST.CNTA_CODE "
    strSQL = strSQL & "   and CBOX_MAST.CBOXM_BOXNO = CBRA_TRAN.CBRAT_BOXNO "
    strSQL = strSQL & "   and CBOX_MAST.CBOXM_EDANO = CBRA_TRAN.CBRAT_EDANO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = YARD_MAST.YARD_CODE "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
    strSQL = strSQL & " ; "
    
    createExcelDataSQL01 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用(未割付)
'        MODULE_ID        :createExcelDataSQL02
'        NOTE             :createBoxDataSQL02を参考に作成
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createExcelDataSQL02() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      YARD_MAST.YARD_CODE as YARD_CODE "
    strSQL = strSQL & "     ,YARD_MAST.YARD_NAME as YARD_NAME "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BOXNO as BOXNO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
    strSQL = strSQL & "     ,NULL as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "      ,YARD_MAST "
    strSQL = strSQL & " where NOT EXISTS (select * from CBRA_TRAN where CBRA_TRAN.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBRA_TRAN.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO) "
    strSQL = strSQL & "   and NOT EXISTS (select * from CBYS_TRAN CBYS01 where CBYS01.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBYS01.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO and CBYS01.CBYST_STATUS = 0 and CBYS01.CBYST_SYUBETSU = 1) "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = YARD_MAST.YARD_CODE "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
    strSQL = strSQL & " ; "
    
    createExcelDataSQL02 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用(割付済（親）)
'        MODULE_ID        :createExcelDataSQL03
'        NOTE             :createBoxDataSQL03を参考に作成
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createExcelDataSQL03() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      YARD_MAST.YARD_CODE as YARD_CODE "
    strSQL = strSQL & "     ,YARD_MAST.YARD_NAME as YARD_NAME "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BOXNO as BOXNO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
    strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and w01.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "      ,(select CBYST_BOXNO,CBYST_EDANO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as w02 "
    strSQL = strSQL & "      ,YARD_MAST "
    strSQL = strSQL & " where w02.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "   and w02.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = YARD_MAST.YARD_CODE "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
    strSQL = strSQL & " ; "
    
    createExcelDataSQL03 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用(割付済（子）)
'        MODULE_ID        :createExcelDataSQL04
'        NOTE             :createBoxDataSQL04を参考に作成
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createExcelDataSQL04() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      YARD_MAST.YARD_CODE as YARD_CODE "
    strSQL = strSQL & "     ,YARD_MAST.YARD_NAME as YARD_NAME "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BOXNO as BOXNO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
    strSQL = strSQL & "     ,CBRA_W.CNTA_NO as CNO "
    strSQL = strSQL & "     ,CBYS_W.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,CBYS_W.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_W.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_W.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from  CBOX_MAST OYA_BOX,CBYS_TRAN OYA_CBYS,CBYS_TRAN "
    strSQL = strSQL & "      ,CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
    strSQL = strSQL & "          LEFT OUTER JOIN  "
    strSQL = strSQL & "            (SELECT CBYS_TRAN.CBYST_BOXNO,CBYS_TRAN.CBYST_EDANO,CBYS_TRAN.CBYST_SETDATE,CBYS_TRAN.CBYST_SETFLOOR, "
    strSQL = strSQL & "                    NAME_MAST7.NAME_RYAK,CBYS_TRAN.CBYST_BIKON "
    strSQL = strSQL & "               FROM CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7  "
    strSQL = strSQL & "                        ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR "
    strSQL = strSQL & "                       AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "              WHERE CBYS_TRAN.CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as CBYS_W  "
    strSQL = strSQL & "       ON CBYS_W.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and CBYS_W.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "          LEFT OUTER JOIN  "
    strSQL = strSQL & "            (SELECT CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CNTA_MAST.CNTA_NO "
    strSQL = strSQL & "               FROM CBRA_TRAN,CNTA_MAST "
    strSQL = strSQL & "              WHERE CBRA_TRAN.CBRAT_CNO = CNTA_MAST.CNTA_NO "
    strSQL = strSQL & "                AND CBRA_TRAN.CBRAT_YCODE = CNTA_MAST.CNTA_CODE) as CBRA_W  "
    strSQL = strSQL & "       ON CBOX_MAST.CBOXM_BOXNO = CBRA_W.CBRAT_BOXNO and CBOX_MAST.CBOXM_EDANO = CBRA_W.CBRAT_EDANO "
    strSQL = strSQL & "      ,YARD_MAST "
    strSQL = strSQL & " where CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 1 "
    strSQL = strSQL & "   and OYA_BOX.CBOXM_BOXNO = CBYS_TRAN.CBYST_YCODE "
    strSQL = strSQL & "   and OYA_CBYS.CBYST_BOXNO = OYA_BOX.CBOXM_BOXNO "
    strSQL = strSQL & "   and OYA_CBYS.CBYST_EDANO = OYA_BOX.CBOXM_EDANO "
    strSQL = strSQL & "   and OYA_CBYS.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and OYA_CBYS.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (OYA_CBYS.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = YARD_MAST.YARD_CODE "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
    strSQL = strSQL & " ; "
    
    createExcelDataSQL04 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用(撤去)
'        MODULE_ID        :createExcelDataSQL05
'        NOTE             :createBoxDataSQL05を参考に作成
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createExcelDataSQL05() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      YARD_MAST.YARD_CODE as YARD_CODE "
    strSQL = strSQL & "     ,YARD_MAST.YARD_NAME as YARD_NAME "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BOXNO as BOXNO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
    strSQL = strSQL & "     ,CBRA_RECD.CBRAR_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_RECD.CBYSR_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_RECD.CBYSR_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_RECD.CBYSR_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
    strSQL = strSQL & "      ,CBYS_RECD LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_RECD.CBYSR_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "          LEFT OUTER JOIN CBRA_RECD ON CBYS_RECD.CBYSR_YCODE = CBRA_RECD.CBRAR_YCODE AND CBYS_RECD.CBYSR_BOXNO = CBRA_RECD.CBRAR_BOXNO AND CBYS_RECD.CBYSR_EDANO = CBRA_RECD.CBRAR_EDANO AND CBYS_RECD.CBYSR_SETDATE = CBRA_RECD.CBRAR_SETDATE "
    strSQL = strSQL & "      ,YARD_MAST "
    strSQL = strSQL & " where CBYS_RECD.CBYSR_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "   and CBYS_RECD.CBYSR_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "   and CBYS_RECD.CBYSR_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_RECD.CBYSR_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
    strSQL = strSQL & "   and REPLACE (CBYS_RECD.CBYSR_YCODE,' ','') = YARD_MAST.YARD_CODE "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
    strSQL = strSQL & " ; "
    
    createExcelDataSQL05 = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :EXCEL出力用(設置中)
'        MODULE_ID        :createExcelDataSQL09
'        NOTE             :createBoxDataSQL09を参考に作成
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createExcelDataSQL09() As String
    
    Dim strSQL        As String
    
    strSQL = ""
    strSQL = strSQL & " select  "
    strSQL = strSQL & "      YARD_MAST.YARD_CODE as YARD_CODE "
    strSQL = strSQL & "     ,YARD_MAST.YARD_NAME as YARD_NAME "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BOXNO as BOXNO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "
    strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "
    strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "
    strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "
    strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "
    strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "
    strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "
    strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "
    strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "
    strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "
    strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "
    strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "
    strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "
    strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "
    strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "
    strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "
    strSQL = strSQL & " from CBOX_MAST LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 ON NAME_MAST1.NAME_CODE = CBOX_MAST.CBOXM_TYPE AND NAME_MAST1.NAME_ID = '156' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 ON NAME_MAST2.NAME_CODE = CBOX_MAST.CBOXM_INTYPE AND NAME_MAST2.NAME_ID = '274' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 ON NAME_MAST3.NAME_CODE = CBOX_MAST.CBOXM_SIZE AND NAME_MAST3.NAME_ID = '271' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 ON NAME_MAST4.NAME_CODE = CBOX_MAST.CBOXM_DOOR AND NAME_MAST4.NAME_ID = '272' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST5 ON NAME_MAST5.NAME_CODE = CBOX_MAST.CBOXM_DOORPSTN AND NAME_MAST5.NAME_ID = '275' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST6 ON NAME_MAST6.NAME_CODE = CBOX_MAST.CBOXM_DOORTYPE AND NAME_MAST6.NAME_ID = '157' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST8 ON NAME_MAST8.NAME_CODE = CBOX_MAST.CBOXM_KNKBN AND NAME_MAST8.NAME_ID = '266' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST9 ON NAME_MAST9.NAME_CODE = CBOX_MAST.CBOXM_CTBNR AND NAME_MAST9.NAME_ID = '267' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST10 ON NAME_MAST10.NAME_CODE = CBOX_MAST.CBOXM_CTYTO AND NAME_MAST10.NAME_ID = '268' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST11 ON NAME_MAST11.NAME_CODE = CBOX_MAST.CBOXM_KNKAK AND NAME_MAST11.NAME_ID = '269' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST12 ON NAME_MAST12.NAME_CODE = CBOX_MAST.CBOXM_SHSHU AND NAME_MAST12.NAME_ID = '270' "
    strSQL = strSQL & "          LEFT OUTER JOIN NAME_MAST AS NAME_MAST13 ON NAME_MAST13.NAME_CODE = CBOX_MAST.CBOXM_KRKBN AND NAME_MAST13.NAME_ID = '195' "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBRA_TRAN.CBRAT_BOXNO,CBRA_TRAN.CBRAT_EDANO,CBRA_TRAN.CBRAT_YCODE,CBRA_TRAN.CBRAT_CNO from CBRA_TRAN,CNTA_MAST where CBRAT_CNO = CNTA_NO and CBRAT_YCODE = CNTA_CODE) as w01 on w01.CBRAT_BOXNO = CBOX_MAST.CBOXM_BOXNO and w01.CBRAT_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "          LEFT OUTER JOIN (select CBYST_BOXNO,CBYST_EDANO from CBYS_TRAN where CBYST_STATUS = 0 and CBYS_TRAN.CBYST_SYUBETSU = 1) as w02 on w02.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO and w02.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "      ,CBYS_TRAN LEFT OUTER JOIN NAME_MAST AS NAME_MAST7 ON NAME_MAST7.NAME_CODE = CBYS_TRAN.CBYST_SETFLOOR AND NAME_MAST7.NAME_ID = '090' "
    strSQL = strSQL & "      ,YARD_MAST "
    strSQL = strSQL & " where CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
    strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & Me.cmb_ヤード & "',' ','') "
    strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = YARD_MAST.YARD_CODE "
    strSQL = strSQL & " order by BOXNO,EDANO,SETDATE,CNO "
    strSQL = strSQL & " ; "
    
    createExcelDataSQL09 = strSQL
    
End Function
'↑INSERT 2013/11/22 MIYAMOTO
'****************************  ended or program ********************************

Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_標題.BackColor = RGB(255, 0, 204)
    End Select

End Sub
