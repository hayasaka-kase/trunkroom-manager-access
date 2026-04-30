VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS720"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード別割引一覧
'   プログラムＩＤ　：　FKS720
'   作　成　日　　　：  2008/07/17
'   作　成　者　　　：  N.IMAI
'   Ver             ：  0.0
'**********************************************
'   修　正　日　　　：  2008/10/25
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  優先順位100の場合、優先が重複しているかチェック
'   　　　　　　　　　　をしないよう修正した。
'   Ver             ：  0.1
'**********************************************
'   修　正　日　　　：  2008/11/18
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  ヤードをフリー入力時にマスタデータを読まないバグを修正
'   Ver             ：  0.2
'**********************************************
'   修　正　日　　　：  2008/12/22
'   修　正　者　　　：  K.KINEBUCHI
'   内　　　容　　　：  終了した割引を削除する際の不具合を修正する。
'   Ver             ：  0.1（これより改修分のVerは、0.1とする。以降+0.1ずづアップする）
'**********************************************
'   修　正　日　　　：  2009/06/30
'   修　正　者　　　：  hirano
'   内　　　容　　　：  割引マスタ　KASE_DB移行に伴う改修
'   Ver             ：  0.2
'**********************************************
'   修　正　日　　　：  2013/06/03
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  終了メッセージを表示させないように修正
'   Ver             ：  0.3
'**********************************************
'   修　正　日　　　：  2019/02/25
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  保存時にサービス内容を保存するように対応
'   Ver             ：  0.4
'**********************************************
'   修　正　日　　　：  2022/05/06
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  割引追加時にヤードのサービス１を更新するように変更
'   Ver             ：  0.5
'**********************************************
'   修　正　日　　　：  2022/11/01
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  フォームのプロパティでスクロールバーを表示するよう変更
'   Ver             ：  0.6
'**********************************************
'   修　正　日　　　：  2024/02/15
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  原因不明なのでワークテーブルの作成・削除をやめる
'   Ver             ：  0.7
'**********************************************

Option Compare Database
Option Explicit

Private Const P_KAKUNIN = "確認"
Private Const P_ERROR = "エラー"
Private Const P_TITLE_UPDATE = "（修正中）"

Private Const PROG_ID = "FKS720"
Private Const P_NEBIKI_LIST_WORK_TABLE = "WK_FKS720"

Private Const BUTTON_MSG_NOW = "実施中割引"
Private Const BUTTON_MSG_END = "終了した割引"

Private Const P_MODE_NORMAL = 1
Private Const P_MODE_END = 0

Private pobjDB                      As Object
Private pstrBumonCode               As String       ' 部門コード
Private gstrUserName                As String
Private pblnRequery                 As Boolean
Private pstrFormTitle               As String
'Private pintMode                    As Integer     'DELETE 2008/12/22 K.KINEBUCHI
Public pintMode                     As Integer      'INSERT 2008/12/22 K.KINEBUCHI
Private pintKensu                   As Integer      '有効割引件数

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

On Error GoTo Exception
    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb
    
    gstrUserName = MSZZ000.LsGetUserName
    pstrFormTitle = Me.Caption
    pintKensu = 0
    pintMode = P_MODE_NORMAL
    
    Call dropWorkTable
    Call createWorkTable
    Call subRequery
    pblnRequery = False
    
Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("Form_Load", True)
    End If
    
End Sub


'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    
    On Error Resume Next
    
    doCmd.Maximize
    doCmd.Requery
    If pblnRequery Then
        Call subRequery
        pblnRequery = False
    End If
   Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_標題.BackColor = RGB(255, 0, 204)
    End Select
    Err.Reset

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤード_AfterUpdate
'        MODULE_ID        :cmb_ヤード_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤード_AfterUpdate()

On Error GoTo Exception

    'I:20081118 N.Imai Start
    If Me.cmb_ヤード <> Format(Nz(Me.cmb_ヤード), "000000") Then
        Me.cmb_ヤード = Format(Nz(Me.cmb_ヤード), "000000")
        Me.cmb_ヤード.Requery
    End If
    'I:20081118 N.Imai End
    If Me.cmb_ヤード.ListIndex = -1 Then
        Me.Caption = Me.cmb_ヤード.Text
        Me.lbl_ヤード名.Caption = ""
        Me.lbl_ヤード所在地.Caption = ""
        Me.lbl_営業開始日.Caption = ""
        Me.lbl_営業終了日.Caption = ""
        Me.lbl_ヤード解約予定日.Caption = ""
        txt_サービス1 = ""
        txt_サービス2 = ""
        txt_サービス3 = ""
        txt_サービス期間 = ""
        lbl_ネット予約対象.Visible = False
    Else
        Me.lbl_ヤード名.Caption = Me.cmb_ヤード.Column(1, Me.cmb_ヤード.ListIndex)
        Me.lbl_ヤード所在地.Caption = Me.cmb_ヤード.Column(2, Me.cmb_ヤード.ListIndex)
        Me.lbl_営業開始日.Caption = Me.cmb_ヤード.Column(3, Me.cmb_ヤード.ListIndex)
        Me.lbl_営業終了日.Caption = Me.cmb_ヤード.Column(4, Me.cmb_ヤード.ListIndex)
        Me.lbl_ヤード解約予定日.Caption = Me.cmb_ヤード.Column(5, Me.cmb_ヤード.ListIndex)
        txt_サービス1 = Me.cmb_ヤード.Column(6, Me.cmb_ヤード.ListIndex)
        txt_サービス2 = Me.cmb_ヤード.Column(7, Me.cmb_ヤード.ListIndex)
        txt_サービス3 = Me.cmb_ヤード.Column(8, Me.cmb_ヤード.ListIndex)
        txt_サービス期間 = Me.cmb_ヤード.Column(9, Me.cmb_ヤード.ListIndex)
        lbl_ネット予約対象.Visible = (Me.cmb_ヤード.Column(10, Me.cmb_ヤード.ListIndex) = 0)
    End If
    
    'Me.cmb_ヤード = Format(Nz(Me.cmb_ヤード), "000000")                        'D:20081118 N.Imai
    Me.cmd_割引追加.Enabled = (Me.cmb_ヤード <> "")
    Me.cmd_表示ステータス.Enabled = (Me.cmb_ヤード <> "")
    
    '検索
    Call subCreateListData_FKS720

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmb_ヤード_AfterUpdate", True)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サービス1_Change
'        MODULE_ID        :txt_サービス1_Change
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サービス1_Change()

    Call subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サービス2_Change
'        MODULE_ID        :txt_サービス2_Change
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サービス2_Change()
    
    Call subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サービス3_Change
'        MODULE_ID        :txt_サービス3_Change
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サービス3_Change()
    
    Call subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サービス期間_Change
'        MODULE_ID        :txt_サービス期間_Change
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サービス期間_Change()
    
    Call subCaptionChange
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :キャプション変更
'        MODULE_ID        :subCaptionChange
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subCaptionChange()
    
    If InStr(1, Me.Caption, P_TITLE_UPDATE) <= 0 Then
        Me.Caption = Me.Caption & P_TITLE_UPDATE
        Me.cmd_保存.Enabled = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_表示ステータス_Click
'        MODULE_ID        :cmd_表示ステータス_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_表示ステータス_Click()

On Error GoTo Exception
    
    If fncUpdateCheck(0) Then
        If pintMode = P_MODE_NORMAL Then
            Me.cmd_表示ステータス.Caption = BUTTON_MSG_NOW
            Me.Refresh
            pintMode = P_MODE_END
        Else
            Me.cmd_表示ステータス.Caption = BUTTON_MSG_END
            Me.Refresh
            pintMode = P_MODE_NORMAL
        End If
        Call subCreateListData_FKS720
    End If

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmd_表示ステータス_Click", True)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧表時用ワークテーブル更新
'        MODULE_ID        :subCreateListData_FKS720
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub subCreateListData_FKS720()                                          'DELETE 2022/05/06 N.IMAI
Public Sub subCreateListData_FKS720(Optional intYardMasterUpdate = 0)           'INSERT 2022/05/06 N.IMAI
    
    Dim objAdoDbConnection  As Object
    Dim rsDBData            As Object       'ADODB.Recordset
    Dim rsListData          As Recordset
    Dim strSQL              As String
    Dim strKaseSqlSvr       As String       'INS 2009/06/30 hirano
    
    On Error GoTo Exception
    '2009/06/30 INS <S> 割引マスタ KASE_DBへ移動
    'SQL-Server名
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    '2009/06/30 INS <E>
    
    'ワークデータを一度削除
    strSQL = ""
    strSQL = strSQL & "DELETE FROM " & P_NEBIKI_LIST_WORK_TABLE
    Call pobjDB.Execute(strSQL)
    
    strSQL = ""
    strSQL = strSQL & "SELECT RLDN_TRAN.RLDNT_BUMOC"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_YCODE"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_NO"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_ENABLE"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_FROM"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_TO"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_ORDER"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_TEXT"
    strSQL = strSQL & "      ,RLDN_TRAN.RLDNT_NOTE"
    strSQL = strSQL & "      ,DCNT_MAST.DCNTM_NO"
    strSQL = strSQL & "      ,DCNT_MAST.DCNTM_TYPE"
    strSQL = strSQL & "      ,DCNT_MAST.DCNTM_PERIOD"
    strSQL = strSQL & "      ,DCNT_MAST.DCNTM_SIZE_FROM"
    strSQL = strSQL & "      ,DCNT_MAST.DCNTM_SIZE_TO"
    strSQL = strSQL & "      ,DCNT_MAST.DCNTM_MTEXT"
    strSQL = strSQL & "      ,NAME_MAST1.NAME_RYAK AS NAME_GENKBN"
    strSQL = strSQL & "      ,NAME_MAST2.NAME_RYAK AS NAME_FLOOR"
    strSQL = strSQL & "      ,NAME_MAST3.NAME_RYAK AS NAME_USAGE"
    strSQL = strSQL & "      ,NAME_MAST4.NAME_RYAK AS NAME_USE_PERIOD "
    '2009/06/30 MOD <S>
    'strSQL = strSQL & "FROM RLDN_TRAN, DCNT_MAST "
    strSQL = strSQL & "FROM RLDN_TRAN, " & strKaseSqlSvr & ".dbo.DCNT_MAST "
    '2009/06/30 MOD <E>
    strSQL = strSQL & "LEFT OUTER JOIN NAME_MAST AS NAME_MAST1 "
    strSQL = strSQL & "  ON NAME_MAST1.NAME_CODE = DCNT_MAST.DCNTM_GENKBN "
    strSQL = strSQL & " AND NAME_MAST1.NAME_ID = '083' "
    strSQL = strSQL & "LEFT OUTER JOIN NAME_MAST AS NAME_MAST2 "
    strSQL = strSQL & "  ON NAME_MAST2.NAME_CODE = DCNT_MAST.DCNTM_FLOOR "
    strSQL = strSQL & " AND NAME_MAST2.NAME_ID = '090' "
    strSQL = strSQL & "LEFT OUTER JOIN NAME_MAST AS NAME_MAST3 "
    strSQL = strSQL & "  ON NAME_MAST3.NAME_CODE = DCNT_MAST.DCNTM_USAGE "
    strSQL = strSQL & " AND NAME_MAST3.NAME_ID = '086' "
    strSQL = strSQL & "LEFT OUTER JOIN NAME_MAST AS NAME_MAST4 "
    strSQL = strSQL & "  ON NAME_MAST4.NAME_CODE = DCNT_MAST.DCNTM_USE_PERIOD "
    strSQL = strSQL & " AND NAME_MAST4.NAME_ID = '249' "
    strSQL = strSQL & " WHERE RLDN_TRAN.RLDNT_BUMOC  =  '" & pstrBumonCode & "'"
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_YCODE  =  '" & Me.cmb_ヤード & "'"
    '2009/06/30 INS <S> 割引マスタ kase_DB
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_BUMOC  =  DCNT_MAST.DCNTM_BUMOC "
    '2009/06/30 INS <E>
'    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE =  '" & pintMode & "'"
    
    If pintMode = P_MODE_NORMAL Then
        strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE =  '" & pintMode & "'"
        strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_TO >= '" & Format$(Now, "yyyymmdd") & "'"
        strSQL = strSQL & "   AND DCNT_MAST.DCNTM_NO = RLDN_TRAN.RLDNT_NO "
        strSQL = strSQL & " ORDER BY RLDNT_ORDER"
    Else
        strSQL = strSQL & "   AND (RLDN_TRAN.RLDNT_ENABLE =  '" & pintMode & "' OR RLDN_TRAN.RLDNT_TO <  '" & Format$(Now, "yyyymmdd") & "')"
        strSQL = strSQL & "   AND DCNT_MAST.DCNTM_NO = RLDN_TRAN.RLDNT_NO "
        strSQL = strSQL & " ORDER BY RLDNT_FROM"
    End If
    
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    objAdoDbConnection.CursorLocation = 3                                       'INSERT 2022/05/06 N.IMAI
    Set rsDBData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)

    pintKensu = 0
    While Not rsDBData.EOF
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM " & P_NEBIKI_LIST_WORK_TABLE
        Set rsListData = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
        While Not rsDBData.EOF
            rsListData.AddNew
            rsListData.Fields("DATA_NO") = pintKensu
            rsListData.Fields("RLDNT_BUMOC") = rsDBData.Fields("RLDNT_BUMOC")
            rsListData.Fields("RLDNT_YCODE") = rsDBData.Fields("RLDNT_YCODE")
            rsListData.Fields("RLDNT_NO") = rsDBData.Fields("RLDNT_NO")
            rsListData.Fields("RLDNT_ENABLE") = rsDBData.Fields("RLDNT_ENABLE")
            rsListData.Fields("RLDNT_FROM") = rsDBData.Fields("RLDNT_FROM")
            rsListData.Fields("RLDNT_TO") = rsDBData.Fields("RLDNT_TO")
            '期間
            If Nz(rsDBData.Fields("RLDNT_FROM")) = "00000000" And Nz(rsDBData.Fields("RLDNT_TO")) = "99999999" Then
                rsListData.Fields("RLDNT_FROMTO") = "無期限"
            ElseIf Nz(rsDBData.Fields("RLDNT_FROM")) <> "00000000" And Nz(rsDBData.Fields("RLDNT_TO")) <> "99999999" Then
                rsListData.Fields("RLDNT_FROMTO") = Mid$(rsDBData.Fields("RLDNT_FROM"), 3, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_FROM"), 5, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_FROM"), 7, 2) _
                                                  & "～" _
                                                  & Mid$(rsDBData.Fields("RLDNT_TO"), 3, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_TO"), 5, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_TO"), 7, 2)
            ElseIf Nz(rsDBData.Fields("RLDNT_FROM")) <> "00000000" Then
                rsListData.Fields("RLDNT_FROMTO") = Mid$(rsDBData.Fields("RLDNT_FROM"), 3, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_FROM"), 5, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_FROM"), 7, 2) _
                                                  & "～"
            ElseIf Nz(rsDBData.Fields("RLDNT_TO")) <> "99999999" Then
                rsListData.Fields("RLDNT_FROMTO") = "～" _
                                                  & Mid$(rsDBData.Fields("RLDNT_TO"), 3, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_TO"), 5, 2) _
                                                  & "/" _
                                                  & Mid$(rsDBData.Fields("RLDNT_TO"), 7, 2)
            Else
                rsListData.Fields("RLDNT_FROMTO") = "-"
            End If
            rsListData.Fields("RLDNT_ORDER") = rsDBData.Fields("RLDNT_ORDER")
            rsListData.Fields("RLDNT_TEXT") = rsDBData.Fields("RLDNT_TEXT")
            rsListData.Fields("RLDNT_NOTE") = rsDBData.Fields("RLDNT_NOTE")
            '組合
            If Nz(rsDBData.Fields("DCNTM_PERIOD")) = 0 Then
                '解約まで適用
                If Nz(rsDBData.Fields("DCNTM_TYPE")) = 1 Then
                    rsListData.Fields("KUMI_1") = "×"
                Else
                    rsListData.Fields("KUMI_1") = "○"
                End If
                rsListData.Fields("KUMI_2") = "－"
            Else
                '指定月まで適用
                If Nz(rsDBData.Fields("DCNTM_TYPE")) = 1 Then
                    rsListData.Fields("KUMI_2") = "×"
                Else
                    rsListData.Fields("KUMI_2") = "○"
                End If
                rsListData.Fields("KUMI_1") = "－"
            End If
            'サイズ
            If Nz(rsDBData.Fields("DCNTM_SIZE_FROM")) <> 0 And Nz(rsDBData.Fields("DCNTM_SIZE_TO")) <> 0 Then
                If Nz(rsDBData.Fields("DCNTM_SIZE_FROM")) = Nz(rsDBData.Fields("DCNTM_SIZE_TO")) Then
                    rsListData.Fields("DCNTM_SIZE") = rsDBData.Fields("DCNTM_SIZE_FROM") & "帖"
                Else
                    rsListData.Fields("DCNTM_SIZE") = rsDBData.Fields("DCNTM_SIZE_FROM") & "～" & rsDBData.Fields("DCNTM_SIZE_TO") & "帖"
                End If
            ElseIf Nz(rsDBData.Fields("DCNTM_SIZE_FROM")) <> 0 Then
                    rsListData.Fields("DCNTM_SIZE") = rsDBData.Fields("DCNTM_SIZE_FROM") & "～"
            ElseIf Nz(rsDBData.Fields("DCNTM_SIZE_TO")) <> 0 Then
                    rsListData.Fields("DCNTM_SIZE") = "～" & rsDBData.Fields("DCNTM_SIZE_TO")
            Else
                    rsListData.Fields("DCNTM_SIZE") = ""
            End If
            rsListData.Fields("NAME_GENKBN") = rsDBData.Fields("NAME_GENKBN")
            rsListData.Fields("NAME_FLOOR") = Nz(rsDBData.Fields("NAME_FLOOR"), "-")
            rsListData.Fields("NAME_USAGE") = Nz(rsDBData.Fields("NAME_USAGE"), "-")
            rsListData.Fields("NAME_USE_PERIOD") = Nz(rsDBData.Fields("NAME_USE_PERIOD"), "-")
            rsListData.Fields("DCNTM_MTEXT") = rsDBData.Fields("DCNTM_MTEXT")
            '契約数取得
            rsListData.Fields("KEIYAKU_SU") = fncGetKeiyaku_SU(rsDBData.Fields("DCNTM_NO"), rsDBData.Fields("RLDNT_ENABLE"), objAdoDbConnection)
            rsListData.UPDATE
            'INSAERT 2022/05/06 N.IMAI Start
            If rsDBData.RecordCount = 1 And intYardMasterUpdate = 1 And pintKensu = 0 Then
                Me.txt_サービス1 = rsDBData.Fields("RLDNT_TEXT")
            End If
            'INSAERT 2022/05/06 N.IMAI End
            rsDBData.MoveNext
            If pintMode = P_MODE_NORMAL Then
                pintKensu = pintKensu + 1
            End If
        Wend
        rsListData.Close
        Set rsListData = Nothing
    Wend
    
    Me.cmd_割引追加.Enabled = (pintMode = P_MODE_NORMAL)
    Me.cmb_ヤード.SetFocus
    Me.cmd_保存.Enabled = False

    '一覧再表示
    Call subRequery

    Me.Caption = pstrFormTitle

    rsDBData.Close
    Set rsDBData = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

Exception:
    If Not rsDBData Is Nothing Then: rsDBData.Close: Set rsDBData = Nothing
    If Not rsListData Is Nothing Then: rsListData.Close: Set rsListData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "subCreateListData_FKS720" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル作成
'        MODULE_ID        :createWorkTable
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub createWorkTable()

'DELETE 2024/02/15 N.IMAI Start
'
'    Dim strSQL As String
'
'On Error GoTo Exception
'
'    '一覧画面用ワークテーブル
'    strSQL = ""
'    strSQL = strSQL & "CREATE TABLE " & P_NEBIKI_LIST_WORK_TABLE & " ("
'    strSQL = strSQL & " DATA_NO         INTEGER"
'    strSQL = strSQL & ",RLDNT_BUMOC     VARCHAR(1)"
'    strSQL = strSQL & ",RLDNT_YCODE     VARCHAR(6)"
'    strSQL = strSQL & ",RLDNT_NO        VARCHAR(6)"
'    strSQL = strSQL & ",RLDNT_ENABLE    VARCHAR(1)"
'    strSQL = strSQL & ",RLDNT_FROM      VARCHAR(8)"
'    strSQL = strSQL & ",RLDNT_TO        VARCHAR(8)"
'    strSQL = strSQL & ",RLDNT_FROMTO    VARCHAR(30)"
'    strSQL = strSQL & ",RLDNT_ORDER     VARCHAR(3)"
'    strSQL = strSQL & ",RLDNT_TEXT      VARCHAR(42)"
'    strSQL = strSQL & ",RLDNT_NOTE      VARCHAR(60)"
'    strSQL = strSQL & ",DCNTM_SIZE      VARCHAR(20)"
'    strSQL = strSQL & ",DCNTM_MTEXT     VARCHAR(42)"
'    strSQL = strSQL & ",NAME_GENKBN     VARCHAR(10)"
'    strSQL = strSQL & ",NAME_FLOOR      VARCHAR(10)"
'    strSQL = strSQL & ",NAME_USAGE      VARCHAR(10)"
'    strSQL = strSQL & ",NAME_USE_PERIOD VARCHAR(10)"
'    strSQL = strSQL & ",KUMI_1          VARCHAR(2)"
'    strSQL = strSQL & ",KUMI_2          VARCHAR(2)"
'    strSQL = strSQL & ",KEIYAKU_SU      VARCHAR(5)"
'    strSQL = strSQL & ")"
'    pobjDB.Execute strSQL, dbFailOnError
'
'Exception:
'    If Err <> 0 Then
'        Call Err.Raise(Err.Number, "createWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
'DELETE 2024/02/15 N.IMAI End

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル削除
'        MODULE_ID        :dropWorkTable
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub dropWorkTable()
    
    On Error Resume Next

    'DELETE 2024/02/15 N.IMAI Start
    'Me.自動割引一覧.Form.RecordSource = ""
    'Me.自動割引一覧.Form.Requery
    'pobjDB.Execute "DROP TABLE " & P_NEBIKI_LIST_WORK_TABLE & " ;", dbFailOnError
    'DELETE 2024/02/15 N.IMAI End
    
    pobjDB.Execute "DELETE FROM " & P_NEBIKI_LIST_WORK_TABLE & " ;", dbFailOnError  'INSERT 2024/02/15 N.IMAI
    pintKensu = 0

    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :割引追加ボタン押下時処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_割引追加_Click()

    Dim strParameter    As String

On Error GoTo Exception

    If fncUpdateCheck(2) = True Then
        pblnRequery = True
        '(0)モード - (1)ヤードコード - (2)ヤード名 - (3)割引数
        strParameter = "割引追加," & cmb_ヤード & "," & lbl_ヤード名.Caption & "," & pintKensu
        doCmd.OpenForm "FKS700", , , , , acDialog, strParameter
        'Call subCreateListData_FKS720                                          'DELETE 2022/05/06 N.IMAI
        Call subCreateListData_FKS720(1)                                        'INSERT 2022/05/06 N.IMAI
    End If

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmd_割引追加_Click", True)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :更新確認処理
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncUpdateCheck(intMessage As Integer) As Boolean

    Dim strMessage  As String

On Error GoTo Exception

    If InStr(1, Me.Caption, P_TITLE_UPDATE) > 0 Then
        Select Case intMessage
            '表示切替
            Case 0
                strMessage = "更新しますか？" + vbCrLf + "※更新しなければ一覧の切替が出来ません。"
            '修正画面遷移
            Case 1
                strMessage = "更新しますか？" + vbCrLf + "※更新しなければデータの修正が出来ません。"
            '割引追加
            Case 2
                strMessage = "更新しますか？" + vbCrLf + "※更新しなければ割引追加が出来ません。"
        End Select
        fncUpdateCheck = False
        If MsgBox(strMessage, vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
            '更新処理
            fncUpdateCheck = fncHozon(False)
        End If
    Else
        fncUpdateCheck = True
    End If

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("subUpdateCheck", True)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :保存ボタン押下時処理
'        MODULE_ID        :cmd_保存_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()
  
    Dim strSQL      As String
    Dim rsData      As Object
    Dim rsDataCheck As Object
    Dim blnError    As Boolean

On Error GoTo Exception

    If fncHozon(True) Then
        '更新処理
        If fncUpdate Then
            Me.Caption = pstrFormTitle
            cmb_ヤード.Requery
        End If
    End If

Exception:

    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    If Not rsDataCheck Is Nothing Then
        rsDataCheck.Close
        Set rsDataCheck = Nothing
    End If

    If Err <> 0 Then
        Call MSZZ024_M00("cmd_保存_Click", True)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :更新前チェック処理
'        MODULE_ID        :fncHozon
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncHozon(bMessage As Boolean) As Boolean
  
    Dim strSQL      As String
    Dim rsData      As Object
    Dim rsDataCheck As Object
    Dim blnError    As Boolean
    
    fncHozon = False
    
    'エラーチェック
    blnError = False
    strSQL = "SELECT * FROM " & P_NEBIKI_LIST_WORK_TABLE
    Set rsData = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    While (Not rsData.EOF) And (blnError = False)
        '優先が重複しているかチェック
        If rsData.Fields("RLDNT_ORDER") <> 100 Then                         'INSERT 2008/10/25 N.IMAI
            strSQL = ""
            strSQL = strSQL & "SELECT * FROM " & P_NEBIKI_LIST_WORK_TABLE
            strSQL = strSQL & " WHERE DATA_NO     <> " & rsData.Fields("DATA_NO")
            strSQL = strSQL & "   AND RLDNT_ORDER = '" & rsData.Fields("RLDNT_ORDER") & "'"
            Set rsDataCheck = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
            If Not rsDataCheck.EOF Then
                Call MsgBox("同じ優先順位が指定されています。" + vbCrLf + "優先順位は重ならないようにして下さい。", vbExclamation + vbOKOnly, P_ERROR)
                blnError = True
            End If
            rsDataCheck.Close
            Set rsDataCheck = Nothing
        End If                                                              'INSERT 2008/10/25 N.IMAI
        rsData.MoveNext
    Wend
    rsData.Close
    Set rsData = Nothing

    If blnError = False Then
        If bMessage Then
            '更新確認
            If MsgBox("画面内容のサービス文言、割引優先順で保存します。" + vbCrLf + "よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
                '更新処理
                If fncUpdate Then
                    Me.Caption = pstrFormTitle
                    cmb_ヤード.Requery
                    fncHozon = True
                End If
            End If
        Else
            '更新処理
            If fncUpdate Then
                Me.Caption = pstrFormTitle
                cmb_ヤード.Requery
                fncHozon = True
            End If
        End If
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :更新処理
'        MODULE_ID        :fncUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdate() As Boolean
  
    Dim strSQL              As String
    Dim rsData              As Object
    Dim rsDBData            As Object
    Dim objAdoDbConnection  As Object
    Dim strDate             As String
    Dim strTime             As String
    Dim dtDate              As Date

On Error GoTo Exception

    '初期値設定
    fncUpdate = False
    dtDate = Now
    strDate = Format(dtDate, "yyyymmdd")
    strTime = Format(dtDate, "hhmmss")

    'トランザクション開始
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    objAdoDbConnection.BeginTrans
    
    'ヤードマスタ更新
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM YARD_MAST"
    strSQL = strSQL & " WHERE YARD_MAST.YARD_CODE =  " & CLng(Nz(Me.cmb_ヤード))
    Set rsDBData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If Not rsDBData.EOF Then
        rsDBData.Fields("YARD_SEV1N") = txt_サービス1
        rsDBData.Fields("YARD_SEV2N") = txt_サービス2
        rsDBData.Fields("YARD_SEV3N") = txt_サービス3
        rsDBData.Fields("YARD_ENDEN") = txt_サービス期間
        rsDBData.Fields("YARD_UPDATE") = dtDate
        rsDBData.UPDATE
    End If
    rsDBData.Close
    Set rsDBData = Nothing
    
    'レンタル物件割引更新
    strSQL = "SELECT * FROM " & P_NEBIKI_LIST_WORK_TABLE
    Set rsData = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    While (Not rsData.EOF)
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM RLDN_TRAN"
        strSQL = strSQL & " WHERE RLDN_TRAN.RLDNT_BUMOC  =  '" & pstrBumonCode & "'"
        strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_YCODE  =  '" & Me.cmb_ヤード & "'"
        strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_NO     =  '" & rsData.Fields("RLDNT_NO") & "'"
        strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE =  '" & pintMode & "'"
        strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_FROM   =  '" & rsData.Fields("RLDNT_FROM") & "'"
        strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_TO     =  '" & rsData.Fields("RLDNT_TO") & "'"
        Set rsDBData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsDBData.EOF Then
            '優先順位を更新する
            '▼2008/12/05 hirano Add エラーハンドリング対応
            On Error GoTo Exception2
            '▲2008/12/05 hirano Add
            '2019/02/25 M.HONDA INS
            rsDBData.Fields("RLDNT_SEV1N") = txt_サービス1
            rsDBData.Fields("RLDNT_SEV2N") = txt_サービス2
            rsDBData.Fields("RLDNT_SEV3N") = txt_サービス3
            rsDBData.Fields("RLDNT_ENDEN") = txt_サービス期間
            '2019/02/25 M.HONDA INS
            rsDBData.Fields("RLDNT_ORDER") = rsData.Fields("RLDNT_ORDER")
            rsDBData.Fields("RLDNT_UPDAD") = strDate
            rsDBData.Fields("RLDNT_UPDAJ") = strTime
            rsDBData.Fields("RLDNT_UPDPB") = PROG_ID
            rsDBData.Fields("RLDNT_UPDUB") = gstrUserName
            rsDBData.UPDATE
        End If
        rsDBData.Close
        Set rsDBData = Nothing
        rsData.MoveNext
    Wend
    rsData.Close
    Set rsData = Nothing
    
    'コミット
    objAdoDbConnection.CommitTrans
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    fncUpdate = True

    '再表示
    Call subCreateListData_FKS720

Exception:

    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    If Not rsDBData Is Nothing Then
        rsDBData.Close
        Set rsDBData = Nothing
    End If

    If Err <> 0 Then
        If Not objAdoDbConnection Is Nothing Then
            'ロールバック
            objAdoDbConnection.RollbackTrans
            objAdoDbConnection.Close
            Set objAdoDbConnection = Nothing
        End If
        Call MSZZ024_M00("cmd_保存_Click", True)
    End If
'▼2008/12/05 hirano Add エラーハンドリング対応
    Exit Function
Exception2:
    If Not rsDBData Is Nothing Then
        rsDBData.Cancel
        Set rsDBData = Nothing
    End If
    GoTo Exception
'▲2008/12/05 hirano Add
End Function

'==============================================================================*
'
'        MODULE_NAME      :終了ボタン押下時処理
'        MODULE_ID        :cmd_終了_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

On Error GoTo Exception
    
    '' 2013/06/03 M.HONDA START
'    If InStr(1, Me.Caption, P_TITLE_UPDATE) > 0 Then
'        If MsgBox("修正中ですが保存せずに終了しますか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
'            '更新処理
'            doCmd.Close acForm, Me.FormName
'        End If
'    Else
'        If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
    '' 2013/06/03 M.HONDA START
'        End If
'    End If
    '' 2013/06/03 M.HONDA END
    doCmd.Close acForm, Me.FormName
    
Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmd_終了_Click", True)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Close
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

On Error GoTo Exception
    
    Call dropWorkTable
    
    'コネクションクローズ
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("Form_Close", True)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subRequery
'        機能             :一覧表示
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subRequery()
    
    Me.自動割引一覧.Form.RecordSource = "SELECT * FROM " & P_NEBIKI_LIST_WORK_TABLE & " ;"
    Me.自動割引一覧.Form.Requery

End Sub
'==============================================================================*
'
'        MODULE_NAME      :fncGetKeiyaku_SU
'        機能             :契約数取得
'        IN               :割引番号
'                          有効可否
'                          コネクション
'        OUT              :契約数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKeiyaku_SU(strNEBIKI_NO As String, strENABLE As String, objConnection As Object) As Integer

    Dim strSQL         As String
    Dim rsKEIYAKU      As Object
    Dim intKEIYAKU_CNT As Integer

    On Error GoTo Exception

    intKEIYAKU_CNT = 0

    '契約数取得
    strSQL = ""
    '2009/06/30 MOD <S> hirano
'    strSQL = strSQL & "SELECT COUNT(DCRAT_DCNT_NO) AS DATA_CNT "
    strSQL = strSQL & "SELECT COUNT(*) AS DATA_CNT FROM ("
    strSQL = strSQL & " SELECT DCRA_TRAN.DCRAT_ACPTNO "
    '2009/06/30 MOD <E> hirano
    strSQL = strSQL & "  FROM RCPT_TRAN, DCRA_TRAN, RLDN_TRAN"
    strSQL = strSQL & " WHERE RCPT_TRAN.RCPT_YCODE       = '" & Me.cmb_ヤード & "'"
    strSQL = strSQL & "   AND RCPT_TRAN.RCPT_CARG_ACPTNO = DCRA_TRAN.DCRAT_ACPTNO "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_YCODE      = RCPT_TRAN.RCPT_YCODE "
'    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE     = '" & strENABLE & "' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_NO         = '" & strNEBIKI_NO & "'"
    strSQL = strSQL & "   AND DCRA_TRAN.DCRAT_DCNT_NO    = RLDN_TRAN.RLDNT_NO"
'    strSQL = strSQL & "   AND DCRA_TRAN.DCRAT_FROM BETWEEN RLDN_TRAN.RLDNT_FROM AND RLDN_TRAN.RLDNT_TO "
    '2009/06/30 ADD <S> hirano
    strSQL = strSQL & " GROUP BY DCRA_TRAN.DCRAT_ACPTNO ) AS A"
    '2009/06/30 ADD <E> hirano
    Set rsKEIYAKU = MSZZ025.ADODB_Recordset(strSQL, objConnection)

    intKEIYAKU_CNT = CStr(Nz(rsKEIYAKU.Fields("DATA_CNT"), 0))
    rsKEIYAKU.Close
    Set rsKEIYAKU = Nothing

    fncGetKeiyaku_SU = intKEIYAKU_CNT

    Exit Function

Exception:
    If Not rsKEIYAKU Is Nothing Then: rsKEIYAKU.Close: Set rsKEIYAKU = Nothing
    Call Err.Raise(Err.Number, "fncGetKeiyaku_SU" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
