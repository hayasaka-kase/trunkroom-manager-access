VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS810"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細）
'   プログラムＩＤ　：　FKS810
'   作　成　日　　　：  2008/06/03
'   作　成　者　　　：  iizuka
'   Ver             ：  0.0
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細）
'   プログラムＩＤ　：　FKS810
'   作　成　日　　　：  2013/11/18
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.1
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細）
'   プログラムＩＤ　：　FKS810
'   作　成　日　　　：  2013/12/06
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.2
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細）
'   プログラムＩＤ　：　FKS810
'   作　成　日　　　：  2014/01/17
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.3
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細）
'   プログラムＩＤ　：　FKS810
'   作　成　日　　　：  2014/01/23
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.4
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細）
'   プログラムＩＤ　：　FKS810
'   作　成　日　　　：  2018/07/19
'   作　成　者　　　：  TRIPLE FIELD PARK
'   Ver             ：  0.5
'**********************************************
Option Compare Database
Option Explicit

Private Const PROG_ID = "FKS810"

Private Const P_ROOM_WORK_TABLE          As String = "WK_FKS800"
Private Const P_COMBO_BOX_WORK_TABLE     As String = "WK_FKS810_CMBD"

Private Const MODE_INSERT = 1
Private Const MODE_UPDATE = 2
Private Const MODE_Y_MOVE = 3

Private pobjDB                      As Object
Private Const adCmdText = 1
Private pstrBumonCode               As String       ' 部門コード
Private pstrProcMode                As String       ' 処理モード

Private pstrType        As String                   ' タイプ
Private pstrSize        As String                   ' サイズ
Private pstrDoor        As String                   ' ドア数
Private pstrDoorType    As String                   ' ドアタイプ
Private pstrDoorIti     As String                   ' ドア位置
Private pstrInDesign    As String                   ' 内装
Private pstrFloorLevel  As String                   ' 設置階層
Private pstrSetDate     As String                   ' 設置日
'↓INSERT 2013/11/18 MIYAMOTO
Private pstrOldEdaNo    As String                   '枝番
Private pstrKanriKbn    As String                   '所有
Private pstrKounyuKbn   As String                   '購入区分
Private pstrCtBunrui    As String                   'コンテナ分類
Private pstrCtYouto     As String                   '設置用途
Private pstrKenKakunin  As String                   '建築確認
Private pstrSySyubetu   As String                   'コンテナ内通路
Private pstrBunkatuSu   As String                   '分割数
'↑INSERT 2013/11/18 MIYAMOTO

Private pstrOldYCode        As String
Private pstrOldBoxNo        As String
Private pstrOldFloorLevel   As String

Private pstrDateValue   As String                   ' 日付の入力チェック用
Private pboolFrmActive  As Boolean
Private pstrParameter() As String                   '初期表示データ

'==============================================================================*
'
'        MODULE_NAME      :cmb_サイズ_AfterUpdate
'        MODULE_ID        :cmb_サイズ_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_サイズ_AfterUpdate()
    
    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    If (Me.cmb_サイズ.ListIndex = -1) Then
        pstrSize = ""
    Else
        pstrSize = Me.cmb_サイズ.Column(1, Me.cmb_サイズ.ListIndex)
    End If
    
    Call setKontComboDataSQL
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_サイズ_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_タイプ_AfterUpdate
'        MODULE_ID        :cmb_タイプ_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_タイプ_AfterUpdate()

    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    If (Me.cmb_タイプ.ListIndex = -1) Then
        pstrType = ""
    Else
        pstrType = Me.cmb_タイプ.Column(1, Me.cmb_タイプ.ListIndex)
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_タイプ_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ドアタイプ_AfterUpdate
'        MODULE_ID        :cmb_ドアタイプ_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ドアタイプ_AfterUpdate()

    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    If (Me.cmb_ドアタイプ.ListIndex = -1) Then
        pstrInDesign = ""
    Else
        pstrInDesign = Me.cmb_ドアタイプ.Column(1, Me.cmb_ドアタイプ.ListIndex)
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_ドアタイプ_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ドア位置_AfterUpdate
'        MODULE_ID        :cmb_ドア位置_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ドア位置_AfterUpdate()

    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    If (Me.cmb_ドア位置.ListIndex = -1) Then
        pstrDoorIti = ""
    Else
        pstrDoorIti = Me.cmb_ドア位置.Column(1, Me.cmb_ドア位置.ListIndex)
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_ドア位置_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ドア数_AfterUpdate
'        MODULE_ID        :cmb_ドア数_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ドア数_AfterUpdate()
    
    Dim isException    As Boolean
    isException = False
    Dim intLoopNo As Integer

On Error GoTo Exception
    
    If (Me.cmb_ドア数.ListIndex = -1) Then
        pstrDoor = ""
    Else
        pstrDoor = Me.cmb_ドア数.Column(1, Me.cmb_ドア数.ListIndex)
    End If
    
    '2019/04/17 M.HONDA INS
    'ドア数 = 1D
    If pstrDoor = "1" Then
        Me.cmb_ドア位置.VALUE = "縦"
        Me.cmb_内装.VALUE = "無"
       
    ElseIf pstrDoor = "99" Then
        cmb_設置用途.VALUE = "―"
        cmb_コンテナ分類.VALUE = "―"
        cmb_サイズ.VALUE = "―"
        cmb_ドア位置.VALUE = "―"
        cmb_タイプ.VALUE = "―"
        cmb_ドアタイプ.VALUE = "―"
        cmb_内装.VALUE = "―"
        cmb_コンテナ内通路.VALUE = "―"
        
    Else
        Me.cmb_ドア位置.VALUE = "横"
        Me.cmb_内装.VALUE = "有"

    End If
'    '2019/04/17 M.HONDA INS
    
    
    Call setKontComboDataSQL
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_ドア数_AfterUpdate", True)
    
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
    isException = False

On Error GoTo Exception

    If Me.cmb_ヤード.ListIndex = -1 Then
        Me.lbl_ヤードコード.Caption = ""
        Me.lbl_ヤード名称.Caption = ""
        Me.lbl_営業開始日.Caption = ""
        Me.lbl_営業終了日.Caption = ""
        Me.lbl_ヤード解約予定日.Caption = ""
    Else
        Me.lbl_ヤード名称.Caption = Me.cmb_ヤード.Column(1, Me.cmb_ヤード.ListIndex)
        Me.lbl_営業開始日.Caption = Me.cmb_ヤード.Column(3, Me.cmb_ヤード.ListIndex)
        Me.lbl_営業終了日.Caption = Me.cmb_ヤード.Column(4, Me.cmb_ヤード.ListIndex)
        Me.lbl_ヤード解約予定日.Caption = Me.cmb_ヤード.Column(5, Me.cmb_ヤード.ListIndex)
        Me.lbl_ヤードコード.Caption = Format(Nz(Me.cmb_ヤード), "000000")
    End If
    If pboolFrmActive Then
        '初期表示は起動しない（Form_Loadイベントから発生してくる為、回避）
        Me.sub_部屋割付.Form.RecordSource = ""
        Call Form_FKS800.dropWorkTable(2)
        Call Form_FKS800.createWorkTable(2)
        Me.sub_部屋割付.Form.RecordSource = "SELECT * FROM " & P_ROOM_WORK_TABLE & " AS W01;"
        Me.sub_部屋割付.Form.Requery
    End If
    
    Me.cmb_ヤード = Format(Nz(Me.cmb_ヤード), "000000")
    
    Call setKontComboDataSQL
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_ヤード_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_設置階層_AfterUpdate
'        MODULE_ID        :cmb_設置階層_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_設置階層_AfterUpdate()
    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    If pboolFrmActive Then
        '初期表示は起動しない（Form_Loadイベントから発生してくる為、回避）
        Me.sub_部屋割付.Form.RecordSource = ""
        Call Form_FKS800.dropWorkTable(2)
        Call Form_FKS800.createWorkTable(2)
        Me.sub_部屋割付.Form.RecordSource = "SELECT * FROM " & P_ROOM_WORK_TABLE & " AS W01;"
        Me.sub_部屋割付.Form.Requery
    End If
    If Me.cmb_設置階層.ListIndex <> -1 Then
        Call setKontComboDataSQL
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_設置階層_AfterUpdate", True)
    
End Sub

'↓INSERT 2014/01/23 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :cmb_設置用途_AfterUpdate
'        MODULE_ID        :cmb_設置用途_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2014/01/23
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_設置用途_AfterUpdate()

    Dim isException    As Boolean
    isException = False

On Error GoTo Exception

    If Nz(Me.txt_ボックス番号.VALUE) = "" And Nz(Me.cmb_設置用途.Text) <> "" Then
        If Form_FKS800.gstrAutoBoxNoKaHi = -1 Then
            Me.txt_ボックス番号.VALUE = fncGetAutoBoxNo
        Else
            Me.txt_ボックス番号.VALUE = ""
        End If
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_内装_AfterUpdate", True)

End Sub
'↑INSERT 2014/01/23 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :cmb_内装_AfterUpdate
'        MODULE_ID        :cmb_内装_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_内装_AfterUpdate()

    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    If (Me.cmb_内装.ListIndex = -1) Then
        pstrInDesign = ""
    Else
        pstrInDesign = Me.cmb_内装.Column(1, Me.cmb_内装.ListIndex)
    End If
    
    Call setKontComboDataSQL
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_内装_AfterUpdate", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :保存ボタン押下時処理
'        MODULE_ID        :cmd_保存_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()
    
    Dim objAdoDbConnection As Object
    Dim isExecuteFlg As Boolean
    Dim isException  As Boolean
    isException = False

On Error GoTo Exception
    
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    '入力チェック
    'エラーチェック
    '    ・割付られた部屋が撤去済みだった場合
    '    ・サイズ（帖）÷ドア数以上の部屋を割付けている場合
    '    ・他ボックスで割付済の部屋を選択している場合
    '    ・指定したドア数以上の部屋割付がされている場合
    If fncInputCheck(objAdoDbConnection) Then

        '更新処理
        If (pstrProcMode = "0" Or pstrProcMode = "3") Then
            If MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, "確認") = vbYes Then
                '新規
                isExecuteFlg = fncHozon(MODE_INSERT, objAdoDbConnection)
                Call setDefaultWindow
                Call setKontComboDataSQL
            End If
        ElseIf (pstrProcMode = "1") Then
            If MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, "確認") = vbYes Then
                '修正
                isExecuteFlg = fncHozon(MODE_UPDATE, objAdoDbConnection)
            End If
        ElseIf (pstrProcMode = "2") Then
            If MsgBox("移動します。よろしいですか？", vbQuestion + vbYesNo, "確認") = vbYes Then
                If (Me.tab_移動先指定.VALUE = 0) Then
                    'ヤード移動
                    isExecuteFlg = fncHozon(MODE_Y_MOVE, objAdoDbConnection)
                ElseIf (Me.tab_移動先指定.VALUE = 1) Then
                    'ボックス移動
                    MsgBox "ボックス間移動"
                End If
            End If
        End If
    End If
    
    If isExecuteFlg Then
        '一覧再描画
        Call Form_FKS800.createBoxData
        Call Form_FKS800.Requery
        If (pstrProcMode = "1" Or pstrProcMode = "2" Or pstrProcMode = "3") Then
            '一覧に戻る
            Call CloseWindow
        End If
    End If
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objAdoDbConnection Is Nothing Then
        objAdoDbConnection.Close
        Set objAdoDbConnection = Nothing
    End If
    If isException Then Call MSZZ024_M00("cmd_保存_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :更新処理
'        MODULE_ID        :fncHozon
'        NOTE             :
'        CREATE_DATE      :2008/06/27
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncHozon(iMode As Integer, objAdoDbConnection As Object) As Boolean

    objAdoDbConnection.BeginTrans
    
On Error GoTo Exception

    fncHozon = False

    'DB更新処理
    Select Case iMode
        Case MODE_INSERT
            fncHozon = INS_RTN(objAdoDbConnection)
        Case MODE_UPDATE
            fncHozon = UPD_RTN(objAdoDbConnection)
        Case MODE_Y_MOVE
            fncHozon = MOVE_Y_RTN(objAdoDbConnection)
    End Select

Exception:

    If fncHozon = True Then
        'コミット
        objAdoDbConnection.CommitTrans
    Else
        objAdoDbConnection.RollbackTrans
    End If
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncHozon" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :戻るボタン押下時処理
'        MODULE_ID        :cmd_戻る_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_戻る_Click()
    Dim isException As Boolean
    isException = False

On Error GoTo Exception

    Call CloseWindow
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_戻る_Click", True)
    
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
'        MODULE_NAME      :フォームロード時処理
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Dim isException    As Boolean
    isException = False

On Error GoTo Exception
    
    pboolFrmActive = False
    
    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb

    cmd_移動履歴.Enabled = False
    
    pstrParameter = Split(Nz(Me.OpenArgs), ",")
    
    Call setDefaultWindow
    pboolFrmActive = True
    Call setKontComboDataSQL
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :部屋割付用ＳＱＬ作成
'        MODULE_ID        :setKontComboDataSQL
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub setKontComboDataSQL()
    
    Dim objAdoDbConnection As Object
    Dim strSQL             As String
    Dim sinSize            As Single
    Dim rsData             As Object       'ADODB.Recordset
    Dim rsDataLocal        As Object
    Dim rsComboBoxData     As Recordset

    On Error GoTo Exception
    
    If ((pstrSize <> "") And _
        (pstrDoor <> "") And _
        (Me.cmb_設置階層.ListIndex <> -1) And _
        (Me.lbl_ヤードコード.Caption <> "") And _
        (Me.txt_ボックス番号.VALUE <> "")) And pboolFrmActive Then
        
        pstrFloorLevel = Me.cmb_設置階層.Column(1, Me.cmb_設置階層.ListIndex)
        sinSize = CSng(pstrSize) / CSng(pstrDoor)
        
        strSQL = ""
        strSQL = strSQL & " SELECT '       （未割付）' AS TEXTDATA , 0 as CNO , 0 as SIZE , 0 as ORDER_NO FROM CNTA_MAST "
        strSQL = strSQL & " UNION "
        strSQL = strSQL & " (SELECT "
        strSQL = strSQL & " RIGHT('000000' + CAST(CNTA_MAST.CNTA_NO AS varchar), 6) + ' ' + NAME1.NAME_RYAK + ' ' + CAST(CNTA_MAST.CNTA_SIZE AS varchar) + '帖' + ' ' + NAME2.NAME_RYAK  AS TEXTDATA "
        strSQL = strSQL & " ,CNTA_MAST.CNTA_NO AS CNO,CNTA_MAST.CNTA_SIZE as SIZE , 1 as ORDER_NO "
        strSQL = strSQL & "  FROM (CNTA_MAST as CNTA_MAST "
        strSQL = strSQL & "  LEFT JOIN (select NAME_MAST1.NAME_RYAK,NAME_MAST1.NAME_CODE from NAME_MAST as NAME_MAST1 where (((NAME_MAST1.NAME_CODE) <> 10) AND ((NAME_MAST1.NAME_CODE) <> 30) AND ((NAME_MAST1.NAME_ID) = '086'))) as NAME1 on (NAME1.NAME_CODE = CNTA_MAST.CNTA_USAGE)) "
        strSQL = strSQL & "  LEFT JOIN (select NAME_MAST2.NAME_RYAK,NAME_MAST2.NAME_CODE from NAME_MAST as NAME_MAST2 where ((NAME_MAST2.NAME_ID)='090')) as NAME2 on (NAME2.NAME_CODE = CNTA_MAST.CNTA_STEP) "
        strSQL = strSQL & "  ,CBRA_TRAN "
        strSQL = strSQL & "  WHERE CNTA_MAST.CNTA_NO = CBRA_TRAN.CBRAT_CNO "
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_CODE = CBRA_TRAN.CBRAT_YCODE "
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_STEP = " & pstrFloorLevel & " "
        strSQL = strSQL & "    AND CBRA_TRAN.CBRAT_BOXNO = '" & pstrOldBoxNo & "'"
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "    AND CBRA_TRAN.CBRAT_EDANO = '" & pstrOldEdaNo & "'"
'↑INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "    AND ((CBRA_TRAN.CBRAT_YCODE) = " & Me.lbl_ヤードコード.Caption & " ) "
        strSQL = strSQL & " UNION ALL "
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " RIGHT('000000' + CAST(CNTA_MAST.CNTA_NO AS varchar), 6) + ' ' + NAME1.NAME_RYAK + ' ' + CAST(CNTA_MAST.CNTA_SIZE AS varchar) + '帖' + ' ' + NAME2.NAME_RYAK  AS TEXTDATA "
        strSQL = strSQL & " ,CNTA_MAST.CNTA_NO AS CNO ,CNTA_MAST.CNTA_SIZE as SIZE , 2 as ORDER_NO "
        strSQL = strSQL & "  FROM (CNTA_MAST as CNTA_MAST "
        strSQL = strSQL & "  LEFT JOIN (select NAME_MAST1.NAME_RYAK,NAME_MAST1.NAME_CODE from NAME_MAST as NAME_MAST1 where (((NAME_MAST1.NAME_CODE) <> 10) AND ((NAME_MAST1.NAME_CODE) <> 30) AND ((NAME_MAST1.NAME_ID) = '086'))) as NAME1 on (NAME1.NAME_CODE = CNTA_MAST.CNTA_USAGE)) "
        strSQL = strSQL & "  LEFT JOIN (select NAME_MAST2.NAME_RYAK,NAME_MAST2.NAME_CODE from NAME_MAST as NAME_MAST2 where ((NAME_MAST2.NAME_ID)='090')) as NAME2 on (NAME2.NAME_CODE = CNTA_MAST.CNTA_STEP) "
        strSQL = strSQL & "  WHERE CNTA_MAST.CNTA_NO NOT IN (SELECT CNTA_MAST2.CNTA_NO FROM CBRA_TRAN,CNTA_MAST as CNTA_MAST2 "
        strSQL = strSQL & "                    WHERE CBRA_TRAN.CBRAT_CNO = CNTA_MAST2.CNTA_NO"
        strSQL = strSQL & "                      AND CBRA_TRAN.CBRAT_YCODE = CNTA_MAST2.CNTA_CODE"
        strSQL = strSQL & "                      AND CNTA_MAST2.CNTA_USE <> 9 "
        strSQL = strSQL & "                      AND CNTA_MAST2.CNTA_STEP = " & pstrFloorLevel & " "
        strSQL = strSQL & "                      AND ((CNTA_MAST2.CNTA_CODE) = " & Me.lbl_ヤードコード.Caption & " ) )"
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_SIZE >= " & sinSize & " "
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_USE <> 9 "
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_STEP = " & pstrFloorLevel & " "
        strSQL = strSQL & "    AND ((CNTA_MAST.CNTA_CODE) = " & Me.lbl_ヤードコード.Caption & " ) "
        strSQL = strSQL & " UNION ALL "
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " RIGHT('000000' + CAST(CNTA_MAST.CNTA_NO AS varchar), 6) + ' ' + NAME1.NAME_RYAK + ' ' + CAST(CNTA_MAST.CNTA_SIZE AS varchar) + '帖' + ' ' + NAME2.NAME_RYAK  AS TEXTDATA "
        strSQL = strSQL & " ,CNTA_MAST.CNTA_NO AS CNO ,CNTA_MAST.CNTA_SIZE as SIZE , 3 as ORDER_NO "
        strSQL = strSQL & "  FROM (CNTA_MAST as CNTA_MAST "
        strSQL = strSQL & "  LEFT JOIN (select NAME_MAST1.NAME_RYAK,NAME_MAST1.NAME_CODE from NAME_MAST as NAME_MAST1 where (((NAME_MAST1.NAME_CODE) <> 10) AND ((NAME_MAST1.NAME_CODE) <> 30) AND ((NAME_MAST1.NAME_ID) = '086'))) as NAME1 on (NAME1.NAME_CODE = CNTA_MAST.CNTA_USAGE)) "
        strSQL = strSQL & "  LEFT JOIN (select NAME_MAST2.NAME_RYAK,NAME_MAST2.NAME_CODE from NAME_MAST as NAME_MAST2 where ((NAME_MAST2.NAME_ID)='090')) as NAME2 on (NAME2.NAME_CODE = CNTA_MAST.CNTA_STEP) "
        strSQL = strSQL & "  WHERE CNTA_MAST.CNTA_NO NOT IN (SELECT CNTA_MAST2.CNTA_NO FROM CBRA_TRAN,CNTA_MAST as CNTA_MAST2 "
        strSQL = strSQL & "                    WHERE CBRA_TRAN.CBRAT_CNO = CNTA_MAST2.CNTA_NO"
        strSQL = strSQL & "                      AND CBRA_TRAN.CBRAT_YCODE = CNTA_MAST2.CNTA_CODE"
        strSQL = strSQL & "                      AND CNTA_MAST2.CNTA_USE <> 9 "
        strSQL = strSQL & "                      AND CNTA_MAST2.CNTA_STEP = " & pstrFloorLevel & " "
        strSQL = strSQL & "                      AND ((CNTA_MAST2.CNTA_CODE) = " & Me.lbl_ヤードコード.Caption & " ) )"
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_SIZE < " & sinSize & " "
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_USE <> 9 "
        strSQL = strSQL & "    AND CNTA_MAST.CNTA_STEP = " & pstrFloorLevel & " "
        strSQL = strSQL & "    AND ((CNTA_MAST.CNTA_CODE) = " & Me.lbl_ヤードコード.Caption & " ) )"
        strSQL = strSQL & "  ORDER BY ORDER_NO , SIZE;"

        Me.sub_部屋割付.Form![cmb_部屋割付].RowSource = ""
        Call Form_FKS800.dropWorkTable(3)
        Call Form_FKS800.createWorkTable(3)
      
        Set rsComboBoxData = pobjDB.OpenRecordset(P_COMBO_BOX_WORK_TABLE, dbOpenDynaset, dbAppendOnly)
        Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
        
        With rsData
            Do Until .EOF
                '他ボックスで選択済みのデータ以外を追加
                strSQL = ""
                strSQL = strSQL & "SELECT * FROM " & P_ROOM_WORK_TABLE
                strSQL = strSQL & " WHERE CNO = " & .Fields("CNO")
                strSQL = strSQL & " AND   CNO <> 0 "
                Set rsDataLocal = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
                If rsDataLocal.EOF Then
                    rsComboBoxData.AddNew
                    rsComboBoxData.Fields("TEXTDATA") = .Fields("TEXTDATA")
                    rsComboBoxData.Fields("CNO") = .Fields("CNO")
                    rsComboBoxData.Fields("SIZE") = .Fields("SIZE")
                    rsComboBoxData.UPDATE
                End If
                rsDataLocal.Close
                Set rsDataLocal = Nothing
                .MoveNext
            Loop
            rsData.Close
            rsComboBoxData.Close
            Set rsData = Nothing
            Set rsComboBoxData = Nothing
        End With
        Me.sub_部屋割付.Form![cmb_部屋割付].RowSource = "SELECT TEXTDATA, CNO, SIZE FROM " & P_COMBO_BOX_WORK_TABLE & ";"
        Me.sub_部屋割付.Form![cmb_部屋割付].Requery
        Me.sub_部屋割付.Form![lbl_ドア数].Caption = pstrDoor
        sub_部屋割付.Enabled = True
    Else
        sub_部屋割付.Enabled = False
    End If
    
    
    Exit Sub

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
'    If Not rsDataLocal Is Nothing Then: rsDataLocal.Close: Set rsDataLocal = Nothing
    If Not rsComboBoxData Is Nothing Then: rsComboBoxData.Close: Set rsComboBoxData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "setKontComboDataSQL" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :fncInputCheck
'        NOTE             :objAdoDbConnection コネクション
'        CREATE_DATE      :2008/06/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck(objAdoDbConnection As Object) As Boolean
    
    Dim strSQL             As String
    Dim rsData             As Object
    Dim rsDataLocal        As Object
    Dim strNewYCode        As String
    Dim strNewBoxNo        As String
    Dim iRoomCount         As Integer
    Dim dRoomSize          As Double
    Dim bRoomOver          As Boolean
    Dim strSetDate         As String
'↓INSERT 2013/11/18 MIYAMOTO
    Dim strNewEdaNo        As String
    Dim iBunkatuSu         As Integer
'↑INSERT 2013/11/18 MIYAMOTO
'↓INSERT 2018/07/17 PARK
    Dim iHukusyaSu         As Integer
'↑INSERT 2018/07/17 PARK

    On Error GoTo Exception
    
    '初期値設定
    fncInputCheck = False
    
    '必須入力チェック
    If pstrProcMode = "2" Then
        '移動
        If Trim(Nz(Me.cmb_ヤード.VALUE)) = "" Then
            MsgBox "ヤードは必須入力です。", vbInformation, "確認"
            Me.cmb_ヤード.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_ヤード.VALUE)) = pstrOldYCode _
        And pstrFloorLevel = pstrOldFloorLevel Then
            MsgBox "ヤードまたは設置階層が変更されていません。" & vbCrLf & "移動先のヤード、設置階層を選択して下さい。", vbInformation, "確認"
            Me.cmb_ヤード.SetFocus
            Exit Function
        End If
        If Trim(Nz(pstrSetDate)) >= Trim(Nz(Me.txt_設置日.VALUE)) Then
            MsgBox "" & pstrSetDate & "　＜　設置日として下さい。", vbInformation, "確認"
            Me.txt_設置日.SetFocus
            Exit Function
        End If
    Else
        '新規・修正
        If Trim(Nz(Me.lbl_ヤードコード.Caption)) = "" Then
            Exit Function
        End If
        If Not fncRequireControl(Me.txt_ボックス番号, _
                                Nz(Me.txt_ボックス番号.VALUE), _
                                "ボックス番号", _
                                16) _
            Then Exit Function

'↓INSERT 2018/07/17 PARK
        If Trim(Nz(Me.txt_複写件数.VALUE)) = "" Then
            MsgBox "複写件数は必須入力です。", vbInformation, "確認"
            Me.txt_複写件数.SetFocus
            Exit Function
        Else
            If Not IsNumeric(Trim(Nz(Me.txt_複写件数.VALUE))) Then
                MsgBox "複写件数は数値を入力してください。", vbInformation, "確認"
                Me.txt_複写件数.SetFocus
                Exit Function
            Else
                iHukusyaSu = Trim(Val(Nz(Me.txt_複写件数.VALUE)))
                If (iHukusyaSu < 1 Or iHukusyaSu > 1000) Then
                    MsgBox "複写件数は1～1000を入力してください。", vbInformation, "確認"
                    Me.txt_複写件数.SetFocus
                    Exit Function
                End If
                If (Form_FKS800.gstrAutoBoxNoKaHi <> -1 And iHukusyaSu > 1) Then
                    MsgBox "複写件数が複数の場合は自動発番をチェックしてください。", vbInformation, "確認"
                    Me.txt_複写件数.SetFocus
                    Exit Function
                End If
            End If
        End If
'↑INSERT 2018/07/17 PARK
'↓INSERT 2013/11/18 MIYAMOTO
        If Trim(Nz(Me.txt_枝番.VALUE)) = "" Then
            MsgBox "枝番は必須入力です。", vbInformation, "確認"
            Me.txt_枝番.SetFocus
            Exit Function
        Else
            If Not IsNumeric(Trim(Nz(Me.txt_枝番.VALUE))) Then
                MsgBox "枝番は数値を入力してください。", vbInformation, "確認"
                Me.txt_枝番.SetFocus
                Exit Function
            Else
                iBunkatuSu = Trim(Val(Nz(Me.txt_枝番.VALUE)))
                If (iBunkatuSu < 0 Or iBunkatuSu > 99) Then
                    MsgBox "枝番は00～99を入力してください。", vbInformation, "確認"
                    Me.txt_枝番.SetFocus
                    Exit Function
                End If
                Me.txt_枝番 = Format(Me.txt_枝番, "00")
            End If
        End If
'↑INSERT 2013/11/18 MIYAMOTO
        If Trim(Nz(Me.cmb_タイプ.VALUE)) = "" Then
            MsgBox "タイプは必須入力です。", vbInformation, "確認"
            Me.cmb_タイプ.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_ドア数.VALUE)) = "" Then
            MsgBox "ドア数は必須入力です。", vbInformation, "確認"
            Me.cmb_ドア数.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_内装.VALUE)) = "" Then
            MsgBox "内装は必須入力です。", vbInformation, "確認"
            Me.cmb_内装.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_ドア位置.VALUE)) = "" Then
            MsgBox "ドア位置は必須入力です。", vbInformation, "確認"
            Me.cmb_ドア位置.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_サイズ.VALUE)) = "" Then
            MsgBox "サイズは必須入力です。", vbInformation, "確認"
            Me.cmb_サイズ.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_ドアタイプ.VALUE)) = "" Then
            MsgBox "ドアタイプは必須入力です。", vbInformation, "確認"
            Me.cmb_ドアタイプ.SetFocus
            Exit Function
        End If
        If Not fncRequireControl(Me.txt_入庫日, _
                                Nz(Me.txt_入庫日.VALUE), _
                                "入庫日", _
                                10) _
            Then Exit Function
        If Not IsDate(Nz(Me.txt_入庫日.VALUE)) Then
            MsgBox "入庫日はYYYY/MM/DD形式です。", vbInformation, "確認"
            Me.txt_入庫日.SetFocus
            Exit Function
        End If
'↓INSERT 2013/11/18 MIYAMOTO
        If Trim(Nz(Me.cmb_所有.VALUE)) = "" Then
            MsgBox "所有は必須入力です。", vbInformation, "確認"
            Me.cmb_所有.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_購入区分.VALUE)) = "" Then
            MsgBox "購入区分は必須入力です。", vbInformation, "確認"
            Me.cmb_購入区分.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_コンテナ分類.VALUE)) = "" Then
            MsgBox "コンテナ分類は必須入力です。", vbInformation, "確認"
            Me.cmb_コンテナ分類.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.cmb_設置用途.VALUE)) = "" Then
            MsgBox "設置用途は必須入力です。", vbInformation, "確認"
            Me.cmb_設置用途.SetFocus
            Exit Function
        End If
'当バージョン(0.1)では非表示とする
'        If Trim(Nz(Me.cmb_建築確認.VALUE)) = "" Then
'            MsgBox "建築確認は必須入力です。", vbInformation, "確認"
'            Me.cmb_建築確認.SetFocus
'            Exit Function
'        End If
        If Trim(Nz(Me.cmb_コンテナ内通路.VALUE)) = "" Then
            MsgBox "コンテナ内通路は必須入力です。", vbInformation, "確認"
            Me.cmb_コンテナ内通路.SetFocus
            Exit Function
        End If
        If Trim(Nz(Me.txt_分割数.VALUE)) = "" Then
            MsgBox "分割数は必須入力です。", vbInformation, "確認"
            Me.txt_分割数.SetFocus
            Exit Function
        Else
            If Not IsNumeric(Trim(Nz(Me.txt_分割数.VALUE))) Then
                MsgBox "分割数は数値を入力してください。", vbInformation, "確認"
                Me.txt_分割数.SetFocus
                Exit Function
            Else
                iBunkatuSu = Trim(Val(Nz(Me.txt_分割数.VALUE)))
                If (iBunkatuSu < 0 Or iBunkatuSu > 99) Then
                    MsgBox "分割数は0～99を入力してください。", vbInformation, "確認"
                    Me.txt_分割数.SetFocus
                    Exit Function
                End If
            End If
        End If
'↑INSERT 2013/11/18 MIYAMOTO
    End If
    If Not fncRequireControl(Me.txt_設置日, _
                            Nz(Me.txt_設置日.VALUE), _
                            "設置日", _
                            10) _
        Then Exit Function
    If Not IsDate(Nz(Me.txt_設置日.VALUE)) Then
        MsgBox "設置日はYYYY/MM/DD形式です。", vbInformation, "確認"
        Me.txt_設置日.SetFocus
        Exit Function
    End If
    If Trim(Nz(Me.cmb_設置階層.VALUE)) = "" Then
        MsgBox "設置階層は必須入力です。", vbInformation, "確認"
        Me.cmb_設置階層.SetFocus
        Exit Function
    End If
    If Trim(Nz(Me.txt_入庫日.VALUE)) > Trim(Nz(Me.txt_設置日.VALUE)) Then
        MsgBox "入庫日　≦　設置日として下さい。", vbInformation, "確認"
        Me.txt_設置日.SetFocus
        Exit Function
    End If
    If (CmLenb(Me.txt_備考.VALUE) > 50) Then
        MsgBox "備考は桁数オーバーです。", vbInformation, "確認"
        Me.txt_備考.SetFocus
        Exit Function
    End If
    '初期値設定
    iRoomCount = 0
    dRoomSize = Nz(pstrSize) / Nz(pstrDoor)
    If pstrProcMode = "2" Then
        '移動
        strNewYCode = Me.cmb_ヤード
        strNewBoxNo = Me.txt_ボックス番号
'↓INSERT 2013/11/18 MIYAMOTO
        strNewEdaNo = Me.txt_枝番
'↑INSERT 2013/11/18 MIYAMOTO
    Else
        '新規・修正
        strNewYCode = Me.lbl_ヤードコード.Caption
        strNewBoxNo = Me.txt_ボックス番号
'↓INSERT 2013/11/18 MIYAMOTO
        strNewEdaNo = Me.txt_枝番
'↑INSERT 2013/11/18 MIYAMOTO
    End If
    'コンテナ重複チェック
    strSQL = ""
    strSQL = strSQL & "SELECT CNO, COUNT(*) AS DATA_COUNT FROM " & P_ROOM_WORK_TABLE
    strSQL = strSQL & " WHERE CNO <> 0"
    strSQL = strSQL & " GROUP BY CNO"
    Set rsDataLocal = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    While Not rsDataLocal.EOF
        If rsDataLocal("DATA_COUNT") > 1 Then
            MsgBox "コンテナ " & Format$(rsDataLocal.Fields("CNO"), "000000") & " が二つ割付けられています。", vbInformation, "確認"
            Exit Function
        End If
        rsDataLocal.MoveNext
    Wend
    rsDataLocal.Close
    Set rsDataLocal = Nothing
    'ボックス番号重複チェック
'↓UPDATE 2013/11/18 MIYAMOTO
'    If Trim(pstrOldBoxNo) <> strNewBoxNo Then
    If Trim(pstrOldBoxNo) & Trim(pstrOldEdaNo) <> strNewBoxNo & strNewEdaNo Then
'↑UPDATE 2013/11/18 MIYAMOTO
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBOX_MAST "
        strSQL = strSQL & " WHERE CBOXM_BOXNO = '" & strNewBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "   AND CBOXM_EDANO = '" & strNewEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsData.EOF Then
'↓UPDATE 2013/11/18 MIYAMOTO
'            MsgBox "ボックス番号が重複しています。" + vbCrLf + "重複しない番号を入力して下さい。", vbInformation, "確認"
            MsgBox "ボックス番号・枝番が重複しています。" + vbCrLf + "重複しない番号を入力して下さい。", vbInformation, "確認"
'↑UPDATE 2013/11/18 MIYAMOTO
            Me.txt_ボックス番号.SetFocus
            Exit Function
        End If
        rsData.Close
        Set rsData = Nothing
    End If
    'コンテナ撤去チェック
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & P_ROOM_WORK_TABLE
    strSQL = strSQL & " WHERE CNO <> 0"
    Set rsDataLocal = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    While Not rsDataLocal.EOF
        iRoomCount = iRoomCount + 1
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CNTA_MAST "
        strSQL = strSQL & " WHERE CNTA_CODE = " & strNewYCode
        strSQL = strSQL & " AND   CNTA_NO   = " & rsDataLocal.Fields("CNO")
        strSQL = strSQL & " AND   CNTA_STEP = " & Nz(pstrFloorLevel)
        strSQL = strSQL & " AND   CNTA_USE  = 9 "
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsData.EOF Then
            MsgBox Format$(rsDataLocal.Fields("CNO"), "000000") & " は既に撤去済みのため割付出来ません。", vbInformation, "確認"
            Exit Function
        End If
        rsData.Close
        Set rsData = Nothing
        '他ボックス割付済みチェック
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBRA_TRAN "
'↓UPDATE 2013/11/18 MIYAMOTO
'        strSQL = strSQL & " WHERE CBRAT_BOXNO <> '" & strNewBoxNo & "' "
'        strSQL = strSQL & " AND   CBRAT_BOXNO <> '" & pstrOldBoxNo & "' "
        strSQL = strSQL & " WHERE CBRAT_BOXNO + CBRAT_EDANO <> '" & strNewBoxNo & strNewEdaNo & "' "
        strSQL = strSQL & " AND   CBRAT_BOXNO + CBRAT_EDANO <> '" & pstrOldBoxNo & pstrOldEdaNo & "' "
'↑UPDATE 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBRAT_YCODE =  '" & strNewYCode & "' "
        strSQL = strSQL & " AND   CBRAT_CNO    =  " & rsDataLocal.Fields("CNO")
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsData.EOF Then
            MsgBox Format$(rsDataLocal.Fields("CNO"), "000000") & " は他ボックスで既に割付されています。", vbInformation, "確認"
            Exit Function
        End If
        rsData.Close
        Set rsData = Nothing
        '実貼オーバーチェック
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CNTA_MAST "
        strSQL = strSQL & " WHERE CNTA_CODE = " & strNewYCode
        strSQL = strSQL & " AND   CNTA_NO   = " & rsDataLocal.Fields("CNO")
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsData.EOF Then
            If dRoomSize < rsData.Fields("CNTA_SIZE").VALUE Then
                bRoomOver = True
            End If
        End If
        rsDataLocal.MoveNext
    Wend
    rsDataLocal.Close
    Set rsDataLocal = Nothing
    'ドア数チェック
    If Nz(pstrDoor) < iRoomCount Then
        MsgBox "指定した部屋（ドア）数以上の部屋は割付できません。", vbInformation, "確認"
        If pstrProcMode <> "2" Then
            Me.cmb_ドア数.SetFocus
        End If
        Exit Function
    End If
    If pstrProcMode = "2" Then
        '移動時のみ
        '設置日重複チェック
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBYS_TRAN "
        strSQL = strSQL & " WHERE CBYST_BOXNO = '" & pstrOldBoxNo & "'"
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "   AND CBYST_EDANO = '" & pstrOldEdaNo & "'"
'↑INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBYST_YCODE = '" & pstrOldYCode & "'"
        strSQL = strSQL & " AND   CBYST_SETDATE = '" & Format(Me.txt_設置日, "yyyymmdd") & "' "
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsData.EOF Then
            MsgBox "指定した設置日にて既に登録済みです。" + vbCrLf + "設置日を変更して下さい。", vbInformation, "確認"
            Me.txt_設置日.SetFocus
            Exit Function
        End If
        rsData.Close
        Set rsData = Nothing
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBYS_RECD "
        strSQL = strSQL & " WHERE CBYSR_BOXNO = '" & strNewBoxNo & "'"
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "   AND CBYSR_EDANO = '" & strNewEdaNo & "'"
'↑INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBYSR_YCODE = '" & strNewYCode & "'"
        strSQL = strSQL & " AND   CBYSR_SETDATE = '" & Format(Me.txt_設置日, "yyyymmdd") & "' "
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsData.EOF Then
            MsgBox "指定した設置日にて既に登録済みです。" + vbCrLf + "設置日を変更して下さい。", vbInformation, "確認"
            Me.txt_設置日.SetFocus
            Exit Function
        End If
        rsData.Close
        Set rsData = Nothing
    End If
    '修正時、履歴に登録済みの設置日より過去の入力不可チェック
    If pstrProcMode = "1" Then
        strSQL = ""
        strSQL = strSQL & "SELECT MAX(CBYSR_SETDATE) MAX_DATE FROM CBYS_RECD "
        strSQL = strSQL & " WHERE CBYSR_BOXNO = '" & strNewBoxNo & "'"
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "   AND CBYSR_EDANO = '" & strNewEdaNo & "'"
'↑INSERT 2013/11/18 MIYAMOTO
'↓UPDATE 2013/11/18 MIYAMOTO
'        strSQL = strSQL & " GROUP BY CBYSR_BOXNO                    "
        strSQL = strSQL & " GROUP BY CBYSR_BOXNO,CBYSR_EDANO        "
'↑UPDATE 2013/11/18 MIYAMOTO
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If Not rsData.EOF Then
            If rsData.Fields("MAX_DATE").VALUE >= Trim(Nz(Format(Me.txt_設置日.VALUE, "yyyymmdd"))) Then
                strSetDate = DateSerial(Left(rsData.Fields("MAX_DATE").VALUE, 4), Mid(rsData.Fields("MAX_DATE").VALUE, 5, 2), Right(rsData.Fields("MAX_DATE").VALUE, 2))
                MsgBox "" & strSetDate & "　＜　設置日として下さい。", vbInformation, "確認"
                Me.txt_設置日.SetFocus
                Exit Function
            End If
        End If
        rsData.Close
        Set rsData = Nothing
    End If
    '実貼オーバー確認
    If bRoomOver Then
        If MsgBox("１部屋あたり" & Format$(dRoomSize, "0.00") & "帖です。" & vbCrLf & Format$(dRoomSize, "0.00") & "帖以上の部屋を割り付けていますがよろしいですか？", vbInformation + vbYesNo, "確認") <> vbYes Then
            Exit Function
        End If
    End If
    
    fncInputCheck = True
    
    Exit Function

Exception:
    
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    If Not rsDataLocal Is Nothing Then
        rsDataLocal.Close
        Set rsDataLocal = Nothing
    End If
    Call Err.Raise(Err.Number, "fncInputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :必須チェックコントロール
'        MODULE_ID        :fncRequireControl
'        NOTE             :引数からの値を空データかチェックする
'                          objControl    フォーカスをセットするコントロール
'                          strValue      必須チェックする元
'                          strMsgText    エラーメッセージで使用
'        CREATE_DATE      :2008/06/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRequireControl( _
                    objControl As Object, _
                    strValue As String, _
                    strMsgText As String, _
                    intLength As Integer) As Boolean

    On Error GoTo Exception

    fncRequireControl = False
    
    If (strValue = "") Then
        MsgBox strMsgText & "は必須入力です。", vbInformation, "確認"
        objControl.SetFocus
        Exit Function
    End If
    
    If (CmLenb(strValue) > intLength) Then
        MsgBox strMsgText & "は桁数オーバーです。", vbInformation, "確認"
        objControl.SetFocus
        Exit Function
    End If
    
    fncRequireControl = True
    
    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncRequireControl" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :txt_ボックス番号_AfterUpdate
'        MODULE_ID        :txt_ボックス番号_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_ボックス番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Char2(Me.txt_ボックス番号, KeyAscii, 16)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_設置日_BeforeUpdate
'        MODULE_ID        :txt_設置日_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_設置日_BeforeUpdate(Cancel As Integer)
    If Nz(Me.txt_設置日) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_設置日.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_設置日_AfterUpdate
'        MODULE_ID        :txt_設置日_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_設置日_AfterUpdate()
    Me.txt_設置日.VALUE = pstrDateValue
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_入庫日_BeforeUpdate
'        MODULE_ID        :txt_入庫日_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_入庫日_BeforeUpdate(Cancel As Integer)
    If Nz(Me.txt_入庫日.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_入庫日.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "日付を入力して下さい。（YYYY/MM/DD）"
            Cancel = True
            SendKeys "{ESC}"
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_入庫日_AfterUpdate
'        MODULE_ID        :txt_入庫日_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_入庫日_AfterUpdate()
    Me.txt_入庫日.VALUE = pstrDateValue
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 新規追加処理
'        MODULE_ID        : INS_RTN
'        PARAMETER        : objAdoDbConnection コネクション
'        戻り値           : True=正常 False=異常
'        CREATE_DATE      : 2008/06/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function INS_RTN(objAdoDbConnection As Object) As Boolean
  
    Dim strSQL             As String
    Dim rsData             As Object
    Dim rsDataLocal        As Object       'ワークファイル（部屋割付情報）読み込み用
    Dim strDate            As String
    Dim strTime            As String
'↓INSERT 2018/07/17 PARK
    Dim strBoxNumber       As String
    Dim iHukusyaSu         As Integer
    Dim intCount           As Integer
    Dim iBoxNumber         As Integer
    Dim strPrefix          As String
    Dim iAdd               As Integer
    Dim bEnd               As Boolean
    Dim rsDataChk          As Object
'↑INSERT 2018/07/17 PARK

On Error GoTo Exception
    
    INS_RTN = False

    '初期値設定
    strDate = Format(DATE, "yyyymmdd")
    strTime = Format(Now, "hhmmss")
'↓INSERT 2018/07/17 PARK
    strBoxNumber = Me.txt_ボックス番号
    iHukusyaSu = Me.txt_複写件数
    
    intCount = 0
    iAdd = 1
    strPrefix = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'FKS800' AND INTIF_RECFB = 'AUTO_PREFIX_" & Format(Me.cmb_設置用途.Column(1, Me.cmb_設置用途.ListIndex), "00") & "'"), "")
    iBoxNumber = Val(Replace(strBoxNumber, strPrefix, "", , , vbTextCompare))
    For intCount = 1 To iHukusyaSu
        bEnd = False
        If intCount > 1 Then
            While Not bEnd
                strBoxNumber = strPrefix & (iBoxNumber + iAdd)
                '発番したボックス番号が存在するか確認
                strSQL = ""
                strSQL = strSQL & "SELECT * FROM CBOX_MAST "
                strSQL = strSQL & " WHERE CBOXM_BOXNO = '" & strBoxNumber & "'"
                Set rsDataChk = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoAppendOnly)
                bEnd = rsDataChk.EOF
                rsDataChk.Close
                Set rsDataChk = Nothing
                iAdd = iAdd + 1
            Wend
        End If
'↑INSERT 2018/07/17 PARK

    'コンテナボックスマスタ登録
    Set rsData = MSZZ025.ADODB_Recordset("CBOX_MAST", objAdoDbConnection, adoAppendOnly)
    rsData.AddNew
    rsData.Fields("CBOXM_BOXNO") = strBoxNumber
'↓INSERT 2013/11/18 MIYAMOTO
    rsData.Fields("CBOXM_EDANO") = Me.txt_枝番
'↑INSERT 2013/11/18 MIYAMOTO
    rsData.Fields("CBOXM_TYPE") = Me.cmb_タイプ.Column(1, Me.cmb_タイプ.ListIndex)
    rsData.Fields("CBOXM_INTYPE") = Me.cmb_内装.Column(1, Me.cmb_内装.ListIndex)
    rsData.Fields("CBOXM_SIZE") = Me.cmb_サイズ.Column(2, Me.cmb_サイズ.ListIndex)
    rsData.Fields("CBOXM_DOOR") = Me.cmb_ドア数.Column(1, Me.cmb_ドア数.ListIndex)
    rsData.Fields("CBOXM_DOORPSTN") = Me.cmb_ドア位置.Column(1, Me.cmb_ドア位置.ListIndex)
    rsData.Fields("CBOXM_DOORTYPE") = Me.cmb_ドアタイプ.Column(1, Me.cmb_ドアタイプ.ListIndex)
    rsData.Fields("CBOXM_STOCKDATE") = Format(txt_入庫日, "yyyymmdd")
    rsData.Fields("CBOXM_STATUS") = 0
'↓INSERT 2013/11/18 MIYAMOTO
    rsData.Fields("CBOXM_BKTSU") = Me.txt_分割数
    rsData.Fields("CBOXM_KNKBN") = Me.cmb_購入区分.Column(1, Me.cmb_購入区分.ListIndex)
    rsData.Fields("CBOXM_CTBNR") = Me.cmb_コンテナ分類.Column(1, Me.cmb_コンテナ分類.ListIndex)
    rsData.Fields("CBOXM_CTYTO") = Me.cmb_設置用途.Column(1, Me.cmb_設置用途.ListIndex)
'当バージョン(0.1)では非表示とする
'    rsData.Fields("CBOXM_KNKAK") = Me.cmb_建築確認.Column(1, Me.cmb_建築確認.ListIndex)
    rsData.Fields("CBOXM_KNKAK") = "0"      'とりあえず"0"をセット
    rsData.Fields("CBOXM_SHSHU") = Me.cmb_コンテナ内通路.Column(1, Me.cmb_コンテナ内通路.ListIndex)
    rsData.Fields("CBOXM_KRKBN") = Me.cmb_所有.Column(1, Me.cmb_所有.ListIndex)
'↑INSERT 2013/11/18 MIYAMOTO
    rsData.Fields("CBOXM_INSED") = strDate
    rsData.Fields("CBOXM_INSEJ") = strTime
    rsData.Fields("CBOXM_INSPB") = PROG_ID
    rsData.Fields("CBOXM_INSUB") = Form_FKS800.gstrUserName
    rsData.UPDATE

    rsData.Close
    Set rsData = Nothing
  
    'コンテナボックスヤード設置トラン登録
    Set rsData = MSZZ025.ADODB_Recordset("CBYS_TRAN", objAdoDbConnection, adoAppendOnly)
    rsData.AddNew
    rsData.Fields("CBYST_BOXNO") = strBoxNumber
'↓INSERT 2013/11/18 MIYAMOTO
    rsData.Fields("CBYST_EDANO") = Me.txt_枝番
'↑INSERT 2013/11/18 MIYAMOTO
'    If Me.tab_移動先指定.Value = 0 Then                                                'PHESE2にて対応
'    'ヤード情報                                                                        'PHESE2にて対応
        rsData.Fields("CBYST_SYUBETSU") = 0
        rsData.Fields("CBYST_YCODE") = Me.lbl_ヤードコード.Caption
'    Else                                                                              'PHESE2にて対応
'    'ボックス情報                                                                      'PHESE2にて対応
'        rsData.Fields("CBYST_SYUBETSU") = 1                                           'PHESE2にて対応
'        rsData.Fields("CBYST_YCODE") = Me.lbl_ボックス番号.Caption                     'PHESE2にて対応
'    End If                                                                            'PHESE2にて対応
    rsData.Fields("CBYST_STATUS") = 0
    rsData.Fields("CBYST_SETDATE") = Format(txt_設置日, "yyyymmdd")
    rsData.Fields("CBYST_SETFLOOR") = Me.cmb_設置階層.Column(1, Me.cmb_設置階層.ListIndex)
    rsData.Fields("CBYST_BIKON") = Me.txt_備考
    rsData.Fields("CBYST_INSED") = strDate
    rsData.Fields("CBYST_INSEJ") = strTime
    rsData.Fields("CBYST_INSPB") = PROG_ID
    rsData.Fields("CBYST_INSUB") = Form_FKS800.gstrUserName
    rsData.UPDATE
    
    rsData.Close
    Set rsData = Nothing
  
    'コンテナボックス部屋割付トラン登録
    '※部屋の割付があったときのみ登録（未割付は除外）
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM " & P_ROOM_WORK_TABLE
    strSQL = strSQL & " WHERE  CNO <> 0                         "

    Set rsDataLocal = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

    If Not rsDataLocal.EOF Then
        
        Set rsData = MSZZ025.ADODB_Recordset("CBRA_TRAN", objAdoDbConnection, adoAppendOnly)
    
        While Not rsDataLocal.EOF
            rsData.AddNew
            rsData.Fields("CBRAT_BOXNO") = Me.txt_ボックス番号
'↓INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAT_EDANO") = Me.txt_枝番
'↑INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAT_YCODE") = Me.lbl_ヤードコード.Caption
            rsData.Fields("CBRAT_CNO") = rsDataLocal.Fields("CNO")
            rsData.Fields("CBRAT_STATUS") = 0
            rsData.Fields("CBRAT_SETDATE") = Format(txt_設置日, "yyyymmdd")
            rsData.Fields("CBRAT_INSED") = strDate
            rsData.Fields("CBRAT_INSEJ") = strTime
            rsData.Fields("CBRAT_INSPB") = PROG_ID
            rsData.Fields("CBRAT_INSUB") = Form_FKS800.gstrUserName
            rsData.UPDATE
            rsDataLocal.MoveNext
        Wend
    
        rsData.Close
        Set rsData = Nothing

    End If
    

    rsDataLocal.Close
    Set rsDataLocal = Nothing

'↓INSERT 2018/07/17 PARK
    Next intCount
'↑INSERT 2018/07/17 PARK
    INS_RTN = True

Exception:

    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If

    If Not rsDataLocal Is Nothing Then
        rsDataLocal.Close
        Set rsDataLocal = Nothing
    End If

    If Err <> 0 Then
        Call Err.Raise(Err.Number, "INS_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      : ヤード移動処理
'        MODULE_ID        : MOVE_Y_RTN
'        PARAMETER        : objAdoDbConnection コネクション
'        戻り値　　　　　 : True=正常 False=異常
'        CREATE_DATE      : 2008/06/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MOVE_Y_RTN(objAdoDbConnection As Object) As Boolean
  
    Dim strSQL               As String
    Dim rsData               As Object
    Dim rsData2              As Object
    Dim rsDataLocal          As Object   'ワークファイル（部屋割付情報）読み込み用
    Dim strNewBoxNo          As String
'↓INSERT 2013/11/18 MIYAMOTO
    Dim strNewEdaNo          As String
'↑INSERT 2013/11/18 MIYAMOTO
    Dim strNewYCode          As String
    Dim strDate              As String
    Dim strTime              As String
  
On Error GoTo Exception
    
    MOVE_Y_RTN = False
    
    '初期値設定
    strDate = Format(DATE, "yyyymmdd")
    strTime = Format(Now, "hhmmss")

    strNewBoxNo = Me.txt_ボックス番号
'↓INSERT 2013/11/18 MIYAMOTO
    strNewEdaNo = Me.txt_枝番
'↑INSERT 2013/11/18 MIYAMOTO
    strNewYCode = Me.cmb_ヤード

    'コンテナボックスマスタ更新
'    strSQL = ""
'    strSQL = strSQL & "SELECT * FROM CBOX_MAST "
'    strSQL = strSQL & "WHERE CBOXM_BOXNO = '" & pstrOldBoxNo & "'"
'    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
'    If Not rsData.EOF Then
'        rsData.Fields("CBOXM_UPDAD") = strDate
'        rsData.Fields("CBOXM_UPDAJ") = strTime
'        rsData.Fields("CBOXM_UPDPB") = PROG_ID
'        rsData.Fields("CBOXM_UPDUB") = Form_FKS800.gstrUserName
'        rsData.UPDATE
'    End If
'    rsData.Close
'    Set rsData = Nothing
  
    'コンテナボックスヤード設置履歴トラン追加
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM CBYS_TRAN "
    strSQL = strSQL & " WHERE CBYST_BOXNO    = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & "   AND CBYST_EDANO    = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & " AND   CBYST_SYUBETSU = 0"
    strSQL = strSQL & " AND   CBYST_YCODE    = '" & pstrOldYCode & "' "
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If Not rsData.EOF Then
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBYS_RECD "
        strSQL = strSQL & "WHERE CBYSR_BOXNO = '" & strNewBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "AND CBYSR_EDANO = '" & strNewEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        Set rsData2 = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        While Not rsData.EOF
            rsData2.AddNew
            rsData2.Fields("CBYSR_BOXNO") = rsData.Fields("CBYST_BOXNO")
'↓INSERT 2013/11/18 MIYAMOTO
            rsData2.Fields("CBYSR_EDANO") = rsData.Fields("CBYST_EDANO")
'↑INSERT 2013/11/18 MIYAMOTO
            rsData2.Fields("CBYSR_SYUBETSU") = rsData.Fields("CBYST_SYUBETSU")
            rsData2.Fields("CBYSR_YCODE") = rsData.Fields("CBYST_YCODE")
            rsData2.Fields("CBYSR_SETDATE") = rsData.Fields("CBYST_SETDATE")
            rsData2.Fields("CBYSR_SETFLOOR") = rsData.Fields("CBYST_SETFLOOR")
            rsData2.Fields("CBYSR_BIKON") = rsData.Fields("CBYST_BIKON")
            rsData2.Fields("CBYSR_INSED") = strDate
            rsData2.Fields("CBYSR_INSEJ") = strTime
            rsData2.Fields("CBYSR_INSPB") = PROG_ID
            rsData2.Fields("CBYSR_INSUB") = Form_FKS800.gstrUserName
            rsData2.UPDATE
            '旧コンテナボックスヤード設置トラン削除
            rsData.Delete
            rsData.MoveNext
        Wend
        rsData2.Close
        Set rsData2 = Nothing
    End If
    rsData.Close
    Set rsData = Nothing

    'コンテナボックスヤード設置トラン追加
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM CBYS_TRAN "
    strSQL = strSQL & " WHERE CBYST_BOXNO    = '" & strNewBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & "   AND CBYST_EDANO    = '" & strNewEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & " AND   CBYST_SYUBETSU = 0"
    strSQL = strSQL & " AND   CBYST_YCODE    = " & strNewYCode
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If rsData.EOF Then
        rsData.AddNew
        rsData.Fields("CBYST_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBYST_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBYST_SYUBETSU") = 0
        rsData.Fields("CBYST_YCODE") = strNewYCode
        rsData.Fields("CBYST_STATUS") = 0
        rsData.Fields("CBYST_SETDATE") = Format(txt_設置日, "yyyymmdd")
        rsData.Fields("CBYST_SETFLOOR") = Me.cmb_設置階層.Column(1, Me.cmb_設置階層.ListIndex)
        rsData.Fields("CBYST_BIKON") = Me.txt_備考
        rsData.Fields("CBYST_INSED") = strDate
        rsData.Fields("CBYST_INSEJ") = strTime
        rsData.Fields("CBYST_INSPB") = PROG_ID
        rsData.Fields("CBYST_INSUB") = Form_FKS800.gstrUserName
        rsData.UPDATE
    End If
  
    'コンテナボックス部屋割付履歴トラン追加
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM CBRA_TRAN "
    strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & "   AND CBRAT_EDANO = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & " AND   CBRAT_YCODE =  " & pstrOldYCode
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If Not rsData.EOF Then
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBRA_RECD "
        strSQL = strSQL & " WHERE CBRAR_BOXNO = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "   AND CBRAR_EDANO = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBRAR_YCODE =  " & pstrOldYCode
        Set rsData2 = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        While Not rsData.EOF
            rsData2.AddNew
            rsData2.Fields("CBRAR_BOXNO") = rsData.Fields("CBRAT_BOXNO")
'↓INSERT 2013/11/18 MIYAMOTO
            rsData2.Fields("CBRAR_EDANO") = rsData.Fields("CBRAT_EDANO")
'↑INSERT 2013/11/18 MIYAMOTO
            rsData2.Fields("CBRAR_YCODE") = rsData.Fields("CBRAT_YCODE")
            rsData2.Fields("CBRAR_CNO") = rsData.Fields("CBRAT_CNO")
            rsData2.Fields("CBRAR_SETDATE") = rsData.Fields("CBRAT_SETDATE")
            rsData2.Fields("CBRAR_INSED") = strDate
            rsData2.Fields("CBRAR_INSEJ") = strTime
            rsData2.Fields("CBRAR_INSPB") = PROG_ID
            rsData2.Fields("CBRAR_INSUB") = Form_FKS800.gstrUserName
            rsData2.UPDATE
            '旧コンテナボックス部屋割付トラン削除
            rsData.Delete
            rsData.MoveNext
        Wend
        rsData2.Close
        Set rsData2 = Nothing
    End If
    rsData.Close
    Set rsData = Nothing
  
    'コンテナボックス部屋割付トラン追加
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM " & P_ROOM_WORK_TABLE
    strSQL = strSQL & " WHERE CNO <> 0"
    Set rsDataLocal = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rsDataLocal.EOF Then
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBRA_TRAN "
        strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & strNewBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "   AND CBRAT_EDANO = '" & strNewEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBRAT_YCODE =  " & strNewYCode
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        '割付入力あり
        While Not rsDataLocal.EOF
            rsData.AddNew
            rsData.Fields("CBRAT_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAT_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAT_YCODE") = strNewYCode
            rsData.Fields("CBRAT_CNO") = rsDataLocal.Fields("CNO")
            rsData.Fields("CBRAT_STATUS") = 0
            rsData.Fields("CBRAT_SETDATE") = Format(txt_設置日, "yyyymmdd")
            rsData.Fields("CBRAT_INSED") = strDate
            rsData.Fields("CBRAT_INSEJ") = strTime
            rsData.Fields("CBRAT_INSPB") = PROG_ID
            rsData.Fields("CBRAT_INSUB") = Form_FKS800.gstrUserName
            rsData.UPDATE
            rsDataLocal.MoveNext
        Wend
        rsData.Close
        Set rsData = Nothing
    End If
    rsDataLocal.Close
    Set rsDataLocal = Nothing
    
    '正常終了
    MOVE_Y_RTN = True
    
Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    If Not rsData2 Is Nothing Then
        rsData2.Close
        Set rsData2 = Nothing
    End If
    If Not rsDataLocal Is Nothing Then
        rsDataLocal.Close
        Set rsDataLocal = Nothing
    End If
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "MOVE_Y_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      : 修正処理
'        MODULE_ID        : UPD_RTN
'        PARAMETER        : objAdoDbConnection コネクション
'        戻り値　　　　　 : True=正常 False=異常
'        CREATE_DATE      : 2008/06/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function UPD_RTN(objAdoDbConnection As Object) As Boolean
  
    Dim strSQL               As String
    Dim rsData               As Object
    Dim rsData2              As Object
    Dim rsDataLocal          As Object   'ワークファイル（部屋割付情報）読み込み用
    Dim strNewBoxNo          As String
'↓INSERT 2013/11/18 MIYAMOTO
    Dim strNewEdaNo          As String
'↑INSERT 2013/11/18 MIYAMOTO
    Dim strNewYCode          As String
    Dim strDate              As String
    Dim strTime              As String
  
On Error GoTo Exception
    
    UPD_RTN = False

    '初期値設定
    strDate = Format(DATE, "yyyymmdd")
    strTime = Format(Now, "hhmmss")

    strNewBoxNo = Me.txt_ボックス番号
'↓INSERT 2013/11/18 MIYAMOTO
    strNewEdaNo = Me.txt_枝番
'↑INSERT 2013/11/18 MIYAMOTO
    strNewYCode = Me.lbl_ヤードコード.Caption

    'コンテナボックスマスタ更新
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM CBOX_MAST "
    strSQL = strSQL & "WHERE CBOXM_BOXNO = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & "AND CBOXM_EDANO = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If Not rsData.EOF Then
        rsData.Fields("CBOXM_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBOXM_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBOXM_TYPE") = Me.cmb_タイプ.Column(1, Me.cmb_タイプ.ListIndex)
        rsData.Fields("CBOXM_INTYPE") = Me.cmb_内装.Column(1, Me.cmb_内装.ListIndex)
        rsData.Fields("CBOXM_SIZE") = Me.cmb_サイズ.Column(2, Me.cmb_サイズ.ListIndex)
        rsData.Fields("CBOXM_DOOR") = Me.cmb_ドア数.Column(1, Me.cmb_ドア数.ListIndex)
        rsData.Fields("CBOXM_DOORPSTN") = Me.cmb_ドア位置.Column(1, Me.cmb_ドア位置.ListIndex)
        rsData.Fields("CBOXM_DOORTYPE") = Me.cmb_ドアタイプ.Column(1, Me.cmb_ドアタイプ.ListIndex)
        rsData.Fields("CBOXM_STOCKDATE") = Format(txt_入庫日, "yyyymmdd")
'↓INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBOXM_BKTSU") = Me.txt_分割数
        rsData.Fields("CBOXM_KNKBN") = Me.cmb_購入区分.Column(1, Me.cmb_購入区分.ListIndex)
        rsData.Fields("CBOXM_CTBNR") = Me.cmb_コンテナ分類.Column(1, Me.cmb_コンテナ分類.ListIndex)
        rsData.Fields("CBOXM_CTYTO") = Me.cmb_設置用途.Column(1, Me.cmb_設置用途.ListIndex)
'当バージョン(0.1)では非表示とする
'        rsData.Fields("CBOXM_KNKAK") = Me.cmb_建築確認.Column(1, Me.cmb_建築確認.ListIndex)
        rsData.Fields("CBOXM_SHSHU") = Me.cmb_コンテナ内通路.Column(1, Me.cmb_コンテナ内通路.ListIndex)
        rsData.Fields("CBOXM_KRKBN") = Me.cmb_所有.Column(1, Me.cmb_所有.ListIndex)
        If pstrOldBoxNo <> strNewBoxNo Then     '更新前後のボックス番号が異なる時に更新前ボックス番号をセット
            rsData.Fields("CBOXM_OLDBN") = pstrOldBoxNo
        End If
'↑INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBOXM_UPDAD") = strDate
        rsData.Fields("CBOXM_UPDAJ") = strTime
        rsData.Fields("CBOXM_UPDPB") = PROG_ID
        rsData.Fields("CBOXM_UPDUB") = Form_FKS800.gstrUserName
        rsData.UPDATE
    End If
    rsData.Close
    Set rsData = Nothing
  
    'ボックス番号を更新
'↓UPDATE 2013/11/18 MIYAMOTO
'    If pstrOldBoxNo <> strNewBoxNo Then
    If pstrOldBoxNo & pstrOldEdaNo <> strNewBoxNo & strNewEdaNo Then
'↑UPDATE 2013/11/18 MIYAMOTO
        'コンテナボックスヤード設置トラン
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBYS_TRAN "
        strSQL = strSQL & "WHERE CBYST_BOXNO = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "AND CBYST_EDANO = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        While Not rsData.EOF
            rsData.Fields("CBYST_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBYST_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBYST_UPDAD") = strDate
            rsData.Fields("CBYST_UPDAJ") = strTime
            rsData.Fields("CBYST_UPDPB") = PROG_ID
            rsData.Fields("CBYST_UPDUB") = Form_FKS800.gstrUserName
            rsData.UPDATE
            rsData.MoveNext
        Wend
        'コンテナボックスヤード設置履歴トラン
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBYS_RECD "
        strSQL = strSQL & "WHERE CBYSR_BOXNO = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "AND CBYSR_EDANO = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        While Not rsData.EOF
            rsData.Fields("CBYSR_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBYSR_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBYSR_UPDAD") = strDate
            rsData.Fields("CBYSR_UPDAJ") = strTime
            rsData.Fields("CBYSR_UPDPB") = PROG_ID
            rsData.Fields("CBYSR_UPDUB") = Form_FKS800.gstrUserName
            rsData.UPDATE
            rsData.MoveNext
        Wend
        'コンテナボックス部屋割付トラン
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBRA_TRAN "
        strSQL = strSQL & "WHERE CBRAT_BOXNO = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "AND CBRAT_EDANO = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        While Not rsData.EOF
            rsData.Fields("CBRAT_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAT_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAT_UPDAD") = strDate
            rsData.Fields("CBRAT_UPDAJ") = strTime
            rsData.Fields("CBRAT_UPDPB") = PROG_ID
            rsData.Fields("CBRAT_UPDUB") = Form_FKS800.gstrUserName
            rsData.UPDATE
            rsData.MoveNext
        Wend
        'コンテナボックス部屋割付履歴トラン
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBRA_RECD "
        strSQL = strSQL & "WHERE CBRAR_BOXNO = '" & pstrOldBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "AND CBRAR_EDANO = '" & pstrOldEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        While Not rsData.EOF
            rsData.Fields("CBRAR_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAR_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
            rsData.Fields("CBRAR_UPDAD") = strDate
            rsData.Fields("CBRAR_UPDAJ") = strTime
            rsData.Fields("CBRAR_UPDPB") = PROG_ID
            rsData.Fields("CBRAR_UPDUB") = Form_FKS800.gstrUserName
            rsData.UPDATE
            rsData.MoveNext
        Wend
    End If
  
    'コンテナボックスヤード設置トラン更新
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM CBYS_TRAN "
    strSQL = strSQL & " WHERE CBYST_BOXNO = '" & strNewBoxNo & "' "
'↓INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & "   AND CBYST_EDANO = '" & strNewEdaNo & "' "
'↑INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & " AND   CBYST_SYUBETSU = 0 "
    strSQL = strSQL & " AND   CBYST_YCODE = '" & strNewYCode & "'"
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If Not rsData.EOF Then
        rsData.Fields("CBYST_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBYST_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBYST_SYUBETSU") = 0
        rsData.Fields("CBYST_YCODE") = strNewYCode
        rsData.Fields("CBYST_STATUS") = 0
        rsData.Fields("CBYST_SETDATE") = Format(txt_設置日, "yyyymmdd")
        rsData.Fields("CBYST_SETFLOOR") = Me.cmb_設置階層.Column(1, Me.cmb_設置階層.ListIndex)
        rsData.Fields("CBYST_BIKON") = Me.txt_備考
        rsData.Fields("CBYST_UPDAD") = strDate
        rsData.Fields("CBYST_UPDAJ") = strTime
        rsData.Fields("CBYST_UPDPB") = PROG_ID
        rsData.Fields("CBYST_UPDUB") = Form_FKS800.gstrUserName
        rsData.UPDATE
    End If
    rsData.Close
    Set rsData = Nothing
  
    'コンテナボックス部屋割付トランデータ確認
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM CBRA_TRAN "
    strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & strNewBoxNo & "'"
'↓INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & "   AND CBRAT_EDANO = '" & strNewEdaNo & "'"
'↑INSERT 2013/11/18 MIYAMOTO
    strSQL = strSQL & " AND   CBRAT_YCODE = " & strNewYCode
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    If Not rsData.EOF Then
        '追加用レコードセット
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBRA_RECD "
        strSQL = strSQL & " WHERE CBRAR_BOXNO = '" & strNewBoxNo & "'"
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & "   AND CBRAR_EDANO = '" & strNewEdaNo & "'"
'↑INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBRAR_YCODE =  " & strNewYCode
        Set rsData2 = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        While Not rsData.EOF
            '割付が削除されたか確認
            strSQL = ""
            strSQL = strSQL & "SELECT * FROM " & P_ROOM_WORK_TABLE
            strSQL = strSQL & " WHERE CNO = " & rsData.Fields("CBRAT_CNO")
            strSQL = strSQL & " AND   CNO <> 0"
            Set rsDataLocal = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
            If rsDataLocal.EOF Then
'部屋変更の際、履歴は取らない（履歴は移動の場合のみ出力）
'                'コンテナボックス部屋割付履歴トラン追加
'                rsData2.AddNew
'                rsData2.Fields("CBRAR_BOXNO") = rsData.Fields("CBRAT_BOXNO")
'                rsData2.Fields("CBRAR_YCODE") = rsData.Fields("CBRAT_YCODE")
'                rsData2.Fields("CBRAR_CNO") = rsData.Fields("CBRAT_CNO")
'                rsData2.Fields("CBRAR_SETDATE") = rsData.Fields("CBRAT_SETDATE")
'                rsData2.Fields("CBRAR_INSED") = strDate
'                rsData2.Fields("CBRAR_INSEJ") = strTime
'                rsData2.Fields("CBRAR_INSPB") = PROG_ID
'                rsData2.Fields("CBRAR_INSUB") = Form_FKS800.gstrUserName
'                rsData2.UPDATE
                '旧コンテナボックス部屋割付トラン削除
                rsData.Delete
            End If
            rsDataLocal.Close
            Set rsDataLocal = Nothing
            rsData.MoveNext
        Wend
        rsData.Close
        Set rsData = Nothing
        rsData2.Close
        Set rsData2 = Nothing
    End If
  
    'コンテナボックス部屋割付トラン登録
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM " & P_ROOM_WORK_TABLE
    strSQL = strSQL & " WHERE CNO <> 0"
    Set rsDataLocal = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    '割付入力あり
    While Not rsDataLocal.EOF
        'コンテナボックス部屋割付トランデータ確認
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBRA_TRAN "
        strSQL = strSQL & " WHERE CBRAT_BOXNO = '" & strNewBoxNo & "'"
'↓INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBRAT_EDANO = '" & strNewEdaNo & "'"
'↑INSERT 2013/11/18 MIYAMOTO
        strSQL = strSQL & " AND   CBRAT_YCODE = " & strNewYCode
        strSQL = strSQL & " AND   CBRAT_CNO = " & rsDataLocal.Fields("CNO")
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        If rsData.EOF Then
            'データが無い時Insertする
            rsData.AddNew
            rsData.Fields("CBRAT_INSED") = strDate
            rsData.Fields("CBRAT_INSEJ") = strTime
            rsData.Fields("CBRAT_INSPB") = PROG_ID
            rsData.Fields("CBRAT_INSUB") = Form_FKS800.gstrUserName
        Else
            rsData.Fields("CBRAT_UPDAD") = strDate
            rsData.Fields("CBRAT_UPDAJ") = strTime
            rsData.Fields("CBRAT_UPDPB") = PROG_ID
            rsData.Fields("CBRAT_UPDUB") = Form_FKS800.gstrUserName
        End If
        rsData.Fields("CBRAT_BOXNO") = strNewBoxNo
'↓INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBRAT_EDANO") = strNewEdaNo
'↑INSERT 2013/11/18 MIYAMOTO
        rsData.Fields("CBRAT_YCODE") = strNewYCode
        rsData.Fields("CBRAT_CNO") = rsDataLocal.Fields("CNO")
        rsData.Fields("CBRAT_STATUS") = 0
        rsData.Fields("CBRAT_SETDATE") = Format(txt_設置日, "yyyymmdd")
        rsData.UPDATE
        rsData.Close
        Set rsData = Nothing
        rsDataLocal.MoveNext
    Wend
    rsDataLocal.Close
    Set rsDataLocal = Nothing
    
    '正常終了
    UPD_RTN = True

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    If Not rsData2 Is Nothing Then
        rsData2.Close
        Set rsData2 = Nothing
    End If
    If Not rsDataLocal Is Nothing Then
        rsDataLocal.Close
        Set rsDataLocal = Nothing
    End If
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "UPD_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      : ボックス番号自動発番処理
'        MODULE_ID        : fncGetAutoBoxNo
'        PARAMETER        : 引数なし
'        戻り値　　　　　 : 発番したボックス番号
'        CREATE_DATE      : 2008/06/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetAutoBoxNo() As String
  
    Dim objAdoDbConnection   As Object
    Dim strSQL               As String
    Dim rsData               As Object
    Dim rsDataChk            As Object
    Dim iAdd                 As Integer
    Dim bEnd                 As Boolean
'↓INSERT 2014/01/17 MIYAMOTO
    Dim strPrefix            As String
'↑INSERT 2014/01/17 MIYAMOTO

On Error GoTo Exception

'↓INSERT 2014/01/17 MIYAMOTO
    '接頭文字の取得
'↓UPDATE 2014/01/23 MIYAMOTO
'    strPrefix = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'FKS800' AND INTIF_RECFB = 'AUTO_PREFIX'"), "")
    strPrefix = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'FKS800' AND INTIF_RECFB = 'AUTO_PREFIX_" & Format(Me.cmb_設置用途.Column(1, Me.cmb_設置用途.ListIndex), "00") & "'"), "")
'↑UPDATE 2014/01/23 MIYAMOTO
'↑INSERT 2014/01/17 MIYAMOTO
    
    '初期値設定
    iAdd = 1
    bEnd = False
    fncGetAutoBoxNo = "1"

    'コンテナボックスマスタ件数＋ｎを初期ボックス番号とする
    strSQL = ""
    strSQL = strSQL & "SELECT COUNT(*) AS DATA_COUNT FROM CBOX_MAST "
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    While Not bEnd

'↓UPDATE 2014/01/17 MIYAMOTO
'        fncGetAutoBoxNo = Nz(rsData.Fields("DATA_COUNT")) + iAdd
        fncGetAutoBoxNo = strPrefix & (Nz(rsData.Fields("DATA_COUNT")) + iAdd)
'↑UPDATE 2014/01/17 MIYAMOTO
        '発番したボックス番号が存在するか確認
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM CBOX_MAST "
        strSQL = strSQL & " WHERE CBOXM_BOXNO = '" & fncGetAutoBoxNo & "'"
        Set rsDataChk = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        bEnd = rsDataChk.EOF
        rsDataChk.Close
        Set rsDataChk = Nothing
        iAdd = iAdd + 1
    Wend
    rsData.Close
    objAdoDbConnection.Close
    Set rsData = Nothing
    Set objAdoDbConnection = Nothing

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    If Not rsDataChk Is Nothing Then
        rsDataChk.Close
        Set rsDataChk = Nothing
    End If
    If Err <> 0 Then
        If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
        Call Err.Raise(Err.Number, "fncGetAutoBoxNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :一覧に戻る処理
'        MODULE_ID        :CloseWindow
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CloseWindow()

On Error GoTo Exception
    
    doCmd.Close acForm, Me.FormName
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "CloseWindow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォーム初期表示処理
'        MODULE_ID        :setDefaultWindow
'        NOTE             :
'        CREATE_DATE      :2008/06/20
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setDefaultWindow()

    Dim bMoveYard      As Boolean
    Dim intLoopNo      As Integer
'↓INSERT 2013/11/18 MIYAMOTO
    Dim objAdoDbConnection  As Object
    Dim strSQL              As String
    Dim rsData              As Object
'↑INSERT 2013/11/18 MIYAMOTO
    
On Error GoTo Exception
    
    pstrProcMode = pstrParameter(23)   '２３番目は処理モード
    
    Me.lbl_ヤードコード.Caption = pstrParameter(0)
    pstrOldYCode = pstrParameter(0)
    Me.lbl_ヤード名称.Caption = pstrParameter(1)
    Me.lbl_営業開始日.Caption = pstrParameter(2)
    Me.lbl_営業終了日.Caption = pstrParameter(3)
    Me.lbl_ヤード解約予定日.Caption = pstrParameter(4)
    If pstrParameter(5) = "" Then
'↓UPDATE 2014/01/23 MIYAMOTO
'        '自動発番
''↓UPDATE 2013/11/18 MIYAMOTO
''        Me.txt_ボックス番号.VALUE = fncGetAutoBoxNo
'        '前画面(FKS800)の「コンテナ番号自動発番可否」がチェックオンの場合は自動発番、チェックオフの場合は空欄とする
'        If Form_FKS800.gstrAutoBoxNoKaHi = -1 Then
'            Me.txt_ボックス番号.VALUE = fncGetAutoBoxNo
'        Else
'            Me.txt_ボックス番号.VALUE = ""
'        End If
''↑UPDATE 2013/11/18 MIYAMOTO
        Me.txt_ボックス番号.VALUE = ""
'↑UPDATE 2014/01/23 MIYAMOTO
'↓INSERT 2013/11/18 MIYAMOTO
        Me.txt_枝番.VALUE = "00"
'↑INSERT 2013/11/18 MIYAMOTO

    Else
'↓UPDATE 2013/11/18 MIYAMOTO
'        Me.txt_ボックス番号.VALUE = pstrParameter(5)
        Me.txt_ボックス番号.VALUE = Mid(pstrParameter(5), 1, InStr(pstrParameter(5), " ") - 1)
        Me.txt_枝番.VALUE = Mid(pstrParameter(5), InStr(pstrParameter(5), " ") + 1)
'↑UPDATE 2013/11/18 MIYAMOTO
    End If
    pstrOldBoxNo = Me.txt_ボックス番号.VALUE
'↓INSERT 2013/11/18 MIYAMOTO
    pstrOldEdaNo = Me.txt_枝番.VALUE
    If pstrProcMode = "0" Then
        '新規時、パラメータより取得
'↑INSERT 2013/11/18 MIYAMOTO
        pstrType = pstrParameter(6)
        Me.lbl_タイプ.Caption = pstrParameter(7)
        pstrInDesign = pstrParameter(8)
        Me.lbl_内装.Caption = pstrParameter(9)
        pstrSize = pstrParameter(10)
        Me.lbl_サイズ.Caption = pstrParameter(11)
        pstrDoor = pstrParameter(12)
        If pstrParameter(13) = "" Then
            Me.lbl_ドア数.Caption = ""
        Else
            Me.lbl_ドア数.Caption = Split(pstrParameter(13), "#")(0)
        End If
        pstrDoorIti = pstrParameter(14)
        Me.lbl_ドア位置.Caption = pstrParameter(15)
        pstrDoorType = pstrParameter(16)
        Me.lbl_ドアタイプ.Caption = pstrParameter(17)
        Me.txt_入庫日.VALUE = IIf(pstrParameter(18) = "", Format(DATE, "YYYY/MM/DD"), pstrParameter(18))
        Me.txt_設置日.VALUE = IIf((pstrParameter(19) = "") Or (pstrProcMode = "2"), _
                                  Format(DATE, "YYYY/MM/DD"), _
                                  pstrParameter(19))
        pstrSetDate = pstrParameter(19)
        pstrFloorLevel = pstrParameter(20)
        pstrOldFloorLevel = pstrParameter(20)
        Me.txt_備考.VALUE = pstrParameter(21)
'↓INSERT 2013/11/18 MIYAMOTO
        pstrKanriKbn = ""               '所有
        Me.lbl_所有.Caption = ""
        pstrKounyuKbn = ""              '購入区分
        Me.lbl_購入区分.Caption = ""
        pstrCtBunrui = ""               'コンテナ分類
        Me.lbl_コンテナ分類.Caption = ""
        pstrCtYouto = ""                '設置用途
        Me.lbl_設置用途.Caption = ""
'当バージョン(0.1)では非表示とする
'        pstrKenKakunin = ""             '建築確認
'        Me.lbl_建築確認.Caption = ""
        pstrSySyubetu = ""              'コンテナ内通路
        Me.lbl_コンテナ内通路.Caption = ""
        pstrBunkatuSu = "0"             '分割数
        Me.txt_分割数.VALUE = pstrBunkatuSu
    Else
        '修正、移動時、コンテナボックスマスタより該当データを取得
        strSQL = ""
        strSQL = strSQL & " select  "
        strSQL = strSQL & "      CBOX_MAST.CBOXM_BOXNO as BOXNO "               'ボックス番号
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_EDANO as EDANO "               '枝番
        strSQL = strSQL & "     ,NAME_MAST1.NAME_RYAK as TYPE "                 'タイプ(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_TYPE as TYPE_VARUE "           'タイプ
        strSQL = strSQL & "     ,NAME_MAST2.NAME_RYAK as INTYPE "               '内装(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_INTYPE as INTYPE_VARUE "       '内装
        strSQL = strSQL & "     ,NAME_MAST3.NAME_RYAK as SIZE "                 'サイズ(名称)
        strSQL = strSQL & "     ,NAME_MAST3.NAME_VALUE_FROM as SIZE_VARUE "     'サイズ
        strSQL = strSQL & "     ,NAME_MAST4.NAME_RYAK as DOOR "                 'ドア数(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOOR as DOOR_VARUE "           'ドア数
        strSQL = strSQL & "     ,NAME_MAST5.NAME_RYAK as DOORPSTN "             'ドア位置(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORPSTN as DOORPSTN_VARUE "   'ドア位置
        strSQL = strSQL & "     ,NAME_MAST6.NAME_RYAK as DOORTYPE "             'ドアタイプ(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_DOORTYPE as DOORTYPE_VARUE "   'ドアタイプ
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STOCKDATE as STOCKDATE "       '入庫日
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_STATUS as STATUS "             '状態
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_BKTSU as BKTSU "               '分割数
        strSQL = strSQL & "     ,NAME_MAST8.NAME_RYAK as KNKBN "                '購入区分(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKBN as KNKBN_VALUE "         '購入区分
        strSQL = strSQL & "     ,NAME_MAST9.NAME_RYAK as CTBNR "                'コンテナ分類(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTBNR as CTBNR_VALUE "         'コンテナ分類
        strSQL = strSQL & "     ,NAME_MAST10.NAME_RYAK as CTYTO "               '設置用途(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_CTYTO as CTYTO_VALUE "         '設置用途
        strSQL = strSQL & "     ,NAME_MAST11.NAME_RYAK as KNKAK "               '建築確認(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KNKAK as KNKAK_VALUE "         '建築確認
        strSQL = strSQL & "     ,NAME_MAST12.NAME_RYAK as SHSHU "               'コンテナ内通路(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_SHSHU as SHSHU_VALUE "         'コンテナ内通路
        strSQL = strSQL & "     ,NAME_MAST13.NAME_RYAK as KRKBN "               '所有(名称)
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_KRKBN as KRKBN_VALUE "         '所有
        strSQL = strSQL & "     ,CBOX_MAST.CBOXM_OLDBN as OLDBN "               '旧ボックス番号
        strSQL = strSQL & "     ,w01.CBRAT_CNO as CNO "                         '
        strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETDATE as SETDATE "           '設置日
        strSQL = strSQL & "     ,NAME_MAST7.NAME_RYAK as SETFLOOR "             '設置階層(名称)
        strSQL = strSQL & "     ,CBYS_TRAN.CBYST_SETFLOOR as SETFLOOR_VARUE "   '設置階層
        strSQL = strSQL & "     ,CBYS_TRAN.CBYST_BIKON as BIKON "               '備考
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
        strSQL = strSQL & " where CBYS_TRAN.CBYST_BOXNO = CBOX_MAST.CBOXM_BOXNO "
        strSQL = strSQL & "   and CBYS_TRAN.CBYST_EDANO = CBOX_MAST.CBOXM_EDANO "
        strSQL = strSQL & "   and CBYS_TRAN.CBYST_STATUS = 0 "
        strSQL = strSQL & "   and CBYS_TRAN.CBYST_SYUBETSU = 0 "
        strSQL = strSQL & "   and CBOX_MAST.CBOXM_BOXNO = '" & pstrOldBoxNo & "' "
        strSQL = strSQL & "   and CBOX_MAST.CBOXM_EDANO = '" & pstrOldEdaNo & "' "
        strSQL = strSQL & "   and REPLACE (CBYS_TRAN.CBYST_YCODE,' ','') = REPLACE ('" & pstrOldYCode & "',' ','') "
        strSQL = strSQL & " order by BOXNO,SETDATE,CNO "
        Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
        If Not rsData.EOF Then
            pstrType = Nz(rsData.Fields("TYPE_VARUE").VALUE)
            Me.lbl_タイプ.Caption = Nz(rsData.Fields("TYPE").VALUE)
            pstrInDesign = Nz(rsData.Fields("INTYPE_VARUE").VALUE)
            Me.lbl_内装.Caption = Nz(rsData.Fields("INTYPE").VALUE)
            pstrSize = Nz(rsData.Fields("SIZE_VARUE").VALUE)
            Me.lbl_サイズ.Caption = Nz(rsData.Fields("SIZE").VALUE)
            pstrDoor = Nz(rsData.Fields("DOOR_VARUE").VALUE)
            If Nz(rsData.Fields("DOOR").VALUE) = "" Then
                Me.lbl_ドア数.Caption = ""
            Else
                Me.lbl_ドア数.Caption = Nz(rsData.Fields("DOOR").VALUE)
            End If
            pstrDoorIti = Nz(rsData.Fields("DOORPSTN_VARUE").VALUE)
            Me.lbl_ドア位置.Caption = Nz(rsData.Fields("DOORPSTN").VALUE)
            pstrDoorType = Nz(rsData.Fields("DOORTYPE_VARUE").VALUE)
            Me.lbl_ドアタイプ.Caption = Nz(rsData.Fields("DOORTYPE").VALUE)
            Me.txt_入庫日.VALUE = IIf(Nz(rsData.Fields("STOCKDATE").VALUE) = "", Format(DATE, "YYYY/MM/DD"), Format(Nz(rsData.Fields("STOCKDATE").VALUE), "@@@@\/@@\/@@"))
            Me.txt_設置日.VALUE = IIf((Nz(rsData.Fields("SETDATE").VALUE) = "") Or (pstrProcMode = "2"), _
                                      Format(DATE, "YYYY/MM/DD"), _
                                      Format(Nz(rsData.Fields("SETDATE").VALUE), "@@@@\/@@\/@@"))
            pstrSetDate = Format(Nz(rsData.Fields("SETDATE").VALUE), "@@@@\/@@\/@@")
            pstrFloorLevel = Nz(rsData.Fields("SETFLOOR_VARUE").VALUE)
            pstrOldFloorLevel = Nz(rsData.Fields("SETFLOOR_VARUE").VALUE)
            If Nz(rsData.Fields("BIKON").VALUE) = "" Then
                Me.txt_備考.VALUE = ""
            Else
                Me.txt_備考.VALUE = Nz(rsData.Fields("BIKON").VALUE)
            End If
            pstrKounyuKbn = Nz(rsData.Fields("KNKBN_VALUE").VALUE)
            Me.lbl_購入区分.Caption = Nz(rsData.Fields("KNKBN").VALUE)
            pstrCtBunrui = Nz(rsData.Fields("CTBNR_VALUE").VALUE)
            Me.lbl_コンテナ分類.Caption = Nz(rsData.Fields("CTBNR").VALUE)
            pstrCtYouto = Nz(rsData.Fields("CTYTO_VALUE").VALUE)
            Me.lbl_設置用途.Caption = Nz(rsData.Fields("CTYTO").VALUE)
'当バージョン(0.1)では非表示とする
'            pstrKenKakunin = Nz(rsData.Fields("KNKAK_VALUE").VALUE)
'            Me.lbl_建築確認.Caption = Nz(rsData.Fields("KNKAK").VALUE)
            pstrSySyubetu = Nz(rsData.Fields("SHSHU_VALUE").VALUE)
            Me.lbl_コンテナ内通路.Caption = Nz(rsData.Fields("SHSHU").VALUE)
            pstrKanriKbn = Nz(rsData.Fields("KRKBN_VALUE").VALUE)
            Me.lbl_所有.Caption = Nz(rsData.Fields("KRKBN").VALUE)
            pstrBunkatuSu = Nz(rsData.Fields("BKTSU").VALUE)
            Me.txt_分割数.VALUE = pstrBunkatuSu
        End If
    End If
'↑INSERT 2013/11/18 MIYAMOTO
    Me.lbl_タイトル.Caption = pstrParameter(22)
    bMoveYard = (pstrProcMode = "2")
    
    Me.cmb_ヤード.Visible = bMoveYard
    Me.lbl_タイプ.Visible = bMoveYard
    Me.lbl_サイズ.Visible = bMoveYard
    Me.lbl_ドア数.Visible = bMoveYard
    Me.lbl_内装.Visible = bMoveYard
    Me.lbl_ドアタイプ.Visible = bMoveYard
    Me.lbl_ドア位置.Visible = bMoveYard
    Me.lbl_ヤードコード.Visible = (Not bMoveYard)
    Me.cmb_タイプ.Visible = (Not bMoveYard)
    Me.cmb_サイズ.Visible = (Not bMoveYard)
    Me.cmb_ドア数.Visible = (Not bMoveYard)
    Me.cmb_内装.Visible = (Not bMoveYard)
    Me.cmb_ドアタイプ.Visible = (Not bMoveYard)
    Me.cmb_ドア位置.Visible = (Not bMoveYard)
    Me.txt_ボックス番号.Enabled = (Not bMoveYard)
    Me.txt_ボックス番号.Locked = bMoveYard
    Me.txt_入庫日.Enabled = (Not bMoveYard)
    Me.txt_入庫日.Locked = bMoveYard
'↓INSERT 2013/11/18 MIYAMOTO
    Me.lbl_所有.Visible = bMoveYard
    Me.lbl_購入区分.Visible = bMoveYard
    Me.lbl_コンテナ分類.Visible = bMoveYard
    Me.lbl_設置用途.Visible = bMoveYard
'当バージョン(0.1)では非表示とする
'    Me.lbl_建築確認.Visible = bMoveYard
    Me.lbl_コンテナ内通路.Visible = bMoveYard
    Me.cmb_所有.Visible = (Not bMoveYard)
    Me.cmb_購入区分.Visible = (Not bMoveYard)
    Me.cmb_コンテナ分類.Visible = (Not bMoveYard)
    Me.cmb_設置用途.Visible = (Not bMoveYard)
'当バージョン(0.1)では非表示とする
'    Me.cmb_建築確認.Visible = (Not bMoveYard)
    Me.cmb_コンテナ内通路.Visible = (Not bMoveYard)
    Me.txt_枝番.Enabled = (Not bMoveYard)
    Me.txt_枝番.Locked = bMoveYard
    Me.txt_分割数.Enabled = (Not bMoveYard)
    Me.txt_分割数.Locked = bMoveYard
'↑INSERT 2013/11/18 MIYAMOTO
'↓INSERT 2018/07/17 PARK
    Me.txt_複写件数.VALUE = "1"
    Me.txt_複写件数.Enabled = (Not (pstrProcMode = "1"))
    Me.txt_複写件数.Locked = (pstrProcMode = "1")
    Me.txt_複写件数.Enabled = (Not bMoveYard)
    Me.txt_複写件数.Locked = bMoveYard
    Me.txt_複写件数.Enabled = (Not Form_FKS800.gstrAutoBoxNoKaHi <> -1)
    Me.txt_複写件数.Locked = Form_FKS800.gstrAutoBoxNoKaHi <> -1
'↑INSERT 2018/07/17 PARK
    If (pstrProcMode = "0") Then
        '新規登録
        Me.cmb_タイプ.SetFocus
        Me.cmb_タイプ.Text = ""
        Me.cmb_内装.SetFocus
        Me.cmb_内装.ListIndex = 1
        Me.cmb_サイズ.SetFocus
        Me.cmb_サイズ.Text = ""
        Me.cmb_ドア数.SetFocus
        Me.cmb_ドア数.Text = ""
        Me.cmb_ドアタイプ.SetFocus
        Me.cmb_ドアタイプ.ListIndex = 0
        Me.cmb_ドア位置.SetFocus
        Me.cmb_ドア位置.ListIndex = 0
        Me.cmb_設置階層.SetFocus
        Me.cmb_設置階層.Text = ""

'↓INSERT 2013/11/18 MIYAMOTO
        Me.cmb_所有.SetFocus
        Me.cmb_所有.Text = ""
        Me.cmb_購入区分.SetFocus
        Me.cmb_購入区分.Text = ""
        Me.cmb_コンテナ分類.SetFocus
        Me.cmb_コンテナ分類.Text = ""
        Me.cmb_設置用途.SetFocus
        Me.cmb_設置用途.Text = ""
'当バージョン(0.1)では非表示とする
'        Me.cmb_建築確認.SetFocus
'        Me.cmb_建築確認.text = ""
        Me.cmb_コンテナ内通路.SetFocus
        Me.cmb_コンテナ内通路.Text = ""
'↑INSERT 2013/11/18 MIYAMOTO

        Me.txt_ボックス番号.BackColor = 16777215
        Me.txt_入庫日.BackColor = 16777215
'↓INSERT 2013/11/18 MIYAMOTO
        Me.txt_枝番.BackColor = 16777215
        Me.txt_分割数.BackColor = 16777215
'↑INSERT 2013/11/18 MIYAMOTO
        Me.pge_ヤード.Caption = "ヤード"
        Me.txt_ボックス番号.SetFocus

    ElseIf (pstrProcMode = "1") Then
        '修正
        For intLoopNo = 0 To Me.cmb_タイプ.ListCount - 1
            If Me.cmb_タイプ.Column(1, intLoopNo) = pstrType Then
                Me.cmb_タイプ.SetFocus
                Me.cmb_タイプ.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_内装.ListCount - 1
            If Me.cmb_内装.Column(1, intLoopNo) = pstrInDesign Then
                Me.cmb_内装.SetFocus
                Me.cmb_内装.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_サイズ.ListCount - 1
            If Me.cmb_サイズ.Column(1, intLoopNo) = pstrSize Then
                Me.cmb_サイズ.SetFocus
                Me.cmb_サイズ.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_ドアタイプ.ListCount - 1
            If Me.cmb_ドアタイプ.Column(1, intLoopNo) = pstrDoorType Then
                Me.cmb_ドアタイプ.SetFocus
                Me.cmb_ドアタイプ.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_ドア位置.ListCount - 1
            If Me.cmb_ドア位置.Column(1, intLoopNo) = pstrDoorIti Then
                Me.cmb_ドア位置.SetFocus
                Me.cmb_ドア位置.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_設置階層.ListCount - 1
            If Me.cmb_設置階層.Column(1, intLoopNo) = pstrFloorLevel Then
                Me.cmb_設置階層.SetFocus
                Me.cmb_設置階層.ListIndex = intLoopNo
                Exit For
            End If
        Next
'↓INSERT 2013/11/18 MIYAMOTO
        For intLoopNo = 0 To Me.cmb_所有.ListCount - 1
            If Me.cmb_所有.Column(1, intLoopNo) = pstrKanriKbn Then
                Me.cmb_所有.SetFocus
                Me.cmb_所有.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_購入区分.ListCount - 1
            If Me.cmb_購入区分.Column(1, intLoopNo) = pstrKounyuKbn Then
                Me.cmb_購入区分.SetFocus
                Me.cmb_購入区分.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_コンテナ分類.ListCount - 1
            If Me.cmb_コンテナ分類.Column(1, intLoopNo) = pstrCtBunrui Then
                Me.cmb_コンテナ分類.SetFocus
                Me.cmb_コンテナ分類.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_設置用途.ListCount - 1
            If Me.cmb_設置用途.Column(1, intLoopNo) = pstrCtYouto Then
                Me.cmb_設置用途.SetFocus
                Me.cmb_設置用途.ListIndex = intLoopNo
                Exit For
            End If
        Next
'当バージョン(0.1)では非表示とする
'        For intLoopNo = 0 To Me.cmb_建築確認.ListCount - 1
'            If Me.cmb_建築確認.Column(1, intLoopNo) = pstrKenKakunin Then
'                Me.cmb_建築確認.SetFocus
'                Me.cmb_建築確認.ListIndex = intLoopNo
'                Exit For
'            End If
'        Next
        For intLoopNo = 0 To Me.cmb_コンテナ内通路.ListCount - 1
            If Me.cmb_コンテナ内通路.Column(1, intLoopNo) = pstrSySyubetu Then
                Me.cmb_コンテナ内通路.SetFocus
                Me.cmb_コンテナ内通路.ListIndex = intLoopNo
                Exit For
            End If
        Next
'↑INSERT 2013/11/18 MIYAMOTO
        For intLoopNo = 0 To Me.cmb_ドア数.ListCount - 1
            If Me.cmb_ドア数.Column(1, intLoopNo) = pstrDoor Then
                Me.cmb_ドア数.SetFocus
                Me.cmb_ドア数.ListIndex = intLoopNo
                Exit For
            End If
        Next

        Me.txt_ボックス番号.BackColor = 16777215
        Me.txt_入庫日.BackColor = 16777215
'↓INSERT 2013/11/18 MIYAMOTO
        Me.txt_枝番.BackColor = 16777215
        Me.txt_分割数.BackColor = 16777215
'↑INSERT 2013/11/18 MIYAMOTO
        Me.pge_ヤード.Caption = "ヤード"
        Me.txt_ボックス番号.SetFocus
    ElseIf (pstrProcMode = "2") Then
        '移動登録
        For intLoopNo = 0 To Me.cmb_設置階層.ListCount - 1
            If Me.cmb_設置階層.Column(1, intLoopNo) = pstrFloorLevel Then
                Me.cmb_設置階層.SetFocus
                Me.cmb_設置階層.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_ヤード.ListCount - 1
            If Me.cmb_ヤード.Column(0, intLoopNo) = pstrParameter(0) Then
                Me.cmb_ヤード.SetFocus
                Me.cmb_ヤード.ListIndex = intLoopNo
                Exit For
            End If
        Next

        Me.txt_ボックス番号.BackColor = 14680063
        Me.txt_入庫日.BackColor = 14680063
'↓INSERT 2013/11/18 MIYAMOTO
        Me.txt_枝番.BackColor = 14680063
        Me.txt_分割数.BackColor = 14680063
'↑INSERT 2013/11/18 MIYAMOTO
        Me.pge_ヤード.Caption = "ヤード移動指定"
        Me.cmd_保存.Caption = "移動"
    
    '2019/04/17 M.HONDA INS
    ElseIf (pstrProcMode = "3") Then
        '複写
        For intLoopNo = 0 To Me.cmb_タイプ.ListCount - 1
            If Me.cmb_タイプ.Column(1, intLoopNo) = pstrType Then
                Me.cmb_タイプ.SetFocus
                Me.cmb_タイプ.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_内装.ListCount - 1
            If Me.cmb_内装.Column(1, intLoopNo) = pstrInDesign Then
                Me.cmb_内装.SetFocus
                Me.cmb_内装.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_サイズ.ListCount - 1
            If Me.cmb_サイズ.Column(1, intLoopNo) = pstrSize Then
                Me.cmb_サイズ.SetFocus
                Me.cmb_サイズ.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_ドアタイプ.ListCount - 1
            If Me.cmb_ドアタイプ.Column(1, intLoopNo) = pstrDoorType Then
                Me.cmb_ドアタイプ.SetFocus
                Me.cmb_ドアタイプ.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_ドア位置.ListCount - 1
            If Me.cmb_ドア位置.Column(1, intLoopNo) = pstrDoorIti Then
                Me.cmb_ドア位置.SetFocus
                Me.cmb_ドア位置.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_設置階層.ListCount - 1
            If Me.cmb_設置階層.Column(1, intLoopNo) = pstrFloorLevel Then
                Me.cmb_設置階層.SetFocus
                Me.cmb_設置階層.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_所有.ListCount - 1
            If Me.cmb_所有.Column(1, intLoopNo) = pstrKanriKbn Then
                Me.cmb_所有.SetFocus
                Me.cmb_所有.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_購入区分.ListCount - 1
            If Me.cmb_購入区分.Column(1, intLoopNo) = pstrKounyuKbn Then
                Me.cmb_購入区分.SetFocus
                Me.cmb_購入区分.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_コンテナ分類.ListCount - 1
            If Me.cmb_コンテナ分類.Column(1, intLoopNo) = pstrCtBunrui Then
                Me.cmb_コンテナ分類.SetFocus
                Me.cmb_コンテナ分類.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_設置用途.ListCount - 1
            If Me.cmb_設置用途.Column(1, intLoopNo) = pstrCtYouto Then
                Me.cmb_設置用途.SetFocus
                Me.cmb_設置用途.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_コンテナ内通路.ListCount - 1
            If Me.cmb_コンテナ内通路.Column(1, intLoopNo) = pstrSySyubetu Then
                Me.cmb_コンテナ内通路.SetFocus
                Me.cmb_コンテナ内通路.ListIndex = intLoopNo
                Exit For
            End If
        Next
        For intLoopNo = 0 To Me.cmb_ドア数.ListCount - 1
            If Me.cmb_ドア数.Column(1, intLoopNo) = pstrDoor Then
                Me.cmb_ドア数.SetFocus
                Me.cmb_ドア数.ListIndex = intLoopNo
                Exit For
            End If
        Next

        Me.txt_ボックス番号.BackColor = 16777215
        Me.txt_入庫日.BackColor = 16777215
        Me.txt_枝番.BackColor = 16777215
        Me.txt_分割数.BackColor = 16777215
        Me.pge_ヤード.Caption = "ヤード"
        Me.txt_ボックス番号.SetFocus
        Me.txt_ボックス番号.VALUE = Left(Mid(pstrParameter(5), 1, InStr(pstrParameter(5), " ") - 1), 4)
        Me.txt_枝番.VALUE = "00"

'2019/04/17 M.HONDA INS
        
        
    End If
    
'↓INSERT 2018/07/17 PARK
    If (pstrProcMode = "1") Or (pstrProcMode = "2") Or (Form_FKS800.gstrAutoBoxNoKaHi <> -1) Then
        Me.txt_複写件数.BackColor = 14680063
        Me.txt_複写件数.Enabled = False
        Me.txt_複写件数.Locked = True
    End If
'↑INSERT 2018/07/17 PARK
    
    Form_FKS810.sub_部屋割付.Form.RecordSource = "SELECT * FROM " & P_ROOM_WORK_TABLE & " AS W01;"
    Form_FKS810.sub_部屋割付.Form.Requery
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "setDefaultWindow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub



