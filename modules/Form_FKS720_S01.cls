VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS720_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード別割引一覧（サブ）
'   プログラムＩＤ　：　FKS720_S01
'   作　成　日　　　：  2008/07/17
'   作　成　者　　　：  N.IMAI
'   Ver            ：  0.0
'**********************************************
'   修　正　日　　　：  2008/10/25
'   修　正　者　　　：  K.KINEBUCHI
'   内　　　容　　　：  優先順位のコンボボックスに、小数点を入力した際
'   　　　　　　　　　　エラーメッセージを出力するように修正した。
'   Ver            ：  0.1
'**********************************************
'   修　正　日　　　：  2008/11/13
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  次画面へ文言を引き数で渡す際に
'   　　　　　　　　　　区切り文字がカンマの場合、不具合があるので
'   　　　　　　　　　　区切り文字をタブに変更
'   Ver            ：  0.2
'**********************************************
'   修　正　日　　　：  2008/12/22
'   修　正　者　　　：  K.KINEBUCHI
'   内　　　容　　　：  終了した割引を削除する際の不具合を修正する。
'   Ver            ：  0.1（これより改修分のVerは、0.1とする。以降+0.1ずづアップする）
'**********************************************
Option Compare Database
Option Explicit

Private Const P_KAKUNIN = "確認"
Private Const P_ERROR = "エラー"
Private Const P_TITLE_UPDATE = "（修正中）"
Private Const P_INIT_ROWSOURCE = "1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30;31;32;33;34;35;36;37;38;39;40;41;42;43;44;45;46;47;48;49;50;51;52;53;54;55;56;57;58;59;60;61;62;63;64;65;66;67;68;69;70;71;72;73;74;75;76;77;78;79;80;81;82;83;84;85;86;87;88;89;90;91;92;93;94;95;96;97;98;99"
Private Const P_NEBIKI_LIST_WORK_TABLE = "WK_FKS720"

' 桁数
Private Const P_YUUSEN_LENGTH        As Integer = 2 '区分

'コントロールの種類
Private Const P_CONTROL_数値         As Integer = 1 '数値

Private Const PROG_ID = "FKS70_S01"

Private pobjDB                       As Object

Private pstrBumonCode                As String      '部門コード

'==============================================================================*
'
'        MODULE_NAME      :cmb_優先_AfterUpdate
'        MODULE_ID        :cmb_優先_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_優先_AfterUpdate()

    Call Me.Parent.subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_優先_BeforeUpdate
'        MODULE_ID        :cmb_優先_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_優先_BeforeUpdate(Cancel As Integer)

    Dim vCode   As Variant
    Dim vRet    As Variant                                  'INSERT 2008/10/25 K.KINEBUCHI

    On Error GoTo Exception
    
    vCode = Nz(Me.cmb_優先, "")
    
    '入力チェック
    If Not IsNumeric(vCode) Then
        '数値以外の入力はｴﾗｰ
        MsgBox "数値を入力して下さい。", vbInformation, P_ERROR
        Cancel = True
        Exit Sub
    End If
    If CmLenb(vCode) = 0 Or CmLenb(vCode) > 2 Then
        '桁数ｴﾗｰ
        MsgBox "桁数が正しくありません。", vbInformation, P_ERROR
        Cancel = True
        Exit Sub
    End If

    'INSERT 2008/10/25 K.KINEBUCHI start
    '小数点が入力されたかチェック
    vRet = Nz(InStr(Abs(Me.cmb_優先), "."), 0)
        
    If vRet <> 0 Then
        '小数点の入力はエラー
        MsgBox "小数点は入力できません。", vbInformation, P_ERROR
        Cancel = True
        Exit Sub
    End If
    'INSERT 2008/10/25 K.KINEBUCHI end
    
Exception:

    If Err <> 0 Then
        Call MSZZ024_M00("cmb_優先_BeforeUpdate", True)
        Err.Clear
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_優先_Enter
'        MODULE_ID        :cmb_優先_Enter
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_優先_Enter()

    On Error Resume Next

    'エラーは無視する
    If cmb_優先 = "0" Then
        cmb_優先.RowSource = "0"
    Else
        '選択できる値を設定
        cmb_優先.RowSource = P_INIT_ROWSOURCE
    End If

    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_削除_Click
'        MODULE_ID        :cmd_削除_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_Click()
    
On Error GoTo Exception
    
    If txt_契約数 <> "0" Then
        '契約あり
        Call MsgBox("この割引は適用した受付または契約があるため削除できません。", vbExclamation + vbOKOnly, P_ERROR)
    ElseIf InStr(1, Me.Parent.Caption, P_TITLE_UPDATE) > 0 Then
        '修正中
        Call MsgBox("修正中のため削除できません。" & vbCrLf & "一度保存するか再度読込みし直して下さい。", vbQuestion + vbOK, P_KAKUNIN)
    Else
        '削除処理
        If MsgBox("データを削除します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, P_KAKUNIN) = vbYes Then
            If DEL_RTN Then
                '一覧再表示（実施中割引）
                Call Parent.subCreateListData_FKS720
            End If
        End If
    End If

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmd_削除_Click", True)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_修正_Click
'        MODULE_ID        :cmd_修正_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_修正_Click()
    
    Dim strParameter As String
    
    On Error GoTo Exception
    
    If Parent.fncUpdateCheck(1) Then
        ' (0)タイトル - (1)モード       - (2)ヤードコード - (3)部門コード - (4)割引番号
        ' (5)有効可否 - (6)実施期間From - (7)実施期間To   - (8)提示文言   - (9)自動文言
        ' (9)備考
'D:2008.11.12 N.Imai Start
'        strParameter = "修正" _
'                     & ",1" _
'                     & "," & Me.Parent.cmb_ヤード _
'                     & "," & pstrBumonCode _
'                     & "," & Nz(txt_割引番号, "") _
'                     & "," & Nz(txt_有効可否, "") _
'                     & "," & Nz(txt_実施期間From, "") _
'                     & "," & Nz(txt_実施期間To, "") _
'                     & "," & Nz(txt_内容, "") _
'                     & "," & Nz(txt_自動文言, "") _
'                     & "," & Nz(txt_備考, "")
'D:2008.11.12 N.Imai End
'I:2008.11.12 N.Imai Start
        strParameter = "修正" _
                     & vbTab & "1" _
                     & vbTab & Me.Parent.cmb_ヤード _
                     & vbTab & pstrBumonCode _
                     & vbTab & Nz(txt_割引番号, "") _
                     & vbTab & Nz(txt_有効可否, "") _
                     & vbTab & Nz(txt_実施期間From, "") _
                     & vbTab & Nz(txt_実施期間To, "") _
                     & vbTab & Nz(txt_内容, "") _
                     & vbTab & Nz(txt_自動文言, "") _
                     & vbTab & Nz(txt_備考, "")
'I:2008.11.12 N.Imai Start
        doCmd.OpenForm "FKS721", , , , , acDialog, strParameter
        Me.Parent.subCreateListData_FKS720
    End If
    
Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("cmd_修正_Click", True)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Current
'        MODULE_ID        :Form_Current
'        NOTE             :
'        CREATE_DATE      :2008/07/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Current()

    cmb_優先.Visible = Me.Parent.Form.cmd_割引追加.Enabled
    txt_優先.Visible = Not Me.Parent.Form.cmd_割引追加.Enabled
'    cmd_修正.Enabled = Me.Parent.Form.cmd_割引追加.Enabled
'    cmd_削除.Enabled = Me.Parent.Form.cmd_割引追加.Enabled

End Sub

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
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)
    
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
'        MODULE_NAME      : 削除処理
'        MODULE_ID        : DEL_RTN
'        PARAMETER        : 引数なし
'        戻り値           : True=正常 False=異常
'        CREATE_DATE      : 2008/**/**
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DEL_RTN() As Boolean
  
    Dim objAdoDbConnection   As Object
    Dim strSQL               As String
    Dim rsDataDel            As Object

On Error GoTo Exception

    DEL_RTN = False
    
    'トランザクション開始
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    objAdoDbConnection.BeginTrans

    'レンタル物件割引削除
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM RLDN_TRAN"
    strSQL = strSQL & " WHERE RLDN_TRAN.RLDNT_BUMOC  =  '" & pstrBumonCode & "'"
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_YCODE  =  '" & Me.Parent.Form.cmb_ヤード & "'"
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_NO     =  '" & txt_割引番号 & "'"
'    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE =  '1'"                           'DELETE 2008/12/22 K.KINEBUCHI
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE =  '" & Form_FKS720.pintMode & "'" 'INSERT 2008/12/22 K.KINEBUCHI
    Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    rsDataDel.Delete
    rsDataDel.Close
    Set rsDataDel = Nothing

    DEL_RTN = True
    
    'コミット
    objAdoDbConnection.CommitTrans
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

Exception:

    If Not rsDataDel Is Nothing Then
        rsDataDel.Close
        Set rsDataDel = Nothing
    End If

    If Err <> 0 Then
        If Not objAdoDbConnection Is Nothing Then
            'ロールバック
            objAdoDbConnection.RollbackTrans
            objAdoDbConnection.Close
            Set objAdoDbConnection = Nothing
        End If
        Call Err.Raise(Err.Number, "DEL_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
