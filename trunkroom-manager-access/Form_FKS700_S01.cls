VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS700_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　割引マスタ一覧（サブ）
'   プログラムＩＤ　：　FKS700_S01
'   作　成　日　　　：  2008/07/25
'   作　成　者　　　：  T.Suzuki
'   Ver            ：  0.0
'**********************************************
'   修　正　日　　　：  2008/11/13
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  次画面へ文言を引き数で渡す際に
'   　　　　　　　　　　区切り文字がカンマの場合、不具合があるので
'   　　　　　　　　　　区切り文字をタブに変更
'   Ver            ：  0.1
'**********************************************
'   修　正　日　　　：  2009/06/30
'   修　正　者　　　：  hirano
'   内　　　容　　　：  割引マスタ KASE_DBへ移行に伴う変更
'   Ver            ：  0.3
'**********************************************
'   修　正　日　　　：  2013/04/08
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  画面の並び順を変更
'   Ver            ：  0.4
'**********************************************
'   修　正　日　　　：  2017/12/01
'   修　正　者　　　：  イーグルソフト
'   内　　　容　　　：  キャンペーン一括設定対応
'   Ver            ：  0.6
'**********************************************

Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS700_S01"
Private Const P_ROOM_WORK_TABLE         As String = "WK_FKS700"
Private Const P_NEBIKI_LIST_WORK_TABLE  As String = "WK_FKS700_S01_TEMP"

Private Const P_KAKUNIN                 As String = "確認"
Private Const P_ERROR                   As String = "エラー"

Private pobjDB                          As Object

Private pstrBumonCode                   As String       ' 部門コード

'==============================================================================*
'
'        MODULE_NAME      :cmd_削除_Click
'        MODULE_ID        :cmd_削除_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_Click()

    Dim isException        As Boolean

    isException = False

    On Error GoTo Exception

    If Nz(Me.txt_契約数, "0") <> "0" Then
        MsgBox "この割引は適用した受付または契約が既にあるため削除は出来ません。", vbInformation, P_ERROR
        Exit Sub
    End If

    If MsgBox("データを削除します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, P_KAKUNIN) = vbYes Then
        '削除
        Call DEL_RTN
        Call Form_FKS700.createNebikiData
        Call Form_FKS700.Requery
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_削除_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_修正_Click
'        MODULE_ID        :cmd_修正_Click
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_修正_Click()
    
    Dim anStrParameter As String
    Dim isException    As Boolean

    isException = False

    On Error GoTo Exception

    anStrParameter = ""

    ' (0)割引番号 - (1)割引期間 - (2)割引形式 - (3)割引値 - (4)受付種別 - (5)サイズFrom - (6)サイズTo(5) - (7)段 - (8)用途 -
    ' (9)利用期間 - (10)文言自動生成 - (11)お客様提示文言 - (12)タイトル - (13)処理モード
'D:2008.11.13 N.IMAI Start
'    anStrParameter = Nz(Me.txt_割引番号.VALUE, "") & "," & _
'                     Nz(Me.txt_割引期間値.VALUE, "") & "," & _
'                     Nz(Me.txt_割引形式値.VALUE, "") & "," & _
'                     Nz(Me.txt_割引値.VALUE, "") & "," & _
'                     Nz(Me.txt_受付値.VALUE, "") & "," & _
'                     Nz(Me.txt_サイズFrom.VALUE, "") & "," & _
'                     Nz(Me.txt_サイズTo.VALUE, "") & "," & _
'                     Nz(Me.txt_段値.VALUE, "") & "," & _
'                     Nz(Me.txt_用途値.VALUE, "") & "," & _
'                     Nz(Me.txt_利用期間値.VALUE, "") & "," & _
'                     Nz(Me.txt_文言自動生成.VALUE, "") & "," & _
'                     Nz(Me.txt_内容.VALUE, "") & "," & _
'                     "修正" & "," & _
'                     "1"
'D:2008.11.13 N.IMAI End
'I:2008.11.13 N.IMAI Start
    anStrParameter = Nz(Me.txt_割引番号.VALUE, "") & vbTab & _
                     Nz(Me.txt_割引期間値.VALUE, "") & vbTab & _
                     Nz(Me.txt_割引形式値.VALUE, "") & vbTab & _
                     Nz(Me.txt_割引値.VALUE, "") & vbTab & _
                     Nz(Me.txt_受付値.VALUE, "") & vbTab & _
                     Nz(Me.txt_サイズFrom.VALUE, "") & vbTab & _
                     Nz(Me.txt_サイズTo.VALUE, "") & vbTab & _
                     Nz(Me.txt_段値.VALUE, "") & vbTab & _
                     Nz(Me.txt_用途値.VALUE, "") & vbTab & _
                     Nz(Replace(Me.txt_利用期間値.VALUE, ",", vbTab), "") & vbTab & _
                     Nz(Replace(Me.txt_文言自動生成.VALUE, ",", vbTab), "") & vbTab & _
                     Nz(Me.txt_内容.VALUE, "") & vbTab & _
                     "修正" & vbTab & _
                     "1"
'I:2008.11.13 N.IMAI End
    doCmd.OpenForm "FKS710", , , , , acDialog, anStrParameter

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_修正_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_追加_Click
'        MODULE_ID        :cmd_追加_Click
'        NOTE             :
'        CREATE_DATE      :2008/08/20
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_追加_Click()
    
    Dim strParameter As String
    Dim isException  As Boolean

    isException = False
    
    On Error GoTo Exception
    
'2017/11/01 start ADD EGL
    If Me.Parent.Form.lbl_標題.Caption = "割引選択" Then
        Form_FKS450.p450_RET = True
        Form_FKS450.p450_WaribikiNo = Nz(txt_割引番号, "")
        Form_FKS450.p450_Youto = Nz(Me.txt_用途値, "")
        Form_FKS450.p450_Mongon = Nz(txt_内容, "")
        doCmd.Close acForm, Me.Parent.NAME
        Exit Sub
    End If
'2017/11/01 end ADD EGL
    
    ' (0)タイトル - (1)モード   - (2)ヤードコード - (3)部門コード - (4)割引番号
    ' (5)手動文言
'D:2008.11.13 N.IMAI Start
'    strParameter = "新規追加" _
'                 & ",0" _
'                 & "," & Me.Parent.Form.txt_ヤード _
'                 & "," & pstrBumonCode _
'                 & "," & Nz(txt_割引番号, "") _
'                 & "," & Nz(txt_内容)
'D:2008.11.13 N.IMAI End
'I:2008.11.13 N.IMAI Start
    strParameter = "新規追加" & vbTab & _
                   "0" & vbTab & _
                   Me.Parent.Form.txt_ヤード & vbTab & _
                   pstrBumonCode & vbTab & _
                   Nz(txt_割引番号, "") & vbTab & _
                   Nz(txt_内容) & vbTab
'I:2008.11.13 N.IMAI End
    doCmd.OpenForm "FKS721", , , , , acDialog, strParameter

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_追加_Click", True)

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
    
    '画面内容切替
    lbl_設定有無.Visible = Not IsNull(Me.Parent.OpenArgs)
    lbl_設定有無背景.Visible = Not IsNull(Me.Parent.OpenArgs)
    lbl_追加.Visible = Not IsNull(Me.Parent.OpenArgs)
    lbl_追加背景.Visible = Not IsNull(Me.Parent.OpenArgs)
    txt_設定有無.Visible = Not IsNull(Me.Parent.OpenArgs)
    cmd_追加.Visible = Not IsNull(Me.Parent.OpenArgs)
    
    lbl_変更.Visible = IsNull(Me.Parent.OpenArgs)
    lbl_変更背景.Visible = IsNull(Me.Parent.OpenArgs)
    cmd_修正.Visible = IsNull(Me.Parent.OpenArgs)
    cmd_削除.Visible = IsNull(Me.Parent.OpenArgs)

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
    Dim rsDataDel            As Object    ' マスタ、トラン削除用
    Dim strDCNTM_NO          As String    ' 割引番号

    On Error GoTo Exception

    DEL_RTN = False
    '2009/06/30 MOD <S> 割引マスタ KASE_DBへ移行
    'Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set objAdoDbConnection = MSZZ025.ADODB_Connection()
    '2009/06/30 MOD <E>

    ' トランザクション開始
    objAdoDbConnection.BeginTrans

    ' 割引番号設定
    strDCNTM_NO = Me.txt_割引番号

    strSQL = ""
    strSQL = strSQL & " SELECT * FROM DCNT_MAST "
    strSQL = strSQL & " WHERE DCNTM_NO = '" & strDCNTM_NO & "' "
    '2009/06/30 INS <S> 割引マスタ KASE_DBへ移行
    strSQL = strSQL & " AND  DCNTM_BUMOC = '" & pstrBumonCode & "' "
    '2009/06/30 INS <E>
    Set rsDataDel = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)

    rsDataDel.Delete

    rsDataDel.Close
    Set rsDataDel = Nothing
    
    DEL_RTN = True
    ' コミット
    objAdoDbConnection.CommitTrans
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

Exception:
    If Not rsDataDel Is Nothing Then
        rsDataDel.Close
        Set rsDataDel = Nothing
    End If

    If Err <> 0 Then
        objAdoDbConnection.RollbackTrans
        objAdoDbConnection.Close
        Set objAdoDbConnection = Nothing
        Call Err.Raise(Err.Number, "DEL_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
