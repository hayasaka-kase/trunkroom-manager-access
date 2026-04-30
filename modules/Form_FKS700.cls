VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　割引マスタ一覧
'   プログラムＩＤ　：　FKS700
'   作　成　日　　　：  2008/07/25
'   作　成　者　　　：  T.Suzuki
'   Ver             ：  0.0
'**********************************************
'   修　正　日　　　：  2008/11/13
'   修　正　者　　　：  K.KINEBUCHI
'   内　　　容　　　：  次画面へ文言を引き数で渡す際に
'   　　　　　　　　　　区切り文字がカンマの場合、不具合があるので
'   　　　　　　　　　　区切り文字をタブに変更
'   Ver             ：  0.1
'**********************************************
'   修　正　日　　　：  2009/01/14
'   修　正　者　　　：  hirano
'   内　　　容　　　：  割引形式のラベル変更「ｎ円割引またはn%引き」→「ｎ円割引」とする。
'                       処理は変更なし。
'   Ver             ：  0.2
'**********************************************
'   修　正　日　　　：  2009/06/30
'   修　正　者　　　：  hirano
'   内　　　容　　　：  割引マスタKASE_DB移行ﾆ伴う変更
'   Ver             ：  0.3
'**********************************************
'   修　正　日　　　：  2011/08/29
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  事務手数料割引
'   Ver             ：  0.4
'**********************************************
'   修　正　日　　　：  2013/06/03
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  終了メッセージを表示させないように修正
'   Ver             ：  0.5
'**********************************************
'   修　正　日　　　：  2017/12/01
'   修　正　者　　　：  イーグルソフト
'   内　　　容　　　：  キャンペーン一括設定対応
'   Ver             ：  0.6
'**********************************************
'   修　正　日　　　：  2019/03/25
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  表示する割引をPGPAで制御するように修正
'   Ver             ：  0.7
'**********************************************
'   修　正　日　　　：  2023/12/08
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  契約数でなく割引中の契約数に変更
'   Ver             ：  0.8
'**********************************************

Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS700"

Private Const P_TITLE_NEBIKI_MASTER     As String = "割引マスタ一覧"
Private Const P_TITLE_NEBIKI_ADD        As String = "追加割引マスタ 検索"

Private Const P_ROOM_WORK_TABLE         As String = "WK_FKS700"
Private Const P_NEBIKI_LIST_WORK_TABLE  As String = "WK_FKS700_S01_TEMP"

Private Const P_ALL_NINE                As String = "9999999999"

' 不足桁０埋め用
Private Const P_ZERO_SUPPRESS           As String = "0000000000"

' 桁数
Private Const P_DIVISION_LENGTH         As Integer = 2       ' 区分
Private Const P_STEP_LENGTH             As Integer = 1       ' 段
Private Const P_SIZE_LENGTH             As Integer = 5       ' サイズ
Private Const P_YOTO_LENGTH             As Integer = 2       ' 用途
Private Const P_RIYO_KIKAN_LENGTH       As Integer = 1       ' 利用期間

Private Const P_SIZE_MAX_VALUE          As Single = 999.99   ' サイズ

Private Const P_MINUS                   As Integer = 45      ' マイナス符号

'コントロールの種類
Private Const P_CONTROL_数値            As Integer = 1       ' 数値

Private Const P_KAKUNIN                 As String = "確認"
Private Const P_ERROR                   As String = "エラー"

Private pobjDB                          As Object

Private pstrBumonCode                   As String            ' 部門コード
Private anStrParameter                  As String            ' パラメーター

Public gstrUserName                     As String
Private pstrParameter()                 As String            ' 初期表示データ

Private var変更日付                     As Variant  '日付入力チェック用

'==============================================================================*
'
'        MODULE_NAME      :cmb_割引期間_AfterUpdate
'        MODULE_ID        :cmb_割引期間_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_割引期間_AfterUpdate()

    Me.txt_割引期間 = fncAfterUpdateCommon(Me.cmb_割引期間, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_割引期間_BeforeUpdate
'        MODULE_ID        :cmb_割引期間_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_割引期間_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_割引期間, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_AfterUpdate
'        MODULE_ID        :cmb_受付種別_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_AfterUpdate()

    Me.txt_受付種別 = fncAfterUpdateCommon(Me.cmb_受付種別, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_BeforeUpdate
'        MODULE_ID        :cmb_受付種別_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_受付種別, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_KeyPress
'        MODULE_ID        :cmb_受付種別_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_受付種別, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズFrom_AfterUpdate
'        MODULE_ID        :txt_サイズFrom_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズFrom_AfterUpdate()

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズFrom_BeforeUpdate
'        MODULE_ID        :txt_サイズFrom_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズFrom_BeforeUpdate(Cancel As Integer)

    If fncChkTextInput2(Me.txt_サイズFrom, P_SIZE_MAX_VALUE, P_SIZE_LENGTH) = False Then
        MsgBox "桁数オーバーしています｡", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズFrom_KeyPress
'        MODULE_ID        :txt_サイズFrom_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズFrom_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    ' マイナス符号が入力された場合、入力を無効にする。
    If KeyAscii = P_MINUS Then KeyAscii = 0: Exit Sub
    
    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.txt_サイズFrom, 3, 2, KeyAscii)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズTo_AfterUpdate
'        MODULE_ID        :txt_サイズTo_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズTo_AfterUpdate()

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズTo_BeforeUpdate
'        MODULE_ID        :txt_サイズTo_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズTo_BeforeUpdate(Cancel As Integer)

    If fncChkTextInput2(Me.txt_サイズTo, P_SIZE_MAX_VALUE, P_SIZE_LENGTH) = False Then
        MsgBox "桁数オーバーしています｡", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_サイズTo_KeyPress
'        MODULE_ID        :txt_サイズTo_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_サイズTo_KeyPress(KeyAscii As Integer)

    On Error Resume Next

    ' マイナス符号が入力された場合、入力を無効にする。
    If KeyAscii = P_MINUS Then KeyAscii = 0: Exit Sub

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.txt_サイズTo, 3, 2, KeyAscii)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_段_AfterUpdate
'        MODULE_ID        :cmb_段_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_段_AfterUpdate()

    Me.txt_段 = fncAfterUpdateCommon(Me.cmb_段, P_STEP_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_段_BeforeUpdate
'        MODULE_ID        :cmb_段_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_段_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_段, P_STEP_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_段_KeyPress
'        MODULE_ID        :cmb_段_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_段_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_段, KeyAscii, P_STEP_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_AfterUpdate
'        MODULE_ID        :cmb_用途_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_AfterUpdate()

    Me.txt_用途 = fncAfterUpdateCommon(Me.cmb_用途, P_YOTO_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_BeforeUpdate
'        MODULE_ID        :cmb_用途_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_用途, P_YOTO_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_KeyPress
'        MODULE_ID        :cmb_用途_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_用途, KeyAscii, P_YOTO_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_利用期間_AfterUpdate
'        MODULE_ID        :cmb_利用期間_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用期間_AfterUpdate()

    Me.txt_利用期間 = fncAfterUpdateCommon(Me.cmb_利用期間, P_RIYO_KIKAN_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_利用期間_BeforeUpdate
'        MODULE_ID        :cmb_利用期間_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用期間_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(Me.cmb_利用期間, P_RIYO_KIKAN_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_利用期間_KeyPress
'        MODULE_ID        :cmb_利用期間_KeyPress
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用期間_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_利用期間, KeyAscii, P_RIYO_KIKAN_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_検索_Click
'        MODULE_ID        :cmd_検索_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索_Click()

    ' 検索条件入力チェック
    If fncSearchInputCheck = False Then
        Me.txt_サイズFrom.SetFocus
        Exit Sub
    End If

    ' 検索
    Call createNebikiData

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        MODULE_ID        :cmd_終了_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

    Dim isException As Boolean

    isException = False

    On Error GoTo Exception
    '' 2013/06/03 M.HONDA START
'    If lbl_ヤード情報.Visible = False Then
'        If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) <> vbYes Then
'            GoTo Finally
'        End If
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
'        MODULE_NAME      :cmd_新規作成_Click
'        MODULE_ID        :cmd_新規作成_Click
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_新規作成_Click()

    Dim isException As Boolean

    isException = False

    On Error GoTo Exception

'    Call deleteWorkTable

    ' (0)割引番号 - (1)割引期間 - (2)割引形式 - (3)割引値 - (4)受付種別 - (5)サイズFrom - (6)サイズTo(5) - (7)段 - (8)用途 -
    ' (9)利用期間 - (10)文言自動生成 - (11)お客様提示文言 - (12)タイトル - (13)処理モード
'    anStrParameter = ",,,,,,,,,,,,新規登録,0"                      'DELETE 2008/11/13 K.KINEBUCHI
'INSERT 2008/11/13 K.KINEBUCHI Start
    anStrParameter = vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     vbTab & _
                     "新規登録" & vbTab & _
                     "0"
'INSERT 2008/11/13 K.KINEBUCHI end
    doCmd.OpenForm "FKS710", , , , , acDialog, anStrParameter

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_新規作成_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        MODULE_ID        :Form_Activate
'        NOTE             :
'        CREATE_DATE      :2008/07/25
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
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Dim isException As Boolean

    isException = False

    On Error GoTo Exception

    ' テーブル削除
    Call dropWorkTable

    ' コネクションクローズ
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
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Dim isException As Boolean

    isException = False

    On Error GoTo Exception
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_標題.BackColor = RGB(255, 0, 204)
    End Select

    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        ' テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb

    ' ワークテーブルあったら削除
    If fncTableExist(P_NEBIKI_LIST_WORK_TABLE) Then
        Call dropWorkTable
    End If

    ' ワークテーブル作成
    Call createWorkTable

    gstrUserName = MSZZ000.LsGetUserName

    ' パラメータ設定
    txt_ヤード = ""
    If Not IsNull(Me.OpenArgs) Then
        pstrParameter = Split(Nz(Me.OpenArgs), ",")
        
'2017/12/01 start ADD EGL
        If pstrParameter(0) = "割引選択" Then
            txt_ヤード = pstrParameter(1)
            lbl_標題.Caption = pstrParameter(0)
            cmd_新規作成.Visible = False
'2017/12/01 end ADD EGL
        Else
            txt_ヤード = pstrParameter(1)
            lbl_標題.Caption = P_TITLE_NEBIKI_ADD
            cmd_新規作成.Visible = False
            lbl_ヤード情報.Caption = "現在 " _
                                   & pstrParameter(1) _
                                   & ":" _
                                   & pstrParameter(2) _
                                   & " ヤードには合計" _
                                   & pstrParameter(3) _
                                   & "件の割引が設定されています。"
        End If
    Else
        lbl_標題.Caption = P_TITLE_NEBIKI_MASTER
        lbl_ヤード情報.Visible = False
    End If

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
'        CREATE_DATE      :2008/07/25
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
'                         :補足※ダイアログのRecordSource代入（２つ）は
'                         :      ダイアログの方で代入する。ここでは行わない。
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub dropWorkTable()

    On Error GoTo Exception

    Me.割引マスタ一覧.Form.RecordSource = ""
    Me.割引マスタ一覧.Form.OrderBy = ""
    Me.割引マスタ一覧.Form.Requery
    pobjDB.Execute "DROP TABLE " & P_NEBIKI_LIST_WORK_TABLE & " ;", dbFailOnError

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "dropWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル作成
'        MODULE_ID        :createWorkTable
'        NOTE             :
'                         :補足※ダイアログのRecordSource代入（２つ）は
'                         :      ダイアログの方で代入する。ここでは行わない。
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub createWorkTable()

    Dim strSQL As String

    On Error GoTo Exception

    ' 一覧画面用ワークテーブル
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE " & P_NEBIKI_LIST_WORK_TABLE & " ("
    strSQL = strSQL & " DCNTM_NO             VARCHAR(6) "
    strSQL = strSQL & ",DCNTM_NO_SORT_KEY    VARCHAR(6) "
    strSQL = strSQL & ",DCNTM_MTEXT          VARCHAR(42) "
    strSQL = strSQL & ",NEBIKI_KIKAN_VALUE   VARCHAR(3) "
    strSQL = strSQL & ",NBIKI_KEISHIKI_VALUE VARCHAR(2) "
    strSQL = strSQL & ",NEBIKI_VALUE         VARCHAR(9) "
    strSQL = strSQL & ",NEBIKI_KIKAN         VARCHAR(20) "
    strSQL = strSQL & ",UKETSUKE_VALUE       VARCHAR(2)"
    strSQL = strSQL & ",UKETSUKE             VARCHAR(20) "
    strSQL = strSQL & ",SIZE_FROM_VALUE      VARCHAR(6) "
    strSQL = strSQL & ",SIZE_TO_VALUE        VARCHAR(6) "
    strSQL = strSQL & ",SIZE                 VARCHAR(20) "
    strSQL = strSQL & ",STEP_VALUE           VARCHAR(2) "
    strSQL = strSQL & ",STEP                 VARCHAR(20) "
    strSQL = strSQL & ",YOTO_VALUE           VARCHAR(2) "
    strSQL = strSQL & ",YOTO                 VARCHAR(20) "
    strSQL = strSQL & ",RIYO_KIKAN_VALUE     VARCHAR(2) "
    strSQL = strSQL & ",RIYO_KIKAN           VARCHAR(20) "
    strSQL = strSQL & ",DCNTM_ATEXT          VARCHAR(150) "
    strSQL = strSQL & ",KEIYAKU_SU           VARCHAR(5) "
    strSQL = strSQL & ",SETTEI_SU            VARCHAR(10) "
    strSQL = strSQL & ")"
    pobjDB.Execute strSQL, dbFailOnError

    ' RecordSource設定
    ' ソート順：割引番号の降順
    Me.割引マスタ一覧.Form.RecordSource = "SELECT * FROM " & P_NEBIKI_LIST_WORK_TABLE & " ORDER BY DCNTM_NO DESC;"
    Me.割引マスタ一覧.Form.OrderBy = "DCNTM_NO_SORT_KEY desc"

    ' サブフォームのリフレッシュ
    Me.割引マスタ一覧.Form.Requery

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "createWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルデータ削除
'        MODULE_ID        :deleteWorkTable
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub deleteWorkTable()
    
    Dim strSQL              As String
    
    On Error GoTo Exception

    Call dropWorkTable
    Call createWorkTable
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "deleteWorkTable" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧作成
'        MODULE_ID        :createNebikiData
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub createNebikiData()

    On Error GoTo Exception

    ' ワークテーブル削除
    Call deleteWorkTable

    ' 割引一覧データ取得
    Call getSelectNebikiData

    ' サブフォームのリフレッシュ
    Me.割引マスタ一覧.Form.Requery

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "createNebikiData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧表時用ワークテーブル更新
'        MODULE_ID        :getSelectNebikiData
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getSelectNebikiData()

    Dim objAdoDbConnection   As Object
    Dim rsData               As Object       'ADODB.Recordset
    Dim rsNebikiData         As Recordset
    Dim strSQL               As String
    Dim strSize              As String
    Dim intKEIYAKU_CNT       As Integer
    Dim strKaseSqlSvr        As String       'INS 2009/06/30 hirano
    On Error GoTo Exception
    '2009/06/30 INS <S> 割引マスタ KASE_DBへ移行
    'SQL-Server名
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
     '2009/06/30 INS <E>
    ' 割引一覧データ取得
    strSQL = ""
    strSQL = " SELECT    DCNT_MAST.DCNTM_NO         DCNTM_NO "              ' 割引番号
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_MTEXT      DCNTM_MTEXT "           ' 内容(手動文言)
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_PERIOD     NEBIKI_KIKAN_VALUE "    ' 割引期間(値)
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_TYPE       NBIKI_KEISHIKI_VALUE "  ' 割引形式(値)
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_VALUE      NEBIKI_VALUE "          ' 割引値
    strSQL = strSQL & " ,NAME1.NAME_RYAK            NEBIKI_KIKAN "          ' 割引期間
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_GENKBN     UKETSUKE_VALUE "        ' 受付(値)
    strSQL = strSQL & " ,NAME2.NAME_RYAK            UKETSUKE "              ' 受付
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_SIZE_FROM  SIZE_FROM "             ' サイズ(From)
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_SIZE_TO    SIZE_TO "               ' サイズ(To)
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_FLOOR      STEP_VALUE "            ' 段(値)
    strSQL = strSQL & " ,NAME3.NAME_RYAK            STEP "                  ' 段
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_USAGE      YOTO_VALUE "            ' 用途(値)
    strSQL = strSQL & " ,NAME4.NAME_RYAK            YOTO "                  ' 用途
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_USE_PERIOD RIYO_KIKAN_VALUE "      ' 利用期間(値)
    strSQL = strSQL & " ,NAME5.NAME_RYAK            RIYO_KIKAN "            ' 利用期間
    strSQL = strSQL & " ,DCNT_MAST.DCNTM_ATEXT      DCNTM_ATEXT "           ' 自動文言
    'strSQL = strSQL & " ,COUNT(*)                   KEIYAKU_CNT "
    strSQL = strSQL & " ,COUNT(DCRAT_DCNT_NO)       KEIYAKU_CNT "               'INSERT 2023/12/08 N.IMAI
    
    strSQL = strSQL & " FROM " & strKaseSqlSvr & ".dbo.DCNT_MAST LEFT OUTER JOIN NAME_MAST AS NAME1 ON NAME1.NAME_CODE = DCNT_MAST.DCNTM_PERIOD     AND NAME1.NAME_ID = '276' "
    'DELETE 2023/12/08 N.IMAI strSQL = strSQL & "     LEFT OUTER JOIN DCRA_TRAN ON DCRAT_DCNT_NO = DCNTM_NO"
    strSQL = strSQL & "                LEFT OUTER JOIN NAME_MAST AS NAME2 ON NAME2.NAME_CODE = DCNT_MAST.DCNTM_GENKBN     AND NAME2.NAME_ID = '083' "
    strSQL = strSQL & "                LEFT OUTER JOIN NAME_MAST AS NAME3 ON NAME3.NAME_CODE = DCNT_MAST.DCNTM_FLOOR      AND NAME3.NAME_ID = '090' "
    strSQL = strSQL & "                LEFT OUTER JOIN NAME_MAST AS NAME4 ON NAME4.NAME_CODE = DCNT_MAST.DCNTM_USAGE      AND NAME4.NAME_ID = '086' "
    strSQL = strSQL & "                LEFT OUTER JOIN NAME_MAST AS NAME5 ON NAME5.NAME_CODE = DCNT_MAST.DCNTM_USE_PERIOD AND NAME5.NAME_ID = '249' "
    
    'INSERT 2023/12/08 N.IMAI Start
    strSQL = strSQL & "  LEFT OUTER JOIN "
    strSQL = strSQL & "  ( "
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "     DCRAT_DCNT_NO"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "     KASE_DB.dbo.KEIY_TRAN"
'    strSQL = strSQL & " INNER JOIN"
'    strSQL = strSQL & "     KASE_DB.dbo.URIA_TRAN"
'    strSQL = strSQL & " ON"
'    strSQL = strSQL & "     URIAT_KEIYB = KEIYT_KEIYB"
'    strSQL = strSQL & " AND"
'    strSQL = strSQL & "     URIAT_URYMD = '202312'"
'    strSQL = strSQL & " INNER JOIN"
'    strSQL = strSQL & "     KASE_DB.dbo.JUKU_TRAN"
'    strSQL = strSQL & " ON"
'    strSQL = strSQL & "     JUKUT_KEIYB = KEIYT_KEIYB"
'    strSQL = strSQL & " AND"
'    strSQL = strSQL & "     JUKUT_SEIKI = URIAT_URIAI"
'    strSQL = strSQL & " AND"
'    strSQL = strSQL & "     JUKUT_ZAPII = '91'"
    strSQL = strSQL & " INNER JOIN"
    strSQL = strSQL & "     DCRA_TRAN"
    strSQL = strSQL & " ON"
    strSQL = strSQL & "     DCRAT_ACPTNO = KEIYT_KEIYB"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "     DCRAT_SEIKYU_KBN = '1'"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "     ISNULL(KEIYT_KAIYD, '29991231') >= '" & Format(Now, "yyyyMMdd") & "'"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "     KEIYT_KAISD <= ISNULL(KEIYT_KAIYD, '29991231')"
    strSQL = strSQL & " )   AS DCRAT"
    strSQL = strSQL & "    ON DCRAT_DCNT_NO = DCNTM_NO "
    'INSERT 2023/12/08 N.IMAI Start
    strSQL = strSQL & " WHERE DCNT_MAST.DCNTM_BUMOC = '" & pstrBumonCode & "' "  ' 部門コード
    strSQL = strSQL & "   AND DCNT_MAST.DCNTM_KIND  = 1 "                        ' 割引種別(契約系割引)
    
    '2019/03/19 M.HONDA INS
    '利用している割引を絞る
    Dim sDCNTM_NO As String
    sDCNTM_NO = Nz(DLookup("PGPAT_PARAN", "PGPA_TABL", "PGPAT_PGP1B = 'FKS700' AND PGPAT_PGP2B =" & "'" & pstrBumonCode & "'"))
    If sDCNTM_NO <> "" Then
        strSQL = strSQL & "   AND DCNT_MAST.DCNTM_NO IN  (" & sDCNTM_NO & ")"
    End If

    ' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ' Where句作成
    ' << 割引期間 >>
    If cmb_割引期間.VALUE <> "" Then
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_PERIOD = '" & cmb_割引期間.VALUE & "'"
    End If
    
'    Select Case Me.opg_割引期間.VALUE
'        Case 1  ' *** 全て ***
'            ' 条件指定なし
'
'        Case 2  ' *** nヵ月後は定額 ***
'            strSQL = strSQL & " AND DCNT_MAST.DCNTM_PERIOD <> '0' "
'
'        Case 3  ' *** 解約まで割引 ***
'            strSQL = strSQL & " AND DCNT_MAST.DCNTM_PERIOD = '0' "
'    End Select






    ' << 割引形式 >>
    Select Case Me.opg_割引形式.VALUE
        Case 1  ' *** 全て ***
            ' 条件指定なし

        Case 2  ' *** 円指定 ***
            strSQL = strSQL & " AND DCNT_MAST.DCNTM_TYPE = '1' "

        Case 3  ' *** n円割引またはn%引き ***
            strSQL = strSQL & " AND DCNT_MAST.DCNTM_TYPE IN ('2', '3')"
        '----↓↓↓↓----20110829--M.HONDA--Ins-------↓↓↓↓---<s>
        Case 10  ' *** 事務手数料割引 ***
            strSQL = strSQL & " AND DCNT_MAST.DCNTM_TYPE ='10'"
        '----↑↑↑↑----20110829--M.HONDA--Ins-------↑↑↑↑---<e>
    End Select

    ' << 受付種別 >>
    If Nz(Me.cmb_受付種別.VALUE, "") <> "" Then
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_GENKBN = '" & Me.cmb_受付種別.VALUE & "' "
    End If

    ' << サイズ >>
    If Nz(Me.txt_サイズFrom.VALUE, "") <> "" And Nz(Me.txt_サイズTo.VALUE, "") = "" Then
        ' Fromのみ入力
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_SIZE_FROM >= " & Me.txt_サイズFrom.VALUE

    ElseIf Nz(Me.txt_サイズFrom.VALUE, "") = "" And Nz(Me.txt_サイズTo.VALUE, "") <> "" Then
        ' Toのみ入力
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_SIZE_TO <= " & Me.txt_サイズTo.VALUE

    ElseIf Nz(Me.txt_サイズFrom.VALUE, "") <> "" And Nz(Me.txt_サイズTo.VALUE, "") <> "" Then
        ' FromTo入力
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_SIZE_FROM >= " & Me.txt_サイズFrom.VALUE
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_SIZE_TO   <= " & Me.txt_サイズTo.VALUE

    ElseIf Nz(Me.txt_サイズFrom.VALUE, "") = "" And Nz(Me.txt_サイズTo.VALUE, "") = "" Then
        ' 未入力
        ' 条件指定なし
    End If

    ' << 段 >>
    If Nz(Me.cmb_段, "") <> "" Then
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_FLOOR = '" & Me.cmb_段.VALUE & "' "
    End If

    ' << 用途 >>
    If Nz(Me.cmb_用途, "") <> "" Then
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_USAGE = '" & Me.cmb_用途.VALUE & "' "
    End If

    ' << 利用期間 >>
    If Nz(Me.cmb_利用期間, "") <> "" Then
        strSQL = strSQL & " AND DCNT_MAST.DCNTM_USE_PERIOD = '" & Me.cmb_利用期間.VALUE & "' "
    End If
    
    strSQL = strSQL & " Group BY"
    strSQL = strSQL & "     DCNT_MAST.DCNTM_NO , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_MTEXT , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_PERIOD , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_TYPE , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_VALUE , "
    strSQL = strSQL & "     NAME1.NAME_RYAK , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_GENKBN , "
    strSQL = strSQL & "     NAME2.NAME_RYAK , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_SIZE_FROM , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_SIZE_TO , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_FLOOR , "
    strSQL = strSQL & "     NAME3.NAME_RYAK , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_USAGE , "
    strSQL = strSQL & "     NAME4.NAME_RYAK , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_USE_PERIOD , "
    strSQL = strSQL & "     NAME5.NAME_RYAK , "
    strSQL = strSQL & "     DCNT_MAST.DCNTM_ATEXT "
   
    ' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)

    If rsData.EOF = True Then
        MsgBox "指定した検索条件に該当するデータはありません。", vbInformation, P_ERROR
    Else
        While Not rsData.EOF
            strSQL = ""
            strSQL = strSQL & "SELECT * FROM " & P_NEBIKI_LIST_WORK_TABLE
            Set rsNebikiData = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    
            rsNebikiData.AddNew
    
            rsNebikiData.Fields("DCNTM_NO") = rsData.Fields("DCNTM_NO")
            rsNebikiData.Fields("DCNTM_NO_SORT_KEY") = Format(rsData.Fields("DCNTM_NO"), "000000")
            rsNebikiData.Fields("DCNTM_MTEXT") = rsData.Fields("DCNTM_MTEXT")
            rsNebikiData.Fields("NEBIKI_KIKAN_VALUE") = rsData.Fields("NEBIKI_KIKAN_VALUE")
            rsNebikiData.Fields("NBIKI_KEISHIKI_VALUE") = rsData.Fields("NBIKI_KEISHIKI_VALUE")
            rsNebikiData.Fields("NEBIKI_VALUE") = rsData.Fields("NEBIKI_VALUE")
            rsNebikiData.Fields("NEBIKI_KIKAN") = rsData.Fields("NEBIKI_KIKAN")
            rsNebikiData.Fields("UKETSUKE_VALUE") = rsData.Fields("UKETSUKE_VALUE")
            rsNebikiData.Fields("UKETSUKE") = rsData.Fields("UKETSUKE")
    
            rsNebikiData.Fields("SIZE_FROM_VALUE") = rsData.Fields("SIZE_FROM")
            rsNebikiData.Fields("SIZE_TO_VALUE") = rsData.Fields("SIZE_TO")
    
            If (Nz(rsData.Fields("SIZE_FROM"), "") = "" And Nz(rsData.Fields("SIZE_TO"), "") = "") Then
                ' From～Toの両方が未設定
                strSize = "－"
    
            ElseIf (rsData.Fields("SIZE_FROM") = rsData.Fields("SIZE_TO")) Then
                ' From～Toの両方が同じ値
                strSize = rsData.Fields("SIZE_FROM") & "帖"
    
            Else
                ' From～Toの両方が異なる値
                strSize = rsData.Fields("SIZE_FROM") & "～" & rsData.Fields("SIZE_TO") & "帖"
            End If
    
            rsNebikiData.Fields("SIZE") = strSize
    
            rsNebikiData.Fields("STEP_VALUE") = rsData.Fields("STEP_VALUE")
            rsNebikiData.Fields("STEP") = Nz(rsData.Fields("STEP"), "－")
            rsNebikiData.Fields("YOTO_VALUE") = rsData.Fields("YOTO_VALUE")
            rsNebikiData.Fields("YOTO") = Nz(rsData.Fields("YOTO"), "－")
            rsNebikiData.Fields("RIYO_KIKAN_VALUE") = rsData.Fields("RIYO_KIKAN_VALUE")
            rsNebikiData.Fields("RIYO_KIKAN") = Nz(rsData.Fields("RIYO_KIKAN"), "－")
            rsNebikiData.Fields("DCNTM_ATEXT") = rsData.Fields("DCNTM_ATEXT")
            rsNebikiData.Fields("KEIYAKU_SU") = rsData.Fields("KEIYAKU_CNT")
    
            ' 契約数取得
'            intKEIYAKU_CNT = fncGetKeiyaku_SU(rsData.Fields("DCNTM_NO"), objAdoDbConnection)
'            rsNebikiData.Fields("KEIYAKU_SU") = intKEIYAKU_CNT
            ' 設定数取得
'            If txt_ヤード <> "" Then
'                rsNebikiData.Fields("SETTEI_SU") = fncGetSettei_SU(rsData.Fields("DCNTM_NO"), objAdoDbConnection)
'            End If
    
            rsNebikiData.UPDATE
            rsNebikiData.Close
            Set rsNebikiData = Nothing
    
            rsData.MoveNext
        Wend
    End If
    
    rsData.Close
    Set rsData = Nothing
    objAdoDbConnection.Close
    Set objAdoDbConnection = Nothing

    Exit Function

Exception:
    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not rsNebikiData Is Nothing Then: rsNebikiData.Close: Set rsNebikiData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    Call Err.Raise(Err.Number, "getSelectNebikiData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :検索条件入力チェック
'        MODULE_ID        :fncSearchInputCheck
'        NOTE             :
'        CREATE_DATE      :2008/07/25
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSearchInputCheck() As Boolean

    On Error GoTo Exception

    fncSearchInputCheck = False

    If (Nz(Me.txt_サイズFrom.VALUE, "") <> "" And Nz(Me.txt_サイズTo.VALUE, "") <> "") Then
        If Me.txt_サイズFrom.VALUE > Me.txt_サイズTo.VALUE Then
            ' 大小チェックエラー
            ' サイズの範囲指定が矛盾しています。 \n From <= To となるように入力して下さい。
            MsgBox "サイズの範囲指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbInformation, P_ERROR
            Exit Function
        End If
    End If

    fncSearchInputCheck = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncSearchInputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncBeforeUpdateCommon
'        機能             :更新前イベント時の処理
'                         :入力チェックを行う
'        IN               :①対象コントロール
'                         :②入力最大桁数
'                         :③定数値「コントロールの種類」参照
'                         :④更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncBeforeUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer, ByRef Cancel As Integer) As Boolean

    Dim strTemp     As String
    Dim blnRet      As Boolean

    If Left(objTarget.NAME, 3) = "cmb" Then
        'コンボボックス入力チェック
        fncBeforeUpdateCommon = fncChkComboInput(objTarget, intLength, Cancel)
    Else
        '入力桁数チェック
        fncBeforeUpdateCommon = fncChkTextInput(objTarget, intLength, Cancel)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncAfterUpdateCommon
'        機能             :更新後イベント時の処理
'                         :数値項目は０埋めまたは０消しした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'                         :③定数値「コントロールの種類」参照
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :年月は"yyyy/mm"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncAfterUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer) As String

    fncAfterUpdateCommon = ""

    '体裁を整える
    Select Case intKind
        Case P_CONTROL_数値 '数値項目は０埋め
            objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))

        Case Else
            fncAfterUpdateCommon = Nz(objTarget)
    End Select

    ' コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        If Nz(objTarget) <> "" And _
           objTarget.ListCount <> 0 _
        Then
            fncAfterUpdateCommon = Nz(objTarget.Column(1))
        End If
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDateCheck
'        機能             :日付入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    fncDateCheck = True
    
    var変更日付 = CmDateChecker(txtTarget)
    If IsNull(var変更日付) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkComboInput
'        機能             :コンボボックス入力チェック
'        IN               :第一引数：対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :0=正常
'                          1=入力規則に反している
'                          2=入力値がリストに存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer, ByRef Cancel As Integer) As Integer

    Dim strCode     As String
    Dim intRet      As Integer
    Dim intCount    As Integer
    Dim strValue    As String
    Dim blnFound    As Boolean
    
    intRet = 0
    
    strCode = Nz(cmbCheck, "")
    
    ' 空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkComboInput_End
    End If
    
    If Not fncNumericCheck(cmbCheck, intMaxLength, Cancel) Then
        intRet = 1
        GoTo ChkComboInput_End
    End If
    
    ' 入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        blnFound = False
        ' 入力値を０サプレス
        strValue = Format(cmbCheck, Mid(P_ZERO_SUPPRESS, 1, intMaxLength))
        ' コンボボックス内にコードが存在してる？
        For intCount = 0 To cmbCheck.ListCount - 1
            ' コード値比較
            If cmbCheck.Column(0, intCount) = strValue Then
                blnFound = True
                Exit For
            End If
        Next
        
        If Not blnFound Then
            MsgBox "該当するコードがありません。", vbInformation, P_ERROR
            SendKeys "{ESC}"
            intRet = 2
            GoTo ChkComboInput_End
        End If
    End If

ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If

    fncChkComboInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkTextInput
'        機能             :テキストボックス入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTextInput(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean

    If Nz(txtCheck) = "" Then
        fncChkTextInput = True
        Exit Function
    End If

    If CmLenb(txtCheck) > intMaxLength Then
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        fncChkTextInput = False
    Else
        fncChkTextInput = True
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkSingleByte
'        機能             :半角文字入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkSingleByte(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean

    fncChkSingleByte = False

    If Nz(txtCheck) = "" Then
        fncChkSingleByte = True
        Exit Function
    End If

    ' 桁数チェック
    If Not fncChkTextInput(txtCheck, intMaxLength, Cancel) Then
        Exit Function
    End If

    ' 半角文字チェック
    If CmChkCher(Nz(txtCheck.VALUE), "A") Then
        fncChkSingleByte = True
    Else
        MsgBox "半角英数字で入力してください。", vbExclamation, P_KAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :小数入力項目の入力チェック
'        MODULE_ID        :fncChkTextInput2
'        Parameter        :ctlTxt      = 対象となるテキストボックス
'                          varMaxValue = 最大値
'                          intLength   = 長さ
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTextInput2(ByRef ctlTxt As TextBox, _
                                 ByVal varMaxValue As Variant, ByVal intLength As Integer, _
                                 Optional ByVal intKbn As Integer = 0) As Boolean

    fncChkTextInput2 = False

    If IsNull(ctlTxt) = True Then
        fncChkTextInput2 = True: Exit Function
    End If

    Select Case intKbn
        Case 0
            If (ctlTxt > varMaxValue Or _
                Len(Replace(Replace(ctlTxt, ".", ""), "-", "")) > intLength) Then
                Exit Function
            End If

        Case 1
            ' *** フォームのエラーイベントでのエラートラップのとき
            If (ctlTxt > varMaxValue Or _
                Len(Replace(Replace(ctlTxt.Text, ".", ""), "-", "")) > intLength) Then
                Exit Function
            End If
    End Select

    fncChkTextInput2 = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncNumericCheck
'        機能             :数値入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncNumericCheck(objTarget As Object, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean

    Dim strNumber       As String
    Dim lngCompare      As Long

    fncNumericCheck = False

    strNumber = Nz(objTarget, "")
    ' 空白なら入力チェックしない
    If strNumber = "" Then
        fncNumericCheck = True
        Exit Function
    End If

    lngCompare = Val(Mid(P_ALL_NINE, 1, intMaxLength))
    If Val(Nz(objTarget, 0)) > lngCompare Then
        ' 桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        Exit Function
    End If

    fncNumericCheck = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetKeiyaku_SU
'        機能             :契約数取得
'        IN               :割引番号
'                          コネクション
'        OUT              :契約数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncGetKeiyaku_SU(strNEBIKI_NO As String, objConnection As Object) As Integer

    Dim strSQL         As String
    Dim rsKEIYAKU      As Object
    Dim intKEIYAKU_CNT As Integer

    On Error GoTo Exception

    intKEIYAKU_CNT = 0

    ' 契約数取得
    strSQL = ""
    '2009/06/30 MOD <S>
    'strSQL = "SELECT COUNT(*) Cnt FROM DCRA_TRAN "
    'strSQL = strSQL & " WHERE DCRAT_DCNT_NO = '" & strNEBIKI_NO & "' "
    strSQL = "SELECT count(*) Cnt"
    strSQL = strSQL & " FROM (SELECT DCRAT_ACPTNO"
    strSQL = strSQL & "  FROM DCRA_TRAN"
    strSQL = strSQL & "  WHERE DCRAT_DCNT_NO = '" & strNEBIKI_NO & "'"
    strSQL = strSQL & "group by DCRAT_ACPTNO) AS A"
    '2009/06/30 MOD <E>
    Set rsKEIYAKU = MSZZ025.ADODB_Recordset(strSQL, objConnection)

    intKEIYAKU_CNT = CStr(Nz(rsKEIYAKU.Fields("Cnt"), 0))
    rsKEIYAKU.Close
    Set rsKEIYAKU = Nothing

    fncGetKeiyaku_SU = intKEIYAKU_CNT

    Exit Function

Exception:
    If Not rsKEIYAKU Is Nothing Then: rsKEIYAKU.Close: Set rsKEIYAKU = Nothing
    Call Err.Raise(Err.Number, "fncGetKeiyaku_su" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetSettei_SU
'        機能             :割引数取得
'        IN               :割引番号
'                          コネクション
'        OUT              :契約数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncGetSettei_SU(strNEBIKI_NO As String, objConnection As Object) As String

    Dim strSQL         As String
    Dim rsSETTEI       As Object
    Dim strSETTEI_CNT  As String

    On Error GoTo Exception

    fncGetSettei_SU = ""

    ' 有効割引数取得
    strSQL = ""
    strSQL = "SELECT COUNT(*) Cnt FROM RLDN_TRAN "
    strSQL = strSQL & " WHERE RLDN_TRAN.RLDNT_BUMOC  =  '" & pstrBumonCode & "' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_YCODE  =  '" & Trim(Val(Nz(txt_ヤード, 0))) & "' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_NO     =  '" & strNEBIKI_NO & "' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE =  '1' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_TO     >= '" & Format$(Now, "yyyymmdd") & "'"
    Set rsSETTEI = MSZZ025.ADODB_Recordset(strSQL, objConnection)
    strSETTEI_CNT = CStr(Nz(rsSETTEI.Fields("Cnt"), 0)) & ":"
    rsSETTEI.Close
    Set rsSETTEI = Nothing

    ' 無効割引数取得
    strSQL = ""
    strSQL = "SELECT COUNT(*) Cnt FROM RLDN_TRAN "
    strSQL = strSQL & " WHERE RLDN_TRAN.RLDNT_BUMOC  = '" & pstrBumonCode & "' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_YCODE  = '" & Trim(Val(Nz(txt_ヤード, 0))) & "' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_NO     = '" & strNEBIKI_NO & "' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_ENABLE = '0' "
    strSQL = strSQL & "   AND RLDN_TRAN.RLDNT_TO     < '" & Format$(Now, "yyyymmdd") & "'"
    Set rsSETTEI = MSZZ025.ADODB_Recordset(strSQL, objConnection)
    strSETTEI_CNT = strSETTEI_CNT & CStr(Nz(rsSETTEI.Fields("Cnt"), 0))
    rsSETTEI.Close
    Set rsSETTEI = Nothing

    fncGetSettei_SU = strSETTEI_CNT

    Exit Function

Exception:
    If Not rsSETTEI Is Nothing Then: rsSETTEI.Close: Set rsSETTEI = Nothing
    Call Err.Raise(Err.Number, "fncGetSettei_SU" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

