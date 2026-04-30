VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS770"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　精算依頼リスト印刷
'   プログラムＩＤ　：　FVS760
'   作　成　日　　　：  2005/08/10
'   作　成　者　　　：  T.SUZUKI
'   Ver            ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/11/11
'   修　正　者　　　：　T.SUZUKI
'   修　正　内　容　：  1.売掛トラン、回収トランのデータ抽出時の不具合を修正
'                       2.ワークテーブルの物理削除処理、作成処理を追加
'                       3.帳票のソート順を変更
'                       4.帳票の出力条件を変更
'   Ver             ：　0.1
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/12/02
'   修　正　者　　　：　T.SUZUKI
'   修　正　内　容　：  画面に帳票出力条件である｢解約区分｣、｢顧客コード｣を追加
'   Ver             ：　0.2
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/03/20
'   修　正　者　　　：　N.MIURA
'   修　正　内　容　：  ＥＸＣＥＬシート出力処理、契約金額を税込み金額とする。
'   Ver             ：　0.3
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/04/27
'   修　正　者　　　：　tajima
'   修　正　内　容　：  保証会社ならば預り金は０円とする。
'   Ver             ：　0.4
'**********************************************
'修正履歴
'   修　正　日　　　：　2010/01/28
'   修　正　者　　　：　M.HONDA
'   修　正　内　容　：  保証金取得方法を変更
'   Ver             ：　0.5
'**********************************************
'修正履歴
'   修　正　日　　　：　2012/10/09
'   修　正　者　　　：　M.HONDA
'   修　正　内　容　：  保会社を表示させるように修正
'   Ver             ：　0.6
'**********************************************
'修正履歴
'   修　正　日　　　：　2012/10/31
'   修　正　者　　　：　M.HONDA
'   修　正　内　容　：  パルマ包括保証会社の場合には過去の契約から
'                       保証金を取得する。
'   Ver             ：　0.7
'**********************************************
'修正履歴
'   修　正　日　　　：　2013/02/05
'   修　正　者　　　：　M.HONDA
'   修　正　内　容　：  保証金有のみを出力するOPボタンを追加
'   Ver             ：　0.8
'**********************************************
'修正履歴
'   修　正　日　　　：　2017/10/10
'   修　正　者　　　：　N.IMAI
'   修　正　内　容　：  預り金取得（パルマ包括限定）でデータが取得出来ない場合
'                       メッセージが表示されてしまう不具合に対応。
'   Ver             ：　0.9
'**********************************************
'修正履歴
'   修　正　日　　　：　2023/04/06
'   修　正　者　　　：　N.IMAI
'   修　正　内　容　：  保証金のカラム変更に対応(KEIYT_HOSYA→KEIYT_HKSYA)
'   Ver             ：　1.0
'**********************************************
'修正履歴
'   修　正　日　　　：　2024/12/20
'   修　正　者　　　：　N.IMAI
'   修　正　内　容　：  KAIS_TRANに契約日以降の過剰が存在する場合もあるので、検索条件から年月を外す
'   Ver             ：　1.1
'**********************************************

Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const pstrCstRPT_ID        As String = "RVS770"

Private Const pstrCstKAKUNIN       As String = "確認"
Private Const pstrCstERROR         As String = "エラー"

Private pobjDB                     As Database
Private objKASE_DB                 As Database      ' 加瀬DB
Private objKONT_DB                 As Database      ' コンテナDB
Private Const pintCstFirstPage     As Integer = 26
Private Const pintCstOtherPage     As Integer = 34
Private Const pintCstKAKUTEI       As Integer = -1  ' 請求済フラグ(確定)
Private Const pintCstMIKAKUTEI     As Integer = 0   ' 請求済フラグ(未確定)
Private Const pintCARG_AGRE        As Integer = 9   ' 契約状態(完了)

Private pstrKYDATE                 As String        ' 解約日付(YYYYMM)
Private pstrKYDATE_F               As String        ' 解約日付(From)
Private pstrKYDATE_T               As String        ' 解約日付(To)

' 2005/11/11 ADD T.SUZUKI Start
Private Const pstrCstWK_TABLE_NM   As String = "FVS770_W01"
' 2005/11/11 ADD T.SUZUKI End

' 2005/12/02 ADD T.SUZUKI Start
Private Const pstrCstMSGKetaOver   As String = "桁数オーバーしています。"
Private Const pstrCstMSGInputError As String = "を入力してください。"

Private pstrBumonCd                As String        ' 部門コード
Private pstrBumonNm                As String        ' 部門名
' 2005/12/02 ADD T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    ' 初期処理
    Call Frm_Init
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Title.BackColor = RGB(255, 0, 204)
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_KaiyakuYYYYMM_AfterUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuYYYYMM_AfterUpdate()

    ' 『YYYY/MM』の形に変換
    Me.txt_KaiyakuYYYYMM = Left(Check_Date(Me.txt_KaiyakuYYYYMM & "01"), 7)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_KaiyakuYYYYMM_BeforeUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuYYYYMM_BeforeUpdate(Cancel As Integer)

    Dim vDATE   As Variant
    Dim VDateL  As Variant

    On Error GoTo txt_KaiyakuYYYYMM_BeforeUpdate_Err

    vDATE = Nz(Me.txt_KaiyakuYYYYMM, "")

    ' 何も入力されていない時は、処理をしない
    If vDATE = "" Then
        GoTo txt_KaiyakuYYYYMM_BeforeUpdate_Exit
    End If

    VDateL = vDATE & "01"
    VDateL = Check_Date(VDateL)
    VDateL = CMCheck_Date(VDateL)

    ' 入力チェック
    Select Case CmBCheck_Date(VDateL, 10)
        Case 0

        Case 2
            ' 桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, pstrCstERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo txt_KaiyakuYYYYMM_BeforeUpdate_Exit

        Case 3, 9
            MsgBox "日付を正しく入力してください。", vbInformation, pstrCstERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo txt_KaiyakuYYYYMM_BeforeUpdate_Exit
    End Select

txt_KaiyakuYYYYMM_BeforeUpdate_Exit:
    Exit Sub

txt_KaiyakuYYYYMM_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo txt_KaiyakuYYYYMM_BeforeUpdate_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_KaiyakuYYYYMM_Change
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuYYYYMM_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_KaiyakuYYYYMM_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KaiyakuYYYYMM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_KaiyakuYYYYMM, KeyAscii, 7)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_AfterUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_F_AfterUpdate()
    Me.txt_Seisan_ReqDate_F = Check_Date(Me.txt_Seisan_ReqDate_F)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_BeforeUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_F_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Seisan_ReqDate_F, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_Change
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_F_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_GotFocus
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_F_GotFocus()
        Call Date_TextFormat(Me.txt_Seisan_ReqDate_F, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_LostFocus
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_F_LostFocus()
    Call Date_TextFormat(Me.txt_Seisan_ReqDate_F, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Seisan_ReqDate_F, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_T_AfterUpdate()
    Me.txt_Seisan_ReqDate_T = Check_Date(Me.txt_Seisan_ReqDate_T)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_T_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Seisan_ReqDate_T, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_T_Change
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_T_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_T_GotFocus
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_T_GotFocus()
        Call Date_TextFormat(Me.txt_Seisan_ReqDate_T, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_T_LostFocus()
        Call Date_TextFormat(Me.txt_Seisan_ReqDate_T, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_Seisan_ReqDate_F_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Seisan_ReqDate_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Seisan_ReqDate_T, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_BeforeUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_YCODE_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_YCODE_F, "")
    If varCode = "" Then
        Me.txt_YCODE_F = ""
        Exit Sub
    End If

    ' ヤード名取得
    strCARG = fncGetYardData(varCode)
    If strCARG = "" Then
        MsgBox "該当するヤードマスタがありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_YCODE_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_AfterUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_AfterUpdate()

    Dim varCode   As Variant
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_YCODE_F_AfterUpdate_Err

    varCode = Nz(Me.cmb_YCODE_F, "")
    strCARG = ""
    If varCode = "" Then Exit Sub

    If Trim(varCode) <> "" Then
        ' ヤード名取得
        strCARG = fncGetYardData(varCode)
    End If

    Me.cmb_YCODE_F = Format(Me.cmb_YCODE_F, "000000")
    Me.txt_YCODE_F = strCARG
    Me.cmb_YCODE_T.SetFocus  ' ヤードコード(TO)にフォーカスセット

    Exit Sub

cmb_YCODE_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_Change
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YCODE_F, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_AfterUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_AfterUpdate()

    Dim varCode   As Variant
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_YCODE_T_AfterUpdate_Err

    varCode = Nz(Me.cmb_YCODE_T, "")
    strCARG = ""
    If varCode = "" Then Exit Sub

    If Trim(varCode) <> "" Then
        strWhere = " YARD_CODE = " & varCode & ""
        strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
    End If

    Me.cmb_YCODE_T = Format(Me.cmb_YCODE_T, "000000")
    Me.txt_YCODE_T = strCARG


    Exit Sub

cmb_YCODE_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_BeforeUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_YCODE_T_BeforeUpdate_Err

    varCode = Nz(Me.cmb_YCODE_T, "")
    If varCode = "" Then
        Me.txt_YCODE_T = ""
        Exit Sub
    End If

    ' ヤード名取得
    strCARG = fncGetYardData(varCode)
    If strCARG = "" Then
        MsgBox "該当するヤードマスタがありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_YCODE_T_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
'    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_Change
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_KeyPress
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YCODE_T, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_F_AfterUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_F_AfterUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_F_AfterUpdate()

    Dim strCode   As String
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_KAIYAKU_F_AfterUpdate_Err

    strCARG = ""
    strCode = Nz(Me.cmb_KAIYAKU_F, "")
    If strCode = "" Then Exit Sub

    If Trim(strCode) <> "" Then
        ' 解約区分名取得
        strCARG = fncGetName(0, "089", strCode)  ' 2005/12/02 ADD T.SUZUKI
'        strCARG = fncGETNAME("089", strCode)    ' 2005/12/02 DEL T.SUZUKI
    End If

    Me.cmb_KAIYAKU_F = Format(Me.cmb_KAIYAKU_F, "00")
    Me.txt_KAIYAKU_F = strCARG  ' 解約区分名をセット
    Me.cmb_KAIYAKU_T.SetFocus   ' 解約区分(TO)にフォーカスセット

    Exit Sub

cmb_KAIYAKU_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_F_BeforeUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_F_BeforeUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_F_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_KAIYAKU_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_KAIYAKU_F, "")
    If varCode = "" Then
        Me.txt_KAIYAKU_F = ""
        Exit Sub
    End If

    ' 解約区分名取得
    strCARG = fncGetName(0, "089", varCode)
'    strCARG = fncGETNAME("089", varCode)
    If strCARG = "" Then
        MsgBox "該当する解約区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_KAIYAKU_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_F_Changeイベント
'        MODULE_ID        :cmb_KAIYAKU_F_Change
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_F_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_F_KeyPressイベント
'        MODULE_ID        :cmb_KAIYAKU_F_KeyPress
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_KAIYAKU_F, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_T_AfterUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_T_AfterUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_T_AfterUpdate()

    Dim strCode   As String
    Dim strCARG   As String

    On Error GoTo cmb_KAIYAKU_T_AfterUpdate_Err

    strCARG = ""
    strCode = Nz(Me.cmb_KAIYAKU_T, "")
    If strCode = "" Then Exit Sub

    If Trim(strCode) <> "" Then
        ' 解約区分名取得
        strCARG = fncGetName(0, "089", strCode)  ' 2005/12/02 ADD T.SUZUKI
'        strCARG = fncGETNAME("089", strCode)    ' 2005/12/02 DEL T.SUZUKI
    End If

    Me.cmb_KAIYAKU_T = Format(Me.cmb_KAIYAKU_T, "00")
    Me.txt_KAIYAKU_T = strCARG  ' 解約区分名をセット
    Me.txt_顧客カナ.SetFocus     ' 顧客名カナにフォーカスセット

    Exit Sub

cmb_KAIYAKU_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_T_BeforeUpdateイベント
'        MODULE_ID        :cmb_KAIYAKU_T_BeforeUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_T_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_KAIYAKU_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_KAIYAKU_T, "")
    If varCode = "" Then
        Me.txt_KAIYAKU_T = ""
        Exit Sub
    End If

    ' 解約区分名取得
    strCARG = fncGetName(0, "089", varCode)  ' 2005/12/02 ADD T.SUZUKI
'    strCARG = fncGETNAME("089", varCode)    ' 2005/12/02 DEL T.SUZUKI
    If strCARG = "" Then
        MsgBox "該当する解約区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_KAIYAKU_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_T_Changeイベント
'        MODULE_ID        :cmb_KAIYAKU_T_Change
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_T_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_KAIYAKU_T_KeyPressイベント
'        MODULE_ID        :cmb_KAIYAKU_T_KeyPress
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAIYAKU_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_KAIYAKU_T, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客カナ_AfterUpdateイベント
'        MODULE_ID        :txt_顧客カナ_AfterUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客カナ_AfterUpdate()

    ' 顧客名が空白の場合、顧客コードコンボボックス、顧客名を空白に設定
    If Nz(Me.txt_顧客カナ, "") = "" Then
        Me.cmb_顧客コード.RowSource = ""
        Me.cmb_顧客コード = Null
        Me.txt_顧客名 = Null
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客カナ_BeforeUpdateイベント
'        MODULE_ID        :txt_顧客カナ_BeforeUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客カナ_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    vCode = Nz(txt_顧客カナ, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox pstrCstMSGKetaOver, vbInformation, pstrCstERROR
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "顧客名カナ" & pstrCstMSGInputError, vbInformation, pstrCstERROR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客カナ_Changeイベント
'        MODULE_ID        :txt_顧客カナ_Change
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客カナ_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_AfterUpdateイベント
'        MODULE_ID        :cmb_顧客コード_AfterUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_AfterUpdate()

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_顧客コード, "")

    If vCode = "" Then
        Exit Sub
    End If

    'ファイルへの存在チェック
    If fncChkUSER_MAST(1) = False Then
        Exit Sub
    End If

    Me.cmb_顧客コード = Format(Me.cmb_顧客コード, "000000")
    Me.txt_顧客カナ.SetFocus
    Me.cmb_顧客コード.SetFocus

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_BeforeUpdateイベント
'        MODULE_ID        :cmb_顧客コード_BeforeUpdate
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_顧客コード, "")

    If vCode = "" Then
        Me.txt_顧客名 = ""  ' 顧客名を空白に設定
        Exit Sub
    End If

    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox pstrCstMSGKetaOver, vbInformation, pstrCstERROR
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "コードを入力してください。", vbInformation, pstrCstERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

    'ファイルへの存在チェック
    If fncChkUSER_MAST(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_GotFocusイベント
'        MODULE_ID        :cmb_顧客コード_GotFocus
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_GotFocus()

    Dim strSQL As String

    ' 顧客名が空白の場合、顧客コードコンボボックス、顧客名を空白に設定
    If Nz(Me.txt_顧客カナ, "") = "" And Nz(Me.txt_顧客名, "") = "" Then  ' 2005/12/02 ADD T.SUZUKI
'    If Nz(Me.txt_顧客カナ, "") = "" Then                                ' 2005/12/02 DEL T.SUZUKI
        Me.cmb_顧客コード.RowSource = ""
        Me.cmb_顧客コード = Null
        Me.txt_顧客名 = Null
        Exit Sub

    Else
        ' 顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ表示のSQL文作成
        strSQL = "SELECT TOP 100 Format([USER_CODE],'000000') AS 式1, "
        strSQL = strSQL & " dbo_USER_MAST.USER_NAME "
        strSQL = strSQL & "  FROM dbo_USER_MAST "
        strSQL = strSQL & " WHERE (((dbo_USER_MAST.USER_KANA) LIKE '*" & Me.txt_顧客カナ & "*')) "
        strSQL = strSQL & " ORDER BY Format([USER_CODE],'000000');"
    
        Me.cmb_顧客コード.RowSource = strSQL
        Me.cmb_顧客コード.Requery

        If Me.cmb_顧客コード.ListCount = 0 Then
            Me.cmb_顧客コード = Null
            Me.txt_顧客名 = Null
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_Changeイベント
'        MODULE_ID        :cmb_顧客コード_Change
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_Change()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_KeyPressイベント
'        MODULE_ID        :cmb_顧客コード_KeyPress
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_顧客コード, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Seisan_Click
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Seisan_Click()

    Dim intRtn    As Integer
    Dim strMsg    As String
    Dim blnUpdFlg As Boolean

    doCmd.Hourglass True
    blnUpdFlg = False

    ' 確認メッセージ
    strMsg = "精算依頼済みに更新しますか？"
    intRtn = MsgBox(strMsg, vbQuestion + vbYesNo, pstrCstKAKUNIN)
    If intRtn = vbNo Then GoTo cmd_Seisan_Click_Exit

    ' コンテナ契約ファイル更新処理
    If fncUPD_CARG_FILE = False Then
        GoTo cmd_Seisan_Click_Exit
    End If

    ' 終了メッセージ
    intRtn = MsgBox("更新処理が終了しました。", vbInformation + vbOKOnly, pstrCstKAKUNIN)
    blnUpdFlg = True

cmd_Seisan_Click_Exit:
    doCmd.Hourglass False

    Me.txt_KaiyakuYYYYMM.SetFocus

    If blnUpdFlg = True Then
        Me.cmd_Seisan.Enabled = False
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :chk_Seisan_Req_AfterUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_Seisan_Req_AfterUpdate()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :chk_Seisan_Req_AfterUpdate
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_選択_AfterUpdate()
    ' 精算依頼済みにするコマンドボタンの制御
    Call subControlcmdSeisan
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Preview_Click
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Preview_Click()

    On Error GoTo cmd_Preview_Click_Err

    Dim strMsg       As String
    Dim intRtn       As Long
    Dim lngDataCount As Long     ' データ数
    Dim intPageCount As Integer  ' ページ数

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Preview_Click_Exit

    strMsg = "印刷内容を画面に出力します。よろしいですか？"
    intRtn = MsgBox(strMsg, vbQuestion + vbYesNo, pstrCstKAKUNIN)
    If intRtn = vbNo Then GoTo cmd_Preview_Click_Exit

    doCmd.Hourglass True

'''' ★★★★★デバック用★★★★★★★★★★★★★★★★★★★★
'''    Debug.Print "帳票プレビューStart"
'''    Debug.Print Now
'''' ★★★★★デバック用★★★★★★★★★★★★★★★★★★★★

    ' 帳票出力用データ作成
    If FncCreateWork(lngDataCount, intPageCount) = False Then GoTo cmd_Preview_Click_Exit
'''' ★★★★★デバック用★★★★★★★★★★★★★★★★★★★★
'''    MsgBox "該当件数：" & Format(lngDataCount, "###,##0") & "件" & Space(1) & "ページ数：" & intPageCount
'''' ★★★★★デバック用★★★★★★★★★★★★★★★★★★★★

    ' 対象件数セット
    Me.txt_Count = lngDataCount

    ' レポートオープン(ページ数を渡す)
    doCmd.OpenReport "RVS770", acPreview, "", "", acWindowNormal, CStr(intPageCount)
'''    doCmd.OpenReport "RVS770", acPreview, "", ""

    ' 精算依頼済みにするコマンドボタンを使用可能にする
    If Me.chk_Seisan_Req.VALUE = 0 Then
        ' 『精算依頼済みを対象』がOffのとき
        Me.cmd_Seisan.Enabled = True
    End If

'''' ★★★★★デバック用★★★★★★★★★★★★★★★★★★★★
'''    Debug.Print "帳票プレビューEnd"
'''    Debug.Print Now
'''' ★★★★★デバック用★★★★★★★★★★★★★★★★★★★★

cmd_Preview_Click_Exit:
    doCmd.Hourglass False
    Exit Sub

cmd_Preview_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo cmd_Preview_Click_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Excel_Click
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    Dim lngDataCount As Long     ' データ数
    Dim intPageCount As Integer  ' ページ数

    On Error GoTo cmd_Excel_Click_Err

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Excel_Click_Exit

    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, pstrCstKAKUNIN) = vbYes) Then

        doCmd.Hourglass True

        ' 帳票出力用データ作成
        If FncCreateWork(lngDataCount, intPageCount) = True Then
           doCmd.OutputTo acQuery, "FVS770_Q01", "MicrosoftExcel(*.xls)", "", True, ""
        End If
    End If

cmd_Excel_Click_Exit:
    doCmd.Hourglass False
    Exit Sub

cmd_Excel_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo cmd_Excel_Click_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_End_Click
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Frm_Init
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Frm_Init()

    With Me
        .txt_KaiyakuYYYYMM = Null       ' 対象解約年月
        .txt_Seisan_ReqDate_F = Null    ' 精算依頼日(From)
        .txt_Seisan_ReqDate_T = Null    ' 精算依頼日(To)
        .cmb_YCODE_F = Null             ' ヤードコード(From)
        .txt_YCODE_F = Null             ' ヤード名(From)
        .cmb_YCODE_T = Null             ' ヤードコード(To)
        .txt_YCODE_T = Null             ' ヤード名(To)
        .txt_Count = Null
        .cmd_Seisan.Enabled = False     ' 精算依頼済みにするコマンドボタン
        .chk_Seisan_Req.VALUE = 0       ' 精算依頼済みを対象にするチェックボックスOff
        .chk_Parma.VALUE = 0
        .opg_選択.VALUE = 0             ' 全てをOn

        .cmb_顧客コード.RowSource = ""  ' 顧客コード  ' 2005/12/02 ADD T.SUZUKI
    End With

    ' 部門名取得
    Call subGetBumonName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subControlcmdSeisan
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subControlcmdSeisan()

    On Error Resume Next

    If Me.cmd_Seisan.Enabled = True Then
        Me.cmd_Seisan.Enabled = False
        Me.txt_Count = Null
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub dateCheck(ByRef ctlTxt As TextBox, Cancel As Integer)

    Dim vCode As Variant
    Dim vDATE As Variant

    On Error GoTo DateCheck_Err

    vCode = Nz(ctlTxt, "")

    ' 何も入力されていない時は、処理をしない
    If vCode = "" Then GoTo DateCheck_Exit
    vCode = Check_Date(vCode)
    vCode = CMCheck_Date(vCode)

    ' 入力チェック
    Select Case CmBCheck_Date(vCode)
        Case 2
             ' 桁数エラー
             MsgBox "桁数オーバーしています。", vbExclamation, pstrCstERROR
             Cancel = True
             SendKeys "{ESC}"
             GoTo DateCheck_Exit

        Case 3, 9
             MsgBox "日付を正しく入力してください。", vbExclamation, pstrCstERROR
             Cancel = True
             SendKeys "{ESC}"
             GoTo DateCheck_Exit
        Case Else
    End Select

DateCheck_Exit:
    Exit Sub

DateCheck_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Date_TextFormat
'        機能             :日付データのフォーマット
'        IN               :
'        OUT              :
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Date_TextFormat(ByRef ctlTxt As TextBox, ByVal lngMode As Long)

    Dim vCode As Variant

    vCode = Nz(ctlTxt, "")
    If vCode <> "" Then
        Select Case lngMode
            Case 0
                ctlTxt = Format(vCode, "yy/mm/dd")
            Case 1
                ctlTxt = Format(vCode, "yyyy/mm/dd")
            Case 2
                ctlTxt = Format(vCode, "yy/mm")
            Case 3
                ctlTxt = Format(vCode, "yyyy/mm")
        End Select
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)

    Dim sYear As String

    CMCheck_Date = Null

    If Nz(vDATE, "") = "" Then Exit Function
    If IsDate(vDATE) = False Then Exit Function

    sYear = Format(vDATE, "yyyy")
    If CmLenb(sYear) <> 4 Then Exit Function

    CMCheck_Date = vDATE
End Function

'==============================================================================*
'
'        MODULE_NAME      :日付項目のチェックを一括でする
'        MODULE_ID        :
'        CREATE_DATE      :
'
'引数
'oText 対象項目
'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
'返り値は、
'２：桁数が溢れている時
'３：日付と判断できない時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer

    Dim sYear   As String
    Dim intLen  As Integer

    On Error GoTo CmBCheck_Date_Err

    CmBCheck_Date = 2

    If IsMissing(ilen) = False Then
       intLen = Nz(CmNumeric(ilen), 0)
    Else
       intLen = 10
    End If

    oText = Nz(oText, "")
    ' 桁数チェック
    If CmLenb(oText) > intLen Then GoTo CmBCheck_Date_Exit
    CmBCheck_Date = 3
    If Nz(oText, "") = "" Then GoTo CmBCheck_Date_Exit

    ' 論理チェック
    If IsDate(oText) = False Then GoTo CmBCheck_Date_Exit

    sYear = Format(oText, "yyyy")
    If CmLenb(sYear) <> 4 Then GoTo CmBCheck_Date_Exit

    CmBCheck_Date = 0

CmBCheck_Date_Exit:
    Exit Function

CmBCheck_Date_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    CmBCheck_Date = 9
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetYardData
'        機能             :ヤードマスタよりヤード名取得
'        IN               :
'        OUT              :
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetYardData(ByVal varCode As Variant) As String

    Dim strWhere As String
    Dim strCARG  As String

    On Error GoTo fncGetYardData_Err

    fncGetYardData = ""

    strWhere = strWhere & " dbo_YARD_MAST.YARD_CODE = " & varCode & ""
    strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")

    fncGetYardData = strCARG

    Exit Function

fncGetYardData_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    fncInputCheck = False

    With Me
        If Nz(.txt_KaiyakuYYYYMM, "") = "" Then
            MsgBox "対象解約年月を入力して下さい。", vbInformation, pstrCstERROR
            .txt_KaiyakuYYYYMM.SetFocus
            Exit Function
        End If
    End With

    fncInputCheck = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :日付入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkDateFromTo
'        IN               :
'        CREATE_DATE      :2005/08/10
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkDateFromTo(ByRef ctlTextF As TextBox, ByRef ctlTextT As TextBox)

    Dim strDayF As String
    Dim strDayT As String
    Dim strSwap As String

    If ctlTextF <> "" And ctlTextT <> "" Then
        strDayF = Format(ctlTextF, "yyyy/mm/dd")  ' 精算依頼日(From)
        strDayT = Format(ctlTextT, "yyyy/mm/dd")  ' 精算依頼日(To)

        If strDayT <> "" And strDayF > strDayT Then
            ' 大小入力エラー時
            strSwap = strDayF
            strDayF = strDayT
            strDayT = strSwap

            ' 参照元へ再セット
            ctlTextF = strDayF
            ctlTextT = strDayT
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ヤードコード入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkYARDFromTo
'        IN               :
'        CREATE_DATE      :2005/08/10
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkYARDFromTo(ByRef ctlcmbF As ComboBox, ByRef ctlTextF As TextBox, _
                          ByRef ctlcmbT As ComboBox, ByRef ctlTextT As TextBox)

    Dim strYARDF       As String
    Dim strYARDT       As String
    Dim strSwap        As String

    strSwap = ""
    If Trim(ctlcmbF) <> "" And Trim(ctlcmbT) <> "" Then
        strYARDF = Format(ctlcmbF, "000000")  ' ヤードコード(From)
        strYARDT = Format(ctlcmbT, "000000")  ' ヤードコード(To)

        If strYARDF > strYARDT Then
            ' 大小入力エラー時
            ' *** ヤードコード
            strSwap = strYARDF
            strYARDF = strYARDT
            strYARDT = strSwap

            '　参照元へ再セット
            ctlcmbF = strYARDF
            ctlcmbT = strYARDT

            ' 名称
            ' *** ヤード名
            strSwap = ""
            strYARDF = ""
            strYARDT = ""

            strSwap = ctlTextF
            strYARDF = ctlTextT
            strYARDT = strSwap

            '　参照元へ再セット
            ctlTextF = strYARDF
            ctlTextT = strYARDT
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル削除
'        MODULE_ID        :fncDeleteWorkTable
'        Parameter        :ワークテーブル名
'        CREATE_DATE      :2005/11/11
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDeleteWorkTable(ByVal strTableNm As String) As Boolean

    Dim strSQL As String

    On Error GoTo fncDeleteWorkTable_Err

    fncDeleteWorkTable = False

    ' テーブルの存在チェック
    If fncChkTableExist(strTableNm) = True Then

        ' テーブル削除
        pobjDB.TableDefs.Delete strTableNm
        Select Case Err.Number
            Case 0
            Case 3265
                Err.Clear
            Case Else
                MsgBox "subDropWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
                Err.Clear
        End Select
    End If

    fncDeleteWorkTable = True

    Exit Function

fncDeleteWorkTable_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの存在チェック
'        MODULE_ID        :fncChkTableExist
'        Parameter        :ワークテーブル名
'        CREATE_DATE      :2005/11/11
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTableExist(ByVal strTableNm As String) As Boolean

    Dim tdf As TableDef

    fncChkTableExist = False

    For Each tdf In pobjDB.TableDefs
        If tdf.NAME = strTableNm Then
            fncChkTableExist = True
            Exit For
        End If
    Next tdf
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル作成
'        MODULE_ID        :fncCreateWorkTable
'        Parameter        :ワークテーブル名
'        CREATE_DATE      :2005/11/11
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable(ByVal strTableNm As String) As Boolean

    Dim tdfNew    As TableDef
    Dim intCount  As Integer

    On Error GoTo fncCreateWorkTable_Err

    fncCreateWorkTable = False

    Set tdfNew = pobjDB.CreateTableDef(pstrCstWK_TABLE_NM)
    With tdfNew
        .Fields.Append .CreateField("FURIKOMIKIN_FLG", dbLong)        ' 振込金額フラグ(0 OR 1 OR 2)
        .Fields.Append .CreateField("SortKey", dbLong)                ' ページフッター制御用
        .Fields.Append .CreateField("REPORT_ID", dbText, 50)          ' レポートID
        .Fields.Append .CreateField("BUMOM_BUMOC", dbText, 3)         ' 部門コード
        .Fields.Append .CreateField("BUMOM_BUMON", dbText, 50)        ' 部門名称
        .Fields.Append .CreateField("CARG_LIQ_REQDATE", dbDate)       ' 精算依頼日
        .Fields.Append .CreateField("CARG_UCODE", dbLong)             ' 顧客コード
        .Fields.Append .CreateField("USER_NAME", dbText, 36)          ' 顧客名
        .Fields.Append .CreateField("USER_SKBN", dbInteger)           ' 集金区分
        .Fields.Append .CreateField("USER_SKBN_NAME", dbText, 20)     ' 集金区分名
        .Fields.Append .CreateField("CARG_YCODE", dbLong)             ' ヤードコード
        .Fields.Append .CreateField("CARG_NO", dbLong)                ' コンテナ番号
        .Fields.Append .CreateField("CARG_KYDATE", dbDate)            ' 解約日付
        .Fields.Append .CreateField("KAIYAKU_NAME", dbText, 20)       ' 解約区分名(略称)  ' 2005/12/02 ADD T.SUZUKI
        .Fields.Append .CreateField("CARG_RENTKG", dbLong)            ' 月額料金
        .Fields.Append .CreateField("CARG_LIQ_NOTE", dbText, 60)      ' 精算備考
        .Fields.Append .CreateField("KEIYAKU_HONSU", dbLong)          ' 他契約本数
        .Fields.Append .CreateField("CARG_HOSYICD", dbText, 2)        ' 保証区分
        .Fields.Append .CreateField("CARG_HOSYICD_NAME", dbText, 20)  ' 保証区分名
        .Fields.Append .CreateField("CARG_HOSYO_CD", dbText, 6)        ' 保証会社ｺｰﾄﾞ  ' 2012/10/09 M.HDONDA INS
        .Fields.Append .CreateField("CARG_HOSYO_CD_NAME", dbText, 20)  ' 保証会社名称  ' 2012/10/09 M.HDONDA INS
        .Fields.Append .CreateField("KEIYT_HOZUA", dbLong)            ' 預り金
        .Fields.Append .CreateField("NYUUKIN_ZAN", dbLong)            ' 入金残
        .Fields.Append .CreateField("FURIKOMI_KIN", dbLong)           ' 振込金額
        .Fields.Append .CreateField("CARG_ACPTNO", dbText, 10)        ' 契約受注番号

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With

    ' テーブルの追加
    pobjDB.TableDefs.Append tdfNew

    fncCreateWorkTable = True

    Exit Function

fncCreateWorkTable_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Function

' 2005/11/11 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :ワークテーブルのデータ削除
''        MODULE_ID        :fncDeleteWork
''        CREATE_DATE      :2005/08/10
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncDeleteWork() As Boolean
'
'    Dim strSQL As String
'
'    On Error GoTo fncDeleteWork_Err
'
'    fncDeleteWork = False
'
'    ' ワークテーブルの内容を削除
'    strSQL = "DELETE FROM FVS770_W01; "
'    pobjDB.Execute strSQL, dbFailOnError
'
'    fncDeleteWork = True
'
'    Exit Function
'
'fncDeleteWork_Err:
'    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
'    Err.Clear
'End Function
' 2005/11/11 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルへ対象データのInsert
'        MODULE_ID        :fncCreateWork
'        Parameter        :lngDataCount データ数
'                          intPageCount ページ数
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function FncCreateWork(ByRef lngDataCount As Long, ByRef intPageCount As Integer) As Boolean

    Dim strSQL          As String     ' SELECT句格納
    Dim strWhere        As String     ' Where句格納

    Dim objRsKONT       As Recordset  ' メインレコードセット(コンテナDB)
    Dim objWk           As Recordset  ' ワークテーブルレコードセット(FVS770_W01)

'    Dim strBumoCD       As String     ' 部門コード  ' 2005/12/02 DEL T.SUZUKI
'    Dim strBumoNm       As String     ' 部門名      ' 2005/12/02 DEL T.SUZUKI

    Dim strSaveKokyc    As String     ' 保存用顧客コード

    Dim lngURI_TOTAL    As Long       ' 売上金額合計
    Dim lngHOSYA_TOTAL  As Long       ' 預保証金合計
    Dim lngKAIS_TOTAL   As Long       ' 回収金額合計
    Dim lngNYUKIN_ZAN   As Long       ' 入金残

    Dim blnChgFlg       As Boolean    ' 顧客コード変更フラグ
    Dim blnInsFlg       As Boolean    ' データ作成フラグ

    Dim intRtn          As Integer    ' 2005/11/11 ADD T.SUZUKI

    On Error GoTo fncCreateWork_Err

    FncCreateWork = False

    doCmd.SetWarnings False

    ' DB接続
    Set pobjDB = CurrentDb

    ' ワークテーブルの削除処理
    If fncDeleteWorkTable(pstrCstWK_TABLE_NM) = False Then GoTo fncCreateWork_Err  ' 2005/11/11 ADD T.SUZUKI
'    If fncDeleteWork = False Then GoTo fncCreateWork_Err                          ' 2005/11/11 DEL T.SUZUKI

    ' ワークテーブルの作成処理
    If fncCreateWorkTable(pstrCstWK_TABLE_NM) = False Then GoTo fncCreateWork_Err  ' 2005/11/11 ADD T.SUZUKI

'    ' 部門名取得
'    Call subGetBumonName(pobjDB, strBumoCD, strBumoNm)

    ' コンテナDB、加瀬DB接続
    Call subConnectDB(pstrBumonCd)
'    Call subConnectDB(strBumoCD)

    ' コンテナ契約ファイルよりデータ取得
    strSQL = ""
    strSQL = strSQL & "SELECT " & Chr(13)
    strSQL = strSQL & " CARG_LIQ_REQDATE, " & Chr(13)  ' 精算依頼日
    strSQL = strSQL & " CARG_UCODE, " & Chr(13)        ' 顧客コード
    strSQL = strSQL & " USER_NAME, " & Chr(13)         ' 顧客名
    strSQL = strSQL & " USER_SKBN, " & Chr(13)         ' 集金区分
    strSQL = strSQL & " CARG_YCODE, " & Chr(13)        ' ヤードコード
    strSQL = strSQL & " CARG_NO, " & Chr(13)           ' コンテナ番号
    strSQL = strSQL & " CARG_KYDATE, " & Chr(13)       ' 解約日付
    strSQL = strSQL & " CARG_KAICD, " & Chr(13)        ' 解約区分  ' 2005/12/02 ADD T.SUZUKI
    strSQL = strSQL & " CARG_RENTKG, " & Chr(13)       ' 月額料金
    strSQL = strSQL & " CARG_SYOZEI, " & Chr(13)       ' 消費税    ' 2007/03/20 ADD T.MIURA
    strSQL = strSQL & " CARG_LIQ_NOTE, " & Chr(13)     ' 精算備考
    strSQL = strSQL & " CARG_HOSYICD, " & Chr(13)      ' 保証区分
    strSQL = strSQL & " CARG_HOSYO_CD, " & Chr(13)     ' 保証会社ｺｰﾄﾞ  '2012/10/09 M.HONDA INS
    strSQL = strSQL & " SHIRM_SHIRN, " & Chr(13)       ' 保証会社名称  '2012/10/09 M.HONDA INS
    strSQL = strSQL & " CARG_ACPTNO " & Chr(13)        ' 受注契約番号
    strSQL = strSQL & " FROM CARG_FILE " & Chr(13)  ' **コンテナ契約ファイル
    strSQL = strSQL & " LEFT JOIN KASE_DB.dbo.SHIR_MAST ON "   '2012/10/09 M.HONDA INS
    strSQL = strSQL & " SHIRM_BUMOC = 'A' AND "                '2012/10/09 M.HONDA INS
    strSQL = strSQL & " SHIRM_SHIRC = CARG_HOSYO_CD, "         '2012/10/09 M.HONDA INS
    strSQL = strSQL & "      YARD_MAST, " & Chr(13)  ' **ヤードマスタ
    strSQL = strSQL & "      USER_MAST " & Chr(13)   ' **顧客マスタ
    strSQL = strSQL & " WHERE CARG_YCODE = YARD_CODE " & Chr(13)
    strSQL = strSQL & "   AND CARG_UCODE = USER_CODE " & Chr(13)

    ' Where句生成
    strWhere = fncCreateWhere

    If strWhere <> "" Then
        strSQL = strSQL & strWhere
    End If

    ' ソート
    strSQL = strSQL & " ORDER BY CARG_UCODE, " & Chr(13)       ' 顧客コード
    strSQL = strSQL & "          CARG_YCODE DESC, " & Chr(13)  ' ヤードコード
    strSQL = strSQL & "          CARG_NO DESC" & Chr(13)       ' コンテナ番号

    ' コンテナ契約ファイル読込
    Set objRsKONT = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    If objRsKONT.RecordCount = 0 Then
        GoTo No_Data_Exit
    End If

    ' ワークテーブルへ書き込み
    strSQL = "SELECT * FROM FVS770_W01 "
    Set objWk = pobjDB.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

    strSaveKokyc = ""
    blnChgFlg = False
    blnInsFlg = False
    lngDataCount = 0
    Do Until objRsKONT.EOF
        ' *** 顧客コード毎に出力 ***
        If strSaveKokyc <> objRsKONT.Fields("CARG_UCODE") Then
            blnInsFlg = False    ' Insertフラグ
            lngNYUKIN_ZAN = 0    ' 入金残
            lngURI_TOTAL = 0     ' 売上金額合計
            lngKAIS_TOTAL = 0    ' 回収金額合計

            blnChgFlg = True  ' 顧客コード変更フラグ

            ' 売上金額合計・預保証金・回収金額合計取得
            ' *** 売上金額合計・預保証金取得 ***
            lngURI_TOTAL = fncGetURI_TOTAL(pstrBumonCd, Format(objRsKONT.Fields("CARG_UCODE"), "000000"), pstrKYDATE)    ' 205/12/02 ADD T.SUZUKI
'            lngURI_TOTAL = fncGetURI_TOTAL(strBumoCD, Format(objRsKONT.Fields("CARG_UCODE"), "000000"), pstrKYDATE)    ' 205/11/11 ADD T.SUZUKI  ' 205/12/02 DEL T.SUZUKI
'            lngURI_TOTAL = fncGetURI_TOTAL(strBumoCD, objRsKONT.Fields("CARG_UCODE"), pstrKYDATE)                     ' 205/11/11 DEL T.SUZUKI

            ' *** 回収金額合計取得 ***
            lngKAIS_TOTAL = fncGetKAIS_TOTAL(pstrBumonCd, Format(objRsKONT.Fields("CARG_UCODE"), "000000"), pstrKYDATE)  ' 205/12/02 ADD T.SUZUKI
'            lngKAIS_TOTAL = fncGetKAIS_TOTAL(strBumoCD, Format(objRsKONT.Fields("CARG_UCODE"), "000000"), pstrKYDATE)  ' 205/11/11 ADD T.SUZUKI  ' 205/12/02 DEL T.SUZUKI
'            lngKAIS_TOTAL = fncGetKAIS_TOTAL(strBumoCD, objRsKONT.Fields("CARG_UCODE"), pstrKYDATE)                   ' 205/11/11 DEL T.SUZUKI

            ' 入金残算出(売上金額合計 - 回収金額合計)
            lngNYUKIN_ZAN = lngURI_TOTAL - lngKAIS_TOTAL

' 2005/11/11 DEL T.SUZUKI Start
'            ' 処理選択
'            Select Case Me.opg_選択
'                Case 0  ' *** 全て ***
'                    blnInsFlg = True
'                Case 1  ' *** 未収 ***
'                    If lngNYUKIN_ZAN > 0 Then
'                        blnInsFlg = True
'                    End If
'                Case 2  ' *** 過入金 ***
'                    If lngNYUKIN_ZAN < 0 Then
'                        blnInsFlg = True
'                    End If
'                Case 3  ' *** 精算なし ***
'                    If lngNYUKIN_ZAN = 0 Then
'                        blnInsFlg = True
'                    End If
'            End Select
' 2005/11/11 DEL T.SUZUKI End

            ' 保存用顧客コード格納
            strSaveKokyc = objRsKONT.Fields("CARG_UCODE")
        End If
        ' **************************

        ' データ作成処理
' 2005/11/11 DEL T.SUZUKI Start
        If fncInsertFVS770_W01(objWk, objRsKONT, blnChgFlg, pstrBumonCd, pstrBumonNm, lngNYUKIN_ZAN) = False Then GoTo fncCreateWork_Exit  ' 2005/12/02 ADD T.SUZUKI
'        If fncInsertFVS770_W01(objWk, objRsKONT, blnChgFlg, strBumoCD, strBumoNm, lngNYUKIN_ZAN) = False Then GoTo fncCreateWork_Exit     ' 2005/12/02 DEL T.SUZUKI
        lngDataCount = lngDataCount + 1  ' データ数カウントアップ
' 2005/11/11 DEL T.SUZUKI End
' 2005/11/11 DEL T.SUZUKI Start
'        If blnInsFlg = True Then
'            ' データ作成対象の場合
'            If fncInsertFVS770_W01(objWk, objRsKONT, blnChgFlg, strBumoCD, strBumoNm, lngNYUKIN_ZAN) = False Then GoTo fncCreateWork_Exit '
'            lngDataCount = lngDataCount + 1  ' データ数カウントアップ
'        End If
' 2005/11/11 DEL T.SUZUKI End

        ' 顧客コード変更フラグ初期化
        If blnChgFlg = True Then blnChgFlg = False

        objRsKONT.MoveNext
    Loop

    ' 該当データ数0件のとき
    If lngDataCount = 0 Then GoTo No_Data_Exit

' 205/11/11 MOVE T.SUZUKI Start
    ' 振込金額設定
    intRtn = fncUPDATE_FURIKOMIKIN
    Select Case intRtn
        Case 0   ' *** 該当データ数0件 ***
            lngDataCount = 0
            GoTo No_Data_Exit

        Case -1  ' *** システムエラー ***
            lngDataCount = 0
            GoTo fncCreateWork_Exit

        Case Else  ' *** 該当データあり ***
            ' 該当データ数セット
            lngDataCount = intRtn
    End Select
' 205/11/11 MOVE T.SUZUKI End

    ' ソートキー設定
    If fncSetPageSortKey(intPageCount) = False Then GoTo fncCreateWork_Err

' 205/11/11 MOVE T.SUZUKI Start
'    ' 振込金額設定
'    If fncUPDATE_FURIKOMIKIN() = False Then GoTo fncCreateWork_Err
' 205/11/11 MOVE T.SUZUKI End

    FncCreateWork = True

fncCreateWork_Exit:
    If Not objRsKONT Is Nothing Then objRsKONT.Close: Set objRsKONT = Nothing
    If Not objWk Is Nothing Then objWk.Close: Set objWk = Nothing

    ' DB切断
    Call subDisConnectDB

    doCmd.SetWarnings True
    Exit Function

No_Data_Exit:
    MsgBox "該当するデータがありません。", vbInformation, pstrCstERROR
    GoTo fncCreateWork_Exit
    Exit Function

fncCreateWork_Err:
'''    objWs.Rollback
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncCreateWork_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :データ登録処理
'        MODULE_ID        :fncInsertFVS770_W01
'        Parameter        :objWk          ：カレントDBレコードセット
'                         :objRsKONT      ：コンテナDBレコードセット
'                         :blnChgFlg      ：顧客コード変更フラグ
'                         :strBumoCD      ：部門コード
'                         :strBumoNm      ：部門名
'                         :lngNYUKIN_ZAN  ：入金残
'                         :lngFURIKOMI_KIN：振込金額
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInsertFVS770_W01(ByRef objWk As Recordset, ByVal objRsKONT As Recordset, _
                                     ByVal blnChgFlg As Boolean, ByVal strBumoCD As String, ByVal strBumoNm As String, _
                                     ByVal lngNYUKIN_ZAN As Long) As Boolean

    Dim varCode        As Variant

    Dim lngKEIYT_HOZUA As Long

    On Error GoTo fncInsertFVS770_W01_Err

    fncInsertFVS770_W01 = False

    With objWk
        .AddNew
        .Fields("REPORT_ID") = pstrCstRPT_ID                            ' 帳票ID
        .Fields("BUMOM_BUMOC") = strBumoCD                              ' 部門コード
        .Fields("BUMOM_BUMON") = strBumoNm                              ' 部門名

        If Nz(objRsKONT.Fields("CARG_LIQ_REQDATE"), "") <> "" Then
            .Fields("CARG_LIQ_REQDATE") = Format(objRsKONT.Fields("CARG_LIQ_REQDATE"), "YYYY/MM/DD")   ' 精算依頼日
        End If

        .Fields("CARG_UCODE") = objRsKONT.Fields("CARG_UCODE")          ' 顧客コード
        .Fields("USER_NAME") = objRsKONT.Fields("USER_NAME")            ' 顧客名

        ' 集金区分取得
        If Nz(objRsKONT.Fields("USER_SKBN"), "") <> "" Then
            .Fields("USER_SKBN") = objRsKONT.Fields("USER_SKBN")        ' 集金区分

            varCode = fncGetName(0, "030", objRsKONT.Fields("USER_SKBN"))  ' 2005/12/02 ADD T.SUZUKI
'            varCode = fncGETNAME("030", objRsKONT.Fields("USER_SKBN"))    ' 2005/12/02 DEL T.SUZUKI
            .Fields("USER_SKBN_NAME") = varCode
        End If

        .Fields("CARG_YCODE") = objRsKONT.Fields("CARG_YCODE")          ' ヤードコード
        .Fields("CARG_NO") = objRsKONT.Fields("CARG_NO")                ' コンテナ番号
        .Fields("CARG_KYDATE") = Format(objRsKONT.Fields("CARG_KYDATE"), "YYYY/MM/DD") ' 解約日付

' 2005/12/02 ADD T.SUZUKI Start
        ' 解約区分名(略称)取得
        varCode = fncGetName(1, "089", objRsKONT.Fields("CARG_KAICD"))  ' 解約区分
        .Fields("KAIYAKU_NAME") = varCode                               ' 解約区分名(略称)
' 2005/12/02 ADD T.SUZUKI End

        '.Fields("CARG_RENTKG") = objRsKONT.Fields("CARG_RENTKG")        ' 月額料金 ' 2007/03/20 DEL N.MIURA
        .Fields("CARG_RENTKG") = objRsKONT.Fields("CARG_RENTKG").VALUE + _
                                 objRsKONT.Fields("CARG_SYOZEI").VALUE   ' 月額料金 ' 2007/03/20 ADD N.MIRUA
        
        .Fields("CARG_LIQ_NOTE") = objRsKONT.Fields("CARG_LIQ_NOTE")    ' 精算備考

        ' 保証区分取得
        If Nz(objRsKONT.Fields("CARG_HOSYICD"), "") <> "" Then
            .Fields("CARG_HOSYICD") = objRsKONT.Fields("CARG_HOSYICD")  ' 保証区分

            varCode = fncGetName(0, "200", objRsKONT.Fields("CARG_HOSYICD"))  ' 2005/12/02 ADD T.SUZUKI
'            varCode = fncGETNAME("200", objRsKONT.Fields("CARG_HOSYICD"))    ' 2005/12/02 DEL T.SUZUKI
            .Fields("CARG_HOSYICD_NAME") = varCode
        End If

        .Fields("CARG_HOSYO_CD_NAME") = objRsKONT.Fields("SHIRM_SHIRN")  ' 保証会社名    2012/10/09 M.HONDA INS

        .Fields("CARG_HOSYO_CD") = objRsKONT.Fields("CARG_HOSYO_CD")     ' 保証会社ｺｰﾄﾞ  2012/10/09 M.HONDA INS
        
        '' 2012/10/31 M.HONDA START
        '' ﾊﾟﾙﾏ包括の保証会社の場合には保証金の取得方法を変える。
        If Nz(objRsKONT.Fields("CARG_HOSYO_CD"), 0) = "970031" Or Nz(objRsKONT.Fields("CARG_HOSYO_CD"), 0) = "970032" Then
            '' 預かり金取得
            lngKEIYT_HOZUA = fncGetKEIYT_HOZUA_PAR(objRsKONT.Fields("CARG_ACPTNO"))
        
        
        Else
            ' 預り金取得
            '' 2010/01/28 Ver0.5 M.HONDA START
            Select Case CInt(Nz(.Fields("CARG_HOSYICD"), "0"))
                Case 1, 3, 4, 6, 7   '日本賃貸、リプラス、日本セーフティー、保証会社あり
                    lngKEIYT_HOZUA = 0
                Case Else     '上記以外は預り金取得
                    lngKEIYT_HOZUA = fncGetKEIYT_HOZUA(objRsKONT.Fields("CARG_ACPTNO"))
            End Select
        End If
        '▼ 2007/04/27 add tajima
        ' 保証区分によっては取得する コード設計書 200 を参照
'        Select Case CInt(Nz(.Fields("CARG_HOSYICD"), "0"))
'            Case 3, 4, 6  '日本賃貸、リプラス、日本セーフティーは預金は０
'                lngKEIYT_HOZUA = 0
'            Case Else     '上記以外は預り金取得
'                lngKEIYT_HOZUA = fncGetKEIYT_HOZUA(objRsKONT.Fields("CARG_ACPTNO"))
'        End Select
        '▲ 2007/04/27 add tajima
        '' 2010/01/28 Ver0.5 M.HONDA END
        
        
        .Fields("KEIYT_HOZUA") = lngKEIYT_HOZUA                         ' 預り金

        If blnChgFlg = True Then
            ' 顧客コード変更時に下記の値を設定
            .Fields("NYUUKIN_ZAN") = lngNYUKIN_ZAN                      ' 入金残

            ' 他契約本数取得
            .Fields("KEIYAKU_HONSU") = fncGetKeiyaku_Honsu(objRsKONT.Fields("CARG_UCODE"))
        End If

' 205/11/11 DEL T.SUZUKI Start
'        If lngNYUKIN_ZAN > 0 Then
'            ' 『+』
'            .Fields("NYUKINZAN_FLG") = 0  ' 入金残フラグ
'
'        ElseIf lngNYUKIN_ZAN < 0 Then
'            ' 『-』
'            .Fields("NYUKINZAN_FLG") = 1  ' 入金残フラグ
'
'        ElseIf lngNYUKIN_ZAN = 0 Then
'            ' 『±0』
'            .Fields("NYUKINZAN_FLG") = 2  ' 入金残フラグ
'        End If
' 205/11/11 DEL T.SUZUKI End

        ' 受注契約番号
        .Fields("CARG_ACPTNO") = objRsKONT.Fields("CARG_ACPTNO")

        .UPDATE
    End With

    fncInsertFVS770_W01 = True

fncInsertFVS770_W01_Exit:

    Exit Function

fncInsertFVS770_W01_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncInsertFVS770_W01_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :subConnectDB
'        Parameter        :strBumoCD  部門コード
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectDB(ByVal strBumoCD As String)

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo subConnectDB_Err

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME_" & strBumoCD)
    If strDataSource = "" Then
        ' テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If

    ' DB接続
    Set objKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(CStr(strBumoCD)))
    ' -----------------------------------------------------------------------------------------------------------

    ' -----------------------------------------------------------------------------------------------------------
    ' 加瀬DB接続
    strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME")

    If strDataSource = "" Then
        ' テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If

    Set objKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
    ' -----------------------------------------------------------------------------------------------------------

    Exit Sub

subConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :subDisConnectDB
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDisConnectDB()

    On Error GoTo subDisConnectDB_Err

    ' 加瀬DB
    objKASE_DB.Close
    Set objKASE_DB = Nothing

    ' コンテナDB
    objKONT_DB.Close
    Set objKONT_DB = Nothing

    Exit Sub

subDisConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ODBC接続文字列取得
'        MODULE_ID        :fncGetDataSource
'        Parameter        :strODBCNAME ODBC名
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetDataSource(ByVal strODBCNAME As String) As String

    Dim strName As String

    On Error GoTo fncGetDataSource_Err

    fncGetDataSource = ""

    strName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & strODBCNAME & "'"))

    fncGetDataSource = strName

    Exit Function

fncGetDataSource_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :subGetBumonName
'        IN               :objDB = DB接続 /strBumonCd = 部門コード /strBumonNm = 部門名
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetBumonName()
'Private Sub subGetBumonName(ByRef objDb As Database, ByRef strBumonCd As String, ByRef strBumonNm As String)

    Dim strSQL  As String
    Dim objRs   As Recordset
    Dim objDb   As Database          ' 2005/12/02 ADD T.SUZUKI

    On Error GoTo subGetBumonName_Err

    Set objDb = CurrentDb  ' 2005/12/02 ADD T.SUZUKI

    strSQL = "SELECT "
    strSQL = strSQL & "BUMOM_BUMOC, "  ' 部門コード
    strSQL = strSQL & "BUMOM_BUMON "   ' 部門名称
    strSQL = strSQL & "FROM dbo_CONT_MAST INNER JOIN BUMO_MAST ON "
    strSQL = strSQL & "dbo_CONT_MAST.CONT_BUMOC = BUMO_MAST.BUMOM_BUMOC "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            pstrBumonCd = .Fields("BUMOM_BUMOC")
            pstrBumonNm = .Fields("BUMOM_BUMON")
        Else
            pstrBumonCd = Null
            pstrBumonNm = Null
        End If
'        If Not .EOF Then
'            strBumonCd = .Fields("BUMOM_BUMOC")
'            strBumonNm = .Fields("BUMOM_BUMON")
'        Else
'            strBumonCd = Null
'            strBumonNm = Null
'        End If
    End With

subGetBumonName_Exit:
' 2005/12/02 ADD T.SUZUKI Start
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
' 2005/12/02 ADD T.SUZUKI End
    Exit Sub

subGetBumonName_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo subGetBumonName_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Where句生成
'        MODULE_ID        :fncCreateWhere
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWhere() As String

    Dim strWhere As String

    With Me
        ' *** 精算依頼日の大小チェック ***
        Call ChkDateFromTo(.txt_Seisan_ReqDate_F, .txt_Seisan_ReqDate_T)

        ' *** ヤードコードの大小チェック ***
        Call ChkYARDFromTo(.cmb_YCODE_F, .txt_YCODE_F, .cmb_YCODE_T, .txt_YCODE_T)

' 2005/12/02 ADD T.SUZUKI Start
        ' *** 解約区分の大小チェック ***
        Call ChkYARDFromTo(.cmb_KAIYAKU_F, .txt_KAIYAKU_F, .cmb_KAIYAKU_T, .txt_KAIYAKU_T)
' 2005/12/02 ADD T.SUZUKI End

        ' 対象解約年月
        pstrKYDATE = Replace(.txt_KaiyakuYYYYMM, "/", "")                 ' 解約日付(YYYYMM)
        pstrKYDATE_F = .txt_KaiyakuYYYYMM & "/01"                         ' 解約日付(From)
        pstrKYDATE_T = DateAdd("D", "-1", DateAdd("M", 1, pstrKYDATE_F))  ' 解約日付(To)

        strWhere = " AND CARG_KYDATE BETWEEN '" & pstrKYDATE_F & "' AND '" & pstrKYDATE_T & "' "  ' 解約日付(YYYYMM)の初日～末日

        ' 精算依頼日
        If (Nz(.txt_Seisan_ReqDate_F, "") <> "" And Nz(.txt_Seisan_ReqDate_T, "") <> "") Then
            ' 精算依頼日(From)：入力
            ' 精算依頼日(To)  ：入力
            strWhere = strWhere & " AND CARG_LIQ_REQDATE BETWEEN '" & .txt_Seisan_ReqDate_F & "' And '" & .txt_Seisan_ReqDate_T & "' "

        ElseIf (Nz(.txt_Seisan_ReqDate_F, "") <> "" And Nz(.txt_Seisan_ReqDate_T, "") = "") Then
            ' 精算依頼日(From)：入力
            ' 精算依頼日(To)  ：未入力
            strWhere = strWhere & " AND CARG_LIQ_REQDATE >= '" & .txt_Seisan_ReqDate_F & "' "

        ElseIf (Nz(.txt_Seisan_ReqDate_F, "") = "" And Nz(.txt_Seisan_ReqDate_T, "") <> "") Then
            ' 精算依頼日(From)：未入力
            ' 精算依頼日(To)  ：入力
            strWhere = strWhere & " AND CARG_LIQ_REQDATE <= '" & txt_Seisan_ReqDate_T & "' "
        End If

        ' 精算依頼日
        strWhere = strWhere & " AND CARG_LIQ_REQDATE IS NOT NULL "

        ' ヤードコード
        If (Nz(.cmb_YCODE_F, "") <> "" And Nz(.cmb_YCODE_T, "") <> "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：入力
            strWhere = strWhere & " AND CARG_YCODE BETWEEN " & .cmb_YCODE_F & " AND " & .cmb_YCODE_T

        ElseIf (Nz(.cmb_YCODE_F, "") <> "" And Nz(.cmb_YCODE_T, "") = "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：未入力
            strWhere = strWhere & " AND CARG_YCODE >= " & .cmb_YCODE_F

        ElseIf (Nz(.cmb_YCODE_F, "") = "" And Nz(.cmb_YCODE_T, "") <> "") Then
            ' ヤードコード(From)：未入力
            ' ヤードコード(To)  ：入力
            strWhere = strWhere & " AND CARG_YCODE <= " & .cmb_YCODE_T
        End If

' 2005/12/02 ADD T.SUZUKI Start
        ' 解約区分
        If (Nz(.cmb_KAIYAKU_F, "") <> "" And Nz(.cmb_KAIYAKU_T, "") <> "") Then
            ' 解約区分(From)：入力
            ' 解約区分(To)  ：入力
            strWhere = strWhere & " AND CARG_KAICD BETWEEN " & .cmb_KAIYAKU_F & " AND " & .cmb_KAIYAKU_T

        ElseIf (Nz(.cmb_KAIYAKU_F, "") <> "" And Nz(.cmb_KAIYAKU_T, "") = "") Then
            ' 解約区分(From)：入力
            ' 解約区分(To)  ：未入力
            strWhere = strWhere & " AND CARG_KAICD >= " & .cmb_KAIYAKU_F

        ElseIf (Nz(.cmb_KAIYAKU_F, "") = "" And Nz(.cmb_KAIYAKU_T, "") <> "") Then
            ' 解約区分(From)：未入力
            ' 解約区分(To)  ：入力
            strWhere = strWhere & " AND CARG_KAICD <= " & .cmb_KAIYAKU_T
        End If

        ' 顧客コード
        If Nz(.cmb_顧客コード, "") <> "" Then
            strWhere = strWhere & " AND CARG_UCODE = " & .cmb_顧客コード
        End If
' 2005/12/02 ADD T.SUZUKI End

        ' 精算依頼済みを対象
        If .chk_Seisan_Req.VALUE = -1 Then
            ' ON
            strWhere = strWhere & " AND CARG_LIQ_FLG = " & pintCstKAKUTEI    ' 精算依頼済み(確定)
        Else
            ' OFF
            strWhere = strWhere & " AND CARG_LIQ_FLG = " & pintCstMIKAKUTEI  ' 精算依頼済み(未確定)
        End If
        
        ' 2012/10/09 M.HONDA START
        ' ﾊﾟﾙﾏ包括のみを対象
        If .chk_Parma.VALUE = -1 Then
            ' ON
            strWhere = strWhere & " AND CARG_HOSYO_CD IN ('970031','970032') "
        End If
        ' 2012/10/09 M.HONDA END
        
        
    End With

    fncCreateWhere = strWhere
End Function

'==============================================================================*
'
'        MODULE_NAME      :名称取得
'        MODULE_ID        :fncGetName
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetName(ByVal intKbn As Integer, ByVal strNAME_ID As String, ByVal varNAME_CODE As Variant) As String
'Private Function fncGETNAME(ByVal strNAME_ID As String, ByVal varNAME_CODE As Variant) As String

    Dim strWhere  As String
    Dim strName   As String

    On Error GoTo fncGetName_Err

    fncGetName = ""

    ' 名称マスタより略称を取得
    strWhere = " NAME_ID = '" & strNAME_ID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & varNAME_CODE

    strName = ""
    Select Case intKbn
        Case 0  ' *** 名称 ***
            strName = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
        Case 1  ' *** 略称 ***
            strName = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWhere), "")
    End Select

    If strName <> "" Then
        fncGetName = strName
    End If

fncGetName_Exit:
    Exit Function

fncGetName_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncGetName_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :預り金取得（パルマ包括限定）
'        MODULE_ID        :fncGetKEIYT_HOZUA_PAR
'        Parameter        :strCARG_ACPTNO コンテナ契約ファイル.受注契約番号
'        CREATE_DATE      :2012/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKEIYT_HOZUA_PAR(ByVal strCARG_ACPTNO As String) As Long

    Dim strSQL         As String
    Dim objRs          As Recordset
    Dim lngKEIYT_HOZUA As Long

    On Error GoTo fncGetKEIYT_HOZUA_PAR_Err

    fncGetKEIYT_HOZUA_PAR = 0
    lngKEIYT_HOZUA = 0

    strSQL = "          select b.KEIYT_KEIYB, "
    strSQL = strSQL & "        b.KEIYT_HOSYI, "
    strSQL = strSQL & "        b.KEIYT_HOKAC, "
    'strSQL = strSQL & "        b.KEIYT_HOSYA "
    strSQL = strSQL & "        b.KEIYT_HKSYA "                                  'INSERT 2023/04/06 N.IMAI
    strSQL = strSQL & "   from KEIY_TRAN a, "
    strSQL = strSQL & "        KEIY_TRAN b "
    strSQL = strSQL & "  where a.KEIYT_KEIYB = '" & strCARG_ACPTNO & "'"
    strSQL = strSQL & "  and   a.KEIYT_KKOSB = b.KEIYT_KEIYB "
    strSQL = strSQL & "  and   b.KEIYT_STATI <> '10' "

    Set objRs = objKASE_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    If objRs.EOF = False Then                                                   'INSERT 2017/10/10 N.IMAI
        '' 一つ前の契約の保証会社がパルマ包括の場合には再度、本処理を呼びだす。
        If Nz(objRs.Fields("KEIYT_HOKAC"), 0) = "970031" Or Nz(objRs.Fields("KEIYT_HOKAC"), 0) = "970032" Then
            Call fncGetKEIYT_HOZUA_PAR(objRs.Fields("KEIYT_KEIYB"))
        End If
    
        Select Case CInt(Nz(objRs.Fields("KEIYT_HOSYI"), "0"))
            Case 1, 3, 4, 6, 7   '日本賃貸、リプラス、日本セーフティー、保証会社あり
                lngKEIYT_HOZUA = 0
            Case Else     '上記以外は預り金取得
                'lngKEIYT_HOZUA = Nz(objRs.Fields("KEIYT_HOSYA"), 0)
                lngKEIYT_HOZUA = Nz(objRs.Fields("KEIYT_HKSYA"), 0)             'INSERT 2023/04/06 N.IMAI
        End Select
    End If                                                                      'INSERT 2017/10/10 N.IMAI

    fncGetKEIYT_HOZUA_PAR = lngKEIYT_HOZUA

fncGetKEIYT_HOZUA_PAR_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Function

fncGetKEIYT_HOZUA_PAR_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo fncGetKEIYT_HOZUA_PAR_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :預り金取得
'        MODULE_ID        :fncGetKEIYT_HOZUA
'        Parameter        :strCARG_ACPTNO コンテナ契約ファイル.受注契約番号
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKEIYT_HOZUA(ByVal strCARG_ACPTNO As String) As Long

    Dim strSQL         As String
    Dim objRs          As Recordset
    Dim lngKEIYT_HOZUA As Long

    On Error GoTo fncGetKEIYT_HOZUA_Err

    fncGetKEIYT_HOZUA = 0
    lngKEIYT_HOZUA = 0

    '' 2010/01/28 M.HONDA START
    'strSQL = "SELECT KEIYT_HOSYA " & Chr(13)     ' 預り保証金
    strSQL = "SELECT KEIYT_HKSYA " & Chr(13)     ' 預り保証金                   'INSERT 2023/04/06 N.IMAI
    '' strSQL = "SELECT KEIYT_HOZUA " & Chr(13)  ' 回収済保証金
    '' 2010/01/28 M.HONDA END
    strSQL = strSQL & " FROM KEIY_TRAN " & Chr(13)  ' **受注契約トラン
    strSQL = strSQL & " WHERE KEIYT_KEIYB = '" & strCARG_ACPTNO & "'"
    Set objRs = objKASE_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    If objRs.EOF = False Then
        'lngKEIYT_HOZUA = Nz(objRs.Fields("KEIYT_HOSYA"), 0)
        lngKEIYT_HOZUA = Nz(objRs.Fields("KEIYT_HKSYA"), 0)                     'INSERT 2023/04/06 N.IMAI
    End If

    fncGetKEIYT_HOZUA = lngKEIYT_HOZUA

fncGetKEIYT_HOZUA_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Function

fncGetKEIYT_HOZUA_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo fncGetKEIYT_HOZUA_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :売上合計金額取得
'        MODULE_ID        :fncGetURI_TOTAL
'        Parameter        :varBumoCD      部門コード
'                          varKOK_CD      顧客コード
'                          varKAIYAKU_YM  解約年月(YYYY/MM)
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetURI_TOTAL(ByVal strBumoCD As String, ByVal strUSER_CODE As Variant, ByVal strKYDATE As String) As Long

    Dim strSQL       As String
    Dim objRs        As Recordset
    Dim lngURI_TOTAL As Long

    On Error GoTo fncGetURI_TOTAL_Err

    fncGetURI_TOTAL = 0

    strSQL = " SELECT SUM(URIAT_URIAG) AS TOTAL " & Chr(13)   ' 売上金額
    strSQL = strSQL & " FROM URIA_TRAN " & Chr(13)  ' ***売掛トラン
    strSQL = strSQL & " WHERE URIAT_BUMOC = '" & strBumoCD & "' " & Chr(13)     ' 部門コード
    strSQL = strSQL & "   AND URIAT_KOKYC = '" & strUSER_CODE & "' " & Chr(13)  ' 顧客コード
    'strSQL = strSQL & "   AND URIAT_URYMD <= '" & strKYDATE & "' " & Chr(13)    ' 売上年月 'DELETE 2024/12/20 N.IMAI
    Set objRs = objKASE_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    ' 売上金額合計
    lngURI_TOTAL = Nz(objRs.Fields("TOTAL").VALUE, 0)

    fncGetURI_TOTAL = lngURI_TOTAL

fncGetURI_TOTAL_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Function

fncGetURI_TOTAL_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo fncGetURI_TOTAL_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :回収合計金額取得
'        MODULE_ID        :fncGetKAIS_TOTAL
'        Parameter        :varBumoCD      部門コード
'                          varKOK_CD      顧客コード
'                          varKAIYAKU_YM  解約年月(YYYY/MM)
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKAIS_TOTAL(ByVal strBumoCD As String, ByVal strUSER_CODE As String, ByVal strKYDATE As String) As Long

    Dim strSQL        As String
    Dim objRs         As Recordset
    Dim lngKAIS_TOTAL As Long

    On Error GoTo fncGetKAIS_TOTAL_Err

    fncGetKAIS_TOTAL = 0

    strSQL = " SELECT SUM(KAIST_KAISA) AS TOTAL " & Chr(13)  ' 回収金額
    strSQL = strSQL & " FROM KAIS_TRAN " & Chr(13)  ' **回収トラン
    strSQL = strSQL & " WHERE KAIST_BUMOC = '" & strBumoCD & "' " & Chr(13)     ' 部門コード
    strSQL = strSQL & "   AND KAIST_KOKYC = '" & strUSER_CODE & "' " & Chr(13)  ' 顧客コード
    'strSQL = strSQL & "   AND KAIST_KAYMD <= '" & strKYDATE & "' " & Chr(13)    ' 回収年月 'DELETE 2024/12/20 N.IMAI
    Set objRs = objKASE_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    lngKAIS_TOTAL = Nz(objRs.Fields("TOTAL").VALUE, 0)

    fncGetKAIS_TOTAL = lngKAIS_TOTAL

fncGetKAIS_TOTAL_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Function

fncGetKAIS_TOTAL_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo fncGetKAIS_TOTAL_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :他契約本数取得
'        MODULE_ID        :fncGetKeiyaku_Honsu
'        Parameter        :varCARG_UCODE 顧客コード
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKeiyaku_Honsu(ByVal varCARG_UCODE As Variant) As Integer

    Dim strSQL       As String
    Dim objRs        As Recordset
    Dim intDataCount As Integer

    On Error GoTo fncGetKeiyaku_Honsu_Err

    fncGetKeiyaku_Honsu = 0

    intDataCount = 0

    ' 契約本数取得
    strSQL = "SELECT COUNT(*) AS Cnt " & Chr(13)
    strSQL = strSQL & " FROM dbo_CARG_FILE " & Chr(13)  ' **コンテナ契約ファイル
    strSQL = strSQL & " WHERE CARG_AGRE <> " & pintCARG_AGRE & Chr(13)  ' 完了以外
    strSQL = strSQL & "   AND (format(CARG_KYDATE, 'YYYYMM') > '" & pstrKYDATE & "' " & Chr(13)
    strSQL = strSQL & "    OR  CARG_KYDATE IS NULL) " & Chr(13)
    strSQL = strSQL & "   AND CARG_UCODE = " & Val(varCARG_UCODE)       ' 顧客コード
    Set objRs = pobjDB.OpenRecordset(strSQL, dbReadOnly)

    intDataCount = objRs.Fields("Cnt")  ' 契約本数

    fncGetKeiyaku_Honsu = intDataCount

fncGetKeiyaku_Honsu_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Function

fncGetKeiyaku_Honsu_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncGetKeiyaku_Honsu_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約ファイル.請求済フラグの更新
'        MODULE_ID        :fncUPD_CARG_FILE
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUPD_CARG_FILE() As Boolean

    Dim strSQL   As String
    Dim objRs    As Recordset
    Dim objWKRs  As Recordset

    Dim objWs    As Workspace

    On Error GoTo fncUPD_CARG_FILE_Err

    fncUPD_CARG_FILE = False

    Set objWs = DBEngine.Workspaces(0)
    objWs.BeginTrans

    strSQL = "SELECT CARG_ACPTNO " & Chr(13)
    strSQL = strSQL & " FROM FVS770_W01 " & Chr(13)
    strSQL = strSQL & " ORDER BY FURIKOMIKIN_FLG, " & Chr(13)  ' 2005/11/11 ADD T.SUZUKI
'    strSQL = strSQL & " ORDER BY NYUKINZAN_FLG, " & Chr(13)   ' 2005/11/11 DEL T.SUZUKI
    strSQL = strSQL & "          SortKey, " & Chr(13)
    strSQL = strSQL & "          CARG_UCODE, " & Chr(13)
    strSQL = strSQL & "          CARG_YCODE, " & Chr(13)
    strSQL = strSQL & "          CARG_NO "
    Set objRs = pobjDB.OpenRecordset(strSQL, dbReadOnly)

    Do Until objRs.EOF
        strSQL = "SELECT CARG_LIQ_FLG, " & Chr(13)
        strSQL = strSQL & " CARG_UPDAD, " & Chr(13)
        strSQL = strSQL & " CARG_UPDAJ, " & Chr(13)
        strSQL = strSQL & " CARG_UPDPB, " & Chr(13)
        strSQL = strSQL & " CARG_UPDUB " & Chr(13)
        strSQL = strSQL & " FROM dbo_CARG_FILE " & Chr(13)
        strSQL = strSQL & " WHERE CARG_ACPTNO = '" & objRs.Fields("CARG_ACPTNO") & "' "
        Set objWKRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

        ' 請求済フラグの更新
        With objWKRs
            .Edit
            .Fields("CARG_LIQ_FLG") = pintCstKAKUTEI           ' 請求済フラグ

            .Fields("CARG_UPDAD") = Format$(DATE, "YYYYMMDD")  ' 更新日付
            .Fields("CARG_UPDAJ") = Format$(time, "hhmmss")    ' 更新時刻
            .Fields("CARG_UPDPB") = Me.Form.NAME               ' 更新プログラムID
            .Fields("CARG_UPDUB") = MSZZ000.LsGetUserName      ' 更新ユーザーID

            .UPDATE
        End With

        objWKRs.Close
        Set objWKRs = Nothing

        objRs.MoveNext
    Loop

    objWs.CommitTrans

    fncUPD_CARG_FILE = True

fncUPD_CARG_FILE_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If

    If Not objWKRs Is Nothing Then
        objWKRs.Close
        Set objWKRs = Nothing
    End If

    Exit Function

fncUPD_CARG_FILE_Err:
    objWs.Rollback

    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo fncUPD_CARG_FILE_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルへ対象データのInsert
'        MODULE_ID        :fncSetPageSortKey
'        Parameter        :intPageCount ページ数
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetPageSortKey(ByRef intPageCount As Integer) As Boolean

    Dim strSQL             As String
    Dim objRs              As Recordset
    Dim objRsWk            As Recordset

    Dim strSaveKokyc       As String     ' 保存用顧客コード
    Dim intUserCount       As Integer    ' 顧客データ数

    Dim lngSortKey         As Long       ' ソートキー
    Dim intRowCount        As Long       ' 行数
    Dim lngCurrentPage     As Long       ' カレントページ
    Dim intFURIKOMIKIN_FLG As Integer    ' 205/11/11 ADD T.SUZUKI
'    Dim intNYUKINZAN_FLG   As Integer   ' 205/11/11 DEL T.SUZUKI

    Dim blnKaiPageFlg    As Boolean    ' 改ページフラグ

    Dim blnMaxFlg        As Boolean

    On Error GoTo fncSetPageSortKey_Err

    fncSetPageSortKey = False

    ' データ取得
    strSQL = "SELECT CARG_UCODE, " & Chr(13)          ' 顧客コード
    strSQL = strSQL & " FURIKOMIKIN_FLG, " & Chr(13)  ' 入金フラグ  ' 2005/11/11 ADD T.SUZUKI
'    strSQL = strSQL & " NYUKINZAN_FLG, " & Chr(13)  ' 入金フラグ   ' 2005/11/11 DEL T.SUZUKI
    strSQL = strSQL & " SortKey " & Chr(13)         ' ソートキー
    strSQL = strSQL & "FROM FVS770_W01 " & Chr(13)
    strSQL = strSQL & " ORDER BY FURIKOMIKIN_FLG, " & Chr(13)       ' 2005/11/11 ADD T.SUZUKI
'    strSQL = strSQL & " ORDER BY NYUKINZAN_FLG, " & Chr(13)        ' 2005/11/11 DEL T.SUZUKI
    strSQL = strSQL & "          CARG_UCODE, " & Chr(13)
    strSQL = strSQL & "          CARG_YCODE, " & Chr(13)
    strSQL = strSQL & "          CARG_NO "
    
    
    
    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

    strSaveKokyc = ""      ' 顧客コード
    intRowCount = 0        ' 行数カウント
    lngCurrentPage = 1     ' ページ数カウント
    lngSortKey = 0         ' ソートキー
    blnKaiPageFlg = False  ' 改ページフラグ
    blnMaxFlg = False      ' 最大行数オーバーフラグ
    Do Until objRs.EOF
        If intRowCount = 0 Then
' 2005/11/11 REP T.SUZUKI Start
            ' 入金残フラグ(0 OR 1 OR 2)
            intFURIKOMIKIN_FLG = objRs.Fields("FURIKOMIKIN_FLG")
'            ' 入金残フラグ(0 OR 1 OR 2)
'            intNYUKINZAN_FLG = objRs.Fields("NYUKINZAN_FLG")
' 2005/11/11 REP T.SUZUKI End
        End If

        If strSaveKokyc <> objRs.Fields("CARG_UCODE") Then
            ' 顧客コードが変更になったとき
            ' 顧客データ数取得
            strSQL = "SELECT COUNT(*) AS Cnt " & Chr(13)
            strSQL = strSQL & " FROM FVS770_W01 " & Chr(13)
            strSQL = strSQL & " WHERE CARG_UCODE = " & objRs.Fields("CARG_UCODE")
            Set objRsWk = pobjDB.OpenRecordset(strSQL, dbReadOnly)
    
            intUserCount = Nz(objRsWk.Fields("Cnt"), 0)  ' ユーザーデータ数格納
            objRsWk.Clone
            Set objRsWk = Nothing

            If (lngCurrentPage = 1 And (intRowCount + intUserCount) > pintCstFirstPage) Then
                ' 1ページ目
                If intRowCount = 0 Then
                    ' ページ先頭行で最大行数をオーバーしてしまった場合
                    blnMaxFlg = True
                Else
                    blnKaiPageFlg = True
                End If

            ElseIf (lngCurrentPage > 1 And (intRowCount + intUserCount) > pintCstOtherPage) Then
                ' 2ページ目以降
                If intRowCount = 0 Then
                    ' ページ先頭行で最大行数をオーバーしてしまった場合
                    blnMaxFlg = True
             Else
                    blnKaiPageFlg = True
                End If
            End If

            ' 変更後顧客コード格納
            strSaveKokyc = objRs.Fields("CARG_UCODE")
        End If

        If blnMaxFlg = True Then
            If lngCurrentPage = 1 Then
                ' 1ページ目
                If intRowCount >= pintCstFirstPage Then
                    blnKaiPageFlg = True
                    blnMaxFlg = False
                Else
                    ' 改ページはさせない
                End If

            Else
                ' 2ページ目以降
                If intRowCount >= pintCstOtherPage Then
                    blnKaiPageFlg = True
                    blnMaxFlg = False
                Else
                    ' 改ページはさせない
                End If
            End If
        End If

        If blnKaiPageFlg = False Then
            If (lngCurrentPage = 1 And intRowCount >= pintCstFirstPage) Then
                ' 1ページ目のMax行数に達したとき
                blnKaiPageFlg = True
            ElseIf (lngCurrentPage > 1 And intRowCount >= pintCstOtherPage) Then
                ' 2ページ目以降のMax行数に達したとき
                blnKaiPageFlg = True
' 205/11/11 ADD T.SUZUKI Start
            ElseIf intFURIKOMIKIN_FLG <> objRs.Fields("FURIKOMIKIN_FLG") Then
                ' 振込金額フラグが変更になったとき
                blnKaiPageFlg = True
                intFURIKOMIKIN_FLG = objRs.Fields("FURIKOMIKIN_FLG")  ' 振込金額フラグ設定
' 205/11/11 ADD T.SUZUKI End
' 205/11/11 DEL T.SUZUKI Start
'            ElseIf intNYUKINZAN_FLG <> objRs.Fields("NYUKINZAN_FLG") Then
'                ' 入金フラグが変更になったとき
'                blnKaiPageFlg = True
'                intNYUKINZAN_FLG = objRs.Fields("NYUKINZAN_FLG")    ' 入金残フラグ設定
' 205/11/11 DEL T.SUZUKI End
            End If
        End If

        If blnKaiPageFlg = True Then
            lngSortKey = lngSortKey + 1                       ' ソートキーカウントアップ
            lngCurrentPage = lngCurrentPage + 1               ' ページ数カウントアップ

            intRowCount = 0

            blnKaiPageFlg = False  ' 改ページフラグ初期化
        End If

        ' ソートキー更新
        objRs.Edit
        objRs.Fields("SortKey") = lngSortKey
        objRs.UPDATE

        intRowCount = intRowCount + 1  ' 行数カウントアップ

        objRs.MoveNext
    Loop

    intPageCount = lngSortKey + 1  ' ページ数セット
    fncSetPageSortKey = True

fncSetPageSortKey_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If

    If Not objRsWk Is Nothing Then
        objRsWk.Close
        Set objRsWk = Nothing
    End If
    Exit Function

fncSetPageSortKey_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncSetPageSortKey_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :振込金額の更新
'        MODULE_ID        :fncUPDATE_FURIKOMIKIN
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUPDATE_FURIKOMIKIN() As Integer
'Private Function fncUPDATE_FURIKOMIKIN() As Boolean  ' 2005/11/11 DEL T.SUZUKI

    Dim strSQL          As String
    Dim objRs           As Recordset
    Dim objRsWk         As Recordset

' 2005/11/11 ADD T.SUZUKI Start
    Dim lngFURIKOMI_KIN As Long     ' 振込金額
    Dim blnInsFlg       As Boolean  ' Insertフラグ
    Dim intCount        As Integer  ' カウンター
    Dim intOutPutCount  As Integer  ' 出力データ数
' 2005/11/11 ADD T.SUZUKI End

    On Error GoTo fncUPDATE_FURIKOMIKIN_Err

    fncUPDATE_FURIKOMIKIN = -1

    ' 顧客コード毎に預り金、入金残を取得
    strSQL = "SELECT CARG_UCODE, "
    strSQL = strSQL & " SUM(NZ(KEIYT_HOZUA, 0)) AS HOZUA, " & Chr(13)  ' 預り金
    strSQL = strSQL & " SUM(NZ(NYUUKIN_ZAN, 0)) AS ZAN " & Chr(13)     ' 入金残
    strSQL = strSQL & " FROM FVS770_W01 " & Chr(13)
    strSQL = strSQL & " GROUP BY CARG_UCODE " & Chr(13)
    strSQL = strSQL & " ORDER BY CARG_UCODE "
    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

    intOutPutCount = 0  ' 2005/11/11 ADD T.SUZUKI
    Do Until objRs.EOF
' 2005/11/11 ADD T.SUZUKI Start
        lngFURIKOMI_KIN = 0  ' 振込金額
        blnInsFlg = False    ' Insertフラグ
' 2005/11/11 ADD T.SUZUKI End

        ' ワークテーブルの内容を抽出
' 2005/11/11 ADD T.SUZUKI Start
        strSQL = " SELECT FURIKOMI_KIN, " & Chr(13)      ' 振込金額
        strSQL = strSQL & " FURIKOMIKIN_FLG " & Chr(13)  ' 振込金額フラグ
        strSQL = strSQL & " FROM FVS770_W01 " & Chr(13)
        strSQL = strSQL & " WHERE CARG_UCODE = " & objRs.Fields("CARG_UCODE") & Chr(13)  ' 顧客コード
        strSQL = strSQL & " ORDER BY CARG_YCODE DESC, " & Chr(13)  ' ヤードコード
        strSQL = strSQL & "          CARG_NO DESC "                ' コンテナ番号
' 2005/11/11 ADD T.SUZUKI End
' 2005/11/11 DEL T.SUZUKI Start
'        strSQL = " SELECT FURIKOMI_KIN " & Chr(13)  ' 振込金額
'        strSQL = strSQL & " FROM FVS770_W01 " & Chr(13)
'        strSQL = strSQL & " WHERE CARG_UCODE = " & objRs.Fields("CARG_UCODE") & Chr(13)
'        strSQL = strSQL & "   AND KEIYAKU_HONSU IS NOT NULL " & Chr(13)
' 2005/11/11 DEL T.SUZUKI End
        Set objRsWk = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

' 2005/11/11 ADD T.SUZUKI Start
        lngFURIKOMI_KIN = objRs.Fields("HOZUA") - objRs.Fields("ZAN")  ' 振込金額(預り金額 - 入金残)

        ' 振込金額の判定
        Select Case Me.opg_選択
            Case 0  ' *** 全て ***
                blnInsFlg = True
            Case 1  ' *** 未収 ***
                If lngFURIKOMI_KIN > 0 Then
                    blnInsFlg = True
                End If
            Case 2  ' *** 過入金 ***
                If lngFURIKOMI_KIN < 0 Then
                    blnInsFlg = True
                End If
            Case 3  ' *** 精算なし ***
                If lngFURIKOMI_KIN = 0 Then
                    blnInsFlg = True
                End If
            Case 4  ' *** 保証金 ***
                If objRs.Fields("HOZUA") <> 0 Then
                    blnInsFlg = True
                End If
        End Select

        ' 振込金額設定
        intCount = 0
        With objRsWk
            If blnInsFlg = True Then
                Do Until .EOF
                    ' 該当する顧客レコードの先頭に振込金額を設定
                    .Edit

                    ' フラグの判定
' 205/11/11 DEL T.SUZUKI Start
                    If lngFURIKOMI_KIN > 0 Then
                        ' 『+』
                        .Fields("FURIKOMIKIN_FLG") = 0  ' 入金残フラグ

                    ElseIf lngFURIKOMI_KIN < 0 Then
                        ' 『-』
                        .Fields("FURIKOMIKIN_FLG") = 1  ' 入金残フラグ

                    ElseIf lngFURIKOMI_KIN = 0 Then
                        ' 『±0』
                        .Fields("FURIKOMIKIN_FLG") = 2  ' 入金残フラグ
                    End If
 '205/11/11 DEL T.SUZUKI End

                    If intCount = 0 Then
                        .Fields("FURIKOMI_KIN") = lngFURIKOMI_KIN  ' 振込金額
                        intCount = intCount + 1
                    End If

                    .UPDATE

                    intOutPutCount = intOutPutCount + 1  ' 出力データ数カウントアップ

                    .MoveNext
                Loop

            Else
                ' 出力対象外のため、ワークテーブルから削除
                Do Until .EOF
                    .Edit
                    .Delete

                    .MoveNext
                Loop
            End If
        End With
' 2005/11/11 ADD T.SUZUKI End
' 2005/11/11 DEL T.SUZUKI Start
'        With objRsWk
'            .Edit
'            .Fields("FURIKOMI_KIN") = objRs.Fields("HOZUA") - objRs.Fields("ZAN")  ' 預り金額 - 入金残
'            .Update
'        End With
' 2005/11/11 DEL T.SUZUKI End

        objRsWk.Close
        Set objRsWk = Nothing

        objRs.MoveNext
    Loop

    fncUPDATE_FURIKOMIKIN = intOutPutCount  ' 2005/11/11 ADD T.SUZUKI
'    fncUPDATE_FURIKOMIKIN = True           ' 2005/11/11 DEL T.SUZUKI

fncUPDATE_FURIKOMIKIN_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If

    If Not objRsWk Is Nothing Then
        objRsWk.Close
        Set objRsWk = Nothing
    End If
    Exit Function

fncUPDATE_FURIKOMIKIN_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncUPDATE_FURIKOMIKIN_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :文字項目のチェックを一括でする。
'        MODULE_ID        :CmBCheck_Char
'        CREATE_DATE      :
'
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

    On Error GoTo err_rtn

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
       CmBCheck_Char = 2
       GoTo Exit_rtn
    End If

    CmBCheck_Char = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmBCheck_Char = 9
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :数値項目のチェックを一括でする。
'        MODULE_ID        :CmBCheck_Num
'        CREATE_DATE      :
'
  '返り値は、
  '１：数値と判断できない時
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
    
  '引数 oText ←判定する値
  '     iLen  ←整数桁
  '     isLen ←小数桁(省略された時は、小数点桁数無し)
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, _
                              Optional ByVal isLEN As Variant) As Integer

    Dim vINT    As Variant
    Dim vsINT   As Variant
    Dim i       As Integer
      
    Dim dTEXT   As Double
    Dim dFROM   As Double
    Dim dTO     As Double
    Dim vRet    As Variant

    On Error GoTo err_rtn

    CmbCheck_Num = 9

    If Nz(CmNumeric(oText), "") = "" Then              ' 数値と判断できないのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    If CmChkCher(CStr(Nz(oText)), "A") = False Then    ' 全角文字が入力されているのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    ' 小数点の値が設定されている時のチェック
    vRet = Nz(InStr(Abs(oText), "."), 0)

    If IsMissing(isLEN) = False Then
        If vRet <> 0 Then
            If isLEN < CmLenb(Abs(oText)) - vRet Then  ' 小数点桁数が設定桁数を超えているエラー
                CmbCheck_Num = 2
                GoTo Exit_rtn
            End If
        End If

    Else
        If CmLenb(Abs(CmNumeric(oText))) > ilen Then   ' 文字数超過のためエラー
            CmbCheck_Num = 2
            GoTo Exit_rtn
        End If
    
        If vRet <> 0 Then                              ' 小数点の入力があるのでエラー
            CmbCheck_Num = 2
            GoTo Exit_rtn
        End If
    End If

    ' 桁チェック
    i = 0
    For i = 1 To ilen
        vINT = vINT & "9"
    Next

    If IsMissing(isLEN) = False Then
        vsINT = "."
        i = 0
        For i = 1 To isLEN
            vsINT = vsINT & "9"
        Next
    End If

    vINT = vINT & vsINT
    vINT = CmNumeric(vINT)

    dTEXT = CmNumeric(oText)
    dFROM = vINT * -1
    dTO = vINT
  
    If dFROM > dTEXT Or dTO < dTEXT Then               ' 論理範囲超過のためエラー
        CmbCheck_Num = 2
        GoTo Exit_rtn
    End If

    CmbCheck_Num = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmbCheck_Num = 9
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :顧客マスタ問い合わせ
'        MODULE_ID        :fncChkUSER_MAST
'        Parameter        :intKbn
'                                 0 = 存在チェック
'                                 1 = 顧客データ取得
'        CREATE_DATE      :2005/12/02
'
'==============================================================================*
Private Function fncChkUSER_MAST(ByVal intKbn As Integer) As Boolean

    Dim vCode    As Variant
    Dim vRet     As Variant
    Dim objRs    As Recordset
    Dim strSQL   As String
    Dim stWHERE  As String

    On Error GoTo fncChkUSER_MAST_Err

    fncChkUSER_MAST = False

    vCode = Nz(Me.cmb_顧客コード, "")

    If vCode = "" Then: fncChkUSER_MAST = True: GoTo fncChkUSER_MAST_Exit

    doCmd.Hourglass True
    doCmd.SetWarnings False

    Select Case intKbn
        Case 0  ' *** 顧客マスタ存在チェック ***
            ' 顧客名取得
            stWHERE = ""
            stWHERE = " USER_CODE =" & Val(vCode)
            vRet = Nz(DLookup("USER_NAME", "dbo_USER_MAST", stWHERE), "")

            If vRet = "" Then
                MsgBox "顧客マスタが登録されておりません", vbInformation, pstrCstKAKUNIN
                GoTo fncChkUSER_MAST_Exit
            End If

        Case 1  ' *** 顧客データ設定 ***
            Call subConnectDB(pstrBumonCd)

            strSQL = " SELECT USER_KANA, " & Chr(13)
            strSQL = strSQL & " USER_NAME " & Chr(13)
            strSQL = strSQL & "  FROM USER_MAST " & Chr(13)
            strSQL = strSQL & " WHERE USER_CODE = " & Val(vCode)
            Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

            If objRs.EOF = False Then
                Me.txt_顧客カナ = Nz(objRs.Fields("USER_KANA"), "")  ' 顧客名カナ
                Me.txt_顧客名 = Nz(objRs.Fields("USER_NAME"), "")    ' 顧客名
            Else
                Me.txt_顧客カナ = ""  ' 顧客名カナ
                Me.txt_顧客名 = ""    ' 顧客名
            End If

            objRs.Close
            Set objRs = Nothing

            ' DB切断
            Call subDisConnectDB
    End Select

    fncChkUSER_MAST = True

fncChkUSER_MAST_Exit:
    doCmd.Hourglass False
    doCmd.SetWarnings True

    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objKONT_DB Is Nothing Then
        ' DB切断
        Call subDisConnectDB
    End If
    Exit Function

fncChkUSER_MAST_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncChkUSER_MAST_Exit
End Function
'****************************  ended or program ********************************
