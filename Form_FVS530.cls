VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS530"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ご紹介連絡表印刷
'   プログラムＩＤ　：　FVS530
'   作　成　日　　　：  2005/09/12
'   作　成　者　　　：  T.SUZUKI
'   Ver            ：  0.0
'**********************************************
'   更　新　日　　　：  2007/04/06
'   更　新　者　　　：  tajima
'                   ：  予約状態の初期値を 02:取置に
'   Ver             ：  0.1
'**********************************************
'   更　新　日　　　：  2017/04/01
'   更　新　者　　　：  K.SATO
'                   ：  ご紹介メール送信機能追加
'   Ver             ：  0.2
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const pstrCstRPT_ID     As String = "RVS530"

Private Const pstrCstKAKUNIN    As String = "確認"
Private Const pstrCstERROR      As String = "エラー"
Private Const pstrCstInputError As String = "を入力してください。"
Private Const pstrCstKetaOver   As String = "桁数オーバーしています。"
Private Const pstrCstNoDataErr  As String = "入力した受付番号は存在しません。"
Private Const pstrCstNotInputError As String = "をクリアしてください。"         'INSERT 2017/04/01 K.SATO START

Private Const pstrCstViewNew    As String = "新規紹介"
Private Const pstrCstViewOnly   As String = "紹介可能のみ"
Private Const pstrCstViewAll    As String = "全て"

Private pobjDB                  As Database
Private pobjKONT_DB             As Database

Private Const pcstr_YUKBN_00    As String = "00"    '予約未確定
 
'INSERT 2017/04/01 K.SATO START
'プログラムID
Private Const P_PROGRAM_ID      As String = "FVS530"
'INSERT 2017/04/01 K.SATO END

'==============================================================================*
'
'        MODULE_NAME      :メール送信ボタン押下時の処理
'        MODULE_ID        :cmd_Mail_Click
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Mail_Click()
    
    Dim strMsg        As String
    Dim lngRc         As Long

    On Error GoTo cmd_Mail_Click_Err

    ' 日付入力チェック
    If fncInputCheck = False Then
    Else
        If (Nz(Me.txt_MLDATE_F, "") <> "" Or Nz(Me.txt_MLDATE_T, "") <> "") Then
            MsgBox "メール送信日Fromまたはメール送信日To" & pstrCstNotInputError, vbInformation, pstrCstERROR
            Me.txt_MLDATE_F.SetFocus
            Exit Sub
        End If
    
        If pfncRptPrintFVS530(pintCstFVS530, pintCstMail, Nz(Me.cmb_PrintSyu, 0), pobjKONT_DB) = False Then Exit Sub
    End If

    Exit Sub

cmd_Mail_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear

End Sub


'==============================================================================*
'
'        MODULE_NAME      :Form_Activateイベント
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Loadイベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    ' カレントDB設定
    Set pobjDB = CurrentDb

    ' 初期処理
    Call Frm_Init
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Closeイベント
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Dim blnRTN As Boolean

    ' カレントDB切断
    pobjDB.Close
    Set pobjDB = Nothing

    ' コンテナDB切断
    blnRTN = fncDisConnectDB
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォーム初期化処理
'        MODULE_ID        :Frm_Init
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Frm_Init()

    Dim strSQL    As String
    Dim objRs     As Recordset
    Dim intCount  As Integer

    ' コンテナDB接続
    If fncConnectDB = False Then Exit Sub

    With Me
        .txt_INTRT_INSED_F = ""                          ' 紹介データ作成日(From)
        .txt_INTRT_INSED_F = Format(DATE, "YYYY/MM/DD")  ' システム日付
        .txt_INTRT_INSED_T = ""                          ' 紹介データ作成日(To)
        .cmb_Tantoc_F = Null                             ' 受付担当コード(From)
        .txt_Tantoc_F = Null                             ' 受付担当名(From)
        .cmb_Tantoc_T = Null                             ' 受付担当コード(To)
        .txt_Tantoc_T = Null                             ' 受付担当名(To)

        ' 印刷種別データ設定
        For intCount = 1 To 3
            Select Case intCount
                Case pintCstViewNew   ' 新規紹介
                    .cmb_PrintSyu.AddItem Item:=CStr(pintCstViewNew) & ";" & pstrCstViewNew

                Case pintCstViewOnly  ' 紹介可能のみ
                    .cmb_PrintSyu.AddItem Item:=CStr(pintCstViewOnly) & ";" & pstrCstViewOnly

                Case pintCstViewAll   ' 全て
                    .cmb_PrintSyu.AddItem Item:=CStr(pintCstViewAll) & ";" & pstrCstViewAll
            End Select
        Next intCount

        .cmb_PrintSyu = pintCstViewNew                   ' 印刷種別

        ' 連絡区分データ設定
        strSQL = "SELECT NAME_CODE, "
        strSQL = strSQL & " NAME_RYAK "
        strSQL = strSQL & " FROM NAME_MAST "
        strSQL = strSQL & " WHERE NAME_ID = '085' "
        strSQL = strSQL & " ORDER BY NAME_CODE"
        Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

        intCount = 0
        Do Until objRs.EOF

            If intCount = 0 Then
                ' 先頭行には固定列を設定
                .cmb_Renraku.AddItem Item:=CStr(pintCstRenraku) & ";全て"
            End If

            ' 名称コード、名称
            .cmb_Renraku.AddItem Item:=objRs.Fields(0) & ";" & objRs.Fields(1)

            intCount = intCount + 1

            objRs.MoveNext
        Loop
        objRs.Close
        Set objRs = Nothing

            ' 先頭行を選択
        .cmb_Renraku = .cmb_Renraku.ItemData(0)

        ' 予約受付状態(FROM)
        .cmb_Yukbn_F = "02" '2007/04/06 tajima 元→pcstr_YUKBN_00  ' 予約未確定
        Call cmb_Yukbn_F_AfterUpdate

        ' 予約受付状態(TO)
        .cmb_Yukbn_T = "02" '2007/04/06 tajima 元→pcstr_YUKBN_00  ' 予約未確定
        Call cmb_Yukbn_T_AfterUpdate

        .txt_INTRT_INSED_F.SetFocus
    End With
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
'        MODULE_NAME      :txt_INTRT_INSED_F_AfterUpdateイベント
'        MODULE_ID        :txt_INTRT_INSED_F_AfterUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_F_AfterUpdate()
    Me.txt_INTRT_INSED_F = Check_Date(Me.txt_INTRT_INSED_F)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_F_BeforeUpdateイベント
'        MODULE_ID        :txt_INTRT_INSED_F_BeforeUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_F_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_INTRT_INSED_F, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_F_GotFocusイベント
'        MODULE_ID        :txt_INTRT_INSED_F_GotFocus
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_F_GotFocus()
    Call Date_TextFormat(Me.txt_INTRT_INSED_F, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_F_LostFocusイベント
'        MODULE_ID        :txt_INTRT_INSED_F_LostFocus
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_F_LostFocus()
    Call Date_TextFormat(Me.txt_INTRT_INSED_F, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_F_KeyPressイベント
'        MODULE_ID        :txt_INTRT_INSED_F_KeyPress
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_INTRT_INSED_F, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_T_AfterUpdateイベント
'        MODULE_ID        :txt_INTRT_INSED_T_AfterUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_T_AfterUpdate()
    Me.txt_INTRT_INSED_T = Check_Date(Me.txt_INTRT_INSED_T)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_T_BeforeUpdateイベント
'        MODULE_ID        :txt_INTRT_INSED_T_BeforeUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_T_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_INTRT_INSED_T, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_T_GotFocusイベント
'        MODULE_ID        :txt_INTRT_INSED_T_GotFocus
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_T_GotFocus()
    Call Date_TextFormat(Me.txt_INTRT_INSED_T, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_T_LostFocusイベント
'        MODULE_ID        :txt_INTRT_INSED_T_LostFocus
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_T_LostFocus()
    Call Date_TextFormat(Me.txt_INTRT_INSED_T, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_INTRT_INSED_T_KeyPressイベント
'        MODULE_ID        :txt_INTRT_INSED_T_KeyPress
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_INTRT_INSED_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_INTRT_INSED_T, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_F_AfterUpdateイベント
'        MODULE_ID        :cmb_Tantoc_F_AfterUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_F_AfterUpdate()

    Dim varCode   As Variant
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_YCODE_F_AfterUpdate_Err

    varCode = Nz(Me.cmb_Tantoc_F, "")
    strCARG = ""
    If varCode = "" Then Exit Sub

    If Trim(varCode) <> "" Then
        ' 受付担当名取得
        strCARG = pfncGetTantoData(pobjDB, varCode)
    End If

    Me.cmb_Tantoc_F = Format(Me.cmb_Tantoc_F, "000")
    Me.txt_Tantoc_F = strCARG
    Me.cmb_Tantoc_T.SetFocus  ' 受付担当コード(TO)にフォーカスセット

    Exit Sub

cmb_YCODE_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_F_BeforeUpdateイベント
'        MODULE_ID        :cmb_Tantoc_F_BeforeUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_F_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_YCODE_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_Tantoc_F, "")
    If varCode = "" Then
        Me.txt_Tantoc_F = ""
        Exit Sub
    End If

    ' 受付担当名取得
    strCARG = pfncGetTantoData(pobjDB, varCode)
    If strCARG = "" Then
        MsgBox "該当する受付担当データがありません。", vbInformation, pstrCstKAKUNIN
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
'        MODULE_NAME      :cmb_Tantoc_F_KeyPressイベント
'        MODULE_ID        :cmb_Tantoc_F_KeyPress
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Tantoc_F, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_T_AfterUpdateイベント
'        MODULE_ID        :cmb_Tantoc_T_AfterUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_T_AfterUpdate()

    Dim varCode   As Variant
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_YCODE_T_AfterUpdate_Err

    varCode = Nz(Me.cmb_Tantoc_T, "")
    strCARG = ""
    If varCode = "" Then Exit Sub

    If Trim(varCode) <> "" Then
        ' 受付担当名取得
        strCARG = pfncGetTantoData(pobjDB, varCode)
    End If

    Me.cmb_Tantoc_T = Format(Me.cmb_Tantoc_T, "000")
    Me.txt_Tantoc_T = strCARG
    Me.cmb_PrintSyu.SetFocus  ' 印刷種別コンボボックスにフォーカスセット

    Exit Sub

cmb_YCODE_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_T_BeforeUpdateイベント
'        MODULE_ID        :cmb_Tantoc_T_BeforeUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_T_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_YCODE_T_BeforeUpdate_Err

    varCode = Nz(Me.cmb_Tantoc_T, "")
    If varCode = "" Then
        Me.txt_Tantoc_T = ""
        Exit Sub
    End If

    ' 受付担当名取得
    strCARG = pfncGetTantoData(pobjDB, varCode)
    If strCARG = "" Then
        MsgBox "該当する受担当データがありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_YCODE_T_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YCODE_T_KeyPressイベント
'        MODULE_ID        :cmb_YCODE_T_KeyPress
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Tantoc_T, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_PrintSyu_KeyPressイベント
'        MODULE_ID        :cmb_PrintSyu_KeyPress
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_PrintSyu_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UKNO_BeforeUpdateイベント
'        MODULE_ID        :txt_UKNO_BeforeUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UKNO_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_UKNO_BeforeUpdate_Err

    vCode = Nz(Me.txt_Ukno, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 11)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox pstrCstKetaOver, vbInformation, pstrCstERROR
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "受付番号" & pstrCstInputError, vbInformation, pstrCstERROR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    ' ファイルへの存在チェック
    If fncCheckYOUK_TRAN(vCode) = False Then
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If

    Exit Sub

txt_UKNO_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :受付番号検索ボタン押下時の処理
'        MODULE_ID        :cmd_UKNOSearch_Click
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_UKNOSearch_Click()
    Call psubSearchUKdata(Me.NAME, Me.txt_Ukno.NAME)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Renraku_BeforeUpdateイベント
'        MODULE_ID        :cmb_Renraku_BeforeUpdate
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Renraku_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strRyak  As String
    Dim strWhere As String

    On Error GoTo cmb_Renraku_BeforeUpdate_Err

    varCode = Nz(Me.cmb_Renraku, "")
    If (varCode = "" Or varCode = pintCstRenraku) Then
        Exit Sub
    End If

    ' 連絡区分名取得
    strRyak = fncGetName_MAST(1, "085", CStr(varCode))

    If strRyak = "" Then
        MsgBox "連絡区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_Renraku_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Renraku_KeyPressイベント
'        MODULE_ID        :cmb_Renraku_KeyPress
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Renraku_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_F_AfterUpdateイベント
'        MODULE_ID        :cmb_Yukbn_F_AfterUpdate
'        CREATE_DATE      :2005/11/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_F_AfterUpdate()

    Dim varCode   As Variant
    Dim strYukbn  As String
    Dim strWhere  As String

    On Error GoTo cmb_Yukbn_F_AfterUpdate_Err

    varCode = Nz(Me.cmb_Yukbn_F, "")
    strYukbn = ""
    If varCode = "" Then Exit Sub

    If Trim(varCode) <> "" Then
        ' 予約受付状態区分名取得
        strYukbn = fncGetName_MAST(0, "080", CStr(varCode))
    End If

    Me.cmb_Yukbn_F = Format(Me.cmb_Yukbn_F, "00")
    Me.txt_Yukbn_F = strYukbn
    Me.cmb_Yukbn_T.SetFocus  ' 予約受付状態区分(TO)にフォーカス移動

    Exit Sub

cmb_Yukbn_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_F_BeforeUpdateイベント
'        MODULE_ID        :cmb_Yukbn_F_BeforeUpdate
'        CREATE_DATE      :2005/11/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_F_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strName  As String
    Dim strWhere As String

    On Error GoTo cmb_Yukbn_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_Yukbn_F, "")
    If varCode = "" Then
        Me.txt_Yukbn_F = ""
        Exit Sub
    End If

    ' 予約受付状態区分名取得
    strName = fncGetName_MAST(0, "080", CStr(varCode))

    If strName = "" Then
        MsgBox "予約受付状態区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_Yukbn_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_F_KeyPressイベント
'        MODULE_ID        :cmb_Yukbn_F_KeyPress
'        CREATE_DATE      :2005/11/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Yukbn_F, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_T_AfterUpdateイベント
'        MODULE_ID        :cmb_Yukbn_T_AfterUpdate
'        CREATE_DATE      :2005/11/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_T_AfterUpdate()

    Dim varCode   As Variant
    Dim strYukbn  As String
    Dim strWhere  As String

    On Error GoTo cmb_Yukbn_T_AfterUpdate_Err

    varCode = Nz(Me.cmb_Yukbn_T, "")
    strYukbn = ""
    If varCode = "" Then Exit Sub

    If Trim(varCode) <> "" Then
        ' 予約受付状態区分名取得
        strYukbn = fncGetName_MAST(0, "080", CStr(varCode))
    End If

    Me.cmb_Yukbn_T = Format(Me.cmb_Yukbn_T, "00")
    Me.txt_Yukbn_T = strYukbn

    Exit Sub

cmb_Yukbn_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_T_BeforeUpdateイベント
'        MODULE_ID        :cmb_Yukbn_T_BeforeUpdate
'        CREATE_DATE      :2005/11/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_T_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strYukbn As String
    Dim strWhere As String

    On Error GoTo cmb_Yukbn_T_BeforeUpdate_Err

    varCode = Nz(Me.cmb_Yukbn_T, "")
    If varCode = "" Then
        Me.txt_Yukbn_T = ""
        Exit Sub
    End If

    ' 予約受付状態区分名取得
    strYukbn = fncGetName_MAST(0, "080", CStr(varCode))

    If strYukbn = "" Then
        MsgBox "予約受付状態区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_Yukbn_T_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_T_KeyPressイベント
'        MODULE_ID        :cmb_Yukbn_T_KeyPress
'        CREATE_DATE      :2005/11/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Yukbn_T, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :プレビューボタン押下時の処理
'        MODULE_ID        :cmd_Preview_Click
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Preview_Click()

    On Error GoTo cmd_Preview_Click_Err

    Dim strMsg        As String
    Dim lngRc         As Long

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Preview_Click_Exit
    
    strMsg = "印刷内容を画面に出力します。よろしいですか？"
    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo, pstrCstKAKUNIN)
    If lngRc = vbNo Then GoTo cmd_Preview_Click_Exit

    ' 帳票出力用データ出力処理
    If pfncRptPrintFVS530(pintCstFVS530, pintCstPreview, Nz(Me.cmb_PrintSyu, 0), pobjKONT_DB) = False Then Exit Sub


cmd_Preview_Click_Exit:
    Exit Sub

cmd_Preview_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo cmd_Preview_Click_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :EXCELボタン押下時の処理
'        MODULE_ID        :cmd_Excel_Click
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    Dim intDataCount  As Integer

    On Error GoTo cmd_Excel_Click_Err

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Excel_Click_Exit

    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, pstrCstKAKUNIN) = vbYes) Then
        ' 帳票出力用データ出力処理
        If pfncRptPrintFVS530(pintCstFVS530, pintCstExcel, Nz(Me.cmb_PrintSyu, 0), pobjKONT_DB) = False Then Exit Sub
    End If

cmd_Excel_Click_Exit:
    Exit Sub

cmd_Excel_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
    GoTo cmd_Excel_Click_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :終了ボタン押下時の処理
'        MODULE_ID        :cmd_End_Click
'        CREATE_DATE      :2005/09/12
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
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :2005/09/12
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
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'        CREATE_DATE      :2005/09/12
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
'        MODULE_NAME      :Date_TextFormat
'        機能             :日付データのフォーマット
'        IN               :
'        OUT              :
'        CREATE_DATE      :2005/09/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Date_TextFormat(ByRef ctlTxt As TextBox, ByVal lngMode As Long)

    Dim vCode As Variant

    vCode = Nz(ctlTxt, "")
    If vCode <> "" Then
        If lngMode = 0 Then
            ctlTxt = Format(vCode, "yy/mm/dd")
        Else
            ctlTxt = Format(vCode, "yyyy/mm/dd")
        End If
    End If
End Sub

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
'        MODULE_NAME      :
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2005/09/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    fncInputCheck = False

    With Me
        If (Nz(.txt_INTRT_INSED_F, "") = "" And Nz(.txt_INTRT_INSED_T, "") = "") Then
            MsgBox "紹介データ作成日Fromまたは紹介データ作成日To" & pstrCstInputError, vbInformation, pstrCstERROR
            .txt_INTRT_INSED_F.SetFocus
            Exit Function
        End If
    End With

    fncInputCheck = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :予約受付トラン存在チェック
'        MODULE_ID        :fncCheckYOUK_TRAN
'        CREATE_DATE      :2005/09/12
'        処理内容         ：予約受付トランの存在チェック
'        引数             ：varUKNO：受付番号
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckYOUK_TRAN(varUKNO As Variant) As Boolean

    Dim strSQL As String
    Dim vRet   As Variant

    On Error GoTo fncCheckYOUK_TRAN_Err

    fncCheckYOUK_TRAN = False

    If varUKNO <> "" Then
        strSQL = " YOUKT_UKNO = '" & varUKNO & "'"
        vRet = Nz(DLookup("YOUKT_UKNO", "dbo_YOUK_TRAN", strSQL), "")

        If vRet = "" Then
            MsgBox pstrCstNoDataErr, vbInformation, pstrCstKAKUNIN
            Exit Function
        End If
    End If

    fncCheckYOUK_TRAN = True

    Exit Function

fncCheckYOUK_TRAN_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Function

''==============================================================================*
''
''        MODULE_NAME      :名称マスタデータ取得
''        MODULE_ID        :fncGetNAME_MAST
''        CREATE_DATE      :2005/09/12
''        処理内容         ：予約受付トランの存在チェック
''        引数             ：varUKNO：受付番号
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncGetName_MAST(ByVal intKbna As Integer, ByVal strNAMEID As Variant, Optional strCODE As String = "") As String
'
'    Dim strWHERE As String
'    Dim strNAME  As String
'
'    On Error GoTo fncGetName_MAST_Err
'
'    fncGetName_MAST = ""
'
'    strWHERE = ""
'    strWHERE = strWHERE & " NAME_ID = '" & "085" & "'"
'
'    If Nz(Trim(strCODE), "") <> "" Then
'        strWHERE = strWHERE & " AND NAME_CODE = " & (strCODE)
'    End If
'
'    Select Case intKbna
'        Case 0  ' *** 名称 ***
'            strNAME = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWHERE), "")
'
'        Case 1  ' *** 略称 ***
'            strNAME = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWHERE), "")
'    End Select
'
'    fncGetName_MAST = strNAME
'
'    Exit Function
'
'fncGetName_MAST_Err:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectDB() As Boolean

    Dim strBumoCD     As String

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo fncConnectDB_Err

    fncConnectDB = False

    ' 部門コード取得
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumoCD & "'"), "")
    
    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, pstrCstRPT_ID
        Set pobjKONT_DB = Nothing
    Else
        Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumoCD))
    End If
    ' -----------------------------------------------------------------------------------------------------------

    fncConnectDB = True

    Exit Function

fncConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :fncDisConnectDB
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDisConnectDB() As Boolean

    On Error GoTo fncDisConnectDB_Err

    fncDisConnectDB = False

    ' コンテナDB
    pobjKONT_DB.Close
    Set pobjKONT_DB = Nothing

    fncDisConnectDB = True

    Exit Function

fncDisConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_F_BeforeUpdateイベント
'        MODULE_ID        :txt_MLDATE_F_BeforeUpdate
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_F_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_MLDATE_F, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_F_GotFocusイベント
'        MODULE_ID        :txt_MLDATE_F_GotFocus
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_F_GotFocus()
    Call Date_TextFormat(Me.txt_MLDATE_F, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_F_LostFocusイベント
'        MODULE_ID        :txt_MLDATE_F_LostFocus
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_F_LostFocus()
    Call Date_TextFormat(Me.txt_MLDATE_F, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_F_KeyPressイベント
'        MODULE_ID        :txt_MLDATE_F_KeyPress
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_MLDATE_F, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_T_BeforeUpdateイベント
'        MODULE_ID        :txt_MLDATE_T_BeforeUpdate
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_T_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_MLDATE_T, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_T_GotFocusイベント
'        MODULE_ID        :txt_MLDATE_T_GotFocus
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_T_GotFocus()
    Call Date_TextFormat(Me.txt_MLDATE_T, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_T_LostFocusイベント
'        MODULE_ID        :txt_MLDATE_T_LostFocus
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_T_LostFocus()
    Call Date_TextFormat(Me.txt_MLDATE_T, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MLDATE_T_KeyPressイベント
'        MODULE_ID        :txt_MLDATE_T_KeyPress
'        CREATE_DATE      :2017/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MLDATE_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_MLDATE_T, KeyAscii, 10)
End Sub
'****************************  ended or program ********************************
