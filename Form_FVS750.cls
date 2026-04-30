VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS750"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　延長リスト印刷
'   プログラムＩＤ　：　FVS750
'   作　成　日　　　：  2005/08/08
'   作　成　者　　　：  T.SUZUKI
'   Ver            ：  0.0
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const pstrCstRPT_ID   As String = "RVS750"

Private Const pstrCstKAKUNIN  As String = "確認"
Private Const pstrCstERROR    As String = "エラー"

Public pobjDB                 As Database

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2005/08/08
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
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    ' 初期処理
    Call Frm_Init
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_F_AfterUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_F_AfterUpdate()
    Me.txt_CYDATE_F = Check_Date(Me.txt_CYDATE_F)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_F_BeforeUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_F_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_CYDATE_F, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_F_GotFocus
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_F_GotFocus()
    Call Date_TextFormat(Me.txt_CYDATE_F, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_F_LostFocus
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_F_LostFocus()
    Call Date_TextFormat(Me.txt_CYDATE_F, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_F_KeyPress
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_CYDATE_F, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_T_AfterUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_T_AfterUpdate()
    Me.txt_CYDATE_T = Check_Date(Me.txt_CYDATE_T)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_T_BeforeUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_T_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_CYDATE_T, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_T_GotFocus
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_T_GotFocus()
    Call Date_TextFormat(Me.txt_CYDATE_T, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_T_GotFocus
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_T_LostFocus()
        Call Date_TextFormat(Me.txt_CYDATE_T, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_CYDATE_T_KeyPress
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CYDATE_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_CYDATE_T, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_KAICD_F_AfterUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAICD_F_AfterUpdate()

    Dim varCode   As Variant
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_KAICD_F_AfterUpdate_Err

    varCode = Nz(Me.cmb_KAICD_F, "")
    strCARG = ""
    If varCode = "" Then Exit Sub

    ' 解約区分名取得
    strCARG = fncGetName("089", varCode)

    Me.cmb_KAICD_F = Format(Me.cmb_KAICD_F, "00")
    Me.txt_KAICD_F = strCARG
    Me.cmb_KAICD_T.SetFocus  ' 解約区分(TO)にフォーカスセット

    Exit Sub

cmb_KAICD_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_KAICD_F_BeforeUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAICD_F_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_KAICD_F_BeforeUpdate_Err

    varCode = Nz(Me.cmb_KAICD_F, "")
    If varCode = "" Then
        Me.txt_KAICD_F = ""
        Exit Sub
    End If

    ' 解約区分名取得
    strCARG = fncGetName("089", varCode)
    If strCARG = "" Then
        MsgBox "該当する解約区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_KAICD_F_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_KAICD_F_KeyPress
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAICD_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_KAICD_F, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_KAICD_T_AfterUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAICD_T_AfterUpdate()

    Dim varCode   As Variant
    Dim strCARG   As String
    Dim strWhere  As String

    On Error GoTo cmb_KAICD_T_AfterUpdate_Err

    varCode = Nz(Me.cmb_KAICD_T, "")
    strCARG = ""
    If varCode = "" Then Exit Sub

    ' 解約区分名取得
    strCARG = fncGetName("089", varCode)

    Me.cmb_KAICD_T = Format(Me.cmb_KAICD_T, "00")
    Me.txt_KAICD_T = strCARG
    Me.cmb_YCODE_F.SetFocus  ' ヤードコード(FROM)にフォーカスセット

    Exit Sub

cmb_KAICD_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_KAICD_T_BeforeUpdate
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAICD_T_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant
    Dim strCARG  As String

    On Error GoTo cmb_KAICD_T_BeforeUpdate_Err

    varCode = Nz(Me.cmb_KAICD_T, "")
    If varCode = "" Then
        Me.txt_KAICD_T = ""
        Exit Sub
    End If

    ' 解約区分名取得
    strCARG = fncGetName("089", varCode)
    If strCARG = "" Then
        MsgBox "該当する解約区分がありません。", vbInformation, pstrCstKAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

cmb_KAICD_T_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_KAICD_T_KeyPress
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KAICD_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_KAICD_T, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_F_AfterUpdate
'        CREATE_DATE      :2005/08/08
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
'        MODULE_ID        :cmb_YCODE_F_BeforeUpdate
'        CREATE_DATE      :2005/08/08
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
'        MODULE_ID        :cmb_YCODE_F_KeyPress
'        CREATE_DATE      :2005/08/08
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
'        CREATE_DATE      :2005/08/08
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
'        CREATE_DATE      :2005/08/08
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
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_YCODE_T_KeyPress
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YCODE_T, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Preview_Click
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Preview_Click()

    On Error GoTo cmd_Preview_Click_Err

    Dim strMsg As String
    Dim lngRc  As Long

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Preview_Click_Exit
    
    strMsg = "印刷内容を画面に出力します。よろしいですか？"
    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo, pstrCstKAKUNIN)
    If lngRc = vbNo Then GoTo cmd_Preview_Click_Exit

    doCmd.Hourglass True

    ' 帳票出力用データ作成
    If FncCreateWork = False Then GoTo cmd_Preview_Click_Exit

    ' 延長リストオープン
    doCmd.OpenReport "RVS750", acPreview, "", ""

    GoTo cmd_Preview_Click_Exit

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
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    On Error GoTo cmd_Excel_Click_Err

    ' 日付入力チェック
    If fncInputCheck = False Then GoTo cmd_Excel_Click_Exit

    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, pstrCstKAKUNIN) = vbYes) Then

        doCmd.Hourglass True

        ' 帳票出力用データ作成
        If FncCreateWork = True Then
           doCmd.OutputTo acQuery, "FVS750_Q01", "MicrosoftExcel(*.xls)", "", True, ""
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
'        CREATE_DATE      :2005/08/08
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
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Frm_Init()

    With Me
        .txt_CYDATE_F = Null  ' 対象年月日(FROM)
        .txt_CYDATE_T = Null  ' 対象年月日(TO)
        .cmb_KAICD_F = Null   ' 解約区分(FROM)
        .txt_KAICD_F = Null   ' 解約区分名(FROM)
        .cmb_KAICD_T = Null   ' 解約区分(TO)
        .txt_KAICD_T = Null   ' 解約区分名(TO)
        .cmb_YCODE_F = Null   ' ヤードコード(FROM)
        .txt_YCODE_F = Null   ' ヤード名(FROM)
        .cmb_YCODE_T = Null   ' ヤードコード(TO)
        .txt_YCODE_T = Null   ' ヤード名(TO)
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :2005/08/08
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
'        MODULE_NAME      :
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    fncInputCheck = False

    With Me
        If (Nz(.txt_CYDATE_F, "") = "" And Nz(.txt_CYDATE_T, "") = "") Then
            MsgBox "対象年月日Fromまたは対象年月日Toを入力して下さい。", vbInformation, pstrCstERROR
            .txt_CYDATE_F.SetFocus
            Exit Function
        End If
    End With

    fncInputCheck = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルのデータ削除
'        MODULE_ID        :fncCreateWork
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDeleteWork() As Boolean

    Dim strSQL As String

    On Error GoTo fncDeleteWork_Err

    fncDeleteWork = False

    ' ワークテーブルの内容を削除
    strSQL = "DELETE FROM FVS750_W01 ; "
    pobjDB.Execute strSQL, dbFailOnError

    fncDeleteWork = True

    Exit Function

fncDeleteWork_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルへ対象データのInsert
'        MODULE_ID        :fncCreateWork
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function FncCreateWork() As Boolean

    Dim strSQL      As String
    Dim strWhere    As String

    Dim strSQLWK    As String

    Dim objRs       As Recordset
    Dim objWk       As Recordset

    Dim varBumoCD   As Variant  ' 部門ｺｰﾄﾞ
    Dim varBumoNm   As Variant  ' 部門名

    Dim strAddr     As String   ' 住所
    Dim strRenraku  As String   ' 連絡先

    Dim varCode     As Variant

    On Error GoTo fncCreateWork_Err

    FncCreateWork = False

    doCmd.SetWarnings False

    ' DB接続
    Set pobjDB = CurrentDb

    ' ワークテーブルの内容を削除
    If fncDeleteWork = False Then GoTo fncCreateWork_Err

    ' 部門名取得
    Call subGetBumonName(pobjDB, varBumoCD, varBumoNm)

    ' コンテナ契約ファイルよりデータ取得
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " CARG_CYDATE, " & Chr(13)        ' 解約予定日
    strSQL = strSQL & " CARG_KAICD, " & Chr(13)         ' 解約区分
    strSQL = strSQL & " CARG_UCODE," & Chr(13)          ' 顧客コード
    strSQL = strSQL & " USER_NAME," & Chr(13)           ' 顧客名
    strSQL = strSQL & " CARG_YCODE, " & Chr(13)         ' ヤードコード
    strSQL = strSQL & " CARG_NO, " & Chr(13)            ' コンテナ番号
    strSQL = strSQL & " CARG_KAITANTO," & Chr(13)       ' 解約受付担当
    strSQL = strSQL & " CARG_KAIUKEDATE," & Chr(13)     ' 解約受付日
    strSQL = strSQL & " CARG_KYDATE, " & Chr(13)        ' 解約日付
    strSQL = strSQL & " CARG_RENTKG, " & Chr(13)        ' 月額料金
    strSQL = strSQL & " CARG_HOSYICD, " & Chr(13)       ' 保証区分
    strSQL = strSQL & " CARG_NIRCD, " & Chr(13)         ' 荷入区分
    strSQL = strSQL & " USER_ADR_1, " & Chr(13)         ' 住所1
    strSQL = strSQL & " USER_ADR_2, " & Chr(13)         ' 住所2
    strSQL = strSQL & " USER_ADR_3, " & Chr(13)         ' 住所3
    strSQL = strSQL & " USER_TEL, " & Chr(13)           ' 電話番号
    strSQL = strSQL & " USER_FAX, " & Chr(13)           ' FAX番号
    strSQL = strSQL & " USER_KEITAI " & Chr(13)         ' 携帯番号
    strSQL = strSQL & " FROM ("
    strSQL = strSQL & "   dbo_CARG_FILE INNER JOIN dbo_YARD_MAST " & Chr(13)
    strSQL = strSQL & "   ON dbo_CARG_FILE.CARG_YCODE = dbo_YARD_MAST.YARD_CODE) " & Chr(13)
    strSQL = strSQL & "   INNER JOIN dbo_USER_MAST " & Chr(13)
    strSQL = strSQL & "   ON dbo_CARG_FILE.CARG_UCODE = dbo_USER_MAST.USER_CODE " & Chr(13)
    strSQL = strSQL & "  WHERE CARG_AGRE = 3 " & Chr(13) ' 契約状態：3(解約延長)
    strSQL = strSQL & "    AND CARG_KAICD IS NOT NULL " & Chr(13)

    ' Where句生成
    strWhere = fncCreateWhere

    If strWhere <> "" Then
        strSQL = strSQL & strWhere
    End If

    ' ソート
    strSQL = strSQL & " ORDER BY CARG_YCODE, " & Chr(13)       ' ヤードコード
    strSQL = strSQL & "          CARG_NO, " & Chr(13)          ' コンテナ番号
    strSQL = strSQL & "          CARG_MNT_REQDATE " & Chr(13)  ' メンテ依頼日

    ' コンテナ契約ファイル読込
    Set objRs = pobjDB.OpenRecordset(strSQL, dbReadOnly)

    If objRs.RecordCount = 0 Then
        MsgBox "該当するデータがありません。", vbInformation, pstrCstERROR
        objRs.Close
        Set objRs = Nothing
        Exit Function
    End If

    ' ワークテーブルへ書き込み
    strSQLWK = "SELECT * FROM FVS750_W01 "
    Set objWk = pobjDB.OpenRecordset(strSQLWK, dbOpenDynaset, dbAppendOnly)

    Do Until objRs.EOF
        With objWk
            .AddNew

            .Fields("REPORT_ID") = pstrCstRPT_ID                          ' 帳票ID
            .Fields("BUMOM_BUMOC") = varBumoCD                            ' 部門コード
            .Fields("BUMOM_BUMON") = varBumoNm                            ' 部門名
            .Fields("CARG_CYDATE") = objRs.Fields("CARG_CYDATE")          ' 解約予定日

            ' 解約区分
            .Fields("CARG_KAICD") = objRs.Fields("CARG_KAICD")

            ' 解約区分名取得
            If Nz(objRs.Fields("CARG_KAICD"), "") <> "" Then
                varCode = fncGetName("089", objRs.Fields("CARG_KAICD"))
                .Fields("CARG_KAICD_NAME") = varCode                      ' 解約区分名
            End If

            .Fields("CARG_UCODE") = objRs.Fields("CARG_UCODE")            ' 顧客コード
            .Fields("USER_NAME") = objRs.Fields("USER_NAME")              ' 顧客名
            .Fields("CARG_YCODE") = objRs.Fields("CARG_YCODE")            ' ヤードコード
            .Fields("CARG_NO") = objRs.Fields("CARG_NO")                  ' コンテナ番号

            ' 解約受付担当
            varCode = Nz(objRs.Fields("CARG_KAITANTO"), "")
            If varCode <> "" Then
                .Fields("CARG_KAITANTO") = varCode

                ' 解約受付担当者名
                .Fields("CARG_KAITANTO_NAME") = fncGetTantoData(varCode)
            End If

            .Fields("CARG_KAIUKEDATE") = objRs.Fields("CARG_KAIUKEDATE")  ' 解約受付日
            .Fields("CARG_KYDATE") = objRs.Fields("CARG_KYDATE")          ' 解約日付
            .Fields("CARG_RENTKG") = objRs.Fields("CARG_RENTKG")          ' 月額料金

            ' 保証区分
            .Fields("CARG_HOSYICD") = objRs.Fields("CARG_HOSYICD")

            ' 保証区分名取得
            If Nz(objRs.Fields("CARG_HOSYICD"), "") <> "" Then
                varCode = fncGetName("200", objRs.Fields("CARG_HOSYICD"))
                .Fields("CARG_HOSYICD_NAME") = varCode                    ' 保証区分名
            End If

            ' 荷入区分
            Select Case Nz(objRs.Fields("CARG_NIRCD"), "")
                Case ""
                    .Fields("CARG_NIRCD") = "有・無"
                Case 0, 1
                    varCode = fncGetName("093", objRs.Fields("CARG_NIRCD"))
                    .Fields("CARG_NIRCD") = varCode
                Case Else
                    .Fields("CARG_NIRCD") = Null
            End Select

            ' 契約者住所
            strAddr = ""
            strAddr = objRs.Fields("USER_ADR_1") & objRs.Fields("USER_ADR_2") & objRs.Fields("USER_ADR_3")
            .Fields("CARG_ADDR") = strAddr

            ' 連絡先
            strRenraku = ""
            ' 電話番号
            If Nz(objRs.Fields("USER_TEL"), "") <> "" Then
                strRenraku = "TEL:" & objRs.Fields("USER_TEL")
            Else
                strRenraku = Chr(13) & Chr(10)
            End If

            ' FAX番号
            If Nz(objRs.Fields("USER_FAX"), "") <> "" Then
                strRenraku = strRenraku & Chr(13) & Chr(10) & "FAX:" & objRs.Fields("USER_FAX")
            Else
                strRenraku = strRenraku & Chr(13) & Chr(10)
            End If

            ' 携帯番号
            If Nz(objRs.Fields("USER_KEITAI"), "") <> "" Then
                strRenraku = strRenraku & Chr(13) & Chr(10) & "CEL:" & objRs.Fields("USER_KEITAI")
            Else
                strRenraku = strRenraku & Chr(13) & Chr(10)
            End If
            .Fields("CARG_RENRAKU") = strRenraku

            .UPDATE
        End With

        objRs.MoveNext
    Loop

    FncCreateWork = True

fncCreateWork_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    If Not objWk Is Nothing Then
        objWk.Close
        Set objWk = Nothing
    End If
    Exit Function

fncCreateWork_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncCreateWork_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :Where句生成
'        MODULE_ID        :fncCreateWhere
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWhere() As String

    Dim strSwap  As String
    Dim strWhere As String

    With Me
        ' *** 解約予定日の大小チェック ***
        Call ChkDateFromTo(txt_CYDATE_F, txt_CYDATE_T)

        ' *** 解約区分コードの大小チェック ***
        Call ChkCmbFromTo(0, .cmb_KAICD_F, .txt_KAICD_F, .cmb_KAICD_T, .txt_KAICD_T)

        ' *** ヤードコードの大小チェック ***
        Call ChkCmbFromTo(1, .cmb_YCODE_F, .txt_YCODE_F, .cmb_YCODE_T, .txt_YCODE_T)

        ' 解約予定日
        If (Nz(.txt_CYDATE_F, "") <> "" And Nz(.txt_CYDATE_T, "") <> "") Then
            ' 解約予定日(From)：入力
            ' 解約予定日(To)  ：入力
            strWhere = " AND CARG_CYDATE BETWEEN CDate(#" & .txt_CYDATE_F & "#) And CDATE(#" & .txt_CYDATE_T & "#)"

        ElseIf (Nz(.txt_CYDATE_F, "") <> "" And Nz(.txt_CYDATE_T, "") = "") Then
            ' 解約予定日(From)：入力
            ' 解約予定日(To)  ：未入力
            strWhere = " AND CARG_CYDATE >= CDate(#" & .txt_CYDATE_F & "#)"

        ElseIf (Nz(.txt_CYDATE_F, "") = "" And Nz(.txt_CYDATE_T, "") <> "") Then
            ' 解約予定日(From)：未入力
            ' 解約予定日(To)  ：入力
            strWhere = " AND CARG_CYDATE <= CDate(#" & txt_CYDATE_T & "#)"
        End If

        ' *** 解約区分コード ***
        If (Nz(.cmb_KAICD_F, "") <> "" And Nz(.cmb_KAICD_T, "") <> "") Then
            ' 解約区分コード(From)：入力
            ' 解約区分コード(To)  ：入力
            strWhere = strWhere & " AND CARG_KAICD BETWEEN " & .cmb_KAICD_F & " AND " & .cmb_KAICD_T

        ElseIf (Nz(.cmb_KAICD_F, "") <> "" And Nz(.cmb_KAICD_T, "") = "") Then
            ' 解約区分コード(From)：入力
            ' 解約区分コード(To)  ：未入力
            strWhere = strWhere & " AND CARG_KAICD >= " & .cmb_KAICD_F

        ElseIf (Nz(.cmb_KAICD_F, "") = "" And Nz(.cmb_KAICD_T, "") <> "") Then
            ' 解約区分コード(From)：未入力
            ' 解約区分コード(To)  ：入力
            strWhere = strWhere & " AND CARG_KAICD <= " & .cmb_KAICD_T
        End If

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
    End With

    fncCreateWhere = strWhere
End Function

'==============================================================================*
'
'        MODULE_NAME      :名称取得
'        MODULE_ID        :fncGetName
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetName(strNAME_ID As String, varNAME_CODE As Variant) As String

    Dim strWhere  As String
    Dim strName   As String

    On Error GoTo fncGetName_Err

    fncGetName = ""

    ' 名称マスタより略称を取得
    strWhere = " NAME_ID = '" & strNAME_ID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & varNAME_CODE

    strName = ""
    strName = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")

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
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :subGetBumonName
'        IN               :objDB = DB接続 /strBumonCd = 部門ｺｰﾄﾞ /strBumonNm = 部門名
'        CREATE_DATE      :2005/08/08
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetBumonName(ByRef objDb As Database, ByRef varBumonCd As Variant, ByRef varBumonNm As Variant)

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo subGetBumonName_Err

    strSQL = "SELECT "
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMOC, "  ' 部門コード
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMON "   ' 部門名称
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "dbo_CONT_MAST INNER JOIN BUMO_MAST ON "
    strSQL = strSQL & "dbo_CONT_MAST.CONT_BUMOC = BUMO_MAST.BUMOM_BUMOC; "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            varBumonCd = .Fields("BUMOM_BUMOC")
            varBumonNm = .Fields("BUMOM_BUMON")
        Else
            varBumonCd = Null
            varBumonNm = Null
        End If
    End With

subGetBumonName_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Sub

subGetBumonName_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    GoTo subGetBumonName_Exit
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :日付入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkDateFromTo
'        IN               :
'        CREATE_DATE      :2005/08/08
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkDateFromTo(ByRef ctlTextF As TextBox, ByRef ctlTextT As TextBox)

    Dim strDayF As String
    Dim strDayT As String
    Dim strSwap As String

    If ctlTextF <> "" And ctlTextT <> "" Then
        strDayF = Format(ctlTextF, "yyyy/mm/dd")  ' 解約予定日(From)
        strDayT = Format(ctlTextT, "yyyy/mm/dd")  ' 解約予定日(To)

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
'        CREATE_DATE      :2005/08/08
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkCmbFromTo(ByRef intKbn As Integer, ByRef ctlcmbF As ComboBox, ByRef ctlTextF As TextBox, _
                          ByRef ctlcmbT As ComboBox, ByRef ctlTextT As TextBox)

    Dim strCmbF  As String
    Dim strcmbT  As String
    Dim strSwap  As String

    strSwap = ""
    If Trim(ctlcmbF) <> "" And Trim(ctlcmbT) <> "" Then
        Select Case intKbn
            Case 0  ' *** 解約区分 ***
                strCmbF = Format(ctlcmbF, "00")      ' 解約区分コード(From)
                strcmbT = Format(ctlcmbT, "00")      ' 解約区分コード(To)

            Case 1  ' *** ヤードコード ***
                strCmbF = Format(ctlcmbF, "000000")  ' ヤードコード(From)
                strcmbT = Format(ctlcmbT, "000000")  ' ヤードコード(To)                a
        End Select

        If strCmbF > strcmbT Then
            ' 大小入力エラー時
            ' *** ヤードコード
            strSwap = strCmbF
            strCmbF = strcmbT
            strcmbT = strSwap

            '　参照元へ再セット
            ctlcmbF = strCmbF
            ctlcmbT = strcmbT

            ' 名称
            ' *** ヤード名
            strSwap = ""
            strCmbF = ""
            strcmbT = ""

            strSwap = ctlTextF
            strCmbF = ctlTextT
            strcmbT = strSwap

            '　参照元へ再セット
            ctlTextF = strCmbF
            ctlTextT = strcmbT
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'        CREATE_DATE      :2005/08/08
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
    GoTo DateCheck_Exit
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
    CmBCheck_Date = 9
    Err.Clear
    GoTo CmBCheck_Date_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :Date_TextFormat
'        機能             :日付データのフォーマット
'        IN               :
'        OUT              :
'        CREATE_DATE      :2005/08/08
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
'        MODULE_NAME      :fncGetYardData
'        機能             :ヤードマスタよりヤード名取得
'        IN               :
'        OUT              :
'        CREATE_DATE      :2005/08/08
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetYardData(varCode As Variant) As String

    Dim strWhere As String
    Dim strCARG  As String

    On Error GoTo fncGetYardData_Err

    fncGetYardData = ""

    strWhere = strWhere & " YARD_CODE = " & varCode & ""
    strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")

    fncGetYardData = strCARG

    Exit Function

fncGetYardData_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetTantoData
'        機能             :担当者マスタ情報取得
'        IN               :vTanto = 担当者ｺｰﾄﾞ / vTanNm = 担当者名
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetTantoData(ByRef vTanto As Variant) As String

    Dim strSQL  As String
    Dim objRs   As Recordset

    fncGetTantoData = ""

    strSQL = "SELECT "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC, "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTN "
    strSQL = strSQL & "FROM TANT_MAST INNER JOIN "
    strSQL = strSQL & "dbo_CONT_MAST ON "
    strSQL = strSQL & "TANT_MAST.TANTM_BUMOC = dbo_CONT_MAST.CONT_BUMOC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"

    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Exit Function
    Else
        If Nz(objRs("TANTM_TANTN"), "") <> "" Then fncGetTantoData = objRs("TANTM_TANTN")
    End If

fncGetTantoData_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Function

fncGetTantoData_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo fncGetTantoData_Exit
End Function
'****************************  ended or program ********************************
