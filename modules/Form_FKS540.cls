VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS540"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　近隣ヤードリスト印刷
'   プログラムＩＤ　：　FKS540
'   作　成　日　　　：  2005/10/17
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
Private Const pstrCstRPT_ID       As String = "FKS540"

Private Const pstrCstKAKUNIN      As String = "確認"
Private Const pstrCstERROR        As String = "エラー"
Private Const pstrCstCOMMENT      As String = "なし"

Private Const pstrCstWK_TABLE_NM1 As String = "FKS540_W01"  ' ワークテーブル(メイン)
Private Const pstrCstWK_TABLE_NM2 As String = "FKS540_W02"  ' ワークテーブル(サブ)

Private pobjDB                    As Database               ' カレントDB
Private pobjKONT_DB               As Database               ' コンテナDB

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
'        MODULE_NAME      :txt_UPDAD_F_AfterUpdateイベント
'        MODULE_ID        :txt_UPDAD_F_AfterUpdate
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_F_AfterUpdate()
    Me.txt_UPDAD_F = Check_Date(Me.txt_UPDAD_F)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_F_BeforeUpdateイベント
'        MODULE_ID        :txt_UPDAD_F_BeforeUpdate
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_F_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_UPDAD_F, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_F_GotFocusイベント
'        MODULE_ID        :txt_UPDAD_F_GotFocus
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_F_GotFocus()
    Call Date_TextFormat(Me.txt_UPDAD_F, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_F_LostFocusイベント
'        MODULE_ID        :txt_UPDAD_F_LostFocus
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_F_LostFocus()
    Call Date_TextFormat(Me.txt_UPDAD_F, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_F_KeyPressイベント
'        MODULE_ID        :txt_UPDAD_F_KeyPress
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_UPDAD_F, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_T_AfterUpdateイベント
'        MODULE_ID        :txt_UPDAD_T_AfterUpdate
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_T_AfterUpdate()
    Me.txt_UPDAD_T = Check_Date(Me.txt_UPDAD_T)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_T_BeforeUpdateイベント
'        MODULE_ID        :txt_UPDAD_T_BeforeUpdate
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_T_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_UPDAD_T, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_T_GotFocusイベント
'        MODULE_ID        :txt_UPDAD_T_GotFocus
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_T_GotFocus()
    Call Date_TextFormat(Me.txt_UPDAD_T, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_T_LostFocusイベント
'        MODULE_ID        :txt_UPDAD_T_LostFocus
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_T_LostFocus()
        Call Date_TextFormat(Me.txt_UPDAD_T, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UPDAD_T_KeyPressイベント
'        MODULE_ID        :txt_UPDAD_T_KeyPress
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UPDAD_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_UPDAD_T, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YCODE_F_AfterUpdateイベント
'        MODULE_ID        :cmb_YCODE_F_AfterUpdate
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :cmb_YCODE_F_BeforeUpdateイベント
'        MODULE_ID        :cmb_YCODE_F_BeforeUpdate
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :cmb_YCODE_F_KeyPressイベント
'        MODULE_ID        :cmb_YCODE_F_KeyPress
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YCODE_F, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YCODE_T_AfterUpdateイベント
'        MODULE_ID        :cmb_YCODE_T_AfterUpdate
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :cmb_YCODE_T_BeforeUpdateイベント
'        MODULE_ID        :cmb_YCODE_T_BeforeUpdate
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :cmb_YCODE_T_KeyPressイベント
'        MODULE_ID        :cmb_YCODE_T_KeyPress
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YCODE_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YCODE_T, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Preview_Clickイベント
'        MODULE_ID        :cmd_Preview_Click
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Preview_Click()

    On Error GoTo cmd_Preview_Click_Err

    Dim strMsg As String
    Dim lngRc  As Long
    
    strMsg = "印刷内容を画面に出力します。よろしいですか？"
    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo, pstrCstKAKUNIN)
    If lngRc = vbNo Then GoTo cmd_Preview_Click_Exit

    doCmd.Hourglass True

    ' 帳票出力用データ作成
    If FncCreateWork = False Then GoTo cmd_Preview_Click_Exit

    ' 延長リストオープン
    doCmd.OpenReport "RKS540", acPreview, "", ""

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
'        MODULE_NAME      :cmd_Excel_Clickイベント
'        MODULE_ID        :cmd_Excel_Click
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    On Error GoTo cmd_Excel_Click_Err

'    ' 日付入力チェック
'    If fncInputCheck = False Then GoTo cmd_Excel_Click_Exit

    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, pstrCstKAKUNIN) = vbYes) Then

        doCmd.Hourglass True

        ' 帳票出力用データ作成
        If FncCreateWork = True Then
           doCmd.OutputTo acQuery, "FKS540_Q01", "MicrosoftExcel(*.xls)", "", True, ""
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
'        MODULE_NAME      :cmd_End_Clickイベント
'        MODULE_ID        :cmd_End_Click
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

    On Error Resume Next

    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Frm_Initイベント
'        MODULE_ID        :Frm_Init
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Frm_Init()
    With Me
        .cmb_YCODE_F = Null  ' ヤードコード(FROM)
        .txt_YCODE_F = Null  ' ヤード名(FROM)
        .cmb_YCODE_T = Null  ' ヤードコード(TO)
        .txt_YCODE_T = Null  ' ヤード名(TO)
        .txt_UPDAD_F = Null  ' 対象年月日(FROM)
        .txt_UPDAD_T = Null  ' 対象年月日(TO)
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :ワークテーブルのデータ削除
'        MODULE_ID        :fncDeleteWorkTable
'        Parameter        :strTableNm = テーブル名
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDeleteWorkTable(strTableNm As String) As Boolean

    On Error GoTo fncDeleteWorkTable_Err

    fncDeleteWorkTable = False

    ' ワークテーブル存在チェック
    If fncTableExist(strTableNm) = True Then

        ' ワークテーブルを削除
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
'        MODULE_NAME      :ACCESSテーブル存在チェック
'        MODULE_ID        :fncTableExist
'        Parameter        :IN strTableName = テーブル名
'        RETURN           :True=存在する False=存在しない
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :ワークテーブルのデータ削除
'        MODULE_ID        :fncDeleteWorkTable
'        Parameter        :strTableNm = テーブル名
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable(strTableNm As String) As Boolean

    Dim tdfNew    As TableDef
    Dim intCount  As Integer

    On Error GoTo fncCreateWorkTable_Err

    fncCreateWorkTable = False

    Set tdfNew = pobjDB.CreateTableDef(strTableNm)

    Select Case strTableNm
        Case pstrCstWK_TABLE_NM1  '「FVS540_W01」
            With tdfNew
                .Fields.Append .CreateField("REPORT_ID", dbText, 50)       ' レポートID
                .Fields.Append .CreateField("YARD_CODE", dbLong)           ' ヤードコード
                .Fields.Append .CreateField("YARD_NAME", dbText, 36)       ' ヤード名
                .Fields.Append .CreateField("YARD_ADDRESS", dbText, 100)   ' ヤード住所
                .Fields.Append .CreateField("YARD_NOTE", dbText, 36)       ' ヤード付近メモ
        
                For intCount = 0 To .Fields.Count - 1
                    If .Fields(intCount).Type = dbText Then
                        .Fields(intCount).AllowZeroLength = True
                    End If
                Next intCount
            End With

        Case pstrCstWK_TABLE_NM2  '「FVS540_W02」
            With tdfNew
                .Fields.Append .CreateField("YARD_CODE", dbLong)           ' ヤードコード
                .Fields.Append .CreateField("NYAR_NCODE", dbLong)          ' 近隣ヤードコード
                .Fields.Append .CreateField("NYAR_NAME", dbText, 36)       ' 近隣ヤード名
                .Fields.Append .CreateField("NYAR_ADDRESS", dbText, 100)   ' 予約ヤード住所
                .Fields.Append .CreateField("NYAR_NOTE", dbText, 36)       ' 近隣ヤード付近メモ
                .Fields.Append .CreateField("COMMENT", dbText, 10)         ' コメント
        
                For intCount = 0 To .Fields.Count - 1
                    If .Fields(intCount).Type = dbText Then
                        .Fields(intCount).AllowZeroLength = True
                    End If
                Next intCount
            End With
    End Select

    ' テーブルの追加
    pobjDB.TableDefs.Append tdfNew

    fncCreateWorkTable = True

    Exit Function

fncCreateWorkTable_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, pstrCstERROR
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルへ対象データのInsert
'        MODULE_ID        :fncCreateWork
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function FncCreateWork() As Boolean

    Dim strSQL           As String
    Dim strWhere         As String

    Dim objRsMain        As Recordset
    Dim objRsDetail      As Recordset
    Dim objWk1           As Recordset
    Dim objWk2           As Recordset

    Dim strAddress       As String

    Dim strAddr          As String   ' 住所
    Dim strRenraku       As String   ' 連絡先

    Dim varCode          As Variant

    Dim strYARD_CODE()   As String  ' ヤードコード
    Dim strYard_Name()   As String  ' ヤード名
    Dim strYARD_YUBINO() As String  ' 郵便番号
    Dim strYARD_ADDR_1() As String  ' 住所１
    Dim strYARD_ADDR_2() As String  ' 住所２
    Dim strYARD_NOTE()   As String  ' 付近メモ

    Dim intCount         As Integer

    On Error GoTo fncCreateWork_Err

    FncCreateWork = False

    doCmd.SetWarnings False

    ' DB接続
    Set pobjDB = CurrentDb

    ' コンテナDB接続
    If fncConnectDB = False Then GoTo fncCreateWork_Exit

    ' ワークテーブルを削除
    If fncDeleteWorkTable(pstrCstWK_TABLE_NM1) = False Then GoTo fncCreateWork_Err
    If fncDeleteWorkTable(pstrCstWK_TABLE_NM2) = False Then GoTo fncCreateWork_Err

    ' ワークテーブルを作成
    If fncCreateWorkTable(pstrCstWK_TABLE_NM1) = False Then GoTo fncCreateWork_Err
    If fncCreateWorkTable(pstrCstWK_TABLE_NM2) = False Then GoTo fncCreateWork_Err

    ' 「ヤードマスタ」、「近隣ヤードマスタ」よりデータメインデータ取得
    strSQL = ""
    strSQL = " SELECT DISTINCT YARD_MAST.YARD_CODE, " & Chr(13)
    strSQL = strSQL & " YARD_MAST.YARD_NAME, " & Chr(13)
    strSQL = strSQL & " YARD_MAST.YARD_YUBINO, " & Chr(13)
    strSQL = strSQL & " YARD_MAST.YARD_ADDR_1, " & Chr(13)
    strSQL = strSQL & " YARD_MAST.YARD_ADDR_2, " & Chr(13)
    strSQL = strSQL & " YARD_MAST.YARD_NOTE " & Chr(13)
    strSQL = strSQL & "  FROM YARD_MAST, " & Chr(13)
    strSQL = strSQL & "       NYAR_MAST " & Chr(13)

    ' Where句生成
    strWhere = fncCreateWhere

    If Nz(strWhere, "") <> "" Then
        ' 出力条件付加
        strSQL = strSQL & strWhere

        ' ヤードマスタと近隣ヤードマスタの結合条件
        strSQL = strSQL & " AND YARD_MAST.YARD_CODE = NYAR_MAST.NYAR_YCODE "
    Else
        ' ヤードマスタと近隣ヤードマスタの結合条件
        strSQL = strSQL & " WHERE YARD_MAST.YARD_CODE = NYAR_MAST.NYAR_YCODE "
    End If

    ' ソート
    strSQL = strSQL & " ORDER BY YARD_MAST.YARD_CODE " & Chr(13)  ' ヤードコード

    ' メインデータ読込
    Set objRsMain = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

    If objRsMain.EOF = True Then
        MsgBox "該当するデータがありません。", vbInformation, pstrCstERROR
        objRsMain.Close
        Set objRsMain = Nothing
        Exit Function
    End If

    ' 対象データ数取得
    objRsMain.MoveLast
    intCount = objRsMain.RecordCount
    objRsMain.MoveFirst

    ' 動的配列の宣言
    ReDim strYARD_CODE(intCount - 1)
    ReDim strYard_Name(intCount - 1)
    ReDim strYARD_YUBINO(intCount - 1)
    ReDim strYARD_ADDR_1(intCount - 1)
    ReDim strYARD_ADDR_2(intCount - 1)
    ReDim strYARD_NOTE(intCount - 1)

    ' 配列に値を格納
    intCount = 0
    Do Until objRsMain.EOF

        strYARD_CODE(intCount) = objRsMain.Fields("YARD_CODE")
        strYard_Name(intCount) = objRsMain.Fields("YARD_NAME")
        strYARD_YUBINO(intCount) = Nz(objRsMain.Fields("YARD_YUBINO"), "")
        strYARD_ADDR_1(intCount) = Nz(objRsMain.Fields("YARD_ADDR_1"), "")
        strYARD_ADDR_2(intCount) = Nz(objRsMain.Fields("YARD_ADDR_2"), "")
        strYARD_NOTE(intCount) = objRsMain.Fields("YARD_NOTE")

        intCount = intCount + 1  ' カウントアップ

        objRsMain.MoveNext
    Loop
    objRsMain.Close
    Set objRsMain = Nothing

    ' ワークテーブルへ書き込み
    strSQL = "SELECT * FROM FKS540_W01 "
    Set objWk1 = pobjDB.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

    ' ワークテーブルへ書き込み
    strSQL = "SELECT * FROM FKS540_W02 "
    Set objWk2 = pobjDB.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

    For intCount = 0 To UBound(strYARD_CODE)
        ' メイン部 -------------------------------------------------------------------------------------------
        With objWk1
            .AddNew

            .Fields("REPORT_ID") = pstrCstRPT_ID                   ' 帳票ID
            .Fields("YARD_CODE") = strYARD_CODE(intCount)          ' ヤードコード
            .Fields("YARD_NAME") = strYard_Name(intCount)          ' ヤード名

            strAddress = ""
            ' 郵便番号
            If Nz(strYARD_YUBINO(intCount), "") <> "" Then
                strAddress = strYARD_YUBINO(intCount)
            End If

            ' 住所１
            If Nz(strYARD_ADDR_1(intCount), "") <> "" Then
                If Nz(strAddress, "") <> "" Then strAddress = strAddress & Space(1)

                strAddress = strAddress & strYARD_ADDR_1(intCount)
            End If

            ' 住所２
            If Nz(strYARD_ADDR_2(intCount), "") <> "" Then
                If Nz(strAddress, "") <> "" Then strAddress = strAddress & Space(1)
    
                strAddress = strAddress & strYARD_ADDR_2(intCount)
            End If

            .Fields("YARD_ADDRESS") = strAddress                   ' 住所(郵便番号-住所１-住所2)

            .Fields("YARD_NOTE") = strYARD_NOTE(intCount)          ' 付近メモ

            .UPDATE
        End With
        ' ---------------------------------------------------------------------------------------------------

        ' 明細データ読込(メインヤード以外)
        strSQL = "SELECT NYAR_NCODE, " & Chr(13)          ' 近隣ヤードコード
        strSQL = strSQL & " YARD_NAME, " & Chr(13)        ' ヤード名
        strSQL = strSQL & " YARD_YUBINO, " & Chr(13)      ' 郵便番号
        strSQL = strSQL & " YARD_ADDR_1, " & Chr(13)      ' 住所１
        strSQL = strSQL & " YARD_ADDR_2, " & Chr(13)      ' 住所2
        strSQL = strSQL & " YARD_NOTE " & Chr(13)         ' 補足
        strSQL = strSQL & "  FROM YARD_MAST, " & Chr(13)  ' *** ヤードマスタ ***
        strSQL = strSQL & "       NYAR_MAST " & Chr(13)   ' *** 近隣ヤード ***
        strSQL = strSQL & " WHERE YARD_MAST.YARD_CODE  = NYAR_MAST.NYAR_NCODE " & Chr(13)
        strSQL = strSQL & "   AND NYAR_MAST.NYAR_YCODE = " & strYARD_CODE(intCount) & Chr(13)
        strSQL = strSQL & "   AND NYAR_MAST.NYAR_NCODE <> " & strYARD_CODE(intCount)
        Set objRsDetail = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)

        ' 明細部 ---------------------------------------------------------------------------------------------
        With objWk2
            If objRsDetail.EOF = True Then
                ' 近隣ヤードが存在しないとき
                .AddNew

                .Fields("YARD_CODE") = strYARD_CODE(intCount)                     ' ヤードコード
                .Fields("COMMENT") = pstrCstCOMMENT                               ' コメント

                .UPDATE

            Else
                Do Until objRsDetail.EOF
                    .AddNew

                    ' 近隣ヤードが存在するとき
                    ' レポート出力用
                    .Fields("YARD_CODE") = strYARD_CODE(intCount)                 ' ヤードコード
                    .Fields("NYAR_NCODE") = objRsDetail.Fields("NYAR_NCODE")      ' 近隣ヤードコード
                    .Fields("NYAR_NAME") = objRsDetail.Fields("YARD_NAME")        ' 近隣ヤード名

                    strAddress = ""
                    ' 郵便番号
                    If Nz(objRsDetail.Fields("YARD_YUBINO"), "") <> "" Then
                        strAddress = objRsDetail.Fields("YARD_YUBINO")
                    End If

                    ' 住所１
                    If Nz(objRsDetail.Fields("YARD_ADDR_1"), "") <> "" Then
                        If Nz(strAddress, "") <> "" Then strAddress = strAddress & Space(1)
    
                        strAddress = strAddress & objRsDetail.Fields("YARD_ADDR_1")
                    End If

                    ' 住所２
                    If Nz(objRsDetail.Fields("YARD_ADDR_2"), "") <> "" Then
                        If Nz(strAddress, "") <> "" Then strAddress = strAddress & Space(1)

                        strAddress = strAddress & objRsDetail.Fields("YARD_ADDR_2")
                    End If

                    .Fields("NYAR_ADDRESS") = strAddress                          ' 近隣ヤード住所(郵便番号-住所１-住所2)

                    .Fields("NYAR_NOTE") = objRsDetail.Fields("YARD_NOTE")        ' 近隣ヤード付近メモ

                    .Fields("COMMENT") = Null                                     ' コメント

                    .UPDATE

                    objRsDetail.MoveNext  ' 次近隣ヤードコードへ
                Loop
            End If

            objRsDetail.Close
            Set objRsDetail = Nothing
        End With
        ' ---------------------------------------------------------------------------------------------------
    Next intCount

    objWk1.Close
    Set objWk1 = Nothing

    objWk2.Close
    Set objWk2 = Nothing

    pobjDB.Close
    Set pobjDB = Nothing

    ' コンテナDB切断
    If fncDisConnectDB = False Then GoTo fncCreateWork_Exit

    FncCreateWork = True

fncCreateWork_Exit:
    If Not objRsMain Is Nothing Then
        objRsMain.Close
        Set objRsMain = Nothing
    End If
    If Not objRsDetail Is Nothing Then
        objRsDetail.Close
        Set objRsDetail = Nothing
    End If
    If Not objWk1 Is Nothing Then
        objWk1.Close
        Set objWk1 = Nothing
    End If
    If Not objWk2 Is Nothing Then
        objWk2.Close
        Set objWk2 = Nothing
    End If

    If Not pobjDB Is Nothing Then
        pobjDB.Close
        Set pobjDB = Nothing
    End If

    If Not pobjKONT_DB Is Nothing Then
        pobjKONT_DB.Close
        Set pobjKONT_DB = Nothing
    End If

    ' 配列の解放
    Erase strYARD_CODE()
    Erase strYard_Name()
    Erase strYARD_YUBINO()
    Erase strYARD_ADDR_1()
    Erase strYARD_ADDR_2()
    Erase strYARD_NOTE()

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
'        CREATE_DATE      :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWhere() As String

    Dim strWhere     As String
    Dim strUpd_From  As String
    Dim strUpd_To    As String

    With Me
        ' *** ヤードコードの大小チェック ***
        Call ChkCmbFromTo(1, .cmb_YCODE_F, .txt_YCODE_F, .cmb_YCODE_T, .txt_YCODE_T)

        ' *** 更新日の大小チェック ***
        Call ChkDateFromTo(.txt_UPDAD_F, .txt_UPDAD_T)

        ' *** 更新日From、更新日Toの取得 ***
        ' 更新日From
        strUpd_From = Nz(.txt_UPDAD_F, "")
        If strUpd_From <> "" Then strUpd_From = Replace(strUpd_From, "/", "")

        ' 更新日(To)
        strUpd_To = Nz(.txt_UPDAD_T, "")
        If strUpd_To <> "" Then strUpd_To = Replace(strUpd_To, "/", "")

        ' 更新日(or 作成日)
        If (Nz(strUpd_From, "") <> "" And Nz(strUpd_To, "") <> "") Then
            ' 更新日(From)：入力
            ' 更新日(To)  ：入力
            strWhere = " WHERE " & Chr(13)
            strWhere = strWhere & " CASE " & Chr(13)
            strWhere = strWhere & " WHEN NYAR_UPDAD IS NULL     AND NYAR_INSED BETWEEN '" & strUpd_From & "' AND '" & strUpd_To & "' THEN 1 " & Chr(13)  ' 作成日
            strWhere = strWhere & " WHEN NYAR_UPDAD IS NOT NULL AND NYAR_UPDAD BETWEEN '" & strUpd_From & "' AND '" & strUpd_To & "' THEN 1 " & Chr(13)  ' 更新日
            strWhere = strWhere & " END = 1 " & Chr(13)

        ElseIf (Nz(strUpd_From, "") <> "" And Nz(strUpd_To, "") = "") Then
            ' 更新日(From)：入力
            ' 更新日(To)  ：未入力
            strWhere = " WHERE " & Chr(13)
            strWhere = strWhere & " CASE " & Chr(13)
            strWhere = strWhere & " WHEN NYAR_UPDAD IS NULL     AND NYAR_INSED >= '" & strUpd_From & "' THEN 1 " & Chr(13)  ' 作成日
            strWhere = strWhere & " WHEN NYAR_UPDAD IS NOT NULL AND NYAR_UPDAD >= '" & strUpd_From & "' THEN 1 " & Chr(13)  ' 更新日
            strWhere = strWhere & " END = 1 " & Chr(13)

        ElseIf (Nz(strUpd_From, "") = "" And Nz(strUpd_To, "") <> "") Then
            ' 更新日(From)：未入力
            ' 更新日(To)  ：入力
            strWhere = " WHERE " & Chr(13)
            strWhere = strWhere & " CASE " & Chr(13)
            strWhere = strWhere & " WHEN NYAR_UPDAD IS NULL     AND NYAR_INSED <= '" & strUpd_To & "' THEN 1 " & Chr(13)    ' 作成日
            strWhere = strWhere & " WHEN NYAR_UPDAD IS NOT NULL AND NYAR_UPDAD <= '" & strUpd_To & "' THEN 1 " & Chr(13)    ' 更新日
            strWhere = strWhere & " END = 1 " & Chr(13)
        End If

        ' ヤードコード
        If (Nz(.cmb_YCODE_F, "") <> "" And Nz(.cmb_YCODE_T, "") <> "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：入力
            If Nz(strWhere, "") <> "" Then
                strWhere = strWhere & " AND YARD_CODE BETWEEN " & .cmb_YCODE_F & " AND " & .cmb_YCODE_T & Chr(13)
            Else
                strWhere = " WHERE YARD_CODE BETWEEN " & .cmb_YCODE_F & " AND " & .cmb_YCODE_T & Chr(13)
            End If

        ElseIf (Nz(.cmb_YCODE_F, "") <> "" And Nz(.cmb_YCODE_T, "") = "") Then
            ' ヤードコード(From)：入力
            ' ヤードコード(To)  ：未入力
            If Nz(strWhere, "") <> "" Then
                strWhere = strWhere & " AND YARD_CODE >= " & .cmb_YCODE_F & Chr(13)
            Else
                strWhere = " WHERE YARD_CODE >= " & .cmb_YCODE_F & Chr(13)
            End If

        ElseIf (Nz(.cmb_YCODE_F, "") = "" And Nz(.cmb_YCODE_T, "") <> "") Then
            ' ヤードコード(From)：未入力
            ' ヤードコード(To)  ：入力
            If Nz(strWhere, "") <> "" Then
                strWhere = strWhere & " AND YARD_CODE <= " & .cmb_YCODE_T & Chr(13)
            Else
                strWhere = " WHERE YARD_CODE <= " & .cmb_YCODE_T & Chr(13)
            End If
        End If

'        ' 更新日(or 作成日)
'        If (Nz(strUpd_From, "") <> "" And Nz(strUpd_To, "") <> "") Then
'            ' 更新日(From)：入力
'            ' 更新日(To)  ：入力
'            strWHERE = strWHERE & "AND IIF(NYAR_UPDAD IS NULL, NYAR_INSED BETWEEN '" & strUpd_From & "' AND '" & strUpd_To & "', " & Chr(13)  ' 作成日
'            strWHERE = strWHERE & "                            NYAR_UPDAD BETWEEN '" & strUpd_From & "' AND '" & strUpd_To & "') " & Chr(13)  ' 更新日
'
'        ElseIf (Nz(strUpd_From, "") <> "" And Nz(strUpd_To, "") = "") Then
'            ' 更新日(From)：入力
'            ' 更新日(To)  ：未入力
'            strWHERE = strWHERE & "AND IIF(NYAR_UPDAD IS NULL, NYAR_INSED >= '" & strUpd_From & "', " & Chr(13)  ' 作成日
'            strWHERE = strWHERE & "                            NYAR_UPDAD >= '" & strUpd_From & "') " & Chr(13)  ' 更新日
'
'        ElseIf (Nz(strUpd_From, "") = "" And Nz(strUpd_To, "") <> "") Then
'            ' 更新日(From)：未入力
'            ' 更新日(To)  ：入力
'            strWHERE = strWHERE & "AND IIF(NYAR_UPDAD IS NULL, NYAR_INSED <= '" & strUpd_To & "', " & Chr(13)  ' 作成日
'            strWHERE = strWHERE & "                            NYAR_UPDAD <= '" & strUpd_To & "') " & Chr(13)  ' 更新日
'        End If
    End With

    fncCreateWhere = strWhere
End Function

'==============================================================================*
'
'        MODULE_NAME      :名称取得
'        MODULE_ID        :fncGetName
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :日付入力チェック（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkDateFromTo
'        IN               :
'        CREATE_DATE      :2005/10/17
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
'        CREATE_DATE      :2005/10/17
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
'        CREATE_DATE      :2005/10/17
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
'        CREATE_DATE      :2005/10/17
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
'        CREATE_DATE      :2005/10/17
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
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        CREATE_DATE      :2005/10/17
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
'        CREATE_DATE      :2005/10/17
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
'****************************  ended or program ********************************
