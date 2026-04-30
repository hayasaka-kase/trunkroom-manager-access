VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS480"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　倉庫管理システム
'**********************************************
'   プログラム名　　：　顧客マスタ一覧表
'   プログラムＩＤ　：
'   作　成　日　　　：
'   作　成　者　　　：　αシステム
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：2004/11/24
'   修　正　者　　　：N.MIURA
'   修　正　内　容　：フォーム名変更
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：2006/01/25
'   修　正　者　　　：T.SUZUKI
'   修　正　内　容　：印刷出力項目に『新顧客コード』を追加
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：2013/06/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：終了メッセージを表示させないように修正
'   Ver             ：  0.3
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID            As String = "FKS480"

Private Const strMsg_kaku        As String = "確認"
Private Const strMsg_err         As String = "エラー"

Private Const strMSGNoDataError  As String = "存在しないコードです。"

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate イベント
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open イベント
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    '＊＊＊画面の最大化＊＊＊
    doCmd.Maximize

    '＊＊＊画面表示クリア＊＊＊
    Me.cmb_CODEFROM = ""
    Me.cmb_CODETO = ""
    Me.txt_CODEFROM = ""
    Me.txt_CODETO = ""
    Me.txt_更新日FROM = ""
    Me.txt_更新日TO = ""
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Error イベント
'        MODULE_ID        :Form_Error
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
        Case 2113
            Response = acDataErrContinue
            MsgBox "数字を入力してください。", 0 + 64, "確認"
            SendKeys "{ESC}"
            Exit Sub
        Case 2279
            Response = acDataErrContinue
            MsgBox "正しく入力してください。", 0 + 64, "確認"
            SendKeys "{ESC}"
            Exit Sub
        Case 3316
            Response = acDataErrContinue
            SendKeys "{ESC}{ESC}"
            Exit Sub
        Case Else
            MsgBox DataErr
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown イベント
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Err_Form_KeyDown

    'SHIFT、CTRL、ALTのいずれかが押されている場合は何もしない
    If Shift > 0 Then
        Exit Sub
    End If

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Me.cmd_印刷.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

Exit_Form_KeyDown:
    Exit Sub

Err_Form_KeyDown:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_KeyDown
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload イベント
'        MODULE_ID        :Form_Unload
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)

    Dim strMsg  As String
    Dim intAns  As Integer
    Dim dbMydb  As Database
    Dim strSQL  As String

    On Error GoTo Err_Form_Unload

    '' 2013/06/03 M.HONDA START
'    strMsg = "終了します。" '& vbCr
'    strMsg = strMsg & "よろしいですか？"
'    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)
'
'    If intAns = vbNo Then
'        Cancel = True
'        Me.cmb_CODEFROM.SetFocus
'        GoTo Exit_Form_Unload
'    Else
'        '*** SQLSERVERのテーブルよりJETのワークテーブルに取り込み ***
'        Set dbMydb = currentdb()
'
'        '*** ワーク削除 ***
'        strSQL = ""
'        strSQL = "delete from FKS480_W01"       ' 2006/01/25 ADD T.SUZUKI
''        strSQL = "delete from 顧客マスタＷＫ"  ' 2006/01/25 DEL T.SUZUKI
'
'        dbMydb.Execute strSQL, dbFailOnError
'        dbMydb.Close
'    End If
    '' 2013/06/03 M.HONDA END
    '*** SQLSERVERのテーブルよりJETのワークテーブルに取り込み ***
    Set dbMydb = CurrentDb()

    '*** ワーク削除 ***
    strSQL = ""
    strSQL = "delete from FKS480_W01"       ' 2006/01/25 ADD T.SUZUKI
'        strSQL = "delete from 顧客マスタＷＫ"  ' 2006/01/25 DEL T.SUZUKI

    dbMydb.Execute strSQL, dbFailOnError
    dbMydb.Close

Exit_Form_Unload:
    Exit Sub

Err_Form_Unload:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_Unload
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CODEFROM_AfterUpdate イベント
'        MODULE_ID        :cmb_CODEFROM_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODEFROM_AfterUpdate()

    Dim varCode     As Variant
    Dim strUser     As String
    Dim strWhere    As String

    On Error GoTo Err_cmb_CODEFROM_AfterUpdate

    varCode = Nz(Me.cmb_CODEFROM, "")
    strUser = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strUser = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If

    Me.cmb_CODEFROM = Format(Me.cmb_CODEFROM, "000000")
    Me.txt_CODEFROM = strUser
    Me.cmb_CODETO.SetFocus

Exit_cmb_CODEFROM_AfterUpdate:
    Exit Sub

Err_cmb_CODEFROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODEFROM_AfterUpdate
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CODEFROM_BeforeUpdate イベント
'        MODULE_ID        :cmb_CODEFROM_BeforeUpdate
'        CREATE_DATE      :2006/01/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODEFROM_BeforeUpdate(Cancel As Integer)

    Dim varCode     As Variant
    Dim strUser     As String
    Dim strWhere    As String

    On Error GoTo cmb_CODEFROM_BeforeUpdate_Err

    varCode = Nz(Me.cmb_CODEFROM, "")
    strUser = ""

    If varCode = "" Then
        Me.txt_CODEFROM = "": Exit Sub
    End If

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strUser = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If

    If Nz(strUser, "") = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_CODEFROM_BeforeUpdate_Exit
    End If

cmb_CODEFROM_BeforeUpdate_Exit:
    Exit Sub

cmb_CODEFROM_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume cmb_CODEFROM_BeforeUpdate_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CODEFROM_KeyPress イベント
'        MODULE_ID        :cmb_CODEFROM_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODEFROM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODEFROM, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CODETO_AfterUpdate イベント
'        MODULE_ID        :cmb_CODETO_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODETO_AfterUpdate()

    Dim varCode As Variant
    Dim varChCode As Variant
    Dim strUser As String
    Dim strWhere As String

    On Error GoTo Err_cmb_CODETO_AfterUpdate

    varCode = Nz(Me.cmb_CODETO, "")
    strUser = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strUser = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If
    Me.cmb_CODETO = Format(Me.cmb_CODETO, "000000")
    Me.txt_CODETO = strUser
'    Me.txt_更新日ＦＲＯＭ.SetFocus

Exit_cmb_CODETO_AfterUpdate:
    Exit Sub

Err_cmb_CODETO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODETO_AfterUpdate
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CODEto_BeforeUpdate イベント
'        MODULE_ID        :cmb_CODEto_BeforeUpdate
'        CREATE_DATE      :2006/01/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODEto_BeforeUpdate(Cancel As Integer)
            
    Dim varCode     As Variant
    Dim strUser     As String
    Dim strWhere    As String

    On Error GoTo cmb_CODEto_BeforeUpdate_Err

    varCode = Nz(Me.cmb_CODETO, "")
    strUser = ""

    If varCode = "" Then
        Me.txt_CODETO = "": Exit Sub
    End If

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strUser = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If

    If Nz(strUser, "") = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_CODEto_BeforeUpdate_Exit
    End If

cmb_CODEto_BeforeUpdate_Exit:
    Exit Sub

cmb_CODEto_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume cmb_CODEto_BeforeUpdate_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_CODETO_KeyPress イベント
'        MODULE_ID        :cmb_CODETO_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODETO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODETO, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日FROM_AfterUpdate イベント
'        MODULE_ID        :txt_更新日FROM_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_AfterUpdate()

    On Error GoTo Err_txt_更新日FROM_AfterUpdate

    '*-Nullチェック-*
    If IsNull(Me.txt_更新日FROM) Or Me.txt_更新日FROM = "" Then
        Exit Sub
    End If

    '*-Check_Date関数を使用-*
    Me.txt_更新日FROM = Check_Date(Me.txt_更新日FROM)

    Exit Sub

Err_txt_更新日FROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日FROM_BeforeUpdate イベント
'        MODULE_ID        :txt_更新日FROM_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_BeforeUpdate(Cancel As Integer)

    Dim vDATE As Variant

    On Error GoTo Err_txt_更新日FROM_BeforeUpdate

    vDATE = Nz(Me.txt_更新日FROM, "")

    '*-Nullチェック-*
    If vDATE = "" Then Exit Sub

    '*-Check_Date関数を使用-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)

    '入力チェック
    Select Case CmBCheck_Date(vDATE)
        Case 0

        Case 2
        '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_更新日FROM_BeforeUpdate

        Case 3, 9
            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_更新日FROM_BeforeUpdate
    End Select

Exit_txt_更新日FROM_BeforeUpdate:
    Exit Sub

Err_txt_更新日FROM_BeforeUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日FROM_GotFocus イベント
'        MODULE_ID        :txt_更新日FROM_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_GotFocus()
    '*-YY/MM/DDの形に変換-*
    Me.txt_更新日FROM = Format(Me.txt_更新日FROM, "YYYY/MM/DD")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日FROM_KeyPress イベント
'        MODULE_ID        :txt_更新日FROM_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_KeyPress(KeyAscii As Integer)
'    Call KeyPress_For_Date(Me.txt_更新日ＦＲＯＭ, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日FROM_LostFocus イベント
'        MODULE_ID        :txt_更新日FROM_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_LostFocus()
    '*-YY/MM/DDの形に変換-*
    Me.txt_更新日FROM = Format(Me.txt_更新日FROM, "YYYY/MM/DD")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日TO_AfterUpdate イベント
'        MODULE_ID        :txt_更新日TO_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_AfterUpdate()

    On Error GoTo Err_txt_更新日TO_AfterUpdate

    '*-Nullチェック-*
    If IsNull(Me.txt_更新日TO) Or Me.txt_更新日TO = "" Then
        Exit Sub
    End If

    '*-Check_Date関数-*
    Me.txt_更新日TO = Check_Date(Me.txt_更新日TO)

    Exit Sub

Err_txt_更新日TO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日TO_BeforeUpdate イベント
'        MODULE_ID        :txt_更新日TO_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_BeforeUpdate(Cancel As Integer)

    Dim vDATE As Variant

    On Error GoTo Err_txt_更新日TO_BeforeUpdate

    vDATE = Nz(Me.txt_更新日TO, "")

    '*-Nullチェック-*
    If vDATE = "" Then Exit Sub

    '*-Check_Date関数をしよう-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)

    '*-入力チェック-*
    Select Case CmBCheck_Date(vDATE)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_更新日TO_BeforeUpdate

        Case 3, 9
            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_更新日TO_BeforeUpdate
    End Select

Exit_txt_更新日TO_BeforeUpdate:
    Exit Sub

Err_txt_更新日TO_BeforeUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日TO_GotFocus イベント
'        MODULE_ID        :txt_更新日TO_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_GotFocus()
    '**--YY/MM/DDの形に変換--**
    Me.txt_更新日TO = Format(Me.txt_更新日TO, "YYYY/MM/DD")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日TO_LostFocus イベント
'        MODULE_ID        :txt_更新日TO_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_LostFocus()
    '**--YYYY/MM/DDの形に変換--**
    Me.txt_更新日TO = Format(Me.txt_更新日TO, "YYYY/MM/DD")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日TO_KeyDown イベント
'        MODULE_ID        :txt_更新日TO_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim strChdate As String

    On Error GoTo Err_txt_更新日TO_KeyDown

    strChdate = Nz(Me.txt_更新日TO, "")

    If Shift <> 0 And KeyCode = vbKeyTab Then
        Me.txt_更新日FROM.SetFocus
        KeyCode = 0
        Exit Sub
    End If

    '$$$ start $$$  10Keyが使用できなくなるので削除 01-07-2000 Araki
    'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    '    Me.cmd_印刷.SetFocus
    '    KeyCode = 0
    '    Exit Sub
    'End If
    '$$$ end $$$

Exit_txt_更新日TO_KeyDown:
   Exit Sub

Err_txt_更新日TO_KeyDown:
    If Err.Number <> 2110 Then
        MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    End If
    Resume Exit_txt_更新日TO_KeyDown
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_更新日TO_KeyPress イベント
'        MODULE_ID        :txt_更新日TO_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_KeyPress(KeyAscii As Integer)
'    Call KeyPress_For_Date(Me.txt_更新日ＴＯ, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_プレビュー_Click イベント
'        MODULE_ID        :cmd_プレビュー_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_プレビュー_Click()

    Dim strMsg As String
    Dim intAns As Integer

    On Error GoTo Err_cmd_プレビュー_Click

    strMsg = "印刷内容を画面に出力します。" '& vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_プレビュー_Click
    End If

    '*** ワーク作成 ***
    If Make_Work = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_プレビュー_Click
    End If

    doCmd.OpenReport "RKS480_顧客マスタ一覧表", acPreview

Exit_cmd_プレビュー_Click:
    Me.cmb_CODEFROM.SetFocus
    Exit Sub

Err_cmd_プレビュー_Click:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmd_プレビュー_Click
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_印刷_Click イベント
'        MODULE_ID        :cmd_印刷_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_印刷_Click()

    Dim strMsg As String
    Dim intAns As Integer
       
    On Error GoTo Err_cmd_印刷_Click

    strMsg = "印刷内容をプリンタに印字します。" & vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_印刷_Click
    End If

    '*** ワーク作成 ***
    If Make_Work = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_印刷_Click
    End If

    '印刷実行
    doCmd.OpenReport "RKS480_顧客マスタ一覧表", acNormal

Exit_cmd_印刷_Click:
    Me.cmb_CODEFROM.SetFocus
    Exit Sub

Err_cmd_印刷_Click:
    Select Case Err
        Case 2501:
            MsgBox "印刷を中止しました!!", vbInformation + vbOKOnly, strMsg_kaku
        Case 2201:
            MsgBox "プリンタドライバーの設定に誤りがあります。", vbInformation + vbOKOnly, strMsg_kaku
        Case Else:
            MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    End Select
  
    Resume Exit_cmd_印刷_Click
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click イベント
'        MODULE_ID        :cmd_終了_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next
    'doCmd.Close acForm, "FKS480_顧客マスタ一覧表" 'DELETE 20041124 N.MIURA
    doCmd.Close acForm, "FKS480"                   'INSERT 20041124 N.MIURA
End Sub

'==============================================================================*
'
'        MODULE_NAME      :帳票出力用ワークテーブル作成
'        MODULE_ID        :Make_Work
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_Work() As Boolean

    '*** 顧客マスタ一覧表作成 ***
    Dim strWhere    As String
    Dim lngKensu    As Long
    Dim strSQL      As String
    Dim dbMydb      As Database
    Dim strMsg      As String

    On Error GoTo Err_Make_Work

    Make_Work = False
    doCmd.Hourglass True

    '*** 抽出条件作成 ***
    strWhere = Make_Where

    '*** SQLSERVERのテーブルよりJETのワークテーブルに取り込み ***
    Set dbMydb = CurrentDb()

    '*** ワーク削除 ***
    strSQL = ""
    strSQL = "delete from FKS480_W01"                  ' 2006/01/25 ADD T.SUZUKI
'    strSQL = "delete from 顧客マスタＷＫ"             ' 2006/01/25 DEL T.SUZUKI
    dbMydb.Execute strSQL, dbFailOnError

    '*** ワーク追加 ***
    strSQL = ""
    strSQL = strSQL & " INSERT INTO FKS480_W01("       ' 2006/01/25 ADD T.SUZUKI
'    strSQL = strSQL & " INSERT INTO 顧客マスタＷＫ("  ' 2006/01/25 DEL T.SUZUKI
    strSQL = strSQL & " USER_CODE,"       '0
' 2006/01/25 ADD T.SUZUKI Start
    strSQL = strSQL & " USER_NEW_CODE,"   '1 新顧客コード
' 2006/01/25 ADD T.SUZUKI End
    strSQL = strSQL & " USER_NAME,"       '2
    strSQL = strSQL & " USER_KANA,"       '3
    strSQL = strSQL & " USER_TANM,"       '4
    strSQL = strSQL & " USER_TAKA,"       '5
    strSQL = strSQL & " USER_YUBINO,"     '6
    strSQL = strSQL & " USER_ADR_1,"      '7
    strSQL = strSQL & " USER_ADR_2,"      '8
    strSQL = strSQL & " USER_TEL,"        '9
    strSQL = strSQL & " USER_KEITAI,"     '10
    strSQL = strSQL & " USER_FAX,"        '11
    strSQL = strSQL & " USER_SDAY,"       '12
    strSQL = strSQL & " USER_KKBN,"       '13
    strSQL = strSQL & " USER_SKBN,"       '14
    strSQL = strSQL & " USER_BIKO,"       '15
    strSQL = strSQL & " USER_UPDATE)"     '16
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " dbo_USER_MAST.USER_CODE,"       '0
    strSQL = strSQL & " IIF(nz(dbo_USER_MAST.USER_NEW_CODE,'') = '', NULL, "
    strSQL = strSQL & " '新：' & format(dbo_USER_MAST.USER_NEW_CODE,'000000')), "   '1 新顧客コード
    strSQL = strSQL & " dbo_USER_MAST.USER_NAME,"       '2
    strSQL = strSQL & " dbo_USER_MAST.USER_KANA,"       '3
    strSQL = strSQL & " dbo_USER_MAST.USER_TANM,"       '4
    strSQL = strSQL & " dbo_USER_MAST.USER_TAKA,"       '5
    strSQL = strSQL & " dbo_USER_MAST.USER_YUBINO,"     '6
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_1,"      '7
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_2,"      '8
    strSQL = strSQL & " dbo_USER_MAST.USER_TEL,"        '9
    strSQL = strSQL & " dbo_USER_MAST.USER_KEITAI,"     '10
    strSQL = strSQL & " dbo_USER_MAST.USER_FAX,"        '11
    strSQL = strSQL & " dbo_USER_MAST.USER_SDAY,"       '12
    strSQL = strSQL & " dbo_USER_MAST.USER_KKBN,"       '13
    strSQL = strSQL & " dbo_USER_MAST.USER_SKBN,"       '14
    strSQL = strSQL & " dbo_USER_MAST.USER_BIKO,"       '15
    strSQL = strSQL & " dbo_USER_MAST.USER_UPDATE"      '16
    strSQL = strSQL & " FROM dbo_USER_MAST"

    If strWhere <> "" Then
        strSQL = strSQL & " WHERE " & strWhere
    End If

    dbMydb.Execute strSQL, dbFailOnError
    
    '*** データチェック ***
    'アクションクエリーの結果で判断
    lngKensu = dbMydb.RecordsAffected
    If lngKensu < 1 Then
        strMsg = "対象となるデータが存在しませんでした。"
        MsgBox strMsg, vbInformation + vbOKOnly, strMsg_kaku
        GoTo Exit_Make_Work
    End If

''''' ADD 99/12/20  名称マスタより略称をセット
    strSQL = ""
    strSQL = " UPDATE FKS480_W01 INNER JOIN dbo_NAME_MAST ON "                 ' 2006/01/25 ADD T.SUZUKI
    strSQL = strSQL & " FKS480_W01.USER_KKBN = dbo_NAME_MAST.NAME_CODE "       ' 2006/01/25 ADD T.SUZUKI
    strSQL = strSQL & " SET FKS480_W01.USER_RYAK1 = [NAME_RYAK] "              ' 2006/01/25 ADD T.SUZUKI
'    strSQL = " UPDATE 顧客マスタＷＫ INNER JOIN dbo_NAME_MAST ON "            ' 2006/01/25 DEL T.SUZUKI
'    strSQL = strSQL & " 顧客マスタＷＫ.USER_KKBN = dbo_NAME_MAST.NAME_CODE    " 2006/01/25 DEL T.SUZUKI
'    strSQL = strSQL & " SET 顧客マスタＷＫ.USER_RYAK1 = [NAME_RYAK] "         ' 2006/01/25 DEL T.SUZUKI
    strSQL = strSQL & " WHERE (((dbo_NAME_MAST.NAME_ID) = '020' )) "
    
    dbMydb.Execute strSQL, dbFailOnError

    strSQL = ""
    strSQL = " UPDATE FKS480_W01 INNER JOIN dbo_NAME_MAST ON "                 ' 2006/01/25 ADD T.SUZUKI
    strSQL = strSQL & " FKS480_W01.USER_SKBN = dbo_NAME_MAST.NAME_CODE "       ' 2006/01/25 ADD T.SUZUKI
    strSQL = strSQL & " SET FKS480_W01.USER_RYAK2 = [NAME_RYAK] "              ' 2006/01/25 ADD T.SUZUKI
'    strSQL = " UPDATE 顧客マスタＷＫ INNER JOIN dbo_NAME_MAST ON "            ' 2006/01/25 DEL T.SUZUKI
'    strSQL = strSQL & " 顧客マスタＷＫ.USER_SKBN = dbo_NAME_MAST.NAME_CODE "  ' 2006/01/25 DEL T.SUZUKI
'    strSQL = strSQL & " SET 顧客マスタＷＫ.USER_RYAK2 = [NAME_RYAK] "         ' 2006/01/25 DEL T.SUZUKI
    strSQL = strSQL & " WHERE (((dbo_NAME_MAST.NAME_ID) = '030' )) "
    
    dbMydb.Execute strSQL, dbFailOnError

''''' ADD 99/12/20  名称マスタより略称をセット

    Make_Work = True

Exit_Make_Work:
    doCmd.Hourglass False
    Exit Function

Err_Make_Work:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_Work
End Function

'==============================================================================*
'
'        MODULE_NAME      :条件式作成
'        MODULE_ID        :Make_Where
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_Where() As String

    Dim strCodefr  As String
    Dim strCODETO  As String
    Dim strDayfr   As String
    Dim strDayto   As String
    Dim strSwap    As String
    Dim strWhere   As String

    On Error GoTo Err_Make_Where

    '*** 画面指定値取得 ***'
    strCodefr = Nz(Me.cmb_CODEFROM, "")
    strCODETO = Nz(Me.cmb_CODETO, "")

    '*** ＩＤの大小チェック ***
    If Trim(strCodefr) <> "" And Trim(strCODETO) <> "" Then
        '*** ＣＯＤＥ入替 ***
        If strCodefr > strCODETO Then
            strSwap = strCodefr
            strCodefr = strCODETO
            strCODETO = strSwap
        End If
    End If

    '更新日のチェック
    If Me.txt_更新日FROM <> "" And Me.txt_更新日TO <> "" Then
        strDayfr = Format(Nz(Me.txt_更新日FROM, ""), "YYYY/MM/DD")
        strDayto = Format(Nz(Me.txt_更新日TO, ""), "YYYY/MM/DD")
        If strDayto <> "" And strDayfr > strDayto Then
            'strSwap = strDayto                        'Del 01-11-2000 Araki
            'Me.txt_更新日ＦＲＯＭ = strSwap            'Del 01-11-2000 Araki
            'txt_更新日ＴＯ = strDayfr                  'Del 01-11-2000 Araki
            
            '$$$ start $$$ 更新日選択範囲の大小逆転時のデータ抽出ミス修正のため追加 01-11-2000 Araki
            strSwap = strDayfr
            strDayfr = strDayto
            strDayto = strSwap
            Me.txt_更新日FROM = strDayfr
            Me.txt_更新日TO = strDayto
            '$$$ end $$$
        End If
    End If

    '*** ＣＯＤＥ条件指定 ***
    strWhere = ""
    If Trim(strCodefr) <> "" And Trim(strCODETO) = "" Then
        strWhere = strWhere & " USER_CODE >= " & Trim(strCodefr) & ""
    ElseIf Trim(strCodefr) = "" And Trim(strCODETO) <> "" Then
        strWhere = strWhere & "USER_CODE <= " & strCODETO & ""
    ElseIf Trim(strCodefr) <> "" And Trim(strCODETO) <> "" Then
        strWhere = strWhere & " USER_CODE BETWEEN " & strCodefr & " "
        strWhere = strWhere & " AND " & strCODETO & " "
    ElseIf Trim(strCodefr) = "" And Trim(strCODETO) = "" Then
    End If

    If strWhere <> "" And strDayfr <> "" Then
        strWhere = strWhere & " AND "
    End If

    If strDayfr <> "" Then
        strWhere = strWhere & "USER_UPDATE >=cvdate(#" & strDayfr & "#)"
    End If

    If strWhere <> "" And strDayto <> "" Then
        strWhere = strWhere & " AND "
    End If

    If strDayto <> "" Then
        strWhere = strWhere & "USER_UPDATE <=cvdate(#" & strDayto & "#)"
    End If
    
    Make_Where = strWhere

Exit_Make_Where:
    Exit Function

Err_Make_Where:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_Where
End Function

'==============================================================================*
'
'        MODULE_NAME      :txt_END_Enter イベント
'        MODULE_ID        :txt_END_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_END_Enter()
    Me.cmd_印刷.SetFocus
End Sub

' @f
'
' 機　能　 :キーの取消 日付のみ対応
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
Private Sub KeyPress_For_Date(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                     ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc("/")   ' 0～9迄、/ を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                If Len(.Text) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

Private Function CMCheck_Date(ByVal vDATE As Variant)
'-----------------------------------------------------------'
'  Check_Dateでは'198/03といった形が変換されるための補助検査
'-----------------------------------------------------------'

    Dim sYear As String

    CMCheck_Date = Null

    If Nz(vDATE, "") = "" Then Exit Function

    If IsDate(vDATE) = False Then
        Exit Function
    End If

    sYear = Format(vDATE, "yyyy")

    If CmLenb(sYear) <> 4 Then
        Exit Function
    End If

    CMCheck_Date = vDATE
End Function

Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer
'------------------------'
'  入金年月日のチェック
'------------------------'
  '引数
  'oText 対象項目
  'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
    
  '返り値は、
  '２：桁数が溢れている時
  '３：日付と判断できない時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時

    Dim sYear  As String
    Dim intLen As Integer

    On Error GoTo Err_CmBCheck_Date

    CmBCheck_Date = 2

    If IsMissing(ilen) = False Then
        intLen = Nz(CmNumeric(ilen), 0)
    Else
        intLen = 10
    End If

    oText = Nz(oText, "")

    '桁数チェック
    If CmLenb(oText) > intLen Then
        GoTo Exit_CmBCheck_Date
    End If

    CmBCheck_Date = 3

    If Nz(oText, "") = "" Then
        GoTo Exit_CmBCheck_Date
    End If

    '論理チェック
    If IsDate(oText) = False Then
        GoTo Exit_CmBCheck_Date
    End If

    sYear = Format(oText, "yyyy")

    If CmLenb(sYear) <> 4 Then
        GoTo Exit_CmBCheck_Date
    End If

    CmBCheck_Date = 0

Exit_CmBCheck_Date:
    Exit Function

Err_CmBCheck_Date:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    CmBCheck_Date = 9
    GoTo Exit_CmBCheck_Date
End Function
'****************************  ended or program ********************************
