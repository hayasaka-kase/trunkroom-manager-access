VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS490"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　倉庫管理システム
'**********************************************
'   プログラム名　　：　ヤードマスタ一覧表
'   プログラムＩＤ　：
'   作　成　日　　　：
'   作　成　者　　　：　αシステム
'
'   修　正　日　　　： 2004/11/06
'   修　正　者　　　： N.MIURA
'   Ver             ： 0.2
'
'   修　正　日　　　： 2009/12/08
'   修　正　者　　　： M.RYU
'   Ver             ： 0.3
'   修　正　内　容　：①ワークTBL「ヤードマスタＷＫ」⇒作成するとき、借上№を追加
'                     ②レポート（RKS490_ヤードマスタ一覧表）⇒借上№列を追加
'                     ③クエリ「QKS490_M00」⇒借上№を追加
'                     ④印刷ボタン無効になってるから、ﾌﾟﾚﾋﾞｭｰボタンにフォーカスを設定
'
'   修　正　日　　　： 2009/12/08
'   修　正　者　　　： M.RYU
'   Ver             ： 0.3
'   修　正　内　容　：①ワークTBL「ヤードマスタＷＫ」⇒作成するとき、借上№を追加
'                     ②レポート（RKS490_ヤードマスタ一覧表）⇒借上№列を追加
'                     ③クエリ「QKS490_M00」⇒借上№を追加
'                     ④印刷ボタン無効になってるから、ﾌﾟﾚﾋﾞｭｰボタンにフォーカスを設定
'
'   修　正　日　　　： 2011/04/11
'   修　正　者　　　： M.RYU
'   Ver             ： 0.4
'   修　正　内　容　：①印刷ボタンを削除
'                     ②検索条件を追加：HP更新日、ネット予約対象、自動ページup対象、
'                                       営業終了、ヤード解約予定、ヤード非表示
'                     ③画面初期表示：ヤードコードを最小～最大、更新日は当日
'                     ④QKS490_M00      V0.2 2011/04/11 ヤードマスタ一覧
'
'
'   修　正　日　　　： 2013/06/03
'   修　正　者　　　： M.HONDA
'   Ver             ： 0.5
'   修　正　内　容　：終了メッセージを表示させないように修正
'
'
'   修　正　日　　　： 2013/09/07
'   修　正　者　　　： M.HONDA
'   Ver             ： 0.6
'   修　正　内　容　： エクセル出力に巡回担当、メンテ担当、賃貸種別を追加
'
'   修　正　日　　　： 2015/10/26
'   修　正　者　　　： M.HONDA
'   Ver             ： 0.7
'   修　正　内　容　： エクセル出力に最寄駅、営業時間、物件詳細ページを追加
'
'   修　正　日　　　： 2017/07/29
'   修　正　者　　　： M.HONDA
'   Ver             ： 0.8
'   修　正　内　容　： 項目の追加
'
'   修　正　日　　　： 2017/10/12
'   修　正　者　　　： S.WATANABE
'   Ver             ： 0.9
'   修　正　内　容　： 価格テンプレートを追加
'
'   修　正　日　　　： 2018/01/20
'   修　正　者　　　： N.IMAI
'   Ver             ： 1.0
'   修　正　内　容　： ヤード担当を追加
'
'   修　正　日　　　： 2019/07/09
'   修　正　者　　　： M.HONDA
'   Ver             ： 1.1
'   修　正　内　容　： データ取得部分をごっそりと書き直し
'
'   修　正　日　　　： 2022/09/27
'   修　正　者　　　： N.IMAI
'   Ver             ： 1.2
'   修　正　内　容　： ヤード解約予定日を追加
'
'   修　正　日　　　： 2024/10/23
'   修　正　者　　　： N.IMAI
'   Ver             ： 1.3
'   修　正　内　容　： セキュリティ会社を追加
'
'   修　正　日　　　： 2025/11/21
'   修　正　者　　　： K.KINEBUCHI
'   Ver             ： 1.4
'   修　正　内　容　： Excel出力に、借上物件の「所有区分」を追加する
'
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS490"

Const strMsg_kaku = "確認"
Const strMsg_err = "エラー"
Private Const WK_YARD_MAST = "ヤードマスタＷＫ"           'INSERT 2011/04/11 M.RYU
Private pstrBumoCD As String                              '2013/09/07 M.HONDA INS
Private Const C_TEMP_TABLE  As String = "FKS490W"
'
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODEFROM_AfterUpdate()
On Error GoTo Err_cmb_CODEFROM_AfterUpdate
'    Dim varId As Variant                           '---UPDATE 2011/04/11 M.RYU------<S>
'    Dim varCode As Variant
'    Dim strYard As String
'    Dim varChCode As Variant
'    Dim strWhere As String
'
'    varCode = Me.cmb_CODEFROM
'    strYard = ""
'
'    If Trim(varCode) <> "" Then
'        strWhere = ""
'        strWhere = strWhere & " YARD_CODE=" & varCode & ""
'        strYard = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
'    End If
'    Me.cmb_CODEFROM = Format(Me.cmb_CODEFROM, "000000")
'    Me.txt_CODEFROM = strYard
    
    Me.txt_CODEFROM = Me.cmb_CODEFROM.Column(1)     '---UPDATE 2011/04/11 M.RYU------<E>
    Me.cmb_CODETO.SetFocus
    
Exit_cmb_CODEFROM_AfterUpdate:
    Exit Sub

Err_cmb_CODEFROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODEFROM_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODEFROM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODEFROM, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODETO_AfterUpdate()
On Error GoTo Err_cmb_CODETO_AfterUpdate
'    Dim varCode As Variant                           '---UPDATE 2011/04/11 M.RYU------<S>
'    Dim varChCode As Variant
'    Dim strYard As String
'    Dim strWhere As String
'
'    varCode = Me.cmb_CODETO
'    strYard = ""
'
'    If Trim(varCode) <> "" Then
'        strWhere = ""
'        strWhere = strWhere & " YARD_CODE=" & varCode & ""
'        strYard = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
'    End If
'    Me.cmb_CODETO = Format(Me.cmb_CODETO, "000000")
'    Me.txt_CODETO = strYard

    Me.txt_CODETO = Me.cmb_CODETO.Column(1)           '---UPDATE 2011/04/11 M.RYU------<E>
'    Me.txt_更新日FROM.SetFocus

Exit_cmb_CODETO_AfterUpdate:
    Exit Sub

Err_cmb_CODETO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODETO_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODETO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODETO, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_プレビュー_Click()
On Error GoTo Err_cmd_プレビュー_Click
    Dim strMsg As String
    Dim intAns As Integer

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

    doCmd.OpenReport "RKS490_ヤードマスタ一覧表", acPreview

Exit_cmd_プレビュー_Click:
    Me.cmb_CODEFROM.SetFocus
    Exit Sub

Err_cmd_プレビュー_Click:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmd_プレビュー_Click

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_EXCEL_ClicK
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()
    On Error GoTo Err_cmd_excel_Click
    
    Dim strMsg     As String
    Dim intAns     As Integer

    strMsg = "印刷内容をEXCELに出力します。" '& vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_excel_Click
    End If

    '*** ワーク作成 ***
    If Make_Work = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_excel_Click
    End If

    doCmd.OutputTo acOutputTable, C_TEMP_TABLE, "MicrosoftExcel(*.xls)", "", True, ""
            
Exit_cmd_excel_Click:
    doCmd.Hourglass False
    Me.cmb_CODEFROM.SetFocus
    Exit Sub

Err_cmd_excel_Click:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmd_excel_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
On Error Resume Next
    'doCmd.Close acForm, "FKS490_ヤードマスタ一覧表" 'DELETE 20041106 N.MIURA
    doCmd.Close acForm, "FKS490"                       'INSERT 20041106 N.MIURA

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_EXCEL_ClicK
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_EXCEL_ClicK
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
'        MODULE_NAME      :
'        MODULE_ID        :
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
'            Me.cmd_印刷.SetFocus           'DELETE 2011/04/11 M.RYU
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
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    '＊＊＊画面の最大化＊＊＊
    doCmd.Maximize
        
    '＊＊＊画面表示クリア＊＊＊
'    Me.cmb_CODEFROM = ""                       '---UPDATE 2011/04/11 M.RYU------<S>
'    Me.cmb_CODETO = ""
'    Me.txt_CODEFROM = ""
'    Me.txt_CODETO = ""
'    Me.txt_更新日FROM = ""
'    Me.txt_更新日TO = ""
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Title.BackColor = RGB(255, 0, 204)
    End Select
    
    ' 部門コード取得
    pstrBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")    '' 2013/09/07 M.HONDA INS
    
    Me.cmb_CODEFROM = Me.cmb_CODEFROM.ItemData(0)
    Me.txt_CODEFROM = Me.cmb_CODEFROM.Column(1)
    Me.cmb_CODETO = Me.cmb_CODETO.ItemData(Me.cmb_CODETO.ListCount - 1)
    Me.txt_CODETO = Me.cmb_CODETO.Column(1)
    Me.txt_更新日FROM = Format(DATE, "yyyy/mm/dd")
    Me.txt_更新日TO = Format(DATE, "yyyy/mm/dd")
    Me.txt_WebPageFROM = ""
    Me.txt_WebPageTO = ""                       '---UPDATE 2011/04/11 M.RYU------<E>
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload
    Dim strMsg As String
    Dim intAns As Integer
    Dim dbMydb As Database
    Dim strSQL As String
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
''        strSQL = ""                                        'UPDATE 2011/04/11 M.RYU------<S>
''        strSQL = "delete from ヤードマスタＷＫ"
''        dbMydb.Execute strSQL, dbFailOnError
'        Call subDropWorkTable(dbMydb)                       'UPDATE 2011/04/11 M.RYU------<E>
'
'        dbMydb.Close
'    End If
    '' 2013/06/03 M.HONDA END
    '*** SQLSERVERのテーブルよりJETのワークテーブルに取り込み ***
    Set dbMydb = CurrentDb()
    '*** ワーク削除 ***
    Call subDropWorkTable(dbMydb)                       'UPDATE 2011/04/11 M.RYU------<E>
    dbMydb.Close

Exit_Form_Unload:
    Exit Sub
    
Err_Form_Unload:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_Unload

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
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
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_txt_更新日FROM_BeforeUpdate

    Dim vDATE As Variant
    vDATE = Nz(Me.txt_更新日FROM, "")
    
    '*-Nullチェック-*
    If vDATE = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数を使用-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)

    '*-入力チェック-*
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
'        MODULE_NAME      :
'        MODULE_ID        :
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
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_KeyPress(KeyAscii As Integer)

    Call KeyPress_For_Date(Me.txt_更新日FROM, KeyAscii, 8)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日FROM_LostFocus()

    '*-YY/MM/DDの形に変換-*
    Me.txt_更新日FROM = Format(Me.txt_更新日FROM, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
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
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_txt_更新日TO_BeforeUpdate

    Dim vDATE As Variant
    
    vDATE = Nz(Me.txt_更新日TO, "")
    
    '*-Nullチェック-*
    If vDATE = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数をしよう-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)

    '*-入力チェック-*
    '入力チェック
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
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_GotFocus()

    '**--YY/MM/DDの形に変換--**
    Me.txt_更新日TO = Format(Me.txt_更新日TO, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_LostFocus()

    '**--YYYY/MM/DDの形に変換--**
    Me.txt_更新日TO = Format(Me.txt_更新日TO, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_txt_更新日TO_KeyDown
    
    Dim strChdate As String

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
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_更新日TO_KeyPress(KeyAscii As Integer)
    
    Call KeyPress_For_Date(Me.txt_更新日TO, KeyAscii, 8)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_END_Enter()
    '---20091205---M.RYU----update--------<s>
    '印刷ボタン無効になってるから、ﾌﾟﾚﾋﾞｭｰボタンにフォーカスを設定
'    Me.cmd_印刷.SetFocus
    Me.cmd_プレビュー.SetFocus
    '---20091205---M.RYU----update--------<s>
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_Where() As String
    On Error GoTo Err_Make_Where

    Dim lngCodefr As String
    Dim lngCodeto As String
    Dim strDayfr As String
    Dim strDayto As String
    Dim strSwap As String
    Dim strWhere As String
    Dim strTcode As String   'INS 2017/10/12 S.WATANABE

    '*** 画面指定値取得 ***'
    lngCodefr = Nz(Me.cmb_CODEFROM, 0)
    lngCodeto = Nz(Me.cmb_CODETO, 0)
    strTcode = Nz(Me.cmb_価格テンプコード, "")   'INS 2017/10/12 S.WATANABE

    '*** ＩＤの大小チェック ***
    If Trim(lngCodefr) <> "" And Trim(lngCodeto) <> "" Then
        '*** ＣＯＤＥ入替 ***
        If lngCodefr > lngCodeto Then
            strSwap = lngCodefr
            lngCodefr = lngCodeto
            lngCodeto = strSwap
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
    If Trim(lngCodefr) <> "" And Trim(lngCodeto) = "" Then
        strWhere = strWhere & " YARD_CODE >= " & Trim(lngCodefr) & ""
    ElseIf Trim(lngCodefr) = "" And Trim(lngCodeto) <> "" Then
        strWhere = strWhere & "YARD_CODE <= " & lngCodeto & ""
    ElseIf Trim(lngCodefr) <> "" And Trim(lngCodeto) <> "" Then
        strWhere = strWhere & " YARD_CODE BETWEEN " & lngCodefr & " "
        strWhere = strWhere & " AND " & lngCodeto & " "

    ElseIf Trim(lngCodefr) = "" And Trim(lngCodeto) = "" Then

    End If

    If strWhere <> "" And strDayfr <> "" Then
        strWhere = strWhere & " AND "
    End If

    If strDayfr <> "" Then
        strWhere = strWhere & "YARD_UPDATE >=cvdate(#" & strDayfr & "#)"
    End If

    If strWhere <> "" And strDayto <> "" Then
        strWhere = strWhere & " AND "
    End If

    If strDayto <> "" Then
        strWhere = strWhere & "YARD_UPDATE <=cvdate(#" & strDayto & "#)"
    End If
    
    
    If Me.txt_WebPageFROM <> "" And Me.txt_WebPageTO <> "" Then     '---INSERT 2011/04/11 M.RYU--------<S>--'検索条件を追加
        strDayfr = Format(Nz(Me.txt_WebPageFROM, ""), "YYYY/MM/DD")
        strDayto = Format(Nz(Me.txt_WebPageTO, ""), "YYYY/MM/DD")
        If strDayto <> "" And strDayfr > strDayto Then
            '更新日選択範囲の大小逆転時
            strSwap = strDayfr
            strDayfr = strDayto
            strDayto = strSwap
            Me.txt_WebPageFROM = strDayfr
            Me.txt_WebPageTO = strDayto
        End If
    End If
    
    If Nz(Me.txt_WebPageFROM) <> "" Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_WP_UPDATE WebPage >= cvdate(#" & Nz(Me.txt_WebPageFROM) & "#)"    'WEB更新日
    End If
    
    If Nz(Me.txt_WebPageTO) <> "" Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_WP_UPDATE WebPage <= cvdate(#" & Nz(Me.txt_WebPageTO) & "#)"      'WEB更新日
    End If
            
    If frm_NetUse = optNetUse_not.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_NETUSE_KBN <> -1     " 'ネット予約対象可否
    ElseIf frm_NetUse = optNetUse_only.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_NETUSE_KBN = -1      " 'ネット予約対象可否
    End If
    
    If frm_WebUpLoad = optWebUpLoad_not.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_WP_UPLOAD <> -1      " 'アップロード対象可否
    ElseIf frm_WebUpLoad = optWebUpLoad_only.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_WP_UPLOAD = -1       " 'アップロード対象可否
    End If
    
    If frm_RentendDay = optRentendDay_not.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_RENTEND_DAY IS NULL                "   '営業終了日
    ElseIf frm_RentendDay = optRentendDay_only.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_RENTEND_DAY IS NOT NULL            "   '営業終了日
    End If
    
    If frm_EndDay = optEndDay_not.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_END_DAY IS NULL                    "   'ヤード解約予定日
    ElseIf frm_EndDay = optEndDay_only.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_END_DAY IS NOT NULL                "   'ヤード解約予定日
    End If
    
    If frm_NondispDay = optNondispDay_not.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_NONDISP_DAY IS NULL                "  'ヤード非表示日
    ElseIf frm_NondispDay = optNondispDay_only.OptionValue Then
        strWhere = IIf(strWhere = "", "", strWhere & "AND") & " YARD_NONDISP_DAY IS NOT NULL            "  'ヤード非表示日
    End If                                                          '---INSERT 2011/04/11 M.RYU--------<E>--'検索条件を追加
    
    'strWhere = strWhere & "AND ((dbo_NAME_MAST.NAME_ID)='207') AND ((dbo_TANT_MAST.TANTM_BUMOC)='" & pstrBumoCD & "') AND ((dbo_TANT_MAST_1.TANTM_BUMOC)='" & pstrBumoCD & "') "    '' 2013/09/07 M.HONDA INS
    'strWhere = strWhere & "AND ((N207.NAME_ID)='207') AND ((N041.NAME_ID)='041') AND ((N279.NAME_ID)='279') AND ((dbo_TANT_MAST.TANTM_BUMOC)='" & pstrBumoCD & "') AND ((dbo_TANT_MAST_1.TANTM_BUMOC)='" & pstrBumoCD & "') "    '' 2013/09/07 M.HONDA INS
    strWhere = strWhere & "AND ((N207.NAME_ID)='207') AND ((N041.NAME_ID)='041') AND ((N279.NAME_ID)='279') AND ((dbo_TANT_MAST.TANTM_BUMOC)='" & pstrBumoCD & "') AND ((dbo_TANT_MAST_1.TANTM_BUMOC)='" & pstrBumoCD & "') AND ((nz(dbo_TANT_MAST_2.TANTM_BUMOC, '" & pstrBumoCD & "')='" & pstrBumoCD & "')) "    '' 2018/01/20 N.IMAI INS
    
    'INS 2017/10/12 S.WATANABE<S>
    '価格テンプレートが選択されている場合
    If strTcode <> "" Then
        strWhere = strWhere & "AND KTPS_KKTH_TBL.KKTHM_TCODE = '" & strTcode & "' "                        '適用価格テンプレート
    End If
    'INS 2017/10/12 S.WATANABE<E>
    
    Make_Where = strWhere

Exit_Make_Where:
    Exit Function

Err_Make_Where:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_Where

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_Work() As Boolean
On Error GoTo Err_Make_Work
'*** ヤードマスタ一覧表作成 ***
    Dim strWhere  As String
    Dim lngKensu As Long
    Dim strSQL As String
    Dim dbMydb As Database
    Dim strMsg As String
    Dim intAns As Integer
    Dim varChData As Variant
    Dim strBumonCode As String
    Dim strDataSource As String
    Dim lngCnt As Integer
    Dim rsSqlServer         As Object       'ADODB.Recordset
    '部門コード取得
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    Make_Work = False
    doCmd.Hourglass True

'    '*** 抽出条件作成 ***
    strWhere = Make_Where
'
    
   'SQL-Server接続
    Dim objCon         As Object       'ADODB.Connection
    Set objCon = ADODB_Connection(strBumonCode)
    
    strSQL = strSQL & " select  "
    strSQL = strSQL & "         YARD_CODE as ヤードコード, "
    strSQL = strSQL & "         YARD_NAME as ヤード名, "
    strSQL = strSQL & "         YARD_YUBINO as 郵便番号, "
    strSQL = strSQL & "         YARD_ADDR_1 as 住所, "
    strSQL = strSQL & "         YARD_ADDR_2 as 住所2, "
    strSQL = strSQL & "         YARD_ADDR_3 as 住所3, "
    strSQL = strSQL & "         YARD_BIKON as 備考, "
    strSQL = strSQL & "         YARD_NOTE  as 補足, "
    strSQL = strSQL & "         HKEIT_BUKHC as 物件No, "
    strSQL = strSQL & "         SHIRM_SHIRN as オーナー名,  "
    strSQL = strSQL & "         HKEIT_KEIYD as 契約日, "
    strSQL = strSQL & "         HKEIT_SIKAD as 賃料発生日, "
    strSQL = strSQL & "         HKEIT_KIKAD as 契約期間, "
    strSQL = strSQL & "         HKEIT_KAIYD as 解約予定日, "
    strSQL = strSQL & "         HKEIT_KEIYQ as 坪数, "
    strSQL = strSQL & "         HKEIT_KEIYU as 単価, "
    strSQL = strSQL & "         HKEIT_KEIYA as 契約金額, "
    strSQL = strSQL & "         C210.NAME_NAME AS 所有区分, "               'INSERT 2025/11/21 K.KINEBUCHI
    strSQL = strSQL & "         CODET_NAMEN as 管理区分, "
    strSQL = strSQL & "         NAME_207.NAME_NAME AS 賃貸種別, "
    strSQL = strSQL & "         NAME_041.NAME_NAME AS 建築確認, "
    strSQL = strSQL & "         YARD_TANTO as リーダーコード, "
    strSQL = strSQL & "         isnull(TANT.TANTM_TANTN,'') as ヤードリーダ, "
    strSQL = strSQL & "         YARD_MNT_TANTO as メンテコード, "
    strSQL = strSQL & "         isnull(MNT_TANT.TANTM_TANTN,'') as メンテ担当, "
    strSQL = strSQL & "         YARD_JUN_TANTO as 巡回コード, "
    strSQL = strSQL & "         isnull(JUN_TANT.TANTM_TANTN,'') as 巡回担当, "
    strSQL = strSQL & "         NAME_062.NAME_NAME AS セキュリティ会社, "    'INSERT 2024/10/23 N.IMAI
    strSQL = strSQL & "         YARD_BEGIN_DAY as 営業開始日, "
    strSQL = strSQL & "         YARD_RENTEND_DAY    as 営業終了日, "
    strSQL = strSQL & "         YARD_END_DAY    as ヤード解約予定日, "  'INSERT 2022/09/27 N.IMAI
    strSQL = strSQL & "         isnull(YARD_IDO,0) as 緯度, "
    strSQL = strSQL & "         isnull(YARD_KEIDO,0) as 経度, "
    strSQL = strSQL & "         isnull(YARD_NETUSE_KBN,0) as ネット対象, "
    strSQL = strSQL & "         isnull(YARD_TIHO_KBN,0) as 地方ヤード, "
    strSQL = strSQL & "         YARD_HOSYO_CD as 保証会社コード, "
    strSQL = strSQL & "         YARD_NET_HOSYO_CD as ネット保証会社コード, "
    strSQL = strSQL & "         YARD_NET_HOSCD_KIN2 as 適用賃料, "
    strSQL = strSQL & "         YARD_NET_HOSCD_SHR2 as 適用支払 , "
    strSQL = strSQL & "         YARD_NET_HOSCD2 as コード, "
    strSQL = strSQL & "         RLDNT_NO as 割引No, "
    strSQL = strSQL & "         YARD_SEV1N as サービス1,  "
    strSQL = strSQL & "         YARD_SEV2N as サービス2, "
    strSQL = strSQL & "         YARD_SEV3N as サービス3, "
    strSQL = strSQL & "         YARD_ENDEN as サービス期間, "
    strSQL = strSQL & "         YARD_SEV_EXMONTH as 満了月数, "
    strSQL = strSQL & "         KTPST_STDTE as テンプレート設定日,"
    strSQL = strSQL & "         KKTHM_TNAME as テンプレート名称 "
    strSQL = strSQL & " from    YARD_MAST "
    strSQL = strSQL & "         LEFT JOIN NAME_MAST NAME_207 ON "
    strSQL = strSQL & "            NAME_207.NAME_ID = '207' AND "
    strSQL = strSQL & "            NAME_207.NAME_CODE = YARD_TINTI "
    
    strSQL = strSQL & "         LEFT JOIN RLDN_TRAN ON "
    strSQL = strSQL & "             RLDNT_ENABLE = 1 AND "
    strSQL = strSQL & "             RLDNT_YCODE = YARD_CODE and"
    strSQL = strSQL & "            RLDNT_TO >= '" & Format(DATE, "YYYYMMDD") & "'"
    
    
    
    
    strSQL = strSQL & "         LEFT JOIN NAME_MAST NAME_041 ON "
    strSQL = strSQL & "            NAME_041.NAME_ID = '041' AND "
    strSQL = strSQL & "            NAME_041.NAME_CODE = YARD_KENTIKU "
    'INSERT 2024/10/23 N.IMAI Start
    strSQL = strSQL & "         LEFT JOIN NAME_MAST NAME_062 ON "
    strSQL = strSQL & "            NAME_062.NAME_ID = '062' AND "
    strSQL = strSQL & "            NAME_062.NAME_CODE = YARD_SECUCOMP "
    'INSERT 2024/10/23 N.IMAI End
    strSQL = strSQL & "         LEFT JOIN " & strDataSource & ".dbo.TANT_MAST TANT ON "
    strSQL = strSQL & "            TANT.TANTM_TANTC = YARD_TANTO AND "
    strSQL = strSQL & "            TANT.TANTM_BUMOC = '" & strBumonCode & "' "
    strSQL = strSQL & "         LEFT JOIN " & strDataSource & ".dbo.TANT_MAST MNT_TANT ON "
    strSQL = strSQL & "            MNT_TANT.TANTM_TANTC = YARD_MNT_TANTO AND "
    strSQL = strSQL & "            MNT_TANT.TANTM_BUMOC = '" & strBumonCode & "' "
    strSQL = strSQL & "         LEFT JOIN " & strDataSource & ".dbo.TANT_MAST JUN_TANT ON "
    strSQL = strSQL & "            JUN_TANT.TANTM_TANTC = YARD_JUN_TANTO AND "
    strSQL = strSQL & "            JUN_TANT.TANTM_BUMOC = '" & strBumonCode & "' "
    strSQL = strSQL & "         LEFT JOIN " & strDataSource & ".dbo.HKEI_TRAN  ON "
    strSQL = strSQL & "            HKEIT_BUMOC = '" & strBumonCode & "'  AND "
    strSQL = strSQL & "            HKEIT_SYO1C = YARD_CODE AND "
    'strSQL = strSQL & "            ISNULL(HKEIT_KAIYD,'29990101') >= '" & Replace(DateSerial(Year(DATE), Month(DATE), 0), "/", "") & "' "
    strSQL = strSQL & "            HKEIT_KAIYD IS NULL "
    strSQL = strSQL & "         LEFT JOIN " & strDataSource & ".dbo.SHIR_MAST  ON "
    strSQL = strSQL & "            SHIRM_BUMOC = 'ｦ' AND "
    strSQL = strSQL & "            SHIRM_SHIRC = HKEIT_SHIRC  "
    strSQL = strSQL & "         LEFT JOIN " & strDataSource & ".dbo.CODE_TABL  ON "
    strSQL = strSQL & "            CODET_SIKBC = '221' AND "
    strSQL = strSQL & "            CODET_CODEC = HKEIT_KANRI "
    
'INSERT 2025/11/21 K.KINEBUCHI start
    strSQL = strSQL & "         LEFT JOIN "
    strSQL = strSQL & "            " & strDataSource & ".dbo.TINH_MAST "
    strSQL = strSQL & "         ON "
    strSQL = strSQL & "             TINHM_BUMOC = '" & strBumonCode & "' "
    strSQL = strSQL & "         AND "
    strSQL = strSQL & "             TINHM_TINTC = YARD_CODE "

    strSQL = strSQL & "         LEFT JOIN "
    strSQL = strSQL & "            " & strDataSource & ".dbo.BUKE_MAST "
    strSQL = strSQL & "         ON "
    strSQL = strSQL & "             BUKEM_BUKHC = TINHM_BUKHC "
    
    strSQL = strSQL & "         LEFT JOIN "
    strSQL = strSQL & "            NAME_MAST C210 "
    strSQL = strSQL & "         ON "
    strSQL = strSQL & "         C210.NAME_ID = '210' "
    strSQL = strSQL & "         AND "
    strSQL = strSQL & "             C210.NAME_CODE = BUKEM_SYOUI "
'INSERT 2025/11/21 K.KINEBUCHI end
    
    strSQL = strSQL & "    LEFT JOIN "
    strSQL = strSQL & "    (     "
    strSQL = strSQL & "    SELECT"
    strSQL = strSQL & "           TRANA.KTPST_YCODE as KTPST_YCODE,"
    strSQL = strSQL & "           TRANA.KTPST_STDTE,"
    strSQL = strSQL & "            KKTHM_TNAME,"
    strSQL = strSQL & "           KKTHM_TCODE"
    strSQL = strSQL & "               FROM KTPS_TRAN TRANA, KKTH_MAST"
    strSQL = strSQL & "            WHERE TRANA.KTPST_TCODE = KKTHM_TCODE"
    strSQL = strSQL & "           AND TRANA.KTPST_STDTE =  ( SELECT MAX(TRANB.KTPST_STDTE) FROM KTPS_TRAN TRANB WHERE TRANB.KTPST_YCODE = TRANA.KTPST_YCODE )) KTPS_KKTH_TBL "
    strSQL = strSQL & "       ON KTPS_KKTH_TBL.KTPST_YCODE = YARD_MAST.YARD_CODE"
    strSQL = strSQL & " Where 1 = 1 "
    strSQL = strSQL & " AND YARD_CODE between " & cmb_CODEFROM & " AND " & cmb_CODETO
    If frm_RentendDay = optRentendDay_not.OptionValue Then
        strSQL = strSQL & " AND YARD_RENTEND_DAY IS NULL                "   '営業終了日
    ElseIf frm_RentendDay = optRentendDay_only.OptionValue Then
        strSQL = strSQL & "  AND  YARD_RENTEND_DAY IS NOT NULL            "   '営業終了日
    End If
    If frm_EndDay = optEndDay_not.OptionValue Then
        strSQL = strSQL & "  AND  YARD_END_DAY IS NULL                    "   'ヤード解約予定日
    ElseIf frm_EndDay = optEndDay_only.OptionValue Then
        strSQL = strSQL & "  AND  YARD_END_DAY IS NOT NULL                "   'ヤード解約予定日
    End If
    
   If Nz(Me.cmb_価格テンプコード, "") <> "" Then
        strSQL = strSQL & "  AND KTPS_KKTH_TBL.KKTHM_TCODE = '" & Nz(Me.cmb_価格テンプコード, "") & "' "                         '適用価格テンプレート
    End If
    
    
    strSQL = strSQL & " ORDER BY YARD_CODE "
    
    lngCnt = ADODB_CreateWorkTable_CurrentDb(C_TEMP_TABLE, strSQL, objCon)

 
    '*** データチェック ***
    'アクションクエリーの結果で判断
    If lngCnt < 0 Then
        strMsg = "対象となるデータが存在しませんでした。"
        MsgBox strMsg, vbInformation + vbOKOnly
        GoTo Exit_Make_Work
    End If
    
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
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
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
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)
'-----------------------------------------------------------'
'  Check_Dateでは'198/03といった形が変換されるための補助検査
'-----------------------------------------------------------'
  
    Dim sYear As String
      
    CMCheck_Date = Null
      
    If Nz(vDATE, "") = "" Then
        Exit Function
    End If
      
    If IsDate(vDATE) = False Then
        Exit Function
    End If
      
    sYear = Format(vDATE, "yyyy")
    
    If CmLenb(sYear) <> 4 Then
        Exit Function
    End If
      
    CMCheck_Date = vDATE

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
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
On Error GoTo Err_CmBCheck_Date
  
    Dim sYear  As String
    Dim intLen As Integer
    
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

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルの削除
'       MODULE_ID       : subDropWorkTable
'       CREATE_DATE     : 2011/04/11        M.RYU
'       PARAM           : dbAccess          Access側データベース
'                       : strTable          テーブル名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDropWorkTable(ByVal dbAccess As Database)
    On Error Resume Next
    
    dbAccess.TableDefs.Delete WK_YARD_MAST
    Select Case Err.Number
    Case 0
    Case 3265
        Err.Clear
    Case Else
        MsgBox "subDropWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
        Err.Clear
    End Select
End Sub


'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_WebPageFROM_AfterUpdate
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageFROM_AfterUpdate()
On Error GoTo Err_txt_WebPageFROM_AfterUpdate

    '*-Nullチェック-*
    If IsNull(Me.txt_WebPageFROM) Or Me.txt_WebPageFROM = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数を使用-*
    Me.txt_WebPageFROM = Check_Date(Me.txt_WebPageFROM)
    
    Exit Sub
    
Err_txt_WebPageFROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageFROM_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_txt_WebPageFROM_BeforeUpdate

    Dim vDATE As Variant
    vDATE = Nz(Me.txt_WebPageFROM, "")
    
    '*-Nullチェック-*
    If vDATE = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数を使用-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)

    '*-入力チェック-*
    '入力チェック
    Select Case CmBCheck_Date(vDATE)
        Case 0
      
        Case 2
        '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_WebPageFROM_BeforeUpdate
        Case 3, 9
            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_WebPageFROM_BeforeUpdate
    End Select

Exit_txt_WebPageFROM_BeforeUpdate:
    
    Exit Sub
    
Err_txt_WebPageFROM_BeforeUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageFROM_GotFocus()

    '*-YY/MM/DDの形に変換-*
    Me.txt_WebPageFROM = Format(Me.txt_WebPageFROM, "YYYY/MM/DD")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageFROM_KeyPress(KeyAscii As Integer)

    Call KeyPress_For_Date(Me.txt_WebPageFROM, KeyAscii, 8)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageFROM_LostFocus()

    '*-YY/MM/DDの形に変換-*
    Me.txt_WebPageFROM = Format(Me.txt_WebPageFROM, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageTO_AfterUpdate()
On Error GoTo Err_txt_WebPageTO_AfterUpdate

    '*-Nullチェック-*
    If IsNull(Me.txt_WebPageTO) Or Me.txt_WebPageTO = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数-*
    Me.txt_WebPageTO = Check_Date(Me.txt_WebPageTO)

    Exit Sub

Err_txt_WebPageTO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageTO_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_txt_WebPageTO_BeforeUpdate

    Dim vDATE As Variant
    
    vDATE = Nz(Me.txt_WebPageTO, "")
    
    '*-Nullチェック-*
    If vDATE = "" Then
        Exit Sub
    End If
    
    '*-Check_Date関数をしよう-*
    vDATE = Check_Date(vDATE)
    vDATE = CMCheck_Date(vDATE)

    '*-入力チェック-*
    '入力チェック
    Select Case CmBCheck_Date(vDATE)
        Case 0
      
        Case 2
        '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_WebPageTO_BeforeUpdate
        Case 3, 9
            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_txt_WebPageTO_BeforeUpdate
    End Select

Exit_txt_WebPageTO_BeforeUpdate:
    Exit Sub
    
Err_txt_WebPageTO_BeforeUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageTO_GotFocus()

    '**--YY/MM/DDの形に変換--**
    Me.txt_WebPageTO = Format(Me.txt_WebPageTO, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageTO_LostFocus()

    '**--YYYY/MM/DDの形に変換--**
    Me.txt_WebPageTO = Format(Me.txt_WebPageTO, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageTO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_txt_WebPageTO_KeyDown
    
    Dim strChdate As String

    strChdate = Nz(Me.txt_WebPageTO, "")
    
    If Shift <> 0 And KeyCode = vbKeyTab Then
        Me.txt_WebPageFROM.SetFocus
        KeyCode = 0
        Exit Sub
    End If
    
Exit_txt_WebPageTO_KeyDown:
   Exit Sub

Err_txt_WebPageTO_KeyDown:
    If Err.Number <> 2110 Then
        MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    End If

    Resume Exit_txt_WebPageTO_KeyDown
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :2011/04/11        M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_WebPageTO_KeyPress(KeyAscii As Integer)
    
    Call KeyPress_For_Date(Me.txt_WebPageTO, KeyAscii, 8)

End Sub
'INS 2017/10/12 S.WATANABE<S>
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_価格テンプコード_AfterUpdate()
On Error GoTo Err_cmb_価格テンプコード_AfterUpdate
    
    Me.txt_価格テンプ名称 = Me.cmb_価格テンプコード.Column(1)
    
Exit_cmb_価格テンプコード_AfterUpdate:
    Exit Sub

Err_cmb_価格テンプコード_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_価格テンプコード_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_価格テンプコード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_価格テンプコード, KeyAscii, 6)
End Sub
'INS 2017/10/12 S.WATANABE<E>
'****************************  ended or program ********************************

