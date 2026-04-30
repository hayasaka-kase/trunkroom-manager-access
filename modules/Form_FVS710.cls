VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS710"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　解約条件入力
'   プログラムＩＤ　：　FVS710
'   作　成　日　　　：  2005/07/11
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：2006/2/3
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：１．顧客カナをコンボボックスからテキストボックスに変更
'                   ：２．顧客コードコンボボックスを顧客カナで部分一致でリスト設定する
'                   ：３．顧客コードコンボボックスが空の場合にリスト存在チェックは行わない
'                   ：４．顧客名のテキストボックス追加
'   Ver             ： 0.2
'
'**********************************************
'修正履歴
'   修　正　日　　　：2012/07/07
'   修　正　者      ：M.HONDA
'   修　正　内　容　：１．部屋Noを検索条件に追加
'   Ver             ： 0.3
'**********************************************
'修正履歴
'   修　正　日　　　：2013/06/03
'   修　正　者      ：M.HONDA
'   修　正　内　容　：１．終了メッセージを表示させないように対応。
'   Ver             ： 0.4
'**********************************************
'修正履歴
'   修　正　日　　　：2019/04/10
'   修　正　者      ：Y_SUZUKI
'   修　正　内　容　：検索条件に「解約担当者」を追加
'   Ver             ： 0.5
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Const P_PROG_ID             As String = "FVS710"
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

'桁数
Private Const P_DIVISION_LENGTH     As Integer = 2      '区分
Private Const P_PERSON_LENGTH       As Integer = 3      '担当コード
Private Const P_CODE_LENGTH         As Integer = 6      'ヤードコード／コンテナ番号
Private Const P_DATE_LENGTH         As Integer = 10      '日付
Private Const P_YM_LENGTH           As Integer = 7       '年月
Private Const P_ZERO_SUPPRESS       As String = "0000000000"    '不足桁０埋め用

'2006.2.3 SHIBAZAKI 追加
Private Const P_KANA_LENGTH         As Integer = 18       '顧客カナ

'==============================================================================*
'
'        MODULE_NAME      :chk_MaintenanceEndNull_AfterUpdate
'        機能             :メンテ完了日未入力のみ AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_MaintenanceEndNull_AfterUpdate()

    txt_MaintenanceEndDateFrom.Enabled = Not chk_MaintenanceEndNull.VALUE
    txt_MaintenanceEndDateTo.Enabled = Not chk_MaintenanceEndNull.VALUE

End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_MaintenanceRequestNull_AfterUpdate
'        機能             :メンテ依頼日未入力のみ AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_MaintenanceRequestNull_AfterUpdate()

    txt_MaintenanceRequestDateFrom.Enabled = Not chk_MaintenanceRequestNull.VALUE
    txt_MaintenanceRequestDateTo.Enabled = Not chk_MaintenanceRequestNull.VALUE
    
End Sub

'==============================================================================*
'ADD 2019/04/10 Y_SUZUKI EGL
'        MODULE_NAME      :cmb_KaiyakuTanto_AfterUpdate
'        機能             :解約受付担当 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KaiyakuTanto_AfterUpdate()
    txt_KaiyakuTanto = AfterUpdateCommon(cmb_KaiyakuTanto, P_PERSON_LENGTH, 1)

End Sub

'==============================================================================*
'ADD 2019/04/10 Y_SUZUKI EGL
'        MODULE_NAME      :cmb_KaiyakuTanto_BeforeUpdate
'        機能             :解約受付担当 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KaiyakuTanto_BeforeUpdate(Cancel As Integer)
    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_KaiyakuTanto, P_PERSON_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する担当者がありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If

End Sub

'==============================================================================*
'ADD 2019/04/10 Y_SUZUKI EGL
'        MODULE_NAME      :cmb_KaiyakuTanto_KeyPress
'        機能             :解約受付担当 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KaiyakuTanto_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_KaiyakuTanto, KeyAscii, P_PERSON_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_UserCode_AfterUpdate
'        機能             :顧客コード AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_UserCode_AfterUpdate()

'2006.2.3 SHIBAZAKI 変更
'    Me.cmb_UserCode = Format(Me.cmb_UserCode, Mid(P_ZERO_SUPPRESS, 1, P_CODE_LENGTH))
    txt_UserName = AfterUpdateCommon(cmb_UserCode, P_CODE_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_UserCode_BeforeUpdate
'        機能             :顧客コード BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_UserCode_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_UserCode, P_CODE_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する顧客マスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_UserCode_KeyPress
'        機能             :顧客コード KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_UserCode_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num(cmb_UserCode, KeyAscii, P_CODE_LENGTH)

End Sub

'2006.2.3 SHIBAZAKI 削除
''==============================================================================*
''
''        MODULE_NAME      :cmb_UserKana_AfterUpdate
''        機能             :顧客カナ AfterUpdate ｲﾍﾞﾝﾄ処理
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_UserKana_AfterUpdate()
'
'    Me.cmb_UserCode.Requery
'
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YardFrom_AfterUpdate
'        機能             :ヤードコードFROM AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardFrom_AfterUpdate()

    txt_YardFrom = AfterUpdateCommon(cmb_YardFrom, P_CODE_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YardFrom_BeforeUpdate
'        機能             :ヤードコードFROM BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardFrom_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_YardFrom, P_CODE_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YardFrom_KeyPress
'        機能             :ヤードコードFROM KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardFrom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_YardFrom, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YardTo_AfterUpdate
'        機能             :ヤードコードTO AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardTo_AfterUpdate()

    txt_YardTo = AfterUpdateCommon(cmb_YardTo, P_CODE_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YardTo_BeforeUpdate
'        機能             :ヤードコードTO BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardTo_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_YardTo, P_CODE_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YardTo_KeyPress
'        機能             :ヤードコードTO KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardTo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_YardTo, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_検索開始_Click
'        機能             :検索開始 Click ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索開始_Click()

    Dim strParam As String
    
    strParam = strParam & txt_CancelDateFrom & ","
    strParam = strParam & txt_CancelDateTo & ","
    strParam = strParam & cmb_YardFrom & ","
    strParam = strParam & cmb_YardTo & ","
    strParam = strParam & cmb_UserCode & ","
    strParam = strParam & txt_MaintenanceRequestDateFrom & ","
    strParam = strParam & txt_MaintenanceRequestDateTo & ","
    strParam = strParam & txt_MaintenanceEndDateFrom & ","
    strParam = strParam & txt_MaintenanceEndDateTo & ","
    
    strParam = strParam & cmb_ContractState & ","
    strParam = strParam & opg_SortOrder.VALUE & ","
    strParam = strParam & chk_MaintenanceRequestNull.VALUE & ","
    strParam = strParam & chk_MaintenanceEndNull.VALUE
    
    '2006.2.3 SHIBAZAKI 追加
    strParam = strParam & "," & IIf(Nz(cmb_UserCode, "") = "", txt_UserKana, "")
    
    strParam = strParam & "," & txt_CON        '' 2012/07/07 M.HONDA INS
    
    strParam = strParam & "," & cmb_KaiyakuTanto        'ADD 2019/04/10 Y_SUZUKI EGL
    
    Call FormEnable(False)
    doCmd.OpenForm "FVS720", , , , , , strParam

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        機能             :終了 Click ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next


    doCmd.Close acForm, Me.FormName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form Activate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()

    doCmd.Maximize

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form KeyDown ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'F1=「終了」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload
'        機能             :Form Unload ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Form_Unload
    
    Dim strMsg  As String
    Dim lngRc   As Long
    
    '' 2013/06/03 M.HONDA START
'    '終了メッセージを表示
'    strMsg = "終了します。よろしいですか？"
'    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, P_KAKUNIN)
'
'    If lngRc = vbNo Then
'       Cancel = True
'    End If
    '' 2013/06/03 M.HONDA END
    
    Exit Sub
    
Err_Form_Unload:
    MsgBox Error$

End Sub
    
'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function dateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    dateCheck = True
    
    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        dateCheck = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :AfterUpdateCommon
'        機能             :更新後イベント時の処理
'        IN               :①対象コントロール
'                         :②入力最大桁数
'                         :③1=数値項目 2=日付項目 3=テキスト項目
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function AfterUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer) As String

    AfterUpdateCommon = ""
    
    '体裁を整える
    Select Case intKind
    Case 1
        objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    Case 2
        AfterUpdateCommon = Format(VarChangeDate, "yyyy/mm/dd")
    Case Else
        AfterUpdateCommon = Nz(objTarget)
    End Select
    
    'コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        AfterUpdateCommon = Nz(objTarget.Column(1, objTarget.ListIndex))
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateFrom_AfterUpdate
'        機能             :解約日FROM AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateFrom_AfterUpdate()

    txt_CancelDateFrom = AfterUpdateCommon(txt_CancelDateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateFrom_BeforeUpdate
'        機能             :解約日FROM BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_CancelDateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateFrom_GotFocus
'        機能             :解約日FROM GotFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateFrom_GotFocus()

    '体裁を整える
    txt_CancelDateFrom = Format(txt_CancelDateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateFrom_KeyPress
'        機能             :解約日FROM KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(txt_CancelDateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateFrom_LostFocus
'        機能             :解約日FROM LostFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateFrom_LostFocus()

    '体裁を整える
    txt_CancelDateFrom = Format(txt_CancelDateFrom, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateTo_AfterUpdate
'        機能             :解約日TO AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateTo_AfterUpdate()

    txt_CancelDateTo = AfterUpdateCommon(txt_CancelDateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateTo_BeforeUpdate
'        機能             :解約日TO BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_CancelDateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateTo_GotFocus
'        機能             :解約日TO GotFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateTo_GotFocus()

    '体裁を整える
    txt_CancelDateTo = Format(txt_CancelDateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateTo_KeyPress
'        機能             :解約日TO KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(txt_CancelDateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_CancelDateTo_LostFocus
'        機能             :解約日TO LostFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CancelDateTo_LostFocus()

    '体裁を整える
    txt_CancelDateTo = Format(txt_CancelDateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ChkComboInput
'        機能             :コンボボックス入力チェック
'        IN               :対象コンボボックス
'        OUT              :0=正常
'                          1=入力規則に反している
'                          2=入力値がリストに存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer) As Integer

    Dim strCode As String
    Dim intRet  As Integer
    Dim intCount    As Integer  '06.02.3 SHIBAZAKI 追加
    Dim strValue    As String   '06.02.3 SHIBAZAKI 追加
    Dim blnFound    As Boolean  '06.02.3 SHIBAZAKI 追加
    
    intRet = 0
    
    strCode = Nz(cmbCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkComboInput_End
    End If
    
    '入力チェック
    Select Case CmNumCheck(strCode, intMaxLength, 0)
        Case 3
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            intRet = 1
            GoTo ChkComboInput_End
        Case 2
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            SendKeys "{ESC}"
            intRet = 1
            GoTo ChkComboInput_End
    End Select
    
    '入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        '06.2.3 SHIBAZAKI 追加↓
        blnFound = False
        '入力値を０サプレス
        strValue = Format(cmbCheck, Mid(P_ZERO_SUPPRESS, 1, intMaxLength))
        'コンボボックス内にコードが存在してる？
        For intCount = 0 To cmbCheck.ListCount - 1
            'コード値比較
            If cmbCheck.Column(0, intCount) = strValue Then
                blnFound = True
                Exit For
            End If
        Next
        '06.2.3 SHIBAZAKI 追加↑
        
        '06.2.3 SHIBAZAKI 変更
        'If cmbCheck.ListIndex = -1 Then
        If Not blnFound Then
            intRet = 2
            GoTo ChkComboInput_End
        End If
    End If
    
ChkComboInput_End:
    ChkComboInput = intRet

End Function


'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateFrom_AfterUpdate
'        機能             :メンテ依頼日FROM AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateFrom_AfterUpdate()

    txt_MaintenanceRequestDateFrom = AfterUpdateCommon(txt_MaintenanceRequestDateFrom, P_DATE_LENGTH, 2)
    
    If Nz(txt_MaintenanceRequestDateFrom) <> "" Then
        opg_SortOrder.VALUE = 1
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateFrom_BeforeUpdate
'        機能             :メンテ依頼日FROM BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_MaintenanceRequestDateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateFrom_GotFocus
'        機能             :メンテ依頼日FROM GotFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateFrom_GotFocus()

    '体裁を整える
    txt_MaintenanceRequestDateFrom = Format(txt_MaintenanceRequestDateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateFrom_KeyPress
'        機能             :メンテ依頼日FROM KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(txt_MaintenanceRequestDateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateFrom_LostFocus
'        機能             :メンテ依頼日FROM LostFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateFrom_LostFocus()

    '体裁を整える
    txt_MaintenanceRequestDateFrom = Format(txt_MaintenanceRequestDateFrom, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateTo_AfterUpdate
'        機能             :メンテ依頼日TO AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateTo_AfterUpdate()

    txt_MaintenanceRequestDateTo = AfterUpdateCommon(txt_MaintenanceRequestDateTo, P_DATE_LENGTH, 2)

    If Nz(txt_MaintenanceRequestDateTo) <> "" Then
        opg_SortOrder.VALUE = 1
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateTo_BeforeUpdate
'        機能             :メンテ依頼日TO BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_MaintenanceRequestDateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateTo_GotFocus
'        機能             :メンテ依頼日TO GotFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateTo_GotFocus()

    '体裁を整える
    txt_MaintenanceRequestDateTo = Format(txt_MaintenanceRequestDateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateTo_KeyPress
'        機能             :メンテ依頼日TO KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(txt_MaintenanceRequestDateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDateTo_LostFocus
'        機能             :メンテ依頼日TO LostFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDateTo_LostFocus()

    '体裁を整える
    txt_MaintenanceRequestDateTo = Format(txt_MaintenanceRequestDateTo, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateFrom_AfterUpdate
'        機能             :メンテ完了日FROM AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateFrom_AfterUpdate()

    txt_MaintenanceEndDateFrom = AfterUpdateCommon(txt_MaintenanceEndDateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateFrom_BeforeUpdate
'        機能             :メンテ完了日FROM BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_MaintenanceEndDateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateFrom_GotFocus
'        機能             :メンテ完了日FROM GotFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateFrom_GotFocus()

    '体裁を整える
    txt_MaintenanceEndDateFrom = Format(txt_MaintenanceEndDateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateFrom_KeyPress
'        機能             :メンテ完了日FROM KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(txt_MaintenanceEndDateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateFrom_LostFocus
'        機能             :メンテ完了日FROM LostFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateFrom_LostFocus()

    '体裁を整える
    txt_MaintenanceEndDateFrom = Format(txt_MaintenanceEndDateFrom, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateTo_AfterUpdate
'        機能             :メンテ完了日TO AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateTo_AfterUpdate()

    txt_MaintenanceEndDateTo = AfterUpdateCommon(txt_MaintenanceEndDateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateTo_BeforeUpdate
'        機能             :メンテ完了日TO BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_MaintenanceEndDateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateTo_GotFocus
'        機能             :メンテ完了日TO GotFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateTo_GotFocus()

    '体裁を整える
    txt_MaintenanceEndDateTo = Format(txt_MaintenanceEndDateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateTo_KeyPress
'        機能             :メンテ完了日TO KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(txt_MaintenanceEndDateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDateTo_LostFocus
'        機能             :メンテ完了日TO LostFocus ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDateTo_LostFocus()

    '体裁を整える
    txt_MaintenanceEndDateTo = Format(txt_MaintenanceEndDateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ContractState_AfterUpdate
'        機能             :状態区分 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ContractState_AfterUpdate()

    txt_ContractState = AfterUpdateCommon(cmb_ContractState, P_DIVISION_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ContractState_BeforeUpdate
'        機能             :状態区分 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ContractState_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_ContractState, P_DIVISION_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する状態区分がありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ContractState_KeyPress
'        機能             :状態区分 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ContractState_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_ContractState, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :subReload
'        CREATE_DATE      :2005/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subReload(Optional intFlg As Integer)

    Call FormEnable(True)
    Call chk_MaintenanceRequestNull_AfterUpdate
    Call chk_MaintenanceEndNull_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :項目使用可不可設定
'        MODULE_ID        :FormEnable
'        CREATE_DATE      :2005/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FormEnable(blnEnable As Boolean)

    txt_CancelDateFrom.Enabled = blnEnable
    txt_CancelDateTo.Enabled = blnEnable
    cmb_YardFrom.Enabled = blnEnable
    cmb_YardTo.Enabled = blnEnable
'    cmb_UserKana.Enabled = blnEnable   '2006.2.3 SHIBAZAKI 削除
    txt_UserKana.Enabled = blnEnable    '2006.2.3 SHIBAZAKI 追加
    cmb_UserCode.Enabled = blnEnable
    txt_MaintenanceRequestDateFrom.Enabled = blnEnable
    txt_MaintenanceRequestDateTo.Enabled = blnEnable
    txt_MaintenanceEndDateFrom.Enabled = blnEnable
    txt_MaintenanceEndDateTo.Enabled = blnEnable
    cmb_ContractState.Enabled = blnEnable
    opg_SortOrder.Enabled = blnEnable
    cmd_終了.Enabled = blnEnable
    
    txt_CON.Enabled = blnEnable     '' 2012/07/07
    cmb_KaiyakuTanto.Enabled = blnEnable    'ADD 2019/04/10 Y_SUZUKI EGL

End Sub

'2006.2.3 SHIBAZAKI 追加↓
'==============================================================================*
'
'        MODULE_NAME      :txt_UserKana_AfterUpdate
'        機能             :顧客カナ AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UserKana_AfterUpdate()

    Call AfterUpdateCommon(txt_UserKana, P_KANA_LENGTH, 3)
    
    cmb_UserCode = ""
    Call cmb_UserCode_AfterUpdate
    cmb_UserCode.Requery

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UserKana_BeforeUpdate
'        機能             :顧客カナ BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UserKana_BeforeUpdate(Cancel As Integer)
    
    Call chkTextInput(txt_UserKana, P_KANA_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ChkTextInput
'        機能             :テキストボックス入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function chkTextInput(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean
    
    If Nz(txtCheck) = "" Then
        chkTextInput = True
        Exit Function
    End If
    
    If CmLenb(txtCheck) > intMaxLength Then
        '桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        chkTextInput = False
    Else
        chkTextInput = True
    End If

End Function
'2006.2.3 SHIBAZAKI 追加↑

