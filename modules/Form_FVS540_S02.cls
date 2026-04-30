VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS540_S02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約一覧サブ(検索条件)
'   プログラムＩＤ　：　FVS540_S02
'   作　成　日　　　：  2005/12/20
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/1/24
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.1
'   修　正　内　容　：  1.検索条件に移動理由追加
'
'修正履歴
'   修　正　日　　　：  2006/04/06
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.2
'   修　正　内　容　：  1.条件追加(4/6)
'                   :    ・鍵配送日
'                   :    ・鍵配送状況
'                   :    ・伝票番号
'
'   修　正　日　　　：  2007/08/23
'   修　正　者　　　：  tajima
'   Ver             ：  0.3
'   修　正　内　容　：  1.条件追加
'                   :    ・受付種別
'                   :    ・申込書送付方法
'                   :    ・受付番号
'
'   修　正　日　　　：  2015/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  1.3
'   修　正　内　容　：  条件に身分証明書を確認する必要のあるもののみ表示を追加
'
'   修　正　日　　　：  2018/05/23
'   修　正　者　　　：  K.SHIBUYA
'   Ver             ：  1.4
'   修　正　内　容　：  顧客フリガナ・顧客コードの表示場所変更
'
'   修　正　日　　　：  2022/08/22
'   修　正　者　　　：  N.IMAI
'   Ver             ：  1.6
'   修　正　内　容　：  検索条件追加：初期費用請求方法、入金確認日、メール発行日
'
'**********************************************
Option Compare Database
Option Explicit

'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant

Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

Private Const P_ZERO_SUPPRESS       As String = "0000000000"    '不足桁０埋め用
'桁数
Private Const P_AUTO_LENGTH         As Integer = 1      '自動
Private Const P_DIVISION_LENGTH     As Integer = 2      '区分
Private Const P_PERSON_LENGTH       As Integer = 3      '担当コード
Private Const P_CODE_LENGTH         As Integer = 6      'ヤードコード/顧客コード
Private Const P_DATE_LENGTH         As Integer = 10      '日付
Private Const P_TEL_LENGTH          As Integer = 15     '電話番号
Private Const P_KANA_LENGTH         As Integer = 18       '顧客カナ

'2006/04/06 SHIBAZAKI 追加
Private Const P_SLIPNO_LENGTH       As Integer = 14      '伝票番号

'==============================================================================*
'
'        MODULE_NAME      :AfterUpdateCommon
'        機能             :更新後イベント時の処理
'                         :数値項目は０サプレスした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
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
    Case 1  '数値項目は０埋め
        objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    Case 2
        AfterUpdateCommon = Format(VarChangeDate, "yyyy/mm/dd")
    Case Else
        AfterUpdateCommon = Nz(objTarget)
    End Select
    
    'コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        If Nz(objTarget) <> "" And _
           objTarget.ListCount <> 0 _
        Then
            AfterUpdateCommon = Nz(objTarget.Column(1, objTarget.ListIndex))
        End If
    End If

End Function

    
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
'        MODULE_NAME      :ChkComboInput
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
Private Function ChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer, ByRef Cancel As Integer) As Integer

    Dim strCode     As String
    Dim intRet      As Integer
    Dim intCount    As Integer  '06.01.22 SHIBAZAKI 追加
    Dim strValue    As String   '06.01.22 SHIBAZAKI 追加
    Dim blnFound    As Boolean  '06.01.22 SHIBAZAKI 追加
    
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
        '06.01.22 SHIBAZAKI 追加↓
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
        '06.01.22 SHIBAZAKI 追加↑
        
        '06.01.22 SHIBAZAKI 変更
        'If cmbCheck.ListIndex = -1 Then
        If Not blnFound Then
            MsgBox "該当するコードがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
            intRet = 2
            GoTo ChkComboInput_End
        End If
    End If
    
ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If
    
    ChkComboInput = intRet

End Function

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


'==============================================================================*
'
'        MODULE_NAME      :cmb_Auto_AfterUpdate
'        機能             :自動 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Auto_AfterUpdate()

    txt_Auto = AfterUpdateCommon(cmb_AUTO, P_AUTO_LENGTH, 1)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Auto_BeforeUpdate
'        機能             :自動 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Auto_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_AUTO, P_AUTO_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Auto_KeyPress
'        機能             :自動 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_AUTO_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_AUTO, KeyAscii, P_AUTO_LENGTH)

End Sub

'06.1.24 SHIBAZAKI 追加↓
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveFrom_AfterUpdate
'        機能             :移動理由FROM AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveFrom_AfterUpdate()
    
    txt_MoveFrom = AfterUpdateCommon(cmb_MoveFrom, P_DIVISION_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveFrom_BeforeUpdate
'        機能             :移動理由FROM BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveFrom_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_MoveFrom, P_DIVISION_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveFrom_KeyPress
'        機能             :移動理由FROM KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveFrom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_MoveFrom, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveTo_AfterUpdate
'        機能             :移動理由To AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveTo_AfterUpdate()
    
    txt_MoveTo = AfterUpdateCommon(cmb_MoveTo, P_DIVISION_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveTo_BeforeUpdate
'        機能             :移動理由To BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveTo_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_MoveTo, P_DIVISION_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveTo_KeyPress
'        機能             :移動理由To KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveTo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_MoveTo, KeyAscii, P_DIVISION_LENGTH)

End Sub
'06.1.24 SHIBAZAKI 追加↑

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tanto_AfterUpdate
'        機能             :受付担当者 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tanto_AfterUpdate()

    txt_Tanto = AfterUpdateCommon(cmb_Tanto, P_PERSON_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tanto_BeforeUpdate
'        機能             :受付担当者 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tanto_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_Tanto, P_PERSON_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tanto_KeyPress
'        機能             :受付担当者 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tanto_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_Tanto, KeyAscii, P_PERSON_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Ucode_AfterUpdate
'        機能             :顧客コード AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ucode_AfterUpdate()
    
    txt_Uname = AfterUpdateCommon(cmb_Ucode, P_CODE_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Ucode_BeforeUpdate
'        機能             :顧客コード BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ucode_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_Ucode, P_CODE_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Ucode_KeyPress
'        機能             :顧客コード KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ucode_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_Ucode, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YukbnFrom_AfterUpdate
'        機能             :受付状態FROM AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YukbnFrom_AfterUpdate()
    
    txt_UkkbnFrom = AfterUpdateCommon(cmb_YukbnFrom, P_DIVISION_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YukbnFrom_BeforeUpdate
'        機能             :受付状態FROM BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YukbnFrom_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_YukbnFrom, P_DIVISION_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YukbnFrom_KeyPress
'        機能             :受付状態FROM KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YukbnFrom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_YukbnFrom, KeyAscii, P_DIVISION_LENGTH)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_YukbnTo_AfterUpdate
'        機能             :受付状態To AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YukbnTo_AfterUpdate()
    
    txt_UkkbnTo = AfterUpdateCommon(cmb_YukbnTo, P_DIVISION_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YukbnTo_BeforeUpdate
'        機能             :受付状態To BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YukbnTo_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_YukbnTo, P_DIVISION_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_YukbnTo_KeyPress
'        機能             :受付状態To KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YukbnTo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_YukbnTo, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_AfterUpdate
'        機能             :ヤードコード AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_AfterUpdate()
    
    txt_Yard = AfterUpdateCommon(cmb_Yard, P_CODE_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_BeforeUpdate
'        機能             :ヤードコード BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_Yard, P_CODE_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_KeyPress
'        機能             :ヤードコード KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_Yard, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSeikyuCd_AfterUpdate
'        機能             :初期費用請求方法 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSeikyuCd_AfterUpdate()

    Call ComboBox_AfterUpdate(cmbSeikyuCd, txtSeikyuName.NAME)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'F1=「終了」 F2=「検索」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Parent.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Parent.cmd_検索.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kana_AfterUpdate
'        機能             :顧客カナ AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kana_AfterUpdate()

    Call AfterUpdateCommon(txt_Kana, P_KANA_LENGTH, 3)
    
    cmb_Ucode = ""
    txt_Uname = ""
    cmb_Ucode.Requery
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kana_BeforeUpdate
'        機能             :顧客カナ BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kana_BeforeUpdate(Cancel As Integer)

    Call chkTextInput(txt_Kana, P_KANA_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Kana_KeyPress
'        機能             :顧客カナ KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kana_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Char(txt_Kana, KeyAscii, P_KANA_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Tel_AfterUpdate
'        機能             :電話番号 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_AfterUpdate()

    Call AfterUpdateCommon(txt_Tel, P_TEL_LENGTH, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Tel_BeforeUpdate
'        機能             :電話番号 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_BeforeUpdate(Cancel As Integer)

    Call chkTextInput(txt_Tel, P_TEL_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Tel_KeyPress
'        機能             :電話番号 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Tel(txt_Tel, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateFrom_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateFrom_AfterUpdate()

    txt_UkdateFrom = AfterUpdateCommon(txt_UkdateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateFrom_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_UkdateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateFrom_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateFrom_GotFocus()

    '体裁を整える
    txt_UkdateFrom = Format(txt_UkdateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateFrom_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_UkdateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateFrom_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateFrom_LostFocus()

    '体裁を整える
    txt_UkdateFrom = Format(Me.txt_UkdateFrom, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateTo_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateTo_AfterUpdate()

    txt_UkdateTo = AfterUpdateCommon(txt_UkdateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateTo_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_UkdateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateTo_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateTo_GotFocus()

    '体裁を整える
    txt_UkdateTo = Format(txt_UkdateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateTo_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_UkdateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateTo_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkdateTo_LostFocus()

    '体裁を整える
    txt_UkdateTo = Format(Me.txt_UkdateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateFrom_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateFrom_AfterUpdate()

    txt_KudateFrom = AfterUpdateCommon(txt_KudateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateFrom_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_KudateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateFrom_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateFrom_GotFocus()

    '体裁を整える
    txt_KudateFrom = Format(txt_KudateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateFrom_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_KudateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateFrom_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateFrom_LostFocus()

    '体裁を整える
    txt_KudateFrom = Format(Me.txt_KudateFrom, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateTo_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateTo_AfterUpdate()

    txt_KudateTo = AfterUpdateCommon(txt_KudateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateTo_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_KudateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateTo_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateTo_GotFocus()

    '体裁を整える
    txt_KudateTo = Format(txt_KudateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateTo_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_KudateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KudateTo_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KudateTo_LostFocus()

    '体裁を整える
    txt_KudateTo = Format(Me.txt_KudateTo, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateFrom_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateFrom_AfterUpdate()

    txt_YtdateFrom = AfterUpdateCommon(txt_YtdateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateFrom_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_YtdateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateFrom_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateFrom_GotFocus()

    '体裁を整える
    txt_YtdateFrom = Format(txt_YtdateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateFrom_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_YtdateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateFrom_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateFrom_LostFocus()

    '体裁を整える
    txt_YtdateFrom = Format(Me.txt_YtdateFrom, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateTo_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateTo_AfterUpdate()

    txt_YtdateTo = AfterUpdateCommon(txt_YtdateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateTo_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_YtdateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateTo_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateTo_GotFocus()

    '体裁を整える
    txt_YtdateTo = Format(txt_YtdateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateTo_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_YtdateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YtdateTo_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtdateTo_LostFocus()

    '体裁を整える
    txt_YtdateTo = Format(Me.txt_YtdateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateFrom_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateFrom_AfterUpdate()

    txt_KkdateFrom = AfterUpdateCommon(txt_KkdateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateFrom_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_KkdateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateFrom_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateFrom_GotFocus()

    '体裁を整える
    txt_KkdateFrom = Format(txt_KkdateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateFrom_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_KkdateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateFrom_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateFrom_LostFocus()

    '体裁を整える
    txt_KkdateFrom = Format(Me.txt_KkdateFrom, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateTo_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateTo_AfterUpdate()

    txt_KkdateTo = AfterUpdateCommon(txt_KkdateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateTo_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_KkdateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateTo_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateTo_GotFocus()

    '体裁を整える
    txt_KkdateTo = Format(txt_KkdateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateTo_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_KkdateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KkdateTo_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkdateTo_LostFocus()

    '体裁を整える
    txt_KkdateTo = Format(Me.txt_KkdateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateFrom_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateFrom_AfterUpdate()

    txt_TkdateFrom = AfterUpdateCommon(txt_TkdateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateFrom_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_TkdateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateFrom_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateFrom_GotFocus()

    '体裁を整える
    txt_TkdateFrom = Format(txt_TkdateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateFrom_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_TkdateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateFrom_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateFrom_LostFocus()

    '体裁を整える
    txt_TkdateFrom = Format(Me.txt_TkdateFrom, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateTo_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateTo_AfterUpdate()

    txt_TkdateTo = AfterUpdateCommon(txt_TkdateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateTo_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_TkdateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateTo_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateTo_GotFocus()

    '体裁を整える
    txt_TkdateTo = Format(txt_TkdateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateTo_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_TkdateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TkdateTo_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkdateTo_LostFocus()

    '体裁を整える
    txt_TkdateTo = Format(Me.txt_TkdateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateFrom_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateFrom_AfterUpdate()

    txt_TadateFrom = AfterUpdateCommon(txt_TadateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateFrom_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_TadateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateFrom_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateFrom_GotFocus()

    '体裁を整える
    txt_TadateFrom = Format(txt_TadateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateFrom_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_TadateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateFrom_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateFrom_LostFocus()

    '体裁を整える
    txt_TadateFrom = Format(Me.txt_TadateFrom, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_TadateTo_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateTo_AfterUpdate()

    txt_TadateTo = AfterUpdateCommon(txt_TadateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateTo_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_TadateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateTo_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateTo_GotFocus()

    '体裁を整える
    txt_TadateTo = Format(txt_TadateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateTo_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_TadateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_TadateTo_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TadateTo_LostFocus()

    '体裁を整える
    txt_TadateTo = Format(Me.txt_TadateTo, "yyyy/mm/dd")

End Sub


'2006/04/06 SHIBAZAKI 追加↓
'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateFrom_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateFrom_AfterUpdate()

    txt_KeySendDateFrom = AfterUpdateCommon(txt_KeySendDateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateFrom_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_KeySendDateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateFrom_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateFrom_GotFocus()

    '体裁を整える
    txt_KeySendDateFrom = Format(txt_KeySendDateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateFrom_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_KeySendDateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateFrom_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateFrom_LostFocus()

    '体裁を整える
    txt_KeySendDateFrom = Format(Me.txt_KeySendDateFrom, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateTo_AfterUpdate
'        機能             :受付日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateTo_AfterUpdate()

    txt_KeySendDateTo = AfterUpdateCommon(txt_KeySendDateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateTo_BeforeUpdate
'        機能             :受付日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_KeySendDateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateTo_GotFocus
'        機能             :受付日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateTo_GotFocus()

    '体裁を整える
    txt_KeySendDateTo = Format(txt_KeySendDateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateTo_KeyPress
'        機能             :受付日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_KeySendDateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeySendDateTo_LostFocus
'        機能             :受付日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeySendDateTo_LostFocus()

    '体裁を整える
    txt_KeySendDateTo = Format(Me.txt_KeySendDateTo, "yyyy/mm/dd")

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_SlipNo_AfterUpdate
'        機能             :txt_SlipNo_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SlipNo_AfterUpdate()

    Call AfterUpdateCommon(txt_SlipNo, P_SLIPNO_LENGTH, 3)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_SlipNo_BeforeUpdate
'        機能             :txt_SlipNo_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SlipNo_BeforeUpdate(Cancel As Integer)

    Call chkTextInput(txt_SlipNo, P_SLIPNO_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_SlipNo_KeyPress
'        機能             :txt_SlipNo_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SlipNo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char(txt_SlipNo, KeyAscii, P_SLIPNO_LENGTH)

End Sub
'2006/04/06 SHIBAZAKI 追加↑

'★☆★☆★☆★☆★☆★★☆☆ 2007/08/23 add tajima ★☆★☆★☆★☆★☆★☆★☆
'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_AfterUpdate
'        機能             :cmb_受付種別_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_AfterUpdate()
    Me.txt_受付種別 = AfterUpdateCommon(cmb_受付種別, P_DIVISION_LENGTH, 1)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_BeforeUpdate
'        機能             :cmb_受付種別_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_BeforeUpdate(Cancel As Integer)
    Call ChkComboInput(cmb_受付種別, P_DIVISION_LENGTH, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_KeyPress
'        機能             :cmb_受付種別_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_受付種別, KeyAscii, P_DIVISION_LENGTH)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_申込書送付方_AfterUpdate
'        機能             :cmb_申込書送付方_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_申込書送付方_AfterUpdate()
    Me.txt_申込書送付方 = AfterUpdateCommon(cmb_申込書送付方, P_DIVISION_LENGTH, 1)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_申込書送付方_BeforeUpdate
'        機能             :cmb_申込書送付方_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_申込書送付方_BeforeUpdate(Cancel As Integer)
    Call ChkComboInput(cmb_申込書送付方, P_DIVISION_LENGTH, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_申込書送付方_KeyPress
'        機能             :cmb_申込書送付方_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_申込書送付方_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_申込書送付方, KeyAscii, P_DIVISION_LENGTH)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_受付番号_AfterUpdate
'        機能             :txt_受付番号_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_受付番号_AfterUpdate()
    Call AfterUpdateCommon(txt_受付番号, 11, 3)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_受付番号_BeforeUpdate
'        機能             :txt_受付番号_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_受付番号_BeforeUpdate(Cancel As Integer)
    Call chkTextInput(txt_受付番号, 11, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_受付番号_KeyPress
'        機能             :txt_受付番号_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_受付番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Char(txt_受付番号, KeyAscii, 11)
End Sub

'INSERT 2022/08/22 N.IMAI Start
'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateFrom_AfterUpdate
'        機能             :入金確認日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateFrom_AfterUpdate()

    txt_PaydateFrom = AfterUpdateCommon(txt_PaydateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateFrom_BeforeUpdate
'        機能             :入金確認日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_PaydateFrom, Cancel)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateFrom_GotFocus
'        機能             :入金確認日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateFrom_GotFocus()

    '体裁を整える
    txt_PaydateFrom = Format(txt_PaydateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateFrom_KeyPress
'        機能             :入金確認日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_PaydateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateFrom_LostFocus
'        機能             :入金確認日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateFrom_LostFocus()

    '体裁を整える
    txt_PaydateFrom = Format(Me.txt_PaydateFrom, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateTo_AfterUpdate
'        機能             :入金確認日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateTo_AfterUpdate()

    txt_PaydateTo = AfterUpdateCommon(txt_PaydateTo, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateTo_BeforeUpdate
'        機能             :入金確認日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_PaydateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateTo_GotFocus
'        機能             :入金確認日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateTo_GotFocus()

    '体裁を整える
    txt_PaydateTo = Format(txt_PaydateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateTo_KeyPress
'        機能             :入金確認日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_PaydateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_PaydateTo_LostFocus
'        機能             :入金確認日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PaydateTo_LostFocus()

    '体裁を整える
    txt_PaydateTo = Format(Me.txt_PaydateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateFrom_AfterUpdate
'        機能             :メール配信日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateFrom_AfterUpdate()

    txt_YamatodateFrom = AfterUpdateCommon(txt_YamatodateFrom, P_DATE_LENGTH, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateFrom_GotFocus
'        機能             :メール配信日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateFrom_GotFocus()

    '体裁を整える
    txt_YamatodateFrom = Format(txt_YamatodateFrom, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateFrom_KeyPress
'        機能             :メール配信日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateFrom_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_YamatodateFrom, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateFrom_LostFocus
'        機能             :メール配信日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateFrom_LostFocus()

    '体裁を整える
    txt_YamatodateFrom = Format(Me.txt_YamatodateFrom, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateFrom_BeforeUpdate
'        機能             :メール配信日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateFrom_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_YamatodateFrom, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateTo_GotFocus
'        機能             :メール配信日 GotFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateTo_GotFocus()

    '体裁を整える
    txt_YamatodateTo = Format(txt_YamatodateTo, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateTo_KeyPress
'        機能             :メール配信日 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateTo_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_YamatodateTo, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateTo_LostFocus
'        機能             :メール配信日 LostFocus
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateTo_LostFocus()

    '体裁を整える
    txt_YamatodateTo = Format(Me.txt_YamatodateTo, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateTo_BeforeUpdate
'        機能             :メール配信日 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateTo_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_YamatodateTo, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_YamatodateTo_AfterUpdate
'        機能             :メール配信日 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YamatodateTo_AfterUpdate()

    txt_YamatodateTo = AfterUpdateCommon(txt_YamatodateTo, P_DATE_LENGTH, 2)

End Sub
'INSERT 2022/08/22 N.IMAI End
