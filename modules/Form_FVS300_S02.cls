VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS300_S02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : コンテナ管理システム
'
'        PROGRAM_NAME    : クレジットカード登録メール送信サブ
'        PROGRAM_ID      : FVS300_S01
'
'        CREATE          : 2011/02/23
'        CERATER         : m.honda
'        Ver             : 0.0
'
'        UPDATE          : 2011/07/02
'        UPDATER         : M.RYU
'        Ver             : V0.1
'                        : 検索条件追加　⇒　保証会社コード、保証会社区分
'
'        UPDATE          : 2011/07/21
'        UPDATER         : M.RYU
'        Ver             : V0.2
'                        : メール送信ボタン動作修正（送信後の画面再読込をコメント）
'
'        UPDATE          : 2011/08/30
'        UPDATER         : M.RYU
'        Ver             : V0.3
'                        : 検索条件を追加⇒カード支払開始月・初回カード支払対象抽出
'
'        UPDATE          : 2011/09/27
'        UPDATER         : M.RYU
'        Ver             : V0.4
'                        : 口座振替メール送信を追加
'
'        UPDATE          : 2011/10/03
'        UPDATER         : M.RYU
'        Ver             : V0.5
'                        : 口座振替メール送信ボタンを追加
'
'        UPDATE          : 2011/11/11
'        UPDATER         : M.RYU
'        Ver             : V0.6
'                        : コメント内容を修正
'
'        UPDATE          : 2012/01/05
'        UPDATER         : M.RYU
'        Ver             : V0.7
'                        : 口座振替のurl引数を顧客カナに変更
'
'        UPDATE          : 2013/08/21
'        UPDATER         : M.HONDA
'        Ver             : V0.8
'                        : 月額料金を表示しないように修正
'
'==============================================================================*
Option Compare Database
Option Explicit

Private Const PROG_ID = "FVS300"

'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant

Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

Private Const P_ZERO_SUPPRESS       As String = "0000000000"    '不足桁０埋め用
'桁数
Private Const P_DIVISION_LENGTH     As Integer = 2      '区分
Private Const P_CODE_LENGTH         As Integer = 6      'ヤードコード/顧客コード/保証会社コード
Private Const P_DATE_LENGTH         As Integer = 10     '日付

'Private Const C_MAIL_TXT            As String = "card.txt" 'DELETE 2011/09/28 M.RYU
Private Const C_MAIL_TXT            As String = "mail.txt"  'INSERT 2011/09/28 M.RYU
'Private Const C_MAIL_ADDR           As String = "<mei.ryu@kasegroup.co.jp>"
    
Enum E_PAYMENT      'INSERT 2011/10/03
    p_CARD = 2      'INSERT 2011/10/03
    p_KOUZA = 3     'INSERT 2011/10/03
End Enum

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
'        MODULE_NAME      :chk_CardObj_Click
'        機能             :チェックボックス
'        CREAT            :2011/08/30 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_CardObj_Click()
    Call setObjEnabled(Me.chk_CardObj)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :setObjEnabled
'        機能             :項目使用可能の設定
'        引数             :falg ⇒ チェックボックスの選択状態（初回カード支払対象抽出）
'        CREAT            :2011/08/30 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setObjEnabled(flag As Double)
    If flag = True Then
        Me.cmb_受付種別 = ""
        Me.cmb_毎月回収区分 = ""
        Me.cmb_YukbnFrom = ""
        Me.cmb_Ucode = ""
        Me.cmb_HosyId = ""
        Me.cmb_HosyC = ""

        Call cmb_受付種別_AfterUpdate
        Call cmb_毎月回収区分_AfterUpdate
        Call cmb_YukbnFrom_AfterUpdate
        Call cmb_Ucode_AfterUpdate
        Call cmb_HosyId_AfterUpdate
        Call cmb_HosyC_AfterUpdate

        Me.txt_KeySendDateFrom = ""
        Me.txt_KeySendDateTo = ""
        Me.txt_UkdateFrom = ""
        Me.txt_UkdateTo = ""

        Me.cmb_集金区分 = "04"      'クレジットカード
        Call cmb_集金区分_AfterUpdate

    Else
        Me.txt_CardReqFrom = ""

    End If

    Me.cmb_受付種別.Enabled = Not flag
    Me.cmb_毎月回収区分.Enabled = Not flag
    Me.cmb_集金区分.Enabled = Not flag
    Me.cmb_YukbnFrom.Enabled = Not flag
    Me.cmb_Ucode.Enabled = Not flag
    Me.cmb_HosyId.Enabled = Not flag
    Me.cmb_HosyC.Enabled = Not flag
    Me.txt_KeySendDateFrom.Enabled = Not flag
    Me.txt_KeySendDateTo.Enabled = Not flag
    Me.txt_UkdateFrom.Enabled = Not flag
    Me.txt_UkdateTo.Enabled = Not flag
    Me.txt_CardReqFrom.Enabled = flag
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_HosyC_AfterUpdate
'        機能             :保証会社 AfterUpdate
'        CREAT            :2011/07/02 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_HosyC_AfterUpdate()
    txt_HosyN = AfterUpdateCommon(cmb_HosyC, P_CODE_LENGTH, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_HosyId_AfterUpdate
'        機能             :保証区分 AfterUpdate
'        CREAT            :2011/07/02 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_HosyId_AfterUpdate()
    txt_HosyIdn = AfterUpdateCommon(cmb_HosyId, 2, 1)
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
'        MODULE_NAME      :cmd_Kouza_Click
'        機能             :口座登録のメール送信
'        CREAT            :2011/10/06 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kouza_Click()
    If MsgBox("口座振替登録のメールを送信します。宜しいですか？", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If

    'チェックが付いてるデータに対しメール送信を行う。
    If sendMailCard(E_PAYMENT.p_KOUZA) = False Then
        MsgBox ("メール送信失敗。")
        Exit Sub
    Else
        MsgBox ("メール送信完了。")
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Mail
'        機能             :クレジット登録のメール送信
'        CREAT            :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Mail_Click()
    
    ' 確認メッセージ表示
'    If MsgBox("対象者に対してメールを送信します。宜しいですか？", vbQuestion + vbYesNo) = vbNo Then                        'DELETE 2011/09/28 M.RYU
    If MsgBox("クレジットカード登録のメールを送信します。宜しいですか？", vbQuestion + vbYesNo) = vbNo Then    'INSERT 2011/09/28 M.RYU
        Exit Sub
    End If

    'チェックが付いてるデータに対しメール送信を行う。
    If sendMailCard(E_PAYMENT.p_CARD) = False Then
        MsgBox ("メール送信失敗。")         'INSERT 2011/09/28 M.RYU
        Exit Sub
    Else
        MsgBox ("メール送信完了。")         'INSERT 2011/09/28 M.RYU
    End If
    
'    MsgBox ("メールの送信が完了しました。")'DELETE 2011/09/28 M.RYU

'    '再読み込みさせる      '--DELETE 2011/07/21 M.RYU--<S>
'    If Not Form_FVS300.MakeWork(False) Then
'        Exit Sub
'    End If
'
'    Form_FVS300.SubForm.Requery    '-------------------<E>

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
'        MODULE_NAME      :cmb_HosyId_AfterUpdate
'        機能             :保証区分 AfterUpdate
'        CREAT            :2011/07/02 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_CardReqFrom_KeyPress(KeyAscii As Integer)
    '###############    キー入力の規制（6桁の数字のみ）      ####################
    With Me.txt_CardReqFrom
        Select Case KeyAscii
            Case 8, 9, 10, 13   '8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9")
                If Len(.Text) > 5 And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkdateFrom_AfterUpdate
'        機能             :予約 AfterUpdate
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
'        機能             :予約 BeforeUpdate
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
'        機能             :予約 GotFocus
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
'        機能             :予約 KeyPress
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
'        機能             :予約 LostFocus
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
'        機能             :予約 AfterUpdate
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
'        機能             :予約 BeforeUpdate
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
'        機能             :予約 GotFocus
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
'        機能             :予約 KeyPress
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
'        機能             :予約 LostFocus
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
'        MODULE_NAME      :txt_KeySendDateFrom_AfterUpdate
'        機能             :鍵送付日 AfterUpdate
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
'        機能             :鍵送付日 BeforeUpdate
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
'        機能             :鍵送付日 GotFocus
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
'        機能             :鍵送付日 KeyPress
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
'        機能             :鍵送付日 LostFocus
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
'        MODULE_NAME      :txt_KeySendDateTo_BeforeUpdate
'        機能             :鍵送付日 BeforeUpdate
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
'        MODULE_NAME      :txt_KeySendDateTo_AfterUpdate
'        機能             :鍵送付日 AfterUpdate
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
'        MODULE_NAME      :txt_KeySendDateTo_GotFocus
'        機能             :鍵送付日 GotFocus
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
'        機能             :鍵送付日 KeyPress
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
'        機能             :鍵送付日 LostFocus
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
'        MODULE_NAME      :sendMailCard
'        機能             :クレジットカードURLメール送信
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function sendMailCard() As Boolean                             'DELETE 2011/10/03 M.RYU
Private Function sendMailCard(intPayment As Integer) As Boolean         'INSERT 2011/10/03 M.RYU

On Error GoTo ErrorHandle

    Dim dbAccess        As Database
    Dim dbRecordset     As Recordset
    Dim objCon          As Object
    Dim sql             As String
    Dim strBUMOC        As String
    Dim lngCnt          As Long
    
'    sendMailCard = False
    
    '部門コード取得
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBUMOC = "" Then
        Call MSZZ024_M10("DLookup(""CONT_BUMOC"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", "テーブル[CONT_MAST]の設定不正。")
    End If

    'SQLサーバー接続(kasedb)
    Set objCon = ADODB_Connection
    Set dbAccess = CurrentDb
    
    ' メールアドレスが入っていないデータは対象外とする(@で判断)。
    sql = "      SELECT   * "
    sql = sql & "FROM     FVS300_W01 "
    sql = sql & "WHERE    SEND = -1 "
    sql = sql & "AND      MAIL LIKE '*@*' "
    sql = sql & "ORDER BY UCODE "
     
    Set dbRecordset = dbAccess.OpenRecordset(sql)
    
    If dbRecordset.RecordCount = 0 Then
        MsgBox "対象者が選択されていません。。", vbInformation, P_KAKUNIN
        sendMailCard = False
        Exit Function
    End If
    
    
    '' メール送信でエラーが発生したらその時点で処理をとめる。
    '' 以降のメール送信もエラーになるため。
    Do Until dbRecordset.EOF
        
        ' メール送信
        If sendMail(dbRecordset, intPayment) = False Then
            sendMailCard = False    'INSERT 2011/09/28 M.RYU
            GoTo Exit_rtn
        End If
    
        ' 顧客コメントを登録する
        lngCnt = ADODB_Execute(makeComment(dbRecordset.Fields("UCODE"), strBUMOC))
         
         dbRecordset.MoveNext
     Loop
     
    sendMailCard = True             'INSERT 2011/09/28 M.RYU
    
GoTo Exit_rtn

ErrorHandle:
    Call MSZZ024_M00("sendMailCard", True)
    sendMailCard = False
    Err.Clear
    
Exit_rtn:
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
'    sendMailCard = True            'DELETE 2011/09/28 M.RYU
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :sendMail
'        機能             :メール送信
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function sendMail(dbRecordset As Object, intPayment As Integer) As Boolean

On Error GoTo ErrorHandle
    
    Dim strMail     As String
    Dim strMailPath As String
    Dim iFile       As Integer
    Dim strBuff     As String
    Dim strResult   As String
    Dim strSubject  As String
    Dim strApplyUrl As String
    Dim strSendTo   As String
    Dim objSC       As Object
    Dim strUserName As String
    Dim strUserKana As String           'INSERT 20120105 M.RYU
    Dim urlPath     As String
    Dim strBCC      As String
    Dim strTxt      As String           'INSERT 2011/09/28 M.RYU
          
    sendMail = False
    
    Set objSC = CreateObject("ScriptControl")
    objSC.Language = "Jscript"
    strUserName = objSC.CodeObject.encodeURIComponent(dbRecordset!USER)
    strUserKana = objSC.CodeObject.encodeURIComponent(dbRecordset!USERKANA)         'INSERT 20120105 M.RYU
    Set objSC = Nothing

    'メールテンプレート読込み
    strMailPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='FVS300' and INTIF_RECFB='MAIL_PATH'"))
    
    iFile = FreeFile()
    
    If intPayment = E_PAYMENT.p_CARD Then                           'INSERT 2011/09/28 M.RYU--<S>
        strTxt = "クレジットカード"
        ' 登録URL取得
        urlPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='FVS300' and INTIF_RECFB='entry.php'"))
        ' URL
        strApplyUrl = urlPath & "entry.php" & "?memberid=" & dbRecordset!BUMOC & dbRecordset!UCODE & "&membername=" & strUserName
        
    ElseIf intPayment = E_PAYMENT.p_KOUZA Then
        strTxt = "口座振替"
        ' 登録URL取得
        urlPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='FVS300' and INTIF_RECFB='transfer.php'"))
        ' URL
        strApplyUrl = urlPath & "transfer.php" & "?mi=" & dbRecordset!BUMOC & dbRecordset!UCODE & "&mn=" & strUserKana
        
    End If                                                      'INSERT 2011/09/28 M.RYU--<E>
    
    Open strMailPath & C_MAIL_TXT For Input As #iFile
    
    While Not EOF(iFile)
        Line Input #iFile, strBuff
        strResult = strResult & strBuff & vbCrLf
    Wend
    Close #iFile

'    strSubject = "月々のクレジットカード支払いお申込み案内"    'DELETE  2011/09/28 M.RYU
    strSubject = "【加瀬倉庫】月々の" & strTxt & "支払いお申込み案内"       'INSERT  2011/09/28 M.RYU
    
    ' 内容の置換え
    ' 宛名(BCC - 予約案内)
    strBCC = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='FVS300' and INTIF_RECFB='BCC'"))
    strSendTo = "<" & dbRecordset!MAIL & ">" & vbTab & "bcc" & vbTab & strBCC

    ' メール本文作成
    ' 顧客名
    strResult = Replace(strResult, "%YOUKT_NAME%", dbRecordset!USER)
    ' ヤード住所
    'strResult = Replace(strResult, "%YARD_ADDR%", dbRecordset!ADDR)
    ' ヤード名
    'strResult = Replace(strResult, "%YARD_NAME%", dbRecordset!NAME)
    ' 使用用途
    'strResult = Replace(strResult, "%USAGE_NAME%", dbRecordset!USAGE)
    ' コンテナサイズ
    'strResult = Replace(strResult, "%CNTA_SIZE%", dbRecordset!SIZE)
    ' 段
    'If dbRecordset!BUMOC = "8" Then
    '    strResult = Replace(strResult, "段：%STEP_NAME%" & vbCrLf, "")
    'Else
    '    strResult = Replace(strResult, "%STEP_NAME%", dbRecordset!STEP)
    'End If
    ' 月額料金
    'strResult = Replace(strResult, "%RCPT_RENTKG%", dbRecordset!RENTKG)
    
    ' 支払方法
    strResult = Replace(strResult, "%N_NAME%", strTxt)          'INSERT  2011/09/28 M.RYU
    
'    ' 登録URL取得
'    urlPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='FVS300' and INTIF_RECFB='entry.php'"))  'DELETE  2011/09/28 M.RYU
    ' URL
'    strApplyUrl = urlPath & "entry.php" & "?memberid=" & dbRecordset!BUMOC & dbRecordset!UCODE & "&membername=" & strUserName  'DELETE  2011/09/28 M.RYU
    strResult = Replace(strResult, "%APPLYURL%", strApplyUrl)
    
    ' メール送信
    Call MSZZ021_M10(strBCC, strSendTo, strSubject, strResult)

    sendMail = True
    
    Exit Function

ErrorHandle:
    Call MSZZ024_M00("sendMailCard", True)
    sendMail = False
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :makeComment
'        機能             :顧客コメントに送信日を登録する。
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function makeComment(UserCode As String, strBUMOC As String) As String
    
    Dim strSQL       As String

    strSQL = strSQL & "INSERT INTO "
    strSQL = strSQL & " KOTO_TRAN  "
    strSQL = strSQL & "( "
    strSQL = strSQL & " KOTOT_BUMOC, "
    strSQL = strSQL & " KOTOT_KOKYC, "
    strSQL = strSQL & " KOTOT_TANTC, "
    strSQL = strSQL & " KOTOT_STANC, "
    strSQL = strSQL & " KOTOT_GENTD, "
    strSQL = strSQL & " KOTOT_GENTJ, "
    strSQL = strSQL & " KOTOT_TYPEC, "
    strSQL = strSQL & " KOTOT_CONTN, "
    strSQL = strSQL & " KOTOT_CNT1N, "
    strSQL = strSQL & " KOTOT_CNT2N, "
    strSQL = strSQL & " KOTOT_DISPI, "
    strSQL = strSQL & " KOTOT_OTKB1, "
    strSQL = strSQL & " KOTOT_OTKB2, "
    strSQL = strSQL & " KOTOT_OTKB3, "
    strSQL = strSQL & " KOTOT_OTKB4, "
    strSQL = strSQL & " KOTOT_OTKB5, "
    strSQL = strSQL & " KOTOT_INSED, "
    strSQL = strSQL & " KOTOT_INSEJ, "
    strSQL = strSQL & " KOTOT_INSPB, "
    strSQL = strSQL & " KOTOT_INSUB  "
    strSQL = strSQL & " ) VALUES ( "
    strSQL = strSQL & "'" & strBUMOC & "', "
    strSQL = strSQL & "'" & UserCode & "', "
    strSQL = strSQL & " '999', "
    strSQL = strSQL & " '999', "
    strSQL = strSQL & Format$(DATE, "YYYYMMDD") & ", "
    strSQL = strSQL & " 000000, "
    strSQL = strSQL & " '09', "
'    strSQL = strSQL & " '申込み用紙発行', "    'DELETE 2011/11/11 M.RYU
    strSQL = strSQL & " '申込メール送信', "     'INSERT 2011/11/11 M.RYU
    strSQL = strSQL & " '', "
    strSQL = strSQL & " '', "
    strSQL = strSQL & " 1, "
    strSQL = strSQL & " '99', "
    strSQL = strSQL & " '99', "
    strSQL = strSQL & " '99', "
    strSQL = strSQL & " '99', "
    strSQL = strSQL & " '99', "
    strSQL = strSQL & Format$(DATE, "YYYYMMDD") & ", "
    strSQL = strSQL & " '000000', "
    strSQL = strSQL & "'" & PROG_ID & "', "
    strSQL = strSQL & "'" & LsGetUserName() & "' "
    strSQL = strSQL & " )"

    makeComment = strSQL
    
End Function


'==============================================================================*
'
'        MODULE_NAME      :cmb_集金区分_AfterUpdate
'        機能             :cmb_集金区分_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_AfterUpdate()
    Me.txt_集金区分 = AfterUpdateCommon(cmb_集金区分, P_DIVISION_LENGTH, 1)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_集金区分_BeforeUpdate
'        機能             :cmb_集金区分_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_BeforeUpdate(Cancel As Integer)
    Call ChkComboInput(cmb_集金区分, P_DIVISION_LENGTH, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_集金区分_KeyPress
'        機能             :cmb_集金区分_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_集金区分, KeyAscii, P_DIVISION_LENGTH)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_毎月回収区分_AfterUpdate
'        機能             :cmb_毎月回収区分_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_毎月回収区分_AfterUpdate()
    Me.txt_毎月回収区分 = AfterUpdateCommon(cmb_毎月回収区分, P_DIVISION_LENGTH, 1)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_毎月回収区分_BeforeUpdate
'        機能             :cmb_毎月回収区分_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_毎月回収区分_BeforeUpdate(Cancel As Integer)
    Call ChkComboInput(cmb_毎月回収区分, P_DIVISION_LENGTH, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_毎月回収区分_KeyPress
'        機能             :cmb_毎月回収区分_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_毎月回収区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_毎月回収区分, KeyAscii, P_DIVISION_LENGTH)
End Sub


