VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS720_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　解約一覧サブ
'   プログラムＩＤ　：　FVS720_S01
'   作　成　日　　　：  2005/07/29
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：2005/10/02
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：入力不可項目に対する必須チェックを行わない
'   Ver             ：  0.1
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
Public blnLoading   As Boolean

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const DETAIL_FORM As String = "FVS700"
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

'桁数
Private Const P_DIVISION_LENGTH     As Integer = 2      '区分
Private Const P_PERSON_LENGTH       As Integer = 3      '担当コード
Private Const P_CODE_LENGTH         As Integer = 6      'ヤードコード／コンテナ番号
Private Const P_DATE_LENGTH         As Integer = 10      '日付
Private Const P_YM_LENGTH           As Integer = 7       '年月
Private Const P_ZERO_SUPPRESS       As String = "0000000000"    '不足桁０埋め用

'契約状態
Private Const P_STATUS_UNSETTING    As String = "00"    '契約状態「未設定」
Private Const P_STATUS_CONTRACT     As String = "01"    '契約状態「契約」
Private Const P_STATUS_CANCEL       As String = "02"    '契約状態「解約予定」
Private Const P_STATUS_EXTEND       As String = "03"    '契約状態「解約延長」
Private Const P_STATUS_MAINTENANCE  As String = "04"    '契約状態「メンテ中」
Private Const P_STATUS_COMPLETE     As String = "09"    '契約状態「完了」

'背景色
Private Const P_BACKCOLOR_YELLOW    As Long = 14680063      '黄
Private Const P_BACKCOLOR_WHITE     As Long = 16777215      '白

'表示データよりも新しい契約データが存在する場合Trueになる
Private blnNewContract  As Boolean

'詳細画面表示中フラグ
Private blnDispDetail   As Boolean

Public blnStatusNotSpecifyOrComplete  As Boolean

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

    txt_ContractState = AfterUpdateCommon(cmb_ContractState, P_DIVISION_LENGTH)
    If Not blnStatusNotSpecifyOrComplete Then
        Call SelectStatus
    End If
    
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
            MsgBox "該当する状態区分がありません。", vbInformation, P_ERROR
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
'        MODULE_NAME      :cmb_MaintenancePerson_AfterUpdate
'        機能             :メンテ担当 AfterUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MaintenancePerson_AfterUpdate()
    
    txt_MaintenancePerson = AfterUpdateCommon(cmb_MaintenancePerson, P_PERSON_LENGTH)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MaintenancePerson_BeforeUpdate
'        機能             :メンテ担当 BeforeUpdate ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MaintenancePerson_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Dim intRet As Integer
    intRet = ChkComboInput(cmb_MaintenancePerson, P_PERSON_LENGTH)
    If intRet <> 0 Then
        If intRet = 2 Then
            MsgBox "該当する担当者マスタがありません。", vbInformation, P_ERROR
            SendKeys "{ESC}"
        End If
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MaintenancePerson_KeyPress
'        機能             :メンテ担当 KeyPress ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MaintenancePerson_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_MaintenancePerson, KeyAscii, P_PERSON_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Detail_Click
'        機能             :詳細 Click ｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Detail_Click()

    Dim blnRet      As Boolean
    Dim strContNo   As String
    
    strContNo = txt_ContractNo
    If blnDispDetail Then
        blnRet = True
    Else
        If Not RequireCheck Then
            Exit Sub
        End If
        
        If Me.Dirty Then    'カレントレコードが変更されていたら保存する
            Me.Dirty = False
        End If
    
        
        If Parent.blnChangeMode Then
            blnRet = Parent.WorkToCarg
            If blnRet Then
                If Not Parent.DataCopy("FVS720_W01", "FVS720_W02", True) Then
                    Exit Sub
                End If
                Me.Requery
                Call Parent.EndChange
            End If
        Else
            blnRet = True
        End If
    End If
    
    If blnRet Then
        cmd_Detail.SetFocus
        Call FormEnable(False)
        blnDispDetail = True
        doCmd.OpenForm DETAIL_FORM, , , , , , strContNo & "," & Parent.NAME
    End If
    
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
    If cmbCheck.ListIndex = -1 Then
        intRet = 2
        GoTo ChkComboInput_End
    End If

ChkComboInput_End:
    ChkComboInput = intRet

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
'        MODULE_NAME      :AfterUpdateCommon
'        機能             :更新後イベント時の処理
'        IN               :①対象コントロール
'                         :②入力最大桁数
'                         :③省略可　日付項目の場合True（規定値False)
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :コンボボックスは名称
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function AfterUpdateCommon(objTarget As Object, intLength As Integer, Optional blnDate As Boolean = False) As String

    AfterUpdateCommon = ""
    
    '体裁を整える
    If blnDate Then
        AfterUpdateCommon = Format(VarChangeDate, "yyyy/mm/dd")
    Else
        objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    End If
    
    txt_ChangeFlag = "1"
    
    Call Parent.InputChange
    
    If Left(objTarget.NAME, 3) = "cmb" Then
        AfterUpdateCommon = Nz(objTarget.Column(1, objTarget.ListIndex))
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_BeforeUpdate
'        機能             :Form_BeforeUpdateイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_BeforeUpdate(Cancel As Integer)

    If Not RequireCheck Then
            Cancel = True
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :RequireCheck
'        機能             :必須入力チェック
'        IN               :
'        OUT              :True=必須入力が全て入力されている False=されていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function RequireCheck() As Boolean

    RequireCheck = False
    
    If Not RequireControl(cmb_ContractState, "契約区分") Then Exit Function
    
    '状態区分「契約」または、「未設定」または、新契約なしの「完了」（完了※）の場合は必須項目無し
    If cmb_ContractState = P_STATUS_CONTRACT Or _
       cmb_ContractState = P_STATUS_UNSETTING Or _
       (cmb_ContractState = P_STATUS_COMPLETE And blnNewContract) _
    Then
        RequireCheck = True
        Exit Function
    End If
    
    'メンテ中はメンテ依頼日は必須
    If cmb_ContractState = P_STATUS_MAINTENANCE Then
        If Not RequireControl(txt_MaintenanceRequestDate, "メンテ依頼日") Then Exit Function
    End If
    
    '完了時は、読み込み時にメンテナンス依頼日が格納されている場合、
    'メンテ依頼日、メンテ完了日、メンテ担当は必須
    If cmb_ContractState = P_STATUS_COMPLETE And Nz(txt_HideMaintenanceRequestDate) <> "" And txt_NewContractFlag = "0" Then
        If Not RequireControl(txt_MaintenanceRequestDate, "メンテ依頼日") Then Exit Function
        
        If Not RequireControl(txt_MaintenanceEndDate, "メンテ完了日") Then Exit Function
        
        If Not RequireControl(cmb_MaintenancePerson, "メンテ担当") Then Exit Function
    End If
    
    RequireCheck = True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :RequireControl
'        機能             :必須チェック
'        IN               :①対象コントロール　②エラーメッセージに出す文言
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function RequireControl(objTarget As Object, strMsgText As String) As Boolean

    RequireControl = True
    
    '20051002 SHIBAZAKI Enabledプロパティをチェックする
    If objTarget.Enabled And Nz(objTarget.VALUE) = "" Then
        MsgBox strMsgText & "は必須入力です。", vbInformation, P_KAKUNIN
        objTarget.SetFocus
        RequireControl = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_Current
'        機能             :Form_Currentイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Current()

    If Not Parent.blnLoading And Not blnDispDetail And Not blnStatusNotSpecifyOrComplete Then
        Call SelectStatus
    End If
    
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

    Call Parent.Form_KeyDown(KeyCode, Shift)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form_Loadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    'サブフォーム／親フォーム共にLOADが完了するまでカレントレコード移動処理はスキップ
    Parent.blnLoading = True
    blnDispDetail = False
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_AfterUpdate
'        機能             :精算依頼日 AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_AfterUpdate()

    txt_AccountRequestDate = AfterUpdateCommon(txt_AccountRequestDate, P_DATE_LENGTH, True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_BeforeUpdate
'        機能             :精算依頼日 BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_AccountRequestDate, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_AccountRequestDate_KeyPress
'        機能             :精算依頼日 KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_AccountRequestDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_AccountRequestDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_AfterUpdate
'        機能             :メンテ完了日 AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_AfterUpdate()

    txt_MaintenanceEndDate = AfterUpdateCommon(txt_MaintenanceEndDate, P_DATE_LENGTH, True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_BeforeUpdate
'        機能             :メンテ完了日 BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_MaintenanceEndDate, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceEndDate_KeyPress
'        機能             :メンテ完了日 KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceEndDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_MaintenanceEndDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_AfterUpdate
'        機能             :メンテ依頼日 AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_AfterUpdate()

    txt_MaintenanceRequestDate = AfterUpdateCommon(txt_MaintenanceRequestDate, P_DATE_LENGTH, True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_BeforeUpdate
'        機能             :メンテ依頼日 BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    Call dateCheck(txt_MaintenanceRequestDate, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_MaintenanceRequestDate_KeyPress
'        機能             :メンテ依頼日 KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_MaintenanceRequestDate_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_MaintenanceRequestDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectStatus
'        機能             :ステータス変更時の画面制御
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
 Sub SelectStatus()

    Select Case cmb_ContractState
    Case P_STATUS_CONTRACT
        Call SelectContract
    Case P_STATUS_CANCEL
        Call SelectCacnel
    Case P_STATUS_EXTEND
        Call SelectExtend
    Case P_STATUS_MAINTENANCE
        Call SelectMaintenance
    Case P_STATUS_COMPLETE
        If txt_NewContractFlag = "1" Then
            Call SelectCompleteNewContract
        Else
            Call SelectComplete
        End If
    End Select

End Sub


'==============================================================================*
'
'        MODULE_NAME      :SelectContract
'        機能             :ステータス=「契約」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectContract()

    txt_CancelDate = ""
    txt_MaintenanceRequestDate = ""
    txt_MaintenanceEndDate = ""
    cmb_MaintenancePerson = ""
    txt_MaintenancePerson = ""
    
    Call SetCtlProperty(txt_MaintenanceRequestDate, False)
    Call SetCtlProperty(txt_MaintenanceEndDate, False)
    Call SetCtlProperty(cmb_MaintenancePerson, False)
    Call SetCtlProperty(txt_AccountRequestDate, False)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :SelectCacnel
'        機能             :ステータス=「解約予定」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectCacnel()

    txt_MaintenanceRequestDate = ""
    txt_MaintenanceEndDate = ""
    cmb_MaintenancePerson = ""
    txt_MaintenancePerson = ""
    
    Call SetCtlProperty(txt_MaintenanceRequestDate, False)
    Call SetCtlProperty(txt_MaintenanceEndDate, False)
    Call SetCtlProperty(cmb_MaintenancePerson, False)
    Call SetCtlProperty(txt_AccountRequestDate, True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectExtend
'        機能             :ステータス=「解約延長」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectExtend()
    
    Call SelectCacnel

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectMaintenance
'        機能             :ステータス=「メンテ中」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectMaintenance()


    txt_MaintenanceEndDate = ""
    
    Call SetCtlProperty(txt_MaintenanceRequestDate, True)
    Call SetCtlProperty(txt_MaintenanceEndDate, False)
    Call SetCtlProperty(cmb_MaintenancePerson, True)
    Call SetCtlProperty(txt_AccountRequestDate, True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectComplete
'        機能             :ステータス=「完了」
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectComplete()

    Call SetCtlProperty(txt_MaintenanceRequestDate, True)
    Call SetCtlProperty(txt_MaintenanceEndDate, True)
    Call SetCtlProperty(cmb_MaintenancePerson, True)
    Call SetCtlProperty(txt_AccountRequestDate, True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SelectCompleteNewContract
'        機能             :ステータス=「完了」＆表示データより新しい契約がある
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SelectCompleteNewContract()

    Call SetCtlProperty(txt_MaintenanceRequestDate, False)
    Call SetCtlProperty(txt_MaintenanceEndDate, False)
    Call SetCtlProperty(cmb_MaintenancePerson, False)
    Call SetCtlProperty(txt_AccountRequestDate, True)
    
End Sub


'==============================================================================*
'
'        MODULE_NAME      :SetCtlProperty
'        機能             :ﾃｷｽﾄﾌﾟﾛﾊﾟﾃｨ設定
'        IN               :txtCtl = text / blnEnabled = Enabled設定
'        INSERT DATE      :2005/05/27
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetCtlProperty(txtCtl As Object, blnEnabled As Boolean)

    With txtCtl
        .Enabled = blnEnabled
        .Locked = Not blnEnabled
    End With

End Sub

'==============================================================================*
'
'        MODULE_NAME      :項目使用可不可設定
'        MODULE_ID        :FormEnable
'        CREATE_DATE      :2005/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subFormReload()

    Call FormEnable(True)
    
    If blnStatusNotSpecifyOrComplete Then
        Call OnlyAccountDate
    Else
        Call SelectStatus
    End If
    blnDispDetail = False
    
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

    cmb_ContractState.Enabled = blnEnable
    txt_MaintenanceRequestDate.Enabled = blnEnable
    txt_MaintenanceEndDate.Enabled = blnEnable
    cmb_MaintenancePerson.Enabled = blnEnable
    txt_AccountRequestDate.Enabled = blnEnable
    
    Parent.cmd_保存.Enabled = blnEnable
    Parent.cmd_取消.Enabled = blnEnable
    Parent.cmd_終了.Enabled = blnEnable

End Sub

'==============================================================================*
'
'        MODULE_NAME      :OnlyAccountDate
'        機能             :精算依頼日のみ使用可にする
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub OnlyAccountDate()

    cmb_ContractState.Enabled = False
    Call YellowControl(txt_MaintenanceRequestDate, True)
    Call YellowControl(txt_MaintenanceEndDate, True)
    cmb_MaintenancePerson.Enabled = False
    Call YellowControl(txt_AccountRequestDate, False)
  
End Sub


'==============================================================================*
'
'        MODULE_NAME      :YellowControl
'        機能             :コントロール黄色化
'        IN               :第一引数：対象コントロール
'                         :第二引数：省略可 True=黄色化（規定値） False=白化
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub YellowControl(objTarget As Object, Optional blnYellow As Boolean = True)

    If blnYellow Then
        '背景色を黄色
        objTarget.BackColor = P_BACKCOLOR_YELLOW
    Else
        '背景色を黄色
        objTarget.BackColor = P_BACKCOLOR_WHITE
    End If
    
    '編集ロック
    objTarget.Enabled = Not blnYellow
    objTarget.Locked = blnYellow

End Sub

