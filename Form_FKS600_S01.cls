VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS600_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　セキュリティカードメンテ（サブ）
'   プログラムＩＤ　：　FKS600_S01
'   作　成　日　　　：  2006/05/01
'   作　成　者　　　：  tajima
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/06/20
'   修　正　者　　　：  tajima
'   修　正　内　容　：  削除ボタンをタブストップしないようにした
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：  2012/07/07
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  部屋No.を入力した際に、用途を基本へ変更する。
'   Ver             ：  0.2
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_KAKUNIN           As String = "確認"
Private Const P_ERROR             As String = "エラー"
Private Const P_DATE_LENGTH       As Integer = 10      '日付

Private Const P_BACKCOLOR_YELLOW  As Long = 14680063      '黄
Private Const P_BACKCOLOR_WHITE   As Long = 16777215      '白
Private Const P_部屋_無割当       As String = "無し"
Private Const P_チェック済み      As Integer = -1
Private Const P_未チェック        As Integer = 0

'==============================================================================*
'
'        MODULE_NAME      :フォームの基本設定
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2006/06/04
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    

End Sub
'==============================================================================*
'
'       MODULE_NAME      :フォームキー押下時処理
'       MODULE_ID        :Form_KeyDown
'       NOTE             :サブフォーム内の状況に応じコントロールプロパティを変える
'       CREATE_DATE      :2006/06/04
'       UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Me.Parent.KeyDown(KeyCode, Shift)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :サブフォームに移動してきたとき
'        MODULE_ID        :Form_Current
'        NOTE             :サブフォーム内の状況に応じコントロールプロパティを変える
'        CREATE_DATE      :2006/06/04
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Current()
 '使用用途によるコントロール変更
  If Me.Parent.pFROM_IDforFKS600_S01 <> "" Then Call setControlProperty
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_BeforeUpdate(Cancel As Integer)

    Dim errorMsg As String
    
    errorMsg = ""
    
    ' 入力項目・必須チェック
    If Nz(Me.txt_カード№, "") = "" Then
        errorMsg = "カード№は必須入力です。"
        Me.txt_カード№.SetFocus
    
    ElseIf Nz(Me.lst_部屋№.VALUE, "") = "" Then
        errorMsg = "部屋№は必須入力です。"
        Me.lst_部屋№.SetFocus
    
    ElseIf Nz(Me.lst_用途.VALUE, "") = "" Then
        errorMsg = "用途は必須入力です。"
        Me.lst_用途.SetFocus
    
    ElseIf Nz(Me.lst_用途.VALUE, "") = P_カード用途_業者貸値 And _
           Nz(Me.txt_貸出先.VALUE, "") = "" Then
        errorMsg = "業者貸しならば貸出先は必須入力です。"
        Me.txt_貸出先.SetFocus
    
    ElseIf Nz(Me.lst_用途.VALUE, "") = P_カード用途_業者貸値 And _
           Nz(Me.txt_貸出日.VALUE, "") = "" Then
        errorMsg = "業者貸しならば貸出日は必須入力です。"
        If Nz(Me.chk_LendingDay.VALUE, 0) = P_チェック済み Then
            Me.txt_貸出日.SetFocus
        Else
            Me.chk_LendingDay.SetFocus
        End If
    
    ElseIf Me.Parent.IsSelectModeFKS600() = False And _
           Nz(Me.lst_用途.VALUE, "") = P_カード用途_基本値 And _
           Nz(Me.lst_部屋№.VALUE, "") = P_部屋_無割当 Then
        errorMsg = "用途が基本ならば部屋№は必須指定です。"
        Me.lst_部屋№.SetFocus
    
    ' 入力項目・矛盾チェック
    ElseIf Nz(Me.chk_LendingDay.VALUE, 0) <> P_未チェック And _
           Nz(Me.txt_貸出日.VALUE, "") = "" Then
        errorMsg = "貸出チェック済みならば貸出日は必須入力です。" & vbCr & _
                   "※貸出日を入れないのならば貸出チェックを外して下さい。"
        Me.chk_LendingDay.SetFocus
    
    ElseIf Nz(Me.chk_ReturnDay.VALUE, 0) <> P_未チェック And _
           Nz(Me.txt_返却日.VALUE, "") = "" Then
        errorMsg = "返却チェック済みならば返却日は必須入力です。" & vbCr & _
                   "※返却日を入れないのならば返却チェックを外して下さい。"
        Me.chk_ReturnDay.SetFocus
    
    ElseIf Nz(Me.chk_StopDay.VALUE, 0) <> P_未チェック And _
           Nz(Me.txt_停止日.VALUE, "") = "" Then
        errorMsg = "停止チェック済みならば停止日は必須入力です。" & vbCr & _
                   "※停止日を入れないのならば停止チェックを外して下さい。"
        Me.chk_StopDay.SetFocus
    
    ElseIf Nz(Me.lst_用途.VALUE, "") = P_カード用途_追加値 And _
           Nz(Me.chk_LendingDay.VALUE, 0) = P_未チェック Then
        errorMsg = "用途が追加ならば貸出日は必須入力です。" & vbCr & _
                   "※貸出チェックをして下さい。"
        Me.chk_LendingDay.SetFocus
    End If
        
    If errorMsg <> "" Then
        MsgBox errorMsg, vbExclamation, P_ERROR
        Cancel = True
        Exit Sub
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_AfterUpdate()
    ' ★ここで画面項目（ワークテーブル）に変更をかけると無限ループっぽくなるよ
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_LendingDay_BeforeUpdate(Cancel As Integer)
    
    ' 貸出日チェックの有効可否チェック
    
    ' 仕様として
    ' 貸出中としているときの変更(つまり貸出取消)には警告を出す
    If Is顧客貸出() = True Then
        If vbNo = MsgBox("貸出取消をしますか？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, P_KAKUNIN) Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :貸出チェックをクリック
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/06/05
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_LendingDay_Click()
    
    If Nz(chk_LendingDay.VALUE, 0) = P_チェック済み Then
        '貸出チェックを付けたら、システム日付を設定
        txt_貸出日.VALUE = Format(DATE, "yyyy/mm/dd")
        ' 受付か契約から来たのならば
        If Me.Parent.IsSelectModeFKS600() = True Then
          ' 部屋№の設定と契約番号、契約客先名の設定を行う
            lst_部屋№ = Me.Parent.pUseRoomNoforFKS600_S01
            txt_貸出先契約 = Me.Parent.pAcptNoforFKS600_S01
            txt_貸出先 = Me.Parent.pUserNameforFKS600_S01
          ' 用途の自動設定
            If lst_用途 <> CmKomsMod.P_カード用途_基本値 Then
                lst_用途.VALUE = CmKomsMod.P_カード用途_追加値  ' 基本以外は全て追加とする
            End If
        Else
        ' メンテ系ならば業者貸しとする
            ' 業者名の設定を促す
            If Nz(txt_貸出先, "") = "" Then txt_貸出先 = "業者名を入力して下さい"
            ' 用途の自動設定
            lst_用途.VALUE = CmKomsMod.P_カード用途_業者貸値    ' 業者貸しとする
        End If
    Else
    '貸出チェックを外したら、貸出日はクリアする
        txt_貸出日.VALUE = Null
      ' 部屋№は元に、契約番号、契約客先名はクリアする
        lst_部屋№ = IIf(Nz(txt_元部屋番号, P_部屋_無割当) = P_部屋_無割当, P_部屋_無割当, txt_元部屋番号)
        txt_貸出先契約 = Null
        txt_貸出先 = Null
      ' 用途は基本以外は全て予備とする
        If Nz(lst_用途, "") <> CmKomsMod.P_カード用途_基本値 Then lst_用途 = CmKomsMod.P_カード用途_予備値
     ' また返却日、停止日もクリアとする※貸出事実を消すんだから
        txt_返却日.VALUE = Null
        chk_ReturnDay.VALUE = P_未チェック
        txt_停止日.VALUE = Null
        chk_StopDay.VALUE = P_未チェック
    End If
        
    Call commonAfterUpdate
    Call setControlProperty   'コントロール入力可否の属性変更
    
    'ガイダンスへの反映（選択系のときのみ動作）ちょっと保留
    'If Me.Parent.IsSelectModeFKS600 = True Then Call Me.Parent.NoticeChangeUseType

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_ReturnDay_BeforeUpdate(Cancel As Integer)

    ' 返却日チェックの有効可否チェック
    
    ' ①貸出チェック済みであること
    If Nz(chk_LendingDay.VALUE, 0) <> P_チェック済み Then
        MsgBox "貸出していないものは返却の必要はありません。", vbExclamation, P_ERROR
        SendKeys "{ESC}"
        Cancel = True
    ' ②停止中チェック済みでないこと
    ElseIf Nz(chk_StopDay.VALUE, 0) = P_チェック済み Then
        MsgBox "停止中なので返却はできません。" & vbCr & _
                "返却するのならば停止を取消して下さい。", vbExclamation, P_ERROR
        SendKeys "{ESC}"
        Cancel = True
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :返却チェックをクリック
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_ReturnDay_Click()

    If Nz(chk_ReturnDay.VALUE, 0) = P_チェック済み Then
    '返却チェックを付けたら、システム日付を設定
        txt_返却日.VALUE = Format(DATE, "yyyy/mm/dd")
    Else
    '返却チェックを外したら、返却日はクリアする
        txt_返却日.VALUE = Null
    End If

    Call commonAfterUpdate
    Call setControlProperty   'コントロール入力可否の属性変更
   
   'ガイダンスへの反映（選択系のときのみ動作） ちょっと保留
   ' If Me.Parent.IsSelectModeFKS600 = True Then Call Me.Parent.NoticeChangeUseType

End Sub
'==============================================================================*
'
'        MODULE_NAME      :停止日チェック
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_StopDay_BeforeUpdate(Cancel As Integer)

    ' 停止日チェックの有効可否チェック
    
    ' 仕様として
    ' 停止日はどのような状態でも（貸出可否に関わらず）チェックできる必要がある
    ' よってとくにチェックはしない
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_StopDay_Click()
    
    If Nz(chk_StopDay.VALUE, 0) = P_チェック済み Then
    '停止チェックを付けたら、システム日付を設定する
        txt_停止日.VALUE = Format(DATE, "yyyy/mm/dd")
    '返却日もクリアする
        txt_返却日.VALUE = Null
        chk_ReturnDay.VALUE = P_未チェック
    Else
    '停止チェックを外したら、停止日はクリアとする
        txt_停止日.VALUE = Null
    End If

    Call commonAfterUpdate
    Call setControlProperty   'コントロール入力可否の属性変更
    
    'ガイダンスへの反映（選択系のときのみ動作）ちょっと保留
    ' If Me.Parent.IsSelectModeFKS600 = True Then Call Me.Parent.NoticeChangeUseType
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lst_部屋№_AfterUpdate()
    Call commonAfterUpdate
    
    '' 2012/07/07 M.HONDA START
    If lst_部屋№ <> "無し" Then
        lst_用途 = "01"
    End If
    '' 2012/07/07 M.HONDA END
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :部屋№が変更できるのはメンテ系且つ顧客未貸出
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lst_部屋№_BeforeUpdate(Cancel As Integer)
    
    '貸出中の場合は変更不可
    If Is顧客貸出() = True Then
        If vbNo = MsgBox("既にお客様に貸出中としていますが宜しいですか？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, P_KAKUNIN) Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If
    
    ' 空白入力と "無し" はここでは許す　※Form_BeforeUpdate で空白不可の最終チェックはします
    If Nz(lst_部屋№.VALUE, "") = "" Or Nz(lst_部屋№.VALUE, "") = P_部屋_無割当 Then Exit Sub
    
    ' 部屋番号のチェック
    Dim sqlText As String
    sqlText = " CNTA_CODE = " & Val(Me.Parent.cmb_Yard.VALUE) & _
              " AND CNTA_NO =" & Val(lst_部屋№.VALUE)
    If "" = Nz(DLookup("CNTA_CODE", "dbo_CNTA_MAST", sqlText), "") Then
        MsgBox "指定された部屋№は登録されていません。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If


End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lst_部屋№_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(lst_部屋№, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lst_用途_AfterUpdate()
    Call commonAfterUpdate
    Call setControlProperty   'コントロールの属性変更
    
    ' 編集系で顧客貸出ではなく業者貸しでもなければ貸出先はクリアする
    If Me.Parent.IsSelectModeFKS600() = False And _
       Nz(Me.txt_元契約番号, "") = "" And Nz(Me.lst_用途.VALUE, "") <> P_カード用途_業者貸値 Then
        Me.txt_貸出先 = Null
    Else
    '用途が変わるのならばガイダンスへの反映（選択系のときのみ動作）
    '    Call Me.Parent.NooticeChangeUseType  保留
    End If
    
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub lst_用途_BeforeUpdate(Cancel As Integer)
    
    Dim errorMessage As String
    Dim alertMessage As String
    
    errorMessage = ""
    alertMessage = ""
    
    ' 既に顧客に貸出にした場合
    If Is顧客貸出() = True Then
        If lst_用途.VALUE = CmKomsMod.P_カード用途_業者貸値 Then
            errorMessage = "既にお客様に貸出中の為「業者貸し」は選択できません。"
        
        ElseIf lst_用途.VALUE = CmKomsMod.P_カード用途_予備値 Then
            errorMessage = "お客様に貸出中の場合は「予備」は選択出来ません。"
        
        ElseIf lst_用途.VALUE <> txt_元使用用途.VALUE Then
            alertMessage = "お客様に貸出中の用途と異なりますが宜しいのですか？"
        End If
    
    Else
    ' 顧客に未貸出のパターンで
    ' ・選択系の場合
        If Me.Parent.IsSelectModeFKS600() = True Then
            If lst_用途.VALUE = CmKomsMod.P_カード用途_業者貸値 Then
                errorMessage = "ここでは「業者貸」は選択できません。"
            ElseIf lst_用途.VALUE = CmKomsMod.P_カード用途_予備値 And _
                   Nz(chk_LendingDay.VALUE, 0) = P_チェック済み Then
                errorMessage = "「予備」は貸出できません。"
            End If
    ' ・メンテ系の場合
        Else
            If lst_用途.VALUE = CmKomsMod.P_カード用途_追加値 Then
                errorMessage = "ここでは「追加」には選択できません。"
            End If
        End If
    End If
    
    'エラーメッセージに値が入ればエラー
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
    
    '警戒メッセージに値が入れば確認
    If alertMessage <> "" Then
        If vbNo = MsgBox(alertMessage, vbYesNo + vbExclamation + vbDefaultButton2, "確認") Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If


End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード№_AfterUpdate()
    Call commonAfterUpdate
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード№_BeforeUpdate(Cancel As Integer)

    Dim strSQL  As String
    Dim vRet    As Variant

    '必須チェック
    If IsNull(txt_カード№.VALUE) Then
        MsgBox "必須入力です。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
    
    '停止中は変更不可
    If Nz(chk_StopDay.VALUE, 0) = P_チェック済み Then
        MsgBox "このカードは停止中としています。" & vbCr & _
                "カード番号を変更するなら停止チェックを外して下さい。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
    
    '顧客に貸出中は変更不可
    If Is顧客貸出() = True Then
        If vbNo = MsgBox("貸出中としていますが宜しいですか？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, P_KAKUNIN) Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If
    
    '全角文字入力チェック
    If CmChkCher(txt_カード№.VALUE, "A") = False Then
        MsgBox "全角文字が入力されています。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
    
    '入力桁数チェック
    If chkTextInput(txt_カード№, 11) <> 0 Then
        Cancel = True
        Exit Sub
    End If

    '存在チェック
    strSQL = strSQL & " SCRDM_CARD_NO = '" & CStr(txt_カード№.VALUE) & "'"
    vRet = Nz(DLookup("SCRDM_CARD_NO", "FKS600_W01", strSQL), "")
    
    If vRet <> "" Then
        MsgBox "カード番号が重複しています。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード№_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_カード№, KeyAscii, 11)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出先_AfterUpdate()
    Call commonAfterUpdate
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出先_BeforeUpdate(Cancel As Integer)

    If Nz(lst_用途.VALUE, "") <> CmKomsMod.P_カード用途_業者貸値 Then
        MsgBox "用途が【業者貸】以外は編集できません。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    Else
          If chkTextInput(txt_貸出先, 36) <> 0 Then
              Cancel = True
          End If
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出先_KeyPress(KeyAscii As Integer)

    ' 用途で業者貸しのときのみ入力可
    If Nz(lst_用途.VALUE, "") <> CmKomsMod.P_カード用途_業者貸値 Then
        KeyAscii = 0
        Exit Sub
    Else
        Call CmKeyPress.CmKeyPress_For_Char2(txt_貸出先, KeyAscii, 36)
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function dateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    Dim VarChangeDate As Variant
    dateCheck = True
    
    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        dateCheck = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出日_AfterUpdate()
    
    Me.txt_貸出日 = Format(Me.txt_貸出日, "yyyy/mm/dd")  '体裁を整える
    Call commonAfterUpdate
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出日_BeforeUpdate(Cancel As Integer)

    '日付入力チェック
    If dateCheck(txt_貸出日, Cancel) = False Then
        Exit Sub
    End If
  
    ' 日付矛盾チェック
    If Nz(chk_ReturnDay.VALUE, 0) = P_チェック済み Then
        If Format(txt_貸出日, "yyyymmdd") > Format(txt_返却日, "yyyymmdd") Then
            MsgBox "貸出日は返却日より過去を入力してください。", vbExclamation, P_ERROR
            SendKeys "{ESC}"
            Cancel = True
        End If
    End If

    If Nz(chk_StopDay.VALUE, 0) = P_チェック済み Then
        If Format(txt_貸出日, "yyyymmdd") > Format(txt_停止日, "yyyymmdd") Then
            MsgBox "貸出日は停止日より過去を入力してください。", vbExclamation, P_ERROR
            SendKeys "{ESC}"
            Cancel = True
        End If
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出日_GotFocus()

    '体裁を整える
      txt_貸出日 = Format(txt_貸出日, "yy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出日_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_貸出日, KeyAscii, P_DATE_LENGTH)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出日_LostFocus()

    '体裁を整える
    txt_貸出日 = Format(Me.txt_貸出日, "yyyy/mm/dd")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日_AfterUpdate()

    txt_停止日 = Format(Me.txt_停止日, "yyyy/mm/dd")  '体裁を整える
    Call commonAfterUpdate
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    If dateCheck(txt_停止日, Cancel) = False Then
        Exit Sub
    End If
    
    If Format(txt_貸出日, "yyyymmdd") > Format(txt_停止日, "yyyymmdd") Then
        MsgBox "停止日は貸出日より未来を入力してください。", vbExclamation, P_ERROR
        Cancel = True
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日_GotFocus()
    txt_停止日 = Format(txt_停止日, "yy/mm/dd")    '体裁を整える
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_停止日, KeyAscii, P_DATE_LENGTH)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日_LostFocus()
    txt_停止日 = Format(Me.txt_停止日, "yyyy/mm/dd") '体裁を整える
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_備考_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2006/06/05
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_備考_AfterUpdate()
    Call commonAfterUpdate
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_備考_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2006/06/05
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_備考_BeforeUpdate(Cancel As Integer)

    If chkTextInput(txt_備考, 60) <> 0 Then
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_備考_KeyPress
'        NOTE             :
'        CREATE_DATE      :2006/06/05
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_備考_KeyPress(KeyAscii As Integer)
    Call CmKeyPress.CmKeyPress_For_Char2(txt_備考, KeyAscii, 60)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却日_AfterUpdate()
    txt_返却日 = Format(Me.txt_返却日, "yyyy/mm/dd")  '体裁を整える
    Call commonAfterUpdate
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却日_BeforeUpdate(Cancel As Integer)
    
    '日付入力チェック
    If dateCheck(txt_返却日, Cancel) = False Then
        Exit Sub
    End If
    
    If Format(txt_貸出日, "yyyymmdd") > Format(txt_返却日, "yyyymmdd") Then
        MsgBox "返却日は貸出日より未来を入力してください。", vbExclamation, P_ERROR
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却日_GotFocus()

    txt_返却日 = Format(txt_返却日, "yy/mm/dd")  '体裁を整える

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_返却日, KeyAscii, P_DATE_LENGTH)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_返却日_LostFocus()
    txt_返却日 = Format(Me.txt_返却日, "yyyy/mm/dd") '体裁を整える
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_削除_Click()
    
    Dim sqlText   As String
    Dim dbObject  As Database
    
    On Error GoTo Exception
    
    If Nz(txt_カード№, "") = "" Then Exit Sub
    
    '貸出中か確認
    If Is顧客貸出() = True Then
        If Nz(chk_ReturnDay.VALUE, 0) <> P_チェック済み And Nz(chk_StopDay.VALUE, 0) <> P_チェック済み Then
            MsgBox "現在お客様に貸出中としている為、削除できません。" & vbCr & _
                   "※返却済みならば返却チェックを付けてから削除して下さい。", vbExclamation, P_ERROR
            Exit Sub
        End If
    End If
    
    '選択データ削除確認
    If MsgBox("データを削除します。よろしいですか？", _
                vbCritical + vbYesNo + vbDefaultButton2, "確認") = vbNo Then
        Exit Sub
    End If
    
    ' ワークテーブル該当データ削除
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
    
    Set dbObject = CurrentDb()
    sqlText = "DELETE FROM FKS600_W01 WHERE SCRDM_CARD_NO = '" & txt_カード№.VALUE & "'"
    dbObject.Execute sqlText, dbFailOnError                 'SQL文実行
    Me.Requery                                          '一覧再検索
    
    Call commonAfterUpdate
    
    dbObject.Close
    Set dbObject = Nothing
    
    Exit Sub
Exception:
    If Not dbObject Is Nothing Then dbObject.Close: Set dbObject = Nothing
    Call MSZZ024_M00("cmb_削除_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :サブフォームの基本コントロールのプロパティ設定
'        MODULE_ID        :SetBaseProperty
'        NOTE             :親フォームが呼ぶ
'        CREATE_DATE      :2006/06/04
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetBaseProperty()

    If Me.Parent.IsSelectModeFKS600() = False Then

        'メニューから & ヤードメンテから
        txt_カード№.BackColor = P_BACKCOLOR_WHITE
        txt_カード№.Enabled = True
        txt_カード№.Locked = False
        txt_貸出先.BackColor = P_BACKCOLOR_WHITE
        '※貸出先の,Enabled,Lockedは Form_Current で行う
        lst_部屋№.Enabled = True
        lst_部屋№.Locked = False
        lst_部屋№.BackColor = P_BACKCOLOR_WHITE

        cmb_削除.Enabled = True
        Me.AllowAdditions = True
        
    Else
      
        '受付入力 & 契約入力から
        ' カード№、貸出先、部屋№を入力不可とする
        txt_カード№.BackColor = P_BACKCOLOR_YELLOW
        txt_カード№.Enabled = False
        txt_カード№.Locked = True
        txt_貸出先.BackColor = P_BACKCOLOR_YELLOW
        '※貸出先の,Enabled,Lockedは Form_Current で行う
        lst_部屋№.Enabled = False
        lst_部屋№.Locked = True
        lst_部屋№.BackColor = P_BACKCOLOR_YELLOW
        
        cmb_削除.Enabled = False
        Me.AllowAdditions = False
        
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面コントロールプロパティの設定
'        MODULE_ID        :setControlProperty
'        NOTE             :状態(用途)によってのコントロールプロパティの設定
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setControlProperty()
    
    '使用用途によるコントロール変更
        
    '①貸出先のプロパティ変更
    If Nz(Me.lst_用途.VALUE, "") = CmKomsMod.P_カード用途_業者貸値 Then
        ' 編集可能に
        Me.txt_貸出先.Enabled = True
        Me.txt_貸出先.Locked = False
    Else
        ' 編集不可に
        Me.txt_貸出先.Enabled = True
        Me.txt_貸出先.Locked = True
    End If
    
    '②貸出日のプロパティ変更
    If Nz(Me.chk_LendingDay.VALUE, 0) = P_チェック済み Then
        ' 編集可能に
        Me.txt_貸出日.Enabled = True
        Me.txt_貸出日.Locked = False
    Else
        ' 編集不可に
        Me.txt_貸出日.Enabled = False
        Me.txt_貸出日.Locked = True
    End If
    
    '③返却日のプロパティ変更
    If Nz(Me.chk_ReturnDay.VALUE, 0) = P_チェック済み Then
        ' 編集可能に
        Me.txt_返却日.Enabled = True
        Me.txt_返却日.Locked = False
    Else
        ' 編集不可に
        Me.txt_返却日.Enabled = False
        Me.txt_返却日.Locked = True
    End If
    
    '④停止日のプロパティ変更
    If Nz(Me.chk_StopDay.VALUE, 0) = P_チェック済み Then
        ' 編集可能に
        Me.txt_停止日.Enabled = True
        Me.txt_停止日.Locked = False
    Else
        ' 編集不可に
        Me.txt_停止日.Enabled = False
        Me.txt_停止日.Locked = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :AfterUpdate共通処理
'        MODULE_ID        :commonAfterUpdate
'        NOTE             :自フォームに修正が入ったならば親フォームに修正有無を通知
'                         :自フォームの修正フラグを立てる
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub commonAfterUpdate()
    
    If Nz(Me.chk_修正可否.VALUE, 0) = 0 Then
        Me.chk_修正可否.VALUE = -1
        Call Parent.SetFKS600EditChange '標題に"（修正中）"を付ける
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :chkTextInput
'        機能             :コンボボックス入力チェック
'        IN               :対象コンボボックス
'        OUT              :
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function chkTextInput(txtCheck As TextBox, intMaxLength As Integer) As Integer

    Dim strCode As String
    Dim intRet  As Integer
    
  On Error GoTo Exception
    intRet = 0
    
    strCode = Nz(txtCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo Finally
    End If
    
    '入力チェック
    intRet = check_Char(strCode, intMaxLength)
    Select Case intRet
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo Finally

    End Select
    
Finally:
    chkTextInput = intRet
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "chkTextInput" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :文字項目のチェックを一括でする。
'        MODULE_ID        :check_Char
'        CREATE_DATE      :
'
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function check_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

    On Error GoTo Exception

    check_Char = 9

    If CmLenb(oText) > ilen Then
       check_Char = 2
       Exit Function
    End If

    check_Char = 0
    Exit Function
    
Exception:
    Call Err.Raise(Err.Number, "check_Char" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :顧客貸出済み可否
'        MODULE_ID        :Is顧客貸出
'        NOTE             :元々の貸出と画面の変更内容を元に本カードを顧客に
'                         :貸出といているのか、それとも貸し出していないかを判断する
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Is顧客貸出() As Boolean

    Dim ret As Boolean

    ret = True

    ' 元々マスタでは顧客貸出で登録しているか
    If Nz(Me.txt_元契約番号, "") <> "" Then
    
        ' 画面編集上、貸出先契約が消えていれば
        If Nz(Me.txt_貸出先契約, "") = "" Then ret = False
        ' 返却日が入っていれば戻ってきた→貸し出してはいない
        ElseIf Nz(Me.txt_返却日, "") = "" Then ret = False
        ' 停止日が入っていれば無くした→貸し出してはいない
        ElseIf Nz(Me.txt_停止日, "") = "" Then ret = False

    Else
        ret = False
    End If
    
    Is顧客貸出 = ret

End Function

