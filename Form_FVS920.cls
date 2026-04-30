VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS920"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　メンテナンス依頼書出力条件入力
'   プログラムＩＤ　：　FVS920
'   作　成　日　　　：  2006/09/01
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：  2007/02/14
'   修　正　者      ：  tajima
'   修　正　内　容　：  ヤード修理依頼書を印刷対応
'                       機能名を修理依頼書とした
'   Ver             ：  0.1
'
'**********************************************
'修正履歴
'   修　正　日　　　：  2011/07/30
'   修　正　者      ：  M.RYU
'   修　正　内　容　：  検索条件を追加
'                       修理区分と完了日
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：  2013/06/03
'   修　正　者      ：  M.HONDA
'   修　正　内　容　：  終了メッセージを表示させないように修正
'   Ver             ：  0.3
'
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
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

'桁数
Private Const P_CODE_LENGTH         As Integer = 6              'ヤードコード
Private Const P_ZERO_SUPPRESS       As String = "0000000000"    '不足桁０埋め用

'印刷種別 2007/02/14 add tajima
Private Const P_印刷種別_部屋修理   As Integer = 1
Private Const P_印刷種別_ヤード修理   As Integer = 2

'==============================================================================*
'
'        MODULE_NAME      :chkComplete_AfterUpdate
'        機能             :chkComplete_AfterUpdate イベント
'        CREAT            :2011/07/30 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkComplete_AfterUpdate()
    Me.txtCompdFrom.Enabled = chkComplete.VALUE
    Me.txtCompdTo.Enabled = chkComplete.VALUE
    
    If chkComplete.VALUE = False Then
        Me.txtCompdFrom = ""
        Me.txtCompdTo = ""
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chkRoom_AfterUpdate
'        機能             :chkRoom_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkRoom_AfterUpdate()
    
    'チェックされたら部屋コンボを使用可、ヤードTOを使用不可
    'チェックを外されたら部屋コンボを使用不可、ヤードTOを使用可
    cmbRoom.Enabled = chkRoom.VALUE
    cmbYardTo.Enabled = Not chkRoom.VALUE
    
    cmbRoom = ""
    cmbYardTo = ""
    Call cmbYardTo_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      : 修理区分From更新後
'        MODULE_ID        : cmbReprFrom_AfterUpdate
'        CREAT_DATE       : 2011/07/30 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReprFrom_AfterUpdate()
    Me.txtReprFrom = IIf(Nz(Me.cmbReprFrom) = "", "", Nz(Me.cmbReprFrom.Column(1)))
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 修理区分To更新後
'        MODULE_ID        : cmbReprTo_AfterUpdate
'        CREAT_DATE       : 2011/07/30 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReprTo_AfterUpdate()
    Me.txtReprTo = IIf(Nz(Me.cmbReprTo) = "", "", Nz(Me.cmbReprTo.Column(1)))
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbRoom_AfterUpdate
'        機能             :cmbRoom_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRoom_AfterUpdate()

    Call fncCmbAfterUpdate(cmbRoom, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbRoom_BeforeUpdate
'        機能             :cmbRoom_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRoom_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call fncChkComboInput(cmbRoom, P_CODE_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbRoom_KeyPress
'        機能             :cmbRoom_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRoom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbRoom, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardFrom_AfterUpdate
'        機能             :cmbYardFrom_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardFrom_AfterUpdate()
    
    Dim strSQL      As String
    
    txtYardFrom = fncCmbAfterUpdate(cmbYardFrom, P_CODE_LENGTH)
    cmbRoom = ""
    If Nz(cmbYardFrom) = "" Then
        cmbRoom.RowSource = ""
    Else
        strSQL = strSQL & " SELECT DISTINCT "
        strSQL = strSQL & "        Format([CNTA_NO],""000000"") "
        strSQL = strSQL & "   FROM dbo_CNTA_MAST "
        strSQL = strSQL & "  WHERE [CNTA_CODE]=[Forms]![FVS920]![cmbYardFrom] "
        strSQL = strSQL & "  ORDER BY 1; "
        cmbRoom.RowSource = strSQL
    End If
    cmbRoom.Requery

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardFrom_BeforeUpdate
'        機能             :cmbYardFrom_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardFrom_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call fncChkComboInput(cmbYardFrom, P_CODE_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardFrom_KeyPress
'        機能             :cmbYardFrom_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardFrom_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmbYardFrom, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardTo_AfterUpdate
'        機能             :cmbYardTo_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardTo_AfterUpdate()
    
    txtYardTo = fncCmbAfterUpdate(cmbYardTo, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardTo_BeforeUpdate
'        機能             :cmbYardTo_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardTo_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call fncChkComboInput(cmbYardTo, P_CODE_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbYardTo_KeyPress
'        機能             :cmbYardTo_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbYardTo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmbYardTo, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncCmbAfterUpdate
'        機能             :更新後イベント時の処理
'                         :０サプレスした値を対象コントロールに設定し、名称を返却
'        IN               :①対象コントロール
'                         :②入力最大桁数
'        OUT              :名称
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCmbAfterUpdate(objTarget As Object, intLength As Integer) As String

    fncCmbAfterUpdate = ""
    
    '０埋め
    objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    
    '名称を返却する
    If objTarget.ColumnCount > 1 Then
        If Nz(objTarget) <> "" And _
           objTarget.ListCount <> 0 _
        Then
            fncCmbAfterUpdate = Nz(objTarget.Column(1, objTarget.ListIndex))
        End If
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkComboInput
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
Private Function fncChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer, ByRef Cancel As Integer) As Integer

    Dim strCode     As String
    Dim intRet      As Integer
    Dim intCount    As Integer
    Dim strValue    As String
    Dim blnFound    As Boolean
    
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
        If cmbCheck.ListIndex = -1 Then
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
            
            If Not blnFound Then
                MsgBox "該当するコードがありません。", vbInformation, P_KAKUNIN
                SendKeys "{ESC}"
                intRet = 2
                GoTo ChkComboInput_End
            End If
        End If
    End If
    
ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If
    
    fncChkComboInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :cmdEnd_Click
'        機能             :cmdEnd_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdEnd_Click()

    '' 2013/06/03 M.HONDA START
    ' 確認メッセージ表示
'    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
'        On Error Resume Next
'        Call psubClearWork
'        doCmd.Close acForm, Me.FormName
'    End If
    doCmd.Close acForm, Me.FormName
    '' 2013/06/03 M.HONDA END

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdExcel_Click
'        機能             :cmdExcel_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdExcel_Click()
On Error GoTo ErrorHandler
    
    Call subOutput(P910_MODE_EXCEL)     'Excel出力

    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdExcel_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdPreview_Click
'        機能             :cmdPreview_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdPreview_Click()
On Error GoTo ErrorHandler
    
    Call subOutput(P910_MODE_PREVIEW)   '印刷プレビュー表示
    
    Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmdPreview_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subOutput
'        機能             :出力処理
'        IN               :1=印刷プレビュー 2=EXCEL出力
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subOutput(intMode As Integer)
On Error GoTo ErrorHandler

    Dim blnRet  As Boolean
    Dim strMsg  As String
    
    '必須チェック
    If Not fncRequireCheck() Then
        Exit Sub
    End If
    
    Select Case intMode
    Case P910_MODE_PREVIEW:
        strMsg = "印刷内容を画面に出力します。よろしいですか？"
    Case P910_MODE_EXCEL:
        strMsg = "Excel出力を実行します。よろしいですか？"
    End Select
    
    '実行確認メッセージ
    If MsgBox(strMsg, vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    
    '出力
    '▼ 2007/02/14 変更 tajima
    Select Case frm_印刷種別.VALUE
          Case P_印刷種別_部屋修理
          
          '----UPDATE 2011/07/30 M.RYU--------<S>
'              blnRet = Print910.fncPrintMaintenanceRequest(intMode, , _
'                                                  Nz(cmbYardFrom, ""), _
'                                                  Nz(cmbYardTo, ""), _
'                                                  chkComplete.VALUE, _
'                                                  Nz(cmbRoom, ""))
                blnRet = Print910.fncPrintMaintenanceRequest(intMode, , _
                            Nz(cmbYardFrom, ""), Nz(cmbYardTo, ""), chkComplete.VALUE, _
                            Nz(cmbRoom, ""), Nz(Me.cmbReprFrom), Nz(Me.cmbReprTo), _
                            Nz(Me.txtCompdFrom), Nz(Me.txtCompdTo))
          '----UPDATE 2011/07/30 M.RYU--------<E>
          
          Case P_印刷種別_ヤード修理
          '----UPDATE 2011/07/30 M.RYU--------<S>
'              blnRet = Print901.fncPrintYardMaintenanceRequest(intMode, , _
'                                                  Nz(cmbYardFrom, ""), _
'                                                  Nz(cmbYardTo, ""), _
'                                                  chkComplete.VALUE)
    
                blnRet = Print901.fncPrintYardMaintenanceRequest(intMode, , _
                            Nz(cmbYardFrom, ""), Nz(cmbYardTo, ""), chkComplete.VALUE, _
                            Nz(Me.cmbReprFrom), Nz(Me.cmbReprTo), Nz(Me.txtCompdFrom), Nz(Me.txtCompdTo))
          '----UPDATE 2011/07/30 M.RYU--------<E>
    End Select
    '▲ 2007/02/14

    If Not blnRet Then
        MsgBox "該当データが存在しません。", vbCritical, P_ERROR
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subOutput" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :fncRequireCheck
'        機能             :必須入力チェック
'        IN               :
'        OUT              :True=必須入力が全て入力されている False=されていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRequireCheck() As Boolean

    fncRequireCheck = False
    
    '▼ 2007/02/14 変更 tajima
    Select Case frm_印刷種別.VALUE
        Case P_印刷種別_部屋修理
    
              If (cmbYardFrom.Enabled And Nz(cmbYardFrom.VALUE) = "") And _
                 (cmbYardTo.Enabled And Nz(cmbYardTo.VALUE) = "") _
              Then
                  MsgBox "ヤードコードＦＲＯＭ／ＴＯのいずれかは必須入力です。", vbInformation, P_KAKUNIN
                  cmbYardFrom.SetFocus
                  Exit Function
              End If
              
              If cmbRoom.Enabled And Nz(cmbRoom, "") = "" Then
                  MsgBox "部屋指定の場合、部屋番号は必須入力です。", vbInformation, P_KAKUNIN
                  cmbRoom.SetFocus
                  Exit Function
              End If
        Case P_印刷種別_ヤード修理
              
              If (cmbYardFrom.Enabled And Nz(cmbYardFrom.VALUE) = "") And _
                 (cmbYardTo.Enabled And Nz(cmbYardTo.VALUE) = "") _
              Then
                  MsgBox "ヤードコードＦＲＯＭ／ＴＯのいずれかは必須入力です。", vbInformation, P_KAKUNIN
                  cmbYardFrom.SetFocus
                  Exit Function
              End If
    
    End Select
    '▲ 2007/02/14
    
    fncRequireCheck = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activate イベント
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
'        機能             :Form_KeyDown イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'Ｆ１：終了へ，Ｆ２：印刷へ，Ｆ３：印刷へ
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmdEnd.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Me.cmdPreview.SetFocus
        Case vbKeyF3
            KeyCode = 0
            Me.cmdExcel.SetFocus
        Case vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form_Open イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    doCmd.Maximize
    Call chkRoom_AfterUpdate
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
    End Select
    
    Call frm_印刷種別_AfterUpdate   'INSERT 2011/07/30 M.RYU
    Call chkComplete_AfterUpdate    'INSERT 2011/07/30 M.RYU

End Sub

'2007/02/14 add new method
'==============================================================================*
'
'        MODULE_NAME      :frm_印刷種別_AfterUpdateイベント
'        MODULE_ID        :frm_印刷種別_AfterUpdate
'        CREATE_DATE      :2007/02/14
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub frm_印刷種別_AfterUpdate()
  
    Select Case frm_印刷種別.VALUE
        Case P_印刷種別_部屋修理
            chkRoom.Visible = True
            cmbRoom.Visible = True
            
            'INSERT 2011/07/30 M.RYU
            cmbReprFrom.RowSource = " SELECT DISTINCT " & _
                        " CODE_TABL.CODET_CODEC, CODE_TABL.CODET_NAMEN " & _
                        " FROM CODE_TABL WHERE CODE_TABL.CODET_SIKBC = '251';"

        
        Case P_印刷種別_ヤード修理
            chkRoom.Visible = False
            cmbRoom.Visible = False
              
            'INSERT 2011/07/30 M.RYU
            cmbReprFrom.RowSource = " SELECT DISTINCT " & _
                        " CODE_TABL.CODET_CODEC, CODE_TABL.CODET_NAMEN " & _
                        " FROM CODE_TABL WHERE CODE_TABL.CODET_SIKBC = '254';"
    
    End Select
    
    cmbReprTo.RowSource = cmbReprFrom.RowSource 'INSERT 2011/07/30 M.RYU----<s>
    cmbReprFrom = ""
    txtReprFrom = ""
    cmbReprTo = ""
    txtReprTo = ""                              'INSERT 2011/07/30 M.RYU----<e>
    
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :完了日From
'        MODULE_ID        :txtCompdFrom_KeyPress
'        CREATE_DATE      :2011/07/30 M.RYU
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtCompdFrom_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txtCompdFrom, KeyAscii, 8)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :完了日To
'        MODULE_ID        :txtCompdTo_KeyPress
'        CREATE_DATE      :2011/07/30 M.RYU
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtCompdTo_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txtCompdTo, KeyAscii, 8)
End Sub
