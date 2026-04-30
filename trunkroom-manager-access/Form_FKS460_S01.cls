VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS460_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　価格テンプレートマスタメンテ
'   プログラムＩＤ　：
'   作　成　日　　　：  2017/11/07
'   作　成　者　　　：  Y.SUZUKI
'
'   更　新　日　　　：  2018/08/11
'   更　新　者　　　：  EGL
'
'**********************************************
Option Compare Database
Option Explicit

'**********************************************
' 定数宣言
'**********************************************
Private Const strMsg_err            As String = "エラー"
Private Const strMSG桁数オーバー    As String = "桁数オーバーしています。"
Private Const strmsg重複            As String = "段区分・用途・実帖の組合せが重複しています。"
Private Const strMSG入力不正        As String = "必須入力です。"

Private Const pcsglSizeMaxValue     As Single = 999.99   ' 実帖の入力最大値
Private Const pcsglMoneyMaxValue    As Long = 999999999  ' 金額項目の入力最大値
Private Const pcintMINUS            As Integer = 45      ' マイナス符号

Private Const pcstrfmtSTANDARD      As String = "Standard"  ' 書式：標準
Private Const pcstrfmtFIXED         As String = "Fixed"     ' 書式：固定

Private Const pcstrZAPPI_ErrMsg_F   As String = "初回雑費nnで「無し」を選んだ場合「0」より大きい金額は入力できません。"
Private Const pcstrZAPPI_ErrMsg_M   As String = "毎月雑費nnで「無し」を選んだ場合「0」より大きい金額は入力できません。"
Private Const str事務手数料         As String = "63"    '2018/08/11
'==============================================================================*
'   MODULE_NAME :   「行削除」クリック時処理
'   MODULE_ID   :   cmd_削除
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'行削除
Private Sub cmd_削除_Click()
    
    On Error GoTo Exception
    
    Dim strSEQ      As String
    Dim lngRTN      As Long
    
    '入力チェック
    strSEQ = Nz(KKTPM_SEQ_NUM.VALUE, "")      '連番
    If "" = strSEQ Then
        Exit Sub
    End If
    
    '確認メッセージ
    lngRTN = MsgBox("削除します、よろしいですか？", vbCritical + vbYesNo + vbDefaultButton2, "確認")
    If vbNo = lngRTN Then
        Exit Sub
    End If
    
    '警告非表示
    doCmd.SetWarnings False
    
    doCmd.RunCommand acCmdDeleteRecord
    
    '警告表示
    doCmd.SetWarnings True
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME : チェックを押された場合
'   MODULE_ID   : KKTPM_FZAPPI_1_CHK_AfterUpdate
'   CREATE_DATE : 2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_CHK_AfterUpdate()
    If KKTPM_FZAPPI_1_CHK.VALUE = True Then
        KKTPM_FZAPPI_1_COST.VALUE = pcsglMoneyMaxValue
    Else
        KKTPM_FZAPPI_1_COST.VALUE = 0
    End If
End Sub

'==============================================================================*
'   MODULE_NAME : 重複コードチェック
'   MODULE_ID   : KKTPM_FZAPPI_1_CODE_BeforeUpdate
'   CREATE_DATE : 2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_CODE_BeforeUpdate(Cancel As Integer)
    Dim blnErrFlg As Boolean
    blnErrFlg = False
    If KKTPM_FZAPPI_1_CODE.VALUE = KKTPM_FZAPPI_2_CODE.VALUE And Nz(KKTPM_FZAPPI_1_CODE.VALUE, 0) <> 0 And Nz(KKTPM_FZAPPI_2_CODE.VALUE, 0) <> 0 Then
        blnErrFlg = True
    ElseIf KKTPM_FZAPPI_1_CODE.VALUE = KKTPM_EZAPPI_1_CODE.VALUE And Nz(KKTPM_FZAPPI_1_CODE.VALUE, 0) <> 0 And Nz(KKTPM_EZAPPI_1_CODE.VALUE, 0) <> 0 Then
        blnErrFlg = True
    End If
    If blnErrFlg = True Then
        MsgBox "１つの価格内で同じ雑費区分は指定できません。", vbExclamation, strMsg_err
        KKTPM_FZAPPI_1_CODE.Undo
        Cancel = True
    End If
End Sub

'==============================================================================*
'   MODULE_NAME : 重複コードチェック
'   MODULE_ID   : KKTPM_FZAPPI_2_CODE_BeforeUpdate
'   CREATE_DATE : 2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_CODE_BeforeUpdate(Cancel As Integer)
    Dim blnErrFlg As Boolean
    blnErrFlg = False
    If KKTPM_FZAPPI_1_CODE.VALUE = KKTPM_FZAPPI_2_CODE.VALUE And Nz(KKTPM_FZAPPI_1_CODE.VALUE, 0) <> 0 And Nz(KKTPM_FZAPPI_2_CODE.VALUE, 0) <> 0 Then
        blnErrFlg = True
    ElseIf KKTPM_FZAPPI_2_CODE.VALUE = KKTPM_EZAPPI_1_CODE.VALUE And Nz(KKTPM_FZAPPI_1_CODE.VALUE, 0) <> 0 And Nz(KKTPM_EZAPPI_1_CODE.VALUE, 0) <> 0 Then
        blnErrFlg = True
    End If
    If blnErrFlg = True Then
        MsgBox "１つの価格内で同じ雑費区分は指定できません。", vbExclamation, strMsg_err
        KKTPM_FZAPPI_2_CODE.Undo
        Cancel = True
    End If
End Sub

'==============================================================================*
'   MODULE_NAME : 重複コードチェック
'   MODULE_ID   : KKTPM_EZAPPI_1_CODE_BeforeUpdate
'   CREATE_DATE : 2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_CODE_BeforeUpdate(Cancel As Integer)
    Dim blnErrFlg As Boolean
    blnErrFlg = False
    If KKTPM_EZAPPI_1_CODE.VALUE = KKTPM_FZAPPI_1_CODE.VALUE And Nz(KKTPM_EZAPPI_1_CODE.VALUE, 0) <> 0 And Nz(KKTPM_FZAPPI_1_CODE.VALUE, 0) <> 0 Then
        blnErrFlg = True
    ElseIf KKTPM_EZAPPI_1_CODE.VALUE = KKTPM_FZAPPI_2_CODE.VALUE And Nz(KKTPM_EZAPPI_1_CODE.VALUE, 0) <> 0 And Nz(KKTPM_FZAPPI_2_CODE.VALUE, 0) <> 0 Then
        blnErrFlg = True
    End If
    If blnErrFlg = True Then
        MsgBox "１つの価格内で同じ雑費区分は指定できません。", vbExclamation, strMsg_err
        KKTPM_EZAPPI_1_CODE.Undo
        Cancel = True
    End If
End Sub

'==============================================================================*
'   MODULE_NAME :   「段区分」更新後処理
'   MODULE_ID   :   KKTPM_STEP
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_STEP_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「用途」更新後処理
'   MODULE_ID   :   KKTPM_USAGE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_USAGE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「用途」キー入力時処理
'   MODULE_ID   :   KKTPM_USAGE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_USAGE_KeyPress(KeyAscii As Integer)
    '入力を許可しない
    KeyAscii = 0
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「実帖」更新前処理
'   MODULE_ID   :   KKTPM_SIZE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_SIZE_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If Not Parent.fncTxtInputCheck(KKTPM_SIZE, pcsglSizeMaxValue, 5) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「実帖」更新後処理
'   MODULE_ID   :   KKTPM_SIZE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_SIZE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「実帖」キー入力時処理
'   MODULE_ID   :   KKTPM_SIZE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_SIZE_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    If KeyAscii = pcintMINUS Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If Not Parent.fncTxtInputCheck(KKTPM_SIZE, pcsglSizeMaxValue, 5) Then
        KeyAscii = 0: Exit Sub
    End If
    
    ' 桁数チェック
    Call ETC_CUREN_RTN(KKTPM_SIZE, 3, 2, KeyAscii)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「価格」更新前処理
'   MODULE_ID   :   KKTPM_PRICE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_PRICE_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If Not Parent.fncTxtInputCheck(KKTPM_PRICE, pcsglMoneyMaxValue, 9) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「価格」更新後処理
'   MODULE_ID   :   KKTPM_PRICE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_PRICE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「価格」フォーカス取得後処理
'   MODULE_ID   :   KKTPM_PRICE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_PRICE_GotFocus()
    
    On Error GoTo Exception
    
    '書式：固定(カンマ編集なし)
    KKTPM_PRICE.Format = pcstrfmtFIXED
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「価格」フォーカス喪失後処理
'   MODULE_ID   :   KKTPM_PRICE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_PRICE_LostFocus()
    
    On Error GoTo Exception
    
    '書式：標準(カンマ編集あり)
    KKTPM_PRICE.Format = pcstrfmtSTANDARD
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「価格」キー入力時処理
'   MODULE_ID   :   KKTPM_PRICE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_PRICE_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '桁数チェック
    Call CmKeyPress_For_Num1(KKTPM_PRICE, KeyAscii, 9)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費1区分」更新後処理
'   MODULE_ID   :   KKTPM_FZAPPI_1_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_CODE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    If KKTPM_FZAPPI_1_CODE = 0 And Nz(KKTPM_FZAPPI_1_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_F, "nn", "1"), vbInformation, strMsg_err
        KKTPM_FZAPPI_1_COST.SetFocus
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費1区分」キー入力時処理
'   MODULE_ID   :   KKTPM_FZAPPI_1_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_CODE_KeyPress(KeyAscii As Integer)
    '入力を許可しない
    KeyAscii = 0
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費1価格」更新前処理
'   MODULE_ID   :   KKTPM_FZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_COST_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If KKTPM_FZAPPI_1_CODE = 0 And Nz(KKTPM_FZAPPI_1_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_F, "nn", "1"), vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    If Not Parent.fncTxtInputCheck(KKTPM_FZAPPI_1_COST, pcsglMoneyMaxValue, 9) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費1価格」更新後処理
'   MODULE_ID   :   KKTPM_FZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_COST_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費1価格」フォーカス取得後処理
'   MODULE_ID   :   KKTPM_FZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_COST_GotFocus()
    
    On Error GoTo Exception
    
    '書式：固定(カンマ編集なし)
    KKTPM_FZAPPI_1_COST.Format = pcstrfmtFIXED
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費1価格」フォーカス喪失後処理
'   MODULE_ID   :   KKTPM_FZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_COST_LostFocus()
    
    On Error GoTo Exception
    
    '書式：標準(カンマ編集あり)
    KKTPM_FZAPPI_1_COST.Format = pcstrfmtSTANDARD
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費1価格」キー入力時処理
'   MODULE_ID   :   KKTPM_FZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_1_COST_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '桁数チェック
    Call CmKeyPress_For_Num1(KKTPM_FZAPPI_1_COST, KeyAscii, 9)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「初回雑費2区分」更新後処理
'   MODULE_ID   :   KKTPM_FZAPPI_2_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_CODE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    If KKTPM_FZAPPI_2_CODE = 0 And Nz(KKTPM_FZAPPI_2_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_F, "nn", "2"), vbInformation, strMsg_err
        KKTPM_FZAPPI_2_COST.SetFocus
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「初回雑費2区分」キー入力時処理
'   MODULE_ID   :   KKTPM_FZAPPI_2_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_CODE_KeyPress(KeyAscii As Integer)
    '入力を許可しない
    KeyAscii = 0
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費2価格」更新前処理
'   MODULE_ID   :   KKTPM_FZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_COST_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If KKTPM_FZAPPI_2_CODE = 0 And Nz(KKTPM_FZAPPI_2_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_F, "nn", "2"), vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    If Not Parent.fncTxtInputCheck(KKTPM_FZAPPI_2_COST, pcsglMoneyMaxValue, 9) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費2価格」更新後処理
'   MODULE_ID   :   KKTPM_FZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_COST_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費2価格」フォーカス取得後処理
'   MODULE_ID   :   KKTPM_FZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_COST_GotFocus()
    
    On Error GoTo Exception
    
    '書式：固定(カンマ編集なし)
    KKTPM_FZAPPI_2_COST.Format = pcstrfmtFIXED
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費2価格」フォーカス喪失後処理
'   MODULE_ID   :   KKTPM_FZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_COST_LostFocus()
    
    On Error GoTo Exception
    
    '書式：標準(カンマ編集あり)
    KKTPM_FZAPPI_2_COST.Format = pcstrfmtSTANDARD
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費2価格」キー入力時処理
'   MODULE_ID   :   KKTPM_FZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_2_COST_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '桁数チェック
    Call CmKeyPress_For_Num1(KKTPM_FZAPPI_2_COST, KeyAscii, 9)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「初回雑費3区分」更新後処理
'   MODULE_ID   :   KKTPM_FZAPPI_3_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_3_CODE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    If KKTPM_FZAPPI_3_CODE = 0 And Nz(KKTPM_FZAPPI_3_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_F, "nn", "3"), vbInformation, strMsg_err
        KKTPM_FZAPPI_3_COST.SetFocus
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「初回雑費3区分」キー入力時処理
'   MODULE_ID   :   KKTPM_FZAPPI_3_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_3_CODE_KeyPress(KeyAscii As Integer)
    '入力を許可しない
    KeyAscii = 0
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費3価格」更新前処理
'   MODULE_ID   :   KKTPM_FZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_3_COST_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If KKTPM_FZAPPI_3_CODE = 0 And Nz(KKTPM_FZAPPI_3_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_F, "nn", "3"), vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    If Not Parent.fncTxtInputCheck(KKTPM_FZAPPI_3_COST, pcsglMoneyMaxValue, 9) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費3価格」更新後処理
'   MODULE_ID   :   KKTPM_FZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_3_COST_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費3価格」フォーカス取得後処理
'   MODULE_ID   :   KKTPM_FZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_3_COST_GotFocus()
    
    On Error GoTo Exception
    
    '書式：固定(カンマ編集なし)
    KKTPM_FZAPPI_3_COST.Format = pcstrfmtFIXED
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費3価格」フォーカス喪失後処理
'   MODULE_ID   :   KKTPM_FZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_3_COST_LostFocus()
    
    On Error GoTo Exception
    
    '書式：標準(カンマ編集あり)
    KKTPM_FZAPPI_3_COST.Format = pcstrfmtSTANDARD
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「初回雑費3価格」キー入力時処理
'   MODULE_ID   :   KKTPM_FZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_FZAPPI_3_COST_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '桁数チェック
    Call CmKeyPress_For_Num1(KKTPM_FZAPPI_3_COST, KeyAscii, 9)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「毎月雑費1区分」更新後処理
'   MODULE_ID   :   KKTPM_EZAPPI_1_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_CODE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    If KKTPM_EZAPPI_1_CODE = 0 And Nz(KKTPM_EZAPPI_1_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_M, "nn", "1"), vbInformation, strMsg_err
        KKTPM_EZAPPI_1_COST.SetFocus
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「毎月雑費1区分」キー入力時処理
'   MODULE_ID   :   KKTPM_EZAPPI_1_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_CODE_KeyPress(KeyAscii As Integer)
    '入力を許可しない
    KeyAscii = 0
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費1価格」更新前処理
'   MODULE_ID   :   KKTPM_EZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_COST_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If KKTPM_EZAPPI_1_CODE = 0 And Nz(KKTPM_EZAPPI_1_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_M, "nn", "1"), vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    If Not Parent.fncTxtInputCheck(KKTPM_EZAPPI_1_COST, pcsglMoneyMaxValue, 9) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費1価格」更新後処理
'   MODULE_ID   :   KKTPM_EZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_COST_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費1価格」フォーカス取得後処理
'   MODULE_ID   :   KKTPM_EZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_COST_GotFocus()
    
    On Error GoTo Exception
    
    '書式：固定(カンマ編集なし)
    KKTPM_EZAPPI_1_COST.Format = pcstrfmtFIXED
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費1価格」フォーカス喪失後処理
'   MODULE_ID   :   KKTPM_EZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_COST_LostFocus()
    
    On Error GoTo Exception
    
    '書式：標準(カンマ編集あり)
    KKTPM_EZAPPI_1_COST.Format = pcstrfmtSTANDARD
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費1価格」キー入力時処理
'   MODULE_ID   :   KKTPM_EZAPPI_1_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_1_COST_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '桁数チェック
    Call CmKeyPress_For_Num1(KKTPM_EZAPPI_1_COST, KeyAscii, 9)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「毎月雑費2区分」更新後処理
'   MODULE_ID   :   KKTPM_EZAPPI_2_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_2_CODE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    If KKTPM_EZAPPI_2_CODE = 0 And Nz(KKTPM_EZAPPI_2_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_M, "nn", "2"), vbInformation, strMsg_err
        KKTPM_EZAPPI_2_COST.SetFocus
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「毎月雑費2区分」キー入力時処理
'   MODULE_ID   :   KKTPM_EZAPPI_2_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_2_CODE_KeyPress(KeyAscii As Integer)
    '入力を許可しない
    KeyAscii = 0
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費2価格」更新前処理
'   MODULE_ID   :   KKTPM_EZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_2_COST_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If KKTPM_EZAPPI_2_CODE = 0 And Nz(KKTPM_EZAPPI_2_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_M, "nn", "2"), vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    If Not Parent.fncTxtInputCheck(KKTPM_EZAPPI_2_COST, pcsglMoneyMaxValue, 9) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費2価格」更新後処理
'   MODULE_ID   :   KKTPM_EZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_2_COST_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費2価格」フォーカス取得後処理
'   MODULE_ID   :   KKTPM_EZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_2_COST_GotFocus()
    
    On Error GoTo Exception
    
    '書式：固定(カンマ編集なし)
    KKTPM_EZAPPI_2_COST.Format = pcstrfmtFIXED
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費2価格」フォーカス喪失後処理
'   MODULE_ID   :   KKTPM_EZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_2_COST_LostFocus()
    
    On Error GoTo Exception
    
    '書式：標準(カンマ編集あり)
    KKTPM_EZAPPI_2_COST.Format = pcstrfmtSTANDARD
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費2価格」キー入力時処理
'   MODULE_ID   :   KKTPM_EZAPPI_2_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_2_COST_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '桁数チェック
    Call CmKeyPress_For_Num1(KKTPM_EZAPPI_2_COST, KeyAscii, 9)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「毎月雑費3区分」更新後処理
'   MODULE_ID   :   KKTPM_EZAPPI_3_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_3_CODE_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    If KKTPM_EZAPPI_3_CODE = 0 And Nz(KKTPM_EZAPPI_3_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_M, "nn", "3"), vbInformation, strMsg_err
        KKTPM_EZAPPI_3_COST.SetFocus
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'=============================================================================*
'   MODULE_NAME :   「毎月雑費3区分」キー入力時処理
'   MODULE_ID   :   KKTPM_EZAPPI_3_CODE
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_3_CODE_KeyPress(KeyAscii As Integer)
    '入力を許可しない
    KeyAscii = 0
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費3価格」更新前処理
'   MODULE_ID   :   KKTPM_EZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_3_COST_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    If KKTPM_EZAPPI_3_CODE = 0 And Nz(KKTPM_EZAPPI_3_COST, 0) > 0 Then
        MsgBox Replace(pcstrZAPPI_ErrMsg_M, "nn", "3"), vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    If Not Parent.fncTxtInputCheck(KKTPM_EZAPPI_3_COST, pcsglMoneyMaxValue, 9) Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費3価格」更新後処理
'   MODULE_ID   :   KKTPM_EZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_3_COST_AfterUpdate()
    
    On Error GoTo Exception
    
    Call Parent.set_ModifyFlg(True)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費3価格」フォーカス取得後処理
'   MODULE_ID   :   KKTPM_EZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_3_COST_GotFocus()
    
    On Error GoTo Exception
    
    '書式：固定(カンマ編集なし)
    KKTPM_EZAPPI_3_COST.Format = pcstrfmtFIXED
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費3価格」フォーカス喪失後処理
'   MODULE_ID   :   KKTPM_EZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_3_COST_LostFocus()
    
    On Error GoTo Exception
    
    '書式：標準(カンマ編集あり)
    KKTPM_EZAPPI_3_COST.Format = pcstrfmtSTANDARD
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   「毎月雑費3価格」キー入力時処理
'   MODULE_ID   :   KKTPM_EZAPPI_3_COST
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KKTPM_EZAPPI_3_COST_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Exception
    
    '桁数チェック
    Call CmKeyPress_For_Num1(KKTPM_EZAPPI_3_COST, KeyAscii, 9)
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   更新前処理
'   MODULE_ID   :   Form_FKS460_S01
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo Exception
    
    '入力チェック
    If Not CHK_INPUT_DATA Then
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME :   キー入力時処理
'   MODULE_ID   :   Form_FKS460_S01
'   CREATE_DATE :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Exception
    
    'F1=「終了」、 F2=「保存・削除」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Parent.cmd_終了.SetFocus
            
        Case vbKeyF2
            KeyCode = 0
            If Parent.cmd_保存.Enabled Then Parent.cmd_保存.SetFocus
            If Parent.cmd_削除.Enabled Then Parent.cmd_削除.SetFocus
            
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
            
    End Select
    
    Exit Sub
    
Exception:
    'メッセージ・ログ出力
    Call MSZZ024_M00("opg_処理選択_BeforeUpdate", True)
    
End Sub

'==============================================================================*
'   MODULE_NAME     :   必須入力項目の入力チェック
'   Parameter       :   なし
'   Return Value    :   True：正常終了、False：異常終了
'   CREATE_DATE     :   2017/11/07
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_INPUT_DATA() As Boolean
    
    Dim strSEQ      As String
    Dim strStep     As String
    Dim strUsage    As String
    Dim strSize     As String
    
    Dim dbAcc       As Database
    Dim strDML      As String
    Dim rsWork      As Recordset
    Dim lngCnt      As Long
    
    Dim strMsg      As String
    Dim blnRTN      As Boolean
    
    On Error GoTo Exception
    
    Set dbAcc = CurrentDb
    
    '******************************
    '必須チェック
    '******************************
    strSEQ = Nz(KKTPM_SEQ_NUM.VALUE, "")      '連番
    strStep = Nz(KKTPM_STEP.VALUE, "")        '段区分
    strUsage = Nz(KKTPM_USAGE.VALUE, "")      '用途
    strSize = Nz(KKTPM_SIZE.VALUE, "")        '実帖
    
    '2018/08/11 事務手数料の場合は必須入力不要
    If Me.Parent!cmb_テンプレートタイプ.VALUE <> str事務手数料 Then
    
        If "" = strStep Then
            MsgBox "段区分は" & strMSG入力不正, vbInformation, strMsg_err
            KKTPM_STEP.SetFocus
            Exit Function
        End If
        
        If "" = strUsage Then
            MsgBox "用途は" & strMSG入力不正, vbInformation, strMsg_err
            KKTPM_USAGE.SetFocus
            Exit Function
        End If
        
        If "" = strSize Then
            MsgBox "実帖は" & strMSG入力不正, vbInformation, strMsg_err
            KKTPM_SIZE.SetFocus
            Exit Function
        End If
        
    End If
    
    '******************************
    '重複チェック
    '******************************
    '「段区分」「用途」「実帖」で同じ明細は重複エラーとする
    'SQL編集して検索
    
'2018/08/11 null対応
'    strDML = "SELECT COUNT(*) AS CNT" _
'           & " FROM FKS460_W01" _
'           & " WHERE KKTPM_STEP = " & strSTEP _
'           & " AND KKTPM_USAGE = " & strUSAGE _
'           & " AND KKTPM_SIZE = " & strSize _
'           & " AND KKTPM_SEQ_NUM <> " & strSEQ
    
    strDML = "SELECT COUNT(*) AS CNT" _
           & " FROM FKS460_W01" _
           & " WHERE KKTPM_STEP = " & fncSqlNull(strStep) _
           & " AND KKTPM_USAGE = " & fncSqlNull(strUsage) _
           & " AND KKTPM_SIZE = " & fncSqlNull(strSize) _
           & " AND KKTPM_SEQ_NUM <> " & strSEQ

    Set rsWork = dbAcc.OpenRecordset(strDML, dbOpenSnapshot)
    
    'データがない場合
    If Not rsWork.EOF Then
        '重複データが在る場合
        lngCnt = Nz(rsWork.Fields("CNT"), 0)
        If 0 < lngCnt Then
            MsgBox strmsg重複, vbInformation, strMsg_err
            GoTo Exit_rtn
            
        End If
        
    End If
    
    CHK_INPUT_DATA = True
    
Exit_rtn:
    If Not rsWork Is Nothing Then rsWork.Close: Set rsWork = Nothing
    If Not dbAcc Is Nothing Then dbAcc.Close: Set dbAcc = Nothing
    
    Exit Function
    
Exception:
    If Not rsWork Is Nothing Then rsWork.Close: Set rsWork = Nothing
    If Not dbAcc Is Nothing Then dbAcc.Close: Set dbAcc = Nothing
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "fncDisConnectDB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Function

'==============================================================================*
'   MODULE_NAME :   null変換
'   MODULE_ID   :   fncSqlNull
'   CREATE_DATE :   2018/08/11
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSqlNull(strTgt As String) As String
    If Trim(fncSqlNull) = "" Then
        fncSqlNull = "null"
    Else
        fncSqlNull = strTgt
    End If
End Function
