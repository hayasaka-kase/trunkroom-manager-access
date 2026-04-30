VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS700_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　解約入力サブ(カード入力)
'   プログラムＩＤ　：　FVS700_S01
'   作　成　日　　　：  2006/05/29
'   作　成　者　　　：  イーグルソフト 田島
'   Ver             ：  0.0
'   備考            ：
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   Ver             ：  0.1
'   修　正　内　容　：
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
Private Const P_ERROR               As String = "エラー"

'==============================================================================*
'
'        MODULE_NAME      :ChkTextInput
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
    
    intRet = 0
    
    strCode = Nz(txtCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkTextInput_End
    End If
    
    '入力チェック
    intRet = CmBCheck_Char(strCode, intMaxLength)
    Select Case intRet
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo ChkTextInput_End

        Case 9
            MsgBox "備考を入力してください。", vbInformation, P_ERROR
            SendKeys "{ESC}"
            GoTo ChkTextInput_End
    End Select
    
ChkTextInput_End:
    chkTextInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :文字項目のチェックを一括でする。
'        MODULE_ID        :CmBCheck_Char
'        CREATE_DATE      :
'
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

    On Error GoTo err_rtn

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
       CmBCheck_Char = 2
       GoTo Exit_rtn
    End If

    CmBCheck_Char = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmBCheck_Char = 9
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :chk_カード停止_AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_カード停止_AfterUpdate()
    Call Form_FVS700.ChangeSubForm  '親にサブフォームが変更された事を通知
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_カード停止_BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_カード停止_BeforeUpdate(Cancel As Integer)

    If Me.chk_カード返却.VALUE = -1 Then
        MsgBox "このカードは返却済みとしています。" & _
         vbCrLf & "停止とするならば先に返却チェックを外して下さい。", vbInformation, P_ERROR
        Cancel = True
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_カード返却_AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_カード返却_AfterUpdate()
    Call Form_FVS700.ChangeSubForm  '親にサブフォームが変更された事を通知
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_カード返却_BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_カード返却_BeforeUpdate(Cancel As Integer)

    If Me.chk_カード停止.VALUE = -1 Then
        MsgBox "このカードは停止中としています。" & _
         vbCrLf & "返却とするならば先に停止チェックを外して下さい。", vbInformation, P_ERROR
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_カード備考_AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード備考_AfterUpdate()
    Call Form_FVS700.ChangeSubForm  '親にサブフォームが変更された事を通知
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_カード備考_BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード備考_BeforeUpdate(Cancel As Integer)
    
    If chkTextInput(txt_カード備考, 60) <> 0 Then
        Cancel = True
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_カード備考_KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_カード備考_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char2(txt_カード備考, KeyAscii, 60)
    
End Sub


