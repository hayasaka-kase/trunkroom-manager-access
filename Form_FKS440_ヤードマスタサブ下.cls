VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS440_ヤードマスタサブ下"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　倉庫管理システム
'**********************************************
'   プログラム名　　：　ヤードマスタサブ上
'   プログラムＩＤ　：
'   作　成　日　　　：
'   作　成　者　　　：  αシステム
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/10/24
'   修　正　者　　　：　T.SUZUKI
'   修　正　内　容　：  コンテナ番号入力時、上下段の重複チェックを行うよう、処理を追加
'   Ver             ：　0.1
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/05/25
'   修　正　者　　　：　S.SHIBAZAKI
'   修　正　内　容　：  コンテナ追加の初期値を利用可にする(デザインのみ変更)
'   Ver             ：　0.2
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/06/09
'   修　正　者　　　：　S.SHIBAZAKI
'   修　正　内　容　：  鍵情報追加
'   Ver             ：　0.3
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/06/24
'   修　正　者　　　：　tajima
'   修　正　内　容　：  修繕呼出対応
'   Ver             ：　0.4
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/09/12
'   修　正　者　　　：　S.SHIBAZAKI
'   修　正　内　容　：  デザインのみ変更
'                  ：  1.削除ボタンをその他タブに移動
'                  ：  2.状態をその他タブからタブの外側に移動
'                  ：  3.修理中可否を上体の下に追加
'   Ver             ：　0.5
'**********************************************
'修正履歴
'   修　正　日　　　：　2008/10/15
'   修　正　者　　　：　iizuka
'   修　正　内　容　：  デザインのみ変更
'   Ver             ：　0.6
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2009/04/01
'   修　正　者　　　：　kihara
'   修　正　内　容　：  修繕入力しているコンテナ番号は、削除・変更できないようにする
'   Ver             ：　0.7
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2011/03/05
'   修　正　者　　　：　S.SHIBAZAKI
'   Ver             ：　2.0
'   修　正　内　容  ：  バイク可否フラグ追加
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2013/05/17
'   修　正　者　　　：　M.HONDA
'   Ver             ：　2.1
'   修　正　内　容  ：  幅・奥行きの位置を変更
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2014/03/07
'   修　正　者　　　：　MIYAMOTO
'   Ver             ：　2.2
'   修　正　内　容  ：  自動鍵用の解除番号を追加
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2014/04/16
'   修　正　者　　　：　MIYAMOTO
'   Ver             ：　2.3
'   修　正　内　容  ：  自動鍵用の解除番号発番の引数(コンテナ番号)を追加
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2014/04/18
'   修　正　者　　　：　MIYAMOTO
'   Ver             ：　2.4
'   修　正　内　容  ：  ダイヤル南京の場合に一度設定したあと変更出来ないように変更
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2016/05/16
'   修　正　者　　　：　SUGIMURA
'   Ver             ：　2.5
'   修　正　内　容  ：  鍵区分情報にて解除番号手動入力設定の場合は解除番号は手入力とするように追加・変更
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2016/12/06
'   修　正　者　　　：　M.HONDA
'   Ver             ：　2.6
'   修　正　内　容  ：  0.8帖～4帖の場合のみバイクにデフォルトのチェック行う。
'                   ：  帖に応じて自動的にサイズを登録する。
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2018/08/20
'   修　正　者　　　：　K.WADA
'   Ver             ：　2.7
'   修　正　内　容  ：  事務手数料登録機能、階数入力機能の追加
'
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：　2021/06/02
'   修　正　者　　　：　EGL
'   Ver             ：　2.8
'   修　正　内　容  ：  QR番号対応
'**********************************************

Option Compare Database
Option Explicit
    
Dim strMsg As String                         'メッセージ表示用
Dim intAns As Integer                        '確認入力用
Private Const strMsg_kaku           As String = "確認"
Private Const strMsg_err            As String = "エラー"
Private Const strmsg更新エラー      As String = "更新中にエラーが発生しました。" & vbCrLf & "（処理は中断されました。）"
Private Const strmsg重複            As String = "既に存在するコードです。"
Private Const strMSGデータなし      As String = "存在しないコードです。"
Private Const strMSG入力不正        As String = "を入力してください。"
Private Const strMSG桁数オーバー    As String = "桁数オーバーしています。"

Dim objDb                           As Database
Dim objRs                           As Recordset

Dim WRI_ObjWS                       As Workspace  ' 2005/10/24 DEL T.SUZUKI
Dim WRI_objDb                       As Database

' 2005/10/24 ADD T.SUZUKI Start
Private Const pcstrfmtSTANDARD      As String = "Standard"  ' 書式：標準
Private Const pcstrfmtFIXED         As String = "Fixed"     ' 書式：固定

Private Const strMSGCHENGE_ERR      As String = "このコンテナは変更できません。"

'Trueの場合は更新時の必須項目チェック等は行わない
Private blnCheckSkip                As Boolean
' 2005/10/24 ADD T.SUZUKI End

Private Const pcsglMaxValue         As Single = 999.99   ' 実帖の入力最大値      ' 2005/12/15 ADD T.SUZUKI
Private Const pcsglMoneyMaxValue    As Long = 999999999  ' 金額項目の入力最大値  ' 2005/12/17 ADD T.SUZUKI
Private Const pcsglFloorMaxValue    As Integer = 999     ' 階数の入力最大値      ' 2018/08/20 ADD K.WADA
Private Const pcintMINUS            As Integer = 45      ' マイナス符号          ' 2005/12/15 ADD T.SUZUKI

'↓INSERT 2011/03/05 SHIBAZAKI
Private Const pcstrUSAGE_コンテナ = "0"
Private Const pcstrUSAGE_バイク = "3"
Private Const pcstrUSAGE_バイク屋外 = "31"
Private Const pcstrUSAGE_バイク屋内 = "32"
Private Const pcstrUSAGE_バイクコンテナ = "33"
Private Const pcstrUSAGE_コンテナバイク = "39"

Private Const pcstrSTEP_上段 = "0"
Private Const pcstrSTEP_下段 = "1"
'↑INSERT 2011/03/05 SHIBAZAKI

'==============================================================================*
'
'        MODULE_NAME      :chk_CNTA_BIKE_OK_AfterUpdate
'        MODULE_ID        :chk_CNTA_BIKE_OK_AfterUpdate
'        CREATE_DATE      :2011/03/05 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_CNTA_BIKE_OK_AfterUpdate()

    Me.Parent.blnDataUpdate = True
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk_CNTA_BIKE_OK_BeforeUpdate
'        MODULE_ID        :chk_CNTA_BIKE_OK_BeforeUpdate
'        CREATE_DATE      :2011/03/05 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_CNTA_BIKE_OK_BeforeUpdate(Cancel As Integer)

    '下段コンテナならば変更可能
    Cancel = Not (Nz(Me.CNTA_USAGE.VALUE, "99") = pcstrUSAGE_コンテナ And Nz(txt_段区分) = pcstrSTEP_下段)
    
End Sub

'↓ INSERT 2014/03/07 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :chk_KaiKeyChange_AfterUpdateイベント
'        MODULE_ID        :chk_KaiKeyChange_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_KaiKeyChange_AfterUpdate()
    Dim strNameValFrom                  As String

'↓ INSERT 2016/05/16 SUGIMURA
    Dim strKaiKey                       As String
'↑ INSERT 2016/05/16 SUGIMURA


    strNameValFrom = Nz(cmb_KeyType.Column(2))
    '自動鍵タイプの場合(1=自動鍵、2=手動鍵)
'↓ UPDATE 2016/05/16 SUGIMURA
'    If strNameValFrom <> "1" Then
'       MsgBox "自動鍵以外はチェック出来ません。", vbInformation
    If MSZZ070_KaiKey(CInt(Parent.chk_地方ヤード), strNameValFrom) <> "1" And MSZZ070_KaiKey(CInt(Parent.chk_地方ヤード), strNameValFrom) <> "2" Then
        MsgBox "解除番号の入力は出来ません。", vbInformation
'↑ UPDATE 2016/05/16 SUGIMURA
        chk_KaiKeyChange = False
        Exit Sub
    End If
    
'↓ INSERT 2016/05/16 SUGIMURA
    strKaiKey = InputBox("解除番号を入力してください。" & Chr(13) & "(半角数字４桁)", "解除番号入力", txt_KaiKey)
      
    If StrPtr(strKaiKey) = 0 Then
        chk_KaiKeyChange = False
        Exit Sub
    ElseIf Len(strKaiKey) <> 0 And (IsNumeric(strKaiKey) = False Or Len(strKaiKey) > 4) Then
        MsgBox "半角数字4桁以内で入力してください。", vbExclamation
        chk_KaiKeyChange = False
        Exit Sub
    Else
        chk_KaiKeyChange = False
    End If
    txt_KaiKey = strKaiKey
'↑ INSERT 2016/05/16 SUGIMURA
End Sub
'↑ INSERT 2014/03/07 MIYAMOTO

'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :cmb_KeyType_BeforeUpdateイベント
'        MODULE_ID        :cmb_KeyType_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KeyType_AfterUpdate()
'↓ INSERT 2014/03/07 MIYAMOTO
    Dim strNameValFrom                  As String
'↑ INSERT 2014/03/07 MIYAMOTO

    txt_KeyTypeName = Parent.fncComboAfterUpdate(cmb_KeyType, 2)
'↓ INSERT 2014/03/07 MIYAMOTO
    strNameValFrom = Nz(cmb_KeyType.Column(2))
    '自動鍵タイプの場合(1=自動鍵)
'↓ UPDATE 2016/05/16 SUGIMURA
    'If strNameValFrom = "1" Then
    If MSZZ070_KaiKey(CInt(Parent.chk_地方ヤード), strNameValFrom) = "1" Then
'↑ UPDATE 2016/05/16 SUGIMURA
'↓ INSERT 2016/05/16 SUGIMURA
        txt_KaiKey.Locked = True
        txt_KaiKey.Enabled = False
'↑ INSERT 2016/05/16 SUGIMURA
'↓ UPDATE 2014/04/16 MIYAMOTO
'        txt_KaiKey = MSZZ070_M00(Parent.pstrBumoCD, CLng(txt_ヤードコード), cmb_KeyType)
         txt_KaiKey = MSZZ070_M00(Parent.pstrBumoCD, CLng(txt_ヤードコード), CLng(txt_コンテナ番号), cmb_KeyType)
'↑ UPDATE 2014/04/16 MIYAMOTO
    
'↓insert start EGL 2021/06/02
        If cmb_KeyType.VALUE = 16 Then
            txt_KAINO_QR = MSZZ070_M01(Parent.pstrBumoCD, cmb_KeyType)
            If Parent.txt_解除番号権限 = 0 Then
                txt_KeyNo = "******"
            Else
                txt_KeyNo = txt_KAINO_QR
            End If
        Else
            txt_KeyNo.VALUE = ""
        End If
'↑insert end EGL 2021/06/02
    

'↓ INSERT 2016/05/16 SUGIMURA
    ElseIf MSZZ070_KaiKey(Parent.chk_地方ヤード, strNameValFrom) = "2" Then
        txt_KaiKey.Locked = False
        txt_KaiKey.Enabled = True
        txt_KaiKey = ""
        txt_KeyNo.VALUE = "" ' insert start EGL 2021/06/02
'↑ INSERT 2016/05/16 SUGIMURA
    Else
'↓ INSERT 2016/05/16 SUGIMURA
        txt_KaiKey.Locked = True
        txt_KaiKey.Enabled = False
'↑ INSERT 2016/05/16 SUGIMURA
        txt_KaiKey = ""
    End If
'↑ INSERT 2014/03/07 MIYAMOTO
    
'↓ insert start EGL 2021/06/02
    If Parent.txt_解除番号権限 = 0 Then
        txt_KaiKey.InputMask = "Password"
    Else
        txt_KaiKey.InputMask = ""
    End If
'↑ insert end EGL 2021/06/02

    Me.Parent.blnDataUpdate = True

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeyNo_AfterUpdateイベント
'        MODULE_ID        :txt_KeyNo_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyNo_AfterUpdate()
    
    Me.Parent.blnDataUpdate = True
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_KeyNo_KeyDownイベント
'        MODULE_ID        :txt_KeyNo_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyNo_KeyDown(KeyCode As Integer, Shift As Integer)
    'insert start EGL 2021/06/02
    If cmb_KeyType.VALUE = 16 And txt_KeyNo.VALUE <> "" Then
        KeyCode = 0
        Exit Sub
    End If
    'insert end EGL 2021/06/02
End Sub

'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :cmb_KeyType_BeforeUpdateイベント
'        MODULE_ID        :cmb_KeyType_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KeyType_BeforeUpdate(Cancel As Integer)

'↓ INSERT 2014/04/18 MIYAMOTO
    '変更前の鍵区分が自動鍵の場合、鍵区分を変更できない様にする
    If Me.txt_自動鍵区分 = "1" Then
        MsgBox "自動鍵が指定されている為、他の鍵区分には変更出来ません。", vbInformation
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
'↑ INSERT 2014/04/18 MIYAMOTO

    Call Parent.fncChkComboInput(cmb_KeyType, 2, Cancel)
    
End Sub

'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :cmb_KeyType_KeyPressイベント
'        MODULE_ID        :cmb_KeyType_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KeyType_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmb_KeyType, KeyAscii, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_FLOOR_AfterUpdateイベント
'        MODULE_ID        :CNTA_FLOOR_AfterUpdate
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_FLOOR_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_FLOOR_BeforeUpdateイベント
'        MODULE_ID        :CNTA_FLOOR_BeforeUpdate
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_FLOOR_BeforeUpdate(Cancel As Integer)
    If Me.Parent.fncTxtInputCheck(Me.CNTA_FLOOR, pcsglFloorMaxValue, 3) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_FLOOR_KeyPressイベント
'        MODULE_ID        :CNTA_FLOOR_KeyPress
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_FLOOR_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If Me.Parent.fncTxtInputCheck(Me.CNTA_FLOOR, pcsglFloorMaxValue, 3) = False Then
        KeyAscii = 0: Exit Sub
    End If
    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.CNTA_FLOOR, 3, 0, KeyAscii)
End Sub

'↓ INSERT 2014/03/07 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :Form_Currentイベント
'        MODULE_ID        :Form_Current
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Current()

    If Me.Recordset Is Nothing Then Exit Sub
    If Me.Recordset.RecordCount = 0 Then Exit Sub

    If chk_KaiKeyChange = True Then
        txt_KaiKey.Enabled = True
        txt_KaiKey.Locked = False
    Else
        txt_KaiKey.Enabled = False
        txt_KaiKey.Locked = True
    End If
    
'↓ INSERT 2016/05/16 SUGIMURA
    If Me.Parent.txt_解除番号権限 = "4" Then
        chk_KaiKeyChange.Enabled = True
        chk_KaiKeyChange.Locked = False
    Else
        chk_KaiKeyChange.Enabled = False
        chk_KaiKeyChange.Locked = False
    End If
'↑ INSERT 2016/05/16 SUGIMURA
    
End Sub
'↑ INSERT 2014/03/07 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :Form_Loadイベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    Set objDb = CurrentDb

    blnCheckSkip = False  ' 2005/11/26 ADD T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDownイベント
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode <> vbKeyF12 Then 'F12キー以外を押下したら                                     ' 2009/04/01 INSERT KIHARA
        Call CmReprMod.OpenFVS90n(KeyCode, Me.Parent.cmb_ヤードコード, Me.txt_コンテナ番号)  ' 2006/06/24 Add tajima
    End If                                                                                   ' 2009/04/01 INSERT KIHARA
    
    On Error Resume Next  ' 2005/12/15 ADD T.SUZUKI

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.Parent.cmd_終了.SetFocus

        Case vbKeyF2
            KeyCode = 0
            If Me.Parent.opg_処理選択.VALUE = 3 Then
                Me.Parent.cmd_削除.SetFocus
            Else
                Me.Parent.cmd_保存.SetFocus
            End If

        Case vbKeyF5
            KeyCode = 0
' 2005/11/21 REP T.SUZUKI Start
            ' F5キー押下によりヤードコードへフォーカス移動
            Me.Parent.cmb_ヤードコード.SetFocus
            Call Me.Parent.cmd_Top_Click
'            Me.Parent.FKS440_ヤードマスタサブ下.SetFocus
' 2005/11/21 REP T.SUZUKI End

'        Case vbKeyF3, vbKeyF6 To vbKeyF12  ' DELETE 2009/04/01 KIHARA
        Case vbKeyF3, vbKeyF6 To vbKeyF11   ' INSERT 2009/04/01 KIHARA
            KeyCode = 0
            
'INSERT 2009/04/01 KIHARA Start
        Case vbKeyF12
            ' 必須チェック
            If Not Me.Parent.RequireCheck(1) Then
            KeyCode = 0
                Exit Sub
            End If
            Call CmReprMod.OpenFVS90n(KeyCode, Me.Parent.cmb_ヤードコード, Me.txt_コンテナ番号)
            KeyCode = 0
'INSERT 2009/04/01 KIHARA End

        Case Else
            Exit Sub
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_BeforeUpdateイベント
'        MODULE_ID        :Form_BeforeUpdate
'        CREATE_DATE      :2005/11/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo err_rtn

    If blnCheckSkip Then
        Exit Sub
    End If
    
        ' 必須チェック
        If Not Me.Parent.RequireCheck(1) Then
            Cancel = True
            Exit Sub
        End If
        
    Exit Sub

err_rtn:
    MsgBox Err.Description, , strMsg_err
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Errorイベント
'        MODULE_ID        :Form_Error
'        CREATE_DATE      :2005/12/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Error(DataErr As Integer, Response As Integer)

    Dim strCtrlName As String

    Select Case DataErr
        Case 2113, 3761 '定型入力エラー（アプリケーションエラー）
            strCtrlName = Me.ActiveControl.NAME  ' アクティブなコントロール名を取得
            Select Case strCtrlName
                Case Me.CNTA_PRICE_DIFF.NAME
                    ' *** 料金差額 ***
                    If Me.Parent.fncTxtInputCheck(Me.CNTA_PRICE_DIFF, pcsglMoneyMaxValue, 9, 1) = False Then
                        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
                        SendKeys "{ESC}"
                    End If
                    Response = acDataErrContinue
            End Select
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Closeイベント
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    Set objRs = Nothing
    Set objDb = Nothing
End Sub


'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :txt_KeyNo_BeforeUpdateイベント
'        MODULE_ID        :txt_KeyNo_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyNo_BeforeUpdate(Cancel As Integer)

    Dim vCode       As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.txt_KeyNo, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case Parent.CmBCheck_Char(vCode, 10)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn

End Sub

'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :txt_KeyNo_KeyPressイベント
'        MODULE_ID        :txt_KeyNo_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyNo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char(txt_KeyNo, KeyAscii, 10)

End Sub

'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :txt_KeyNum_AfterUpdateイベント
'        MODULE_ID        :txt_KeyNum_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyNum_AfterUpdate()
    
    Me.Parent.blnDataUpdate = True

End Sub

'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :txt_KeyNum_BeforeUpdateイベント
'        MODULE_ID        :txt_KeyNum_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyNum_BeforeUpdate(Cancel As Integer)
    
    If Me.Parent.fncTxtInputCheck(Me.txt_KeyNum, 99, 2) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If

End Sub

'2006/06/09 SHIBAZAKI 鍵情報追加
'==============================================================================*
'
'        MODULE_NAME      :txt_KeyNum_KeyPressイベント
'        MODULE_ID        :txt_KeyNum_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KeyNum_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txt_KeyNum, KeyAscii, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_コンテナ番号_AfterUpdateイベント
'        MODULE_ID        :txt_コンテナ番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_コンテナ番号_AfterUpdate()

    Dim vCode       As Variant

    vCode = Nz(Me.txt_コンテナ番号, "")
    If vCode = "" Then
' 2005/11/26 DEL T.SUZUKI Start
'        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
''        Cancel = True
'        Me.txt_コンテナ番号.SetFocus
'        SendKeys "{ESC}"
'        Exit Sub
' 2005/11/26 DEL T.SUZUKI End
    Else
        If ((Nz(Me.txt_コンテナ番号, "") <> Nz(Me.txt_コンテナ番号_Old, "")) Or _
            (Nz(Me.txt_コンテナ番号, "") <> Nz(Me.txt_コンテナ番号.OldValue, ""))) Then

            ' コンテナマスタ重複チェック
            If CHK_CNTA_EXIST Then
                MsgBox strmsg重複, vbInformation + vbOKOnly, strMsg_kaku
                Me.txt_コンテナ番号.SetFocus
                SendKeys "{ESC}"
                Exit Sub
            End If
        End If
        ' INSERT START 2018/08/20 K.WADA
        Select Case Parent.pstrBumoCD
            Case "H"
                Me.CNTA_FLOOR = 1
            Case "8", "O"
                If IsNumeric(Left$(vCode, 1)) Then
                    Me.CNTA_FLOOR = Left$(vCode, 1)
                End If
        End Select
        ' INSERT END 2018/08/20 K.WADA
    End If
    
    Me.txt_コンテナ番号 = Format(Me.txt_コンテナ番号, "000000") 'Add 01-07-2000 Araki

    Me.Parent.blnDataUpdate = True  ' 2005/11/21 ADD T.SUZUKI

' 2005/11/26 DEL T.SUZUKI Start
'' 2005/10/24 ADD T.SUZUKI Start
'    ' コンテナ番号を入力した際、レコードを確定する
'    doCmd.Echo False
'    Me.Requery
'    doCmd.GoToRecord , , acLast
'    doCmd.Echo True
'' 2005/10/24 ADD T.SUZUKI End
' 2005/11/26 DEL T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_コンテナ番号_BeforeUpdateイベント
'        MODULE_ID        :txt_コンテナ番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_コンテナ番号_BeforeUpdate(Cancel As Integer)

    Dim vCode       As Variant

    On Error Resume Next

' 2005/11/04 ADD T.SUZUKI Start
' 2005/11/29 REP T.SUZUKI Start
    If ((Nz(Me.txt_コンテナ番号, "") <> Nz(Me.txt_コンテナ番号_Old, "")) Or _
        (Nz(Me.txt_コンテナ番号, "") <> Nz(Me.txt_コンテナ番号.OldValue, ""))) Then
'    If (Parent.opg_処理選択.Value = 2 And Me.txt_コンテナ番号 <> Me.txt_コンテナ番号.Tag) Then
' 2005/11/29 REP T.SUZUKI End
        If Parent.opg_処理選択.VALUE = 2 Then
            If (Nz(Me.txt_コンテナ番号_Old, "") <> "" Or Nz(Me.txt_コンテナ番号.OldValue, "")) <> "" Then  ' 2005/11/29 ADD T.SUZUKI
    '        If Nz(Me.txt_コンテナ番号.Tag, "") <> "" Then                                                ' 2005/11/29 DEL T.SUZUKI
                ' 削除チェック
                If CHK_DEL_CNTA(1) = -1 Then      ' 2005/11/29 ADD T.SUZUKI
    '            If CHK_DEL_CNTA(1) = False Then  ' 2005/11/29 DEL T.SUZUKI
                    MsgBox strMSGCHENGE_ERR, vbInformation + vbOKOnly, strMsg_kaku
                    Cancel = True
                    SendKeys "{ESC}"
                    Exit Sub
                End If
            End If
        End If

        If (Parent.opg_処理選択.VALUE = 1 Or Parent.opg_処理選択.VALUE = 2) Then

            vCode = Nz(Me.txt_コンテナ番号, "")
            If vCode = "" Then
            Else
                ' コンテナマスタ重複チェック
                If CHK_CNTA_EXIST Then
                    MsgBox strmsg重複, vbInformation + vbOKOnly, strMsg_kaku
                    Cancel = True
                    SendKeys "{ESC}"
                    Exit Sub
                End If
            End If
        End If
    End If
' 2005/11/04 ADD T.SUZUKI End

'    vCode = Nz(Me.txt_コンテナ番号, "")
'    If vCode = "" Then
'' 2005/11/26 DEL T.SUZUKI Start
''        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
''        Cancel = True
'''        Me.txt_コンテナ番号.SetFocus
''        SendKeys "{ESC}"
''        Exit Sub
'' 2005/11/26 DEL T.SUZUKI End
'    Else
'        ' コンテナマスタ重複チェック
'        If CHK_CNTA_EXIST Then
'            MsgBox strmsg重複, vbInformation + vbOKOnly, strMsg_kaku
'            Cancel = True
''            Me.txt_コンテナ番号.SetFocus
'            SendKeys "{ESC}"
'            Exit Sub
'        End If
'    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_コンテナ番号_Changeイベント
'        MODULE_ID        :txt_コンテナ番号_Change
'        CREATE_DATE      :2005/11/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_コンテナ番号_Change()
    On Error Resume Next
    ' 変更時、変更前の値をtagに格納しておく
    Me.txt_コンテナ番号_Old = Me.txt_コンテナ番号.OldValue   ' 205/11/29 ADD T.SUZUKI
'    Me.txt_コンテナ番号.Tag = Me.txt_コンテナ番号.OldValue  ' 205/11/29 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_コンテナ番号_DblClickイベント
'        MODULE_ID        :txt_コンテナ番号_DblClick
'        CREATE_DATE      :2006/06/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_コンテナ番号_DblClick(Cancel As Integer)
    Me.Parent.strSznCntNo = ""                                'INSERT 2009/04/01 KIHARA
    Call Form_KeyDown(CmReprMod.P_呼出_部屋, 0)
    Me.Parent.strSznCntNo = Me.txt_コンテナ番号               'INSERT 2009/04/01 KIHARA
End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_コンテナ番号_KeyPressイベント
'        MODULE_ID        :txt_コンテナ番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_コンテナ番号_KeyPress(KeyAscii As Integer)
    ' 桁数チェック
    Call CmKeyPress_For_Num1(Me.txt_コンテナ番号, KeyAscii, 6)  '桁数チェック追加 01-06-2000 Araki
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_種別_AfterUpdateイベント
'        MODULE_ID        :cmb_種別_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_種別_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_種別_BeforeUpdateイベント
'        MODULE_ID        :cmb_種別_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_種別_BeforeUpdate(Cancel As Integer)

    Dim vCode       As Variant
    Dim strRyak     As String
    Dim strWhere    As String

    On Error GoTo err_rtn

' 2005/11/26 DEL T.SUZUKI Start
'    vCode = Nz(Me.cmb_種別, "")
'    If vCode = "" Then
'        MsgBox "種別" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'        Exit Sub
'    End If
'
'    strWHERE = ""
'    strWHERE = strWHERE & " NAME_ID = '" & "040" & "'"
'    strWHERE = strWHERE & " AND NAME_CODE = " & (vCode)
'
'    strRYAK = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWHERE), "")
'
'    If strRYAK = "" Then
'        MsgBox strMSGデータなし, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
'
'    If Trim(Me.cmb_種別) <> "" Then
'        If Trim(Len(Me.cmb_種別)) > 2 Then
'            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
' 2005/11/26 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_種別_KeyPressイベント
'        MODULE_ID        :cmb_種別_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_種別_KeyPress(KeyAscii As Integer)

    ' 入力を許可しない
    KeyAscii = 0: Exit Sub

' 2005/10/24 DEL T.SUZUKI Start
'    ' 桁数チェック
'    Call CmKeyPress_For_Num1(Me.cmb_種別, KeyAscii, 2)  '桁数チェック追加 01-06-2000 Araki
' 2005/10/24 DEL T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_USAGE_AfterUpdateイベント
'        MODULE_ID        :CNTA_USAGE_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_USAGE_AfterUpdate()
    Me.Parent.blnDataUpdate = True

    Call subBikeOkCheckBox              'INSERT 2011/03/05 SHIBAZAKI

End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_USAGE_BeforeUpdateイベント
'        MODULE_ID        :CNTA_USAGE_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_USAGE_BeforeUpdate(Cancel As Integer)

    Dim vCode       As Variant
    Dim strRyak     As String
    Dim strWhere    As String

    On Error GoTo err_rtn

' 2005/11/26 DEL T.SUZUKI Start
'    vCode = Nz(Me.CNTA_USAGE, "")
'    If vCode = "" Then
'        MsgBox "用途" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'        Exit Sub
'    End If
'
'    strWHERE = ""
'    strWHERE = strWHERE & " NAME_ID = '" & "086" & "'"
'    strWHERE = strWHERE & " AND NAME_CODE = " & (vCode)
'
'    strRYAK = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWHERE), "")
'
'    If strRYAK = "" Then
'        MsgBox strMSGデータなし, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
'
'    If Trim(Me.CNTA_USAGE) <> "" Then
'        If Trim(Len(Me.CNTA_USAGE)) > 2 Then
'            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
' 2005/11/26 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_USAGE_KeyPressイベント
'        MODULE_ID        :CNTA_USAGE_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_USAGE_KeyPress(KeyAscii As Integer)

    ' 入力を許可しない
    KeyAscii = 0: Exit Sub

'    ' 桁数チェック
'    Call CmKeyPress_For_Num1(Me.CNTA_USAGE, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :tab_DownHeadder_Changeイベント
'        MODULE_ID        :tab_DownHeadder_Change
'        CREATE_DATE      :2005/11/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub tab_DownHeadder_Change()

    Dim intDetailPage As Integer
    Dim intHeaderPage As Integer

    ' タブのページが変更された際にタブのヘッダーが異なる場合の対応

    ' ヘッダーのページを取得
    intHeaderPage = Me.tab_DownHeadder.VALUE

    ' 詳細のページを取得
    intDetailPage = Me.tab_Down.VALUE

    ' ヘッダーのページと詳細のページと異なる場合、詳細の選択ページを変更する
    If intHeaderPage <> intDetailPage Then
        Me.tab_Down.Pages(intHeaderPage).SetFocus
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :tab_Down_Changeイベント
'        MODULE_ID        :tab_Down_Change
'        CREATE_DATE      :2005/10/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub tab_Down_Change()

    Dim intDetailPage As Integer
    Dim intHeaderPage As Integer

    ' タブのページが変更された際にタブのヘッダーが異なる場合の対応

    ' 詳細のページを取得
    intDetailPage = Me.tab_Down.VALUE

    ' ヘッダーのページを取得
    intHeaderPage = Me.tab_DownHeadder.VALUE

    ' 詳細のページとヘッダーのページが異なる場合、ヘッダーの選択ページを変更する
    If intDetailPage <> intHeaderPage Then
        Me.tab_DownHeadder.Pages(intDetailPage).SetFocus
    End If
End Sub

' 2005/11/21 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :CNTA_KEI_RENTKG_AfterUpdateイベント
''        MODULE_ID        :CNTA_KEI_RENTKG_AfterUpdate
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_KEI_RENTKG_AfterUpdate()
'
'    ' 保証金が未入力の場合、月額料金を設定する。
'    If Nz(Me.CNTA_SECUKG, "") = "" Then
'        Me.CNTA_SECUKG = Me.CNTA_KEI_RENTKG
'    End If
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_KEI_RENTKG_BeforeUpdateイベント
''        MODULE_ID        :CNTA_KEI_RENTKG_BeforeUpdate
''        CREATE_DATE      :2005/11/08
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_KEI_RENTKG_BeforeUpdate(Cancel As Integer)
'    If Nz(Me.txt_コンテナ番号, "") = "" Then
'        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        Exit Sub
'    End If
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_KEI_RENTKG_GotFocusイベント
''        MODULE_ID        :CNTA_KEI_RENTKG_GotFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_KEI_RENTKG_GotFocus()
'    Me.CNTA_KEI_RENTKG.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_KEI_RENTKG_LostFocusイベント
''        MODULE_ID        :CNTA_KEI_RENTKG_LostFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_KEI_RENTKG_LostFocus()
'    Me.CNTA_KEI_RENTKG.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_KEI_RENTKG_KeyPressイベント
''        MODULE_ID        :CNTA_KEI_RENTKG_KeyPress
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_KEI_RENTKG_KeyPress(KeyAscii As Integer)
'    ' 桁数チェック
'    Call CmKeyPress_For_Num1(Me.CNTA_KEI_RENTKG, KeyAscii, 9)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_SECUKG_BeforeUpdateイベント
''        MODULE_ID        :CNTA_SECUKG_BeforeUpdate
''        CREATE_DATE      :2005/11/08
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_SECUKG_BeforeUpdate(Cancel As Integer)
'    If Nz(Me.txt_コンテナ番号, "") = "" Then
'        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        Exit Sub
'    End If
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_SECUKG_KeyPressイベント
''        MODULE_ID        :CNTA_SECUKG_KeyPress
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_SECUKG_KeyPress(KeyAscii As Integer)
'    ' 桁数チェック
'    Call CmKeyPress_For_Num1(Me.CNTA_SECUKG, KeyAscii, 9)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ECUKG_GotFocusイベント
''        MODULE_ID        :CNTA_ECUKG_GotFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_SECUKG_GotFocus()
'    Me.CNTA_SECUKG.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_SECUKG_LostFocusイベント
''        MODULE_ID        :CNTA_SECUKG_LostFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_SECUKG_LostFocus()
'    Me.CNTA_SECUKG.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :cmb_ZAPPI_CODE1_BeforeUpdateイベント
''        MODULE_ID        :cmb_ZAPPI_CODE1_BeforeUpdate
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_ZAPPI_CODE1_BeforeUpdate(Cancel As Integer)
'
'    Dim vCode       As Variant
'    Dim strRYAK     As String
'    Dim strWHERE    As String
'
'    On Error GoTo err_rtn
'
'' 2005/11/08 ADD T.SUZUKI Start
'    If Nz(Me.txt_コンテナ番号, "") = "" Then
'        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        Exit Sub
'    End If
'' 2005/11/08 ADD T.SUZUKI End
'
'    vCode = Nz(Me.cmb_ZAPPI_CODE1, "")
'    If vCode = "" Then
'        MsgBox "雑費コード１" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'        Exit Sub
'    End If
'
'    strWHERE = ""
'    strWHERE = strWHERE & " NAME_ID = '" & "096" & "'"
'    strWHERE = strWHERE & " AND NAME_CODE = " & (vCode)
'
'    strRYAK = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWHERE), "")
'
'    If strRYAK = "" Then
'        MsgBox strMSGデータなし, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
'
'    If Trim(Me.cmb_ZAPPI_CODE1) <> "" Then
'        If Trim(Len(Me.cmb_ZAPPI_CODE1)) > 1 Then
'            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
'
'Exit_rtn:
'    Exit Sub
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'    GoTo Exit_rtn
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :cmb_ZAPPI_CODE1_KeyPressイベント
''        MODULE_ID        :cmb_ZAPPI_CODE1_KeyPress
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_ZAPPI_CODE1_KeyPress(KeyAscii As Integer)
'
'    ' 入力を許可しない
'    KeyAscii = 0: Exit Sub
'
''    ' 桁数チェック
''    Call CmKeyPress_For_Num1(Me.cmb_ZAPPI_CODE1, KeyAscii, 1)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_1_KeyPressイベント
''        MODULE_ID        :CNTA_ZAPPI_1_KeyPress
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_1_KeyPress(KeyAscii As Integer)
'    ' 桁数チェック
'    Call CmKeyPress_For_Num1(Me.CNTA_ZAPPI_1, KeyAscii, 9)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_1_BeforeUpdateイベント
''        MODULE_ID        :CNTA_ZAPPI_1_BeforeUpdate
''        CREATE_DATE      :2005/11/08
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_1_BeforeUpdate(Cancel As Integer)
'    If Nz(Me.txt_コンテナ番号, "") = "" Then
'        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        Exit Sub
'    End If
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_1_GotFocusイベント
''        MODULE_ID        :CNTA_ZAPPI_1_GotFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_1_GotFocus()
'    Me.CNTA_ZAPPI_1.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_1_LostFocusイベント
''        MODULE_ID        :CNTA_ZAPPI_1_LostFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_1_LostFocus()
'    Me.CNTA_ZAPPI_1.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :cmb_ZAPPI_CODE2_BeforeUpdateイベント
''        MODULE_ID        :cmb_ZAPPI_CODE2_BeforeUpdate
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_ZAPPI_CODE2_BeforeUpdate(Cancel As Integer)
'
'    Dim vCode       As Variant
'    Dim strRYAK     As String
'    Dim strWHERE    As String
'
'    On Error GoTo err_rtn
'
'' 2005/11/08 ADD T.SUZUKI Start
'    If Nz(Me.txt_コンテナ番号, "") = "" Then
'        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        Exit Sub
'    End If
'' 2005/11/08 ADD T.SUZUKI End
'
'    vCode = Nz(Me.cmb_ZAPPI_CODE2, "")
'    If vCode = "" Then
'        MsgBox "雑費コード２" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'        Exit Sub
'    End If
'
'    strWHERE = ""
'    strWHERE = strWHERE & " NAME_ID = '" & "096" & "'"
'    strWHERE = strWHERE & " AND NAME_CODE = " & (vCode)
'
'    strRYAK = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWHERE), "")
'
'    If strRYAK = "" Then
'        MsgBox strMSGデータなし, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
'
'    If Trim(Me.cmb_ZAPPI_CODE2) <> "" Then
'        If Trim(Len(Me.cmb_ZAPPI_CODE2)) > 1 Then
'            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
'
'Exit_rtn:
'    Exit Sub
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'    GoTo Exit_rtn
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :cmb_ZAPPI_CODE2_KeyPressイベント
''        MODULE_ID        :cmb_ZAPPI_CODE2_KeyPress
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_ZAPPI_CODE2_KeyPress(KeyAscii As Integer)
'
'    ' 入力を許可しない
'    KeyAscii = 0: Exit Sub
'
''    ' 桁数チェック
''    Call CmKeyPress_For_Num1(Me.cmb_ZAPPI_CODE2, KeyAscii, 1)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_2_KeyPressイベント
''        MODULE_ID        :CNTA_ZAPPI_2_KeyPress
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_2_KeyPress(KeyAscii As Integer)
'    ' 桁数チェック
'    Call CmKeyPress_For_Num1(Me.CNTA_ZAPPI_2, KeyAscii, 9)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_2_BeforeUpdateイベント
''        MODULE_ID        :CNTA_ZAPPI_2_BeforeUpdate
''        CREATE_DATE      :2005/11/08
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_2_BeforeUpdate(Cancel As Integer)
'    If Nz(Me.txt_コンテナ番号, "") = "" Then
'        MsgBox "コンテナ番号" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        Exit Sub
'    End If
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_2_GotFocusイベント
''        MODULE_ID        :CNTA_ZAPPI_2_GotFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_2_GotFocus()
'    Me.CNTA_ZAPPI_2.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :CNTA_ZAPPI_2_LostFocusイベント
''        MODULE_ID        :CNTA_ZAPPI_2_LostFocus
''        CREATE_DATE      :2005/10/24
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CNTA_ZAPPI_2_LostFocus()
'    Me.CNTA_ZAPPI_2.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
'End Sub
' 2005/11/21 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :CNTA_SIZE_AfterUpdateイベント
'        MODULE_ID        :CNTA_SIZE_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_SIZE_AfterUpdate()
    Me.Parent.blnDataUpdate = True
    Call subBikeOkCheckBox              'INSERT 2011/03/05 SHIBAZAKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_SIZE_BeforeUpdateイベント
'        MODULE_ID        :CNTA_SIZE_BeforeUpdate
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_SIZE_BeforeUpdate(Cancel As Integer)

    If Me.Parent.fncTxtInputCheck(Me.CNTA_SIZE, pcsglMaxValue, 5) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_SIZE_KeyPressイベント
'        MODULE_ID        :CNTA_SIZE_KeyPress
'        CREATE_DATE      :2005/10/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_SIZE_KeyPress(KeyAscii As Integer)
    On Error Resume Next  ' 2005/11/08 ADD T.SUZUKI

' 2005/12/15 ADD T.SUZUKI Start
    If KeyAscii = pcintMINUS Then KeyAscii = 0: Exit Sub

    If Me.Parent.fncTxtInputCheck(Me.CNTA_SIZE, pcsglMaxValue, 5) = False Then
        KeyAscii = 0: Exit Sub
    End If
' 2005/12/15 ADD T.SUZUKI End

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.CNTA_SIZE, 3, 2, KeyAscii)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_PRICE_DIFF_AfterUpdateイベント
'        MODULE_ID        :CNTA_PRICE_DIFF_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_PRICE_DIFF_AfterUpdate()
    Me.Parent.blnDataUpdate = True
    Call subBikeOkCheckBox              'INSERT 2011/03/05 SHIBAZAKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_PRICE_DIFF_GotFocusイベント
'        MODULE_ID        :CNTA_PRICE_DIFF_GotFocus
'        CREATE_DATE      :2005/11/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_PRICE_DIFF_GotFocus()
    Me.CNTA_PRICE_DIFF.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_PRICE_DIFF_LostFocusイベント
'        MODULE_ID        :CNTA_PRICE_DIFF_LostFocus
'        CREATE_DATE      :2005/11/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_PRICE_DIFF_LostFocus()
    Me.CNTA_PRICE_DIFF.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_PRICE_DIFF_KeyPressイベント
'        MODULE_ID        :CNTA_PRICE_DIFF_KeyPress
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_PRICE_DIFF_KeyPress(KeyAscii As Integer)
    ' 桁数チェック
    Call CmKeyPress_For_Num(Me.CNTA_PRICE_DIFF, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_REASON_AfterUpdateイベント
'        MODULE_ID        :CNTA_REASON_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_REASON_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_REASON_BeforeUpdateイベント
'        MODULE_ID        :CNTA_REASON_BeforeUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_REASON_BeforeUpdate(Cancel As Integer)

    Dim vCode       As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.CNTA_REASON, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case Parent.CmBCheck_Char(vCode, 60)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_REASON_KeyPressイベント
'        MODULE_ID        :CNTA_REASON_KeyPress
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_REASON_KeyPress(KeyAscii As Integer)
    ' 桁数チェック
    Call CmKeyPress_For_Char2(Me.CNTA_REASON, KeyAscii, 60)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_DEPTH_AfterUpdateイベント
'        MODULE_ID        :CNTA_DEPTH_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_DEPTH_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_DEPTH_BeforeUpdateイベント
'        MODULE_ID        :CNTA_DEPTH_BeforeUpdate
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_DEPTH_BeforeUpdate(Cancel As Integer)

    If Me.Parent.fncTxtInputCheck(Me.CNTA_DEPTH, pcsglMaxValue, 5) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_DEPTH_KeyPressイベント
'        MODULE_ID        :CNTA_DEPTH_KeyPress
'        CREATE_DATE      :2005/11/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_DEPTH_KeyPress(KeyAscii As Integer)
    On Error Resume Next  ' 2005/11/08 ADD T.SUZUKI

' 2005/12/15 ADD T.SUZUKI Start
    If KeyAscii = pcintMINUS Then KeyAscii = 0: Exit Sub

    If Me.Parent.fncTxtInputCheck(Me.CNTA_DEPTH, pcsglMaxValue, 5) = False Then
        KeyAscii = 0: Exit Sub
    End If
' 2005/12/15 ADD T.SUZUKI End

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.CNTA_DEPTH, 3, 2, KeyAscii)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_WIDTH_AfterUpdateイベント
'        MODULE_ID        :CNTA_WIDTH_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_WIDTH_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_WIDTH_BeforeUpdateイベント
'        MODULE_ID        :CNTA_WIDTH_BeforeUpdate
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_WIDTH_BeforeUpdate(Cancel As Integer)

    If Me.Parent.fncTxtInputCheck(Me.CNTA_WIDTH, pcsglMaxValue, 5) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_WIDTH_KeyPressイベント
'        MODULE_ID        :CNTA_WIDTH_KeyPress
'        CREATE_DATE      :2005/11/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_WIDTH_KeyPress(KeyAscii As Integer)
    On Error Resume Next  ' 2005/11/08 ADD T.SUZUKI

' 2005/12/15 ADD T.SUZUKI Start
    If KeyAscii = pcintMINUS Then KeyAscii = 0: Exit Sub

    If Me.Parent.fncTxtInputCheck(Me.CNTA_WIDTH, pcsglMaxValue, 5) = False Then
        KeyAscii = 0: Exit Sub
    End If
' 2005/12/15 ADD T.SUZUKI End

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.CNTA_WIDTH, 3, 2, KeyAscii)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_HEIGHT_AfterUpdateイベント
'        MODULE_ID        :CNTA_HEIGHT_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_HEIGHT_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_HEIGHT_BeforeUpdateイベント
'        MODULE_ID        :CNTA_HEIGHT_BeforeUpdate
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_HEIGHT_BeforeUpdate(Cancel As Integer)

    If Me.Parent.fncTxtInputCheck(Me.CNTA_HEIGHT, pcsglMaxValue, 5) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_HEIGHT_KeyPressイベント
'        MODULE_ID        :CNTA_HEIGHT_KeyPress
'        CREATE_DATE      :2005/11/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_HEIGHT_KeyPress(KeyAscii As Integer)
    On Error Resume Next  ' 2005/11/08 ADD T.SUZUKI

' 2005/12/15 ADD T.SUZUKI Start
    If KeyAscii = pcintMINUS Then KeyAscii = 0: Exit Sub

    If Me.Parent.fncTxtInputCheck(Me.CNTA_HEIGHT, pcsglMaxValue, 5) = False Then
        KeyAscii = 0: Exit Sub
    End If
' 2005/12/15 ADD T.SUZUKI End

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.CNTA_HEIGHT, 3, 2, KeyAscii)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_USE_AfterUpdateイベント
'        MODULE_ID        :CNTA_USE_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_USE_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_USE_BeforeUpdateイベント
'        MODULE_ID        :CNTA_USE_BeforeUpdate
'        CREATE_DATE      :2005/10/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_USE_BeforeUpdate(Cancel As Integer)

    Dim vCode       As Variant
    Dim strRyak     As String
    Dim strWhere    As String

    On Error GoTo err_rtn

' 2005/12/05 ADD T.SUZUKI Start
    ' 「9:撤去」は選択不可とする
    If Nz(Me.CNTA_USE, "") = 9 Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
        Exit Sub
    End If
' 2005/12/05 ADD T.SUZUKI End

' 2005/11/26 DEL T.SUZUKI Start
'    vCode = Nz(Me.CNTA_USE, "")
'    If vCode = "" Then
'        MsgBox "利用" & strMSG入力不正, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'        Exit Sub
'    End If
'
'    strWHERE = ""
'    strWHERE = strWHERE & " NAME_ID = '" & "094" & "'"
'    strWHERE = strWHERE & " AND NAME_CODE = " & (vCode)
'
'    strRYAK = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWHERE), "")
'
'    If strRYAK = "" Then
'        MsgBox strMSGデータなし, vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
'
'    If Trim(Me.CNTA_USE) <> "" Then
'        If Trim(Len(Me.CNTA_USE)) > 1 Then
'            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
' 2005/11/26 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :CNTA_USE_KeyPressイベント
'        MODULE_ID        :CNTA_USE_KeyPress
'        CREATE_DATE      :2005/10/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CNTA_USE_KeyPress(KeyAscii As Integer)

    ' 入力を許可しない
    KeyAscii = 0: Exit Sub

'    ' 桁数チェック
'    Call CmKeyPress_For_Num1(Me.CNTA_USE, KeyAscii, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_削除_Clickイベント
'        MODULE_ID        :cmd_削除_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_Click()

    Dim intAns  As Integer
    Dim intRtn  As Integer  ' 2005/12/27 ADD T.SUZUKI

    On Error GoTo Err_cmd_削除_Click

    ' レコードが確定されていない場合には、処理を行わない
    If Nz(Me.CNTA_Number, "") = "" Then Exit Sub

    ' 削除チェック
    intAns = CHK_DEL_CNTA
    If intAns = -1 Then
        MsgBox "このコンテナは削除できません。", vbInformation + vbOKOnly, "確認"
        GoTo Exit_cmd_削除_Click
    End If

' 2005/12/27 ADD T.SUZUKI Start
    ' コンテナ削除チェック
    intRtn = Me.Parent.fncKONT_DelCheck(CLng(Me.Parent.cmb_ヤードコード), CLng(Me.txt_コンテナ番号))
    Select Case intRtn
        Case -1
            ' 削除不可 OR システムエラー
            GoTo Exit_cmd_削除_Click
        Case 0
            ' 削除可能
    End Select
' 2005/12/27 ADD T.SUZUKI End

    ' 削除前の確認メッセージ出力
    intAns = MsgBox("このコンテナを削除します。" & vbCrLf _
        & "よろしいですか？", vbCritical + vbYesNo + vbDefaultButton2, strMsg_kaku)
    If intAns <> vbYes Then
        Me.txt_コンテナ番号.SetFocus
        GoTo Exit_cmd_削除_Click
    End If

    ' データ削除処理
    Call WRI_RTN

    doCmd.Echo False  ' 2005/11/08 ADD T.SUZUKI
    Me.Requery
    doCmd.Echo True   ' 2005/11/08 ADD T.SUZUKI
    Me.txt_コンテナ番号.SetFocus

Exit_cmd_削除_Click:
    Exit Sub

Err_cmd_削除_Click:
    MsgBox Err.Number & Space(1) & Err.Description
    Resume Exit_cmd_削除_Click
End Sub
' 2005/12/14 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :cmd_削除_Clickイベント
''        MODULE_ID        :cmd_削除_Click
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmd_削除_Click()
'
'    Dim strWHERE        As String
'    Dim Ans             As Integer
'    Dim MaxCnt          As Variant
'    Dim intCnt          As Integer
'    Dim vCode(0 To 1)   As Variant
'
'    On Error GoTo Err_cmd_削除_Click
'
'    ' レコードが確定されていない場合には、処理を行わない
'    If Nz(Me.CNTA_Number, "") = "" Then Exit Sub  ' 2005/11/29 ADD T.SUZUKI
'
'' 2005/11/29 ADD T.SUZUKI Start
'    Dim intRtn As Integer
'
'    ' 削除チェック
'    intRtn = CHK_DEL_CNTA
'    If intRtn <> 0 Then
'        If intRtn = -1 Then: MsgBox "このコンテナは削除できません。", vbInformation + vbOKOnly, "確認"
'        SendKeys "{ESC}"
'        GoTo Exit_cmd_削除_Click
'    End If
'' 2005/11/29 ADD T.SUZUKI End
'
'' 2005/11/13 ADD T.SUZUKI Start
'    ' コンテナ番号未入力時の処理
'    If (Nz(Me.txt_コンテナ番号) = "" And Nz(Me.CNTA_Number, "") <> "") Then
'        ' コンテナ番号未入力かつ、空レコードのデータ削除時
'        Call WRI_RTN2
'        doCmd.Echo False
'        Me.Requery
'        doCmd.Echo True
'        Me.txt_コンテナ番号.SetFocus
'        GoTo Exit_cmd_削除_Click
'    End If
'' 2005/11/13 ADD T.SUZUKI End
'
'    '*---- CODEの入力チェック ----*
'    If IsNull(Me.txt_コンテナ番号) Or Me.txt_コンテナ番号 = "" Then
'        'レコードを元に戻す
'        SendKeys "{ESC}"
'        GoTo Exit_cmd_削除_Click
'    End If
'
'' 2005/11/29 DEL T.SUZUKI Start
''    If CHK_DEL_CNTA = False Then
'''''''        MsgBox "このデータは削除できません。", vbOKOnly, "確認"
''        MsgBox "このコンテナは削除できません。", vbInformation + vbOKOnly, "確認"
''        GoTo Exit_cmd_削除_Click
''    End If
'' 2005/11/29 DEL T.SUZUKI End
'
'' ヤードコード、コンテナ番号格納
'    vCode(0) = Nz(Me.txt_ヤードコード, "")
'    vCode(1) = Nz(Me.txt_コンテナ番号, "")
'
'    intCnt = 0
'    For intCnt = 0 To 1
'        If vCode(intCnt) = "" Then
'            GoTo Exit_cmd_削除_Click
'        End If
'    Next
'
'' 2005/11/30 ADD T.SUZUKI Start
'    Ans = MsgBox("このコンテナを削除します。" & vbCrLf _
'        & "よろしいですか？", vbCritical + vbYesNo + vbDefaultButton2, strMsg_kaku)
'    If Ans <> vbYes Then
'        Me.txt_コンテナ番号.SetFocus
'        GoTo Exit_cmd_削除_Click
'    End If
'    Call WRI_RTN
'' 2005/11/30 ADD T.SUZUKI End
'' 2005/11/30 DEL T.SUZUKI Start
''    strWHERE = ""
''    strWHERE = strWHERE & " CNTA_STEP = 1 "
''
''    MaxCnt = DCount("CNTA_CODE", "ヤード下ＷＫ", strWHERE)   ' 2005/10/24 ADD T.SUZUKI
''    If MaxCnt = 1 Then
'''''''        Ans = MsgBox("最後のデータです。" & vbCrLf
''        Ans = MsgBox("最後のコンテナです。" & vbCrLf _
''            & "削除しても本当によろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)
''        If Ans <> vbYes Then
''            Me.txt_コンテナ番号.SetFocus
''            GoTo Exit_cmd_削除_Click                         ' 2005/11/26 ADD T.SUZUKI
'''            Exit Sub                                        ' 2005/11/26 DEL T.SUZUKI
''        End If
''        Call WRI_RTN
''    Else
'''''''        Ans = MsgBox("このデータを削除します。" & vbCrLf
''        Ans = MsgBox("このコンテナを削除します。" & vbCrLf _
''            & "よろしいですか？", vbCritical + vbYesNo + vbDefaultButton2, strMsg_kaku)
''        If Ans <> vbYes Then
''            Me.txt_コンテナ番号.SetFocus
''            GoTo Exit_cmd_削除_Click                         ' 2005/11/26 ADD T.SUZUKI
'''            Exit Sub                                        ' 2005/11/26 DEL T.SUZUKI
''        End If
''        Call WRI_RTN
''    End If
'' 2005/11/30 DEL T.SUZUKI End
'
'    doCmd.Echo False  ' 2005/11/08 ADD T.SUZUKI
'    Me.Requery
'    doCmd.Echo True   ' 2005/11/08 ADD T.SUZUKI
'    Me.txt_コンテナ番号.SetFocus
'
'Exit_cmd_削除_Click:
'    Exit Sub
'
'Err_cmd_削除_Click:
'    MsgBox Err.Number & Space(1) & Err.Description
'    Resume Exit_cmd_削除_Click
'End Sub
' 2005/12/14 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :DB更新処理
'        MODULE_ID        :WRI_RTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function WRI_RTN() As Boolean

    On Error GoTo err_rtn

    WRI_RTN = False

    doCmd.Hourglass True
    doCmd.SetWarnings False

    Set WRI_ObjWS = DBEngine(0)
    Set WRI_objDb = CurrentDb

    'トランザクション開始
    WRI_ObjWS.BeginTrans

' 2005/11/26 ADD T.SUZUKI Start
    blnCheckSkip = True

    '入力中のレコードをいったん保存する
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
' 2005/11/26 ADD T.SUZUKI End

    '処理モード毎にSQLの発行
    If WRI_CNTA_MAST_RTN = False Then
        GoTo Update_Err
    End If

    'コミット完了
    WRI_ObjWS.CommitTrans
    WRI_RTN = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set WRI_objDb = Nothing
    Set WRI_ObjWS = Nothing
    blnCheckSkip = False     ' 2005/11/26 ADD T.SUZUKI
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Update_Err:
    WRI_ObjWS.Rollback
    MsgBox strmsg更新エラー, vbCritical, strMsg_kaku
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB更新処理
'        MODULE_ID        :WRI_CNTA_MAST_RTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function WRI_CNTA_MAST_RTN() As Boolean

'    Dim vCode(0 To 1)   As Variant   ' 2005/12/14 DEL T.SUZUKI
'    Dim intCnt          As Integer   ' 2005/12/14 DEL T.SUZUKI
    
    Dim strSQL          As String
    Dim objRsWk         As Recordset  ' 2005/10/24 ADD T.SUZUKI

    On Error GoTo err_rtn

    WRI_CNTA_MAST_RTN = False

' 2005/12/14 DEL T.SUZUKI Start
'    vCode(0) = Nz(Me.txt_ヤードコード, "")
'    vCode(1) = Nz(Me.txt_コンテナ番号, "")
'
'    intCnt = 0
'    For intCnt = 0 To 1
'        If vCode(intCnt) = "" Then
'            GoTo err_rtn
'        End If
'    Next
' 2005/12/14 DEL T.SUZUKI End

' 2005/10/24 ADD T.SUZUKI Start
    ' ヤード下ＷＫのデータ削除
    strSQL = ""
    strSQL = "SELECT * FROM ヤード下ＷＫ "
' 2005/12/14 ADD T.SUZUKI Start
    strSQL = strSQL & " WHERE CNTA_Number = " & Me.CNTA_Number  ' 連番
' 2005/12/14 ADD T.SUZUKI End
' 2005/12/14 DEL T.SUZUKI Start
'    strSQL = strSQL & " WHERE CNTA_CODE = " & (vCode(0))
'    strSQL = strSQL & "   AND CNTA_NO = " & (vCode(1))
'    strSQL = strSQL & "   AND CNTA_STEP = 1 "
' 2005/12/14 DEL T.SUZUKI End
    Set objRsWk = WRI_objDb.OpenRecordset(strSQL)

    With objRsWk
        Do Until .EOF
            .Edit
            .Delete

            .MoveNext
        Loop
    End With
    objRsWk.Close
    Set objRsWk = Nothing
' 2005/10/24 ADD T.SUZUKI End
' 2005/10/24 DEL T.SUZUKI Start
'    strSQL = ""
'    strSQL = strSQL & " DELETE FROM ヤード下ＷＫ"
'    strSQL = strSQL & " WHERE CNTA_CODE = " & (vCode(0))
'    strSQL = strSQL & " AND CNTA_NO = " & (vCode(1))
'    strSQL = strSQL & " AND CNTA_STEP = 1 "
'
'    WRI_objDb.Execute strSQL
' 2005/10/24 DEL T.SUZUKI End

    WRI_CNTA_MAST_RTN = True

Exit_rtn:
    If Not objRsWk Is Nothing Then objRsWk.Close: Set objRsWk = Nothing        ' 2005/10/24 ADD T.SUZUKI
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :DB更新処理
'        MODULE_ID        :WRI_CNTA_MAST_RTN2
'        CREATE_DATE      :2005/11/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function WRI_CNTA_MAST_RTN2() As Boolean

    Dim vCode           As Variant
    Dim strSQL          As String
    Dim objRsWk         As Recordset

    On Error GoTo err_rtn
    WRI_CNTA_MAST_RTN2 = False

    ' AutoNumber型のデータ取得
    vCode = Me.CNTA_Number

    If Nz(vCode, "") = "" Then
        WRI_CNTA_MAST_RTN2 = True
        GoTo Exit_rtn
    End If

    ' ヤード下ＷＫのデータ削除
    strSQL = ""
    strSQL = "SELECT * FROM ヤード下ＷＫ "
    strSQL = strSQL & " WHERE CNTA_Number = " & (vCode)
    strSQL = strSQL & "   AND CNTA_STEP = 1 "
    Set objRsWk = WRI_objDb.OpenRecordset(strSQL)

    With objRsWk
        Do Until .EOF
            .Edit
            .Delete

            .MoveNext
        Loop
    End With
    objRsWk.Close
    Set objRsWk = Nothing

    WRI_CNTA_MAST_RTN2 = True

Exit_rtn:
    If Not objRsWk Is Nothing Then objRsWk.Close: Set objRsWk = Nothing        ' 2005/10/24 ADD T.SUZUKI
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :コンテナ存在チェック
'        MODULE_ID        :CHK_CNTA_EXIST
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_CNTA_EXIST() As Boolean

    Dim strWhere        As String
    Dim ans             As Integer
    Dim MaxCnt          As Variant
    Dim intCnt          As Integer
    Dim vCode(0 To 1)   As Variant
    Dim CNTACnt          As Variant

    CHK_CNTA_EXIST = False
    vCode(0) = Nz(Me.txt_ヤードコード, "")
    vCode(1) = Nz(Me.txt_コンテナ番号, "")

    intCnt = 0
    For intCnt = 0 To 1
        If vCode(intCnt) = "" Then
            Exit Function
        End If
    Next

' 2005/10/24 ADD T.SUZUKI Start
    '存在チェック(上段)
    strWhere = ""
    strWhere = strWhere & " CNTA_CODE = " & (vCode(0))
    strWhere = strWhere & " AND CNTA_NO = " & (vCode(1))
    strWhere = strWhere & " AND CNTA_STEP = 0"

    MaxCnt = DCount("CNTA_NO", "ヤード上ＷＫ", strWhere)
    If MaxCnt <> 0 Then
        CHK_CNTA_EXIST = True
        Exit Function
    End If
' 2005/10/24 ADD T.SUZUKI End

    '存在チェック(下段)
    strWhere = ""
    strWhere = strWhere & " CNTA_CODE = " & (vCode(0))
    strWhere = strWhere & " AND CNTA_NO = " & (vCode(1))
    strWhere = strWhere & " AND CNTA_STEP = 1"

    MaxCnt = DCount("CNTA_NO", "ヤード下ＷＫ", strWhere)
    If MaxCnt <> 0 Then
        CHK_CNTA_EXIST = True
        Exit Function
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :削除チェック
'        MODULE_ID        :CHK_DEL_CNTA
'        Return           : 0 = 削除チェックOK
'                           1 = コンテナ番号未入力
'                          -1 = 削除チェックNG
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_DEL_CNTA(Optional intKbn As Integer = 0) As Integer  ' 2005/10/24 ADD T.SUZUKI
'Private Function CHK_DEL_CNTA() As Boolean                               ' 2005/10/24 DEL T.SUZUKI

    Dim strWhere        As String
    Dim varCARG_AGRE    As Variant   ' 2005/10/24 ADD T.SUZUKI
    Dim intCnt          As Integer
    Dim vCode(0 To 1)   As Variant
    Dim varYOUKT_UKNO   As Variant   ' 2005/10/24 ADD T.SUZUKI
    Dim varREPR_TRAN    As Variant   ' 2009/04/01 ADD KIHARA

    CHK_DEL_CNTA = -1
'    CHK_DEL_CNTA = False

    vCode(0) = Nz(Me.txt_ヤードコード, "")
' 2005/11/04 REP T.SUZUKI Start
    If intKbn = 0 Then
        ' 削除時
' 2005/11/29 REP T.SUZUKI Start
        If Nz(Me.txt_コンテナ番号, "") = "" Then
            vCode(1) = Nz(Me.txt_コンテナ番号_Old, "")
        Else
            vCode(1) = Nz(Me.txt_コンテナ番号, "")
        End If
'        vCode(1) = Nz(Me.txt_コンテナ番号, "")
' 2005/11/29 REP T.SUZUKI End
    Else
        ' 変更時
        vCode(1) = Nz(Me.txt_コンテナ番号_Old, "")   ' 2005/11/29 ADD T.SUZUKI
'        vCode(1) = Nz(Me.txt_コンテナ番号.Tag, "")  ' 2005/11/29 DEL T.SUZUKI
    End If
'    vCode(1) = Nz(Me.txt_コンテナ番号, "")
' 2005/11/04 REP T.SUZUKI End

    intCnt = 0
    For intCnt = 0 To 1
' 2005/11/29 REP T.SUZUKI Start
        If vCode(intCnt) = "" Then
            ' ヤードコードまたはコンテナ番号未入力時
            CHK_DEL_CNTA = 1
            Exit Function
        End If
'        If vCode(intCnt) = "" Then Exit Function
' 2005/11/29 REP T.SUZUKI End
    Next

    '存在チェック
    '●コンテナ契約ファイル
    strWhere = ""
    strWhere = strWhere & " CNTA_CODE = " & (vCode(0))
    strWhere = strWhere & " AND CNTA_NO = " & (vCode(1))

' 2005/10/24 ADD T.SUZUKI Start
    ' 「契約状態(名称)」
    varCARG_AGRE = DLookup("AGRE_NAME", "ヤード下ＷＫ", strWhere)

    If Nz(varCARG_AGRE, "") <> "" Then Exit Function

    '●予約受付トラン
    strWhere = ""
    strWhere = strWhere & " YOUKT_YCODE = " & (vCode(0))
    strWhere = strWhere & " AND YOUKT_NO = " & (vCode(1))

    ' 「予約受付番号」
    varYOUKT_UKNO = DLookup("YOUKT_UKNO", "dbo_YOUK_TRAN", strWhere)

    If Nz(varYOUKT_UKNO, "") <> "" Then Exit Function
' 2005/10/24 ADD T.SUZUKI End
' 2005/10/24 DEL T.SUZUKI Start
'    MaxCnt = DLookup("CNTA_AGRE", "ヤード下ＷＫ", strWhere)
'    If MaxCnt <> 0 Then
'        Exit Function
'    End If
' 2005/10/24 DEL T.SUZUKI End
' 2009/04/01 INSERT KIHARA Start
    '●修繕トラン
    strWhere = ""
    strWhere = strWhere & " CNTA_CODE = " & (vCode(0))
    strWhere = strWhere & " AND CNTA_NO = " & (vCode(1))

    ' 「修繕状態(名称)」
    varREPR_TRAN = DLookup("REPR_NAME", "ヤード下ＷＫ", strWhere)
    If Nz(varREPR_TRAN, "") <> "" Then Exit Function
' 2009/04/01 INSERT KIHARA End

    CHK_DEL_CNTA = 0
'    CHK_DEL_CNTA = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subCtlTab
'        CREATE_DATE      :2005/10/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subCtlTab(intKbn As Integer)
    ' INSERT START 2018/09/05 K.WADA
    Dim blnVisible  As Boolean
    blnVisible = False
    If intKbn = 0 Then
        blnVisible = True
    End If
    Me.lbl_階数.Visible = blnVisible
    Me.CNTA_FLOOR.Visible = blnVisible
    ' INSERT END 2018/09/05 K.WADA
    Me.tab_DownHeadder.Pages(intKbn).SetFocus
    Me.tab_Down.Pages(intKbn).SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :バイク可否チェックボックスの値設定
'        MODULE_ID        :subBikeOkCheckBox
'        CREATE_DATE      :2011/03/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subBikeOkCheckBox()
    
    Dim iCnt   As Integer
    Dim sWork  As String
    Dim tmp    As Variant
    Dim i      As Integer
    
    Select Case Nz(Me.CNTA_USAGE.VALUE, "99")
    Case pcstrUSAGE_コンテナ
        '用途＝コンテナ、下段、４帖以下、差額なし ならばバイク可
        '2016/12/06 M.HONDA UPD
        'chk_CNTA_BIKE_OK.VALUE = (Nz(txt_段区分) = pcstrSTEP_下段 And Nz(CNTA_SIZE, 99) <= 4 And Nz(CNTA_PRICE_DIFF, 0) = 0)
        '0.7帖の場合にはデフォルトのチェックを外す
        If (Nz(txt_段区分) = pcstrSTEP_下段 And Nz(CNTA_SIZE, 99) <= 4 And Nz(CNTA_SIZE, 99) > 0.7 And Nz(CNTA_PRICE_DIFF, 0) = 0) Then
            chk_CNTA_BIKE_OK.VALUE = True
        Else
            chk_CNTA_BIKE_OK.VALUE = False
        End If
        '2016/12/06 M.HONDA UPD
    
    Case pcstrUSAGE_バイク, pcstrUSAGE_バイク屋外, pcstrUSAGE_バイク屋内, pcstrUSAGE_バイクコンテナ, pcstrUSAGE_コンテナバイク
        '用途＝バイク関係 ならばバイク可
        chk_CNTA_BIKE_OK.VALUE = True
    
    Case Else
        chk_CNTA_BIKE_OK.VALUE = False
    End Select

    iCnt = 0
    '奥行・幅・高さを自動的にセットする。
    For iCnt = 0 To UBound(strSize())
        If CNTA_SIZE = strSize(iCnt) Then
            sWork = strBox(iCnt)
            tmp = Split(sWork, ",")
            CNTA_WIDTH = tmp(1)
            CNTA_DEPTH = tmp(0)
            CNTA_HEIGHT = tmp(2)
            Exit For
        End If
    Next

End Sub

'****************************  ended or program ********************************




