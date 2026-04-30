VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS440_ヤードマスタサブ価格表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　倉庫管理システム
'**********************************************
'   プログラム名　　：　ヤードマスタサブ価格表
'   プログラムＩＤ　：
'   作　成　日　　　：  2005/11/18
'   作　成　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.1
'**********************************************
'**********************************************
'修正履歴
'   修　正　日　　　：2006/03/18
'   修　正　者　　　：T.SUZUKI
'   修　正　内　容　：雑費で「無し」を選らんだ場合、「0」以上の雑費金額が入力できない
'   Ver             ：0.2
'**********************************************
'修正履歴
'   修　正　日　　　：2006/06/24
'   修　正　者　　　：tajima
'   修　正　内　容　：修繕呼出対応
'   Ver             ：0.3
'**********************************************
'修正履歴
'   修　正　日　　　：2007/05/25
'   修　正　者　　　：tajima
'   修　正　内　容　： デザインのみ変更
'                   ： 1.バイク対応（30:バイク全般は選択肢に含まず）
'   Ver             ：0.4
'**********************************************
'修正履歴
'   修　正　日　　　：2008/10/15
'   修　正　者　　　：iizuka
'   修　正　内　容  ：価格表とコンテナマスタに
'                     「割引き対象外」を追加
'   Ver             ：0.5
'**********************************************
'修正履歴
'   修　正　日　　　：2008/11/14
'   修　正　者　　　：K.KINEBUCHI
'   修　正　内　容  ：未設定行項目に対し「割引き対象外」をチェックした時の不具合を修正する
'   Ver            ：0.6
'**********************************************
'修正履歴
'   修　正　日　　　：2012/06/12
'   修　正　者　　　：M.HONDA
'   修　正　内　容  ：ｺﾝﾃﾅ(ﾊﾞｲｸ)選択時に初回雑費に配送料をｾｯﾄする
'   Ver            ：0.7
'**********************************************
'修正履歴
'   修　正　日　　　：　2012/12/26
'   修　正　者　　　：　M.HONDA
'   Ver             ：　0.8
'   修　正　内　容  ：  用途 : バイク以外が選択選択された際に
'                       毎月雑費に保険料相当額を自動的にセット。
'
'**********************************************
'修正履歴
'   修　正　日　　　：　2013/03/17
'   修　正　者　　　：　M.HONDA
'   Ver             ：　0.9
'   修　正　内　容  ：  保証加入費の価格を変更
'
'**********************************************
'修正履歴
'   修　正　日　　　：　2018/08/20
'   修　正　者　　　：　K.WADA
'   Ver             ：　1.0
'   修　正　内　容  ：  事務手数料登録機能、階数入力機能の追加
'
'**********************************************
'修正履歴
'   修　正　日　　　：　2019/07/08
'   修　正　者　　　：　M.HONDA
'   Ver             ：　1.1
'   修　正　内　容  ：  賃料が変更されたら事務手数料も変更するように修正
'**********************************************

Option Compare Database
Option Explicit

Private Const strMsg_kaku           As String = "確認"
Private Const strMsg_err            As String = "エラー"
Private Const strmsg重複            As String = "既に同じ用途/同じ実帖が設定されています。"
Private Const strMSG入力不正        As String = "必須入力です。"
Private Const strMSG桁数オーバー    As String = "桁数オーバーしています。"  ' 2005/12/15 ADD T.SUZUKI

Private Const pcstrfmtSTANDARD      As String = "Standard"  ' 書式：標準
Private Const pcstrfmtFIXED         As String = "Fixed"     ' 書式：固定

Private Const 雑費区分_無し = 0

'Trueの場合は更新時の必須項目チェック等は行わない
Dim blnCheckSkip                    As Boolean

Private Const pstrName_Up           As String = "FKS440_ヤードマスタサブ価格表上"
Private Const pstrName_Down         As String = "FKS440_ヤードマスタサブ価格表下"

Private Const pcsglSizeMaxValue     As Single = 999.99   ' 実帖の入力最大値      ' 2005/12/15 ADD T.SUZUKI
Private Const pcsglMoneyMaxValue    As Long = 999999999  ' 金額項目の入力最大値  ' 2005/12/17 ADD T.SUZUKI
Private Const pcintMINUS            As Integer = 45      ' マイナス符号          ' 2005/12/15 ADD T.SUZUKI

Private Const pcstrZAPPI_ErrMsg_F   As String = "初回雑費１で「無し」を選んだ場合「0」以上の金額は入力できません。"
Private Const pcstrZAPPI2_ErrMsg_F   As String = "初回雑費２で「無し」を選んだ場合「0」以上の金額は入力できません。"    ' 2018/08/20 ADD K.WADA
Private Const pcstrZAPPI_ErrMsg_M   As String = "毎月雑費１で「無し」を選んだ場合「0」以上の金額は入力できません。"

'==============================================================================*
'
'        MODULE_NAME      :cmb_初回雑費区分_AfterUpdate イベント
'        MODULE_ID        :cmb_初回雑費区分_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_初回雑費区分_AfterUpdate()

    Me.Parent.blnDataUpdate = True

' 2006/03/18 ADD T.SUZUKI Start
    If Me.cmb_初回雑費区分 = 0 And Nz(Me.txt_初回雑費, 0) > 0 Then
        MsgBox pcstrZAPPI_ErrMsg_F, vbInformation, strMsg_err
        Me.txt_初回雑費.SetFocus
    End If
' 2006/03/18 ADD T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_初回雑費区分_KeyPress イベント
'        MODULE_ID        :cmb_初回雑費区分_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_初回雑費区分_KeyPress(KeyAscii As Integer)
    ' 入力を許可しない
    KeyAscii = 0: Exit Sub
End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_初回雑費区分2_AfterUpdate イベント
'        MODULE_ID        :cmb_初回雑費区分2_AfterUpdate
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_初回雑費区分2_AfterUpdate()
    Me.Parent.blnDataUpdate = True
    If Me.cmb_初回雑費区分2 = 0 And Nz(Me.txt_初回雑費2, 0) > 0 Then
        MsgBox pcstrZAPPI2_ErrMsg_F, vbInformation, strMsg_err
        Me.txt_初回雑費2.SetFocus
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_初回雑費区分2_KeyPress イベント
'        MODULE_ID        :cmb_初回雑費区分2_KeyPress
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_初回雑費区分2_KeyPress(KeyAscii As Integer)
    ' 入力を許可しない
    KeyAscii = 0: Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_毎月雑費区分_AfterUpdate イベント
'        MODULE_ID        :cmb_毎月雑費区分_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_毎月雑費区分_AfterUpdate()

    Me.Parent.blnDataUpdate = True

' 2006/03/18 ADD T.SUZUKI Start
    If Me.cmb_毎月雑費区分 = 0 And Nz(Me.txt_毎月雑費, 0) > 0 Then
        MsgBox pcstrZAPPI_ErrMsg_M, vbInformation, strMsg_err
        Me.txt_毎月雑費.SetFocus
    End If
' 2006/03/18 ADD T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_毎月雑費区分_KeyPress イベント
'        MODULE_ID        :cmb_毎月雑費区分_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_毎月雑費区分_KeyPress(KeyAscii As Integer)
    ' 入力を許可しない
    KeyAscii = 0: Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_BeforeUpdate イベント
'        MODULE_ID        :Form_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_rtn

    If blnCheckSkip Then
        Exit Sub
    End If
    
    '必須チェック
    If Not RequireCheck() Then
        Cancel = True
        Exit Sub
    End If
    
    '同一用途／サイズが存在してる？
    If CHK_PRIC_EXIST() Then
        MsgBox strmsg重複, , strMsg_err
        Cancel = True
        Exit Sub
    End If
    
    '雑費関係初期値
    If Nz(cmb_初回雑費区分, "") = "" Then
        cmb_初回雑費区分 = 雑費区分_無し
    End If

    ' INSERT START 2018/08/20 K.WADA
    If Nz(cmb_初回雑費区分2, "") = "" Then
        cmb_初回雑費区分2 = 雑費区分_無し
    End If
    ' INSERT END 2018/08/20 K.WADA

    If Nz(cmb_毎月雑費区分, "") = "" Then
        cmb_毎月雑費区分 = 雑費区分_無し
    End If

    If Nz(txt_初回雑費, "") = "" Then
        txt_初回雑費 = 0
    End If

    ' INSERT START 2018/08/20 K.WADA
    If Nz(txt_初回雑費2, "") = "" Then
        txt_初回雑費2 = 0
    End If
    ' INSERT END 2018/08/20 K.WADA

    If Nz(txt_毎月雑費, "") = "" Then
        txt_毎月雑費 = 0
    End If
    
    Exit Sub

err_rtn:
    MsgBox Err.Description, , strMsg_err
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Loadイベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    blnCheckSkip = False
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

    Dim strName As String

    On Error Resume Next  ' 2005/12/15 ADD T.SUZUKI

    Call CmReprMod.OpenFVS90n(KeyCode, Me.Parent.cmb_ヤードコード)  ' 2006/06/24 Add tajima
   
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

' 2005/11/21 ADD T.SUZUKI Start
        Case vbKeyF5
            strName = Me.Parent.ActiveControl.NAME
            Select Case strName
                Case pstrName_Up
                    Me.Parent.FKS440_ヤードマスタサブ価格表下.SetFocus

                Case pstrName_Down
                    Me.Parent.FKS440_ヤードマスタサブ上.SetFocus
            End Select
' 2005/11/21 ADD T.SUZUKI End

        Case vbKeyF3, vbKeyF6 To vbKeyF12
            KeyCode = 0

        Case Else
            Exit Sub
    End Select
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
        Case 2113  '定型入力エラー（アプリケーションエラー）
            strCtrlName = Me.ActiveControl.NAME  ' アクティブなコントロール名を取得
            Select Case strCtrlName
                Case Me.txt_価格.NAME
                    ' *** 賃料 ***
                    If Me.Parent.fncTxtInputCheck(Me.txt_価格, pcsglMoneyMaxValue, 9, 1) = False Then
                        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
                        SendKeys "{ESC}"
                    End If

                Case Me.txt_初回雑費.NAME
                    ' *** 初回雑費 ***
                    If Me.Parent.fncTxtInputCheck(Me.txt_初回雑費, pcsglMoneyMaxValue, 9, 1) = False Then
                        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
                        SendKeys "{ESC}"
                    End If
                
                Case Me.txt_初回雑費2.NAME
                    ' *** 初回雑費2 ***
                    If Me.Parent.fncTxtInputCheck(Me.txt_初回雑費2, pcsglMoneyMaxValue, 9, 1) = False Then
                        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
                        SendKeys "{ESC}"
                    End If

                Case Me.txt_毎月雑費.NAME
                    ' *** 毎月雑費 ***
                    If Me.Parent.fncTxtInputCheck(Me.txt_毎月雑費, pcsglMoneyMaxValue, 9, 1) = False Then
                        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
                        SendKeys "{ESC}"
                    End If
            End Select

            Response = acDataErrContinue
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_AfterUpdateイベント
'        MODULE_ID        :cmb_用途_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_AfterUpdate()
    
    '' INS 2012/06/13 M.HONDA START
    '' ｺﾝﾃﾅ(ﾊﾞｲｸ)の場合には初回雑費に配送料をｾｯﾄする
    If cmb_用途.VALUE = "33" Then
        cmb_初回雑費区分 = 95
    Else
        cmb_初回雑費区分 = 0
    End If
    '' INS 2012/06/13 M.HONDA END
    
    ' INSERT START 2018/08/20 K.WADA
    cmb_初回雑費区分2 = 63
    txt_初回雑費2 = ""
    ' INSERT END 2018/08/20 K.WADA
    
    ' DELETE START 2018/08/20 K.WADA
    '' INS 2012/12/26 M.HONDA START
    '' ﾊﾞｲｸ・ﾊﾞｰﾁｬﾙｵﾌｨｽ以外の場合には毎月雑費に保険料相当額をセット
'    If cmb_用途.VALUE = "0" Or cmb_用途.VALUE = "1" Or cmb_用途.VALUE = "2" Or cmb_用途.VALUE = "4" Or cmb_用途.VALUE = "5" Or cmb_用途.VALUE = "6" _
'       Or cmb_用途.VALUE = "39" Or cmb_用途.VALUE = "98" Or cmb_用途.VALUE = "99" Then
'        cmb_毎月雑費区分 = 68
'        '2014/03/17 M.HONDA UPD
'        'txt_毎月雑費 = 450
'        txt_毎月雑費 = 480
'        '2014/03/17 M.HONDA UPD
'    Else
'        cmb_毎月雑費区分 = 0
'        txt_毎月雑費 = 0
'    End If
    '' INS 2012/12/26 M.HONDA END
    ' DELETE END 2018/08/20 K.WADA
    
    ' INSERT START 2018/08/20 K.WADA
    cmb_毎月雑費区分 = 68
    txt_毎月雑費 = fncGetInitZappi(68)
    ' INSERT END 2018/08/20 K.WADA
    
    Me.Parent.blnDataUpdate = True
    
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :雑費コードより雑費初期値を取得する
'        MODULE_ID        :fncGetZappi
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetInitZappi(intZappiCode As Integer) As String
    Dim strWhere            As String
    Dim strNameValueFrom    As String
    
    fncGetInitZappi = 0

    strWhere = ""
    strWhere = strWhere & " NAME_ID = '096'"
    strWhere = strWhere & " AND NAME_CODE = " & intZappiCode
    strNameValueFrom = Nz(DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", strWhere), "")
    If CLng(strNameValueFrom) = 999999999 Then
        strNameValueFrom = Nz(Me.PRIC_PRICE, "")
    End If
    fncGetInitZappi = strNameValueFrom

End Function
'==============================================================================*
'
'        MODULE_NAME      :cmb_用途_KeyPressイベント
'        MODULE_ID        :cmb_用途_KeyPress
'        CREATE_DATE      :
'

'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_用途_KeyPress(KeyAscii As Integer)
    ' 入力を許可しない
    KeyAscii = 0: Exit Sub
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
On Error GoTo err_rtn
    
    Dim objDb                           As Database

    If Nz(txt_連番, "") = "" Then
        Exit Sub
    End If

    If MsgBox("削除します、よろしいですか？", vbCritical + vbYesNo + vbDefaultButton2, strMsg_kaku) = vbNo Then
'    If MsgBox("削除します。よろしいですか？", vbYesNo, strMsg_kaku) = vbNo Then
        Exit Sub
    End If
    
    blnCheckSkip = True
    
    '入力中のレコードをいったん保存する
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
    On Error GoTo err_rtn
    
    blnCheckSkip = False
    
    'ワークテーブルから削除
    Set objDb = CurrentDb
    objDb.Execute "DELETE FROM FKS440_W02 WHERE PRIC_SEQ = " & txt_連番, dbFailOnError

    Me.Requery

Exit_rtn:
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    Exit Sub

err_rtn:
    MsgBox Err.Description, , strMsg_err
    Resume Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :価格表存在チェック
'        MODULE_ID        :CHK_PRIC_EXIST
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_PRIC_EXIST() As Boolean

    Dim strWhere        As String
    Dim valYcode          As Variant

    CHK_PRIC_EXIST = False
    
    '存在チェック
    strWhere = ""
    strWhere = strWhere & " PRIC_YCODE= " & Val(txt_ヤードコード)
    strWhere = strWhere & " AND PRIC_USAGE = " & Val(cmb_用途)
    strWhere = strWhere & " AND PRIC_SIZE = " & Val(txt_実帖)
    strWhere = strWhere & " AND PRIC_STEP = " & Val(txt_段区分)
    strWhere = strWhere & " AND PRIC_SEQ <> " & Val(txt_連番)

    '存在する場合はTrueを返却する
    valYcode = DLookup("PRIC_YCODE", "FKS440_W02", strWhere)
    If Nz(valYcode, "") <> "" Then
        CHK_PRIC_EXIST = True
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :txt_価格_AfterUpdate イベント
'        MODULE_ID        :txt_価格_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_価格_AfterUpdate()
    Me.Parent.blnDataUpdate = True
    
    ' 2019/0708 M.HONDA UPD
    ' INSERT START 2018/08/20 K.WADA
    'If cmb_初回雑費区分2 = 63 Then  ' DELETE 2018/08/28 K.WADA
    'If cmb_初回雑費区分2 = 63 And Nz(txt_初回雑費2, "") = "" Then  ' INSERT 2018/08/28 K.WADA
    If cmb_初回雑費区分2 = 63 Then
        txt_初回雑費2 = fncGetInitZappi(63)
    End If
    ' INSERT END 2018/08/20 K.WADA
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_価格_BeforeUpdate イベント
'        MODULE_ID        :txt_価格_BeforeUpdate
'        CREATE_DATE      :2005/12/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_価格_BeforeUpdate(Cancel As Integer)

    If Me.Parent.fncTxtInputCheck(Me.txt_価格, pcsglMoneyMaxValue, 9) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_価格_GotFocus イベント
'        MODULE_ID        :txt_価格_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_価格_GotFocus()
    Me.txt_価格.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_価格_KeyPress イベント
'        MODULE_ID        :txt_価格_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_価格_KeyPress(KeyAscii As Integer)
    ' 桁数チェック
    Call CmKeyPress_For_Num1(Me.txt_価格, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_価格_LostFocus イベント
'        MODULE_ID        :txt_価格_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_価格_LostFocus()
    Me.txt_価格.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_実帖_AfterUpdate イベント
'        MODULE_ID        :txt_実帖_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_実帖_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_実帖_BeforeUpdate イベント
'        MODULE_ID        :txt_実帖_BeforeUpdate
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_実帖_BeforeUpdate(Cancel As Integer)

    If Me.Parent.fncTxtInputCheck(Me.txt_実帖, pcsglSizeMaxValue, 5) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_実帖_KeyPress イベント
'        MODULE_ID        :txt_実帖_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_実帖_KeyPress(KeyAscii As Integer)
    On Error Resume Next

' 2005/12/15 ADD T.SUZUKI Start
    If KeyAscii = pcintMINUS Then KeyAscii = 0: Exit Sub

    If Me.Parent.fncTxtInputCheck(Me.txt_実帖, pcsglSizeMaxValue, 5) = False Then
        KeyAscii = 0: Exit Sub
    End If
' 2005/12/15 ADD T.SUZUKI End

    ' 桁数チェック
    Call ETC_CUREN_RTN(Me.txt_実帖, 3, 2, KeyAscii)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費_AfterUpdate イベント
'        MODULE_ID        :txt_初回雑費_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費_BeforeUpdate イベント
'        MODULE_ID        :txt_初回雑費_BeforeUpdate
'        CREATE_DATE      :2005/12/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費_BeforeUpdate(Cancel As Integer)

' 2006/03/18 ADD T.SUZUKI Start
'    MsgBox Me.cmb_初回雑費区分
    If Me.cmb_初回雑費区分 = 0 And Nz(Me.txt_初回雑費, 0) > 0 Then
        MsgBox pcstrZAPPI_ErrMsg_F, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
' 2006/03/18 ADD T.SUZUKI End

    If Me.Parent.fncTxtInputCheck(Me.txt_初回雑費, pcsglMoneyMaxValue, 9) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費_GotFocus イベント
'        MODULE_ID        :txt_初回雑費_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費_GotFocus()
    Me.txt_初回雑費.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費_KeyPress イベント
'        MODULE_ID        :txt_初回雑費_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    ' 桁数チェック
    Call CmKeyPress_For_Num1(Me.txt_初回雑費, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費_LostFocus イベント
'        MODULE_ID        :txt_初回雑費_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費_LostFocus()
    Me.txt_初回雑費.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費2_AfterUpdate イベント
'        MODULE_ID        :txt_初回雑費2_AfterUpdate
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費2_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費2_BeforeUpdate イベント
'        MODULE_ID        :txt_初回雑費2_BeforeUpdate
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費2_BeforeUpdate(Cancel As Integer)
    If Me.cmb_初回雑費区分2 = 0 And Nz(Me.txt_初回雑費2, 0) > 0 Then
        MsgBox pcstrZAPPI2_ErrMsg_F, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
    If Me.Parent.fncTxtInputCheck(Me.txt_初回雑費2, pcsglMoneyMaxValue, 9) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費2_GotFocus イベント
'        MODULE_ID        :txt_初回雑費2_GotFocus
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費2_GotFocus()
    Me.txt_初回雑費2.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費2_KeyPress イベント
'        MODULE_ID        :txt_初回雑費2_KeyPress
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費2_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    ' 桁数チェック
    Call CmKeyPress_For_Num1(Me.txt_初回雑費2, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_初回雑費2_LostFocus イベント
'        MODULE_ID        :txt_初回雑費2_LostFocus
'        CREATE_DATE      :2018/08/20
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_初回雑費2_LostFocus()
    Me.txt_初回雑費2.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_毎月雑費_AfterUpdate イベント
'        MODULE_ID        :txt_毎月雑費_AfterUpdate
'        CREATE_DATE      :2005/11/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_毎月雑費_AfterUpdate()
    Me.Parent.blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_毎月雑費_BeforeUpdate イベント
'        MODULE_ID        :txt_毎月雑費_BeforeUpdate
'        CREATE_DATE      :2005/12/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_毎月雑費_BeforeUpdate(Cancel As Integer)

' 2006/03/18 ADD T.SUZUKI Start
'    MsgBox Me.cmb_初回雑費区分
    If Me.cmb_毎月雑費区分 = 0 And Nz(Me.txt_毎月雑費, 0) > 0 Then
        MsgBox pcstrZAPPI_ErrMsg_M, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
' 2006/03/18 ADD T.SUZUKI End

    If Me.Parent.fncTxtInputCheck(Me.txt_毎月雑費, pcsglMoneyMaxValue, 9) = False Then
        MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_毎月雑費_GotFocus イベント
'        MODULE_ID        :txt_毎月雑費_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_毎月雑費_GotFocus()
    Me.txt_毎月雑費.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_毎月雑費_KeyPress イベント
'        MODULE_ID        :txt_毎月雑費_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_毎月雑費_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    ' 桁数チェック
    Call CmKeyPress_For_Num1(Me.txt_毎月雑費, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_毎月雑費_LostFocus イベント
'        MODULE_ID        :txt_毎月雑費_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_毎月雑費_LostFocus()
    Me.txt_毎月雑費.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :必須チェック
'        MODULE_ID        :RequireCheck
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function RequireCheck() As Boolean

    RequireCheck = False
    
    If Nz(cmb_用途) = "" Then
        MsgBox strMSG入力不正, vbInformation, strMsg_err
'        MsgBox strMSG入力不正, , strMsg_err
        cmb_用途.SetFocus
        Exit Function
    End If

    If Nz(txt_実帖) = "" Then
        MsgBox strMSG入力不正, vbInformation, strMsg_err
'        MsgBox strMSG入力不正, , strMsg_err
        txt_実帖.SetFocus
        Exit Function
    End If
    
    If Nz(txt_価格) = "" Then
        MsgBox strMSG入力不正, vbInformation, strMsg_err
'        MsgBox strMSG入力不正, , strMsg_err
        txt_価格.SetFocus
        Exit Function
    End If

    RequireCheck = True
End Function

' 2008/10/15 ADD iizuka Start
'==============================================================================*
'
'        MODULE_NAME      :chk_割引対象外_AfterUpdate イベント
'        MODULE_ID        :chk_割引対象外_AfterUpdate
'        CREATE_DATE      :2008/10/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_割引対象外_AfterUpdate()
    On Error GoTo ErrorHandler
    
    If (Nz(cmb_用途) <> "" And Nz(txt_実帖) <> "") Then                 'INSERT 2008/11/14 K.KINEBUCHI
    
        If chk_割引対象外.VALUE Then
            'チェック付けた時だけ処理を行う
    
            blnCheckSkip = True
            
            '入力中のレコードをいったん保存する
            On Error Resume Next
            doCmd.RunCommand acCmdSaveRecord
            On Error GoTo ErrorHandler
            
            blnCheckSkip = False
            
            Call Me.Parent.setNebikiOnOff(Val(txt_ヤードコード), Val(txt_段区分), _
                                          Val(cmb_用途), Val(txt_実帖.VALUE), chk_割引対象外.VALUE)
        End If
    End If                                                              'INSERT 2008/11/14 K.KINEBUCHI

    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("chk_割引対象外_AfterUpdate", True)

End Sub
' 2008/10/15 ADD iizuka End

'****************************  ended or program ********************************

