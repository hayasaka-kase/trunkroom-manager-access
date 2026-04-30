VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS221_S02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　解約ヤード詳細(ヘッダ)
'   プログラムＩＤ　：　FVS221_S02
'   作　成　日　　　：  2009/04/30
'   作　成　者　　　：  KIHARA
'   Ver             ：  0.0
'   備考            ：
'
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   Ver             ：
'   備考            ：
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private VarChangeDate   As Variant

Private blnEditMode         As Boolean  'True=編集中

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_KAKUNIN     As String = "確認"
Private Const P_ERROR       As String = "エラー"
Private Const P_CONTROL_半角        As Integer = 7       '半角文字
Private Const P_KEYNO_LENGTH        As Integer = 10      '鍵番号

Private Const P_ALL_NINE              As String = "9999999999"
'フォームキャプション
Private Const P_CAPTION_EDITING     As String = "(修正中)"
Private strFormCaption      As String   '初期表示時のフォームのキャプション

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
        Case vbKeyF1                'F1=「終了」
            KeyCode = 0
            Parent.cmd_終了.SetFocus
'▼ヤード解約情報保存時に使用▼
'        Case vbKeyF2               'F2=「保存」
'            Parent.cmd_保存.SetFocus
        Case vbKeyF2 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_解約日_BeforeUpdate
'        機能             :txt_解約日_更新前処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日_BeforeUpdate(Cancel As Integer)

    Dim vDATE        As Variant
    Dim vCompareDate As Variant

    On Error GoTo err_rtn

    ' 解約日付
    Call dateCheck(Me.txt_解約日, Cancel): If Cancel = True Then GoTo Exit_rtn

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub dateCheck(txtTarget As Variant, ByRef Cancel As Integer)

    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_解約日_AfterUpdate
'        機能             :txt_解約日_更新後処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日_AfterUpdate()

    Me.txt_解約日 = Format(VarChangeDate, "yyyy/mm/dd")
    Me.txt_解約日付 = Format(VarChangeDate, "yyyy/mm/dd") 'ご契約情報の解約日にも表示する

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_解約日_GotFocus
'        機能             :txt_解約日_フォーカス取得後
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_解約日, "")

    If vCode <> "" Then
        Me.txt_解約日 = Format(vCode, "yyyy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_解約日_KeyPress
'        機能             :txt_解約日_キー入力時
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_解約日, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_解約日_LostFocus
'        機能             :txt_解約日_フォーカス喪失後
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_解約日, "")

    If vCode <> "" Then
        Me.txt_解約日 = Format(vCode, "yyyy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_BeforeUpdateイベント
'        機能             :cmb_MoveKbn_更新前処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_BeforeUpdate(Cancel As Integer)
    
    Dim varCode As Variant
    
    If ControlCrear(Me.cmb_MoveKbn, Me.txt_MoveKbn) = False Then Exit Sub
    
    varCode = Me.cmb_MoveKbn

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo Exception

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exception
    End Select

    If Chk_NameBefore(Me.cmb_MoveKbn, "084") = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       GoTo Exception
    End If

    Exit Sub

Exception:
    Cancel = True
    SendKeys "{ESC}"

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_AfterUpdate
'        機能        　　 :cmb_MoveKbn更新後処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_AfterUpdate()
    On Error GoTo Exception

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_MoveKbn, Me.txt_MoveKbn, "084") = False Then Exit Sub

    Exit Sub

Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_KeyPress
'        機能             :cmb_MoveKbn_キー入力時
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_MoveKbn, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ControlCrear
'        機能             :ｺﾝﾎﾞﾎﾞｯｸｽの内容が空の場合、名称ﾃｷｽﾄも空にする
'        IN               :ctlCmb = ｺﾝﾎﾞﾎﾞｯｸｽ / ctlTxt = ﾃｷｽﾄ
'        OUT              :ｺﾝﾎﾞﾎﾞｯｸｽが空だったかの判断
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ControlCrear(ByRef ctlCmb As ComboBox, ByRef ctlTxt As TextBox) As Boolean
    
    If Nz(ctlCmb, "") = "" Then
        ctlTxt = ""
        Exit Function
    Else
        ControlCrear = True
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME        :CmBCheck_Num
'        機能               :数値項目のチェックを一括でする。☆★他フォームでも個々に宣言されている
'        CREATE_DATE        :2009/04/30
'
  '返り値は、
  '１：数値と判断できない時
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
    
  '引数 oText ←判定する値
  '     iLen  ←整数桁
  '     isLen ←小数桁(省略された時は、小数点桁数無し)
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, Optional ByVal isLEN As Variant) As Integer

    Dim vINT    As Variant
    Dim vsINT   As Variant
    Dim i       As Integer

    Dim dTEXT   As Double
    Dim dFROM   As Double
    Dim dTO     As Double
    Dim vRet    As Variant

    On Error GoTo err_rtn

    CmbCheck_Num = 9

    If Nz(CmNumeric(oText), "") = "" Then       '数値と判断できないのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    If CmChkCher(CStr(Nz(oText)), "A") = False Then '全角文字が入力されているのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    '小数点の値が設定されている時のチェック
    vRet = Nz(InStr(Abs(oText), "."), 0)
    
    If IsMissing(isLEN) = False Then
        If vRet <> 0 Then
            If isLEN < CmLenb(Abs(oText)) - vRet Then    '小数点桁数が設定桁数を超えているエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
            End If
        End If
    Else
        If CmLenb(Abs(CmNumeric(oText))) > ilen Then           '文字数超過のためエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
        End If
    
        If vRet <> 0 Then                         '小数点の入力があるのでエラー
          CmbCheck_Num = 2
          GoTo Exit_rtn
        End If
    End If

    '桁チェック
    i = 0
    For i = 1 To ilen: vINT = vINT & "9": Next

    If IsMissing(isLEN) = False Then
        vsINT = "."
        i = 0
        For i = 1 To isLEN
            vsINT = vsINT & "9"
        Next
    End If

    vINT = vINT & vsINT
    vINT = CmNumeric(vINT)

    dTEXT = CmNumeric(oText)
    dFROM = vINT * -1
    dTO = vINT

    If dFROM > dTEXT Or dTO < dTEXT Then        '論理範囲超過のためエラー
        CmbCheck_Num = 2
        GoTo Exit_rtn
    End If

    CmbCheck_Num = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmbCheck_Num = 9
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :Chk_NameBefore
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :ctl = 対象ｺﾝﾎﾞﾎﾞｯｸｽ / lngID = 名称ｺｰﾄﾞ
'        OUT              :ｺｰﾄﾞ名称
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_NameBefore(ByVal ctlCmb As ComboBox, ByVal strID As String) As Variant

    Dim strWhere    As String
    Dim vCode       As Variant

    'ファイルへの存在チェック
    vCode = Nz(ctlCmb, "")
    If vCode <> "" Then
        strWhere = "NAME_ID = '" & strID & "'"
        strWhere = strWhere & " AND NAME_CODE =" & Format$(vCode, "000")
        Chk_NameBefore = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :Chk_NameAfter
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :ctl = 対象ｺﾝﾎﾞﾎﾞｯｸｽ / ctlText = 名称ｾｯﾄ先ﾃｷｽﾄ / strID = 名称ｺｰﾄﾞ
'        OUT              :取得結果
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_NameAfter(ByRef ctlCmb As ComboBox, ByRef ctlText As TextBox, ByVal strID As String, _
                                Optional ctlText2 As TextBox = "", Optional ctlText3 As TextBox = "") As Boolean
    
    Dim vCode       As Variant
    Dim strWhere    As String
    
    Chk_NameAfter = False
    
    vCode = Nz(ctlCmb, "")
    
    If vCode = "" Then GoTo Chk_NameAfter_Err
    
    'ファイルへの存在チェック
    strWhere = "NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE =" & Format$(vCode, "000")
    
    ctlText = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
    
    ctlCmb = Format$(ctlCmb, "00")
    
    GoTo Chk_NameAfter_End
    
Chk_NameAfter_Err:
    Exit Function
    
Chk_NameAfter_End:
    Chk_NameAfter = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiNo_AfterUpdate
'        機能             :txtKeyKagiNo_AfterUpdate イベント
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiNo_AfterUpdate()

    Call fncAfterUpdateCommon(txtKeyKagiNo, P_KEYNO_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiNo_BeforeUpdate
'        機能             :txtKeyKagiNo_BeforeUpdate イベント
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiNo_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtKeyKagiNo, P_KEYNO_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiNo_KeyPress
'        機能             :txtKeyKagiNo_KeyPress イベント
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiNo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char(txtKeyKagiNo, KeyAscii, P_KEYNO_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncAfterUpdateCommon
'        機能             :更新後イベント時の処理
'                         :数値項目は０埋めまたは０消しした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'                         :③定数値「コントロールの種類」参照
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :年月は"yyyy/mm"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncAfterUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer) As String

    fncAfterUpdateCommon = ""
    
    '体裁を整える
        fncAfterUpdateCommon = Nz(objTarget)
    
    'コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        If Nz(objTarget) <> "" And _
           objTarget.ListCount <> 0 _
        Then
            fncAfterUpdateCommon = Nz(objTarget.Column(1))
        End If
    End If
    
    Call subEditMode(True)

End Function

'==============================================================================*
'
'        MODULE_NAME      :subEditMode
'        機能             :変更モード設定
'        IN               :True=変更モードに設定
'                          False=変更モード解除
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subEditMode(blnEditing As Boolean)

    If blnEditing Then
        '編集モードに設定（初回のみ）
        If Not blnEditMode Then
            Me.Caption = Me.Caption & P_CAPTION_EDITING
        End If
    Else
        Me.Caption = strFormCaption
    End If
    
    blnEditMode = blnEditing
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncBeforeUpdateCommon
'        機能             :更新前イベント時の処理
'                         :入力チェックを行う
'        IN               :①対象コントロール
'                         :②入力最大桁数
'                         :③定数値「コントロールの種類」参照
'                         :④更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncBeforeUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer, ByRef Cancel As Integer) As Boolean

    Dim strTemp     As String
    Dim blnRet      As Boolean
    
    If Left(objTarget.NAME, 3) = "cmb" Then
        'コンボボックス入力チェック
        fncBeforeUpdateCommon = fncChkComboInput(objTarget, intLength, Cancel)
    Else
        '入力桁数チェック
        fncBeforeUpdateCommon = fncChkTextInput(objTarget, intLength, Cancel)
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
'        CREATE_DATE      :2009/04/30
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
    
    If Not fncNumericCheck(cmbCheck, intMaxLength, Cancel) Then
        intRet = 1
        GoTo ChkComboInput_End
    End If
    
    '入力値がリストに存在しているかチェック
'    If cmbCheck.ListCount <> 0 Then
'        blnFound = False
'        '入力値を０サプレス
'        strValue = Format(cmbCheck, Mid(P_ZERO_SUPPRESS, 1, intMaxLength))
'        'コンボボックス内にコードが存在してる？
'        For intcount = 0 To cmbCheck.ListCount - 1
'            'コード値比較
'            If cmbCheck.Column(0, intcount) = strValue Then
'                blnFound = True
'                Exit For
'            End If
'        Next
'        If Not blnFound Then
'            MsgBox "該当するコードがありません。", vbInformation, P_KAKUNIN
'            SendKeys "{ESC}"
'            intRet = 2
'            GoTo ChkComboInput_End
'        End If
'    End If
    
ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If
    
    fncChkComboInput = intRet

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkTextInput
'        機能             :テキストボックス入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTextInput(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean
    
    If Nz(txtCheck) = "" Then
        fncChkTextInput = True
        Exit Function
    End If
    
    If CmLenb(txtCheck) > intMaxLength Then
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        fncChkTextInput = False
    Else
        fncChkTextInput = True
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_Kaito_AfterUpdate
'        機能             :cmb_Kaito_更新後処理
'        CREATE_DATE      :2009/04/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kaito_AfterUpdate()
    
'パラメータ     0   使用可
'               1   使用不可

    Select Case Me.cmb_Kaito.Column(0)
        Case 0, 2 '未回答、移動
            Call CmLockControl(Me.txt_解約日, 0)
            Call CmLockControl(Me.cmb_MoveKbn, 0)
        Case 1  '解約
            Call CmLockControl(Me.txt_解約日, 0)
            Call CmLockControl(Me.cmb_MoveKbn, 1)
            Call CmLockControl(Me.txt_MoveKbn, 1)
        Case 3  '他希望
            Call CmLockControl(Me.txt_解約日, 1)
            Call CmLockControl(Me.cmb_MoveKbn, 1)
            Call CmLockControl(Me.txt_MoveKbn, 1)
    End Select
    
End Sub

'==============================================================================*
'
'        MODULE_ID        :CmLockControl
'        機能             :入力区分によるコントロールのロックを設定
'        CREATE_DATE      :2009/04/30
'
'引数 ０  使用可
'     １  使用不可
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CmLockControl(ByVal ContNM As Object, ByVal iMode As Integer)

    Dim objControl As Object
  
    Set objControl = ContNM
  
    If iMode = 0 Then
        objControl.Enabled = True
        objControl.Locked = False
    Else
        objControl.Enabled = False
        objControl.Locked = True
    End If
  
End Sub


'==============================================================================*
'
'        MODULE_NAME      :fncNumericCheck
'        機能             :数値入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncNumericCheck(objTarget As Object, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean

    Dim strNumber       As String
    Dim lngCompare      As Long

    fncNumericCheck = False

    strNumber = Nz(objTarget, "")
    ' 空白なら入力チェックしない
    If strNumber = "" Then
        fncNumericCheck = True
        Exit Function
    End If

    lngCompare = Val(Mid(P_ALL_NINE, 1, intMaxLength))
    If Val(Nz(objTarget, 0)) > lngCompare Then
        ' 桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        Exit Function
    End If

    fncNumericCheck = True

End Function

