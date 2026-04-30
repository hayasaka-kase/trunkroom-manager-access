VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS020_コンテナ予約入力サブ上1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　地主予約入力サブ上
'   プログラムＩＤ　：　FKS020
'   作　成　日　　　：  1999/10/25
'   作　成　者　　　：  αシステム　栗原  洋行
'**********************************************
'修正履歴
'   修　正　日　　　：　2004/07/10
'   修　正　者　　　：  N.MIURA
'   修　正　内　容　：  共益区分から代引き区分へ変更
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/05/06
'   修　正　者　　　：  K.MITSUNO
'   修　正　内　容　：  新予約入力画面追加対応
'   Ver            ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/07/13
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  契約入力画面遷移方法修正
'   Ver             ：  0.3
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  画面遷移する前に保存を確認
'   修　正　内　容　：  タイトルに修正中かどうかを表示
'   Ver             ：  0.4
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/08/03
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  解約日付には無条件にコンテナ契約ファイルの解約日付を設定するように修正
'   Ver             ：  0.5
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/10/19
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  コンテナの種別ではなく、コンテナの実帖を設定するように修正
'   Ver             ：  0.6
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/11/21
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  ｢使用料｣、｢保証金｣を入力不可に変更
'   Ver             ：  0.6
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/06/24
'   修　正　者　　　：  tajima
'   修　正　内　容　：  修繕呼出対応
'   Ver             ：  0.7
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/09/24
'   修　正　者　　　：  tajima
'   修　正　内　容　：  修繕対応その２
'                   ：  ①営業終了日以降は予約ボタン押下不可
'                   ：  ②修理中コンテナは予約ボタン押下不可
'                   ：  ③修繕文言出力対応
'   Ver             ：  0.8
'**********************************************
'修正履歴
'   修　正　日      ：2007/02/08
'   修　正　者      ：tajima
'   修　正　内　容  ：カード依頼書対応
'                     1.セキュリティカードメンテをダイアログで開くとプレビューの
'                       制御ができないので通常ウインドウとした
'   Ver             ：  0.9
'**********************************************
Option Compare Database
Option Explicit
    
Dim strMsg As String                         'メッセージ表示用
Dim intAns As Integer                        '確認入力用

Const strMsg_kaku = "確認"
Const strMsg_err = "エラー"

Dim objDb                     As Database
Dim objRs                     As Recordset

Dim WRI_ObjWS                 As Workspace
Dim WRI_objDb                 As Database

Dim vGotText                  As Variant                'フォーカスを取得した時の内容
Dim bSYKYMD                   As Boolean
Dim vHData                    As Variant

Public bOpenFlag_FKG010       As Boolean                'フォームが読込まれたかどうかのフラグ
Public bChangeFlag_FKG010     As Boolean                '入力に変更があったかどうかのフラグ
Public sJYUNO_FKG010          As String                 '複製ボタンを押した時に返り値をセットする
'==============================================================================*
'
'        MODULE_NAME      :Form_Currentイベント
'        MODULE_ID        :Form_Current
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Current()
    
    '▼2006/09/24 tajima add▼
    ' 修繕対応により入力は不可！
    '2005/05/27 K.MITSUNO INSERT
'    If txt_UkeNo <> "" Then
'        Call SetCtltxt(txt_予約客先名, False, True)
'        Call SetCtltxt(txt_電話番号, False, True)
'    Else
'        Call SetCtltxt(txt_予約客先名, True, False)
'        Call SetCtltxt(txt_電話番号, True, False)
'
'    End If
    '2005/05/27 K.MITSUNO END
   '▲2006/09/24 tajima add▲

    Call SetCtltxt(txt_賃料, False, True)    ' 2005/11/21 ADD T.SUZUKI
    Call SetCtltxt(txt_保証金, False, True)  ' 2005/11/21 ADD T.SUZUKI

    '▼2006/09/24 tajima add▼
    If Format(Nz(Parent.txt_営業終了日.VALUE, "9999/12/31"), "yyyymmdd") <= Format(Now, "yyyymmdd") Then
        Me.cmd_Uke.Enabled = False
    Else
        Me.cmd_Uke.Enabled = True
    End If
   '▲2006/09/24 tajima add▲

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

   On Error Resume Next

  ' ▼2006/06/24 Add tajima
  ' 利用状況では契約顧客無しは 0 (画面上は"000000") なので修繕呼出に対しては null を渡す
   Dim var顧客コード As Variant
   var顧客コード = IIf(Me.txt_顧客コード = 0, Null, Me.txt_顧客コード)
   Call CmReprMod.OpenFVS90n(KeyCode, _
                             Me.Parent.cmb_ヤードコード, _
                             Me.txt_コンテナ, _
                             var顧客コード, _
                             Me.txt_ACPTNO)
  ' ▲2006/06/24 Add tajima

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Parent.cmd_終了.SetFocus
            Case vbKeyF2
            KeyCode = 0
            Parent.cmd_保存.SetFocus
        Case vbKeyF5
            '2005/06/10 MITSUNO START
            With Parent.サブ下
                .SetFocus
                DoEvents
                .txt_予約客先名.SetFocus
                If Err.Number <> 0 Then
                    .txt_賃料.SetFocus
                End If
            End With
        '2005/06/10 MITSUNO END
'        KeyCode = 0
'        Parent.サブ上.SetFocus
        '2005/05/30 MITSUNO
        '予約名/電話番号変更時、予約トランに受付番号が入っていたら
        'ｺﾝﾃﾅﾏｽﾀへUpdateしない処理追加の為、以下の処理をコメント
'        Parent.サブ上.Form!txt_予約客先名.SetFocus
        Case vbKeyF3, vbKeyF6 To vbKeyF12
            KeyCode = 0
    End Select
End Sub

'2005/05/06 K.MTISUNO Insert Start
'********************************************
'機能：予約ｺﾝﾃﾅ確定 予約受付(ｺﾝﾃﾅ確定)
'詳細：新予約入力画面へｺﾝﾃﾅ番号・ﾔｰﾄﾞ番号を渡す
'********************************************
Private Sub cmd_Uke_Click()
    
    doCmd.RunCommand acCmdSaveRecord                                            'INSERT 20050718 K.ISHIZAKA
    If Parent.SaveData() = False Then                                           'INSERT 20050718 K.ISHIZAKA
        Parent.cmb_ヤードコード.SetFocus                                        'INSERT 20050718 K.ISHIZAKA
        Exit Sub                                                                'INSERT 20050718 K.ISHIZAKA
    End If                                                                      'INSERT 20050718 K.ISHIZAKA
    
    '▼2006/09/24 tajima add▼
    If Nz(Me.txt_修理中可否.VALUE, 0) = 1 Then
        MsgBox "修理中のため予約は出来ません", vbInformation, strMsg_err
        Exit Sub
    End If
    '▲2006/09/24 tajima add▲
     
    vHData = Format$(Me.Parent.cmb_ヤードコード, "000000") & Format$(Me.txt_コンテナ, "000000") & Me.txt_UkeNo
    
    '予約入力画面　起動
    doCmd.OpenForm "FVS500", , , , , , vHData
End Sub
'2005/05/06 K.MTISUNO Insert End

'==============================================================================*
'
'        MODULE_NAME      :cmd_詳細_Clickイベント
'        MODULE_ID        :cmd_詳細_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_詳細_Click()

    Dim strACPTNO As String                                                     ' 2005/07/19 ADD T.SUZUKI

    doCmd.RunCommand acCmdSaveRecord                                            'INSERT 20050718 K.ISHIZAKA
    If Parent.SaveData() = False Then                                           'INSERT 20050718 K.ISHIZAKA
        Parent.cmb_ヤードコード.SetFocus                                        'INSERT 20050718 K.ISHIZAKA
        Exit Sub                                                                'INSERT 20050718 K.ISHIZAKA
    End If                                                                      'INSERT 20050718 K.ISHIZAKA

' 2005/07/19 ADD T.SUZUKI Start
    strACPTNO = Nz(Me.txt_ACPTNO, "")  ' 受注契約番号セット

    If strACPTNO <> "" Then
        ' 契約データが存在する場合
        'doCmd.OpenForm "FVS601", , , , , acDialog, 3 & strAcptNo
        ' 2007/02/08 ↑旧↓新 chg Tajima
        doCmd.OpenForm "FVS601", , , , , , 3 & strACPTNO
    Else
        ' 契約データが存在しない場合
        MsgBox "契約情報がありません", vbInformation, strMsg_err
    End If
' 2005/07/19 ADD T.SUZUKI End

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txt_予約客先名_BeforeUpdateイベント
'        MODULE_ID        :txt_予約客先名_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_予約客先名_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant
  
    On Error GoTo err_rtn

    vCode = Nz(Me.txt_予約客先名, "")

    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
    
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
            MsgBox "名称を入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select
  
Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_予約客先名_Changeイベント
'        MODULE_ID        :txt_予約客先名_Change
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_予約客先名_Change()
    Me.txt_変更 = 1
    Call Parent.SetChangeFlag                                                   'INSERT 20050718 K.ISHIZAKA
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_電話番号_BeforeUpdateイベント
'        MODULE_ID        :txt_電話番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.txt_電話番号, "")

    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0
    
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
            MsgBox "電話番号を入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_電話番号_Changeイベント
'        MODULE_ID        :txt_電話番号_Change
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_Change()
    Me.txt_変更 = 1
    Call Parent.SetChangeFlag                                                   'INSERT 20050718 K.ISHIZAKA
End Sub
'==============================================================================*
'
'        MODULE_NAME      :chk_取消_AfterUpdateイベント
'        MODULE_ID        :chk_取消_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_取消_AfterUpdate()
    '$$$ 近隣ヤード選択機能追加による追加 (01-05-2000 Araki)
    Forms!FKS020_コンテナ予約入力!flag_ch = True
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_END_Enterイベント
'        MODULE_ID        :txt_END_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_END_Enter()
    Parent.cmd_保存.SetFocus
End Sub
' 205/11/21 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :txt_賃料_AfterUpdateイベント
''        MODULE_ID        :txt_賃料_AfterUpdate
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_賃料_AfterUpdate()
'    '$$$ 近隣ヤード選択機能追加による追加 (01-05-2000 Araki)
'    Forms!FKS020_コンテナ予約入力!flag_ch = True
'End Sub
''==============================================================================*
''
''        MODULE_NAME      :txt_賃料_Changeイベント
''        MODULE_ID        :txt_賃料_Change
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_賃料_Change()
'    Me.txt_変更 = 1
'    Call Parent.SetChangeFlag                                                   'INSERT 20050718 K.ISHIZAKA
'End Sub
''==============================================================================*
''
''        MODULE_NAME      :txt_賃料_KeyPressイベント
''        MODULE_ID        :txt_賃料_KeyPress
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_賃料_KeyPress(KeyAscii As Integer)
'    Call CmKeyPress_For_Num1(Me.txt_賃料, KeyAscii, 9)  '桁数チェック追加 01-06-2000 Araki
'End Sub
''==============================================================================*
''
''        MODULE_NAME      :txt_保証金_AfterUpdateイベント
''        MODULE_ID        :txt_保証金_AfterUpdate
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_保証金_AfterUpdate()
'    '$$$ 近隣ヤード選択機能追加による追加 (01-05-2000 Araki)
'    Forms!FKS020_コンテナ予約入力!flag_ch = True
'End Sub
''==============================================================================*
''
''        MODULE_NAME      :txt_保証金_Changeイベント
''        MODULE_ID        :txt_保証金_Change
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_保証金_Change()
'    Me.txt_変更 = 1
'    Call Parent.SetChangeFlag                                                   'INSERT 20050718 K.ISHIZAKA
'End Sub
''==============================================================================*
''
''        MODULE_NAME      :txt_保証金_KeyPressイベント
''        MODULE_ID        :txt_保証金_KeyPress
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_保証金_KeyPress(KeyAscii As Integer)
'    Call CmKeyPress_For_Num1(Me.txt_保証金, KeyAscii, 9)    '桁数チェック追加 01-06-2000 Araki
'End Sub
' 205/11/21 DEL T.SUZUKI End
'==============================================================================*
'
'        MODULE_NAME      :文字項目入力チェック
'        MODULE_ID        :CmBCheck_Char
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer
  '文字項目のチェックを一括でする。

  '返り値は、
  '
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
  
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
'        MODULE_NAME      :日付項目入力チェック
'        MODULE_ID        :CmBCheck_Date
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer
    
  '日付項目のチェックを一括でする
    
  '引数
  'oText 対象項目
  'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
    
  '返り値は、
  '２：桁数が溢れている時
  '３：日付と判断できない時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
  Dim sYear As String
  Dim intLen As Integer
On Error GoTo err_rtn
  
  CmBCheck_Date = 2
  
  If IsMissing(ilen) = False Then
    intLen = Nz(CmNumeric(ilen), 0)
  Else
    intLen = 10
  End If
  
  oText = Nz(oText, "")
  
  '桁数チェック
  If CmLenb(oText) > intLen Then
    Exit Function
  End If
  
  CmBCheck_Date = 3
  
  If Nz(oText, "") = "" Then
    Exit Function
  End If
  
  '論理チェック
  If IsDate(oText) = False Then
    Exit Function
  End If
  
  sYear = Format(oText, "yyyy")
  
  If CmLenb(sYear) <> 4 Then
    Exit Function
  End If
  
  CmBCheck_Date = 0

Exit_rtn:
  Exit Function
err_rtn:
  MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
  Err.Clear
  CmBCheck_Date = 9
  GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :数値項目入力チェック
'        MODULE_ID        :CmBCheck_Num
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, Optional ByVal isLEN As Variant) As Integer
  '数値項目のチェックを一括でする。
  '返り値は、
  '１：数値と判断できない時
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
    
  '引数 oText ←判定する値
  '     iLen  ←整数桁
  '     isLen ←小数桁(省略された時は、小数点桁数無し)
    
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
  For i = 1 To ilen
    vINT = vINT & "9"
  Next
  
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
'        MODULE_NAME      :日付項目入力チェック
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)
  
  Dim sYear As String
  'Check_Dateでは'198/03といった形が変換されるための補助検査
  
  CMCheck_Date = Null
  
  If Nz(vDATE, "") = "" Then
    Exit Function
  End If
  
  If IsDate(vDATE) = False Then
    Exit Function
  End If
  
  sYear = Format(vDATE, "yyyy")

  If CmLenb(sYear) <> 4 Then
    Exit Function
  End If
  
  CMCheck_Date = vDATE

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmParEmpty
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmParEmpty(ByVal oEmpty As Variant, Optional ByVal oEmptyFull As Variant) As Variant

  Dim vText As Variant
  
  If IsMissing(oEmptyFull) = False Then
    vText = oEmptyFull
  Else
    vText = 0
  End If
  
  If IsEmpty(oEmpty) = True Then
    CmParEmpty = vText
  Else
    CmParEmpty = oEmpty
  End If
  
End Function
'==============================================================================*
'
'        MODULE_NAME      :SetCtltxt
'        機能             :ﾃｷｽﾄﾌﾟﾛﾊﾟﾃｨ設定
'        IN               :txtCtl = text / blnEnabled = Enabled設定 / blnLocked = Locked設定
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetCtltxt(txtCtl As TextBox, blnEnabled As Boolean, blnLocked As Boolean)
    With txtCtl
        .Enabled = blnEnabled
        .Locked = blnLocked
    End With
End Sub

'★☆★☆★☆★☆★☆★☆★☆ 2006/06/24 修繕呼出対応 ★☆★☆★☆★☆★☆★☆★
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_コンテナ_DblClick
'        CREATE_DATE      :2006/06/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_コンテナ_DblClick(Cancel As Integer)
    Call Form_KeyDown(CmReprMod.P_呼出_部屋, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客コード_DblClick
'        CREATE_DATE      :2006/06/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客コード_DblClick(Cancel As Integer)
    
    If Me.txt_顧客コード.VALUE = 0 Then
        MsgBox "このコンテナ（部屋）に契約している顧客はいません。", vbInformation, strMsg_kaku
    Else
        Call Form_KeyDown(CmReprMod.P_呼出_顧客, 0)
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客コード_DblClick
'        CREATE_DATE      :2006/06/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客コード_GotFocus()
    
    If Me.txt_顧客コード.VALUE = 0 Then
        txt_顧客コード.StatusBarText = _
          "このコンテナ（部屋）を契約している顧客はいません。"
    Else
        txt_顧客コード.StatusBarText = _
          "マウスの左ボタンをダブルクリックすると、この契約顧客のｺﾒﾝﾄ･ｸﾚｰﾑ一覧が表示されます。"
    End If

End Sub
'****************************  ended or program ********************************
