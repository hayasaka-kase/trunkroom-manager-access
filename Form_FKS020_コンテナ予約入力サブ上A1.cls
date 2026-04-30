VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS020_コンテナ予約入力サブ上A1"
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
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'**********************************************
Option Compare Database
Option Explicit
'
'Dim strMSG As String                         'メッセージ表示用
'Dim intAns As Integer                        '確認入力用
'
'Const strMsg_kaku = "確認"
'Const strMsg_Err = "エラー"
'
'Dim objDB                     As Database
'Dim objRS                     As Recordset
'
'Dim WRI_ObjWS                 As Workspace
'Dim WRI_objDb                 As Database
'
'Dim vGotText                  As Variant                'フォーカスを取得した時の内容
'Dim bSYKYMD                   As Boolean
'Dim vHData                    As Variant
'
'Public bOpenFlag_FKG010       As Boolean                'フォームが読込まれたかどうかのフラグ
'Public bChangeFlag_FKG010     As Boolean                '入力に変更があったかどうかのフラグ
'Public sJYUNO_FKG010          As String                 '複製ボタンを押した時に返り値をセットする
'
'Private Function CmBCheck_Char(ByVal oText As Variant, ByVal iLen As Integer) As Integer
'  '文字項目のチェックを一括でする。
'
'  '返り値は、
'  '
'  '２：桁数が溢れている時
'  '９：それ以外のエラー
'  '０：すべてのチェックがＯＫの時
'
'On Error GoTo err_rtn
'  CmBCheck_Char = 9
'
'  If CmLenb(oText) > iLen Then
'    CmBCheck_Char = 2
'    GoTo Exit_rtn
'  End If
'
'  CmBCheck_Char = 0
'
'Exit_rtn:
'  Exit Function
'err_rtn:
'  MsgBox "ｴﾗｰ番号" & ERR.Number & vbCrLf & ERR.Description
'  CmBCheck_Char = 9
'  GoTo Exit_rtn
'
'End Function
'
'Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal iLen As Variant) As Integer
'
'  '日付項目のチェックを一括でする
'
'  '引数
'  'oText 対象項目
'  'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
'
'  '返り値は、
'  '２：桁数が溢れている時
'  '３：日付と判断できない時
'  '９：それ以外のエラー
'  '０：すべてのチェックがＯＫの時
'  Dim sYear As String
'  Dim intLen As Integer
'On Error GoTo err_rtn
'
'  CmBCheck_Date = 2
'
'  If IsMissing(iLen) = False Then
'    intLen = Nz(CmNumeric(iLen), 0)
'  Else
'    intLen = 10
'  End If
'
'  oText = Nz(oText, "")
'
'  '桁数チェック
'  If CmLenb(oText) > intLen Then
'    Exit Function
'  End If
'
'  CmBCheck_Date = 3
'
'  If Nz(oText, "") = "" Then
'    Exit Function
'  End If
'
'  '論理チェック
'  If IsDate(oText) = False Then
'    Exit Function
'  End If
'
'  sYear = Format(oText, "yyyy")
'
'  If CmLenb(sYear) <> 4 Then
'    Exit Function
'  End If
'
'  CmBCheck_Date = 0
'
'Exit_rtn:
'  Exit Function
'err_rtn:
'  MsgBox "ｴﾗｰ番号" & ERR.Number & vbCrLf & ERR.Description
'  ERR.Clear
'  CmBCheck_Date = 9
'  GoTo Exit_rtn
'End Function
'
'Private Function CmBCheck_Num(ByVal oText As Variant, ByVal iLen As Integer, Optional ByVal isLEN As Variant) As Integer
'  '数値項目のチェックを一括でする。
'  '返り値は、
'  '１：数値と判断できない時
'  '２：桁数が溢れている時
'  '９：それ以外のエラー
'  '０：すべてのチェックがＯＫの時
'
'  '引数 oText ←判定する値
'  '     iLen  ←整数桁
'  '     isLen ←小数桁(省略された時は、小数点桁数無し)
'
'  Dim vINT    As Variant
'  Dim vsINT   As Variant
'  Dim i       As Integer
'
'  Dim dTEXT   As Double
'  Dim dFROM   As Double
'  Dim dTO     As Double
'  Dim vRET    As Variant
'
'On Error GoTo err_rtn
'  CmBCheck_Num = 9
'
'  If Nz(CmNumeric(oText), "") = "" Then       '数値と判断できないのでエラー
'    CmBCheck_Num = 1
'    GoTo Exit_rtn
'  End If
'
'  If CmChkCher(CStr(Nz(oText)), "A") = False Then '全角文字が入力されているのでエラー
'    CmBCheck_Num = 1
'    GoTo Exit_rtn
'  End If
'
'  '小数点の値が設定されている時のチェック
'  vRET = Nz(InStr(Abs(oText), "."), 0)
'
'  If IsMissing(isLEN) = False Then
'    If vRET <> 0 Then
'      If isLEN < CmLenb(Abs(oText)) - vRET Then    '小数点桁数が設定桁数を超えているエラー
'        CmBCheck_Num = 2
'        GoTo Exit_rtn
'      End If
'    End If
'  Else
'
'    If CmLenb(Abs(CmNumeric(oText))) > iLen Then           '文字数超過のためエラー
'      CmBCheck_Num = 2
'      GoTo Exit_rtn
'    End If
'
'    If vRET <> 0 Then                         '小数点の入力があるのでエラー
'      CmBCheck_Num = 2
'      GoTo Exit_rtn
'    End If
'  End If
'
'  '桁チェック
'  i = 0
'  For i = 1 To iLen
'    vINT = vINT & "9"
'  Next
'
'  If IsMissing(isLEN) = False Then
'    vsINT = "."
'    i = 0
'    For i = 1 To isLEN
'      vsINT = vsINT & "9"
'    Next
'  End If
'
'  vINT = vINT & vsINT
'  vINT = CmNumeric(vINT)
'
'  dTEXT = CmNumeric(oText)
'  dFROM = vINT * -1
'  dTO = vINT
'
'  If dFROM > dTEXT Or dTO < dTEXT Then        '論理範囲超過のためエラー
'    CmBCheck_Num = 2
'    GoTo Exit_rtn
'  End If
'
'  CmBCheck_Num = 0
'
'Exit_rtn:
'  Exit Function
'err_rtn:
'  MsgBox "ｴﾗｰ番号" & ERR.Number & vbCrLf & ERR.Description
'  CmBCheck_Num = 9
'  GoTo Exit_rtn
'End Function
'
'Private Function CMCheck_Date(ByVal vDATE As Variant)
'
'  Dim sYear As String
'  'Check_Dateでは'198/03といった形が変換されるための補助検査
'
'  CMCheck_Date = Null
'
'  If Nz(vDATE, "") = "" Then
'    Exit Function
'  End If
'
'  If IsDate(vDATE) = False Then
'    Exit Function
'  End If
'
'  sYear = Format(vDATE, "yyyy")
'
'  If CmLenb(sYear) <> 4 Then
'    Exit Function
'  End If
'
'  CMCheck_Date = vDATE
'
'End Function
'
'Private Function CmParEmpty(ByVal oEmpty As Variant, Optional ByVal oEmptyFull As Variant) As Variant
'
'  Dim vText As Variant
'
'  If IsMissing(oEmptyFull) = False Then
'    vText = oEmptyFull
'  Else
'    vText = 0
'  End If
'
'  If IsEmpty(oEmpty) = True Then
'    CmParEmpty = vText
'  Else
'    CmParEmpty = oEmpty
'  End If
'
'End Function
'
'
'Private Sub chk_取消_AfterUpdate()
'
'  '$$$ 近隣ヤード選択機能追加による追加 (01-05-2000 Araki)
'  Forms!fks020_コンテナ予約入力!flag_ch = True
'
'End Sub
'
'Private Sub cmd_詳細_Click()
'
'    vHData = "0," & Format(Me.txt_コンテナ, "000000") & Me.txt_CNTA_DATE
'
'    'コンテナ詳細　起動
'    doCmd.OpenForm "FKS030_コンテナ詳細", , , , , acDialog, vHData
'
'
'End Sub
'
'Private Sub Form_Current()
'
'    If Me.txt_顧客コード = 0 Then
'        Me.txt_賃料.Enabled = True
'        Me.txt_賃料.Locked = False
'        Me.txt_保証金.Enabled = True
'        Me.txt_保証金.Locked = False
'    Else
'        Me.txt_賃料.Enabled = False
'        Me.txt_賃料.Locked = True
'        Me.txt_保証金.Enabled = False
'        Me.txt_保証金.Locked = True
'    End If
'
'
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
'On Error Resume Next
'
'  Select Case KeyCode
'    Case vbKeyF1
'      KeyCode = 0
'      Parent.cmd_終了.SetFocus
'    Case vbKeyF2
'      KeyCode = 0
'      Parent.cmd_保存.SetFocus
'    Case vbKeyF5
'      KeyCode = 0
'        Parent.サブ下.SetFocus
'        Parent.サブ下.Form!txt_予約客先名.SetFocus
'
'    Case vbKeyF3, vbKeyF6 To vbKeyF12
'      KeyCode = 0
'  End Select
'
'
'End Sub
'
'Private Sub txt_賃料_AfterUpdate()
'
'  '$$$ 近隣ヤード選択機能追加による追加 (01-05-2000 Araki)
'  Forms!fks020_コンテナ予約入力!flag_ch = True
'
'End Sub
'
'Private Sub txt_賃料_Change()
'
'    Me.txt_変更 = 1
'
'End Sub
'
'Private Sub txt_賃料_KeyPress(keyascii As Integer)
'    Call CmKeyPress_For_Num1(Me.txt_賃料, keyascii, 9)  '桁数チェック追加 01-06-2000 Araki
'End Sub
'
'Private Sub txt_電話番号_BeforeUpdate(Cancel As Integer)
'  Dim vCODE As Variant
'
'On Error GoTo err_rtn
'
'  vCODE = Nz(Me.txt_電話番号, "")
'
'  If vCODE = "" Then
'    Exit Sub
'  End If
'
'  '入力チェック
'  Select Case CmBCheck_Char(vCODE, 15)
'    Case 0
'
'    Case 2
'      '桁数ｴﾗｰ
'      MsgBox "桁数オーバーしています。", vbInformation, strMsg_Err
'      Cancel = True
'      GoTo Exit_rtn
'    Case 9
'      MsgBox "電話番号を入力してください。", vbInformation, strMsg_Err
'      Cancel = True
'      SendKeys "{ESC}"
'      GoTo Exit_rtn
'  End Select
'
'Exit_rtn:
'  doCmd.Hourglass False
'  doCmd.SetWarnings True
'  Set objRS = Nothing
'  Exit Sub
'err_rtn:
'  MsgBox "ｴﾗｰ番号" & ERR.numberr & vbCrLf & ERR.Description
'  ERR.Clear
'  GoTo Exit_rtn
'
'
'End Sub
'
'Private Sub txt_電話番号_Change()
'
'    Me.txt_変更 = 1
'
'End Sub
'Private Sub txt_保証金_AfterUpdate()
'
'  '$$$ 近隣ヤード選択機能追加による追加 (01-05-2000 Araki)
'  Forms!fks020_コンテナ予約入力!flag_ch = True
'
'End Sub
'
'Private Sub txt_保証金_Change()
'
'    Me.txt_変更 = 1
'
'End Sub
'
'Private Sub txt_保証金_KeyPress(keyascii As Integer)
'    Call CmKeyPress_For_Num1(Me.txt_保証金, keyascii, 9)    '桁数チェック追加 01-06-2000 Araki
'End Sub
'
'Private Sub txt_予約客先名_BeforeUpdate(Cancel As Integer)
'  Dim vCODE As Variant
'
'On Error GoTo err_rtn
'
'  vCODE = Nz(Me.txt_予約客先名, "")
'
'  If vCODE = "" Then
'    Exit Sub
'  End If
'
'  '入力チェック
'  Select Case CmBCheck_Char(vCODE, 36)
'    Case 0
'
'    Case 2
'      '桁数ｴﾗｰ
'      MsgBox "桁数オーバーしています。", vbInformation, strMsg_Err
'      Cancel = True
'      GoTo Exit_rtn
'    Case 9
'      MsgBox "名称を入力してください。", vbInformation, strMsg_Err
'      Cancel = True
'      SendKeys "{ESC}"
'      GoTo Exit_rtn
'  End Select
'
'Exit_rtn:
'  doCmd.Hourglass False
'  doCmd.SetWarnings True
'  Set objRS = Nothing
'  Exit Sub
'err_rtn:
'  MsgBox "ｴﾗｰ番号" & ERR.numberr & vbCrLf & ERR.Description
'  ERR.Clear
'  GoTo Exit_rtn
'
'End Sub
'
'Private Sub txt_予約客先名_Change()
'
'    Me.txt_変更 = 1
'
'End Sub
'****************************  ended or program ********************************
