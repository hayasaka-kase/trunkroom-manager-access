VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS820_S03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約入力サブ-上段一覧
'   プログラムＩＤ　：　FVS820_S03
'   作　成　日　　　：  2003/06/09
'   作　成　者　　　：  Eagle Soft 田島
'**********************************************
'修正履歴
'   修　正　日　　　：　2005/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  画面遷移する前に保存を確認
'   修　正　内　容　：  タイトルに修正中かどうかを表示
'   Ver             ：  0.2
'**********************************************
Option Compare Database
Option Explicit
    
Const strMsg_err = "エラー"

Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer
  '文字項目のチェックを一括でする。

  '返り値は、
  '
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
  
On Error GoTo err_rtn
  CmBCheck_Char = 9
  
  If FvsModLenb(oText) > ilen Then
    CmBCheck_Char = 2
    GoTo Exit_rtn
  End If
  
  CmBCheck_Char = 0
  GoTo Exit_rtn

err_rtn:
  MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
  CmBCheck_Char = 9
Exit_rtn:
  Exit Function

End Function

'2005/05/06 K.MTISUNO Insert Start
'********************************************
'機能：予約ｺﾝﾃﾅ確定 予約受付
'詳細：新予約入力画面へｺﾝﾃﾅ番号・ﾔｰﾄﾞ番号を渡す
'********************************************
Private Sub cmdUke_Click()
    
    Dim vHData As Variant
    
    doCmd.RunCommand acCmdSaveRecord                                            'INSERT 20050718 K.ISHIZAKA
    If Parent.SaveData() = False Then                                           'INSERT 20050718 K.ISHIZAKA
        Parent.cmb_ヤードコード.SetFocus                                        'INSERT 20050718 K.ISHIZAKA
        Exit Sub                                                                'INSERT 20050718 K.ISHIZAKA
    End If                                                                      'INSERT 20050718 K.ISHIZAKA
    
    vHData = Format$(Me.Parent.txt_ヤードコード, "000000") & Format$(Me.txt_コンテナ, "000000") & Me.txt_UkeNo
    
    '予約入力画面　起動
    doCmd.OpenForm "FVS500", , , , , , vHData
    
End Sub
'2005/05/06 K.MTISUNO Insert End
'==============================================================================*
'
'        MODULE_NAME      :SetCtltxt
'        機能             :ﾃｷｽﾄﾌﾟﾛﾊﾟﾃｨ設定
'        IN               :txtCtl = text / blnEnabled = Enabled設定 / blnLocked = Locked設定
'        OUT              :
'                         :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Current()
    If txt_UkeNo <> "" Then
        Call SetCtltxt(txt_予約客先名, False, True)
        Call SetCtltxt(txt_電話番号, False, True)
    Else
        Call SetCtltxt(txt_予約客先名, True, False)
        Call SetCtltxt(txt_電話番号, True, False)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SetCtltxt
'        機能             :ﾃｷｽﾄﾌﾟﾛﾊﾟﾃｨ設定
'        IN               :txtCtl = text / blnEnabled = Enabled設定 / blnLocked = Locked設定
'        INSERT DATE      :2005/05/27
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetCtltxt(txtCtl As TextBox, blnEnabled As Boolean, blnLocked As Boolean)

    With txtCtl
        .Enabled = blnEnabled
        .Locked = blnLocked
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 
    Call Parent.Form_KeyDown(KeyCode, Shift)
    
End Sub

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
  
  GoTo Exit_rtn
  
err_rtn:
  MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
  Err.Clear
Exit_rtn:
  doCmd.Hourglass False
  doCmd.SetWarnings True
  Exit Sub
  

End Sub

Private Sub txt_電話番号_Change()

    Me.txt_変更 = 1
    Call Parent.SetChangeFlag                                                   'INSERT 20050718 K.ISHIZAKA
    
End Sub

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
  GoTo Exit_rtn
    
err_rtn:
  MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
  Err.Clear
Exit_rtn:
  doCmd.Hourglass False
  doCmd.SetWarnings True
  Exit Sub

End Sub

Private Sub txt_予約客先名_Change()

    Me.txt_変更 = 1
    Call Parent.SetChangeFlag                                                   'INSERT 20050718 K.ISHIZAKA
    
End Sub
