VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS200_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード紹介制限・サブ
'   プログラムＩＤ　：　FVS200_S01
'   作　成　日　　　：  2007/03/01
'   作　成　者　　　：  tajima
'   Ver            ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'   Ver             ：
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   Private定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_BACKCOLOR_YELLOW  As Long = 14680063      '黄
Private Const P_BACKCOLOR_WHITE   As Long = 16777215      '白
Private Const P_チェック済み      As Integer = -1
Private Const P_未チェック        As Integer = 0

'==============================================================================*
'
' 画面最大表示
'
'
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub
'==============================================================================*
'
' キー押下処理
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Me.Parent.FunctionKeyDown(KeyCode, Shift)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :サブフォームに移動してきたとき
'        MODULE_ID        :Form_Current
'        NOTE             :サブフォーム内の状況に応じコントロールプロパティを変える
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Current()
    Call setControlProperty
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_紹介制限_BeforeUpdate(Cancel As Integer)
    
  On Error GoTo Exception
  
    If Me.chk_紹介制限.VALUE = P_チェック済み And _
        Nz(Me.Parent.txt_ヤード解約予定日, "") = "" Then
        MsgBox "元のヤードに【ヤード解約予定日】を設定していませんので紹介制限を行うことは出来ません。", vbInformation, "エラー"
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("chk_紹介制限_BeforeUpdate", True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        NOTE             :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_紹介制限_Click()

  On Error GoTo Exception
  
    If Me.chk_紹介制限.VALUE = P_チェック済み Then
        txt_制限日.VALUE = Me.Parent.txt_ヤード解約予定日.VALUE
        txt_制限元ヤード.VALUE = Me.Parent.cmb_ヤードコード.VALUE
        txt_制限元ヤード名.VALUE = Me.Parent.txt_ヤード名 & "にて停止"
    Else
        txt_制限日.VALUE = Null
        txt_制限元ヤード.VALUE = Null
        txt_制限元ヤード名.VALUE = Null
    End If
    
    Call commonAfterUpdate
    Call setControlProperty   'コントロール入力可否の属性変更
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("chk_紹介制限_BeforeUpdate", True)
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
        MsgBox "日付を正しく入力してください。", vbExclamation, "エラー"
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
Private Sub txt_制限日_AfterUpdate()
    
    Me.txt_制限日 = Format(Me.txt_制限日, "yyyy/mm/dd")  '体裁を整える
    
    If Nz(Me.txt_制限日, "") = "" Then
        ' 制限日が未入力ならチェックも外す
        Me.chk_紹介制限.SetFocus
        Me.chk_紹介制限.VALUE = P_未チェック
        Call chk_紹介制限_Click
    Else
        Call commonAfterUpdate
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
Private Sub txt_制限日_BeforeUpdate(Cancel As Integer)

    '日付入力チェック
    If dateCheck(txt_制限日, Cancel) = False Then
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
Private Sub txt_制限日_GotFocus()

    '体裁を整える
      txt_制限日 = Format(txt_制限日, "yy/mm/dd")

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
Private Sub txt_制限日_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Date(Me.txt_制限日, KeyAscii, 10)

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
Private Sub txt_制限日_LostFocus()

    '体裁を整える
    txt_制限日 = Format(Me.txt_制限日, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面コントロールプロパティの設定
'        MODULE_ID        :setControlProperty
'        NOTE             :状態(用途)によってのコントロールプロパティの設定
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setControlProperty()
    
    ' 2007/02/27 新横での業務との打ち合わせにより、制限日は入力不可とした
    ' よって以下ロジックは無効としておく
    
'    If Me.chk_紹介制限.Value = P_チェック済み Then
'        ' 編集可能に
'        Me.txt_制限日.Enabled = True
'        Me.txt_制限日.Locked = False
'    Else
'        ' 編集不可に
'        Me.txt_制限日.Enabled = False
'        Me.txt_制限日.Locked = True
'    End If
    
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
        Call Parent.SetFVS200EditChange '標題に"（修正中）"を付ける
    End If

End Sub

'****************************  ended or program ********************************
