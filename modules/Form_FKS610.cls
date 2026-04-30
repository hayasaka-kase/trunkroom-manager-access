VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : セキュリティカード停止依頼書印刷
'
'        PROGRAM_NAME    : 日付範囲の入力
'        PROGRAM_ID      : FKS610
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2007/02/08
'        CERATER         : tajima
'        Ver             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        Ver             :
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private VarChangeDate               As Variant

'==============================================================================*
'
'        MODULE_NAME      :フォーム初期値
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2007/02/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim iParamCnt   As Integer
    Dim strParam()  As String

  On Error GoTo Exception
    
    ' 起動パラメータ取得
    strParam = Split(Nz(Me.OpenArgs), ",")
    
    ' パラメータ数チェック
    iParamCnt = UBound(strParam)
    If iParamCnt = 1 Then
        Me.txt_停止日From.VALUE = strParam(0)
        Me.txt_停止日To.VALUE = strParam(1)
    Else
        Cancel = True
        Call MSZZ024_M10("OpenArgs", "起動パラメータが正しくありません。")
        Exit Sub
    End If

    Exit Sub

Exception:
    Call MSZZ024_M00("Form_Open", True)
    Cancel = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_閉じる_Click
'        MODULE_ID        :cmd_閉じる_Click
'        CREATE_DATE      :2007/02/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_閉じる_Click()
    On Error Resume Next
    
    Call Form_FKS600.subSetStopDay("", "")  ' 印刷を実行しない印
    
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :決定
'        MODULE_ID        :cmd_決定_Click
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_決定_Click()

    On Error Resume Next

    If Nz(Me.txt_停止日From, "") = "" And _
       Nz(Me.txt_停止日To, "") = "" Then

       MsgBox "印刷対象とする停止日を入力して下さい。"
       Me.txt_停止日From.SetFocus
    End If

   Call Form_FKS600.subSetStopDay(Me.txt_停止日From, Me.txt_停止日To)
   
   doCmd.Close acForm, Me.FormName
 End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_停止日From_BeforeUpdate
'        機能             :txt_停止日From BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日From_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_停止日From, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日From_AfterUpdate
'        機能             :txt_停止日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日From_AfterUpdate()
    Me.txt_停止日From = Format(VarChangeDate, "yyyy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日From_KeyPress
'        機能             :txt_停止日From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日From_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_停止日From, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日From_GotFocus
'        機能             :txt_停止日From GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日From_GotFocus()
    Call Date_TextFormat(Me.txt_停止日From, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日From_LostFocus
'        機能             :txt_停止日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日From_LostFocus()
    If Nz(Me.txt_停止日From, "") <> "" Then
        Call Date_TextFormat(Me.txt_停止日From, 1)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_停止日To_BeforeUpdate
'        機能             :txt_停止日To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日To_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_停止日To, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日To_AfterUpdate
'        機能             :txt_停止日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日To_AfterUpdate()
    Me.txt_停止日To = Format(VarChangeDate, "yyyy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日To_KeyPress
'        機能             :txt_停止日To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日To_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_停止日To, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日To_GotFocus
'        機能             :txt_停止日To GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日To_GotFocus()
    Call Date_TextFormat(Me.txt_停止日To, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_停止日To_LostFocus
'        機能             :txt_停止日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_停止日To_LostFocus()
    If Nz(Me.txt_停止日To, "") <> "" Then
        Call Date_TextFormat(Me.txt_停止日To, 1)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub dateCheck(txtTarget As Variant, ByRef Cancel As Integer)

    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, "エラー"
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Date_TextFormat
'        機能             :日付ﾃﾞｰﾀのﾌｫｰﾏｯﾄ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Date_TextFormat(ByRef ctlTxt As TextBox, ByVal lngMode As Long)

    Dim vCode As Variant

    vCode = Nz(ctlTxt, "")
    If vCode <> "" Then
        If lngMode = 0 Then
            ctlTxt = Format(vCode, "yy/mm/dd")
        Else
            ctlTxt = Format(vCode, "yyyy/mm/dd")
        End If
    End If
    
End Sub


'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subRtnFVS500
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subRtnFVS500()

'    ' 受付番号セット
'    If Me.List_Kokyaku.ListIndex = -1 Then
'        pstrUKNO = List_Kokyaku.Column(5, 0)
'    Else
'        pstrUKNO = List_Kokyaku.Column(5, Me.List_Kokyaku.ListIndex)
'    End If
'
'    ' 遷移元画面への受付番号セット
'    Call psubSetDataFVS500
'
'    ' 画面を閉じる
'    doCmd.Close acForm, Me.FormName
End Sub
'****************************  ended or program ********************************
