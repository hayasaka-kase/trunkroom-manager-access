VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　検索条件入力
'   プログラムＩＤ　：　FVS000
'   作　成　日　　　：  2003/06/09
'   作　成　者　　　：  Eagle Soft 田島
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FVS000"

'ヤード一覧画面ID
Const YARD_LIST = "FVS010"
Const strMsg_kaku = "確認"
Const strMsg_err = "エラー"
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :chk_近隣ヤード_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_近隣ヤード_Click()

'Start 2003/07/07 azegami@eagle-soft.co.jp
    Dim vCode As Variant

    vCode = Nz(txt_ヤード住所)

    If vCode = "" Then
        Me.chk_近隣ヤード = False
    Else
'        Me.chk_近隣ヤード = True
    End If
'End 2003/07/07 azegami@eagle-soft.co.jp

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :chk_近隣ヤード_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_近隣ヤード_KeyPress(KeyAscii As Integer)

'Start 2003/07/08 azegami@eagle-soft.co.jp
    If KeyAscii = vbKeySpace Then
        Me.chk_近隣ヤード = True
    End If
'End 2003/07/08 azegami@eagle-soft.co.jp

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :chk_種別_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_種別_Click()

'Start 2003/07/07 azegami@eagle-soft.co.jp
    Dim strfrom As String
    Dim strTo As String

    strfrom = Format(Nz(Me.txt_種別FROM, ""), "00")
    strTo = Format(Nz(Me.txt_種別TO, ""), "00")

    If strfrom = "" And strTo = "" Then
        Me.chk_種別 = False
    End If

'End 2003/07/07 azegami@eagle-soft.co.jp

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :chk_種別_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_種別_KeyPress(KeyAscii As Integer)

'Start 2003/07/08 azegami@eagle-soft.co.jp
    If KeyAscii = vbKeySpace Then
        Me.chk_種別 = True
    End If
'End 2003/07/08 azegami@eagle-soft.co.jp

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_検索開始_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_検索開始_Click()

    Dim vCode     As Variant
    Dim stDocName As String
    Dim strKinrin As String
    Dim strSyu    As String
    Dim strBuf    As String
    Dim strfrom   As String
    Dim strTo     As String
    Dim blnFLG    As Boolean
    Const strMsg = "種別は 1 ～ 40 以内で入力して下さい。"

'Start 2003/07/08 azegami@eagle-soft.co.jp
    vCode = Nz(txt_ヤード住所, "")
    strfrom = Format(Nz(Me.txt_種別FROM, ""), "00")
    strTo = Format(Nz(Me.txt_種別TO, ""), "00")

    If vCode = "" And strfrom = "" And strTo = "" Then
        MsgBox "入力して下さい。", vbInformation, strMsg_kaku
        txt_ヤード住所.SetFocus
        Exit Sub
    End If

    If vCode = "" And Me.chk_種別 = False Then
        MsgBox "種別のチェックを入れて下さい。", vbInformation, strMsg_kaku
        txt_ヤード住所.SetFocus
        Exit Sub
    End If
'End 2003/07/08 azegami@eagle-soft.co.jp

    blnFLG = False
    '住所か種別は必須入力
    If Nz(txt_ヤード住所) <> "" Then
        blnFLG = True
    End If
    
'(2003/07/02 ADD) >>>
    If chk_種別 = True And (Nz(txt_種別FROM) <> "" Or Nz(txt_種別TO) <> "") Then
        blnFLG = True
    End If
    If blnFLG = False Then
        Exit Sub
    End If
'(2003/07/02 ADD) <<<
    
    If chk_近隣ヤード = True Then
        strKinrin = "ON "
    Else
        strKinrin = "OFF"
    End If

'(2003/07/02 ADD) >>>
    If chk_種別 = True Then
        strBuf = Trim$(Nz(txt_種別FROM))
        If strBuf = "" Then
            strSyu = "01"
        Else
            '数値/範囲チェック
            If Trim$(Nz(FVSModNumeric(strBuf))) = "" Then Exit Sub
            '範囲チェック
            If CInt(strBuf) > 40 Or CInt(strBuf) < 1 Then
                MsgBox "種別：" & vbCrLf & "入力は1～40までです。", vbInformation, strMsg_kaku
                txt_種別FROM.SetFocus
                Exit Sub
            End If
            strSyu = Format$(strBuf, "00")
        End If
        
        strBuf = Trim$(Nz(txt_種別TO))
        If strBuf = "" Then
            strSyu = strSyu & "40"
        Else
            '数値チェック
            If Trim$(Nz(FVSModNumeric(strBuf))) = "" Then Exit Sub
            '範囲チェック
            If CInt(strBuf) > 40 Or CInt(strBuf) < 1 Then
                MsgBox "種別：" & vbCrLf & "入力は1～40までです。", vbInformation, strMsg_kaku
                txt_種別TO.SetFocus
                Exit Sub
            End If
            strSyu = strSyu & Format$(strBuf, "00")
        End If
    Else
        strSyu = "0000"
    End If
'(2003/07/02 ADD) <<<

'Start 2003/07/07 azegami@eagle-soft.co.jp
'矛盾チェック
    If strfrom > strTo Then
        MsgBox "種別：" & vbCrLf & "入力した値が不適当です。", vbExclamation, strMsg_err
        txt_種別FROM.SetFocus
        Exit Sub
    End If
'End 2003/07/07 azegami@eagle-soft.co.jp

    'ヤード一覧表示
    'パラメータは近隣ON/OFFと住所
    doCmd.OpenForm "FVS010", , , , , , strKinrin & strSyu & Nz(Me.txt_ヤード住所)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_終了_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
On Error Resume Next
    
    '_/_/_/ ｼｽﾃﾑを終了
    'doCmd.Quit acQuitSaveAll
    doCmd.Close acForm, Me.FormName
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()

    'フルスクリーン化
    doCmd.Maximize

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
    
    'F1:終了
    'F4:検索開始
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Call cmd_終了_Click
    Case vbKeyF4
        Dim ctlCurrentControl As Control
        KeyCode = 0
        Set ctlCurrentControl = Screen.ActiveControl
        cmd_検索開始.SetFocus
        ctlCurrentControl.SetFocus
        Call cmd_検索開始_Click
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    txt_ヤード住所.SetFocus

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    'フルスクリーン化
    doCmd.Maximize
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Unload
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload

    doCmd.Restore
    Exit Sub
    
Err_Form_Unload:
    MsgBox Error$
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_ヤード住所_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_ヤード住所_LostFocus()

'Start 2003/07/07 azegami@eagle-soft.co.jp
    Dim vCode As Variant
    
    vCode = Nz(Me.txt_ヤード住所, "")
    
    If vCode = "" Then
        Me.chk_近隣ヤード = 0
    End If
'End 2003/07/07 azegami@eagle-soft.co.jp

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_種別FROM_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_種別FROM_KeyPress(KeyAscii As Integer)

'(2003/07/02 ADD) >>>
    Call CmKeyPress_For_Num2(Me.txt_種別FROM, KeyAscii, 2)
'(2003/07/02 ADD) <<<

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_種別FROM_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_種別FROM_LostFocus()

'Start 2003/07/07 azegami@eagle-soft.co.jp
    Dim strfrom As String
    Dim strTo As String
    
    strfrom = Format(Nz(Me.txt_種別FROM, ""), "00")
    strTo = Format(Nz(Me.txt_種別TO, ""), "00")
    
    If strfrom = "" And strTo = "" Then
        Me.chk_種別 = False
    End If
'End 2003/07/07 azegami@eagle-soft.co.jp

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_種別TO_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_種別TO_KeyPress(KeyAscii As Integer)

'(2003/07/02 ADD) >>>
    Call CmKeyPress_For_Num2(Me.txt_種別TO, KeyAscii, 2)
'(2003/07/02 ADD) <<<

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_種別TO_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_種別TO_LostFocus()

'Start 2003/07/07 azegami@eagle-soft.co.jp
    Dim strTo As String
    Dim strfrom As String
    
    strTo = Format(Nz(Me.txt_種別TO, ""), "00")
    strfrom = Format(Nz(Me.txt_種別FROM, ""), "00")
    
    If strTo = "" And strfrom = "" Then
        Me.chk_種別 = False
    End If
'End 2003/07/07 azegami@eagle-soft.co.jp

End Sub
'****************************  ended or program ********************************

