VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS220_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード解約一覧画面
'   プログラムＩＤ　：　FVS220_S01
'   作　成　日　　　：  2009/04/30
'   作　成　者　　　：
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2011/09/30
'   修　正　者　　　：  M.RYU
'   修　正　内　容　：  ﾌｫｰﾏｯﾄ調整
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'   Ver             ：
'**********************************************

Option Compare Database
Private Const P_PROG_ID     As String = "FVS220_S01"
Private Const P_PREVIEW     As Integer = 1

'==================================================
'
'        MODULE_NAME      :明細処理
'        MODULE_ID        :cmd_明細_Click
'        CREATE_DATE      :2009/04/30
'        戻り値           :
'
'==================================================
Private Sub cmd_プレビュー_Click()
    Dim strYcode As String
    Dim strCNTANO As String
    Dim strMoveYCODE As String
    Dim strMoveCNTANO As String
    On Error GoTo ErrorHandle
    
    '状態チェック
    If Not funcDataChk Then
        Exit Sub
    End If
    'プレビュー実行
    strYcode = Nz(txt_YCODE, "")
    strCNTANO = Nz(txt_CNTANO, "")
    strMoveYCODE = Nz(txt_MOVE_YCODE, "")
    strMoveCNTANO = Nz(txt_MOVE_CNTANO, "")
    Call Me.Parent.funcPrintHelperKyk(strYcode, strCNTANO, strMoveYCODE, strMoveCNTANO, P_PREVIEW)
    Exit Sub
ErrorHandle:
        Call MSZZ024_M00("cmd_プレビュー_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :レコード選択チェック件数
'        MODULE_ID        :funcSelectChk
'        CREATE_DATE      :2009/04/30
'        戻り値           :funcSelectChk,Output,Boolean,True 正常,False エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function funcDataChk() As Boolean
    If Nz(txt_SITUATION, "") = "未紹介" Then
        MsgBox "未紹介です｡ 紹介を行ってください｡", vbOKOnly, P_MSG_KAKUNIN
        funcDataChk = False
        GoTo funcDataChk_Exit
    End If
    funcDataChk = True
    Exit Function
funcDataChk_Exit:
End Function

'==================================================
'   イベント名：cmd_明細_Click
'   処理      ：ボタンクリック時
'   ボタン処理：明細
'   作成日    ：2009/04/30  新規作成
'   更新日    ：0000/00/00
'==================================================
Private Sub cmd_明細_Click()
    Dim str_Param As String
    
    If txt_SITUATION = "未紹介" Then
        MsgBox "未紹介です。紹介してください。", vbInformation, P_MSG_KAKUNIN
        Exit Sub
    End If
    str_Param = P_PROG_ID & "," & txt_YCODE & "," & txt_CNTANO & "," & txt_MOVE & "," & IIf(txt_MOVE_RENTAL = "", "0", Format$(txt_MOVE_RENTAL, "########0"))
    doCmd.OpenForm "FVS221", , , , , , str_Param

End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームキー押下時処理
'       MODULE_ID       : Form_KeyDown
'       CREATE_DATE     : 2009/04/30
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Parent.KeyDown(KeyCode, Shift)
End Sub
