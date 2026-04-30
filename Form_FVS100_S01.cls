VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS100_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約契約解約帳票　用途一覧画面
'   プログラムＩＤ　：　FVS100_S01
'   作　成　日　　　：  2009/10/23
'   作　成　者　　　：  kihara
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'   Ver             ：
'**********************************************

Option Compare Database
Private Const P_PROG_ID     As String = "FVS100_S01"
Private Const P_PREVIEW     As Integer = 1

'==============================================================================*
'
'        MODULE_NAME      :chk_用途_KeyDown
'        機能             :chk_用途_KeyDownイベント
'        CREATE_DATE      :2009/10/23
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_用途_KeyDown(KeyCode As Integer, Shift As Integer)

    'フォーカス移動制御（最終行の時、TabキーでEXCELボタンへ）
    If Me.RecordsetClone.RecordCount > 0 Then
        Me.RecordsetClone.MoveLast
        If Me.CurrentRecord = Me.RecordsetClone.RecordCount Then
            Select Case KeyCode
                Case vbKeyTab
                    KeyCode = 0
                    Form_FVS100.cmd_Excel.SetFocus
            End Select
        End If
    End If

End Sub
