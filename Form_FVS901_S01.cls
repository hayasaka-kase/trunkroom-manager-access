VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS901_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード修繕一覧画面
'   プログラムＩＤ　：　FVS901
'   作　成　日　　　：  2006/05/13
'   作　成　者　　　：  T.SUZUKI
'   Ver            ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  FVS901_S01
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  修理時に完了日付を出すように修正。
'   Ver             ：  Ver0.1
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   Private定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_MSG_MODE_INS          As String = "1"                           ' 登録
Private Const P_MSG_MODE_UPD          As String = "2"                           ' 修正

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
'        MODULE_NAME      :cmd_修正_Click イベント
'        MODULE_ID        :cmd_修正_Click
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_修正_Click()

  On Error GoTo Exception

    ' 修繕入力ダイアログ表示
    Call Parent.subOpen_FVS910(P_MSG_MODE_UPD, Me.txt_ユニークコード)

    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_修正_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cnd_削除_Click イベント
'        MODULE_ID        :cnd_削除_Click
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cnd_削除_Click()

    Dim intRtn As Integer

  On Error GoTo Exception

    ' ユニークコードの編集
    If Nz(Me.txt_ユニークコード, "") = "" Then Exit Sub

    Call Parent.subUpdFVS901(Me.txt_ユニークコード)

    Exit Sub

Exception:
    Call MSZZ024_M00("cnd_削除_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load イベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2010/12/24
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim errorMessage As String
    
  On Error GoTo Exception
    errorMessage = ""
    
    
    
    
    
    
    Exit Sub

Exception:
    Call MSZZ024_M00("Form_Load", True)
    Err.Clear

End Sub
'****************************  ended or program ********************************
