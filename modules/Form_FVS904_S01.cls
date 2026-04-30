VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS904_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード一括修繕一覧画面
'   プログラムＩＤ　：　FVS904
'   作　成　日　　　：  2017/11/03
'   作　成　者　　　：  T.SUZUKI
'   Ver            ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  FVS904_S01
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  修理時に完了日付を出すように修正。
'   Ver             ：  Ver0.0
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

Private Const P_MSG_ERR               As String = "エラー"
Private Const P_MSG_KAKUNIN           As String = "確認"

Private Sub chk_選択_Click()

'    If chk_選択.VALUE And cntOnData() > 0 Then
'        If Me.txt_info = getOnDataOne() Then
'        Else
'            MsgBox P_MSG_ERR_SELECT
'            chk_選択.VALUE = False
'            Exit Sub
'        End If
'    End If

    Call FVS904_sub.chkOnOffData(Me.Parent.cmb_ヤードコード, txt_コンテナ番号, txt_STEP, chk_選択)
    'MsgBox cntOnData

End Sub

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
'        CREATE_DATE      :2017/11/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_修正_Click()

    Dim fvs910Form  As Form

  On Error Resume Next
        Set fvs910Form = Forms("FVS910")
        If Not fvs910Form Is Nothing Then
            MsgBox "修繕入力画面が開いています。" & vbCr & "修繕入力画面を閉じて下さい。", vbInformation, P_MSG_KAKUNIN
            fvs910Form.SetFocus
            Exit Sub
        End If

  On Error GoTo Exception


    ' 修繕入力ダイアログ表示
    If Nz(Me.FVS904_UNIQEC, "") = "" Then Exit Sub
    
    If Me.FVS904_UNIQEC <> "" Then
    
        Call Parent.subOpen_FVS910(P_MSG_MODE_UPD, Me.FVS904_UNIQEC, False, Me.FVS904_CNTA_NO)
    End If

    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_修正_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cnd_削除_Click イベント
'        MODULE_ID        :cnd_削除_Click
'        CREATE_DATE      :2017/11/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cnd_削除_Click()

    Dim intRtn As Integer
    Dim fvs910Form  As Form

On Error Resume Next
    Set fvs910Form = Forms("FVS910")
    If Not fvs910Form Is Nothing Then
        MsgBox "修繕入力画面が開いているため削除はできません。" & vbCr & "修繕入力画面を閉じて下さい。", vbInformation, P_MSG_KAKUNIN
        fvs910Form.SetFocus
        Exit Sub
    End If
    
  On Error GoTo Exception

    ' ユニークコードの編集
    If Nz(Me.FVS904_UNIQEC, "") = "" Then Exit Sub

    Call Parent.subUpdFVS904(Me.FVS904_UNIQEC, Me.FVS904_CNTA_NO)

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
