VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0051_S00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : レンタル収納システム
'
'        PROGRAM_NAME    : 契約検索-検索条件
'        PROGRAM_ID      : LF0051_S00
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2009/10/05
'        CERATER         : N.Miura
'        VAR             : 0.0
'
'        UPDATE          :
'        UPDATER         :
'        VAR             :
'                        :
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Const strMsg_err = "エラー"

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'画面ID
Private Const P_FORM_契約随時入力       As String = "MFHK000"

Private Const P_FORM_旧契約入力         As String = "MFHK009"

'コード値
Private Const P_SESSI_連続              As Integer = 1      '請求サイクル＝連続
Private Const P_SESSI_随時              As Integer = 2      '請求サイクル＝随時

Private Const P_STATI_解約予定          As Integer = 2      '契約状態＝解約予定
Private Const P_STATI_完了              As Integer = 9      '契約状態＝完了

'==============================================================================*
'
'        MODULE_NAME      :cmdKEI_Click
'        機能             :cmdKEI_Click
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdKEI_Click()

    Dim strACPTNO As String

    strACPTNO = Nz(Me.txtKeiyb, "")  ' 受注契約番号セット

    If strACPTNO <> "" Then
        doCmd.OpenForm "FVS601", , , , , , 3 & strACPTNO
    Else
        MsgBox "契約情報がありません", vbInformation, strMsg_err
    End If

End Sub
