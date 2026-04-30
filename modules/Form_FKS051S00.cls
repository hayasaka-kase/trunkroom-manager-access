VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS051S00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : コンテナ管理サブシステム
'
'        PROGRAM_NAME    : 売上登録
'        PROGRAM_ID      : FKS051S00
'        PROGRAM_KBN     : FORM
'
'        CREATE          : N.MIURA
'        CERATER         : 2003/10/17
'                        :
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'==============================================================================*
'
'画面最大表示
'
'
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    On Error GoTo Err_Form_Activate
    
    doCmd.Maximize

Exit_Form_Activate:
    Exit Sub

Err_Form_Activate:
    MsgBox Err.Description
    Resume Exit_Form_Activate
    
End Sub
'****************************  ended or program ********************************

