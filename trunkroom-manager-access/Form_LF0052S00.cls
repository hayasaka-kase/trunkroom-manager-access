VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0052S00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 未収管理サブシステム
'
'        PROGRAM_NAME    : 宛名明細
'        PROGRAM_ID      : LF0052S00
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2001/05/11
'        CERATER         : N.MIURA
'
'        UPDATE          : 2011/06/30
'        UPDATER         : M.RYU
'        Ver             : 0.1
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
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
'    MsgBox Err.Description                                                     'DELETE 2007/02/20 K.KINEBUCHI
    Call MSZZ024_M00("Form_Activate", True)                                     'INSERT 2007/02/20 K.KINEBUCHI
    Resume Exit_Form_Activate
    
End Sub
'****************************  ended or program ********************************

