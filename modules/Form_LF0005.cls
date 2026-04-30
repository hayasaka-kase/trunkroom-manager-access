VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME : 分析帳票
'
'        PROGRAM_NAME    : 坪単価一覧表
'        PROGRAM_ID      : LF0005
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/03/02
'        CERATER         : N.MIURA
'
'        UPDATE          :
'        UPDATER         :
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
Private Const PROG_ID = "LF0005"
Private USER_ID         As String
Private MODL_ID         As String
'
Private dbs             As Database
Private strSQL          As String
'
Private INPCHEK_STS     As Boolean
'
'==============================================================================*
'
'        MODULE_NAME      :フォームのオープン
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Form_Open
    
    Call MSHS010_M00(PROG_ID, "1")
    
    Call LF0005_DBOPEN
    
    USER_ID = LsGetUserName()

Exit_Form_Open:
    Exit Sub

Err_Form_Open:
    MsgBox Err.Description
    Resume Exit_Form_Open
End Sub
'==============================================================================*
'
'        MODULE_NAME      :実行
'        MODULE_ID        :LF0005_EXE1B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0005_EXE1B_Click()
    On Error GoTo Err_LF0005_EXE1B_Click
    
    Select Case Me.LF0005_OUTPI
           Case 0
                doCmd.OpenReport "LR0005_M00", acPreview, "", ""
           Case 1
                doCmd.OpenReport "LR0005_M10", acPreview, "", ""
           Case 2
                doCmd.OpenReport "LR0005_M20", acPreview, "", ""
    End Select
    
Exit_LF0005_EXE1B_Click:
    Exit Sub

Err_LF0005_EXE1B_Click:
    MsgBox Err.Description
    Resume Exit_LF0005_EXE1B_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :LF0005_EXE2B_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0005_EXE2B_Click()
    On Error GoTo Err_LF0005_EXE2B_Click
    
    doCmd.OutputTo acQuery, "LQ0005_M00", "MicrosoftExcel(*.xls)", "", True, ""
   'doCmd.OutputTo acOutputTable, "WK_MQAM010_TEMP", "MicrosoftExcel(*.xls)", "", True, ""

Exit_LF0005_EXE2B_Click:
    Exit Sub

Err_LF0005_EXE2B_Click:
    MsgBox Err.Description
    Resume Exit_LF0005_EXE2B_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :LF0005_EXITB_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0005_EXITB_Click()
    On Error GoTo Err_LF0005_EXITB_Click

    Call LF0005_DBCLOS
    
    Call MSHS010_M00(PROG_ID, "2")
    
    doCmd.Close

Exit_LF0005_EXITB_Click:
    Exit Sub

Err_LF0005_EXITB_Click:
    MsgBox Err.Description
    Resume Exit_LF0005_EXITB_Click
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＯＰＥＮ
'        MODULE_ID        :LF0005_DBOPEN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0005_DBOPEN()
    On Error GoTo Err_LF0005_DBOPEN
    
    Set dbs = CurrentDb()
    
Exit_LF0005_DBOPEN:
    Exit Sub

Err_LF0005_DBOPEN:
    MsgBox Err.Description
    Resume Exit_LF0005_DBOPEN

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＤＢ　ＣＬＯＳＥ
'        MODULE_ID        :LF0005_DBCLOS
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LF0005_DBCLOS()
    On Error GoTo Err_LF0005_DBCLOS
    
    dbs.Close
    
    Set dbs = Nothing

Exit_LF0005_DBCLOS:
    Exit Sub

Err_LF0005_DBCLOS:
    MsgBox Err.Description
    Resume Exit_LF0005_DBCLOS

End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面最大表示
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
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

