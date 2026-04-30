VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS810_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス（詳細サブ）
'   プログラムＩＤ　：　FKS810
'   作　成　日　　　：  2008/06/03
'   作　成　者　　　：  iizuka
'   Ver             ：  0.0
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'
'        MODULE_NAME      :cmb_部屋割付_AfterUpdate
'        MODULE_ID        :cmb_部屋割付_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_部屋割付_AfterUpdate()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception

    If cmb_部屋割付.ListIndex = -1 Then
        Me.txt_コンテナ番号.VALUE = ""
    Else
        Me.txt_コンテナ番号.VALUE = cmb_部屋割付.Column(1, cmb_部屋割付.ListIndex)
    End If
    
    doCmd.Echo False
    
    doCmd.GoToRecord , , acNewRec
    
    doCmd.Echo True
    
    Call Form_FKS810.setKontComboDataSQL
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmb_部屋割付_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_部屋割付_KeyDown
'        MODULE_ID        :cmb_部屋割付_KeyDown
'        NOTE             :
'        CREATE_DATE      :2008/06/02
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'※未完成
'Private Sub cmb_部屋割付_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Me.cmb_部屋割付.ListIndex <> -1 Then
'        If KeyCode = vbKeyTab Then
'            If Shift = 0 Then
'                KeyCode = 0
'                Call Form_FKS810.cmd_保存.SetFocus
'            Else
'                KeyCode = 0
'                Call Form_FKS810.txt_備考.SetFocus
'            End If
'        End If
'    End If
'End Sub
