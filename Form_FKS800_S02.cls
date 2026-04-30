VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS800_S02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  start or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800_S02
'   作　成　日　　　：  2013/11/22
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.0
'**********************************************
'   プログラム名　　：　ボックスマスタメンテナンス
'   プログラムＩＤ　：　FKS800_S02
'   作　成　日　　　：  2014/01/17
'   作　成　者　　　：  MIYAMOTO
'   Ver             ：  0.1
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        MODULE_ID        :Form_Activate
'        NOTE             :
'        CREATE_DATE      :2013/11/22
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
'↓INSERT 2014/01/17 MIYAMOTO
    Dim isException    As Boolean
    isException = False
'↑INSERT 2014/01/17 MIYAMOTO

On Error GoTo Exception
    
    doCmd.Maximize

'↓INSERT 2014/01/17 MIYAMOTO
    GoTo Finally
'↑INSERT 2014/01/17 MIYAMOTO

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Activate", True)
    
End Sub
'****************************  ended or program ********************************
