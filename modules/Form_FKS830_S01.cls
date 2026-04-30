VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS830_S01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  start or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　割引マスタ一覧（サブ）
'   プログラムＩＤ　：　FKS830_S01
'   作　成　日　　　：  2014/02/03
'   作　成　者　　　：  EGL
'   Ver             ：  0.0
'**********************************************
'   プログラム名　　：　割引マスタ一覧（サブ）
'   プログラムＩＤ　：　FKS830_S01
'   作　成　日　　　：  2014/04/21
'   作　成　者　　　：  EGL
'   Ver             ：  0.1
'                   ：  設置用途に「看板」を追加(フォームのみ)
'                   ：  列見出しの「その他」を全て「不明」に変更(フォームのみ)
'**********************************************

Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS830_S01"
Private Const P_ROOM_WORK_TABLE         As String = "WK_FKS830"
Private Const P_NEBIKI_LIST_WORK_TABLE  As String = "WK_FKS830_S01_TEMP"

Private Const P_KAKUNIN                 As String = "確認"
Private Const P_ERROR                   As String = "エラー"

Private pobjDB                          As Object

Private pstrBumonCode                   As String       ' 部門コード

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Dim isException As Boolean

    isException = False

On Error GoTo Exception

    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb

    '画面内容切替

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Close
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2014/02/03
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Dim isException As Boolean

    isException = False

On Error GoTo Exception

    'コネクションクローズ
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Close", True)

End Sub
'****************************  ended or program ********************************

