VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS550_S02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　取置き履歴(ヘッダ)
'   プログラムＩＤ　：　FVS550_S02
'   作　成　日　　　：  2005/12/20
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/09/11
'   修　正　者　　　：  S.SHIBAZAKI
'   Ver             ：  0.1
'   修　正　内　容　：  1.受付日を追加（デザインのみ変更）
'
'**********************************************
'修正履歴
'   修　正　日　　　：  2007/04/17
'   修　正　者　　　：  tajima
'   Ver             ：  0.2
'   修　正　内　容　：  近隣紹介ラベルを拡張（デザインのみ変更）
'
'**********************************************
Option Compare Database
Option Explicit


'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'F1=「終了」 F2=「検索」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Parent.cmd_終了.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

