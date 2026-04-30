VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RVS750"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　延長リスト印刷
'   プログラムＩＤ　：　RVS750
'   作　成　日　　　：  2005/08/08
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.0
'**********************************************
Option Compare Database
Option Explicit

Private Const pstrCYDATE As String = "解約予定日"

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Report_Open
'        CREATE_DATE      :2005/08/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Report_Open(Cancel As Integer)

    Dim strCYDATE_F As String
    Dim strCYDATE_T As String

    On Error Resume Next

    If Nz(Forms![FVS750]!txt_CYDATE_F, "") <> "" Then
        strCYDATE_F = Forms![FVS750]!txt_CYDATE_F  ' 解約予定日From
    End If

    If Nz(Forms![FVS750]!txt_CYDATE_T, "") <> "" Then
        strCYDATE_T = Forms![FVS750]!txt_CYDATE_T  ' 解約予定日To
    End If

    ' 解約予定日From～To設定
    If (strCYDATE_F <> "" And strCYDATE_T <> "") Then
        Me.lbl_CYDATE_FT.Caption = pstrCYDATE & Space(2) & Format(strCYDATE_F, "YYYY/MM/DD") & "～" & Format(strCYDATE_T, "YYYY/MM/DD")
    ElseIf (strCYDATE_F <> "" And strCYDATE_T = "") Then
        Me.lbl_CYDATE_FT.Caption = pstrCYDATE & Space(2) & Format(strCYDATE_F, "YYYY/MM/DD") & "～"
    ElseIf (strCYDATE_F = "" And strCYDATE_T <> "") Then
        Me.lbl_CYDATE_FT.Caption = pstrCYDATE & Space(2) & "～" & Format(strCYDATE_T, "YYYY/MM/DD")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :グループフッター0_Format
'        CREATE_DATE      :2005/08/08
'        処理内容          :解約区分のグループフッターフォーマット時、解約区分計タイトルの文字列編集
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub グループフッター0_Format(Cancel As Integer, FormatCount As Integer)

    Me.lbl_解約区分計.Caption = Me.txt_解約区分名.VALUE & "計（" & Me.txt_解約区分計.VALUE & "）件"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :レポートフッター_Format
'        CREATE_DATE      :2005/08/08
'        処理内容          :レポートフッターフォーマット時、合計タイトルの文字列編集
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub レポートフッター_Format(Cancel As Integer, FormatCount As Integer)

    Me.lbl_合計.Caption = "合計（" & Me.txt_解約区分合計.VALUE & "）件"
End Sub
