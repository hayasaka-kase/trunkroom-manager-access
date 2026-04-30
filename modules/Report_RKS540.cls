VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RKS540"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　メンテナンス表印刷
'   プログラムＩＤ　：　RKS520
'   作　成　日　　　：  2005/10/17
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.0
'**********************************************
Option Compare Database
Option Explicit

Private Const pstrUPD As String = "更新年月日"

'==============================================================================*
'
'        MODULE_NAME      :Report_Open
'        機能             :レポートをオープンする際、更新年月日の設定を行う
'        CREATE           :2005/10/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Report_Open(Cancel As Integer)

    Dim strUpd_F  As String  ' 更新年月日From
    Dim strUpd_T  As String  ' 更新年月日To

    On Error Resume Next

    strUpd_F = ""
    strUpd_T = ""
    If Nz(Forms![FKS540]!txt_UPDAD_F, "") <> "" Then
        strUpd_F = Nz(Forms![FKS540]!txt_UPDAD_F, "") ' 更新年月日From
    End If

    If Nz(Forms![FKS540]!txt_UPDAD_T, "") <> "" Then
        strUpd_T = Nz(Forms![FKS540]!txt_UPDAD_T, "") ' 更新年月日To
    End If

    ' 更新年月日From～To設定
    If (strUpd_F <> "" And strUpd_T <> "") Then
        Me.lbl_UPD.Caption = pstrUPD & Space(2) & Format(strUpd_F, "YYYY/MM/DD") & "～" & Format(strUpd_T, "YYYY/MM/DD")
    ElseIf (strUpd_F <> "" And strUpd_T = "") Then
        Me.lbl_UPD.Caption = pstrUPD & Space(2) & Format(strUpd_F, "YYYY/MM/DD") & "～"
    ElseIf (strUpd_F = "" And strUpd_T <> "") Then
        Me.lbl_UPD.Caption = pstrUPD & Space(2) & "～" & Format(strUpd_T, "YYYY/MM/DD")
    Else
        Me.lbl_UPD.Caption = Space(2)
    End If
End Sub
