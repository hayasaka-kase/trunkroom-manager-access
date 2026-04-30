VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RVS760_BK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　メンテナンス表印刷
'   プログラムＩＤ　：　RVS760
'   作　成　日　　　：  2005/08/03
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   Ver             ：0.1
'   修　正　日      ：2005/10/15
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：ワークテーブル変更に伴う修正
'**********************************************
'修正履歴
'   Ver             ：0.2
'   修　正　日      ：2005/10/31
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：解約年月日From～Toの表示処理を追加
'**********************************************
'修正履歴
'   Ver             ：0.4
'   修　正　日      ：2006/05/12
'   修　正　者      ：N.MIURA
'   修　正　内　容  ：コンテナ番号のフォントサイズ変更（１２－＞１１）
'**********************************************
'修正履歴
'   Ver             ：0.7
'   修　正　日      ：2010/12/13
'   修　正　者      ：M.HONDA
'   修　正　内　容  ：刻印欄を追加
'**********************************************
'修正履歴
'   Ver             ：0.8
'   修　正　日      ：2011/06/03
'   修　正　者      ：M.RYU
'   修　正　内　容  ：ヤード住所・コンテナ用途を追加
'**********************************************
Option Compare Database
Option Explicit

Private Const pstrMNT_REQDATE As String = "メンテナンス依頼日"
Private Const pstrKAIYAKUDATE As String = "解約年月日"  ' 2005/10/31 ADD T.SUZUKI

'==============================================================================*
'
'        MODULE_NAME      :Report_Open
'        機能             :レポートをオープンする際、改ページ制御を行う
'        CREATE           :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Report_Open(Cancel As Integer)

    Dim strMNT_REQDATE_F As String
    Dim strMNT_REQDATE_T As String

' 2005/10/31 ADD T.SUZUKI Start
    Dim strKAIYAKUDATE_F As String
    Dim strKAIYAKUDATE_T As String

    Dim strTITTLE        As String
' 2005/10/31 ADD T.SUZUKI End

    On Error Resume Next

    If Forms![FVS760]!chk_KaiPage.VALUE = -1 Then
        Me.groupHeader.forceNewPage = 1
    End If

' 2005/10/31 ADD T.SUZUKI Start
    strKAIYAKUDATE_F = ""
    strKAIYAKUDATE_T = ""
    strMNT_REQDATE_F = ""
    strMNT_REQDATE_T = ""

    If Nz(Forms![FVS760]!txt_KaiyakuDate_F, "") <> "" Then
        strKAIYAKUDATE_F = Forms![FVS760]!txt_KaiyakuDate_F  ' 解約年月日From
    End If

    If Nz(Forms![FVS760]!txt_KaiyakuDate_T, "") <> "" Then
        strKAIYAKUDATE_T = Forms![FVS760]!txt_KaiyakuDate_T  ' 解約年月日To
    End If
' 2005/10/31 ADD T.SUZUKI End

    If Nz(Forms![FVS760]!txt_Mnt_ReqDate_F, "") <> "" Then
        strMNT_REQDATE_F = Forms![FVS760]!txt_Mnt_ReqDate_F  ' メンテナンス依頼日From
    End If

    If Nz(Forms![FVS760]!txt_Mnt_ReqDate_T, "") <> "" Then
        strMNT_REQDATE_T = Forms![FVS760]!txt_Mnt_ReqDate_T  ' メンテナンス依頼日To
    End If

' 2005/10/31 ADD T.SUZUKI Start
    ' 解約年月日From～To設定
    If (strKAIYAKUDATE_F <> "" And strKAIYAKUDATE_T <> "") Then
        strTITTLE = pstrKAIYAKUDATE & Space(2) & Format(strKAIYAKUDATE_F, "YYYY/MM/DD") & "～" & Format(strKAIYAKUDATE_T, "YYYY/MM/DD")
    ElseIf (strKAIYAKUDATE_F <> "" And strKAIYAKUDATE_T = "") Then
        strTITTLE = pstrKAIYAKUDATE & Space(2) & Format(strKAIYAKUDATE_F, "YYYY/MM/DD") & "～"
    ElseIf (strKAIYAKUDATE_F = "" And strKAIYAKUDATE_T <> "") Then
        strTITTLE = pstrKAIYAKUDATE & Space(2) & "～" & Format(strKAIYAKUDATE_T, "YYYY/MM/DD")
    End If

    If strTITTLE <> "" Then strTITTLE = strTITTLE & Space(2)

    ' メンテ依頼日From～To設定
    If (strMNT_REQDATE_F <> "" And strMNT_REQDATE_T <> "") Then
        strTITTLE = strTITTLE & pstrMNT_REQDATE & Space(2) & Format(strMNT_REQDATE_F, "YYYY/MM/DD") & "～" & Format(strMNT_REQDATE_T, "YYYY/MM/DD")
    ElseIf (strMNT_REQDATE_F <> "" And strMNT_REQDATE_T = "") Then
        strTITTLE = strTITTLE & pstrMNT_REQDATE & Space(2) & Format(strMNT_REQDATE_F, "YYYY/MM/DD") & "～"
    ElseIf (strMNT_REQDATE_F = "" And strMNT_REQDATE_T <> "") Then
        strTITTLE = strTITTLE & pstrMNT_REQDATE & Space(2) & "～" & Format(strMNT_REQDATE_T, "YYYY/MM/DD")
    ElseIf (strMNT_REQDATE_F = "" And strMNT_REQDATE_T = "") Then
          strTITTLE = strTITTLE & "メンテナンス依頼日未入力"
    End If
' 2005/10/31 ADD T.SUZUKI End

' 2005/10/31 DEL T.SUZUKI Start
'    ' メンテ依頼日From～To設定
'    If (strMNT_REQDATE_F <> "" And strMNT_REQDATE_T <> "") Then
'        Me.lbl_MNT_REQDATE_FT.Caption = pstrMNT_REQDATE & Space(2) & Format(strMNT_REQDATE_F, "YYYY/MM/DD") & "～" & Format(strMNT_REQDATE_T, "YYYY/MM/DD")
'    ElseIf (strMNT_REQDATE_F <> "" And strMNT_REQDATE_T = "") Then
'        Me.lbl_MNT_REQDATE_FT.Caption = pstrMNT_REQDATE & Space(2) & Format(strMNT_REQDATE_F, "YYYY/MM/DD") & "～"
'    ElseIf (strMNT_REQDATE_F = "" And strMNT_REQDATE_T <> "") Then
'        Me.lbl_MNT_REQDATE_FT.Caption = pstrMNT_REQDATE & Space(2) & "～" & Format(strMNT_REQDATE_T, "YYYY/MM/DD")
'    End If
' 2005/10/31 DEL T.SUZUKI End

    Me.lbl_MNT_REQDATE_FT.Caption = strTITTLE  ' 2005/10/31 ADD T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Report_Close
'        機能             :
'        CREATE           :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Report_Close()
    If Me.groupHeader.forceNewPage = 1 Then
        Me.groupHeader.forceNewPage = 0
    End If
End Sub
