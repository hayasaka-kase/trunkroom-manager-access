VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_RVS770"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　精算依頼リスト印刷
'   プログラムＩＤ　：　RVS770
'   作　成　日　　　：  2005/08/10
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.0
'**********************************************
Option Compare Database
Option Explicit

Private Const pstrSeisan_ReqDate As String = "精算依頼日"

Private Const pintCstFirstPage   As Integer = 26          ' 1ページMax行数
Private Const pintCstOtherPage   As Integer = 34          ' 2ページ目以降Max行数

Private pintPageCount            As Integer               ' ページ数
Private pintRowCount             As Integer               ' 行数
Private pintMaxPageCount         As Integer               ' MAXページ数

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Report_Open
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Report_Open(Cancel As Integer)

    Dim strSeisan_ReqDate_F As String
    Dim strSeisan_ReqDate_T As String
    Dim varForm_OpenArgs    As Variant

    On Error Resume Next

    ' 画面からのパラメータを取得
    varForm_OpenArgs = Nz(OpenArgs, "")
    If varForm_OpenArgs <> "" Then
        pintMaxPageCount = CInt(varForm_OpenArgs)
    End If

    strSeisan_ReqDate_F = ""
    strSeisan_ReqDate_T = ""
    If Nz(Forms![FVS770]!txt_Seisan_ReqDate_F, "") <> "" Then
        strSeisan_ReqDate_F = Nz(Forms![FVS770]!txt_Seisan_ReqDate_F, "") ' 精算依頼日From
    End If

    If Nz(Forms![FVS770]!txt_Seisan_ReqDate_T, "") <> "" Then
        strSeisan_ReqDate_T = Nz(Forms![FVS770]!txt_Seisan_ReqDate_T, "") ' 精算依頼日To
    End If

    ' 精算依頼日From～To設定
    If (strSeisan_ReqDate_F <> "" And strSeisan_ReqDate_T <> "") Then
        Me.lbl_strSeisan_ReqDate_FT.Caption = pstrSeisan_ReqDate & Space(2) & Format(strSeisan_ReqDate_F, "YYYY/MM/DD") & "～" & Format(strSeisan_ReqDate_T, "YYYY/MM/DD")
    ElseIf (strSeisan_ReqDate_F <> "" And strSeisan_ReqDate_T = "") Then
        Me.lbl_strSeisan_ReqDate_FT.Caption = pstrSeisan_ReqDate & Space(2) & Format(strSeisan_ReqDate_F, "YYYY/MM/DD") & "～"
    ElseIf (strSeisan_ReqDate_F = "" And strSeisan_ReqDate_T <> "") Then
        Me.lbl_strSeisan_ReqDate_FT.Caption = pstrSeisan_ReqDate & Space(2) & "～" & Format(strSeisan_ReqDate_T, "YYYY/MM/DD")
    Else
        Me.lbl_strSeisan_ReqDate_FT.Caption = Space(2)
    End If

    pintPageCount = 0  ' ページ数初期化
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Report_Page
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Report_Page()
    ' MAXページに達したらページ数初期化
    If pintPageCount = pintMaxPageCount Then pintPageCount = 0
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :ページヘッダーセクション_Print
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ページヘッダーセクション_Print(Cancel As Integer, PrintCount As Integer)
    pintPageCount = pintPageCount + 1  ' ページ数カウントアップ
    pintRowCount = 0                   ' 行数初期化
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :グループフッター0_Print
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub グループフッター0_Print(Cancel As Integer, PrintCount As Integer)
    Me.lbl_頁計.Caption = "頁計（" & Format(Nz(Me.txt_頁計.VALUE, 0), "###,##0") & "）件"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :レポートフッター_Print
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub レポートフッター_Print(Cancel As Integer, PrintCount As Integer)
    Me.lbl_合計.Caption = "合計（" & Format(Nz(Me.txt_合計.VALUE, 0), "###,###0") & "）件"
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :詳細_Print
'        CREATE_DATE      :2005/08/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub 詳細_Print(Cancel As Integer, PrintCount As Integer)

    pintRowCount = pintRowCount + 1  ' 行数カウントアップ

    If Nz(Me.KEIYAKU_HONSU, "") = "" Then
        ' 契約本数がNullのとき
        If (pintPageCount = 1 And pintRowCount = pintCstFirstPage) Or _
           (pintPageCount > 1 And pintRowCount = pintCstOtherPage) Then

            ' 各ページのMax行に達したとき
            Me.LineDetail02.Visible = True
            Me.LineDetail04.Visible = True

        Else
            Me.LineDetail02.Visible = False
            Me.LineDetail04.Visible = False
        End If

    Else
        ' 契約本数がNullでないとき
        Me.LineDetail02.Visible = True
        Me.LineDetail04.Visible = True
    End If

    ' 各ページ目のMAX行数に達したら、行数初期化
    If (pintPageCount = 1 And pintRowCount = pintCstFirstPage) Or _
       (pintPageCount > 1 And pintRowCount = pintCstOtherPage) Then
       pintRowCount = 0  ' 行数初期化
    End If
End Sub
