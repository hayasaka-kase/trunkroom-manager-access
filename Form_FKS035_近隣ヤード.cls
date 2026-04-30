VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS035_近隣ヤード"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　近隣ヤード
'   プログラムＩＤ　：　FKS035
'   作　成　日　　　：  2000/01/04
'   作　成　者　　　：  αシステム　(Araki)
'**********************************************
'修正履歴
'   修　正　日　　　：2005/10/31
'   修　正　者　　　：T.SUZUKI
'   修　正　内　容　：1.近隣ヤードテーブルの追加に伴い、データの抽出方法を変更
'                    2.画面レイアウトの変更
'**********************************************
'   修　正　日　　　：2006/04/11  Ver 0.2
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：1.距離を追加
'**********************************************
Option Compare Database
Option Explicit

Private Const pcstrMsg_Kaku As String = "確認"
Private Const pcstrMsg_Err  As String = "エラー"

'==============================================================================*
'
'        MODULE_NAME      :Form_Loadイベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Dim strYcode As String
    Dim strSQL   As String

    ' ヤードコード取得
    strYcode = Nz(Forms!FKS020_コンテナ予約入力!cmb_ヤードコード, "")

    ' ヤードコードチェック
    If strYcode = "" Then
        Call cmd_Cancel_Click
        Exit Sub
    End If

    ' 表示用近隣ヤードデータ抽出のSQL作成
    strSQL = fncCreateSQL(strYcode)

    ' リストボックスにデータにSQL設定
    Me.List_NYAR.RowSource = strSQL
    Me.List_NYAR.Requery

    Select Case Me.List_NYAR.ListCount
        Case 0
            ' 該当データ0件
            MsgBox "該当データが存在しません。", vbInformation, pcstrMsg_Kaku
            Call cmd_Cancel_Click
            Exit Sub

'        Case 1
'            ' 該当する近隣ヤードが1件の場合
'            strYcode = Me.List_NYAR.Column(0, 0)
'
'            ' 呼び出し画面にヤードコードを返す
'            Call subReturnYCODE(strYcode)
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Decision_Clickイベント
'        MODULE_ID        :cmd_Decision_Click
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Decision_Click()

    Dim strYcode As String  ' ヤードコード

    ' 選択入力チェック
    If Me.List_NYAR.ListIndex = -1 Then
        MsgBox "データを選択して下さい。", vbInformation, pcstrMsg_Err
        Exit Sub
    End If

    ' 選択された近隣ヤードを取得
    If Me.List_NYAR.ListIndex = -1 Then
        strYcode = Me.List_NYAR.Column(0, 0)
    Else
        strYcode = Me.List_NYAR.Column(0, Me.List_NYAR.ListIndex)
    End If

    ' 呼び出し画面にヤードコードを返す
    Call subReturnYCODE(strYcode)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Cancel_Clickイベント
'        MODULE_ID        :cmd_Cancel_Click
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Cancel_Click()

    ' 画面を閉じる
    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :リストボックス表示のSQL作成
'        MODULE_ID        :fncCreateSQL
'        Parameter        :strYCDDE = ヤードコード
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateSQL(ByVal strYCDDE As String) As String

    Dim strSQL   As String

    On Error GoTo fncCreateSQL_Err

    fncCreateSQL = ""

    ' パラメータチェック
    If Nz(strYCDDE, "") = "" Then Exit Function

    ' ヤードコードをもとに近隣ヤードコードを取得
    strSQL = " SELECT format(dbo_NYAR_MAST.NYAR_NCODE,'000000'), " & Chr(13)  ' 近隣ヤードコード
    strSQL = strSQL & " dbo_YARD_MAST.YARD_NAME, " & Chr(13)                  ' ヤード名
    strSQL = strSQL & " format(dbo_NYAR_MAST.NYAR_KIRO,'0.00'), " & Chr(13)  ' 距離  'INSERT 2006/04/11 K.ISHIZAKA
    strSQL = strSQL & " NZ(dbo_YARD_MAST.YARD_ADDR_1, '') " & Chr(13)         ' ヤード住所
    strSQL = strSQL & " FROM dbo_NYAR_MAST, " & Chr(13)  ' *** 近隣ヤードマスタ ***
    strSQL = strSQL & "      dbo_YARD_MAST " & Chr(13)   ' *** ヤードマスタ
    strSQL = strSQL & " WHERE dbo_NYAR_MAST.NYAR_NCODE = dbo_YARD_MAST.YARD_CODE " & Chr(13)
    strSQL = strSQL & "   AND dbo_NYAR_MAST.NYAR_NCODE <> " & strYCDDE & Chr(13)
    strSQL = strSQL & "   AND dbo_NYAR_MAST.NYAR_YCODE =  " & strYCDDE
    strSQL = strSQL & " ORDER BY NZ(dbo_NYAR_MAST.NYAR_KIRO,9999) "             'INSERT 2006/04/11 K.ISHIZAKA

    fncCreateSQL = strSQL

    Exit Function

fncCreateSQL_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :SetYCodeイベント
'        MODULE_ID        :subReturnYCODE
'        Parameter        :strYCDDE = ヤードコード
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subReturnYCODE(ByVal strYcode As String)

    ' パラメータチェック
    If Nz(strYcode, "") = "" Then Exit Sub

    ' マスタチェック
    If fncCHK_CNTA_MAST(strYcode) = True Then
        Forms!FKS020_コンテナ予約入力!cmb_ヤードコード = Format(strYcode, "000000")
        doCmd.Close acForm, Me.FormName
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :コンテナ登録データ有無のチェック
'        MODULE_ID        :CHK_CNTA_MAST
'        Parameter        :strYCDDE = ヤードコード
'        CREATE_DATE      :2005/10/31
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCHK_CNTA_MAST(ByVal strYcode As String) As Boolean

    Dim stWHERE As String
    Dim vRet    As Variant

    On Error GoTo fncCHK_CNTA_MAST_Err

    fncCHK_CNTA_MAST = False

    If strYcode = "" Then
        fncCHK_CNTA_MAST = True
        GoTo fncCHK_CNTA_MAST_Exit
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    stWHERE = ""
    stWHERE = " CNTA_CODE =" & strYcode

    vRet = Nz(DLookup("CNTA_NO", "dbo_CNTA_MAST", stWHERE), "")

    If vRet = "" Then
       MsgBox "コンテナ登録がされておりません", vbInformation, pcstrMsg_Kaku
       GoTo fncCHK_CNTA_MAST_Exit
    End If

    fncCHK_CNTA_MAST = True

fncCHK_CNTA_MAST_Exit:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Exit Function

fncCHK_CNTA_MAST_Err:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear
    GoTo fncCHK_CNTA_MAST_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :ダブルクリックで決定
'        MODULE_ID        :List_NYAR_DblClick
'        CREATE_DATE      :2006/04/11
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub List_NYAR_DblClick(Cancel As Integer)
    If Me.List_NYAR.ListIndex >= 0 Then
        Call cmd_Decision_Click
    End If
End Sub
'****************************  ended or program ********************************
