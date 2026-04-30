VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS522_契約入力検索"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 予約受付検索
'
'        PROGRAM_NAME    : 予約受付検索
'        PROGRAM_ID      : FVS520
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2004/07/24
'        CERATER         : T.SUZUKI
'        Ver             : 0.0
'
'        UPDATE          : 2011/11/24
'        UPDATER         : M.RYU
'        Ver             : 0.1
'                          検索条件を修正
'
'        UPDATE          : 2012/10/04
'        UPDATER         : M.HONDA
'        Ver             : 0.2
'                          部屋Noを追加
'
'        UPDATE          : 2019/05/07
'        UPDATER         : M.HONDA
'        Ver             : 0.3
'
'        UPDATE          : 2020/02/28
'        UPDATER         : N.IMAI
'        Ver             : 0.4
'                          ｽﾃｰﾀｽ取得方法変更
'                          リストダブルクリックで選択処理を行う
'
'        UPDATE          : 2022/02/10
'        UPDATER         : N.IMAI
'        Ver             : 0.5
'                          ｽﾃｰﾀｽ取得条件変更
'
'        UPDATE          : 2023/01/19
'        UPDATER         : M.HONDA
'        Ver             : 0.6
'                          メールアドレスで検索可能へ
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const strMsg_kaku         As String = "確認"
Private Const strMsg_err          As String = "エラー"
Private Const strMSGKetaOver      As String = "桁数オーバーしています。"
Private Const strMSGInputError    As String = "を入力してください。"

'==============================================================================*
'
'        MODULE_NAME      :フォーム初期値
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    ' フォーム初期化
    Call subFormInit
End Sub






'==============================================================================*
'
'        MODULE_NAME      :検索
'        MODULE_ID        :cmd_Search_Click
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Search_Click()

    Dim strSQL As String

    On Error GoTo cmd_Search_Click_Err

    doCmd.Hourglass True

    If fncInputCheck = False Then GoTo cmd_Search_Click_Exit

    ' リストボックス表示用SQL生成
    If fncCreateSQL(strSQL) = False Then GoTo cmd_Search_Click_Exit

    Me.List_Kokyaku.RowSource = strSQL
    Me.List_Kokyaku.Requery

    Select Case Me.List_Kokyaku.ListCount
        Case 0  ' *** 該当データなし ***
            MsgBox "該当データがありません。", vbInformation, strMsg_kaku
            Me.txt_Search.SetFocus
    
        Case Else
            '  データあり
            Me.cmd_Decision.Enabled = True
    End Select

cmd_Search_Click_Exit:
    doCmd.Hourglass False
    Exit Sub

cmd_Search_Click_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    GoTo cmd_Search_Click_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :決定
'        MODULE_ID        :cmd_Decision_Click
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Decision_Click()

    On Error Resume Next

    ' 選択入力チェック
    If Me.List_Kokyaku.ListIndex = -1 Then
        MsgBox "データを選択して下さい。", vbInformation, strMsg_err
        Exit Sub
    End If

    ' 遷移元画面（予約受付入力画面）の戻り値設定
    Call subRtnFVS500
End Sub

'==============================================================================*
'
'        MODULE_NAME      :キャンセル
'        MODULE_ID        :Cmd_Cansel_Click
'        CREATE_DATE      :2005/07/26
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
'        MODULE_NAME      :フォーム初期化
'        MODULE_ID        :subFormInit
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFormInit()

    With Me
        .txt_Search = ""

        .cmd_Decision.Enabled = False
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    fncInputCheck = False

    If Nz(Trim(Me.txt_Search), "") = "" Then

       MsgBox "検索入力項目を入力して下さい。"
       Me.txt_Search.SetFocus
       GoTo InputErr
    End If

    fncInputCheck = True

    Exit Function

InputErr:
    Exit Function

fncInputCheck_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
End Function

'==============================================================================*
'
'        MODULE_NAME      :データ取得SQL生成
'        MODULE_ID        :fncCreateSQL
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateSQL(strSQL As String) As Boolean

    On Error GoTo fncCreateSQL_Err

    fncCreateSQL = False
    strSQL = ""
    strSQL = strSQL & " select YARD_CODE AS ヤードコード,"
    strSQL = strSQL & "        YARD_NAME AS ヤード名,"
    strSQL = strSQL & "        YARD_ADDR_1 AS 住所1, "
    strSQL = strSQL & "        YARD_ADDR_2 AS 住所2, "
    strSQL = strSQL & "        YARD_ADDR_3 AS 住所3, "
    strSQL = strSQL & "        YARD_NOTE As メモ,"
    strSQL = strSQL & "        YARD_BEGIN_DAY AS 営業開始日, "
    strSQL = strSQL & "        YARD_END_DAY AS 営業終了日 "
    strSQL = strSQL & " from dbo_YARD_MAST"
    strSQL = strSQL & " where  (YARD_NAME like '%" & Me.txt_Search & "%') "
    strSQL = strSQL & " OR (YARD_ADDR_1 like '%" & Me.txt_Search & "%') "
    strSQL = strSQL & " OR (YARD_ADDR_2 like '%" & Me.txt_Search & "%') "
    strSQL = strSQL & " OR (YARD_ADDR_3 like '%" & Me.txt_Search & "%') "
    strSQL = strSQL & " AND YARD_BEGIN_DAY IS NOT NULL "
    'strSQL = strSQL & " AND  IIF(ISNULL(YARD_END_DAY) , '29990101', YARD_END_DAY) > '" & Format(DateSerial(Year(DATE), Month(DATE), 0), "YYYYMMDD") & "'"
   ' strSQL = strSQL & " AND  YARD_END_DAY > '" & Format(DateSerial(Year(DATE), Month(DATE), 0), "YYYYMMDD") & "'"
    strSQL = strSQL & "   order by YARD_CODE "

    fncCreateSQL = True

    Exit Function

fncCreateSQL_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
End Function

'==============================================================================*
'
'        MODULE_NAME      :文字項目のチェックを一括でする。
'        MODULE_ID        :CmBCheck_Char
'        CREATE_DATE      :
'
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

    On Error GoTo CmBCheck_Char_Err

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
       CmBCheck_Char = 2
       GoTo CmBCheck_Char_Exit
    End If

    CmBCheck_Char = 0

CmBCheck_Char_Exit:
    Exit Function

CmBCheck_Char_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    CmBCheck_Char = 9
    GoTo CmBCheck_Char_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subRtnFVS500
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subRtnFVS500()

    ' 受付番号セット
    If Me.List_Kokyaku.ListIndex = -1 Then
        pstrUKNO = List_Kokyaku.Column(0, 1)
    Else
        pstrUKNO = List_Kokyaku.Column(0, Me.List_Kokyaku.ListIndex + 1)
    End If

   ' 遷移元画面の受付番号コンボボックスに値をセット
    With Forms("FVS600").Controls("cmb_ヤードコード")
        .VALUE = pstrUKNO
    End With

    ' 画面を閉じる
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :顧客コード
'        MODULE_ID        :List_Kokyaku_DblClick
'        CREATE_DATE      :2020/02/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub List_Kokyaku_DblClick(Cancel As Integer)

    If Me.List_Kokyaku.ListIndex <> -1 Then
        Call cmd_Decision_Click
    End If

End Sub
'****************************  ended or program ********************************
