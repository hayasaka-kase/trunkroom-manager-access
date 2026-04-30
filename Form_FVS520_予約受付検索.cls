VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS520_予約受付検索"
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
'        MODULE_NAME      :顧客コード
'        MODULE_ID        :txt_KCode_AfterUpdate
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KCode_AfterUpdate()

    If Nz(Me.txt_KCode, "") = "" Then Exit Sub
    Me.txt_KCode = Format(txt_KCode, "000000")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :顧客コード
'        MODULE_ID        :txt_KCode_BeforeUpdate
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KCode_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_KCode_BeforeUpdate_Err

    vCode = Nz(Me.txt_KCode, "")

    ' 入力チェック
    Select Case CmBCheck_Char(vCode, 6)
        Case 0

        Case 2
            ' 桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub
    
        Case 9
            MsgBox "コードを入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_KCode_BeforeUpdate_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :顧客コード
'        MODULE_ID        :txt_KCode_KeyPress
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KCode_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_KCode, KeyAscii, 6)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :電話番号
'        MODULE_ID        :txt_TEL_BeforeUpdate
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_TEL_BeforeUpdate_Err

    vCode = Nz(Me.txt_Tel, "")

    If vCode = "" Then Exit Sub

    ' 入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0

        Case 2
            ' 桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "電話番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_TEL_BeforeUpdate_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :電話番号
'        MODULE_ID        :txt_TEL_KeyPress
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_Tel, KeyAscii, 15)    '←DB定義より15桁まで入力可。
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
            Me.txt_KKana.SetFocus

'        Case 1  ' *** 該当データ1件        'DELETE 2011/11/24 M.RYU
'            ' 遷移元画面（予約受付入力画面）の戻り値設定
'            Call subRtnFVS500              'DELETE 2011/11/24 M.RYU
        
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
        .txt_KKana = ""   ' 顧客名ｶﾅ
        .txt_KCode = ""   ' 顧客コード
        .txt_Tel = ""     ' 電話番号 Or 携帯番号 Or FAX番号

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

    If Nz(Trim(Me.txt_KKana), "") = "" And _
       Nz(Trim(Me.txt_KCode), "") = "" And _
       Nz(Trim(Me.txt_Mail), "") = "" And _
       Nz(Trim(Me.txt_Tel), "") = "" Then

       MsgBox "検索入力項目を入力して下さい。"
       Me.txt_KKana.SetFocus
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
    strSQL = " SELECT   "
    strSQL = strSQL & "  iif(NAME_C088N.ADMEN <> '' AND YOUKT_YUKBN NOT IN(10,51,52,53) , NAME_C088N.ADMEN, NAME_C080N.ADMEN)   AS ｽﾃｰﾀｽ,  " & Chr(13)  'INSERT 2022/02/10 N.IMAI
    strSQL = strSQL & " YOUKT_UKNO      AS 受付番号 , " & Chr(13)                                   ' 受付番号
    strSQL = strSQL & " YOUKT_UKDATE    AS 受付日   , " & Chr(13)                                   ' 受付日
    strSQL = strSQL & " CARG_KYDATE     AS 解約日   , " & Chr(13)                                   ' 解約日
    strSQL = strSQL & " YOUKT_NAME      AS 顧客名   , " & Chr(13)                                   ' 顧客名
    strSQL = strSQL & " YOUKT_KANA      AS ｶﾅ       , " & Chr(13)                                   ' 顧客名ｶﾅ
    strSQL = strSQL & " YOUKT_YCODE     AS ﾔｰﾄﾞNo   , " & Chr(13)                                   '
    strSQL = strSQL & " YARD_NAME       AS ヤード名 , " & Chr(13)                                   ' ヤード名
    strSQL = strSQL & " YOUKT_NO        AS 部屋番号 , " & Chr(13)                                   ' 部屋番号
    strSQL = strSQL & " YOUKT_MAIL      AS メール ,   " & Chr(13)                                   ' メール
    
    strSQL = strSQL & " YARD_ADDR_1 & YARD_ADDR_2 & YARD_ADDR_3     AS ヤード所在地 , " & Chr(13)   ' ヤード所在地
    strSQL = strSQL & " iif(YOUKT_TEL IS NULL, '', YOUKT_TEL)       AS 電話番号     , " & Chr(13)   ' 電話番号
    strSQL = strSQL & " iif(YOUKT_KEITAI IS NULL, '', YOUKT_KEITAI) AS 携帯番号     , " & Chr(13)   ' 携帯番号
    strSQL = strSQL & " iif(YOUKT_FAX IS NULL, '',YOUKT_FAX)        AS FAX番号        " & Chr(13)   ' FAX番号
    strSQL = strSQL & " FROM ((( dbo_YOUK_TRAN  " & Chr(13)
    strSQL = strSQL & " INNER JOIN dbo_YARD_MAST ON dbo_YOUK_TRAN.YOUKT_YCODE = dbo_YARD_MAST.YARD_CODE ) "
    
    strSQL = strSQL & "  INNER JOIN NAME_C080N  ON dbo_YOUK_TRAN.YOUKT_YUKBN = NAME_C080N.NAME_CODE ) "
    strSQL = strSQL & " LEFT JOIN dbo_CARG_FILE  ON dbo_CARG_FILE.CARG_UKNO = dbo_YOUK_TRAN.YOUKT_UKNO ) "
    strSQL = strSQL & " LEFT JOIN NAME_C088N  ON dbo_CARG_FILE.CARG_AGRE = NAME_C088N.NAME_CODE "
    strSQL = strSQL & " WHERE 1=1 "
    ' 顧客名ｶﾅ
    If Nz(Trim(Me.txt_KKana), "") <> "" Then
        strSQL = strSQL & " AND  REPLACE(REPLACE(NZ(YOUKT_KANA),' ',''),'　','') LIKE '*" & _
                    Replace(Replace(Me.txt_KKana, " ", ""), "　", "") & "*'  " & Chr(13)
    End If

    ' 顧客コード
    If Nz(Trim(Me.txt_KCode), "") <> "" Then
        strSQL = strSQL & " AND YOUKT_UCODE = " & Me.txt_KCode & Chr(13)
    End If

    ' TEL Or 携帯 Or FAX
    If Nz(Trim(Me.txt_Tel), "") <> "" Then
        strSQL = strSQL & "  AND ( REPLACE(NZ(YOUKT_TEL), '-', '') = '" & Replace(Me.txt_Tel, "-", "") & "'     " & Chr(13)
        strSQL = strSQL & "  OR  REPLACE(NZ(YOUKT_KEITAI), '-', '') = '" & Replace(Me.txt_Tel, "-", "") & "'    " & Chr(13)
        strSQL = strSQL & "  OR  REPLACE(NZ(YOUKT_FAX), '-', '') = '" & Replace(Me.txt_Tel, "-", "") & "'   )   " & Chr(13)
    End If
    
    ' メール
    If Nz(Trim(Me.txt_Mail), "") <> "" Then
        strSQL = strSQL & " AND  YOUKT_MAIL LIKE '*" & Me.txt_Mail & "*'  " & Chr(13)
    End If
    
    strSQL = strSQL & "   order by CARG_AGRE,YOUKT_YUKBN,YOUKT_UKDATE DESC "

    fncCreateSQL = True

    Exit Function

fncCreateSQL_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
End Function

'==============================================================================*
'
'        MODULE_NAME      :データ取得SQL生成
'        MODULE_ID        :txt_KKana_BeforeUpdate
'        CREATE_DATE      :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KKana_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_KKana_BeforeUpdate_Err

    vCode = Nz(Me.txt_KKana, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "顧客名カナ" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_KKana_BeforeUpdate_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Sub

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
'        pstrUKNO = List_Kokyaku.Column(5, 0)       'DELETE 2011/11/24 M.RYU
        pstrUKNO = List_Kokyaku.Column(1, 1)        'INSERT 2011/11/24 M.RYU
    Else
'        pstrUKNO = List_Kokyaku.Column(5, Me.List_Kokyaku.ListIndex)       'DELETE 2011/11/24 M.RYU
        pstrUKNO = List_Kokyaku.Column(1, Me.List_Kokyaku.ListIndex + 1)    'INSERT 2011/11/24 M.RYU
    End If

    ' 遷移元画面への受付番号セット
    Call psubSetDataFVS500

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
