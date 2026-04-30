VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS560"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 既存顧客検索
'
'        PROGRAM_NAME    : 既存顧客検索
'        PROGRAM_ID      : FVS560
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2006/02/18
'        CERATER         : T.SUZUKI
'        Ver             : 0.0
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

Private Const P_USEKANE           As Integer = 0
Private Const P_RENRAKU           As Integer = 1

'==============================================================================*
'
'        MODULE_NAME      :フォーム初期値
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2006/02/18
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim var_OpenArgs As Variant
    Dim strParam()   As String

    ' フォーム初期化
    Call subFormInit

    ' パラメータセット
    var_OpenArgs = Nz(Me.OpenArgs, "")

    If var_OpenArgs <> "" Then
        strParam = Split(var_OpenArgs, ",")

        Me.txt_KKana = strParam(P_USEKANE)  ' 顧客名カナ
        Me.txt_Tel = strParam(P_RENRAKU)    ' 連絡先

        If (Nz(Me.txt_KKana, "") <> "" Or Nz(Me.txt_Tel, "") <> "") Then
            ' 検索実行
            Call cmd_Search_Click
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :電話番号
'        MODULE_ID        :txt_TEL_BeforeUpdate
'        CREATE_DATE      :2006/02/18
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
'        CREATE_DATE      :2006/02/18
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
'        CREATE_DATE      :2006/02/18
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
            MsgBox "検索条件に該当する顧客データはありません。" & vbCrLf & "検索条件を変えて再検索か、キャンセルボタンで戻ってください。", vbInformation, strMsg_kaku
            Me.txt_KKana.SetFocus

'        Case 1  ' *** 該当データ1件
'            ' 遷移元画面の戻り値設定
'            Call subRtnMotoGamen

        Case Else
            '  複数件データあり
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
'        CREATE_DATE      :2006/02/18
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

    ' 遷移元画面の戻り値設定
    Call subRtnMotoGamen
End Sub

'==============================================================================*
'
'        MODULE_NAME      :キャンセル
'        MODULE_ID        :Cmd_Cansel_Click
'        CREATE_DATE      :2006/02/18
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
'        CREATE_DATE      :2006/02/18
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFormInit()

    With Me
        .txt_KKana = ""   ' 顧客名ｶﾅ
        .txt_Tel = ""     ' 電話番号 Or 携帯番号 Or FAX番号

        .cmd_Decision.Enabled = False
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2006/02/18
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    fncInputCheck = False

    If Nz(Trim(Me.txt_KKana), "") = "" And _
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

    Dim strWhere As String

    On Error GoTo fncCreateSQL_Err

    fncCreateSQL = False

    strSQL = ""
    strSQL = " SELECT   USER_NAME, " & Chr(13)                                                      ' 顧客名
    strSQL = strSQL & " USER_KANA, " & Chr(13)                                                      ' 顧客名ｶﾅ
    strSQL = strSQL & " iif(USER_TEL IS NULL, '', 'TEL:' & USER_TEL) AS TEL, " & Chr(13)            ' 電話番号
    strSQL = strSQL & " iif(USER_KEITAI IS NULL, '', '携帯:' & USER_KEITAI) AS KEITAI, " & Chr(13)  ' 携帯番号
    strSQL = strSQL & " iif(USER_FAX IS NULL, '','FAX:' & USER_FAX) AS FAX, " & Chr(13)             ' FAX番号
    strSQL = strSQL & " USER_ADR_1 & USER_ADR_2 & USER_ADR_3, " & Chr(13)                           ' 住所
    strSQL = strSQL & " USER_CODE " & Chr(13)                                                       ' 顧客名コード
    strSQL = strSQL & " FROM dbo_USER_MAST " & Chr(13)

    ' 抽出条件設定
    strWhere = ""

    ' 顧客名ｶﾅ
    If Nz(Trim(Me.txt_KKana), "") <> "" Then
        strWhere = " WHERE Replace(Replace(USER_KANA,' ',''), '　','') LIKE '*" & Replace(Replace(Me.txt_KKana, " ", ""), "　", "") & "*'" & Chr(13)
    End If

    ' TEL Or 携帯 Or FAX
    If Nz(Trim(Me.txt_Tel), "") <> "" Then
        If Nz(strWhere, "") <> "" Then
            ' 上記で条件が指定されている場合
            strWhere = strWhere & " AND "
        Else
            ' 上記で条件が指定されていない場合
            strWhere = " WHERE "
        End If

        strWhere = strWhere & "     (Replace(USER_TEL,'-','')    = '" & Replace(Me.txt_Tel, "-", "") & "' " & Chr(13)
        strWhere = strWhere & "  OR  Replace(USER_KEITAI,'-','') = '" & Replace(Me.txt_Tel, "-", "") & "' " & Chr(13)
        strWhere = strWhere & "  OR  Replace(USER_FAX,'-','')    = '" & Replace(Me.txt_Tel, "-", "") & "') " & Chr(13)
    End If

    strSQL = strSQL & strWhere

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
'        MODULE_ID        :subRtnMotoGamen
'        CREATE_DATE      :2006/02/18
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subRtnMotoGamen()

    Dim strUserCd As String

    strUserCd = ""

    ' 顧客コードセット
    If Me.List_Kokyaku.ListIndex = -1 Then
        strUserCd = List_Kokyaku.Column(6, 0)
    Else
        strUserCd = List_Kokyaku.Column(6, Me.List_Kokyaku.ListIndex)
    End If

    ' 遷移元画面への受付番号セット
    Call UserReflection(strUserCd)

    ' 画面を閉じる
    doCmd.Close acForm, Me.FormName
End Sub
'****************************  ended or program ********************************
