VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　紹介制限ヤード状況一覧画面
'   プログラムＩＤ　：　FVS200
'   作　成　日　　　：  2007/03/01
'   作　成　者　　　：  tajima
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'   Ver             ：
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   Private定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_WK_TABLE_NM           As String = "FVS200_W01"  ' ワークテーブル名
Private Const P_SUBFRM_NM             As String = "FVS200_S01"  ' サブフォーム名

Private Const P_MSG_ERR               As String = "エラー"
Private Const P_MSG_KAKUNIN           As String = "確認"

Private Const P_MAX_YARD_NO           As String = "aaaaaa"
'==============================================================================*
'   Private変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private pstrBumonCd                   As String   ' 部門コード

Private pobjKONT_DB                   As Database ' コンテナDBオブジェクト
Private pobjKASE_DB                   As Database ' 加瀬DBオブジェクト

Private pstrYardCode                  As String   ' パラメータで渡ってきたヤードコード
Private pstrFormCaption               As String   ' この画面の標題
Private pstrUserName                  As String   ' ユーザID
Private pIsAuthority                  As Boolean  ' FVS200更新可否権限
Private pIsChanged                    As Boolean  ' 変更可否フラグ

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FunctionKeyDown(KeyCode, Shift)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load イベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim errorMessage As String
    
  On Error GoTo Exception
    errorMessage = ""
    
     ' 初期処理
    errorMessage = initializFunction()
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation, P_MSG_ERR
        doCmd.Close acForm, Me.FormName
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("Form_Load", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open イベント
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim strParam()  As String
    Dim iParamCnt   As Integer
    
  On Error GoTo Exception

    ' 起動パラメータの取得
    strParam = Split(Nz(Me.OpenArgs), ",")
    
    ' 起動パラメータは以下のパターンがある
    '①なし
    '利用状況からきた場合
    '② ヤードコード
    
    ' パラメータ数チェックと取込
    iParamCnt = UBound(strParam)
    
    
    If iParamCnt < 0 Then
        pstrYardCode = ""
    Else
        pstrYardCode = strParam(0)
    End If
    
    Exit Sub
        
Exception:
    Call MSZZ024_M00("Form_Open", True)
    Cancel = True
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_End_Click イベント
'        MODULE_ID        :cmd_End_Click
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

  On Error GoTo Exception
  
    ' 確認メッセージ表示
    If pIsChanged = True Then
    
        If MsgBox("修正中ですが保存せずに終了しますか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
    End If

    ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙﾃﾞｰﾀ削除
    Call clearFVS200_W01

    ' 本画面を閉じる
    doCmd.Close acForm, Me.FormName
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_End_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_印刷_Click イベント
'        MODULE_ID        :cmd_印刷_Click
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_印刷_Click()
    
  On Error GoTo Exception
  
    If Nz(Me.cmb_ヤードコード.VALUE, "") = "" Then
        MsgBox "ヤードを選択して下さい。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
      
    If pIsChanged = True Then
        If MsgBox("修正中の為、一度保存してから印刷プレビューを行います。" & vbCr & "よろしいですか？", _
                vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
        ' ☆保存処理の実行
        Call sub保存処理
        Call clearChangeFlgFVS200_W01
    Else
        If MsgBox("画面内容にて印刷プレビューを行います。" & vbCr & "よろしいですか？", _
                    vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' 印刷の実行
    If Print200.fncPrintYardInlimit(Print200.P901_MODE_PREVIEW, Me.cmb_ヤードコード.VALUE) = False Then
           MsgBox "印刷対象となるデータがありません。", vbInformation, P_MSG_KAKUNIN
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_印刷_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Excel_Click イベント
'        MODULE_ID        :cmd_Excel_Click
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

  On Error GoTo Exception
  
    If Nz(Me.cmb_ヤードコード.VALUE, "") = "" Then
        MsgBox "ヤードを選択して下さい。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
      
    If pIsChanged = True Then
        If MsgBox("修正中の為、一度保存してからExcel出力を行います。" & vbCr & "よろしいですか？", _
                vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
        ' ☆保存処理の実行
        Call sub保存処理
        Call clearChangeFlgFVS200_W01
    Else
        If MsgBox("画面内容にてExcel出力を行います。" & vbCr & "よろしいですか？", _
                    vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' 印刷の実行
    If Print200.fncPrintYardInlimit(Print200.P901_MODE_EXCEL, Me.cmb_ヤードコード.VALUE) = False Then
           MsgBox "Excel出力対象となるデータがありません。", vbInformation, P_MSG_KAKUNIN
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_Excel_Click", True)
End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_Click イベント
'        MODULE_ID        :cmd_保存_Click
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()

  On Error GoTo Exception

    If pIsChanged = False Then
        MsgBox "修正はされておりません｡", vbInformation, P_MSG_KAKUNIN
        Exit Sub
    End If

    If MsgBox("画面内容を保存します。" & vbCr & "よろしいですか？", _
            vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If

    ' 保存実行
    Call sub保存処理
    
    If MsgBox("終了します。" & vbCr & "よろしいですか？", _
            vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbYes Then
        Call clearFVS200_W01
        doCmd.Close acForm, Me.FormName    ' 本画面を閉じる
        Exit Sub
    End If
    
    ' 終了しないならワークの変更フラグだけハズしたい
    Call clearChangeFlgFVS200_W01
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_保存_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_取消_Click イベント
'        MODULE_ID        :cmd_取消_Click
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()

  On Error GoTo Exception
  
    If pIsChanged = True Then

        If MsgBox("修正内容を取消ます。" & vbCr & "よろしいですか？", _
                vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
            Exit Sub
        End If
        
        ' データの読み直し
        Call setAllData
      
    End If
      
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_取消_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤードコード_AfterUpdate イベント
'        MODULE_ID        :cmb_ヤードコード_AfterUpdate
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_AfterUpdate()

    Dim varCode As Variant

  On Error GoTo Exception

    ' ヤードコード取得
    varCode = Nz(Me.cmb_ヤードコード, "")

    If varCode = "" Then Exit Sub

    Me.cmb_ヤードコード = Format(varCode, "000000")

    ' ヤードコードを元に全データを読込、画面に設定する
    Call setAllData
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_ヤードコード_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤードコード_BeforeUpdate イベント
'        MODULE_ID        :cmb_ヤードコード_BeforeUpdate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant

  On Error GoTo Exception

    ' ヤードコード取得
    varCode = Nz(Me.cmb_ヤードコード, "")

    ' 入力チェック
    Select Case fncCheckNum(varCode, 6)
        Case 0
            ' 正常時
        Case 2
            ' 桁数エラー
            MsgBox "桁数オーバーしています｡", vbInformation, P_MSG_ERR
            Cancel = True: SendKeys "{ESC}"
            Exit Sub
        Case 9
            ' その他のエラー
            MsgBox "ヤードコードを入力してください。", vbInformation, P_MSG_ERR
            Cancel = True: SendKeys "{ESC}"
            Exit Sub
    End Select

    ' ヤードマスタチェック(存在チェック)
    If isExistYardCode(varCode) = False Then
        MsgBox "該当するヤードマスタが登録がされておりません。", vbInformation, P_MSG_KAKUNIN
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_ヤードコード_BeforeUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤードコード_KeyPress イベント
'        MODULE_ID        :cmb_ヤードコード_KeyPress
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_ヤードコード, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :数値項目のチェック
'        MODULE_ID        :fncCheckNum
'        PARAMETER        :varValue(IN)  = 判定する値
'                          intLenNum(IN) = 整数桁
'                          varLenDec(IN) = 小数桁(省略された時は、小数点桁数無し)
'        RETURN           :
'                          1 = 数値チェックエラー
'                          2 = 桁溢れエラー
'                          9 = 上記以外のエラー
'                          0 = 正常時
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckNum(ByVal varValue As Variant, ByVal intLenNum As Integer, _
                             Optional ByVal varLenDec As Variant) As Integer

    Dim intCount    As Integer
    Dim varRtn      As Variant
    Dim varNumeric  As Variant
    Dim varDecimal  As Variant
    Dim dblValue    As Double
    Dim dblValue_F  As Double
    Dim dblValue_T  As Double

  On Error GoTo Exception

    fncCheckNum = 9

    ' 数値とチェックエラー
    If Nz(CmNumeric(varValue), "") = "" Then
        fncCheckNum = 1: GoTo fncCheckNum_Exit
    End If

    ' 全角チェックエラー
    If CmChkCher(CStr(Nz(varValue)), "A") = False Then
        fncCheckNum = 1: GoTo fncCheckNum_Exit
    End If

    ' 小数点指定時のチェック
    varRtn = Nz(InStr(Abs(varValue), "."), 0)

    If IsMissing(varLenDec) = False Then
        If varRtn <> 0 Then
            If varLenDec < CmLenb(Abs(varValue)) - varRtn Then
                ' 小数点桁数が設定桁数を超えているエラー
                fncCheckNum = 2: GoTo fncCheckNum_Exit
            End If
        End If

    Else
        If CmLenb(Abs(CmNumeric(varValue))) > intLenNum Then
            ' 文字数超過のためエラー
            fncCheckNum = 2: GoTo fncCheckNum_Exit
        End If

        If varRtn <> 0 Then
            ' 小数点の入力があるのでエラー
            fncCheckNum = 2: GoTo fncCheckNum_Exit
        End If
    End If

    ' 桁チェック
    intCount = 0
    For intCount = 1 To intLenNum
        varNumeric = varNumeric & "9"
    Next intCount

    If IsMissing(varLenDec) = False Then
        varNumeric = ".": intCount = 0
        For intCount = 1 To varLenDec
            varDecimal = varDecimal & "9"
        Next intCount
    End If

    varNumeric = varNumeric & varDecimal
    varNumeric = CmNumeric(varNumeric)

    dblValue = CmNumeric(varValue)
    dblValue_F = varNumeric * -1
    dblValue_T = varNumeric

    If (dblValue_F > dblValue) Or (dblValue_T < dblValue) Then
        ' 論理範囲超過のためエラー
        fncCheckNum = 2: GoTo fncCheckNum_Exit
    End If

    fncCheckNum = 0

fncCheckNum_Exit:
    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncCheckNum" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :文字列のﾁｪｯｸ
'        MODULE_ID        :fncCheckChar
'        PARAMETER        :varValue = 処理対象Text
'                          intLen   = 制限文字数
'        REYURN           :
'                          2 = 桁溢れエラー
'                          0 = 正常時
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckChar(ByVal varValue As Variant, ByVal intLen As Integer) As Integer

    On Error GoTo Exception

    fncCheckChar = 9

    If CmLenb(varValue) > intLen Then
        fncCheckChar = 2
    Else
        fncCheckChar = 0
    End If

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncCheckChar" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :Textエリア入力文字数ﾁｪｯｸ
'        MODULE_ID        :fncCheckText
'        PARAMETER        :ctlTxt(IN)    = 処理対象Text
'                          lngLength(IN) = 制限文字数
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckText(ByRef ctlTxt As TextBox, ByVal lngLength As Long) As Boolean

    Dim strwork As String

  On Error GoTo Exception

    fncCheckText = False
    strwork = Nz(ctlTxt, "")

    Select Case fncCheckChar(strwork, lngLength)
        Case 0
            ' 正常時
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_MSG_ERR
            Exit Function

        Case 9
            ' その他のエラー
            MsgBox "内容を正しく入力してください。", vbInformation, P_MSG_ERR
            Exit Function
    End Select

    fncCheckText = True
    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncCheckText" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :KASE_DB接続
'        MODULE_ID        :connectKaseDB
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub connectKaseDB()

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo Exception

    strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME")

    If strDataSource = "" Then
        ' テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If

    Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "connectKaseDB" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :KONT_DB接続
'        MODULE_ID        :connectKontDB
'        PARAMETER        :a部門CD
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub connectKontDB(ByVal a部門CD As String)

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo Exception

    strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME_" & a部門CD)
    If strDataSource = "" Then
        ' テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If

    ' DB接続
    Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(CStr(a部門CD)))
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "connectKontDB" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ODBC接続文字列取得
'        MODULE_ID        :fncGetDataSource
'        PARAMETER        :strODBCNAME(IN) ODBC名
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetDataSource(ByVal strODBCNAME As String) As String

    Dim strName As String

    On Error GoTo fncGetDataSource_Err

    fncGetDataSource = ""

    strName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & strODBCNAME & "'"))

    fncGetDataSource = strName

    Exit Function

fncGetDataSource_Err:
    Call Err.Raise(Err.Number, "fncGetDataSource" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :KASE_DB切断
'        MODULE_ID        :disConnectKaseDB
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub disConnectKaseDB()

  On Error GoTo Exception

    pobjKASE_DB.Close
    Set pobjKASE_DB = Nothing

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "disConnectKaseDB" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :KONT_DB切断
'        MODULE_ID        :disConnectKontDB
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub disConnectKontDB()

  On Error GoTo Exception

    pobjKONT_DB.Close
    Set pobjKONT_DB = Nothing

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "disConnectKontDB" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :部門コード取得
'        MODULE_ID        :fncGetBumonName
'        PARAMETER        :
'        RETURN           :部門コード
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBumonName() As String

    Dim objDb       As Database
    Dim strSQL      As String
    Dim objRs       As Recordset
    Dim strBumonCd  As String
    Dim blnErrFlg   As Boolean

    On Error GoTo fncGetBumonName_Err

    strBumonCd = ""

    Set objDb = CurrentDb

    ' 部門コード取得
    strSQL = "SELECT BUMOM_BUMOC "             ' 部門コード
    strSQL = strSQL & " FROM  dbo_CONT_MAST "  ' コントロールマスタ(コンテナDB)
    strSQL = strSQL & "      ,BUMO_MAST "      ' 部門マスタ(加瀬DB)
    strSQL = strSQL & " WHERE CONT_BUMOC = BUMOM_BUMOC "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            strBumonCd = Nz(.Fields("BUMOM_BUMOC"), "")
        End If
    End With

    fncGetBumonName = strBumonCd

    GoTo fncGetBumonName_Exit

fncGetBumonName_Err:
    blnErrFlg = True

fncGetBumonName_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If

    ' カレントDBクローズ
    objDb.Close
    Set objDb = Nothing

    If blnErrFlg = True Then
        ' エラー発生時
        Call Err.Raise(Err.Number, "fncGetBumonName" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタ存在可否
'        MODULE_ID        :isExistYardCode
'        PARAMETER        :ヤードコード
'        RETURN           :True = 存在するコード／False = 存在しないコード
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isExistYardCode(ByVal aYardCode As Variant) As Boolean

    Dim varRtn     As Variant
    Dim strSQL     As String

  On Error GoTo Exception

    isExistYardCode = False

    ' ヤードコード未入力はエラー
    If Nz(aYardCode, "") = "" Then Exit Function

    ' 存在チェック
    strSQL = " YARD_CODE = " & Val(aYardCode)
    varRtn = Nz(DCount("YARD_CODE", "dbo_YARD_MAST", strSQL), "")

    If varRtn < 1 Then
        Exit Function
    Else
        isExistYardCode = True 'そんざいスマス
    End If

    Exit Function

Exception:
    ' エラー発生時
    Call Err.Raise(Err.Number, "isExistYardCode" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタ情報を画面に設定
'        MODULE_ID        :getYardMastData
'        PARAMETER        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数  ヤードコード
'        RETURN           :True = 正常終了／False = 異常終了
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function setYardMastData(dbSQLServer As DAO.Database, _
                                 ByVal aYardCode As Variant) As Boolean

    Dim varCode    As Variant
    Dim varRtn     As Variant
    Dim sqlText    As String
    Dim objRs      As Recordset
    Dim isException  As Boolean

  On Error GoTo Exception

    setYardMastData = False
    isException = False

      ' *** ヤードマスタデータ取得 ***
      sqlText = "  SELECT   YARD_NAME "         ' ヤード名
      sqlText = sqlText & " ,YARD_ADDR_1 "      ' ヤード住所１
      sqlText = sqlText & " ,YARD_ADDR_2 "      ' ヤード住所２
      sqlText = sqlText & " ,YARD_ADDR_3 "      ' ヤード住所３
      sqlText = sqlText & " ,YARD_BEGIN_DAY "   ' 営業開始日
      sqlText = sqlText & " ,YARD_END_DAY "     ' ヤード解約予定日
      sqlText = sqlText & " ,YARD_RENTEND_DAY " ' 営業終了日
      sqlText = sqlText & "  FROM YARD_MAST "
      sqlText = sqlText & " WHERE YARD_CODE = " & Val(aYardCode)
      Set objRs = dbSQLServer.OpenRecordset(sqlText, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

      If Not (objRs.BOF And objRs.EOF) Then
          Me.txt_ヤード名 = objRs.Fields("YARD_NAME")
          Me.txt_ヤード住所 = objRs.Fields("YARD_ADDR_1") & _
                              objRs.Fields("YARD_ADDR_2") & _
                              objRs.Fields("YARD_ADDR_3")
          Me.txt_営業開始日 = objRs.Fields("YARD_BEGIN_DAY")
          Me.txt_ヤード解約予定日 = objRs.Fields("YARD_END_DAY")
          Me.txt_営業終了日 = objRs.Fields("YARD_RENTEND_DAY")
          setYardMastData = True '取得OK
      End If

    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "setYardMastData" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :テーブルの存在チェック
'        MODULE_ID        :fncCHKTableExist
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isExistTable(strTableName As String) As Boolean

    Dim objDb        As Database
    Dim tdf          As TableDef
    Dim isException  As Boolean

  On Error GoTo Exception
    isExistTable = False
    isException = False

    Set objDb = CurrentDb

    For Each tdf In objDb.TableDefs
        If tdf.NAME = strTableName Then
            isExistTable = True
            Exit For
        End If
    Next tdf

    GoTo Finally

Exception:
    isException = True

Finally:
    objDb.Close
    Set objDb = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "isExistTable" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ作成
'        MODULE_ID        :createWorkTable
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createWorkTable() As Boolean

    Dim objDb       As Database
    Dim tdfNew      As TableDef
    Dim intCount    As Integer
    Dim isException As Boolean

    On Error GoTo Exception
    
    createWorkTable = False
    isException = False

    Set objDb = CurrentDb

    Set tdfNew = objDb.CreateTableDef(P_WK_TABLE_NM)
    With tdfNew
        .Fields.Append .CreateField("FKS200_YCODE", dbText, 6)      ' ヤードコード
        .Fields.Append .CreateField("FKS200_YNAME", dbText, 36)     ' ヤード名称
        .Fields.Append .CreateField("FKS200_KIRO", dbDouble)        ' 距離
        .Fields.Append .CreateField("FKS200_ADDR", dbText, 108)     ' 住所
        .Fields.Append .CreateField("FKS200_CHECK", dbBoolean)      ' 制限済み
        .Fields.Append .CreateField("FKS200_LIMITDATE", dbDate)     ' 制限日
        .Fields.Append .CreateField("FKS200_LYCODE", dbText, 6)     ' 制限元ヤードコード
        .Fields.Append .CreateField("FKS200_LYNAME", dbText, 50)    ' 制限元ヤード名
        .Fields.Append .CreateField("FKS200_CHANGE", dbBoolean)     ' 修正可否

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With

    ' テーブルの追加
    objDb.TableDefs.Append tdfNew

    ' ｻﾌﾞﾌｫｰﾑのｿｰｽｵﾌﾞｼﾞｪｸﾄ設定
    Me.FVS200_FVS200S01.SourceObject = P_SUBFRM_NM
    createWorkTable = True
    GoTo Finally

Exception:
    isException = True

Finally:
    objDb.Close
    Set objDb = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "createWorkTable" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :初期処理（画面オープン時に一度だけ実行する）
'        MODULE_ID        :initializFunction
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'        NOTE             :戻り値としてエラーメッセージを返す。""なら正常完了
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function initializFunction() As String

    Dim isException As Boolean
    Dim errorMessage As String
    
  On Error GoTo Exception
    isException = False
    errorMessage = ""
    
    'ユーザID取得
    pstrUserName = ""
    pstrUserName = MSZZ000.LsGetUserName
    
    ' 部門コード取得
    pstrBumonCd = fncGetBumonName
    
    ' 加瀬DBに接続
    Call connectKaseDB

    ' 更新権限取得を行う
    pIsAuthority = isFVS200Authority(pobjKASE_DB, pstrUserName)
    ' 更新権限によってボタンを隠したりサブフォームを編集不可にす
    If pIsAuthority Then
        Me.cmd_保存.Visible = True
        Me.cmd_取消.Visible = True
        Me.FVS200_FVS200S01.Locked = False
    Else
        Me.cmd_保存.Visible = False
        Me.cmd_取消.Visible = False
        Me.FVS200_FVS200S01.Locked = True
        Me.Caption = Me.Caption & "（参照モード）"
    End If
    
    ' 変更フラグ初期化
    pIsChanged = False
      
    ' 画面の標題格納
    pstrFormCaption = Me.Caption
      
    ' 画面設定など....
    ' メニューから呼ばれたとき
    If pstrYardCode = "" Then
        Call clearFVS200_W01
        Me.FVS200_FVS200S01.Requery
    Else
    ' 他画面から呼ばれたとき...引数のヤードコードで画面表示
        Me.cmb_ヤードコード = Format(pstrYardCode, "000000")
        Call setAllData
    End If

    GoTo Finally
    
Exception:
    isException = True
    errorMessage = "system error"

Finally:
    If Not pobjKASE_DB Is Nothing Then Call disConnectKaseDB
    initializFunction = errorMessage
    If isException = True Then
        Call Err.Raise(Err.Number, "initializFunction" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ削除
'        MODULE_ID        :clearFVS2600_W01
'        PARAMETER        :
'        CREATE_DATE      :2006/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub clearFVS200_W01()

    Dim strSQL      As String
    Dim objDb       As Database
    Dim isException As Boolean

  On Error GoTo Exception
    isException = False
    
    Set objDb = CurrentDb

    strSQL = "DELETE FROM " & P_WK_TABLE_NM
    objDb.Execute strSQL, dbFailOnError

    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "clearFVS2600_W01" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ作成
'        MODULE_ID        :createWorkData
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function createWorkData() As Boolean

    Dim objDb      As Database
    Dim sqlText    As String
    Dim readRs     As Recordset
    Dim writeRs    As Recordset
    Dim isException  As Boolean

  On Error GoTo Exception
    isException = False
    createWorkData = False
    
    ' ワークテーブルの為のローカルに接続
    Set objDb = CurrentDb

    ' KONT_DBから読み込む
    sqlText = "SELECT NYAR_NCODE"
    sqlText = sqlText & " ,YARD_MAST.YARD_NAME " & Chr(13)
    sqlText = sqlText & " ,NYAR_KIRO " & Chr(13)
    sqlText = sqlText & " ,YARD_MAST.YARD_ADDR_1 " & Chr(13)
    sqlText = sqlText & " ,YARD_MAST.YARD_ADDR_2 " & Chr(13)
    sqlText = sqlText & " ,YARD_MAST.YARD_ADDR_3 " & Chr(13)
    sqlText = sqlText & " ,YARD_MAST.YARD_INLIMIT_DAY " & Chr(13)
    sqlText = sqlText & " ,YARD_MAST.YARD_INLIMIT_YCODE " & Chr(13)
    sqlText = sqlText & " ,INLIMIT_YARD.YARD_NAME NYAR_NAME " & Chr(13)
    sqlText = sqlText & " FROM NYAR_MAST " & Chr(13)
    sqlText = sqlText & "   INNER JOIN YARD_MAST ON " & Chr(13)
    sqlText = sqlText & "       YARD_MAST.YARD_CODE = NYAR_NCODE " & Chr(13)
    sqlText = sqlText & "   LEFT OUTER JOIN YARD_MAST AS INLIMIT_YARD ON " & Chr(13)
    sqlText = sqlText & "       INLIMIT_YARD.YARD_CODE = YARD_MAST.YARD_INLIMIT_YCODE " & Chr(13)
    sqlText = sqlText & " WHERE " & Chr(13)
    sqlText = sqlText & "       NYAR_YCODE = " & Me.cmb_ヤードコード.VALUE & Chr(13)
    sqlText = sqlText & "   AND NYAR_NCODE <> " & Me.cmb_ヤードコード.VALUE & Chr(13)
    sqlText = sqlText & "   AND ISNULL(YARD_MAST.YARD_NONDISP_DAY,'9999/12/31') > GETDATE()"
    sqlText = sqlText & " ORDER BY NYAR_KIRO"
    Set readRs = pobjKONT_DB.OpenRecordset(sqlText, dbOpenSnapshot, dbSQLPassThrough)

    If readRs.EOF = False Then
        ' 書込先となるワークテーブルを開く
        sqlText = "SELECT * FROM " & P_WK_TABLE_NM
        Set writeRs = objDb.OpenRecordset(sqlText, dbOpenDynaset, dbAppendOnly)

        Do Until readRs.EOF
            writeRs.AddNew
            ' 一般的な情報を設定
            writeRs.Fields("FKS200_YCODE").VALUE = Format$(readRs.Fields("NYAR_NCODE"), "000000")
            writeRs.Fields("FKS200_YNAME").VALUE = readRs.Fields("YARD_NAME")
            writeRs.Fields("FKS200_KIRO").VALUE = readRs.Fields("NYAR_KIRO")
            writeRs.Fields("FKS200_ADDR").VALUE = Nz(readRs.Fields("YARD_ADDR_1"), "") & _
                                                  Nz(readRs.Fields("YARD_ADDR_2"), "") & _
                                                  Nz(readRs.Fields("YARD_ADDR_3"), "")
            ' 制限があるならば制限情報を設定
            If Nz(readRs.Fields("YARD_INLIMIT_DAY"), "") <> "" Then
                writeRs.Fields("FKS200_CHECK").VALUE = True
                writeRs.Fields("FKS200_LIMITDATE").VALUE = readRs.Fields("YARD_INLIMIT_DAY")
                writeRs.Fields("FKS200_LYCODE").VALUE = Format$(readRs.Fields("YARD_INLIMIT_YCODE"), "000000")
                writeRs.Fields("FKS200_LYNAME").VALUE = Nz(readRs.Fields("NYAR_NAME"), "") & "にて停止"
            Else
                writeRs.Fields("FKS200_CHECK").VALUE = False
            End If
            writeRs.Fields("FKS200_CHANGE").VALUE = False           'レコード修正有無
            writeRs.UPDATE

            ' DBの内容を次に
            readRs.MoveNext
        Loop
        createWorkData = True
    Else
        ' 該当データなし
    End If
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not readRs Is Nothing Then readRs.Close: Set readRs = Nothing
    If Not writeRs Is Nothing Then writeRs.Close: Set writeRs = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "createWorkData" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :紹介制限権限可否
'        MODULE_ID        :isFVS200Authority
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = ユーザID
'        戻り値           : True...更新可
'                         : False...更新不可
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isFVS200Authority(dbSQLServer As DAO.Database, _
                                   aユーザID As String _
                                  ) As Boolean
                              
    Dim strSQL     As String
    Dim rsObject   As Recordset
     
    isFVS200Authority = False
    
    On Error GoTo Exception
  
    ' オンライン管理テーブルから権限取得
    strSQL = "SELECT ONLET_ACCEI FROM ONLE_TABL WHERE ONLET_PROGB = 'FVS200'"
    strSQL = strSQL & "AND ONLET_USERB = '" & aユーザID & "'"
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    ' 対象確認
    If rsObject.EOF = False Then
      ' オンライン更新区分が"1"ならば更新権限がある
      If "1" = Nz(rsObject.Fields("ONLET_ACCEI"), "") Then
        isFVS200Authority = True
      End If
    End If
    rsObject.Close
    Set rsObject = Nothing
   
    Exit Function

Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "isFVS200Authority" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :全データを読込、画面に設定する
'        MODULE_ID        :setAllData
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setAllData()

    Dim isException As Boolean
    
  On Error GoTo Exception
    isException = False

    ' KONT_DBに接続
    Call connectKontDB(pstrBumonCd)

    ' ヤードの情報を設定
    If setYardMastData(pobjKONT_DB, Me.cmb_ヤードコード.VALUE) = False Then
        Call MSZZ024_M10("setAllData", "ヤード情報の取得と画面設定が失敗しました。")
    End If

    ' ワークテーブルのクリア
    Call clearFVS200_W01
    
    ' ワークテーブルにデータを読み込む
    If createWorkData = False Then
        MsgBox "このヤードには近隣ヤードが登録がされておりません。", vbInformation, P_MSG_KAKUNIN
    End If

    Me.FVS200_FVS200S01.Requery

    ' 変更可否は初期化
    pIsChanged = False
    Me.Caption = pstrFormCaption
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not pobjKONT_DB Is Nothing Then Call disConnectKontDB
    If isException = True Then
        Call Err.Raise(Err.Number, "setAllData" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :全データを読込、画面に設定する
'        MODULE_ID        :sub保存処理()
'        PARAMETER        :
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub保存処理()

    Dim isException As Boolean
    
  On Error GoTo Exception
    isException = False

    ' KONT_DBに接続
    Call connectKontDB(pstrBumonCd)

    ' 画面内容を反映ＤＢに反映
    Call writeYARD_MAST(pobjKONT_DB, Me.cmb_ヤードコード.VALUE)
    
    ' 変更可否は初期化
    pIsChanged = False
    Me.Caption = pstrFormCaption
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not pobjKONT_DB Is Nothing Then Call disConnectKontDB
    If isException = True Then
        Call Err.Raise(Err.Number, "sub保存処理" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタの書込
'        MODULE_ID        :writeYARD_MAST
'        NOTE             :変更のあったヤードに対し紹介制限を行うか解除
'        CREATE_DATE      :2007/03/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub writeYARD_MAST(dbSQLServer As DAO.Database, _
                           ByVal aヤードコード As String _
                          )

    Dim objDb         As Database
    Dim rsAccess      As Recordset  ' ワークテーブルレコードセット
    Dim rsSqlServer   As Recordset  ' ＳＱＬサーバテーブルのレコードセット
    Dim sqlText       As String
    Dim serverYardCode As String
    Dim accessYardCode  As String
    Dim isException   As Boolean
    
    On Error GoTo Exception
    isException = False
    
    ' ワークテーブルの為のローカルに接続
    Set objDb = CurrentDb
    
    ' ヤードマスタから更新対象に成りうるレコードを取得する
    sqlText = " SELECT * FROM YARD_MAST WHERE YARD_CODE IN  " & Chr(13)
    sqlText = sqlText & "( SELECT NYAR_NCODE FROM NYAR_MAST "
    sqlText = sqlText & "  WHERE NYAR_YCODE = " & aヤードコード
    sqlText = sqlText & "    AND YARD_CODE <> " & aヤードコード & ")"
    sqlText = sqlText & " ORDER BY YARD_CODE "
    Set rsSqlServer = pobjKONT_DB.OpenRecordset(sqlText, dbOpenDynaset)

    ' ワークテーブルの変更されたデータの抽出
    sqlText = " SELECT * FROM " & P_WK_TABLE_NM & " WHERE FKS200_CHANGE =  -1 ORDER BY FKS200_YCODE "
    Set rsAccess = objDb.OpenRecordset(sqlText, dbReadOnly)

    ' ワークテーブルの変更数分処理を行う
    accessYardCode = ""
    serverYardCode = ""
    Do Until (rsAccess.EOF)
        ' ワークテーブルのキー作成
        If rsAccess.EOF = True Then
            accessYardCode = P_MAX_YARD_NO
        Else
            accessYardCode = rsAccess.Fields("FKS200_YCODE") ' ゼロサプレスのテキストで設定済み
        End If

        ' ヤードマスタのキー作成
        If rsSqlServer.EOF = True Then
            serverYardCode = P_MAX_YARD_NO
        Else
            serverYardCode = Format$(rsSqlServer.Fields("YARD_CODE"), "000000")
        End If

        ' ワークテーブルとヤードマスタのデータ比較
        Select Case StrComp(accessYardCode, serverYardCode, vbBinaryCompare)
            Case 0  'accessYardCode = serverYardCode
                ' ワークテーブルのデータとヤードマスタのデータが一致する場合
                ' 修正があった場合のみ更新をする
                If rsAccess.Fields("FKS200_CHANGE") = True Then
                    With rsSqlServer
                        .Edit
                        .Fields("YARD_INLIMIT_DAY") = rsAccess.Fields("FKS200_LIMITDATE")
                        .Fields("YARD_INLIMIT_YCODE") = rsAccess.Fields("FKS200_LYCODE")
                        .Fields("YARD_UPDATE") = DATE
                        .UPDATE
                    End With
                End If
                ' → 次レコードへ
                rsAccess.MoveNext     ' ワークテーブル
                rsSqlServer.MoveNext  ' ヤードマスタ

            Case -1  'accessYardCode < serverYardCode
                ' ワークテーブルのデータがヤードマスタに存在しない場合
                ' ※あり得ぬけど
                ' → 次レコードへ
                rsAccess.MoveNext    ' ワークテーブルのみ
  
            Case 1 ' accessYardCode > serverYardCode
              ' ヤードマスタのデータがワークテーブルに存在しない場合
               '  そのヤードには変更がなかったんよね
                ' → 次レコードへ
               rsSqlServer.MoveNext   ' ヤードマスタのみ
        End Select
    Loop

    GoTo Finally

Exception:
    isException = True

Finally:
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "writeYARD_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙの変更フラグを初期化
'        MODULE_ID        :clearChangeFlgFVS200_W01
'        PARAMETER        :
'        CREATE_DATE      :2006/03/01
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub clearChangeFlgFVS200_W01()

    Dim strSQL      As String
    Dim objDb       As Database
    Dim isException As Boolean

  On Error GoTo Exception
    isException = False
    
    Set objDb = CurrentDb

    strSQL = "UPDATE " & P_WK_TABLE_NM & " SET FKS200_CHANGE = False "
    objDb.Execute strSQL, dbFailOnError

    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "clearChangeFlgFVS200_W01" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

' ☆☆☆ Public関数群 ☆☆☆
'==============================================================================*
'
'        MODULE_NAME      :ファンクションキー押下処理
'        MODULE_ID        :FunctionKeyDown
'        PARAMETER        :
'        NOTE             :サブフォームからも呼ばれる
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub FunctionKeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
              
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_End.SetFocus

        Case vbKeyF2 To vbKeyF12
          KeyCode = 0
        
        Case Else
            Exit Sub
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :FVS200の編集変更処理
'        MODULE_ID        :
'        NOTE             :サブフォームから呼ばれる
'        CREATE_DATE      :2006/05/01
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetFVS200EditChange()

    If pIsChanged = False Then
        Me.Caption = Me.Caption & "(修正中)"
        pIsChanged = True
    End If

End Sub
' ☆☆☆ Public関数群 ☆☆☆
'****************************  ended or program ********************************
