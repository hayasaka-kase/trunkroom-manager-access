VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード解約業務画面
'   プログラムＩＤ　：　FVS210
'   作　成　日　　　：  2007/04/09
'   作　成　者　　　：  h.ueno
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2009/04/30
'   修　正　者　　　：  hirano
'   修　正　内　容　：  ヤード解約一覧遷移を追加
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：  2012/01/11
'   修　正　者　　　：  M.RYU
'   修　正　内　容　：  FVS220に遷移するときのパラメータに営業終了日を解約予定日に変更
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：  2013/06/03
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  終了メッセージを表示させないように修正。
'   Ver             ：  0.3
'**********************************************
'修正履歴
'   修　正　日　　　：  2018/02/17
'   修　正　者　　　：  k.sato
'   修　正　内　容　：  解約ヤード振分表の出力ボタンを追加
'   Ver             ：  0.4
'**********************************************
'修正履歴
'   修　正　日　　　：  2018/05/08
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  解約ヤード振分表を解約ヤードじゃなくても出力できるように修正。
'   Ver             ：  0.5
'**********************************************
'修正履歴
'   修　正　日　　　：  2018/06/17
'   修　正　者　　　：  k.sato
'   修　正　内　容　：  解約ヤード振分表の選択条件に「営業日を含める」を追加
'   Ver             ：  0.6
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   Private定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_WK_TABLE_NM           As String = "FVS210_W01"  ' ワークテーブル名

Private Const P_MSG_ERR               As String = "エラー"
Private Const P_MSG_KAKUNIN           As String = "確認"

Private Const P_MAX_YARD_NO           As String = "aaaaaa"
Private Const P_PROG_ID     As String = "FVS210"
'==============================================================================*
'   Private変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private pobjKONT_DB                   As Database ' コンテナDBオブジェクト
Private pstrBumonCd                   As String   ' 部門コード

Private pstrYardCode                  As String   ' パラメータで渡ってきたヤードコード

'
'        MODULE_NAME      :cmd_PrintC_Click イベント
'        MODULE_ID        :cmd_PrintC_Click
'        CREATE_DATE      :2018/02/17
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_PrintC_Click()

  On Error GoTo Exception
  
    If Nz(Me.cmb_ヤードコード.VALUE, "") = "" Then
        MsgBox "ヤードを選択して下さい。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
      
#If Not test Then
'DELETE 2018/05/08 M.HONDA Start
'    If Nz(Me.txt_解約予定日.VALUE, "") = "" Then
'        MsgBox "このヤードは解約予定ではありません。", vbInformation, P_MSG_KAKUNIN
'        Me.cmb_ヤードコード.SetFocus
'        Exit Sub
'    End If
'DELETE 2018/05/08 M.HONDA End
#End If
      
    If MsgBox("解約ヤード契約者一覧表をExcelに出力しますか？", _
                vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    
    ' 印刷の実行
    If Print210.fncPrintYardInlimit(Print210.P210_MODE_EXCEL, pstrBumonCd, Me.cmb_ヤードコード.VALUE, Me.txt_ヤード名.VALUE, "3") = False Then
           MsgBox "このヤードには契約者はいません。", vbInformation, P_MSG_KAKUNIN
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_PrintC_Click", True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim strParam()  As String
    Dim iParamCnt   As Integer
    
  On Error GoTo Exception
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_タイトル.BackColor = RGB(255, 0, 204)
    End Select

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
    
    'INSERT 2018-06-17
    Me.chk_営業日.VALUE = False
    
    Exit Sub
        
Exception:
    Call MSZZ024_M00("Form_Open", True)
    Cancel = True
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_End_Click イベント
'        MODULE_ID        :cmd_End_Click
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

  On Error GoTo Exception
  
    '' 2013/06/03 M.HONDA START
    ' 確認メッセージ表示
'    If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
'        Exit Sub
'    End If
    '' 2013/06/03 M.HONDA END

    ' 本画面を閉じる
    doCmd.Close acForm, Me.FormName
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_End_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_PrintA_Click イベント
'        MODULE_ID        :cmd_PrintA_Click
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_PrintA_Click()

  On Error GoTo Exception
  
    If Nz(Me.cmb_ヤードコード.VALUE, "") = "" Then
        MsgBox "ヤードを選択して下さい。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
      
    If Nz(Me.txt_解約予定日.VALUE, "") = "" Then
        MsgBox "このヤードは解約予定ではありません。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
      
    If MsgBox("解約ヤード契約者一覧表をExcelに出力しますか？", _
                vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    
    ' 印刷の実行
    If Print210.fncPrintYardInlimit(Print210.P210_MODE_EXCEL, pstrBumonCd, Me.cmb_ヤードコード.VALUE, Me.txt_ヤード名.VALUE, "1") = False Then
           MsgBox "このヤードには契約者はいません。", vbInformation, P_MSG_KAKUNIN
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_PrintA_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_PrintB_Click イベント
'        MODULE_ID        :cmd_PrintB_Click
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_PrintB_Click()

  On Error GoTo Exception
  
    If Nz(Me.cmb_ヤードコード.VALUE, "") = "" Then
        MsgBox "ヤードを選択して下さい。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
      
    If Nz(Me.txt_解約予定日.VALUE, "") = "" Then
        MsgBox "このヤードは解約予定ではありません。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
            
    If MsgBox("代替候補一覧表をExcelに出力しますか？", _
                vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    
    ' 印刷の実行
    If Print210.fncPrintYardInlimit(Print210.P210_MODE_EXCEL, pstrBumonCd, Me.cmb_ヤードコード.VALUE, Me.txt_ヤード名.VALUE, "2") = False Then
           MsgBox "このヤードの移動候補（紹介制限ヤード）が" & vbCr & "無いまたは契約者はいません。", vbInformation, P_MSG_KAKUNIN
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_PrintB_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_FVS200_Click イベント
'        MODULE_ID        :cmd_FVS200_Click
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_FVS200_Click()

  On Error GoTo Exception
    
    If Nz(Me.cmb_ヤードコード.VALUE, "") <> "" Then
    
      If Nz(Me.txt_解約予定日.VALUE, "") = "" Then
          MsgBox "このヤードは解約ではありません。", vbInformation, P_MSG_KAKUNIN
          Me.cmb_ヤードコード.SetFocus
          Exit Sub
      End If
    
      '紹介制限ヤード状況 オープン
      doCmd.OpenForm "FVS200", , , , , , Format$(Me.cmb_ヤードコード, "000000")
        
    Else
        
      doCmd.OpenForm "FVS200"
        
    End If
        
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_FVS200_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Print220_Click イベント
'        MODULE_ID        :cmd_Print220_Click
'        CREATE_DATE      :2007/05/07
'        UPDATE_DATE      :2009/04/30 by hirano cmd_FVS220_Click -> cmd_Print220_Click
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmd_FVS220_Click()
Private Sub cmd_Print220_Click()

  On Error GoTo Exception
    
    ' 印刷の実行
'    If Print220.fncPrintYardListInlimit(Print220.P220_MODE_PREVIEW, Me.cmb_ヤードコード.Value, "") = False Then
    If Print220.fncPrintYardListInlimit(Print220.P220_MODE_PREVIEW, "", "") = False Then
           MsgBox "対象なし！", vbInformation, P_MSG_KAKUNIN
    End If
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_FVS220_Click", True)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤードコード_AfterUpdate イベント
'        MODULE_ID        :cmb_ヤードコード_AfterUpdate
'        CREATE_DATE      :2007/04/09
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
    
    ' ヤードの情報を設定
    If fncSetYardMastData(Me.cmb_ヤードコード.VALUE) = False Then
        Call MSZZ024_M10("setAllData", "ヤード情報の取得が失敗しました。")
    End If

    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_ヤードコード_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤードコード_BeforeUpdate イベント
'        MODULE_ID        :cmb_ヤードコード_BeforeUpdate
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
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
'        MODULE_NAME      :KONT_DB接続
'        MODULE_ID        :connectKontDB
'        PARAMETER        :a部門CD
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
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
'        MODULE_NAME      :部門コード取得
'        MODULE_ID        :fncGetBumonName
'        PARAMETER        :
'        RETURN           :部門コード
'        CREATE_DATE      :2007/04/09
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
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isExistYardCode(ByVal aYardCode As Variant) As Boolean

    Dim varRtn     As Variant
    Dim strSQL     As String
    Dim rsGetData  As Recordset
    
  On Error GoTo Exception

    isExistYardCode = False

    ' ヤードコード未入力はエラー
    If Nz(aYardCode, "") = "" Then Exit Function

    ' 存在チェック
    strSQL = " SELECT COUNT(*) YARD_COUNT " & Chr(13)
    strSQL = strSQL & " FROM   YARD_MAST " & Chr(13)
    strSQL = strSQL & " WHERE  YARD_CODE = " & Val(aYardCode) & Chr(13)
    
    Set rsGetData = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    If rsGetData("YARD_COUNT") < 1 Then
        rsGetData.Close: Set rsGetData = Nothing
        Exit Function
    Else
        isExistYardCode = True 'そんざいスマス
    End If

    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    
    Exit Function

Exception:
    ' エラー発生時
    Call Err.Raise(Err.Number, "isExistYardCode" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタ情報を画面に設定
'        MODULE_ID        :fncSetYardMastData
'        PARAMETER        :strYardCode        - ヤードコード
'        RETURN           :True = 正常終了／False = 異常終了
'        CREATE_DATE      :2007/04/09
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetYardMastData(ByVal strYardCode As String) As Boolean

    Dim strSQL        As String
    Dim rsGetData     As Recordset
    Dim strYard_Name  As String
    Dim blnErrFlg     As Boolean

    On Error GoTo Exception

    strYard_Name = ""
    blnErrFlg = False

    ' ヤード名取得
    strSQL = "SELECT    YARD_NAME " & Chr(13)
    strSQL = strSQL & "         ,YARD_RENTEND_DAY " & Chr(13)
    strSQL = strSQL & "         ,YARD_END_DAY " & Chr(13)
    strSQL = strSQL & "         ,(SELECT COUNT(YARD_INLIMIT_YCODE) FROM YARD_MAST WHERE YARD_INLIMIT_YCODE = " & Val(strYardCode) & ") LIMIT_CNT " & Chr(13)
    strSQL = strSQL & " FROM     YARD_MAST " & Chr(13)
    strSQL = strSQL & " WHERE    YARD_CODE = " & Val(strYardCode)

    Set rsGetData = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    If Not rsGetData.EOF Then
        Me.txt_ヤード名 = rsGetData.Fields("YARD_NAME")
        Me.txt_営業終了日 = rsGetData.Fields("YARD_RENTEND_DAY")
        Me.txt_解約予定日 = rsGetData.Fields("YARD_END_DAY")
        Me.txt_紹介制限ヤード数 = rsGetData.Fields("LIMIT_CNT")
        fncSetYardMastData = True   '取得OK
    End If
    
    GoTo Finaly

Exception:
    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing
    blnErrFlg = True

Finaly:
    'DB切断
    If Not rsGetData Is Nothing Then rsGetData.Close: Set rsGetData = Nothing

    If blnErrFlg = True Then
        ' エラー発生時
        Call Err.Raise(Err.Number, "fncGetYardMastData" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :初期処理（画面オープン時に一度だけ実行する）
'        MODULE_ID        :initializFunction
'        PARAMETER        :
'        CREATE_DATE      :2007/04/09
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
    
    ' 部門コード取得
    pstrBumonCd = fncGetBumonName
    
    ' KONT_DBに接続
    Call connectKontDB(pstrBumonCd)
    
    ' 画面設定など....
    ' メニューから呼ばれたとき
    If pstrYardCode <> "" Then
    ' 他画面から呼ばれたとき...引数のヤードコードで画面表示
        Me.cmb_ヤードコード = Format(pstrYardCode, "000000")
    End If

    GoTo Finally
    
Exception:
    isException = True
    errorMessage = "system error"

Finally:
    initializFunction = errorMessage
    If isException = True Then
        Call Err.Raise(Err.Number, "initializFunction" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmd_FVS220_Click イベント
'        MODULE_ID        :cmd_FVS220_Click
'        CREATE_DATE      :2009/04/30
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_FVS220_Click()

    Dim str_Param As String
  
    If Nz(Me.cmb_ヤードコード.VALUE, "") = "" Then
        MsgBox "ヤードを選択して下さい。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
    
    If Nz(Me.txt_解約予定日.VALUE, "") = "" Then
        MsgBox "このヤードは解約予定ではありません。", vbInformation, P_MSG_KAKUNIN
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If
    
'    str_Param = P_PROG_ID & "," & Nz(Me.cmb_ヤードコード.VALUE, "") & "," & Me.txt_ヤード名 & "," & Nz(txt_営業終了日, "") 'DELETE 2012/01/11 M.RYU
    str_Param = P_PROG_ID & "," & Nz(Me.cmb_ヤードコード.VALUE, "") & "," & Me.txt_ヤード名 & "," & Nz(txt_解約予定日, "")  'INSERT 2012/01/11 M.RYU
    doCmd.OpenForm "FVS220", , , , , , str_Param

End Sub

' ☆☆☆ Public関数群 ☆☆☆
'==============================================================================*
'
'        MODULE_NAME      :ファンクションキー押下処理
'        MODULE_ID        :FunctionKeyDown
'        PARAMETER        :
'        NOTE             :サブフォームからも呼ばれる
'        CREATE_DATE      :2007/04/09
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
' ☆☆☆ Public関数群 ☆☆☆
'****************************  ended or program ********************************

