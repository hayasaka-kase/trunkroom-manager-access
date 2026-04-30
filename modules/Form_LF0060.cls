VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LF0060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : コンテナ管理システム
'        SUB_SYSTEM_NAME : 帳票
'
'        PROGRAM_NAME    : コンテナ在庫一覧表
'        PROGRAM_ID      : LF0060
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2019/11/20
'        CERATER         : Y.WADA
'        Ver             : 0.0
'
'        UPDATE          : 2020/02/15
'        UPDATER         : Y.WADA
'        Ver             : 0.1
'                        : モジュール（MS0060A）のみ修正
'
'        UPDATE          : 2020/06/09
'        UPDATER         : Y.WADA
'        Ver             : 0.2
'                        : モジュール（MS0060A）のみ修正
'
'        UPDATE          : 2020/06/12
'        UPDATER         : Y.WADA
'        Ver             : 0.3
'                        : 帳表種類を追加、コンテナ在庫一覧表（サマリ）を追加
'
'        UPDATE          : 2020/07/07
'        UPDATER         : Y.WADA
'        Ver             : 0.4
'                        : モジュール（MS0060A, MS0060B）のみ修正
'
'        UPDATE          : 2021/05/20
'        UPDATER         : K.KINEBUCHI
'        Ver             : 0.5
'                        : リスト種類の「コンテナ在庫一覧表（種類毎）」を非表示にする。
'
'        UPDATE          : 2022/09/30
'        UPDATER         : K.KINEBUCHI
'        Ver             : 0.6
'                        : モジュール「LS0060、LS0060A、LS0060B」を修正
'
'        UPDATE          : 2022/10/26
'        UPDATER         : N.IMAI
'        Ver             : 1.0
'                        : 在庫総合計の作成、出力を追加
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "LF0060"

'==============================================================================*
'
'        MODULE_NAME      :EXCEL
'        MODULE_ID        :cmd_excel_Click
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()
    On Error GoTo ErrorHandler
    
    Dim strYM As String
       
'DELETE 2020/06/12 Y.WADA Start
'    If fncInputCheck() Then
'
'        If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, Me.Caption) = vbYes) Then
'
'            doCmd.Hourglass True '砂時計    'INSERT 2020/06/09 Y.WADA
'
'            strYM = Replace(txt_zaikoYM.VALUE, "/", "")
'            If strYM = LS0060A_SETDATE_YM Then
'                '入力日＝締日を考慮した当月の場合は、在庫集計トランデータを作成
'                If Not LS0060A_M00(strYM) Then
'                    doCmd.Hourglass False    '砂時計解除    'INSERT 2020/06/09 Y.WADA
'                    MsgBox "在庫集計トランデータの作成ができませんでした。", vbOKOnly + vbInformation, Me.Caption
'                    Exit Sub
'                End If
'            End If
'
'            If Not LS0060_xlPreview(strYM) Then
'                doCmd.Hourglass False    '砂時計解除    'INSERT 2020/06/09 Y.WADA
'                MsgBox "対象データがありませんでした。", vbOKOnly + vbInformation, Me.Caption
'                Exit Sub
'            End If
'
'        End If
'
'        doCmd.Hourglass False    '砂時計解除    'INSERT 2020/06/09 Y.WADA
'    End If
'DELETE 2020/06/12 Y.WADA End

    'INSERT 2020/06/12 Y.WADA Start
    If Not fncInputCheck() Then
        Exit Sub
    End If
        
    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, Me.Caption) = vbNo) Then
        Exit Sub
    End If
            
    doCmd.Hourglass True '砂時計
    
    strYM = Replace(txt_zaikoYM.VALUE, "/", "")
    If strYM = LS0060A_SETDATE_YM Then
        If (MsgBox("在庫集計トランデータを作成しますか？", vbYesNo + vbQuestion, Me.Caption) = vbYes) Then
            '入力日＝締日を考慮した当月の場合は、在庫集計トランデータを作成
            If Not LS0060A_M00(strYM) Then
                doCmd.Hourglass False    '砂時計解除
                MsgBox "在庫集計トランデータの作成ができませんでした。", vbOKOnly + vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
            
    'Excel出力
    Select Case opg_LIST_SYURUI
    Case opb_LIST_SYURUI.OptionValue
        If Not LS0060_xlPreview(strYM) Then
            doCmd.Hourglass False    '砂時計解除
            MsgBox "対象データがありませんでした。", vbOKOnly + vbInformation, Me.Caption
            Exit Sub
        End If
    
    Case opb_LIST_SUMM.OptionValue
        If Not LS0060B_xlPreview(strYM) Then
            doCmd.Hourglass False    '砂時計解除
            MsgBox "対象データがありませんでした。", vbOKOnly + vbInformation, Me.Caption
            Exit Sub
        End If
    End Select
        
    doCmd.Hourglass False    '砂時計解除
    'INSERT 2020/06/12 Y.WADA End

Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmd_excel_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面最大表示
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    On Error Resume Next
    doCmd.Maximize
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームのクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    On Error GoTo ErrorHandler
    Call MSHS010_M00(PROG_ID, "2")
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("Form_Close", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームのオープン
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    On Error GoTo ErrorHandler
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル44.BackColor = RGB(255, 0, 204)
    End Select
    Call MSHS010_M00(PROG_ID, "1")
    
    '対象年月の初期値は前月
    txt_zaikoYM = Format(DateSerial(Year(Now()), Month(Now()) - 1, 1), "yyyy/mm")

Exit Sub

ErrorHandler:
    Call MSZZ024_M00("Form_Open", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :閉じる
'        MODULE_ID        :MFAM009_CLOSB_Click
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_close_Click()
    On Error Resume Next
    doCmd.Close
End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :FunInputCheck
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean
    On Error GoTo ErrorHandler
    
    fncInputCheck = True
    
    If Not fncInputCheckHiss() Then
        fncInputCheck = False
        Exit Function
    End If

'    If Not fncInputCheckType() Then
'        fncInputCheck = False
'        Exit Function
'    End If
Exit Function
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncInputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :必須入力チェック
'        MODULE_ID        :InputCheckHiss
'        CREATE_DATE      :2019/11/20 Y.WADA
'        RETURN           :正常(True)／異常(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheckHiss() As Boolean
    On Error GoTo ErrorHandler

    fncInputCheckHiss = False

    If Not checkInput(txt_zaikoYM) Then
        Exit Function
    End If
    
    fncInputCheckHiss = True
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "fncInputCheckHiss" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
''
''        MODULE_NAME     :型入力チェック
''        MODULE_ID       :fncInputCheckType
''        CREATE_DATE     :2019/11/20 Y.WADA
''        RETURN          :正常(True)／異常(False)
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncInputCheckType() As Boolean
'    On Error GoTo ErrorHandler
'
'    fncInputCheckType = False
'
'    If Not checkIsDate(txt_ZaikoYM, 7) Then
'        Exit Function
'    End If
'
'    fncInputCheckType = True
'Exit Function
'ErrorHandler:                   '↓自分の関数名
'    Call Err.Raise(Err.Number, "fncInputCheckType" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Function

'==============================================================================*
'
'        MODULE_NAME      :fncFirstDay
'        機能             :月初日
'        IN               :対象年月のコントロール
'        OUT              :月初日
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncFirstDay(objTarget As Object) As String

    If Nz(objTarget) = "" Then
        fncFirstDay = ""
    Else
        fncFirstDay = objTarget & "/01"
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDateCheck
'        機能             :日付入力チェック
'        IN               :①対象テキスト
'                          ②更新後処理のキャンセル
'                          ③対象テキストのコントロール
'        OUT              :True=正常 False=入力エラー
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateCheck(txtTarget As Variant, ByRef Cancel As Integer, ctlTarget As Control) As Boolean

    fncDateCheck = True
    
    If IsNull(CmDateChecker(txtTarget)) Then
        Call checkMsgBox(ctlTarget, "日付を正しく入力してください。")
        
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :txt_zaikoYM_AfterUpdate
'        MODULE_ID        :txt_zaikoYM_AfterUpdate
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_zaikoYM_AfterUpdate()

    txt_zaikoYM = Format(CmDateChecker(fncFirstDay(txt_zaikoYM)), "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_zaikoYM_BeforeUpdate
'        MODULE_ID        :txt_zaikoYM_BeforeUpdate
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_zaikoYM_BeforeUpdate(Cancel As Integer)
        
    Call fncDateCheck(fncFirstDay(txt_zaikoYM), Cancel, txt_zaikoYM)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_zaikoYM_GotFocus
'        MODULE_ID        :txt_zaikoYM_GotFocus
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_zaikoYM_GotFocus()

    txt_zaikoYM = Format(fncFirstDay(txt_zaikoYM), "yy/mm")
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_zaikoYM_KeyPress
'        MODULE_ID        :txt_zaikoYM_KeyPress
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_zaikoYM_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txt_zaikoYM, KeyAscii, 7)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_zaikoYM_LostFocus
'        MODULE_ID        :txt_zaikoYM_LostFocus
'        CREATE_DATE      :2019/11/20 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_zaikoYM_LostFocus()

    txt_zaikoYM = Format(fncFirstDay(txt_zaikoYM), "yyyy/mm")

End Sub
'****************************  ended or program ********************************
