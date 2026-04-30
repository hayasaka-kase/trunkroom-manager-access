VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約受付リスト印刷
'   プログラムＩＤ　：　FVS510
'   作　成　日　　　：  2005/05/27
'   作　成　者　　　：  ChronoSystem 三ツ野
'   Ver             ：  0.0
'**********************************************
'   プログラム名　　：　希望サイズ対応
'   プログラムＩＤ　：　FVS510
'   作　成　日　　　：  2005/06/26
'   作　成　者　　　：  K.MITSUNO
'   Ver             ：  0.1
'**********************************************
'   プログラム名　　：　検索条件追加(ﾔｰﾄﾞ)対応
'   プログラムＩＤ　：　FVS510
'   作　成　日　　　：  2005/07/23
'   作　成　者　　　：  K.MITSUNO
'   Ver             ：  0.2
'**********************************************
'   プログラム名　　：　予約受付リスト2の出力機能追加
'   プログラムＩＤ　：　FVS510
'   作　成　日　　　：  2005/09/15
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.3
'**********************************************
'   プログラム名　　：　画面に｢自動区分｣追加
'   プログラムＩＤ　：　FVS510
'   作　成　日　　　：  2005/11/04
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.4
'**********************************************
'   プログラム名　　：　予約受付リスト印刷
'   プログラムＩＤ　：　FVS510
'   修　正　日　　　：  2006/01/12
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  1.画面に移動区分追加
'                       2.帳票出力に移動区分追加
'   Ver             ：  0.5
'**********************************************
'**********************************************
'   プログラム名　　：　予約受付リスト印刷
'   プログラムＩＤ　：　FVS510
'   修　正　日　　　：  2007/04/03
'   修　正　者　　　：  H.UENO
'   修　正　内　容　：  画面に予約担当者追加
'                   ：  近隣紹介可否を帳票に追加
'   Ver             ：  0.6
'**********************************************
'
'   修　正　日　　　：  2007/08/23
'   修　正　者　　　：  tajima
'   修　正　内　容　：  ネット予約対応
'                   ：  検索条件の追加
'   Ver             ：  0.7
'**********************************************
'
'   修　正　日　　　：  2012/07/04
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  ログ出力対応
'                   ：
'   Ver             ：  0.8
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Const PROG_ID = "FVS510"

Private Const P_SPACE       As String = " "
Private Const P_TITLE       As String = "予約受付リスト"

'状態区分ｺｰﾄﾞ
Private Const P_YUKBN_00    As String = "00"    '予約未確定
Private Const P_YUKBN_01    As String = "01"    '予約確定
Private Const P_YUKBN_10    As String = "10"    '受付     ' 2006/01/12 ADD T.SUZUKI

'名称ﾏｽﾀｺｰﾄﾞ
Private Const P_NAME_020    As String = "020"   '顧客区分
Private Const P_NAME_040    As String = "040"   'コンテナ種別
Private Const P_NAME_070    As String = "070"   '広告媒体
Private Const P_NAME_080    As String = "080"   '受付状態区分
Private Const P_NAME_085    As String = "085"   '連絡区分
Private Const P_NAME_086    As String = "086"   '用途
Private Const P_NAME_087    As String = "087"   'レンタルサイズ範囲
Private Const P_NAME_090    As String = "090"   '段区分
Private Const P_NAME_098    As String = "098"   '移動区分                      ' 2006/01/12 ADD T.SUZUKI

'連絡区分
Private Const P_REN_00      As String = "00"    '電話
Private Const P_REN_01      As String = "01"    '携帯番号
Private Const P_REN_02      As String = "02"    'FAX
Private Const P_REN_10      As String = "10"    'Eメール

'段区分データ無し時表示文言
Private Const P_STEP_NASHI  As String = "指定無し"

'状態ステータス
Private Const P_JOUTAI_OK   As String = "貸出OK"
Private Const P_JOUTAI_MATI As String = "解約待ち"

Private Const P_FOMAT_KANMA As String = "#,###"
Private Const P_PROG_ID     As String = "FVS510"
Private Const P_RPT_ID      As String = "RVS510"
Private Const P_RPT_ID2     As String = "RVS512"                              ' 2005/09/15 ADD T.SUZUKI
Private Const P_KAKUNIN     As String = "確認"
Private Const P_ERROR       As String = "エラー"

' 2005/11/04 ADD T.SUZUKI Start
Private Const P_ALL_VALUE      As Integer = 0
Private Const P_AUTO_VALUE     As Integer = 1
Private Const P_HAND_VALUE     As Integer = 2


Private Const P_ALL            As String = "全て"
Private Const P_AUTO           As String = "自動のみ"
Private Const P_HAND           As String = "自動を除く"

Private Const P_TABLENM_FVS510 As String = "FVS510_W01"
Private Const P_TABLENM_FVS512 As String = "FVS512_W01"
' 2005/11/04 ADD T.SUZUKI End

' 2007/04/03 ADD H.UENO Start
Private Const P_PERSON_LENGTH       As Integer = 3              '担当コード
Private Const P_ZERO_SUPPRESS       As String = "0000000000"    '不足桁０埋め用
Private VarChangeDate As Variant                                '変更前処理で入力した日付を保存する
' 2007/04/03 ADD H.UENO End

Private Const P_OUTPUT_CANCEL  As Integer = 2501                              ' 2005/12/14 ADD T.SUZUKI
' 2007/04/03 ADD H.UENO Start
'==============================================================================*
'
'        MODULE_NAME      :cmb_Tanto_AfterUpdate
'        機能             :受付担当者 AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tanto_AfterUpdate()

    txt_Tanto = AfterUpdateCommon(cmb_Tanto, P_PERSON_LENGTH, 1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tanto_BeforeUpdate
'        機能             :受付担当者 BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tanto_BeforeUpdate(Cancel As Integer)

    '入力値をチェックする
    Call ChkComboInput(cmb_Tanto, P_PERSON_LENGTH, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_Tanto_KeyPress
'        機能             :受付担当者 KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tanto_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(Me.cmb_Tanto, KeyAscii, P_PERSON_LENGTH)

End Sub
' 2007/04/03 ADD H.UENO End

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form Activate時処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :2004/06/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    Call MSHS010_M00(PROG_ID, "2")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form Load時処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    '初期処理
    Call Frm_Init
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_End_Click
'        機能             :終了ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()
    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Excel_Click
'        機能             :Excelボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()

    Dim lngDataCount As Long  ' 対象データ数

    On Error GoTo cmd_Excel_Click_Err

' 2005/12/14 ADD T.SUZUKI Start
    ' 画面の入力チェック
    If fncInputCheck = False Then GoTo cmd_Excel_Click_End
' 2005/12/14 ADD T.SUZUKI End

    If (MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, P_KAKUNIN) = vbYes) Then

' 2005/09/15 ADD T.SUZUKI Start
        doCmd.Hourglass True

        Select Case Me.opg_Select
            Case 1  ' *** 予約受付リスト1 ***
                If Create_Wk = True Then
                   doCmd.OutputTo acQuery, "FVS510_Q01", "MicrosoftExcel(*.xls)", "", True, ""
                End If
    
            Case 2  ' *** 予約受付リスト2 ***
                If Create_Wk2 = True Then
                   doCmd.OutputTo acQuery, "FVS512_Q01", "MicrosoftExcel(*.xls)", "", True, ""
                End If
        End Select
' 2005/09/15 ADD T.SUZUKI End
' 2005/09/15 DEL T.SUZUKI Start
'        If Create_Wk = True Then
'           doCmd.OutputTo acQuery, "FVS510_Q01", "MicrosoftExcel(*.xls)", "", True, ""
'        End If
' 2005/09/15 DEL T.SUZUKI End
    End If

    GoTo cmd_Excel_Click_End

cmd_Excel_Click_Err:
    Select Case Err.Number
        Case P_OUTPUT_CANCEL
            ' Excel出力のキャンセル時のエラー
        Case Else
            MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, P_ERROR
    End Select

cmd_Excel_Click_End:
    doCmd.Hourglass False  ' 2005/09/15 ADD T.SUZUKI
    Me.cmb_YuKbnF.SetFocus
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Preview_Click
'        機能             :プレビューボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Preview_Click()
    
    On Error GoTo cmd_Preview_Click_Err

    Dim strMsg       As String
    Dim lngRc        As Long

' 2005/12/14 ADD T.SUZUKI Start
    ' 画面の入力チェック
    If fncInputCheck = False Then GoTo cmd_Preview_Click_End
' 2005/12/14 ADD T.SUZUKI End

    strMsg = "印刷内容を画面に出力します。よろしいですか？"
    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo, P_KAKUNIN)
    If lngRc = vbNo Then GoTo cmd_Preview_Click_End
    
   ' 2005/09/15 ADD T.SUZUKI Start
    doCmd.Hourglass True

    Select Case Me.opg_Select
        Case 1  ' *** 予約受付リスト1 ***
            If Create_Wk = False Then GoTo cmd_Preview_Click_End

            doCmd.OpenReport "RVS510_予約受付リスト", acPreview, "", ""

        Case 2  ' *** 予約受付リスト2 ***
            If Create_Wk2 = False Then GoTo cmd_Preview_Click_End

            doCmd.OpenReport "RVS512", acPreview, "", ""
    End Select
   ' 2005/09/15 ADD T.SUZUKI End
'    If Create_Wk = False Then GoTo cmd_Preview_Click_End  ' 2005/09/15 DEL T.SUZUKI

    GoTo cmd_Preview_Click_End
    
cmd_Preview_Click_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, P_ERROR

cmd_Preview_Click_End:
    doCmd.Hourglass False  ' 2005/09/15 ADD T.SUZUKI
    Me.cmb_YuKbnF.SetFocus
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YardF_BeforeUpdate
'        機能             :ヤード番号FROM BeforeUpdate時
'        IN               :2005/07/23 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardF_BeforeUpdate(Cancel As Integer)
    
    Dim intChk As Integer
    
    'ｺｰﾄﾞが空の場合、顧客区分名欄も空欄表示
    If ControlCrear(Me.cmb_YardF, Me.txt_YardF) = False Then Exit Sub
    
    If Yard_BeforeUp(Me.cmb_YardF, intChk) = False Then
        Cancel = True
        If intChk = 2 Then SendKeys "{ESC}"
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YardF_AfterUpdate
'        機能             :ヤード番号FROM AfterUpdate時
'        IN               :2005/07/23 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardF_AfterUpdate()
    Call Yard_AfterUp(cmb_YardF, txt_YardF)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YardF_KeyPress
'        機能             :ヤード番号FROM KeyPress時
'        IN               :2005/07/23 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YardF, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YardT_BeforeUpdate
'        機能             :ヤード番号TO BeforeUpdate時
'        IN               :2005/07/23 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardT_BeforeUpdate(Cancel As Integer)
    
    Dim intChk As Integer
    
    'ｺｰﾄﾞが空の場合、顧客区分名欄も空欄表示
    If ControlCrear(Me.cmb_YardT, Me.txt_YardT) = False Then Exit Sub

    If Yard_BeforeUp(Me.cmb_YardT, intChk) = False Then
        Cancel = True
        If intChk = 2 Then SendKeys "{ESC}"
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YardT_AfterUpdate
'        機能             :ヤード番号TO AfterUpdate時
'        IN               :2005/07/23 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardT_AfterUpdate()
    Call Yard_AfterUp(cmb_YardT, txt_YardT)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YardT_KeyPress
'        機能             :ヤード番号TO KeyPress時
'        IN               :2005/07/23 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YardT, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YuKbnF_KeyPress
'        機能             :状態受付区分From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YuKbnF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YuKbnF, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YuKbnF_BeforeUpdate
'        機能             :状態受付区分From BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YuKbnF_BeforeUpdate(Cancel As Integer)
    
    'ｺｰﾄﾞが空の場合、受付区分名欄も空欄表示
    If ControlCrear(Me.cmb_YuKbnF, Me.txt_YuKbnF) = False Then Exit Sub
    
    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_YuKbnF, P_NAME_080) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YuKbnF_AfterUpdate
'        機能             :状態受付区分From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YuKbnF_AfterUpdate()

    On Error GoTo cmb_YuKbnF_AfterUpdate_Err
    
    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_YuKbnF, Me.txt_YuKbnF, P_NAME_080) = False Then GoTo cmb_YuKbnF_AfterUpdate_End

' 2006/01/12 ADD T.SUZUKI Start
    If Format(Me.cmb_YuKbnF, "00") = P_YUKBN_10 Then
        ' 受付が選択された場合
        If Format(Me.cmb_YuKbnT, "00") <> P_YUKBN_10 Then
            Call subSetMoveKbn(True)
        End If
    Else
        ' 上記以外
        If Format(Me.cmb_YuKbnT, "00") <> P_YUKBN_10 Then
            Call subSetMoveKbn(False)
        End If
    End If
' 2006/01/12 ADD T.SUZUKI End

    GoTo cmb_YuKbnF_AfterUpdate_End
    
cmb_YuKbnF_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
cmb_YuKbnF_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YuKbnT_KeyPress
'        機能             :状態受付区分To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YuKbnT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_YuKbnT, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YuKbnT_BeforeUpdate
'        機能             :状態受付区分To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YuKbnT_BeforeUpdate(Cancel As Integer)

    'ｺｰﾄﾞが空の場合、受付区分名欄も空欄表示
    If ControlCrear(Me.cmb_YuKbnT, Me.txt_YuKbnT) = False Then Exit Sub
    
    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_YuKbnT, P_NAME_080) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_YuKbnT_AfterUpdate
'        機能             :状態受付区分To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YuKbnT_AfterUpdate()

    On Error GoTo cmb_YuKbnT_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_YuKbnT, Me.txt_YuKbnT, P_NAME_080) = False Then GoTo cmb_YuKbnT_AfterUpdate_End

' 2006/01/12 ADD T.SUZUKI Start
    If Format(Me.cmb_YuKbnT, "00") = P_YUKBN_10 Then
        ' 受付が選択された場合
        If Format(Me.cmb_YuKbnF, "00") <> P_YUKBN_10 Then
            Call subSetMoveKbn(True)
        End If
    Else
        ' 上記以外
        If Format(Me.cmb_YuKbnF, "00") <> P_YUKBN_10 Then
            Call subSetMoveKbn(False)
        End If
    End If
' 2006/01/12 ADD T.SUZUKI End

    GoTo cmb_YuKbnT_AfterUpdate_End

cmb_YuKbnT_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
cmb_YuKbnT_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_F_AfterUpdate
'        機能             :cmb_MoveKbn_F_AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_F_AfterUpdate()

    On Error GoTo cmb_MoveKbn_F_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_MoveKbn_F, Me.txt_MoveKbn_F, P_NAME_098) = False Then GoTo cmb_MoveKbn_F_AfterUpdate_End

    GoTo cmb_MoveKbn_F_AfterUpdate_End

cmb_MoveKbn_F_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_MoveKbn_F_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_F_BeforeUpdate
'        機能             :cmb_MoveKbn_F_BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_F_BeforeUpdate(Cancel As Integer)

    'ｺｰﾄﾞが空の場合、移動区分名欄も空欄表示
    If ControlCrear(Me.cmb_MoveKbn_F, Me.txt_MoveKbn_F) = False Then Exit Sub

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_MoveKbn_F, P_NAME_098) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_F_KeyPress
'        機能             :cmb_MoveKbn_F_KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_F_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_MoveKbn_F, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_T_AfterUpdate
'        機能             :cmb_MoveKbn_T_AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_T_AfterUpdate()

    On Error GoTo cmb_MoveKbn_T_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_MoveKbn_T, Me.txt_MoveKbn_T, P_NAME_098) = False Then GoTo cmb_MoveKbn_T_AfterUpdate_End

    GoTo cmb_MoveKbn_T_AfterUpdate_End

cmb_MoveKbn_T_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_MoveKbn_T_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_T_BeforeUpdate
'        機能             :cmb_MoveKbn_T_BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_T_BeforeUpdate(Cancel As Integer)

    'ｺｰﾄﾞが空の場合、移動区分名欄も空欄表示
    If ControlCrear(Me.cmb_MoveKbn_T, Me.txt_MoveKbn_T) = False Then Exit Sub

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_MoveKbn_T, P_NAME_098) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_T_KeyPress
'        機能             :cmb_MoveKbn_T_KeyPress
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_T_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_MoveKbn_T, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームオープン
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Title.BackColor = RGB(255, 0, 204)
    End Select
    Call MSHS010_M00(PROG_ID, "1")
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateF_KeyPress
'        機能             :受付日From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_UkDateF, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateF_GotFocus
'        機能             :受付日From GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateF_GotFocus()
    Call Date_TextFormat(Me.txt_UkDateF, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateF_LostFocus
'        機能             :受付日From LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateF_LostFocus()
    If Nz(Me.txt_UkDateF, "") <> "" Then
        Call Date_TextFormat(Me.txt_UkDateF, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateF_BeforeUpdate
'        機能             :受付日From BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateF_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_UkDateF, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateF_AfterUpdate
'        機能             :受付日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateF_AfterUpdate()
    Me.txt_UkDateF = Format(Me.txt_UkDateF, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateT_KeyPress
'        機能             :受付日To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_UkDateT, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateT_GotFocus
'        機能             :受付日To GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateT_GotFocus()
    Call Date_TextFormat(Me.txt_UkDateT, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateT_LostFocus
'        機能             :受付日To LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateT_LostFocus()
    If Nz(Me.txt_UkDateT, "") <> "" Then
        Call Date_TextFormat(Me.txt_UkDateT, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateT_BeforeUpdate
'        機能             :受付日To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateT_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_UkDateT, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_UkDateT_AfterUpdate
'        機能             :受付日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_UkDateT_AfterUpdate()
    Me.txt_UkDateT = Format(Me.txt_UkDateT, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateF_KeyPress
'        機能             :契約受付日From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_KuDateF, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateF_GotFocus
'        機能             :契約受付日From GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateF_GotFocus()
    Call Date_TextFormat(Me.txt_KuDateF, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateF_LostFocus
'        機能             :契約受付日From LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateF_LostFocus()
    If Nz(Me.txt_KuDateF, "") <> "" Then
        Call Date_TextFormat(Me.txt_KuDateF, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateF_AfterUpdate
'        機能             :契約受付日From LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateF_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_KuDateF, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateF_AfterUpdate
'        機能             :契約受付日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateF_AfterUpdate()
    Me.txt_KuDateF = Format(Me.txt_KuDateF, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateT_KeyPress
'        機能             :契約受付日To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_KuDateT, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateT_GotFocus
'        機能             :契約受付日To GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateT_GotFocus()
    Call Date_TextFormat(Me.txt_KuDateT, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateT_LostFocus
'        機能             :契約受付日To LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateT_LostFocus()
    If Nz(Me.txt_KuDateT, "") <> "" Then
        Call Date_TextFormat(Me.txt_KuDateT, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateT_BeforeUpdate
'        機能             :契約受付日To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateT_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_KuDateT, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KuDateT_AfterUpdate
'        機能             :契約受付日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KuDateT_AfterUpdate()
    Me.txt_KuDateT = Format(Me.txt_KuDateT, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateF_KeyPress
'        機能             :予約取消日From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_YtDateF, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateF_GotFocus
'        機能             :予約取消日From GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateF_GotFocus()
    Call Date_TextFormat(Me.txt_YtDateF, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateF_LostFocus
'        機能             :予約取消日From LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateF_LostFocus()
    If Nz(Me.txt_YtDateF, "") <> "" Then
        Call Date_TextFormat(Me.txt_YtDateF, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateF_BeforeUpdate
'        機能             :予約取消日From BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateF_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_YtDateF, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateF_AfterUpdate
'        機能             :予約取消日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateF_AfterUpdate()
    Me.txt_YtDateF = Format(Me.txt_YtDateF, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateT_KeyPress
'        機能             :予約取消日To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_YtDateT, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateT_GotFocus
'        機能             :予約取消日To GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateT_GotFocus()
    Call Date_TextFormat(Me.txt_YtDateT, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateT_LostFocus
'        機能             :予約取消日To LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateT_LostFocus()
    If Nz(Me.txt_YtDateT, "") <> "" Then
        Call Date_TextFormat(Me.txt_YtDateT, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateT_BeforeUpdate
'        機能             :予約取消日To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateT_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_YtDateT, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_YtDateT_AfterUpdate
'        機能             :予約取消日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YtDateT_AfterUpdate()
    Me.txt_YtDateT = Format(Me.txt_YtDateT, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateF_KeyPress
'        機能             :希望開始日From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_KkDateF, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateF_GotFocus
'        機能             :希望開始日From GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateF_GotFocus()
    Call Date_TextFormat(Me.txt_KkDateF, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateF_LostFocus
'        機能             :希望開始日From LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateF_LostFocus()
    If Nz(Me.txt_KkDateF, "") <> "" Then
        Call Date_TextFormat(Me.txt_KkDateF, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateF_BeforeUpdate
'        機能             :希望開始日From BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateF_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_KkDateF, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateF_AfterUpdate
'        機能             :希望開始日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateF_AfterUpdate()
    Me.txt_KkDateF = Format(Me.txt_KkDateF, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateT_KeyPress
'        機能             :希望開始日To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_KkDateT, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateT_GotFocus
'        機能             :希望開始日To GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateT_GotFocus()
    Call Date_TextFormat(Me.txt_KkDateT, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateT_LostFocus
'        機能             :希望開始日To LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateT_LostFocus()
    If Nz(Me.txt_KkDateT, "") <> "" Then
        Call Date_TextFormat(Me.txt_KkDateT, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateT_BeforeUpdate
'        機能             :希望開始日To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateT_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_KkDateT, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_KkDateT_AfterUpdate
'        機能             :希望開始日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_KkDateT_AfterUpdate()
    Me.txt_KkDateT = Format(Me.txt_KkDateT, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateF_KeyPress
'        機能             :取置期限日From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_TkDateF, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateF_GotFocus
'        機能             :取置期限日From GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateF_GotFocus()
    Call Date_TextFormat(Me.txt_TkDateF, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateF_LostFocus
'        機能             :取置期限日From LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateF_LostFocus()
    If Nz(Me.txt_TkDateF, "") <> "" Then
        Call Date_TextFormat(Me.txt_TkDateF, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateF_BeforeUpdate
'        機能             :取置期限日From BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateF_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_TkDateF, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateF_AfterUpdate
'        機能             :取置期限日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateF_AfterUpdate()
    Me.txt_TkDateF = Format(Me.txt_TkDateF, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateT_KeyPress
'        機能             :取置期限日To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_TkDateT, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateT_GotFocus
'        機能             :取置期限日To GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateT_GotFocus()
    Call Date_TextFormat(Me.txt_TkDateT, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateT_LostFocus
'        機能             :取置期限日To LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateT_LostFocus()
    If Nz(Me.txt_TkDateT, "") <> "" Then
        Call Date_TextFormat(Me.txt_TkDateT, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateT_BeforeUpdate
'        機能             :取置期限日To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateT_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_TkDateT, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TkDateT_AfterUpdate
'        機能             :取置期限日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TkDateT_AfterUpdate()
    Me.txt_TkDateT = Format(Me.txt_TkDateT, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateF_KeyPress
'        機能             :問合日From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateF_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_TaDateF, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateF_GotFocus
'        機能             :問合日From GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateF_GotFocus()
    Call Date_TextFormat(Me.txt_TaDateF, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateF_LostFocus
'        機能             :問合日From LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateF_LostFocus()
    If Nz(Me.txt_TaDateF, "") <> "" Then
        Call Date_TextFormat(Me.txt_TaDateF, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateF_BeforeUpdate
'        機能             :問合日From BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateF_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_TaDateF, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateF_AfterUpdate
'        機能             :問合日From AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateF_AfterUpdate()
    Me.txt_TaDateF = Format(Me.txt_TaDateF, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateT_KeyPress
'        機能             :問合日To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateT_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_TaDateT, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateT_GotFocus
'        機能             :問合日To GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateT_GotFocus()
    Call Date_TextFormat(Me.txt_TaDateT, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateT_LostFocus
'        機能             :問合日To LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateT_LostFocus()
    If Nz(Me.txt_TaDateT, "") <> "" Then
        Call Date_TextFormat(Me.txt_TaDateT, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateT_BeforeUpdate
'        機能             :問合日To BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateT_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_TaDateT, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_TaDateT_AfterUpdate
'        機能             :問合日To AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_TaDateT_AfterUpdate()
    Me.txt_TaDateT = Format(Me.txt_TaDateT, "yyyy/mm/dd")  ' 2007/08/23 fix tajima
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_AUTO_KeyPress
'        機能             :自動区分 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        CREATE           :2005/11/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_AUTO_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

'■□■□■□■□■□■□■□■□■□ ↓↓↓　Private 関数群 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□
'==============================================================================*
'
'        MODULE_NAME      :Date_TextFormat
'        機能             :日付ﾃﾞｰﾀのﾌｫｰﾏｯﾄ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Date_TextFormat(ByRef ctlTxt As TextBox, ByVal lngMode As Long)
    
    Dim vCode As Variant
    
    vCode = Nz(ctlTxt, "")
    If vCode <> "" Then
        If lngMode = 0 Then
            ctlTxt = Format(vCode, "yy/mm/dd")
        Else
            ctlTxt = Format(vCode, "yyyy/mm/dd")
        End If
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :ControlCrear
'        機能             :ｺﾝﾎﾞﾎﾞｯｸｽの内容が空の場合、名称ﾃｷｽﾄも空にする
'        IN               :ctlCmb = ｺﾝﾎﾞﾎﾞｯｸｽ / ctlTxt = ﾃｷｽﾄ
'        OUT              :ｺﾝﾎﾞﾎﾞｯｸｽが空だったかの判断
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ControlCrear(ByRef ctlCmb As ComboBox, ByRef ctlTxt As TextBox) As Boolean
    
    ControlCrear = False
    If Nz(ctlCmb, "") = "" Then
        ctlTxt = ""
        Exit Function
    Else
        ControlCrear = True
    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :Chk_NameAfter
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :ctl = 対象ｺﾝﾎﾞﾎﾞｯｸｽ / ctlText = 名称ｾｯﾄ先ﾃｷｽﾄ / strID = 名称ｺｰﾄﾞ
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_NameAfter(ByRef ctlCmb As ComboBox, ByRef ctlText As TextBox, ByVal strID As String) As Boolean
    
    Dim vCode       As Variant
    Dim strWhere    As String
    
    Chk_NameAfter = False
    
    vCode = Nz(ctlCmb, "")
    If vCode = "" Then GoTo Chk_NameAfter_Err
    
    'ファイルへの存在チェック
    strWhere = "NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE =" & Format$(vCode, "000")
    
    ctlText = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
    ctlCmb = Format$(ctlCmb, "00")
    
    GoTo Chk_NameAfter_End
    
Chk_NameAfter_Err:
    Exit Function
    
Chk_NameAfter_End:
    Chk_NameAfter = True
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :ﾌｫｰﾑ初期化処理
'        MODULE_ID        :Frm_Init
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Frm_Init()

    Dim objDb   As Database
    Dim strName As String    ' 2005/12/14 ADD T.SUZUKI

    Set objDb = CurrentDb

' 2005/12/14 ADD T.SUZUKI Start
    cmb_YuKbnF.VALUE = P_YUKBN_00
    cmb_YuKbnT.VALUE = P_YUKBN_00

    strName = ""
    ' 予約受付状態区分名取得
    strName = GetNameData(objDb, P_NAME_080, P_YUKBN_00)
    txt_YuKbnF = strName
    txt_YuKbnT = strName
' 2005/12/14 ADD T.SUZUKI End
' 2005/12/14 DEL T.SUZUKI Start
'    cmb_YuKbnF.Value = P_YUKBN_00
'    txt_YuKbnF = GetNameData(objDb, P_NAME_080, P_YUKBN_00)
'    cmb_YuKbnT.Value = P_YUKBN_01                            ' 2005/12/14 DEL T.SUZUKI
'    txt_YuKbnT = GetNameData(objDb, P_NAME_080, P_YUKBN_01)  ' 2005/12/14 DEL T.SUZUKI
' 2005/12/14 DEL T.SUZUKI End

    Set objDb = Nothing

    txt_UkDateF = ""
    txt_UkDateT = ""
    txt_KuDateF = ""
    txt_KuDateT = ""
    txt_YtDateF = ""
    txt_YtDateT = ""
    txt_KkDateF = ""
    txt_KkDateT = ""
    txt_TkDateF = ""
    txt_TkDateT = ""
    txt_TaDateF = ""
    txt_TaDateT = ""

' 2005/11/04 ADD T.SUZUKI Start
    ' 自動または手動
    cmb_AUTO = ""

    cmb_AUTO.AddItem Item:=P_ALL_VALUE & ";" & P_ALL
    cmb_AUTO.AddItem Item:=P_AUTO_VALUE & ";" & P_AUTO
    cmb_AUTO.AddItem Item:=P_HAND_VALUE & ";" & P_HAND

    ' 初期値として｢全て｣を選択
    cmb_AUTO = P_ALL_VALUE
' 2005/11/04 ADD T.SUZUKI End

    ' 移動区分の可視設定
    Call subSetMoveKbn(False)                                 ' 2006/01/12 ADD T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :金額のカンマ編集処理
'        MODULE_ID        :Chg_Kanma
'        IN               :varValue = 金額ﾃﾞｰﾀ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chg_Kanma(ByVal varValue As Variant) As Variant
    
    Dim varWork As Variant
    If Nz(varValue, "") <> "" Then
        varWork = Format(varValue, P_FOMAT_KANMA)
    Else
        varWork = Null
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :電話番号の取得(dataなし時：電話番号⇒携帯番号⇒Fax番号で取得)
'        MODULE_ID        :SetTelData
'        IN               :strTel = 電話番号 /strKeitai = 携帯番号 /strFax = FAX番号
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SetTelData(ByVal varRenKbn As Variant, ByVal varMail As Variant, _
                            ByVal varTel As Variant, ByVal varKeitai As Variant, ByVal varFax As Variant) As Variant
    
    Dim varWork As Variant
    
    Select Case varRenKbn
        Case P_REN_00: varWork = varTel
        Case P_REN_01: varWork = varKeitai
        Case P_REN_02: varWork = varFax
        Case P_REN_10: varWork = varMail
        Case Else:     varWork = Null
    End Select
    SetTelData = varWork
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :ﾔｰﾄﾞ番号存在ﾁｪｯｸ(BeforeUpdate)
'        MODULE_ID        :Yard_BeforeUp
'        IN               :2005/07/23 K.MITSUNO
'        OUT              :intChk=1(桁数ｵｰﾊﾞ時)/2(ｺｰﾄﾞ未入力時/0(その他)
'==============================================================================*
Private Function Yard_BeforeUp(ByVal cmbCtl As ComboBox, ByRef intChk As Integer) As Boolean

    On Error GoTo Yard_BeforeUp_Err

    Dim vCode As Variant
    
    
    Yard_BeforeUp = False
    vCode = Nz(cmbCtl, "")
    
    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            intChk = 1
            GoTo Yard_BeforeUp_End
        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            intChk = 2
            GoTo Yard_BeforeUp_End
        Case Else
            intChk = 0
    End Select
    
    'ファイルへの存在チェック
    If GetYardData(0, cmbCtl) = False Then
        intChk = 2
        Exit Function
    End If
  
  GoTo Yard_BeforeUp_End
  
Yard_BeforeUp_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Yard_BeforeUp_End:
    Yard_BeforeUp = True
    Exit Function
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :ﾔｰﾄﾞ番号存在ﾁｪｯｸ(AfterUpdate)
'        MODULE_ID        :Yard_AfterUp
'        IN               :2005/07/23 K.MITSUNO
'
'==============================================================================*
Private Sub Yard_AfterUp(ByVal cmbCtl As ComboBox, ByVal txtCtl As TextBox)

    On Error GoTo Yard_AfterUp_Err

    Dim vCode   As Variant
    Dim lngRc   As Long

    vCode = Nz(cmbCtl, "")
    If vCode = "" Then
        
        GoTo Yard_AfterUp_End
    End If
    
    cmbCtl = Format(cmbCtl, "000000")
    
    'ファイルへの存在チェック
    lngRc = GetYardData(1, cmbCtl, txtCtl)
    
    GoTo Yard_AfterUp_End
    
Yard_AfterUp_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
Yard_AfterUp_End:
    Exit Sub
    
End Sub
'■□■□■□■□■□■□■□■□■□ ↓↓↓　DB Access系 関数群 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□
'==============================================================================*
'
'        MODULE_NAME      :Chk_NameBefore
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :ctl = 対象ｺﾝﾎﾞﾎﾞｯｸｽ / lngID = 名称ｺｰﾄﾞ
'        OUT              :ｺｰﾄﾞ名称
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_NameBefore(ByVal ctlCmb As ComboBox, ByVal strID As String) As Variant

    Dim strWhere    As String
    Dim vCode       As Variant
    
    'ファイルへの存在チェック
    vCode = Nz(ctlCmb, "")
    If vCode <> "" Then
        strWhere = "NAME_ID = '" & strID & "'"
        strWhere = strWhere & " AND NAME_CODE =" & Format$(vCode, "000")
        
        Chk_NameBefore = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
    End If
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :GetNameData
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :strID = 名称ｺｰﾄﾞ / strCD = 名称ｺｰﾄﾞ
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetNameData(ByRef objDb As Database, ByVal strID As String, _
                                ByVal varCD As Variant, Optional varValue As Variant = "") As Variant
    
    Dim strSQL  As String
    Dim varWork As Variant
    Dim objRs   As Recordset
    
    strSQL = "SELECT NAME_NAME,NAME_RYAK "
    strSQL = strSQL & "FROM dbo_NAME_MAST "
    strSQL = strSQL & "WHERE "
    
    If Nz(varCD, "") <> "" Then
        strSQL = strSQL & "NAME_ID = '" & strID & "' AND NAME_CODE = " & varCD
        
        Set objRs = objDb.OpenRecordset(strSQL, dbDenyRead)
        If Not objRs.EOF Then
            With objRs
                If strID = P_NAME_086 Then
                    varWork = .Fields("NAME_RYAK")
                Else
                    varWork = .Fields("NAME_NAME")
                End If
                .Close
            End With
        Else
            varWork = Null
            objRs.Close
        End If
        Set objRs = Nothing
    Else
        '希望サイズ名対応 2005/06/26 K.MITSUNO
        If varValue <> "" Then
            strSQL = strSQL & "NAME_VALUE_FROM = " & varValue
            Set objRs = objDb.OpenRecordset(strSQL, dbDenyRead)
            If Not objRs.EOF Then
                With objRs
                    varWork = .Fields("NAME_NAME")
                    .Close
                End With
            Else
                varWork = Null
                objRs.Close
            End If
        Else
            varWork = Null
        End If
    End If
    
    GetNameData = varWork
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :ﾜｰｸﾃｰﾌﾞﾙへ対象ﾃﾞｰﾀのInsert
'        MODULE_ID        :Create_Wk
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Create_Wk() As Boolean

    On Error GoTo Create_Wk_On_Err
    
    Dim strSQL        As String
    Dim strSQLWK      As String
    Dim strWhere      As String
    Dim varBumoCD     As Variant  '部門ｺｰﾄﾞ
    Dim varBumoNm     As Variant  '部門名
    Dim varKontS      As Variant  'ｺﾝﾃﾅ番号
    Dim varKyDate     As Variant  '解約日
    Dim varWork       As Variant
    Dim lngRc         As Long
    Dim errLoop       As Error
'    Dim objWs       As Workspace  ' 2005/09/15 DEL T.SUZUKI
    Dim objRs         As Recordset
    Dim objWk         As Recordset
    Dim objDb         As Database
' 2006/01/12 ADD T.SUZUKI Start
    Dim strYukbnName  As String   '受付状態区分名
    Dim strMoveName   As String   '移動区分名
' 2006/01/12 ADD T.SUZUKI End

    Create_Wk = False

    doCmd.SetWarnings False
    
    'DB接続
    Set objDb = CurrentDb

' 2005/11/04 ADD T.SUZUKI Start
    ' ﾜｰｸﾃｰﾌﾞﾙ削除処理
    If fncDeleteWorkTable(objDb, P_TABLENM_FVS510) = False Then Exit Function

    ' ﾜｰｸﾃｰﾌﾞﾙ作成処理
    If fncCreateWorkTable(objDb, P_TABLENM_FVS510) = False Then Exit Function
' 2005/11/04 ADD T.SUZUKI End

' 2005/11/04 DEL T.SUZUKI Start
'    'ﾜｰｸﾃｰﾌﾞﾙの内容を削除
'    strSQL = "DELETE FROM FVS510_W01 ; "
'    objDb.Execute strSQL, dbFailOnError
' 2005/11/04 DEL T.SUZUKI End

' 2005/09/15 DEL T.SUZUKI Start
'    Set objWs = DBEngine.Workspaces(0)
'    objWs.BeginTrans
' 2005/09/15 DEL T.SUZUKI End
    
    '部門名取得
    Call GetBumonName(objDb, varBumoCD, varBumoNm)
    
    '新規受付トランより対象ﾃﾞｰﾀ抽出⇒ﾜｰｸﾃｰﾌﾞﾙへ
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_YUKBN, "       '受付状態区分
' 2006/01/12 ADD T.SUZUKI Start
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_MOVEKBN, "     '移動区分
' 2006/01/12 ADD T.SUZUKI End
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_UKNO, "        '受付番号
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_UKTANTO, "     '受付担当者ｺｰﾄﾞ
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_UKDATE, "      '受付日
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_UKKBN, "       '受付区分
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_KKBN, "        '顧客区分
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_RENKBN, "      '連絡区分
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_TANM, "        '代表者名
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_TAKA, "        '代表者カナ
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_YUBINO, "      '郵便番号
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_ADR_1, "       '顧客住所１
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_ADR_2, "       '顧客住所２
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_ADR_3, "       '顧客住所３
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_BUTU, "        '予定収納物
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_KIBOUSU, "     '希望本数
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_RENJFROM, "    '連絡時間From
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_RENJTO, "      '連絡時間To
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_UCODE, "       '顧客ｺｰﾄﾞ
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_NAME, "        '顧客名
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_KANA, "        '顧客名カナ
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_TEL, "         '電話番号
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_KEITAI, "      '携帯番号
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_FAX, "         'FAX番号
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_MAIL, "        'ﾒｰﾙ
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_YCODE, "       'ﾔｰﾄﾞｺｰﾄﾞ
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_BIKO, "        '備考
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_NO, "          'ｺﾝﾃﾅ番号
    '2005/06/26追加 K.MITSUNO START
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_USAGE, "       '用途
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_SIZE_FROM, "   '希望サイズFrom
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_SIZE_TO, "     '希望サイズTo
    '2005/06/26追加 K.MITSUNO END
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_STEP, "        '段区分
    'strSql = strSql & "dbo_YOUK_TRAN.YOUKT_SYUFROM, "     'ｺﾝﾃﾅ種別From→項目削除2005/06/25
    'strSql = strSql & "dbo_YOUK_TRAN.YOUKT_SYUTO, "       'ｺﾝﾃﾅ種別To→項目削除2005/06/25
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_PRICFROM, "    '使用料From
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_PRICTO,"       '使用料To
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_KUDATE, "      '契約受付日
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_YTDATE, "      '予約取消日
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_KKDATE, "      '希望開始日
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_TKDATE, "      '取置期限日
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_TADATE, "      '問合せ日
    strSQL = strSQL & "dbo_YARD_MAST.YARD_NAME, "         'ヤード名
    strSQL = strSQL & "dbo_YARD_MAST.YARD_ADDR_1, "       'ヤード住所１
    strSQL = strSQL & "dbo_YARD_MAST.YARD_ADDR_2, "       'ヤード住所２
    strSQL = strSQL & "dbo_YARD_MAST.YARD_ADDR_3,"        'ヤード住所３
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_AUTOKBN "      '自動区分  ' 2005/11/04 ADD T.SUZUKI
    strSQL = strSQL & ",dbo_YOUK_TRAN.YOUKT_NEARKBN "      '2007/04/03 ADD Tajima
    strSQL = strSQL & ",dbo_YOUK_TRAN.YOUKT_SEND_CD "      '送付方  2007/08/23 ADD Tajima
    strSQL = strSQL & ",dbo_YOUK_TRAN.YOUKT_GENKBN "       '発生区分2007/08/23 ADD Tajima
    
    strSQL = strSQL & "FROM dbo_YOUK_TRAN LEFT JOIN "
    strSQL = strSQL & "dbo_YARD_MAST ON "
    strSQL = strSQL & "dbo_YOUK_TRAN.YOUKT_YCODE = dbo_YARD_MAST.YARD_CODE "
    
    'Where句生成
    strWhere = Make_Where()
    If Trim$(strWhere) <> "" Then
        strSQL = strSQL & "WHERE " & strWhere
    End If

    'ORDER BY句生成                                                             '2005/07/23 修正 K.MITSUNO
    strSQL = strSQL & " ORDER BY dbo_YOUK_TRAN.YOUKT_YCODE,dbo_YOUK_TRAN.YOUKT_YUKBN ,dbo_YOUK_TRAN.YOUKT_NO,  "
    strSQL = strSQL & " dbo_YOUK_TRAN.YOUKT_STEP,dbo_YOUK_TRAN.YOUKT_SIZE_FROM,dbo_YOUK_TRAN.YOUKT_UKDATE,"
    strSQL = strSQL & " dbo_YOUK_TRAN.YOUKT_UKNO;"

    '予約トラン読込
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

' 2005/09/15 ADD T.SUZUKI Start
    If objRs.EOF = True Then
        MsgBox "該当するデータがありません。", vbInformation, P_ERROR
        objRs.Close
        Set objRs = Nothing
        Exit Function
    End If
' 2005/09/15 ADD T.SUZUKI End

    'ワークテーブルへ書き込み
    strSQLWK = "SELECT * FROM FVS510_W01 "
    Set objWk = objDb.OpenRecordset(strSQLWK, dbOpenDynaset, dbAppendOnly)

    Do Until objRs.EOF
        With objWk
            .AddNew
            .Fields("YOUKT_UKNO") = objRs.Fields("YOUKT_UKNO")                          '受付番号
            .Fields("REPORT_ID") = P_RPT_ID                                             '部門コード
            .Fields("BUMOM_BUMOC") = varBumoCD                                          '部門コード
            .Fields("BUMOM_BUMON") = varBumoNm                                          '部門名
            '受付状態区分コード/名称
            .Fields("YOUKT_YUKBN") = objRs.Fields("YOUKT_YUKBN")
' 2006/01/12 ADD T.SUZUKI Start
            '受付状態区分名取得
            strYukbnName = ""
            strMoveName = ""
            strYukbnName = GetNameData(objDb, P_NAME_080, Nz(objRs.Fields("YOUKT_YUKBN"), Null))

            If Nz(objRs.Fields("YOUKT_MOVEKBN"), 0) <> 0 Then
                '移動区分名取得
                strMoveName = GetNameData(objDb, P_NAME_098, Nz(objRs.Fields("YOUKT_MOVEKBN"), Null))
                If Nz(strMoveName, "") <> "" Then
                    strMoveName = "(" & strMoveName & ")"
                End If
            End If

            strYukbnName = strYukbnName & strMoveName
            .Fields("YOUKT_YUKBN_NM") = strYukbnName                                    '受付状態区分名 + (移動区分名)
' 2006/01/12 ADD T.SUZUKI End
' 2006/01/12 DEL T.SUZUKI Start
'            .Fields("YOUKT_YUKBN_NM") = GetNameData(objDb, P_NAME_080, Nz(objRs.Fields("YOUKT_YUKBN"), Null))
' 2006/01/12 DEL T.SUZUKI End
' 2006/01/12 ADD T.SUZUKI Start
            .Fields("YOUKT_MOVEKBN") = objRs.Fields("YOUKT_MOVEKBN")                    '移動区分
' 2006/01/12 ADD T.SUZUKI End
            '受付担当者名
            varWork = ""                                                                ' 2006/01/27 ADD T.SUZUKI
            lngRc = GetTantoData(objDb, objRs.Fields("YOUKT_UKTANTO"), varWork)
            .Fields("YOUKT_UKTANTO") = varWork
            .Fields("YOUKT_UKDATE") = Format$(objRs.Fields("YOUKT_UKDATE"), "yyyy/mm/dd") '2007/08/23 fix tajima 受付日
            '受付区分
            .Fields("YOUKT_UKKBN") = Nz(GetNameData(objDb, P_NAME_070, objRs.Fields("YOUKT_UKKBN")), Null)
            '顧客区分
            .Fields("YOUKT_KKBN") = Nz(GetNameData(objDb, P_NAME_020, objRs.Fields("YOUKT_KKBN")), Null)
            '連絡区分
            .Fields("YOUKT_RENKBN") = Nz(GetNameData(objDb, P_NAME_085, objRs.Fields("YOUKT_RENKBN")), Null)
            .Fields("YOUKT_TANM") = objRs.Fields("YOUKT_TANM")                          '代表者名
            .Fields("YOUKT_TAKA") = objRs.Fields("YOUKT_TAKA")                          '代表者カナ
            .Fields("YOUKT_YUBINO") = objRs.Fields("YOUKT_YUBINO")                      '郵便番号
            .Fields("YOUKT_ADR_1") = objRs.Fields("YOUKT_ADR_1")                        '顧客住所１
            .Fields("YOUKT_ADR_2") = objRs.Fields("YOUKT_ADR_2")                        '顧客住所２
            .Fields("YOUKT_ADR_3") = objRs.Fields("YOUKT_ADR_3")                        '顧客住所３
            '' 20120704 M.HONDA START
            '' .Fields("YOUKT_BUTU") = objRs.Fields("YOUKT_BUTU")                          '予定収納物
            .Fields("YOUKT_BUTU") = ""                          '予定収納物
            '' 20120704 M.HONDA END
            .Fields("YOUKT_KIBOUSU") = objRs.Fields("YOUKT_KIBOUSU")                    '希望本数
'▼ 2007/08/23 tajima
'   del      .Fields("YOUKT_RENJFROM") = objRs.Fields("YOUKT_RENJFROM")                  '連絡時間From
'   del      .Fields("YOUKT_RENJTO") = objRs.Fields("YOUKT_RENJTO")                      '連絡時間To
'▲ 2007/08/23 tajima
            .Fields("YOUKT_UCODE") = objRs.Fields("YOUKT_UCODE")                        '顧客ｺｰﾄﾞ
            .Fields("YOUKT_NAME") = objRs.Fields("YOUKT_NAME")                          '顧客名
            .Fields("YOUKT_KANA") = objRs.Fields("YOUKT_KANA")                          '顧客名カナ
            '希望連絡先
            .Fields("YOUKT_TEL_PRI") = SetTelData(Format(objRs.Fields("YOUKT_RENKBN"), "00"), objRs.Fields("YOUKT_MAIL"), _
                                                    objRs.Fields("YOUKT_TEL"), objRs.Fields("YOUKT_KEITAI"), objRs.Fields("YOUKT_FAX"))
            .Fields("YOUKT_TEL") = objRs.Fields("YOUKT_TEL")                            '電話番号
            .Fields("YOUKT_KEITAI") = objRs.Fields("YOUKT_KEITAI")                      '携帯番号
            .Fields("YOUKT_FAX") = objRs.Fields("YOUKT_FAX")                            'FAX番号
            .Fields("YOUKT_MAIL") = objRs.Fields("YOUKT_MAIL")                          'ﾒｰﾙ
            .Fields("YOUKT_YCODE") = objRs.Fields("YOUKT_YCODE")                        'ﾔｰﾄﾞｺｰﾄﾞ
            .Fields("YARD_NAME") = objRs.Fields("YARD_NAME")                            'ﾔｰﾄﾞ名
            'ﾔｰﾄﾞ住所
            varWork = ""                                                                ' 2006/01/27 ADD T.SUZUKI
            varWork = objRs.Fields("YARD_ADDR_1") & objRs.Fields("YARD_ADDR_2") & objRs.Fields("YARD_ADDR_3")
            .Fields("YARD_ADDR") = Nz(varWork, Null)
            .Fields("YOUKT_BIKO") = objRs.Fields("YOUKT_BIKO")                          '備考
            .Fields("YOUKT_NO") = objRs.Fields("YOUKT_NO")                              'ｺﾝﾃﾅ番号
            .Fields("YOUKT_STEP") = objRs.Fields("YOUKT_STEP")                          '段区分ｺｰﾄﾞ
            '段区分名
            .Fields("YOUKT_STEP_NM") = Nz(GetNameData(objDb, P_NAME_090, objRs.Fields("YOUKT_STEP")), P_STEP_NASHI)
            
            '追加 2005/06/25 K.MITSUNO START
            '用途名称
            .Fields("YOUKT_USAGE") = Nz(GetNameData(objDb, P_NAME_086, objRs.Fields("YOUKT_USAGE")), Null)
            '希望サイズ名
            .Fields("YOUKT_SIZE_NM") = Nz(GetNameData(objDb, P_NAME_087, "", objRs.Fields("YOUKT_SIZE_FROM")), Null)
            .Fields("YOUKT_SIZE_FROM") = objRs.Fields("YOUKT_SIZE_FROM")
            .Fields("YOUKT_SIZE_TO") = objRs.Fields("YOUKT_SIZE_TO")
            '追加 2005/06/25 K.MITSUNO END
            
            '項目削除→2005/06/26 START
            'ｺﾝﾃﾅ種別From名
            '.Fields("YOUKT_SYUFROM") = GetNameData(objDB, P_NAME_040, Nz(objRs.Fields("YOUKT_SYUFROM"), Null))
            'ｺﾝﾃﾅ種別To名
            '.Fields("YOUKT_SYUTO") = GetNameData(objDB, P_NAME_040, Nz(objRs.Fields("YOUKT_SYUTO"), Null))
            '項目削除→2005/06/26 END
            
            .Fields("YOUKT_PRICFROM") = objRs.Fields("YOUKT_PRICFROM")                  '使用料From
            .Fields("YOUKT_PRICTO") = objRs.Fields("YOUKT_PRICTO")                      '使用料To
            
            'ｺﾝﾃﾅの状態・解約日を取得
            Call GetJouStatus(objDb, objRs.Fields("YOUKT_YCODE"), objRs.Fields("YOUKT_NO"), varKontS, varKyDate)
            
            .Fields("YOUKT_YUKBNNM") = varKontS                                         '状態
            .Fields("CARG_KYDATE") = varKyDate                                          '解約日
            .Fields("YOUKT_KUDATE") = Nz(objRs.Fields("YOUKT_KUDATE"), "")              '契約受付日
            .Fields("YOUKT_YTDATE") = Nz(objRs.Fields("YOUKT_YTDATE"), "")              '予約取消日
            .Fields("YOUKT_KKDATE") = Nz(objRs.Fields("YOUKT_KKDATE"), "")              '希望開始日
            .Fields("YOUKT_TKDATE") = Nz(objRs.Fields("YOUKT_TKDATE"), "")              '取置期限日
            .Fields("YOUKT_TADATE") = Nz(objRs.Fields("YOUKT_TADATE"), "")              '問合せ日

' 2005/11/04 ADD T.SUZUKI Start
            If (objRs.Fields("YOUKT_AUTOKBN") <> 0 And Not IsNull(objRs.Fields("YOUKT_AUTOKBN")) = True) Then
                ' 自動区分が｢NOT NULL｣ または ｢0｣以外
                .Fields("YOUKT_AUTOKBN") = "○"

            ElseIf (objRs.Fields("YOUKT_AUTOKBN") = 0 Or IsNull(objRs.Fields("YOUKT_AUTOKBN")) = True) Then
                ' 自動区分が｢NULL｣ または ｢0｣
                .Fields("YOUKT_AUTOKBN") = Null
            End If
' 2005/11/04 ADD T.SUZUKI End

    '▼2007/04/03 ADD tajima
            If Nz(objRs.Fields("YOUKT_NEARKBN"), 0) = 0 Then
                  .Fields("YOUKT_NEARKBN") = "近隣ヤード紹介不要"
            Else
                  .Fields("YOUKT_NEARKBN") = "近隣ヤード紹介希望"
            End If
    '▲2007/04/03 ADD tajima

    '▼2007/08/23 ADD tajima
        '1.連絡時間fromに受付種別を設定する
            .Fields("YOUKT_RENJFROM") = GetNameData(objDb, "083", Nz(objRs.Fields("YOUKT_GENKBN"), Null)) & " 予約"
        '2.連絡時間toに申込書送付方を設定する
            If Nz(objRs.Fields("YOUKT_SEND_CD"), "") <> "" Then
                .Fields("YOUKT_RENJTO") = "申込書" & _
                    GetNameData(objDb, "012", Nz(objRs.Fields("YOUKT_SEND_CD"), Null)) & "希望"
            End If
    '▲2007/08/23 ADD tajima

            .UPDATE
        End With

        objRs.MoveNext
    Loop
    
    objWk.Close
    objRs.Close
    
'    objWs.CommitTrans  ' 2005/09/15 DEL T.SUZUKI
    
    Create_Wk = True
    GoTo Create_Wk_End

Create_Wk_On_Err:
    If DBEngine.Errors.Count > 0 Then
        For Each errLoop In DBEngine.Errors
            MsgBox "Error Number:" + CStr(errLoop.Number) + vbCrLf + errLoop.Description
        Next
    End If
'    objWs.Rollback  ' 2005/09/15 DEL T.SUZUKI

Create_Wk_End:
    doCmd.SetWarnings True
    Set objWk = Nothing
    Set objRs = Nothing
'    Set objWs = Nothing  ' 2005/09/15 DEL T.SUZUKI
    Set objDb = Nothing
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブル削除処理
'        MODULE_ID        :fncDeleteWorkTable
'        Parameter        :objDb     ：データベースオブジェクト
'                          strTableNm：ワークテーブル名
'        CREATE_DATE      :2005/11/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDeleteWorkTable(ByRef objDb As Database, ByVal strTableNm As String) As Boolean

    Dim strSQL As String

    On Error GoTo fncDeleteWorkTable_Err

    fncDeleteWorkTable = False

    ' テーブルの存在チェック
    If fncChkTableExist(objDb, strTableNm) = True Then

        ' テーブル削除
        objDb.TableDefs.Delete strTableNm
        Select Case Err.Number
            Case 0
            Case 3265
                Err.Clear
            Case Else
                MsgBox "subDropWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
                Err.Clear
        End Select
    End If

    fncDeleteWorkTable = True

fncDeleteWorkTable_Exit:
    Exit Function

fncDeleteWorkTable_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, P_ERROR
    Err.Clear
    GoTo fncDeleteWorkTable_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの存在チェック
'        MODULE_ID        :fncChkTableExist
'        Parameter        :strTableNm：ワークテーブル名
'        CREATE_DATE      :2005/11/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTableExist(ByRef objDb As Database, strTableName As String) As Boolean

    Dim tdf As TableDef

    fncChkTableExist = False

    For Each tdf In objDb.TableDefs
        If tdf.NAME = strTableName Then
            fncChkTableExist = True
            Exit For
        End If
    Next tdf
End Function

'==============================================================================*
'
'        MODULE_NAME      :ワークテーブルの作成
'        MODULE_ID        :fncCreateWorkTable
'        Parameter        :objDb     ：データベースオブジェクト
'                          strTableNm：ワークテーブル名
'        CREATE_DATE      :2005/11/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable(ByRef objDb As Database, ByVal strTableNm As String) As Boolean

    Dim tdfNew    As TableDef
    Dim intCount  As Integer

    On Error GoTo fncCreateWorkTable_Err

    fncCreateWorkTable = False

    Set tdfNew = objDb.CreateTableDef(strTableNm)

    With tdfNew
        Select Case strTableNm
            Case P_TABLENM_FVS510  ' *** 予約受付リスト(ヤード別) ***
                .Fields.Append .CreateField("YOUKT_UKNO", dbText, 11)      ' 受付番号
                .Fields.Append .CreateField("BUMOM_BUMOC", dbText, 3)      ' 部門コード
                .Fields.Append .CreateField("REPORT_ID", dbText, 50)       ' レポートID
                .Fields.Append .CreateField("BUMOM_BUMON", dbText, 50)     ' 部門名称
                .Fields.Append .CreateField("YOUKT_YUKBN", dbLong)         ' 予約受付状態区分コード'(Sortで使用)
' 2006/01/12 ADD T.SUZUKI Start
                .Fields.Append .CreateField("YOUKT_MOVEKBN", dbInteger)    ' 移動区分
                .Fields.Append .CreateField("YOUKT_YUKBN_NM", dbText, 100) ' 予約受付状態区分名
' 2006/01/12 ADD T.SUZUKI End
' 2006/01/12 DEL T.SUZUKI Start
'                .Fields.Append .CreateField("YOUKT_YUKBN_NM", dbText, 50)  ' 予約受付状態区分名
' 2006/01/12 DEL T.SUZUKI End
                .Fields.Append .CreateField("YOUKT_UKTANTO", dbText, 30)   ' 受付担当者名
                .Fields.Append .CreateField("YOUKT_UKDATE", dbText, 15)    ' 受付日
                .Fields.Append .CreateField("YOUKT_UKKBN", dbText, 50)     ' 受付区分名
                .Fields.Append .CreateField("YOUKT_UCODE", dbLong)         ' 顧客ｺｰﾄﾞ
                .Fields.Append .CreateField("YOUKT_NAME", dbText, 36)      ' 顧客名
                .Fields.Append .CreateField("YOUKT_KANA", dbText, 50)      ' 顧客名カナ
                .Fields.Append .CreateField("YOUKT_TEL_PRI", dbText, 60)   ' 希望連絡先
                .Fields.Append .CreateField("YOUKT_TEL", dbText, 15)       ' 電話番号
                .Fields.Append .CreateField("YOUKT_KEITAI", dbText, 15)    ' 携帯番号
                .Fields.Append .CreateField("YOUKT_FAX", dbText, 15)       ' FAX番号
                .Fields.Append .CreateField("YOUKT_KKBN", dbText, 30)      ' 顧客区分名
                .Fields.Append .CreateField("YOUKT_RENKBN", dbText, 30)    ' 連絡区分名
                .Fields.Append .CreateField("YOUKT_TANM", dbText, 18)      ' 代表者名
                .Fields.Append .CreateField("YOUKT_TAKA", dbText, 18)      ' 代表社名カナ
                .Fields.Append .CreateField("YOUKT_YUBINO", dbText, 10)    ' 郵便番号
                .Fields.Append .CreateField("YOUKT_ADR_1", dbText, 36)     ' 顧客住所１
                .Fields.Append .CreateField("YOUKT_ADR_2", dbText, 36)     ' 顧客住所2
                .Fields.Append .CreateField("YOUKT_ADR_3", dbText, 36)     ' 顧客住所3
                .Fields.Append .CreateField("YOUKT_MAIL", dbText, 60)      ' メールアドレス
                .Fields.Append .CreateField("YOUKT_YCODE", dbLong)         ' ヤードコード
                .Fields.Append .CreateField("YARD_NAME", dbText, 36)       ' ヤード名
                .Fields.Append .CreateField("YARD_ADDR", dbText, 54)       ' ヤード住所
                .Fields.Append .CreateField("YOUKT_NEARKBN", dbText, 20)   ' 近隣ﾔｰﾄﾞ 2007/04/03 tajima chg
                .Fields.Append .CreateField("YOUKT_BUTU", dbText, 60)      ' 予定収納物
                .Fields.Append .CreateField("YOUKT_BIKO", dbText, 60)      ' 備考
                .Fields.Append .CreateField("YOUKT_KIBOUSU", dbLong)       ' 希望本数
                .Fields.Append .CreateField("YOUKT_NO", dbLong)            ' コンテナ番号
                .Fields.Append .CreateField("YOUKT_USAGE", dbText, 50)     ' 用途名称
                .Fields.Append .CreateField("YOUKT_SIZE_NM", dbText, 50)   ' 希望サイズ名称
                .Fields.Append .CreateField("YOUKT_SIZE_FROM", dbSingle)   ' 希望サイズFROM
                .Fields.Append .CreateField("YOUKT_SIZE_TO", dbSingle)     ' 希望サイズTO
                .Fields.Append .CreateField("YOUKT_STEP", dbLong)          ' 段区分ｺｰﾄﾞ
                .Fields.Append .CreateField("YOUKT_STEP_NM", dbText, 50)   ' 段区分名 2005/07/23 追加　K.MITSUNO
                .Fields.Append .CreateField("YOUKT_PRICFROM", dbLong)      ' 使用料FROM（要カンマ編集）
                .Fields.Append .CreateField("YOUKT_PRICTO", dbLong)        ' 使用料TO（要カンマ編集）
                .Fields.Append .CreateField("YOUKT_RENJFROM", dbText, 50)  ' 2007/08/23 受付種別とした旧：連絡時間From
                .Fields.Append .CreateField("YOUKT_RENJTO", dbText, 50)    ' 2007/08/23 申込書送付方とした旧：連絡時間To
                .Fields.Append .CreateField("YOUKT_YUKBNNM", dbText, 50)   ' BOX状態
                .Fields.Append .CreateField("CARG_KYDATE", dbText, 50)     ' 解約日
                .Fields.Append .CreateField("YOUKT_KUDATE", dbText, 50)    ' 契約受付日
                .Fields.Append .CreateField("YOUKT_YTDATE", dbText, 50)    ' 予約取消日
                .Fields.Append .CreateField("YOUKT_KKDATE", dbText, 50)    ' 希望開始日
                .Fields.Append .CreateField("YOUKT_TKDATE", dbText, 50)    ' 取置期限日
                .Fields.Append .CreateField("YOUKT_TADATE", dbText, 50)    ' 問合せ日
                .Fields.Append .CreateField("YOUKT_AUTOKBN", dbText, 2)    ' 自動区分(○ OR NULL)

                For intCount = 0 To .Fields.Count - 1
                    If .Fields(intCount).Type = dbText Then
                        .Fields(intCount).AllowZeroLength = True
                    End If
                Next intCount

            Case P_TABLENM_FVS512  ' *** 予約受付リスト(受付日別) ***
                .Fields.Append .CreateField("REPORT_ID", dbText, 50)         ' レポートID
                .Fields.Append .CreateField("BUMOM_BUMOC", dbText, 3)        ' 部門コード
                .Fields.Append .CreateField("BUMOM_BUMON", dbText, 50)       ' 部門名称
                .Fields.Append .CreateField("YOUKT_UKDATE", dbText, 15)      ' 受付日
                .Fields.Append .CreateField("YOUKT_YCODE", dbLong)           ' ヤードコード
                .Fields.Append .CreateField("YARD_NAME", dbText, 36)         ' ヤード名
                .Fields.Append .CreateField("YOUKT_NO", dbLong)              ' コンテナ番号
                .Fields.Append .CreateField("YOUKT_SIZE_NM", dbText, 50)     ' 希望サイズ名称
                .Fields.Append .CreateField("YOUKT_SIZE_FROM", dbSingle)     ' 希望サイズFROM
                .Fields.Append .CreateField("YOUKT_SIZE_TO", dbSingle)       ' 希望サイズTO
                .Fields.Append .CreateField("YOUKT_STEP", dbLong)            ' 段区分ｺｰﾄﾞ
                .Fields.Append .CreateField("YOUKT_STEP_NM", dbText, 50)     ' 段区分名'
' del 2007/08/23 tajima .Fields.Append .CreateField("YOUKT_TADATE", dbText, 15)      ' 問合せ日
                .Fields.Append .CreateField("YOUKT_NAME", dbText, 36)        ' 顧客名
                .Fields.Append .CreateField("YOUKT_TEL", dbText, 15)         ' 電話番号
                .Fields.Append .CreateField("YOUKT_KEITAI", dbText, 15)      ' 携帯番号
                .Fields.Append .CreateField("YOUKT_RENKBN", dbText, 30)      ' 連絡区分名
                .Fields.Append .CreateField("YOUKT_BIKO", dbText, 60)        ' 備考
                .Fields.Append .CreateField("YOUKT_UKTANTO", dbText, 30)     ' 受付担当者名
                .Fields.Append .CreateField("YOUKT_YUKBN", dbLong)           ' 予約受付状態区分コード'(Sortで使用)
                .Fields.Append .CreateField("YOUKT_YUKBN_NM", dbText, 50)    ' 予約受付状態区分名
' 2006/01/12 ADD T.SUZUKI Start
                .Fields.Append .CreateField("YOUKT_MOVEKBN_NM", dbText, 50)  ' 移動区分名
' 2006/01/12 ADD T.SUZUKI End
                .Fields.Append .CreateField("YOUKT_UKNO", dbText, 11)        ' 受付番号
                .Fields.Append .CreateField("YOUKT_AUTOKBN", dbText, 2)      ' 自動区分
                .Fields.Append .CreateField("YOUKT_NEARKBN", dbText, 20)     ' 近隣ﾔｰﾄﾞ 2007/04/03 tajima chg
        
                .Fields.Append .CreateField("YOUKT_FAX", dbText, 15)         ' FAX 2007/08/23 add tajima
                .Fields.Append .CreateField("YOUKT_GEN_NM", dbText, 20)      ' 発生種別 2007/08/23 add tajima
                .Fields.Append .CreateField("YOUKT_SEND_NM", dbText, 20)     ' 申込書送付方 2007/08/23 add tajima
        
                For intCount = 0 To .Fields.Count - 1
                    If .Fields(intCount).Type = dbText Then
                        .Fields(intCount).AllowZeroLength = True
                    End If
                Next intCount
        End Select
    End With

    ' テーブルの追加
    objDb.TableDefs.Append tdfNew

    fncCreateWorkTable = True

fncCreateWorkTable_Exit:
    Exit Function

fncCreateWorkTable_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, P_ERROR
    Err.Clear
    GoTo fncCreateWorkTable_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :ﾜｰｸﾃｰﾌﾞﾙへ対象ﾃﾞｰﾀのInsert
'        MODULE_ID        :Create_Wk2
'        CREATE_DATE      :2005/09/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Create_Wk2() As Boolean

    Dim objDb        As Database
    Dim objRs        As Recordset
    Dim objWk        As Recordset

    Dim strSQL       As String
    Dim strWhere     As String

    Dim varBumoNm    As Variant  ' 部門名
    Dim varBumoCD    As Variant  ' 部門コード

    Dim errLoop      As Error
    Dim varWork      As Variant
    Dim lngRc        As Long

' 2006/01/12 ADD T.SUZUKI Start
    Dim strMoveName   As String   '移動区分名
' 2006/01/12 ADD T.SUZUKI End

    On Error GoTo Create_Wk2_Err

    Create_Wk2 = False
    doCmd.SetWarnings False

    'DB接続
    Set objDb = CurrentDb

' 2005/11/04 ADD T.SUZUKI Start
    ' ﾜｰｸﾃｰﾌﾞﾙ削除処理
    If fncDeleteWorkTable(objDb, P_TABLENM_FVS512) = False Then Exit Function

    ' ﾜｰｸﾃｰﾌﾞﾙ作成処理
    If fncCreateWorkTable(objDb, P_TABLENM_FVS512) = False Then Exit Function
' 2005/11/04 ADD T.SUZUKI End

' 2005/11/04 DEL T.SUZUKI Start
'    'ﾜｰｸﾃｰﾌﾞﾙの内容を削除
'    strSQL = "DELETE FROM FVS512_W01 ; "
'    objDb.Execute strSQL, dbFailOnError
' 2005/11/04 DEL T.SUZUKI End

    '部門名取得
    Call GetBumonName(objDb, varBumoCD, varBumoNm)

    '予約受付トランより対象ﾃﾞｰﾀ抽出⇒ﾜｰｸﾃｰﾌﾞﾙへ
    strSQL = ""
    strSQL = strSQL & "SELECT dbo_YOUK_TRAN.YOUKT_UKDATE, " & Chr(13)     '受付日
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_YCODE, " & Chr(13)      'ﾔｰﾄﾞｺｰﾄﾞ
    strSQL = strSQL & "       dbo_YARD_MAST.YARD_NAME, " & Chr(13)        'ﾔｰﾄﾞ名
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_NO, " & Chr(13)         'ｺﾝﾃﾅ番号
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_SIZE_FROM, " & Chr(13)  '希望サイズFrom
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_SIZE_TO, " & Chr(13)    '希望サイズTo
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_STEP, " & Chr(13)       '段区分
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_TADATE, " & Chr(13)     '問合せ日
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_NAME, " & Chr(13)       '顧客名
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_TEL, " & Chr(13)        '電話番号
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_KEITAI, " & Chr(13)     '携帯番号
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_RENKBN, " & Chr(13)     '連絡区分
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_BIKO, " & Chr(13)       '備考
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_UKTANTO, " & Chr(13)    '受付担当者ｺｰﾄﾞ
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_YUKBN, " & Chr(13)      '受付状態区分
' 2006/01/12 ADD T.SUZUKI Start
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_MOVEKBN, " & Chr(13)    '移動区分
' 2006/01/12 ADD T.SUZUKI End
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_UKNO, " & Chr(13)       '受付番号
    strSQL = strSQL & "       dbo_YOUK_TRAN.YOUKT_AUTOKBN "               '自動区分  ' 2005/11/04 ADD T.SUZUKI
    strSQL = strSQL & "      ,dbo_YOUK_TRAN.YOUKT_NEARKBN "               '2007/04/03 ADD Tajima
    strSQL = strSQL & "      ,dbo_YOUK_TRAN.YOUKT_SEND_CD "               '送付方  2007/08/23 ADD Tajima
    strSQL = strSQL & "      ,dbo_YOUK_TRAN.YOUKT_GENKBN "                '発生区分2007/08/23 ADD Tajima
    strSQL = strSQL & "      ,dbo_YOUK_TRAN.YOUKT_FAX "                   'fax     2007/08/23 ADD Tajima
    strSQL = strSQL & " FROM dbo_YOUK_TRAN LEFT JOIN " & Chr(13)
    strSQL = strSQL & "      dbo_YARD_MAST ON " & Chr(13)
    strSQL = strSQL & "      dbo_YOUK_TRAN.YOUKT_YCODE = dbo_YARD_MAST.YARD_CODE " & Chr(13)

    'Where句生成
    strWhere = Make_Where()
    If Trim$(strWhere) <> "" Then
        strSQL = strSQL & "WHERE " & strWhere
    End If

    'ORDER BY句生成
    strSQL = strSQL & " ORDER BY dbo_YOUK_TRAN.YOUKT_UKDATE, " & Chr(13)  ' 受付日
    strSQL = strSQL & "          dbo_YOUK_TRAN.YOUKT_YCODE, " & Chr(13)   ' ﾔｰﾄﾞｺｰﾄﾞ
    strSQL = strSQL & "          dbo_YOUK_TRAN.YOUKT_NO; "                ' ｺﾝﾃﾅ番号

    '予約受付トラン読込
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    If objRs.EOF = True Then
        MsgBox "該当するデータがありません。", vbInformation, P_ERROR
        objRs.Close
        Set objRs = Nothing
        Exit Function
    End If

    'ワークテーブルへ書き込み
    strSQL = "SELECT * FROM FVS512_W01 "
    Set objWk = objDb.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

    Do Until objRs.EOF
        With objWk

            .AddNew

            .Fields("REPORT_ID") = P_RPT_ID2                                            '帳票ID
            .Fields("BUMOM_BUMOC") = varBumoCD                                          '部門コード
            .Fields("BUMOM_BUMON") = varBumoNm                                          '部門名

            .Fields("YOUKT_UKDATE") = Format$(objRs.Fields("YOUKT_UKDATE"), "yyyy/mm/dd") '2007/08/23 fix tajima 受付日
            .Fields("YOUKT_YCODE") = objRs.Fields("YOUKT_YCODE")                        'ﾔｰﾄﾞｺｰﾄﾞ
            .Fields("YARD_NAME") = objRs.Fields("YARD_NAME")                            'ﾔｰﾄﾞ名
            .Fields("YOUKT_NO") = objRs.Fields("YOUKT_NO")                              'ｺﾝﾃﾅ番号

            '希望サイズ名
            .Fields("YOUKT_SIZE_NM") = Nz(GetNameData(objDb, P_NAME_087, "", objRs.Fields("YOUKT_SIZE_FROM")), Null)
            .Fields("YOUKT_SIZE_FROM") = objRs.Fields("YOUKT_SIZE_FROM")
            .Fields("YOUKT_SIZE_TO") = objRs.Fields("YOUKT_SIZE_TO")

            .Fields("YOUKT_STEP") = objRs.Fields("YOUKT_STEP")                          '段区分ｺｰﾄﾞ
            '段区分名
            .Fields("YOUKT_STEP_NM") = Nz(GetNameData(objDb, P_NAME_090, objRs.Fields("YOUKT_STEP")), P_STEP_NASHI)

'2007/08/23 del tajima            .Fields("YOUKT_TADATE") = objRs.Fields("YOUKT_TADATE")                      '問合せ日
            .Fields("YOUKT_NAME") = objRs.Fields("YOUKT_NAME")                          '顧客名
            .Fields("YOUKT_TEL") = objRs.Fields("YOUKT_TEL")                            '電話番号
            .Fields("YOUKT_KEITAI") = objRs.Fields("YOUKT_KEITAI")                      '携帯番号

            '連絡区分
            .Fields("YOUKT_RENKBN") = Nz(GetNameData(objDb, P_NAME_085, objRs.Fields("YOUKT_RENKBN")), Null)

            .Fields("YOUKT_BIKO") = objRs.Fields("YOUKT_BIKO")                          '備考

            '受付担当者名
            lngRc = GetTantoData(objDb, objRs.Fields("YOUKT_UKTANTO"), varWork)
            .Fields("YOUKT_UKTANTO") = varWork

            '受付状態区分コード/名称
            .Fields("YOUKT_YUKBN") = objRs.Fields("YOUKT_YUKBN")
            .Fields("YOUKT_YUKBN_NM") = GetNameData(objDb, P_NAME_080, Nz(objRs.Fields("YOUKT_YUKBN"), Null))

' 2006/01/12 ADD T.SUZUKI Start
            '移動区分名
            strMoveName = ""
            If Nz(objRs.Fields("YOUKT_MOVEKBN"), 0) <> 0 Then
                '移動区分名取得
                strMoveName = GetNameData(objDb, P_NAME_098, Nz(objRs.Fields("YOUKT_MOVEKBN"), Null))

                If Nz(strMoveName, "") <> "" Then
                    strMoveName = "(" & strMoveName & ")"
                End If
            End If
            .Fields("YOUKT_MOVEKBN_NM") = strMoveName
' 2006/01/12 ADD T.SUZUKI End

            .Fields("YOUKT_UKNO") = objRs.Fields("YOUKT_UKNO")                          '受付番号

' 2005/11/04 ADD T.SUZUKI Start
            If (objRs.Fields("YOUKT_AUTOKBN") <> 0 And Not IsNull(objRs.Fields("YOUKT_AUTOKBN")) = True) Then
                ' 自動区分が｢NOT NULL｣ または ｢0｣以外
                .Fields("YOUKT_AUTOKBN") = "○"

            ElseIf (IsNull(objRs.Fields("YOUKT_AUTOKBN")) = True Or objRs.Fields("YOUKT_AUTOKBN") = 0) Then
                ' 自動区分が｢NULL｣ または ｢0｣
                .Fields("YOUKT_AUTOKBN") = Null
            End If
' 2005/11/04 ADD T.SUZUKI End

    '▼2007/04/03 ADD tajima
            If Nz(objRs.Fields("YOUKT_NEARKBN"), 0) = 0 Then
                  .Fields("YOUKT_NEARKBN") = "紹介不要"
            Else
                  .Fields("YOUKT_NEARKBN") = "紹介希望"
            End If
    '▲2007/04/03 ADD tajima


    '▼2007/08/23 ADD tajima
        '1.受付種別を設定する
            .Fields("YOUKT_GEN_NM") = GetNameData(objDb, "083", Nz(objRs.Fields("YOUKT_GENKBN"), Null)) & "予約"
        '2.申込書送付方を設定する
            If Nz(objRs.Fields("YOUKT_SEND_CD"), "") <> "" Then
                .Fields("YOUKT_SEND_NM") = _
                    GetNameData(objDb, "012", Nz(objRs.Fields("YOUKT_SEND_CD"), Null)) & "希望"
            End If
        '3.FAXを設定する
            .Fields("YOUKT_FAX") = objRs.Fields("YOUKT_FAX")
            
    '▲2007/08/23 ADD tajima

            .UPDATE
        End With

        objRs.MoveNext
    Loop

    objWk.Close
    objRs.Close

    Create_Wk2 = True

    GoTo Create_Wk2_End

Create_Wk2_Err:
    If DBEngine.Errors.Count > 0 Then
        For Each errLoop In DBEngine.Errors
            MsgBox "Error Number:" + CStr(errLoop.Number) + vbCrLf + errLoop.Description
        Next
    End If

Create_Wk2_End:
    doCmd.SetWarnings True
    Set objWk = Nothing
    Set objRs = Nothing
    Set objDb = Nothing
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナ状態の取得
'        MODULE_ID        :ChgDateStyle
'        IN               :objDB = DB接続/varYard = ﾔｰﾄﾞ番号/varKon = ｺﾝﾃﾅ番号
'                          varKontS = ｺﾝﾃﾅ状態/varKyDate = 解約日
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ChgDateStyle(ByVal strDate As String) As Variant
    Dim varWork As Variant
    If strDate <> "" Then
        varWork = Format(strDate, "yyyymmdd")
    Else
        varWork = Null
    End If
    ChgDateStyle = varWork
End Function
'==============================================================================*
'
'        MODULE_NAME      :コンテナ状態の取得
'        MODULE_ID        :GetJouStatus
'        IN               :objDB = DB接続/varYard = ﾔｰﾄﾞ番号/varKon = ｺﾝﾃﾅ番号
'                          varKontS = ｺﾝﾃﾅ状態/varKyDate = 解約日
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub GetJouStatus(ByRef objDb As Database, ByVal varYard As Variant, ByVal varKon As Variant, _
                                ByRef varKontS As Variant, ByRef varKyDate As Variant)
    
    Dim strSQL  As String
    Dim varSta  As Variant
    Dim varDate As Variant
    Dim blnFln  As Boolean      'ｺﾝﾃﾅ確定=True/未確定False
    Dim objKon  As Recordset
    
    
    strSQL = "SELECT TOP 1 "
    strSQL = strSQL & "CARG_KYDATE,CARG_CONTNA "
    strSQL = strSQL & "FROM dbo_CARG_FILE "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CARG_YCODE = " & varYard
    If Nz(varKon, "") <> "" Then
        blnFln = True
        strSQL = strSQL & " AND CARG_NO = " & varKon
    Else
        blnFln = False
    End If
    'ORDER BY句生成
    strSQL = strSQL & " ORDER BY "
    strSQL = strSQL & " NZ(CARG_KYDATE,CARG_FSDATE) DESC "
    
    Set objKon = objDb.OpenRecordset(strSQL, dbReadOnly)
    With objKon
        If Not .EOF Then
            '解約日がNullの時
            If Nz(.Fields("CARG_KYDATE"), "") = "" Then
                varSta = P_JOUTAI_MATI
                varDate = Null
            Else
                '解約日がNull以外でコンテナ状態"ﾚ"無し
                If .Fields("CARG_CONTNA") = -1 Then
                    varSta = P_JOUTAI_OK
                    varDate = Null
                Else
                    varSta = P_JOUTAI_MATI
                    varDate = .Fields("CARG_KYDATE")
                End If
            End If
        Else
            varSta = P_JOUTAI_OK
            varDate = Null
        End If
        .Close
    End With
    
    'コンテナ未確定時は状態は空
    If blnFln = False Then varSta = Null: varDate = Null
    
    Set objKon = Nothing
    
    'ｺﾝﾃﾅ状態
    varKontS = varSta
    '解約日
    varKyDate = varDate
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :GetBumonName
'        IN               :objDB = DB接続 /strBumonCd = 部門ｺｰﾄﾞ /strBumonNm = 部門名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub GetBumonName(ByRef objDb As Database, ByRef varBumonCd As Variant, ByRef varBumonNm As Variant)
    
    Dim strSQL  As String
    Dim objRs   As Recordset
    
    strSQL = "SELECT "
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMOC, " '部門ｺｰﾄﾞ
    strSQL = strSQL & "BUMO_MAST.BUMOM_BUMON "  '部門名称
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "dbo_CONT_MAST INNER JOIN BUMO_MAST ON "
    strSQL = strSQL & "dbo_CONT_MAST.CONT_BUMOC = BUMO_MAST.BUMOM_BUMOC; "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)
    With objRs
        If Not .EOF Then
            varBumonCd = .Fields("BUMOM_BUMOC")
            varBumonNm = .Fields("BUMOM_BUMON")
        Else
            varBumonCd = Null
            varBumonNm = Null
        End If
        .Close
    End With
    Set objRs = Nothing
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :GetTantoData
'        機能             :担当者マスタ情報取得
'        IN               :vTanto = 担当者ｺｰﾄﾞ / vTanNm = 担当者名
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetTantoData(ByRef objDb As Database, ByRef vTanto As Variant, Optional ByRef vTanNm As Variant) As Boolean
    
    Dim strSQL  As String
    Dim objRs   As Recordset
    
    GetTantoData = False
    
    strSQL = "SELECT "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC, "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTN "
    strSQL = strSQL & "FROM TANT_MAST INNER JOIN "
    strSQL = strSQL & "dbo_CONT_MAST ON "
    strSQL = strSQL & "TANT_MAST.TANTM_BUMOC = dbo_CONT_MAST.CONT_BUMOC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"
    
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Exit Function
    Else
        If Nz(objRs("TANTM_TANTN"), "") <> "" Then vTanNm = objRs("TANTM_TANTN")
    End If
    
    GetTantoData = True
    
GetTantoData_End:
    objRs.Close
    Set objRs = Nothing
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :GetYardData
'        機能             :ヤードマスタ問い合わせ
'        IN               :2005/07/23
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetYardData(ByVal iMode As Integer, ByVal cmbCtl As ComboBox, Optional ByRef txtCtl As TextBox) As Boolean
    
    Dim strSQL  As String
    Dim vRet    As Variant
    Dim rsYard  As Recordset
    Dim objDb   As Database
    
On Error GoTo GetYardData_Err
  
    GetYardData = False
  
    If Nz(cmbCtl, "") = "" Then
        GetYardData = True
        GoTo GetYardData_End
    End If
      
    doCmd.SetWarnings False
    
    If iMode = 0 Then
        '存在チェック
        strSQL = strSQL & " YARD_CODE =" & Val(cmbCtl.VALUE)
        vRet = Nz(DLookup("YARD_CODE", "dbo_YARD_MAST", strSQL), "")
        
        If vRet = "" Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, P_KAKUNIN
            GoTo GetYardData_End
        End If
    Else
        'DB接続
        Set objDb = CurrentDb
        
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "YARD_NAME "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "dbo_YARD_MAST "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "YARD_CODE =" & Val(cmbCtl.VALUE)
        
        Set rsYard = objDb.OpenRecordset(strSQL, dbOpenSnapshot)
        If Not (rsYard.BOF And rsYard.EOF) Then
            txtCtl = rsYard.Fields("YARD_NAME")
        End If
        
        rsYard.Close
        Set rsYard = Nothing
        Set objDb = Nothing
    End If
    
    GetYardData = True
    GoTo GetYardData_End

GetYardData_Err:
    If Not rsYard Is Nothing Then rsYard.Close: Set rsYard = Nothing
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

GetYardData_End:
    doCmd.SetWarnings True
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :予約トランよりﾃﾞｰﾀ取得する際のWhere句生成
'        MODULE_ID        :Make_Where
'        IN               :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_Where() As String

    On Error GoTo Make_Where_Err
    
    Dim strYuF      As String
    Dim strYuT      As String
    Dim strSwap     As String
    Dim strDayF     As String
    Dim strDayT     As String
    Dim strWhere    As String
    Dim strYARDF    As String
    Dim strYARDT    As String
    Dim strSwapY    As String

    Dim strMoveF    As String                                                   ' 2006/01/12 ADD T.SUZUKI
    Dim strMoveT    As String                                                   ' 2006/01/12 ADD T.SUZUKI
    
    Dim strTanto    As String                                                   ' 2007/04/03 ADD H.UENO

                                                                                '2005/07/23 In Start K.MITSUNO
    '画面指定値取得(ヤードコード)
    strYARDF = Nz(Me.cmb_YardF, "")
    strYARDT = Nz(Me.cmb_YardT, "")

    'ﾔｰﾄﾞｺｰﾄﾞの大小ﾁｪｯｸ
    If Trim$(strYARDF) <> "" And Trim$(strYARDT) <> "" Then
        'FromとToの選択が逆だった時の為
        If strYARDF > strYARDT Then
            strSwapY = strYARDF
            strYARDF = strYARDT
            strYARDT = strSwapY
        End If
    End If

    'ﾔｰﾄﾞｺｰﾄﾞ条件指定
    If Trim$(strYARDF) <> "" And Trim$(strYARDT) = "" Then
        strWhere = "(YOUKT_YCODE = " & Val(strYARDF) & ") "
    ElseIf Trim$(strYARDF) = "" And Trim$(strYARDT) <> "" Then
        strWhere = "(YOUKT_YCODE = " & Val(strYARDT) & ") "
    ElseIf Trim$(strYARDF) <> "" And Trim$(strYARDT) <> "" Then
        strWhere = " (YOUKT_YCODE BETWEEN " & Val(strYARDF)
        strWhere = strWhere & " AND " & Val(strYARDT) & ") "
    ElseIf Trim$(strYARDF) = "" And Trim$(strYARDT) = "" Then

    End If

    If strWhere <> "" Then
        strWhere = strWhere & " AND "
    End If
                                                                                '2005/07/23 In End K.MITSUNO
    '画面指定値取得(受付状態区分)
    strYuF = Nz(Me.cmb_YuKbnF, "")
    strYuT = Nz(Me.cmb_YuKbnT, "")

    '受付状態ｺｰﾄﾞの大小ﾁｪｯｸ
    If Trim$(strYuF) <> "" And Trim$(strYuT) <> "" Then
        'FromとToの選択が逆だった時の為
        If strYuF > strYuT Then
            strSwap = strYuF
            strYuF = strYuT
            strYuT = strSwap
        End If
    End If

    '受付状態ｺｰﾄﾞ条件指定
    If Trim$(strYuF) <> "" And Trim$(strYuT) = "" Then
        strWhere = strWhere & "(YOUKT_YUKBN = " & Trim$(strYuF) & ") "
    ElseIf Trim$(strYuF) = "" And Trim$(strYuT) <> "" Then
        strWhere = strWhere & "(YOUKT_YUKBN = " & strYuT & ") "
    ElseIf Trim$(strYuF) <> "" And Trim$(strYuT) <> "" Then
        strWhere = strWhere & " (YOUKT_YUKBN BETWEEN " & strYuF
        strWhere = strWhere & " AND " & strYuT & ") "
    ElseIf Trim$(strYuF) = "" And Trim$(strYuT) = "" Then

    End If

' 2006/01/12 ADD T.SUZUKI Start
    '画面指定値取得(移動区分)
    strMoveF = Nz(Me.cmb_MoveKbn_F, "")
    strMoveT = Nz(Me.cmb_MoveKbn_T, "")

    '移動区分ｺｰﾄﾞの大小ﾁｪｯｸ
    If Trim$(strMoveF) <> "" And Trim$(strMoveT) <> "" Then
        'FromとToの選択が逆だった時の為
        If strMoveF > strMoveT Then
            strSwap = strMoveF
            strYuF = strMoveT
            strMoveT = strSwap
        End If
    End If

    '移動区分ｺｰﾄﾞ条件指定
    If Trim$(strMoveF) <> "" And Trim$(strMoveT) = "" Then
        If strWhere <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & "(YOUKT_MOVEKBN = " & Trim$(strMoveF) & ") "
    ElseIf Trim$(strMoveF) = "" And Trim$(strMoveT) <> "" Then
        If strWhere <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & "(YOUKT_MOVEKBN = " & strMoveT & ") "
    ElseIf Trim$(strMoveF) <> "" And Trim$(strMoveT) <> "" Then
        If strWhere <> "" Then strWhere = strWhere & " AND "
        strWhere = strWhere & " (YOUKT_MOVEKBN BETWEEN " & strMoveF
        strWhere = strWhere & " AND " & strMoveT & ") "
    ElseIf Trim$(strMoveF) = "" And Trim$(strMoveF) = "" Then
        ' 指定なし
    End If
' 2006/01/12 ADD T.SUZUKI End

    '予約受付日のチェック
    Call ChkDateFromTo(Me.txt_UkDateF, Me.txt_UkDateT)
    strWhere = strWhere & Make_DateWhere(strWhere, Me.txt_UkDateF, Me.txt_UkDateT, "YOUKT_UKDATE", True) '2007/08/23 fix tajima

    '契約受付日のチェック
    Call ChkDateFromTo(Me.txt_KuDateF, Me.txt_KuDateT)
    strWhere = strWhere & Make_DateWhere(strWhere, Me.txt_KuDateF, Me.txt_KuDateT, "YOUKT_KUDATE")

    '予約取消日のチェック
    Call ChkDateFromTo(Me.txt_YtDateF, Me.txt_YtDateT)
    strWhere = strWhere & Make_DateWhere(strWhere, Me.txt_YtDateF, Me.txt_YtDateT, "YOUKT_YTDATE")

    '希望開始日のチェック
    Call ChkDateFromTo(Me.txt_KkDateF, Me.txt_KkDateT)
    strWhere = strWhere & Make_DateWhere(strWhere, Me.txt_KkDateF, Me.txt_KkDateT, "YOUKT_KKDATE")

    '取置期限日のチェック
    Call ChkDateFromTo(Me.txt_TkDateF, Me.txt_TkDateT)
    strWhere = strWhere & Make_DateWhere(strWhere, Me.txt_TkDateF, Me.txt_TkDateT, "YOUKT_TKDATE")

    '問合せ日のチェック
    Call ChkDateFromTo(Me.txt_TaDateF, Me.txt_TaDateT)
    strWhere = strWhere & Make_DateWhere(strWhere, Me.txt_TaDateF, Me.txt_TaDateT, "YOUKT_TADATE")

' 2005/11/04 ADD T.SUZUKI Start
    ' 自動または手動
    Select Case Me.cmb_AUTO
        Case P_ALL_VALUE   ' *** 全て ***
            ' 条件指定なし

        Case P_AUTO_VALUE  ' *** 自動のみ ***
            ' ｢0｣以外 または ｢Not Null｣
            If Nz(strWhere, "") <> "" Then
                strWhere = strWhere & " AND "
            End If

            strWhere = strWhere & " (YOUKT_AUTOKBN <> 0 AND YOUKT_AUTOKBN IS NOT NULL) "

        Case P_HAND_VALUE  ' *** 自動を除く ***
            ' ｢0｣以外 または ｢Null｣
            If Nz(strWhere, "") <> "" Then
                strWhere = strWhere & " AND "
            End If

            strWhere = strWhere & " (YOUKT_AUTOKBN = 0 OR YOUKT_AUTOKBN IS NULL) "
    End Select
' 2005/11/04 ADD T.SUZUKI End

' 2007/04/03 ADD H.UENO Start
    strTanto = Nz(Me.cmb_Tanto)
    If Trim$(strTanto) <> "" Then
        strWhere = strWhere & " AND YOUKT_UKTANTO = '" & strTanto & "' "
    End If
' 2007/04/03 ADD H.UENO End

'▼ 2007/08/23 add tajima
    If Nz(Me.cmb_受付種別.VALUE, "") <> "" Then
        If CInt(Me.cmb_受付種別.VALUE) = 0 Then
            strWhere = strWhere & " AND ( ISNULL(YOUKT_GENKBN) OR YOUKT_GENKBN = 0 ) "
        Else
            strWhere = strWhere & " AND YOUKT_GENKBN = " & CInt(Me.cmb_受付種別.VALUE)
        End If
    End If
    
    If Nz(Me.cmb_書類送付方.VALUE, "") <> "" Then
        strWhere = strWhere & " AND YOUKT_SEND_CD = " & CInt(Me.cmb_書類送付方.VALUE)
    End If
'▲ 2007/08/23 add tajima
    
    Make_Where = strWhere
    GoTo Make_Where_End

Make_Where_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, P_ERROR
    
Make_Where_End:
    Exit Function

End Function
'==============================================================================*
'
'        MODULE_NAME      :日付欄Where句生成
'        MODULE_ID        :Make_DateWhere
'        IN               :
'          2007/08/23 fix tajima ...anIs時間 対応
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_DateWhere(ByVal strSQL As String, _
                                ByRef ctlTextF As TextBox, _
                                ByRef ctlTextT As TextBox, _
                                ByVal strDBField As String, _
                                Optional anIs時間 As Boolean = False) As String
    
    Dim strWhere As String
    
    'この処理までにWhere句生成されてなくて、日付From条件有りの時
    If strSQL <> "" And ctlTextF <> "" Then
        strWhere = strWhere & " AND "
    End If
    '日付From指定有り
    If ctlTextF <> "" Then
        If anIs時間 = True Then
            strWhere = strWhere & strDBField & " >= CDATE(#" & ctlTextF & " 00:00:00#) "
        Else
            strWhere = strWhere & strDBField & " >= CDATE(#" & ctlTextF & "#) "
        End If
    End If
    '既にWhere句作成されており、日付To条件有り
    If strSQL <> "" And ctlTextT <> "" Then
        strWhere = strWhere & " AND "
    End If
    '日付To条件有り
    If ctlTextT <> "" Then
        If anIs時間 = True Then
            strWhere = strWhere & strDBField & " <= CDATE(#" & ctlTextT & " 23:59:59#) "
        Else
            strWhere = strWhere & strDBField & " <= CDATE(#" & ctlTextT & "#) "
        End If
    End If
    
    Make_DateWhere = strWhere
    
End Function

'■□■□■□■□■□■□■□■□■□ ↓↓↓　入力ﾁｪｯｸ 関数 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□
'==============================================================================*
'
'        MODULE_NAME      :日付入力ﾁｪｯｸ（From/Toを逆に入力した時の対応）
'        MODULE_ID        :ChkDateFromTo
'        IN               :
'
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ChkDateFromTo(ByRef ctlTextF As TextBox, ByRef ctlTextT As TextBox)
    
    Dim strDayF As String
    Dim strDayT As String
    Dim strSwap As String
    
    If ctlTextF <> "" And ctlTextT <> "" Then
        strDayF = Format(ctlTextF, "yyyy/mm/dd")
        strDayT = Format(ctlTextT, "yyyy/mm/dd")
        If strDayT <> "" And strDayF > strDayT Then
            '日付選択範囲の大小逆転時のデータ抽出ミス修正のため追加
            strSwap = strDayF
            strDayF = strDayT
            strDayT = strSwap
            '参照元へ再セット
            ctlTextF = strDayF
            ctlTextT = strDayT
        End If
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub dateCheck(ByRef ctlTxt As TextBox, Cancel As Integer)
    
    On Error GoTo DateCheck_Err
    
    Dim vCode As Variant
    Dim vDATE As Variant
      
    vCode = Nz(ctlTxt, "")
    
    '何も入力されていない時は、処理をしない
    If vCode = "" Then GoTo DateCheck_End
    vCode = Check_Date(vCode)
    vCode = CMCheck_Date(vCode)
    
    '入力チェック
    Select Case CmBCheck_Date(vCode)
        Case 2
             '桁数ｴﾗｰ
             MsgBox "桁数オーバーしています。", vbExclamation, P_ERROR
             Cancel = True
             SendKeys "{ESC}"
             GoTo DateCheck_End
        Case 3, 9
             MsgBox "日付を正しく入力してください。", vbExclamation, P_ERROR
             Cancel = True
             SendKeys "{ESC}"
             GoTo DateCheck_End
        Case Else
    End Select
    
    GoTo DateCheck_End
    
DateCheck_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
DateCheck_End:
    Exit Sub
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :日付項目のチェックを一括でする
'        MODULE_ID        :
'        CREATE_DATE      :
'
'引数
'oText 対象項目
'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
'返り値は、
'２：桁数が溢れている時
'３：日付と判断できない時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer

    On Error GoTo CmBCheck_Date_Err
    
    Dim sYear   As String
    Dim intLen  As Integer
  
    CmBCheck_Date = 2
  
    If IsMissing(ilen) = False Then
       intLen = Nz(CmNumeric(ilen), 0)
    Else
       intLen = 10
    End If
    
    oText = Nz(oText, "")
    '桁数チェック
    If CmLenb(oText) > intLen Then GoTo CmBCheck_Date_End
    CmBCheck_Date = 3
    If Nz(oText, "") = "" Then GoTo CmBCheck_Date_End
    
    '論理チェック
    If IsDate(oText) = False Then GoTo CmBCheck_Date_End
    
    sYear = Format(oText, "yyyy")
    If CmLenb(sYear) <> 4 Then GoTo CmBCheck_Date_End
    
    CmBCheck_Date = 0
    
    GoTo CmBCheck_Date_End
    
CmBCheck_Date_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    CmBCheck_Date = 9
    
CmBCheck_Date_End:
    Exit Function
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :
'
'Check_Dateでは'198/03といった形が変換されるための補助検査
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)
  
    Dim sYear As String
    
    CMCheck_Date = Null
    
    If Nz(vDATE, "") = "" Then Exit Function
    If IsDate(vDATE) = False Then Exit Function
    
    sYear = Format(vDATE, "yyyy")
    If CmLenb(sYear) <> 4 Then Exit Function
    
    CMCheck_Date = vDATE
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :数値項目のチェックを一括でする。☆★他フォームでも個々に宣言されている
'        MODULE_ID        :CmBCheck_Num
'        CREATE_DATE      :
'
'返り値は、
'１：数値と判断できない時
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'
'引数 oText ←判定する値
'     iLen  ←整数桁
'     isLen ←小数桁(省略された時は、小数点桁数無し)
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, Optional ByVal isLEN As Variant) As Integer
    
On Error GoTo err_rtn
    
    Dim vINT    As Variant
    Dim vsINT   As Variant
    Dim i       As Integer
      
    Dim dTEXT   As Double
    Dim dFROM   As Double
    Dim dTO     As Double
    Dim vRet    As Variant

    CmbCheck_Num = 9
    
    If Nz(CmNumeric(oText), "") = "" Then       '数値と判断できないのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If
    
    If CmChkCher(CStr(Nz(oText)), "A") = False Then '全角文字が入力されているのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If
        
    '小数点の値が設定されている時のチェック
    vRet = Nz(InStr(Abs(oText), "."), 0)
    
    If IsMissing(isLEN) = False Then
        If vRet <> 0 Then
            If isLEN < CmLenb(Abs(oText)) - vRet Then    '小数点桁数が設定桁数を超えているエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
            End If
        End If
    Else
        If CmLenb(Abs(CmNumeric(oText))) > ilen Then           '文字数超過のためエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
        End If
    
        If vRet <> 0 Then                         '小数点の入力があるのでエラー
          CmbCheck_Num = 2
          GoTo Exit_rtn
        End If
    End If
        
    '桁チェック
    i = 0
    For i = 1 To ilen: vINT = vINT & "9": Next
    
    If IsMissing(isLEN) = False Then
        vsINT = "."
        i = 0
        For i = 1 To isLEN
            vsINT = vsINT & "9"
        Next
    End If
      
    vINT = vINT & vsINT
    vINT = CmNumeric(vINT)
    
    dTEXT = CmNumeric(oText)
    dFROM = vINT * -1
    dTO = vINT
    
    If dFROM > dTEXT Or dTO < dTEXT Then        '論理範囲超過のためエラー
        CmbCheck_Num = 2
        GoTo Exit_rtn
    End If
    
    CmbCheck_Num = 0

Exit_rtn:
    Exit Function
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmbCheck_Num = 9
    GoTo Exit_rtn
End Function
'==============================================================================*
'
'        MODULE_NAME      :画面項目の入力チェック
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2005/12/14
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck() As Boolean

    fncInputCheck = False

    With Me
        ' 必須入力チェック
        ' 受付状態区分(From-To)
        If (Nz(.cmb_YuKbnF, "") = "" And Nz(.cmb_YuKbnT, "") = "") Then
            MsgBox "受付状態区分を入力して下さい。", vbInformation, P_ERROR
            .cmb_YuKbnF.SetFocus
            Exit Function
        End If
    End With

    fncInputCheck = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :移動区分の可視設定
'        MODULE_ID        :fncInputCheck
'        CREATE_DATE      :2006/01/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetMoveKbn(ByVal blnFLG As Boolean)

    With Me
        ' 移動区分の可視設定
        .cmb_MoveKbn_F = ""
        .txt_MoveKbn_F = ""
        .cmb_MoveKbn_T = ""
        .txt_MoveKbn_T = ""

        .lbl_MoveKbn.Visible = blnFLG
        .cmb_MoveKbn_F.Visible = blnFLG
        .txt_MoveKbn_F.Visible = blnFLG
        .cmb_MoveKbn_T.Visible = blnFLG
        .txt_MoveKbn_T.Visible = blnFLG
    End With
End Sub

' 2007/04/03 ADD H.UENO Start
'==============================================================================*
'
'        MODULE_NAME      :AfterUpdateCommon
'        機能             :更新後イベント時の処理
'                         :数値項目は０サプレスした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'                         :③1=数値項目 2=日付項目 3=テキスト項目
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function AfterUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer) As String

    AfterUpdateCommon = ""
    
    '体裁を整える
    Select Case intKind
    Case 1  '数値項目は０埋め
        objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    Case 2
        AfterUpdateCommon = Format(VarChangeDate, "yyyy/mm/dd")
    Case Else
        AfterUpdateCommon = Nz(objTarget)
    End Select
    
    'コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        If Nz(objTarget) <> "" And _
           objTarget.ListCount <> 0 _
        Then
            AfterUpdateCommon = Nz(objTarget.Column(1, objTarget.ListIndex))
        End If
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :ChkComboInput
'        機能             :コンボボックス入力チェック
'        IN               :第一引数：対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :0=正常
'                          1=入力規則に反している
'                          2=入力値がリストに存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer, ByRef Cancel As Integer) As Integer

    Dim strCode     As String
    Dim intRet      As Integer
    Dim intCount    As Integer  '06.01.22 SHIBAZAKI 追加
    Dim strValue    As String   '06.01.22 SHIBAZAKI 追加
    Dim blnFound    As Boolean  '06.01.22 SHIBAZAKI 追加
    
    intRet = 0
    
    strCode = Nz(cmbCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkComboInput_End
    End If
    
    '入力チェック
    Select Case CmNumCheck(strCode, intMaxLength, 0)
        Case 3
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            intRet = 1
            GoTo ChkComboInput_End
        Case 2
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            SendKeys "{ESC}"
            intRet = 1
            GoTo ChkComboInput_End
    End Select
    
    '入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        '06.01.22 SHIBAZAKI 追加↓
        blnFound = False
        '入力値を０サプレス
        strValue = Format(cmbCheck, Mid(P_ZERO_SUPPRESS, 1, intMaxLength))
        'コンボボックス内にコードが存在してる？
        For intCount = 0 To cmbCheck.ListCount - 1
            'コード値比較
            If cmbCheck.Column(0, intCount) = strValue Then
                blnFound = True
                Exit For
            End If
        Next
        '06.01.22 SHIBAZAKI 追加↑
        
        '06.01.22 SHIBAZAKI 変更
        'If cmbCheck.ListIndex = -1 Then
        If Not blnFound Then
            MsgBox "該当するコードがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
            intRet = 2
            GoTo ChkComboInput_End
        End If
    End If
    
ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If
    
    ChkComboInput = intRet

End Function
' 2007/04/03 ADD H.UENO End
'★☆★☆★☆★☆★☆★★☆☆ 2007/08/23 add tajima ★☆★☆★☆★☆★☆★☆★☆
'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_AfterUpdate
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_AfterUpdate()
    
    Call Chk_NameAfter(Me.cmb_受付種別, Me.txt_受付種別, "083")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_BeforeUpdate
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_BeforeUpdate(Cancel As Integer)

    If ControlCrear(Me.cmb_受付種別, Me.txt_受付種別) = False Then Exit Sub
    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_受付種別, "083") = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_受付種別_KeyPress
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付種別_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_受付種別, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_書類送付方_AfterUpdate
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_書類送付方_AfterUpdate()

    Call Chk_NameAfter(Me.cmb_書類送付方, Me.txt_書類送付方, "012")

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_書類送付方_BeforeUpdate
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_書類送付方_BeforeUpdate(Cancel As Integer)

    If ControlCrear(Me.cmb_書類送付方, Me.txt_書類送付方) = False Then Exit Sub
    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_書類送付方, "012") = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_書類送付方_KeyPress
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_書類送付方_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_書類送付方, KeyAscii, 2)
End Sub

