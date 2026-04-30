VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS430"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'==============================================================================*
'   プログラム名　　：　顧客マスタ保守
'   プログラムＩＤ　：
'   作　成　日　　　：  1999/10/22
'   作　成　者　　　：  αシステム
'   Ver             ：  0.0
'==============================================================================*
'修正履歴
'   修　正　日　　　：2004/11/24
'   修　正　者　　　：N.MIURA
'   修　正　内　容　：フォーム名変更
'   Ver             ：  0.1
'==============================================================================*
'修正履歴
'   修　正　日　　　：2004/12/07
'   修　正　者　　　：N.MIURA
'   修　正　内　容　：備考桁数変更
'   Ver             ：  0.2
'==============================================================================*
'修正履歴
'   修　正　日　　　：2005/06/17
'   修　正　者　　　：c.iguchi
'   修　正　内　容　：備考桁数変更
'   Ver             ：  0.3
'==============================================================================*
'修正履歴
'   修　正　日　　　：2005/07/13
'   修　正　者　　　：T.SUZUKI
'   修　正　内　容　：1.顧客マスタの項目追加による画面への項目追加
'                    2.新規項目へのデータ登録・更新処理追加
'   Ver             ：  0.4
'==============================================================================*
'修正履歴
'   修　正　日　　　：2006/01/28
'   修　正　者　　　：T.SUZUKI
'   修　正　内　容　：画面に新顧客コード追加
'   Ver             ：  0.5
'==============================================================================*
'修正履歴
'   修　正　日　　　：2006/02/07
'   修　正　者　　　：T.SUZUKI
'   修　正　内　容　：全画面から呼び出された際の処理を修正(受付入力対応)
'   Ver             ：  0.6
'==============================================================================*
'修正履歴
'   修　正　日　　　：2006/05/17
'   修　正　者　　　：tajima
'   修　正　内　容　：1.顧客マスタ・連絡先担当者名を36→18へ
'                 　：2.既存の連絡先担当者名を保存するときも18バイトチェックを行う
'                 　：3.郵便番号変換による住所１の桁溢れ対応
'                 　：4.削除にこの顧客コードで予約と契約があるかチェック
'   Ver             ：  0.7
'==============================================================================*
'修正履歴
'   修　正　日　　　：2006/06/24
'   修　正　者　　　：tajima
'   修　正　内　容　：顧客問合わせ呼出対応
'   Ver             ：  0.8
'==============================================================================*
'修正履歴
'   修　正　日　　　：2007/01/12
'   修　正　者　　　：tajima
'   修　正　内　容　：削除ボタンを一時無効化※プロパティ変更のみ
'   Ver             ：  0.8a
'==============================================================================*
'修正履歴
'   修　正　日　　　：2007/06/10
'   修　正　者　　　：tajima
'   修　正　内　容　：顧客情報変更対応
'                   ：1.顧客マスタ作成情報、更新情報対応
'                   ：2.画面項目「変更内容」「変更日付」追加
'   Ver             ：  0.9
'==============================================================================*
'修正履歴
'   修　正　日　　　：2010/04/19
'   修　正　者　　　：SHIBAZAKI
'   修　正　内　容　：勤務先を追加
'   Ver             ：  1.0
'==============================================================================*
'修正履歴
'   修　正　日　　　：2010/04/28
'   修　正　者　　　：SHIBAZAKI
'   修　正　内　容　：生年月日のバグを修正
'   Ver             ：  1.1
'==============================================================================*
'修正履歴
'   修　正　日　　　：2011/07/29
'   修　正　者　　　：M.RYU
'   修　正　内　容　：保存するとき、電話番号・携帯番号・FAX番号をTRIMでｽﾍﾟｰｽを弾く
'   Ver             ：  1.2
'==============================================================================*
'==============================================================================*
'修正履歴
'   修　正　日　　　：2011/11/29
'   修　正　者　　　：M.RYU
'   修　正　内　容　：書類送付先担当者名⇒36バイトチェック
'   Ver             ：1.3
'==============================================================================*
'修正履歴
'   修　正　日　　　：2012/11/21
'   修　正　者　　　：K.SHIBUYA
'   修　正　内　容　：顧客名カナのしたに文言追加
'   Ver             ：1.4
'==============================================================================*
'修正履歴
'   修　正　日　　　：2013/06/11
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：勤務先カナの使用不可
'   Ver             ：1.5
'==============================================================================*
'修正履歴
'   修　正　日　　　：2013/09/19
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：支払方法の初期表示をコンビニから振込へ変更
'                 　：他の支払い方法からコンビニへの変更不可へ
'   Ver             ：1.6
'==============================================================================*
'修正履歴
'   修　正　日　　　：2014/11/15
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：性別区分を追加
'                 　：
'   Ver             ：1.7
'==============================================================================*
'修正履歴
'   修　正　日　　　：2018/01/28
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：仮想口座マスタへのデータ追加
'                 　：
'   Ver             ：1.8
'==============================================================================*
'修正履歴
'   修　正　日　　　：2018/04/23
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：保存時に印刷メッセージを表示されないように対応
'                 　：
'   Ver             ：1.9
'==============================================================================*
'修正履歴
'   修　正　日　　　：2019/05/01
'   修　正　者　　　：Y.WADA
'   修　正　内　容　：あおぞら銀行仮想口座対応・仮想口座のテーブル名変更(KOUZ_MAST->KOUZ_MAST_J)
'                 　：
'   Ver             ：2.0
'==============================================================================*
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS430"

Private Const strmsg保存確認        As String = "保存します。よろしいですか？"
Private Const strmsg取消確認        As String = "データ入力を取り消します。よろしいですか？"
Private Const strmsg削除確認        As String = "データを削除します。よろしいですか？"
Private Const strmsg終了確認        As String = "終了します。よろしいですか？"
Private Const strmsg更新エラー      As String = "更新中にエラーが発生しました。" & vbCrLf & "（処理は中断されました。）"
Private Const strmsgクリア確認      As String = "クリアされます。よろしいですか？"
Private Const strMSGデータなし      As String = "存在しないコードです。"
Private Const strmsg重複            As String = "既に存在するコードです。"
Private Const strMSG桁数オーバー    As String = "桁数オーバーしています。"
Private Const strmsg締日            As String = "締日は３０日以内で入力してください。"
Private Const strMsg_kaku           As String = "確認"
Private Const strMsg_err            As String = "エラー"
Private Const strMSG入力不正        As String = "を入力してください。"
Private Const strMSGコンビニ        As String = "コンビニへは変更できません。"             ''2013/09/19 M.HONDA

Private intModeflg                  As Integer      '1=新規,2=更新,3=削除
Private Const intModeInsert         As Integer = 1
Private Const intModeUpdate         As Integer = 2
Private Const intModeDelete         As Integer = 3

Private blnDataUpdate               As Boolean

Dim objDb                           As Database
Dim objRs                           As Recordset

Dim WRI_ObjWS                       As Workspace
Dim WRI_objDb                       As Database

Public blnFormLoadFlg               As Boolean      'フォームが読込まれたかどうかのフラグ
Private oCurrent                    As Object       '前回のオブジェクト格納変数

Dim vBack_Key                       As Variant

' 2005/07/13 ADD T/SUZUKI Start
Private Form_OpenArgs               As Variant      ' 呼出元からのパラメータ
'Private pintOpenMode                As Integer     ' 2006/02/07 DEL T.SUZUKI
Private pstrUSER_CODE               As String
' 2005/07/13 ADD T/SUZUKI End

' 2006/02/07 ADD T.SUZUKI Start
' 全画面から遷移している
' ※True：前画面から顧客コードを受け取っている
Private pblnReceiptMode             As Boolean

' パラメータでもらった呼び出し元画面名
Private pstrFormID                  As String
' 2006/02/07 ADD T.SUZUKI End

' 2007/06/10 ADD tajima
'ADO定数（アクティブXを参照設定しないためローカルで定義する）
Private Const adCmdText = 1
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1
Private Const adOpenDynamic = 2
Private Const adLockPessimistic = 2
Private Const adCmdStoredProc = 4
'変更前処理で入力した日付を保存する
Private var変更日付 As Variant
Private is顧客情報変更 As Boolean

'変更前の集金区分を保持する
Private strKkbn As String   ''213/09/19 M.HONDA INS

'==============================================================================*
'
'        MODULE_NAME      :cmb_集金区分_AfterUpdate イベント
'        MODULE_ID        :cmb_集金区分_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_性別区分_AfterUpdate()

    Dim strRyak     As String
    Dim strWhere    As String
    Dim vCode       As Variant
    
    vCode = Nz(Me.cmb_性別区分, "")
    If vCode = "" Then
        Me.txt_性別区分 = ""
        Exit Sub
    End If
    
    strRyak = s_SelectRyak(Me.cmb_性別区分, "302")
    Me.txt_性別区分 = strRyak

    blnDataUpdate = True
End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_性別区分_BeforeUpdate イベント
'        MODULE_ID        :cmb_性別区分_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_性別区分_BeforeUpdate(Cancel As Integer)
    
    Dim vCode       As Variant
    Dim strRyak     As String
    Dim strWhere    As String

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.cmb_性別区分, "")
    If vCode = "" Then
        Cancel = True
        MsgBox "性別区分" & strMSG入力不正, vbInformation, strMsg_kaku
        Me.txt_性別区分 = ""
        Exit Sub
    End If

    strRyak = s_SelectRyak(Me.cmb_性別区分, "302")

    If strRyak = "" Then
        MsgBox strMSGデータなし, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    If Trim(Me.cmb_性別区分) <> "" Then
        If Trim(Len(Me.cmb_性別区分)) > 2 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub



'==============================================================================*
'
'        MODULE_NAME      :cmd_KYubin_Click イベント
'        MODULE_ID        :cmd_KYubin_Click
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_KYubin_Click()
    Call MSZ0000_YUBIB(Me.NAME, Me.txt_USER_KPOST.NAME, Me.txt_USER_KADR_1.NAME, _
                                Nz(Me.txt_USER_KPOST.VALUE), Nz(Me.txt_USER_KADR_1.VALUE))

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate イベント
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load イベント
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    doCmd.Maximize
    blnFormLoadFlg = False

    Set objDb = CurrentDb
    
    ' 処理モード保持(1：新規/2：修正/3：削除)
    intModeflg = Me.opg_処理選択.VALUE

    'フォームの初期化処理
    Call CLR_RTN

    ' 画面項目の使用可否設定
    Call Etc_Lock_Rtn

    ' 【その他連絡先】の使用可否、背景色設定
    Call subTxtControl
'''    Call chk_入力_Click

    Call s_bottunset(intModeflg)

' 2006/02/07 REP T.SUZUKI Start
    If Nz(pstrUSER_CODE, "") <> "" Then
'' 2005/07/13 ADD T.SUZUKI Start
''    If pintOpenMode = 1 Then
' 2006/02/07 REP T.SUZUKI End
        Me.opt_新規.Enabled = False
        Me.opt_削除.Enabled = False

        Me.cmb_顧客コード = pstrUSER_CODE
        Call cmb_顧客コード_AfterUpdate
    End If
' 2005/07/13 ADD T.SUZUKI End

    blnFormLoadFlg = True
End Sub

' 2005/07/13 ADD T/SUZUKI Start
'==============================================================================*
'
'        MODULE_NAME      :Form_Open イベント
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim lngRc As Long                                       ' 2006/02/07 ADD T.SUZUKI

    Form_OpenArgs = Nz(OpenArgs, "")
'    pintOpenMode = 0                                       ' 2006/02/07 DEL T.SUZUKI
    pstrUSER_CODE = ""
    ' 変更フラグを初期化
    pblnChangeFlg = False                                   ' 2006/02/07 ADD T.SUZUKI
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
    End Select

    If Form_OpenArgs <> "" Then
' 2006/02/07 ADD T.SUZUKI Start
        lngRc = InStr(Form_OpenArgs, ",")
        If lngRc = 0 Then
            pstrUSER_CODE = Form_OpenArgs                   ' 顧客コード
        Else
            pstrUSER_CODE = Left(Form_OpenArgs, lngRc - 1)  ' 顧客コード
            pstrFormID = Mid(Form_OpenArgs, lngRc + 1)      ' 呼出元画面ID
        End If
' 2006/02/07 ADD T.SUZUKI End
' 2006/02/07 DEL T.SUZUKI Start
'        pintOpenMode = Mid(Form_OpenArgs, 1, 1)   ' 画面オープンモード
'        pstrUSER_CODE = Mid(Form_OpenArgs, 2, 6)  ' 顧客コード
'        pblnChangeFlg = False                     ' 変更フラグを初期化
' 2006/02/07 DEL T.SUZUKI End

        ' 修正モードに設定
        Me.opg_処理選択.VALUE = 2
    End If

' 2006/02/07 ADD T.SUZUKI Start
    If Nz(pstrUSER_CODE, "") <> "" Then
        pblnReceiptMode = True
    Else
        pblnReceiptMode = False
    End If
' 2006/02/07 ADD T.SUZUKI End
End Sub
' 2005/07/13 ADD T/SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :Form_KeyDown イベント
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_Form_KeyDown

    'SHIFT、CTRL、ALTのいずれかが押されている場合は何もしない
    If Shift > 0 Then
        Exit Sub
    End If

    ' 顧客問合わせの呼出 2006/06/24 add tajima
    Call CmReprMod.OpenFVS90n(KeyCode, Null, Null, Me.cmb_顧客コード.VALUE)
    
    ' キーコードチェック
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            If Me.opg_処理選択.VALUE = 3 Then
                Me.cmd_削除.SetFocus
            Else
                Me.cmd_保存.SetFocus
            End If
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
        Case Else
    End Select
Exit_Form_KeyDown:
    Exit Sub

Err_Form_KeyDown:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_KeyDown
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Close イベント
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    Set objRs = Nothing
    Set objDb = Nothing

' 2006/02/07 DEL T.SUZUKI
'' 2005/07/13 ADD T.SUZUKI Start
'    If pintOpenMode = 1 Then
'        Call Forms("FVS600").SetFocus
'    End If
'' 2005/07/13 ADD T.SUZUKI End
' 2006/02/07 DEL T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload イベント
'        MODULE_ID        :Form_Unload
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
    Call ParentReload
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ParentReload
'        機能             :呼び出し元更新処理
'        CREATE_DATE      :2006/02/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub ParentReload()

    Dim objForm         As Form
    Dim intReloadParam  As Integer

    If pblnChangeFlg = True Then
        intReloadParam = 1  ' 変更あり
    Else
        intReloadParam = 0  ' 変更無し
    End If

    If pstrFormID <> "" Then
        Set objForm = Forms(pstrFormID)
        If Not objForm Is Nothing Then
            Call objForm.subReload(intReloadParam)
        End If
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_新顧客コード_AfterUpdate イベント
'        MODULE_ID        :cmb_新顧客コード_AfterUpdate
'        CREATE_DATE      :2006/01/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_新顧客コード_AfterUpdate()

    Dim strSQL      As String
    Dim vCode       As Variant
    Dim intRet      As Integer
    Dim intAns      As Integer
    Dim blnRet      As Boolean

    On Error GoTo cmb_新顧客コード_AfterUpdate_Err

    vCode = Nz(Me.cmb_新顧客コード, "")

    ' 何も入力されていない時は、処理をしない
    If vCode = "" Then Exit Sub

    ' 新顧客コードコンボボックスにて選択した顧客コードのデータを画面に表示
    Call DSP_MODEUPDATE_USER_MAST(1)

    ' 新顧客コードを6桁のゼロサプレス
    Me.cmb_新顧客コード = Format(Me.cmb_新顧客コード, "000000")

cmb_新顧客コード_AfterUpdate_Exit:
    On Error Resume Next
    Me.cmb_新顧客コード.SetFocus
    Exit Sub

cmb_新顧客コード_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo cmb_新顧客コード_AfterUpdate_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_新顧客コード_BeforeUpdate イベント
'        MODULE_ID        :cmb_新顧客コード_BeforeUpdate
'        CREATE_DATE      :2006/01/28
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_新顧客コード_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo cmb_新顧客コード_BeforeUpdate_Err

    If blnFormLoadFlg = False Then Exit Sub

    vCode = Nz(Me.cmb_新顧客コード, "")

    ' 何も入力されていない時は、処理をしない
    If vCode = "" Then
        Me.txt_新顧客名 = "": Exit Sub
    End If

    ' 入力チェック
    Select Case CmBCheck_Char(vCode, 6)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo cmb_新顧客コード_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_新顧客コード_BeforeUpdate_Exit
    End Select

    ' マスタへの存在チェック
    If CHK_USER_CODE_EXIST(1) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_新顧客コード_BeforeUpdate_Exit
    End If

cmb_新顧客コード_BeforeUpdate_Exit:
    Exit Sub

cmb_新顧客コード_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo cmb_新顧客コード_BeforeUpdate_Exit
End Sub

'==============================================================================*
'存在チェック
'顧客ｺｰﾄﾞがＤＢに既に存在するかチェックする。
'True=存在する、False=存在しない
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_USER_CODE_EXIST(ByVal intKbn As Integer) As Boolean             ' 2006/01/28 ADD T.SUZUKI
'Private Function CHK_USER_CODE_EXIST() As Boolean                                   ' 2006/01/28 DEL T.SUZUKI

    Dim stWHERE As String
    Dim vRet    As Variant
    Dim vCode   As Variant

    On Error GoTo err_rtn
    CHK_USER_CODE_EXIST = False

    vCode = IIf(intKbn = 0, Nz(Me.cmb_顧客コード, ""), Nz(Me.cmb_新顧客コード, ""))  ' 2006/01/28 ADD T.SUZUKI
'    vCode = Nz(Me.cmb_顧客コード, "")                                               ' 2006/01/28 DEL T.SUZUKI
    If vCode = "" Then
        GoTo Exit_rtn
    End If

    stWHERE = ""
    stWHERE = stWHERE & " USER_CODE = " & (vCode)
    vRet = Nz(DLookup("USER_CODE", "dbo_USER_MAST", stWHERE), "")

    '()
    Select Case intKbn                                                               ' 2006/01/28 ADD T.SUZUKI
        Case 0  ' *** 顧客コード ***
            Select Case Me.opg_処理選択
                Case 1  ' *** 新規 ***
                        If vRet <> "" Then
                            MsgBox strmsg重複, vbInformation, strMsg_err
                            GoTo Exit_rtn
                        End If
        
                Case 2, 3  ' *** 修正 OR 削除 ***
                    If vRet = "" Then
                        MsgBox strMSGデータなし, vbInformation, strMsg_err
                        GoTo Exit_rtn
                    End If
            End Select

        Case 1  ' *** 新顧客コード ***
            If vRet = "" Then
                MsgBox strMSGデータなし, vbInformation, strMsg_err
                GoTo Exit_rtn
            End If
    End Select                                                                       ' 2006/01/28 ADD T.SUZUKI
    '()

'()    If vRET = "" Then
'()        GoTo EXIT_RTN
'()    End If

    CHK_USER_CODE_EXIST = True

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'入力チェック
'true=正常,false=エラー
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_INPUTDATA() As Boolean

    CHK_INPUTDATA = False

    '更新前の入力チェック

    '顧客ｺｰﾄﾞ
    If Nz(CmNumeric(Me.cmb_顧客コード), "") = "" Then
        MsgBox "顧客ｺｰﾄﾞが正しく入力されていません。", vbInformation, strMsg_kaku
        Me.cmb_顧客コード.SetFocus
        Exit Function
    End If
    '顧客名
    '顧客名ｶﾅ
    '代表者名
    '代表者名ｶﾅ
    '郵便番号
    '住所１
    ' 住所１は36桁とする
    If Nz(Me.txt_住所1, "") <> "" Then
      If CmLenb(Me.txt_住所1) > 36 Then
        MsgBox "都道府県市町村は全角18文字以内で入力して下さい", vbInformation, strMsg_err
        SendKeys "{ESC}"
        Me.txt_住所1.SetFocus
        Exit Function
      End If
    End If
    '住所２
    '電話番号
    '携帯番号
    'FAX番号
    '締日
    If Nz(Me.txt_締日, "") = "" Then
        MsgBox "締日" & strMSG入力不正, vbInformation, strMsg_err
        SendKeys "{ESC}"
        Me.txt_締日.SetFocus
        Exit Function
    End If

    '顧客区分
    If Nz(Me.cmb_顧客区分, "") = "" Then
        MsgBox "顧客区分" & strMSG入力不正, vbInformation, strMsg_err
        Me.txt_顧客区分 = ""
        SendKeys "{ESC}"
        Me.cmb_顧客区分.SetFocus
        Exit Function
    End If

    '集金区分
    If Nz(Me.cmb_集金区分, "") = "" Then
        MsgBox "集金区分" & strMSG入力不正, vbInformation, strMsg_err
        Me.txt_集金区分 = ""
        SendKeys "{ESC}"
        Me.cmb_集金区分.SetFocus
        Exit Function
    End If

    '備考

' 2005/07/13 ADD T.SUZUKI Start
    If Me.chk_入力.VALUE = -1 Then
        ' 連絡先担当者名
        If Nz(Me.txt_連絡先担当者名, "") = "" Then
            MsgBox "連絡先担当者名" & strMSG入力不正, vbInformation, strMsg_err
            Me.txt_連絡先担当者名 = ""
            SendKeys "{ESC}"
            Me.txt_連絡先担当者名.SetFocus
            Exit Function
        End If

        ' 連絡先担当者名カナ
        If Nz(Me.txt_連絡先担当者名カナ, "") = "" Then
            MsgBox "連絡先担当者名カナ" & strMSG入力不正, vbInformation, strMsg_err
            Me.txt_連絡先担当者名カナ = ""
            SendKeys "{ESC}"
            Me.txt_連絡先担当者名カナ.SetFocus
            Exit Function
        End If

        ' 連絡先郵便番号
        If Nz(Me.txt_連絡先郵便番号, "") = "" Then
            MsgBox "連絡先郵便番号" & strMSG入力不正, vbInformation, strMsg_err
            Me.txt_連絡先郵便番号 = ""
            SendKeys "{ESC}"
            Me.txt_連絡先郵便番号.SetFocus
            Exit Function
        End If

        ' 連絡先住所１
        If Nz(Me.txt_連絡先都道府県市町村, "") = "" Then
            MsgBox "連絡先都道府県市町村" & strMSG入力不正, vbInformation, strMsg_err
            Me.txt_連絡先都道府県市町村 = ""
            SendKeys "{ESC}"
            Me.txt_連絡先都道府県市町村.SetFocus
            Exit Function
        End If

        ' 連絡先住所２
        If Nz(Me.txt_連絡先番地, "") = "" Then
            MsgBox "連絡先番地" & strMSG入力不正, vbInformation, strMsg_err
            Me.txt_連絡先番地 = ""
            SendKeys "{ESC}"
            Me.txt_連絡先番地.SetFocus
            Exit Function
        End If

        ' 連絡先電話番号
        If Nz(Me.txt_連絡先電話番号, "") = "" Then
            MsgBox "連絡先電話番号" & strMSG入力不正, vbInformation, strMsg_err
            Me.txt_連絡先電話番号 = ""
            SendKeys "{ESC}"
            Me.txt_連絡先電話番号.SetFocus
            Exit Function
        End If
        
        ' 2006/05/17 Add Tajima Start
        ' 連絡先担当者名は18桁とする
        If Nz(Me.txt_連絡先担当者名, "") <> "" Then
'          If CmLenb(Me.txt_連絡先担当者名) > 18 Then                                           'DELETE 2011/11/29 M.RYU
'            MsgBox "連絡先担当者名は全角９文字以内として下さい", vbInformation, strMsg_err     'DELETE 2011/11/29 M.RYU
          If CmLenb(Me.txt_連絡先担当者名) > 36 Then                                            'INSERT 2011/11/29 M.RYU
            MsgBox "連絡先担当者名は全角18文字以内として下さい", vbInformation, strMsg_err      'INSERT 2011/11/29 M.RYU
            SendKeys "{ESC}"
            Me.txt_連絡先担当者名.SetFocus
            Exit Function
          End If
        End If
        ' 連絡先住所１は36桁とする
        If Nz(Me.txt_連絡先都道府県市町村, "") <> "" Then
          If CmLenb(Me.txt_連絡先都道府県市町村) > 36 Then
            MsgBox "都道府県市町村は全角18文字以内で入力して下さい", vbInformation, strMsg_err
            SendKeys "{ESC}"
            Me.txt_連絡先都道府県市町村.SetFocus
            Exit Function
          End If
        End If
        ' 2006/05/17 Add Tajima End
        
    End If
' 2005/07/13 ADD T.SUZUKI End

    '▼2007/06/10 add tajima
    If Nz(Me.cmb_顧客情報変更理由, "00") <> "00" Then
        If Nz(Me.txt_顧客情報変更日, "") = "" Then
            MsgBox "顧客情報を変更したならば変更日を入力して下さい", vbInformation, strMsg_err
            SendKeys "{ESC}"
            Me.txt_顧客情報変更日.SetFocus
            Exit Function
        End If
    End If
    '▲2007/06/10 add tajima
                
    '↓INSERT 2010/04/19 SHIBAZAKI
    If Nz(Me.txt_USER_KKANA, "") <> "" Then
      If CmLenb(Me.txt_USER_KKANA) > 40 Then
        MsgBox "勤務先カナは40文字以内で入力して下さい", vbInformation, strMsg_err
        SendKeys "{ESC}"
        Me.txt_USER_KKANA.SetFocus
        Exit Function
      End If
    End If
    If Nz(Me.txt_USER_KADR_1, "") <> "" Then
      If CmLenb(Me.txt_USER_KADR_1) > 36 Then
        MsgBox "緊急連絡先都道府県市町村は全角18文字以内で入力して下さい", vbInformation, strMsg_err
        SendKeys "{ESC}"
        Me.txt_USER_KADR_1.SetFocus
        Exit Function
      End If
    End If
    '↑INSERT 2010/04/19 SHIBAZAKI

    CHK_INPUTDATA = True
End Function

'==============================================================================*
'更新時
'顧客コードで検索されたレコードを画面にセットする
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DSP_MODEUPDATE_USER_MAST(ByVal iMode As Integer) As Boolean

  '取引先マスタへの問合せ
  '引数 ０：存在チェック
  '     １：データ参照

    Dim strSQL As String
    Dim vCode  As Variant

    Dim stid   As String

    On Error GoTo err_rtn

    DSP_MODEUPDATE_USER_MAST = False

' 2006/01/28 REP T.SUZUKI Start
    ' intKbn = 0：顧客コード／intKbn = 1：新顧客コード
    vCode = IIf(iMode = 0, Nz(Me.cmb_顧客コード, ""), Nz(Me.cmb_新顧客コード, ""))
'    vCode = Nz(Me.cmb_顧客コード, "")
' 2006/01/28 REP T.SUZUKI End

    If vCode = "" Then
        DSP_MODEUPDATE_USER_MAST = True
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    Select Case iMode
        Case 0  ' *** 顧客コード ***
            ' 顧客マスタよりデータ取得
            strSQL = ""
            strSQL = "SELECT"
            strSQL = strSQL & " dbo_USER_MAST.USER_CODE,"       '0
            strSQL = strSQL & " dbo_USER_MAST.USER_NAME,"       '1
            strSQL = strSQL & " dbo_USER_MAST.USER_KANA,"       '2
            strSQL = strSQL & " dbo_USER_MAST.USER_TANM,"       '3
            strSQL = strSQL & " dbo_USER_MAST.USER_TAKA,"       '4
            strSQL = strSQL & " dbo_USER_MAST.USER_YUBINO,"     '5
            strSQL = strSQL & " dbo_USER_MAST.USER_ADR_1,"      '6
            strSQL = strSQL & " dbo_USER_MAST.USER_ADR_2,"      '7
            strSQL = strSQL & " dbo_USER_MAST.USER_ADR_3,"      '8
            strSQL = strSQL & " dbo_USER_MAST.USER_TEL,"        '9
            strSQL = strSQL & " dbo_USER_MAST.USER_KEITAI,"     '10
            strSQL = strSQL & " dbo_USER_MAST.USER_FAX,"        '11
            strSQL = strSQL & " dbo_USER_MAST.USER_MAIL,"       '12
            strSQL = strSQL & " dbo_USER_MAST.USER_SDAY,"       '13
            strSQL = strSQL & " dbo_USER_MAST.USER_KKBN,"       '14
            strSQL = strSQL & " dbo_USER_MAST.USER_SKBN,"       '15
            strSQL = strSQL & " dbo_USER_MAST.USER_BIKO,"       '16
' 2005/07/13 ADD C.IGUCHI Start
            strSQL = strSQL & " dbo_USER_MAST.USER_RNAME,"      '17
            strSQL = strSQL & " dbo_USER_MAST.USER_RKANA,"      '18
            strSQL = strSQL & " dbo_USER_MAST.USER_RPOST,"      '19
            strSQL = strSQL & " dbo_USER_MAST.USER_RADR_1,"     '20
            strSQL = strSQL & " dbo_USER_MAST.USER_RADR_2,"     '21
            strSQL = strSQL & " dbo_USER_MAST.USER_RADR_3,"     '22
            strSQL = strSQL & " dbo_USER_MAST.USER_RTEL,"       '23
            strSQL = strSQL & " dbo_USER_MAST.USER_RKEITAI ,"   '24
            strSQL = strSQL & " dbo_USER_MAST.USER_RFAX, "      '25
            strSQL = strSQL & " dbo_USER_MAST.USER_NEW_CODE "   '26          ' 2006/01/28 ADD T.SUZUKI
' 2005/07/13 ADD C.IGUCHI End
            strSQL = strSQL & " ,dbo_USER_MAST.USER_CHG_RKBN "  '27 '2007/06/10 add tajima
            strSQL = strSQL & " ,dbo_USER_MAST.USER_CHG_DATE "  '28 '2007/06/10 add tajima
            strSQL = strSQL & " ,dbo_USER_MAST.USER_SEIBETU "   '29 '2014/11/15 INS M.HONDA
            
            '↓INSERT 2010/04/19 SHIBAZAKI
            strSQL = strSQL & " ,dbo_USER_MAST.USER_RMAIL "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KNAME "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KKANA "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_JOB "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KPOST "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KADR_1 "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KADR_2 "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KADR_3 "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KTEL "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KHOUR_START "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_KHOUR_END "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_NENSHU "
            strSQL = strSQL & " ,dbo_USER_MAST.USER_BIRTHDAY "
            '↑INSERT 2010/04/19 SHIBAZAKI
            
            strSQL = strSQL & " FROM dbo_USER_MAST"
            strSQL = strSQL & " WHERE dbo_USER_MAST.USER_CODE = " & (vCode)
            Set objRs = objDb.OpenRecordset(strSQL, dbOpenForwardOnly)

            If Not (objRs.EOF) Then                                          ' 2006/01/28 ADD T.SUZUKI
'            If Not (objRs.BOF And objRs.EOF) Then                           ' 2006/01/28 DEL T.SUZUKI
                '画面に値をセット
                Me.cmb_顧客コード = Format(Nz(objRs(0), ""), "000000")       ' 2006/01/28 ADD T.SUZUKI
'                Me.cmb_顧客コード = Nz(objRs(0), "")                        ' 2006/01/28 DEL T.SUZUKI
                Me.txt_顧客名 = Nz(objRs(1), "")
                Me.txt_顧客名カナ = Nz(objRs(2), "")
                Me.txt_代表者名 = Nz(objRs(3), "")
                Me.txt_代表者名カナ = Nz(objRs(4), "")
                Me.txt_郵便番号 = Nz(objRs(5), "")
                Me.txt_住所1 = Nz(objRs(6), "")
                Me.txt_住所2 = Nz(objRs(7), "")
                Me.txt_住所3 = Nz(objRs(8), "")    '(2005/07/13 iguchi ADD) <<<
                Me.txt_電話番号 = Nz(objRs(9), "")
                Me.txt_携帯番号 = Nz(objRs(10), "")
                Me.txt_FAX番号 = Nz(objRs(11), "")
                Me.txt_メールアドレス = Nz(objRs(12), "")
                Me.txt_締日 = Nz(objRs(13), "")
                Me.cmb_顧客区分 = Format(Nz(objRs(14), ""), "00")            ' 2006/01/28 ADD T.SUZUKI
                Me.cmb_集金区分 = Format(Nz(objRs(15), ""), "00")            ' 2006/01/28 ADD T.SUZUKI
                strKkbn = Format(Nz(objRs(15), ""), "00")                    ' 2013/09/19 M.HONDA INS
'                Me.cmb_顧客区分 = Nz(objRs(14), "")                         ' 2006/01/28 DEL T.SUZUKI
'                Me.cmb_集金区分 = Nz(objRs(15), "")                         ' 2006/01/28 DEL T.SUZUKI
                Me.txt_備考 = Nz(objRs(16), "")
                stid = "020"
                Me.txt_顧客区分 = s_SelectRyak(Me.cmb_顧客区分, stid)
                stid = "030"
                Me.txt_集金区分 = s_SelectRyak(Me.cmb_集金区分, stid)

                '2014/11/15 M.HONDA INS
                Me.cmb_性別区分 = Format(Nz(objRs(29), ""), "00")
                stid = "302"
                Me.txt_性別区分 = s_SelectRyak(Me.cmb_性別区分, stid)
                '2014/11/15 M.HONDA INS

' 2005/07/13 ADD C.IGUCHI Start
                If Nz(Trim(objRs(17)), "") <> "" Then
                    ' 顧客マスタ.連絡先担当者名がNullでなければ、入力するチェックボックスを
                    ' 選択状態に設定する
                    Me.chk_入力 = -1
                End If

                Me.txt_連絡先担当者名 = Nz(objRs(17), "")
                Me.txt_連絡先担当者名カナ = Nz(objRs(18), "")
                Me.txt_連絡先郵便番号 = Nz(objRs(19), "")
                Me.txt_連絡先都道府県市町村 = Nz(objRs(20), "")
                Me.txt_連絡先番地 = Nz(objRs(21), "")
                Me.txt_連絡先建物名等 = Nz(objRs(22), "")
                Me.txt_連絡先電話番号 = Nz(objRs(23), "")
                Me.txt_連絡先携帯番号 = Nz(objRs(24), "")
                Me.txt_連絡先FAX番号 = Nz(objRs(25), "")
' 2005/07/13 ADD C.IGUCHI End
                Me.cmb_新顧客コード = Format(Nz(objRs(26), ""), "000000")    ' 2006/01/28 ADD T.SUZUKI
                Me.txt_新顧客名 = ""                                         ' 2006/01/28 ADD T.SUZUKI
            
                Me.cmb_顧客情報変更理由 = Format(Nz(objRs(27), 0), Me.cmb_顧客情報変更理由.Format) '2007/06/10 add tajima
                Me.txt_顧客情報変更理由 = Me.cmb_顧客情報変更理由.Column(1, Me.cmb_顧客情報変更理由.ListIndex)
                Me.txt_顧客情報変更日 = objRs(28)   '2007/06/10 add tajima
            
                '↓INSERT 2010/04/19 SHIBAZAKI
                With objRs
                    txt_USER_KNAME = .Fields("USER_KNAME")
'                    txt_USER_KKANA = .Fields("USER_KKANA")                     'DELETE 2013/06/11 K.ISHIZAKA
                    txt_USER_JOB = .Fields("USER_JOB")
                    txt_USER_KPOST = .Fields("USER_KPOST")
                    txt_USER_KADR_1 = .Fields("USER_KADR_1")
                    txt_USER_KADR_2 = .Fields("USER_KADR_2")
                    txt_USER_KADR_3 = .Fields("USER_KADR_3")
                    txt_USER_KTEL = .Fields("USER_KTEL")
                    txt_USER_KHOUR_START = .Fields("USER_KHOUR_START")
                    txt_USER_KHOUR_END = .Fields("USER_KHOUR_END")
                    txt_USER_NENSHU = .Fields("USER_NENSHU")
                    If Trim(Nz(.Fields("USER_BIRTHDAY"), "")) = "" Then
                        txt_USER_BIRTHDAY = .Fields("USER_BIRTHDAY")
                    Else
                        txt_USER_BIRTHDAY = Left(.Fields("USER_BIRTHDAY"), 4) & "/" & Mid(.Fields("USER_BIRTHDAY"), 5, 2) & "/" & Right(.Fields("USER_BIRTHDAY"), 2)
                    End If
                End With
                '↑INSERT 2010/04/19 SHIBAZAKI
            End If

        Case 1  ' *** 新顧客コード ***
            strSQL = ""
            strSQL = "SELECT dbo_USER_MAST.USER_NAME "
            strSQL = strSQL & "  FROM dbo_USER_MAST "
            strSQL = strSQL & " WHERE dbo_USER_MAST.USER_CODE = " & (vCode)
            Set objRs = objDb.OpenRecordset(strSQL, dbOpenForwardOnly)

            If Not (objRs.EOF) Then
                Me.txt_新顧客名 = Nz(objRs(0), "")
            End If
    End Select

    DSP_MODEUPDATE_USER_MAST = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    objRs.Close
    Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'顧客区分、集金区分の値を取得する。
'パラメータ ContNM Object
'戻り値 strRyak String [名称マスタ].略称
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function s_SelectRyak(ByVal ContNM As Object, stid As String) As String

    Dim strRyak     As String
    Dim strWhere    As String
    Dim vCode       As Variant
    Dim objControl  As Object
    Dim strNameID   As String
    Dim strID        As String
    
    s_SelectRyak = ""
    Set objControl = ContNM

    vCode = Nz(objControl, "")
    If vCode = "" Then
        Exit Function
    End If

    strID = stid
    
 'Select Case ObjControl
 '       Case Me.cmb_顧客区分
 '           strNameID = "020"   '顧客区分 NAME_ID="020" 固定
 '       Case Me.cmb_集金区分
 '           strNameID = "030"   '集金区分 NAME_ID="030" 固定
 '       Case Else
 '           s_SelectRyak = ""
 '           Exit Function
 '   End Select

    strWhere = ""
  '  strWhere = strWhere & " NAME_ID = '" & strNameID & "'"
    strWhere = strWhere & " NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & (vCode)

    strRyak = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWhere), "")

    s_SelectRyak = strRyak
End Function

'==============================================================================*
'chk_入力_Clickイベント
'==============================================================================*
Private Sub chk_入力_Click()

    Dim lngBackColor As Long
    Dim intControl   As Integer

    ' 【その他連絡先】の使用可否、背景色設定
    Call subTxtControl

    ' 入力するチェックボックスが未チェックの場合
    If Me.chk_入力.VALUE = 0 Then
        ' 【その他連絡先】を初期化
        Call CLR_RTN2
    End If
End Sub

' 2005/07/13 ADD T.SUZUKI Start
'==============================================================================*
' chk_入力_KeyDownイベント
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_入力_KeyDown(KeyCode As Integer, Shift As Integer)

    If Me.txt_連絡先担当者名.Enabled = True Then
        Exit Sub
    End If

    If Shift = 0 And KeyCode = vbKeyTab Then
        Select Case intModeflg
            Case intModeInsert, intModeUpdate
                Me.cmd_保存.SetFocus
            Case intModeDelete
                Me.cmd_削除.SetFocus
        End Select
        KeyCode = 0
    End If
End Sub
' 2005/07/13 ADD T.SUZUKI End

'==============================================================================*
'【その他連絡先】の使用可否、背景色の設定を行う
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subTxtControl()

    Dim lngBackColor As Long
    Dim intControl   As Integer

    If (chk_入力.VALUE = -1 And Me.opg_処理選択.VALUE <> 3) Then
        'チェックされてるとき(白)
        lngBackColor = 16777215
        intControl = 0

    Else
        'チェックされていないとき(黄色)
        lngBackColor = 14680063
        intControl = 1
    End If

    With Me
        ' 【その他連絡先】の背景色設定
        .txt_連絡先担当者名.BackColor = lngBackColor
        .txt_連絡先担当者名カナ.BackColor = lngBackColor
        .txt_連絡先郵便番号.BackColor = lngBackColor
        .txt_連絡先都道府県市町村.BackColor = lngBackColor
        .txt_連絡先番地.BackColor = lngBackColor
        .txt_連絡先建物名等.BackColor = lngBackColor
        .txt_連絡先電話番号.BackColor = lngBackColor
        .txt_連絡先携帯番号.BackColor = lngBackColor
        .txt_連絡先FAX番号.BackColor = lngBackColor

        ' 【その他連絡先】の使用可否設定
        Call CmLockControl(.txt_連絡先担当者名, intControl)
        Call CmLockControl(.txt_連絡先担当者名カナ, intControl)
        Call CmLockControl(.txt_連絡先郵便番号, intControl)

        If intControl = 0 Then
            Me.cmd_RYubin.Enabled = True
        Else
            Me.cmd_RYubin.Enabled = False
        End If

        Call CmLockControl(.txt_連絡先都道府県市町村, intControl)
        Call CmLockControl(.txt_連絡先番地, intControl)
        Call CmLockControl(.txt_連絡先建物名等, intControl)
        Call CmLockControl(.txt_連絡先電話番号, intControl)
        Call CmLockControl(.txt_連絡先携帯番号, intControl)
        Call CmLockControl(.txt_連絡先FAX番号, intControl)
    End With
    
    If Nz(Me.cmb_顧客情報変更理由.VALUE, "00") = "00" Then
        Me.txt_顧客情報変更日.BackColor = 14680063
        Call CmLockControl(Me.txt_顧客情報変更日, 1)
    Else
        Me.txt_顧客情報変更日.BackColor = 16777215
        Call CmLockControl(Me.txt_顧客情報変更日, 0)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_集金区分_KeyPress イベント
'        MODULE_ID        :cmb_集金区分_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_集金区分, KeyAscii, 2)      '桁数チェック追加 01-06-2000 Araki
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_削除_KeyDown イベント
'        MODULE_ID        :cmd_削除_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            cmd_削除_Click
        Case vbKeyTab
            If Shift > 0 Then
                Me.cmb_顧客コード.SetFocus
            Else
                Me.cmd_取消.SetFocus
            End If
        Case vbKeyUp, vbKeyRight
            Me.cmd_取消.SetFocus
        Case vbKeyDown, vbKeyLeft
            If intModeflg <> intModeDelete Then
                Me.cmd_保存.SetFocus
            Else
                Me.cmd_終了.SetFocus
            End If
    End Select
    KeyCode = 0

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_取消_KeyDown イベント
'        MODULE_ID        :cmd_取消_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmd_取消_Click
        Case vbKeyTab
            If Shift > 0 Then
                Me.cmb_顧客コード.SetFocus
            Else
                Me.cmd_終了.SetFocus
            End If
        Case vbKeyUp, vbKeyRight
            Me.cmd_終了.SetFocus
        Case vbKeyDown, vbKeyLeft
            If intModeflg <> intModeDelete Then
                Me.cmd_保存.SetFocus
            Else
                Me.cmd_削除.SetFocus
            End If
    End Select
    KeyCode = 0
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_KeyDown イベント
'        MODULE_ID        :cmd_終了_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmd_終了_Click
        Case vbKeyTab
            If Shift > 0 Then
                Me.cmb_顧客コード.SetFocus
            Else
                If intModeflg <> intModeDelete Then
                    Me.cmd_保存.SetFocus
                Else
                    Me.cmd_削除.SetFocus
                End If
            End If
        Case vbKeyLeft, vbKeyDown
            Me.cmd_取消.SetFocus
        Case vbKeyRight, vbKeyUp
            If intModeflg <> intModeDelete Then
                Me.cmd_保存.SetFocus
            Else
                Me.cmd_削除.SetFocus
            End If
    End Select
    KeyCode = 0
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_MouseDown イベント
'        MODULE_ID        :cmd_終了_MouseDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_KeyDown イベント
'        MODULE_ID        :cmd_保存_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            cmd_保存_Click
        Case vbKeyTab
            If Shift > 0 Then
                Me.cmb_顧客コード.SetFocus
            Else
                If intModeflg <> intModeDelete Then
                    Me.cmd_取消.SetFocus
                Else
                    Me.cmd_削除.SetFocus
                End If
            End If
        Case vbKeyUp, vbKeyRight
            If intModeflg <> intModeDelete Then
                Me.cmd_取消.SetFocus
            Else
                Me.cmd_削除.SetFocus
            End If
        Case vbKeyLeft, vbKeyDown
            Me.cmd_終了.SetFocus
    End Select
    KeyCode = 0
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opt_削除_KeyDown イベント
'        MODULE_ID        :opt_削除_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opt_削除_KeyDown(KeyCode As Integer, Shift As Integer)
    '$$$ start $$$ 変更有無に関わらず処理選択が無条件に変更されてしまうため削除 01-11-2000 Araki
    'Select Case KeyCode
    '    Case vbKeyReturn
    '        Me.cmb_顧客コード.SetFocus
    '    Case vbKeyUp, vbKeyLeft
    '        Me.opg_処理選択.Value = 2
    '    Case vbKeyDown, vbKeyRight
    '        Me.opg_処理選択.Value = 1
    '    Case vbKeyTab
    '        If Shift <> 0 Then
    '            Me.opg_処理選択.Value = 2
    '        Else
    '            Me.opg_処理選択.Value = 1
    '        End If
    '    Case Else
    'End Select
    '$$$ end $$$
    
    '$$$ start $$$ クリア=Noの時に処理選択が変更されてしまうので削除 01-11-2000 Araki
    'KeyCode = 0
    'intModeflg = Me.opg_処理選択.Value
    'Call s_bottunset(intModeflg)
    '$$$ end $$$
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opt_修正_KeyDown イベント
'        MODULE_ID        :opt_修正_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opt_修正_KeyDown(KeyCode As Integer, Shift As Integer)
    '$$$ start $$$ 変更有無に関わらず処理選択が無条件に変更されてしまうため削除 01-11-2000 Araki
    'Select Case KeyCode
    '    Case vbKeyReturn
    '        Me.cmb_顧客コード.SetFocus
    '    Case vbKeyDown, vbKeyRight
    '        Me.opg_処理選択.Value = 3
    '    Case vbKeyUp, vbKeyLeft
    '        Me.opg_処理選択.Value = 1
    '    Case vbKeyTab
    '        If Shift <> 0 Then
    '            Me.opg_処理選択.Value = 1
    '        Else
    '            Me.opg_処理選択.Value = 3
    '        End If
    '    Case Else
    'End Select
    '$$$ end $$$
    
    '$$$ start $$$ クリア=Noの時に処理選択が変更されてしまうので削除 01-11-2000 Araki
    'KeyCode = 0
    'intModeflg = Me.opg_処理選択.Value
    'Call s_bottunset(intModeflg)
    '$$$ end $$$
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opt_新規_KeyDown イベント
'        MODULE_ID        :opt_新規_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opt_新規_KeyDown(KeyCode As Integer, Shift As Integer)
    '$$$ start $$$ 変更有無に関わらず処理選択が無条件に変更されてしまうため削除 01-11-2000 Araki
    'Select Case KeyCode
    '    Case vbKeyReturn
    '        Me.cmb_顧客コード.SetFocus
    '    Case vbKeyDown, vbKeyRight
    '        Me.opg_処理選択.Value = 2
    '    Case vbKeyLeft, vbKeyUp
    '        Me.opg_処理選択.Value = 3
    '    Case vbKeyTab
    '        If Shift <> 0 Then
    '            Me.opg_処理選択.Value = 3
    '        Else
    '            Me.opg_処理選択.Value = 2
    '        End If
    '    Case Else
    'End Select
    '$$$ end $$$
    
    '$$$ start $$$ クリア=Noの時に処理選択が変更されてしまうので削除 01-11-2000 Araki
    'KeyCode = 0
    'intModeflg = Me.opg_処理選択.Value
    'Call s_bottunset(intModeflg)
    '$$$ end $$$
End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面初期化
'        MODULE_ID        :CLR_RTN
'        CREATE_DATE      :N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CLR_RTN()

    On Error Resume Next  'エラーがあっても処理続行
  
    '画面上のテキストボックスをクリアする
    With Me
        .cmb_顧客コード = ""
        .txt_顧客名 = ""
        .txt_顧客名カナ = ""
        .txt_代表者名 = ""
        .txt_代表者名カナ = ""
        .txt_郵便番号 = ""
        .txt_住所1 = ""
        .txt_住所2 = ""
        .txt_住所3 = ""  '(2005/07/13 iguchi ADD) <<<
        .txt_電話番号 = ""
        .txt_携帯番号 = ""
        .txt_FAX番号 = ""
        .txt_メールアドレス = ""
        .txt_締日 = "31"
        .cmb_顧客区分 = "01"
        Call cmb_顧客区分_AfterUpdate
        
        '' 2013/09/19 M.HONDA START
        '' 初期表示をコンビニから振込へ変更
        ''.cmb_集金区分 = "07"
        .cmb_集金区分 = "01"
        strKkbn = "01"
        '' 2013/09/19 M.HONDA END
        
        cmb_集金区分_AfterUpdate
        .txt_備考 = ""

' 2005/07/06 ADD C.IGUCHI Start
        .chk_入力.VALUE = 0  ' 未選択状態
'        .chk_入力 = ""
        .txt_連絡先担当者名 = ""
        .txt_連絡先担当者名カナ = ""
        .txt_連絡先郵便番号 = ""
        .txt_連絡先都道府県市町村 = ""
        .txt_連絡先番地 = ""
        .txt_連絡先建物名等 = ""
        .txt_連絡先電話番号 = ""
        .txt_連絡先携帯番号 = ""
        .txt_連絡先FAX番号 = ""
' 2005/07/13 ADD C.IGUCHI End
        .cmb_新顧客コード = ""                                                ' 2006/01/28 ADD T.SUZUKI
        .txt_新顧客名 = ""                                                    ' 2006/01/28 ADD T.SUZUKI
        .cmb_顧客情報変更理由 = Null  ' 2007/06/10 ADD tajima
        .txt_顧客情報変更理由 = ""    ' 2007/06/10 ADD tajima
        .txt_顧客情報変更日 = Null    ' 2007/06/10 ADD tajima
        
        '↓INSERT 2010/04/19 SHIBAZAKI
        txt_USER_KNAME = Null
        txt_USER_KKANA = Null
        txt_USER_JOB = Null
        txt_USER_KPOST = Null
        txt_USER_KADR_1 = Null
        txt_USER_KADR_2 = Null
        txt_USER_KADR_3 = Null
        txt_USER_KTEL = Null
        txt_USER_KHOUR_START = Null
        txt_USER_KHOUR_END = Null
        txt_USER_NENSHU = Null
        txt_USER_BIRTHDAY = Null
        '↑INSERT 2010/04/19 SHIBAZAKI
   
    End With

    '変更フラグの初期化
    blnDataUpdate = False
    is顧客情報変更 = False '2007/06/10 add tajima

    '退避変数の初期化
    vBack_Key = Null
End Sub

' 2005/07/13 ADD C.IGUCHI Start
'==============================================================================*
'
'        MODULE_NAME      :画面初期化(その他連絡先)
'        MODULE_ID        :CLR_RTN2
'        CREATE_DATE      :C.IGUCHI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CLR_RTN2()
    
    With Me
        .txt_連絡先担当者名 = ""
        .txt_連絡先担当者名カナ = ""
        .txt_連絡先郵便番号 = ""
        .txt_連絡先都道府県市町村 = ""
        .txt_連絡先番地 = ""
        .txt_連絡先建物名等 = ""
        .txt_連絡先電話番号 = ""
        .txt_連絡先携帯番号 = ""
        .txt_連絡先FAX番号 = ""
    End With
End Sub
' 2005/07/13 ADD C.IGUCHI End

'==============================================================================*
'文字項目のチェックを一括でする。
'==============================================================================*
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

  '返り値は、
  '
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
  
On Error GoTo err_rtn
  CmBCheck_Char = 9
  
  If CmLenb(oText) > ilen Then
    CmBCheck_Char = 2
    GoTo Exit_rtn
  End If
  
  CmBCheck_Char = 0

Exit_rtn:
  Exit Function

err_rtn:
  MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
  CmBCheck_Char = 9
  GoTo Exit_rtn
End Function

'==============================================================================*
'数値項目のチェックを一括でする。
'==============================================================================*
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, Optional ByVal isLEN As Variant) As Integer
  '返り値は、
  '１：数値と判断できない時
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
    
  '引数 oText ←判定する値
  '     iLen  ←整数桁
  '     isLen ←小数桁(省略された時は、小数点桁数無し)
    
  Dim vINT    As Variant
  Dim vsINT   As Variant
  Dim i       As Integer
    
  Dim dTEXT   As Double
  Dim dFROM   As Double
  Dim dTO     As Double
  Dim vRet    As Variant
    
On Error GoTo err_rtn
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
  For i = 1 To ilen
    vINT = vINT & "9"
  Next
  
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
'==============================================================================*
Private Sub CmFormatConma(ByVal oText As TextBox, ByVal iMode As Integer)
  '引数 ０：カンマ表示
  '     １：カンマ無し
  
  Dim objCONT As TextBox
  
  Set objCONT = oText
  
  If iMode = 0 Then
    objCONT.VALUE = Format(Nz(CmNumeric(objCONT), ""), "#,##0")
  Else
    objCONT.VALUE = Format(Nz(CmNumeric(objCONT), ""), "0")
  End If
End Sub

'==============================================================================*
'コントロールの入力切替設定
'==============================================================================*
Private Sub CmLockControl(ByVal ContNM As Object, ByVal iMode As Integer)

  '引数 ０  使用可
  '     １  使用不可
  Dim objControl As Object
  
  Set objControl = ContNM
  
  If iMode = 0 Then
    objControl.Enabled = True
    objControl.Locked = False
  Else
    objControl.Enabled = False
    objControl.Locked = True
  End If
End Sub

'==============================================================================*
'入力区分によるコントロールのロックを設定
'==============================================================================*
Private Sub Etc_Lock_Rtn()

'パラメータ　０  使用可
'           １  使用不可

    Select Case intModeflg
        Case intModeInsert, intModeUpdate      '新規、更新
            Call CmLockControl(Me.cmb_顧客コード, 0)
            Call CmLockControl(Me.txt_顧客名, 0)
            Call CmLockControl(Me.txt_顧客名カナ, 0)
            Call CmLockControl(Me.txt_代表者名, 0)
            Call CmLockControl(Me.txt_代表者名カナ, 0)
            Call CmLockControl(Me.txt_郵便番号, 0)
            Me.cmd_Yubin.Enabled = True
            Call CmLockControl(Me.txt_住所1, 0)
            Call CmLockControl(Me.txt_住所2, 0)
            Call CmLockControl(Me.txt_住所3, 0)
            Call CmLockControl(Me.txt_電話番号, 0)
            Call CmLockControl(Me.txt_携帯番号, 0)
            Call CmLockControl(Me.txt_FAX番号, 0)
            Call CmLockControl(Me.txt_メールアドレス, 0)
            Call CmLockControl(Me.txt_締日, 0)
            Call CmLockControl(Me.cmb_顧客区分, 0)
            Call CmLockControl(Me.txt_顧客区分, 1)
            Call CmLockControl(Me.cmb_集金区分, 0)
            Call CmLockControl(Me.txt_集金区分, 1)
            Call CmLockControl(Me.txt_備考, 0)

' 2005/07/13 ADD C.IGUCHI Start
            Me.chk_入力.Enabled = True
            Call CmLockControl(Me.txt_連絡先担当者名, 0)
            Call CmLockControl(Me.txt_連絡先担当者名カナ, 0)
            Call CmLockControl(Me.txt_連絡先郵便番号, 0)
            Me.cmd_RYubin.Enabled = True
            Call CmLockControl(Me.txt_連絡先都道府県市町村, 0)
            Call CmLockControl(Me.txt_連絡先番地, 0)
            Call CmLockControl(Me.txt_連絡先建物名等, 0)
            Call CmLockControl(Me.txt_連絡先電話番号, 0)
            Call CmLockControl(Me.txt_連絡先携帯番号, 0)
            Call CmLockControl(Me.txt_連絡先FAX番号, 0)
' 2005/07/13 ADD C.IGUCHI End

            Call CmLockControl(Me.cmb_新顧客コード, 0)                             ' 2006/01/28 ADD T.SUZUKI
            Call CmLockControl(Me.txt_新顧客名, 1)                                 ' 2006/01/28 ADD T.SUZUKI
            '▼ 2007/06/10 add tajima
            If intModeflg = intModeInsert Then
                '新規モードでは選択不可！
                Me.cmb_顧客情報変更理由.Enabled = False
                Me.cmb_顧客情報変更理由.Locked = False
            Else
                Call CmLockControl(Me.cmb_顧客情報変更理由, 0)
            End If
            Call CmLockControl(Me.txt_顧客情報変更日, 1)
            '▲ 2007/06/10 add tajima
            
            Call CmLockControl(Me.txt_USER_BIRTHDAY, 0)                             'INSERT 2010/04/19 SHIBAZAKI
            Call LockKinmusaki(0)                                                   'INSERT 2010/04/19 SHIBAZAKI
            
            Me.cmd_保存.Enabled = True
            Me.cmd_削除.Enabled = False

        Case intModeDelete      '削除 顧客コードのみ入力可
            Call CmLockControl(Me.cmb_顧客コード, 0)
            Call CmLockControl(Me.txt_顧客名, 1)
            Call CmLockControl(Me.txt_顧客名カナ, 1)
            Call CmLockControl(Me.txt_代表者名, 1)
            Call CmLockControl(Me.txt_代表者名カナ, 1)
            Call CmLockControl(Me.txt_郵便番号, 1)
            Me.cmd_Yubin.Enabled = False
            Call CmLockControl(Me.txt_住所1, 1)
            Call CmLockControl(Me.txt_住所2, 1)
            Call CmLockControl(Me.txt_住所3, 1)
            Call CmLockControl(Me.txt_電話番号, 1)
            Call CmLockControl(Me.txt_携帯番号, 1)
            Call CmLockControl(Me.txt_FAX番号, 1)
            Call CmLockControl(Me.txt_メールアドレス, 1)
            Call CmLockControl(Me.txt_締日, 1)
            Call CmLockControl(Me.cmb_顧客区分, 1)
            Call CmLockControl(Me.txt_顧客区分, 1)
            Call CmLockControl(Me.cmb_集金区分, 1)
            Call CmLockControl(Me.txt_集金区分, 1)
            Call CmLockControl(Me.txt_備考, 1)

' 2005/07/13 ADD C.IGUCHI Start
            Me.chk_入力.Enabled = False
            Call CmLockControl(Me.txt_連絡先担当者名, 1)
            Call CmLockControl(Me.txt_連絡先担当者名カナ, 1)
            Call CmLockControl(Me.txt_連絡先郵便番号, 1)
            Me.cmd_RYubin.Enabled = False
            Call CmLockControl(Me.txt_連絡先都道府県市町村, 1)
            Call CmLockControl(Me.txt_連絡先番地, 1)
            Call CmLockControl(Me.txt_連絡先建物名等, 1)
            Call CmLockControl(Me.txt_連絡先電話番号, 1)
            Call CmLockControl(Me.txt_連絡先携帯番号, 1)
            Call CmLockControl(Me.txt_連絡先FAX番号, 1)
' 2005/07/13 ADD C.IGUCHI End

            Call CmLockControl(Me.cmb_新顧客コード, 1)                             ' 2006/01/28 ADD T.SUZUKI
            Call CmLockControl(Me.txt_新顧客名, 1)                                 ' 2006/01/28 ADD T.SUZUKI

            Call CmLockControl(Me.cmb_顧客情報変更理由, 1)    '2007/06/10 add tajima
            Call CmLockControl(Me.txt_顧客情報変更日, 1)      '2007/06/10 add tajima

            Call CmLockControl(Me.txt_USER_BIRTHDAY, 1)                             'INSERT 2010/04/19 SHIBAZAKI
            Call LockKinmusaki(1)                                                   'INSERT 2010/04/19 SHIBAZAKI

            Me.cmd_保存.Enabled = False
            Me.cmd_削除.Enabled = True

        Case Else
            Exit Sub
    End Select
End Sub

'==============================================================================*
'コンボボックスに値を設定する
'顧客区分に使用する。→まだ決まってないので保留。
'==============================================================================*
Private Sub READ_CMB_TXT(ByVal oName As Object)

  Dim objControl As ComboBox
  Dim vRet       As Variant

  Set objControl = oName
  vRet = objControl.Column(1, objControl.ListIndex)

  Select Case objControl.NAME
    Case "cmb_顧客区分"
'      Me.txt_TORKBNNM = vRet
  End Select
End Sub

'==============================================================================*
'ＤＢの更新処理
'==============================================================================*
Private Function WRI_RTN() As Boolean

On Error GoTo err_rtn

    WRI_RTN = False

    doCmd.Hourglass True
    doCmd.SetWarnings False

    Set WRI_ObjWS = DBEngine(0)
    Set WRI_objDb = CurrentDb

    'トランザクション開始
    WRI_ObjWS.BeginTrans

    '処理モード毎にＳＱＬの発行
    Select Case intModeflg
        Case intModeInsert              '新規 INSERT
            If WRI_USER_MAST_RTN(intModeInsert) = False Then
                GoTo Update_Err
            End If
        Case intModeUpdate              '更新 DELETE→INSERT
            If WRI_USER_MAST_RTN(intModeUpdate) = False Then
                GoTo Update_Err
            End If
            If WRI_USER_MAST_RTN(intModeInsert) = False Then
                GoTo Update_Err
            End If
        Case intModeDelete              '削除 DELETE
            If WRI_USER_MAST_RTN(intModeDelete) = False Then
                GoTo Update_Err
            End If
    End Select

    'コミット完了
    WRI_ObjWS.CommitTrans
    WRI_RTN = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set WRI_objDb = Nothing
    Set WRI_ObjWS = Nothing

    Exit Function
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Update_Err:
    WRI_ObjWS.Rollback
    MsgBox strmsg更新エラー, vbCritical, strMsg_kaku
    GoTo Exit_rtn
End Function

'==============================================================================*
'ＤＢの更新処理
'==============================================================================*
Private Function WRI_USER_MAST_RTN(ByVal intMode As Integer) As Boolean
  '引数 ０：正更新
  '     １：逆更新
    Dim vCode           As Variant
    Dim strSQL          As String
    Dim vPar(0 To 26)   As Variant                                              ' 2006/01/28 ADD T.SUZUKI
'    Dim vPar(0 To 25)   As Variant                                             ' 2006/01/28 DEL T.SUZUKI

    On Error GoTo err_rtn
    WRI_USER_MAST_RTN = False

    vCode = Nz(Me.cmb_顧客コード, "")

    If vCode = "" Then
        GoTo Exit_rtn
    End If

    If intMode = intModeInsert Then
        vPar(0) = Nz(CmNumeric(Me.cmb_顧客コード), 0)
        vPar(1) = Me.txt_顧客名
        vPar(2) = Me.txt_顧客名カナ
        vPar(3) = Me.txt_代表者名
        vPar(4) = Me.txt_代表者名カナ
        vPar(5) = Me.txt_郵便番号
        vPar(6) = Me.txt_住所1
        vPar(7) = Me.txt_住所2
        vPar(8) = Me.txt_住所3    '(2005/07/13 iguchi ADD) <<<
        vPar(9) = Me.txt_電話番号
        vPar(10) = Me.txt_携帯番号
        vPar(11) = Me.txt_FAX番号
        vPar(12) = Me.txt_メールアドレス
        vPar(13) = Nz(CmNumeric(Me.txt_締日), 0)
        vPar(14) = Nz(CmNumeric(Me.cmb_顧客区分), 0)
        vPar(15) = Nz(CmNumeric(Me.cmb_集金区分), 0)
        vPar(16) = Me.txt_備考

' 2005/07/13 ADD C.IGUCHI Start
        vPar(17) = Me.txt_連絡先担当者名
        vPar(18) = Me.txt_連絡先担当者名カナ
        vPar(19) = Me.txt_連絡先郵便番号
        vPar(20) = Me.txt_連絡先都道府県市町村
        vPar(21) = Me.txt_連絡先番地
        vPar(22) = Me.txt_連絡先建物名等
        vPar(23) = Me.txt_連絡先電話番号
        vPar(24) = Me.txt_連絡先携帯番号
        vPar(25) = Me.txt_連絡先FAX番号
' 2005/07/13 ADD C.IGUCHI End
        vPar(26) = Me.cmb_新顧客コード                                          ' 2006/01/28 ADD T.SUZUKI

        strSQL = ""
        strSQL = " INSERT INTO dbo_USER_MAST "
        strSQL = strSQL & " ( "
        strSQL = strSQL & " USER_CODE,"         '0
        strSQL = strSQL & " USER_NAME,"         '1
        strSQL = strSQL & " USER_KANA,"         '2
        strSQL = strSQL & " USER_TANM,"         '3
        strSQL = strSQL & " USER_TAKA,"         '4
        strSQL = strSQL & " USER_YUBINO,"       '5
        strSQL = strSQL & " USER_ADR_1,"        '6
        strSQL = strSQL & " USER_ADR_2,"        '7
        strSQL = strSQL & " USER_ADR_3,"        '8
        strSQL = strSQL & " USER_TEL,"          '9
        strSQL = strSQL & " USER_KEITAI,"       '10
        strSQL = strSQL & " USER_FAX,"          '11
        strSQL = strSQL & " USER_MAIL,"         '12
        strSQL = strSQL & " USER_SDAY,"         '13
        strSQL = strSQL & " USER_KKBN,"         '14
        strSQL = strSQL & " USER_SKBN,"         '15
        strSQL = strSQL & " USER_BIKO,"         '16
        strSQL = strSQL & " USER_RNAME,"        '17
        strSQL = strSQL & " USER_RKANA,"        '18
        strSQL = strSQL & " USER_RPOST,"        '19
        strSQL = strSQL & " USER_RADR_1,"       '20
        strSQL = strSQL & " USER_RADR_2,"       '21
        strSQL = strSQL & " USER_RADR_3,"       '22
        strSQL = strSQL & " USER_RTEL,"         '23
        strSQL = strSQL & " USER_RKEITAI,"      '24
        strSQL = strSQL & " USER_RFAX,"         '25
        strSQL = strSQL & " USER_UPDATE,"       '26
        strSQL = strSQL & " USER_NEW_CODE"      '27                             ' 2006/01/28 ADD T.SUZUKI
        strSQL = strSQL & " )"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & vPar(0) & ","
        strSQL = strSQL & CmParChar(vPar(1)) & ","
        strSQL = strSQL & CmParChar(vPar(2)) & ","
        strSQL = strSQL & CmParChar(vPar(3)) & ","
        strSQL = strSQL & CmParChar(vPar(4)) & ","
        strSQL = strSQL & CmParChar(vPar(5)) & ","
        strSQL = strSQL & CmParChar(vPar(6)) & ","
        strSQL = strSQL & CmParChar(vPar(7)) & ","
        strSQL = strSQL & CmParChar(vPar(8)) & ","
        strSQL = strSQL & CmParChar(vPar(9)) & ","
        strSQL = strSQL & CmParChar(vPar(10)) & ","
        strSQL = strSQL & CmParChar(vPar(11)) & ","
        strSQL = strSQL & CmParChar(vPar(12)) & ","
        strSQL = strSQL & vPar(13) & ","
        strSQL = strSQL & vPar(14) & ","
        strSQL = strSQL & vPar(15) & ","
        strSQL = strSQL & CmParChar(vPar(16)) & ","
' 2005/07/13 ADD C.IGUCHI Start
        strSQL = strSQL & CmParChar(vPar(17)) & ","
        strSQL = strSQL & CmParChar(vPar(18)) & ","
        strSQL = strSQL & CmParChar(vPar(19)) & ","
        strSQL = strSQL & CmParChar(vPar(20)) & ","
        strSQL = strSQL & CmParChar(vPar(21)) & ","
        strSQL = strSQL & CmParChar(vPar(22)) & ","
        strSQL = strSQL & CmParChar(vPar(23)) & ","
        strSQL = strSQL & CmParChar(vPar(24)) & ","
        strSQL = strSQL & CmParChar(vPar(25)) & ","
' 2005/07/13 ADD C.IGUCHI End
' 2006/01/28 REP T.SUZUKI Start
        strSQL = strSQL & " Cdate(#" & Format(DATE, "yyyy/mm/dd") & "#), "
        strSQL = strSQL & CmParChar(vPar(26))
'        strSQL = strSQL & " Cdate(#" & Format(DATE, "yyyy/mm/dd") & "#)"
' 2006/01/28 REP T.SUZUKI End
        strSQL = strSQL & " )"

        WRI_objDb.Execute strSQL

    Else
        strSQL = ""
        strSQL = " DELETE FROM dbo_USER_MAST"
        strSQL = strSQL & " WHERE USER_CODE = " & (vCode) 'CmParCharかけると落ちる。

        WRI_objDb.Execute strSQL
    End If

    WRI_USER_MAST_RTN = True

Exit_rtn:
    Erase vPar
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_KeyDown イベント
'        MODULE_ID        :cmb_顧客コード_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intAns      As Integer

    If intModeflg <> intModeDelete Then
        Select Case KeyCode
            Case vbKeyUp
                Me.cmb_顧客コード.SetFocus
            Case vbKeyLeft
                Me.cmb_顧客コード.SetFocus
            'シフト＆タブのときだけ、処理選択に戻す。
            Case vbKeyTab
                If Shift > 0 Then
                    '変更フラグがTrueの時は確認メッセージを出す
                    If blnDataUpdate = True Then
                        intAns = MsgBox(strmsgクリア確認, vbQuestion + vbYesNo, strMsg_kaku)
                        If intAns = vbYes Then
                            ' フォームの初期化処理
                            Call CLR_RTN
                            Me.opg_処理選択.SetFocus
                        End If
                    Else
                        ' フォームの初期化処理
                        Call CLR_RTN
                        Me.opg_処理選択.SetFocus
                    End If
                End If
            Case Else
        End Select
    Else
        Select Case KeyCode
            Case vbKeyTab
                If Shift > 0 Then
                    '変更フラグがTrueの時は確認メッセージを出す
                    If blnDataUpdate = True Then
                        intAns = MsgBox(strmsgクリア確認, vbQuestion + vbYesNo, strMsg_kaku)
                        If intAns = vbYes Then
                            ' フォームの初期化処理
                            Call CLR_RTN
                            Me.opg_処理選択.SetFocus
                        End If
                    Else
                        ' フォームの初期化処理
                        Call CLR_RTN
                        Me.opg_処理選択.SetFocus
                    End If
                End If
        End Select
    End If
    intModeflg = Me.opg_処理選択.VALUE
    Call s_bottunset(intModeflg)
   
End Sub

'==============================================================================*
'ボタンの切り替え
'==============================================================================*
Private Function s_bottunset(intMode As Integer)

    If intMode <> intModeDelete Then
        Me.cmd_保存.Enabled = True
        Me.cmd_削除.Enabled = False
    Else
        Me.cmd_保存.Enabled = False
        Me.cmd_削除.Enabled = True
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_KeyPress イベント
'        MODULE_ID        :cmb_顧客コード_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客コード, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客区分_AfterUpdate イベント
'        MODULE_ID        :cmb_顧客区分_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客区分_AfterUpdate()
    
    Dim strRyak     As String
    Dim vCode       As Variant

    vCode = Nz(Me.cmb_顧客区分, "")
    If vCode = "" Then
        Me.txt_顧客区分 = ""
        Exit Sub
    End If

    strRyak = s_SelectRyak(Me.cmb_顧客区分, "020")
    Me.txt_顧客区分 = strRyak

    blnDataUpdate = True

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客区分_BeforeUpdate イベント
'        MODULE_ID        :cmb_顧客区分_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客区分_BeforeUpdate(Cancel As Integer)
    
    Dim vCode       As Variant
    Dim strRyak     As String

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.cmb_顧客区分, "")
    If vCode = "" Then
        Cancel = True
        MsgBox "顧客区分" & strMSG入力不正, vbInformation, strMsg_kaku
        Me.txt_顧客区分 = ""
        Exit Sub
    End If

    strRyak = s_SelectRyak(Me.cmb_顧客区分, "020")

    If strRyak = "" Then
        MsgBox strMSGデータなし, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.cmb_顧客区分) <> "" Then
        If Trim(Len(Me.cmb_顧客区分)) > 2 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_集金区分_AfterUpdate イベント
'        MODULE_ID        :cmb_集金区分_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_AfterUpdate()

    Dim strRyak     As String
    Dim strWhere    As String
    Dim vCode       As Variant
    
    vCode = Nz(Me.cmb_集金区分, "")
    If vCode = "" Then
        Me.txt_集金区分 = ""
        Exit Sub
    End If
    
    strRyak = s_SelectRyak(Me.cmb_集金区分, "030")
    Me.txt_集金区分 = strRyak

    blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_集金区分_BeforeUpdate イベント
'        MODULE_ID        :cmb_集金区分_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_BeforeUpdate(Cancel As Integer)
    
    Dim vCode       As Variant
    Dim strRyak     As String
    Dim strWhere    As String

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.cmb_集金区分, "")
    If vCode = "" Then
        Cancel = True
        MsgBox "集金区分" & strMSG入力不正, vbInformation, strMsg_kaku
        Me.txt_集金区分 = ""
        Exit Sub
    End If

    strRyak = s_SelectRyak(Me.cmb_集金区分, "030")

    If strRyak = "" Then
        MsgBox strMSGデータなし, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    If Trim(Me.cmb_集金区分) <> "" Then
        If Trim(Len(Me.cmb_集金区分)) > 2 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

    '' 2013/09/19 M.HONDA INS
    '' 元がコンビニの場合にはチェックしない。
    If strKkbn <> "07" Then
        If vCode = "07" Or vCode = "7" Then
            Cancel = True
            MsgBox "集金区分" & strMSGコンビニ, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
        End If
    End If
    '' 2013/09/19 M.HONDA INS

Exit_rtn:
    Exit Sub
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客区分_GotFocus イベント
'        MODULE_ID        :cmb_顧客区分_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客区分_GotFocus()
    'コンボボックスにはRequery
    Me.cmb_顧客区分.Requery
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_AfterUpdate イベント
'        MODULE_ID        :cmb_顧客コード_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_AfterUpdate()

    Dim strSQL      As String
    Dim vCode       As Variant
    Dim intRet      As Integer
    Dim intAns      As Integer
    Dim blnRet      As Boolean

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_顧客コード, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        ' フォームの初期化処理
        Call CLR_RTN

        Me.cmb_顧客コード = vBack_Key
        
'''''        Me.opg_処理選択.SetFocus
        Me.cmb_顧客コード.SetFocus
        Exit Sub
    End If

    ' 顧客コードコンボボックスにて選択した顧客コードのデータを画面に表示
' 2006/01/28 ADD T.SUZUKI Start
    Call DSP_MODEUPDATE_USER_MAST(0)
    If Nz(Me.cmb_新顧客コード, "") <> "" Then
        Call DSP_MODEUPDATE_USER_MAST(1)
    End If
' 2006/01/28 ADD T.SUZUKI End
'    Call DSP_MODEUPDATE_USER_MAST(Me.opg_処理選択.Value)                    ' 2006/01/28 DEL T.SUZUKI

    ' 【その他連絡先】の使用可否、背景色設定
    Call subTxtControl  ' 2005/07/13 ADD C.IGUCHII

    '()'存在チェック
    '()Select Case intModeflg
    '()    Case intModeUpdate, intModeDelete
    '()        blnRet = CHK_USER_CODE_EXIST
    '()        If blnRet = False Then
    '()            MsgBox strmsgデータなし, vbOKOnly, strMsg_kaku
    '()        Else
    '()            'データの呼び出し
    '()            Call DSP_MODEUPDATE_USER_MAST(intModeUpdate)
    '()        End If
    '()    Case intModeInsert
    '()        blnRet = CHK_USER_CODE_EXIST
    '()        If blnRet Then
    '()            MsgBox strmsg重複, vbOKOnly, strMsg_kaku
    '()        End If
    '()End Select
    '()Me.cmb_顧客コード = VCode

    vBack_Key = Me.cmb_顧客コード
    Me.cmb_顧客コード = Format(Me.cmb_顧客コード, "000000")

    '()Call ETC_LOCK_RTN

Exit_rtn:
    On Error Resume Next
    Me.cmb_顧客コード.SetFocus
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_BeforeUpdate イベント
'        MODULE_ID        :cmb_顧客コード_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_BeforeUpdate(Cancel As Integer)
  
    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.cmb_顧客コード, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        MsgBox "顧客コードを入力してください。", vbInformation, strMsg_kaku
        Exit Sub
    End If

    '入力チェック
    '()If intModeflg <> intModeInsert Then
        Select Case CmBCheck_Char(vCode, 6)
            Case 0
    
            Case 2
                '桁数ｴﾗｰ
                MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
                Cancel = True
                GoTo Exit_rtn
            Case 9
                MsgBox "コードを入力してください。", vbInformation, strMsg_err
                Cancel = True
                SendKeys "{ESC}"
                GoTo Exit_rtn
        End Select

    'マスタへの存在チェック
        If CHK_USER_CODE_EXIST(0) = False Then                                        ' 2006/01/28 ADD T.SUZUKI
'        If CHK_USER_CODE_EXIST = False Then                                          ' 2006/01/28 DEL T.SUZUKI
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
        End If
    '()End If
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_GotFocus イベント
'        MODULE_ID        :cmb_顧客コード_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_GotFocus()
  'コンボボックスにはRequery
    Me.cmb_顧客コード.Requery
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客区分_KeyPress イベント
'        MODULE_ID        :cmb_顧客区分_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客区分, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_削除_Click イベント
'        MODULE_ID        :cmd_削除_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_Click()

    Dim intAns      As Integer
    
    If Nz(Me.cmb_顧客コード, "") = "" Then
        GoTo Exit_rtn
    End If

' 2006/05/17 Rep Tajima Start
'   If CHK_INPUTDATA = False Then
'        Exit Sub
'    End If

    If IsDeletePossible = False Then
        MsgBox "この顧客コードで契約か予約がされているため削除は出来ません。", vbInformation + vbOKOnly, strMsg_err
        Exit Sub
    End If
' 2006/05/17 Rep Tajima End

    intAns = MsgBox(strmsg削除確認, vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_rtn
    End If

    If WRI_RTN = True Then
        ' フォームの初期化処理
        Call CLR_RTN

' 2005/07/13 ADD C.IGUSHI Start
        ' 【その他連絡先】の使用可否、背景色設定
        Call subTxtControl
'''        Call chk_入力_Click
' 2005/07/13 ADD C.IGUSHI End
    End If

Exit_rtn:
    Me.cmb_顧客コード.SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_取消_Click イベント
'        MODULE_ID        :cmd_取消_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()

    Dim intAns      As Integer

    intAns = MsgBox(strmsg取消確認, vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)
    If intAns = vbNo Then
        GoTo Exit_rtn
    End If

    ' フォームの初期化処理
    Call CLR_RTN

' 2005/07/13 ADD C.IGUSHI Start
    ' 【その他連絡先】の使用可否、背景色設定
    Call subTxtControl
'''    Call chk_入力_Click
' 2005/07/13 ADD C.IGUSHI End

    blnDataUpdate = False
    Me.cmb_顧客コード.Requery
    is顧客情報変更 = False '2007/06/10 add tajima

Exit_rtn:
    Me.cmb_顧客コード.SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click イベント
'        MODULE_ID        :cmd_終了_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

    Dim intAns      As Integer

    On Error GoTo Err_cmd_終了_Click
    
    If blnDataUpdate = True Then
        '初期値：「いいえ」
        intAns = MsgBox("修正内容を保存せずに" & strmsg終了確認, vbQuestion + vbYesNo + vbDefaultButton1, "確認")
        If intAns = vbNo Then
            Exit Sub
        End If
    End If
    
    doCmd.Close acForm, Me.FormName
    Exit Sub

Err_cmd_終了_Click:
    MsgBox Error$
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_Click イベント
'        MODULE_ID        :cmd_保存_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()

    Dim intAns  As Integer
    Dim nextParams  As String

    On Error GoTo err_rtn

    If CHK_INPUTDATA = False Then
        Exit Sub
    End If

    '確認メッセージ
    intAns = MsgBox(strmsg保存確認, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_rtn
    End If

    If func保存処理 = True Then '2007/06/10 fix tajima
' 2006/02/07 ADD T.SUZUKI Start
        ' 変更フラグを設定
        pblnChangeFlg = True                                                ' 2006/02/07 ADD T.SUZUKI

        '2019/04/23 M.HONDA DEL
        '▼ 2007/06/10 add tajima
        'If Nz(Me.cmb_顧客情報変更理由.VALUE, "00") <> "00" And is顧客情報変更 = True Then
        '    If vbYes = MsgBox("顧客情報変更として登録しました" & vbCr & _
        '                      "契約リストへの印刷を行いますか？", _
        '                        vbQuestion + vbYesNo, strMsg_kaku) Then
        '        nextParams = Me.cmb_顧客コード.VALUE & "," _
        '                      & Me.txt_顧客情報変更日.VALUE & "," _
        '                      & Me.cmb_顧客情報変更理由.VALUE & "," _
        '                      & "04"  '保証区分は04:リプラス前提としておく
        '        doCmd.OpenForm "FKS140", , , , , , nextParams
        '    End If
        'End If
        '▲ 2007/06/10 add tajima
        '2019/04/23 M.HONDA DEL
        
        If pblnReceiptMode = True Then
            ' 画面を閉じる
            doCmd.Close acForm, Me.FormName: Exit Sub
        End If
' 2006/02/07 ADD T.SUZUKI End

        ' フォームの初期化処理
        Call CLR_RTN

        ' 【その他連絡先】の使用可否、背景色設定
        Call subTxtControl
    
    End If

    blnDataUpdate = False
    is顧客情報変更 = False    ' 2007/06/10 add tajima

Exit_rtn:
    Me.cmb_顧客コード.SetFocus
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opg_処理選択_AfterUpdate イベント
'        MODULE_ID        :opg_処理選択_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_処理選択_AfterUpdate()

    ' フラグ初期化
    blnDataUpdate = False
    is顧客情報変更 = False '2007/06/10 add tajima

    ' フォームの初期化処理
    Call CLR_RTN

    Me.cmb_顧客コード.Requery

    ' 処理モード保持
    intModeflg = Me.opg_処理選択.VALUE

    Call s_bottunset(intModeflg)

    ' 画面項目の使用可否設定
    Call Etc_Lock_Rtn

    ' 【その他連絡先】の使用可否、背景色設定
    Call subTxtControl
'''    Call chk_入力_Click

    Me.cmb_顧客コード.SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opg_処理選択_BeforeUpdate イベント
'        MODULE_ID        :opg_処理選択_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_処理選択_BeforeUpdate(Cancel As Integer)

    Dim intAns      As Integer

    '変更フラグがTrueの時は確認メッセージを出す
    If blnDataUpdate = True Then
        intAns = MsgBox(strmsgクリア確認, vbQuestion + vbYesNo, strMsg_kaku)
        If intAns = vbNo Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If

       ' フォームの初期化処理
       Call CLR_RTN
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_FAX番号_AfterUpdate イベント
'        MODULE_ID        :txt_FAX番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_FAX番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_FAX番号_BeforeUpdate イベント
'        MODULE_ID        :txt_FAX番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_FAX番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_FAX番号) <> "" Then
        If Trim(Len(Me.txt_FAX番号)) > 15 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_FAX番号_KeyPress イベント
'        MODULE_ID        :txt_FAX番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_FAX番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_FAX番号, KeyAscii, 15)     '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_BIRTHDAY_AfterUpdate
'        機能             :
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_BIRTHDAY_AfterUpdate()
    Me.txt_USER_BIRTHDAY = Format(var変更日付, "yyyy/mm/dd")
    blnDataUpdate = True
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_USER_BIRTHDAY_BeforeUpdate
'        機能             :
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_BIRTHDAY_BeforeUpdate(Cancel As Integer)
    Call isDateFormat(Me.txt_USER_BIRTHDAY, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_USER_BIRTHDAY_GotFocus
'        機能             :
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_BIRTHDAY_GotFocus()
    txt_USER_BIRTHDAY = Format(txt_USER_BIRTHDAY, "yy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_USER_BIRTHDAY_KeyPress
'        機能             :
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_BIRTHDAY_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(txt_USER_BIRTHDAY, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_USER_BIRTHDAY_LostFocus
'        機能             :
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_BIRTHDAY_LostFocus()
    txt_USER_BIRTHDAY = Format(txt_USER_BIRTHDAY, "yyyy/mm/dd")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_JOB_AfterUpdate イベント
'        MODULE_ID        :txt_USER_JOB_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_JOB_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_JOB_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_JOB_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_JOB_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_USER_JOB, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 50)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KADR_1_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KADR_1_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KADR_1_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KADR_1_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KADR_1_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KADR_1_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_USER_KADR_1, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KADR_2_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KADR_2_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KADR_2_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KADR_2_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KADR_2_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KADR_2_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_USER_KADR_2, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KADR_3_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KADR_3_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KADR_3_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KADR_3_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KADR_3_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KADR_3_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_USER_KADR_3, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KHOUR_END_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KHOUR_END_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KHOUR_END_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KHOUR_END_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KHOUR_END_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KHOUR_END_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_USER_KHOUR_END) <> "" Then
        If Trim(Len(Me.txt_USER_KHOUR_END)) > 5 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KHOUR_START_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KHOUR_START_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KHOUR_START_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KHOUR_START_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KHOUR_START_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KHOUR_START_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_USER_KHOUR_START) <> "" Then
        If Trim(Len(Me.txt_USER_KHOUR_START)) > 5 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KNAME_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KNAME_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KNAME_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KNAME_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KNAME_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KNAME_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_USER_KNAME, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 50)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KKANA_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KKANA_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KKANA_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KKANA_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KKANA_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KKANA_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_USER_KKANA, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 40)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KPOST_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KPOST_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KPOST_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KPOST_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KPOST_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KPOST_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_USER_KPOST) <> "" Then
        If Trim(Len(Me.txt_USER_KPOST)) > 10 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KPOST_KeyPress イベント
'        MODULE_ID        :txt_USER_KPOST_KeyPress
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KPOST_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_USER_KPOST, KeyAscii, 10)  '←DB定義より10桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KTEL_AfterUpdate イベント
'        MODULE_ID        :txt_USER_KTEL_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KTEL_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KTEL_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_KTEL_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KTEL_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_USER_KTEL) <> "" Then
        If Trim(Len(Me.txt_USER_KTEL)) > 15 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_KTEL_KeyPress イベント
'        MODULE_ID        :txt_USER_KTEL_KeyPress
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_KTEL_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_USER_KTEL, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_NENSHU_AfterUpdate イベント
'        MODULE_ID        :txt_USER_NENSHU_AfterUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_NENSHU_AfterUpdate()

  '変更フラグ
  blnDataUpdate = True

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_NENSHU_BeforeUpdate イベント
'        MODULE_ID        :txt_USER_NENSHU_BeforeUpdate
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_NENSHU_BeforeUpdate(Cancel As Integer)

    Dim strNumber       As String
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If
    
    strNumber = Nz(txt_USER_NENSHU, "")
    '空白なら入力チェックしない
    If Nz(txt_USER_NENSHU, "") = "" Then
        Exit Sub
    End If
    
    If Val(Nz(txt_USER_NENSHU, 0)) > 999999999 Then
        '桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation
        Cancel = True
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_NENSHU_GotFocus イベント
'        MODULE_ID        :txt_USER_NENSHU_GotFocus
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_NENSHU_GotFocus()
    
    txt_USER_NENSHU.Format = "Fixed"

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_NENSHU_KeyPress イベント
'        MODULE_ID        :txt_USER_NENSHU_KeyPress
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_NENSHU_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_USER_NENSHU, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_USER_NENSHU_LostFocus イベント
'        MODULE_ID        :txt_USER_NENSHU_LostFocus
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_USER_NENSHU_LostFocus()

    txt_USER_NENSHU.Format = "Standard"

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_メールアドレス_AfterUpdate イベント
'        MODULE_ID        :txt_メールアドレス_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_メールアドレス_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_メールアドレス_BeforeUpdate イベント
'        MODULE_ID        :txt_メールアドレス_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_メールアドレス_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_メールアドレス) <> "" Then
        If Trim(Len(Me.txt_メールアドレス)) > 60 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_携帯番号_AfterUpdate イベント
'        MODULE_ID        :txt_携帯番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_携帯番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_携帯番号_BeforeUpdate イベント
'        MODULE_ID        :txt_携帯番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_携帯番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_携帯番号) <> "" Then
        If Trim(Len(Me.txt_携帯番号)) > 15 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_携帯番号_KeyPress イベント
'        MODULE_ID        :txt_携帯番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_携帯番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_携帯番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_最後_GotFocus イベント
'        MODULE_ID        :txt_最後_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_最後_GotFocus()
    If Me.cmd_保存.Enabled = True Then
'''''        Me.cmb_顧客コード.SetFocus     'DEL 99/12/20
        Me.cmd_保存.SetFocus          'ADD 99/12/20
    Else
        Me.cmd_削除.SetFocus
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_住所2_AfterUpdate イベント
'        MODULE_ID        :txt_住所2_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所2_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_住所2_BeforeUpdate イベント
'        MODULE_ID        :txt_住所2_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所2_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_住所2, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客名_AfterUpdate イベント
'        MODULE_ID        :txt_顧客名_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客名_BeforeUpdate イベント
'        MODULE_ID        :txt_顧客名_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_顧客名, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客名カナ_AfterUpdate イベント
'        MODULE_ID        :txt_顧客名カナ_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名カナ_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客名カナ_BeforeUpdate イベント
'        MODULE_ID        :txt_顧客名カナ_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名カナ_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_顧客名カナ, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_住所3_AfterUpdate イベント
'        MODULE_ID        :txt_住所3_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'(2005/6/17iguchi ADD)START
Private Sub txt_住所3_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_住所3_BeforeUpdate イベント
'        MODULE_ID        :txt_住所3_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所3_BeforeUpdate(Cancel As Integer)
 Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If
    vCode = Nz(Me.txt_住所3, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'(2005/6/17iguchi ADD)END

'==============================================================================*
'
'        MODULE_NAME      :txt_代表者名カナ_AfterUpdate イベント
'        MODULE_ID        :txt_代表者名カナ_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名カナ_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_代表者名カナ_BeforeUpdate イベント
'        MODULE_ID        :txt_代表者名カナ_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名カナ_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_代表者名カナ, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_締日_AfterUpdate イベント
'        MODULE_ID        :txt_締日_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_締日_AfterUpdate()
    If Nz(Me.txt_締日, "") = "" Then
        Me.txt_締日 = "30"
    End If
    '変更フラグ
    blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_締日_BeforeUpdate イベント
'        MODULE_ID        :txt_締日_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_締日_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If
    'CmBCheck_Numをかける。
    If Nz(Me.txt_締日, "") = "" Then
'        MsgBox "締日" & strmsg入力不正, vbInformation, strMsg_Err
'        Cancel = True
        GoTo Exit_rtn
    End If

    If Trim(Me.txt_締日) <> "" Then
        If Trim(Len(Me.txt_締日)) > 2 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
        End If
    End If
    If CInt(Me.txt_締日) > 30 Then
        MsgBox strmsg締日, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_締日_KeyPress イベント
'        MODULE_ID        :txt_締日_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_締日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_締日, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_電話番号_AfterUpdate イベント
'        MODULE_ID        :txt_電話番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_電話番号_BeforeUpdate イベント
'        MODULE_ID        :txt_電話番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_電話番号) <> "" Then
        If Trim(Len(Me.txt_電話番号)) > 15 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_電話番号_KeyPress イベント
'        MODULE_ID        :txt_電話番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_電話番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_備考_AfterUpdate イベント
'        MODULE_ID        :txt_備考_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_備考_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_備考_BeforeUpdate イベント
'        MODULE_ID        :txt_備考_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_備考_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_備考, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    'Select Case CmBCheck_Char(vCode, 36) 'DELETE 20041207 N.MIURA
    Select Case CmBCheck_Char(vCode, 60)  'INSERT 20041207 N.MIURA
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

' 2005/07/13 DEL T.SUZUKI Start
''==============================================================================*
''txt_備考_KeyDownイベント
''==============================================================================*
'Private Sub txt_備考_KeyDown(KeyCode As Integer, Shift As Integer)

'    If Shift = 0 And KeyCode = vbKeyTab Then
'        Select Case intModeflg
'            Case intModeInsert, intModeUpdate
'                Me.cmd_保存.SetFocus
'            Case intModeDelete
'                Me.cmd_削除.SetFocus
'        End Select
'        KeyCode = 0
'    End If
'End Sub
' 2005/07/13 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :txt_郵便番号_AfterUpdate イベント
'        MODULE_ID        :txt_郵便番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_郵便番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_郵便番号_BeforeUpdate イベント
'        MODULE_ID        :txt_郵便番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_郵便番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_郵便番号) <> "" Then
        If Trim(Len(Me.txt_郵便番号)) > 10 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_郵便番号_KeyPress イベント
'        MODULE_ID        :txt_郵便番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_郵便番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_郵便番号, KeyAscii, 10)  '←DB定義より10桁まで入力可。
End Sub

' 2005/07/13 ADD T.SUZUKI Start
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Yubin_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Yubin_Click()
    Call MSZ0000_YUBIB(Me.NAME, Me.txt_郵便番号.NAME, Me.txt_住所1.NAME, _
                                Nz(Me.txt_郵便番号.VALUE), Nz(Me.txt_住所1.VALUE))
End Sub
' 2005/07/13 ADD T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :txt_住所1_AfterUpdate イベント
'        MODULE_ID        :txt_住所1_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所1_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_住所1_BeforeUpdate イベント
'        MODULE_ID        :txt_住所1_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所1_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_住所1, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_代表者名_AfterUpdate イベント
'        MODULE_ID        :txt_代表者名_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_代表者名_BeforeUpdate イベント
'        MODULE_ID        :txt_代表者名_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_代表者名, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先FAX番号_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先FAX番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先FAX番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先FAX番号_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先FAX番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先FAX番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_連絡先FAX番号) <> "" Then
        If Trim(Len(Me.txt_連絡先FAX番号)) > 15 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

' 2005/07/13 ADD T.SUZUKI Start
'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先FAX番号_KeyPress イベント
'        MODULE_ID        :txt_連絡先FAX番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先FAX番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_連絡先FAX番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先FAX番号_KeyDown イベント
'        MODULE_ID        :txt_連絡先FAX番号_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先FAX番号_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 0 And KeyCode = vbKeyTab Then
        Select Case intModeflg
            Case intModeInsert, intModeUpdate
                Me.cmd_保存.SetFocus
            Case intModeDelete
                Me.cmd_削除.SetFocus
        End Select
        KeyCode = 0
    End If
End Sub
' 2005/07/13 ADD T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先携帯番号_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先携帯番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先携帯番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先携帯番号_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先携帯番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先携帯番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_連絡先携帯番号) <> "" Then
        If Trim(Len(Me.txt_連絡先携帯番号)) > 15 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先携帯番号_KeyPress イベント
'        MODULE_ID        :txt_連絡先携帯番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先携帯番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_連絡先携帯番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先建物名等_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先建物名等_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先建物名等_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先建物名等_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先建物名等_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先建物名等_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_連絡先建物名等, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先担当者名_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先担当者名_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当者名_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先担当者名_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先担当者名_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当者名_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_連絡先担当者名, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
'    Select Case CmBCheck_Char(vCode, 18)  '2006/05/17 フィールドサイズ36→18へ tajima
    Select Case CmBCheck_Char(vCode, 36)  '2011/11/29 フィールドサイズ18→36へ M.RYU
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先担当者名カナ_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先担当者名カナ_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当者名カナ_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先担当者名カナ_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先担当者名カナ_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当者名カナ_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant
    
    On Error GoTo err_rtn
      
    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_連絡先担当者名カナ, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先電話番号_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先電話番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先電話番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先電話番号_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先電話番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先電話番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_連絡先電話番号) <> "" Then
        If Trim(Len(Me.txt_連絡先電話番号)) > 15 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先電話番号_KeyPress イベント
'        MODULE_ID        :txt_連絡先電話番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先電話番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_連絡先電話番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先都道府県市町村_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先都道府県市町村_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先都道府県市町村_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先都道府県市町村_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先都道府県市町村_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先都道府県市町村_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_連絡先都道府県市町村, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先番地_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先番地_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先番地_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先番地_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先番地_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先番地_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant

    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    vCode = Nz(Me.txt_連絡先番地, "")

    '何も入力されていない時は、処理をしない
    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0
        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        Case 9
    End Select
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先郵便番号_AfterUpdate イベント
'        MODULE_ID        :txt_連絡先郵便番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先郵便番号_AfterUpdate()
  '変更フラグ
  blnDataUpdate = True
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先郵便番号_BeforeUpdate イベント
'        MODULE_ID        :txt_連絡先郵便番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先郵便番号_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo err_rtn

    If blnFormLoadFlg = False Then
        Exit Sub
    End If

    If Trim(Me.txt_連絡先郵便番号) <> "" Then
        If Trim(Len(Me.txt_連絡先郵便番号)) > 10 Then
            MsgBox strMSG桁数オーバー, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_連絡先郵便番号_KeyPress イベント
'        MODULE_ID        :txt_連絡先郵便番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先郵便番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_連絡先郵便番号, KeyAscii, 10)  '←DB定義より10桁まで入力可。
End Sub

' 2005/07/13 ADD T.SUZUKI Start
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_RYubin_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_RYubin_Click()
    Call MSZ0000_YUBIB(Me.NAME, Me.txt_連絡先郵便番号.NAME, Me.txt_連絡先都道府県市町村.NAME, _
                                Nz(Me.txt_連絡先郵便番号.VALUE), Nz(Me.txt_連絡先都道府県市町村.VALUE))
End Sub
' 2005/07/13 ADD T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :削除可能可否
'        MODULE_ID        :IsDeletePossible
'        CREATE_DATE      :2006/05/17
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function IsDeletePossible() As Boolean

    Dim vRet   As Variant
    Dim strSQL As String
    
    IsDeletePossible = False

    On Error GoTo Finally

    vRet = Nz(Me.cmb_顧客コード, "")
    If vRet = "" Then Exit Function

    ' 契約トランと予約トランでこの顧客コードを扱っているかチェック
    strSQL = " SELECT COUNT(*) FROM dbo_CARG_FILE WHERE dbo_CARG_FILE.CARG_UCODE = " & vRet
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenForwardOnly)
    If Not (objRs.EOF) Then
      If 0 < Nz(objRs(0), 0) Then
        GoTo Finally
      End If
   End If
    
    strSQL = " SELECT COUNT(*) FROM dbo_YOUK_TRAN WHERE dbo_YOUK_TRAN.YOUKT_UCODE = " & vRet
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenForwardOnly)
    If Not (objRs.EOF) Then
      If 0 < Nz(objRs(0), 0) Then
        GoTo Finally
      End If
   End If
   
   IsDeletePossible = True

Finally:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
End Function
'★☆★☆★☆★☆★☆★★☆☆ 2007/06/10 add tajima ★☆★☆★☆★☆★☆★☆★☆
'==============================================================================*
'
'        MODULE_NAME      :func保存処理
'        機能             :保存処理の実行
'        IN               :
'        OUT              :True=正常 False=異常
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function func保存処理()

    Dim isException     As Boolean
    Dim bumonCode       As String
    Dim adoDbConnection As Object       'ADODB.Connection
    Dim userId          As String
    
    isException = False
    func保存処理 = True

On Error GoTo Exception

    '部門コード取得
    bumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If bumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("DLookup(""CONT_BUMOC"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", "テーブル[CONT_MAST]の設定不正。")
    End If

    'ユーザ情報取得
    userId = LsGetUserName()

    'ADO接続Object生成
    Set adoDbConnection = MSZZ025.ADODB_Connection(bumonCode)

    '処理モード毎の処理
    Select Case intModeflg
        '新規 INSERT
        Case intModeInsert
              Call insert_USER_MAST(adoDbConnection, userId)
        
        '更新 UPDATE
        Case intModeUpdate
              Call update_USER_MAST(adoDbConnection, Nz(Me.cmb_顧客コード.VALUE, ""), userId)
        
        '削除 DELETE
        Case intModeDelete
             ' 削除処理は2007/06/15現在無し
    End Select

    func保存処理 = True

    GoTo Finally
Exception:
    isException = True
Finally:
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing
    If isException Then
        Call Err.Raise(Err.Number, "func保存処理" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :update_USER_MAST
'        機能             :USER_MASTへの更新
'        IN               :
'        OUT              :True=正常 False=異常
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function update_USER_MAST(aConnection As Object, a顧客コード As String, a更新ユーザ As String)

    Dim isException     As Boolean
    Dim rsData          As Object       'ADODB.Recordset
    Dim sqlText         As String
    isException = False
    update_USER_MAST = False
    
On Error GoTo Exception
    
    sqlText = "SELECT * "
    sqlText = sqlText & " FROM USER_MAST "
    sqlText = sqlText & " WHERE USER_CODE = '" & a顧客コード & "' "
        
    Set rsData = MSZZ025.ADODB_Recordset(sqlText, aConnection, adoReadWrite)

    If Not rsData.EOF Then
        'データ更新
        
'----↓↓↓↓----DELETE 2011/07/29 M.RYU------↓↓↓↓--------<S>
'        rsData.Fields("USER_NAME") = Me.txt_顧客名
'        rsData.Fields("USER_KANA") = Me.txt_顧客名カナ
'        rsData.Fields("USER_TANM") = Me.txt_代表者名
'        rsData.Fields("USER_TAKA") = Me.txt_代表者名カナ
'        rsData.Fields("USER_YUBINO") = Me.txt_郵便番号
'        rsData.Fields("USER_ADR_1") = Me.txt_住所1
'        rsData.Fields("USER_ADR_2") = Me.txt_住所2
'        rsData.Fields("USER_ADR_3") = Me.txt_住所3
'        rsData.Fields("USER_TEL") = Me.txt_電話番号
'        rsData.Fields("USER_KEITAI") = Me.txt_携帯番号
'        rsData.Fields("USER_FAX") = Me.txt_FAX番号
'        rsData.Fields("USER_SDAY") = Nz(CmNumeric(Me.txt_締日), 0)
'        rsData.Fields("USER_KKBN") = Nz(CmNumeric(Me.cmb_顧客区分), 0)
'        rsData.Fields("USER_SKBN") = Nz(CmNumeric(Me.cmb_集金区分), 0)
'        rsData.Fields("USER_BIKO") = Me.txt_備考
'        rsData.Fields("USER_MAIL") = Me.txt_メールアドレス
'        rsData.Fields("USER_RNAME") = Me.txt_連絡先担当者名
'        rsData.Fields("USER_RKANA") = Me.txt_連絡先担当者名カナ
'        rsData.Fields("USER_RPOST") = Me.txt_連絡先郵便番号
'        rsData.Fields("USER_RADR_1") = Me.txt_連絡先都道府県市町村
'        rsData.Fields("USER_RADR_2") = Me.txt_連絡先番地
'        rsData.Fields("USER_RADR_3") = Me.txt_連絡先建物名等
'        rsData.Fields("USER_RTEL") = Me.txt_連絡先電話番号
'        rsData.Fields("USER_RKEITAI") = Me.txt_連絡先携帯番号
'        rsData.Fields("USER_RFAX") = Me.txt_連絡先FAX番号
'        rsData.Fields("USER_UPDATE") = DATE
'        If Nz(Me.cmb_新顧客コード, "") = "" Then
'            rsData.Fields("USER_NEW_CODE") = Null
'        Else
'            rsData.Fields("USER_NEW_CODE") = CmNumeric(Me.cmb_新顧客コード)
'        End If
'        rsData.Fields("USER_CHG_RKBN") = Nz(CmNumeric(Me.cmb_顧客情報変更理由), 0)
'        rsData.Fields("USER_CHG_DATE") = Me.txt_顧客情報変更日
'        '↓INSERT 2010/04/19 SHIBAZAKI
'        With rsData
'            .Fields("USER_KNAME") = txt_USER_KNAME
'            .Fields("USER_KKANA") = txt_USER_KKANA
'            .Fields("USER_JOB") = txt_USER_JOB
'            .Fields("USER_KPOST") = txt_USER_KPOST
'            .Fields("USER_KADR_1") = txt_USER_KADR_1
'            .Fields("USER_KADR_2") = txt_USER_KADR_2
'            .Fields("USER_KADR_3") = txt_USER_KADR_3
'            .Fields("USER_KTEL") = txt_USER_KTEL
'            .Fields("USER_KHOUR_START") = txt_USER_KHOUR_START
'            .Fields("USER_KHOUR_END") = txt_USER_KHOUR_END
'            .Fields("USER_NENSHU") = txt_USER_NENSHU
'            If Trim(Nz(txt_USER_BIRTHDAY, "")) <> "" Then
'                .Fields("USER_BIRTHDAY") = Replace(txt_USER_BIRTHDAY, "/", "")
'            Else                                                                    'INSERT 2010/04/28 SHIBAZAKI
'                .Fields("USER_BIRTHDAY") = txt_USER_BIRTHDAY                        'INSERT 2010/04/28 SHIBAZAKI
'            End If
'        End With
'        '↑INSERT 2010/04/19 SHIBAZAKI
'----↑↑↑↑----DELETE 2011/07/29 M.RYU------↑↑↑↑--------<E>

        Call setDataSave(rsData)        'INSERT 2011/07/29 M.RYU
        
        rsData.Fields("USER_UPDAD") = Format(DATE, "yyyymmdd")
        rsData.Fields("USER_UPDAJ") = Format(Now, "hhmmss")
        rsData.Fields("USER_UPDPB") = Me.NAME
        rsData.Fields("USER_UPDUB") = a更新ユーザ
    
        rsData.UPDATE
    End If
    update_USER_MAST = True

    GoTo Finally

Exception:
    isException = True
Finally:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If isException Then
        Call Err.Raise(Err.Number, "update_USER_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :insert_USER_MAST
'        機能             :USER_MASTへの追記
'        IN               :
'        OUT              :True=正常 False=異常
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function insert_USER_MAST(aConnection As Object, a作成ユーザ As String)

    Dim isException     As Boolean
    Dim rsData          As Object       'ADODB.Recordset
    Dim sqlText         As String
    isException = False
    insert_USER_MAST = False
    
On Error GoTo Exception
    
    Set rsData = MSZZ025.ADODB_Recordset("USER_MAST", aConnection, adoAppendOnly)

    'データ追加
    rsData.AddNew
    rsData.Fields("USER_CODE") = Me.cmb_顧客コード
    
'----↓↓↓↓----DELETE 2011/07/29 M.RYU------↓↓↓↓--------<S>
'    rsData.Fields("USER_NAME") = Me.txt_顧客名
'    rsData.Fields("USER_KANA") = Me.txt_顧客名カナ
'    rsData.Fields("USER_TANM") = Me.txt_代表者名
'    rsData.Fields("USER_TAKA") = Me.txt_代表者名カナ
'    rsData.Fields("USER_YUBINO") = Me.txt_郵便番号
'    rsData.Fields("USER_ADR_1") = Me.txt_住所1
'    rsData.Fields("USER_ADR_2") = Me.txt_住所2
'    rsData.Fields("USER_ADR_3") = Me.txt_住所3
'    rsData.Fields("USER_TEL") = Me.txt_電話番号
'    rsData.Fields("USER_KEITAI") = Me.txt_携帯番号
'    rsData.Fields("USER_FAX") = Me.txt_FAX番号
'    rsData.Fields("USER_SDAY") = Nz(CmNumeric(Me.txt_締日), 0)
'    rsData.Fields("USER_KKBN") = Nz(CmNumeric(Me.cmb_顧客区分), 0)
'    rsData.Fields("USER_SKBN") = Nz(CmNumeric(Me.cmb_集金区分), 0)
'    rsData.Fields("USER_BIKO") = Me.txt_備考
'    rsData.Fields("USER_MAIL") = Me.txt_メールアドレス
'    rsData.Fields("USER_RNAME") = Me.txt_連絡先担当者名
'    rsData.Fields("USER_RKANA") = Me.txt_連絡先担当者名カナ
'    rsData.Fields("USER_RPOST") = Me.txt_連絡先郵便番号
'    rsData.Fields("USER_RADR_1") = Me.txt_連絡先都道府県市町村
'    rsData.Fields("USER_RADR_2") = Me.txt_連絡先番地
'    rsData.Fields("USER_RADR_3") = Me.txt_連絡先建物名等
'    rsData.Fields("USER_RTEL") = Me.txt_連絡先電話番号
'    rsData.Fields("USER_RKEITAI") = Me.txt_連絡先携帯番号
'    rsData.Fields("USER_RFAX") = Me.txt_連絡先FAX番号
'    rsData.Fields("USER_UPDATE") = DATE
'    If Nz(Me.cmb_新顧客コード, "") = "" Then
'        rsData.Fields("USER_NEW_CODE") = Null
'    Else
'        rsData.Fields("USER_NEW_CODE") = CmNumeric(Me.cmb_新顧客コード)
'    End If
'    rsData.Fields("USER_CHG_RKBN") = Nz(CmNumeric(Me.cmb_顧客情報変更理由), 0)
'    rsData.Fields("USER_CHG_DATE") = Me.txt_顧客情報変更日
'    '↓INSERT 2010/04/19 SHIBAZAKI
'    With rsData
'        .Fields("USER_KNAME") = txt_USER_KNAME
'        .Fields("USER_KKANA") = txt_USER_KKANA
'        .Fields("USER_JOB") = txt_USER_JOB
'        .Fields("USER_KPOST") = txt_USER_KPOST
'        .Fields("USER_KADR_1") = txt_USER_KADR_1
'        .Fields("USER_KADR_2") = txt_USER_KADR_2
'        .Fields("USER_KADR_3") = txt_USER_KADR_3
'        .Fields("USER_KTEL") = txt_USER_KTEL                   'DELETE 2011/07/29 M.RYU
'        .Fields("USER_KHOUR_START") = txt_USER_KHOUR_START
'        .Fields("USER_KHOUR_END") = txt_USER_KHOUR_END
'        .Fields("USER_NENSHU") = txt_USER_NENSHU
'        If Trim(Nz(txt_USER_BIRTHDAY, "")) <> "" Then
'            .Fields("USER_BIRTHDAY") = Replace(txt_USER_BIRTHDAY, "/", "")
'        Else
'            .Fields("USER_BIRTHDAY") = txt_USER_BIRTHDAY
'        End If
'    End With
'    '↑INSERT 2010/04/19 SHIBAZAKI
'----↑↑↑↑----DELETE 2011/07/29 M.RYU------↑↑↑↑--------<E>

    Call setDataSave(rsData)        'INSERT 2011/07/29 M.RYU
    
    rsData.Fields("USER_INSED") = Format(DATE, "yyyymmdd")
    rsData.Fields("USER_INSEJ") = Format(Now, "hhmmss")
    rsData.Fields("USER_INSPB") = Me.NAME
    rsData.Fields("USER_INSUB") = a作成ユーザ
    
    rsData.UPDATE
    
    'INSERT 2018/01/26 N.IMAI Start
    '加瀬DBの仮想口座を更新
    Dim objKASE     As Object
    Dim strSQL      As String
    Dim strBUMOC    As String
    Dim rsKOUZM     As Recordset
    Set objKASE = MSZZ025.ADODB_Connection()
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    strSQL = ""
    strSQL = strSQL & " SELECT *"
'    strSql = strSql & "   FROM KOUZ_MAST"      'DELETE 2019/05/01 Y.WADA
    strSQL = strSQL & "   FROM KOUZ_MAST_J"     'INSERT 2019/05/01 Y.WADA
    strSQL = strSQL & "  WHERE KOUZM_BUMOC = '" & strBUMOC & "'"
    strSQL = strSQL & "    AND NZ(KOUZM_KOKYC, '') = '' "
    Set rsKOUZM = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If rsKOUZM.EOF = False Then
        rsKOUZM.Edit
        rsKOUZM.Fields("KOUZM_KOKYC") = Format(cmb_顧客コード, "000000")
        rsKOUZM.Fields("KOUZM_UPDAD") = Format(Now(), "yyyymmdd")
        rsKOUZM.Fields("KOUZM_UPDAJ") = Format(Now(), "hhnnss")
        rsKOUZM.Fields("KOUZM_UPDPB") = Me.NAME
        rsKOUZM.Fields("KOUZM_UPDUB") = a作成ユーザ
        rsKOUZM.UPDATE
    End If
    'INSERT 2018/01/26 N.IMAI End
    
    insert_USER_MAST = True

    GoTo Finally

Exception:
    isException = True
Finally:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If isException Then
        Call Err.Raise(Err.Number, "insert_USER_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客情報変更理由_AfterUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客情報変更理由_AfterUpdate()
    
    With Me.cmb_顧客情報変更理由
        If .ListIndex >= 0 Then
            .VALUE = Format(.Text, .Format)
            Me.txt_顧客情報変更理由.VALUE = .Column(1, .ListIndex)
            ' 変更理由が"0：無し" OR 未入力ならば変更日は入力×
            If Nz(.VALUE, "00") = "00" Then
                Me.txt_顧客情報変更日.BackColor = 14680063
                Call CmLockControl(Me.txt_顧客情報変更日, 1)
                '変更日はクリアしてしまうよ
                Me.txt_顧客情報変更日.VALUE = Null
            Else
                Call CmLockControl(Me.txt_顧客情報変更日, 0)
                Me.txt_顧客情報変更日.BackColor = 16777215
                ' 変更日が未入力なら現在の年月日をデフォルト表示
                If Nz(Me.txt_顧客情報変更日.VALUE, "") = "" Then
                    Me.txt_顧客情報変更日.VALUE = Format(DATE, "yyyy/mm/dd")
                End If
            End If
        End If
    End With
    blnDataUpdate = True
    is顧客情報変更 = True
   
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客情報変更理由_BeforeUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客情報変更理由_BeforeUpdate(Cancel As Integer)

  If Me.cmb_顧客情報変更理由.ListIndex < 0 Then
       MsgBox "そのコードは登録されていません。", vbInformation, "確認"
       Cancel = True
       SendKeys "{ESC}"
  End If

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客情報変更理由_KeyPress
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客情報変更理由_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客情報変更理由, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日_AfterUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日_AfterUpdate()
    Me.txt_顧客情報変更日 = Format(var変更日付, "yyyy/mm/dd")
    blnDataUpdate = True
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日_BeforeUpdate
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日_BeforeUpdate(Cancel As Integer)
    Call isDateFormat(Me.txt_顧客情報変更日, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日_GotFocus
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日_GotFocus()
    txt_顧客情報変更日 = Format(txt_顧客情報変更日, "yy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日_KeyPress
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(txt_顧客情報変更日, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_顧客情報変更日_LostFocus
'        機能             :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客情報変更日_LostFocus()
    txt_顧客情報変更日 = Format(txt_顧客情報変更日, "yyyy/mm/dd")
End Sub
'==============================================================================*
'
'        MODULE_NAME      :isDateFormat
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :
'        OUT              :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isDateFormat(aTarget As Variant, ByRef Cancel As Integer)

    isDateFormat = False
    var変更日付 = CmDateChecker(aTarget)
    If IsNull(var変更日付) Then
        MsgBox "日付を正しく入力してください。", vbInformation, "エラー"
        Cancel = True
        SendKeys "{ESC}"
    Else
        isDateFormat = True
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :勤務先の入力可不可
'        MODULE_ID        :LockKinmusaki
'        CREATE_DATE      :2010/04/19 SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub LockKinmusaki(ByVal iMode As Integer)
  '引数 ０  使用可
  '     １  使用不可

    Call CmLockControl(Me.txt_USER_KNAME, iMode)
'    Call CmLockControl(Me.txt_USER_KKANA, iMode)                               'DELETE 2013/06/11 K.ISHIZAKA
    Call CmLockControl(Me.txt_USER_JOB, iMode)
    Call CmLockControl(Me.txt_USER_KPOST, iMode)
    Call CmLockControl(Me.txt_USER_KADR_1, iMode)
    Call CmLockControl(Me.txt_USER_KADR_2, iMode)
    Call CmLockControl(Me.txt_USER_KADR_3, iMode)
    Call CmLockControl(Me.txt_USER_KTEL, iMode)
    Call CmLockControl(Me.txt_USER_KHOUR_START, iMode)
    Call CmLockControl(Me.txt_USER_KHOUR_END, iMode)
    Call CmLockControl(Me.txt_USER_NENSHU, iMode)
    Call CmLockControl(Me.txt_USER_BIRTHDAY, iMode)

    cmd_KYubin.Enabled = (iMode = 0)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :データ保存
'        MODULE_ID        :setDataSave
'        CREATE_DATE      :2011/07/29 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setDataSave(ByRef rsData As Object)
    With rsData
        .Fields("USER_NAME") = Me.txt_顧客名
        .Fields("USER_KANA") = Me.txt_顧客名カナ
        .Fields("USER_TANM") = Me.txt_代表者名
        .Fields("USER_TAKA") = Me.txt_代表者名カナ
        .Fields("USER_YUBINO") = Me.txt_郵便番号
        .Fields("USER_ADR_1") = Me.txt_住所1
        .Fields("USER_ADR_2") = Me.txt_住所2
        .Fields("USER_ADR_3") = Me.txt_住所3
        .Fields("USER_TEL") = Trim(Me.txt_電話番号)
        .Fields("USER_KEITAI") = Trim(Me.txt_携帯番号)
        .Fields("USER_FAX") = Trim(Me.txt_FAX番号)
        .Fields("USER_SDAY") = Nz(CmNumeric(Me.txt_締日), 0)
        .Fields("USER_KKBN") = Nz(CmNumeric(Me.cmb_顧客区分), 0)
        .Fields("USER_SKBN") = Nz(CmNumeric(Me.cmb_集金区分), 0)
        .Fields("USER_SEIBETU") = Nz(CmNumeric(Me.cmb_性別区分), 0)         '2014/11/15 M.HONDA INS
        .Fields("USER_BIKO") = Me.txt_備考
        .Fields("USER_MAIL") = Me.txt_メールアドレス
        .Fields("USER_RNAME") = Me.txt_連絡先担当者名
        .Fields("USER_RKANA") = Me.txt_連絡先担当者名カナ
        .Fields("USER_RPOST") = Me.txt_連絡先郵便番号
        .Fields("USER_RADR_1") = Me.txt_連絡先都道府県市町村
        .Fields("USER_RADR_2") = Me.txt_連絡先番地
        .Fields("USER_RADR_3") = Me.txt_連絡先建物名等
        .Fields("USER_RTEL") = Trim(Me.txt_連絡先電話番号)
        .Fields("USER_RKEITAI") = Trim(Me.txt_連絡先携帯番号)
        .Fields("USER_RFAX") = Trim(Me.txt_連絡先FAX番号)
        .Fields("USER_UPDATE") = DATE
        .Fields("USER_CHG_RKBN") = Nz(CmNumeric(Me.cmb_顧客情報変更理由), 0)
        .Fields("USER_CHG_DATE") = Me.txt_顧客情報変更日
        
        If Nz(Me.cmb_新顧客コード, "") = "" Then
            .Fields("USER_NEW_CODE") = Null
        Else
            .Fields("USER_NEW_CODE") = CmNumeric(Me.cmb_新顧客コード)
        End If
        
        .Fields("USER_KNAME") = txt_USER_KNAME
        .Fields("USER_KKANA") = txt_USER_KKANA
        .Fields("USER_JOB") = txt_USER_JOB
        .Fields("USER_KPOST") = txt_USER_KPOST
        .Fields("USER_KADR_1") = txt_USER_KADR_1
        .Fields("USER_KADR_2") = txt_USER_KADR_2
        .Fields("USER_KADR_3") = txt_USER_KADR_3
        .Fields("USER_KTEL") = Trim(txt_USER_KTEL)
        .Fields("USER_KHOUR_START") = txt_USER_KHOUR_START
        .Fields("USER_KHOUR_END") = txt_USER_KHOUR_END
        .Fields("USER_NENSHU") = txt_USER_NENSHU
        
        If Trim(Nz(txt_USER_BIRTHDAY, "")) <> "" Then
            .Fields("USER_BIRTHDAY") = Replace(txt_USER_BIRTHDAY, "/", "")
        Else
            .Fields("USER_BIRTHDAY") = txt_USER_BIRTHDAY
        End If
        
    End With

End Sub
