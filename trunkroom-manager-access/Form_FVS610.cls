VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat of program ********************************

'*******************************************************************************
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'*******************************************************************************
'   プログラム名　　：　Web申込手動取込処理
'   プログラムＩＤ　：  Form_FVS600_S01
'   作　成　日　　　：  2019/04/10
'                       2019/04/23
'   作　成　者　　　：  TAKENOUCHI
'                       Y_SUZUKI
'   Ver             ：  1.0
'
'*******************************************************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'==============================================================================*
Private Type parentDispInfo
    strType         As String           '親画面から渡された申込トラン.申込種別
    intAgre         As Integer          '親画面から渡されたCARG_FILE.契約状態
    strKeiyakuNo    As String           '親画面から渡された契約No
    strLoginUser    As String           '親画面から渡されたログインユーザ
    strLoginPass    As String           '親画面から渡されたログインパスワード
    strKaiYtiYmd    As String           '親画面から渡された解約予定日
    strShinNitiji   As String           '親画面から渡された申込日
    strShinJi       As String           '親画面から渡された申込日時(時間のみ)
    strFormName     As String           '親画面のフォーム名
End Type

Private parentInfo As parentDispInfo

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID               As String = "FVS610"
Private Const FORM_NM_FVS600        As String = "FVS600"
Private Const FORM_NM_FVS601        As String = "FVS601"

Private Const KYK_KBN               As String = "01"            '解約区分
Private Const KYK_RIYUU_KBN         As String = "01"            '解約理由区分
Private Const KYK_UKE_TAN_CD        As String = "888"           '解約受付担当区分
Private Const TYPE_01               As String = "01"            '解約申込
Private Const TYPE_02               As String = "02"            '解約取消申込
Private Const AGRE_1                As Integer = 1              '契約中
Private Const AGRE_2                As Integer = 2              '解約予定
Private Const AGRE_3                As Integer = 3              '解約延長
Private Const AGRE_4                As Integer = 4              'メンテ中
Private Const AGRE_9                As Integer = 9              '完了
Private Const STS_CANCEL            As Integer = -1             '取消

Private Const MSG_PARTS_3           As String = "が、現契約では取込できませんのでネットの申込をキャンセルするため破棄して下さい"

'==============================================================================*
'
'       MODULE_NAME     : 「何もしない」ボタン押下時処理
'       MODULE_ID       : cmb_何もしない_Click()
'       CREATE_DATE     : 2019/04/10
'       PARAM           : なし
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_何もしない_Click()

    doCmd.Close acForm, Me.NAME

End Sub

'==============================================================================*
'
'       MODULE_NAME     : 「取込」ボタン押下時処理
'       MODULE_ID       : cmb_取込_Click()
'       CREATE_DATE     : 2019/04/10
'       PARAM           : なし
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_取込_Click()

    '親画面に解約日等を反映する
    Call subSetObjValParent
    MsgBox "契約入力画面へ申込の反映が完了しました" & vbCrLf & "契約入力画面で「保存」ボタンで取込完了です"
    doCmd.Close acForm, Me.NAME

End Sub

'==============================================================================*
'
'       MODULE_NAME     : 「破棄」ボタン押下時処理
'       MODULE_ID       : cmb_破棄_Click()
'       CREATE_DATE     : 2019/04/10
'       PARAM           : なし
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_破棄_Click()
    Dim strMsg      As String

    '確認メッセージ
    strMsg = "ネットの申込をキャンセルします。宜しいですか？" & vbCrLf & "お客様には別途連絡して下さい。"
    If vbCancel = MsgBox(strMsg, vbOKCancel, "申込みキャンセル確認") Then Exit Sub

    '申込トランの更新メイン処理
    If fncUpdEnterTrnMain(parentInfo.strKeiyakuNo, STS_CANCEL) = False Then
        MsgBox "申込トランの更新処理に失敗しました"
        Exit Sub
    End If
    
    MsgBox "ネットの申込を破棄しました。"
    doCmd.Close acForm, Me.NAME
    
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 画面起動時処理
'       MODULE_ID       : Form_Open
'       CREATE_DATE     : 2019/04/10
'       PARAM           : なし
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim iParamCnt   As Integer
    Dim strParam()  As String

    On Error GoTo Exception
    
    'OpenArgsをカンマごとに分解して配列に格納する
    strParam = Split(Nz(Me.OpenArgs), ",")

    ' パラメータ数チェック
    iParamCnt = UBound(strParam)

    If iParamCnt = 8 Then
        parentInfo.strType = strParam(0)                 '申込種別
        parentInfo.intAgre = strParam(1)                 '契約状態
        parentInfo.strKeiyakuNo = strParam(2)            '契約No
        parentInfo.strLoginUser = strParam(3)            'ログインユーザ
        parentInfo.strLoginPass = strParam(4)            'ログインパスワード
        parentInfo.strKaiYtiYmd = strParam(5)            '解約予定日
        parentInfo.strShinNitiji = strParam(6)           '申込日時
        parentInfo.strShinJi = strParam(7)              '申込日時(時)
        parentInfo.strFormName = strParam(8)            '親画面のフォーム名
    Else
        Cancel = True
        Call MSZZ024_M10("OpenArgs", "起動パラメータが正しくありません。")
        Exit Sub
    End If
    
'↓ADD 2019/04/23 Y_SUZUKI EGL
    'ボタン表示位置(横)の調整
    cmb_取込.Left = (Form.WIDTH / 4) - (cmb_取込.WIDTH / 2)
    cmb_破棄.Left = (Form.WIDTH / 2) - (cmb_破棄.WIDTH / 2)
    cmb_何もしない.Left = ((Form.WIDTH / 4) * 3) - (cmb_何もしない.WIDTH / 2)
    
    '解約申込(01)で、契約状態(4:メンテ中、9:完了)のデータ
    '解約取消申込(02)で、契約状態(1:契約中、4:メンテ中、9:完了)のデータ
    'の場合
    If (parentInfo.strType = TYPE_01 And (parentInfo.intAgre = AGRE_4 Or parentInfo.intAgre = AGRE_9)) Or _
       (parentInfo.strType = TYPE_02 And (parentInfo.intAgre = AGRE_1 Or parentInfo.intAgre = AGRE_4 Or parentInfo.intAgre = AGRE_9)) Then
        '画面項目の表示/非表示制御を行う。
        lbl_契約状況表示欄.Caption = lbl_契約状況表示欄.Caption & vbCrLf & MSG_PARTS_3
        lbl_お伺い.Visible = False
        cmb_取込.Visible = False
        
        'ボタン表示位置(横)の調整
        cmb_破棄.Left = (Form.WIDTH / 3) - (cmb_破棄.WIDTH / 2)
        cmb_何もしない.Left = ((Form.WIDTH / 3) * 2) - (cmb_何もしない.WIDTH / 2)
        
    End If
    
    '解約取消申込(02)の場合
    If parentInfo.strType = TYPE_02 Then
        '画面サイズ、項目位置の調整
        '解約予定日時を非表示にして
        lbl_解約予定日時_title.Visible = False
        lbl_解約予定日時.Visible = False
        
        '申込日時を解約予定日時の高さに持ってくる
        lbl_申込日時_title.Top = lbl_解約予定日時_title.Top
        lbl_申込日時.Top = lbl_解約予定日時_title.Top
        
    End If
    
    '解約予定日時が表示されている場合は値を設定
    If lbl_解約予定日時.Visible Then lbl_解約予定日時.Caption = parentInfo.strKaiYtiYmd
    
    '申込日時の値を設定
    lbl_申込日時.Caption = parentInfo.strShinNitiji & " " & parentInfo.strShinJi
    
'↑ADD 2019/04/23 Y_SUZUKI EGL

    Exit Sub

Exception:
    Call MSZZ024_M00("Form_Open", True)
    Cancel = True

End Sub

'==============================================================================*
'
'       MODULE_NAME     : 申込トランの更新メイン処理
'       MODULE_ID       : fncUpdEnterTrnMain
'       CREATE_DATE     : 2019/04/10
'       PARAM           : strKeiyakuNo  契約No
'                         intStatus     処理Status
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdEnterTrnMain(ByVal strKeiyakuNo As String, ByVal intStatus As Integer) As Boolean

    Dim objDoc              As Object
    Dim strErrMsg           As String
    Dim strLoginKey         As String
    Dim iniFile             As Collection
    Dim strUser             As String
    Dim strDate             As String
    Dim strTime             As String
    
    fncUpdEnterTrnMain = False
    
    On Error GoTo ErrorHandler
    
    '更新日時
    strUser = LsGetUserName()
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    
    'Config.iniを取得
    Set iniFile = MSZZ073K.get_ini_file()
    
    'キー取得
    strLoginKey = koms_login(parentInfo.strLoginUser, parentInfo.strLoginPass, iniFile)
    Set objDoc = koms_createDocument(parentInfo.strLoginUser, strLoginKey)
    
    '申込トランの更新準備
    Call upd_ENTER_TRAN_Documents(objDoc, strKeiyakuNo, intStatus, strUser, strDate, strTime)
    
    '申込トランの更新実行
    Call send_table(iniFile, objDoc, "ENTER_TRAN")
    
    fncUpdEnterTrnMain = True
    Exit Function
    
ErrorHandler:
    strErrMsg = Err.Description
    Call MSZZ024_M00("fncUpdEnterTrnMain", False)
    Call MSZZ003_M00(PROG_ID, "9", strErrMsg)

End Function

'==============================================================================*
'
'       MODULE_NAME     : ENTER_TRAN更新
'       MODULE_ID       : upd_ENTER_TRAN_Documents
'       CREATE_DATE     : 2019/04/10
'       PARAM           : objDoc                ドキュメントオブジェクト(I)
'                       : strUser               契約No
'                       : intStatus             処理Status
'                       : strUser               ユーザー(I)
'                       : strDate               日付(I)
'                       : strTime               時刻(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub upd_ENTER_TRAN_Documents( _
          objDoc As Object, _
    ByVal strKeiyakuNo As String, _
    ByVal intStatus As Integer, _
    ByVal strUser As String, _
    ByVal strDate As String, _
    ByVal strTime As String _
)
    
    Dim objTbl      As Object
    
    On Error GoTo ErrorHandler

    'アクションメソッド設定
    Set objTbl = koms_createTable(objDoc, "ENTER_TRAN", "UPDATE")
    
    Call koms_setColumn(objTbl, "STATUS", intStatus, "tinyint")
    Call koms_setColumn(objTbl, "UPDATE_DATE", strDate, "varchar")
    Call koms_setColumn(objTbl, "UPDATE_TIME", strTime, "varchar")
    Call koms_setColumn(objTbl, "UPDATE_PRG_ID", "FVS600", "varchar")
    Call koms_setColumn(objTbl, "UPDATE_USER_ID", strUser, "varchar")
    Call koms_whereColumn(objTbl, "VAR_01", strKeiyakuNo, "varchar")

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "upd_ENTER_TRAN_Documents" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'       MODULE_NAME     : tableの更新
'       MODULE_ID       : send_table
'       CREATE_DATE     : 2019/04/10
'       PARAM           : iniFile               接続情報等(I)
'                       : objDoc                ドキュメントオブジェクト(I)
'                       : strTbl                テーブル名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub send_table(iniFile As Collection, objDoc As Object, strTbl As String)
    
    Dim strResponseXml      As String
    Dim objRet              As Object
    Dim objTbl              As Object
    Dim strFile             As String
    
    On Error GoTo ErrorHandler
    
    If koms_table_ctrl(objDoc, strResponseXml, iniFile) Then
        Set objRet = XmlCreateDocument(strResponseXml)
        Set objTbl = objRet.selectSingleNode("/root/" & strTbl)
        
        If CInt(objTbl.Text) > 0 Then
            '正常
            Exit Sub
        End If
    End If
    
    Call MSZZ003_M80("次のコマンドに失敗しました。")
    Call MSZZ003_M80(objDoc.XML)
                                                                                
    If InStr(1, LCase(strResponseXml), LCase("charset=EUC-JP"), vbTextCompare) > 0 Then
        strFile = CurrentProject.Path & "\" & PROG_ID & "_" & Format(Now, "yyyymmdd_hhnnss") & "_ErrorLog.html"
        Call EUCSave(strFile, strResponseXml)
        strResponseXml = "結果は外部ファイル[" & strFile & "]を参照してください。"
    End If
                                                                                
    Call MSZZ003_M80(strResponseXml)
    Call MSZZ024_M10("koms_table_ctrl", "PHP側の" & strTbl & "の更新に失敗しました。")

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "send_table" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'       MODULE_NAME     : 解約日等を親画面に反映する
'       MODULE_ID       : subSetObjValParent
'       CREATE_DATE     : 2019/04/10
'       PARAM           : なし
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetObjValParent()
    Dim AnyForm As Form     'add
    
    On Error GoTo ErrorHandler
    
    '遷移元の親画面に値を反映する
    If parentInfo.strFormName = FORM_NM_FVS600 Then
        Forms(FORM_NM_FVS600).txt_申込種別.VALUE = parentInfo.strType
        
        '解約申込の場合
        If parentInfo.strType = TYPE_01 Then
            Forms(FORM_NM_FVS600).txt_解約日付.VALUE = parentInfo.strKaiYtiYmd
            Forms(FORM_NM_FVS600).txt_解約受付日.VALUE = parentInfo.strShinNitiji
            Forms(FORM_NM_FVS600).cmb_解約区分 = KYK_KBN
            Forms(FORM_NM_FVS600).cmb_解約理由 = KYK_RIYUU_KBN
            Forms(FORM_NM_FVS600).cmb_解約受付担当者コード = KYK_UKE_TAN_CD
        Else
            Forms(FORM_NM_FVS600).txt_解約日付.VALUE = ""
            Forms(FORM_NM_FVS600).txt_解約受付日.VALUE = ""
            Forms(FORM_NM_FVS600).cmb_解約区分 = KYK_KBN
            Forms(FORM_NM_FVS600).cmb_解約理由 = ""
            Forms(FORM_NM_FVS600).cmb_解約受付担当者コード = ""
        End If
        
        Forms(FORM_NM_FVS600).Call_AfterUpdate          'ADD 2019/04/23 Y_SUZUKI EGL
        
    ElseIf parentInfo.strFormName = FORM_NM_FVS601 Then
        Forms(FORM_NM_FVS601).txt_申込種別.VALUE = parentInfo.strType
        
        '解約申込の場合
        If parentInfo.strType = TYPE_01 Then
            Forms(FORM_NM_FVS601).txt_解約日付.VALUE = parentInfo.strKaiYtiYmd
            Forms(FORM_NM_FVS601).txt_解約受付日.VALUE = parentInfo.strShinNitiji
            Forms(FORM_NM_FVS601).cmb_解約区分 = KYK_KBN
            Forms(FORM_NM_FVS601).cmb_解約理由 = KYK_RIYUU_KBN
            Forms(FORM_NM_FVS601).cmb_解約受付担当者コード = KYK_UKE_TAN_CD
        Else
            Forms(FORM_NM_FVS601).txt_解約日付.VALUE = ""
            Forms(FORM_NM_FVS601).txt_解約受付日.VALUE = ""
            Forms(FORM_NM_FVS601).cmb_解約区分 = KYK_KBN
            Forms(FORM_NM_FVS601).cmb_解約理由 = ""
            Forms(FORM_NM_FVS601).cmb_解約受付担当者コード = ""
        End If
    
        Forms(FORM_NM_FVS601).Call_AfterUpdate          'ADD 2019/04/23 Y_SUZUKI EGL
        
    End If
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetObjValParent" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub
