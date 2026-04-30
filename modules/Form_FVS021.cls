VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　空状況検索
'   プログラムＩＤ　：　FVS021
'   作　成　日　　　：  2005/07/28
'   作　成　者　　　：  K.ISHIZAKA
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：  20050829
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.1
'   修　正　内　容　：  解約予定/取り置きを表示する場合、しない場合の対応
'**********************************************
'   修　正　日　　　：  2005/11/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.2
'   修　正　内　容　：  デザイン項目追加と色変更
'**********************************************
'   修　正　日　　　：  2005/12/05
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.3
'   修　正　内　容　：  コンテナマスタ・利用可否の追加と自社理由の対応
'**********************************************
'   修　正　日　　　：  2005/12/15
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  取り置きされているかの判断は、予約紹介トランを参照する
'**********************************************
'   修　正　日　　　：  2006/01/20
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.4
'   修　正　内　容　：  1.一時テーブルの制約を一時インデックスに変更する
'**********************************************
'   修　正　日　　　：  2006/01/22
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.5
'   修　正　内　容　：  1.自社理由のデータを非表示とする
'**********************************************
'   修　正　日　　　：  2006/08/31
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.9
'   修　正　内　容　：  ・ヤード非表示日を過ぎたヤードは選択リストから外す
'                       ・営業終了日を過ぎたヤードは検索対象外
'                       ・営業開始日を過ぎていなければ一覧上に表示する
'                       ・営業終了日が入力されていれば一覧上に表示する
'                       ・地主解約のチェックを外す
'                       ・コンテナ種別=0の条件をはずす
'                       ・備考欄に部屋と顧客の直近修繕を出力
'                       ・修繕入力に遷移可能にする
'**********************************************
'   修　正　日　　　：  2006/10/22
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.a
'   修　正　内　容　：  DAO→ADO変更
'**********************************************
'   修　正　日　　　：  2007/01/12
'   修　正　者　　　：  tajima
'   Ver             ：  0.b
'   修　正　内　容　：  備考のフォントを赤字に変更 ※プロパティのみ変更
'**********************************************
'   修　正　日　　　：  2007/02/07
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.c
'   修　正　内　容　：  修繕のデータがヌルのとき不正終了していた
'                   ：  営業開始日に達しなければ部屋指定予約は不可
'**********************************************
'   修　正　日　　　：  2007/04/14
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.d
'   修　正　内　容　：  Ver0.5の対応不具合
'**********************************************
'   修　正　日　　　：  2007/06/02
'   修　正　者　　　：  tajima
'   Ver             ：  0.c
'   修　正　内　容　：  レンタル用途が選ばれていればそれを予約入力へ渡す
'                       ※部屋指定ではなくヤード予約の場合のみ
'**********************************************
'   修　正　日　　　：  2009/06/13
'   修　正　者　　　：  tajima
'   Ver             ：  0.d
'   修　正　内　容　：  詳細ページ（ホームページ）の拡張子変更
'
'**********************************************
'   修　正　日　　　：  2010/06/10
'   修　正　者　　　：  SHIBAZAKI
'   Ver             ：  1.0
'   修　正　内　容　：  直近の顧客問合せを取得する際、表示区分が「表示」のデータのみを対象とする。
'**********************************************
'   修　正　日　　　：  2010/07/25
'   修　正　者　　　：  M.RYU
'   Ver             ：  1.1
'   修　正　内　容　：  3535ホームページ切替により、詳細ページのリンクも変わります。
'**********************************************
'   修　正　日　　　：  2019/04/22
'   修　正　者　　　：  N.IMAI
'   Ver             ：  1.2
'   修　正　内　容　：  KOTO_TRANから取得した文言がオーバーフローする問題に対応
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FVS021"
Private Const constMsgKaku = "確認"

Private Const C_YARD_CODE       As Integer = 0  'ヤードコード
Private Const C_CNTA_SYU        As Integer = 1  'コンテナ種別条件
Private Const C_CONT_BUMOC      As Integer = 2  '部門コード
Private Const C_CONT_DISKBACK   As Integer = 3  'WEBパス
Private Const C_CONT_EXCEL_PATH As Integer = 4  'EXCELパス
Private Const C_CONNECTION      As Integer = 5  'SQL-Server接続文字列
Private Const C_USER_ID         As Integer = 6  'ユーザー名

Private Const adOpenKeyset = 1
Private Const adLockPessimistic = 2
Private Const adCmdText = 1
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1

'ADO DataType
Private Const adNumeric = 131
Private Const adVarChar = 200
Private Const adDBTimeStamp = 135

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private strOpenArgs()       As String       '引数配列
Private lngYardCode         As Long         'ヤードコード
Private strCaption          As String       'タイトル
Private bChangeFlag         As Boolean      '編集中フラグ

'==============================================================================*
'   イベント関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : WEBボタンが押されたとき
'       MODULE_ID       : cmd_WEB_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_WEB_Click()
    On Error GoTo Err_Handle
    
    Call MSZZ009_M00(Me.cmd_WEB, Me.txt_SubAddress, Me.txt_Address)
    Exit Sub
    
Err_Handle:
    MsgBox "cmd_WEB_Click:" & Err.Number & vbCrLf & Err.Description
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 取消ボタン押下時処理
'       MODULE_ID       : cmd_取消_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
    bChangeFlag = False
    Call Reload(Format(lngYardCode), "")
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 予約受付ボタン押下時処理
'       MODULE_ID       : cmd_受付_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_受付_Click()
    Call ShowFVS500("", "", Nz(Form_FVS011.cmb_Usage.VALUE, ""))  'fix 2007/06/02 tajima
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 終了ボタンが押されたとき
'       MODULE_ID       : cmd_終了_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next
    'Call MSHS010_M00(PROG_ID, "2")
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 配置ボタンが押されたとき
'       MODULE_ID       : cmd_配置_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_配置_Click()
    Dim strExcel            As String
    On Error GoTo Err_Handle

    strExcel = Dir(strOpenArgs(C_CONT_EXCEL_PATH) & strOpenArgs(C_YARD_CODE) & "*.XLS")
    If strExcel <> "" Then
        Call Y_Exec(Me.hwnd, strOpenArgs(C_CONT_EXCEL_PATH) & strExcel, "", "", SW_SHOWNORMAL)
    Else
        MsgBox "ｺﾝﾃﾅ配置EXCELが存在しません。", vbOK, constMsgKaku
    End If
    Exit Sub
    
Err_Handle:
    MsgBox "cmd_配置_Click:" & Err.Number & vbCrLf & Err.Description
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 保存ボタンが押されたとき
'       MODULE_ID       : cmd_保存_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
    Call fncSaveData
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームアクティブ時処理
'       MODULE_ID       : Form_Activate
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    doCmd.Maximize
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームクローズ時処理
'       MODULE_ID       : Form_Close
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    Call subDropWorkTable(CurrentDb())
    'Call MSHS010_M00(PROG_ID, "2")
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームデアクティブ時処理
'       MODULE_ID       : Form_Deactivate
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Deactivate()
    On Error Resume Next
    '編集レコードをワークに保存
    doCmd.RunCommand acCmdSaveRecord
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームエラー時処理
'       MODULE_ID       : Form_Error
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : DataErr
'                       : Response
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
    Case 2169       '「レコードは保存されません。」という時
        DataErr = 0
        Response = 0
    End Select
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームキー押下時処理
'       MODULE_ID       : Form_KeyDown
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeyDown(KeyCode, Shift)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : フォームロード時処理
'       MODULE_ID       : Form_Load
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    On Error GoTo Err_Handle
    
    'Call MSHS010_M00(PROG_ID, "1")
    strCaption = Me.Caption
    '引数は ヤードコード、コンテナ種別条件、部門コード、WEBパス、EXCELパス、SQL-Server接続文字列、ユーザー名
    strOpenArgs = Split(Me.OpenArgs, ",")
    Me.txt_YARD_CODE = strOpenArgs(C_YARD_CODE)
    lngYardCode = Val(strOpenArgs(C_YARD_CODE))
    Call subSelFVS011
    Me.cmd_受付.SetFocus
    If strOpenArgs(C_CONT_DISKBACK) = "" Then
        Me.cmd_WEB.Enabled = False
    Else
'        If strOpenArgs(C_CONT_BUMOC) = "8" Then                               'DELETE 20090613 N.Miura START
'            Me.txt_Address = strOpenArgs(C_CONT_DISKBACK) & Format(lngYardCode) & ".htm"
'        Else
'            Me.txt_Address = strOpenArgs(C_CONT_DISKBACK) & strOpenArgs(C_YARD_CODE) & ".htm"
'        End If                                                                'DELETE 20090613 N.Miura ENDED
        
        Select Case strOpenArgs(C_CONT_BUMOC)                                  'INSERT 20090613 N.Miura START
        Case "1":  Me.txt_Address = strOpenArgs(C_CONT_DISKBACK) & strOpenArgs(C_YARD_CODE) & ".php"
'        Case "8":  Me.txt_Address = strOpenArgs(C_CONT_DISKBACK) & Format(lngYardCode) & ".htm"        'DELETE 2011/07/25 M.RYU
'        Case Else: Me.txt_Address = strOpenArgs(C_CONT_DISKBACK) & strOpenArgs(C_YARD_CODE) & ".htm"   'DELETE 2011/07/25 M.RYU
        Case "8":  Me.txt_Address = strOpenArgs(C_CONT_DISKBACK) & Format(lngYardCode) & ".html"        'INSERT 2011/07/25 M.RYU
        Case Else: Me.txt_Address = strOpenArgs(C_CONT_DISKBACK) & strOpenArgs(C_YARD_CODE) & ".html"   'INSERT 2011/07/25 M.RYU
        End Select                                                             'INSERT 20090613 N.Miura ENDED
    End If
    If strOpenArgs(C_CONT_EXCEL_PATH) = "" Then
        Me.cmd_配置.Enabled = False
    End If
    Call fncGetData
    
    Exit Sub
    
Err_Handle:
    MsgBox "Form_Load:" & Err.Number & vbCrLf & Err.Description
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ダブルクリック時処理
'       MODULE_ID       : txt_YARD_CODE_DblClick
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_YARD_CODE_DblClick(Cancel As Integer)
    Call Form_KeyDown(CmReprMod.P_呼出_ヤード, 0)
End Sub

'==============================================================================*
'   内部関数（イベントでない）
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'==============================================================================*
'
'       MODULE_NAME     : ヤード情報表示
'       MODULE_ID       : subSelFVS011
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSelFVS011()
    Dim rsAccess            As Recordset
    
    Set rsAccess = CurrentDb.OpenRecordset(fncSelectFVS011(), dbOpenForwardOnly, dbReadOnly)
    With rsAccess
        Me.txt_YARD_NAME = Nz(.Fields("YARD_NAME"))
        Me.txt_YARD_ADDR = Nz(.Fields("YARD_ADDR_1")) & Nz(.Fields("YARD_ADDR_2")) & Nz(.Fields("YARD_ADDR_3"))
        Me.txt_YARD_NOTE = Nz(.Fields("YARD_NOTE"))
        Me.txt_YARD_BIKON = Nz(.Fields("YARD_BIKON"))
        Me.txt_YARD_SEV1N = Nz(.Fields("YARD_SEV1N"))
        Me.txt_YARD_SEV2N = Nz(.Fields("YARD_SEV2N"))
        Me.txt_YARD_SEV3N = Nz(.Fields("YARD_SEV3N"))
        Me.txt_YARD_ENDEN = Nz(.Fields("YARD_ENDEN"))
        Me.txt_KAIYAKU = Nz(.Fields("KAIYAKU"))                                 'INSERT 20050829 K.ISHIZAKA
        Me.txt_TORIOKI = Nz(.Fields("TORIOKI"))                                 'INSERT 20050829 K.ISHIZAKA
        Me.txt_MIKAKUTEI = Nz(.Fields("MIKAKUTEI"))
        Me.txt_YARD_BEGIN_DAY = IIf(IsNull(.Fields("YARD_BEGIN_DAY")), "", Format(.Fields("YARD_BEGIN_DAY"), "yyyy/mm/dd"))         'INSERT 2006/08/31 K.ISHIZAKA
        Me.txt_YARD_RENTEND_DAY = IIf(IsNull(.Fields("YARD_RENTEND_DAY")), "", Format(.Fields("YARD_RENTEND_DAY"), "yyyy/mm/dd"))   'INSERT 2006/08/31 K.ISHIZAKA
        .Close
    End With
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 作業用テーブル取得ＳＱＬ作成
'       MODULE_ID       : fncCreateTempTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectFVS011() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_ADDR_1,"
    strSQL = strSQL & " YARD_ADDR_2,"
    strSQL = strSQL & " YARD_ADDR_3,"
    strSQL = strSQL & " YARD_NOTE,"
    strSQL = strSQL & " YARD_BIKON,"
    strSQL = strSQL & " YARD_SEV1N,"
    strSQL = strSQL & " YARD_SEV2N,"
    strSQL = strSQL & " YARD_SEV3N,"
    strSQL = strSQL & " YARD_ENDEN,"
    strSQL = strSQL & " YARD_BEGIN_DAY,"                                        'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " YARD_RENTEND_DAY,"                                      'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " KAIYAKU,"                                               'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & " TORIOKI,"                                               'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & " MIKAKUTEI "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " WK_FVS011_TEMP "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & " YARD_CODE = " & Format(lngYardCode)

    fncSelectFVS011 = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルの削除
'       MODULE_ID       : subDropWorkTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : dbAccess          Access側データベース
'                       : strTable          テーブル名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDropWorkTable(ByVal dbAccess As Database)
    On Error Resume Next

    Me.サブ下.SourceObject = ""
    Me.サブ上.SourceObject = ""
    dbAccess.TableDefs.Delete "WK_FVS021_TEMP"
    Select Case Err.Number
    Case 0
    Case 3265
        Err.Clear
    Case Else
        MsgBox "subDropWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
        Err.Clear
    End Select
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用テーブル作成ＳＱＬ作成
'       MODULE_ID       : fncCreateTempTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateTempTable() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "CREATE TABLE #WK_FVS021_TEMP ("
    strSQL = strSQL & " CNTA_CODE       NUMERIC(6)  NOT NULL,"
    strSQL = strSQL & " CNTA_NO         NUMERIC(6),"
    strSQL = strSQL & " CNTA_STEP       NUMERIC(1),"
    strSQL = strSQL & " CNTA_SYU_NAME   VARCHAR(20),"
    strSQL = strSQL & " CNTA_SIZE       NUMERIC(5,2),"                          'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_PRICE_DIFF NUMERIC(9),"                            'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_REASON     VARCHAR(60),"                           'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_HEIGHT     NUMERIC(5,2),"                          'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_WIDTH      NUMERIC(5,2),"                          'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_DEPTH      NUMERIC(5,2),"                          'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " YOUKT_YUKBN     NUMERIC(2),"
    strSQL = strSQL & " YOUKT_YUKBN_NAME    VARCHAR(20),"
    strSQL = strSQL & " YOUKT_TKDATE    DATETIME,"
    strSQL = strSQL & " PRIC_PRICA      NUMERIC(9),"
    strSQL = strSQL & " PRIC_FZAPPI     NUMERIC(9),"                            'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " PRIC_EZAPPI     NUMERIC(9),"                            'INSERT 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & " KAKUTEI         NUMERIC(6),"                           'DELETE 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " KAKUTEI         VARCHAR(20),"                           'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " YOUKT_UCODE     NUMERIC(6),"                            'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " YOUKT_NAME      VARCHAR(36),"
'    strSql = strSql & " YOUKT_RENKBN    VARCHAR(60),"                          'DELETE 2019/04/22 N.IMAI
    strSQL = strSQL & " YOUKT_RENKBN    VARCHAR(120),"                          'INSERT 2019/04/22 N.IMAI
    strSQL = strSQL & " YOUKT_BIKO      VARCHAR(60),"
    strSQL = strSQL & " YOUKT_UKNO      VARCHAR(11),"
    strSQL = strSQL & " CHANGE          NUMERIC(1) "
'>>> DELETE START 2006/01/20 K.ISHIZAKA >>>
'    strSQL = strSQL & ",CONSTRAINT WK_KEY0B PRIMARY KEY"                        'INSERT 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & "     ("                                                  'INSERT 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & "     CNTA_CODE, CNTA_NO"                                 'INSERT 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & "     )"                                                  'INSERT 2005/11/17 K.ISHIZAKA
'<<< DELETE END   2006/01/20 K.ISHIZAKA <<<
    strSQL = strSQL & ")"
    fncCreateTempTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用テーブル作成ＳＱＬ作成
'       MODULE_ID       : fncCreateTempIndex
'       CREATE_DATE     : 2006/01/20        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateTempIndex() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "CREATE INDEX #WK_FVS021_KEY0B ON #WK_FVS021_TEMP ("
    strSQL = strSQL & "     CNTA_CODE, CNTA_NO"
    strSQL = strSQL & ")"
    
    fncCreateTempIndex = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ作成ＳＱＬ作成
'       MODULE_ID       : fncInsertTempTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInsertTempTable() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "INSERT"
    strSQL = strSQL & " #WK_FVS021_TEMP "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_NO,"
    strSQL = strSQL & " CNTA_STEP,"
    strSQL = strSQL & "(SELECT NAME_NAME"
    strSQL = strSQL & " FROM   NAME_MAST"
'    strSQL = strSQL & " WHERE  NAME_ID   = '040'"                              'DELETE 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & " AND    NAME_CODE = CNTA_SYU"                           'DELETE 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " WHERE  NAME_ID   = '086'"                               'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " AND    NAME_CODE = CNTA_USAGE"                          'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & ")            AS CNTA_SYU_NAME,"
    strSQL = strSQL & " CNTA_SIZE,"                                             'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_PRICE_DIFF,"                                       'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_REASON,"                                           'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_HEIGHT,"                                           'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_WIDTH,"                                            'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " CNTA_DEPTH,"                                            'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " NULL        AS YOUKT_YUKBN,"
    strSQL = strSQL & " '空き'      AS YOUKT_YUKBN_NAME,"
    strSQL = strSQL & " NULL        AS YOUKT_TKDATE,"
'    strSQL = strSQL & "(SELECT PRICA"                                          'DELETE 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & " FROM   PRIC_TABL"                                      'DELETE 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & " WHERE  YARDC = CNTA_CODE"                              'DELETE 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & " AND    SAIZC = CNTA_SYU"                               'DELETE 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & " AND    SETKI = CNTA_STEP"                              'DELETE 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & ")            AS PRIC_PRICA,"                            'DELETE 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " PRIC_PRICE,"                                            'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " PRIC_FZAPPI,"                                           'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " PRIC_EZAPPI,"                                           'INSERT 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & "(SELECT COUNT(*)"                                       'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & " FROM   YOUK_TRAN"                                      'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & " WHERE  YOUKT_YCODE = CNTA_CODE"                        'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & " AND    YOUKT_NO    = CNTA_NO"                          'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & " AND    YOUKT_YUKBN = 1"                                'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & ")            AS KAKUTEI,"                               'DELETE 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " NULL        AS KAKUTEI,"                                'INSERT 2005/12/05 K.ISHIZAKA

    strSQL = strSQL & " NULL        AS YOUKT_UCODE,"                            'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " NULL        AS YOUKT_NAME,"
'    strSQL = strSQL & " CNTA_YTEL   AS YOUKT_RENKBN,"                          'DELETE 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " NULL        AS YOUKT_RENKBN,"                           'INSERT 2006/08/31 K.ISHIZAKA
'    strSQL = strSQL & " CNTA_YNAME  AS YOUKT_BIKO,"                            'DELETE 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " NULL        AS YOUKT_BIKO,"                             'INSERT 2006/08/31 K.ISHIZAKA
    
    strSQL = strSQL & " NULL        AS YOUKT_UKNO,"
    strSQL = strSQL & " 0           AS CHANGE "
    strSQL = strSQL & " FROM   CNTA_MAST"
    strSQL = strSQL & " LEFT OUTER JOIN"                                        'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "        PRIC_TABL"                                       'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " ON (   PRIC_YCODE = CNTA_CODE"                          'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " AND    PRIC_STEP  = CNTA_STEP"                          'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " AND    PRIC_USAGE = CNTA_USAGE"                         'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " AND    PRIC_SIZE  = CNTA_SIZE )"                        'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & " WHERE  CNTA_CODE  = " & Format(lngYardCode) & strOpenArgs(C_CNTA_SYU)
    strSQL = strSQL & " AND    CNTA_USE   = 1"                                  'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " AND    NOT EXISTS"
    strSQL = strSQL & "    (SELECT *"
    strSQL = strSQL & "     FROM   CARG_FILE"
    strSQL = strSQL & "     WHERE  CARG_YCODE  =  CNTA_CODE"
    strSQL = strSQL & "     AND    CARG_NO     =  CNTA_NO"
    
    If Val(Nz(Me.txt_KAIYAKU)) > 0 Then                                         '復活   2006/01/22 K.ISHIZAKA   'DELETE 2005/12/05 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'    If (Nz(Me.txt_TORIOKI) <> "") Or (Nz(Me.txt_KAIYAKU) <> "") Then           'DELETE 2006/01/22 K.ISHIZAKA   'INSERT 2005/12/15 K.ISHIZAKA
        strSQL = strSQL & " AND    CARG_KYDATE IS NULL )"
    Else                                                                        'INSERT 2005/12/15 K.ISHIZAKA   'DELETE 2005/12/05 K.ISHIZAKA
        strSQL = strSQL & " AND    CARG_CONTNA = 0 )"                           'INSERT 2005/12/15 K.ISHIZAKA   'DELETE 2005/12/05 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
    End If                                                                      'INSERT 2005/12/15 K.ISHIZAKA   'DELETE 2005/12/05 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & " AND    NOT EXISTS"
    strSQL = strSQL & "    (SELECT *"
'    strSQL = strSQL & "     FROM   YOUK_TRAN"                                  'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "     WHERE  YOUKT_YCODE =  CNTA_CODE"                   'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "     AND    YOUKT_NO    =  CNTA_NO"                     'DELETE 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "     FROM   INTR_TRAN "                                  'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "     INNER  JOIN"                                        'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "            YOUK_TRAN "                                  'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "     ON (   YOUKT_UKNO = INTRT_UKNO "                    'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "     )"                                                  'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "     WHERE  INTRT_YCODE =  CNTA_CODE"                    'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "     AND    INTRT_NO    =  CNTA_NO"                      'INSERT 2005/12/15 K.ISHIZAKA
    If Val(Nz(Me.txt_TORIOKI)) > 0 Then                                         '復活   2006/01/22 K.ISHIZAKA   'DELETE 2005/12/15 K.ISHIZAKA   'INSERT 20050829 K.ISHIZAKA
'    If (Nz(Me.txt_TORIOKI) <> "") Or (Nz(Me.txt_KAIYAKU) <> "") Then           'DELETE 2006/01/22 K.ISHIZAKA   'INSERT 2005/12/15 K.ISHIZAKA
'        strSQL = strSQL & " AND    YOUKT_YUKBN =  10)"                         'DELETE 2005/12/15 K.ISHIZAKA
'        strSQL = strSQL & " AND    INTRT_INTROKBN = 2)"                        'DELETE 2006/01/22 K.ISHIZAKA   'INSERT 2005/12/15 K.ISHIZAKA
'        strSQL = strSQL & " AND    YOUKT_YUKBN IN(1, 10))"                     'DELETE 2007/04/14 K.ISHIZAKA 'INSERT 2006/01/22 K.ISHIZAKA   'DELETE 2005/12/15 K.ISHIZAKA
        strSQL = strSQL & " AND  ( INTRT_INTROKBN = 2 "                         'INSERT 2007/04/14 K.ISHIZAKA
        strSQL = strSQL & " OR     YOUKT_YUKBN    = 1 ))"                       'INSERT 2007/04/14 K.ISHIZAKA
    Else                                                                        'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " AND    YOUKT_YUKBN IN(2, 10))"                     'DELETE 2005/12/15 K.ISHIZAKA   'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " AND    INTRT_INTROKBN IN(1, 2))"                   'DELETE 2006/01/22 K.ISHIZAKA   'INSERT 2005/12/15 K.ISHIZAKA
'        strSQL = strSQL & " AND    YOUKT_YUKBN IN(1, 2, 10))"                  'DELETE 2007/04/14 K.ISHIZAKA 'INSERT 2006/01/22 K.ISHIZAKA   'DELETE 2005/12/15 K.ISHIZAKA   'INSERT 20050829 K.ISHIZAKA
        strSQL = strSQL & " AND    INTRT_INTROKBN IN(1, 2))"                    'INSERT 2007/04/14 K.ISHIZAKA
    End If                                                                      'INSERT 20050829 K.ISHIZAKA
    
    fncInsertTempTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ更新ＳＱＬ作成（解約予定データ）
'       MODULE_ID       : fncUpdateTempTable_20
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateTempTable_20() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & " #WK_FVS021_TEMP "
    strSQL = strSQL & "SET"

    strSQL = strSQL & " YOUKT_YUKBN      = NAME_CODE,"
    strSQL = strSQL & " YOUKT_YUKBN_NAME = NAME_NAME,"
    strSQL = strSQL & " YOUKT_TKDATE     = CARG_KYDATE,"
'    strSQL = strSQL & " YOUKT_RENKBN     = CNTA_YTEL,"                         'DELETE 2006/08/31 K.ISHIZAKA 'INSERT 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & " YOUKT_BIKO       = CNTA_YNAME,"                        'DELETE 2006/08/31 K.ISHIZAKA 'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " YOUKT_UCODE      = CARG_UCODE,"                         'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " YOUKT_NAME       = USER_NAME "
    
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "     #WK_FVS021_TEMP "
    
    strSQL = strSQL & "INNER JOIN"                                              'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "     CNTA_MAST "                                         'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "ON ( CNTA_MAST.CNTA_CODE = #WK_FVS021_TEMP.CNTA_CODE "   'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "AND  CNTA_MAST.CNTA_NO   = #WK_FVS021_TEMP.CNTA_NO  )"   'INSERT 2005/12/05 K.ISHIZAKA
    
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "     CARG_FILE "
'    strSQL = strSQL & "ON ( CARG_YCODE  =  CNTA_CODE "                         'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & "AND  CARG_NO     =  CNTA_NO  )"                         'DELETE 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "ON ( CARG_YCODE  =  #WK_FVS021_TEMP.CNTA_CODE "          'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "AND  CARG_NO     =  #WK_FVS021_TEMP.CNTA_NO  )"          'INSERT 2005/12/05 K.ISHIZAKA
    
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "     USER_MAST "
    strSQL = strSQL & "ON ( USER_CODE   =  CARG_UCODE )"
    
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "     NAME_MAST "
    strSQL = strSQL & "ON ( NAME_ID   = '080' "
    strSQL = strSQL & "AND  NAME_CODE = 20 )"
    
    strSQL = strSQL & "WHERE CARG_CONTNA =  0 "
    strSQL = strSQL & "AND   CARG_KYDATE IS NOT NULL "
    
    fncUpdateTempTable_20 = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ更新ＳＱＬ作成（予約取置きデータ）
'       MODULE_ID       : fncUpdateTempTable_02
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateTempTable_02() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & " #WK_FVS021_TEMP "
    strSQL = strSQL & "SET"

    strSQL = strSQL & " #WK_FVS021_TEMP.YOUKT_YUKBN  = YOUK_TRAN.YOUKT_YUKBN,"
    strSQL = strSQL & " YOUKT_YUKBN_NAME             = NAME_YUKBN.NAME_NAME,"
    strSQL = strSQL & " KAKUTEI ="                                              'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " CASE WHEN YOUK_TRAN.YOUKT_YUKBN = 1"                    'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " THEN"                                                   'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "     NAME_KAKUTEI.NAME_NAME"                             'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " ELSE"                                                   'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "     NULL"                                               'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " END,"                                                   'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " #WK_FVS021_TEMP.YOUKT_TKDATE = YOUK_TRAN.YOUKT_TKDATE,"
    strSQL = strSQL & " #WK_FVS021_TEMP.YOUKT_UCODE  = YOUK_TRAN.YOUKT_UCODE,"  'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " #WK_FVS021_TEMP.YOUKT_NAME   = YOUK_TRAN.YOUKT_NAME,"
    strSQL = strSQL & " #WK_FVS021_TEMP.YOUKT_RENKBN = NAME_RENKBN.NAME_NAME + ':' +"
    strSQL = strSQL & " CASE YOUK_TRAN.YOUKT_RENKBN"
    strSQL = strSQL & " WHEN  0 THEN"
    strSQL = strSQL & "     YOUK_TRAN.YOUKT_TEL"
    strSQL = strSQL & " WHEN  1 THEN"
    strSQL = strSQL & "     YOUK_TRAN.YOUKT_KEITAI"
    strSQL = strSQL & " WHEN  2 THEN"
    strSQL = strSQL & "     YOUK_TRAN.YOUKT_FAX"
    strSQL = strSQL & " WHEN 10 THEN"
    strSQL = strSQL & "     YOUK_TRAN.YOUKT_MAIL"
    strSQL = strSQL & " ELSE"
    strSQL = strSQL & "     NULL"
    strSQL = strSQL & " END,"
'    strSQL = strSQL & " #WK_FVS021_TEMP.YOUKT_BIKO   = YOUK_TRAN.YOUKT_BIKO,"  'DELETE 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " #WK_FVS021_TEMP.YOUKT_UKNO   = YOUK_TRAN.YOUKT_UKNO "
    
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "     #WK_FVS021_TEMP "
    
    strSQL = strSQL & "INNER JOIN"                                              'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "     INTR_TRAN "                                         'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "ON ( INTRT_YCODE  =  CNTA_CODE "                         'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "AND  INTRT_NO     =  CNTA_NO  )"                         'INSERT 2005/12/15 K.ISHIZAKA
    
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "     YOUK_TRAN "
'    strSQL = strSQL & "ON ( YOUKT_YCODE  =  CNTA_CODE "                        'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "AND  YOUKT_NO     =  CNTA_NO  )"                        'DELETE 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "ON ( YOUK_TRAN.YOUKT_UKNO = INTRT_UKNO )"                'INSERT 2005/12/15 K.ISHIZAKA
    
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "     NAME_MAST AS NAME_YUKBN "
    strSQL = strSQL & "ON ( NAME_YUKBN.NAME_ID   = '080' "
'    strSQL = strSQL & "AND  NAME_YUKBN.NAME_CODE = YOUK_TRAN.YOUKT_YUKBN )"    'DELETE 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "AND  NAME_YUKBN.NAME_CODE = 2 )"                         'INSERT 2005/12/05 K.ISHIZAKA

    strSQL = strSQL & "INNER JOIN"                                              'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "     NAME_MAST AS NAME_KAKUTEI "                         'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "ON ( NAME_KAKUTEI.NAME_ID   = '080' "                    'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "AND  NAME_KAKUTEI.NAME_CODE = 1 )"                       'INSERT 2005/12/05 K.ISHIZAKA

    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "     NAME_MAST AS NAME_RENKBN "
    strSQL = strSQL & "ON ( NAME_RENKBN.NAME_ID   = '085' "
    strSQL = strSQL & "AND  NAME_RENKBN.NAME_CODE = YOUK_TRAN.YOUKT_RENKBN )"

'    strSQL = strSQL & "WHERE YOUK_TRAN.YOUKT_YUKBN = 2 "                       'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & "WHERE YOUK_TRAN.YOUKT_YUKBN IN(1, 2) "                  'DELETE 2005/12/15 K.ISHIZAKA   'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "WHERE INTRT_INTROKBN = 1 "                               'INSERT 2005/12/15 K.ISHIZAKA
    
    fncUpdateTempTable_02 = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルの作成
'       MODULE_ID       : fncCreateWorkTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : dbAccess          Access側データベース
'                       : rsSqlserver       SQL-Server側レコードセット
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable(ByVal dbAccess As Database, ByVal rsSqlServer As Object) As Boolean
    Dim tdfNew              As TableDef
    Dim i                   As Integer
    On Error GoTo Err_Handler
    
    Set tdfNew = dbAccess.CreateTableDef("WK_FVS021_TEMP")
    For i = 0 To rsSqlServer.Fields.Count - 1
        With rsSqlServer.Fields(i)
            Select Case .Type
            Case adVarChar
                tdfNew.Fields.Append tdfNew.CreateField(.NAME, dbText, .DefinedSize)
                tdfNew.Fields(.NAME).AllowZeroLength = True
            Case adDBTimeStamp
                tdfNew.Fields.Append tdfNew.CreateField(.NAME, dbDate)
            Case Else
                If .NumericScale > 0 Then                                       'INSERT 2005/11/17 K.ISHIZAKA
                    tdfNew.Fields.Append tdfNew.CreateField(.NAME, dbDouble)    'INSERT 2005/11/17 K.ISHIZAKA
                Else                                                            'INSERT 2005/11/17 K.ISHIZAKA
                    tdfNew.Fields.Append tdfNew.CreateField(.NAME, dbLong)
                End If                                                          'INSERT 2005/11/17 K.ISHIZAKA
            End Select
        End With
    Next
    dbAccess.TableDefs.Append tdfNew
    Me.サブ上.SourceObject = "FVS021_S03"
    Me.サブ下.SourceObject = "FVS021_S04"
    Me.Refresh
    fncCreateWorkTable = True
    Exit Function

Err_Handler:
    MsgBox "fncCreateWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
    Me.サブ上.SourceObject = "FVS021_S03"
    Me.サブ下.SourceObject = "FVS021_S04"
    Me.Refresh
    fncCreateWorkTable = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ取得ＳＱＬ作成
'       MODULE_ID       : fncSelectTempTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectTempTable() As String
    fncSelectTempTable = "SELECT * FROM #WK_FVS021_TEMP"
    If (Nz(Me.txt_TORIOKI) = "") And (Nz(Me.txt_KAIYAKU) = "") Then             'INSERT 2005/12/05 K.ISHIZAKA
        fncSelectTempTable = fncSelectTempTable & " WHERE YOUKT_YUKBN IS NULL"  'INSERT 2005/12/05 K.ISHIZAKA
    ElseIf Nz(Me.txt_TORIOKI) = "" Then                                         'INSERT 2005/12/05 K.ISHIZAKA
        fncSelectTempTable = fncSelectTempTable & " WHERE ISNULL(YOUKT_YUKBN, 20) = 20" 'INSERT 2005/12/05 K.ISHIZAKA
    ElseIf Nz(Me.txt_KAIYAKU) = "" Then                                         'INSERT 2005/12/05 K.ISHIZAKA
        fncSelectTempTable = fncSelectTempTable & " WHERE YOUKT_YUKBN IS NULL"  'INSERT 2005/12/05 K.ISHIZAKA
        fncSelectTempTable = fncSelectTempTable & " OR YOUKT_UKNO IS NOT NULL"  'INSERT 2005/12/05 K.ISHIZAKA
    End If                                                                      'INSERT 2005/12/05 K.ISHIZAKA
    fncSelectTempTable = fncSelectTempTable & " ORDER BY CNTA_NO"               'INSERT 2006/08/31 K.ISHIZAKA
End Function

'==============================================================================*
'
'       MODULE_NAME     : 詳細データ検索処理
'       MODULE_ID       : fncGetData
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetData() As Boolean
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim dbAccess            As Database
    Dim rsAccess            As Recordset
    Dim i                   As Integer
'    Dim dbKase              As Database                                        'DELETE 2006/10/22 K.ISHIZAKA 'INSERT 2006/08/31 K.ISHIZAKA
'    Dim rsKase              As Recordset                                       'DELETE 2006/10/22 K.ISHIZAKA  'INSERT 2006/08/31 K.ISHIZAKA
'    Dim strDNS              As String                                          'DELETE 2006/10/22 K.ISHIZAKA  'INSERT 2006/08/31 K.ISHIZAKA
    Dim dbKase              As Object                                           'INSERT 2006/10/22 K.ISHIZAKA
    Dim rsKase              As Object                                           'INSERT 2006/10/22 K.ISHIZAKA
    On Error GoTo Err_Handler
    
    'SQL-Server接続
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(strOpenArgs(C_CONNECTION))
    '詳細データ作成
    dbSQLServer.Execute fncCreateTempTable(), adCmdText
    dbSQLServer.Execute fncCreateTempIndex(), adCmdText                         'INSERT 2006/01/20 K.ISHIZAKA
    dbSQLServer.Execute fncInsertTempTable(), adCmdText
'    dbSqlServer.Execute fncUpdateTempTable_20(), adCmdText                     'DELETE 2005/12/05 K.ISHIZAKA
    dbSQLServer.Execute fncUpdateTempTable_02(), adCmdText
    dbSQLServer.Execute fncUpdateTempTable_20(), adCmdText                      'INSERT 2005/12/05 K.ISHIZAKA
    'データ取得
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    rsSqlServer.Open fncSelectTempTable(), dbSQLServer, adOpenForwardOnly, adLockReadOnly
    'Access接続
    Set dbAccess = CurrentDb()
    'Access側ワークテーブル削除
    Call subDropWorkTable(dbAccess)
    'Access側ワークテーブル作成
    If fncCreateWorkTable(dbAccess, rsSqlServer) = False Then
        'SQL-Server側切断
        rsSqlServer.Close
        Set rsSqlServer = Nothing
        dbSQLServer.Close
        Set dbSQLServer = Nothing
        'Access側切断
        dbAccess.Close
        Set dbAccess = Nothing
        fncGetData = False
        Exit Function
    End If
    
    '加瀬DB接続
'    strDNS = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB='ODBC_DATA_SOURCE_NAME'")     'DELETE 2006/10/22 K.ISHIZAKA  'INSERT 2006/08/31 K.ISHIZAKA
'    Set dbKase = Workspaces(0).OpenDatabase(strDNS, dbDriverNoPrompt, False, MSZZ007_M00()) 'DELETE 2006/10/22 K.ISHIZAKA  'INSERT 2006/08/31 K.ISHIZAKA
'    Set rsKase = dbKase.OpenRecordset(fncSelectREPR_TRAN(), dbOpenForwardOnly, dbSQLPassThrough, dbReadOnly) 'DELETE 2006/10/22 K.ISHIZAKA  'INSERT 2006/08/31 K.ISHIZAKA
    Set dbKase = CreateObject("ADODB.Connection")                               'INSERT 2006/10/22 K.ISHIZAKA
    dbKase.CommandTimeout = 180                                                 'INSERT 2006/10/22 K.ISHIZAKA
    Call dbKase.Open(MSZZ007_M10())                                             'INSERT 2006/10/22 K.ISHIZAKA
    Set rsKase = CreateObject("ADODB.Recordset")                                'INSERT 2006/10/22 K.ISHIZAKA
    rsKase.Open fncSelectREPR_TRAN(), dbKase, adOpenForwardOnly, adLockReadOnly 'INSERT 2006/10/22 K.ISHIZAKA

    'Access側ワークテーブル
    Set rsAccess = dbAccess.OpenRecordset("WK_FVS021_TEMP", dbOpenDynaset, dbAppendOnly)
    With rsSqlServer
        Do While Not .EOF
            rsAccess.AddNew
            For i = 0 To .Fields.Count - 1
                With .Fields(i)
                    rsAccess.Fields(.NAME) = .VALUE
                End With
            Next
'>>> INSERT START 2006/08/31 K.ISHIZAKA >>>
            Do While Not rsKase.EOF
                Select Case StrComp(rsKase.Fields("REPRT_ROOMC"), Format(.Fields("CNTA_NO"), "000000"))
                Case 0
                    rsAccess.Fields("YOUKT_BIKO") = rsKase.Fields("REPRT_CNT1N")
                    rsKase.MoveNext
                    Exit Do
                Case 1
                    Exit Do
                End Select
                rsKase.MoveNext
            Loop
            '顧客修繕を呼び出す
            If rsAccess.Fields("YOUKT_YUKBN") = "20" Then
                'rsAccess.Fields("YOUKT_RENKBN") = fncSelectKOTOT_CONTN(dbKase, .Fields("YOUKT_UCODE"))             'DELETE 2019/04/22 N.IMAI
                rsAccess.Fields("YOUKT_RENKBN") = LeftB(fncSelectKOTOT_CONTN(dbKase, .Fields("YOUKT_UCODE")), 120)  'INSERT 2019/04/22 N.IMAI
            End If
'<<< INSERT END   2006/08/31 K.ISHIZAKA <<<
            rsAccess.UPDATE
            .MoveNext
        Loop
        'SQL-Server側切断
        .Close
    End With
    Set rsSqlServer = Nothing
    dbSQLServer.Close
    Set dbSQLServer = Nothing
    '加瀬DB切断
    rsKase.Close                                                                'INSERT 2006/08/31 K.ISHIZAKA
    Set rsKase = Nothing                                                        'INSERT 2006/08/31 K.ISHIZAKA
    dbKase.Close                                                                'INSERT 2006/08/31 K.ISHIZAKA
    Set dbKase = Nothing                                                        'INSERT 2006/08/31 K.ISHIZAKA
    'Access側切断
    rsAccess.Close
    Set rsAccess = Nothing
    dbAccess.Close
    Set dbAccess = Nothing
    bChangeFlag = False
    Me.Caption = strCaption
    Me.サブ上.Requery
    Me.サブ下.Requery
    
    '▼ 2007/02/07 tajima
    ' 営業開始日に達してなければ部屋指定の予約ボタンは押せないようにする
    If Nz(Me.txt_YARD_BEGIN_DAY, "1900/01/01") > Format$(DATE, "yyyy/mm/dd") Then
        Me.サブ下.Form!cmdUke.Enabled = False
        Me.サブ上.Form!cmdUke.Enabled = False
     End If
    '▲ 2007/02/07 tajima
    
    fncGetData = True
    Exit Function

Err_Handler:
    MsgBox "fncGetData:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
    'SQL-Server側切断
    If Not rsSqlServer Is Nothing Then
        If rsSqlServer.State > 0 Then
            rsSqlServer.Close
        End If
    End If
    If Not dbSQLServer Is Nothing Then
        dbSQLServer.Close
    End If
    '加瀬DB切断
    If Not rsKase Is Nothing Then                                               'INSERT 2006/08/31 K.ISHIZAKA
        If rsKase.State > 0 Then                                                'INSERT 2006/10/22 K.ISHIZAKA
            rsKase.Close                                                        'INSERT 2006/08/31 K.ISHIZAKA
        End If                                                                  'INSERT 2006/10/22 K.ISHIZAKA
    End If                                                                      'INSERT 2006/08/31 K.ISHIZAKA
    If Not dbKase Is Nothing Then                                               'INSERT 2006/08/31 K.ISHIZAKA
        dbKase.Close                                                            'INSERT 2006/08/31 K.ISHIZAKA
    End If                                                                      'INSERT 2006/08/31 K.ISHIZAKA
    'Access側切断
    If Not rsAccess Is Nothing Then
        rsAccess.Close
    End If
    If Not dbAccess Is Nothing Then
        dbAccess.Close
    End If
    fncGetData = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : 確認と保存
'       MODULE_ID       : fncSaveData
'       CREATE_DATE     : 20050718 K.ISHIZAKA
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncSaveData() As Boolean
    If bChangeFlag Then
        If MsgBox("修正中の内容を保存します。よろしいですか？", vbQuestion + vbYesNo, constMsgKaku) = vbYes Then
            fncSaveData = fncUpdateCNTA_MAST()
            Exit Function
        End If
    End If
    fncSaveData = True
End Function

'==============================================================================*
'
'       MODULE_NAME     : 編集データ取得ＳＱＬ作成
'       MODULE_ID       : fncSelectWorkTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectWorkTable() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_NO,"
    strSQL = strSQL & " YOUKT_RENKBN,"
    strSQL = strSQL & " YOUKT_BIKO,"
    strSQL = strSQL & " YOUKT_UKNO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " WK_FVS021_TEMP "
    strSQL = strSQL & "WHERE CHANGE = 1 "
    strSQL = strSQL & "ORDER BY CNTA_NO "

    fncSelectWorkTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナマスタデータ取得ＳＱＬ作成
'       MODULE_ID       : fncSelectCNTA_MAST
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectCNTA_MAST() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_NO,"
    strSQL = strSQL & " CNTA_YTEL,"
    strSQL = strSQL & " CNTA_YNAME,"
    strSQL = strSQL & " CNTA_UPDATE "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " CNTA_MAST "
    strSQL = strSQL & "WHERE CNTA_CODE = " & Format(lngYardCode) & strOpenArgs(C_CNTA_SYU)
    strSQL = strSQL & " ORDER BY CNTA_NO "

    fncSelectCNTA_MAST = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナマスタ更新処理
'       MODULE_ID       : fncSelectCNTA_MAST
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateCNTA_MAST() As Boolean
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim dbAccess            As Database
    Dim rsAccess            As Recordset
    On Error GoTo Err_Handler
    
    'Access接続
    Set dbAccess = CurrentDb()
    'Access側ワークテーブル
    Set rsAccess = dbAccess.OpenRecordset(fncSelectWorkTable(), dbOpenForwardOnly)
    If rsAccess.EOF Then
        'Access側切断
        rsAccess.Close
        Set rsAccess = Nothing
        dbAccess.Close
        Set dbAccess = Nothing
        fncUpdateCNTA_MAST = True
        Exit Function
    End If
    'SQL-Server接続
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(strOpenArgs(C_CONNECTION))
    'データ取得
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    With rsSqlServer
        .Open fncSelectCNTA_MAST(), dbSQLServer, adOpenKeyset, adLockPessimistic
        Do While (Not .EOF) And (Not rsAccess.EOF)
            If .Fields("CNTA_NO") = rsAccess("CNTA_NO") Then
                '.Edit
                .Fields("CNTA_YTEL") = rsAccess("YOUKT_RENKBN")
                .Fields("CNTA_YNAME") = rsAccess("YOUKT_BIKO")
                .Fields("CNTA_UPDATE") = DATE
                .UPDATE
                .MoveNext
                rsAccess.MoveNext
            Else
                If .Fields("CNTA_NO") < rsAccess("CNTA_NO") Then
                    .MoveNext
                Else
                    rsAccess.MoveNext
                End If
            End If
        Loop
        'SQL-Server側切断
        .Close
    End With
    Set rsSqlServer = Nothing
    dbSQLServer.Close
    Set dbSQLServer = Nothing
    'Access側切断
    rsAccess.Close
    Set rsAccess = Nothing
    dbAccess.Close
    Set dbAccess = Nothing
    bChangeFlag = False
    Me.Caption = strCaption
    fncUpdateCNTA_MAST = True
    Exit Function

Err_Handler:
    MsgBox "fncUpdateCNTA_MAST:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
    'SQL-Server側切断
    If Not rsSqlServer Is Nothing Then
        If rsSqlServer.State > 0 Then
            rsSqlServer.Close
        End If
    End If
    If Not dbSQLServer Is Nothing Then
        dbSQLServer.Close
    End If
    'Access側切断
    If Not rsAccess Is Nothing Then
        rsAccess.Close
    End If
    If Not dbAccess Is Nothing Then
        dbAccess.Close
    End If
    fncUpdateCNTA_MAST = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナ修繕直近取得処理
'       MODULE_ID       : fncSelectREPR_TRAN
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectREPR_TRAN() As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " REPRT_ROOMC,"
    strSQL = strSQL & " SUBSTRING(REPRT_GENTD, 3, 2) + '/' +"
    strSQL = strSQL & " SUBSTRING(REPRT_GENTD, 5, 2) + ' ' +"
    strSQL = strSQL & " CASE WHEN REPRT_TYPEC = '02'"
    strSQL = strSQL & " THEN"
    strSQL = strSQL & " ("
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "  CODET_NAMEN"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  CODE_TABL"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  CODET_SIKBC = '251'"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  CODET_CODEC = CAST(REPRT_REPRC AS NUMERIC)"
    strSQL = strSQL & " )"
    strSQL = strSQL & " ELSE "
'    strSql = strSql & "  REPRT_CNT1N"                                          'DELETE 2007/02/07 K.ISHIZAKA
    strSQL = strSQL & "  ISNULL(REPRT_CNT1N,'')"                                'INSERT 2007/02/07 K.ISHIZAKA
    strSQL = strSQL & " END AS REPRT_CNT1N "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " REPR_TRAN oya "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " REPRT_BUMOC = '" & strOpenArgs(C_CONT_BUMOC) & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " REPRT_YARDC = '" & strOpenArgs(C_YARD_CODE) & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " REPRT_INPTI = '03' "
'    strSql = strSql & "AND"                                                    'DELETE 2006/09/11 K.ISHIZAKA
'    strSql = strSql & " REPRT_TYPEC != '02' "                                  'DELETE 2006/09/11 K.ISHIZAKA
    strSQL = strSQL & "AND"
    strSQL = strSQL & " REPRT_UNIQEC ="
    strSQL = strSQL & " ("
    strSQL = strSQL & " SELECT TOP 1"
    strSQL = strSQL & "  REPRT_UNIQEC"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  REPR_TRAN ko"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  ko.REPRT_BUMOC = oya.REPRT_BUMOC"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ko.REPRT_YARDC = oya.REPRT_YARDC"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ko.REPRT_ROOMC = oya.REPRT_ROOMC"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ko.REPRT_INPTI = oya.REPRT_INPTI"
    strSQL = strSQL & " AND"                                                    'INSERT 2006/09/11 K.ISHIZAKA
    strSQL = strSQL & "  ko.REPRT_TYPEC != '02' "                               'INSERT 2006/09/11 K.ISHIZAKA
    strSQL = strSQL & " ORDER BY"
    strSQL = strSQL & "  REPRT_GENTD DESC,"
    strSQL = strSQL & "  REPRT_INSED DESC,"
    strSQL = strSQL & "  REPRT_INSEJ DESC"
    strSQL = strSQL & " ) "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " REPRT_ROOMC"

    fncSelectREPR_TRAN = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 顧客修繕直近取得更新処理
'       MODULE_ID       : fncSelectKOTOT_CONTN
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'       PARAM           : dbSqlServer       データベース(I)
'                       : strKOTOT_KOKYC    顧客コード(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncSelectKOTOT_CONTN(dbSqlServer As Database, ByVal strKOTOT_KOKYC As String) As String   'DELETE 2006/10/22 K.ISHIZAKA
Private Function fncSelectKOTOT_CONTN(dbSQLServer As Object, ByVal strKOTOT_KOKYC As String) As String      'INSERT 2006/10/22 K.ISHIZAKA
'    Dim rsObject            As Recordset                                       'DELETE 2006/10/22 K.ISHIZAKA
    Dim rsObject            As Object                                           'INSERT 2006/10/22 K.ISHIZAKA
    
'    Set rsObject = dbSqlServer.OpenRecordset(fncSelectKOTO_TRAN(strKOTOT_KOKYC), dbOpenForwardOnly, dbSQLPassThrough, dbReadOnly)  'DELETE 2006/10/22 K.ISHIZAKA
    Set rsObject = CreateObject("ADODB.Recordset")                              'INSERT 2006/10/22 K.ISHIZAKA
    With rsObject
        .Open fncSelectKOTO_TRAN(Format(strKOTOT_KOKYC, "000000")), dbSQLServer, adOpenForwardOnly, adLockReadOnly 'INSERT 2006/10/22 K.ISHIZAKA
        If Not .EOF Then
            fncSelectKOTOT_CONTN = .Fields(0)
        End If
        .Close
    End With
End Function

'==============================================================================*
'
'       MODULE_NAME     : 顧客修繕直近取得処理
'       MODULE_ID       : fncSelectKOTO_TRAN
'       CREATE_DATE     : 2006/08/31        K.ISHIZAKA
'       PARAM           : strKOTOT_KOKYC    顧客コード(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSelectKOTO_TRAN(ByVal strKOTOT_KOKYC As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT TOP 1"
    strSQL = strSQL & " SUBSTRING(KOTOT_GENTD, 3, 2) + '/' +"
    strSQL = strSQL & " SUBSTRING(KOTOT_GENTD, 5, 2) + ' ' +"
'    strSql = strSql & " KOTOT_CONTN "                                          'DELETE 2007/02/07 K.ISHIZAKA
    strSQL = strSQL & " ISNULL(KOTOT_CONTN, '') "                               'INSERT 2007/02/07 K.ISHIZAKA
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " KOTO_TRAN "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " KOTOT_BUMOC = '" & strOpenArgs(C_CONT_BUMOC) & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " KOTOT_KOKYC = '" & strKOTOT_KOKYC & "' "
    strSQL = strSQL & "AND"                                                     'INSERT 2010/06/10 SHIBAZAKI
    strSQL = strSQL & " KOTOT_DISPI = 1 "                                       'INSERT 2010/06/10 SHIBAZAKI
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " KOTOT_GENTD DESC,"
    strSQL = strSQL & " KOTOT_INSED DESC,"
    strSQL = strSQL & " KOTOT_INSEJ DESC "

    fncSelectKOTO_TRAN = strSQL
End Function

'==============================================================================*
'   外部関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : キー押下処理
'       MODULE_ID       : KeyDown
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : KeyCode           キーコード
'                       : Shift             シフトキー状態
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    'F1:戻る
    'F2:保存
    'F3:取消
    'F5:段切り替え
    'F8:配置
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Me.cmd_終了.SetFocus
        'Call cmd_終了_Click
    Case vbKeyF2
        KeyCode = 0
        Me.cmd_保存.SetFocus
        'Call cmd_保存_Click
    Case vbKeyF3
        KeyCode = 0
        Me.cmd_取消.SetFocus
        'Call cmd_取消_Click
    Case vbKeyF5
        If Me.ActiveControl.Properties("NAME") = "サブ上" Then
            With Me.サブ下
                .SetFocus
                .Form!cmdUke.SetFocus
            End With
        Else
            With Me.サブ上
                .SetFocus
                .Form!cmdUke.SetFocus
            End With
        End If
    Case vbKeyF8
        Call CmReprMod.OpenFVS90n(KeyCode, strOpenArgs(C_YARD_CODE))            'INSERT 2006/08/31 K.ISHIZAKA
        KeyCode = 0
'        cmd_配置.SetFocus                                                      'DELETE 2006/08/31 K.ISHIZAKA
        'Call cmd_配置_Click
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        Select Case Me.ActiveControl.Properties("NAME")                         'INSERT 2006/08/31 K.ISHIZAKA
        Case "サブ上", "サブ下"                                                 'INSERT 2006/08/31 K.ISHIZAKA
            Call CmReprMod.OpenFVS90n(KeyCode, strOpenArgs(C_YARD_CODE), Me.ActiveControl.Form!txt_YOUKT_NO, Me.ActiveControl.Form!txt_YOUKT_UCODE) 'INSERT 2006/08/31 K.ISHIZAKA
        End Select                                                              'INSERT 2006/08/31 K.ISHIZAKA
        KeyCode = 0
    End Select
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 修正中に設定
'       MODULE_ID       : SetChangeFlag
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetChangeFlag()
    If bChangeFlag = False Then
        Me.Caption = strCaption & "(修正中)"
        bChangeFlag = True
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 詳細データ検索処理
'       MODULE_ID       : Reload
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : strYard1          ヤードコード１
'                       : strYard2          ヤードコード２
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function Reload(ByVal strYard1 As String, ByVal strYard2 As String) As Boolean
    Dim beforForm As Form
    
    If bChangeFlag = False Then
        If (lngYardCode = Val(strYard1)) Or (lngYardCode = Val(strYard2)) Then
            Call subSelFVS011                                                   'INSERT 20050829 K.ISHIZAKA
            If fncGetData() = False Then
                Reload = False
                Exit Function
            End If

            On Error Resume Next
            Set beforForm = Forms("FVS011")
            If Not beforForm Is Nothing Then
                If beforForm.Reload(strYard1, strYard2) = False Then
                    Reload = False
                    Exit Function
                End If
            End If
        End If
    End If
    Reload = True
End Function

'==============================================================================*
'
'       MODULE_NAME     : 予約受付画面表示処理
'       MODULE_ID       : ShowFVS500
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : strYOUKT_NO       コンテナ番号
'                       : strYOUKT_UKNO     予約受付番号
'                       : aレンタル用途     add 2007/06/02 tajima
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub ShowFVS500(ByVal strYOUKT_NO As String, ByVal strYOUKT_UKNO As String, _
                      Optional aレンタル用途 As String = "")
    On Error Resume Next
    doCmd.RunCommand acCmdSaveRecord
    If fncSaveData() Then
        If strYOUKT_NO <> "" Then
            strYOUKT_NO = Format(Val(strYOUKT_NO), "000000")
        End If
        doCmd.OpenForm "FVS500", , , , , , strOpenArgs(C_YARD_CODE) & strYOUKT_NO & strYOUKT_UKNO & "," & aレンタル用途
    End If
End Sub

'****************************  ended or program ********************************
