VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　空状況一覧
'   プログラムＩＤ　：　FVS011
'   作　成　日　　　：  2005/07/28
'   作　成　者　　　：  K.ISHIZAKA
'   Ver             ：  0.0
'
'   MSZZ007, MSZZ013 を KAGTOS からインポートしなおすこと
'   CNTA_MAST にインデックスを追加する(CNTA_CODE, CNTA_STEP)
'**********************************************
'修正履歴
'   修　正　日　　　：  20050829
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.1
'   修　正　内　容　：  1.コンテナ種別未選択は1～40とする
'                       2.検索条件ヤード名、ヤード住所を選択可能とする
'                       3.ヤードコードをコンボボックスで選択可能とする
'                       4.解約予定を表示するかをチェックボックスで選択可能とする
'                       5.取り置きを表示するかをチェックボックスで選択可能とする
'   CARG_FILE にインデックスを追加する(CARG_CONTNA, CARG_YCODE, CARG_KYDATE)
'**********************************************
'   修　正　日　　　：  20051003
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.2
'   修　正　内　容　：  1.コンテナ種別未選択は0より大きいとする
'                       2.解約予定数、取置数の算出時にコンテナ種別の条件が抜けていたので追加
'**********************************************
'   修　正　日　　　：  2005/11/17
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.3
'   修　正　内　容　：  1.コンテナ種別を実帖に変更する
'**********************************************
'   修　正　日　　　：  2005/12/05
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.4
'   修　正　内　容　：  1.コンテナ確定予約を自社理由に変更する対応
'                       2.コンテナマスタ・利用可否を条件に含める（種別ゼロの判断はまだ残す）
'**********************************************
'   修　正　日　　　：  2005/12/15
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  3.取り置きされているかの判断は、予約紹介トランを参照する
'**********************************************
'   修　正　日　　　：  2006/01/20
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.5
'   修　正　内　容　：  1.一時テーブルの制約を一時インデックスに変更する
'**********************************************
'   修　正　日　　　：  2006/01/22
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.6
'   修　正　内　容　：  1.自社理由のデータを非表示とする
'**********************************************
'   修　正　日　　　：  2006/01/26
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.7
'   修　正　内　容　：  1.ヤードをコンボボックスから選択したとき、コンテナ条件が未設定だった不具合の改修
'**********************************************
'   修　正　日　　　：  2006/03/31
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.8
'   修　正　内　容　：  1.地図表示対応
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
'**********************************************
'   修　正　日　　　：  2007/02/07
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  0.a
'   修　正　内　容　：  1.紹介制限期限日を過ぎたヤードであること
'**********************************************
'   修　正　日　　　：  2007/05/21
'   修　正　者　　　：  tajima
'   Ver             ：  0.b
'   修　正　内　容　：  1.検索条件にレンタル用途を追加
'**********************************************
'   修　正　日　　　：  2008/03/23
'   修　正　者　　　：  tajima
'   Ver             ：  0.c
'   修　正　内　容　：  1.ヤードコード指定検索の場合は距離順とする
'**********************************************
'   修　正　日　　　：  2009/12/02
'   修　正　者　　　：  K.ISHIZAKA
'   Ver             ：  1.0
'   修　正　内　容　：  1.統合検索MDBのときはコンテナとトランク両方を検索する
'**********************************************
'   修　正　日　　　：  2012/07/04
'   修　正　者　　　：  M.HONDA
'   Ver             ：  1.1
'   修　正　内　容　：  1.ログ出力対応
'**********************************************
'   修　正　日　　　：  2012/11/28
'   修　正　者　　　：  M.HONDA
'   Ver             ：  1.2
'   修　正　内　容　：  1.ｷｬﾝﾍﾟｰﾝ文字数対応
'**********************************************
'   修　正　日　　　：  2025/01/27
'   修　正　者　　　：  N.IMAI
'   Ver             ：  1.3
'   修　正　内　容　：  開始日がNULLのヤードは表示しない
'**********************************************

Option Compare Database
Option Explicit
'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FVS011"            'プログラムＩＤ
Private Const constMsgErr = "エラー"

Private Const adCmdText = 1
Private Const adOpenForwardOnly = 0
Private Const adLockReadOnly = 1

Private Const C_GREEN_COLOR         As Long = 8454016                           'INSERT 2009/12/02 K.ISHIZAKA

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private strConnection       As String       'SQL-Server接続文字列
Private strCNTA_SYU         As String       'コンテナ種別の条件(一覧表示したときの条件でReloadするためここで宣言する)
Private strArgments         As String       '次画面に渡す各種引数
Private bKaiyaku            As Boolean      '解約予定                           'INSERT 20050829 K.ISHIZAKA
Private bTorioki            As Boolean      '取り置き                           'INSERT 20050829 K.ISHIZAKA

Private colKey              As New Collection   'コレクションキー               'INSERT 2009/12/02 K.ISHIZAKA
Private colConnection       As New Collection   'SQL-Server接続文字列           'INSERT 2009/12/02 K.ISHIZAKA
Private colArgments         As New Collection   '次画面に渡す各種引数           'INSERT 2009/12/02 K.ISHIZAKA

'==============================================================================*
'   イベント関数
'==============================================================================*
'==============================================================================*
'
'       MODULE_NAME     : クリックされたとき
'       MODULE_ID       : chk_NearYard_Click
'       CREATE_DATE     : 2006/04/06        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_NearYard_Click()
    Me.txt_Km.Enabled = Me.chk_NearYard
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キーが押されたとき
'       MODULE_ID       : chk_NearYard_KeyPress
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_NearYard_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 12288, vbKeySpace
        Me.chk_NearYard = Not Me.chk_NearYard
    End Select
End Sub
'==============================================================================*
'
'       MODULE_NAME     : 種別が選択されたとき
'       MODULE_ID       : cmb_StepFrom_AfterUpdate
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_StepFrom_AfterUpdate()                                         'DELETE 2005/11/17 K.ISHIZAKA
'    With Me.cmb_StepFrom                                                       'DELETE 2005/11/17 K.ISHIZAKA
'        Me.txt_StepFrom = IIf(.ListIndex >= 0, .Column(1, .ListIndex), "")     'DELETE 2005/11/17 K.ISHIZAKA
'    End With                                                                   'DELETE 2005/11/17 K.ISHIZAKA
'End Sub                                                                        'DELETE 2005/11/17 K.ISHIZAKA

'==============================================================================*
'
'       MODULE_NAME     : 種別が選択されたとき
'       MODULE_ID       : cmb_StepTo_AfterUpdate
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_StepTo_AfterUpdate()                                           'DELETE 2005/11/17 K.ISHIZAKA
'    With Me.cmb_StepTo                                                         'DELETE 2005/11/17 K.ISHIZAKA
'        Me.txt_StepTo = IIf(.ListIndex >= 0, .Column(1, .ListIndex), "")       'DELETE 2005/11/17 K.ISHIZAKA
'    End With                                                                   'DELETE 2005/11/17 K.ISHIZAKA
'End Sub                                                                        'DELETE 2005/11/17 K.ISHIZAKA

'==============================================================================*
'
'       MODULE_NAME     : ヤードコードが選択されたとき
'       MODULE_ID       : cmb_YardCode_AfterUpdate
'       CREATE_DATE     : 2005/08/29        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_YardCode_AfterUpdate()
    With Me.cmb_YardCode
        If .ListIndex >= 0 Then
            .VALUE = Format(.Text, .Format)
            Me.txt_YardAddr = .Column(1, .ListIndex)
            Me.fra_Yard.VALUE = False
'            strCNTA_SYU = " AND CNTA_SYU > 0" & fncWhere_CNTA_SIZE()           'DELETE 2006/08/31 K.ISHIZAKA 'INSERT 2006/01/26 K.ISHIZAKA
'            strCNTA_SYU = fncWhere_CNTA_SIZE()                                 'DELETE 2007/05/21 tajima     'INSERT 2006/08/31 K.ISHIZAKA
            strCNTA_SYU = fncWhere_CNTA_SIZE() & fncWhere_CNTA_USAGE()          'INSERT 2007/05/21 tajima
            
            Call fncGetData(.VALUE)
'            Call MapOpen                                                        'INSERT 2006/03/31 K.ISHIZAKA
        End If
    End With
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 地図ボタンが押されたとき
'       MODULE_ID       : cmd_MAP_Click
'       CREATE_DATE     : 2006/03/31        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_MAP_Click()
    Call MapOpen                                                                'INSERT 2006/03/31 K.ISHIZAKA
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 検索ボタンが押されたとき
'       MODULE_ID       : cmd_Serach_Click
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Serach_Click()
    Dim dbAccess            As Database
    
    Me.cmb_YardCode.VALUE = ""                                                  'INSERT 20050829 K.ISHIZAKA
    If Trim(Me.txt_YardAddr.VALUE) <> "" Then
'        If fncCheck_CNTA_SYU() Then                                            'DELETE 2005/11/17 K.ISHIZAKA
'            strCNTA_SYU = fncWhere_CNTA_SYU()                                  'DELETE 2005/11/17 K.ISHIZAKA
        If fncCheck_CNTA_SIZE() Then                                            'INSERT 2005/11/17 K.ISHIZAKA
'            strCNTA_SYU = " AND CNTA_SYU > 0" & fncWhere_CNTA_SIZE()           'DELETE 2006/08/31 K.ISHIZAKA 'INSERT 2005/11/17 K.ISHIZAKA
'            strCNTA_SYU = fncWhere_CNTA_SIZE()                                 'DELETE 2007/05/21 tajima     'INSERT 2006/08/31 K.ISHIZAKA
            strCNTA_SYU = fncWhere_CNTA_SIZE() & fncWhere_CNTA_USAGE()          'INSERT 2007/05/21 tajima

'            Call fncGetData                                                    'DELETE 20050829 K.ISHIZAKA
            Call fncGetData("")                                                 'INSERT 20050829 K.ISHIZAKA
'            Call MapOpen                                                        'INSERT 2006/03/31 K.ISHIZAKA
        End If
    Else
        Set dbAccess = CurrentDb()
        'Access側ワークテーブル削除
        Call subDropWorkTable(dbAccess)
        'Access側ワークテーブル作成
        Call fncCreateWorkTable(dbAccess)
        dbAccess.Close
    End If
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
    If Not fncFormIsLoaded("FKS000") Then                                       'INSERT START 2009/12/02 K.ISHIZAKA
        Call CmSysEnd("統合検索システム")
    Else                                                                        'INSERT END   2009/12/02 K.ISHIZAKA
        On Error Resume Next
        doCmd.Close acForm, Me.FormName
    End If                                                                      'INSERT 2009/12/02 K.ISHIZAKA
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
    Call MSHS010_M00(PROG_ID, "2")
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
    Dim dbAccess            As Database
    Dim rsAccess            As Recordset
    Dim strSQL              As String
    Dim strCONT_BUMOC       As String       '部門コード
    Dim strCONT_DISKBACK    As String
    Dim strCONT_EXCEL_PATH  As String

    On Error GoTo Err_Handle
    
    'Call MSHS010_M00(PROG_ID, "1")
    Set dbAccess = CurrentDb()
    If fncCreateWorkTable(dbAccess) = False Then
        Me.cmd_Serach.Enabled = False
        Exit Sub
    End If
    If (MSZZ013_M00("ADODB.Connection") = False) Or (MSZZ013_M00("ADODB.Recordset") = False) Then
        Me.cmd_Serach.Enabled = False
        MsgBox "「ActiveX Data Objects」がインストールされてません!!"
        Exit Sub
    End If
    If Not fncFormIsLoaded("FKS000") Then                                       'INSERT START 2009/12/02 K.ISHIZAKA
        Me.[フォームヘッダー].BackColor = C_GREEN_COLOR
        cmb_YardCode.RowSource = ""             'ヤードコード検索をできないようにする
        cmb_YardCode.Enabled = False            'リンクテーブルを見ないようにする
        Call set_cmb_Usage("H")                 'レンタル用途でリンクテーブルを見ないようにする
        Call makeArgmentsCollection("H")
        Call makeArgmentsCollection("8")
    Else                                                                        'INSERT END   2009/12/02 K.ISHIZAKA
        strSQL = strSQL & "SELECT "
        strSQL = strSQL & " CONT_BUMOC,"
        strSQL = strSQL & " CONT_DISKBACK,"
        strSQL = strSQL & " CONT_EXCEL_PATH "
        strSQL = strSQL & "FROM dbo_CONT_MAST "
        strSQL = strSQL & "WHERE CONT_KEY = 1 "
        Set dbAccess = CurrentDb()
        Set rsAccess = dbAccess.OpenRecordset(strSQL, dbOpenSnapshot)
        With rsAccess
            If Not .EOF Then
                strCONT_BUMOC = Trim(Nz(.Fields("CONT_BUMOC")))
                strCONT_DISKBACK = Trim(Nz(.Fields("CONT_DISKBACK")))
                strCONT_EXCEL_PATH = Trim(Nz(.Fields("CONT_EXCEL_PATH")))
            Else
                strCONT_BUMOC = ""
                strCONT_DISKBACK = ""
                strCONT_EXCEL_PATH = ""
            End If
            .Close
        End With
        Set rsAccess = Nothing
        dbAccess.Close
        Set dbAccess = Nothing
        If strCONT_BUMOC = "" Then
            Me.cmd_Serach.Enabled = False
            MsgBox "[CONT_MAST]の設定不備です。"
            Exit Sub
        End If
        If strCONT_EXCEL_PATH <> "" Then
            If Right(strCONT_EXCEL_PATH, 1) <> "\" Then
                strCONT_EXCEL_PATH = strCONT_EXCEL_PATH & "\"
            End If
        End If
        strConnection = MSZZ007_M10(strCONT_BUMOC)
        strArgments = Join(Array(strCONT_BUMOC, strCONT_DISKBACK, strCONT_EXCEL_PATH, strConnection, LsGetUserName()), ",")
        
        colKey.Add strCONT_BUMOC, strCONT_BUMOC                                 'INSERT START 2009/12/02 K.ISHIZAKA
        colConnection.Add strConnection, strCONT_BUMOC
        colArgments.Add strArgments, strCONT_BUMOC
    End If                                                                      'INSERT END   2009/12/02 K.ISHIZAKA
    Me.fra_Yard.VALUE = True                                                    'INSERT 20050829 K.ISHIZAKA
    Exit Sub
    
Err_Handle:
    MsgBox "Form_Load:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
    Me.cmd_Serach.Enabled = False
    If Not rsAccess Is Nothing Then
        rsAccess.Close
    End If
    If Not dbAccess Is Nothing Then
        dbAccess.Close
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : Form_Open
'       MODULE_ID       : Form_Open
'       CREATE_DATE     :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")      'ADD 2006/03/18 Tajima
    Select Case strCONT_BUMOC                                                   'REP 2006/03/18 Tajima
    Case "H"
        ラベル2.BackColor = RGB(255, 0, 204)
    End Select
    
    
    Call MSHS010_M00(PROG_ID, "1")
    
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : txt_Km_KeyPress
'       CREATE_DATE     : 2006/04/06        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Km_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Dec(Me.txt_Km, KeyAscii, 6, 2)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : txt_SizeFrom_KeyPress
'       CREATE_DATE     : 2005/11/17        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SizeFrom_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Dec(Me.txt_SizeFrom, KeyAscii, 6, 2)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : txt_SizeTo_KeyPress
'       CREATE_DATE     : 2005/11/17        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SizeTo_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Dec(Me.txt_SizeTo, KeyAscii, 6, 2)
End Sub

'==============================================================================*
'   内部関数（イベントでない）
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
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

    Me.一覧サブ.SourceObject = ""
    dbAccess.TableDefs.Delete "WK_FVS011_TEMP"
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
'       MODULE_NAME     : ワークテーブルの作成
'       MODULE_ID       : fncCreateWorkTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : dbAccess          Access側データベース
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable(ByVal dbAccess As Database) As Boolean
    Dim tdfNew              As TableDef
    Dim i                   As Integer
    On Error GoTo Err_Handler
    
    Set tdfNew = dbAccess.CreateTableDef("WK_FVS011_TEMP")
    With tdfNew
        .Fields.Append .CreateField("NO", dbLong)                               'INSERT 2006/03/31 K.ISHIZAKA
        
        .Fields.Append .CreateField("YARD_CODE", dbLong)
        .Fields.Append .CreateField("YARD_NAME", dbText, 36)
        .Fields.Append .CreateField("YARD_ADDR_1", dbText, 36)
        .Fields.Append .CreateField("YARD_ADDR_2", dbText, 36)
        .Fields.Append .CreateField("YARD_ADDR_3", dbText, 36)
        .Fields.Append .CreateField("YARD_NOTE", dbText, 36)
        .Fields.Append .CreateField("YARD_BIKON", dbText, 60)
        
        '' 2012/11/28 M.HONDA START
        '' .Fields.Append .CreateField("YARD_SEV1N", dbText, 60)
        '' .Fields.Append .CreateField("YARD_SEV2N", dbText, 60)
        '' .Fields.Append .CreateField("YARD_SEV3N", dbText, 60)
        .Fields.Append .CreateField("YARD_SEV1N", dbText, 80)
        .Fields.Append .CreateField("YARD_SEV2N", dbText, 80)
        .Fields.Append .CreateField("YARD_SEV3N", dbText, 80)
        '' 2012/11/28 M.HONDA END
        
        .Fields.Append .CreateField("YARD_ENDEN", dbText, 60)
        
        .Fields.Append .CreateField("YARD_BEGIN_DAY", dbDate)                   'INSERT 2006/08/31 K.ISHIZAKA
        .Fields.Append .CreateField("YARD_RENTEND_DAY", dbDate)                 'INSERT 2006/08/31 K.ISHIZAKA
        .Fields.Append .CreateField("YARD_BEGIN_DAY2", dbText, 7)               'INSERT 2006/08/31 K.ISHIZAKA
        .Fields.Append .CreateField("YARD_RENTEND_DAY2", dbText, 7)             'INSERT 2006/08/31 K.ISHIZAKA
        
        .Fields.Append .CreateField("UP_AKI", dbLong)
        .Fields.Append .CreateField("DW_AKI", dbLong)
'        .Fields.Append .CreateField("UP_KAKUTEI", dbLong)                      'DELETE 2005/12/05 K.ISHIZAKA
'        .Fields.Append .CreateField("DW_KAKUTEI", dbLong)                      'DELETE 2005/12/05 K.ISHIZAKA
        .Fields.Append .CreateField("MIKAKUTEI", dbLong)
        
        .Fields.Append .CreateField("KAIYAKU", dbLong)                          'INSERT 20050829 K.ISHIZAKA
        .Fields.Append .CreateField("TORIOKI", dbLong)                          'INSERT 20050829 K.ISHIZAKA
        
        .Fields.Append .CreateField("YARD_IDO", dbDouble)                       'INSERT 2006/03/31 K.ISHIZAKA
        .Fields.Append .CreateField("YARD_KEIDO", dbDouble)                     'INSERT 2006/03/31 K.ISHIZAKA
        .Fields.Append .CreateField("YARD_KIRO", dbDouble)                      'INSERT 2006/03/31 K.ISHIZAKA
        
        .Fields.Append .CreateField("SERVICE", dbInteger)
        .Fields.Append .CreateField("SUM_AKI", dbLong)
        
        .Fields.Append .CreateField("SORT_VALUE", dbLong)   '2008/03/23 add tajima
        .Fields.Append .CreateField("KEY", dbText, 1)                           'INSERT 2009/12/02 K.ISHIZAKA
        
        For i = 0 To .Fields.Count - 1
            If .Fields(i).Type = dbText Then
                .Fields(i).AllowZeroLength = True
            End If
        Next
    End With
    dbAccess.TableDefs.Append tdfNew
    Me.一覧サブ.SourceObject = "FVS011_S01"
    fncCreateWorkTable = True
    Exit Function

Err_Handler:
    MsgBox "fncCreateWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
    Me.一覧サブ.SourceObject = "FVS011_S01"
    fncCreateWorkTable = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : ワークテーブルデータ取得ＳＱＬ作成
'       MODULE_ID       : fncSelectWorkTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : strYard1          ヤードコード１
'                       : strYard2          ヤードコード２
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncSelectWorkTable(ByVal strYard1 As String, ByVal strYard2 As String) As String
Private Function fncSelectWorkTable(ByVal strKey As String, ByVal strYard1 As String, ByVal strYard2 As String) As String
    Dim strSQL              As String
    
    strSQL = "SELECT * FROM WK_FVS011_TEMP "
    If (strYard2 = "") Or (strYard1 = strYard2) Then
        strSQL = strSQL & "WHERE  YARD_CODE = " & strYard1
        strSQL = strSQL & " AND KEY = '" & strKey & "' "                        'INSERT 2009/12/02 K.ISHIZAKA
    Else
        strSQL = strSQL & "WHERE  YARD_CODE IN (" & strYard1 & "," & strYard2 & ") "
        strSQL = strSQL & " AND KEY = '" & strKey & "' "                        'INSERT 2009/12/02 K.ISHIZAKA
        strSQL = strSQL & "ORDER BY YARD_CODE"
    End If
    fncSelectWorkTable = strSQL
End Function

'DELETE START 2005/11/17 K.ISHIZAKA
'==============================================================================*
'
'       MODULE_NAME     : コンテナ種別入力チェック処理
'       MODULE_ID       : fncCheck_CNTA_SYU
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncCheck_CNTA_SYU() As Boolean
'    With Me.cmb_StepFrom
'        If Trim(.Value) <> "" And .ListIndex < 0 Then
'            .SetFocus
'            MsgBox "種別From：" & vbCrLf & "入力した値が不適当です。", vbExclamation, constMsgErr
'            fncCheck_CNTA_SYU = False
'            Exit Function
'        End If
'    End With
'    With Me.cmb_StepTo
'        If Trim(.Value) <> "" And .ListIndex < 0 Then
'            .SetFocus
'            MsgBox "種別To：" & vbCrLf & "入力した値が不適当です。", vbExclamation, constMsgErr
'            fncCheck_CNTA_SYU = False
'            Exit Function
'        End If
'    End With
'    fncCheck_CNTA_SYU = True
'End Function

'==============================================================================*
'
'       MODULE_NAME     : コンテナ種別条件作成処理
'       MODULE_ID       : fncWhere_CNTA_SYU
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncWhere_CNTA_SYU() As String
'    '両方とも空
'    If (Me.cmb_StepFrom.ListIndex < 0) And (Me.cmb_StepTo.ListIndex < 0) Then
''        fncWhere_CNTA_SYU = ""                                                 'DELETE 20050829 K.ISHIZAKA
''        fncWhere_CNTA_SYU = " AND CNTA_SYU BETWEEN 1 AND 40"                   'DELETE 20051003 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        fncWhere_CNTA_SYU = " AND CNTA_SYU > 0"                                 'INSERT 20051003 K.ISHIZAKA
'        Exit Function
'    End If
'    'どっちか空
'    If (Me.cmb_StepFrom.ListIndex < 0) Or (Me.cmb_StepTo.ListIndex < 0) Then
'        fncWhere_CNTA_SYU = " AND CNTA_SYU = " & Nz(Me.cmb_StepFrom) & Nz(Me.cmb_StepTo)
'        Exit Function
'    End If
'    '一緒
'    If Me.cmb_StepFrom.ListIndex = Me.cmb_StepTo.ListIndex Then
'        fncWhere_CNTA_SYU = " AND CNTA_SYU = " & Nz(Me.cmb_StepFrom)
'        Exit Function
'    End If
'    '両方入力
'    If Me.cmb_StepFrom.ListIndex < Me.cmb_StepTo.ListIndex Then
'        fncWhere_CNTA_SYU = " AND CNTA_SYU BETWEEN " & Nz(Me.cmb_StepFrom) & " AND " & Nz(Me.cmb_StepTo)
'    Else
'        fncWhere_CNTA_SYU = " AND CNTA_SYU BETWEEN " & Nz(Me.cmb_StepTo) & " AND " & Nz(Me.cmb_StepFrom)
'    End If
'End Function
'DELETE END   2005/11/17 K.ISHIZAKA

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
    
    strSQL = strSQL & "CREATE TABLE #WK_FVS011_TEMP ("
    strSQL = strSQL & " YARD_CODE       NUMERIC(6)  NOT NULL,"
    strSQL = strSQL & " YARD_NAME       VARCHAR(36),"
    strSQL = strSQL & " YARD_ADDR_1     VARCHAR(36),"
    strSQL = strSQL & " YARD_ADDR_2     VARCHAR(36),"
    strSQL = strSQL & " YARD_ADDR_3     VARCHAR(36),"
    strSQL = strSQL & " YARD_NOTE       VARCHAR(36),"
    strSQL = strSQL & " YARD_BIKON      VARCHAR(60),"
    
    '' 2012/11/28 M.HONDA START
    '' strSQL = strSQL & " YARD_SEV1N      VARCHAR(60),"
    '' strSQL = strSQL & " YARD_SEV2N      VARCHAR(60),"
    '' strSQL = strSQL & " YARD_SEV3N      VARCHAR(60),"
    strSQL = strSQL & " YARD_SEV1N      VARCHAR(80),"
    strSQL = strSQL & " YARD_SEV2N      VARCHAR(80),"
    strSQL = strSQL & " YARD_SEV3N      VARCHAR(80),"
    '' 2012/11/28 M.HONDA END
    
    strSQL = strSQL & " YARD_ENDEN      VARCHAR(60),"
    strSQL = strSQL & " YARD_BEGIN_DAY   DATETIME,"                             'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " YARD_RENTEND_DAY DATETIME,"                             'INSERT 2006/08/31 K.ISHIZAKA
    strSQL = strSQL & " UP_AKI          NUMERIC(9),"
    strSQL = strSQL & " DW_AKI          NUMERIC(9),"
'    strSQL = strSQL & " UP_KAKUTEI      NUMERIC(9),"                           'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & " DW_KAKUTEI      NUMERIC(9),"                           'DELETE 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & " MIKAKUTEI       NUMERIC(9),"
    strSQL = strSQL & " KAIYAKU         NUMERIC(9),"                            'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & " TORIOKI         NUMERIC(9),"                            'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & " YARD_IDO        NUMERIC(8,6),"                          'INSERT 2006/03/31 K.ISHIZAKA
    strSQL = strSQL & " YARD_KEIDO      NUMERIC(9,6),"                          'INSERT 2006/03/31 K.ISHIZAKA
    strSQL = strSQL & " YARD_KIRO       NUMERIC(7,4),"                          'INSERT 2006/03/31 K.ISHIZAKA
    strSQL = strSQL & " SERVICE AS CAST(DATALENGTH("
    strSQL = strSQL & "            ISNULL(YARD_SEV1N,'') +"
    strSQL = strSQL & "            ISNULL(YARD_SEV2N,'') +"
    strSQL = strSQL & "            ISNULL(YARD_SEV3N,'') ) AS BIT),"
'    strSQL = strSQL & " SUM_AKI AS (UP_AKI + DW_AKI) "                         'DELETE 20050829 K.ISHIZAKA
    strSQL = strSQL & " SUM_AKI AS (UP_AKI + DW_AKI + ISNULL(KAIYAKU,0) + ISNULL(TORIOKI,0)) " 'INSERT 20050829 K.ISHIZAKA
'>>> DELETE START 2006/01/20 K.ISHIZAKA >>>
'    strSQL = strSQL & ",CONSTRAINT WK_KEY0B PRIMARY KEY"                        'INSERT 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & "     ("                                                  'INSERT 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & "     YARD_CODE"                                          'INSERT 2005/11/17 K.ISHIZAKA
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
    
    strSQL = strSQL & "CREATE INDEX #WK_FVS011_KEY0B ON #WK_FVS011_TEMP ("
    strSQL = strSQL & "     YARD_CODE"
    strSQL = strSQL & ")"
    
    fncCreateTempIndex = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ作成ＳＱＬ作成
'       MODULE_ID       : fncInsertTempTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : strYard1          ヤードコード１
'                       : strYard2          ヤードコード２
'                       : bNearYard         近隣有無区分
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncInsertTempTable(ByVal strYard1 As String, ByVal strYard2 As String) As String                              'DELETE 20050829 K.ISHIZAKA
Private Function fncInsertTempTable(ByVal strYard1 As String, ByVal strYard2 As String, ByVal bNearYard As Boolean) As String   'INSERT 20050829 K.ISHIZAKA
    Dim strSQL              As String
    
    strSQL = strSQL & "INSERT"
    strSQL = strSQL & " #WK_FVS011_TEMP "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_CODE,"
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
    
'    strSQL = strSQL & " NULL AS UP_AKI,"                                       'DELETE 20050829 K.ISHIZAKA
'    strSQL = strSQL & " NULL AS DW_AKI,"                                       'DELETE 20050829 K.ISHIZAKA
    strSQL = strSQL & " 0 AS UP_AKI,"                                           'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & " 0 AS DW_AKI,"                                           'INSERT 20050829 K.ISHIZAKA
    
'    strSQL = strSQL & "(SELECT COUNT(*)"                                       'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " FROM   CNTA_MAST"                                      'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " INNER JOIN"                                            'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & "        YOUK_TRAN"                                      'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " ON (   YOUKT_YCODE = CNTA_CODE"                        'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " AND    YOUKT_NO    = CNTA_NO"                          'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " AND    YOUKT_YUKBN = 1 )"                              'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " WHERE  CNTA_CODE   = YARD_CODE"                        'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " AND    CNTA_STEP   = 0" & strCNTA_SYU                  'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & ") AS UP_KAKUTEI,"                                       'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " 0 AS UP_KAKUTEI,"                                      'DELETE 2005/12/05 K.ISHIZAKA   'INSERT 20051006 K.ISHIZAKA
    
'    strSQL = strSQL & "(SELECT COUNT(*)"                                       'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " FROM   CNTA_MAST"                                      'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " INNER JOIN"                                            'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & "        YOUK_TRAN"                                      'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " ON (   YOUKT_YCODE = CNTA_CODE"                        'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " AND    YOUKT_NO    = CNTA_NO"                          'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " AND    YOUKT_YUKBN = 1 )"                              'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " WHERE  CNTA_CODE   = YARD_CODE"                        'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " AND    CNTA_STEP   = 1" & strCNTA_SYU                  'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & ") AS DW_KAKUTEI,"                                       'DELETE 20051006 K.ISHIZAKA
'    strSQL = strSQL & " 0 AS DW_KAKUTEI,"                                      'DELETE 2005/12/05 K.ISHIZAKA   'INSERT 20051006 K.ISHIZAKA
    
    strSQL = strSQL & "(SELECT COUNT(*)"
    strSQL = strSQL & " FROM   YOUK_TRAN"
    strSQL = strSQL & " WHERE  YOUKT_YCODE = YARD_CODE"
    strSQL = strSQL & " AND    YOUKT_YUKBN = 0"
    strSQL = strSQL & ") AS MIKAKUTEI,"
    
'    If bKaiyaku Then                                                           'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & "(SELECT COUNT(*)"                                   'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " FROM   CARG_FILE"                                  'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " WHERE  CARG_YCODE  = YARD_CODE"                    'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " AND    CARG_CONTNA = 0"                            'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " AND    CARG_KYDATE IS NOT NULL"                    'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & ") AS KAIYAKU,"                                      'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'    Else                                                                       'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
        strSQL = strSQL & "NULL AS KAIYAKU,"                                    'INSERT 20050829 K.ISHIZAKA
'    End If                                                                     'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
    
'    If bTorioki Then                                                           'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & "(SELECT COUNT(*)"                                   'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " FROM   YOUK_TRAN"                                  'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " WHERE  YOUKT_YCODE = YARD_CODE"                    'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & " AND    YOUKT_YUKBN = 2"                            'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & ") AS TORIOKI "                                      'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'    Else                                                                       'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
'        strSQL = strSQL & "NULL AS TORIOKI "                                   'DELETE 2006/03/31 K.ISHIZAKA   'INSERT 20050829 K.ISHIZAKA
'    End If                                                                     'DELETE 20051006 K.ISHIZAKA 'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & " NULL AS TORIOKI,"                                       'INSERT 2006/03/31 K.ISHIZAKA
    strSQL = strSQL & " YARD_IDO,"                                              'INSERT 2006/03/31 K.ISHIZAKA
    strSQL = strSQL & " YARD_KEIDO,"                                            'INSERT 2006/03/31 K.ISHIZAKA
    If (strYard1 <> "") And bNearYard Then                                      'INSERT 2006/03/31 K.ISHIZAKA
        strSQL = strSQL & "(SELECT NYAR_KIRO"                                   'INSERT 2006/03/31 K.ISHIZAKA
        strSQL = strSQL & " FROM   NYAR_MAST"                                   'INSERT 2006/03/31 K.ISHIZAKA
        strSQL = strSQL & " WHERE  NYAR_YCODE = " & strYard1                    'INSERT 2006/03/31 K.ISHIZAKA
        strSQL = strSQL & " AND    NYAR_NCODE = YARD_CODE"                      'INSERT 2006/03/31 K.ISHIZAKA
        strSQL = strSQL & ") AS YARD_KIRO "                                     'INSERT 2006/03/31 K.ISHIZAKA
    Else                                                                        'INSERT 2006/03/31 K.ISHIZAKA
        strSQL = strSQL & " NULL AS YARD_KIRO "                                 'INSERT 2006/03/31 K.ISHIZAKA
    End If                                                                      'INSERT 2006/03/31 K.ISHIZAKA

    strSQL = strSQL & "FROM "
    strSQL = strSQL & " YARD_MAST "
    
    If strYard1 = "" Then
'        If Me.chk_NearYard.Value Then                                          'DELETE 20050829 K.ISHIZAKA
        If bNearYard Then                                                       'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "WHERE  YARD_CODE IN"
            strSQL = strSQL & "    (  SELECT DISTINCT NYAR_NCODE"
            strSQL = strSQL & "       FROM   NYAR_MAST"
            strSQL = strSQL & "       WHERE  NYAR_YCODE IN"
            strSQL = strSQL & "           (  SELECT YARD_CODE"
            strSQL = strSQL & "              FROM   YARD_MAST "
        End If
        If Me.fra_Yard.VALUE Then                                               'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "WHERE (YARD_ADDR_1 LIKE '%" & Trim(Me.txt_YardAddr.VALUE) & "%' "
            strSQL = strSQL & " OR    YARD_ADDR_2 LIKE '%" & Trim(Me.txt_YardAddr.VALUE) & "%') "
        Else                                                                    'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "WHERE (YARD_NAME LIKE '%" & Trim(Me.txt_YardAddr.VALUE) & "%') " 'INSERT 20050829 K.ISHIZAKA
        End If                                                                  'INSERT 20050829 K.ISHIZAKA
'        If Me.chk_NearYard.Value Then                                          'DELETE 20050829 K.ISHIZAKA
        If bNearYard Then                                                       'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "           )"
            If IsNumeric(Trim(Nz(Me.txt_Km))) Then                              'INSERT 2006/03/31 K.ISHIZAKA
                strSQL = strSQL & " AND ISNULL(NYAR_KIRO,0) <= " & Me.txt_Km & " "  'INSERT 2006/03/31 K.ISHIZAKA
            End If                                                              'INSERT 2006/03/31 K.ISHIZAKA
            strSQL = strSQL & "    )"
        End If
'>>> DELETE START 2006/08/31 K.ISHIZAKA >>>
'        strSQL = strSQL & " AND EXISTS( "
'        strSQL = strSQL & "     SELECT * "
'        strSQL = strSQL & "     FROM   JARG_FILE "
'        strSQL = strSQL & "     WHERE JARG_YCODE = YARD_CODE "
'        strSQL = strSQL & "     AND   ISNULL(JARG_KYDATE, '9999/12/31') > GETDATE()"
'        strSQL = strSQL & "     )"
'<<< DELETE END   2006/08/31 K.ISHIZAKA <<<
        strSQL = strSQL & " AND ISNULL(YARD_RENTEND_DAY,  '9999/12/31') > GETDATE() "   'INSERT 2006/08/31 K.ISHIZAKA
        strSQL = strSQL & " AND ISNULL(YARD_INLIMIT_DAY,  '1900/01/01') < GETDATE() -1 "   'INSERT 2007/02/07 K.ISHIZAKA
        strSQL = strSQL & " AND YARD_BEGIN_DAY IS NOT NULL "                    'INSERT 2025/01/27 N.IMAI
    Else
        If bNearYard Then                                                       'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "WHERE  YARD_CODE IN"                             'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "    (  SELECT DISTINCT NYAR_NCODE"               'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "       FROM   NYAR_MAST"                         'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & "       WHERE  NYAR_YCODE = " & strYard1          'INSERT 20050829 K.ISHIZAKA
            If IsNumeric(Trim(Nz(Me.txt_Km))) Then                              'INSERT 2006/03/31 K.ISHIZAKA
                strSQL = strSQL & " AND ISNULL(NYAR_KIRO,0) <= " & Me.txt_Km & " "  'INSERT 2006/03/31 K.ISHIZAKA
            End If                                                              'INSERT 2006/03/31 K.ISHIZAKA
            strSQL = strSQL & "    )"                                           'INSERT 20050829 K.ISHIZAKA
            strSQL = strSQL & " AND ISNULL(YARD_RENTEND_DAY,  '9999/12/31') > GETDATE() "   'INSERT 2006/08/31 K.ISHIZAKA
            strSQL = strSQL & " AND ISNULL(YARD_INLIMIT_DAY,  '1900/01/01') < GETDATE() -1 "   'INSERT 2007/02/07 K.ISHIZAKA
            strSQL = strSQL & " AND YARD_BEGIN_DAY IS NOT NULL "                'INSERT 2025/01/27 N.IMAI
        Else                                                                    'INSERT 20050829 K.ISHIZAKA
            If (strYard2 = "") Or (strYard1 = strYard2) Then
                strSQL = strSQL & "WHERE  YARD_CODE = " & strYard1
                strSQL = strSQL & " AND ISNULL(YARD_RENTEND_DAY,  '9999/12/31') > GETDATE() "   'INSERT 2006/08/31 K.ISHIZAKA
                strSQL = strSQL & " AND ISNULL(YARD_INLIMIT_DAY,  '1900/01/01') < GETDATE() -1 "   'INSERT 2007/02/07 K.ISHIZAKA
                strSQL = strSQL & " AND YARD_BEGIN_DAY IS NOT NULL "            'INSERT 2025/01/27 N.IMAI
            Else
                strSQL = strSQL & "WHERE  YARD_CODE IN (" & strYard1 & "," & strYard2 & ") "
                strSQL = strSQL & "ORDER BY YARD_CODE"
            End If
        End If                                                                  'INSERT 20050829 K.ISHIZAKA
    End If
    
    fncInsertTempTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ更新ＳＱＬ作成（空き数）
'       MODULE_ID       : fncUpdateTempTable
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateTempTable() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & "    #WK_FVS011_TEMP "
    strSQL = strSQL & "SET"
    strSQL = strSQL & "    UP_AKI = UPAKI,"
    strSQL = strSQL & "    DW_AKI = DWAKI "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "    #WK_FVS011_TEMP "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "    CNTA_CODE,"
'    strSQL = strSQL & "    COUNT(NULLIF(CNTA_STEP, 1)) AS UPAKI,"              'DELETE 2005/11/17 K.ISHIZAKA
'    strSQL = strSQL & "    COUNT(NULLIF(CNTA_STEP, 0)) As DWAKI "              'DELETE 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "    COUNT(CASE WHEN CNTA_STEP = 0" & strCNTA_SYU         'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "     THEN 1 ELSE NULL END"                               'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "    ) AS UPAKI,"                                         'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "    COUNT(CASE WHEN CNTA_STEP = 1" & strCNTA_SYU         'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "     THEN 1 ELSE NULL END"                               'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "    ) AS DWAKI "                                         'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "    CNTA_MAST "
    strSQL = strSQL & "WHERE"
'    strSQL = strSQL & "    CNTA_CODE IN (SELECT YARD_CODE FROM #WK_FVS011_TEMP)" & strCNTA_SYU 'DELETE 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "    CNTA_CODE IN (SELECT YARD_CODE FROM #WK_FVS011_TEMP) " 'INSERT 2005/11/17 K.ISHIZAKA
    strSQL = strSQL & "AND CNTA_USE = 1 "                                       'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "AND NOT EXISTS"
    strSQL = strSQL & "   (SELECT *"
    strSQL = strSQL & "    FROM  CARG_FILE"
    strSQL = strSQL & "    WHERE CARG_YCODE  = CNTA_CODE"
    strSQL = strSQL & "    AND   CARG_NO     = CNTA_NO"
'    strSQL = strSQL & "    AND   CARG_KYDATE IS NULL )"                        'DELETE 20050829 K.ISHIZAKA
    strSQL = strSQL & "    AND   CARG_CONTNA = 0 )"                             'INSERT 20050829 K.ISHIZAKA
    strSQL = strSQL & "AND NOT EXISTS"
    strSQL = strSQL & "   (SELECT *"
'    strSQL = strSQL & "    FROM  YOUK_TRAN"                                    'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "    WHERE YOUKT_YCODE = CNTA_CODE"                      'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "    AND   YOUKT_NO    = CNTA_NO"                        'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "    AND   YOUKT_YUKBN = 10)"                            'DELETE 20050829 K.ISHIZAKA
'    strSQL = strSQL & "    AND   YOUKT_YUKBN IN(2, 10) )"                      'DELETE 2005/12/05 K.ISHIZAKA   'INSERT 20050829 K.ISHIZAKA
'    strSQL = strSQL & "    AND   YOUKT_YUKBN IN(1, 2, 10) )"                   'DELETE 2005/12/15 K.ISHIZAKA   'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "    FROM  INTR_TRAN"                                     'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "    WHERE INTRT_YCODE    = CNTA_CODE"                    'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "    AND   INTRT_NO       = CNTA_NO"                      'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "    AND   INTRT_INTROKBN IN (1, 2) )"                    'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "GROUP BY CNTA_CODE"
    strSQL = strSQL & ") AS WK "
    strSQL = strSQL & "ON CNTA_CODE = YARD_CODE"

    fncUpdateTempTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ更新ＳＱＬ作成（確定数）
'       MODULE_ID       : fncUpdateTempKakutei
'       CREATE_DATE     : 2005/10/06        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'>>> DELETE START 2005/12/05 K.ISHIZAKA >>>
'Private Function fncUpdateTempKakutei() As String
'    Dim strSQL              As String
'
'    strSQL = strSQL & "UPDATE"
'    strSQL = strSQL & "    #WK_FVS011_TEMP "
'    strSQL = strSQL & "SET"
'    strSQL = strSQL & "    UP_KAKUTEI = UP_KAKU,"
'    strSQL = strSQL & "    DW_KAKUTEI = DW_KAKU "
'    strSQL = strSQL & "FROM"
'    strSQL = strSQL & "    #WK_FVS011_TEMP "
'    strSQL = strSQL & "INNER JOIN"
'    strSQL = strSQL & "("
'    strSQL = strSQL & "SELECT"
'    strSQL = strSQL & "    CNTA_CODE,"
'    strSQL = strSQL & "    COUNT(NULLIF(CNTA_STEP, 1)) AS UP_KAKU,"
'    strSQL = strSQL & "    COUNT(NULLIF(CNTA_STEP, 0)) AS DW_KAKU "
'    strSQL = strSQL & "FROM"
'    strSQL = strSQL & "    CNTA_MAST "
'    strSQL = strSQL & "INNER JOIN"
'    strSQL = strSQL & "    YOUK_TRAN "
'    strSQL = strSQL & "ON (   YOUKT_YCODE = CNTA_CODE "
'    strSQL = strSQL & "AND    YOUKT_NO    = CNTA_NO "
'    strSQL = strSQL & ")"
'    strSQL = strSQL & "WHERE"
'    strSQL = strSQL & "    CNTA_CODE IN (SELECT YARD_CODE FROM #WK_FVS011_TEMP)" & strCNTA_SYU
'    strSQL = strSQL & "AND YOUKT_YUKBN = 1 "
'    strSQL = strSQL & "GROUP BY CNTA_CODE"
'    strSQL = strSQL & ") AS WK "
'    strSQL = strSQL & "ON CNTA_CODE  =  YARD_CODE"
'
'    fncUpdateTempKakutei = strSQL
'End Function
'<<< DELETE END   2005/12/05 K.ISHIZAKA <<<
'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ更新ＳＱＬ作成（解約数）
'       MODULE_ID       : fncUpdateTempKaiyaku
'       CREATE_DATE     : 2005/10/06        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateTempKaiyaku() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & "    #WK_FVS011_TEMP "
    strSQL = strSQL & "SET"
    strSQL = strSQL & "    KAIYAKU = KAI "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "    #WK_FVS011_TEMP "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "    CNTA_CODE,"
    strSQL = strSQL & "    COUNT(*) AS KAI "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "    CARG_FILE "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "    CNTA_MAST "
    strSQL = strSQL & "ON (   CNTA_CODE = CARG_YCODE "
    strSQL = strSQL & "AND    CNTA_NO   = CARG_NO )"
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & "    CARG_YCODE IN (SELECT YARD_CODE FROM #WK_FVS011_TEMP)" & strCNTA_SYU
    strSQL = strSQL & "AND CARG_CONTNA = 0 "
    strSQL = strSQL & "AND CARG_KYDATE IS NOT NULL "
    
    strSQL = strSQL & "AND NOT EXISTS"                                          'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "   (SELECT *"                                            'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "    FROM  INTR_TRAN"                                     'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "    WHERE INTRT_YCODE    = CNTA_CODE"                    'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "    AND   INTRT_NO       = CNTA_NO"                      'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "    AND   INTRT_INTROKBN IN (1, 2) )"                    'INSERT 2006/01/22 K.ISHIZAKA

    strSQL = strSQL & "GROUP BY CNTA_CODE"
    strSQL = strSQL & ") AS WK "
    strSQL = strSQL & "ON CNTA_CODE  =  YARD_CODE"

    fncUpdateTempKaiyaku = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一時作業用データ更新ＳＱＬ作成（取置数）
'       MODULE_ID       : fncUpdateTempTorioki
'       CREATE_DATE     : 2005/10/06        K.ISHIZAKA
'       RETURN          : SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateTempTorioki() As String
    Dim strSQL              As String
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & "    #WK_FVS011_TEMP "
    strSQL = strSQL & "SET"
    strSQL = strSQL & "    TORIOKI = TORI "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "    #WK_FVS011_TEMP "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & "    CNTA_CODE,"
    strSQL = strSQL & "    COUNT(*) AS TORI "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "    CNTA_MAST "
    strSQL = strSQL & "INNER JOIN"
'    strSQL = strSQL & "    YOUK_TRAN "                                         'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "ON (   YOUKT_YCODE = CNTA_CODE "                        'DELETE 2005/12/15 K.ISHIZAKA
'    strSQL = strSQL & "AND    YOUKT_NO    = CNTA_NO "                          'DELETE 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "    INTR_TRAN "                                          'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "ON (   INTRT_YCODE = CNTA_CODE "                         'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "AND    INTRT_NO    = CNTA_NO "                           'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & ")"
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "    YOUK_TRAN "                                          'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "ON (   YOUKT_UKNO = INTRT_UKNO "                         'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & ")"                                                       'INSERT 2006/01/22 K.ISHIZAKA
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & "    CNTA_CODE IN (SELECT YARD_CODE FROM #WK_FVS011_TEMP)" & strCNTA_SYU
    strSQL = strSQL & "AND YOUKT_YUKBN = 2 "                                    '復活   2006/01/22 K.ISHIZAKA   'DELETE 2005/12/05 K.ISHIZAKA
'    strSQL = strSQL & "AND YOUKT_YUKBN IN(1, 2) "                              'DELETE 2005/12/15 K.ISHIZAKA   'INSERT 2005/12/05 K.ISHIZAKA
    strSQL = strSQL & "AND INTRT_INTROKBN = 1 "                                 'INSERT 2005/12/15 K.ISHIZAKA
    strSQL = strSQL & "GROUP BY CNTA_CODE"
    strSQL = strSQL & ") AS WK "
    strSQL = strSQL & "ON CNTA_CODE  =  YARD_CODE"

    fncUpdateTempTorioki = strSQL
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
Private Function fncSelectTempTable(ByVal bAll As Boolean) As String
    Dim strSQL              As String
    
    If bAll Then
        If Nz(Me.cmb_YardCode.VALUE, "") = "" Then  '2008/03/23 add tajima
            strSQL = strSQL & "SELECT"
            strSQL = strSQL & " * "
            strSQL = strSQL & "FROM #WK_FVS011_TEMP "
           ' strSQL = strSQL & "WHERE SUM_AKI > 0 "                                 'DEL    2017/08/02 M.HONDA
            strSQL = strSQL & "ORDER BY"                                            'INSERT 2006/03/31 K.ISHIZAKA
            strSQL = strSQL & " SERVICE DESC,"                                      'INSERT 2006/03/31 K.ISHIZAKA
            strSQL = strSQL & " [UP_AKI]+[DW_AKI] DESC"                             'INSERT 2006/03/31 K.ISHIZAKA
        '▼2008/03/23 add tajima
        '  ヤードコード指定の場合は基本距離順とする
        Else
            strSQL = strSQL & "SELECT"
            strSQL = strSQL & " * "
            strSQL = strSQL & ",CASE WHEN YARD_CODE = " & CLng(Me.cmb_YardCode.VALUE)
            strSQL = strSQL & " THEN 1 "            '自ヤードが先頭
            strSQL = strSQL & " WHEN SERVICE = 1 "
            strSQL = strSQL & " THEN 2 "            'キャンペーンが次
            strSQL = strSQL & " ELSE 3 "            'それ以外はその他
            strSQL = strSQL & " END SORT_VALUE "
            strSQL = strSQL & "FROM #WK_FVS011_TEMP "
          '  strSQL = strSQL & "WHERE SUM_AKI > 0 "                                    'DEL    2017/08/02 M.HONDA
            strSQL = strSQL & "ORDER BY"
            strSQL = strSQL & " SORT_VALUE,"
            strSQL = strSQL & " YARD_KIRO "
        End If
        '▲2008/03/23 add tajima
        
    Else
        strSQL = strSQL & "SELECT"
        strSQL = strSQL & " YARD_CODE,"
        strSQL = strSQL & " UP_AKI,"
        strSQL = strSQL & " DW_AKI,"
'        strSQL = strSQL & " UP_KAKUTEI,"                                       'DELETE 2005/12/05 K.ISHIZAKA
'        strSQL = strSQL & " DW_KAKUTEI,"                                       'DELETE 2005/12/05 K.ISHIZAKA
        strSQL = strSQL & " MIKAKUTEI,"
        strSQL = strSQL & " KAIYAKU,"                                           'INSERT 20050829 K.ISHIZAKA
        strSQL = strSQL & " TORIOKI,"                                           'INSERT 20050829 K.ISHIZAKA
        strSQL = strSQL & " SUM_AKI "
        strSQL = strSQL & "FROM #WK_FVS011_TEMP "
    End If
    fncSelectTempTable = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一覧データ検索処理
'       MODULE_ID       : fncGetData
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : strYardCode       ヤードコード
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncGetData() As Boolean                                       'DELETE 20050829 K.ISHIZAKA
Private Function fncGetData(ByVal strYardCode As String) As Boolean             'INSERT 20050829 K.ISHIZAKA
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim dbAccess            As Database
    Dim rsAccess            As Recordset
    Dim i                   As Integer
    Dim iNo                 As Long                                             'INSERT 2006/03/31 K.ISHIZAKA
    Dim strKey              As Variant                                          'INSERT 2009/12/02 K.ISHIZAKA
    On Error GoTo Err_Handler
    
    bKaiyaku = Me.chk_KAIYAKU                                                   'INSERT 20050829 K.ISHIZAKA
    bTorioki = Me.chk_TORIOKI                                                   'INSERT 20050829 K.ISHIZAKA
    
    'Access接続                                                                 'INSERT START 2009/12/02 K.ISHIZAKA
    Set dbAccess = CurrentDb()
    'Access側ワークテーブル削除
    Call subDropWorkTable(dbAccess)
    'Access側ワークテーブル作成
    If fncCreateWorkTable(dbAccess) = False Then
        dbAccess.Close
        Set dbAccess = Nothing
        fncGetData = False
        Exit Function
    End If
    For Each strKey In colKey
        strConnection = colConnection(strKey)                                   'INSERT END   2009/12/02 K.ISHIZAKA
    
    'SQL-Server接続
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(strConnection)
    '詳細データ作成
    dbSQLServer.Execute fncCreateTempTable(), adCmdText
    dbSQLServer.Execute fncCreateTempIndex(), adCmdText                         'INSERT 2006/01/20 K.ISHIZAKA
'    dbSqlserver.Execute fncInsertTempTable("", ""), adCmdText                                  'DELETE 20050829 K.ISHIZAKA
    dbSQLServer.Execute fncInsertTempTable(strYardCode, "", Me.chk_NearYard.VALUE), adCmdText   'INSERT 20050829 K.ISHIZAKA
    dbSQLServer.Execute fncUpdateTempTable(), adCmdText
'    dbSqlServer.Execute fncUpdateTempKakutei(), adCmdText                      'DELETE 2005/12/05 K.ISHIZAKA   'INSERT 20051006 K.ISHIZAKA
    If bKaiyaku Then                                                            'INSERT 20051006 K.ISHIZAKA
        dbSQLServer.Execute "UPDATE #WK_FVS011_TEMP SET KAIYAKU = 0", adCmdText 'INSERT 2005/12/05 K.ISHIZAKA
        dbSQLServer.Execute fncUpdateTempKaiyaku(), adCmdText                   'INSERT 20051006 K.ISHIZAKA
    End If                                                                      'INSERT 20051006 K.ISHIZAKA
    If bTorioki Then                                                            'INSERT 20051006 K.ISHIZAKA
        dbSQLServer.Execute "UPDATE #WK_FVS011_TEMP SET TORIOKI = 0", adCmdText 'INSERT 2005/12/05 K.ISHIZAKA
        dbSQLServer.Execute fncUpdateTempTorioki(), adCmdText                   'INSERT 20051006 K.ISHIZAKA
    End If                                                                      'INSERT 20051006 K.ISHIZAKA
    'データ取得
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    rsSqlServer.Open fncSelectTempTable(True), dbSQLServer, adOpenForwardOnly, adLockReadOnly
'    'Access接続                                                                'DELETE START 2009/12/02 K.ISHIZAKA
'    Set dbAccess = CurrentDb()
'    'Access側ワークテーブル削除
'    Call subDropWorkTable(dbAccess)
'    'Access側ワークテーブル作成
'    If fncCreateWorkTable(dbAccess) = False Then
'        dbAccess.Close
'        Set dbAccess = Nothing
'        'SQL-Server側切断
'        rsSqlServer.Close
'        Set rsSqlServer = Nothing
'        dbSqlServer.Close
'        Set dbSqlServer = Nothing
'        fncGetData = False
'        Exit Function
'    End If                                                                     'DELETE END   2009/12/02 K.ISHIZAKA
    iNo = 0                                                                     'INSERT 2006/03/31 K.ISHIZAKA
    'Access側ワークテーブル
    Set rsAccess = dbAccess.OpenRecordset("WK_FVS011_TEMP", dbOpenDynaset, dbAppendOnly)
    With rsSqlServer
        Do While Not .EOF
            iNo = iNo + 1                                                       'INSERT 2006/03/31 K.ISHIZAKA
            rsAccess.AddNew
            rsAccess.Fields("NO") = iNo
            For i = 0 To .Fields.Count - 1
                With .Fields(i)
                    rsAccess.Fields(.NAME) = .VALUE
                End With
            Next
            If Nz(rsAccess.Fields("YARD_BEGIN_DAY"), DATE) > DATE Then          'INSERT 2006/08/31 K.ISHIZAKA
                rsAccess.Fields("YARD_BEGIN_DAY2") = Format(rsAccess.Fields("YARD_BEGIN_DAY"), "yy/mm") & "～"      'INSERT 2006/08/31 K.ISHIZAKA
            End If                                                              'INSERT 2006/08/31 K.ISHIZAKA
            If Not IsNull(rsAccess.Fields("YARD_RENTEND_DAY")) Then             'INSERT 2006/08/31 K.ISHIZAKA
                rsAccess.Fields("YARD_RENTEND_DAY2") = "～" & Format(rsAccess.Fields("YARD_RENTEND_DAY"), "yy/mm")  'INSERT 2006/08/31 K.ISHIZAKA
            End If                                                              'INSERT 2006/08/31 K.ISHIZAKA
            rsAccess.Fields("KEY") = strKey                                     'INSERT 2009/12/02 K.ISHIZAKA
            rsAccess.UPDATE
            .MoveNext
        Loop
        'SQL-Server側切断
        .Close
    End With
    Set rsSqlServer = Nothing
    dbSQLServer.Close
    Set dbSQLServer = Nothing
    
    Next                                                                        'INSERT 2009/12/02 K.ISHIZAKA
    
    'Access側切断
    rsAccess.Close
    Set rsAccess = Nothing
    dbAccess.Close
    Set dbAccess = Nothing
    With Me.一覧サブ.Form                                                       'INSERT START 2009/12/02 K.ISHIZAKA
        If Nz(Me.cmb_YardCode.VALUE, "") = "" Then
            .OrderBy = "SERVICE, [UP_AKI]+[DW_AKI] DESC"
        Else
            .OrderBy = "SORT_VALUE, YARD_KIRO"
        End If
        .OrderByOn = True
    End With                                                                    'INSERT END   2009/12/02 K.ISHIZAKA
    Me.一覧サブ.Requery
    Me.一覧サブ.SetFocus
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
'       MODULE_NAME     : 一覧データ検索処理
'       MODULE_ID       : fncCheck_CNTA_SIZE
'       CREATE_DATE     : 2005/11/17        K.ISHIZAKA
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheck_CNTA_SIZE() As Boolean
    If fncCheck_CNTA_SIZE_1(Nz(Me.txt_SizeFrom)) Then
        fncCheck_CNTA_SIZE = fncCheck_CNTA_SIZE_1(Nz(Me.txt_SizeTo))
    Else
        fncCheck_CNTA_SIZE = False
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : 一覧データ検索処理
'       MODULE_ID       : fncCheck_CNTA_SIZE_1
'       CREATE_DATE     : 2005/11/17        K.ISHIZAKA
'       PARAM           : strSize           サイズ
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheck_CNTA_SIZE_1(ByVal strSize As String) As Boolean
    If Trim(strSize) <> "" Then
        If (Val(strSize) <= 0) Or (Val(strSize) >= 1000) Then
            fncCheck_CNTA_SIZE_1 = False
            Exit Function
        End If
    End If
    fncCheck_CNTA_SIZE_1 = True
End Function

'==============================================================================*
'
'       MODULE_NAME     : サイズ条件作成処理
'       MODULE_ID       : fncWhere_CNTA_SIZE
'       CREATE_DATE     : 2005/11/17        K.ISHIZAKA
'       RETURN          : サイズ条件SQL
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncWhere_CNTA_SIZE() As String
    Dim strSizeFrom         As String
    Dim strSizeTo           As String
    
    strSizeFrom = Trim(Nz(Me.txt_SizeFrom))
    strSizeTo = Trim(Nz(Me.txt_SizeTo))
    '両方とも空
    If (strSizeFrom = "") And (strSizeTo = "") Then
        fncWhere_CNTA_SIZE = ""
        Exit Function
    End If
    'どっちか空
    If (strSizeFrom = "") Or (strSizeTo = "") Then
        fncWhere_CNTA_SIZE = " AND CNTA_SIZE = " & strSizeFrom & strSizeTo
        Exit Function
    End If
    '一緒
    If strSizeFrom = strSizeTo Then
        fncWhere_CNTA_SIZE = " AND CNTA_SIZE = " & strSizeFrom
        Exit Function
    End If
    '両方入力
    If Val(strSizeFrom) < Val(strSizeTo) Then
        fncWhere_CNTA_SIZE = " AND CNTA_SIZE BETWEEN " & strSizeFrom & " AND " & strSizeTo
    Else
        fncWhere_CNTA_SIZE = " AND CNTA_SIZE BETWEEN " & strSizeTo & " AND " & strSizeFrom
    End If
    
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
    Select Case KeyCode
    Case vbKeyF1
        KeyCode = 0
        Me.cmd_終了.SetFocus
        'Call cmd_終了_Click
    Case vbKeyF3, vbKeyF5 To vbKeyF12
        KeyCode = 0
    End Select
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 詳細画面表示
'       MODULE_ID       : ShowFVS021
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : strYardCode       ヤードコード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Sub ShowFVS021(ByVal strYardCode As String)
Public Sub ShowFVS021(ByVal strKey As String, ByVal strYardCode As String)
    If Not fncFormIsLoaded("FKS000") Then                                       'INSERT START 2009/12/02 K.ISHIZAKA
        If fncFormIsLoadedMsg("FVS400", Me.Caption) Then
            Exit Sub
        End If
        If fncFormIsLoadedMsg("FVS550", Me.Caption) Then
            Exit Sub
        End If
        If fncFormIsLoadedMsg("FVS500", Me.Caption) Then
            Exit Sub
        End If
        If fncFormIsLoadedMsg("FVS021", Me.Caption) Then
            Exit Sub
        End If
        If DLookup("CONT_BUMOC", "dbo_CONT_MAST") <> strKey Then
            Call MSHS012_M10(strKey)
        End If
    End If
    strArgments = colArgments(strKey)                                           'INSERT END   2009/12/02 K.ISHIZAKA
    '引数は ヤードコード、コンテナ種別条件、部門コード、WEBパス、EXCELパス、SQL-Server接続文字列、ユーザー名
    doCmd.OpenForm "FVS021", , , , , , Join(Array(Format(Val(strYardCode), "000000"), strCNTA_SYU, strArgments), ",")
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 一覧データ検索処理
'       MODULE_ID       : Reload
'       CREATE_DATE     : 2005/07/28        K.ISHIZAKA
'       PARAM           : strYard1          ヤードコード１
'                       : strYard2          ヤードコード２
'       RETURN          : 正常(True)/エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function Reload(ByVal strYard1 As String, ByVal strYard2 As String) As Boolean
Public Function Reload(ByVal strKey As String, ByVal strYard1 As String, ByVal strYard2 As String) As Boolean
    Dim dbSQLServer         As Object       'ADODB.Connection
    Dim rsSqlServer         As Object       'ADODB.Recordset
    Dim dbAccess            As Database
    Dim rsAccess            As Recordset
    Dim i                   As Integer
    On Error GoTo Err_Handler
    
    strConnection = colConnection(strKey)                                       'INSERT 2009/12/02 K.ISHIZAKA
    'ヤードコードの先頭ゼロを削除
    strYard1 = Format(Val(strYard1))
    strYard2 = Format(Val(strYard2))
    'Access接続
    Set dbAccess = CurrentDb()
'    Set rsAccess = dbAccess.OpenRecordset(fncSelectWorkTable(strYard1, strYard2), dbOpenDynaset)   'DELETE 2009/12/02 K.ISHIZAKA
    Set rsAccess = dbAccess.OpenRecordset(fncSelectWorkTable(strKey, strYard1, strYard2), dbOpenDynaset) 'INSERT 2009/12/02 K.ISHIZAKA
    If rsAccess.EOF Then
        rsAccess.Close
        Set rsAccess = Nothing
        dbAccess.Close
        Set dbAccess = Nothing
        Reload = True
        Exit Function
    End If
    'SQL-Server接続
    Set dbSQLServer = CreateObject("ADODB.Connection")
    dbSQLServer.CommandTimeout = 180
    Call dbSQLServer.Open(strConnection)
    '詳細データ作成
    dbSQLServer.Execute fncCreateTempTable(), adCmdText
    dbSQLServer.Execute fncCreateTempIndex(), adCmdText                         'INSERT 2006/01/20 K.ISHIZAKA
'    dbSqlserver.Execute fncInsertTempTable(strYard1, strYard2), adCmdText          'DELETE 20050829 K.ISHIZAKA
    dbSQLServer.Execute fncInsertTempTable(strYard1, strYard2, False), adCmdText    'INSERT 20050829 K.ISHIZAKA
    dbSQLServer.Execute fncUpdateTempTable(), adCmdText
'    dbSqlServer.Execute fncUpdateTempKakutei(), adCmdText                      'DELETE 2005/12/05 K.ISHIZAKA   'INSERT 20051006 K.ISHIZAKA
    If bKaiyaku Then                                                            'INSERT 20051006 K.ISHIZAKA
        dbSQLServer.Execute "UPDATE #WK_FVS011_TEMP SET KAIYAKU = 0", adCmdText 'INSERT 2005/12/05 K.ISHIZAKA
        dbSQLServer.Execute fncUpdateTempKaiyaku(), adCmdText                   'INSERT 20051006 K.ISHIZAKA
    End If                                                                      'INSERT 20051006 K.ISHIZAKA
    If bTorioki Then                                                            'INSERT 20051006 K.ISHIZAKA
        dbSQLServer.Execute "UPDATE #WK_FVS011_TEMP SET TORIOKI = 0", adCmdText 'INSERT 2005/12/05 K.ISHIZAKA
        dbSQLServer.Execute fncUpdateTempTorioki(), adCmdText                   'INSERT 20051006 K.ISHIZAKA
    End If                                                                      'INSERT 20051006 K.ISHIZAKA
    'データ取得
    Set rsSqlServer = CreateObject("ADODB.Recordset")
    rsSqlServer.Open fncSelectTempTable(False), dbSQLServer, adOpenForwardOnly, adLockReadOnly
    With rsSqlServer
        Do While Not .EOF
'            If rsAccess.Fields("YARD_CODE") = rsSqlServer("YARD_CODE") Then    'DELETE 2009/12/02 K.ISHIZAKA
            If (rsAccess.Fields("YARD_CODE") = rsSqlServer("YARD_CODE")) And (rsAccess.Fields("KEY") = strKey) Then 'INSERT 2009/12/02 K.ISHIZAKA
                rsAccess.Edit
                For i = 0 To .Fields.Count - 1
                    With .Fields(i)
                        rsAccess.Fields(.NAME) = .VALUE
                    End With
                Next
                rsAccess.UPDATE
                rsAccess.MoveNext
                If rsAccess.EOF Then
                    Exit Do
                End If
            End If
            .MoveNext
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
    Me.一覧サブ.Requery
    Reload = True
    Exit Function

Err_Handler:
    MsgBox "Reload:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
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
    Reload = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : 地図表示
'       MODULE_ID       : MapOpen
'       CREATE_DATE     : 2006/03/31        K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub MapOpen()
    Dim rsAccess            As Recordset
    Dim IEApp               As Object
    Dim Elmnt               As Object
    Dim strMark             As String
    On Error GoTo Err_Handler

    Set rsAccess = CurrentDb().OpenRecordset("SELECT * FROM WK_FVS011_TEMP ORDER BY NO DESC", dbOpenForwardOnly)
    With rsAccess
        If .EOF Then
            .Close
            Exit Sub
        End If
        Call MNN020.WebGoogleOpen(IEApp, True)
        While Not .EOF
            If Nz(.Fields("SERVICE")) = -1 Then
                strMark = "r"
            ElseIf (.Fields("UP_AKI") + .Fields("DW_AKI")) >= 10 Then
                strMark = "b"
            Else
                strMark = "g"
            End If
            If .Fields("NO") < 25 Then
                strMark = strMark & Format(.Fields("NO"))
            End If
            If (Nz(.Fields("YARD_IDO")) > 0) And (Nz(.Fields("YARD_KEIDO")) > 0) Then
                Call MNN020.WebGoogleMarker(IEApp, .Fields("YARD_IDO"), .Fields("YARD_KEIDO"), _
                    strMark, .Fields("YARD_CODE"), .Fields("YARD_NAME"), _
                    Nz(.Fields("YARD_ADDR_1")) & Nz(.Fields("YARD_ADDR_2")) & Nz(.Fields("YARD_ADDR_3")), Nz(.Fields("YARD_NOTE")))
            End If
            .MoveNext
        Wend
        .Close
    End With
    Set Elmnt = IEApp.Document.getElementById("divHelp")
    Elmnt.Style.Display = ""
Exit Sub

Err_Handler:
    rsAccess.Close
    MsgBox "MapOpen:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
End Sub

'############################ 2007/05/21 Add tajima ############################
'==============================================================================*
'
'       MODULE_NAME     : キー入力時処理
'       MODULE_ID       : cmb_Usage_KeyPress
'       CREATE_DATE     : 2007/05/21 tajima
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Usage_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Usage, KeyAscii, 2)
End Sub
'==============================================================================*
'
'       MODULE_NAME     : 更新後処理
'       MODULE_ID       : cmb_Usage_AfterUpdate
'       CREATE_DATE     : 2007/05/21 tajima
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Usage_AfterUpdate()
    
    With Me.cmb_Usage
        If .ListIndex >= 0 Then
            .VALUE = Format(.Text, .Format)
            Me.txt_Usage.VALUE = .Column(1, .ListIndex)
        Else
            Me.txt_Usage.VALUE = "" '未入力なら用途文言クリア
        End If
    End With
    
    ' サイズ＆レンタル用途の取得条件を変更
    strCNTA_SYU = fncWhere_CNTA_SIZE() & fncWhere_CNTA_USAGE()

End Sub
'==============================================================================*
'
'       MODULE_NAME     : 更新前処理
'       MODULE_ID       : cmb_Usage_BeforeUpdate
'       CREATE_DATE     : 2007/05/21 tajima
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Usage_BeforeUpdate(Cancel As Integer)
    
  If Nz(Me.cmb_Usage.VALUE, "") = "" Then Exit Sub
    
  If Me.cmb_Usage.ListIndex < 0 Then
       MsgBox "そのコードは登録されていません。", vbInformation, "確認"
       Cancel = True
       SendKeys "{ESC}"
  End If

End Sub
'==============================================================================*
'
'       MODULE_NAME     : レンタル用途条件作成処理
'       MODULE_ID       : fncWhere_CNTA_USAGE
'       CREATE_DATE     : 2007/05/21 tajima
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncWhere_CNTA_USAGE()

    'NOTE 空詳細へのパラメータになる関係でIN(x,y,z)は使えない※カンマが×
    Select Case Nz(Me.cmb_Usage.VALUE, "")
        '条件無し
        Case ""
              fncWhere_CNTA_USAGE = ""
        
        'ボックスorトランク
        Case "10"
              fncWhere_CNTA_USAGE = " AND (CNTA_USAGE = 0 OR CNTA_USAGE = 1) "
              
        'バイク関連全て
        Case "30"
              fncWhere_CNTA_USAGE = " AND (CNTA_USAGE = 3 OR CNTA_USAGE = 31 OR CNTA_USAGE = 32) "
        
        ' その他コード特定
        Case Else
              fncWhere_CNTA_USAGE = " AND CNTA_USAGE = " & Me.cmb_Usage.VALUE
    
    End Select
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : 引数リスト作成処理
'       MODULE_ID       : makeArgmentsCollection
'       CREATE_DATE     : 2009/12/02        K.ISHIZAKA
'       PARAM           : strCONT_BUMOC     部門コード(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub makeArgmentsCollection(ByVal strCONT_BUMOC As String)
    Dim objCon              As Object
    Dim objRst              As Object
    Dim strSQL              As String
    Dim strCONT_DISKBACK    As String
    Dim strCONT_EXCEL_PATH  As String
    On Error GoTo ErrorHandler
    
    strSQL = "SELECT CONT_DISKBACK, CONT_EXCEL_PATH FROM CONT_MAST WHERE CONT_KEY = 1"
    Set objCon = ADODB_Connection(strCONT_BUMOC)
    On Error GoTo ErrorHandler1
    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler2
    With objRst
        If Not .EOF Then
            strCONT_DISKBACK = Trim(Nz(.Fields("CONT_DISKBACK")))
            strCONT_EXCEL_PATH = Trim(Nz(.Fields("CONT_EXCEL_PATH")))
        Else
            strCONT_DISKBACK = ""
            strCONT_EXCEL_PATH = ""
        End If
        .Close
    End With
    If strCONT_DISKBACK = "" Then
        Me.cmd_Serach.Enabled = False
        MsgBox "[CONT_MAST]の設定不備です。"
        Exit Sub
    End If
    If strCONT_EXCEL_PATH <> "" Then
        If Right(strCONT_EXCEL_PATH, 1) <> "\" Then
            strCONT_EXCEL_PATH = strCONT_EXCEL_PATH & "\"
        End If
    End If
    strConnection = MSZZ007_M10(strCONT_BUMOC)
    strArgments = Join(Array(strCONT_BUMOC, strCONT_DISKBACK, strCONT_EXCEL_PATH, strConnection, LsGetUserName()), ",")
    
    colKey.Add strCONT_BUMOC, strCONT_BUMOC
    colConnection.Add strConnection, strCONT_BUMOC
    colArgments.Add strArgments, strCONT_BUMOC
Exit Sub

ErrorHandler2:
    objRst.Close
ErrorHandler1:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "set_cmb_Usage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : レンタル用途設定処理
'       MODULE_ID       : set_cmb_Usage
'       CREATE_DATE     : 2009/12/02        K.ISHIZAKA
'       PARAM           : strBumoc          部門コード(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub set_cmb_Usage(ByVal strBUMOC As String)
    Dim objCon              As Object
    Dim objRst              As Object
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = "SELECT NAME_CODE, NAME_NAME FROM NAME_MAST WHERE NAME_ID ='086' ORDER BY NAME_CODE"
    Set objCon = ADODB_Connection(strBUMOC)
    On Error GoTo ErrorHandler1
    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler2
    strSQL = ""
    With objRst
        While Not .EOF
            strSQL = strSQL & Format(.Fields("NAME_CODE"), "00") & ";"
            strSQL = strSQL & .Fields("NAME_NAME") & ";"
            .MoveNext
        Wend
        .Close
    End With
    On Error GoTo ErrorHandler1
    objCon.Close
    On Error GoTo ErrorHandler
    With cmb_Usage
        .RowSourceType = "Value List"
        .RowSource = strSQL
    End With
Exit Sub

ErrorHandler2:
    objRst.Close
ErrorHandler1:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "set_cmb_Usage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'****************************  ended or program ********************************
