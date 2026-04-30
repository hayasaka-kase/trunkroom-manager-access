VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS903"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード修繕一覧画面
'   プログラムＩＤ　：　FVS901
'   作　成　日　　　：  2006/05/13
'   作　成　者　　　：  T.SUZUKI
'   Ver            ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：2006/09/06
'   修　正　者　　　：tajima
'   修　正　内　容　：1.操作権限変更
'                   ：2.修理中＆撤去中を削除したらコンテナの情報も戻す
'                   ：3.ヤードの開始日～終了日を出す
'   Ver             ：0.1

'   修　正　日　　　：2007/01/26
'   修　正　者　　　：tajima@eagle-soft
'   修　正　内　容　：①ヤード修理入力機能追加対応
'   Ver             ：0.2
'
'   修　正　日　　　：2010/06/10
'   修　正　者　　　：SHIBAZAKI
'   修　正　内　容　：顧客コメントクレーム履歴検索時、表示区分が「表示」のレコードのみ検索対象とする。
'   Ver             ：0.3
'
'   修　正　日　　　：2010/12/22
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：修理時に完了日付を出すように修正。
'   Ver             ：0.4
'
'   修　正　日　　　：2013/06/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：終了メッセージを表示させないように修正。
'   Ver             ：0.5
'
'   修　正　日　　　：2013/11/02
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：部屋修繕時に鍵情報を出力
'   Ver             ：0.6
'
'   修　正　日　　　：2013/11/15
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：部屋修繕時に鍵情報を出力のSQLを修正
'   Ver             ：0.7
'
'   修　正　日　　　：2014/09/29
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：部屋修繕時に鍵情報を出力のSQLを修正
'   Ver             ：0.8
'
'   修　正　日　　　：2015/03/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：鍵メンテを自動発番に対応
'   Ver             ：0.9
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   Private定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_FORM_ID_FVS901        As String = "FVS901"                      ' ヤード修繕一覧画面
Private Const P_FORM_ID_FVS902        As String = "FVS902"                      ' 顧客修繕一覧画面
Private Const P_FORM_ID_FVS903        As String = "FVS903"                      ' コンテナ(部屋)修繕一覧画面
Private Const P_FORM_ID_FVS910        As String = "FVS910"                      ' 修繕入力ダイアログ

Private Const P_FORM_TITTLE_FVS901    As String = "ヤード修繕"
Private Const P_FORM_TITTLE_FVS902    As String = "顧客ｺﾒﾝﾄ･ｸﾚｰﾑ履歴"
Private Const P_FORM_TITTLE_FVS903    As String = "コンテナ(部屋)修繕"

Private Const P_WK_TABLE_NM           As String = "FVS901_W01"                  ' ワークテーブル名
Private Const P_SUBFRM_NM             As String = "FVS901_S01"                  ' サブフォーム名

Private Const P_DB_KONT               As Integer = 1                            ' コンテナDB
Private Const P_DB_KASE               As Integer = 2                            ' 加瀬DB

Private Const P_NAME_ID_090           As String = "090"                         ' 名称取得ID(段区分)
Private Const P_NAME_ID_040           As String = "040"                         ' 名称取得ID(種別)
Private Const P_NAME_ID_250           As String = "250"                         ' 名称取得ID(区分)

Private Const P_MSG_ERR               As String = "エラー"
Private Const P_MSG_KAKUNIN           As String = "確認"

Private Const P_MSG_ERR_PARAM         As String = "起動パラメータが正しくありません。"
Private Const P_MSG_ERR_INPUT_KETA    As String = "桁数オーバーしています。"
Private Const P_MSG_ERR_INPUT_VALUE   As String = "コードを入力してください。"
Private Const P_MSG_ERR_INPUT_NAIYO   As String = "内容を正しく入力してください。"

Private Const P_MSG_ERR_NODATA_YARD   As String = "該当するヤードマスタが登録がされておりません。"
Private Const P_MSG_ERR_NODATA_CNTA   As String = "該当するコンテナマスタ登録がされておりません。"
Private Const P_MSG_ERR_NODATA_USER   As String = "該当する顧客マスタが登録されておりません。"

Private Const P_MSG_ERR_INPUT_YARD    As String = "ヤードコードを選択して下さい。"
Private Const P_MSG_ERR_INPUT_CNTA    As String = "コンテナ番号を選択して下さい。"
Private Const P_MSG_ERR_INPUT_USER    As String = "顧客コードを選択して下さい。"

Private Const P_MSG_END_KAKUNIN       As String = "終了します。よろしいですか？"

Private Const P_MSG_MODE_INS          As String = "1"                           ' 登録
Private Const P_MSG_MODE_UPD          As String = "2"                           ' 修正

'==============================================================================*
'   Private変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private pstrOpenForm                  As String                                 ' 「FVS901」OR「FVS902」OR「FVS903」

Private pstrBumonCd                   As String                                 ' 部門コード

Private pobjKONT_DB                   As Database                               ' コンテナDBオブジェクト
Private pobjKASE_DB                   As Database                               ' 加瀬DBオブジェクト

Private pstrYardCode                  As String   ' パラメータで渡ってきたヤードコード
Private pstrRoomNo                    As String   ' パラメータで渡ってきた部屋番号
Private pstrUserCode                  As String   ' パラメータで渡ってきた顧客コード
Private pstrAcptNo                    As String   ' パラメータで渡ってきた契約番号

Private pstrUserName                  As String   ' ユーザID
Private pIsAuthority                  As Boolean  ' FVS900更新可否権限
Private pIsAutoRead                   As Boolean  ' 自動読込可否
Private pIsChanged                    As Boolean  ' 変更可否フラグ  2006/09/06

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
    ' 適正サイズにする ヘルプによると 1cm は 567twip
    'doCmd.MoveSize , , 25.25 * 567, 18 * 567
    
    ' とりあえずMAXサイズで
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
    
    ' 画面項目初期設定
    Call initFormItems(pstrOpenForm)
    
    ' 初期ＤＢ読込
    errorMessage = initDBLoad()
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

    ' フォーム名をセット
    pstrOpenForm = Me.NAME   ' フォーム名
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
    'ヤード修繕の場合
    '② ヤードコード
    '顧客修繕の場合
    '③ 顧客コード
    '④ 顧客コード,契約番号
    '部屋修繕の場合
    '⑤ ヤードコード,部屋番号
    '⑥ ヤードコード,部屋番号,契約番号
    
'strParam = Split("000012,", ",") ' 新規 5/30確認
'strParam = Split("000012,870000,", ",") ' 新規 5/30確認
'strParam = Split("000010,", ",") ' 新規 5/30確認
    
    ' パラメータ数チェックと取込
    iParamCnt = UBound(strParam)
    
    pIsAutoRead = False       ' 引数指定があれば自動起動としたい
    Select Case pstrOpenForm
        ' ヤードの場合
        Case P_FORM_ID_FVS901
            If iParamCnt < 0 Then
                pstrYardCode = ""
            Else
                pstrYardCode = strParam(0)
                pIsAutoRead = True
            End If
        ' 顧客の場合
        Case P_FORM_ID_FVS902
            If iParamCnt < 0 Then
                pstrUserCode = ""
            Else
                pstrUserCode = strParam(0)
                If iParamCnt = 1 Then pstrAcptNo = strParam(1)  ' 契約番号付きなら契約番号取得
                pIsAutoRead = True
            End If
        ' コンテナ（部屋）の場合
        Case P_FORM_ID_FVS903
            If iParamCnt < 1 Then
                pstrYardCode = ""
                pstrRoomNo = ""
            Else
                pstrYardCode = strParam(0)
                pstrRoomNo = strParam(1)
                If iParamCnt = 2 Then pstrAcptNo = strParam(2)  ' 契約番号付きなら契約番号取得
                pIsAutoRead = True
            End If
      Case Else
            Call MSZZ024_M10("Form_Open", "画面ID" & pstrOpenForm & "は処理対象外です。")
    End Select
    
    Exit Sub
    
    pIsChanged = False  '2006/09/06
        
Exception:
    Call MSZZ024_M00("Form_Open", True)
    Cancel = True
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_End_Click イベント
'        MODULE_ID        :cmd_End_Click
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

  On Error GoTo Exception
  
    '' 2013/06/03 M.HONDA START
    ' 確認メッセージ表示
'    If MsgBox(P_MSG_END_KAKUNIN, vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
'        Exit Sub
'    End If
    '' 2013/06/03 M.HONDA END

    ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙﾃﾞｰﾀ削除
    Call subDeleteWork

    ' 部屋顧客修繕モードで変更があった場合、関係する他画面の再読込を指示する 2006/09/06
    If pIsChanged = True And (pstrOpenForm = P_FORM_ID_FVS903 Or pstrOpenForm = P_FORM_ID_FVS902) Then
        Call beforFormReload
    End If
    
    ' 本画面を閉じる
    doCmd.Close acForm, Me.FormName
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_End_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_ヤードコード_AfterUpdate イベント
'        MODULE_ID        :cmb_ヤードコード_AfterUpdate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_AfterUpdate()

    Dim varCode As Variant

  On Error GoTo Exception

    ' ヤードコード取得
    varCode = Nz(Me.cmb_ヤードコード, "")

    ' コンテナ関連の初期化
    Me.cmb_コンテナ番号 = ""
    Me.txt_段区分名 = ""
    Me.txt_サイズ = ""

    If varCode = "" Then Exit Sub

    Me.cmb_ヤードコード = Format(varCode, "000000")

    ' ヤードマスタチェック(データ取得)
    If fncCheck_Yard_Mast(1) = False Then Exit Sub

    Select Case pstrOpenForm
        Case P_FORM_ID_FVS901
            ' 修繕トランデータ設定処理
            Call subFormLoad
            
        Case P_FORM_ID_FVS903
            ' *** コンテナ(部屋)修繕一覧画面 ***
            ' コンテナ番号のリクエリー
            Call subCNTA_Requery
    End Select
    
    pIsChanged = False  '2006/09/06
    
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

    ' コードが空の場合入力区分欄も空欄表示
    Select Case pstrOpenForm
        Case P_FORM_ID_FVS901
            ' *** ヤード修繕一覧画面 ***
            If fncControlCrear(1) = False Then Exit Sub

        Case P_FORM_ID_FVS903
            ' *** コンテナ(部屋)修繕一覧画面 ***
            If fncControlCrear(2) = False Then Exit Sub
    End Select

    ' ヤードコード取得
    varCode = Nz(Me.cmb_ヤードコード, "")

    ' 入力チェック
    Select Case fncCheckNum(varCode, 6)
        Case 0
            ' 正常時
        Case 2
            ' 桁数エラー
            MsgBox P_MSG_ERR_INPUT_KETA, vbInformation, P_MSG_ERR
            Cancel = True: SendKeys "{ESC}"
            Exit Sub
        Case 9
            ' その他のエラー
            MsgBox P_MSG_ERR_INPUT_VALUE, vbInformation, P_MSG_ERR
            Cancel = True: SendKeys "{ESC}"
            Exit Sub
    End Select

    ' ヤードマスタチェック(存在チェック)
    If fncCheck_Yard_Mast(0) = False Then
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
'        MODULE_NAME      :cmb_コンテナ番号_AfterUpdate イベント
'        MODULE_ID        :cmb_コンテナ番号_AfterUpdate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_コンテナ番号_AfterUpdate()

    Dim varCode    As Variant

  On Error GoTo Exception

    ' コンテナ番号取得
    varCode = Nz(Me.cmb_コンテナ番号, "")
    Me.cmb_コンテナ番号 = Format(varCode, "000000")

    ' コンテナマスタチェック(データ取得)
    If fncCheck_Cnta_Mast(1) = False Then Exit Sub

    ' 修繕トランデータ設定処理
    Call subFormLoad
    
    pIsChanged = False  '2006/09/06
    
    Exit Sub
    
Exception:
    Call MSZZ024_M00("cmb_コンテナ番号_AfterUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_コンテナ番号_BeforeUpdate イベント
'        MODULE_ID        :cmb_コンテナ番号_BeforeUpdate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_コンテナ番号_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant

    On Error GoTo Exception

    ' コードが空の場合入力区分欄も空欄表示
    If fncControlCrear(3) = False Then Exit Sub

    ' ヤードコードが未入力の状態ではコンテナ番号は入力させない
    If Nz(Me.cmb_ヤードコード, "") = "" Then
        MsgBox "ヤードコードを選択して下さい。"
        Cancel = True: SendKeys "{ESC}"
        On Error Resume Next
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If

    ' コンテナ番号取得
    varCode = Nz(Me.cmb_コンテナ番号, "")

    ' 入力チェック
    Select Case fncCheckNum(varCode, 6)
        Case 0
            ' 正常時
        Case 2
            ' 桁数エラー
            MsgBox P_MSG_ERR_INPUT_KETA, vbInformation, P_MSG_ERR
            Cancel = True: SendKeys "{ESC}"
            Exit Sub
        Case 9
            ' その他のエラー
            MsgBox P_MSG_ERR_INPUT_VALUE, vbInformation, P_MSG_ERR
            Cancel = True: SendKeys "{ESC}"
            Exit Sub
    End Select

    ' コンテナマスタチェック(存在チェック)
    If fncCheck_Cnta_Mast(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If

    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_コンテナ番号_BeforeUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_コンテナ番号_KeyPress イベント
'        MODULE_ID        :cmb_コンテナ番号_KeyPress
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_コンテナ番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_コンテナ番号, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_顧客名_BeforeUpdate イベント
'        MODULE_ID        :txt_顧客名_BeforeUpdate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名_BeforeUpdate(Cancel As Integer)

  On Error GoTo Exception

    ' 文字数チェック
    If fncCheckText(Me.txt_顧客名, 18) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If

    Exit Sub

Exception:
    Call MSZZ024_M00("txt_顧客名_BeforeUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_AfterUpdate イベント
'        MODULE_ID        :cmb_顧客コード_AfterUpdate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_AfterUpdate()

    Dim varCode  As Variant

  On Error GoTo Exception

    varCode = Nz(Me.cmb_顧客コード.VALUE, "")
    If varCode = "" Then Exit Sub
    
    ' ファイルへの存在チェック
    If fncCheck_User_Mast(1) = False Then Exit Sub
    Me.cmb_顧客コード = Format(Me.cmb_顧客コード.VALUE, "000000")

    ' 顧客問合わせトランデータ設定処理
    Call subFormLoad
    
    pIsChanged = False  '2006/09/06
    
    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_顧客コード_AfterUpdate", True)
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_BeforeUpdate イベント
'        MODULE_ID        :cmb_顧客コード_BeforeUpdate
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_BeforeUpdate(Cancel As Integer)

    Dim varCode  As Variant

  On Error GoTo Exception
    
    varCode = Nz(Me.cmb_顧客コード.VALUE, "")
    If varCode = "" Then
        MsgBox "顧客コードは必須入力です。", vbInformation, P_MSG_ERR
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If
    
    '入力チェック
    Select Case fncCheckNum(varCode, 6)
        Case 2
            ' 桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_MSG_ERR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub

        Case 9
            ' その他のエラー
            MsgBox P_MSG_ERR_INPUT_VALUE, vbInformation, P_MSG_ERR
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        Case Else
    End Select

    ' ファイルへの存在チェック
    If fncCheck_User_Mast(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If

    Exit Sub

Exception:
    Call MSZZ024_M00("cmb_顧客コード_BeforeUpdate", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_GotFocus イベント
'        MODULE_ID        :cmb_顧客コード_GotFocus
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_GotFocus()

    Dim strSQL As String

  On Error GoTo cmb_顧客コード_GotFocus_Err

    If Nz(Me.txt_顧客名, "") = "" Then
        Me.cmb_顧客コード.RowSource = ""
        Exit Sub

    Else
        ' 顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ表示のSQL文作成
        strSQL = "SELECT TOP 100 Format([USER_CODE],'000000') AS 式1, "
        strSQL = strSQL & " dbo_USER_MAST.USER_NAME "
        strSQL = strSQL & "  FROM dbo_USER_MAST "
        strSQL = strSQL & " WHERE (((dbo_USER_MAST.USER_KANA) LIKE '*" & Me.txt_顧客名 & "*')) "
        strSQL = strSQL & " ORDER BY Format([USER_CODE],'000000');"

        Me.cmb_顧客コード.RowSource = strSQL
        Me.cmb_顧客コード.Requery

        If Me.cmb_顧客コード.ListCount = 0 Then Me.cmb_顧客コード = Null
    End If

    Exit Sub

cmb_顧客コード_GotFocus_Err:
    Call MSZZ024_M00("cmb_顧客コード_GotFocus", True)
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客コード_KeyPress イベント
'        MODULE_ID        :cmb_顧客コード_KeyPress
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_顧客コード, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_新規入力_Click イベント
'        MODULE_ID        :cmd_新規入力_Click
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_新規入力_Click()

  On Error GoTo Exception

    Call subOpen_FVS910(P_MSG_MODE_INS)

    Exit Sub

Exception:
    Call MSZZ024_M00("cmd_新規入力_Click", True)
End Sub

' ☆☆☆ Private関数群 ☆☆☆
'==============================================================================*
'
'        MODULE_NAME      :画面項目初期設定
'        MODULE_ID        :initFormItems
'        PARAMETER        :strOpenForm(IN) オープンフォーム
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub initFormItems(ByVal strOpenForm As String)

    On Error GoTo Exception

    ' 初期処理
    With Me
        Select Case strOpenForm
            Case P_FORM_ID_FVS901
                ' *** ヤード修繕 ***
                ' タイトル設定
                .lbl_タイトル.Caption = P_FORM_TITTLE_FVS901

                ' 表示項目
                .lbl_ヤード名.Visible = True
                .cmb_ヤードコード.Visible = True
                .txt_ヤード名.Visible = True
                .lbl_所在地名.Visible = True
                .txt_ヤード住所.Visible = True
                .lbl_営業開始日.Visible = True
                .txt_営業開始日.Visible = True
                .lbl_営業終了日.Visible = True
                .txt_営業終了日.Visible = True
                .lbl_ヤード解約予定日.Visible = True
                .txt_ヤード解約予定日.Visible = True

                ' 非表示項目
                .lbl_顧客コード.Visible = False
                .cmb_顧客コード.Visible = False
                .lbl_顧客名.Visible = False
                .txt_顧客名.Visible = False
                .lbl_コンテナ番号.Visible = False
                .cmb_コンテナ番号.Visible = False
                .lbl_段区分.Visible = False
                .txt_段区分名.Visible = False
                .lbl_種別.Visible = False
                .txt_サイズ.Visible = False
                
                '2013/10/02 M.HONDA INS
                .txt_KEY_TYPE.Visible = False
                .txt_KEY_NO.Visible = False
                .txt_NUM.Visible = False
                .lbl_KEY_TYPE.Visible = False
                .lbl_KEY_NO__.Visible = False
                .lbl_NUM.Visible = False
                '2013/10/02 M.HONDA INS

                ' ヤードコードをパラメータで受けた場合
                If pIsAutoRead = True Then
                    .cmb_ヤードコード.VALUE = pstrYardCode
                    .cmb_ヤードコード.Enabled = False
                Else
                ' パラメータ無しモード
                    .cmb_ヤードコード.VALUE = ""
                    .cmb_ヤードコード.Enabled = True
                End If
                
            Case P_FORM_ID_FVS902
                ' *** 顧客修繕 ***
                ' タイトル設定
                .Caption = Replace(Replace(Me.Caption, P_FORM_TITTLE_FVS901, P_FORM_TITTLE_FVS902), P_FORM_ID_FVS901, P_FORM_ID_FVS902)
                .lbl_タイトル.Caption = P_FORM_TITTLE_FVS902

                ' 表示項目
                .lbl_顧客コード.Visible = True
                .cmb_顧客コード.Visible = True
                .lbl_顧客名.Visible = True
                .txt_顧客名.Visible = True

                ' 非表示項目
                .lbl_ヤード名.Visible = False
                .cmb_ヤードコード.Visible = False
                .txt_ヤード名.Visible = False
                .lbl_所在地名.Visible = False
                .txt_ヤード住所.Visible = False
                .lbl_営業開始日.Visible = False
                .txt_営業開始日.Visible = False
                .lbl_営業終了日.Visible = False
                .txt_営業終了日.Visible = False
                .lbl_ヤード解約予定日.Visible = False
                .txt_ヤード解約予定日.Visible = False
                .lbl_コンテナ番号.Visible = False
                .cmb_コンテナ番号.Visible = False
                .lbl_段区分.Visible = False
                .txt_段区分名.Visible = False
                .lbl_種別.Visible = False
                .txt_サイズ.Visible = False
                '2013/10/02 M.HONDA INS
                .txt_KEY_TYPE.Visible = False
                .txt_KEY_NO.Visible = False
                .txt_NUM.Visible = False
                .lbl_KEY_TYPE.Visible = False
                .lbl_KEY_NO__.Visible = False
                .lbl_NUM.Visible = False
                '2013/10/02 M.HONDA INS

                ' 顧客コードをパラメータで受けた場合
                If pIsAutoRead = True Then
                    .txt_顧客名.Enabled = False
                    .cmb_顧客コード.VALUE = pstrUserCode
                    .cmb_顧客コード.Enabled = False
                Else
                ' パラメータ無しモード
                    .cmb_顧客コード.VALUE = ""
                    .cmb_顧客コード.Enabled = True
                    .txt_顧客名.Enabled = True
                    .txt_顧客名.SetFocus
                End If


            Case P_FORM_ID_FVS903
                ' *** コンテナ(部屋)修繕 ***
                ' タイトル設定
                
                .Caption = Replace(Replace(Me.Caption, P_FORM_TITTLE_FVS901, P_FORM_TITTLE_FVS903), P_FORM_ID_FVS901, P_FORM_ID_FVS903)
                .lbl_タイトル.Caption = P_FORM_TITTLE_FVS903

                ' 表示項目
                .lbl_ヤード名.Visible = True
                .cmb_ヤードコード.Visible = True
                .txt_ヤード名.Visible = True
                .lbl_所在地名.Visible = True
                .txt_ヤード住所.Visible = True
                .lbl_コンテナ番号.Visible = True
                .cmb_コンテナ番号.Visible = True
                .lbl_段区分.Visible = True
                .txt_段区分名.Visible = True
                .lbl_種別.Visible = True
                .txt_サイズ.Visible = True
                '2013/10/02 M.HONDA INS
                .txt_KEY_TYPE.Visible = True
                .txt_KEY_NO.Visible = True
                .txt_NUM.Visible = True
                .lbl_KEY_TYPE.Visible = True
                .lbl_KEY_NO__.Visible = True
                .lbl_NUM.Visible = True
                '2013/10/02 M.HONDA INS

                '' ▼ 20101224 M.HONDA
                Form_FVS901_S01.lbl_完了日1.Visible = True
                Form_FVS901_S01.lbl_内容1.Visible = True
                Form_FVS901_S01.txt_完了日1.Visible = True
                Form_FVS901_S01.txt_内容1.Visible = True
                '' ▲ 20101224 M.HONDA

                ' 非表示項目
                .lbl_営業開始日.Visible = False
                .txt_営業開始日.Visible = False
                .lbl_営業終了日.Visible = False
                .txt_営業終了日.Visible = False
                .lbl_ヤード解約予定日.Visible = False
                .txt_ヤード解約予定日.Visible = False
                .lbl_顧客コード.Visible = False
                .cmb_顧客コード.Visible = False
                .lbl_顧客名.Visible = False
                .txt_顧客名.Visible = False
                
                '' ▼ 20101224 M.HONDA
                Form_FVS901_S01.lbl_内容.Visible = False
                Form_FVS901_S01.txt_内容.Visible = False
                '' ▲ 20101224 M.HONDA
                
                ' ヤードコードと部屋番号をパラメータで受けた場合
                If pIsAutoRead = True Then
                    .cmb_ヤードコード.VALUE = pstrYardCode
                    .cmb_ヤードコード.Enabled = False
                    .cmb_コンテナ番号.VALUE = pstrRoomNo
                    .cmb_コンテナ番号.Enabled = False
                Else
                ' パラメータ無しモード
                    .cmb_ヤードコード.VALUE = ""
                    .cmb_ヤードコード.Enabled = True
                    .cmb_コンテナ番号.VALUE = ""
                    .cmb_コンテナ番号.Enabled = True
                End If
        End Select
    End With

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "initFormItems" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :コンテナ番号リクエリー処理
'        MODULE_ID        :subOPEN_FVS910
'        PARAMETER        :strKbn(IN) 処理区分(1：新規／2：修正)
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subOpen_FVS910(ByVal strKbn As String, Optional ByVal strUNIQUE_CD As String = "")

    Dim strParam  As String
    Dim mode      As String

    On Error GoTo Exception

    ' パラメータ作成
    ' 処理区分(01：ヤード修繕／02：顧客修繕／03：コンテナ(部屋)修繕)
    ' モード(0:参照／1：新規／2：修正)
    ' ユニークコード
    ' ヤードコード
    ' コンテナ番号
    ' 顧客コード
    ' 契約番号
    ' 画面ID
    ' 権限可否  Add 2006/09/06
    strParam = ""
    
    ' 権限に関係なく新規登録か修正モードとなる Add 2006/09/06
    mode = strKbn
    
    With Me
        Select Case pstrOpenForm
            Case P_FORM_ID_FVS901
                ' *** ヤード修繕 ***
                If Nz(.cmb_ヤードコード, "") = "" Then
                    MsgBox P_MSG_ERR_INPUT_YARD, vbInformation, P_MSG_ERR: Exit Sub
                End If

                strParam = CmReprMod.P_対象_ヤード & "," & mode & "," & strUNIQUE_CD & "," & Format(.cmb_ヤードコード, "000000") & ",,," & pstrAcptNo

            Case P_FORM_ID_FVS902
                ' *** 顧客修繕 ***
                If Nz(.cmb_顧客コード, "") = "" Then
                    MsgBox P_MSG_ERR_INPUT_USER, vbInformation, P_MSG_ERR: Exit Sub
                End If

                strParam = CmReprMod.P_対象_顧客 & "," & mode & "," & strUNIQUE_CD & ",,," & Format(.cmb_顧客コード, "000000") & "," & pstrAcptNo

            Case P_FORM_ID_FVS903
                ' *** コンテナ(部屋)修繕 ***
                If Nz(.cmb_ヤードコード, "") = "" Then
                    MsgBox P_MSG_ERR_INPUT_YARD, vbInformation, P_MSG_ERR: Exit Sub
                End If

                If Nz(.cmb_コンテナ番号, "") = "" Then
                    MsgBox P_MSG_ERR_INPUT_CNTA, vbInformation, P_MSG_ERR: Exit Sub
                End If

                strParam = CmReprMod.P_対象_部屋 & "," & mode & "," & strUNIQUE_CD & "," & Format(.cmb_ヤードコード, "000000") & "," & Format(.cmb_コンテナ番号, "000000") & ",," & pstrAcptNo
        End Select
    End With

    ' 画面ID 2006/09/06 権限を追加
    strParam = strParam & "," & pstrOpenForm & "," & pIsAuthority

    ' 修繕入力ダイアログ表示
    doCmd.OpenForm P_FORM_ID_FVS910, , , , , , strParam

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "subOpen_FVS910" & vbRightAllow & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :コンテナ番号リクエリー処理
'        MODULE_ID        :subCNTA_Requery
'        PARAMETER        :strOpenForm(IN) オープンフォーム
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCNTA_Requery()

    Dim strSQL  As String

    strSQL = "  SELECT   Format(CNTA_NO,'000000') AS 式2 "
    strSQL = strSQL & " ,IIf(CNTA_STEP = 0, '上段', '下段') As 式1 "
    strSQL = strSQL & "  FROM dbo_CNTA_MAST "
    strSQL = strSQL & " WHERE CNTA_CODE = " & Val(Me.cmb_ヤードコード)
    strSQL = strSQL & " ORDER BY Format(CNTA_NO, '000000');"

    Me.cmb_コンテナ番号.RowSource = strSQL
    Me.cmb_コンテナ番号.Requery
End Sub
    
'==============================================================================*
'
'        MODULE_NAME      :ｺﾝﾎﾞﾎﾞｯｸｽの内容が空の場合、名称テキストも空にする
'        MODULE_ID        :fncControlCrear
'        PARAMETER        :intKbn(IN)   = 処理区分
'        RETURN           :True = 入力項目がNULL以外／False = 入力項目がNULL
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncControlCrear(ByVal intKbn As Integer) As Boolean

    With Me
        Select Case intKbn
            Case 1
                If Nz(.cmb_ヤードコード, "") = "" Then
                    .txt_ヤード名 = ""
                    .txt_ヤード住所 = ""
                    .txt_営業開始日 = ""
                    .txt_営業終了日 = ""
                    GoTo fncControlCrear_Exit
                End If

            Case 2
                If Nz(.cmb_ヤードコード, "") = "" Then
                    .txt_ヤード名 = ""
                    .txt_ヤード住所 = ""
                    .txt_営業開始日 = ""
                    .txt_営業終了日 = ""
                    .cmb_コンテナ番号 = ""
                    .txt_段区分名 = ""
                    .txt_サイズ = ""
                     GoTo fncControlCrear_Exit
                End If

            Case 3
                If Nz(.cmb_コンテナ番号, "") = "" Then
                    .txt_段区分名 = ""
                    .txt_サイズ = ""
                    GoTo fncControlCrear_Exit
                End If
        End Select
    End With

    fncControlCrear = True

    Exit Function

fncControlCrear_Exit:
End Function

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
            MsgBox P_MSG_ERR_INPUT_KETA, vbInformation, P_MSG_ERR
            Exit Function

        Case 9
            ' その他のエラー
            MsgBox P_MSG_ERR_INPUT_NAIYO, vbInformation, P_MSG_ERR
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
'        MODULE_NAME      :サブフォームの表示
'        MODULE_ID        :subFormLoad
'        PARAMETER        :
'        CREATE_DATE      :2006/06/24
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFormLoad()

    Dim isException As Boolean
    Dim isThisDBOpen  As Boolean  ' ここでDBオープンしたかのフラグ

  On Error GoTo Exception
    isException = False
    
    If pobjKASE_DB Is Nothing Then
         Call subConnectDB(P_DB_KASE)
         isThisDBOpen = True
    End If

    Call subDeleteWork            ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙﾃﾞｰﾀ削除
    Call subCreateWork            ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙﾃﾞｰﾀ作成
    Me.FVS901_FVS901S01.Requery
    
    GoTo Finally

Exception:
    isException = True
    
Finally:
    If isThisDBOpen = True And Not pobjKASE_DB Is Nothing Then Call subDisConnectDB(P_DB_KASE)
    If isException = True Then
        Call Err.Raise(Err.Number, "subFormLoad" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub
' ☆☆☆ Private関数群 ☆☆☆

' ☆☆☆ DB ACCESS系関数群 ☆☆☆
'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :subConnectDB
'        PARAMETER        :intKbn(IN)     DB接続区分(0 = コンテナDB／1 = 加瀬DB)
'                          strBumoCD(IN)  部門コード
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectDB(ByVal intKbn As Integer, Optional ByVal strBumoCD As String = "")

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo subConnectDB_Err

    Select Case intKbn
        Case P_DB_KONT
            ' -----------------------------------------------------------------------------------------------------------
            ' コンテナDB接続
            strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME_" & strBumoCD)
            If strDataSource = "" Then
                ' テーブル[SETU_TABL]の設定不正
                Exit Sub
            End If

            ' DB接続
            Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(CStr(strBumoCD)))
            ' -----------------------------------------------------------------------------------------------------------

        Case P_DB_KASE
            ' -----------------------------------------------------------------------------------------------------------
            ' 加瀬DB接続
            strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME")
        
            If strDataSource = "" Then
                ' テーブル[SETU_TABL]の設定不正
                Exit Sub
            End If

            Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
            ' -----------------------------------------------------------------------------------------------------------
    End Select

    Exit Sub

subConnectDB_Err:
    Call Err.Raise(Err.Number, "subConnectDB" & vbRightAllow & Err.Source, _
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
'        MODULE_NAME      :DB切断
'        MODULE_ID        :subDisConnectDB
'        PARAMETER        :intKbn(IN)
'                            1 = コンテナDB
'                            2 = 加瀬DB
'        CREATE_DATE      :2005/08/10
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDisConnectDB(ByVal intKbn As Integer)

    On Error GoTo subDisConnectDB_Err

    Select Case intKbn
        Case P_DB_KONT
            ' コンテナDB
            pobjKONT_DB.Close
            Set pobjKONT_DB = Nothing

        Case P_DB_KASE
            ' 加瀬DB
            pobjKASE_DB.Close
            Set pobjKASE_DB = Nothing
    End Select

    Exit Sub

subDisConnectDB_Err:
    Call Err.Raise(Err.Number, "subDisConnectDB" & vbRightAllow & Err.Source, _
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
'        MODULE_NAME      :ヤードマスタ問い合わせ
'        MODULE_ID        :fncCheck_Yard_Mast
'        PARAMETER        :intMode(IN)  = 処理モード(0：存在チェック／1：ヤードマスタデータ取得)
'        RETURN           :True = 正常終了／False = 異常終了
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheck_Yard_Mast(ByVal intMode As Integer) As Boolean

    Dim varCode    As Variant
    Dim varRtn     As Variant
    Dim strSQL     As String
    Dim objRs      As Recordset
    Dim blnErrFlg  As Boolean

    On Error GoTo fncCheck_Yard_Mast_Err

    fncCheck_Yard_Mast = False
    blnErrFlg = False

    ' ヤードコード取得
    varCode = Nz(Me.cmb_ヤードコード, "")
    If varCode = "" Then Exit Function


    Select Case intMode
        Case 0
            ' *** 存在チェック ***
            ' 存在チェック
            strSQL = " YARD_CODE = " & Val(varCode)
            varRtn = Nz(DCount("YARD_CODE", "dbo_YARD_MAST", strSQL), "")

            If varRtn < 1 Then
                MsgBox P_MSG_ERR_NODATA_YARD, vbInformation, P_MSG_KAKUNIN
                fncCheck_Yard_Mast = False: Exit Function
            Else
                fncCheck_Yard_Mast = True 'そんざいスマス
            End If

        Case 1
            ' コンテナDB接続
            Call subConnectDB(P_DB_KONT, pstrBumonCd)

            ' *** ヤードマスタデータ取得 ***
            strSQL = "  SELECT   YARD_NAME "       ' ヤード名
            strSQL = strSQL & " ,YARD_ADDR_1 "     ' ヤード住所１
            strSQL = strSQL & " ,YARD_ADDR_2 "     ' ヤード住所２
            strSQL = strSQL & " ,YARD_ADDR_3 "     ' ヤード住所３
            strSQL = strSQL & " ,YARD_BEGIN_DAY "  ' 営業開始日
            strSQL = strSQL & " ,YARD_END_DAY "    ' ヤード解約予定日
            strSQL = strSQL & " ,YARD_RENTEND_DAY " ' 営業終了日 2006/09/06
            strSQL = strSQL & " ,(SELECT TOP 1 CONT_ENDCAL_MM FROM CONT_MAST) AS ENDCAL_MM " '営業終了計算月
            strSQL = strSQL & "  FROM YARD_MAST "
            strSQL = strSQL & " WHERE YARD_CODE = " & Val(varCode)
            Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

            If Not (objRs.BOF And objRs.EOF) Then
                Me.txt_ヤード名 = objRs.Fields("YARD_NAME")
                Me.txt_ヤード住所 = objRs.Fields("YARD_ADDR_1") & _
                                    objRs.Fields("YARD_ADDR_2") & _
                                    objRs.Fields("YARD_ADDR_3")
                Me.txt_営業開始日 = objRs.Fields("YARD_BEGIN_DAY")
                Me.txt_ヤード解約予定日 = objRs.Fields("YARD_END_DAY")
                Me.txt_営業終了日 = objRs.Fields("YARD_RENTEND_DAY")  '2006/09/06 営業終了日はフィールドとしてもつ
                fncCheck_Yard_Mast = True 'そんざいスマス
            End If
    End Select

    GoTo fncCheck_Yard_Mast_Exit

fncCheck_Yard_Mast_Err:
    ' エラーフラグセット
    blnErrFlg = True

fncCheck_Yard_Mast_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKONT_DB Is Nothing Then Call subDisConnectDB(P_DB_KONT)

    If blnErrFlg = True Then
        ' エラー発生時
        Call Err.Raise(Err.Number, "fncCheck_Yard_Mast" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ問い合わせ
'        MODULE_ID        :fncCheck_Cnta_Mast
'        PARAMETER        :intMode(IN)  = 処理モード(0：存在チェック／1：ヤードマスタデータ取得)
'        RETURN           :True = 正常終了／False = 異常終了
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheck_Cnta_Mast(ByVal intMode As Integer) As Boolean

    Dim varCode    As Variant
    Dim varRtn     As Variant
    Dim strSQL     As String
    Dim objRs      As Recordset
    Dim blnErrFlg  As Boolean

    On Error GoTo fncCheck_Cnta_Mast_Err
    blnErrFlg = False
    fncCheck_Cnta_Mast = False

    ' コンテナ番号取得
    varCode = Nz(Me.cmb_コンテナ番号, "")
    If varCode = "" Then Exit Function
    
    
    Select Case intMode
        Case 0
            ' *** 存在チェック ***
            strSQL = " CNTA_CODE = " & Val(Me.cmb_ヤードコード)
            strSQL = strSQL & " AND CNTA_NO = " & Val(varCode)
            varRtn = Nz(DLookup("CNTA_STEP", "dbo_CNTA_MAST", strSQL), "")

            If varRtn = "" Then
                MsgBox P_MSG_ERR_NODATA_CNTA, vbInformation, P_MSG_KAKUNIN
                GoTo fncCheck_Cnta_Mast_Exit
            Else
                fncCheck_Cnta_Mast = True '存在する
            End If

        Case 1
            ' コンテナDB接続
            Call subConnectDB(P_DB_KONT, pstrBumonCd)

            ' *** コンテナマスタデータ取得 ***
            strSQL = " SELECT    CNTA_STEP "                                    ' 段区分
            strSQL = strSQL & " ,STEP_NAME.NAME_NAME AS STEP_NAME"              ' 段名
            strSQL = strSQL & " ,CNTA_SIZE "                                    ' サイズ
            
            strSQL = strSQL & " ,C060.NAME_NAME "
            'strSQL = strSQL & " ,CNTA_KEY_NO "       '2014/09/29 M.HONDA DEL
            strSQL = strSQL & " ,CNTA_KEY_NUM "
            '2014/09/29 M.HONDA INS
            strSQL = strSQL & " ,isnull(( "
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "         AKKNT.AKKNT_KAINO "
            strSQL = strSQL & "     FROM "
            strSQL = strSQL & "         ( "
            strSQL = strSQL & "         SELECT "
            strSQL = strSQL & "             AKKNT_YARD, "
            strSQL = strSQL & "             AKKNT_CTNO, "
            strSQL = strSQL & "             AKKNT_KTYPE, "
            strSQL = strSQL & "             AKKNT_HATUD, "
            strSQL = strSQL & "             ROW_NUMBER() OVER (PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE ORDER BY AKKNT_HATUD DESC) AS SEQ, "
            strSQL = strSQL & "             AKKNT_KAINO "
            strSQL = strSQL & "         FROM "
            strSQL = strSQL & "             AKKN_TRAN "
            strSQL = strSQL & "         WHERE "
            strSQL = strSQL & "             AKKNT_YARD = CNTA_MAST.CNTA_CODE AND "
            strSQL = strSQL & "             AKKNT_CTNO = CNTA_MAST.CNTA_NO AND "
            strSQL = strSQL & "             AKKNT_KTYPE = CNTA_MAST.CNTA_KEY_TYPE "
            strSQL = strSQL & "         ) AKKNT "
            strSQL = strSQL & "     WHERE "
            strSQL = strSQL & "         AKKNT.SEQ = 1 "
            strSQL = strSQL & "     ) ,CNTA_KEY_NO) CNTA_KEY_NO "
            '2014/09/29 M.HONDA INS
            strSQL = strSQL & " FROM CNTA_MAST "                                ' コンテナマスタ
            strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS STEP_NAME "       ' 段区分を外部結合
            strSQL = strSQL & "  ON STEP_NAME.NAME_CODE = CNTA_MAST.CNTA_STEP " ' 段区分を外部結合
            strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST AS C060 "
            'strSQL = strSQL & "  ON C060.NAME_CODE = CNTA_MAST.CNTA_KEY_TYPE "  ' 2013/11/12 M.HONDA
            strSQL = strSQL & "  ON C060.NAME_CODE = CONVERT(int, CNTA_MAST.CNTA_KEY_TYPE) "
            strSQL = strSQL & "  AND C060.NAME_ID = '060'"
            strSQL = strSQL & " WHERE CNTA_CODE = " & Val(Me.cmb_ヤードコード)
            strSQL = strSQL & "   AND CNTA_NO   = " & Val(varCode)
            strSQL = strSQL & "   AND STEP_NAME.NAME_ID = '090'"
            Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

            If Not (objRs.BOF And objRs.EOF) Then
                ' 段区分
                Me.txt_段区分名 = objRs.Fields("STEP_NAME")
                ' サイズ
                If Nz(objRs.Fields("CNTA_SIZE"), "") <> "" Then
                    Me.txt_サイズ = Format$(objRs.Fields("CNTA_SIZE"), "0.00")
                End If
                
                ' 鍵区分
                Me.txt_KEY_TYPE = Nz(objRs.Fields("NAME_NAME"), "")
                ' 鍵番号
                Me.txt_KEY_NO = Nz(objRs.Fields("CNTA_KEY_NO"), "")
                ' 本数
                Me.txt_NUM = Nz(objRs.Fields("CNTA_KEY_NUM"), "")
                                
                fncCheck_Cnta_Mast = True '存在する
            End If
            
    End Select
    
    GoTo fncCheck_Cnta_Mast_Exit

fncCheck_Cnta_Mast_Err:
    blnErrFlg = True

fncCheck_Cnta_Mast_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKONT_DB Is Nothing Then Call subDisConnectDB(P_DB_KONT)

    If blnErrFlg = True Then
        ' エラー発生時
        Call Err.Raise(Err.Number, "fncCheck_Cnta_Mast" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :顧客マスタ問い合わせ
'        MODULE_ID        :fncCheck_User_Mast
'        PARAMETER        :intMode(IN)  = 処理モード(0：存在チェック／1：ヤードマスタデータ取得)
'        RETURN           :True = 正常終了／False = 異常終了
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheck_User_Mast(ByVal intMode As Integer) As Boolean

    Dim varCode    As Variant
    Dim varRtn     As Variant
    Dim strSQL     As String
    Dim objRs      As Recordset
    Dim blnErrFlg  As Boolean

  On Error GoTo fncCheck_User_Mast_Err

    fncCheck_User_Mast = False
    blnErrFlg = False

    varCode = Nz(Me.cmb_顧客コード, "")
    If varCode = "" Then Exit Function

    Select Case intMode
        Case 0
            ' *** 存在チェック ***
            strSQL = ""
            strSQL = " USER_CODE =" & Val(varCode)
            varRtn = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strSQL), "")

            If varRtn = "" Then
                MsgBox P_MSG_ERR_NODATA_USER, vbInformation, P_MSG_KAKUNIN
                GoTo fncCheck_User_Mast_Exit
            Else
                fncCheck_User_Mast = True
            End If

        Case 1
            ' コンテナDB接続
            Call subConnectDB(P_DB_KONT, pstrBumonCd)

            ' *** ヤードマスタデータ取得 ***
            strSQL = "SELECT USER_KANA "
            strSQL = strSQL & " FROM USER_MAST "
            strSQL = strSQL & " WHERE USER_CODE = '" & Val(varCode) & "' "
            Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

            With Me
                If Not (objRs.BOF And objRs.EOF) Then
                    ' 顧客名カナ
                    .txt_顧客名 = objRs.Fields("USER_KANA")
                    fncCheck_User_Mast = True
                Else
                    ' 顧客名カナ
                    .txt_顧客名 = ""
                End If
            End With
            objRs.Close: Set objRs = Nothing
    End Select

    GoTo fncCheck_User_Mast_Exit

fncCheck_User_Mast_Err:
    blnErrFlg = True

fncCheck_User_Mast_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKONT_DB Is Nothing Then Call subDisConnectDB(P_DB_KONT)

    If blnErrFlg = True Then
        ' エラー発生時
        Call Err.Raise(Err.Number, "fncCheck_User_Mast" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :名称マスタのデータ取得
'        MODULE_ID        :fncGetName
'        PARAMETER        :intKbn(IN)    = (0：名称／1：略称)
'                          strNameID(IN) = マスタID
'                          strCode(IN)   = コード
'                          objDB(IN)     = データベースオブジェクト
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'        備考             ：コードテーブルは加瀬DBを参照
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetName(ByVal intKbn As Integer, _
                            ByVal strNameID As String, _
                            ByVal strCode As String, _
                            ByVal objDb As Database) As String

    Dim strName    As String
    Dim strSQL     As String
    Dim objRs      As Recordset
    Dim blnErrFlg  As Boolean

    On Error GoTo fncGetName_Err

    fncGetName = "": strName = ""
    blnErrFlg = False

    If strCode = "" Then Exit Function

    ' 名称マスタより名称を取得
    If intKbn = 0 Then
        strSQL = "SELECT CODET_NAMEN AS NAME "  ' 名称
    ElseIf intKbn = 1 Then
        strSQL = "SELECT CODET_RYAKN AS NAME "  ' 略称
    End If

    strSQL = strSQL & "  FROM CODE_TABL "
    strSQL = strSQL & " WHERE CODET_SIKBC = '" & strNameID & "'"
    strSQL = strSQL & "   AND CODET_CODEC = " & strCode
    Set objRs = objDb.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

    With objRs
        If Not .EOF Then
            strName = Nz(.Fields("NAME"), "")
        End If
    End With
    objRs.Close
    Set objRs = Nothing

    fncGetName = strName

    GoTo fncCheck_Cnta_Mast_Exit

fncGetName_Err:
    blnErrFlg = True

fncCheck_Cnta_Mast_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing

    If blnErrFlg = True Then
        ' エラー発生時
        Call Err.Raise(Err.Number, "fncGetName" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ削除
'        MODULE_ID        :fncDeleteWorkTable
'        PARAMETER        :
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDeleteWorkTable() As Boolean

    Dim objDb      As Database
    Dim strSQL     As String
    Dim blnErrFlg  As Boolean

    On Error GoTo fncDeleteWorkTable_Err

    fncDeleteWorkTable = False
    blnErrFlg = False

    Set objDb = CurrentDb

    ' テーブルの存在チェック
    If fncChkTableExist(P_WK_TABLE_NM) = True Then
        ' ｻﾌﾞﾌｫｰﾑのｿｰｽｵﾌﾞｼﾞｪｸﾄ無効化
        Me.FVS901_FVS901S01.SourceObject = ""

        ' テーブル削除
        objDb.TableDefs.Delete P_WK_TABLE_NM
        Select Case Err.Number
            Case 0
            Case 3265
                Err.Clear
            Case Else
                MsgBox "fncDeleteWorkTable:" & Err.Number & vbCrLf & Err.Description, , "Error!!"
                Err.Clear
        End Select
    End If

    fncDeleteWorkTable = True
    GoTo fncDeleteWorkTable_Exit

fncDeleteWorkTable_Err:
    blnErrFlg = True

fncDeleteWorkTable_Exit:
    objDb.Close
    Set objDb = Nothing

    If blnErrFlg = True Then
        Call Err.Raise(Err.Number, "fncDeleteWorkTable" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :テーブルの存在チェック
'        MODULE_ID        :fncCHKTableExist
'        PARAMETER        :
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTableExist(strTableName As String) As Boolean

    Dim objDb      As Database
    Dim tdf        As TableDef
    Dim blnErrFlg  As Boolean

    On Error GoTo fncCHKTableExist_Err

    fncChkTableExist = False
    blnErrFlg = False

    Set objDb = CurrentDb

    For Each tdf In objDb.TableDefs
        If tdf.NAME = strTableName Then
            fncChkTableExist = True
            Exit For
        End If
    Next tdf

    GoTo fncCHKTableExist_Exit

fncCHKTableExist_Err:
    blnErrFlg = True

fncCHKTableExist_Exit:
    objDb.Close
    Set objDb = Nothing

    If blnErrFlg = True Then
        Call Err.Raise(Err.Number, "fncChkTableExist" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ作成
'        MODULE_ID        :fncCreateWorkTable
'        PARAMETER        :
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable() As Boolean

    Dim objDb      As Database
    Dim tdfNew     As TableDef
    Dim intCount   As Integer
    Dim blnErrFlg  As Boolean

    On Error GoTo fncCreateWorkTable_Err

    fncCreateWorkTable = False
    blnErrFlg = False

    Set objDb = CurrentDb

    Set tdfNew = objDb.CreateTableDef(P_WK_TABLE_NM)
    With tdfNew
        .Fields.Append .CreateField("FVS901_UNIQEC", dbText, 100)      ' ユニークコード
        .Fields.Append .CreateField("FVS901_GENTD", dbText, 10)        ' 発生日(YYYY/MM/DD)
        .Fields.Append .CreateField("FVS901_TYPE", dbText, 20)         ' 区分名
        .Fields.Append .CreateField("FVS901_NAIYO", dbText, 250)       ' 内容
        .Fields.Append .CreateField("FVS901_TANTNM", dbText, 36)       ' 入力担当名

        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With

    ' テーブルの追加
    objDb.TableDefs.Append tdfNew

    ' ｻﾌﾞﾌｫｰﾑのｿｰｽｵﾌﾞｼﾞｪｸﾄ設定
    Me.FVS901_FVS901S01.SourceObject = P_SUBFRM_NM
    fncCreateWorkTable = True
    GoTo fncCreateWorkTable_Exit

fncCreateWorkTable_Err:
    blnErrFlg = True

fncCreateWorkTable_Exit:
    objDb.Close
    Set objDb = Nothing

    If blnErrFlg = True Then
        Call Err.Raise(Err.Number, "fncCreateWorkTable" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :初期ＤＢ読込
'        MODULE_ID        :initDBLoad
'        PARAMETER        :
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'        NOTE             :戻り値としてエラーメッセージを返す。""なら正常完了
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function initDBLoad() As String

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
    Call subConnectDB(P_DB_KASE)

    ' 更新権限取得を行う
    pIsAuthority = CmReprMod.isAuthority(pobjKASE_DB, pstrUserName)
    ' 権限による項目設定                  Del 2006/09/06 tajima 新規入力ダイアログは誰でも呼べる
    ' cmd_新規入力.Enabled = pIsAuthority Del 2006/09/06 tajima
    
    ' メニューから呼ばれたとき
    If pIsAutoRead = False Then
        Call subDeleteWork
        Me.FVS901_FVS901S01.Requery
    Else
    ' パラメータ設定可否による自動処理の実行
        Select Case pstrOpenForm
            Case P_FORM_ID_FVS901
                ' ヤード情報の取得と有無チェック
                If fncCheck_Yard_Mast(1) = False Then
                    errorMessage = "ヤードコード【" & pstrYardCode & "】は存在しません" & _
                                      vbCr & "※ヤードマスタに保存されているか確認して下さい。"
                End If
            Case P_FORM_ID_FVS902
                ' 顧客情報の取得と有無チェック
                If fncCheck_User_Mast(1) = False Then
                    errorMessage = "顧客コード【" & pstrUserCode & "】は存在しません" & _
                                      vbCr & "※顧客マスタに保存されているか確認して下さい。"
                End If
            Case P_FORM_ID_FVS903
                ' ヤードとコンテナ情報の取得と有無チェック
                If fncCheck_Yard_Mast(1) = False Then
                    errorMessage = "ヤードコード【" & pstrYardCode & "】は存在しません" & _
                                      vbCr & "※ヤードマスタに保存されているか確認して下さい。"
                ElseIf fncCheck_Cnta_Mast(1) = False Then
                    errorMessage = "ヤード【" & pstrYardCode & "】に部屋番号【" & pstrRoomNo & "】は存在しません" & _
                                      vbCr & "※コンテナマスタに保存されているか確認して下さい。"
                End If
        End Select
        
        If errorMessage = "" Then
          ' エラーがなければサブフォーム生成
            Call subDeleteWork
            Call subCreateWork
            Me.FVS901_FVS901S01.Requery
            Me.cmd_新規入力.SetFocus  'Rep 2006/09/06 tajima
        End If
    End If

    GoTo Finally
    
Exception:
    isException = True
    errorMessage = "system error"

Finally:
    If Not pobjKASE_DB Is Nothing Then Call subDisConnectDB(P_DB_KASE)
    initDBLoad = errorMessage
    If isException = True Then
        Call Err.Raise(Err.Number, "initDBLoad" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ削除
'        MODULE_ID        :subDeleteWork
'        PARAMETER        :
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDeleteWork()

    Dim strSQL     As String
    Dim objDb      As Database
    Dim blnErrFlg  As Boolean

    On Error GoTo subDeleteWork_Err
    blnErrFlg = False
    
    Set objDb = CurrentDb

    strSQL = "DELETE FROM " & P_WK_TABLE_NM
    objDb.Execute strSQL, dbFailOnError

    GoTo subDeleteWork_Exit

subDeleteWork_Err:
    blnErrFlg = True

subDeleteWork_Exit:
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing

    If blnErrFlg = True Then
        Call Err.Raise(Err.Number, "subDeleteWork" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ作成
'        MODULE_ID        :fncCreateWorkTable
'        PARAMETER        :
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreateWork()

    Dim objDb      As Database
    Dim strSQL     As String
    Dim objRs      As Recordset
    Dim objRsWk    As Recordset
    Dim varName    As Variant
    Dim strUniqec  As String
    Dim blnErrFlg  As Boolean

    On Error GoTo subCreateWork_Err
    blnErrFlg = False
    
    ' ワークテーブルの為のローカルに接続
    Set objDb = CurrentDb

    Select Case pstrOpenForm
        Case P_FORM_ID_FVS901, P_FORM_ID_FVS903
            ' *** ヤード修繕 OR コンテナ(部屋)修繕 ***
            strSQL = "SELECT REPRT_UNIQEC " & Chr(13)     ' ユニークコード
            strSQL = strSQL & " ,REPRT_GENTD " & Chr(13)  ' 発生日
            strSQL = strSQL & " ,REPRT_TYPEC " & Chr(13)  ' 入力区分
            strSQL = strSQL & " ,TYPEC_NAME.CODET_NAMEN AS TYPEC_NAME" & Chr(13)  ' 入力区分名
            strSQL = strSQL & " ,REPRC_NAME.CODET_NAMEN AS REPRC_NAME" & Chr(13)  ' 修理区分名
            strSQL = strSQL & " ,REPRT_CNT1N " & Chr(13)  ' 内容1
            strSQL = strSQL & " ,REPRT_CNT2N " & Chr(13)  ' 内容2
            strSQL = strSQL & " ,REPRT_KINGA " & Chr(13)  ' 金額
            strSQL = strSQL & " ,REPRT_GYOUN " & Chr(13)  ' 業者名
            strSQL = strSQL & " ,REPRT_TANTC " & Chr(13)  ' 入力担当
            strSQL = strSQL & " ,TANTM_TANTN " & Chr(13)  ' 入力担当名
            '' ▼ 20101224 M.HONDA
            strSQL = strSQL & " ,REPRT_COMPD " & Chr(13)  ' 完了日
            '' ▲ 20101224 M.HONDA
            strSQL = strSQL & " FROM REPR_TRAN " & Chr(13)
            strSQL = strSQL & "      INNER JOIN TANT_MAST ON " & Chr(13)
            strSQL = strSQL & "      REPRT_TANTC  = TANTM_TANTC " & Chr(13)
            strSQL = strSQL & "      LEFT OUTER JOIN CODE_TABL AS TYPEC_NAME ON " & Chr(13)
            strSQL = strSQL & "      REPRT_TYPEC = TYPEC_NAME.CODET_CODEC " & Chr(13)
            strSQL = strSQL & "       AND TYPEC_NAME.CODET_SIKBC = '250'" & Chr(13)
            strSQL = strSQL & "      LEFT OUTER JOIN CODE_TABL AS REPRC_NAME ON " & Chr(13)
            strSQL = strSQL & "      REPRT_REPRC = REPRC_NAME.CODET_CODEC " & Chr(13)
            '▼ 2007/01/26 add tajima 修理区分名称の取得
            If pstrOpenForm = P_FORM_ID_FVS903 Then
                strSQL = strSQL & "       AND REPRC_NAME.CODET_SIKBC = '251'" & Chr(13)
            Else
                strSQL = strSQL & "       AND REPRC_NAME.CODET_SIKBC = '254'" & Chr(13)
            End If
            '▲ 2007/01/26 add
            strSQL = strSQL & " WHERE TANTM_BUMOC = '" & pstrBumonCd & "'   " & Chr(13)
            strSQL = strSQL & "   AND REPRT_BUMOC = '" & pstrBumonCd & "'   " & Chr(13)   '--2010/09/11--ryu---add---
            strSQL = strSQL & "   AND REPRT_YARDC = '" & Me.cmb_ヤードコード & "' " & Chr(13)

            If pstrOpenForm = P_FORM_ID_FVS901 Then
                ' 入力対象
                strSQL = strSQL & " AND  REPRT_INPTI = '" & CmReprMod.P_対象_ヤード & "' "

            ElseIf pstrOpenForm = P_FORM_ID_FVS903 Then
                ' コンテナ(部屋)修繕のとき
                strSQL = strSQL & " AND REPRT_ROOMC = '" & Me.cmb_コンテナ番号 & "' " & Chr(13)

                ' 入力対象
                strSQL = strSQL & " AND  REPRT_INPTI = '" & CmReprMod.P_対象_部屋 & "' "
            End If

            strSQL = strSQL & " ORDER BY REPRT_GENTD DESC"

        Case P_FORM_ID_FVS902
            ' *** 顧客修繕 ***
            strSQL = "SELECT KOTOT_UNIQEC " & Chr(13)     ' ユニークコード
            strSQL = strSQL & " ,KOTOT_GENTD " & Chr(13)  ' 発生日
            strSQL = strSQL & " ,KOTOT_TYPEC " & Chr(13)  ' 入力区分
            strSQL = strSQL & " ,CODET_NAMEN " & Chr(13)  ' 入力区分名
            strSQL = strSQL & " ,KOTOT_CONTN " & Chr(13)  ' 内容
            strSQL = strSQL & " ,KOTOT_TANTC " & Chr(13)  ' 入力担当
            strSQL = strSQL & " ,TANTM_TANTN " & Chr(13)  ' 入力担当名
            strSQL = strSQL & " FROM KOTO_TRAN " & Chr(13)
            strSQL = strSQL & "      INNER JOIN TANT_MAST ON                    " & Chr(13)
            strSQL = strSQL & "      KOTOT_TANTC  = TANTM_TANTC                 " & Chr(13)
            strSQL = strSQL & "      INNER JOIN CODE_TABL ON                    " & Chr(13)
            strSQL = strSQL & "      KOTOT_TYPEC = CODET_CODEC                  " & Chr(13)
            strSQL = strSQL & " WHERE TANTM_BUMOC = '" & pstrBumonCd & "'       " & Chr(13)
            strSQL = strSQL & "   AND KOTOT_BUMOC = '" & pstrBumonCd & "'       " & Chr(13)  '--2010/09/11--ryu---add---
            strSQL = strSQL & "   AND CODET_SIKBC = '" & P_NAME_ID_250 & "'     " & Chr(13)
            strSQL = strSQL & "   AND KOTOT_KOKYC = '" & Me.cmb_顧客コード & "' " & Chr(13)
            strSQL = strSQL & "   AND KOTOT_DISPI = 1                           " & Chr(13)  'INSERT 2010/06/10 SHIBAZAKI
            strSQL = strSQL & " ORDER BY KOTOT_GENTD DESC "
    End Select

    ' SQL文発行
    Set objRs = pobjKASE_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    If objRs.EOF = False Then
        ' ワークテーブルオープン
        strSQL = "SELECT * FROM FVS901_W01"
        Set objRsWk = objDb.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

        Do Until objRs.EOF
            With objRsWk
                .AddNew

                Select Case pstrOpenForm
                    Case P_FORM_ID_FVS901, P_FORM_ID_FVS903
                        strUniqec = "": strUniqec = StringFromGUID(objRs.Fields("REPRT_UNIQEC").VALUE)           ' ユニークコード
                        .Fields("FVS901_UNIQEC").VALUE = Mid(strUniqec, 7, Len(strUniqec) - 7)
                        .Fields("FVS901_GENTD").VALUE = Mid(objRs.Fields("REPRT_GENTD").VALUE, 1, 4) & "/" & _
                                                        Mid(objRs.Fields("REPRT_GENTD").VALUE, 5, 2) & "/" & _
                                                        Mid(objRs.Fields("REPRT_GENTD").VALUE, 7, 2)             ' 発生日
                        .Fields("FVS901_TYPE").VALUE = objRs.Fields("TYPEC_NAME").VALUE                         ' 区分名


                        '' ▼ 20101224 M.HONDA
                        ' 【完了日】
                        If Nz(objRs.Fields("REPRT_COMPD").VALUE, "") <> "" Then
                            .Fields("FVS901_COMPD").VALUE = Mid(objRs.Fields("REPRT_COMPD").VALUE, 1, 4) & "/" & _
                                                            Mid(objRs.Fields("REPRT_COMPD").VALUE, 5, 2) & "/" & _
                                                            Mid(objRs.Fields("REPRT_COMPD").VALUE, 7, 2)             ' 発生日
                        End If
                        '' ▼ 20101224 M.HONDA


                         .Fields("FVS901_GENTJ").VALUE = "00:00:00"    '' 2013/11/15 M.HONDA INS

                        ' 内容編集処理
                        varName = ""
                        ' 【修理区分名】
                        If Nz(objRs.Fields("REPRC_NAME").VALUE, "") <> "" Then
                            varName = objRs.Fields("REPRC_NAME").VALUE & Space(1)
                        End If
                        
                        ' 【内容1】
                        If Nz(objRs.Fields("REPRT_CNT1N").VALUE, "") <> "" Then
                            varName = varName & objRs.Fields("REPRT_CNT1N").VALUE & Space(1)
                        End If
                        ' 【内容2】
                        If Nz(objRs.Fields("REPRT_CNT2N").VALUE, "") <> "" Then
                            varName = varName & objRs.Fields("REPRT_CNT2N").VALUE & Space(1)
                        End If
                        ' 【金額】
                        If Nz(objRs.Fields("REPRT_KINGA").VALUE, "") <> "" Then
                            varName = varName & "\" & Format(objRs.Fields("REPRT_KINGA").VALUE, "###,###,###") & Space(1)
                        End If
                        ' 【業者名】
                        If Nz(objRs.Fields("REPRT_GYOUN").VALUE, "") <> "" Then
                            varName = varName & objRs.Fields("REPRT_GYOUN").VALUE & Space(1)
                        End If
                        .Fields("FVS901_NAIYO").VALUE = varName                                                  ' 内容

                        .Fields("FVS901_TANTNM").VALUE = objRs.Fields("TANTM_TANTN").VALUE                       ' 担当者名
                        
                    Case P_FORM_ID_FVS902
                        strUniqec = "": strUniqec = StringFromGUID(objRs.Fields("KOTOT_UNIQEC").VALUE)           ' ユニークコード
                        strUniqec = Mid(strUniqec, 7, Len(strUniqec) - 7)
                        .Fields("FVS901_UNIQEC").VALUE = strUniqec
                        .Fields("FVS901_GENTD").VALUE = Mid(objRs.Fields("KOTOT_GENTD").VALUE, 1, 4) & "/" & _
                                                        Mid(objRs.Fields("KOTOT_GENTD").VALUE, 5, 2) & "/" & _
                                                        Mid(objRs.Fields("KOTOT_GENTD").VALUE, 7, 2)             ' 発生日
                        .Fields("FVS901_TYPE").VALUE = objRs.Fields("CODET_NAMEN").VALUE                         ' 区分
                        .Fields("FVS901_NAIYO").VALUE = objRs.Fields("KOTOT_CONTN").VALUE                        ' 内容
                        .Fields("FVS901_TANTNM").VALUE = objRs.Fields("TANTM_TANTN").VALUE                       ' 担当者名
                End Select

                .UPDATE
            End With

            objRs.MoveNext
        Loop
        
        objRsWk.Close
        Set objRsWk = Nothing
    Else
        ' 該当データなし
    End If
    objRs.Close
    Set objRs = Nothing

    GoTo subCreateWork_Exit

subCreateWork_Err:
    blnErrFlg = True

subCreateWork_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objRsWk Is Nothing Then objRsWk.Close: Set objRsWk = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing

    If blnErrFlg = True Then
        Call Err.Raise(Err.Number, "subCreateWork" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :修繕トランまたは顧客問合わせトランの削除
'        MODULE_ID        :fncUpdFVS901
'        PARAMETER        :strUNIQEC(IN) = ユニークコード
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subUpdFVS901(ByVal strUniqec As String)

    Dim strSQL      As String
    Dim objRs       As Recordset
    Dim WRI_ObjWS   As Workspace
    Dim blnErrFlg   As Boolean
    Dim reprtTYPEC  As String
    Dim fvs910Form  As Form

    If Nz(strUniqec, "") = "" Then Exit Sub
    
    On Error Resume Next
    Set fvs910Form = Forms("FVS910")
    If Not fvs910Form Is Nothing Then
        MsgBox "修繕入力画面が開いているため削除はできません。" & vbCr & "修繕入力画面を閉じて下さい。", vbInformation, P_MSG_KAKUNIN
        fvs910Form.SetFocus
        Exit Sub
    End If
    
    On Error GoTo subUpdFVS901_Err
    
    If pIsAuthority = False Then
        MsgBox "削除をする権限がありません。", vbInformation, P_MSG_KAKUNIN
        Exit Sub
    End If

    ' 確認メッセージ
    If MsgBox("削除します、よろしいですか？", vbQuestion + vbYesNo, P_MSG_KAKUNIN) = vbNo Then
        Exit Sub
    End If

    ' 加瀬DBに接続
    Call subConnectDB(P_DB_KASE)
    ' 2006/09/06 KOMS系接続 部屋修繕のときのみ開く
    If pstrOpenForm = P_FORM_ID_FVS903 Then Call subConnectDB(P_DB_KONT, pstrBumonCd)

    ' トランザクション開始
    Set WRI_ObjWS = DBEngine.Workspaces(0)
    WRI_ObjWS.BeginTrans

    strSQL = "SELECT * FROM "

    Select Case pstrOpenForm
        Case P_FORM_ID_FVS901, P_FORM_ID_FVS903
            ' *** ヤード修繕一覧画面 OR コンテナ(部屋)修繕一覧画面 ***
            strSQL = strSQL & " REPR_TRAN "
            strSQL = strSQL & " WHERE REPRT_UNIQEC = '" & strUniqec & "' "

        Case P_FORM_ID_FVS902
            ' *** 顧客修繕一覧画面 ***
            strSQL = strSQL & " KOTO_TRAN "
            strSQL = strSQL & " WHERE KOTOT_UNIQEC = '" & strUniqec & "' "
    End Select

    Set objRs = pobjKASE_DB.OpenRecordset(strSQL)

    ' データ削除
    With objRs
        If pstrOpenForm = P_FORM_ID_FVS903 Then reprtTYPEC = .Fields("REPRT_TYPEC")
        .Edit
        .Delete
    End With

    ' 2006/09/06 部屋修繕のときのみKOMS系への反映を行う
    If pstrOpenForm = P_FORM_ID_FVS903 Then Call subKont系反映(strUniqec, reprtTYPEC)

    ' コミット
    WRI_ObjWS.CommitTrans

    ' サブフォーム再作成
    Call subFormLoad

    ' 削除が成功したならば変更された！
    pIsChanged = True  '2006/09/06
    
    GoTo subUpdFVS901_Exit

subUpdFVS901_Err:
    ' ロールバック
    WRI_ObjWS.Rollback
    blnErrFlg = True

subUpdFVS901_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKASE_DB Is Nothing Then Call subDisConnectDB(P_DB_KASE)
    If Not pobjKONT_DB Is Nothing Then: Call subDisConnectDB(P_DB_KONT)

    If blnErrFlg = True Then
        Call Err.Raise(Err.Number, "subUpdFVS901" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :KONT系ＤＢへの修繕内容反映処理
'        MODULE_ID        :subKont系反映
'        PARAMETER        :
'        CREATE_DATE      :2006/09/06
'        NOTE             :KONT系DBとしてデータ削除に伴うCNTA_MASTへの反映を行う
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKont系反映(ByVal a削除ユニークコード As String, ByVal a修繕タイプ As String)

    Dim strSQL            As String
    Dim rsObject          As Recordset
    Dim isNowRepairing    As Boolean ' 現在修理中なのか
    
  On Error GoTo Exception
    
    isNowRepairing = CmReprMod.IsRepairing(pobjKASE_DB, pstrBumonCd, Me.cmb_ヤードコード, Me.cmb_コンテナ番号, a削除ユニークコード)
    
    ' ヤードコード＆コンテナコードで更新対象抽出
    strSQL = "SELECT * "
    strSQL = strSQL & " FROM CNTA_MAST WHERE CNTA_CODE = " & Me.cmb_ヤードコード
    strSQL = strSQL & " AND CNTA_NO = " & Me.cmb_コンテナ番号
    
    Set rsObject = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    rsObject.Edit
    Select Case a修繕タイプ
        Case CmReprMod.P_修繕_修理
            ' 現状、修理中があるかチェック
            If isNowRepairing = True Then
               rsObject.Fields("CNTA_REPR_KBN") = 1
               rsObject.Fields("CNTA_USE") = 0
            Else
              rsObject.Fields("CNTA_REPR_KBN") = 0
              rsObject.Fields("CNTA_USE") = 1
            End If
        
        Case CmReprMod.P_修繕_撤去予定
            ' 撤去が削除されたらコンテナは利用可に戻す
             rsObject.Fields("CNTA_USE") = 1
    End Select
    ' 更新日設定
    rsObject.Fields("CNTA_UPDATE") = Format(Now, "yyyy/mm/dd")
    rsObject.UPDATE

    rsObject.Close
    Set rsObject = Nothing
    
    Exit Sub
   
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "subKont系反映" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
' ☆☆☆ DB ACCESS系関数群 ☆☆☆

'==============================================================================*
'
'        MODULE_NAME      :前画面の再表示
'        MODULE_ID        :beforFormReload
'        PARAMETER        :
'        CREATE_DATE      :2006/09/24
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub beforFormReload()

    Dim beforForm As Form

    On Error Resume Next

    Set beforForm = Forms("FKS020_コンテナ予約入力")
    If Not beforForm Is Nothing Then
        Call beforForm.Reload(beforForm.cmb_ヤードコード, beforForm.cmb_ヤードコード)
        Set beforForm = Nothing
    End If
    
    Set beforForm = Forms("FVS500")
    If Not beforForm Is Nothing Then
        Call beforForm.SubReloadfromFVS90n
        Set beforForm = Nothing
    End If
    
    Set beforForm = Forms("FVS400")
    If Not beforForm Is Nothing Then
        Call beforForm.SubRefreshfromFVS90n
        Set beforForm = Nothing
    End If
    
    Set beforForm = Forms("FVS600")
    If Not beforForm Is Nothing Then
        Call beforForm.SubReloadfromFVS90n
        Set beforForm = Nothing
    End If
    
    Set beforForm = Forms("FVS700")
    If Not beforForm Is Nothing Then
        Call beforForm.SubReloadfromFVS90n
        Set beforForm = Nothing
    End If
    
    
    Set beforForm = Forms("FVS021")
    If Not beforForm Is Nothing Then
        Call beforForm.Reload(beforForm.txt_YARD_CODE, beforForm.txt_YARD_CODE)
    Else
        Set beforForm = Forms("FVS011")
        If Not beforForm Is Nothing And pstrOpenForm = P_FORM_ID_FVS903 Then
          ' 空き検索画面は部屋修繕のみ（修理にしたり撤去が出来るので）
            Call beforForm.Reload(Me.cmb_ヤードコード, Me.cmb_ヤードコード)
        End If
    End If
    
End Sub
' ☆☆☆ Public関数群 ☆☆☆
'==============================================================================*
'
'        MODULE_NAME      :外部から再表示指令
'        MODULE_ID        :OutSideReload
'        PARAMETER        :
'        CREATE_DATE      :2006/05/13
'        UPDATE_DATE      :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub OutSideReload()
    
    Dim isException As Boolean

  On Error GoTo Exception
    isException = False

    ' 加瀬DBに接続
    Call subConnectDB(P_DB_KASE)

    ' 再表示
    Call subFormLoad
    
    ' 再表示依頼があったといことは変更された！
    pIsChanged = True  '2006/09/06
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not pobjKASE_DB Is Nothing Then Call subDisConnectDB(P_DB_KASE)
    If isException = True Then
        Call Err.Raise(Err.Number, "OutSideReload" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub
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

' ☆☆☆ Public関数群 ☆☆☆
'****************************  ended or program ********************************
