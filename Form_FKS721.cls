VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS721"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　ヤード別割引設定
'   プログラムＩＤ　：　FKS721
'   作　成　日　　　：  2008/08/09
'   作　成　者　　　：  K.KINEBUCHI
'   Ver            ：  0.0
'**********************************************
'   修　正　日　　　：  2008/10/25
'   修　正　者　　　：  K.KINEBUCHI
'   内　　　容　　　：  １）適用優先順位を発番する際、レンタル物件割引きトランの
'   　　　　　　　　　　　　ヤードコードのみで発番するよう修正した。
'                  ：  ２）実施期間を「無期限」にした場合適用優先順を100にしてい
'                       　 た箇所を修正する。
'   Ver            ：  0.1
'**********************************************
'   修　正　日　　　：  2008/11/13
'   修　正　者　　　：  N.IMAI
'   内　　　容　　　：  次画面へ文言を引き数で渡す際に
'   　　　　　　　　　　区切り文字がカンマの場合、不具合があるので
'   　　　　　　　　　　区切り文字をタブに変更
'   Ver            ：  0.2
'**********************************************
'**********************************************
'   修　正　日　　　：  2009/11/06
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  画面初期表示の割引実施期間をデフォルト無期限に設定
'   Ver             ：  0.3
'**********************************************
'**********************************************
'   修　正　日　　　：  2013/06/04
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  コンテナの割引を選択時に、コンテナ（バイク）で同一の割引があれば適用させる。
'   Ver             ：  0.4
'**********************************************
'**********************************************
'   修　正　日　　　：  2013/07/31
'   修　正　者　　　：  M.HONDA
'   内　　　容　　　：  期間のチェックを外す
'   Ver             ：  0.5
'**********************************************

Option Compare Database
Option Explicit

Private Const P_KAKUNIN = "確認"
Private Const P_ERROR = "エラー"
Private Const P_TITLE_UPDATE = "（修正中）"

Private Const PROG_ID = "FKS721"

Private Const P_NEBIKI_WORK_TABLE   As String = "WK_FKS721"

Private Const MODE_INSERT = 1
Private Const MODE_UPDATE = 2

Private Const P_ORDER_99            As Integer = 99
Private Const P_ORDER_100           As Integer = 100
Private Const P_DATE_MIN            As String = "00000000"  '実施期間無期限MIN
Private Const P_DATE_MAX            As String = "99999999"  '実施期間無期限MAX
Private Const P_MONGON_LENGTH       As Integer = 42         '提示文言
Private Const P_BIKO_LENGTH         As Integer = 60         '備考

Private pobjDB                      As Object

'引数格納用
Private pstrParameter()             As String               '初期表示データ

'各引数格納用
Private pstrBumonCode               As String               '部門コード
Private pstrProcMode                As String               '処理モード
Private pstrYCode                   As String               'ヤードコード
Private pstrNo                      As String               '割引番号
Private pstrEnable                  As String               '有効可否
Private pstrDateFrom                As String               '実施期間From
Private pstrDateTo                  As String               '実施期間To

Private vDatefrom                   As Date                 '初期設定用
Private vDateto                     As Date                 '初期設定用
Private vSysDate                    As String               'システム日付

Private pboolFrmActive              As Boolean
Private pstrDateValue               As String               ' 日付の入力チェック用
Private gstrUserName                As String

'==============================================================================*
'
'        MODULE_NAME      :フォームロード
'        MODULE_ID        :Form_Load
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Dim isException    As Boolean

    isException = False

On Error GoTo Exception

    pboolFrmActive = False

    pstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If pstrBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    Set pobjDB = CurrentDb
    
    gstrUserName = MSZZ000.LsGetUserName
 
    ' パラメータ設定
    'D:2008.11.12 N.imai pstrParameter = Split(Nz(Me.OpenArgs), ",")
    pstrParameter = Split(Nz(Me.OpenArgs), vbTab)                               'I:2008.11.12 N.imai

    '初期設定
    vDatefrom = CDate(Format$(Now, "yyyy/mm") & "/01")          'システム日付の翌月1日
    vDatefrom = DateAdd("m", 1, vDatefrom)
    vDateto = CDate(Format$(Now, "yyyy/mm") & "/01")            'システム日付の翌月末日
    vDateto = DateAdd("m", 2, vDateto)
    vDateto = DateAdd("d", -1, vDateto)
    vSysDate = CStr(Format$(Now, "yyyymmdd"))                   'システム日付
    
    ' 初期表示項目設定処理
    Call setDefaultWindow

    pboolFrmActive = True

    GoTo Finally

Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Load", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :初期表示項目設定処理
'        MODULE_ID        :setDefaultWindow
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setDefaultWindow()

    Dim bMoveYard      As Boolean
    Dim intLoopNo      As Integer
    
On Error GoTo Exception
    
    '割引マスタ一覧より（追加）　(0)タイトル - (1)モード - (2)ヤードコード - (3)部門コード - (4)割引番号 - (5)手動文言
    'ヤード別割引一覧より（修正）(0)タイトル - (1)モード - (2)ヤードコード - (3)部門コード - (4)割引番号 - (5)有効可否 - (6)実施期間From - (7)実施期間To - (8)提示文言 - (9)自動文言 - (10)備考
    
    '共通項目セット
    Me.lbl_タイトル.Caption = pstrParameter(0)                      'タイトル（新規追加 or 修正）
    pstrProcMode = pstrParameter(1)                                '処理モード（0=新規追加 1=修正）
    pstrYCode = pstrParameter(2)                                   'ヤードコード
    pstrNo = pstrParameter(4)                                      '割引番号
    
    If pstrProcMode = "0" Then
        '新規追加モード初期設定
        Me.cmd_処理.Caption = "追加"                                '処理ボタン
        Me.opt_無効.Enabled = False                                 '有効可否「無効」を選択不可とする
        
        '' 2009/11/06 Ver0.3 M.HONDA START
        'Me.fra_実施期間グループ.VALUE = 2                            '実施期間指定を規定値
        'Me.txt_実施期間from.VALUE = vDatefrom                       '実施期間fromにシステム日付の翌月１日を規定値
        'Me.txt_実施期間to.VALUE = vDateto                           '実施期間toにシステム日付の翌月末日を規定値
        'Me.txt_実施期間from.Enabled = True                          '実施期間from入力可
        'Me.txt_実施期間to.Enabled = True                            '実施期間to入力可
        Me.fra_実施期間グループ.VALUE = 1
        Me.txt_実施期間From.VALUE = ""                           '無期限につき未入力
        Me.txt_実施期間From.VALUE = ""                           '無期限につき未入力
        Me.txt_実施期間From.Enabled = False                      '実施期間from入力不可
        Me.txt_実施期間To.Enabled = False                        '実施期間to入力可不可
        '' 2009/11/06 Ver0.3 M.HONDA END
        Me.lbl_標準文言.Caption = pstrParameter(5)                  '割引マスタ手動文言（割引マスタ一覧_FKS700より）を規定値
        Me.txt_提示文言.VALUE = pstrParameter(5)                    '割引マスタ手動文言（割引マスタ一覧_FKS700より）を規定値
        Me.txt_備考.VALUE = ""                                      '空白を規定値
    Else
        '修正モード初期設定
        Me.cmd_処理.Caption = "保存"                                '処理ボタン
        Me.opt_無効.Enabled = True                                 '有効可否「無効」を選択可とする
        pstrEnable = pstrParameter(5)                              '有効可否
        '有効可否振り分け
        If pstrEnable = "1" Then
            '有効
            Me.fra_有効可否グループ.VALUE = 1
        Else
            '無効
            Me.fra_有効可否グループ.VALUE = 2
        End If
        
        pstrDateFrom = pstrParameter(6)                             '実施期間FROM
        pstrDateTo = pstrParameter(7)                               '実施期間TO
        '実施期間振り分け
        If pstrDateFrom = P_DATE_MIN And pstrDateTo = P_DATE_MAX Then
            Me.fra_実施期間グループ.VALUE = 1
            Me.txt_実施期間From.VALUE = ""                           '無期限につき未入力
            Me.txt_実施期間From.VALUE = ""                           '無期限につき未入力
            Me.txt_実施期間From.Enabled = False                      '実施期間from入力不可
            Me.txt_実施期間To.Enabled = False                        '実施期間to入力可不可
        Else
            Me.fra_実施期間グループ.VALUE = 2
            Me.txt_実施期間From.VALUE = Format$(Left(pstrDateFrom, 4) & "/" & Mid(pstrDateFrom, 5, 2) & "/" & Right(pstrDateFrom, 2)) '実施期間from
            Me.txt_実施期間To.VALUE = Format$(Left(pstrDateTo, 4) & "/" & Mid(pstrDateTo, 5, 2) & "/" & Right(pstrDateTo, 2))         '実施期間from
            Me.txt_実施期間From.Enabled = True                      '実施期間from入力可
            Me.txt_実施期間To.Enabled = True                        '実施期間to入力可
        End If
        
        Me.txt_提示文言.VALUE = pstrParameter(8)                    'レンタル物件割引トラン提示文言（ヤード別割引一覧_FKS720より）を規定値
        lbl_標準文言.Caption = pstrParameter(9)                     '割引マスタ手動文言（ヤード別割引一覧_FKS720より）を規定値
        Me.txt_備考.VALUE = pstrParameter(10)                       'レンタル物件割引トラン備考（ヤード別割引一覧_FKS720より）を規定値
    
    End If
    
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "setDefaultWindow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fra_有効可否グループ_AfterUpdate
'        MODULE_ID        :fra_有効可否グループ_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub fra_有効可否グループ_AfterUpdate()
  
    Call subCaptionChange
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fra_実施期間グループ_AfterUpdate
'        MODULE_ID        :fra_実施期間グループ_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UODATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub fra_実施期間グループ_AfterUpdate()
  
    If fra_実施期間グループ.VALUE = 1 Then
        txt_実施期間From.VALUE = ""
        txt_実施期間To.VALUE = ""
        txt_実施期間From.Enabled = False
        txt_実施期間To.Enabled = False
    Else
        txt_実施期間From.VALUE = vDatefrom
        txt_実施期間To.VALUE = vDateto
        txt_実施期間From.Enabled = True
        txt_実施期間To.Enabled = True
    End If
    
    Call subCaptionChange
  
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_実施期間from_BeforeUpdate
'        MODULE_ID        :txt_実施期間from_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_実施期間from_BeforeUpdate(Cancel As Integer)
    
    If Nz(Me.txt_実施期間From.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_実施期間From.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "実施期間はYYYY/MM/DD形式です。"
            Cancel = True
            SendKeys "{ESC}"
        End If
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_実施期間from_AfterUpdate
'        MODULE_ID        :txt_実施期間from_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_実施期間from_AfterUpdate()

    Me.txt_実施期間From.VALUE = pstrDateValue
    Call subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_実施期間to_BeforeUpdate
'        MODULE_ID        :txt_実施期間to_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_実施期間to_BeforeUpdate(Cancel As Integer)
    
    If Nz(Me.txt_実施期間To.VALUE) <> "" Then
        pstrDateValue = Nz(CmDateChecker(Nz(Me.txt_実施期間To.VALUE)))
        If pstrDateValue = "" Then
            MsgBox "実施期間はYYYY/MM/DD形式です。"
            Cancel = True
            SendKeys "{ESC}"
        End If
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_実施期間to_AfterUpdate
'        MODULE_ID        :txt_実施期間to_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_実施期間to_AfterUpdate()

    Me.txt_実施期間To.VALUE = pstrDateValue
    Call subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_提示文言_BeforeUpdate
'        MODULE_ID        :txt_提示文言_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_提示文言_BeforeUpdate(Cancel As Integer)

    '入力桁数チェック
    Call fncChkTextInput(txt_提示文言, P_MONGON_LENGTH, Cancel, _
                               "桁数オーバーです。" & vbCrLf & "提示文言の入力は半角42桁までです。")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_提示文言_AfterUpdate
'        MODULE_ID        :txt_提示文言_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_提示文言_AfterUpdate()

    Call fncAfterUpdateCommon(txt_提示文言, P_MONGON_LENGTH)
    Call subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_備考_BeforeUpdate
'        MODULE_ID        :txt_備考_BeforeUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_備考_BeforeUpdate(Cancel As Integer)

    '入力桁数チェック
    Call fncChkTextInput(txt_備考, P_BIKO_LENGTH, Cancel, _
                               "桁数オーバーです。" & vbCrLf & "備考の入力は半角60桁までです。")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_備考_AfterUpdate
'        MODULE_ID        :txt_備考_AfterUpdate
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_備考_AfterUpdate()

    Call fncAfterUpdateCommon(txt_備考, P_BIKO_LENGTH)
    Call subCaptionChange

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncAfterUpdateCommon
'        機能             :更新後イベント時の処理
'                         :数値項目は０埋めまたは０消しした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :年月は"yyyy/mm"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncAfterUpdateCommon(objTarget As Object, intLength As Integer) As String
    
On Error GoTo Exception

    fncAfterUpdateCommon = ""
    
    '体裁を整える
    fncAfterUpdateCommon = Nz(objTarget)

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("fncAfterUpdateCommon", True)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkTextInput
'        機能             :テキストボックス入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTextInput(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer, Optional strErrMsg As String) As Boolean
    
On Error GoTo Exception

    If Nz(txtCheck) = "" Then
        fncChkTextInput = True
        Exit Function
    End If

    If CmLenb(txtCheck) > intMaxLength Then
        MsgBox strErrMsg, vbInformation, P_ERROR
        Cancel = True
        fncChkTextInput = False
    Else
        fncChkTextInput = True
    End If

Exception:
    If Err <> 0 Then
        Call MSZZ024_M00("fncChkTextInput", True)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :キャプション変更処理
'        MODULE_ID        :subCaptionChange
'        NOTE             :画面タイトル（修正中）を表示する。
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subCaptionChange()
    
    If InStr(1, Me.Caption, P_TITLE_UPDATE) <= 0 Then
        Me.Caption = Me.Caption & P_TITLE_UPDATE
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :処理ボタン押下時処理
'        MODULE_ID        :cmd_処理_Click
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_処理_Click()
    
    Dim objAdoDbConnection As Object
    Dim isExecuteFlg As Boolean
    Dim isException  As Boolean
    isException = False

On Error GoTo Exception
    
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(pstrBumonCode)
    
    '入力チェック
    If fncInputCheck(objAdoDbConnection) Then
        '追加処理
        If pstrProcMode = "0" Then
            If MsgBox("追加します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
                '新規
                isExecuteFlg = fncHozon(MODE_INSERT, objAdoDbConnection)
            End If
        Else
            If MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
                '修正
                isExecuteFlg = fncHozon(MODE_UPDATE, objAdoDbConnection)
            End If
        End If
    End If
    
    If isExecuteFlg Then
        If pstrProcMode = "0" Then
            '新規追加（正常終了後、メッセージを表示し、一覧再読み込み）
            Call MsgBox("追加しました。", vbInformation + vbOKOnly, P_KAKUNIN)
            Call Form_FKS700.createNebikiData
            Call Form_FKS700.Requery
        Else
            '修正（正常終了後、メッセージを表示し、一覧再読み込み）
            Call MsgBox("更新しました。", vbInformation + vbOKOnly, P_KAKUNIN)
        End If
        '各一覧に戻る
        Call CloseWindow
    End If
    
    GoTo Finally

Exception:
    isException = True

Finally:
    If Not objAdoDbConnection Is Nothing Then
        objAdoDbConnection.Close
        Set objAdoDbConnection = Nothing
    End If
    
    If isException Then Call MSZZ024_M00("cmd_処理_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力チェック
'        MODULE_ID        :fncInputCheck
'        NOTE             :objAdoDbConnection コネクション
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInputCheck(objAdoDbConnection As Object) As Boolean
    
    Dim strSQL             As String
    Dim rsData             As Object
    Dim strNewDateFrom     As String        '実施期間from格納用
    Dim strNewDateTo       As String        '実施期間to格納用
    Dim strSetDate         As String
    
On Error GoTo Exception
    
    '初期値設定
    fncInputCheck = False
    
    '追加モード（新規作成）で「有効可否」の変更はあり得ないが、常にチェックする。
    If pstrProcMode = "0" Then
        If fra_有効可否グループ.VALUE = 2 Then
            Call MsgBox("追加処理での有効可否は「無効」にはできません。", vbExclamation + vbOKOnly, P_ERROR)
            fra_有効可否グループ.VALUE = 1
            Me.fra_有効可否グループ.SetFocus
            Exit Function
        End If
    End If
    
    '実施期間選択時チェック
    If Me.fra_実施期間グループ.VALUE = 2 Then
        
        '実施期間from必須入力チェック
        If Nz(Me.txt_実施期間From.VALUE, "") = "" Then
            Call MsgBox("必須設定としています。日付を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.txt_実施期間From.SetFocus
            Exit Function
        End If
        
        '実施期間to必須入力チェック
        If Nz(Me.txt_実施期間To.VALUE, "") = "" Then
            Call MsgBox("必須設定としています。日付を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.txt_実施期間To.SetFocus
            Exit Function
        End If
    
        '実施期間from日付形式チェック
        If Not IsDate(Nz(Me.txt_実施期間From.VALUE)) Then
            Call MsgBox("実施期間はYYYY/MM/DD形式です。", vbExclamation + vbOKOnly, P_ERROR)
            Me.txt_実施期間From.SetFocus
            Exit Function
        End If
    
        '実施期間to日付形式チェック
        If Not IsDate(Nz(Me.txt_実施期間To.VALUE)) Then
            Call MsgBox("実施期間はYYYY/MM/DD形式です。", vbExclamation + vbOKOnly, P_ERROR)
            Me.txt_実施期間To.SetFocus
            Exit Function
        End If

    End If
    
    '提示文言必須入力チェック
    If Nz(txt_提示文言.VALUE, "") = "" Then
        Call MsgBox("必須設定としています。提示文言を入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
        Me.txt_提示文言.SetFocus
        Exit Function
    End If
    
    '実施期間の選択により、日付をセット
    If Me.fra_実施期間グループ.VALUE = 1 Then
        strNewDateFrom = P_DATE_MIN
        strNewDateTo = P_DATE_MAX
    Else
        strNewDateFrom = Format(Me.txt_実施期間From.VALUE, "yyyymmdd")
        strNewDateTo = Format(Me.txt_実施期間To.VALUE, "yyyymmdd")
        
        '実施期間from～to大小チェック
        If strNewDateFrom > strNewDateTo Then
            Call MsgBox("実施期間の指定が矛盾しています。" & vbCrLf & "From <= To となるように入力して下さい。", vbExclamation + vbOKOnly, P_ERROR)
            Me.txt_実施期間To.SetFocus
            Exit Function
        End If
        
        '' 2013/07/31 M.HONDA UPD
'        '実施期間toの入力で、システム日付より過去の日付はエラーとする。
'        If vSysDate > strNewDateTo Then
'            Call MsgBox("期間指定で過去の日付入力は行えません。", vbExclamation + vbOKOnly, P_ERROR)
'            Me.txt_実施期間To.SetFocus
'            Exit Function
'        End If
        ''  2013/07/31 M.HONDA END
    End If
    
    'レンタル物件割引きトランの実施期間が設定中で、同一の割引きと実施日の重りをチェック
    '※新規追加 or 修正且つ、画面入力した日付と引数で受け取った日付が違った場合チェックする
    If (pstrProcMode = "0") Or (pstrProcMode = "1" And pstrDateFrom <> strNewDateFrom Or pstrDateTo <> strNewDateTo) Then
        strSQL = ""
        strSQL = strSQL & "SELECT * FROM RLDN_TRAN "
        strSQL = strSQL & " WHERE RLDNT_YCODE = '" & pstrYCode & "' "
        strSQL = strSQL & " AND   RLDNT_NO    = '" & pstrNo & "' "
        If pstrProcMode = "1" Then
            strSQL = strSQL & " AND   RLDNT_FROM    <> '" & pstrDateFrom & "' "
        End If
        strSQL = strSQL & " AND   RLDNT_FROM <= '" & strNewDateTo & "' "
        strSQL = strSQL & " AND   RLDNT_TO   >= '" & strNewDateFrom & "' "
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        
        If Not rsData.EOF Then
            Call MsgBox("同じ割引が複数適用される可能性があります。" + vbCrLf + "実施期間は重ならないようにして下さい。", vbExclamation + vbOKOnly, P_ERROR)
            If Me.fra_実施期間グループ.VALUE <> 1 Then
                Me.txt_実施期間From.SetFocus
            End If
            Exit Function
        End If
        
        rsData.Close
        Set rsData = Nothing
    End If
    
    '有効可否が「有効」の選択で、レンタル物件割引きトランの適用優先順をチェック
    If fra_有効可否グループ.VALUE = 1 Then
        strSQL = ""
        strSQL = strSQL & "SELECT MAX(RLDNT_ORDER) RLDNT_ORDER FROM RLDN_TRAN "
        strSQL = strSQL & " WHERE RLDNT_YCODE = '" & pstrYCode & "' "
        strSQL = strSQL & " AND   RLDNT_NO    = '" & pstrNo & "' "
        If pstrProcMode = "1" Then
            strSQL = strSQL & " AND   RLDNT_FROM    = '" & pstrDateFrom & "' "
        End If
        strSQL = strSQL & " AND   RLDNT_ORDER <> " & P_ORDER_100 & " "
        
        Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
        
        If Not rsData.EOF Then
            If rsData.Fields("RLDNT_ORDER").VALUE = P_ORDER_99 Then
                Call MsgBox("最後尾の優先順位「99」を割り振れません。" + vbCrLf + "一覧に戻り優先順位「99」を前の順位にして下さい。", vbExclamation + vbOKOnly, P_ERROR)
                Me.fra_有効可否グループ.SetFocus
                Exit Function
            End If
        End If
    
        rsData.Close
        Set rsData = Nothing
    
    End If
    
    '入力チェック正常
    fncInputCheck = True
    Exit Function

Exception:
    
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    
    Call Err.Raise(Err.Number, "fncInputCheck" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :更新処理
'        MODULE_ID        :fncHozon
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncHozon(iMode As Integer, objAdoDbConnection As Object) As Boolean

    objAdoDbConnection.BeginTrans
    
On Error GoTo Exception

    fncHozon = False

    'DB更新処理
    Select Case iMode
        Case MODE_INSERT
            fncHozon = INS_RTN(objAdoDbConnection)
        Case MODE_UPDATE
            fncHozon = UPD_RTN(objAdoDbConnection)
    End Select

Exception:

    If fncHozon = True Then
        'コミット
        objAdoDbConnection.CommitTrans
    Else
        objAdoDbConnection.RollbackTrans
    End If
    
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncHozon" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      : 新規追加処理
'        MODULE_ID        : INS_RTN
'        PARAMETER        : objAdoDbConnection コネクション
'        戻り値           : True=正常 False=異常
'        CREATE_DATE      : 2008/08/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function INS_RTN(objAdoDbConnection As Object) As Boolean
  
    Dim strSQL             As String
    Dim rsData             As Object
    Dim rsDCNTData         As Object
    Dim strDate            As String
    Dim strTime            As String
    Dim strKaseSqlSvr      As String

On Error GoTo Exception
    
    INS_RTN = False

    'システム日付セット
    strDate = Format(DATE, "yyyymmdd")
    strTime = Format(Now, "hhmmss")
  
    'レンタル物件割引トラン登録
    Set rsData = MSZZ025.ADODB_Recordset("RLDN_TRAN", objAdoDbConnection, adoAppendOnly)

    rsData.AddNew
    rsData.Fields("RLDNT_BUMOC") = pstrBumonCode                                        '部門コード
    rsData.Fields("RLDNT_YCODE") = pstrYCode                                            'ヤードコード
    rsData.Fields("RLDNT_NO") = pstrNo                                                  '割引番号
    rsData.Fields("RLDNT_ENABLE") = "1"                                                 '有効可否（新規=1固定）
    
    '実施期間選択により以下振り分け
    If Me.fra_実施期間グループ.VALUE = "1" Then
        rsData.Fields("RLDNT_FROM") = P_DATE_MIN                                        '実施期間From（実施期間=無期限）
        rsData.Fields("RLDNT_TO") = P_DATE_MAX                                          '実施期間To（実施期間=無期限）
'        rsData.Fields("RLDNT_ORDER") = 100                                             '適用優先順（実施期間無期限=100固定）  'DELETE 2008/10/25 K.KINEBUCHI
    Else
        rsData.Fields("RLDNT_FROM") = Format(Me.txt_実施期間From.VALUE, "yyyymmdd")     '実施期間From（実施期間=指定有り）
        rsData.Fields("RLDNT_TO") = Format(Me.txt_実施期間To.VALUE, "yyyymmdd")         '実施期間From（実施期間=指定有り）
'        rsData.Fields("RLDNT_ORDER") = fncGetAutoOrderNo(objAdoDbConnection)           '適用優先順（現有効割引の最後尾 + 1）  'DELETE 2008/10/25 K.KINEBUCHI
    End If
    
    rsData.Fields("RLDNT_ORDER") = fncGetAutoOrderNo(objAdoDbConnection)               '適用優先順（現有効割引の最後尾 + 1）  'INSERT 2008/10/25 K.KINEBUCHI
    rsData.Fields("RLDNT_TEXT") = Me.txt_提示文言.VALUE                                 '提示文言
    rsData.Fields("RLDNT_NOTE") = Me.txt_備考.VALUE                                     '備考
    rsData.Fields("RLDNT_INSED") = strDate                                              '作成日付
    rsData.Fields("RLDNT_INSEJ") = strTime                                              '作成時刻
    rsData.Fields("RLDNT_INSPB") = PROG_ID                                              '作成プログラムID
    rsData.Fields("RLDNT_INSUB") = gstrUserName                                         '作成ユーザID
    
    '' 同一割引でコンテナ（バイク）が存在する場合には強制的に適用させる。
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    strSQL = ""
    strSQL = strSQL & "select B.* "
    strSQL = strSQL & "FROM  " & strKaseSqlSvr & ".dbo.DCNT_MAST B, "
    strSQL = strSQL & "    ( "
    strSQL = strSQL & "        select A.* "
    strSQL = strSQL & "        from " & strKaseSqlSvr & ".dbo.DCNT_MAST A "
    strSQL = strSQL & "        where A.DCNTM_NO = '" & pstrNo & "' "
    strSQL = strSQL & "        AND   A.DCNTM_USAGE = 0 "
    strSQL = strSQL & "        AND   DCNTM_BUMOC = '" & pstrBumonCode & "' "
    strSQL = strSQL & "    ) A "
    strSQL = strSQL & "where "
    strSQL = strSQL & "b.DCNTM_BUMOC = a.DCNTM_BUMOC "
    strSQL = strSQL & "AND B.DCNTM_KIND = A.DCNTM_KIND "
    strSQL = strSQL & "AND B.DCNTM_TYPE = A.DCNTM_TYPE "
    strSQL = strSQL & "AND B.DCNTM_VALUE = A.DCNTM_VALUE "
    strSQL = strSQL & "AND B.DCNTM_PERIOD = A.DCNTM_PERIOD "
    strSQL = strSQL & "AND ISNULL(B.DCNTM_SIZE_FROM,0) = ISNULL(A.DCNTM_SIZE_FROM,0) "
    strSQL = strSQL & "AND ISNULL(B.DCNTM_SIZE_TO,0) = ISNULL(A.DCNTM_SIZE_TO,0) "
    strSQL = strSQL & "AND ISNULL(B.DCNTM_FLOOR,0) = ISNULL(A.DCNTM_FLOOR,0) "
    strSQL = strSQL & "AND B.DCNTM_GENKBN = A.DCNTM_GENKBN "
    strSQL = strSQL & "AND B.DCNTM_USE_PERIOD = A.DCNTM_USE_PERIOD "
    strSQL = strSQL & "AND B.DCNTM_USAGE = 33 "
    
    Set rsDCNTData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    
    If Not rsDCNTData.EOF Then
        rsData.AddNew
        rsData.Fields("RLDNT_BUMOC") = pstrBumonCode                                        '部門コード
        rsData.Fields("RLDNT_YCODE") = pstrYCode                                            'ヤードコード
        rsData.Fields("RLDNT_NO") = rsDCNTData.Fields("DCNTM_NO")                               '割引番号
        rsData.Fields("RLDNT_ENABLE") = "1"                                                 '有効可否（新規=1固定）
        
        '実施期間選択により以下振り分け
        If Me.fra_実施期間グループ.VALUE = "1" Then
            rsData.Fields("RLDNT_FROM") = P_DATE_MIN                                        '実施期間From（実施期間=無期限）
            rsData.Fields("RLDNT_TO") = P_DATE_MAX                                          '実施期間To（実施期間=無期限）
        Else
            rsData.Fields("RLDNT_FROM") = Format(Me.txt_実施期間From.VALUE, "yyyymmdd")     '実施期間From（実施期間=指定有り）
            rsData.Fields("RLDNT_TO") = Format(Me.txt_実施期間To.VALUE, "yyyymmdd")         '実施期間From（実施期間=指定有り）
        End If
        
        rsData.Fields("RLDNT_ORDER") = fncGetAutoOrderNo(objAdoDbConnection)                '適用優先順（現有効割引の最後尾 + 1）  'INSERT 2008/10/25 K.KINEBUCHI
        rsData.Fields("RLDNT_TEXT") = Me.txt_提示文言.VALUE                                 '提示文言
        rsData.Fields("RLDNT_NOTE") = Me.txt_備考.VALUE                                     '備考
        rsData.Fields("RLDNT_INSED") = strDate                                              '作成日付
        rsData.Fields("RLDNT_INSEJ") = strTime                                              '作成時刻
        rsData.Fields("RLDNT_INSPB") = PROG_ID                                              '作成プログラムID
        rsData.Fields("RLDNT_INSUB") = gstrUserName                                         '作成ユーザID
        rsDCNTData.Close
        Set rsDCNTData = Nothing
    
    End If
    
    rsData.UPDATE
    
    rsData.Close
    Set rsData = Nothing

    
    
    
    '正常終了
    INS_RTN = True

Exception:

    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If

    If Err <> 0 Then
        Call Err.Raise(Err.Number, "INS_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      : 修正処理
'        MODULE_ID        : UPD_RTN
'        PARAMETER        : objAdoDbConnection コネクション
'        戻り値　　　　　  : True=正常 False=異常
'        CREATE_DATE      : 2008/08/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function UPD_RTN(objAdoDbConnection As Object) As Boolean
  
    Dim strSQL               As String
    Dim rsData               As Object
    Dim rsData2              As Object
    Dim rsDataLocal          As Object   'ワークファイル（部屋割付情報）読み込み用
    Dim strNewBoxNo          As String
    Dim strNewYCode          As String
    Dim strDate              As String
    Dim strTime              As String
  
On Error GoTo Exception
    
    UPD_RTN = False

    '初期値設定
    strDate = Format(DATE, "yyyymmdd")
    strTime = Format(Now, "hhmmss")
  
    'レンタル物件割引トラン登録
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM RLDN_TRAN "
    strSQL = strSQL & " WHERE RLDNT_YCODE = '" & pstrYCode & "' "
    strSQL = strSQL & " AND   RLDNT_NO    = '" & pstrNo & "' "
    strSQL = strSQL & " AND   RLDNT_FROM  = '" & pstrDateFrom & "' "
    
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    
    If Not rsData.EOF Then
        
        '有効可否選択により以下振り分け
        If Me.fra_有効可否グループ.VALUE = "1" Then
            rsData.Fields("RLDNT_ENABLE") = "1"                                             '有効可否（1=有効）
'INSERT 2008/10/25 K.KINEBUCHI start
            If pstrEnable <> Me.fra_有効可否グループ.VALUE Then
                rsData.Fields("RLDNT_ORDER") = fncGetAutoOrderNo(objAdoDbConnection)        '適用優先順（現有効割引の最後尾 + 1）
            End If
'INSERT 2008/10/25 K.KINEBUCHI end
        Else
            rsData.Fields("RLDNT_ENABLE") = "0"                                             '有効可否（0=無効）
            rsData.Fields("RLDNT_ORDER") = 100                                              '適用優先順（実施期間無期限=100固定）   'INSERT 2008/10/25 K.KINEBUCHI
        End If
        
        '実施期間選択により以下振り分け
        If Me.fra_実施期間グループ.VALUE = "1" Then
            rsData.Fields("RLDNT_FROM") = P_DATE_MIN                                        '実施期間From（実施期間=無期限）
            rsData.Fields("RLDNT_TO") = P_DATE_MAX                                          '実施期間To（実施期間=無期限）
'            rsData.Fields("RLDNT_ORDER") = 100                                             '適用優先順（実施期間無期限=100固定）   'DELETE 2008/10/25 K.KINEBUCHI
        Else
            rsData.Fields("RLDNT_FROM") = Format(Me.txt_実施期間From.VALUE, "yyyymmdd")     '実施期間From（実施期間=指定有り）
            rsData.Fields("RLDNT_TO") = Format(Me.txt_実施期間To.VALUE, "yyyymmdd")         '実施期間From（実施期間=指定有り）
'DELETE 2008/10/25 K.KINEBUCHI start
'            If pstrDateFrom = P_DATE_MIN And pstrDateTo = P_DATE_MAX Then
'                rsData.Fields("RLDNT_ORDER") = fncGetAutoOrderNo(objAdoDbConnection)       '適用優先順（現有効割引の最後尾 + 1）
'            End If
'DELETE 2008/10/25 K.KINEBUCHI end
        End If
        
        rsData.Fields("RLDNT_TEXT") = Me.txt_提示文言.VALUE                                 '提示文言
        rsData.Fields("RLDNT_NOTE") = Me.txt_備考.VALUE                                     '備考
        rsData.Fields("RLDNT_UPDAD") = strDate                                              '更新日付
        rsData.Fields("RLDNT_UPDAJ") = strTime                                              '更新時刻
        rsData.Fields("RLDNT_UPDPB") = PROG_ID                                              '更新プログラムID
        rsData.Fields("RLDNT_UPDUB") = gstrUserName                                         '更新ユーザID
        
        rsData.UPDATE
    
    End If
    
    rsData.Close
    Set rsData = Nothing
    
    '正常終了
    UPD_RTN = True

Exception:
    If Not rsData Is Nothing Then
        rsData.Close
        Set rsData = Nothing
    End If
    
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "UPD_RTN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :閉じるボタン押下時処理
'        MODULE_ID        :cmd_終了_Click
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_閉じる_Click()

    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    If InStr(1, Me.Caption, P_TITLE_UPDATE) > 0 Then
        If MsgBox("修正中ですが保存せずに終了しますか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
            '更新処理
            Call CloseWindow
        End If
    Else
        If MsgBox("終了します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
            Call CloseWindow
        End If
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("cmd_閉じる_Click", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :一覧に戻る処理
'        MODULE_ID        :CloseWindow
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CloseWindow()

On Error GoTo Exception

    doCmd.Close acForm, Me.FormName

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "CloseWindow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        NOTE             :
'        CREATE_DATE      :2008/08/09
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    
    Dim isException As Boolean
    isException = False

On Error GoTo Exception
    
    'コネクションクローズ
    If Not pobjDB Is Nothing Then
       pobjDB.Close
       Set pobjDB = Nothing
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If isException Then Call MSZZ024_M00("Form_Close", True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      : 適用優先順発番処理
'        MODULE_ID        : fncGetAutoOrder
'        PARAMETER        : 引数なし
'        戻り値　　　　　  : 発番した適用優先順
'        CREATE_DATE      : 2008/08/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetAutoOrderNo(objAdoDbConnection As Object) As Integer
  
    Dim strSQL               As String
    Dim rsODData             As Object

On Error GoTo Exception
    
    '初期値設定
    fncGetAutoOrderNo = 0

    'レンタル物件割引きトランの適用優先順 ＋1をセットする
    strSQL = ""
    strSQL = strSQL & "SELECT MAX(RLDNT_ORDER) AS RLDNT_ORDER FROM RLDN_TRAN "
    strSQL = strSQL & " WHERE RLDNT_YCODE = '" & pstrYCode & "' "
    'strSQL = strSQL & " AND   RLDNT_NO    = '" & pstrNo & "' "         'DELETE 2008/10/25 K.KINEBUCHI
    strSQL = strSQL & " AND   RLDNT_ORDER <> " & P_ORDER_100 & " "
    
    Set rsODData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection, adoReadWrite)
    
    If Not rsODData.EOF Then
        fncGetAutoOrderNo = Nz(rsODData.Fields("RLDNT_ORDER")) + 1
    End If

    '正常終了
    rsODData.Close
    Set rsODData = Nothing
    Exit Function

Exception:
    If Not rsODData Is Nothing Then
        rsODData.Close
        Set rsODData = Nothing
    End If
    
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncGetAutoOrderNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
