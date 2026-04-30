VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS720"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　解約一覧
'   プログラムＩＤ　：　FVS720
'   作　成　日　　　：  2005/07/29
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴
'   修　正　日　　　：2005/10/02
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：サブフォームで保存完了時の処理追加
'   Ver             ：  0.1
'
'   修　正　日　　　：2006/02/03
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：条件に顧客カナ（部分一致）追加
'   Ver             ：  0.3
'
'   修　正　日　　　：2012/07/07
'   修　正　者      ：M.HONDA
'   修　正　内　容　：検索条件の部屋追加
'   Ver             ：  0.4
'
'   修　正　日　　　：2019/04/10
'   修　正　者      ：Y_SUZUKI
'   修　正　内　容　：検索条件に「解約担当者」を追加
'   Ver             ：  0.5
'
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'接続ＤＢオブジェクト
Private dbSQLServer     As Database
Private dbAccess        As Database

'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'部門コード
Private strBumonCode    As String

'編集モードフラグ　入力値が変更された場合ONになる
Public blnChangeMode As Boolean

'フォームのキャプションを保存する
Private strFormCaption As String

'検索条件
Private strCondCancelDateFrom                   As String
Private strCondCancelDateTo                     As String
Private strCondYardFrom                         As String
Private strCondYardTo                           As String
Private strCondUserCode                         As String
Private strCondMaintenanceRequestDateFrom       As String
Private strCondMaintenanceRequestDateTo         As String
Private strCondMaintenanceEndDateFrom           As String
Private strCondMaintenanceEndDateTo             As String
Private strCondContractState                    As String
Private strCondSortOrder                        As String
Private blnMaintenanceReqNull                   As String
Private blnMaintenanceEndNull                   As String
Private strCondUserKana                         As String   '2006.2.3 SHIBAZAKI 追加
Private strCondCon                              As String   '2012/07/07 M.HONDA
Private strCondKaitanto                         As String   'ADD 2019/04/10 Y_SUZUKI EGL

'画面ロード中（サブフォームのLOADイベントでTrueになり、親フォームのLoad終了時にOFFになる）
Public blnLoading   As Boolean

'True=終了ボタンでの問い合わせを行わない
Private blnSkipFormUnload As Boolean

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_PROG_ID             As String = "FVS720"
Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

'値が変更された際にタイトル表示する
Private Const P_EDIT_MODE           As String = "(修正中)"

'契約状態
Private Const P_STATUS_UNSETTING    As String = "00"    '契約状態「未設定」
Private Const P_STATUS_CONTRACT     As String = "01"    '契約状態「契約」
Private Const P_STATUS_CANCEL       As String = "02"    '契約状態「解約予定」
Private Const P_STATUS_EXTEND       As String = "03"    '契約状態「解約延長」
Private Const P_STATUS_MAINTENANCE  As String = "04"    '契約状態「メンテ中」
Private Const P_STATUS_COMPLETE     As String = "09"    '契約状態「完了」

'予約受付状態区分
Private Const P_ACCEPT_STATUS_DECISION  As Integer = 1      '予約受付状態区分「予約確定」
Private Const P_ACCEPT_STATUS_KEEP      As Integer = 2      '予約受付状態区分「予約取置」
Private Const P_ACCEPT_STATUS_CONTRACT  As Integer = 10     '予約受付状態区分「契約受付」

'2005/12/14 SHIBAZAKI 追加↓
Private Const P_UKNO_HEAD           As String = "@" '紹介トラン作成時の契約番号の頭につける文字
'2005/12/14 SHIBAZAKI 追加↑

'2006.2.3 SHIBAZAKI 追加
'↓MOD 2019/04/10 Y_SUZUKI EGL
'Private Const P_PARAM_COUNT         As Integer = 15     'パラメータ数
Private Const P_PARAM_COUNT         As Integer = 16     'パラメータ数
'↑MOD 2019/04/10 Y_SUZUKI EGL

'==============================================================================*
'ADD 2019/04/10 Y_SUZUKI EGL
'        MODULE_NAME      :cmd_Excel_Click
'        機能             :EXCELボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_excel_Click()
    
    On Error GoTo cmd_Excel_Click_Err
    
    '出力確認
    If vbNo = MsgBox("Excel出力を実行しますか？", vbYesNo + vbQuestion, P_KAKUNIN) Then
        Exit Sub
    
    End If

    'FVS720_W01のデータを、FVS720_W03へコピー
    DataCopy2
    
    doCmd.Hourglass True
    doCmd.OutputTo acOutputTable, "FVS720_W03", acFormatXLS, , True
    
    GoTo cmd_Excel_Click_End
    
cmd_Excel_Click_Err:
    Select Case Err.Number
        Case 2501
            ' Excel出力のキャンセル時のエラー
        Case Else
            MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbExclamation + vbOKOnly, P_ERROR
    End Select

cmd_Excel_Click_End:
    doCmd.Hourglass False
    Me.cmd_Excel.SetFocus

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_取消_Click
'        機能             :取り消しボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()
On Error GoTo err_rtn
    
    Dim lngRc As Long
    
    '確認ﾒｯｾｰｼﾞ
    If blnChangeMode Then
        lngRc = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
        If lngRc = vbNo Then Exit Sub
    End If

    blnLoading = True
    
    doCmd.Hourglass True
    
    'W02をコピーしてW01
    Call DataCopy("FVS720_W02", "FVS720_W01")
    
    '再表示
    SubForm.Requery
    
    If SubForm.Form.blnStatusNotSpecifyOrComplete Then
        Call SubForm.Form.OnlyAccountDate
        SubForm.SetFocus
        SubForm.Form.txt_AccountRequestDate.SetFocus
    Else
        SubForm.SetFocus
        SubForm.Form.cmb_ContractState.SetFocus
        Call SubForm.Form.SelectStatus
    End If
    
    blnLoading = False
    
    Me.Caption = strFormCaption
    blnChangeMode = False
    
    GoTo Exit_rtn

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

Exit_rtn:
    doCmd.Hourglass False

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Click
'        機能             :終了ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next
    doCmd.Close acForm, Me.FormName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_保存_Click
'        機能             :保存ボタンクリック処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()

    'ワークファイルの内容を契約トランに書き込む
    If WorkToCarg() Then
        blnSkipFormUnload = True
        Call cmd_終了_Click
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :WorkToCarg
'        機能             :ワークファイルの内容を契約トランに書き込む
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function WorkToCarg() As Boolean

On Error GoTo err_rtn

    Dim strSQL              As String
    Dim rsSqlServer         As Recordset
    Dim rsAccess            As Recordset
    Dim lngRc               As Long
    Dim strExistingStatus   As String
    Dim strMsg              As String
    Dim objWs               As Workspace
    
    WorkToCarg = False
    strMsg = ""
    
    lngRc = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
    If lngRc = vbNo Then GoTo Exit_rtn
    
    doCmd.Hourglass True
    
    blnLoading = True
    
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    'ACCESS接続
    Set dbAccess = CurrentDb
    
    '2005/12/14 SHIBAZAKI 追加
    Set objWs = DBEngine.Workspaces(0)
    objWs.BeginTrans
    
    'ワークテーブルから変更フラグがONになっているレコード取得
    strSQL = "SELECT * FROM FVS720_W01 WHERE CHANGE_FLG = ""1"" "
    Set rsAccess = dbAccess.OpenRecordset(strSQL, dbOpenDynaset)
    
    Do While Not rsAccess.EOF
        '変更対象検索
        strSQL = ""
        strSQL = strSQL & " SELECT * FROM CARG_FILE "
        strSQL = strSQL & "  WHERE CARG_ACPTNO = '" & rsAccess.Fields("CARG_ACPTNO") & "'"
        
        Set rsSqlServer = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
        If rsSqlServer.EOF Then
            strMsg = "コンテナ契約ファイル更新情報読み込み中にエラーが発生しました"
            GoTo err_rtn
        End If
        '変更前の状態区分
        strExistingStatus = rsSqlServer.Fields("CARG_AGRE")
        
        'コンテナ契約トラン更新
        Call SetCargFile(rsAccess, rsSqlServer)
        
        '2005/12/14 SHIBAZAKI 追加↓
        '状態区分「完了」で保存する場合、紹介待ち情報を作成する。
        If strExistingStatus <> P_STATUS_COMPLETE And rsAccess.Fields("CARG_AGRE") = P_STATUS_COMPLETE Then
            If Not CmnUpdateCntaMast(dbSQLServer, _
                                     Val(rsSqlServer.Fields("CARG_YCODE")), _
                                     Val(rsSqlServer.Fields("CARG_NO")), _
                                     strMsg, _
                                     P_UKNO_HEAD & rsSqlServer.Fields("CARG_ACPTNO")) _
            Then
                GoTo err_rtn
            End If
        End If
        '2005/12/14 SHIBAZAKI 追加↑
        
        'TODO 一時機能カット
       '状態区分「完了」で保存する場合、予約自動取置き処理を行う
'        If strExistingStatus <> P_STATUS_COMPLETE And rsAccess.Fields("CARG_AGRE") = P_STATUS_COMPLETE Then
'            Call AutoKeep(rsSqlServer.Fields("CARG_YCODE"), rsSqlServer.Fields("CARG_NO"))
'        End If
        
        rsAccess.MoveNext
    Loop

    '2005/12/14 SHIBAZAKI 追加
    objWs.CommitTrans

    blnLoading = False
        
    WorkToCarg = True

    GoTo Exit_rtn
    
err_rtn:
    '2005/12/14 SHIBAZAKI 追加
    If Not objWs Is Nothing Then
        objWs.Rollback
    End If
    
    If strMsg = "" Then
        strMsg = "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    End If
    
    Err.Clear

Exit_rtn:
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    '2005/12/14 SHIBAZAKI 追加
    If Not objWs Is Nothing Then Set objWs = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, P_ERROR
    End If
    
    doCmd.Hourglass False

End Function

'==============================================================================*
'
'        MODULE_NAME      :SetCargFile
'        機能             :契約トラン更新
'        IN               :rsSource=設定先　rsDestination=設定元
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetCargFile(rsSource As Recordset, rsDestination As Recordset)

    rsDestination.Edit
    
    rsDestination.Fields("CARG_AGRE") = zeroByteToNull(rsSource.Fields("CARG_AGRE"))
    
    '契約状態が「完了」の場合、コンテナ状態を－１にする
    If rsSource.Fields("CARG_AGRE") = P_STATUS_COMPLETE Then
        rsDestination.Fields("CARG_CONTNA") = -1
    Else
        rsDestination.Fields("CARG_CONTNA") = 0
    End If
    
    rsDestination.Fields("CARG_KYDATE") = zeroByteToNull(rsSource.Fields("CARG_KYDATE"))
    rsDestination.Fields("CARG_MNT_REQDATE") = zeroByteToNull(rsSource.Fields("CARG_MNT_REQDATE"))
    rsDestination.Fields("CARG_MNT_CMPDATE") = zeroByteToNull(rsSource.Fields("CARG_MNT_CMPDATE"))
    rsDestination.Fields("CARG_MNT_TANTO") = zeroByteToNull(rsSource.Fields("CARG_MNT_TANTO"))
    rsDestination.Fields("CARG_LIQ_REQDATE") = zeroByteToNull(rsSource.Fields("CARG_LIQ_REQDATE"))
    rsDestination.Fields("CARG_UPDATE") = Now
    rsDestination.Fields("CARG_UPDAD") = Format$(DATE, "yyyymmdd")          '更新日付(var/Null=Y)
    rsDestination.Fields("CARG_UPDAJ") = Format$(time, "hhmmss")            '更新時刻(var/Null=Y)
    rsDestination.Fields("CARG_UPDPB") = Me.NAME                            '更新プログラムID(var/Null=Y)
    rsDestination.Fields("CARG_UPDUB") = LsGetUserName()                    '更新ユーザーID(var/Null=Y)
    
    rsDestination.UPDATE

End Sub

'==============================================================================*
'
'        MODULE_NAME      :ZeroByteToNull
'        機能             :０バイト文字をNULLにする
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function zeroByteToNull(varInput As Variant) As Variant

    zeroByteToNull = varInput
    
    If varInput = "" Then
        zeroByteToNull = Null
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :AutoKeep
'        機能             :予約自動取置き
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function AutoKeep(strYardCode As String, strCNTANO As String) As Boolean

On Error GoTo AutoKeep_Err

    Dim strSQL      As String
    Dim rsYouk      As Recordset
    Dim intToriDd   As Integer
    
    AutoKeep = True
    
    If Not PossibleKeep(strYardCode, strCNTANO) Then
        AutoKeep = False
        Exit Function
    End If
    
    '取置き日数をコントロールマスタから取得
    intToriDd = Nz(DLookup("CONT_TORI_DD", "dbo_CONT_MAST", "CONT_KEY=1"))

    '予約受付トランから同コンテナの「予約確定」データを取得
    strSQL = ""
    strSQL = strSQL & " SELECT TOP 1 * FROM YOUK_TRAN "
    strSQL = strSQL & "  WHERE YOUKT_YCODE = " & strYardCode
    strSQL = strSQL & "    AND YOUKT_NO = " & strCNTANO
    strSQL = strSQL & "    AND YOUKT_YUKBN = " & P_ACCEPT_STATUS_DECISION
    strSQL = strSQL & "  ORDER BY YOUKT_UKDATE, YOUKT_UKNO "
    
    Set rsYouk = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    If Not rsYouk.EOF Then
        With rsYouk
            .Edit
            '予約受付状態区分
            .Fields("YOUKT_YUKBN") = P_ACCEPT_STATUS_KEEP
            '自動区分
            .Fields("YOUKT_AUTOKBN") = P_ACCEPT_STATUS_KEEP
            '契約受付日
            .Fields("YOUKT_KUDATE") = DATE
            '取置き期限日
            .Fields("YOUKT_TKDATE") = DateAdd("d", intToriDd, DATE)
            .Fields("YOUKT_UPDAD") = Format$(DATE, "yyyymmdd")          '更新日付(var/Null=Y)
            .Fields("YOUKT_UPDAJ") = Format$(time, "hhmmss")            '更新時刻(var/Null=Y)
            .Fields("YOUKT_UPDPB") = Me.NAME                            '更新プログラムID(var/Null=Y)
            .Fields("YOUKT_UPDUB") = LsGetUserName()                    '更新ユーザーID(var/Null=Y)
            .UPDATE
        End With
    End If
    
    GoTo AutoKeep_End

AutoKeep_Err:
    MsgBox "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr _
            & Err.Description, vbCritical + vbOKOnly, P_ERROR
    
AutoKeep_End:
    If Not rsYouk Is Nothing Then rsYouk.Close: Set rsYouk = Nothing

End Function

'==============================================================================*
'
'        MODULE_NAME      :PossibleKeep
'        機能             :自動取置き可否チェック
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function PossibleKeep(strYardCode As String, strCNTANO As String) As Boolean

On Error GoTo PossibleKeep_Err

    Dim strSQL      As String
    Dim rsYouk      As Recordset
    
    PossibleKeep = True
    
    strSQL = strSQL & " SELECT * FROM YOUK_TRAN "
    strSQL = strSQL & "  WHERE YOUKT_YCODE = " & strYardCode
    strSQL = strSQL & "    AND YOUKT_NO = " & strCNTANO
    strSQL = strSQL & "    AND ( YOUKT_YUKBN = " & P_ACCEPT_STATUS_KEEP
    strSQL = strSQL & "       OR YOUKT_YUKBN = " & P_ACCEPT_STATUS_CONTRACT & " ) "

    Set rsYouk = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    If Not rsYouk.EOF Then
        PossibleKeep = False
    End If
    
    GoTo PossibleKeep_End

PossibleKeep_Err:
    MsgBox "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr _
            & Err.Description, vbCritical + vbOKOnly, P_ERROR
    
PossibleKeep_End:
    If Not rsYouk Is Nothing Then rsYouk.Close: Set rsYouk = Nothing


End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activateイベント
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
'        MODULE_NAME      :Form_Error
'        機能             :Form_Errorイベント
'        IN               :
'        OUT              :
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
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDownイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    'F1=「終了」F2=「保存」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            Me.cmd_保存.SetFocus
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        機能             :Form Loadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()
    
    Dim strOpenArg  As String
    Dim blnRet      As Boolean
    
    blnMaintenanceReqNull = False
    blnMaintenanceEndNull = False
    
    blnSkipFormUnload = False
    
    strOpenArg = Nz(Me.OpenArgs)
    Dim lngInStr    As Long
    
    'パラメータ取得
    If strOpenArg <> "" Then
        Call GetCondition(strOpenArg)
    End If
    
    blnChangeMode = False
    strFormCaption = Me.Caption
    
    'SQL-Server接続
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Exit Sub
    End If
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumonCode & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If
    '接続文字列取得
    strConnect = MSZZ007_M00(strBumonCode)
    If strConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If
    
    'ワークテーブル作成
    blnRet = MakeWork()
    
    If strCondContractState = "" Or strCondContractState = P_STATUS_COMPLETE Then
        SubForm.Form.blnStatusNotSpecifyOrComplete = True
    Else
        SubForm.Form.blnStatusNotSpecifyOrComplete = False
    End If
    
    SubForm.Requery
    
    If blnRet Then
        If SubForm.Form.blnStatusNotSpecifyOrComplete Then
            SubForm.SetFocus
            SubForm.Form.txt_AccountRequestDate.SetFocus
            Call SubForm.Form.OnlyAccountDate
        Else
            SubForm.SetFocus
'↓DEL 2019/04/10 Y_SUZUKI EGL
'該当の項目は使用不可となっている為、SetFocusでエラーが出る。
'            SubForm.Form.cmb_ContractState.SetFocus
'↑DEL 2019/04/10 Y_SUZUKI EGL
            Call SubForm.Form.SelectStatus
        End If
    Else
        Call BtnDesable
    End If
    
    'サブフォーム／親フォーム共にLOADが完了するまでカレントレコード移動処理はスキップ
    blnLoading = False

End Sub

'==============================================================================*
'
'        MODULE_NAME      :GetCondition
'        機能             :検索条件取得
'        IN               :前画面からのパラメータ
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub GetCondition(strArg As String)

    Dim strTemp     As String
    Dim lngInStr    As Long
    Dim varValue   As Variant
    Dim intCount    As Integer
    
    '前画面からカンマ区切りで条件が渡される
    'それをバラして格納する
    varValue = Split(strArg, ",")
    
    If UBound(varValue) <> P_PARAM_COUNT - 1 Then
        Exit Sub
    End If
    
    strCondCancelDateFrom = varValue(0)
    strCondCancelDateTo = varValue(1)
    strCondYardFrom = varValue(2)
    strCondYardTo = varValue(3)
    strCondUserCode = varValue(4)
    strCondMaintenanceRequestDateFrom = varValue(5)
    strCondMaintenanceRequestDateTo = varValue(6)
    strCondMaintenanceEndDateFrom = varValue(7)
    strCondMaintenanceEndDateTo = varValue(8)
    strCondContractState = varValue(9)
    strCondSortOrder = varValue(10)
    blnMaintenanceReqNull = varValue(11)
    blnMaintenanceEndNull = varValue(12)
    '2006.2.3 SHIBAZAKI 追加
    strCondUserKana = varValue(13)
    
    strCondCon = varValue(14)       '' 2012/07/07 M.HONDA
    
    strCondKaitanto = varValue(15)       'ADD 2019/04/10 Y_SUZUKI EGL
    
    Call SetSearchList
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SetSearchList
'        機能             :検索条件表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetSearchList()

    Dim strTemp     As String
    
    '検索条件文字列作成
    strTemp = "解約日　　　  : " & strCondCancelDateFrom
    If strCondCancelDateTo <> "" Then
        strTemp = strTemp & " ～ " & strCondCancelDateTo
    End If
    lst_SearchCondition.AddItem (strTemp)
    
'↓ADD 2019/04/10 Y_SUZUKI EGL
    strTemp = "解約受付担当  : " & strCondKaitanto
    lst_SearchCondition.AddItem (strTemp)
'↑ADD 2019/04/10 Y_SUZUKI EGL

    strTemp = "ヤードコード  : " & strCondYardFrom
    If strCondYardTo <> "" Then
        strTemp = strTemp & " ～ " & strCondYardTo
    End If
    lst_SearchCondition.AddItem (strTemp)
    
    '2006.2.3 SHIBAZAKI 追加
    strTemp = "顧客カナ　　  : " & strCondUserKana
    lst_SearchCondition.AddItem (strTemp)
    
    strTemp = "顧客コード　  : " & strCondUserCode
    lst_SearchCondition.AddItem (strTemp)
    
    If blnMaintenanceReqNull Then
        strTemp = "メンテ依頼日  : 未入力のみ対象"
    Else
        strTemp = "メンテ依頼日  : " & strCondMaintenanceRequestDateFrom
        If strCondMaintenanceRequestDateTo <> "" Then
            strTemp = strTemp & " ～ " & strCondMaintenanceRequestDateTo
        End If
    End If
    lst_SearchCondition.AddItem (strTemp)
    
    If blnMaintenanceEndNull Then
        strTemp = "メンテ完了日  : 未入力のみ対象"
    Else
        strTemp = "メンテ完了日  : " & strCondMaintenanceEndDateFrom
        If strCondMaintenanceEndDateTo <> "" Then
            strTemp = strTemp & " ～ " & strCondMaintenanceEndDateTo
        End If
    End If
    lst_SearchCondition.AddItem (strTemp)
    
    strTemp = "状態区分　　  : " & strCondContractState
    lst_SearchCondition.AddItem (strTemp)
    
    If strCondSortOrder = "1" Then
        strTemp = "並び順　　　  : 顧客コード"
    Else
        strTemp = "並び順　　　  : ヤードコード"
    End If
    lst_SearchCondition.AddItem (strTemp)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :MakeWork
'        機能             :ワークテーブル作成
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeWork() As Boolean
On Error GoTo err_rtn
    
    'FVS720_W01：サブフォームとリンクするテーブル
    'FVS720_W02：読み込み時の情報を保存する。取り消し時にW01にコピーして再表示する
    
    Dim strSQL          As String
    Dim rsSqlServer     As Recordset
    Dim rsAccess        As Recordset
    Dim lngRecordCount  As Long

    MakeWork = True
    
    doCmd.Hourglass True

    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    
    'ワークテーブルの元ネタ取得
    strSQL = MakeCargSql()
    Set rsSqlServer = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    'ACCESS接続
    Set dbAccess = CurrentDb
    
    'ワークテーブルが無ければ作成
    If Not TableExist("FVS720_W01") Then
        dbAccess.Execute CreateWork("FVS720_W01"), dbFailOnError
    End If
    
    'Access側ワークテーブル削除
    strSQL = MakeDeleteSql("FVS720_W01")
    dbAccess.Execute strSQL, dbFailOnError
    
    'ワークテーブルレコード追加
    Set rsAccess = dbAccess.OpenRecordset("FVS720_W01", dbOpenDynaset, dbAppendOnly)
    lngRecordCount = 0
    Do While Not rsSqlServer.EOF
        lngRecordCount = lngRecordCount + 1
        Call AddWork(rsSqlServer, rsAccess)
        rsSqlServer.MoveNext
    Loop
    
    If lngRecordCount = 0 Then
        MsgBox "該当データが存在しません。", vbInformation, P_KAKUNIN
        MakeWork = False
        GoTo Exit_rtn
    End If
    
    '担当者マスタから担当者名を取得してワークテーブル更新
    strSQL = MakeUpdateTantouNameSql()
    dbAccess.Execute strSQL, dbFailOnError
    
    'W01をコピーしてW02（取り消しボタン用）作成
    If Not DataCopy("FVS720_W01", "FVS720_W02") Then
        MakeWork = False
    End If
    
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    MakeWork = False
    Err.Clear

Exit_rtn:
    If Not rsSqlServer Is Nothing Then rsSqlServer.Close: Set rsSqlServer = Nothing
    If Not rsAccess Is Nothing Then rsAccess.Close: Set rsAccess = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    If Not dbAccess Is Nothing Then dbAccess.Close: Set dbAccess = Nothing
    
    doCmd.Hourglass False
    
    Exit Function


End Function

'==============================================================================*
'
'        MODULE_NAME      :CreateWork
'        機能             :ワークテーブルクリエイトSQL文作成
'        IN               :テーブル名
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CreateWork(strTableName As String) As String

    Dim strSQL  As String
    
    strSQL = ""
    strSQL = strSQL & " CREATE TABLE " & strTableName & " ( "
    strSQL = strSQL & "        CARG_ACPTNO            varchar(10), "
    strSQL = strSQL & "        CARG_YCODE         varchar(13), "
    strSQL = strSQL & "        YARD_NAME          varchar(36), "
    strSQL = strSQL & "        USER_KANA         varchar(18), "
    strSQL = strSQL & "        USER_NAME          varchar(36), "
    strSQL = strSQL & "        CARG_AGRE          varchar(2), "
    strSQL = strSQL & "        AGRE_NAME          varchar(20), "
    strSQL = strSQL & "        CARG_KAIUKEDATE          varchar(20), "   '2019/05/23 M.HONDA INS
    strSQL = strSQL & "        CARG_KYDATE            varchar(10), "
    strSQL = strSQL & "        CARG_MNT_REQDATE       varchar(10), "
    strSQL = strSQL & "        CARG_MNT_CMPDATE       varchar(10), "
    strSQL = strSQL & "        CARG_LIQ_REQDATE       varchar(10), "
    strSQL = strSQL & "        CARG_MNT_TANTO     varchar(3), "
    strSQL = strSQL & "        TANTM_TANTN            varchar(36), "
    strSQL = strSQL & "        CARG_LIQ_FLG           varchar(10), "
    strSQL = strSQL & "        CHANGE_FLG           char(1), "
    strSQL = strSQL & "        HIDE_CARG_MNT_REQDATE           varchar(10), "
    strSQL = strSQL & "        NEW_CONTRACT_FLAG           char(1) "
    strSQL = strSQL & " ) "

    CreateWork = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :TableExist
'        機能             :ACCESSテーブル存在チェック
'        IN               :テーブル名
'        OUT              :True=存在する False=存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function TableExist(strTableName As String) As Boolean

    Dim tdf As TableDef
    
    TableExist = False
    
    For Each tdf In dbAccess.TableDefs
        If tdf.NAME = strTableName Then
            TableExist = True
            Exit For
        End If
    Next tdf

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeCargSql
'        機能             :コンテナ契約トラン検索SQL文作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeCargSql() As String

    Dim strSQL      As String
    
    strSQL = strSQL & " SELECT CARG_ACPTNO, "
    strSQL = strSQL & "        CARG_YCODE, "
    strSQL = strSQL & "        CARG_NO, "
    strSQL = strSQL & "        YARD_NAME, "
    strSQL = strSQL & "        USER_KANA, "
    strSQL = strSQL & "        USER_NAME, "
    strSQL = strSQL & "        CARG_AGRE, "
    strSQL = strSQL & "        NAME_NAME, "
    strSQL = strSQL & "        CARG_KAIUKEDATE, "  '2019/05/23 M.HONDA INS
    strSQL = strSQL & "        CARG_KYDATE, "
    strSQL = strSQL & "        CARG_MNT_REQDATE, "
    strSQL = strSQL & "        CARG_MNT_CMPDATE, "
    strSQL = strSQL & "        CARG_LIQ_REQDATE, "
    strSQL = strSQL & "        CARG_MNT_TANTO, "
    strSQL = strSQL & "        CASE WHEN CARG_LIQ_FLG = -1 "
    strSQL = strSQL & "             THEN '精算済' "
    strSQL = strSQL & "             ELSE '' "
    strSQL = strSQL & "        END AS LIQ_FLG,"
    strSQL = strSQL & " ( SELECT COUNT(*) FROM CARG_FILE CARG_COUNT "
    strSQL = strSQL & "    WHERE CARG_COUNT.CARG_YCODE = CARG.CARG_YCODE  "
    strSQL = strSQL & "      AND CARG_COUNT.CARG_NO = CARG.CARG_NO "
    strSQL = strSQL & "      AND CARG_COUNT.CARG_ACPTNO > CARG.CARG_ACPTNO ) AS NEW_CONTRACT_CONUT "

    strSQL = strSQL & "   FROM CARG_FILE CARG  "
    strSQL = strSQL & "        INNER JOIN YARD_MAST "
    strSQL = strSQL & "           ON YARD_MAST.YARD_CODE = CARG.CARG_YCODE "
    strSQL = strSQL & "        INNER JOIN USER_MAST "
    strSQL = strSQL & "           ON USER_MAST.USER_CODE = CARG.CARG_UCODE "
    strSQL = strSQL & "        INNER JOIN NAME_MAST "
    strSQL = strSQL & "           ON NAME_MAST.NAME_CODE = CARG.CARG_AGRE "

    strSQL = strSQL & "  WHERE NAME_ID = '088' "
    strSQL = strSQL & "    AND CARG_KYDATE IS NOT NULL "
    
    '条件文作成
    strSQL = strSQL & MakeWhere()
    
    strSQL = strSQL & " ORDER BY "
    If strCondSortOrder = "1" Then
        strSQL = strSQL & " CARG_UCODE,CARG_YCODE,CARG_NO,CARG_ACPTNO "
    Else
        strSQL = strSQL & " CARG_YCODE,CARG_NO,CARG_UCODE,CARG_ACPTNO "
    End If
    
    MakeCargSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeWhere
'        機能             :条件文作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeWhere() As String

    Dim strSQL      As String
    
    '解約日
    If strCondCancelDateFrom <> "" And strCondCancelDateTo <> "" Then
        strSQL = strSQL & " AND CARG_KYDATE BETWEEN '" & strCondCancelDateFrom & "' AND '" & strCondCancelDateTo & "' "
    ElseIf strCondCancelDateFrom <> "" Then
        strSQL = strSQL & " AND CARG_KYDATE = '" & strCondCancelDateFrom & "' "
    ElseIf strCondCancelDateTo <> "" Then
        strSQL = strSQL & " AND CARG_KYDATE <= '" & strCondCancelDateTo & "' "
    End If
    
'↓ADD 2019/04/10 Y_SUZUKI EGL
    '解約受付担当
    If strCondKaitanto <> "" Then
        strSQL = strSQL & " AND CARG_KAITANTO = " & strCondKaitanto
    End If
'↑ADD 2019/04/10 Y_SUZUKI EGL
    
    'ヤードコード
    If strCondYardFrom <> "" And strCondYardTo <> "" Then
        strSQL = strSQL & " AND CARG_YCODE BETWEEN " & strCondYardFrom & " AND " & strCondYardTo
    ElseIf strCondYardFrom <> "" Then
        strSQL = strSQL & " AND CARG_YCODE = " & strCondYardFrom
    ElseIf strCondYardTo <> "" Then
        strSQL = strSQL & " AND CARG_YCODE <= " & strCondYardTo
    End If
    
    '顧客コード
    If strCondUserCode <> "" Then
        strSQL = strSQL & " AND CARG_UCODE = " & strCondUserCode
    ElseIf strCondUserKana <> "" Then   '2006.2.3 SHIBAZAKI 追加
        strSQL = strSQL & " AND USER_KANA LIKE '%" & strCondUserKana & "%' "   '2006.2.3 SHIBAZAKI 追加
    End If
    
    'メンテ依頼日
    If blnMaintenanceReqNull Then
        strSQL = strSQL & " AND CARG_MNT_REQDATE IS NULL "
    ElseIf strCondMaintenanceRequestDateFrom <> "" And strCondMaintenanceRequestDateTo <> "" Then
        strSQL = strSQL & " AND CARG_MNT_REQDATE BETWEEN '" & strCondMaintenanceRequestDateFrom & "' AND '" & strCondMaintenanceRequestDateTo & "' "
    ElseIf strCondMaintenanceRequestDateFrom <> "" Then
        strSQL = strSQL & " AND CARG_MNT_REQDATE = '" & strCondMaintenanceRequestDateFrom & "' "
    ElseIf strCondMaintenanceRequestDateTo <> "" Then
        strSQL = strSQL & " AND CARG_MNT_REQDATE <= '" & strCondMaintenanceRequestDateTo & "' "
    End If
    
    'メンテ完了日
    If blnMaintenanceEndNull Then
        strSQL = strSQL & " AND CARG_MNT_CMPDATE IS NULL "
    ElseIf strCondMaintenanceEndDateFrom <> "" And strCondMaintenanceEndDateTo <> "" Then
        strSQL = strSQL & " AND CARG_MNT_CMPDATE BETWEEN '" & strCondMaintenanceEndDateFrom & "' AND '" & strCondMaintenanceEndDateTo & "' "
    ElseIf strCondMaintenanceEndDateFrom <> "" Then
        strSQL = strSQL & " AND CARG_MNT_CMPDATE = '" & strCondMaintenanceEndDateFrom & "' "
    ElseIf strCondMaintenanceEndDateTo <> "" Then
        strSQL = strSQL & " AND CARG_MNT_CMPDATE <= '" & strCondMaintenanceEndDateTo & "' "
    End If
    
    '状態区分
    If strCondContractState <> "" Then
        strSQL = strSQL & " AND CARG_AGRE = " & strCondContractState
    End If
    
    '部屋番号
    If strCondCon <> "" Then
        strSQL = strSQL & " AND CARG_NO = " & strCondCon
    End If
    
    
    
    MakeWhere = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :AddWork
'        機能             :作業テーブル追加
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub AddWork(rsSource As Recordset, rsDestination As Recordset)

    rsDestination.AddNew
    
    rsDestination.Fields("CARG_ACPTNO") = rsSource.Fields("CARG_ACPTNO")
    rsDestination.Fields("CARG_YCODE") = Format(rsSource.Fields("CARG_YCODE"), "000000") & "-" & Format(rsSource.Fields("CARG_NO"), "000000")
    rsDestination.Fields("YARD_NAME") = rsSource.Fields("YARD_NAME")
    rsDestination.Fields("USER_KANA") = rsSource.Fields("USER_KANA")
    rsDestination.Fields("USER_NAME") = rsSource.Fields("USER_NAME")
    rsDestination.Fields("CARG_AGRE") = Format(rsSource.Fields("CARG_AGRE"), "00")
    rsDestination.Fields("AGRE_NAME") = rsSource.Fields("NAME_NAME")
    rsDestination.Fields("CARG_KAIUKEDATE") = Format(rsSource.Fields("CARG_KAIUKEDATE"), "yyyy/mm/dd") '2019/05/23 M.HONDA INS
    rsDestination.Fields("CARG_KYDATE") = Format(rsSource.Fields("CARG_KYDATE"), "yyyy/mm/dd")
    rsDestination.Fields("CARG_MNT_REQDATE") = Format(rsSource.Fields("CARG_MNT_REQDATE"), "yyyy/mm/dd")
    rsDestination.Fields("CARG_MNT_CMPDATE") = Format(rsSource.Fields("CARG_MNT_CMPDATE"), "yyyy/mm/dd")
    rsDestination.Fields("CARG_LIQ_REQDATE") = Format(rsSource.Fields("CARG_LIQ_REQDATE"), "yyyy/mm/dd")
    rsDestination.Fields("CARG_MNT_TANTO") = Format(rsSource.Fields("CARG_MNT_TANTO"), "000")
    rsDestination.Fields("TANTM_TANTN") = ""
    rsDestination.Fields("CARG_LIQ_FLG") = rsSource.Fields("LIQ_FLG")
    rsDestination.Fields("CHANGE_FLG") = "0"
    rsDestination.Fields("HIDE_CARG_MNT_REQDATE") = Format(rsSource.Fields("CARG_MNT_REQDATE"), "yyyy/mm/dd")
    
    If rsSource.Fields("NEW_CONTRACT_CONUT") > 0 Then
        rsDestination.Fields("NEW_CONTRACT_FLAG") = "1"
    Else
        rsDestination.Fields("NEW_CONTRACT_FLAG") = "0"
    End If
    
   
    rsDestination.UPDATE

End Sub

'==============================================================================*
'
'        MODULE_NAME      :MakeUpdateTantouNameSql
'        機能             :作業用テーブル担当者名更新SQL文作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeUpdateTantouNameSql() As String
    
    Dim strSQL      As String
    
    strSQL = strSQL & " UPDATE FVS720_W01 "
    strSQL = strSQL & "        INNER JOIN TANT_MAST "
    strSQL = strSQL & "           ON TANT_MAST.TANTM_TANTC = FVS720_W01.CARG_MNT_TANTO "
    strSQL = strSQL & "    SET FVS720_W01.TANTM_TANTN = TANT_MAST.TANTM_TANTN "
    strSQL = strSQL & "  WHERE TANT_MAST.TANTM_BUMOC = """ & strBumonCode & """"

    MakeUpdateTantouNameSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :DataCopy
'        機能             :作業用テーブルデータコピー
'        IN               :①コピー元テーブル名　②コピー先テーブル名
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function DataCopy(strSourceTable As String, strDestinationTable As String, Optional blnClearChange As Boolean = False) As Boolean
On Error GoTo err_rtn
    
    Dim strSQL      As String
    Dim blnConnect  As Boolean
    
    DataCopy = True
    
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
        blnConnect = True
    Else
        blnConnect = False
    End If
    
    'コピー先ワークテーブルが無ければ作成
    If Not TableExist(strDestinationTable) Then
        dbAccess.Execute CreateWork(strDestinationTable), dbFailOnError
    End If
    
    'コピー先ワークテーブル削除
    strSQL = MakeDeleteSql(strDestinationTable)
    dbAccess.Execute strSQL, dbFailOnError
    
    'データコピー
    If blnClearChange Then
        strSQL = ""
        strSQL = strSQL & " UPDATE " & strSourceTable
        strSQL = strSQL & "    SET CHANGE_FLG = '0' "
        dbAccess.Execute strSQL, dbFailOnError
    End If
    
    strSQL = MakeDataCopySql(strSourceTable, strDestinationTable)
    dbAccess.Execute strSQL, dbFailOnError
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    DataCopy = False
    Err.Clear

Exit_rtn:
    If blnConnect Then
        dbAccess.Close
        Set dbAccess = Nothing
    End If
    
    Exit Function

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeDataCopySql
'        機能             :作業用テーブルデータコピーSQL文作成
'        IN               :①コピー元テーブル名　②コピー先テーブル名
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeDataCopySql(strSourceTable As String, strDestinationTable As String) As String

    Dim strSQL      As String
    
    strSQL = strSQL & " INSERT INTO " & strDestinationTable
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & " FROM " & strSourceTable
    
    MakeDataCopySql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :MakeDeleteSql
'        機能             :作業用テーブル削除SQL文作成
'        IN               :対象テーブル名
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeDeleteSql(strTableName As String) As String

    Dim strSQL      As String
    
    strSQL = strSQL & " DELETE FROM " & strTableName
    
    MakeDeleteSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload
'        機能             :Form_Unloadイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Dim strMsg  As String
    Dim lngRc   As Long
    
    '契約番号指定モードで保存が押下された場合、終了時の問い合わせは行わない
    If blnSkipFormUnload Then
        lngRc = vbYes
    Else
        '終了メッセージを表示
        strMsg = "終了します。よろしいですか？"
        lngRc = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, P_KAKUNIN)
    End If
    
    If lngRc = vbNo Then
       Cancel = True
    End If
    
    Dim objForm     As Form
    
    Set objForm = Forms("FVS710")
    If Not objForm Is Nothing Then
        Call objForm.subReload
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :InputChange
'        機能             :サブフォームで入力変更時の処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub InputChange()

    '変更された場合（初回だけ）タイトルにその旨を表示する
    If Not blnChangeMode Then
        blnChangeMode = True
        Me.Caption = Me.Caption & P_EDIT_MODE
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :EndChange
'        機能             :サブフォームで保存完了時の処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub EndChange()

    Me.Caption = strFormCaption
    blnChangeMode = False

End Sub

'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :subReload
'        CREATE_DATE      :2005/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subReload(intFlg As Integer)

    Dim blnRet      As Boolean
    blnRet = True
    
    If intFlg = 1 Then
        If SubForm.Form.Dirty Then    'カレントレコードが変更されていたら保存する
            SubForm.Form.Dirty = False
        End If
        ' 解約入力画面で修正があった場合、画面を再表示する
        blnRet = MakeWork
        SubForm.Requery
        
    End If

    If blnRet Then
        Call SubForm.Form.subFormReload
    Else
        cmd_終了.Enabled = True
    End If

    If Not blnRet Then
        Call BtnDesable
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :終了ボタン以外を使用不可にする
'        MODULE_ID        :BtnDesable
'        CREATE_DATE      :2005/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub BtnDesable()

    cmd_保存.Enabled = False
    cmd_取消.Enabled = False
    
End Sub

'==============================================================================*
'ADD 2019/04/10 Y_SUZUKI EGL
'        MODULE_NAME      :DataCopy2
'        機能             :エクセルに出力するデータをFVS720_W01からFVS720_W03へコピーする。
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function DataCopy2() As Boolean
On Error GoTo err_rtn
    
    Dim strSQL      As String
    Dim blnConnect  As Boolean
    
    DataCopy2 = True
    
    If dbAccess Is Nothing Then
        Set dbAccess = CurrentDb
        blnConnect = True
    Else
        blnConnect = False
    End If
    
    'コピー先ワークテーブルが無ければ作成
    If Not TableExist("FVS720_W03") Then
        'CREATE文作成
        strSQL = "CREATE TABLE FVS720_W03 (" _
               & "  契約番号        varchar(10)," _
               & "  ヤード-コンテナ varchar(13)," _
               & "  ヤード名        varchar(36)," _
               & "  顧客カナ        varchar(18)," _
               & "  顧客名          varchar(36)," _
               & "  状態            varchar(20)," _
               & "  解約受付日      varchar(20)," _
               & "  解約日          varchar(10)," _
               & "  解約受付担当    varchar(36)," _
               & "  メンテ依頼日    varchar(10)," _
               & "  メンテ完了日    varchar(10)," _
               & "  メンテ担当      varchar(36)," _
               & "  精算依頼日      varchar(10)" _
               & ")"
        'ワークテーブル作成
        dbAccess.Execute CreateWork(strSQL), dbFailOnError
    End If
    
    'コピー先ワークテーブル削除
    strSQL = MakeDeleteSql("FVS720_W03")
    dbAccess.Execute strSQL, dbFailOnError
    
    'データコピー
    strSQL = "INSERT INTO FVS720_W03" & vbCrLf _
           & " SELECT W01.CARG_ACPTNO AS 契約番号" & vbCrLf _
           & "      , W01.CARG_YCODE AS ヤードコンテナ" & vbCrLf _
           & "      , W01.YARD_NAME AS ヤード名" & vbCrLf _
           & "      , W01.USER_KANA AS 顧客カナ" & vbCrLf _
           & "      , W01.USER_NAME AS 顧客名" & vbCrLf _
           & "      , W01.AGRE_NAME AS 状態" & vbCrLf _
           & "      , W01.CARG_KAIUKEDATE AS 解約受付日" & vbCrLf _
           & "      , W01.CARG_KYDATE AS 解約日" & vbCrLf _
           & "      , TBL.TANTM_KAITANTN AS 解約受付担当" & vbCrLf _
           & "      , W01.CARG_MNT_REQDATE AS メンテ依頼日" & vbCrLf _
           & "      , W01.CARG_MNT_CMPDATE AS メンテ完了日" & vbCrLf _
           & "      , W01.TANTM_TANTN AS メンテ担当" & vbCrLf _
           & "      , W01.CARG_LIQ_REQDATE AS 精算依頼日" & vbCrLf _
           & "   FROM FVS720_W01 W01 INNER JOIN (SELECT CF.CARG_ACPTNO" & vbCrLf _
           & "                                        , TM.TANTM_TANTN AS TANTM_KAITANTN" & vbCrLf _
           & "                                     FROM dbo_CARG_FILE CF" & vbCrLf _
           & "                                          LEFT JOIN dbo_TANT_MAST TM" & vbCrLf _
           & "                                                  ON TM.TANTM_TANTC = CF.CARG_KAITANTO" & vbCrLf _
           & "                                    WHERE TM.TANTM_BUMOC = """ & strBumonCode & """ ) TBL" & vbCrLf _
           & "                               ON TBL.CARG_ACPTNO = W01.CARG_ACPTNO"
    dbAccess.Execute strSQL, dbFailOnError
    GoTo Exit_rtn
    
err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    DataCopy2 = False
    Err.Clear

Exit_rtn:
    If blnConnect Then
        dbAccess.Close
        Set dbAccess = Nothing
    End If
    
    Exit Function

End Function

