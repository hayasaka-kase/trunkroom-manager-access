VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS441"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : ヤードページ自動生成＆ＨＰ反映機能
'
'        PROGRAM_NAME    : ヤードページ自動生成＆ＨＰ反映機能
'        PROGRAM_ID      : FKS441
'        PROGRAM_KBN     : FORM
'
'        CREATE          : 2008/02/02
'        CERATER         : EGL tajima
'        Ver             : 0.0
'
'        UPDATE          : 2008/10/06
'        UPDATER         : EGL hirano
'        Ver             : 0.1
'
'        UPDATE          : 2008/10/30
'        UPDATER         : EGL hirano
'        Ver             : 0.2
'
'        UPDATE          : 2011/10/14
'        UPDATER         : M.RYU
'        Ver             : 0.3
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   定数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const P_ページ生成_MODE       As Integer = 1
Private Const P_FTPアップ_MODE        As Integer = 2
Private Const P_コンテナ部門          As String = "H"
Private Const P_トランク部門          As String = "8"

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private g_CONT_BUMOC As String

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    'Debug.Print Nz(Me.OpenArgs)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :Form_Load
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    Me.cmd_OK.Visible = False
    Me.cmd_実行.Visible = True
    Me.cmd_Cancel.Visible = True
    
    If isAuthority() = True Then
        If Me.OpenArgs = P_ページ生成_MODE Then
            Me.lbl_メッセージ.Caption = "現情報を元にヤード詳細ページを生成します。よろしいですか？"
        Else
            Me.lbl_メッセージ.Caption = "生成済みのヤード詳細ページをＷｅｂに反映します。よろしいですか？"
        End If
    Else
        Me.lbl_メッセージ.Caption = "あなたにはヤード詳細ページ生成またはＷｅｂ反映の権限がありません。"
        ' ボタントランスフォーム
        Me.cmd_OK.Visible = True
        Call Me.cmd_OK.SetFocus
        Me.cmd_実行.Visible = False
        Me.cmd_Cancel.Visible = False
    End If


End Sub

'==============================================================================*
'
'        MODULE_NAME      :実行ボタン押下
'        MODULE_ID        :cmd_実行_Click
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_実行_Click()
    
    Dim isException As Boolean
    isException = False
    
On Error GoTo Exception

    doCmd.Hourglass True
            
    '部門コードの取得
    g_CONT_BUMOC = fncGetBumonCode
        
    'モードによって主処理分岐
    If Me.OpenArgs = P_ページ生成_MODE Then
        Call mainGeneratedPage  'ページ生成
    Else
        Call mainFtpUpload      'FTPアップ（反映処理）
    End If
    
    GoTo Finally

Exception:
    isException = True

Finally:
    doCmd.Hourglass False
    ' ボタントランスフォーム
    Me.cmd_OK.Visible = True
    Call Me.cmd_OK.SetFocus
    Me.cmd_実行.Visible = False
    Me.cmd_Cancel.Visible = False

    If isException Then Call MSZZ024_M00("cmd_実行_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :キャンセルボタン押下
'        MODULE_ID        :Cmd_Cansel_Click
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Cancel_Click()

    ' 画面を閉じる
    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ＯＫボタン押下
'        MODULE_ID        :cmd_OK_Click
'        CREATE_DATE      :22008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_OK_Click()
    ' 画面を閉じる
    On Error Resume Next
    doCmd.Close acForm, Me.FormName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ページ生成主処理
'        MODULE_ID        :mainGeneratedPage
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub mainGeneratedPage()

    Dim isException     As Boolean
    Dim stYardMast      As GenYardPage.Type_YARD_WP_INF
    Dim stNyarMast(10)  As GenYardPage.Type_NYAR_WP_INF
    Dim stTypeData()    As GenYardPage.Type_TYPE_WP_INF 'コンテナのみ取得
    Dim arTrnkPrices(4) As String   'トランクのみ取得
    
    isException = False
    
On Error GoTo Exception

    ' タイプ情報は動的配列としておく
    ReDim Preserve stTypeData(GenYardPage.TYPE_配列標準_SIZE)
    'ヤード関連情報の読込
    Me.lbl_メッセージ.Caption = "ヤード情報読込中"
    
    stYardMast.code = Forms!FKS440_ヤードマスタ保守!cmb_ヤードコード
    If readData(stYardMast, stNyarMast, stTypeData, arTrnkPrices) = False Then GoTo Finally
    
''''------------DELETE 2011/10/14 M.RYU---------<S>
''''    'テンプレートを元にヤード詳細ページ生成
''''    Me.lbl_メッセージ.Caption = "ヤード詳細ページ生成中"
''''    If GenYardPage.GenerateYardPage(g_CONT_BUMOC, stYardMast, stNyarMast, stTypeData, arTrnkPrices) = False Then
''''        Me.lbl_メッセージ.Caption = "ヤード詳細ページ生成は失敗しました。"
''''        GoTo Finally
''''    End If
''''
''''    Me.lbl_メッセージ.Caption = "画像ファイル準備中"
''''    If GenYardPage.PreparationImageFile(g_CONT_BUMOC, stYardMast.CODE) = False Then
''''        Me.lbl_メッセージ.Caption = "画像ファイルの準備が失敗しました。"
''''        GoTo Finally
''''    End If
''''------------DELETE 2011/10/14 M.RYU---------<E>

    ' 配置図生成 & 寸法表生成（トランクのみ）
    Me.lbl_メッセージ.Caption = "Ｅｘｃｅｌ配置図寸法表読込ＰＤＦ変換中"
    If GenYardPage.GeneratePDF(g_CONT_BUMOC, stYardMast.code) = False Then
        Me.lbl_メッセージ.Caption = "このヤード用の配置図＆寸法表生成のＥｘｃｅｌファイルがありません。" & vbCrLf & _
                                    "確認後、再度ページ生成を行って下さい。"
        GoTo Finally
    End If
    Me.lbl_メッセージ.Caption = "ヤード詳細ページを生成しました"
    
    GoTo Finally

Exception:
    isException = True

Finally:
    Erase stTypeData ' 動的配列の解放

    If isException = True Then
        Me.lbl_メッセージ.Caption = Me.lbl_メッセージ.Caption & "にてエラーが発生しました。"
        Call Err.Raise(Err.Number, "mainGeneratedPage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
                   
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ＦＴＰへのUP
'        MODULE_ID        :mainFtpUpload
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub mainFtpUpload()

On Error GoTo Exception

    Me.lbl_メッセージ.Caption = "各ファイルをＷｅｂに反映中"
    If GenYardPage.UploadWebPage(g_CONT_BUMOC, _
                                 Forms!FKS440_ヤードマスタ保守!cmb_ヤードコード) = False Then
        Me.lbl_メッセージ.Caption = "Ｗｅｂへ反映するファイルが足りません。" & vbCrLf & _
                                    "ページ生成を行ってから再度更新を行って下さい。"
        Exit Sub
    End If
    
    Me.lbl_メッセージ.Caption = "Ｗｅｂへの反映日を記録中"
    
    'FTP更新時を刻む
    Call setYardWpUpdate
    
    Me.lbl_メッセージ.Caption = "Ｗｅｂへの反映が完了しました"

    Exit Sub

Exception:
    Me.lbl_メッセージ.Caption = Me.lbl_メッセージ.Caption & "にてエラーが発生しました。"
    Call Err.Raise(Err.Number, "mainGeneratedPage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :ＵＰＬＯＡＤ日付設定
'        MODULE_ID        :setYardWpUpdate
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub setYardWpUpdate()

    Dim adoDbConnection As Object
    Dim isException     As Boolean
    Dim SETDATE         As String
    
    isException = False

On Error GoTo Exception
    
    'DB接続
    Call subConnectServer(adoDbConnection)
    
    SETDATE = GenYardPage.SetUploadInfo(adoDbConnection, Forms!FKS440_ヤードマスタ保守!cmb_ヤードコード)
    Forms!FKS440_ヤードマスタ保守!txt_HP更新日 = SETDATE
    GoTo Finally

Exception:
    isException = True
                   
Finally:
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "setYardWpUpdate" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub


'==============================================================================*
'
'        MODULE_NAME      :データ読込処理
'        MODULE_ID        :readData
'        CREATE_DATE      :2008/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function readData(ByRef aYardMast As GenYardPage.Type_YARD_WP_INF, _
                          ByRef aNyarMast() As GenYardPage.Type_NYAR_WP_INF, _
                          ByRef aTypeData() As GenYardPage.Type_TYPE_WP_INF, _
                          ByRef aTrnkPrices() As String)

    Dim adoDbConnection As Object
    Dim isException     As Boolean
    
    readData = False
    isException = False

On Error GoTo Exception
    
    'DB接続
    Call subConnectServer(adoDbConnection)
    
    'ヤード情報の取得・・・処理速度を稼ぐため？親画面の情報を取得
    aYardMast.NAME = Forms!FKS440_ヤードマスタ保守!txt_ヤード名
    aYardMast.ADDRESS = Forms!FKS440_ヤードマスタ保守!txt_住所1 & _
                        Forms!FKS440_ヤードマスタ保守!txt_住所2 & _
                        Forms!FKS440_ヤードマスタ保守!txt_住所3
    aYardMast.CAMPAIGN = Forms!FKS440_ヤードマスタ保守!txt_SEV1N & _
                        Forms!FKS440_ヤードマスタ保守!txt_SEV2N & _
                        Forms!FKS440_ヤードマスタ保守!txt_ENDEN
    aYardMast.NOTE = Forms!FKS440_ヤードマスタ保守!txt_備考
    aYardMast.HPNOTE = Forms!FKS440_ヤードマスタ保守!txt_HP用備考
    aYardMast.IDO = Forms!FKS440_ヤードマスタ保守!txt_YARD_IDO
    aYardMast.KEIDO = Forms!FKS440_ヤードマスタ保守!txt_YARD_KEIDO
    aYardMast.SEV1N = Forms!FKS440_ヤードマスタ保守!txt_SEV1N   '2008/10/06 ADD Rev0.1
    aYardMast.SEV2N = Forms!FKS440_ヤードマスタ保守!txt_SEV2N   '2008/10/06 ADD Rev0.1
    aYardMast.SEV3N = Forms!FKS440_ヤードマスタ保守!txt_SEV3N   '2008/10/10 ADD Rev0.1
    aYardMast.ENDEN = Forms!FKS440_ヤードマスタ保守!txt_ENDEN   '2008/10/06 ADD Rev0.1
    
    ' 参考にDB取得版
'    If GenYardPage.ReadYardMast(adoDbConnection, aYardMast) = False Then GoTo Exception
    '近隣情報の取得
    '2008/10/30 近隣情報不具合
    'Call GenYardPage.ReadNyarMast(adoDbConnection, CLng(aYardMast.CODE), aNyarMast)
    If g_CONT_BUMOC = P_コンテナ部門 Then
        Call GenYardPage.ReadNyarMast(adoDbConnection, CLng(aYardMast.code), aNyarMast)
    ElseIf g_CONT_BUMOC = P_トランク部門 Then
        Call GenYardPage.ReadNyarMast(adoDbConnection, CLng(aYardMast.code), aNyarMast, g_CONT_BUMOC)
    End If
    '2008/10/30 近隣情報不具合
    
    'タイプ情報の取得...部門によって取得内容が変わる....
    If g_CONT_BUMOC = P_コンテナ部門 Then
        Call GenYardPage.ReadTypeData(adoDbConnection, CLng(aYardMast.code), aTypeData)
    ElseIf g_CONT_BUMOC = P_トランク部門 Then
        Call GenYardPage.ReadTrnkPrices(adoDbConnection, CLng(aYardMast.code), aTrnkPrices)
    End If
    
    readData = True
    GoTo Finally

Exception:
    isException = True
                   
Finally:
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "readData" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'       MODULE_NAME     : データベース接続
'       MODULE_ID       : subConnectServer
'       CREATE_DATE     :
'       PARAM           : データベースオブジェクト
'       RETURN          :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectServer(adoDbConnection As Object)

    Dim strBUMOC        As String

On Error GoTo ErrorHandler

    '部門コード取得
    'strBumoc = fncGetBumonCode()
    strBUMOC = g_CONT_BUMOC  '既に部門コード取得済みなんで

    'ADO接続Object生成
    Set adoDbConnection = MSZZ025.ADODB_Connection(strBUMOC)

    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subConnectServer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 部門コード取得
'       MODULE_ID       : fncGetBumonCode
'       CREATE_DATE     :
'       PARAM           :
'       RETURN          : 部門コード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBumonCode() As String

    Dim strBumonCode        As String

On Error GoTo ErrorHandler

    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "部門コードの設定不正。")
    End If

    fncGetBumonCode = strBumonCode

    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetBumonCode" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'       MODULE_NAME     : 権限可否チェック
'       MODULE_ID       : isAuthority
'       CREATE_DATE     :
'       PARAM           :
'       RETURN          : true...あり、false...なし
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function isAuthority() As Boolean

    Dim adoDbConnection As Object
    Dim rsData          As Object
    Dim isException     As Boolean
    Dim strSQL          As String
    Dim userName        As String

    isAuthority = False
    isException = False
    
    On Error GoTo Exception
    
    'ユーザID取得
    userName = ""
    userName = MSZZ000.LsGetUserName
    
    'ADO接続Object生成
    Set adoDbConnection = MSZZ025.ADODB_Connection()

    ' コンテナマスタ、価格表テーブルから取得
    strSQL = "SELECT ONLET_ACCEI FROM ONLE_TABL WHERE ONLET_PROGB = 'FKS441'"
    strSQL = strSQL & "AND ONLET_USERB = '" & userName & "'"
    
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, adoDbConnection)
    ' 対象確認
    If rsData.EOF = False Then
        ' オンライン更新区分が"1"ならば権限がある
        If "1" = Nz(rsData.Fields("ONLET_ACCEI"), "") Then
            isAuthority = True
        End If
    End If
    GoTo Finally
                              
Exception:
    isException = True
                   
Finally:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not adoDbConnection Is Nothing Then adoDbConnection.Close: Set adoDbConnection = Nothing

    If isException = True Then
        Call Err.Raise(Err.Number, "isAuthority" & vbRightAllow & Err.Source, _
                       Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'****************************  ended or program ********************************
