VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FKS240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　倉庫管理システム
'**********************************************
'   プログラム名　　：　請求明細書
'   プログラムＩＤ　：
'   作　成　日　　　：
'   作　成　者　　　：　αシステム  Ishii Kaoru
'
'   修　正　日　　　：2000/08/01
'   修　正　者　　　：KASE MIURA
'   Ver             ：0.1
'   修　正　内　容　：共益費対応
'
'   修　正　日　　　：2003/05/06
'   修　正　者　　　：KASE MIURA
'   Ver             ：0.2
'   修　正　内　容　：SQLSERVER対応
'
'   修　正　日　　　：2004/10/30
'   修　正　者　　　：KASE MIURA
'   Ver             ：0.4
'   修　正　内　容　：共益費機能の停止
'
'   修　正　日　　　：2004/11/24
'   修　正　者　　　：KASE MIURA
'   Ver             ：0.5
'   修　正　内　容　：フォーム名変更
'
'   修　正　日　　　：2005/07/26
'   修　正　者　　　：KASE MIURA
'   Ver             ：0.6
'   修　正　内　容　：住所３追加
'
'   修　正　日　　　：2005/12/19
'   修　正　者　　　：KASE MIURA
'   Ver             ：0.7
'   修　正　内　容　：請求書レイアウト変更
'
'   修　正　日　　　：2006/05/10
'   修　正　者　　　：KASE SHIBAZAKI
'   Ver             ：0.8
'   修　正　内　容　：その他連絡先対応
'
'   修　正　日　　　：2008/02/04
'   修　正　者　　　：N.MIURA
'   Ver             ：0.9
'   修　正　内　容　：顧客振込先専用口座印字対応（仮想口座）
''
'   修　正　日　　　：2008/06/16
'   修　正　者　　　：N.Miura
'   Ver             ：0.A
'   修　正　内　容　：社印印字機能装備
'
'   修　正　日　　　：2009/12/02
'   修　正　者　　　：KASE M.Ryu
'   Ver             ：0.B
'   修　正　内　容　：請求書レイアウト変更
'
'   修　正　日　　　：2010/03/25
'   修　正　者　　　：K.ISHIZAKA
'   Ver             ：1.0
'   修　正　内　容　：請求書発行区分＝1:発行のものだけを対象とする
'
'   修　正　日　　　：2010/04/10
'   修　正　者　　　：K.ISHIZAKA
'   Ver             ：1.1
'   修　正　内　容　：請求書発行区分は契約ファイルの項目をみる
'
'   修　正　日　　　：2011/01/12
'   修　正　者　　　：M.RYU
'   Ver             ：1.2
'   修　正　内　容　：画面項目追加（発行日、振込期限）
'
'   修　正　日　　　：2013/06/03
'   修　正　者　　　：M.HONDA
'   Ver             ：1.3
'   修　正　内　容　：終了メッセージを表示させないように修正
'
'   修　正　日　　　：2014/09/30
'   修　正　者　　　：M.HONDA
'   Ver             ：1.4
'   修　正　内　容　：デフォルトで計上年月を表示するように修正
'
'   修　正　日　　　：2015/10/31
'   修　正　者　　　：M.HONDA
'   Ver             ：1.5
'   修　正　内　容　：前払い月数を見るように対応
'
'   修　正　日　　　：2017/11/27
'   修　正　者　　　：k.sato
'   Ver             ：1.6
'   修　正　内　容　：未取込警告初期表示、カード取込日を表示
'
'   修　正　日　　　：2017/12/01
'   修　正　者　　　：N.IMAI
'   Ver             ：1.7
'   修　正　内　容　：請求書出力した顧客一覧をワークに保存する
'
'   修　正　日　　　：2018/07/13
'   修　正　者　　　：N.IMAI
'   Ver             ：1.8
'   修　正　内　容　：ワークへの保存時に「'」が存在するとエラーになる問題に対応
'
'   修　正　日　　　：2018/08/04
'   修　正　者　　　：N.IMAI
'   Ver             ：1.9
'   修　正　内　容　：既存顧客新保証プラン対応
'
'   修　正　日　　　：2018/10/15
'   修　正　者　　　：N.IMAI
'   Ver             ：2.0
'   修　正　内　容　：仮想口座取得時の条件から会社コードを削除
'
'   修　正　日　　　：2018/12/15
'   修　正　者　　　：N.IMAI
'   Ver             ：2.1
'   修　正　内　容　：カード未取込警告が常に表示される問題に対応
'                     帳票の送付先名称について出力条件を変更
'
'   修　正　日　　　：2019/01/12
'   修　正　者　　　：TFC
'   Ver             ：2.2
'   修　正　内　容　：顧客CSV・帳票CSV出力、WebへのUpload追加
'
'   修　正　日　　　：2019/04/26
'   修　正　者　　　：N.IMAI
'   Ver             ：2.4
'   修　正　内　容　：新保証会社切替え顧客の場合、初月でも請求書を出力する
'
'   修　正　日　　　：2019/08/15
'   修　正　者　　　：Y.WADA
'   Ver             ：2.5
'   修　正　内　容　：消費税対応
'
'   修　正　日　　　：2019/10/19
'   修　正　者　　　：N.IMAI
'   Ver             ：2.6
'   修　正　内　容　：消費税対応（Web請求書）
'
'   修　正　日　　　：2020/03/24
'   修　正　者　　　：M.HONDA
'   Ver             ：2.7
'   修　正　内　容　：強制発行フラグを見るように修正
'
'   修　正　日　　　：2020/04/06
'   修　正　者　　　：N.IMAI
'   Ver             ：2.8
'   修　正　内　容　：ワークへの保存時に「'」が存在するとエラーになる問題に対応
'
'   修　正　日　　　：2022/05/31
'   修　正　者　　　：K.KINEBUCHI
'   Ver             ：3.0
'   修　正　内　容　：22.6.1より分社化による角印取得方法変更　※2.9のコメントがない
'
'   修　正　日　　　：2022/08/04
'   修　正　者　　　：N.IMAI
'   Ver             ：3.2
'   修　正　内　容　：圧着ボタン追加
'
'   修　正　日　　　：2022/08/10
'   修　正　者　　　：N.IMAI
'   Ver             ：3.3
'   修　正　内　容　：請求書の口座名対応
'
'   修　正　日　　　：2023/02/24
'   修　正　者　　　：N.IMAI
'   Ver             ：3.4
'   修　正　内　容　：複数のKOUZ_MASTを持つ顧客に対応
''
''**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "FKS240"

Private Const strMsg_kaku As String = "確認"
Private Const strMsg_err As String = "エラー"

Private Const strMsg_Carderr As String = "クレジットカードが取り込まれていませんが、請求書の発行を行ってよろしいでしょうか"  '2017-11-27 add k.sato

Private m_strTimeStp    As String                           '実行日時
Private m_strZTimeStp   As String                           '１日前の実行日時

Private FstrBumonCode   As String                                               'INSERT 2023/02/24 N.IMAI

'---> Rev 1.8 START
'Sleep使用のための宣言(2010以降は64ビット記述でOK)
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If
'---> Rev 1.8 END

'==============================================================================*
'
'        MODULE_NAME      :cmd_DataMake_Clickイベント
'        MODULE_ID        :cmd_DataMake_Click
'        CREATE_DATE      :2018.06.30   TFC
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_DataMake_Click()

    Dim intAns      As Integer              'MsgBox 復帰値
    Dim dt_Now      As Date                 '現在日時
    Dim strWhere    As String               'Where構文
    Dim strBUMOC    As String               '部門コード
    Dim strFileCT   As String               '顧客CSVファイル名
    Dim strFilePT   As String               '帳票CSVファイル名
    Dim isubRet     As Integer              '呼び出し関数復帰値
    Dim strMsg      As String               'メッセージBox用メッセージ
    Dim strPath     As String               'CSV格納パス
    Dim strSQL      As String               'SQLスクリプト格納

    On Error GoTo Err_cmd_DataMake_Click

    strMsg = "請求書データを作成します。" '& vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_DataMake_Click
    End If

    '*** ワーク作成 ***
    doCmd.Hourglass True

    Call DEL_RTN
    
    '===============================
    ' 画面指定計上年月の前月分作成
    '===============================
    If Make_Work(1, 1) = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_DataMake_Click
    End If
    
    '前月分でデータ取得できなくもエラーとしない
    '*** Jetワーク作成 ***
    If Make_JetWork = False Then
        'エラー発生またはデータなし
'        GoTo Exit_cmd_DataMake_Click
    End If
    
    '--- 請求書ヘッダーWKの複製
    'システムメッセージ抑制
    doCmd.SetWarnings WarningsOn:=False
    
    If ExistTable("前月顧客ＷＫ") = True Then
        '存在する場合、削除
        strSQL = "drop table 前月顧客ＷＫ;"
        doCmd.RunSQL strSQL
    End If

    strSQL = "select * into 前月顧客ＷＫ from 請求書ヘッダーＷＫ;"
    doCmd.RunSQL strSQL
    'システムメッセージ抑制解除
    doCmd.SetWarnings WarningsOn:=True
    
    Call DEL_RTN
    '===============================
    ' 画面指定計上年月分作成
    '===============================
    If Make_Work(0, 1) = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_DataMake_Click
    End If

    '*** Jetワーク作成 ***
    If Make_JetWork = False Then
        'エラー発生またはデータなし
        GoTo Exit_cmd_DataMake_Click
    End If


    '現在日付
    dt_Now = Now()
    
    '部門コード取得
    strBUMOC = Get_Bumoc()
    
    'PATH取得
    strPath = Get_CsvPath(strBUMOC)
    
    If strPath = "" Then
        strMsg = "INTI_FILEの設定不足です 識別子[" & GetProgId() & "]"
        Call MsgBox(strMsg, vbOKOnly, "データ作成")
        GoTo Exit_cmd_DataMake_Click
    End If
    
    '各CSVファイル名作成
    strFileCT = strPath & "顧客_" & strBUMOC & "_" & _
                    Left(Format(Me.txt_日付FROM & "/01", "yyyymmdd"), 6) & "_" & Format(dt_Now, "YYYYMMDDHHMM") & ".csv"
    strFilePT = strPath & "帳票_" & strBUMOC & "_" & _
                    Left(Format(Me.txt_日付FROM & "/01", "yyyymmdd"), 6) & "_" & Format(dt_Now, "YYYYMMDDHHMM") & ".csv"

    'CSV出力
    isubRet = FncOut_Csv(strFileCT, strFilePT, Me.txt_日付FROM)
    
    doCmd.Hourglass False
    
    If isubRet <> 0 Then
        strMsg = "CSV出力に失敗しました。" '& vbCr
        strMsg = strMsg & "エラーログを確認してください"
        Call MsgBox(strMsg, vbOKOnly, "データ作成")
    Else
        strMsg = "【顧客CSV】" & vbCr & _
                    strFileCT & vbCr & _
                    "【帳票CSV】" & vbCr & _
                    strFilePT & vbCr & _
                    "を作成しました"
        Call MsgBox(strMsg, vbOKOnly, "データ作成")
    End If


Exit_cmd_DataMake_Click:
    doCmd.Hourglass False
    Exit Sub

Err_cmd_DataMake_Click:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmd_DataMake_Click

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_DataMake2_Clickイベント
'        MODULE_ID        :cmd_DataMake2_Click
'        CREATE_DATE      :
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_DataMake2_Click()
    Dim intAns      As Integer              'MsgBox 復帰値
    Dim dt_Now      As Date                 '現在日時
    Dim strWhere    As String               'Where構文
    Dim strBUMOC    As String               '部門コード
    Dim strFileCT   As String               '顧客CSVファイル名
    Dim strFilePT   As String               '帳票CSVファイル名
    Dim isubRet     As Integer              '呼び出し関数復帰値
    Dim strMsg      As String               'メッセージBox用メッセージ
    Dim strPath     As String               'CSV格納パス
    Dim strSQL      As String               'SQLスクリプト格納

    On Error GoTo Err_cmd_DataMake_Click

    strMsg = "圧着葉書用請求書データを作成します。" '& vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_DataMake_Click
    End If

    '*** ワーク作成 ***
    doCmd.Hourglass True

    Call DEL_RTN
    
    '===============================
    ' 画面指定計上年月の前月分作成
    '===============================
    If Make_Work(1, 0) = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_DataMake_Click
    End If
    
    '前月分でデータ取得できなくもエラーとしない
    '*** Jetワーク作成 ***
    If Make_JetWork = False Then
        'エラー発生またはデータなし
'        GoTo Exit_cmd_DataMake_Click
    End If
    
    '--- 請求書ヘッダーWKの複製
    'システムメッセージ抑制
    doCmd.SetWarnings WarningsOn:=False
    
    If ExistTable("前月顧客ＷＫ") = True Then
        '存在する場合、削除
        strSQL = "drop table 前月顧客ＷＫ;"
        doCmd.RunSQL strSQL
    End If

    strSQL = "select * into 前月顧客ＷＫ from 請求書ヘッダーＷＫ;"
    doCmd.RunSQL strSQL
    'システムメッセージ抑制解除
    doCmd.SetWarnings WarningsOn:=True
    
    Call DEL_RTN
    '===============================
    ' 画面指定計上年月分作成
    '===============================
    If Make_Work(0, 0) = False Then
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_DataMake_Click
    End If

    '*** Jetワーク作成 ***
    If Make_JetWork = False Then
        'エラー発生またはデータなし
        GoTo Exit_cmd_DataMake_Click
    End If


    '現在日付
    dt_Now = Now()
    
    '部門コード取得
    strBUMOC = Get_Bumoc()
    
    'PATH取得
    strPath = Get_CsvPath(strBUMOC)
    
    If strPath = "" Then
        strMsg = "INTI_FILEの設定不足です 識別子[" & GetProgId() & "]"
        Call MsgBox(strMsg, vbOKOnly, "データ作成")
        GoTo Exit_cmd_DataMake_Click
    End If
    
    '各CSVファイル名作成
    strFileCT = strPath & "顧客_圧着_" & strBUMOC & "_" & _
                    Left(Format(Me.txt_日付FROM & "/01", "yyyymmdd"), 6) & "_" & Format(dt_Now, "YYYYMMDDHHMM") & ".csv"
    strFilePT = strPath & "帳票_圧着_" & strBUMOC & "_" & _
                    Left(Format(Me.txt_日付FROM & "/01", "yyyymmdd"), 6) & "_" & Format(dt_Now, "YYYYMMDDHHMM") & ".csv"

    'CSV出力
    isubRet = FncOut_Csv2(strFileCT, strFilePT, Me.txt_日付FROM)
    
    doCmd.Hourglass False
    
    If isubRet <> 0 Then
        strMsg = "CSV出力に失敗しました。" '& vbCr
        strMsg = strMsg & "エラーログを確認してください"
        Call MsgBox(strMsg, vbOKOnly, "データ作成")
    Else
        strMsg = "【顧客CSV】" & vbCr & _
                    strFileCT & vbCr & _
                    "【帳票CSV】" & vbCr & _
                    strFilePT & vbCr & _
                    "を作成しました"
        Call MsgBox(strMsg, vbOKOnly, "データ作成")
    End If


Exit_cmd_DataMake_Click:
    doCmd.Hourglass False
    Exit Sub

Err_cmd_DataMake_Click:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmd_DataMake_Click


End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_Upload_Clickイベント
'        MODULE_ID        :cmd_Upload_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Upload_Click()

    Dim iCaptureID(2)   As Integer          '取込ID
                                            '(0):顧客CSV 取込ID
                                            '(1):帳票CSV 取込ID
    Dim bRet            As Boolean          '呼び出し関数復帰値
    Dim strMsg(2)       As String           'メッセージBoxメッセージ
    Dim intAns          As Integer          'メッセージBox返却値
    Dim strFileCT       As String           '顧客CSVファイル名
    Dim strFilePT       As String           '帳票CSVファイル名
    Dim strPath         As String           'CSV出力先Path
    Dim strBUMOC        As String           '部門コード
    Dim strFilePrefix   As String           'CSVPrefix部
    Dim intRet(4)       As Integer          '関数復帰値([0]:顧客,[1]:帳票,[2]:顧客状況確認,[3]:帳票状況確認)
    Dim ilp1            As Integer          'Loopカウンタ
    
    '---------------------------------------
    ' URLをセット
    '---------------------------------------
'    Call Set_Account
'    Call Set_APItoken
'    Call Set_Domain
    '---------------------------------------
    ' 最新更新日時の各CSVを指定
    '---------------------------------------
    '部門コード取得
    strBUMOC = Get_Bumoc()
    'CSV出力Path取得
    strPath = Get_CsvPath(strBUMOC)
    'CSVPreFix部
    strFilePrefix = "顧客_" & strBUMOC
    '最新顧客CSVファイル名取得
    strFileCT = FncGetNewFile(strPath, strFilePrefix)
    
    If strFileCT = "" Then
        strMsg(0) = "顧客CSVファイルを取得できませんでした"
        Call MsgBox(strMsg(0), vbOKOnly, "アップロード")
        GoTo Exit_cmd_Upload_Click
    End If
    
    strFileCT = strPath & strFileCT
    
    'CSVPreFix部
    strFilePrefix = "帳票_" & strBUMOC
    '最新顧客CSVファイル名取得
    strFilePT = FncGetNewFile(strPath, strFilePrefix)
    strFilePT = strPath & strFilePT
    If strFilePT = "" Then
        strMsg(0) = "帳票CSVファイルを取得できませんでした"
        Call MsgBox(strMsg(0), vbOKOnly, "アップロード")
        GoTo Exit_cmd_Upload_Click
    End If
    
    strMsg(0) = "【顧客CSV】" & vbCr & _
                strFileCT & vbCr & _
                "【帳票CSV】" & vbCr & _
                strFilePT & vbCr & _
                "をアップロードします。よろしいですか？"
    intAns = MsgBox(strMsg(0), vbQuestion + vbYesNo, "アップロード")
    
    If intAns = vbNo Then
        GoTo Exit_cmd_Upload_Click
    End If
        
    doCmd.Hourglass True
    '---------------------------------------
    ' ファイルUpLoad
    '---------------------------------------
    '顧客CSV Upload
    intRet(0) = Fnc_BulkCapture(1, strFileCT, iCaptureID(0))
    Debug.Print Now() & " 顧客CSV Upload Ret:" & intRet(0)
    '帳票CSV Upload
    intRet(1) = Fnc_BulkCapture(3, strFilePT, iCaptureID(1))
    Debug.Print Now() & " 帳票CSV Upload Ret:" & intRet(1)
    
    
    If intRet(0) = -1 Then
        'アップロードエラー
        strMsg(0) = "顧客CSVアップロードに失敗しました"
        If GetLog() = 1 Then
            strMsg(0) = strMsg(0) & vbCr & _
                    "ログファイルを確認してください"
        End If
        Call MsgBox(strMsg(0), vbOKOnly, "アップロード")
        GoTo Exit_cmd_Upload_Click
    End If
    If intRet(1) = -1 Then
        strMsg(0) = "帳票CSVアップロードに失敗しました"
        If GetLog() = 1 Then
            strMsg(0) = strMsg(0) & vbCr & _
                    "ログファイルを確認してください"
        End If
        Call MsgBox(strMsg(0), vbOKOnly, "アップロード")
        GoTo Exit_cmd_Upload_Click
    End If
        
    If intRet(0) >= 2 Or _
        intRet(1) >= 2 Then
        
        '状況確認までSleep
        'レスポンス結果待ち
        For ilp1 = 0 To 4
            Sleep 1000
            DoEvents
        Next
        
        '対象データあり
        For ilp1 = 0 To 3
            '顧客CSV 取込状況確認
            If intRet(0) >= 2 Then
                intRet(2) = Fnc_BulkStatus(1, iCaptureID(0))
                Debug.Print Now() & " 顧客CSV状況確認:" & ilp1 & " Ret:" & intRet(2)
            End If
            
            '帳票CSV 取込状況確認
            If intRet(1) >= 2 Then
                intRet(3) = Fnc_BulkStatus(3, iCaptureID(1))
                Debug.Print Now() & " 帳票CSV状況確認:" & ilp1 & " Ret:" & intRet(3)
            End If
            
            If intRet(2) <= 0 And _
                intRet(3) <= 0 Then
                Exit For
            End If
            '3秒Sleep
            Sleep 3000
        Next
        If intRet(2) < 0 Then
            strMsg(0) = "顧客CSVアップロードに失敗しました"
            If GetLog() = 1 Then
                strMsg(0) = strMsg(0) & vbCr & _
                        "ログファイルを確認してください"
            End If
            Call MsgBox(strMsg(0), vbOKOnly, "アップロード")
            GoTo Exit_cmd_Upload_Click
        End If
        If intRet(3) < 0 Then
            strMsg(0) = "帳票CSVアップロードに失敗しました"
            If GetLog() = 1 Then
                strMsg(0) = strMsg(0) & vbCr & _
                        "ログファイルを確認してください"
            End If
            Call MsgBox(strMsg(0), vbOKOnly, "アップロード")
            GoTo Exit_cmd_Upload_Click
        End If
    Else
        '対象データなし
    End If
    
    doCmd.Hourglass False
    
    If intRet(0) >= 2 Then
        strMsg(0) = "顧客CSV取込ID:" & iCaptureID(0) & " データ件数(" & intRet(0) - 1 & ") "
    Else
        strMsg(0) = "顧客CSVアップロードデータ無し"
    End If
    If intRet(1) >= 2 Then
        strMsg(1) = "帳票CSV取込ID:" & iCaptureID(1) & " データ件数(" & intRet(1) - 1 & ") "
    Else
        strMsg(1) = "帳票CSVアップロードデータ無し "
    End If
                    
    Call MsgBox(strMsg(0) & vbCr & strMsg(1), vbOKOnly, "アップロード")

Exit_cmd_Upload_Click:
    doCmd.Hourglass False

End Sub

'
'==============================================================================*
'
'        MODULE_NAME      :Form_Openイベント
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

'    Resume Next
    
    Dim wkDate As Date
    Dim intAns As Integer '2017-11-29 add k.sato
     Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Title.BackColor = RGB(255, 0, 204)
    End Select
    '実行日時
    wkDate = Now()
    m_strTimeStp = Format(wkDate, "yyyy/mm/dd hh:nn:ss")
    m_strZTimeStp = Format(DateAdd("d", -1, wkDate), "yyyy/mm/dd hh:nn:ss")

    '＊＊＊画面の最大化＊＊＊
    doCmd.Maximize

    '＊＊＊画面表示クリア＊＊＊
    Me.cmb_CODEFROM = ""
    Me.cmb_CODETO = ""
    Me.txt_CODEFROM = ""
    Me.txt_CODETO = ""
    
    Me.txt_日付FROM = ""
    
    
    Me.txt_日付FROM = Format(DateAdd("d", 1, DATE), "yyyy/mm")
    
    Me.cmb_顧客FROM = ""
    Me.cmb_顧客TO = ""
    Me.Txt_顧客FROM = ""
    Me.txt_顧客TO = ""
    Me.cmb_区分FROM = ""
    Me.txt_区分FROM = ""
    Me.cmb_区分TO = ""
    Me.txt_区分TO = ""

'    Me.txt_CAMPC = DLookup("CONT_SENDEN1", "dbo_CONT_MAST", "CONT_KEY=1")       'INSERT 20080616 N.MIURA
    Me.txt_CAMPC = DLookup("CONT_CAMPC", "dbo_CONT_MAST", "CONT_KEY=1")          'INSERT 2022/05/31 K.KINEBUCHI

    Me.txt_Hkobi = Format(DATE, "yyyymmdd")         'INSERT 20110108 M.RYU
    
    'INSERT 2017-11-27 add start k.sato
    Me.txt_card_imp_date = getCardImpDate()
    
    If Format(DATE, "yyyymm") > Mid(Me.txt_card_imp_date, 1, 6) Then
        intAns = MsgBox(strMsg_Carderr, vbQuestion + vbYesNo, strMsg_kaku)
        If intAns = vbNo Then
            Cancel = True
        End If
    End If

    'INSERT 2017-11-27 add end k.sato
    
    If Nz(Me.txt_日付FROM) <> "" Then
        '振込期限日 = 計上年月の前月末日
        Me.txt_Fkbi = Format(DateAdd("d", -1, Me.txt_日付FROM & "/01"), "yyyymmdd")
    End If

    '部門コード取得
    FstrBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1"))  'INSERT 2023/02/24 N.IMAI
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'        UPDATE_DATE      : 2018.06.07 TFC
'        Rev1.8           : パラメータ追加
'                           (0:画面指定計上月,1:画面指定計上月の前月)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'画面入力より
'SQL の WHERE文作成
'**********************************************
'--- 2018.06.07 Rev1.8 UPD START
'Private Function Make_Where() As String
Private Function Make_Where(intFlg As Integer) As String
'--- 2018.06.07 Rev1.8 UPD END


    Dim strKbnfr    As String
    Dim strKbnto    As String
    Dim strCodefr   As String
    Dim strCODETO   As String
    Dim strUserfr   As String
    Dim strUserto   As String
    Dim strDayfr    As String
    Dim strSwap     As String
    Dim strWhere    As String

    On Error GoTo Err_Make_Where

    '*** 画面指定値取得 ***'
    strKbnfr = Nz(Me.cmb_区分FROM, "")
    strKbnto = Nz(Me.cmb_区分TO, "")
    strCodefr = Nz(Me.cmb_CODEFROM, "")
    strCODETO = Nz(Me.cmb_CODETO, "")
    strUserfr = Nz(Me.cmb_顧客FROM, "")
    strUserto = Nz(Me.cmb_顧客TO, "")
    
'''''    strDayfr = Nz(Me.txt_日付ＦＲＯＭ, "")             'DEL 2000/04/20
    If Nz(Me.txt_日付FROM, "") <> "" Then                   'INSERT 2011/01/12 M.RYU
'--- 2018.06.07 Rev1.8 UPD START
'        strDayfr = Nz(Me.txt_日付FROM, "") & "/01"          'ADD 2000/04/20
        If intFlg = 0 Then
            strDayfr = Nz(Me.txt_日付FROM, "") & "/01"          'ADD 2000/04/20
        Else
            strDayfr = Format(DateAdd("m", -1, Me.txt_日付FROM & "/01"), "yyyy/mm/dd")
        End If
'--- 2018.06.07 Rev1.8 UPD END
    Else
        strDayfr = ""                                       'INSERT 2011/01/12 M.RYU
    End If
    
    strSwap = ""

    '*** ＩＤの大小チェック ***
    If Trim(strCodefr) <> "" And Trim(strCODETO) <> "" Then
        '*** ＣＯＤＥ入替 ***
        If strCodefr > strCODETO Then
            strSwap = strCodefr
            strCodefr = strCODETO
            strCODETO = strSwap
        End If
    End If

    strSwap = ""
    If Trim(strUserfr) <> "" And Trim(strUserto) <> "" Then
        '*** ＣＯＤＥ入替 ***
        If strUserfr > strUserto Then
            strSwap = strUserfr
            strUserfr = strUserto
            strUserto = strSwap
        End If
    End If

    strSwap = ""
    If Trim(strKbnfr) <> "" And Trim(strKbnto) <> "" Then
        '*** ＣＯＤＥ入替 ***
        If strKbnfr > strKbnto Then
            strSwap = strKbnfr
            strKbnfr = strKbnto
            strKbnto = strSwap
        End If
    End If

    strWhere = ""
    '集金区分
     '← TO
    If Trim(strKbnfr) = "" And Trim(strKbnto) <> "" Then
        strWhere = strWhere & " USER_SKBN <= " & strKbnto & ""
    'FROM ⇔ TO
    ElseIf Trim(strKbnfr) <> "" And Trim(strKbnto) <> "" Then
        If Trim(strKbnfr) <> Trim(strKbnto) Then
            strWhere = strWhere & " USER_SKBN BETWEEN " & strKbnfr & " "
            strWhere = strWhere & " AND " & strKbnto & " "
        Else
            strWhere = strWhere & " USER_SKBN = " & Trim(strKbnfr) & ""
        End If
    'FROM →
    ElseIf Trim(strKbnfr) <> "" And Trim(strKbnto) = "" Then
        strWhere = strWhere & "  USER_SKBN >= " & Trim(strKbnfr) & ""
    '指定なし（全件）
    ElseIf Trim(strKbnfr) = "" And Trim(strKbnto) = "" Then

    End If

    If strWhere <> "" Then
        If strDayfr <> "" Then
            strWhere = strWhere & " AND "
        End If
    End If

    '請求年月日（計上年月に変更 2000/04/20）
    If strDayfr <> "" Then
'''''        strWhere = strWhere & " REQU_YYYYMM = CONVERT(datetime," & Chr$(34) & strDayfr & Chr$(34) & ")"
        strWhere = strWhere & " REQU_KJDATE = CONVERT(datetime," & Chr$(34) & strDayfr & Chr$(34) & ")"
    End If

    If strWhere <> "" Then
        If strUserfr <> "" Or strUserto <> "" Then
            strWhere = strWhere & " AND "
        End If
    End If

    '顧客コード
    '← TO
    If Trim(strUserfr) = "" And Trim(strUserto) <> "" Then
        strWhere = strWhere & " REQU_KCODE <= " & strUserto & ""
    'FROM ⇔ TO
    ElseIf Trim(strUserfr) <> "" And Trim(strUserto) <> "" Then
        If Trim(strUserfr) <> Trim(strUserto) Then
            strWhere = strWhere & " REQU_KCODE BETWEEN " & strUserfr & " "
            strWhere = strWhere & " AND " & strUserto & " "
        Else
            strWhere = strWhere & " REQU_KCODE = " & Trim(strUserfr) & ""
        End If
    'FROM →
    ElseIf Trim(strUserfr) <> "" And Trim(strUserto) = "" Then
        strWhere = strWhere & "  REQU_KCODE >= " & Trim(strUserfr) & ""
    '指定なし（全件）
    ElseIf Trim(strUserfr) = "" And Trim(strUserto) = "" Then

    End If

    If strWhere <> "" Then
        If strCodefr <> "" Or strCODETO <> "" Then
            strWhere = strWhere & " AND "
        End If
    End If

    'ヤードコード
    'FROM →
    If Trim(strCodefr) <> "" And Trim(strCODETO) = "" Then
        strWhere = strWhere & " REQU_YCODE >= " & Trim(strCodefr) & ""
    '← TO
    ElseIf Trim(strCodefr) = "" And Trim(strCODETO) <> "" Then
        strWhere = strWhere & "REQU_YCODE <= " & strCODETO & ""
    'FROM ⇔ TO
    ElseIf Trim(strCodefr) <> "" And Trim(strCODETO) <> "" Then
        If Trim(strCodefr) <> Trim(strCODETO) Then
            strWhere = strWhere & " REQU_YCODE BETWEEN " & strCodefr & " "
            strWhere = strWhere & " AND " & strCODETO & " "
        Else
            strWhere = strWhere & " REQU_YCODE = " & Trim(strCodefr) & ""
        End If
    '指定なし（全件）
    ElseIf Trim(strCodefr) = "" And Trim(strCODETO) = "" Then

    End If

    Make_Where = strWhere

Exit_Make_Where:
    Exit Function

Err_Make_Where:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_Where

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'        UPDATE_DATE      : 2018.06.07
'
'        Rev1.8           : パラメータを追加
'                           ( 0:画面からの指定月, 1:画面からの指定月の前月)
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'SQLServer側にワークテーブル作成
'戻り値 ： strWhere,INPUT,String ,SQL条件文
'**********************************************
'Private Function Make_Work() As Boolean
Private Function Make_Work(intFlg As Integer, Optional webFlg As Integer = 0) As Boolean

    Dim objDb       As Database
    Dim ObjQDF      As QueryDef
    Dim objRst      As Recordset
    Dim strConnect  As String               'Connect情報
    Dim strError    As String               'エラーメッセージ用
    Dim errObj      As Error
    Dim strWhere    As String
    Dim strSQL      As String
'--- 2018.06.07 Rev1.8 ADD START
    Dim strPrvMon   As String               '指定計上月の前月
'--- 2018.06.07 Rev1.8 ADD END
    On Error GoTo Err_Make_Work

    Make_Work = False

'--- 2018.06.07 Rev1.8 UPD START
'    'ワークの削除
'    Call DEL_RTN
    If intFlg = 0 Then
        'ワークの削除
        Call DEL_RTN
    End If
'--- 2018.06.07 Rev1.8 UPD END

'--- 2018.06.07 Rev1.8 UPD START
'    strWhere = Make_Where
    strWhere = Make_Where(intFlg)
'--- 2018.06.07 Rev1.8 UPD END

    Set objDb = CurrentDb

    strConnect = DLookup("CONT_SERVER_PATH", "dbo_CONT_MAST", "CONT_KEY=1")
    If IsNull(strConnect) Then
        MsgBox "ＳＱＬＳｅｒｖｅｒのＰＡＳＳが見つかりません。", vbInformation + vbOKOnly, strMsg_err
        GoTo Exit_Make_Work
    End If

    ' 規定の QueryTimeout プロパティを変更します。
    Debug.Print "Defauit QueryTimeout of Database: " & objDb.QueryTimeout
    objDb.QueryTimeout = 1200                             '２０分（０：タイムアウト無し）
    Debug.Print "New     QueryTimeout of Database: " & objDb.QueryTimeout

    Set ObjQDF = objDb.CreateQueryDef("")
    ObjQDF.Connect = strConnect

    ObjQDF.ReturnsRecords = True

    ' 規定の ODBCTimeout プロパティを変更します。
    Debug.Print "Defauit ODBCTimeout of Database: " & ObjQDF.ODBCTimeout
    ObjQDF.ODBCTimeout = 1200                             '２０分（０：タイムアウト無し）
    Debug.Print "New     ODBCTimeout of Database: " & ObjQDF.ODBCTimeout

    strSQL = ""
    strSQL = "DECLARE @Day_Time DateTime "                            '実行日時

    '文字列の中にｼﾝｸﾞﾙｸｫｰﾃｰｼｮﾝが含まれる場合、エラーになってしまうので、
    'ﾀﾞﾌﾞﾙｸｫｰﾃｰｼｮﾝで囲み、文字として扱うための設定
    '必ず一番最後で、設定をONに戻すこと！
    '戻さないと、ﾃｰﾌﾞﾙを開いたとき、ODBC呼び出しは失敗しましたがでて、見れません！！！
    strSQL = strSQL & "SET QUOTED_IDENTIFIER OFF "
    strSQL = strSQL & "SELECT @Day_Time =CONVERT(DateTime,'" & Trim(m_strTimeStp) & "') "

    'RKS170_WORK_02へ追加
    strSQL = strSQL & "INSERT INTO RKS170_WORK_02"
    
    strSQL = strSQL & "( RKS170_DATETIME,"        '0
    strSQL = strSQL & "RKS170_REQNO,"             '1
    strSQL = strSQL & "RKS170_YYYYMM,"            '2
    strSQL = strSQL & "RKS170_KJDATE,"            '3    1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_SDATE,"             '4
    strSQL = strSQL & "RKS170_NKDATE,"            '5    1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_KCODE,"             '6
    strSQL = strSQL & "RKS170_NKBN,"              '7
'    strSQL = strSQL & "RKS170_KOUM,"              '8   1999/12/15 DELETE Ishii
    strSQL = strSQL & "RKS170_YCODE,"             '8
    strSQL = strSQL & "RKS170_NO,"                '9
    strSQL = strSQL & "RKS170_TTANKA,"            '10
    strSQL = strSQL & "RKS170_TUBOSU,"            '11
    strSQL = strSQL & "RKS170_KINGAK,"            '12
    strSQL = strSQL & "RKS170_SYOZEI,"            '13
    strSQL = strSQL & "RKS170_SECUKG,"            '14
    strSQL = strSQL & "RKS170_TOTAL,"             '15
    strSQL = strSQL & "RKS170_TEKI,"              '16
    strSQL = strSQL & "RKS170_FLG,"               '17   1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_UPDATE,"            '16
    strSQL = strSQL & "RKS170_PAGE,"              '17
    strSQL = strSQL & "RKS170_YARDNAME,"          '18
    strSQL = strSQL & "RKS170_ZEIFLG)"            '19

    strSQL = strSQL & " SELECT  "
    strSQL = strSQL & "@Day_Time,"              '0
    strSQL = strSQL & "REQU_REQNO,"             '1
    strSQL = strSQL & "REQU_YYYYMM,"            '2
    strSQL = strSQL & "REQU_KJDATE,"            '3    1999/12/15 ADD Ishii
    strSQL = strSQL & "REQU_SDATE,"             '4
    strSQL = strSQL & "REQU_NKDATE,"            '5    1999/12/15 ADD Ishii
    strSQL = strSQL & "REQU_KCODE,"             '6
    strSQL = strSQL & "REQU_NKBN,"              '7
'    strSQL = strSQL & "REQU_KOUM,"             '8   1999/12/15 DELETE Ishii
    strSQL = strSQL & "REQU_YCODE,"             '8
    strSQL = strSQL & "REQU_NO,"                '9
    strSQL = strSQL & "REQU_TTANKA,"            '10
    strSQL = strSQL & "REQU_TUBOSU,"            '11
    strSQL = strSQL & "REQU_KINGAK,"            '12
    strSQL = strSQL & "REQU_SYOZEI,"            '13
    strSQL = strSQL & "REQU_SECUKG,"            '14
    strSQL = strSQL & "REQU_TOTAL,"             '15
    strSQL = strSQL & "REQU_TEKI,"              '16
    strSQL = strSQL & "0 AS REQU_FLG,"          '17   1999/12/15 ADD Ishii
    strSQL = strSQL & "REQU_UPDATE,"            '16
    strSQL = strSQL & "0 AS REQU_PAGE,"         '17   ←初期値０
    
'''''    strSQL = strSQL & "B.YARD_NAME,"            '18
    strSQL = strSQL & "ISNULL(B.YARD_NAME,' '),"            '18
    
    strSQL = strSQL & "0 AS RKS170_ZEIFLG"      '19
    strSQL = strSQL & " FROM (((REQU_FILE"
    strSQL = strSQL & " LEFT JOIN USER_MAST A ON REQU_FILE.REQU_KCODE = A.USER_CODE) "
    strSQL = strSQL & " LEFT JOIN YARD_MAST B ON REQU_FILE.REQU_YCODE = B.YARD_CODE) "
    strSQL = strSQL & " LEFT JOIN PAID_FILE C ON REQU_FILE.REQU_REQNO = C.PAID_REQNO) "
'    strSQL = strSQL & " LEFT JOIN PAID_FILE D ON C.PAID_NYU_NO = D.PAIH_NYU_NO) "
    '2015/10/31 M.HONDA INS
    strSQL = strSQL & " INNER Join "
    strSQL = strSQL & " CARG_FILE D ON "
    strSQL = strSQL & " REQU_ACPTNO = CARG_ACPTNO "
    '2015/10/31 M.HONDA INS
    'INSERT 2018/08/04 N.IMAI Start
    strSQL = strSQL & " INNER JOIN RCPT_TRAN"
    strSQL = strSQL & " ON CARG_ACPTNO  = RCPT_CARG_ACPTNO"
    'INSERT 2020/07/03 N.IMAI Start
    strSQL = strSQL & " INNER JOIN YOUK_TRAN"
    strSQL = strSQL & " ON YOUKT_UKNO  = RCPT_NO"
    strSQL = strSQL & " AND YOUKT_YUKBN = 20"
    'INSERT 2020/07/03 N.IMAI End
    strSQL = strSQL & " WHERE ISNULL(REQU_FILE.REQU_TOTAL,0) > ISNULL(C.PAID_NKINGAK,0)"
    'INSERT 2018/08/04 N.IMAI End
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
    strSQL = strSQL & " AND EXISTS"                                             'INSERT START 2010/03/25 K.ISHIZAKA
    strSQL = strSQL & "("
    strSQL = strSQL & " SELECT *"
'    strSQL = strSQL & " FROM RCPT_TRAN"                                        'DELETE 2010/04/10 K.ISHIZAKA
'    strSQL = strSQL & " WHERE RCPT_CARG_ACPTNO = REQU_ACPTNO "                 'DELETE 2010/04/10 K.ISHIZAKA
'    strSQL = strSQL & " AND RCPT_SEIKI = 1"                                    'DELETE 2010/04/10 K.ISHIZAKA
    strSQL = strSQL & " FROM CARG_FILE"                                         'INSERT 2010/04/10 K.ISHIZAKA
    strSQL = strSQL & " WHERE CARG_ACPTNO = REQU_ACPTNO "                       'INSERT 2010/04/10 K.ISHIZAKA
    
    If webFlg = 0 Then
        strSQL = strSQL & " AND CARG_SEIKI IN (1,4)"                                     '2020/03/24 M.HONDA UPD
    Else
        strSQL = strSQL & " AND CARG_SEIKI = 3"
    End If
    strSQL = strSQL & ")"                                                       'INSERT END   2010/03/25 K.ISHIZAKA
    
'--- 2018.06.07 Rev1.8 UPD START
'    strSQL = strSQL & " AND CONVERT(VARCHAR(6),DATEADD(m,ISNULL(CARG_PREMTH_SUM,0),CONVERT(DATETIME,CARG_STDATE,112)),112) < '" & Replace(txt_日付FROM, "/", "") & "'"    '2015/10/31 M.HONDA INS
    If intFlg = 0 Then
        'strSql = strSql & " AND CONVERT(VARCHAR(6),DATEADD(m,ISNULL(CARG_PREMTH_SUM,0),CONVERT(DATETIME,CARG_STDATE,112)),112) < '" & Replace(txt_日付FROM, "/", "") & "'"     'DELETE 2019/04/26 N.IMAI 2015/10/31 M.HONDA INS
        strSQL = strSQL & " AND (CONVERT(VARCHAR(6),DATEADD(m,ISNULL(CARG_PREMTH_SUM,0),CONVERT(DATETIME,CARG_STDATE,112)),112) < '" & Replace(txt_日付FROM, "/", "") & "'"     'INSERT 2019/04/26 N.IMAI
        strSQL = strSQL & " OR ISNULL(RCPT_NEW_HOSPLN_IKOU, '0') = '1' OR  CARG_INSPB = 'TaxREnew10' )  "                                                                                                       'INSERT 2019/04/26 N.IMAI
    Else
        strPrvMon = Format(DateAdd("m", -1, Me.txt_日付FROM & "/01"), "yyyymm")
'        strSql = strSql & " AND CONVERT(VARCHAR(6),DATEADD(m,ISNULL(CARG_PREMTH_SUM,0),CONVERT(DATETIME,CARG_STDATE,112)),112) < '" & Replace(strPrvMon, "/", "") & "'"        'DELETE 2019/04/26 N.IMAI
        strSQL = strSQL & " AND (CONVERT(VARCHAR(6),DATEADD(m,ISNULL(CARG_PREMTH_SUM,0),CONVERT(DATETIME,CARG_STDATE,112)),112) < '" & Replace(strPrvMon, "/", "") & "'"        'INSERT 2019/04/26 N.IMAI
        strSQL = strSQL & " OR ISNULL(RCPT_NEW_HOSPLN_IKOU, '0') = '1' OR  CARG_INSPB = 'TaxREnew10' )  "                                                                                                       'INSERT 2019/04/26 N.IMAI
    End If
'--- 2018.06.07 Rev1.8 UPD END
    If strWhere <> "" Then
        strSQL = strSQL & " AND " & strWhere
    End If

    '必ず必要！！！
    strSQL = strSQL & " SET QUOTED_IDENTIFIER ON "

    ObjQDF.sql = strSQL
    Set objRst = ObjQDF.OpenRecordset(dbOpenSnapshot)

    If objRst(0) <> 0 Then
        MsgBox "ワーク生成中にエラーが発生しました。", vbCritical + vbOKOnly, strMsg_err
        GoTo Exit_Make_Work
    End If

    objDb.Close
    Set objDb = Nothing
    
    Make_Work = True

Exit_Make_Work:
    Exit Function

Err_Make_Work:
    Errors.Refresh
    strError = "ワーク生成中にエラーが発生しました。" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    For Each errObj In DBEngine.Errors
        If Format$(errObj.Number) <> 3325 Then
            strError = strError & Format$(errObj.Number)
            strError = strError & " : " & errObj.Description
            strError = strError & " (" & errObj.Source & ") . "
            strError = strError & Chr$(13) & Chr$(10)
        Else
            'レコードが返されなかったのエラーだから、正常
            Make_Work = True
            GoTo Exit_Make_Work
        End If
    Next
    MsgBox strError, vbCritical + vbOKOnly, strMsg_err
    GoTo Exit_Make_Work

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'SQLSERVERのテーブルよりJETのワークテーブルに取り込み
'戻り値 ： strWhere,INPUT,String ,SQL条件文
'**********************************************
Private Function Make_JetWork() As Boolean

    Dim lngKensu    As Long
    Dim strSQL      As String
    Dim objDb       As Database
    Dim strMsg      As String
    Dim strStampPath    As String                                               'INSERT 20080616 N.MIURA
    
    Dim objRSH      As Recordset
    Dim objRSM      As Recordset

    On Error GoTo Err_Make_JetWork

    Make_JetWork = False

    If Me.frm_社印種別 = 1 Then                                                'INSERT 20080616 N.MIURA START
       strStampPath = fncGetStampPath()
    Else
       strStampPath = ""
    End If                                                                     'INSERT 20080616 N.MIURA ENDED
    
    Set objDb = CurrentDb

    strSQL = ""
    strSQL = strSQL & " INSERT INTO 請求書ＷＫ("
    strSQL = strSQL & "RKS170_DATETIME,"          '0
    strSQL = strSQL & "RKS170_REQNO,"             '1
    strSQL = strSQL & "RKS170_YYYYMM,"            '2
    strSQL = strSQL & "RKS170_KJDATE,"            '3    1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_SDATE,"             '4
    strSQL = strSQL & "RKS170_NKDATE,"            '5    1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_KCODE,"             '6
    strSQL = strSQL & "RKS170_NKBN,"              '7
'    strSQL = strSQL & "RKS170_KOUM,"              '8   1999/12/15 DELETE Ishii
    strSQL = strSQL & "RKS170_YCODE,"             '8
    strSQL = strSQL & "RKS170_NO,"                '9
    strSQL = strSQL & "RKS170_TTANKA,"            '10
    strSQL = strSQL & "RKS170_TUBOSU,"            '11
    strSQL = strSQL & "RKS170_KINGAK,"            '12
    strSQL = strSQL & "RKS170_SYOZEI,"            '13
    strSQL = strSQL & "RKS170_SECUKG,"            '14
    strSQL = strSQL & "RKS170_TOTAL,"             '15
    strSQL = strSQL & "RKS170_TEKI,"              '16
    strSQL = strSQL & "RKS170_FLG,"               '17   1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_UPDATE,"            '16
    strSQL = strSQL & "RKS170_PAGE,"              '17
    strSQL = strSQL & "RKS170_YARDNAME,"          '18
    strSQL = strSQL & "RKS170_ZEIFLG)"            '19

    strSQL = strSQL & " SELECT "
    strSQL = strSQL & "RKS170_DATETIME,"          '0
    strSQL = strSQL & "RKS170_REQNO,"             '1
    strSQL = strSQL & "RKS170_YYYYMM,"            '2
    strSQL = strSQL & "RKS170_KJDATE,"            '3    1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_SDATE,"             '4
    strSQL = strSQL & "RKS170_NKDATE,"            '5    1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_KCODE,"             '6
    strSQL = strSQL & "RKS170_NKBN,"              '7
'    strSQL = strSQL & "RKS170_KOUM,"              '8   1999/12/15 DELETE Ishii
    strSQL = strSQL & "RKS170_YCODE,"             '8
    strSQL = strSQL & "RKS170_NO,"                '9
    strSQL = strSQL & "RKS170_TTANKA,"            '10
    strSQL = strSQL & "RKS170_TUBOSU,"            '11
    strSQL = strSQL & "RKS170_KINGAK,"            '12
    strSQL = strSQL & "RKS170_SYOZEI,"            '13
    strSQL = strSQL & "RKS170_SECUKG,"            '14
    strSQL = strSQL & "RKS170_TOTAL,"             '15
    strSQL = strSQL & "RKS170_TEKI,"              '16
    strSQL = strSQL & "RKS170_FLG,"               '17   1999/12/15 ADD Ishii
    strSQL = strSQL & "RKS170_UPDATE,"            '16
    strSQL = strSQL & "RKS170_PAGE,"              '17
    strSQL = strSQL & "RKS170_YARDNAME,"          '18
    strSQL = strSQL & "RKS170_ZEIFLG"             '19
    strSQL = strSQL & " FROM dbo_RKS170_WORK_02"
    strSQL = strSQL & " WHERE RKS170_DATETIME = CDate(#" & m_strTimeStp & "#)"

    objDb.Execute strSQL, dbFailOnError


   '*** データチェック ***
    'アクションクエリーの結果で判断
    lngKensu = objDb.RecordsAffected
    If lngKensu < 1 Then
        strMsg = "対象となるデータが存在しませんでした。"
        MsgBox strMsg, vbInformation + vbOKOnly, strMsg_kaku
        GoTo Exit_Make_JetWork
    End If

    objDb.Close
    Set objDb = Nothing
    
    '消費税明細レコードの作成
    'If f_InsertTax = False Then                                                'DELETE 20051219 N.MIURA
    '    GoTo Err_Make_JetWork                                                  'DELETE 20051219 N.MIURA
    'End If                                                                     'DELETE 20051219 N.MIURA

    '共益費明細レコードの作成
    'If f_Insertson = False Then      '2000/08/01 insert miura                  DELETE 20041030 N.MIURA START
    '    GoTo Err_Make_JetWork        '2000/08/01 insert miura
    'End If                           '2000/08/01 insert miura                  DELETE 20041030 N.MIURA ENDED

    'ﾍﾟｰｼﾞ項目の更新
    Call PageCount_RTN

    Set objDb = CurrentDb

    '#####
    'strSQL = ""                                                                'DELETE 20050726 N.MIURA START
    'strSQL = strSQL & " INSERT INTO 請求書ヘッダーＷＫ ( RKS170_KCODE, RKS170_KINGAKの合計, USER_CODE, USER_NAME, USER_KANA, USER_TANM, USER_TAKA, USER_YUBINO, USER_ADR_1, USER_ADR_2, USER_TEL, USER_KEITAI, USER_FAX, USER_SDAY, USER_KKBN, USER_SKBN, USER_BIKO, USER_UPDATE, key, RKS170_PAGE )"
    'strSQL = strSQL & " SELECT 請求書ＷＫ.RKS170_KCODE, Sum(請求書ＷＫ.RKS170_KINGAK), dbo_USER_MAST.USER_CODE, dbo_USER_MAST.USER_NAME, dbo_USER_MAST.USER_KANA, dbo_USER_MAST.USER_TANM, dbo_USER_MAST.USER_TAKA, dbo_USER_MAST.USER_YUBINO, dbo_USER_MAST.USER_ADR_1, dbo_USER_MAST.USER_ADR_2, dbo_USER_MAST.USER_TEL, dbo_USER_MAST.USER_KEITAI, dbo_USER_MAST.USER_FAX, dbo_USER_MAST.USER_SDAY, dbo_USER_MAST.USER_KKBN, dbo_USER_MAST.USER_SKBN, dbo_USER_MAST.USER_BIKO, dbo_USER_MAST.USER_UPDATE, Format([RKS170_KCODE],""000000"") , 請求書ＷＫ.RKS170_PAGE"
    'strSQL = strSQL & " FROM 請求書ＷＫ LEFT JOIN dbo_USER_MAST ON 請求書ＷＫ.RKS170_KCODE = dbo_USER_MAST.USER_CODE"
    'strSQL = strSQL & " GROUP BY 請求書ＷＫ.RKS170_KCODE, dbo_USER_MAST.USER_CODE, dbo_USER_MAST.USER_NAME, dbo_USER_MAST.USER_KANA, dbo_USER_MAST.USER_TANM, dbo_USER_MAST.USER_TAKA, dbo_USER_MAST.USER_YUBINO, dbo_USER_MAST.USER_ADR_1, dbo_USER_MAST.USER_ADR_2, dbo_USER_MAST.USER_TEL, dbo_USER_MAST.USER_KEITAI, dbo_USER_MAST.USER_FAX, dbo_USER_MAST.USER_SDAY, dbo_USER_MAST.USER_KKBN, dbo_USER_MAST.USER_SKBN, dbo_USER_MAST.USER_BIKO, dbo_USER_MAST.USER_UPDATE, Format([RKS170_KCODE],""000000""), 請求書ＷＫ.RKS170_PAGE;"
    '                                                                           'DELETE 20050726 N.MIURA ENDED

    strSQL = ""
    strSQL = strSQL & " INSERT INTO 請求書ヘッダーＷＫ "
    strSQL = strSQL & " ( "
    strSQL = strSQL & " RKS170_KCODE, "
    strSQL = strSQL & " RKS170_KINGAKの合計, "
    strSQL = strSQL & " USER_CODE, "
    strSQL = strSQL & " USER_NAME, "
    strSQL = strSQL & " USER_KANA, "
    strSQL = strSQL & " USER_TANM, "
    strSQL = strSQL & " USER_TAKA, "
    strSQL = strSQL & " USER_YUBINO, "
    strSQL = strSQL & " USER_ADR_1, "
    strSQL = strSQL & " USER_ADR_2, "
    strSQL = strSQL & " USER_ADR_3, "
    strSQL = strSQL & " USER_TEL, "
    strSQL = strSQL & " USER_KEITAI, "
    strSQL = strSQL & " USER_FAX, "
    strSQL = strSQL & " USER_SDAY, "
    strSQL = strSQL & " USER_KKBN, "
    strSQL = strSQL & " USER_SKBN, "
    strSQL = strSQL & " USER_BIKO, "
    strSQL = strSQL & " USER_UPDATE, "
    
    '2006/05/10 SHIBAZAKI 追加
    strSQL = strSQL & " USER_RNAME, "
    strSQL = strSQL & " USER_RPOST, "
    strSQL = strSQL & " USER_RADR_1, "
    strSQL = strSQL & " USER_RADR_2, "
    strSQL = strSQL & " USER_RADR_3, "
    
    strSQL = strSQL & " key, "
    strSQL = strSQL & " STAMP_PATH,"                                           'INSERT 20080616 N.MIURA
    strSQL = strSQL & " RKS170_PAGE "
    strSQL = strSQL & " )"
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " 請求書ＷＫ.RKS170_KCODE, "
    'strSQL = strSQL & " Sum(請求書ＷＫ.RKS170_KINGAK), "                       'DELETE 20051219 N.MIURA
    strSQL = strSQL & " Sum(請求書ＷＫ.RKS170_KINGAK + 請求書ＷＫ.RKS170_SYOZEI), " 'INSERT 20051219 N.MIURA
    strSQL = strSQL & " dbo_USER_MAST.USER_CODE, "
    strSQL = strSQL & " dbo_USER_MAST.USER_NAME, "
    strSQL = strSQL & " dbo_USER_MAST.USER_KANA, "
    strSQL = strSQL & " dbo_USER_MAST.USER_TANM, "
    strSQL = strSQL & " dbo_USER_MAST.USER_TAKA, "
    strSQL = strSQL & " dbo_USER_MAST.USER_YUBINO, "
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_1, "
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_2, "
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_3, "
    strSQL = strSQL & " dbo_USER_MAST.USER_TEL, "
    strSQL = strSQL & " dbo_USER_MAST.USER_KEITAI, "
    strSQL = strSQL & " dbo_USER_MAST.USER_FAX, "
    strSQL = strSQL & " dbo_USER_MAST.USER_SDAY, "
    strSQL = strSQL & " dbo_USER_MAST.USER_KKBN, "
    strSQL = strSQL & " dbo_USER_MAST.USER_SKBN, "
    strSQL = strSQL & " dbo_USER_MAST.USER_BIKO, "
    strSQL = strSQL & " dbo_USER_MAST.USER_UPDATE, "
    
    '2006/05/10 SHIBAZAKI 追加
    strSQL = strSQL & " dbo_USER_MAST.USER_RNAME, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RPOST, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RADR_1, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RADR_2, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RADR_3, "
    
    strSQL = strSQL & " Format([RKS170_KCODE],""000000"") , "
    strSQL = strSQL & "'" & strStampPath & "',"                                'INSERT 20080616 N.MIURA
    strSQL = strSQL & " 請求書ＷＫ.RKS170_PAGE"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & " 請求書ＷＫ "
    strSQL = strSQL & " LEFT JOIN dbo_USER_MAST ON 請求書ＷＫ.RKS170_KCODE = dbo_USER_MAST.USER_CODE"
    strSQL = strSQL & " GROUP BY "
    strSQL = strSQL & " 請求書ＷＫ.RKS170_KCODE, "
    strSQL = strSQL & " dbo_USER_MAST.USER_CODE, "
    strSQL = strSQL & " dbo_USER_MAST.USER_NAME, "
    strSQL = strSQL & " dbo_USER_MAST.USER_KANA, "
    strSQL = strSQL & " dbo_USER_MAST.USER_TANM, "
    strSQL = strSQL & " dbo_USER_MAST.USER_TAKA, "
    strSQL = strSQL & " dbo_USER_MAST.USER_YUBINO, "
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_1, "
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_2, "
    strSQL = strSQL & " dbo_USER_MAST.USER_ADR_3, "
    strSQL = strSQL & " dbo_USER_MAST.USER_TEL, "
    strSQL = strSQL & " dbo_USER_MAST.USER_KEITAI, "
    strSQL = strSQL & " dbo_USER_MAST.USER_FAX, "
    strSQL = strSQL & " dbo_USER_MAST.USER_SDAY, "
    strSQL = strSQL & " dbo_USER_MAST.USER_KKBN, "
    strSQL = strSQL & " dbo_USER_MAST.USER_SKBN, "
    strSQL = strSQL & " dbo_USER_MAST.USER_BIKO, "
    strSQL = strSQL & " dbo_USER_MAST.USER_UPDATE, "
    
    '2006/05/10 SHIBAZAKI 追加
    strSQL = strSQL & " dbo_USER_MAST.USER_RNAME, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RPOST, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RADR_1, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RADR_2, "
    strSQL = strSQL & " dbo_USER_MAST.USER_RADR_3, "
    
    strSQL = strSQL & " Format([RKS170_KCODE],""000000""), "
    strSQL = strSQL & " 請求書ＷＫ.RKS170_PAGE"
    strSQL = strSQL & " ;"

    objDb.Execute strSQL, dbFailOnError

    'objDb.Execute strKOUZ_MAST, dbFailOnError                                                    'INSERT 20080204 N.MIURA
    Call subKOUZ_MAST_Set                                                       'INSERT 2023/02/24 N.IMAI
    objDb.Execute strBANK_TABL, dbFailOnError                                                    'INSERT 20080204 N.MIURA

    Set objRSH = objDb.OpenRecordset("請求書ヘッダーＷＫ", dbOpenDynaset)

    If objRSH.EOF = True And objRSH.BOF = True Then
    Else
        objRSH.MoveFirst

        Do Until objRSH.EOF
            strSQL = ""
            'strSQL = strSQL & " SELECT Sum(請求書ＷＫ.RKS170_KINGAK)"                           'DELETE 20051219 N.MIURA
            strSQL = strSQL & " SELECT Sum(請求書ＷＫ.RKS170_KINGAK + 請求書ＷＫ.RKS170_SYOZEI)" 'INSERT 20051219 N.MIURA
            'INSERT 2019/08/15 Y.WADA Start
            strSQL = strSQL & " , Sum(請求書ＷＫ.RKS170_SYOZEI) As SYOZEI"
            strSQL = strSQL & " , Max(RKS170_KJDATE) As KJDATE"
            'INSERT 2019/08/15 Y.WADA End
            strSQL = strSQL & " FROM 請求書ＷＫ"
            strSQL = strSQL & " WHERE 請求書ＷＫ.RKS170_KCODE =" & objRSH![RKS170_KCODE]
    
            Set objRSM = objDb.OpenRecordset(strSQL, dbOpenSnapshot)
    
            objRSH.Edit
            objRSH![RKS170_KINGAKの合計] = objRSM.Fields(0)
            
            'INSERT 2019/08/15 Y.WADA Start
            Dim strDate1        As String
            Dim dblZeiRitu1     As Double
            Dim strRoundKbn1    As String
            
            '税率
            strDate1 = Replace(Nz(objRSM.Fields("KJDATE")), "/", "")
            dblZeiRitu1 = MSZZ004_M20(strDate1, strRoundKbn1)
            
            objRSH![TAX_RATE] = dblZeiRitu1 * 100
            objRSH![TAX] = objRSM.Fields("SYOZEI")
            'INSERT 2019/08/15 Y.WADA End
            
            objRSH.UPDATE
    
            objRSH.MoveNext
    
            objRSM.Close
            Set objRSM = Nothing
        Loop
    End If

    objRSH.Close
    Set objRSH = Nothing

    objDb.Close
    Set objDb = Nothing

    Make_JetWork = True

Exit_Make_JetWork:
    Exit Function

Err_Make_JetWork:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Make_JetWork

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :strKOUZ_MAST
'        CREATE_DATE      :2008/02/04 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function strKOUZ_MAST()
    Dim strSQL          As String

    strSQL = ""
    strSQL = strSQL & "UPDATE     請求書ヘッダーＷＫ,"
'DELETE 2019/10/19 N.IMAI Start
'    strSQL = strSQL & "           KOUZ_MAST "
'    strSQL = strSQL & "INNER JOIN dbo_CONT_MAST "
'    strSQL = strSQL & "ON        (KOUZ_MAST.KOUZM_BUMOC          = dbo_CONT_MAST.CONT_BUMOC) "
''    strSQL = strSQL & "AND       (KOUZ_MAST.KOUZM_CAMPC          = dbo_CONT_MAST.CONT_SENDEN1) "   'DELETE 2018/10/15 N.IMAI
'    strSQL = strSQL & "SET        請求書ヘッダーＷＫ.KOUZM_KINYC = KOUZ_MAST.KOUZM_KINYC,"
'    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_SHITC = KOUZ_MAST.KOUZM_SHITC,"
'    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_YOKII = VAL(KOUZ_MAST.KOUZM_YOKII),"
'    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_KOUZB = KOUZ_MAST.KOUZM_KOUZB "
'    strSQL = strSQL & "WHERE      Format([USER_CODE],'000000')   = KOUZ_MAST.KOUZM_KOKYC "
'DELETE 2019/10/19 N.IMAI End
    strSQL = strSQL & "          (KOUZ_MAST_J "
    strSQL = strSQL & "INNER JOIN dbo_CONT_MAST "
    strSQL = strSQL & "ON        (KOUZ_MAST_J.KOUZM_BUMOC          = dbo_CONT_MAST.CONT_BUMOC)) "
    strSQL = strSQL & "INNER JOIN CAMP_MAST "                                                   'INSERT 2022/08/10 N.IMAI
    strSQL = strSQL & "ON        (KOUZ_MAST_J.KOUZM_CAMPC          = CAMP_MAST.CAMPM_CAMPC) "   'INSERT 2022/08/10 N.IMAI
    strSQL = strSQL & "SET        請求書ヘッダーＷＫ.KOUZM_KINYC = KOUZ_MAST_J.KOUZM_KINYC,"
    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_SHITC = KOUZ_MAST_J.KOUZM_SHITC,"
    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_YOKII = VAL(KOUZ_MAST_J.KOUZM_YOKII),"
    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_KOUZB = KOUZ_MAST_J.KOUZM_KOUZB,"
    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_KOUZN = CAMPM_CAMPN "                'INSERT 2022/08/10 N.IMAI
    strSQL = strSQL & "WHERE      Format([USER_CODE],'000000')   = KOUZ_MAST_J.KOUZM_KOKYC "
    strSQL = strSQL & ";"
    
    strKOUZ_MAST = strSQL
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :strBANK_TABL
'        CREATE_DATE      :2008/02/04 N.MIURA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Function strBANK_TABL()
    Dim strSQL          As String

    strSQL = ""
    strSQL = strSQL & "UPDATE     CODE_TABL, "
    strSQL = strSQL & "           請求書ヘッダーＷＫ "
    strSQL = strSQL & "INNER JOIN BANK_TABL "
    strSQL = strSQL & "ON        (請求書ヘッダーＷＫ.KOUZM_SHITC  = BANK_TABL.BANKT_SHITC) "
    strSQL = strSQL & "AND       (請求書ヘッダーＷＫ.KOUZM_KINYC  = BANK_TABL.BANKT_KINYC) "
    
    strSQL = strSQL & "SET        請求書ヘッダーＷＫ.KOUZM_KINYN  = BANK_TABL.BANKT_KINYN,"
    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_SHITN  = BANK_TABL.BANKT_SHITN,"
    strSQL = strSQL & "           請求書ヘッダーＷＫ.KOUZM_YOKIN  = CODE_TABL.CODET_NAMEN "
    
    'strSQL = strSQL & "WHERE  Val(請求書ヘッダーＷＫ.KOUZM_YOKII) = CODE_TABL.CODET_CODEC "
    strSQL = strSQL & "WHERE      請求書ヘッダーＷＫ.KOUZM_YOKII  = CODE_TABL.CODET_CODEC "
    strSQL = strSQL & "AND        CODE_TABL.CODET_SIKBC           = '121';"
    
    strBANK_TABL = strSQL

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20051219 N.MIURA
'**********************************************
'消費税明細の作成・・・顧客ｺｰﾄﾞごとに作成する。
'**********************************************
'Private Function f_InsertTax() As Boolean
'    Dim strSQL      As String
'    Dim objDb       As Database
'    Dim strMsg      As String
'    On Error GoTo Err_f_InsertTax
'    f_InsertTax = False
'    Set objDb = CurrentDb
'    'STRSQL = ""
'    'STRSQL = STRSQL & " INSERT INTO 請求書ＷＫ("                                                 '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "RKS170_KCODE,"               '0                                          '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "RKS170_KINGAK,"              '1                                          '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "RKS170_PAGE,"                '2                                          '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "RKS170_YARDNAME,"            '3                                          '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "RKS170_ZEIFLG)"              '4                                          '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & " SELECT "                                                                '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "RKS170_KCODE  AS RKS170_KCODE,"                                       '0 '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "SUM(IIF(RKS170_SYOZEI IS NULL ,0,RKS170_SYOZEI)) AS RKS170_KINGAKU,"  '1 '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "0             AS RKS170_PAGE,"                                        '2 '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "'消費税'      AS RKS170_YARDNAME,"                                    '3 '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & "1           AS RKS170_ZEIFLG"                                         '4 '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & " FROM 請求書ＷＫ"                                                         '2000/08/01 DELETE MIURA
'    'STRSQL = STRSQL & " GROUP BY RKS170_KCODE"                                                   '2000/08/01 DELETE MIURA
'    strSQL = ""
'    strSQL = strSQL & "INSERT INTO 請求書ＷＫ ( "
'    strSQL = strSQL & "RKS170_KCODE, "
'    strSQL = strSQL & "RKS170_KJDATE, "
'    strSQL = strSQL & "RKS170_YCODE, "
'    strSQL = strSQL & "RKS170_KINGAK, "
'    strSQL = strSQL & "RKS170_PAGE, "
'    strSQL = strSQL & "RKS170_YARDNAME, "
'    strSQL = strSQL & "RKS170_ZEIFLG "
'    strSQL = strSQL & ") "
'    strSQL = strSQL & "SELECT "
'    strSQL = strSQL & "請求書ＷＫ.RKS170_KCODE                          AS RKS170_KCODE,  "
'    strSQL = strSQL & "First(請求書ＷＫ.RKS170_KJDATE)                  AS RKS170_KJDATE, "
'    strSQL = strSQL & "First(999999)                                    AS RKS170_YCODE,  "
'    strSQL = strSQL & "Sum(IIf(RKS170_SYOZEI Is Null,0,RKS170_SYOZEI))  AS RKS170_KINGAK, "
'    strSQL = strSQL & "First(0)                                         AS RKS170_PAGE,   "
'    strSQL = strSQL & "First('消費税') AS RKS170_YARDNAME, First(2)     AS RKS170_ZEIFLG  "
'    strSQL = strSQL & "FROM 請求書WK "
'    strSQL = strSQL & "Group BY "
'    strSQL = strSQL & "請求書WK.RKS170_KCODE "
'    strSQL = strSQL & "HAVING "
'    strSQL = strSQL & "Sum(IIf(RKS170_SYOZEI Is Null,0,RKS170_SYOZEI)) <> 0 "
'    strSQL = strSQL & "; "
'    objDb.Execute strSQL, dbFailOnError
'    objDb.Close
'    Set objDb = Nothing
'    f_InsertTax = True
'Exit_f_InsertTax:
'    Exit Function
'Err_f_InsertTax:
'    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
'    Resume Exit_f_InsertTax
'End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20041030 N.MIURA
'**********************************************
'共益費明細の作成・・・明細毎に作成する。
'**********************************************
'Private Function f_Insertson() As Boolean                                           '2000/08/01 INSER MIURA
'    Dim strSQL      As String                                                       '2000/08/01 INSER MIURA
'    Dim objDB       As Database                                                     '2000/08/01 INSER MIURA
'    Dim strMSG      As String                                                       '2000/08/01 INSER MIURA
'    'On Error GoTo Err_f_Insertson                                                  '2000/08/01 INSER MIURA
'    f_Insertson = False                                                             '2000/08/01 INSER MIURA
'    Set objDB = CurrentDb                                                           '2000/08/01 INSER MIURA
'    strSQL = ""                                                                     '2000/08/01 INSER MIURA
'    strSQL = strSQL & "INSERT INTO 請求書ＷＫ( "                                    '2000/08/01 INSER MIURA
'    strSQL = strSQL & "RKS170_KJDATE, "
'    strSQL = strSQL & "RKS170_KCODE, "                                              '2000/08/01 INSER MIURA
'    strSQL = strSQL & "RKS170_YCODE, "                                              '2000/08/01 INSER MIURA
'    strSQL = strSQL & "RKS170_NO, "                                                 '2000/08/01 INSER MIURA
'    strSQL = strSQL & "RKS170_KINGAK, "                                             '2000/08/01 INSER MIURA
'    strSQL = strSQL & "RKS170_PAGE, "                                               '2000/08/01 INSER MIURA
'    strSQL = strSQL & "RKS170_YARDNAME, "                                           '2000/08/01 INSER MIURA
'    strSQL = strSQL & "RKS170_ZEIFLG "                                              '2000/08/01 INSER MIURA
'    strSQL = strSQL & ") "                                                          '2000/08/01 INSER MIURA
'    strSQL = strSQL & "SELECT "                                                     '2000/08/01 INSER MIURA
'    strSQL = strSQL & "請求書ＷＫ.RKS170_KJDATE  AS RKS170_KJDATE, "                '2000/08/01 INSER MIURA
'    strSQL = strSQL & "請求書ＷＫ.RKS170_KCODE   AS RKS170_KCODE, "                 '2000/08/01 INSER MIURA
'    strSQL = strSQL & "請求書ＷＫ.RKS170_YCODE   AS RKS170_YCODE, "                 '2000/08/01 INSER MIURA
'    strSQL = strSQL & "請求書ＷＫ.RKS170_NO      AS RKS170_NO, "                    '2000/08/01 INSER MIURA
'    'strSQL = strSQL & "KYOU_TRAN.KYOUT_KYOUA    AS RKS170_KINGAK, "                '2000/08/01 INSER MIURA
'    strSQL = strSQL & "dbo_KYOU_TRAN.KYOUT_KYOUA AS RKS170_KINGAK, "                '2003/05/06 INSER MIURA
'    strSQL = strSQL & "0                        AS RKS170_PAGE, "                   '2000/08/01 INSER MIURA
'    strSQL = strSQL & "'共益費（税込み）'       AS RKS170_YARDNAME, "               '2000/08/01 INSER MIURA
'    strSQL = strSQL & "1                        AS RKS170_ZEIFLG "                  '2000/08/01 INSER MIURA
'    strSQL = strSQL & "FROM ((請求書ＷＫ INNER JOIN dbo_CARG_FILE "                 '2000/08/01 INSER MIURA
'    strSQL = strSQL & "ON  (請求書ＷＫ.RKS170_YCODE = dbo_CARG_FILE.CARG_YCODE) "   '2000/08/01 INSER MIURA
'    strSQL = strSQL & "AND (請求書ＷＫ.RKS170_NO    = dbo_CARG_FILE.CARG_NO) "      '2000/08/01 INSER MIURA
'    strSQL = strSQL & "AND (請求書ＷＫ.RKS170_KCODE = dbo_CARG_FILE.CARG_UCODE)) "  '2000/08/01 INSER MIURA
'    strSQL = strSQL & "INNER JOIN dbo_CNTA_MAST "                                   '2000/08/01 INSER MIURA
'    strSQL = strSQL & "ON  (dbo_CARG_FILE.CARG_YCODE = dbo_CNTA_MAST.CNTA_CODE) "   '2000/08/01 INSER MIURA
'    strSQL = strSQL & "AND (dbo_CARG_FILE.CARG_NO = dbo_CNTA_MAST.CNTA_NO)) "       '2000/08/01 INSER MIURA
'    'strSQL = strSQL & "LEFT JOIN KYOU_TRAN "                                       '2000/08/01 INSER MIURA
'    'strSQL = strSQL & "ON   dbo_CNTA_MAST.CNTA_SYU = KYOU_TRAN.KYOUT_KYOUC "       '2000/08/01 INSER MIURA
'    strSQL = strSQL & "LEFT JOIN dbo_KYOU_TRAN "                                    '2003/05/06 INSER MIURA
'    strSQL = strSQL & "ON   dbo_CNTA_MAST.CNTA_SYU = dbo_KYOU_TRAN.KYOUT_KYOUC "    '2003/05/06 INSER MIURA
'    strSQL = strSQL & "WHERE (((請求書ＷＫ.RKS170_ZEIFLG)=0) "                      '2000/08/01 INSER MIURA
'    strSQL = strSQL & "AND   ((dbo_CARG_FILE.CARG_DOCU2)=-1)) "                     '2000/08/01 INSER MIURA
'    objDB.Execute strSQL, dbFailOnError                                             '2000/08/01 INSER MIURA
'    objDB.Close                                                                     '2000/08/01 INSER MIURA
'    Set objDB = Nothing                                                             '2000/08/01 INSER MIURA
'    f_Insertson = True                                                             '2000/08/01 INSER MIURA
'Exit_f_Insertson:                                                                   '2000/08/01 INSER MIURA
'    Exit Function                                                                   '2000/08/01 INSER MIURA
'Err_f_Insertson:                                                                    '2000/08/01 INSER MIURA
'    MsgBox "ｴﾗｰ番号:" & err.Number & vbCr & err.Description, vbInformation + vbOKOnly, strMsg_err '2000/08/01 INSER MIURA
'    Resume Exit_f_Insertson                                                         '2000/08/01 INSER MIURA
'End Function                                                                        '2000/08/01 INSER MIURA
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'請求書のページカウント
'**********************************************
Private Sub PageCount_RTN()

    Dim strSQL          As String
    Dim ojDyn           As Recordset
    Dim varUSER         As Variant      '顧客ｺｰﾄﾞ
    Dim lngPage         As Long         'ﾍﾟｰｼﾞ
    Dim lngGyo          As Long         '行
    Dim sql             As String

    'STRSQL = ""                                                '2000/08/01 DELETE MIURA
    'STRSQL = "SELECT 請求書ＷＫ.* "                            '2000/08/01 DELETE MIURA
    'STRSQL = STRSQL & "FROM 請求書ＷＫ "                       '2000/08/01 DELETE MIURA
    'STRSQL = STRSQL & "ORDER BY RKS170_KCODE,RKS170_ZEIFLG;"   '2000/08/01 DELETE MIURA

    strSQL = ""                                                 '2000/08/01 INSERT MIURA
    strSQL = "SELECT 請求書ＷＫ.* "                             '2000/08/01 INSERT MIURA
    strSQL = strSQL & "FROM 請求書ＷＫ "                        '2000/08/01 INSERT MIURA
    strSQL = strSQL & "ORDER BY "                               '2000/08/01 INSERT MIURA
    strSQL = strSQL & "請求書ＷＫ.RKS170_KJDATE, "              '2000/08/01 INSERT MIURA
    strSQL = strSQL & "請求書ＷＫ.RKS170_KCODE, "               '2000/08/01 INSERT MIURA
    strSQL = strSQL & "請求書ＷＫ.RKS170_YCODE, "               '2000/08/01 INSERT MIURA
    strSQL = strSQL & "請求書ＷＫ.RKS170_NO, "                  '2000/08/01 INSERT MIURA
    strSQL = strSQL & "請求書ＷＫ.RKS170_ZEIFLG "               '2000/08/01 INSERT MIURA
    strSQL = strSQL & ";"                                       '2000/08/01 INSERT MIURA

    Set ojDyn = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)

    lngPage = 0
    lngGyo = 0

    If ojDyn.EOF Then
        Else
            ojDyn.MoveFirst
            Do Until ojDyn.EOF

            ojDyn.Edit  'カレントレコードを編集できるように編集バッファにコピー

            If varUSER <> ojDyn!RKS170_KCODE Then
                'もし顧客コードが変わったらページカウントアップ
                lngPage = 1
                lngGyo = 1
            ElseIf varUSER = ojDyn!RKS170_KCODE Then
                '顧客コードが変わらなかったら・・・
                lngGyo = lngGyo + 1
            End If

            'If lngGyo > 9 Then                                                 2001/04/11 INSERT N.MIURA
            'If lngGyo > 32 Then                                                2005/12/19 DELETE N.MIURA
            If lngGyo > 27 Then
                lngPage = lngPage + 1
                lngGyo = 1
            End If

            'カレントレコードの編集
            ojDyn!RKS170_PAGE = lngPage
            ojDyn.UPDATE

            varUSER = ojDyn!RKS170_KCODE

      ojDyn.MoveNext
    Loop

  End If

  ojDyn.Close

  Set ojDyn = Nothing

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'*********************************
'ワーク削除
'*********************************
Private Sub DEL_RTN()

    Dim objDb   As Database
    Dim strSQL  As String

    On Error Resume Next

    Set objDb = CurrentDb()

    '*** SQLServer ***
    strSQL = ""
    strSQL = "DELETE FROM dbo_RKS170_WORK_02"
    strSQL = strSQL & " WHERE RKS170_DATETIME = CDate(#" & m_strTimeStp & "#)"
    strSQL = strSQL & " OR RKS170_DATETIME <= CDate(#" & m_strZTimeStp & "#)"
    objDb.Execute strSQL

    '*** JET ***
    strSQL = ""
    strSQL = "DELETE FROM 請求書ヘッダーＷＫ"
    objDb.Execute strSQL, dbFailOnError

    '*** JET ***
    strSQL = ""
    strSQL = "DELETE FROM 請求書ＷＫ"
    objDb.Execute strSQL, dbFailOnError

    objDb.Close
    Set objDb = Nothing

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'cmb_CODEFROM_AfterUpdateイベント
'**********************************************
Private Sub cmb_CODEFROM_AfterUpdate()

    Dim varCode As Variant
    Dim strCARG As String
    Dim strWhere As String

    On Error GoTo Err_cmb_CODEFROM_AfterUpdate

    varCode = Me.cmb_CODEFROM
    strCARG = ""
    
    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " dbo_YARD_MAST.YARD_CODE=" & varCode & ""
        strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
    End If
    Me.cmb_CODEFROM = Format(Me.cmb_CODEFROM, "000000")
    Me.txt_CODEFROM = strCARG
    Me.cmb_CODETO.SetFocus

Exit_cmb_CODEFROM_AfterUpdate:
    Exit Sub

Err_cmb_CODEFROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODEFROM_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'cmb_CODEFROM_KeyPressイベント
'**********************************************
Private Sub cmb_CODEFROM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODEFROM, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'cmb_CODETO_AfterUpdateイベント
'**********************************************
Private Sub cmb_CODETO_AfterUpdate()

    Dim varCode As Variant
    Dim strCARG As String
    Dim strWhere As String

    On Error GoTo Err_cmb_CODETO_AfterUpdate

    varCode = Me.cmb_CODETO
    strCARG = ""
    
    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " YARD_CODE=" & varCode & ""
        strCARG = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strWhere), "")
    End If
    Me.cmb_CODETO = Format(Me.cmb_CODETO, "000000")
    Me.txt_CODETO = strCARG
    Me.cmb_顧客FROM.SetFocus

Exit_cmb_CODETO_AfterUpdate:
    Exit Sub

Err_cmb_CODETO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODETO_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'cmb_顧客TO_AfterUpdateイベント
'**********************************************
Private Sub cmb_区分FROM_AfterUpdate()
    
    Dim varCode As Variant
    Dim strCARG As String
    Dim strWhere As String
    
    On Error GoTo Err_cmb_区分FROM_AfterUpdate

    varCode = Me.cmb_区分FROM
    strCARG = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " NAME_ID= '" & "030" & "'"
        strWhere = strWhere & " AND NAME_CODE=" & varCode & ""
        strCARG = Nz(DLookup("NAME_RYAk", "dbo_NAME_MAST", strWhere), "")
    End If
    Me.txt_区分FROM = strCARG
    Me.cmb_区分TO.SetFocus

Exit_cmb_区分FROM_AfterUpdate:
    Exit Sub

Err_cmb_区分FROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_区分FROM_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'cmb_区分ＦＲＯＭ_KeyPressイベント
'**********************************************
Private Sub cmb_区分FROM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_区分FROM, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'cmb_顧客TO_AfterUpdateイベント
'**********************************************
Private Sub cmb_区分TO_AfterUpdate()
    
    Dim varCode As Variant
    Dim strCARG As String
    Dim strWhere As String
    
    On Error GoTo Err_cmb_区分TO_AfterUpdate
    
    varCode = Me.cmb_区分TO
    strCARG = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " NAME_ID= '" & "030" & "'"
        strWhere = strWhere & " AND NAME_CODE=" & varCode & ""
        strCARG = Nz(DLookup("NAME_RYAk", "dbo_NAME_MAST", strWhere), "")
    End If
    Me.txt_区分TO = strCARG
    Me.cmb_CODEFROM.SetFocus

Exit_cmb_区分TO_AfterUpdate:
    Exit Sub

Err_cmb_区分TO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_区分TO_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_区分ＴＯ_KeyPressイベント
'        MODULE_ID        :cmb_区分TO_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_区分TO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_区分TO, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客FROM_KeyPressイベント
'        MODULE_ID        :cmb_顧客FROM_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客FROM_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客FROM, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客TO_AfterUpdateイベント
'        MODULE_ID        :cmb_顧客TO_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客TO_AfterUpdate()

    Dim varCode As Variant
    Dim strCARG As String
    Dim strWhere As String

    On Error GoTo Err_cmb_顧客TO_AfterUpdate

    varCode = Me.cmb_顧客TO
    strCARG = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strCARG = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If
    Me.cmb_顧客TO = Format(Me.cmb_顧客TO, "000000")
    Me.txt_顧客TO = strCARG
    'Me.cmd_印刷.SetFocus                                                       'DELETE 20051219 N.MIURA
    Me.cmd_プレビュー.SetFocus                                                  'INSERT 20051219 N.MIURA

Exit_cmb_顧客TO_AfterUpdate:
    Exit Sub

Err_cmb_顧客TO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_顧客TO_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客FROM_AfterUpdateイベント
'        MODULE_ID        :cmb_顧客FROM_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客FROM_AfterUpdate()

    Dim varCode As Variant
    Dim strCARG As String
    Dim strWhere As String

    On Error GoTo Err_cmb_CODETO_AfterUpdate

    varCode = Me.cmb_顧客FROM
    strCARG = ""

    If Trim(varCode) <> "" Then
        strWhere = ""
        strWhere = strWhere & " USER_CODE=" & varCode & ""
        strCARG = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")
    End If
    Me.cmb_顧客FROM = Format(Me.cmb_顧客FROM, "000000")
    Me.Txt_顧客FROM = strCARG
    Me.cmb_顧客TO.SetFocus

Exit_cmb_CODETO_AfterUpdate:
    Exit Sub

Err_cmb_CODETO_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmb_CODETO_AfterUpdate

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_CODETO_KeyPressイベント
'        MODULE_ID        :cmb_CODETO_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_CODETO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_CODETO, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_顧客TO_KeyPressイベント
'        MODULE_ID        :cmb_顧客TO_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客TO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_顧客TO, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_プレビュー_Clickイベント
'        MODULE_ID        :cmd_プレビュー_Click
'        CREATE_DATE      :
'        UPDATE_DATE      : 2018.06.07 TFC
'        Rev1.8           : Make_Workへパラメータ追加
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_プレビュー_Click()

    Dim strMsg As String
    Dim intAns As Integer

    On Error GoTo Err_cmd_プレビュー_Click

    strMsg = "印刷内容を画面に出力します。" '& vbCr
    strMsg = strMsg & "よろしいですか？"
    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)

    If intAns = vbNo Then
        GoTo Exit_cmd_プレビュー_Click
    End If

    '*** ワーク作成 ***
    doCmd.Hourglass True

    Call DEL_RTN
 '--- 2018.06.07 Rev1.8 UPD START
'    If Make_Work = False Then
    If Make_Work(0) = False Then
'--- 2018.06.07 Rev1.8 UPD END
        '*** エラー発生またはデータなし ***
        GoTo Exit_cmd_プレビュー_Click
    End If
    '*** Jetワーク作成 ***
    If Make_JetWork = False Then
        'エラー発生またはデータなし
        GoTo Exit_cmd_プレビュー_Click
    End If

    'doCmd.OpenReport "RKS240_請求明細書", acPreview                            'DELETE 20051219 N.MIURA
    doCmd.OpenReport "RKS240", acPreview                                        'INSERT 20051219 N.MIURA
    
    'INSERT 2017/12/01 N.IMAI Start
    Dim strDate         As String
    Dim strTime         As String
    Dim strFilename     As String
    strDate = Format$(Now, "yyyyMMdd")
    strTime = Format$(Now, "hhnnss")
    strFilename = "FKS240_" & strDate & "_" & strTime & ".xls"
    Call Make_Output_List(strDate, strTime)
    'INSERT 2017/12/01 N.IMAI End

Exit_cmd_プレビュー_Click:
    doCmd.Hourglass False
    Me.txt_日付FROM.SetFocus
    Exit Sub

Err_cmd_プレビュー_Click:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_cmd_プレビュー_Click
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_印刷_Clickイベント
'        MODULE_ID        :cmd_印刷_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8'DELETE 20051219 N.MIURA
'Private Sub cmd_印刷_Click()
'    Dim strMsg As String
'    Dim intAns As Integer
'    On Error GoTo Err_cmd_印刷_Click
'    strMsg = "印刷内容をプリンタに印字します。" & vbCr
'    strMsg = strMsg & "よろしいですか？"
'    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)
'    If intAns = vbNo Then
'        GoTo Exit_cmd_印刷_Click
'    End If
'    '*** ワーク作成 ***
'    doCmd.Hourglass True
'    Call DEL_RTN
'    If Make_Work = False Then
'        'エラー発生またはデータなし
'        GoTo Exit_cmd_印刷_Click
'    End If
'    '*** Jetワーク作成 ***
'    If Make_JetWork = False Then
'        'エラー発生またはデータなし
'        GoTo Exit_cmd_印刷_Click
'    End If
'    '印刷実行
'    doCmd.OpenReport "RKS240_請求明細書", acNormal
'Exit_cmd_印刷_Click:
'    doCmd.Hourglass False
'    Me.txt_日付FROM.SetFocus
'    Exit Sub
'Err_cmd_印刷_Click:
'    Select Case Err
'        Case 2501:
'            MsgBox "印刷を中止しました!!", vbInformation + vbOKOnly, strMsg_kaku
'        Case 2201:
'            MsgBox "プリンタドライバーの設定に誤りがあります。", vbInformation + vbOKOnly, strMsg_kaku
'        Case Else:
'            MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
'    End Select
'    Resume Exit_cmd_印刷_Click
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_終了_Clickイベント
'        MODULE_ID        :cmd_終了_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()
    On Error Resume Next
    'doCmd.Close acForm, "FKS240_請求明細書"                                    'DELETE 20041124 N.MIURA
    doCmd.Close acForm, "FKS240"                                                'INSERT 20041124 N.MIURA
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'Form_Activateイベント
'**********************************************
Private Sub Form_Activate()
    doCmd.Maximize
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Error
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'Form_Errorイベント
'**********************************************
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
        Case 2113
            Response = acDataErrContinue
            MsgBox "数字を入力してください。", 0 + 64, strMsg_kaku
            SendKeys "{ESC}"
            Exit Sub
        Case 2279
            Response = acDataErrContinue
            MsgBox "正しく入力してください。", 0 + 64, strMsg_kaku
            SendKeys "{ESC}"
            Exit Sub
        Case 3316
            Response = acDataErrContinue
            SendKeys "{ESC}{ESC}"
            Exit Sub
        Case Else
            MsgBox DataErr
    End Select

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'Form_KeyDownイベント
'**********************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Err_Form_KeyDown

    'SHIFT、CTRL、ALTのいずれかが押されている場合は何もしない
    If Shift > 0 Then
        Exit Sub
    End If

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus
        Case vbKeyF2
            KeyCode = 0
            'Me.cmd_印刷.SetFocus                                               'DELETE 20051219 N.MIURA
            Me.cmd_プレビュー.SetFocus                                          'INSERT 20051219 N.MIURA
    End Select

Exit_Form_KeyDown:
    Exit Sub

Err_Form_KeyDown:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_KeyDown

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Unload
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'Form_Unloadイベント
'**********************************************
Private Sub Form_Unload(Cancel As Integer)

    Dim strMsg As String
    Dim intAns As Integer

    On Error GoTo Err_Form_Unload

    '' 2013/06/03 M.HONDA START
'    strMsg = "終了します。" '& vbCr
'    strMsg = strMsg & "よろしいですか？"
'    intAns = MsgBox(strMsg, vbQuestion + vbYesNo, strMsg_kaku)
'
'    If intAns = vbNo Then
'        Cancel = True
'        Me.txt_日付FROM.SetFocus
'        GoTo Exit_Form_Unload
'    Else
'        Call DEL_RTN
'    End If
    '' 2013/06/03 M.HONDA END

Exit_Form_Unload:
    Exit Sub

Err_Form_Unload:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Resume Exit_Form_Unload

End Sub
'''==============================================================================*
'''
'''        MODULE_NAME      :
'''        MODULE_ID        :txt_日付FROM_AfterUpdate
'''        CREATE_DATE      :2011/01/08 M.RYU
'''
'''==============================================================================*
'''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'''**********************************************
'''txt_日付FROM_AfterUpdateイベント
'''**********************************************
''Private Sub txt_日付FROM_AfterUpdate()
''On Error GoTo Err_txt_日付FROM_AfterUpdate
''    Dim strY, strM  As String
''    Dim strMsg      As String
''
''    If Nz(Me.txt_日付FROM) = "" Then Exit Sub
''
''    strMsg = "計上年月に入力不正。"
''    If Len(Trim(Nz(Me.txt_日付FROM))) < 6 Then
''        MsgBox strMsg, vbInformation + vbOKOnly, strMsg_err
''        Exit Sub
''    End If
''
''    If Left(Me.txt_日付FROM, 4) <= 0 Or Right(Me.txt_日付FROM, 2) > 12 Or _
''            Right(Me.txt_日付FROM, 2) < 1 Then
''        MsgBox strMsg, vbInformation + vbOKOnly, strMsg_err
''        Exit Sub
''    End If
''
''    If Nz(Me.txt_日付FROM) <> "" Then
''        '振込期限日 = 計上年月の前月末日
''        Me.txt_Fkbi = Format(DateAdd("d", -1, Me.txt_日付FROM & "/01"), "yyyymmdd")
''    End If
''
''    Exit Sub
''
''Err_txt_日付FROM_AfterUpdate:
''    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
''    Exit Sub
''
''End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_日付FROM_LostFocus
'        CREATE_DATE      :2011/01/08 M.RYU
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'txt_日付FROM_LostFocusイベント
'**********************************************
Private Sub txt_日付FROM_LostFocus()
On Error GoTo Err_txt_日付FROM_AfterUpdate
    Dim strMsg      As String
    
    If Nz(Me.txt_日付FROM) = "" Then
        Me.txt_Fkbi = ""
        Exit Sub
    End If
    
    strMsg = "計上年月に入力不正。"
    If Len(Trim(Nz(Me.txt_日付FROM))) < 6 Then GoTo msgHandle
    If Left(Me.txt_日付FROM, 4) <= 0 Or Right(Me.txt_日付FROM, 2) > 12 Or _
            Right(Me.txt_日付FROM, 2) < 1 Then GoTo msgHandle
    
    If Nz(Me.txt_日付FROM) <> "" Then
        '振込期限日 = 計上年月の前月末日
        Me.txt_Fkbi = Format(DateAdd("d", -1, Me.txt_日付FROM & "/01"), "yyyymmdd")
    End If
       
    Exit Sub

msgHandle:
    MsgBox strMsg, vbInformation + vbOKOnly, strMsg_err
    Me.cmd_終了.SetFocus
    Me.txt_日付FROM.SetFocus
    Exit Sub
    
Err_txt_日付FROM_AfterUpdate:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    Exit Sub
End Sub

''Private Sub txt_日付FROM_AfterUpdate()        ''---------------------DELETE 2011/01/08 M.RYU
''
''    On Error GoTo Err_txt_日付FROM_AfterUpdate
''
''    '*-Nullチェック-*
''    If IsNull(Me.txt_日付FROM) Or Me.txt_日付FROM = "" Then
''        Exit Sub
''    End If
''
''    '*-Check_Date関数を使用-*
'''''''    Me.txt_日付ＦＲＯＭ = Check_Date(Me.txt_日付ＦＲＯＭ)     'DEL 2000/04/20
''    Me.txt_日付FROM = Check_Date(Me.txt_日付FROM & "01")           'ADD 2000/04/20
''
''    Exit Sub
''
''Err_txt_日付FROM_AfterUpdate:
''    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
''    Exit Sub
''
''End Sub
'''==============================================================================*
'''
'''        MODULE_NAME      :
'''        MODULE_ID        :txt_日付FROM_BeforeUpdate
'''        CREATE_DATE      :
''''----------------------------DELETE 2011/01/08 M.RYU
'''==============================================================================*
'''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'''**********************************************
'''txt_日付FROM_BeforeUpdateイベント
'''**********************************************
''Private Sub txt_日付FROM_BeforeUpdate(Cancel As Integer)
''    On Error GoTo Err_txt_日付FROM_BeforeUpdate
''
''    Dim vDATE As Variant
'''''''    vDATE = Nz(Me.txt_日付ＦＲＯＭ, "")     'DEL 2000/04/20
''    vDATE = Nz(Me.txt_日付FROM, "") & "01"       'ADD 2000/04/20
''
''    '*-Nullチェック-*
''    If vDATE = "" Then
''        Exit Sub
''    End If
''
''    '*-Check_Date関数を使用-*
''    vDATE = Check_Date(vDATE)
''    vDATE = CMCheck_Date(vDATE)
''
''    '*-入力チェック-*
''    Select Case CmBCheck_Date(vDATE)
''        Case 0
''
''        Case 2
''        '桁数ｴﾗｰ
''            MsgBox "桁数オーバーしています。", vbInformation, strMsg_err
''            Cancel = True
''            SendKeys "{ESC}"
''            GoTo Exit_txt_日付FROM_BeforeUpdate
''        Case 3, 9
''            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
''            Cancel = True
''            SendKeys "{ESC}"
''            GoTo Exit_txt_日付FROM_BeforeUpdate
''    End Select
''
''Exit_txt_日付FROM_BeforeUpdate:
''
''    Exit Sub
''
''Err_txt_日付FROM_BeforeUpdate:
''    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
''    Exit Sub
''
''End Sub
'''==============================================================================*
'''
'''        MODULE_NAME      :
'''        MODULE_ID        :txt_日付FROM_GotFocus
'''        CREATE_DATE      :
'''----------------------------DELETE 2011/01/08 M.RYU
'''==============================================================================*
'''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'''**********************************************
'''txt_日付FROM_GotFocusイベント
'''**********************************************
''Private Sub txt_日付FROM_GotFocus()
''
'''''''    Me.txt_日付ＦＲＯＭ = Format(Me.txt_日付ＦＲＯＭ, "YYYY/MM/DD")   'DEL 2000/04/20
''    Me.txt_日付FROM = Format(Me.txt_日付FROM, "YYYY/MM")                   'ADD 2000/04/20
''
''End Sub
'''==============================================================================*
'''
'''        MODULE_NAME      :
'''        MODULE_ID        :txt_日付FROM_KeyPress
'''        CREATE_DATE      :
'''----------------------------DELETE 2011/01/08 M.RYU
'''==============================================================================*
'''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'''**********************************************
'''txt_日付FROM_KeyPressイベント
'''**********************************************
''Private Sub txt_日付FROM_KeyPress(KeyAscii As Integer)
''
'''''''    Call KeyPress_For_Date(Me.txt_日付ＦＲＯＭ, KeyAscii, 10)   'DEL 2000/04/20
''    Call KeyPress_For_Date(Me.txt_日付FROM, KeyAscii, 7)             'ADD 2000/04/20
''
''End Sub
'''==============================================================================*
'''
'''        MODULE_NAME      :
'''        MODULE_ID        :txt_日付FROM_LostFocus
'''        CREATE_DATE      :
'''----------------------------DELETE 2011/01/08 M.RYU
'''==============================================================================*
'''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'''**********************************************
'''txt_日付FROM_LostFocusイベント
'''**********************************************
''Private Sub txt_日付FROM_LostFocus()
''
'''''''    Me.txt_日付ＦＲＯＭ = Format(Me.txt_日付ＦＲＯＭ, "yyyy/mm/dd")   'DEL 2000/04/20
''    Me.txt_日付FROM = Format(Me.txt_日付FROM, "YYYY/MM")                   'ADD 2000/04/20
''
''End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_END_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'**********************************************
'txt_End_Enterイベント
'**********************************************
Private Sub txt_END_Enter()
    'Me.cmd_印刷.SetFocus                                                       'DELETE 20051219 N.MIURA
    Me.cmd_プレビュー.SetFocus                                                  'INSERT 20051219 N.MIURA
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :KeyPress_For_Date
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓共通関数に入れたい・・・。
'**********************************************
' @f
'
' 機　能　 :キーの取消 日付のみ対応
' 返り値　 :Keyacii
' 引　数　 :ojCtrl   - コントロール
' 　　　    Keyascii - Keyascii
' 　　　　  inLen    - 入力する最大文字長
'
' 機能説明 :引数で指定したコントロールに入力したAsciiCodeを有効か無効か判断し
' 　　　　  返り値をKeyasciiに格納する
'
'**********************************************
Private Sub KeyPress_For_Date(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                     ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc("/")   ' 0～9迄、/ を許可する
                ''' 入力された文字列が引数の最大文字長に達した場合、入力したキーを無効にする
                If Len(.Text) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)
'-----------------------------------------------------------'
'  Check_Dateでは'198/03といった形が変換されるための補助検査
'-----------------------------------------------------------'
  
    Dim sYear As String
      
    CMCheck_Date = Null
      
    If Nz(vDATE, "") = "" Then
        Exit Function
    End If
      
    If IsDate(vDATE) = False Then
        Exit Function
    End If
      
    sYear = Format(vDATE, "yyyy")
    
    If CmLenb(sYear) <> 4 Then
        Exit Function
    End If
      
    CMCheck_Date = vDATE

End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmBCheck_Date
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer
'------------------------'
'  入金年月日のチェック
'------------------------'
  '引数
  'oText 対象項目
  'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
    
  '返り値は、
  '２：桁数が溢れている時
  '３：日付と判断できない時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
On Error GoTo Err_CmBCheck_Date
  
    Dim sYear  As String
    Dim intLen As Integer
    
    CmBCheck_Date = 2
    
    If IsMissing(ilen) = False Then
        intLen = Nz(CmNumeric(ilen), 0)
    Else
        intLen = 10
    End If
    
    oText = Nz(oText, "")
    
    '桁数チェック
    If CmLenb(oText) > intLen Then
        GoTo Exit_CmBCheck_Date
    End If
    
    CmBCheck_Date = 3
    
    If Nz(oText, "") = "" Then
        GoTo Exit_CmBCheck_Date
    End If
    
    '論理チェック
    If IsDate(oText) = False Then
        GoTo Exit_CmBCheck_Date
    End If
    
    sYear = Format(oText, "yyyy")
    
    If CmLenb(sYear) <> 4 Then
        GoTo Exit_CmBCheck_Date
    End If
    
    CmBCheck_Date = 0

Exit_CmBCheck_Date:
  Exit Function
  
Err_CmBCheck_Date:
  MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
  Err.Clear
  CmBCheck_Date = 9
  GoTo Exit_CmBCheck_Date
  
End Function
'==============================================================================*
'
'        MODULE_NAME      :社印画像格納パス取得
'        MODULE_ID        :fncGetStampPath
'        CREATE_DATE      :2008/06/16 N.Miura
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetStampPath() As String
On Error GoTo err_rtn

    Dim strWhere            As String
    
    fncGetStampPath = ""
    
    strWhere = " INTIF_PROGB = '" & Me.NAME & "' AND " & _
               " INTIF_RECFB = '" & Me.txt_CAMPC & "' "
    
    fncGetStampPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere), "")
    GoTo Exit_rtn

err_rtn:
    Call Err.Raise(Err.Number, "fncGetStampPath" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

Exit_rtn:

End Function

'2017-11-27 add start k.sato
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getCardImpDate() As String

    Dim objDb       As Database
    Dim ObjQDF      As QueryDef
    Dim objRst      As Recordset
    Dim strConnect  As String               'Connect情報
    Dim strSQL      As String
    Dim strError    As String
    Dim errObj      As Error
    
    
    On Error GoTo Err_getCardImpDate

    getCardImpDate = ""

    Set objDb = CurrentDb

    'データベース名取得：KONTENA
    strConnect = DLookup("CONT_SERVER_PATH", "dbo_CONT_MAST", "CONT_KEY=1")
    If IsNull(strConnect) Then
        MsgBox "ＳＱＬＳｅｒｖｅｒのＰＡＳＳが見つかりません。", vbInformation + vbOKOnly, strMsg_err
        GoTo Exit_getCardImpDate
    End If

    objDb.QueryTimeout = 1200
    Set ObjQDF = objDb.CreateQueryDef("")
    ObjQDF.Connect = strConnect
    ObjQDF.ReturnsRecords = True

    'strSql = " SELECT * FROM CLOG_TRAN ORDER BY CLOGT_INSED, CLOGT_INSEJ "             'DELETE 2018/12/15 N.IMAI
    strSQL = " SELECT * FROM CLOG_TRAN ORDER BY CLOGT_INSED DESC, CLOGT_INSEJ DESC "    'INSERT 2018/12/15 N.IMAI

    ObjQDF.sql = strSQL
    Set objRst = ObjQDF.OpenRecordset(dbOpenSnapshot)

    If objRst.EOF Then
        getCardImpDate = ""
        GoTo Exit_getCardImpDate
    Else
        getCardImpDate = objRst.Fields("CLOGT_IMPD").VALUE
    End If

    objDb.Close
    Set objDb = Nothing
    
Exit_getCardImpDate:
    If Not objDb Is Nothing Then
        objDb.Close
        Set objDb = Nothing
    End If
    
    Exit Function

Err_getCardImpDate:
    Errors.Refresh
    strError = "エラーが発生しました。" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    For Each errObj In DBEngine.Errors
        If Format$(errObj.Number) <> 3325 Then
            strError = strError & Format$(errObj.Number)
            strError = strError & " : " & errObj.Description
            strError = strError & " (" & errObj.Source & ") . "
            strError = strError & Chr$(13) & Chr$(10)
        Else
            'レコードが返されなかったのエラーだから、正常
            getCardImpDate = ""
            GoTo Exit_getCardImpDate
        End If
    Next
    MsgBox strError, vbCritical + vbOKOnly, strMsg_err
    GoTo Exit_getCardImpDate

End Function
'2017-11-27 add end k.sato
'==============================================================================*
'
'        MODULE_NAME      :出力した情報をワークに保存
'        MODULE_ID        :Make_Output_List
'        CREATE_DATE      :2017/12/01 N.IMAI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Make_Output_List(strDate As String, strTime As String) As Boolean

    Dim objCon          As Object
    Dim objRst          As Object
    Dim strError        As String               'エラーメッセージ用
    Dim errObj          As Error
    Dim strSQL          As String
    'Dim strBumonCode    As String                                              'DELETE 2023/02/24 N.IMAI

    On Error GoTo ErrorHandler

    Make_Output_List = False

    '部門コード取得
    'strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1"))  'DELETE 2023/02/24 N.IMAI
    
    '出力データ取得
    Set objRst = CurrentDb.OpenRecordset(strOutputHeader, dbOpenSnapshot)
    
    On Error GoTo ErrorHandler2

    Set objCon = ADODB_Connection(FstrBumonCode)
    On Error GoTo ErrorHandler3
    
    'ワークを削除
    strSQL = ""
    strSQL = strSQL & "DELETE RKS240_WORK_SAVE "
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "   DISP_BUMOC = '" & FstrBumonCode & "' "
    strSQL = strSQL & " AND"
    strSQL = strSQL & "   DISP_HAKOD = '" & txt_Hkobi & "' "
    Call ADODB_Execute(strSQL, objCon)

    'ワークへ保存
    Do Until objRst.EOF
        strSQL = ""
        strSQL = strSQL & "INSERT INTO RKS240_WORK_SAVE "
        strSQL = strSQL & "(DISP_BUMOC "
        strSQL = strSQL & ",DISP_HAKOD "
        strSQL = strSQL & ",DISP_KEIJD "
        strSQL = strSQL & ",RKS170_KCODE "
        strSQL = strSQL & ",RKS170_KINGAK "
        strSQL = strSQL & ",USER_CODE "
        strSQL = strSQL & ",USER_NAME "
        strSQL = strSQL & ",USER_KANA "
        strSQL = strSQL & ",USER_TANM "
        strSQL = strSQL & ",USER_TAKA "
        strSQL = strSQL & ",USER_YUBINO "
        strSQL = strSQL & ",USER_ADR_1 "
        strSQL = strSQL & ",USER_ADR_2 "
        strSQL = strSQL & ",USER_ADR_3 "
        strSQL = strSQL & ",USER_TEL "
        strSQL = strSQL & ",USER_KEITAI "
        strSQL = strSQL & ",USER_FAX "
        strSQL = strSQL & ",USER_SDAY "
        strSQL = strSQL & ",USER_KKBN "
        strSQL = strSQL & ",USER_SKBN "
        strSQL = strSQL & ",USER_BIKO "
        strSQL = strSQL & ",USER_UPDATE "
        strSQL = strSQL & ",RKS170_PAGE "
        strSQL = strSQL & ",USER_RNAME "
        strSQL = strSQL & ",USER_RPOST "
        strSQL = strSQL & ",USER_RADR_1 "
        strSQL = strSQL & ",USER_RADR_2 "
        strSQL = strSQL & ",USER_RADR_3 "
        strSQL = strSQL & ",KOUZM_KINYC "
        strSQL = strSQL & ",KOUZM_KINYN "
        strSQL = strSQL & ",KOUZM_SHITC "
        strSQL = strSQL & ",KOUZM_SHITN "
        strSQL = strSQL & ",KOUZM_YOKII "
        strSQL = strSQL & ",KOUZM_YOKIN "
        strSQL = strSQL & ",KOUZM_KOUZB "
        strSQL = strSQL & ",WORK_INSED "
        strSQL = strSQL & ",WORK_INSEJ "
        strSQL = strSQL & ") "
        strSQL = strSQL & "VALUES "
        strSQL = strSQL & "('" & FstrBumonCode & "'"
        strSQL = strSQL & ",'" & Nz(txt_Hkobi, "") & "'"
        strSQL = strSQL & ",'" & Nz(txt_日付FROM, "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("RKS170_KCODE"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("RKS170_KINGAKの合計"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_CODE"), "") & "'"
'        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_NAME"), "") & "'"                      'DELETE 2018/07/13 N.IMAI
'        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_KANA"), "") & "'"                      'DELETE 2018/07/13 N.IMAI
        strSQL = strSQL & ",'" & Replace(Nz(objRst.Fields("USER_NAME"), ""), "'", "''") & "'"   'INSERT 2018/07/13 N.IMAI
        strSQL = strSQL & ",'" & Replace(Nz(objRst.Fields("USER_KANA"), ""), "'", "''") & "'"   'INSERT 2018/07/13 N.IMAI
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_TANM"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_TAKA"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_YUBINO"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_ADR_1"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_ADR_2"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_ADR_3"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_TEL"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_KEITAI"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_FAX"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_SDAY"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_KKBN"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_SKBN"), "") & "'"
'        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_BIKO"), "") & "'"
        strSQL = strSQL & ",'" & Replace(Nz(objRst.Fields("USER_BIKO"), ""), "'", "''") & "'"   'INSERT 2020/04/06 N.IMAI
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_UPDATE"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("RKS170_PAGE"), "") & "'"
'        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_RNAME"), "") & "'"
        strSQL = strSQL & ",'" & Replace(Nz(objRst.Fields("USER_RNAME"), ""), "'", "''") & "'"  'INSERT 2020/04/06 N.IMAI
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_RPOST"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_RADR_1"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_RADR_2"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("USER_RADR_3"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("KOUZM_KINYC"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("KOUZM_KINYN"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("KOUZM_SHITC"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("KOUZM_SHITN"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("KOUZM_YOKII"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("KOUZM_YOKIN"), "") & "'"
        strSQL = strSQL & ",'" & Nz(objRst.Fields("KOUZM_KOUZB"), "") & "'"
        strSQL = strSQL & ",'" & strDate & "'"
        strSQL = strSQL & ",'" & strTime & "'"
        strSQL = strSQL & ") "
        Call ADODB_Execute(strSQL, objCon)
        objRst.MoveNext
    Loop

ErrorHandler3:
    objCon.Close
    Set objCon = Nothing

ErrorHandler2:
    objRst.Close
    Set objRst = Nothing

ExitHandler:
    Exit Function

ErrorHandler:
    Errors.Refresh
    strError = "ワーク生成中にエラーが発生しました。" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    For Each errObj In DBEngine.Errors
        If Format$(errObj.Number) <> 3325 Then
            strError = strError & Format$(errObj.Number)
            strError = strError & " : " & errObj.Description
            strError = strError & " (" & errObj.Source & ") . "
            strError = strError & Chr$(13) & Chr$(10)
        Else
            'レコードが返されなかったのエラーだから、正常
            Make_Output_List = True
            GoTo ExitHandler
        End If
    Next
    MsgBox strError, vbCritical + vbOKOnly, strMsg_err

End Function

'==============================================================================*
'
'        MODULE_NAME      :出力したヘッダ情報を取得するSQL
'        MODULE_ID        :strOutputHeader
'        CREATE_DATE      :2017/12/01 N.IMAI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function strOutputHeader() As String
    Dim strSQL  As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "    請求書ヘッダーＷＫ.RKS170_KCODE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.RKS170_KINGAKの合計"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_CODE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_NAME"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_KANA"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_TANM"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_TAKA"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_YUBINO"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_ADR_1"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_ADR_2"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_ADR_3"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_TEL"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_KEITAI"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_FAX"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_SDAY"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_KKBN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_SKBN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_BIKO"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_UPDATE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.RKS170_PAGE"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RNAME"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RPOST"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RADR_1"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RADR_2"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.USER_RADR_3"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_KINYC"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_KINYN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_SHITC"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_SHITN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_YOKII"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_YOKIN"
    strSQL = strSQL & "  , 請求書ヘッダーＷＫ.KOUZM_KOUZB "
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "    請求書ヘッダーＷＫ"
    strSQL = strSQL & "  , FQS240 "
    strSQL = strSQL & " ORDER BY 請求書ヘッダーＷＫ.USER_CODE "

    strOutputHeader = strSQL

End Function
'==============================================================================*
'
'        MODULE_NAME      :該当フォルダ内の最新更新日付ファイルを取得
'        MODULE_ID        :FncGetNewFile
'        PARAMETER        : I/O     strFilePath     フォルダPATH
'                           I/      strFilePrefix   ファイル名接頭子
'        CREATE_DATE      :2018/06/01 TFC
'        UPDATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function FncGetNewFile(strFilePath As String, strFilePrefix As String) As String

    Dim strFilename As String
    Dim tmpFileName As String
    Dim time As Date
    Dim pretime As Date
     
    tmpFileName = Dir(strFilePath & strFilePrefix & "*.csv")
    
    Do Until tmpFileName = ""
        pretime = FileDateTime(strFilePath & tmpFileName)
        If pretime > time Then
            time = pretime
            strFilename = tmpFileName
        End If
        tmpFileName = Dir()
    Loop
    
    FncGetNewFile = strFilename

End Function
'==============================================================================*
'
'        MODULE_NAME      :口座マスタ(KOUZ_MAST)の情報をセット
'        MODULE_ID        :subKOUZ_MAST_Set
'        CREATE_DATE      :2023/02/24   N.IMAI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKOUZ_MAST_Set()

    Dim objCon          As Object
    Dim objRst          As Object
    Dim objWk           As Recordset
    Dim strSQL          As String
    
    On Error GoTo ErrorHandler
    
    Set objCon = ADODB_Connection()
    On Error GoTo ErrorHandler2
    
    Set objWk = CurrentDb.OpenRecordset("請求書ヘッダーＷＫ", dbOpenDynaset)

    Do Until objWk.EOF
        strSQL = ""
        strSQL = strSQL & vbCrLf & "SELECT * "
        strSQL = strSQL & vbCrLf & "  FROM KOUZ_MAST "
        strSQL = strSQL & vbCrLf & " INNER JOIN CAMP_MAST "
        strSQL = strSQL & vbCrLf & "    ON CAMPM_CAMPC = KOUZM_CAMPC "
        strSQL = strSQL & vbCrLf & "WHERE KOUZM_BUMOC = '" & FstrBumonCode & "' "
        strSQL = strSQL & vbCrLf & "  AND KOUZM_KOKYC = '" & Format(objWk.Fields("USER_CODE").VALUE, "000000") & "'"
        Set objRst = ADODB_Recordset(strSQL, objCon)
        On Error GoTo ErrorHandler3
        If objRst.EOF = False Then
            objWk.Edit
            objWk.Fields("KOUZM_KINYC").VALUE = objRst.Fields("KOUZM_KINYC").VALUE
            objWk.Fields("KOUZM_SHITC").VALUE = objRst.Fields("KOUZM_SHITC").VALUE
            objWk.Fields("KOUZM_YOKII").VALUE = objRst.Fields("KOUZM_YOKII").VALUE
            objWk.Fields("KOUZM_KOUZB").VALUE = objRst.Fields("KOUZM_KOUZB").VALUE
            objWk.Fields("KOUZM_KOUZN").VALUE = objRst.Fields("CAMPM_CAMPN").VALUE
            objWk.UPDATE
        End If
        objWk.MoveNext
    Loop

ErrorHandler3:
    objRst.Close
    Set objRst = Nothing

ErrorHandler2:
    objCon.Close
    Set objCon = Nothing

ErrorHandler:
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "subKOUZ_MAST_Set" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub
'****************************  ended or program ********************************
