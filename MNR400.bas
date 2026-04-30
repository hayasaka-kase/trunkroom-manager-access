Attribute VB_Name = "MNR400"
'****************************  strat or program ********************************
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：  サービス内容夜間設定
'   プログラムＩＤ　：　MNR400
'   作　成　日　　　：  2017/12/01
'   作　成　者　　　：  Y.SUZUKI
'   Ver             ：  1.0
'**********************************************
'修正履歴
'   修　正　日　　　：
'   修　正　者　　　：
'   修　正　内　容　：
'   Ver             ：
'**********************************************
Option Compare Database
Option Explicit

'**********************************************
'定数宣言
'**********************************************
Private Const strPROG_ID    As String = "MNR400"

'**********************************************
'変数宣言
'**********************************************
Private strUSER_ID          As String       'ユーザーコード

Private datUpdate           As Date         'システム日付(date型)
Private strDate             As String       'システム日付(文字列)
Private strTime             As String       'システム時刻(文字列)

Private lngCnt              As Long         '処理件数(割引終了)
Private lngCnt2             As Long         '処理件数(実施期間到来)

'==============================================================================*
'   MODULE_NAME     :   サービス内容夜間設定バッチ
'   Parameters      :   strBUMOC      部門コード
'   Return Value    :   True：正常終了、False：異常終了
'   CREATE_DATE     :   2017/12/01
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MNR400Main(ByVal strBUMOC As String) As Boolean
    
    Dim objCon              As Object
    Dim blnRTN              As Boolean
    
    On Error GoTo Exception
    
    '*************************************************
    '初期処理
    '*************************************************
    'カウンタ初期化
    lngCnt = 0
    lngCnt2 = 0

    '処理開始ログを出力
    Call MSZZ003_M00(strPROG_ID, "0", "")
    
    'システム日時を取得
    datUpdate = DATE                                        'システム日付
    strDate = Format(datUpdate, "yyyymmdd")                 'システム日付(yyyymmdd型式)
    strTime = Format(Now(), "hhmmss")                       'システム時刻(hhmmdd型式)
    
    '各種データ取得
    strUSER_ID = LsGetUserName()                            'ユーザーID
    
    'SQLサーバーに接続
    Set objCon = MSZZ025.ADODB_Connection(strBUMOC)
    
    '*************************************************
    'メイン処理
    '*************************************************
    'トランザクション開始
    objCon.BeginTrans
    
    '割引終了更新
    blnRTN = upd_Waribiki_Shuryou(objCon)
    If Not blnRTN Then
        GoTo Exit_Function
    End If
    
    '実施期間到来更新
    blnRTN = upd_Jisshi_Kikan(objCon)
    If Not blnRTN Then
        GoTo Exit_Function
    End If
    
    '*************************************************
    '終了処理
    '*************************************************
    objCon.CommitTrans
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    
    '各処理件数ログ出力方法
    Call MSZZ003_M00(strPROG_ID, "8", "割引終了更新件数：" & Format(lngCnt))
    Call MSZZ003_M00(strPROG_ID, "8", "割引開始登録件数：" & Format(lngCnt2))
    
    '終了ログ出力方法
    Call MSZZ003_M00(strPROG_ID, "1", "")
    
    MNR400Main = True
    
    Exit Function

Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "MNR400Main" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
Exit_Function:
    'ロールバック
    objCon.RollbackTrans
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    
End Function

'==============================================================================*
'   MODULE_NAME     :   割引終了更新処理
'   Parameters      :   objCon      接続情報
'   Return Value    :   boolean     TRUE=正常終了、FALSE=異常終了
'   CREATE_DATE     :   2017/12/01
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function upd_Waribiki_Shuryou(objCon As Object) As Boolean
    
    Dim strDML                  As String               'SQL編集
    Dim strYcode                As String
    Dim strYCODE_BK             As String
    
    Dim objRLDN                 As Object
    Dim objYARD                 As Object
    
    On Error GoTo Exception
    
    '*************************************************
    '無効化または実施期間が過ぎた自動処理対象を読込
    '*************************************************
    'SQL編集 - 検索
        strDML = "SELECT RLDNT_YCODE, RLDNT_NO, RLDNT_FROM, RLDNT_CTLFLG, RLDNT_UPDAD, RLDNT_UPDAJ, RLDNT_UPDPB, RLDNT_UPDUB" _
               & " FROM RLDN_TRAN" _
               & " WHERE RLDNT_GNKBN  = '1' AND RLDNT_CTLFLG = '1' AND RLDNT_ORDER = '1' AND ( RLDNT_ENABLE = '0' OR CONVERT(NVARCHAR, GETDATE(), 112) > RLDNT_TO )" _
               & " ORDER BY RLDNT_YCODE"
    Set objRLDN = MSZZ025.ADODB_Recordset(strDML, objCon, adoReadWrite)             'adoReadOnly, adoAppendOnly, adoReadWrite
    
    '*************************************************
    '割引終了更新
    '*************************************************
    Do Until objRLDN.EOF
        'ヤードコード取得
        strYcode = objRLDN.Fields("RLDNT_YCODE")
        
        '*************************************************
        'ヤードマスタのサービス内容をクリアする
        '*************************************************
        If strYcode <> strYCODE_BK Then
            'SQL編集 - 検索
            strDML = "SELECT * FROM YARD_MAST WHERE YARD_CODE = '" & strYcode & "'"
            Set objYARD = MSZZ025.ADODB_Recordset(strDML, objCon, adoReadWrite)         'adoReadOnly, adoAppendOnly, adoReadWrite
            
            'ヤードマスタ更新
            With objYARD
                .Fields("YARD_SEV1N") = vbNullString            'サービス内容１
                .Fields("YARD_SEV2N") = vbNullString            'サービス内容２
                .Fields("YARD_SEV3N") = vbNullString            'サービス内容３
                .Fields("YARD_ENDEN") = vbNullString            'サービス期間
                .Fields("YARD_UPDATE") = datUpdate              '更新日
                
                .UPDATE
            End With
            'レコードセットクリア
            If Not objYARD Is Nothing Then objYARD.Close: Set objYARD = Nothing
            
            strYCODE_BK = strYcode
        End If
        
        '*************************************************
        '自動処理済みの更新を行う
        '*************************************************
        'レンタル物件割引更新
        objRLDN.Fields("RLDNT_CTLFLG") = "0"                '制御フラグ
        objRLDN.Fields("RLDNT_UPDAD") = strDate             '更新日付
        objRLDN.Fields("RLDNT_UPDAJ") = strTime             '更新時刻
        objRLDN.Fields("RLDNT_UPDPB") = strPROG_ID          '更新プログラムID
        objRLDN.Fields("RLDNT_UPDUB") = strUSER_ID          '更新ユーザーID
        
        objRLDN.UPDATE
        
        '次のデータ
        objRLDN.MoveNext
        
        '処理件数カウント
        lngCnt = lngCnt + 1
    Loop
    
    'レコードセットクリア
    If Not objRLDN Is Nothing Then objRLDN.Close: Set objRLDN = Nothing
    
    '正常終了
    upd_Waribiki_Shuryou = True
    Exit Function
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "upd_Waribiki_Shuryou" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
    'レコードセットクリア
    If Not objRLDN Is Nothing Then objRLDN.Close: Set objRLDN = Nothing
    If Not objYARD Is Nothing Then objYARD.Close: Set objYARD = Nothing
    
End Function

'==============================================================================*
'   MODULE_NAME     :   実施期間到来更新処理
'   Parameters      :   objCon      接続情報
'   Return Value    :   boolean     TRUE=正常終了、FALSE=異常終了
'   CREATE_DATE     :   2017/12/01
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function upd_Jisshi_Kikan(objCon As Object) As Boolean
    
    Dim strDML                  As String               'SQL編集
    
    Dim strYcode                As String               'ヤードコード
    Dim strSEV1N                As String               'サービス内容１
    Dim strSEV2N                As String               'サービス内容２
    Dim strSEV3N                As String               'サービス内容３
    Dim strENDEN                As String               'サービス期間
    Dim strYCODE_BK             As String
    
    Dim objRLDN                 As Object
    Dim objYARD                 As Object
    
    On Error GoTo Exception
    
    '*************************************************
    '無効化または実施期間が過ぎた自動処理対象を読込
    '*************************************************
    'SQL編集 - 検索
    strDML = "SELECT RLDNT_YCODE, RLDNT_NO, RLDNT_FROM, RLDNT_SEV1N, RLDNT_SEV2N, RLDNT_SEV3N, RLDNT_ENDEN, RLDNT_UPDAD, RLDNT_UPDAJ, RLDNT_UPDPB, RLDNT_UPDUB" _
           & " FROM RLDN_TRAN" _
           & " WHERE RLDNT_GNKBN  = '1' AND RLDNT_ENABLE = '1' AND RLDNT_CTLFLG = '1' AND RLDNT_ORDER  = '1' AND CONVERT(NVARCHAR, GETDATE(), 112) BETWEEN RLDNT_FROM AND RLDNT_TO"
    Set objRLDN = MSZZ025.ADODB_Recordset(strDML, objCon, adoReadWrite)             'adoReadOnly, adoAppendOnly, adoReadWrite
    
    '*************************************************
    '実施期間到来更新
    '*************************************************
    Do Until objRLDN.EOF
        '取得
        strYcode = objRLDN.Fields("RLDNT_YCODE")            'ヤードコード
        strSEV1N = objRLDN.Fields("RLDNT_SEV1N")            'サービス内容１
        strSEV2N = objRLDN.Fields("RLDNT_SEV2N")            'サービス内容２
        strSEV3N = objRLDN.Fields("RLDNT_SEV3N")            'サービス内容３
        strENDEN = objRLDN.Fields("RLDNT_ENDEN")            'サービス期間
        
        '*************************************************
        'ヤードマスタのサービス内容をクリアする
        '*************************************************
        'SQL編集 - 検索
        strDML = "SELECT * FROM YARD_MAST WHERE YARD_CODE = '" & strYcode & "'"
        Set objYARD = MSZZ025.ADODB_Recordset(strDML, objCon, adoReadWrite)         'adoReadOnly, adoAppendOnly, adoReadWrite
        
        'ヤードマスタ更新
        With objYARD
            .Fields("YARD_SEV1N") = strSEV1N                'サービス内容１
            .Fields("YARD_SEV2N") = strSEV2N                'サービス内容２
            .Fields("YARD_SEV3N") = strSEV3N                'サービス内容３
            .Fields("YARD_ENDEN") = strENDEN                'サービス期間
            .Fields("YARD_UPDATE") = datUpdate              '更新日
            
            .UPDATE
        End With
        'レコードセットクリア
        If Not objYARD Is Nothing Then objYARD.Close: Set objYARD = Nothing
        
        '*************************************************
        '自動処理済みの更新を行う
        '*************************************************
        'レンタル物件割引更新
        objRLDN.Fields("RLDNT_UPDAD") = strDate             '更新日付
        objRLDN.Fields("RLDNT_UPDAJ") = strTime             '更新時刻
        objRLDN.Fields("RLDNT_UPDPB") = strPROG_ID          '更新プログラムID
        objRLDN.Fields("RLDNT_UPDUB") = strUSER_ID          '更新ユーザーID
        
        objRLDN.UPDATE
        
        '次のデータ
        objRLDN.MoveNext
        
        '処理件数カウント
        lngCnt2 = lngCnt2 + 1
    Loop
    
    'レコードセットクリア
    If Not objYARD Is Nothing Then objYARD.Close: Set objYARD = Nothing
    
    '正常終了
    upd_Jisshi_Kikan = True
    Exit Function
    
Exception:
    'メッセージ・ログ出力
    Call Err.Raise(Err.Number, "upd_Jisshi_Kikan" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
    'レコードセットクリア
    If Not objRLDN Is Nothing Then objRLDN.Close: Set objRLDN = Nothing
    If Not objYARD Is Nothing Then objYARD.Close: Set objYARD = Nothing
    
End Function



