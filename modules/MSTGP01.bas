Attribute VB_Name = "MSTGP01"
'****************************  strat of program ********************************
'==============================================================================*
'
'       SYSTEM_NAME     : コンテナ統合ＤＢシステム
'       SUB_SYSTEM_NAME : 統合ＤＢ
'
'       PROGRAM_NAME    : コンテナ統合ＤＢ更新
'       PROGRAM_ID      : MSTGP01
'       PROGRAM_KBN     : MODULE
'
'       CREATE          : 2011/02/06
'       CERATER         : K.ISHIZAKA
'       Ver             : 0.0
'
'       UPDATE          : 2011/03/02
'       UPDATER         : K.ISHZIAKA
'       Ver             : 0.1
'                         ログ出力
'
'       UPDATE          : 2011/03/08
'       UPDATER         : K.ISHZIAKA
'       Ver             : 0.2
'                         営業中のヤードのみとする
'                         過去の割引は対象データに含めない
'
'       UPDATE          : 2011/03/10
'       UPDATER         : K.ISHZIAKA
'       Ver             : 0.3
'                         コンテナ(バイク)の価格テーブルが存在しない場合
'                         コンテナの価格で作成する
'
'       UPDATE          : 2011/03/12
'       UPDATER         : K.ISHZIAKA
'       Ver             : 0.4
'                         接続部門を受取れるようにする
'
'       UPDATE          : 2011/03/15
'       UPDATER         : K.ISHZIAKA
'       Ver             : 0.5
'                         用途＝39:バイク(コンテナ)に対応する
'
'       UPDATE          : 2011/06/18
'       UPDATER         : M.HONDA
'       Ver             : 0.6
'                         1.TOGO_YARD_MAST作成にADDR_TABLとの結合条件を
'                           住所の名称ではなく郵便番号で結合させるように修正。
'                         2.差額金額が入っている件数を取得するように修正
'
'       UPDATE          : 2011/09/06
'       UPDATER         : M.HONDA
'       Ver             : 0.7
'                         1.部屋毎の奥・幅・高さを最小値ではなく
'                           幅の最小値に合わせた奥・高さを取得するように修正。
'
'       UPDATE          : 2011/09/30
'       UPDATER         : M.HONDA
'       Ver             : 0.8
'                         1.割引データ作成時に事務手数料割引は取得しないように修正。
'                         2.終了している割引データは作成しないように修正。
'                         3.円指定・n円引きの求め方が逆の為修正。
'
'       UPDATE          : 2011/10/06
'       UPDATER         : M.HONDA
'       Ver             : 0.9
'                         1.3535.comのお客様サイト用のログイン情報を作成
'
'       UPDATE          : 2012/06/26
'       UPDATER         : K.ISHZIAKA
'       Ver             : 1.0
'                         Webの空き数を減らす
'
'       UPDATE          : 2013/06/13
'       UPDATER         : K.ISHZIAKA
'       Ver             : 1.1
'                         受付中のデータも対象にする
'
'       UPDATE          : 2013/04/06
'       UPDATER         : M.HONDA
'       Ver             : 1.2
'                         コンテナ（バイク）価格表作成時に雑費コード、金額を０へ修正
'
'       UPDATE          : 2013/11/03
'       UPDATER         : K.ISHZIAKA
'       Ver             : 1.3
'                         TOGO_YARD_MASTの項目追加
'
'       UPDATE          : 2014/10/06
'       UPDATER         : M.HONDA
'       Ver             : 1.4
'                         TOGO_YARD_MASTの項目追加
'
'       UPDATE          : 2014/11/12
'       UPDATER         : M.HONDA
'       Ver             : 1.5
'                         TOGO_YARD_MASTの項目追加
'
'       UPDATE          : 2014/12/04
'       UPDATER         : K.ISHZIAKA
'       Ver             : 1.6
'                         TOGO_RLDN_TRANの項目追加
'                         事務手数料割引、保証委託料割引に対応
'
'       UPDATE          : 2014/12/12
'       UPDATER         : K.ISHZIAKA
'       Ver             : 1.7
'                         Webも考慮した統合ＤＢの空き件数を取得する関数追加
'
'       UPDATE          : 2015/01/26
'       UPDATER         : K.ISHZIAKA
'       Ver             : 1.8
'                         保証委託料割引の算出方法改定
'                           毎月の割引額（電話以外）と補償加入費を含めて算出する
'
'       UPDATE          : 2015/01/30
'       UPDATER         : K.ISHZIAKA
'       Ver             : 1.9
'                         割引後額、割引額を算出する
'
'       UPDATE          : 2015/01/31
'       UPDATER         : K.ISHZIAKA
'       Ver             : 2.0
'
'                         保証委託料の計算ミス
'       UPDATE          : 2015/02/12
'       UPDATER         : M.HONDA
'       Ver             : 2.1
'                         SPAC_TABLをTOGO_SPAC_TABLへ修正
'
'       UPDATE          : 2016/08/17
'       UPDATER         : M.HONDA
'       Ver             : 2.2
'                         部屋指定区分を追加
'
'       UPDATE          : 2016/12/14
'       UPDATER         : M.HONDA
'       Ver             : 2.3
'                         割引で10単位の切り捨てを廃止
'
'       UPDATE          : 2018/01/14
'       UPDATER         : M.HONDA
'       Ver             : 2.4
'                         ネット契約割引を追加
'                         事務手数料割引/保証委託料割引をコメントアウト
'                         ヤードマスタにカラムを追加
'
'       UPDATE          : 2018/07/19
'       UPDATER         : K.ISHZIAKA
'       Ver             : 2.5
'                         空き状況に階数追加
'
'       UPDATE          : 2018/08/22
'       UPDATER         : K.ISHZIAKA
'       Ver             : 2.6
'                         ヤードマスタのセキュリティ、換気を追加
'                         価格表の雑費２、３を追加。初回、毎月それぞれ
'
'       UPDATE          : 2018/11/08
'       UPDATER         : K.ISHZIAKA
'       Ver             : 2.7
'                         空き状況に貸止数を追加
'                         ヤードマスタに貸止開始日時、貸止終了日時を追加
'
'       UPDATE          : 2018/12/20
'       UPDATER         : K.ISHZIAKA
'       Ver             : 2.8
'                         空き状況貸止数の間違い
'
'       UPDATE          : 2018/12/22
'       UPDATER         : K.ISHZIAKA
'       Ver             : 2.9
'                       　※Ver2.8のコードがあると見辛いのでなかったことにしました。
'                         空き状況貸止数の間違い
'
'       UPDATE          : 2019/01/28
'       UPDATER         : tajima
'       Ver             : 3.0
'                       　部門Hのネット割引を全て1080とする※トランクは変更なし
'
'       UPDATE          : 2019/04/05
'       UPDATER         : Y_SUZUKI EGL
'       Ver             : 3.1
'                       　TOGO_YARD_MASTの項目追加
'
'       UPDATE          : 2019/08/08
'       UPDATER         : M.HONDA
'       Ver             : 3.2
'                       　消費税対応 TOGO_PRIC_TABLに項目追加
'
'       UPDATE          : 2019/08/17
'       UPDATER         : Y_SUZUKI EGL
'       Ver             : 3.3
'                       　ヤードマスタの新規項目「集客契約有無」を、統合ヤードマスタに反映させる
'
'       UPDATE          : 2019/11/08
'       UPDATER         : K.SATO
'       Ver             : 3.3
'                       　賃貸クレカ対応 wase3535に渡すマスタLOGIN、KYKAKに賃貸部門追加(F,5,R)
'
'       UPDATE          : 2020/03/11
'       UPDATER         : M.HONDA
'       Ver             : 3.4
'                       　空き作る際に撤去コンテナは条件から外す
'
'       UPDATE          : 2020/06/01
'       UPDATER         : M.HONDA
'       Ver             : 3.5
'                       　コンテナ（バイクを）部屋単位の条件に入れる。
'
'       UPDATE          : 2020/09/05
'       UPDATER         : K.ISHIZAKA
'       Ver             : 3.4
'                       　検索キーワード追加
'
'       UPDATE          : 2020/11/20
'       UPDATER         : K.ISHIZAKA
'       Ver             : 4.0
'                       　インターネット予約バッチ、ご紹介バッチ、KOMS2002、MyPocetをマージしました。
'
'       UPDATE          : 2020/11/30
'       UPDATER         : K.KINEBUCHI
'       Ver             : 4.1
'                       　サイズが1畳以上のネット割引額を2200→3300に変更する
'                       　※価格は3300円以下にならないように設定してもらった
'
'       UPDATE          : 2021/03/24
'       UPDATER         : N.IMAI
'       Ver             : 4.2
'                       　2021/03/26 19:00:00 から 2021/04/01 09:00:00 のネット割引額を0にする
'
'       UPDATE          : 2021/06/16
'       UPDATER         : N.IMAI
'       Ver             : 4.3
'                       　事務手数料5500円固定→0円固定
'
'       UPDATE          : 2021/10/29
'       UPDATER         : N.IMAI
'       Ver             : 4.4
'                       　CARG_FILEの条件変更
'
'       UPDATE          : 2022/01/07
'       UPDATER         : N.IMAI
'       Ver             : 4.5
'                       　TOSO_SPAC_TABL更新時に部屋指定の場合は「空き状況（サイズ別）」の更新をやめる
'                       　ご紹介バッチのMSTGP01モジュールを利用（KOMSは古いままだった）
'
'       UPDATE          : 2026/04/22
'       UPDATER         : N.IMAI
'       Ver             : 4.6
'                       　更新系はTOSO_SPAC_TABL⇒TOGO_SPAC_TABL_Jとする
'
'
'       ※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※
'       ※
'       ※  注  意
'       ※
'       ※  ↓に変更 : このモジュールはKOMS_2002でも利用しているので、server11で修正した場合は注意して下さい。
'       ※
'       ※  3535リニューアルに伴ってバッチ系とは別モジュール管理とします。ご紹介バッチと統合禁止
'       ※
'       ※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※
'
'==============================================================================*
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID       As String = "MSTGP01"

'統合ＤＢ接続部門コード
Private Const C_TOGO_CONNENCT_BUMOC As String = "TOGO"

'用途区分のグループ
Private Const C_USAGE_GRP_COUNT As String = "0,1,3,4,10,30,31,32"   'サイズ別集計
Private Const C_USAGE_GRP_KBIKE As String = "33"                    'コンテナ(バイク）
'Private Const C_USAGE_GRP_ALONE As String = "0,1,3,4,10,30,31,32"   '部屋別
Private Const C_USAGE_GRP_ALONE As String = "0,1,3,4,10,30,31,32,33"   '部屋別 '2020/06/01 M.HONDA


Private Const C_WORK_NYAR1  As String = "WORK_NYAR_MAST1"
Private Const C_WORK_NYAR2  As String = "WORK_NYAR_MAST2"

Private Const C_2kSEV_LIMIT     As Long = 5000                                  'INSERT 2014/12/04 2000円引限界値

Private Const C_JIMU_TESUURYO   As Long = 5500  '事務手数料                     'INSERT 2015/01/30 K.ISHIZAKA

Private strDate             As String
Private strTime             As String
Private strUser             As String

#Const cnsTest = 0      ' ←本番
'#Const cnsTest = 1      ' ←テスト

Sub test_aa()
    Call MSTGP01_M00
End Sub

Private Sub test()
    Dim strNowMonthBeforeLast As Variant                                        'INSERT 2019/11/08 K.sato
    strNowMonthBeforeLast = DateSerial(Year(DATE), Month(DATE), 0)
MsgBox strNowMonthBeforeLast
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 統合ＤＢの内容を最新状態にする
'       MODULE_ID       : MSTGP01_M00
'       CREATE_DATE     : 2011/02/06            K.ISHIZAKA
'       PARAM           : [strBUMOC]            KOMS接続部門(I)
'                                               カンマ区切りで複数指定可
'                                               省略時は INTI_FILE より取得
'       RETURN          : 正常(True)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSTGP01_M00(Optional ByVal strBUMOC As String = "") As Boolean
    Dim i                   As Long
    Dim i2                  As Long                                             'INSERT 2019/11/08 K.sato
    Dim lngCnt              As Long
    Dim objCon              As Object
    Dim strKaseDb           As String
    Dim varKomsDB           As Variant
    Dim strMsg              As String                                           'INSERT 2011/03/02 K.ISHIZAKA
    Dim strSQL              As String                                           'INSERT 2014/12/04 K.ISHIZAKA
    Dim strNowMonthBeforeLast As Variant                                        'INSERT 2019/11/08 K.sato
    strNowMonthBeforeLast = Format(DateSerial(Year(DATE), Month(DATE), 0), "yyyymmdd") 'INSERT 2019/11/08 K.sato
    Dim varTintaiBumon      As Variant                                          'INSERT 2019/11/08 K.sato
    varTintaiBumon = Array("F", "5", "R")                                       'INSERT 2019/11/08 K.sato
    
    On Error GoTo ErrorHandler
    
    Call MSZZ003_M00(PROG_ID, "0", "")                                          'INSERT 2011/03/02 K.ISHIZAKA
    
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    strUser = Left(LsGetUserName(), 8)
    
    '事務手数料割引の条件取得
    Call MNR100.GetOffiveFeeGetRange                                            'INSERT 2014/12/04 K.ISHIZAKA
    '２千円割引の条件取得
    Call MNR100.Get2kDiscountServiceGetRange                                    'INSERT 2014/12/04 K.ISHIZAKA
    
    If strBUMOC = "" Then
        strBUMOC = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & PROG_ID & "'"), "")
        If strBUMOC = "" Then
            Call MSZZ024_M10("DLookup", "テーブル[INTI_FILE]の設定不正です。")
        End If
    End If
    strKaseDb = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'"), "")
    If strKaseDb = "" Then
        Call MSZZ024_M10("DLookup", "テーブル[SETU_TABL]の設定不正です。")
    End If
    varKomsDB = GetKomsDbNames(strBUMOC)
    
    Set objCon = ADODB_Connection(C_TOGO_CONNENCT_BUMOC)
    On Error GoTo ErrorHandler1
    
    lngCnt = ADODB_Execute("TRUNCATE TABLE TOGO_SPAC_TABL", objCon) '空き状況
    lngCnt = ADODB_Execute("TRUNCATE TABLE TOGO_PRIC_TABL", objCon) '価格表
    lngCnt = ADODB_Execute("TRUNCATE TABLE TOGO_RLDN_TRAN", objCon) 'レンタル物件割引
    lngCnt = ADODB_Execute("TRUNCATE TABLE TOGO_YARD_MAST", objCon) 'ヤードマスタ
    lngCnt = ADODB_Execute("TRUNCATE TABLE TOGO_NYAR_MAST", objCon) '近隣マスタ
    
    '----↓↓↓↓----20111006--M.HONDA--ins-------↓↓↓↓---<s>
    lngCnt = ADODB_Execute("TRUNCATE TABLE TOGO_LOGIN_TABL", objCon) 'ログインテーブル
    lngCnt = ADODB_Execute("TRUNCATE TABLE TOGO_KOKY_MAST", objCon)  '顧客マスタ
    '----↑↑↑↑----20111006--M.HONDA--ins-------↑↑↑↑---<e>
    
    'MsgBox (1)

'#If cnsTest = 0 Then
    For i = LBound(varKomsDB) To UBound(varKomsDB)
        '空き状況（サイズ別）
        lngCnt = ADODB_Execute(Insert_SPAC_TABL_COUNT(varKomsDB(i)), objCon)
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " 空き状況（サイズ別）：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
        '空き状況（コンテナバイク）
        lngCnt = ADODB_Execute(Insert_SPAC_TABL_KBIKE(varKomsDB(i)), objCon)
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " 空き状況（コンテナバイク）：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
        '空き状況（部屋別）
        If Trim(C_USAGE_GRP_ALONE) <> "" Then
            lngCnt = ADODB_Execute(Insert_SPAC_TABL_ALONE(varKomsDB(i)), objCon)
            Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " 空き状況（部屋別）：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
        End If
        lngCnt = ADODB_Execute(Insert_PRIC_TABL(varKomsDB(i)), objCon) '価格表
        lngCnt = lngCnt + ADODB_Execute(Insert_PRIC_TABL(varKomsDB(i), True), objCon) 'INSERT 2011/03/10 K.ISHIZAKA
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " 価格表：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
        lngCnt = ADODB_Execute(Insert_RLDN_TRAN(strKaseDb, varKomsDB(i)), objCon) 'レンタル物件割引
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " レンタル物件割引：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
        lngCnt = ADODB_Execute(Insert_YARD_MAST(strKaseDb, varKomsDB(i)), objCon) 'ヤードマスタ
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " ヤードマスタ：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
        lngCnt = ADODB_Execute(Insert_NYAR_MAST(varKomsDB(i)), objCon) '近隣マスタ
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " 近隣マスタ：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
        
        '----↓↓↓↓----20111006--M.HONDA--ins-------↓↓↓↓---<s>
        Dim varBumonc      As Variant
        varBumonc = Split(strBUMOC, ",")
        lngCnt = ADODB_Execute(Insert_LOGIN_TABL(varKomsDB(i), varBumonc(i)), objCon) 'ログインテーブル
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " ログインテーブル：" & Format(lngCnt))
        lngCnt = ADODB_Execute(Insert_KOKY_MAST(varKomsDB(i), varBumonc(i)), objCon) '顧客マスタ
        Call MSZZ003_M00(PROG_ID, "8", varKomsDB(i) & " 顧客マスタ：" & Format(lngCnt))
        '----↑↑↑↑----20111006--M.HONDA--ins-------↑↑↑↑---<e>
        
    Next
'#End If

    'INSERT 2019/11/08 K.SATO
    'strNowMonthBeforeLast
    
    'MsgBox (2)
    
    For i2 = LBound(varTintaiBumon) To UBound(varTintaiBumon)
        lngCnt = ADODB_Execute(Insert_LOGIN_TABL_from_KASEDB(strKaseDb, varTintaiBumon(i2), strNowMonthBeforeLast), objCon) 'ログインテーブル
        Call MSZZ003_M00(PROG_ID, "8", "KASE_DB" & " ログインテーブル：" & Format(lngCnt))
        lngCnt = ADODB_Execute(Insert_KOKY_MAST_from_KASEDB(strKaseDb, varTintaiBumon(i2), strNowMonthBeforeLast), objCon) '顧客マスタ
        Call MSZZ003_M00(PROG_ID, "8", "KASE_DB" & " 顧客マスタ：" & Format(lngCnt))
    Next
        
'#If cnsTest = 0 Then
    lngCnt = Insert_NYAR_MAST_by_USAG_MAST(objCon) '近隣マスタ（データベースを跨る）
    Call MSZZ003_M00(PROG_ID, "8", "近隣マスタ（データベースを跨る）：" & Format(lngCnt)) 'INSERT 2011/03/02 K.ISHIZAKA
'#End If
    
'    '事務手数料割引                                                             'INSERT START 2014/12/04 K.ISHIZAKA
'    strSQL = Insert_TOGO_RLDN_TRAN_Jimute()
'    If strSQL <> "" Then
'        lngCnt = ADODB_Execute(strSQL, objCon)
'        Call MSZZ003_M00(PROG_ID, "8", "事務手数料割引：" & Format(lngCnt))
'    End If
    
    'ネット契約割引                                                             'INSERT START 2018/01/14 M.HONDA
    strSQL = Insert_TOGO_RLDN_TRAN_Net()
    If strSQL <> "" Then
        lngCnt = ADODB_Execute(strSQL, objCon)
        Call MSZZ003_M00(PROG_ID, "8", "ネット契約割引：" & Format(lngCnt))
    End If
    
    '保証委託料割引
'    lngCnt = ADODB_Execute(Insert_TOGO_RLDN_TRAN_itakuryo(strKaseDb), objCon)
'    Call MSZZ003_M00(PROG_ID, "8", "保証委託料割引：" & Format(lngCnt))         'INSERT END   2014/12/04 K.ISHIZAKA


    objCon.Close
    On Error GoTo ErrorHandler
    Call MSZZ003_M00(PROG_ID, "1", "")                                          'INSERT 2011/03/02 K.ISHIZAKA
    MSTGP01_M00 = True
Exit Function

ErrorHandler1:
    objCon.Close
ErrorHandler:           '↓自分の関数名
'    Call MSZZ024_M00("MSTGP01_M00", False)                                     'DELETE 2011/03/02 K.ISHIZAKA
    strMsg = MSZZ024_M00("MSTGP01_M00", False)                                  'INSERT 2011/03/02 K.ISHIZAKA
    Call MSZZ003_M00(PROG_ID, "9", "Error!!" & vbCrLf & strMsg)                 'INSERT 2011/03/02 K.ISHIZAKA
    MSTGP01_M00 = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : カンマ区切りの接続部門コードをKOMSＤＢ名配列に変換する
'       MODULE_ID       : GetKomsDbNames
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strBUMOC              接続部門(I)カンマ区切りで複数指定可
'       RETURN          : KOMSＤＢ名配列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetKomsDbNames(ByVal strBUMOC As String) As Variant
    Dim i                   As Long
    Dim j                   As Long
    Dim varKomsDB           As Variant
    On Error GoTo ErrorHandler

    varKomsDB = Split(strBUMOC, ",")
    j = LBound(varKomsDB)
    For i = j To UBound(varKomsDB)
        If Trim(varKomsDB(i)) <> "" Then
            varKomsDB(j) = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME_" & Trim(varKomsDB(i)) & "'"), "")
            If varKomsDB(j) = "" Then
                Call MSZZ024_M10("DLookup", "テーブル[SETU_TABL]の設定不正です。")
            End If
            j = j + 1
        End If
    Next
    ReDim Preserve varKomsDB(j - 1)
    GetKomsDbNames = varKomsDB
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "GetKomsDbNames" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヤードマスタ登録
'       MODULE_ID       : Insert_YARD_MAST
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKaseDB             KASEＤＢ名(I)
'                       : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_YARD_MAST(ByVal strKaseDb As String, ByVal strKomsDb As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "INSERT INTO TOGO_YARD_MAST "
    strSQL = strSQL & "("                                                      'INSERT START 2020/09/05 K.ISHIZAKA
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " YARD_USAGE,"
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_YUBINO,"
    strSQL = strSQL & " YARD_ADDR_1,"
    strSQL = strSQL & " YARD_ADDR_2,"
    strSQL = strSQL & " YARD_TUBOSU,"
    strSQL = strSQL & " YARD_TTANKA,"
    strSQL = strSQL & " YARD_RENTGK,"
    strSQL = strSQL & " YARD_UPDATE,"
    strSQL = strSQL & " YARD_NOTE,"
    strSQL = strSQL & " YARD_SEV1N,"
    strSQL = strSQL & " YARD_SEV2N,"
    strSQL = strSQL & " YARD_SEV3N,"
    strSQL = strSQL & " YARD_ENDEN,"
    strSQL = strSQL & " YARD_BIKON,"
    strSQL = strSQL & " YARD_ADDR_3,"
    strSQL = strSQL & " YARD_MNT_TANTO,"
    strSQL = strSQL & " YARD_SEV_EXMONTH,"
    strSQL = strSQL & " YARD_DC_MAX,"
    strSQL = strSQL & " YARD_ZAHYOKEI,"
    strSQL = strSQL & " YARD_IDO,"
    strSQL = strSQL & " YARD_KEIDO,"
    strSQL = strSQL & " YARD_BEGIN_DAY,"
    strSQL = strSQL & " YARD_STOP_FROM,"
    strSQL = strSQL & " YARD_STOP_TO,"
    strSQL = strSQL & " YARD_END_DAY,"
    strSQL = strSQL & " YARD_NONDISP_DAY,"
    strSQL = strSQL & " YARD_RENTEND_DAY,"
    strSQL = strSQL & " YARD_INLIMIT_DAY,"
    strSQL = strSQL & " YARD_INLIMIT_YCODE,"
    strSQL = strSQL & " YARD_NETUSE_KBN,"
    strSQL = strSQL & " YARD_WP_UPDATE,"
    strSQL = strSQL & " YARD_WP_NOTE,"
    strSQL = strSQL & " YARD_NYAR_DATE,"
    strSQL = strSQL & " YARD_WP_UPLOAD,"
    strSQL = strSQL & " YARD_TIHO_KBN,"
    strSQL = strSQL & " YARD_HOSYO_CD,"
    strSQL = strSQL & " YARD_TEL_HOSCD_KIN2,"
    strSQL = strSQL & " YARD_TEL_HOSCD_SHR2,"
    strSQL = strSQL & " YARD_TEL_HOSCD2,"
    strSQL = strSQL & " YARD_NET_HOSYO_CD,"
    strSQL = strSQL & " YARD_NET_HOSCD_KIN2,"
    strSQL = strSQL & " YARD_NET_HOSCD_SHR2,"
    strSQL = strSQL & " YARD_NET_HOSCD2,"
    strSQL = strSQL & " YARD_LEASE_NO,"
    strSQL = strSQL & " YARD_FUZZY_NAME,"
    strSQL = strSQL & " YARD_FUZZY_ADR1,"
    strSQL = strSQL & " YARD_MIKAKUTEI,"
    strSQL = strSQL & " YARD_TODOF,"
    strSQL = strSQL & " YARD_SITSF,"
    strSQL = strSQL & " YARD_CHINF,"
    strSQL = strSQL & " YARD_TODON,"
    strSQL = strSQL & " YARD_SITSN,"
    strSQL = strSQL & " YARD_CHINN,"
    strSQL = strSQL & " YARD_BUMOC,"
    strSQL = strSQL & " YARD_RIYOUJIKAN,"
    strSQL = strSQL & " YARD_KANOUSYA,"
    strSQL = strSQL & " YARD_ZENMEN,"
    strSQL = strSQL & " YARD_SYOUMEI,"
    strSQL = strSQL & " YARD_HOSOU,"
    strSQL = strSQL & " YARD_ELEVATOR,"
    strSQL = strSQL & " YARD_AIRCON,"
    strSQL = strSQL & " YARD_PARKING,"
    strSQL = strSQL & " YARD_SECURITY,"
    strSQL = strSQL & " YARD_VENTILATION,"
    strSQL = strSQL & " YARD_MOYORI_EKI1,"
    strSQL = strSQL & " YARD_MOYORI_EKI2,"
    strSQL = strSQL & " YARD_MOYORI_EKI3,"
    strSQL = strSQL & " YARD_MOYORI_BUS1,"
    strSQL = strSQL & " YARD_MOYORI_BUS2,"
    strSQL = strSQL & " YARD_MOYORI_BUS3,"
    strSQL = strSQL & " YARD_TOKKI1,"
    strSQL = strSQL & " YARD_TOKKI2,"
    strSQL = strSQL & " YARD_TOKKI3,"
    strSQL = strSQL & " YARD_TOKKI4,"
    strSQL = strSQL & " YARD_HEYA_SHITEI_KBN,"
    strSQL = strSQL & " YARD_INSURANCE,"
    strSQL = strSQL & " YARD_RENEWAL_FEE,"
    strSQL = strSQL & " YARD_24HOUR_ACCESS,"
    strSQL = strSQL & " YARD_NEW_PROPERTY,"
    strSQL = strSQL & " YARD_HIGH_QUAL_PARTITION,"
    strSQL = strSQL & " YARD_REASONABLE,"
    strSQL = strSQL & " YARD_CAMPAIGN,"
    strSQL = strSQL & " YARD_HUMIDITY_CTRL,"
    strSQL = strSQL & " YARD_STORAGE_CAPA,"
    strSQL = strSQL & " YARD_TRANSPORT_SERVICE,"
    strSQL = strSQL & " YARD_KEYLESS,"
    strSQL = strSQL & " YARD_SYUKYAKU_KBN,"
    strSQL = strSQL & " YARD_KEYWORD "
    strSQL = strSQL & ") "                                                      'INSERT END   2020/09/05 K.ISHIZAKA
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " SPAC_USAGE AS YARD_USAGE,"
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_YUBINO,"
    strSQL = strSQL & " YARD_ADDR_1,"
    strSQL = strSQL & " YARD_ADDR_2,"
    strSQL = strSQL & " YARD_TUBOSU,"
    strSQL = strSQL & " YARD_TTANKA,"
    strSQL = strSQL & " YARD_RENTGK,"
    strSQL = strSQL & " YARD_UPDATE,"
    strSQL = strSQL & " YARD_NOTE,"
    strSQL = strSQL & " CASE WHEN RLDNT_CNT > 0 THEN NULLIF(YARD_SEV1N,'') ELSE NULL END AS YARD_SEV1N, "
    strSQL = strSQL & " CASE WHEN RLDNT_CNT > 0 THEN NULLIF(YARD_SEV2N,'') ELSE NULL END AS YARD_SEV2N,"
    strSQL = strSQL & " CASE WHEN RLDNT_CNT > 0 THEN NULLIF(YARD_SEV3N,'') ELSE NULL END AS YARD_SEV3N,"
    strSQL = strSQL & " CASE WHEN RLDNT_CNT > 0 THEN NULLIF(YARD_ENDEN,'') ELSE NULL END AS YARD_ENDEN,"
    strSQL = strSQL & " YARD_BIKON,"
    strSQL = strSQL & " YARD_ADDR_3,"
    strSQL = strSQL & " YARD_MNT_TANTO,"
    strSQL = strSQL & " YARD_SEV_EXMONTH,"
    strSQL = strSQL & " CONT_DC_MAX AS YARD_DC_MAX,"
    strSQL = strSQL & " YARD_ZAHYOKEI,"
    strSQL = strSQL & " YARD_IDO,"
    strSQL = strSQL & " YARD_KEIDO,"
    strSQL = strSQL & " YARD_BEGIN_DAY,"
    strSQL = strSQL & " YARD_STOP_FROM,"                                        'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & " YARD_STOP_TO,"                                          'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & " YARD_END_DAY,"
    strSQL = strSQL & " YARD_NONDISP_DAY,"
    strSQL = strSQL & " YARD_RENTEND_DAY,"
    strSQL = strSQL & " YARD_INLIMIT_DAY,"
    strSQL = strSQL & " YARD_INLIMIT_YCODE,"
    strSQL = strSQL & " YARD_NETUSE_KBN,"
    strSQL = strSQL & " YARD_WP_UPDATE,"
    strSQL = strSQL & " YARD_WP_NOTE,"
    strSQL = strSQL & " YARD_NYAR_DATE,"
    strSQL = strSQL & " YARD_WP_UPLOAD,"
    strSQL = strSQL & " YARD_TIHO_KBN, "      '2014/11/12 M.HONDA INS
    strSQL = strSQL & " YARD_HOSYO_CD, "
    strSQL = strSQL & " YARD_TEL_HOSCD_KIN2, "   '2018/01/31 M.HONDA INS
    strSQL = strSQL & " YARD_TEL_HOSCD_SHR2, "   '2018/01/31 M.HONDA INS
    strSQL = strSQL & " YARD_TEL_HOSCD2, "       '2018/01/31 M.HONDA INS
    strSQL = strSQL & " YARD_NET_HOSYO_CD, "     '2014/11/12 M.HONDA INS
    strSQL = strSQL & " YARD_NET_HOSCD_KIN2, "   '2018/01/14 M.HONDA INS
    strSQL = strSQL & " YARD_NET_HOSCD_SHR2, "   '2018/01/14 M.HONDA INS
    strSQL = strSQL & " YARD_NET_HOSCD2, "       '2018/01/14 M.HONDA INS
    strSQL = strSQL & " YARD_LEASE_NO,"
    strSQL = strSQL & " " & changeFuzzySQL("YARD_NAME") & " AS YARD_FUZZY_NAME,"
    strSQL = strSQL & " " & changeFuzzySQL("YARD_ADDR_1") & " AS YARD_FUZZY_ADR1,"
    
    strSQL = strSQL & "(SELECT COUNT(*)"
    strSQL = strSQL & " FROM   " & strKomsDb & ".dbo.YOUK_TRAN"
    strSQL = strSQL & " WHERE  YOUKT_YCODE = YARD_CODE"
    strSQL = strSQL & " AND    YOUKT_YUKBN = 0"
    strSQL = strSQL & " AND ( (ISNULL(YOUKT_USAGE, SPAC_USAGE) = SPAC_USAGE)"
    strSQL = strSQL & "    OR (YOUKT_USAGE = 10 AND SPAC_USAGE IN(0,1))"
    strSQL = strSQL & "    OR (YOUKT_USAGE = 30 AND SPAC_USAGE IN(3,31,32,33))"
    strSQL = strSQL & "     )"
    strSQL = strSQL & ") AS YARD_MIKAKUTEI,"
    
    strSQL = strSQL & " ADDRT_TODOF AS YARD_TODOF,"
    strSQL = strSQL & " ADDRT_SITSF AS YARD_SITSF,"
    strSQL = strSQL & " ADDRT_CHINF AS YARD_CHINF,"
    strSQL = strSQL & " ADDRT_TODON AS YARD_TODON,"
    strSQL = strSQL & " ADDRT_SITSN AS YARD_SITSN,"
    strSQL = strSQL & " ADDRT_CHINN AS YARD_CHINN,"
    strSQL = strSQL & " CONT_BUMOC AS YARD_BUMOC "
    
    strSQL = strSQL & " ,YARD_RIYOUJIKAN"                                       'INSERT START 2013/11/03 K.ISHIZAKA
    strSQL = strSQL & " ,YARD_KANOUSYA"
    strSQL = strSQL & " ,YARD_ZENMEN"
    strSQL = strSQL & " ,YARD_SYOUMEI"
    strSQL = strSQL & " ,YARD_HOSOU"
    '2014/10/06 M.HONDA INS
    strSQL = strSQL & " ,YARD_ELEVATOR"
    strSQL = strSQL & " ,YARD_AIRCON"
    strSQL = strSQL & " ,YARD_PARKING"
    '2014/10/06 M.HONDA INS
    strSQL = strSQL & " ,YARD_SECURITY "                                        'INSERT 2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " ,YARD_VENTILATION "                                     'INSERT 2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " ,YARD_MOYORI_EKI1"
    strSQL = strSQL & " ,YARD_MOYORI_EKI2"
    strSQL = strSQL & " ,YARD_MOYORI_EKI3"
    strSQL = strSQL & " ,YARD_MOYORI_BUS1"
    strSQL = strSQL & " ,YARD_MOYORI_BUS2"
    strSQL = strSQL & " ,YARD_MOYORI_BUS3"
    strSQL = strSQL & " ,YARD_TOKKI1"
    strSQL = strSQL & " ,YARD_TOKKI2"
    strSQL = strSQL & " ,YARD_TOKKI3"
    strSQL = strSQL & " ,YARD_TOKKI4 "                                           'INSERT END   2013/11/03 K.ISHIZAKA
    '2016/08/17 M.HONDA INS
    strSQL = strSQL & " ,(CASE YARD_HEYA_SHITEI_KBN "
    strSQL = strSQL & "   WHEN -1 THEN 1 "
    strSQL = strSQL & "   ELSE 0    END )  "
    strSQL = strSQL & "   YARD_HEYA_SHITEI_KBN "
    '2016/08/17 M.HONDA INS
'↓ADD 2019/04/05 Y_SUZUKI EGL
    strSQL = strSQL & " ,YARD_INSURANCE "
    strSQL = strSQL & " ,YARD_RENEWAL_FEE "
    strSQL = strSQL & " ,YARD_24HOUR_ACCESS "
    strSQL = strSQL & " ,YARD_NEW_PROPERTY "
    strSQL = strSQL & " ,YARD_HIGH_QUAL_PARTITION "
    strSQL = strSQL & " ,YARD_REASONABLE "
    strSQL = strSQL & " ,YARD_CAMPAIGN "
    strSQL = strSQL & " ,YARD_HUMIDITY_CTRL "
    strSQL = strSQL & " ,YARD_STORAGE_CAPA "
    strSQL = strSQL & " ,YARD_TRANSPORT_SERVICE "
    strSQL = strSQL & " ,YARD_KEYLESS "
'↑ADD 2019/04/05 Y_SUZUKI EGL
'↓ADD 2019/08/17 Y_SUZUKI EGL
    strSQL = strSQL & " ,YARD_SYUKYAKU_KBN "
'↑ADD 2019/08/17 Y_SUZUKI EGL
    strSQL = strSQL & " ,YARD_KEYWORD "                                         'INSERT 2020/09/05 K.ISHIZAKA
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CONT_MAST, "
    strSQL = strSQL & " " & strKomsDb & ".dbo.YARD_MAST src "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " SPAC_YCODE,"
    strSQL = strSQL & " SPAC_USAGE,"
    strSQL = strSQL & " COUNT(RLDNT_YCODE) AS RLDNT_CNT "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " TOGO_SPAC_TABL "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " TOGO_RLDN_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " RLDNT_YCODE = SPAC_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_USAGE = SPAC_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_SIZE = SPAC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_STEP = SPAC_STEP "
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " SPAC_YCODE,"
    strSQL = strSQL & " SPAC_USAGE "
    strSQL = strSQL & ") spac "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " SPAC_YCODE = YARD_CODE "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKaseDb & ".dbo.ADDR_TABL "
    strSQL = strSQL & "ON"
    ''INSERT 2011/06/18 M.HONDA START
    ''strSQL = strSQL & " ADDRT_TODON + ISNULL(ADDRT_SITSN,'') + ISNULL(ADDRT_CHINN,'') = YARD_ADDR_1 "
    strSQL = strSQL & " ADDRT_NYUBB =  REPLACE(YARD_YUBINO,'-','') "
    ''INSERT 2011/06/18 M.HONDA END
    
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CONT_KEY = 1 "

    Insert_YARD_MAST = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 近隣マスタ登録（同一データベース内）
'       MODULE_ID       : Insert_NYAR_MAST
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_NYAR_MAST(ByVal strKomsDb As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "INSERT INTO TOGO_NYAR_MAST "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " base.YARD_CODE AS NYAR_YCODE,"
    strSQL = strSQL & " base.YARD_USAGE AS NYAR_USAGE,"
    strSQL = strSQL & " nyar.YARD_CODE AS NYAR_NCODE,"
    strSQL = strSQL & " nyar.YARD_USAGE AS NYAR_NSAGE,"
    strSQL = strSQL & " NYAR_INSED,"
    strSQL = strSQL & " NYAR_INSEJ,"
    strSQL = strSQL & " NYAR_INSPB,"
    strSQL = strSQL & " NYAR_INSUB,"
    strSQL = strSQL & " NYAR_UPDAD,"
    strSQL = strSQL & " NYAR_UPDAJ,"
    strSQL = strSQL & " NYAR_UPDPB,"
    strSQL = strSQL & " NYAR_UPDUB,"
    strSQL = strSQL & " NYAR_KIRO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.NYAR_MAST "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_YARD_MAST base "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " base.YARD_CODE = NYAR_YCODE "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_YARD_MAST nyar "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " nyar.YARD_CODE = NYAR_NCODE "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_USAG_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " USAG_USAGE = base.YARD_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " USAG_NSAGE = nyar.YARD_USAGE "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " base.YARD_BUMOC = nyar.YARD_BUMOC "

    Insert_NYAR_MAST = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 価格表登録
'       MODULE_ID       : Insert_PRIC_TABL
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function Insert_PRIC_TABL(ByVal strKomsDb As String) As String         'DELETE 2011/03/10 K.ISHIZAKA
Private Function Insert_PRIC_TABL(ByVal strKomsDb As String, Optional blKBike As Boolean = False) As String 'INSERT 2011/03/10 K.ISHIZAKA
    Dim strSQL              As String

    strSQL = strSQL & "INSERT INTO TOGO_PRIC_TABL "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " PRIC_YCODE,"
'    strSQL = strSQL & " PRIC_USAGE,"                                           'DELETE 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & " SPAC_USAGE AS PRIC_USAGE,"                              'INSERT 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & " PRIC_SIZE,"
    strSQL = strSQL & " PRIC_STEP,"
    '----↓↓↓↓----20110906--M.HONDA--update-------↓↓↓↓---<s>
    'strSQL = strSQL & " MIN(CNTA_DEPTH) AS PRIC_DEPTH,"
    'strSQL = strSQL & " MIN(CNTA_WIDTH) AS PRIC_WIDTH,"
    'strSQL = strSQL & " MIN(CNTA_HEIGHT) AS PRIC_HEIGHT,"
    strSQL = strSQL & " CNTA_DEPTH AS PRIC_DEPTH,"
    strSQL = strSQL & " CNTA_WIDTH AS PRIC_WIDTH,"
    strSQL = strSQL & " CNTA_HEIGHT AS PRIC_HEIGHT,"
    '----↑↑↑↑----20110906--M.HONDA--update-------↑↑↑↑---<e>
    strSQL = strSQL & " PRIC_FPRICE, "                                          '2019/08/08 M.HONDA INS
    strSQL = strSQL & " PRIC_PRICE,"
    strSQL = strSQL & " PRIC_FZAPPI_CODE,"
    strSQL = strSQL & " PRIC_FZAPPI,"
    strSQL = strSQL & " PRIC_FZAPPI_CODE2,"                                     'INSERT START 2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " PRIC_FZAPPI2,"
    strSQL = strSQL & " PRIC_FZAPPI_CODE3,"
    strSQL = strSQL & " PRIC_FZAPPI3,"                                          'INSERT END   2018/08/22 K.ISHIZAKA
    '2016/09/28 M.HONDA UPD
'    If blKBike Then                                                             'UPD START 2013/04/06 M.HONDA
'        strSQL = strSQL & " 0,"
'        strSQL = strSQL & " 0,"
'    Else
'        strSQL = strSQL & " PRIC_EZAPPI_CODE,"
'        strSQL = strSQL & " PRIC_EZAPPI,"
'    End If                                                                      'UPD START 2013/04/06 M.HONDA
    strSQL = strSQL & " PRIC_EZAPPI_CODE,"
    strSQL = strSQL & " PRIC_FEZAPPI, "                                          '2019/08/08 M.HONDA INS
    strSQL = strSQL & " PRIC_EZAPPI,"
    '2016/09/28 M.HONDA UPD
    strSQL = strSQL & " PRIC_EZAPPI_CODE2,"                                     'INSERT START 2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " PRIC_EZAPPI2,"
    strSQL = strSQL & " PRIC_EZAPPI_CODE3,"
    strSQL = strSQL & " PRIC_EZAPPI3,"                                          'INSERT END   2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " PRIC_NEBIKI_DISABLE,"
    strSQL = strSQL & " PRIC_INSED,"
    strSQL = strSQL & " PRIC_INSEJ,"
    strSQL = strSQL & " PRIC_INSPB,"
    strSQL = strSQL & " PRIC_INSUB,"
    strSQL = strSQL & " PRIC_UPDAD,"
    strSQL = strSQL & " PRIC_UPDAT,"
    strSQL = strSQL & " PRIC_UPDPB,"
    strSQL = strSQL & " PRIC_UPDUB "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.PRIC_TABL "
    strSQL = strSQL & "INNER JOIN"
    
    '----↓↓↓↓----20110906--M.HONDA--update-------↓↓↓↓---<s>
    strSQL = strSQL & "( SELECT  HEIGHT.CNTA_CODE, "
    strSQL = strSQL & "        HEIGHT.CNTA_USAGE, "
    strSQL = strSQL & "        HEIGHT.CNTA_SIZE, "
    strSQL = strSQL & "        HEIGHT.CNTA_STEP, "
    strSQL = strSQL & "        CNTA_DEPTH.CNTA_WIDTH, "
    strSQL = strSQL & "        CNTA_DEPTH.CNTA_DEPTH, "
    strSQL = strSQL & "        MIN(CNTA_HEIGHT) As CNTA_HEIGHT, "
    strSQL = strSQL & "        HEIGHT.CNTA_USE, "
    strSQL = strSQL & "        HEIGHT.CNTA_BIKE_OK "
    strSQL = strSQL & "FROM " & strKomsDb & ".DBO.CNTA_MAST HEIGHT , "
    strSQL = strSQL & "     ( "
    strSQL = strSQL & "        SELECT DEPTH.CNTA_CODE, "
    strSQL = strSQL & "               DEPTH.CNTA_USAGE, "
    strSQL = strSQL & "               DEPTH.CNTA_SIZE, "
    strSQL = strSQL & "               DEPTH.CNTA_STEP, "
    strSQL = strSQL & "               WIDTH.CNTA_WIDTH, "
    strSQL = strSQL & "               MIN(CNTA_DEPTH) As CNTA_DEPTH "
    strSQL = strSQL & "        FROM " & strKomsDb & ".DBO.CNTA_MAST DEPTH , "
    strSQL = strSQL & "             ( "
    strSQL = strSQL & "                SELECT   CNTA_CODE, "
    strSQL = strSQL & "                         CNTA_USAGE, "
    strSQL = strSQL & "                         CNTA_SIZE, "
    strSQL = strSQL & "                         CNTA_STEP, "
    strSQL = strSQL & "                         MIN(CNTA_WIDTH) As CNTA_WIDTH "
    strSQL = strSQL & "                FROM " & strKomsDb & ".DBO.CNTA_MAST "
    strSQL = strSQL & "                GROUP BY CNTA_CODE, "
    strSQL = strSQL & "                         CNTA_USAGE, "
    strSQL = strSQL & "                         CNTA_SIZE, "
    strSQL = strSQL & "                         CNTA_STEP ) WIDTH "
    strSQL = strSQL & "        WHERE DEPTH.CNTA_CODE = Width.CNTA_CODE "
    strSQL = strSQL & "         AND    DEPTH.CNTA_USAGE = WIDTH.CNTA_USAGE "
    strSQL = strSQL & "         AND    DEPTH.CNTA_SIZE = WIDTH.CNTA_SIZE "
    strSQL = strSQL & "         AND    DEPTH.CNTA_STEP = WIDTH.CNTA_STEP "
    strSQL = strSQL & "         AND    DEPTH.CNTA_WIDTH = WIDTH.CNTA_WIDTH "
    strSQL = strSQL & "        GROUP BY DEPTH.CNTA_CODE, "
    strSQL = strSQL & "               DEPTH.CNTA_USAGE, "
    strSQL = strSQL & "               DEPTH.CNTA_SIZE, "
    strSQL = strSQL & "               DEPTH.CNTA_STEP, "
    strSQL = strSQL & "               WIDTH.CNTA_WIDTH ) CNTA_DEPTH "
    strSQL = strSQL & "WHERE Height.CNTA_CODE = CNTA_DEPTH.CNTA_CODE "
    strSQL = strSQL & " AND    HEIGHT.CNTA_USAGE = CNTA_DEPTH.CNTA_USAGE "
    strSQL = strSQL & " AND    HEIGHT.CNTA_SIZE = CNTA_DEPTH.CNTA_SIZE "
    strSQL = strSQL & " AND    HEIGHT.CNTA_STEP = CNTA_DEPTH.CNTA_STEP "
    strSQL = strSQL & " AND    HEIGHT.CNTA_WIDTH = CNTA_DEPTH.CNTA_WIDTH "
    strSQL = strSQL & " AND    HEIGHT.CNTA_DEPTH = CNTA_DEPTH.CNTA_DEPTH "
    strSQL = strSQL & "GROUP BY HEIGHT.CNTA_CODE, "
    strSQL = strSQL & "        HEIGHT.CNTA_USAGE, "
    strSQL = strSQL & "        HEIGHT.CNTA_SIZE, "
    strSQL = strSQL & "        HEIGHT.CNTA_STEP, "
    strSQL = strSQL & "        CNTA_DEPTH.CNTA_WIDTH, "
    strSQL = strSQL & "        CNTA_DEPTH.CNTA_DEPTH, "
    strSQL = strSQL & "        HEIGHT.CNTA_USE, "
    strSQL = strSQL & "        HEIGHT.CNTA_BIKE_OK ) CNTA "
    ''strSQL = strSQL & " " & strKomsDb & ".dbo.CNTA_MAST "
    '----↑↑↑↑----20110906--M.HONDA--update-------↑↑↑↑---<e>
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CNTA_CODE = PRIC_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = PRIC_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_SIZE = PRIC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_STEP = PRIC_STEP "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_SPAC_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " SPAC_YCODE = PRIC_YCODE "
    strSQL = strSQL & "AND"
    If blKBike Then                                                             'INSERT START 2011/03/10 K.ISHIZAKA
        strSQL = strSQL & " SPAC_USAGE = 33 "
    Else                                                                        'INSERT END   2011/03/10 K.ISHIZAKA
        strSQL = strSQL & " SPAC_USAGE = PRIC_USAGE "
    End If                                                                      'INSERT 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_SIZE = PRIC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_STEP = PRIC_STEP "
    
    strSQL = strSQL & "WHERE"                                                   'INSERT START 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & " CNTA_USE = 1 "
    If blKBike Then
        strSQL = strSQL & "AND"
        strSQL = strSQL & " CNTA_USAGE = 0 "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " CNTA_BIKE_OK = 1 "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " NOT EXISTS"
        strSQL = strSQL & "("
        strSQL = strSQL & "SELECT"
        strSQL = strSQL & " * "
        strSQL = strSQL & "FROM"
        strSQL = strSQL & " TOGO_PRIC_TABL ch "
        strSQL = strSQL & "WHERE"
        strSQL = strSQL & " ch.PRIC_YCODE = CNTA_CODE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " ch.PRIC_USAGE = 33 "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " ch.PRIC_SIZE = CNTA_SIZE "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " ch.PRIC_STEP = CNTA_STEP "
        strSQL = strSQL & ")"
    End If                                                                      'INSERT END   2011/03/10 K.ISHIZAKA
    
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " PRIC_YCODE,"
'    strSQL = strSQL & " PRIC_USAGE,"                                           'DELETE 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & " SPAC_USAGE,"                                            'INSERT 2011/03/10 K.ISHIZAKA
    strSQL = strSQL & " PRIC_SIZE,"
    
    
    strSQL = strSQL & " PRIC_FPRICE, "                                          '2019/08/08 M.HONDA INS
    strSQL = strSQL & " PRIC_FEZAPPI, "                                          '2019/08/08 M.HONDA INS
    
    strSQL = strSQL & " PRIC_STEP,"
    strSQL = strSQL & " PRIC_PRICE,"
    strSQL = strSQL & " PRIC_FZAPPI_CODE,"
    strSQL = strSQL & " PRIC_FZAPPI,"
    strSQL = strSQL & " PRIC_FZAPPI_CODE2,"                                     'INSERT START 2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " PRIC_FZAPPI2,"
    strSQL = strSQL & " PRIC_FZAPPI_CODE3,"
    strSQL = strSQL & " PRIC_FZAPPI3,"                                          'INSERT END   2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " PRIC_EZAPPI_CODE,"
    strSQL = strSQL & " PRIC_EZAPPI,"
    strSQL = strSQL & " PRIC_EZAPPI_CODE2,"                                     'INSERT START 2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " PRIC_EZAPPI2,"
    strSQL = strSQL & " PRIC_EZAPPI_CODE3,"
    strSQL = strSQL & " PRIC_EZAPPI3,"                                          'INSERT END   2018/08/22 K.ISHIZAKA
    strSQL = strSQL & " PRIC_NEBIKI_DISABLE,"
    '----↓↓↓↓----20110906--M.HONDA--update-------↓↓↓↓---<s>
    strSQL = strSQL & " CNTA_DEPTH,"
    strSQL = strSQL & " CNTA_WIDTH,"
    strSQL = strSQL & " CNTA_HEIGHT,"
    '----↑↑↑↑----20110906--M.HONDA--update-------↑↑↑↑---<e>
    strSQL = strSQL & " PRIC_INSED,"
    strSQL = strSQL & " PRIC_INSEJ,"
    strSQL = strSQL & " PRIC_INSPB,"
    strSQL = strSQL & " PRIC_INSUB,"
    strSQL = strSQL & " PRIC_UPDAD,"
    strSQL = strSQL & " PRIC_UPDAT,"
    strSQL = strSQL & " PRIC_UPDPB,"
    strSQL = strSQL & " PRIC_UPDUB "

    Insert_PRIC_TABL = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : レンタル物件割引登録
'       MODULE_ID       : Insert_RLDN_TRAN
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKaseDB             KASEＤＢ名(I)
'                       : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_RLDN_TRAN(ByVal strKaseDb As String, ByVal strKomsDb As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "INSERT INTO TOGO_RLDN_TRAN "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " RLDNT_BUMOC,"
    strSQL = strSQL & " PRIC_YCODE AS RLDNT_YCODE,"
    strSQL = strSQL & " PRIC_USAGE AS RLDNT_USAGE,"
    strSQL = strSQL & " PRIC_SIZE AS RLDNT_SIZE,"
    strSQL = strSQL & " PRIC_STEP AS RLDNT_STEP,"
    strSQL = strSQL & " RLDNT_NO,"
    strSQL = strSQL & " RLDNT_ENABLE,"
    strSQL = strSQL & " CASE WHEN DCNTM_TYPE BETWEEN 1 AND 3 THEN 0 ELSE DCNTM_TYPE END AS RLDNT_TYPE,"
    
    '2019/08/08 M.HONDA INS
    strSQL = strSQL & " CASE DCNTM_TYPE"
    strSQL = strSQL & " WHEN 1 THEN DCNTM_VALUE "  '円指定
    strSQL = strSQL & " WHEN 2 THEN PRIC_FPRICE - DCNTM_VALUE "    '円引き
    strSQL = strSQL & " WHEN 3 THEN PRIC_FPRICE - (FLOOR(PRIC_FPRICE * (DCNTM_VALUE / 100))) " '％引き
    'strSQL = strSQL & " ELSE DCNTM_VALUE "  '事務手数料割引                                    'DELETE 2021/06/16 N.IMAI
    strSQL = strSQL & " ELSE 0 "             '事務手数料割引                                    'INSERT 2021/06/16 N.IMAI
    strSQL = strSQL & " END AS RLDNT_FPRICE,"
    
    strSQL = strSQL & " CASE DCNTM_TYPE"
    strSQL = strSQL & " WHEN 1 THEN PRIC_FPRICE - DCNTM_VALUE "  '円指定
    strSQL = strSQL & " WHEN 2 THEN DCNTM_VALUE "    '円引き
    strSQL = strSQL & " WHEN 3 THEN (FLOOR(PRIC_FPRICE * (DCNTM_VALUE / 100))) " '％引き
    'strSQL = strSQL & " ELSE " & Format(C_JIMU_TESUURYO) & " - DCNTM_VALUE "  '事務手数料割引  'DELETE 2021/06/16 N.IMAI
    strSQL = strSQL & " ELSE 0 "  '事務手数料割引                                               'INSERT 2021/06/16 N.IMAI
    strSQL = strSQL & " END AS RLDNT_FWARIBIKI,"                                 '
    '2019/08/08 M.HONDA INS
    
    strSQL = strSQL & " CASE DCNTM_TYPE"
    strSQL = strSQL & " WHEN 1 THEN DCNTM_VALUE "  '円指定
    strSQL = strSQL & " WHEN 2 THEN PRIC_PRICE - DCNTM_VALUE "    '円引き
    strSQL = strSQL & " WHEN 3 THEN PRIC_PRICE - (FLOOR(PRIC_PRICE * (DCNTM_VALUE / 100))) " '％引き
    'strSQL = strSQL & " ELSE DCNTM_VALUE "  '事務手数料割引      'INSERT 2014/12/04 K.ISHIZAKA 'DELETE 2021/06/16 N.IMAI
    strSQL = strSQL & " ELSE 0 "             '事務手数料割引                                    'INSERT 2021/06/16 N.IMAI
    strSQL = strSQL & " END AS RLDNT_PRICE,"
    strSQL = strSQL & " CASE DCNTM_TYPE"                                        'INSERT START 2015/01/30 K.ISHIZAKA
    strSQL = strSQL & " WHEN 1 THEN PRIC_PRICE - DCNTM_VALUE "  '円指定
    strSQL = strSQL & " WHEN 2 THEN DCNTM_VALUE "    '円引き
    strSQL = strSQL & " WHEN 3 THEN (FLOOR(PRIC_PRICE * (DCNTM_VALUE / 100))) " '％引き
    'strSQL = strSQL & " ELSE " & Format(C_JIMU_TESUURYO) & " - DCNTM_VALUE "  '事務手数料割引  'DELETE 2021/06/16 N.IMAI
    strSQL = strSQL & " ELSE 0 "  '事務手数料割引                                               'INSERT 2021/06/16 N.IMAI
    strSQL = strSQL & " END AS RLDNT_WARIBIKI,"                                 'INSERT END   2015/01/30 K.ISHIZAKA
    strSQL = strSQL & " DCNTM_PERIOD AS RLDNT_PERIOD,"
    strSQL = strSQL & " RLDNT_FROM,"
    strSQL = strSQL & " RLDNT_TO,"
    strSQL = strSQL & " RLDNT_ORDER,"
    strSQL = strSQL & " RLDNT_TEXT,"
    strSQL = strSQL & " RLDNT_NOTE,"
    strSQL = strSQL & " DCNTM_GENKBN AS RLDNT_GENKBN,"
    strSQL = strSQL & " DCNTM_USE_PERIOD AS RLDNT_USE_PERIOD,"
    strSQL = strSQL & " RLDNT_INSED,"
    strSQL = strSQL & " RLDNT_INSEJ,"
    strSQL = strSQL & " RLDNT_INSPB,"
    strSQL = strSQL & " RLDNT_INSUB,"
    strSQL = strSQL & " RLDNT_UPDAD,"
    strSQL = strSQL & " RLDNT_UPDAJ,"
    strSQL = strSQL & " RLDNT_UPDPB,"
    strSQL = strSQL & " RLDNT_UPDUB "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKaseDb & ".dbo.DCNT_MAST "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.RLDN_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " RLDNT_BUMOC = DCNTM_BUMOC "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_NO = DCNTM_NO "
    
    If MNR100.IsOffiveFeeGet() Then                                             'INSERT 2014/12/04 K.ISHIZAKA
        '事務手数料割引の期間内のときはここで作成しない
        '----↓↓↓↓----20110930--M.HONDA--Ins-------↓↓↓↓---<s>
        strSQL = strSQL & "AND"
        strSQL = strSQL & " DCNTM_TYPE <> 10 "
        '----↑↑↑↑----20110930--M.HONDA--Ins-------↑↑↑↑---<e>
    End If                                                                      'INSERT 2014/12/04 K.ISHIZAKA
    
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_PRIC_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PRIC_YCODE = CAST(RLDNT_YCODE AS NUMERIC(6)) "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PRIC_USAGE = ISNULL(DCNTM_USAGE,PRIC_USAGE) "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PRIC_SIZE BETWEEN ISNULL(DCNTM_SIZE_FROM,0) AND ISNULL(DCNTM_SIZE_TO,999.99) "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PRIC_STEP = ISNULL(DCNTM_FLOOR,PRIC_STEP) "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PRIC_NEBIKI_DISABLE = 0 " '値引き対象
    strSQL = strSQL & "AND"                                                     'INSERT 2011/03/08 K.ISHIZAKA
    strSQL = strSQL & " RLDNT_TO >= '" & strDate & "' "                         'INSERT 2011/03/08 K.ISHIZAKA
    '----↓↓↓↓----20110930--M.HONDA--Ins-------↓↓↓↓---<s>
    strSQL = strSQL & " AND "
    strSQL = strSQL & " RLDNT_ENABLE = 1 "
    '----↑↑↑↑----20110930--M.HONDA--Ins-------↑↑↑↑---<e>

    If MNR100.Is2kDiscountServiceGet() Then                                     'INSERT START 2014/12/04 K.ISHIZAKA
        '２千円割引が適用される場合は価格の安いものだけが事務手数料の対象となる
        strSQL = strSQL & "WHERE"
        strSQL = strSQL & " PRIC_PRICE < " & Format(C_2kSEV_LIMIT) & " "
        strSQL = strSQL & "OR"
        strSQL = strSQL & " DCNTM_TYPE <> 10 "
    End If                                                                      'INSERT END   2014/12/04 K.ISHIZAKA

    Insert_RLDN_TRAN = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（サイズ別）
'       MODULE_ID       : Insert_SPAC_TABL_COUNT
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_SPAC_TABL_COUNT(ByVal strKomsDb As String) As String
    Dim strSQL              As String

    'strSQL = strSQL & "INSERT INTO TOGO_SPAC_TABL "
    strSQL = strSQL & "INSERT INTO TOGO_SPAC_TABL_J "                           'INSERT 2026/04/22 N.IMAI
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE AS SPAC_YCODE,"
    strSQL = strSQL & " CNTA_USAGE AS SPAC_USAGE,"
    strSQL = strSQL & " CNTA_SIZE AS SPAC_SIZE,"
    strSQL = strSQL & " CNTA_STEP AS SPAC_STEP,"
    strSQL = strSQL & " CAST(0 AS NUMERIC(6,0)) AS SPAC_NO,"
    strSQL = strSQL & " CNTA_FLOOR AS SPAC_FLOOR,"                              'INSERT 2018/07/19 K.ISHIZAKA
    
    strSQL = strSQL & Insert_SPAC_TABL_COMM(strKomsDb)
    
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE IN (" & C_USAGE_GRP_COUNT & ") "
    strSQL = strSQL & ") AS wk "                                                'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP "
    strSQL = strSQL & ",CNTA_FLOOR "                                            'INSERT 2018/07/19 K.ISHIZAKA

    Insert_SPAC_TABL_COUNT = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（コンテナバイク）
'       MODULE_ID       : Insert_SPAC_TABL_KBIKE
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_SPAC_TABL_KBIKE(ByVal strKomsDb As String) As String
    Dim strSQL              As String

    'strSQL = strSQL & "INSERT INTO TOGO_SPAC_TABL "
    strSQL = strSQL & "INSERT INTO TOGO_SPAC_TABL_J "                           'INSERT 2026/04/22 N.IMAI
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE AS SPAC_YCODE,"
    strSQL = strSQL & " " & C_USAGE_GRP_KBIKE & " AS SPAC_USAGE,"
    strSQL = strSQL & " CNTA_SIZE AS SPAC_SIZE,"
    strSQL = strSQL & " CNTA_STEP AS SPAC_STEP,"
    strSQL = strSQL & " CAST(0 AS NUMERIC(6,0)) AS SPAC_NO,"
    strSQL = strSQL & " CAST(1 AS NUMERIC(3,0)) AS SPAC_FLOOR,"                 'INSERT 2018/07/19 K.ISHIZAKA
    
    strSQL = strSQL & Insert_SPAC_TABL_COMM(strKomsDb)
    
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = 0 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_BIKE_OK = 1 "
    strSQL = strSQL & ") AS wk "                                                'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP "

    Insert_SPAC_TABL_KBIKE = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（部屋別）
'       MODULE_ID       : Insert_SPAC_TABL_ALONE
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_SPAC_TABL_ALONE(ByVal strKomsDb As String) As String
    Dim strSQL              As String

    'strSQL = strSQL & "INSERT INTO TOGO_SPAC_TABL "
    strSQL = strSQL & "INSERT INTO TOGO_SPAC_TABL_J "                           'INSERT 2026/04/22 N.IMAI
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE AS SPAC_YCODE,"
    strSQL = strSQL & " CNTA_USAGE AS SPAC_USAGE,"
    strSQL = strSQL & " CNTA_SIZE AS SPAC_SIZE,"
    strSQL = strSQL & " CNTA_STEP AS SPAC_STEP,"
    strSQL = strSQL & " CNTA_NO AS SPAC_NO,"
    strSQL = strSQL & " CNTA_FLOOR AS SPAC_FLOOR,"                              'INSERT 2018/07/19 K.ISHIZAKA
    '2016/08/18 M.HONDA UPD
    'strSQL = strSQL & Insert_SPAC_TABL_COMM(strKomsDb)
    strSQL = strSQL & Insert_SPAC_TABL_ALONE_COMM(strKomsDb)
    '2016/08/18 M.HONDA UPD
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE IN (" & C_USAGE_GRP_ALONE & ") "
    strSQL = strSQL & ") AS wk "                                                'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & "  WHERE CNTA_USE != 9 "      '2020/03/11 M.HONDA
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP,"
    strSQL = strSQL & " CNTA_NO "
    strSQL = strSQL & ",CNTA_FLOOR "                                            'INSERT 2018/07/19 K.ISHIZAKA

    Insert_SPAC_TABL_ALONE = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（部屋別）
'       MODULE_ID       : Insert_SPAC_TABL_ALONE_COMM
'       CREATE_DATE     : 2016/08/18            M.HONDA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_SPAC_TABL_ALONE_COMM(ByVal strKomsDb As String) As String
    Dim strSQL              As String

    '設置数
    strSQL = strSQL & " CAST(1 AS NUMERIC(6,0)) AS SPAC_SETTI, "
    '空き数
    'strSQL = strSQL & " CAST(COUNT( CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_AKI,"
    strSQL = strSQL & " CAST(COUNT( CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL AND CNTA_USE = 1 THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_AKI,"
    '差額数
   ' strSQL = strSQL & "  CAST(0 AS NUMERIC(6,0)) AS SPAC_SAGA, "
    
    strSQL = strSQL & " CAST(COUNT(CASE WHEN (INTRT_INTROKBN IS NULL AND CARG_NO IS NULL AND CNTA_PRICE_DIFF <> 0) or (CNTA_REASON LIKE '%柱%') or (CNTA_REASON LIKE '%変則%') THEN  CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_SAGA,"
   
    
    
    '取り置き
    strSQL = strSQL & " CAST(COUNT( CASE WHEN YOUKT_UKNO IS NOT NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_TORI,"
    '解約予定、取り置きされていたら除外
    strSQL = strSQL & " CAST(COUNT( CASE WHEN YOUKT_UKNO IS NULL AND CARG_KYDATE IS NOT NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_KAI,"
    '貸止数
'    strSQL = strSQL & " CAST(COUNT(AKKNT_KAINO) AS NUMERIC(9,0)) AS SPAC_TOME," 'DELETE 2018/12/22 K.ISHIZAKA 'INSERT 2018/11/08 K.ISHIZAKA
                                                                                'INSERT START 2018/12/22 K.ISHIZAKA
    strSQL = strSQL & " CAST(COUNT( CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL AND CNTA_USE = 1 " '空きの状態で
    strSQL = strSQL & " AND CONVERT(varchar,GETDATE(),112) BETWEEN CONVERT(varchar,DATEADD(d,-1,YARD_STOP_FROM),112) AND CONVERT(varchar,YARD_STOP_TO,112) " '今日が貸し止めの期間内で
    strSQL = strSQL & " AND CNTA_USAGE != 31 "      'バイク野外置き場以外で
    strSQL = strSQL & " AND AKKNT_KAINO IS NULL "   '自動鍵でない
    strSQL = strSQL & " THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_TOME,"
                                                                                'INSERT END   2018/12/22 K.ISHIZAKA
    strSQL = strSQL & " CAST('" & strDate & "' AS VARCHAR(8)) AS SPAC_INSED,"
    strSQL = strSQL & " CAST('" & strTime & "' AS VARCHAR(6)) AS SPAC_INSEJ,"
    strSQL = strSQL & " CAST('" & PROG_ID & "' AS VARCHAR(11)) AS SPAC_INSPB,"
    strSQL = strSQL & " CAST('" & strUser & "' AS VARCHAR(8)) AS SPAC_INSUB,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(8)) AS SPAC_UPDAD,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(6)) AS SPAC_UPDAJ,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(11)) AS SPAC_UPDPB,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(8)) AS SPAC_UPDUB "
    strSQL = strSQL & "FROM"
    
    strSQL = strSQL & " ( "                                                     'INSERT START 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " ( "
    strSQL = strSQL & "SELECT TOP 1"
    strSQL = strSQL & " AKKNT_KAINO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.AKKN_TRAN "
    strSQL = strSQL & "WHERE"               '今日が貸し止めの期間内             'DELETE START 2018/12/22 K.ISHIZAKA
'    strSQL = strSQL & " CONVERT(varchar,GETDATE(),112) BETWEEN CONVERT(varchar,DATEADD(d,-1,YARD_STOP_FROM),112) AND CONVERT(varchar,YARD_STOP_TO,112) "
'    strSQL = strSQL & "AND"                 'バイク野外置き場はいつでも利用可能
'    strSQL = strSQL & " CNTA_USAGE != 31 "
'    strSQL = strSQL & "AND"                                                    'DELETE END   2018/12/22 K.ISHIZAKA
    strSQL = strSQL & " AKKNT_YARD = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " AKKNT_CTNO = CNTA_NO "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " AKKNT_HATUD DESC "
    strSQL = strSQL & ") AS AKKNT_KAINO,"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"                                                    'INSERT END   2018/11/08 K.ISHIZAKA
    
    strSQL = strSQL & " " & strKomsDb & ".dbo.YARD_MAST "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.PRIC_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PRIC_YCODE = YARD_CODE "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CNTA_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CNTA_CODE = PRIC_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = PRIC_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_SIZE = PRIC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_STEP = PRIC_STEP "
    '解約予定
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CARG_FILE "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CARG_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_NO = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_CONTNA = 0 "
    '取り置き と受付
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.INTR_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " INTRT_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_NO = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_INTROKBN IN(1, 2) "
    '取り置き
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.YOUK_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " YOUKT_UKNO = INTRT_UKNO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " YOUKT_YUKBN = 2 "
    '表示可能なヤード
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " YARD_HEYA_SHITEI_KBN = -1 "
    strSQL = strSQL & "AND "
    strSQL = strSQL & " ISNULL(YARD_NONDISP_DAY,0) < CONVERT(datetime, '" & strDate & "', 112) "
    '営業中のヤード
    strSQL = strSQL & "AND"
    strSQL = strSQL & " ISNULL(YARD_RENTEND_DAY, CONVERT(datetime, '99991231', 112)) > CONVERT(datetime, '" & strDate & "', 112) "
    '価格が設定されている
    strSQL = strSQL & "AND"
    strSQL = strSQL & " ISNULL(PRIC_PRICE,0) > 0 "
    '利用可能なコンテナ
    '全コンテナを対象
    '2016/08/22 M.HONDA DEL
    'strSQL = strSQL & "AND"
    'strSQL = strSQL & " CNTA_USE = 1 "

    Insert_SPAC_TABL_ALONE_COMM = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（共通部分）
'       MODULE_ID       : Insert_SPAC_TABL_COMM
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_SPAC_TABL_COMM(ByVal strKomsDb As String) As String
    Dim strSQL              As String

    '設置数
    strSQL = strSQL & " CAST(COUNT(DISTINCT CNTA_NO) AS NUMERIC(9,0)) AS SPAC_SETTI,"
    '空き数
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_AKI,"
    ''INSERT 2011/07/18 M.HONDA START
    '差額数
    '20150212 M.HONDA UPD
    'strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL AND CNTA_PRICE_DIFF <> 0 THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_SAGA,"
    strSQL = strSQL & "  CAST(COUNT(DISTINCT CASE WHEN (INTRT_INTROKBN IS NULL AND CARG_NO IS NULL AND CNTA_PRICE_DIFF <> 0) or (CNTA_REASON LIKE '%柱%') or (CNTA_REASON LIKE '%変則%') THEN  CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_SAGA,"
    ''INSERT 2011/07/18 M.HONDA START
    '取り置き
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN YOUKT_UKNO IS NOT NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_TORI,"
    '解約予定、取り置きされていたら除外
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN YOUKT_UKNO IS NULL AND CARG_KYDATE IS NOT NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_KAI,"
    '貸止数
'    strSQL = strSQL & " CAST(COUNT(AKKNT_KAINO) AS NUMERIC(9,0)) AS SPAC_TOME," 'DELETE 2018/12/22 K.ISHIZAKA 'INSERT 2018/11/08 K.ISHIZAKA
                                                                                'INSERT START 2018/12/22 K.ISHIZAKA
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL " '空きの状態で
    strSQL = strSQL & " AND CONVERT(varchar,GETDATE(),112) BETWEEN CONVERT(varchar,DATEADD(d,-1,YARD_STOP_FROM),112) AND CONVERT(varchar,YARD_STOP_TO,112) " '今日が貸し止めの期間内で
    strSQL = strSQL & " AND CNTA_USAGE != 31 "      'バイク野外置き場以外で
    strSQL = strSQL & " AND AKKNT_KAINO IS NULL "   '自動鍵でない
    strSQL = strSQL & " THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_TOME,"
                                                                                'INSERT END   2018/12/22 K.ISHIZAKA
    strSQL = strSQL & " CAST('" & strDate & "' AS VARCHAR(8)) AS SPAC_INSED,"
    strSQL = strSQL & " CAST('" & strTime & "' AS VARCHAR(6)) AS SPAC_INSEJ,"
    strSQL = strSQL & " CAST('" & PROG_ID & "' AS VARCHAR(11)) AS SPAC_INSPB,"
    strSQL = strSQL & " CAST('" & strUser & "' AS VARCHAR(8)) AS SPAC_INSUB,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(8)) AS SPAC_UPDAD,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(6)) AS SPAC_UPDAJ,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(11)) AS SPAC_UPDPB,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(8)) AS SPAC_UPDUB "
    strSQL = strSQL & "FROM"
    
    strSQL = strSQL & " ( "                                                     'INSERT START 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " ( "
    strSQL = strSQL & "SELECT TOP 1"
    strSQL = strSQL & " AKKNT_KAINO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.AKKN_TRAN "
    strSQL = strSQL & "WHERE"               '今日が貸し止めの期間内             'DELETE START 2018/12/22 K.ISHIZAKA
'    strSQL = strSQL & " CONVERT(varchar,GETDATE(),112) BETWEEN CONVERT(varchar,DATEADD(d,-1,YARD_STOP_FROM),112) AND CONVERT(varchar,YARD_STOP_TO,112) "
'    strSQL = strSQL & "AND"                 'バイク野外置き場はいつでも利用可能
'    strSQL = strSQL & " CNTA_USAGE != 31 "
'    strSQL = strSQL & "AND"                                                    'DELETE END   2018/12/22 K.ISHIZAKA
    strSQL = strSQL & " AKKNT_YARD = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " AKKNT_CTNO = CNTA_NO "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " AKKNT_HATUD DESC "
    strSQL = strSQL & ") AS AKKNT_KAINO,"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"                                                    'INSERT END   2018/11/08 K.ISHIZAKA
    
    strSQL = strSQL & " " & strKomsDb & ".dbo.YARD_MAST "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.PRIC_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PRIC_YCODE = YARD_CODE "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CNTA_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CNTA_CODE = PRIC_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = PRIC_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_SIZE = PRIC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_STEP = PRIC_STEP "
    '解約予定
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CARG_FILE "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CARG_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_NO = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_CONTNA = 0 "
    '取り置き と受付(2013/06/13)
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.INTR_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " INTRT_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_NO = CNTA_NO "
    strSQL = strSQL & "AND"
'    strSQL = strSQL & " INTRT_INTROKBN = 1 "                                   'DELETE 2013/06/13 K.ISHIZAKA
    strSQL = strSQL & " INTRT_INTROKBN IN(1, 2) "                               'INSERT 2013/06/13 K.ISHIZAKA
    '取り置き
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.YOUK_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " YOUKT_UKNO = INTRT_UKNO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " YOUKT_YUKBN = 2 "
    '表示可能なヤード
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " (YARD_HEYA_SHITEI_KBN = 0 OR YARD_HEYA_SHITEI_KBN IS NULL ) " 'INSERT 2016/08/18 M.HONDA
    strSQL = strSQL & "AND "
    strSQL = strSQL & " ISNULL(YARD_NONDISP_DAY,0) < CONVERT(datetime, '" & strDate & "', 112) "
    '営業中のヤード
    strSQL = strSQL & "AND"                                                     'INSERT 2011/03/08 K.ISHIZAKA
    strSQL = strSQL & " ISNULL(YARD_RENTEND_DAY, CONVERT(datetime, '99991231', 112)) > CONVERT(datetime, '" & strDate & "', 112) " 'INSERT 2011/03/08 K.ISHIZAKA
    '価格が設定されている
    strSQL = strSQL & "AND"
    strSQL = strSQL & " ISNULL(PRIC_PRICE,0) > 0 "
    '利用可能なコンテナ
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USE = 1 "
'    '受付は除外                                                                'DELETE START 2013/06/13 K.ISHIZAKA
'    strSQL = strSQL & "AND NOT EXISTS"
'    strSQL = strSQL & "("
'    strSQL = strSQL & "SELECT * "
'    strSQL = strSQL & "FROM  " & strKomsDb & ".dbo.INTR_TRAN "
'    strSQL = strSQL & "WHERE INTRT_YCODE    = CNTA_CODE "
'    strSQL = strSQL & "AND   INTRT_NO       = CNTA_NO "
'    strSQL = strSQL & "AND   INTRT_INTROKBN = 2 "
'    strSQL = strSQL & ")"                                                      'DELETE END   2013/06/13 K.ISHIZAKA

    Insert_SPAC_TABL_COMM = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 近隣マスタ登録（データベースを跨る）
'       MODULE_ID       : Insert_NYAR_MAST_by_USAG_MAST
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : objCon                コネクションオブジェクト(I)
'       RETURN          : 件数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_NYAR_MAST_by_USAG_MAST(objCon As Object) As Long
    Dim lngCnt              As Long
    Dim objRst1             As Object
    Dim objRst2             As Object
    Dim dblKiro             As Double
    On Error GoTo ErrorHandler
    
    Call ADODB_DropTable(C_WORK_NYAR1, objCon)
    Call ADODB_DropTable(C_WORK_NYAR2, objCon)
    lngCnt = ADODB_Execute(Create_WORK_NYAR_MAST1(), objCon)
    lngCnt = ADODB_Execute(Create_WORK_NYAR_MAST2(), objCon)
    
    Set objRst1 = ADODB_Recordset(Select_WORK_NYAR_MAST1(), objCon)
    On Error GoTo ErrorHandler1
    Set objRst2 = ADODB_Recordset(C_WORK_NYAR2, objCon, adoAppendOnly)
    On Error GoTo ErrorHandler2
    
    With objRst1
        While Not .EOF
            dblKiro = KmFromDo(.Fields("BASE_ZAHYOKEI"), _
                .Fields("BASE_IDO"), .Fields("BASE_KEIDO"), _
                .Fields("NYAR_IDO"), .Fields("NYAR_KEIDO"))
            If dblKiro <= .Fields("USAG_KIRO") Then
                objRst2.AddNew
                On Error GoTo ErrorHandler3
                objRst2.Fields("NYAR_YCODE") = .Fields("NYAR_YCODE")
                objRst2.Fields("NYAR_NCODE") = .Fields("NYAR_NCODE")
                objRst2.Fields("NYAR_KIRO") = dblKiro
                objRst2.UPDATE
                On Error GoTo ErrorHandler2
            End If
            .MoveNext
        Wend
        objRst2.Close
        On Error GoTo ErrorHandler1
        .Close
        On Error GoTo ErrorHandler
    End With

    lngCnt = 0
    lngCnt = lngCnt + ADODB_Execute(Insert_NYAR_MAST_Work("NYAR_YCODE", "NYAR_NCODE"), objCon)
    lngCnt = lngCnt + ADODB_Execute(Insert_NYAR_MAST_Work("NYAR_NCODE", "NYAR_YCODE"), objCon)
    
    Call ADODB_DropTable(C_WORK_NYAR2, objCon)
    Call ADODB_DropTable(C_WORK_NYAR1, objCon)

    Insert_NYAR_MAST_by_USAG_MAST = lngCnt
Exit Function

ErrorHandler3:
    objRst2.CancelUpdate
ErrorHandler2:
    objRst2.Close
ErrorHandler1:
    objRst1.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Insert_NYAR_MAST_by_USAG_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : データベースを跨る近隣ヤード抽出
'       MODULE_ID       : Create_WORK_NYAR_MAST1
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Create_WORK_NYAR_MAST1() As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " base.YARD_CODE AS NYAR_YCODE,"
    strSQL = strSQL & " base.YARD_USAGE AS NYAR_USAGE,"
    strSQL = strSQL & " nyar.YARD_CODE AS NYAR_NCODE,"
    strSQL = strSQL & " nyar.YARD_USAGE AS NYAR_NSAGE,"
    strSQL = strSQL & " base.YARD_ZAHYOKEI AS BASE_ZAHYOKEI,"
    strSQL = strSQL & " base.YARD_IDO AS BASE_IDO,"
    strSQL = strSQL & " base.YARD_KEIDO AS BASE_KEIDO,"
    strSQL = strSQL & " nyar.YARD_IDO AS NYAR_IDO,"
    strSQL = strSQL & " nyar.YARD_KEIDO AS NYAR_KEIDO,"
    strSQL = strSQL & " USAG_KIRO "
    strSQL = strSQL & "INTO"
    strSQL = strSQL & " " & C_WORK_NYAR1 & " "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " TOGO_YARD_MAST base "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_USAG_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " USAG_USAGE = base.YARD_USAGE "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_YARD_MAST nyar "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " USAG_NSAGE = nyar.YARD_USAGE "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " base.YARD_BUMOC != nyar.YARD_BUMOC "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " nyar.YARD_IDO BETWEEN base.YARD_IDO - (1.0 / 110.0 * USAG_KIRO) AND base.YARD_IDO + (1.0 / 110.0 * USAG_KIRO) "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " nyar.YARD_KEIDO BETWEEN base.YARD_KEIDO - (1.0 / 90.0 * USAG_KIRO) AND base.YARD_KEIDO + (1.0 / 90.0 * USAG_KIRO) "

    Create_WORK_NYAR_MAST1 = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 距離算出ワークテーブル作成
'       MODULE_ID       : Select_WORK_NYAR_MAST1
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Create_WORK_NYAR_MAST2() As String
    Dim strSQL          As String

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " TOP 0 "
    strSQL = strSQL & " NYAR_YCODE,"
    strSQL = strSQL & " NYAR_NCODE,"
    strSQL = strSQL & " USAG_KIRO AS NYAR_KIRO "
    strSQL = strSQL & "INTO"
    strSQL = strSQL & " " & C_WORK_NYAR2 & " "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & C_WORK_NYAR1 & " "

    Create_WORK_NYAR_MAST2 = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 距離算出対象データ抽出
'       MODULE_ID       : Select_WORK_NYAR_MAST1
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Select_WORK_NYAR_MAST1() As String
    Dim strSQL          As String

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " DISTINCT"
    strSQL = strSQL & " NYAR_YCODE,"
    strSQL = strSQL & " NYAR_NCODE,"
    strSQL = strSQL & " BASE_ZAHYOKEI,"
    strSQL = strSQL & " BASE_IDO,"
    strSQL = strSQL & " BASE_KEIDO,"
    strSQL = strSQL & " NYAR_IDO,"
    strSQL = strSQL & " NYAR_KEIDO,"
    strSQL = strSQL & " USAG_KIRO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & C_WORK_NYAR1 & " par "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " NYAR_YCODE < NYAR_NCODE "
    strSQL = strSQL & "OR"
    strSQL = strSQL & "("
    strSQL = strSQL & " NYAR_YCODE > NYAR_NCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " NOT EXISTS"
    strSQL = strSQL & " ("
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & "  *"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & "  " & C_WORK_NYAR1 & " ch"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "  ch.NYAR_YCODE = par.NYAR_NCODE"
    strSQL = strSQL & " AND"
    strSQL = strSQL & "  ch.NYAR_NCODE = par.NYAR_YCODE"
    strSQL = strSQL & " )"
    strSQL = strSQL & ")"

    Select_WORK_NYAR_MAST1 = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 近隣マスタ登録
'       MODULE_ID       : Insert_NYAR_MAST_Work
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strYocde              ヤードコード１(I)
'                       : strNocde              ヤードコード２(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_NYAR_MAST_Work(ByVal strYcode As String, ByVal strNcode As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "INSERT INTO TOGO_NYAR_MAST "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " wk1.NYAR_YCODE,"
    strSQL = strSQL & " wk1.NYAR_USAGE,"
    strSQL = strSQL & " wk1.NYAR_NCODE,"
    strSQL = strSQL & " wk1.NYAR_NSAGE,"
    strSQL = strSQL & " CAST('" & strDate & "' AS VARCHAR(8)) AS NYAR_INSED,"
    strSQL = strSQL & " CAST('" & strTime & "' AS VARCHAR(6)) AS NYAR_INSEJ,"
    strSQL = strSQL & " CAST('" & PROG_ID & "' AS VARCHAR(11)) AS NYAR_INSPB,"
    strSQL = strSQL & " CAST('" & strUser & "' AS VARCHAR(8)) AS NYAR_INSUB,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(8)) AS NYAR_UPDAD,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(6)) AS NYAR_UPDAJ,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(11)) AS NYAR_UPDPB,"
    strSQL = strSQL & " CAST(NULL AS VARCHAR(8)) AS NYAR_UPDUB,"
    strSQL = strSQL & " wk2.NYAR_KIRO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & C_WORK_NYAR1 & " wk1 "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " " & C_WORK_NYAR2 & " wk2 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " wk1.NYAR_YCODE = wk2." & strYcode & " "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " wk1.NYAR_NCODE = wk2." & strNcode & " "

    Insert_NYAR_MAST_Work = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 統合ＤＢの内容を最新状態にする
'       MODULE_ID       : MSTGP01_M00
'       CREATE_DATE     : 2011/02/06            K.ISHIZAKA
'       PARAM           : strYCODE              ヤードコード(I)
'                       : [strNO]               コンテナ番号(I)
'       RETURN          : 正常(True)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function MSTGP01_M10(ByVal strYcode As String, Optional ByVal strNO As String = "") As Boolean 'DELETE 2011/03/12 K.ISHIZAKA
Public Function MSTGP01_M10(ByVal strYcode As String, Optional ByVal strNO As String = "", _
    Optional ByVal strBUMOC As String = "") As Boolean                          'INSERT 2011/03/12 K.ISHIZAKA
    Dim lngCnt              As Long
    Dim objCon              As Object
    Dim objRst              As Object
    Dim strKomsDb           As String
    Dim strWhere            As String
    Dim strUsage            As String
    Dim strStep             As String
    Dim strSQL              As String
    Dim blWebUpdate         As Boolean                                          'INSERT 2012/06/26 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    strUser = Left(LsGetUserName(), 8)
    
    blWebUpdate = CBool(Nz(DLookup("INTIF_RECFB", "INTI_FILE", "INTIF_PROGB = 'FTG011'"), "False"))
    
    If strBUMOC = "" Then
        strBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    End If
     
    strKomsDb = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME_" & strBUMOC & "'")
    Set objCon = ADODB_Connection(C_TOGO_CONNENCT_BUMOC)
    On Error GoTo ErrorHandler1
    
    strWhere = "AND CNTA_CODE = " & strYcode & " "
    'コンテナ番号指定の時は、用途、サイズ、段を取得
    If strNO <> "" Then
        strSQL = strSQL & "SELECT"
        strSQL = strSQL & " CNTA_USAGE,"
        strSQL = strSQL & " CNTA_SIZE,"
        strSQL = strSQL & " CNTA_STEP "
        strSQL = strSQL & "FROM"
        strSQL = strSQL & " " & strKomsDb & ".dbo.CNTA_MAST "
        strSQL = strSQL & "WHERE"
        strSQL = strSQL & " CNTA_CODE = " & strYcode & " "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " CNTA_NO = " & strNO & " "
        
        Set objRst = ADODB_Recordset(strSQL, objCon)
        On Error GoTo ErrorHandler2
        
        With objRst
            If Not .EOF Then
                strUsage = .Fields("CNTA_USAGE")
                strStep = .Fields("CNTA_STEP")
                If CheckUsage(strUsage, C_USAGE_GRP_ALONE) Then
                    strWhere = strWhere & "AND CNTA_NO = " & strNO & " "
                Else
                    strWhere = strWhere & "AND CNTA_USAGE = " & .Fields("CNTA_USAGE") & " "
                    strWhere = strWhere & "AND CNTA_SIZE = " & .Fields("CNTA_SIZE") & " "
                    strWhere = strWhere & "AND CNTA_STEP = " & .Fields("CNTA_STEP") & " "
                End If
            End If
            .Close
        End With
        On Error GoTo ErrorHandler1
        If blWebUpdate Then
            lngCnt = ADODB_Execute(Befoer_Update_SPAC_TABL(strWhere), objCon)
        End If
    Else
        blWebUpdate = False
    End If

    '空き状況（サイズ別）
    If strNO = "" Then
         If strUsage = "" Or CheckUsage(strUsage, C_USAGE_GRP_COUNT & ",39") Then
            lngCnt = ADODB_Execute(Update_SPAC_TABL_COUNT(strKomsDb, strWhere), objCon)
        End If
    End If
'    '空き状況（コンテナバイク）
'    If strUsage = "" Or strUsage = "0" Then
'        lngCnt = ADODB_Execute(Update_SPAC_TABL_KBIKE(strKomsDb, strWhere), objCon)
'    End If
    '空き状況（部屋別）
    If Trim(C_USAGE_GRP_ALONE) <> "" Then
        If strUsage = "" Or CheckUsage(strUsage, C_USAGE_GRP_ALONE) Then
            lngCnt = ADODB_Execute(Update_SPAC_TABL_ALONE(strKomsDb, strWhere), objCon)
        End If
    End If
    '未確定数
    lngCnt = ADODB_Execute(Update_YARD_MAST(strKomsDb, strYcode, strUsage), objCon)
    
'    If blWebUpdate Then                                                     'INSERT START 2012/06/26 K.ISHIZAKA
'        Set objRst = ADODB_Recordset(After_Update_SPAC_TABL(), objCon)
'        On Error GoTo ErrorHandler2
'        With objRst
'            While Not .EOF
'                strSQL = "SELECT CountDown(" & Format(.Fields("SPAC_YCODE").VALUE) & ","
'                strSQL = strSQL & Format(.Fields("SPAC_USAGE").VALUE) & ","
'                strSQL = strSQL & Format(.Fields("SPAC_SIZE").VALUE) & ","
'                strSQL = strSQL & Format(.Fields("SPAC_STEP").VALUE) & ","
'                strSQL = strSQL & Format(.Fields("SPAC_NO").VALUE) & ");"
'                'カウントダウンできなくても、ここでは無視する
'                lngCnt = MySql_ExecGetLong(strSQL)
'                .MoveNext
'            Wend
'            .Close
''        End With
'        On Error GoTo ErrorHandler1
'    End If                                                                  'INSERT END   2012/06/26 K.ISHIZAKA
    
    objCon.Close
    On Error GoTo ErrorHandler
    MSTGP01_M10 = True
Exit Function

ErrorHandler2:
    objRst.Close
ErrorHandler1:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call MSZZ024_M00("MSTGP01_M10", False)
    MSTGP01_M10 = False
End Function

'==============================================================================*
'
'       MODULE_NAME     : カンマ区切りの用途区分に存在するかチェックする
'       MODULE_ID       : CheckUsage
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strUsage              用途(I)
'                       : strGroup              用途グループ(I)カンマ区切りで複数指定可
'       RETURN          : 存在する(True)／しない(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CheckUsage(ByVal strUsage As String, ByVal strGroup As String) As Boolean
    Dim varUsage            As Variant
    On Error GoTo ErrorHandler
    
    For Each varUsage In Split(strGroup, ",")
        If varUsage = strUsage Then
            CheckUsage = True
            Exit Function
        End If
    Next
    CheckUsage = False
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "CheckUsage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（サイズ別）
'       MODULE_ID       : Update_SPAC_TABL_COUNT
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'                       : strWhere              条件(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Update_SPAC_TABL_COUNT(ByVal strKomsDb As String, ByVal strWhere As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "UPDATE"
    'strSQL = strSQL & " TOGO_SPAC_TABL "
    strSQL = strSQL & " TOGO_SPAC_TABL_J "                                      'INSERT 2026/04/22 N.IMAI
    strSQL = strSQL & "SET"
    strSQL = strSQL & " SPAC_SETTI = CNTA_SETTI,"
    strSQL = strSQL & " SPAC_AKI = CNTA_AKI,"
    strSQL = strSQL & " SPAC_TORI = CNTA_TORI,"
    strSQL = strSQL & " SPAC_KAI = CNTA_KAI,"
    strSQL = strSQL & " SPAC_TOME = CNTA_TOME,"                                                'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & " SPAC_UPDAD = '" & strDate & "',"
    strSQL = strSQL & " SPAC_UPDAJ = '" & strTime & "',"
    strSQL = strSQL & " SPAC_UPDPB = '" & PROG_ID & "',"
    strSQL = strSQL & " SPAC_UPDUB = '" & strUser & "' "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP,"
    
    strSQL = strSQL & Update_SPAC_TABL_COMM(strKomsDb) & strWhere
    
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE IN (" & C_USAGE_GRP_COUNT & ") "
    strSQL = strSQL & ") AS wk "                                                'INSERT 2018/11/08 K.ISHIZAKA
    
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP "
    strSQL = strSQL & ") cnta "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CNTA_CODE = SPAC_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = SPAC_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_SIZE = SPAC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_STEP = SPAC_STEP "

    Update_SPAC_TABL_COUNT = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（コンテナバイク）
'       MODULE_ID       : Update_SPAC_TABL_KBIKE
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'                       : strWhere              条件(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Update_SPAC_TABL_KBIKE(ByVal strKomsDb As String, ByVal strWhere As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "UPDATE"
    'strSQL = strSQL & " TOGO_SPAC_TABL "
    strSQL = strSQL & " TOGO_SPAC_TABL_J "                                      'INSERT 2026/04/22 N.IMAI
    strSQL = strSQL & "SET"
    strSQL = strSQL & " SPAC_SETTI = CNTA_SETTI,"
    strSQL = strSQL & " SPAC_AKI = CNTA_AKI,"
    strSQL = strSQL & " SPAC_TORI = CNTA_TORI,"
    strSQL = strSQL & " SPAC_KAI = CNTA_KAI,"
    strSQL = strSQL & " SPAC_TOME = CNTA_TOME,"                                                'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & " SPAC_UPDAD = '" & strDate & "',"
    strSQL = strSQL & " SPAC_UPDAJ = '" & strTime & "',"
    strSQL = strSQL & " SPAC_UPDPB = '" & PROG_ID & "',"
    strSQL = strSQL & " SPAC_UPDUB = '" & strUser & "' "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " " & C_USAGE_GRP_KBIKE & " AS CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP,"
    
    strSQL = strSQL & Update_SPAC_TABL_COMM(strKomsDb) & strWhere
    
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = 0 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_BIKE_OK = 1 "
    strSQL = strSQL & ") AS wk "                                                'INSERT 2018/11/08 K.ISHIZAKA
    
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP "
    strSQL = strSQL & ") cnta "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CNTA_CODE = SPAC_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = SPAC_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_SIZE = SPAC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_STEP = SPAC_STEP "

    Update_SPAC_TABL_KBIKE = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（部屋別）
'       MODULE_ID       : Update_SPAC_TABL_ALONE
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'                       : strWhere              条件(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Update_SPAC_TABL_ALONE(ByVal strKomsDb As String, ByVal strWhere As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & " TOGO_SPAC_TABL_J "
    strSQL = strSQL & "SET"
    strSQL = strSQL & " SPAC_SETTI = CNTA_SETTI,"
    strSQL = strSQL & " SPAC_AKI = CNTA_AKI,"
    strSQL = strSQL & " SPAC_TORI = CNTA_TORI,"
    strSQL = strSQL & " SPAC_KAI = CNTA_KAI,"
    strSQL = strSQL & " SPAC_TOME = CNTA_TOME,"                                                'INSERT 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & " SPAC_UPDAD = '" & strDate & "',"
    strSQL = strSQL & " SPAC_UPDAJ = '" & strTime & "',"
    strSQL = strSQL & " SPAC_UPDPB = '" & PROG_ID & "',"
    strSQL = strSQL & " SPAC_UPDUB = '" & strUser & "' "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP,"
    strSQL = strSQL & " CNTA_NO,"
    
    strSQL = strSQL & Update_SPAC_TABL_COMM(strKomsDb) & strWhere
    
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE IN (" & C_USAGE_GRP_ALONE & ") "
    strSQL = strSQL & ") AS wk "                                                'INSERT 2018/11/08 K.ISHIZAKA
    
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " CNTA_CODE,"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP,"
    strSQL = strSQL & " CNTA_NO "
    strSQL = strSQL & ") cnta "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CNTA_CODE = SPAC_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_USAGE = SPAC_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_SIZE = SPAC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_STEP = SPAC_STEP "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_NO = SPAC_NO "

    Update_SPAC_TABL_ALONE = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 空き状況登録（共通部分）
'       MODULE_ID       : Update_SPAC_TABL_COMM
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Update_SPAC_TABL_COMM(ByVal strKomsDb As String) As String
    Dim strSQL              As String

    '設置数
    strSQL = strSQL & " CAST(COUNT(DISTINCT CNTA_NO) AS NUMERIC(9,0)) AS CNTA_SETTI,"
    '空き数
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS CNTA_AKI,"
    ''INSERT 2011/07/18 M.HONDA START
    '差額数
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN INTRT_INTROKBN IS NULL AND CARG_NO IS NULL AND CNTA_PRICE_DIFF <> 0 THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS SPAC_SAGA,"
    ''INSERT 2011/07/18 M.HONDA START
    '取り置き
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN YOUKT_UKNO IS NOT NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS CNTA_TORI,"
    '貸止数
    strSQL = strSQL & " CAST(COUNT(AKKNT_KAINO) AS NUMERIC(9,0)) AS CNTA_TOME," 'INSERT 2018/11/08 K.ISHIZAKA
    
    '解約予定、取り置きされていたら除外
    strSQL = strSQL & " CAST(COUNT(DISTINCT CASE WHEN YOUKT_UKNO IS NULL AND CARG_KYDATE IS NOT NULL THEN CNTA_NO ELSE NULL END) AS NUMERIC(9,0)) AS CNTA_KAI "
    strSQL = strSQL & "FROM"
    
    strSQL = strSQL & " ( "                                                     'INSERT START 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " ( "
    strSQL = strSQL & "SELECT TOP 1"
    strSQL = strSQL & " AKKNT_KAINO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.AKKN_TRAN "
    strSQL = strSQL & "WHERE"               '今日が貸し止めの期間内
    strSQL = strSQL & " CONVERT(varchar,GETDATE(),112) BETWEEN CONVERT(varchar,DATEADD(d,-1,YARD_STOP_FROM),112) AND CONVERT(varchar,YARD_STOP_TO,112) "
    strSQL = strSQL & "AND"                 'バイク野外置き場はいつでも利用可能
    strSQL = strSQL & " CNTA_USAGE != 31 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " AKKNT_YARD = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " AKKNT_CTNO = CNTA_NO "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " AKKNT_HATUD DESC "
    strSQL = strSQL & ") AS AKKNT_KAINO,"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"                                                    'INSERT END   2018/11/08 K.ISHIZAKA
    
    strSQL = strSQL & " " & strKomsDb & ".dbo.CNTA_MAST "
    strSQL = strSQL & "INNER JOIN"                                              'INSERT START 2018/11/08 K.ISHIZAKA
    strSQL = strSQL & " " & strKomsDb & ".dbo.YARD_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CNTA_CODE = YARD_CODE "                                 'INSERT END   2018/11/08 K.ISHIZAKA
    '解約予定
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CARG_FILE "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CARG_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_NO = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CARG_CONTNA = 0 "
    '取り置き
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.INTR_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " INTRT_YCODE = CNTA_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_NO = CNTA_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " INTRT_INTROKBN IN ('1','2') "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " " & strKomsDb & ".dbo.YOUK_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " YOUKT_UKNO = INTRT_UKNO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " YOUKT_YUKBN = 2 "
    '利用可能なコンテナ
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CNTA_USE = 1 "
    '受付は除外
'    strSQL = strSQL & "AND NOT EXISTS"
'    strSQL = strSQL & "("
'    strSQL = strSQL & "SELECT *"
'    strSQL = strSQL & "FROM  " & strKomsDb & ".dbo.INTR_TRAN "
'    strSQL = strSQL & "WHERE INTRT_YCODE    = CNTA_CODE "
'    strSQL = strSQL & "AND   INTRT_NO       = CNTA_NO "
'    strSQL = strSQL & "AND   INTRT_INTROKBN = 2 "
'    strSQL = strSQL & ") "

    Update_SPAC_TABL_COMM = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 未確定数更新
'       MODULE_ID       : Update_YARD_MAST
'       CREATE_DATE     : 2011/01/24            K.ISHIZAKA
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'                       : strYcode              ヤードコード(I)
'                       : strUsage              用途(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Update_YARD_MAST(ByVal strKomsDb As String, ByVal strYcode As String, ByVal strUsage As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & " TOGO_YARD_MAST "
    strSQL = strSQL & "SET"
    strSQL = strSQL & " YARD_MIKAKUTEI = "
    strSQL = strSQL & "(SELECT COUNT(*)"
    strSQL = strSQL & " FROM   " & strKomsDb & ".dbo.YOUK_TRAN"
    strSQL = strSQL & " WHERE  YOUKT_YCODE = YARD_CODE"
    strSQL = strSQL & " AND    YOUKT_YUKBN = 0"
    strSQL = strSQL & " AND ( (ISNULL(YOUKT_USAGE, YARD_USAGE) = YARD_USAGE)"
    strSQL = strSQL & "    OR (YOUKT_USAGE = 10 AND YARD_USAGE IN(0,1))"
    strSQL = strSQL & "    OR (YOUKT_USAGE = 30 AND YARD_USAGE IN(3,31,32,33))"
    strSQL = strSQL & "     )"
    strSQL = strSQL & ")"
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " YARD_CODE = " & strYcode & " "
    If strUsage <> "" Then
        strSQL = strSQL & "AND"
        strSQL = strSQL & " YARD_USAGE = " & strUsage & " "
    End If

    Update_YARD_MAST = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : ログインテーブル
'       MODULE_ID       : Insert_LOGIN_TABL
'       CREATE_DATE     :
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'                       : strBUMOC              部門ｺｰﾄﾞ
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_LOGIN_TABL(ByVal strKomsDb As String, strBUMOC As Variant) As String
    Dim strSQL              As String

    strSQL = "INSERT INTO TOGO_LOGIN_TABL "
    strSQL = strSQL & "SELECT " & "'" & strBUMOC & "',"
    strSQL = strSQL & "       RCPT_UCODE, "
    strSQL = strSQL & "       RCPT_YCODE, "
    strSQL = strSQL & "       RCPT_CNO, "
    strSQL = strSQL & "       ISNULL(CARG_AGRE,0) "
    strSQL = strSQL & "FROM ( "
    strSQL = strSQL & "SELECT RCPT_UCODE, "
    strSQL = strSQL & "       RCPT_YCODE, "
    strSQL = strSQL & "       RCPT_CNO, "
    strSQL = strSQL & "       CARG_AGRE, "
    strSQL = strSQL & "       CARG_KYDATE, "
    strSQL = strSQL & "       CARG_HOSYO_CD, "
    strSQL = strSQL & "       YOUKT_YUKBN "
    strSQL = strSQL & "FROM " & strKomsDb & ".dbo.RCPT_TRAN "
    strSQL = strSQL & "       LEFT JOIN " & strKomsDb & ".dbo.CARG_FILE ON "
    strSQL = strSQL & "       RCPT_CARG_ACPTNO = CARG_ACPTNO "
    strSQL = strSQL & "       LEFT JOIN " & strKomsDb & ".dbo.YOUK_TRAN ON "
    strSQL = strSQL & "       RCPT_NO = YOUKT_UKNO  ) A "
'    strSQL = strSQL & "WHERE a.CARG_KYDATE IS NULL "                                        'DELETE 2021/10/29 N.IMAI
    strSQL = strSQL & "WHERE ISNULL(a.CARG_KYDATE,'2100-12-31 00:00:00.000') >= GETDATE() "  'INSERT 2021/10/29 N.IMAI
    strSQL = strSQL & "AND   A.YOUKT_YUKBN IN (10,20) "
    strSQL = strSQL & "AND   A.CARG_HOSYO_CD NOT IN ('970033','970031') "
    strSQL = strSQL & "ORDER BY RCPT_UCODE "
    
    Insert_LOGIN_TABL = strSQL

End Function

'==============================================================================*
'
'       MODULE_NAME     : 顧客マスタ
'       MODULE_ID       : Insert_KOKY_MAST
'       CREATE_DATE     :
'       PARAM           : strKomsDb             KOMSＤＢ名(I)
'                       : strBUMOC              部門ｺｰﾄﾞ
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_KOKY_MAST(ByVal strKomsDb As String, strBUMOC As Variant) As String
    Dim strSQL              As String

    strSQL = "INSERT INTO TOGO_KOKY_MAST "
    strSQL = strSQL & " SELECT  DISTINCT  " & "'" & strBUMOC & "', "
    strSQL = strSQL & "       USER_CODE , "
    strSQL = strSQL & "       USER_NAME, "
    strSQL = strSQL & "       USER_KANA, "
    strSQL = strSQL & "       USER_SKBN, "
    strSQL = strSQL & "       USER_INSED, "
    strSQL = strSQL & "       USER_INSEJ, "
    strSQL = strSQL & "       USER_INSPB, "
    strSQL = strSQL & "       USER_INSUB, "
    strSQL = strSQL & "       USER_UPDAD, "
    strSQL = strSQL & "       USER_UPDAJ, "
    strSQL = strSQL & "       USER_UPDPB, "
    strSQL = strSQL & "       USER_UPDUB "
    strSQL = strSQL & "       FROM ( "
    strSQL = strSQL & "             SELECT RCPT_UCODE, "
    strSQL = strSQL & "                    RCPT_YCODE, "
    strSQL = strSQL & "                    RCPT_CNO, "
    strSQL = strSQL & "                    CARG_AGRE, "
    strSQL = strSQL & "                    CARG_KYDATE, "
    strSQL = strSQL & "                    CARG_HOSYO_CD, "
    strSQL = strSQL & "                    YOUKT_YUKBN "
    strSQL = strSQL & "               FROM  " & strKomsDb & ".dbo.RCPT_TRAN "
    strSQL = strSQL & "                    LEFT JOIN  " & strKomsDb & ".dbo.CARG_FILE ON "
    strSQL = strSQL & "                               RCPT_CARG_ACPTNO = CARG_ACPTNO "
    strSQL = strSQL & "                    LEFT JOIN  " & strKomsDb & ".dbo.YOUK_TRAN ON "
    strSQL = strSQL & "                               RCPT_NO = YOUKT_UKNO  ) A, "
    strSQL = strSQL & "            " & strKomsDb & ".dbo.USER_MAST "
'    strSQL = strSQL & "       WHERE a.CARG_KYDATE IS NULL "                                        'DELETE 2021/10/29 N.IMAI
    strSQL = strSQL & "       WHERE ISNULL(a.CARG_KYDATE,'2100-12-31 00:00:00.000') >= GETDATE() "  'INSERT 2021/10/29 N.IMAI
    strSQL = strSQL & "       AND   A.YOUKT_YUKBN IN (10,20) "
    strSQL = strSQL & "       AND   RCPT_UCODE = USER_CODE "
    strSQL = strSQL & "       AND   A.CARG_HOSYO_CD NOT IN ('970033','970031') "
    
    Insert_KOKY_MAST = strSQL

End Function

'==============================================================================*
'
'       MODULE_NAME     : ログインテーブル(from KASEDB)
'       MODULE_ID       : Insert_LOGIN_TABL_from_KASEDB
'       CREATE_DATE     :
'       PARAM           : strKomsDb             KASEDB名
'                       : strBUMOC              部門ｺｰﾄﾞ
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_LOGIN_TABL_from_KASEDB(ByVal strKomsDb As String, strBUMOC As Variant, ByVal strNowMonthBeforeLast As String) As String
    Dim strSQL              As String

    strSQL = "INSERT INTO TOGO_LOGIN_TABL "
    strSQL = strSQL & "select KEIYT_BUMOC, "
    strSQL = strSQL & "KEIYT_KOKYC, "
    strSQL = strSQL & "KEIYT_BUKHC, "
    strSQL = strSQL & "KEIYT_NUMBC, "
    strSQL = strSQL & "ISNULL(KEIYT_STATI,0) "
    strSQL = strSQL & "from " & strKomsDb & ".dbo.KEIY_TRAN "
    strSQL = strSQL & "where KEIYT_BUMOC = '" & strBUMOC & "'"
    strSQL = strSQL & "AND ISNULL(KEIYT_KAIYD,'29990101') > '" & strNowMonthBeforeLast & "' "
    strSQL = strSQL & "AND KEIYT_SESSI = 1 "
    strSQL = strSQL & "AND KEIYT_STATI <> 10 "
    strSQL = strSQL & "AND ISNULL(KEIYT_HOKAC,'') NOT IN ('970152','970090','970180','970040','970041','970042','970043') "
    strSQL = strSQL & "AND  KEIYT_BUKHC <> '001940' "
    
    strSQL = strSQL & "UNION  "
    strSQL = strSQL & "select KEIYT_BUMOC, "
    strSQL = strSQL & "KEIYT_KOKYC, "
    strSQL = strSQL & "KEIYT_SYO1C, "
    strSQL = strSQL & "KEIYT_NUMBC, "
    strSQL = strSQL & "ISNULL(KEIYT_STATI,0) "
    strSQL = strSQL & "from " & strKomsDb & ".dbo.KEIY_TRAN "
    strSQL = strSQL & "where KEIYT_BUMOC = '" & strBUMOC & "'"
    strSQL = strSQL & "AND ISNULL(KEIYT_KAIYD,'29990101') > '" & strNowMonthBeforeLast & "' "
    strSQL = strSQL & "AND KEIYT_SESSI = 1 "
    strSQL = strSQL & "AND KEIYT_STATI <> 10 "
    strSQL = strSQL & "AND ISNULL(KEIYT_HOKAC,'') NOT IN ('970152','970090','970180','970040','970041','970042','970043') "
    strSQL = strSQL & "AND KEIYT_SYO1C <> '001940' "
    strSQL = strSQL & "order by KEIYT_BUMOC "
    Insert_LOGIN_TABL_from_KASEDB = strSQL
    
    'Debug.Print strSQL

End Function

'==============================================================================*
'
'       MODULE_NAME     : 顧客マスタ
'       MODULE_ID       : Insert_KOKY_MAST_from_KASEDB
'       CREATE_DATE     :
'       PARAM           : strKomsDb             KASEDB名
'                       : strBUMOC              部門ｺｰﾄﾞ
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_KOKY_MAST_from_KASEDB(ByVal strKomsDb As String, strBUMOC As Variant, ByVal strNowMonthBeforeLast As String) As String
    Dim strSQL              As String

    strSQL = "INSERT INTO TOGO_KOKY_MAST "
    strSQL = strSQL & " select DISTINCT  "
    strSQL = strSQL & " KOKYM_BUMOC, "
    strSQL = strSQL & " KOKYM_KOKYC, "
    strSQL = strSQL & " LEFT(KOKYM_KOKYN,20), "
    strSQL = strSQL & " LEFT(KOKYM_KOKYF,20), "
    strSQL = strSQL & " KOKYM_SKBNI, "
    strSQL = strSQL & " KOKYM_INSED, "
    strSQL = strSQL & " KOKYM_INSEJ, "
    strSQL = strSQL & " KOKYM_INSPB, "
    strSQL = strSQL & " KOKYM_INSUB, "
    strSQL = strSQL & " KOKYM_UPDAD, "
    strSQL = strSQL & " KOKYM_UPDAJ, "
    strSQL = strSQL & " KOKYM_UPDPB, "
    strSQL = strSQL & " KOKYM_UPDUB "
    strSQL = strSQL & " from " & strKomsDb & ".dbo.KEIY_TRAN "
    strSQL = strSQL & " INNER JOIN " & strKomsDb & ".dbo.KOKY_MAST ON "
    strSQL = strSQL & " KOKYM_BUMOC = KEIYT_BUMOC AND "
    strSQL = strSQL & " KOKYM_KOKYC = KEIYT_KOKYC "
    strSQL = strSQL & " where KEIYT_BUMOC = '" & strBUMOC & "'"
    strSQL = strSQL & " AND ISNULL(KEIYT_KAIYD,'29990101') > '" & strNowMonthBeforeLast & "' "
    strSQL = strSQL & " AND KEIYT_SESSI = 1 "
    strSQL = strSQL & " AND KEIYT_STATI <> 10 "
    strSQL = strSQL & " AND ISNULL(KEIYT_HOKAC,'') NOT IN ('970152','970090','970180','970040','970041','970042','970043') "
    strSQL = strSQL & " AND  KEIYT_BUKHC <> '001940' "
    Insert_KOKY_MAST_from_KASEDB = strSQL
    
    'Debug.Print strSQL
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : 更新前空き数
'       MODULE_ID       : Befoer_Update_SPAC_TABL
'       CREATE_DATE     : 2012/06/26            K.ISHIZAKA
'       PARAM           : strWhere              更新条件(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Befoer_Update_SPAC_TABL(ByVal strWhere As String) As String
    Dim strSQL              As String

    strWhere = Replace(strWhere, "AND CNTA_CODE ", "WHERE SPAC_YCODE ")
    strWhere = Replace(strWhere, "AND CNTA_", "AND SPAC_")
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " SPAC_YCODE AS BF_YCODE,"
    strSQL = strSQL & " SPAC_USAGE AS BF_USAGE,"
    strSQL = strSQL & " SPAC_SIZE AS BF_SIZE,"
    strSQL = strSQL & " SPAC_STEP AS BF_STEP,"
    strSQL = strSQL & " SPAC_NO  AS BF_NO,"
    strSQL = strSQL & " SPAC_AKI AS BF_AKI "
    strSQL = strSQL & "INTO #BF_SPAC "
    '20150212 M.HONDA
    'strSQL = strSQL & "FROM SPAC_TABL "
    strSQL = strSQL & "FROM TOGO_SPAC_TABL "
    strSQL = strSQL & strWhere
    
    Befoer_Update_SPAC_TABL = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 更新後空き数
'       MODULE_ID       : After_Update_SPAC_TABL
'       CREATE_DATE     : 2012/06/26            K.ISHIZAKA
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function After_Update_SPAC_TABL() As String
    Dim strSQL              As String

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " SPAC_YCODE,"
    strSQL = strSQL & " SPAC_USAGE,"
    strSQL = strSQL & " SPAC_SIZE,"
    strSQL = strSQL & " SPAC_STEP,"
    strSQL = strSQL & " SPAC_NO "
    strSQL = strSQL & "FROM #BF_SPAC "
    '20150212 M.HONDA
    'strSQL = strSQL & "INNER JOIN SPAC_TABL "
    strSQL = strSQL & "INNER JOIN TOGO_SPAC_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " SPAC_YCODE = BF_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_USAGE = BF_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_SIZE = BF_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_STEP = BF_STEP "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_NO  = BF_NO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SPAC_AKI <= BF_AKI "

    After_Update_SPAC_TABL = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : 事務手数料割引データ作成
'       MODULE_ID       : Insert_TOGO_RLDN_TRAN_Jimute
'       CREATE_DATE     : 2014/12/04            K.ISHIZAKA
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_TOGO_RLDN_TRAN_Jimute() As String
    Dim strSQL              As String
    Dim P_GET_OFFICE_FEE_FROM As String
    Dim P_GET_OFFICE_FEE_TO As String
    Const MODULE_ID         As String = "MNR100"
    
    '事務手数料割引の期間内でないときはここで作成しない
    If Not MNR100.IsOffiveFeeGet() Then
        Insert_TOGO_RLDN_TRAN_Jimute = ""
        Exit Function
    End If
    '期間がプライベート宣言されているので取り直す
    P_GET_OFFICE_FEE_FROM = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='OFFICE_FEE_FROM'"))
    P_GET_OFFICE_FEE_TO = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & MODULE_ID & "' AND INTIF_RECFB='OFFICE_FEE_TO'"))

    strSQL = strSQL & "INSERT INTO TOGO_RLDN_TRAN "
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " YARD_BUMOC AS RLDNT_BUMOC, "
    strSQL = strSQL & " YARD_CODE AS RLDNT_YCODE,"
    strSQL = strSQL & " YARD_USAGE AS RLDNT_USAGE, "
    strSQL = strSQL & " PRIC_SIZE AS RLDNT_SIZE, "
    strSQL = strSQL & " PRIC_STEP AS RLDNT_STEP, "
    strSQL = strSQL & " '------' AS RLDNT_NO, "
    strSQL = strSQL & " '1' AS RLDNT_ENABLE, "
    strSQL = strSQL & " 10 AS RLDNT_TYPE, "
    strSQL = strSQL & " " & MNR100.P_GET_OFFICE_FEE & " AS RLDNT_PRICE, "
    strSQL = strSQL & " " & Format(C_JIMU_TESUURYO - CLng(MNR100.P_GET_OFFICE_FEE)) & " AS RLDNT_WARIBIKI, " 'INSERT 2015/01/30 K.ISHIZAKA
    strSQL = strSQL & " 1 AS RLDNT_PERIOD, "
    strSQL = strSQL & " '" & P_GET_OFFICE_FEE_FROM & "' AS RLDNT_FROM, "
    strSQL = strSQL & " '" & P_GET_OFFICE_FEE_TO & "' AS RLDNT_TO, "
    strSQL = strSQL & " 1 AS RLDNT_ORDER, "
    strSQL = strSQL & " '事務手数料:" & MNR100.P_GET_OFFICE_FEE & "円サービス' AS RLDNT_TEXT, "
    strSQL = strSQL & " NULL AS RLDNT_NOTE, "
    strSQL = strSQL & " 99 AS RLDNT_GENKBN, "
    strSQL = strSQL & " YARD_SEV_EXMONTH AS RLDNT_USE_PERIOD, "
    strSQL = strSQL & " '" & strDate & "' AS RLDNT_INSED, "
    strSQL = strSQL & " '" & strTime & "' AS RLDNT_INSEJ, "
    strSQL = strSQL & " '" & PROG_ID & "' AS RLDNT_INSPB, "
    strSQL = strSQL & " '" & strUser & "' AS RLDNT_INSUB, "
    strSQL = strSQL & " NULL AS RLDNT_UPDAD, "
    strSQL = strSQL & " NULL AS RLDNT_UPDAJ, "
    strSQL = strSQL & " NULL AS RLDNT_UPDPB, "
    strSQL = strSQL & " NULL AS RLDNT_UPDUB  "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " TOGO_YARD_MAST  "
    strSQL = strSQL & "INNER JOIN "
    strSQL = strSQL & " TOGO_PRIC_TABL "
    strSQL = strSQL & "ON"
    'strSQL = strSQL & " PRIC_YCODE = SPAC_YCODE "  '2015/01/05 M.HONDA UPD
    strSQL = strSQL & " PRIC_YCODE = YARD_CODE "
    strSQL = strSQL & "AND"
    'strSQL = strSQL & " PRIC_USAGE = SPAC_USAGE "  '2015/01/05 M.HONDA UPD
    strSQL = strSQL & " PRIC_USAGE = YARD_USAGE "
    
    If MNR100.Is2kDiscountServiceGet() Then
        '２千円割引が適用される場合は価格の安いものだけが事務手数料の対象となる
        strSQL = strSQL & "WHERE"
        strSQL = strSQL & " PRIC_PRICE < " & Format(C_2kSEV_LIMIT) & " "
    End If
    
    Insert_TOGO_RLDN_TRAN_Jimute = strSQL
End Function
'==============================================================================*
'
'       MODULE_NAME     : ネット契約割引データ作成
'       MODULE_ID       : Insert_TOGO_RLDN_TRAN_Net
'       CREATE_DATE     : 2018/01/15            M.HONDA
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_TOGO_RLDN_TRAN_Net() As String
    Dim strSQL              As String
    Const MODULE_ID         As String = "MNR100"
    
    strSQL = strSQL & "INSERT INTO TOGO_RLDN_TRAN "
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "    YARD_BUMOC AS RLDNT_BUMOC, "
    strSQL = strSQL & "    YARD_CODE AS RLDNT_YCODE, "
    strSQL = strSQL & "    YARD_USAGE AS RLDNT_USAGE, "
    strSQL = strSQL & "    PRIC_SIZE AS RLDNT_SIZE, "
    strSQL = strSQL & "    PRIC_STEP AS RLDNT_STEP, "
    strSQL = strSQL & "    'ZZZZZZ' AS RLDNT_NO, "
    strSQL = strSQL & "    '1' AS RLDNT_ENABLE, "
    strSQL = strSQL & "    10 AS RLDNT_TYPE, "
    strSQL = strSQL & "    0 AS RLDNT_FPRICE, "    '2019/08/08 M.HONDA INS
    strSQL = strSQL & "    0 AS RLDNT_FWARIBIKI, " '2019/08/08 M.HONDA INS
    strSQL = strSQL & "    0 AS RLDNT_PRICE, "
    If Now >= "2021/03/26 19:00:00" And Now <= "2021/04/01 09:00:00" Then       'INSERT 2021/03/24 N.IMAI
        strSQL = strSQL & " 0 AS RLDNT_WARIBIKI"
    Else
        strSQL = strSQL & "    CASE WHEN PRIC_PRICE < 1100 THEN 0 "
    'chg ↓ 2018/01/28 tajima：strSQL = strSQL & "         WHEN PRIC_SIZE <= '0.9' THEN 1080 "
        strSQL = strSQL & "         WHEN PRIC_SIZE <= '2.9' THEN 1100 "
    'chg ↑ 2018/01/28
    '    strSQL = strSQL & "         ELSE 2200 END RLDNT_WARIBIKI, "                    'DELETE 2020/11/30 K.KINEBUCHI
        strSQL = strSQL & "         ELSE 3300 END RLDNT_WARIBIKI, "                     'INSERT 2020/11/30 K.KINEBUCHI
    End If
    strSQL = strSQL & "    1 AS RLDNT_PERIOD, "
    strSQL = strSQL & "    '00000000' AS RLDNT_FROM, "
    strSQL = strSQL & "    '99999999' AS RLDNT_TO, "
    strSQL = strSQL & "    1 AS RLDNT_ORDER, "
    strSQL = strSQL & "    'ネット契約割引サービス' AS RLDNT_TEXT, "
    strSQL = strSQL & "    NULL AS RLDNT_NOTE, "
    strSQL = strSQL & "    99 AS RLDNT_GENKBN, "
    strSQL = strSQL & "    99 AS RLDNT_USE_PERIOD, "
    strSQL = strSQL & " '" & strDate & "' AS RLDNT_INSED,  "
    strSQL = strSQL & " '" & strTime & "' AS RLDNT_INSEJ,  "
    strSQL = strSQL & " '" & PROG_ID & "' AS RLDNT_INSPB,  "
    strSQL = strSQL & " '" & strUser & "' AS RLDNT_INSUB,  "
    strSQL = strSQL & "    NULL AS RLDNT_UPDAD, "
    strSQL = strSQL & "    NULL AS RLDNT_UPDAJ, "
    strSQL = strSQL & "    NULL AS RLDNT_UPDPB, "
    strSQL = strSQL & "    NULL As RLDNT_UPDUB "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    TOGO_YARD_MAST "
    strSQL = strSQL & "    INNER Join "
    strSQL = strSQL & "    TOGO_PRIC_TABL ON "
    strSQL = strSQL & "    PRIC_YCODE = YARD_CODE AND "
    strSQL = strSQL & "    PRIC_USAGE = YARD_USAGE "
    
    Insert_TOGO_RLDN_TRAN_Net = strSQL

End Function

'==============================================================================*
'
'       MODULE_NAME     : 保証委託料割引データ作成
'       MODULE_ID       : Insert_TOGO_RLDN_TRAN_itakuryo
'       CREATE_DATE     : 2014/12/04            K.ISHIZAKA
'       PARAM           : strKaseDb             KASEＤＢ名(I)
'       RETURN          : ＳＱＬ文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Insert_TOGO_RLDN_TRAN_itakuryo(ByVal strKaseDb As String) As String
    Dim strSQL              As String

    strSQL = strSQL & "DECLARE"
    strSQL = strSQL & " @NET_HOSHO_CD AS VARCHAR(7);"
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " @NET_HOSHO_CD = CASE WHEN PGPAT_PARAN <= CONVERT(varchar(8), GETDATE(), 112) + REPLACE(CONVERT(varchar(8), GETDATE(), 114),':','')"
    strSQL = strSQL & " THEN"
    strSQL = strSQL & "  CASE WHEN EXISTS"
    strSQL = strSQL & "  ("
    strSQL = strSQL & "  SELECT"
    strSQL = strSQL & "   * "
    strSQL = strSQL & "  FROM"
    strSQL = strSQL & "   " & strKaseDb & ".dbo.PGPA_TABL "
    strSQL = strSQL & "  WHERE"
    strSQL = strSQL & "   PGPAT_PGP1B LIKE 'MNR210%' "
    strSQL = strSQL & "  AND"
    strSQL = strSQL & "   PGPAT_PGP2B = 'NET_HOSHO_FLG' "
    strSQL = strSQL & "  AND"
    strSQL = strSQL & "   PGPAT_PARAN = 0"
    strSQL = strSQL & "  )"
    strSQL = strSQL & "  THEN"
    strSQL = strSQL & "   ("
    strSQL = strSQL & "   SELECT"
    strSQL = strSQL & "    PGPAT_PARAN "
    strSQL = strSQL & "   FROM"
    strSQL = strSQL & "    " & strKaseDb & ".dbo.PGPA_TABL "
    strSQL = strSQL & "   WHERE"
    strSQL = strSQL & "    PGPAT_PGP1B LIKE 'MNR210%' "
    strSQL = strSQL & "   AND"
    strSQL = strSQL & "    PGPAT_PGP2B = 'NET_HOSHO_CD' "
    strSQL = strSQL & "   )"
    strSQL = strSQL & "  ELSE"
    strSQL = strSQL & "   NULL"
    strSQL = strSQL & "  END"
    strSQL = strSQL & " ELSE '970030'"
    strSQL = strSQL & " END "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKaseDb & ".dbo.PGPA_TABL "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " PGPAT_PGP1B LIKE 'MNR210%' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PGPAT_PGP2B = 'NET_HOSHO_SDATE' "
    strSQL = strSQL & ";"
    strSQL = strSQL & "INSERT INTO TOGO_RLDN_TRAN "
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YARD_BUMOC AS RLDNT_BUMOC,"
    strSQL = strSQL & " YARD_CODE AS RLDNT_YCODE,"
    strSQL = strSQL & " YARD_USAGE AS RLDNT_USAGE,"
    strSQL = strSQL & " PRIC_SIZE AS RLDNT_SIZE,"
    strSQL = strSQL & " PRIC_STEP AS RLDNT_STEP,"
    strSQL = strSQL & " DCNT_NO AS RLDNT_NO,"
    strSQL = strSQL & " '1' AS RLDNT_ENABLE,"
    strSQL = strSQL & " CASE TEKIYO_HOU"
    strSQL = strSQL & " WHEN '1' THEN '21'"
    strSQL = strSQL & " WHEN '2' THEN '22'"
    strSQL = strSQL & " WHEN '3' THEN '23'"
    strSQL = strSQL & " END AS RLDNT_TYPE,"
'    strSQL = strSQL & " CAST(ROUND(PRIC_PRICE / 100 * KYAKU_RATE,0,1) AS NUMERIC(9,0)) AS RLDNT_PRICE," 'DELETE 2015/01/26 K.ISHIZAKA
'    strSQL = strSQL & " CAST(ROUND((PRIC_PRICE + ISNULL(RLDNT_PRICE,0) + CASE WHEN PRIC_EZAPPI_CODE = '68' THEN PRIC_EZAPPI ELSE 0 END) / 100 * KYAKU_RATE,0,1) AS NUMERIC(9,0)) AS RLDNT_PRICE," 'DELETE 2015/01/30 K.ISHIZAKA 'INSERT 2015/01/26 K.ISHIZAKA
'    strSQL = strSQL & " CAST(ISNULL(RLDNT_PRICE,PRIC_PRICE) - ROUND((ISNULL(RLDNT_PRICE,PRIC_PRICE) + CASE WHEN PRIC_EZAPPI_CODE = '68' THEN PRIC_EZAPPI ELSE 0 END) / 100 * KYAKU_RATE,0,1) AS NUMERIC(9,0)) AS RLDNT_PRICE," 'DELETE 2015/01/31 K.ISHIZAKA 'INSERT 2015/01/30 K.ISHIZAKA
    strSQL = strSQL & " CAST(ISNULL(RLDNT_PRICE,PRIC_PRICE) + CASE WHEN PRIC_EZAPPI_CODE = '68' THEN PRIC_EZAPPI ELSE 0 END "
    strSQL = strSQL & "    - ROUND((ISNULL(RLDNT_PRICE,PRIC_PRICE) + CASE WHEN PRIC_EZAPPI_CODE = '68' THEN PRIC_EZAPPI ELSE 0 END) / 100 * KYAKU_RATE,0,1) AS NUMERIC(9,0)) AS RLDNT_PRICE," 'INSERT 2015/01/31 K.ISHIZAKA
    strSQL = strSQL & " CAST(ROUND((ISNULL(RLDNT_PRICE,PRIC_PRICE) + CASE WHEN PRIC_EZAPPI_CODE = '68' THEN PRIC_EZAPPI ELSE 0 END) / 100 * KYAKU_RATE,0,1) AS NUMERIC(9,0)) AS RLDNT_WARIBIKI," 'INSERT 2015/01/30 K.ISHIZAKA
    strSQL = strSQL & " 1 AS RLDNT_PERIOD,"
    strSQL = strSQL & " SDATE AS RLDNT_FROM,"
    strSQL = strSQL & " EDATE AS RLDNT_TO,"
    strSQL = strSQL & " 1 AS RLDNT_ORDER,"
    strSQL = strSQL & " MONGON AS RLDNT_TEXT,"
    strSQL = strSQL & " NULL AS RLDNT_NOTE,"
    strSQL = strSQL & " 1 AS RLDNT_GENKBN,"
    strSQL = strSQL & " NULL AS RLDNT_USE_PERIOD,"
    strSQL = strSQL & " '" & strDate & "' AS RLDNT_INSED,"
    strSQL = strSQL & " '" & strTime & "' AS RLDNT_INSEJ,"
    strSQL = strSQL & " '" & PROG_ID & "' AS RLDNT_INSPB,"
    strSQL = strSQL & " '" & strUser & "' AS RLDNT_INSUB,"
    strSQL = strSQL & " NULL AS RLDNT_UPDAD,"
    strSQL = strSQL & " NULL AS RLDNT_UPDAJ,"
    strSQL = strSQL & " NULL AS RLDNT_UPDPB,"
    strSQL = strSQL & " NULL AS RLDNT_UPDUB "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " MAX(CASE COLID WHEN 1 THEN PGPAT_PARAN ELSE NULL END) AS DCNT_NO,"
    strSQL = strSQL & " MAX(CASE COLID WHEN 2 THEN PGPAT_PARAN ELSE NULL END) AS EDATE,"
    strSQL = strSQL & " MAX(CASE COLID WHEN 3 THEN PGPAT_PARAN ELSE NULL END) AS JITSU_RATE,"
    strSQL = strSQL & " MAX(CASE COLID WHEN 4 THEN PGPAT_PARAN ELSE NULL END) AS KYAKU_RATE,"
    strSQL = strSQL & " MAX(CASE COLID WHEN 5 THEN PGPAT_PARAN ELSE NULL END) AS MONGON,"
    strSQL = strSQL & " MAX(CASE COLID WHEN 6 THEN PGPAT_PARAN ELSE NULL END) AS SDATE,"
    strSQL = strSQL & " MAX(CASE COLID WHEN 7 THEN PGPAT_PARAN ELSE NULL END) AS TEKIYO_HCD,"
    strSQL = strSQL & " MAX(CASE COLID WHEN 8 THEN PGPAT_PARAN ELSE NULL END) AS TEKIYO_HOU "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " ROW_NUMBER() OVER (PARTITION BY PGPAT_PGP1B ORDER BY PGPAT_PGP2B) AS COLID,"
    strSQL = strSQL & " * "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKaseDb & ".dbo.PGPA_TABL "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " PGPAT_PGP1B like 'MSZZ0068%' "
    strSQL = strSQL & ") wk "
    strSQL = strSQL & "GROUP BY"
    strSQL = strSQL & " PGPAT_PGP1B "
    strSQL = strSQL & ") pgpa "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_YARD_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " ISNULL(@NET_HOSHO_CD,YARD_NET_HOSYO_CD) = TEKIYO_HCD "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " TOGO_PRIC_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PRIC_YCODE = YARD_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PRIC_USAGE = YARD_USAGE "
    strSQL = strSQL & "LEFT JOIN"                                               'INSERT START 2015/01/26 K.ISHIZAKA
    strSQL = strSQL & " TOGO_RLDN_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " RLDNT_YCODE = YARD_CODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_USAGE = YARD_USAGE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_SIZE = PRIC_SIZE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_STEP = PRIC_STEP "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_TYPE = 0 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_PERIOD = 0 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_GENKBN > 0 " '電話以外（ネット or 区別なし）
    strSQL = strSQL & "AND"
    strSQL = strSQL & " RLDNT_ENABLE = '1' "                                    'INSERT END   2015/01/26 K.ISHIZAKA
    strSQL = strSQL & ";"

    Insert_TOGO_RLDN_TRAN_itakuryo = strSQL
End Function

'==============================================================================*
'
'       MODULE_NAME     : Webも考慮した統合ＤＢの空き件数を取得する
'       MODULE_ID       : MSTGP01_M20
'       CREATE_DATE     : 2014/12/12            K.ISHIZAKA
'       PARAM           : strYCODE              ヤードコード(I)
'                       : strNO                 コンテナ番号(I)
'                       : [strBumoc]            接続部門コード(I)
'       RETURN          : 空き件数(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function MSTGP01_M20(ByVal strYcode As String, ByVal strNO As String, _
    Optional ByVal strBUMOC As String = "", Optional ByVal intHEYA As Integer = 0) As Long
    Dim lngCnt              As Long
    Dim objCon              As Object
    Dim objRst              As Object
    Dim strKomsDb           As String
    Dim strWhere            As String
    Dim strUsage            As String
    Dim strSQL              As String
    On Error GoTo ErrorHandler
    
    'KOMSのDB名取得
    If strBUMOC = "" Then
        strBUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    End If
    strKomsDb = DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME_" & strBUMOC & "'")
    
    '統合DB
    Set objCon = ADODB_Connection(C_TOGO_CONNENCT_BUMOC)
    On Error GoTo ErrorHandler1
    
    '用途、サイズ、段を取得
    strSQL = "SELECT"
    strSQL = strSQL & " CNTA_USAGE,"
    strSQL = strSQL & " CNTA_SIZE,"
    strSQL = strSQL & " CNTA_STEP "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " " & strKomsDb & ".dbo.CNTA_MAST "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " CNTA_CODE = " & strYcode & " "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_NO = " & strNO & " "
    
    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler2
    With objRst
        If Not .EOF Then
            strWhere = "WHERE SPAC_YCODE = " & strYcode & " "
            strUsage = .Fields("CNTA_USAGE")
           
           If intHEYA = -1 Then
                strWhere = strWhere & "AND SPAC_NO = " & strNO & " "
            Else
                strWhere = strWhere & "AND SPAC_USAGE = " & strUsage & " "
                strWhere = strWhere & "AND SPAC_SIZE = " & .Fields("CNTA_SIZE") & " "
                strWhere = strWhere & "AND SPAC_STEP = " & .Fields("CNTA_STEP") & " "
            End If
        Else
            strWhere = ""
        End If
        .Close
    End With
    On Error GoTo ErrorHandler1
    
    If strWhere <> "" Then
        strSQL = "SELECT SPAC_AKI FROM TOGO_SPAC_TABL " & strWhere
        lngCnt = ADODB_ExecGetLong(strSQL, objCon)
    Else
        lngCnt = -1
    End If
    
'        If CBool(Nz(DLookup("INTIF_RECFB", "INTI_FILE", "INTIF_PROGB = 'FTG011'"), "False")) Then
'            lngCnt = MySql_ExecGetLong(strSQL)
'        Else
'            lngCnt = ADODB_ExecGetLong(strSQL, objCon)
'        End If
    
    objCon.Close
    On Error GoTo ErrorHandler
    MSTGP01_M20 = lngCnt
Exit Function

ErrorHandler2:
    objRst.Close
ErrorHandler1:
    objCon.Close
ErrorHandler:                   '↓自分の関数名
    Call MSZZ024_M00("MSTGP01_M20", False)
End Function

'****************************  ended of program ********************************

