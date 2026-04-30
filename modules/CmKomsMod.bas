Attribute VB_Name = "CmKomsMod"
'****************************  strat or program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬総合システム
'        SUB_SYSTEM_NAME :
'
'        PROGRAM_NAME    :
'        PROGRAM_ID      : CmKomsMethods
'        PROGRAM_KBN     :
'
'        CREATE          : 2005/10/14
'        CERATER         : T.SUZUKI
'        Ver             : 0.0
'
'        UPDATE          : 2005/12/03
'        UPDATER         : H.TAJIMA & S.SHIBAZAKI
'        Ver             : 0.0
'                        : 紹介待ち対応
'
'        UPDATE          : 2006/01/20
'        UPDATER         : H.TAJIMA
'        Ver             : 0.1
'                        : 移動予約対応
'
'        UPDATE          : 2006/02/03
'        UPDATER         : S.SHIBAZAKI
'        Ver             : 0.2
'                        : 新規取置きデータ作成時に初回印刷日を当日に設定する
'
'        UPDATE          : 2006/02/06
'        UPDATER         : T.SUZUKI
'        Ver             : 0.3
'                        : 請求ファイル新規作成関数追加
'                        : 予約受付トラン更新関数追加
'                        : 予約ご紹介トラン更新関数追加
'                        : コンテナ契約ファイル新規作成関数追加
'        UPDATE          : 2006/02/16
'        UPDATER         : H.TAJIMA
'                        : 移動用備考生成関数追加
'        UPDATE          : 2006/02/18
'        UPDATER         : T.SUZUKI
'                        : 既存顧客検索画面制御用関数追加
'        UPDATE          : 2006/03/02
'        UPDATER         : TAJIMA
'                        : コンテナ契約ファイル新規作成関数修正
'                        : キャンペーン適用情報の表示機能追加
'        Ver             : 0.4
'        UPDATE          : 2006/04/24
'        UPDATER         : TAJIMA
'                        : 消費税対応 KASE_DBから情報取得
'        Ver             : 0.5
'        UPDATE          : 2006/05/12
'        UPDATER         : TAJIMA
'                        : 予約入力時の移動予約制限解除
'        Ver             : 0.6
'        UPDATE          : 2006/05/24
'        UPDATER         : TAJIMA
'                        : 鍵＆セキュリティカード対応
'        Ver             : 0.7
'        UPDATE          : 2006/09/27
'        UPDATER         : TAJIMA
'                        : 修繕対応
'
'        Ver             : 0.8
'        UPDATE          : 2007/04/15
'        UPDATER         : TAJIMA
'                        : 近隣紹介対応
'
'        Ver             : 0.9
'        UPDATE          : 2008/03/25
'        UPDATER         : TAJIMA
'                        : 移動区分の追加対応
'
'        Ver             : 1.0
'        UPDATE          : 2009/04/01
'        UPDATER         : hirano
'                        : 保証会社コードの追加対応
'
'        Ver             : 1.1
'        UPDATE          : 2010/04/10
'        UPDATER         : K.ISHIZAKA
'                        : 請求書発行区分の追加対応
'
'        Ver             : 1.2
'        UPDATE          : 2010/12/21
'        UPDATER         : M.HONDA
'                        : 名義変更の際には、清算依頼日に翌月5日をセット
'
'        Ver             : 1.3
'        UPDATE          : 2011/09/27
'        UPDATER         : M.RYU
'                        : 移動区分を追加⇒07紹介者金額変更
'
'        Ver             : 1.4
'        UPDATE          : 2012/07/07
'        UPDATER         : M.HONDA
'                        : 移動処理に発生区分を電話をデフォルトで設定
'
'        Ver             : 1.5
'        UPDATE          : 2013/07/08
'        UPDATER         : M.HONDA
'                        : 清算依頼日の設定条件を変更
'
'        Ver             : 1.6
'        UPDATE          : 2013/07/25
'        UPDATER         : M.HONDA
'                        : 契約トラン保存に貸出用途を追加
'
'        Ver             : 1.7
'        UPDATE          : 2014/06/20
'        UPDATER         : MIYAMOTO
'                        : 契約トラン構造体に鍵変更理由コードを追加
'                          コンテナ契約ファイル新規作成処理、移動元コンテナ契約トラン更新処理に鍵変更理由コードを設定する処理を追加
'
'        Ver             : 1.8
'        UPDATE          : 2015/07/16
'        UPDATER         : K.ISHIZAKA
'                        : コンテナ契約ファイルの新規作成時、今後は既に存在する場合があるので、そのときは更新する
'                        : 請求ファイルデータの新規作成時、今後は既に存在する場合があるので、そのときは何もしない
'
'        Ver             : 1.9
'        UPDATE          : 2015/07/26
'        UPDATER         : K.ISHIZAKA
'                        : Ver1.8への修正
'
'        Ver             : 1.10
'        UPDATE          : 2017/07/31 V1.10
'        UPDATER         : YSUZUKI
'                        : 既存契約（顧客コードで他の契約）があるかの確認( メソッド追加 )
'
'        Ver             : 2.0
'        UPDATE          : 2017/07/29
'        UPDATER         : M.HONDA
'                        : 受付キャンセル時にコンテナ（バイク）をコンテナに戻す
'
'        Ver             : 2.1
'        UPDATE          : 2018/03/19
'        UPDATER         : EGL
'                        : 請求削除ファイルへの追記メソッド追加
'
'        Ver             : 2.2
'        UPDATE          : 2018/05/23
'        UPDATER         : N.IMAI
'                        : 契約済みチェックでチェック除外の予約番号を使用するように変更
'
'        Ver             : 2.3
'        UPDATE          : 2018/09/25
'        UPDATER         : EGL
'                        : 分社化対応
'
'        Ver             : 2.4
'        UPDATE          : 2019/03/06
'        UPDATER         : Y.WADA
'                        : IsMoveReserveEntry:契約変更で更新元契約を複数回可能とする
'
'        Ver             : 2.5
'        UPDATE          : 2019/06/15
'        UPDATER         : N.IMAI
'                        : 受付キャンセル時にバイク（コンテナ）をバイクに戻す
'
'        Ver             : 2.6
'        UPDATE          : 2019/08/02
'        UPDATER         : K.ISHIZAKA
'                        : 消費税計算の変更
'
'        Ver             : 2.7
'        UPDATE          : 2020/03/25
'        UPDATER         : EGL
'                        : 請求ファイル削除の不具合対応
'
'        Ver             : 2.8
'        UPDATE          : 2020/04/01
'        UPDATER         : N.IMAI
'                        : セキュリティーカードの取得をADO接続に変更
'
'        Ver             : 2.9
'        UPDATE          : 2020/09/28
'        UPDATER         : EGL
'                        : プランコード対応
'
'        Ver             : 3.0
'        UPDATE          : 2021/03/31
'        UPDATER         : N.IMAI
'                        : CheckRegiTargetをODBCに変更
'
'        Ver             : 3.1
'        UPDATE          : 2022/03/02
'        UPDATER         : N.IMAI
'                        : 請求ファイルのデータを削除が全件正しく削除されない問題に対応
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const PROG_ID = "CmKomsMod"

' CheckLending 用の定数
Public Const pcintCONTENA_UNITCHECK   As Integer = 1    ' コンテナチェック  0000:0000:0001
Public Const pcintSHOUMATI_UNITCHECK  As Integer = 2    ' 紹介待ちチェック  0000:0000:0010
Public Const pcintYARDKADO_UNITCHECK  As Integer = 4    ' ヤード稼働チェック0000:0000:0100
Public Const pcintYOYAKU_UNITCHECK    As Integer = 8    ' 予約中チェック    0000:0000:1000
Public Const pcintKEIYAKU_UNITCHECK   As Integer = 16   ' 契約中チェック    0000:0001:0000
Public Const pcintKAIYAKU_UNITCHECK   As Integer = 32   ' 解約予定チェック  0000:0010:0000
Public Const pcint自社予約理由_FLG    As Integer = 256  ' 自社理由除外ﾌﾗｸﾞ  0001:0000:0000

Public Const pcintCONUSEWAIT_CHECK    As Integer = 3    ' ｺﾝﾃﾅ.紹介待ちﾁｪｯｸ 0000:0000:0011
Public Const pcintYOYAKU_CHECK        As Integer = 31   ' 予約系基本ﾁｪｯｸ    0000:0001:1111
Public Const pcintUKETUYOYAKU_CHECK   As Integer = 23   ' 契約受付予約ﾁｪｯｸ  0000:0001:0111
Public Const pcintJISYAYOYAKU_CHECK   As Integer = 47   ' 自社理由予約ﾁｪｯｸ  0000:0010:1011 '2006/09/27 0100を落とした
Public Const pcintKEIYAKU_CHECK       As Integer = 15   ' 契約系チェック    0000:0000:1111
Public Const pcintKAIMODO_CHECK       As Integer = 280  ' 解約戻しチェック  0001:0001:1000
Public Const pcintALL_CHECK           As Integer = 223  ' オールチェック    0000:1101:1111
Public Const pcintTEST_ALL_CHECK      As Integer = 255  ' テストオール用    0000:1111:1111

Private Const pcstrYUKBN_01         As String = "01" ' 自社理由
Private Const pcstrYUKBN_02         As String = "02" ' 取置
Private Const pcstrYUKBN_10         As String = "10" ' 受付
Private Const pcstrYUKBN_53         As String = "53" ' 受付キャンセル

Private Const pcintAGRE_01          As Integer = 1  ' 契約状態(契約)
Private Const pcintAGRE_02          As Integer = 2  ' 契約状態(解約)
Private Const pcintAGRE_09          As Integer = 9  ' 契約状態(完了)
Private Const pcintCNTA_USE_NO      As Integer = 0  ' コンテナ利用否
Private Const pcintCNTA_USE_OK      As Integer = 1  ' コンテナ利用可
Private Const pcintCNTA_USE_DEL     As Integer = 9  ' コンテナ利用撤去

' チェック文言
Private Const pcstrYOYAKU_ErrMsg    As String = "このコンテナは他で予約中です。"
Private Const pcstrKEIYAKU_ErrMsg   As String = "このコンテナは他で契約中です。"
Private Const pcstrKADO_ErrMsg      As String = "このヤードは既に稼働していません。" _
                                       & vbCr & "ヤード解約予定日を過ぎています。"  '2006/09/27 文言追加
Private Const pcstrSHOUMATI_ErrMsg  As String = "このコンテナは紹介待ちです。"
Private Const pcstrCUSING_ErrMsg    As String = "このコンテナは現在利用できません。"
Private Const pcstrTEKYO_ErrMsg     As String = "このコンテナは撤去されました。"
Private Const pcstrKAIYOTEI_ErrMsg  As String = "このコンテナは解約予定ではありません。"

' チェック結果
Public Const pcintSYSTEM_ERROR     As Integer = -1  ' システムエラー
Public Const pcintCHECK_OK         As Integer = 0   ' 利用可能
Public Const pcintSHOUKAI_WAIT     As Integer = 1   ' 紹介待ち
Public Const pcintCNTA_NONUSE      As Integer = 2   ' 利用不可のコンテナ
Public Const pcintCNTA_DELETED     As Integer = 3   ' 撤去されたコンテナ
Public Const pcintYOYAKU_Err       As Integer = 4   ' 他で予約(自社理由、取置、受付)済み
Public Const pcintNOTKAIYAKU       As Integer = 5   ' 非解約契約（自社理由予約不可）
Public Const pcintKEIYAKU_Err      As Integer = 6   ' 他で契約済み
Public Const pcintKADO_Err         As Integer = 7   ' 稼働ヤードではない

' 2005/12/7 S.SHIBAZAKI 追加↓
Private Const pcstrIntroSts_KEEP    As String = "1" '紹介区分「取置きした」
Private Const pcstrIntroSts_CANCEL  As String = "6" '紹介区分「キャンセルした」

'処理モード
Private Const pcintProcMode_単一コンテナ指定    As Integer = 1  '単一コンテナ指定
Private Const pcintProcMode_全取消し          As Integer = 2  '全コンテナ
Private Const pcintProcMode_紹介区分変更_紹介待ち作成    As Integer = 3  '指定コンテナ除外
Private Const pcintProcMode_単純追加    As Integer = 4  '紹介データ新規作成
Private Const pcintProcMode_単純更新    As Integer = 5  '紹介データ紹介区分更新
' 2005/12/7 S.SHIBAZAKI 追加↑

' 2006/01/28 T.SUZUKI 追加↓
' 【移動区分】
Private Const pcstrMOVEKBN_00       As String = "00"          ' 無し
Private Const pcstrMOVEKBN_01       As String = "01"          ' 自社理由移動
Private Const pcstrMOVEKBN_02       As String = "02"          ' 顧客理由移動
Private Const pcstrMOVEKBN_03       As String = "03"          ' 未使用移動
Private Const pcstrMOVEKBN_04       As String = "04"          ' 名義変更
Private Const pcstrMOVEKBN_05       As String = "05"          ' 金額変更
Private Const pcstrMOVEKBN_06       As String = "06"          ' 保障会社変更
Private Const pcstrMOVEKBN_07       As String = "07"          ' 紹介者金額変更  'INSERT 2011/09/27 M.RYU

' 【契約状態】
Private Const pcstrAGRE_01          As String = "01"          ' 契約
Private Const pcstrAGRE_02          As String = "02"          ' 解約予定
Private Const pcstrAGRE_03          As String = "03"          ' 解約延長
Private Const pcstrAGRE_04          As String = "04"          ' メンテ中
Private Const pcstrAGRE_09          As String = "09"          ' 完了

' 【解約区分】
Private Const pcstrKAICD_01         As String = "01"          ' 通常
Private Const pcstrKAICD_03         As String = "03"          ' 未使用
Private Const pcstrKAICD_04         As String = "04"          ' 移動
Private Const pcstrKAICD_05         As String = "05"          ' 契約変更
Private Const pcstrKAICD_06         As String = "06"          ' 保証会社理由

' 【解約理由区分コード】
Private Const pcstrKAIRIUCD_07      As String = "07"          ' 加瀬の別のボックスに移動
Private Const pcstrKAIRIUCD_99      As String = "99"          ' その他
' 2006/01/28 T.SUZUKI 追加↑

' 2006/02/06 T.SUZUKI 追加↓
Private Const pcsglMoneySumMaxValue As Double = 9999999999#  ' 請求ファイル作成時の金額合計値の最大値
' コンテナ契約ファイル構造体
Public Type Type_CARG_FILE
    YCODE      As Long
    No         As Long
    UCODE      As Long
    AGRE       As Integer
    FSDATE     As Variant
    STDATE     As Variant
    EDDATE     As Variant
    CYDATE     As Variant
    KYDATE     As Variant
    USAGE      As Integer
    DOCU1      As Integer
    DOCU2      As Integer
    RENTKG     As Double
    SYOZEI     As Double
    FRSTKG     As Double
    FSYOZEI    As Double
    FRST_BILL  As Double  'Add 2005/03/02 tajima 初回請求金額
    PREMTH_SUM As Double  'Add 2005/03/02 tajima 前払月数
    SECUKG     As Double
    BIKO       As String
    UPDATE     As Variant
    MDATE      As Variant
    ACPTNO     As String
    CAMPC      As String  '会社コード 2018/09/25 EGL INS
    HOSYICD    As String
    HOSYO_CD   As String  '保証会社コード 2009/04/01 Add
    KAGIICD    As String  '鍵区分
    DAHIB      As String
    KAIJB      As String  '鍵解除番号
    HOKAI      As String
    HOSYD      As Variant
    HOSYB      As String
    ADMEI      As String
    KEITI      As String
    KAGIA      As Double
    UKNO       As String
    UKTANTO    As String
    KAIUKEDATE As Variant
    KAITANTO   As String
    KAICD      As Integer
    KAIRIYUCD  As Variant
    CMP_EXDATE As Variant  'Add 2005/03/02 tajima キャンペーン適用満了日
    KEY_LENTNUM As Variant 'Add 2005/05/24 tajima 鍵貸出本数
    SEIKI      As String   'Add 2010/04/10 K.ISHIZAKA 請求書発行区分
    KEY_CHGRIYUCD As Integer    '鍵変更理由コード       'INSERT 2014/06/20 MIYAMOTO
    PLANCD      As String   'add 2020/09/28 tajima プランコード
End Type

Public CARG_FILE          As Type_CARG_FILE
' 2006/02/06 T.SUZUKI 追加↑

' 2006/02/06 T.SUZUKI 追加↓
Private pstrParentFormName As String
Private pstrParentUserCd   As String
' 2006/02/06 T.SUZUKI 追加↑

' 2006/04/25 tajima 追加
Private pvalBeforTaxRate As Variant '変更前消費税率
Private pvalAfterTaxRate As Variant '変更後消費税率
Private pstrTaxJudgDate As String   '消費税率適用開始年月
Private pstrRoundType As String     '消費税額端数区分

' 2006/06/07 tajima 追加
Public Const P_カード部屋割当無し値 As Integer = -1
Public Const P_カード用途_予備値    As String = "00"
Public Const P_カード用途_基本値    As String = "01"
Public Const P_カード用途_追加値    As String = "02"
Public Const P_カード用途_業者貸値  As String = "03"
'==============================================================================*
'
'        MODULE_NAME      :物件チェック
'        MODULE_ID        :CheckLending
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(Integer) = チェック区分
'                         :第3引数(Long) = ヤードコード
'                         :第4引数(Long) = コンテナ番号
'                         :第5引数(ByRef String) = エラーメッセージ格納
'                         :第6引数(String)[省略可] = チェック除外の予約番号（省略時は全予約対象）
'                         :第7引数(String)[省略可] = チェック除外の契約番号（省略時は全契約対象）
'        戻り値
'                           0            = 可能
'                          -1            = システムエラー
'                           上記以外     = エラーとなった内容
'        CREATE_DATE      :2005/10/19
'        UPDATE_DATE      :2005/12/12 引数変更、DataSetを外部でオープン
'        UPDATE_DATE      :2006/01/21 移動予約対応のため、第７引数追加
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CheckLending(dbSQLServer As DAO.Database, _
                              ByVal intCheck As Integer, _
                              ByVal lngYardCode As Long, _
                              ByVal lngCntaCode As Long, _
                              ByRef strErrMsg As String, _
                              Optional strUkno As String = "", _
                              Optional strACPTNO As String = "" _
                              ) As Integer

    Dim objRs    As Recordset
    Dim strSQL   As String
    Dim intCount As Integer

    Dim intKbn   As Integer
    Dim strMsg   As String   ' エラーメッセージ格納ワーク変数

    On Error GoTo CheckLending_Err

    CheckLending = -1  ' 初初期化
    strErrMsg = ""
    intKbn = pcintCHECK_OK  ' 初期化

    ' コンテナ利用可否 or 紹介待ちのチェック
    If (intCheck And pcintCONTENA_UNITCHECK) > 0 Or (intCheck And pcintSHOUMATI_UNITCHECK) > 0 Then
        ' コンテナマスタの利用可否区分と紹介待ち情報を取得する
        strSQL = "SELECT CNTA_USE, CNTA_UKNO "
        strSQL = strSQL & " FROM CNTA_MAST "
        strSQL = strSQL & "WHERE CNTA_CODE  = " & lngYardCode & Chr(13)  ' ヤードコード
        strSQL = strSQL & "  AND CNTA_NO    = " & lngCntaCode & Chr(13)  ' コンテナ番号
        Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
        Dim intUse   As Integer  ' コンテナ利用可否
        Dim strWaitUkNo  As String   ' 紹介待ちにした予約番号
        intUse = Nz(objRs.Fields("CNTA_USE"), 0)
        strWaitUkNo = Nz(objRs.Fields("CNTA_UKNO"), "")
        objRs.Close
        Set objRs = Nothing

        ' コンテナ利用可否のチェック
        If (intCheck And pcintCONTENA_UNITCHECK) > 0 Then
          Select Case intUse
            ' 利用不可コンテナ
            Case pcintCNTA_USE_NO
                  intKbn = pcintCNTA_NONUSE
                  strErrMsg = pcstrCUSING_ErrMsg
                  If (intCheck And pcintSHOUMATI_UNITCHECK) = 0 Then
                    GoTo CheckLending_Exit  ' 紹介待ちのチェックをしないのならば抜けます
                  End If
            ' 撤去コンテナである
            Case pcintCNTA_USE_DEL
                  intKbn = pcintCNTA_DELETED
                  strErrMsg = pcstrTEKYO_ErrMsg
                  GoTo CheckLending_Exit    ' 撤去コンテナなので抜ける
            ' どれにも当てはまらないのは利用出来る
            Case Else
                  intKbn = pcintCHECK_OK
                  strErrMsg = ""
          End Select
        End If
        
        ' 紹介待ちチェックのチェック
        If (intCheck And pcintSHOUMATI_UNITCHECK) > 0 Then
          ' 紹介待ちにした予約番号に何か値が入っていれば紹介待ち状態です。
          If strWaitUkNo <> "" Then
            intKbn = pcintSHOUKAI_WAIT
            strErrMsg = pcstrSHOUMATI_ErrMsg
            GoTo CheckLending_Exit
          '2006/04/19 tajima Add Start
          ' 紹介待ちにした予約番号に何も入ってなくてもコンテナ利用可でなければ利用不可
          ElseIf intUse <> pcintCNTA_USE_OK Then
            intKbn = pcintCNTA_NONUSE
            strErrMsg = pcstrCUSING_ErrMsg
            GoTo CheckLending_Exit    ' 利用可能ではないので抜ける
          End If
          '2006/04/19 tajima Add End
        End If
        
    End If
    
    ' 予約済みチェック
    If (intCheck And pcintYOYAKU_UNITCHECK) > 0 Then
        ' 以下のデータが存在するかチェックする
        ' 予約受付トランに同一のヤードコード、コンテナ番号かつ、予約状態：受付
        ' 予約紹介トランに同一のヤードコード、コンテナ番号かつ、紹介区分：取置
        ' 2007/04/15 紹介トランの評価のみで良い気がする・・・
        strSQL = "SELECT COUNT(*) AS Cnt "
        strSQL = strSQL & " FROM YOUK_TRAN LEFT OUTER JOIN INTR_TRAN ON YOUKT_UKNO = INTRT_UKNO "
        strSQL = strSQL & " WHERE " & Chr(13)                     ' 2007/04/15 chg tajima
        strSQL = strSQL & "   ( YOUKT_YUKBN IN(" & pcstrYUKBN_10  ' 予約状態：受付  ' 2007/04/15 chg tajima

        If (intCheck And pcint自社予約理由_FLG) = 0 Then
          strSQL = strSQL & "," & pcstrYUKBN_01 ' 自社理由除外ﾌﾗｸﾞが立っていなければﾁｪｯｸする
        End If

        strSQL = strSQL & " ) AND YOUKT_YCODE = " & lngYardCode & "AND YOUKT_NO = " & lngCntaCode & Chr(13)  ' 予約・コンテナ番号

        If strUkno <> "" Then
          strSQL = strSQL & "   AND YOUKT_UKNO <> '" & strUkno & "'" & Chr(13)    ' チェック除外の予約番号の指定
        End If

        strSQL = strSQL & " ) OR ( YOUKT_YUKBN = " & pcstrYUKBN_02 & Chr(13)      ' 予約状態：取置
        strSQL = strSQL & "     AND INTRT_YCODE = " & lngYardCode & Chr(13)       ' 紹介・ヤードコード 2007/04/15 add tajima
        strSQL = strSQL & "     AND INTRT_NO   = " & lngCntaCode & Chr(13)        ' 紹介・コンテナ番号
        strSQL = strSQL & "     AND INTRT_INTROKBN = " & pcstrIntroSts_KEEP       ' 紹介区分：取置

        If strUkno <> "" Then
          strSQL = strSQL & "   AND INTRT_UKNO <> '" & strUkno & "'" & Chr(13)    ' チェック除外の予約番号の指定
        End If
        strSQL = strSQL & ") "                                       ' 2007/04/15 chg tajima
        
        Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

        intCount = 0
        intCount = Nz(objRs.Fields("Cnt"), 0)
        objRs.Close
        Set objRs = Nothing

        If intCount > 0 Then
            ' 予約済みチェックエラー
            intKbn = pcintYOYAKU_Err
            strErrMsg = pcstrYOYAKU_ErrMsg
            GoTo CheckLending_Exit
        Else
            intKbn = pcintCHECK_OK
        End If
    End If


    ' 契約済みチェック
    If (intCheck And pcintKEIYAKU_UNITCHECK) > 0 Then
        ' (コンテナ契約ファイルに同一のヤードコード、コンテナ番号かつ、契約状態が完了以外のデータが
        '  存在するか否かをチェック)
        strSQL = " SELECT COUNT(CARG_ACPTNO) AS Cnt " & Chr(13)
        strSQL = strSQL & " FROM CARG_FILE " & Chr(13)
        strSQL = strSQL & " WHERE CARG_ACPTNO = (SELECT TOP 1 CARG_ACPTNO " & Chr(13)
        strSQL = strSQL & "                         FROM CARG_FILE " & Chr(13)
        strSQL = strSQL & "                        WHERE CARG_YCODE = " & lngYardCode & Chr(13)
        strSQL = strSQL & "                          AND CARG_NO    = " & lngCntaCode & Chr(13)
        'INSERT 2018/05/23 N.IMAI Start
        If strUkno <> "" Then
          strSQL = strSQL & "                        AND CARG_UKNO <> '" & strUkno & "'" & Chr(13)    ' チェック除外の予約番号の指定
        End If
        'INSERT 2018/05/23 N.IMAI End
        strSQL = strSQL & "                        ORDER BY CARG_FSDATE DESC) " & Chr(13)
        ' 最新契約は初回契約開始日が一番大きいものとする  →△△△△△△ 2006/02/04 H.Tajima
        strSQL = strSQL & "   AND CARG_AGRE <> " & pcintAGRE_09
        
        '▽▼2006/01/21 ADD Tajima▽▼
        If strACPTNO <> "" Then
          strSQL = strSQL & "   AND CARG_ACPTNO <> '" & strACPTNO & "'"    ' チェック除外の契約番号の指定
        End If
        '△▲2006/01/21 ADD Tajima△▲
        
        Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

        intCount = 0
        intCount = Nz(objRs.Fields("Cnt"), 0)
        objRs.Close
        Set objRs = Nothing

        If intCount > 0 Then
            ' 契約済みチェックエラー
            intKbn = pcintKEIYAKU_Err
            strErrMsg = pcstrKEIYAKU_ErrMsg
            GoTo CheckLending_Exit
        Else
            intKbn = pcintCHECK_OK
        End If
    End If

    ' 稼働ヤードチェック
    If (intCheck And pcintYARDKADO_UNITCHECK) > 0 Then
        ' (地主契約トランに、稼働ヤードがあるか否かのチェック)
        intKbn = pcintKADO_Err
        '▼ 2006/09/27 add tajima ▼
'        strSql = "SELECT COUNT(JARG_YCODE) AS Cnt " & Chr(13)
'        strSql = strSql & "  FROM JARG_FILE " & Chr(13)
'        strSql = strSql & " WHERE JARG_YCODE = " & lngYardCode & Chr(13)  ' ヤードコード
'        strSql = strSql & "   AND (JARG_KYDATE Is Null " & Chr(13)
'        strSql = strSql & "    OR  JARG_KYDATE > '" & Format(DATE, "YYYY/MM/DD") & "') "
        ' ↑地主契約は参照せず、↓ヤード解約予定日で判断する 2006/09/27
        strSQL = "SELECT COUNT(YARD_CODE) AS Cnt " & Chr(13)
        strSQL = strSQL & "  FROM YARD_MAST " & Chr(13)
        strSQL = strSQL & " WHERE YARD_CODE = " & lngYardCode & Chr(13)  ' ヤードコード
        strSQL = strSQL & "   AND ISNULL(YARD_END_DAY,'9999/12/31') > GETDATE()"
        '▲ 2006/09/27 add tajima ▲
       
        Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

        intCount = 0
        intCount = Nz(objRs.Fields("Cnt"), 0)
        objRs.Close
        Set objRs = Nothing

        If intCount = 0 Then
          ' 稼働チェックエラー
            intKbn = pcintKADO_Err
            strErrMsg = pcstrKADO_ErrMsg
            GoTo CheckLending_Exit
        Else
            intKbn = pcintCHECK_OK
        End If
    End If

    ' 解約予定コンテナチェック
    If (intCheck And pcintKAIYAKU_UNITCHECK) > 0 Then
        ' (コンテナ契約ファイルに同一のヤードコード、コンテナ番号かつ、解約予定日が空で
        '  契約中のものの存在をチェック)
        strSQL = " SELECT COUNT(CARG_ACPTNO) Cnt" & Chr(13)
        strSQL = strSQL & " FROM CARG_FILE " & Chr(13)
        strSQL = strSQL & " WHERE CARG_ACPTNO = (SELECT TOP 1 CARG_ACPTNO " & Chr(13)
        strSQL = strSQL & "                         FROM CARG_FILE " & Chr(13)
        strSQL = strSQL & "                      WHERE CARG_YCODE = " & lngYardCode & Chr(13)
        strSQL = strSQL & "                        AND CARG_NO    = " & lngCntaCode & Chr(13)
        strSQL = strSQL & "                      ORDER BY CARG_FSDATE DESC) " & Chr(13)
        ' 最新契約は初回契約開始日が一番大きいものとする→△△△△△△ 2006/02/04 H.Tajiam
        strSQL = strSQL & " AND CARG_AGRE = " & pcintAGRE_01 & Chr(13)
        strSQL = strSQL & " AND CARG_KYDATE IS NULL"
        Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

        intCount = 0
        intCount = Nz(objRs.Fields("Cnt"), 0)
        objRs.Close
        Set objRs = Nothing

        If intCount > 0 Then
            ' 解約予定の契約ではない
            intKbn = pcintKEIYAKU_Err
            strErrMsg = pcstrKAIYOTEI_ErrMsg
            GoTo CheckLending_Exit
        Else
            intKbn = pcintCHECK_OK
        End If
    
    End If

CheckLending_Exit:
    CheckLending = intKbn
    Exit Function

CheckLending_Err:
    strErrMsg = "ｴﾗｰ番号:" & Err.Number & vbCrLf & Err.Description
    CheckLending = pcintSYSTEM_ERROR
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        Parameter        :strErrMsg = エラーメッセージ格納
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncConnectDB(ByRef strErrMsg As String) As Boolean
'
'    Dim strBumoCD     As String
'
'    Dim strConnect    As String
'    Dim strDataSource As String
'
'    On Error GoTo fncConnectDB_Err
'
'    fncConnectDB = False
'
'    ' 部門コード取得
'    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
'
'    ' -----------------------------------------------------------------------------------------------------------
'    ' コンテナDB接続
'    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumoCD & "'"), "")
'
'    If strDataSource = "" Then
'        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, PROG_ID
'        Set pobjKONT_DB = Nothing
'    Else
'        Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumoCD))
'    End If
'    ' -----------------------------------------------------------------------------------------------------------
'
'    fncConnectDB = True
'
'    Exit Function
'
'fncConnectDB_Err:
'    strErrMsg = "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
'    Err.Clear
'End Function

'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :fncDisConnectDB
'        Parameter        :strErrMsg = エラーメッセージ格納
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncDisConnectDB(ByRef strErrMsg As String) As Boolean
'
'    On Error GoTo fncDisConnectDB_Err
'
'    fncDisConnectDB = False
'
'    ' コンテナDB
'    pobjKONT_DB.Close
'    Set pobjKONT_DB = Nothing
'
'    fncDisConnectDB = True
'
'    Exit Function
'
'fncDisConnectDB_Err:
'    strErrMsg = "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
'    Err.Clear
'End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤード最終使用日付の取得
'                         :このメソッドは指定ヤードの地主契約ファイルから一番未来の解約日付を
'                         :特定し、その解約日付－１を返します。
'        MODULE_ID        :GetYardUseEndDate
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(Long) = ヤードコード
'                         :第3引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :最終使用日付(yyyy/mm/ddのString or ""※何時までも使用可)
'        CREATE_DATE      :2005/12/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetYardUseEndDate(dbSQLServer As DAO.Database, _
                                  lngYardCode As Long, _
                                  ByRef strMessage As String) As String

On Error GoTo GetYardUseEndDate_Err

    Dim lastDate  As Variant
    Dim strSQL    As String
    Dim objRs     As Recordset
    
    strMessage = ""
    '解約日付の一番大きいものを取る。NULLはMAX日として2100/12/31としてみる
    strSQL = "SELECT TOP 1 " & Chr(13)
    strSQL = strSQL & "ISNULL(JARG_KYDATE, CONVERT(datetime,'2100/12/31') ) MAXDATE ,JARG_KYDATE " & Chr(13)
    strSQL = strSQL & "FROM JARG_FILE "
    strSQL = strSQL & "WHERE JARG_YCODE = " & lngYardCode & Chr(13)  ' ヤードコード
    strSQL = strSQL & "ORDER BY MAXDATE DESC"
    
    '対象の検索を行う
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
    lastDate = Nz(objRs.Fields("JARG_KYDATE"), "")
    objRs.Close
    
    '対象がNULLでなければその解約日を見る
    If lastDate <> "" Then
      lastDate = DateAdd("d", -1, lastDate) '解約日の前日までが最終使用日付
    End If
    
    GetYardUseEndDate = CStr(lastDate)
    Exit Function
    
GetYardUseEndDate_Err:
    strMessage = "GetYardUseEndDateｴﾗｰｺｰﾄﾞ:" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ:" & Err.Description
    Err.Clear

End Function
'☆
'==============================================================================*
'
'        MODULE_NAME      :紹介待ち作成－単一コンテナ指定
'        MODULE_ID        :IntroWaitSingle
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(String) = 予約紹介区分
'                         :第4引数(ByRef String) = 異常終了時にエラーメッセージ格納
'                         :第5引数(Long) = ヤードコード
'                         :第6引数(Long) = コンテナ番号
'                         :第7引数(String)[省略可] = 更新元の紹介区分（省略時は「取置き」）
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function IntroWaitSingle(dbSQLServer As DAO.Database, _
                                strBookingNo As String, _
                                strSTATUS As String, _
                                ByRef strMessage As String, _
                                lngYardCode As Long, _
                                lngContainerNo As Long, _
                                Optional strSourceStatus As String = "") As Boolean
On Error GoTo ErrRtn
    
    IntroWaitSingle = False
    strMessage = ""
    
    '【コンテナマスタ更新】
    Call UpdateCnta(dbSQLServer, _
                    strBookingNo, _
                    lngYardCode, _
                    lngContainerNo, _
                    pcintProcMode_単一コンテナ指定)
    
    '【予約紹介トラン更新】
    Call UpdateIntr(dbSQLServer, _
                    strBookingNo, _
                    lngYardCode, _
                    lngContainerNo, _
                    strSTATUS, _
                    pcintProcMode_単一コンテナ指定, _
                    strSourceStatus)
    
    IntroWaitSingle = True
    GoTo EndRtn
    
ErrRtn:
    strMessage = "IntroWaitSingle(" & Err.Number & ")" & Err.Description
    Err.Clear

EndRtn:

End Function

'==============================================================================*
'
'        MODULE_NAME      :紹介待ち作成－全取消し
'        MODULE_ID        :IntroWaitAll
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(ByRef String) = 異常終了時にエラーメッセージ格納
'                         :第4引数(Variant)[省略可能] = ヤードコード
'                         :第5引数(Variant)[省略可能] = コンテナ番号
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function IntroWaitAll(dbSQLServer As DAO.Database, _
                             strBookingNo As String, _
                             ByRef strMessage As String, _
                             Optional varYardCode As Variant = Null, _
                             Optional varContainerNo As Variant = Null) As Boolean
On Error GoTo ErrRtn
    
    IntroWaitAll = False
    strMessage = ""
    
    '【コンテナマスタ更新】
    Call UpdateCnta(dbSQLServer, _
                    strBookingNo, _
                    varYardCode, _
                    varContainerNo, _
                    pcintProcMode_全取消し)
    
    '【予約紹介トラン更新】
    Call UpdateIntr(dbSQLServer, _
                    strBookingNo, _
                    varYardCode, _
                    varContainerNo, _
                    pcstrIntroSts_CANCEL, _
                    pcintProcMode_全取消し)
    
    IntroWaitAll = True
    GoTo EndRtn
    
ErrRtn:
    strMessage = "IntroWaitAll(" & Err.Number & ")" & Err.Description
    Err.Clear

EndRtn:

End Function

'==============================================================================*
'
'        MODULE_NAME      :紹介待ち作成－紹介区分変更＆紹介待ち作成
'        MODULE_ID        :IntroChgAndIntroWait
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(String) = 予約紹介区分
'                         :第4引数(ByRef String) = 異常終了時にエラーメッセージ格納
'                         :第5引数(Long) = ヤードコード
'                         :第6引数(Long) = コンテナ番号
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function IntroChgAndIntroWait(dbSQLServer As DAO.Database, _
                                     strBookingNo As String, _
                                     strSTATUS As String, _
                                     ByRef strMessage As String, _
                                     lngYardCode As Long, _
                                     lngContainerNo As Long) As Boolean
On Error GoTo ErrRtn
    
    IntroChgAndIntroWait = False
    strMessage = ""
    
    '【コンテナマスタ更新】
    Call UpdateCnta(dbSQLServer, _
                    strBookingNo, _
                    lngYardCode, _
                    lngContainerNo, _
                    pcintProcMode_紹介区分変更_紹介待ち作成)
    
    '【予約紹介トラン更新】
    Call UpdateIntr(dbSQLServer, _
                    strBookingNo, _
                    lngYardCode, _
                    lngContainerNo, _
                    strSTATUS, _
                    pcintProcMode_紹介区分変更_紹介待ち作成)
    
    IntroChgAndIntroWait = True
    GoTo EndRtn
    
ErrRtn:
    strMessage = "IntroChgAndIntroWait(" & Err.Number & ")" & Err.Description
    Err.Clear

EndRtn:

End Function

'==============================================================================*
'
'        MODULE_NAME      :紹介トラン新規作成
'        MODULE_ID        :CmnInsertIntrTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(Long) = ヤードコード
'                         :第4引数(Long) = コンテナ番号
'                         :第5引数(String) = 予約紹介区分
'                         :第6引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2005/12/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CmnInsertIntrTran(dbSQLServer As DAO.Database, _
                                  strBookingNo As String, _
                                  lngYardCode As Long, _
                                  lngContainerNo As Long, _
                                  strSTATUS As String, _
                                  ByRef strMessage As String) As Boolean
On Error GoTo ErrRtn
    
    CmnInsertIntrTran = False
    strMessage = ""
    
    '【予約紹介トラン更新】
    Call UpdateIntr(dbSQLServer, _
                    strBookingNo, _
                    lngYardCode, _
                    lngContainerNo, _
                    strSTATUS, _
                    pcintProcMode_単純追加, _
                    strSTATUS)
    
    CmnInsertIntrTran = True
    GoTo EndRtn
    
ErrRtn:
    strMessage = "CmnInsertIntrTran(" & Err.Number & ")" & Err.Description
    Err.Clear

EndRtn:

End Function

'==============================================================================*
'
'        MODULE_NAME      :紹介トラン紹介区分更新
'        MODULE_ID        :CmnInsertIntrTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(String) = 更新元紹介区分
'                         :第4引数(String) = 更新後紹介区分
'                         :第5引数(ByRef String) = 異常終了時にエラーメッセージ格納
'                         :第6引数(Variant)[省略可] = 紹介番号（指定時は更新元区分は無視）
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2005/12/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CmnUpdateIntrTran(dbSQLServer As DAO.Database, _
                                  strBookingNo As String, _
                                  strSourceStatus As String, _
                                  strDestinationStatus As String, _
                                  ByRef strMessage As String, _
                                  Optional varIntroNo As Variant = Null) As Boolean
On Error GoTo ErrRtn
    
    CmnUpdateIntrTran = False
    strMessage = ""
    
    '【予約紹介トラン更新】
    Call UpdateIntr(dbSQLServer, _
                    strBookingNo, _
                    Null, _
                    Null, _
                    strDestinationStatus, _
                    pcintProcMode_単純更新, _
                    strSourceStatus, _
                    varIntroNo)
    
    CmnUpdateIntrTran = True
    GoTo EndRtn
    
ErrRtn:
    strMessage = "CmnUpdateIntrTran(" & Err.Number & ")" & Err.Description
    Err.Clear

EndRtn:

End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ更新
'        MODULE_ID        :CmnUpdateCntaMast
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(Long) = ヤードコード
'                         :第3引数(Long) = コンテナ番号
'                         :第4引数(ByRef String) = 異常終了時にエラーメッセージ格納
'                         :第5引数(String)[省略可能] = 予約番号
'                         :                   ・省略時は利用可否を「可」にして
'                         :                   　紹介待ちにした予約番号をNULLにする。
'                         :                   ・指定時は省略時は利用可否を「否」にして
'                         :                   　紹介待ちにした予約番号を設定する。
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2005/12/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CmnUpdateCntaMast(dbSQLServer As DAO.Database, _
                                  lngYardCode As Long, _
                                  lngContainerNo As Long, _
                                  ByRef strMessage As String, _
                                  Optional strBookingNo As String = "") As Boolean
On Error GoTo ErrRtn
    
    CmnUpdateCntaMast = False
    strMessage = ""
    
    '【コンテナマスタ更新】
    Call UpdateCnta(dbSQLServer, _
                    strBookingNo, _
                    lngYardCode, _
                    lngContainerNo, _
                    pcintProcMode_単一コンテナ指定)
    
    CmnUpdateCntaMast = True
    GoTo EndRtn
    
ErrRtn:
    strMessage = "CmnUpdateCntaMast(" & Err.Number & ")" & Err.Description
    Err.Clear

EndRtn:

End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ更新
'        MODULE_ID        :UpdateCnta
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(Variant) = ヤードコード
'                         :第4引数(Variant) = コンテナ番号
'                         :第5引数(Integer) = 処理モード
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub UpdateCnta(dbSQLServer As DAO.Database, _
                       strBookingNo As String, _
                       varYardCode As Variant, _
                       varContainerNo As Variant, _
                       intProcMode As Integer)
    
    Dim rsCnta          As Recordset
    Dim strKey          As String
    Dim blnFound        As Boolean
    
    '全取消しモードでコンテナが指定されている場合、
    '紹介トランに存在しないコンテナは後で検索して更新する
    '紹介トランに存在する場合はフラグはTrueになる。
    blnFound = False
    
    strKey = Nz(varYardCode, "") & Nz(varContainerNo, "")
    
    '１．更新対象検索
    Set rsCnta = dbSQLServer.OpenRecordset( _
                        MakeCntaSql(strBookingNo, varYardCode, varContainerNo, intProcMode), _
                        dbOpenDynaset)
    
    '２．コンテナマスタの利用可否区分を「否」に更新
    With rsCnta
        While Not .EOF
            '紹介区分変更モードの場合は、指定されたコンテナは更新しない
            If Not (intProcMode = pcintProcMode_紹介区分変更_紹介待ち作成 And _
                   (.Fields("CNTA_CODE") & .Fields("CNTA_NO")) = strKey) _
            Then
                Call SetCntaFields(rsCnta, strBookingNo)
            End If
            
            '指定コンテナ？
            If (.Fields("CNTA_CODE") & .Fields("CNTA_NO")) = strKey Then
                blnFound = True
            End If
            
            .MoveNext
        Wend
        
        .Close
    End With

    '全取消しモードで指定コンテナが見つからない場合、そのコンテナを検索して更新する
    If intProcMode = pcintProcMode_全取消し And Not blnFound And Nz(varYardCode, "") <> "" Then
        Set rsCnta = dbSQLServer.OpenRecordset( _
                            MakeCntaSql(strBookingNo, varYardCode, varContainerNo, pcintProcMode_単一コンテナ指定), _
                            dbOpenDynaset)
        
        If Not rsCnta.EOF Then
            Call SetCntaFields(rsCnta, strBookingNo)
        End If
        rsCnta.Close
        Set rsCnta = Nothing
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ検索SQL作成
'        MODULE_ID        :UpdateCnta
'        Parameter        :第1引数(String) = 予約番号
'                         :第2引数(Variant) = ヤードコード
'                         :第3引数(Variant) = コンテナ番号
'                         :第4引数(Integer) = 処理モード
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeCntaSql(strBookingNo As String, _
                             varYardCode As Variant, _
                             varContainerNo As Variant, _
                             intProcMode As Integer) As String

    Dim strSQL      As String
    
    '条件は、利用可否＝「可」
    strSQL = strSQL & " SELECT CNTA_CODE, "
    strSQL = strSQL & "        CNTA_NO, "
    strSQL = strSQL & "        CNTA_UPDATE, "
    strSQL = strSQL & "        CNTA_USE, "
    strSQL = strSQL & "        CNTA_UKNO, "
    strSQL = strSQL & "        CNTA_USAGE "   '2017/07/29 M.HONDA INS
    strSQL = strSQL & "   FROM CNTA_MAST "
    strSQL = strSQL & "  WHERE  "
    
    If intProcMode = pcintProcMode_単一コンテナ指定 Then
        '単一コンテナ指定モード
        'ヤードコード＆コンテナ番号が一致するデータを検索する
        strSQL = strSQL & "        CNTA_CODE = " & varYardCode
        strSQL = strSQL & "    AND CNTA_NO = " & varContainerNo
    Else
        'それ以外のモード
        '紹介トランに存在するコンテナを検索する
        strSQL = strSQL & "    EXISTS ( "
        
        strSQL = strSQL & MakeIntrSql(strBookingNo, Null, Null, intProcMode)
        
        
        strSQL = strSQL & "    AND INTRT_YCODE = CNTA_CODE "
        strSQL = strSQL & "    AND INTRT_NO = CNTA_NO ) "
    End If
    
    MakeCntaSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ値設定
'        MODULE_ID        :SetCntaFields
'        Parameter        :第1引数(Recordset) = 更新対象
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetCntaFields(rsCnta As Recordset, strBookingNo As String)

    With rsCnta
        .Edit
        
        .Fields("CNTA_UPDATE") = DATE
        
        If strBookingNo = "" Then
            .Fields("CNTA_USE") = pcintCNTA_USE_OK   '「可」
            .Fields("CNTA_UKNO") = Null
        Else
            .Fields("CNTA_USE") = pcintCNTA_USE_NO   '「否」
            .Fields("CNTA_UKNO") = strBookingNo
        End If
        
        '2017/07/29 M.HONDA INS
        If .Fields("CNTA_USAGE") = "33" Then
            .Fields("CNTA_USAGE") = 0
        End If
        '2017/07/29 M.HONDA INS
        
        '2019/06/15 N.IMAI Start
        If .Fields("CNTA_USAGE") = "39" Then
            .Fields("CNTA_USAGE") = 3
        End If
        '2019/06/15 N.IMAI End
        
        .UPDATE
    End With

End Sub

'==============================================================================*
'
'        MODULE_NAME      :紹介トラン更新
'        MODULE_ID        :UpdateIntr
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(Variant) = ヤードコード
'                         :第4引数(Variant) = コンテナ番号
'                         :第5引数(String) = 予約紹介区分
'                         :第6引数(Integer) = 処理モード
'                         :第7引数(String) = 更新元の紹介区分
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub UpdateIntr(dbSQLServer As DAO.Database, _
                       strBookingNo As String, _
                       varYardCode As Variant, _
                       varContainerNo As Variant, _
                       strSTATUS As String, _
                       intProcMode As Integer, _
                       Optional strSourceStatus As String = "", _
                       Optional varIntroNo As Variant = Null)
    
    Dim rsIntr          As Recordset
    Dim strUserName     As String
    Dim strProgramName  As String
    Dim blnInsert       As Boolean
    Dim strKey          As String
    Dim strKbn          As String
    Dim blnFound        As Boolean
    
    strKey = Nz(varYardCode, "") & Nz(varContainerNo, "")

    strUserName = LsGetUserName()
    strProgramName = GetProgramName()
    
    '１．更新対象検索
    Set rsIntr = dbSQLServer.OpenRecordset( _
                        MakeIntrSql(strBookingNo, _
                                    varYardCode, _
                                    varContainerNo, _
                                    intProcMode, _
                                    strSourceStatus, _
                                    varIntroNo), _
                        dbOpenDynaset)
    
    '２．ご紹介区分を、引数で指定された区分に更新
    If intProcMode = pcintProcMode_単一コンテナ指定 Then
        '単一コンテナ指定モード
        'EOFならば追加、
        blnInsert = rsIntr.EOF
    Else
        'それ以外のモード
        blnInsert = True
        
        While Not rsIntr.EOF
            '単純追加モード時は、更新は行わない。対象データが存在するか否かのみを見る。
            If intProcMode <> pcintProcMode_単純追加 Then
                '紹介区分変更モードの場合は、指定されたコンテナだけパラメータ指定の区分にする
                'それ以外は「キャンセル」
                If intProcMode = pcintProcMode_紹介区分変更_紹介待ち作成 Then
                    If (rsIntr.Fields("INTRT_YCODE") & rsIntr.Fields("INTRT_NO")) = strKey Then
                        strKbn = strSTATUS
                    Else
                        strKbn = pcstrIntroSts_CANCEL
                    End If
                Else
                    strKbn = strSTATUS
                End If
                
                '紹介トラン更新
                Call SetIntrFields(dbSQLServer, _
                                   rsIntr, _
                                   strBookingNo, _
                                   varYardCode, _
                                   varContainerNo, _
                                   strKbn, _
                                   strUserName, _
                                   strProgramName, _
                                   False)
            End If
            
            '指定コンテナ？
            If (rsIntr.Fields("INTRT_YCODE") & rsIntr.Fields("INTRT_NO")) = strKey Then
                blnFound = True
            End If
            
            rsIntr.MoveNext
        Wend
    End If
    
    '単一コンテナ指定
    'もしくは全取消しか紹介区分変更か単純追加モードで指定コンテナが見つからなかった場合
    If intProcMode = pcintProcMode_単一コンテナ指定 Or _
       ( _
         ( _
           intProcMode = pcintProcMode_全取消し Or _
           intProcMode = pcintProcMode_紹介区分変更_紹介待ち作成 Or _
           intProcMode = pcintProcMode_単純追加 _
         ) And _
         Not blnFound And _
         Nz(varYardCode, "") <> "" _
       ) _
    Then
        '最後のパラメータがTrueならば新規追加する
        Call SetIntrFields(dbSQLServer, _
                           rsIntr, _
                           strBookingNo, _
                           varYardCode, _
                           varContainerNo, _
                           strSTATUS, _
                           strUserName, _
                           strProgramName, _
                           blnInsert)
    
    End If
    
    rsIntr.Close
    Set rsIntr = Nothing

End Sub

'==============================================================================*
'
'        MODULE_NAME      :紹介トラン検索SQL作成
'        MODULE_ID        :UpdateIntr
'        Parameter        :第1引数(String) = 予約番号
'                         :第2引数(Variant) = ヤードコード
'                         :第3引数(Variant) = コンテナ番号
'                         :第4引数(Integer) = 処理モード
'                         :第5引数(String) = 更新元の紹介区分（省略時は「取置き」）
'                         :第6引数(Variant) = 紹介番号（指定時は第2,3,5引数は無視）
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MakeIntrSql(strBookingNo As String, _
                             varYardCode As Variant, _
                             varContainerNo As Variant, _
                             intProcMode As Integer, _
                             Optional strSourceStatus As String = "", _
                             Optional varIntroNo As Variant = Null) As String

    Dim strSQL      As String
    Dim strSign     As String
    

    '条件は、予約番号＆ヤードコード＆コンテナ番号＆紹介区分＝「取置き」
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM INTR_TRAN "
    strSQL = strSQL & "  WHERE INTRT_UKNO  = '" & strBookingNo & "' "
    
    If Nz(varIntroNo, "") = "" Then
        '紹介番号省略
        If Nz(strSourceStatus, "") = "" Then
            strSQL = strSQL & "    AND INTRT_INTROKBN = '" & pcstrIntroSts_KEEP & "' "  '「取置き」
        Else
            strSQL = strSQL & "    AND INTRT_INTROKBN = '" & strSourceStatus & "' "  '指定された区分
        End If
        '単一コンテナ指定モード時はヤード＆コンテナが一致することが条件
        If intProcMode = pcintProcMode_単一コンテナ指定 Then
            strSQL = strSQL & "    AND INTRT_YCODE = " & varYardCode
            strSQL = strSQL & "    AND INTRT_NO = " & varContainerNo
        End If
    Else
        '紹介番号指定
        strSQL = strSQL & " AND INTRT_INTRONO = " & varIntroNo
    End If
    
    MakeIntrSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :紹介トラン値設定
'        MODULE_ID        :SetIntrFields
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 更新対象RecordSet
'                         :第3引数(String) = 予約番号
'                         :第4引数(Variant) = ヤードコード
'                         :第5引数(Variant) = コンテナ番号
'                         :第6引数(String) = 予約紹介区分
'                         :第7引数(String) = 更新ユーザーID
'                         :第8引数(String) = 更新プログラムID
'                         :第9引数(Boolean) = Trueなら新規追加 Falseなら変更
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetIntrFields(dbSQLServer As DAO.Database, _
                          rsIntr As Recordset, _
                          strBookingNo As String, _
                          varYardCode As Variant, _
                          varContainerNo As Variant, _
                          strSTATUS As String, _
                          strUserName As String, _
                          strProgramName As String, _
                          blnAddNew As Boolean)

    Dim lngIntroNo      As Long
    
    With rsIntr
        If blnAddNew Then
            '新規紹介番号発番
            lngIntroNo = GetNewIntroNo(dbSQLServer, strBookingNo)
            
            '紹介トラン新規作成
            .AddNew
            .Fields("INTRT_UKNO") = strBookingNo
            .Fields("INTRT_INTRONO") = lngIntroNo
            .Fields("INTRT_YCODE") = varYardCode
            .Fields("INTRT_NO") = varContainerNo
            .Fields("INTRT_NEARKBN") = "0"
            .Fields("INTRT_INSED") = Format$(DATE, "yyyymmdd")
            .Fields("INTRT_INSEJ") = Format$(time, "hhmmss")
            .Fields("INTRT_INSPB") = strProgramName
            .Fields("INTRT_INSUB") = strUserName
            '2006.2.3 SHIBAZAKI 追加
            .Fields("INTRT_FOUTD") = DATE
        Else
            '紹介トラン変更
            .Edit
            .Fields("INTRT_UPDAD") = Format$(DATE, "yyyymmdd")
            .Fields("INTRT_UPDAJ") = Format$(time, "hhmmss")
            .Fields("INTRT_UPDPB") = strProgramName
            .Fields("INTRT_UPDUB") = strUserName
        End If
        .Fields("INTRT_INTROKBN") = strSTATUS   '紹介区分
        .UPDATE
    End With
        
End Sub

'==============================================================================*
'
'        MODULE_NAME      :新規紹介番号発番
'        MODULE_ID        :SetIntrFields
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'        Return           :初番した紹介番号
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetNewIntroNo(dbSQLServer As DAO.Database, strBookingNo As String) As Long

    Dim rsMaxIntroNo        As Recordset
    
    GetNewIntroNo = 0

    '予約番号を条件に、紹介番号の最大値を検索(一件も存在しない場合、NULLが検索される)
    Set rsMaxIntroNo = dbSQLServer.OpenRecordset( _
                            "SELECT MAX(INTRT_INTRONO) AS MAX_NO " & _
                            "  FROM INTR_TRAN " & _
                            " WHERE INTRT_UKNO = '" & strBookingNo & "'", _
                            dbOpenSnapshot, dbSQLPassThrough)
    
    '最大値に１加算
    GetNewIntroNo = Nz(rsMaxIntroNo.Fields("MAX_NO"), 0) + 1

    rsMaxIntroNo.Close
    Set rsMaxIntroNo = Nothing
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :呼び出し元プログラムID取得
'        MODULE_ID        :GetProgramName
'        Return           :プログラムID
'        CREATE_DATE      :2005/12/6
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetProgramName() As String
On Error Resume Next

    Dim objCallFrom As Object
    Dim objParent   As Object
    
    '呼び出し元Object取得
    Set objCallFrom = Application.CodeContextObject
    
    If objCallFrom Is Nothing Then
        '呼び出し元が、フォームやレポートではない場合、（標準モジュールから呼び出されたりとか）
        'CodeContextObjectプロパティを参照できないので、この標準モジュールのIDを返却する。
        GetProgramName = PROG_ID
    Else
        '名前を返却
        '呼び出し元がサブフォームならば親フォームの名前を返却
        Set objParent = objCallFrom.Parent
        If objParent Is Nothing Then
            '親フォームが存在しない
            GetProgramName = objCallFrom.NAME
        Else
            '親フォームが存在する
            GetProgramName = objParent.NAME
        End If
    End If
    
    If Not objCallFrom Is Nothing Then
        Set objCallFrom = Nothing
    End If
    If Not objParent Is Nothing Then
        Set objParent = Nothing
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :指定したコンテナを紹介待ちにする
'        MODULE_ID        :WaitUpdateCntaMast
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(Long) = ヤードコード
'                         :第4引数(Long) = コンテナ番号
'                         :第5引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2005/12/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function WaitUpdateCntaMast(dbSQLServer As DAO.Database, _
                                   strBookingNo As String, _
                                   lngYardCode As Long, _
                                   lngContainerNo As Long, _
                                   ByRef strMessage As String) As Boolean

    Dim strSQL As String
    Dim rsCnta As Recordset

On Error GoTo ErrRtn

    WaitUpdateCntaMast = False
    strMessage = ""

    ' コンテナマスタの対象データ取得
    strSQL = strSQL & " SELECT CNTA_USE, "
    strSQL = strSQL & "        CNTA_UKNO, "
    strSQL = strSQL & "        CNTA_UPDATE "
    strSQL = strSQL & "   FROM CNTA_MAST "
    strSQL = strSQL & "WHERE CNTA_CODE  = " & lngYardCode & Chr(13)     ' ヤードコード
    strSQL = strSQL & "  AND CNTA_NO    = " & lngContainerNo & Chr(13)  ' コンテナ番号
    Set rsCnta = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    With rsCnta
        .Edit

        .Fields("CNTA_USE") = pcintCNTA_USE_NO   '「否」
        .Fields("CNTA_UPDATE") = DATE
        .Fields("CNTA_UKNO") = strBookingNo

        .UPDATE
    End With
    rsCnta.Close
    Set rsCnta = Nothing

    WaitUpdateCntaMast = True
    GoTo EndRtn

ErrRtn:
    strMessage = "CmnUpdateIntrTran(" & Err.Number & ")" & Err.Description
    Err.Clear

EndRtn:
End Function
'==============================================================================*
'
'        MODULE_NAME      :旧顧客マスタの更新
'                         :既に新顧客コードが設定されていたら何もしない
'        MODULE_ID        :UpdateOldUserCode
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 設定する新顧客コード
'                         :第3引数(String) = 対象顧客コード
'                         :第4引数(ByRef String) = 格納メッセージ
'        Return           :True = 正常終了  False = 異常終了
'        Note             :必要なら上位呼び出し先でトランザクションをすること
'        CREATE_DATE      :2006/01/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UpdateOldUserMast(dbSQLServer As DAO.Database, _
                                   ByVal aNewUserCode As String, _
                                   ByVal anOldUserCode As String, _
                                   ByRef aReturnMessage As String _
                                  ) As Boolean
    Dim strSQL As String
    Dim rsUser As Recordset

On Error GoTo Exception

    aReturnMessage = ""

    ' コンテナマスタの対象データ取得
    strSQL = strSQL & "SELECT USER_NEW_CODE, USER_UPDATE "
    strSQL = strSQL & "FROM   USER_MAST "
    strSQL = strSQL & "WHERE  USER_CODE  = " & anOldUserCode
    Set rsUser = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    With rsUser
        ' 新顧客コードに値が設定されていなければ新顧客コードを設定する
        If Nz(.Fields("USER_NEW_CODE"), "") = "" Then
          .Edit
          .Fields("USER_UPDATE") = DATE
          .Fields("USER_NEW_CODE") = aNewUserCode
          .UPDATE
        End If
    End With

    UpdateOldUserMast = True
    GoTo Finally

Exception:
    UpdateOldUserMast = False
    aReturnMessage = "UpdateOldUserCode(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not rsUser Is Nothing Then rsUser.Close: Set rsUser = Nothing
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :予約受付番号発番処理
'                         :※この処理はFVS500.UkeNo_Numbering をエレガントにPublic化Publicしたものである
'                         :※処理内容は全く同じなので、あるタイミングで統一したいにゃー
'        MODULE_ID        :NumberingUKNO
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef String) = 発番した受付番号
'                         :※識別(1)+部門ｺｰﾄﾞ(1)+受付年(2)+受付月(1)+ｼｰｹﾝｽ(6)[11桁]
'                         :第3引数(ByRef String) = 格納メッセージ
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2006/01/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function NumberingUKNO(dbSQLServer As DAO.Database, _
                                  ByRef anUkNo As String, _
                                  ByRef aReturnMessage As String _
                                ) As Boolean

    Dim strSQL As String
    Dim rsObject As Recordset
    Dim lngSequenceNo As Long   '予約シーケンス
    Dim strBumonCode As String  '部門コード
    Dim strYY       As String   '年
    Dim strMM       As String   '月
    Dim strUkno     As String   '生成される予約番号

On Error GoTo Exception
    
    anUkNo = ""
    aReturnMessage = ""
    
    '①コントロールマスタから使用する項目を取得
    strSQL = "SELECT CONT_BUMOC, CONT_YOYA_NO, CONT_UPDATE FROM CONT_MAST "
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    If rsObject.EOF Then
      'コントロールマスタが無い！？
      aReturnMessage = "NumberingUKNO()にて例外発生：コントロールマスタが存在しない"
      NumberingUKNO = False
      GoTo Finally
    Else
    '②受付番号を発番して取得
        With rsObject
            .Edit
            .Fields("CONT_YOYA_NO") = .Fields("CONT_YOYA_NO") + 1   '受付ｼｰｹﾝｽ№CountUp
            .Fields("CONT_UPDATE") = Now()                          '更新日
            lngSequenceNo = .Fields("CONT_YOYA_NO")                 '予約シーケンゲットだぜ
            strBumonCode = .Fields("CONT_BUMOC")                    '部門ｺｰﾄﾞゲットだぜ
            .UPDATE             ' ここでちゃっちゃと更新して
            .Close              ' そしてクローズして
        End With
        Set rsObject = Nothing  ' 解放しちゃいマホ
    End If

    '③受付番号の生成
    '識別子セット
    strUkno = "U" '予約受付番号の識別子 'U' ※定数はあえて使わない。理由はあえて述べない
    '部門コードセット
    strUkno = strUkno & strBumonCode
    '受付年２桁セット
    Select Case Format(DATE, "yyyy")
        Case "1998": strYY = "&&"
        Case "1999": strYY = "**"
        Case Else
            If Mid$(Trim$(Format$(DATE, "yyyy")), 1, 2) = "20" Then
                strYY = Mid$(Trim$(Format$(DATE, "yyyy")), 3)
            End If
    End Select
    strUkno = strUkno & strYY
    '受付月セット
    Select Case Format(DATE, "m")
        Case "10": strMM = "A"
        Case "11": strMM = "B"
        Case "12": strMM = "C"
        Case Else
            strMM = Trim$(Format$(DATE, "m"))
    End Select
    strUkno = strUkno & strMM
    'シーケンスＮＯセット
    strUkno = strUkno & CStr(Format(lngSequenceNo, "000000"))

    ' 生成した予約受付番号を渡す
    anUkNo = strUkno

    NumberingUKNO = True
    GoTo Finally

Exception:
    NumberingUKNO = False
    aReturnMessage = "NumberingUKNO(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing

End Function
'==============================================================================*
'
'        MODULE_NAME      :期限日取得
'                         :※この処理はFVS500.subSetTKDATE をよりエレガントにPublic化したものである
'                         :※処理内容は全く同じなので、あるタイミングで統一したいにゃー
'        MODULE_ID        :GetTimeLimitDay
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByVal String) = 取得タイプ
'                         :第3引数(ByVal Date)   = 基準日
'                         :第3引数(ByRef Date)   = 取得した期限日
'                         :第4引数(ByRef String) = 格納メッセージ
'        Return           :True = 正常終了  False = 異常終了
'        Note             :期限日はヤード解約日以前でないといけない。
'                         :別途ヤードの解約期限日を取得し比較すること！
'        CREATE_DATE      :2006/01/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetTimeLimitDay(dbSQLServer As DAO.Database, _
                                  ByVal aGetType As String, _
                                  ByVal aBaseDate As Date, _
                                  ByRef aTLD As Date, _
                                  ByRef aReturnMessage As String _
                                ) As Boolean
                                
    Dim strSQL As String
    Dim rsObject As Recordset
    Dim intTORI_DD       As Integer   '取置期限日数
    Dim intKEIUKE_TLD    As Integer   '契約受付期限日数
    Dim intJISYARIYU_TLM As Integer   '自社理由取置期限月数

On Error GoTo Exception
    
    aReturnMessage = ""
    intTORI_DD = 0
    intKEIUKE_TLD = 0
    intJISYARIYU_TLM = 0
    
    '①コントロールマスタから使用する項目を取得
    strSQL = "  SELECT  CONT_TORI_DD, " & Chr(13)
    strSQL = strSQL & " CONT_KEIUKE_TLD, " & Chr(13)
    strSQL = strSQL & " CONT_JISYARIYU_TLM " & Chr(13)
    strSQL = strSQL & " FROM CONT_MAST "
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    Do Until rsObject.EOF
        intTORI_DD = Nz(rsObject.Fields("CONT_TORI_DD"), 0)              ' 取置期限日数のゲット
        intKEIUKE_TLD = Nz(rsObject.Fields("CONT_KEIUKE_TLD"), 0)        ' 契約受付期限日数のゲット
        intJISYARIYU_TLM = Nz(rsObject.Fields("CONT_JISYARIYU_TLM"), 0)  ' 自社理由取置期限月数のゲット
        rsObject.MoveNext
    Loop
    
    '②コントロールマスタから取得した日数をもとに基準日から算出した日付を取得
    Select Case aGetType
        Case pcstrYUKBN_01  ' 自社理由
            aTLD = DateAdd("m", intJISYARIYU_TLM, aBaseDate)  ' 自社理由取置期限月数
        Case pcstrYUKBN_02  ' 取置
            aTLD = DateAdd("d", intTORI_DD, aBaseDate)        ' 取置期限日数
        Case pcstrYUKBN_10  ' 受付
            aTLD = DateAdd("d", intKEIUKE_TLD, aBaseDate)     ' 契約受付期限日数
        Case Else
            GetTimeLimitDay = False
            aReturnMessage = "GetTimeLimitDay()：認識できない取得タイプが指定されています→" & aGetType
            GoTo Finally
    End Select
    
    GetTimeLimitDay = True
    GoTo Finally

Exception:
    GetTimeLimitDay = False
    aReturnMessage = "GetTimeLimitDay(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
                                
End Function
'==============================================================================*
'
'        MODULE_NAME      :解約ヤード用の予約状態１（自社理由）の予約情報作成
'                         :解約ヤード用に予約受付トラン＆紹介トランにて自社理由の予約情報を作成する                        :紹介トランを作成する
'        MODULE_ID        :AddnReserveType1forCancelYard
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = ヤードコード
'                         :第3引数(String) = コンテナコード
'                         :第4引数(String) = 担当者コード
'                         :第5引数(String) = 作成プログラムID
'                         :第6引数(String) = 作成新ユーザID
'                         :第7引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :True = 正常終了  False = 異常終了
'        Note             :このFunction内でトランザクションを行っているので注意
'        CREATE_DATE      :2006/01/25
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function AddnReserveType1forCancelYard(dbSQLServer As DAO.Database, _
                                   ByVal aYardCode As String, _
                                   ByVal aContainerCode As String, _
                                   ByVal anUketukeTantoCode As String, _
                                   ByVal aMakingProgramID As String, _
                                   ByVal aMakingUserID As String, _
                                   ByRef aReturnMessage As String _
                                  ) As Boolean
    Dim strSQL As String
    Dim strReturnMessage As String
    Dim rsObject As Recordset
    Dim wsObject  As Workspace  'トランザクション処理の為のWorkSpaceオブジェクト
    Dim dbgTrap As String '2008/03/25 add tajima

On Error GoTo Exception
    strReturnMessage = ""
    dbgTrap = "0" '2008/03/25 add tajima

    Set wsObject = Nothing

    '①先に使用する予約番号を発番する
    Dim strBookingNo As String             '予約番号
    If NumberingUKNO(dbSQLServer, strBookingNo, strReturnMessage) = False Then
      strReturnMessage = "AddYoukOnReserveType1(" & strReturnMessage & ")"
      GoTo Exception
    End If
    
    dbgTrap = "1" '2008/03/25 add tajima
    
    'α．トランザクションの開始
    Set wsObject = DBEngine.Workspaces(0)
    wsObject.BeginTrans
    
    dbgTrap = "2" '2008/03/25 add tajima
    
    '②自社理由期限日を取得する
    Dim dtTLD As Date                     '自社理由期限日
    If GetTimeLimitDay(dbSQLServer, pcstrYUKBN_01, DATE, dtTLD, strReturnMessage) = False Then
      strReturnMessage = "AddYoukOnReserveType1(" & strReturnMessage & ")"
      GoTo Exception
    End If
    
    dbgTrap = "3" '2008/03/25 add tajima
    
    '③ヤード最終使用日付の取得し､②で取得した期限日とどちらを使用するか判断する
    Dim strYardLastDate As String
    strYardLastDate = GetYardUseEndDate(dbSQLServer, CLng(aYardCode), strReturnMessage)
    If Nz(strYardLastDate, "") <> "" Then strYardLastDate = Format(strYardLastDate, "YYYY/MM/DD")

    If Nz(strYardLastDate, "") <> "" Then
        ' ヤード最終使用日付が取得できたとき
        ' 取得期限日以前にヤードが解約されてしまうならば（ヤード最終使用日付 ＜ 取得した期限日）
        If CDate(strYardLastDate) < dtTLD Then
            dtTLD = CDate(strYardLastDate) 'ヤード最終使用日付が期限日となる
        End If
    End If
    
    dbgTrap = "4" '2008/03/25 add tajima
    
    '④予約受付トランにパラメータのヤード・コンテナで自社理由予約があるかチェックする
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM   YOUK_TRAN "
    strSQL = strSQL & "WHERE  YOUKT_YCODE  = " & aYardCode
    strSQL = strSQL & " AND   YOUKT_NO     = " & aContainerCode
    strSQL = strSQL & " AND   YOUKT_YUKBN  = " & pcstrYUKBN_01
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    dbgTrap = "5" '2008/03/25 add tajima
    
    If rsObject.EOF = True Then
    '⑤無ければ予約受付トランの新規作成を行う
      With rsObject
        .AddNew
        .Fields("YOUKT_UKNO") = strBookingNo                ' 予約受付番号　※①で発番したモノ
        .Fields("YOUKT_UKDATE") = DATE                      ' 予約受付日
        .Fields("YOUKT_UKTANTO") = anUketukeTantoCode       ' 予約受付担当者コード
        .Fields("YOUKT_UKKBN") = 99                         ' 受付区分 = 99
        .Fields("YOUKT_YCODE") = aYardCode                  ' ヤードコード
        .Fields("YOUKT_NEARKBN") = 0                        ' 近隣ヤード指定=0
        .Fields("YOUKT_NO") = aContainerCode                ' コンテナＮｏ
        .Fields("YOUKT_YUKBN") = pcstrYUKBN_01              ' 予約状態 = 1:自社理由
        .Fields("YOUKT_KUDATE") = DATE                      ' 契約受付日
        .Fields("YOUKT_TKDATE") = dtTLD                     ' 取置期限日　※②③で決めたモノ
        .Fields("YOUKT_KIBOUSU") = 1                        ' 希望本数 = 1
        .Fields("YOUKT_NAME") = "解約の為、貸禁"            ' 顧客名
        .Fields("YOUKT_KANA") = "ｶｲﾔｸﾉﾀﾒ ｶｼｷﾝ"              ' 顧客名カナ
        .Fields("YOUKT_TEL") = "0"                          ' 電話番号 = "0"
        .Fields("YOUKT_KKBN") = 1                           ' 顧客区分 = 1
        .Fields("YOUKT_RENKBN") = 0                         ' 連絡区分 = 0
        .Fields("YOUKT_AUTOKBN") = 0                        ' 自動区分 = 0
        .Fields("YOUKT_MOVEKBN") = 0                        ' 移動区分 = 0
        .Fields("YOUKT_GENKBN") = 0                         ' 発生区分 = 0      '' 2012/07/07 M.HONDA INS
        .Fields("YOUKT_INSED") = Format$(DATE, "yyyymmdd")  ' 作成日
        .Fields("YOUKT_INSEJ") = Format$(time, "hhmmss")     ' 作成時間
        .Fields("YOUKT_INSPB") = aMakingProgramID           ' 作成プログラムＩＤ
        .Fields("YOUKT_INSUB") = aMakingUserID              ' 作成ユーザＩＤ
        .UPDATE
      End With
      
    dbgTrap = "6" '2008/03/25 add tajima
      
    '⑥紹介トランの作成
      If CmnInsertIntrTran(dbSQLServer, _
                            strBookingNo, _
                            CLng(aYardCode), _
                            CLng(aContainerCode), _
                            pcstrIntroSts_KEEP, _
                            strReturnMessage) = False Then
        strReturnMessage = "AddYoukOnReserveType1(" & strReturnMessage & ")"
        GoTo Exception
      End If
    Else
    dbgTrap = "7" '2008/03/25 add tajima
    
    '⑦既に自社理由の予約があれば何もしない
      strReturnMessage = "AddYoukOnReserveType1(既に自社理由予約があります)"
    End If
    dbgTrap = "8" '2008/03/25 add tajima
    
    '⑧処理正常終了
    AddnReserveType1forCancelYard = True
    'β．コミット
    wsObject.CommitTrans
    dbgTrap = "9" '2008/03/25 add tajima
    
    GoTo Finally

Exception:
    'γ．ロールバック
    wsObject.Rollback
    'エラー処理
    AddnReserveType1forCancelYard = False
    If strReturnMessage = "" Then
      strReturnMessage = "AddYoukOnReserveType1(" & Err.Number & ")dbgTrap:" & dbgTrap & Err.Description
    End If
    Err.Clear

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    'If Not wsObject Is Nothing Then wsObject.Close: Set wsObject = Nothing
    aReturnMessage = strReturnMessage
End Function

'==============================================================================*
'        MODULE_NAME      :移動元コンテナ契約トラン更新
'        MODULE_ID        :UpdateMoveOldCARG
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 更新対象の契約番号
'                         :第3引数(String) = 移動理由(移動予約受付トラン・移動区分）
'                         :第4引数(Date)   = 新契約開始日（画面で入力された契約開始日）
'                         :第5引数(Date)   = 予約受付日(移動予約受付トラン・受付日）
'                         :第6引数(String) = 受付担当コード(移動予約受付トラン・受付担当者コード）
'                         :第7引数(String) = 備考
'                         :第8引数(String) = 更新プログラムID
'                         :第9引数(String) = 更新ユーザID
'                         :第10引数(ByRef String) = エラーメッセージ格納
'                         :第11引数(Optional String) = 鍵変更理由コード
'        戻り値
'                           0            = 正常終了
'                          -1            = システムエラー
'        CREATE_DATE      :2006/01/28
'        UPDATE_DATE      :2014/06/20 移動元鍵区分変換対応のため、第11引数追加し、鍵変更理由コードを設定する処理を追加
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UpdateMoveOldCARG(dbSQLServer As DAO.Database, _
                              ByVal anACPTNO As String, _
                              ByVal aMovReasonCode As String, _
                              ByVal aNewContractDate As Date, _
                              ByVal aReservationDate As Date, _
                              ByVal aUketukeTantoCode As String, _
                              ByVal aNote As String, _
                              ByVal anUpdateProgramID As String, _
                              ByVal anUpdateUserID As String, _
                              ByRef anErrMsg As String, _
                              Optional anKeyChgRiyuCd As Integer = -1 _
                              ) As Integer

    Dim strSQL         As String
    Dim objRs          As Recordset
    Dim strAGRE        As String   ' 契約状態
    Dim varCONTNA      As Variant  ' コンテナ状態
    Dim strKAICD       As String   ' 解約区分
    Dim strKAIRIYUCD   As String   ' 解約理由区分コード
    Dim varMNT_REQDATE As Variant  ' メンテ依頼日
    Dim varMNT_TANTO   As Variant  ' メンテ担当者コード
    Dim strTargetDate  As String
    
    '' 20101220 Ver1.2 M.HONDA START
    Dim datLIQREQDATE  As Variant  ' 清算依頼日
    Dim strLIQFLG      As String   ' 清算済フラグ
    '' 20101220 Ver1.2 M.HONDA END

    On Error GoTo UpdateMoveOldCARG_Err
    UpdateMoveOldCARG = pcintSYSTEM_ERROR  ' 初初期化
    anErrMsg = ""

    If anACPTNO = "" Then Exit Function

    ' コンテナ契約ファイルの対象データ抽出
    strSQL = "  SELECT  CARG_AGRE, " & Chr(13)         ' 契約状態
    strSQL = strSQL & " CARG_STDATE, " & Chr(13)       ' 契約開始日
    strSQL = strSQL & " CARG_CYDATE, " & Chr(13)       ' 解約予定日
    strSQL = strSQL & " CARG_KYDATE, " & Chr(13)       ' 解約日付
    strSQL = strSQL & " CARG_CONTNA, " & Chr(13)       ' コンテナ状態
    strSQL = strSQL & " CARG_UPDATE, " & Chr(13)       ' 更新日
    strSQL = strSQL & " CARG_KAIUKEDATE, " & Chr(13)   ' 解約受付日
    strSQL = strSQL & " CARG_KAITANTO, " & Chr(13)     ' 解約担当者コード
    strSQL = strSQL & " CARG_KAICD, " & Chr(13)        ' 解約区分
    strSQL = strSQL & " CARG_KAIRIYUCD, " & Chr(13)    ' 解約理由区分コード
    strSQL = strSQL & " CARG_MNT_CMPDATE, " & Chr(13)  ' メンテ完了日
    strSQL = strSQL & " CARG_MNT_TANTO, " & Chr(13)    ' メンテ担当者コード
    strSQL = strSQL & " CARG_MNT_MEMO, " & Chr(13)     ' メンテメモ
    
    '' 20101220 Ver1.2 M.HONDA START
    strSQL = strSQL & " CARG_LIQ_REQDATE, " & Chr(13)  ' 清算依頼日
    strSQL = strSQL & " CARG_LIQ_FLG, " & Chr(13)      ' 清算済フラグ
    '' 20101220 Ver1.2 M.HONDA END
    
'↓ INSERT 2014/06/20 MIYAMOTO
    strSQL = strSQL & " CARG_KEY_CHGRIYUCD, " & Chr(13) ' 鍵変更理由コード
'↑ INSERT 2014/06/20 MIYAMOTO
    
    strSQL = strSQL & " CARG_UPDAD, " & Chr(13)        ' 更新日付
    strSQL = strSQL & " CARG_UPDAJ, " & Chr(13)        ' 更新時刻
    strSQL = strSQL & " CARG_UPDPB, " & Chr(13)        ' 更新プログラムID
    strSQL = strSQL & " CARG_UPDUB " & Chr(13)         ' 更新ユーザーID
    strSQL = strSQL & " FROM CARG_FILE " & Chr(13)
    strSQL = strSQL & " WHERE CARG_ACPTNO = '" & anACPTNO & "' "
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    With objRs
        Do Until .EOF
            .Edit

            ' 第3引数(移動区分)により各値を特定
            strAGRE = ""
            varCONTNA = Null
            strKAICD = "": strKAIRIYUCD = ""
            varMNT_REQDATE = Null: varMNT_TANTO = Null
            
            '' 20101220 Ver1.2 M.HONDA START
            datLIQREQDATE = Null
            strLIQFLG = 0
            '' 20101220 Ver1.2 M.HONDA END
            
            Select Case aMovReasonCode
                Case pcstrMOVEKBN_01, pcstrMOVEKBN_02
                    ' 自社理由 OR 顧客理由
                    strAGRE = pcstrAGRE_02             ' 解約予定
                    strKAICD = pcstrKAICD_04           ' 移動
                    strKAIRIYUCD = pcstrKAIRIUCD_07    ' 加瀬の別のボックスに移動

                Case pcstrMOVEKBN_03
                    ' 未使用移動
                    strAGRE = pcstrAGRE_02             ' 解約予定
                    strKAICD = pcstrKAICD_03           ' 未使用
                    strKAIRIYUCD = pcstrKAIRIUCD_07    ' 加瀬の別のボックスに移動


                '▼2010/12/20 Ver1.2 M.HONDA
                '名義変更
                Case pcstrMOVEKBN_04
                    strAGRE = pcstrAGRE_09                                         ' 完了
                    varCONTNA = -1                                                 ' コンテナ状態
                    strKAICD = pcstrKAICD_05                                       ' 契約変更
                    strKAIRIYUCD = pcstrKAIRIUCD_99                                ' その他
                    varMNT_REQDATE = aReservationDate                              ' メンテ依頼日(メンテ完了日)
                    varMNT_TANTO = aUketukeTantoCode
                    ''datLIQREQDATE = CDate(Format(DateAdd("m", 1, DATE), "yyyy/mm/") + "05") ' 清算依頼日      '' 2013/07/08 M.HONDA DEL
                    ''strLIQFLG = 0                                                           ' 清算済フラグ    '' 2013/07/08 M.HONDA DEL
                '▲2010/12/20 Ver1.2 M.HONDA

                '▼2008/03/25 add tajima
                '  契約変更系として一括りにする
                'Case pcstrMOVEKBN_04, pcstrMOVEKBN_05, pcstrMOVEKBN_06
                Case Else
                '▲2008/03/25 add tajima
                    ' 名義変更 OR 金額変更 OR 保障会社変更 OR 紹介者金額変更
                    strAGRE = pcstrAGRE_09             ' 完了
                    varCONTNA = -1                     ' コンテナ状態
                    strKAICD = pcstrKAICD_05           ' 契約変更
                    strKAIRIYUCD = pcstrKAIRIUCD_99    ' その他
                    varMNT_REQDATE = aReservationDate  ' メンテ依頼日(メンテ完了日)
                    varMNT_TANTO = aUketukeTantoCode
            End Select
            
            
            '' 2013/07/08 M.HONDA INS
            '' 自社理由移動は清算依頼日をセットする。
            Select Case aMovReasonCode
                Case pcstrMOVEKBN_01
                    datLIQREQDATE = Null
                    strLIQFLG = 0
                Case Else
                    datLIQREQDATE = CDate(Format(DateAdd("m", 1, DATE), "yyyy/mm/") + "05") ' 清算依頼日
                    strLIQFLG = 0                                                           ' 清算済フラグ
            End Select
            '' 2013/07/08 M.HONDA INS
            

            ' 契約状態セット
            .Fields("CARG_AGRE") = strAGRE

            ' 解約予定日・解約日付
            strTargetDate = ""
            Select Case aMovReasonCode
                ' 2005/02/07 DEL H.Tajima ▼▽ 顧客理由移動も前月末日！
                ' Case pcstrMOVEKBN_02,
                '     ' システム日付の1ヶ月後
                '     strTargetDate = DateAdd("m", 1, DATE)
                ' 2005/02/07 DEL H.Tajima △▲
                Case pcstrMOVEKBN_03
                    ' 移動元(自分自身)契約開始日の前月末日)
                    If IsNull(.Fields("CARG_STDATE")) = False Then
                        strTargetDate = .Fields("CARG_STDATE")
                        strTargetDate = DateSerial(Year(strTargetDate), _
                                                   Month(strTargetDate), _
                                                   1) - 1
                    End If
                '▼2008/03/25 add tajima
                '  契約変更系として一括りにする
                'Case pcstrMOVEKBN_01, pcstrMOVEKBN_02, pcstrMOVEKBN_04, pcstrMOVEKBN_05, pcstrMOVEKBN_06
                Case Else
                '▲2008/03/25 add tajima
                    ' 新契約開始日の前月末日
                    strTargetDate = DateSerial(Year(aNewContractDate), _
                                               Month(aNewContractDate), _
                                               1) - 1
            End Select

            .Fields("CARG_CYDATE") = Format(strTargetDate, "YYYY/MM/DD")              ' 解約予定日
            .Fields("CARG_KYDATE") = Format(strTargetDate, "YYYY/MM/DD")              ' 解約日付

            If IsNull(varCONTNA) = False Then
                .Fields("CARG_CONTNA") = varCONTNA                                    ' コンテナ状態
            End If

            .Fields("CARG_UPDATE") = Format(DATE, "YYYY/MM/DD")                       ' 更新日(システム日付)
            .Fields("CARG_KAIUKEDATE") = Format(aReservationDate, "YYYY/MM/DD")       ' 解約受付日(第5引数)
            .Fields("CARG_KAITANTO") = aUketukeTantoCode                              ' 解約担当者コード(第6引数)
                .Fields("CARG_KAICD") = strKAICD                                          ' 解約区分
            .Fields("CARG_KAIRIYUCD") = strKAIRIYUCD                                  ' 解約理由区分コード
            
            '' 20101220 Ver1.2 M.HONDA START
            .Fields("CARG_LIQ_REQDATE") = datLIQREQDATE         ' 清算依頼日
            .Fields("CARG_LIQ_FLG") = strLIQFLG                 ' 清算済フラグ
            '' 20101220 Ver1.2 M.HONDA END

            ' メンテ完了日
            If IsNull(varMNT_REQDATE) = False Then
                .Fields("CARG_MNT_CMPDATE") = Format(varMNT_REQDATE, "YYYY/MM/DD")    ' メンテ完了日(第5引数)
            End If

            ' メンテ担当者コード(第6引数)
            If IsNull(varMNT_TANTO) = False Then
                .Fields("CARG_MNT_TANTO") = varMNT_TANTO
            End If

            .Fields("CARG_MNT_MEMO") = aNote                                          ' メンテメモ(第7引数)
            .Fields("CARG_UPDAD") = Format(DATE, "YYYYMMDD")                          ' 更新日付
            .Fields("CARG_UPDAJ") = Format(time, "HHMMSS")                            ' 更新時刻
            .Fields("CARG_UPDPB") = anUpdateProgramID                                 ' 更新プログラムID(第7引数)
            .Fields("CARG_UPDUB") = anUpdateUserID                                    ' 更新ユーザーID(第8引数)

'↓ INSERT 2014/06/20 MIYAMOTO
            If anKeyChgRiyuCd <> -1 Then
                .Fields("CARG_KEY_CHGRIYUCD") = anKeyChgRiyuCd                        ' 鍵変更理由コード(第11引数)
            End If
'↑ INSERT 2014/06/20 MIYAMOTO

            .UPDATE

            .MoveNext
        Loop
    End With
    objRs.Close
    Set objRs = Nothing

    UpdateMoveOldCARG = pcintCHECK_OK

UpdateMoveOldCARG_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

UpdateMoveOldCARG_Err:
    anErrMsg = "ｴﾗｰ番号:" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo UpdateMoveOldCARG_Exit
End Function

'==============================================================================*
'        MODULE_NAME      :移動予約可否
'        MODULE_ID        :IsMoveReserveEntry
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 移動理由(移動予約受付トラン・移動区分）
'                         :第3引数(String) = 移動元契約番号（画面で選択されている契約番号）
'                         :第4引数(String) = 移動先(新)ヤードコード（画面で選択されたヤードコード）
'                         :第5引数(String) = 移動先(新)コンテナ番号(画面で選択されたヤードコード）
'                         :第6引数(ByRef String) = エラーメッセージ格納
'                         :第7引数(ByRef integer) = 契約変更の回数                              'INSERT 2019/03/06 Y.WADA
'                         :第8引数(String)[省略可] = チェック除外の予約番号（省略時は全予約対象）
'        戻り値           true  = 予約可
'                         false = 予約不可（且つエラーメッセージ <> "" ならシステムエラー）
'        CREATE_DATE      :2006/01/30
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'INSERT 2019/03/06 Y.WADA Start
Public Function IsMoveReserveEntry(dbSQLServer As DAO.Database, _
                              ByVal aMoveReasonCode As String, _
                              ByVal anOldACPTNO As String, _
                              ByVal aNewYardCode As String, _
                              ByVal aNewContainerCode As String, _
                              ByRef aReturnMessage As String, _
                              ByRef aMotoACPTNOCnt As Integer, _
                              Optional strUkno As String = "" _
                              ) As Boolean
'INSERT 2019/03/06 Y.WADA End
'DELETE 2019/03/06 Y.WADA Start
'Public Function IsMoveReserveEntry(dbSqlServer As DAO.Database, _
'                              ByVal aMoveReasonCode As String, _
'                              ByVal anOldACPTNO As String, _
'                              ByVal aNewYardCode As String, _
'                              ByVal aNewContainerCode As String, _
'                              ByRef aReturnMessage As String, _
'                              Optional strUkno As String = "" _
'                              ) As Boolean
'DELETE 2019/03/06 Y.WADA End
    Dim strSQL As String
    Dim rsObject As Recordset
    Dim intCount As Integer
    Dim isRet As Boolean
    Dim strReturnMessage As String
    Dim dtSTDATE As Date
    Dim lngYCODE As Long
    Dim lngNo As Long
    Dim strMoveName As String
    
On Error GoTo Exception
    
    aReturnMessage = ""
    strReturnMessage = ""
    intCount = 0
    isRet = False
    
    '元コンテナ契約の情報を取得
    strSQL = "  SELECT  CARG_YCODE, CARG_NO, CARG_STDATE, " & Chr(13)
    strSQL = strSQL & " ( SELECT NAME_NAME FROM NAME_MAST "
    strSQL = strSQL & "   WHERE NAME_ID = '098' AND NAME_CODE = " & aMoveReasonCode
    strSQL = strSQL & " ) MOVE_NAME "
    strSQL = strSQL & "FROM CARG_FILE " & Chr(13)
    strSQL = strSQL & "WHERE CARG_ACPTNO = '" & anOldACPTNO & "'"
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    '取得した可否をチェック
    If rsObject.EOF = True Then
      isRet = False
      aReturnMessage = "指定された移動元契約[" & anOldACPTNO & "]は存在しません。"
      GoTo Finally
    End If
    
    '取得した情報を変数にゲット
    lngYCODE = Nz(rsObject.Fields("CARG_YCODE"), 0)     ' ヤード番号
    lngNo = Nz(rsObject.Fields("CARG_NO"), 0)           ' コンテナNO
    dtSTDATE = Nz(rsObject.Fields("CARG_STDATE"), "")   ' 初回契約開始日
    strMoveName = Nz(rsObject.Fields("MOVE_NAME"), "")  ' 移動区分名称
    rsObject.Close
    Set rsObject = Nothing
    
    '①元契約の初回契約日の翌月経っているかチェック
' 2006/05/12 Del Start tajima
' 予約での移動予約制限解除、受付入力での契約日付で制限を行う
'    Select Case aMoveReasonCode
'          ' コンテナ移動系：自社理由 OR 顧客理由 OR 未使用移動
'          ' の場合は、契約変更系のような縛りは不要
'          Case pcstrMOVEKBN_04, pcstrMOVEKBN_05, pcstrMOVEKBN_06
'          ' 契約変更系：名義変更 OR 金額変更 OR 保証会社変更
'          ' の場合は初回契約日の翌月１日以降ならば移動予約可能
'            If Format(DATE, "YYYY/MM") <= Format(dtSTDATE, "YYYY/MM") Then
'              isRet = False
'              strReturnMessage = strMoveName & "では契約開始日の翌月でなければ契約変更は出来ません。"
'              GoTo Finally
'            End If
'    End Select
' 2006/05/12 Del End tajima
    
    '②コンテナ契約での存在可否でチェック
    ' 前準備・・・
    If lngYCODE = CLng(aNewYardCode) And lngNo = CLng(aNewContainerCode) Then
      intCount = 1
    Else
      intCount = 0
    End If
    ' 実際のチェック開始
    Select Case aMoveReasonCode
          Case pcstrMOVEKBN_01, pcstrMOVEKBN_02, pcstrMOVEKBN_03
          ' コンテナ移動系：自社理由 OR 顧客理由 OR 未使用移動
            If intCount = 0 Then
              isRet = True  '移動元契約と違うヤード・コンテナを指定しているのでＯＫ
            Else
              strReturnMessage = strMoveName & "で同じヤード・コンテナは指定できません。"
            End If
         '▼2008/03/25 add tajima
         '  契約変更系として一括りにする
          'Case pcstrMOVEKBN_04, pcstrMOVEKBN_05, pcstrMOVEKBN_06
          Case Else
         '▲2008/03/25 add tajima
          ' 契約変更系：名義変更 OR 金額変更 OR 保証会社変更
            If intCount = 1 Then
              isRet = True  '移動元契約と同じヤード・コンテナを指定しているのでＯＫ
            Else
              strReturnMessage = strMoveName & "でヤード・コンテナは変更できません。"
            End If
    End Select
    ' チェック結果処理
    If isRet = False Then
      aReturnMessage = strReturnMessage
      GoTo Finally
    End If
    
    '③生きている移動予約があるかのチェック
    isRet = False
    strSQL = "SELECT  COUNT(YOUKT_MOTO_ACPTNO ) Cnt " & Chr(13)
    strSQL = strSQL & "FROM YOUK_TRAN " & Chr(13)
    strSQL = strSQL & "WHERE " & Chr(13)
    strSQL = strSQL & "  YOUKT_MOTO_ACPTNO = '" & anOldACPTNO & "' AND"
    strSQL = strSQL & "  YOUKT_YUKBN <> " & pcstrYUKBN_53
    If strUkno <> "" Then
      strSQL = strSQL & " AND YOUKT_UKNO <> '" & strUkno & "'" ' チェック除外の予約番号の指定
    End If
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    intCount = Nz(rsObject.Fields("Cnt"), 0)
    rsObject.Close
    Set rsObject = Nothing
    
'DELETE 2019/03/06 Y.WADA Start
'    If intCount = 0 Then
'      isRet = True  '０件なら移動予約はされていない
'    Else
'      strReturnMessage = "この契約では既に契約変更がされています。" & Chr(13) & "契約変更は１度しか出来ません。"
'    End If
'DELETE 2019/03/06 Y.WADA End
    
    'INSERT 2019/03/06 Y.WADA Start
    aMotoACPTNOCnt = intCount
    isRet = True
    'INSERT 2019/03/06 Y.WADA End
    
    GoTo Finally

Exception:
    strReturnMessage = "IsMoveReserveEntry(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    aReturnMessage = strReturnMessage
    IsMoveReserveEntry = isRet
End Function

'==============================================================================*
'
'        MODULE_NAME      :前契約の日付取得
'        MODULE_ID        :GetFormerContractKYADTE
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = ヤードコード
'                         :第3引数(String) = コンテナコード
'                         :第5引数(ByRef String) = 取得した解約日付
'                         :第4引数(ByRef String) = 格納メッセージ
'        Return           :取得した日付
'        Note             :新契約の契約開始日と評価する何かの日付を取得する
'        CREATE_DATE      :2006/01/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetFormerContractDate(dbSQLServer As DAO.Database, _
                                   ByVal aYardCode As String, _
                                   ByVal aContainerCode As String, _
                                   ByRef aKYDATE As Variant, _
                                   ByRef aReturnMessage As String _
                                  ) As Boolean
    Dim strSQL As String
    Dim rsObject As Recordset

On Error GoTo Exception

    aReturnMessage = ""

    ' とりあえず現状では前回契約の完了日付を取得します
    strSQL = " SELECT CARG_MNT_CMPDATE GET_DATE" & Chr(13)
    strSQL = strSQL & " FROM CARG_FILE " & Chr(13)
    strSQL = strSQL & " WHERE CARG_ACPTNO = (SELECT TOP 1 CARG_ACPTNO " & Chr(13)
    strSQL = strSQL & "                         FROM CARG_FILE " & Chr(13)
    strSQL = strSQL & "                        WHERE CARG_YCODE = " & aYardCode & Chr(13)
    strSQL = strSQL & "                          AND CARG_NO    = " & aContainerCode & Chr(13)
    strSQL = strSQL & "                        ORDER BY CARG_FSDATE DESC)"
    ' 最新の契約かの判断は初回契約開始日が一番大きいものとする

    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    If rsObject.EOF = False Then
      aKYDATE = Nz(rsObject.Fields("GET_DATE"), "")
    Else
      aKYDATE = ""
    End If
    
    GetFormerContractDate = True
    GoTo Finally

Exception:
    GetFormerContractDate = False
    aReturnMessage = "GetFormerContractDate(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    
End Function

'*********************** これ以降は 0.3 2006/02/06 Add *************************
'==============================================================================*
'
'        MODULE_NAME      :請求ファイル新規作成
'        MODULE_ID        :InsertRequFile
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(Long)    = 請求No
'                         :第3引数(String)  = 請求年月
'                         :第4引数(String)  = 請求対象年月（売上計上）
'                         :第5引数(Long)    = 顧客コード
'                         :第6引数(String)  = ヤードコード
'                         :第7引数(String)  = コンテナ番号
'                         :第8引数(Double)  = 金額
'                         :第9引数(Double)  = 消費税
'                         :第10引数(Double) = 保証金
'                         :第11引数(String) = 受注契約番号
'                         :第12引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2006/02/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function InsertRequFile(dbSQLServer As DAO.Database, _
                                   ByVal aReqNo As Long, _
                                   ByVal aSeikyuDate As String, _
                                   ByVal aKeijyoDate As String, _
                                   ByVal anUserCode As Long, _
                                   ByVal aYardCode As String, _
                                   ByVal aContainerCode As String, _
                                   ByVal aKingaku As Double, _
                                   ByVal aSyohizei As Double, _
                                   ByVal aSecukingaku As Double, _
                                   ByVal anACPTNO As String, _
                                   ByRef aReturnMessage As String _
                                  ) As Boolean

    Dim strSQL    As String
    Dim objRs     As Recordset
    Dim strNKDATE As String
    Dim dblTOTAL  As Double
    Dim dbl金額   As Double

    On Error GoTo Exception

    InsertRequFile = False
    aReturnMessage = ""

   ' 請求ファイルデータ抽出
    strSQL = "SELECT * FROM REQU_FILE "
    strSQL = strSQL & "WHERE REQU_ACPTNO = '" & anACPTNO & "' "                 'INSERT 2015/07/16 K.ISHIZAKA
'    strSQL = strSQL & "AND REQU_KJDATE = DateValue('" & aKeijyoDate & "') "    'DELETE 2015/07/26 K.ISHIZAKA 'INSERT 2015/07/16 K.ISHIZAKA
    strSQL = strSQL & "AND REQU_KJDATE = #" & aKeijyoDate & "# "                'INSERT 2015/07/26 K.ISHIZAKA
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)                'INSERT 2015/07/26 K.ISHIZAKA 'DELETE 2015/07/16 K.ISHIZAKA
'    Set objRs = dbSqlServer.OpenRecordset(strSQL, dbOpenDynaset, dbDenyWrite)  'DELETE 2015/07/26 K.ISHIZAKA 'INSERT 2015/07/16 K.ISHIZAKA

    With objRs
        'もし存在していたら終了処理へ
        If Not .EOF Then GoTo ExistsData                                        'INSERT 2015/07/16 K.ISHIZAKA
        .AddNew

        .Fields("REQU_REQNO") = aReqNo                                    ' 請求№
        .Fields("REQU_YYYYMM") = aSeikyuDate                              ' 請求年月
        .Fields("REQU_KJDATE") = aKeijyoDate                              ' 請求対象年月（売上計上）
        .Fields("REQU_SDATE") = aSeikyuDate                               ' 請求年月日

        ' 入金予定日
        strNKDATE = ""
        strNKDATE = Format(DateAdd("D", -1, CDate(aKeijyoDate)), "YYYY/MM/DD")
        .Fields("REQU_NKDATE") = strNKDATE

        .Fields("REQU_KCODE") = anUserCode                                ' 顧客コード

        .Fields("REQU_NKBN") = 0                                          ' 入金区分(0固定)
        .Fields("REQU_YCODE") = aYardCode                                 ' ﾔｰﾄﾞｺｰﾄﾞ
        .Fields("REQU_NO") = aContainerCode                               ' ｺﾝﾃﾅ番号
        .Fields("REQU_TTANKA") = 0                                        ' 坪単価(0固定)
        .Fields("REQU_TUBOSU") = 0                                        ' 坪数(0固定)

        '金額
        If aKingaku > pcsglMoneySumMaxValue Then
            dbl金額 = pcsglMoneySumMaxValue
        Else
            dbl金額 = aKingaku
        End If

        ' 消費税 2006/04/27 消費税は金額から内税で求める
        .Fields("REQU_SYOZEI") = GetIncludeTax(dbl金額, Format$(aKeijyoDate, "YYYYMM"))
        dbl金額 = dbl金額 - .Fields("REQU_SYOZEI")
        .Fields("REQU_KINGAK") = dbl金額
        
        .Fields("REQU_SECUKG") = aSecukingaku                             ' 保証金
        
        ' 合計(金額 + 消費税 + 保証金)
        dblTOTAL = dbl金額 + .Fields("REQU_SYOZEI") + aSecukingaku
        If dblTOTAL > pcsglMoneySumMaxValue Then
            .Fields("REQU_TOTAL") = pcsglMoneySumMaxValue
        Else
            .Fields("REQU_TOTAL") = dblTOTAL
        End If

        .Fields("REQU_TEKI") = Null                                       ' 摘要(Null固定)
        .Fields("REQU_FLG") = 1                                           ' 請求修正ﾌﾗｸﾞ(1固定)
        .Fields("REQU_UPDATE") = Format(DATE, "YYYY/MM/DD")               ' 更新日
        .Fields("REQU_ACPTNO") = anACPTNO                                 ' 受注契約番号

        .UPDATE
    End With
    
ExistsData:                                                                     'INSERT 2015/07/16 K.ISHIZAKA
    objRs.Close
    Set objRs = Nothing

    InsertRequFile = True
    GoTo Finally

Exception:
    aReturnMessage = "InsertRequFile(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
End Function

'==============================================================================*
'
'        MODULE_NAME      :予約受付トラン更新
'        MODULE_ID        :UpdateYoukTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String)  = 予約番号
'                         :第3引数(Integer) = 予約受付状態区分
'                         :第4引数(String)  = 更新プログラムID
'                         :第5引数(String)  = 更新ユーザID
'                         :第6引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2006/02/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UpdateYoukTran(dbSQLServer As DAO.Database, _
                               ByVal anUkNo As String, _
                               ByVal aYuKbn As Integer, _
                               ByVal anUpdateProgramID As String, _
                               ByVal anUpdateUserID As String, _
                               ByRef aReturnMessage As String) As Boolean

    Dim strSQL As String
    Dim objRs  As Recordset

    On Error GoTo Exception

    UpdateYoukTran = False
    aReturnMessage = ""

    ' 予約受付トランデータ抽出
    strSQL = " SELECT YOUKT_YUKBN, " & Chr(13)    ' 予約受付状態区分
    strSQL = strSQL & " YOUKT_UPDAD, " & Chr(13)  ' 更新日付
    strSQL = strSQL & " YOUKT_UPDAJ, " & Chr(13)  ' 更新時刻
    strSQL = strSQL & " YOUKT_UPDPB, " & Chr(13)  ' 更新プログラムID
    strSQL = strSQL & " YOUKT_UPDUB " & Chr(13)   ' 更新ユーザーID
    strSQL = strSQL & "  FROM YOUK_TRAN " & Chr(13)
    strSQL = strSQL & " WHERE YOUKT_UKNO = '" & anUkNo & "' "
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    If objRs.EOF = False Then
        With objRs
            .Edit
    
            .Fields("YOUKT_YUKBN") = aYuKbn
            .Fields("YOUKT_UPDAD") = Format$(DATE, "YYYYMMDD")
            .Fields("YOUKT_UPDAJ") = Format$(time, "HHMMSS")
            .Fields("YOUKT_UPDPB") = anUpdateProgramID
            .Fields("YOUKT_UPDUB") = anUpdateUserID
    
            .UPDATE
        End With
    End If
    objRs.Close
    Set objRs = Nothing

    UpdateYoukTran = True
    GoTo Finally

Exception:
    aReturnMessage = "UpdateYoukTran(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
End Function

'==============================================================================*
'
'        MODULE_NAME      :予約ご紹介トラン更新 ★この関数は使用しない 02/06 by tajima
'        MODULE_ID        :UpdateIntrTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String)  = 予約番号
'                         :第3引数(String)  = ヤードコード
'                         :第4引数(String)  = コンテナ番号
'                         :第5引数(String)  = ご紹介区分(From)
'                         :第6引数(String)  = ご紹介区分(To)
'                         :第7引数(String)  = 更新プログラムID
'                         :第8引数(String)  = 更新ユーザID
'                         :第9引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2006/02/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function UpdateIntrTran(dbSQLServer As DAO.Database, _
                               ByVal anUkNo As String, _
                               ByVal aYardCode As String, _
                               ByVal aContainerCode As String, _
                               ByVal anIntrokbn_F As String, _
                               ByVal anIntrokbn_T As String, _
                               ByVal anUpdateProgramID As String, _
                               ByVal anUpdateUserID As String, _
                               ByRef aReturnMessage As String) As Boolean

    Dim strSQL As String
    Dim objRs  As Recordset

    UpdateIntrTran = False
    aReturnMessage = ""

    ' 予約ご紹介トランデータ抽出
    strSQL = " SELECT INTRT_INTROKBN, " & Chr(13)  ' ご紹介区分
    strSQL = strSQL & " INTRT_UPDAD, " & Chr(13)   ' 更新日付
    strSQL = strSQL & " INTRT_UPDAJ, " & Chr(13)   ' 更新時刻
    strSQL = strSQL & " INTRT_UPDPB, " & Chr(13)   ' 更新プログラムID
    strSQL = strSQL & " INTRT_UPDUB " & Chr(13)    ' 更新ユーザーID
    strSQL = strSQL & "  FROM INTR_TRAN " & Chr(13)
    strSQL = strSQL & " WHERE INTRT_UKNO  = '" & anUkNo & "' " & Chr(13)    ' 予約番号
    strSQL = strSQL & "   AND INTRT_YCODE = " & aYardCode & Chr(13)         ' ヤードコード
    strSQL = strSQL & "   AND INTRT_NO    = " & aContainerCode              ' コンテナ番号
    strSQL = strSQL & "   AND INTRT_INTROKBN = '" & anIntrokbn_F & "' "     ' ご紹介区分
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    If objRs.EOF = False Then
        With objRs
            .Edit

            .Fields("INTRT_INTROKBN") = anIntrokbn_T
            .Fields("INTRT_UPDAD") = Format$(DATE, "YYYYMMDD")
            .Fields("INTRT_UPDAJ") = Format$(time, "HHMMSS")
            .Fields("INTRT_UPDPB") = anUpdateProgramID
            .Fields("INTRT_UPDUB") = anUpdateUserID

            .UPDATE
        End With
    End If
    objRs.Close
    Set objRs = Nothing

    UpdateIntrTran = True
    GoTo Finally

Exception:
    aReturnMessage = "UpdateIntrTran(" & Err.Number & ")" & Err.Description
    Err.Clear

Finally:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約ファイル新規作成
'        MODULE_ID        :InsertCargFile
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(ByRef構造体)  = 契約情報 CARG_FILE
'                         :第3引数(String)       = 更新プログラムID
'                         :第4引数(String)       = 更新ユーザID
'                         :第5引数(ByRef String) = 異常終了時にエラーメッセージ格納
'        Return           :True = 正常終了  False = 異常終了
'        CREATE_DATE      :2006/02/06
'        UPDATE_DATE      :2014/06/20 鍵変更理由コードを設定する処理を追加
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function InsertCargFile(dbSQLServer As DAO.Database, _
                                   ByRef aCARG_FILE As Type_CARG_FILE, _
                                   ByVal anUpdateProgramID As String, _
                                   ByVal anUpdateUserID As String, _
                                   ByRef aReturnMessage As String) As Boolean

    Dim strSQL As String
    Dim objRs  As Recordset
    Dim nowDate As Date
    
    On Error GoTo Exception
    
    aReturnMessage = ""
    nowDate = DATE

   ' コンテナ契約ファイルデータ抽出
    strSQL = "SELECT * FROM CARG_FILE "
    strSQL = strSQL & "WHERE CARG_ACPTNO = '" & aCARG_FILE.ACPTNO & "' "        'INSERT 2015/07/16 K.ISHIZAKA
'    Set objRs = dbSqlServer.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly) 'DELETE 2015/07/16 K.ISHIZAKA
'    Set objRs = dbSqlServer.OpenRecordset(strSQL, dbOpenDynaset, dbDenyWrite)  'DELETE 2015/07/26 K.ISHIZAKA 'INSERT 2015/07/16 K.ISHIZAKA
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)                'INSERT 2015/07/26 K.ISHIZAKA

    With objRs
'        .AddNew                                                                'DELETE 2015/07/16 K.ISHIZAKA
        If .EOF Then                                                            'INSERT START 2015/07/16 K.ISHIZAKA
            .AddNew
            .Fields("CARG_INSED") = Format$(nowDate, "YYYYMMDD") ' 作成日付
            .Fields("CARG_INSEJ") = Format$(time, "HHMMSS")     ' 作成時刻
            .Fields("CARG_INSPB") = anUpdateProgramID           ' 作成プログラムID
            .Fields("CARG_INSUB") = anUpdateUserID              ' 作成ユーザーID
            .Fields("CARG_CAMPC") = aCARG_FILE.CAMPC            ' 会社コード 2018/09/25 EGL INS
        Else
            .Edit
            .Fields("CARG_UPDAD") = Format$(nowDate, "YYYYMMDD") ' 更新日付
            .Fields("CARG_UPDAJ") = Format$(time, "HHMMSS")     ' 更新時刻
            .Fields("CARG_UPDPB") = anUpdateProgramID           ' 更新プログラムID
            .Fields("CARG_UPDUB") = anUpdateUserID              ' 更新ユーザーID
        End If                                                                  'INSERT END   2015/07/16 K.ISHIZAKA

        .Fields("CARG_YCODE") = aCARG_FILE.YCODE            ' ヤードコード
        .Fields("CARG_NO") = aCARG_FILE.No                  ' コンテナ番号
        .Fields("CARG_UCODE") = aCARG_FILE.UCODE            ' 顧客コード
        .Fields("CARG_AGRE") = aCARG_FILE.AGRE              ' 契約状態
        .Fields("CARG_FSDATE") = aCARG_FILE.FSDATE          ' 初回契約日
        .Fields("CARG_STDATE") = aCARG_FILE.STDATE          ' 契約開始日
        .Fields("CARG_EDDATE") = aCARG_FILE.EDDATE          ' 契約満了日
        .Fields("CARG_CYDATE") = aCARG_FILE.CYDATE          ' 解約予定日
        .Fields("CARG_KYDATE") = aCARG_FILE.KYDATE          ' 解約日付
        .Fields("CARG_CONTNA") = 0                          ' コンテナ状態(0固定)
        
        .Fields("CARG_USAGE") = aCARG_FILE.USAGE            ' 貸出用途                '' 2013/07/25 M.HONDA INS
        
        .Fields("CARG_DOCU1") = aCARG_FILE.DOCU1            ' 新契約書切替区分
        .Fields("CARG_DOCU2") = aCARG_FILE.DOCU2            ' 代引き区分
        .Fields("CARG_TUBOSU") = 0                          ' コンテナ坪数(0固定)
        .Fields("CARG_TTANKA") = 0                          ' コンテナ坪単価(0固定)
        ' 2006/04/27 内税対応 start tajima
        ' 消費税は月額料金から内税で求める
        .Fields("CARG_SYOZEI") = GetIncludeTax(aCARG_FILE.RENTKG, Format$(aCARG_FILE.STDATE, "YYYYMM"))
        .Fields("CARG_FSYOZEI") = GetIncludeTax(aCARG_FILE.FRSTKG, Format$(aCARG_FILE.STDATE, "YYYYMM"))
        ' 使用料は税抜き金額をいれる
        .Fields("CARG_RENTKG") = aCARG_FILE.RENTKG - .Fields("CARG_SYOZEI")   ' 月額料金
        .Fields("CARG_FRSTKG") = aCARG_FILE.FRSTKG - .Fields("CARG_FSYOZEI")  ' 初回金額
        ' 2006/04/27 End tajima
        .Fields("CARG_FRST_BILL") = aCARG_FILE.FRST_BILL    ' 初回請求金額 Add 2005/03/02 tajima
        .Fields("CARG_PREMTH_SUM") = aCARG_FILE.PREMTH_SUM  ' 前払月数  Add 2005/03/02 tajima
        .Fields("CARG_SECUKG") = aCARG_FILE.SECUKG          ' 預り金
        .Fields("CARG_BIKO") = aCARG_FILE.BIKO              ' 備考
        .Fields("CARG_UPDATE") = nowDate                    ' 更新日
        .Fields("CARG_HDATE") = Null                        ' 保険加入日(Null固定)
        .Fields("CARG_MDATE") = aCARG_FILE.MDATE            ' 代引伝票送付日
        .Fields("CARG_ACPTNO") = aCARG_FILE.ACPTNO          ' 受注契約番号
        .Fields("CARG_HOSYICD") = aCARG_FILE.HOSYICD        ' 保証区分
        .Fields("CARG_HOSYO_CD") = aCARG_FILE.HOSYO_CD      ' 保証会社コード 2009/04/01 Add
        .Fields("CARG_KAGIICD") = aCARG_FILE.KAGIICD        ' 鍵区分
        .Fields("CARG_DAHIB") = aCARG_FILE.DAHIB            ' 代引き伝票番号
        .Fields("CARG_HOKEB") = Null                        ' 保険番号(Null固定)
        .Fields("CARG_KAIJB") = aCARG_FILE.KAIJB            ' 鍵解除番号
        .Fields("CARG_HOKAI") = aCARG_FILE.HOKAI            ' 保証会社区分
        .Fields("CARG_HOSYD") = aCARG_FILE.HOSYD            ' 保証承認日
        .Fields("CARG_HOSYB") = aCARG_FILE.HOSYB            ' 保証承認番号
        .Fields("CARG_ADMEI") = aCARG_FILE.ADMEI            ' 広告媒体
        .Fields("CARG_KEITI") = aCARG_FILE.KEITI            ' 契約形態
        .Fields("CARG_KAGIA") = aCARG_FILE.KAGIA            ' 鍵代
        .Fields("CARG_UKNO") = aCARG_FILE.UKNO              ' 予約受付番号
        .Fields("CARG_UKTANTO") = aCARG_FILE.UKTANTO        ' 契約受付担当者コード
        .Fields("CARG_KAIUKEDATE") = aCARG_FILE.KAIUKEDATE  ' 解約受付日
        .Fields("CARG_KAITANTO") = aCARG_FILE.KAITANTO      ' 解約担当者コード
        If aCARG_FILE.KAICD <= 0 Then
            aCARG_FILE.KAICD = 1                            ' 解約区分の初期値は"1"です。
        End If
        .Fields("CARG_KAICD") = aCARG_FILE.KAICD            ' 解約区分
        If aCARG_FILE.KAIRIYUCD > 0 Then                    ' 解約理由=0ならテーブルには書かない
          .Fields("CARG_KAIRIYUCD") = aCARG_FILE.KAIRIYUCD  ' 解約理由区分コード
        End If
        .Fields("CARG_CMP_EXDATE") = aCARG_FILE.CMP_EXDATE  ' キャンペーン適用満了日 Add 2005/03/02 tajima
        .Fields("CARG_KEY_LENTNUM") = aCARG_FILE.KEY_LENTNUM ' 鍵貸出本数 Add 2005/05/24 tajima
        .Fields("CARG_SEIKI") = aCARG_FILE.SEIKI            ' 請求書発行区分 Add 2010/04/10 K.ISHIZAKA
'↓ INSERT 2014/06/20 MIYAMOTO
        If aCARG_FILE.KEY_CHGRIYUCD <> -1 Then
            .Fields("CARG_KEY_CHGRIYUCD") = aCARG_FILE.KEY_CHGRIYUCD    ' 鍵変更理由コード
        End If
'↑ INSERT 2014/06/20 MIYAMOTO
'        .Fields("CARG_INSED") = Format$(nowDate, "YYYYMMDD") ' 作成日付        'DELETE START 2015/07/16 K.ISHIZAKA
'        .Fields("CARG_INSEJ") = Format$(Time, "HHMMSS")     ' 作成時刻
'        .Fields("CARG_INSPB") = anUpdateProgramID           ' 作成プログラムID
'        .Fields("CARG_INSUB") = anUpdateUserID              ' 作成ユーザーID   'DELETE END   2015/07/16 K.ISHIZAKA

        .Fields("CARG_PLAN_CD") = aCARG_FILE.PLANCD         'プランコード対応 Add 2020/09/28 tajima

        .UPDATE
    End With
    objRs.Close
    Set objRs = Nothing
    InsertCargFile = True
    Exit Function

Exception:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    InsertCargFile = False
    Call Err.Raise(Err.Number, "InsertCargFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function


'==============================================================================*
'
'        MODULE_NAME      :新契約の「備考」の文言作成
'        MODULE_ID        :CreateNewMoveMemo
'        Parameter        :
'        Parameter        :第1引数(Integer)      = 移動区分
'                         :第2引数(String)       = 移動区分名称
'                         :第3引数(long)         = ヤードコード
'                         :第4引数(long)         = コンテナ番号
'                         :第5引数(long)         = 顧客コード
'                         :第6引数(String)       = 顧客名カナ
'                         :第7引数(Double)       = 月額料金
'                         :第8引数(String)       = 保証区分略称
'                         :第9引数(String)       = 受注契約番号
'                         :第10引数(ByRef String) = 生成された移動備考文言
'        Return           :True = 正常終了  False = 異常終了
'        Note             :新契約に書込む、旧契約の概要を作る
'        CREATE_DATE      :2006/02/16
''==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CreateNewMoveMemo(ByVal aMoveType As Integer, _
                                  ByVal aMoveTypeName As String, _
                                  ByVal aYardCode As Long, _
                                  ByVal aContainerNo As Long, _
                                  ByVal aUserCode As Long, _
                                  ByVal aUserKana As String, _
                                  ByVal aMonthlyFee As Double, _
                                  ByVal aGuarantyName As String, _
                                  ByVal anACPTNO As String, _
                                  ByRef aMoveMemo As String _
                                  ) As Boolean
    On Error GoTo Exception
    aMoveMemo = ""
    
    '①それぞれの文言組み立て
    Select Case aMoveType
        ' *** 自社理由移動 OR 顧客理由移動 OR 未使用移動
        Case pcstrMOVEKBN_01, pcstrMOVEKBN_02, pcstrMOVEKBN_03
            ' (旧契約ヤードコード-旧契約コンテナ番号)
            aMoveMemo = Format(aYardCode, "000000") & "-" & Format(aContainerNo, "000000")
            
        Case pcstrMOVEKBN_04  ' *** 名義変更 ***
            ' (旧契約番号「旧顧客コード:旧顧客カナ名様」)
            aMoveMemo = anACPTNO & "「" & Format(aUserCode, "000000") & ":" & aUserKana & "様」"

'        Case pcstrMOVEKBN_05  ' *** 金額変更 ***  'DELETE 2011/09/27 M.RYU
        Case pcstrMOVEKBN_05, pcstrMOVEKBN_07      'INSERT 2011/09/27 M.RYU
            ' (旧契約番号「\旧契約月額」)
            aMoveMemo = anACPTNO & "「\" & Format(aMonthlyFee, "#,##0") & "」"

        Case pcstrMOVEKBN_06  ' *** 保障会社変更 ***
            ' (旧契約番号「保証区分略称」)
            aMoveMemo = anACPTNO & "「" & aGuarantyName & "」"
        '▼2008/03/25 add tajima
        Case Else
            aMoveMemo = anACPTNO
        '▲2008/03/25 add tajima
    End Select
                                  
    '②共通文言(から移動理由名)で〆
    aMoveMemo = aMoveMemo & "から" & aMoveTypeName
    CreateNewMoveMemo = True
    Exit Function
                                  
Exception:
    CreateNewMoveMemo = False
    Call Err.Raise(Err.Number, "CreateNewMoveMemo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :旧契約の「メンテ備考」の文言作成
'        MODULE_ID        :CreateOldMoveMemo
'        Parameter        :
'        Parameter        :第1引数(Integer)      = 移動区分
'                         :第2引数(String)       = 移動区分名称
'                         :第3引数(long)         = ヤードコード
'                         :第4引数(long)         = コンテナ番号
'                         :第5引数(long)         = 顧客コード
'                         :第6引数(String)       = 顧客名カナ
'                         :第7引数(Double)       = 月額料金
'                         :第8引数(String)       = 保証区分略称
'                         :第9引数(String)       = 受注契約番号
'                         :第10引数(String)      = 契約開始日
'                         :第11引数(ByRef String) = 生成された移動備考文言
'        Return           :True = 正常終了  False = 異常終了
'        Note             :旧契約に書込む、新契約の概要を作る
'        CREATE_DATE      :2006/02/16
''==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function CreateOldMoveMemo(ByVal aMoveType As Integer, _
                                  ByVal aMoveTypeName As String, _
                                  ByVal aYardCode As Long, _
                                  ByVal aContainerNo As Long, _
                                  ByVal aUserCode As Long, _
                                  ByVal aUserKana As String, _
                                  ByVal aMonthlyFee As Double, _
                                  ByVal aGuarantyName As String, _
                                  ByVal anACPTNO As String, _
                                  ByVal aStartDate As Date, _
                                  ByRef aMoveMemo As String _
                                  ) As Boolean
    On Error GoTo Exception
    aMoveMemo = ""
    
    '①それぞれの文言組み立て
    Select Case aMoveType
        ' *** 自社理由移動 OR 顧客理由移動 OR 未使用移動
        Case pcstrMOVEKBN_01, pcstrMOVEKBN_02, pcstrMOVEKBN_03
            ' (新契約ヤードコード-新契約コンテナ番号 )
            aMoveMemo = Format(aYardCode, "000000") & "-" & Format(aContainerNo, "000000")
            
        Case pcstrMOVEKBN_04  ' *** 名義変更 ***
            ' (新契約番号「新顧客コード:新顧客カナ名様」)
            aMoveMemo = anACPTNO & "「" & Format(aUserCode, "000000") & ":" & aUserKana & "様」"

'        Case pcstrMOVEKBN_05  ' *** 金額変更 ***  'DELETE 2011/09/27 M.RYU
        Case pcstrMOVEKBN_05, pcstrMOVEKBN_07      'INSERT 2011/09/27 M.RYU
            ' (新契約番号「\新契約月額」)
            aMoveMemo = anACPTNO & "「\" & Format(aMonthlyFee, "#,##0") & "」"

        Case pcstrMOVEKBN_06  ' *** 保障会社変更 ***
            ' (新契約番号「新保証区分略称 」)
            aMoveMemo = anACPTNO & "「" & aGuarantyName & "」"
        '▼2008/03/25 add tajima
        Case Else
            aMoveMemo = anACPTNO
        '▲2008/03/25 add tajima
    End Select
                                  
    '②共通文言(に移動区分名称 契約開始年月)で〆
    aMoveMemo = aMoveMemo & "に" & aMoveTypeName & Space(1) & Format(aStartDate, "YY/MM")
    CreateOldMoveMemo = True
    Exit Function
                                  
Exception:
    CreateOldMoveMemo = False
    Call Err.Raise(Err.Number, "CreateOldMoveMemo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :既存顧客検索ダイアログ表示
'        MODULE_ID        :ControlFVS560
'        Parameter        :第1引数(String)      = フォーム名
'                         :第2引数(String)      = コントロール名
'                         :第3引数(String)      = 顧客名カナ
'                         :第4引数(String)      = 顧客連絡先
'        CREATE_DATE      :2006/02/18
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub ControlFVS560(ByVal aFormName As String, ByVal aParentUserCdName As String, _
                         ByVal anUserKana As String, ByVal anUserRenraku As String)

    pstrParentFormName = aFormName        ' フォーム名
    pstrParentUserCd = aParentUserCdName  ' コントロール名

    ' 既存顧客検索画面表示
    doCmd.OpenForm "FVS560", acNormal, , , , acDialog, Nz(anUserKana, "") & "," & Nz(anUserRenraku, "")
End Sub

'==============================================================================*
'
'        MODULE_NAME      :顧客コード反映
'        MODULE_ID        :UserReflection
'        Parameter        :第1引数(String)      = 顧客コード
'        CREATE_DATE      :2006/02/18
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub UserReflection(ByRef strUSER_CODE As String)

    On Error GoTo Exception

    If Nz(strUSER_CODE, "") <> "" Then
        ' 呼び出し元画面の対象となる項目に顧客コードをセットする
        With Forms(pstrParentFormName).Controls(pstrParentUserCd)
            .VALUE = strUSER_CODE
        End With
    End If

    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "UserReflection" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    On Error Resume Next
End Sub
'==============================================================================*
'
'        MODULE_NAME      :キャンペーン適用情報文言取得
'        MODULE_ID        :GetCampaignInfomationText
'        Parameter        :関数宣言参照
'        Retrun           :キャンペーン適用情報
'        CREATE_DATE      :2006/03/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetCampaignInfomationText(ByVal aサービス内容1 As String, _
                                          ByVal aサービス内容2 As String, _
                                          ByVal aサービス内容3 As String, _
                                          ByVal aサービス期間 As String, _
                                          ByVal aキャンペーン適用満了日 As String _
                                       ) As String
  
  Dim returnText As String
  Dim tabText As String
  
  On Error GoTo Exception
  
  tabText = "・"
  returnText = "本契約は" & vbCrLf
  
  If aサービス内容1 <> "" Then
      returnText = returnText & tabText & aサービス内容1 & vbCrLf
  End If
  If aサービス内容2 <> "" Then
      returnText = returnText & tabText & aサービス内容2 & vbCrLf
  End If
  If aサービス内容3 <> "" Then
      returnText = returnText & tabText & aサービス内容3 & vbCrLf
  End If
  If aサービス期間 <> "" Then
      returnText = returnText & tabText & aサービス期間 & vbCrLf
  End If
  
  returnText = returnText & vbCrLf & "によるキャンペーンを適用しています。" & vbCrLf
  
  If Nz(aキャンペーン適用満了日, "") <> "" Then
    returnText = returnText & "よって【" & aキャンペーン適用満了日
    returnText = returnText & "】まで解約制限があります。"
  Else
    returnText = returnText & "解約制限は特にありません。"
  End If
 
  GetCampaignInfomationText = returnText
  Exit Function
  
Exception:
  Call Err.Raise(Err.Number, "GetCampaignInfomationText" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :キャンペーン適用警告文言取得
'        MODULE_ID        :GetCampaignWarningText
'        Parameter        :関数宣言参照
'        Retrun           :キャンペーン適用情報
'        CREATE_DATE      :2006/03/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetCampaignWarningText(ByVal aサービス内容1 As String, _
                                       ByVal aサービス内容2 As String, _
                                       ByVal aサービス内容3 As String, _
                                       ByVal aサービス期間 As String, _
                                       ByVal aキャンペーン適用満了日 As String _
                                       ) As String
  Dim returnText As String
  
  On Error GoTo Exception
  
  returnText = GetCampaignInfomationText(aサービス内容1, _
                                         aサービス内容2, _
                                         aサービス内容3, _
                                         aサービス期間, _
                                         aキャンペーン適用満了日)
  
  returnText = returnText & vbCrLf & "それでも解約を受付ますか？"
 
  GetCampaignWarningText = returnText
  Exit Function
  
Exception:
  Call Err.Raise(Err.Number, "GetCampaignWarningText" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'*********************** これ以降は 0.4 2006/04/25 Add *************************
'==============================================================================*
'
'        MODULE_NAME      :内税を求める
'        MODULE_ID        :GetIncludeTax
'        Parameter        :第1引数(Double) = 税込金額
'                         :第2引数(String) = 請求年月日(yyyymm形式であること)
'        Return           :計算した税金
'        CREATE_DATE      :2006/04/26
'        Note             :請求する年月日に応じた内税額を返す
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetIncludeTax(ByVal a税込金額 As Double, _
                              ByVal a請求年月 As String) As Double
'                                                                               'DELETE START 2019/08/02 K.ISHIZAKA
'  GetIncludeTax = 0
'
'  On Error GoTo Exception
'
'  ' 税計算するための情報取得可否のチェック
'  If IsNumeric(pvalBeforTaxRate) = False Or IsNumeric(pvalAfterTaxRate) = False Or _
'     Nz(pstrTaxJudgDate, "") = "" Or Nz(pstrRoundType, "") = "" Then
'
'    ' 税計算情報が未取得ならばKASE_DBから情報を取得する
'    If GetTaxMethod() = False Then
'        Call MSZZ024_M10("GetTaxMethod", "消費税計算するための情報が取得できません。")
'    End If
'
'  End If
'
'  Dim dRate As Double
'  Dim dTax As Double
'
'  ' 使用税率の判断
'
'  If a請求年月 < pstrTaxJudgDate Then
'    dRate = pvalBeforTaxRate  '変更前消費税率を使用
'  Else
'    dRate = pvalAfterTaxRate  '変更後消費税率を使用
'  End If
'
'  dTax = a税込金額 - (a税込金額 / (1 + dRate))
'
'  Select Case pstrRoundType
'
'    Case "0"  '四捨五入
'        GetIncludeTax = Fix(dTax + Sgn(dTax) * 0.5)
'
'    Case "1"  '切捨て
'        GetIncludeTax = Fix(dTax)
'
'    Case "2"  '切上げ
'        GetIncludeTax = Fix(dTax + Sgn(dTax) * 0.9)
'
'  End Select
'                                                                               'DELETE END   2019/08/02 K.ISHIZAKA
'                                                                               'INSERT START 2019/08/02 K.ISHIZAKA
    Dim lngPrice            As Long
    Dim lngTax              As Long
    On Error GoTo Exception
    
    Call MSZZ004_M10(a税込金額, a請求年月, "2", lngPrice, lngTax)
    GetIncludeTax = lngTax
'                                                                               'INSERT END   2019/08/02 K.ISHIZAKA
  Exit Function
                             
Exception:
  Call Err.Raise(Err.Number, "GetIncludeTax" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
  
'                                                                               'DELETE START 2019/08/02 K.ISHIZAKA
''==============================================================================*
''
''        MODULE_NAME      :税計算方法の取得
''        MODULE_ID        :GetTaxMethod
''        Parameter        :Nothing
''        Return           :True = 正常終了  False = 異常終了(情報取得失敗)
''        CREATE_DATE      :2006/04/26
''        Note             :KASE_DBのから税計算に関する情報を取得し、グローバルに納める
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function GetTaxMethod() As Boolean
'
'  On Error GoTo Exception
'
'    Dim strSQL     As String
'    Dim dbObject   As Database  '加瀬DBオブジェクト
'    Dim objRs      As Recordset 'レコードセット
'
'    Dim strDataSource As String
'
'    ' 加瀬DB接続
'    strDataSource = GetDataSource("ODBC_DATA_SOURCE_NAME")
'
'    If strDataSource = "" Then
'        ' テーブル[SETU_TABL]の設定不正
'      GetTaxMethod = False
'      GoTo Finally
'    End If
'
'    Set dbObject = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
'
'    ' 設定テーブルから消費税関連の情報を取得する
'    strSQL = " SELECT CONFT_SYYMD,CONFT_SYOLR,CONFT_SYNWR,CONFT_SYKBI"
'    strSQL = strSQL & " FROM CONF_TABL "
'
'    Set objRs = dbObject.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
'
'    If objRs.EOF = False Then
'      pstrTaxJudgDate = Nz(objRs.Fields("CONFT_SYYMD"), "")
'      pstrRoundType = Nz(objRs.Fields("CONFT_SYKBI"), "")
'      pvalBeforTaxRate = objRs.Fields("CONFT_SYOLR")
'      pvalAfterTaxRate = objRs.Fields("CONFT_SYNWR")
'    Else
'      GetTaxMethod = False
'      GoTo Finally
'    End If
'
'    ' 取得結果の評価
'    If pstrTaxJudgDate = "" Or pstrRoundType = "" Or _
'      IsNumeric(pvalBeforTaxRate) = False Or IsNumeric(pvalAfterTaxRate) = False Then
'
'      GetTaxMethod = False
'      GoTo Finally
'    End If
'
'    GetTaxMethod = True
'
'Finally:
'  If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
'  If Not dbObject Is Nothing Then dbObject.Close: Set dbObject = Nothing
'  Exit Function
'
'Exception:
'  GetTaxMethod = False
'  If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
'  If Not dbObject Is Nothing Then dbObject.Close: Set dbObject = Nothing
'  Call Err.Raise(Err.Number, "GetTaxMethod" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'
'End Function
'
''==============================================================================*
''
''        MODULE_NAME      :ODBC接続文字列取得
''        MODULE_ID        :GetDataSource
''        Parameter        :anODBCNAME ODBC名
''        CREATE_DATE      :2005/08/10
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function GetDataSource(ByVal anODBCNAME As String) As String
'
'    Dim strName As String
'
'    On Error GoTo Exception
'
'    GetDataSource = ""
'
'    strName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = '" & anODBCNAME & "'"))
'
'    GetDataSource = strName
'
'    Exit Function
'
'Exception:
'  Call Err.Raise(Err.Number, "GetDataSource" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'
'End Function
'                                                                               'DELETE END   2019/08/02 K.ISHIZAKA

'==============================================================================*
'
'        MODULE_NAME      :INTR_TRANの存在チェック
'        MODULE_ID        :IsExistenceIntrTran
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 予約番号
'                         :第3引数(Long) = ヤードコード
'                         :第4引数(Long) = コンテナ番号
'        Return           :TRUE...あり、FALSE...なし
'        CREATE_DATE      :2006/05/01
'        Note             :指定したコンテナがINTR_TRANにあるかチェックをする
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function IsExistenceIntrTran(dbSQLServer As DAO.Database, _
                                    a予約番号 As String, _
                                    aヤードコード As Long, _
                                    aコンテナ番号 As Long _
                                    ) As Boolean

    Dim strSQL As String
    Dim rsObject As Recordset 'レコードセット
    Dim intCount As Integer
    
On Error GoTo Exception

    ' 紹介トランにあるかチェック
    strSQL = strSQL & "SELECT COUNT(*) CNT "
    strSQL = strSQL & "  FROM INTR_TRAN "
    strSQL = strSQL & "WHERE INTRT_UKNO = '" & a予約番号 & "'"
    strSQL = strSQL & "  AND INTRT_YCODE = " & aヤードコード & Chr(13)
    strSQL = strSQL & "  AND INTRT_NO    = " & aコンテナ番号 & Chr(13)
    strSQL = strSQL & "  AND INTRT_INTROKBN IN('1','2')" 'ご紹介区分、01:取置、02:受付 参照
    
    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
    intCount = 0
    intCount = Nz(rsObject.Fields("CNT"), 0)
    rsObject.Close
    Set rsObject = Nothing
  
    If intCount > 0 Then
        IsExistenceIntrTran = True
    Else
        IsExistenceIntrTran = False
    End If
    
    Exit Function
    
Exception:
    IsExistenceIntrTran = False
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "IsExistenceIntrTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :紹介待ちコンテナのクリア
'        MODULE_ID        :GetIncludeTax
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 紹介待ちにした予約番号
'        CREATE_DATE      :2006/05/01
'        Note             :紹介待ちをクリアする
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub ClearWaitCntaMast(dbSQLServer As DAO.Database, _
                             a紹介待ちにした予約番号 As String)

    Dim strSQL As String
    Dim rsObject As Recordset
    
On Error GoTo Exception

    strSQL = strSQL & "SELECT CNTA_CODE, CNTA_NO, CNTA_UPDATE, CNTA_USE, CNTA_UKNO"
    strSQL = strSQL & "  FROM CNTA_MAST "
    strSQL = strSQL & "WHERE CNTA_UKNO = '" & a紹介待ちにした予約番号 & "'"

    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    With rsObject
      While Not .EOF
          .Edit
          .Fields("CNTA_USE") = pcintCNTA_USE_OK
          .Fields("CNTA_UPDATE") = DATE
          .Fields("CNTA_UKNO") = Null
          .UPDATE
          .MoveNext
      Wend
      .Close
    End With

    Set rsObject = Nothing
    Exit Sub
    
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "ClearWaitCntaMast" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'*********************** これ以降は 0.5 2006/05/24 Add *************************
'==============================================================================*
'
'        MODULE_NAME      :セキュリティカードの初期化
'        MODULE_ID        :InitializSCards
'        Parameter        :第1引数(Recordset) = 対象レコード
'        Return           :
'        CREATE_DATE      :2006/05/01
'        Note             :セキュリティカードを初期状態（未貸出）に戻す
'                         :初期化手順を共通とするので初期化は本サブを使用すること
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub InitializSCards(ByRef aRecordSet As Recordset)
      
      ' 使用用途が基本ならば使用用途、割当部屋番、は変わらない
      ' But 追加の場合は予備に戻し、且つ初期設定の割当に戻す
      If aRecordSet.Fields("SCRDM_USAGE") = P_カード用途_追加値 Then
          ' 使用用途が追加ならば予備とする
          aRecordSet.Fields("SCRDM_USAGE") = P_カード用途_予備値
          ' 部屋番は返却先部屋番号が入る
          aRecordSet.Fields("SCRDM_ROOM_NO") = aRecordSet.Fields("SCRDM_RETURN_RNO")
      End If
      aRecordSet.Fields("SCRDM_OTHER_NAME") = Null
      aRecordSet.Fields("SCRDM_ACPTNO") = Null
      aRecordSet.Fields("SCRDM_LENDING_DAY") = Null
      aRecordSet.Fields("SCRDM_RETURN_DAY") = Null
      '更新情報はそれぞれ上位Subで設定してほしい

End Sub

'==============================================================================*
'
'        MODULE_NAME      :契約で貸出たセキュリティカードの取消
'        MODULE_ID        :CancelContractSCards
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 契約番号
'                         :第3引数(String) = 更新ユーザID
'                         :第4引数(String) = 更新プログラムID
'                         :第5引数(String) = 予約番号※予約入力が呼ぶ場合のみ設定
'        Return           :
'        CREATE_DATE      :2006/05/01
'        Note             :取消というか元に戻す
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub CancelContractSCards(dbSQLServer As DAO.Database, _
                                        ByVal a契約番号 As String, _
                                        Optional aユーザID As String = "", _
                                        Optional aプログラムID As String = "", _
                                        Optional a予約番号 As String = "" _
                                       )
    Dim strSQL As String
    Dim rsObject As Recordset
    Dim strDate As String
    Dim strTime As String
  
  On Error GoTo Exception
    
    ' 更新情報を取得しておく
    strDate = Format$(DATE, "yyyymmdd")
    strTime = Format$(time, "hhmmss")
    If aユーザID = "" Then
        aユーザID = LsGetUserName()
    End If
    If aプログラムID = "" Then
        aプログラムID = GetProgramName()
    End If
    
    ' 指定ヤード＆部屋のカードを集めるSQL
    strSQL = "SELECT * FROM SCRD_MAST "
    strSQL = strSQL & " WHERE SCRDM_ACPTNO = "
    If a契約番号 <> "" Then
        strSQL = strSQL & "'" & a契約番号 & "'"
    ElseIf a予約番号 <> "" Then
        strSQL = strSQL & "(SELECT RCPT_CARG_ACPTNO FROM RCPT_TRAN WHERE RCPT_NO = '" & a予約番号 & "' )"
    Else
        Call MSZZ024_M10("CancelContractSCards", "パラメータエラー・予約番号も契約番号指定されていない")
    End If
    strSQL = strSQL & "   AND SCRDM_STOP_DAY IS NULL"  ' 停止中でないもの

    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    ' 取得したカードを未使用にしていく
    While Not rsObject.EOF
        rsObject.Edit
        Call InitializSCards(rsObject) ' 未使用は手順共通化
        rsObject.Fields("SCRDM_UPDAD") = strDate
        rsObject.Fields("SCRDM_UPDAJ") = strTime
        rsObject.Fields("SCRDM_UPDPB") = aユーザID
        rsObject.Fields("SCRDM_UPDUB") = aプログラムID
        rsObject.UPDATE
        rsObject.MoveNext
    Wend
    rsObject.Close

    Set rsObject = Nothing
    Exit Sub
    
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "CancelContractSCards" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :契約使用中セキュリティカードの取得
'        MODULE_ID        :GetContractSCards
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 契約番号
'                         :第3引数(Integer)= 基本枚数(out)
'                         :第4引数(Integer)= 追加枚数(out)
'                         :第5引数(String) = 契約番に割当ているカード達(out)
'                         :                  ※例）AAAA;BBBB;CCCC そのままRowSourceに入れて頂けます
'        CREATE_DATE      :2006/05/01
'        Note             :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub GetContractSCards(dbSQLServer As DAO.Database, _
                                        ByVal a契約番号 As String, _
                                        ByRef a基本枚数 As Integer, _
                                        ByRef a追加枚数 As Integer, _
                                        ByRef a割当カード達 As String _
                                       )
    Dim strSQL As String
    Dim rsObject As Object
    Dim isException As Boolean
    
  On Error GoTo Exception
    isException = False
    a基本枚数 = 0
    a追加枚数 = 0
    a割当カード達 = ""

    ' 指定した契約番号で使っている（停止中は除くよ）カードのみ集める
    strSQL = strSQL & "SELECT SCRDM_CARD_NO, SCRDM_USAGE"
    strSQL = strSQL & "  FROM SCRD_MAST"
    strSQL = strSQL & " WHERE SCRDM_ACPTNO = '" & a契約番号 & "'"
    strSQL = strSQL & "   AND SCRDM_STOP_DAY IS NULL"
    strSQL = strSQL & " ORDER BY SCRDM_CARD_NO"

    'INSERT 2020/04/01 N.IMAI Start
    Dim objDb   As Object
    Dim objRst  As Object
    Set objDb = ADODB_Connection(Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1")))
    Set objRst = ADODB_Recordset(strSQL, objDb)
    
    ' 取得したカード番号を詰めていく
    While Not objRst.EOF
        If a割当カード達 <> "" Then a割当カード達 = a割当カード達 & ";"
        a割当カード達 = a割当カード達 & Nz(objRst.Fields("SCRDM_CARD_NO"), "")
        If objRst.Fields("SCRDM_USAGE") = P_カード用途_基本値 Then
            a基本枚数 = a基本枚数 + 1
        Else
            a追加枚数 = a追加枚数 + 1
        End If
        objRst.MoveNext
    Wend
    'INSERT 2020/04/01 N.IMAI End

    'DELETE 2020/04/01 N.IMAI Start
'    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
'
'    ' 取得したカード番号を詰めていく
'    While Not rsObject.EOF
'        If a割当カード達 <> "" Then a割当カード達 = a割当カード達 & ";"
'        a割当カード達 = a割当カード達 & Nz(rsObject.Fields("SCRDM_CARD_NO"), "")
'        If rsObject.Fields("SCRDM_USAGE") = P_カード用途_基本値 Then
'            a基本枚数 = a基本枚数 + 1
'        Else
'            a追加枚数 = a追加枚数 + 1
'        End If
'        rsObject.MoveNext
'    Wend
    'DELETE 2020/04/01 N.IMAI End
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    'DELETE 2020/04/01 N.IMAI Start
'    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
'
'    If isException = True Then
'        Call Err.Raise(Err.Number, "GetContractSCards" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
    'DELETE 2020/04/01 N.IMAI End
    
    'INSERT 2020/04/01 N.IMAI Start
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    'INSERT 2020/04/01 N.IMAI End
    
    If isException = True Then
        Call Err.Raise(Err.Number, "GetContractSCards" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :貸出可能部屋割当セキュリティカード情報の取得
'        MODULE_ID        :GetRoomAllotSCards
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(Long)   = ヤードコード
'                         :第3引数(Long)   = 部屋番号
'                         :第4引数(Integer)= 基本枚数(out)
'                         :第5引数(Integer)= 追加枚数(out)
'                         :第6引数(String) = ヤードと部屋に割当ているカード達(out)
'                         :                  ※例）AAAA;BBBB;CCCC そのままRowSourceに入れて頂けます
'        CREATE_DATE      :2006/05/01
'        Note             :貸出可能なカード情報を取得する※受付入力初読込時使用
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub GetRoomAllotSCards(dbSQLServer As DAO.Database, _
                                        ByVal aヤードコード As Long, _
                                        ByVal a部屋番号 As Long, _
                                        ByRef a基本枚数 As Integer, _
                                        ByRef a追加枚数 As Integer, _
                                        ByRef a割当カード達 As String _
                                       )
    Dim strSQL As String
    Dim rsObject As Recordset
    Dim isException As Boolean
    
  On Error GoTo Exception
    isException = False
    a基本枚数 = 0
    a追加枚数 = 0
    a割当カード達 = ""

    ' 次に貸出可能なカード情報のみを集める
    ' 契約番号が入っているものは貸出不可ですよ
    strSQL = strSQL & "SELECT SCRDM_CARD_NO, SCRDM_USAGE"
    strSQL = strSQL & "  FROM SCRD_MAST"
    strSQL = strSQL & " WHERE SCRDM_YCODE = " & aヤードコード
    strSQL = strSQL & "   AND SCRDM_ROOM_NO = " & a部屋番号
    strSQL = strSQL & "   AND SCRDM_ACPTNO IS NULL"       ' ←契約番号無しが対象 ↓基本と追加を対象に
    strSQL = strSQL & "   AND SCRDM_USAGE IN('" & P_カード用途_基本値 & "','" & P_カード用途_追加値 & "')"
    strSQL = strSQL & " ORDER BY SCRDM_CARD_NO"

    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

    ' 取得したカード番号を詰めていく
    While Not rsObject.EOF
        If a割当カード達 <> "" Then a割当カード達 = a割当カード達 & ";"
        a割当カード達 = a割当カード達 & Nz(rsObject.Fields("SCRDM_CARD_NO"), "")
        If rsObject.Fields("SCRDM_USAGE") = "01" Then
            a基本枚数 = a基本枚数 + 1
        Else
            a追加枚数 = a追加枚数 + 1
        End If
        rsObject.MoveNext
    Wend
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "GetRoomAllotSCards" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :セキュリティカードの貸出確定
'        MODULE_ID        :SetContractSCards
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = ヤードコード
'                         :第3引数(String) = 部屋番号
'                         :第4引数(String) = 契約番号
'                         :第5引数(String) = 更新ユーザID
'                         :第6引数(String) = 更新プログラムID
'        Return           :
'        CREATE_DATE      :2006/06/01
'        Note             :受付入力で初めて登録するときに呼ぶ
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetContractSCards(dbSQLServer As DAO.Database, _
                                        ByVal aヤードコード As Long, _
                                        ByVal a部屋番号 As Long, _
                                        ByVal a契約番号 As String, _
                                        Optional aユーザID As String = "", _
                                        Optional aプログラムID As String = "" _
                                       )
    Dim strSQL As String
    Dim rsObject As Recordset
    Dim strDate As String
    Dim strTime As String

  On Error GoTo Exception
    
    ' 更新情報を取得しておく
    strDate = Format$(DATE, "yyyymmdd")
    strTime = Format$(time, "hhmmss")
    If aユーザID = "" Then
        aユーザID = LsGetUserName()
    End If
    If aプログラムID = "" Then
        aプログラムID = GetProgramName()
    End If

    ' いまこの部屋に割り当てている基本貸出のカードを集める
    ' 契約番号が入っているものは貸出不可ですよ←停止中のモノだ
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "  FROM SCRD_MAST"
    strSQL = strSQL & " WHERE SCRDM_YCODE = " & aヤードコード
    strSQL = strSQL & "  AND SCRDM_ROOM_NO = " & a部屋番号
    strSQL = strSQL & "  AND SCRDM_USAGE = '" & P_カード用途_基本値 & "'"  ' 基本貸しが対象"
    strSQL = strSQL & "  AND SCRDM_ACPTNO IS NULL"       ' 契約番号無しが対象

    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    ' 取得したカードを基本貸出にしていく
    With rsObject
      While Not .EOF
          .Edit
          .Fields("SCRDM_OTHER_NAME") = Null
          .Fields("SCRDM_ACPTNO") = a契約番号
          .Fields("SCRDM_LENDING_DAY") = DATE
          .Fields("SCRDM_RETURN_DAY") = Null
          .Fields("SCRDM_STOP_DAY") = Null
          .Fields("SCRDM_UPDAD") = strDate
          .Fields("SCRDM_UPDAJ") = strTime
          .Fields("SCRDM_UPDPB") = aユーザID
          .Fields("SCRDM_UPDUB") = aプログラムID
          .UPDATE
          .MoveNext
      Wend
      .Close
    End With

    Set rsObject = Nothing
    Exit Sub
    
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "SetContractSCards" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :セキュリティカードの貸出先付け替え
'        MODULE_ID        :SetChangeSCards
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String) = 旧契約番号
'                         :第3引数(String) = 新契約番号
'                         :第4引数(String) = 更新ユーザID
'                         :第5引数(String) = 更新プログラムID
'        Return           :
'        CREATE_DATE      :2006/06/01
'        Note             :受付入力(契約変更)で初めて登録するときに呼ぶ
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetChangeSCards(dbSQLServer As DAO.Database, _
                                        ByVal a旧契約番号 As String, _
                                        ByVal a新契約番号 As String, _
                                        Optional aユーザID As String = "", _
                                        Optional aプログラムID As String = "" _
                                       )
    Dim strSQL As String
    Dim rsObject As Recordset
    Dim strDate As String
    Dim strTime As String

  On Error GoTo Exception
    
    ' 更新情報を取得しておく
    strDate = Format$(DATE, "yyyymmdd")
    strTime = Format$(time, "hhmmss")
    If aユーザID = "" Then
        aユーザID = LsGetUserName()
    End If
    If aプログラムID = "" Then
        aプログラムID = GetProgramName()
    End If

    ' 指定した契約で貸出中のカードを集める
    ' 契約番号が入っているものは貸出不可ですよ←停止中のモノだ
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "  FROM SCRD_MAST"
    strSQL = strSQL & " WHERE SCRDM_ACPTNO = '" & a旧契約番号 & "'"
    strSQL = strSQL & "  AND SCRDM_STOP_DAY IS NULL"  ' 停止カードは引き継がない

    Set rsObject = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)

    ' 取得したカードの契約番号を書き換える
    ' ☆貸出日は変えないで契約番号のみ書き換える。
    With rsObject
      While Not .EOF
          .Edit
          .Fields("SCRDM_ACPTNO") = a新契約番号
          .Fields("SCRDM_UPDAD") = strDate
          .Fields("SCRDM_UPDAJ") = strTime
          .Fields("SCRDM_UPDPB") = aユーザID
          .Fields("SCRDM_UPDUB") = aプログラムID
          .UPDATE
          .MoveNext
      Wend
      .Close
    End With

    Set rsObject = Nothing
    Exit Sub
    
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    Call Err.Raise(Err.Number, "SetChangeSCards" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'****************************  ended or program ********************************

' ***** Start 2017/07/31 add by ysuzuki

'==============================================================================*
'   MODULE_NAME     : 既存契約（顧客コードで他の契約）があるか確認
'   MODULE_ID       : CheckOtherAgreement
'   Parameter       : 第1引数(dao.DataBase)  : dbSqlServer = SqlServerにDAO接続したDataBase
'                     第2引数(String)        : a_UCODE     = 顧客コード
'                     第3引数(String)        : a_ACPTNO    = 受注契約番号
'   Return          : ( エラー=-1、他の契約なし=0、他の契約あり=1 )
'   CREATE_DATE     : 2017/07/28
'   Note            : 既存契約（顧客コードで他の契約）があるか確認
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function CheckOtherAgreement(dbSQLServer As DAO.Database, ByVal lngUcode As Long, ByVal strACPTNO As String) As Long
Private Function CheckOtherAgreement(dbSQLServer As Object, ByVal lngUcode As Long, ByVal strACPTNO As String) As Long
    
    Dim strSQL As String
    'Dim rsObject As Recordset                                                  'DELETE 2021/03/31 N.IMAI
    Dim rsObject As Object                                                      'INSERT 2021/03/31 N.IMAI

    CheckOtherAgreement = -1
    
    On Error GoTo Exception
    
    ' ﾊﾟﾗﾒｰﾀﾁｪｯｸ
    If dbSQLServer Is Nothing Then Exit Function
    
    If "" = strACPTNO Or IsNull(lngUcode) Then
        CheckOtherAgreement = 0
        GoTo Exception
    End If
    
    ' SQL編集
    strSQL = strSQL & "SELECT * " & vbCrLf
    strSQL = strSQL & "  FROM CARG_FILE" & vbCrLf
    strSQL = strSQL & " WHERE CARG_UCODE  = " & lngUcode & vbCrLf
    strSQL = strSQL & "   AND CARG_ACPTNO != '" & strACPTNO & "'" & vbCrLf
    strSQL = strSQL & "   AND ( ISNULL(CARG_CYDATE,'9999/12/31') > GETDATE() )" & vbCrLf
    strSQL = strSQL & "   AND CARG_AGRE  IN ( 1, 2, 3 )" & vbCrLf
    
    'Set rsObject = dbSQLServer.OpenRecordset(strSql, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly) 'DELETE 2021/03/31 N.IMAI
    Set rsObject = ADODB_Recordset(strSQL, dbSQLServer)                                             'INSERT 2021/03/31 N.IMAI
    
    ' 該当データ無
    If rsObject.EOF Then
        CheckOtherAgreement = 0
        GoTo Exception
        
    End If
    
    CheckOtherAgreement = 1
    
Exception:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    If -1 = CheckOtherAgreement Then
        Call Err.Raise(Err.Number, "CheckOtherAgreement" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        
    End If
    
End Function

'==============================================================================*
'   MODULE_NAME : コンテナ契約ファイル、請求ファイルへのデータ登録の有無判定
'   MODULE_ID   : CheckRegi_target
'   Parameter   : 第1引数(dao.DataBase)  : dbSqlServer = SqlServerにDAO接続したDataBase
'   Return      : Long = ( 異常終了=-1、登録しない=0、登録する=1 )
'
'   CREATE_DATE : 2017/07/28
'   Note        :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function CheckRegiTarget(dbSQLServer As DAO.Database
Public Function CheckRegiTarget(dbSQLServer As Object _
                                 , ByVal strHassei_Kbn As String _
                                 , ByVal strShoki_Seikyu As String _
                                 , ByVal lngUcode As Long _
                                 , ByVal strACPTNO As String) As Long

    ' 変数宣言
    Dim lngRTN  As Long
    
    On Error GoTo Exception
    
    'コンテナ契約ファイル、請求ファイルへのデータ登録対象外とする。
    CheckRegiTarget = -1
    
    '同じ顧客コードで他の契約が在るか確認( エラー=-1、他の契約なし=0、他の契約あり=1 )
    lngRTN = CheckOtherAgreement(dbSQLServer, lngUcode, strACPTNO)
    
    '発生区分="0"(電話)、初期費用請求方法="1"(振込)、同じ顧客コードで他の契約が在る場合
    
    strHassei_Kbn = Right(String(2, "0") & Trim(strHassei_Kbn), 2)
    strShoki_Seikyu = Right(String(2, "0") & Trim(strShoki_Seikyu), 2)
    
    If ("00" = strHassei_Kbn And "01" = strShoki_Seikyu And 1 = lngRTN) Then
        'コンテナ契約ファイル、請求ファイルへのデータ登録対象とする。
        CheckRegiTarget = 1
        
    Else
        'コンテナ契約ファイル、請求ファイルへのデータ登録対象外とする。
        CheckRegiTarget = 0
        
    End If
    
    Exit Function
    
Exception:
    'エラー処理
    If -1 = CheckRegiTarget Then
        Call Err.Raise(Err.Number, "CheckRegiTarget" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'   MODULE_NAME     : 請求ファイルのデータを削除
'   MODULE_ID       : DELETE_REQU_FILE
'   Parameter       : 第1引数(dao.DataBase)  : dbSqlServer = SqlServerにDAO接続したDataBase
'                     第2引数(String)       : strACPTNO = 画面．受注契約番号
'   Return          : Long = ( 異常終了=False、正常終了=True )
'
'   CREATE_DATE     : 2017/07/28
'   Note            :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function DeleteRequFile(dbSQLServer As DAO.Database, ByVal strACPTNO As String) As Boolean

    Dim rsREQU_FILE As Recordset
    Dim strSQL As String
    Dim i As Integer '2020/03/25
    
    DeleteRequFile = False
    
    On Error GoTo Exception
    
    ' パラメータチェック
    If "" = strACPTNO Then GoTo End_Function
    
    ' 対象データが請求ファイルにいるか確認
    strSQL = "SELECT * FROM REQU_FILE WHERE REQU_ACPTNO = '" & strACPTNO & "'"
    Set rsREQU_FILE = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
'DELETE 2023/03/02 N.IMAI Start
    ' 請求ファイルに削除対象データがいない場合は処理を終了
'    If rsREQU_FILE.EOF Then GoTo End_Function
'    For i = 1 To rsREQU_FILE.RecordCount '2020/03/25 add 存在した件数分の削除
'
'        ' 請求ファイルの削除
'            With rsREQU_FILE
'            ' 削除
'            .Edit
'            .Delete
'        End With
'
'    Next '2020/03/25 add 存在した件数分の削除
'DELETE 2023/03/02 N.IMAI End
   
    'INSERT 2023/03/02 N.IMAI Start
    If rsREQU_FILE.EOF = False Then
        dbSQLServer.Execute ("DELETE FROM REQU_FILE WHERE REQU_ACPTNO = '" & strACPTNO & "'")
    End If
    'INSERT 2023/03/02 N.IMAI End
    rsREQU_FILE.Close
    Set rsREQU_FILE = Nothing
    
End_Function:
    DeleteRequFile = True
    
Exception:
    If Err <> 0 Then
        If Not rsREQU_FILE Is Nothing Then rsREQU_FILE.Close: Set rsREQU_FILE = Nothing     ' 請求ファイル
        If Not DeleteRequFile Then
            ' エラー処理
            Call Err.Raise(Err.Number, "DeleteRequFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        End If
    End If
    
End Function

'==============================================================================*
'   MODULE_NAME     : 請求ファイル新規作成２
'   MODULE_ID       : InsertRequFile2
'   Parameter       : 第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                     第2引数(Long)    = 請求No
'                     第3引数(String)  = 請求年月
'                     第4引数(String)  = 請求対象年月（売上計上）
'                     第5引数(Long)    = 顧客コード
'                     第6引数(String)  = ヤードコード
'                     第7引数(String)  = コンテナ番号
'                     第8引数(Double)  = 金額
'                     第9引数(Double)  = 消費税
'                     第10引数(Double) = 保証金
'                     第11引数(String) = 受注契約番号
'                     第12引数(ByRef String) = 異常終了時にエラーメッセージ格納
'
'   Return          : True = 正常終了  False = 異常終了
'   CREATE_DATE     : 2017/08/04
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function InsertRequFile2(dbSQLServer As DAO.Database, _
                                   ByVal aReqNo As Long, _
                                   ByVal aSeikyuDate As String, _
                                   ByVal aKeijyoDate As String, _
                                   ByVal anUserCode As Long, _
                                   ByVal aYardCode As String, _
                                   ByVal aContainerCode As String, _
                                   ByVal aKingaku As Double, _
                                   ByVal aSyohizei As Double, _
                                   ByVal aSecukingaku As Double, _
                                   ByVal anACPTNO As String, _
                                   ByRef aReturnMessage As String _
                                  ) As Boolean

    Dim strSQL    As String
    Dim objRs     As Recordset
    Dim strNKDATE As String
    Dim dblTOTAL  As Double
    Dim dbl金額   As Double

    On Error GoTo Exception

    '初期処理
    InsertRequFile2 = False
    aReturnMessage = ""

   ' 請求ファイルデータ抽出
    strSQL = "SELECT * FROM REQU_FILE WHERE 1=2"
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    With objRs
        .AddNew

        .Fields("REQU_REQNO") = aReqNo                                    ' 請求№
        .Fields("REQU_YYYYMM") = aSeikyuDate                              ' 請求年月
        .Fields("REQU_KJDATE") = aKeijyoDate                              ' 請求対象年月（売上計上）
        .Fields("REQU_SDATE") = aSeikyuDate                               ' 請求年月日

        ' 入金予定日
        strNKDATE = ""
        strNKDATE = Format(DateAdd("D", -1, CDate(aKeijyoDate)), "YYYY/MM/DD")
        .Fields("REQU_NKDATE") = strNKDATE

        .Fields("REQU_KCODE") = anUserCode                                ' 顧客コード
        .Fields("REQU_NKBN") = 0                                          ' 入金区分(0固定)
        .Fields("REQU_YCODE") = aYardCode                                 ' ﾔｰﾄﾞｺｰﾄﾞ
        .Fields("REQU_NO") = aContainerCode                               ' ｺﾝﾃﾅ番号
        .Fields("REQU_TTANKA") = 0                                        ' 坪単価(0固定)
        .Fields("REQU_TUBOSU") = 0                                        ' 坪数(0固定)

        '金額
        If aKingaku > pcsglMoneySumMaxValue Then
            dbl金額 = pcsglMoneySumMaxValue
        Else
            dbl金額 = aKingaku
        End If

        ' 消費税 2006/04/27 消費税は金額から内税で求める
        .Fields("REQU_SYOZEI") = GetIncludeTax(dbl金額, Format$(aKeijyoDate, "YYYYMM"))
        dbl金額 = dbl金額 - .Fields("REQU_SYOZEI")
        .Fields("REQU_KINGAK") = dbl金額
        
        .Fields("REQU_SECUKG") = aSecukingaku                             ' 保証金
        
        ' 合計(金額 + 消費税 + 保証金)
        dblTOTAL = dbl金額 + .Fields("REQU_SYOZEI") + aSecukingaku
        If dblTOTAL > pcsglMoneySumMaxValue Then
            .Fields("REQU_TOTAL") = pcsglMoneySumMaxValue
        Else
            .Fields("REQU_TOTAL") = dblTOTAL
        End If

        .Fields("REQU_TEKI") = Null                                       ' 摘要(Null固定)
        .Fields("REQU_FLG") = 1                                           ' 請求修正ﾌﾗｸﾞ(1固定)
        .Fields("REQU_UPDATE") = Format(DATE, "YYYY/MM/DD")               ' 更新日
        .Fields("REQU_ACPTNO") = anACPTNO                                 ' 受注契約番号

        .UPDATE
        
    End With
    
    InsertRequFile2 = True

Exception:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not InsertRequFile2 Then
        aReturnMessage = "InsertRequFile2(" & Err.Number & ")" & Err.Description
        Call Err.Raise(Err.Number, "InsertRequFile2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        
    End If
    
End Function


' ***** end 2017/07/31 add by ysuzuki

' ***** Start 2018/03/19 add by EGL
'==============================================================================*
'   MODULE_NAME     : 請求削除へInert
'   MODULE_ID       : InsertRequDele
'   Parameter       : 第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                     第2引数(String) = 受注契約番号
'                     第6引数(String) = 削除ユーザ
'                     第5引数(String) = 削除機能
'   Return          : true:正常終了（対象有無は不明だが)、false:失敗(例外発生)
'   CREATE_DATE     : 2018/03/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function InsertRequDele(dbSQLServer As DAO.Database, _
                                    ByVal anACPTNO As String, _
                                    ByVal a削除ユーザ As String, _
                                    ByVal a削除機能 As String _
                            ) As Boolean
    Dim strDate As String
    Dim strTime As String
    Dim strSQL As String
    
    InsertRequDele = False
    
    On Error GoTo Exception
    
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    
    ' パラメータチェック
    If "" = anACPTNO Then GoTo Exception
    
     ' インサート実行
    strSQL = "insert into REQU_DELE ( "
    strSQL = strSQL & " REQD_REQNO "
    strSQL = strSQL & ",REQD_YYYYMM"
    strSQL = strSQL & ",REQD_KJDATE"
    strSQL = strSQL & ",REQD_SDATE"
    strSQL = strSQL & ",REQD_NKDATE"
    strSQL = strSQL & ",REQD_KCODE"
    strSQL = strSQL & ",REQD_NKBN"
    strSQL = strSQL & ",REQD_YCODE"
    strSQL = strSQL & ",REQD_NO"
    strSQL = strSQL & ",REQD_TTANKA"
    strSQL = strSQL & ",REQD_TUBOSU"
    strSQL = strSQL & ",REQD_KINGAK"
    strSQL = strSQL & ",REQD_SYOZEI"
    strSQL = strSQL & ",REQD_SECUKG"
    strSQL = strSQL & ",REQD_TOTAL"
    strSQL = strSQL & ",REQD_TEKI"
    strSQL = strSQL & ",REQD_FLG"
    strSQL = strSQL & ",REQD_UPDATE"
    strSQL = strSQL & ",REQD_ACPTNO"
    strSQL = strSQL & ",REQD_DELED"
    strSQL = strSQL & ",REQD_DELEJ"
    strSQL = strSQL & ",REQD_DELPB"
    strSQL = strSQL & ",REQD_DELUB"
    strSQL = strSQL & " ) select "
    strSQL = strSQL & " REQU_REQNO "
    strSQL = strSQL & ",REQU_YYYYMM"
    strSQL = strSQL & ",REQU_KJDATE"
    strSQL = strSQL & ",REQU_SDATE"
    strSQL = strSQL & ",REQU_NKDATE"
    strSQL = strSQL & ",REQU_KCODE"
    strSQL = strSQL & ",REQU_NKBN"
    strSQL = strSQL & ",REQU_YCODE"
    strSQL = strSQL & ",REQU_NO"
    strSQL = strSQL & ",REQU_TTANKA"
    strSQL = strSQL & ",REQU_TUBOSU"
    strSQL = strSQL & ",REQU_KINGAK"
    strSQL = strSQL & ",REQU_SYOZEI"
    strSQL = strSQL & ",REQU_SECUKG"
    strSQL = strSQL & ",REQU_TOTAL"
    strSQL = strSQL & ",REQU_TEKI"
    strSQL = strSQL & ",REQU_FLG"
    strSQL = strSQL & ",REQU_UPDATE"
    strSQL = strSQL & ",REQU_ACPTNO"
    strSQL = strSQL & ",'" & strDate & "'"
    strSQL = strSQL & ",'" & strTime & "'"
    strSQL = strSQL & ",'" & a削除機能 & "'"
    strSQL = strSQL & ",'" & a削除ユーザ & "'"
    strSQL = strSQL & " from REQU_FILE where REQU_ACPTNO = '" & anACPTNO & "'"
    
    '請求削除へインサート
    dbSQLServer.Execute (strSQL)
    
    InsertRequDele = True
    
Exception:
    If InsertRequDele = False Then
        Call Err.Raise(Err.Number, "InsertRequDele" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function
' ***** end 2018/03/19 add by EGL
