VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************  strat of program ********************************
'**********************************************
'   ユーザ名　　　　：　（株）加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　契約入力
'   プログラムＩＤ　：
'   作　成　日　　　：  2005/07/13
'   作　成　者　　　：  T.SUZUKI
'   Ver             ：  0.0
'**********************************************
'修正履歴
'   修　正　日　　　：2005/08/03
'   修　正　者      ：T.SUZUKI
'   修　正　内　容　：1.解約関連の項目追加
'                   : 2.請求ファイルへのデータ更新処理追加
'                   : 3.解約日付入力時、コンテナ契約ファイルのコンテナ状態の値は『-1』に更新しないよう修正
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日      ：2005/08/31
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：1.コンテナ契約削除ファイルへのデータ登録時のSQL発行方法を修正
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   Ver             ：0.3
'   修　正　日      ：2005/09/15
'   修　正　者      ：K.ISHIZAKA
'   修　正　内　容  ：顧客マスタ変更からのデータ転送に漏れている項目を追加
'**********************************************
'修正履歴
'   Ver             ：0.4
'   修　正　日      ：2005/09/16
'   修　正　者      ：K.ISHIZAKA
'   修　正　内　容  ：修正モードでのデータ表示に漏れている項目を追加
'                     ヤードコード、コンテナ番号は必須入力
'                     そのたいろいろ
'**********************************************
'修正履歴
'   Ver             ：0.5
'   修　正　日      ：2005/10/07
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：[広告媒体]、[鍵区分]、[保証区分]を必須入力とするよう修正
'**********************************************
'修正履歴
'   Ver             ：0.6
'   修　正　日      ：2005/10/19
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：1.新規モードで予約受付からデータ取得時、既存顧客の場合、
'                       既存の顧客データが画面に正しく設定されるように処理を修正
'                     2.顧客カナコンボボックスをテキストボックスに変更
'                       2005/10/21 K.ISHIZAKA
'                     3.顧客カナの背景色を青→白にする
'                     4.顧客選択時リストを更新する
'                     5.cmb_顧客コードのリストは100件までとする
'**********************************************
'修正履歴
'   Ver             ：0.7
'   修　正　日      ：2005/11/10
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：請求ファイルの更新の更新条件を修正
'**********************************************
'修正履歴
'   Ver             ：0.8
'   修　正　日      ：2005/12/19
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：1.画面項目「種別」(コンテナマスタ.コンテナ種別) →「サイズ」(コンテナマスタ.実帖)に変更
'                     2.契約データ作成の際、物件チェックを行うように処理を追加
'**********************************************
'修正履歴
'   Ver             ：0.9
'   修　正　日      ：2006/01/05
'   修　正　者      ：T.SUZUKI
'   修　正　内　容  ：1.移動予約対応
'                     2.日付項目の入力チェック変更(｢/｣ありでの入力を可能とした)
'**********************************************
'修正履歴
'   Ver             ：1.0
'   修　正　日      ：2006/02/17
'   修　正　者      ：H.TAJIMA
'   修　正　内　容  ：受付対応機能対応
'                     1.新規ボタン廃止
'                     2.入力項目→表示項目に
'                     3.キャンペーン満了日追加
'**********************************************
'修正履歴
'   Ver             ：1.1
'   修　正　日      ：2006/05/09
'   修　正　者      ：H.TAJIMA
'   修　正　内　容  ：受付対応機能対応（改）
'                     1.承認日を入力可能に
'                     2.承認番号を入力可能に
'**********************************************
'修正履歴
'   Ver             ：1.2
'   修　正　日      ：2006/05/18
'   修　正　者      ：H.TAJIMA
'   修　正　内　容  ：受付対応機能対応（改）
'                     1.移動予約チェック外し
'**********************************************'修正履歴
'   Ver             ：1.3
'   修　正　日      ：2006/05/24
'   修　正　者      ：H.TAJIMA
'   修　正　内　容  ：鍵＆カード対応
'                     1.鍵情報修正可
'                     2.カード貸出状態表示
'**********************************************
'   Ver             ：1.4
'   修　正　日      ：2006/07/28
'   修　正　者      ：H.TAJIMA
'   修　正　内　容  ：鍵＆カード対応
'                     1.FVS601モードでもカード追加可能とした
'**********************************************
'   Ver             ：1.5
'   修　正　日      ：2006/09/22
'   修　正　者      ：SHIBAZAKI
'   修　正　内　容  ：修繕対応
'                     1.ヤード非表示日を過ぎたヤードは選択リストから外す
'                     2.ヤード営業終了日、ヤード解約予定日追加
'                     3.解約日の入力はヤード解約予定日までとする
'                     4.予約客先名を削除
'                     5.電話番号を削除
'                     6.コンテナの直近修繕を表示
'                     7.顧客クレームを表示
'                     8.地主情報削除
'**********************************************
'   Ver             ：1.6
'   修　正　日      ：2007/02/08
'   修　正　者      ：tajima
'   修　正　内　容  ：カード依頼書対応
'                     1.セキュリティカードメンテをダイアログで開くとプレビューの
'                       制御ができないので通常ウインドウとした
'**********************************************
'   Ver             ：1.7
'   修　正　日      ：2007/07/11
'   修　正　者      ：tajima
'   修　正　内　容  ：契約リスト
'                     1.顧客マスタメンテをダイアログで開くとプレビューの
'                       制御ができないので通常ウインドウとした
'**********************************************
'   Ver             ：1.8
'   修　正　日      ：2008/03/23
'   修　正　者      ：tajima
'   修　正　内　容  ：解約区分を「未使用キャンセル」にすれば自動的に解約日を
'                     契約開始日の前月末日とする
'**********************************************
'   Ver             ：1.9
'   修　正　日      ：2008/05/23
'   修　正　者      ：tajima
'   修　正　内　容  ：解約区分が「通常:01」を選ばれているだけでは解約入力必須としない
'**********************************************
'   Ver             ：2.0
'   修　正　日      ：2008/09/25
'   修　正　者      ：SHIBAZAKI
'   修　正　内　容  ：臨時保証区分対応
'**********************************************
'   Ver             ：2.1
'   修　正　日      ：2008/11/17
'   修　正　者      ：SHIBAZAKI
'   修　正　内　容  ：保証金のラベルはNAME_MASTの保証区分のTOの値から雑費名を取得し、表示する。
'**********************************************
'   Ver             ：2.2
'   修　正　日      ：2009/04/01
'   修　正　者      ：KIHARA
'   修　正　内　容  ：保険会社コードを追加
'**********************************************
'   Ver             ：2.3
'   修　正　日      ：2009/12/08
'   修　正　者      ：M.RYU
'   修　正　内　容  ：鍵種別を追加、フォームで鍵コンボボックスの幅を広く修正
'**********************************************
'   Ver             ：2.4
'   修　正　日      ：2010/04/10
'   修　正　者      ：K.ISHIZAKA
'   修　正　内　容  ：請求書発行区分を追加、削除時に保証会社コードが不足していたので追加
'**********************************************
'   Ver             ：2.5
'   修　正　日      ：2011/04/11
'   修　正　者      ：K.ISHIZAKA
'   修　正　内　容  ：緊急連絡先追加
'**********************************************
'   Ver　　         ：2.6
'   修　正　日　　　：2011/09/27
'   修　正　者　　　：M.RYU
'   修　正　内　容　：移動区分を追加⇒07紹介者金額変更
'**********************************************
'   Ver　　         ：2.7
'   修　正　日　　　：2012/04/11
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：スロープ貸出を追加
'**********************************************
'   Ver　　         ：2.8
'   修　正　日　　　：2012/07/04
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：ログ出力対応
'**********************************************
'   Ver　　         ：3.0
'   修　正　日　　　：2013/05/09
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：画面上に種別を追加
'**********************************************
'   Ver　　         ：3.1
'   修　正　日　　　：2013/06/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：終了メッセージを表示させないように修正
'**********************************************
'   Ver　　         ：3.2
'   修　正　日　　　：2013/09/19
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：集金区分の初期表示をコンビニから振込へ変更
'**********************************************
'   Ver　　         ：3.3
'   修　正　日　　　：2014/03/04
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：保証加入費の金額を修正
'**********************************************
'   Ver　　         ：3.4
'   修　正　日　　　：2014/04/16
'   修　正　者　　　：MIYAMOTO
'   修　正　内　容　：解約入力時、鍵区分が自動鍵、且つ契約状態が解約予定の場合、自動鍵解除番号発番トランを登録する
'**********************************************
'   Ver　　         ：3.5
'   修　正　日　　　：2014/06/24
'   修　正　者　　　：MIYAMOTO
'   修　正　内　容　：解約入力ボタン押下時、契約状態が解約予定の場合、鍵区分が変換鍵対象ならコンテナマスタの更新、
'                     且つ変換鍵区分が自動鍵なら自動鍵解除番号発番トランを登録する(地方ヤード以外が対象)
'                   ：スペアキー有無の表示を追加
'**********************************************
'   Ver　　         ：3.6
'   修　正　日　　　：2014/12/12
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：.Netの顧客検索から連動されたときにも、受付参照や契約更新などを可能とする
'**********************************************
'   Ver　　         ：3.7
'   修　正　日　　　：2015/03/26
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：自動鍵の際の、コンテナマスタの更新を修正
'**********************************************
'   Ver　　         ：3.8
'   修　正　日　　　：2016/05/19
'   修　正　者　　　：MIYAMOTO
'   修　正　内　容　：地方ヤードの鍵区分を考慮した条件に変更
'**********************************************
'   Ver　　         ：3.9
'   修　正　日　　　：2017/02/10
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：予約トランが契約のデータのみを表示
'**********************************************
'   Ver　　         ：4.0
'   修　正　日　　　：2017/07/29
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：未使用キャンセルの場合にコンテナ（バイク）をｺﾝﾃﾅに戻す
'**********************************************
'   Ver　　         ：4.1
'   修　正　日　　　：2018/02/10
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：保証会社新プラン対応
'**********************************************
'   Ver　　         ：4.2
'   修　正　日　　　：2018/08/21
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：解約日の入力チェック条件を変更
'**********************************************
'   Ver　　         ：4.3
'   修　正　日　　　：2019/02/19
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：保証加入の判定を金額から雑費コードに変更
'**********************************************
'   Ver　　         ：4.4
'   修　正　日　　　：2019/04/23
'   修　正　者　　　：TAKENOUCHI/Y_SUZUKI
'   修　正　内　容　：ネット解約で入力後、電話(解約届書）で訂正された場合の対応
'                     (1) 顧客問合わせトラン登録処理
'                     (2) ネット解約申込されたデータが有る場合の確認画面表示とボタン押下時の処理を追加
'**********************************************
'   Ver　　         ：4.5
'   修　正　日　　　：2019/06/15
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：未使用キャンセルの場合にバイク（コンテナ）をバイクに戻す
'**********************************************
'   Ver　　         ：4.6
'   修　正　日　　　：2019/08/08
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：SELECT文で*を使用していたがAccessのRecordSetは255カラムまでの制限があり、
'                     テーブルへのカラム追加によりその上限をオーバーしてしまった問題に対応
'**********************************************
'   Ver　　         ：4.7
'   修　正　日　　　：2020/12/15
'   修　正　者　　　：EGL
'   修　正　内　容　：雑費コードが「68」でかつ共益費が0円の場合、本来なら補償加入は「加入済」になるべきだが
'                     「未加入」になってしまう問題を改善
'**********************************************
'   Ver　　         ：4.8
'   修　正　日　　　：2021/06/02
'   修　正　者　　　：EGL
'   修　正　内　容　：QR番号対応
'**********************************************
'   Ver　　         ：4.9
'   修　正　日　　　：2021/08/20
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：契約番号を選択可能とする
'**********************************************
'   Ver　　         ：5.0
'   修　正　日　　　：2023/05/02
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：fncChkExistEnterTrnをコメントアウト
'                     メンテ業者と巡回業者を表示
'**********************************************
'   Ver　　         ：5.1
'   修　正　日　　　：2023/06/23
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：契約形態・スペア・セキュリティーを追加
'**********************************************
'   Ver　　         ：5.2
'   修　正　日　　　：2024/01/26
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：セキュリティー会社取得条件変更
'**********************************************
Option Compare Database
Option Explicit

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Const PROG_ID             As String = "FVS600"   ' フォームID  ' 2005/08/03 DEL T.SUZUKI

Private Const PROG_ID             As String = "FVS600"

Private pintAns                   As Integer             ' 確認入力用

Private Const strMsg_kaku           As String = "確認"
Private Const strMsg_err            As String = "エラー"

Private Const strMSGInputError      As String = "を入力してください。"
Private Const strMSGNoDataError     As String = "存在しないコードです。"
Private Const strMSGKetaOver        As String = "桁数オーバーしています。"
' 2006/01/05 ADD T.SUZUKI Start
Private Const strMSGMoveError       As String = "移動元がメンテ中または完了のため、" & vbCrLf & _
                                              "契約登録できません。"
Private Const strMSGMoveBiko        As String = "より移動"
Private Const strKeiyakuChgError    As String = "契約開始日の翌月で無ければ契約変更は出来ません。"
' 2006/01/05 ADD T.SUZUKI End

Private pWRI_objDb                  As Database
Private pWRI_ObjWS                  As Workspace
Private pobjDB                      As Database

Private pvarHenrit                  As Variant
Private pvarMaerit                  As Variant
Private pvarAtorit                  As Variant
Private pvarDATE                    As Variant

Private pvarKYDATE                  As Variant                  ' 地主契約マスタ.解約日付→ ヤードの解約予定日とした2006/09/22

Private pvarKcode                   As Variant                  ' 顧客コード格納
Private pvarContna                  As Variant
Private pintError                   As Integer

Private pintJARG_CHK                As Integer                  ' コンテナ契約ファイルチェックフラグ

Private pblnOpenFlag_FVS600         As Boolean                  ' フォームが読込まれたかどうかのフラグ
Private pblnChangeFlag_FVS600       As Boolean                  ' 入力に変更があったかどうかのフラグ

Private svFSDATE                    As Variant                  ' 契約開始日
Private pstrJUCHUKEIYAKU_NO         As String                  ' 受注契約番号保存用

Private Form_OpenArgs               As Variant                  ' 呼出元からのパラメータ
                                                              ' （オープンモード + 受付番号 または オープンモード + ヤードコード + コンテナ番号）
Private pintOpenMode                As Integer                  ' フォームオープンモード
                                                              ' ※0：メインメニューから遷移（新規契約）
                                                              '   1：ヤード利用状況画面から遷移（新規契約・取込あり）
                                                              '      契約新規かつ受付番号を選択したとき
                                                              '   2：メインメニューから遷移（契約修正）
                                                              '   3：ヤード利用状況画面からの遷移（契約参照）
                                                              '   4：契約削除
                                                              '   9：予約一覧画面から遷移（新規契約・取込あり）
Private pstrUKNO                    As String                   ' 予約受付トラン.受付番号
Private pstrYCode                   As String                   ' 予約受付トラン.ヤードコード
Private pstrNo                      As String                   ' 予約受付トラン.コンテナ番号

Private Const plngInputColor        As Long = 16777215          ' 顧客情報の背景色(入力可)
Private Const plngLockColor         As Long = 14680063          ' 顧客情報の背景色(入力不可)

Private plngUSER_NO                 As Long                     ' 新規顧客データ登録用顧客コード格納変数

Private pblnMoneyChgFlg             As Boolean                  ' 2005/08/03 ADD T.SUZUKI

Private Const pstrFVS600            As String = "FVS600"        ' 2005/08/03 ADD T.SUZUKI
Private Const pstrFVS601            As String = "FVS601"        ' 2005/08/03 ADD T.SUZUKI
Private Const pstrFVS602            As String = "FVS602"        ' 2006/03/18 ADD tajima
Private Const pstrFVS700            As String = "FVS700"        ' 2005/08/03 ADD T.SUZUKI
Private Const pstrFVS701            As String = "FVS701"        ' 2005/08/03 ADD T.SUZUKI
Private Const pstrFVS500            As String = "FVS500"        ' 2006/01/06 ADD T.SUZUKI
Private Const pstrFVS400            As String = "FVS400"        ' 2006/03/04 ADD T.SUZUKI
Private pstrFormName                As String                   ' 2005/08/03 ADD T.SUZUKI フォームID格納
Private Const pstrKaiyakuMsg        As String = "(解約入力修正中)" ' 2005/08/03 ADD T.SUZUKI 解約入力画面表示中の画面の標題
Private pstrAcptNo                  As String                   ' 2005/08/03 ADD T.SUZUKI 受注契約番号格納
Private pintAGRE                    As Integer                  ' 2005/08/03 ADD T.SUZUKI コンテナ契約ファイル.契約状態格納

Private objKONT_DB                  As Database                 ' 2005/08/31 ADD T.SUZUKI コンテナDB
Private pobjKASE_DB                 As Database                 ' 2009/04/01 INS KIHARA   加瀬DB

' 2006/01/05 ADD T.SUZUKI Start
Private Const P_TITTLE_KEIYAKU      As String = "契約入力"

Private Const pcstrfmtSTANDARD      As String = "Standard"    ' 書式：標準
Private Const pcstrfmtFIXED         As String = "Fixed"       ' 書式：固定
Private Const pcsglMoneyMaxValue    As Long = 999999999       ' 金額項目の入力最大値
Private Const pcsglMoneySumMaxValue As Double = 9999999999#   ' 請求ファイル作成時の金額合計値の最大値
Private Const pcintBikoMaxLength    As Long = 60              ' コンテナ契約ファイル.備考の最大桁数

' 【予約受付状態区分】
Private Const P_YUKBN_10            As Integer = 10           ' 受付(予約受付トラン.予約受付状態区分)
Private Const P_YUKBN_20            As Integer = 20           ' 契約(予約受付トラン.予約受付状態区分)
Private Const P_YUKBN_51            As Integer = 51           ' 予約キャンセル(予約受付トラン.予約受付状態区分)

Private Const P_ZERO                As Integer = 0            ' 固定値(0)
Private Const P_FORMAT_YMD          As String = "yyyy/mm/dd"  ' 日付書式

' 【移動区分】
Private Const P_MOVEKBN_00          As String = "00"          ' 無し
Private Const P_MOVEKBN_01          As String = "01"          ' 自社理由移動
Private Const P_MOVEKBN_02          As String = "02"          ' 顧客理由移動
Private Const P_MOVEKBN_03          As String = "03"          ' 未使用移動
Private Const P_MOVEKBN_04          As String = "04"          ' 名義変更
Private Const P_MOVEKBN_05          As String = "05"          ' 金額変更
Private Const P_MOVEKBN_06          As String = "06"          ' 保証会社変更
Private Const P_MOVEKBN_07          As String = "07"          ' 紹介者金額変更  'INSERT 2011/09/27 M.RYU

Private Const P_NAME_098            As String = "098"

' 【契約状態】
Private Const P_AGRE_01             As String = "01"          ' 契約
Private Const P_AGRE_02             As String = "02"          ' 解約予定
Private Const P_AGRE_03             As String = "03"          ' 解約延長
Private Const P_AGRE_04             As String = "04"          ' メンテ中
Private Const P_AGRE_09             As String = "09"          ' 完了

' 【広告媒体】
Private Const P_ADMEI_99            As String = "99"          ' その他

' 【契約形態】
Private Const P_AGRE_99             As String = "99"          ' その他

' 【ご紹介区分】
Private Const P_INTROKBN_02         As String = "2"           ' 受付
Private Const P_INTROKBN_03         As String = "3"           ' 契約

' 移動予約画面遷移区分
Private Const P_MOVEYOYAKU_0        As Integer = 0             ' 移動予約(新規モード)
Private Const P_MOVEYOYAKU_1        As Integer = 1             ' 移動予約(修正モード)

' 追加予約画面遷移区分
Private Const P_YOYAKUADD           As Integer = 2             ' 追加予約

Private VarChangeDate               As Variant

' 予約受付トラン構造体
Private Type Type_YOUK_TRAN
    strMOTO_ACPTNO                  As String                 ' 予約受付トラン.元受注契約番号
    strMOVEKBN                      As String                 ' 予約受付トラン.移動区分
    strUkDate                       As String                 ' 予約受付トラン.受付日
    strUKTANTO                      As String                 ' 予約受付トラン.受付担当者コード
End Type

Private SAKI_YOUK_TRAN              As Type_YOUK_TRAN

' コンテナ契約ファイル構造体
Private Type Type_CARG_FILE
    strYcode                        As String                 ' コンテナ契約ファイル.ヤードコード
    strNO                           As String                 ' コンテナ契約ファイル.コンテナ番号
    strUCODE                        As String                 ' コンテナ契約ファイル.顧客コード
    strAGRE                         As String                 ' コンテナ契約ファイル.契約状態
    dblRENTKG                       As Double                 ' コンテナ契約ファイル.月額料金
    strACPTNO                       As String                 ' コンテナ契約ファイル.受注契約番号
    strHOSYICD                      As String                 ' コンテナ契約ファイル.保証区分
    strUKTANTO                      As String                 ' コンテナ契約ファイル.契約受付担当者コード
End Type

Private MOTO_CARG_FILE              As Type_CARG_FILE

Private pblnUpdateFlg               As Boolean                ' 変更フラグ
' 2006/01/05 ADD T.SUZUKI End
Private LimitationStartDate         As Variant                ' 契約開始日制限値 by Tajima

' 2006/03/04 ADD T.SUZUKI Start
Private Const P_SHU                 As Integer = 0            ' 契約修正モード
Private Const P_DEL                 As Integer = 1            ' 契約削除モード

Private pblnFKS430_Flag             As Boolean                ' 顧客マスタメンテ画面Open時：True
' 2006/03/04 ADD T.SUZUKI End

' 2006/03/12 ADD tajima
Private isRCPT_TRAN                 As Boolean                ' 受付トラン有無
Private pstrCampaignInfomationText  As String                 ' キャンペーン適用情報文言
Private pstrCampaignWarningText     As String                 ' キャンペーン解約警告文言
' 2006/05/10 ADD tajima
Private pOriginalKeyType            As String                ' 読込時の鍵区分
Private pOriginalKeyNo              As String                ' 読込時の鍵番号
Private pOriginalKeyNum             As Integer               ' 読込時の鍵本数
' 2019/04/23 ADD TAKENOUCHI Start
Private pstrLoginUser               As String                'ネット解約DB用の接続ユーザ
Private pstrLoginPass               As String                'ネット解約DB用の接続パスワード
Private pstrLoginKey                As String                'ネット解約DB用のログインキー
Private piniFile                    As Collection
Private pstrKaiYtiYmd               As String                '申込トラン.解約予定日

' 2019/04/23 ADD TAKENOUCHI End
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_スロープ_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
Private Sub cmb_スロープ_AfterUpdate()

    Dim vCode       As Variant
    Dim strName     As String

    vCode = Nz(Me.cmb_スロープ, "")
    If vCode = "" Then
        Me.txt_スロープ = ""
        Exit Sub
    End If
    Me.cmb_スロープ = Format(vCode, "00")

    strName = fncSelectName("294", Me.cmb_スロープ)

    Me.txt_スロープ = strName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_スロープ_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
Private Sub cmb_スロープ_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num(Me.cmb_スロープ, KeyAscii, 2)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_スロープ_Change
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_スロープ_Change()

    Call subSetChangeFlag
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_スロープ_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
Private Sub cmb_スロープ_BeforeUpdate(Cancel As Integer)

    Dim strName     As String

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

    ' コードが空の場合解約区分欄も空欄表示
    If fncControlCrear(Me.cmb_スロープ, Me.txt_スロープ) = False Then Exit Sub

    '  スロープ名取得
    strName = fncSelectName("294", Me.cmb_スロープ)
    If strName = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub


'==============================================================================*
'
'       MODULE_NAME     : データが更新後
'       MODULE_ID       : cmb_請求書発行区分_AfterUpdate
'       CREATE_DATE     : 2010/04/10            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_請求書発行区分_AfterUpdate()
    On Error GoTo ErrorHandler
    
    txt_請求書発行区分.VALUE = cmb_請求書発行区分.Column(1)
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_請求書発行区分_AfterUpdate", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : データが変わったとき
'       MODULE_ID       : cmb_請求書発行区分_Change
'       CREATE_DATE     : 2010/04/10            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_請求書発行区分_Change()
    On Error GoTo ErrorHandler
    
    Call subSetChangeFlag
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_請求書発行区分_Change", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キーを押したとき
'       MODULE_ID       : cmb_請求書発行区分_KeyPress
'       CREATE_DATE     : 2010/04/10            K.ISHIZAKA
'       PARAM           : KeyAscii              アスキーコード(I/O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_請求書発行区分_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    
    Call CmKeyPress_For_Num(Me.cmb_請求書発行区分, KeyAscii, 1)
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_請求書発行区分_KeyPress", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : リストにない入力をしたとき
'       MODULE_ID       : cmb_請求書発行区分_NotInList
'       CREATE_DATE     : 2010/04/10            K.ISHIZAKA
'       PARAM           : NewData               データ(I/O)
'                       : Response              レスポンス(I/O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_請求書発行区分_NotInList(NewData As String, Response As Integer)
    On Error GoTo ErrorHandler
    
    MsgBox strMSGNoDataError, vbInformation, strMsg_err
    Response = acDataErrContinue
    cmb_請求書発行区分.Undo
Exit Sub
    
ErrorHandler:          '↓自分の関数名
    Call MSZZ024_M00("cmb_請求書発行区分_NotInList", True)   '←親となる関数に対してだけ呼び出しを記述
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 郵便番号変換
'       MODULE_ID       : cmd_KinkyuYubin_Click
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_KinkyuYubin_Click()
    Call MSZ0000_YUBIB(Me.NAME, txt_RCPT_KREN_YUBINO.NAME, txt_RCPT_KREN_ADR_1.NAME, _
                                Nz(txt_RCPT_KREN_YUBINO.VALUE), Nz(txt_RCPT_KREN_ADR_1.VALUE))
End Sub

Private Sub cmd_Search_Click()
  doCmd.OpenForm "FKS522_契約入力検索", acNormal, , , , acDialog
  
  Call cmb_ヤードコード_AfterUpdate
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_UketsukeView_Clickイベント
'        MODULE_ID        :cmd_UketsukeView_Click
'        CREATE_DATE      :2006/03/04
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_UketsukeView_Click()
        
    ' ヤードコード、コンテナ番号が未選択の状態では次画面に遷移させない
    If (Replace(Nz(Me.cmb_ヤードコード, ""), "0", "") = "" And Replace(Nz(Me.cmb_コンテナ番号, ""), "0", "") = "") Then
        Exit Sub
    End If

    ' 受注契約番号がNullの状態のとき、処理を抜ける
    If Nz(Me.txt_受注契約番号, "") = "" Then Exit Sub     ' 2006/03/18 ADD T.SUZUKI

    If isRCPT_TRAN = False Then
        MsgBox "この契約には受付情報はありません。": Exit Sub
    End If

    If Nz(pstrUKNO, "") = "" Then
        MsgBox "受付番号がNullです。": Exit Sub
    End If

    '受付入力画面起動
    Select Case Me.Form.NAME
        Case pstrFVS600, pstrFVS602
            doCmd.OpenForm pstrFVS400, , , , , , Me.NAME & "," & pstrUKNO & "," & pstrUKNO & ","

        Case pstrFVS601
            doCmd.OpenForm pstrFVS400, , , , , acDialog, Me.NAME & "," & pstrUKNO & "," & pstrUKNO & ","
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_YoyakuAdd_Clickイベント
'        MODULE_ID        :cmd_YoyakuAdd_Click
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_YoyakuAdd_Click()

    Dim vHData As Variant

' 2006/03/04 ADD T.SUZUKI Start
    ' ヤードコード、コンテナ番号が未選択の状態では次画面に遷移させない
    If (Replace(Nz(Me.cmb_ヤードコード, ""), "0", "") = "" And Replace(Nz(Me.cmb_コンテナ番号, ""), "0", "") = "") Then
        Exit Sub
    End If
' 2006/03/04 ADD T.SUZUKI End

    ' 予約入力画面を追加予約モードで開く
    ' ※パラメータ(ヤードコード・顧客コード・画面ID)
    vHData = Format$(Me.cmb_ヤードコード, "000000") & Space(6) & Space(11) & _
             P_YOYAKUADD & Format(Me.cmb_顧客コード, "000000") & Me.Form.NAME

    '予約入力画面起動(予約追加)
    doCmd.OpenForm pstrFVS500, , , , , , vHData
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd_KeiyakuRep_Clickイベント
'        MODULE_ID        :cmd_KeiyakuRep_Click
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_KeiyakuRep_Click()

    Dim strACPTNO As String  ' 受注契約番号
    Dim lngYCODE  As Long    ' ヤードコード
    Dim lngNo     As Long    ' コンテナ番号
    Dim strUkno   As String  ' 予約受付番号
    Dim intRtn    As Integer
    Dim vHData    As Variant

    strACPTNO = Nz(Me.txt_受注契約番号, "")
    If strACPTNO = "" Then Exit Sub

' 2006/05/18 DEL tajima Start
'' 2006/02/02 ADD T.SUZUKI Start
'    If Format(DATE, "YYYY/MM") > Format(Me.txt_契約開始日, "YYYY/MM") Then
'        ' チェックOK
'    Else
'        ' チェックNG
'        MsgBox strKeiyakuChgError, vbInformation, strMsg_err
'        Exit Sub
'    End If
'' 2006/02/02 ADD T.SUZUKI End
' 2006/05/18 DEL tajima End

    ' 予約受付トランデータチェック
    intRtn = fncCheckMoveYoyaku(strACPTNO, lngYCODE, lngNo, strUkno)

    vHData = ""
    Select Case intRtn
        Case P_MOVEYOYAKU_0
            ' *** 該当データなし ***
            ' → 予約入力画面を新規(受付)モードで開く
            '    ※パラメータ(ヤードコード・コンテナ番号・顧客コード・受注契約番号・画面ID)
            vHData = Format$(Me.cmb_ヤードコード, "000000") & Format$(Me.cmb_コンテナ番号, "000000") & Space(11) & _
                     P_MOVEYOYAKU_0 & Format(Me.cmb_顧客コード, "000000") & strACPTNO
'            vHData = Format$(lngYCODE, "000000") & Format$(lngNO, "000000") & Space(11) & _
'                     P_MOVEYOYAKU_0 & Format(Me.cmb_顧客コード, "000000") & strACPTNO & Me.Form.NAME

        Case P_MOVEYOYAKU_1
            ' *** 該当データあり ***
            ' → 予約入力画面を修正モードで開く
            '    ※パラメータ(ヤードコード・コンテナ番号・予約受付番号・画面ID)
            vHData = Format$(lngYCODE, "000000") & Format$(lngNo, "000000") & strUkno & _
                     P_MOVEYOYAKU_1
'            vHData = Format$(lngYCODE, "000000") & Format$(lngNO, "000000") & strUKNO

        Case -1
            ' システムエラー
            Exit Sub
    End Select

    ' 予約入力画面　起動
    doCmd.OpenForm pstrFVS500, , , , , , vHData
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_カードメンテ_Click
'        CREATE_DATE      :2006/06/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_カードメンテ_Click()
    
    If Nz(Me.txt_受注契約番号, "") = "" Then Exit Sub

    '起動！カードメンテ画面
'    doCmd.OpenForm "FKS600", , , , , acDialog, pstrFVS600 & "," _
' 2006/07/28 Rep ↑旧↓新
'    doCmd.OpenForm "FKS600", , , , , acDialog, Me.Form.NAME & "," _
'                                     & Me.cmb_ヤードコード & "," _
'                                     & Me.cmb_コンテナ番号 & "," _
'                                     & Me.txt_受注契約番号
' 2007/02/08 Rep ↑旧↓新
    doCmd.OpenForm "FKS600", , , , , , Me.Form.NAME & "," _
                                     & Me.cmb_ヤードコード & "," _
                                     & Me.cmb_コンテナ番号 & "," _
                                     & Me.txt_受注契約番号

End Sub

'==============================================================================*
'
'        MODULE_NAME      :適用キャンペーンの表示
'        MODULE_ID        :cmd_キャンペーン内容_Click
'        CREATE_DATE      :2006/03/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_キャンペーン内容_Click()

  If isRCPT_TRAN = True And pstrCampaignInfomationText <> "" Then
      MsgBox pstrCampaignInfomationText, vbOKOnly + vbInformation, "この契約で適用したキャンペーンについて"
  End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Activate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Activate()
    ' フルスクリーン化
    'If Me.Form.NAME <> pstrFVS601 Then  2007/02/08 del tajima
      doCmd.Maximize
    'End If  2007/02/08 del tajima
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Open
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
    Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Keiyaku.BackColor = RGB(255, 0, 204)
    End Select
    ' 初期化
    Form_OpenArgs = ""
    pintOpenMode = 2                                                            '2006/02/17 Rep Tajima 以前は 0
    pstrUKNO = ""
    pstrYCode = ""
    pstrNo = ""
    pblnChangeFlg = False
    pblnFKS430_Flag = False                                                     ' 2006/03/04 ADD T.SUZUKI

    Call MSHS010_M00(PROG_ID, "1")


' 2005/08/03 REP T.SUZUKI Start
    Form_OpenArgs = Nz(OpenArgs, "")
    If Form_OpenArgs <> "" Then
        pintOpenMode = CInt(Mid(Form_OpenArgs, 1, 1))  ' オープンモード

        Select Case pintOpenMode
            Case 1
                pstrUKNO = Mid(Form_OpenArgs, 2, 11)   ' 受付番号

'            Case 3  ' *** ヤード利用状況画面から遷移した場合 ***               'DELETE 2014/12/12 K.ISHIZAKA
            Case 2, 3  ' *** ヤード利用状況画面から遷移した場合 ***             'INSERT 2014/12/12 K.ISHIZAKA
                    '     (画面ID：FKS020_コンテナ予約入力)
'                pstrYCODE = Mid(Form_OpenArgs, 2, 6)   ' ヤードコード
'                pstrNO = Mid(Form_OpenArgs, 8, 6)      ' コンテナ番号
                pstrAcptNo = Mid(Form_OpenArgs, 2, 10)  ' 受注契約番号

                ' コンテナ契約ファイル存在チェック
                ' (ヤードコード、コンテナ番号をセット)
                If fncChkCARG_FILE(pstrAcptNo) = False Then
                    Cancel = 1  ' フォーム読込をキャンセル
                End If

' 2006/01/05 ADD T.SUZUKI Start
            Case 9  ' *** 予約一覧から遷移した場合 ***
                    '     (画面ID：FVS540)
                pstrUKNO = Mid(Form_OpenArgs, 2, 11)   ' 受付番号
' 2006/01/05 ADD T.SUZUKI End
        End Select
    End If

'    Form_OpenArgs = Nz(OpenArgs, "")
'    If Form_OpenArgs <> "" Then
'        pintOpenMode = CInt(Mid(Form_OpenArgs, 1, 1))  ' オープンモード
'        Select Case pintOpenMode
'            Case 1
'                pstrUKNO = Mid(Form_OpenArgs, 2, 11)   ' 受付番号
'
'            Case 2  ' 2005/08/03 DEL T.SUZUKI
'                pstrYCODE = Mid(Form_OpenArgs, 2, 6)   ' ヤードコード
'                pstrNO = Mid(Form_OpenArgs, 8, 6)      ' コンテナ番号
'        End Select
'    End If
' 2005/08/03 REP T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Load
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Load()

    ' 画面を最大化する
    ' If Me.Form.NAME <> pstrFVS601 Then  2007/02/08 del tajima
      doCmd.Maximize
    ' End If  2007/02/08 del tajima

    pblnOpenFlag_FVS600 = False  ' フォーム読込フラグOFF

    Set pobjDB = CurrentDb

    ' フォームタイトル設定
    Call subSetFormTittle        ' 2005/08/03 ADD T.SUZUKI

    ' 画面の初期化処理(全項目初期化)
    Call CLR_RTN
    Call CLR_RTN3(0)  ' 【顧客情報】
    Call CLR_RTN3(1)  ' 【その他連絡先】
    Call CLR_RTN3(2)  ' 【緊急連絡先】                                          'INSERT 2011/04/11 K.ISHIZAKA

    Call subLock_Main        ' ヤードコード、コンテナ番号の使用可否設定
'    Call subLock_YoyakuItem  ' 予約受付からチェックボックス、予約番号の使用可否設定 'DEL 2006/03/03 Tajima

    Select Case pintOpenMode
        Case 0  '*** メインメニューより遷移 ***
            ' ---> 契約新規・取込なし
            Call subLock_Etc         ' ヤードコード、コンテナ番号以外の項目を使用不可に設定
'ADD 2006/03/26 Tajima Start
        Case 2  ' FormLoadでいきなり修正モード "2" になるんで
            '引数があったときは 3 と同じことをさせる
            If Nz(Me.OpenArgs) <> "" Then                                       'INSERT START 2014/12/12 K.ISHIZAKA
                Call subSetKaiyakuControl
                If Not fncSetYoyakuData(2) Then Exit Sub
                Me.cmd_削除.Enabled = False
            End If                                                              'INSERT END   2014/12/12 K.ISHIZAKA
            Me.cmb_ヤードコード.SetFocus
'ADD 2006/03/26 Tajima End
        Case 3  ' *** ヤード利用状況画面(コンテナ契約ファイル読込) ***
                ' ※解約入力項目のみ入力可能
            ' ---> 契約修正
'            Call subOptGroupControl    ' 2005/08/03 ADD T.SUZUKI  ' 2006/03/04 DEL T.SUZUKI

            ' 解約時入力不可項目の背景色設定
            Call subSetKaiyakuControl  ' 2005/08/03 ADD T.SUZUKI

            ' コンテナ契約ファイルのデータを画面にセット
            If fncSetYoyakuData(2) = False Then Exit Sub

            Me.cmd_削除.Enabled = False                            ' 2006/03/04 ADD T.SUZUKI

            ' ※画面の各項目の使用可否は『cmb_コンテナ番号_AfterUpdate』イベントにて『LOCK_RTN』をコール
            '   (フォーカス設定も含む)

' 2006/01/05 ADD T.SUZUKI Start
        Case 9  ' *** 予約一覧画面 ***
            ' ---> 契約新規・取込あり
'            Call subOptGroupControl                               ' 2006/03/04 DEL T.SUZUKI

' DEL 2006/03/03 Tajima Start
'            Me.chk_予約受付 = -1                    ' 予約受付からチェックボックスを選択状態に設定
'            Me.cmb_受付番号 = pstrUKNO              ' 予約一覧で指定された「予約受付番号」をセット
'            Call cmb_受付番号_AfterUpdate           ' 指定された「予約受付番号」の該当データを表示
'            ' 予約受付からチェックボックス、受付番号をし使用不可に設定
'            Me.chk_予約受付.Enabled = False         ' 予約受付からチェックボックス：OFF
'            Call CmLockControl(Me.cmb_受付番号, 1)  ' 受付番号コンボボックス      ：OFF
' DEL 2006/03/03 Tajima End
            
            ' フォーカスを移動しておく
            '2006/09/22 SHIBAZAKI 予約名削除
'            Me.txt_予約名.SetFocus    '元はtxt_契約開始日 2006/02/18 Add TAJIMA
            Me.cmd_終了.SetFocus
            
' 2006/01/05 ADD T.SUZUKI End
    End Select

    pblnOpenFlag_FVS600 = True  ' フォーム読込フラグON

    ' 変更フラグ初期化
    pblnUpdateFlg = False       ' 2006/01/05 ADD T.SUZUKI
   
End Sub

' 2006/03/04 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :subOptGroupControl
''        CREATE_DATE      :2004/07/19
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub subOptGroupControl()
'
'    With Me
'        Select Case pintOpenMode
'            Case 1
'            Case 2
'            Case 3
'                .opt_新規.Enabled = False
'                .opt_修正.Enabled = True
'                .opt_削除.Enabled = False
'
'                ' 修正ボタンをアクティブに設定
'                .opg_処理選択 = 2
'
'' 2006/01/26 ADD T.SUZUKI Start
'            Case 9
'                .opt_新規.Enabled = True
'                .opt_修正.Enabled = False
'                .opt_削除.Enabled = False
'
'                ' 新規ボタンをアクティブに設定
'                .opg_処理選択 = 1
'' 2006/01/26 ADD T.SUZUKI End
'        End Select
'    End With
'End Sub
' 2006/03/04 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subSetKaiyakuControl
'        CREATE_DATE      :2004/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetKaiyakuControl()

    With Me
        Call subSetColor(.txt_初回契約日, plngLockColor)
        Call subSetColor(.txt_契約開始日, plngLockColor)
        Call subSetColor(.txt_契約満了日, plngLockColor)
        Call subSetColor(.txt_解除番号, plngLockColor)
        '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'        Call subSetColor(.txt_予約名, plngLockColor)
'        Call subSetColor(.txt_予約番号, plngLockColor)
        Call subSetColor(.txt_承認日, plngLockColor)
        Call subSetColor(.txt_承認番号, plngLockColor)
        Call subSetColor(.txt_月額料金, plngLockColor)
        Call subSetColor(.txt_消費税, plngLockColor)
        'Call subSetColor(.txt_預り保証金, plngLockColor) DEL 2006/03/03 Tajima
        Call subSetColor(.txt_契約備考, plngLockColor)
' 2006/05/07 ADD Tajima Start
        Call subSetColor(.txt_スペアキー備考, plngLockColor)
        Call subSetColor(.txt_貸出本数, plngLockColor)
' 2006/05/07 ADD Tajima Endt
    End With
End Sub


'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :opg_処理選択_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_処理選択_BeforeUpdate(Cancel As Integer)

    ' 変更フラグがTrueの時は確認メッセージを出す
    If pblnChangeFlag_FVS600 = True Then
        pintAns = MsgBox("クリアされます。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku)
        If pintAns = vbNo Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    End If
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_ADR_1_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ADR_1_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_ADR_1.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "都道府県市町村" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_ADR_1.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_ADR_1_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ADR_1_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_ADR_2_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ADR_2_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_ADR_2.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "番地" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_ADR_2.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_ADR_2_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ADR_2_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_ADR_3_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ADR_3_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_ADR_3.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "建物名等" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_ADR_3.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_ADR_3_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ADR_3_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_FAX_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_FAX_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_FAX.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "FAX番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_FAX.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_FAX_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_FAX_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー押下時
'       MODULE_ID       : txt_RCPT_KREN_FAX_KeyPress
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_FAX_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_RCPT_KREN_FAX, KeyAscii, 15)  '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_KANA_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_KANA_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_KANA.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 40)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先名(カナ)" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_KANA.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_KANA_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_KANA_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_KEITAI_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_KEITAI_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_KEITAI.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "携帯番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_KEITAI.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_KEITAI_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_KEITAI_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー押下時
'       MODULE_ID       : txt_RCPT_KREN_KEITAI_KeyPress
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_KEITAI_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_RCPT_KREN_KEITAI, KeyAscii, 15)  '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_NAME_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_NAME_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_NAME.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 50)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先名" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_NAME.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_NAME_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_NAME_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_TEL_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_TEL_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_TEL.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "電話番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_TEL.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_TEL_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_TEL_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー押下時
'       MODULE_ID       : txt_RCPT_KREN_TEL_KeyPress
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_TEL_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_RCPT_KREN_TEL, KeyAscii, 15)  '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_YUBINO_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_YUBINO_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_YUBINO.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 10)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "郵便番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_YUBINO.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_YUBINO_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_YUBINO_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'       MODULE_NAME     : キー押下時
'       MODULE_ID       : txt_RCPT_KREN_YUBINO_KeyPress
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_YUBINO_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_RCPT_KREN_YUBINO, KeyAscii, 10)  '←DB定義より10桁まで入力可。
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新前
'       MODULE_ID       : txt_RCPT_KREN_ZOKUGARA_BeforeUpdate
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ZOKUGARA_BeforeUpdate(Cancel As Integer)
    Dim strCode             As String
    On Error GoTo ErrorHandler

    strCode = Nz(txt_RCPT_KREN_ZOKUGARA.VALUE, "")
    If strCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(strCode, 20)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "続柄" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            txt_RCPT_KREN_ZOKUGARA.Undo
            Exit Sub
    End Select
Exit Sub

ErrorHandler:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 更新時
'       MODULE_ID       : txt_RCPT_KREN_ZOKUGARA_Change
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RCPT_KREN_ZOKUGARA_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_スペアキー備考_BeforeUpdate
'        CREATE_DATE      :2006/05/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_スペアキー備考_BeforeUpdate(Cancel As Integer)
    
    Dim vCode As Variant

    vCode = Nz(Me.txt_スペアキー備考, "")

    '入力チェック
    Select Case CmBCheck_Char(vCode, 60)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "スペアキー備考を入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_スペアキー備考_Change
'        CREATE_DATE      :2006/05/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_スペアキー備考_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_スペアキー備考_KeyPress
'        CREATE_DATE      :2006/05/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_スペアキー備考_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Char2(Me.txt_スペアキー備考, KeyAscii, 60)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解除番号_KeyPress
'        CREATE_DATE      :2006/05/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解除番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Char2(Me.txt_解除番号, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_承認番号_KeyPress
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Char2(Me.txt_承認番号, KeyAscii, 20)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_貸出本数_BeforeUpdate
'        CREATE_DATE      :2006/05/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出本数_BeforeUpdate(Cancel As Integer)
    Dim vNum As Variant

    On Error GoTo err_rtn

    vNum = Nz(Me.txt_貸出本数, "")

    If vNum = "" Then Exit Sub

    If fncTxtInputCheck(Me.txt_貸出本数, "99", 2) = False Then
        MsgBox strMSGKetaOver, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_貸出本数_Change
'        CREATE_DATE      :2006/05/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出本数_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_貸出本数_KeyPress
'        CREATE_DATE      :2006/05/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_貸出本数_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_貸出本数, KeyAscii, 2)
End Sub

' 2006/03/04 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :opg_処理選択_AfterUpdate
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub opg_処理選択_AfterUpdate()
'
'    Select Case opg_処理選択
'        Case 1
'            ' *** 新規 ***
'            pintOpenMode = 0
'
'        Case 2
'            ' *** 修正 ***
'            pintOpenMode = 2
'
'        Case 3
'            ' *** 削除 ***
'            pintOpenMode = 4   ' 2005/08/03 ADD T.SUZUKI
''            pintOpenMode = 3  ' 2005/08/03 DEL T.SUZUKI
'    End Select
'
'    ' 画面の初期化処理(全項目初期化)
'    Call CLR_RTN
'    Call CLR_RTN3(0)  ' 【顧客情報】
'    Call CLR_RTN3(1)  ' 【その他連絡先】
'
'    ' ヤードコード、コンテナ番号の使用可否設定
'    Call subLock_Main
'
'    Call subLock_Etc
'
'    ' 予約受付チェックボックス、受付番号コンボの制御
''    Call subLock_YoyakuItem DEL 2006/03/03 Tajima
'
'    ' 処理モードによってフォーカスの移動先を変更する
'    If Me.opg_処理選択 = 1 Then
''        Me.chk_予約受付.SetFocus DEL 2006/03/03 Tajima
'    Else
'        Me.cmb_ヤードコード.SetFocus
'    End If
'End Sub
' 2006/03/04 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_預り保証金_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_預り保証金_AfterUpdate()

'' 2005/08/03 ADD T.SUZUKI Start
'    If Me.opg_処理選択.Value = 2 Then
'        pblnMoneyChgFlg = True
'    End If
'' 2005/08/03 ADD T.SUZUKI End
' 2006/03/04 DEL T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_預り保証金_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_預り保証金_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :フレームヤード_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub フレームヤード_Enter()

    If Me.cmb_ヤードコード.Enabled = True Then
        Me.cmb_ヤードコード.SetFocus
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_受付担当者コード_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付担当者コード_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    If pblnOpenFlag_FVS600 = False Then Exit Sub

    ' コードが空の場合担当者名欄も空欄表示
    If fncControlCrear(Me.cmb_受付担当者コード, Me.txt_受付担当者名) = False Then Exit Sub  ' 2005/08/03 ADD T.SUZUKI

' 2005/08/03 DEL T.SUZUKI Start
'    vCode = Nz(Me.cmb_受付担当者コード, "")
'    If vCode = "" Then
'        Exit Sub
'    End If
' 2005/08/03 DEL T.SUZUKI End

    '選択された担当者データの存在チェック
    If GetTantoData(Nz(Me.cmb_受付担当者コード, "")) = False Then
        MsgBox "登録されていません。", vbInformation, strMsg_kaku
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_受付担当者コード_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付担当者コード_AfterUpdate()

    Dim vCode   As Variant  ' 受付担当者コード
    Dim vTanto  As Variant  ' 受付担当者名

    On Error GoTo cmb_受付担当者コード_AfterUpdate_Err

    vCode = Nz(Me.cmb_受付担当者コード, "")
    If vCode = "" Then GoTo cmb_受付担当者コード_AfterUpdate_End

    ' ファイルへの存在チェック
    If GetTantoData(vCode, vTanto) = False Then GoTo cmb_受付担当者コード_AfterUpdate_End

    Me.txt_受付担当者名 = vTanto
    Me.cmb_受付担当者コード = Format$(Me.cmb_受付担当者コード, "000")

    GoTo cmb_受付担当者コード_AfterUpdate_End

cmb_受付担当者コード_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_受付担当者コード_AfterUpdate_End:
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_受付担当者コード_KeyPress
'        機能             :受付担当 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付担当者コード_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_受付担当者コード, KeyAscii, 3)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_受付担当者コード_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_受付担当者コード_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_ヤードコード_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_ヤードコード, "")

    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "コードを入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

    'ファイルへの存在チェック
    If Chk_Yard_Mast(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    '2006/09/22 SHIBAZAKI 地主情報削除
'    If CHK_JARG_FILE(0) = False Then
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If

' 2006/03/04 DEL T.SUZUKI Start
'    If Me.opg_処理選択.Value = 1 Then
'        If pintJARG_CHK <> 1 Then
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
' 2006/03/04 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_ヤードコード_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_AfterUpdate()

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_ヤードコード, "")

    If vCode = "" Then
        Call CLR_RTN                                                            'INSERT 20050916 K.ISHIZAKA
        Call CLR_RTN3(0)  ' 【顧客情報】                                        'INSERT 20050916 K.ISHIZAKA
        Call CLR_RTN3(1)  ' 【その他連絡先】                                    'INSERT 20050916 K.ISHIZAKA
        Call CLR_RTN3(2)  ' 【緊急連絡先】                                      'INSERT 2011/04/11 K.ISHIZAKA
        ' 画面項目のロック処理                                                  'INSERT 20050916 K.ISHIZAKA
        Call subLock_Main                                                       'INSERT 20050916 K.ISHIZAKA
        Call subLock_Etc                                                        'INSERT 20050916 K.ISHIZAKA
        Exit Sub
    End If
    cmb_コンテナ番号 = "000000"                                                 'INSERT 20050916 K.ISHIZAKA
    Call CLR_RTN2(2)                                                            'INSERT 20050916 K.ISHIZAKA
    ' 画面項目のロック処理                                                      'INSERT 20050916 K.ISHIZAKA
    Call subLock_Main                                                           'INSERT 20050916 K.ISHIZAKA
    Call subLock_Etc                                                            'INSERT 20050916 K.ISHIZAKA

    ' ファイルへの存在チェック（ヤードマスタ）
    If Chk_Yard_Mast(1) = False Then
        Exit Sub
    End If

    '2006/09/22 SHIBAZAKI 地主情報削除
'    ' 地主契約マスタデータ取得
'    If CHK_JARG_FILE(1) = False Then
'        Exit Sub
'    End If

    Me.cmb_ヤードコード = Format(Me.cmb_ヤードコード, "000000")

    Me.cmb_コンテナ番号.Requery
    On Error Resume Next  ' 2005/08/03 ADD T.SUZUKI
    If Me.cmb_コンテナ番号.Enabled = True Then
        Me.cmb_コンテナ番号.SetFocus
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_ヤードコード_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_ヤードコード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_ヤードコード, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_コンテナ番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_コンテナ番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_コンテナ番号, "")

    If vCode = "" Then Exit Sub

    ' 入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 0

        Case 2
            ' 桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "コードを入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

    ' ファイルへの存在チェック
    If CHK_CNTA_MAST(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    ' 契約ファイルの存在チェック
' 2006/03/04 ADD T.SUZUKI Start
    If CHK_CARG_FILE(0) = False Then
        MsgBox "契約されていません", vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_コンテナ番号_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_コンテナ番号_AfterUpdate()

    Dim vCode   As Variant
    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_コンテナ番号, "")

    Call CLR_RTN2(2)                                                            'INSERT 20050916 K.ISHIZAKA
    If vCode = "" Then
        ' 画面項目のロック処理                                                  'INSERT 20050916 K.ISHIZAKA
        Call subLock_Main                                                       'INSERT 20050916 K.ISHIZAKA
        Call subLock_Etc                                                        'INSERT 20050916 K.ISHIZAKA
        Exit Sub
    End If

'    Me.cmb_コンテナ番号 = Format(Me.cmb_コンテナ番号, "000000")                'DELETE 20050916 K.ISHIZAKA
    Me.cmb_コンテナ番号 = Format(vCode, "000000")                               'INSERT 20050916 K.ISHIZAKA

    ' ファイルへの存在チェック（コンテナマスタ）
    If CHK_CNTA_MAST(1) = False Then
       GoTo Exit_rtn
    End If

' 2006/03/04 ADD T.SUZUKI Start
    If CHK_CARG_FILE(1) = False Then
        MsgBox "契約されていません", vbInformation, strMsg_err
        Me.cmb_コンテナ番号.SetFocus
        GoTo Exit_rtn
    End If
' 2006/03/04 ADD T.SUZUKI End
' 2006/03/04 DEL T.SUZUKI Start
'    ' 契約ファイルの存在チェック
'    Select Case Me.opg_処理選択.Value
'        Case 1  ' *** 新規 ***
'            ' コンテナ契約ファイルチェック
'            If CHK_CARG_FILE(1) = False Then
'                ' → コンテナ契約ファイルが存在しない場合
'                If CHK_CNTA_MAST(2) = False Then
'                    GoTo Exit_rtn
'                End If
'            Else
'                If pvarContna = 0 Then
'                    ' コンテナ契約ファイルチェック結果、既存契約データのコンテナ状態が「0」のとき
'                    MsgBox "コンテナは契約可能状態でありません", vbInformation, strMsg_err
'                    Me.cmb_コンテナ番号.SetFocus
'                Else
'                    MsgBox "既に契約されています", vbInformation, strMsg_err
'                    Me.cmb_コンテナ番号.SetFocus
'                End If
'                GoTo Exit_rtn
'            End If
'
'            If (pintOpenMode <> 1 And pintOpenMode <> 9) Then  ' 2006/01/26 ADD T.SUZUKI
''            If pintOpenMode <> 1 Then                         ' 2006/01/26 DEL T.SUZUKI
'                ' 顧客区分名取得
'                strSql = "SELECT NZ(COUNT(*),0) " & Chr(13)
'                strSql = strSql & "  FROM dbo_YOUK_TRAN " & Chr(13)
'                strSql = strSql & " WHERE YOUKT_YCODE = " & Me.cmb_ヤードコード & Chr(13)  ' ヤードコード
'                strSql = strSql & "   AND YOUKT_NO    = " & Me.cmb_コンテナ番号 & Chr(13)  ' コンテナ番号
'                strSql = strSql & "   AND YOUKT_YUKBN IN (10, 2) "                         ' 予約受付状態区分
'
'                Set objRs = pobjDB.OpenRecordSet(strSql, dbOpenDynaset)
'
'                If objRs.Fields(0).Value <> 0 Then
'                    MsgBox "既に予約取置または契約受付されています。", vbInformation, strMsg_err
'                    Me.cmb_コンテナ番号.SetFocus
'                    GoTo Exit_rtn
'                Else
'                    objRs.Close
'                    Set objRs = Nothing
'                End If
'            End If
'
'        Case 2  ' *** 修正 ***
'            If CHK_CARG_FILE(1) = False Then
'                MsgBox "契約されていません", vbInformation, strMsg_err
'                Me.cmb_コンテナ番号.SetFocus
'                GoTo Exit_rtn
'            End If
'
'        Case 3  ' *** 削除 ***
'            If CHK_CARG_FILE(1) = False Then
'                MsgBox "契約されていません", vbInformation, strMsg_err
'                Me.cmb_コンテナ番号.SetFocus
'                GoTo Exit_rtn
'            End If
'
'            If CHK_REQU_FILE(0) = False Then
'            Else
'                Me.cmb_コンテナ番号.SetFocus
'                GoTo Exit_rtn
'            End If
'    End Select
' 2006/03/04 DEL T.SUZUKI End

    ' 画面の各コントロールの制御
    Call subLock_formItem
' 2023/05/02 M.HODNA
' 2019/04/23 ADD TAKENOUCHI Start
'    If fncChkExistEnterTrn = False Then
'        GoTo err_rtn
'    End If
' 2019/04/23 ADD TAKENOUCHI End
' 2023/05/02 M.HODNA

    If Nz(Me.cmb_ヤードコード, "") <> "" And Nz(vCode, "") <> "" Then
        Me.txt_FaceAuth.VALUE = GetFaceAuthStatus(CStr(Me.cmb_ヤードコード), CStr(vCode))
    Else
        Me.txt_FaceAuth.VALUE = ""
    End If


' 2006/03/04 ADD T.SUZUKI Start
    If pintOpenMode = 3 Then
        Me.txt_解約日付.SetFocus
    Else
        If pvarKYDATE < DATE Then
            Me.txt_解約日付.SetFocus
        Else
            '2006/09/22 SHIBAZAKI 予約名／予約番号削除
            'Me.txt_予約名.SetFocus  ' 2006/02/18 Add TAJIMA
        End If
    End If
' 2006/03/04 ADD T.SUZUKI End
' 2006/03/04 DEL T.SUZUKI Start
'    Select Case Me.opg_処理選択.Value
'        Case 3  ' *** 削除 ***
'            Me.cmd_削除.SetFocus
'
'        Case 2  ' *** 修正 ***
'' 2005/08/03 REP T.SUZUKI Start
'            If pintOpenMode = 3 Then
'                Me.txt_解約日付.SetFocus
'            Else
'                If pvarKYDATE < DATE Then
'                    Me.txt_解約日付.SetFocus
'                Else
'                    Me.txt_予約名.SetFocus  ' 2006/02/18 Add TAJIMA
'                End If
'            End If
'' 2005/08/03 REP T.SUZUKI End
'
'        Case 1  ' *** 新規 ***
'            Me.txt_予約名.SetFocus  'txt_契約開始日  2006/02/18 Add TAJIMA
'    End Select
' 2006/03/04 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_コンテナ番号_Enter
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_コンテナ番号_Enter()

    If IsNull(Me.cmb_ヤードコード) Then
        Me.cmb_ヤードコード.SetFocus
        Exit Sub
    End If

' DEL 2006/03/03 Tajima Strat
'    If Val(Me.cmb_ヤードコード) = 0 Then
'        If Me.chk_予約受付.Enabled = True Then
'            Me.chk_予約受付.SetFocus
'        Else
'            Me.cmb_ヤードコード.SetFocus
'        End If
'    End If
' DEL 2006/03/03 Tajima End
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_コンテナ番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_コンテナ番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_コンテナ番号, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約開始日_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約開始日_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_契約開始日, "")

    If vCode <> "" Then
        Me.txt_契約開始日 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約開始日_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約開始日_BeforeUpdate(Cancel As Integer)

'    Dim vCode As Variant
    Dim vDATE As Variant

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

' 2006/01/05 ADD T.SUZUKI Start
    ' 契約開始日
    Call dateCheck(Me.txt_契約開始日, Cancel): If Cancel = True Then GoTo Exit_rtn

    ' 初回契約日
'    Call DateCheck(Me.txt_初回契約日, Cancel): If Cancel = True Then GoTo Exit_rtn
'    Me.txt_初回契約日 = Format(varChangeDate, "yyyy/mm/dd")
    vDATE = Nz(Me.txt_初回契約日, "")

    If vDATE = "" Then
        MsgBox "初回契約日を入力してください。", vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If
' 2006/01/05 ADD T.SUZUKI End
' 2006/01/31 ADD H.Tajima Start
  '新規モードのときのみ　契約開始日＞前回契約の解約日付　のチェックをする
'  If Me.opg_処理選択.Value = 1 Then                                                 ' 2006/03/04 DEL T.SUZUKI
    '常にチェックを行う
    If IsNull(LimitationStartDate) = True Then
    ' 契約開始制限日が空ならば取得を行う※キャッシュみたいなもん
      If fncGetLimitationStartDate = False Then
        Cancel = True    ' Falseが返ってきたらシステム的なエラー
        SendKeys "{ESC}"
        GoTo Exit_rtn
      End If
      '↓デバッグ用、ブレークポイント張ると挙動不審なので
      'Call MSZZ003_M00("FVS600", 0, "LimitationStartDate = " & LimitationStartDate)
    End If
    If LimitationStartDate <> "" Then
      '↓varChangeDateにMe.txt_契約開始日を変換した日付が入っているのでそれで評価
      If CStr(VarChangeDate) <= CStr(LimitationStartDate) Then
          MsgBox "前回契約の完了日 " & LimitationStartDate & "以降の日付で入力してください。", vbInformation, strMsg_err
          Cancel = True
          SendKeys "{ESC}"
          GoTo Exit_rtn
      End If
    End If
'  End If                                                                            ' 2006/03/04 DEL T.SUZUKI
' 2006/01/31 ADD T.Tajima End

' 2006/01/05 DEL T.SUZUKI Start
'    vCode = Nz(Me.txt_契約開始日, "")
'
'    ' 何も入力されていない時は、処理をしない
'    If vCode = "" Then GoTo Exit_rtn
'
'    vCode = Check_Date(vCode)
'    vCode = CMCheck_Date(vCode)
'
'    ' 入力チェック
'    Select Case CmBCheck_Date(vCode)
'        Case 0
'
'        Case 2
'            ' 桁数エラー
'            MsgBox strMSGKetaOver, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'
'        Case 3, 9
'            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'    End Select
'
'    ' 初回契約日チェック
'    vDATE = Nz(Check_Date(Me.txt_初回契約日), "")
'
'    If vDATE <> "" Then
'    Else
'        MsgBox "初回契約日を入力してください。", vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
' 2006/01/05 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約開始日_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約開始日_AfterUpdate()

'    ' 変更フラグ                                            ' 2005/08/03 DEL T.SUZUKI
'    pblnChangeFlag_FVS600 = True                            ' 2005/08/03 DEL T.SUZUKI

    Me.txt_契約開始日 = Format(VarChangeDate, P_FORMAT_YMD)  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_契約開始日 = Check_Date(Me.txt_契約開始日)       ' 2006/01/05 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約開始日_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約開始日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_契約開始日, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約開始日_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約開始日_LostFocus()

    Dim vCode     As Variant
    Dim intkikan  As Long

    vCode = Nz(Me.txt_契約開始日, "")
  
    If vCode <> "" Then
        Me.txt_契約開始日 = Format(vCode, P_FORMAT_YMD)                                  ' 2006/01/05 ADD T.SUZUKI
'        Me.txt_契約開始日 = Format(vCode, "yyyy/mm/dd")                                 ' 2006/01/05 DEL T.SUZUKI
    Else
        Exit Sub
    End If

'    If Me.opg_処理選択.Value = 1 Then                                                   ' 2006/03/04 DEL T.SUZUKI
        intkikan = Nz(DLookup("CONT_ARGEYY", "dbo_CONT_MAST", "CONT_KEY = 1"), "0")
        intkikan = intkikan * 12 - 1
        Me.txt_契約満了日 = Format(End_Day(Me.txt_契約開始日, intkikan), P_FORMAT_YMD)   ' 2006/01/05 ADD T.SUZUKI
'        Me.txt_契約満了日 = Format(End_Day(Me.txt_契約開始日, intkikan), "YYYY/MM/DD")  ' 2006/01/05 DEL T.SUZUKI
'    End If                                                                              ' 2006/03/04 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約開始日_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約開始日_Change()
'    MsgBox "契約開始日が変更されました。"
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約満了日_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約満了日_AfterUpdate()

'    ' 変更フラグ                                            ' 2005/08/03 DEL T.SUZUKI
'    pblnChangeFlag_FVS600 = True                            ' 2005/08/03 DEL T.SUZUKI

    Me.txt_契約満了日 = Format(VarChangeDate, P_FORMAT_YMD)  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_契約満了日 = Check_Date(Me.txt_契約満了日)       ' 2006/01/05 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約満了日_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約満了日_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant
    Dim vDATE As Variant

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

' 2006/01/05 ADD T.SUZUKI Start
    ' 契約満了日
    Call dateCheck(Me.txt_契約満了日, Cancel): If Cancel = True Then GoTo Exit_rtn
    vCode = Me.txt_契約満了日

    ' 契約開始日
'    Call DateCheck(Me.txt_契約開始日, Cancel): If Cancel = True Then GoTo Exit_rtn
'    Me.txt_契約開始日 = Format(varChangeDate, "yyyy/mm/dd")
    vDATE = Nz(Me.txt_契約開始日, "")

    If vDATE <> "" Then
        If vCode < vDATE Then
            MsgBox "契約満了日は、契約開始日以降を入力してください。", vbInformation + vbOKOnly, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
        End If
    Else
        MsgBox "契約開始日を入力してください。", vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    ' 末日チェック
    If vCode <> Format(Matu_Day(VarChangeDate), P_FORMAT_YMD) Then  ' 2006/03/04 ADD T.SUZUKI
' 2006/03/04 DEL T.SUZUKI Start
'    If vCode <> Format(Matu_Day(vCode), P_FORMAT_YMD) Then         ' 2006/01/05 ADD T.SUZUKI
''    If vCode <> Format(Matu_Day(vCode), "yyyy/mm/dd") Then        ' 2006/01/05 DEL T.SUZUKI
' 2006/03/04 DEL T.SUZUKI End

        MsgBox "契約満了日は、末日の日付を入力してください。", vbInformation + vbOKOnly, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If
' 2006/01/05 ADD T.SUZUKI End
' 2006/01/05 DEL T.SUZUKI Start
'    Dim vCode As Variant
'    Dim vDATE As Variant
'
'    On Error GoTo err_rtn
'
'    If pblnOpenFlag_FVS600 = False Then Exit Sub
'
'    vCode = Nz(Me.txt_契約満了日, "")
'
'    ' 何も入力されていない時は、処理をしない
'    If vCode = "" Then GoTo Exit_rtn
'
'    vCode = Check_Date(vCode)
'    vCode = CMCheck_Date(vCode)
'
'    ' 入力チェック
'    Select Case CmBCheck_Date(vCode)
'        Case 0
'
'        Case 2
'            ' 桁数エラー
'            MsgBox strMSGKetaOver, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'
'        Case 3, 9
'            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'    End Select
'
'    ' 契約開始日チェック
'    vDATE = Nz(Check_Date(Me.txt_契約開始日), "")
'
'    If vDATE <> "" Then
'        If vCode < vDATE Then
'            MsgBox "契約満了日は、契約開始日以降を入力してください。", vbInformation + vbOKOnly, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    Else
'        MsgBox "契約開始日を入力してください。", vbInformation, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
'
'    ' 末日チェック
'    If vCode <> Format(Matu_Day(vCode), "yyyy/mm/dd") Then
'        MsgBox "契約満了日は、末日の日付を入力してください。", vbInformation + vbOKOnly, strMsg_err
'        Cancel = True
'        SendKeys "{ESC}"
'        GoTo Exit_rtn
'    End If
' 2006/01/05 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約満了日_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約満了日_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_契約満了日, "")

    If vCode <> "" Then
        Me.txt_契約満了日 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約満了日_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約満了日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_契約満了日, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約満了日_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約満了日_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_契約満了日, "")

    If vCode <> "" Then
        Me.txt_契約満了日 = Format(vCode, P_FORMAT_YMD)      ' 2006/01/05 ADD T.SUZUKI
'        Me.txt_契約満了日 = Format(vCode, "yyyy/mm/dd")     ' 2006/01/05 DEL T.SUZUKI
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約満了日_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約満了日_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約日付_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日付_BeforeUpdate(Cancel As Integer)

    Dim vDATE        As Variant
    Dim vCompareDate As Variant

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

' 2006/01/05 ADD T.SUZUKI Start
    ' 解約日付
    Call dateCheck(Me.txt_解約日付, Cancel): If Cancel = True Then GoTo Exit_rtn
' 2006/01/05 ADD T.SUZUKI End

' 2005/08/03 ADD T.SUZUKI Start
    ' 解約修正のとき
'    If Nz(Me.txt_解約日付.Tag) <> "" Then                           ' 2006/03/04 ADD T.SUZUKI
''    If Me.opg_処理選択 = 2 And Nz(Me.txt_解約日付.Tag) <> "" Then  ' 2006/03/04 DEL T.SUZUKI
'        vCompareDate = Me.txt_解約日付.Tag
'
'        If Mid(Replace(Nz(Me.txt_解約日付, ""), "/", ""), 1, 6) <> Mid(Replace(vCompareDate, "/", ""), 1, 6) Then
''        If Format(Me.txt_解約日付, "YYYYMM") <> Format(vCompareDate, "YYYYMM") Then
'            MsgBox "解約日付は日のみ変更して下さい。", vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
' 2005/08/03 ADD T.SUZUKI End

'2006/09/22 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↓
'    If txt_ヤード解約予定 <> "" Then
'        If DateDiff("d", txt_ヤード解約予定, VarChangeDate) > 0 Then
'            MsgBox "解約日はヤード解約予定日 [" & txt_ヤード解約予定 & "] 以前で無ければいけません。", vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'        End If
'    End If
'2006/09/22 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↑


' 2006/03/12 ADD Tajima Start
  ' キャンペーン適用満了日とのチェック
  If Nz(Me.txt_キャンペーン適用満了日, "") <> "" Then
        '↓varChangeDateにMe.txt_解約日を変換した日付が入っているのでそれで評価
      If CStr(VarChangeDate) <= CStr(Me.txt_キャンペーン適用満了日) Then
        Dim ret As Integer
        ret = MsgBox(pstrCampaignWarningText, vbYesNo + vbExclamation + vbDefaultButton2, "この契約で適用したキャンペーンについて")
        If ret = vbNo Then
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
        End If
      End If
  End If
' 2006/03/12 ADD Tajima End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約日付_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日付_AfterUpdate()

    Dim vDATE As Variant

'    ' 変更フラグ                                          ' 2005/08/03 DEL T.SUZUKI
'    pblnChangeFlag_FVS600 = True                          ' 2005/08/03 DEL T.SUZUKI

    Me.txt_解約日付 = Format(VarChangeDate, P_FORMAT_YMD)  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_解約日付 = Check_Date(Me.txt_解約日付)         ' 2006/01/05 DEL T.SUZUKI

    vDATE = Nz(Me.txt_解約日付, "")

    ' 何も入力されていない場合、解約関連入力項目を使用不可に設定
    ' 解約日付に入力されている場合、解約関連入力項目を使用可に設定
    If vDATE = "" Then
        ' 未入力の場合
        Me.txt_解約受付日 = Null
        Me.cmb_解約区分 = Null
        Me.txt_解約区分 = Null
        Me.cmb_解約理由 = Null
        Me.txt_解約理由 = ""
        Me.cmb_解約受付担当者コード = Null
        Me.txt_解約受付担当者名 = Null

        Call CmLockControl(Me.txt_解約受付日, 1)
        'Call CmLockControl(Me.cmb_解約区分, 1) 2008/03/23 del tajima
        Call CmLockControl(Me.cmb_解約理由, 1)
        Call CmLockControl(Me.cmb_解約受付担当者コード, 1)
'        Me.cmd_解約入力.Enabled = False

    Else
        If Me.txt_解約日付.Tag = "" Then
            ' 入力されている場合
            ' 解約受付日・解約区分・解約理由に初期値設定
            ' ※タグに解約日付が設定されていない場合(解約日付が新規に入力されたとき)はデフォルト設定を行う
            Me.txt_解約受付日 = Format(Now, P_FORMAT_YMD)      ' 2006/01/05 ADD T.SUZUKI
'            Me.txt_解約受付日 = Format(Now, "YYYY/MM/DD")     ' 2006/01/05 DEL T.SUZUKI
            Me.cmb_解約区分 = "01"                         ' デフォルト値：01
            Call cmb_解約区分_AfterUpdate
            Me.cmb_解約理由 = Null
            Call cmb_解約理由_AfterUpdate

            Me.txt_解約日付.Tag = Me.txt_解約日付
        End If

        Call CmLockControl(Me.txt_解約受付日, 0)
        'Call CmLockControl(Me.cmb_解約区分, 0) 2008/03/23 del tajima
        Call CmLockControl(Me.cmb_解約理由, 0)
        Call CmLockControl(Me.cmb_解約受付担当者コード, 0)
'        Me.cmd_解約入力.Enabled = True
    End If

'    If Me.opg_処理選択.Value = 2 Then                         ' 2006/03/04 DEL T.SUZUKI
        Me.cmd_解約入力.Enabled = True
'    End If                                                    ' 2006/03/04 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約日付_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日付_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_解約日付, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約日付_LostFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日付_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_解約日付, "")

    If vCode <> "" Then
        Me.txt_解約日付 = Format(vCode, P_FORMAT_YMD)      ' 2006/01/05 ADD T.SUZUKI
'        Me.txt_解約日付 = Format(vCode, "yyyy/mm/dd")     ' 2006/01/05 DEL T.SUZUKI
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約日付_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日付_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_解約日付, "")

    If vCode <> "" Then
        Me.txt_解約日付 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約日付_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約日付_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約受付日_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約受付日_AfterUpdate()

    Dim vCode As Variant

'    '変更フラグ                                             ' 2005/08/03 DEL T.SUZUKI
'    pblnChangeFlag_FVS600 = True                            ' 2005/08/03 DEL T.SUZUKI

    Me.txt_解約受付日 = Format(VarChangeDate, P_FORMAT_YMD)  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_解約受付日 = Check_Date(Me.txt_解約受付日)       ' 2006/01/05 DEL T.SUZUKI
'    vCode = Nz(Me.txt_解約受付日, "")                       ' 2006/01/05 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約受付日_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約受付日_BeforeUpdate(Cancel As Integer)

'    Dim vCode  As Variant
'    Dim vDATE  As Variant

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

' 2006/01/05 ADD T.SUZUKI Start
    ' 解約受付日
    Call dateCheck(Me.txt_解約受付日, Cancel): If Cancel = True Then GoTo Exit_rtn
' 2006/01/05 ADD T.SUZUKI End

' 2006/01/05 DEL T.SUZUKI Start
'    vCode = Nz(Me.txt_解約受付日, "")
'
'    ' 何も入力されていない時は、処理をしない
'    If vCode = "" Then GoTo Exit_rtn
'
'    vCode = Check_Date(vCode)
'    vCode = CMCheck_Date(vCode)
'
'    ' 入力チェック
'    Select Case CmBCheck_Date(vCode)
'        Case 0
'' 2005/08/03 DEL T.SUZUKI Start
'            'txt_解約日付 が入力されれば chk_コンテナ状態 チェックOK
''            Call CmLockControl(Me.chk_コンテナ状態, 0)
'' 2005/08/03 DEL T.SUZUKI End
'
'        Case 2
'            ' 桁数エラー
'            MsgBox strMSGKetaOver, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'
'        Case 3, 9
'            MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            GoTo Exit_rtn
'    End Select
' 2006/01/05 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約受付日_GotFocus
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約受付日_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_解約受付日, "")

    If vCode <> "" Then
        Me.txt_解約受付日 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約受付日_LostFocus
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約受付日_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_解約受付日, "")

    If vCode <> "" Then
        Me.txt_解約受付日 = Format(vCode, P_FORMAT_YMD)      ' 2006/01/05 ADD T.SUZUKI
'        Me.txt_解約受付日 = Format(vCode, "yyyy/mm/dd")     ' 2006/01/05 DEL T.SUZUKI
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解約受付日_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解約受付日_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約区分_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約区分_AfterUpdate()

    Dim vCode       As Variant
    Dim strName     As String

    vCode = Nz(Me.cmb_解約区分, "")
    If vCode = "" Then
        Me.txt_解約区分 = ""
        Exit Sub
    End If
    Me.cmb_解約区分 = Format(vCode, "00")

    strName = fncSelectName("089", Me.cmb_解約区分)           ' 2006/01/05 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_解約区分, "089")          ' 2006/01/05 DEL T.SUZUKI

    Me.txt_解約区分 = strName

    '▼2008/03/23 add tajima
    '  未使用キャンセル且つ契約開始日設定済み、解約日付未入力ならば
    '  解約日を契約開始日前月末日とする
    Dim strTargetDate  As String
    If Me.cmb_解約区分.VALUE = "03" And _
        Nz(Me.txt_契約開始日.VALUE, "") <> "" And _
        Nz(Me.txt_解約日付.VALUE, "") = "" Then
        
        strTargetDate = Me.txt_契約開始日.VALUE
        strTargetDate = DateSerial(Year(strTargetDate), _
                                   Month(strTargetDate), _
                                    1) - 1
        Me.txt_解約日付 = Format(strTargetDate, "YYYY/MM/DD")
        '解約受付日の初期設定
        If Nz(Me.txt_解約受付日, "") = "" Then
            Me.txt_解約受付日 = Format(Now, P_FORMAT_YMD)
        End If
    End If
    '他入力可状態にチェンジ
    Call CmLockControl(Me.txt_解約受付日, 0)
    Call CmLockControl(Me.cmb_解約理由, 0)
    Call CmLockControl(Me.cmb_解約受付担当者コード, 0)
    '▲2008/03/23 add tajima

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約区分_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約区分_BeforeUpdate(Cancel As Integer)
    
    Dim vCode       As Variant
    Dim strName     As String

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

    ' コードが空の場合解約区分欄も空欄表示
    If fncControlCrear(Me.cmb_解約区分, Me.txt_解約区分) = False Then Exit Sub  ' 2005/08/03 ADD T.SUZUKI

    '  解約区分名取得
    strName = fncSelectName("089", Me.cmb_解約区分)                             ' 2006/01/05 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_解約区分, "089")                            ' 2006/01/05 DEL T.SUZUKI
    If strName = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約区分_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約区分_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約区分_KeyPress
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_解約区分, KeyAscii, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約理由_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
Private Sub cmb_解約理由_AfterUpdate()

    Dim vCode       As Variant
    Dim strName     As String

    vCode = Nz(Me.cmb_解約理由, "")
    If vCode = "" Then
        Me.txt_解約理由 = ""
        Exit Sub
    End If
    Me.cmb_解約理由 = Format(vCode, "00")

    strName = fncSelectName("092", Me.cmb_解約理由)           ' 2006/01/05 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_解約理由, "092")          ' 2006/01/05 DEL T.SUZUKI

    Me.txt_解約理由 = strName
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約理由_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
Private Sub cmb_解約理由_BeforeUpdate(Cancel As Integer)

    Dim strName     As String

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

    ' コードが空の場合解約区分欄も空欄表示
    If fncControlCrear(Me.cmb_解約区分, Me.txt_解約理由) = False Then Exit Sub

    '  解約区分名取得
    strName = fncSelectName("092", Me.cmb_解約理由)           ' 2006/01/08 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_解約理由, "092")          ' 2006/01/08 DEL T.SUZUKI
    If strName = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約理由_KeyPress
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
Private Sub cmb_解約理由_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_解約理由, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約理由_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約理由_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約受付担当者コード_AfterUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約受付担当者コード_AfterUpdate()

    Dim vCode   As Variant  ' 受付担当者コード
    Dim vTanto  As Variant  ' 受付担当者名

    On Error GoTo cmb_解約受付担当者コード_AfterUpdate_Err

    vCode = Nz(Me.cmb_解約受付担当者コード, "")
    If vCode = "" Then GoTo cmb_解約受付担当者コード_AfterUpdate_End

    ' ファイルへの存在チェック
    If GetTantoData(vCode, vTanto) = False Then GoTo cmb_解約受付担当者コード_AfterUpdate_End

    Me.txt_解約受付担当者名 = vTanto
    Me.cmb_解約受付担当者コード = Format$(Me.cmb_解約受付担当者コード, "000")

    GoTo cmb_解約受付担当者コード_AfterUpdate_End

cmb_解約受付担当者コード_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_解約受付担当者コード_AfterUpdate_End:
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約受付担当者コード_BeforeUpdate
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約受付担当者コード_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    If pblnOpenFlag_FVS600 = False Then Exit Sub

    ' コードが空の場合担当者名欄も空欄表示
    If fncControlCrear(Me.cmb_解約受付担当者コード, Me.txt_解約受付担当者名) = False Then Exit Sub  ' 2005/08/03 ADD T.SUZUKI

' 2005/08/03 DEL T.SUZUKI Start
'    vCode = Nz(Me.cmb_解約受付担当者コード, "")
'    If vCode = "" Then
'        Exit Sub
'    End If
' 2005/08/03 DEL T.SUZUKI End

    ' 選択された担当者データの存在チェック
    If GetTantoData(Nz(Me.cmb_解約受付担当者コード, "")) = False Then
        MsgBox "登録されていません。", vbInformation, strMsg_kaku
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約受付担当者コード_KeyPress
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約受付担当者コード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_解約受付担当者コード, KeyAscii, 3)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_解約受付担当者コード_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_解約受付担当者コード_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_解約入力_Click
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_解約入力_Click()

' 2006/03/04 ADD T.SUZUKI Start
    If Nz(Me.txt_受注契約番号, "") = "" Then Exit Sub
' 2006/03/04 ADD T.SUZUKI End

    If pblnChangeFlag_FVS600 = True Then
        ' 変更フラグがTrueの時は確認メッセージを出す
        pintAns = MsgBox("更新しますか？" & vbCrLf & _
                  "※更新しなければ解約入力は表示できません。", _
                  vbQuestion + vbYesNo, strMsg_kaku)

        If pintAns = vbYes Then
            ' 更新前の入力チェックを行う
            If CHK_RTN_For_UpDate = False Then
                Exit Sub
            End If

            ' データ登録処理を行う
            If WRI_RTN(P_SHU) = False Then Exit Sub                             ' 2006/03/04 ADD T.SUZUKI
'            If WRI_RTN = False Then Exit Sub                                   ' 2006/03/04 DEL T.SUZUKI
        Else
            Exit Sub
        End If

        pblnChangeFlag_FVS600 = False  ' 変更フラグ(OFF)
        Me.Caption = pstrFormName      ' 標題設定
    End If

    Select Case Me.Form.NAME
        Case pstrFVS600, pstrFVS602
            ' 画面コントロールの無効化
            Call subFormLockControl(False)

            ' 解約入力画面
            doCmd.OpenForm pstrFVS700, , , , , , Me.txt_受注契約番号 & "," & pstrFVS600

        Case pstrFVS601
            ' 解約入力画面(参照モード)
            doCmd.OpenForm pstrFVS701, , , , , acDialog, Me.txt_受注契約番号 & "," & pstrFVS601
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_鍵区分_AfterUpdate
'        CREATE_DATE      :2006/05/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵区分_AfterUpdate()

    Dim strRyak     As String
    Dim vCode       As Variant

    vCode = Nz(Me.cmb_鍵区分, "")
    If vCode = "" Then
        Me.txt_鍵区分 = ""
        Me.txt_スペアキー有無 = ""      'INSERT 2014/06/24 MIYAMOTO
        Exit Sub
    End If

    strRyak = fncSelectRyak("060", Me.cmb_鍵区分)           ' 2006/01/05 ADD T.SUZUKI

    Me.txt_鍵区分 = strRyak

'↓ INSERT 2014/06/24 MIYAMOTO
'↓ UPDATE 2016/05/19 MIYAMOTO
'    If vCode = 15 Then
    If MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(2)) <> 0 Then
'↑ UPDATE 2016/05/19 MIYAMOTO
        Me.txt_スペアキー有無 = "無"
    Else
        Me.txt_スペアキー有無 = "有"
    End If
'↑ INSERT 2014/06/24 MIYAMOTO

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_鍵区分_BeforeUpdate
'        CREATE_DATE      :2006/05/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵区分_BeforeUpdate(Cancel As Integer)
    
    Dim vCode       As Variant
    Dim strRyak     As String

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

    vCode = Nz(Me.cmb_鍵区分, "")
    If vCode = "" Then
        Cancel = True
        MsgBox "鍵区分" & strMSGInputError, vbInformation, strMsg_kaku
        Me.txt_鍵区分 = ""
        Me.txt_スペアキー有無 = ""      'INSERT 2014/06/24 MIYAMOTO
        Exit Sub
    End If

    ' 鍵区分名取得
    strRyak = fncSelectRyak("060", Me.cmb_鍵区分)                            ' 2006/01/05 ADD T.SUZUKI
    If strRyak = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_鍵区分_KeyPress
'        CREATE_DATE      :2006/05/10
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_鍵区分, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_鍵区分_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_鍵区分_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :解除番号
'        MODULE_ID        :txt_解除番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解除番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    vCode = Nz(Me.txt_解除番号, "")
  
    '入力チェック
    Select Case CmBCheck_Char(vCode, 10)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "鍵番号を入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_解除番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_解除番号_Change()
    Call subSetChangeFlag
End Sub

'2005/09/22 SHIBAZAKI 予約名／予約番号削除↓
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_予約名_BeforeUpdate
''        CREATE_DATE      :2006/01/05
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_予約名_BeforeUpdate(Cancel As Integer)
'
'    Dim vCode As Variant
'
'    On Error GoTo txt_予約名_BeforeUpdate_Err
'
'    vCode = Nz(Me.txt_予約名, "")
'
'    If vCode = "" Then Exit Sub
'
'    '入力チェック
'    Select Case CmBCheck_Char(vCode, 36)
'        Case 0
'
'        Case 2
'            '桁数ｴﾗｰ
'            MsgBox strMSGKetaOver, vbInformation, strMsg_err
'            Cancel = True
'            Exit Sub
'
'        Case 9
'            MsgBox "予約名" & strMSGInputError, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            Exit Sub
'    End Select
'
'    Exit Sub
'
'txt_予約名_BeforeUpdate_Err:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_予約名_Change
''        CREATE_DATE      :2005/08/03
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_予約名_Change()
'    Call subSetChangeFlag
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_予約名_KeyPress
''        CREATE_DATE      :2005/01/05
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_予約名_KeyPress(KeyAscii As Integer)
'    Call CmKeyPress_For_Char2(Me.txt_予約名, KeyAscii, 36)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_予約番号_BeforeUpdate
''        CREATE_DATE      :2006/01/05
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_予約番号_BeforeUpdate(Cancel As Integer)
'
'    Dim vCode As Variant
'
'    On Error GoTo txt_予約番号_BeforeUpdate_Err
'
'    vCode = Nz(Me.txt_予約番号, "")
'
'    If vCode = "" Then Exit Sub
'
'    '入力チェック
'    Select Case CmBCheck_Char(vCode, 15)
'        Case 0
'
'        Case 2
'            '桁数ｴﾗｰ
'            MsgBox strMSGKetaOver, vbInformation, strMsg_err
'            Cancel = True
'            Exit Sub
'
'        Case 9
'            MsgBox "予約番号" & strMSGInputError, vbInformation, strMsg_err
'            Cancel = True
'            SendKeys "{ESC}"
'            Exit Sub
'    End Select
'
'    Exit Sub
'
'txt_予約番号_BeforeUpdate_Err:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_予約番号_Change
''        CREATE_DATE      :2005/08/03
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_予約番号_Change()
'    Call subSetChangeFlag
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_予約番号_KeyPress
''        CREATE_DATE      :2006/01/05
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_予約番号_KeyPress(KeyAscii As Integer)
'    Call CmKeyPress_For_Char2(Me.txt_予約番号, KeyAscii, 15)
'End Sub
'2005/09/22 SHIBAZAKI 予約名／予約番号削除↑

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_保証区分_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_保証区分, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_保証区分_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_AfterUpdate()

    Dim strRyak     As String
    Dim vCode       As Variant

    vCode = Nz(Me.cmb_保証区分, "")

    If vCode = "" Then
        Me.txt_保証区分 = ""
        Exit Sub
    End If

    strRyak = fncSelectRyak("200", Me.cmb_保証区分)           ' 2006/01/05 ADD T.SUZUKI
'    strRyak = fncSelectRyak(Me.cmb_保証区分, "200")          ' 2006/01/05 DEL T.SUZUKI

    Me.txt_保証区分 = strRyak

    'INSERT 2008/11/17 SHIBAZAKI
    Me.lbl_預り保証金 = cmb_保証区分.Column(3, cmb_保証区分.ListIndex)
    
    '↓DELETE 2008/11/17 SHIBAZAKI
''     2006/03/03 Fix Tajima Start
'    Select Case Me.cmb_保証区分
'      Case "3", "4", "6"
'        Me.lbl_預り保証金 = "保証委託料"
'      Case "7"                              'INSERT 2008/09/25 SHIBAZAKI
'        Me.lbl_預り保証金 = "事務手数料"    'INSERT 2008/09/25 SHIBAZAKI
'      Case Else
'        Me.lbl_預り保証金 = "保証金"
'    End Select
''     2006/03/03 Fix Tajima End
    '↑DELETE 2008/11/17 SHIBAZAKI

End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_保証区分_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_BeforeUpdate(Cancel As Integer)

    Dim vCode       As Variant
    Dim strRyak     As String

    On Error GoTo err_rtn

    If pblnOpenFlag_FVS600 = False Then Exit Sub

' 2005/10/07 DEL T.SUZUKI Start
'    If fncControlCrear(Me.cmb_保証区分, Me.txt_保証区分) = False Then Exit Sub  ' 2005/08/03 ADD T.SUZUKI
' 2005/10/07 DEL T.SUZUKI End
' 2005/10/07 ADD T.SUZUKI Start
    vCode = Nz(Me.cmb_保証区分, "")
    If vCode = "" Then
        Cancel = True
        MsgBox "保証区分" & strMSGInputError, vbInformation, strMsg_kaku
        Me.txt_保証区分 = ""
        Exit Sub
    End If
' 2005/10/07 ADD T.SUZUKI End

    ' 保証区分名取得
    strRyak = fncSelectRyak("200", Me.cmb_保証区分)                              ' 2006/01/05 ADD T.SUZUKI
'    strRyak = fncSelectRyak(Me.cmb_保証区分, "200")                             ' 2006/01/05 DEL T.SUZUKI
    If strRyak = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

' 2005/08/03 DEL T.SUZUKI Start
'    On Error GoTo err_rtn
'
'    If bOpenFlag_FVS600 = False Then
'        Exit Sub
'    End If
'
'    If Trim(Me.cmb_保証区分) = "" Then
'        Cancel = True
'        GoTo Exit_rtn
'    End If
' 2005/08/03 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_保証区分_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証区分_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :承認日
'        MODULE_ID        :txt_承認日_BeforeUpdate
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認日_BeforeUpdate(Cancel As Integer)

    On Error GoTo err_rtn
  
    If pblnOpenFlag_FVS600 = False Then Exit Sub

    ' 承認日
    Call dateCheck(Me.txt_承認日, Cancel): If Cancel = True Then GoTo Exit_rtn

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :承認日
'        MODULE_ID        :txt_承認日_AfterUpdate
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認日_AfterUpdate()
    Me.txt_承認日 = Format(VarChangeDate, P_FORMAT_YMD)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :承認日
'        MODULE_ID        :txt_承認日_GotFocus
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認日_GotFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_承認日, "")

    If vCode <> "" Then
        Me.txt_承認日 = Format(vCode, "yy/mm/dd")
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :承認日
'        MODULE_ID        :txt_承認日_LostFocus
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認日_LostFocus()

    Dim vCode As Variant

    vCode = Nz(Me.txt_承認日, "")
    
    If vCode <> "" Then
        Me.txt_承認日 = Format(vCode, P_FORMAT_YMD)      ' 2006/01/05 ADD T.SUZUKI
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :承認日
'        MODULE_ID        :txt_承認日_KeyPress
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_承認日, KeyAscii, 10)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_承認日_Change
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認日_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_承認番号_BeforeUpdate
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    vCode = Nz(Me.txt_承認番号, "")

    '入力チェック
    Select Case CmBCheck_Char(vCode, 20)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "承認番号を入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_承認番号_Change
'        CREATE_DATE      :2006/05/09
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_承認番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_月額料金_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_月額料金_BeforeUpdate(Cancel As Integer)

    Dim vNum As Variant

    On Error GoTo err_rtn

    vNum = Nz(Me.txt_月額料金, "")

    If vNum = "" Then Exit Sub

' 2006/01/05 ADD T.SUZUKI Start
    If fncTxtInputCheck(Me.txt_月額料金, pcsglMoneyMaxValue, 9) = False Then
        MsgBox strMSGKetaOver, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
' 2006/01/05 ADD T.SUZUKI End
' 2006/01/05 DEL T.SUZUKI Start
'    Select Case CHK_Numeric(vNum, 9)
'        Case 0
'
'        Case 1
'            MsgBox "月額料金は、数値で、入力してください。", vbInformation, strMsg_kaku
'            Cancel = True
'            SendKeys "{ESC}"  '入力のキャンセル処理
'            Exit Sub
'
'        Case 2
'            MsgBox "月額料金は、9桁以内で、入力してください。", vbInformation, strMsg_kaku
'            Cancel = True
'            SendKeys "{ESC}"  '入力のキャンセル処理
'            Exit Sub
'
'        Case Else
'            Cancel = True
'            SendKeys "{ESC}"
'            Exit Sub
'    End Select
' 2006/01/05 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_月額料金_GotFocus
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_月額料金_GotFocus()
    Me.txt_月額料金.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_月額料金_LostFocus
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_月額料金_LostFocus()
    Me.txt_月額料金.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_月額料金_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_月額料金_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_月額料金, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_月額料金_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_月額料金_AfterUpdate()

    Call subCalculation_TAX      ' 2005/08/03 REP T.SUZUKI
    Call subCalculation_TOTAL    ' 2005/08/03 REP T.SUZUKI
' 2006/03/03 DEL Tajima Start
'    Call subCalculation_SECUKG   ' 2005/08/03 ADD T.SUZUKI
'    If Me.opg_処理選択.Value = 1 Then
'        Me.txt_初回金額 = Me.txt_月額料金
'        Call txt_初回金額_AfterUpdate
'    End If
'
'    ' 初回使用料算出
'    Call subGetFirstMoney
' 2006/03/03 DEL Tajima End
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_月額料金_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_月額料金_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subCalculation_TAX
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCalculation_TAX()

    Dim varTAX As Variant

' 2005/08/03 REP T.SUZUKI Start
    varTAX = ""
    If pvarHenrit = "" Then
       varTAX = Me.txt_月額料金 * pvarMaerit / 100
    Else
       If Me.txt_契約開始日 < pvarDATE Then
          varTAX = Me.txt_月額料金 * pvarMaerit / 100
       Else
          varTAX = Me.txt_月額料金 * pvarAtorit / 100
       End If
    End If

    ' 小数点以下は切り捨て
    Me.txt_消費税 = Fix(varTAX)

'    If pvarHenrit = "" Then
'       Me.txt_消費税 = Me.txt_月額料金 * pvarMaerit / 100
'    Else
'       If Me.txt_契約開始日 < pvarDATE Then
'          Me.txt_消費税 = Me.txt_月額料金 * pvarMaerit / 100
'       Else
'          Me.txt_消費税 = Me.txt_月額料金 * pvarAtorit / 100
'       End If
'    End If
' 2005/08/03 REP T.SUZUKI End
End Sub

' 2005/08/03 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_消費税計算
''        CREATE_DATE      :2004/11/25
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_消費税計算()
'    If pvarHenrit = "" Then
'       Me.txt_消費税 = Me.txt_月額料金 * pvarMaerit / 100
'    Else
'       If Me.txt_契約開始日 < pvarDATE Then
'          Me.txt_消費税 = Me.txt_月額料金 * pvarMaerit / 100
'       Else
'          Me.txt_消費税 = Me.txt_月額料金 * pvarAtorit / 100
'       End If
'    End If
'End Sub
' 2005/08/03 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subCalculation_TOTAL
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCalculation_TOTAL()
    ' 月額料金合計算出 & セット
    Me.txt_合計 = Nz(Me.txt_月額料金, 0) + Nz(txt_消費税, 0)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subCalculation_FstTAX
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCalculation_FstTAX()

' 2006/03/03 DEL Tajima Start
'    Dim varTAX As Variant
'' 2005/08/03 REP T.SUZUKI Start
'    varTAX = ""  ' 初期化
'    If pvarHenrit = "" Then
'        varTAX = (Me.txt_初回金額 + Me.txt_初回鍵代) * pvarMaerit / 100
'    Else
'        If Me.txt_契約開始日 < pvarDATE Then
'           varTAX = (Me.txt_初回金額 + Me.txt_初回鍵代) * pvarMaerit / 100
'        Else
'           varTAX = (Me.txt_初回金額 + Me.txt_初回鍵代) * pvarAtorit / 100
'        End If
'    End If
'
'    ' 小数点以下は四捨五入する
'    Me.txt_初回消費税 = Fix(varTAX)
'
'' 2005/08/03 REP T.SUZUKI End
' 2006/03/03 DEL Tajima End

End Sub

' 2005/08/03 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :txt_初回消費税計算
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txt_初回消費税計算()
'
'    If pvarHenrit = "" Then
'        Me.txt_初回消費税 = (Me.txt_初回金額 + Me.txt_初回鍵代) * pvarMaerit / 100
'    Else
'        If Me.txt_契約開始日 < pvarDATE Then
'           Me.txt_初回消費税 = (Me.txt_初回金額 + Me.txt_初回鍵代) * pvarMaerit / 100
'        Else
'           Me.txt_初回消費税 = (Me.txt_初回金額 + Me.txt_初回鍵代) * pvarAtorit / 100
'        End If
'    End If
'End Sub
' 2005/08/03 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subCalculation_FstTOTAL
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCalculation_FstTOTAL()
' 2006/03/03 Del Tajima Start
'    ' 初回合計金額算出 & セット
'    Me.txt_初回合計 = Nz(Me.txt_初回金額, 0) + Nz(txt_初回鍵代, 0) + Nz(txt_初回消費税, 0)
' 2006/03/03 Del Tajima End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subGetFirstMoney
'        CREATE_DATE      :
'
'        処理内容         ：月額使用料を元に初回使用料を算出する
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetFirstMoney()

    Dim dteFSDate   As Date     ' 契約開始日の年月の初日
    Dim intLastDay  As Integer  ' 本月日数
    Dim intDay      As Integer  ' 本日
    Dim varMoney    As Variant  ' 初回使用料

    On Error GoTo subGetFirstMoney_Err

    ' 契約開始日が未入力の場合、処理を行わない？
    If Nz(txt_契約開始日, "") = "" Then Exit Sub

    ' 契約開始年月の初日セット
    dteFSDate = CDate(Format(txt_契約開始日, "YYYY/MM") & "/01")

    '本月日数算出
    intLastDay = Val(Format(DateAdd("d", -1, CDate(Format(dteFSDate, "YYYY") & "/" & Format(DateAdd("m", 1, dteFSDate), "MM") & "/01")), "DD"))
'    intLastDay = Val(Format(DateAdd("d", -1, CDate(Format(DATE, "YYYY") & "/" & Format(DateAdd("m", 1, DATE), "MM") & "/01")), "DD"))

    ' 本日取得
    intDay = Mid(Replace(Nz(txt_契約開始日, "0"), "/", "", 1, 10), 7, 2)

    ' 初回使用料 = (月額料金 / 本月日数) * (本月日数 - 本日)
' 2005/09/12 REP T.SUZUKI Start
    varMoney = Fix((Val(Me.txt_月額料金) * ((intLastDay - intDay + 1) / intLastDay)) / 100) * 100
'    varMoney = Fix((Val(Me.txt_月額料金) / intLastDay) * (intLastDay - intDay + 1))
' 2005/09/12 REP T.SUZUKI End
'    varMoney = Fix((Val(Me.txt_月額料金) / intLastDay) * (intLastDay - intDay))

' 2006/03/03 DEL Tajima Start
'    Me.txt_初回金額 = varMoney
'
'' 2005/08/03 ADD T.SUZUKI Start
'    ' 初回消費税算出
'    Call txt_初回金額_AfterUpdate
'' 2005/08/03 ADD T.SUZUKI End
' 2006/03/03 DEL Tajima End

    Exit Sub

subGetFirstMoney_Err:
'    Me.txt_初回金額 = 0 2006/03/03 DEL Tajima End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_消費税_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_消費税_AfterUpdate()
    Call subCalculation_TOTAL  ' 2005/08/03 ADD T.SUZUKI
'    Call txt_合計計算          ' 2005/08/03 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_消費税_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_消費税_BeforeUpdate(Cancel As Integer)

    Dim vNum As Variant

    On Error GoTo err_rtn

    vNum = Nz(Me.txt_消費税, "")

    If vNum = "" Then Exit Sub

' 2006/01/05 ADD T.SUZUKI Start
    If fncTxtInputCheck(Me.txt_消費税, pcsglMoneyMaxValue, 9) = False Then
        MsgBox strMSGKetaOver, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
' 2006/01/05 ADD T.SUZUKI End
' 2006/01/05 DEL T.SUZUKI Start
'    Select Case CHK_Numeric(vNum, 9)
'        Case 0
'
'        Case 1
'            MsgBox "消費税は、数値で、入力してください。", vbInformation, strMsg_kaku
'            Cancel = True
'            SendKeys "{ESC}"  '入力のキャンセル処理
'            Exit Sub
'        Case 2
'            MsgBox "消費税は、9桁以内で、入力してください。", vbInformation, strMsg_kaku
'            Cancel = True
'            SendKeys "{ESC}"  '入力のキャンセル処理
'            Exit Sub
'
'        Case Else
'            Cancel = True
'            SendKeys "{ESC}"
'            Exit Sub
'    End Select
' 2006/01/05 DEL T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_消費税_GotFocus
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_消費税_GotFocus()
    Me.txt_消費税.Format = pcstrfmtFIXED  ' 書式：固定(カンマ編集なし)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_消費税_LostFocus
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_消費税_LostFocus()
    Me.txt_消費税.Format = pcstrfmtSTANDARD  ' 書式：標準(カンマ編集あり)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_消費税_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_消費税_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_消費税, KeyAscii, 9)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_消費税_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_消費税_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約備考_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約備考_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    vCode = Nz(Me.txt_契約備考, "")

    '入力チェック
    Select Case CmBCheck_Char(vCode, 60)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "備考を入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約備考_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約備考_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Char2(Me.txt_契約備考, KeyAscii, 60)           ' 2006/01/27 REP T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_契約備考_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_契約備考_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_顧客マスタ修正_Click
'        CREATE_DATE      :
'        処理内容         ：画面で選択されている顧客コードを元に、顧客マスタ登録画面を修正モードで表示する
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_顧客マスタ修正_Click()

    On Error Resume Next

    ' ヤードコード、コンテナ番号が未選択の状態では次画面に遷移させない  ADD 2006/03/11 tajima
    If (Replace(Nz(Me.cmb_ヤードコード, ""), "0", "") = "" And Replace(Nz(Me.cmb_コンテナ番号, ""), "0", "") = "") Then
        Exit Sub
    End If

    ' 顧客名カナ入力チェック
    If Nz(Me.txt_顧客カナ, "") = "" Then                              ' 2005/10/19 ADD T.SUZUKI
'    If Nz(Me.cmb_顧客カナ, "") = "" Then                             ' 2005/10/19 DEL T.SUZUKI
        MsgBox "顧客名を選択して下さい。", vbInformation, strMsg_err
        Me.txt_顧客カナ.SetFocus                                      ' 2005/10/19 ADD T.SUZUKI
'        Me.cmb_顧客カナ.SetFocus                                     ' 2005/10/19 DEL T.SUZUKI
        Exit Sub
    End If

    ' 顧客コード入力チェック
    If (Nz(cmb_顧客コード, "") = "" Or cmb_顧客コード = 0) Then
        MsgBox "顧客コードを選択して下さい。", vbInformation, strMsg_err
        Me.cmb_顧客コード.SetFocus
        Exit Sub
    End If

    ' 顧客マスタ保守画面を表示
    ' ※画面で入力された顧客コードを顧客マスタ保守画面にパラメータとして渡す
    pblnFKS430_Flag = True                                                        ' 2006/03/04 ADD T.SUZUKI
'2007/07/11 del tajima：doCmd.OpenForm "FKS430", , , , , acDialog, Me.cmb_顧客コード & "," & Me.NAME  ' 2006/03/04 ADD T.SUZUKI
    doCmd.OpenForm "FKS430", , , , , , Me.cmb_顧客コード & "," & Me.NAME          ' 2007/07/11 ADD tajima
'    doCmd.OpenForm "FKS430", , , , , acDialog, 1 & Me.cmb_顧客コード             ' 2006/03/04 DEL T.SUZUKI

' 2006/03/04 DEL T.SUZUKI Start
'    If pblnChangeFlg = True Then
'        If CHK_USER_MAST(1) = False Then Exit Sub
'    End If
' 2006/03/04 DEL T.SUZUKI End

cmd_顧客マスタ修正_Click_Exit:
'    pblnChangeFlg = False                                                        ' 2006/03/04 DEL T.SUZUKI
End Sub

' 2005/10/19 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :cmb_顧客カナ_AfterUpdate
''        CREATE_DATE      :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_顧客カナ_AfterUpdate()
'
'    Me.cmb_顧客コード.Requery
'
'    Call CLR_RTN3(0, 1)
'    Call CLR_RTN3(1, 1)
'End Sub

''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :cmb_顧客カナ_Change
''        CREATE_DATE      :2005/08/03
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_顧客カナ_Change()
'    Call subSetChangeFlag
'End Sub
' 2005/10/19 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客カナ_AfterUpdate
'        CREATE_DATE      :2005/10/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客カナ_AfterUpdate()
    Call CLR_RTN3(0, 1) ' 【顧客情報】
    Call CLR_RTN3(1, 1) ' 【その他連絡先】
    Call CLR_RTN3(2, 1) ' 【緊急連絡先】                                        'INSERT 2011/04/11 K.ISHIZAKA
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客カナ_BeforeUpdate
'        CREATE_DATE      :2005/10/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客カナ_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    vCode = Nz(txt_顧客カナ, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "顧客名カナ" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客カナ_Change
'        CREATE_DATE      :2005/10/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客カナ_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_顧客コード_AfterUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_AfterUpdate()

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_顧客コード, "")

    If vCode = "" Then
        Exit Sub
    End If

    'ファイルへの存在チェック
    If CHK_USER_MAST(1) = False Then
        Exit Sub
    End If

    ' タブのページ設定
    If Me.tab_顧客情報.VALUE <> 0 Then
        Me.tab_顧客情報.VALUE = 0
    End If

    Me.cmb_顧客コード = Format(Me.cmb_顧客コード, "000000")
    Me.txt_顧客カナ.SetFocus                                                    'INSERT 2005/10/21 K.ISHIZAKA
    Me.cmb_顧客コード.SetFocus                                                  'INSERT 2005/10/21 K.ISHIZAKA

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_顧客コード_GotFocus
'        CREATE_DATE      :2005/10/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_GotFocus()

    Dim strSQL As String

    If Nz(Me.txt_顧客カナ, "") = "" Then
        Me.cmb_顧客コード.RowSource = ""
        Me.cmb_顧客コード = Null
        Exit Sub

    Else
        ' 顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ表示のSQL文作成
        'strSQL = "SELECT Format([USER_CODE],'000000') AS 式1, "                'DELETE 2005/10/21 K.ISHIZAKA
        strSQL = "SELECT TOP 100 Format([USER_CODE],'000000') AS 式1, "         'INSERT 2005/10/21 K.ISHIZAKA
        strSQL = strSQL & " dbo_USER_MAST.USER_NAME "
        strSQL = strSQL & "  FROM dbo_USER_MAST "
        strSQL = strSQL & " WHERE (((dbo_USER_MAST.USER_KANA) LIKE '*" & Me.txt_顧客カナ & "*')) "
        strSQL = strSQL & " ORDER BY Format([USER_CODE],'000000');"
    
        Me.cmb_顧客コード.RowSource = strSQL
        Me.cmb_顧客コード.Requery

        If Me.cmb_顧客コード.ListCount = 0 Then Me.cmb_顧客コード = Null
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_顧客コード_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo err_rtn

    vCode = Nz(Me.cmb_顧客コード, "")

    If vCode = "" Then
        Exit Sub
    End If

    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn

        Case 9
            MsgBox "コードを入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exit_rtn
    End Select

    'ファイルへの存在チェック
    If CHK_USER_MAST(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

    pintError = 0
    If Me.cmb_顧客コード <> pvarKcode Then
'        If Me.opg_処理選択 > 1 Then                                                             ' 2006/03/04 DEL T.SUZUKI
            If CHK_REQU_FILE(0) = False Then
                Else
                    MsgBox "請求処理がされているため変更できません。", vbInformation, strMsg_kaku
                    Cancel = True
                    SendKeys "{ESC}"
                    GoTo Exit_rtn
            End If
'        End If                                                                                  ' 2006/03/04 DEL T.SUZUKI
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_顧客コード_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_顧客コード, KeyAscii, 6)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_顧客コード_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客コード_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客名カナ_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名カナ_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_顧客名カナ_BeforeUpdate_Err

    vCode = Nz(Me.txt_顧客名カナ, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "顧客名カナ" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_顧客名カナ_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客名カナ_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名カナ_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客名_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_顧客名_BeforeUpdate_Err

    vCode = Nz(Me.txt_顧客名, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "顧客名" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_顧客名_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客名_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客名_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_代表者名カナ_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名カナ_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_代表者名カナ_BeforeUpdate_Err

    vCode = Nz(Me.txt_代表者名カナ, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "代表者名カナ" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_代表者名カナ_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_代表者名カナ_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名カナ_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_代表者名_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_代表者名_BeforeUpdate_Err

    vCode = Nz(Me.txt_代表者名, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "代表者名" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_代表者名_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_代表者名_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_代表者名_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_郵便番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_郵便番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_郵便番号_BeforeUpdate_Err

    vCode = Nz(Me.txt_郵便番号, "")

    '入力チェック
    Select Case CmBCheck_Char(vCode, 10)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "郵便番号を入力してください。", vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_郵便番号_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_郵便番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_郵便番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_郵便番号, KeyAscii, 10)  '←DB定義より10桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_郵便番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_郵便番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_Yubin_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Yubin_Click()

    Call MSZ0000_YUBIB(Me.NAME, Me.txt_郵便番号.NAME, Me.txt_住所1.NAME, _
                                Nz(Me.txt_郵便番号.VALUE), Nz(Me.txt_住所1.VALUE))
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_住所1_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所1_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_住所1_BeforeUpdate_Err

    vCode = Nz(Me.txt_住所1, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "都道府県市町村" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_住所1_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_住所1_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所1_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_住所2_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所2_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_住所2_BeforeUpdate_Err

    vCode = Nz(Me.txt_住所2, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "番地" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_住所2_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_住所2_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所2_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_住所3_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所3_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_住所3_BeforeUpdate_Err

    vCode = Nz(Me.txt_住所3, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "建物名等" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_住所3_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_住所3_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_住所3_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_電話番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_電話番号_BeforeUpdate_Err

    vCode = Nz(Me.txt_電話番号, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "電話番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_電話番号_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_電話番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_電話番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_電話番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_電話番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_携帯番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_携帯番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_携帯番号_BeforeUpdate_Err

    vCode = Nz(Me.txt_携帯番号, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "携帯番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_携帯番号_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_携帯番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_携帯番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_携帯番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_携帯番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_携帯番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_FAX番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_FAX番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_FAX番号_BeforeUpdate_Err

    vCode = Nz(Me.txt_FAX番号, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "携帯番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_FAX番号_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_FAX番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_FAX番号_KeyPress(KeyAscii As Integer)
  Call CmKeyPress_For_Tel(Me.txt_FAX番号, KeyAscii, 15)    '←DB定義より15桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_FAX番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_FAX番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_メールアドレス_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_メールアドレス_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_メールアドレス_BeforeUpdate_Err

    vCode = Nz(Me.txt_メールアドレス, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 60)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "メールアドレス" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_メールアドレス_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_メールアドレス_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_メールアドレス_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        機能             :cmb_顧客区分_BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客区分_BeforeUpdate(Cancel As Integer)

    Dim strName  As String
'    Dim vCode    As Variant
'    Dim strSQL   As String  ' 2005/08/03 DEL T.SUZUKI

    On Error GoTo err_rtn

    ' 画面読み込み時は未処理
    If pblnOpenFlag_FVS600 = False Then Exit Sub

    ' コードが空の場合解約区分欄も空欄表示
    If fncControlCrear(Me.cmb_顧客区分, Me.txt_顧客区分名) = False Then Exit Sub  ' 2005/08/03 ADD T.SUZUKI

'    vCode = Nz(Me.cmb_顧客区分, "")
'    If vCode = "" Then
'        Cancel = True
'        MsgBox "顧客区分" & strMSGInputError, vbInformation, strMsg_kaku
'        Me.txt_顧客区分名 = ""
'        Exit Sub
'    End If

    If Trim(Me.cmb_顧客区分) <> "" Then
        If Trim(Len(Me.cmb_顧客区分)) > 2 Then
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

' 2005/08/03 REP T.SUZUKI Start
    ' 顧客区分名取得
    strName = fncSelectRyak("020", Me.cmb_顧客区分)   ' 2006/01/05 ADD T.SUZUKI
'    strNAME = fncSelectRyak(Me.cmb_顧客区分, "020")   ' 2005/08/09 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_顧客区分, "020")  ' 2005/08/09 DEL T.SUZUKI
'    strSQL = " NAME_ID ='020'"
'    strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_顧客区分)
'    strNAME = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")

    If strName = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If
' 2005/08/03 REP T.SUZUKI End

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        機能             :cmb_顧客区分_AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客区分_AfterUpdate()

    Dim vCode    As Variant
    Dim strName  As String   ' 2005/08/03 DEL T.SUZUKI
'    Dim strSQL   As String  ' 2005/08/03 DEL T.SUZUKI

    vCode = Nz(Me.cmb_顧客区分, "")
    If vCode = "" Then
        Me.txt_顧客区分名 = ""
        Exit Sub
    End If

' 2005/08/03 REP T.SUZUKI Start
    ' 顧客区分名取得
    strName = fncSelectRyak("020", Me.cmb_顧客区分)   ' 2006/01/05 ADD T.SUZUKI
'    strNAME = fncSelectRyak(Me.cmb_顧客区分, "020")   ' 2005/08/09 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_顧客区分, "020")  ' 2005/08/09 DEL T.SUZUKI
'    strSQL = " NAME_ID ='020'"
'    strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_顧客区分)
'    strRYAK = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
' 2005/08/03 REP T.SUZUKI End

    Me.cmb_顧客区分 = Format(vCode, "00")
    Me.txt_顧客区分名 = strName

'    ' 変更フラグ                   ' 2005/08/03 DEL T.SUZUKI
'    pblnChangeFlag_FVS600 = True  ' 2005/08/03 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_顧客区分_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_顧客区分_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        機能             :cmb_集金区分_BeforeUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_BeforeUpdate(Cancel As Integer)

    Dim strName  As String
'    Dim vCode    As Variant  ' 2005/08/03 DEL T.SUZUKI
'    Dim strSQL   As String   ' 2005/08/03 DEL T.SUZUKI

    On Error GoTo err_rtn

    ' 画面読み込み時は未処理
    If pblnOpenFlag_FVS600 = False Then Exit Sub

    If fncControlCrear(Me.cmb_集金区分, Me.txt_集金区分名) = False Then Exit Sub  ' 2005/08/03 ADD T.SUZUKI

' 2005/08/03 DEL T.SUZUKI Start
'    vCode = Nz(Me.cmb_集金区分, "")
'    If vCode = "" Then
'        Cancel = True
'        MsgBox "集金区分" & strMSGInputError, vbInformation, strMsg_kaku
'        Me.txt_集金区分名 = ""
'        Exit Sub
'    End If
' 2005/08/03 DEL T.SUZUKI End

    If Trim(Me.cmb_集金区分) <> "" Then
        If Trim(Len(Me.cmb_集金区分)) > 2 Then
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            GoTo Exit_rtn
        End If
    End If

' 2005/08/03 REP T.SUZUKI Start
    ' 集金区分名取得
    strName = fncSelectRyak("030", Me.cmb_集金区分)   ' 2006/01/05 ADD T.SUZUKI
'    strNAME = fncSelectRyak(Me.cmb_集金区分, "030")   ' 2005/08/09 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_集金区分, "030")  ' 2005/08/09 DEL T.SUZUKI
'    strSQL = " NAME_ID ='030'"
'    strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_集金区分)
'    strRYAK = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
' 2005/08/03 REP T.SUZUKI End

    If strName = "" Then
        MsgBox strMSGNoDataError, vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
        GoTo Exit_rtn
    End If

Exit_rtn:
    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        機能             :cmb_集金区分_AfterUpdate
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_AfterUpdate()

    Dim vCode     As Variant
    Dim strName   As String
'    Dim strSQL    As String  ' 2005/08/03 DEL T.SUZUKI
'    Dim strWhere  As String  ' 2005/08/03 DEL T.SUZUKI

    vCode = Nz(Me.cmb_集金区分, "")
    If vCode = "" Then
        Me.txt_集金区分名 = ""
        Exit Sub
    End If

' 2005/08/03 REP T.SUZUKI Start
    ' 集金区分名取得
    strName = fncSelectRyak("030", Me.cmb_集金区分)   ' 2006/01/05 ADD T.SUZUKI
'    strNAME = fncSelectRyak(Me.cmb_集金区分, "030")   ' 2005/08/09 ADD T.SUZUKI
'    strNAME = fncSelectName(Me.cmb_集金区分, "030")  ' 2005/08/09 DEL T.SUZUKI
'    strSQL = " NAME_ID ='030'"
'    strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_集金区分)
'    strRYAK = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
' 2005/08/03 REP T.SUZUKI End

    Me.cmb_集金区分 = Format(vCode, "00")
    Me.txt_集金区分名 = strName

'    ' 変更フラグ                   ' 2005/08/03 DEL T.SUZUKI
'    pblnChangeFlag_FVS600 = True  ' 2005/08/03 DEL T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_集金区分_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_集金区分_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_締日_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_締日_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_締日_BeforeUpdate_Err

    vCode = Nz(Me.txt_締日, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmbCheck_Num(vCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "締日" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_締日_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_締日_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_締日_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_締日, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_締日_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_締日_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客備考_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客備考_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_顧客備考_BeforeUpdate_Err

    vCode = Nz(Me.txt_顧客備考, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 60)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "顧客備考" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_顧客備考_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_顧客備考_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_顧客備考_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :chk_入力する_Click
'        CREATE_DATE      :
'
'        処理内容         ：入力するチェックボックスのON/OFFにより、【その他連絡先】の使用の可否を設定する
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_入力する_Click()

    Dim intControl   As Integer

    ' 【その他連絡先】の使用可否、背景色の設定
    Call subTxtControl

    If Me.chk_入力する.VALUE = 0 Then
        ' 【その他連絡先】を初期化
        Call CLR_RTN3(1)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subTxtControl
'        CREATE_DATE      :
'
'        処理内容         ：【その他連絡先】の使用可否、背景色の設定を行う
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subTxtControl()

    Dim intControl    As Integer
    Dim lngBackColor  As Long
    Dim blnFLG        As Boolean

    With Me
'DEL 2006/03/03 Tajima Start
'        If (.opg_処理選択.Value = 1 And .chk_入力する.Value = -1 And .opg_顧客情報取得選択.Value = 1) Then
'            'チェックされてるとき(白)
'            intControl = 0
'            lngBackColor = plngInputColor
'            blnFlg = True
'        Else
'            'チェックされていないとき(黄色)
'            intControl = 1
'            lngBackColor = plngLockColor
'            blnFlg = False
'        End If
'DEL 2006/03/03 Tajima End

'Add 2006/03/03 Tajima Start
        intControl = 1
        lngBackColor = plngLockColor
        blnFLG = False
'Add 2006/03/03 Tajima End
        
        ' 【その他連絡先】
        ' 使用可否設定
        Call CmLockControl(.txt_連絡先担当名カナ, intControl)
        Call CmLockControl(.txt_連絡先担当者名, intControl)
        Call CmLockControl(.txt_連絡先郵便番号, intControl)

        ' 郵便番号変換コマンドボタン
        .cmd_RYubin.Enabled = blnFLG

        Call CmLockControl(.txt_連絡先住所1, intControl)
        Call CmLockControl(.txt_連絡先住所2, intControl)
        Call CmLockControl(.txt_連絡先住所3, intControl)
        Call CmLockControl(.txt_連絡先電話番号, intControl)
        Call CmLockControl(.txt_連絡先携帯番号, intControl)
        Call CmLockControl(.txt_連絡先FAX, intControl)

        ' 背景色設定処理
        Call subSetColor(.txt_連絡先担当名カナ, lngBackColor)
        Call subSetColor(.txt_連絡先担当者名, lngBackColor)
        Call subSetColor(.txt_連絡先郵便番号, lngBackColor)
        Call subSetColor(.txt_連絡先住所1, lngBackColor)
        Call subSetColor(.txt_連絡先住所2, lngBackColor)
        Call subSetColor(.txt_連絡先住所3, lngBackColor)
        Call subSetColor(.txt_連絡先電話番号, lngBackColor)
        Call subSetColor(.txt_連絡先携帯番号, lngBackColor)
        Call subSetColor(.txt_連絡先FAX, lngBackColor)
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先担当名カナ_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当名カナ_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先担当名カナ_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先担当名カナ, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 18)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先担当名カナ" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先担当名カナ_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先担当名カナ_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当名カナ_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先担当者名_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当者名_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先担当者名_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先担当者名, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先担当者名" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先担当者名_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先担当者名_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先担当者名_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先郵便番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先郵便番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先郵便番号_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先郵便番号, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 10)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先郵便番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先郵便番号_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先郵便番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先郵便番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_連絡先郵便番号, KeyAscii, 10)  '←DB定義より10桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先郵便番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先郵便番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_RYubin_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_RYubin_Click()

    Call MSZ0000_YUBIB(Me.NAME, Me.txt_連絡先郵便番号.NAME, Me.txt_連絡先住所1.NAME, _
                                Nz(Me.txt_連絡先郵便番号.VALUE), Nz(Me.txt_連絡先住所1.VALUE))
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先住所1_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先住所1_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先住所1_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先住所1, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先都道府県市町村" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先住所1_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先住所1_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先住所1_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先住所2_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先住所2_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先住所2_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先住所2, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先番地" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先住所2_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先住所2_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先住所2_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先住所3_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先住所3_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先住所3_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先住所3, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 36)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先建物名等" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先住所3_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先住所3_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先住所3_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先電話番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先電話番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先電話番号_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先電話番号, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先電話番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先電話番号_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先電話番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先電話番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_連絡先電話番号, KeyAscii, 15)  '←DB定義より10桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先電話番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先電話番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先携帯番号_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先携帯番号_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先携帯番号_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先携帯番号, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先携帯番号" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先携帯番号_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先携帯番号_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先携帯番号_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_連絡先携帯番号, KeyAscii, 15)  '←DB定義より10桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先携帯番号_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先携帯番号_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先FAX_BeforeUpdate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先FAX_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo txt_連絡先FAX_BeforeUpdate_Err

    vCode = Nz(Me.txt_連絡先FAX, "")

    If vCode = "" Then Exit Sub

    '入力チェック
    Select Case CmBCheck_Char(vCode, 15)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox strMSGKetaOver, vbInformation, strMsg_err
            Cancel = True
            Exit Sub

        Case 9
            MsgBox "連絡先ＦＡＸ" & strMSGInputError, vbInformation, strMsg_err
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
    End Select

    Exit Sub

txt_連絡先FAX_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先FAX_KeyPress
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先FAX_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_連絡先FAX, KeyAscii, 15)    '←DB定義より10桁まで入力可。
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_連絡先FAX_Change
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_連絡先FAX_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_保存_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_保存_Click()

    Dim beforForm As Form  ' 2005/09/12 ADD

    On Error GoTo err_rtn

    ' 受注契約番号がNullの状態のとき、処理を抜ける
    If Nz(Me.txt_受注契約番号, "") = "" Then Exit Sub     ' 2006/03/18 ADD T.SUZUKI

    ' ヤードコード、コンテナ番号が未選択の状態では次画面に遷移させない ADD 2006/03/11 tajima
    If (Replace(Nz(Me.cmb_ヤードコード, ""), "0", "") = "" And Replace(Nz(Me.cmb_コンテナ番号, ""), "0", "") = "") Then
        Exit Sub
    End If

    ' 更新前の入力チェックを行う
    If CHK_RTN_For_UpDate = False Then
        Exit Sub
    End If

' 2006/01/05 ADD T.SUZUKI Start
    ' 移動元契約データチェック
    If fncCheckMoveMotoData = False Then
        GoTo Exit_rtn
    End If
' 2006/01/05 ADD T.SUZUKI End

    ' 確認メッセージ
    pintAns = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, strMsg_kaku)
    If pintAns = vbNo Then
        GoTo Exit_rtn
    End If
    
' 2019/04/23 ADD TAKENOUCHI Start
    If Me.txt_申込種別.VALUE <> "" Then
        'ネット解約DBの更新
        If fncUpdNetDBMain = False Then
            GoTo err_rtn
        Else
            '顧客マスタ(KOTO_TRAN)登録
            If Not fncInsKOTO_TRAN(Me.txt_受注契約番号) Then GoTo err_rtn
        End If
    End If
' 2019/04/23 ADD TAKENOUCHI End

    ' データ登録処理を行う
    If WRI_RTN(P_SHU) = True Then                                               ' 2006/03/04 ADD T.SUZUKI
'    If WRI_RTN = True Then                                                     ' 2006/03/04 DEL T.SUZUKI
'        Call CLR_RTN2                                                          'DELETE 20050916 K.ISHIZAKA
        Call CLR_RTN2(2)                                                                                     ' 2006/03/04 ADD T.SUZUKI
'        Call CLR_RTN2(opg_処理選択.Value)                                      'INSERT 20050916 K.ISHIZAKA  ' 2006/03/04 DEL T.SUZUKI

        ' 画面項目のロック処理
        Call subLock_Main
        Call subLock_Etc
    End If

    ' 変更フラグセット
    pblnUpdateFlg = True                                                        ' 2006/01/05 ADD T.SUZUKI

' 2005/08/03 ADD T.SUZUKI Start
    ' 契約参照モードの場合、画面にデータをリセットする
    If pintOpenMode = 3 Then
        If fncSetYoyakuData(2) = False Then Exit Sub
    End If
' 2005/08/03 ADD T.SUZUKI End

Exit_rtn:
    On Error Resume Next                                                        ' 2006/01/28 MOVE T.SUZUKI

    Me.cmb_コンテナ番号.SetFocus                                                ' 2006/03/04 ADD T.SUZUKI
' 2006/03/04 DEL T.SUZUKI Start
'' 2005/08/03 REP T.SUZUKI Start
'    Select Case Me.opg_処理選択.Value
'        Case 1  ' *** 新規 ***
'            If pintOpenMode = 1 Then
''                Me.cmb_受付番号.SetFocus DEL 2006/03/03 Tajima
'
'            ElseIf 3 Then
''                Me.txt_解約日付.SetFocus
'
'            ElseIf 9 Then                                                       ' 2006/01/28 ADD T.SUZUKI
'                ' 予約一覧から契約入力へ遷移してきた場合
'            Else
'                Me.cmb_コンテナ番号.SetFocus
'            End If
'
'        Case 2  ' *** 修正 ***
'            If pintOpenMode = 2 Then
'                Me.cmb_コンテナ番号.SetFocus
'            ElseIf pintOpenMode = 3 Then
''                Me.txt_解約日付.SetFocus
'            End If
'
'        Case 3  ' *** 削除 ***
'            Me.cmb_コンテナ番号.SetFocus
'    End Select
' 2006/03/04 DEL T.SUZUKI End

'    If (Me.opg_処理選択.Value = 1 And pintOpenMode = 1) Then
'        ' 新規契約 かつ 取込ありのとき
'        Me.cmb_受付番号.SetFocus
'    Else
'        ' 上記以外
'        Me.cmb_コンテナ番号.SetFocus
'    End If
' 2005/08/03 REP T.SUZUKI End

' 2006/01/05 ADD T.SUZUKI Start
'    On Error Resume Next                                                       ' 2006/01/28 MOVE T.SUZUKI

    ' *** 利用状況からの画面遷移の場合 ***
    Set beforForm = Forms("FKS020_コンテナ予約入力")
    If Not beforForm Is Nothing Then
        Call beforForm.Reload(Me.cmb_ヤードコード, Me.cmb_ヤードコード)
    End If

    ' *** 予約一覧からの画面遷移の場合 ***
    Set beforForm = Forms("FVS540")
    If Not beforForm Is Nothing Then
        ' 保存処理後、画面を閉じる
        Call cmd_終了_Click
    End If
' 2006/01/05 ADD T.SUZUKI End
' 2006/01/05 DEL T.SUZUKI Start
'' 2005/09/12 ADD T.SUZUKI Start
'    On Error Resume Next
'    Set beforForm = Forms("FKS020_コンテナ予約入力")
'    If Not beforForm Is Nothing Then
'        Call beforForm.Reload(Me.cmb_ヤードコード, Me.cmb_ヤードコード)
'    End If
'' 2005/09/12 ADD T.SUZUKI End
' 2006/01/15 DEL T.SUZUKI End

    Exit Sub

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_削除_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_削除_Click()
  
    Dim strMsg    As String
    Dim lCount    As Long
    Dim stWHERE   As String

    Dim intErr    As Integer

    ' 受注契約番号がNullの状態のとき、処理を抜ける
    If Nz(Me.txt_受注契約番号, "") = "" Then Exit Sub     ' 2006/03/18 ADD T.SUZUKI

    ' 入力チェック
' 2006/03/04 ADD T.SUZUKI Start
    strMsg = ""
    If Replace(Nz(Me.cmb_ヤードコード, ""), "0", "") = "" Then
        intErr = 1
        strMsg = "ヤードコードが入力されていません。"
    ElseIf Replace(Nz(Me.cmb_コンテナ番号, ""), "0", "") = "" Then
        intErr = 2
        strMsg = "コンテナ番号が入力されていません。"
    End If
    If Nz(strMsg, "") <> "" Then GoTo Exit_err
' 2006/03/04 ADD T.SUZUKI End
' 2006/03/04 DEL T.SUZUKI Start
'    If Nz(Me.cmb_コンテナ番号, "") = "" Then                                  ' 2006/03/04 DEL T.SUZUKI
'        MsgBox "コンテナ番号が入力されていません。", vbInformation, strMsg_kaku
'        GoTo Exit_rtn
'    End If
' 2006/03/04 DEL T.SUZUKI End

' 2006/03/18 ADD T.SUZUKI Start
    ' 請求ファイルチェック
    strMsg = "": strMsg = strMsg & "請求処理がされているため変更できません。"
    If CHK_REQU_FILE(0) = True Then
        MsgBox strMsg, vbInformation, strMsg_kaku
        GoTo Exit_rtn
    End If
' 2006/03/18 ADD T.SUZUKI End

    strMsg = "": strMsg = strMsg & "契約を取消し、受付状態に戻します。よろしいですか？"

    pintAns = MsgBox(strMsg, vbCritical + vbYesNo + vbDefaultButton2, strMsg_kaku)
    If pintAns = vbNo Then
        GoTo Exit_rtn
    End If

    ' データ削除処理を行う
    If WRI_RTN(P_DEL) = True Then                                               ' 2006/03/04 ADD T.SUZUKI
'    If WRI_RTN = True Then                                                     ' 2006/03/04 DEL T.SUZUKI
'        Call CLR_RTN2                                                          'DELETE 20050916 K.ISHIZAKA
        Call CLR_RTN2(2)                                                                                      ' 2006/03/04 ADD T.SUZUKI
'        Call CLR_RTN2(opg_処理選択.Value)                                       'INSERT 20050916 K.ISHIZAKA  ' 2006/03/04 DEL T.SUZUKI

        ' 画面項目のロック処理
        Call subLock_Main
        Call subLock_Etc
    End If

Exit_rtn:
    Me.cmb_コンテナ番号.SetFocus
    Exit Sub

Exit_err:
' 2006/03/04 ADD T.SUZUKI Start
    MsgBox strMsg, vbInformation, strMsg_kaku

    Select Case intErr
        Case 1
            Me.cmb_ヤードコード.SetFocus
        Case 2
            Me.cmb_コンテナ番号.SetFocus
    End Select
' 2006/03/04 ADD T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_取消_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_取消_Click()

    ' ヤードコード、コンテナ番号が未選択の状態では次画面に遷移させない ADD 2006/03/11 tajima
    If (Replace(Nz(Me.cmb_ヤードコード, ""), "0", "") = "" And Replace(Nz(Me.cmb_コンテナ番号, ""), "0", "") = "") Then
        Exit Sub
    End If
    
    ' 受注契約番号がNullの状態のとき、処理を抜ける
    If Nz(Me.txt_受注契約番号, "") = "" Then Exit Sub     ' 2006/03/18 ADD T.SUZUKI

    pintAns = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo + vbDefaultButton2, strMsg_kaku)
    If pintAns = vbNo Then
        GoTo Exit_rtn
    End If

' 2005/08/03 ADD T.SUZUKI Start
    If pintOpenMode = 2 Then
'        pstrYCODE = Me.cmb_ヤードコード  ' ヤードコード                        'DELETE 20050916 K.ISHIZAKA
'        pstrNO = Me.cmb_コンテナ番号     ' コンテナ番号                        'DELETE 20050916 K.ISHIZAKA
        pstrYCode = Nz(Me.cmb_ヤードコード)  ' ヤードコード                     'INSERT 20050916 K.ISHIZAKA
        pstrNo = Nz(Me.cmb_コンテナ番号)     ' コンテナ番号                     'INSERT 20050916 K.ISHIZAKA
    End If
' 2005/08/03 ADD T.SUZUKI End

    ' 画面の初期化処理(全項目初期化)
    Call CLR_RTN
    Call CLR_RTN3(0)  ' 【顧客情報】
    Call CLR_RTN3(1)  ' 【その他連絡先】
    Call CLR_RTN3(2)  ' 【緊急連絡先】                                          'INSERT 2011/04/11 K.ISHIZAKA

    Select Case pintOpenMode
        Case 0     '*** メインメニューより遷移 ***
            ' 契約新規・取込なし
            Call subLock_Etc

'            Me.chk_予約受付.SetFocus  DEL 2006/03/03 Tajima
                   
        Case 1  ' *** ヤード利用状況画面(予約受付トラン読込)
            Call subLock_Main

            ' 契約新規・取込あり
            ' 予約受付トランから該当データを取り込み、画面に設定
            ' 画面の各項目の使用可否は『cmb_コンテナ番号_AfterUpdate』イベントにて『LOCK_RTN』をコール
            If fncGetYOUK_TRAN = False Then Exit Sub

        Case 2, 3  ' *** ヤード利用状況画面(コンテナ契約ファイル読込)  ' 2005/08/03 ADD T.SUZUKI
'        Case 2  ' *** ヤード利用状況画面(コンテナ契約ファイル読込)    ' 2005/08/03 ADD T.SUZUKI
            ' 契約修正（解約関連項目のみ入力可）

            ' コンテナ契約ファイルのデータを画面にセット
            If fncSetYoyakuData(2) = False Then Exit Sub

' 2005/08/03 ADD T.SUZUKI Start
            If pintOpenMode = 2 Then
                pstrYCode = ""  ' ヤードコード
                pstrNo = ""     ' コンテナ番号
            End If
' 2005/08/03 ADD T.SUZUKI End

            ' ※画面の各項目の使用可否は『cmb_コンテナ番号_AfterUpdate』イベントにて『LOCK_RTN』をコール
            '   (フォーカス設定も含む)

'            Me.cmd_削除.Enabled = False                                ' 2006/03/04 ADD T.SUZUKI

' 2006/01/05 ADD T.SUZUKI Start
        Case 9  ' *** 予約一覧(予約受付トラン読込)
            Call subLock_Main

            ' フォーカスを移動しておく
            '2005/09/22 SHIBAZAKI 予約名／予約番号削除
            'Me.txt_予約名.SetFocus  ' 元はtxt_契約開始日 2006/02/18 Add TAJIMA

            ' 予約受付からチェックボックス、受付番号をし使用不可に設定
' DEL 2006/03/03 Tajima Start
'            Me.chk_予約受付.Enabled = False         ' 予約受付からチェックボックス：OFF
'            Call CmLockControl(Me.cmb_受付番号, 1)  ' 受付番号コンボボックス      ：OFF
' DEL 2006/03/03 Tajima End
            ' 契約新規・取込あり
            ' 予約受付トランから該当データを取り込み、画面に設定
            ' 画面の各項目の使用可否は『cmb_コンテナ番号_AfterUpdate』イベントにて『LOCK_RTN』をコール
            If fncGetYOUK_TRAN = False Then Exit Sub
' 2006/01/05 ADD T.SUZUKI End
    End Select

Exit_rtn:
    ' 予約受付からチェックボックスにフォーカス移動
'    Me.cmb_ヤードコード.SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmd_終了_Click
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_終了_Click()

    Dim beforForm As Form

    On Error Resume Next

    doCmd.Close acForm, Me.FormName

' 2006/01/05 ADD T.SUZUKI Start
    Set beforForm = Forms("FVS540")
    If Not beforForm Is Nothing Then
        Call beforForm.subReload(IIf(pblnUpdateFlg = True, 1, 0))
    End If
' 2006/01/05 ADD T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_End_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_End_GotFocus()

    Me.cmd_保存.SetFocus                                                       ' 2006/03/04 ADD T.SUZUKI
' 2006/03/04 DEL T.SUZUKI STart
'    Select Case Me.opg_処理選択.Value
'        Case 3
'            Me.cmd_削除.SetFocus
'        Case Else
'            Me.cmd_保存.SetFocus
'    End Select
' 2006/03/04 DEL T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :txt_HEnd_GotFocus
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_HEnd_GotFocus()
    Me.cmb_ヤードコード.SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

'    Dim beforForm As Form  ' 2005/09/12 DEL T.SUZUKI

    On Error Resume Next
    
    Call MSHS010_M00(PROG_ID, "2")

    'オブジェクト変数の解放
'    Set pobjRs = Nothing   ' 2006/001/05 DEL T.SUZUKI
    Set pobjDB = Nothing
    Set pWRI_ObjWS = Nothing
    Set pWRI_objDb = Nothing
'    Set pobjRS2 = Nothing  ' 2006/001/05 DEL T.SUZUKI

' 2005/09/12 DEL T.SUZUKI Start
'    Set beforForm = Forms("FKS020_コンテナ予約入力")
'    If Not beforForm Is Nothing Then
'        Call beforForm.Reload
'    End If
' 2005/09/12 DEL T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_KeyDown
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_終了.SetFocus

        Case vbKeyF2
            KeyCode = 0

            Me.cmd_保存.SetFocus                                                ' 2006/03/04 ADD T.SUZUKI

      ' ▼2006/09/26 Add tajima
        Case vbKeyF8, vbKeyF9, vbKeyF12
          If Me.Form.NAME = pstrFVS600 Then
              Call CmReprMod.OpenFVS90n(KeyCode, _
                                        Me.cmb_ヤードコード, _
                                        Me.cmb_コンテナ番号, _
                                        Me.cmb_顧客コード _
                                        )
          ElseIf CmReprMod.P_呼出_ヤード Or CmReprMod.P_呼出_顧客 Or CmReprMod.P_呼出_部屋 Then
              MsgBox "参照モードでは修繕関連の呼出はできません。" & vbCr & "利用状況にて呼出をしてください。", vbInformation, strMsg_kaku
          End If
          KeyCode = 0 '2006/10/12 add tajima
      ' ▲2006/09/26 Add tajima


        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Unload
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)

    Dim strMsg As String
    Dim varDefBtn As Variant

    On Error GoTo Err_Form_Unload

    ' 予約一覧からの遷移の場合、以降の処理は行わない
    If pintOpenMode = 9 Then Exit Sub                                          ' 2006/01/26 ADD T.SUZUKI

' 2006/03/18 ADD tajima Start
    If pblnChangeFlag_FVS600 = True And Nz(Me.txt_受注契約番号, "") <> "" Then
      strMsg = "修正中ですが保存せずに終了しますか？"
      varDefBtn = vbQuestion + vbYesNo + vbDefaultButton2
    Else
      strMsg = "終了します。よろしいですか？"
      varDefBtn = vbQuestion + vbYesNo + vbDefaultButton1
    End If
    
' 2006/03/18 ADD tajima End
    
    '' 2013/06/03 M.HONDA START
    '****************************************
    '*  終了メッセージを表示　　　　　　　　*
    '****************************************
'    pintAns = MsgBox(strMsg, varDefBtn, "確認")
'    If pintAns = vbNo Then
'        Cancel = True
'
'' 2005/08/03 REP T.SUZUKI Start
'        Select Case pintOpenMode
'            Case 1
'                ' 予約受付からチェックボックスにフォーカス移動
''                Me.chk_予約受付.SetFocus DEL 2006/03/03 Tajima Start
'            Case 2
'                Me.cmb_ヤードコード.SetFocus
'            Case 3
'                Me.txt_解約日付.SetFocus
'            Case Else
'                Me.cmb_ヤードコード.SetFocus
'        End Select
''        If pintOpenMode = 1 Then
''            ' 予約受付からチェックボックスにフォーカス移動
''            Me.chk_予約受付.SetFocus
''        Else
''            Me.cmb_ヤードコード.SetFocus
''        End If
'' 2005/08/03 REP T.SUZUKI End
'
'        Exit Sub
'    End If
    '' 2013/06/13 M.HONDA END

    Exit Sub

Err_Form_Unload:
    MsgBox Error$
    Exit Sub
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :Form_Error
'        CREATE_DATE      :
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

'■□■□■□■□■□■□■□■□■□ ↓↓↓　Private 関数群 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□
' <<<画面制御系>>>
'==============================================================================*
'
'        MODULE_NAME      :画面の標題設定
'        MODULE_ID        :subSetFormTittle
'        CREATE_DATE      :2004/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetFormTittle()

    pstrFormName = Me.Caption

    Select Case Me.Form.NAME
        Case pstrFVS600, pstrFVS602
            Me.Caption = pstrFormName

        Case pstrFVS601
            pstrFormName = Replace(pstrFormName, "600", "601")
            Me.Caption = pstrFormName
    End Select
End Sub

'==============================================================================*
'
'        MODULE_NAME      :各入力項目修正時、フォームの標題を『修正中』に設定
'        MODULE_ID        :subSetChangeFlag
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetChangeFlag()

    If pblnChangeFlag_FVS600 = False Then
        Me.Caption = pstrFormName & "(修正中)"
        pblnChangeFlag_FVS600 = True  ' 変更フラグ(ON)
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subFormLockControl
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFormLockControl(blnFLG As Boolean)
    Dim ctrl                As Variant                                          'INSERT 2011/04/11 K.ISHIZAKA

    With Me
        ' フォームのプロパティ設定
        .AllowEdits = blnFLG      ' 更新の許可
        .AllowDeletions = blnFLG  ' 削除の許可
        .AllowAdditions = blnFLG  ' 追加の許可
        .DataEntry = blnFLG       ' データ入力用

        .txt_解約日付.SetFocus     ' 解約日付にフォーカスを移動

        .cmb_受付担当者コード.Enabled = blnFLG
        .cmb_受付担当者コード.Locked = IIf(blnFLG = False, True, False)
        .cmb_ヤードコード.Enabled = blnFLG
        .cmb_ヤードコード.Locked = IIf(blnFLG = False, True, False)
        .cmb_コンテナ番号.Enabled = blnFLG
        .cmb_コンテナ番号.Locked = IIf(blnFLG = False, True, False)
        .cmb_解約区分.Enabled = blnFLG
        .cmb_解約区分.Locked = IIf(blnFLG = False, True, False)
        .cmb_解約理由.Enabled = blnFLG
        .cmb_解約理由.Locked = IIf(blnFLG = False, True, False)
        .cmb_解約受付担当者コード.Enabled = blnFLG
        .cmb_解約受付担当者コード.Locked = IIf(blnFLG = False, True, False)
        .cmb_鍵区分.Enabled = blnFLG
        .cmb_鍵区分.Locked = IIf(blnFLG = False, True, False)
       
        
        '' 20120411 M.HONDA ADD
        .cmb_スロープ.Enabled = blnFLG
        .cmb_スロープ.Locked = IIf(blnFLG = False, True, False)
        
' DEL 2006/06/01 Tajima Start 入力不可項目なので制御しない
'        .cmb_保証区分.Enabled = blnFlg
'        .cmb_保証区分.Locked = IIf(blnFlg = False, True, False)
' DEL 2006/06/01 Tajima End
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
        .cmb_請求書発行区分.Enabled = blnFLG
'       .cmb_請求書発行区分.Enabled = False
       
        .cmb_請求書発行区分.Locked = Not blnFLG                                 'INSERT 2010/04/10 K.ISHIZAKA
        .cmd_カードメンテ.Enabled = blnFLG  '2006/06/08 ADD tajima
        .cmd_削除.Enabled = blnFLG          '2006/06/08 ADD tajima
        .cmd_顧客マスタ修正.Enabled = blnFLG
' DEL 2006/03/03 Tajima Start
'        .txt_顧客カナ.Enabled = blnFlg                            ' 2005/10/19 ADD T.SUZUKI
'        .txt_顧客カナ.Locked = IIf(blnFlg = False, True, False)   ' 2005/10/19 ADD T.SUZUKI
'        .cmb_顧客コード.Enabled = blnFlg
'        .cmb_顧客コード.Locked = IIf(blnFlg = False, True, False)
' DEL 2006/03/03 Tajima End
        .cmd_Yubin.Enabled = blnFLG
        .cmd_RYubin.Enabled = blnFLG
' DEL 2006/06/01 Tajima Start 入力不可項目なので制御しない
'        .cmb_顧客区分.Enabled = blnFlg
'        .cmb_顧客区分.Locked = IIf(blnFlg = False, True, False)
'        .cmb_集金区分.Enabled = blnFlg
'        .cmb_集金区分.Locked = IIf(blnFlg = False, True, False)
' DEL 2006/06/01 Tajima End

'        .opg_処理選択.Enabled = blnFlg                                     ' 2006/03/04 DEL T.SUZUKI

'        .opg_顧客情報取得選択.Enabled = blnFlg                             DEL 2006/03/03 Tajima
  '        .opg_顧客情報取得選択.Locked = IIf(blnFlg = False, True, False)  DEL 2006/03/03 Tajima

        For Each ctrl In pge_緊急連絡先.Controls                                'INSERT START 2011/04/11 K.ISHIZAKA
            Select Case UCase(TypeName(ctrl))
            Case "TEXTBOX"
                ctrl.Enabled = (blnFLG And isRCPT_TRAN)
                ctrl.Locked = Not (blnFLG And isRCPT_TRAN)
            Case "COMMANDBUTTON"
                ctrl.Enabled = (blnFLG And isRCPT_TRAN)
            End Select
        Next                                                                    'INSERT END   2011/04/11 K.ISHIZAKA

        .cmd_保存.Enabled = blnFLG
        .cmd_取消.Enabled = blnFLG
        .cmd_終了.Enabled = blnFLG

        ' 標題設定
        If blnFLG = True Then
            .Caption = pstrFormName
        Else
            .Caption = pstrFormName & pstrKaiyakuMsg
        End If
        
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面上のテキストボックスをクリアする
'        MODULE_ID        :CLR_RTN
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CLR_RTN()
  
    Dim vRet  As Variant

    On Error Resume Next  'エラーがあっても処理続行

    With Me
        ' 上記以外
        If pintOpenMode = 1 Then
            ' 契約新規・取込
        Else
            ' 上記以外
'            .chk_予約受付.Value = 0  DEL 2006/03/03 Tajima Start
'            .cmb_受付番号 = Null     DEL 2006/03/03 Tajima Start
        End If

        .cmb_受付担当者コード = ""
        .txt_受付担当者名 = ""

'        .cmb_ヤードコード = 0                                                  'DELETE 20050916 K.ISHIZAKA
        .cmb_ヤードコード = "000000"                                            'INSERT 20050916 K.ISHIZAKA
        .txt_ヤード名 = Null
        .txt_ヤード住所 = Null
'        .txt_地主コード = Null                                                 ' 2005/08/03 DEL T.SUZUKI
'        .txt_地主名 = Null                                                     ' 2006/09/22 DEL SHIBAZAKI
'        .txt_物件番号 = Null                                                   ' 2006/09/22 DEL SHIBAZAKI

        '2006/09/22 SHIBAZAKI ヤード営業終了日､ヤード解約予定日追加
        .txt_営業終了日 = ""
        .txt_ヤード解約予定 = ""
        
        
        .txt_mente = ""
        .txt_jun = ""
        .txt_Keitai = ""
        .txt_spe = ""
        .txt_seq = ""
        
        .txt_受注契約番号 = Null
        pstrJUCHUKEIYAKU_NO = ""                                                ' 2005/08/03 ADD T.SUZUKI
'        strJUCHUKEIYAKU_NO = ""                                                ' 2005/08/03 DEL T.SUZUKI

'        .cmb_コンテナ番号 = 0                                                  'DELETE 20050916 K.ISHIZAKA
        .cmb_コンテナ番号 = "000000"                                            'INSERT 20050916 K.ISHIZAKA
'        .txt_段区分 = 0                                                        ' 2005/08/03 DEL T.SUZUKI
        .txt_段区分名 = Null
'        .txt_コンテナ種別 = 0                                                  ' 2005/08/03 DEL T.SUZUKI
        .txt_コンテナ種別名 = Null

        .txt_種別 = Null                                                        ' 2013/05/09 M.HONDA INS
        Me.txt_種別.BackColor = 14680063                                        ' 2013/05/09 M.HONDA INS

        '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'        .txt_予約名 = Null
'        .txt_予約番号 = Null

        '2005/09/22 SHIBAZAKI 直近修繕表示
        .txt_直近修繕 = Null
        .txt_クレーム = Null

        .txt_初回契約日 = Null
        .txt_契約開始日 = Null
        .txt_契約満了日 = Null

        .txt_解約日付 = Null
        .txt_解約日付.Tag = ""                                                  ' 2005/08/03 ADD T.SUZUKI
'        .chk_コンテナ状態 = 0                                                  ' 2005/08/03 DEL T.SUZUKI

        .txt_解約受付日 = Null                                                  ' 2005/08/03 ADD T.SUZUKI
        .cmb_解約区分 = Null                                                    ' 2005/08/03 ADD T.SUZUKI
        .txt_解約区分 = Null                                                    ' 2005/08/03 ADD T.SUZUKI
'        .txt_解約受付日 = Format(Now, "YYYY/MM/DD")
'        .cmb_解約区分 = "01"
'        Call cmb_解約区分_AfterUpdate
        .cmb_解約理由 = Null                                                    ' 2005/08/03 ADD T.SUZUKI
        .txt_解約理由 = Null                                                    ' 2005/08/03 ADD T.SUZUKI
        .cmb_解約受付担当者コード = Null                                        ' 2005/08/03 ADD T.SUZUKI
        .txt_解約受付担当者名 = Null                                            ' 2005/08/03 ADD T.SUZUKI

'2012/04/11 INS M.HONDA
        .cmb_スロープ = Null
        .txt_スロープ = Null

        .cmb_鍵区分 = 0
        Call cmb_鍵区分_AfterUpdate
        .txt_解除番号 = Null
        .chk_書類1 = -1

        .cmb_保証区分 = 4
        Call cmb_保証区分_AfterUpdate
        .cmb_保証会社コード = Null                                              ' 2009/04/01 INS KIHARA
        Call cmb_保証会社コード_AfterUpdate                                     ' 2009/04/01 INS KIHARA
        
        .txt_承認日 = Null
        .txt_承認番号 = Null

        .txt_月額料金 = 0
        .txt_消費税 = 0
        .txt_合計 = 0

        .txt_預り保証金 = 0

        .txt_契約備考 = Null
        .txt_FaceAuth = Null
        
        
        .txt_RCPT_BIKE_NUMBER = Null
        .txt_RCPT_BIKE_SHASYU = Null
        
        .txt_mente = Null
        .txt_jun = Null
'2006/05/07 ADD tajima Start
        .txt_スペアキー備考 = Null
        .txt_貸出本数 = 0
        .txt_セキュリティカード枚数 = Null
'2006/05/07 ADD tajima End
        
'        .opg_顧客情報取得選択.Value = 1  ' デフォルト：新規  DEL 2006/03/03 Tajima
        .txt_キャンペーン適用満了日 = Null      'ADD 2006/03/03 Tajima
        .cmd_キャンペーン内容.Enabled = False   'ADD 2006/03/03 Tajima

        .lbl_Keiyaku.Caption = P_TITTLE_KEIYAKU              ' 2006/01/26 ADD T.SUZUKI
    End With

' 2006/01/05 ADD T.SUZUKI Start
    ' 構造体初期化
    ' [予約受付トラン]
    With SAKI_YOUK_TRAN
        .strMOTO_ACPTNO = ""             ' 元受注契約番号
        .strMOVEKBN = P_MOVEKBN_00       ' 移動区分(無し)
        .strUkDate = ""                  ' 受付日
        .strUKTANTO = ""                 ' 受付担当者コード
    End With

    ' [コンテナ契約ファイル]
    With MOTO_CARG_FILE
        .strYcode = ""                   ' ヤードコード
        .strNO = ""                      ' コンテナ番号
        .strUCODE = ""                   ' 顧客コード
        .strAGRE = ""                    ' 契約状態
        .dblRENTKG = 0                   ' 月額料金
        .strACPTNO = ""                  ' 受注契約番号
        .strHOSYICD = ""                 ' 保証区分
        .strUKTANTO = ""                 ' 契約受付担当者コード
    End With
' 2006/01/05 ADD T.SUZUKI End

    ' コントロールマスタより取得 ------------------------------------------
    ' 消費税切替日
    vRet = Nz(DLookup("CONT_HENRIT", "dbo_CONT_MAST", "CONT_KEY = 1"), "")
    pvarHenrit = vRet

    ' 切替前消費税率
    vRet = Nz(DLookup("CONT_MAERIT", "dbo_CONT_MAST", "CONT_KEY = 1"), "")
    pvarMaerit = vRet

    ' 切替後消費税率
    vRet = Nz(DLookup("CONT_ATORIT", "dbo_CONT_MAST", "CONT_KEY = 1"), "")
    pvarAtorit = vRet
    ' ---------------------------------------------------------------------




    ' システム日付
    pvarDATE = Format(Now, P_FORMAT_YMD)                                        ' 2006/01/05 ADD T.SUZUKI
'    pvarDATE = Format(Now, "yyyy/mm/dd")                                       ' 2006/01/05 DEL T.SUZUKI

    ' 変更フラグの初期化
    pblnChangeFlag_FVS600 = False
    pblnMoneyChgFlg = False                                                     ' 2005/08/03 ADD T.SUZUKI
    Me.Caption = pstrFormName                                                   ' 2005/08/03 ADD T.SUZUKI
    pintAGRE = 0                                                                ' 2005/08/03 ADD T.SUZUKI
    LimitationStartDate = Null                                                  ' 2006/01/31 ADD H.TAJIMA
End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面上のテキストボックスをクリアする
'        MODULE_ID        :CLR_RTN2
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub CLR_RTN2()                                                         'DELETE 20050916 K.ISHIZAKA
Private Sub CLR_RTN2(ByVal intOpg As Integer)                                   'INSERT 20050916 K.ISHIZAKA

    Dim vRet  As Variant

    On Error Resume Next  'エラーがあっても処理続行

    With Me
        ' 受付番号コンボの再作成
'        .cmb_受付番号.Requery  DEL 2006/03/03 Tajima

'        Select Case .opg_処理選択.Value                                        'DELETE 20050916 K.ISHIZAKA
        Select Case intOpg                                                      'INSERT 20050916 K.ISHIZAKA
            Case 1  ' *** 新規 ***
                If pintOpenMode = 1 Then
                    ' 取込あり
'                    .cmb_ヤードコード = Null                                   'DELETE 20050916 K.ISHIZAKA
                    .cmb_ヤードコード = "000000"                                'INSERT 20050916 K.ISHIZAKA
                    .txt_ヤード名 = Null
                    .txt_ヤード住所 = Null
'                    .txt_地主コード = Null  ' 2005/08/03 DEL T.SUZUKI
'                    .txt_地主名 = Null                                         ' 2006/09/22 DEL SHIBAZAKI
'                    .txt_物件番号 = Null                                       ' 2006/09/22 DEL SHIBAZAKI
'                    .cmb_コンテナ番号 = Null                                   'DELETE 20050916 K.ISHIZAKA
                    .cmb_コンテナ番号 = "000000"                                'INSERT 20050916 K.ISHIZAKA

                Else
                    ' 取込なし
'                    .cmb_コンテナ番号 = Null                                   'DELETE 20050916 K.ISHIZAKA
                    .cmb_コンテナ番号 = "000000"                                'REP 20050916 K.ISHIZAKA
'                    .chk_予約受付.Value = 0    ' 未選択状態  DEL 2006/03/03 Tajima
'                    .cmb_受付番号 = Null                     DEL 2006/03/03 Tajima
                End If

            Case 2, 3 ' *** 修正 OR 削除 ***
                ' ヤードコード、コンテナ番号ともに選択状態とする。
                .txt_受注契約番号 = Null
        End Select

        ' 受注契約番号を初期化
        pstrJUCHUKEIYAKU_NO = ""  ' 2005/08/03 REP T.SUZUKI

        .cmb_受付担当者コード = ""
        .txt_受付担当者名 = ""
        .txt_段区分名 = Null        ' 2005/08/03 REP T.SUZUKI
        .txt_コンテナ種別名 = Null  ' 2005/08/03 REP T.SUZUKI
        
        .txt_種別 = Null  ' 2013/05/09 M.HONDA INS
        .txt_種別.BackColor = 14680063 ' 2013/05/09 M.HONDA INS

        '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'        .txt_予約名 = Null
'        .txt_予約番号 = Null

        '2005/09/22 SHIBAZAKI 直近修繕表示
        .txt_直近修繕 = Null
        .txt_クレーム = Null
        
        .txt_初回契約日 = Null
        .txt_契約開始日 = Null
        .txt_契約満了日 = Null

        .txt_解約日付 = Null
        .txt_解約日付.Tag = ""  ' 2005/08/03 ADD T.SUZUKI
'        .chk_コンテナ状態 = 0   ' 2005/08/03 DEL T.SUZUKI

' 2005/08/03 ADD T.SUZUKI Start
        .txt_解約受付日 = Null
'        .txt_解約受付日 = Format(Now, "YYYY/MM/DD")
        .cmb_解約区分 = Null
        .txt_解約区分 = Null
        .cmb_解約理由 = Null
        .txt_解約理由 = Null
        .cmb_解約受付担当者コード = Null
        .txt_解約受付担当者名 = Null
' 2005/08/03 ADD T.SUZUKI End

'2012/04/11 INS M.HONDA
        .cmb_スロープ = Null
        .txt_スロープ = Null

        .cmb_鍵区分 = 0
        Call cmb_鍵区分_AfterUpdate
        .txt_解除番号 = Null
        .chk_書類1 = -1

        .cmb_保証区分 = 4
        Call cmb_保証区分_AfterUpdate
        
        .cmb_保証会社コード = Null              '2009/04/01 INS KIHARA
        Call cmb_保証会社コード_AfterUpdate     '2009/04/01 INS KIHARA

        .txt_承認日 = Null
        .txt_承認番号 = Null
        
        .cmb_請求書発行区分 = Null                                              'INSERT 2010/04/10 K.ISHIZAKA
        Call cmb_請求書発行区分_AfterUpdate                                     'INSERT 2010/04/10 K.ISHIZAKA

        .txt_月額料金 = 0
        .txt_消費税 = 0
        .txt_合計 = 0
        .txt_預り保証金 = 0
        .txt_契約備考 = Null

'2006/05/07 ADD tajima Start
        .txt_スペアキー備考 = Null
        .txt_貸出本数 = 0
        .txt_セキュリティカード枚数 = Null
'2006/05/07 ADD tajima End

        ' タブのページ設定
        If Me.tab_顧客情報.VALUE <> 0 Then
            Me.tab_顧客情報.VALUE = 0
        End If

        .txt_顧客カナ = Null   ' 2005/10/19 REP T.SUZUKI
        .cmb_顧客コード = "000000"                                              'INSERT 20050916 K.ISHIZAKA

        ' 【顧客情報】
        .txt_顧客名 = Null
        .txt_顧客名カナ = Null
        .txt_代表者名 = Null
        .txt_代表者名カナ = Null
        .txt_郵便番号 = Null
        .txt_住所1 = Null
        .txt_住所2 = Null
        .txt_住所3 = Null
        .txt_メールアドレス = Null
        .txt_電話番号 = Null
        .txt_携帯番号 = Null
        .txt_FAX番号 = Null
        .txt_締日 = "31"   ' デフォルト：31

        .cmb_顧客区分 = Null
        .txt_顧客区分名 = Null
        '' 2013/09/19 M.HONDA INS
        ''.cmb_集金区分 = "07"
        .cmb_集金区分 = "01"
        '' 2013/09/19 M.HONDA INS
        
        Call cmb_集金区分_AfterUpdate
        .txt_顧客備考 = Null

        ' 【その他連絡先】
        .txt_連絡先担当名カナ = Null
        .txt_連絡先担当者名 = Null
        .txt_連絡先郵便番号 = Null
        .txt_連絡先住所1 = Null
        .txt_連絡先住所2 = Null
        .txt_連絡先住所3 = Null
        .txt_連絡先電話番号 = Null
        .txt_連絡先携帯番号 = Null
        .txt_連絡先FAX = Null

        '【緊急連絡先】                                                         'INSERT START 2011/04/11 K.ISHIZAKA
        .txt_RCPT_KREN_NAME = Null
        .txt_RCPT_KREN_KANA = Null
        .txt_RCPT_KREN_ZOKUGARA = Null
        .txt_RCPT_KREN_YUBINO = Null
        .txt_RCPT_KREN_ADR_1 = Null
        .txt_RCPT_KREN_ADR_2 = Null
        .txt_RCPT_KREN_ADR_3 = Null
        .txt_RCPT_KREN_TEL = Null
        .txt_RCPT_KREN_KEITAI = Null
        .txt_RCPT_KREN_FAX = Null                                               'INSERT END   2011/04/11 K.ISHIZAKA

        .txt_RCPT_BIKE_SHASYU = Null
        .txt_RCPT_BIKE_NUMBER = Null

'        .opg_顧客情報取得選択.Value = 1  DEL 2006/03/03 Tajima
        .txt_キャンペーン適用満了日 = Null      'ADD 2006/03/03 Tajima
        .cmd_キャンペーン内容.Enabled = False   'ADD 2006/03/03 Tajima
        .chk_入力する.VALUE = 0
    End With

    ' コントロールマスタより取得 -----------------------------------------
    ' 消費税切替日
    vRet = Nz(DLookup("CONT_HENRIT", "dbo_CONT_MAST", "CONT_KEY = 1"), "")
    pvarHenrit = vRet

    ' 切替前消費税率
    vRet = Nz(DLookup("CONT_MAERIT", "dbo_CONT_MAST", "CONT_KEY = 1"), "")
    pvarMaerit = vRet

    ' 切替後消費税率
    vRet = Nz(DLookup("CONT_ATORIT", "dbo_CONT_MAST", "CONT_KEY = 1"), "")
    pvarAtorit = vRet
    ' ---------------------------------------------------------------------

    ' システム日付
    pvarDATE = Format(Now, P_FORMAT_YMD)      ' 2006/01/05 ADD T.SUZUKI
'    pvarDATE = Format(Now, "yyyy/mm/dd")     ' 2006/01/05 DEL T.SUZUKI

    ' 変更フラグの初期化
    pblnChangeFlag_FVS600 = False
    pblnMoneyChgFlg = False                   ' 2005/08/03 ADD T.SUZUKI
    Me.Caption = pstrFormName                 ' 2005/08/03 ADD T.SUZUKI
    pintAGRE = 0                              ' 2005/08/03 ADD T.SUZUKI
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CLR_RTN3
'        CREATE_DATE      :
'
'        処理内容         ：【顧客情報】、【その他連絡先】の内容をクリアする。
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CLR_RTN3(intMode As Integer, Optional intKey As Integer = 0)

    With Me
        Select Case intMode
            Case 0
                ' タブのページ設定
                If Me.tab_顧客情報.VALUE <> 0 Then
                    Me.tab_顧客情報.VALUE = 0
                End If

                ' 【顧客情報】
                If intKey <> 1 Then
                    .cmb_顧客コード = "000000"  ' 2005/09/12 REP T.SUZUKI
                    .txt_顧客カナ = Null        ' 2005/10/19 REP T.SUZUKI
                End If

                .txt_顧客名 = Null
                .txt_顧客名カナ = Null
                .txt_代表者名 = Null
                .txt_代表者名カナ = Null
                .txt_郵便番号 = Null
                .txt_住所1 = Null
                .txt_住所2 = Null
                .txt_住所3 = Null
                .txt_メールアドレス = Null
                .txt_電話番号 = Null
                .txt_携帯番号 = Null
                .txt_FAX番号 = Null
                .txt_締日 = 31
                .cmb_顧客区分 = Null
                .txt_顧客区分名 = Null
                '' 2013/09/19 M.HONDA INS
                ''.cmb_集金区分 = "07"
                .cmb_集金区分 = "01"
                '' 2013/09/19 M.HONDA INS
                
                Call cmb_集金区分_AfterUpdate
                .txt_顧客備考 = Null

            Case 1
                ' 【その他連絡先】
                .chk_入力する.VALUE = 0            ' 未選択状態にする
                .txt_連絡先担当名カナ = Null
                .txt_連絡先担当者名 = Null
                .txt_連絡先郵便番号 = Null
                .txt_連絡先住所1 = Null
                .txt_連絡先住所2 = Null
                .txt_連絡先住所3 = Null
                .txt_連絡先電話番号 = Null
                .txt_連絡先携帯番号 = Null
                .txt_連絡先FAX = Null
                
            Case 2                                                              'INSERT START 2011/04/11 K.ISHIZAKA
                '【緊急連絡先】
                .txt_RCPT_KREN_NAME = Null
                .txt_RCPT_KREN_KANA = Null
                .txt_RCPT_KREN_ZOKUGARA = Null
                .txt_RCPT_KREN_YUBINO = Null
                .txt_RCPT_KREN_ADR_1 = Null
                .txt_RCPT_KREN_ADR_2 = Null
                .txt_RCPT_KREN_ADR_3 = Null
                .txt_RCPT_KREN_TEL = Null
                .txt_RCPT_KREN_KEITAI = Null
                .txt_RCPT_KREN_FAX = Null                                       'INSERT END   2011/04/11 K.ISHIZAKA
                
        End Select
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmLockControl
'        CREATE_DATE      :
'
'引数 ０  使用可
'     １  使用不可
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CmLockControl(ByVal ContNM As Object, ByVal iMode As Integer)

    Dim objControl As Object

    On Error Resume Next                                                 ' 2006/01/05 ADD T.SUZUKI

    Set objControl = ContNM

    If iMode = 0 Then
        objControl.Enabled = True
        objControl.Locked = False
    Else
        objControl.Enabled = False
        objControl.Locked = True
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subSetColor
'        CREATE_DATE      :
'
'        引       数      ：ContNM = コントロール名
'                           lngColor = 背景色
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetColor(ByVal ContNM As Object, ByVal lngColor As Long)

    Dim objControl As Object

    Set objControl = ContNM

    objControl.BackColor = lngColor
End Sub

'==============================================================================*
'
'        MODULE_NAME      :入力区分によるコントロールのロックを設定
'        MODULE_ID        :subLock_Etc
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subLock_Etc()

    With Me
'        Call CmLockControl(.cmb_受付担当者コード, 1)  ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_契約開始日, 1)        ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_契約満了日, 1)        ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.cmd_広告媒体, 1)          ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
        Call CmLockControl(.txt_解約日付, 1)
        Call CmLockControl(.txt_解約受付日, 1)            ' 2005/08/03 ADD T.SUZUKI
        Call CmLockControl(.cmb_解約区分, 1)              ' 2005/08/03 ADD T.SUZUKI
        Call CmLockControl(.cmb_解約理由, 1)              ' 2005/08/03 ADD T.SUZUKI
        
        '2012/04/11 INS M.HONDA
        Call CmLockControl(.cmb_スロープ, 1)
        Call CmLockControl(.txt_スロープ, 1)
        
        
        
        Call CmLockControl(.cmb_解約受付担当者コード, 1)  ' 2005/08/03 ADD T.SUZUKI
        .cmd_解約入力.Enabled = False                     ' 2005/08/03 ADD T.SUZUKI
'        Call CmLockControl(.chk_コンテナ状態, 1)         ' 2005/08/03 DEL T.SUZUKI
        Call CmLockControl(.cmb_鍵区分, 1)
        Call CmLockControl(.txt_解除番号, 1)
        Call CmLockControl(.chk_書類1, 1)
'        Call CmLockControl(.chk_書類2, 1)            ' 2006/03/03 Add TAJIMA 受付対応による入力項目削除
'        Call CmLockControl(.txt_戻り日, 1)           ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_番号, 1)             ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.cmb_保証区分, 1)         ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
        Call CmLockControl(.txt_承認日, 1)
        Call CmLockControl(.txt_承認番号, 1)
        Call CmLockControl(.cmb_請求書発行区分, 1)                              'INSERT 2010/04/10 K.ISHIZAKA
'        Call CmLockControl(.txt_月額料金, 1)         ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_消費税, 1)           ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_初回金額, 1)         ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_初回鍵代, 1)         ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_初回消費税, 1)       ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        Call CmLockControl(.txt_預り保証金, 1)       ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
        Call CmLockControl(.txt_契約備考, 1)
'        .txt_顧客カナ.Enabled = False                ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'        .cmb_顧客コード.Enabled = False              ' 2006/02/18 Add TAJIMA 受付対応による入力絶対無効化
'2006/05/07 ADD tajima Start
        Call CmLockControl(.txt_スペアキー備考, 1)
        Call CmLockControl(.txt_貸出本数, 1)
'2006/05/07 ADD tajima End

        '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'        Call CmLockControl(.txt_予約名, 1)
'        Call CmLockControl(.txt_予約番号, 1)

        ' 顧客情報取得選択オプションボタンを使用不可に設定
'        .opt_新規顧客.Enabled = False   DEL 2006/03/03 Tajima
'        .opt_既存顧客.Enabled = False  DEL 2006/03/03 Tajima

        ' 顧客マスタ修正ボタン、【顧客情報】、【その他連絡先】を使用不可に設定
        Call subLock_UserItem(0, 1)

        ' 【その他連絡先】の使用可否設定、背景色設定
        Call subTxtControl

        .cmd_保存.Enabled = True                      ' 2006/03/04 ADD T.SUZUKI
        .cmd_削除.Enabled = True                      ' 2006/03/04 ADD T.SUZUKI
' 2006/03/04 DEL T.SUZUKI Start
'        Select Case .opg_処理選択.Value
'            Case 3     ' *** 削除 ***
'                .cmd_保存.Enabled = False
'                .cmd_削除.Enabled = True
'
'            Case Else  ' *** 上記以外 ***
'                .cmd_保存.Enabled = True
'                .cmd_削除.Enabled = False
'        End Select
' 2006/03/04 DEL T.SUZUKI End

        ' 予約追加ボタン
        Me.cmd_YoyakuAdd.Enabled = False   ' 2006/01/05 ADD T.SUZUKI
        ' 契約変更ボタン
        Me.cmd_KeiyakuRep.Enabled = False  ' 2006/01/05 ADD T.SUZUKI
        ' 受付参照ボタン
        Me.cmd_UketsukeView.Enabled = False ' 2006/03/18 ADD Tajima
       
'2006/07/28 DEL  Me.cmd_カードメンテ.Enabled = False ' 2006/05/07 ADD tajima
       
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subLock_Main
'        CREATE_DATE      :
'
'        処理内容         ：処理モードにより、ヤードコード、コンテナ番号の使用可否を設定
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subLock_Main()

    With Me
        Select Case pintOpenMode
            Case 1, 3, 9                                    ' 2006/01/26 ADD T.SUZUKI
                ' *** 契約新規・取込あり／契約参照 ***
'            Case 1, 3                                      ' 2006/01/26 DEL T.SUZUKI
'                ' *** 契約新規・取込あり／契約参照 ***
                Call CmLockControl(Me.cmb_ヤードコード, 1)
                Call CmLockControl(Me.cmb_コンテナ番号, 1)
            Case Else
                ' *** 上記以外 ***
                Call CmLockControl(Me.cmb_ヤードコード, 0)
                Call CmLockControl(Me.cmb_コンテナ番号, 0)
        End Select
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :画面の各コントロールの制御
'        MODULE_ID        :subLock_formItem
'        CREATE_DATE      :
'        処理内容         ：処理選択状態により、画面の各コントロールの制御を行う
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subLock_formItem()
    Dim ctrl                As Variant                                          'INSERT 2011/04/11 K.ISHIZAKA

    With Me
        If pintOpenMode = 2 Then
            ' *** 契約修正 ***
            Call CmLockControl(.cmb_受付担当者コード, 0)
    
            If pvarKYDATE < DATE Then
                ' 地主契約マスタ.解約日付 < システム日付のとき
                ' → 解約日のみ入力可能とする
                Call CmLockControl(.txt_解約日付, 0)
                Call CmLockControl(.cmb_解約区分, 0)    '2008/03/23 add tajima 解約区分選択何時でもOK
    
                If Nz(.txt_解約日付, "") = "" Then
                    Call CmLockControl(.txt_解約受付日, 1)
                    'Call CmLockControl(.cmb_解約区分, 1)  2008/03/23 del tajima
                    Call CmLockControl(.cmb_解約理由, 1)
                    Call CmLockControl(.cmb_解約受付担当者コード, 1)
                Else
                    Call CmLockControl(.txt_解約受付日, 0)
                    'Call CmLockControl(.cmb_解約区分, 0)  2008/03/23 del tajima
                    Call CmLockControl(.cmb_解約理由, 0)
                    Call CmLockControl(.cmb_解約受付担当者コード, 0)
                End If
                .cmd_解約入力.Enabled = True  ' 修正時は使用可能とする
    
                Call CmLockControl(.chk_書類1, 1)
                Call CmLockControl(.txt_契約備考, 1)
                '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'                Call CmLockControl(.txt_予約名, 1)
'                Call CmLockControl(.txt_予約番号, 1)
'2006/05/07 ADD tajima Start
                Call CmLockControl(.txt_スペアキー備考, 1)
'2006/05/07 ADD tajima End

                Call CmLockControl(.txt_スロープ, 1)
                Call CmLockControl(.cmb_スロープ, 1)
            Else
                ' 上記以外
                Call CmLockControl(.txt_解約日付, 0)
    
                If Nz(.txt_解約日付, "") = "" Then
                    Call CmLockControl(.txt_解約受付日, 1)
                    'Call CmLockControl(.cmb_解約区分, 1)   2008/03/23 del tajima
                    Call CmLockControl(.cmb_解約理由, 1)
                    Call CmLockControl(.cmb_解約受付担当者コード, 1)
                Else
                    Call CmLockControl(.txt_解約受付日, 0)
                    'Call CmLockControl(.cmb_解約区分, 0)      2008/03/23 del tajima
                    Call CmLockControl(.cmb_解約理由, 0)
                    Call CmLockControl(.cmb_解約受付担当者コード, 0)
                End If
                .cmd_解約入力.Enabled = True  ' 修正時は使用可能とする
    
'2006/05/10 ADD tajima Start
                Call CmLockControl(.cmb_鍵区分, 0)
                Call CmLockControl(.txt_解除番号, 0)
                Call CmLockControl(.txt_貸出本数, 0)
'2006/05/10 ADD tajima End
                Call CmLockControl(.chk_書類1, 0)
                Call CmLockControl(.txt_契約備考, 0)
'2006/05/07 ADD tajima Start
                Call CmLockControl(.txt_スペアキー備考, 0)
'2006/05/07 ADD tajima End
                '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'                Call CmLockControl(.txt_予約名, 0)
'                Call CmLockControl(.txt_予約番号, 0)
                
                Call CmLockControl(.txt_承認日, 0)  '2006/05/09 Add Tajima
                Call CmLockControl(.txt_承認番号, 0) '2006/05/09 Add Tajima
                
                Call CmLockControl(.txt_スロープ, 0)
                Call CmLockControl(.cmb_スロープ, 0)
            
                Dim rCONT_BUMOC As String
                rCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")

                If rCONT_BUMOC = "1" Then
                    Call CmLockControl(.cmb_請求書発行区分, 0)
                Else
                    Call CmLockControl(.cmb_請求書発行区分, 1)
                End If
            
            
            End If
    
            ' 顧客情報取得選択オプションボタン
            Call subLock_UserItem(1, 1) 'Add 2006/03/03 Tajima
            
            ' 【その他連絡先】の使用可否設定、背景色設定
            Call subTxtControl
            
            ' 【緊急連絡先】の使用可否設定（基本は使用可でないといけないので、わざと背景色は変更しない＝気づいてパッチを充てるべきだから）
            For Each ctrl In pge_緊急連絡先.Controls                            'INSERT START 2011/04/11 K.ISHIZAKA
                Select Case UCase(TypeName(ctrl))
                Case "TEXTBOX"
                    ctrl.Enabled = isRCPT_TRAN
                    ctrl.Locked = Not isRCPT_TRAN
                Case "COMMANDBUTTON"
                    ctrl.Enabled = isRCPT_TRAN
                End Select
            Next                                                                'INSERT END   2011/04/11 K.ISHIZAKA
            
            ' 予約追加ボタン
            Me.cmd_YoyakuAdd.Enabled = True               ' 2006/01/05 ADD T.SUZUKI
            ' 受付参照ボタン
            Me.cmd_UketsukeView.Enabled = True            ' 2006/03/18 ADD Tajima

'2006/07/28 DEL  Me.cmd_カードメンテ.Enabled = True            ' 2006/05/07 ADD Tajima

        Else
            ' *** 契約参照 ***
            Call CmLockControl(.cmb_受付担当者コード, 1)
            Call CmLockControl(.txt_契約開始日, 1)
            Call CmLockControl(.txt_契約満了日, 1)
            Call CmLockControl(.txt_解約日付, 0)
            Call CmLockControl(.cmb_解約区分, 0)    '2008/03/23 add tajima 解約区分選択何時でもOK
            Call CmLockControl(.cmb_請求書発行区分, 1)                          'INSERT 2010/04/10 K.ISHIZAKA

            If Nz(.txt_解約日付, "") = "" Then
                Call CmLockControl(.txt_解約受付日, 1)
                'Call CmLockControl(.cmb_解約区分, 1)  2008/03/23 del tajima
                Call CmLockControl(.cmb_解約理由, 1)
                Call CmLockControl(.cmb_解約受付担当者コード, 1)
            Else
                Call CmLockControl(.txt_解約受付日, 0)
                'Call CmLockControl(.cmb_解約区分, 0)  2008/03/23 del tajima
                Call CmLockControl(.cmb_解約理由, 0)
                Call CmLockControl(.cmb_解約受付担当者コード, 0)
            End If
            .cmd_解約入力.Enabled = True  ' 修正時は使用可能とする

            Call CmLockControl(.txt_スロープ, 1)
            Call CmLockControl(.cmb_スロープ, 1)

            Call CmLockControl(.cmb_鍵区分, 1)
            Call CmLockControl(.txt_解除番号, 1)
            Call CmLockControl(.chk_書類1, 1)
            Call CmLockControl(.cmb_保証区分, 1)
            Call CmLockControl(.cmb_保証会社コード, 1)    ' 2009/04/01 INS KIHARA
            Call CmLockControl(.txt_承認日, 1)
            Call CmLockControl(.txt_承認番号, 1)
            Call CmLockControl(.txt_月額料金, 1)
            Call CmLockControl(.txt_消費税, 1)
            Call CmLockControl(.txt_契約備考, 1)
'2006/05/07 ADD tajima Start
            Call CmLockControl(.txt_スペアキー備考, 1)
            Call CmLockControl(.txt_貸出本数, 1)
'2006/05/07 ADD tajima End
            '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'            Call CmLockControl(.txt_予約名, 1)
'            Call CmLockControl(.txt_予約番号, 1)

            ' 【顧客情報】、【その他連絡先】を使用不可能に設定（既存）
            Call subLock_UserItem(0, 1)

            ' 【その他連絡先】の使用可否設定、背景色設定
            Call subTxtControl
            
            ' 【緊急連絡先】の使用可否設定（基本は使用可でないといけないので、わざと背景色は変更しない＝気づいてパッチを充てるべきだから）
            For Each ctrl In pge_緊急連絡先.Controls                            'INSERT START 2011/04/11 K.ISHIZAKA
                Select Case UCase(TypeName(ctrl))
                Case "TEXTBOX"
                    ctrl.Enabled = isRCPT_TRAN
                    ctrl.Locked = Not isRCPT_TRAN
                Case "COMMANDBUTTON"
                    ctrl.Enabled = isRCPT_TRAN
                End Select
            Next                                                                'INSERT END   2011/04/11 K.ISHIZAKA

            ' 予約追加ボタン
            Me.cmd_YoyakuAdd.Enabled = False
            ' 契約変更ボタン
            Me.cmd_KeiyakuRep.Enabled = False
            ' 受付参照ボタン
            Me.cmd_UketsukeView.Enabled = False            ' 2006/03/18 ADD Tajima
            
'2006/07/28 DEL Me.cmd_カードメンテ.Enabled = False            ' 2006/05/07 ADD Tajima
        End If
     
    End With
End Sub

'DEL 2006/03/03 Tajima Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :subLock_YoyakuItem
''        CREATE_DATE      :
''
''        処理内容         ：予約受付チェックボックス、受付番号の使用可否の設定
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub subLock_YoyakuItem()
'
'    With Me
'        If .opg_処理選択.Value = 1 Then
'            ' *** 新規 ***
'            If .chk_予約受付.Enabled = False Then
'                .chk_予約受付.Enabled = True
'            End If
'
'            If .chk_予約受付 = -1 Then
'                Call CmLockControl(.cmb_受付番号, 0)  ' 予約受付からチェックボックス：ON
'            Else
'                Call CmLockControl(.cmb_受付番号, 1)  ' 予約受付からチェックボックス：OFF
'            End If
'
'        Else
'            ' *** 修正 OR 削除
'' 2005/08/03 REP T.SUZUKI Start
'            If .cmb_ヤードコード.Enabled = True Then
'                .cmb_ヤードコード.SetFocus            ' ヤードコードにセットフォーカス
'            End If
'            .cmb_ヤードコード.SetFocus            ' ヤードコードにセットフォーカス
'' 2005/08/03 REP T.SUZUKI End
'            .chk_予約受付.Enabled = False         ' 予約受付から
'            Call CmLockControl(.cmb_受付番号, 1)  ' 受付番号
'        End If
'    End With
'End Sub
'DEL 2006/03/03 Tajima Start

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :subLock_UserItem
'        CREATE_DATE      :
'
'        処理内容         ：顧客マスタ修正ボタン、【顧客情報】、【その他連絡先】の制御を行う。
'        引    数         ：intMode   ：0 = 新規顧客／1 = 既存顧客
'                           intControl：0 = 使用可能／1：使用不可
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subLock_UserItem(intMode As Integer, intControl As Integer)

    Dim lngBackColor  As Long
    Dim blnFLG        As Boolean

    With Me
        ' 顧客マスタ修正ボタン制御
        
' DEL 2006/03/03 Tajima Start
'        Select Case intMode
'            Case 0  ' *** 新規顧客 ***
'                .cmd_顧客マスタ修正.Enabled = False
'                .txt_顧客カナ.Enabled = False    ' 顧客カナ   ' 2005/10/19 ADD T.SUZUKI
'                .cmb_顧客コード.Enabled = False  ' 顧客コード
'
'            Case 1  ' *** 既存顧客 ***
'                .cmd_顧客マスタ修正.Enabled = True
'                .txt_顧客カナ.Enabled = True     ' 顧客名カナ   ' 2005/10/19 ADD T.SUZUKI
'
'' 2005/08/03 REP T.SUZUKI Start
'                If pintOpenMode = 3 Then
'                    .cmb_顧客コード.Enabled = False  ' 顧客コード
'                Else
'                    .cmb_顧客コード.Enabled = True   ' 顧客コード
'                End If
''                .cmb_顧客コード.Enabled = True   ' 顧客コード
'' 2005/08/03 REP T.SUZUKI End
'        End Select
' DEL 2006/03/03 Tajima End

        ' 項目の背景色決定
        If intControl = 0 Then
            lngBackColor = plngInputColor
            blnFLG = True
        Else
            lngBackColor = plngLockColor
            blnFLG = False
        End If

        ' 【顧客情報】
        '  --- > 使用可否設定
        Call CmLockControl(Me.txt_顧客名カナ, intControl)
        Call CmLockControl(Me.txt_顧客名, intControl)
        Call CmLockControl(Me.txt_代表者名カナ, intControl)
        Call CmLockControl(Me.txt_代表者名, intControl)
        Call CmLockControl(Me.txt_郵便番号, intControl)

        ' 郵便番号変換コマンドボタン
        .cmd_Yubin.Enabled = blnFLG

        Call CmLockControl(.txt_住所1, intControl)
        Call CmLockControl(.txt_住所2, intControl)
        Call CmLockControl(.txt_住所3, intControl)
        Call CmLockControl(.txt_メールアドレス, intControl)
        Call CmLockControl(.txt_電話番号, intControl)
        Call CmLockControl(.txt_携帯番号, intControl)
        Call CmLockControl(.txt_FAX番号, intControl)
        Call CmLockControl(.cmb_顧客区分, intControl)
        Call CmLockControl(.cmb_集金区分, intControl)
' ↓↓↓↓↓名称は常に使用不可
'        Call CmLockControl(.txt_顧客区分名, intControl)
'        Call CmLockControl(.txt_集金区分名, intControl)
' ↑↑↑↑↑名称は常に使用不可
        Call CmLockControl(.txt_締日, intControl)
        Call CmLockControl(.txt_顧客備考, intControl)
    
        '  --- > 背景色設定
        Call subSetColor(.txt_顧客名カナ, lngBackColor)
        Call subSetColor(.txt_顧客名, lngBackColor)
        Call subSetColor(.txt_代表者名カナ, lngBackColor)
        Call subSetColor(.txt_代表者名, lngBackColor)
        Call subSetColor(.txt_郵便番号, lngBackColor)
        Call subSetColor(.txt_住所1, lngBackColor)
        Call subSetColor(.txt_住所2, lngBackColor)
        Call subSetColor(.txt_住所3, lngBackColor)
        Call subSetColor(.txt_メールアドレス, lngBackColor)
        Call subSetColor(.txt_電話番号, lngBackColor)
        Call subSetColor(.txt_携帯番号, lngBackColor)
        Call subSetColor(.txt_FAX番号, lngBackColor)
' ↓↓↓↓↓デフォルト色を活かす
'        Call subSetColor(.cmb_顧客区分, lngBackColor)
'        Call subSetColor(.txt_顧客区分名, lngBackColor)
'        Call subSetColor(.cmb_集金区分, lngBackColor)
'        Call subSetColor(.txt_集金区分名, lngBackColor)
' ↑↑↑↑↑デフォルト色を活かす
        Call subSetColor(.txt_締日, lngBackColor)
        Call subSetColor(.txt_顧客備考, lngBackColor)
    
        ' 【その他連絡先】
        '  --- > 使用可否設定
        .chk_入力する.Enabled = blnFLG
    End With
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncControlCrear
'        機能             :ｺﾝﾎﾞﾎﾞｯｸｽの内容が空の場合、名称ﾃｷｽﾄも空にする
'        IN               :ctlCmb = ｺﾝﾎﾞﾎﾞｯｸｽ / ctlTxt = ﾃｷｽﾄ
'        OUT              :ｺﾝﾎﾞﾎﾞｯｸｽが空だったかの判断
'        CREATE           :2005/08/03
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncControlCrear(ByRef ctlCmb As ComboBox, ByRef ctlTxt As TextBox) As Boolean

    If Nz(ctlCmb, "") = "" Then
        ctlTxt = ""
        Exit Function
    Else
        fncControlCrear = True
    End If
End Function

' <<<入力チェック系>>>
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CmKeyPress_For_D
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CmKeyPress_For_D(ojCtrl As Control, KeyAscii As Integer, inLen As Integer)

    On Error GoTo CmKeyPress_For_Dec_Err

    Dim stValue As String

    With ojCtrl
        Select Case KeyAscii
            Case 8, 9, 10, 13                               ' 8:BS 9:TAB 10:LF 13:CR
            Case Asc("0") To Asc("9"), Asc("."), Asc("-")   ' 0～9迄と .と - を許可する
            
                If Len(.Text) > 1 Then
                  stValue = CStr(Abs(CLng(Format$(.Text, "#0.00"))))
                End If

                If Len(stValue) > (inLen - 1) And .SelLength = 0 Then
                    KeyAscii = 0
                End If
            Case Else
                KeyAscii = 0
        End Select
    End With

    Exit Sub

CmKeyPress_For_Dec_Err:
  KeyAscii = 0
End Sub

'==============================================================================*
'
'        MODULE_NAME      :小数入力項目の入力チェック
'        MODULE_ID        :fncTxtInputCheck
'        Parameter        :ctlTxt      = 対象となるテキストボックス
'                          varMaxValue = 最大値
'                          intLength   = 長さ
'        CREATE_DATE      :2005/12/15
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTxtInputCheck(ByRef ctlTxt As TextBox, _
                                  ByVal varMaxValue As Variant, ByVal intLength As Integer) As Boolean

    fncTxtInputCheck = False

    If IsNull(ctlTxt) = True Then
        fncTxtInputCheck = True: Exit Function
    End If

'    Select Case intKbn
'        Case 0
            If (ctlTxt > varMaxValue Or _
                Len(Replace(Replace(ctlTxt, ".", ""), "-", "")) > intLength) Then
                Exit Function
            End If

'        Case 1
'            ' *** フォームのエラーイベントでのエラートラップのとき
'            If (ctlTxt > varMaxValue Or _
'                Len(Replace(Replace(ctlTxt.Text, ".", ""), "-", "")) > intLength) Then
'                Exit Function
'            End If
'    End Select
    fncTxtInputCheck = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub dateCheck(txtTarget As Variant, ByRef Cancel As Integer)

    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, strMsg_err
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :文字項目のチェックを一括でする。
'        MODULE_ID        :CmBCheck_Char
'        CREATE_DATE      :
'
'返り値は、
'２：桁数が溢れている時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

    On Error GoTo err_rtn

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
       CmBCheck_Char = 2
       GoTo Exit_rtn
    End If

    CmBCheck_Char = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmBCheck_Char = 9
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :日付項目のチェックを一括でする
'        MODULE_ID        :CmBCheck_Date
'        CREATE_DATE      :
'
'引数
'oText 対象項目
'iLen  桁数(省略時は、"yyyy/mm/dd"として10バイト)
'返り値は、
'２：桁数が溢れている時
'３：日付と判断できない時
'９：それ以外のエラー
'０：すべてのチェックがＯＫの時
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Date(ByVal oText As Variant, Optional ByVal ilen As Variant) As Integer

    Dim sYear   As String
    Dim intLen  As Integer

    On Error GoTo err_rtn

    CmBCheck_Date = 2

    If IsMissing(ilen) = False Then
        intLen = Nz(CmNumeric(ilen), 0)
    Else
        intLen = 10
    End If

    oText = Nz(oText, "")

    ' 桁数チェック
    If CmLenb(oText) > intLen Then
        Exit Function
    End If

    CmBCheck_Date = 3

    If Nz(oText, "") = "" Then
        Exit Function
    End If

    ' 論理チェック
    If IsDate(oText) = False Then
        Exit Function
    End If

    sYear = Format(oText, "yyyy")

    If CmLenb(sYear) <> 4 Then
        Exit Function
    End If

    CmBCheck_Date = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    CmBCheck_Date = 9
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :数値項目のチェックを一括でする。
'        MODULE_ID        :CmBCheck_Num
'        CREATE_DATE      :
'
  '返り値は、
  '１：数値と判断できない時
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時
    
  '引数 oText ←判定する値
  '     iLen  ←整数桁
  '     isLen ←小数桁(省略された時は、小数点桁数無し)
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, _
                              Optional ByVal isLEN As Variant) As Integer

    Dim vINT    As Variant
    Dim vsINT   As Variant
    Dim i       As Integer
      
    Dim dTEXT   As Double
    Dim dFROM   As Double
    Dim dTO     As Double
    Dim vRet    As Variant

    On Error GoTo err_rtn

    CmbCheck_Num = 9

    If Nz(CmNumeric(oText), "") = "" Then              ' 数値と判断できないのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    If CmChkCher(CStr(Nz(oText)), "A") = False Then    ' 全角文字が入力されているのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    ' 小数点の値が設定されている時のチェック
    vRet = Nz(InStr(Abs(oText), "."), 0)

    If IsMissing(isLEN) = False Then
        If vRet <> 0 Then
            If isLEN < CmLenb(Abs(oText)) - vRet Then  ' 小数点桁数が設定桁数を超えているエラー
                CmbCheck_Num = 2
                GoTo Exit_rtn
            End If
        End If

    Else
        If CmLenb(Abs(CmNumeric(oText))) > ilen Then   ' 文字数超過のためエラー
            CmbCheck_Num = 2
            GoTo Exit_rtn
        End If
    
        If vRet <> 0 Then                              ' 小数点の入力があるのでエラー
            CmbCheck_Num = 2
            GoTo Exit_rtn
        End If
    End If

    ' 桁チェック
    i = 0
    For i = 1 To ilen
        vINT = vINT & "9"
    Next

    If IsMissing(isLEN) = False Then
        vsINT = "."
        i = 0
        For i = 1 To isLEN
            vsINT = vsINT & "9"
        Next
    End If

    vINT = vINT & vsINT
    vINT = CmNumeric(vINT)

    dTEXT = CmNumeric(oText)
    dFROM = vINT * -1
    dTO = vINT
  
    If dFROM > dTEXT Or dTO < dTEXT Then               ' 論理範囲超過のためエラー
        CmbCheck_Num = 2
        GoTo Exit_rtn
    End If

    CmbCheck_Num = 0

Exit_rtn:
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmbCheck_Num = 9
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :
'Check_Dateでは'198/03といった形が変換されるための補助検査
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)

    Dim sYear  As String

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
'        MODULE_ID        :CmParEmpty
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmParEmpty(ByVal oEmpty As Variant, _
                            Optional ByVal oEmptyFull As Variant) As Variant

    Dim vText  As Variant

    If IsMissing(oEmptyFull) = False Then
        vText = oEmptyFull
    Else
        vText = 0
    End If

    If IsEmpty(oEmpty) = True Then
        CmParEmpty = vText
    Else
        CmParEmpty = oEmpty
    End If
End Function

' 2006/01/05 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :数値項目のチェックを一括でする。
''        MODULE_ID        :CHK_Numeric
''        CREATE_DATE      :
''
''返り値は、
''１：数値と判断できない時
''２：桁数が溢れている時
''９：それ以外のエラー
''０：すべてのチェックがＯＫの時
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function CHK_Numeric(ByVal oText As Variant, ByVal iLen As Integer) As Integer
'
'    On Error GoTo err_rtn
'
'    CHK_Numeric = 9
'
'    If Nz(CmNumeric(oText), "") = "" Then
'        CHK_Numeric = 1
'        GoTo Exit_rtn
'    End If
'
'    If CmLenb(oText) > iLen Then
'        CHK_Numeric = 2
'        GoTo Exit_rtn
'    End If
'
'    CHK_Numeric = 0
'
'Exit_rtn:
'    Exit Function
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    CHK_Numeric = 9
'    GoTo Exit_rtn
'End Function
' 2006/01/05 DEL T.SUZUKI End

' <<<DBAccess系>>>
'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約ファイルのデータ取得
'        MODULE_ID        :fncGetCARG_FILE
'        Parameter        :strACPTNO = 受注契約番号
'                          objRs     = レコードセットオブジェクト
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetCARG_FILE(strACPTNO As String, ByRef objRsCarg As Recordset) As Boolean

    Dim strSQL As String

    On Error GoTo fncGetCARG_FILE_Err

    fncGetCARG_FILE = False

    doCmd.Hourglass True

    strSQL = "SELECT    *  " & Chr(13)
    strSQL = strSQL & " FROM dbo_CARG_FILE " & Chr(13)           ' コンテナ契約ファイル
    strSQL = strSQL & " WHERE CARG_ACPTNO = '" & strACPTNO & "'" ' 受注契約番号
    Set objRsCarg = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)

    If objRsCarg.EOF = True Then
        objRsCarg.Close
        Set objRsCarg = Nothing
        GoTo fncGetCARG_FILE_Exit
    End If

    ' 戻り値設定
    fncGetCARG_FILE = True

fncGetCARG_FILE_Exit:
    doCmd.Hourglass False
    Exit Function

fncGetCARG_FILE_Err:
  MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
  Err.Clear
  GoTo fncGetCARG_FILE_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncGetYOUK_TRAN
'        CREATE_DATE      :2005/07/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetYOUK_TRAN() As Boolean

    Dim strSQL As String
    Dim objRs  As Recordset                                      ' 2006/01/05 ADD T.SUZUKI

    On Error GoTo err_rtn

    fncGetYOUK_TRAN = False

    doCmd.Hourglass True

    strSQL = "SELECT    YOUKT_UKNO,  " & Chr(13)      ' 受付番号
    strSQL = strSQL & " YOUKT_YCODE, " & Chr(13)      ' ヤード番号
    strSQL = strSQL & " YOUKT_NO, " & Chr(13)         ' コンテナ番号
    strSQL = strSQL & " YOUKT_UKKBN, " & Chr(13)      ' 受付区分
    strSQL = strSQL & " YOUKT_KUDATE, " & Chr(13)     ' 契約受付日
    strSQL = strSQL & " YOUKT_KKDATE, " & Chr(13)     ' 希望開始日
    ' 予約名
    ' 番号
    strSQL = strSQL & " YOUKT_BIKO, " & Chr(13)       ' 備考
    ' 顧客情報フリガナ
    ' 顧客情報コード
    strSQL = strSQL & " YOUKT_UCODE,  " & Chr(13)     ' 顧客ｺｰﾄﾞ
    strSQL = strSQL & " YOUKT_NAME, " & Chr(13)       ' 顧客名
    strSQL = strSQL & " YOUKT_KANA, " & Chr(13)       ' 顧客名ｶﾅ
    strSQL = strSQL & " YOUKT_TANM, " & Chr(13)       ' 代表者名
    strSQL = strSQL & " YOUKT_TAKA, " & Chr(13)       ' 代表者名ｶﾅ
    strSQL = strSQL & " YOUKT_YUBINO, " & Chr(13)     ' 郵便番号
    strSQL = strSQL & " YOUKT_ADR_1, " & Chr(13)      ' 住所１
    strSQL = strSQL & " YOUKT_ADR_2, " & Chr(13)      ' 住所２
    strSQL = strSQL & " YOUKT_ADR_3, " & Chr(13)      ' 住所３
    strSQL = strSQL & " YOUKT_MAIL, " & Chr(13)       ' メールアドレス          ' 2006/01/31 ADD T.SUZUKI
    strSQL = strSQL & " YOUKT_TEL, " & Chr(13)        ' 電話番号
    strSQL = strSQL & " YOUKT_KEITAI, " & Chr(13)     ' 携帯番号
    strSQL = strSQL & " YOUKT_FAX, " & Chr(13)        ' FAX番号
    strSQL = strSQL & " YOUKT_KKBN, " & Chr(13)       ' 顧客区分
' 2006/01/05 ADD T.SUZUKI Start
    strSQL = strSQL & " YOUKT_MOTO_ACPTNO, " & Chr(13)  ' 元受注契約番号
    strSQL = strSQL & " YOUKT_MOVEKBN, " & Chr(13)      ' 移動区分
    strSQL = strSQL & " YOUKT_UKDATE, " & Chr(13)       ' 受付日
    strSQL = strSQL & " YOUKT_UKTANTO "                 ' 受付担当者コード
' 2006/01/05 ADD T.SUZUKI End
    strSQL = strSQL & " FROM dbo_YOUK_TRAN " & Chr(13)
    strSQL = strSQL & " WHERE YOUKT_UKNO = '" & pstrUKNO & "'" ' 受付番号
    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)       ' 2006/01/05 ADD T.SUZUKI
'    Set pobjRS2 = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)    ' 2006/01/05 DEL T.SUZUKI

    If objRs.RecordCount = 0 Then                                  ' 2006/01/05 ADD T.SUZUKI
'    If pobjRS2.RecordCount = 0 Then                               ' 2006/01/05 DEL T.SUZUKI
        GoTo Exit_rtn
    End If

    ' 予約受付トランの読込データを画面にセット
    If fncSetYoyakuData(1, objRs) = False Then Exit Function       ' 2006/01/05 ADD T.SUZUKI
'    If fncSetYoyakuData(1) = False Then Exit Function             ' 2006/01/05 DEL T.SUZUKI

    ' 戻り値設定
    fncGetYOUK_TRAN = True

Exit_rtn:
    doCmd.Hourglass False
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing  ' 2006/01/05 ADD T.SUZUKI
'    Set pobjRS2 = Nothing                                         ' 2006/01/05 DEL T.SUZUKI
    Exit Function

err_rtn:
  MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
  Err.Clear
  GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :予約受付トランデータ表示処理
'        MODULE_ID        :fncSetYoyakuData
'        Parameter        :
'                          intmode = 契約新規取込あり(予約受付トランのデータを画面に設定)
'                          objRsYK = 契約修正
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetYoyakuData(ByVal intMode As Integer, Optional ByVal objRsYK As Recordset) As Boolean

    Dim vRet   As Variant
    Dim objRs  As Recordset

    On Error GoTo err_rtn

    fncSetYoyakuData = False

    Select Case intMode
        Case 1  ' *** 契約新規取込あり ***
            'コンテナ状況区分
            pvarContna = -1
            pvarKcode = ""

            With Me
' 2006/01/05 ADD T.SUZUKI Start
                With SAKI_YOUK_TRAN
                    .strMOTO_ACPTNO = Nz(objRsYK![YOUKT_MOTO_ACPTNO], "")       ' 元受注契約番号
                    .strMOVEKBN = Format(Nz(objRsYK![YOUKT_MOVEKBN], 0), "00")  ' 移動区分
                    .strUkDate = CStr(Nz(objRsYK![YOUKT_UKDATE], ""))           ' 受付日
                    .strUKTANTO = Nz(objRsYK![YOUKT_UKTANTO], "")               ' 受付担当者コード
                End With
' 2006/01/05 ADD T.SUZUKI Start

                ' 予約受付(選択状態に設定)
'                .chk_予約受付.Value = -1 DEL 2006/03/03 Tajima
'
'                ' 受付番号
'                .cmb_受付番号 = pstrUKNO DEL 2006/03/03 Tajima

                ' ヤードコード
                .cmb_ヤードコード = objRsYK![YOUKT_YCODE]
                ' ※cmb_ヤードコード_AfterUpdateにてクリア処理を行うように修正されている！！！
                Call cmb_ヤードコード_AfterUpdate

                ' 新規登録の場合、【cmb_コンテナ番号_AfterUpdate】処理にて『顧客コード』が必要なため、事前にセット
                ' 顧客コード
                .cmb_顧客コード = Format(Nz(objRsYK![YOUKT_UCODE], 0), "000000")

                ' コンテナ番号
                .cmb_コンテナ番号 = objRsYK![YOUKT_NO]
                pvarContna = objRsYK![YOUKT_NO]                       ' (フォームモジュール内で使用)
                Call cmb_コンテナ番号_AfterUpdate

                .txt_解約日付 = ""                      ' 解約日付

                ' 移動元データ設定処理
                If (Nz(SAKI_YOUK_TRAN.strMOTO_ACPTNO, "") <> "" And SAKI_YOUK_TRAN.strMOVEKBN <> P_MOVEKBN_00) Then
                    If fncSetMoveMotoData(SAKI_YOUK_TRAN.strMOTO_ACPTNO, SAKI_YOUK_TRAN.strMOVEKBN) = False Then GoTo Exit_rtn
                End If

                ' --- 顧客データ設定 -------------------------------------------------
                ' ※既存：顧客マスタ
                '   新規：予約受付トラン
                If Nz(objRsYK![YOUKT_UCODE], 0) = 0 Then
                    ' 顧客コードがNULLの場合、新規顧客
'                    Me.opg_顧客情報取得選択.Value = 1  DEL 2006/03/03 Tajima

                    ' 予約受付トラン.顧客コードがNULLの場合、予約受付トランの顧客情報を設定し、
                    ' 画面の各顧客情報入力項目を使用可能に設定する
                    .cmb_顧客コード = Format(Nz(objRsYK![YOUKT_UCODE], 0), "000000") ' 顧客コード

                    pvarKcode = objRsYK![YOUKT_UCODE]          ' (フォームモジュール内で使用)
                    If Nz(objRsYK![YOUKT_UCODE], 0) <> 0 Then
                        .txt_顧客名 = objRsYK![YOUKT_NAME]     ' 顧客名
                    Else
                        .txt_顧客名 = objRsYK![YOUKT_NAME]     ' 顧客名
                    End If

                    .txt_顧客名カナ = objRsYK![YOUKT_KANA]     ' 顧客名ｶﾅ
                    .txt_顧客カナ = objRsYK![YOUKT_KANA]       ' 顧客名ｶﾅ
                    .txt_代表者名 = objRsYK![YOUKT_TANM]       ' 代表者名
                    .txt_代表者名カナ = objRsYK![YOUKT_TAKA]   ' 代表者名ｶﾅ
                    .txt_郵便番号 = objRsYK![YOUKT_YUBINO]     ' 郵便番号
                    .txt_住所1 = objRsYK![YOUKT_ADR_1]         ' 住所1
                    .txt_住所2 = objRsYK![YOUKT_ADR_2]         ' 住所2
                    .txt_住所3 = objRsYK![YOUKT_ADR_3]         ' 住所3
                    .txt_メールアドレス = objRsYK![YOUKT_MAIL] ' メールアドレス  ' 2006/01/31 ADD T.SUZUKI
                    .txt_電話番号 = objRsYK![YOUKT_TEL]        ' 電話番号
                    .txt_携帯番号 = objRsYK![YOUKT_KEITAI]     ' 携帯番号
                    .txt_FAX番号 = objRsYK![YOUKT_FAX]         ' FAX番号

                    .cmb_顧客区分 = objRsYK![YOUKT_KKBN]       ' 顧客区分

                    ' 顧客区分名取得
                    vRet = fncSelectRyak("020", Me.cmb_顧客区分)                 ' 2006/01/05 ADD T.SUZUKI
'                    vRet = fncSelectRyak(Me.cmb_顧客区分, "020")                ' 2006/01/05 DEL T.SUZUKI
                    .txt_顧客区分名 = vRet

                Else
                    ' 既存の顧客データを設定
'                    Me.opg_顧客情報取得選択.Value = 2  DEL 2006/03/03 Tajima
                    .cmb_顧客コード = Format(Nz(objRsYK![YOUKT_UCODE], 0), "000000") ' 顧客コード

                    ' 顧客データ取得
                    Set objRs = pobjDB.OpenRecordset(SEL_USER_QUERY(objRsYK![YOUKT_UCODE]), dbOpenSnapshot)

                    If Not (objRs.BOF And objRs.EOF) Then
                        ' 顧客マスタより取得した値を画面に設定
                        ' 【顧客情報】
                        .txt_顧客名 = objRs![USER_NAME]
                        .txt_電話番号 = objRs![USER_TEL]
                        .txt_顧客名カナ = objRs![USER_KANA]
                        .txt_顧客カナ = objRs![USER_KANA]
                        .txt_代表者名 = objRs![USER_TANM]
                        .txt_代表者名カナ = objRs![USER_TAKA]
                        .txt_郵便番号 = objRs![USER_YUBINO]
                        .txt_住所1 = objRs![USER_ADR_1]
                        .txt_住所2 = objRs![USER_ADR_2]
                        .txt_住所3 = objRs![USER_ADR_3]
                        .txt_メールアドレス = objRs![USER_MAIL]
                        .txt_携帯番号 = objRs![USER_KEITAI]
                        .txt_FAX番号 = objRs![USER_FAX]
                        .txt_締日 = objRs![USER_SDAY]
                        .cmb_顧客区分 = objRs![USER_KKBN]
                        .cmb_集金区分 = objRs![USER_SKBN]
                        .txt_顧客備考 = objRs![USER_BIKO]

                        ' 顧客区分名取得
                        vRet = fncSelectRyak("020", Me.cmb_顧客区分)   ' 2006/01/05 ADD T.SUZUKI
'                        vRet = fncSelectRyak(Me.cmb_顧客区分, "020")  ' 2006/01/05 DEL T.SUZUKI
                        .txt_顧客区分名 = vRet

                        ' 集金区分名取得
                        vRet = fncSelectRyak("030", Me.cmb_集金区分)   ' 2006/01/05 ADD T.SUZUKI
'                        vRet = fncSelectRyak(Me.cmb_集金区分, "030")  ' 2006/01/05 DEL T.SUZUKI
                        .txt_集金区分名 = vRet

                        ' 顧客マスタ.連絡先担当者名 IS NOT NULL のとき、入力するチェックボックスを選択状態に設定
                        If Nz(Trim(objRs![USER_RNAME]), "") <> "" Then
                            .chk_入力する = -1
                        End If

                        ' 【その他連絡先】を設定
                        .txt_連絡先担当名カナ = objRs![USER_RKANA]
                        .txt_連絡先担当者名 = objRs![USER_RNAME]
                        .txt_連絡先郵便番号 = objRs![USER_RPOST]
                        .txt_連絡先住所1 = objRs![USER_RADR_1]
                        .txt_連絡先住所2 = objRs![USER_RADR_2]
                        .txt_連絡先住所3 = objRs![USER_RADR_3]
                        .txt_連絡先電話番号 = objRs![USER_RTEL]
                        .txt_連絡先携帯番号 = objRs![USER_RKEITAI]
                        .txt_連絡先FAX = objRs![USER_RFAX]
                    End If
                End If
                ' --- 顧客データ設定 -------------------------------------------------
            End With

        Case 2  ' *** 契約修正 ***
            With Me
                .cmb_ヤードコード = pstrYCode      ' ヤードコード
                Call cmb_ヤードコード_AfterUpdate

                .cmb_コンテナ番号 = pstrNo         ' コンテナ番号
                pvarContna = ""                    ' (フォームモジュール内で使用)

                ' ↓ 以降はイベントプロシージャー『cmb_コンテナ番号_AfterUpdate』で処理を行う。
                ' ※コンテナ契約ファイルより取得した値を画面に設定
                Call cmb_コンテナ番号_AfterUpdate
            End With
    End Select

    ' 戻り値設定
    fncSetYoyakuData = True

Exit_rtn:
    doCmd.Hourglass False
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function
' 2006/01/05 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :fncSetYoyakuData
''        CREATE_DATE      :2005/07/13
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncSetYoyakuData(intmode As Integer) As Boolean
'
'    Dim vRet    As Variant
'
'    On Error GoTo err_rtn
'
'    fncSetYoyakuData = False
'
'    Select Case intmode
'        Case 1  ' *** 契約新規取込あり ***
'            'コンテナ状況区分
'            pvarContna = -1
'            pvarKcode = ""
'
'            With Me
'                ' 予約受付(選択状態に設定)
'                .chk_予約受付.Value = -1
'
'                ' 受付番号
'                .cmb_受付番号 = pstrUKNO
'
'                ' ヤードコード
'                .cmb_ヤードコード = pobjRS2![YOUKT_YCODE]
'' hana cmb_ヤードコード_AfterUpdateにてクリア処理を行うように修正されている！！！
'                Call cmb_ヤードコード_AfterUpdate
'
'                ' 新規登録の場合、【cmb_コンテナ番号_AfterUpdate】処理にて『顧客コード』が必要なため、事前にセット
'                ' 顧客コード
'                .cmb_顧客コード = Format(Nz(pobjRS2![YOUKT_UCODE], 0), "000000")  ' 2005/09/12 ADD T.SUZUKI
''                .cmb_顧客コード = Nz(pobjRS2![YOUKT_UCODE], 0)                   ' 2005/09/12 DEL T.SUZUKI
'
'                ' コンテナ番号
'                .cmb_コンテナ番号 = pobjRS2![YOUKT_NO]
'                pvarContna = pobjRS2![YOUKT_NO]                       ' (フォームモジュール内で使用)
'                Call cmb_コンテナ番号_AfterUpdate
'
'                ' 受付区分
'                .cmd_広告媒体 = Format(pobjRS2![YOUKT_UKKBN], "0")
'                Call cmd_広告媒体_AfterUpdate
'
'                .txt_解約日付 = ""                      ' 解約日付
''                .chk_コンテナ状態 = 0                   ' コンテナ状態  ' 2005/08/03 DEL T.SUZUKI
'
'                ' --- 顧客データ設定 -------------------------------------------------
'                ' ※既存：顧客マスタ
'                '   新規：予約受付トラン
'                If Nz(pobjRS2![YOUKT_UCODE], 0) = 0 Then
'                    ' 顧客コードがNULLの場合、新規顧客
'                    Me.opg_顧客情報取得選択.Value = 1
'
'                    ' 予約受付トラン.顧客コードがNULLの場合、予約受付トランの顧客情報を設定し、
'                    ' 画面の各顧客情報入力項目を使用可能に設定する
'' 2005/10/19 REP T.SUZUKI Start
'                    .cmb_顧客コード = Format(Nz(pobjRS2![YOUKT_UCODE], 0), "000000") ' 顧客コード
''                    .cmb_顧客コード = Format(pobjRS2![YOUKT_UCODE], "000000") ' 顧客コード  ' 2005/09/12 ADD T.SUZUKI
''                    .cmb_顧客コード = pobjRS2![YOUKT_UCODE]  ' 顧客コード                  ' 2005/09/12 DEL T.SUZUKI
'' 2005/10/19 REP T.SUZUKI End
'
'                    pvarKcode = pobjRS2![YOUKT_UCODE]          ' (フォームモジュール内で使用)
'                    If Nz(pobjRS2![YOUKT_UCODE], 0) <> 0 Then
'                        .txt_顧客名 = pobjRS2![YOUKT_NAME]   ' 顧客名
''                        .cmb_顧客カナ.Requery  ' 2005/10/19 DEL T.SUZUKI
'                    Else
'                        .txt_顧客名 = pobjRS2![YOUKT_NAME]   ' 顧客名
'                    End If
'
'                    .txt_顧客名カナ = pobjRS2![YOUKT_KANA]   ' 顧客名ｶﾅ
'                    .txt_顧客カナ = pobjRS2![YOUKT_KANA]     ' 顧客名ｶﾅ   ' 2005/10/19 ADD T.SUZUKI
''                    .cmb_顧客カナ = pobjRS2![YOUKT_KANA]     ' 顧客名ｶﾅ  ' 2005/10/19 DEL T.SUZUKI
'                    .txt_代表者名 = pobjRS2![YOUKT_TANM]     ' 代表者名
'                    .txt_代表者名カナ = pobjRS2![YOUKT_TAKA] ' 代表者名ｶﾅ
'                    .txt_郵便番号 = pobjRS2![YOUKT_YUBINO]   ' 郵便番号
'                    .txt_住所1 = pobjRS2![YOUKT_ADR_1]       ' 住所1
'                    .txt_住所2 = pobjRS2![YOUKT_ADR_2]       ' 住所2
'                    .txt_住所3 = pobjRS2![YOUKT_ADR_3]       ' 住所3
'                    .txt_電話番号 = pobjRS2![YOUKT_TEL]      ' 電話番号
'                    .txt_携帯番号 = pobjRS2![YOUKT_KEITAI]   ' 携帯番号
'                    .txt_FAX番号 = pobjRS2![YOUKT_FAX]       ' FAX番号
'
'                    .cmb_顧客区分 = pobjRS2![YOUKT_KKBN]     ' 顧客区分
'
'' 2005/08/03 REP T.SUZUKI Start
'                    ' 顧客区分名取得
'                    vRet = fncSelectRyak(Me.cmb_顧客区分, "020")   ' 2005/08/09 ADD T.SUZUKI
''                    vRet = fncSelectName(Me.cmb_顧客区分, "020")  ' 2005/08/09 DEL T.SUZUKI
''                    strSQL = " NAME_ID ='020'"
''                    strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_顧客区分)
''                    vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'                    .txt_顧客区分名 = vRet
'' 2005/08/03 REP T.SUZUKI End
'
'                Else
'                    ' 既存の顧客データを設定
'                    Me.opg_顧客情報取得選択.Value = 2
'' 2005/10/19 REP T.SUZUKI Start
'                    .cmb_顧客コード = Format(Nz(pobjRS2![YOUKT_UCODE], 0), "000000") ' 顧客コード
'
'                    ' 顧客データ取得
'                    Set pobjRs = pobjDB.OpenRecordset(SEL_USER_QUERY(pobjRS2![YOUKT_UCODE]), dbOpenSnapshot)
''                    Set pobjRs = pobjDB.OpenRecordset(SEL_USER_QUERY(Me.cmb_顧客コード), dbOpenSnapshot)
'' 2005/10/19 REP T.SUZUKI End
'
'                    If Not (pobjRs.BOF And pobjRs.EOF) Then
'                        ' 顧客マスタより取得した値を画面に設定
'                        ' 【顧客情報】
''                        .cmb_顧客カナ.Requery  ' 2005/10/19 DEL T.SUZUKI
'                        .txt_顧客名 = pobjRs![USER_NAME]
'                        .txt_電話番号 = pobjRs![USER_TEL]
'                        .txt_顧客名カナ = pobjRs![USER_KANA]
'                        .txt_顧客カナ = pobjRs![USER_KANA]   ' 2005/10/19 ADD T.SUZUKI
''                        .cmb_顧客カナ = pobjRs![USER_KANA]  ' 2005/10/19 DEL T.SUZUKI
'                        .txt_代表者名 = pobjRs![USER_TANM]
'                        .txt_代表者名カナ = pobjRs![USER_TAKA]
'                        .txt_郵便番号 = pobjRs![USER_YUBINO]
'                        .txt_住所1 = pobjRs![USER_ADR_1]
'                        .txt_住所2 = pobjRs![USER_ADR_2]
'                        .txt_住所3 = pobjRs![USER_ADR_3]
'                        .txt_メールアドレス = pobjRs![USER_MAIL]
'                        .txt_携帯番号 = pobjRs![USER_KEITAI]
'                        .txt_FAX番号 = pobjRs![USER_FAX]
'                        .txt_締日 = pobjRs![USER_SDAY]
'                        .cmb_顧客区分 = pobjRs![USER_KKBN]
'                        .cmb_集金区分 = pobjRs![USER_SKBN]
'                        .txt_顧客備考 = pobjRs![USER_BIKO]
'
'' 2005/08/03 REP T.SUZUKI Start
'                        ' 顧客区分名取得
'                        vRet = fncSelectRyak(Me.cmb_顧客区分, "020")   ' 2005/08/09 ADD T.SUZUKI
''                        vRet = fncSelectName(Me.cmb_顧客区分, "020")  ' 2005/08/09 DEL T.SUZUKI
''                        strSQL = " NAME_ID ='020'"
''                        strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_顧客区分)
''                        vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'                        .txt_顧客区分名 = vRet
'' 2005/08/03 REP T.SUZUKI End
'
'' 2005/08/03 REP T.SUZUKI Start
'                        ' 集金区分名取得
'                        vRet = fncSelectRyak(Me.cmb_集金区分, "030")   ' 2005/08/09 ADD T.SUZUKI
''                        vRet = fncSelectName(Me.cmb_集金区分, "030")  ' 2005/08/09 DEL T.SUZUKI
''                        strSQL = " NAME_ID ='030'"
''                        strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_集金区分)
''                        vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'                        .txt_集金区分名 = vRet
'' 2005/08/03 REP T.SUZUKI End
'
'                        ' 顧客マスタ.連絡先担当者名 IS NOT NULL のとき、入力するチェックボックスを選択状態に設定
'                        If Nz(Trim(pobjRs![USER_RNAME]), "") <> "" Then
'                            .chk_入力する = -1
'                        End If
'
'                        ' 【その他連絡先】を設定
'                        .txt_連絡先担当名カナ = pobjRs![USER_RKANA]
'                        .txt_連絡先担当者名 = pobjRs![USER_RNAME]
'                        .txt_連絡先郵便番号 = pobjRs![USER_RPOST]
'                        .txt_連絡先住所1 = pobjRs![USER_RADR_1]
'                        .txt_連絡先住所2 = pobjRs![USER_RADR_2]
'                        .txt_連絡先住所3 = pobjRs![USER_RADR_3]
'                        .txt_連絡先電話番号 = pobjRs![USER_RTEL]
'                        .txt_連絡先携帯番号 = pobjRs![USER_RKEITAI]
'                        .txt_連絡先FAX = pobjRs![USER_RFAX]
'                    End If
'                End If
'                ' --- 顧客データ設定 -------------------------------------------------
'            End With
'
'        Case 2  ' *** 契約修正 ***
'            With Me
'                .cmb_ヤードコード = pstrYCODE      ' ヤードコード
'                Call cmb_ヤードコード_AfterUpdate
'
'                .cmb_コンテナ番号 = pstrNO         ' コンテナ番号
'                pvarContna = ""                   ' (フォームモジュール内で使用)
'
'                ' ↓ 以降はイベントプロシージャー『cmb_コンテナ番号_AfterUpdate』で処理を行う。
'                ' ※コンテナ契約ファイルより取得した値を画面に設定
'                Call cmb_コンテナ番号_AfterUpdate
'            End With
'    End Select
'
'    ' 戻り値設定
'    fncSetYoyakuData = True
'
'Exit_rtn:
'    doCmd.Hourglass False
'    Set pobjRS2 = Nothing
'    Exit Function
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'    GoTo Exit_rtn
'End Function
' 2006/01/05 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :移動元データ設定処理
'        MODULE_ID        :fncSetMoveMotoData
'        Parameter        :
'                          strMoto_AcptNo = 予約受付トラン.元受注契約番号
'                          strMoveKbn     = 予約受付トラン.移動元区分
'        CREATE_DATE      :2005/07/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSetMoveMotoData(ByVal strMOTO_ACPTNO As String, ByVal strMOVEKBN As String) As Boolean

    Dim objRsCarg As Recordset
    Dim strBIKO   As String

    On Error GoTo fncSetMoveMotoData_Err

    fncSetMoveMotoData = False

    ' 移動元コンテナ契約ファイルデータ取得
    If fncGetCARG_FILE(strMOTO_ACPTNO, objRsCarg) = False Then GoTo fncSetMoveMotoData_Exit

    With objRsCarg
        MOTO_CARG_FILE.strYcode = Format(.Fields("CARG_YCODE"), "000000")    ' ヤードコード
        MOTO_CARG_FILE.strNO = Format(.Fields("CARG_NO"), "000000")          ' コンテナ番号
        MOTO_CARG_FILE.strUCODE = Format(.Fields("CARG_UCODE"), "000000")    ' 顧客コード
        MOTO_CARG_FILE.strAGRE = Format(.Fields("CARG_AGRE"), "00")          ' 契約状態
        MOTO_CARG_FILE.dblRENTKG = .Fields("CARG_RENTKG")                    ' 月額料金
        MOTO_CARG_FILE.strACPTNO = .Fields("CARG_ACPTNO")                    ' 受注契約番号
        MOTO_CARG_FILE.strHOSYICD = .Fields("CARG_HOSYICD")                  ' 保証区分
        MOTO_CARG_FILE.strUKTANTO = .Fields("CARG_UKTANTO")                  ' 契約受付担当者コード

        ' 初期表示用「備考」文言作成
        strBIKO = ""
        Call subCreateMemo(0, strMOVEKBN, strBIKO)

        Select Case strMOVEKBN
            Case P_MOVEKBN_01
                ' *** 自社理由 ***
                Me.txt_解約日付 = Format(.Fields("CARG_KYDATE"), P_FORMAT_YMD)
                Me.txt_月額料金 = .Fields("CARG_RENTKG")
                Call txt_月額料金_AfterUpdate
                'Me.txt_初回金額 = .Fields("CARG_RENTKG") 'DEL 2006/03/03 Tajima
                'Call txt_初回金額_AfterUpdate            'DEL 2006/03/03 Tajima
                Me.txt_契約備考 = Format(.Fields("CARG_YCODE"), "000000") & "-" & _
                                  Format(.Fields("CARG_NO"), "000000") & _
                                strMSGMoveBiko
                Me.cmb_保証区分 = .Fields("CARG_HOSYICD"): Call cmb_保証区分_AfterUpdate
                'Me.txt_預り保証金 = .Fields("CARG_SECUKG") 2006/03/03 DEL Tajima
                Me.cmb_鍵区分 = .Fields("CARG_KAGIICD"): Call cmb_鍵区分_AfterUpdate
                Me.txt_解約受付日 = Format(.Fields("CARG_KAIUKEDATE"), P_FORMAT_YMD)
                Me.cmb_解約受付担当者コード = .Fields("CARG_KAITANTO"): Call cmb_解約受付担当者コード_AfterUpdate
                Me.cmb_解約区分 = .Fields("CARG_KAICD"): Call cmb_解約区分_AfterUpdate
                Me.cmb_解約理由 = .Fields("CARG_KAIRIYUCD"): Call cmb_解約理由_AfterUpdate

            Case P_MOVEKBN_02
                ' *** 顧客理由 ***
                Me.txt_解約日付 = Format(.Fields("CARG_KYDATE"), P_FORMAT_YMD)
                Me.txt_契約備考 = Format(.Fields("CARG_YCODE"), "000000") & "-" & _
                                  Format(.Fields("CARG_NO"), "000000") & _
                                strMSGMoveBiko
                Me.cmb_保証区分 = .Fields("CARG_HOSYICD"): Call cmb_保証区分_AfterUpdate
                'Me.txt_預り保証金 = .Fields("CARG_SECUKG") 2006/03/03 DEL Tajima
                Me.cmb_鍵区分 = .Fields("CARG_KAGIICD"): Call cmb_鍵区分_AfterUpdate
                Me.txt_解約受付日 = Format(.Fields("CARG_KAIUKEDATE"), P_FORMAT_YMD)
                Me.cmb_解約受付担当者コード = .Fields("CARG_KAITANTO"): Call cmb_解約受付担当者コード_AfterUpdate
                Me.cmb_解約区分 = .Fields("CARG_KAICD"): Call cmb_解約区分_AfterUpdate
                Me.cmb_解約理由 = .Fields("CARG_KAIRIYUCD"): Call cmb_解約理由_AfterUpdate

            Case P_MOVEKBN_03
                ' *** 未使用移動 ***
                Me.txt_解約日付 = Format(.Fields("CARG_KYDATE"), P_FORMAT_YMD)
                Me.txt_月額料金 = .Fields("CARG_RENTKG")
                Call txt_月額料金_AfterUpdate
                'Me.txt_初回金額 = .Fields("CARG_RENTKG") DEL 2006/03/03 Tajima
                'Call txt_初回金額_AfterUpdate            DEL 2006/03/03 Tajima
                Me.txt_契約備考 = Format(.Fields("CARG_YCODE"), "000000") & "-" & _
                                  Format(.Fields("CARG_NO"), "000000") & _
                                strMSGMoveBiko
                Me.cmb_保証区分 = .Fields("CARG_HOSYICD"): Call cmb_保証区分_AfterUpdate
                'Me.txt_預り保証金 = .Fields("CARG_SECUKG") 2006/03/03 DEL Tajima
                Me.cmb_鍵区分 = .Fields("CARG_KAGIICD"): Call cmb_鍵区分_AfterUpdate
                Me.txt_解約受付日 = Format(.Fields("CARG_KAIUKEDATE"), P_FORMAT_YMD)
                Me.cmb_解約受付担当者コード = .Fields("CARG_KAITANTO"): Call cmb_解約受付担当者コード_AfterUpdate
                Me.cmb_解約区分 = .Fields("CARG_KAICD"): Call cmb_解約区分_AfterUpdate
                Me.cmb_解約理由 = .Fields("CARG_KAIRIYUCD"): Call cmb_解約理由_AfterUpdate

            Case P_MOVEKBN_04
                ' *** 名義変更 ***
                Me.txt_解約日付 = Format(.Fields("CARG_KYDATE"), P_FORMAT_YMD)
                Me.txt_月額料金 = .Fields("CARG_RENTKG")
                Call txt_月額料金_AfterUpdate
                'Me.txt_初回金額 = .Fields("CARG_RENTKG") DEL 2006/03/03 Tajima
                'Call txt_初回金額_AfterUpdate            DEL 2006/03/03 Tajima
                Me.txt_契約備考 = Format(.Fields("CARG_YCODE"), "000000") & "-" & _
                                  Format(.Fields("CARG_NO"), "000000") & _
                                strMSGMoveBiko
                Me.cmb_保証区分 = .Fields("CARG_HOSYICD"): Call cmb_保証区分_AfterUpdate
                'Me.txt_預り保証金 = .Fields("CARG_SECUKG") 2006/03/03 DEL Tajima
                Me.cmb_鍵区分 = .Fields("CARG_KAGIICD"): Call cmb_鍵区分_AfterUpdate
                Me.txt_解約受付日 = Format(.Fields("CARG_KAIUKEDATE"), P_FORMAT_YMD)
                Me.cmb_解約受付担当者コード = .Fields("CARG_KAITANTO"): Call cmb_解約受付担当者コード_AfterUpdate
                Me.cmb_解約区分 = .Fields("CARG_KAICD"): Call cmb_解約区分_AfterUpdate
                Me.cmb_解約理由 = .Fields("CARG_KAIRIYUCD"): Call cmb_解約理由_AfterUpdate

'            Case P_MOVEKBN_05                  'DELETE 2011/09/27 M.RYU
            Case P_MOVEKBN_05, P_MOVEKBN_07     'INSERT 2011/09/27 M.RYU
                ' *** 金額変更 ***
                Me.txt_解約日付 = Format(.Fields("CARG_KYDATE"), P_FORMAT_YMD)
                Me.txt_契約備考 = Format(.Fields("CARG_YCODE"), "000000") & "-" & _
                                  Format(.Fields("CARG_NO"), "000000") & _
                                strMSGMoveBiko
                Me.cmb_保証区分 = .Fields("CARG_HOSYICD"): Call cmb_保証区分_AfterUpdate
                'Me.txt_預り保証金 = .Fields("CARG_SECUKG") 2006/03/03 DEL Tajima
                Me.cmb_鍵区分 = .Fields("CARG_KAGIICD"): Call cmb_鍵区分_AfterUpdate
                Me.txt_解約受付日 = Format(.Fields("CARG_KAIUKEDATE"), P_FORMAT_YMD)
                Me.cmb_解約受付担当者コード = .Fields("CARG_KAITANTO"): Call cmb_解約受付担当者コード_AfterUpdate
                Me.cmb_解約区分 = .Fields("CARG_KAICD"): Call cmb_解約区分_AfterUpdate
                Me.cmb_解約理由 = .Fields("CARG_KAIRIYUCD"): Call cmb_解約理由_AfterUpdate

            Case P_MOVEKBN_06
                ' *** 保障会社変更 ***
                Me.txt_解約日付 = Format(.Fields("CARG_KYDATE"), P_FORMAT_YMD)
                Me.txt_月額料金 = .Fields("CARG_RENTKG")
                Call subCalculation_TAX                                          ' 月額使用料の消費税額算出
                Call subCalculation_TOTAL                                        ' 月額使用量の合計金額算出
                Me.txt_契約備考 = Format(.Fields("CARG_YCODE"), "000000") & "-" & _
                                  Format(.Fields("CARG_NO"), "000000") & _
                                strMSGMoveBiko
                Me.cmb_鍵区分 = .Fields("CARG_KAGIICD"): Call cmb_鍵区分_AfterUpdate
                Me.txt_解約受付日 = Format(.Fields("CARG_KAIUKEDATE"), P_FORMAT_YMD)
                Me.cmb_解約受付担当者コード = .Fields("CARG_KAITANTO"): Call cmb_解約受付担当者コード_AfterUpdate
                Me.cmb_解約区分 = .Fields("CARG_KAICD"): Call cmb_解約区分_AfterUpdate
                Me.cmb_解約理由 = .Fields("CARG_KAIRIYUCD"): Call cmb_解約理由_AfterUpdate
        End Select
    End With
    objRsCarg.Close
    Set objRsCarg = Nothing

    ' 初期表示用備考設定
    Me.txt_契約備考 = strBIKO

    fncSetMoveMotoData = True

fncSetMoveMotoData_Exit:
    If Not objRsCarg Is Nothing Then objRsCarg.Close: Set objRsCarg = Nothing
    doCmd.Hourglass False
    Exit Function

fncSetMoveMotoData_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncSetMoveMotoData_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :GetTantoData
'        機能             :担当者マスタ情報取得
'        IN               :vTanto = 担当者ｺｰﾄﾞ / vTanNm = 担当者名
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetTantoData(ByRef vTanto As Variant, _
                              Optional ByRef vTanNm As Variant) As Boolean

    Dim strSQL  As String
    Dim objRs   As Recordset

    GetTantoData = False

    strSQL = "SELECT "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC, "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTN "
    strSQL = strSQL & "FROM TANT_MAST INNER JOIN "
    strSQL = strSQL & "dbo_CONT_MAST ON "
    strSQL = strSQL & "TANT_MAST.TANTM_BUMOC = dbo_CONT_MAST.CONT_BUMOC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"

    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)
    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Exit Function
    Else
        If Nz(objRs("TANTM_TANTN"), "") <> "" Then vTanNm = objRs("TANTM_TANTN")
    End If

GetTantoData_End:
    objRs.Close
    Set objRs = Nothing
    GetTantoData = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :部門コード/部門名称の取得
'        MODULE_ID        :subGetBumonName
'        IN               :objDB = DB接続 /strBumonCd = 部門コード /strBumonNm = 部門名
'        CREATE_DATE      :2005/08/31
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetBumonName(ByRef objDb As Database, ByRef strBumonCd As String, _
                            Optional ByRef strBumonNm As String = "")
'Private Sub subGetBumonName(ByRef objDb As dataBase, ByRef strBumonCd As String, ByRef strBumonNm As String)

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo subGetBumonName_Err

    strSQL = "SELECT "
    strSQL = strSQL & "BUMOM_BUMOC, "  ' 部門コード
    strSQL = strSQL & "BUMOM_BUMON "   ' 部門名称
    strSQL = strSQL & "FROM dbo_CONT_MAST INNER JOIN BUMO_MAST ON "
    strSQL = strSQL & "dbo_CONT_MAST.CONT_BUMOC = BUMO_MAST.BUMOM_BUMOC "
    Set objRs = objDb.OpenRecordset(strSQL, dbReadOnly)

    With objRs
        If Not .EOF Then
            strBumonCd = .Fields("BUMOM_BUMOC")
            strBumonNm = .Fields("BUMOM_BUMON")
        Else
            strBumonCd = Null
            strBumonNm = Null
        End If
    End With

subGetBumonName_Exit:
    If Not objRs Is Nothing Then
        objRs.Close
        Set objRs = Nothing
    End If
    Exit Sub

subGetBumonName_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
    GoTo subGetBumonName_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncSelectRyak
'        CREATE_DATE      :
'
'保証区分、鍵区分の値を取得する。
'パラメータ ContNM Object
'戻り値 strRyak String [名称マスタ].略称
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2006/01/05 REP T.SUZUKI Start
Private Function fncSelectRyak(stid As String, Optional ByVal ContNM As Object, _
                               Optional ByRef strCode As String) As String
'Private Function fncSelectRyak(ByVal ContNM As Object, stid As String) As String
' 2006/01/05 REP T.SUZUKI End

    Dim objControl  As Object
    Dim vCode       As Variant
    Dim strID       As String
    Dim strWhere    As String
    Dim strRyak     As String

    fncSelectRyak = ""
    Set objControl = ContNM

' 2006/01/05 ADD T.SUZUKI Start
    If Nz(strCode, "") = "" Then
        vCode = Nz(objControl, "")
    Else
        vCode = Nz(strCode, "")
    End If
' 2006/01/05 ADD T.SUZUKI End
'    vCode = Nz(ObjControl, "")          ' 2006/01/05 DEL T.SUZUKI

    If vCode = "" Then Exit Function

    strID = stid

    ' 名称マスタより略称を取得
    strWhere = ""
    strWhere = strWhere & " NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & (vCode)
    strRyak = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strWhere), "")

    fncSelectRyak = strRyak
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncSelectName
'        CREATE_DATE      :
'
'パラメータ ContNM Object
'戻り値 strRyak String [名称マスタ].名称
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2006/01/05 REP T.SUZUKI Start
Private Function fncSelectName(stid As String, Optional ByVal ContNM As Object, _
                               Optional ByRef strCode As String) As String
'Private Function fncSelectName(ByVal ContNM As Object, stid As String) As String
' 2006/01/05 REP T.SUZUKI End

    Dim objControl  As Object
    Dim vCode       As Variant
    Dim strID       As String
    Dim strWhere    As String
    Dim strName     As String

    fncSelectName = ""
    Set objControl = ContNM

' 2006/01/05 ADD T.SUZUKI Start
    If Nz(strCode, "") = "" Then
        vCode = Nz(objControl, "")
    Else
        vCode = Nz(strCode, "")
    End If
' 2006/01/05 ADD T.SUZUKI End
'    vCode = Nz(ObjControl, "")          ' 2006/01/05 DEL T.SUZUKI

    If vCode = "" Then Exit Function

    strID = stid

    ' 名称マスタより名称を取得
    strWhere = ""
    strWhere = strWhere & " NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE = " & (vCode)
    strName = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")

    fncSelectName = strName
End Function

'==============================================================================*
'
'        MODULE_NAME      :ヤードマスタ問い合わせ
'        MODULE_ID        :CHK_YARD_MAST
'        CREATE_DATE      :
'
  '引数 ０：存在チェック
  '     １：データ参照
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_Yard_Mast(ByVal iMode As Integer) As Boolean
    
    Dim strSQL As String
    Dim objRs  As Recordset                                     ' 2006/01/05 ADD T.SUZUKI
    Dim vCode  As Variant
    Dim vRet   As Variant
    Dim i      As Integer
    Dim vRET2  As Variant

    On Error GoTo err_rtn
  
    Chk_Yard_Mast = False
  
    vCode = Nz(Me.cmb_ヤードコード, "")
  
    If vCode = "" Then
        Chk_Yard_Mast = True
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    If iMode = 0 Then
        '存在チェック
        strSQL = " YARD_CODE =" & Val(vCode)

        vRet = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strSQL), "")
      
        If vRet = "" Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, strMsg_kaku
            Chk_Yard_Mast = True
        End If

    Else
        strSQL = "  SELECT"
        strSQL = strSQL & " YARD_NAME,"
        strSQL = strSQL & " YARD_ADDR_1,"
        strSQL = strSQL & " YARD_ADDR_2,"
        strSQL = strSQL & " YARD_ADDR_3 "
    
        '2006/09/22 SHIBAZAKI ヤード営業終了日、ヤード解約予定日追加
        strSQL = strSQL & ","
        strSQL = strSQL & "YARD_RENTEND_DAY, "
        strSQL = strSQL & "YARD_END_DAY "
'↓ INSERT 2014/06/24 MIYAMOTO
        strSQL = strSQL & ","
        strSQL = strSQL & "YARD_TIHO_KBN "
'↑ INSERT 2014/06/24 MIYAMOTO
        
        strSQL = strSQL & " FROM"
        strSQL = strSQL & " dbo_YARD_MAST"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " YARD_CODE =" & Val(vCode)
        Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)    ' 2006/01/05 ADD T.SUZUKI
'        Set pobjRs = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)  ' 2006/01/05 DEL T.SUZUKI

        If Not (objRs.BOF And objRs.EOF) Then
            Me.txt_ヤード名 = objRs.Fields("YARD_NAME")
            Me.txt_ヤード住所 = objRs.Fields("YARD_ADDR_1") & _
                                objRs.Fields("YARD_ADDR_2") & _
                                objRs.Fields("YARD_ADDR_3")
            
            '2006/09/22 SHIBAZAKI ヤード営業期間とヤード解約予定日追加
            Me.txt_営業終了日 = Nz(objRs.Fields("YARD_RENTEND_DAY"), "")
            Me.txt_ヤード解約予定 = Nz(objRs.Fields("YARD_END_DAY"), "")
            pvarKYDATE = Nz(objRs.Fields("YARD_END_DAY"), "9999/12/31")
            Me.txt_地方ヤード = Nz(objRs.Fields("YARD_TIHO_KBN"), "0")      'INSERT 2014/06/24 MIYAMOTO
        End If
'        If Not (pobjRs.BOF And pobjRs.EOF) Then
'            Me.txt_ヤード名 = pobjRs.Fields("YARD_NAME")
'            Me.txt_ヤード住所 = pobjRs.Fields("YARD_ADDR_1") & _
'                                pobjRs.Fields("YARD_ADDR_2") & _
'                                pobjRs.Fields("YARD_ADDR_3")
'        End If
    End If

    Chk_Yard_Mast = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing   ' 2006/01/05 ADD T.SUZUKI
'    Set pobjRs = Nothing                                           ' 2006/01/05 DEL T.SUZUKI
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'2006/09/22 SHIBAZAKI 地主情報削除
''==============================================================================*
''
''        MODULE_NAME      :地主契約問い合わせ
''        MODULE_ID        :CHK_JARG_FILE
''        CREATE_DATE      :
''
'  '引数 ０：存在チェック
'  '     １：データ参照
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function CHK_JARG_FILE(ByVal iMode As Integer) As Boolean
'
'    Dim strSql  As String
'    Dim objRs   As Recordset                                        ' 2006/01/05 ADD T.SUZUKI
'    Dim vCode   As Variant
'    Dim vRet    As Variant
'
'    On Error GoTo err_rtn
'
'    CHK_JARG_FILE = False
'
'    vCode = Nz(Me.cmb_ヤードコード, "")
'
'    If vCode = "" Then
'        CHK_JARG_FILE = True
'        GoTo Exit_rtn
'    End If
'
'    doCmd.Hourglass True
'    doCmd.SetWarnings False
'
'    If iMode = 0 Then
'        ' *** 地主契約マスタ存在チェック ***
'        pintJARG_CHK = 0
'
'        '存在チェック①
'        strSql = " JARG_YCODE =" & Val(vCode)
'        strSql = strSql & " AND JARG_STDATE <= CDATE(#" & DATE & "#)"
'        strSql = strSql & " AND JARG_EDDATE >= CDATE(#" & DATE & "#)"
'        strSql = strSql & " AND ( JARG_KYDATE >= CDATE(#" & DATE & "#) OR JARG_KYDATE IS NULL )"
'        vRet = Nz(DLookup("JARG_JCODE", "dbo_JARG_FILE", strSql), "")
'
'        pintJARG_CHK = 1
'
'        If vRet = "" Then
'            pintJARG_CHK = 2
'
'            '存在チェック②
'            strSql = " JARG_YCODE =" & Val(vCode)
'            strSql = strSql & " AND JARG_STDATE <= CDATE(#" & DATE & "#)"
'            vRet = Nz(DLookup("JARG_JCODE", "dbo_JARG_FILE", strSql), "")
'
'            If vRet = "" Then
'               GoTo Exit_rtn
'            End If
'        End If
'
'    Else
'        ' *** 地主契約マスタデータ参照 ***
'        strSql = " SELECT "
'        strSql = strSql & " JARG_BKNO, "
'        strSql = strSql & " JARG_JCODE, "
'        strSql = strSql & " JARG_KYDATE "
'        strSql = strSql & " FROM "
'        strSql = strSql & " dbo_JARG_FILE "
'        strSql = strSql & " WHERE"
'        strSql = strSql & " JARG_YCODE =" & Val(vCode)
'        strSql = strSql & " AND JARG_STDATE <= CDATE(#" & DATE & "#)"
'        strSql = strSql & " AND JARG_EDDATE >= CDATE(#" & DATE & "#)"
'        strSql = strSql & " AND ( JARG_KYDATE >= CDATE(#" & DATE & "#) OR JARG_KYDATE IS NULL )"
'        strSql = strSql & " ORDER by "
'        strSql = strSql & " JARG_JCODE "
'        Set objRs = pobjDB.OpenRecordset(strSql, dbOpenSnapshot)    ' 2006/01/05 ADD T.SUZUKI
''        Set pobjRs = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)  ' 2006/01/05 DEL T.SUZUKI
'
'' 2006/01/05 ADD T.SUZUKI Start
'        If Not (objRs.BOF And objRs.EOF) Then
'            Me.txt_物件番号 = objRs![JARG_BKNO]
'            pvarKYDATE = objRs![JARG_KYDATE]
'
'            strSql = " JINU_CODE =" & objRs![JARG_JCODE]
'            Me.txt_地主名 = Nz(DLookup("JINU_NAME", "dbo_JINU_MAST", strSql), "")
'' 2006/01/05 ADD T.SUZUKI End
'' 2006/01/05 DEL T.SUZUKI Start
''        If Not (pobjRs.BOF And pobjRs.EOF) Then
''
'''            Me.txt_地主コード = pobjRs![JARG_JCODE]                ' 2005/08/03 DEL T.SUZUKI
''            Me.txt_物件番号 = pobjRs![JARG_BKNO]
''            pvarKYDATE = pobjRs![JARG_KYDATE]
''
''            strSQL = " JINU_CODE =" & pobjRs![JARG_JCODE]           ' 2005/08/03 ADD T.SUZUKI
'''            strSQL = strSQL & " JINU_CODE =" & Me.txt_地主コード   ' 2005/08/03 DEL T.SUZUKI
''            Me.txt_地主名 = Nz(DLookup("JINU_NAME", "dbo_JINU_MAST", strSQL), "")
'' 2006/01/05 DEL T.SUZUKI End
'        Else
'            objRs.Close                                              ' 2006/01/05 ADD T.SUZUKI
'            Set objRs = Nothing                                      ' 2006/01/05 ADD T.SUZUKI
'
'            strSql = " SELECT "
'            strSql = strSql & " JARG_BKNO , "
'            strSql = strSql & " JARG_JCODE , "
'            strSql = strSql & " JARG_KYDATE "
'            strSql = strSql & " FROM "
'            strSql = strSql & " dbo_JARG_FILE "
'            strSql = strSql & " WHERE "
'            strSql = strSql & " JARG_YCODE =" & Val(vCode)
'            strSql = strSql & " AND JARG_STDATE <= CDATE(#" & DATE & "#)"
'            strSql = strSql & " ORDER by "
'            strSql = strSql & " JARG_JCODE "
'            Set objRs = pobjDB.OpenRecordset(strSql, dbOpenSnapshot)    ' 2006/01/05 ADD T.SUZUKI
''            Set pobjRs = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)  ' 2006/01/05 DEL T.SUZUKI
'
'' 2006/01/05 ADD T.SUZUKI Start
'            If Not (objRs.BOF And objRs.EOF) Then
'                Me.txt_物件番号 = objRs![JARG_BKNO]
'                pvarKYDATE = objRs![JARG_KYDATE]
'
'                strSql = " JINU_CODE =" & objRs![JARG_JCODE]
'                Me.txt_地主名 = Nz(DLookup("JINU_NAME", "dbo_JINU_MAST", strSql), "")
'            End If
'' 2006/01/05 ADD T.SUZUKI End
'' 2006/01/05 DEL T.SUZUKI Start
''            If Not (pobjRs.BOF And pobjRs.EOF) Then
'''                Me.txt_地主コード = pobjRs![JARG_JCODE]  ' 2005/08/03 DEL T.SUZUKI
''                Me.txt_物件番号 = pobjRs![JARG_BKNO]
''                pvarKYDATE = pobjRs![JARG_KYDATE]
''
''                strSQL = " JINU_CODE =" & pobjRs![JARG_JCODE]  ' 2005/08/03 ADD T.SUZUKI
'''                strSQL = strSQL & " JINU_CODE =" & Me.txt_地主コード   ' 2005/08/03 DEL T.SUZUKI
''                Me.txt_地主名 = Nz(DLookup("JINU_NAME", "dbo_JINU_MAST", strSQL), "")
''            End If
'' 2006/01/05 DEL T.SUZUKI End
'        End If
'    End If
'
'    CHK_JARG_FILE = True
'
'Exit_rtn:
'    doCmd.Hourglass False
'    doCmd.SetWarnings True
'    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing       ' 2006/01/05 ADD T.SUZUKI
''    Set pobjRs = Nothing                                               ' 2006/01/05 DEL T.SUZUKI
'    Exit Function
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'    GoTo Exit_rtn
'End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナマスタ問い合わせ
'        MODULE_ID        :CHK_CNTA_MAST
'        CREATE_DATE      :2005/01/31
'
'        引数             :０ = 存在チェック
'                          １ = データ参照
'                          ２ = 予約情報表示
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_CNTA_MAST(ByVal iMode As Integer) As Boolean

    Dim vCode    As Variant
    Dim vRet     As Variant
    Dim stWHERE  As String
    Dim objRs    As Recordset

    On Error GoTo err_rtn

    CHK_CNTA_MAST = False

    vCode = Nz(Me.cmb_コンテナ番号, "")

    If vCode = "" Then
        CHK_CNTA_MAST = True
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False
  
    Select Case iMode
        Case 0
            '存在チェック
            stWHERE = " CNTA_CODE =" & Val(Me.cmb_ヤードコード)
            stWHERE = stWHERE & " AND CNTA_NO =" & Val(vCode)
            vRet = Nz(DLookup("CNTA_STEP", "dbo_CNTA_MAST", stWHERE), "")

            If vRet = "" Then
                MsgBox "コンテナ登録がされておりません", vbInformation, strMsg_kaku
                GoTo Exit_rtn
            End If

        Case 1
            'データ参照
            stWHERE = " SELECT * "
            stWHERE = stWHERE & " FROM dbo_CNTA_MAST"
            '' 2013/05/09 INS M.HONDA
            stWHERE = stWHERE & " INNER Join "
            stWHERE = stWHERE & " dbo_NAME_MAST ON "
            stWHERE = stWHERE & " dbo_CNTA_MAST.CNTA_USAGE = dbo_NAME_MAST.NAME_CODE "
            '' 2013/05/09 INS M.HONDA
            stWHERE = stWHERE & " where CNTA_CODE =" & Val(Me.cmb_ヤードコード)
            stWHERE = stWHERE & " AND CNTA_NO = " & Val(vCode)
            '' 2013/05/09 INS M.HONDA
            stWHERE = stWHERE & " AND ((dbo_NAME_MAST.NAME_ID)='086') "
            '' 2013/05/09 INS M.HONDA
            Set objRs = pobjDB.OpenRecordset(stWHERE, dbOpenSnapshot)

            If Not (objRs.BOF And objRs.EOF) Then

                ' 段区分名のみ表示
                If Nz(objRs![CNTA_STEP], 0) = 0 Then
                   Me.txt_段区分名 = "上段"
                Else
                   Me.txt_段区分名 = "下段"
                End If

                ' コンテナマスタ.実帖をセット
                vRet = Format(objRs![CNTA_SIZE], "0.00")

                Me.txt_コンテナ種別名 = vRet
                
                Me.txt_種別 = objRs![NAME_NAME] '' 2013/05/09 INS M.HONDA
                '' バイク系の場合には色を変える。
                If objRs![NAME_CODE] = 3 Or objRs![NAME_CODE] = 30 Or objRs![NAME_CODE] = 31 Or objRs![NAME_CODE] = 32 Or objRs![NAME_CODE] = 33 Then
                    Me.txt_種別.BackColor = 16776960
                Else
                    Me.txt_種別.BackColor = 14680063
                End If
                
 
                '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'                Me.txt_予約名 = objRs![CNTA_YNAME]
'                Me.txt_予約番号 = objRs![CNTA_YTEL]

                '2006/05/07 del Start Tajima コンテナマスタからは取れないデータ
'                Me.txt_初回契約日 = Format(DATE, P_FORMAT_YMD)      ' 2006/01/05 ADD T.SUZUKI
'                Me.txt_承認日 = Format(DATE, P_FORMAT_YMD)          ' 2006/01/05 ADD T.SUZUKI
'
'
'                If Nz(Me.cmb_鍵区分, "") <> "" Then
'                    ' 鍵区分名取得
'                    vRet = fncSelectRyak("060", Me.cmb_鍵区分)      ' 2006/01/05 REP T.SUZUKI
'                    Me.txt_鍵区分 = vRet
'                End If
'
'                If Me.cmb_保証区分 <> "" Then
'                    ' 保証区分名取得
'                    vRet = fncSelectRyak("200", Me.cmb_保証区分)    ' 2006/01/05 REP T.SUZUKI
'                    Me.txt_保証区分 = vRet
'                End If
                '2006/05/07 del End
          
            Else
                GoTo Exit_rtn
            End If

        Case 2
            '20050222 N.MIURA
            '以下のコーディングはCNTA_MASTのCNTA_YCODE（予約顧客コード）が未使用と思われる為
            '動作しない可能性がある。次回の修正にて対応とする。

            '予約情報表示
            stWHERE = ""
            stWHERE = stWHERE & " SELECT * "
            stWHERE = stWHERE & " FROM dbo_CNTA_MAST"
            stWHERE = stWHERE & " where CNTA_CODE =" & Val(Me.cmb_ヤードコード)
            stWHERE = stWHERE & " AND CNTA_NO = " & Val(vCode)
            stWHERE = stWHERE & " AND CNTA_YCODE <> 0"
            Set objRs = pobjDB.OpenRecordset(stWHERE, dbOpenSnapshot)

            If Not (objRs.BOF And objRs.EOF) Then
                Me.cmb_顧客コード = Format(objRs![CNTA_YCODE], "000000")
                Me.txt_電話番号 = objRs![CNTA_YTEL]
                Me.txt_月額料金 = objRs![CNTA_YRENTKG]
                'Me.txt_預り保証金 = objRs![CNTA_SECUKG]  2006/03/03 DEL Tajima

                ' 顧客データ取得
                Set objRs = pobjDB.OpenRecordset(SEL_USER_QUERY(Me.cmb_顧客コード), dbOpenSnapshot)

                If Not (objRs.BOF And objRs.EOF) Then
                    ' 顧客マスタより取得した値を画面に設定
                    With Me
                        .txt_顧客名 = objRs![USER_NAME]
                        .txt_顧客カナ = objRs![USER_KANA]
                        .txt_顧客名カナ = objRs![USER_KANA]
                        .txt_代表者名 = objRs![USER_TANM]
                        .txt_代表者名カナ = objRs![USER_TAKA]
                        .txt_郵便番号 = objRs![USER_YUBINO]
                        .txt_住所1 = objRs![USER_ADR_1]
                        .txt_住所2 = objRs![USER_ADR_2]
                        .txt_住所3 = objRs![USER_ADR_3]
                        .txt_メールアドレス = objRs![USER_MAIL]
                        .txt_電話番号 = objRs![USER_TEL]
                        .txt_携帯番号 = objRs![USER_KEITAI]
                        .txt_FAX番号 = objRs![USER_FAX]
                        .txt_締日 = objRs![USER_SDAY]
                        .cmb_顧客区分 = objRs![USER_KKBN]
                        .cmb_集金区分 = objRs![USER_SKBN]
                        .txt_顧客備考 = objRs![USER_BIKO]

                        ' 顧客区分名取得
                        vRet = fncSelectRyak("020", Me.cmb_顧客区分)           ' 2006/01/05 ADD T.SUZUKI
'                        vRet = fncSelectRyak(Me.cmb_顧客区分, "020")          ' 2006/01/05 DEL T.SUZUKI
                        .txt_顧客区分名 = vRet

                        ' 集金区分名取得
                        vRet = fncSelectRyak("030", Me.cmb_集金区分)           ' 2006/01/05 ADD T.SUZUKI
'                        vRet = fncSelectRyak(Me.cmb_集金区分, "030")          ' 2006/01/05 DEL T.SUZUKI
                        .txt_集金区分名 = vRet

                        If Nz(Trim(objRs![USER_RNAME]), "") <> "" Then
                            .chk_入力する = -1
                        End If

                        ' 【その他連絡先】を設定
                        .txt_連絡先担当名カナ = objRs![USER_RKANA]
                        .txt_連絡先担当者名 = objRs![USER_RNAME]
                        .txt_連絡先郵便番号 = objRs![USER_RPOST]
                        .txt_連絡先住所1 = objRs![USER_RADR_1]
                        .txt_連絡先住所2 = objRs![USER_RADR_2]
                        .txt_連絡先住所3 = objRs![USER_RADR_3]
                        .txt_連絡先電話番号 = objRs![USER_RTEL]
                        .txt_連絡先携帯番号 = objRs![USER_RKEITAI]
                        .txt_連絡先FAX = objRs![USER_RFAX]
                    End With
                End If

            Else
                '予約情報確認
                stWHERE = ""
                stWHERE = " CNTA_CODE =" & Val(Me.cmb_ヤードコード)
                stWHERE = stWHERE & " AND CNTA_NO =" & Val(vCode)
                ' 画面.顧客コードがNULLの場合があるので、下記のように対応
                stWHERE = stWHERE & " AND CNTA_YCODE =" & Val(Nz(Me.cmb_顧客コード, 0))
                vRet = Nz(DLookup("CNTA_STEP", "dbo_CNTA_MAST", stWHERE), "")

                 If vRet = "" Then GoTo Exit_rtn
            End If
    End Select
    
    CHK_CNTA_MAST = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

' 2006/01/05 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :コンテナマスタ問い合わせ
''        MODULE_ID        :CHK_CNTA_MAST
''        CREATE_DATE      :2005/01/31
''
''        引数             :０ = 存在チェック
''                          １ = データ参照
''                          ２ = 予約情報表示
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function CHK_CNTA_MAST(ByVal iMode As Integer) As Boolean
'
'    Dim vCode           As Variant
'    Dim vRet            As Variant
''    Dim strSQL          As String  ' 2005/08/03 DEL T.SUZUKI
'    Dim stWHERE         As String
'
'    On Error GoTo err_rtn
'
'    CHK_CNTA_MAST = False
'
'    vCode = Nz(Me.cmb_コンテナ番号, "")
'
'    If vCode = "" Then
'        CHK_CNTA_MAST = True
'        GoTo Exit_rtn
'    End If
'
'    doCmd.Hourglass True
'    doCmd.SetWarnings False
'
'    Select Case iMode
'        Case 0
'            '存在チェック
'            stWHERE = " CNTA_CODE =" & Val(Me.cmb_ヤードコード)
'            stWHERE = stWHERE & " AND CNTA_NO =" & Val(vCode)
'            vRet = Nz(DLookup("CNTA_STEP", "dbo_CNTA_MAST", stWHERE), "")
'
'            If vRet = "" Then
'                MsgBox "コンテナ登録がされておりません", vbInformation, strMsg_kaku
'                GoTo Exit_rtn
'            End If
'
'        Case 1
'            'データ参照
'            stWHERE = " SELECT * "
'            stWHERE = stWHERE & " FROM dbo_CNTA_MAST"
'            stWHERE = stWHERE & " where CNTA_CODE =" & Val(Me.cmb_ヤードコード)
'            stWHERE = stWHERE & " AND CNTA_NO = " & Val(vCode)
'            Set pobjRs = pobjDB.OpenRecordset(stWHERE, dbOpenSnapshot)
'
'            If Not (pobjRs.BOF And pobjRs.EOF) Then
'
'' 2005/08/03 REP T.SUZUKI Start
'' 段区分名のみ表示
'                If Nz(pobjRs![CNTA_STEP], 0) = 0 Then
'                   Me.txt_段区分名 = "上段"
'                Else
'                   Me.txt_段区分名 = "下段"
'                End If
''                Me.txt_段区分 = pobjRs![CNTA_STEP]
''                If Me.txt_段区分 = 0 Then
''                   Me.txt_段区分名 = "上段"
''                Else
''                   Me.txt_段区分名 = "下段"
''                End If
'' 2005/08/03 REP T.SUZUKI End
'
''                Me.txt_コンテナ種別 = pobjRs![CNTA_SYU]  ' 2005/08/03 DEL T.SUZUKI
'
'' 2005/08/03 REP T.SUZUKI Start
'                ' コンテナマスタ.実帖をセット
'                vRet = Format(pobjRs![CNTA_SIZE], "0.00")         ' 2005/11/19 ADD T.SUZUKI
''                ' コンテナ種別名取得
''                vRet = fncSelectRyak(pobjRs![CNTA_SYU], "040")   ' 2005/11/19 DEL T.SUZUKI
''                vRet = fncSelectRyak(pobjRs![CNTA_SYU], "040")   ' 2005/08/09 ADD T.SUZUKI
''                vRet = fncSelectName(pobjRs![CNTA_SYU], "040")  ' 2005/08/09 DEL T.SUZUKI
''                strSQL = " NAME_ID ='040'"
''                strSQL = strSQL & " AND NAME_CODE =" & Val(pobjRs![CNTA_SYU])     ' 2005/08/03 ADD T.SUZUKI
'''                strSQL = strSQL & " AND NAME_CODE =" & Val(Me.txt_コンテナ種別)  ' 2005/08/03 DEL T.SUZUKI
''                vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'' 2005/08/03 REP T.SUZUKI End
'
'                Me.txt_コンテナ種別名 = vRet
'                Me.txt_初回契約日 = Format(DATE, "YYYY/MM/DD")
'
'                Me.txt_予約名 = pobjRs![CNTA_YNAME]
'                Me.txt_予約番号 = pobjRs![CNTA_YTEL]
'
'                Me.txt_承認日 = Format(DATE, "YYYY/MM/DD")
'
'                If Nz(Me.cmb_鍵区分, "") <> "" Then
'' 2005/08/03 REP T.SUZUKI Start
'                    ' 鍵区分名取得
'                    vRet = fncSelectRyak(Me.cmb_鍵区分, "060")
''                    strSQL = " NAME_ID ='060'"
''                    strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_鍵区分)
''                    vRet = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strSQL), "")
'                    Me.txt_鍵区分 = vRet
'' 2005/08/03 REP T.SUZUKI End
'                End If
'
'                If Me.cmb_保証区分 <> "" Then
'' 2005/08/03 REP T.SUZUKI Start
'                    ' 保証区分名取得
'                    vRet = fncSelectRyak(Me.cmb_保証区分, "200")
''                    strSQL = " NAME_ID ='200'"
''                    strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_保証区分)
''                    vRet = Nz(DLookup("NAME_RYAK", "dbo_NAME_MAST", strSQL), "")
'                    Me.txt_保証区分 = vRet
'' 2005/08/03 REP T.SUZUKI End
'                End If
'            Else
'                GoTo Exit_rtn
'            End If
'
'        Case 2
'            '20050222 N.MIURA
'            '以下のコーディングはCNTA_MASTのCNTA_YCODE（予約顧客コード）が未使用と思われる為
'            '動作しない可能性がある。次回の修正にて対応とする。
'
'            '予約情報表示
'            stWHERE = ""
'            stWHERE = stWHERE & " SELECT * "
'            stWHERE = stWHERE & " FROM dbo_CNTA_MAST"
'            stWHERE = stWHERE & " where CNTA_CODE =" & Val(Me.cmb_ヤードコード)
'            stWHERE = stWHERE & " AND CNTA_NO = " & Val(vCode)
'            stWHERE = stWHERE & " AND CNTA_YCODE <> 0"
'            Set pobjRs = pobjDB.OpenRecordset(stWHERE, dbOpenSnapshot)
'
'            If Not (pobjRs.BOF And pobjRs.EOF) Then
'                Me.cmb_顧客コード = Format(pobjRs![CNTA_YCODE], "000000")  ' 2005/09/12 ADD T.SUZUKI
''                Me.cmb_顧客コード = pobjRs![CNTA_YCODE]                   ' 2005/09/12 DEL T.SUZUKI
'                Me.txt_電話番号 = pobjRs![CNTA_YTEL]
'                Me.txt_月額料金 = pobjRs![CNTA_YRENTKG]
'                Me.txt_預り保証金 = pobjRs![CNTA_SECUKG]
'
'                ' 顧客データ取得
'                Set pobjRs = pobjDB.OpenRecordset(SEL_USER_QUERY(Me.cmb_顧客コード), dbOpenSnapshot)
'
'                If Not (pobjRs.BOF And pobjRs.EOF) Then
'                    ' 顧客マスタより取得した値を画面に設定
'                    With Me
''                        .cmb_顧客カナ.Requery  ' 2005/10/19 DEL T.SUZUKI
'                        .txt_顧客名 = pobjRs![USER_NAME]
'                        .txt_顧客カナ = pobjRs![USER_KANA]   ' 2005/10/19 ADD T.SUZUKI
''                        .cmb_顧客カナ = pobjRs![USER_KANA]  ' 2005/10/19 DEL T.SUZUKI
'                        .txt_顧客名カナ = pobjRs![USER_KANA]
'                        .txt_代表者名 = pobjRs![USER_TANM]
'                        .txt_代表者名カナ = pobjRs![USER_TAKA]
'                        .txt_郵便番号 = pobjRs![USER_YUBINO]
'                        .txt_住所1 = pobjRs![USER_ADR_1]
'                        .txt_住所2 = pobjRs![USER_ADR_2]
'                        .txt_住所3 = pobjRs![USER_ADR_3]                        'INSERT 20050915 K.ISHIZAKA
'                        .txt_メールアドレス = pobjRs![USER_MAIL]                'INSERT 20050915 K.ISHIZAKA
'                        .txt_電話番号 = pobjRs![USER_TEL]                       'INSERT 20050915 K.ISHIZAKA
'                        .txt_携帯番号 = pobjRs![USER_KEITAI]
'                        .txt_FAX番号 = pobjRs![USER_FAX]
'                        .txt_締日 = pobjRs![USER_SDAY]
'                        .cmb_顧客区分 = pobjRs![USER_KKBN]
'                        .cmb_集金区分 = pobjRs![USER_SKBN]
'                        .txt_顧客備考 = pobjRs![USER_BIKO]
'
'' 2005/08/03 REP T.SUZUKI Start
'                        ' 顧客区分名取得
'                        vRet = fncSelectRyak(Me.cmb_顧客区分, "020")   ' 2005/08/09 ADD T.SUZUKI
''                        vRet = fncSelectName(Me.cmb_顧客区分, "020")  ' 2005/08/09 DEL T.SUZUKI
''                        strSQL = " NAME_ID ='020'"
''                        strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_顧客区分)
''                        vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'                        .txt_顧客区分名 = vRet
'' 2005/08/03 REP T.SUZUKI End
'
'' 2005/08/03 REP T.SUZUKI Start
'                        ' 集金区分名取得
'                        vRet = fncSelectRyak(Me.cmb_集金区分, "030")   ' 2005/08/09 ADD T.SUZUKI
''                        vRet = fncSelectName(Me.cmb_集金区分, "030")  ' 2005/08/09 DEL T.SUZUKI
''                        strSQL = " NAME_ID ='030'"
''                        strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_集金区分)
''                        vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'                        .txt_集金区分名 = vRet
'' 2005/08/03 REP T.SUZUKI End
'
'                        If Nz(Trim(pobjRs![USER_RNAME]), "") <> "" Then
'                            .chk_入力する = -1
'                        End If
'
'                        ' 【その他連絡先】を設定
'                        .txt_連絡先担当名カナ = pobjRs![USER_RKANA]
'                        .txt_連絡先担当者名 = pobjRs![USER_RNAME]
'                        .txt_連絡先郵便番号 = pobjRs![USER_RPOST]
'                        .txt_連絡先住所1 = pobjRs![USER_RADR_1]
'                        .txt_連絡先住所2 = pobjRs![USER_RADR_2]
'                        .txt_連絡先住所3 = pobjRs![USER_RADR_3]
'                        .txt_連絡先電話番号 = pobjRs![USER_RTEL]
'                        .txt_連絡先携帯番号 = pobjRs![USER_RKEITAI]
'                        .txt_連絡先FAX = pobjRs![USER_RFAX]
'                    End With
'                End If
'
'            Else
'                '予約情報確認
'                stWHERE = ""
'                stWHERE = " CNTA_CODE =" & Val(Me.cmb_ヤードコード)
'                stWHERE = stWHERE & " AND CNTA_NO =" & Val(vCode)
'                ' 画面.顧客コードがNULLの場合があるので、下記のように対応
'                stWHERE = stWHERE & " AND CNTA_YCODE =" & Val(Nz(Me.cmb_顧客コード, 0))
'                vRet = Nz(DLookup("CNTA_STEP", "dbo_CNTA_MAST", stWHERE), "")
'
'                If vRet = "" Then
'                    GoTo Exit_rtn
'                End If
'            End If
'    End Select
'
'    CHK_CNTA_MAST = True
'
'Exit_rtn:
'    doCmd.Hourglass False
'    doCmd.SetWarnings True
'    Set pobjRs = Nothing
'    Exit Function
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'    GoTo Exit_rtn
'End Functions
' 2006/01/05 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :顧客マスタ問い合わせ
'        MODULE_ID        :CHK_USER_MAST
'        CREATE_DATE      :
  '引数 ０：存在チェック
  '     １：データ参照
  '     ２：
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_USER_MAST(ByVal iMode As Integer) As Boolean

    Dim vCode    As Variant
    Dim vRet     As Variant
    Dim stWHERE  As String
    Dim objRs    As Recordset

    On Error GoTo err_rtn

    CHK_USER_MAST = False

    vCode = Nz(Me.cmb_顧客コード, "")

    If vCode = "" Then
        CHK_USER_MAST = True
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    If iMode = 0 Then
        ' 顧客マスタ存在チェック
        stWHERE = ""
        stWHERE = " USER_CODE =" & Val(vCode)
        vRet = Nz(DLookup("USER_NAME", "dbo_USER_MAST", stWHERE), "")

        If vRet = "" Then
            MsgBox "顧客マスタが登録されておりません", vbInformation, strMsg_kaku
            GoTo Exit_rtn
        End If

    Else
        ' 顧客データ取得
        Set objRs = pobjDB.OpenRecordset(SEL_USER_QUERY(vCode), dbOpenSnapshot)

        If Not (objRs.BOF And objRs.EOF) Then
            ' 顧客マスタより取得した値を画面に設定
            With Me
                .txt_顧客名 = objRs![USER_NAME]
                .txt_顧客名カナ = objRs![USER_KANA]
                .txt_顧客カナ = objRs![USER_KANA]
                .txt_代表者名 = objRs![USER_TANM]
                .txt_代表者名カナ = objRs![USER_TAKA]
                .txt_郵便番号 = objRs![USER_YUBINO]
                .txt_住所1 = objRs![USER_ADR_1]
                .txt_住所2 = objRs![USER_ADR_2]
                .txt_住所3 = objRs![USER_ADR_3]
                .txt_メールアドレス = objRs![USER_MAIL]
                .txt_電話番号 = objRs![USER_TEL]
                .txt_携帯番号 = objRs![USER_KEITAI]
                .txt_FAX番号 = objRs![USER_FAX]
                .txt_締日 = objRs![USER_SDAY]
                .cmb_顧客区分 = objRs![USER_KKBN]
                .cmb_集金区分 = objRs![USER_SKBN]
                .txt_顧客備考 = objRs![USER_BIKO]

                ' 顧客区分名取得
                vRet = fncSelectRyak("020", Me.cmb_顧客区分)    ' 2006/01/05 ADD T.SUZUKI
'                vRet = fncSelectRyak(Me.cmb_顧客区分, "020")   ' 2005/08/09 ADD T.SUZUKI
                .txt_顧客区分名 = vRet

                ' 集金区分名取得
                vRet = fncSelectRyak("030", Me.cmb_集金区分)    ' 2006/01/05 ADD T.SUZUKI
'                vRet = fncSelectRyak(Me.cmb_集金区分, "030")   ' 2005/08/09 ADD T.SUZUKI
                .txt_集金区分名 = vRet

                ' 顧客マスタ.連絡先担当者名 IS NOT NULL のとき、入力するチェックボックスを選択状態に設定
                If Nz(Trim(objRs![USER_RNAME]), "") <> "" Then
                    .chk_入力する = -1
                Else
                    .chk_入力する = 0
                End If

                ' 【その他連絡先】を設定
                .txt_連絡先担当名カナ = objRs![USER_RKANA]
                .txt_連絡先担当者名 = objRs![USER_RNAME]
                .txt_連絡先郵便番号 = objRs![USER_RPOST]
                .txt_連絡先住所1 = objRs![USER_RADR_1]
                .txt_連絡先住所2 = objRs![USER_RADR_2]
                .txt_連絡先住所3 = objRs![USER_RADR_3]
                .txt_連絡先電話番号 = objRs![USER_RTEL]
                .txt_連絡先携帯番号 = objRs![USER_RKEITAI]
                .txt_連絡先FAX = objRs![USER_RFAX]
            End With
        End If
    End If

    CHK_USER_MAST = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

' 2006/01/05 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :顧客マスタ問い合わせ
''        MODULE_ID        :CHK_USER_MAST
''        CREATE_DATE      :
'  '引数 ０：存在チェック
'  '     １：データ参照
'  '     ２：
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function CHK_USER_MAST(ByVal iMode As Integer) As Boolean
'
'    Dim vCode           As Variant
'    Dim vRet            As Variant
''    Dim strSQL          As String  ' 2005/08/03 DEL T.SUZUKI
'    Dim stWHERE         As String
'
'    On Error GoTo err_rtn
'
'    CHK_USER_MAST = False
'
'    vCode = Nz(Me.cmb_顧客コード, "")
'
'    If vCode = "" Then
'        CHK_USER_MAST = True
'        GoTo Exit_rtn
'    End If
'
'    doCmd.Hourglass True
'    doCmd.SetWarnings False
'
'    If iMode = 0 Then
'        ' 顧客マスタ存在チェック
'        stWHERE = ""
'        stWHERE = " USER_CODE =" & Val(vCode)
'        vRet = Nz(DLookup("USER_NAME", "dbo_USER_MAST", stWHERE), "")
'
'        If vRet = "" Then
'            MsgBox "顧客マスタが登録されておりません", vbInformation, strMsg_kaku
'            GoTo Exit_rtn
'        End If
'
'    Else
'        ' 顧客データ取得
'        Set pobjRs = pobjDB.OpenRecordset(SEL_USER_QUERY(vCode), dbOpenSnapshot)
'
'        If Not (pobjRs.BOF And pobjRs.EOF) Then
'            ' 顧客マスタより取得した値を画面に設定
'            With Me
''                .cmb_顧客カナ.Requery               ' 2005/10/19 DEL T.SUZUKI
'                .txt_顧客名 = pobjRs![USER_NAME]
'                .txt_顧客名カナ = pobjRs![USER_KANA]
'                .txt_顧客カナ = pobjRs![USER_KANA]   ' 2005/10/19 ADD T.SUZUKI
''                .cmb_顧客カナ = pobjRs![USER_KANA]  ' 2005/10/19 DEL T.SUZUKI
'                .txt_代表者名 = pobjRs![USER_TANM]
'                .txt_代表者名カナ = pobjRs![USER_TAKA]
'                .txt_郵便番号 = pobjRs![USER_YUBINO]
'                .txt_住所1 = pobjRs![USER_ADR_1]
'                .txt_住所2 = pobjRs![USER_ADR_2]
'                .txt_住所3 = pobjRs![USER_ADR_3]
'                .txt_メールアドレス = pobjRs![USER_MAIL]
'                .txt_電話番号 = pobjRs![USER_TEL]
'                .txt_携帯番号 = pobjRs![USER_KEITAI]
'                .txt_FAX番号 = pobjRs![USER_FAX]
'                .txt_締日 = pobjRs![USER_SDAY]
'                .cmb_顧客区分 = pobjRs![USER_KKBN]
'' 2005/08/12 REP T.SUZUKI Start
'                .cmb_集金区分 = pobjRs![USER_SKBN]
''                .cmb_集金区分 = pobjRs![USER_KKBN]
'' 2005/09/12 REP T.SUZUKI End
'                .txt_顧客備考 = pobjRs![USER_BIKO]
'
'' 2005/08/03 REP T.SUZUKI Start
'                ' 顧客区分名取得
'                vRet = fncSelectRyak(Me.cmb_顧客区分, "020")   ' 2005/08/09 ADD T.SUZUKI
''                vRet = fncSelectName(Me.cmb_顧客区分, "020")  ' 2005/08/09 DEL T.SUZUKI
''                strSQL = " NAME_ID ='020'"
''                strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_顧客区分)
''                vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'                .txt_顧客区分名 = vRet
'' 2005/08/03 REP T.SUZUKI End
'
'' 2005/08/03 REP T.SUZUKI Start
'                ' 集金区分名取得
'                vRet = fncSelectRyak(Me.cmb_集金区分, "030")   ' 2005/08/09 ADD T.SUZUKI
''                vRet = fncSelectName(Me.cmb_集金区分, "030")  ' 2005/08/09 DEL T.SUZUKI
''                strSQL = " NAME_ID ='030'"
''                strSQL = strSQL & " AND NAME_CODE =" & Val(Me.cmb_集金区分)
''                vRet = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'                .txt_集金区分名 = vRet
'' 2005/08/03 REP T.SUZUKI End
'
'                ' 顧客マスタ.連絡先担当者名 IS NOT NULL のとき、入力するチェックボックスを選択状態に設定
'                If Nz(Trim(pobjRs![USER_RNAME]), "") <> "" Then
'                    .chk_入力する = -1
'' 2005/09/12 ADD T.SUZUKI Start
'                Else
'                    .chk_入力する = 0
'' 2005/09/12 ADD T.SUZUKI End
'                End If
'
'                ' 【その他連絡先】を設定
'                .txt_連絡先担当名カナ = pobjRs![USER_RKANA]
'                .txt_連絡先担当者名 = pobjRs![USER_RNAME]
'                .txt_連絡先郵便番号 = pobjRs![USER_RPOST]
'                .txt_連絡先住所1 = pobjRs![USER_RADR_1]
'                .txt_連絡先住所2 = pobjRs![USER_RADR_2]
'                .txt_連絡先住所3 = pobjRs![USER_RADR_3]
'                .txt_連絡先電話番号 = pobjRs![USER_RTEL]
'                .txt_連絡先携帯番号 = pobjRs![USER_RKEITAI]
'                .txt_連絡先FAX = pobjRs![USER_RFAX]
'            End With
'        End If
'    End If
'
'    CHK_USER_MAST = True
'
'Exit_rtn:
'    doCmd.Hourglass False
'    doCmd.SetWarnings True
'    Set pobjRs = Nothing
'    Exit Function
'
'err_rtn:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'    GoTo Exit_rtn
'End Function
' 2006/01/05 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CHK_REQU_FILE
'        CREATE_DATE      :
'        UPDATE_DATE      :2006/03/18 T.SUZUKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_REQU_FILE(ByVal iMode As Integer) As Boolean

    Dim strWhere        As String
    Dim vCode           As Variant
    Dim vRet            As Variant
    Dim strSQL          As String

    On Error GoTo err_rtn

    CHK_REQU_FILE = False

    vCode = Me.txt_受注契約番号
    If vCode = "" Then
        CHK_REQU_FILE = True
        GoTo Exit_rtn
    End If

    doCmd.Hourglass True
    doCmd.SetWarnings False

    If iMode = 0 Then
        strWhere = " REQU_ACPTNO = '" & vCode & "'"
        vRet = Nz(DCount("REQU_REQNO", "dbo_REQU_FILE", strWhere), 0)

        If vRet = 0 Then
            GoTo Exit_rtn
        End If
    End If

    CHK_REQU_FILE = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約ファイル問い合わせ
'        MODULE_ID        :CHK_CARG_FILE
'        CREATE_DATE      :
'
'                          引数 ０：存在チェック
'                               １：データ参照
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_CARG_FILE(ByVal iMode As Integer) As Boolean

    Dim vCode   As Variant
    Dim vRet    As Variant
    Dim objRs   As Recordset

    On Error GoTo err_rtn

    CHK_CARG_FILE = False

    doCmd.Hourglass True
    doCmd.SetWarnings False

    'コンテナ状況区分
    pvarContna = -1
    pvarKcode = ""

    If iMode = 0 Then
        ' *** 存在チェック ***
        Me.txt_受注契約番号 = ""

        Set objRs = pobjDB.OpenRecordset(SEL_CARG_QUERY(Me.cmb_ヤードコード, Me.cmb_コンテナ番号), dbOpenSnapshot)

        If Not (objRs.BOF And objRs.EOF) Then
            ' 契約データ有
            pvarKcode = objRs![CARG_UCODE]
' 2006/01/30 DEL T.SUZUKI Start
'
'            If (objRs![CARG_STDATE] <= DATE) And _
'               (objRs![CARG_KYDATE] >= DATE Or IsNull(objRs![CARG_KYDATE])) Then
'                CHK_CARG_FILE = True
'                GoTo Exit_rtn
'            End If
'
'            pvarContna = objRs![CARG_CONTNA]
'
'            If objRs![CARG_KYDATE] < DATE And objRs![CARG_CONTNA] = -1 Then  ' 契約可能
'                GoTo Exit_rtn
'            End If
' 2006/01/30 DEL T.SUZUKI End
        Else
            '契約データ無し
            GoTo Exit_rtn
        End If

    Else
        ' *** データ取得処理 ***
        Set objRs = pobjDB.OpenRecordset(SEL_CARG_QUERY(Me.cmb_ヤードコード, Me.cmb_コンテナ番号), dbOpenSnapshot)

        If Not (objRs.BOF And objRs.EOF) Then
            ' 契約データ有
            pvarContna = objRs![CARG_CONTNA]

' 2006/03/04 DEL T.SUZUKI Start
'            If Me.opg_処理選択 = 1 Then
'' 2006/01/05 ADD T.SUZUKI Start
'                ' 移動先データの契約入力の場合、処理を抜ける
'                If (objRs![CARG_ACPTNO] = SAKI_YOUK_TRAN.strMOTO_ACPTNO) Then
'                    GoTo Exit_rtn
'                End If
'' 2006/01/05 ADD T.SUZUKI End
'
'                If (objRs![CARG_KYDATE] < DATE) And _
'                    objRs![CARG_CONTNA] = -1 Then      ' 契約可能
'                    GoTo Exit_rtn
'                End If
'            End If
' 2006/03/04 DEL T.SUZUKI End

            pintAGRE = Nz(objRs![CARG_AGRE], 0)        ' 契約状態

' 2006/01/05 ADD T.SUZUKI Start
            ' 契約状態がメンテ中または完了以外のとき、【契約変更】ボタンの使用を可能とする
            If Format(pintAGRE, "00") = P_AGRE_04 Or Format(pintAGRE, "00") = P_AGRE_09 Then
                Me.cmd_KeiyakuRep.Enabled = False
            Else
                Me.cmd_KeiyakuRep.Enabled = True
            End If
' 2006/01/05 ADD T.SUZUKI End

            ' 受付番号セット
            'Me.cmb_受付番号 = Nz(objRs![CARG_UKNO], "")  DEL 2006/03/03 Tajima

            ' コンテナ契約ファイル.受付番号
            pstrUKNO = Nz(Trim(objRs![CARG_UKNO]), "")

            ' コンテナ契約ファイル.受付担当者コード
            Me.cmb_受付担当者コード = objRs![CARG_UKTANTO]
            Call cmb_受付担当者コード_AfterUpdate

            Me.txt_受注契約番号 = objRs![CARG_ACPTNO]
            pstrJUCHUKEIYAKU_NO = objRs![CARG_ACPTNO]

            Me.txt_初回契約日 = Format(objRs![CARG_FSDATE], P_FORMAT_YMD)      ' 2006/01/05 REP T.SUZUKI
            svFSDATE = objRs![CARG_FSDATE]

            Me.txt_契約開始日 = Format(objRs![CARG_STDATE], P_FORMAT_YMD)      ' 2006/01/05 REP T.SUZUKI
            Me.txt_契約満了日 = Format(objRs![CARG_EDDATE], P_FORMAT_YMD)      ' 2006/01/05 REP T.SUZUKI

' 2006/01/05 REP T.SUZUKI Start
            Me.txt_解約日付 = Format(objRs![CARG_KYDATE], P_FORMAT_YMD)
            Me.txt_解約日付.Tag = Format(objRs![CARG_KYDATE], P_FORMAT_YMD)  ' ダミーとして解約日付のタグにコンテナ契約ファイル.解約日付を設定
' 2006/01/05 REP T.SUZUKI End

            If Nz(objRs![CARG_KAIUKEDATE], "") = "" Then
                Me.txt_解約受付日 = Null
            Else
                Me.txt_解約受付日 = Format(objRs![CARG_KAIUKEDATE], P_FORMAT_YMD)      ' 2006/01/05 REP T.SUZUKI
            End If

            Me.cmb_解約区分 = objRs![CARG_KAICD]
            Call cmb_解約区分_AfterUpdate
            Me.cmb_解約理由 = objRs![CARG_KAIRIYUCD]
            Call cmb_解約理由_AfterUpdate
            Me.cmb_解約受付担当者コード = objRs![CARG_KAITANTO]
            Call cmb_解約受付担当者コード_AfterUpdate

            Me.cmb_鍵区分 = Format(objRs![CARG_KAGIICD], "0")
            Call cmb_鍵区分_AfterUpdate
            Me.txt_解除番号 = objRs![CARG_KAIJB]
            '2006/05/07 ADD Start
            Me.txt_貸出本数 = objRs![CARG_KEY_LENTNUM]
            Me.txt_スペアキー備考 = objRs![CARG_SPAREKEY_NOTE]
            ' 読込時のデータ保持
            pOriginalKeyType = Nz(Me.cmb_鍵区分, "")
            pOriginalKeyNo = Nz(Me.txt_解除番号, "")
            pOriginalKeyNum = Nz(Me.txt_貸出本数, 0)
            '2006/05/07 ADD End

            Me.chk_書類1 = objRs![CARG_DOCU1]
' DEL 2006/03/03 Tajima Start
'            Me.chk_書類2 = objRs![CARG_DOCU2]
'            Me.txt_戻り日 = Format(objRs![CARG_MDATE], P_FORMAT_YMD)                   ' 2006/01/05 REP T.SUZUKI
'            Me.txt_番号 = objRs![CARG_DAHIB]
' DEL 2006/03/03 Tajima Start

            Me.txt_承認日 = objRs![CARG_HOSYD]
            Me.txt_承認番号 = objRs![CARG_HOSYB]
            
            Me.cmb_請求書発行区分 = objRs![CARG_SEIKI]                          'INSERT 2010/04/10 K.ISHIZAKA
            Call cmb_請求書発行区分_AfterUpdate                                 'INSERT 2010/04/10 K.ISHIZAKA
            
            Me.cmb_保証区分 = Format(objRs![CARG_HOSYICD], "0")
            Call cmb_保証区分_AfterUpdate
            Me.cmb_保証会社コード = objRs![CARG_HOSYO_CD]             ' 2009/04/01 INS KIHARA
            Call cmb_保証会社コード_AfterUpdate                       ' 2009/04/01 INS KIHARA

            Me.txt_月額料金 = objRs![CARG_RENTKG]
            Me.txt_消費税 = objRs![CARG_SYOZEI]
            Me.txt_合計 = objRs![CARG_RENTKG] + objRs![CARG_SYOZEI]
            Me.txt_預り保証金 = objRs![CARG_SECUKG]
            Me.txt_mente = objRs![MTANTM_TANTN]
            Me.txt_jun = objRs![JTANTM_TANTN]
            
          
            '' 2014/03/04 M.HONDA
            '' 受付トランに雑費コードを追加する予定なので、その時雑費コードを見るように直す
            '' 2013/02/22 M.HONDA START
            'If Nz(objRs![RCPT_EZAPPI], 0) = 450 Or Nz(objRs![RCPT_EZAPPI], 0) = 460 Or Nz(objRs![RCPT_EZAPPI], 0) = 480 Or Nz(objRs![RCPT_EZAPPI], 0) = 1080 Then                                      'DELETE 2018/02/10 N.IMAI
            'If Nz(objRs![RCPT_EZAPPI], 0) = 450 Or Nz(objRs![RCPT_EZAPPI], 0) = 460 Or Nz(objRs![RCPT_EZAPPI], 0) = 480 Or Nz(objRs![RCPT_EZAPPI], 0) = 1080 Or Nz(objRs![RCPT_EZAPPI], 0) = 2160 Then 'DEKETE 2019/02/19 N.IMAI INSERT 2018/02/10 N.IMAI
            '2020/12/15 [EGL] UPD Start
            'If Nz(objRs![RCPT_EZAPPI_CODE], 0) = 68 And Nz(objRs![RCPT_EZAPPI], 0) <> 0 Then                                                                                                            'INSERT 2019/02/19 N.IMAI
            '    Me.txt_補償加入 = "加入済"
            'Else
            '    Me.txt_補償加入 = "未加入"
            'End If
            '価格改定前はRCPT_PLAN_CDは 0000。念のため Null も考慮
            If Nz(objRs![RCPT_PLAN_CD], "0000") = "0000" Then
                If Nz(objRs![RCPT_EZAPPI_CODE], 0) = 68 And Nz(objRs![RCPT_EZAPPI], 0) <> 0 Then
                   Me.txt_補償加入 = "加入済"
                Else
                   Me.txt_補償加入 = "未加入"
                End If
            Else
                '価格改定後はRCPT_PLAN_CDに値が入っている
                If Nz(objRs![RCPT_EZAPPI_CODE], 0) = 68 Then
                   Me.txt_補償加入 = "加入済"
                Else
                   Me.txt_補償加入 = "未加入"
                End If
            End If
            '2020/12/15 [EGL] UPD End
            '' 2013/02/22 M.HONDA END
           Me.txt_Keitai = objRs![CODET_NAMEN]
           Me.txt_spe = objRs![SPARE]
           Me.txt_seq = objRs![SECUCOMP]

           Me.txt_RCPT_BIKE_NUMBER = objRs![RCPT_BIKE_NUMBER]
           Me.txt_RCPT_BIKE_SHASYU = objRs![RCPT_BIKE_SHASYU]

' DEL 2006/03/03 Tajima Start
'            Me.txt_初回金額 = objRs![CARG_FRSTKG]
'            Me.txt_初回鍵代 = objRs![CARG_KAGIA]
'            Me.txt_初回消費税 = objRs![CARG_FSYOZEI]
'            Me.txt_初回合計 = objRs![CARG_FRSTKG] + objRs![CARG_KAGIA] + objRs![CARG_FSYOZEI]
' DEL 2006/03/03 Tajima End

            Me.txt_契約備考 = objRs![CARG_BIKO]

            Me.cmb_顧客コード = Format(objRs![CARG_UCODE], "000000")
            pvarKcode = objRs![CARG_UCODE]
            
            ' 2006/03/03 ADD Tajima Start
            Me.txt_キャンペーン適用満了日 = Format(objRs![CARG_CMP_EXDATE], P_FORMAT_YMD)
            If Nz(Me.txt_キャンペーン適用満了日, "") <> "" Then
              Me.cmd_キャンペーン内容.Enabled = True
            End If
            ' 2006/03/03 ADD Tajima End

            ' ADD 2006/05/25 Tajima Start
            ' 受付トランの有無チェックとキャンペーン文言を取得しておく
            isRCPT_TRAN = False           '受付トラン有無を無しに
            pstrCampaignInfomationText = ""  'キャンペーン情報文言クリア
            pstrCampaignWarningText = ""     'キャンペーン警告文言クリア
            ' その他のテーブルから情報取得 ※受付トラン、セキュリティマスタなど
            isRCPT_TRAN = fncGetOtherTableData(Nz(objRs![CARG_UKNO], ""), Nz(objRs![CARG_ACPTNO], "")) '2005/05/24 Rep
            ' ADD 2006/05/25 Tajima End

            ' 顧客データ取得
            Set objRs = pobjDB.OpenRecordset(SEL_USER_QUERY(Me.cmb_顧客コード), dbOpenSnapshot)

            If Not (objRs.BOF And objRs.EOF) Then
                ' 顧客マスタより取得した値を画面に設定
                With Me
'                    .opg_顧客情報取得選択.Value = 2  DEL 2006/03/03 Tajima
                    .txt_顧客名 = objRs![USER_NAME]
                    .txt_電話番号 = objRs![USER_TEL]
                    .txt_顧客名カナ = objRs![USER_KANA]
                    .txt_顧客カナ = objRs![USER_KANA]
                    .txt_代表者名 = objRs![USER_TANM]
                    .txt_代表者名カナ = objRs![USER_TAKA]
                    .txt_郵便番号 = objRs![USER_YUBINO]
                    .txt_住所1 = objRs![USER_ADR_1]
                    .txt_住所2 = objRs![USER_ADR_2]
                    .txt_住所3 = objRs![USER_ADR_3]
                    .txt_メールアドレス = objRs![USER_MAIL]
                    .txt_携帯番号 = objRs![USER_KEITAI]
                    .txt_FAX番号 = objRs![USER_FAX]
                    .txt_締日 = objRs![USER_SDAY]
                    .cmb_顧客区分 = objRs![USER_KKBN]
                    .cmb_集金区分 = objRs![USER_SKBN]
                    .txt_顧客備考 = objRs![USER_BIKO]

                    ' 顧客区分名取得
                    vRet = fncSelectRyak("020", Me.cmb_顧客区分)           ' 2006/01/05 REP T.SUZUKI
                    .txt_顧客区分名 = vRet

                    ' 集金区分名取得
                    vRet = fncSelectRyak("030", Me.cmb_集金区分)           ' 2006/01/05 REP T.SUZUKI
                    .txt_集金区分名 = vRet

                    ' 顧客マスタ.連絡先担当者名 IS NOT NULL のとき、入力するチェックボックスを選択状態に設定
                    If Nz(Trim(objRs![USER_RNAME]), "") <> "" Then
                        .chk_入力する = -1
                    End If

                    ' 【その他連絡先】を設定
                    .txt_連絡先担当名カナ = objRs![USER_RKANA]
                    .txt_連絡先担当者名 = objRs![USER_RNAME]
                    .txt_連絡先郵便番号 = objRs![USER_RPOST]
                    .txt_連絡先住所1 = objRs![USER_RADR_1]
                    .txt_連絡先住所2 = objRs![USER_RADR_2]
                    .txt_連絡先住所3 = objRs![USER_RADR_3]
                    .txt_連絡先電話番号 = objRs![USER_RTEL]
                    .txt_連絡先携帯番号 = objRs![USER_RKEITAI]
                    .txt_連絡先FAX = objRs![USER_RFAX]
                End With
            End If
            
  
            
            
            '2006/09/22 SHIBAZAKI コンテナの直近修繕と顧客クレームを表示
            Call DispReprInfo
        Else
            '契約データ無し
            GoTo Exit_rtn
        End If
    End If

    CHK_CARG_FILE = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

err_rtn:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function


'2006/09/22 SHIBAZAKI　コンテナの直近修繕と顧客クレームを表示
'==============================================================================*
'
'        MODULE_NAME      :DispMntInfo
'        機能             :修繕情報表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub DispReprInfo()
On Error GoTo ErrRtn

    Dim strKagtosSqlserver      As String
    Dim strKagtosConnect        As String
    Dim dbKagtos                As Database
    Dim strBUMOC                As String
    Dim blnError                As Boolean
    
    blnError = False
    
    '加瀬ＤＢに接続
    strKagtosSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    strKagtosConnect = MSZZ007_M00()
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    Set dbKagtos = Workspaces(0).OpenDatabase(strKagtosSqlserver, dbDriverNoPrompt, False, strKagtosConnect)

    txt_直近修繕 = GetRoomLatestReprText(dbKagtos, strBUMOC, cmb_ヤードコード, cmb_コンテナ番号)
    txt_クレーム = GetUserLatestKotoText(dbKagtos, strBUMOC, cmb_顧客コード)
    
    dbKagtos.Close
    Set dbKagtos = Nothing
    
    GoTo ExitRtn

ErrRtn:
    blnError = True
    
ExitRtn:
    If Not dbKagtos Is Nothing Then dbKagtos.Close: Set dbKagtos = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncCheckYOUK_TRAN
'        CREATE_DATE      :2005/07/13
'        処理内容         ：予約受付トランの存在チェック
'        引数             ：varUKNO：受付番号
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckYOUK_TRAN(varUKNO As Variant) As Boolean

    Dim strSQL As String
    Dim vRet   As Variant

    On Error GoTo fncCheckYOUK_TRAN_Err

    fncCheckYOUK_TRAN = False

    If varUKNO <> "" Then
        strSQL = " YOUKT_UKNO = '" & varUKNO & "'"
        strSQL = strSQL & " AND YOUKT_YUKBN = 10 "
        vRet = Nz(DLookup("YOUKT_UKNO", "dbo_YOUK_TRAN", strSQL), "")

        If vRet = "" Then
            MsgBox "該当する予約受付トランがありません。", vbInformation, strMsg_kaku
            Exit Function
        End If
    End If

    fncCheckYOUK_TRAN = True

    Exit Function

fncCheckYOUK_TRAN_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Function

'==============================================================================*
'
'        MODULE_NAME      :移動予約チェック
'        MODULE_ID        :fncCheckMoveYoyaku
'        Parameter        :
'                          strACPTNO(IN) = 受注契約番号
'                          lngYCODE(OUT) = ヤードコード
'                          lngNO(OUT)    = コンテナ番号
'                          strUKNO(OUT)  = 予約受付番号
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckMoveYoyaku(ByVal strACPTNO As String, ByRef lngYCODE As Long, _
                                    ByRef lngNo As Long, ByRef strUkno As String) As Integer

    Dim strSQL As String
    Dim objRs  As Recordset

    On Error GoTo fncCheckMoveYoyaku_Err

    fncCheckMoveYoyaku = -1

    If Nz(strACPTNO, "") = "" Then Exit Function

    ' 初期化
    lngYCODE = 0
    lngNo = 0
    strUkno = ""

    ' 予約受付トランデータチェック
    strSQL = " SELECT   YOUKT_YCODE, " & Chr(13)  ' ヤードコード
    strSQL = strSQL & " YOUKT_NO, " & Chr(13)     ' コンテナ番号
    strSQL = strSQL & " YOUKT_UKNO " & Chr(13)    ' 予約受付番号
    strSQL = strSQL & "  FROM dbo_YOUK_TRAN " & Chr(13)
    strSQL = strSQL & " WHERE YOUKT_MOTO_ACPTNO = '" & strACPTNO & "' "  ' 元受注契約番号
    strSQL = strSQL & "   AND YOUKT_YUKBN       = " & P_YUKBN_10         ' 予約受付状態区分(10：受付)
    strSQL = strSQL & "   AND YOUKT_MOVEKBN    <> " & P_MOVEKBN_00       ' 移動区分(00：無し)
    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)

    With objRs
        Do Until .EOF
            
            lngYCODE = .Fields("YOUKT_YCODE")  ' ヤードコード
            lngNo = .Fields("YOUKT_NO")        ' コンテナ番号
            strUkno = .Fields("YOUKT_UKNO")    ' 予約受付番号

            objRs.MoveNext
        Loop
    End With
    objRs.Close
    Set objRs = Nothing

    If Nz(strUkno, "") <> "" Then
        ' 該当データあり
        fncCheckMoveYoyaku = 1
    Else
        ' 該当データなし
        fncCheckMoveYoyaku = 0
    End If

fncCheckMoveYoyaku_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

fncCheckMoveYoyaku_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncCheckMoveYoyaku_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :コンテナ契約ファイル問い合わせ
'        MODULE_ID        :fncChkCARG_FILE
'        引       数      ：strACPTNO(IN) = 受注契約番号
'        CREATE_DATE      :2005/06/17
'
'==============================================================================*
Public Function fncChkCARG_FILE(ByVal strACPTNO As String) As Boolean

    Dim objRs  As Recordset
    Dim strSQL As String

    On Error GoTo fncChkCARG_FILE_Err

    fncChkCARG_FILE = False

    Set pobjDB = CurrentDb

    doCmd.Hourglass True
    doCmd.SetWarnings False

    ' 存在チェック
    strSQL = ""
    strSQL = strSQL & "SELECT CARG_YCODE, " & Chr(13)  ' ヤードコード
    strSQL = strSQL & "       CARG_NO " & Chr(13)      ' コンテナ番号
    strSQL = strSQL & "  FROM dbo_CARG_FILE " & Chr(13)
    strSQL = strSQL & " WHERE CARG_ACPTNO = '" & strACPTNO & "' "
    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)

    If objRs.Fields(0).VALUE = 0 Then
        GoTo fncChkCARG_FILE_Exit
    Else
        pstrYCode = objRs.Fields("CARG_YCODE") ' ヤードコード
        pstrNo = objRs.Fields("CARG_NO")       ' コンテナ番号
    End If
    fncChkCARG_FILE = True

fncChkCARG_FILE_Exit:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    Set objRs = Nothing
    pobjDB.Close
    Set pobjDB = Nothing
    Exit Function

fncChkCARG_FILE_Err:
    MsgBox "ｴﾗｰ番号" & Err.numberr & vbCrLf & Err.Description
    Err.Clear
    GoTo fncChkCARG_FILE_Exit
End Function

'==============================================================================*
'
'        MODULE_NAME      :移動元契約データチェック
'        MODULE_ID        fncCheckMoveMotoData
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCheckMoveMotoData() As Boolean

    fncCheckMoveMotoData = False

    ' 予約受付トランからの読み込み時
    If (Nz(SAKI_YOUK_TRAN.strMOTO_ACPTNO, "") <> "" And SAKI_YOUK_TRAN.strMOVEKBN <> P_MOVEKBN_00) Then
        ' 移動契約データの場合
'''        ' ①移動元コンテナ契約ファイル存在チェック
'''        If fncGetCARG_AGRE(Nz(SAKI_YOUK_TRAN.strMOTO_ACPTNO, "")) = False Then
'''            GoTo fncCheckMoveMotoData_ChkErr
'''        End If

        ' ②移動元コンテナ契約ファイル.契約状態チェック
        If (MOTO_CARG_FILE.strAGRE = P_AGRE_01 Or MOTO_CARG_FILE.strAGRE = P_AGRE_02 Or MOTO_CARG_FILE.strAGRE = P_AGRE_03) Then
            ' チェックOK
        Else
            ' チェックNG
            GoTo fncCheckMoveMotoData_ChkErr
        End If

    Else
        ' 通常契約データ
    End If

    fncCheckMoveMotoData = True

fncCheckMoveMotoData_ChkExit:
    Exit Function

fncCheckMoveMotoData_ChkErr:
    MsgBox strMSGMoveError, vbInformation, strMsg_err
    Exit Function

fncCheckMoveMotoData_Err:
    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncCheckMoveMotoData_ChkExit
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :WRI_RTN
'        Parameter        :intLbn
'                                 0 = 修正モードで保存
'                                 1 = 削除モードで保存
'        CREATE_DATE      :
'
'        処理内容         ：コンテナ契約ファイルの新規・修正・削除登録
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function WRI_RTN(ByVal intLbn As Integer) As Boolean

    Dim blnRTN        As Boolean
    Dim strKIYAKU_NO  As String   ' 受注契約番号
    Dim lngREQU_NO    As Long     ' 請求NO
    Dim strBumoCD     As String   ' 部門コード
    Dim strErrMsg     As String   ' エラーメッセージ格納

    Dim strUserName   As String   ' ユーザー名            ' 2006/02/03 ADD T.SUZUKI

    On Error GoTo err_rtn

    WRI_RTN = False

    doCmd.Hourglass True
    doCmd.SetWarnings False

    ' 部門コード／部門名取得
    Call subGetBumonName(pobjDB, strBumoCD)

    ' コンテナDB接続
    Call subConnectDB(strBumoCD)

    ' ユーザー名取得
    strUserName = MSZZ000.LsGetUserName                   ' 2006/02/03 ADD T.SUZUKI

' 2006/03/04 ADD T.SUZUKI Start
    Select Case intLbn
      Case 0  ' *** 修正 ***
        If WRI_RTN_For_CARG_FILE(2, strUserName) = False Then GoTo Exit_rtn  ' 2006/02/03 ADD T.SUZUKI

      Case 1  ' *** 削除 ***
        If WRI_RTN_For_CARG_FILE(1, strUserName) = False Then GoTo Exit_rtn  ' 2006/02/03 ADD T.SUZUKI
    End Select
' 2006/03/04 ADD T.SUZUKI End

' 2006/03/04 DEL T.SUZUKI Start
'    Select Case Me.opg_処理選択.Value
'        Case 1  ' *** 新規 ***
'            ' 初回契約日チェック
'            If Not CHK_CARG_FILE2 Then GoTo Exit_rtn
'
'            ' 受注契約番号発番
'            blnRtn = WRI_RTN_For_CTL_MST(strKIYAKU_NO, lngREQU_NO)
'            If Not blnRtn Then
'                GoTo Exit_rtn
'            End If
'
''DEL 2006/03/03 Tajima Start
''            ' 新規登録 かつ 新規顧客の場合の処理 --------------------------------------
''            If Me.opg_顧客情報取得選択.Value = 1 Then
''                ' コントロールマスタデータ抽出 & データ更新処理
''                If fncSEL_UPD_CONT_MAST = False Then
''                    Exit Function
''                End If
''
''                ' 顧客マスタ登録処理
''                If fncINS_USER_MAST = False Then
''                    Exit Function
''                End If
''            End If
''            ' ----------------------------------------------------------------------
''DEL 2006/03/03 Tajima End
'
'            ' 新規登録処理
'            If WRI_RTN_For_CARG_FILE(0, strUserName, strKIYAKU_NO, lngREQU_NO) = False Then GoTo Exit_rtn  ' 2006/02/03 ADD T.SUZUKI
''            If WRI_RTN_For_CARG_FILE(0, strKIYAKU_NO, lngREQU_NO) = False Then GoTo Exit_rtn              ' 2006/02/03 DEL T.SUZUKI
'            plngUSER_NO = 0  ' 顧客コード初期化(新規顧客時使用)
'
'' 2006/01/05 ADD T.SUZUKI Start
'            ' 予約受付トランからのデータ読み込み時
'            If (Nz(SAKI_YOUK_TRAN.strMOTO_ACPTNO, "") <> "" And SAKI_YOUK_TRAN.strMOVEKBN <> P_MOVEKBN_00) Then
'                strErrMsg = "": blnRtn = True
'                Select Case SAKI_YOUK_TRAN.strMOVEKBN
'                    Case P_MOVEKBN_01  ' *** 自社理由移動 ***
'                        ' → 旧契約のヤード・コンテナで自社理由移動のデータを作成
'                        strErrMsg = ""
'                        blnRtn = CmKomsMod.AddnReserveType1forCancelYard(objKONT_DB, MOTO_CARG_FILE.strYCODE, _
'                                                                         MOTO_CARG_FILE.strNO, _
'                                                                         MOTO_CARG_FILE.strUKTANTO, _
'                                                                         Me.Form.NAME, strUserName, strErrMsg)             ' 2006/02/03 ADD T.SUZUKI
''                                                                         Me.Form.NAME, MSZZ000.LsGetUserName, strErrMsg)  ' 2006/02/03 DEL T.SUZUKI
'
'                    Case P_MOVEKBN_04  ' *** 名義変更 ***
'                        blnRtn = CmKomsMod.UpdateOldUserMast(objKONT_DB, Me.cmb_顧客コード, _
'                                                             MOTO_CARG_FILE.strUCODE, strErrMsg)
'                End Select
'
'                If blnRtn = False Then
'                    ' エラー時
'                    MsgBox strErrMsg, vbInformation, strMsg_err
'                End If
'            End If
'' 2006/01/05 ADD T.SUZUKI End
'
'        Case 2  ' *** 修正 ***
'            ' 修正登録処理
'            If WRI_RTN_For_CARG_FILE(2, strUserName) = False Then GoTo Exit_rtn  ' 2006/02/03 ADD T.SUZUKI
''            If WRI_RTN_For_CARG_FILE(2) = False Then GoTo Exit_rtn              ' 2006/02/03 DEL T.SUZUKI
'
'        Case 3  ' *** 削除 ***
'            ' 削除登録処理
'            If WRI_RTN_For_CARG_FILE(1, strUserName) = False Then GoTo Exit_rtn  ' 2006/02/03 ADD T.SUZUKI
''            If WRI_RTN_For_CARG_FILE(1) = False Then GoTo Exit_rtn              ' 2006/02/03 DEL T.SUZUKI
'    End Select
' 2006/03/04 DEL T.SUZUKI End

    ' DB切断
    Call subDisConnectDB

    WRI_RTN = True

Exit_rtn:
    If Not objKONT_DB Is Nothing Then
        ' コンテナDB切断
        Call subDisConnectDB
    End If

    doCmd.Hourglass False
    doCmd.SetWarnings True

    Set pWRI_objDb = Nothing
    Set pWRI_ObjWS = Nothing
    Exit Function

err_rtn:
    MsgBox "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbCritical + vbOKOnly, strMsg_err
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :更新前の入力チェック
'        MODULE_ID        :CHK_RTN_For_UpDate
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_RTN_For_UpDate() As Boolean

    Dim strKaiyaku  As String
    Dim strzenKaiyaku As String
    Dim strzengesyo As String

    CHK_RTN_For_UpDate = False

'>>> INSERT START 20050916 K.ISHIZAKA >>>
    ' ヤードコード
    If Val(Nz(Me.cmb_ヤードコード)) = 0 Then
        MsgBox "ヤードコードが正しく入力されていません。", vbInformation, strMsg_kaku
        If Me.cmb_ヤードコード.Enabled = True Then Me.cmb_ヤードコード.SetFocus
        Exit Function
    End If

    ' コンテナ番号
    If Val(Nz(Me.cmb_コンテナ番号)) = 0 Then
        MsgBox "コンテナ番号が正しく入力されていません。", vbInformation, strMsg_kaku
        If Me.cmb_コンテナ番号.Enabled = True Then Me.cmb_コンテナ番号.SetFocus
        Exit Function
    End If
'<<< INSERT END   20050916 K.ISHIZAKA <<<

    ' 受付担当者コード
    If Nz(Me.cmb_受付担当者コード, "") = "" Then
        MsgBox "受付担当者コードが正しく入力されていません。", vbInformation, strMsg_kaku
        If Me.cmb_受付担当者コード.Enabled = True Then Me.cmb_受付担当者コード.SetFocus
        Exit Function
    End If

' DEL 2006/03/03 Tajima Start 受付対応に伴い顧客入力は行わないのでチェックしない
'    ' 顧客情報関連入力チェック
'    If (Me.opg_処理選択 = 1 And Me.opg_顧客情報取得選択 = 2) Or _
'       (Me.opg_処理選択 = 2 And Me.opg_顧客情報取得選択 = 2) Then
''    If Not (Me.opg_処理選択.Value = 1 And Me.opg_顧客情報取得選択.Value = 1) Then
'        ' *** 既存顧客の場合の入力チェック ***
'        ' ①新規 - 既存顧客
'        ' ②修正 - 既存顧客
'
'        ' 顧客コード
'        If Nz(Me.cmb_顧客コード, "") = "" Then
'            MsgBox "顧客コードが正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.cmb_顧客コード.SetFocus
'            Exit Function
'        End If
'
'        If Val(Nz(Me.cmb_顧客コード, 0)) = 0 Then
'            MsgBox "顧客コードが正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.cmb_顧客コード.SetFocus
'            Exit Function
'        End If
'
'' 2005/08/03 ADD T.SUZUKI Start
'    ElseIf (Me.opg_処理選択 = 1 And Me.opg_顧客情報取得選択.Value = 1) Then
'        ' *** 新規顧客の場合の入力チェック ***
'        ' 顧客名
'        If Nz(Me.txt_顧客名, "") = "" Then
'            MsgBox "顧客名を入力して下さい。", vbInformation, strMsg_kaku
'            Me.txt_顧客名.SetFocus
'            Exit Function
'        End If
'
'        ' 顧客名カナ
'        If Nz(Me.txt_顧客名カナ, "") = "" Then
'            MsgBox "顧客名カナを入力して下さい。", vbInformation, strMsg_kaku
'            Me.txt_顧客名カナ.SetFocus
'            Exit Function
'        End If
'
'        ' 郵便番号
'        If Nz(Me.txt_郵便番号, "") = "" Then
'            MsgBox "郵便番号を入力して下さい。", vbInformation, strMsg_kaku
'            Me.txt_郵便番号.SetFocus
'            Exit Function
'        End If
'
'        ' 都道府県市町村
'        If Nz(Me.txt_住所1, "") = "" Then
'            MsgBox "都道府県市町村を入力して下さい。", vbInformation, strMsg_kaku
'            Me.txt_住所1.SetFocus
'            Exit Function
'        End If
'
'        ' 番地
'        If Nz(Me.txt_住所2, "") = "" Then
'            MsgBox "番地を入力して下さい。", vbInformation, strMsg_kaku
'            Me.txt_住所2.SetFocus
'            Exit Function
'        End If
'
'        ' 電話番号
'        If Nz(Me.txt_電話番号, "") = "" Then
'            MsgBox "電話番号を入力して下さい。", vbInformation, strMsg_kaku
'            Me.txt_電話番号.SetFocus
'            Exit Function
'        End If
'
'        ' 顧客区分
'        If Nz(Me.cmb_顧客区分, "") = "" Then
'            MsgBox "顧客区分を入力して下さい。", vbInformation, strMsg_kaku
'            Me.cmb_顧客区分.SetFocus
'            Exit Function
'        End If
'
'        ' 集金区分
'        If Nz(Me.cmb_集金区分, "") = "" Then
'            MsgBox "集金区分を入力して下さい。", vbInformation, strMsg_kaku
'            Me.cmb_集金区分.SetFocus
'            Exit Function
'        End If
'
'        ' 締日
'        If Nz(Me.txt_締日, "") = "" Then
'            MsgBox "締日を入力して下さい。", vbInformation, strMsg_kaku
'            Me.txt_締日.SetFocus
'            Exit Function
'        End If
'    End If
'' 2005/08/03 ADD T.SUZUKI End
' DEL 2006/03/03 Tajima End 受付対応に伴い顧客入力は行わないのでチェックしない

    ' 契約開始日
    If Nz(Me.txt_契約開始日, "") = "" Then
        MsgBox "契約開始日が正しく入力されていません。", vbInformation, strMsg_kaku
        Me.txt_契約開始日.SetFocus
        Exit Function
    End If

    ' 契約満了日
    If Nz(Me.txt_契約満了日, "") = "" Then
        MsgBox "契約満了日が正しく入力されていません。", vbInformation, strMsg_kaku
        Me.txt_契約満了日.SetFocus
        Exit Function
    End If

    '解約日付が入力された場合
    strKaiyaku = Nz(Me.txt_解約日付, "")

    If strKaiyaku <> "" _
        Or Nz(Me.cmb_解約区分, "") <> "01" Then   '2008/05/23 add tajima 解約区分"01"以外か
        '2008/05/23 del tajima ↑ORG→ Or Nz(Me.cmb_解約区分, "") <> "" Then   '2008/03/23 add tajima 解約区分入力済みか
        If Not IsDate(strKaiyaku) Then
            Me.txt_解約日付.SetFocus
            MsgBox "解約日付が正しく入力されていません。", vbInformation, strMsg_kaku
            Exit Function
        End If

        If StrComp(Format(DateAdd("m", "-1", Nz(Me.txt_契約開始日, "")), "yyyy/mm") & "/01", strKaiyaku) > 0 Then
            Me.txt_解約日付.SetFocus
            MsgBox "解約日付に「" & Format(DateAdd("m", "-1", Nz(Me.txt_契約開始日, "")), "yyyy/mm") & "/01」より前は指定できません。", _
                   vbInformation, _
                   strMsg_kaku
            Exit Function
        End If

' 2005/08/03 ADD T.SUZUKI Start
        If Nz(Me.txt_解約受付日, "") = "" Then
            MsgBox "解約受付日を入力して下さい。", vbInformation, strMsg_kaku
            Me.txt_解約受付日.SetFocus
            Exit Function
        End If

        If Nz(Me.cmb_解約区分, "") = "" Then
            MsgBox "解約区分を入力して下さい。", vbInformation, strMsg_kaku
            Me.cmb_解約区分.SetFocus
            Exit Function
        End If

        If Nz(Me.cmb_解約理由, "") = "" Then
            MsgBox "解約理由を入力して下さい。", vbInformation, strMsg_kaku
            Me.cmb_解約理由.SetFocus
            Exit Function
        End If

        If Nz(Me.cmb_解約受付担当者コード, "") = "" Then
            MsgBox "解約受付担当者コードを入力して下さい。", vbInformation, strMsg_kaku
            Me.cmb_解約受付担当者コード.SetFocus
            Exit Function
        End If
' 2005/08/03 ADD T.SUZUKI End

        '今月から３ヶ月以内であること
        If StrComp(strKaiyaku, Format(DateAdd("m", C_KAIYAKU_KIKAN_FROM + 1, Now()), "yyyy/mm") & "/01") < 0 Then
            Me.txt_解約日付.SetFocus
            MsgBox "解約日付に「" & Format(DateAdd("m", C_KAIYAKU_KIKAN_FROM + 1, Now()), "yyyy/mm") & "/01」より前は指定できません。", _
                   vbInformation, _
                   strMsg_kaku
            Exit Function
        End If

        '今月から後20ヶ月以内であること
        If StrComp(strKaiyaku, Format(DateAdd("m", C_KAIYAKU_KIKAN_TO, Now()), "yyyy/mm") & "/01") >= 0 Then
            Me.txt_解約日付.SetFocus
            MsgBox "解約日付に「" & Format(DateAdd("m", C_KAIYAKU_KIKAN_TO, Now()), "yyyy/mm") & "/01」以降は指定できません。", _
                   vbInformation, _
                   strMsg_kaku
            Exit Function
        End If
        
        '2006/09/22 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↓
        If txt_ヤード解約予定 <> "" Then
            If DateDiff("d", txt_ヤード解約予定, strKaiyaku) > 0 Then
                Me.txt_解約日付.SetFocus
                MsgBox "解約日はヤード解約予定日 [" & txt_ヤード解約予定 & "] 以前で無ければいけません。", vbInformation, strMsg_kaku
                Exit Function
            End If
        End If
        '2006/09/22 SHIBAZAKI 解約日の入力はヤード解約予定日までとする↑
        
        '2018/08/21 M.HONDA INS
        '当月以前に解約日は入力不可へ（未使用キャンセルのみ可能）
        '前月末日を求める。
'        strzenKaiyaku = DateSerial(Year(strKaiyaku), Month(strKaiyaku), 0)
'        strzengesyo = DateSerial(Year(DATE), Month(DATE), 1)
''        '未使用キャンセルの場合にはエラーとしない
'        If DateDiff("d", CDate(txt_契約開始日), CDate(strKaiyaku)) > 0 Then
'            If DateDiff("d", CDate(strzengesyo), CDate(strzenKaiyaku)) >= 0 Then
'                Me.txt_解約日付.SetFocus
'                MsgBox "売上が確定しているため解約日は当月以降でなければなりません。", vbInformation, strMsg_kaku
'                Exit Function
'            End If
'        End If
        '2018/08/21 M.HONDA INS
    End If

    If pintError = 1 Then
        MsgBox "請求処理がされているため変更できません。", vbInformation, strMsg_kaku
        Me.cmb_顧客コード.SetFocus
        Me.cmb_顧客コード = Format(pvarKcode, "000000")  ' 2005/09/12 ADD T.SUZUKI
'        Me.cmb_顧客コード = pvarKcode                   ' 2005/09/12 DEL T.SUZUKI
        pintError = 0
        Exit Function
    End If

' DEL 2006/03/03 Tajima Start 受付対応に伴い顧客入力は行わないのでチェックしない
'    ' 【その他連絡先】入力チェック
'    If Me.chk_入力する.Value = -1 Then
'        ' 連絡先担当者名カナ
'        If Nz(Me.txt_連絡先担当名カナ, "") = "" Then
'            MsgBox "連絡先担当名カナが正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.txt_連絡先担当名カナ.SetFocus
'            Exit Function
'        End If
'
'        ' 連絡先担当者名
'        If Nz(Me.txt_連絡先担当者名, "") = "" Then
'            MsgBox "連絡先担当名カナが正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.txt_連絡先担当者名.SetFocus
'            Exit Function
'        End If
'
'        ' 連絡先郵便番号
'        If Nz(Me.txt_連絡先郵便番号, "") = "" Then
'            MsgBox "連絡先郵便番号が正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.txt_連絡先郵便番号.SetFocus
'            Exit Function
'        End If
'
'        ' 連絡先住所１
'        If Nz(Me.txt_連絡先住所1, "") = "" Then
'            MsgBox "都道府県市町村が正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.txt_連絡先住所1.SetFocus
'            Exit Function
'        End If
'
'        ' 連絡先住所２
'        If Nz(Me.txt_連絡先住所2, "") = "" Then
'            MsgBox "番地が正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.txt_連絡先住所2.SetFocus
'            Exit Function
'        End If
'
'        ' 連絡先電話番号
'        If Nz(Me.txt_連絡先電話番号, "") = "" Then
'            MsgBox "連絡先電話番号が正しく入力されていません。", vbInformation, strMsg_kaku
'            Me.txt_連絡先電話番号.SetFocus
'            Exit Function
'        End If
'    End If
' DEL 2006/03/03 Tajima End 受付対応に伴い顧客入力は行わないのでチェックしない
    
'ADD 2006/05/22 Tajima Start
    If pintAGRE >= 4 Then
      ' メンテ以降は契約入力での鍵変更は推奨しない
      If Nz(pOriginalKeyType, "") <> Nz(Me.cmb_鍵区分.VALUE, "") Or _
         Nz(pOriginalKeyNo, "") <> Nz(Me.txt_解除番号, "") Or _
         Nz(pOriginalKeyNum, 0) <> Nz(Me.txt_貸出本数, 0) Then
         
         MsgBox "鍵関係の変更はメンテ以前でなければ、この契約にしか反映されません。" & _
                vbCrLf & "次契約には変更した鍵情報を引き継ぐことはできません。", vbInformation, strMsg_kaku
      End If
    End If
'ADD 2006/05/22 Tajima End
    
    ' 請求書発行区分
    If Nz(Me.cmb_請求書発行区分, "") = "" Then                                  'INSERT START 2010/04/10 K.ISHIZAKA
        MsgBox "請求書発行区分を入力して下さい。", vbInformation, strMsg_kaku
        Me.cmb_請求書発行区分.SetFocus
        Exit Function
    End If                                                                      'INSERT END   2010/04/10 K.ISHIZAKA
    
    CHK_RTN_For_UpDate = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :新規登録時のデータチェック
'        MODULE_ID        :CHK_CARG_FILE2
'        CREATE_DATE      :
'        処理内容         ：同一、ヤード、コンテナ、顧客、契約開始日（計上年月）が存在した場合は、
'                          エラーとする
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CHK_CARG_FILE2() As Boolean
    
    Dim strSQL As String
    Dim objRs  As Recordset  ' 2006/01/05 ADD T.SUZUKI
'    Dim vCode  As Variant   ' 2005/08/03 DEL T.SUZUKI
'    Dim vRet   As Variant   ' 2005/08/03 DEL T.SUZUKI
'    Dim i      As Integer   ' 2005/08/03 DEL T.SUZUKI

    On Error GoTo Err_rtn2

    CHK_CARG_FILE2 = False

    doCmd.Hourglass True
    doCmd.SetWarnings False

    ' 存在チェック
    strSQL = " SELECT * " & Chr(13)
    strSQL = strSQL & "  FROM CARG_FILE " & Chr(13)                                             ' 2006/01/05 ADD T.SUZUKI
'    strSQL = strSQL & "  FROM dbo_CARG_FILE " & Chr(13)                                        ' 2006/01/05 DEL T.SUZUKI
    strSQL = strSQL & " WHERE CARG_YCODE  = " & Val(Me.cmb_ヤードコード) & Chr(13)
    strSQL = strSQL & "   AND CARG_NO     = " & Val(Me.cmb_コンテナ番号) & Chr(13)
    strSQL = strSQL & "   AND CARG_FSDATE = " & Format(Me.txt_初回契約日, P_FORMAT_YMD) & Chr(13)                   ' 2006/01/05 ADD T.SUZUKI
'    strSQL = strSQL & "   AND CARG_FSDATE = CDATE(#" & Format(Me.txt_初回契約日, "yyyy/mm/dd") & "# ) " & Chr(13)  ' 2006/01/05 DEL T.SUZUKI
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)                      ' 2006/01/05 ADD T.SUZUKI
'    Set pobjRs = pobjDB.OpenRecordset(strSQL, dbOpenSnapshot)                                                      ' 2006/01/05 DEL T.SUZUKI

    If Not (objRs.BOF And objRs.EOF) Then                                                       ' 2006/01/05 ADD T.SUZUKI
'    If Not (pobjRs.BOF And pobjRs.EOF) Then                                                    ' 2006/01/05 DEL T.SUZUKI
        MsgBox "このコンテナは既に同じ初回契約日で登録されています｡" + vbCrLf + _
               "このコンテナに対しての契約は翌日にならないと登録できません｡", _
               vbInformation, strMsg_err
        GoTo Exit_rtn
    End If

    CHK_CARG_FILE2 = True

Exit_rtn:
    doCmd.Hourglass False
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing                               ' 2006/01/05 ADD T.SUZUKI
'    Set pobjRs = Nothing                                                                       ' 2006/01/05 DEL T.SUZUKI
    Exit Function

Err_rtn2:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo Exit_rtn
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :WRI_RTN_For_CARG_FILE
'        CREATE_DATE      :
'
'        引数              ０、正更新
'                          １、逆更新
'                           strUserName ：ユーザー名
'                           strKIYAKU_NO：受注契約番号
'                           lngREQU_NO  ：請求NO
'
'        コンテナ予約ＷＫからコンテナマスタに予約をセットする。
'        解約チェックがされている時は、コンテナマスタからクリアする。
'        既に予約されている場合は、変更のあったコンテナのみ変更する。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2006/02/03 REP T.SUZUKI Start
Private Function WRI_RTN_For_CARG_FILE(ByVal iMode As Integer, _
                                       ByVal strUserName As String, _
                                       Optional strKIYAKU_NO As String, _
                                       Optional lngREQU_NO As Long) As Boolean
'Private Function WRI_RTN_For_CARG_FILE(ByVal iMode As Integer, _
                                       Optional strKIYAKU_NO As String, _
                                       Optional lngREQU_NO As Long) As Boolean
' 2006/02/03 REP T.SUZUKI End

    Dim strSQL        As String
    Dim strErrMsg     As String
    Dim objRs         As Recordset
    Dim objRsKONT     As Recordset

' 2006/01/05 ADD T.SUZUKI Start
    Dim strMenteMemo  As String   ' メンテメモ
' 2006/01/05 ADD T.SUZUKI End

    Dim blnKaiYotei   As Boolean    'INSERT 2014/04/16 MIYAMOTO
    Dim strNameFrom   As String     'INSERT 2014/06/24 MIYAMOTO

    On Error GoTo err_rtn

    WRI_RTN_For_CARG_FILE = False
    blnKaiYotei = False             'INSERT 2014/04/16 MIYAMOTO

    doCmd.SetWarnings False

    Set pWRI_ObjWS = DBEngine.Workspaces(0)
    Set pWRI_objDb = CurrentDb

    ' トランザクション開始
    pWRI_ObjWS.BeginTrans

    Select Case iMode
        Case 0          ' *** 追加 ***
            ' コンテナ契約ファイル作成
            ' コンテナ契約ファイルオープン
            strSQL = "SELECT * FROM CARG_FILE "
            Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

            With objRs
                .AddNew

                .Fields("CARG_YCODE") = Format(Me.cmb_ヤードコード, "000000")                      ' ヤードコード
                .Fields("CARG_NO") = Format(Me.cmb_コンテナ番号, "000000")                         ' コンテナ番号

'2006/03/03 DEL Tajima Start  受付対応により新規登録は無いので
'                ' 顧客コード
'                If (Me.opg_処理選択.Value = 1 And Me.opg_顧客情報取得選択.Value = 1) Then
'                    ' *** 新規契約 かつ 新規顧客 ***
'                    .Fields("CARG_UCODE") = plngUSER_NO
'                Else
'                    ' *** 新規契約 かつ 既存顧客 ***
'                    .Fields("CARG_UCODE") = Format(Me.cmb_顧客コード, "000000")
'                End If
'2006/03/03 DEL Tajima End

                ' 契約状態
'2006/01/05 ADD T.SUZUKI Start
                ' 予約受付トランからの読み込み時
                If (Nz(SAKI_YOUK_TRAN.strMOTO_ACPTNO, "") <> "" And SAKI_YOUK_TRAN.strMOVEKBN <> P_MOVEKBN_00) Then
                    Select Case MOTO_CARG_FILE.strAGRE
                        Case P_AGRE_01
'                            ' *** 契約 ***
                            .Fields("CARG_AGRE") = P_AGRE_01    ' 1(契約)
                        Case P_AGRE_02, P_AGRE_03
                            ' *** 解約予定 OR 解約延長 ***
                            .Fields("CARG_AGRE") = P_AGRE_02    ' 2(解約予定)
                            blnKaiYotei = True                                          'INSERT 2014/04/16 MIYAMOTO
                    End Select

                Else
                    If Nz(Me.txt_解約日付, "") = "" Then
                        .Fields("CARG_AGRE") = P_AGRE_01        ' 1(契約)
                    Else
                        .Fields("CARG_AGRE") = P_AGRE_02        ' 2(解約予定)
                        blnKaiYotei = True                                              'INSERT 2014/04/16 MIYAMOTO
                    End If
                End If
'2006/01/05 ADD T.SUZUKI End
'2006/01/05 DEL T.SUZUKI Start
'                If Nz(Me.txt_解約日付, "") = "" Then
'                    .Fields("CARG_AGRE") = P_KEIYAKU        ' 1(契約)
'                Else
'                    .Fields("CARG_AGRE") = P_KAIYAKU_YOTEI  ' 2(解約予定)
'                End If
'2006/01/05 DEL T.SUZUKI End

                .Fields("CARG_FSDATE") = Format(Me.txt_初回契約日, P_FORMAT_YMD)                   ' 初回契約日
                .Fields("CARG_STDATE") = Format(Me.txt_契約開始日, P_FORMAT_YMD)                   ' 契約開始日
                .Fields("CARG_EDDATE") = Format(Me.txt_契約満了日, P_FORMAT_YMD)                   ' 契約満了日

                If Nz(Me.txt_解約日付, "") = "" Then
                    .Fields("CARG_CYDATE") = Null                                                  ' 解約予定日
                    .Fields("CARG_KYDATE") = Null                                                  ' 解約日付
                Else
                    .Fields("CARG_CYDATE") = Format(Me.txt_解約日付, P_FORMAT_YMD)                 ' 解約予定日
                    .Fields("CARG_KYDATE") = Format(Me.txt_解約日付, P_FORMAT_YMD)                 ' 解約日付
                End If

                .Fields("CARG_CONTNA") = P_ZERO                                                    ' コンテナ状態
                .Fields("CARG_DOCU1") = Me.chk_書類1                                               ' 新契約書切替区分
                .Fields("CARG_TUBOSU") = P_ZERO                                                    ' 坪数
                .Fields("CARG_TTANKA") = P_ZERO                                                    ' 坪単価
                .Fields("CARG_RENTKG") = Val(Nz(Me.txt_月額料金, P_ZERO))                          ' 月額料金
                .Fields("CARG_SYOZEI") = Val(Nz(Me.txt_消費税, P_ZERO))                            ' 消費税
'2006/03/03 DEL Tajima Start
'                .Fields("CARG_DOCU2") = Me.chk_書類2                                               ' 代引き区分
'                .Fields("CARG_FRSTKG") = Val(Nz(Me.txt_初回金額, P_ZERO))                          ' 初回金額
'                .Fields("CARG_FSYOZEI") = Val(Nz(Me.txt_初回消費税, P_ZERO))                       ' 初回金額消費税
'                ' 代引伝票送付日
'                If Nz(Me.txt_戻り日, "") = "" Then
'                    .Fields("CARG_MDATE") = Null
'                Else
'                    .Fields("CARG_MDATE") = Format(Me.txt_戻り日, P_FORMAT_YMD)
'                End If
'                .Fields("CARG_DAHIB") = Me.txt_番号                                                 ' 代引き伝票番号
'                .Fields("CARG_KAGIA") = Val(Nz(Me.txt_初回鍵代, P_ZERO))                            ' 鍵代
'                .Fields("CARG_SECUKG") = Val(Nz(Me.txt_預り保証金, P_ZERO))                        ' 預かり保証金
'2006/03/03 DEL Tajima End
'2006/05/07 ADD tajima Start
                .Fields("CARG_SPAREKEY_NOTE") = Nz(Trim(Me.txt_スペアキー備考), "")
                If .Fields("CARG_SPAREKEY_NOTE") <> "" Then
                    .Fields("CARG_SPAREKEY") = 0  'ｽﾍﾟｱｷｰ備考に何か入力があればｽﾍﾟｱｷｰ無し
                Else
                    .Fields("CARG_SPAREKEY") = 1 'ｽﾍﾟｱｷｰ備考に何も入力なければｽﾍﾟｱｷｰ有り
                End If
'2006/05/07 ADD tajima End
                .Fields("CARG_BIKO") = Nz(Me.txt_契約備考, "")                                     ' 備考
                .Fields("CARG_UPDATE") = Format(Now, P_FORMAT_YMD)                                 ' 更新日
                .Fields("CARG_HDATE") = Null                                                       ' 保険加入日

                .Fields("CARG_ACPTNO") = strKIYAKU_NO                                               ' 受注契約番号

                ' 画面の受注契約番号セット --------
                Me.txt_受注契約番号 = strKIYAKU_NO
                ' ---------------------------------

                .Fields("CARG_HOSYICD") = Format(Me.cmb_保証区分, "0")                              ' 保証区分
                .Fields("CARG_HOSYO_CD") = Format(Me.cmb_保証会社コード, "0")                       ' 保証会社コード　'2009/04/01 INS KIHARA
                .Fields("CARG_KAGIICD") = Format(Me.cmb_鍵区分, "0")                                ' 鍵区分
                .Fields("CARG_HOKEB") = Null                                                        ' 保険番号
                .Fields("CARG_KAIJB") = Me.txt_解除番号                                             ' 鍵解除番号
                .Fields("CARG_HOKAI") = 9                                                           ' 保証会社区分

                ' 保証承認日
                If Nz(Me.txt_承認日, "") = "" Then
                    .Fields("CARG_HOSYD") = Null
                Else
                    .Fields("CARG_HOSYD") = Me.txt_承認日
                End If

                .Fields("CARG_HOSYB") = Me.txt_承認番号                                             ' 保証承認番号
                .Fields("CARG_KEITI") = P_AGRE_99                                                   ' 契約形態

                ' [ヤード利用状況]画面または[予約一覧]画面からの遷移の場合、
                ' 予約受付トラン.受付番号をセット
                If (pintOpenMode = 1 Or pintOpenMode = 9) Then                                      ' 2006/01/05 ADD T.SUZUKI
'                ' ヤード利用状況画面からの遷移の場合、予約受付トラン.受付番号をセット
'                If pintOpenMode = 1 Then                                                           ' 2006/01/05 DEL T.SUZUKI
                    ' 受付番号
                    .Fields("CARG_UKNO") = pstrUKNO
                End If

                .Fields("CARG_UKTANTO") = Me.cmb_受付担当者コード                                   ' 受付担当者コード

                If Nz(Me.txt_解約日付, "") = "" Then
                    .Fields("CARG_KAIUKEDATE") = Null                                               ' 解約受付日
                    .Fields("CARG_KAITANTO") = Null                                                 ' 解約受付担当者コード
                    .Fields("CARG_KAICD") = 1                                                       ' 解約区分
                    .Fields("CARG_KAIRIYUCD") = Null                                                ' 解約理由区分コード
                Else
                    .Fields("CARG_KAIUKEDATE") = Format(Me.txt_解約受付日, P_FORMAT_YMD)            ' 解約受付日
                    .Fields("CARG_KAITANTO") = Me.cmb_解約受付担当者コード                          ' 解約受付担当者コード
                    .Fields("CARG_KAICD") = Nz(Me.cmb_解約区分, 1)                                  ' 解約区分
                    .Fields("CARG_KAIRIYUCD") = Nz(Me.cmb_解約理由, Null)                           ' 解約理由区分コード
                End If
                .Fields("CARG_SEIKI") = cmb_請求書発行区分.VALUE                'INSERT 2010/04/10 K.ISHIZAKA

                .Fields("CARG_INSED") = Format$(DATE, "yyyymmdd")                                   ' 作成日付
                .Fields("CARG_INSEJ") = Format$(time, "hhmmss")                                     ' 作成時刻
                .Fields("CARG_INSPB") = Me.Form.NAME                                                ' 作成プログラムID
                .Fields("CARG_INSUB") = strUserName                                                 ' 作成ユーザーID   ' 2006/02/03 ADD T.SUZUKI
'                .Fields("CARG_INSUB") = MSZZ000.LsGetUserName                                       ' 作成ユーザーID  ' 2006/02/03 DEL T.SUZUKI

                .UPDATE
            End With
            objRs.Close
            Set objRs = Nothing

            ' 予約受付トランの更新
            If (pintOpenMode = 1 Or pintOpenMode = 9) Then                                          ' 2006/01/05 ADD T.SUZUKI
'            If pintOpenMode = 1 Then                                                               ' 2006/01/05 DEL T.SUZUKI
                ' 予約受付トランよりデータを取り込んだ場合、予約受付トランの予約受付状態区分を"20"(契約)に更新する。
                If fncUPD_YOUK_TRAN(P_YUKBN_20, strUserName, strErrMsg) = False Then GoTo WRI_Err   ' 2006/02/03 ADD T.SUZUKI
'                If fncUPD_YOUK_TRAN(P_YUKBN_20, strErrMsg) = False Then GoTo WRI_Err               ' 2006/02/03 DEL T.SUZUKI

' 2006/02/02 ADD T.SUZUKI Start
                ' 予約受付トランよりデータを取り込んだ場合、予約ご紹介トランのご紹介区分を"3"(契約)に更新する。
                If fncUPD_INTR_TRAN(P_INTROKBN_02, P_INTROKBN_03, CLng(Me.cmb_ヤードコード), _
                                    CLng(Me.cmb_コンテナ番号), strUserName, strErrMsg) = False Then
                    GoTo WRI_Err
                End If
' 2006/02/02 ADD T.SUZUKI End
            End If

' DEL 2006/03/03 Tajima End 新規契約登録はあり得ないので呼ばない
'           ' 請求ファイルの作成
'            If WRI_RTN_For_REQU_FILE(lngREQU_NO, strKIYAKU_NO, strErrMsg) = False Then GoTo WRI_Err
' DEL 2006/03/03 Tajima Start


' 2006/01/05 ADD T.SUZUKI Start
            ' 予約受付トランからの読み込み時
            If (Nz(SAKI_YOUK_TRAN.strMOTO_ACPTNO, "") <> "" And SAKI_YOUK_TRAN.strMOVEKBN <> P_MOVEKBN_00) Then
                ' メンテメモの文言作成
                strMenteMemo = ""
                Call subCreateMemo(1, SAKI_YOUK_TRAN.strMOVEKBN, strMenteMemo)

                ' 移動元コンテナ契約トラン更新
                ' ①dao.DataBase
                ' ②受注契約番号
                ' ③移動区分
                ' ④契約開始日
                ' ⑤受付日
                ' ⑥受付担当者コード
                ' ⑦更新プログラムID
                ' ⑧更新ユーザーID
                ' ⑨エラーメッセージ格納
                If UpdateMoveOldCARG(objKONT_DB, _
                                     Nz(SAKI_YOUK_TRAN.strMOTO_ACPTNO, ""), _
                                     SAKI_YOUK_TRAN.strMOVEKBN, _
                                     CDate(Me.txt_契約開始日), _
                                     SAKI_YOUK_TRAN.strUkDate, _
                                     SAKI_YOUK_TRAN.strUKTANTO, _
                                     strMenteMemo, _
                                     Me.Form.NAME, _
                                     strUserName, _
                                     strErrMsg) <> 0 Then

                    If Nz(strErrMsg, "") <> "" Then
                        MsgBox strErrMsg, vbInformation, strMsg_err
                        GoTo WRI_Err
                    End If
                End If
            End If
' 2006/01/05 ADD T.SUZUKI End


        Case 2  ' *** 修正 ***
            ' コンテナ契約対象ファイルデータの抽出
            strSQL = "SELECT * " & Chr(13)
            strSQL = strSQL & "  FROM CARG_FILE " & Chr(13)
            strSQL = strSQL & " WHERE CARG_ACPTNO = '" & pstrJUCHUKEIYAKU_NO & "'"  ' 受注契約番号
            Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

            If objRs.EOF = True Then
                objRs.Close
                Set objRs = Nothing
                GoTo Exit_rtn
            End If

            With objRs
                .Edit

                .Fields("CARG_UCODE") = Format(Me.cmb_顧客コード, "000000")       ' 顧客コード

                ' 契約状態
                If Nz(Me.txt_解約受付日, "") = "" Then
                    .Fields("CARG_AGRE") = P_AGRE_01                              ' 1(契約)
'                    .Fields("CARG_AGRE") = P_KEIYAKU                              ' 1(契約)
                Else
                    ' 解約日付が入力されているとき
                    ' 契約状態が0(未設定)または1(契約)のときには『契約状態=2(解約予定)』に設定する
                    If (pintAGRE = 1) Then
                        .Fields("CARG_AGRE") = P_AGRE_02                          ' 2(解約予定)
'                        .Fields("CARG_AGRE") = P_KAIYAKU_YOTEI                    ' 2(解約予定)
                        blnKaiYotei = True                                              'INSERT 2014/04/16 MIYAMOTO
                    End If
                End If

                .Fields("CARG_FSDATE") = Format(Me.txt_初回契約日, P_FORMAT_YMD)          ' 初回契約日
                .Fields("CARG_STDATE") = Format(Me.txt_契約開始日, P_FORMAT_YMD)          ' 契約開始日
                .Fields("CARG_EDDATE") = Format(Me.txt_契約満了日, P_FORMAT_YMD)          ' 契約満了日

                If Nz(Me.txt_解約日付, "") = "" Then
                    .Fields("CARG_CYDATE") = Null                                         ' 解約予定日
                    .Fields("CARG_KYDATE") = Null                                         ' 解約日付
                    .Fields("CARG_KAIUKEDATE") = Null                                     ' 解約受付日
                    .Fields("CARG_KAITANTO") = Null                                       ' 解約受付担当者コード
                    .Fields("CARG_KAICD") = 1                                             ' 解約区分
                    .Fields("CARG_KAIRIYUCD") = Null                                      ' 解約理由
                Else
                    If (pintAGRE = 1 Or pintAGRE = 2) Then
                        ' 解約日付が入力されているとき
                        ' 契約状態が0(未設定)または1(契約)のときには解約予定日に解約日付を設定する
                        .Fields("CARG_CYDATE") = Format(Me.txt_解約日付, P_FORMAT_YMD)    ' 解約予定日
                    Else
                        ' 解約予定日の更新は行わない
                    End If

                    .Fields("CARG_KYDATE") = Format(Me.txt_解約日付, P_FORMAT_YMD)        ' 解約日付
                    .Fields("CARG_KAIUKEDATE") = Format(Me.txt_解約受付日, P_FORMAT_YMD)  ' 解約受付日
                    .Fields("CARG_KAITANTO") = Me.cmb_解約受付担当者コード                ' 解約受付担当者コード
                    .Fields("CARG_KAICD") = Nz(Me.cmb_解約区分, 1)                        ' 解約区分
                    .Fields("CARG_KAIRIYUCD") = Nz(Me.cmb_解約理由, Null)                 ' 解約理由
                End If

                .Fields("CARG_DOCU1") = Me.chk_書類1                                      ' 新契約書切替区分
                .Fields("CARG_TUBOSU") = 0                                                ' 坪数
                .Fields("CARG_TTANKA") = 0                                                ' 坪単価
                .Fields("CARG_RENTKG") = Val(Nz(Me.txt_月額料金, P_ZERO))                 ' 月額料金
                .Fields("CARG_SYOZEI") = Val(Nz(Me.txt_消費税, P_ZERO))                   ' 消費税
'DEL 2006/03/03 Tajima Start
'                .Fields("CARG_DOCU2") = Me.chk_書類2                                      ' 代引き区分
'                .Fields("CARG_FRSTKG") = Val(Nz(Me.txt_初回金額, P_ZERO))                 ' 初回金額
'                .Fields("CARG_FSYOZEI") = Val(Nz(Me.txt_初回消費税, P_ZERO))              ' 初回金額消費税
'                .Fields("CARG_KAGIA") = Val(Nz(Me.txt_初回鍵代, P_ZERO))                  ' 鍵代
'                ' 代引き伝票発送日
'                If Nz(Me.txt_戻り日, "") = "" Then
'                    .Fields("CARG_MDATE") = Null
'                Else
'                    .Fields("CARG_MDATE") = Format(Me.txt_戻り日, P_FORMAT_YMD)
'                End If
'                .Fields("CARG_DAHIB") = Me.txt_番号                                       ' 代引伝票番号
'                .Fields("CARG_SECUKG") = Val(Nz(Me.txt_預り保証金, P_ZERO))               ' 預り金
'DEL 2006/03/03 Tajima End
'2006/05/07 ADD tajima Start
                .Fields("CARG_SPAREKEY_NOTE") = Nz(Trim(Me.txt_スペアキー備考), "")
                If .Fields("CARG_SPAREKEY_NOTE") <> "" Then
                    .Fields("CARG_SPAREKEY") = 0  'ｽﾍﾟｱｷｰ備考に何か入力があればｽﾍﾟｱｷｰ無し
                Else
                    .Fields("CARG_SPAREKEY") = 1 'ｽﾍﾟｱｷｰ備考に何も入力なければｽﾍﾟｱｷｰ有り
                End If
'2006/05/07 ADD tajima End
                .Fields("CARG_BIKO") = Nz(Me.txt_契約備考, "")                            ' 備考
                .Fields("CARG_UPDATE") = Format(Now, P_FORMAT_YMD)                        ' 更新日
                .Fields("CARG_HDATE") = Null                                              ' 保険加入日

                .Fields("CARG_HOSYICD") = Format(Me.cmb_保証区分, "0")                    ' 保証区分
                .Fields("CARG_KAGIICD") = Format(Me.cmb_鍵区分, "0")                      ' 鍵区分
                .Fields("CARG_HOKEB") = Null                                              ' 保険番号
                .Fields("CARG_KAIJB") = Me.txt_解除番号                                   ' 鍵解除番号
                .Fields("CARG_KEY_LENTNUM") = Me.txt_貸出本数                             ' 鍵貸出本数 2006/05/22 add tajima

                ' 保証承認日
                If Nz(Me.txt_承認日, "") = "" Then
                    .Fields("CARG_HOSYD") = Null
                    strSQL = strSQL & "CARG_HOSYD = NULL, "
                Else
                    .Fields("CARG_HOSYD") = Format(Me.txt_承認日, P_FORMAT_YMD)
                End If

                .Fields("CARG_HOSYB") = Me.txt_承認番号                                   ' 保証承認番号
                .Fields("CARG_UKTANTO") = Me.cmb_受付担当者コード                         ' 受付担当者コード
                .Fields("CARG_SEIKI") = cmb_請求書発行区分.VALUE                'INSERT 2010/04/10 K.ISHIZAKA
                
'↓ INSERT 2014/06/24 MIYAMOTO
                '地方ヤード以外で解約予定になった場合
                If Me.txt_地方ヤード <> -1 And blnKaiYotei = True Then
                    '選択されている鍵区分が変換対象の場合、鍵変更理由コードに3をセットする
'↓ UPDATE 2016/05/19 MIYAMOTO
'                    If Nz(Me.cmb_鍵区分.Column(3)) <> -1 Then
                    If MSZZ070_KaiKey(txt_地方ヤード, Nz(Me.cmb_鍵区分.Column(3))) <> 0 Then
'↑ UPDATE 2016/05/19 MIYAMOTO
                        .Fields("CARG_KEY_CHGRIYUCD") = 3                                 ' 鍵変更理由コード
                    End If
                End If
'↑ INSERT 2014/06/24 MIYAMOTO
                
                .Fields("CARG_UPDAD") = Format$(DATE, "yyyymmdd")                         ' 更新日付
                .Fields("CARG_UPDAJ") = Format$(time, "hhmmss")                           ' 更新時刻
                .Fields("CARG_UPDPB") = Me.Form.NAME                                      ' 更新プログラムID
                .Fields("CARG_UPDUB") = strUserName                                       ' 更新ユーザーID   ' 2006/02/03 REP T.SUZUKI

                '2006/05/10 鍵情報の保存 Add Start Tajima
                .Fields("CARG_KEY_LENTNUM") = Nz(Me.txt_貸出本数, 0)
                '2006/05/10 鍵情報の保存 Add End Tajima

                .UPDATE
            End With
            objRs.Close
            Set objRs = Nothing

'DEL 2006/03/03 Tajima Strat 受付対応によって金額情報の変更が出来なくなるので削除
'            If pblnMoneyChgFlg = True Then
'                ' 請求ファイルの更新
'                If fncUPD_REQU_FILE(pstrJUCHUKEIYAKU_NO, Format(Me.txt_契約開始日, "YYYY/MM"), _
'                                    strErrMsg) = False Then GoTo WRI_Err
'                pblnMoneyChgFlg = False
'            End If
'DEL 2006/03/03 Tajima End

        Case 1          ' *** 削除 ***
            ' コンテナ契約ファイルのデータをコンテナ契約削除ファイルに待避する
            strSQL = "INSERT INTO CARG_DELE "
            strSQL = strSQL & " ( "
            strSQL = strSQL & " CARD_YCODE, "         ' ヤードコード
            strSQL = strSQL & " CARD_NO, "            ' コンテナ番号
            strSQL = strSQL & " CARD_UCODE, "         ' 顧客コード
            strSQL = strSQL & " CARD_AGRE, "          ' 契約状態
            strSQL = strSQL & " CARD_FSDATE, "        ' 初回契約日
            strSQL = strSQL & " CARD_STDATE, "        ' 契約開始日
            strSQL = strSQL & " CARD_EDDATE, "        ' 契約満了日
            strSQL = strSQL & " CARD_CYDATE, "        ' 解約予定日
            strSQL = strSQL & " CARD_KYDATE, "        ' 解約日付
            strSQL = strSQL & " CARD_CONTNA, "        ' コンテナ状態
            strSQL = strSQL & " CARD_DOCU1, "         ' 新契約書切替区分
            strSQL = strSQL & " CARD_DOCU2, "         ' 代引き区分
            strSQL = strSQL & " CARD_TUBOSU, "        ' 坪数コンテナ坪数
            strSQL = strSQL & " CARD_TTANKA, "        ' コンテナ坪単価
            strSQL = strSQL & " CARD_RENTKG, "        ' 月額料金
            strSQL = strSQL & " CARD_SYOZEI, "        ' 消費税
            strSQL = strSQL & " CARD_FRSTKG, "        ' 初回金額
            strSQL = strSQL & " CARD_FSYOZEI, "       ' 初回金額消費税
            strSQL = strSQL & " CARD_FRST_BILL , "    ' 初回請求金額 2006/03/18 Add Tajima
            strSQL = strSQL & " CARD_PREMTH_SUM, "    ' 前払月数     2006/03/18 Add Tajima
            strSQL = strSQL & " CARD_SECUKG, "        ' 預り金
            strSQL = strSQL & " CARD_BIKO, "          ' 備考
            strSQL = strSQL & " CARD_UPDATE, "        ' 更新日
            strSQL = strSQL & " CARD_HDATE, "         ' 保険加入日
            strSQL = strSQL & " CARD_MDATE, "         ' 代引伝票送付日
            strSQL = strSQL & " CARD_ACPTNO, "        ' 受注契約番号
            strSQL = strSQL & " CARD_HOSYICD, "       ' 保証区分
            strSQL = strSQL & " CARD_HOSYO_CD, "      ' 保証会社                  'INSERT 2010/04/10 K.ISHIZAKA
            strSQL = strSQL & " CARD_KAGIICD, "       ' 鍵区分
            strSQL = strSQL & " CARD_DAHIB, "         ' 代引き伝票番号
            strSQL = strSQL & " CARD_HOKEB, "         ' 保険番号
            strSQL = strSQL & " CARD_KAIJB, "         ' 鍵解除番号
            strSQL = strSQL & " CARD_HOKAI, "         ' 保証会社区分
            strSQL = strSQL & " CARD_HOSYD, "         ' 保証承認日
            strSQL = strSQL & " CARD_HOSYB, "         ' 保証承認番号
            strSQL = strSQL & " CARD_ADMEI, "         ' 広告媒体
            strSQL = strSQL & " CARD_KEITI, "         ' 契約携帯
            strSQL = strSQL & " CARD_KAGIA, "         ' 鍵代
            strSQL = strSQL & " CARD_UKNO, "          ' 予約受付番号
            strSQL = strSQL & " CARD_UKTANTO, "       ' 契約受付担当者コード
            strSQL = strSQL & " CARD_KAIUKEDATE, "    ' 解約受付日
            strSQL = strSQL & " CARD_KAITANTO, "      ' 解約担当者コード
            strSQL = strSQL & " CARD_KAICD, "         ' 解約区分
            strSQL = strSQL & " CARD_KAIRIYUCD, "     ' 解約理由区分コード
            strSQL = strSQL & " CARD_CMP_EXDATE, "    ' キャンペーン適用満了日 2006/03/18 Add Tajima
            strSQL = strSQL & " CARD_KEY_RETDATE, "   ' 鍵返却日
            strSQL = strSQL & " CARD_SPAREKEY, "      ' スペアキー
            strSQL = strSQL & " CARD_MNT_REQDATE, "   ' メンテ依頼日
            strSQL = strSQL & " CARD_MNT_CMPDATE, "   ' メンテ完了日
            strSQL = strSQL & " CARD_NIRCD, "         ' 荷入区分
            strSQL = strSQL & " CARD_SYOCD, "         ' 処分区分
            strSQL = strSQL & " CARD_HOSCD, "         ' 補修区分
            strSQL = strSQL & " CARD_MNT_TANTO, "     ' メンテ担当者コード
            strSQL = strSQL & " CARD_MNT_MEMO, "      ' メンテメモ
            strSQL = strSQL & " CARD_LIQ_REQDATE, "   ' 清算依頼日
            strSQL = strSQL & " CARD_LIQ_NOTE, "      ' 清算備考
            strSQL = strSQL & " CARD_LIQ_FLG, "       ' 清算済フラグ
            '2006/05/24 Add tajima Start
            strSQL = strSQL & " CARD_SPAREKEY_NOTE, " ' スペアキー備考
            strSQL = strSQL & " CARD_KEY_LENTNUM,"    ' 鍵貸出本数
            strSQL = strSQL & " CARD_KEY_RETNUM,"     ' 鍵返却本数
            strSQL = strSQL & " CARD_KEY_CHGDATE,"    ' 鍵変更日
            strSQL = strSQL & " CARD_KEY_CHGTANTO,"   ' 鍵変更担当
            strSQL = strSQL & " CARD_KEY_CHGRIYUCD,"  ' 鍵変更理由区分コード
            strSQL = strSQL & " CARD_KEY_CHGNOTE,"    ' 鍵変更備考
            '2006/05/24 Add tajima End
            strSQL = strSQL & " CARD_SEIKI, "         ' 請求書発行区分          'INSERT 2010/04/10 K.ISHIZAKA
            strSQL = strSQL & " CARD_INSED, "         ' 作成日付
            strSQL = strSQL & " CARD_INSEJ, "         ' 作成時刻
            strSQL = strSQL & " CARD_INSPB, "         ' 作成プログラムID
            strSQL = strSQL & " CARD_INSUB, "         ' 作成ユーザーID
            strSQL = strSQL & " CARD_UPDAD, "         ' 更新日付
            strSQL = strSQL & " CARD_UPDAJ, "         ' 更新時刻
            strSQL = strSQL & " CARD_UPDPB, "         ' 更新プログラムID
            strSQL = strSQL & " CARD_UPDUB, "         ' 更新ユーザーID
            strSQL = strSQL & " CARD_DELED, "         ' 削除日付
            strSQL = strSQL & " CARD_DELEJ, "         ' 削除時刻
            strSQL = strSQL & " CARD_DELPB, "         ' 削除プログラムID
            strSQL = strSQL & " CARD_DELUB "          ' 削除ユーザーID
            strSQL = strSQL & " ) "
            strSQL = strSQL & " SELECT CARG_YCODE, "
            strSQL = strSQL & "        CARG_NO, "
            strSQL = strSQL & "        CARG_UCODE, "
            strSQL = strSQL & "        CARG_AGRE, "
            strSQL = strSQL & "        CARG_FSDATE, "
            strSQL = strSQL & "        CARG_STDATE, "
            strSQL = strSQL & "        CARG_EDDATE, "
            strSQL = strSQL & "        CARG_CYDATE, "
            strSQL = strSQL & "        CARG_KYDATE, "
            strSQL = strSQL & "        CARG_CONTNA, "
            strSQL = strSQL & "        CARG_DOCU1, "
            strSQL = strSQL & "        CARG_DOCU2, "
            strSQL = strSQL & "        CARG_TUBOSU, "
            strSQL = strSQL & "        CARG_TTANKA, "
            strSQL = strSQL & "        CARG_RENTKG, "
            strSQL = strSQL & "        CARG_SYOZEI, "
            strSQL = strSQL & "        CARG_FRSTKG, "
            strSQL = strSQL & "        CARG_FSYOZEI, "
            strSQL = strSQL & "        CARG_FRST_BILL, "  '2006/03/18 Add Tajima
            strSQL = strSQL & "        CARG_PREMTH_SUM, " '2006/03/18 Add Tajima
            strSQL = strSQL & "        CARG_SECUKG, "
            strSQL = strSQL & "        CARG_BIKO, "
            strSQL = strSQL & "        CARG_UPDATE, "
            strSQL = strSQL & "        CARG_HDATE, "
            strSQL = strSQL & "        CARG_MDATE, "
            strSQL = strSQL & "        CARG_ACPTNO, "
            strSQL = strSQL & "        CARG_HOSYICD, "
            strSQL = strSQL & "        CARG_HOSYO_CD, "                         'INSERT 2010/04/10 K.ISHIZAKA
            strSQL = strSQL & "        CARG_KAGIICD, "
            strSQL = strSQL & "        CARG_DAHIB, "
            strSQL = strSQL & "        CARG_HOKEB, "
            strSQL = strSQL & "        CARG_KAIJB, "
            strSQL = strSQL & "        CARG_HOKAI, "
            strSQL = strSQL & "        CARG_HOSYD, "
            strSQL = strSQL & "        CARG_HOSYB, "
            strSQL = strSQL & "        CARG_ADMEI, "
            strSQL = strSQL & "        CARG_KEITI, "
            strSQL = strSQL & "        CARG_KAGIA, "
            strSQL = strSQL & "        CARG_UKNO, "
            strSQL = strSQL & "        CARG_UKTANTO, "
            strSQL = strSQL & "        CARG_KAIUKEDATE, "
            strSQL = strSQL & "        CARG_KAITANTO, "
            strSQL = strSQL & "        CARG_KAICD, "
            strSQL = strSQL & "        CARG_KAIRIYUCD, "
            strSQL = strSQL & "        CARG_CMP_EXDATE, " '2006/03/18 Add Tajima
            strSQL = strSQL & "        CARG_KEY_RETDATE, "
            strSQL = strSQL & "        CARG_SPAREKEY, "
            strSQL = strSQL & "        CARG_MNT_REQDATE, "
            strSQL = strSQL & "        CARG_MNT_CMPDATE, "
            strSQL = strSQL & "        CARG_NIRCD, "
            strSQL = strSQL & "        CARG_SYOCD, "
            strSQL = strSQL & "        CARG_HOSCD, "
            strSQL = strSQL & "        CARG_MNT_TANTO, "
            strSQL = strSQL & "        CARG_MNT_MEMO, "
            strSQL = strSQL & "        CARG_LIQ_REQDATE, "
            strSQL = strSQL & "        CARG_LIQ_NOTE, "
            strSQL = strSQL & "        CARG_LIQ_FLG, "
            '2006/05/24 Add tajima Start
            strSQL = strSQL & "        CARG_SPAREKEY_NOTE,"
            strSQL = strSQL & "        CARG_KEY_LENTNUM,"
            strSQL = strSQL & "        CARG_KEY_RETNUM,"
            strSQL = strSQL & "        CARG_KEY_CHGDATE,"
            strSQL = strSQL & "        CARG_KEY_CHGTANTO,"
            strSQL = strSQL & "        CARG_KEY_CHGRIYUCD,"
            strSQL = strSQL & "        CARG_KEY_CHGNOTE,"
            '2006/05/24 Add tajima End
            strSQL = strSQL & "        CARG_SEIKI, "    ' 請求書発行区分        'INSERT 2010/04/10 K.ISHIZAKA
            strSQL = strSQL & "        CARG_INSED, "
            strSQL = strSQL & "        CARG_INSEJ, "
            strSQL = strSQL & "        CARG_INSPB, "
            strSQL = strSQL & "        CARG_INSUB, "
            strSQL = strSQL & "        CARG_UPDAD, "
            strSQL = strSQL & "        CARG_UPDAJ, "
            strSQL = strSQL & "        CARG_UPDPB, "
            strSQL = strSQL & "        CARG_UPDUB, "
            strSQL = strSQL & " '" & Format$(DATE, "YYYYMMDD") & "', "   ' 削除日付
            strSQL = strSQL & " '" & Format$(time, "HHMMSS") & "', "     ' 削除時刻
            strSQL = strSQL & " '" & Me.Form.NAME & "', "                ' 削除プログラムID
            strSQL = strSQL & " '" & strUserName & "' "                  ' 削除ユーザーID   ' 2006/02/03 ADD T.SUZUKI
'            strSQL = strSQL & " '" & MSZZ000.LsGetUserName & "' "        ' 削除ユーザーID  ' 2006/02/03 DEL T.SUZUKI
            strSQL = strSQL & "   FROM CARG_FILE "
            strSQL = strSQL & "  WHERE CARG_YCODE  = " & Val(Me.cmb_ヤードコード)
            strSQL = strSQL & "    AND CARG_NO     = " & Val(Me.cmb_コンテナ番号)
            strSQL = strSQL & "    AND CARG_UCODE  = " & pvarKcode
            strSQL = strSQL & "    AND CARG_FSDATE = '" & Format(Me.txt_初回契約日, "YYYY/MM/DD") & "' "
            objKONT_DB.Execute strSQL, dbSQLPassThrough

            ' コンテナ契約ファイルの対象データを抽出
            strSQL = "SELECT * " & Chr(13)
            strSQL = strSQL & "  FROM dbo_CARG_FILE " & Chr(13)
            strSQL = strSQL & " WHERE CARG_YCODE  = " & Val(Me.cmb_ヤードコード)
            strSQL = strSQL & "   AND CARG_NO     = " & Val(Me.cmb_コンテナ番号)
            strSQL = strSQL & "   AND CARG_UCODE  = " & pvarKcode
            strSQL = strSQL & "   AND CARG_FSDATE = CDATE(#" & Format(Me.txt_初回契約日, P_FORMAT_YMD) & "#) "
            Set objRsKONT = pWRI_objDb.OpenRecordset(strSQL, dbOpenDynaset)

            If objRsKONT.EOF = True Then
                objRsKONT.Close
                Set objRsKONT = Nothing
                GoTo Exit_rtn
            End If

            ' コンテナ契約ファイルの削除
            With objRsKONT
                .Edit
                .Delete
            End With
            objRsKONT.Close
            Set objRsKONT = Nothing

            ' 予約受付トランに存在するデータを削除した場合、
            ' 予約受付トランの『予約受付状態区分』を"10"(契約受付)に更新
            If pstrUKNO <> "" Then
                If fncUPD_YOUK_TRAN(P_YUKBN_10, strUserName, strErrMsg) = False Then GoTo WRI_Err  ' 2006/02/03 ADD T.SUZUKI
'                If fncUPD_YOUK_TRAN(P_YUKBN_10, strErrMsg) = False Then GoTo WRI_Err              ' 2006/02/03 DEL T.SUZUKI

' 2006/02/03 ADD T.SUZUKI Start
                ' 予約受付トランよりデータを取り込んだ場合、予約ご紹介トランのご紹介区分を"2"(受付)に更新する。
                If fncUPD_INTR_TRAN(P_INTROKBN_03, P_INTROKBN_02, CLng(Me.cmb_ヤードコード), _
                                    CLng(Me.cmb_コンテナ番号), strUserName, strErrMsg) = False Then
                    GoTo WRI_Err
                End If
' 2006/02/03 ADD T.SUZUKI End
            End If
    End Select

    If iMode = 0 Or iMode = 2 Then
        Call fncUPD_RCPT_TRAN(pstrUKNO)                                         'INSERT 2011/04/11 K.ISHIZAKA
        ' 予約データ更新（追加 または 修正時）
        ' コンテナマスタ更新
'↓ UPDATE 2014/06/24 MIYAMOTO
'        If fncUPD_CNTA_MAST(strErrMsg) = False Then GoTo WRI_Err
        If fncUPD_CNTA_MAST(blnKaiYotei, strErrMsg) = False Then GoTo WRI_Err
'↑ UPDATE 2014/06/24 MIYAMOTO
    
'↓ INSERT 2014/04/16 MIYAMOTO
        '鍵区分が自動鍵、且つ契約状態が解約予定の場合
'↓ UPDATE 2016/05/19 MIYAMOTO
'        If Me.cmb_鍵区分.Column(2) = "1" And _
'           blnKaiYotei = True Then
        If MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(2)) = "1" And _
           blnKaiYotei = True Then
'↑ UPDATE 2016/05/19 MIYAMOTO
            '自動鍵解除番号発番トランの登録
'↓ UPDATE 2014/06/24 MIYAMOTO
'            If fncINS_AKKN_TRAN(strUserName, strErrMsg) = False Then GoTo WRI_Err
            If fncINS_AKKN_TRAN(strUserName, Me.cmb_鍵区分, strErrMsg) = False Then GoTo WRI_Err
'↑ UPDATE 2014/06/24 MIYAMOTO
'↓ INSERT 2014/06/24 MIYAMOTO
        Else
            '地方ヤード以外、且つ解約予定で変換対象の鍵区分の場合
'↓ UPDATE 2016/05/19 MIYAMOTO
'            If Me.txt_地方ヤード <> -1 And blnKaiYotei = True And Me.cmb_鍵区分.Column(3) <> -1 Then
            If blnKaiYotei = True And _
                MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(3)) <> 0 Then
'↑ UPDATE 2016/05/19 MIYAMOTO
                '変換鍵区分より自動鍵区分を取得
'↓ UPDATE 2016/05/19 MIYAMOTO
'                strNameFrom = DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & Me.cmb_鍵区分.Column(3))
                strNameFrom = MSZZ070_KaiKey(txt_地方ヤード, DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(3))))
'↑ UPDATE 2016/05/19 MIYAMOTO
                If strNameFrom = "1" Then
                    '自動鍵解除番号発番トランの登録
'↓ UPDATE 2016/05/19 MIYAMOTO
'                    If fncINS_AKKN_TRAN(strUserName, Me.cmb_鍵区分.Column(3), strErrMsg) = False Then GoTo WRI_Err
                    If fncINS_AKKN_TRAN(strUserName, MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(3)), strErrMsg) = False Then GoTo WRI_Err
'↑ UPDATE 2016/05/19 MIYAMOTO
                End If
            End If
'↑ INSERT 2014/06/24 MIYAMOTO
        End If
'↑ INSERT 2014/04/16 MIYAMOTO
    End If

    ' コミット
    pWRI_ObjWS.CommitTrans

    WRI_RTN_For_CARG_FILE = True

Exit_rtn:
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objRsKONT Is Nothing Then objRsKONT.Close: Set objRsKONT = Nothing
    Exit Function

WRI_Err:
    ' ロールバック
    pWRI_ObjWS.Rollback
    ' エラーメッセージ出力
    If strErrMsg <> "" Then MsgBox strErrMsg, vbCritical + vbOKOnly, strMsg_err
    GoTo Exit_rtn

err_rtn:
    strErrMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description
    GoTo WRI_Err
End Function

'==============================================================================*
'
'        MODULE_NAME      :受注契約番号発番
'        MODULE_ID        :WRI_RTN_For_CTL_MST
'        CREATE_DATE      :
'
'1.コントロールマスタ更新
'2.受注契約番号作成
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function WRI_RTN_For_CTL_MST(ByRef strNO As String, lngReqNo As Long) As Boolean  ' 2006/01/05 ADD T.SUZUKI
'Private Function WRI_RTN_For_CTL_MST(strNo As String, lngREQNO As Long) As Boolean       ' 2006/01/05 DEL T.SUZUKI

    Dim strSQL   As String

    Dim objRs    As Recordset
    Dim lngYKNO  As Long       ' 受注契約番号
    Dim strYY    As String
    Dim strMM    As String

    On Error GoTo Err_rtn_ctl

    WRI_RTN_For_CTL_MST = False
    doCmd.SetWarnings False

    ' コントロールマスタより｢契約№｣取得
    strSQL = "SELECT CONT_KYNO, " & Chr(13)
    strSQL = strSQL & " CONT_REQNO, " & Chr(13)
    strSQL = strSQL & " CONT_SKDATE, " & Chr(13)
    strSQL = strSQL & " CONT_BUMOC " & Chr(13)                                                  ' 2006/01/05 ADD T.SUZUKI
    strSQL = strSQL & "  FROM CONT_MAST " & Chr(13)                                             ' 2006/01/05 ADD T.SUZUKI
'    strSQL = strSQL & "  FROM dbo_CONT_MAST " & Chr(13)                                        ' 2006/01/05 DEL T.SUZUKI
    strSQL = strSQL & " WHERE CONT_KEY = 1 "
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)                                 ' 2006/01/05 ADD T.SUZUKI
'    Set objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)                                    ' 2006/01/05 DEL T.SUZUKI

    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        GoTo Exit_rtn_ctl
    End If

    '*** データ更新（コントロールマスタ）***
    With objRs
        .Edit
        .Fields("CONT_KYNO") = .Fields("CONT_KYNO") + 1      ' 顧客No
        .Fields("CONT_REQNO") = .Fields("CONT_REQNO") + 1    ' 請求NO
        .Fields("CONT_SKDATE") = Format(DATE, P_FORMAT_YMD)  ' 更新日                           ' 2006/01/05 ADD T.SUZUKI
'        .Fields("CONT_SKDATE") = Format(DATE, "yyyy/mm/dd")  ' 更新日                          ' 2006/01/05 DEL T.SUZUKI
        .UPDATE
        .Close
    End With
    Set objRs = Nothing

    'データの取得
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)  ' 2006/01/05 ADD T.SUZUKI
'    Set objRs = pobjDB.OpenRecordset("dbo_CONT_MAST", dbOpenSnapshot)                          ' 2006/01/05 DEL T.SUZUKI
    objRs.MoveFirst

    '契約NO
    lngYKNO = objRs.Fields("CONT_KYNO")

    '請求NO
    lngReqNo = objRs.Fields("CONT_REQNO")

    '年取得
    strYY = ""
    Select Case Format(Me.txt_契約開始日, "yyyy")
        Case "1998"
            strYY = "&&"
        Case "1999"
            strYY = "**"
        Case Else
            If Mid(Trim(Format(Me.txt_契約開始日, "yyyy")), 1, 2) = "20" Then
                strYY = Mid(Trim(Format(Me.txt_契約開始日, "yyyy")), 3)
            End If
    End Select

    '月取得
    strMM = ""
    Select Case Format(Me.txt_契約開始日, "m")
        Case "10"  ' 10月
            strMM = "A"
        Case "11"  ' 11月
            strMM = "B"
        Case "12"  ' 12月
            strMM = "C"
        Case Else
            strMM = Trim(Format(Me.txt_契約開始日, "m"))
    End Select

    ' 戻り値｢受注契約番号｣セット
    strNO = objRs.Fields("CONT_BUMOC") + strYY & strMM & Trim(Format$(lngYKNO, "000000"))       ' 2006/01/05 ADD T.SUZUKI
'    strNo = objRs!CONT_BUMOC + strYY + strMM + Trim(Format$(lngYKNO, "000000"))                ' 2006/01/05 DEL T.SUZUKI

    objRs.Close
    Set objRs = Nothing

    WRI_RTN_For_CTL_MST = True

Exit_rtn_ctl:
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

Err_rtn_ctl:
    '*** エラー処理（Ｒｏｌｌｂａｃｋ） ***
    Dim errLoop As Error

    MsgBox "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description, vbInformation + vbOKOnly, strMsg_err
    If DBEngine.Errors.Count > 0 Then
        For Each errLoop In DBEngine.Errors
            MsgBox "Error Number:" + CStr(errLoop.Number) + vbCrLf + errLoop.Description
        Next
    End If
    Err.Clear

    GoTo Exit_rtn_ctl
End Function

'==============================================================================*
'
'        MODULE_NAME      :新規顧客データ登録
'        MODULE_ID        :fncINS_USER_MAST
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncINS_USER_MAST() As Boolean

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo fncINS_USER_MAST_Err

    fncINS_USER_MAST = False
    doCmd.SetWarnings False

    strSQL = " SELECT USER_CODE, " & Chr(13)       ' 顧客ｺｰﾄﾞ
    strSQL = strSQL & " USER_NAME, " & Chr(13)     ' 顧客名
    strSQL = strSQL & " USER_KANA, " & Chr(13)     ' 顧客名ｶﾅ
    strSQL = strSQL & " USER_TANM, " & Chr(13)     ' 代表者名
    strSQL = strSQL & " USER_TAKA, " & Chr(13)     ' 代表者名ｶﾅ
    strSQL = strSQL & " USER_YUBINO, " & Chr(13)   ' 郵便番号
    strSQL = strSQL & " USER_ADR_1, " & Chr(13)    ' 住所1
    strSQL = strSQL & " USER_ADR_2, " & Chr(13)    ' 住所2
    strSQL = strSQL & " USER_ADR_3, " & Chr(13)    ' 住所3
    strSQL = strSQL & " USER_TEL, " & Chr(13)      ' 電話番号
    strSQL = strSQL & " USER_KEITAI, " & Chr(13)   ' 携帯番号
    strSQL = strSQL & " USER_FAX, " & Chr(13)      ' ＦＡＸ
    strSQL = strSQL & " USER_SDAY, " & Chr(13)     ' 締日
    strSQL = strSQL & " USER_KKBN, " & Chr(13)     ' 顧客区分
    strSQL = strSQL & " USER_SKBN, " & Chr(13)     ' 集金区分
    strSQL = strSQL & " USER_BIKO, " & Chr(13)     ' 備考
    strSQL = strSQL & " USER_MAIL, " & Chr(13)     ' メールアドレス
    strSQL = strSQL & " USER_RNAME, " & Chr(13)    ' 連絡先担当者名
    strSQL = strSQL & " USER_RKANA, " & Chr(13)    ' 連絡先担当名カナ
    strSQL = strSQL & " USER_RPOST, " & Chr(13)    ' 連絡先郵便番号
    strSQL = strSQL & " USER_RADR_1, " & Chr(13)   ' 連絡先住所１
    strSQL = strSQL & " USER_RADR_2, " & Chr(13)   ' 連絡先住所２
    strSQL = strSQL & " USER_RADR_3, " & Chr(13)   ' 連絡先住所３
    strSQL = strSQL & " USER_RTEL, " & Chr(13)     ' 連絡先電話番号
    strSQL = strSQL & " USER_RKEITAI, " & Chr(13)  ' 連絡先携帯番号
    strSQL = strSQL & " USER_RFAX, " & Chr(13)     ' 連絡先ＦＡＸ
    strSQL = strSQL & " USER_UPDATE " & Chr(13)    ' 更新日
    strSQL = strSQL & " FROM USER_MAST "
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    With objRs
        .Edit

        .Fields("USER_CODE") = plngUSER_NO
        .Fields("USER_NAME") = Me.txt_顧客名
        .Fields("USER_KANA") = Me.txt_顧客名カナ
        .Fields("USER_TANM") = Me.txt_代表者名
        .Fields("USER_TAKA") = Me.txt_代表者名カナ
        .Fields("USER_YUBINO") = Me.txt_郵便番号
        .Fields("USER_ADR_1") = Me.txt_住所1
        .Fields("USER_ADR_2") = Me.txt_住所2
        .Fields("USER_ADR_3") = Me.txt_住所3
        .Fields("USER_TEL") = Nz(Me.txt_電話番号, "")
        .Fields("USER_KEITAI") = Nz(Me.txt_携帯番号, "")
        .Fields("USER_FAX") = Nz(Me.txt_FAX番号, "")
        .Fields("USER_SDAY") = Nz(Me.txt_締日, 0)
        .Fields("USER_KKBN") = Nz(Me.cmb_顧客区分, 0)
        .Fields("USER_SKBN") = Nz(Me.cmb_集金区分, 0)
        .Fields("USER_BIKO") = Nz(Me.txt_顧客備考, "")
        .Fields("USER_MAIL") = Nz(Me.txt_メールアドレス, "")
        .Fields("USER_RNAME") = Nz(Me.txt_連絡先担当者名, "")
        .Fields("USER_RKANA") = Nz(Me.txt_連絡先担当名カナ, "")
        .Fields("USER_RPOST") = Nz(Me.txt_連絡先郵便番号, "")
        .Fields("USER_RADR_1") = Nz(Me.txt_連絡先住所1, "")
        .Fields("USER_RADR_2") = Nz(Me.txt_連絡先住所2, "")
        .Fields("USER_RADR_3") = Nz(Me.txt_連絡先住所3, "")
        .Fields("USER_RTEL") = Nz(Me.txt_連絡先電話番号, "")
        .Fields("USER_RKEITAI") = Nz(Me.txt_連絡先携帯番号, "")
        .Fields("USER_RFAX") = Nz(Me.txt_連絡先FAX, "")
        .Fields("USER_UPDATE") = Format(DATE, P_FORMAT_YMD)

        .UPDATE
    End With
    objRs.Close
    Set objRs = Nothing

    fncINS_USER_MAST = True

fncINS_USER_MAST_Exit:
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

fncINS_USER_MAST_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

    GoTo fncINS_USER_MAST_Exit
End Function
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :fncINS_USER_MAST
''        CREATE_DATE      :2005/07/13
''        処理内容         ：新規顧客データ登録
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncINS_USER_MAST() As Boolean
'
'    Dim strSQL  As String
''    Dim vRet    As Variant  ' 2005/08/03 DEL T.SUZUKI
'
'    On Error GoTo fncINS_USER_MAST_Err
'
'    fncINS_USER_MAST = False
'
'    doCmd.SetWarnings False
'
'    strSQL = "INSERT INTO dbo_USER_MAST " & Chr(13)
'    strSQL = strSQL & " ( " & Chr(13)
'    strSQL = strSQL & "  USER_CODE, " & Chr(13)     ' 顧客ｺｰﾄﾞ
'    strSQL = strSQL & "  USER_NAME, " & Chr(13)     ' 顧客名
'    strSQL = strSQL & "  USER_KANA, " & Chr(13)     ' 顧客名ｶﾅ
'    strSQL = strSQL & "  USER_TANM, " & Chr(13)     ' 代表者名
'    strSQL = strSQL & "  USER_TAKA, " & Chr(13)     ' 代表者名ｶﾅ
'    strSQL = strSQL & "  USER_YUBINO, " & Chr(13)   ' 郵便番号
'    strSQL = strSQL & "  USER_ADR_1, " & Chr(13)    ' 住所1
'    strSQL = strSQL & "  USER_ADR_2, " & Chr(13)    ' 住所2
'    strSQL = strSQL & "  USER_ADR_3, " & Chr(13)    ' 住所3
'    strSQL = strSQL & "  USER_TEL, " & Chr(13)      ' 電話番号
'    strSQL = strSQL & "  USER_KEITAI, " & Chr(13)   ' 携帯番号
'    strSQL = strSQL & "  USER_FAX, " & Chr(13)      ' ＦＡＸ
'    strSQL = strSQL & "  USER_SDAY, " & Chr(13)     ' 締日
'    strSQL = strSQL & "  USER_KKBN, " & Chr(13)     ' 顧客区分
'    strSQL = strSQL & "  USER_SKBN, " & Chr(13)     ' 集金区分
'    strSQL = strSQL & "  USER_BIKO, " & Chr(13)     ' 備考
'    strSQL = strSQL & "  USER_MAIL, " & Chr(13)     ' メールアドレス
'    strSQL = strSQL & "  USER_RNAME, " & Chr(13)    ' 連絡先担当者名
'    strSQL = strSQL & "  USER_RKANA, " & Chr(13)    ' 連絡先担当名カナ
'    strSQL = strSQL & "  USER_RPOST, " & Chr(13)    ' 連絡先郵便番号
'    strSQL = strSQL & "  USER_RADR_1, " & Chr(13)   ' 連絡先住所１
'    strSQL = strSQL & "  USER_RADR_2, " & Chr(13)   ' 連絡先住所２
'    strSQL = strSQL & "  USER_RADR_3, " & Chr(13)   ' 連絡先住所３
'    strSQL = strSQL & "  USER_RTEL, " & Chr(13)     ' 連絡先電話番号
'    strSQL = strSQL & "  USER_RKEITAI, " & Chr(13)  ' 連絡先携帯番号
'    strSQL = strSQL & "  USER_RFAX, " & Chr(13)     ' 連絡先ＦＡＸ
'    strSQL = strSQL & "  USER_UPDATE " & Chr(13)    ' 更新日
'    strSQL = strSQL & " ) " & Chr(13)
'    strSQL = strSQL & " VALUES " & Chr(13)
'    strSQL = strSQL & " ( " & Chr(13)
'    strSQL = strSQL & plngUSER_NO & ", " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_顧客名 & "', " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_顧客名カナ & "', " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_代表者名 & "', " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_代表者名カナ & "', " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_郵便番号 & "', " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_住所1 & "', " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_住所2 & "', " & Chr(13)
'    strSQL = strSQL & "'" & Me.txt_住所3 & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_電話番号, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_携帯番号, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_FAX番号, "") & "', " & Chr(13)
'    strSQL = strSQL & Nz(Me.txt_締日, 0) & ", " & Chr(13)
'    strSQL = strSQL & Nz(Me.cmb_顧客区分, 0) & ", " & Chr(13)
'    strSQL = strSQL & Nz(Me.cmb_集金区分, 0) & ", " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_顧客備考, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_メールアドレス, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先担当者名, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先担当名カナ, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先郵便番号, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先住所1, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先住所2, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先住所3, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先電話番号, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先携帯番号, "") & "', " & Chr(13)
'    strSQL = strSQL & "'" & Nz(Me.txt_連絡先FAX, "") & "', " & Chr(13)
'    strSQL = strSQL & " Cdate(#" & Format(DATE, "yyyy/mm/dd") & "#)" & Chr(13)
'    strSQL = strSQL & " ) " & Chr(13)
'
'    doCmd.RunSQL strSQL, False
'
'    fncINS_USER_MAST = True
'
'fncINS_USER_MAST_Exit:
'    doCmd.SetWarnings True
'    Exit Function
'
'fncINS_USER_MAST_Err:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    Err.Clear
'
'    GoTo fncINS_USER_MAST_Exit
'End Function

'==============================================================================*
'
'        MODULE_NAME      :新規顧客データ登録
'        MODULE_ID        :fncSEL_UPD_CONT_MAST
'        CREATE_DATE      :2005/07/13
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncSEL_UPD_CONT_MAST() As Boolean

    Dim strSQL  As String
    Dim objRs   As Recordset
    Dim vRet    As Variant

    On Error GoTo fncSEL_UPD_CONT_MAST_Err

    fncSEL_UPD_CONT_MAST = False
    doCmd.SetWarnings False

    ' 顧客No初期化
    plngUSER_NO = 0

    ' コントロールマスタより顧客Noを取得
    strSQL = "SELECT CONT_USER_NO, " & Chr(13)
    strSQL = strSQL & " CONT_UPDATE " & Chr(13)
    strSQL = strSQL & "  FROM CONT_MAST " & Chr(13)                                             ' 2006/01/05 ADD T.SUZUKI
'    strSQL = strSQL & "  FROM dbo_CONT_MAST " & Chr(13)                                        ' 2006/01/05 DEL T.SUZUKI
    strSQL = strSQL & " WHERE CONT_KEY = 1 "
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)                                 ' 2006/01/05 ADD T.SUZUKI
'    sSet objRs = pobjDB.OpenRecordset(strSQL, dbOpenDynaset)                                   ' 2006/01/05 DEL T.SUZUKI

    If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        GoTo fncSEL_UPD_CONT_MAST_Exit
    End If

    With objRs
        .Edit
        .Fields("CONT_USER_NO") = .Fields("CONT_USER_NO") + 1        ' 顧客No
        .Fields("CONT_UPDATE") = Format(Now, "YYYY/MM/DD HH:MM:SS")  ' 更新日
        .UPDATE
        .Close
    End With
    Set objRs = Nothing

    'データの取得
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)  ' 2006/01/05 ADD T.SUZUKI
'    Set objRs = pobjDB.OpenRecordset("dbo_CONT_MAST", dbOpenSnapshot)                          ' 2006/01/05 DEL T.SUZUKI
    objRs.MoveFirst

    ' 顧客No
    plngUSER_NO = objRs!CONT_USER_NO

    objRs.Close
    Set objRs = Nothing

    ' 存在チェック
    vRet = ""
    vRet = Nz(DLookup("USER_CODE", "dbo_USER_MAST", "USER_CODE = " & plngUSER_NO), "")

    If vRet <> "" Then
        MsgBox "顧客コードが重複しています。", vbInformation, strMsg_err
        plngUSER_NO = 0
        Exit Function
    End If

' 2005/09/12 ADD T.SUZUKI Start
    ' 新規に発番した顧客コードをコンボボックスに設定する
    Me.cmb_顧客コード = Format(plngUSER_NO, "000000")
' 2005/09/12 ADD T.SUZUKI Start

    fncSEL_UPD_CONT_MAST = True

fncSEL_UPD_CONT_MAST_Exit:
    doCmd.SetWarnings True
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

fncSEL_UPD_CONT_MAST_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    plngUSER_NO = 0

    GoTo fncSEL_UPD_CONT_MAST_Exit
End Function

' DEL 2006/03/03 Tajima Start 受付対応に伴い新規契約は無くなるので削除
''==============================================================================*
''
''        MODULE_NAME      :請求データ出力
''        MODULE_ID        :WRI_RTN_For_REQU_FILE
''        CREATE_DATE      :2006/01/05
''
'' AddNew UpdateをInsert文に変更（複数回実行するとAddNew Updateができなくなるため)
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function WRI_RTN_For_REQU_FILE(ByVal lngREQU_NO As Long, _
'                                       ByVal strKIYAKU_NO As String, _
'                                       ByRef strErrMsg As String) As Boolean
'
'    Dim strSql      As String
'    Dim objRs       As Recordset
'    Dim strUSER_NO  As String     ' 顧客ｺｰﾄﾞ
'    Dim strNKDATE   As String     ' 入金予定日
'    Dim dblKINGAK   As Double     ' 金額(金額 + 鍵代)
'    Dim dblTOTAL    As Double     ' 合計
'    Dim dtSYSDATE   As Date
'    Dim dtWKYYMM    As Date
'
'    On Error GoTo WRI_RTN_For_REQU_FILE_Err
'
'    strErrMsg = ""
'
'    dtSYSDATE = CDate(Format$(DATE, P_FORMAT_YMD))                  ' システム日付(YYYY/MM/DD)
'    dtWKYYMM = CDate(Format(Me.txt_契約開始日, "YYYY/MM") & "/01")  ' 請求対象年月（売上計上）(YYYY/MM/DD)
'
'    ' 請求ファイル抽出
'    strSql = "SELECT * FROM REQU_FILE "
'    Set objRs = objKONT_DB.OpenRecordSet(strSql, dbOpenDynaset)
'
'    With objRs
'        .AddNew
'
'        .Fields("REQU_REQNO") = CStr(lngREQU_NO)                          ' 請求№
'        .Fields("REQU_YYYYMM") = Format(Me.txt_初回契約日, P_FORMAT_YMD)  ' 請求年月
'        .Fields("REQU_KJDATE") = Format(dtWKYYMM, P_FORMAT_YMD)           ' 請求対象年月（売上計上）
'        .Fields("REQU_SDATE") = Format(Me.txt_初回契約日, P_FORMAT_YMD)   ' 請求年月日
'
'        ' 入金予定日
'        strNKDATE = ""
'        strNKDATE = Format(DateAdd("D", -1, CDate(Format(Me.txt_契約開始日, "YYYY/MM") & "/01")), "YYYY/MM/DD")
'        .Fields("REQU_NKDATE") = strNKDATE
'
'        ' 顧客コード
'        strUSER_NO = ""
'        If (Me.opg_処理選択.Value = 1 And Me.opg_顧客情報取得選択.Value = 1) Then
'            strUSER_NO = Format(plngUSER_NO, "000000")
'        Else
'            strUSER_NO = Format(Me.cmb_顧客コード, "000000")
'        End If
'        .Fields("REQU_KCODE") = strUSER_NO
'
'        .Fields("REQU_NKBN") = P_ZERO                                     ' 入金区分
'        .Fields("REQU_YCODE") = Format(Me.cmb_ヤードコード, "000000")     ' ﾔｰﾄﾞｺｰﾄﾞ
'        .Fields("REQU_NO") = Format(Me.cmb_コンテナ番号, "000000")        ' ｺﾝﾃﾅ番号
'        .Fields("REQU_TTANKA") = P_ZERO                                   ' 坪単価
'        .Fields("REQU_TUBOSU") = P_ZERO                                   ' 坪数
'
'        '金額(金額 + 鍵代)
'        dblKINGAK = P_ZERO
'        dblKINGAK = Val(Nz(Me.txt_初回金額, P_ZERO)) + Val(Nz(Me.txt_初回鍵代, P_ZERO))
'        If dblKINGAK > pcsglMoneySumMaxValue Then
'            .Fields("REQU_KINGAK") = pcsglMoneySumMaxValue
'        Else
'            .Fields("REQU_KINGAK") = dblKINGAK
'        End If
'
'        .Fields("REQU_SYOZEI") = Val(Nz(Me.txt_初回消費税, P_ZERO))       ' 消費税
'        .Fields("REQU_SECUKG") = Val(Nz(Me.txt_預り保証金, P_ZERO))       ' 保証金
'
'        ' 合計
'        dblTOTAL = Val(Nz(Me.txt_初回金額, P_ZERO)) + _
'                   Val(Nz(Me.txt_初回鍵代, P_ZERO)) + _
'                   Val(Nz(Me.txt_初回消費税, P_ZERO)) + _
'                   Val(Nz(Me.txt_預り保証金, P_ZERO))
'        If dblTOTAL > pcsglMoneySumMaxValue Then
'            .Fields("REQU_TOTAL") = pcsglMoneySumMaxValue
'        Else
'            .Fields("REQU_TOTAL") = dblTOTAL
'        End If
'
'        .Fields("REQU_TEKI") = Null                                       ' 摘要
'        .Fields("REQU_FLG") = 1                                           ' 請求修正ﾌﾗｸﾞ
'        .Fields("REQU_UPDATE") = Format(dtSYSDATE, P_FORMAT_YMD)          ' 更新日
'        .Fields("REQU_ACPTNO") = strKIYAKU_NO                             ' 受注契約番号
'
'        .UPDATE
'    End With
'    objRs.Close
'    Set objRs = Nothing
'
'    WRI_RTN_For_REQU_FILE = True
'
'WRI_RTN_For_REQU_FILE_Exit:
'    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
'    Exit Function
'
'WRI_RTN_For_REQU_FILE_Err:
'    strErrMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description
'  GoTo WRI_RTN_For_REQU_FILE_Exit
'End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncUPD_REQU_FILE
'        Parameter        :strKIYAKU_NO = 受注契約番号
'                          strYYYYMM    = 初回契約年月(YYYY/MM)
'                          strErrMsg    = エラーメッセージ格納
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncUPD_REQU_FILE(ByVal strKIYAKU_NO As String, ByVal strYYYYMM As String, _
'                                  ByRef strErrMsg As String) As Boolean
'
'    Dim strSql    As String
'    Dim objRs     As Recordset
'    Dim dblKINGAK As Double
'    Dim dblTOTAL  As Double
'
'    On Error GoTo fncUPD_REQU_FILE_Err
'
'    fncUPD_REQU_FILE = False
'    strErrMsg = ""
'
'    strSql = " SELECT REQU_KINGAK, " & Chr(13)    ' 金額
'    strSql = strSql & " REQU_SYOZEI, " & Chr(13)  ' 消費税
'    strSql = strSql & " REQU_SECUKG, " & Chr(13)  ' 保証金
'    strSql = strSql & " REQU_TOTAL, " & Chr(13)   ' 合計
'    strSql = strSql & " REQU_FLG, " & Chr(13)     ' 請求修正フラグ
'    strSql = strSql & " REQU_UPDATE " & Chr(13)   ' 更新日
'    strSql = strSql & "  FROM REQU_FILE " & Chr(13)
'    strSql = strSql & " WHERE REQU_ACPTNO                        = '" & strKIYAKU_NO & "' " & Chr(13)  ' 受注契約番号
'    strSql = strSql & "   AND format(REQU_KJDATE,'yyyy/mm') = '" & strYYYYMM & "' "                    ' 請求対象年月売上計上
''    strSQL = strSQL & "   AND CONVERT(CHAR(7), REQU_KJDATE, 111) = '" & strYYYYMM & "' "               ' 請求対象年月売上計上
'    Set objRs = objKONT_DB.OpenRecordSet(strSql, dbOpenDynaset)
''    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough)
'
'    If objRs.EOF = False Then
'        With objRs
'            .Edit
'
'            '金額(金額 + 鍵代)
'            dblKINGAK = 0
'            dblKINGAK = Val(Nz(Me.txt_初回金額, 0)) + Val(Nz(Me.txt_初回鍵代, P_ZERO))
'            If dblKINGAK > pcsglMoneySumMaxValue Then
'                .Fields("REQU_KINGAK") = pcsglMoneySumMaxValue
'            Else
'                .Fields("REQU_KINGAK") = dblKINGAK
'            End If
'
'            .Fields("REQU_SYOZEI") = Val(Nz(Me.txt_初回消費税, P_ZERO))       ' 消費税
'            .Fields("REQU_SECUKG") = Val(Nz(Me.txt_預り保証金, P_ZERO))       ' 保証金
'
'            ' 合計
'            dblTOTAL = Val(Nz(Me.txt_初回金額, P_ZERO)) + _
'                       Val(Nz(Me.txt_初回鍵代, P_ZERO)) + _
'                       Val(Nz(Me.txt_初回消費税, P_ZERO)) + _
'                       Val(Nz(Me.txt_預り保証金, P_ZERO))
'            If dblTOTAL > pcsglMoneySumMaxValue Then
'                .Fields("REQU_TOTAL") = pcsglMoneySumMaxValue
'            Else
'                .Fields("REQU_TOTAL") = dblTOTAL
'            End If
'
'            .Fields("REQU_FLG") = 2                                           ' 請求修正ﾌﾗｸﾞ
'            .Fields("REQU_UPDATE") = Format(DATE, P_FORMAT_YMD)               ' 更新日
'
'            .UPDATE
'        End With
'    End If
'    objRs.Close
'    Set objRs = Nothing
'
'    fncUPD_REQU_FILE = True
'
'fncUPD_REQU_FILE_Exit:
'    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
'    Exit Function
'
'fncUPD_REQU_FILE_Err:
'    strErrMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description
'End Function
'DEL 2006/03/03 Tajima End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncUPD_YOUK_TRAN
'        Parameter        :intYOUKT_YUKBN P_YUKBN_20 = 契約
'                                         P_YUKBN_10 = 契約受付
'                          strUserName    ユーザー名
'                          strErrMsg      エラーメッセージ
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2006/02/03 REP T.SUZUKI Start
Private Function fncUPD_YOUK_TRAN(ByVal intYOUKT_YUKBN As Integer, ByVal strUserName As String, _
                                  ByRef strErrMsg As String) As Boolean
'Private Function fncUPD_YOUK_TRAN(ByVal intYOUKT_YUKBN As Integer, ByRef strErrMsg As String) As Boolean
' 2006/02/03 REP T.SUZUKI End

    Dim strSQL As String
    Dim objRs  As Recordset

    On Error GoTo fncUPD_YOUK_TRAN_Err

    fncUPD_YOUK_TRAN = False
    strErrMsg = ""

    strSQL = " SELECT YOUKT_YUKBN, " & Chr(13)    ' 予約受付状態区分 = 20(契約)
    strSQL = strSQL & " YOUKT_UPDAD, " & Chr(13)  ' 更新日付
    strSQL = strSQL & " YOUKT_UPDAJ, " & Chr(13)  ' 更新時刻
    strSQL = strSQL & " YOUKT_UPDPB, " & Chr(13)  ' 更新プログラムID
    strSQL = strSQL & " YOUKT_UPDUB " & Chr(13)   ' 更新ユーザーID
    strSQL = strSQL & "  FROM YOUK_TRAN " & Chr(13)
    strSQL = strSQL & " WHERE YOUKT_UKNO = '" & pstrUKNO & "' "
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    With objRs
        If .EOF = False Then
            .Edit
    
            .Fields("YOUKT_YUKBN") = intYOUKT_YUKBN
            .Fields("YOUKT_UPDAD") = Format$(DATE, "YYYYMMDD")
            .Fields("YOUKT_UPDAJ") = Format$(time, "HHMMSS")
            .Fields("YOUKT_UPDPB") = Me.Form.NAME
            .Fields("YOUKT_UPDUB") = strUserName             ' 2006/02/03 REP T.SUZUKI
    
            .UPDATE
        End If
    End With
    objRs.Close
    Set objRs = Nothing

    fncUPD_YOUK_TRAN = True

fncUPD_YOUK_TRAN_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

fncUPD_YOUK_TRAN_Err:
    strErrMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description  ' 2006/01/05 ADD T.SUZUKI
    GoTo fncUPD_YOUK_TRAN_Exit
End Function
' 2006/01/05 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :
''        MODULE_ID        :fncUPD_YOUK_TRAN
''        Parameter        :intYOUKT_YUKBN 20 = 契約
''                                         10 = 契約受付
''        CREATE_DATE      :2005/08/03
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncUPD_YOUK_TRAN(intYOUKT_YUKBN As Integer) As Boolean
'
'    Dim strSQL As String
'
'    On Error GoTo fncUPD_YOUK_TRAN_Err
'
'    fncUPD_YOUK_TRAN = False
'
'    strSQL = "UPDATE  dbo_YOUK_TRAN " & Chr(13)
'    strSQL = strSQL & "   SET  YOUKT_YUKBN = " & intYOUKT_YUKBN & Chr(13)                 ' 予約受付状態区分 = 20(契約)
'    strSQL = strSQL & "       ,YOUKT_UPDAD = " & Format$(DATE, "yyyymmdd") & Chr(13)      ' 更新日付
'    strSQL = strSQL & "       ,YOUKT_UPDAJ = " & Format$(Time, "hhmmss") & Chr(13)        ' 更新時刻
'' 2005/08/03 REP T.SUZUKI Start
'    strSQL = strSQL & "       ,YOUKT_UPDPB = '" & pstrFormName & "'" & Chr(13)            ' 更新プログラムID
''    strSQL = strSQL & "       ,YOUKT_UPDPB = '" & PROG_ID & "'" & Chr(13)                 ' 更新プログラムID
'' 2005/08/03 REP T.SUZUKI End
'    strSQL = strSQL & "       ,YOUKT_UPDUB = '" & MSZZ000.LsGetUserName & "' " & Chr(13)  ' 更新ユーザーID
'    strSQL = strSQL & " WHERE YOUKT_UKNO = '" & pstrUKNO & "' "
'
'    pWRI_objDb.Execute strSQL, dbFailOnError
'
'    fncUPD_YOUK_TRAN = True
'
'    Exit Function
'
'fncUPD_YOUK_TRAN_Err:
'    Exit Function
'End Function
' 2006/01/05 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncUPD_INTR_TRAN
'        Parameter        :strINTROKBN_F：ご紹介区分(更新条件)  P_INTROKBN_02 = 受付
'                                                               P_INTROKBN_03 = 契約
'                          strINTROKBN_T：ご紹介区分(更新値)    P_INTROKBN_02 = 受付
'                                                               P_INTROKBN_03 = 契約
'                          lngYCODE     ：ヤードコード
'                          lngNO        ：コンテナ番号
'                          strUserName  ：ユーザー名
'                          strErrMsg    ：エラーメッセージ
'        CREATE_DATE      :2006/02/02
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUPD_INTR_TRAN(ByVal strINTROKBN_F As String, ByVal strINTROKBN_T As String, _
                                  ByVal lngYCODE As Long, ByVal lngNo As Long, _
                                  ByVal strUserName As String, ByRef strErrMsg As String) As Boolean

    Dim strSQL As String
    Dim objRs  As Recordset

    On Error GoTo fncUPD_INTR_TRAN_Err

    fncUPD_INTR_TRAN = False

    strErrMsg = ""

    strSQL = " SELECT INTRT_INTROKBN, " & Chr(13)  ' ご紹介区分 = 3(契約)
    strSQL = strSQL & " INTRT_UPDAD, " & Chr(13)   ' 更新日付
    strSQL = strSQL & " INTRT_UPDAJ, " & Chr(13)   ' 更新時刻
    strSQL = strSQL & " INTRT_UPDPB, " & Chr(13)   ' 更新プログラムID
    strSQL = strSQL & " INTRT_UPDUB " & Chr(13)    ' 更新ユーザーID
    strSQL = strSQL & "  FROM INTR_TRAN " & Chr(13)
    strSQL = strSQL & " WHERE INTRT_UKNO  = '" & pstrUKNO & "' " & Chr(13)  ' 予約受付番号
    strSQL = strSQL & "   AND INTRT_YCODE = " & lngYCODE & Chr(13)          ' ヤードコード
    strSQL = strSQL & "   AND INTRT_NO    = " & lngNo                       ' コンテナ番号
    strSQL = strSQL & "   AND INTRT_INTROKBN = '" & strINTROKBN_F & "' "    ' ご紹介区分(2：受付)
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    If objRs.EOF = False Then
        With objRs
            .Edit

            .Fields("INTRT_INTROKBN") = strINTROKBN_T
            .Fields("INTRT_UPDAD") = Format$(DATE, "YYYYMMDD")
            .Fields("INTRT_UPDAJ") = Format$(time, "HHMMSS")
            .Fields("INTRT_UPDPB") = Me.Form.NAME
            .Fields("INTRT_UPDUB") = strUserName

            .UPDATE
        End With
    End If
    objRs.Close
    Set objRs = Nothing

    fncUPD_INTR_TRAN = True

fncUPD_INTR_TRAN_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

fncUPD_INTR_TRAN_Err:
    strErrMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description
    GoTo fncUPD_INTR_TRAN_Exit
End Function

'↓ INSERT 2014/04/16 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncINS_AKKN_TRAN
'        Parameter        :strErrMsg エラーメッセージ格納
'        CREATE_DATE      :2014/04/16
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓ UPDATE 2014/06/24 MIYAMOTO
Private Function fncINS_AKKN_TRAN(ByVal strUserName As String, _
                                  ByVal strKagiKbn As String, _
                                  ByRef strErrMsg As String) As Boolean
'Private Function fncINS_AKKN_TRAN(ByVal strUserName As String, _
'                                  ByRef strErrMsg As String) As Boolean
'↑ UPDATE 2014/06/24 MIYAMOTO

    Dim strSQL              As String
    Dim strNewKaiNo         As String
    Dim strBumoCD           As String   ' 部門コード
    Dim objRs               As Object
    Dim dbSQLServer         As Object 'INSERT 2021/06/04 EGL

    On Error GoTo fncINS_AKKN_TRAN_Err

    fncINS_AKKN_TRAN = False

    ' 部門コード／部門名取得
    Call subGetBumonName(pobjDB, strBumoCD)

    'SQLサーバー接続
    Set dbSQLServer = ADODB_Connection(strBumoCD)  '2021/06/02 insert EGL

    '新解除番号の発番
    strNewKaiNo = MSZZ070_M00(strBumoCD, Me.cmb_ヤードコード, Me.cmb_コンテナ番号, Me.cmb_鍵区分)

    '自動鍵解除番号発番トランの登録
    strSQL = ""
    strSQL = strSQL & "SELECT      *" & vbCrLf
    strSQL = strSQL & "FROM        AKKN_TRAN" & vbCrLf
    strSQL = strSQL & "WHERE       AKKNT_YARD = " & Val(Me.cmb_ヤードコード) & vbCrLf
    strSQL = strSQL & "AND         AKKNT_CTNO = " & Val(Me.cmb_コンテナ番号) & vbCrLf
'↓ UPDATE 2014/06/24 MIYAMOTO
'    strSQL = strSQL & "AND         AKKNT_KTYPE = '" & Nz(Me.cmb_鍵区分) & "'" & vbCrLf
    strSQL = strSQL & "AND         AKKNT_KTYPE = '" & Nz(strKagiKbn) & "'" & vbCrLf
'↑ UPDATE 2014/06/24 MIYAMOTO
    strSQL = strSQL & "AND         AKKNT_STYPE = '4'" & vbCrLf 'INSERT 2021/06/04 EGL
    'Set objRs = objKONT_DB.OpenRecordset(strSql, dbOpenDynaset)
    Set objRs = ADODB_Recordset(strSQL, dbSQLServer, adoAppendOnly) '2021/06/02 insert EGL

    With objRs
        .AddNew

        .Fields("AKKNT_YARD") = Val(Me.cmb_ヤードコード)                                    ' ヤードコード
        .Fields("AKKNT_CTNO") = Val(Me.cmb_コンテナ番号)                                    ' コンテナ番号
'↓ UPDATE 2014/06/24 MIYAMOTO
'        .Fields("AKKNT_KTYPE") = Nz(Me.cmb_鍵区分)                                          ' 鍵区分
        .Fields("AKKNT_KTYPE") = Nz(strKagiKbn)                                             ' 鍵区分
'↑ UPDATE 2014/06/24 MIYAMOTO
        .Fields("AKKNT_STYPE") = "4"                                                        ' 副種別 INSERT 2021/06/04 EGL
        .Fields("AKKNT_HATUD") = Format$(DATE, "YYYYMMDD") & Format$(time, "HHMMSS")        ' 解除番号発番日時
        .Fields("AKKNT_KAINO") = Nz(strNewKaiNo)                                            ' 解除番号
        .Fields("AKKNT_INSED") = Format$(DATE, "YYYYMMDD")                                  ' 作成日付
        .Fields("AKKNT_INSEJ") = Format$(time, "HHMMSS")                                    ' 作成時刻
        .Fields("AKKNT_INSPB") = PROG_ID                                                    ' 作成プログラムID
        .Fields("AKKNT_INSUB") = LsGetUserName()                                            ' 作成ユーザーID

        .UPDATE
    End With

'↓ INSERT 2021/06/04 EGL
    If Me.cmb_鍵区分 = 16 Then
        '新解除番号の発番
        strNewKaiNo = MSZZ070_M01(strBumoCD, Me.cmb_鍵区分)
        '自動鍵解除番号発番トランの登録
        strSQL = ""
        strSQL = strSQL & "SELECT      *" & vbCrLf
        strSQL = strSQL & "FROM        AKKN_TRAN" & vbCrLf
        strSQL = strSQL & "WHERE       AKKNT_YARD = " & Val(Me.cmb_ヤードコード) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_CTNO = " & Val(Me.cmb_コンテナ番号) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_KTYPE = '" & Nz(strKagiKbn) & "'" & vbCrLf
        strSQL = strSQL & "AND         AKKNT_STYPE = '6'" & vbCrLf
        Set objRs = ADODB_Recordset(strSQL, dbSQLServer, adoAppendOnly)
        With objRs
            .AddNew
            .Fields("AKKNT_YARD") = Val(Me.cmb_ヤードコード)                                    ' ヤードコード
            .Fields("AKKNT_CTNO") = Val(Me.cmb_コンテナ番号)                                    ' コンテナ番号
            .Fields("AKKNT_KTYPE") = Nz(strKagiKbn)                                             ' 鍵区分
            .Fields("AKKNT_STYPE") = "6"                                                        ' 副種別 INSERT 2021/06/04 EGL
            .Fields("AKKNT_HATUD") = Format$(DATE, "YYYYMMDD") & Format$(time, "HHMMSS")        ' 解除番号発番日時
            .Fields("AKKNT_KAINO") = Nz(strNewKaiNo)                                            ' 解除番号
            .Fields("AKKNT_INSED") = Format$(DATE, "YYYYMMDD")                                  ' 作成日付
            .Fields("AKKNT_INSEJ") = Format$(time, "HHMMSS")                                    ' 作成時刻
            .Fields("AKKNT_INSPB") = PROG_ID                                                    ' 作成プログラムID
            .Fields("AKKNT_INSUB") = LsGetUserName()                                            ' 作成ユーザーID
            .UPDATE
        End With
End If
'↑ INSERT 2021/06/04 EGL


    fncINS_AKKN_TRAN = True

fncINS_AKKN_TRAN_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

fncINS_AKKN_TRAN_Err:
    strErrMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description
    GoTo fncINS_AKKN_TRAN_Exit

End Function
'↑ INSERT 2014/04/16 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :fncUPD_CNTA_MAST
'        Parameter        :strErrMsg エラーメッセージ格納
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓ UPDATE 2014/06/24 MIYAMOTO
'Private Function fncUPD_CNTA_MAST(ByRef strErrMsg As String) As Boolean
Private Function fncUPD_CNTA_MAST(ByVal blnKaiYotei As Boolean, ByRef strErrMsg As String) As Boolean
'↑ UPDATE 2014/06/24 MIYAMOTO

    Dim strSQL As String
    Dim objRs  As Recordset

    fncUPD_CNTA_MAST = False
    strErrMsg = ""

    strSQL = " SELECT CNTA_YNAME, " & Chr(13)
    strSQL = strSQL & " CNTA_YTEL, " & Chr(13)
    strSQL = strSQL & " CNTA_UPDATE " & Chr(13)
    '2006/05/10 鍵情報の保存 Add Start Tajima
    strSQL = strSQL & ",CNTA_KEY_TYPE " & Chr(13)
    strSQL = strSQL & ",CNTA_KEY_NO " & Chr(13)
    strSQL = strSQL & ",CNTA_KEY_NUM " & Chr(13)
    '2006/05/10 鍵情報の保存 Add End Tajima
    strSQL = strSQL & ",CNTA_USAGE " & Chr(13)  '2017/07/29
    strSQL = strSQL & "  FROM CNTA_MAST " & Chr(13)
    strSQL = strSQL & " WHERE CNTA_NO   = " & Me.cmb_コンテナ番号 & Chr(13)
    strSQL = strSQL & "   AND CNTA_CODE = " & Me.cmb_ヤードコード
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    With objRs
        .Edit

        '2005/09/22 SHIBAZAKI 予約名／予約番号削除
'        .Fields("CNTA_YNAME") = Me.txt_予約名
'        .Fields("CNTA_YTEL") = Me.txt_予約番号
        .Fields("CNTA_UPDATE") = Format(Now, P_FORMAT_YMD)

        '2015/03/26 M.HONDA INS
        '既に自動発番されているデータは鍵区分を書き換えない。
'↓ UPDATE 2016/05/19 MIYAMOTO
'        If .Fields("CNTA_KEY_TYPE") <> 15 Then
        If MSZZ070_KaiKey(txt_地方ヤード, DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & .Fields("CNTA_KEY_TYPE"))) = 0 Then
'↑ UPDATE 2016/05/19 MIYAMOTO
            '2006/05/10 鍵情報の保存 Add Start Tajima
            If pintAGRE < 4 Then
              'メンテ前ならコンテナマスタに反映する
    '↓ UPDATE 2014/06/24 MIYAMOTO
    '          .Fields("CNTA_KEY_TYPE") = Me.cmb_鍵区分
                '鍵区分が自動鍵の場合
'↓ UPDATE 2016/05/19 MIYAMOTO
'                If Me.cmb_鍵区分.Column(2) = 1 Then
                If MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(2)) = 1 Then
'↑ UPDATE 2016/05/19 MIYAMOTO
                    .Fields("CNTA_KEY_TYPE") = Me.cmb_鍵区分
                '鍵区分が自動鍵以外の場合
                Else
                    '鍵区分が変換対象ではない場合
'↓ UPDATE 2016/05/19 MIYAMOTO
'                    If Me.cmb_鍵区分.Column(3) = -1 Then
                    If MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(3)) = 0 Then
'↑ UPDATE 2016/05/19 MIYAMOTO
                        .Fields("CNTA_KEY_TYPE") = Me.cmb_鍵区分
                    '鍵区分が変換対象の場合
                    Else
'↓ UPDATE 2016/05/19 MIYAMOTO
'                        '地方ヤード以外で解約予定の場合
'                        If Me.txt_地方ヤード <> -1 And blnKaiYotei = True Then
'                            .Fields("CNTA_KEY_TYPE") = Me.cmb_鍵区分.Column(3)
                        '解約予定の場合
                        If blnKaiYotei = True Then
                            .Fields("CNTA_KEY_TYPE") = MSZZ070_KaiKey(txt_地方ヤード, Me.cmb_鍵区分.Column(3))
'↑ UPDATE 2016/05/19 MIYAMOTO
                        Else
                            .Fields("CNTA_KEY_TYPE") = Me.cmb_鍵区分
                        End If
                    End If
                End If
    '↑ UPDATE 2014/06/24 MIYAMOTO
              .Fields("CNTA_KEY_NO") = Me.txt_解除番号
              .Fields("CNTA_KEY_NUM") = Nz(Me.txt_貸出本数, 0)
            End If
            '2006/05/10 鍵情報の保存 Add End Tajima
        End If
        
        '2017/07/29 M.HONDA INS
        '未使用ｷｬﾝｾﾙの場合にはｺﾝﾃﾅ(ﾊﾞｲｸ)をｺﾝﾃﾅに戻す
        If Me.txt_解約日付 < Me.txt_初回契約日 Then
            If .Fields("CNTA_USAGE") = "33" Then
                .Fields("CNTA_USAGE") = "0"
            End If
        End If
        '2017/07/29 M.HONDA INS
        
        '2019/06/15 N.IMAI Start
        '未使用ｷｬﾝｾﾙの場合にはﾊﾞｲｸ(ｺﾝﾃﾅ)をバイクに戻す
        If Me.txt_解約日付 < Me.txt_初回契約日 Then
            If .Fields("CNTA_USAGE") = "39" Then
                .Fields("CNTA_USAGE") = "3"
            End If
        End If
        '2019/06/15 N.IMAI End
        
        .UPDATE
    End With
    objRs.Close
    Set objRs = Nothing

    fncUPD_CNTA_MAST = True

fncUPD_CNTA_MAST_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    Exit Function

fncUPD_CNTA_MAST_Err:
    strErrMsg = "エラーが発生しました。" & vbCrLf & "ｴﾗｰ番号:" & Err.Number & vbCr & Err.Description  ' 2006/01/05 ADD T.SUZUKI
    GoTo fncUPD_CNTA_MAST_Exit
End Function

'==============================================================================*
'
'       MODULE_NAME     : 受付トラン更新（緊急連絡先）
'                         ※受付取込の対象にしたくないので更新日等は更新しない
'       MODULE_ID       : fncUPD_RCPT_TRAN
'       CREATE_DATE     : 2011/04/11            K.ISHIZAKA
'       PARAM           : strRCPT_NO            受付番号(I)
'       RETURN          : 正常(True)のみ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUPD_RCPT_TRAN(ByVal strRCPT_NO As String) As Boolean
    Dim strSQL              As String
    Dim objRs               As Recordset
    Dim objFld              As Field
    On Error GoTo ErrorHandler

    strSQL = "SELECT " & Chr(13)
    strSQL = strSQL & " RCPT_KREN_NAME,RCPT_KREN_KANA," & Chr(13)
    strSQL = strSQL & " RCPT_KREN_YUBINO,RCPT_KREN_ADR_1,RCPT_KREN_ADR_2,RCPT_KREN_ADR_3," & Chr(13)
    strSQL = strSQL & " RCPT_KREN_TEL,RCPT_KREN_KEITAI,RCPT_KREN_FAX," & Chr(13)
    strSQL = strSQL & " RCPT_KREN_ZOKUGARA, RCPT_SLOPE, RCPT_BIKE_NUMBER, RCPT_BIKE_SHASYU " & Chr(13)
    strSQL = strSQL & "FROM RCPT_TRAN" & Chr(13)
    strSQL = strSQL & "WHERE RCPT_NO = '" & strRCPT_NO & "'"
    
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
    On Error GoTo ErrorHandler1
    With objRs
        If Not .EOF Then
            .Edit
            On Error GoTo ErrorHandler2
            For Each objFld In .Fields
            
                If objFld.NAME = "RCPT_SLOPE" Then
                    objFld.VALUE = Me.Controls("cmb_スロープ").VALUE
                Else
                    objFld.VALUE = Me.Controls("txt_" & objFld.NAME).VALUE
                End If
            Next
            
            .UPDATE
            On Error GoTo ErrorHandler1
        End If
        .Close
    End With
    fncUPD_RCPT_TRAN = True
Exit Function

ErrorHandler2:
    objRs.CancelUpdate
ErrorHandler1:
    objRs.Close
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "fncUPD_RCPT_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :subConnectDB
'        Parameter        :strBumoCD  部門コード
'        CREATE_DATE      :2005/08/31
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subConnectDB(ByVal strBumoCD As String)

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo subConnectDB_Err

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = fncGetDataSource("ODBC_DATA_SOURCE_NAME_" & strBumoCD)
    If strDataSource = "" Then
        ' テーブル[SETU_TABL]の設定不正
        Exit Sub
    End If

    ' DB接続
    Set objKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(CStr(strBumoCD)))
    ' -----------------------------------------------------------------------------------------------------------

    Exit Sub

subConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :subDisConnectDB
'        CREATE_DATE      :2005/08/31
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDisConnectDB()

    On Error GoTo subDisConnectDB_Err

    ' コンテナDB
    If Not objKONT_DB Is Nothing Then   'MODIFY 2009/04/01 KIHARA
        objKONT_DB.Close
        Set objKONT_DB = Nothing
    End If                              'MODIFY 2009/04/01 KIHARA
'INSERT 2009/04/01 KIHARA Start
    ' 加瀬DB
    If Not pobjKASE_DB Is Nothing Then
        pobjKASE_DB.Close
        Set pobjKASE_DB = Nothing
    End If
'INSERT 2009/04/01 KIHARA End
    Exit Sub

subDisConnectDB_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ODBC接続文字列取得
'        MODULE_ID        :fncGetDataSource
'        Parameter        :strODBCNAME ODBC名
'        CREATE_DATE      :2005/08/31
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
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function

' <<<その他>>>
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :SEL_USER_QUERY
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SEL_USER_QUERY(strKOKYC As Variant) As String

    Dim strSQL As String

    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM dbo_USER_MAST "
    strSQL = strSQL & "WHERE USER_CODE =" & Val(strKOKYC)
    strSQL = strSQL & ";"

    SEL_USER_QUERY = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :SEL_CARG_QUERY
'        Parameter        :strYARDC  ヤードコード
'                          strKONTC  コンテナ番号
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2005/08/03 REP T.SUZUKI Start
Private Function SEL_CARG_QUERY(strYardc As Variant, _
                                strKONTC As Variant) As String

' 2005/08/03 REP T.SUZUKI End

    Dim strSQL As String
    Dim strBumoCD As String
    
    Call subGetBumonName(pobjDB, strBumoCD)
    

    'INSERT 2019/08/08 N.IMAI Start
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " YOUKT_UKNO "
    strSQL = strSQL & ",YOUKT_UKDATE "
    strSQL = strSQL & ",YOUKT_UKTANTO "
    strSQL = strSQL & ",YOUKT_UKKBN "
    strSQL = strSQL & ",YOUKT_YCODE "
    strSQL = strSQL & ",YOUKT_NEARKBN "
    strSQL = strSQL & ",YOUKT_NO "
    strSQL = strSQL & ",YOUKT_USAGE "
    strSQL = strSQL & ",YOUKT_SIZE_FROM "
    strSQL = strSQL & ",YOUKT_SIZE_TO "
    strSQL = strSQL & ",YOUKT_STEP "
    strSQL = strSQL & ",YOUKT_PRICFROM "
    strSQL = strSQL & ",YOUKT_PRICTO "
    strSQL = strSQL & ",YOUKT_YUKBN "
    strSQL = strSQL & ",YOUKT_KUDATE "
    strSQL = strSQL & ",YOUKT_YTDATE "
    strSQL = strSQL & ",YOUKT_KKDATE "
    strSQL = strSQL & ",YOUKT_TKDATE "
    strSQL = strSQL & ",YOUKT_TADATE "
    strSQL = strSQL & ",YOUKT_MLDATE "
    strSQL = strSQL & ",YOUKT_BUTU "
    strSQL = strSQL & ",YOUKT_BIKO "
    strSQL = strSQL & ",YOUKT_KIBOUSU "
    strSQL = strSQL & ",YOUKT_UCODE "
    strSQL = strSQL & ",YOUKT_NAME "
    strSQL = strSQL & ",YOUKT_KANA "
    strSQL = strSQL & ",YOUKT_TANM "
    strSQL = strSQL & ",YOUKT_TAKA "
    strSQL = strSQL & ",YOUKT_YUBINO "
    strSQL = strSQL & ",YOUKT_ADR_1 "
    strSQL = strSQL & ",YOUKT_ADR_2 "
    strSQL = strSQL & ",YOUKT_ADR_3 "
    strSQL = strSQL & ",YOUKT_TEL "
    strSQL = strSQL & ",YOUKT_KEITAI "
    strSQL = strSQL & ",YOUKT_FAX "
    strSQL = strSQL & ",YOUKT_KKBN "
    strSQL = strSQL & ",YOUKT_RENKBN "
    strSQL = strSQL & ",YOUKT_RENJFROM "
    strSQL = strSQL & ",YOUKT_RENJTO "
    strSQL = strSQL & ",YOUKT_MAIL "
    strSQL = strSQL & ",YOUKT_AUTOKBN "
    strSQL = strSQL & ",YOUKT_MOVEKBN "
    strSQL = strSQL & ",YOUKT_MOTO_ACPTNO "
    strSQL = strSQL & ",YOUKT_RCPT_NO "
    strSQL = strSQL & ",YOUKT_REQ_YCODE "
    strSQL = strSQL & ",YOUKT_MOVE_TANTO "
    strSQL = strSQL & ",YOUKT_GENKBN "
    strSQL = strSQL & ",YOUKT_GENBNR "
    strSQL = strSQL & ",YOUKT_CAMPC "
    strSQL = strSQL & ",YOUKT_ORG_YUKBN "
    strSQL = strSQL & ",YOUKT_SEND_CD "
    strSQL = strSQL & ",YOUKT_USE_PERIOD "
    strSQL = strSQL & ",YOUKT_ZEN_KKDATE "
    strSQL = strSQL & ",YOUKT_UPDAD "
    strSQL = strSQL & ",YOUKT_UPDAJ "
    strSQL = strSQL & ",YOUKT_UPDPB "
    strSQL = strSQL & ",YOUKT_UPDUB "
    strSQL = strSQL & ",RCPT_NO "
    strSQL = strSQL & ",RCPT_CAMPC "
    strSQL = strSQL & ",RCPT_UKTANTO "
    strSQL = strSQL & ",RCPT_UKDATE "
    strSQL = strSQL & ",RCPT_CYDATE "
    strSQL = strSQL & ",RCPT_CREASON "
    strSQL = strSQL & ",RCPT_PAYDATE "
    strSQL = strSQL & ",RCPT_UKNOTE "
    strSQL = strSQL & ",RCPT_COMMENT_1 "
    strSQL = strSQL & ",RCPT_COMMENT_2 "
    strSQL = strSQL & ",RCPT_UCODE "
    strSQL = strSQL & ",RCPT_SEND_CD "
    strSQL = strSQL & ",RCPT_BUTU "
    strSQL = strSQL & ",RCPT_YCODE "
    strSQL = strSQL & ",RCPT_CNO "
    strSQL = strSQL & ",RCPT_HOSYICD "
    strSQL = strSQL & ",RCPT_HOSYO_CD "
    strSQL = strSQL & ",RCPT_NEW_HOSPLN_IKOU "
    strSQL = strSQL & ",RCPT_PLAN_CD "          '2020/12/15 [EGL] INS
    strSQL = strSQL & ",RCPT_SEV_EXPIRE "
    strSQL = strSQL & ",RCPT_SEV_EXPIRE_MONTH "
    strSQL = strSQL & ",RCPT_SEV1N "
    strSQL = strSQL & ",RCPT_SEV2N "
    strSQL = strSQL & ",RCPT_SEV3N "
    strSQL = strSQL & ",RCPT_ENDEN "
    strSQL = strSQL & ",RCPT_TPOINT_RITU "
    strSQL = strSQL & ",RCPT_MITU_CD "
    strSQL = strSQL & ",RCPT_SEIKYU_CD "
    strSQL = strSQL & ",RCPT_STDATE "
    strSQL = strSQL & ",RCPT_DC_STDATE "
    strSQL = strSQL & ",RCPT_DC_ENABLE "
    strSQL = strSQL & ",RCPT_SYOKISAN_DATE "
    strSQL = strSQL & ",RCPT_KISAN_DATE "
    strSQL = strSQL & ",RCPT_RENTKG "
    strSQL = strSQL & ",RCPT_RENTKG_OLR "
    strSQL = strSQL & ",RCPT_FIRSTKG "
    strSQL = strSQL & ",RCPT_EZAPPI_CODE "
    strSQL = strSQL & ",RCPT_EZAPPI "
    strSQL = strSQL & ",RCPT_EZAPPI_DAILY "
    strSQL = strSQL & ",RCPT_ADD_EZAPPI_CODE1 "
    strSQL = strSQL & ",RCPT_ADD_EZAPPI1 "
    strSQL = strSQL & ",RCPT_ADD_EZAPPI_DAILY1 "
    strSQL = strSQL & ",RCPT_ADD_EZAPPI_CODE2 "
    strSQL = strSQL & ",RCPT_ADD_EZAPPI2 "
    strSQL = strSQL & ",RCPT_ADD_EZAPPI_DAILY2 "
    strSQL = strSQL & ",RCPT_FZAPPI_CODE "
    strSQL = strSQL & ",RCPT_FZAPPI "
    strSQL = strSQL & ",RCPT_ADD_FZAPPI_CODE1 "
    strSQL = strSQL & ",RCPT_ADD_FZAPPI1 "
    strSQL = strSQL & ",RCPT_ADD_FZAPPI_CODE2 "
    strSQL = strSQL & ",RCPT_ADD_FZAPPI2 "
    strSQL = strSQL & ",RCPT_NET_WARIBIKI "
    strSQL = strSQL & ",RCPT_SECUKG "
    strSQL = strSQL & ",RCPT_SECUKG_KASE "
    strSQL = strSQL & ",RCPT_SECUKG_FLG "
    strSQL = strSQL & ",RCPT_HOSHOU_WARIBIKI "
    strSQL = strSQL & ",RCPT_SEIKYU_KIKAN "
    strSQL = strSQL & ",RCPT_SUMIKG "
    strSQL = strSQL & ",RCPT_KAISYUG "
    strSQL = strSQL & ",RCPT_KAGIICD "
    strSQL = strSQL & ",RCPT_KAGI_NO "
    strSQL = strSQL & ",RCPT_SECURITY_CARD "
    strSQL = strSQL & ",RCPT_KAGI_LENDNUM "
    strSQL = strSQL & ",RCPT_KAINO "
    strSQL = strSQL & ",RCPT_HOSYD "
    strSQL = strSQL & ",RCPT_HOSYB "
    strSQL = strSQL & ",RCPT_SEND_SAME_DIFFER "
    strSQL = strSQL & ",RCPT_SEND_NAME "
    strSQL = strSQL & ",RCPT_SEND_HONORIFIC "
    strSQL = strSQL & ",RCPT_SEND_YUBINO "
    strSQL = strSQL & ",RCPT_SEND_ADR_1 "
    strSQL = strSQL & ",RCPT_SEND_ADR_2 "
    strSQL = strSQL & ",RCPT_SEND_ADR_3 "
    strSQL = strSQL & ",RCPT_SEND_TEL "
    strSQL = strSQL & ",RCPT_KEY_SEND_DATE "
    strSQL = strSQL & ",RCPT_ARRIVAL_EXPECT "
    strSQL = strSQL & ",RCPT_SPECIFY_TIME "
    strSQL = strSQL & ",RCPT_SLIP_NO "
    strSQL = strSQL & ",RCPT_DERIVERY_STATUS "
    strSQL = strSQL & ",RCPT_ARRIVAL_DATE "
    strSQL = strSQL & ",RCPT_CARG_ACPTNO "
    strSQL = strSQL & ",RCPT_FWARIBIKI "
    strSQL = strSQL & ",RCPT_EWARIBIKI "
    strSQL = strSQL & ",RCPT_SEIKI "
    strSQL = strSQL & ",RCPT_KAIHI "
    strSQL = strSQL & ",RCPT_SLOPE "
    strSQL = strSQL & ",RCPT_BIKE_KOUKBN "
    strSQL = strSQL & ",RCPT_BIKE_NUMBER "
    strSQL = strSQL & ",RCPT_BIKE_SHASYU "
    strSQL = strSQL & ",RCPT_BIKE_HAIKIRYO "
    strSQL = strSQL & ",RCPT_TPOINT "
    strSQL = strSQL & ",RCPT_KOUKBN "
    strSQL = strSQL & ",RCPT_KREN_NAME "
    strSQL = strSQL & ",RCPT_KREN_KANA "
    strSQL = strSQL & ",RCPT_KREN_YUBINO "
    strSQL = strSQL & ",RCPT_KREN_ADR_1 "
    strSQL = strSQL & ",RCPT_KREN_ADR_2 "
    strSQL = strSQL & ",RCPT_KREN_ADR_3 "
    strSQL = strSQL & ",RCPT_KREN_TEL "
    strSQL = strSQL & ",RCPT_KREN_KEITAI "
    strSQL = strSQL & ",RCPT_KREN_FAX "
    strSQL = strSQL & ",RCPT_KREN_ZOKUGARA "
    strSQL = strSQL & ",RCPT_SHOUMEI_UP_UM "
    strSQL = strSQL & ",RCPT_SHOUMEI_KDATE "
    strSQL = strSQL & ",RCPT_SHOUMEI_FDATE "
    strSQL = strSQL & ",RCPT_SHOUMEI_FCNT "
    strSQL = strSQL & ",RCPT_NEN_KBN "
    strSQL = strSQL & ",RCPT_NEN_JIREI "
    strSQL = strSQL & ",RCPT_NEN_SEIKI "
    strSQL = strSQL & ",RCPT_NEN_BIKO "
    strSQL = strSQL & ",RCPT_MOUSIKOMI_DATE "
    strSQL = strSQL & ",RCPT_SEC_MAILSND_DATE "
    strSQL = strSQL & ",RCPT_YAMATO_DATE "
    strSQL = strSQL & ",RCPT_UPDAD "
    strSQL = strSQL & ",RCPT_UPDAJ "
    strSQL = strSQL & ",RCPT_UPDPB "
    strSQL = strSQL & ",RCPT_UPDUB "
    strSQL = strSQL & ",CARG_YCODE "
    strSQL = strSQL & ",CARG_NO "
    strSQL = strSQL & ",CARG_UCODE "
    strSQL = strSQL & ",CARG_AGRE "
    strSQL = strSQL & ",CARG_FSDATE "
    strSQL = strSQL & ",CARG_STDATE "
    strSQL = strSQL & ",CARG_EDDATE "
    strSQL = strSQL & ",CARG_CYDATE "
    strSQL = strSQL & ",CARG_KYDATE "
    strSQL = strSQL & ",CARG_CONTNA "
    strSQL = strSQL & ",CARG_USAGE "
    strSQL = strSQL & ",CARG_DOCU1 "
    strSQL = strSQL & ",CARG_DOCU2 "
    strSQL = strSQL & ",CARG_TUBOSU "
    strSQL = strSQL & ",CARG_TTANKA "
    strSQL = strSQL & ",CARG_RENTKG "
    strSQL = strSQL & ",CARG_SYOZEI "
    strSQL = strSQL & ",CARG_FRSTKG "
    strSQL = strSQL & ",CARG_FSYOZEI "
    strSQL = strSQL & ",CARG_FRST_BILL "
    strSQL = strSQL & ",CARG_PREMTH_SUM "
    strSQL = strSQL & ",CARG_SECUKG "
    strSQL = strSQL & ",CARG_BIKO "
    strSQL = strSQL & ",CARG_UPDATE "
    strSQL = strSQL & ",CARG_HDATE "
    strSQL = strSQL & ",CARG_MDATE "
    strSQL = strSQL & ",CARG_ACPTNO "
    strSQL = strSQL & ",CARG_CAMPC "
    strSQL = strSQL & ",CARG_HOSYICD "
    strSQL = strSQL & ",CARG_HOSYO_CD "
    strSQL = strSQL & ",CARG_KAGIICD "
    strSQL = strSQL & ",CARG_DAHIB "
    strSQL = strSQL & ",CARG_HOKEB "
    strSQL = strSQL & ",CARG_KAIJB "
    strSQL = strSQL & ",CARG_HOKAI "
    strSQL = strSQL & ",CARG_HOSYD "
    strSQL = strSQL & ",CARG_HOSYB "
    strSQL = strSQL & ",CARG_ADMEI "
    strSQL = strSQL & ",CARG_KEITI "
    strSQL = strSQL & ",CARG_KAGIA "
    strSQL = strSQL & ",CARG_UKNO "
    strSQL = strSQL & ",CARG_UKTANTO "
    strSQL = strSQL & ",CARG_KAIUKEDATE "
    strSQL = strSQL & ",CARG_KAITANTO "
    strSQL = strSQL & ",CARG_KAICD "
    strSQL = strSQL & ",CARG_KAIRIYUCD "
    strSQL = strSQL & ",CARG_CMP_EXDATE "
    strSQL = strSQL & ",CARG_KEY_RETDATE "
    strSQL = strSQL & ",CARG_SPAREKEY "
    strSQL = strSQL & ",CARG_MNT_REQDATE "
    strSQL = strSQL & ",CARG_MNT_CMPDATE "
    strSQL = strSQL & ",CARG_NIRCD "
    strSQL = strSQL & ",CARG_SYOCD "
    strSQL = strSQL & ",CARG_HOSCD "
    strSQL = strSQL & ",CARG_MNT_TANTO "
    strSQL = strSQL & ",CARG_MNT_MEMO "
    strSQL = strSQL & ",CARG_LIQ_REQDATE "
    strSQL = strSQL & ",CARG_LIQ_NOTE "
    strSQL = strSQL & ",CARG_LIQ_FLG "
    strSQL = strSQL & ",CARG_SPAREKEY_NOTE "
    strSQL = strSQL & ",CARG_KEY_LENTNUM "
    strSQL = strSQL & ",CARG_KEY_RETNUM "
    strSQL = strSQL & ",CARG_KEY_CHGDATE "
    strSQL = strSQL & ",CARG_KEY_CHGTANTO "
    strSQL = strSQL & ",CARG_KEY_CHGRIYUCD "
    strSQL = strSQL & ",CARG_KEY_CHGNOTE "
    strSQL = strSQL & ",CARG_SEIKI "
    strSQL = strSQL & ",CARG_UPDAD "
    strSQL = strSQL & ",CARG_UPDAJ "
    strSQL = strSQL & ",CARG_UPDPB "
    strSQL = strSQL & ",CARG_UPDUB "
    strSQL = strSQL & ",MTANT.TANTM_TANTN AS MTANTM_TANTN"
    strSQL = strSQL & ",JTANT.TANTM_TANTN AS JTANTM_TANTN"
    
    strSQL = strSQL & ",CODET_NAMEN "
    
    strSQL = strSQL & ",NAME61.NAME_NAME AS SPARE "
    strSQL = strSQL & ",NAME62.NAME_NAME AS SECUCOMP "
    
    
    
    'INSERT 2019/08/08 N.IMAI Wnd
'    strSQL = "SELECT * " & Chr(13)                                             'DELETE 2019/08/08 N.IMAI
    strSQL = strSQL & " FROM dbo_CARG_FILE, " & Chr(13)
    strSQL = strSQL & "      dbo_RCPT_TRAN, " & Chr(13)
    strSQL = strSQL & "      dbo_YOUK_TRAN  " & Chr(13)
    
    strSQL = strSQL & "      ,dbo_YARD_MAST  " & Chr(13)
    strSQL = strSQL & "      ,TANT_MAST MTANT " & Chr(13)
    strSQL = strSQL & "      ,TANT_MAST JTANT " & Chr(13)
    
    strSQL = strSQL & "      ,dbo_TINH_MAST " & Chr(13)
    strSQL = strSQL & "      ,dbo_BUKE_MAST " & Chr(13)
    strSQL = strSQL & "      ,CODE_TABL " & Chr(13)
    
    
    strSQL = strSQL & "     ,dbo_NAME_MAST NAME61 "
    strSQL = strSQL & "     ,dbo_NAME_MAST NAME62 "
    
    strSQL = strSQL & "WHERE CARG_YCODE = " & Val(strYardc) & Chr(13)
    strSQL = strSQL & "  AND CARG_NO    = " & Val(strKONTC) & Chr(13)
    strSQL = strSQL & "  AND CARG_ACPTNO = RCPT_CARG_ACPTNO " & Chr(13)
    '2017/02/10 M.HONDA INS
    strSQL = strSQL & "  AND YOUKT_UKNO  = RCPT_NO " & Chr(13)
    strSQL = strSQL & "  AND YOUKT_YUKBN = 20 " & Chr(13)
    '2017/02/10 M.HONDA INS
    strSQL = strSQL & "  AND CARG_YCODE = YARD_CODE "
    strSQL = strSQL & "  AND YARD_MNT_TANTO = MTANT.TANTM_TANTC " & Chr(13)
    strSQL = strSQL & "  AND MTANT.TANTM_BUMOC = '" & strBumoCD & "'"
    
    strSQL = strSQL & "  AND YARD_JUN_TANTO = JTANT.TANTM_TANTC " & Chr(13)
    strSQL = strSQL & "  AND JTANT.TANTM_BUMOC = '" & strBumoCD & "'"
    
    
    strSQL = strSQL & "  AND TINHM_TINTC = '" & strYardc & "'" & Chr(13)
    strSQL = strSQL & "  AND TINHM_BUMOC = '" & strBumoCD & "'" & Chr(13)
    strSQL = strSQL & "  AND TINHM_BUKHC = BUKEM_BUKHC " & Chr(13)
    
    strSQL = strSQL & "  AND CODET_SIKBC = '210' " & Chr(13)
    strSQL = strSQL & "  AND Val(BUKEM_SYOUI) = CODET_CODEC " & Chr(13)
    
    
    strSQL = strSQL & "  AND NAME61.NAME_ID = '061' " & Chr(13)
    strSQL = strSQL & "  AND NAME61.NAME_CODE = YARD_SPARE_KEY " & Chr(13)
    
    strSQL = strSQL & "  AND NAME62.NAME_ID = '062' " & Chr(13)
    'strSQL = strSQL & "  AND NAME62.NAME_CODE = Val(YARD_SECUCOMP) " & Chr(13) 'DELETE 2024/01/26 N.IMAI
    strSQL = strSQL & "  AND NAME62.NAME_CODE = NZ(YARD_SECUCOMP,2) " & Chr(13) 'DELETE 2024/01/26 N.IMAI
    
    strSQL = strSQL & "ORDER BY CARG_FSDATE DESC "
    strSQL = strSQL & ";"


    SEL_CARG_QUERY = strSQL
End Function

'==============================================================================*
'
'        MODULE_NAME      :移動元契約データの初期表示用の「備考」の文言または
'                          移動元契約データ設定用の「メンテメモ」の文言作成
'        MODULE_ID        :subCreateMemo
'        Parameter        :
'                          intMode(IN)       = 処理区分(0:備考作成／1:メンテメモ作成)
'                          strMoveKbn(IN)    = 移動区分
'                          strMenteMemo(OUT) = メンテメモ文言格納
'        CREATE_DATE      :2006/01/05
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreateMemo(ByVal intMode As Integer, ByVal strMOVEKBN As String, _
                          ByRef strMenteMemo As String)

    Dim strMemo       As String
    Dim strName       As String
    Dim strYcode      As String
    Dim strNO         As String
    Dim strACPTNO     As String
    Dim strUCODE      As String
    Dim strUKANA      As String
    Dim strMoney      As String
    Dim strRENTKG     As String
    Dim strHOSYICD    As String
    Dim strRyak       As String
    Dim intLenb       As Integer
    Dim strFromTo     As String

    strMenteMemo = ""
    strMemo = ""
    If Nz(strMOVEKBN, "") = "" Then Exit Sub

    ' 初期化
    strYcode = ""
    strNO = ""
    strACPTNO = ""
    strUCODE = ""
    strUKANA = ""
    strMoney = ""
    strRENTKG = ""
    strHOSYICD = ""
    strFromTo = ""

    ' 移動区分名取得
    strName = ""
    strName = fncSelectName(P_NAME_098, , strMOVEKBN)
    If Nz(strName, "") = "" Then Exit Sub

    ' 画面のタイトル設定
    Me.lbl_Keiyaku.Caption = P_TITTLE_KEIYAKU & "(" & strName & ")"

    Select Case intMode
        Case 0  ' 備考作成
            With MOTO_CARG_FILE
                strYcode = Format(.strYcode, "000000")             ' 旧契約ヤードコード
                strNO = Format(.strNO, "000000")                   ' 旧契約コンテナ番号
                strACPTNO = .strACPTNO                             ' 旧契約番号
                strUCODE = Format(.strUCODE, "000000")             ' 旧顧客コード
                strRENTKG = Format(.dblRENTKG, "#,##0")            ' 旧月額料金
                strHOSYICD = .strHOSYICD                           ' 旧保証区分
            End With
            strFromTo = "から"

        Case 1  ' メンテメモ作成
            With Me
                strYcode = Format(.cmb_ヤードコード, "000000")     ' 新契約ヤードコード
                strNO = Format(.cmb_コンテナ番号, "000000")        ' 新契約コンテナ番号
                strACPTNO = .txt_受注契約番号                      ' 新契約番号
                strUCODE = Format(.cmb_顧客コード, "000000")       ' 新顧客コード
                strRENTKG = Format(.txt_月額料金, "#,##0")         ' 新月額料金
                strHOSYICD = .cmb_保証区分                         ' 新保証区分
            End With

            strFromTo = "に"
    End Select

    Select Case strMOVEKBN
        ' *** 自社理由移動 OR 顧客理由移動 OR 未使用移動
        Case P_MOVEKBN_01, P_MOVEKBN_02, P_MOVEKBN_03
            ' 備考 OR メンテメモ設定
            ' (新(旧)契約ヤードコード + "-" + 新(旧)契約コンテナ番号 + 移動理由名 + YY/MM)
            strMemo = strYcode & "-" & strNO & strFromTo & strName

        Case P_MOVEKBN_04  ' *** 名義変更 ***
            ' メンテメモ設定
            ' 顧客名取得
            strUKANA = Nz(DLookup("USER_KANA", "dbo_USER_MAST", "USER_CODE = " & strUCODE), "")

            ' 新(旧)顧客名カナを除いた文字数を取得
            intLenb = CmLenb(strACPTNO & "「" & strUCODE & ":" & "様」" & strFromTo & strName) + 6
            intLenb = pcintBikoMaxLength - intLenb

            ' 備考の最大文字数にあわせて新(旧)顧客名カナを編集
            strUKANA = Mid(strUKANA, 1, intLenb)

            ' (新(旧)契約番号 + "「" + 新(旧)顧客コード + ":" & 新(旧)顧客名カナ + "様に" + 移動理由名 + YY/MM)
            strMemo = strACPTNO & "「" & strUCODE & ":" & strUKANA & "様」" & strFromTo & strName

'        Case P_MOVEKBN_05  ' *** 金額変更 ***  'DELETE 2011/09/27 M.RYU
        Case P_MOVEKBN_05, P_MOVEKBN_07         'INSERT 2011/09/27 M.RYU
            ' メンテメモ設定
            ' (新(旧)契約番号 + "「\" + 新(旧)契約月額 + "」に" + 移動理由名 + YY/MM)
            strMemo = strACPTNO & "「\" & Format(strRENTKG, "#,##0") & "」" & strFromTo & strName

        Case P_MOVEKBN_06  ' *** 保障会社変更 ***
            ' 保証区分略称取得
            strRyak = ""
            strRyak = fncSelectRyak("200", , strHOSYICD)

            ' メンテメモ設定
            ' (新(旧)契約番号 + "「" + 保証区分略称 + "」に" + 移動理由名 + YY/MM)
            strMemo = strACPTNO & "「" & strRyak & "」" & strFromTo & strName
    End Select

' 2006/02/02 REP T.SUZUKI Start
    ' 年月付加
    If intMode = 1 Then
        ' メンテメモ作成時に、画面の新契約開始日の年月を設定する
        ' (※備考作成時は不要！！！)
        strMemo = strMemo & Space(1) & Format(Me.txt_契約開始日, "YY/MM")
    End If
' 2006/02/02 REP T.SUZUKI End

    strMenteMemo = strMemo
End Sub
'==============================================================================*
'
'        MODULE_NAME      :開始制限日の取得
'        MODULE_ID        :fncGetLimitationStartDate
'        引       数      ：なし
'        CREATE_DATE      :2006/01/31
'
'==============================================================================*
Private Function fncGetLimitationStartDate() As Boolean

    Dim strBumoCD As String
    Dim isRet     As Boolean
    Dim strErrMsg As String

    On Error GoTo Exception

    ' 部門コード／部門名取得
    Call subGetBumonName(pobjDB, strBumoCD)
    ' コンテナDB接続
    Call subConnectDB(strBumoCD)

    ' 開始制限日の取得う
    isRet = CmKomsMod.GetFormerContractDate(objKONT_DB, _
                                    Me.cmb_ヤードコード, _
                                    Me.cmb_コンテナ番号, _
                                    LimitationStartDate, _
                                    strErrMsg)
    ' コンテナDB切断
    Call subDisConnectDB

    If isRet = False Then
        MsgBox strErrMsg, vbInformation, strMsg_err
    End If
    fncGetLimitationStartDate = isRet

    Exit Function

Exception:
    If Not objKONT_DB Is Nothing Then
        ' コンテナDB切断
        Call subDisConnectDB
    End If
    
    MsgBox "ｴﾗｰｺｰﾄﾞ:" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ:" & Err.Description
    Err.Clear
    fncGetLimitationStartDate = False
End Function
'==============================================================================*
'
'        MODULE_NAME      :その他のテーブルからの情報取得
'        MODULE_ID        :fncGetOtherTableData
'        引       数      :予約受付番号、受注契約番号
'        戻り値           :受付トランの有無
'        CREATE_DATE      :2006/05/24
'
'==============================================================================*
Private Function fncGetOtherTableData(ByVal a予約受付番号 As String, _
                                      ByVal a受注契約番号 As String _
                                      ) As Boolean

    Dim strBumoCD As String
    Dim strSQL    As String
    Dim rsObject  As Recordset
    
    Dim isException As Boolean
    
    On Error GoTo Exception
    fncGetOtherTableData = False
    isException = False

    ' 部門コード／部門名取得
    Call subGetBumonName(pobjDB, strBumoCD)
    ' コンテナDB接続
    Call subConnectDB(strBumoCD)

    '①予約受付番号を元に、受付トランの有無と適用キャンペーン内容を取得する
    ' ※古い契約は予約も受付もないのでチェックしてから取得準備をなせ
    If a予約受付番号 <> "" Then
        strSQL = " SELECT " & Chr(13)
        strSQL = strSQL & " RCPT_KREN_NAME,RCPT_KREN_KANA,RCPT_KREN_ZOKUGARA," & Chr(13)    'INSERT START 2011/04/11 K.ISHIZAKA
        strSQL = strSQL & " RCPT_KREN_YUBINO,RCPT_KREN_ADR_1,RCPT_KREN_ADR_2,RCPT_KREN_ADR_3," & Chr(13)
        strSQL = strSQL & " RCPT_KREN_TEL,RCPT_KREN_KEITAI,RCPT_KREN_FAX," & Chr(13)        'INSERT END   2011/04/11 K.ISHIZAKA
        strSQL = strSQL & " RCPT_SEV1N,RCPT_SEV2N,RCPT_SEV3N,RCPT_ENDEN ,RCPT_SLOPE " & Chr(13)
        strSQL = strSQL & " FROM RCPT_TRAN WHERE RCPT_NO = '" & a予約受付番号 & "'"
    
        Set rsObject = objKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    
        If rsObject.EOF = False Then
            pstrCampaignInfomationText = CmKomsMod.GetCampaignInfomationText( _
                                                    Nz(rsObject.Fields("RCPT_SEV1N"), ""), _
                                                    Nz(rsObject.Fields("RCPT_SEV2N"), ""), _
                                                    Nz(rsObject.Fields("RCPT_SEV3N"), ""), _
                                                    Nz(rsObject.Fields("RCPT_ENDEN"), ""), _
                                                    Me.txt_キャンペーン適用満了日 _
                                                    )
            pstrCampaignWarningText = CmKomsMod.GetCampaignWarningText( _
                                                    Nz(rsObject.Fields("RCPT_SEV1N"), ""), _
                                                    Nz(rsObject.Fields("RCPT_SEV2N"), ""), _
                                                    Nz(rsObject.Fields("RCPT_SEV3N"), ""), _
                                                    Nz(rsObject.Fields("RCPT_ENDEN"), ""), _
                                                    Me.txt_キャンペーン適用満了日 _
                                                    )
            
            txt_RCPT_KREN_NAME.VALUE = Nz(rsObject.Fields("RCPT_KREN_NAME"), "")    'INSERT START 2011/04/11 K.ISHIZAKA
            txt_RCPT_KREN_KANA.VALUE = Nz(rsObject.Fields("RCPT_KREN_KANA"), "")
            txt_RCPT_KREN_ZOKUGARA.VALUE = Nz(rsObject.Fields("RCPT_KREN_ZOKUGARA"), "")
            txt_RCPT_KREN_YUBINO.VALUE = Nz(rsObject.Fields("RCPT_KREN_YUBINO"), "")
            txt_RCPT_KREN_ADR_1.VALUE = Nz(rsObject.Fields("RCPT_KREN_ADR_1"), "")
            txt_RCPT_KREN_ADR_2.VALUE = Nz(rsObject.Fields("RCPT_KREN_ADR_2"), "")
            txt_RCPT_KREN_ADR_3.VALUE = Nz(rsObject.Fields("RCPT_KREN_ADR_3"), "")
            txt_RCPT_KREN_TEL.VALUE = Nz(rsObject.Fields("RCPT_KREN_TEL"), "")
            txt_RCPT_KREN_KEITAI.VALUE = Nz(rsObject.Fields("RCPT_KREN_KEITAI"), "")
            txt_RCPT_KREN_FAX.VALUE = Nz(rsObject.Fields("RCPT_KREN_FAX"), "")      'INSERT END   2011/04/11 K.ISHIZAKA
            
            cmb_スロープ.VALUE = Nz(rsObject.Fields("RCPT_SLOPE"), "0"): Call cmb_スロープ_AfterUpdate      'INSERT 20120411 M.HONDA
            
            fncGetOtherTableData = True
        End If
    End If
    
    '②受注契約番号を元に、セキュリティカードの表示文言を取得し設定す
    ' ※受注契約番号は無いって事は絶対ないのでチェックはせんよ
    Dim intBaseNum As Integer
    Dim intAddNum  As Integer
    Dim strCards   As String
   
    Call CmKomsMod.GetContractSCards(objKONT_DB, a受注契約番号, intBaseNum, intAddNum, strCards)
    Call SetCardInfomation(intBaseNum, intAddNum, strCards)
        
    GoTo Finally
    
Exception:
    isException = True

Finally:
    If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
    If Not objKONT_DB Is Nothing Then Call subDisConnectDB
    
    If isException = True Then
        Call Err.Raise(Err.Number, "fncGetOtherTableData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'■□■□■□■□■□■□■□■□■□ ↓↓↓　Public 関数群 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□
'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :subReload
'        CREATE_DATE      :2005/08/03
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subReload(intFlg As Integer)

' 2006/03/04 ADD T.SUZUKI Start
    ' 顧客マスタ保守画面へ遷移した際の対応
    If pblnFKS430_Flag = True Then
        If pblnChangeFlg = True Then
            ' 顧客マスタ登録画面での変更があった場合の処理
            If CHK_USER_MAST(1) = False Then Exit Sub
            pblnChangeFlg = False
        End If

        pblnFKS430_Flag = False: Exit Sub
    End If
' 2006/03/04 ADD T.SUZUKI End

    Call subFormLockControl(True)

    If intFlg = 1 Then
        ' 解約入力画面で修正があった場合、契約入力画面を再表示する
        Call cmb_コンテナ番号_AfterUpdate
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      :カード情報の設定
'        MODULE_ID        :SetCardInfomation
'        CREATE_DATE      :2006/06/01
'        NOTE             :このメソッドは外部と内部両方から呼ばれる
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetCardInfomation( _
                              ByVal a基本枚数 As Integer, _
                              ByVal a追加枚数 As Integer, _
                              ByVal aカード達 As String _
                            )
    Dim textValue  As String
    Dim totalSum   As Integer
    Dim idx        As Integer
    
    totalSum = a基本枚数 + a追加枚数
 
    If totalSum = 0 Then
      ' 本契約に該当するカードが無ければ
        textValue = "貸出中のカードはありません"
    Else
      ' 該当するカードがあればどのように貸しているのか文言生成
        idx = InStr(aカード達, ";")
        If idx > 0 Then
            textValue = Left$(aカード達, InStr(aカード達, ";") - 1)
        Else
            textValue = aカード達
        End If
        
        If totalSum = 1 Then
            textValue = textValue & "...1枚貸出"
        Else
            textValue = textValue & "...他" & CStr(totalSum - 1)
            textValue = textValue & "枚、合計" & CStr(totalSum) & "枚貸出"
        End If
        
    End If
    Me.txt_セキュリティカード枚数 = textValue
End Sub
'==============================================================================*
'
'        MODULE_NAME      :修繕情報の再表示
'        MODULE_ID        :SubReloadfromFVS90n
'        CREATE_DATE      :2006/09/26
'        NOTE             :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SubReloadfromFVS90n()
    Call DispReprInfo
End Sub
'ADD 2009/04/01 KIHARA
'==============================================================================*
'
'        MODULE_NAME      :KASE-DB接続
'        MODULE_ID        :fncConnectKaseDB
'        CREATE_DATE      :2009/04/01
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectKaseDB() As Boolean

    Dim strDataSource As String

    On Error GoTo Exception

    fncConnectKaseDB = False

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"), "")

    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, pstrFVS601
        Set pobjKASE_DB = Nothing
        Exit Function
    Else
        Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
    End If
    ' -----------------------------------------------------------------------------------------------------------

    fncConnectKaseDB = True

    Exit Function

Exception:
    MsgBox "ｴﾗｰｺｰﾄﾞ" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ" & Err.Description
    Err.Clear
End Function
'ADD 2009/04/01 KIHARA
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :cmb_保証会社コード_AfterUpdate
'        CREATE_DATE      :2009/04/01
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_保証会社コード_AfterUpdate()
On Error GoTo ErrRtn

    Dim vCode    As Variant     '保証会社コード
    Dim strBumoCD   As String   '部門コード
    Dim strSQL      As String
    Dim objRs       As Recordset
    Dim blnError    As Boolean
    
    blnError = False
    vCode = Nz(Me.cmb_保証会社コード, "")

    If vCode = "" Then
        Me.txt_保証会社コード = ""
        Exit Sub
    End If

    ' 部門コード／部門名取得
    Call subGetBumonName(pobjDB, strBumoCD)

    '加瀬DB接続
    If fncConnectKaseDB = False Then Exit Sub

    strSQL = ""
    strSQL = strSQL & "SELECT TOP 1 SHIRM_SHIRN" & Chr(13)
    strSQL = strSQL & "  FROM SHIR_MAST" & Chr(13)
    strSQL = strSQL & " WHERE SHIRM_SHIRC = '" & vCode & "'"
    If Me.cmb_保証区分 = 1 Then
        strSQL = strSQL & Chr(13) & "   AND SHIRM_BUMOC = '" & strBumoCD & "'"
    End If
    
    Set objRs = pobjKASE_DB.OpenRecordset(strSQL, dbOpenDynaset, dbSQLPassThrough, dbReadOnly)
    
    If Not (objRs.BOF And objRs.EOF) Then
        Me.txt_保証会社コード = objRs.Fields("SHIRM_SHIRN")
    Else
        Me.txt_保証会社コード = ""
    End If
    
    objRs.Close
    Set objRs = Nothing
    pobjKASE_DB.Close
    Set pobjKASE_DB = Nothing

   'DB切断
    Call subDisConnectDB
    GoTo ExitRtn
       
ErrRtn:
    blnError = True
    
ExitRtn:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKASE_DB Is Nothing Then Call subDisConnectDB
   
    If blnError Then
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
   
End Sub
'2019/04/10 ADD TAKENOUCHI Start
'==============================================================================*
'
'        MODULE_NAME     : 申込トランの存在チェック処理
'        MODULE_ID       : fncChkExistEnterTrn
'        CREATE_DATE     : 2019/04/10
'        PARAM           : なし
'        RETURN          : True(正常) or False(異常)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkExistEnterTrn() As Boolean
    
    Dim objRow              As Object
    Dim strErrMsg           As String
    Dim strType             As String
    Dim strKaiYtiYmd        As String
    Dim strShinNitiji       As String
    Dim strShinNiti         As String
    Dim strShinJi           As String
    Dim strOpenArgs         As String
    
    On Error GoTo ErrorHandler

    fncChkExistEnterTrn = False
    
    Me.txt_申込種別.VALUE = ""
    
    'ログインユーザー取得
    pstrLoginUser = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & PROG_ID & "' AND INTIF_RECFB = 'KOMS2020_ID'"))
    If pstrLoginUser = "" Then
        Call MSZZ024_M10("DLookup", "INTI_FILEの設定不足です。")
    End If
    
    'パスワード取得
    pstrLoginPass = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = '" & PROG_ID & "' AND INTIF_RECFB = 'KOMS2020_PW'"))
    If pstrLoginPass = "" Then
        Call MSZZ024_M10("DLookup", "INTI_FILEの設定不足です。")
    End If
    
    'Config.iniを取得
    Set piniFile = MSZZ073K.get_ini_file()
    
    'キー取得
    pstrLoginKey = koms_login(pstrLoginUser, pstrLoginPass, piniFile)

    '処理中の契約Noを検索キーにして申込トランを検索する
    Set objRow = sel_ENTER_TRAN_row(piniFile, pstrLoginUser, pstrLoginKey, Me.txt_受注契約番号)

    If Not objRow Is Nothing Then
        strType = getNodeText(objRow, "TYPE")                                '申込種別
        pstrKaiYtiYmd = getNodeText(objRow, "VAR_02")                        '解約予定日
        strKaiYtiYmd = Format(pstrKaiYtiYmd, "@@@@/@@/@@")
        strShinNiti = Format(getNodeText(objRow, "VAR_03"), "@@@@/@@/@@") '申込日時
        strShinJi = Format(getNodeText(objRow, "VAR_04"), "@@:@@:@@") '申込日時
        
        strOpenArgs = strType & "," & pintAGRE & "," & Me.txt_受注契約番号 & "," & pstrLoginUser & "," & pstrLoginPass _
                        & "," & strKaiYtiYmd & "," & strShinNiti & "," & strShinJi & "," & Me.Form.NAME
        doCmd.OpenForm "FVS610", , , , , , strOpenArgs
    End If

    fncChkExistEnterTrn = True
    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncChkExistEnterTrn" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
   
End Function

'==============================================================================*
'
'        MODULE_NAME     : ENTER_TRAN取得
'        MODULE_ID       : sel_ENTER_TRAN_row
'        CREATE_DATE     : 2019/04/10
'        PARAM           : iniFile               接続情報等(I)
'                        : strLoginUser          ログインユーザー(I)
'                        : strLoginKey           ログインキー(I)
'                        : strKEIYB              連動番号(I)
'        RETURN          : 行ノード(MSXML2.IXMLDOMElement)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function sel_ENTER_TRAN_row( _
            iniFile As Collection _
    , ByVal strLoginUser As String _
    , ByVal strLoginKey As String _
    , ByVal strKEIYAKU_NO As String _
) As Object
    
    Dim strResponseXml      As String
    Dim objDoc              As Object
    Dim objTbl              As Object
    Dim strFile             As String
    
    On Error GoTo ErrorHandler
    
    Set objDoc = koms_createDocument(strLoginUser, strLoginKey)
    Set objTbl = koms_createTable(objDoc, "ENTER_TRAN", "SELECT", 1)
    
    Call koms_whereColumn(objTbl, "VAR_01", strKEIYAKU_NO, "varchar")
    Call koms_whereColumn(objTbl, "STATUS", 0, "tinyint")
    
    If Not koms_table_ctrl(objDoc, strResponseXml, iniFile) Then
        Call MSZZ003_M80("次のコマンドに失敗しました。")
        Call MSZZ003_M80(objDoc.XML)
        
        If InStr(1, LCase(strResponseXml), LCase("charset=EUC-JP"), vbTextCompare) > 0 Then
            strFile = CurrentProject.path & "\" & PROG_ID & "_" & Format(Now, "yyyymmdd_hhnnss") & "_ErrorLog.html"
            Call EUCSave(strFile, strResponseXml)
            strResponseXml = "結果は外部ファイル[" & strFile & "]を参照してください。"
        End If
        
        Call MSZZ003_M80(strResponseXml)
        Call MSZZ024_M10("koms_table_ctrl", "KOMS2020側のCORE_TRANの取得に失敗しました。")
    End If
    
    Set objTbl = Nothing
    Set objDoc = Nothing
    Set objDoc = XmlCreateDocument(strResponseXml)
    Set sel_ENTER_TRAN_row = objDoc.selectSingleNode("/root/table/row")

Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number, "sel_ENTER_TRAN_row" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME     : ネット解約DB関連の更新メイン処理
'        MODULE_ID       : fncUpdNetDBMain
'        CREATE_DATE     : 2019/04/10
'        PARAM           : なし
'        RETURN          : True(正常) or False(異常)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdNetDBMain() As Boolean

    Const STS_NOT_SEND_MAIL     As Integer = 2
    
    On Error GoTo ErrorHandler

    fncUpdNetDBMain = False
    
    '申込トランの更新
    If fncUpdEnterTrnMain(Me.txt_受注契約番号, STS_NOT_SEND_MAIL) = False Then
        GoTo ErrorHandler
    End If
    
    'KOMS契約情報の更新
    If fncUpdKomsCargTranMain(Me.txt_受注契約番号, Me.txt_申込種別) = False Then
        GoTo ErrorHandler
    End If
    
    fncUpdNetDBMain = True
    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncUpdNetDBMain" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
   
End Function

'==============================================================================*
'
'       MODULE_NAME     : 申込トランの更新メイン処理
'       MODULE_ID       : fncUpdEnterTrnMain
'       CREATE_DATE     : 2019/04/10
'       PARAM           : strKeiyakuNo  契約No
'                         intStatus     処理Status
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdEnterTrnMain(ByVal strKeiyakuNo As String, ByVal intStatus As Integer) As Boolean

    Dim objDoc              As Object
    Dim strErrMsg           As String
    Dim strLoginKey         As String
    Dim iniFile             As Collection
    Dim strUser             As String
    Dim strDate             As String
    Dim strTime             As String
    
    fncUpdEnterTrnMain = False
    
    On Error GoTo ErrorHandler
    
    '更新日時
    strUser = LsGetUserName()
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    
    Set objDoc = koms_createDocument(pstrLoginUser, pstrLoginKey)
    
    '申込トランの更新準備
    Call upd_ENTER_TRAN_Documents(objDoc, strKeiyakuNo, intStatus, strUser, strDate, strTime)
    
    '申込トランの更新実行
    Call send_table(piniFile, objDoc, "ENTER_TRAN")
    
    fncUpdEnterTrnMain = True
    Exit Function
    
ErrorHandler:
    strErrMsg = Err.Description
    Call MSZZ024_M00("fncUpdEnterTrnMain", False)
    Call MSZZ003_M00(PROG_ID, "9", strErrMsg)

End Function

'==============================================================================*
'
'       MODULE_NAME     : ENTER_TRAN更新
'       MODULE_ID       : upd_ENTER_TRAN_Documents
'       CREATE_DATE     : 2019/04/10
'       PARAM           : objDoc                ドキュメントオブジェクト(I)
'                       : strUser               契約No
'                       : intStatus             処理Status
'                       : strUser               ユーザー(I)
'                       : strDate               日付(I)
'                       : strTime               時刻(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub upd_ENTER_TRAN_Documents( _
          objDoc As Object, _
    ByVal strKeiyakuNo As String, _
    ByVal intStatus As Integer, _
    ByVal strUser As String, _
    ByVal strDate As String, _
    ByVal strTime As String _
)
    
    Dim objTbl      As Object
    
    On Error GoTo ErrorHandler

    'アクションメソッド設定
    Set objTbl = koms_createTable(objDoc, "ENTER_TRAN", "UPDATE")
    
    Call koms_setColumn(objTbl, "STATUS", intStatus, "tinyint")
    Call koms_setColumn(objTbl, "UPDATE_DATE", strDate, "varchar")
    Call koms_setColumn(objTbl, "UPDATE_TIME", strTime, "varchar")
    Call koms_setColumn(objTbl, "UPDATE_PRG_ID", "FVS600", "varchar")
    Call koms_setColumn(objTbl, "UPDATE_USER_ID", strUser, "varchar")
    Call koms_whereColumn(objTbl, "VAR_01", strKeiyakuNo, "varchar")

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "upd_ENTER_TRAN_Documents" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'       MODULE_NAME     : tableの更新
'       MODULE_ID       : send_table
'       CREATE_DATE     : 2019/04/10
'       PARAM           : iniFile               接続情報等(I)
'                       : objDoc                ドキュメントオブジェクト(I)
'                       : strTbl                テーブル名
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub send_table(iniFile As Collection, objDoc As Object, strTbl As String)
    
    Dim strResponseXml      As String
    Dim objRet              As Object
    Dim objTbl              As Object
    Dim strFile             As String
    
    On Error GoTo ErrorHandler
    
    If koms_table_ctrl(objDoc, strResponseXml, iniFile) Then
        Set objRet = XmlCreateDocument(strResponseXml)
        Set objTbl = objRet.selectSingleNode("/root/" & strTbl)
        
        If CInt(objTbl.Text) > 0 Then
            '正常
            Exit Sub
        End If
    End If
    
    Call MSZZ003_M80("次のコマンドに失敗しました。")
    Call MSZZ003_M80(objDoc.XML)
                                                                                
    If InStr(1, LCase(strResponseXml), LCase("charset=EUC-JP"), vbTextCompare) > 0 Then
        strFile = CurrentProject.path & "\" & PROG_ID & "_" & Format(Now, "yyyymmdd_hhnnss") & "_ErrorLog.html"
        Call EUCSave(strFile, strResponseXml)
        strResponseXml = "結果は外部ファイル[" & strFile & "]を参照してください。"
    End If
                                                                                
    Call MSZZ003_M80(strResponseXml)
    Call MSZZ024_M10("koms_table_ctrl", "PHP側の" & strTbl & "の更新に失敗しました。")

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "send_table" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'       MODULE_NAME     : KOMS契約情報の更新メイン処理
'       MODULE_ID       : fncUpdKomsCargTranMain
'       CREATE_DATE     : 2019/04/10
'       PARAM           : strKeiyakuNo  契約No
'                         strType       申込種別
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdKomsCargTranMain(ByVal strKeiyakuNo As String, ByVal strType As String) As Boolean

    Dim objDoc              As Object
    Dim strErrMsg           As String
    Dim strUser             As String
    Dim strDate             As String
    Dim strTime             As String
    
    fncUpdKomsCargTranMain = False
    
    On Error GoTo ErrorHandler
    
    '更新日時
    strUser = LsGetUserName()
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    
    Set objDoc = koms_createDocument(pstrLoginUser, pstrLoginKey)
    
    'KOMS契約情報の更新準備
    Call upd_KOMS_CARG_TRAN_Documents(objDoc, strKeiyakuNo, strType, strUser, strDate, strTime)
    
    'KOMS契約情報の更新実行
    Call send_table(piniFile, objDoc, "KOMS_CARG_TRAN")
    
    fncUpdKomsCargTranMain = True
    Exit Function
    
ErrorHandler:
    strErrMsg = Err.Description
    Call MSZZ024_M00("fncUpdEnterTrnMain", False)
    Call MSZZ003_M00(PROG_ID, "9", strErrMsg)

End Function

'2019/04/23 ADD TAKENOUCHI Start
'==============================================================================*
'
'       MODULE_NAME     : KOMS_CARG_TRAN更新
'       MODULE_ID       : upd_KOMS_CARG_TRAN_Documents
'       CREATE_DATE     : 2019/04/23
'       PARAM           : objDoc                ドキュメントオブジェクト(I)
'                       : strUser               契約No
'                       : strType               申込種別
'                       : strUser               ユーザー(I)
'                       : strDate               日付(I)
'                       : strTime               時刻(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub upd_KOMS_CARG_TRAN_Documents( _
          objDoc As Object, _
    ByVal strKeiyakuNo As String, _
    ByVal strType As String, _
    ByVal strUser As String, _
    ByVal strDate As String, _
    ByVal strTime As String _
)
    
    Dim objTbl          As Object
    Dim strKykSts       As String
    
    Const TYPE_01       As String = "01"         '解約申込
    Const KYK_STS_1     As String = "1"          '契約
    Const KYK_STS_2     As String = "2"          '解約
'    Const KAI_YTI_YMD   As Variant = Null
    
    On Error GoTo ErrorHandler

    'アクションメソッド設定
    Set objTbl = koms_createTable(objDoc, "KOMS_CARG_TRAN", "UPDATE")
    
    If strType = TYPE_01 Then
        strKykSts = KYK_STS_2
        Call koms_setColumn(objTbl, "KAIYAKU_YOTEI_YMD", pstrKaiYtiYmd, "char")
    Else
        strKykSts = KYK_STS_1
        Call koms_setColumn(objTbl, "KAIYAKU_YOTEI_YMD", "", "char")
    End If
    
    Call koms_setColumn(objTbl, "KEIYAKU_STS", strKykSts, "char")
    Call koms_setColumn(objTbl, "UPDATE_DATE", strDate, "varchar")
    Call koms_setColumn(objTbl, "UPDATE_TIME", strTime, "varchar")
    Call koms_setColumn(objTbl, "UPDATE_PRG_ID", "FVS600", "varchar")
    Call koms_setColumn(objTbl, "UPDATE_USER_ID", strUser, "varchar")
    Call koms_whereColumn(objTbl, "KEIYAKU_NO", strKeiyakuNo, "char")

Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "upd_KOMS_CARG_TRAN_Documents" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub
'==============================================================================*
'
'       MODULE_NAME     : 顧客問合わせトラン登録処理
'       MODULE_ID       : fncInsKOTO_TRAN
'       CREATE_DATE     : 2019/04/23
'       PARAM           : strKeiyakuNo  契約No
'       RETURN          : 正常(true)／エラー(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInsKOTO_TRAN(ByVal strKeiyakuNo As String) As Boolean
    '
    Dim dbKagtos                As Database
    Dim strKagtosSqlserver      As String
    Dim strKagtosConnect        As String
    Dim pstKotoTran             As CmReprMod.Type_KOTO_TRAN
    Dim objRsCarg As Recordset
    
    Dim strUser                 As String
    Dim strDate                 As String
    Dim strTime                 As String
    
    Dim strBumonCd              As String       '部門コード
    Dim strUCODE                As String       '顧客コード
    Dim strKotoMsg              As String       'メッセージ( 内容、内容１、内容２ )
    
    Dim strErrMsg               As String
    
    Dim strKYNO As String
    Dim strKYCD As String
    
    fncInsKOTO_TRAN = False
    
    On Error GoTo ErrorHandler
    
    '更新日時
    strUser = LsGetUserName()
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    
    '部門コード取得
    Call subGetBumonName(pobjDB, strBumonCd)
    
    ' 移動元コンテナ契約ファイルデータ取得
    If Not fncGetCARG_FILE(strKeiyakuNo, objRsCarg) Then GoTo ErrorHandler

    strUCODE = Format(objRsCarg.Fields("CARG_UCODE"), "000000")     '顧客コード取得
    objRsCarg.Close
    Set objRsCarg = Nothing
    
    'メッセージの作成
    If Me.txt_申込種別 = "01" Then
    '解約申込ならば
        strKotoMsg = "ネットからの解約申込(契約番号:" & Me.txt_受注契約番号 & ")について契約入力で手動取込しました。"
    ElseIf Me.txt_申込種別 = "02" Then
    '解約取消申込ならば
        strKotoMsg = "ネットからの解約取消申込(契約番号:" & Me.txt_受注契約番号 & ")について契約入力で手動取込しました。"
    Else
        strKotoMsg = "ネットからの未知の申込(契約番号:" & Me.txt_受注契約番号 & ")について契約入力で手動取込しました。"
    End If
    
    '加瀬DB接続
    strKagtosSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    strKagtosConnect = MSZZ007_M00()
    Set dbKagtos = Workspaces(0).OpenDatabase(strKagtosSqlserver, dbDriverNoPrompt, False, strKagtosConnect)
    
    '顧客問合わせトラン構造体初期化
    Call CmReprMod.InitKotoTranType(pstKotoTran)
    
    '顧客問合わせトラン構造体に該当項目を設定
    pstKotoTran.BUMOC = strBumonCd                  '部門コード
    pstKotoTran.KOKYC = strUCODE                    '顧客コード
    pstKotoTran.TANTC = "888"                       '入力担当
    pstKotoTran.GENTD = strDate                     '発生日
    pstKotoTran.GENTJ = strTime                     '発生時刻
    pstKotoTran.TYPEC = "03"                        '入力区分
    pstKotoTran.CONTN = MidB(strKotoMsg, 1, 80)      '内容
    pstKotoTran.CNT1N = MidB(strKotoMsg, 81, 160)    '内容サブ１
    pstKotoTran.CNT2N = MidB(strKotoMsg, 161, 240)   '内容サブ２
    pstKotoTran.INSED = strDate                     '作成日付
    pstKotoTran.INSEJ = strTime                     '作成時刻
    pstKotoTran.INSPB = PROG_ID                     '作成プログラムID
    pstKotoTran.INSUB = strUser                     '作成ユーザーID
    
    Call CmReprMod.InsertKotoTran(dbKagtos, pstKotoTran)
    
    '加瀬DB切断
    dbKagtos.Close
    Set dbKagtos = Nothing
    
    fncInsKOTO_TRAN = True
    Exit Function
    
ErrorHandler:
    If Not objRsCarg Is Nothing Then objRsCarg.Close: Set objRsCarg = Nothing
    'KASEDB解放
    If Not dbKagtos Is Nothing Then dbKagtos.Close: Set dbKagtos = Nothing
    
    strErrMsg = Err.Description
    Call MSZZ024_M00("fncInsKOTO_TRAN", False)
    Call MSZZ003_M00(PROG_ID, "9", strErrMsg)
    
End Function
'==============================================================================*
'
'        MODULE_NAME    : FVS610 より起動するためのファンクション
'                         設定した値の名称を表示するための処理
'                         対象項目 → 解約区分、解約理由、解約受付担当者
'        MODULE_ID      : Call_AfterUpdate
'        CREATE_DATE    : 2019/04/23
'
'==============================================================================*
Public Function Call_AfterUpdate()

        Call cmb_解約区分_AfterUpdate
        Call cmb_解約理由_AfterUpdate
        Call cmb_解約受付担当者コード_AfterUpdate

End Function

'↑ADD 2019/04/23 Y_SUZUKI EGL
'****************************  ended of program ********************************

'==============================================================================*
'        MODULE_NAME      : 顔認証ステータス取得
'        MODULE_ID        : GetFaceAuthStatus
'        CREATE_DATE      : 2026/04/28
'        処理内容         : APIを叩いて顔認証の登録状態を返す
'==============================================================================*
Private Function GetFaceAuthStatus(yardCode As String, roomCode As String) As String
    Dim req As Object
    Dim url As String
    Dim resText As String
    Dim regEx As Object
    
    On Error GoTo ErrHandler
    
    GetFaceAuthStatus = "" ' 初期値
    
    ' API URLの組み立て
    url = "https://dev-facekey-admin.kase3535.com/api/v1/external/system/rooms/detail?yard_code=" & yardCode & "&room_code=" & roomCode
    
    ' HTTPリクエスト送信 (遅延バインディング)
    Set req = CreateObject("MSXML2.XMLHTTP")
    req.Open "GET", url, False
    req.setRequestHeader "Content-Type", "application/json"
    
    ' APIキーをヘッダーにセット（※API側の仕様に合わせてヘッダー名は調整してください）
    req.setRequestHeader "x-api-key", "ff246566d3981d6d571f0011cbc1668f15172d6e73d6a65b827e681b6c6de492"
    
    req.send
    
    ' レスポンスの確認
    If req.Status = 200 Then
        resText = req.responseText
        
        ' 正規表現で "has_face": true を検索
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.Pattern = """has_face""\s*:\s*true"
        
        ' 1つでも true が見つかった場合、登録済みと判定
        If regEx.test(resText) Then
            GetFaceAuthStatus = "登録済み"
        Else
            GetFaceAuthStatus = "未登録"
        End If
    Else
        ' 部屋情報がない、APIエラーなどの場合
        GetFaceAuthStatus = "未登録"
    End If
    
    Set req = Nothing
    Set regEx = Nothing
    Exit Function

ErrHandler:
    ' 通信エラー等のハンドリング
    GetFaceAuthStatus = "取得エラー"
    Set req = Nothing
    Set regEx = Nothing
End Function


