VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　予約入力
'   プログラムＩＤ　：　FVS500
'   作　成　日　　　：  2005/05/06
'   作　成　者　　　：  ChronoSystem 三ツ野
'   Ver             ：  0.0
'   備考            ：　近隣ヤードchkﾎﾞｯｸｽの値設定がvalueで出来ない。2005/05/27
'**********************************************
'   プログラム名　　：　予約入力
'   プログラムＩＤ　：　FVS500
'   作　成　日　　　：  2005/06/25
'   作　成　者　　　：  K.MITSUNO
'   Ver             ：  0.1
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/07/17
'   修　正　者　　　：  K.MITSUNO
'   修　正　内　容　：  地主契約の検索条件変更 ほか
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/07/17
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  閉じるときに前画面を再表示していたが、保存後に変更
'   Ver             ：  0.2
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/07/21
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  入力がなければNullに変換する（日付型なのに空文字に変換されていた）
'   Ver             ：  0.3
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/07/26
'   修　正　者　　　：  K.MITSUNO
'                      T.SUZUKI
'   修　正　内　容　：  1.収納物入力支援処理追加/顧客検索ﾌﾘｶﾞﾅﾘｽﾄ選択変更
'                      2.受付版番号ボタン押下時、予約受付検索画面を表示する機能を追加
'   Ver             ：  0.4
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/07/30
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  空き検索新版対応
'                       ヘルパー連動
'   Ver             ：  0.5
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/08/29
'   修　正　者　　　：  S.SHIBAZAKI
'   修　正　内　容　：  ヘルパーファイル名変更対応
'   Ver             ：  0.6
'**********************************************
'修正履歴
'   修　正　日　　　：  2005/10/06
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  1.変数の初期化を追加
'                       2.近隣ﾔｰﾄﾞ紹介チェックボックスの使用可否設定処理の追加
'                       3.[状態区分]に関わらずコンテナ番号を使用可能とするよう修正
'                       4.[用途]のデフォルト表示機能を追加
'                       5.[希望サイズ]のデフォルト表示機能を追加
'                       4.[紹介件数]の取得、[紹介履歴ボタン]の使用可否設定処理の追加
'                       5.予約入力データ登録時の予約済みチェック、契約済みチェック、稼働ヤードチェックを
'                         共通関数で行うように修正
'                       6.データ登録時の「状態区分」の必須入力チェックの不具合を修正
'                       7.ｻﾌﾞﾌｫｰﾑ表示用のﾃﾞｰﾀ作成、削除処理を追加
'                       8.予約ご紹介画面表示のロジックを追加
'                       9.「紹介件数」→「ご紹介件数」に項目名称変更
'                      10.「紹介履歴」→「ご紹介一覧」に項目名称変更
'                      11.顧客名カナコンボボックスを廃止し、テキストボックスに変更
'                      12.画面に自動区分の内容を表示する処理を追加
'
'   修　正　日　　　：  2005/12/12
'   修　正　者　　　：  T.SUZUKI
'                       1.予約受付状態区分のｺﾝﾎﾞﾎﾞｯｸｽのデータを動的に変更するように処理を追加
'                       2.予約受付状態区分により画面の入力制限を行うよう処理を追加
'                       3.予約受付トラン.自動区分の更新内容を変更
'                       5.画面のラベルの標題を変更
'   修　正　日　　　：  2006/01/05
'                       1.日付項目の入力チェック変更(｢/｣ありでの入力を可能とした)
'   修　正　日　　　：  2006/01/20
'                       1.予約受付状態区分が自動紹介の場合、希望サイズ(From～To)を必須入力とするよう修正
'   修　正　日　　　：  2006/02/18
'                       1.修正中に終了ボタンを押下した場合、確認メッセージを表示するように処理を追加
'                       2.画面に【既存チェック】ボタンを追加し、押下されたら『既存顧客検索画面』を表示する
'                         処理を追加
'   Ver             ：  0.7
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/01/28
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  移動予約対応
'   Ver             ：  0.9
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/02/07
'   修　正　者　　　：  T.SUZUKI
'   修　正　内　容　：  受付入力対応
'   Ver             ：  1.0
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/05/02
'   修　正　者　　　：  tajima
'   修　正　内　容　：  紹介待ち障害対応 ※2006/4/20メール【予約入力不具合】参照
'   Ver             ：  1.1
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/05/12
'   修　正　者　　　：  tajima
'   修　正　内　容　：  受付入力課題対応
'                       1.郵便番号変換による住所１の桁溢れ対応
'                       2.近隣ヤード紹介チェックの保存機能付加
'                         ※プロパティとSetFrmEnabledConfで無効化している！
'                       3.移動予約時の不具合、顧客を変えたら元契約番号をクリア
'   Ver             ：  1.2
'**********************************************
'修正履歴
'   修　正　日　　　：  2006/05/31
'   修　正　者　　　：  tajima
'   修　正　内　容　：  セキュリティカード対応
'                       1.削除やキャンセルしたときに貸出カードを元に戻す
'                       2.契約状態では予約削除不可能とした
'                       3.新規/既存オプションをキーボードで変更時で対応するようにｼﾀ
'   Ver             ：  1.3
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/06/24
'   修　正　者　　　：  tajima
'   修　正　内　容　：  修繕呼出対応
'   Ver             ：  1.4
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/09/26
'   修　正　者　　　：  tajima
'   修　正　内　容　：  修繕対応その２
'                       ①部屋と顧客の直近修繕を表示
'                       ②営業期間の表示
'                       ③希望開始日を制御にする
'   Ver             ：  1.5
'**********************************************
'修正履歴
'   修　正　日　　　：　2006/11/15
'   修　正　者　　　：  tajima
'   修　正　内　容　：  利用状況、空き検索からコンテナ確定で遷移した場合
'                       予約状態は「受付」とする
'   Ver             ：  1.6
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/04/06
'   修　正　者　　　：  tajima
'   修　正　内　容　：  ①近隣紹介チェックの開放、デフォルトでチェック済みとする
'                   ：  ②元々どのヤードが希望だったのか判るようにする
'                   ：  ③営業開始日前のヤードならば近隣紹介チェックを制御する
'                   ：  ④入力項目・移動担当追加
'   Ver             ：  1.7
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/05/25
'   修　正　者　　　：  tajima
'   修　正　内　容　：  ①新規登録時に「発生区分」を "0"(電話)で設定する
'                   ：  ②新規登録時に「登録時予約状態」を設定する
'   Ver             ：  1.8
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/06/02
'   修　正　者　　　：  tajima
'   修　正　内　容　：  ①空詳細からレンタル用途をパラメータで受けれるようにする
'   Ver             ：  1.9
'**********************************************
'修正履歴
'   修　正　日　　　：　2007/08/09
'   修　正　者　　　：  tajima
'   修　正　内　容　：  ①ネット予約対応
'   Ver             ：  2.0
'**********************************************
'修正履歴
'   修　正　日　　　：　2008/03/23
'   修　正　者　　　：  tajima
'   修　正　内　容　：  ①移動区分追加コード「10:満了更新」対応
'   Ver             ：  2.1
'**********************************************
'修正履歴
'   修　正　日　　　：　2008/08/18
'   修　正　者　　　：  S.SHIBAZAKI
'   修　正　内　容　：  「受付」→「受付キャンセル」の際にエラーになる不具合を修正
'   Ver             ：  2.2
'**********************************************
'修正履歴
'   修　正　日　　　：　2008/11/19
'   修　正　者　　　：  hirano
'   修　正　内　容　：  「受付」→「受付キャンセル」の際に適用割引割付を物理削除する。
'                       予約「削除」の際、適用割引割付を物理削除する。
'   Ver             ：  2.3
'**********************************************
'修正履歴
'   修　正　日　　　：　2010/02/15
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  顧客情報の桁チェックエラー時に文字を消さないように修正
'
'   Ver             ：  2.4
'**********************************************
'修正履歴
'   修　正　日　　　：　2011/02/21
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  統合ＤＢへの書き込み、統合新画面のリロード
'   Ver             ：  2.5
'**********************************************
'修正履歴
'   修　正　日　　　：　2011/03/31
'   修　正　者　　　：  S.SHIBAZAKI
'   修　正　内　容　：  用途から「コンテナ（バイク）」「バイク（コンテナ）」を除外する
'   Ver             ：  2.6
'**********************************************
'修正履歴
'   修　正　日　　　：　2011/07/19
'   修　正　者　　　：  tajijma
'   修　正　内　容　：  発生区分を選択可能にする
'   Ver             ：  2.7
'**********************************************
'   修　正　日　　　：  2011/09/27
'   修　正　者　　　：  M.RYU
'   修　正　内　容　：  移動区分を追加⇒07紹介者金額変更
'   Ver　　         ：  2.8
'**********************************************
'   修　正　日　　　：  2012/11/21
'   修　正　者　　　：  K.SHIBUYA
'   修　正　内　容　：  フリガナの下に文言追加
'   Ver　　         ：  2.9
'**********************************************
'   修　正　日　　　：  2013/06/03
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  終了ボタンのメッセージを削除
'   Ver　　         ：  3.0
'**********************************************
'   修　正　日　　　：  2013/11/27
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  見積もりのときは、コンテナ番号は引き継いできたものとする
'   Ver　　         ：  3.1
'**********************************************
'   修　正　日　　　：  2014/08/01
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  近隣ヤード紹介のデフォルトチェックボックスをOFF
'   Ver　　         ：  3.2
'**********************************************
'   修　正　日　　　：  2014/12/12
'   修　正　者　　　：  K.ISHIZAKA
'   修　正　内　容　：  Webも考慮し空き件数をチェックする
'   Ver　　         ：  3.3
'**********************************************
'   修　正　日　　　：  2015/01/31
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  空数件数ﾁｪｯｸ条件を修正
'   Ver　　         ：  3.4
'**********************************************
'   修　正　日　　　：  2015/02/06
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  空数件数ﾁｪｯｸ条件を修正
'   Ver　　         ：  3.5
'**********************************************
'   修　正　日　　　：  2015/02/13
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  オフィスの場合には空数ﾁｪｯｸを行わない。
'   Ver　　         ：  3.6
'**********************************************
'   修　正　日　　　：  2015/08/05
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  空数件数ﾁｪｯｸ条件を修正
'   Ver　　         ：  3.7
'**********************************************
'   修　正　日　　　：  2017/08/07
'   修　正　者　　　：  ysuzuki
'   修　正　内　容　：  以下の条件の場合、エラーメッセージを出力して処理を中断。
'                       修正モード and 予約状態=53:受注キャンセル and コンテナ契約にデータが在る場合
'   Ver　　         ：  3.8
'**********************************************
'   修　正　日　　　：  2018/02/10
'   修　正　者　　　：  sato
'   修　正　内　容　：  収納物＝オートバイのとき、コンテナマスタの用途をコンテナ（バイク）に変更する。
'                       同一サイズの価格表がなかったら用途＝コンテナのレコードをコピーして作る。
'   Ver　　         ：  3.9
'**********************************************
'   修　正　日　　　：  2018/03/01
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  受付状態で内容を修正した後に保存を行うと
'                       「該当する段、サイズでの空きがありません。」エラーが発生する問題に対応
'   Ver　　         ：  4.0
'**********************************************
'   修　正　日　　　：  2018/03/17
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  収納物の用途を変えるのはコンテナのみに修正
'
'   Ver　　         ：  4.1
'**********************************************
'   修　正　日　　　：  2018/04/16
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  賃料のFrom～Toチェックが文字列でチェックを行っているので正しくないバグに対応
'   Ver　　         ：  4.2
'**********************************************
'   修　正　日　　　：  2018/08/30
'   修　正　者　　　：  TAKENOUCHI
'   修　正　内　容　：  No32.予約入力の帖数指定方変更、No80.事務手数料対応
'   Ver　　         ：  4.3
'**********************************************
'**********************************************
'   修　正　日　　　：  2018/09/25
'   修　正　者　　　：  WATANABE
'   修　正　内　容　：  分社化対応
'   Ver　　         ：  4.4
'**********************************************
'   修　正　日　　　：  2019/02/05
'   修　正　者　　　：  tajima
'   修　正　内　容　：  バイク予約時の価格表作成不具合修正(Ver4.3不具合対応)
'   Ver　　         ：  4.5
'**********************************************
'   修　正　日　　　：  2019/03/06
'   修　正　者　　　：  Y.WADA
'   修　正　内　容　：　１．広告媒体のリストに「８１：他部署からの移動」を追加する。
'　                     ２．チェックボックス「他部署からの移動」を追加する。
'　                     ３．リストとチェックボックスは連動する。
'　　                       ・チェックボックスONで、リストを選択不可、「８１：他部署からの移動」をセットする。
'　　                       ・チェックボックスOFFで、リストを選択可、「９９：その他」をセットする。
'　　                       ・リストで「８１：他部署からの移動」を選択時に、チェックボックスOFFにする。
'　                     ４.コンテナ契約で移動の入力をすると､既に移動したというエラーになるが､入力できるようにする｡
'   Ver　　         ：  4.6
'**********************************************
'   修　正　日　　　：  2019/03/25
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  担当者・広告媒体・予定収納物のコンボボックを整理
'   Ver　　         ：  4.7
'**********************************************
'   修　正　日　　　：  2019/05/20
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  収納物でバイクを選択時の不具合対応
'   Ver　　         ：  4.9
'**********************************************
'   修　正　日　　　：  2019/08/19
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  予約入力で元々入力されていた都道府県市町村を空白に変更して保存するとエラーが発生する問題に対応
'   Ver　　         ：  5.0
'**********************************************
'   修　正　日　　　：  2020/01/01
'   修　正　者　　　：  Y.WADA
'   修　正　内　容　： １．必須入力の項目にする
'                            ・予約担当、ヤード、用途、希望サイズ（コンボボックスのみ）
'                              段区分､顧客名､携帯番号
'                      ２．用途のコンボボックス
'                            予約のヤードに存在する､用途のみを表示する
'                            ただし、コンテナ（バイク）、社内（自動販売機）は、表示しない。
'                      ３．希望サイズのコンボボックス
'                            予約のヤード､用途に存在する､希望サイズのみを表示する｡
'                            ただし、コンテナ（バイク）、社内（自動販売機）は、表示しない。
'                      ４．段区分のコンボボックス
'                            予約のヤード､用途､希望サイズに存在する､段区分のみを表示する｡
'                            ただし、コンテナ（バイク）、社内（自動販売機）は、表示しない。
'                      ５．賃料
'                            段区分を入力時に､プライスマスタから賃料を取得して表示する｡
'   Ver　　         ：  5.1
'**********************************************
'   修　正　日　　　：  2020/01/10
'   修　正　者　　　：  Y.WADA
'   修　正　内　容　：  利用状況-予約入力で呼ばれた場合に、用途以下のコンボボックスのデータが空で表示されるのを修正
'   Ver　　         ：  5.2
'**********************************************
'   修　正　日　　　：  2020/02/06
'   修　正　者　　　：  M.HONDA
'   修　正　内　容　：  上段・下段データ取得時の条件を追加
'   Ver　　         ：  5.3
'**********************************************
'   修　正　日　　　：  2020/02/28
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  CARG_FILE、REQU_FILEの削除を追加
'   Ver　　         ：  5.4
'**********************************************
'   修　正　日　　　：  2020/08/07
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  削除処理時もCARG_FILE、REQU_FILEの削除を追加
'   Ver　　         ：  5.5
'**********************************************
'   修　正　日　　　：  2020/08/27
'   修　正　者　　　：  Takenouchi
'   修　正　内　容　：  料金改定対応
'   Ver　　         ：  5.6
'**********************************************
'   修　正　日　　　：  2020/09/30
'   修　正　者　　　：  EGL
'   修　正　内　容　：  制御テーブル参照方ミス修正
'   Ver　　         ：  5.7
'**********************************************
'   修　正　日　　　：  2021/04/28
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  「自動紹介」を画面から「受付」に変更した際に起算日へ月初をセット
'   Ver　　         ：  5.8
'**********************************************
'   修　正　日　　　：  2022/02/14
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  起算日空白の場合、エラーが発生していた問題に対応
'   Ver　　         ：  5.9
'**********************************************
'   修　正　日　　　：  2024/08/27
'   修　正　者　　　：  N.IMAI
'   修　正　内　容　：  予約番号コンボにセットする予約番号を全件→過去10日分(INTI_FILEで指定)に変更
'   Ver　　         ：  6.0
'**********************************************
'   修　正　日　　　：  2025/12/04
'   修　正　者　　　：  T.KAWABATA
'   修　正　内　容　：  YARD_MASTカラム追加によるカラム上限エラーの解消 https://kaseit.backlog.com/view/IT-1208
'   Ver　　         ：  6.1
'**********************************************
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Const WORK_TAB01 = "WK_FVS500_TEMP1"    '用途コンボボックス用          'INSERT 2020/01/01 Y.WADA
Private Const WORK_TAB02 = "WK_FVS500_TEMP2"    '希望サイズコンボボックス用    'INSERT 2020/01/01 Y.WADA
Private Const WORK_TAB03 = "WK_FVS500_TEMP3"    '段区分コンボボックス用        'INSERT 2020/01/01 Y.WADA


'フォーム読込フラグ
Private bln_FormLoad    As Boolean

'処理モード管理変数(詳細)
Private mdl_Mode        As Long
'顧客情報ﾓｰﾄﾞ管理変数
Private mdl_Kok         As Long

'呼出先のプログラムID
Private var_OpenArgs    As Variant

Private objDb           As Database

Private svKcode         As Variant
Private svContna        As Variant

'呼出先のデータ
Private strForm_Yard    As String  'ﾔｰﾄﾞ番号
Private strForm_Kont    As String  'ｺﾝﾃﾅ番号
Private strForm_UkeNo   As String  '受付№
Private strForm_User    As String  '顧客コード      ' 2005/12/27 ADD T.SUZUKI
Private strFrom_Yoyaku  As String  '予約入力モード  ' 2006/01/18 ADD T.SUZUKI
Private strFrom_AcptNo  As String  '受注契約番号    ' 2006/01/18 ADD T.SUZUKI

Private strBeforYard    As String  'DBから読み込んだヤード

Public pIs受付作成済   As Boolean  ' 2006/06/14 ※サブフォームも見る

'▼ 2006/09/26 add tajima ▼
Private str営業開始日  As String
Private str営業終了日  As String
Private pobjKASE_DB    As Database ' 加瀬DB
'▼ 2007/06/02 add tajima ▼
Private str空検索時レンタル用途 As String
'▼ 2007/08/09 add tajima ▼
Private isネット予約  As Boolean

Private strCampCd     As String   '2018/09/25 Add WATANABE

Private isAfterUpdate  As Boolean    'INSERT 2019/03/06 Y.WADA

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Const P_PROG_ID     As String = "FVS500"
Private Const P_KAKUNIN     As String = "確認"
Private Const P_ERROR       As String = "エラー"

'識別(受付番号採番用)
Private Const P_UKE_U       As String = "U"

'状態区分
Private Const P_JOU_MI      As String = "00"    'ｺﾝﾃﾅ未確定初期値
'Private Const P_JOU_KAKU    As String = "01"    'ｺﾝﾃﾅ確定初期値
Private Const P_JOU_KAKU    As String = "10"    '2006/11/15 初期値は受付に


'名称ﾏｽﾀｺｰﾄﾞcarg_
Private Const P_NAME_020    As String = "020"
Private Const P_NAME_040    As String = "040"
Private Const P_NAME_070    As String = "070"
Private Const P_NAME_080    As String = "080"
Private Const P_NAME_085    As String = "085"
Private Const P_NAME_086    As String = "086"
Private Const P_NAME_087    As String = "087"
Private Const P_NAME_090    As String = "090"
Private Const P_NAME_091    As String = "091"   '2005/07/24 K.MITSUNO

Private Const P_NAME_095    As String = "095"   ' 2005/10/06 ADD T.SUZUKI 自動区分
Private Const P_NAME_098    As String = "098"   ' 2006/01/18 ADD T.SUZUKI 移動区分
Private Const P_NAME_083    As String = "083"   ' 2011/07/19 ADD tajima 発生区分

'状態区分ｺｰﾄﾞ
Private Const P_YUKBN_00    As String = "00"    '自動紹介
Private Const P_YUKBN_01    As String = "01"    '自社理由
Private Const P_YUKBN_02    As String = "02"    '取置
Private Const P_YUKBN_10    As String = "10"    '受付
Private Const P_YUKBN_20    As String = "20"    '契約
Private Const P_YUKBN_51    As String = "51"    '予約キャンセル
Private Const P_YUKBN_52    As String = "52"    '取置キャンセル
Private Const P_YUKBN_53    As String = "53"    '受付キャンセル  ' 2005/12/12 ADD T.SUZUKI

'連絡区分
Private Const P_REN_00      As String = "00"
Private Const P_REN_01      As String = "01"
Private Const P_REN_02      As String = "02"
Private Const P_REN_10      As String = "10"

'受付区分
Private Const P_UKKBN_99    As String = "99"
'顧客区分
Private Const P_KKBN_1      As String = "01"
'連絡区分
Private Const P_RENKBN_1    As String = "00"

'遷移元ﾌｫｰﾑｺｰﾄﾞ
Private Const P_FRM_YRD     As Long = 1     'ﾔｰﾄﾞ契約状況予約画面より
Private Const P_FRM_KNT     As Long = 2     '空ｺﾝﾃﾅ検索予約画面より
Private Const P_FRM_MNU     As Long = 3     'ﾒﾆｭｰより

'処理区分(画面左上ボタン用)
Private Const P_NEW         As Long = 1     '新規
Private Const P_SHU         As Long = 2     '修正
Private Const P_DEL         As Long = 3     '削除

'顧客処理ｺｰﾄﾞ
Private Const P_SHINKI      As Long = 0     '新規
Private Const P_KIZON       As Long = 1     '既存

' 2005/10/06 ADD T.SUZUKI Start
Private pobjKONT_DB               As Database             ' コンテナDB

'使用用途
Private Const P_TRANK            As String = "00"         ' トランクルーム
Private Const P_RENTAL           As String = "03"         ' レンタルボックス
Private Const P_TorR             As String = "10"         ' レンタルorトランク

Private Const P_SHU_Caption      As String = "(修正中)"
Private Const P_SHO_Caption      As String = "(他画面編集中)"                 ' 2006/03/11 ADD T.SUZUKI
'Private Const P_SHO_Caption      As String = "(取置履歴照会中)"              ' 2006/03/11 DEL T.SUZUKI

Private Const pcstrUPDMsg        As String = "更新しますか？"
Private Const pcstrMoveMsg       As String = "受付で保存して下さい。"
Private Const pcstrKAKUMsg       As String = "※更新しなければ取置履歴は照会できません。"

Private Const pcstrFVS540        As String = "FVS540"     ' 予約一覧画面
Private Const pcstrFVS600        As String = "FVS600"     ' 契約入力画面
Private Const pstrCstWK_TABLE_NM As String = "FVS500_W01"
Private Const pstrCstSubFromName As String = "FVS500S01"  ' サブフォーム名

Private Const UKNO_WK_TABLE As String = "FVS500_WK_UKNO"    'INSERT 2024/08/27 N.IMAI

Private pstrFormCaption          As String                ' 画面標題格納
Private pblnChangeFlag_FVS500    As Boolean               ' 入力に変更があったかどうかのフラグ
Private pstrFormID               As String                ' 呼出元画面ID
Private pblnUpdateFlg            As Boolean               ' 画面にて変更を保存した際にTrue
Private blnFVS540FLG             As Boolean               ' 呼出元画面が「FVS540｣のときにTrue
' 2005/10/06 ADD T.SUZUKI End

' 2005/12/12 ADD T.SUZUKI Start
Private Const pcstrfmtSTANDARD   As String = "Standard"   ' 書式：標準
Private Const pcstrfmtFIXED      As String = "Fixed"      ' 書式：固定
Private pintAUTOKBN              As Integer               ' 自動区分
Private pstrYUKBN                As String                ' 読み込み時の予約受付状態区分
Private pstrKONT_NO              As String                ' 読み込み時のコンテナ番号

Private Const P_INTROKBN_01      As String = "1"          '取置
Private Const P_INTROKBN_02      As String = "2"          '契約受付
Private Const P_INTROKBN_06      As String = "6"          'キャンセル
' 2005/12/12 ADD T.SUZUKI End

' 2005/12/19 ADD T.SUZUKI Start
Private Const pcstrSelErrMsg     As String = "指定した予約状態は選択できません。"
Private Const pcstrHonsuErrMsg   As String = "本数は、｢1｣以上の値を入力して下さい。"
Private Const pcintHonsu         As Long = 0              '本数の入力エラー値
Private VarChangeDate            As Variant
' 2005/12/19 ADD T.SUZUKI End

' 2006/01/18 ADD T.SUZUKI Start
' 移動予約画面遷移区分
Private Const P_YOYAKU_0         As String = "0"          ' 移動予約(新規モード)
Private Const P_YOYAKU_1         As String = "1"          ' 移動予約(修正モード)
Private Const P_YOYAKU_2         As String = "2"          ' 追加予約

' 画面タイトル
Private Const P_TITTLE_YOYAKU    As String = "予約入力"

' 【移動区分】
Private Const P_MOVEKBN_00          As String = "00"          ' 無し
Private Const P_MOVEKBN_01          As String = "01"          ' 自社理由移動
Private Const P_MOVEKBN_02          As String = "02"          ' 顧客理由移動
Private Const P_MOVEKBN_03          As String = "03"          ' 未使用移動
Private Const P_MOVEKBN_04          As String = "04"          ' 名義変更
Private Const P_MOVEKBN_05          As String = "05"          ' 金額変更
Private Const P_MOVEKBN_06          As String = "06"          ' 保証会社変更
Private Const P_MOVEKBN_07          As String = "07"          ' 紹介者金額変更  'INSERT 2011/09/27 M.RYU
Private Const P_MOVEKBN_満了更新    As Integer = 10           '2008/03/23 add tajima

' 移動予約入力エラーメッセージ
Private Const pcstrMoveErrMsg01     As String = "の場合は予約状態を" & vbCrLf & "「受付」にして下さい。"
Private Const pcstrMoveErrMsg02     As String = "契約一覧から「移動元契約番号」を選択して下さい。"
' 2006/01/18 ADD T.SUZUKI End

' 2006/02/07 ADD T.SUZUKI Start
Private Const pcstrNoDataMsg        As String = "該当データが存在しません。"
Private Const pcstrUkeErrMsg01      As String = "※更新しなければ受付入力は表示できません。"
Private Const pcstrUkeErrMsg02      As String = "受付でなければ表示できません。"
Private Const pcstrUkeErrMsg03      As String = "受付データが存在しません。"
' 2006/02/07 ADD T.SUZUKI Start

' 2006/02/18 ADD T.SUZUKI Start
Private Const pcstrEndEnterMsg      As String = "修正中ですが保存せずに終了しますか？"
Private Const pcstrEndMsg           As String = "終了します。よろしいですか？"
' 2006/02/18 ADD T.SUZUKI End

Private pWRI_ObjWS                  As Workspace              ' 2006/04/07 ADD T.SUZUKI

'▼ 2006/09/26 add tajima ▼
Private Const pcstr最大年月日       As String = "9999/12/31"

'INSERT 2018-02-10 add...2018/08/30 FZ2,FZ3,EZ2,EZ3を追加した...2019/02/05 FZ2,FZ3,EZ2,EZ3をvariant型にした
Private Type Type_PRIC
    PRIC_PRICE                As String
    PRIC_FZAPPI_CODE          As String
    PRIC_FZAPPI               As String
    PRIC_EZAPPI_CODE          As String
    PRIC_EZAPPI               As String
    PRIC_NEBIKI_DISABLE       As String
    PRIC_FZAPPI_CODE2         As String
    PRIC_FZAPPI2              As Variant
    PRIC_FZAPPI_CODE3         As Variant
    PRIC_FZAPPI3              As Variant
    PRIC_EZAPPI_CODE2         As Variant
    PRIC_EZAPPI2              As Variant
    PRIC_EZAPPI_CODE3         As Variant
    PRIC_EZAPPI3              As Variant
End Type
Private recPric As Type_PRIC

'==============================================================================*
'
'        MODULE_NAME      :chk_Tabusyo_AfterUpdateイベント
'        MODULE_ID        :chk_Tabusyo_AfterUpdate
'        CREATE_DATE      :2019/03/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_Tabusyo_AfterUpdate()
    If chk_Tabusyo.VALUE = True Then
        cmb_Ukkbn.VALUE = "81"      '他部署から移動
        cmb_Ukkbn.Enabled = False
    Else
        cmb_Ukkbn.VALUE = "99"      'そのた
        cmb_Ukkbn.Enabled = True
    End If
    
    isAfterUpdate = False
    Call cmb_Ukkbn_AfterUpdate
    isAfterUpdate = True

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form Openイベント
'        CREATE_DATE      :2020/01/01 Y.WADA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()

    'コンボボックスのワークテーブル削除
    cmb_Youto.RowSource = ""
    Call DropWorkTable_CurrentDb(WORK_TAB01)
    cmb_Kibou.RowSource = ""
    Call DropWorkTable_CurrentDb(WORK_TAB02)
    cmb_Step.RowSource = ""
    Call DropWorkTable_CurrentDb(WORK_TAB03)
    cmb_Ukno.RowSource = ""                                                     'INSERT 2024/08/27 N.IMAI
    Call DropWorkTable_CurrentDb(UKNO_WK_TABLE)                                 'INSERT 2024/08/27 N.IMAI

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form Openイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)

    Dim intKbn As Integer
     Dim strCONT_BUMOC As String
    strCONT_BUMOC = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
    Select Case strCONT_BUMOC
    Case "H"
        lbl_Youk.BackColor = RGB(255, 0, 204)
'        フォームヘッダー.BackColor = RGB(137, 189, 222)
    End Select

    '▼ add 2007/06/02  tajima ▼
    ' 空詳細からキタ場合レンタル用途を第二パラメータとして設定したいが
    ' 元々の作りとして１パラメータに全情報を載っけようとするなんてどうかしているので
    ' 本追加ロジックで対応してみたよ
    Dim args()  As String
    args = Split(Nz(Me.OpenArgs), ",")

    str空検索時レンタル用途 = ""
    
    If UBound(args) > 0 Then
      var_OpenArgs = args(0)
      str空検索時レンタル用途 = Nz(args(1), "")
    ElseIf UBound(args) = 0 Then
      var_OpenArgs = Nz(args(0), "")
    Else
      var_OpenArgs = ""
    End If
    '▲ add 2007/06/02 tajima ▲

    ' パラメータセット
    ' var_OpenArgs = Nz(OpenArgs, "") del 2007/06/02 tajima

    blnFVS540FLG = False

    '前画面からは、渡ってくるデータパターン
    '①予約確定受付  ：ﾔｰﾄﾞ番号(6桁)+ｺﾝﾃﾅ番号(6桁※ｺﾝﾃﾅ確定時)+受付番号(11桁)
    '②予約未確定受付：ﾔｰﾄﾞ番号(6桁)+受付番号(11桁)
    '③ｺﾝﾃﾅ未確定    ：ﾔｰﾄﾞ番号(6桁)
    '④ﾒﾆｭｰより      ：なし
    If var_OpenArgs = "" Then
       doCmd.Maximize
    Else
        strForm_Yard = Trim$(Left$(var_OpenArgs, 6))                         ' ﾔｰﾄﾞｺｰﾄﾞ
        strForm_Kont = Trim$(Mid$(var_OpenArgs, 7, 6))                       ' ｺﾝﾃﾅ番号
        strForm_UkeNo = Trim$(Mid$(var_OpenArgs, 13, 11))                    ' 受付番号

' 2005/10/06 ADD T.SUZUKI Start
        If Len(var_OpenArgs) > 23 Then
            If Len(var_OpenArgs) < 30 Then
                ' 予約一覧画面(FVS540)からの遷移の対応
                pstrFormID = Trim$(Mid$(var_OpenArgs, 24, Len(var_OpenArgs) - 23))   ' 画面ID(呼び出し元)
                blnFVS540FLG = True

            Else
                ' 契約入力画面からの遷移のパターン
                strFrom_Yoyaku = Trim$(Mid$(var_OpenArgs, 24, 1))

                Select Case strFrom_Yoyaku
                    Case P_YOYAKU_0
                        ' *** 移動予約(新規) ***
                        strForm_User = Trim$(Mid$(var_OpenArgs, 25, 6))              ' 顧客コード
                        strFrom_AcptNo = Trim$(Mid$(var_OpenArgs, 31, 10))           ' 受注契約番号

                    Case P_YOYAKU_1
                        ' *** 移動予約(修正) ***
                        ' 上記でパラメータ設定済

                    Case P_YOYAKU_2
                        ' *** 追加予約 ***
                        strForm_User = Trim$(Mid$(var_OpenArgs, 25, 6))              ' 顧客コード
                End Select

'                ' 契約入力画面(FVS600)からの遷移の対応
'                pstrFormID = Trim$(Mid$(var_OpenArgs, 30, Len(var_OpenArgs) - 29))   ' 画面ID(呼び出し元)
            End If
' 2005/12/27 DEL T.SUZUKI Start
'        If LenB(var_OpenArgs) > 23 Then
'            pstrFormID = Trim$(Mid$(var_OpenArgs, 24, LenB(var_OpenArgs) - 23))  ' 画面ID(呼び出し元)
'            If pstrFormID = pcstrFVS540 Then
'                blnFVS540FLG = True
'            Else
'                blnFVS540FLG = False
'            End If
' 2005/12/27 DEL T.SUZUKI End
        End If
' 2005/10/06 ADD T.SUZUKI End
    End If

'DELETE 2020/01/10 Y.WADA Start
'    'INSERT 2020/01/01 Y.WADA Start
'    Call make_cmb_Youto(True)       '用途コンボボックス用ワークテーブル作成
'    Call make_cmb_Kibou(True)       '希望サイズコンボボックス用ワークテーブル作成
'    Call make_cmb_Step(True)        '段区分コンボボックス用ワークテーブル作成
'    'INSERT 2020/01/01 Y.WADA End
'DELETE 2020/01/10 Y.WADA End

End Sub
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
    'フルスクリーン化
    doCmd.Maximize
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

    Dim lngRc        As Long
    Dim blnEnableFlg As Boolean     ' 2005/12/27 ADD T.SUZUKI

    doCmd.Maximize
    bln_FormLoad = False

    isAfterUpdate = True    'INSERT 2019/03/06 Y.WADA

' 2006/02/07 DEL T.SUZUKI Start
'    Me.cmd_Excel.Enabled = MSZZ013_M00("Excel.Application")                     'INSERT 20050730 K.ISHIZAKA
' 2006/02/07 DEL T.SUZUKI End

    Set objDb = CurrentDb

' 2005/10/06 ADD T.SUZUKI Start
    ' 画面の標題格納
    pstrFormCaption = Me.Caption

    ' 変更フラグ初期化
    pblnUpdateFlg = False
' 2005/10/06 ADD T.SUZUKI End

    '処理ﾓｰﾄﾞ決定
    ' mdl_Mode(1：新規／2：修正／3：削除)
    Call SetModeConf(strForm_Yard, strForm_Kont, strForm_UkeNo)

    '処理ﾓｰﾄﾞにより処理ﾓｰﾄﾞﾎﾞﾀﾝの設定及び対応した処理へFormの初期化
    Call SetModeBtn(mdl_Mode)

    'ｺﾝﾄﾛｰﾙのﾛｯｸ
    Call Etc_Lock_Rtn

' 2006/01/18 ADD T.SUZUKI Start
    blnEnableFlg = False

    '前画面から受付番号が渡ってきた時は、読み込んで画面へセット
    If strForm_UkeNo <> "" Then Call SetFrmEnabledConf(True): lngRc = GetYoukData(1)

    '契約入力画面からの遷移(移動予約(新規) OR 追加予約)
    '(※受付番号はNullで渡ってくる)
    If (strFrom_Yoyaku = P_YOYAKU_0 Or strFrom_Yoyaku = P_YOYAKU_2) Then
        Call SetFrmEnabledConf(True): Call SetYoyakuData(strFrom_Yoyaku)
    End If
'    If strFrom_Yoyaku = P_YOYAKU_2 Then SetFrmEnabledConf (True): Call SetYoyakuData
' 2006/01/18 ADD T.SUZUKI End
' 2006/01/18 DEL T.SUZUKI Start
'    '前画面から受付番号が渡ってきた時は、読み込んで画面へセット
'    If strForm_UkeNo <> "" Then
'        Call SetFrmEnabledConf(True)
'        lngRc = GetYoukData(1)
'    End If
' 2006/01/18 DEL T.SUZUKI End

    ' 受付番号使用可否設定
    Call SetUKNOBtnConf            ' 2005/07/26 ADD T.SUZUKI

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI

    'INSERT 2020/01/10 Y.WADA Start
    Call make_cmb_Youto(True)       '用途コンボボックス用ワークテーブル作成
    Call make_cmb_Kibou(True)       '希望サイズコンボボックス用ワークテーブル作成
    Call make_cmb_Step(True)        '段区分コンボボックス用ワークテーブル作成
    'INSERT 2020/01/10 Y.WADA End
    
    bln_FormLoad = True
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

   ' ▼2006/06/24 Add tajima
   Call CmReprMod.OpenFVS90n(KeyCode, _
                             Me.cmb_Yard, _
                             Me.cmb_No, _
                             Me.cmb_KokCd _
                             )
  ' ▲2006/06/24 Add tajima

    'F1=「終了」F2=「保存」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmd_End.SetFocus
        Case vbKeyF2
            KeyCode = 0
            If mdl_Mode = P_DEL Then
                Me.cmd_Delete.SetFocus
            Else
                Me.cmd_Enter.SetFocus
            End If
        Case vbKeyF3, vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select
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
'        MODULE_NAME      :Form_Close
'        機能             :Form_Closeイベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub Form_Close()                                                       'DELETE START   20050717 K.ISHIZAKA
'    On Error Resume Next
'    Dim beforForm As Form
'
'    Set objDB = Nothing
'    Set beforForm = Forms("FKS020_コンテナ予約入力")
'    If Not beforForm Is Nothing Then
'        Call beforForm.Reload
'    End If
'    Set beforForm = Forms("FVS020")
'    If Not beforForm Is Nothing Then
'        Call beforForm.Reload
'    End If
'End Sub                                                                        'DELETE END     20050717 K.ISHIZAKA
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

    Dim strMsg  As String
    Dim lngRc   As Long

    On Error GoTo Err_Form_Unload

' 2006/02/18 DEL T.SUZUKI Start
'' 2005/10/06 ADD T.SUZUKI Start
'    If blnFVS540FLG = True Then
'        ' 予約一覧画面からの遷移の場合には終了ﾒｯｾｰｼﾞの表示を行わない。
'        lngRc = vbYes
'    Else
'        '終了メッセージを表示
'        strMsg = "終了します。よろしいですか？"
'        lngRc = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, P_KAKUNIN)
'    End If
'' 2005/10/06 ADD T.SUZUKI End
' 2006/02/18 DEL T.SUZUKI End

' 2005/10/06 DEL T.SUZUKI Start
'    '終了メッセージを表示
'    strMsg = "終了します。よろしいですか？"
'    lngRc = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, P_KAKUNIN)
' 2005/10/06 DEL T.SUZUKI End

' 2006/02/18 DEL T.SUZUKI Start
'    If lngRc = vbNo Then
'       Cancel = True
'       If mdl_Mode = P_NEW Then Me.cmb_Yard.SetFocus
'    End If
' 2006/02/18 DEL T.SUZUKI End

    Exit Sub

Err_Form_Unload:
    MsgBox Error$
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opt_New_MouseDown
'        機能             :新規処理(画面左上ボタン）
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opt_New_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 2005/10/06 ADD T.SUZUKI Start
    If bln_FormLoad = True Then
        '各変数の初期化
        strForm_Yard = ""
        strForm_Kont = ""
        strForm_UkeNo = ""
        strFrom_Yoyaku = ""                  ' 2006/01/18 ADD T.SUZUKI
    End If

    Me.Caption = pstrFormCaption
    Call subSetFVS500(P_MOVEKBN_00)         ' 2006/01/18 ADD T.SUZUKI
' 2005/10/06 ADD T.SUZUKI End

    mdl_Mode = P_NEW              ' → 新規
    Call BtnConf(mdl_Mode)        ' 保存/削除ボタンの使用可否設定
    Call SetFrmEnabledConf(True)  ' 画面項目の使用可否設定
    Call FormInit(mdl_Mode)       ' 画面項目の初期化  ' 2005/10/25 ADD T.SUZUKI
'    Call FormInit(P_NEW)          ' 画面項目の初期化  ' 2005/10/25 DEL T.SUZUKI

    ' 受付番号使用可否設定
    Call SetUKNOBtnConf            ' 2005/07/26 ADD T.SUZUKI

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI

'    ' 予約受付状態区分ｺﾝﾎﾞﾎﾞｯｸｽのSQL作成
'    Call subCreate_YukbnSQL(P_NEW, Me.cmb_Yukbn)  ' 2005/12/12 ADD T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :opt_Correction_MouseDown
'        機能             :修正処理(画面左上ボタン）
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opt_Correction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 2005/10/06 ADD T.SUZUKI Start
    If bln_FormLoad = True Then
        '各変数の初期化
        strForm_Yard = ""
        strForm_Kont = ""
        strForm_UkeNo = ""
        strFrom_Yoyaku = ""                  ' 2006/01/18 ADD T.SUZUKI
    End If

    Me.Caption = pstrFormCaption
    Call subSetFVS500(P_MOVEKBN_00)          ' 2006/01/18 ADD T.SUZUKI
' 2005/10/06 ADD T.SUZUKI End

    mdl_Mode = P_SHU               ' → 修正
    Call FormInit(mdl_Mode)        ' 画面項目の初期化
    Call BtnConf(mdl_Mode)         ' 保存/削除ボタンの使用可否設定
    Me.cmb_Ukno.Enabled = True
    Me.cmb_Ukno.Requery
    Call SetFrmEnabledConf(False)  ' 画面項目の使用可否設定

    ' 受付番号使用可否設定
    Call SetUKNOBtnConf            ' 2005/07/26 ADD T.SUZUKI

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :opt_Del_MouseDown
'        機能             :削除処理(画面左上ボタン）
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opt_Del_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 2005/12/20 ADD T.SUZUKI Start
    ' [削除]ボタンが追うかされた場合の処理は、[取消]ボタンが押下されたときの
    ' 処理と同様にする
    Dim blnRTN As Boolean
    Dim lngRTN As Long

    If pblnChangeFlag_FVS500 = True Then
        '確認ﾒｯｾｰｼﾞ
        lngRTN = MsgBox("画面をクリアします。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
        If lngRTN = vbNo Then
            Me.opg_Mode.VALUE = mdl_Mode
            Exit Sub
        Else
            If bln_FormLoad = True Then
                '各変数の初期化
                strForm_Yard = ""
                strForm_Kont = ""
                strForm_UkeNo = ""
                strFrom_Yoyaku = ""  ' 2006/01/18 ADD T.SUZUKI
            End If
        End If
    End If

    mdl_Mode = P_DEL               ' → 削除
    Call FormInit(mdl_Mode)        ' 画面項目の初期化
    Call BtnConf(mdl_Mode)         ' 保存/削除ボタンの使用可否設定

    ' 既に予約受付番号が選択されていたら、予約受付トランのデータを読み込む
    If strForm_UkeNo <> "" Then
        blnRTN = GetYoukData(1)
        Call SetFrmEnabledConf(True)   ' 画面項目の使用可否設定
    Else
        Call SetFrmEnabledConf(False)  ' 画面項目の使用可否設定
    End If

    pblnChangeFlag_FVS500 = False

    Me.cmb_Ukno.Enabled = True
    Me.cmb_Ukno.Requery
'    Call SetFrmEnabledConf(False)  ' 画面項目の使用可否設定

    ' 受付番号使用可否設定
    Call SetUKNOBtnConf

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf
' 2005/12/20 ADD T.SUZUKI End

' 2005/12/20 DEL T.SUZUKI Start
'' 2005/10/06 ADD T.SUZUKI Start
'    If bln_FormLoad = True Then
'        '各変数の初期化
'        strForm_Yard = ""
'        strForm_Kont = ""
'        strForm_UkeNo = ""
'    End If
'
'    Me.Caption = pstrFormCaption
'' 2005/10/06 ADD T.SUZUKI End
'
'    mdl_Mode = P_DEL               ' → 削除
'    Call FormInit(mdl_Mode)        ' 画面項目の初期化
'    Call BtnConf(mdl_Mode)         ' 保存/削除ボタンの使用可否設定
'    Me.cmb_Ukno.Enabled = True
'    Me.cmb_Ukno.Requery
'    Call SetFrmEnabledConf(False)  ' 画面項目の使用可否設定
'
'    ' 受付番号使用可否設定
'    Call SetUKNOBtnConf  ' 2005/07/26 ADD T.SUZUKI
'
'    ' 紹介履歴ボタン使用可否設定
'    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI
' 2005/12/20 DEL T.SUZUKI End
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Delete_Click
'        機能             :削除ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Delete_Click()

    Dim lngRc   As Long

    On Error GoTo cmd_Delete_Click_Err

    '入力チェック
    If Nz(Me.cmb_Ukno.VALUE, "") = "" Then
        MsgBox "受付番号が選択されていません。", vbExclamation, P_ERROR
        GoTo cmd_Delete_Click_End
    End If

    ' 2006/06/16 Add Tajima Start
    ' 契約状態の場合は削除できない
    If Nz(Me.cmb_Yukbn.VALUE, "") = P_YUKBN_20 Then
        MsgBox "予約状態が契約のため削除はできません。", vbExclamation, P_ERROR
        GoTo cmd_Delete_Click_End
    End If
    ' 2006/06/16 Add Tajima End

    Select Case mdl_Mode
        Case P_NEW, P_SHU
            lngRc = MsgBox("削除モードでないので、削除できません。", vbExclamation, P_ERROR)
            GoTo cmd_Delete_Click_End
        Case P_DEL
            lngRc = MsgBox("削除します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
            If lngRc = vbNo Then GoTo cmd_Delete_Click_End
            'Delete処理
            If DB_Access(mdl_Mode) = False Then GoTo cmd_Delete_Click_End       ' 206/04/07 REP T.SUZUKI
    End Select

' 2005/10/06 REP T.SUZUKI Start
    If blnFVS540FLG = False Then
        lngRc = MsgBox("削除処理が終了しました。", vbInformation + vbOKOnly, P_KAKUNIN)
    End If
'    lngRc = MsgBox("削除処理が終了しました。", vbInformation + vbOKOnly, P_KAKUNIN)
' 2005/10/06 REP T.SUZUKI End

    '処理後は受付№変数をリセット
    strForm_Yard = ""
    strForm_Kont = ""
    strForm_UkeNo = ""
    strFrom_Yoyaku = ""                      ' 2006/01/18 ADD T.SUZUKI

    Me.Caption = pstrFormCaption             ' 2005/10/06 ADD T.SUZUKI
    Call subSetFVS500(P_MOVEKBN_00)          ' 2006/01/18 ADD T.SUZUKI

    '正常終了した時はﾌｫｰﾑを初期化する
    mdl_Mode = P_NEW
    Call SetFrmEnabledConf(True)
    Call FormInit(mdl_Mode)

    ' 受付番号使用可否設定
    Call SetUKNOBtnConf  ' 2005/07/26 ADD T.SUZUKI Start

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI

' 2005/10/06 ADD T.SUZUKI Start
    If blnFVS540FLG = True Then
        ' 予約一覧画面からの遷移の場合、保存処理終了後に画面を閉じる
        Call cmd_End_Click
    End If
' 2005/10/06 ADD T.SUZUKI End

    GoTo cmd_Delete_Click_End

cmd_Delete_Click_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
cmd_Delete_Click_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Enter_Click
'        機能             :保存ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Enter_Click()

    Dim strMsg  As String
    Dim lngRc   As Long
    Dim strBumon As String

    On Error GoTo cmd_Enter_Click_On_Err

    strBumon = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")

    Select Case mdl_Mode
        Case P_NEW, P_SHU
            If mdl_Mode = P_NEW Then strMsg = "新規登録処理" Else strMsg = "修正処理"
            '入力ﾁｪｯｸ
            If ChkDispData = False Then GoTo cmd_Enter_Click_End
            
            '2015/01/31 M.HONDA INS
            '新規の時のみ空数ﾁｪｯｸ
            If mdl_Mode = P_NEW Then
                'オフィスはチェックしない。
                If strBumon <> "1" Then
                    '空き件数チェック
                    If Not Chk_SPAC_AKI() Then GoTo cmd_Enter_Click_End                 'INSERT 2014/12/12 K.ISHIZAKA
                End If
            End If
            '2015/01/31 M.HONDA INS
            
            '確認ﾒｯｾｰｼﾞ
            lngRc = MsgBox("保存します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
            If lngRc = vbNo Then GoTo cmd_Enter_Click_End

            'Insert/Update処理
            If DB_Access(mdl_Mode) = False Then GoTo cmd_Enter_Click_End        ' 2006/04/07 CHG T.SUZUKI

        Case P_DEL
            lngRc = MsgBox("新規・修正時のみ保存可能です。", vbExclamation, P_KAKUNIN)
            GoTo cmd_Enter_Click_End
    End Select

' 2005/10/06 REP T.SUZUKI Start
    If blnFVS540FLG = False Then
        lngRc = MsgBox(strMsg & "が終了しました。", vbInformation + vbOKOnly, P_KAKUNIN)
    End If
'    lngRc = MsgBox(strMsg & "が終了しました。", vbInformation + vbOKOnly, P_KAKUNIN)
' 2005/10/06 REP T.SUZUKI End

    '処理後は受付№変数をリセット
    strForm_Yard = ""
    strForm_Kont = ""
    strForm_UkeNo = ""
    strFrom_Yoyaku = ""                      ' 2006/01/18 ADD T.SUZUKI
    str空検索時レンタル用途 = ""             ' 2007/06/02 add tajima

    Me.Caption = pstrFormCaption             ' 2005/10/06 ADD T.SUZUKI
    Call subSetFVS500(P_MOVEKBN_00)          ' 2006/01/18 ADD T.SUZUKI

    '正常終了した時はﾌｫｰﾑを初期化する
    mdl_Mode = P_NEW
    Call SetFrmEnabledConf(True)
    Call FormInit(mdl_Mode)

    ' 受付番号使用可否設定
    Call SetUKNOBtnConf            ' 2005/07/26 ADD T.SUZUKI

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI

    isネット予約 = False                                '2007/08/09 add tajima
    Me.lbl_Youk.Caption = P_TITTLE_YOYAKU & "：電話"    '2007/08/09 add tajima

' 2005/10/06 ADD T.SUZUKI Start
    If blnFVS540FLG = True Then
        ' 予約一覧画面からの遷移の場合、保存処理終了後に画面を閉じる
        Call cmd_End_Click
    End If
' 2005/10/06 ADD T.SUZUKI End

    GoTo cmd_Enter_Click_End

cmd_Enter_Click_On_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmd_Enter_Click_End:
    'Me.cmb_Yard.SetFocus
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Kesi_Click
'        機能             :取消ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Kesi_Click()

    Dim lngRc As Long

    '確認ﾒｯｾｰｼﾞ
    lngRc = MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN)
    If lngRc = vbNo Then Exit Sub

    '初期化
    Select Case mdl_Mode
        Case P_NEW: Call FormInit(P_NEW)
        Case P_SHU, P_DEL
' 2005/10/06 REP T.SUZUKI Start
            If strForm_UkeNo <> "" Then lngRc = GetYoukData(1)
            pblnChangeFlag_FVS500 = False ' 2005/10/06 ADD T.SUZUKI
'        Case P_SHU, P_DEL: If strForm_UkeNo <> "" Then lngRc = GetYoukData(1)
' 2005/10/06 REP T.SUZUKI End
    End Select

    Me.Caption = pstrFormCaption                                  ' 2005/10/06 ADD T.SUZUKI
    Call subSetFVS500(Format(Nz(Me.cmb_MoveKbn, 0), "00"))        ' 2006/01/18 ADD T.SUZUKI

    ' 受付番号使用可否設定
    Call SetUKNOBtnConf            ' 2005/07/26 ADD T.SUZUKI

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_End_Click
'        機能             :終了ボタン押下時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_End_Click()

' 2005/02/18 ADD T.SUZUKI Start
    Dim strMessage    As String
    Dim varDefBtn     As Variant

    ' 予約一覧画面からの遷移の場合には終了ﾒｯｾｰｼﾞの表示を行わない。
    If (pblnChangeFlag_FVS500 = False And blnFVS540FLG = True) Then GoTo cmd_End_Click_Exit

    ' 変更フラグがTrueの時は更新の確認メッセージを表示、それ以外は終了の確認メッセージを表示
    strMessage = IIf(pblnChangeFlag_FVS500 = True, pcstrEndEnterMsg, pcstrEndMsg)
    varDefBtn = vbYesNo + vbQuestion + IIf(pblnChangeFlag_FVS500 = True, vbDefaultButton2, vbDefaultButton1)

    '' 2013/06/03 M.HONDA START
    ' 確認メッセージ表示
'    If MsgBox(strMessage, varDefBtn, P_KAKUNIN) = vbNo Then
'        Exit Sub
'    End If
    '' 2013/06/03 M.HONDA END
' 2005/02/18 ADD T.SUZUKI End

cmd_End_Click_Exit:
    On Error Resume Next
    doCmd.Close acForm, Me.FormName

    Call BeforForm_Reload  ' 2005/11/02 ADD T.SUZUKI
End Sub

' 2005/07/26 ADD T.SUZUKI Start
'==============================================================================*
'
'        MODULE_NAME      :cmd_UKNO_Click
'        機能             :受付番号コンボボックスクリック時の処理
'        IN               :
'        OUT              :
'        CREATE           :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_UKNO_Click()

    Call psubSearchUKdata(Me.NAME, Me.cmb_Ukno.NAME)

    If pstrUKNO <> "" Then
        Call cmb_Ukno_AfterUpdate
        pstrUKNO = ""
    End If
End Sub
' 2005/07/26 ADD T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :cmb_No_BeforeUpdate
'        機能             :ｺﾝﾃﾅ番号BeforeUpdate時
'        IN               :2005/07/17 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_No_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo cmb_No_BeforeUpdate_Err

    vCode = Nz(Me.cmb_No, "")

    If vCode = "" Then GoTo cmb_No_BeforeUpdate_End

    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            GoTo cmb_No_BeforeUpdate_End

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_No_BeforeUpdate_End
    End Select

    'ファイルへの存在チェック
    If GetCntaData(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_No_BeforeUpdate_End
    End If

    GoTo cmb_No_BeforeUpdate_End

cmb_No_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_No_BeforeUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_No_AfterUpdate
'        機能             :ｺﾝﾃﾅ番号AfterUpdate時
'        IN               :2005/07/17 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_No_AfterUpdate()
  
  On Error GoTo Exception
  
    Me.cmb_No = Format(Me.cmb_No, "000000")
    '▼ 2006/09/26 add tajima ▼
    ' コンテナの修繕情報をゲット
    Dim strBumoCD As String
    If fncConnectKaseDB = False Then Exit Sub ' 2006/09/24 tajima add
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    Me.txt_部屋修繕 = CmReprMod.GetRoomLatestReprText(pobjKASE_DB, strBumoCD, Me.cmb_Yard, Me.cmb_No)
    Call fncDisConnectKaseDB
    '▲ 2006/09/26 add tajima ▲

' 2005/12/12 DEL T.SUZUKI Start
' ｺﾝﾃﾅ番号選択時、希望ｻｲｽﾞ、用途のﾃﾞﾌｫﾙﾄ設定を行わない
'' 2005/10/06 ADD T.SUZUKI Start
'    ' 新規モードのときのみ希望ｻｲｽﾞ、用途のﾃﾞﾌｫﾙﾄ設定を行う
'    If opg_Mode.Value = P_NEW Then
'        ' 希望ｻｲｽﾞのﾃﾞﾌｫﾙﾄ表示処理
'        Call subSetDefaultValue
'    End If
'' 2005/10/06 ADD T.SUZUKI End
' 2005/12/12 DEL T.SUZUKI End
    Exit Sub
Exception:
    Call fncDisConnectKaseDB
    Call MSZZ024_M00("cmb_No_AfterUpdate", True)
    Err.Clear
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_No_Change
'        機能             :ｺﾝﾃﾅ番号変更時
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_No_Change()
    'NZ関数追加 2005/07/17 K.MITSUNO
    strForm_Kont = Nz(Me.cmb_No, "")

    Call subSetChangeFlag  ' 2005/10/06 ADD T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_No_Enter
'        機能             :ｺﾝﾃﾅ番号ﾌｫｰｶｽ取得時
'        IN               :2005/07/17 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_No_Enter()

    If IsNull(Me.cmb_Yard) Then
        Me.cmb_Yard.SetFocus
        Exit Sub
    Else
        If Val(Me.cmb_Yard) = 0 Then Me.cmb_Yard.SetFocus
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_No_Enter
'        機能             :ｺﾝﾃﾅ番号ﾌｫｰｶｽ喪失時
'        IN               :2005/07/17 K.MITSUNO
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_No_Exit(Cancel As Integer)

    If Nz(Me.cmb_No, "") <> "" Then
        Me.cmb_No = Format(Me.cmb_No, "000000")
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_No_KeyPress
'        機能             :ｺﾝﾃﾅ番号KeyPressｲﾍﾞﾝﾄ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_No_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_No, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kokya_BeforeUpdateイベント
'        MODULE_ID        :txt_Kokya_BeforeUpdate
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kokya_BeforeUpdate(Cancel As Integer)

    If ChkTextLength(Me.txt_Kokya, 18) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kokya_Changeイベント
'        MODULE_ID        :txt_Kokya_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kokya_Change()
    Call subSetChangeFlag
End Sub
' 2005/10/06 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :cmb_Kokya_AfterUpdate
''        機能             :顧客情報AfterUpdateｲﾍﾞﾝﾄ処理(顧客ﾌﾘｶﾞﾅ)
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_Kokya_AfterUpdate()
'    cmb_KokCd.Value = ""
'    Me.cmb_KokCd.Requery
'End Sub
''==============================================================================*
''
''        MODULE_NAME      :cmb_Kokya_Changeイベント
''        MODULE_ID        :cmb_Kokya_Change
''        CREATE_DATE      :2005/10/06
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_Kokya_Change()
'    Call subSetChangeFlag
'End Sub
' 2005/10/06 DEL T.SUZUKI End
'==============================================================================*
'
'        MODULE_NAME      :cmb_KokCd_KeyPress
'        機能             :顧客情報KeyPressｲﾍﾞﾝﾄ処理(顧客ｺｰﾄﾞ)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KokCd_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.cmb_KokCd, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_KokCd_GotFocusイベント
'        MODULE_ID        :cmb_KokCd_GotFocus
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KokCd_GotFocus()

    Dim strSQL As String

    If Nz(Me.txt_Kokya, "") = "" Then
        Me.cmb_KokCd.RowSource = ""
        Exit Sub

    Else
        ' 顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ表示のSQL文作成
        strSQL = "SELECT TOP 100 Format([USER_CODE],'000000') AS 式1, "
        strSQL = strSQL & " dbo_USER_MAST.USER_NAME "
        strSQL = strSQL & "  FROM dbo_USER_MAST "
        strSQL = strSQL & " WHERE (((dbo_USER_MAST.USER_KANA) LIKE '*" & Me.txt_Kokya & "*')) "
        strSQL = strSQL & " ORDER BY Format([USER_CODE],'000000');"
    
        Me.cmb_KokCd.RowSource = strSQL
        Me.cmb_KokCd.Requery

        If Me.cmb_KokCd.ListCount = 0 Then Me.cmb_KokCd = Null
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_KokCd_BeforeUpdate
'        機能             :顧客情報BeforeUpdateｲﾍﾞﾝﾄ処理(顧客ｺｰﾄﾞ)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KokCd_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo cmb_KokCd_BeforeUpdate_Err

    If bln_FormLoad = False Then GoTo cmb_KokCd_BeforeUpdate_End
    
    vCode = Nz(Me.cmb_KokCd.VALUE, "")
    If vCode = "" Then GoTo cmb_KokCd_BeforeUpdate_End

    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            GoTo cmb_KokCd_BeforeUpdate_End
        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_KokCd_BeforeUpdate_End
        Case Else
    End Select

    'ファイルへの存在チェック
    If Chk_UserMast(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_KokCd_BeforeUpdate_End
    End If

    GoTo cmb_KokCd_BeforeUpdate_End

cmb_KokCd_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_KokCd_BeforeUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_KokCd_AfterUpdate
'        機能             :顧客情報AfterUpdateｲﾍﾞﾝﾄ処理(顧客ｺｰﾄﾞ)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KokCd_AfterUpdate()

    Dim vCode As Variant

    On Error GoTo Exception

    vCode = Nz(Me.cmb_KokCd.VALUE, "")
    If vCode = "" Then GoTo cmb_KokCd_AfterUpdate_End

    'ファイルへの存在チェック
    If Chk_UserMast(1) = False Then GoTo cmb_KokCd_AfterUpdate_End
    Me.cmb_KokCd = Format(Me.cmb_KokCd.VALUE, "000000")

    On Error Resume Next

' 2005/10/06 ADD T.SUZUKI Start
    Me.txt_Kokya.SetFocus
    Me.cmb_KokCd.SetFocus

    doCmd.Hourglass True
    doCmd.Echo False
    If fncDeleteWorkTable = False Then Exit Sub  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ削除
    If fncCreateWorkTable = False Then Exit Sub  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ作成
    Call subCreateWork                           ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙﾃﾞｰﾀ作成
    Me.FVS500_FVS500S01.Requery                  ' 2005/10/25 ADD T.SUZUKI
'    doCmd.Requery "FVS500_FVS500S01"  ' ｻﾌﾞﾌｫｰﾑ再表示  ' 2005/10/25 DEL T.SUZUKI
    doCmd.Echo True
' 2005/10/06 ADD T.SUZUKI End

' 2006/05/12 ADD tajima Start
' 名変以外で顧客コードを変えたらクリアする
    If Nz(Me.cmb_MoveKbn, "") <> "04" Then
      txt_MOTO_ACPTNO = ""
    End If
' 2006/05/12 ADD tajima End

    '▼ 2006/09/26 add tajima ▼
    ' 顧客修繕情報をゲット
    Dim strBumoCD As String
    If fncConnectKaseDB = False Then Exit Sub ' 2006/09/24 tajima add
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    Me.txt_顧客コメントクレーム = CmReprMod.GetUserLatestKotoText(pobjKASE_DB, strBumoCD, Me.cmb_KokCd)
    Call fncDisConnectKaseDB
    '▲ 2006/09/26 add tajima ▲

    GoTo cmb_KokCd_AfterUpdate_End

Exception:
    Call fncDisConnectKaseDB
    Call MSZZ024_M00("cmb_No_AfterUpdate", True)
    Err.Clear

cmb_KokCd_AfterUpdate_End:
    doCmd.Hourglass False  ' 2005/10/06 ADD T.SUZUKI
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_KokCd_Changeイベント
'        MODULE_ID        :cmb_KokCd_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_KokCd_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Ukno_BeforeUpdate
'        機能             :受付番号BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ukno_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo cmb_Ukno_BeforeUpdate_Err

    vCode = Nz(Me.cmb_Ukno, "")

    If vCode = "" Then GoTo cmb_Ukno_BeforeUpdate_End

    '入力チェック
    If ChkCmbLength(Me.cmb_Ukno, 11) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_Ukno_BeforeUpdate_End
    End If

    'ファイルへの存在チェック
    If GetYoukData(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_Ukno_BeforeUpdate_End
    End If

    GoTo cmb_Ukno_BeforeUpdate_End

cmb_Ukno_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Ukno_BeforeUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Ukno_AfterUpdate
'        機能             :受付番号AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ukno_AfterUpdate()

    Dim vCode   As Variant
    Dim lngRc   As Long

    On Error GoTo cmb_Ukno_AfterUpdate_Err

    If bln_FormLoad = False Then Exit Sub

    vCode = Nz(Me.cmb_Ukno, "")
    If vCode = "" Then
        '空の時は初期化
        Call FormInit(mdl_Mode)
        GoTo cmb_Ukno_AfterUpdate_End
    End If

    Call SetFrmEnabledConf(True)
    strForm_UkeNo = Me.cmb_Ukno.VALUE

    ' 紹介履歴ボタン使用可否設定
    Call SetSyoukaiBtnEnabledConf  ' 2005/10/06 ADD T.SUZUKI

    '予約トラン読み込み
    lngRc = GetYoukData(1)

    ' 画面の標題設定
    Me.Caption = pstrFormCaption             ' 205/10/06 ADD T.SUZUKI
'    Call subSetFVS500(Me.cmb_Yukbn)          ' 2006/01/24 ADD T.SUZUKI
'    Call subSetFVS500(P_MOVEKBN_00)          ' 2006/01/18 ADD T.SUZUKI

    GoTo cmb_Ukno_AfterUpdate_End

cmb_Ukno_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Ukno_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_Change
'        機能             :状態区分Changeｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmb_Yukbn_Change()                                                 DELETE 2005/07/17 K.MITSUNO
'    '状態区分により入力可否設定
'    Call Yukbn_EnableConf(mdl_Mode, Me.cmb_Yukbn)
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_BeforeUpdate
'        機能             :状態区分BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、状態区分名欄も空欄表示
    If ControlCrear(Me.cmb_Yukbn, Me.txt_Yukbn) = False Then Exit Sub

' 2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Yukbn

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Yukbn_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Yukbn_BeforeUpdate_Exit
    End Select
' 2005/12/19 ADD T.SUZUKI End

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_Yukbn, P_NAME_080) = "" Then   ' 2005/12/19 ADD T.SUZUKI
'    If Chk_NameBefore(Me.cmb_Yukbn, P_NAME_080) = "" Then  ' 2005/12/19 DEL T.SUZUKI
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
       GoTo cmb_Yukbn_BeforeUpdate_Exit
    End If

' 2005/12/19 ADD T.SUZUKI Start
'    ' 受付番号が未選択状態のときは処理を行わない
'    If Nz(Me.cmb_Ukno, "") = "" Then Exit Sub

    varCode = Format(varCode, "00")  ' 予約受付状態区分セット
    Select Case mdl_Mode
        Case P_NEW  ' *** 新規 ***
            ' (自動紹介 OR 自社理由 OR 取置 OR 受付)
            If (varCode <> P_YUKBN_00 And _
                varCode <> P_YUKBN_01 And _
                varCode <> P_YUKBN_02 And _
                varCode <> P_YUKBN_10) Then
                 GoTo cmb_Yukbn_BeforeUpdate_SelectErr
            End If

        Case P_SHU  ' *** 修正 ***
            Select Case pstrYUKBN
                Case P_YUKBN_00
                    ' 自動紹介
                    ' (自動紹介 OR 取置 OR 受付 OR 予約キャンセル)
                    If (varCode <> P_YUKBN_00 And _
                        varCode <> P_YUKBN_02 And _
                        varCode <> P_YUKBN_10 And _
                        varCode <> P_YUKBN_51) Then
                        GoTo cmb_Yukbn_BeforeUpdate_SelectErr
                    End If

                Case P_YUKBN_01
                    ' 自社理由
                    ' (自社理由 OR 受付 OR 予約キャンセル)
                    If (varCode <> P_YUKBN_01 And _
                        varCode <> P_YUKBN_10 And _
                        varCode <> P_YUKBN_51) Then
                        GoTo cmb_Yukbn_BeforeUpdate_SelectErr
                    End If

                Case P_YUKBN_02
                    ' 取置
                    ' (自動紹介 OR 取置 OR 受付 OR 予約キャンセル)
                    If (varCode <> P_YUKBN_00 And _
                        varCode <> P_YUKBN_02 And _
                        varCode <> P_YUKBN_10 And _
                        varCode <> P_YUKBN_51) Then
                        GoTo cmb_Yukbn_BeforeUpdate_SelectErr
                    End If

                Case P_YUKBN_10
                    ' 受付
                    ' (受付 OR 受付キャンセル)
                    If (varCode <> P_YUKBN_10 And _
                        varCode <> P_YUKBN_53) Then
                        GoTo cmb_Yukbn_BeforeUpdate_SelectErr
                    End If

                Case P_YUKBN_20
                    ' 契約
                    ' (契約のみ)
                    If varCode <> P_YUKBN_20 Then GoTo cmb_Yukbn_BeforeUpdate_SelectErr

                Case P_YUKBN_51
                    ' 予約キャンセル
                    ' (予約キャンセルのみ)
                    If varCode <> P_YUKBN_51 Then GoTo cmb_Yukbn_BeforeUpdate_SelectErr

                Case P_YUKBN_52
                    ' 取置キャンセル
                    ' (自動紹介 OR 予約キャンセル OR 取置キャンセル)
                    If (varCode <> P_YUKBN_00 And _
                        varCode <> P_YUKBN_51 And _
                        varCode <> P_YUKBN_52) Then
                        GoTo cmb_Yukbn_BeforeUpdate_SelectErr
                    End If

                Case P_YUKBN_53
                    ' 受付キャンセル
                    ' (受付キャンセルのみ)
                    If varCode <> P_YUKBN_53 Then GoTo cmb_Yukbn_BeforeUpdate_SelectErr
                Case Else
            End Select

        Case P_DEL  ' *** 削除 ***
            ' (選択している予約受付状態区分のみ)
            If varCode <> pstrYUKBN Then GoTo cmb_Yukbn_BeforeUpdate_SelectErr
    End Select
' 2005/12/19 ADD T.SUZUKI End

cmb_Yukbn_BeforeUpdate_Exit:
    Exit Sub

' 2005/12/19 ADD T.SUZUKI Start
cmb_Yukbn_BeforeUpdate_SelectErr:
    ' 予約状態指定エラーメッセージ
    MsgBox pcstrSelErrMsg, vbExclamation, P_ERROR
    Cancel = True
    SendKeys "{ESC}"
' 2005/12/19 ADD T.SUZUKI End
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_AfterUpdate
'        機能             :状態区分AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_AfterUpdate()

    On Error GoTo cmb_Yukbn_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_Yukbn, Me.txt_Yukbn, P_NAME_080) = False Then GoTo cmb_Yukbn_AfterUpdate_End

    '状態区分により入力可否設定                                                 変更 2005/07/17 K.MITSUNO
    Call Yukbn_EnableConf(mdl_Mode, Me.cmb_Yukbn)

' 2006/01/18 ADD T.SUZUKI Start
    If Me.cmb_Yukbn = P_YUKBN_10 Then
        Call subSetTKDATE
    End If
' 2006/01/18 ADD T.SUZUKI End

    GoTo cmb_Yukbn_AfterUpdate_End

cmb_Yukbn_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Yukbn_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_KeyPress
'        機能             :状態区分KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_Yukbn, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yukbn_Changeイベント
'        MODULE_ID        :cmb_Yukbn_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yukbn_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Shunou_KeyPress
'        機能             :収納物KeyPressｲﾍﾞﾝﾄ処理
'        IN               :2005/07/24
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Shunou_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Shunou, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Shunou_BeforeUpdate
'        機能             :収納物BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :2005/07/24
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Shunou_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、顧客区分名欄も空欄表示
    If ControlCrear(Me.cmb_Shunou, Me.txt_Butu) = False Then Exit Sub

'2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Shunou

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Shunou_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Shunou_BeforeUpdate_Exit
    End Select
'2005/12/19 ADD T.SUZUKI End

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_Shunou, P_NAME_091) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_Shunou_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Shunou_AfterUpdate
'        機能             :収納物AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :2005/07/24
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Shunou_AfterUpdate()

    On Error GoTo cmb_Shunou_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_Shunou, Me.txt_Butu, P_NAME_091) = False Then GoTo cmb_Shunou_AfterUpdate_End

    GoTo cmb_Shunou_AfterUpdate_End

cmb_Shunou_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Shunou_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Shunou_Changeイベント
'        MODULE_ID        :cmb_Shunou_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Shunou_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kkbn_BeforeUpdate
'        機能             :顧客区分BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kkbn_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、顧客区分名欄も空欄表示
    If ControlCrear(Me.cmb_Kkbn, Me.txt_Kkbn) = False Then Exit Sub

'2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Kkbn

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Kkbn_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Kkbn_BeforeUpdate_Exit
    End Select
'2005/12/19 ADD T.SUZUKI End

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_Kkbn, P_NAME_020) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_Kkbn_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kkbn_AfterUpdate
'        機能             :顧客区分AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kkbn_AfterUpdate()

    On Error GoTo cmb_Kkbn_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_Kkbn, Me.txt_Kkbn, P_NAME_020) = False Then GoTo cmb_Kkbn_AfterUpdate_End

    GoTo cmb_Kkbn_AfterUpdate_End

cmb_Kkbn_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Kkbn_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kkbn_KeyPress
'        機能             :顧客区分KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kkbn_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_Kkbn, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kkbn_Changeイベント
'        MODULE_ID        :cmb_Kkbn_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kkbn_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Step_BeforeUpdate
'        機能             :段区分BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Step_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、段区分名欄も空欄表示
    If ControlCrear(Me.cmb_Step, Me.txt_STEP) = False Then Exit Sub

'2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Step

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Step_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Step_BeforeUpdate_Exit
    End Select
'2005/12/19 ADD T.SUZUKI End

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_Step, P_NAME_090) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_Step_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Step_KeyPress
'        機能             :段区分KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Step_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Step, KeyAscii, 1)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Step_AfterUpdate
'        機能             :段区分AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Step_AfterUpdate()

    On Error GoTo cmb_Step_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_Step, Me.txt_STEP, P_NAME_090) = False Then GoTo cmb_Step_AfterUpdate_End

    'INSERT 2020/01/01 Y.WADA Start
    txt_PricFrom = cmb_Step.Column(2)
    txt_PricTo = cmb_Step.Column(3)
    'INSERT 2020/01/01 Y.WADA End

    GoTo cmb_Step_AfterUpdate_End

cmb_Step_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
cmb_Step_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Step_Changeイベント
'        MODULE_ID        :cmb_Step_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Step_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Youto_BeforeUpdate
'        機能             :用途 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        INSERT         　：2005/06/25 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Youto_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、サイズFrom名欄も空欄表示
    If ControlCrear(Me.cmb_Youto, Me.txt_Youto) = False Then Exit Sub

'2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Youto

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Youto_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Youto_BeforeUpdate_Exit
    End Select
'2005/12/19 ADD T.SUZUKI End

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_Youto, P_NAME_086) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_Youto_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Youto_AfterUpdate
'        機能             :用途 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        INSERT         　：2005/06/25 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Youto_AfterUpdate()

    On Error GoTo cmb_Youto_AfterUpdate_Err

'DELETE 2020/01/01 Y.WADA Start
'    'ファイルへの存在チェック
'    If Chk_NameAfter(Me.cmb_Youto, Me.txt_Youto, P_NAME_086) = False Then GoTo cmb_Youto_AfterUpdate_End
'DELETE 2020/01/01 Y.WADA End

    'INSERT 2020/01/01 Y.WADA Start
    'ファイルへの存在チェック
    Call Chk_NameAfter(Me.cmb_Youto, Me.txt_Youto, P_NAME_086)
    
    Dim dummy_i As Integer
    Call make_cmb_Kibou
    cmb_Kibou.VALUE = ""
    Call cmb_Kibou_BeforeUpdate(dummy_i)
    Call cmb_Kibou_AfterUpdate
    'INSERT 2020/01/01 Y.WADA End

    GoTo cmb_Youto_AfterUpdate_End

cmb_Youto_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Youto_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Youto_KeyPress
'        機能             :用途 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        INSERT         　：2005/06/25 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Youto_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_Youto, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Youto_Changeイベント
'        MODULE_ID        :cmb_Youto_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Youto_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kibou_BeforeUpdate
'        機能             :希望ｻｲｽﾞ BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        INSERT         　：2005/06/25 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kibou_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、ｻｲｽﾞTo名欄も空欄表示
    If ControlCrear(Me.cmb_Kibou, Me.txt_SizeNm) = False Then
        Me.txt_SyzeF = ""
        Me.txt_SyzeT = ""
        Exit Sub
    End If

'2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Kibou

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Kibou_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Kibou_BeforeUpdate_Exit
    End Select
'2005/12/19 ADD T.SUZUKI End

    If Chk_NameBefore(Me.cmb_Kibou, P_NAME_087) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_Kibou_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kibou_AfterUpdate
'        機能             :希望ｻｲｽﾞ AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        INSERT         　：2005/06/25 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kibou_AfterUpdate()

    On Error GoTo cmb_Kibou_AfterUpdate_Err

'DELETE 2020/01/01 Y.WADA Start
'    'ファイルへの存在チェック
'    If Chk_NameAfter(Me.cmb_Kibou, Me.txt_SyzeF, P_NAME_087, Me.txt_SyzeT, Me.txt_SizeNm) = False Then
'        GoTo cmb_Kibou_AfterUpdate_End
'    End If
'DELETE 2020/01/01 Y.WADA End

    'INSERT 2020/01/01 Y.WADA Start
    'ファイルへの存在チェック
    Call Chk_NameAfter(Me.cmb_Kibou, Me.txt_SyzeF, P_NAME_087, Me.txt_SyzeT, Me.txt_SizeNm)
    
    Dim dummy_i As Integer
    Call make_cmb_Step
    cmb_Step.VALUE = ""
    txt_PricFrom.VALUE = ""
    txt_PricTo.VALUE = ""
    Call cmb_Step_BeforeUpdate(dummy_i)
    Call cmb_Step_AfterUpdate
    'INSERT 2020/01/01 Y.WADA End

    GoTo cmb_Kibou_AfterUpdate_End

cmb_Kibou_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Kibou_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kibou_KeyPress
'        機能             :希望ｻｲｽﾞ KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        INSERT         　：2005/06/25 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kibou_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Kibou, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Kibou_Changeイベント
'        MODULE_ID        :cmb_Kibou_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Kibou_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kibousu_KeyPress
'        機能             :希望本数 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kibousu_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_Kibousu, KeyAscii, 2)  ' 2005/12/21 ADD T.SUZUKI
'    Call CmKeyPress_For_Num(Me.txt_Kibousu, KeyAscii, 2)  ' 2005/12/21 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kibousu_Changeイベント
'        MODULE_ID        :txt_Kibousu_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kibousu_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kibousu_BeforeUpdateイベント
'        MODULE_ID        :txt_Kibousu_BeforeUpdate
'        CREATE_DATE      :2005/12/21
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kibousu_BeforeUpdate(Cancel As Integer)

' 2005/12/21 ADD T.SUZUKI Start
    ' 入力値チェック
    If Nz(Me.txt_Kibousu, "") = pcintHonsu Then
        MsgBox pcstrHonsuErrMsg, vbExclamation, P_KAKUNIN
        GoTo Input_Err
    End If
' 2005/12/21 ADD T.SUZUKI End

    If ChkTextLength(Me.txt_Kibousu, 2) = False Then
        GoTo Input_Err     ' 2005/12/21 ADD T.SUZUKI
'        Cancel = True     ' 2005/12/21 DEL T.SUZUKI
'        SendKeys "{ESC}"  ' 2005/12/21 DEL T.SUZUKI
    End If

    Exit Sub

Input_Err:
    ' 入力エラー時
    Cancel = True
    SendKeys "{ESC}"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :chk_Kinrin_AfterUpdateイベント
'        MODULE_ID        :chk_Kinrin_AfterUpdate
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_Kinrin_AfterUpdate()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :chk_Kinrin_BeforeUpdateイベント
'        MODULE_ID        :chk_Kinrin_BeforeUpdate
'        CREATE_DATE      :2007/04/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk_Kinrin_BeforeUpdate(Cancel As Integer)
    
    If chk_Kinrin.VALUE = True And _
        str営業開始日 > Format$(DATE, "YYYY/MM/DD") Then
    
        If MsgBox("新規オープンヤードですが近隣紹介を行いますか？", _
                  vbQuestion + vbYesNo + vbDefaultButton2, P_KAKUNIN) = vbNo Then
            Cancel = True
            SendKeys "{ESC}"
            Exit Sub
        End If
    
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricFrom_BeforeUpdate
'        機能             :使用料Fromｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricFrom_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_PricFrom, 9) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricFrom_GotFocusイベント
'        MODULE_ID        :txt_PricFrom_GotFocus
'        CREATE_DATE      :2005/12/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricFrom_GotFocus()
    Me.txt_PricFrom.Format = pcstrfmtFIXED
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricFrom_LostFocusイベント
'        MODULE_ID        :txt_PricFrom_LostFocus
'        CREATE_DATE      :2005/12/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricFrom_LostFocus()
    Me.txt_PricFrom.Format = pcstrfmtSTANDARD
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricFrom_KeyPress
'        機能             :使用料From KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricFrom_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_PricFrom, KeyAscii, 9)  ' 2005/12/21 ADD T.SUZUKI
'    Call CmKeyPress_For_Num(Me.txt_PricFrom, KeyAscii, 9)  ' 2005/12/21 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricFrom_Changeイベント
'        MODULE_ID        :txt_PricFrom_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricFrom_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricTo_BeforeUpdate
'        機能             :使用料Toｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricTo_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_PricTo, 9) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricTo_GotFocusイベント
'        MODULE_ID        :txt_PricTo_GotFocus
'        CREATE_DATE      :2005/12/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricTo_GotFocus()
    Me.txt_PricTo.Format = pcstrfmtFIXED
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricTo_LostFocusイベント
'        MODULE_ID        :txt_PricTo_LostFocus
'        CREATE_DATE      :2005/12/12
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricTo_LostFocus()
    Me.txt_PricTo.Format = pcstrfmtSTANDARD
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricTo_KeyPress
'        機能             :使用料To KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricTo_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.txt_PricTo, KeyAscii, 9)  ' 2005/12/21 ADD T.SUZUKI
'    Call CmKeyPress_For_Num(Me.txt_PricTo, KeyAscii, 9)  ' 2005/12/21 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_PricTo_Changeイベント
'        MODULE_ID        :txt_PricTo_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_PricTo_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_KeyPress
'        機能             :受付担当 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_Tantoc, KeyAscii, 3)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_BeforeUpdate
'        機能             :受付担当 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合担当者名欄も空欄表示
    If ControlCrear(Me.cmb_Tantoc, Me.txt_Tanton) = False Then Exit Sub

    '選択された担当者データの存在チェック
    If GetTantoData(Nz(Me.cmb_Tantoc, "")) = False Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_AfterUpdate
'        機能             :受付担当 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_AfterUpdate()

    Dim vCode   As Variant
    Dim vTanto  As Variant

    On Error GoTo cmb_Tantoc_AfterUpdate_Err

    vCode = Nz(Me.cmb_Tantoc, "")
    If vCode = "" Then GoTo cmb_Tantoc_AfterUpdate_End

    'ファイルへの存在チェック
    If GetTantoData(vCode, vTanto) = False Then GoTo cmb_Tantoc_AfterUpdate_End

    Me.txt_Tanton = vTanto
    Me.cmb_Tantoc = Format$(Me.cmb_Tantoc, "000")

    GoTo cmb_Tantoc_AfterUpdate_End

cmb_Tantoc_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Tantoc_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Tantoc_Changeイベント
'        MODULE_ID        :cmb_Tantoc_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Tantoc_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Ukkbn_KeyPress
'        機能             :受付区分 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ukkbn_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_Ukkbn, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Ukkbn_BeforeUpdate
'        機能             :受付区分 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ukkbn_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、受付区分名欄も空欄表示
    If ControlCrear(Me.cmb_Ukkbn, Me.txt_Ukkbn) = False Then Exit Sub

'2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Ukkbn

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Ukkbn_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Ukkbn_BeforeUpdate_Exit
    End Select
'2005/12/19 ADD T.SUZUKI End

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_Ukkbn, P_NAME_070) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_Ukkbn_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Ukkbn_AfterUpdate
'        機能             :受付区分 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ukkbn_AfterUpdate()

    On Error GoTo cmb_Ukkbn_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_Ukkbn, Me.txt_Ukkbn, P_NAME_070) = False Then GoTo cmb_Ukkbn_AfterUpdate_End

    'INSERT 2019/03/06 Y.WADA Start
    If isAfterUpdate = True Then
        If cmb_Ukkbn.VALUE = "81" Then  '他部署からの移動
            chk_Tabusyo.VALUE = True
            cmb_Ukkbn.Enabled = False
        Else
            chk_Tabusyo.VALUE = False
            cmb_Ukkbn.Enabled = True
        End If
    End If
    'INSERT 2019/03/06 Y.WADA End
    
    GoTo cmb_Ukkbn_AfterUpdate_End

cmb_Ukkbn_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Ukkbn_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Ukkbn_Changeイベント
'        MODULE_ID        :cmb_Ukkbn_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Ukkbn_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_FVS400_Clickイベント
'        MODULE_ID        :cmd_FVS400_Click
'        CREATE_DATE      :2006/02/07
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_FVS400_Click()

    Dim strUkno       As String
    Dim intAns        As Integer
    Dim strMotoAcptNo As String
    Dim strRcptNo     As String
    Dim strErrMsg     As String
    Dim blnAki        As Boolean
    Dim strBumon      As String

    ' 変更フラグがTrueの時は確認メッセージを出す
    If pblnChangeFlag_FVS500 = True Then
        intAns = MsgBox(pcstrUPDMsg & vbCrLf & pcstrUkeErrMsg01, vbQuestion + vbYesNo, P_KAKUNIN)
        If intAns = vbYes Then
            '入力ﾁｪｯｸ
            If ChkDispData = False Then Exit Sub
            
            'オフィスの場合にはチェックしない。
            strBumon = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1")
            If strBumon <> "1" Then
            
                '空き件数チェック
                blnAki = True
                '2015/02/06 M.HONDA UPD
                '部屋が変わらない移動の場合にはチェックしない
                If cmb_MoveKbn = "04" Or cmb_MoveKbn = "05" Or cmb_MoveKbn = "06" Or cmb_MoveKbn = "07" Or cmb_MoveKbn = "10" Then
                    blnAki = False
                End If
                '取置からの受付の場合にチェックしない
                If pstrYUKBN = "02" Then
                    blnAki = False
                End If
                '2015/02/06 M.HONDA UPD
                           
                '2015/08/05 M.HONDA INS
                 'ネットの場合にはチェックしない
                If cmb_発生区分 = "01" Then
                    blnAki = False
                End If
                '2015/08/05 M.HONDA INS
                
                If blnAki = True Then
                    If Not Chk_SPAC_AKI() Then Exit Sub                                 'INSERT 2014/12/12 K.ISHIZAKA
                End If
            End If

            'Insert/Update処理
            If DB_Access(mdl_Mode) = False Then Exit Sub

' 2006/02/18 ADD T.SUZUKI Start
            If mdl_Mode = P_NEW Then
                ' 修正モードを選択状態にする
                strUkno = Me.txt_Ukno
                Call opt_Correction_MouseDown(0, 0, 0, 0)

                Me.cmb_Ukno = strUkno
                Call cmb_Ukno_AfterUpdate
            End If
' 2006/02/18 ADD T.SUZUKI End

            ' 予約受付状態区分ｺﾝﾎﾞﾎﾞｯｸｽのSQL作成
            Call subCreate_YukbnSQL(mdl_Mode, Nz(Me.cmb_Yukbn, ""))

            ' 予約ご紹介トランの件数設定処理
            If fncGetCountINTR_TRAN(IIf(Nz(Me.cmb_Ukno, "") = "", Me.txt_Ukno, Me.cmb_Ukno)) = False Then Exit Sub
        Else
            Exit Sub
        End If

        pblnChangeFlag_FVS500 = False  ' 変更フラグ(OFF)
        Me.Caption = pstrFormCaption   ' 標題設定
    End If

    ' データチェック
    ' 受付入力画面へ遷移する条件
    ' (予約受付状態区分 = 10(受付) または
    '  予約受付状態区分 > 10(受付) かつ 受付番号 <> "")
    If Nz(strUkno, "") = "" Then
        strUkno = IIf(Nz(Me.cmb_Ukno, "") = "", Me.txt_Ukno, Me.cmb_Ukno)
    End If

    ' 受付入力画面への遷移可否チェック
    If dncCheckUketsuke(strUkno, strMotoAcptNo, strRcptNo, strErrMsg) = False Then
        ' エラーメッセージ出力
        If Nz(strErrMsg, "") <> "" Then
            MsgBox strErrMsg, vbExclamation, P_ERROR
        End If
        Exit Sub
    End If

    ' 画面コントロールの無効化
    Call subFormLockControl(False)

    ' 受付入力画面を表示(パラメータ：画面ID、予約番号、受付番号、元受注契約番号)
    doCmd.OpenForm "FVS400", , , , , , P_PROG_ID & "," & strUkno & "," & strRcptNo & "," & strMotoAcptNo
End Sub
' 2006/02/07 DEL T.SUZUKI Start
''==============================================================================*
''
''        MODULE_NAME      :cmd_Excel_Click
''        機能             :ヘルパー連動
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmd_Excel_Click()
'
'    Dim xlApp               As Object
'    Dim strFile             As String
'    Dim strXlPath           As String
'    Dim strBumoc            As String
'
'    On Error GoTo Err_Handle
'
'    strBumoc = UCase(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY = 1"))
'    Select Case strBumoc
'    Case "1"
'        strXlPath = "\\Server03\cntn\set\"
'        strFile = "オフィス用4.3.xls"
'    Case "T"
'        strXlPath = "\\Server03\cntn\set\"
'        strFile = "横浜鋼管4.11.xls"
'    Case "8"
'        strXlPath = "\\Server03\コンテナ部\受付管理sys\"
'        strFile = "トランク用ヘルパーSYS.xls"
'    Case "H"
'        strXlPath = "\\Server03\コンテナ部\受付管理sys\"
'        'If InStr(Nz(Me.txt_Youto), "バイク") > 0 Then
'        '    strFile = "バイク用ヘルパーVr4.7.xls"
'        'Else
'            strFile = "コンテナ用ヘルパーSYS.xls"
'        'End If
'    Case Else
'        Exit Sub
'    End Select
'
'    Set xlApp = CreateObject("Excel.Application")
'    xlApp.workbooks.Open (strXlPath & strFile)
'    With xlApp.workbooks(strFile).Worksheets("基本入力")
'        .Range("ヤード番号").Value = Nz(Me.cmb_Yard, "")
'        .Range("ボックス番号").Value = Nz(Me.cmb_No, "")
'        .Range("顧客コード").Value = Nz(Me.cmb_KokCd, "")
'        .Range("フリガナ").Value = Nz(Me.txt_Kana, "")
'        .Range("契約者名").Value = Nz(Me.txt_Name, "")
'        .Range("郵便番号").Value = Nz(Me.txt_Yubin, "")
'        .Range("住所1").Value = Nz(Me.txt_Adr1, "")
'        .Range("住所2").Value = Nz(Me.txt_Adr2, "") & Nz(Me.txt_Adr3, "")
'        .Range("TEL").Value = Nz(Me.txt_Tel, "")
'        .Range("FAX").Value = Nz(Me.txt_Fax, "")
'        .Range("携帯").Value = Nz(Me.txt_Keitai, "")
'        If (strBumoc = "8") Or (strBumoc = "H") Then
'            .Range("契約代表者名").Value = Nz(Me.txt_Tanm, "")
'            .Range("備考").Value = Nz(Me.txt_Biko, "")
'        End If
'    End With
'    xlApp.Visible = True
'    Exit Sub
'
'Err_Handle:
'    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
'    If Not xlApp Is Nothing Then
'        xlApp.Visible = True
'    End If
'End Sub
' 2006/02/07 DEL T.SUZUKI End

'==============================================================================*
'
'        MODULE_NAME      :cmd_Syoukai_Clickイベント
'        MODULE_ID        :cmd_Syoukai_Click
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Syoukai_Click()

    Dim intAns As Integer

' 2005/12/21 ADD T.SUZUKI Start
    ' 予約受付番号が未選択状態の場合、処理を抜ける
    If Nz(Me.cmb_Ukno, "") = "" Then
        MsgBox "受付番号が選択されていません。", vbExclamation, P_ERROR
        Exit Sub
    End If
' 2005/12/21 ADD T.SUZUKI End

    ' 変更フラグがTrueの時は確認メッセージを出す
    If pblnChangeFlag_FVS500 = True Then
        intAns = MsgBox(pcstrUPDMsg & vbCrLf & pcstrKAKUMsg, vbQuestion + vbYesNo, P_KAKUNIN)
        If intAns = vbYes Then
            '入力ﾁｪｯｸ
            If ChkDispData = False Then Exit Sub

            'Insert/Update処理
            If DB_Access(mdl_Mode) = False Then Exit Sub

' 2005/12/20 ADD T.SUZUKI Start
            ' 予約ご紹介トランの件数設定処理
            If fncGetCountINTR_TRAN(Me.cmb_Ukno) = False Then Exit Sub
' 2005/12/20 ADD T.SUZUKI Start
        Else
            Exit Sub
        End If

        pblnChangeFlag_FVS500 = False  ' 変更フラグ(OFF)
        Me.Caption = pstrFormCaption   ' 標題設定
    End If

    ' 画面コントロールの無効化
    Call subFormLockControl(False)

    ' 予約ご紹介画面を表示(パラメータ：プログラムID、予約受付番号)
    doCmd.OpenForm "FVS550", , , , , , P_PROG_ID & "," & Me.cmb_Ukno
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Renkbn_KeyPress
'        機能             :連絡区分 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Renkbn_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_Renkbn, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Renkbn_AfterUpdate
'        機能             :連絡区分 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Renkbn_AfterUpdate()

    On Error GoTo cmb_Renkbn_BeforeUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_Renkbn, Me.txt_Renkbn, P_NAME_085) = False Then GoTo cmb_Renkbn_BeforeUpdate_End

    GoTo cmb_Renkbn_BeforeUpdate_End

cmb_Renkbn_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Renkbn_BeforeUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Renkbn_BeforeUpdate
'        機能             :連絡区分 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Renkbn_BeforeUpdate(Cancel As Integer)

    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合、連絡区分名欄も空欄表示
    If ControlCrear(Me.cmb_Renkbn, Me.txt_Renkbn) = False Then Exit Sub

'2005/12/19 ADD T.SUZUKI Start
    Dim varCode As Variant

    varCode = Me.cmb_Renkbn

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Renkbn_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Renkbn_BeforeUpdate_Exit
    End Select
'2005/12/19 ADD T.SUZUKI End

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_Renkbn, P_NAME_085) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_Renkbn_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Renkbn_Changeイベント
'        MODULE_ID        :cmb_Renkbn_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Renkbn_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_BeforeUpdate
'        機能             :ヤード番号 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_BeforeUpdate(Cancel As Integer)

    Dim vCode As Variant

    On Error GoTo cmb_Yard_BeforeUpdate_Err

    vCode = Nz(Me.cmb_Yard, "")

    If vCode = "" Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_Yard_BeforeUpdate_End
    End If

    '入力チェック
    Select Case CmbCheck_Num(vCode, 6)
        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Yard_BeforeUpdate_End
        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_Yard_BeforeUpdate_End
        Case Else
    End Select

    'ファイルへの存在チェック
    If GetYardData(0) = False Then
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_Yard_BeforeUpdate_End
    End If

    GoTo cmb_Yard_BeforeUpdate_End

cmb_Yard_BeforeUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Yard_BeforeUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_AfterUpdate
'        機能             :ヤード番号 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_AfterUpdate()

    Dim vCode   As Variant
    Dim lngRc   As Long

    On Error GoTo cmb_Yard_AfterUpdate_Err

    If bln_FormLoad = False Then Exit Sub

    vCode = Nz(Me.cmb_Yard, "")
    If vCode = "" Then GoTo cmb_Yard_AfterUpdate_End

    Me.cmb_Yard = Format(Me.cmb_Yard, "000000")

    'ファイルへの存在チェック
    lngRc = GetYardData(1)

' 2005/12/12 DEL T.SUZUKI Start
'' 2005/10/06 ADD T.SUZUKI Start
'    ' ﾔｰﾄﾞｺｰﾄﾞが選択された後は無条件にｺﾝﾃﾅ番号を使用可能とする
'    cmb_No.Enabled = True
'' 2005/10/06 ADD T.SUZUKI End
' 2005/12/12 DEL T.SUZUKI End

    cmb_No = ""
    cmb_No.Requery

    'INSERT 2020/01/01 Y.WADA Start
    Dim dummy_i As Integer
    Call make_cmb_Youto
    cmb_Youto.VALUE = ""
    Call cmb_Youto_BeforeUpdate(dummy_i)
    Call cmb_Youto_AfterUpdate
    'INSERT 2020/01/01 Y.WADA End
    
    GoTo cmb_Yard_AfterUpdate_End

cmb_Yard_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_Yard_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_KeyPress
'        機能             :ヤード番号 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_Yard, KeyAscii, 6)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_Yard_Changeイベント
'        MODULE_ID        :cmb_Yard_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_Yard_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Biko_BeforeUpdate
'        機能             :備考 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Biko_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Biko, 60) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Biko_Changeイベント
'        MODULE_ID        :txt_Biko_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Biko_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_USERCHK_Click
'        機能             :cmd_USERCHK_Click イベント
'        IN               :
'        OUT              :
'        CREATE           :2006/02/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_USERCHK_Click()

    Dim strUserCd  As String
    Dim strRenraku As String

    ' 検索前の顧客コードを退避
    strUserCd = Nz(Me.cmb_KokCd, "")

    ' 既存顧客検索画面表示
    If Nz(Me.txt_Tel, "") <> "" Then
        strRenraku = Me.txt_Tel

    ElseIf Nz(Me.txt_Keitai, "") <> "" Then
        strRenraku = Me.txt_Keitai

    ElseIf Nz(Me.txt_Fax, "") <> "" Then
        strRenraku = Me.txt_Fax
    End If

    ' 既存顧客検索処理
    Call ControlFVS560(Me.NAME, Me.cmb_KokCd.NAME, Nz(Me.txt_Kana, ""), strRenraku)

    If strUserCd <> Nz(Me.cmb_KokCd, "") Then
        Me.opg_Kokya.VALUE = 2
        Call opg_Kokya_AfterUpdate  'キーボード対応 2006/06/19 Rep tajima
        Call cmb_KokCd_AfterUpdate
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :opg_Kokya_AfterUpdateイベント
'        MODULE_ID        :opg_Kokya_AfterUpdate
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opg_Kokya_AfterUpdate()
    Call subSetChangeFlag
    
    ' キーボード対応 2006/06/19 Add tajima Start
    Select Case Me.opg_Kokya.VALUE
        Case 1  '新規選択
            Call KokClear
            Call SetOptConf(True)
        Case 2  '既存選択
            Call SetOptConf(False)
    End Select
    ' キーボード対応 2006/06/19 Add tajima End
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Butu_BeforeUpdate
'        機能             :予定収納物 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Butu_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Butu, 60) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Butu_Changeイベント
'        MODULE_ID        :txt_Butu_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Butu_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kudate_BeforeUpdate
'        機能             :契約受付日 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kudate_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Kudate, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kudate_AfterUpdate
'        機能             :契約受付日 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kudate_AfterUpdate()
    Me.txt_Kudate = Format(VarChangeDate, "yyyy/mm/dd")  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_Kudate = Check_Date(Me.txt_Kudate)           ' 2006/01/05 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kudate_KeyPress
'        機能             :契約受付日 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kudate_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Kudate, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kudate_GotFocus
'        機能             :契約受付日 GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kudate_GotFocus()
    Call Date_TextFormat(Me.txt_Kudate, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kudate_LostFocus
'        機能             :契約受付日 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kudate_LostFocus()

    Dim lngToriDay  As Long

    If Nz(Me.txt_Kudate, "") <> "" Then
        Call Date_TextFormat(Me.txt_Kudate, 1)

        ' 取置期限日の設定
        Call subSetTKDATE  ' 2005/12/12 ADD T.SUZUKi
' 2005/12/12 ADD T.SUZUKi End
' 2005/12/12 DEL T.SUZUKi Start
'        '取置き日数を足して取置き期限日に日付を自動でセット(ｺﾝﾄﾛｰﾙﾏｽﾀより取得)
'        lngToriDay = DLookup("CONT_TORI_DD", "dbo_CONT_MAST", "")
'        Me.txt_Tkdate = DateAdd("d", lngToriDay, Me.txt_Kudate)
' 2005/12/12 DEL T.SUZUKi End
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kudate_Changeイベント
'        MODULE_ID        :txt_Kudate_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kudate_Change()
    Call subSetChangeFlag
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_Ytdate_BeforeUpdate
'        機能             :予約取消日 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Ytdate_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Ytdate, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Ytdate_AfterUpdate
'        機能             :予約取消日 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Ytdate_AfterUpdate()
    Me.txt_Ytdate = Format(VarChangeDate, "yyyy/mm/dd")  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_Ytdate = Check_Date(Me.txt_Ytdate)           ' 2006/01/05 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Ytdate_KeyPress
'        機能             :予約取消日 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Ytdate_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Ytdate, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Ytdate_GotFocus
'        機能             :予約取消日 GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Ytdate_GotFocus()
    Call Date_TextFormat(Me.txt_Ytdate, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Ytdate_LostFocus
'        機能             :予約取消日 LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Ytdate_LostFocus()
    If Nz(Me.txt_Ytdate, "") <> "" Then
        Call Date_TextFormat(Me.txt_Ytdate, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Ytdate_Changeイベント
'        MODULE_ID        :txt_Ytdate_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Ytdate_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kkdate_BeforeUpdate
'        機能             :希望開始日 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kkdate_BeforeUpdate(Cancel As Integer)
    If dateCheck(Me.txt_Kkdate, Cancel) = False Then Exit Sub '2006/09/26 rep tajima
    
    '入力内容が日付として問題無ければ、営業期間内かのチェックを行う
    '▼ 2006/09/26 add tajima ▼
    If Nz(VarChangeDate, "") <> "" And _
      (str営業開始日 > Format$(VarChangeDate, "YYYY/MM/DD") Or _
       str営業終了日 < Format$(VarChangeDate, "YYYY/MM/DD")) Then
        MsgBox "希望開始日は営業期間内としてください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
    '▲ 2006/09/26 add tajima ▲
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kkdate_AfterUpdate
'        機能             :希望開始日 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kkdate_AfterUpdate()
    Me.txt_Kkdate = Format(VarChangeDate, "yyyy/mm/dd")  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_Kkdate = Check_Date(Me.txt_Kkdate)           ' 2006/01/05 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kkdate_KeyPress
'        機能             :希望開始日 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kkdate_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Kkdate, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kkdate_GotFocus
'        機能             :希望開始日 GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kkdate_GotFocus()
    Call Date_TextFormat(Me.txt_Kkdate, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kkdate_LostFocus
'        機能             :希望開始日 LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kkdate_LostFocus()
    If Nz(Me.txt_Kkdate, "") <> "" Then
        Call Date_TextFormat(Me.txt_Kkdate, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kkdate_Changeイベント
'        MODULE_ID        :txt_Kkdate_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kkdate_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tadate_BeforeUpdate
'        機能             :お客様問合日 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tadate_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Tadate, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tadate_AfterUpdate
'        機能             :お客様問合日 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tadate_AfterUpdate()
    Me.txt_Tadate = Format(VarChangeDate, "yyyy/mm/dd")  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_Tadate = Check_Date(Me.txt_Tadate)           ' 2006/01/05 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tadate_KeyPress
'        機能             :お客様問合日 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tadate_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Tadate, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tadate_GotFocus
'        機能             :お客様問合日 GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tadate_GotFocus()
    Call Date_TextFormat(Me.txt_Tadate, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tadate_LostFocus
'        機能             :お客様問合日 LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tadate_LostFocus()
    If Nz(Me.txt_Tadate, "") <> "" Then
        Call Date_TextFormat(Me.txt_Tadate, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tadate_Changeイベント
'        MODULE_ID        :txt_Tadate_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tadate_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tkdate_BeforeUpdate
'        機能             :取置き期限日 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tkdate_BeforeUpdate(Cancel As Integer)
    Call dateCheck(Me.txt_Tkdate, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tkdate_AfterUpdate
'        機能             :取置き期限日 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tkdate_AfterUpdate()
    Me.txt_Tkdate = Format(VarChangeDate, "yyyy/mm/dd")  ' 2006/01/05 ADD T.SUZUKI
'    Me.txt_Tkdate = Check_Date(Me.txt_Tkdate)           ' 2006/01/05 DEL T.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tkdate_KeyPress
'        機能             :取置き期限日 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tkdate_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Date(Me.txt_Tkdate, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tkdate_GotFocus
'        機能             :取置き期限日 GotFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tkdate_GotFocus()
    Call Date_TextFormat(Me.txt_Tkdate, 0)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tkdate_LostFocus
'        機能             :取置き期限日 LostFocusｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tkdate_LostFocus()
    If Nz(Me.txt_Tkdate, "") <> "" Then
        Call Date_TextFormat(Me.txt_Tkdate, 1)
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tkdate_Changeイベント
'        MODULE_ID        :txt_Tkdate_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tkdate_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_AfterUpdateイベント
'        MODULE_ID        :cmb_MoveKbn_AfterUpdate
'        CREATE_DATE      :2006/01/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_AfterUpdate()

    Dim strMOVEKBN As String

    On Error GoTo cmb_MoveKbn_AfterUpdate_Err

    strMOVEKBN = ""
    strMOVEKBN = Me.txt_MoveKbn.OldValue

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_MoveKbn, Me.txt_MoveKbn, P_NAME_098) = False Then GoTo cmb_MoveKbn_AfterUpdate_End

    ' タイトル設定
    If Me.txt_MoveKbn.VALUE <> strMOVEKBN Then
        Call subSetFVS500(Format(Nz(Me.cmb_MoveKbn, 0), "00"))
    End If

    GoTo cmb_MoveKbn_AfterUpdate_End

cmb_MoveKbn_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_MoveKbn_AfterUpdate_End:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_BeforeUpdateイベント
'        MODULE_ID        :cmb_MoveKbn_BeforeUpdate
'        CREATE_DATE      :2006/01/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_BeforeUpdate(Cancel As Integer)

    Dim varCode As Variant

    If bln_FormLoad = False Then Exit Sub
    
    ' 2006/06/14 Add tajima Start
    If pIs受付作成済 = True Then
        MsgBox "既に受付作成済みのため移動理由は変更できません。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        GoTo cmb_MoveKbn_BeforeUpdate_Exit
    End If
    ' 2006/06/14 Add tajima End

    'ｺｰﾄﾞが空の場合、移動理由名・移動元契約番号欄も空欄表示
    If ControlCrear(Me.cmb_MoveKbn, Me.txt_MoveKbn) = False Then
        Me.txt_MOTO_ACPTNO = ""
        Exit Sub
    End If

    varCode = Me.cmb_MoveKbn

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_MoveKbn_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_MoveKbn_BeforeUpdate_Exit
    End Select

    If Chk_NameBefore(Me.cmb_MoveKbn, P_NAME_098) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_MoveKbn_BeforeUpdate_Exit:
    Exit Sub
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_Changeイベント
'        MODULE_ID        :cmb_MoveKbn_Change
'        CREATE_DATE      :2006/01/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_MoveKbn_KeyPressイベント
'        MODULE_ID        :cmb_MoveKbn_KeyPress
'        CREATE_DATE      :2006/01/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_MoveKbn_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_MoveKbn, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :opt_Shinki_MouseDown
'        機能             :新規顧客 MouseDownｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub opt_Shinki_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Call KokClear
''    Call SetOptConf(True)　★キーボード対応 2006/06/19 Del tajima
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :KokClear
'        機能             :顧客情報クリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub KokClear()

    With Me
        .txt_Kokya = ""     ' 顧客名ｶﾅ      ' 2005/10/06 ADD T.SUZUKI
'        .cmb_Kokya = ""                    ' 2005/10/06 DEL T.SUZUKI
        .cmb_KokCd = ""
        .txt_Name = ""      '顧客名
        .txt_Kana = ""      '顧客名カナ
        .txt_Yubin = ""     '郵便番号
        .txt_Adr1 = ""      '住所１
        .txt_Adr2 = ""      '住所２
        .txt_Adr3 = ""      '住所３         ' 2005/12/12 ADD T.SUZUKI
        .txt_Tanm = ""      '代表社名
        .txt_Taka = ""      '代表者カナ
        .txt_Tel = ""       '電話番号
        .txt_Keitai = ""    '携帯番号
        .txt_Fax = ""       'FAX番号
        .txt_Mail = ""      'ﾒｰﾙｱﾄﾞﾚｽ       ' 2005/12/12 ADD T.SUZUKI
        '顧客区分
        .cmb_Kkbn.VALUE = P_KKBN_1
        '顧客区分名
        .txt_Kkbn = GetNameData(P_NAME_020, Nz(.cmb_Kkbn, ""))
    End With

' 2005/10/06 ADD T.SUZUKI Start
    doCmd.Hourglass True
    doCmd.Echo False
    If fncDeleteWorkTable = False Then Exit Sub  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ削除
    If fncCreateWorkTable = False Then Exit Sub  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ作成
    Call subCreateWork                           ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙﾃﾞｰﾀ作成
    Me.FVS500_FVS500S01.Requery                  ' 2005/10/25 ADD T.SUZUKI
'    doCmd.Requery "FVS500_FVS500S01"  ' ｻﾌﾞﾌｫｰﾑ再表示  ' 2005/10/25 DEL T.SUZUKI
    doCmd.Echo True
    doCmd.Hourglass False
' 2005/10/06 ADD T.SUZUKI End
End Sub

'==============================================================================*
'
'        MODULE_NAME      :opt_Kizon_MouseDown
'        機能             :既存顧客　MouseDownｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub opt_Kizon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Call SetOptConf(False)★ 'キーボード対応 2006/06/19 Del tajima
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Name_BeforeUpdate
'        機能             :顧客名　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Name_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Name, 36) = False Then
        Cancel = True
'        SendKeys "{ESC}"                           ' 2006/02/18 DEL T.SUZUKI
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Name_Changeイベント
'        MODULE_ID        :txt_Name_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Name_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kana_BeforeUpdate
'        機能             :顧客名ｶﾅ　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kana_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Kana, 18) = False Then
        Cancel = True
        ''SendKeys "{ESC}"    2010/02/15 Ver2.4 M.HONDA
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Kana_Changeイベント
'        MODULE_ID        :txt_Kana_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Kana_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Yubin_BeforeUpdate
'        機能             :郵便番号　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Yubin_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Yubin, 10) = False Then
        Cancel = True
        '' SendKeys "{ESC}"  2010/02/15 Ver2.4 M.HONDA
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmd_Yubin_Click
'        機能             :郵便番号　Clickｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        追加日           ：2005/06/25 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Yubin_Click()

    Dim strYUBIN_Bk As String
    Dim strYUBIN    As String

    strYUBIN_Bk = Nz(Me.txt_Yubin, "")

    Call MSZ0000_YUBIB(Me.NAME, Me.txt_Yubin.NAME, Me.txt_Adr1.NAME, _
                                Nz(Me.txt_Yubin.VALUE), Nz(Me.txt_Adr1.VALUE))

    strYUBIN = Nz(Me.txt_Yubin, "")
    If strYUBIN_Bk <> Nz(strYUBIN, "") Then
        Call txt_Yubin_Change
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Yubin_KeyPress
'        機能             :郵便番号　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Yubin_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_Yubin, KeyAscii, 10)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Yubin_Changeイベント
'        MODULE_ID        :txt_Yubin_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Yubin_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Adr1_BeforeUpdate
'        機能             :住所１　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Adr1_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Adr1, 36) = False Then
        Cancel = True
        '' SendKeys "{ESC}"  2010/02/15 Ver2.4 M.HONDA
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Adr1_Changeイベント
'        MODULE_ID        :txt_Adr1_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Adr1_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Adr2_BeforeUpdate
'        機能             :住所２　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Adr2_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Adr2, 36) = False Then
        Cancel = True
        '' SendKeys "{ESC}"  2010/02/15 Ver2.4 M.HONDA
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Adr2_Changeイベント
'        MODULE_ID        :txt_Adr2_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Adr2_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Adr3_BeforeUpdate
'        機能             :住所３　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Adr3_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Adr3, 36) = False Then
        Cancel = True
        '' SendKeys "{ESC}"  2010/02/15 Ver2.4 M.HONDA
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Adr3_Changeイベント
'        MODULE_ID        :txt_Adr3_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Adr3_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tanm_BeforeUpdate
'        機能             :代表社名　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tanm_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Tanm, 18) = False Then
        Cancel = True
'        SendKeys "{ESC}"                             ' 2006/02/18 DEL T.SUZUKI
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tanm_Changeイベント
'        MODULE_ID        :txt_Tanm_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tanm_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Taka_BeforeUpdate
'        機能             :代表者カナ　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Taka_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Taka, 18) = False Then
        Cancel = True
        '' SendKeys "{ESC}"  2010/02/15 Ver2.4 M.HONDA
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Taka_Changeイベント
'        MODULE_ID        :txt_Taka_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Taka_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tel_BeforeUpdate
'        機能             :電話番号　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_BeforeUpdate(Cancel As Integer)

    If ChkTextLength(Me.txt_Tel, 15) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If

' 2006/03/11 ADD T.SUZUKI Start
    If CmChkCher(Nz(Me.txt_Tel.VALUE), "A") = False Then
        MsgBox "半角英数字で入力してください。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
' 2006/03/11 ADD T.SUZUKI End
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tel_KeyPress
'        機能             :電話番号　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_Tel, KeyAscii, 15)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Tel_Changeイベント
'        MODULE_ID        :txt_Tel_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Tel_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Fax_BeforeUpdate
'        機能             :FAX番号　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Fax_BeforeUpdate(Cancel As Integer)

    If ChkTextLength(Me.txt_Fax, 15) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If

' 2006/03/11 ADD T.SUZUKI Start
    If CmChkCher(Nz(Me.txt_Fax.VALUE), "A") = False Then
        MsgBox "半角英数字で入力してください。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
' 2006/03/11 ADD T.SUZUKI End
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Fax_KeyPress
'        機能             :FAX番号　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Fax_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_Fax, KeyAscii, 15)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Fax_Changeイベント
'        MODULE_ID        :txt_Fax_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Fax_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Fax_BeforeUpdate
'        機能             :携帯番号　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Keitai_BeforeUpdate(Cancel As Integer)

    If ChkTextLength(Me.txt_Keitai, 15) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If

' 2006/03/11 ADD T.SUZUKI Start
    If CmChkCher(Nz(Me.txt_Keitai.VALUE), "A") = False Then
        MsgBox "半角英数字で入力してください。", vbExclamation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
' 2006/03/11 ADD T.SUZUKI End
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Keitai_KeyPress
'        機能             :携帯番号　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Keitai_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Tel(Me.txt_Keitai, KeyAscii, 15)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Keitai_Changeイベント
'        MODULE_ID        :txt_Keitai_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Keitai_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Mail_BeforeUpdate
'        機能             :Mail　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mail_BeforeUpdate(Cancel As Integer)
    If ChkTextLength(Me.txt_Mail, 60) = False Then
        Cancel = True
        SendKeys "{ESC}"
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_Mail_Changeイベント
'        MODULE_ID        :txt_Mail_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_Mail_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_RenFrom_BeforeUpdate
'        機能             :連絡時間From　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RenFrom_BeforeUpdate(Cancel As Integer)

    If ChkTextLength(Me.txt_RenFrom, 2) = False Then
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If

    '00～23のみ入力可
    If Nz(Me.txt_RenFrom, "") <> "" Then
        If Val(Me.txt_RenFrom) > 23 Then
            MsgBox "連絡時間は、「00」～「23」の範囲で入力して下さい。", vbExclamation, P_KAKUNIN
            Cancel = True
            SendKeys "{ESC}"
        End If
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_RenFrom_KeyPress
'        機能             :連絡時間From　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RenFrom_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.txt_RenFrom, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_RenFrom_Changeイベント
'        MODULE_ID        :txt_RenFrom_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RenFrom_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_RenTo_BeforeUpdate
'        機能             :連絡時間To　BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RenTo_BeforeUpdate(Cancel As Integer)

    If ChkTextLength(Me.txt_RenTo, 2) = False Then
        Cancel = True
        SendKeys "{ESC}"
        Exit Sub
    End If

    '00～23のみ入力可
    If Nz(Me.txt_RenTo, "") <> "" Then
        If Val(Me.txt_RenTo) > 23 Then
            MsgBox "連絡時間は、「00」～「23」の範囲で入力して下さい。", vbExclamation, P_KAKUNIN
            Cancel = True
            SendKeys "{ESC}"
        End If
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_RenTo_KeyPress
'        機能             :連絡時間To　KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RenTo_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num(Me.txt_RenTo, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :txt_RenTo_Changeイベント
'        MODULE_ID        :txt_RenTo_Change
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_RenTo_Change()
    Call subSetChangeFlag
End Sub

'■□■□■□■□■□■□■□■□■□ ↓↓↓　Private 関数群 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□
'==============================================================================*
'
'        MODULE_NAME      :Date_TextFormat
'        機能             :日付ﾃﾞｰﾀのﾌｫｰﾏｯﾄ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Date_TextFormat(ByRef ctlTxt As TextBox, ByVal lngMode As Long)

    Dim vCode As Variant

    vCode = Nz(ctlTxt, "")
    If vCode <> "" Then
        If lngMode = 0 Then
            ctlTxt = Format(vCode, "yy/mm/dd")
        Else
            ctlTxt = Format(vCode, "yyyy/mm/dd")
        End If
    End If
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :Etc_Lock_Rtn
'        機能             :コントロールのロックを設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Etc_Lock_Rtn()
    FVS500_FVS500S01.Enabled = False
    Me.cmd_Enter.Enabled = True
    Me.cmd_Kesi.Enabled = True
    Me.cmd_End.Enabled = True
End Sub
'==============================================================================*
'
'        MODULE_NAME      :SetOptConf
'        機能             :顧客情報入力制限
'        IN               :blnFlg = True(新規)/False(既存)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetOptConf(ByRef blnFLG As Boolean)

    Dim blnFlgA As Boolean
    Dim blnFlgB As Boolean

    If blnFLG = True Then
        blnFlgA = False  ' 既存
        blnFlgB = True   ' 新規
    Else
        blnFlgA = True   ' 既存
        blnFlgB = False  ' 新規
    End If

    With Me
        ' 既存顧客用 --------------------------------------------------------------
        .txt_Kokya.Enabled = blnFlgA    '顧客検索ﾌﾘｶﾞﾅ  ' 2005/10/06 DEL T.SUZUKI
'        .cmb_Kokya.Enabled = blnFlgA   '顧客検索ﾌﾘｶﾞﾅ  ' 2005/10/06 DEL T.SUZUKI
        .cmb_KokCd.Enabled = blnFlgA    '顧客検索ｺｰﾄﾞ
        ' -------------------------------------------------------------------------

        ' 新規顧客用 --------------------------------------------------------------
        .txt_Name.Enabled = blnFlgB     '顧客名
        .txt_Kana.Enabled = blnFlgB     '顧客名ﾌﾘｶﾞﾅ
        .cmb_Kkbn.Enabled = blnFlgB     '顧客区分ｺｰﾄﾞ
        ' -------------------------------------------------------------------------
    End With
End Sub
'==============================================================================*
'
'        MODULE_NAME      :SetModeConf
'        機能             :前画面より渡ってきたデータにより処理ﾓｰﾄﾞを設定
'        IN               :strYard = ﾔｰﾄﾞ番号 / strKont = ｺﾝﾃﾅ番号 / strUkeNo = 受付番号
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetModeConf(ByVal strYard As String, ByVal strKont As String, ByVal strUkeNo As String)

    Dim lngWork As Long

' 2005/12/12 ADD T.SUZUKI Start
    If Trim$(strUkeNo) <> "" Then
        ' 受付番号が未入力ではないとき = 修正
        lngWork = P_SHU
    Else
        ' 受付番号が未入力のとき
        If Trim$(strYard) = "" Then
            'ﾒﾆｭｰ画面より
            '(※ﾒﾆｭｰ画面からの遷移の場合ﾔｰﾄﾞ、ｺﾝﾃﾅ、予約受付番号は全て空白)
            lngWork = P_SHU
        Else
            '既存受付入力画面より
            lngWork = P_NEW
        End If
    End If
' 2005/12/12 ADD T.SUZUKI End
' 2005/12/12 DEL T.SUZUKI Start
'    If Trim$(strYard) = "" Then
'        'ﾒﾆｭｰ画面より
'        lngWork = P_SHU
'    Else
'        '既存受付入力画面より
'        If Trim$(strKont) = "" Then
'            lngWork = P_NEW
'        Else
'            If Trim$(strUkeNo) = "" Then
'                lngWork = P_NEW
'            Else
'                lngWork = P_SHU
'            End If
'        End If
'    End If
' 2005/12/12 DEL T.SUZUKI End

    mdl_Mode = lngWork
End Sub
'==============================================================================*
'
'        MODULE_NAME      :SetModeBtn
'        機能             :処理ﾓｰﾄﾞﾎﾞﾀﾝ設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetModeBtn(ByVal lngMode As Long)
    Select Case lngMode
        Case P_NEW
            Me.opg_Mode.VALUE = 1
            Call opt_New_MouseDown(1, 0, 0, 0)
        Case P_SHU
            Me.opg_Mode.VALUE = 2
            Call opt_Correction_MouseDown(1, 0, 0, 0)
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :BtnConf
'        機能             :選択された処理区分(左上ﾎﾞﾀﾝ)によりﾎﾞﾀﾝ選択(右上)可否を設定
'        IN               :lngMode = 1:受付新規 / 2:受付修正 / 3:削除
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub BtnConf(ByVal lngMode As Long)
    '保存/削除ボタン
    If lngMode = P_DEL Then
        Me.cmd_Enter.Enabled = False
        Me.cmd_Delete.Enabled = True
    Else
        Me.cmd_Enter.Enabled = True
        Me.cmd_Delete.Enabled = False
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :FormInit
'        機能             :Form初期処理
'        IN               :lngMode = 1:受付新規 / 2:受付修正 / 3:削除
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub FormInit(ByVal lngMode As Long)

    Dim strwork As String
    Dim strMsg  As String
    Dim intDays As Integer                                                      'INSERT 2024/08/27 N.IMAI

    On Error Resume Next
    
    pIs受付作成済 = False '2006/06/14 Add Tajima

    With Me
        'ﾎﾞﾀﾝの設定
        Call BtnConf(lngMode)

' 2005/10/06 ADD T.SUZUKI Start
        If fncDeleteWorkTable = False Then Exit Sub  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ削除
        If fncCreateWorkTable = False Then Exit Sub  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ作成
' 2005/10/06 ADD T.SUZUKI End

        .cmb_Ukno.Visible = (lngMode <> P_NEW)
        .txt_Ukno.Visible = (lngMode = P_NEW)
        .txt_Ukno = ""
        strBeforYard = ""                                                       'INSERT START   20050717 K.ISHIZAKA
        '処理ﾓｰﾄﾞにより初期化
        Select Case lngMode
            Case P_NEW, P_SHU, P_DEL  ' *** 新規 OR 修正 OR 削除 ***
                .opg_Mode = lngMode             'ﾓｰﾄﾞ設定

                .cmb_Ukno = ""                  '受付番号
                .cmb_Ukno.Enabled = False
                .txt_Ukdate = Format(DATE, "yyyy/mm/dd")
                .cmb_Tantoc.VALUE = ""          '受付担当コード
                .txt_Tanton = ""                '受付担当名
                .cmb_Ukkbn.VALUE = P_UKKBN_99   '受付区分コード
                '受付区分名
                .txt_Ukkbn = GetNameData(P_NAME_070, Nz(.cmb_Ukkbn, ""))

                .cmb_Yard = strForm_Yard        'ﾔｰﾄﾞｺｰﾄﾞ(前画面より送られてきたDataのSet)
                .txt_Yard = ""                  'ﾔｰﾄﾞ名
                .txt_Yadd = ""                  'ﾔｰﾄﾞ住所
                '.chk_Kinrin.VALUE = -1          '近隣ﾔｰﾄﾞ(ﾁｪｯｸ状態に設定)  ' 2007/04/06 chg tajima
                .chk_Kinrin.VALUE = 0          '近隣ﾔｰﾄﾞ(ﾁｪｯｸ状態に設定)  2014/08/02 M.HODNA UPD
                .cmb_No.VALUE = strForm_Kont    'ｺﾝﾃﾅ番号(前画面より送られてきたDataのSet)
                .txt_営業期間 = ""              '2006/09/26 add tajima
                
                '▼ add 2007/06/02 tajima ▼
                If str空検索時レンタル用途 = "" Then
                  .cmb_Youto.VALUE = ""           '用途コード ※add前のまま
                  .txt_Youto = ""                 '用途名     ※add前のまま
                Else
                  .cmb_Youto.VALUE = str空検索時レンタル用途
                  .txt_Youto = GetNameData(P_NAME_086, Nz(cmb_Youto.VALUE))
                End If
                '▲ add 2007/06/02 tajima ▲
                
                .cmb_Kibou.VALUE = ""           '希望サイズコード
                .txt_SizeNm = ""                '希望ｻｲｽﾞ名
                .txt_SyzeF = ""                 '希望サイズ実値From
                .txt_SyzeT = ""                 '希望サイズ実値To
                .txt_Kibousu = "1"              '希望本数
                .cmb_Step = ""                  '段区分
                .txt_STEP = ""                  '段名
                .txt_PricFrom = ""              '使用料From
                .txt_PricTo = ""                '使用料To

                '状態区分/状態区分名
                If Nz(cmb_No, "") = "" Then
                    .cmb_Yukbn.VALUE = P_JOU_MI
                Else
                    .cmb_Yukbn.VALUE = P_JOU_KAKU
                End If
                .txt_Yukbn = GetNameData(P_NAME_080, Nz(.cmb_Yukbn, ""))

                '状態区分により入力可否設定
                Call Yukbn_EnableConf(lngMode, .cmb_Yukbn)

                .txt_Kudate = ""                '契約受付日
                .txt_Ytdate = ""                '予約取消日
                .txt_Kkdate = ""                '希望開始日
                .txt_Tadate = ""                'お客様問合せ日
                .txt_Tkdate = ""                '取り置き期限日

                .cmb_MoveKbn = P_MOVEKBN_00     '移動区分
                .txt_MoveKbn = GetNameData(P_NAME_098, Nz(.cmb_MoveKbn, ""))
                .txt_MOTO_ACPTNO = ""           '移動元契約番号

                .cmb_Shunou = ""                '収納物選択     2005/07/24 K.MITSUNO
                .txt_Butu = ""                  '予定収納物
                .txt_Biko = ""                  '備考
              
                '▼ 2007/08/09 add tajima ▼
                .cmb_申込書送付方 = ""
                .txt_申込書送付方法 = ""
                .cmb_予定期間 = "00"
                Call cmb_予定期間_AfterUpdate
                '▲ 2007/08/09 add tajima ▲

                '▼ 2011/07/19 add tajima ▼
                .cmb_発生区分 = "00"
                Call cmb_発生区分_AfterUpdate
                '▲ 2011/07/19 add tajima ▲

                .opg_Kokya.VALUE = 1            '顧客選択
                .opt_Shinki.Enabled = True
                .opt_Kizon.Enabled = True

                .txt_Kokya.VALUE = ""           '顧客ﾌﾘｶﾞﾅ   ' 2005/10/06 ADD T.SUZUKI
                .txt_Kokya.Enabled = False                   ' 2005/10/06 ADD T.SUZUKI
'                .cmb_Kokya.Value = ""       '顧客ﾌﾘｶﾞﾅ      ' 2005/10/06 DEL T.SUZUKI
'                .cmb_Kokya.Enabled = False                  ' 2005/10/06 DEL T.SUZUKI
                .cmb_KokCd.VALUE = ""           '顧客ｺｰﾄﾞ
                .cmb_KokCd.Enabled = False

                .txt_Name = ""                  '顧客名
                .txt_Name.Enabled = True
                .txt_Kana = ""                  '顧客名フリガナ
                .txt_Kana.Enabled = True
                .txt_Tanm = ""                  '代表名
                .txt_Taka = ""                  '代表者フリガナ
                .txt_Yubin = ""                 '郵便番号
                .txt_Adr1 = ""                  '住所1
                .txt_Adr2 = ""                  '住所2
                .txt_Adr3 = ""                  '住所3
                .txt_Tel = ""                   '電話番号
                .txt_Keitai = ""                '携帯番号
                .txt_Fax = ""                   'FAX番号
                .txt_Mail = ""                  'メール

                '顧客区分
                .cmb_Kkbn.VALUE = P_KKBN_1
                '顧客区分名
                .txt_Kkbn = GetNameData(P_NAME_020, Nz(.cmb_Kkbn, ""))
                '連絡区分
                .cmb_Renkbn.VALUE = P_RENKBN_1
                '連絡区分名
                .txt_Renkbn = GetNameData(P_NAME_085, Nz(.cmb_Renkbn, ""))
                .txt_RenFrom = ""       '連絡時間From
                .txt_RenTo = ""         '連絡時間To

                '住所／ヤード名等取得
                If strForm_Yard <> "" Then Call GetYardData(1)

                Me.FVS500_FVS500S01.Requery        ' 2005/10/25 ADD T.SUZUKI
'                doCmd.Requery "FVS500_FVS500S01"  ' 2005/10/25 DEL T.SUZUKI

                Me.lbl_AutoKbn.Caption = Space(1)  ' 2005/10/06 ADD T.SUZUKI

                pintAUTOKBN = 0                    ' 2005/12/12 ADD T.SUZUKI
                pstrYUKBN = ""                     ' 2005/12/12 ADD T.SUZUKI
                pstrKONT_NO = ""                   ' 2005/12/12 ADD T.SUZUKI
                '▼ 2006/09/26 add tajima ▼
                Call SubReloadfromFVS90n
                '▲ 2006/09/26 add tajima ▲

            Case Else
                If strForm_UkeNo <> "" Then
                    '既存予約画面から渡された受付番号をｾｯﾄ
                    .cmb_Ukno = strForm_UkeNo
                Else
                    'Menuより来た時
                    .cmb_Ukno = ""
                End If

                .cmb_Ukno.Enabled = True
        End Select

' 2005/10/06 ADD T.SUZUKI Start
' '紹介件数 ■□３次ﾌｪｰｽﾞ対象 ■□
        .txt_Syoukai = ""   '紹介件数 ■□１次ﾌｪｰｽﾞ対象外 ■□
' 2005/10/06 ADD T.SUZUKI End

'        Call subDeleteWork  ' 2005/10/06 ADD T.SUZUKI  移動↑↑↑

        pblnChangeFlag_FVS500 = False  ' 2005/10/06 ADD T.SUZUKI
        If isネット予約 = True Then chk_Kinrin.Enabled = False '2007/08/09 add tajima
    End With
    
    'INSERT 2024/08/27 N.IMAI Start
    intDays = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'FVS500' AND INTIF_RECFB = 'DAYS'"), 0)
    If intDays > 0 Then
        Call subComboUKNOInit(intDays)
    End If
    'INSERT 2024/08/27 N.IMAI End

    ' 予約受付状態区分ｺﾝﾎﾞﾎﾞｯｸｽのSQL作成
    Call subCreate_YukbnSQL(lngMode, Nz(Me.cmb_Yukbn, ""))  ' 2005/12/12 ADD R.SUZUKI
End Sub
'==============================================================================*
'
'        MODULE_NAME      :SetFrmEnabledConf
'        機能             :ﾌｫｰﾑｺﾝﾄﾛｰﾙ Enabled一括設定
'        IN               :blnFlg = Enabled(True/False)
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetFrmEnabledConf(ByVal blnFLG As Boolean)

    Dim blnLock As Boolean

    On Error Resume Next  ' 2006/02/18 ADD T.SUZUKI

    txt_Ukdate.Enabled = blnFLG
    txt_Ukdate.Locked = blnFLG
    cmb_Tantoc.Enabled = blnFLG
    txt_Tanton.Enabled = blnFLG
    txt_Tanton.Locked = blnFLG
    cmb_Ukkbn.Enabled = blnFLG
    txt_Ukkbn.Enabled = blnFLG
    txt_Ukkbn.Locked = blnFLG

' 2005/10/06 ADD T.SUZUKI Start
    txt_Syoukai.Enabled = blnFLG  ' 紹介件数
    txt_Syoukai.Locked = blnFLG
' 2005/10/06 ADD T.SUZUKI End

    cmb_Yard.Enabled = blnFLG
    txt_Yard.Enabled = blnFLG
    txt_Yard.Locked = blnFLG
    txt_Yadd.Enabled = blnFLG
    txt_Yadd.Locked = blnFLG
    cmb_No.Enabled = blnFLG

    chk_Kinrin.Enabled = blnFLG   ' 近隣ﾔｰﾄﾞ

    '追加 20050625 K.MITSUNO
    cmb_Youto.Enabled = blnFLG
    txt_Youto.Enabled = blnFLG
    txt_Youto.Locked = blnFLG
    cmb_Kibou.Enabled = blnFLG
    txt_SizeNm.Enabled = blnFLG
    txt_SizeNm.Locked = blnFLG
    txt_SyzeF.Enabled = blnFLG
    'txt_SyzeF.Locked = blnFLG       '2018/08/30 TAKENOUCHI DEL
    txt_SyzeT.Enabled = blnFLG
    'txt_SyzeT.Locked = blnFLG       '2018/08/30 TAKENOUCHI DEL
    'ｺﾒﾝﾄ 20050625 K.MITSUNO
    'cmb_Syufrom.Enabled = blnFlg
    'txt_Syufrom.Enabled = blnFlg
    'txt_Syufrom.Locked = blnFlg
    'cmb_Syuto.Enabled = blnFlg
    'txt_Syuto.Enabled = blnFlg
    'txt_Syuto.Locked = blnFlg

    txt_Kibousu.Enabled = blnFLG
    txt_Kibousu.Locked = blnLock
    cmb_Step.Enabled = blnFLG
    txt_STEP.Enabled = blnFLG
    txt_STEP.Locked = blnFLG
    txt_PricFrom.Enabled = blnFLG
    txt_PricTo.Enabled = blnFLG

    cmd_Yubin.Enabled = blnFLG  '20050625 K.MITSUNO
    cmb_Yukbn.Enabled = blnFLG
    txt_Yukbn.Enabled = blnFLG
    txt_Yukbn.Locked = blnFLG
    txt_Kudate.Enabled = blnFLG
    txt_Ytdate.Enabled = blnFLG
    txt_Kkdate.Enabled = blnFLG
    txt_Tadate.Enabled = blnFLG
    txt_Tkdate.Enabled = blnFLG
    chk_Tabusyo.Enabled = blnFLG    'INSERT 2019/03/06 Y.WADA
' 2006/01/18 ADD T.SUZUKI Start
    cmb_MoveKbn.Enabled = blnFLG
    txt_MoveKbn.Enabled = blnFLG
    txt_MoveKbn.Locked = blnFLG
    txt_MOTO_ACPTNO.Enabled = blnFLG
    txt_MOTO_ACPTNO.Locked = blnFLG
' 2006/01/18 ADD T.SUZUKI End
    cmb_Shunou.Enabled = blnFLG '2005/07/17 K.MITSUNO
    txt_Butu.Enabled = blnFLG
    txt_Biko.Enabled = blnFLG

    cmd_USERCHK.Enabled = blnFLG  ' 2006/02/18 ADD T.SUZUKI

    opg_Kokya.Enabled = blnFLG
    txt_Kokya.Enabled = blnFLG    ' 2005/10/06 ADD T.SUZUKI
'    cmb_Kokya.Enabled = blnFlg   ' 2005/10/06 DEL T.SUZUKI
    cmb_KokCd.Enabled = blnFLG
    txt_Name.Enabled = blnFLG
    txt_Kana.Enabled = blnFLG
    txt_Yubin.Enabled = blnFLG
    txt_Adr1.Enabled = blnFLG
    txt_Adr2.Enabled = blnFLG
    txt_Adr3.Enabled = blnFLG
    txt_Tanm.Enabled = blnFLG
    txt_Taka.Enabled = blnFLG
    txt_Tel.Enabled = blnFLG
    txt_Keitai.Enabled = blnFLG
    txt_Fax.Enabled = blnFLG
    txt_Mail.Enabled = blnFLG

    cmb_Kkbn.Enabled = blnFLG
    txt_Kkbn.Enabled = blnFLG
    txt_Kkbn.Locked = blnFLG
    cmb_Renkbn.Enabled = blnFLG
    txt_Renkbn.Enabled = blnFLG
    txt_Renkbn.Locked = blnFLG
    txt_RenFrom.Enabled = blnFLG
    txt_RenTo.Enabled = blnFLG

    cmd_FVS400.Enabled = blnFLG  ' 2006/02/07 ADD T.SUZUKI

    FVS500_FVS500S01.Enabled = blnFLG
    FVS500_FVS500S01.Locked = blnFLG
    
    '▼ 2007/08/09 add tajima ▼
    Me.cmb_申込書送付方.Enabled = blnFLG
    Me.txt_申込書送付方法.Enabled = blnFLG
    Me.txt_申込書送付方法.Locked = blnFLG
    Me.cmb_予定期間.Enabled = blnFLG
    Me.txt_予定期間.Enabled = blnFLG
    Me.txt_予定期間.Locked = blnFLG
    '▲ 2007/08/09 add tajima ▲

    '▼ 2011/07/19 add tajima ▼
    Me.cmb_発生区分.Enabled = blnFLG
    Me.txt_発生区分.Enabled = blnFLG
    Me.txt_発生区分.Locked = blnFLG
    '▲ 2011/07/19 add tajima ▲

End Sub
'==============================================================================*
'
'        MODULE_NAME      :CmLockControl
'        機能             :ｺﾝﾄﾛｰﾙの設定
'        IN               :ContNM = ｺﾝﾄﾛｰﾙ / iMode = 処理ﾓｰﾄﾞ
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub CmLockControl(ByVal ContNM As Object, ByVal iMode As Integer)

    Dim objControl As Object
    
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
'        MODULE_NAME      :Yukbn_EnableConf
'        機能             :状態区分により入力可否設定
'        IN               :lngMode = 処理ﾓｰﾄﾞ / strKbn  = 状態区分 / intKbn = 区分
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Yukbn_EnableConf(ByVal lngMode As Long, ByRef strKbn As String)

    Dim blnFLG As Boolean  ' 2005/12/12 ADD T.SUZUKI

    With Me
' 2005/10/06 ADD T.SUZUKI Start
        ' 予約一覧画面(FVS540)からの遷移の場合
        ' (新規登録モードでの登録はさせない)
        If blnFVS540FLG = True Then
            .opt_New.Enabled = False  '処理選択ボタン(新規)
        End If
' 2005/10/06 ADD T.SUZUKI End

        Select Case lngMode
            Case P_NEW, P_SHU
                Select Case strKbn
                    Case P_YUKBN_00  '*** 自動紹介 ***
                        .cmb_Yard.Enabled = True     'ﾔｰﾄﾞｺｰﾄﾞ●必須
                        .cmb_No = ""                 'ｺﾝﾃﾅ番号×入力不可(項目ｸﾘｱ)
                        .cmb_No.Enabled = False
                        .txt_Kudate = ""             '契約受付日×入力不可(項目ｸﾘｱ)
                        .txt_Kudate.Enabled = False
                        .txt_Tkdate = ""             '取置期限日×入力不可(項目ｸﾘｱ)
                        .txt_Tkdate.Enabled = False
                        .txt_Ytdate = ""             '予約取消日×入力不可(項目ｸﾘｱ)
                        .txt_Ytdate.Enabled = False

                    Case P_YUKBN_01  '*** 自社理由 ***
                        .cmb_Yard.Enabled = True     'ﾔｰﾄﾞｺｰﾄﾞ●必須
                        .cmb_No.Enabled = True       'ｺﾝﾃﾅ番号●必須
                        .txt_Kudate.Enabled = True   '契約受付日●必須
                        .txt_Tkdate.Enabled = True   '取置期限日●必須
                        .txt_Ytdate = ""             '予約取消日×入力不可(項目ｸﾘｱ)
                        .txt_Ytdate.Enabled = False
'                        .txt_Ytdate.Enabled = True   '予約取消日●必須

                    Case P_YUKBN_02  ' *** 取置 ***
                        .cmb_Yard.Enabled = True     'ﾔｰﾄﾞｺｰﾄﾞ●必須
                        .cmb_No.Enabled = True       'ｺﾝﾃﾅ番号●必須
                        .txt_Kudate.Enabled = True   '契約受付日●必須
                        .txt_Tkdate.Enabled = True   '取置期限日●必須
                        .txt_Ytdate = ""             '予約取消日×入力不可(項目ｸﾘｱ)
                        .txt_Ytdate.Enabled = False

                    Case P_YUKBN_10                 ' *** 受付 ***            ' 2005/12/12 ADD T.SUZUKI
'                    Case P_YUKBN_10, P_YUKBN_20  ' *** 契約受付 OR 契約 ***  ' 2005/12/12 DEL T.SUZUKI
                        If (lngMode = P_SHU And pstrYUKBN = P_YUKBN_10) Then
                            ' 修正モード かつ 読み込んだ状態が「契約受付」のとき
                            .cmb_Yard.Enabled = False    'ﾔｰﾄﾞｺｰﾄﾞ×入力不可(項目そのまま)
                            .cmb_No.Enabled = False      'ｺﾝﾃﾅ番号×入力不可(項目そのまま)
                            .txt_Kudate.Enabled = False  '契約受付日×入力不可(項目そのまま)
                            .cmd_USERCHK.Enabled = False '登録チェック押下不可 2006/04/06 ADD Tajima
                            .txt_Kkdate.Enabled = False  '希望開始日×入力不可(項目そのまま)  '2006/09/26 add tajima
                            .cmb_申込書送付方.Enabled = False '2007/08/09 add tajima
                            .cmb_予定期間.Enabled = False     '2007/08/09 add tajima
                            .cmb_発生区分.Enabled = False '2011/07/19 add tajima
                        Else
                            ' 上記以外
                            .cmb_Yard.Enabled = True     'ﾔｰﾄﾞｺｰﾄﾞ●必須
                            .cmb_No.Enabled = True       'ｺﾝﾃﾅ番号●必須
                            .txt_Kudate.Enabled = True   '契約受付日●必須
                            .txt_Kkdate.Enabled = True   '希望開始日●必須  '2006/09/26 add tajima
                         End If

                        .txt_Tkdate.Enabled = True       '取置期限日●必須
                        .txt_Ytdate = ""                 '予約取消日×入力不可(項目ｸﾘｱ)
                        .txt_Ytdate.Enabled = False

' 2005/12/12 ADD T.SUZUKI Start
                    Case P_YUKBN_20                  ' *** 契約 ***
                        .cmb_Yard.Enabled = False         'ﾔｰﾄﾞｺｰﾄﾞ×入力不可(項目そのまま)
                        .cmb_No.Enabled = False           'ｺﾝﾃﾅ番号×入力不可(項目そのまま)
                        .txt_Kudate.Enabled = False       '契約受付日×入力不可(項目そのまま)
                        .txt_Tkdate.Enabled = False       '取置期限日×入力不可(項目そのまま)
                        .txt_Ytdate = ""                  '予約取消日×入力不可(項目そのまま)
                        .txt_Ytdate.Enabled = False
                        .txt_Kkdate.Enabled = False       '希望開始日×入力不可(項目そのまま)  '2006/09/26 add tajima
' 2005/12/12 ADD T.SUZUKI End
' 2006/01/30 ADD T.SUZUKI Start
                        .cmb_MoveKbn.Enabled = False      '移動区分×入力不可
                        .cmb_MoveKbn.Locked = False
                        .txt_MoveKbn.Enabled = False
                        .txt_MoveKbn.Locked = False
                        .txt_MOTO_ACPTNO.Enabled = False  '移動元受注契約番号×入力不可
                        .txt_MOTO_ACPTNO.Locked = False
' 2006/01/30 ADD T.SUZUKI End
                        .cmd_USERCHK.Enabled = False '登録チェック押下不可 2006/04/06 ADD Tajima

                    Case P_YUKBN_51, P_YUKBN_52, P_YUKBN_53 ' *** 予約キャンセル OR 予約取置キャンセル OR 受付キャンセル ***
'                    Case P_YUKBN_51, P_YUKBN_52  ' *** 予約キャンセル OR 予約取置キャンセル ***
                        .cmb_Yard.Enabled = False    'ﾔｰﾄﾞｺｰﾄﾞ×入力不可(項目そのまま)
                        .cmb_No.Enabled = False      'ｺﾝﾃﾅ番号×入力不可(項目そのまま)
                        .txt_Kudate.Enabled = False  '契約受付日×入力不可(項目そのまま)
                        .txt_Tkdate.Enabled = False  '取置期限日×入力不可(項目そのまま)
                        .txt_Ytdate.Enabled = True   '予約取消日●必須
                        .cmd_USERCHK.Enabled = False '登録チェック押下不可 2006/04/06 ADD Tajima
                        .txt_Kkdate.Enabled = False  '希望開始日×入力不可(項目そのまま)  '2006/09/26 add tajima
                End Select
            Case Else
        End Select

' 2005/12/12 ADD T.SUZUKI Start
        ' 契約状態が｢予約未確定｣のときのみ入力可能とする
        If strKbn = P_YUKBN_00 Then blnFLG = True Else blnFLG = False

        .cmb_Youto.Enabled = blnFLG     ' 用途
        .cmb_Kibou.Enabled = blnFLG     ' 希望サイズ
        .txt_Kibousu.Enabled = blnFLG   ' 本数
        .cmb_Step.Enabled = blnFLG      ' 段区分
        .txt_PricFrom.Enabled = blnFLG  ' 使用料(From)
        .txt_PricTo.Enabled = blnFLG    ' 使用料(To)
' ▼2018/08/30 TAKENOUCHI INS <S>
        .txt_SyzeF.Enabled = blnFLG     ' 希望サイズ(From)
        .txt_SyzeT.Enabled = blnFLG     ' 希望サイズ(To)
        .chk_Kinrin.Enabled = blnFLG    ' チェックボックス「近隣ヤードも紹介する」
' ▲2018/08/30 TAKENOUCHI INS <E>
' 2005/12/12 ADD T.SUZUKI End
    End With
End Sub

''==============================================================================*
''
''        MODULE_NAME      :Yukbn_EnableConf
''        機能             :状態区分により入力可否設定
''        IN               :lngMode = 処理ﾓｰﾄﾞ / strKbn  = 状態区分 / intKbn = 区分
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub Yukbn_EnableConf(ByVal lngMode As Long, ByRef strKbn As String)
'
'    With Me
'' 2005/10/06 ADD T.SUZUKI Start
'        ' 予約一覧画面(FVS540)からの遷移の場合
'        ' (新規登録モードでの登録はさせない)
'        If blnFVS540FLG = True Then
'            .opt_New.Enabled = False  '処理選択ボタン(新規)
'        End If
'' 2005/10/06 ADD T.SUZUKI End
'
'        Select Case lngMode
'            Case P_NEW, P_SHU
'                Select Case strKbn
'                    Case P_YUKBN_00  '*** 予約未確定 ***
'                        .cmb_Yard.Enabled = True    'ﾔｰﾄﾞｺｰﾄﾞ●必須
'' 2005/10/06 ADD T.SUZUKI Start
'                        .cmb_No.Enabled = True      'ｺﾝﾃﾅ番号●必須
'' 2005/10/06 ADD T.SUZUKI End
'' 2005/10/06 DEL T.SUZUKI Start
''                        .cmb_No = ""                'ｺﾝﾃﾅ番号×入力不可(項目ｸﾘｱ)
''                        .cmb_No.Enabled = False
'' 2005/10/06 DEL T.SUZUKI End
'                        .txt_Kudate = ""            '契約受付日○入力不可(項目そのまま)
'                        .txt_Kudate.Enabled = False
'                        .txt_Ytdate = ""            '予約取消日×入力不可(項目ｸﾘｱ)
'                        .txt_Ytdate.Enabled = False
'                        .txt_Tkdate = ""            '取置期限日×入力不可(項目ｸﾘｱ)
'                        .txt_Tkdate.Enabled = False
'
'                    Case P_YUKBN_01  '*** 予約確定 ***
'                        .cmb_Yard.Enabled = True    'ﾔｰﾄﾞｺｰﾄﾞ●必須
'                        .cmb_No.Enabled = True      'ｺﾝﾃﾅ番号●必須
'                        .txt_Kudate = ""            '契約受付日×入力不可(項目ｸﾘｱ)
'                        .txt_Kudate.Enabled = False
'                        .txt_Ytdate = ""            '予約取消日×入力不可(項目ｸﾘｱ)
'                        .txt_Ytdate.Enabled = False
'                        .txt_Tkdate = ""            '取置期限日×入力不可(項目ｸﾘｱ)
'                        .txt_Tkdate.Enabled = False
'
'                    Case P_YUKBN_02  ' *** 予約取置 ***
'                        .cmb_Yard.Enabled = True    'ﾔｰﾄﾞｺｰﾄﾞ●必須
'                        .cmb_No.Enabled = True      'ｺﾝﾃﾅ番号●必須
'                        .txt_Kudate.Enabled = True  '契約受付日●必須
'                        .txt_Ytdate = ""            '予約取消日×入力不可(項目ｸﾘｱ)
'                        .txt_Ytdate.Enabled = False
'                        .txt_Tkdate.Enabled = True  '取置期限日●必須
'
'                    Case P_YUKBN_10, P_YUKBN_20  ' *** 契約受付 OR 契約 ***
'                        .cmb_Yard.Enabled = True    'ﾔｰﾄﾞｺｰﾄﾞ●必須
'                        .cmb_No.Enabled = True      'ｺﾝﾃﾅ番号●必須
'                        .txt_Kudate.Enabled = True  '契約受付日●必須
'                        .txt_Ytdate = ""            '予約取消日×入力不可(項目ｸﾘｱ)
'                        .txt_Ytdate.Enabled = False
'                        .txt_Tkdate.Enabled = False '取置期限日○入力不可(項目そのまま)
'
'                    Case P_YUKBN_51, P_YUKBN_52  ' *** 予約キャンセル OR 予約取置キャンセル ***
'                        .cmb_Yard.Enabled = False   'ﾔｰﾄﾞｺｰﾄﾞ○入力不可(項目そのまま)
'                        .cmb_No.Enabled = False     'ｺﾝﾃﾅ番号○入力不可(項目そのまま)
'                        .txt_Kudate.Enabled = False '契約受付日○入力不可(項目そのまま)
'                        .txt_Ytdate.Enabled = True  '予約取消日●必須
'                        .txt_Tkdate.Enabled = False '取置期限日○入力不可(項目そのまま)
'                End Select
'            Case Else
'        End Select
'    End With
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :ControlCrear
'        機能             :ｺﾝﾎﾞﾎﾞｯｸｽの内容が空の場合、名称ﾃｷｽﾄも空にする
'        IN               :ctlCmb = ｺﾝﾎﾞﾎﾞｯｸｽ / ctlTxt = ﾃｷｽﾄ
'        OUT              :ｺﾝﾎﾞﾎﾞｯｸｽが空だったかの判断
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ControlCrear(ByRef ctlCmb As ComboBox, ByRef ctlTxt As TextBox) As Boolean
    
    If Nz(ctlCmb, "") = "" Then
        ctlTxt = ""
        Exit Function
    Else
        ControlCrear = True
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :HidukeFormat
'        機能             :日付欄ﾌｫｰﾏｯﾄ
'        IN               :strData = 日付データ
'        OUT              :ﾌｫｰﾏｯﾄ後の日付ﾃﾞｰﾀ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function HidukeFormat(ByVal strData As String) As Variant
    If Trim$(strData) <> "" Then
        HidukeFormat = Format$(strData, "yyyy/mm/dd")
    Else
        HidukeFormat = Null
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :SetUKNOBtnConf
'        機能             :受付番号ボタン使用可否の設定
'        IN               :
'        OUT              :
'        CREATE           :2005/07/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetUKNOBtnConf()

    On Error Resume Next

    Select Case mdl_Mode
        Case P_NEW  ' *** 新規 ***
            Me.cmd_UKNO.Enabled = False
        Case P_SHU  ' *** 修正 ***
            Me.cmd_UKNO.Enabled = True
        Case P_DEL  ' *** 削除 ***
            Me.cmd_UKNO.Enabled = True
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :SetSyoukaiBtnEnabledConf
'        機能             :紹介履歴ボタン使用可否の設定
'        IN               :
'        OUT              :
'        CREATE           :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetSyoukaiBtnEnabledConf()

    On Error Resume Next

    Select Case mdl_Mode
        Case P_NEW  ' *** 新規 ***
            Me.cmd_Syoukai.Enabled = False

        Case P_SHU, P_DEL ' *** 修正 OR 削除 ***
            If Trim(Nz(Me.cmb_Ukno, "")) <> "" Then
                Me.cmd_Syoukai.Enabled = True
            Else
                Me.cmd_Syoukai.Enabled = False
            End If
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :各入力項目修正時、フォームの標題を『修正中』に設定
'        MODULE_ID        :subSetChangeFlag
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetChangeFlag()

    If pblnChangeFlag_FVS500 = False Then
        Me.Caption = pstrFormCaption & P_SHU_Caption
        pblnChangeFlag_FVS500 = True  ' 変更フラグ(ON)
    End If
End Sub
''==============================================================================*
''
''        MODULE_NAME      :画面の使用可否コントロール
''        MODULE_ID        :subFormLockControl1
''        CREATE_DATE      :2005/10/06
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub subFormLockControl1(blnFlg As Boolean)
'
'    With Me
'        ' フォームのプロパティ設定
'        .AllowEdits = blnFlg      ' 更新の許可
'        .AllowDeletions = blnFlg  ' 削除の許可
'        .AllowAdditions = blnFlg  ' 追加の許可
'        .DataEntry = blnFlg       ' データ入力用
'
'        ' 処理選択ボタン
'        .opg_Mode.Enabled = blnFlg
''
'        .cmd_Enter.Enabled = blnFlg  ' 保存ボタン
'        .cmd_Kesi.Enabled = blnFlg   ' 取消ボタン
'        .cmd_End.Enabled = blnFlg    ' 終了ボタン
'
'        ' 標題設定
'        If blnFlg = True Then
'            .Caption = pstrFormCaption                  ' デフォルト標題
'        Else
'            .Caption = pstrFormCaption & P_SHO_Caption  ' 予約ご紹介紹介中標題
'        End If
'    End With
'End Sub
'==============================================================================*
'
'        MODULE_NAME      :画面の使用可否コントロール
'        MODULE_ID        :subFormLockControl
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subFormLockControl(blnFLG As Boolean)

    On Error Resume Next                                                        ' 2006/02/07 ADD T.SUZUKI

    With Me
        ' フォームのプロパティ設定
        .AllowEdits = blnFLG      ' 更新の許可
        .AllowDeletions = blnFLG  ' 削除の許可
        .AllowAdditions = blnFLG  ' 追加の許可
        .DataEntry = blnFLG       ' データ入力用

        ' ダミー項目にフォーカス移動
        .txt_HEnd.SetFocus                                                      ' 2006/03/11 ADD T.SUZUKI
'        .cmd_Syoukai.SetFocus  ' 紹介履歴ボタンにフォーカス移動                ' 2006/03/11 DEL T.SUZUKI

        ' 予約番号
        .cmd_UKNO.Enabled = blnFLG                                              ' 2006/03/11 ADD T.SUZUKI

        ' 受付番号
        .cmb_Ukno.Enabled = blnFLG
        .cmb_Ukno.Locked = IIf(blnFLG = False, True, False)

        ' 受付担当者コード
        .cmb_Tantoc.Enabled = blnFLG
        .cmb_Tantoc.Locked = IIf(blnFLG = False, True, False)

        ' 受付区分
        .cmb_Ukkbn.Enabled = blnFLG
        .cmb_Ukkbn.Locked = IIf(blnFLG = False, True, False)

        ' 取置履歴ボタン
        .cmd_Syoukai.Enabled = blnFLG                                           ' 2006/03/11 ADD T.SUZUKI

' 2006/02/07 REP T.SUZUKI Start
        .cmd_FVS400.Enabled = blnFLG
'        ' ヘルパー
'        .cmd_Excel.Enabled = blnFlg
' 2006/02/07 REP T.SUZUKI End

        ' ヤードコード
        .cmb_Yard.Enabled = blnFLG
        .cmb_Yard.Locked = IIf(blnFLG = False, True, False)

        ' コンテナ番号
        .cmb_No.Enabled = blnFLG
        .cmb_No.Locked = IIf(blnFLG = False, True, False)

        ' 用途
        .cmb_Youto.Enabled = blnFLG
        .cmb_Youto.Locked = IIf(blnFLG = False, True, False)

        ' 希望サイズ
        .cmb_Kibou.Enabled = blnFLG
        .cmb_Kibou.Locked = IIf(blnFLG = False, True, False)

        ' 本数
        .txt_Kibousu.Enabled = blnFLG
        .txt_Kibousu.Locked = IIf(blnFLG = False, True, False)

        ' 段区分
        .cmb_Step.Enabled = blnFLG
        .cmb_Step.Locked = IIf(blnFLG = False, True, False)

        ' 使用料(From)
        .txt_PricFrom.Enabled = blnFLG
        .txt_PricFrom.Locked = IIf(blnFLG = False, True, False)

        '  使用料(To)
        .txt_PricTo.Enabled = blnFLG
        .txt_PricTo.Locked = IIf(blnFLG = False, True, False)

        ' 状態区分
        .cmb_Yukbn.Enabled = blnFLG
        .cmb_Yukbn.Locked = IIf(blnFLG = False, True, False)

'        ' 契約受付日
'        .txt_Kudate.Enabled = blnFlg
'        .txt_Kudate.Locked = IIf(blnFlg = False, True, False)

'        ' 予約取消日
'        .txt_Ytdate.Enabled = blnFlg
'        .txt_Ytdate.Locked = IIf(blnFlg = False, True, False)

        ' 希望開始日
'        .txt_Kkdate.Enabled = blnFlg                           '2006/09/26 制御項目なので
'        .txt_Kkdate.Locked = IIf(blnFlg = False, True, False)

        ' お客様問合日
        .txt_Tadate.Enabled = blnFLG
        .txt_Tadate.Locked = IIf(blnFLG = False, True, False)

'        ' 取置き期限日
'        .txt_Tkdate.Enabled = blnFlg
'        .txt_Tkdate.Locked = IIf(blnFlg = False, True, False)

        ' 予定収納物(コード)
        .cmb_Shunou.Enabled = blnFLG
        .cmb_Shunou.Locked = IIf(blnFLG = False, True, False)
        
        ' 予定収納物(名称)
        .txt_Butu.Enabled = blnFLG
        .txt_Butu.Locked = IIf(blnFLG = False, True, False)

        ' 備考
        .txt_Biko.Enabled = blnFLG
        .txt_Biko.Locked = IIf(blnFLG = False, True, False)

        ' 顧客区分(新規 OR 既存)
        .opg_Kokya.Enabled = blnFLG
        .opg_Kokya.Locked = IIf(blnFLG = False, True, False)

'        ' 顧客コード(コンボボックス)
'        .cmb_Kokya.Enabled = blnFlg
'        .cmb_Kokya.Locked = IIf(blnFlg = False, True, False)

'        ' 顧客名(コンボボックス)
'        .cmb_KokCd.Enabled = blnFlg
'        .cmb_KokCd.Locked = IIf(blnFlg = False, True, False)

' 2006/02/18 DEL T.SUZUKI Start
'        ' 顧客名
'        .txt_Name.Enabled = blnFlg
'        .txt_Name.Locked = IIf(blnFlg = False, True, False)
' 2006/02/18 DEL T.SUZUKI End

'        ' 顧客名カナ
'        .txt_Kana.Enabled = blnFlg
'        .txt_Kana.Locked = IIf(blnFlg = False, True, False)

        ' 郵便番号変換ボタン
        .cmd_Yubin.Enabled = blnFLG

        ' 郵便番号
        .txt_Yubin.Enabled = blnFLG
        .txt_Yubin.Locked = IIf(blnFLG = False, True, False)

        ' 住所１
        .txt_Adr1.Enabled = blnFLG
        .txt_Adr1.Locked = IIf(blnFLG = False, True, False)

        ' 住所２
        .txt_Adr2.Enabled = blnFLG
        .txt_Adr2.Locked = IIf(blnFLG = False, True, False)

        ' 住所３
        .txt_Adr3.Enabled = blnFLG
        .txt_Adr3.Locked = IIf(blnFLG = False, True, False)

'        ' 顧客区分
'        .cmb_Kkbn.Enabled = blnFlg
'        .cmb_Kkbn.Locked = IIf(blnFlg = False, True, False)

        ' 連絡区分
        .cmb_Renkbn.Enabled = blnFLG
        .cmb_Renkbn.Locked = IIf(blnFLG = False, True, False)

        ' 連絡時間(From)
        .txt_RenFrom.Enabled = blnFLG
        .txt_RenFrom.Locked = IIf(blnFLG = False, True, False)

        ' 連絡時間(To)
        .txt_RenTo.Enabled = blnFLG
        .txt_RenTo.Locked = IIf(blnFLG = False, True, False)

        ' 代表者名
        .txt_Tanm.Enabled = blnFLG
        .txt_Tanm.Locked = IIf(blnFLG = False, True, False)

        ' 代表者名カナ
        .txt_Taka.Enabled = blnFLG
        .txt_Taka.Locked = IIf(blnFLG = False, True, False)

        ' 電話番号
        .txt_Tel.Enabled = blnFLG
        .txt_Tel.Locked = IIf(blnFLG = False, True, False)

        ' 携帯番号
        .txt_Keitai.Enabled = blnFLG
        .txt_Keitai.Locked = IIf(blnFLG = False, True, False)

        ' FAX番号
        .txt_Fax.Enabled = blnFLG
        .txt_Fax.Locked = IIf(blnFLG = False, True, False)

        ' メールアドレス
        .txt_Mail.Enabled = blnFLG
        .txt_Mail.Locked = IIf(blnFLG = False, True, False)

        '▼ 2007/08/09 add tajima ▼
        .cmb_申込書送付方.Enabled = blnFLG
        .cmb_申込書送付方.Locked = IIf(blnFLG = False, True, False)
        .cmb_予定期間.Enabled = blnFLG
        .cmb_予定期間.Locked = IIf(blnFLG = False, True, False)
        '▲ 2007/08/09 add tajima ▲

        '▼ 2011/07/19 add tajima ▼
        .cmb_発生区分.Enabled = blnFLG
        .cmb_発生区分.Locked = IIf(blnFLG = False, True, False)
        '▲ 2011/07/19 add tajima ▲

        ' 処理選択ボタン
        .opg_Mode.Enabled = blnFLG

        ' 既存顧客チェックボタン  2006/03/30 ADD Tajima
        .cmd_USERCHK.Enabled = blnFLG
        
        .cmd_Enter.Enabled = blnFLG  ' 保存ボタン
        .cmd_Kesi.Enabled = blnFLG   ' 取消ボタン
        .cmd_End.Enabled = blnFLG    ' 終了ボタン

        ' 標題設定
        If blnFLG = True Then
            .Caption = pstrFormCaption                  ' デフォルト標題
        Else
            .Caption = pstrFormCaption & P_SHO_Caption  ' 予約ご紹介紹介中標題
        End If

        ' 予約受付状態区分による項目の入力可否制限設定
        Call Yukbn_EnableConf(mdl_Mode, Me.cmb_Yukbn)   ' 2005/12/12 ADD T.SUZUKI
    End With
End Sub

'■□■□■□■□■□■□■□■□■□ ↓↓↓　DB Access系 関数群 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□

'==============================================================================*
'
'        MODULE_NAME      :Chk_NameBefore
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :ctl = 対象ｺﾝﾎﾞﾎﾞｯｸｽ / lngID = 名称ｺｰﾄﾞ
'        OUT              :ｺｰﾄﾞ名称
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_NameBefore(ByVal ctlCmb As ComboBox, ByVal strID As String) As Variant

    Dim strWhere    As String
    Dim vCode       As Variant

    'ファイルへの存在チェック
    vCode = Nz(ctlCmb, "")
    If vCode <> "" Then
        strWhere = "NAME_ID = '" & strID & "'"
        strWhere = strWhere & " AND NAME_CODE =" & Format$(vCode, "000")

        Chk_NameBefore = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
    End If
End Function
'==============================================================================*
'
'        MODULE_NAME      :Chk_NameAfter
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :ctl = 対象ｺﾝﾎﾞﾎﾞｯｸｽ / ctlText = 名称ｾｯﾄ先ﾃｷｽﾄ / strID = 名称ｺｰﾄﾞ
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_NameAfter(ByRef ctlCmb As ComboBox, ByRef ctlText As TextBox, ByVal strID As String, _
                                Optional ctlText2 As TextBox = "", Optional ctlText3 As TextBox = "") As Boolean
    
    Dim vCode       As Variant
    Dim strWhere    As String
    
    Chk_NameAfter = False
    
    vCode = Nz(ctlCmb, "")
    
    If vCode = "" Then GoTo Chk_NameAfter_Err
    
    'ファイルへの存在チェック
    strWhere = "NAME_ID = '" & strID & "'"
    strWhere = strWhere & " AND NAME_CODE =" & Format$(vCode, "000")
    
    If strID = P_NAME_087 Then
        ctlText = Nz(DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", strWhere), "")
        ctlText2 = Nz(DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", strWhere), "")
        ctlText3 = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
    Else
        ctlText = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strWhere), "")
    End If
    
    ctlCmb = Format$(ctlCmb, "00")
    
    GoTo Chk_NameAfter_End
    
Chk_NameAfter_Err:
    Exit Function
    
Chk_NameAfter_End:
    Chk_NameAfter = True
End Function
'==============================================================================*
'
'        MODULE_NAME      :GetNameData
'        機能             :ﾃﾞｰﾀ存在ﾁｪｯｸ名称マスタ
'        IN               :strID = 名称ｺｰﾄﾞ / strCD = 名称ｺｰﾄﾞ
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetNameData(ByVal strID As String, ByVal strCD As String) As String
    
    Dim strSQL As String
    
    If strCD <> "" Then
        strSQL = "NAME_ID = '" & strID & "' AND NAME_CODE = " & strCD
        GetNameData = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
    End If
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
Private Function GetTantoData(ByRef vTanto As Variant, Optional ByRef vTanNm As Variant) As Boolean

    Dim strSQL  As String
    Dim rs      As Recordset

    GetTantoData = False

    strSQL = "SELECT "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC, "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTN "
    strSQL = strSQL & "FROM TANT_MAST INNER JOIN "
    strSQL = strSQL & "dbo_CONT_MAST ON "
    strSQL = strSQL & "TANT_MAST.TANTM_BUMOC = dbo_CONT_MAST.CONT_BUMOC "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "TANT_MAST.TANTM_TANTC = " & "'" & Format$(vTanto, "000") & "'"
    Set rs = objDb.OpenRecordset(strSQL, dbOpenDynaset)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Function
    Else
        If Nz(rs("TANTM_TANTN"), "") <> "" Then vTanNm = rs("TANTM_TANTN")
    End If

GetTantoData_End:
    rs.Close
    Set rs = Nothing
    GetTantoData = True
End Function
'==============================================================================*
'
'        MODULE_NAME      :GetYoukData
'        機能             :予約受付トラン問い合わせ
'        IN               :lngMode = 0(存在ﾁｪｯｸ) / 1(ﾃﾞｰﾀ参照)
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetYoukData(ByVal lngMode As Long, Optional ByVal intKbn As Integer = 0) As Boolean


    Dim strSQL  As String
    Dim rsYouk  As Recordset
    Dim vRet    As Variant
    Dim vData   As Variant
    Dim vAdd1   As Variant
    Dim vAdd2   As Variant
    Dim vAdd3   As Variant
    Dim varWk   As Variant
    Dim lngRc   As Long

    Dim intDataCount As Integer  ' 2005/10/06 ADD T.SUZUKI
    Dim strAUTONAME  As String   ' 2005/10/06 ADD T.SUZUKI

    On Error GoTo GetYoukData_Err

    GetYoukData = False

    doCmd.SetWarnings False

    doCmd.Hourglass True  ' 2005/10/06 ADD T.SUZUKI

    If lngMode = 0 Then
        '存在チェック
        If Me.cmb_Ukno <> "" Then
            strSQL = strSQL & " YOUKT_UKNO = '" & Me.cmb_Ukno & "'"
            vRet = Nz(DLookup("YOUKT_UKNO", "dbo_YOUK_TRAN", strSQL), "")

            If intKbn = 0 Then                                                               ' 2006/04/07 ADD T.SUZUKI
                If vRet = "" Then
                    MsgBox "該当する予約受付トランがありません。", vbInformation, P_KAKUNIN
                    GoTo GetYoukData_End
                End If
            End If                                                                           ' 2006/04/07 ADD T.SUZUKI
        End If
    Else
        '予約トランData取得
        If strForm_UkeNo <> "" Then

            ' コンテナDBに接続
            If fncConnectDB = False Then Exit Function                                        ' 2005/11/02 ADD T.SUZUKI
            'DELETE 2025/12/04 T.KAWABATA START
            'strSQL = "SELECT * FROM YOUK_TRAN"                                                ' 2005/10/06 REP T.SUZUKI
            'DELETE 2025/12/04 T.KAWABATA END
            'INSERT 2025/12/04 T.KAWABATA START
            strSQL = "SELECT YOUK_TRAN.*, USER_MAST.*,"
            strSQL = strSQL & " YARD_MAST.YARD_NAME,"
            strSQL = strSQL & " YARD_MAST.YARD_ADDR_1,"
            strSQL = strSQL & " YARD_MAST.YARD_ADDR_2,"
            strSQL = strSQL & " YARD_MAST.YARD_ADDR_3,"
            strSQL = strSQL & " YARD_MAST.YARD_BEGIN_DAY,"
            strSQL = strSQL & " YARD_MAST.YARD_RENTEND_DAY"
            strSQL = strSQL & " FROM YOUK_TRAN"
            'INSERT 2025/12/04 T.KAWABATA END
            strSQL = strSQL & " LEFT OUTER JOIN USER_MAST ON YOUKT_UCODE = USER_CODE"         ' 2006/04/04 ADD Tajima
            strSQL = strSQL & " LEFT OUTER JOIN YARD_MAST ON YOUKT_YCODE = YARD_CODE"         ' 2006/09/26 add tajima
            strSQL = strSQL & " WHERE YOUKT_UKNO = '" & strForm_UkeNo & "'"
            Set rsYouk = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)  ' 2005/10/06 REP T.SUZUKI

            'ﾌｫｰﾑへﾃﾞｰﾀｾｯﾄ
            If Not rsYouk.EOF Then
                With Me
                    '▼ 2007/08/09 add tajima
                    If Nz(rsYouk("YOUKT_GENKBN"), 0) = 0 Then
                        isネット予約 = False
                        Me.lbl_Youk.Caption = P_TITTLE_YOYAKU & "：電話"
                    Else
                        isネット予約 = True
                        Me.lbl_Youk.Caption = P_TITTLE_YOYAKU & "：ネット"
                    End If
                    '▲ 2007/08/09 add tajima
                    '▼ 2011/07/19 add tajima ▼
                    .cmb_発生区分 = Format$(rsYouk("YOUKT_GENKBN"), "00")
                    Call cmb_発生区分_AfterUpdate
                    '▲ 2011/07/19 add tajima ▲
                
                    .cmb_Ukno = rsYouk("YOUKT_UKNO")
                    '受付日
                    .txt_Ukdate = Format$(rsYouk("YOUKT_UKDATE"), "yyyy/mm/dd")

                    '受付担当ｺｰﾄﾞ
                    .cmb_Tantoc = Format$(rsYouk("YOUKT_UKTANTO"), "000")

                    '受付担当名
                    lngRc = GetTantoData(rsYouk("YOUKT_UKTANTO"), vData)
                    .txt_Tanton = vData
                    '受付区分
                    .cmb_Ukkbn = rsYouk("YOUKT_UKKBN")
                    '受付区分名
'                    .txt_Ukkbn = GetNameData(P_NAME_070, Nz(rsYouk("YOUKT_UKKBN"), ""))    'DELETE 2019/03/06 Y.WADA
                    Call cmb_Ukkbn_AfterUpdate  'INSERT 2019/03/06 Y.WADA

                    'ﾔｰﾄﾞｺｰﾄﾞ
                    .cmb_Yard = Format$(rsYouk("YOUKT_YCODE"), "000000")
                    'strForm_Yard = Format$(rsYouk("YOUKT_YCODE"), "000000")
                    strBeforYard = Format$(rsYouk("YOUKT_YCODE"), "000000")     'INSERT 20050718 K.ISHIZAKA
                    
                    
                    '▼ 2006/09/26 del tajima ▼
                    ' 旧処理DLookupなんぞやめてRecordsetで取得した結果で設定する
                    'ﾔｰﾄﾞ名
'                    strSql = ""
'                    strSql = "YARD_CODE = " & Format(rsYouk("YOUKT_YCODE"), "000000")
'                    .txt_Yard = Nz(DLookup("YARD_NAME", "dbo_YARD_MAST", strSql), "")
'                    'ﾔｰﾄﾞ所在地
'                    vAdd1 = Nz(DLookup("YARD_ADDR_1", "dbo_YARD_MAST", strSql), "")
'                    vAdd2 = Nz(DLookup("YARD_ADDR_2", "dbo_YARD_MAST", strSql), "")
'                    vAdd3 = Nz(DLookup("YARD_ADDR_3", "dbo_YARD_MAST", strSql), "")
'                    .txt_Yadd = vAdd1 & vAdd2 & vAdd3
                    '▲ 2006/09/26 del tajima ▲
                    '▼ 2006/09/26 add tajima ▼
                    .txt_Yard = rsYouk.Fields("YARD_NAME")
                    .txt_Yadd = rsYouk.Fields("YARD_ADDR_1") & rsYouk.Fields("YARD_ADDR_2") & rsYouk.Fields("YARD_ADDR_3")
                    .txt_営業期間 = ""
                    If Nz(rsYouk.Fields("YARD_BEGIN_DAY"), "") <> "" Then
                        str営業開始日 = Format$(rsYouk.Fields("YARD_BEGIN_DAY"), "YYYY/MM/DD")
                        .txt_営業期間 = str営業開始日
                    End If
                    .txt_営業期間 = .txt_営業期間 & "～"
                    If Nz(rsYouk.Fields("YARD_RENTEND_DAY"), "") <> "" Then
                        str営業終了日 = Format$(rsYouk.Fields("YARD_RENTEND_DAY"), "YYYY/MM/DD")
                        .txt_営業期間 = .txt_営業期間 & str営業終了日
                    Else
                        str営業終了日 = pcstr最大年月日
                    End If
                    '▲ 2006/09/26 add tajima ▲

                    ' 近隣ﾔｰﾄﾞ(-1：紹介する／0：しない)
                    .chk_Kinrin.VALUE = rsYouk("YOUKT_NEARKBN")

                    'ｺﾝﾃﾅ番号
                    If rsYouk("YOUKT_YUKBN") = 5 Then                           'INSERT 2013/11/27 K.ISHIZAKA
                        strForm_Kont = Trim$(Mid$(var_OpenArgs, 7, 6))          'INSERT 2013/11/27 K.ISHIZAKA
                        pstrKONT_NO = strForm_Kont                              'INSERT 2013/11/27 K.ISHIZAKA
                    Else                                                        'INSERT 2013/11/27 K.ISHIZAKA
                        pstrKONT_NO = Format$(rsYouk("YOUKT_NO"), "000000")
                    End If                                                      'INSERT 2013/11/27 K.ISHIZAKA
                    .cmb_No = pstrKONT_NO


' 2005/12/12 ADD T.SUZUKI Start
                    ' 初期化
                    .cmb_Youto = "": .txt_Youto = ""   ' 用途
                    .cmb_Kibou = "": .txt_SizeNm = ""  ' 希望サイズ
                    .txt_SyzeF = "": .txt_SyzeT = ""   ' 希望サイズ(From～To)
' 2005/12/12 ADD T.SUZUKI End

                    '追加 2005/06/25 K.MITSUNO
                    '用途
                    If Nz(rsYouk("YOUKT_USAGE"), "") <> "" Then
                        .cmb_Youto = Format(rsYouk("YOUKT_USAGE"), "00")
                        .txt_Youto = GetNameData(P_NAME_086, Nz(rsYouk("YOUKT_USAGE")))
                    End If
                    '追加 2005/06/25 K.MITSUNO
                    '希望ｻｲｽﾞ
                    strSQL = ""
                    If Nz(rsYouk("YOUKT_SIZE_FROM"), "") <> "" Then
                        strSQL = strSQL & " NAME_ID = '" & P_NAME_087 & "' "                     ' 2006/01/20 ADD T.SUZUKI
                        strSQL = strSQL & " AND NAME_VALUE_FROM = " & rsYouk("YOUKT_SIZE_FROM")  ' 2006/01/20 ADD T.SUZUKI
'                        strSQL = "NAME_VALUE_FROM = " & rsYouk("YOUKT_SIZE_FROM")               ' 2006/01/20 DEL T.SUZUKI
                        If Nz(rsYouk("YOUKT_SIZE_TO"), "") <> "" Then
                            strSQL = strSQL & " AND NAME_VALUE_TO = " & rsYouk("YOUKT_SIZE_TO")
                        End If

'▼ 2018/08/30 TAKENOUCHI DEL<S> ▼
'                        varWk = Nz(DLookup("NAME_CODE", "dbo_NAME_MAST", strSQL), "")
'                        If varWk <> "" Then .cmb_Kibou = Format(varWk, "00") Else .cmb_Kibou = ""
'                        .txt_SizeNm = Nz(DLookup("NAME_NAME", "dbo_NAME_MAST", strSQL), "")
'▲ 2018/08/30 TAKENOUCHI DEL<E> ▲

' 2006/01/20 ADD T.SUZUKI Start
                        ' 希望ｻｲｽﾞ(From)・希望ｻｲｽﾞ(To)設定
                        .txt_SyzeF = Format(rsYouk("YOUKT_SIZE_FROM"), "0.00")
                        .txt_SyzeT = Format(rsYouk("YOUKT_SIZE_TO"), "0.00")
' 2006/01/20 ADD T.SUZUKI End
' 2006/01/20 DEL T.SUZUKI Start
'                        ' 希望ｻｲｽﾞ(From)・希望ｻｲｽﾞ(To)設定
'                        If Nz(.txt_SizeNm, "") <> "" Then
'                            .txt_SyzeF = Format(rsYouk("YOUKT_SIZE_FROM"), "0.00")  ' 2005/12/12 ADD T.SUZUKI
'                            .txt_SyzeT = Format(rsYouk("YOUKT_SIZE_TO"), "0.00")    ' 2005/12/12 ADD T.SUZUKI
''                            .txt_SyzeF = rsYouk("YOUKT_SIZE_FROM")                 ' 2005/12/12 DEL T.SUZUKI
''                            .txt_SyzeT = rsYouk("YOUKT_SIZE_TO")                   ' 2005/12/12 DEL T.SUZUKI
'                        Else
'                            .txt_SyzeF = ""
'                            .txt_SyzeT = ""
'                        End If
' 2006/01/20 DEL T.SUZUKI End
                    End If

                    '希望本数
                    .txt_Kibousu = rsYouk("YOUKT_KIBOUSU")
                    '段区分
                    .cmb_Step = rsYouk("YOUKT_STEP")
                    '段名
                    .txt_STEP = GetNameData(P_NAME_090, Nz(rsYouk("YOUKT_STEP")))
                    '使用料From
                    .txt_PricFrom = rsYouk("YOUKT_PRICFROM")
                    '使用料To
                    .txt_PricTo = rsYouk("YOUKT_PRICTO")

                    '状態区分
' 2005/12/12 ADD T.SUZUKI Start
                    pstrYUKBN = Format$(rsYouk("YOUKT_YUKBN"), "00")
                    If rsYouk("YOUKT_YUKBN") = 5 Then                           'INSERT 2013/11/27 K.ISHIZAKA
                        strForm_Kont = Trim$(Mid$(var_OpenArgs, 7, 6))          'INSERT 2013/11/27 K.ISHIZAKA
                        pstrKONT_NO = strForm_Kont                              'INSERT 2013/11/27 K.ISHIZAKA
                    Else                                                        'INSERT 2013/11/27 K.ISHIZAKA
                        pstrKONT_NO = Format$(rsYouk("YOUKT_NO"), "000000")
                    End If                                                      'INSERT 2013/11/27 K.ISHIZAKA
' 2005/12/12 ADD T.SUZUKI End
                    .cmb_Yukbn = pstrYUKBN                                          ' 2005/12/12 ADD T.SUZUKI
'                    .cmb_Yukbn = Format$(rsYouk("YOUKT_YUKBN"), "00")              ' 2005/12/12 DEL T.SUZUKI
                    Call Yukbn_EnableConf(mdl_Mode, .cmb_Yukbn)                     '変更可否の設定

                    ' 予約受付状態区分ｺﾝﾎﾞﾎﾞｯｸｽのSQL作成
                    ' 移動予約(修正モード)の場合、予約受付状態区分のコンボボックスの内容を
                    ' 変更する
                    Call subCreate_YukbnSQL(mdl_Mode, .cmb_Yukbn)                   ' 2005/12/12 ADD T.SUZUKI

                    '状態区分名
                    .txt_Yukbn = GetNameData(P_NAME_080, Nz(rsYouk("YOUKT_YUKBN")))
                    '契約受付日
                    .txt_Kudate = Format$(rsYouk("YOUKT_KUDATE"), "yyyy/mm/dd")
                    '予約取消日
                    .txt_Ytdate = Format$(rsYouk("YOUKT_YTDATE"), "yyyy/mm/dd")
                    '希望開始日
                    .txt_Kkdate = Format$(rsYouk("YOUKT_KKDATE"), "yyyy/mm/dd")
                    'お客様問合日
                    .txt_Tadate = Format$(rsYouk("YOUKT_TADATE"), "yyyy/mm/dd")
                    '取置き期限日
                    .txt_Tkdate = Format$(rsYouk("YOUKT_TKDATE"), "yyyy/mm/dd")
                    '移動理由
                    .cmb_MoveKbn = Nz(rsYouk("YOUKT_MOVEKBN"), "00")                ' 2006/01/18 ADD T.SUZUKI
                    Call cmb_MoveKbn_AfterUpdate                                    ' 2006/01/18 ADD T.SUZUKI
                    '移動元契約番号
                    .txt_MOTO_ACPTNO = Nz(rsYouk("YOUKT_MOTO_ACPTNO"), "")          ' 2006/01/18 ADD T.SUZUKI
                    ' 移動担当
                    .cmb_移動担当 = rsYouk("YOUKT_MOVE_TANTO")                      ' 2007/04/06 ADD tajima
                    Call cmb_移動担当_AfterUpdate                                   ' 2007/04/06 ADD tajima
                    '▼2007/08/09 add tajima▼
                    .cmb_申込書送付方 = Nz(rsYouk("YOUKT_SEND_CD"), "")
                    .txt_申込書送付方法 = ""
                    Call cmb_申込書送付方_AfterUpdate
                    .cmb_予定期間 = Nz(rsYouk("YOUKT_USE_PERIOD"), "00")
                    Call cmb_予定期間_AfterUpdate
                    '▲2007/08/09 add tajima▲

                    '予定収納物
                    .cmb_Shunou = ""                                                '2005/07/24 K.MITSUNO
                    .txt_Butu = rsYouk("YOUKT_BUTU")

                    '備考
                    .txt_Biko = rsYouk("YOUKT_BIKO")

                    '顧客情報取得選択
' 2005/12/20 ADD T.SUZUKI Start
                    .opt_Kizon.Enabled = True  ' 新規顧客ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(使用可)

                    If Nz(rsYouk("YOUKT_UCODE"), "") = "" Then
                        ' *** 新規顧客 ***
                        .opg_Kokya.VALUE = 1         'ｵﾌﾟｼｮﾝ値設定(新規)

                        .txt_Kokya = ""              '顧客検索ﾌﾘｶﾞﾅ
                        .cmb_KokCd = ""              '顧客検索ｺｰﾄﾞ

                        ' 新規・既存顧客入力エリア使用可否設定
                        Call SetOptConf(True)

                    Else
                        ' *** 既存顧客 ***
                        .opg_Kokya.VALUE = 2         'ｵﾌﾟｼｮﾝ値設定(既存)

                        ' 新規・既存顧客入力エリア使用可否設定
                        Call SetOptConf(False)

                        .txt_Kokya = rsYouk("YOUKT_KANA")                     ' 顧客検索ﾌﾘｶﾞﾅ
                        .cmb_KokCd = Format(rsYouk("YOUKT_UCODE"), "000000")  ' 顧客検索ｺｰﾄﾞ
                    End If

                    ' 読み込んできた予約受付トラン.予約受付状態区分により、
                    ' 新規 または 既存顧客のｵﾌﾟｼｮﾝﾎﾞﾀﾝの使用可否を設定する
                    If (pstrYUKBN = P_YUKBN_00 Or _
                        pstrYUKBN = P_YUKBN_01 Or _
                        pstrYUKBN = P_YUKBN_02) Then

                        ' 自動紹介 OR 自社理由 OR 取置
                        ' → 新規・既存ともに使用可能とする
                        .opt_Shinki.Enabled = True   ' 新規顧客ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(使用可)
                        .opt_Kizon.Enabled = True    ' 既存顧客ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(使用可)
                    Else
                        ' 受付 OR 契約 OR 予約ｷｬﾝｾﾙ OR 取置ｷｬﾝｾﾙ OR 受付ｷｬﾝｾﾙ
                        ' → 新規・既存ともに使用不可とする
                        .opt_Shinki.Enabled = False  ' 新規顧客ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(使用不可)
                        .opt_Kizon.Enabled = False   ' 既存顧客ｵﾌﾟｼｮﾝﾎﾞﾀﾝ(使用不可)

                        If .opg_Kokya.VALUE = 2 Then
                            ' 既存顧客のとき使用不可
                            .txt_Kokya.Enabled = False  '顧客検索ﾌﾘｶﾞﾅ
                            .cmb_KokCd.Enabled = False  '顧客検索ｺｰﾄﾞ
                        End If

                        .cmd_USERCHK.Enabled = False    ' 2006/02/18 ADD T.SUZUKi
                    End If
                    
                    FVS500_FVS500S01.Enabled = True
                    FVS500_FVS500S01.Locked = False

' 2005/10/06 ADD T.SUZUKI Start
                    doCmd.Echo False
                    If fncDeleteWorkTable = False Then Exit Function  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ削除
                    If fncCreateWorkTable = False Then Exit Function  ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙ作成
                    Call subCreateWork                ' ｻﾌﾞﾌｫｰﾑ表示用ﾜｰｸﾃｰﾌﾞﾙﾃﾞｰﾀ作成
                    Me.FVS500_FVS500S01.Requery       ' 2005/10/25 ADD T.SUZUKI
                    doCmd.Echo True
' 2005/10/06 ADD T.SUZUKI End

                    ' サブフォームの設定
                    Call Me.FVS500_FVS500S01.Form.subAcptnoSet(1)       ' 2006/01/28 ADD T.SUZUKI

                    ' 2006/04/04 ADD Tajima 新規顧客ならば予約トランの内容を設定
                    If Nz(rsYouk("YOUKT_UCODE"), "") = "" Then  '2006/04/04 ADD Tajima
                      .txt_Name = Nz(rsYouk("YOUKT_NAME"), "")      '顧客名
                      .txt_Kana = Nz(rsYouk("YOUKT_KANA"), "")      '顧客名ﾌﾘｶﾞﾅ
                      .txt_Tanm = Nz(rsYouk("YOUKT_TANM"), "")      '担当者名(代表者名)
                      .txt_Taka = Nz(rsYouk("YOUKT_TAKA"), "")      '担当者ﾌﾘｶﾞﾅ(代表者名ｶﾅ)
                      .txt_Yubin = Nz(rsYouk("YOUKT_YUBINO"), "")   '郵便番号
                      .txt_Adr1 = Nz(rsYouk("YOUKT_ADR_1"), "")     '住所１
                      .txt_Adr2 = Nz(rsYouk("YOUKT_ADR_2"), "")     '住所２
                      .txt_Adr3 = Nz(rsYouk("YOUKT_ADR_3"), "")     '住所３
                      .txt_Tel = Nz(rsYouk("YOUKT_TEL"), "")        '電話番号
                      .txt_Keitai = Nz(rsYouk("YOUKT_KEITAI"), "")  '携帯番号
                      .txt_Fax = Nz(rsYouk("YOUKT_FAX"), "")        'Fax番号
                      .txt_Mail = Nz(rsYouk("YOUKT_MAIL"), "")      'メールアドレス
                      .cmb_Kkbn = Format(rsYouk("YOUKT_KKBN"), "00") '顧客区分
                    ' 2006/04/04 ADD Start Tajima 既存顧客ならば顧客マスタの内容を設定
                    Else
                      .txt_Name = Nz(rsYouk("USER_NAME"), "")      '顧客名
                      .txt_Kana = Nz(rsYouk("USER_KANA"), "")      '顧客名ﾌﾘｶﾞﾅ
                      .txt_Tanm = Nz(rsYouk("USER_TANM"), "")      '担当者名(代表者名)
                      .txt_Taka = Nz(rsYouk("USER_TAKA"), "")      '担当者ﾌﾘｶﾞﾅ(代表者名ｶﾅ)
                      .txt_Yubin = Nz(rsYouk("USER_YUBINO"), "")   '郵便番号
                      .txt_Adr1 = Nz(rsYouk("USER_ADR_1"), "")     '住所１
                      .txt_Adr2 = Nz(rsYouk("USER_ADR_2"), "")     '住所２
                      .txt_Adr3 = Nz(rsYouk("USER_ADR_3"), "")     '住所３
                      .txt_Tel = Nz(rsYouk("USER_TEL"), "")        '電話番号
                      .txt_Keitai = Nz(rsYouk("USER_KEITAI"), "")  '携帯番号
                      .txt_Fax = Nz(rsYouk("USER_FAX"), "")        'Fax番号
                      .txt_Mail = Nz(rsYouk("USER_MAIL"), "")      'メールアドレス
                      .cmb_Kkbn = Format(rsYouk("USER_KKBN"), "00") '顧客区分
                    End If
                    ' 2006/04/04 ADD End Tajima 既存顧客ならば顧客マスタの内容を設定
                    
                    '顧客区分名
                    .txt_Kkbn = GetNameData(P_NAME_020, Nz(rsYouk("YOUKT_KKBN")))
                    '連絡区分
                    .cmb_Renkbn = Format(rsYouk("YOUKT_RENKBN"), "00")
                    '連絡区分名
                    .txt_Renkbn = GetNameData(P_NAME_085, Nz(rsYouk("YOUKT_RENKBN")))
                    '連絡時間From
                    .txt_RenFrom = Format(rsYouk("YOUKT_RENJFROM"), "00")
                    '連絡時間To
                    .txt_RenTo = Format(rsYouk("YOUKT_RENJTO"), "00")

' 2005/10/06 ADD T.SUZUKI Start
                    pintAUTOKBN = Nz(rsYouk("YOUKT_AUTOKBN"), 0)
                    If pintAUTOKBN <> 0 Then                      ' 2005/12/12 ADD T.SUZUKI
' 2006/01/24 ADD T.SUZUKI Start
                        ' 自動区分 <> 0 以外であれば、「自動」を固定出力
                        .lbl_AutoKbn.Caption = "※自" & vbCrLf & "動"
' 2006/01/24 ADD T.SUZUKI End
                    Else
                        .lbl_AutoKbn.Caption = Space(2)
                    End If
' 2005/10/06 ADD T.SUZUKI End
' 2006/06/14 ADD tajima Start
                    If Nz(rsYouk("YOUKT_RCPT_NO"), "") = "" Then
                        pIs受付作成済 = False
                    Else
                        pIs受付作成済 = True
                    End If
' 2006/06/14 ADD tajima End

                End With

                rsYouk.Close
                Set rsYouk = Nothing
            
            
            End If

            ' コンテナDB切断
            If fncDisConnectDB = False Then GoTo GetYoukData_End          ' 2005/11/02 ADD T.SUZUKI

' 2005/12/20 REP T.SUZUKI Start
            ' 予約ご紹介トランの件数設定処理
            If fncGetCountINTR_TRAN(Me.cmb_Ukno) = False Then GoTo GetYoukData_End
'' 2005/10/06 ADD T.SUZUKI Start
'            ' 予約ご紹介トランの件数取得
'            intDataCount = 0
'            intDataCount = fncGetCountINTR_TRAN(Me.cmb_Ukno)
'            Me.txt_Syoukai = Format(intDataCount, "#,##0")
'' 2005/10/06 ADD T.SUZUKI End
' 2005/12/20 REP T.SUZUKI End
        
            '▼ 2006/09/26 add tajima ▼
            ' コンテナと顧客の修繕情報をゲット
            Call SubReloadfromFVS90n
            '▲ 2006/09/26 add tajima ▲
        End If
    End If

    GetYoukData = True
    GoTo GetYoukData_End

GetYoukData_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

GetYoukData_End:
    If Not rsYouk Is Nothing Then rsYouk.Close: Set rsYouk = Nothing

' 2005/11/02 ADD T.SUZUKI Start
    If lngMode = 1 Then
        If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
        Call fncDisConnectKaseDB '2006/09/26 add tajima
    End If
' 2005/11/02 ADD T.SUZUKI End

    doCmd.SetWarnings True
    doCmd.Hourglass False  ' 2005/10/06 ADD T.SUZUKI
End Function
'==============================================================================*
'
'        MODULE_NAME      :SetYoyakuData
'        機能             :追加予約として、ヤードコード、顧客コードをデフォルトで
'                          設定する
'        IN               :
'        OUT              :
'        CREATE           :2005/12/27
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub SetYoyakuData(ByVal strKbn As String)

    Select Case strKbn
        Case P_YOYAKU_0
            ' *** 移動予約入力(新規)
            Me.cmb_Yard = strForm_Yard           ' ヤードコード
            Call cmb_Yard_AfterUpdate

            Me.cmb_No = strForm_Kont             ' コンテナ番号

            Me.opg_Kokya.VALUE = 2               ' 既存選択
            Me.cmb_KokCd = strForm_User          ' 顧客コード
            Call cmb_KokCd_AfterUpdate

            Me.cmb_MoveKbn = P_MOVEKBN_05        ' 移動区分(金額変更)
            Call cmb_MoveKbn_AfterUpdate
            Me.txt_MOTO_ACPTNO = strFrom_AcptNo  ' 元受注契約番号

            Call subCreate_YukbnSQL(P_NEW, P_YUKBN_10, P_YUKBN_10)
            Me.cmb_Yukbn = P_YUKBN_10            ' 予約受付状態区分
            Call cmb_Yukbn_AfterUpdate

            ' サブフォームの設定
            Call Me.FVS500_FVS500S01.Form.subAcptnoSet(1)

        Case P_YOYAKU_2
            ' *** 追加予約 ***
            Me.cmb_Yard = strForm_Yard           ' ヤードコード
            Call cmb_Yard_AfterUpdate

            Me.cmb_Yukbn = P_YUKBN_00            ' 予約受付状態区分
            Call cmb_Yukbn_AfterUpdate

            Me.opg_Kokya.VALUE = 2               ' 既存選択
            Me.cmb_KokCd = strForm_User          ' 顧客コード
            Call cmb_KokCd_AfterUpdate
    End Select
End Sub
'==============================================================================*
'
'        MODULE_NAME      :DB_Access
'        機能             :DBAccess処理
'        IN               :lngMode = 処理ﾓｰﾄﾞ
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DB_Access(ByVal lngMode As Long) As Boolean

    Dim strMsg As String

    On Error GoTo DB_Access_Err

    DB_Access = False
    doCmd.SetWarnings False

    'エラーメッセージ作成
    Select Case lngMode
        Case P_NEW: strMsg = "新規登録処理中に"
        Case P_SHU: strMsg = "修正処理中に"
        Case P_DEL: strMsg = "削除処理中に"
    End Select

    'DBｱｸｾｽ処理
    If DB_AccessMain(lngMode) = False Then GoTo DB_Access_Err
    
    Set objDb = CurrentDb                                                            ' 2006/04/07 ADD T.SUZUKI End

    ' 変更フラグセット
    pblnUpdateFlg = True  ' 2005/10/06 ADD T.SUZUKI

    If pstrFormID <> pcstrFVS540 Then  ' 2006/01/18 ADD T.SUZUKI
        Call BeforForm_Reload                                                       'INSERT 20050717 K.ISHIZAKA
    End If                             ' 2006/01/18 ADD T.SUZUKI

    DB_Access = True

    GoTo DB_Access_End

DB_Access_Err:
    Call MSZZ024_M00("DB_Access", True)
    Err.Clear
    
DB_Access_End:
    doCmd.SetWarnings True
End Function
'==============================================================================*
'
'        MODULE_NAME      :DB_AccessMain
'        機能             :データ新規登録処理
'        IN               :lngMode = 処理ﾓｰﾄﾞ
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function DB_AccessMain(ByVal lngMode As Long) As Boolean

    Dim strSQL      As String
    Dim strUkeNo    As String
    Dim strUser     As String
    Dim strwork     As String   '顧客カナ・代表者カナ対応
    Dim errLoop     As Error
    Dim rsYouk      As Recordset
    Dim rsRcpt      As Recordset                                                'INSERT 2020/02/28 N.IMAI
    Dim blnError    As Boolean  '2006/04/11 Add Tajima
    Dim MOTOACPTNO  As String   '2006/06/14 Add Tajima カードのため
    Dim moveDiv     As String   '2006/06/14 Add Tajima カードのため
    Dim strUKEKEIYAKU_NO As String  '2008/11/17 hirano DRCA_TRAN削除のため
    
    Dim rsCnta      As Recordset 'INSERT 2018-02-10 add
    Dim rsPric      As Recordset 'INSERT 2018-02-10 add
    Dim rsPric3     As Recordset 'INSERT 2018-02-10 add
    Dim blnWarn     As Boolean  'INSERT 2018-02-10 add
    Dim wkSTEP      As String   'INSERT 2018-02-10 add
    Dim wkSIZE      As String   'INSERT 2018-02-10 add
    
    blnError = False            '2006/04/11 Add Tajima

    blnWarn = False              'INSERT 2018-02-10 add

    On Error GoTo DB_AccessMain_Err

    DB_AccessMain = False

    ' コンテナDBに接続
    If fncConnectDB = False Then Exit Function                                                                ' 2005/11/02 ADD T.SUZUKI

' 2006/04/07 ADD T.SUZUKI Start
    Set pWRI_ObjWS = DBEngine.Workspaces(0)

    ' トランザクション開始
    pWRI_ObjWS.BeginTrans

' 2006/04/07 ADD T.SUZUKI End
    
    '作成UserID取得
    strUser = LsGetUserName()
    Select Case lngMode
        Case P_NEW, P_SHU  ' *** 新規 OR 修正 ***
            If lngMode = P_NEW Then
                '受付番号採番処理
                If UkeNo_Numbering(strUkeNo) = False Then GoTo DB_AccessMain_Err

                strSQL = strSQL & "SELECT * FROM YOUK_TRAN "                                                   ' 2005/10/06 REP T.SUZUKI
                Set rsYouk = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)                                  ' 2005/10/06 REP T.SUZUKI
            Else
                'ﾃﾞｰﾀの存在チェック
                If GetYoukData(0, 1) = False Then                                                              ' 2006/04/07 REP T.SUZUKI
                    MsgBox "予約受付トランが登録されておりません", vbInformation, P_KAKUNIN
                    GoTo DB_AccessMain_End
                End If

                '更新処理
                strSQL = "SELECT * FROM YOUK_TRAN "                                                            ' 2005/10/06 REP T.SUZUKI
                strSQL = strSQL & "WHERE YOUKT_UKNO = '" & Me.cmb_Ukno & "'"    '受付ナンバー
                Set rsYouk = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)                                  ' 2005/10/06 REP T.SUZUKI
            End If

            With rsYouk
                If lngMode = P_NEW Then
                    .AddNew
                    .Fields("YOUKT_UKNO") = strUkeNo                            '受付番号(var)
                    Me.txt_Ukno = strUkeNo
                    .Fields("YOUKT_CAMPC") = strCampCd                          '会社コード(var) 2018/09/25 Add WATANABE
                Else
                    MOTOACPTNO = Nz(.Fields("YOUKT_MOTO_ACPTNO"), "")           '2006/06/14 Add Tajima
                    moveDiv = Format$(Nz(.Fields("YOUKT_MOVEKBN"), 0), "00")    '2006/06/14 Add Tajima
                    .Edit
                End If
                .Fields("YOUKT_UKDATE") = Nz(Me.txt_Ukdate, Null)               '受付日(dat)
                .Fields("YOUKT_UKTANTO") = Nz(Me.cmb_Tantoc, Null)              '受付担当者コード(var)
                .Fields("YOUKT_UKKBN") = Val(Format(Me.cmb_Ukkbn, "00"))        '受付区分(num)
                .Fields("YOUKT_YCODE") = Val(Format(Me.cmb_Yard, "000000"))     'ﾔｰﾄﾞｺｰﾄﾞ(num)
                .Fields("YOUKT_NEARKBN") = 0                                    '近隣ﾔｰﾄﾞ(num)※ﾌｪｰｽﾞ1非対応項目
                'ｺﾝﾃﾅ選択不可の時のｺﾝﾎﾞﾎﾞｯｸｽの値を参照できない為
                'ｺﾝﾃﾅ番号(num/Null=Y)
                If Nz(Me.cmb_No, "") <> "" Then
                    .Fields("YOUKT_NO") = Cmb_ChkNull(Me.cmb_No)
                Else
                    .Fields("YOUKT_NO") = Null
                End If

                'INSERT 2005/06/25 K.MITSUNO
                .Fields("YOUKT_USAGE") = Cmb_ChkNull(Me.cmb_Youto)              '用途(num/Null=Y)
                .Fields("YOUKT_SIZE_FROM") = Txt_ChkNull(Me.txt_SyzeF)          '希望ｻｲｽﾞFrom(num/Null=Y)
                .Fields("YOUKT_SIZE_TO") = Txt_ChkNull(Me.txt_SyzeT)            '希望ｻｲｽﾞTo(num/Null=Y)

                'ｺﾒﾝﾄ 2005/06/25　K.MITSUNO
                .Fields("YOUKT_STEP") = Cmb_ChkNull(Me.cmb_Step)                '段区分(num/Null=Y)
                .Fields("YOUKT_PRICFROM") = Txt_ChkNull(Me.txt_PricFrom)        '使用料FROM(num/Null=Y)
                .Fields("YOUKT_PRICTO") = Txt_ChkNull(Me.txt_PricTo)            '使用料TO(num/Null=Y)
                .Fields("YOUKT_YUKBN") = Nz(Me.cmb_Yukbn, 0)                    '予約受付状態区分(num)
                '状態区分に付随する日付ﾃｷｽﾄは入力不可にする事があるので、以下の判断をする
                .Fields("YOUKT_KUDATE") = Txt_EnabledChk(Me.txt_Kudate)         '契約受付日(dat/Null=Y)
                .Fields("YOUKT_YTDATE") = Txt_EnabledChk(Me.txt_Ytdate)         '予約取消日(dat/Null=Y)
                .Fields("YOUKT_KKDATE") = Txt_EnabledChk(Me.txt_Kkdate)         '希望開始日(dat/Null=Y)
                .Fields("YOUKT_TKDATE") = Txt_EnabledChk(Me.txt_Tkdate)         '取り置き期限日(dat/Null=Y)
                .Fields("YOUKT_TADATE") = Txt_EnabledChk(Me.txt_Tadate)         '問合せ日(dat/Null=Y)
                .Fields("YOUKT_BUTU") = Nz(Me.txt_Butu, Null)                   '予定収納物(var/Null=Y)
                .Fields("YOUKT_BIKO") = Nz(Me.txt_Biko, Null)                   '備考(var/Null=Y)
                .Fields("YOUKT_KIBOUSU") = Nz(Me.txt_Kibousu, Null)             '希望本数(num/Null=Y)
                .Fields("YOUKT_UCODE") = Cmb_ChkNull(Me.cmb_KokCd)              '顧客ｺｰﾄﾞ(num/Null=Y)
                .Fields("YOUKT_NAME") = Me.txt_Name                             '顧客名(var)
                .Fields("YOUKT_KANA") = Me.txt_Kana                             '顧客名ｶﾅ(var)
                .Fields("YOUKT_TANM") = Txt_ChkNull(Me.txt_Tanm)                '代表者名(var/Null=Y)
                .Fields("YOUKT_TAKA") = Txt_ChkNull(Me.txt_Taka)                '代表者名ｶﾅ(var/Null=Y)
                .Fields("YOUKT_YUBINO") = Txt_ChkNull(Me.txt_Yubin)             '郵便番号(var/Null=Y)
                .Fields("YOUKT_ADR_1") = Txt_ChkNull(Me.txt_Adr1)               '住所1(var/Null=Y)
                .Fields("YOUKT_ADR_2") = Txt_ChkNull(Me.txt_Adr2)               '住所2(var/Null=Y)
                .Fields("YOUKT_ADR_3") = Txt_ChkNull(Me.txt_Adr3)               '住所3(var/Null=Y)
                .Fields("YOUKT_TEL") = Txt_ChkNull(Me.txt_Tel)                  '電話番号(var/Null=Y)
                .Fields("YOUKT_KEITAI") = Txt_ChkNull(Me.txt_Keitai)            '携帯番号(var/Null=Y)
                .Fields("YOUKT_FAX") = Txt_ChkNull(Me.txt_Fax)                  'ＦＡＸ(var/Null=Y)
                .Fields("YOUKT_KKBN") = Val(Me.cmb_Kkbn)                        '顧客区分(num)
                .Fields("YOUKT_RENKBN") = Val(Me.cmb_Renkbn)                    '連絡区分(num)
                .Fields("YOUKT_RENJFROM") = Txt_ChkNull(Me.txt_RenFrom)         '連絡時間FROM(var/Null=Y)
                .Fields("YOUKT_RENJTO") = Txt_ChkNull(Me.txt_RenTo)             '連絡時間TO(var/Null=Y)
                .Fields("YOUKT_MAIL") = Txt_ChkNull(Me.txt_Mail)                'メールアドレス(var/Null=Y)
                .Fields("YOUKT_NEARKBN") = Me.chk_Kinrin                        '近隣紹介チェック 2006/05/12 Add tajima
                .Fields("YOUKT_MOVE_TANTO") = Cmb_ChkNull(Me.cmb_移動担当)      '移動担当 2007/04/06 Add tajima
                .Fields("YOUKT_SEND_CD") = Cmb_ChkNull(Me.cmb_申込書送付方)     '2007/08/09 Add tajima
                .Fields("YOUKT_USE_PERIOD") = Cmb_ChkNull(Me.cmb_予定期間)      '2007/08/09 Add tajima
                .Fields("YOUKT_GENKBN") = Val(Format(Me.cmb_発生区分, "00"))    '2011/07/19 Add tajima
                
                If lngMode = P_NEW Then
                    .Fields("YOUKT_INSED") = Format$(DATE, "yyyymmdd")          '作成日付(var)
                    .Fields("YOUKT_INSEJ") = Format$(time, "hhmmss")            '作成時刻(var)
                    .Fields("YOUKT_INSPB") = P_PROG_ID                          '作成プログラムID(var)
                    .Fields("YOUKT_INSUB") = strUser                            '作成ユーザーID(var)
                    .Fields("YOUKT_GENBNR") = 0                                 '発生分類0:社内、2018/08/30 Add tajima
                Else
                    .Fields("YOUKT_UPDAD") = Format$(DATE, "yyyymmdd")          '更新日付(var/Null=Y)
                    .Fields("YOUKT_UPDAJ") = Format$(time, "hhmmss")            '更新時刻(var/Null=Y)
                    .Fields("YOUKT_UPDPB") = P_PROG_ID                          '更新プログラムID(var/Null=Y)
                    .Fields("YOUKT_UPDUB") = strUser                            '更新ユーザーID(var/Null=Y)
                End If

                Select Case lngMode
                    Case P_NEW
                        ' *** 新規 ***
                        .Fields("YOUKT_AUTOKBN") = 0                            '自動区分
                        .Fields("YOUKT_REQ_YCODE") = .Fields("YOUKT_YCODE")     '希望ヤード 2007/04/06 add tajima
                        'DEL 2011/07/19 .Fields("YOUKT_GENKBN") = 0                             '発生区分(0:電話) 2007/05/25 add tajima
                        .Fields("YOUKT_ORG_YUKBN") = .Fields("YOUKT_YUKBN")     '登録時予約状態   2007/05/25 add tajima
                    Case P_SHU
                        ' *** 修正 ***
                        If pstrYUKBN = P_YUKBN_02 Then
                            ' 読み込んできた予約受付状態区分 = 取置 のとき
                            If Me.cmb_Yukbn = P_YUKBN_02 Then
                                ' → 「取置」から「取置」に変更する場合
                                If Nz(Me.cmb_No, "") <> "" Then
                                    .Fields("YOUKT_AUTOKBN") = 0                '自動区分
                                End If
                            Else
                                ' 「取置」から別の予約受付状態区分に変更する場合
                                .Fields("YOUKT_AUTOKBN") = 0                '自動区分
                            End If

                        Else
                            ' 読み込んできた予約受付状態区分 <> 取置 のとき
                            .Fields("YOUKT_AUTOKBN") = 0                        '自動区分
                        End If
                End Select
' 2006/01/27 ADD T.SUZUKI Start
                .Fields("YOUKT_MOVEKBN") = Nz(Me.cmb_MoveKbn, 0)                 '移動区分
                .Fields("YOUKT_MOTO_ACPTNO") = Nz(Me.txt_MOTO_ACPTNO)            '移動元契約番号
' 2006/01/27 ADD T.SUZUKI End
                .UPDATE
                .Close
            End With
            
            '2026/01/28 M.HONDA
            'INSERT 2018-02-10 add start
'            If Me.cmb_Shunou.VALUE = "04" And Me.cmb_No <> "" Then
'
'                '更新処理
'                strSQL = "SELECT * FROM CNTA_MAST "
'                strSQL = strSQL & "WHERE CNTA_CODE = " & Me.cmb_Yard
'                strSQL = strSQL & " AND CNTA_NO = " & Me.cmb_No
'                Set rsCnta = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
''Debug.Print strSQL
'                '2018/03/17 M.HONDA INS
'                '用途を変えるのはコンテナのみ
'                If Nz(rsCnta.Fields("CNTA_USAGE"), "") = "0" Then
'
'                    '用途をコンテナ(バイク）に変更
'                    With rsCnta
'                        If Nz(.Fields("CNTA_USAGE"), "") <> "33" Then
'                            .Edit
'                            .Fields("CNTA_USAGE") = "33"
'                            .UPDATE
'                        End If
'                        wkSTEP = Nz(.Fields("CNTA_STEP"), "")
'                        wkSIZE = Nz(.Fields("CNTA_SIZE"), "")
'                        .Close
'                    End With
'                    '2019/05/20 M.HONDA DEL
'                    Set rsCnta = Nothing
'
'                    'コピー処理
'                    strSQL = "SELECT * FROM PRIC_TABL "
'                    strSQL = strSQL & " WHERE PRIC_YCODE = " & Me.cmb_Yard
'                    strSQL = strSQL & " AND PRIC_USAGE = " & "33"
'                    strSQL = strSQL & " AND PRIC_SIZE = " & wkSIZE
'                    strSQL = strSQL & " AND PRIC_STEP = " & wkSTEP
'                    Set rsPric = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
' 'Debug.Print strSQL
'
'                    'コンテナ（バイク）の価格がなかったらコンテナの価格をコピーする
'                    If rsPric.EOF Then
'
'                        strSQL = "SELECT * FROM PRIC_TABL"
'                        strSQL = strSQL & " WHERE PRIC_YCODE = " & Me.cmb_Yard
'                        strSQL = strSQL & " AND PRIC_USAGE = " & "0"
'                        strSQL = strSQL & " AND PRIC_SIZE = " & wkSIZE
'                        strSQL = strSQL & " AND PRIC_STEP = " & wkSTEP
'                        Set rsPric3 = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
'
''Debug.Print strSQL
'
'                        If rsPric3.EOF Then
'                            blnWarn = True
'                        Else
'                            With rsPric3
'                                recPric.PRIC_PRICE = Nz(.Fields("PRIC_PRICE"), "")
'                                recPric.PRIC_EZAPPI_CODE = Nz(.Fields("PRIC_EZAPPI_CODE"), "")
'                                recPric.PRIC_EZAPPI = Nz(.Fields("PRIC_EZAPPI"), "")
'                                recPric.PRIC_FZAPPI_CODE = Nz(.Fields("PRIC_FZAPPI_CODE"), "")
'                                recPric.PRIC_FZAPPI = Nz(.Fields("PRIC_FZAPPI"), "")
'                                recPric.PRIC_NEBIKI_DISABLE = Nz(.Fields("PRIC_NEBIKI_DISABLE"), "")
'                                '2019/02/05 chg start tajima
'                                '2018/08/30 add start
''                                recPric.PRIC_FZAPPI_CODE2 = Nz(.Fields("PRIC_FZAPPI_CODE2"), "")
''                                recPric.PRIC_FZAPPI2 = Nz(.Fields("PRIC_FZAPPI2"), "")
''                                recPric.PRIC_FZAPPI_CODE3 = Nz(.Fields("PRIC_FZAPPI_CODE3"), "")
''                                recPric.PRIC_FZAPPI3 = Nz(.Fields("PRIC_FZAPPI3"), "")
''                                recPric.PRIC_EZAPPI_CODE2 = Nz(.Fields("PRIC_EZAPPI_CODE2"), "")
''                                recPric.PRIC_EZAPPI2 = Nz(.Fields("PRIC_EZAPPI2"), "")
''                                recPric.PRIC_EZAPPI_CODE3 = Nz(.Fields("PRIC_EZAPPI_CODE3"), "")
''                                recPric.PRIC_EZAPPI3 = Nz(.Fields("PRIC_EZAPPI3"), "")
'                                '2018/08/30 add end
'                                recPric.PRIC_FZAPPI_CODE2 = .Fields("PRIC_FZAPPI_CODE2")
'                                recPric.PRIC_FZAPPI2 = .Fields("PRIC_FZAPPI2")
'                                recPric.PRIC_FZAPPI_CODE3 = .Fields("PRIC_FZAPPI_CODE3")
'                                recPric.PRIC_FZAPPI3 = .Fields("PRIC_FZAPPI3")
'                                recPric.PRIC_EZAPPI_CODE2 = .Fields("PRIC_EZAPPI_CODE2")
'                                recPric.PRIC_EZAPPI2 = .Fields("PRIC_EZAPPI2")
'                                recPric.PRIC_EZAPPI_CODE3 = .Fields("PRIC_EZAPPI_CODE3")
'                                recPric.PRIC_EZAPPI3 = .Fields("PRIC_EZAPPI3")
'                                '2019/02/05 chg end
'                            End With
'
'                            With rsPric
'                                .AddNew
'                                .Fields("PRIC_YCODE") = Me.cmb_Yard
'                                .Fields("PRIC_USAGE") = "33"
'                                .Fields("PRIC_SIZE") = wkSIZE
'                                .Fields("PRIC_STEP") = wkSTEP
'                                .Fields("PRIC_PRICE") = recPric.PRIC_PRICE
'                                .Fields("PRIC_EZAPPI_CODE") = recPric.PRIC_EZAPPI_CODE
'                                .Fields("PRIC_EZAPPI") = recPric.PRIC_EZAPPI
'                                .Fields("PRIC_FZAPPI_CODE") = recPric.PRIC_FZAPPI_CODE
'                                .Fields("PRIC_FZAPPI") = recPric.PRIC_FZAPPI
'                                '2018/08/30 add start
'                                .Fields("PRIC_EZAPPI_CODE2") = recPric.PRIC_EZAPPI_CODE2
'                                .Fields("PRIC_EZAPPI2") = recPric.PRIC_EZAPPI2
'                                .Fields("PRIC_FZAPPI_CODE2") = recPric.PRIC_FZAPPI_CODE2
'                                .Fields("PRIC_FZAPPI2") = recPric.PRIC_FZAPPI2
'                                .Fields("PRIC_EZAPPI_CODE3") = recPric.PRIC_EZAPPI_CODE3
'                                .Fields("PRIC_EZAPPI3") = recPric.PRIC_EZAPPI3
'                                .Fields("PRIC_FZAPPI_CODE3") = recPric.PRIC_FZAPPI_CODE3
'                                .Fields("PRIC_FZAPPI3") = recPric.PRIC_FZAPPI3
'                                '2018/08/30 add end
'                                .Fields("PRIC_NEBIKI_DISABLE") = recPric.PRIC_NEBIKI_DISABLE
'                                .Fields("PRIC_INSED") = Format$(DATE, "yyyymmdd")
'                                .Fields("PRIC_INSEJ") = Format$(time, "hhmmss")
'                                .Fields("PRIC_INSPB") = P_PROG_ID
'                                .Fields("PRIC_INSUB") = strUser
'                                .UPDATE
'                            End With
'                        End If
'                        '2019/05/20 M.HONDA DEL
'                        rsPric3.Close
'                        Set rsPric3 = Nothing
'
'                    Else
'                    End If
'                    '2019/05/20 M.HONDA DEL
'                    rsPric.Close
'                    Set rsPric = Nothing
'                End If
'            End If

            'INSERT 2018-02-10 add end
            
            'INSERT 2020/02/28 N.IMAI Start
            '受付キャンセルの場合はCARG_FILE、REQU_FILEを削除する
            If Nz(Me.cmb_Yukbn.VALUE, "") = P_YUKBN_53 Then
                strSQL = "SELECT * FROM RCPT_TRAN "
                strSQL = strSQL & "WHERE RCPT_NO = '" & Me.cmb_Ukno & "'"
                Set rsRcpt = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
                
                If rsRcpt.EOF = False Then
                    Call CmKomsMod.InsertRequDele(pobjKONT_DB, rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE, strUser, P_PROG_ID)
                
                    strSQL = ""
                    strSQL = strSQL & "DELETE FROM REQU_FILE "
                    strSQL = strSQL & "WHERE REQU_ACPTNO = '" & rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE & "' "
                    pobjKONT_DB.Execute strSQL
                    
                    strSQL = "INSERT INTO CARG_DELE "
                    strSQL = strSQL & "(CARD_YCODE"
                    strSQL = strSQL & ",CARD_NO"
                    strSQL = strSQL & ",CARD_UCODE"
                    strSQL = strSQL & ",CARD_AGRE"
                    strSQL = strSQL & ",CARD_FSDATE"
                    strSQL = strSQL & ",CARD_STDATE"
                    strSQL = strSQL & ",CARD_EDDATE"
                    strSQL = strSQL & ",CARD_CYDATE"
                    strSQL = strSQL & ",CARD_KYDATE"
                    strSQL = strSQL & ",CARD_CONTNA"
                    strSQL = strSQL & ",CARD_USAGE"
                    strSQL = strSQL & ",CARD_DOCU1"
                    strSQL = strSQL & ",CARD_DOCU2"
                    strSQL = strSQL & ",CARD_TUBOSU"
                    strSQL = strSQL & ",CARD_TTANKA"
                    strSQL = strSQL & ",CARD_RENTKG"
                    strSQL = strSQL & ",CARD_SYOZEI"
                    strSQL = strSQL & ",CARD_FRSTKG"
                    strSQL = strSQL & ",CARD_FSYOZEI"
                    strSQL = strSQL & ",CARD_FRST_BILL"
                    strSQL = strSQL & ",CARD_PREMTH_SUM"
                    strSQL = strSQL & ",CARD_SECUKG"
                    strSQL = strSQL & ",CARD_BIKO"
                    strSQL = strSQL & ",CARD_UPDATE"
                    strSQL = strSQL & ",CARD_HDATE"
                    strSQL = strSQL & ",CARD_MDATE"
                    strSQL = strSQL & ",CARD_ACPTNO"
                    strSQL = strSQL & ",CARD_CAMPC"
                    strSQL = strSQL & ",CARD_HOSYICD"
                    strSQL = strSQL & ",CARD_HOSYO_CD"
                    strSQL = strSQL & ",CARD_KAGIICD"
                    strSQL = strSQL & ",CARD_DAHIB"
                    strSQL = strSQL & ",CARD_HOKEB"
                    strSQL = strSQL & ",CARD_KAIJB"
                    strSQL = strSQL & ",CARD_HOKAI"
                    strSQL = strSQL & ",CARD_HOSYD"
                    strSQL = strSQL & ",CARD_HOSYB"
                    strSQL = strSQL & ",CARD_ADMEI"
                    strSQL = strSQL & ",CARD_KEITI"
                    strSQL = strSQL & ",CARD_KAGIA"
                    strSQL = strSQL & ",CARD_UKNO"
                    strSQL = strSQL & ",CARD_UKTANTO"
                    strSQL = strSQL & ",CARD_KAIUKEDATE"
                    strSQL = strSQL & ",CARD_KAITANTO"
                    strSQL = strSQL & ",CARD_KAICD"
                    strSQL = strSQL & ",CARD_KAIRIYUCD"
                    strSQL = strSQL & ",CARD_CMP_EXDATE"
                    strSQL = strSQL & ",CARD_KEY_RETDATE"
                    strSQL = strSQL & ",CARD_SPAREKEY"
                    strSQL = strSQL & ",CARD_MNT_REQDATE"
                    strSQL = strSQL & ",CARD_MNT_CMPDATE"
                    strSQL = strSQL & ",CARD_NIRCD"
                    strSQL = strSQL & ",CARD_SYOCD"
                    strSQL = strSQL & ",CARD_HOSCD"
                    strSQL = strSQL & ",CARD_MNT_TANTO"
                    strSQL = strSQL & ",CARD_MNT_MEMO"
                    strSQL = strSQL & ",CARD_LIQ_REQDATE"
                    strSQL = strSQL & ",CARD_LIQ_NOTE"
                    strSQL = strSQL & ",CARD_LIQ_FLG"
                    strSQL = strSQL & ",CARD_SPAREKEY_NOTE"
                    strSQL = strSQL & ",CARD_KEY_LENTNUM"
                    strSQL = strSQL & ",CARD_KEY_RETNUM"
                    strSQL = strSQL & ",CARD_KEY_CHGDATE"
                    strSQL = strSQL & ",CARD_KEY_CHGTANTO"
                    strSQL = strSQL & ",CARD_KEY_CHGRIYUCD"
                    strSQL = strSQL & ",CARD_KEY_CHGNOTE"
                    strSQL = strSQL & ",CARD_SEIKI"
                    strSQL = strSQL & ",CARD_INSED"
                    strSQL = strSQL & ",CARD_INSEJ"
                    strSQL = strSQL & ",CARD_INSPB"
                    strSQL = strSQL & ",CARD_INSUB"
                    strSQL = strSQL & ",CARD_UPDAD"
                    strSQL = strSQL & ",CARD_UPDAJ"
                    strSQL = strSQL & ",CARD_UPDPB"
                    strSQL = strSQL & ",CARD_UPDUB"
                    strSQL = strSQL & ",CARD_DELED"
                    strSQL = strSQL & ",CARD_DELEJ"
                    strSQL = strSQL & ",CARD_DELPB"
                    strSQL = strSQL & ",CARD_DELUB"
                    strSQL = strSQL & ") "
                    strSQL = strSQL & "SELECT "
                    strSQL = strSQL & " CARG_YCODE"
                    strSQL = strSQL & ",CARG_NO"
                    strSQL = strSQL & ",CARG_UCODE"
                    strSQL = strSQL & ",CARG_AGRE"
                    strSQL = strSQL & ",CARG_FSDATE"
                    strSQL = strSQL & ",CARG_STDATE"
                    strSQL = strSQL & ",CARG_EDDATE"
                    strSQL = strSQL & ",CARG_CYDATE"
                    strSQL = strSQL & ",CARG_KYDATE"
                    strSQL = strSQL & ",CARG_CONTNA"
                    strSQL = strSQL & ",CARG_USAGE"
                    strSQL = strSQL & ",CARG_DOCU1"
                    strSQL = strSQL & ",CARG_DOCU2"
                    strSQL = strSQL & ",CARG_TUBOSU"
                    strSQL = strSQL & ",CARG_TTANKA"
                    strSQL = strSQL & ",CARG_RENTKG"
                    strSQL = strSQL & ",CARG_SYOZEI"
                    strSQL = strSQL & ",CARG_FRSTKG"
                    strSQL = strSQL & ",CARG_FSYOZEI"
                    strSQL = strSQL & ",CARG_FRST_BILL"
                    strSQL = strSQL & ",CARG_PREMTH_SUM"
                    strSQL = strSQL & ",CARG_SECUKG"
                    strSQL = strSQL & ",CARG_BIKO"
                    strSQL = strSQL & ",CARG_UPDATE"
                    strSQL = strSQL & ",CARG_HDATE"
                    strSQL = strSQL & ",CARG_MDATE"
                    strSQL = strSQL & ",CARG_ACPTNO"
                    strSQL = strSQL & ",CARG_CAMPC"
                    strSQL = strSQL & ",CARG_HOSYICD"
                    strSQL = strSQL & ",CARG_HOSYO_CD"
                    strSQL = strSQL & ",CARG_KAGIICD"
                    strSQL = strSQL & ",CARG_DAHIB"
                    strSQL = strSQL & ",CARG_HOKEB"
                    strSQL = strSQL & ",CARG_KAIJB"
                    strSQL = strSQL & ",CARG_HOKAI"
                    strSQL = strSQL & ",CARG_HOSYD"
                    strSQL = strSQL & ",CARG_HOSYB"
                    strSQL = strSQL & ",CARG_ADMEI"
                    strSQL = strSQL & ",CARG_KEITI"
                    strSQL = strSQL & ",CARG_KAGIA"
                    strSQL = strSQL & ",CARG_UKNO"
                    strSQL = strSQL & ",CARG_UKTANTO"
                    strSQL = strSQL & ",CARG_KAIUKEDATE"
                    strSQL = strSQL & ",CARG_KAITANTO"
                    strSQL = strSQL & ",CARG_KAICD"
                    strSQL = strSQL & ",CARG_KAIRIYUCD"
                    strSQL = strSQL & ",CARG_CMP_EXDATE"
                    strSQL = strSQL & ",CARG_KEY_RETDATE"
                    strSQL = strSQL & ",CARG_SPAREKEY"
                    strSQL = strSQL & ",CARG_MNT_REQDATE"
                    strSQL = strSQL & ",CARG_MNT_CMPDATE"
                    strSQL = strSQL & ",CARG_NIRCD"
                    strSQL = strSQL & ",CARG_SYOCD"
                    strSQL = strSQL & ",CARG_HOSCD"
                    strSQL = strSQL & ",CARG_MNT_TANTO"
                    strSQL = strSQL & ",CARG_MNT_MEMO"
                    strSQL = strSQL & ",CARG_LIQ_REQDATE"
                    strSQL = strSQL & ",CARG_LIQ_NOTE"
                    strSQL = strSQL & ",CARG_LIQ_FLG"
                    strSQL = strSQL & ",CARG_SPAREKEY_NOTE"
                    strSQL = strSQL & ",CARG_KEY_LENTNUM"
                    strSQL = strSQL & ",CARG_KEY_RETNUM"
                    strSQL = strSQL & ",CARG_KEY_CHGDATE"
                    strSQL = strSQL & ",CARG_KEY_CHGTANTO"
                    strSQL = strSQL & ",CARG_KEY_CHGRIYUCD"
                    strSQL = strSQL & ",CARG_KEY_CHGNOTE"
                    strSQL = strSQL & ",CARG_SEIKI"
                    strSQL = strSQL & ",CARG_INSED"
                    strSQL = strSQL & ",CARG_INSEJ"
                    strSQL = strSQL & ",CARG_INSPB"
                    strSQL = strSQL & ",CARG_INSUB"
                    strSQL = strSQL & ",CARG_UPDAD"
                    strSQL = strSQL & ",CARG_UPDAJ"
                    strSQL = strSQL & ",CARG_UPDPB"
                    strSQL = strSQL & ",CARG_UPDUB"
                    strSQL = strSQL & ",'" & Format$(DATE, "yyyymmdd") & "'"
                    strSQL = strSQL & ",'" & Format$(time, "hhmmss") & "'"
                    strSQL = strSQL & ",'" & P_PROG_ID & "'"
                    strSQL = strSQL & ",'" & strUser & "'"
                    strSQL = strSQL & " FROM CARG_FILE "
                    strSQL = strSQL & " WHERE CARG_ACPTNO = '" & rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE & "' "
                    pobjKONT_DB.Execute strSQL
                    
                    strSQL = ""
                    strSQL = strSQL & "DELETE FROM CARG_FILE "
                    strSQL = strSQL & "WHERE CARG_ACPTNO = '" & rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE & "' "
                    pobjKONT_DB.Execute strSQL
                End If
            End If
            'INSERT 2020/02/28 N.IMAI End

        Case P_DEL  ' *** 削除 ***
            'ﾃﾞｰﾀの存在ﾁｪｯｸ
            If GetYoukData(0, 1) = False Then                                                                  ' 2006/04/07 REP T.SUZUKI
               MsgBox "予約受付トランが登録されておりません", vbInformation, P_KAKUNIN                                            ' 2006/04/07 ADD T.SUZUKI
                GoTo DB_AccessMain_End
            End If

            '更新処理
            strSQL = "SELECT * FROM YOUK_TRAN "                                                       ' 2005/10/06 REP T.SUZUKI
            strSQL = strSQL & "WHERE YOUKT_UKNO = '" & Me.cmb_Ukno & "'"    '受付№
            Set rsYouk = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)                                      ' 2005/10/06 REP T.SUZUKI
            MOTOACPTNO = Nz(rsYouk.Fields("YOUKT_MOTO_ACPTNO"), "")         '2006/06/14 Add Tajima
            moveDiv = Format$(Nz(rsYouk.Fields("YOUKT_MOVEKBN"), 0), "00")  '2006/06/14 Add Tajima
            With rsYouk
                .Edit
                .Delete
                .Close
            End With
    End Select
    Set rsYouk = Nothing                                                                                       ' 2005/10/06 ADD T.SUZUKI

' 2005/12/12 ADD T.SUZUKI Start
    Dim strYOUKT_UKNO As String

    ' ご紹介トランの作成 OR 更新処理
    ' 予約受付番号取得
    If lngMode = P_NEW Then
        strYOUKT_UKNO = strUkeNo
    Else
        strYOUKT_UKNO = IIf(Nz(Me.txt_Ukno, "") <> "", Me.txt_Ukno, Me.cmb_Ukno)
    End If
    If fncUpdINTR_TRAN(strYOUKT_UKNO) = False Then GoTo DB_AccessMain_Err                           ' 2006/04/07 ADD T.SUZUKI
' 2005/12/12 ADD T.SUZUKI End

' 2006/03/30 ADD H.Tajima Start
  ' 顧客トランへの更新
  ' 既存顧客且つ、新規・修正の場合のみ
  If Nz(Me.cmb_KokCd.VALUE, "") <> "" And (lngMode = P_NEW Or lngMode = P_SHU) Then
    If fncUPD_USER_MAST(pobjKONT_DB) = False Then                                                             ' 2006/04/07 DEL T.SUZUKI
      GoTo DB_AccessMain_Err                                                                                   ' 2006/04/07 ADD T.SUZUKI
    End If
  End If
' 2006/03/30 ADD H.Tajima End

' 2006/06/14 ADD H.Tajima Start
  ' 受付キャンセルまたは削除(コンテナ確定済み)した場合、セキュリティカードの貸出情報を初期化
  If (lngMode = P_DEL And strYOUKT_UKNO <> "") Or _
     (lngMode = P_SHU And Me.cmb_Yukbn.VALUE = P_YUKBN_53) Then
      
      ' 移動区分によって削除処理を変える
      ' ※制御する情報(移動区分、移動元契約)は画面で変更してしまう場合もありえるので
      '　 ＤＢから直接もってくるようにしている
      Select Case moveDiv
          ' 契約変更系ならばカードの契約番号を元に戻す
'          Case P_MOVEKBN_04, P_MOVEKBN_05, P_MOVEKBN_06, P_MOVEKBN_満了更新 '2008/03/23 add tajima  'DELETE 2011/09/27 M.RYU
          Case P_MOVEKBN_04, P_MOVEKBN_05, P_MOVEKBN_06, P_MOVEKBN_07, P_MOVEKBN_満了更新 'INSERT 2011/09/27 M.RYU
              ' 付け替え先の新契約番号を取得
              Dim rcptCargAcptno As String
              strSQL = "SELECT RCPT_CARG_ACPTNO FROM RCPT_TRAN "                                                       ' 2005/10/06 REP T.SUZUKI
              strSQL = strSQL & "WHERE RCPT_NO = '" & Me.cmb_Ukno & "'"    '受付№
              Set rsYouk = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)                                      ' 2005/10/06 REP T.SUZUKI
              rcptCargAcptno = ""
              If Not rsYouk.EOF Then rcptCargAcptno = Nz(rsYouk.Fields("RCPT_CARG_ACPTNO"), "")
              rsYouk.Close
              Set rsYouk = Nothing
              ' 付け替え先の新契約番号が取れればカードの付け替えをする
              If rcptCargAcptno <> "" Then
                      Call CmKomsMod.SetChangeSCards(pobjKONT_DB, rcptCargAcptno, MOTOACPTNO, strUser, P_PROG_ID)
              End If
          
          ' それ以外ならばカード取消処理
          Case Else
              Call CmKomsMod.CancelContractSCards(pobjKONT_DB, "", strUser, P_PROG_ID, strYOUKT_UKNO)
      End Select
  End If
' 2006/06/14 ADD H.Tajima End


' 2006/03/26 ADD H.Tajima Start
  ' 受付トランへの更新というか削除
  If lngMode = P_DEL Then
    '▼2008/12/17 DRCA_TRANへの引数を受付契約番号にする
    'If fncDel_RCPT_TRAN(pobjKONT_DB, strYOUKT_UKNO) = False Then GoTo DB_AccessMain_Err
    ''▼2008/11/19 hirano ADD 割引き割付削除
    'If fncDel_DRCA_TRAN(pobjKONT_DB, strYOUKT_UKNO) = False Then GoTo DB_AccessMain_Err
    If fncDel_REQU_FILE(pobjKONT_DB, lngMode, strUser) = False Then GoTo DB_AccessMain_Err                          'INSERT 2020/08/07 N.IMAI
    If fncDel_RCPT_TRAN(pobjKONT_DB, strYOUKT_UKNO, strUKEKEIYAKU_NO) = False Then GoTo DB_AccessMain_Err
    If fncDel_DRCA_TRAN(pobjKONT_DB, strUKEKEIYAKU_NO) = False Then GoTo DB_AccessMain_Err
    '▲2008/12/17
    '▲2008/11/19 hirano ADD
    ' 自分か紹介待ちにしたコンテナをクリアする
    Call CmKomsMod.ClearWaitCntaMast(pobjKONT_DB, strYOUKT_UKNO) '2006/05/02 tajima Add
'▼2008/11/19 hirano ADD 受付キャンセルの場合、割引き割付削除
  ElseIf (lngMode = P_SHU And Me.cmb_Yukbn.VALUE = P_YUKBN_53) Then
    If fncDel_DRCA_TRAN(pobjKONT_DB, strYOUKT_UKNO) = False Then GoTo DB_AccessMain_Err
'▲2008/11/19 hirano ADD
  End If
' 2006/03/26 ADD H.Tajima End

    ' コミット
    pWRI_ObjWS.CommitTrans                                                                                     ' 2006/04/07 ADD T.SUZUKI
    
    ' コンテナDB切断
    If fncDisConnectDB(1) = False Then GoTo DB_AccessMain_Err                                       ' 2006/04/07 ADD T.SUZUKI

    doCmd.SetWarnings False
    DB_AccessMain = True

    Set pWRI_ObjWS = Nothing                                                                                   ' 2006/04/07 ADD T.SUZUKI

    
    '統合ＤＢへの更新
    If Not MSTGP01_M10(cmb_Yard.VALUE, Nz(cmb_No.VALUE, "")) Then               'INSERT 2011/02/21 K.ISHIZAKA
        Call MSZZ024_M10("MSTGP01_M10", "統合ＤＢへの更新に失敗しました。")     'INSERT 2011/02/21 K.ISHIZAKA
    End If                                                                      'INSERT 2011/02/21 K.ISHIZAKA

    GoTo DB_AccessMain_End

DB_AccessMain_Err:
    '*** エラー処理（Ｒｏｌｌｂａｃｋ） ***
    pWRI_ObjWS.Rollback   ' 2006/04/07 ADD T.SUZUKI
    blnError = True       ' 2006/04/11 Add Tajima

DB_AccessMain_End:
    If Not rsYouk Is Nothing Then rsYouk.Close: Set rsYouk = Nothing
'    If Not rsCnta Is Nothing Then rsCnta.Close: Set rsCnta = Nothing          'INSERT 2018-02-10 add
'    If Not rsPric Is Nothing Then rsPric.Close: Set rsPric = Nothing          'INSERT 2018-02-10 add
'    If Not rsPric3 Is Nothing Then rsPric3.Close: Set rsPric3 = Nothing       'INSERT 2018-02-10 add
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing                            ' 2005/11/02 ADD T.SUZUKI
    If Not pWRI_ObjWS Is Nothing Then Set pWRI_ObjWS = Nothing                               ' 2006/04/07 ADD T.SUZUKI

  If blnError Then
    Call Err.Raise(Err.Number, "DB_AccessMain" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
  End If

  'INSERT 2018-02-10 add start
  If blnWarn Then
    MsgBox "予約の保存処理は成功しましたが、コンテナの価格表が存在しないため" & vbCrLf & "コンテナ（バイク）の価格表を作成できませんでした。"
  End If
  'INSERT 2018-02-10 add end

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDel_REQU_FILE 2020/08/07
'        機能             :契約トラン、請求ファイル削除
'        IN               :strUser
'        OUT              :削除結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDel_REQU_FILE(ByRef aDbObject As Database, lngMode As Long, strUser As String) As Boolean

    Dim strSQL      As String
    Dim rsRcpt      As Recordset
    
    On Error GoTo ErrorHandler

    fncDel_REQU_FILE = True

    '削除または受付キャンセルの場合はCARG_FILE、REQU_FILEを削除する
    If Nz(Me.cmb_Yukbn.VALUE, "") = P_YUKBN_53 Or lngMode = P_DEL Then
        strSQL = "SELECT * FROM RCPT_TRAN "
        strSQL = strSQL & "WHERE RCPT_NO = '" & Me.cmb_Ukno & "'"
        Set rsRcpt = aDbObject.OpenRecordset(strSQL, dbOpenDynaset)
        On Error GoTo ErrorHandler2
        
        If rsRcpt.EOF = False Then
            Call CmKomsMod.InsertRequDele(aDbObject, rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE, strUser, P_PROG_ID)
        
            strSQL = ""
            strSQL = strSQL & "DELETE FROM REQU_FILE "
            strSQL = strSQL & "WHERE REQU_ACPTNO = '" & rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE & "' "
            aDbObject.Execute strSQL
            
            strSQL = "INSERT INTO CARG_DELE "
            strSQL = strSQL & "(CARD_YCODE"
            strSQL = strSQL & ",CARD_NO"
            strSQL = strSQL & ",CARD_UCODE"
            strSQL = strSQL & ",CARD_AGRE"
            strSQL = strSQL & ",CARD_FSDATE"
            strSQL = strSQL & ",CARD_STDATE"
            strSQL = strSQL & ",CARD_EDDATE"
            strSQL = strSQL & ",CARD_CYDATE"
            strSQL = strSQL & ",CARD_KYDATE"
            strSQL = strSQL & ",CARD_CONTNA"
            strSQL = strSQL & ",CARD_USAGE"
            strSQL = strSQL & ",CARD_DOCU1"
            strSQL = strSQL & ",CARD_DOCU2"
            strSQL = strSQL & ",CARD_TUBOSU"
            strSQL = strSQL & ",CARD_TTANKA"
            strSQL = strSQL & ",CARD_RENTKG"
            strSQL = strSQL & ",CARD_SYOZEI"
            strSQL = strSQL & ",CARD_FRSTKG"
            strSQL = strSQL & ",CARD_FSYOZEI"
            strSQL = strSQL & ",CARD_FRST_BILL"
            strSQL = strSQL & ",CARD_PREMTH_SUM"
            strSQL = strSQL & ",CARD_SECUKG"
            strSQL = strSQL & ",CARD_BIKO"
            strSQL = strSQL & ",CARD_UPDATE"
            strSQL = strSQL & ",CARD_HDATE"
            strSQL = strSQL & ",CARD_MDATE"
            strSQL = strSQL & ",CARD_ACPTNO"
            strSQL = strSQL & ",CARD_CAMPC"
            strSQL = strSQL & ",CARD_HOSYICD"
            strSQL = strSQL & ",CARD_HOSYO_CD"
            strSQL = strSQL & ",CARD_KAGIICD"
            strSQL = strSQL & ",CARD_DAHIB"
            strSQL = strSQL & ",CARD_HOKEB"
            strSQL = strSQL & ",CARD_KAIJB"
            strSQL = strSQL & ",CARD_HOKAI"
            strSQL = strSQL & ",CARD_HOSYD"
            strSQL = strSQL & ",CARD_HOSYB"
            strSQL = strSQL & ",CARD_ADMEI"
            strSQL = strSQL & ",CARD_KEITI"
            strSQL = strSQL & ",CARD_KAGIA"
            strSQL = strSQL & ",CARD_UKNO"
            strSQL = strSQL & ",CARD_UKTANTO"
            strSQL = strSQL & ",CARD_KAIUKEDATE"
            strSQL = strSQL & ",CARD_KAITANTO"
            strSQL = strSQL & ",CARD_KAICD"
            strSQL = strSQL & ",CARD_KAIRIYUCD"
            strSQL = strSQL & ",CARD_CMP_EXDATE"
            strSQL = strSQL & ",CARD_KEY_RETDATE"
            strSQL = strSQL & ",CARD_SPAREKEY"
            strSQL = strSQL & ",CARD_MNT_REQDATE"
            strSQL = strSQL & ",CARD_MNT_CMPDATE"
            strSQL = strSQL & ",CARD_NIRCD"
            strSQL = strSQL & ",CARD_SYOCD"
            strSQL = strSQL & ",CARD_HOSCD"
            strSQL = strSQL & ",CARD_MNT_TANTO"
            strSQL = strSQL & ",CARD_MNT_MEMO"
            strSQL = strSQL & ",CARD_LIQ_REQDATE"
            strSQL = strSQL & ",CARD_LIQ_NOTE"
            strSQL = strSQL & ",CARD_LIQ_FLG"
            strSQL = strSQL & ",CARD_SPAREKEY_NOTE"
            strSQL = strSQL & ",CARD_KEY_LENTNUM"
            strSQL = strSQL & ",CARD_KEY_RETNUM"
            strSQL = strSQL & ",CARD_KEY_CHGDATE"
            strSQL = strSQL & ",CARD_KEY_CHGTANTO"
            strSQL = strSQL & ",CARD_KEY_CHGRIYUCD"
            strSQL = strSQL & ",CARD_KEY_CHGNOTE"
            strSQL = strSQL & ",CARD_SEIKI"
            strSQL = strSQL & ",CARD_INSED"
            strSQL = strSQL & ",CARD_INSEJ"
            strSQL = strSQL & ",CARD_INSPB"
            strSQL = strSQL & ",CARD_INSUB"
            strSQL = strSQL & ",CARD_UPDAD"
            strSQL = strSQL & ",CARD_UPDAJ"
            strSQL = strSQL & ",CARD_UPDPB"
            strSQL = strSQL & ",CARD_UPDUB"
            strSQL = strSQL & ",CARD_DELED"
            strSQL = strSQL & ",CARD_DELEJ"
            strSQL = strSQL & ",CARD_DELPB"
            strSQL = strSQL & ",CARD_DELUB"
            strSQL = strSQL & ") "
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & " CARG_YCODE"
            strSQL = strSQL & ",CARG_NO"
            strSQL = strSQL & ",CARG_UCODE"
            strSQL = strSQL & ",CARG_AGRE"
            strSQL = strSQL & ",CARG_FSDATE"
            strSQL = strSQL & ",CARG_STDATE"
            strSQL = strSQL & ",CARG_EDDATE"
            strSQL = strSQL & ",CARG_CYDATE"
            strSQL = strSQL & ",CARG_KYDATE"
            strSQL = strSQL & ",CARG_CONTNA"
            strSQL = strSQL & ",CARG_USAGE"
            strSQL = strSQL & ",CARG_DOCU1"
            strSQL = strSQL & ",CARG_DOCU2"
            strSQL = strSQL & ",CARG_TUBOSU"
            strSQL = strSQL & ",CARG_TTANKA"
            strSQL = strSQL & ",CARG_RENTKG"
            strSQL = strSQL & ",CARG_SYOZEI"
            strSQL = strSQL & ",CARG_FRSTKG"
            strSQL = strSQL & ",CARG_FSYOZEI"
            strSQL = strSQL & ",CARG_FRST_BILL"
            strSQL = strSQL & ",CARG_PREMTH_SUM"
            strSQL = strSQL & ",CARG_SECUKG"
            strSQL = strSQL & ",CARG_BIKO"
            strSQL = strSQL & ",CARG_UPDATE"
            strSQL = strSQL & ",CARG_HDATE"
            strSQL = strSQL & ",CARG_MDATE"
            strSQL = strSQL & ",CARG_ACPTNO"
            strSQL = strSQL & ",CARG_CAMPC"
            strSQL = strSQL & ",CARG_HOSYICD"
            strSQL = strSQL & ",CARG_HOSYO_CD"
            strSQL = strSQL & ",CARG_KAGIICD"
            strSQL = strSQL & ",CARG_DAHIB"
            strSQL = strSQL & ",CARG_HOKEB"
            strSQL = strSQL & ",CARG_KAIJB"
            strSQL = strSQL & ",CARG_HOKAI"
            strSQL = strSQL & ",CARG_HOSYD"
            strSQL = strSQL & ",CARG_HOSYB"
            strSQL = strSQL & ",CARG_ADMEI"
            strSQL = strSQL & ",CARG_KEITI"
            strSQL = strSQL & ",CARG_KAGIA"
            strSQL = strSQL & ",CARG_UKNO"
            strSQL = strSQL & ",CARG_UKTANTO"
            strSQL = strSQL & ",CARG_KAIUKEDATE"
            strSQL = strSQL & ",CARG_KAITANTO"
            strSQL = strSQL & ",CARG_KAICD"
            strSQL = strSQL & ",CARG_KAIRIYUCD"
            strSQL = strSQL & ",CARG_CMP_EXDATE"
            strSQL = strSQL & ",CARG_KEY_RETDATE"
            strSQL = strSQL & ",CARG_SPAREKEY"
            strSQL = strSQL & ",CARG_MNT_REQDATE"
            strSQL = strSQL & ",CARG_MNT_CMPDATE"
            strSQL = strSQL & ",CARG_NIRCD"
            strSQL = strSQL & ",CARG_SYOCD"
            strSQL = strSQL & ",CARG_HOSCD"
            strSQL = strSQL & ",CARG_MNT_TANTO"
            strSQL = strSQL & ",CARG_MNT_MEMO"
            strSQL = strSQL & ",CARG_LIQ_REQDATE"
            strSQL = strSQL & ",CARG_LIQ_NOTE"
            strSQL = strSQL & ",CARG_LIQ_FLG"
            strSQL = strSQL & ",CARG_SPAREKEY_NOTE"
            strSQL = strSQL & ",CARG_KEY_LENTNUM"
            strSQL = strSQL & ",CARG_KEY_RETNUM"
            strSQL = strSQL & ",CARG_KEY_CHGDATE"
            strSQL = strSQL & ",CARG_KEY_CHGTANTO"
            strSQL = strSQL & ",CARG_KEY_CHGRIYUCD"
            strSQL = strSQL & ",CARG_KEY_CHGNOTE"
            strSQL = strSQL & ",CARG_SEIKI"
            strSQL = strSQL & ",CARG_INSED"
            strSQL = strSQL & ",CARG_INSEJ"
            strSQL = strSQL & ",CARG_INSPB"
            strSQL = strSQL & ",CARG_INSUB"
            strSQL = strSQL & ",CARG_UPDAD"
            strSQL = strSQL & ",CARG_UPDAJ"
            strSQL = strSQL & ",CARG_UPDPB"
            strSQL = strSQL & ",CARG_UPDUB"
            strSQL = strSQL & ",'" & Format$(DATE, "yyyymmdd") & "'"
            strSQL = strSQL & ",'" & Format$(time, "hhmmss") & "'"
            strSQL = strSQL & ",'" & P_PROG_ID & "'"
            strSQL = strSQL & ",'" & strUser & "'"
            strSQL = strSQL & " FROM CARG_FILE "
            strSQL = strSQL & " WHERE CARG_ACPTNO = '" & rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE & "' "
            aDbObject.Execute strSQL
            
            strSQL = ""
            strSQL = strSQL & "DELETE FROM CARG_FILE "
            strSQL = strSQL & "WHERE CARG_ACPTNO = '" & rsRcpt.Fields("RCPT_CARG_ACPTNO").VALUE & "' "
            aDbObject.Execute strSQL
        End If
    End If

ErrorHandler2:
    If Not rsRcpt Is Nothing Then rsRcpt.Close: Set rsRcpt = Nothing
ErrorHandler:

    If Err <> 0 Then
        fncDel_REQU_FILE = False
        Call Err.Raise(Err.Number, "fncDel_REQU_FILE" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :UkeNo_Numbering
'        機能             :受付番号採番処理
'        IN               :strNo = 識別(1)+部門ｺｰﾄﾞ(1)+受付年(2)+受付月(1)+ｼｰｹﾝｽ(6)[11桁]
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function UkeNo_Numbering(ByRef strNO As String) As Boolean

    Dim strUkeNo    As String
    Dim strSEQ      As String
    Dim strYY       As String
    Dim strMM       As String
    Dim strSQL      As String
    Dim lngUke      As String
    Dim rsCon       As Recordset

    UkeNo_Numbering = False

    '識別
    strUkeNo = P_UKE_U

    '部門ｺｰﾄﾞ(ｺﾝﾄﾛｰﾙﾏｽﾀにある部門ｺｰﾄﾞを使用する)
    strSQL = "SELECT CONT_BUMOC, CONT_YOYA_NO, CONT_UPDATE FROM CONT_MAST "                 ' 2005/10/06 ADD T.SUZUKI
    Set rsCon = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)                            ' 2005/10/06 ADD T.SUZUKI
'    strSQL = "SELECT CONT_BUMOC,CONT_YOYA_NO,CONT_UPDATE FROM dbo_CONT_MAST "              ' 2005/10/06 DEL T.SUZUKI
'    Set rsCon = objDb.OpenRecordset(strSQL, dbOpenDynaset)                                 ' 2005/10/06 DEL T.SUZUKI

    If rsCon.EOF Then
        rsCon.Close
        Set rsCon = Nothing
        Exit Function
    Else
        With rsCon
            .Edit
            .Fields("CONT_YOYA_NO") = .Fields("CONT_YOYA_NO") + 1   '受付ｼｰｹﾝｽ№CountUp
            .Fields("CONT_UPDATE") = Now()                          '更新日
            strSEQ = CStr(Format(.Fields("CONT_YOYA_NO"), "000000")) '受付№保持
            strUkeNo = strUkeNo & .Fields("CONT_BUMOC")              '部門ｺｰﾄﾞｾｯﾄ
            .UPDATE
            .Close
        End With
        Set rsCon = Nothing
    End If

    '受付年
    Select Case Format(DATE, "yyyy")
        Case "1998": strYY = "&&"
        Case "1999": strYY = "**"
        Case Else
            If Mid$(Trim$(Format$(DATE, "yyyy")), 1, 2) = "20" Then
                strYY = Mid$(Trim$(Format$(DATE, "yyyy")), 3)
            End If
    End Select
    strUkeNo = strUkeNo & strYY

    '受付月
    Select Case Format(DATE, "m")
        Case "10": strMM = "A"
        Case "11": strMM = "B"
        Case "12": strMM = "C"
        Case Else
            strMM = Trim$(Format$(DATE, "m"))
    End Select
    strUkeNo = strUkeNo & strMM

    'ｼｰｹﾝｽ№ｾｯﾄ
    strUkeNo = strUkeNo & strSEQ

    strNO = strUkeNo

    GoTo UkeNo_Numbering_End

UkeNo_Numbering_Err:
    Call Err.Raise(Err.Number, "UkeNo_Numbering" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)  ' 2006/04/07 DEL T.SUZUKI
    Exit Function

UkeNo_Numbering_End:
    UkeNo_Numbering = True
End Function
'==============================================================================*
'
'        MODULE_NAME      :GetCntaData
'        機能             :コンテナマスタ問い合わせ
'        IN               :iMode=0(存在ﾁｪｯｸ)/1(ﾃﾞｰﾀ参照)
'        OUT              :取得結果
'        INDATE           :2005/07/17 K.MITSUNO
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetCntaData(ByVal iMode As Integer) As Boolean
    
    Dim stWHERE         As String
    Dim vCode           As Variant
    Dim vRet            As Variant
    Dim strSQL          As String

    On Error GoTo GetCntaData_Err

    GetCntaData = False

    vCode = Nz(Me.cmb_No, "")

    If vCode = "" Then
        GetCntaData = True
        GoTo GetCntaData_End
    End If
    
    Select Case iMode
        Case 0
            '存在チェック
            strSQL = ""
            strSQL = " CNTA_CODE =" & Val(Me.cmb_Yard)
            strSQL = strSQL & " AND CNTA_NO =" & Val(vCode)
            strSQL = strSQL & " AND CNTA_USE <> 9"                        '2006/09/26 tajima add
            vRet = Nz(DLookup("CNTA_STEP", "dbo_CNTA_MAST", strSQL), "")

            If vRet = "" Then
                MsgBox "コンテナ登録がされておりません", vbInformation, P_KAKUNIN
                GoTo GetCntaData_End
            End If
           '▼ 2006/09/26 add tajima ▼
            vRet = Nz(DLookup("CNTA_REPR_KBN", "dbo_CNTA_MAST", strSQL), 0)
            If vRet = 1 Then
                MsgBox "このコンテナは現在修理中のため、予約を行うことが出来ません。", vbInformation, P_KAKUNIN
                GoTo GetCntaData_End
            End If
           '▲ 2006/09/26 add tajima ▲
        Case 1
            '参照
            '予約受付ではｺﾝﾃﾅ番号によりﾃﾞｰﾀｾｯﾄは使用しないので実装なし
    End Select
    
    GetCntaData = True
    GoTo GetCntaData_End
    
GetCntaData_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

GetCntaData_End:
    Exit Function
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :GetYardData
'        機能             :ヤードマスタ問い合わせ
'        IN               :iMode=0(存在ﾁｪｯｸ)/1(ﾃﾞｰﾀ参照)
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetYardData(ByVal iMode As Integer) As Boolean
    
    Dim strSQL  As String
    Dim vRet    As Variant
    Dim i       As Integer
    Dim vRET2   As Variant
    Dim rsYard  As Recordset

    On Error GoTo GetYardData_Err

    GetYardData = False

    If Me.cmb_Yard = "" Then
        GetYardData = True
        GoTo GetYardData_End
    End If

    doCmd.SetWarnings False

    If iMode = 0 Then
        '存在チェック
        strSQL = strSQL & " YARD_CODE =" & Val(cmb_Yard.VALUE)
        strSQL = strSQL & " AND Nz(YARD_NONDISP_DAY,""9999/12/31"") > Date()" '2006/09/26 tajima add
        
        vRet = Nz(DLookup("YARD_CODE", "dbo_YARD_MAST", strSQL), "")

        If vRet = "" Then
            MsgBox "該当するヤードマスタがありません。", vbInformation, P_KAKUNIN
            GoTo GetYardData_End
        End If

        '▼ 2006/09/26 add tajima ▼
        '営業終了日を過ぎたかのチェック
        vRet = Nz(DLookup("YARD_RENTEND_DAY", "dbo_YARD_MAST", strSQL), "9999/12/31")
        If vRet < DATE Then
            MsgBox "このヤードは営業を終了しているため、予約を行うことは出来ません。", vbInformation, P_KAKUNIN
            GoTo GetYardData_End
        End If
        '▲ 2006/09/26 add tajima ▲
    Else
        ' コンテナDBに接続
        If fncConnectDB = False Then Exit Function                                                ' 2005/11/02 ADD T.SUZUKI

        strSQL = strSQL & "SELECT "
        strSQL = strSQL & "YARD_NAME,"
        strSQL = strSQL & "YARD_ADDR_1,"
        strSQL = strSQL & "YARD_ADDR_2, "
        strSQL = strSQL & "YARD_ADDR_3 "
        strSQL = strSQL & ",YARD_BEGIN_DAY,YARD_RENTEND_DAY "                                     ' 2006/09/26 add tajima
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "YARD_MAST "                                                            ' 2005/10/06 rep T.SUZUKI
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "YARD_CODE =" & Val(cmb_Yard.VALUE)
        Set rsYard = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)          ' 2005/10/06 rep T.SUZUKI

        If Not (rsYard.BOF And rsYard.EOF) Then
            With Me
                .txt_Yard = rsYard.Fields("YARD_NAME")
                .txt_Yadd = rsYard.Fields("YARD_ADDR_1") & rsYard.Fields("YARD_ADDR_2") & rsYard.Fields("YARD_ADDR_3")
                '▼ 2006/09/26 add tajima ▼
                .txt_営業期間 = ""
                If Nz(rsYard.Fields("YARD_BEGIN_DAY"), "") <> "" Then
                    str営業開始日 = Format$(rsYard.Fields("YARD_BEGIN_DAY"), "YYYY/MM/DD")
                    .txt_営業期間 = str営業開始日
                Else
                    str営業開始日 = ""
                End If
                .txt_営業期間 = .txt_営業期間 & "～"
                If Nz(rsYard.Fields("YARD_RENTEND_DAY"), "") <> "" Then
                    str営業終了日 = Format$(rsYard.Fields("YARD_RENTEND_DAY"), "YYYY/MM/DD")
                    .txt_営業期間 = .txt_営業期間 & str営業終了日
                Else
                    str営業終了日 = pcstr最大年月日
                End If
                '▲ 2006/09/26 add tajima ▲
                
                '▼ 2018/08/30 TAKENOUCHI UPD<S> ▼
                '▼ 2007/04/06 add tajima ▼
                'If str営業開始日 > Format$(DATE, "YYYY/MM/DD") Then
                '    chk_Kinrin.VALUE = False
                'Else
                '    chk_Kinrin.VALUE = True
                'End If
                '▲ 2007/04/06 add tajima ▲
                chk_Kinrin.VALUE = False
                '▲ 2018/08/30 TAKENOUCHI UPD<E> ▲
           End With
        End If
        rsYard.Close
        Set rsYard = Nothing

        ' コンテナDB切断
        If fncDisConnectDB = False Then GoTo GetYardData_End                                      ' 2005/11/02 ADD T.SUZUKI
    End If

    GetYardData = True
    GoTo GetYardData_End

GetYardData_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

GetYardData_End:
    If Not rsYard Is Nothing Then rsYard.Close: Set rsYard = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing               ' 2005/11/02 ADD T.SUZUKI

    doCmd.SetWarnings True
End Function
'==============================================================================*
'
'        MODULE_NAME      :Chk_UserMast
'        機能             :顧客マスタ問い合わせ
'        IN               :iMode=0(存在ﾁｪｯｸ)/1(ﾃﾞｰﾀ参照)
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_UserMast(ByVal iMode As Integer) As Boolean

    Dim strWhere        As String
    Dim vCode           As Variant
    Dim vRet            As Variant
    Dim strSQL          As String
    Dim rsKok           As Recordset

    On Error GoTo Chk_UserMast_Err

    Chk_UserMast = False

    vCode = Nz(Me.cmb_KokCd, "")

    If vCode = "" Then
        Chk_UserMast = True
        GoTo Chk_UserMast_End
    End If

    doCmd.SetWarnings False

    If iMode = 0 Then
        strWhere = ""
        strWhere = " USER_CODE =" & Val(vCode)
        vRet = Nz(DLookup("USER_NAME", "dbo_USER_MAST", strWhere), "")

        If vRet = "" Then
            MsgBox "顧客マスタが登録されておりません", vbInformation, P_KAKUNIN
            GoTo Chk_UserMast_End
        End If

    Else
        ' コンテナDBに接続
        If fncConnectDB = False Then Exit Function                                                              ' 2005/11/02 ADD T.SUZUKI

        Set rsKok = pobjKONT_DB.OpenRecordset(SEL_USER_QUERY(vCode), dbOpenSnapshot, dbSQLPassThrough)          ' 2005/10/06 ADD T.SUZUKI
'        Set rsKok = objDb.OpenRecordset(SEL_USER_QUERY(vCode), dbOpenSnapshot)                                 ' 2005/10/06 DEL T.SUZUKI

        If Not (rsKok.BOF And rsKok.EOF) Then
            With Me
'                .cmb_Kokya.Requery  ' 2005/10/06 DEL T.SUZUKI この処理の代替えは？？？(確認)
                .txt_Name = rsKok("USER_NAME")       '顧客名
                .txt_Kana = rsKok("USER_KANA")       '顧客名カナ
                .txt_Yubin = rsKok("USER_YUBINO")    '郵便番号
                .txt_Adr1 = rsKok("USER_ADR_1")      '住所１
                .txt_Adr2 = rsKok("USER_ADR_2")      '住所２
                .txt_Adr3 = rsKok("USER_ADR_3")      '住所３  ' 2005/10/06 ADD T.SUZUKI
'                .txt_Adr3 = rsKok("USER_ADR_3")     '住所３　☆★MSTにないので後ほどコメントアウト
                .txt_Tanm = rsKok("USER_TANM")       '代表社名
                .txt_Taka = rsKok("USER_TAKA")       '代表者カナ
                .txt_Tel = rsKok("USER_TEL")         '電話番号
                .txt_Keitai = rsKok("USER_KEITAI")   '携帯番号
                .txt_Fax = rsKok("USER_FAX")         'FAX番号
                .txt_Mail = rsKok("USER_MAIL")       'メール  ' 2005/10/06 ADD T.SUZUKI
                '.txt_Mail = rsKok("USER_MAIL")      'メール　☆★MSTにないので後ほどコメントアウト
                .cmb_Kkbn = rsKok("USER_KKBN")      '顧客区分
                .txt_Kkbn = GetNameData(P_NAME_020, rsKok("USER_KKBN")) '顧客区分名

' 2005/10/06 ADD T.SUZUKI Start
                If Nz(.txt_Kokya, "") <> rsKok("USER_KANA") Then
                    .txt_Kokya = rsKok("USER_KANA")
                End If
' 2005/10/06 ADD T.SUZUKI End
            End With
        Else
            With Me
'                .cmb_Kokya.Requery  ' 2005/10/06 DEL T.SUZUKI この処理の代替えは？？？(確認)
                .txt_Name = ""      '顧客名
                .txt_Kana = ""      '顧客名カナ
                .txt_Yubin = ""     '郵便番号
                .txt_Adr1 = ""      '住所１
                .txt_Adr2 = ""      '住所２
                .txt_Tanm = ""      '代表社名
                .txt_Taka = ""      '代表者カナ
                .txt_Tel = ""       '電話番号
                .txt_Keitai = ""    '携帯番号
                .txt_Fax = ""       'FAX番号
                '顧客区分
                .cmb_Kkbn.VALUE = P_KKBN_1
                '顧客区分名
                .txt_Kkbn = GetNameData(P_NAME_020, Nz(.cmb_Kkbn, ""))

' 2005/10/06 ADD T.SUZUKI Start
                .txt_Kokya = ""
' 2005/10/06 ADD T.SUZUKI End
            End With
        End If
        rsKok.Close
        Set rsKok = Nothing

        ' コンテナDB切断
        If fncDisConnectDB = False Then GoTo Chk_UserMast_End          ' 2005/11/02 ADD T.SUZUKI
    End If

    Chk_UserMast = True
    GoTo Chk_UserMast_End
    
Chk_UserMast_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    
Chk_UserMast_End:
    If Not rsKok Is Nothing Then rsKok.Close: Set rsKok = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing               ' 2005/11/02 ADD T.SUZUKI

    doCmd.SetWarnings True
End Function
'==============================================================================*
'
'        MODULE_NAME      :SEL_USER_QUERY
'        機能             :SEL_USER_QUERY
'        IN               :varKOKYC = 顧客ｺｰﾄﾞ
'        OUT              :顧客取得SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function SEL_USER_QUERY(varKOKYC As Variant) As String

    Dim strSQL  As String

    strSQL = ""
    strSQL = strSQL & "SELECT * "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "USER_MAST "               ' 205/10/06 ADD T.SUZUKI
'    strSQL = strSQL & "dbo_USER_MAST "          ' 205/10/06 DEL T.SUZUKI
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "USER_CODE =" & Val(varKOKYC)
    strSQL = strSQL & ";"

    SEL_USER_QUERY = strSQL
End Function
'==============================================================================*
'
'        MODULE_NAME      :fncGetCountINTR_TRAN
'        機能             :予約ご紹介トランの件数取得
'        IN               :
'        OUT              :
'        CREATE           :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetCountINTR_TRAN(ByVal strUkno As String) As Boolean
'Private Function fncGetCountINTR_TRAN(ByVal strUkNo As String) As Integer

    Dim strSQL       As String
    Dim objRs        As Recordset
    Dim intDataCount As Integer                                    ' 205/12/20 ADD T.SUZUKI

    On Error GoTo fncGetCountINTR_TRAN_Err

    fncGetCountINTR_TRAN = False                                   ' 205/12/20 ADD T.SUZUKI
'    fncGetCountINTR_TRAN = 0                                      ' 205/12/20 DEL T.SUZUKI

    ' 件数初期化
    Me.txt_Syoukai = 0                                             ' 205/12/20 ADD T.SUZUKI

    If Nz(strUkno, "") = "" Then Exit Function

    ' コンテナDBに接続
    If fncConnectDB = False Then Exit Function

    ' 予約ご紹介トランのデータ件数取得
    strSQL = "SELECT COUNT(*) "
    strSQL = strSQL & " FROM INTR_TRAN "
    strSQL = strSQL & " WHERE INTRT_UKNO = '" & strUkno & "' "  ' 予約受付番号
    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    ' 予約ご紹介トランの件数セット
    intDataCount = 0
    intDataCount = Nz(objRs.Fields(0).VALUE, 0)
    Me.txt_Syoukai = Format(intDataCount, "#,##0")
'    fncGetCountINTR_TRAN = Nz(objRs.Fields(0).Value, 0)

    objRs.Close
    Set objRs = Nothing

    ' コンテナDB切断
    If fncDisConnectDB = False Then GoTo fncGetCountINTR_TRAN_Exit

    fncGetCountINTR_TRAN = True                                    ' 205/12/20 ADD T.SUZUKI

fncGetCountINTR_TRAN_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    Exit Function

fncGetCountINTR_TRAN_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncGetCountINTR_TRAN_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :subCreateWork
'        機能             :サブフォーム表示用データ作成
'        IN               :
'        OUT              :
'        CREATE           :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreateWork()

    Dim strSQL    As String
    Dim objRs     As Recordset
    Dim objRsWk   As Recordset
    Dim blnConFlg As Boolean

    On Error GoTo subCreateWork_Err

    blnConFlg = True

    ' 顧客コードが未選択の場合は処理を抜ける
    If Nz(Me.cmb_KokCd, "") = "" Then Exit Sub                       ' 2006/01/12 ADD T.SUZUKI

    If pobjKONT_DB Is Nothing Then
        ' コンテナDBに接続
        If fncConnectDB = False Then Exit Sub
        blnConFlg = False
    End If

    ' SQL文実行
    strSQL = " SELECT   CARG_FILE.CARG_ACPTNO, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_CONTNA, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_YCODE, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_NO, " & Chr(13)
    strSQL = strSQL & " YARD_MAST.YARD_NAME, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_RENTKG, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_BIKO, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_FSDATE, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_STDATE, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_KYDATE, " & Chr(13)
    strSQL = strSQL & " CARG_FILE.CARG_AGRE " & Chr(13)              ' 2006/01/20 ADD T.SUZUKI
    strSQL = strSQL & " FROM CARG_FILE, " & Chr(13)                  ' *** コンテナ契約ファイル ***
    strSQL = strSQL & "      YARD_MAST " & Chr(13)                   ' *** ヤードマスタ ***
    strSQL = strSQL & "     ,RCPT_TRAN " & Chr(13)                              'INSERT 2020/02/28 N.IMAI
    strSQL = strSQL & "     ,YOUK_TRAN " & Chr(13)                              'INSERT 2020/02/28 N.IMAI
    strSQL = strSQL & " WHERE CARG_FILE.CARG_YCODE = YARD_MAST.YARD_CODE " & Chr(13)
    strSQL = strSQL & "   AND CARG_UCODE = " & Val(Me.cmb_KokCd) & Chr(13)
    strSQL = strSQL & "   AND RCPT_CARG_ACPTNO = CARG_ACPTNO" & Chr(13)         'INSERT 2020/02/28 N.IMAI
    strSQL = strSQL & "   AND YOUKT_UKNO = RCPT_NO" & Chr(13)                   'INSERT 2020/02/28 N.IMAI
    strSQL = strSQL & "   AND YOUKT_YUKBN = 20" & Chr(13)                       'INSERT 2020/02/28 N.IMAI
    strSQL = strSQL & " ORDER BY CARG_FILE.CARG_ACPTNO DESC "
    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    If objRs.EOF = False Then
        ' ワークテーブルオープン
        strSQL = "SELECT * FROM FVS500_W01"
        Set objRsWk = objDb.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

        Do Until objRs.EOF
            With objRsWk
                .AddNew

                .Fields("CARG_ACPTNO").VALUE = objRs.Fields("CARG_ACPTNO").VALUE  ' 受注契約番号
                .Fields("CARG_CONTNA").VALUE = objRs.Fields("CARG_CONTNA").VALUE  ' ｺﾝﾃﾅ状態
                .Fields("CARG_YCODE").VALUE = objRs.Fields("CARG_YCODE").VALUE    ' ﾔｰﾄﾞｺｰﾄﾞ
                .Fields("CARG_NO").VALUE = objRs.Fields("CARG_NO").VALUE          ' ｺﾝﾃﾅ番号
                .Fields("YARD_NAME").VALUE = objRs.Fields("YARD_NAME").VALUE      ' ヤード名
                .Fields("CARG_RENTKG").VALUE = objRs.Fields("CARG_RENTKG").VALUE  ' 月額料金
                .Fields("CARG_BIKO").VALUE = objRs.Fields("CARG_BIKO").VALUE      ' 備考
                .Fields("CARG_FSDATE").VALUE = objRs.Fields("CARG_FSDATE").VALUE  ' 初回契約日
                .Fields("CARG_STDATE").VALUE = objRs.Fields("CARG_STDATE").VALUE  ' 契約開始日
                .Fields("CARG_KYDATE").VALUE = objRs.Fields("CARG_KYDATE").VALUE  ' 解約日付
                .Fields("CARG_AGRE").VALUE = objRs.Fields("CARG_AGRE").VALUE      ' 契約状態     ' 2006/01/20 ADD T.SUZUKI

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

    If blnConFlg = False Then
        ' コンテナDB切断
        If fncDisConnectDB = False Then GoTo subCreateWork_Exit
    End If

subCreateWork_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not objRsWk Is Nothing Then objRsWk.Close: Set objRsWk = Nothing
    If blnConFlg = False Then
        If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing
    End If
    Exit Sub

subCreateWork_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo subCreateWork_Exit
End Sub
'==============================================================================*
'
'        MODULE_NAME      :fncDeleteWorkTable
'        機能             :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ削除
'        IN               :
'        OUT              :
'        CREATE           :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDeleteWorkTable() As Boolean

    Dim strSQL As String

    On Error GoTo fncDeleteWorkTable_Err

    fncDeleteWorkTable = False

    ' テーブルの存在チェック
    If fncChkTableExist(pstrCstWK_TABLE_NM) = True Then
        ' ｻﾌﾞﾌｫｰﾑのｿｰｽｵﾌﾞｼﾞｪｸﾄ無効化
        Me.FVS500_FVS500S01.SourceObject = ""

        ' テーブル削除
        objDb.TableDefs.Delete pstrCstWK_TABLE_NM
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

fncDeleteWorkTable_Exit:
    Exit Function

fncDeleteWorkTable_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncDeleteWorkTable_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :fncCHKTableExist
'        機能             :テーブルの存在チェック
'        IN               :
'        OUT              :
'        CREATE           :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkTableExist(strTableName As String) As Boolean

    Dim tdf As TableDef

    On Error GoTo fncCHKTableExist_Err

    fncChkTableExist = False

    For Each tdf In objDb.TableDefs
        If tdf.NAME = strTableName Then
            fncChkTableExist = True
            Exit For
        End If
    Next tdf

fncCHKTableExist_Exit:
    Exit Function

fncCHKTableExist_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncCHKTableExist_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :fncCreateWorkTable
'        機能             :サブフォーム表示用ﾜｰｸﾃｰﾌﾞﾙ作成
'        IN               :
'        OUT              :
'        CREATE           :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncCreateWorkTable() As Boolean

    Dim tdfNew    As TableDef
    Dim intCount  As Integer

    On Error GoTo fncCreateWorkTable_Err

    fncCreateWorkTable = False

    Set tdfNew = objDb.CreateTableDef(pstrCstWK_TABLE_NM)
    With tdfNew
        .Fields.Append .CreateField("CARG_ACPTNO", dbText, 10)        ' 受注契約番号
        .Fields.Append .CreateField("CARG_CONTNA", dbInteger)         ' ｺﾝﾃﾅ状態
        .Fields.Append .CreateField("CARG_YCODE", dbLong)             ' CARG_YCODE
        .Fields.Append .CreateField("CARG_NO", dbLong)                ' ｺﾝﾃﾅ番号
        .Fields.Append .CreateField("YARD_NAME", dbText, 36)          ' ﾔｰﾄﾞ名
        .Fields.Append .CreateField("CARG_RENTKG", dbLong)            ' 月額料金
        .Fields.Append .CreateField("CARG_BIKO", dbText, 60)          ' 備考
        .Fields.Append .CreateField("CARG_FSDATE", dbDate)            ' 初回契約日
        .Fields.Append .CreateField("CARG_STDATE", dbDate)            ' 契約開始日
        .Fields.Append .CreateField("CARG_KYDATE", dbDate)            ' 解約日付

        .Fields.Append .CreateField("CARG_ACPTNO_Dummy", dbText, 10)  ' 受注契約番号ダミー     ' 2006/01/20 ADD T.SUZUKI
        .Fields.Append .CreateField("CARG_AGRE", dbInteger)           ' 契約状態               ' 2006/01/20 ADD T.SUZUKI


        For intCount = 0 To .Fields.Count - 1
            If .Fields(intCount).Type = dbText Then
                .Fields(intCount).AllowZeroLength = True
            End If
        Next intCount
    End With

    ' テーブルの追加
    objDb.TableDefs.Append tdfNew

    ' ｻﾌﾞﾌｫｰﾑのｿｰｽｵﾌﾞｼﾞｪｸﾄ設定
    Me.FVS500_FVS500S01.SourceObject = pstrCstSubFromName

    fncCreateWorkTable = True

fncCreateWorkTable_Exit:
    Exit Function

fncCreateWorkTable_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    GoTo fncCreateWorkTable_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :紹介トランの更新処理への分岐
'        MODULE_ID        :fncUpdINTR_TRAN
'        Parameter        :strYOUKT_UKNO = 予約受付番号
'        CREATE_DATE      :2005/12/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdINTR_TRAN(ByVal strYOUKT_UKNO As String) As Boolean

    Dim strErrMsg As String

    On Error GoTo fncUpdINTR_TRAN_Err

    fncUpdINTR_TRAN = False

    ' ※pstrYUKBN：読み込まれた予約受付状態区分
    strErrMsg = ""
    ' → 新規 または 修正時
    If (mdl_Mode = P_NEW Or mdl_Mode = P_SHU) Then
        Select Case pstrYUKBN
            Case ""  ' *** 新規 ***
                Select Case Me.cmb_Yukbn
                    Case P_YUKBN_00  ' 自動紹介
                        ' 何もしない

                    Case P_YUKBN_01  ' 自社理由
                        ' 処理A.指定ｺﾝﾃﾅにて【01：取置】の紹介ﾄﾗﾝを作る
                        If CmnInsertIntrTran(pobjKONT_DB, strYOUKT_UKNO, _
                                             CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                             P_INTROKBN_01, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If

                    Case P_YUKBN_02  ' 取置
                        ' 処理A.指定ｺﾝﾃﾅにて【01：取置】の紹介ﾄﾗﾝを作る
                        If CmnInsertIntrTran(pobjKONT_DB, strYOUKT_UKNO, _
                                             CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                             P_INTROKBN_01, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If

                    Case P_YUKBN_10  ' 受付
                        ' 処理D'.指定ｺﾝﾃﾅにて【02：契約受付】の紹介ﾄﾗﾝを作る
                        If CmnInsertIntrTran(pobjKONT_DB, strYOUKT_UKNO, _
                                             CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                             P_INTROKBN_02, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If
                End Select
                ' -------------------------------------------------------------------

            Case P_YUKBN_00  ' 自動紹介
                Select Case Me.cmb_Yukbn
                    Case P_YUKBN_00  ' 自動紹介
                        ' 何もしない

                    Case P_YUKBN_02  ' 取置
                        ' 処理A.指定ｺﾝﾃﾅにて【01：取置】の紹介ﾄﾗﾝを作る
                        If CmnInsertIntrTran(pobjKONT_DB, strYOUKT_UKNO, _
                                             CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                             P_INTROKBN_01, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If

                    Case P_YUKBN_10  ' 受付
                        ' 処理D'.指定ｺﾝﾃﾅにて【02：契約受付】の紹介ﾄﾗﾝを作る
                        If CmnInsertIntrTran(pobjKONT_DB, strYOUKT_UKNO, _
                                             CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                             P_INTROKBN_02, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If

                    Case P_YUKBN_51  ' 予約キャンセル
                        ' 何もしない※キャンセルすべき紹介ﾄﾗﾝはないはず
                End Select
                ' -------------------------------------------------------------------

            Case P_YUKBN_01  ' 自社理由
                Select Case Me.cmb_Yukbn
                    Case P_YUKBN_01  ' 自社理由
                        ' 取置での保存時、読み込んだときのｺﾝﾃﾅ番号と画面で選択されている
                        ' ｺﾝﾃﾅ番号が異なるときに下記の処理を行う
                        If Me.cmb_No <> pstrKONT_NO Then
                            
                          '2006/05/02 紹介待ち障害対応 Rep tajima Start
                          ' 指定した新コンテナで新たな紹介を作り、旧コンテナはキャンセル&紹介待ちに
                          If IntroChgAndIntroWait(pobjKONT_DB, _
                                     strYOUKT_UKNO, _
                                     P_INTROKBN_01, _
                                     strErrMsg, _
                                     CLng(Me.cmb_Yard), _
                                     CLng(Me.cmb_No)) = False Then
                                GoTo fncUpdINTR_TRAN_UpdErr
                            End If
                          '2006/05/02 Rep tajima End

                          ' 2005/12/20 ADD T.SUZUKI Start
                            ' 変更した値をリセットする
                            pstrKONT_NO = Me.cmb_No   ' ｺﾝﾃﾅ番号
                            pstrYUKBN = Me.cmb_Yukbn  ' 予約受付状態区分
                          ' 2005/12/20 ADD T.SUZUKI End
                        Else
                           ' 変更がなければ何もしない
                        End If

                    Case P_YUKBN_10  ' 受付
                        
                        '2006/05/02 紹介待ち障害対応 Rep tajima Start
                          ' コンテナが変更されていれば新コンテナで受付作成、旧コンテナはキャンセル&紹介待ちに
                          ' コンテナが変更されていなければ、そのまま受付にする
                          ' 以下の関数で全て行う
                          If IntroChgAndIntroWait(pobjKONT_DB, _
                                     strYOUKT_UKNO, _
                                     P_INTROKBN_02, _
                                     strErrMsg, _
                                     CLng(Me.cmb_Yard), _
                                     CLng(Me.cmb_No)) = False Then
                                GoTo fncUpdINTR_TRAN_UpdErr
                          End If
                        '2006/05/02 紹介待ち障害対応 Rep tajima End
                        ' 2005/12/20 ADD T.SUZUKI Start
                          ' 変更した値をリセットする
                          pstrKONT_NO = Me.cmb_No   ' ｺﾝﾃﾅ番号
                          pstrYUKBN = Me.cmb_Yukbn  ' 予約受付状態区分
                        ' 2005/12/20 ADD T.SUZUKI End

                    Case P_YUKBN_51  ' 予約キャンセル
                        ' ①取置 → キャンセルにする
                        If CmnUpdateIntrTran(pobjKONT_DB, strYOUKT_UKNO, _
                                             P_INTROKBN_01, P_INTROKBN_06, strErrMsg) = False Then
'                        ' 処理B.全ての取置中の紹介ﾄﾗﾝを【06：キャンセル】にする
'                        If IntroWaitAll(pobjKONT_DB, strYOUKT_UKNO, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If
                      ' 2005/12/21 ADD T.SUZUKI Start
                        ' コンテナマスタの更新(紹介待ちにする)
                        ' ※コンテナ番号は読み込んだときの番号とする
                        If WaitUpdateCntaMast(pobjKONT_DB, strYOUKT_UKNO, _
                                              CLng(Me.cmb_Yard), CLng(pstrKONT_NO), _
                                              strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If
                      ' 2005/12/21 ADD T.SUZUKI End
                End Select
                ' -------------------------------------------------------------------

            Case P_YUKBN_02  ' 取置
                Select Case Me.cmb_Yukbn
                    Case P_YUKBN_00  ' 自動紹介
                        ' 処理B.全ての取置中の紹介ﾄﾗﾝを【06：キャンセル】にする
                        If IntroWaitAll(pobjKONT_DB, strYOUKT_UKNO, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If

                    Case P_YUKBN_02  ' 取置
                        ' 取置での保存時、読み込んだときのｺﾝﾃﾅ番号と画面で選択されている
                        ' ｺﾝﾃﾅ番号が異なるときに下記の処理を行う
                        If Me.cmb_No <> pstrKONT_NO Then
                            ' 処理A.指定ｺﾝﾃﾅに【01：取置】の紹介ﾄﾗﾝを作る
                            If CmnInsertIntrTran(pobjKONT_DB, strYOUKT_UKNO, _
                                                 CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                                 P_INTROKBN_01, strErrMsg) = False Then
                                GoTo fncUpdINTR_TRAN_UpdErr
                            End If
                          ' 2005/12/20 ADD T.SUZUKI Start
                            ' 変更した値をリセットする
                            pstrKONT_NO = Me.cmb_No   ' ｺﾝﾃﾅ番号
                            pstrYUKBN = Me.cmb_Yukbn  ' 予約受付状態区分
                          ' 2005/12/20 ADD T.SUZUKI End
                        Else
                            ' 変更がなければ何もしない
                        End If

                    Case P_YUKBN_10  ' 受付
                    
                        '2006/05/02 紹介待ち障害対応 Rep tajima Start
                          ' コンテナが変更されていれば新コンテナで受付作成、旧コンテナはキャンセル&紹介待ちに
                          ' コンテナが変更されていなければ、そのまま受付にする
                          ' 以下の関数で全て行う
                          If IntroChgAndIntroWait(pobjKONT_DB, _
                                     strYOUKT_UKNO, _
                                     P_INTROKBN_02, _
                                     strErrMsg, _
                                     CLng(Me.cmb_Yard), _
                                     CLng(Me.cmb_No)) = False Then
                                GoTo fncUpdINTR_TRAN_UpdErr
                          End If
                        '2006/05/02 紹介待ち障害対応 Rep tajima End
                        ' 2005/12/20 ADD T.SUZUKI Start
                            ' 変更した値をリセットする
                            pstrKONT_NO = Me.cmb_No   ' ｺﾝﾃﾅ番号
                            pstrYUKBN = Me.cmb_Yukbn  ' 予約受付状態区分
                        ' 2005/12/20 ADD T.SUZUKI End

                    Case P_YUKBN_51  ' 予約キャンセル
                        ' 処理B.全ての取置中の紹介ﾄﾗﾝを【06：キャンセル】にする
                        If IntroWaitAll(pobjKONT_DB, strYOUKT_UKNO, strErrMsg) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If
                End Select
                ' -------------------------------------------------------------------

            Case P_YUKBN_10  ' 受付
                Select Case Me.cmb_Yukbn
                    Case P_YUKBN_10  ' 受付
                        ' 何もしない

                    Case P_YUKBN_53  ' 受付キャンセル
                        ' 処理C.契約受付した紹介ﾄﾗﾝを【06：キャンセル】にする
                        If IntroWaitSingle(pobjKONT_DB, strYOUKT_UKNO, P_INTROKBN_06, strErrMsg, _
                                           CLng(Me.cmb_Yard), CLng(Me.cmb_No), P_INTROKBN_02) = False Then
                            GoTo fncUpdINTR_TRAN_UpdErr
                        End If

'↓DELETE 2008/08/18 SHIBAZAKI
'                      ' 2005/12/21 ADD T.SUZUKI Start
'                        ' コンテナマスタの更新(紹介待ちにする)
'                        ' ※コンテナ番号は読み込んだときの番号とする
'                        If WaitUpdateCntaMast(pobjKONT_DB, strYOUKT_UKNO, _
'                                              CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
'                                              strErrMsg) = False Then
'                            GoTo fncUpdINTR_TRAN_UpdErr
'                        End If
'                      ' 2005/12/21 ADD T.SUZUKI End
'↑DELETE 2008/08/18 SHIBAZAKI
                End Select
                ' -------------------------------------------------------------------

            Case P_YUKBN_52  ' 取置キャンセル
                Select Case Me.cmb_Yukbn
                    Case P_YUKBN_00  ' 自動紹介
                        ' 何もしない

                    Case P_YUKBN_51  ' 予約キャンセル
                        ' 何もしない
                End Select
                ' -------------------------------------------------------------------
        End Select

    Else
        ' → 削除時
        ' ご紹介トランのデータを物理削除
        If fncDELETE_INTR_TRAN(pobjKONT_DB, strYOUKT_UKNO) = False Then Exit Function
    End If

    fncUpdINTR_TRAN = True

fncUpdINTR_TRAN_Exit:
    Exit Function

fncUpdINTR_TRAN_UpdErr:
    Call Err.Raise(Err.Number, "fncUpdINTR_TRAN[" & strErrMsg & "]" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)  ' 2006/04/07 DEL T.SUZUKI
    Exit Function

fncUpdINTR_TRAN_Err:
    Call Err.Raise(Err.Number, "fncUpdINTR_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)  ' 2006/04/07 DEL T.SUZUKI
End Function

'■□■□■□■□■□■□■□■□■□ ↓↓↓　入力ﾁｪｯｸ 関数 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□
'==============================================================================*
'
'        MODULE_NAME      :ChkDispData
'        機能             :予約入力可能コンテナチェック
'        IN               :
'        OUT              :取得結果
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ChkDispData() As Boolean

    Dim strWhere    As String
    Dim strMsg      As String
    Dim strSQL      As String
    Dim varWk       As Variant
    Dim ctl         As Control
    Dim objRs       As Recordset

    Dim intRtn      As Integer     ' 2005/10/12 ADD T.SUZUKI
    Dim strErrMsg   As String      ' 2005/10/12 ADD T.SUZUKI

    Dim intCheck    As Integer     ' 2005/12/14 ADD T.SUZUKI
    Dim strName     As String      ' 2006/01/24 ADD T.SUZUKI

' ***** Start 2017/08/07 add by ysuzuki
    Dim strUkno     As String       '予約契約番号
' ***** end 2017/08/07 add by ysuzuki
    Dim aMotoACPTNOCnt As Integer   'INSERT 2019/03/06 Y.WADA

    ChkDispData = False

    '画面項目のチェック
    If Chk_FrmControl(mdl_Mode, Me.cmb_Yukbn, strMsg) = False Then GoTo ChkDispData_Err

' 2006/01/24 ADD T.SUZUKI Start]
    If Me.cmb_MoveKbn <> P_MOVEKBN_00 Then
        ' ●移動区分のチェック

        ' 移動区分名取得
        strName = ""
        strName = GetNameData(P_NAME_098, Me.cmb_MoveKbn)

        Select Case mdl_Mode
            Case P_NEW  ' *** 新規 ***
                ' 新規モード かつ 移動区分が指定されている場合、
                ' 予約受付状態区分が「10：受付」以外は登録不可とする
                If Me.cmb_Yukbn <> P_YUKBN_10 Then
                    MsgBox strName & pcstrMoveErrMsg01, vbExclamation, P_ERROR
                    GoTo ChkDispData_End
                End If

            Case P_SHU  ' *** 修正 ***
                ' 修正モード かつ 移動区分が指定されている場合、
                ' 予約受付状態区分が下記の場合以外は登録不可とする
                ' 「10：受付」
                ' 「20：契約」
                ' 「 53：受付キャンセル」
                If (Me.cmb_Yukbn <> P_YUKBN_10 And Me.cmb_Yukbn <> P_YUKBN_20 And _
                    Me.cmb_Yukbn <> P_YUKBN_53) Then
                    MsgBox strName & pcstrMoveErrMsg01, vbExclamation, P_ERROR
                    GoTo ChkDispData_End
                End If
        End Select

        ' ②移動元契約番号入力チェック
        If Nz(Me.txt_MOTO_ACPTNO, "") = "" Then
            MsgBox pcstrMoveErrMsg02, vbExclamation, P_ERROR
            GoTo ChkDispData_End
        End If

        ' 移動予約可否チェック
        ' コンテナDBに接続
        If fncConnectDB = False Then Exit Function

'DELETE 2019/03/06 Y.WADA Start
'        If CmKomsMod.IsMoveReserveEntry(pobjKONT_DB, Me.cmb_MoveKbn, Nz(Me.txt_MOTO_ACPTNO, ""), _
'                                        Me.cmb_Yard, Me.cmb_No, strErrMsg, Nz(Me.cmb_Ukno, "")) = False Then
'                                   ' 2006/02/05 H.Tajima 修正 除外予約番号の指定を追加↑
'            ' エラーメッセージ出力
'            MsgBox strErrMsg, vbExclamation, P_ERROR
'            GoTo ChkDispData_End
'        End If
'DELETE 2019/03/06 Y.WADA End
        'INSERT 2019/03/06 Y.WADA Start
        If CmKomsMod.IsMoveReserveEntry(pobjKONT_DB, Me.cmb_MoveKbn, Nz(Me.txt_MOTO_ACPTNO, ""), _
                                        Me.cmb_Yard, Me.cmb_No, strErrMsg, aMotoACPTNOCnt, Nz(Me.cmb_Ukno, "")) = False Then
            ' エラーメッセージ出力
            MsgBox strErrMsg, vbExclamation, P_ERROR
            GoTo ChkDispData_End
        End If
        
        If aMotoACPTNOCnt > 0 Then
            '確認ﾒｯｾｰｼﾞ
            If MsgBox("この契約では既に契約変更がされています。" & Chr(13) & "よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
                GoTo ChkDispData_End
            End If
        End If
        
        'INSERT 2019/03/06 Y.WADA End
    
    
    End If
' 2006/01/24 ADD T.SUZUKI End

' 2005/12/12 ADD T.SUZUKI Start
' ***** Start 2017/08/07 add by ysuzuki
'    If Me.cmb_Yukbn = P_YUKBN_01 Or _
'       Me.cmb_Yukbn = P_YUKBN_02 Or _
'       Me.cmb_Yukbn = P_YUKBN_10 Then
    If Me.cmb_Yukbn = P_YUKBN_01 Or _
       Me.cmb_Yukbn = P_YUKBN_02 Or _
       Me.cmb_Yukbn = P_YUKBN_10 Or _
       Me.cmb_Yukbn = P_YUKBN_53 Then
' ***** end 2017/08/07 add by ysuzuki

        If pobjKONT_DB Is Nothing Then                          ' 2006/01/30 ADD T.SUZUKI
            ' コンテナDBに接続
            If fncConnectDB = False Then Exit Function
        End If                                                  ' 2006/01/30 ADD T.SUZUKI
    End If
' 2005/12/12 ADD T.SUZUKI End

    Select Case Me.cmb_Yukbn
        Case P_YUKBN_01, P_YUKBN_02, P_YUKBN_10
            ' *** 自社理由 OR 取置 OR 受付 ***

            intCheck = 0
            If mdl_Mode = P_NEW Then
                ' *** 新規 ***
                If Me.cmb_Yukbn = P_YUKBN_01 Then
                    intCheck = pcintJISYAYOYAKU_CHECK  ' 自社理由
                Else
                    intCheck = pcintYOYAKU_CHECK       ' 取置 OR 受付
                End If

                ' 物件チェック
                intRtn = CheckLending(pobjKONT_DB, intCheck, CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                      strErrMsg, Nz(Me.cmb_Ukno, ""), Nz(Me.txt_MOTO_ACPTNO, ""))

            ElseIf P_SHU Then
                ' *** 修正 ***
                If Me.cmb_Yukbn = P_YUKBN_01 Then
                    intCheck = pcintJISYAYOYAKU_CHECK  ' 自社理由
                Else
                    intCheck = pcintYOYAKU_CHECK       ' 取置 OR 受付
                End If
                
                '2015/08/05 M.HONDA
                'ネットから入ってきたものは、物件のチェックを行わない。
                '予約時にすでに行っているため。
                ' 物件チェック
                If cmb_発生区分 <> "01" Then
                    If Nz(Me.cmb_No, "") <> "" Then
                        intRtn = CheckLending(pobjKONT_DB, intCheck, CLng(Me.cmb_Yard), CLng(Me.cmb_No), _
                                              strErrMsg, Nz(Me.cmb_Ukno, ""), Nz(Me.txt_MOTO_ACPTNO, ""))
                    End If
                End If
            End If

            If intRtn <> 0 Then
                ' システムエラー OR チェックエラー
                ' エラーメッセージ出力
                MsgBox strErrMsg, vbExclamation, P_ERROR
                GoTo ChkDispData_End
            End If
            
' ***** Start 2017/08/07 add by ysuzuki
        ' *** 受注キャンセル ***
        Case P_YUKBN_53
            ' *** 修正 ***
            If P_SHU = mdl_Mode Then
            
                ' コンテナ契約ファイルに該当データが存在するか確認
                strUkno = Me.cmb_Ukno.VALUE
                strSQL = "SELECT CARG_FILE.* FROM YOUK_TRAN, RCPT_TRAN, CARG_FILE" _
                       & " WHERE RCPT_NO = YOUKT_RCPT_NO And CARG_ACPTNO = RCPT_CARG_ACPTNO" _
                       & "   AND YOUKT_UKNO = '" & strUkno & "'"
                       
                Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)
                
                ' コンテナ契約ファイルに該当データが存在する場合
                If Not objRs.EOF Then
                    'エラーメッセージを表示し
                    'MsgBox "既に契約データ作成済みのため受付キャンセルはできません" & vbCrLf _
                         & "キャンセルするためには契約入力で未使用キャンセル登録して下さい" _
                         , vbInformation _
                         , P_KAKUNIN
            
                End If
                ChkDispData = True
                GoTo ChkDispData_End
            
            End If
' ***** end 2017/08/07 add by ysuzuki

        Case Else
    End Select

' 2006/01/30 ADD T.SUZUKI Start
    If Not pobjKONT_DB Is Nothing Then
        ' コンテナDBに接続
        If fncConnectDB = False Then Exit Function
    End If

    ChkDispData = True
    GoTo ChkDispData_End

ChkDispData_Err:
    MsgBox strMsg, vbExclamation + vbOKOnly, P_ERROR

ChkDispData_End:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing                                    ' 2005/11/02 ADD T.SUZUKI
    If Not pobjKONT_DB Is Nothing Then pobjKONT_DB.Close: Set pobjKONT_DB = Nothing                  ' 2005/11/02 ADD T.SUZUKI
    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :Chk_FrmControl
'        機能             :画面入力データの必須チェック
'        IN               :lngMode = 処理ﾓｰﾄﾞ / strKbn  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'        備考             :記述を省けますが、画面上部からチェックする事を優先しました。
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_FrmControl(ByVal lngMode As Long, ByVal strKbn As Variant, ByRef strMsg As String) As Boolean

    ' 賃料大乗チェック用変数
    Dim strPricFrom As String  ' 2005/12/20 ADD T.SUZUKI
    Dim strPricTo   As String  ' 2005/12/20 ADD T.SUZUKI
    Dim strSwap     As String  ' 2005/12/20 ADD T.SUZUKI
    Dim dblSizeFrom As Double  ' 2018/08/30 INS TAKENOUCHI
    Dim dblSizeTo   As Double  ' 2018/08/30 INS TAKENOUCHI


    Chk_FrmControl = False
    With Me
        Select Case lngMode
            Case P_NEW, P_SHU
                '受付番号入力ﾁｪｯｸ
                If lngMode = P_SHU Then
                    strMsg = "受付番号の選択をして下さい。"
                    If Nz(.cmb_Ukno, "") = "" Then .cmb_Ukno.SetFocus: Exit Function
                End If

                strMsg = "受付状態区分の選択をして下さい。"
                If Nz(strKbn, "") = "" Then .cmb_Yukbn.SetFocus: Exit Function  '受付状態区分
                '▼ 2011/07/19 add tajima ▼
                strMsg = "発生区分の選択をして下さい。"
                If Nz(.cmb_発生区分, "") = "" Then .cmb_発生区分.SetFocus: Exit Function
                '▲ 2011/07/19 add tajima ▲
                
                'INSERT 2020/01/01 Y.WADA Start
                If .cmb_Tantoc.Enabled Then
                    strMsg = "予約担当の選択をして下さい。"
                    If Nz(.cmb_Tantoc, "") = "" Then .cmb_Tantoc.SetFocus: Exit Function
                End If
                
                If .cmb_Yard.Enabled Then
                    strMsg = "ヤードコードの選択をして下さい。"
                    If Nz(.cmb_Yard, "") = "" Then .cmb_Yard.SetFocus: Exit Function
                End If
                
                If .cmb_Youto.Enabled Then
                    strMsg = "用途の選択をして下さい。"
                    If Nz(.cmb_Youto, "") = "" Then .cmb_Youto.SetFocus: Exit Function
                End If
                
                If .txt_SyzeF.Enabled Then
                    strMsg = "希望サイズを入力して下さい。"
                    If Nz(.txt_SyzeF, "") = "" Or Nz(txt_SyzeT, "") = "" Then .cmb_Kibou.SetFocus: Exit Function
                End If
                
                If .cmb_Step.Enabled Then
                    strMsg = "段区分の選択をして下さい。"
                    If Nz(.cmb_Step, "") = "" Then .cmb_Step.SetFocus: Exit Function
                End If
        
                Select Case opg_Kokya.VALUE
                    Case 1  ' 新規顧客
                        strMsg = "顧客名を入力して下さい。"
                        If Nz(.txt_Name, "") = "" Then .txt_Name.SetFocus: Exit Function
                        strMsg = "顧客名ﾌﾘｶﾞﾅを入力をして下さい。"
                        If Nz(.txt_Kana, "") = "" Then .txt_Kana.SetFocus: Exit Function

                    Case 2  ' 既存顧客
                        strMsg = "顧客名カナを選択して下さい。"
                        If Nz(.txt_Kokya, "") = "" Then .txt_Kokya.SetFocus: Exit Function
                        strMsg = "顧客コードを選択して下さい。"
                        If Nz(.cmb_KokCd, "") = "" Then .cmb_KokCd.SetFocus: Exit Function
                End Select
                
                'strMSG = "携帯番号を入力してください。"
                'If Nz(.txt_Keitai, "") = "" Then .txt_Keitai.SetFocus: Exit Function
                'INSERT 2020/01/01 Y.WADA End
                
                Select Case strKbn
                    Case P_YUKBN_00  ' *** 自動紹介 ***
'                        strMSG = "予約担当の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI  'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Tantoc, "") = "" Then .cmb_Tantoc.SetFocus: Exit Function    '受付担当                  'DELETE 2020/01/01 Y.WADA
                        strMsg = "広告媒体の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI
                        If Nz(.cmb_Ukkbn, "") = "" Then .cmb_Ukkbn.SetFocus: Exit Function      '受付区分
'                        strMSG = "ヤードコードの選択をして下さい。"                                                        'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Yard, "") = "" Then .cmb_Yard.SetFocus: Exit Function        'ヤードコード              'DELETE 2020/01/01 Y.WADA
'                        strMSG = "希望サイズを入力して下さい。"                                                            'DELETE 2020/01/01 Y.WADA
'                        If Nz(.txt_SyzeF, "") = "" Or Nz(txt_SyzeT, "") = "" Then .cmb_Kibou.SetFocus: Exit Function       'DELETE 2020/01/01 Y.WADA
                        strMsg = "本数の入力をして下さい。"
                        If Nz(.txt_Kibousu, "") = "" Then .txt_Kibousu.SetFocus: Exit Function   '本数
'DELETE 2020/01/01 Y.WADA Start
'                        Select Case opg_Kokya.VALUE
'                            Case 1  ' 新規顧客
'                                strMSG = "顧客名の入力をして下さい。"
'                                If Nz(.txt_Name, "") = "" Then .txt_Name.SetFocus: Exit Function        '顧客名
'                                strMSG = "顧客名ﾌﾘｶﾞﾅの入力をして下さい。"
'                                If Nz(.txt_Kana, "") = "" Then .txt_Kana.SetFocus: Exit Function        '顧客名ﾌﾘｶﾞﾅ
'
'                            Case 2  ' 既存顧客
'                                strMSG = "顧客名カナを選択して下さい。"
'                                If Nz(.txt_Kokya, "") = "" Then .txt_Kokya.SetFocus: Exit Function      '顧客名ｶﾅｺﾝﾎﾞﾎﾞｯｸｽ
'                                strMSG = "顧客コードを選択して下さい。"
'                                If Nz(.cmb_KokCd, "") = "" Then .cmb_KokCd.SetFocus: Exit Function      '顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ
'                        End Select
'DELETE 2020/01/01 Y.WADA End
                        strMsg = "連絡区分の選択をして下さい。"
                        If Nz(.cmb_Renkbn, "") = "" Then .cmb_Renkbn.SetFocus: Exit Function    '連絡区分

                    Case P_YUKBN_01  ' *** 自社理由 ***
'                        strMSG = "予約担当の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI  'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Tantoc, "") = "" Then .cmb_Tantoc.SetFocus: Exit Function    '受付担当                  'DELETE 2020/01/01 Y.WADA
                        strMsg = "広告媒体の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI
                        If Nz(.cmb_Ukkbn, "") = "" Then .cmb_Ukkbn.SetFocus: Exit Function      '受付区分
'                        strMSG = "ヤードコードの選択をして下さい。"                                                        'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Yard, "") = "" Then .cmb_Yard.SetFocus: Exit Function        'ヤードコード              'DELETE 2020/01/01 Y.WADA
                        strMsg = "ｺﾝﾃﾅ番号の選択をして下さい。"
                        If Nz(.cmb_No, "") = "" Then .cmb_No.SetFocus: Exit Function            'ｺﾝﾃﾅ番号
                        strMsg = "受付日の入力をして下さい。"                                   ' 2005/12/19 ADD T.SUZUKI
                        If Nz(.txt_Kudate, "") = "" Then .txt_Kudate.SetFocus: Exit Function    '契約受付日
                        strMsg = "取置き期限日の入力をして下さい。"
                        If Nz(.txt_Tkdate, "") = "" Then .txt_Tkdate.SetFocus: Exit Function    '取置き期限日
'DELETE 2020/01/01 Y.WADA Start
'                        Select Case opg_Kokya.VALUE
'                            Case 1  ' 新規顧客
'                                strMSG = "顧客名の入力をして下さい。"
'                                If Nz(.txt_Name, "") = "" Then .txt_Name.SetFocus: Exit Function        '顧客名
'                                strMSG = "顧客名ﾌﾘｶﾞﾅの入力をして下さい。"
'                                If Nz(.txt_Kana, "") = "" Then .txt_Kana.SetFocus: Exit Function        '顧客名ﾌﾘｶﾞﾅ
'
'                            Case 2  ' 既存顧客
'                                strMSG = "顧客名カナを選択して下さい。"
'                                If Nz(.txt_Kokya, "") = "" Then .txt_Kokya.SetFocus: Exit Function      '顧客名ｶﾅｺﾝﾎﾞﾎﾞｯｸｽ
'                                strMSG = "顧客コードを選択して下さい。"
'                                If Nz(.cmb_KokCd, "") = "" Then .cmb_KokCd.SetFocus: Exit Function      '顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ
'                        End Select
'DELETE 2020/01/01 Y.WADA End
                        strMsg = "連絡区分の選択をして下さい。"
                        If Nz(.cmb_Renkbn, "") = "" Then .cmb_Renkbn.SetFocus: Exit Function    '連絡区分

                    Case P_YUKBN_02  ' *** 取置 ***
'                        strMSG = "予約担当の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI  'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Tantoc, "") = "" Then .cmb_Tantoc.SetFocus: Exit Function    '受付担当                  'DELETE 2020/01/01 Y.WADA
                        strMsg = "広告媒体の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI
                        If Nz(.cmb_Ukkbn, "") = "" Then .cmb_Ukkbn.SetFocus: Exit Function      '受付区分
'                        strMSG = "ヤードコードの選択をして下さい。"                                                        'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Yard, "") = "" Then .cmb_Yard.SetFocus: Exit Function        'ヤードコード              'DELETE 2020/01/01 Y.WADA
                        If pintAUTOKBN = 0 Then
                            ' 自動区分=0のとき、ｺﾝﾃﾅ番号は必須入力
                            strMsg = "ｺﾝﾃﾅ番号の選択をして下さい。"
                            If Nz(.cmb_No, "") = "" Then .cmb_No.SetFocus: Exit Function        'ｺﾝﾃﾅ番号
                        End If
                        strMsg = "受付日の入力をして下さい。"                                   ' 2005/12/19 ADD T.SUZUKI
                        If Nz(.txt_Kudate, "") = "" Then .txt_Kudate.SetFocus: Exit Function    '契約受付日
                        strMsg = "取置き期限日の入力をして下さい。"
                        If Nz(.txt_Tkdate, "") = "" Then .txt_Tkdate.SetFocus: Exit Function    '取置き期限日
'DELETE 2020/01/01 Y.WADA Start
'                        Select Case opg_Kokya.VALUE
'                            Case 1  ' 新規顧客
'                                strMSG = "顧客名の入力をして下さい。"
'                                If Nz(.txt_Name, "") = "" Then .txt_Name.SetFocus: Exit Function        '顧客名
'                                strMSG = "顧客名ﾌﾘｶﾞﾅの入力をして下さい。"
'                                If Nz(.txt_Kana, "") = "" Then .txt_Kana.SetFocus: Exit Function        '顧客名ﾌﾘｶﾞﾅ
'
'                            Case 2  ' 既存顧客
'                                strMSG = "顧客名カナを選択して下さい。"
'                                If Nz(.txt_Kokya, "") = "" Then .txt_Kokya.SetFocus: Exit Function      '顧客名ｶﾅｺﾝﾎﾞﾎﾞｯｸｽ
'                                strMSG = "顧客コードを選択して下さい。"
'                                If Nz(.cmb_KokCd, "") = "" Then .cmb_KokCd.SetFocus: Exit Function      '顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ
'                        End Select
'DELETE 2020/01/01 Y.WADA End
                        strMsg = "連絡区分の選択をして下さい。"
                        If Nz(.cmb_Renkbn, "") = "" Then .cmb_Renkbn.SetFocus: Exit Function    '連絡区分

                    Case P_YUKBN_10, P_YUKBN_20  ' *** 契約受付 OR 契約 ***
'                        strMSG = "予約担当の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI  'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Tantoc, "") = "" Then .cmb_Tantoc.SetFocus: Exit Function     '受付担当                 'DELETE 2020/01/01 Y.WADA
                        strMsg = "広告媒体の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI
                        If Nz(.cmb_Ukkbn, "") = "" Then .cmb_Ukkbn.SetFocus: Exit Function       '受付区分

                        If strKbn = P_YUKBN_10 Then                                              ' 2005/12/12 ADD T.SUZUKI
'                            strMSG = "ヤードコードの選択をして下さい。"                                                    'DELETE 2020/01/01 Y.WADA
'                            If Nz(.cmb_Yard, "") = "" Then .cmb_Yard.SetFocus: Exit Function     'ヤードコード             'DELETE 2020/01/01 Y.WADA
                            strMsg = "ｺﾝﾃﾅ番号の選択をして下さい。"
                            If Nz(.cmb_No, "") = "" Then .cmb_No.SetFocus: Exit Function         'ｺﾝﾃﾅ番号
                            strMsg = "受付日の入力をして下さい。"                                ' 2005/12/19 ADD T.SUZUKI
                            If Nz(.txt_Kudate, "") = "" Then .txt_Kudate.SetFocus: Exit Function '契約受付日

                            strMsg = "取置き期限日の入力をして下さい。"
                            If Nz(.txt_Tkdate, "") = "" Then .txt_Tkdate.SetFocus: Exit Function '取置き期限日
                            '▼ 2006/09/26 add tajima ▼
                            '新規の受付で移動予約でないときのみ希望開始日をチェックする※過去データを考慮
                            If lngMode = P_NEW And .cmb_MoveKbn = P_MOVEKBN_00 Then
                                strMsg = "希望開始日の入力をして下さい。"
                                If Nz(.txt_Kkdate, "") = "" Then .txt_Kkdate.SetFocus: Exit Function
                                ' 希望開始日が営業期間内かのチェック
                                If str営業開始日 > Format$(.txt_Kkdate, "YYYY/MM/DD") Or _
                                   str営業終了日 < Format$(.txt_Kkdate, "YYYY/MM/DD") Then
                                    strMsg = "希望開始日は営業期間内としてください。"
                                    .txt_Kkdate.SetFocus: Exit Function
                                End If
                            End If
                            '▲ 2006/09/26 add tajima ▲
                        End If                                                                   ' 2005/12/12 ADD T.SUZUKI
'DELETE 2020/01/01 Y.WADA Start
'                        Select Case opg_Kokya.VALUE
'                            Case 1  ' 新規顧客
'                                strMSG = "顧客名の入力をして下さい。"
'                                If Nz(.txt_Name, "") = "" Then .txt_Name.SetFocus: Exit Function        '顧客名
'                                strMSG = "顧客名ﾌﾘｶﾞﾅの入力をして下さい。"
'                                If Nz(.txt_Kana, "") = "" Then .txt_Kana.SetFocus: Exit Function        '顧客名ﾌﾘｶﾞﾅ
'
'                            Case 2  ' 既存顧客
'                                strMSG = "顧客名カナを選択して下さい。"
'                                If Nz(.txt_Kokya, "") = "" Then .txt_Kokya.SetFocus: Exit Function      '顧客名ｶﾅｺﾝﾎﾞﾎﾞｯｸｽ
'                                strMSG = "顧客コードを選択して下さい。"
'                                If Nz(.cmb_KokCd, "") = "" Then .cmb_KokCd.SetFocus: Exit Function      '顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ
'                        End Select
'DELETE 2020/01/01 Y.WADA End
                        strMsg = "連絡区分の選択をして下さい。"
                        If Nz(.cmb_Renkbn, "") = "" Then .cmb_Renkbn.SetFocus: Exit Function     '連絡区分

                    Case P_YUKBN_51, P_YUKBN_52, P_YUKBN_53  ' *** 予約キャンセル OR 予約取置キャンセル OR 受付キャンセル ***   ' 2005/12/12 ADD T.SUZUKI
'                        strMSG = "予約担当の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI      'DELETE 2020/01/01 Y.WADA
'                        If Nz(.cmb_Tantoc, "") = "" Then .cmb_Tantoc.SetFocus: Exit Function    '受付担当                      'DELETE 2020/01/01 Y.WADA
                        strMsg = "広告媒体の選択をして下さい。"                                 ' 2005/12/19 ADD T.SUZUKI
                        If Nz(.cmb_Ukkbn, "") = "" Then .cmb_Ukkbn.SetFocus: Exit Function      '受付区分
                        strMsg = "予約取消日の入力をして下さい。"
                        If Nz(.txt_Ytdate, "") = "" Then .txt_Ytdate.SetFocus: Exit Function    '予約取消日
'DELETE 2020/01/01 Y.WADA Start
'                        Select Case opg_Kokya.VALUE
'                            Case 1  ' 新規顧客
'                                strMSG = "顧客名の入力をして下さい。"
'                                If Nz(.txt_Name, "") = "" Then .txt_Name.SetFocus: Exit Function        '顧客名
'                                strMSG = "顧客名ﾌﾘｶﾞﾅの入力をして下さい。"
'                                If Nz(.txt_Kana, "") = "" Then .txt_Kana.SetFocus: Exit Function        '顧客名ﾌﾘｶﾞﾅ
'
'                            Case 2  ' 既存顧客
'                                strMSG = "顧客名カナを選択して下さい。"
'                                If Nz(.txt_Kokya, "") = "" Then .txt_Kokya.SetFocus: Exit Function      '顧客名ｶﾅｺﾝﾎﾞﾎﾞｯｸｽ
'                                strMSG = "顧客コードを選択して下さい。"
'                                If Nz(.cmb_KokCd, "") = "" Then .cmb_KokCd.SetFocus: Exit Function      '顧客ｺｰﾄﾞｺﾝﾎﾞﾎﾞｯｸｽ
'                        End Select
'DELETE 2020/01/01 Y.WADA End
                        strMsg = "連絡区分の選択をして下さい。"
                        If Nz(.cmb_Renkbn, "") = "" Then .cmb_Renkbn.SetFocus: Exit Function    '連絡区分
                End Select

                '連絡区分で選択したﾃｷｽﾄ入力がされているかチェック
                If Nz(.cmb_Renkbn, "") <> "" Then
                    Select Case Format(.cmb_Renkbn, "00")
                        Case P_REN_00
                            strMsg = "連絡区分が電話の時は電話番号を入力してください。"
                            '' 2013/06/03 M.HONDA START
                            'If Nz(.txt_Tel, "") = "" Then .txt_Tel.SetFocus: Exit Function
                            If (Nz(.txt_Tel, "") = "" And Nz(.txt_Keitai, "") = "") Then .txt_Tel.SetFocus: Exit Function
                            '' 2013/06/03 M.HONDA END
'DELETE 2020/01/01 Y.WADA Start
'                        Case P_REN_01
'                            strMSG = "連絡区分が携帯の時は携帯番号を入力してください。"
'                            If Nz(.txt_Keitai, "") = "" Then .txt_Keitai.SetFocus: Exit Function
'DELETE 2020/01/01 Y.WADA End
                        Case P_REN_02
                            strMsg = "連絡区分がFAXの時はFAX番号を入力してください。"
                            If Nz(.txt_Fax, "") = "" Then .txt_Fax.SetFocus: Exit Function
                        Case P_REN_10
                            strMsg = "連絡区分がEﾒｰﾙの時はメールアドレスを入力してください。"
                            If Nz(.txt_Mail, "") = "" Then .txt_Mail.SetFocus: Exit Function
                        Case Else
                    End Select
                End If
                
                '▼ 2007/08/09 add tajima ▼
                ' 申込書送付方法にて 04:e-mail を選択したときの必須項目チェック
                If Nz(cmb_申込書送付方.VALUE, "") = "04" Then
                    strMsg = "申込書送付方がEﾒｰﾙの時はメールアドレスを入力してください。"
                    If Nz(.txt_Mail.VALUE, "") = "" Then .txt_Mail.SetFocus: Exit Function
                End If
                '▲ 2007/08/09 add tajima ▲

                '顧客/代表者カナ項目のみ入力文字数のチェック
                strMsg = "カナ項目は半角18文字以内で入力してください。"
                If Len(Me.txt_Kana) > 18 Then
                    If Nz(.txt_Kana, "") <> "" Then .txt_Kana.SetFocus: Exit Function
                End If
                If Len(Me.txt_Taka) > 18 Then
                    If Nz(.txt_Taka, "") <> "" Then .txt_Taka.SetFocus: Exit Function
                End If

                '▼ 2018/08/30 TAKENOUCHI INS<S> ▼
                ' 希望サイズのFrom～Toチェック(From値がTo値より大きい場合エラーとする)
                If (Nz(.txt_SyzeF.VALUE, "") <> "" And Nz(.txt_SyzeT.VALUE, "") <> "") Then
                    dblSizeFrom = CDbl(.txt_SyzeF.VALUE)
                    dblSizeTo = CDbl(.txt_SyzeT.VALUE)
                    
                    If dblSizeFrom > dblSizeTo Then
                        strMsg = "希望サイズFrom/Toの入力が矛盾しています。"
                        .txt_SyzeF.SetFocus
                        Exit Function
                    End If
                End If
                '▲ 2018/08/30 TAKENOUCHI INS<E> ▲

            Case Else
                If lngMode = P_DEL Then
                    strMsg = "受付番号"
                    If Nz(.cmb_Ukno, "") = "" Then .cmb_Ukno.SetFocus: Exit Function
                End If
        End Select

        '2006/05/12 ADD Tajima Start
        strMsg = "都道府県市町村は全角18文字以内で入力して下さい。"
'        If CmLenb(Me.txt_Adr1) > 36 Then                                       'DELETE 2019/08/19 N.IMAI
        If CmLenb(Nz(Me.txt_Adr1)) > 36 Then                                    'INSERT 2019/08/19 N.IMAI
            .txt_Adr1.SetFocus: Exit Function
            Exit Function
        End If
        '2006/05/12 ADD Tajima End

        '2006/04/15 ADD Tajima Start
        strMsg = "顧客区分の選択をして下さい。"
        If Nz(.cmb_Kkbn, "") = "" Then .cmb_Kkbn.SetFocus: Exit Function    '顧客区分
        '2006/04/15 ADD Tajima End

        '連絡時間テキスト入力矛盾チェック
        If Nz(.txt_RenFrom, "") <> "" And Nz(.txt_RenTo, "") <> "" Then
            If .txt_RenFrom > .txt_RenTo Then
                strMsg = "連絡時間From/Toの入力が矛盾しています。"
                .txt_RenFrom.SetFocus: Exit Function
            End If
        End If



' 2005/12/20 ADD T.SUZUKI Start
        ' 賃料のFrom～Toチェック(逆転していた場合、FromとToを入れ替える)
        If (Me.txt_PricFrom <> "" And Me.txt_PricTo <> "") Then
            strPricFrom = Me.txt_PricFrom  ' 賃料(From)
            strPricTo = Me.txt_PricTo      ' 賃料(To)

            'INSERT 2018/04/16 N.IMAI Start
            If IsNumeric(strPricFrom) = False Then
                strPricFrom = "0"
            End If
            If IsNumeric(strPricTo) = False Then
                strPricTo = "0"
            End If
            'INSERT 2018/04/16 N.IMAI End
'            If strPricTo <> "" And strPricFrom > strPricTo Then                'DELETE 2018/04/16 N.IMAI
            If strPricTo <> "" And CLng(strPricFrom) > CLng(strPricTo) Then     'INSERT 2018/04/16 N.IMAI
                ' 大小入力エラー時
                strSwap = strPricFrom
                strPricFrom = strPricTo
                strPricTo = strSwap

                ' 参照元へ再セット
                Me.txt_PricFrom = strPricFrom
                Me.txt_PricTo = strPricTo
            End If
        End If
' 2005/12/20 ADD T.SUZUKI End
    End With

    Chk_FrmControl = True
End Function

'==============================================================================*
'
'        MODULE_NAME      :数値項目のチェックを一括でする。☆★他フォームでも個々に宣言されている
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
Private Function CmbCheck_Num(ByVal oText As Variant, ByVal ilen As Integer, Optional ByVal isLEN As Variant) As Integer

    Dim vINT    As Variant
    Dim vsINT   As Variant
    Dim i       As Integer

    Dim dTEXT   As Double
    Dim dFROM   As Double
    Dim dTO     As Double
    Dim vRet    As Variant

    On Error GoTo err_rtn

    CmbCheck_Num = 9

    If Nz(CmNumeric(oText), "") = "" Then       '数値と判断できないのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    If CmChkCher(CStr(Nz(oText)), "A") = False Then '全角文字が入力されているのでエラー
        CmbCheck_Num = 1
        GoTo Exit_rtn
    End If

    '小数点の値が設定されている時のチェック
    vRet = Nz(InStr(Abs(oText), "."), 0)
    
    If IsMissing(isLEN) = False Then
        If vRet <> 0 Then
            If isLEN < CmLenb(Abs(oText)) - vRet Then    '小数点桁数が設定桁数を超えているエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
            End If
        End If
    Else
        If CmLenb(Abs(CmNumeric(oText))) > ilen Then           '文字数超過のためエラー
              CmbCheck_Num = 2
              GoTo Exit_rtn
        End If
    
        If vRet <> 0 Then                         '小数点の入力があるのでエラー
          CmbCheck_Num = 2
          GoTo Exit_rtn
        End If
    End If

    '桁チェック
    i = 0
    For i = 1 To ilen: vINT = vINT & "9": Next

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

    If dFROM > dTEXT Or dTO < dTEXT Then        '論理範囲超過のためエラー
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
'        MODULE_NAME      :Textエリア入力文字数ﾁｪｯｸ
'        MODULE_ID        :ChkTextLength
'        IN               :ctlTxt = 処理対象Text lngLength = 制限文字数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ChkTextLength(ByRef ctlTxt As TextBox, ByVal lngLength As Long) As Boolean

    Dim strwork As String

    On Error GoTo ChkTextLength_Err

    ChkTextLength = False
    strwork = Nz(ctlTxt, "")

    Select Case CmBCheck_Char(strwork, lngLength)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo ChkTextLength_End
        Case 9
            MsgBox "内容を正しく入力してください。", vbInformation, P_ERROR
            GoTo ChkTextLength_End
    End Select

    ChkTextLength = True
    GoTo ChkTextLength_End

ChkTextLength_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
ChkTextLength_End:
End Function
'==============================================================================*
'
'        MODULE_NAME      :Comboエリア入力文字数ﾁｪｯｸ
'        MODULE_ID        :ChkCmbLength
'        IN               :ctlCmb = 処理対象Text lngLength = 制限文字数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function ChkCmbLength(ByRef ctlCmb As ComboBox, ByVal lngLength As Long) As Boolean

    Dim strwork As String

    On Error GoTo ChkCmbLength_Err

    ChkCmbLength = False
    strwork = Nz(ctlCmb, "")

    Select Case CmBCheck_Char(strwork, lngLength)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo ChkCmbLength_End
        Case 9
            MsgBox "内容を正しく入力してください。", vbInformation, P_ERROR
            GoTo ChkCmbLength_End
    End Select

    ChkCmbLength = True
    GoTo ChkCmbLength_End

ChkCmbLength_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

ChkCmbLength_End:
End Function
'==============================================================================*
'
'        MODULE_NAME      :文字列のﾁｪｯｸ
'        MODULE_ID        :CmBCheck_Char
'        IN               :oText = 処理対象Text iLen = 制限文字数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CmBCheck_Char(ByVal oText As Variant, ByVal ilen As Integer) As Integer

'-----------------'
'  内容のチェック
'-----------------'
  '返り値は、
  '２：桁数が溢れている時
  '９：それ以外のエラー
  '０：すべてのチェックがＯＫの時

    On Error GoTo Err_CmBCheck_Char

    CmBCheck_Char = 9

    If CmLenb(oText) > ilen Then
        CmBCheck_Char = 2
        GoTo Exit_CmBCheck_Char
    End If

    CmBCheck_Char = 0

Exit_CmBCheck_Char:
    Exit Function

Err_CmBCheck_Char:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    CmBCheck_Char = 9
    GoTo Exit_CmBCheck_Char
End Function
'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :CMCheck_Date
'        CREATE_DATE      :
'
'Check_Dateでは'198/03といった形が変換されるための補助検査
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function CMCheck_Date(ByVal vDATE As Variant)
  
    Dim sYear As String
    
    CMCheck_Date = Null
    
    If Nz(vDATE, "") = "" Then Exit Function
    If IsDate(vDATE) = False Then Exit Function
    
    sYear = Format(vDATE, "yyyy")
    If CmLenb(sYear) <> 4 Then Exit Function
    
    CMCheck_Date = vDATE
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :日付項目のチェックを一括でする
'        MODULE_ID        :
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

    On Error GoTo CmBCheck_Date_Err

    CmBCheck_Date = 2

    If IsMissing(ilen) = False Then
       intLen = Nz(CmNumeric(ilen), 0)
    Else
       intLen = 10
    End If

    oText = Nz(oText, "")
    '桁数チェック
    If CmLenb(oText) > intLen Then GoTo CmBCheck_Date_End
    CmBCheck_Date = 3
    If Nz(oText, "") = "" Then GoTo CmBCheck_Date_End

    '論理チェック
    If IsDate(oText) = False Then GoTo CmBCheck_Date_End

    sYear = Format(oText, "yyyy")
    If CmLenb(sYear) <> 4 Then GoTo CmBCheck_Date_End

    CmBCheck_Date = 0

    GoTo CmBCheck_Date_End

CmBCheck_Date_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    CmBCheck_Date = 9

CmBCheck_Date_End:
    Exit Function
End Function
'==============================================================================*
'
'        MODULE_NAME      :DateCheck
'        機能             :日付入力ﾃﾞｰﾀのチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ / Cancel  = 状態区分
'        OUT              :strMsg = ｴﾗｰﾒｯｾｰｼﾞ / Chk_FrmControl = 入力不備有無
'        NOTE  2006/09/26 :結果を戻り値で返すようにした
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function dateCheck(txtTarget As Variant, ByRef Cancel As Integer)

    dateCheck = False
    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    Else
        dateCheck = True
    End If

End Function
'==============================================================================*
'
'        MODULE_NAME      :Txt_EnabledChk
'        機能             :TextEnabledチェック⇒Nullチェック
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ
'        OUT              :セットするデータ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Txt_EnabledChk(ByRef ctlTxt As TextBox) As Variant

    Dim varData As Variant

' 2005/12/19 ADD T.SUZUKI Start
    ' 使用不可の項目であっても、該当項目に設定されている値を
    ' 戻すように修正
    varData = Txt_ChkNull(ctlTxt, 1)
' 2005/12/19 ADD T.SUZUKI End
' 2005/12/19 DEL T.SUZUKI Start
'    If ctlTxt.Enabled = True Then
'        varData = Txt_ChkNull(ctlTxt, 1)
'    Else
'        varData = Null
'    End If
' 2005/12/19 DEL T.SUZUKI End

    Txt_EnabledChk = varData
End Function
'==============================================================================*
'
'        MODULE_NAME      :Cmb_ChkNull
'        機能             :ｺﾝﾎﾞﾎﾞｯｸｽチェック
'        IN               :ctlCmb = 処理対象ｺﾝﾎﾞ
'        OUT              :セットするデータ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Cmb_ChkNull(ByRef ctlCmb As ComboBox) As Variant

    Dim varData As Variant
    
    If Trim(ctlCmb) = "" Then
        varData = Null
    Else
        varData = Nz(ctlCmb, Null)
    End If
    Cmb_ChkNull = varData
    
End Function
'==============================================================================*
'
'        MODULE_NAME      :Txt_ChkNull
'        機能             :ﾃｷｽﾄﾁｪｯｸ
'        IN               :ctlTxt = 処理対象ﾃｷｽﾄ /lngMode = 0:通常ﾃﾞｰﾀ 1:Date
'        OUT              :セットするデータ
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Txt_ChkNull(ByRef ctlTxt As TextBox, Optional lngMode As Long = 0) As Variant
    
    Dim varData As Variant
    
    'If Trim(ctlTxt) = "" Then                                                  'DELETE 20050721 K.ISHIZAKA
    If Trim(Nz(ctlTxt, "")) = "" Then                                           'INSERT 20050721 K.ISHIZAKA
        varData = Null
    Else
        Select Case lngMode
            Case 0: varData = Nz(ctlTxt, Null)
            Case 1
                varData = Nz(Format(ctlTxt, "yyyy/mm/dd"), Null)
            Case Else
        End Select
    End If

    Txt_ChkNull = varData
End Function
'==============================================================================*
'
'        MODULE_NAME      :BeforForm_Reload
'        機能             :前画面の再表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub BeforForm_Reload()                                                  'INSERT START   20050717 K.ISHIZAKA

    Dim beforForm As Form

    On Error Resume Next

    Set beforForm = Forms("FKS020_コンテナ予約入力")
    If Not beforForm Is Nothing Then
        Call beforForm.Reload(strBeforYard, Me.cmb_Yard)
    End If
'    Set beforForm = Forms("FVS020")                                            'DELETE 20050730 K.ISHIZAKA
'    If Not beforForm Is Nothing Then                                           'DELETE 20050730 K.ISHIZAKA
'        Call beforForm.Reload(strBeforYard, Me.cmb_Yard)                       'DELETE 20050730 K.ISHIZAKA
'    End If                                                                     'DELETE 20050730 K.ISHIZAKA
'    Set beforForm = Forms("FVS820")                                            'DELETE 20050730 K.ISHIZAKA
    Set beforForm = Forms("FVS021")                                             'INSERT 20050730 K.ISHIZAKA
    If Not beforForm Is Nothing Then
        Call beforForm.Reload(strBeforYard, Me.cmb_Yard)
    Else                                                                        'INSERT 20050730 K.ISHIZAKA
        Set beforForm = Forms("FVS011")                                         'INSERT 20050730 K.ISHIZAKA
        If Not beforForm Is Nothing Then                                        'INSERT 20050730 K.ISHIZAKA
            Call beforForm.Reload(strBeforYard, Me.cmb_Yard)                    'INSERT 20050730 K.ISHIZAKA
        End If                                                                  'INSERT 20050730 K.ISHIZAKA
    End If
                                                                                'INSERT START 2011/02/22 K.ISHIZAKA
    Set beforForm = Forms("FTG011")
    If Not beforForm Is Nothing Then
        Call beforForm.Reload(strBeforYard, str空検索時レンタル用途, cmb_Yard.VALUE, Nz(cmb_Youto.VALUE, ""))
    End If
    Set beforForm = Forms("FTG021")
    If Not beforForm Is Nothing Then
        Call beforForm.Reload(strBeforYard, str空検索時レンタル用途, cmb_Yard.VALUE, Nz(cmb_Youto.VALUE, ""))
    End If
                                                                                'INSERT END   2011/02/22 K.ISHIZAKA

' 2005/10/06 ADD T.SUZUKI Start
    ' 予約一覧画面
    If pstrFormID <> "" Then
        Set beforForm = Forms(pcstrFVS540)
        If Not beforForm Is Nothing Then
            Call beforForm.subReload(IIf(pblnUpdateFlg = True, 1, 0))
        End If
    End If
' 2005/10/06 ADD T.SUZUKI End
End Sub                                                                         'INSERT END     20050717 K.ISHIZAKA
'==============================================================================*
'
'        MODULE_NAME      :DB接続
'        MODULE_ID        :fncConnectDB
'        CREATE_DATE      :2005/10/06
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectDB() As Boolean

    Dim strBumoCD     As String

    Dim strConnect    As String
    Dim strDataSource As String

    On Error GoTo fncConnectDB_Err

    fncConnectDB = False

    ' 部門コード取得
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    '会社コード取得処理
    strCampCd = DLookup("CONT_CAMPC", "dbo_CONT_MAST")               ' 2018/09/25 ADD WATANABE

    ' -----------------------------------------------------------------------------------------------------------
    ' コンテナDB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumoCD & "'"), "")
    
    If strDataSource = "" Then
        MsgBox "SETU_TABLの設定が不正です。", vbExclamation, P_PROG_ID
        Set pobjKONT_DB = Nothing
    Else
        Set pobjKONT_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00(strBumoCD))
    End If
    ' -----------------------------------------------------------------------------------------------------------

    fncConnectDB = True

    Exit Function

fncConnectDB_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Function
'==============================================================================*
'
'        MODULE_NAME      :DB切断
'        MODULE_ID        :fncDisConnectDB
'        CREATE_DATE      :2005/10/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' 2006/04/07 REP T.SUZUKI Stsart
Private Function fncDisConnectDB(Optional ByVal intKbn As Integer = 0) As Boolean
'Private Function fncDisConnectDB() As Boolean
' 2006/04/07 REP T.SUZUKI End

    Dim strMsg As String

    On Error GoTo fncDisConnectDB_Err

    fncDisConnectDB = False

    ' コンテナDB
    pobjKONT_DB.Close
    Set pobjKONT_DB = Nothing

    fncDisConnectDB = True

    Exit Function

fncDisConnectDB_Err:
    strMsg = "fncDisConnectDB(ｴﾗｰｺｰﾄﾞ：" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ：" & Err.Description & ")"   ' 2006/04/07 ADD T.SUZUKI

' 2006/04/07 ADD T.SUZUKI Start
    If intKbn = 0 Then
        Err.Clear
        MsgBox strMsg        ' エラーメッセージ出力
    Else
        Call Err.Raise(Err.Number, "fncDisConnectDB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
' 2006/04/07 ADD T.SUZUKI End
End Function
'==============================================================================*
'
'        MODULE_NAME      :予約受付状態区分コンボボックスのSQL作成
'        MODULE_ID        :subCreate_YukbnSQL
'        Parameter        :lngMode = 処理ﾓｰﾄﾞ
'                          strKbn  = 状態区分
'        CREATE_DATE      :2005/12/08
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCreate_YukbnSQL(ByVal lngMode As Long, ByRef strKbn As String, _
                               Optional ByVal strYukbn As String)
'Private Sub subCreate_YukbnSQL(ByVal lngMode As Long, ByRef strKbn As String)

    Dim strSQL As String

    ' メインSQL文作成
    strSQL = "SELECT Format([NAME_CODE],'00') AS 式1, "
    strSQL = strSQL & " dbo_NAME_MAST.NAME_NAME "
    strSQL = strSQL & "  FROM dbo_NAME_MAST "
    strSQL = strSQL & " WHERE dbo_NAME_MAST.NAME_ID = '" & P_NAME_080 & "' "

    If lngMode = P_NEW Then
        ' *** 新規 ***
        If strYukbn <> "" Then
            strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
            strSQL = strSQL & strYukbn & ") "      ' 第3引数で指定された状態区分
        Else
            strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
            strSQL = strSQL & P_YUKBN_00 & ", "    ' 自動紹介
            strSQL = strSQL & P_YUKBN_01 & ", "    ' 自社理由
            strSQL = strSQL & P_YUKBN_02 & ", "    ' 取置
            strSQL = strSQL & P_YUKBN_10 & ") "    ' 受付
        End If

    ElseIf lngMode = P_SHU Then
        ' *** 修正 ***
        If strKbn <> "" Then                                            ' 2005/12/21 ADD T.SUZUKI
            Select Case strKbn
                Case P_YUKBN_00
                    ' 自動紹介
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_00 & ", "    ' 自動紹介
                    strSQL = strSQL & P_YUKBN_02 & ", "    ' 取置
                    strSQL = strSQL & P_YUKBN_10 & ", "    ' 受付
                    strSQL = strSQL & P_YUKBN_51 & ") "    ' 予約キャンセル
    
                Case P_YUKBN_01
                    ' 自社理由
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_01 & ", "    ' 自社理由
                    strSQL = strSQL & P_YUKBN_10 & ", "    ' 受付
                    strSQL = strSQL & P_YUKBN_51 & ") "    ' 予約キャンセル
    
                Case P_YUKBN_02
                    ' 取置
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_00 & ", "    ' 自動紹介
                    strSQL = strSQL & P_YUKBN_02 & ", "    ' 取置
                    strSQL = strSQL & P_YUKBN_10 & ", "    ' 受付
                    strSQL = strSQL & P_YUKBN_51 & ") "    ' 予約キャンセル
    
                Case P_YUKBN_10
                    ' 受付
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_10 & ", "    ' 受付
                    strSQL = strSQL & P_YUKBN_53 & ") "    ' 受付キャンセル
    
                Case P_YUKBN_20
                    ' 契約
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_20 & ") "    ' 契約
    
                Case P_YUKBN_51
                    ' 予約キャンセル
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_51 & ") "    ' 予約キャンセル
    
                Case P_YUKBN_52
                    ' 取置キャンセル
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_00 & ", "    ' 自動紹介
                    strSQL = strSQL & P_YUKBN_51 & ", "    ' 予約キャンセル
                    strSQL = strSQL & P_YUKBN_52 & ") "    ' 取置キャンセル
    
                Case P_YUKBN_53
                    ' 受付キャンセル
                    strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
                    strSQL = strSQL & P_YUKBN_53 & ") "    ' 受付キャンセル
                Case Else
            End Select
        End If                                                          ' 2005/12/21 ADD T.SUZUKI

    ElseIf lngMode = P_DEL Then
        ' *** 削除 ***
        If strKbn <> "" Then                                            ' 2005/12/21 ADD T.SUZUKI
            strSQL = strSQL & "   AND  dbo_NAME_MAST.NAME_CODE IN ("
            strSQL = strSQL & strKbn & ") "
        End If                                                          ' 2005/12/21 ADD T.SUZUKI
    End If

    ' ソートキー指定
    strSQL = strSQL & " ORDER BY dbo_NAME_MAST.NAME_ID;"

    ' 予約受付状態区分コンボボックスのSQL設定
    Me.cmb_Yukbn.RowSource = strSQL
End Sub
'==============================================================================*
'
'        MODULE_NAME      :取置期限日の設定 ※受付の場合は希望開始日も・・・
'        MODULE_ID        :subSetTKDATE
'        CREATE_DATE      :2005/12/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetTKDATE()

    Dim strSQL           As String
    Dim objRs            As Recordset
    Dim strTargetDate    As String
    Dim intTORI_DD       As Integer    ' 取置期限日数
    Dim intKEIUKE_TLD    As Integer    ' 契約受付期限日数
    Dim intJISYARIYU_TLM As Integer    ' 自社理由取置期限月数
    Dim strLastDate      As String     ' ヤード最終使用日付格納変数
    Dim strErrMsg        As String     ' エラーメッセージ格納変数
    Dim intサービス期間  As Integer    ' 2006/09/27 add tajima
    Dim intAddMon        As Integer    ' 2020/08/27 INS Takenouchi 起算日導出時の加算月
    Dim intAddDay        As Integer    ' 2020/08/27 INS Takenouchi 起算日導出時の加算日
    Dim strKkdateTmp     As String     ' 2020/08/27 INS Takenouchi
 
    On Error GoTo subSetTKDATE_Err

    ' ヤードコードが未選択の場合は処理を抜ける
    If Nz(Me.cmb_Yard, "") = "" Then Exit Sub

    If Nz(Me.txt_Kudate, "") = "" Then Exit Sub

    ' コンテナDB接続
    If fncConnectDB = False Then GoTo subSetTKDATE_Exit

    ' コントロールマスタより項目取得
    ' ①取置期限日数
    ' ②契約受付期限日数
    ' ③自社理由取置期限月数
    ' ④サービス期間の取得 2006/09/26 add tajima
    strSQL = "  SELECT  CONT_TORI_DD, " & Chr(13)
    strSQL = strSQL & " CONT_KEIUKE_TLD, " & Chr(13)
    strSQL = strSQL & " CONT_JISYARIYU_TLM " & Chr(13)
    strSQL = strSQL & ",CONT_SERVICE_DD " & Chr(13)
    strSQL = strSQL & " FROM CONT_MAST "
    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)

    intTORI_DD = 0
    intKEIUKE_TLD = 0
    intJISYARIYU_TLM = 0
    intサービス期間 = 0                                               ' 2006/09/27 add tajima
    Do Until objRs.EOF
        intTORI_DD = Nz(objRs.Fields("CONT_TORI_DD"), 0)              ' 取置期限日数
        intKEIUKE_TLD = Nz(objRs.Fields("CONT_KEIUKE_TLD"), 0)        ' 契約受付期限日数
        intJISYARIYU_TLM = Nz(objRs.Fields("CONT_JISYARIYU_TLM"), 0)  ' 自社理由取置期限月数
        intサービス期間 = Nz(objRs.Fields("CONT_SERVICE_DD"), 0)      ' 2006/09/27 add tajima
        objRs.MoveNext
    Loop
    objRs.Close
    Set objRs = Nothing

    '2020/08/27 INS Takenouchi Start
    '汎用マスタ「INTI_FILE」から起算日導出のための日数を取得する
    Call getNissuForKisanDate(Me.txt_Kudate, intAddMon, intAddDay)
    If intAddMon > 0 And intAddDay > 0 Then
        '加算月が1以上の場合、起算日を算出する際の基準日は翌月の指定日とする
        strKkdateTmp = DateAdd("m", intAddMon, Me.txt_Kudate)
        strKkdateTmp = DateSerial(Year(strKkdateTmp), Month(strKkdateTmp), intAddDay)
    ElseIf intAddDay > 0 Then
        '加算月が0の場合、指定日に加算とする
        strKkdateTmp = DateAdd("d", intAddDay, Me.txt_Kudate)
    End If
    '2020/08/27 INS Takenouchi End
    
    ' コントロールマスタから取得した日数をもとに算出した日付を取得
    Select Case Me.cmb_Yukbn
        Case P_YUKBN_01  ' 自社理由
            strTargetDate = DateAdd("m", intJISYARIYU_TLM, Me.txt_Kudate)  ' 自社理由取置期限月数
        Case P_YUKBN_02  ' 取置
            strTargetDate = DateAdd("d", intTORI_DD, Me.txt_Kudate)        ' 取置期限日数
        Case P_YUKBN_10  ' 受付
            strTargetDate = DateAdd("d", intKEIUKE_TLD, Me.txt_Kudate)     ' 契約受付期限日数
            ' 希望開始日が未入力且つ、移動予約でなければ受付日＋サービス期間で初期設定する
            '▼ 2006/09/26 rep tajima ▼
            'If Nz(Me.txt_Kkdate, "") = "" And Me.cmb_MoveKbn = P_MOVEKBN_00 Then   'DELETE 2021/04/28 N.IMAI
            If (Nz(Me.txt_Kkdate, "") = "" And Me.cmb_MoveKbn = P_MOVEKBN_00) _
            Or (pstrYUKBN = P_YUKBN_00 And Me.cmb_Yukbn = P_YUKBN_10) Then          'INSERT 2021/04/28 N.IMAI
               '2020/08/27 UPD Takenouchi Start
                'Me.txt_Kkdate = DateAdd("d", intサービス期間, Me.txt_Kudate)
                Me.txt_Kkdate = strKkdateTmp
                '2020/08/27 UPD Takenouchi End
                ' もし希望開始日が営業終了日を越えていたならば、営業終了日までとする
                If Me.txt_Kkdate <> "" Then                                     'INSERT 2022/02/14 N.IMAI
                    If CDate(str営業終了日) < CDate(Me.txt_Kkdate) Then Me.txt_Kkdate = str営業終了日
                End If
                ' もし希望開始日が営業開始日以前ならば営業開始日とする
                If Me.txt_Kkdate <> "" Then                                     'INSERT 2022/02/14 N.IMAI
                    If str営業開始日 <> "" Then
                        If CDate(str営業開始日) > CDate(Me.txt_Kkdate) Then Me.txt_Kkdate = str営業開始日
                    End If
                End If
            End If
            '▲ 2006/09/26 rep tajima ▲
    End Select

    ' ヤード最終使用日付の取得
    '▼ 2006/09/26 rep tajima ▼
    ' strLastDate = GetYardUseEndDate(pobjKONT_DB, CLng(Me.cmb_Yard), strErrMsg)
    ' If Nz(strLastDate, "") <> "" Then strLastDate = Format(strLastDate, "YYYY/MM/DD")
    ' 地主契約を見るのではなく ヤード営業終了日をあてる
    strLastDate = str営業終了日
    '▲ 2006/09/26 rep tajima ▲

    If Nz(strLastDate, "") <> "" Then
        ' ヤード最終使用日付が取得できたとき
        If CDate(strLastDate) < CDate(strTargetDate) Then
            Me.txt_Tkdate = strLastDate
        Else
            Me.txt_Tkdate = strTargetDate
        End If
    Else
        ' ヤード最終使用日付が取得できなかったとき
        ' 契約受付日 + 自社理由取置期限月数 または
        ' 契約受付日 + 取置期限日数         または
        ' 契約受付日 + 契約受付期限日数
        Me.txt_Tkdate = strTargetDate
    End If

    ' コンテナDB切断
    If fncDisConnectDB = False Then GoTo subSetTKDATE_Exit

subSetTKDATE_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKONT_DB Is Nothing Then
        ' コンテナDB切断
        If fncDisConnectDB = False Then Exit Sub
    End If
    Exit Sub

subSetTKDATE_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ：" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ：" & Err.Description
    Err.Clear
    GoTo subSetTKDATE_Exit
End Sub

'==============================================================================*
'
'        MODULE_NAME      :ご紹介トランのデータ削除
'        MODULE_ID        :fncDELETE_INTR_TRAN
'        CREATE_DATE      :2005/12/12
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDELETE_INTR_TRAN(ByRef objKONT_DB As Database, ByVal strYOUKT_UKNO As String) As Boolean

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo fncDELETE_INTR_TRAN_Err

    fncDELETE_INTR_TRAN = False

    strSQL = "SELECT * FROM INTR_TRAN "
    strSQL = strSQL & " WHERE INTRT_UKNO = '" & strYOUKT_UKNO & "' "
    Set objRs = objKONT_DB.OpenRecordset(strSQL, dbOpenDynaset)

    Do Until objRs.EOF
        objRs.Edit
        objRs.Delete

        objRs.MoveNext
    Loop
    objRs.Close
    Set objRs = Nothing

    fncDELETE_INTR_TRAN = True
    Exit Function

fncDELETE_INTR_TRAN_Err:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    fncDELETE_INTR_TRAN = False
    Call Err.Raise(Err.Number, "fncDELETE_INTR_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'        MODULE_NAME      :予約入力画面設定
'        MODULE_ID        :subSetFVS500
'        Parameter        :
'                          strKbn = 移動区分
'                          intKbn = 移動元契約番号を初期化するか否かのフラグ
'        CREATE_DATE      :2006/01/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetFVS500(ByVal strKbn As String)

    Dim strName As String
    Dim strSQL  As String
    Dim objRs   As Recordset

    '▼ 2007/08/09 add tajima
    If isネット予約 = True Then
        Me.lbl_Youk.Caption = P_TITTLE_YOYAKU & "：ネット"
    Else
        Me.lbl_Youk.Caption = P_TITTLE_YOYAKU & "：電話"
    End If
    '▲ 2007/08/09 add tajima
    
    If strKbn <> P_MOVEKBN_00 Then
        ' 移動区分名取得
        strName = ""
        strName = GetNameData(P_NAME_098, strKbn)

        ' 画面タイトル(予約入力)
        Me.lbl_Youk.Caption = Me.lbl_Youk.Caption & "(" & strName & ")" 'fix 2007/08/09 tajima

        If Me.cmb_Yukbn <> P_YUKBN_20 Then
            ' 元受注契約番号
            Me.txt_MOTO_ACPTNO.Enabled = True
            Me.txt_MOTO_ACPTNO.Locked = True

            ' サブフォームの使用を可能とする
            FVS500_FVS500S01.Enabled = True
            FVS500_FVS500S01.Locked = False

        Else
            ' 元受注契約番号
            Me.txt_MOTO_ACPTNO.Enabled = False
            Me.txt_MOTO_ACPTNO.Locked = False

            ' サブフォームの使用を可能とする
            FVS500_FVS500S01.Enabled = False
            FVS500_FVS500S01.Locked = True
        End If
        
    Else
        ' 画面タイトル(予約入力)
        'Me.lbl_Youk.Caption = P_TITTLE_YOYAKU '既に設定済みよってdel 2007/08/09 tajima

        ' 元受注契約番号
        Me.txt_MOTO_ACPTNO.Enabled = False
        Me.txt_MOTO_ACPTNO.Locked = False

        ' サブフォームの使用を不可とする
        FVS500_FVS500S01.Enabled = False
        FVS500_FVS500S01.Locked = True
    
    End If

'▼ 2007/04/06 add tajima
    '移動担当入力制御
    If strKbn = P_MOVEKBN_01 Or strKbn = P_MOVEKBN_02 Then
        Me.lbl_移動担当ラベル.Visible = True
        Me.cmb_移動担当.Enabled = True
        Me.txt_移動担当.Enabled = True
        If Nz(Me.cmb_移動担当.VALUE, "") = "" Then
           Me.cmb_移動担当.VALUE = 0
        End If
    Else
        Me.lbl_移動担当ラベル.Visible = False
        Me.cmb_移動担当.Enabled = False
        Me.cmb_移動担当.VALUE = ""
        Me.txt_移動担当.Enabled = False
        Me.txt_移動担当.VALUE = ""
    End If
    Call cmb_移動担当_AfterUpdate
'▲ 2007/04/06 add tajima
    
    
    If strKbn = P_MOVEKBN_00 Then
        ' 「移動区分 = 無し」のときの処理
        Me.txt_MOTO_ACPTNO = ""
        
        doCmd.Echo False

        ' サブフォーム表示用のワークテーブルを初期化
        strSQL = "SELECT * FROM FVS500_W01"
        Set objRs = objDb.OpenRecordset(strSQL, dbOpenDynaset)

        With objRs
          Do Until .EOF

            .Edit
            .Fields("CARG_ACPTNO_Dummy") = ""
            .UPDATE
    
            .MoveNext
          Loop
        End With
        objRs.Close
        Set objRs = Nothing

        Me.FVS500_FVS500S01!FVS500S01_YCODE = ""  ' ヤードコード
        Me.FVS500_FVS500S01!FVS500S01_NO = ""     ' コンテナ番号

        Me.FVS500_FVS500S01.Requery
        doCmd.Echo True
        
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :
'        MODULE_ID        :dncCheckUketsuke
'        Parameter        :
'                          strUKNO       = 予約番号
'                          strMotoAcptNo = 元受注契約番号
'                          strRcptNo     = 受付番号
'                          strErrMsg     = エラーメッセージ格納
'        CREATE_DATE      :2006/01/18
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function dncCheckUketsuke(ByVal strUkno As String, _
                                  ByRef strMotoAcptNo As String, ByRef strRcptNo As String, _
                                  ByRef strErrMsg As String) As Boolean

    Dim strSQL   As String
    Dim objRs    As Recordset
    Dim intYUKBN As Integer

    On Error GoTo dncCheckUketsuke_Err

    dncCheckUketsuke = False

    intYUKBN = 0
    strMotoAcptNo = ""
    strRcptNo = ""
    strErrMsg = ""

    If Nz(strUkno, "") = "" Then Exit Function

    ' コンテナDBに接続
    If fncConnectDB = False Then Exit Function

    strSQL = "SELECT  YOUKT_YUKBN, " & Chr(13)          ' 予約受付状態区分
    strSQL = strSQL & " YOUKT_MOTO_ACPTNO, " & Chr(13)  ' 元受注契約番号
    strSQL = strSQL & " YOUKT_RCPT_NO " & Chr(13)       ' 受付番号
    strSQL = strSQL & "  FROM YOUK_TRAN " & Chr(13)
    strSQL = strSQL & " WHERE YOUKT_UKNO = '" & strUkno & "' "
    Set objRs = pobjKONT_DB.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)

    If objRs.EOF = False Then
        With objRs
            intYUKBN = .Fields("YOUKT_YUKBN")                    ' 予約受付状態区分
            strMotoAcptNo = Nz(.Fields("YOUKT_MOTO_ACPTNO"), "") ' 元受注契約番号
            strRcptNo = Nz(.Fields("YOUKT_RCPT_NO"), "")         ' 受付番号
        End With

        Select Case intYUKBN
            Case Is < P_YUKBN_10  ' *** 受付以外( < 10) ***
                strErrMsg = pcstrUkeErrMsg02
                GoTo dncCheckUketsuke_Exit

            Case P_YUKBN_10       ' *** 受付 ***
                ' 受付OK

            Case Is > P_YUKBN_10  ' *** 受付以外( > 10) ***
                If Nz(strRcptNo, "") <> "" Then
                    ' 受付OK
                Else
                    strErrMsg = pcstrUkeErrMsg03
                    GoTo dncCheckUketsuke_Exit
                End If
        End Select
    Else
        strErrMsg = pcstrNoDataMsg
        GoTo dncCheckUketsuke_Exit
    End If
    objRs.Close
    Set objRs = Nothing

    ' コンテナDB切断
    If fncDisConnectDB = False Then GoTo dncCheckUketsuke_Exit

    dncCheckUketsuke = True

dncCheckUketsuke_Exit:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    If Not pobjKONT_DB Is Nothing Then
        If fncDisConnectDB = False Then Exit Function
    End If
    Exit Function

dncCheckUketsuke_Err:
    MsgBox "ｴﾗｰｺｰﾄﾞ：" & Err.Number & Space(1) & "ｴﾗｰﾒｯｾｰｼﾞ：" & Err.Description
    Err.Clear
    GoTo dncCheckUketsuke_Exit
End Function
'==============================================================================*
'
'        MODULE_NAME      :受付トランのデータ削除
'        MODULE_ID        :fncDel_RCPT_TRAN
'        CREATE_DATE      :2006/03/26
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'▼2008/12/19 hirano MOD a受付契約番号をパラメータに追加.DCRA_TRNを削除のため
'Private Function fncDel_RCPT_TRAN(ByRef aDbObject As Database, ByVal a予約受付番号 As String) As Boolean
Private Function fncDel_RCPT_TRAN(ByRef aDbObject As Database, ByVal a予約受付番号 As String, Optional a受付契約番号 As String) As Boolean
'▲2008/12/19 hirano MOD
    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo Exception

    fncDel_RCPT_TRAN = False

    strSQL = "SELECT * FROM RCPT_TRAN "
    strSQL = strSQL & " WHERE RCPT_NO = '" & a予約受付番号 & "' "
    Set objRs = aDbObject.OpenRecordset(strSQL, dbOpenDynaset)

    If objRs.EOF = False Then
    
        objRs.Edit
        '▼2008/12/17 hirano add 削除データの受付契約番号を取得
        a受付契約番号 = objRs.Fields("RCPT_CARG_ACPTNO").VALUE
        '▲2008/12/17 hirano add
        objRs.Delete
    End If
    objRs.Close
    Set objRs = Nothing

    fncDel_RCPT_TRAN = True
    Exit Function

Exception:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    fncDel_RCPT_TRAN = False
    Call Err.Raise(Err.Number, "fncDel_RCPT_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)  ' 2006/04/07 DEL T.SUZUKI
End Function

'==============================================================================*
'
'        MODULE_NAME      :顧客マスタへの更新
'        MODULE_ID        :fncUPD_USER_MAST
'        CREATE_DATE      :2006/03/30
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUPD_USER_MAST(ByRef aDbObject As Database) As Boolean

    Dim strSQL  As String
    Dim objRs   As Recordset
    Dim isUpdate As Boolean

    On Error GoTo Exception

    fncUPD_USER_MAST = False

    strSQL = "SELECT * FROM USER_MAST "
    strSQL = strSQL & " WHERE USER_CODE = " & Me.cmb_KokCd.VALUE
    Set objRs = aDbObject.OpenRecordset(strSQL, dbOpenDynaset)

    isUpdate = False
    
    ' 顧客名カナ、顧客区分（法人or個人)以外は顧客マスタへの更新対象とする
    If objRs.EOF = False Then
      With objRs
       .Edit

        If Nz(Me.txt_Tanm, "") <> Nz(.Fields("USER_TANM"), "") Then
          .Fields("USER_TANM") = Nz(Me.txt_Tanm, "")
          isUpdate = True
        End If
        
        If Nz(Me.txt_Taka, "") <> Nz(.Fields("USER_TAKA"), "") Then
          .Fields("USER_TAKA") = Nz(Me.txt_Taka, "")
          isUpdate = True
        End If
        
        If Nz(Me.txt_Yubin, "") <> Nz(.Fields("USER_YUBINO"), "") Then
          .Fields("USER_YUBINO") = Nz(Me.txt_Yubin, "")
          isUpdate = True
        End If

        If Nz(Me.txt_Adr1, "") <> Nz(.Fields("USER_ADR_1"), "") Then
          .Fields("USER_ADR_1") = Nz(Me.txt_Adr1, "")
          isUpdate = True
        End If
        
        If Nz(Me.txt_Adr2, "") <> Nz(.Fields("USER_ADR_2"), "") Then
          .Fields("USER_ADR_2") = Nz(Me.txt_Adr2, "")
          isUpdate = True
        End If
    
        If Nz(Me.txt_Adr3, "") <> Nz(.Fields("USER_ADR_3"), "") Then
          .Fields("USER_ADR_3") = Nz(Me.txt_Adr3, "")
          isUpdate = True
        End If
   
        If Nz(Me.txt_Tel, "") <> Nz(.Fields("USER_TEL"), "") Then
          .Fields("USER_TEL") = Nz(Me.txt_Tel, "")
          isUpdate = True
        End If
    
        If Nz(Me.txt_Keitai, "") <> Nz(.Fields("USER_KEITAI"), "") Then
          .Fields("USER_KEITAI") = Nz(Me.txt_Keitai, "")
          isUpdate = True
        End If
    
        If Nz(Me.txt_Fax, "") <> Nz(.Fields("USER_FAX"), "") Then
          .Fields("USER_FAX") = Nz(Me.txt_Fax, "")
          isUpdate = True
        End If
        
        If Nz(Me.txt_Mail, "") <> Nz(.Fields("USER_MAIL"), "") Then
          .Fields("USER_MAIL") = Nz(Me.txt_Mail, "")
          isUpdate = True
        End If
    
        If isUpdate = True Then
          .Fields("USER_UPDATE") = DATE
          .UPDATE
        End If
        
      End With
    End If
    objRs.Close
    Set objRs = Nothing

    fncUPD_USER_MAST = True
    Exit Function

Exception:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    fncUPD_USER_MAST = False
    Call Err.Raise(Err.Number, "fncUPD_USER_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)  ' 2006/04/07 DEL T.SUZUKI
End Function

'==============================================================================*
'
'        MODULE_NAME      :KASE-DB接続
'        MODULE_ID        :fncConnectKaseDB
'        CREATE_DATE      :2006/09/24
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncConnectKaseDB() As Boolean

    Dim strDataSource As String

  On Error GoTo Exception
  
    fncConnectKaseDB = False

    ' -----------------------------------------------------------------------------------------------------------
    ' 加瀬DB接続
    strDataSource = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"), "")

    If strDataSource = "" Then
        Set pobjKASE_DB = Nothing
        Call MSZZ024_M10("fncConnectKaseDB", "SETU_TABLの設定が不正です。")
        Exit Function
    Else
        Set pobjKASE_DB = Workspaces(0).OpenDatabase(strDataSource, dbDriverNoPrompt, False, MSZZ007_M00())
    End If
    ' -----------------------------------------------------------------------------------------------------------

    fncConnectKaseDB = True

    Exit Function

Exception:
    Call Err.Raise(Err.Number, "fncConnectKaseDB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :KaseDB切断
'        MODULE_ID        :fncDisConnectKaseDB
'        CREATE_DATE      :2005/11/08
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDisConnectKaseDB() As Boolean

  On Error GoTo Exception

    fncDisConnectKaseDB = False

    If Not pobjKASE_DB Is Nothing Then
        pobjKASE_DB.Close
        Set pobjKASE_DB = Nothing
    End If
  
    fncDisConnectKaseDB = True

    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncDisConnectKaseDB" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'■□■□■□■□■□■□■□■□■□ ↓↓↓　Public 関数群 ↓↓↓ ■□■□■□■□■□■□■□■□■□■□

'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :subReload
'        CREATE_DATE      :2005/10/06
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subReload(intFlg As Integer)

    ' 画面コントロールの有効化
    Call subFormLockControl(True)  ' 画面項目の使用可否設定
'    Call SetFrmEnabledConf(True)    ' 画面項目の使用可否設定
'    Call subFormLockControl(True)

    If intFlg = 1 Then
        ' 遷移先画面にて修正があった場合、変更フラグをONに設定する
        pblnUpdateFlg = True                                                   ' 2006/03/11 ADD T.SUZUKI

        ' 予約ご紹介画面で修正があった場合、予約受付入力画面を再表示する
        Call cmb_Ukno_AfterUpdate

' 2005/12/20 ADD T.SUZUKI Start
' 2005/12/20 ADD T.SUZUKI End
    End If
End Sub
'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させるFVS90n用
'        MODULE_ID        :subReloadfromFVS90n
'        CREATE_DATE      :2006/09/24
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SubReloadfromFVS90n()
    Dim strBumoCD As String
    
  On Error GoTo Exception
    
    Call fncConnectKaseDB
    strBumoCD = DLookup("CONT_BUMOC", "dbo_CONT_MAST")
    If Nz(Me.cmb_Yard, "") <> "" And Nz(Me.cmb_No, "") <> "" Then
        Me.txt_部屋修繕 = CmReprMod.GetRoomLatestReprText(pobjKASE_DB, strBumoCD, Me.cmb_Yard, Me.cmb_No)
    Else
        Me.txt_部屋修繕 = ""
    End If
    If Nz(Me.cmb_KokCd, "") <> "" Then
        Me.txt_顧客コメントクレーム = CmReprMod.GetUserLatestKotoText(pobjKASE_DB, strBumoCD, Me.cmb_KokCd)
    Else
        Me.txt_顧客コメントクレーム = ""
    End If
    Call fncDisConnectKaseDB
    Exit Sub
    
Exception:
    Call Err.Raise(Err.Number, "subReloadfromFVS90n" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'#################### ADD 2007/04/06 ####################
'==============================================================================*
'
'        MODULE_NAME      :cmb_移動担当_AfterUpdateイベント
'        MODULE_ID        :cmb_移動担当_AfterUpdate
'        CREATE_DATE      :2007/04/06
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_移動担当_AfterUpdate()

    On Error GoTo Exception

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_移動担当, Me.txt_移動担当, "084") = False Then Exit Sub

    Exit Sub

Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_移動担当_BeforeUpdateイベント
'        MODULE_ID        :cmb_移動担当_BeforeUpdate
'        CREATE_DATE      :2007/04/06
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_移動担当_BeforeUpdate(Cancel As Integer)
    
    Dim varCode As Variant

    If bln_FormLoad = False Then Exit Sub
    
    If ControlCrear(Me.cmb_移動担当, Me.txt_移動担当) = False Then Exit Sub
    
    varCode = Me.cmb_移動担当

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo Exception

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exception
    End Select

    If Chk_NameBefore(Me.cmb_移動担当, "084") = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       GoTo Exception
    End If

    Exit Sub

Exception:
    Cancel = True
    SendKeys "{ESC}"

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_移動担当_BeforeUpdateイベント
'        MODULE_ID        :cmb_移動担当_BeforeUpdate
'        CREATE_DATE      :2007/04/06
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_移動担当_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_移動担当_BeforeUpdateイベント
'        MODULE_ID        :cmb_移動担当_BeforeUpdate
'        CREATE_DATE      :2007/04/06
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_移動担当_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_移動担当, KeyAscii, 2)
End Sub

'★☆★☆★☆★☆★☆★★☆☆ 2007/08/08 add tajima ★☆★☆★☆★☆★☆★☆★☆
'==============================================================================*
'
'        MODULE_NAME      :cmb_予定期間_AfterUpdateイベント
'        MODULE_ID        :cmb_予定期間_AfterUpdate
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_予定期間_AfterUpdate()
    On Error GoTo Exception

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_予定期間, Me.txt_予定期間, "249") = False Then Exit Sub

    Exit Sub

Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_予定期間_BeforeUpdateイベント
'        MODULE_ID        :cmb_予定期間_BeforeUpdate
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_予定期間_BeforeUpdate(Cancel As Integer)

    Dim varCode As Variant

    If bln_FormLoad = False Then Exit Sub
    
    If ControlCrear(Me.cmb_予定期間, Me.txt_予定期間) = False Then Exit Sub
    
    varCode = Me.cmb_予定期間

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo Exception

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exception
    End Select

    If Chk_NameBefore(Me.cmb_予定期間, "249") = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       GoTo Exception
    End If

    Exit Sub

Exception:
    Cancel = True
    SendKeys "{ESC}"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_予定期間_Changeイベント
'        MODULE_ID        :cmb_予定期間_Change
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_予定期間_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_予定期間_KeyPressイベント
'        MODULE_ID        :cmb_予定期間_KeyPress
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_予定期間_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_予定期間, KeyAscii, 2)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_申込書送付方_AfterUpdateイベント
'        MODULE_ID        :cmb_申込書送付方_AfterUpdate
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_申込書送付方_AfterUpdate()

    On Error GoTo Exception

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_申込書送付方, Me.txt_申込書送付方法, "012") = False Then Exit Sub

    Exit Sub

Exception:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_申込書送付方_BeforeUpdateイベント
'        MODULE_ID        :cmb_申込書送付方_BeforeUpdate
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_申込書送付方_BeforeUpdate(Cancel As Integer)
    Dim varCode As Variant

    If bln_FormLoad = False Then Exit Sub
    
    If ControlCrear(Me.cmb_申込書送付方, Me.txt_申込書送付方法) = False Then Exit Sub
    
    varCode = Me.cmb_申込書送付方

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            GoTo Exception

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo Exception
    End Select

    If Chk_NameBefore(Me.cmb_申込書送付方, "012") = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       GoTo Exception
    End If

    Exit Sub

Exception:
    Cancel = True
    SendKeys "{ESC}"
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_申込書送付方_Changeイベント
'        MODULE_ID        :cmb_申込書送付方_Change
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_申込書送付方_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_申込書送付方_KeyPressイベント
'        MODULE_ID        :cmb_申込書送付方_KeyPress
'        CREATE_DATE      :2007/08/09
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_申込書送付方_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(Me.cmb_申込書送付方, KeyAscii, 2)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :適用割引割付のデータ削除
'        MODULE_ID        :fncDel_DRCA_TRAN
'        CREATE_DATE      :2008/11/19
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDel_DRCA_TRAN(ByRef aDbObject As Database, ByVal a予約受付番号 As String) As Boolean

    Dim strSQL  As String
    Dim objRs   As Recordset

    On Error GoTo Exception

    fncDel_DRCA_TRAN = False

    strSQL = "SELECT * FROM DCRA_TRAN "
    strSQL = strSQL & " WHERE DCRAT_ACPTNO = '" & a予約受付番号 & "' "
    Set objRs = aDbObject.OpenRecordset(strSQL, dbOpenDynaset)

    If objRs.EOF = False Then
    
        Do Until objRs.EOF
            objRs.Edit
            objRs.Delete
            objRs.MoveNext
        Loop
    End If
    objRs.Close
    Set objRs = Nothing

    fncDel_DRCA_TRAN = True
    Exit Function

Exception:
    If Not objRs Is Nothing Then objRs.Close: Set objRs = Nothing
    fncDel_DRCA_TRAN = False
    Call Err.Raise(Err.Number, "fncDel_DRCA_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :cmb_発生区分_AfterUpdate
'        機能             :発生区分 AfterUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_発生区分_AfterUpdate()
    
    On Error GoTo cmb_発生区分_AfterUpdate_Err

    'ファイルへの存在チェック
    If Chk_NameAfter(Me.cmb_発生区分, Me.txt_発生区分, P_NAME_083) = False Then GoTo cmb_発生区分_AfterUpdate_End

    If Me.cmb_発生区分 = "00" Then
        isネット予約 = False
        Me.lbl_Youk.Caption = P_TITTLE_YOYAKU & "：電話"
    Else
        isネット予約 = True
        Me.lbl_Youk.Caption = P_TITTLE_YOYAKU & "：ネット"
    End If
    
    GoTo cmb_発生区分_AfterUpdate_End

cmb_発生区分_AfterUpdate_Err:
    MsgBox "ｴﾗｰ番号" & Err.Number & vbCrLf & Err.Description
    Err.Clear

cmb_発生区分_AfterUpdate_End:
    Exit Sub

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_発生区分_BeforeUpdate
'        機能             :発生区分 BeforeUpdateｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_発生区分_BeforeUpdate(Cancel As Integer)
    
    If bln_FormLoad = False Then Exit Sub

    'ｺｰﾄﾞが空の場合担当者名欄も空欄表示
    If ControlCrear(Me.cmb_発生区分, Me.txt_発生区分) = False Then Exit Sub

    Dim varCode As Variant

    varCode = Me.cmb_発生区分

    '入力チェック
    Select Case CmbCheck_Num(varCode, 2)
        Case 0

        Case 2
            '桁数ｴﾗｰ
            MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_発生区分_BeforeUpdate_Exit

        Case 9
            MsgBox "コードを入力してください。", vbInformation, P_ERROR
            Cancel = True
            SendKeys "{ESC}"
            GoTo cmb_発生区分_BeforeUpdate_Exit
    End Select

    'ファイルへの存在チェック
    If Chk_NameBefore(Me.cmb_発生区分, P_NAME_083) = "" Then
       MsgBox "登録されていません。", vbInformation, P_KAKUNIN
       Cancel = True
       SendKeys "{ESC}"
    End If

cmb_発生区分_BeforeUpdate_Exit:
    Exit Sub
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_発生区分_Changeイベント
'        MODULE_ID        :cmb_発生区分_Change
'        CREATE_DATE      :2011/07/19
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_発生区分_Change()
    Call subSetChangeFlag
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_発生区分_KeyPress
'        機能             :発生区分 KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_発生区分_KeyPress(KeyAscii As Integer)
    '桁数チェック追加
    Call CmKeyPress_For_Num1(Me.cmb_発生区分, KeyAscii, 2)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : Webも考慮し空き件数をチェックする
'       MODULE_ID       : Chk_SPAC_AKI
'       CREATE_DATE     : 2014/12/12            K.ISHIZAKA
'       RETURN          : 空きあり(True)／なし(False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function Chk_SPAC_AKI() As Boolean
    Dim strUkno             As String   '受付番号
    Dim strNO               As String   'コンテナ番号
    Dim strHEYA             As Integer    '部屋指定
    On Error GoTo ErrorHandler
    
    '2016/10/29 M.HONDA INS
    strHEYA = Nz(DLookup("YARD_HEYA_SHITEI_KBN", "dbo_YARD_MAST", "YARD_CODE = " & Me.cmb_Yard), "")
    '2016/10/29 M.HONDA INS
    
    strUkno = IIf(Nz(Me.cmb_Ukno, "") = "", Me.txt_Ukno, Me.cmb_Ukno)
    If strUkno <> "" Then
        strNO = Nz(DLookup("YOUKT_NO", "dbo_YOUK_TRAN", "YOUKT_UKNO = '" & strUkno & "'"), "")
    Else
        strNO = ""
    End If
   'コンテナ番号を指定したときで、変化した場合
    'If (Nz(cmb_No.Value) <> "") And (Nz(cmb_No.Value) <> strNO) Then           'DELETE 2018/03/01 N.IMAI
    If (Nz(cmb_No.VALUE) <> "") And (Val(Nz(cmb_No.VALUE)) <> Val(strNO)) Then  'INSERT 2018/03/01 N.IMAI
        '空き数チェック
        If MSTGP01_M20(cmb_Yard.VALUE, cmb_No.VALUE, , strHEYA) <= 0 Then
            MsgBox "該当する段、サイズでの空きがありません。", vbOKOnly + vbInformation, Me.Caption
            Chk_SPAC_AKI = False
            Exit Function
        End If
    End If

    Chk_SPAC_AKI = True
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Chk_SPAC_AKI" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

' ▼ 2018/08/30 TAKENOUCHI INS<S> ▼
'==============================================================================*
'
'        MODULE_NAME      :txt_SyzeF_KeyPress
'        機能             :希望サイズFrom KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        CREATE_DATE      :2018/08/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SyzeF_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Dec(Me.txt_SyzeF, KeyAscii, 5, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt_SyzeT_KeyPress
'        機能             :希望サイズTo KeyPressｲﾍﾞﾝﾄ処理
'        IN               :
'        OUT              :
'        CREATE_DATE      :2018/08/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SyzeT_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Dec(Me.txt_SyzeT, KeyAscii, 5, 2)

End Sub

'==============================================================================*
'
'        MODULE_NAME      : txt_SyzeF_BeforeUpdateイベント
'        MODULE_ID        : txt_SyzeF_BeforeUpdate
'        CREATE_DATE      : 2018/08/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SyzeF_BeforeUpdate(Cancel As Integer)

    If fncTxtBetweenCheck(txt_SyzeF, 0.01, 99.9) = False Then
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      : txt_SyzeT_BeforeUpdateイベント
'        MODULE_ID        : txt_SyzeT_BeforeUpdate
'        CREATE_DATE      : 2018/08/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt_SyzeT_BeforeUpdate(Cancel As Integer)
    
    If fncTxtBetweenCheck(txt_SyzeT, 0.01, 99.9) = False Then
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      : テキストボックスの入力値の範囲チェック
'        MODULE_ID        : subTxtBetweenCheck
'        CREATE_DATE      : 2018/08/30
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncTxtBetweenCheck(ctlText As TextBox, ByVal dblMin As Double, ByVal dblMax As Double) As Boolean
    
    Dim dblVal    As Double
    
    fncTxtBetweenCheck = False
    
    If Nz(ctlText, "") <> "" Then
        dblVal = CDbl(ctlText)
        
        If (dblVal < dblMin) Or (dblVal > dblMax) Then
            MsgBox "入力値が不正です。0.01～99.9の数値を入力してください。", vbInformation, P_ERROR
            Exit Function
        End If
    End If

    fncTxtBetweenCheck = True

End Function
' ▲ 2018/08/30 TAKENOUCHI INS<E> ▲


'==============================================================================*
'
'        MODULE_NAME      :コンボボックス定義共通
'        MODULE_ID        :make_cmb
'        CREATE_DATE      :2020/01/01 Y.WADA
'        PARAM            :作業テーブルを作成する
'                         :作業テーブル名
'                         :コンボボックス・オブジェクト
'                         :作業テーブルの元SELECT文
'                         :コンボボックスのSELECT文（省略可）
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub make_cmb(blnCreate As Boolean, strWKTblNm As String, ctlCmb As ComboBox, strSQL As String, Optional strSQL2 As String = "")
    
    Dim strBUMOC        As String
    Dim objCon          As Object
    
    On Error GoTo ErrorHandler
    
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    
    On Error GoTo ErrorHandler2
    
    Set objCon = ADODB_Connection(strBUMOC)
    
    If blnCreate Then
        Call ADODB_CreateWorkTable_CurrentDb(strWKTblNm, strSQL, objCon)
    Else
        Call ADODB_Reset_CurrentDb(strWKTblNm, strSQL, objCon)
    End If
    
    objCon.Close
    On Error GoTo ErrorHandler
    
    If strSQL2 = "" Then
        '省略時のSelect文
        strSQL2 = ""
        strSQL2 = strSQL2 & " SELECT cd, nm "
        strSQL2 = strSQL2 & " FROM "
        strSQL2 = strSQL2 & strWKTblNm
        strSQL2 = strSQL2 & " ORDER BY "
        strSQL2 = strSQL2 & " cd "
        strSQL2 = strSQL2 & ";"
    End If

    With ctlCmb
        .RowSourceType = "Table/Query"
        .RowSource = strSQL2
    End With

Exit Sub

ErrorHandler2:
    objCon.Close

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "make_cmb" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :用途コンボボックス定義
'        MODULE_ID        :make_cmb_Youto
'        CREATE_DATE      :2020/01/01 Y.WADA
'        PARAM            :作業テーブルを作成する　（省略値：False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub make_cmb_Youto(Optional blnCreate As Boolean = False)
    
    Dim strSQL          As String
    
    On Error GoTo ErrorHandler

    strSQL = ""
    strSQL = strSQL & vbCrLf & "DECLARE @YARD_CODE AS NUMERIC(6) = " & IIf(Nz(cmb_Yard.VALUE) <> "", cmb_Yard.VALUE, "NULL") & ";"
    strSQL = strSQL & vbCrLf & "SELECT "
    strSQL = strSQL & vbCrLf & "  CAST(FORMAT(NAME_CODE,'00') AS VARCHAR(2)) AS CD"
    strSQL = strSQL & vbCrLf & ", NAME_NAME AS NM"
    strSQL = strSQL & vbCrLf & "FROM "
    strSQL = strSQL & vbCrLf & "    NAME_MAST"
    strSQL = strSQL & vbCrLf & "WHERE NAME_ID = '086'"
    strSQL = strSQL & vbCrLf & "AND NAME_CODE NOT IN (33, 39, 52, 54, 55) " '33コンテナ（バイク）、39バイク(コンテナ)、52自販機（飲料水）、54自販機（ﾊﾞﾗｴﾃｨ）、55自販機（その他）
    strSQL = strSQL & vbCrLf & "AND EXISTS "
    strSQL = strSQL & vbCrLf & "    ("
    strSQL = strSQL & vbCrLf & "         SELECT * "
    strSQL = strSQL & vbCrLf & "         FROM CNTA_MAST "
    strSQL = strSQL & vbCrLf & "         WHERE CNTA_CODE = @YARD_CODE"
    strSQL = strSQL & vbCrLf & "          AND CNTA_USAGE = NAME_CODE"
    strSQL = strSQL & vbCrLf & "    )"
    strSQL = strSQL & vbCrLf & "ORDER BY 1"
    
    Call make_cmb(blnCreate, WORK_TAB01, cmb_Youto, strSQL)

Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "make_cmb_Youto" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub


'==============================================================================*
'
'        MODULE_NAME      :希望サイズコンボボックス定義
'        MODULE_ID        :make_cmb_Youto
'        CREATE_DATE      :2020/01/01 Y.WADA
'        PARAM            :作業テーブルを作成する　（省略値：False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub make_cmb_Kibou(Optional blnCreate As Boolean = False)
    
    Dim strSQL          As String
    
    On Error GoTo ErrorHandler
    
    strSQL = ""
    strSQL = strSQL & vbCrLf & "DECLARE @YARD_CODE AS NUMERIC(6) = " & IIf(Nz(cmb_Yard.VALUE) <> "", cmb_Yard.VALUE, "NULL") & ";"
    strSQL = strSQL & vbCrLf & "DECLARE @USAGE AS NUMERIC(2) = " & IIf(Nz(cmb_Youto.VALUE) <> "", cmb_Youto.VALUE, "NULL") & ";"
    strSQL = strSQL & vbCrLf & "SELECT "
    strSQL = strSQL & vbCrLf & "  CAST(FORMAT(NAME_CODE,'00') AS VARCHAR(2)) AS CD"
    strSQL = strSQL & vbCrLf & ", NAME_NAME AS NM"
    strSQL = strSQL & vbCrLf & "FROM "
    strSQL = strSQL & vbCrLf & "    NAME_MAST"
    strSQL = strSQL & vbCrLf & "WHERE NAME_ID = '087'"
    strSQL = strSQL & vbCrLf & "AND EXISTS"
    strSQL = strSQL & vbCrLf & "    ("
    strSQL = strSQL & vbCrLf & "         SELECT * "
    strSQL = strSQL & vbCrLf & "         FROM CNTA_MAST "
    strSQL = strSQL & vbCrLf & "         WHERE CNTA_CODE = @YARD_CODE"
    strSQL = strSQL & vbCrLf & "          AND CNTA_USAGE = @USAGE"
    strSQL = strSQL & vbCrLf & "          AND CNTA_SIZE BETWEEN NAME_VALUE_FROM AND NAME_VALUE_TO"
    strSQL = strSQL & vbCrLf & "    )"
    strSQL = strSQL & vbCrLf & "ORDER BY 1"
    
    Call make_cmb(blnCreate, WORK_TAB02, cmb_Kibou, strSQL)

Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "make_cmb_Kibou" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :段区分コンボボックス定義
'        MODULE_ID        :make_cmb_Youto
'        CREATE_DATE      :2020/01/01 Y.WADA
'        PARAM            :作業テーブルを作成する　（省略値：False)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub make_cmb_Step(Optional blnCreate As Boolean = False)
    
    Dim strSQL          As String
    Dim strSQL2         As String
    
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & vbCrLf & "DECLARE @YARD_CODE AS NUMERIC(6) = " & IIf(Nz(cmb_Yard.VALUE) <> "", cmb_Yard.VALUE, "NULL") & ";"
    strSQL = strSQL & vbCrLf & "DECLARE @USAGE AS NUMERIC(2) = " & IIf(Nz(cmb_Youto.VALUE) <> "", cmb_Youto.VALUE, "NULL") & ";"
    strSQL = strSQL & vbCrLf & "DECLARE @SIZE AS NUMERIC(2) = " & IIf(Nz(cmb_Kibou.VALUE) <> "", cmb_Kibou.VALUE, "NULL") & ";"
    
    strSQL = strSQL & vbCrLf & "SELECT CD, NM"
    strSQL = strSQL & vbCrLf & ", MIN(PRIC_PRICE) AS MIN_P"
    strSQL = strSQL & vbCrLf & ", MAX(PRIC_PRICE) AS MAX_P"
    strSQL = strSQL & vbCrLf & "FROM"
    strSQL = strSQL & vbCrLf & "("
    strSQL = strSQL & vbCrLf & "    SELECT "
    strSQL = strSQL & vbCrLf & "      CAST(FORMAT(NAME_CODE,'00') AS VARCHAR(2)) AS CD"
    strSQL = strSQL & vbCrLf & "    , NAME_NAME AS NM"
    strSQL = strSQL & vbCrLf & "    , PRIC_PRICE"
    strSQL = strSQL & vbCrLf & "    FROM "
    strSQL = strSQL & vbCrLf & "        NAME_MAST C1"
    strSQL = strSQL & vbCrLf & "     , (SELECT NAME_VALUE_FROM, NAME_VALUE_TO FROM NAME_MAST WHERE NAME_ID = '087' AND NAME_CODE = @SIZE) C3"
    strSQL = strSQL & vbCrLf & "    LEFT JOIN"
    strSQL = strSQL & vbCrLf & "        PRIC_TABL"
    strSQL = strSQL & vbCrLf & "            ON  PRIC_YCODE = @YARD_CODE"
    strSQL = strSQL & vbCrLf & "            AND PRIC_USAGE = @USAGE"
    strSQL = strSQL & vbCrLf & "            AND PRIC_SIZE BETWEEN C3.NAME_VALUE_FROM AND C3.NAME_VALUE_TO"
    strSQL = strSQL & vbCrLf & "    WHERE NAME_ID = '090'"
    strSQL = strSQL & vbCrLf & "    AND  C1.NAME_CODE = PRIC_STEP "   '2020/02/6 M.HONDA INS
    strSQL = strSQL & vbCrLf & "    AND EXISTS"
    strSQL = strSQL & vbCrLf & "        ("
    strSQL = strSQL & vbCrLf & "             SELECT * "
    strSQL = strSQL & vbCrLf & "             FROM CNTA_MAST "
    strSQL = strSQL & vbCrLf & "             , (SELECT NAME_VALUE_FROM, NAME_VALUE_TO FROM NAME_MAST WHERE NAME_ID = '087' AND NAME_CODE = @SIZE) C2"
    strSQL = strSQL & vbCrLf & "             WHERE CNTA_CODE = @YARD_CODE"
    strSQL = strSQL & vbCrLf & "              AND CNTA_USAGE = @USAGE"
    strSQL = strSQL & vbCrLf & "              AND CNTA_SIZE BETWEEN C2.NAME_VALUE_FROM AND C2.NAME_VALUE_TO"
    strSQL = strSQL & vbCrLf & "              AND CNTA_STEP = C1.NAME_CODE"
    strSQL = strSQL & vbCrLf & "        )"
    strSQL = strSQL & vbCrLf & ") A"
    strSQL = strSQL & vbCrLf & "GROUP BY CD, NM"
    strSQL = strSQL & vbCrLf & "ORDER BY 1"
    
    strSQL2 = ""
    strSQL2 = strSQL2 & " SELECT cd, nm, min_p, max_p "
    strSQL2 = strSQL2 & " FROM "
    strSQL2 = strSQL2 & WORK_TAB03
    strSQL2 = strSQL2 & " ORDER BY "
    strSQL2 = strSQL2 & " cd "
    strSQL2 = strSQL2 & ";"
   
    Call make_cmb(blnCreate, WORK_TAB03, cmb_Step, strSQL, strSQL2)

Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "make_cmb_Step" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :起算日導出用の日数取得処理
'        MODULE_ID        :getNissuForKisanDate
'        CREATE_DATE      :2020/07/03 Takenouchi
'        PARAM            :受付日
'                         :起算日導出のための加算月
'                         :起算日導出のための加算日
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub getNissuForKisanDate(ByVal aKudate As String, ByRef anAddMon As Integer, ByRef anAddDay As Integer)

    Dim strWhere     As String
    Dim strKuDate    As String
    Dim intKudateD   As Integer

    On Error GoTo getNissuForKisanDate_On_Err

    anAddMon = 0
    anAddDay = 0
    strKuDate = Format(aKudate, "YYYYMMDD")
    intKudateD = Format(aKudate, "D")

    '制御テーブルから起算日導出用の日数を取得する
    strWhere = "CTRLT_CODE LIKE 'KISAN_PTN*'"
    strWhere = strWhere & " AND CTRLT_KEY01 <= '" & strKuDate & "'"
    strWhere = strWhere & " AND CTRLT_KEY02 >= '" & strKuDate & "'"
    strWhere = strWhere & " AND Val(CTRLT_KEY03) <= " & intKudateD
    strWhere = strWhere & " AND Val(CTRLT_KEY04) >= " & intKudateD
    
    '加算月指定情報取得
    anAddMon = Nz(DLookup("CTRLT_FUNC01", "dbo_CTRL_TABL", strWhere), 0) '2020/09/30 EGL "dbo_"を付加
    '加算日指定情報取得
    anAddDay = Nz(DLookup("CTRLT_FUNC02", "dbo_CTRL_TABL", strWhere), 0) '2020/09/30 EGL "dbo_"を付加

Exit Sub

getNissuForKisanDate_On_Err:
    Call Err.Raise(Err.Number, "getNissuForKisanDate" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :予約番号コンボ初期値設定
'        MODULE_ID        :subComboUKNOInit
'        CREATE_DATE      :2024/08/27           N.IMAI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subComboUKNOInit(intDays As Integer)

    Dim strSQL      As String
    Dim strBUMOC    As String
    Dim objCon      As Object
    
    strSQL = ""
    strSQL = strSQL & vbCrLf & "SELECT YOUKT_UKNO FROM YOUK_TRAN "
    strSQL = strSQL & vbCrLf & " WHERE YOUKT_UKDATE >= GETDATE() - " & intDays
    strSQL = strSQL & vbCrLf & " ORDER BY YOUKT_UKNO DESC "
    
    On Error GoTo ErrorHandler1
    
    strBUMOC = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    Set objCon = ADODB_Connection(strBUMOC)
    
    Call ADODB_CreateWorkTable_CurrentDb(UKNO_WK_TABLE, strSQL, objCon)
    cmb_Ukno.RowSource = UKNO_WK_TABLE

ErrorHandler1:
    objCon.Close

ErrorHandler:
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "subComboUKNOInit" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub
