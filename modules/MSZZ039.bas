Attribute VB_Name = "MSZZ039"
'****************************  strat of program ********************************
'==============================================================================*
'
'        SYSTEM_NAME     : 加瀬システム
'        SUB_SYSTEM_NAME : 共通関数
'
'        PROGRAM_NAME    : 統合ヘルパー出力
'        PROGRAM_ID      : MSZZ039
'        PROGRAM_KBN     : MODULE
'
'        CREATE          : 2007/08/20
'        CERATER         : taima
'        Ver             : 0.0
'
'        UPDATE          : 2007/09/01
'        UPDATER         : K.ISHIZAKA
'        Ver             : 0.1
'                          予約受付状態区分を追加
'
'        UPDATE          : 2007/09/17
'        UPDATER         : tajima
'        Ver             : 0.2
'                          請求期間対応
'
'        UPDATE          : 2007/11/15
'        UPDATER         : SHIBAZAKI
'        Ver             : 0.3
'                          受付画面からの呼び出しに対応
'
'        UPDATE          : 2008/04/01
'        UPDATER         : tajima
'        Ver             : 0.4
'                          バーチャルオフィスヘルパー対応
'
'        UPDATE          : 2008/09/25
'        UPDATER         : SHIBAZAKI
'        Ver             : 0.5
'                          臨時保証区分対応
'                          1.印刷シートを減らす-"保証委託契約書"&"保証会社個人情報取扱"
'                          2.「保証委託料」→「事務手数料」
'
'        UPDATE          : 2008/11/13
'        UPDATER         : SHIBAZAKI
'        Ver             : 0.6
'                          保証区分追加対応
'                          1.統合ヘルパーを保証委託料用と事務手数料用の二つにする。
'                            NAME_MASTの保証区分のＦＲＯＭ値でファイル名を決定する。
'
'        UPDATE          : 2008/10/27
'        UPDATER         : iizuka
'        Ver             : 0.7
'                          割引情報出力対応
'                           受付入力　割引きに対しての「明細をみる」ボタンの場合は、
'                                     ヘルパーの「明細書」シートのみを表示させる。
'
'        UPDATE          : 2008/11/27
'        UPDATER         : K.KINEBUCHI
'        Ver             : 0.8
'                          MSZZ045（自動割引制御）を修正に伴い、セット項目を変更する。
'
'        UPDATE          : 2009/01/16
'        UPDATER         : hirano
'        Ver             : 0.9
'                          明細書、申込書の「月額使用料」は、解約までの割引を加味する。
'                          解約届書、住所変更届書も同様にする。
'                          申込書：初回使用料から初月割引分を加味する。
'
'
'        UPDATE          : 2009/02/02
'        UPDATER         : hirano
'        Ver             : 1.0
'                          明細書、申込書の「月額使用料」は、解約までの割引を加味する箇所の不具合
'                          （月額割引を保持していない為、保証委託料ベースで計算していたが、
'                            保証委託料がゼロがセットされる場合があるので、月額値引き計より算出に変更）。
'
'        UPDATE          : 2009/02/13
'        UPDATER         : hirano
'        Ver             : 1.1
'                          明細書、申込書の「月額使用料」の割引分は、画面上の項目より取得する
'
'        UPDATE          : 2009/02/20
'        UPDATER         : hirano
'        Ver             : 1.2
'                          保証会社複数対応　保証区分：１，２に対してのヘルパーファイル名の指定
'
'        UPDATE          : 2009/04/01
'        UPDATER         : hirano
'        Ver             : 1.3
'                          保証会社複数対応φ２　保証会社コード追加に伴い　保証会社コードファイルからシートコピー
'
'        UPDATE          : 2009/06/30
'        UPDATER         : hirano
'        Ver             : 1.4
'                          ネット予約自動割引　事務手数料ゼロは、印字する
'
'        UPDATE          : 2009/10/02
'        UPDATER         : hirano
'        Ver             : 1.5
'                          レンタル用途33 :コンテナ（バイク） 追加
'
'        UPDATE          : 2009/10/23
'        UPDATER         : hirano
'        Ver             : 1.6
'                          保証会社自動切替対応
'
'        UPDATE          : 2010/01/09
'        UPDATER         : M.RYU
'        Ver             : 1.7
'                         --フォームForm_FVS400と関連修正
'                         1.「保証会社無し」の場合、申込書などを「統合ヘルパー00」で出力
'                         2.臨時請求書に各印を挿入
'                         3.「統合ヘルパー」に、明細書シートのテーマが”請求書”になる場合、各印を挿入
'                         4.新規契約時の明細表（経理用）の【保証委託料】の表示欄に保証会社名の表示を追加
'                         5.統合ヘルパーに”毎月の支払い”シートを削除、出力しない
'
'        UPDATE          : 2010/03/25
'        UPDATER         : M.RYU
'        Ver             : 1.8
'                         1.【受付入力画面⇒申込書印刷】
'                               ⇒ヤード"031102"のみ約款別
'
'        UPDATE          : 2010/03/26
'        UPDATER         : M.RYU
'        Ver             : 1.9
'                         1.【受付入力画面⇒鍵同封書類印刷】
'                               ⇒「解約予定日」がある場合、シート「毎月の支払いについて」を出力しない
'                               ⇒「解約予定日」が無し場合、シート「毎月の支払いについて」を出力
'
'        UPDATE          : 2010/03/29
'        UPDATER         : M.RYU
'        Ver             : 2.0
'                         1.【受付入力画面⇒申込書印刷】⇒シート出力順番調整
'
'        UPDATE          : 2010/05/31
'        UPDATER         : M.RYU
'        Ver             : 2.1
'                         1.バーチャルオフィスの場合、ヘルパー出力についての修正
'                           （統合ヘルパー00を使用）、旧バーチャルオフィスヘルパーを廃棄
'                         2.臨時の場合、旧ヘルパーを使用
'
'        UPDATE          : 2010/06/18
'        UPDATER         : M.RYU
'        Ver             : 2.2
'                         統合ヘルパー出力修正：全部同じ出力方法にまとめ
'
'        UPDATE          : 2010/06/30
'        UPDATER         : M.RYU
'        Ver             : 2.3
'                         1.パルマ・アールエム保証会社の場合、統合ヘルパー変更
'                         2.バイク契約誓約書に契約NO.を追加
'
'        UPDATE          : 2010/09/02
'        UPDATER         : M.RYU
'        Ver             : 2.4
'                         1.統合ヘルパー変更⇒クレジット・口座振替申込書を追加⇒鍵同封書類で出力
'
'        UPDATE          : 2010/12/20
'        UPDATER         : M.HONDA
'        Ver             : 2.5
'                          口座振替・クレジットカード単独出力対応
'                          1.口座振替・クレジットカード申込用紙を単独で出力できるようにする。
'
'        UPDATE          : 2011/02/06
'        UPDATER         : K.ISHIZAKA
'        Ver             : 2.6
'                          ネット予約から毎月支払方法が選択されたときの対応版（丸が表示される）
'
'        UPDATE          : 2011/03/02
'        UPDATER         : K.ISHIZAKA
'        Ver             : 2.7
'                          インターネット予約バッチのMSZZ039と統合する
'                           HelperPrintPDF3 を追加
'
'        UPDATE          : 2011/03/25
'        UPDATER         : M.HONDA
'        Ver             : 2.8
'                          ネット予約でクレジット登録案内を出している顧客には
'                          支払い方法の案内用紙を出力しない。
'
'        UPDATE          : 2011/06/01
'        UPDATER         : M.RYU
'        Ver             : 2.9
'                          オフィスに「自習室」用途を追加したのため
'
'        UPDATE          : 2011/06/18
'        UPDATER         : tajima
'        Ver             : 3.0
'                          保証委託料割引対応
'
'        UPDATE          : 2011/07/20
'        UPDATER         : M.RYU
'        Ver             : 3.1
'                          統合ヘルパーに基本入力設定のコードを保証委託設定コードの前に移動
'
'        UPDATE          : 2011/08/23
'        UPDATER         : M.RYU
'        Ver             : 3.2
'                          統合ヘルパー（回収区分=ｸﾚｼﾞｯﾄ）出力条件を修正
'
'        UPDATE          : 2011/12/21
'        UPDATER         : M.RYU
'        Ver             : 3.3
'                          統合ヘルパー 毎月の支払について　の出力について修正　オフィスは旧hpler　8、Ｈは新helper
'
'        UPDATE          : 2012/06/12
'        UPDATER         : M.HONDA
'        Ver             : 3.4
'                          コンテナ（ﾊﾞｲｸ）対応
'
'        UPDATE          : 2012/06/18
'        UPDATER         : tajima
'        Ver             : 3.5
'                          空室待予約対応
'
'        UPDATE          : 2013/03/30
'        UPDATER         : M.HONDA
'        Ver             : 3.6
'                          割引適用期間を追加
'
'        UPDATE          : 2013/04/09
'        UPDATER         : M.HONDA
'        Ver             : 3.7
'                          契約案内シートに行を挿入したため契約案内シート行削除位置を修正
'
'        UPDATE          : 2013/04/14
'        UPDATER         : K.ISHIZAKA
'        Ver             : 3.8
'                          保証委託料加瀬負担に対応させる
'                          ただし請求書上は初月使用料、雑費を割引し委託料は請求するように見せる
'
'        UPDATE          : 2013/04/22
'        UPDATER         : K.ISHIZAKA
'        Ver             : 3.9
'                          保証委託料加瀬負担は初回使用料の割引分として出力する
'                          初回使用料、毎月雑費初回日割分、毎月追加雑費１初回日割分、毎月追加雑費２初回日割分を対象とする
'                          明細の項目名は、初回使用料などと同じ名称の後ろに貸主会社名をつけ
'                          固定値”負担により無料（移動）”をつける
'
'        UPDATE          : 2013/05/13
'        UPDATER         : K.ISHIZAKA
'        Ver             : 4.0
'                          保証委託料加瀬負担は初回保証委託料として出力する
'
'
'        UPDATE          : 2013/06/18
'        UPDATER         : M.HONDA
'        Ver             : 4.1
'                          南大塚ヤードのみ別途誓約書を出力する。
'
'        UPDATE          : 2013/10/16
'        UPDATER         : M.HONDA
'        Ver             : 4.2
'                          更新区分によりヘルパーの出力条件を変更する。
'
'        UPDATE          : 2013/11/16
'        UPDATER         : M.HONDA
'        Ver             : 4.3
'                          住所変更届書を出力できるように対応
'
'        UPDATE          : 2014/03/21
'        UPDATER         : MIYAMOTO
'        Ver             : 4.4
'                          自動鍵(ダイヤル南京錠)の対応(解除番号の取得、基本入力シートセット)
'
'
'        UPDATE          : 2014/08/02
'        UPDATER         : M.HONDA
'        Ver             : 4.5
'                          バイクの場合、鍵同封書類にバイクの誓約書を出力するように対応
'
'        UPDATE          : 2014/10/09
'        UPDATER         : M.HONDA
'        Ver             : 4.5
'                          EMAILアドレスを住所変更届へ出力するように修正。
'
'        UPDATE          : 2014/10/14
'        UPDATER         : tajima
'        Ver             : 4.6
'                          ネット契約フェーズ１Ｂ対応、TPOINT,誕生日出力対応
'
'        UPDATE          : 2015/01/30
'        UPDATER         : M.HONDA
'        Ver             : 4.7
'                          ヘルパーに項目を追加
'
'        UPDATE          : 2015/07/15
'        UPDATER         : K.ISHIZAKA
'        Ver             : 4.8
'                          PDF印刷に印刷種別を渡せるようにする
'
'        UPDATE          : 2015/07/28
'        UPDATER         : K.ISHIZAKA
'        Ver             : 4.9
'                          P_PRINT_鍵書類のときで、ＰＤＦ化するときは「明細書控え」と「明細書（経理用）」は不要
'                                                  鍵種別が「ダイヤル」の場合は「鍵について」を同時に印刷
'
'        UPDATE          : 2015/08/02
'        UPDATER         : M.HONDA
'        Ver             : 5.0
'                          ネット契約に関するロジックをコメントアウト
'
'        UPDATE          : 2016/05/19
'        UPDATER         : MIYAMOTO
'        Ver             : 5.1
'                          「鍵同封書類」ボタン押下時、レンタル用途が「バイク屋外置場」なら
'                           「バイク契約送付資料」シートを追加する
'                           また、上記条件の場合、「ドアキーの開け方」、「鍵について」は出力対象外とする
'
'        UPDATE          : 2017/12/27
'        UPDATER         : N.IMAI
'        Ver             : 5.2
'                          ヤード"111102"も約款別
'
'       UPDATE　　  　  ：2017/12/20
'       UPDATER　　　   ：EGL
'       Ver　　         ：5.2
'       　              ：新プラン対応(部門別の事務手数料、保証委託料の廃止)
'
'        UPDATE          : 2018/03/10
'        UPDATER         : N.IMAI
'        Ver             : 5.3
'                        : 解約用電話番号を部門毎に変更可能とする(CONT_MAST)
'
'        UPDATE          : 2018/03/19
'        UPDATER         : M.HONDA
'        Ver             : 5.4
'                        : 地方区分を追加
'
'        UPDATE          : 2018/06/13
'        UPDATER         : M.HONDA
'        Ver             : 5.5
'                        : 集客契約のヤードを追加
'
'        UPDATE          : 2018/08/25
'        UPDATER         : tajima
'        Ver             : 5.6
'                        : 事務手数料対応
'
'        UPDATE          : 2018/09/22
'        UPDATER         : tajima
'        Ver             : 5.7
'                          加瀬トランクサービス分社化対応
'
'        UPDATE          : 2018/10/02
'        UPDATER         : tajima
'        Ver             : 5.8
'                          階数表示対応
'
'        UPDATE          : 2019/08/15
'        UPDATER         : Y.WADA
'        Ver             : 5.9
'                          消費税対応
'
'        UPDATE          : 2019/09/19
'        UPDATER         : K.ISHIZAKA
'        Ver             : 6.0
'                          消費税対応で前払い賃料が加算されていなかった
'
'        UPDATE          : 2019/09/25
'        UPDATER         : EGL y
'        Ver             : 6.1
'                          個人情報シート、約款シートの細分化対応
'
'        UPDATE          : 2019/09/26
'        UPDATER         : K.ISHIZAKA
'        Ver             : 6.2
'                          消費税対応で課税対象額に含めないものまで加算されていた
'
'        UPDATE          : 2020/01/24
'        UPDATER         : EGL
'        Ver             : 6.3
'                          保証会社無し対応
'
'        UPDATE          : 2020/04/03
'        UPDATER         : EGL
'        Ver             : 6.4
'                          保証会社無対応の不具合、ネット割引金額の明細が出ない
'
'        UPDATE          : 2020/05/01
'        UPDATER         : EGL
'        Ver             : 6.5
'                          受付入力ヘルパー速度改善
'
'        UPDATE          : 2020/05/29
'        UPDATER         : EGL
'        Ver             : 6.6
'                          オフィス解約届変更対応
'
'        UPDATE          : 2020/06/29
'        UPDATER         : EGL
'        Ver             : 6.7
'                          1.毎月回収方法が「振込」の場合、「毎月の支払いについて」を印刷されないように修正
'                          2.トランク用開錠シート対応
'
'        UPDATE          : 2020/07/14
'        UPDATER         : EGL
'        Ver             : 6.8
'                          毎月回収方法が「振込」の場合、対応による不具合対応(暗黙の型変換に頼ってしまった問題）
'
'        UPDATE          : 2020/07/21
'        UPDATER         : EGL
'        Ver             : 6.9
'                          トランク用開錠シートの付加条件として鍵区分コード(10,15,50,94,95)の場合とする
'
'        UPDATE          : 2020/08/28
'        UPDATER         : EGL
'        Ver             : 7.0
'                          「加瀬倉庫は収納代行会社です」の出力をとりやめ
'
'        UPDATE          : 2020/10/15
'        UPDATER         : tajima
'        Ver             : 7.1
'                          オフィス「シェアオフィス」用ヘルパ対応
'
'        UPDATE          : 2020/11/13
'        UPDATER         : tajima
'        Ver             : 7.2
'                          保証会社用ヘルパ「シェアオフィス用」シート対応
'
'        UPDATE          : 2021/06/02
'        UPDATER         : EGL
'        Ver             : 7.3
'                          QR番号対応
'
'        UPDATE          : 2021/08/06
'        UPDATER         : M.HONDA
'        Ver             : 7.4
'                          物件鍵番号
'
'        UPDATE          : 2022/05/16
'        UPDATER         : N.IMAI
'        Ver             : 7.5
'                          シートを分ける「H:鍵について、8:ダイヤル開錠方法」を使用する
'
'        UPDATE          : 2022/05/17
'        UPDATER         : N.IMAI
'        Ver             : 7.6
'                          Bluetooth対応
'
'        UPDATE          : 2022/09/05
'        UPDATER         : N.IMAI
'        Ver             : 7.7
'                          申込書にあるカード、口座の〇を削除
'
'        UPDATE          : 2022/10/12
'        UPDATER         : N.IMAI
'        Ver             : 7.8
'                          申込書にあるカード、口座の〇を削除はバーチャルオフィスのみ
'
'        UPDATE          : 2022/11/28
'        UPDATER         : N.IMAI
'        Ver             : 7.9
'                          振込先口座に仮想口座を出力
'
'        UPDATE          : 2023/04/21
'        UPDATER         : N.IMAI
'        Ver             : 8.0
'                          オフィスの場合はヘルパーにクレカのシートを出力しない
'
'        UPDATE          : 2023/06/29
'        UPDATER         : N.IMAI
'        Ver             : 8.1
'                          約款、個人情報取扱シートをファイルからコピーする前にリネームする
'
'        UPDATE          : 2023/10/03
'        UPDATER         : N.IMAI
'        Ver             : 8.2
'                          角印の挿入方法を変更
'
'        UPDATE          : 2023/10/12
'        UPDATER         : N.IMAI
'        Ver             : 8.3
'                          保証委託契約書、保証会社個人情報取扱を対象シートから除外
'
'        UPDATE          : 2023/10/16
'        UPDATER         : N.IMAI
'        Ver             : 8.4
'                          解約届の場合、約款は出力しない
'
'        UPDATE          : 2023/10/23
'        UPDATER         : N.IMAI
'        Ver             : 8.5
'                          バーチャルオフィスは保証料がある
'
'        UPDATE          : 2023/12/04
'        UPDATER         : N.IMAI
'        Ver             : 8.6
'                          HELPERのパス（INTI_FILE）を部門毎にする
'
'        UPDATE          : 2025/09/05
'        UPDATER         : N.IMAI
'        Ver             : 8.7
'                          自社保証の場合は保証会社の約款を出力対象シートから除外
'
'        UPDATE          : 2025/11/22
'        UPDATER         : M.HONDA
'        Ver             : 8.8
'                          約款統一対応
'
'        UPDATE          : 2026/03/10
'        UPDATER         : K.KINEBUCHI
'        Ver             : 8.9
'                          オフィス用ヘルパー追加
'
'==============================================================================*
Option Compare Database
Option Explicit
'==============================================================================*
'   変数宣言
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'固定約款のヤード   ----2010/03/25----
Private Const CST_YARD_CODE_31102 = "031102"
Private Const CST_YAKAN_NAME_31102 = "31102約款"
Private Const CST_YARD_CODE_111102 = "111102"                                   'INSERT 2017/12/27 N.IMAI
Private Const CST_YAKAN_NAME_111102 = "111102約款"                              'INSERT 2017/12/27 N.IMAI
Private Const CST_YAKAN_NAME_V = "バーチャルオフィス約款"   '----2010/05/31--M.RYU--add--
Private Const CST_YAKAN_NAME = "約款"

'統合ﾍﾙﾊﾟｰに保証委託書・保証会社個人情報取扱シート挿入位置　----2010/03/29----
'2019/09/25 chg
Private Const CST_SHEETNO_保証委託書 = 12
Private Const CST_SHEETNO_保証会社個人情報取扱 = 10
'Private Const CST_SHEETNO_保証委託書 = 10
'Private Const CST_SHEETNO_保証会社個人情報取扱 = 12

'顧客区分
Private Const P_USER_KKBN_法人          As Integer = 1
'Private Const P_HELPER_統合             As String = "統合ヘルパー.xls"         'DELETE 2008/11/13 SHIBAZAKI

'update 2020/05/01 EGL Private Const P_HELPER_統合             As String = "統合ヘルパー$@$@.xls"
'Private Const P_HELPER_統合             As String = "統合ヘルパー$@$@.xlsm"
Private Const P_HELPER_統合             As String = "統合ヘルパー.xlsm"
Private Const P_HELPER_統合オフィス     As String = "統合ヘルパー_OFFI.xlsm"    'INSERT 2026/03/10 K.KINEBUCHI

''Private Const P_HELPER_ネット           As String = "ネット用申込ヘルパー.xls"  'INSERT 2007/08/28 K.ISHIZAKA--//--2010/05/31--M.RYU--del--
''Private Const P_HELPER_バーチャルオフィス As String = "バーチャルオフィスヘルパー.xls"  'INSERT 2008/04/01 tajima--//--2010/05/31--M.RYU--del--

'update 2020/05/01 EGLPrivate Const P_HELPER_保証会社         As String = "保証会社$@$@$@$@$@$@.xls"
Private Const P_HELPER_保証会社         As String = "保証会社$@$@$@$@$@$@.xlsm"

' 必要コントロールマスタ情報・構造体（部門毎に情報が異なる）
Public Type Type_CONT_MAST
    CONT_KAISYA         As String   '貸主会社名
    CONT_YUBINO         As String   '貸主郵便番号
    CONT_ADDR_1         As String   '貸主住所１
    CONT_ADDR_2         As String   '貸主住所２
    CONT_TEL_NO         As String   '貸主TEL
    CONT_FAX_NO         As String   '貸主FAX
    CONT_CANCEL_TEL_NO  As String   '解約専用TEL INSERT 2018/03/10 N.IMAI
    CONT_KEIUKE_TLD     As Integer  '契約受付期限日数
    CONT_CAMPC          As String   '会社コード  INSERT 2018/09/22 EGL
    CONT_TOIAWASE_MAIL  As String   '問合先メアドINSERT 2018/09/22 EGL
    CONT_SEIKYUSYO_TOUROKU_NO  As String   '適格請求書発行事業者登録番号    'INSERT 2019/08/15 Y.WADA
End Type

' 口座情報・構造体（KASE_DBから、なお部門毎に異なる）
Public Type Type_BANK_INF
    BANKT_KINYN         As String   '金融機関名
    BANKT_SHITN         As String   '支店名
    CODET_NAMEN         As String   '預金種別
    NYKOM_KOUZB         As String   '口座番号
    NYKOM_KOUZN         As String   '口座名義人
End Type

' ヘルパーメイン情報・構造体 ※Null許可項目はVariant
Public Type Type_MAIN_INF
    int予約受付状態区分     As Integer                                          'INSERT 2007/09/01 K.ISHIZAKA
    str部門コード           As String
    strヤード名             As String
    strヤードコード         As String   'ゼロサプレス済みであること
    strヤード住所           As Variant
    strスペースコード       As String   '所謂コンテナ番号、ゼロサプレス済みであること
    strスペースサイズ       As String
    var上下段               As Variant
    intフロア               As Integer
    intレンタル用途コード   As Integer
    strレンタル用途名       As String
    str鍵種別名             As String
    dat受付日               As Date
    str受付番号             As String
    val解約日               As Variant
    str受付担当者名         As String
    dat契約開始日           As Date
    dat契約受付期限日       As Date
    dat起算日               As Date
    str初期費用請求方法     As String
    val明細種別             As Variant
    val請求期間             As Variant
    str顧客郵便番号         As String
    str顧客住所1            As String
    str顧客住所2            As String
    val顧客住所3            As Variant
    str顧客フリガナ         As String
    str顧客名               As String
    int顧客区分コード       As Integer
    str顧客代表者名         As String
    str顧客コード           As String   'ゼロサプレス済みであること
    val顧客TEL              As Variant
    val顧客FAX              As Variant
    val顧客携帯             As Variant
    val顧客MAIL             As Variant
    str契約番号             As String
    val承認番号             As Variant
    lng月額使用料           As Long
    lng初回使用料           As Long
    val保証料               As Variant
    val保証料加瀬負担分     As Variant                                          'INSERT 2013/04/22 K.ISHIZAKA
    val毎月雑費名           As Variant
    val毎月雑費             As Variant
    val日割毎月雑費         As Variant
    val追加毎月雑費名1      As Variant
    val追加毎月雑費1        As Variant
    val追加日割毎月雑費1    As Variant
    val追加毎月雑費名2      As Variant
    val追加毎月雑費2        As Variant
    val追加日割毎月雑費2    As Variant
    val初回雑費名           As Variant
    val初回雑費             As Variant
    val追加初回雑費名1      As Variant
    val追加初回雑費1        As Variant
    val追加初回雑費名2      As Variant
    val追加初回雑費2        As Variant
    val書類送付方法         As Variant  '追加
    val発生区分             As Variant                                          'INSERT 2007/08/28 K.ISHIZAKA
    var鍵区分コード         As Variant                                          'INSERT 2007/11/15 SHIBAZAKI
    var入金済み金額         As Variant                                          'INSERT 2007/11/15 SHIBAZAKI
    str保証区分             As String                                           'INSERT 2008/11/13 SHIBAZAKI
    str保証会社コード       As String                                           'INSERT 2009/04/01 hirano
    var割引適用開始月       As String                                           'INSERT 2008/10/27 iizuka
    var割引有効可否         As Variant                                          'INSERT 2008/10/27 iizuka
    lng毎月値引き額         As Long                                             'INSERT 2009/02/02 Hirano
    str回収方法             As String                                           'INSERT 2010/09/02 ryu
    val保証金割引額         As Variant         'add 2011/06/18 tajima
    strスロープ             As String                                           'INSERT 2012/06/12 HONDA
    dat希望開始日           As Date            'add 2012/06/18 tajima
    lng割引適用期間         As Long                                             'INSERT 2013/03/30 M.HONDA
    str更新区分             As String                                           'INSERT 2013/10/17 M.HONDA
    val解除番号             As Variant                                          'INSERT 2014/03/21 MIYAMOTO
    val鍵番号               As Variant          'add 2020/06/29 Takenouchi
    valTPOINT番号           As Variant          'add 2014/10/14 tajima
    val顧客誕生日           As Variant          'add 2014/10/14 tajima
    val顧客性別             As Variant          'add 2014/10/14 tajima
    val予定収納物           As Variant          'add 2014/10/14 tajima
    val利用予定期間         As Variant          'add 2014/10/14 tajima
    val媒体                 As Variant          'add 2014/10/14 tajima
    val顧客代表者名カナ     As Variant          'add 2014/10/14 tajima
    intスロープ貸出コード   As Variant          'add 2014/10/14 tajima
    'INS 2015/01/30  HONDA
    val勤務先名             As Variant
    val職種                 As Variant
    val勤め先電話番号       As Variant
    val緊急連絡先氏名       As Variant
    val緊急連絡先カナ       As Variant
    val緊急連絡先続柄       As Variant
    val緊急連絡先TEL        As Variant
    val緊急連絡先携帯       As Variant
    val担当者部署           As Variant
    val担当者氏名           As Variant
    val担当者電話番号       As Variant
    val担当者携帯番号       As Variant
    int購入前可否           As Long
    val登録ナンバー         As Variant
    val車種                 As Variant
    val排気量               As Variant
    'INS 2015/01/30  HONDA
    int年払い               As Integer
    valサービス1            As Variant       '2015/10/14 M.HONDA INS
    valサービス2            As Variant       '2015/10/14 M.HONDA INS
    valサービス3            As Variant       '2015/10/14 M.HONDA INS
    valサービス期間         As Variant       '2015/10/14 M.HONDA INS
    int満了月数             As Integer       '2015/10/14 M.HONDA INS
    blnネット予約           As Boolean       '2017/12/20 EGL INS
    val地方                 As Variant       '2018/03/19 M.HONDA INS
    valネット割引額         As Variant       '2018/08/25 EGL INS
    lng階数                 As Long          '2018/10/02 EGL INS
    str会社コード           As Variant       '2019/09/25 EGL INS
    int集客契約区分         As Integer       '2019/09/25 EGL INS
    str物件鍵番号           As String        '2021/06/02 EGL INS
    val初回請求回収金額     As Variant       '2021/06/02 EGL INS
End Type

'割付情報
Private Type MSZZ039Type_DCRA_TRAN_INF
    DCRAT_DCNT_NO       As String
    VALUE               As Long
    DCRAT_TEXT          As String
    NEBIKI_TEXT         As String                                              'INSERT 2008/10/27 iizuka
End Type

Private wk割付情報()     As MSZZ045割付情報                                     'INSERT 2008/10/27 iizuka
Private pst保証料割引   As MSZZ068Type_DCRA_TRAN_INF    'add 2011/06/18 tajima

Private plngZeikomi1     As Long         'INSERT 2019/08/15 Y.WADA
Private plngZeikomi2     As Long         'INSERT 2019/08/15 Y.WADA

'↓INSERT 2007/11/15 SHIBAZAKI
'印刷種別(受付入力画面でも使うのでPublic)
Public Const P_PRINT_申込書                 As Integer = 2      '申込書
Public Const P_PRINT_解約書                 As Integer = 3      '解約書
Public Const P_PRINT_鍵書類                 As Integer = 4      '鍵同封書類

'レンタル用途(受付入力画面でも使うのでPublic)
Public Const P_USAGE_個室オフィス           As Integer = 5
Public Const P_USAGE_自習室                 As Integer = 8      'INSERT 2011/06/01 M.RYU
Public Const P_USAGE_バーチャルオフィス     As Integer = 7
Public Const P_USAGE_シェアオフィス         As Integer = 6      'add 2020/11/13 egl

' ▼20101220 M.HONDA
Public Const P_PRINT_クレジット             As Integer = 8      'クレジット
Public Const P_PRINT_口座                   As Integer = 9      '口座振替
' ▲20101220 M.HONDA

Public Const P_PRINT_住所変更               As Integer = 10      '住所変更  ''2013/11/16 M.HONDA INS

Public Const P_PRINT_QR                     As Integer = 16      'QR印刷    '2021/06/02 EGL

'鍵区分
Private Const P_KAGIICD_ドアキー            As Integer = 30
Private Const P_KAGIICD_ダイヤル            As Integer = 15                     'INSERT 2015/07/28 K.ISHIZAKA
'↑INSERT 2007/11/15 SHIBAZAKI

Private Const P_KAGIICD_QRコード            As Integer = 16     '2021/06/02 EGL INS

'↓INSERT 2008/11/13 SHIBAZAKI
'保証区分FROM-保証金区分(受付入力画面でも使うのでPublic)
Public Const P_HOSYIFROM_保証金             As Integer = 0      '統合ヘルパー00
Public Const P_HOSYIFROM_保証委託           As Integer = 1      '統合ヘルパー
Public Const P_HOSYIFROM_事務手数料         As Integer = 2      '統合ヘルパー02

Private Const P_NAMEID_保証区分             As String = "200"
'2009/04/01 DEL <S> hirano 保証会社情報別ファイル保持に伴う削除
''▼2009/02/20 ADD <S> hirano
'Private Const P_STRING_保証区分1            As String = "1" '保証区分１
'Private Const P_STRING_保証区分2            As String = "2" '保証区分２
'Private Const P_FILE_保証区分1              As String = "03" '保証区分１
'Private Const P_FILE_保証区分2              As String = "04" '保証区分２
''▲2009/02/20 ADD <E> hirano
'2009/04/01 DEL <E> hirano
Private Const P_STRING_保証委託料           As String = "初回保証委託料"
Private Const P_STRING_事務手数料           As String = "事務手数料"
'↑INSERT 2008/11/13 SHIBAZAKI
'▼ 2008/11/19 hirano Add
Public P_CMD明細             As Integer
'▲ 2008/11/19 hirano Add

'---20100109---M.RYU---add---------<s>
Private Const P_STRING_保証金               As String = "保証金"
Private Const P_SHEET_経理用明細            As String = "明細書（経理用）"
Private Const P_明細種類_明細書             As String = "明細書"
Private Const P_明細種類_請求書             As String = "請求書"
Public Const P_CELL_各印                    As String = "各印"
'---20100109---M.RYU---add---------<e>

Public Const P_QRコード                     As String = "QRコードあり" '2021/06/02 EGL INS
Public Const P_CELL_QRコード                As String = "QRコード"     '2021/06/02 EGL INS

'▼ 2020/07/14 chg 評価する値は str回収方法(String)のためStringとする
'---2010/09/02---M.RYU---add---------<s>’支払方法
'Private Const P_KAIHI_CREDIT                As Integer = "2"      'クレジット
'Private Const P_KAIHI_ACCOUNT               As Integer = "1"      '口座振替
'Private Const P_KAIHI_TRANSFER              As Integer = "4"      '振込         '2020/06/29 Takenouchi INS
'---2010/09/02---M.RYU---add---------<e>
Private Const P_KAIHI_CREDIT                As String = "2"      'クレジット
Private Const P_KAIHI_ACCOUNT               As String = "1"      '口座振替
Private Const P_KAIHI_TRANSFER              As String = "4"      '振込
'▲ 2020/07/14 chg

'==============================================================================*
'
'       MODULE_NAME     : 出力シートの取得
'       MODULE_ID       : GetOutSheets
'       CREATE_DATE     : 2007/08/20
'                       :
'       PARAM           : aMAIN_INF  - 主情報（１受付毎に異なる）
'                       : [intPrintKind]        印刷種別(I) 省略可 ※定数宣言参照
'                       :
'       NOTE            : 渡された主情報から出力対象のシートを選択
'                       :
'       RETURN          : 対象シート名の配列
'                       : 不正終了時はnullが設定されている
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function GetOutSheets(aMAIN_INF As Type_MAIN_INF, Optional intPrintKind As Integer = P_PRINT_申込書, Optional strHosyicdFrom As String = P_HOSYIFROM_保証委託, Optional strPdfPath As String = "") As Variant 'INSERT 2015/07/28 K.ISHIZAKA
    
    Dim var対象シート   As Variant
    Dim str約款     As String
    Dim str申込書   As String
    Dim str解約通知書 As String
    
    On Error GoTo Exception
    
    var対象シート = Null
    str約款 = "約款"
    str申込書 = ""
    str解約通知書 = "解約通知書"
     
    '部門"1"の場合、オフィス用解約届書を出力する
    If aMAIN_INF.str部門コード = "1" Then
        str解約通知書 = "解約通知書 (オフィス)"
    End If
    
    If intPrintKind = P_PRINT_申込書 Then
        If aMAIN_INF.int顧客区分コード = P_USER_KKBN_法人 Then
            If aMAIN_INF.val発生区分 = 0 Then
                str申込書 = "申込書法人TEL"
            Else
                str申込書 = "申込書法人NET"
            End If
        Else
            If aMAIN_INF.val発生区分 = 0 Then
                str申込書 = "申込書個人TEL"
            Else
                str申込書 = "申込書個人NET"
            End If
        End If
        'If aMAIN_INF.str部門コード = "1" Then                                                                                              'DELETE 2025/09/05 N.IMAI
        If aMAIN_INF.str部門コード = "1" Or aMAIN_INF.str保証区分 = "2" Then                                                                'INSERT 2025/09/05 N.IMAI
            var対象シート = Array("契約の案内", str約款, str申込書, "個人情報取扱", "明細書")                                               'INSERT 2023/10/12 N.IMAI
        Else
            var対象シート = Array("契約の案内", str約款, str申込書, "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書")     'DELETE 2023/10/12 N.IMAI
        End If
       
'       'レンタル用途がバイク関連ならば車両収納誓約書を同時に印刷
'        If isUsageBike(aMAIN_INF.intレンタル用途コード) = True Then
'            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
'            var対象シート(UBound(var対象シート)) = "車両収納誓約書"
'        End If
        
        
       '2025/11/22  バイクの用途で誓約書を分ける。
        If aMAIN_INF.intレンタル用途コード = 32 Then
            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
            var対象シート(UBound(var対象シート)) = "車両収納誓約書"
        End If
        If aMAIN_INF.intレンタル用途コード = 3 Or aMAIN_INF.intレンタル用途コード = 31 Or aMAIN_INF.intレンタル用途コード = 33 Or (aMAIN_INF.intレンタル用途コード = 0 And aMAIN_INF.val予定収納物 = "オートバイ") Then
            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
            var対象シート(UBound(var対象シート)) = "バイク収納_誓約書"
        End If
        
        '2025/11/22 集約契約の場合には誓約書を追加
        If aMAIN_INF.int集客契約区分 = 1 Then
           ReDim Preserve var対象シート(UBound(var対象シート) + 1)
            var対象シート(UBound(var対象シート)) = "集客契約_誓約書"
        End If
        
        
        
       '南大塚の場合には、特別に誓約書を追加
        If (isUsageBike(aMAIN_INF.intレンタル用途コード) = True Or (aMAIN_INF.intレンタル用途コード = 0 And aMAIN_INF.val予定収納物 = "オートバイ")) And aMAIN_INF.strヤードコード = "400702" Then
            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
            var対象シート(UBound(var対象シート)) = "新特約事項"
        End If
    Else
        var対象シート = fncGetOtherSheets(aMAIN_INF, intPrintKind, strPdfPath)
    End If
    
    GetOutSheets = var対象シート
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "GetOutSheets" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'       MODULE_NAME     : 申込書以外の出力シートの取得
'       MODULE_ID       : fncGetOtherSheets
'       CREATE_DATE     : 2007/11/15 SHIBAZAKI
'                       :
'       PARAM           : aMAIN_INF  - 主情報（１受付毎に異なる）
'                       : intPrintKind        印刷種別(I) ※定数宣言参照
'                       :
'       NOTE            : 渡された主情報から出力対象のシートを選択
'                       :
'       RETURN          : 対象シート名の配列
'                       : 不正終了時はnullが設定されている
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetOtherSheets(aMAIN_INF As Type_MAIN_INF, intPrintKind As Integer, Optional strPdfPath As String = "") As Variant

    Dim var対象シート   As Variant
    Dim strSheets           As String
    Dim str解約通知書   As String
    
    On Error GoTo Exception
    var対象シート = Null
    str解約通知書 = "解約通知書"
    
    '部門"1"の場合、オフィス用解約届書を出力する
    If aMAIN_INF.str部門コード = "1" Then
        str解約通知書 = "解約通知書 (オフィス)"
    End If

    Select Case intPrintKind
        Case P_PRINT_解約書
             var対象シート = Array(str解約通知書)
       Case P_PRINT_鍵書類
            ' システム種別毎に使用するシートを指定
                Select Case aMAIN_INF.str部門コード
            
               
                Case "8", "H"
                    If Nz(aMAIN_INF.val解約日) = "" Then
 
                        If Left(aMAIN_INF.str受付番号, 1) = "E" And Nz(aMAIN_INF.str回収方法) = P_KAIHI_CREDIT Then
                            var対象シート = Array("書類送付の案内", "マイページ案内")
                            
                        Else
                           var対象シート = Array("書類送付の案内", "マイページ案内")
                            
                            '毎月回収方法が「振込」と等しくない場合
                            If Nz(aMAIN_INF.str回収方法) <> P_KAIHI_TRANSFER Then
                                If aMAIN_INF.str部門コード = "1" Then
                                    ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                                    var対象シート(UBound(var対象シート)) = "毎月の支払いについて（オフィス）"
                               End If
                            End If
                        End If
                    Else
                       var対象シート = Array("書類送付の案内", "マイページ案内")
                        
                    End If
                
                Case "1"
                    var対象シート = Array("書類送付の案内", "住所変更")
                
                    
            End Select
            
            '回収方法 =2の場合「支払申込書（口座引落）」を同時に印刷
            If Nz(aMAIN_INF.str回収方法) = P_KAIHI_ACCOUNT Then
           
                If aMAIN_INF.str部門コード = "1" Then
                    ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                    var対象シート(UBound(var対象シート)) = "口振"
                    ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                    var対象シート(UBound(var対象シート)) = "受取人払（空間）"
                Else
                
                
                    If aMAIN_INF.int顧客区分コード = P_USER_KKBN_法人 Then
                        ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                        var対象シート(UBound(var対象シート)) = "口振"
                        ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                        var対象シート(UBound(var対象シート)) = "受取人払（収納）"
                    End If
                End If
            End If
            
            
            
            If Nz(aMAIN_INF.var鍵区分コード) = P_KAGIICD_ドアキー Then
                '鍵種別が「ドアキー」の場合は「ドアキーの開け方」を同時に印刷
                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = "ドアキーの開け方"

             ElseIf Nz(aMAIN_INF.val解除番号) <> "" Then

                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                If Nz(aMAIN_INF.var鍵区分コード) = P_KAGIICD_QRコード Then
                    var対象シート(UBound(var対象シート)) = "QRコードあり"
                Else
            
                    If aMAIN_INF.str部門コード = "H" Then
                        If Nz(aMAIN_INF.var鍵区分コード) = "17" Then
                            var対象シート(UBound(var対象シート)) = "鍵について(2)"
                        Else
                            var対象シート(UBound(var対象シート)) = "鍵について"
                        End If
                    Else
                        aMAIN_INF.val鍵番号 = aMAIN_INF.val解除番号
                        If Nz(aMAIN_INF.var鍵区分コード) = "98" Then
                            var対象シート(UBound(var対象シート)) = "Bluetoothあり"
                        ElseIf Nz(aMAIN_INF.var鍵区分コード) = "17" Then
                            var対象シート(UBound(var対象シート)) = "鍵について(2)"
                        Else
                            var対象シート(UBound(var対象シート)) = "ダイヤル開錠方法"
                        End If
                    End If
   
                End If
            ElseIf Nz(aMAIN_INF.val鍵番号) <> "" Then
                If aMAIN_INF.str部門コード = "8" Then
                    Select Case aMAIN_INF.var鍵区分コード
                        Case 10, 15, 50, 94, 95
                            '部門「8」の場合、トランク鍵用の開錠用シートを同時に印刷
                            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                            var対象シート(UBound(var対象シート)) = "ダイヤル開錠方法"
                        Case 17
                            aMAIN_INF.val鍵番号 = aMAIN_INF.val解除番号
                            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                            var対象シート(UBound(var対象シート)) = "鍵について(2)"
                       
                        Case 98
                            aMAIN_INF.val鍵番号 = aMAIN_INF.val解除番号
                            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                            var対象シート(UBound(var対象シート)) = "Bluetoothあり"
                        
                    End Select
                End If
        
            End If
            
            If aMAIN_INF.str部門コード = "1" Then
                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = str解約通知書
            End If
            
        
  
            '受付入力画面のコンテナ情報欄にある「スロープ貸出：01:貸出」の場合に印刷
            If aMAIN_INF.intスロープ貸出コード = 1 Then
            ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = "バイクスロープ"
            End If
            'insert end 2020/05/01 EGL
                
            'レンタル用途が「バイク屋外置場」の場合、「バイク契約送付資料」を同時に印刷
            If aMAIN_INF.intレンタル用途コード = 31 Then
                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = "バイク契約送付資料"
            End If

            
   
        Case P_PRINT_クレジット
            var対象シート = Array("クレジット")



            If aMAIN_INF.str部門コード = "1" Then
                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = "受取人払（空間）"
            Else
                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = "受取人払（収納）"
            End If
        
        Case P_PRINT_口座
            var対象シート = Array("口振")
            If aMAIN_INF.str部門コード = "1" Then
                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = "受取人払（空間）"
            Else
                ReDim Preserve var対象シート(UBound(var対象シート) + 1)
                var対象シート(UBound(var対象シート)) = "受取人払（収納）"
            End If
  

   
        Case P_PRINT_住所変更
            var対象シート = Array("住所変更")
          
            
        End Select
    
        fncGetOtherSheets = var対象シート
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncGetOtherSheets" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :CallHelperMacro
'        機能             :必要なヘルパーのマクロを呼ぶ
'        IN               :Bookオブジェクト, 主情報
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub CallHelperMacro(aBook As Object, aMAIN_INF As Type_MAIN_INF)

    Dim isRogoPrint As Boolean
    
  On Error GoTo Exception
  
    ' 各マクロへのパラメータ設定
    ' ①システムが横浜鋼管可否
    If aMAIN_INF.str部門コード = "T" Then
        isRogoPrint = False
    ' 上記以外ならTRUE
    Else
        isRogoPrint = True
    End If

    
    ' 各マクロの実行
    Call aBook.sub加瀬ロゴ印刷設定(isRogoPrint)
    
    Exit Sub
    
Exception:
    Call Err.Raise(Err.Number, "CallHelperMacro" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :SetBaseSheet
'        機能             :基本入力シートに値を設定する
'        IN               :値設定するシート、設定する各情報
'                         : [intPrintKind]        印刷種別(I) 省略可 ※定数宣言参照
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SetBaseSheet(aSheet As Object, _
                         aCONT_MAST As Type_CONT_MAST, _
                         aBANK_INF As Type_BANK_INF, _
                         aMAIN_INF As Type_MAIN_INF, _
                         Optional intPrintKind As Integer = P_PRINT_申込書, _
                         Optional strHosyicdFrom As String = P_HOSYIFROM_保証委託)
  On Error GoTo Exception
  
      With aSheet

        ' 部門コードを元にシステム種別を決める
        ' このシステム種別名はヘルパーと密接な関係を持っているので変更には注意！
        Select Case aMAIN_INF.str部門コード
              Case "1"
                   .Range("システム種別名").VALUE = "オフィス"
              Case "8"
                   .Range("システム種別名").VALUE = "トランク"
              Case "H"
                   .Range("システム種別名").VALUE = "コンテナ"
              Case Else
                  '上記以外の部門コードは例外終了する
                    Call MSZZ024_M10("OutPutHelperApplication", "部門コード[" & aMAIN_INF.str部門コード & "]は統合ヘルパーには対応していません。")
        End Select
      
          ' 印刷種別を設定する
        ' この印刷種別名はヘルパーの動作に関係するので変更には注意！
        Select Case intPrintKind
            Case P_PRINT_申込書
                .Range("印刷種別").VALUE = "申込書"
            Case P_PRINT_解約書
                .Range("印刷種別").VALUE = "解約書"
            Case P_PRINT_鍵書類
                .Range("印刷種別").VALUE = "鍵同封書"
            Case P_PRINT_QR 'INSERT 2021/06/02 EGL
                .Range("印刷種別").VALUE = "物件解鍵用QR" 'INSERT 2021/06/02 EGL
            ' ▼ 20121221 M.HONDA
            Case P_PRINT_クレジット
                .Range("印刷種別").VALUE = "クレジット"
            Case P_PRINT_口座
                .Range("印刷種別").VALUE = "口座振替"
            ' ▲ 20121221 M.HONDA
            Case P_PRINT_住所変更
                .Range("印刷種別").VALUE = "住所変更"
            Case Else
                '上記以外の印刷種別は例外終了する
                  Call MSZZ024_M10("subSetIntegrationBook", "印刷種別[" & intPrintKind & "]は統合ヘルパーには対応していません。")
        End Select
        '↑INSERT 2007/11/15 SHIBAZAKI
    
        '★コントロールマスタの値を設定
        .Range("貸主会社名").VALUE = aCONT_MAST.CONT_KAISYA
        .Range("貸主郵便番号").VALUE = aCONT_MAST.CONT_YUBINO
        .Range("貸主住所１").VALUE = aCONT_MAST.CONT_ADDR_1
        .Range("貸主住所２").VALUE = aCONT_MAST.CONT_ADDR_2
        .Range("貸主TEL").VALUE = aCONT_MAST.CONT_TEL_NO
        .Range("貸主FAX").VALUE = aCONT_MAST.CONT_FAX_NO
        .Range("契約受付期限日数").VALUE = aCONT_MAST.CONT_KEIUKE_TLD
        .Range("解約専用TEL").VALUE = aCONT_MAST.CONT_CANCEL_TEL_NO             'INSERT 2018/03/10 N.IMAI
        .Range("貸主メールアドレス").VALUE = aCONT_MAST.CONT_TOIAWASE_MAIL      'INSERT 2018/09/22 EGL
        .Range("貸主登録番号").VALUE = aCONT_MAST.CONT_SEIKYUSYO_TOUROKU_NO     'INSERT 2019/08/15 Y.WADA
    
        '★口座情報の値を設定
        .Range("金融機関名") = aBANK_INF.BANKT_KINYN
        .Range("金融機関支店名") = aBANK_INF.BANKT_SHITN
        .Range("預金項目") = aBANK_INF.CODET_NAMEN
        .Range("口座番号") = aBANK_INF.NYKOM_KOUZB
        .Range("口座名義") = aBANK_INF.NYKOM_KOUZN
        
        '★メイン情報を設定
        .Range("ヤード名").VALUE = aMAIN_INF.strヤード名
        .Range("ヤードコード").VALUE = aMAIN_INF.strヤードコード
        .Range("ヤード住所").VALUE = aMAIN_INF.strヤード住所
        .Range("スペースコード").VALUE = aMAIN_INF.strスペースコード
        .Range("スペースサイズ").VALUE = aMAIN_INF.strスペースサイズ
        ' レンタル用途が コンテナのときのみの上下段を設定する
        If aMAIN_INF.intレンタル用途コード = 0 Then
            .Range("上下段").VALUE = Nz(aMAIN_INF.var上下段, "")
        Else
            .Range("上下段").VALUE = ""
        End If
        '▼2018/10/02 EGL add
        '部門H以外は階数をセット
        If aMAIN_INF.str部門コード <> "H" Then
            .Range("上下段").VALUE = fncGetKaiName(aMAIN_INF.lng階数)
        End If
        '▲2018/10/02 EGL add
        
       If isUsageBike(aMAIN_INF.intレンタル用途コード) = True Or (aMAIN_INF.intレンタル用途コード = 0 And aMAIN_INF.val予定収納物 = "オートバイ") Then
            .Range("レンタル用途").VALUE = "バイク"
        Else
            .Range("レンタル用途").VALUE = aMAIN_INF.strレンタル用途名
        End If
        
        .Range("鍵種別").VALUE = aMAIN_INF.str鍵種別名
        .Range("受付日").VALUE = Format$(aMAIN_INF.dat受付日, "yyyy/mm/dd")
        .Range("受付番号").VALUE = aMAIN_INF.str受付番号
        .Range("解約日").VALUE = Format$(aMAIN_INF.val解約日, "yyyy/mm/dd")
        .Range("受付担当者名").VALUE = aMAIN_INF.str受付担当者名
        .Range("契約開始日").VALUE = Format$(aMAIN_INF.dat契約開始日, "yyyy/mm/dd")
        .Range("契約受付期限日").VALUE = Format$(aMAIN_INF.dat契約受付期限日, "yyyy/mm/dd")
       .Range("起算加算日数").VALUE = "4"
        .Range("起算日").VALUE = Format$(aMAIN_INF.dat起算日, "yyyy/mm/dd")
        .Range("初期費用請求方法").VALUE = aMAIN_INF.str初期費用請求方法
        If Nz(aMAIN_INF.val明細種別, "") = "" Then
            .Range("明細種別").VALUE = "明細書"
        Else
            .Range("明細種別").VALUE = aMAIN_INF.val明細種別
        End If
        If Nz(aMAIN_INF.val請求期間) = "" Then
            '請求期間が空ならば、起算日を格納する
            .Range("支払済月日").VALUE = Format$(aMAIN_INF.dat起算日, "yyyy/mm/dd")
        Else
            .Range("支払済月日").VALUE = Left$(aMAIN_INF.val請求期間, 4) & "/" & Right$(aMAIN_INF.val請求期間, 2) & "/01"
        End If
        .Range("契約者郵便番号").VALUE = aMAIN_INF.str顧客郵便番号
        .Range("契約者住所１").VALUE = aMAIN_INF.str顧客住所1
        .Range("契約者住所２").VALUE = aMAIN_INF.str顧客住所2 & "　" & Nz(aMAIN_INF.val顧客住所3, "")
        .Range("契約者名フリガナ").VALUE = aMAIN_INF.str顧客フリガナ
        .Range("契約者名").VALUE = aMAIN_INF.str顧客名
        '顧客敬称 法人なら｢御中」　個人なら「様」
        If aMAIN_INF.int顧客区分コード = 1 Then
            .Range("契約者敬称").VALUE = "御中"
        Else
            .Range("契約者敬称").VALUE = "様"
        End If
        .Range("契約代表者名").VALUE = aMAIN_INF.str顧客代表者名
        .Range("顧客コード").VALUE = aMAIN_INF.str顧客コード
        .Range("契約者TEL").VALUE = Nz(aMAIN_INF.val顧客TEL, "")
        .Range("契約者FAX").VALUE = Nz(aMAIN_INF.val顧客FAX, "")
        .Range("契約者CEL").VALUE = Nz(aMAIN_INF.val顧客携帯, "")
        .Range("契約番号").VALUE = aMAIN_INF.str契約番号
        .Range("承認番号").VALUE = Nz(aMAIN_INF.val承認番号, "")
        If Nz(aMAIN_INF.val請求期間) = "" Or Nz(aMAIN_INF.val請求期間) = "" Or aMAIN_INF.lng毎月値引き額 = 0 Then
            .Range("月額使用料").VALUE = aMAIN_INF.lng月額使用料
        Else
            .Range("月額使用料").VALUE = Nz(aMAIN_INF.lng月額使用料, 0) + (aMAIN_INF.lng毎月値引き額)
        End If
       .Range("初回使用料").VALUE = aMAIN_INF.lng初回使用料
        .Range("毎月雑費名１").VALUE = aMAIN_INF.val毎月雑費名
        .Range("毎月雑費１").VALUE = aMAIN_INF.val毎月雑費
        .Range("毎月雑費名２").VALUE = aMAIN_INF.val追加毎月雑費名1
        .Range("毎月雑費２").VALUE = aMAIN_INF.val追加毎月雑費1
        .Range("毎月雑費名３").VALUE = aMAIN_INF.val追加毎月雑費名2
        .Range("毎月雑費３").VALUE = aMAIN_INF.val追加毎月雑費2
        .Range("初回雑費名１").VALUE = aMAIN_INF.val初回雑費名
        .Range("初回雑費１").VALUE = aMAIN_INF.val初回雑費
        .Range("初回雑費名２").VALUE = aMAIN_INF.val追加初回雑費名1
        .Range("初回雑費２").VALUE = aMAIN_INF.val追加初回雑費1
        .Range("初回雑費名３").VALUE = aMAIN_INF.val追加初回雑費名2
        .Range("初回雑費３").VALUE = aMAIN_INF.val追加初回雑費2
        
        .Range("初回保証料").VALUE = aMAIN_INF.val保証料 + Nz(aMAIN_INF.val保証料加瀬負担分, 0) 'INSERT 2013/05/13 K.ISHIZAKA
              
        If P_CMD明細 <> 1 And strHosyicdFrom = P_HOSYIFROM_保証委託 Then
            .Range("保証会社CD").VALUE = aMAIN_INF.str保証会社コード
        End If
        
        If strHosyicdFrom <> P_HOSYIFROM_事務手数料 And Trim(Nz(aMAIN_INF.str回収方法, "")) <> "" Then
            .Range("回収方法").VALUE = aMAIN_INF.str回収方法
        End If
        
        .Range("スロープ貸出").VALUE = aMAIN_INF.strスロープ        'INSERT 2012/06/12 HONDA
        .Range("割引適用期間").VALUE = aMAIN_INF.lng割引適用期間    'INS 2013/03/30 M.HONDA
        
        If aMAIN_INF.str部門コード = "8" Then
            If aMAIN_INF.var鍵区分コード = "16" Then                '2021/06/02 EGL INS
                .Range("解除番号").VALUE = aMAIN_INF.val解除番号    '2021/06/02 EGL INS
            Else                                                    '2021/06/02 EGL INS
  
                If Nz(aMAIN_INF.val鍵番号) <> "" Then
                    .Range("解除番号").VALUE = aMAIN_INF.val鍵番号
                Else
                    .Range("解除番号").VALUE = aMAIN_INF.val解除番号
                End If

            End If                                                  '2021/06/02 EGL INS
        Else
            .Range("解除番号").VALUE = aMAIN_INF.val解除番号
        End If
       
       If aMAIN_INF.str物件鍵番号 <> "" Then
            .Range("解除番号QR").VALUE = aMAIN_INF.str物件鍵番号        '2021/06/02 EGL INS
        End If
        
        
        .Range("契約者EMAIL").VALUE = aMAIN_INF.val顧客MAIL         '2014/10/09 M.HONDA INS
        
        .Range("契約者誕生日").VALUE = aMAIN_INF.val顧客誕生日
        .Range("Ｔポイント番号").VALUE = aMAIN_INF.valTPOINT番号
        .Range("契約者性別").VALUE = aMAIN_INF.val顧客性別
        .Range("予定収納物").VALUE = aMAIN_INF.val予定収納物
       .Range("利用予定期間").VALUE = "-"
        
        .Range("媒体").VALUE = aMAIN_INF.val媒体
        .Range("契約代表者名フリガナ").VALUE = aMAIN_INF.val顧客代表者名カナ
        .Range("職業").VALUE = aMAIN_INF.val職種
        .Range("勤務先").VALUE = aMAIN_INF.val勤務先名
        .Range("勤務先電話番号").VALUE = aMAIN_INF.val勤め先電話番号
        .Range("緊急連絡先氏名").VALUE = aMAIN_INF.val緊急連絡先氏名
        .Range("緊急連絡先フリガナ").VALUE = aMAIN_INF.val緊急連絡先カナ
        .Range("緊急連絡先電話番号").VALUE = aMAIN_INF.val緊急連絡先TEL
        .Range("続柄").VALUE = aMAIN_INF.val緊急連絡先続柄
        .Range("担当者部署").VALUE = aMAIN_INF.val担当者部署
        .Range("担当者氏名").VALUE = aMAIN_INF.val担当者氏名
        .Range("担当者電話番号").VALUE = aMAIN_INF.val担当者電話番号
        .Range("サービス1").VALUE = aMAIN_INF.valサービス1
        .Range("サービス2").VALUE = aMAIN_INF.valサービス2
        .Range("サービス3").VALUE = aMAIN_INF.valサービス3
        .Range("サービス期間").VALUE = aMAIN_INF.valサービス期間
        .Range("満了月数").VALUE = aMAIN_INF.int満了月数
        
        .Range("地方").VALUE = aMAIN_INF.val地方   '2018/03/19 M.HONDA INS
        If aMAIN_INF.str部門コード = "H" Then
            .Range("補足文言") = ""
        Else
            .Range("補足文言") = "" '2020/08/27 chg
        End If
        
         .Range("上下段").VALUE = aMAIN_INF.intフロア
        
        
    End With

    '請求明細設定
    Call setSeikyuMeisai(aSheet, aCONT_MAST, aMAIN_INF, strHosyicdFrom)         'INSERT 2013/04/22 K.ISHIZAKA
    
    Exit Sub
    
Exception:
    Call Err.Raise(Err.Number, "SetBaseSheet" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :setSeikyuMeisai
'        機能             :申込書　請求明細設定
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2008/11/13 SHIBAZAKI
'Private Sub setSeikyuMeisai(aSheet As Object, aMAIN_INF As Type_MAIN_INF)
'Private Sub setSeikyuMeisai(aSheet As Object, aMAIN_INF As Type_MAIN_INF, strHosyicdFrom As String) 'DELETE 2013/04/22 K.ISHIZAKA
Private Sub setSeikyuMeisai(aSheet As Object, aCONT_MAST As Type_CONT_MAST, aMAIN_INF As Type_MAIN_INF, strHosyicdFrom As String) 'INSERT 2013/04/22 K.ISHIZAKA

    Dim intCount    As Integer
    Dim strTemp     As String
    Dim int日割日数 As Integer
    Dim lngSECUKG_KASE      As Long                                             'INSERT 2013/04/22 K.ISHIZAKA
    
  On Error GoTo Exception
  
    '明細書に日割日数を印字する
    Dim strDaily    As String
    intCount = 1
    int日割日数 = getDailyRate(aMAIN_INF.dat起算日, aMAIN_INF.dat契約開始日)
    
'▼INSERT 2008/10/27 iizuka
    '割引き機能追加
    '割引料金取得（MSZZ045_fncGetNebiki）
    Dim int適用期間                 As Integer
    Dim int割引適用期間             As Integer
    Dim date割引適用開始日付        As Date
    Dim str割引取得月               As String
    Dim int初月カウンタ             As Integer
    Dim int初月以降カウンタ         As Integer
    Dim int初月以降サマリカウンタ   As Integer
    Dim int取得月カウンタ           As Integer
    Dim str割引変換年月             As String
    Dim str割引開始年月             As Date
    Dim str割引終了年月             As Date
    Dim int割引対象月               As Integer
    Dim int割引実行                 As Integer
    
    '各月の割引情報取得
    Dim Type_初月以降サマリ()       As MSZZ039Type_DCRA_TRAN_INF
    Dim int初月以降サマリ件数()     As Integer
    Dim isサマリ                    As Boolean
    Dim intサマリindex              As Integer
    Dim int初月以降明細             As Integer
    Dim int初月以降開始             As Integer              '2008/12/17 INSERT iizuka
    Dim int年払値引き               As Integer

    Dim blnNetNew                   As Boolean  '2017/12/20 ADD
    
    intサマリindex = -1
    int初月以降明細 = 0
    int割引実行 = 0
    
    'INSERT 2019/08/15 Y.WADA Start
    '税込合計の初期化
    plngZeikomi1 = 0
    plngZeikomi2 = 0
    'INSERT 2019/08/15 Y.WADA End
    
    
    date割引適用開始日付 = CDate(Left$(aMAIN_INF.var割引適用開始月, 4) & "/" & Right$(aMAIN_INF.var割引適用開始月, 2) & "/01")
    int適用期間 = UBound(wk割付情報)
    
    For int割引適用期間 = 0 To (int適用期間)
        If wk割付情報(int割引適用期間).件数 <> 0 Then
            If int割引実行 = 0 Then
                int割引実行 = 1
                int割引対象月 = int割引適用期間
            End If
        End If
    Next int割引適用期間
    
'INSERT 2008/12/17 iizuka start
    If Format(date割引適用開始日付, "yyyy/mm") = Format(aMAIN_INF.dat起算日, "yyyy/mm") Then
        '初月の割引情報がある場合は「wk割付情報」を１番目から取得
        '初月の割引情報がない場合は「wk割付情報」を０番目から取得
        int初月以降開始 = IIf(Format(date割引適用開始日付, "yyyy/mm") = _
                                Format(aMAIN_INF.dat起算日, "yyyy/mm"), 1, 0)
    End If
'INSERT 2008/12/17 iizuka end
    If int割引実行 = 1 Then
'       For int初月以降カウンタ = 1 To int適用期間                          '2008/12/17 DELETE iizuka
        For int初月以降カウンタ = int初月以降開始 To int適用期間            '2008/12/17 CHANGE iizuka  For int初月以降カウンタ = 1 To int適用期間
             For int取得月カウンタ = 0 To (wk割付情報(int初月以降カウンタ).件数 - 1)
                '集計wk割付情報
                isサマリ = False
                For int初月以降サマリカウンタ = 0 To intサマリindex
                    If Type_初月以降サマリ(int初月以降サマリカウンタ).DCRAT_DCNT_NO = _
                                         wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).DCRAT_DCNT_NO Then
                        '割引番号が同じ場合は集計
'DELETE 2008/11/27 K.KINEBUCHI start
'                        If wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).DCNTM_TYPE = "1" Then
'                            '割引形式が円指定の場合は、DCRAT_PRICEを利用する
'                            Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE = Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE _
'                                                                                  + wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).DCRAT_PRICE
'                        Else
'                            Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE = Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE _
'                                                                                  + wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).VALUE
'                        End If
'DELETE 2008/11/27 K.KINEBUCHI end
'INSERT 2008/11/27 K.KINEBUCHI start
                        Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE = Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE _
                                                                              + wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).VALUE
'INSERT 2008/11/27 K.KINEBUCHI end
                        int初月以降サマリ件数(int初月以降サマリカウンタ) = int初月以降サマリ件数(int初月以降サマリカウンタ) + 1
                        isサマリ = True
                        str割引変換年月 = Format(DateAdd("m", int初月以降カウンタ - int割引対象月, date割引適用開始日付), "yyyymm")
                        str割引終了年月 = CDate(Left$(str割引変換年月, 4) & "/" & Right$(str割引変換年月, 2) & "/01")
                        strTemp = Format(DateSerial(Format(str割引開始年月, "yyyy"), Format(str割引開始年月, "m"), 1), "yyyy年m月")
                        
                        If int初月以降サマリ件数(int初月以降サマリカウンタ) >= 2 Then
                            strTemp = strTemp & "～" & _
                                        Format(str割引終了年月, "yyyy年m月") & _
                                        "(" & int初月以降サマリ件数(int初月以降サマリカウンタ) & "ヶ月分)"
                        Else
                            strTemp = strTemp & "分"
                        End If
                        Type_初月以降サマリ(int初月以降サマリカウンタ).NEBIKI_TEXT = strTemp & "使用料"
                    End If
                Next int初月以降サマリカウンタ
                If Not isサマリ Then
                    '集計できない場合は、リストに追加
                    
                    intサマリindex = intサマリindex + 1
                    ReDim Preserve Type_初月以降サマリ(intサマリindex)
                    ReDim Preserve int初月以降サマリ件数(intサマリindex)
                    
'DELETE 2008/11/27 K.KINEBUCHI start
'                    If wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).DCNTM_TYPE = "1" Then
'                        '割引形式が円指定の場合は、DCRAT_PRICEを利用する
'                        Type_初月以降サマリ(intサマリindex).VALUE = wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).DCRAT_PRICE
'                    Else
'                        Type_初月以降サマリ(intサマリindex).VALUE = wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).VALUE
'                    End If
'DELETE 2008/11/27 K.KINEBUCHI end
                    Type_初月以降サマリ(intサマリindex).VALUE = wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).VALUE      'INSERT 2008/11/27 K.KINEBUCHI
                    Type_初月以降サマリ(intサマリindex).DCRAT_TEXT = wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).DCRAT_TEXT
                    Type_初月以降サマリ(intサマリindex).DCRAT_DCNT_NO = wk割付情報(int初月以降カウンタ).TRAN_DATA(int取得月カウンタ).DCRAT_DCNT_NO
                    int初月以降サマリ件数(intサマリindex) = 1
                    str割引変換年月 = Format(DateAdd("m", int初月以降カウンタ - int割引対象月, date割引適用開始日付), "yyyymm")
                    str割引開始年月 = CDate(Left$(str割引変換年月, 4) & "/" & Right$(str割引変換年月, 2) & "/01")
                    strTemp = Format(DateSerial(Format(str割引開始年月, "yyyy"), Format(str割引開始年月, "m"), 1), "yyyy年m月")
                    Type_初月以降サマリ(int初月以降サマリカウンタ).NEBIKI_TEXT = strTemp & "分使用料"
                End If
            Next int取得月カウンタ
        Next int初月以降カウンタ
    End If
'▲INSERT 2008/10/27 iizuka
    
    With aSheet
    
        '2017/12/20 ADD EGL ↓↓↓
        If aMAIN_INF.val発生区分 <> 0 Then
            'ネット契約
            If Form_FVS400.FVS400_fncChkHoshoNewOld(aMAIN_INF.str部門コード, aMAIN_INF.str保証会社コード) = True Then
                '新保証会社
                blnNetNew = True
            End If
        End If
        '2017/12/20 ADD EGL ↑↑↑
        
        strTemp = Format$(aMAIN_INF.dat起算日, "m") & "月分"
        '明細書に日割日数を印字する
        strDaily = Format(aMAIN_INF.dat起算日, " ( m月d日 より ")
        If int日割日数 = 0 Then
            strDaily = strDaily & "１ヶ月分 )"  '０なら日割りはないので
        Else
            strDaily = strDaily & _
                       Day(DateSerial(Year(aMAIN_INF.dat起算日), Month(aMAIN_INF.dat起算日) + 1, 0)) & _
                       "日割り " & _
                       int日割日数 & _
                       "日分 )"
        End If
        
        '初月使用料
        .Range("明細1") = strTemp & "使用料" & strDaily
        .Range("明細金額1") = aMAIN_INF.lng初回使用料
        plngZeikomi1 = plngZeikomi1 + Nz(aMAIN_INF.lng初回使用料, 0)    'INSERT 2019/08/15 Y.WADA
        'INSERT 2023/10/23 N.IMAI Start
        If aMAIN_INF.intレンタル用途コード = 7 Then
            plngZeikomi1 = plngZeikomi1 + Nz(aMAIN_INF.val保証料)
        End If
        'INSERT 2023/10/23 N.IMAI End
        '初回雑費
        Call subSetZappi(aSheet, intCount, Nz(aMAIN_INF.val初回雑費名), Nz(aMAIN_INF.val初回雑費), "明細", "明細金額")
        '追加初回雑費１
        '2018/08/25 ADD EGL ↓↓↓
        Call subSetZappi(aSheet, intCount, Nz(aMAIN_INF.val追加初回雑費名1), Nz(aMAIN_INF.val追加初回雑費1), "明細", "明細金額")
        ' 雑費に各適性値が入っているはずなので以下ロジックは削除、元に↑戻す
        '2017/12/20 ADD EGL ↓↓↓
        'If blnNetNew = False Then
        '    '旧(事務手数料)
        '    Call subSetZappi(aSheet, intcount, Nz(aMAIN_INF.val追加初回雑費名1), Nz(aMAIN_INF.val追加初回雑費1), "明細", "明細金額")
        'Else
        '    '新(1カ月分使用料)
        '    If aMAIN_INF.str部門コード = "H" Then
        '        'ｺﾝﾃﾅ
        '        Call subSetZappi(aSheet, intcount, Nz(aMAIN_INF.val追加初回雑費名1), Nz(aMAIN_INF.lng月額使用料), "明細", "明細金額")
        '    ElseIf aMAIN_INF.str部門コード = "8" Then
        '        'トランク
        '        Call subSetZappi(aSheet, intcount, Nz(aMAIN_INF.val追加初回雑費名1), 5400, "明細", "明細金額")
        '    Else
        '        'オフィス
        '       Call subSetZappi(aSheet, intcount, Nz(aMAIN_INF.val追加初回雑費名1), Nz(aMAIN_INF.val追加初回雑費1), "明細", "明細金額")
        '    End If
        'End If
        '2017/12/20 ADD EGL ↑↑↑
        '2018/08/25 ADD EGL ↑↑↑
        
        '追加初回雑費２
        Call subSetZappi(aSheet, intCount, Nz(aMAIN_INF.val追加初回雑費名2), Nz(aMAIN_INF.val追加初回雑費2), "明細", "明細金額")
    
        '日割毎月雑費
'        Call subSetZappi(aSheet, intcount, strTemp & Nz(aMAIN_INF.val毎月雑費名) & strDaily, Nz(aMAIN_INF.val日割毎月雑費), "明細", "明細金額")                    'DELETE 2019/08/15 Y.WADA
        Call subSetZappi(aSheet, intCount, strTemp & Nz(aMAIN_INF.val毎月雑費名) & strDaily, Nz(aMAIN_INF.val日割毎月雑費), "明細", "明細金額", False)              'INSERT 2019/08/15 Y.WADA
        '日割追加毎月雑費１
'        Call subSetZappi(aSheet, intcount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名1) & strDaily, Nz(aMAIN_INF.val追加日割毎月雑費1), "明細", "明細金額")          'DELETE 2019/08/15 Y.WADA
        Call subSetZappi(aSheet, intCount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名1) & strDaily, Nz(aMAIN_INF.val追加日割毎月雑費1), "明細", "明細金額", False)    'INSERT 2019/08/15 Y.WADA
        '日割追加毎月雑費２
        'Call subSetZappi(aSheet, intCount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名2) & strDaily, Nz(aMAIN_INF.val追加日割毎月雑費2), "明細", "明細金額")
'        Call subSetZappi(aSheet, intcount, Nz(aMAIN_INF.val追加毎月雑費名2), Nz(aMAIN_INF.val追加日割毎月雑費2), "明細", "明細金額")                               'DELETE 2019/08/15 Y.WADA
        Call subSetZappi(aSheet, intCount, Nz(aMAIN_INF.val追加毎月雑費名2), Nz(aMAIN_INF.val追加日割毎月雑費2), "明細", "明細金額", False)                         'INSERT 2019/08/15 Y.WADA
        
        '▼ 2007/09/17 tajima 請求期間（前払い分）対応
        ' まずは前払月数を求める
        Dim int前払月数 As Integer
        Dim dat請求範囲年月 As Date
        
        If Nz(aMAIN_INF.val請求期間, "") = "" Then
            int前払月数 = 0
        Else
            dat請求範囲年月 = CDate(Left$(aMAIN_INF.val請求期間, 4) & "/" & Right$(aMAIN_INF.val請求期間, 2) & "/01")
            int前払月数 = IIf(Format(dat請求範囲年月, "yyyy/mm") = Format(aMAIN_INF.dat起算日, "yyyy/mm"), 0, DateDiff("m", aMAIN_INF.dat起算日, dat請求範囲年月))
        End If
        
        ' 前払月数があれば請求期間対応
        If int前払月数 > 0 Then
            strTemp = Format(DateSerial(Format(aMAIN_INF.dat起算日, "yyyy"), Format(aMAIN_INF.dat起算日, "m") + 1, 1), "yyyy年m月")
            ' 前払い月数が一ヶ月の場合に"yyyy年mm月～yyyy年mm月"と明細書に出さないようにする
            If int前払月数 >= 2 Then
                strTemp = strTemp & "～" & _
                            Format(dat請求範囲年月, "yyyy年m月") & _
                            "(" & int前払月数 & "ヶ月分)"
            Else
                strTemp = strTemp & "分"
            End If

            intCount = intCount + 1
            .Range("明細" & intCount) = strTemp & "使用料"
            int初月以降明細 = intCount                                           'INSERT 2008/10/27 iizuka
            .Range("明細金額" & intCount) = aMAIN_INF.lng月額使用料 * int前払月数
            plngZeikomi2 = plngZeikomi2 + aMAIN_INF.lng月額使用料 * int前払月数  'INSERT 2019/09/19 K.ISHIZAKA
            '毎月雑費
'            Call subSetZappi(aSheet, intcount, strTemp & Nz(aMAIN_INF.val毎月雑費名), Nz(aMAIN_INF.val毎月雑費, 0) * int前払月数, "明細", "明細金額")                  'DELETE 2019/08/15 Y.WADA
            Call subSetZappi(aSheet, intCount, strTemp & Nz(aMAIN_INF.val毎月雑費名), Nz(aMAIN_INF.val毎月雑費, 0) * int前払月数, "明細", "明細金額", False)            'INSERT 2019/08/15 Y.WADA
            '追加毎月雑費１
'            Call subSetZappi(aSheet, intcount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名1), Nz(aMAIN_INF.val追加毎月雑費1, 0) * int前払月数, "明細", "明細金額")        'DELETE 2019/08/15 Y.WADA
            Call subSetZappi(aSheet, intCount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名1), Nz(aMAIN_INF.val追加毎月雑費1, 0) * int前払月数, "明細", "明細金額", False)  'INSERT 2019/08/15 Y.WADA
            '追加毎月雑費２
            '2007/11/15 SHIBAZAKI 「val追加初回雑費2」を「val追加毎月雑費2」に修正
            'Call subSetZappi(aSheet, intCount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名2), Nz(aMAIN_INF.val追加初回雑費2, 0) * int前払月数, "明細", "明細金額")
'            Call subSetZappi(aSheet, intcount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名2), Nz(aMAIN_INF.val追加毎月雑費2, 0) * int前払月数, "明細", "明細金額")        'DELETE 2019/08/15 Y.WADA
            Call subSetZappi(aSheet, intCount, strTemp & Nz(aMAIN_INF.val追加毎月雑費名2), Nz(aMAIN_INF.val追加毎月雑費2, 0) * int前払月数, "明細", "明細金額", False)  'INSERT 2019/08/15 Y.WADA
        End If

        '保証委託料
'        strTemp = IIf(txtBillSecukgLabel = P_LABEL_保証委託料, "初回" & txtBillSecukgLabel, txtBillSecukgLabel)
        'strTemp = "初回保証委託料"     'DELETE 2008/09/25 SHIBAZAKI
        'strTemp = "事務手数料"          'INSERT 2008/09/25 SHIBAZAKI->DELETE 2008/11/13 SHIBAZAKI
        'INSERT 2008/11/13 SHIBAZAKI
        
        '----20100109----M.RYU----upadate-----'保証会社なしの場合、”保証金”で表示----------<s>
        'strTemp = IIf(strHosyicdFrom = P_HOSYIFROM_保証委託, P_STRING_保証委託料, P_STRING_事務手数料)
        If IIf(strHosyicdFrom = P_HOSYIFROM_保証委託, P_STRING_保証委託料, IIf(strHosyicdFrom = P_HOSYIFROM_保証金, P_STRING_保証金, P_STRING_事務手数料)) = P_STRING_保証金 Then '2020/01/24 INSERT EGL
            If aMAIN_INF.str保証区分 = "2" Then   '2020/01/24 INSERT EGL
                strTemp = ""                    '2020/01/24 INSERT EGL
                GoTo lbl_skip:                  '2020/01/24 INSERT EGL
            End If                              '2020/01/24 INSERT EGL
        End If                                  '2020/01/24 INSERT EGL
        strTemp = IIf(strHosyicdFrom = P_HOSYIFROM_保証委託, P_STRING_保証委託料, IIf(strHosyicdFrom = P_HOSYIFROM_保証金, P_STRING_保証金, P_STRING_事務手数料))
lbl_skip:                                       '2020/01/24 INSERT EGL
        '----20100109----M.RYU----upadate-----'保証会社なしの場合、”保証金”で表示----------<e>
        
'▼2011/06/18 tajima
        ' 割引前の金額にするため以下へ修正 Call subSetZappi(aSheet, intcount, strTemp, Nz(aMAIN_INF.val保証料), "明細", "明細金額")
'        Call subSetZappi(aSheet, intCount, strTemp, Nz(aMAIN_INF.val保証料, 0) + Nz(aMAIN_INF.val保証金割引額, 0), "明細", "明細金額") 'DELETE 2013/05/13 K.ISHIZAKA
'        Call subSetZappi(aSheet, intcount, strTemp, Nz(aMAIN_INF.val保証料, 0) + Nz(aMAIN_INF.val保証金割引額, 0) _
                                                                               + Nz(aMAIN_INF.val保証料加瀬負担分, 0), "明細", "明細金額") 'DELEET 2019/09/26 K.ISHZIAKA 'INSERT 2013/05/13 K.ISHIZAKA
        Call subSetZappi(aSheet, intCount, strTemp, Nz(aMAIN_INF.val保証料, 0) + Nz(aMAIN_INF.val保証金割引額, 0) _
                                                                               + Nz(aMAIN_INF.val保証料加瀬負担分, 0), "明細", "明細金額", True, False) 'INSERT 2019/09/26 K.ISHIZAKA
        ' 保証委託料割引の有無をチェック　※実際は保証委託料割引の摘要文言を取りたい
        If pst保証料割引.DCRAT_PRICE <> 0 Then
'            Call subSetZappi(aSheet, intcount, pst保証料割引.DCRAT_TEXT, 0 - Nz(aMAIN_INF.val保証金割引額, 0), "明細", "明細金額") 'DELEET 2019/09/26 K.ISHZIAKA
            Call subSetZappi(aSheet, intCount, pst保証料割引.DCRAT_TEXT, 0 - Nz(aMAIN_INF.val保証金割引額, 0), "明細", "明細金額", True, False) 'INSERT 2019/09/26 K.ISHZIAKA
        End If
'▲2011/06/18 tajima
        'INSERT 2007/11/15 SHIBAZAKI
        '入金済み
'        Call subSetZappi(aSheet, intcount, "入金済み", 0 - Nz(aMAIN_INF.var入金済み金額, 0), "明細", "明細金額") 'DELEET 2019/09/26 K.ISHZIAKA
        Call subSetZappi(aSheet, intCount, "入金済み", 0 - Nz(aMAIN_INF.var入金済み金額, 0), "明細", "明細金額", True, False) 'INSERT 2019/09/26 K.ISHZIAKA
'▼INSERT 2008/10/27 iizuka
        '割引情報出力
        Dim is初月以降出力完了()  As Boolean
        
        If intサマリindex >= 0 Then
            ReDim is初月以降出力完了(intサマリindex)
            For int初月以降サマリカウンタ = 0 To intサマリindex
                is初月以降出力完了(int初月以降サマリカウンタ) = False
            Next int初月以降サマリカウンタ
        End If
        
'INSERT 2008/12/17 iizuka start
        If int初月以降開始 = 1 Then
            '初月の割引情報がある場合
'INSERT 2008/12/17 iizuka end

            For int初月カウンタ = 0 To (wk割付情報(0).件数 - 1)
                '初月の割引情報１件出力
'DELETE 2008/11/27 K.KINEBUCHI start
'            If wk割付情報(0).TRAN_DATA(int初月カウンタ).DCNTM_TYPE = "1" Then
'                '割引形式が円指定の場合は、DCRAT_PRICEを利用する
'                Call subSetZappi(aSheet, intCount, .Range("明細1") & wk割付情報(0).TRAN_DATA(int初月カウンタ).DCRAT_TEXT, _
'                                                             Nz(wk割付情報(0).TRAN_DATA(int初月カウンタ).DCRAT_PRICE, 0), "明細", "明細金額")
'            Else
'                Call subSetZappi(aSheet, intCount, .Range("明細1") & wk割付情報(0).TRAN_DATA(int初月カウンタ).DCRAT_TEXT, _
'                                                             Nz(wk割付情報(0).TRAN_DATA(int初月カウンタ).VALUE, 0), "明細", "明細金額")
'            End If
'DELETE 2008/11/27 K.KINEBUCHI end
'INSERT 2008/11/27 K.KINEBUCHI start
                Call subSetZappi(aSheet, intCount, .Range("明細1") & wk割付情報(0).TRAN_DATA(int初月カウンタ).DCRAT_TEXT, _
                                                             Nz(wk割付情報(0).TRAN_DATA(int初月カウンタ).VALUE, 0), "明細", "明細金額")
'INSERT 2008/11/27 K.KINEBUCHI end
                '▼2009/01/16 Add hirano 初回使用料に初月分を反映させる
                .Range("初回使用料") = .Range("初回使用料") + Nz(wk割付情報(0).TRAN_DATA(int初月カウンタ).VALUE, 0)
                '▼2009/01/16 Add hirano
                For int初月以降サマリカウンタ = 0 To intサマリindex
                    
                    If Type_初月以降サマリ(int初月以降サマリカウンタ).DCRAT_DCNT_NO = wk割付情報(0).TRAN_DATA(int初月カウンタ).DCRAT_DCNT_NO Then
                        '割引番号が同じ場合は初月以降も出力
'DELETE 2019/08/15 Y.WADA Start
'                        Call subSetZappi(aSheet, intcount, Type_初月以降サマリ(int初月以降サマリカウンタ).NEBIKI_TEXT & Type_初月以降サマリ(int初月以降サマリカウンタ).DCRAT_TEXT, _
'                                               Nz(Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE, 0), "明細", "明細金額")
'DELETE 2019/08/15 Y.WADA End
                        'INSERT 2019/08/15 Y.WADA Start
                        Call subSetZappi(aSheet, intCount, Type_初月以降サマリ(int初月以降サマリカウンタ).NEBIKI_TEXT & Type_初月以降サマリ(int初月以降サマリカウンタ).DCRAT_TEXT, _
                                               Nz(Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE, 0), "明細", "明細金額", False)
                        'INSERT 2019/08/15 Y.WADA End
                        
                        is初月以降出力完了(int初月以降サマリカウンタ) = True
                    End If
                Next int初月以降サマリカウンタ
            Next int初月カウンタ
        End If                  'INSERT 2008/12/17 iizuka

        
        '初月以降の出力が残ってる場合は全部出力
        For int初月以降サマリカウンタ = 0 To intサマリindex
            If Not is初月以降出力完了(int初月以降サマリカウンタ) Then
                
'DELETE 2019/08/15 Y.WADA Start
'                Call subSetZappi(aSheet, intcount, Type_初月以降サマリ(int初月以降サマリカウンタ).NEBIKI_TEXT & Type_初月以降サマリ(int初月以降サマリカウンタ).DCRAT_TEXT, _
'                                       Nz(Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE, 0), "明細", "明細金額")
'DELETE 2019/08/15 Y.WADA End
                'INSERT 2019/08/15 Y.WADA Start
                Call subSetZappi(aSheet, intCount, Type_初月以降サマリ(int初月以降サマリカウンタ).NEBIKI_TEXT & Type_初月以降サマリ(int初月以降サマリカウンタ).DCRAT_TEXT, _
                                       Nz(Type_初月以降サマリ(int初月以降サマリカウンタ).VALUE, 0), "明細", "明細金額", False)
                'INSERT 2019/08/15 Y.WADA End
            End If
        Next int初月以降サマリカウンタ
'▲INSERT 2008/10/27 iizuka
    
    End With
    lngSECUKG_KASE = aMAIN_INF.val保証料加瀬負担分                              'INSERT START 2013/04/22 K.ISHIZAKA
    strTemp = Format$(aMAIN_INF.dat起算日, "m") & "月分"
    '初月使用料
    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intCount, aCONT_MAST.CONT_KAISYA, _
        strTemp & "使用料" & strDaily, Nz(aMAIN_INF.lng初回使用料), lngSECUKG_KASE)
'DELETE 2019/08/15 Y.WADA Start
'    '日割毎月雑費
'    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intcount, aCONT_MAST.CONT_KAISYA, _
'        strTemp & Nz(aMAIN_INF.val毎月雑費名) & strDaily, Nz(aMAIN_INF.val日割毎月雑費, 0), lngSECUKG_KASE)
'    '日割追加毎月雑費１
'    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intcount, aCONT_MAST.CONT_KAISYA, _
'        strTemp & Nz(aMAIN_INF.val追加毎月雑費名1) & strDaily, Nz(aMAIN_INF.val追加日割毎月雑費1, 0), lngSECUKG_KASE)
'    '日割追加毎月雑費２
'    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intcount, aCONT_MAST.CONT_KAISYA, _
'        strTemp & Nz(aMAIN_INF.val追加毎月雑費名2) & strDaily, Nz(aMAIN_INF.val追加日割毎月雑費2, 0), lngSECUKG_KASE)
'DELETE 2019/08/15 Y.WADA End
'INSERT 2019/08/15 Y.WADA Start
    '日割毎月雑費
    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intCount, aCONT_MAST.CONT_KAISYA, _
        strTemp & Nz(aMAIN_INF.val毎月雑費名) & strDaily, Nz(aMAIN_INF.val日割毎月雑費, 0), lngSECUKG_KASE, False)
    '日割追加毎月雑費１
    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intCount, aCONT_MAST.CONT_KAISYA, _
        strTemp & Nz(aMAIN_INF.val追加毎月雑費名1) & strDaily, Nz(aMAIN_INF.val追加日割毎月雑費1, 0), lngSECUKG_KASE, False)
    '日割追加毎月雑費２
    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intCount, aCONT_MAST.CONT_KAISYA, _
        strTemp & Nz(aMAIN_INF.val追加毎月雑費名2) & strDaily, Nz(aMAIN_INF.val追加日割毎月雑費2, 0), lngSECUKG_KASE, False)
'INSERT 2019/08/15 Y.WADA End
    
    '余り（本来なら発生しないけど、起算日の入力ミスにより発生する）
    lngSECUKG_KASE = fncSetZappiKaseFutan(aSheet, intCount, aCONT_MAST.CONT_KAISYA, _
        "余り", lngSECUKG_KASE, lngSECUKG_KASE)                                 'INSERT END   2013/04/22 K.ISHIZAKA
       
    '2017/07/01 M.HONDA UPD
    If aMAIN_INF.str部門コード <> "H" And aMAIN_INF.int年払い = -1 Then
    'If aMAIN_INF.int年払い = -1 Then
        '2016/12/15 M.HONDA UPD
        'Call subSetZappi(aSheet, intCount, "年払いの為１ヶ月分サービス", aMAIN_INF.val保証料 * -1, "明細", "明細金額")
        
        If aMAIN_INF.lng割引適用期間 >= 12 Then
'            Call subSetZappi(aSheet, intcount, "年払いの為１ヶ月分サービス", ((wk割付情報(1).割引合計 * -1) + aMAIN_INF.val毎月雑費) * -1, "明細", "明細金額")         'DELETE 2019/08/15 Y.WADA
            Call subSetZappi(aSheet, intCount, "年払いの為１ヶ月分サービス", ((wk割付情報(1).割引合計 * -1) + aMAIN_INF.val毎月雑費) * -1, "明細", "明細金額", False)   'INSERT 2019/08/15 Y.WADA
        Else
'            Call subSetZappi(aSheet, intcount, "年払いの為１ヶ月分サービス", aMAIN_INF.val保証料 * -1, "明細", "明細金額")                                             'DELETE 2019/08/15 Y.WAD
            Call subSetZappi(aSheet, intCount, "年払いの為１ヶ月分サービス", aMAIN_INF.val保証料 * -1, "明細", "明細金額", False)                                       'INSERT 2019/08/15 Y.WADA
        End If
        '2016/12/15 M.HONDA UPD

    End If
    
    '2017/12/20 ADD EGL ↓↓↓
    'If blnNetNew = True Then '2020/04/03 DEL EGL
        '2018/08/25 chg EGL ↓↓↓
        'If aMAIN_INF.str部門コード = "H" Then
        '    'ｺﾝﾃﾅ(=事務手数料-月額使用料)
        '    Call subSetZappi(aSheet, intcount, "ネット契約割引サービス", aMAIN_INF.val追加初回雑費1 - aMAIN_INF.lng月額使用料, "明細", "明細金額")
        'ElseIf aMAIN_INF.str部門コード = "8" Then
        '    'ﾄﾗﾝｸ(=事務手数料-5400)
        '    Call subSetZappi(aSheet, intcount, "ネット契約割引サービス", aMAIN_INF.val追加初回雑費1 - 5400, "明細", "明細金額")
        ' ネット割引金額は受付トランに持つのでそれをそのまま表示とする
        If Nz(aMAIN_INF.valネット割引額, 0) <> 0 Then
            Call subSetZappi(aSheet, intCount, "ネット契約割引サービス", aMAIN_INF.valネット割引額, "明細", "明細金額")
        '2018/08/25 chg EGL ↑↑↑
        End If
    'End If '2020/04/03 DEL EGL
    '2017/12/20 ADD EGL ↑↑↑

    'INSERT 2019/08/15 Y.WADA Start
    Dim strDate1        As String
    Dim dblZeiRitu1     As Double
    Dim strRoundKbn1    As String
    Dim lngZeikomi1     As Long
    Dim lngPrice1       As Long
    Dim lngTax1         As Long
    
    Dim strDate2        As String
    Dim dblZeiRitu2     As Double
    Dim strRoundKbn2    As String
    Dim lngZeikomi2     As Long
    Dim lngPrice2       As Long
    Dim lngTax2         As Long
    
    '税１
    strDate1 = Replace(aMAIN_INF.dat起算日, "/", "")
    dblZeiRitu1 = MSZZ004_M20(strDate1, strRoundKbn1)
    lngZeikomi1 = plngZeikomi1
    Call MSZZ004_M10(lngZeikomi1, strDate1, "2", lngPrice1, lngTax1)
    
    '税２
    strDate2 = Format(DateAdd("m", 1, aMAIN_INF.dat起算日), "yyyymmdd")
    dblZeiRitu2 = MSZZ004_M20(strDate2, strRoundKbn2)
    lngZeikomi2 = plngZeikomi2
    Call MSZZ004_M10(lngZeikomi2, strDate2, "2", lngPrice2, lngTax2)
    
    With aSheet
        If dblZeiRitu1 = dblZeiRitu2 Then
            .Range("税率1").VALUE = dblZeiRitu1 * 100
            .Range("対象税込額1").VALUE = lngZeikomi1 + lngZeikomi2
            .Range("消費税1").VALUE = lngTax1 + lngTax2
        
            .Range("税率2").VALUE = ""
            .Range("対象税込額2").VALUE = ""
            .Range("消費税2").VALUE = ""
        Else
            .Range("税率1").VALUE = dblZeiRitu1 * 100
            .Range("対象税込額1").VALUE = lngZeikomi1
            .Range("消費税1").VALUE = lngTax1
        
            .Range("税率2").VALUE = dblZeiRitu2 * 100
            .Range("対象税込額2").VALUE = lngZeikomi2
            .Range("消費税2").VALUE = lngTax2
        End If
    End With
    'INSERT 2019/08/15 Y.WADA End

    Exit Sub

Exception:
   Call Err.Raise(Err.Number, "setSeikyuMeisai" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetZappi
'        機能             :申込書　雑費等設定
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'DELETE 2019/08/15 Y.WADA Start
'Private Sub subSetZappi(objSeet As Object, _
'                        ByRef intcount As Integer, _
'                        strTekiyou As String, _
'                        varCharge As Variant, _
'                        strTekiyouRange As String, _
'                        strChargeRange As String)
'DELETE 2019/08/15 Y.WADA End
'INSERT 2019/08/15 Y.WADA Start
'Private Sub subSetZappi(objSeet As Object, _
                        ByRef intcount As Integer, _
                        strTekiyou As String, _
                        varCharge As Variant, _
                        strTekiyouRange As String, _
                        strChargeRange As String, _
                        Optional blnSyogetu As Boolean = True _
                        )                                                       'DELEET 2019/09/26 K.ISHZIAKA
Private Sub subSetZappi(objSeet As Object, _
                        ByRef intCount As Integer, _
                        strTekiyou As String, _
                        varCharge As Variant, _
                        strTekiyouRange As String, _
                        strChargeRange As String, _
                        Optional blnSyogetu As Boolean = True, _
                        Optional blnKazeiTaisyou As Boolean = True _
                        )                                                       'INSERT 2019/09/26 K.ISHZIAKA
'INSERT 2019/08/15 Y.WADA End
    '----20100109----M.RYU----------update--------------------<s>
    ''    '2009/06/30 MOD <S> hirano 事務手数料ゼロを出力する
    ''    'If Nz(varCharge, 0) <> 0 Then
    ''    If Nz(varCharge, 0) <> 0 Or (Nz(strTekiyou, "") = P_STRING_事務手数料 And Nz(varCharge, 0) = 0) Then
    ''    '2009/06/30 MOD <E> hirano
    If Nz(varCharge, 0) <> 0 Or _
       Nz(strTekiyou, "") = P_STRING_事務手数料 Or _
       Nz(strTekiyou, "") = P_STRING_保証金 Then
    '----20100109----M.RYU----------update--------------------<e>
    
        intCount = intCount + 1
        If intCount <= 13 Then
            objSeet.Range(strTekiyouRange & intCount) = Nz(strTekiyou)
            objSeet.Range(strChargeRange & intCount) = varCharge
        
            If blnKazeiTaisyou Then                                             'INSERT 2019/09/26 K.ISHZIAKA
                'INSERT 2019/08/15 Y.WADA Start
                If blnSyogetu Then
                    '初月
                    plngZeikomi1 = plngZeikomi1 + Nz(varCharge, 0)
                Else
                    '翌月以降
                    plngZeikomi2 = plngZeikomi2 + Nz(varCharge, 0)
                End If
                'INSERT 2019/08/15 Y.WADA End
            End If                                                              'INSERT 2019/09/26 K.ISHZIAKA
        
        End If
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :getDailyRate
'        機能             :日割りを求める
'        IN               :
'        OUT              :0...日割無し、1<...日割日数
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getDailyRate(a起算日 As Date, a契約開始日 As Date) As Integer

    '日割り計算
    If Day(a起算日) = 1 Then
        '起算日が一日の場合日割りなし
        getDailyRate = 0
    Else
        '月の最終日－起算日の日＋１日（起算日も含めるため）
        getDailyRate = Day(DateSerial(Year(a契約開始日), Month(a契約開始日) + 1, 0)) _
                            - Day(a起算日) + 1
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :isUsageBike
'        機能             :レンタル用途がバイク関連可否判定
'        IN               :レンタル用途
'        OUT              :TRUE...バイク関連、FALSE...違う
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isUsageBike(aレンタル用途 As Integer) As Boolean
    '2009/10/02 レンタル用途33 :コンテナ（バイク） 追加
   If aレンタル用途 = 31 Or _
      aレンタル用途 = 32 Or _
      aレンタル用途 = 33 Or _
      aレンタル用途 = 3 Then
        
        isUsageBike = True
    Else
        isUsageBike = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncFirstDay
'        機能             :月初日
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncFirstDay(aVariant As Variant) As String

    If Nz(aVariant) = "" Then
        fncFirstDay = ""
    Else
        fncFirstDay = aVariant & "/01"
    End If
    
End Function

'==============================================================================*
'
'       MODULE_NAME     : 申込書のプレビュー
'       MODULE_ID       : HelperPrintPreview
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strBumoc              部門コード(I)
'                       : strYOUKT_UKNO         予約番号(I)
'                       : strYOUKT_UKNO         予約番号(I)
'                       : intPrintKind          印刷種別(I) 省略可 ※定数宣言参照
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub HelperPrintPreview(ByVal strBUMOC As String, ByVal strYOUKT_UKNO As String, ByRef 割付情報() As MSZZ045割付情報, Optional intPrintKind As Integer = P_PRINT_申込書, Optional str毎月割引 As String, Optional DCNTM_PERIOD As Integer = 0)
    Dim stCONT              As Type_CONT_MAST
    Dim stBANK              As Type_BANK_INF
    Dim stMAIN              As Type_MAIN_INF
    Dim iCount              As Integer                                  'INSERT 2008/10/27 iizuka
    On Error GoTo ErrorHandler
    
    Call GetHelperData(strBUMOC, strYOUKT_UKNO, stMAIN, stCONT, stBANK)
    
    stMAIN.lng割引適用期間 = DCNTM_PERIOD  ' INS 2013/03/30 M.HONDA
    stMAIN.lng毎月値引き額 = Val(str毎月割引)   '2009/02/13 画面より接続
    
    ReDim wk割付情報(UBound(割付情報))
    For iCount = 0 To UBound(割付情報)
        wk割付情報(iCount) = 割付情報(iCount)
    Next iCount
    
    Call HelperPrintPreview2(stMAIN, stCONT, stBANK, intPrintKind)      'INSERT 2007/11/15 SHIBAZAKI

Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintPreview" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 申込書のプレビュー
'       MODULE_ID       : HelperPrintPreview2
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : stMAIN                ヘルパーメイン情報(I)
'                       : stCONT                コントロールマスタ情報(I)
'                       : stBANK                口座情報(I)
'                       : intPrintKind          印刷種別(I) 省略可 ※定数宣言参照
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub HelperPrintPreview2(stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, stBANK As Type_BANK_INF, Optional intPrintKind As Integer = P_PRINT_申込書)
    On Error GoTo ErrorHandler
    
    Call HelperPrintXX(stMAIN, stCONT, stBANK, , intPrintKind)  'INSERT 2007/11/15 SHIBAZAKI
Exit Sub

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintPreview2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 申込書をＰＤＦ印刷
'       MODULE_ID       : HelperPrintPDF
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strBumoc              部門コード(I)
'                       : strYOUKT_UKNO         予約番号(I)
'                       : strOutputPath         出力先(I)
'       RETURN          : フルパスＰＤＦファイル名(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function HelperPrintPDF(ByVal strBumoc As String, ByVal strYOUKT_UKNO As String, ByVal strOutputPath As String) As String 'DELETE 2015/07/15 K.ISHIZAKA
Public Function HelperPrintPDF(ByVal strBUMOC As String, ByVal strYOUKT_UKNO As String, ByVal strOutputPath As String, Optional intPrintKind As Integer = P_PRINT_申込書) As String 'INSERT 2015/07/15 K.ISHIZAKA
    Dim stCONT              As Type_CONT_MAST
    Dim stBANK              As Type_BANK_INF
    Dim stMAIN              As Type_MAIN_INF
    On Error GoTo ErrorHandler
    
    Call GetHelperData(strBUMOC, strYOUKT_UKNO, stMAIN, stCONT, stBANK)
'    HelperPrintPDF = HelperPrintPDF2(stMAIN, stCONT, stBANK, strOutputPath)    'DELETE 2015/07/15 K.ISHIZAKA
    HelperPrintPDF = HelperPrintPDF2(stMAIN, stCONT, stBANK, strOutputPath, intPrintKind) 'INSERT 2015/07/15 K.ISHIZAKA
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintPDF" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 申込書をＰＤＦ印刷
'       MODULE_ID       : HelperPrintPDF2
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : stMAIN                ヘルパーメイン情報(I)
'                       : stCONT                コントロールマスタ情報(I)
'                       : stBANK                口座情報(I)
'                       : strOutputPath         出力先(I)
'       RETURN          : フルパスＰＤＦファイル名(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function HelperPrintPDF2(stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, stBANK As Type_BANK_INF, ByVal strOutputPath As String) As String 'DELETE 2015/07/15 K.ISHIZAKA
Public Function HelperPrintPDF2(stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, stBANK As Type_BANK_INF, ByVal strOutputPath As String, Optional intPrintKind As Integer = P_PRINT_申込書) As String 'INSERT 2015/07/15 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    If Right(strOutputPath, 1) <> "\" Then
        strOutputPath = strOutputPath & "\"
    End If
    If intPrintKind = P_PRINT_申込書 Then                                       'INSERT 2015/07/15 K.ISHIZAKA
        strOutputPath = strOutputPath & stMAIN.str受付番号 & ".pdf"
    Else                                                                        'INSERT START 2015/07/15 K.ISHIZAKA
        strOutputPath = strOutputPath & stMAIN.str受付番号 & "_" & Format(intPrintKind, "00") & ".pdf"
    End If                                                                      'INSERT END   2015/07/15 K.ISHIZAKA
    HelperPrintPDF2 = strOutputPath
'    Call HelperPrintXX(stMAIN, stCONT, stBANK, strOutputPath)                  'DELETE 2015/07/15 K.ISHIZAKA
    Call HelperPrintXX(stMAIN, stCONT, stBANK, strOutputPath, intPrintKind)     'INSERT 2015/07/15 K.ISHIZAKA
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintPDF2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : 申込書をＰＤＦ印刷
'       MODULE_ID       : HelperPrintPDF3
'       CREATE_DATE     : 2011/03/02            K.ISHIZAKA
'       PARAM           : stMAIN                ヘルパーメイン情報(I)
'                       : stCONT                コントロールマスタ情報(I)
'                       : stBANK                口座情報(I)
'                       : strOutputPath         出力先(I)
'                       : 割付情報()          　割引情報(I)
'       RETURN          : フルパスＰＤＦファイル名(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Public Function HelperPrintPDF3(stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, stBANK As Type_BANK_INF, ByVal strOutputPath As String, ByRef 割付情報() As MSZZ045割付情報, Optional DCNTM_PERIOD As Long = 0) As String 'DELETE 2015/07/15 K.ISHIZAKA
Public Function HelperPrintPDF3(stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, stBANK As Type_BANK_INF, ByVal strOutputPath As String, ByRef 割付情報() As MSZZ045割付情報, Optional DCNTM_PERIOD As Long = 0, Optional intPrintKind As Integer = P_PRINT_申込書) As String 'INSERT 2015/07/15 K.ISHIZAKA
    Dim i                   As Long
    On Error GoTo ErrorHandler
    
    ReDim wk割付情報(UBound(割付情報))
    For i = 0 To UBound(割付情報)
        wk割付情報(i) = 割付情報(i)
    Next
    
    stMAIN.lng割引適用期間 = DCNTM_PERIOD  ' INS 2013/03/30 M.HONDA
    
    If Right(strOutputPath, 1) <> "\" Then
        strOutputPath = strOutputPath & "\"
    End If
    If intPrintKind = P_PRINT_申込書 Then                                       'INSERT 2015/07/15 K.ISHIZAKA
        strOutputPath = strOutputPath & stMAIN.str受付番号 & ".pdf"
    Else                                                                        'INSERT START 2015/07/15 K.ISHIZAKA
        strOutputPath = strOutputPath & stMAIN.str受付番号 & "_" & Format(intPrintKind, "00") & ".pdf"
    End If                                                                      'INSERT END   2015/07/15 K.ISHIZAKA
    HelperPrintPDF3 = strOutputPath
'    Call HelperPrintXX(stMAIN, stCONT, stBANK, strOutputPath)                  'DELETE 2015/07/15 K.ISHIZAKA
    Call HelperPrintXX(stMAIN, stCONT, stBANK, strOutputPath, intPrintKind)     'INSERT 2015/07/15 K.ISHIZAKA
Exit Function

ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintPDF3" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME     : ヘルパーに必要なデータ取得
'       MODULE_ID       : GetHelperData
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : strBumoc              部門コード(I)
'                       : strYOUKT_UKNO         予約番号(I)
'                       : stMAIN                ヘルパーメイン情報(O)
'                       : stCONT                コントロールマスタ情報(O)
'                       : stBANK                口座情報(O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub GetHelperData(ByVal strBUMOC As String, ByVal strYOUKT_UKNO As String, stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, stBANK As Type_BANK_INF)
    Dim objCon              As Object
    Dim strNYKOM_CAMPC      As String
    Dim strNYKOM_KINYC      As String
    On Error GoTo ErrorHandler

    Set objCon = ADODB_Connection(strBUMOC)
    On Error GoTo ErrorHandler1
    Call Select_CONT_MAST(objCon, stMAIN, stCONT, strNYKOM_CAMPC, strNYKOM_KINYC)
    Call Select_RCPT_TRAN(objCon, strYOUKT_UKNO, stMAIN)
    
'▼ 2011/06/16 add tajima
    ' 保証委託料割引を求めておく...pst保証料割引はPrivateグローバルです。
    If False = MSZZ045_getHoshoWaribikiDCRA_TRAN2(objCon, stMAIN.str契約番号, pst保証料割引) Then
        pst保証料割引.DCRAT_PRICE = 0 '割引が無かった場合は０にして取れなかった事を記録
    End If
'▲ 2011/06/16 add tajima
    
    objCon.Close
    On Error GoTo ErrorHandler

    Set objCon = ADODB_Connection()
    On Error GoTo ErrorHandler1
    Call Select_BANK_INF(objCon, stMAIN, stBANK, strNYKOM_CAMPC, strNYKOM_KINYC)
    objCon.Close
    On Error GoTo ErrorHandler
Exit Sub

ErrorHandler1:
    objCon.Close
ErrorHandler:
    Call Err.Raise(Err.Number, "GetHelperData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : コントロールマスタ情報取得
'       MODULE_ID       : Select_CONT_MAST
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objCon                コネクション(I)
'                       : stMAIN                ヘルパーメイン情報(O)
'                       : stCONT                コントロールマスタ情報(O)
'                       : strNYKOM_CAMPC        会社コード(O)
'                       : strNYKOM_KINYC        金融機関コード(O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Select_CONT_MAST(objCon As Object, stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, ByRef strNYKOM_CAMPC As String, ByRef strNYKOM_KINYC As String)
    Dim objRst              As Object
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " CONT_BUMOC,"
    strSQL = strSQL & " CONT_ADDR_1,"
    strSQL = strSQL & " CONT_ADDR_2,"
    strSQL = strSQL & " CONT_TEL_NO,"
    strSQL = strSQL & " CONT_FAX_NO,"
    strSQL = strSQL & " CONT_CANCEL_TEL_NO,"                                    'INSERT 2018/03/10 N.IMAI
    strSQL = strSQL & " CONT_YUBINO,"
    strSQL = strSQL & " CONT_KAISYA,"
    strSQL = strSQL & " CONT_KEIUKE_TLD,"
    strSQL = strSQL & " CONT_CAMPC,"                                            'INSERT 2018/09/22 EGL
    strSQL = strSQL & " CONT_TOIAWASE_MAIL,"                                    'INSERT 2018/09/22 EGL
    strSQL = strSQL & " CONT_SEIKYUSYO_TOUROKU_NO,"                             'INSERT 2019/08/15 Y.WADA
    strSQL = strSQL & " CONT_SENDEN1 AS NYKOM_CAMPC,"
    strSQL = strSQL & " CONT_SENDEN2 AS NYKOM_KINYC "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " CONT_MAST "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & "  CONT_KEY = 1"
    
    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler1
    With objRst
        stMAIN.str部門コード = .Fields("CONT_BUMOC")
        stCONT.CONT_ADDR_1 = .Fields("CONT_ADDR_1")
        stCONT.CONT_ADDR_2 = Nz(.Fields("CONT_ADDR_2"))
        stCONT.CONT_TEL_NO = .Fields("CONT_TEL_NO")
        stCONT.CONT_FAX_NO = .Fields("CONT_FAX_NO")
        stCONT.CONT_YUBINO = .Fields("CONT_YUBINO")
        stCONT.CONT_KAISYA = .Fields("CONT_KAISYA")
        stCONT.CONT_KEIUKE_TLD = .Fields("CONT_KEIUKE_TLD")
        stCONT.CONT_CANCEL_TEL_NO = .Fields("CONT_CANCEL_TEL_NO")               'INSERT 2018/03/10 N.IMAI
        strNYKOM_CAMPC = .Fields("NYKOM_CAMPC")
        strNYKOM_KINYC = .Fields("NYKOM_KINYC")
        stCONT.CONT_CAMPC = .Fields("CONT_CAMPC")                               'INSERT 2018/09/22 EGL
        stCONT.CONT_TOIAWASE_MAIL = .Fields("CONT_TOIAWASE_MAIL")               'INSERT 2018/09/22 EGL
        stCONT.CONT_SEIKYUSYO_TOUROKU_NO = Nz(.Fields("CONT_SEIKYUSYO_TOUROKU_NO")) 'INSERT 2019/08/15 Y.WADA
        .Close
    End With
    On Error GoTo ErrorHandler
Exit Sub
    
ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_CONT_MAST" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : ヘルパーメイン情報取得
'       MODULE_ID       : Select_RCPT_TRAN
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objCon                コネクション(I)
'                       : strYOUKT_UKNO         予約番号(I)
'                       : stMAIN                ヘルパーメイン情報(O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Select_RCPT_TRAN(objCon As Object, ByVal strYOUKT_UKNO As String, stMAIN As Type_MAIN_INF)
    Dim objRst              As Object
    Dim strSQL              As String
'    Dim lngSECUKG_KASE      As Long                                            'DELETE 2013/04/22 K.ISHIZAKA 'INSERT 2013/04/14 K.ISHIZAKA
    On Error GoTo ErrorHandler

    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " YOUKT_YUKBN,"                                           'INSERT 2007/09/01 K.ISHIZAKA
    strSQL = strSQL & " YARD_NAME,"
    strSQL = strSQL & " YARD_CODE,"
    strSQL = strSQL & " YARD_ADDR_1 + ISNULL(YARD_ADDR_2,'') + ISNULL(YARD_ADDR_3,'') AS YADDR,"
    strSQL = strSQL & " RCPT_CNO,"
    strSQL = strSQL & " ISNULL(CNTA_SIZE, YOUKT_SIZE_FROM) AS CNTA_SIZE,"
    strSQL = strSQL & " STEPNAME.NAME_NAME AS STEP_NAME,"
    strSQL = strSQL & " ISNULL(CNTA_USAGE, YOUKT_USAGE) AS CNTA_USAGE,"
    strSQL = strSQL & " CNTA_FLOOR," 'ADD 2018/10/02 EGL
    strSQL = strSQL & " USAGENAME.NAME_NAME AS USAGE_NAME,"
    strSQL = strSQL & " KAGIICDNAME.NAME_NAME AS KAGIICD_NAME,"
    strSQL = strSQL & " YOUKT_KUDATE,"  'fix 2007/12/01 tajima
    strSQL = strSQL & " YOUKT_UKNO,"
    strSQL = strSQL & " RCPT_CYDATE,"
    strSQL = strSQL & " ISNULL(RCPT_UKTANTO, YOUKT_UKTANTO) AS YOUKT_UKTANTO,"
    strSQL = strSQL & " RCPT_STDATE,"
    strSQL = strSQL & " YOUKT_TKDATE,"
    strSQL = strSQL & " RCPT_KISAN_DATE,"
    strSQL = strSQL & " SEIKYUNAME.NAME_NAME AS SEIKYU_NAME,"
    strSQL = strSQL & " MITUNAME.NAME_NAME AS MITU_NAME,"
    strSQL = strSQL & " RCPT_SEIKYU_KIKAN,"
    strSQL = strSQL & " RCPT_KAGIICD,"                                          'INSERT 2007/11/15 SHIBAZAKI
    strSQL = strSQL & " RCPT_SUMIKG,"                                           'INSERT 2007/11/15 SHIBAZAKI
    strSQL = strSQL & " RCPT_HOSYICD, "                                         'INSERT 2008/11/13 SHIBAZAKI
    strSQL = strSQL & " RCPT_HOSYO_CD, "                                        'INSERT 2009/04/01 hirano
    strSQL = strSQL & " RCPT_DC_STDATE,"                                        'INSERT 2008/10/27 iizuka
    strSQL = strSQL & " RCPT_DC_ENABLE,"                                        'INSERT 2008/10/27 iizuka
    strSQL = strSQL & " RCPT_KAIHI,"                                            'INSERT 2010/09/02 ryu
    strSQL = strSQL & " RCPT_UKTANTO, "                                         'INS 2013/04/08 M.HONDA
    strSQL = strSQL & " RCPT_BUIL_KAGI_NO, "                                    '2021/06/02 EGL INS
    strSQL = strSQL & " YOUKT_YUBINO,"
    strSQL = strSQL & " YOUKT_ADR_1,"
    strSQL = strSQL & " YOUKT_ADR_2,"
    strSQL = strSQL & " YOUKT_ADR_3,"
    strSQL = strSQL & " YOUKT_KANA,"
    strSQL = strSQL & " YOUKT_NAME,"
    strSQL = strSQL & " YOUKT_KKBN,"
    strSQL = strSQL & " YOUKT_TANM,"
    strSQL = strSQL & " YOUKT_UCODE,"
    strSQL = strSQL & " YOUKT_TEL,"
    strSQL = strSQL & " YOUKT_FAX,"
    strSQL = strSQL & " YOUKT_KEITAI,"
    strSQL = strSQL & " YOUKT_MAIL,"
    strSQL = strSQL & " YOUKT_KKDATE,"  '希望開始日 add 2012/06/18
    strSQL = strSQL & " USER_YUBINO,"
    strSQL = strSQL & " USER_ADR_1,"
    strSQL = strSQL & " USER_ADR_2,"
    strSQL = strSQL & " USER_ADR_3,"
    strSQL = strSQL & " USER_KANA,"
    strSQL = strSQL & " USER_NAME,"
    strSQL = strSQL & " USER_KKBN,"
    strSQL = strSQL & " USER_TANM,"
    strSQL = strSQL & " USER_CODE,"
    strSQL = strSQL & " USER_TEL,"
    strSQL = strSQL & " USER_FAX,"
    strSQL = strSQL & " USER_KEITAI,"
    strSQL = strSQL & " USER_MAIL,"
    strSQL = strSQL & " RCPT_CARG_ACPTNO,"
    strSQL = strSQL & " RCPT_HOSYB,"
    strSQL = strSQL & " RCPT_CAMPC,"                                        '会社コード   add 2019/09/25 add
    strSQL = strSQL & " ISNULL(YARD_SYUKYAKU_KBN,0) AS YARD_SYUKYAKU_KBN,"  '集客契約区分 add 2019/09/25 add
    strSQL = strSQL & " ISNULL(RCPT_RENTKG, YOUKT_PRICFROM) AS RCPT_RENTKG,"
    strSQL = strSQL & " RCPT_FIRSTKG,"
    strSQL = strSQL & " RCPT_SECUKG,"
    strSQL = strSQL & " RCPT_HOSHOU_WARIBIKI,"  'add 2011/06/18 tajima
    strSQL = strSQL & " RCPT_SECUKG_KASE,"                                      'INSERT 2013/04/14 K.ISHIZAKA
    strSQL = strSQL & " EZAPPINAME.NAME_NAME AS EZAPPI_NAME,"
    strSQL = strSQL & " RCPT_EZAPPI,"
    strSQL = strSQL & " RCPT_EZAPPI_DAILY,"
    strSQL = strSQL & " EZAPPINAME1.NAME_NAME AS ADD_EZAPPI_NAME1,"
    strSQL = strSQL & " RCPT_ADD_EZAPPI1,"
    strSQL = strSQL & " RCPT_ADD_EZAPPI_DAILY1,"
    strSQL = strSQL & " EZAPPINAME2.NAME_NAME AS ADD_EZAPPI_NAME2,"
    strSQL = strSQL & " RCPT_ADD_EZAPPI2,"
    strSQL = strSQL & " RCPT_ADD_EZAPPI_DAILY2,"
    strSQL = strSQL & " FZAPPINAME.NAME_NAME AS FZAPPI_NAME,"
    strSQL = strSQL & " RCPT_FZAPPI,"
    strSQL = strSQL & " FZAPPINAME1.NAME_NAME AS ADD_FZAPPI_NAME1,"
    strSQL = strSQL & " RCPT_ADD_FZAPPI1,"
    strSQL = strSQL & " FZAPPINAME2.NAME_NAME AS ADD_FZAPPI_NAME2,"
    strSQL = strSQL & " RCPT_ADD_FZAPPI2,"
    strSQL = strSQL & " ISNULL(RCPT_SEND_CD, YOUKT_SEND_CD) AS YOUKT_SEND_CD,"
    '2009/02/02 hirano MOD Start　毎月割引額合計　追加
    'strSQL = strSQL & " YOUKT_GENKBN "                                          'INSERT 2007/08/28 K.ISHIZAKA
    strSQL = strSQL & " YOUKT_GENKBN, "                                          'INSERT 2007/08/28 K.ISHIZAKA
    strSQL = strSQL & " RCPT_EWARIBIKI, "
    '2009/02/02 hirano MOD End
    strSQL = strSQL & " RCPT_SLOPE, "                              '' 2012/06/12 M.HONDA INS
    strSQL = strSQL & " RCPT_KOUKBN, "                             '' 2013/10/17 M.HONDA INS
'↓ UPDATE 2014/03/21 MIYAMOTO
'    strSQL = strSQL & " SLOPENAME.NAME_NAME AS RCPT_SLOPE_NAME "   '' 2012/06/12 M.HONDA INS
    strSQL = strSQL & " SLOPENAME.NAME_NAME AS RCPT_SLOPE_NAME, "   '' 2012/06/12 M.HONDA INS
    strSQL = strSQL & " RCPT_KAINO "
'↑ UPDATE 2014/03/21 MIYAMOTO
    strSQL = strSQL & ",RCPT_KAGI_NO "                             '2020/06/29 Takenouchi INS
'↓ INSERT 2014/10/14 tajima
    strSQL = strSQL & ",USER_BIRTHDAY "
    strSQL = strSQL & ",RCPT_TPOINT "
    strSQL = strSQL & ",USEPERIODNAME.NAME_NAME AS USE_PERIOD_NAME "
    strSQL = strSQL & ",YOUKT_BUTU "
    strSQL = strSQL & ",KOUKOKUNAME.NAME_NAME AS KOUKOKU_NAME "
    strSQL = strSQL & ",SEIBETSUNAME.NAME_NAME AS SEIBETSU_NAME "
    strSQL = strSQL & ",USER_TAKA "
    strSQL = strSQL & ",RCPT_SLOPE "
'↑ INSERT 2014/10/14 tajima
'2015/01/30 M.HONDA INS
    strSQL = strSQL & ",USER_KNAME "
    strSQL = strSQL & ",USER_JOB "
    strSQL = strSQL & ",USER_KTEL "
    strSQL = strSQL & ",RCPT_KREN_NAME "
    strSQL = strSQL & ",RCPT_KREN_KANA "
    strSQL = strSQL & ",RCPT_KREN_ZOKUGARA "
    strSQL = strSQL & ",RCPT_KREN_TEL "
    strSQL = strSQL & ",RCPT_KREN_KEITAI "
    strSQL = strSQL & ",USER_TBUSHO "
    strSQL = strSQL & ",USER_TNAME "
    strSQL = strSQL & ",USER_TTEL "
    strSQL = strSQL & ",USER_TKEITAI "
    strSQL = strSQL & ",RCPT_BIKE_KOUKBN "
    strSQL = strSQL & ",RCPT_BIKE_NUMBER "
    strSQL = strSQL & ",RCPT_BIKE_SHASYU "
    strSQL = strSQL & ",RCPT_BIKE_HAIKIRYO "
    strSQL = strSQL & ",RCPT_NEN_KBN "
'2015/01/30 M.HONDA INS
    strSQL = strSQL & ",DCRAT_SEV1N " '2015/10/13 M.HONDA INS
    strSQL = strSQL & ",DCRAT_SEV2N " '2015/10/13 M.HONDA INS
    strSQL = strSQL & ",DCRAT_SEV3N " '2015/10/13 M.HONDA INS
    strSQL = strSQL & ",DCRAT_ENDEN " '2015/10/13 M.HONDA INS
    strSQL = strSQL & ",DCRAT_EXMONTH " '2015/10/13 M.HONDA INS
    strSQL = strSQL & ",YARD_TIHO_KBN " '2018/03/19 M.HONDA INS
    strSQL = strSQL & ",RCPT_NET_WARIBIKI " '2018/08/25 tajima INS
   
    strSQL = strSQL & ",CNTA_FLOOR "
    
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " YOUK_TRAN "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " RCPT_TRAN "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " RCPT_NO = YOUKT_RCPT_NO "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " YARD_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " YARD_CODE = ISNULL(RCPT_YCODE, YOUKT_YCODE) "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " USER_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " USER_CODE = RCPT_UCODE "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " CNTA_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CNTA_CODE = RCPT_YCODE "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CNTA_NO = RCPT_CNO "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST STEPNAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " STEPNAME.NAME_ID = '090' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " STEPNAME.NAME_CODE = ISNULL(CNTA_STEP, YOUKT_STEP) "
' 2018/08/25 del strat 初回毎月雑費共に受付トランから取得
'    strSQL = strSQL & "LEFT OUTER JOIN"
'    strSQL = strSQL & " PRIC_TABL "
'    strSQL = strSQL & "ON"
'    strSQL = strSQL & " PRIC_YCODE = CNTA_CODE "
'    strSQL = strSQL & "AND"
'    strSQL = strSQL & " PRIC_USAGE = CNTA_USAGE "
'    strSQL = strSQL & "AND"
'    strSQL = strSQL & " PRIC_SIZE = CNTA_SIZE "
'    strSQL = strSQL & "AND"
'    strSQL = strSQL & " PRIC_STEP = CNTA_STEP "
' 2018/08/25 del end
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST USAGENAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " USAGENAME.NAME_ID = '086' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " USAGENAME.NAME_CODE = ISNULL(CNTA_USAGE, YOUKT_USAGE) "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST KAGIICDNAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " KAGIICDNAME.NAME_ID = '060' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " KAGIICDNAME.NAME_CODE = RCPT_KAGIICD "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST SEIKYUNAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " SEIKYUNAME.NAME_ID = '014' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SEIKYUNAME.NAME_CODE = RCPT_SEIKYU_CD "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST MITUNAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " MITUNAME.NAME_ID = '013' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " MITUNAME.NAME_CODE = RCPT_MITU_CD "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST EZAPPINAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " EZAPPINAME.NAME_ID = '096' "
'2018/08/25 chg start
'   strSQL = strSQL & "AND PRIC_EZAPPI_CODE <> 0 "
'   strSQL = strSQL & "AND"
'   strSQL = strSQL & " EZAPPINAME.NAME_CODE = PRIC_EZAPPI_CODE "
'価格表でなく受付トランから取得
   strSQL = strSQL & "AND RCPT_EZAPPI_CODE <> 0 "
   strSQL = strSQL & "AND"
   strSQL = strSQL & " EZAPPINAME.NAME_CODE = RCPT_EZAPPI_CODE "
'2018/08/25 chg end
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST EZAPPINAME1 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " EZAPPINAME1.NAME_ID = '096' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " EZAPPINAME1.NAME_CODE = RCPT_ADD_EZAPPI_CODE1 "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST EZAPPINAME2 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " EZAPPINAME2.NAME_ID = '096' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " EZAPPINAME2.NAME_CODE = RCPT_ADD_EZAPPI_CODE2 "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST FZAPPINAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " FZAPPINAME.NAME_ID = '096' "
'2018/08/25 chg start
'    strSQL = strSQL & "AND PRIC_FZAPPI_CODE <> 0 "
'    strSQL = strSQL & "AND"
'    strSQL = strSQL & " FZAPPINAME.NAME_CODE = PRIC_FZAPPI_CODE "
'価格表でなく受付トランから取得
    strSQL = strSQL & "AND RCPT_FZAPPI_CODE <> 0 "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " FZAPPINAME.NAME_CODE = RCPT_FZAPPI_CODE "
'2018/08/25 chg end
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST FZAPPINAME1 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " FZAPPINAME1.NAME_ID = '096' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " FZAPPINAME1.NAME_CODE = RCPT_ADD_FZAPPI_CODE1 "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST FZAPPINAME2 "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " FZAPPINAME2.NAME_ID = '096' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " FZAPPINAME2.NAME_CODE = RCPT_ADD_FZAPPI_CODE2 "
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST SLOPENAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " SLOPENAME.NAME_ID = '294' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SLOPENAME.NAME_CODE = RCPT_SLOPE "
'↓ INSERT 2014/10/14 tajima
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST USEPERIODNAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " USEPERIODNAME.NAME_ID = '249' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " USEPERIODNAME.NAME_CODE = YOUKT_USE_PERIOD "    '利用予定期間
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST KOUKOKUNAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " KOUKOKUNAME.NAME_ID = '070' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " KOUKOKUNAME.NAME_CODE = YOUKT_UKKBN "  '広告媒体
    strSQL = strSQL & "LEFT OUTER JOIN"
    strSQL = strSQL & " NAME_MAST SEIBETSUNAME "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " SEIBETSUNAME.NAME_ID = '302' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " SEIBETSUNAME.NAME_CODE = USER_SEIBETU "  '性別
'↑ INSERT 2014/10/14 tajima
    '2015/09/29M.HONDAINS
    strSQL = strSQL & "LEFT JOIN DCRA_TRAN ON "
    strSQL = strSQL & " DCRAT_ACPTNO = RCPT_CARG_ACPTNO AND "
    strSQL = strSQL & " DCRAT_SEIKYU_KBN = 0 "
    '2015/09/29M.HONDAINS
    strSQL = strSQL & "WHERE"
        strSQL = strSQL & " YOUKT_UKNO = '" & strYOUKT_UKNO & "' "
    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler1
    With objRst
        stMAIN.int予約受付状態区分 = .Fields("YOUKT_YUKBN")                     'INSERT 2007/09/01 K.ISHIZAKA
        stMAIN.strヤード名 = .Fields("YARD_NAME")
        stMAIN.strヤードコード = Format(.Fields("YARD_CODE"), "000000")
        stMAIN.strヤード住所 = .Fields("YADDR")
        stMAIN.strスペースコード = Format(.Fields("RCPT_CNO"), "000000")
        stMAIN.strスペースサイズ = Nz(.Fields("CNTA_SIZE"))
        stMAIN.var上下段 = .Fields("STEP_NAME")
        stMAIN.intレンタル用途コード = Nz(.Fields("CNTA_USAGE"))
        stMAIN.lng階数 = Nz(.Fields("CNTA_FLOOR"), 0)           'add 2018/10/02 EGL
        stMAIN.strレンタル用途名 = Nz(.Fields("USAGE_NAME"))
        stMAIN.str鍵種別名 = Nz(.Fields("KAGIICD_NAME"))
        stMAIN.dat受付日 = Nz(.Fields("YOUKT_KUDATE"))  'FIX 2007/12/01 tajima
        stMAIN.str受付番号 = Nz(.Fields("YOUKT_UKNO"))
        stMAIN.val解約日 = .Fields("RCPT_CYDATE")
        stMAIN.str受付担当者名 = Nz(.Fields("YOUKT_UKTANTO"))
        stMAIN.dat契約開始日 = Nz(.Fields("RCPT_STDATE"))
        stMAIN.dat契約受付期限日 = Nz(.Fields("YOUKT_TKDATE"))
        stMAIN.dat起算日 = Nz(.Fields("RCPT_KISAN_DATE"))
        stMAIN.str初期費用請求方法 = Nz(.Fields("SEIKYU_NAME"))
        stMAIN.val明細種別 = .Fields("MITU_NAME")
        stMAIN.val請求期間 = .Fields("RCPT_SEIKYU_KIKAN")
        If Nz(.Fields("USER_CODE")) = "" Then
            stMAIN.str顧客郵便番号 = Nz(.Fields("YOUKT_YUBINO"))
            stMAIN.str顧客住所1 = Nz(.Fields("YOUKT_ADR_1"))
            stMAIN.str顧客住所2 = Nz(.Fields("YOUKT_ADR_2"))
            stMAIN.val顧客住所3 = Nz(.Fields("YOUKT_ADR_3"))
            stMAIN.str顧客フリガナ = Nz(.Fields("YOUKT_KANA"))
            stMAIN.str顧客名 = Nz(.Fields("YOUKT_NAME"))
            stMAIN.int顧客区分コード = Nz(.Fields("YOUKT_KKBN"))
            stMAIN.str顧客代表者名 = Nz(.Fields("YOUKT_TANM"))
            stMAIN.str顧客コード = Format(.Fields("YOUKT_UCODE"), "000000")
            stMAIN.val顧客TEL = Nz(.Fields("YOUKT_TEL"))
            stMAIN.val顧客FAX = Nz(.Fields("YOUKT_FAX"))
            stMAIN.val顧客携帯 = Nz(.Fields("YOUKT_KEITAI"))
            stMAIN.val顧客MAIL = Nz(.Fields("YOUKT_MAIL"))
        Else
            stMAIN.str顧客郵便番号 = Nz(.Fields("USER_YUBINO"))
            stMAIN.str顧客住所1 = Nz(.Fields("USER_ADR_1"))
            stMAIN.str顧客住所2 = Nz(.Fields("USER_ADR_2"))
            stMAIN.val顧客住所3 = Nz(.Fields("USER_ADR_3"))
            stMAIN.str顧客フリガナ = Nz(.Fields("USER_KANA"))
            stMAIN.str顧客名 = Nz(.Fields("USER_NAME"))
            stMAIN.int顧客区分コード = Nz(.Fields("USER_KKBN"))
            stMAIN.str顧客代表者名 = Nz(.Fields("USER_TANM"))
            stMAIN.str顧客コード = Format(.Fields("USER_CODE"), "000000")
            stMAIN.val顧客TEL = Nz(.Fields("USER_TEL"))
            stMAIN.val顧客FAX = Nz(.Fields("USER_FAX"))
            stMAIN.val顧客携帯 = Nz(.Fields("USER_KEITAI"))
            stMAIN.val顧客MAIL = Nz(.Fields("USER_MAIL"))
        End If
        stMAIN.str契約番号 = Nz(.Fields("RCPT_CARG_ACPTNO"))
        stMAIN.val承認番号 = Nz(.Fields("RCPT_HOSYB"))
        stMAIN.lng月額使用料 = Nz(.Fields("RCPT_RENTKG"))
        stMAIN.lng初回使用料 = Nz(.Fields("RCPT_FIRSTKG"))
'        lngSECUKG_KASE = Nz(.Fields("RCPT_SECUKG_KASE"), 0)                    'DELETE START 2013/04/22 K.ISHIZAKA 'INSERT START 2013/04/14 K.ISHIZAKA
'        If Nz(.Fields("RCPT_FIRSTKG")) >= lngSECUKG_KASE Then
'            stMAIN.lng初回使用料 = stMAIN.lng初回使用料 - lngSECUKG_KASE
'            lngSECUKG_KASE = 0
'        Else
'            lngSECUKG_KASE = lngSECUKG_KASE - stMAIN.lng初回使用料
'            stMAIN.lng初回使用料 = 0
'        End If                                                                 'DELETE END   2013/04/22 K.ISHIZAKA 'INSERT END   2013/04/14 K.ISHIZAKA
        stMAIN.val保証料 = .Fields("RCPT_SECUKG")
        stMAIN.val保証料加瀬負担分 = Nz(.Fields("RCPT_SECUKG_KASE"), 0)         'INSERT 2013/04/22 K.ISHIZAKA
'        stMAIN.val保証料 = Nz(stMAIN.val保証料, 0) + Nz(.Fields("RCPT_SECUKG_KASE"), 0) 'DELETE 2013/04/22 K.ISHIZAKA 'INSERT 2013/04/14 K.ISHIZAKA
        stMAIN.val保証金割引額 = Nz(.Fields("RCPT_HOSHOU_WARIBIKI"), 0) 'add 2011/06/18 tajima
        If Nz(.Fields("RCPT_EZAPPI"), 0) <> 0 Then                  'INSERT 2008/11/13 SHIBAZAKI
            stMAIN.val毎月雑費名 = .Fields("EZAPPI_NAME")
            stMAIN.val毎月雑費 = .Fields("RCPT_EZAPPI")
            stMAIN.val日割毎月雑費 = .Fields("RCPT_EZAPPI_DAILY")
'            stMAIN.val日割毎月雑費 = stMAIN.val日割毎月雑費 - lngSECUKG_KASE   'DELETE 2013/04/22 K.ISHIZAKA 'INSERT 2013/04/14 K.ISHIZAKA
        End If                                                      'INSERT 2008/11/13 SHIBAZAKI
        
        If Nz(.Fields("RCPT_ADD_EZAPPI1"), 0) <> 0 Then             'INSERT 2008/11/13 SHIBAZAKI
            stMAIN.val追加毎月雑費名1 = .Fields("ADD_EZAPPI_NAME1")
            stMAIN.val追加毎月雑費1 = .Fields("RCPT_ADD_EZAPPI1")
            stMAIN.val追加日割毎月雑費1 = .Fields("RCPT_ADD_EZAPPI_DAILY1")
        End If                                                      'INSERT 2008/11/13 SHIBAZAKI
        
        '20170530 M.HONDA UPD
        'If Nz(.Fields("RCPT_ADD_EZAPPI2"), 0) <> 0 Then             'INSERT 2008/11/13 SHIBAZAKI
        If Nz(.Fields("RCPT_ADD_EZAPPI2"), 0) <> 0 Or Nz(.Fields("RCPT_ADD_EZAPPI_DAILY2"), 0) <> 0 Then
        '20170530 M.HONDA UPD
        
            stMAIN.val追加毎月雑費名2 = .Fields("ADD_EZAPPI_NAME2")
            stMAIN.val追加毎月雑費2 = .Fields("RCPT_ADD_EZAPPI2")
            stMAIN.val追加日割毎月雑費2 = .Fields("RCPT_ADD_EZAPPI_DAILY2")
        End If                                                      'INSERT 2008/11/13 SHIBAZAKI
        
        If Nz(.Fields("RCPT_FZAPPI"), 0) <> 0 Then                  'INSERT 2008/11/13 SHIBAZAKI
            stMAIN.val初回雑費名 = .Fields("FZAPPI_NAME")
            stMAIN.val初回雑費 = .Fields("RCPT_FZAPPI")
        End If                                                      'INSERT 2008/11/13 SHIBAZAKI
        '2009/06/30 MOD <S> hirano 事務手数料ゼロも出力する
        'If Nz(.Fields("RCPT_ADD_FZAPPI1"), 0) <> 0 Then             'INSERT 2008/11/13 SHIBAZAKI
        If Nz(.Fields("ADD_FZAPPI_NAME1"), "") <> "" Then              'INSERT 2008/11/13 SHIBAZAKI
        '2009/06/30 MOD <E> hirano
            stMAIN.val追加初回雑費名1 = .Fields("ADD_FZAPPI_NAME1")
        '2018/08/25 chg EGL　↓↓ 追加初回雑費は事務手数料
        '　 stMAIN.val追加初回雑費1 = .Fields("RCPT_ADD_FZAPPI1")
            stMAIN.val追加初回雑費1 = .Fields("RCPT_ADD_FZAPPI1") + (-1 * Nz(.Fields("RCPT_NET_WARIBIKI"), 0)) 'ネット契約ならば事務手数料は割引されているので戻す
        '2018/08/25 chg EGL　↑↑
        End If                                                      'INSERT 2008/11/13 SHIBAZAKI
        
        If Nz(.Fields("RCPT_ADD_FZAPPI2"), 0) <> 0 Then             'INSERT 2008/11/13 SHIBAZAKI
            stMAIN.val追加初回雑費名2 = .Fields("ADD_FZAPPI_NAME2")
            stMAIN.val追加初回雑費2 = .Fields("RCPT_ADD_FZAPPI2")
        End If                                                      'INSERT 2008/11/13 SHIBAZAKI
        stMAIN.val書類送付方法 = .Fields("YOUKT_SEND_CD")
        stMAIN.val発生区分 = .Fields("YOUKT_GENKBN")                            'INSERT 2007/08/28 K.ISHIZAKA
        
        stMAIN.var鍵区分コード = Nz(.Fields("RCPT_KAGIICD"))                    'INSERT 2007/11/15 SHIBAZAKI
        stMAIN.var入金済み金額 = Nz(.Fields("RCPT_SUMIKG"))                     'INSERT 2007/11/15 SHIBAZAKI
        
        stMAIN.str保証区分 = Nz(.Fields("RCPT_HOSYICD"))                        'INSERT 2008/11/13 SHIBAZAKI
        stMAIN.str保証会社コード = Nz(.Fields("RCPT_HOSYO_CD"))                 'INSERT 2009/04/01 hirano
        
        stMAIN.var割引適用開始月 = Nz(.Fields("RCPT_DC_STDATE"), _
                                           Format$(Now, "yyyyMM"))              'INSERT 2008/10/27 iizuka
        stMAIN.var割引有効可否 = Nz(.Fields("RCPT_DC_ENABLE"))                  'INSERT 2008/10/27 iizuka
        stMAIN.lng毎月値引き額 = Nz(.Fields("RCPT_EWARIBIKI"), 0)               'INSERT 2009/02/02 hirano
        
        stMAIN.str回収方法 = Nz(.Fields("RCPT_KAIHI"), "")                       'INSERT 2010/09/02 ryu
        
        stMAIN.strスロープ = Nz(.Fields("RCPT_SLOPE_NAME"), "")                  'INSERT 2012/06/12 HONDA
        
        stMAIN.dat希望開始日 = Nz(.Fields("YOUKT_KKDATE"))                       ' add 2012/06/18
        
        stMAIN.str更新区分 = Nz(.Fields("RCPT_KOUKBN"), 0)                  'INSERT 2013/10/17 M.HONDA
        
        stMAIN.val解除番号 = Nz(.Fields("RCPT_KAINO"), "")                  'INSERT 2014/03/21 MIYAMOTO
        stMAIN.val鍵番号 = Nz(.Fields("RCPT_KAGI_NO"), "")                  'INSERT 2020/06/29 Takenouchi
        
        stMAIN.str物件鍵番号 = Nz(.Fields("RCPT_BUIL_KAGI_NO"), "")         'INSERT 2021/06/02 EGL
        
        '↓ INSERT 2014/10/14 tajima
        stMAIN.val顧客誕生日 = Nz(.Fields("USER_BIRTHDAY"), "")
        stMAIN.valTPOINT番号 = Nz(.Fields("RCPT_TPOINT"), "")
        stMAIN.val顧客性別 = Nz(.Fields("SEIBETSU_NAME"), "")
        stMAIN.val予定収納物 = Nz(.Fields("YOUKT_BUTU"), "")
        stMAIN.val利用予定期間 = Nz(.Fields("USE_PERIOD_NAME"), "")
        stMAIN.val媒体 = Nz(.Fields("KOUKOKU_NAME"), "")
        stMAIN.val顧客代表者名カナ = Nz(.Fields("USER_TAKA"), "")
        stMAIN.intスロープ貸出コード = Nz(.Fields("RCPT_SLOPE"), 0)
        '↑ INSERT 2014/10/14 tajima
        '2015/01/30 M.HONDA INS
        stMAIN.val勤務先名 = Nz(.Fields("USER_KNAME"), "")
        stMAIN.val職種 = Nz(.Fields("USER_JOB"), "")
        stMAIN.val勤め先電話番号 = Nz(.Fields("USER_KTEL"), "")
        stMAIN.val緊急連絡先氏名 = Nz(.Fields("RCPT_KREN_NAME"), "")
        stMAIN.val緊急連絡先カナ = Nz(.Fields("RCPT_KREN_KANA"), "")
        stMAIN.val緊急連絡先続柄 = Nz(.Fields("RCPT_KREN_ZOKUGARA"), "")
        stMAIN.val緊急連絡先TEL = Nz(.Fields("RCPT_KREN_TEL"), "")
        stMAIN.val緊急連絡先携帯 = Nz(.Fields("RCPT_KREN_KEITAI"), "")
        stMAIN.val担当者部署 = Nz(.Fields("USER_TBUSHO"), "")
        stMAIN.val担当者氏名 = Nz(.Fields("USER_TNAME"), "")
        stMAIN.val担当者電話番号 = Nz(.Fields("USER_TTEL"), "")
        stMAIN.val担当者携帯番号 = Nz(.Fields("USER_TKEITAI"), "")
        stMAIN.int購入前可否 = Nz(.Fields("RCPT_BIKE_KOUKBN"), 0)
        stMAIN.val登録ナンバー = Nz(.Fields("RCPT_BIKE_NUMBER"), "")
        stMAIN.val車種 = Nz(.Fields("RCPT_BIKE_SHASYU"), "")
        stMAIN.val排気量 = Nz(.Fields("RCPT_BIKE_HAIKIRYO"), "")
        '2015/01/30 M.HONDA INS
        stMAIN.int年払い = Nz(.Fields("RCPT_NEN_KBN"), 0)
        stMAIN.valサービス1 = Nz(.Fields("DCRAT_SEV1N"), "")        '2015/10/13 M.HONDA INS
        stMAIN.valサービス2 = Nz(.Fields("DCRAT_SEV2N"), "")        '2015/10/13 M.HONDA INS
        stMAIN.valサービス3 = Nz(.Fields("DCRAT_SEV3N"), "")        '2015/10/13 M.HONDA INS
        stMAIN.valサービス期間 = Nz(.Fields("DCRAT_ENDEN"), "")     '2015/10/13 M.HONDA INS
        stMAIN.int満了月数 = Nz(.Fields("DCRAT_EXMONTH"), 0)        '2015/10/13 M.HONDA INS
        stMAIN.val地方 = Nz(.Fields("YARD_TIHO_KBN"), 0)            '2018/03/19 M.HONDA INS
        
        stMAIN.valネット割引額 = Nz(.Fields("RCPT_NET_WARIBIKI"), 0) '2018/08/25 EGL add
        
        stMAIN.str会社コード = Nz(.Fields("RCPT_CAMPC"), 0)             '2019/09/25 EGL add
        stMAIN.int集客契約区分 = Nz(.Fields("YARD_SYUKYAKU_KBN"), 0)    '2019/09/25 EGL add
              
        stMAIN.intフロア = Nz(.Fields("CNTA_FLOOR"), 0)
              
        .Close
    End With
    On Error GoTo ErrorHandler
Exit Sub
    
ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_RCPT_TRAN" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 口座情報取得
'       MODULE_ID       : Select_BANK_INF
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : objCon                コネクション(I)
'                       : stMAIN                ヘルパーメイン情報(I)
'                       : stBANK                口座情報(O)
'                       : strNYKOM_CAMPC        会社コード(I)
'                       : strNYKOM_KINYC        金融機関コード(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Select_BANK_INF(objCon As Object, stMAIN As Type_MAIN_INF, stBANK As Type_BANK_INF, ByVal strNYKOM_CAMPC As String, ByVal strNYKOM_KINYC As String)
    Dim objRst              As Object
    Dim strSQL              As String
    On Error GoTo ErrorHandler

    'INSERT 2022/11/28 N.IMAI Start
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " BANKT_KINYN,"
    strSQL = strSQL & " BANKT_SHITN,"
    strSQL = strSQL & " CODET_NAMEN,"
    strSQL = strSQL & " KOUZM_KOUZB,"
    strSQL = strSQL & " CAMPM_CAMPN NYKOM_KOUZN,"
    strSQL = strSQL & " TANTM_TANTN "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " TANT_MAST,"
    strSQL = strSQL & " KOUZ_MAST "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " CAMP_MAST "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CAMPM_CAMPC = KOUZM_CAMPC "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " BANK_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " BANKT_KINYC = KOUZM_KINYC "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " BANKT_SHITC = KOUZM_SHITC "
    strSQL = strSQL & "INNER JOIN"
    strSQL = strSQL & " CODE_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " CODET_SIKBC = '121' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " CODET_CODEC = KOUZM_YOKII "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " KOUZM_BUMOC = '" & stMAIN.str部門コード & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " KOUZM_KOKYC = '" & stMAIN.str顧客コード & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " TANTM_BUMOC = '" & stMAIN.str部門コード & "' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " TANTM_TANTC = '" & stMAIN.str受付担当者名 & "' "

    Set objRst = ADODB_Recordset(strSQL, objCon)
    On Error GoTo ErrorHandler1
    If objRst.EOF = False Then
        With objRst
            stBANK.BANKT_KINYN = .Fields("BANKT_KINYN")
            stBANK.BANKT_SHITN = .Fields("BANKT_SHITN")
            stBANK.CODET_NAMEN = .Fields("CODET_NAMEN")
            stBANK.NYKOM_KOUZB = .Fields("KOUZM_KOUZB")
            stBANK.NYKOM_KOUZN = .Fields("NYKOM_KOUZN")
            stMAIN.str受付担当者名 = .Fields("TANTM_TANTN")
            .Close
        End With
    Else
    'INSERT 2022/11/28 N.IMAI End
        objRst.Close
        On Error GoTo ErrorHandler
        strSQL = ""
        strSQL = strSQL & "SELECT"
        strSQL = strSQL & " BANKT_KINYN,"
        strSQL = strSQL & " BANKT_SHITN,"
        strSQL = strSQL & " CODET_NAMEN,"
        strSQL = strSQL & " NYKOM_KOUZB,"
        strSQL = strSQL & " NYKOM_KOUZN,"
        strSQL = strSQL & " TANTM_TANTN "
        strSQL = strSQL & "FROM"
        strSQL = strSQL & " TANT_MAST,"
        strSQL = strSQL & " NYKO_MAST "
        strSQL = strSQL & "INNER JOIN"
        strSQL = strSQL & " BANK_TABL "
        strSQL = strSQL & "ON"
        strSQL = strSQL & " BANKT_KINYC = NYKOM_KINYC "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " BANKT_SHITC = NYKOM_SHITC "
        strSQL = strSQL & "INNER JOIN"
        strSQL = strSQL & " CODE_TABL "
        strSQL = strSQL & "ON"
        strSQL = strSQL & " CODET_SIKBC = '121' "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " CODET_CODEC = NYKOM_YOKII "
        strSQL = strSQL & "WHERE"
        strSQL = strSQL & " NYKOM_BUMOC = '" & stMAIN.str部門コード & "' "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " NYKOM_CAMPC = '" & strNYKOM_CAMPC & "' "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " NYKOM_KINYC = '" & strNYKOM_KINYC & "' "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " TANTM_BUMOC = '" & stMAIN.str部門コード & "' "
        strSQL = strSQL & "AND"
        strSQL = strSQL & " TANTM_TANTC = '" & stMAIN.str受付担当者名 & "' "
    
        Set objRst = ADODB_Recordset(strSQL, objCon)
        On Error GoTo ErrorHandler1
        With objRst
            stBANK.BANKT_KINYN = .Fields("BANKT_KINYN")
            stBANK.BANKT_SHITN = .Fields("BANKT_SHITN")
            stBANK.CODET_NAMEN = .Fields("CODET_NAMEN")
            stBANK.NYKOM_KOUZB = .Fields("NYKOM_KOUZB")
            stBANK.NYKOM_KOUZN = .Fields("NYKOM_KOUZN")
            stMAIN.str受付担当者名 = .Fields("TANTM_TANTN")
            .Close
        End With
    End If                                                                      'INSERT 2022/11/28 N.IMAI
    On Error GoTo ErrorHandler
Exit Sub
    
ErrorHandler1:
    objRst.Close
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "Select_BANK_INF" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 申込書をプレビュー／ＰＤＦ印刷
'       MODULE_ID       : HelperPrintXX
'       CREATE_DATE     : 2007/08/14            K.ISHIZAKA
'       PARAM           : stMAIN                ヘルパーメイン情報(I)
'                       : stCONT                コントロールマスタ情報(I)
'                       : stBANK                口座情報(I)
'                       : [strOutputPDF]        出力先(I)
'                       : [intPrintKind]        印刷種別(I) 省略可 ※定数宣言参照
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub HelperPrintXX(stMAIN As Type_MAIN_INF, stCONT As Type_CONT_MAST, stBANK As Type_BANK_INF, Optional strPdfPath As String = "", Optional intPrintKind As Integer = P_PRINT_申込書)
    Dim strPath             As String
    Dim xlApp               As Object
    Dim xlBook              As Object
    Dim xlBook2             As Object       'ADD 2009/04/01 hirano
    Dim xlBook3             As Object       'INSERT 2016/05/19 MIYAMOTO
    Dim varPrintSeets       As Variant
    Dim strHosyicdFrom      As String       'INSERT 2008/11/13 SHIBAZAKI
    Dim strWhere            As String       'INSERT 2008/11/13 SHIBAZAKI
    Dim strErrExcel         As String       'ADD 2009/04/01 hirano
    Dim strHKName           As String       'ADD 2009/04/01 hirano
    Dim strHKAddress        As String       'ADD 2009/04/01 hirano
    Dim strHKTEL            As String       'ADD 2009/04/01 hirano
    Dim strShapeName        As String       'INSERT 2011/02/06 K.ISHIZAKA
    Dim objShape            As Object       'INSERT 2011/02/06 K.ISHIZAKA
    Dim strFileNo           As String       'INSERT 2017/12/20 EGL
    Dim intRowNo            As Integer      'INSERT 2017/12/20 EGL
    Dim strBookNameKojin    As String       'add 2019/09/25 EGL y
    Dim strBookNameYakkan   As String       'add 2019/09/25 EGL y
    Dim strHelperNo         As String       'add 2019/09/25 EGL y
    Dim xlSheet             As Variant
    
    On Error GoTo ErrorHandler
    
    ' ヘルパーファイル名取得
'    strPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""FVS400"" AND INTIF_RECFB = ""HELPER_PATH"""))                              'DELETE 2023/12/04 N.IMAI
        strPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""FVS400"" AND INTIF_RECFB = ""HELPER_PATH_" & stMAIN.str部門コード & """"))  'INSERT 2023/12/04 N.IMAI
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
    strWhere = "NAME_ID = '" & P_NAMEID_保証区分 & "' AND NAME_CODE = " & stMAIN.str保証区分
    strHosyicdFrom = Format(Nz(DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", strWhere), 1), "00")
    
    ' Excelオブジェクトを生成する
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo ErrorHandler1
'    Set xlBook = xlApp.Workbooks.Open(strPath & P_HELPER_統合)                                                         'DELETE 2026/03/10 K.KINEBUCHI
    Set xlBook = xlApp.Workbooks.Open(strPath & IIf(stMAIN.str部門コード = "1", P_HELPER_統合オフィス, P_HELPER_統合))  'INSERT 2026/03/10 K.KINEBUCHI

    '---「基本入力」シートに値設定
    Call MSZZ039.SetBaseSheet(xlBook.Worksheets("基本入力"), stCONT, stBANK, stMAIN, intPrintKind, strHosyicdFrom)
    
    '---◆個人情報ファイルからシートコピーする
    strBookNameKojin = fncGetBookName_Kojin(stMAIN)
    
    'シートのインサート
    Call subInsSheets(xlApp, xlBook, strPath & strBookNameKojin, "個人情報取扱", "契約の案内")
    
    '---◆約款シートファイルからシートコピーする
    strBookNameYakkan = fncGetBookName_Yakkan(stMAIN, strPath)

    'シートのインサート
    Call subInsSheets(xlApp, xlBook, strPath & strBookNameYakkan, "約款", "個人情報取扱")

    '---保証委託分については、保証会社ファイルからシートコピーする
    'If P_CMD明細 <> 1 And strHosyicdFrom = P_HOSYIFROM_保証委託 Then                                   'DELETE 2023/10/16 N.IMAI
    If P_CMD明細 <> 1 And strHosyicdFrom = P_HOSYIFROM_保証委託 And intPrintKind <> P_PRINT_解約書 Then 'INSERT 2023/10/16 N.IMAI
        If Dir$(strPath & Replace(P_HELPER_保証会社, "$@$@$@$@$@$@", Format$(stMAIN.str保証会社コード, "000000"))) = "" Then
            Error 53
            GoTo ErrorHandler2
        End If
        Set xlBook2 = xlApp.Workbooks.Open(strPath & Replace(P_HELPER_保証会社, "$@$@$@$@$@$@", Format$(stMAIN.str保証会社コード, "000000")))
        xlBook2.Activate

        xlBook2.Sheets("保証委託契約書").Select
 
        If strHelperNo = "" Then
            xlBook2.Sheets("保証委託契約書").Copy Before:=xlBook.Sheets("明細書")
        Else
            xlBook2.Sheets("保証委託契約書").Copy After:=xlBook.Sheets(CST_SHEETNO_保証委託書)
        End If
 
        xlBook2.Activate
        xlBook2.Sheets("保証会社個人情報取扱").Select
        
        If strHelperNo = "" Then
            xlBook2.Sheets("保証会社個人情報取扱").Copy Before:=xlBook.Sheets("保証委託契約書")
        Else
            xlBook2.Sheets("保証会社個人情報取扱").Copy After:=xlBook.Sheets(CST_SHEETNO_保証会社個人情報取扱)
        End If
        
        Call xlBook2.subGetHKInfo(strHKName, strHKAddress, strHKTEL, strErrExcel)
        If strErrExcel <> "" Then GoTo ErrorHandler2
        xlBook2.Close False
        
    End If
    
    xlBook.Worksheets("基本入力").Range("保証会社名").VALUE = strHKName
    xlBook.Worksheets("基本入力").Range("保証会社所在地").VALUE = strHKAddress
    xlBook.Worksheets("基本入力").Range("保証会社TEL").VALUE = strHKTEL


    'コンテナ用途が「バイク屋外置場」(31)の場合、バイク契約送付資料ファイルからシートコピーする
    If stMAIN.intレンタル用途コード = 31 Then
        If Dir$(strPath & "バイク契約送付資料.xlsx") = "" Then
            Error 53
            GoTo ErrorHandler3
        End If
        Set xlBook3 = xlApp.Workbooks.Open(strPath & "バイク契約送付資料.xlsx")
        xlBook3.Activate
        xlBook3.Sheets("バイク契約送付資料").Select
        xlBook3.Sheets("バイク契約送付資料").Copy After:=xlBook.Sheets(xlBook.Worksheets.Count)
        
        xlBook3.Close False
        
        'シートに値セット
        xlBook.Worksheets("バイク契約送付資料").Range("B12").VALUE = "=スペースコード"
    End If

    '---出力シート設定
    On Error GoTo ErrorHandler2
    If intPrintKind = P_PRINT_QR Then           'INSERT 2021/06/02 EGL
        varPrintSeets = Array("QRコードあり")   'INSERT 2021/06/02 EGL
    ElseIf P_CMD明細 = 1 Then
        varPrintSeets = Array("明細書")
    Else
        varPrintSeets = MSZZ039.GetOutSheets(stMAIN, intPrintKind, strHosyicdFrom, strPdfPath) 'INSERT 2015/07/28 K.ISHIZAKA
    End If

    
    
    '---請求書に各印を追加
    If strHosyicdFrom <> P_HOSYIFROM_事務手数料 And _
        intPrintKind = P_PRINT_申込書 And _
        Nz(stMAIN.val明細種別, "") = P_明細種類_請求書 Then
        Call subInStamp(xlBook, xlBook.Worksheets(P_明細種類_明細書), P_CELL_各印)
    End If
    
    If Nz(stMAIN.var鍵区分コード) = P_KAGIICD_QRコード Then
         Call subQRStamp(xlBook, xlBook.Worksheets(P_QRコード), P_CELL_QRコード, stMAIN)
    End If
    
    If P_CMD明細 = 1 Then
        xlBook.Sheets("明細書").Select
        xlApp.Visible = True
         Call xlBook.Sheets(varPrintSeets).PrintPreview
        'EXCELファイルを閉じる
        xlBook.Close False
        On Error GoTo ErrorHandler1
        'EXCEL終了
        xlApp.DisplayAlerts = False
        xlApp.Quit
    
    Else
        If strPdfPath = "" Then
            '印刷プレビュー表示
            xlApp.Visible = True
            Call xlBook.Sheets(varPrintSeets).PrintPreview
        Else
            'PDF変換
            Call PDFConvertEx(xlBook.Sheets(varPrintSeets), xlBook.NAME, strPdfPath)
        End If
        
        'EXCELファイルを閉じる
        xlBook.Close False
        On Error GoTo ErrorHandler1
        'EXCEL終了
        xlApp.DisplayAlerts = False
        xlApp.Quit
    End If
    On Error GoTo ErrorHandler
Exit Sub

'↓ INSERT 2016/05/19 MIYAMOTO
ErrorHandler3:
    xlBook.Close False
    If stMAIN.intレンタル用途コード = 31 Then
        xlBook3.Close False
    End If
'↑ INSERT 2016/05/19 MIYAMOTO

ErrorHandler2:
    xlBook.Close False
    '2009/04/01 ADD <S> hirano 保証会社ファイル
    If strHosyicdFrom = P_HOSYIFROM_保証委託 Then
        xlBook2.Close False
    End If
    '2009/04/01 ADD <E> hirano
ErrorHandler1:
    xlApp.Visible = True
    
ErrorHandler:                   '↓自分の関数名
    Call Err.Raise(Err.Number, "HelperPrintXX" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME       : ヘルパーにQRコードを挿入
'       MODULE_ID         : subQRStamp
'       CREATE_DATE       : 2021/06/02
'       PARAM             : 引数(Object)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subQRStamp(xBook As Object, xlSheet As Object, xlCell As String, stMAIN As MSZZ039.Type_MAIN_INF)
On Error GoTo Exception
    Dim strQR_Path          As String
    Dim strQR_PathFile      As String
    Dim strTdlQRCode_Path   As String
    Dim lngResult
    'tdlQRCodeのPATH
    strTdlQRCode_Path = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'MSZZ039' AND INTIF_RECFB = 'PATH_QR_EXE'"))
    'QRのPATH
    strQR_Path = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB = 'MSZZ039' AND INTIF_RECFB = 'PATH_QR_FILE'"))
    strQR_PathFile = strQR_Path & stMAIN.str物件鍵番号 & ".png"
    If Dir(strQR_PathFile, vbNormal) = "" Then
        'QRコード作成 実行
        lngResult = ShellWScriptShell(strTdlQRCode_Path & "tdlQRCode.exe -s 10 " & strQR_PathFile & " " & stMAIN.str物件鍵番号)
    End If
    'QRコード貼り付け(Excel側) 実行
    Call xBook.DoQRAddPicture(strQR_PathFile, xlSheet, xlCell)
    Exit Sub

Exception:
    Call Err.Raise(Err.Number, "subQRStamp" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Resume
End Sub

'==============================================================================*

'==============================================================================*
'
'       MODULE_NAME       : 個人情報ファイル名の取得
'       MODULE_ID         : fncGetBookName_Yakkan
'       CREATE_DATE       : 2019/09/25
'       PARAM             : stMAIN      印刷情報の構造体
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBookName_Kojin(stMAIN As Type_MAIN_INF) As String
    
    Const FILE_NAME = ".xlsx"
    
    Dim strHaseiKbn As String
    
    On Error GoTo Exception:
    
    fncGetBookName_Kojin = "個人情報取扱_" & stMAIN.str会社コード & FILE_NAME
    
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncGetBookName_Kojin" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME       : 約款ファイル名の取得
'       MODULE_ID         : fncGetBookName_Yakkan
'       CREATE_DATE       : 2019/09/25
'       PARAM             : stMAIN      印刷情報の構造体
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetBookName_Yakkan(stMAIN As Type_MAIN_INF, strPath As String) As String
    Dim strKaishaCd As String
    Dim strHaseiKbn As String
    Dim strYoto     As String
    Dim strShukyaku As String
    Dim strFilename As String
    On Error GoTo Exception:
       
    '会社コード
    If stMAIN.str会社コード = "KAS" Then
        'コンテナ
        strKaishaCd = "KAS_用途"
        
    ElseIf stMAIN.str会社コード = "KKK" Then
        strKaishaCd = "KKK_用途"
        
    Else
        'トランク
        strKaishaCd = "KTS_用途"
    End If
    
    '2025/11/22 バイク系の約款は屋内バイク以外はコンテナの約款に集約
    '用途
    If stMAIN.intレンタル用途コード = 3 Or stMAIN.intレンタル用途コード = 31 Or stMAIN.intレンタル用途コード = 33 Or (stMAIN.intレンタル用途コード = 0 And stMAIN.val予定収納物 = "オートバイ") Then
        strYoto = "0"
    Else
        'それ以外
        strYoto = CStr(stMAIN.intレンタル用途コード)
    End If
    
    
'    '用途
'    If Mid(CStr(stMAIN.intレンタル用途コード), 1, 1) = 3 _
'      And stMAIN.intレンタル用途コード <> 31 _
'      And stMAIN.intレンタル用途コード <> 32 Then
'        '頭が"3"のもの(33:コンテナ(バイク)等は) "3"にするが 31:屋外ラインや31:屋内置場は除外(コード値のままとする）
'        strYoto = "0"
'    Else
'        'それ以外
'        strYoto = CStr(stMAIN.intレンタル用途コード)
'    End If
    
    
    
    '2025/11/22 集客契約の約款は廃止
    '集約契約区分
'    If stMAIN.int集客契約区分 = 0 Then
'        '無し
'        strShukyaku = ""
'    Else
'        '有り
'        strShukyaku = "_集客契約"
'    End If
    
    'ファイル名
    strFilename = "約款_" & strKaishaCd & strYoto & strShukyaku & ".xlsx"
    If Dir(strPath & strFilename, vbNormal) <> "" Then
        '正常ファイル
        fncGetBookName_Yakkan = strFilename
    Else
        '例外ファイル
        fncGetBookName_Yakkan = strKaishaCd & strHaseiKbn & "_OTHERS.xlsx"
    End If
    
    Exit Function
Exception:
    Call Err.Raise(Err.Number, "fncGetBookName_Yakkan" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'==============================================================================*
'
'       MODULE_NAME       : 指定の位置にシートをコピー挿入する
'       MODULE_ID         : subInsSheets
'       CREATE_DATE       : 2019/09/25
'       PARAM             : xlApp           Excelアプリケーションオブジェクト
'       　　              : xlBook          ベースブック(挿入される側)
'                         : strBookName     挿入元のブック名
'                         : strSheetName    挿入元のシート名
'                         : strBaseSheet    挿入先のシート場所(挿入する手前のシート名)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subInsSheets(ByRef xlApp As Object, ByRef xlBook As Object, strBookName As String, strSheetName As String, strBaseSheet As String)
    Dim objInsBook  As Object
    On Error GoTo Exception:
    Set objInsBook = xlApp.Workbooks.Open(strBookName)
    With objInsBook
        .Activate
        .Sheets(strSheetName).Select
        .Sheets(strSheetName).Copy After:=xlBook.Sheets(strBaseSheet)
        .Close False
    End With
    Exit Sub
Exception:
    Call Err.Raise(Err.Number, "subInsSheets" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME       : ヘルパーにスタンプを挿入
'       MODULE_ID         : subInStamp
'       CREATE_DATE       : 20100109  M.RYU
'       PARAM             : 引数(Object)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
 Public Sub subInStamp(xBook As Object, xlSheet As Object, xlCell As String)
On Error GoTo Exception
    Dim strCAMPC As String
    Dim strWhere As String
    Dim strPath As String
    Dim shps As Object
    
    Dim shpss As Object
    Dim RITSU As Double
    RITSU = 0.75
     
    strCAMPC = DLookup("CONT_SENDEN1", "dbo_CONT_MAST", "CONT_KEY=1")
    strWhere = " INTIF_PROGB = 'FKS240' AND INTIF_RECFB = '" & strCAMPC & "' "
    strPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", strWhere), "")
    
'    Call xBook.DoAddPicture(strPath, xlSheet, xlCell)                                                                                  'DELETE 2023/10/03 N.IMAI
    xlSheet.Shapes.AddPicture fileName:=strPath, LinkToFile:=True, SaveWithDocument:=True, Left:=520, Top:=95, WIDTH:=50, HEIGHT:=50    'INSERT 2023/10/03 N.IMAI
    
    Exit Sub

Exception:
'2020/05/01 update egl Call Err.Raise(Err.Number, "fncGetNetHosyoInfo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Call Err.Raise(Err.Number, "subInStamp" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME       : 加瀬負担分表示
'       MODULE_ID         : fncSetZappiKaseFutan
'       CREATE_DATE       : 2013/04/22          K.ISHIZAKA
'       PARAM             : aSheet              エクセルシートオブジェクト
'                         : intCount            明細行番号(I/O)
'                         : strKaisya           会社名(I)
'                         : strZappiName        雑費名(I)
'                         : lngKinga            雑費金額(I)
'                         : lngSECUKG_KASE      加瀬負担残金額(I)
'                         : blnSyogetu          初月の場合はTrue(I)(省略時値:True)  'INSERT 2019/08/15 Y.WADA
'       RETURN            : 加瀬負担残金額(Long)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'DELETE 2019/08/15 Y.WADA Start
'Private Function fncSetZappiKaseFutan(aSheet As Object, ByRef intcount As Integer, ByVal strKaisya As String, _
'    ByVal strZappiName As String, ByVal lngKinga As Long, ByVal lngSECUKG_KASE As Long) As Long
'DELETE 2019/08/15 Y.WADA End
Private Function fncSetZappiKaseFutan(aSheet As Object, ByRef intCount As Integer, ByVal strKaisya As String, _
    ByVal strZappiName As String, ByVal lngKinga As Long, ByVal lngSECUKG_KASE As Long, Optional blnSyogetu As Boolean = True) As Long
'INSERT 2019/08/15 Y.WADA End
    
    On Error GoTo ErrorHandler
    
    If lngSECUKG_KASE > 0 Then
        If lngKinga > lngSECUKG_KASE Then
'            Call subSetZappi(aSheet, intcount, strZappiName & " " & strKaisya & "一部負担（移動）", -lngSECUKG_KASE, "明細", "明細金額")               'DELETE 2019/08/15 Y.WADA
            Call subSetZappi(aSheet, intCount, strZappiName & " " & strKaisya & "一部負担（移動）", -lngSECUKG_KASE, "明細", "明細金額", blnSyogetu)    'INSERT 2019/08/15 Y.WADA
            fncSetZappiKaseFutan = 0
        Else
'            Call subSetZappi(aSheet, intcount, strZappiName & " " & strKaisya & "負担により無料（移動）", -lngKinga, "明細", "明細金額")               'DELETE 2019/08/15 Y.WADA
            Call subSetZappi(aSheet, intCount, strZappiName & " " & strKaisya & "負担により無料（移動）", -lngKinga, "明細", "明細金額", blnSyogetu)    'INSERT 2019/08/15 Y.WADA
            fncSetZappiKaseFutan = lngSECUKG_KASE - lngKinga
        End If
    Else
        fncSetZappiKaseFutan = 0
    End If
Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncSetZappiKaseFutan" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :fncGetKaiName
'        CREATE_DATE      :2018/10/02
'        機能             :階数表示
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncGetKaiName(intKaiSu As Long) As String
    '階数判定
    If intKaiSu = 0 Then
        '階数なし
        fncGetKaiName = ""
    ElseIf intKaiSu >= 1 Then
        '地上
        fncGetKaiName = CStr(intKaiSu) & "階"
    ElseIf intKaiSu <= -1 Then
        '地下
        fncGetKaiName = "地下" & CStr(intKaiSu * -1) & "階"
    End If
End Function
'****************************  ended of program ********************************
