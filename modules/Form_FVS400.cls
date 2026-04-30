VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FVS400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'   ユーザ名　　　　：　(株)加瀬倉庫
'   システム名　　　：　コンテナ管理システム
'**********************************************
'   プログラム名　　：　受付入力
'   プログラムＩＤ　：　FVS400
'   作　成　日　　　：  2006/02/01
'   作　成　者　　　：  イーグルソフト 柴崎
'   Ver            ：  0.0
'   備考           ：
'**********************************************
'修正履歴１
'
'   修　正　日　　　：2006/04/27
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：1.受付管理表の備考に受付コメントを出力
'                   ：2.受付完了後も印刷系のボタンは押下可能にする。
'   Ver             ：0.1
'
'   修　正　日　　　：2006/04/28
'   修　正　者      ：イーグルソフト 柴崎
'   修　正　内　容　：1.新規顧客で一度も保存せずに受付完了すると受注契約番号発番時にエラーになる不具合修正
'                   ：2.契約トラン作成時に鍵番号にNULLが入るオペレーションが行われた際、不正終了する不具合修正。
'                   ：3.前払い月数が一ヶ月の場合に"yyyy年mm月～yyyy年mm月"と明細書に出さないようにする。
'                   ：4.解約日が入っている場合は、鍵同封書類に解約通知を含めない。
'                   ：5.新規受付時は、「受付取消」と「保証会社NG」ボタンを使用不可にする。
'                   ：6.受付管理表に書類送付方法を出力する。
'                   ：7.鍵送付先名称桁数と住所桁数（１～３全ての桁数合計）をB2仕様に合わせてチェックする。
'   Ver             ：0.2
'
'   修　正　日　　　：2006/05/10
'   修　正　者　　　：イーグルソフト 柴崎
'   修　正　内　容　：1.移動予約の場合、移動元が「保証金」の場合のみ保証金を引き継ぐ
'　　　　　　　　　 ：2.自社理由のリプラス保証での受付時、承認番号が入力されなくても受付完了を可能にする。
'　　　　　　　　　 ：3.明細書に日割日数を印字する。
'　　　　　　　　　 ：4.住所１の桁数を登録時にチェックする。
'　　　　　　　　　 ：5.解約月以降の請求期間は登録できないようにする。
'　　　　　　　　　 ：6.書類送付先名の桁を３６⇒１８に変更。
'　　　　　　　　　 ：7.契約変更系の受付の場合、契約開始日は移動元契約の契約開始日の翌月以降であること。
'　　　　　　　　　 ：8.システム日付が前回の契約データの初回契約日以降じゃないと受付完了させない。
'   Ver　　         ：0.3
'
'   修　正　日　　　：2006/05/22
'   修　正　者　　　：イーグルソフト 柴崎
'   修　正　内　容　：1.請求書２対応
'　　　　　　　　　 ：2.申込書に契約開始日を印字する。
'   Ver　　         ：0.4
'
'   修　正　日　　　：2006/05/31
'   修　正　者　　　：イーグルソフト 柴崎
'   修　正　内　容　：1.契約番号は新規受付の保存時に発番する。
'   Ver　　         ：0.5
'
'   修　正　日　　　：2006/06/01
'   修　正　者　　　：イーグルソフト 柴崎
'   修　正　内　容　：1.鍵情報の初期値はコンテナマスタから取得する。
'　　　　　　　　　 ：2.契約トランに鍵貸し出し本数追加。
'　　　　　　　　　 ：3.受付完了時に鍵情報をコンテナマスタに更新する。
'　　　　　　　　　 ：4.セキュリティカード対応。
'　　　　　　　　　 ：5.契約トラン作成時に代引伝票送付日に到着日ではなく発送日を設定する。
'　　　　　　　　　 ：6.配置図ボタン追加
'　　　　　　　　　 ：7.ヤード地図をクリックするとWEB地図を開くようにする。
'   Ver　　         ：0.6
'
'   修　正　日　　　：2006/06/24
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.修繕呼出対応
'   Ver　　         ：0.7
'
'   修　正　日　　　：2006/09/06
'   修　正　者　　　：イーグルソフト 柴崎
'   修　正　内　容　：1.営業期間の表示
'　　　　　　　　　 ：2.契約開始日は予約トラン・希望開始日を初期値とする
'　　　　　　　　　 ：3.契約開始日の入力は営業期間内とする
'　　　　　　　　　 ：4.解約予定日の入力は営業終了日までとする
'　　　　　　　　　 ：5.部屋と顧客の直近修繕を出力
'　　　　　　　　　 ：6.移動予約も契約開始日の初期値を表示する
'   Ver　　         ：0.8
'
'   修　正　日　　　：2007/01/12
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.受付管理表にｅメール欄の追加対応
'　　　　　　　　　 ：2.統合ヘルパー対応
'   Ver　　         ：0.9
'
'   修　正　日　　　：2007/02/07
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.保証委託料＝月額使用料＋毎月雑費計
'　　　　　　　　　 ：2.保証委託料変更機能追加
'   Ver　　         ：1.0
'
'   修　正　日　　　：2007/02/08
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.セキュリティカードのダイアログ呼出中止→フォームに
'   Ver　　         ：1.1
'
'   修　正　日　　　：2007/04/13
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.受付取消ボタン無効化※フォームプロパティの修正のみ
'                   ：2.荷物移動担当入力項目の追加
'   Ver　　         ：1.2
'
'   修　正　日　　　：2007/05/21
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.バイク関連レンタル用途対応
'   Ver　　         ：1.3
'
'   修　正　日　　　：2007/06/09
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：ヘルパー「住所・連絡先・名義変更届書」対応
'   Ver　　         ：1.4
'
'
'   修　正　日　　　：2007/08/23
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：ネット予約対応、タイトル可変
'   Ver　　         ：1.5
'
'   修　正　日　　　：2007/09/25
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：ネット予約対9/12要望、ネット割引サービス自動対応
'   Ver　　         ：1.6
'
'   修　正　日　　　：2007/11/15
'   修　正　者　　　：イーグルソフト SHIBAZAKI
'   修　正　内　容　：統合ヘルパー印刷を共通関数にて実行する
'   Ver　　         ：1.7
'
'   修　正　日　　　：2008/03/23
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.初回費用が日割りになった時点で翌月まで請求期間を自動変更
'                   ：2.移動区分追加コード「10:満了更新」対応
'                   ：3.解約ヤードに対し加瀬理由で移動した場合のみ前契約金額を引き継ぐ
'                   ：4.バーチャルオフィス専用ヘルパー対応
'   Ver　　         ：1.8
'
'   修　正　日　　　：2008/06/10
'   修　正　者　　　：イーグルソフト tajima
'   修　正　内　容　：1.前金額引継不具合対応
'                   ：2.保証委託料に初回雑費が入らない不具合対応
'   Ver　　         ：1.9
'
'   修　正　日　　　：2008/09/25
'   修　正　者　　　：SHIBAZAKI
'   修　正　内　容　：1.臨時保証区分対応
'                   ：1-1.保証区分を追加
'                   ：    追加した保証区分の場合、保証金を「事務手数料」にする
'   Ver　　         ：2.0
'
'   修　正　日　　　：2008/09/26
'   修　正　者　　　：SHIBAZAKI
'   修　正　内　容　：保証区分によって鍵同封書類ボタンが押せないようにしていた制御を無くす
'   Ver　　         ：2.1
'
'   修　正　日　　　：2008/11/04
'   修　正　者　　　：N.Miura
'   修　正　内　容　：ヤード備考をキャンペーン情報のとなりに追加表示
'   Ver　　         ：2.2
'
'   修　正　日　　　：2008/11/13
'   修　正　者　　　：SHIBAZAKI
'   修　正　内　容　：1.保証会社追加
'                   ：1-1.保証区分を追加
'                   ：1-2.保証区分による処理の分岐のハードコーディングを無くし、
'                   ：    NAME_MASTのFROMとTOで処理する。
'                   ：    ただし、入力チェックと初期値、及びデータソースはＨＣのまま。
'                   ：2.事務手数料対応
'                   ：2-1.追加初回雑費１を事務手数料専用にする。
'                   ：2-2.保証金を事務手数料として徴収する保証区分以外は事務手数料を追加する。
'   Ver　　         ：2.3
'
'   修　正　日　　　：2008/10/01
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：1.自動割引対応
'                   ：1-1.項目追加
'                   ：    割引適用開始月
'                   ：    割引有効可否
'                   ：1-2.請求内訳に割引情報を表示
'                   ：1-3.適用割引割付テーブル更新を追加
'                   ：1-4.割引するチェックボックス初期表示（コンテナマスタの値を反映。保存前でも割引き計算させる） By hirano
'                   ：1-5.割引詳細表示は、明細書シートのみ By hirano
'                   ：1-6.保証会社NGは、適用割引割付　物理削除 By hirano
'                   ：1-7.請求トラン、契約トランの割引考慮 byPEGA
'                   ：1-8.保証金項目への割引反映 By hirano
'                   ：1-9.保証金変更フラグの受付トランへ保持 By hirano
'   Ver　　         ：2.4
'
'   修　正　日　　　：2009/01/25
'   修　正　者　　　：pega
'   修　正　内　容　：1.値引額を受付トランに保持
'   Ver　　         ：2.5
'
'   修　正　日　　　：2009/02/02
'   修　正　者　　　：hirano
'   修　正　内　容　：1.値引き月額の算出修正
'                     月またがり再計算処理改善
'   Ver　　         ：2.6
'
'   修　正　日　　　：2009/02/13
'   修　正　者　　　：hirano
'   修　正　内　容　：1.契約トラン作成時の月額修正
'                     1ヶ月分の割引額を適用する。
'   Ver　　         ：2.7
'
'   修　正　日　　　：2009/02/20
'   修　正　者　　　：hirano
'   修　正　内　容　：保証会社複数対応
'                     保証区分１，２を設定可能にする。
'   Ver　　         ：2.8
'
'   修　正　日　　　：2009/02/23
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：1.適用割引更新方法変更。
'                     既存データの編集中は以下の場合のみ適用割引を更新する
'                     ・割引するチェック変更
'                     ・月額料金変更
'                     ・契約開始日変更
'                     ・契約予定日変更
'                     ・利用予定期間変更
'                     ・請求期間変更
'                     ・割引再計算が押された時
'   Ver　　         ：2.9
'
'   修　正　日　　　：2009/02/25
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：1.適用割引更新方法変更。
'                     適用割引を更新する条件を以下に変更
'                     ・初回金額を追加
'                     ・月額料金は請求合計欄の月額料金を使用
'                     ・割引するチェックの判定方法を変更
'                    2.「明細を見る」を押すと常に明細表示してしまうバグ対応
'   Ver　　         ：2.9
'
'   修　正　日　　　：2009/03/09
'   修　正　者　　　：hirano
'   修　正　内　容　：適用割引更新方法変更の不具合修正。
'   Ver　　         ：3.0
'
'   修　正　日　　　：2009/04/01
'   修　正　者　　　：hirano
'   修　正　内　容　：送付方法、見積もり請求区分、初期費用請求方法初期セットの変更
'                     保証会社コード追加　受付、契約への反映
'   Ver　　         ：3.1
'
'   修　正　日　　　：2009/04/22
'   修　正　者　　　：hirano
'   修　正　内　容　：RESUME 外す
'   Ver　　         ：3.2
'
'   修　正　日　　　：2009/04/23
'   修　正　者　　　：hirano
'   修　正　内　容　：保証区分･保証会社 ネット予約の場合の初期表示の不具合修正
'                     ネット予約の場合　変更不可とする。
'   Ver　　         ：3.3
'
'   修　正　日　　　：2009/04/30
'   修　正　者　　　：hirano
'   修　正　内　容　：移動で新規作成の場合（移動理由：自社理由・金額変更）
'                     移動元がリプラスでも維持させる。
'   Ver　　         ：3.4
'
'   修　正　日　　　：2009/06/30
'   修　正　者　　　：hirano
'   修　正　内　容　：受付管理表金額欄の修正
'   Ver　　         ：3.5
'
'   修　正　日　　　：2009/09/11
'   修　正　者　　　：hirano
'   修　正　内　容　：ネット予約の場合、利用予定期間を変更してもネット割引２０００円引きを反映させない。
'   Ver　　         ：3.6
'
'   修　正　日　　　：2009/10/02
'   修　正　者　　　：hirano
'   修　正　内　容　：レンタル用途33 :コンテナ（バイク） 追加
'   Ver　　         ：3.7
'
'   修　正　日　　　：2009/10/23
'   修　正　者　　　：hirano
'   修　正　内　容　：保証会社自動切替対応
'   Ver　　         ：3.8
'
'   修　正　日　　　：2009/11/10
'   修　正　者　　　：hirano
'   修　正　内　容　：「顧客理由移動」時の割引計算の不整合、印刷時エラー
'                     追記、完了ボタンでの不具合修正
'   Ver　　         ：3.9
'
'   修　正　日　　　：2009/12/08
'   修　正　者　　　：M.RYU
'   修　正　内　容　：鍵種別を追加、フォームで鍵コンボボックスの幅を広く修正
'                    　鍵コンボボックスから、”南京錠”を選択できるように修正
'   Ver　　         ：4.0
'
'   修　正　日　　　：2010/01/09
'   修　正　者　　　：M.RYU
'   修　正　内　容　：--モジュールMSZZ0039と関連修正
'                     1.「保証会社無し」の場合、申込書などを「統合ヘルパー00」で出力
'                     2.臨時請求書に各印を挿入
'                     3.「統合ヘルパー」に、明細書シートのテーマが”請求書”になる場合、各印を挿入
'                     4.新規契約時の明細表（経理用）の【保証委託料】の表示欄に保証会社名の表示を追加
'   Ver　　         ：4.1
'
'   修　正　日　　　：2010/03/15
'   修　正　者　　　：M.RYU
'   修　正　内　容　：請求書２（臨時請求書）を出力とき、明細部に月額使用料を記入
'   Ver　　         ：4.2
'
'   修　正　日　　　：2010/03/26
'   修　正　者　　　：M.RYU
'   修　正　内　容　：請求書２（臨時請求書）を出力とき、タイトルの設定⇒受付入力画面に見積もり請求区分名称を表示
'   Ver　　         ：4.3
'
'   修　正　日　　　：2010/03/29
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：請求書発行区分、毎月回収方法を追加
'   Ver　　         ：4.4
'
'   修　正　日　　　：2010/04/10
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：請求書発行区分を契約トランにも追加
'   Ver　　         ：4.5
'
'   修　正　日　　　：2010/04/08
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：何ヶ月利用のキャンペーンを割引するチェックボックスと連動させる。
'   Ver　　         ：4.6 (4.4をマージしました by K.I)
'
'   修　正　日　　　：2010/04/20
'   修　正　者　　　：S.SHIBAZAKI
'   修　正　内　容　：緊急連絡先と勤務先を追加
'   Ver　　         ：4.7
'
'   修　正　日　　　：2010/04/28
'   修　正　者　　　：S.SHIBAZAKI
'   修　正　内　容　：顧客マスタ保存時のバグ修正
'   Ver　　         ：4.8
'
'   修　正　日　　　：2010/05/28
'   修　正　者　　　：M.RYU
'   修　正　内　容　：解約書・鍵同封書類出力エラー修正
'   Ver　　         ：4.9
'
'   修　正　日　　　：2010/06/22
'   修　正　者　　　：M.RYU
'   修　正　内　容　：受付管理表修正　⇒　支払方法・配達指定時間をExcelで固定設定
'   Ver　　         ：5.0
'
'   修　正　日　　　：2010/12/21
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：ｸﾚｼﾞｯﾄｶｰﾄﾞ・口座振替出力ボタン追加
'   Ver　　         ：5.1
'
'   修　正　日　　　：2010/12/27
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：受付管理表修正 ⇒ 備考欄に予約コメントが表示されるように修正
'   Ver　　         ：5.2
'
'   修　正　日　　　：2011/01/29
'   修　正　者　　　：M.RYU
'   修　正　内　容　：請求書2発行⇒xlsレ修正
'   Ver　　         ：5.3
'
'   修　正　日　　　：2011/02/21
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：統合ＤＢへの書き込み
'   Ver　　         ：5.4
'
'   修　正　日　　　：2011/02/28
'   修　正　者　　　：M.RYu
'   修　正　内　容　：受付管理表⇒「担当者名」に”担当者（カナ）”で表示
'   Ver　　         ：5.5
'
'   修　正　日　　　：2011/03/14
'   修　正　者　　　：S.SHIBAZAKI
'   修　正　内　容　：バイクボックス他用途貸し出し対応
'   Ver　　         ：5.6
'
'   修　正　日　　　：2011/04/01
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：1.自社理由自動の際に保証会社番号・承認日を引き継ぐように修正。
'                   ：2.メンテナンスが完了していれば解約日以前でも貸し出せるように修正
'   Ver　　         ：5.7
'
'
'   修　正　日　　　：2011/04/04
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：1.金額変更の際に保証会社番号・承認日を引き継ぐように修正。
'   Ver　　         ：5.8
'
'   修　正　日　　　：2011/04/05
'   修　正　者　　　：M.RYU
'   修　正　内　容　：臨時請求書の出力内容修正
'   Ver　　         ：5.9
'
'   修　正　日　　　：2011/06/01
'   修　正　者　　　：M.RYU
'   修　正　内　容　：オフィスに「自習室」用途を追加したのため
'   Ver　　         ：6.0
'
'   修　正　日　　　：2011/06/18
'   修　正　者　　　：tajima
'   修　正　内　容　：保証委託料割引対応
'   Ver　　         ：6.1
'
'   修　正　日　　　：2011/07/21
'   修　正　者　　　：m.honda
'   修　正　内　容　：保証委託料割引対応
'   Ver　　         ：6.2
'
'   修　正　日　　　：2011/08/25
'   修　正　者　　　：M.RYU
'   修　正　内　容　：割引適用月の設定を修正⇒受付日の月で設定
'                     MSZZ045_fncGetNebikiDataの引数を修正⇒起算日を受付日に変更
'   Ver　　         ：6.3
'
'   修　正　日　　　：2011/08/29
'   修　正　者　　　：M.HONA
'   修　正　内　容　：事務手数料割引
'   Ver　　         ：6.4
'
'   修　正　日　　　：2011/09/27
'   修　正　者　　　：M.RYU
'   修　正　内　容　：移動区分を追加⇒07紹介者金額変更
'   Ver　　         ：6.5
'
'   修　正　日　　　：2012/03/24
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：WEB決済ボタンを追加
'   Ver　　         ：6.6
'
'   修　正　日　　　：2012/04/11
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：スロープ対応
'   Ver　　         ：6.7
'
'   修　正　日　　　：2012/04/17
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：自社・金額変更時の保証会社をそのまま引き継ぐ様に修正
'   Ver　　         ：6.8
'
'   修　正　日　　　：2012/05/09
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：スロープ貸出有を選択時に初回雑費に2000円を表示
'   Ver　　         ：6.9
'
'   修　正　日　　　：2012/06/30
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：地方物件スロープ対応
'   Ver　　         ：7.0
'
'   修　正　日　　　：2012/07/04
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：ログ出力対応
'   Ver　　         ：7.1
'
'   修　正　日　　　：2012/07/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：鍵対応「番線」を使えるように
'   Ver　　         ：7.2
'
'   修　正　日　　　：2012/08/30
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：請求書2の書類送付先条件を担当者から住所へ変更
'   Ver　　         ：7.3
'
'   修　正　日　　　：2012/10/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：移動時に毎週回収方法を引き継ぐ様に修正
'   Ver　　         ：7.4
'
'   修　正　日　　　：2012/10/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：電話受付時でも保証会社 : クレデンスを利用可能に変更
'   Ver　　         ：7.5
'
'   修　正　日　　　：2012/12/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：毎月雑費 : 不動産保険の場合には、初月の日割りを０とする。
'   Ver　　         ：7.6
'
'   修　正　日　　　：2013/01/17
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：電話受付の発生区分 = ネットの場合に委託料計算には月額賃料を利用する。
'   Ver　　         ：7.7
'
'   修　正　日　　　：2013/01/28
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：電話受付の発生区分 = ネットの不具合対応+起算日不具合対応
'   Ver　　         ：7.8
'
'   修　正　日　　　：2013/03/26
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：事務手数料を０円にする条件を追加
'   Ver　　         ：7.9
'
'   修　正　日　　　：2013/03/27
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：請求期間・雑費に色を付ける
'   Ver　　         ：8.0
'
'   修　正　日　　　：2013/03/30
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：割引適用期間を追加
'   Ver　　         ：8.1
'
'   修　正　日　　　：2013/04/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：保証委託料割引の際に、保証加入費を委託料割引に含めるように修正。
'   Ver　　         ：8.2
'
'   修　正　日　　　：2013/04/05
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：Ver8.2の不具合修正
'   Ver　　         ：8.3
'
'   修　正　日　　　：2013/04/10
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：保証委託料加瀬負担チェックを追加
'                     移動区分が顧客理由移動、保証会社変更のとき使用可能
'                     初期値は　オン        、オフ      　とする
'   Ver　　         ：8.4
'
'   修　正　日　　　：2013/05/21
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：Tポイント対応
'   Ver　　         ：8.5
'
'   修　正　日　　　：2013/06/04
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：保証会社RMトラストを利用不可へ。
'                   ：保険切替の場合には、すべての保証会社へ変更できるように修正。
'                   ：終了メッセージを表示させないように修正。
'   Ver　　         ：8.6
'
'   修　正　日　　　：2013/06/21
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：受付日を受付日時に変更する
'                   ：T-POINT還元率を設定する
'                   ：有効期間、割増率、割増適用時間をPGPAに設定する
'   Ver　　         ：8.7
'
'
'   修　正　日　　　：2013/07/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：受付完了時に契約トランに貸出時の用途を保存する。
'   Ver　　         ：8.8
'
'
'   修　正　日　　　：2013/09/19
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：新規登録時の集金区分をコンビニから振込へ変更
'   Ver　　         ：8.9
'
'   修　正　日　　　：2013/10/16
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：更新区分を追加
'                   ：更新区分によりヘルパーの内容を変更
'   Ver　　         ：9.0
'
'   修　正　日　　　：2013/10/29
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：保証料加瀬負担の表示条件を変更
'                   ：
'   Ver　　         ：9.1
'
'   修　正　日　　　：2013/10/31
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：自社移動理由の際に、解約の為、貸禁データを作成しないように修正
'   Ver　　         ：9.2
'
'   修　正　日　　　：2013/11/13
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：ネット受付の場合にも事務手数料を5000円貰うように修正
'   Ver　　         ：9.3
'
'   修　正　日　　　：2013/11/16
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：住所変更届書ボタンを追加
'   Ver　　         ：9.4
'
'   修　正　日　　　：2014/02/28
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：消費税対応
'   Ver　　         ：9.5
'
'   修　正　日　　　：2014/03/11
'   修　正　者　　　：MIYAMOTO
'   修　正　内　容　：該当データの鍵区分が自動鍵の場合は鍵区分の変更不可、及び解除番号を表示
'                   ：該当データの鍵区分が自動鍵以外の場合は自動鍵以外の鍵区分は変更可とする
'   Ver　　         ：9.6
'
'   修　正　日　　　：2014/04/17
'   修　正　者　　　：MIYAMOTO
'   修　正　内　容　：移動(自社理由、顧客理由、未使用移動)による解約予定で、移動元データの鍵区分が
'                   ：自動鍵だった場合、新たに解除番号を発番し、自動鍵解除番号発番トランに登録する
'   Ver　　         ：9.7
'
'   修　正　日　　　：2014/06/20
'   修　正　者　　　：MIYAMOTO
'   修　正　内　容　：移動(自社理由、顧客理由、未使用移動)による解約予定で、移動元・先データの鍵区分が
'                   ：変換対象の場合、コンテナマスタの鍵区分が変換鍵区分に更新する
'   Ver　　         ：9.8
'
'   修　正　日　　　：2014/10/09
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：受付管理表に差額理由を出力するように修正。
'   Ver　　         ：9.9
'
'   修　正　日　　　：2014/10/14
'   修　正　者　　　：tajima
'   修　正　内　容　：ネット契約フェーズ１Ｂ対応
'                   ：選択肢「性別」の追加、ヘルパー新項目対応
'   Ver　　         ：10.0
'
'   修　正　日　　　：2015/01/08
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：運送サービス項目を追加
'   Ver　　         ：10.1
'
'   修　正　日　　　：2015/02/13
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：受付管理表にTPOINT番号を表示
'   Ver　　         ：10.2
'
'   修　正　日　　　：2015/03/27
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：年払い対応
'   Ver　　         ：10.3
'
'   修　正　日　　　：2015/04/21
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：配送サービス情報に項目追加
'   Ver　　         ：10.4
'
'   修　正　日　　　：2015/07/16
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：証明書イメージ表示、証明書確認ボタン追加
'                   ：ネットで初回クレジットのときは、金額関係をロックする
'   Ver　　         ：10.5
'
'   修　正　日　　　：2015/07/23-24
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：一時使用確認承諾書ボタンを追加
'                   ：顧客理由移動時の事務手数料を3,240円へ
'                   ：顧客理由移動時に委託料加瀬負担を利用不可へ
'   Ver　　         ：10.6
'
'   修　正　日　　　：2015/07/26
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：証明書確認時の接続先ミスを修正
'                     証明書の画像ファイルはJPG以外も表示する
'                     もし画像ファイルでないときはアイコンを表示
'                     移動時は証明書確認をしなくてもＯＫとする
'   Ver　　         ：10.7
'
'   修　正　日　　　：2015/08/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：顧客理由移動時に承認番号・承認日を引き継ぐ。
'   Ver　　         ：10.8
'
'   修　正　日　　　：2015/08/05
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：受付完了の条件を変更
'   Ver　　         ：10.9
'
'   修　正　日　　　：2015/10/28
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：請求書発行区分のデフォルトを発行へ変更
'   Ver　　         ：11.0
'
'   修　正　日　　　：2016/05/17
'   修　正　者　　　：SUGIMURA
'   修　正　内　容　：地方ヤードの鍵区分条件変更
'                     MsgBoxのメッセージ変更
'   Ver　　         ：11.1
'
'   修　正　日　　　：2017/02/04
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：初回の請求金額に年払いを含める様に修正
'　　　　　　　　　　 画面のチカチカ部分の削除
'   Ver　　         ：11.2
'
'   修　正　日　　　：2017/07/31
'   修　正　者　　　：ysuzuki
'   修　正　内　容　：追加契約時の不具合対応
'                     ある条件のとき､受付前でもコンテナ契約ファイル､請求ファイルへを登録する
'   Ver　　         ：11.3
'
'   修　正　日　　　：2017/08/03
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：追加雑費に配送手数料が指定された際には金額をセット
'   Ver　　         ：11.4
'
'   修　正　日　　　：2017/08/06
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：トランクルーム港区浜松町対応
'   Ver　　         ：11.5
'
'   修　正　日　　　：2017/09/26
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：スロープ配送手数料の変更
'   Ver　　         ：11.6
'
'   修　正　日　　　：2017/10/07
'   修　正　者　　　：tajima
'   修　正　内　容　：追加契約時の不具合対応の移動契約考慮（移動契約時は追加契約と見なさい）
'   Ver　　         ：11.7
'
'   修　正　日　　　：2017/12/20
'   修　正　者　　　：EGL
'   修　正　内　容　：新プラン対応(部門別の事務手数料、保証委託料の廃止)
'   Ver　　         ：11.8
'
'   修　正　日　　　：2018/01/19
'   修　正　者　　　：EGL
'   修　正　内　容　：新プラン対応漏れ(バーチャルオフィスの時の事務手数料)
'   Ver　　         ：11.9
'
'   修　正　日　　　：2018/01/20
'   修　正　者　　　：EGL
'   修　正　内　容　：新規の場合、保証会社承認番号を自動発番
'   Ver　　         ：11.10
'
'   修　正　日　　　：2018/01/26
'   修　正　者　　　：EGL
'   修　正　内　容　：事務手数料変更対応(事務手数料の初期値は電話で新規で移動以外のときのみ)
'   Ver　　         ：12.0
'
'   修　正　日　　　：2018/01/29
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：仮想口座マスタへのデータ追加
'   Ver　　         ：12.1
'
'   修　正　日　　　：2018/02/08
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：請求書２の口座情報は仮想口座を優先して出力
'   Ver　　         ：12.2
'
'   修　正　日　　　：2018/02/15
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：保証会社無しの場合は承認番号を必須入力としない
'   Ver　　         ：12.3
'
'   修　正　日　　　：2018/02/19
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：共益費を修正してもマスタ値で上書きされる問題に対応
'   Ver　　         ：12.4
'
'   修　正　日　　　：2018/03/18
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：顧客理由移動の際に保証会社を引継ぐように修正(新プラン対応)
'                     初回契約日を自動時にも保持するように修正
'   Ver　　         ：12.5
'
'   修　正　日　　　：2018/03/19
'   修　正　者　　　：EGL
'   修　正　内　容　：振込で入金確認日に値が入れば且つ請求トランが無ければ請求トランを作成する
'　　　　　　　　　　 振込で入金確認日が未入力になった場合請求トランを削除する
'   Ver　　         ：12.6
'
'   修　正　日　　　：2018/03/28
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：起算日が一日で移動理由が無しの場合、共益費初回(日割)をセットする
'   Ver　　         ：12.7
'
'   修　正　日　　　：2018/05/08
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：来社現金の際にも契約トランを作成する。
'   Ver　　         ：12.8
'
'   修　正　日　　　：2018/05/17
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：012215(港区大門)のハードコーディングを停止（MHNIKKに一括置換え）
'   Ver　　         ：12.9
'
'   修　正　日　　　：2018/05/21
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：チェックロジックをコメントアウト
'   Ver　　         ：13.0
'
'   修　正　日　　　：2018/06/18
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：契約完了以降の書類の出力時にはデータの保存は行わない。
'   Ver　　         ：13.1
'
'   修　正　日　　　：2018/07/09
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：オフィスの自社理由移動の際の事務手数料を21,600円とする。
'   Ver　　         ：13.2
'
'   Ver　　         ：13.3....モジュールプロパティの説明のみ変更あり
'
'   修　正　日　　　：2018/08/06
'   修　正　者　　　：EGL
'   修　正　内　容　：No49.既存顧客新保証会社プラン対応
'                     No57.請求書2の仮想口座対応
'                     No55.法人担当者フリガナ対応（緊急連絡先の複写ボタン追加）
'   Ver　　         ：13.4
'
'
'   修　正　日　　　：2018/08/30
'   修　正　者　　　：EGL
'   修　正　内　容　：No80.事務手数料対応...初回と毎月の雑費コード１を受付トランに保持する
'   Ver　　         ：13.5
'
'   修　正　日　　　：2018/09/21
'   修　正　者　　　：EGL
'   修　正　内　容　：階数表示(トランクルーム時)の対応
'   Ver　　         ：13.6
'
'   修　正　日　　　：2018/09/25
'   修　正　者　　　：EGL
'   修　正　内　容　：分社化対応
'   Ver　　         ：13.7
'
'   修　正　日　　　：2018/10/02
'   修　正　者　　　：EGL
'   修　正　内　容　：保証委託料強制初期化取り止め、事務手数料不具合対応
'   Ver　　         ：13.8
'
'   修　正　日　　　：2018/10/10
'   修　正　者　　　：EGL
'   修　正　内　容　：階数表示リリース対応
'   Ver　　         ：13.9
'
'   修　正　日　　　：2018/10/23
'   修　正　者　　　：tajima
'   修　正　内　容　：新保証プラン移行データの読込エラー対応（13.8の改修漏れ）
'   Ver　　         ：14.0
'
'   修　正　日　　　：2019/02/06
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：スロープ貸出し選択時にFZAPPI_CODEが設定されない問題に対応
'   Ver　　         ：14.1
'
'   修　正　日　　　：2019/02/14
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：新規契約時に既に存在する顧客の場合には保証会社を引継ぐように対応
'   Ver　　         ：14.2
'
'   修　正　日　　　：2019/03/08
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：請求情報参照のみの場合、スロープ有無で金額が変更されてしまうので使用不可とする
'   Ver　　         ：14.3
'
'   修　正　日　　　：2019/03/14
'   修　正　者　　　：Y.WADA
'   修　正　内　容　：移動予約の場合に解約予定日を元の契約から取得していたが空欄にする
'   Ver　　         ：14.4
'
'   修　正　日　　　：2019/03/22
'   修　正　者　　　：TFC_Kurosu
'   修　正　内　容　：鍵送付先情報のメール発行日追加対応
'   Ver　　         ：14.5
'
'
'   修　正　日　　　：2019/04/08
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：宛名印刷ボタンを追加
'   Ver　　         ：14.7
'
'   修　正　日　　　：2019/05/01
'   修　正　者　　　：Y.WADA
'   修　正　内　容　：あおぞら銀行仮想口座対応・仮想口座のテーブル名変更(KOUZ_MAST->KOUZ_MAST_J)
'   Ver　　         ：14.8
'
'   修　正　日　　　：2019/06/24
'   修　正　者　　　：EGL
'   修　正　内　容　：ネット契約改善対応、身分証のフォローメール送信有無と回数表示
'   Ver　　         ：14.9
'
'   修　正　日　　　：2019/08/08
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：消費税対応
'   Ver　　         ：15.0
'
'   修　正　日　　　：2019/08/15
'   修　正　者　　　：Y.WADA
'   修　正　内　容　：消費税対応
'   Ver　　         ：15.1
'
'   修　正　日　　　：2019/09/06
'   修　正　者　　　：K.ISHIZAKA
'   修　正　内　容　：消費税対応
'   Ver　　         ：15.2
'
'   修　正　日　　　：2019/09/25
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：8%で書かれている金額を10%へ変更
'   Ver　　         ：15.3
'
'   修　正　日　　　：2019/10/20
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：コンテナのみ
'                     ネット契約で且つクレジット支払の場合には身分証が無くても完了出来るようにする。（メールにて対応）
'   Ver　　         ：15.4
'
'   修　正　日　　　：2020/02/15
'   修　正　者　　　：tajima
'   修　正　内　容　：売上タイミング変更対応、移動(自社理由、金額変更、保証会社変更)以外は全て保存時に請求ファイルを作成する
'   Ver　　         ：15.5
'
'   修　正　日　　　：2020/02/25
'   修　正　者　　　：EGL
'   修　正　内　容　：保証会社無し対応(保証区分"2"の場合はヘルパー02Uを使う）
'   Ver　　         ：15.6
'
'   修　正　日　　　：2020/03/17
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：前回の契約データは現表示データを除く
'   Ver　　         ：15.7
'
'   修　正　日　　　：2020/04/01
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：受付管理表の保証会社欄に「トランク自社」も出力する
'   Ver　　         ：15.8
'
'   修　正　日　　　：2020/05/01
'   修　正　者　　　：EGL
'   修　正　内　容　：受付入力ヘルパー速度改善
'   Ver　　         ：15.9
'
'   修　正　日　　　：2020/06/29
'   修　正　者　　　：EGL
'   修　正　内　容　：毎月回収方法を必須チェック対象に変更
'   Ver　　         ：16.0
'
'   修　正　日　　　：2020/07/16
'   修　正　者　　　：EGL
'   修　正　内　容　：毎月回収方法を必須チェック対象に変更
'   Ver　　         ：16.1
'
'   修　正　日　　　：2020/07/21
'   修　正　者　　　：EGL
'   修　正　内　容　：料金改定対応
'   Ver　　         ：16.2
'
'   修　正　日　　　：2020/09/28
'   修　正　者　　　：EGL
'   修　正　内　容　：料金改定プランコード対応
'   Ver　　         ：16.3
'
'   修　正　日　　　：2020/10/15
'   修　正　者　　　：tajima
'   修　正　内　容　：シェアオフィス対応
'   Ver　　         ：16.4
'
'   修　正　日　　　：2020/10/30
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：エラー対応（原因不明）ADODBを使用することで回避
'   Ver　　         ：16.5
'
'   修　正　日　　　：2020/11/16
'   修　正　者　　　：EGL
'   修　正　内　容　：コンテナ保証会社無対応
'   Ver　　         ：16.6
'
'   修　正　日　　　：2020/12/01
'   修　正　者　　　：EGL
'   修　正　内　容　：ご紹介コード入力対応
'   Ver　　         ：16.7
'
'   修　正　日　　　：2020/12/15
'   修　正　者　　　：K.KINEBUCHI
'   修　正　内　容　：fncMakeFrom内の、機能:データ検索SQL作成でパラメータテーブルとのJOIN箇所
'　　　　　　　　　　 を変更する。
'   Ver　　         ：16.8
'
'   修　正　日　　　：2021/03/31
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：CheckRegiTargetをODBCに変更
'   Ver　　         ：16.9
'
'   修　正　日　　　：2021/05/10
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：コンテナ自社に「埼玉・千葉」を追加
'　　　　　　　　　　 移動の場合保証会社を「2」へ変更可能とする
'   Ver　　         ：17.0
'
'   修　正　日　　　：2021/06/02
'   修　正　者　　　：EGL
'   修　正　内　容　：QR番号対応
'   Ver　　         ：17.1
'
'   修　正　日　　　：2021/08/04
'   修　正　者　　　：EGL
'   修　正　内　容　：次変更鍵区分の取得変更ミス & 移動時の自動鍵未発番 対応
'   Ver　　         ：17.2
'
'   修　正　日　　　：2021/08/18
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：保存ボタンが使用不可の場合、帳票出力時などの保存をしないように変更
'   Ver　　         ：17.3
'
'   修　正　日　　　：2021/11/30
'   修　正　者　　　：M.HONDA
'   修　正　内　容　：自社保証でも保証会社を変更出来るように修正
'   Ver　　         ：17.4
'
'   修　正　日　　　：2021/12/01
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：自社保証の入力チェック対応漏れ
'   Ver　　         ：17.5
'
'   修　正　日　　　：2022/02/14
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：自社保証に変更可能とする
'　　　　　　　　　　 新規登録中に保証会社を変更すると保証番号を発番するロジックが常に動作していたが、
'　　　　　　　　　　 移動理由「無し」の場合だけに変更
'   Ver　　         ：17.6
'
'   修　正　日　　　：2022/08/22
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：口座名義人名は仮想口座マスタ、入金口座マスタの値をそのまま利用する
'   Ver　　         ：17.7
'
'   修　正　日　　　：2022/11/30
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：請求書２の口座名義人を仮想口座の会社名に変更
'   Ver　　         ：17.8
'
'   修　正　日　　　：2023/03/20
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：仮想口座割当時はKOUZ_MASTを利用
'   Ver　　         ：17.9
'
'   修　正　日　　　：2023/03/27
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：仮想口座割当時はADODB_EXECUTEで実行
'   Ver　　         ：18.0
'
'   修　正　日　　　：2023/04/21
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：仮想口座割当時の条件を変更
'   Ver　　         ：18.2
'
'   修　正　日　　　：2023/05/08
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：自社保証を選択不可とする
'   Ver　　         ：18.3
'
'   修　正　日　　　：2023/09/06
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：追加顧客（自社理由移動以外）で前回の保証会社が「970025」の場合は引継がない
'   Ver　　         ：18.4
'
'   修　正　日　　　：2023/09/11
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：追加顧客で前回の保証会社が「970025」の場合は引継がない（自社理由移動でも）
'   Ver　　         ：18.5
'
'   修　正　日　　　：2023/09/21
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：追加顧客で前回の保証会社が「970170」の場合は引継がない
'   Ver　　         ：18.6
'
'   修　正　日　　　：2023/09/20(2023/08/24)
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：インボイス対応
'   Ver　　         ：18.7
'
'   修　正　日　　　：2023/09/26
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：やっぱり自社理由移動（現在はヤード解約時のみ設定される）の場合は前回の保証会社を引継ぐ
'   Ver　　         ：18.8
'
'   修　正　日　　　：2023/10/04
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：やっぱり自社理由移動（現在はヤード解約時のみ設定される）の場合は前回の保証会社を引継ぐ
'   Ver　　         ：18.9
'
'   修　正　日　　　：2023/10/26
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：NYKO_MAST取得の条件に利用可否を追加
'   Ver　　         ：19.0
'
'   修　正　日　　　：2023/11/09
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：申込書、臨時請求書２の印刷時に仮想口座がなければ割当てる
'   Ver　　         ：19.1
'
'   修　正　日　　　：2023/12/04
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：HELPERのパス（INTI_FILE）を部門毎にする
'   Ver　　         ：19.2
'
'   修　正　日　　　：2024/02/19
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：CONT_MASTの会社ｺｰﾄﾞで仮想口座がなければ割当てる
'   Ver　　         ：19.3
'
'   修　正　日　　　：2026/04/01
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：セキュカード追加画面から戻ると初回分割引合計がゼロになる問題に対応
'   Ver　　         ：19.4
'
'   修　正　日　　　：2026/04/22
'   修　正　者　　　：N.IMAI
'   修　正　内　容　：宛名印刷の変更に対応
'   Ver　　         ：19.6
'
'**********************************************
Option Compare Database
Option Explicit

#Const DEBUG_MODE = False
#Const HELPER_SERVER = False

Private Const PROG_ID = "FVS400"

'==============================================================================*
'   変数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'接続ＤＢ名
Private strSqlserver    As String

'接続文字列
Private strConnect      As String

'部門コード
Private strBumonCode    As String

'会社コード
Private strCampCode    As String                                                'INSERT 20060530 N.MIURA

'金融機関コード
Private strKINYCode    As String

'呼び出し元画面ID
Private strFormFrom     As String

'変更前処理で入力した日付を保存する
Private VarChangeDate As Variant

'○○情報のTOPプロパティを格納する
'ページ切り替え時の先頭位置として使う
Private varPageTop    As Variant

'各情報のタブインデックスが最小のコントロールを格納する
'ページ切り替え時にフォーカスをセットするコントロール
Private varSetFocus    As Variant

'ヘルパーパス
Private strHelperPath       As String

'割引上限
Private lngDCMax            As Long                                             'I:20081001 N.Imai

Private wk割付情報()        As MSZZ045割付情報                                  'I:20081107 N.Imai
Private wk割引情報          As MSZZ045割引情報                                  'I:20081220 N.Imai
Private wk割付情報2()       As MSZZ045割付情報                                  'I:20081217 hirano

Private strSave月額料金         As String                                       'I:20090223 N.Imai
Private strSave初回料金         As String                                       'I:20090223 N.Imai
Private strSave契約開始日       As String                                       'I:20090223 N.Imai
Private strSave解約予定日       As String                                       'I:20090223 N.Imai
Private strSave利用予定期間     As String                                       'I:20090223 N.Imai
Private strSave請求期間         As String                                       'I:20090223 N.Imai
Private blnSave割引する         As String                                       'I:20090223 N.Imai
Private blnSave割引再計算       As String                                       'I:20090223 N.Imai
Private strSave保証会社         As String   '2011/06/18 add tajima
Private strSave毎月回収方法     As String   '2011/06/18 add tajima

Private strSaveトリプル         As String   '2013/10/24 add M.HONDA

Private blnEditMode         As Boolean  'True=編集中
Private blnUpdate           As Boolean  'True=データが変更された
Private strFormCaption      As String   '初期表示時のフォームのキャプション
Private strFormTitle        As String   '初期表示時のフォームのタイトル
Private strUserCodeOldValue As String   '顧客コードコンボボックス前回入力値
Private blnEveryReload      As Boolean  'Trueならば終了時に必ず呼び出し元画面をリフレッシュする
Private strYardMapPath      As String   'ヤード地図ファイル格納パス

Private blnネット予約                   As Boolean  'True=ネット予約発生データ
Private bln元ヤード解約                 As Boolean  'true=元ヤードがあってそれは解約済み
Private blncmbBillSeikyuCd              As Boolean  '初期費用請求方法　 カーソル制御フラグ 2009/04/01 ADD
Private blncmbContainerHosyicd          As Boolean  '保証区分　         カーソル制御フラグ 2009/04/01 ADD
Private blncmbContainerHosyoCompany     As Boolean  '保証会社コード     カーソル制御フラグ 2009/04/23 ADD
Private blnネット予約初回               As Boolean  'True=ネット予約発生データ 2009/04/23 ADD
Private strHosyicd_Before               As String   '変更前　保証区分
Private strHosyo_cd_Before              As String   '変更前　保証会社コード 2009/04/23 ADD
Private strネット予約初期保証会社Cd     As String   '2011/06/18 add
Private strSkkbn                        As String   '2013/09/19 M.HONDA INS
Private strKaijoNo                      As String   '自動鍵解除番号 INSERT 2014/03/11 MIYAMOTO
Private strKaijoNo_QR                   As String   '自動鍵解除番号 INSERT 2021/06/02 EGL
Private int初回雑費2Cd                  As Integer  '初回雑費２コード 2015/01/08 add
Private bln保証会社新旧                 As Boolean  'True=新保証会社 2017/12/20

'画面起動時、コンテナ契約ファイル、請求ファイルへの登録対象かどうか保持
Private lngStRtn                        As Long     '2017/07/31 add by ysuzuki
'受付トランから取得した既存顧客新保証会社プラン移行カラムの値を保持
Private strNewHosplnIkou                As String   '2018/08/06 EGL INS
'受付トランから取得した保証金カラムの値を保持
Private lngSecuKg                       As Long     '2018/08/06 EGL INS

'受付トランから取得した雑費
Private lngFZAPPI_CODE                  As Long     '2018/08/11 EGL INS
Private lngEZAPPI_CODE                  As Long     '2018/08/11 EGL INS

'コントロールマスタから取得した会社コードの値を保持
Private strCmCampCd                     As String   '2018/09/25 EGL INS


Private tax_from As String        '2019/08/08 M.HONDA INS
Private tax_to As String          '2019/08/08 M.HONDA INS

Private strプランCd                     As String   '2020/09/28 EGL add

Private pvONLET_ACCEI                   As String   '2021/06/02 EGL add

' 顔認証連携エラーメッセージ保持用
Private mstrFaceAuthError               As String
'==============================================================================*
'   定数宣言
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Const P_KAKUNIN             As String = "確認"
Private Const P_ERROR               As String = "エラー"

'前画面からのパラメータのインデックス 上から順にカンマ区切りで渡してもらう
Private Const P_パラメータ数                        As Integer = 4
Private Const P_パラメータ_呼び出し元画面           As Integer = 0
Private Const P_パラメータ_予約番号                 As Integer = 1
Private Const P_パラメータ_受付番号                 As Integer = 2
Private Const P_パラメータ_元契約番号               As Integer = 3

Private Const P_WHITE           As Long = 16777215     ' 顧客情報の背景色(入力可)
Private Const P_YELLOW          As Long = 14680063     ' 顧客情報の背景色(入力不可)
Private Const P_BLACK           As Long = 0
Private Const P_RED             As Long = 255
Private Const P_ORANGE          As Long = 42495        '2018/08/06 EGL INS


'コントロールの種類
Private Const P_CONTROL_数値        As Integer = 1          '数値
Private Const P_CONTROL_日付        As Integer = 2          '数値
Private Const P_CONTROL_テキスト    As Integer = 3          'テキスト
Private Const P_CONTROL_年月        As Integer = 5          '年月
Private Const P_CONTROL_金額        As Integer = 6          '金額
Private Const P_CONTROL_半角        As Integer = 7          '半角文字
Private Const P_CONTROL_頭ゼロ消去  As Integer = 8          '数値頭０なし

'数値項目書式
Private Const P_FORMAT_STANDARD      As String = "Standard"    ' 書式：標準
Private Const P_FORMAT_FIXED         As String = "Fixed"       ' 書式：固定

Private Const P_ALL_NINE            As String = "9999999999"

'不足桁０埋め用
Private Const P_ZERO_SUPPRESS       As String = "0000000000"

'桁数
Private Const P_DIVISION_LENGTH     As Integer = 2      '区分
Private Const P_PERSON_LENGTH       As Integer = 3      '担当コード
Private Const P_CODE_LENGTH         As Integer = 6      'ヤードコード/顧客コード
Private Const P_DATE_LENGTH         As Integer = 10      '日付
Private Const P_YEARMONTH_LENGTH    As Integer = 7      '年月
Private Const P_TEL_LENGTH          As Integer = 15     '電話番号
Private Const P_KANA_LENGTH         As Integer = 18       '顧客カナ
Private Const P_NAME_LENGTH         As Integer = 36       '顧客名
Private Const P_POST_LENGTH         As Integer = 8       '郵便番号
Private Const P_ADR_LENGTH          As Integer = 36       '住所
Private Const P_DAI_NAME_LENGTH     As Integer = 18       '代表者/書類送付先名
Private Const P_MAIL_LENGTH         As Integer = 60       'メールアドレス
Private Const P_HOUR_LENGTH         As Integer = 2        '時
Private Const P_PRICE_LENGTH        As Integer = 7        '金額
Private Const P_NOTE_LENGTH         As Integer = 60       '備考
Private Const P_DAY_LENGTH          As Integer = 2        '日
Private Const P_STRAGE_LENGTH       As Integer = 60        '収納物
Private Const P_KEYNO_LENGTH        As Integer = 10      '鍵番号
Private Const P_LENDNUM_LENGTH      As Integer = 3       '鍵貸し出し本数
Private Const P_HOSYOUNO_LENGTH     As Integer = 20      '保証番号
Private Const P_SLIPNO_LENGTH       As Integer = 14      '伝票番号
Private Const P_TPOINT_LENGTH       As Integer = 16      'Tポイント 2013/05/21

'2006/04/28 S.SHIBAZAKI
' 鍵送付先名称桁数と住所桁数（１～３全ての桁数合計）をB2仕様に合わせてチェックする
Private Const P_KEY_SEND_NAME       As Integer = 30      '鍵送り先住所（B2データ仕様の桁数）
Private Const P_KEY_SEND_ADD        As Integer = 98      '鍵送り先住所（B2データ仕様の桁数）

'コントロールマスタの検索キー
Private Const P_CONT_KEY            As Integer = 1

'ラベル
Private Const P_LABEL_新規受付      As String = "新規受付"
Private Const P_LABEL_既存受付      As String = "既存受付"
Private Const P_LABEL_受付完了      As String = "受付完了"
Private Const P_LABEL_日割りなし    As String = "日割りなし"
'↓DELETE 2008/11/13 SHIBAZAKI
'Private Const P_LABEL_保証委託料    As String = "保証委託料"
'Private Const P_LABEL_保証金        As String = "保証金"
'Private Const P_LABEL_事務手数料    As String = "事務手数料"        'INSERT 2008/09/25 SHIBAZAKI
'↑DELETE 2008/11/13 SHIBAZAKI

'INSERT 2020/07/21 S.WATANABE
Private Const P_LABEL_初回          As String = "初回"
Private Const P_LABEL_初回日割      As String = "初回(日割)"

'保存処理区分
Private Const P_SAVEMODE_保存       As Integer = 1
Private Const P_SAVEMODE_完了       As Integer = 2
Private Const P_SAVEMODE_取り消し   As Integer = 3
Private Const P_SAVEMODE_保証会社NG As Integer = 4

'顧客新規／既存
Private Const P_USER_OPT_新規       As Integer = 1  '新規顧客
Private Const P_USER_OPT_既存       As Integer = 2  '既存顧客

'送付先同／異
Private Const P_SEND_OPT_顧客同     As Integer = 1
Private Const P_SEND_OPT_送付先同   As Integer = 2
Private Const P_SEND_OPT_異なる     As Integer = 3

'フォームキャプション
Private Const P_CAPTION_EDITING     As String = "(修正中)"

'印刷種別
Private Const P_PRINT_受付管理表        As Integer = 1      '受付管理表
'Private Const P_PRINT_申込書            As Integer = 2      '申込書        'DELETE 2007/11/15 SHIBAZAKI
'Private Const P_PRINT_解約書            As Integer = 3      '解約書        'DELETE 2007/11/15 SHIBAZAKI
'Private Const P_PRINT_鍵書類            As Integer = 4      '鍵同封書類    'DELETE 2007/11/15 SHIBAZAKI
Private Const P_PRINT_請求書2           As Integer = 5      '請求書２
Private Const P_PRINT_鍵の開け方        As Integer = 6      '鍵の開け方

'会社コード　横浜鋼管
Private Const P_CAMP_横浜鋼管            As String = "YOK"

'ヘルパーファイル名
'update 2020/05/01 EGL Private Const P_HELPER_受付管理表       As String = "受付管理表.xls"
Private Const P_HELPER_受付管理表       As String = "受付管理表.xlsx"


Private Const P_HELPER_コンテナ用       As String = "コンテナ用ヘルパーSYS.xls"
Private Const P_HELPER_トランク用       As String = "トランク用ヘルパーSYS.xls"
Private Const P_HELPER_バイク用         As String = "バイク用ヘルパーSYS.xls"
Private Const P_HELPER_横浜鋼管用       As String = "横浜鋼管用ヘルパーSYS.xls"
Private Const P_HELPER_オフィス用       As String = "オフィス用ヘルパーSYS.xls"

'-----20100109-----M.RYU-------update------'新規ヘルパー「統合ヘルパー00.xls」
'Private Const P_HELPER_保証金用         As String = "保証金有りヘルパーSYS.xls"
Private Const P_HELPER_保証金用         As String = "統合ヘルパー00.xls"

Private Const P_HELPER_統合             As String = "統合ヘルパー.xls" '2007/01/12 add tajima

'2005/05/22 SHIBAZAKI 請求書２対応
'update 2020/05/01 EGL Private Const P_HELPER_請求書2          As String = "臨時請求書.xls"
Private Const P_HELPER_請求書2          As String = "臨時請求書.xlsm"

'ヘルパーに出す解約予定の文言
Private Const P_HELPER_CANCEL_STRING    As String = "解約予定あり"

'2005/05/22 SHIBAZAKI 請求書２対応
Private Const P_CODE_MAST_預金種別      As String = "121"

'Private Const P_STRING_振込人名         As String = "振込人名は『@CASEQ@@BUSEQ@@UCODE@@UKANA@』でお願いします。"   'DELETE 20110129 M.RYU
Private Const P_STRING_振込人名         As String = "@CASEQ@@BUSEQ@@UCODE@@UKANA@"      'INSERT 20110129 M.RYU

'2006/06/01 SHIBAZAKI セキュリティカード対応
Private Const P_FORM_セキュリティーカード追加   As String = "FKS600"

'INSERT 2008/11/13 SHIBAZAKI
'Private Const P_JIMUTESURYO             As Long = 5000  '2014/02/28 M.HONDA 消費税対応 DEL
'Private Const P_JIMUTESURYO             As Long = 5400   '2014/02/28 M.HONDA 消費税対応 INS
'Private Const P_KMOVEJIMUTE             As Long = 3240   '2015/07/24 M.HODNA 顧客理由移動対応

Private Const P_JIMUTESURYO             As Long = 5500    '2019/09/26 M.HONDA
Private Const P_KMOVEJIMUTE             As Long = 3300    '2019/09/26 M.HONDA



Private Const P_HOSPLN_移行対象         As String = "1"  '2018/08/06 EGL INS
Private Const P_HOSPLN_移行対象外       As String = "0"  '2018/08/06 EGL INS

'---- コード値 ----

'移動区分
Private Const P_MOVEKBN_無し            As Integer = 0
Private Const P_MOVEKBN_自社理由        As Integer = 1
Private Const P_MOVEKBN_顧客理由        As Integer = 2
Private Const P_MOVEKBN_未使用移動      As Integer = 3
Private Const P_MOVEKBN_名義変更        As Integer = 4
Private Const P_MOVEKBN_金額変更        As Integer = 5
Private Const P_MOVEKBN_保障会社変更    As Integer = 6
Private Const P_MOVEKBN_紹介者金額変更  As Integer = 7   ' 紹介者金額変更  'INSERT 2011/09/27 M.RYU
Private Const P_MOVEKBN_満了更新        As Integer = 10 '2008/03/23 add tajima

'保証区分
'Private Const P_HOSYICD_その他          As Integer = 0     'DELETE 2008/11/13 SHIBAZAKI
'Private Const P_HOSYICD_トリオ          As Integer = 3     'DELETE 2008/11/13 SHIBAZAKI
Private Const P_HOSYICD_リプラス        As Integer = 4
'Private Const P_HOSYICD_セーフティー    As Integer = 6     'DELETE 2008/11/13 SHIBAZAKI
'Private Const P_HOSYICD_臨時            As Integer = 7     'INSERT 2008/09/25 SHIBAZAKI-> DELETE 2008/11/13 SHIBAZAKI
Private Const P_HOSYICD_パルマ          As Integer = 8      'INSERT 2008/11/13 SHIBAZAKI
'2009/04/01 ADD <S> 保証区分見直し
Private Const P_HOSYICD_保証会社無      As Integer = 0
Private Const P_HOSYICD_保証会社有      As Integer = 1
Private Const P_HOSYICD_保証会社無_自社 As Integer = 2      'INSERT 2020/01/24 EGL 保証会社無し対応
Private Const P_HOSYICD_リプラス包括    As Integer = 5
Private Const P_HOSYICD_臨時            As Integer = 7
Private Const P_HOSYICD_その他          As Integer = 9
'2009/04/01 ADD <E>
'集金区分
Private Const P_USER_SKBN_コンビニ      As Integer = 7
Private Const P_USER_SKBN_振込          As Integer = 1      '2013/09/19 M.HONDA INS

'初期費用請求方法
Private Const P_SEIKYUCD_振込み         As Integer = 1
Private Const P_SEIKYUCD_代引き         As Integer = 2
Private Const P_SEIKYUCD_来社現金       As Integer = 3
Private Const P_SEIKYUCD_クレジット     As Integer = 4                          'INSERT 2015/07/16 K.ISHIZAKA

'見積もり請求区分　2009/04/01 追加
Private Const P_MITSUCD_明細書       As String = "01"
Private Const P_MITSUCD_見積書       As String = "02"
Private Const P_MITSUCD_請求書       As String = "03"

'予約状態
Private Const P_YUKBN_受付              As Integer = 10
Private Const P_YUKBN_契約              As Integer = 20
Private Const P_YUKBN_受付キャンセル     As Integer = 53

'ご紹介区分
Private Const P_INTROKBN_取置           As Integer = 1
Private Const P_INTROKBN_受付           As Integer = 2
Private Const P_INTROKBN_契約           As Integer = 3
Private Const P_INTROKBN_キャンセル     As Integer = 6

'契約状態
Private Const P_AGRE_契約               As Integer = 1
Private Const P_AGRE_解約予定           As Integer = 2
Private Const P_AGRE_解約延長           As Integer = 3

'解約区分
Private Const P_KAICD_通常              As Integer = 1

'解約理由
Private Const P_KAIRIYUCD_限定契約      As Integer = 9

'コンテナ用途
Private Const P_USAGE_レンタルボックス   As Integer = 0
Private Const P_USAGE_トランクルーム     As Integer = 1
Private Const P_USAGE_レンタルオフィス   As Integer = 2
Private Const P_USAGE_バイク            As Integer = 3
Private Const P_USAGE_バイクコンテナ    As Integer = 39             'INSERT 2011/03/14 SHIBAZAKI

'取り消し理由
Private Const P_CANCEL_REASON_自動取り消し           As Integer = 1
Private Const P_CANCEL_REASON_手動取り消し           As Integer = 2
Private Const P_CANCEL_REASON_保証会社NG             As Integer = 3

'顧客区分
Private Const P_USER_KKBN_法人           As Integer = 1

'敬称
Private Const P_HONORIFIC_様            As Integer = 2
Private Const P_HONORIFIC_御中          As Integer = 3

'雑費
Private Const P_ZAPPI_無し              As Integer = 0
'2006/06/01 SHIBAZAKI セキュリティカード対応
Private Const P_ZAPPI_追加カード代       As Integer = 84
Private Const P_ZAPPI_事務手数料        As Integer = 63     'INSERT 2008/11/13 SHIBAZAKI

'割引計算
Private Const P_MODE_本計算  As Integer = 0  '本計算                            'I:20081107 N.Imai
Private Const P_MODE_仮計算  As Integer = 1  '仮計算                            'I:20081107 N.Imai

Private Const P_SEIKI_未発行            As Integer = 2                          'INSERT 2010/03/29 K.ISHIZAKA
Private Const P_SEIKI_発行              As Integer = 1                          'INSERT 2015/10/28 M.HONDA



'2006/04/28 S.SHIBAZAKI 乗用具チェック追加。チェックされている場合、申込書に誓約書添付
'==============================================================================*
'
'        MODULE_NAME      :chkAppFormVehicle_AfterUpdate
'        機能             :chkAppFormVehicle_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkAppFormVehicle_AfterUpdate()

    Call subEditMode(True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :chkCampaign_Click
'        機能             :chkCampaign_Click イベント
'        IN               :
'        OUT              :
'        備考             :強制的にﾄﾘﾌﾟﾙｷｬﾝﾍﾟｰﾝを適用させる。
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkCampaign_Click()
    
    Call sub割引再計算(P_MODE_仮計算)
    Call subEditMode(True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chkUserRinput_AfterUpdate
'        機能             :chkUserRinput_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkContainerSevExpire_AfterUpdate()

    txtContainerSevExpire = ""
    If Nz(txtBillKisanDate) = "" Then
        chkContainerSevExpire.VALUE = 0
        Exit Sub
    End If
    
    If chkContainerSevExpire.VALUE = -1 Then
        '起算日の年月＋ヤードマスタのキャンペーン満了月数
        txtContainerSevExpire = Format(DateSerial(Format(txtBillKisanDate, "yyyy"), Format(txtBillKisanDate, "m") + Val(Nz(txtSevExmonth, 0)), 1), "yyyy/mm")
    End If
    
    Call subEditMode(True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      : chkNEN_KBN_Click
'        機能             :
'        CREAT            : 2015/03/20          M.HONDA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkNEN_KBN_Click()
    If Me.chkNEN_KBN.VALUE = -1 Then
        Call YellowControl(cmbJIREI, False)
        Call YellowControl(cmbSEIKI, False)
        Call YellowControl(txtSEIKI, False)
        
        Dim a As Integer
        '年払いの金額を割引を考慮して設定する。
        '割引の情報を取得する。
        'txt年払い = (txtBillRentkgKei + txtBillEzappiKei) * -1
         '2017/07/01 M.HONDA UPD
        If strBumonCode <> "H" Then
            If wk割引情報.件数 = 0 Then
               txt年払い = (txtBillRentkgKei + txtBillEzappiKei) * -1
            Else
               If wk割引情報.TRAN_DATA(0).DCNTM_PERIOD >= 12 Then
                   txt年払い = (txtBillRentkgKei + txtBillEzappiKei + wk割引情報.TRAN_DATA(0).VALUE_TRUE) * -1
               Else
                   txt年払い = (txtBillRentkgKei + txtBillEzappiKei) * -1
               End If
            End If
        End If
        '2017/07/01 M.HONDA UPD
       
    Else
        Call YellowControl(cmbJIREI, True)
        Call YellowControl(cmbSEIKI, True)
        Call YellowControl(txtSEIKI, True)
        txt年払い = 0
        Me.cmbJIREI = ""
        Me.txtJIREI = ""
        Me.cmbSEIKI = ""
        Me.txtnenSEIKI = ""
        Me.txtSEIKI = ""
    End If
End Sub

'==============================================================================*
'
'        MODULE_NAME      : chkSecukgKase_AfterUpdate
'        機能             : 保証委託料加瀬負担
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkSecukgKase_AfterUpdate()
    Dim lngValue            As Long
    
    lngValue = Nz(txtBillSecukg.VALUE)
    txtBillSecukg.VALUE = txtBillSecukgKase.VALUE
    txtBillSecukgKase.VALUE = lngValue
    txtBillSecukg.Visible = Not chkSecukgKase.VALUE
    txtBillSecukgKase.Visible = chkSecukgKase.VALUE
    Call subEditMode(True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chkUserRinput_AfterUpdate
'        機能             :chkUserRinput_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chkUserRinput_AfterUpdate()

    'チェックされたらその他連絡先を入力可にする
    Call subUserRenrakusakiEnable(chkUserRinput.VALUE = -1)
    Call subRenrakusakiClear
    
    Call subEditMode(True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_更新区分_KeyPress
'        機能             :cmb_更新区分_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_更新区分_AfterUpdate()

    txt_更新区分.VALUE = fncAfterUpdateCommon(cmb_更新区分, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmb_更新区分_KeyPress
'        機能             :cmb_更新区分_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_更新区分_BeforeUpdate(Cancel As Integer)
    Call fncBeforeUpdateCommon(cmb_更新区分, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddEzappiCode2_AfterUpdate
'        機能             :cmbBillAddEzappiCode2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddEzappiCode2_AfterUpdate()

    txtBillAddEzappiName2 = fncAfterUpdateCommon(cmbBillAddEzappiCode2, P_DIVISION_LENGTH, P_CONTROL_数値)

    '' 2017/08/03 M.HONDA
    If cmbBillAddEzappiCode2 = "95" Then
        '' 地方物件の際には10,000円とする(頭二桁が11～19まで)
        If (Left(txtContainerYcode, 2) >= 11 And Left(txtContainerYcode, 2) <= 19) Then
            'txtBillAddEzappiDaily2.VALUE = 10800
            txtBillAddEzappiDaily2.VALUE = 11000  '2019/09/26 M.HONDA
            
        Else
            'txtBillAddEzappiDaily2.VALUE = 2160
            'txtBillAddEzappiDaily2.VALUE = 3240
            txtBillAddEzappiDaily2.VALUE = 3300  '2019/09/26 M.HONDA
            
        End If
        txtBillAddEzappiDaily2.Enabled = False
    Else
        txtBillAddEzappiDaily2.Enabled = True
        txtBillAddEzappiDaily2.VALUE = 0
    End If

    If cmbBillAddEzappiCode2.ListIndex = -1 Then
        '▼2008/12/04 Add hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。
        'del 2018/10/02 EGL -->txtBillSecukg = txtBillSecukg - Nz(txtBillAddEzappi2, 0)
        '▲2008/12/04 Add hirano
        txtBillAddEzappi2 = Null
        txtBillAddEzappi2.Enabled = False
        txtBillAddEzappiDaily2 = Null
    Else
        txtBillAddEzappi2.Enabled = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddEzappiCode2_BeforeUpdate
'        機能             :cmbBillAddEzappiCode2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddEzappiCode2_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbBillAddEzappiCode2, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
    '▼ 2007/08/23 add tajima
    '   ﾈｯﾄ割引暫定処置
    If Me.cmbBillAddEzappiCode2.VALUE = 57 Then
        MsgBox "ネット割引サービスは手動で選択することは出来ません。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
    '▲ 2007/08/23 add tajima
    
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddEzappiCode2_KeyPress
'        機能             :cmbBillAddEzappiCode2_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddEzappiCode2_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbBillAddEzappiCode2, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddFzappiCode1_AfterUpdate
'        機能             :cmbBillAddFzappiCode1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddFzappiCode1_AfterUpdate()

    txtBillAddFzappiName1 = fncAfterUpdateCommon(cmbBillAddFzappiCode1, P_DIVISION_LENGTH, P_CONTROL_数値)
    
    If cmbBillAddFzappiCode1.ListIndex = -1 Then
        txtBillAddFzappi1 = Null
        txtBillAddFzappi1.Enabled = False
    Else
        txtBillAddFzappi1.Enabled = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddFzappiCode1_BeforeUpdate
'        機能             :cmbBillAddFzappiCode1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddFzappiCode1_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbBillAddFzappiCode1, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
    '▼ 2007/08/23 add tajima
    '   ﾈｯﾄ割引暫定処置
    If Me.cmbBillAddFzappiCode1.VALUE = 57 Then
        MsgBox "ネット割引サービスは手動で選択することは出来ません。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
    '▲ 2007/08/23 add tajima
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddFzappiCode1_KeyPress
'        機能             :cmbBillAddFzappiCode1_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddFzappiCode1_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbBillAddFzappiCode1, KeyAscii, P_DIVISION_LENGTH)

End Sub

'2006/06/01 SHIBAZAKI セキュリティカード対応↓
''==============================================================================*
''
''        MODULE_NAME      :cmbBillAddFzappiCode2_AfterUpdate
''        機能             :cmbBillAddFzappiCode2_AfterUpdate イベント
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmbBillAddFzappiCode2_AfterUpdate()
'
'    txtBillAddFzappiName2 = fncAfterUpdateCommon(cmbBillAddFzappiCode2, P_DIVISION_LENGTH, P_CONTROL_数値)
'
'    If cmbBillAddFzappiCode2.ListIndex = -1 Then
'        txtBillAddFzappi2 = Null
'        txtBillAddFzappi2.Enabled = False
'    Else
'        txtBillAddFzappi2.Enabled = True
'    End If
'
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :cmbBillAddFzappiCode2_BeforeUpdate
''        機能             :cmbBillAddFzappiCode2_BeforeUpdate イベント
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmbBillAddFzappiCode2_BeforeUpdate(Cancel As Integer)
'
'    Call fncBeforeUpdateCommon(cmbBillAddFzappiCode2, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
'
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :cmbBillAddFzappiCode2_KeyPress
''        機能             :cmbBillAddFzappiCode2_KeyPress イベント
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub cmbBillAddFzappiCode2_KeyPress(KeyAscii As Integer)
'
'    Call CmKeyPress_For_Num1(cmbBillAddFzappiCode2, KeyAscii, P_DIVISION_LENGTH)
'
'End Sub
'2006/06/01 SHIBAZAKI セキュリティカード対応↑

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddEzappiCode1_AfterUpdate
'        機能             :cmbBillAddEzappiCode1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddEzappiCode1_AfterUpdate()

    txtBillAddEzappiName1 = fncAfterUpdateCommon(cmbBillAddEzappiCode1, P_DIVISION_LENGTH, P_CONTROL_数値)

    If cmbBillAddEzappiCode1.ListIndex = -1 Then
        txtBillAddEzappi1 = Null
        txtBillAddEzappi1.Enabled = False
        txtBillAddEzappiDaily1 = Null
        txtBillAddEzappiDaily1.Enabled = False
    Else
        txtBillAddEzappi1.Enabled = True
        txtBillAddEzappiDaily1.Enabled = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddEzappiCode1_BeforeUpdate
'        機能             :cmbBillAddEzappiCode1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddEzappiCode1_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbBillAddEzappiCode1, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
    '▼ 2007/08/23 add tajima
    '   ﾈｯﾄ割引暫定処置
    If Me.cmbBillAddEzappiCode1.VALUE = 57 Then
        MsgBox "ネット割引サービスは手動で選択することは出来ません。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
    End If
    '▲ 2007/08/23 add tajima
    
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillAddEzappiCode1_KeyPress
'        機能             :cmbBillAddEzappiCode1_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillAddEzappiCode1_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbBillAddEzappiCode1, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillSeikyuCd_GotFocus
'        機能             :cmbBillSeikyuCd_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillSeikyuCd_GotFocus()
    blncmbBillSeikyuCd = True   'カーソルON
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillSeikyuCd_LostFocus
'        機能             :cmbBillSeikyuCd_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/01 追加
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillSeikyuCd_LostFocus()
    blncmbBillSeikyuCd = False   'カーソルOFF
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyicd_GotFocus
'        機能             :cmbContainerHosyicd_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/01 追加
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyicd_GotFocus()
    blncmbContainerHosyicd = True  'カーソルON
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyicd_LostFocus
'        機能             :cmbContainerHosyicd_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/01 追加
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyicd_LostFocus()
    blncmbContainerHosyicd = False  'カーソルOFF
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyoCompany_AfterUpdate
'        機能             :cmbContainerHosyoCompany_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/01 hirano 保証会社コード追加
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyoCompany_AfterUpdate()
    txtContainerHosyoCampany = fncAfterUpdateCommon(cmbContainerHosyoCompany, P_CODE_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyicd_BeforeUpdate
'        機能             :cmbContainerHosyicd_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/01 hirano 保証会社コード追加
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyoCompany_BeforeUpdate(Cancel As Integer)
    Call fncBeforeUpdateCommon(cmbContainerHosyoCompany, P_CODE_LENGTH, P_CONTROL_数値, Cancel)
'▼ 2011/06/18 tajima 保証会社変更は可能にした
    '2009/04/23 ADD<S> hirano 保証会社コード追加　ネット予約は変更不可
'    If blncmbContainerHosyoCompany Then
'        '画面で変更した場合
'        If blnネット予約 Then
'            'ネット予約
'            If strHosyo_cd_Before = cmbContainerHosyoCompany Then
'                Cancel = True
'            Else
'                MsgBox "変更できません。", vbInformation, P_KAKUNIN
'                Cancel = True
'                SendKeys "{ESC}"
'            End If
'        End If
'    End If
    '2009/04/23 INS <E>
' 但し電話予約の場合はクレデンスは選択出来ないようにする

' 2012/10/25 M.HONDA START
'    If blncmbContainerHosyoCompany Then
'        '画面で変更した場合
'        If blnネット予約 = False Then
'            '電話予約の場合クレデンスへは変更不可
'            If cmbContainerHosyoCompany = "970020" Then
'                MsgBox "電話予約の場合、クレデンスへは変更できません。", vbInformation, P_KAKUNIN
'                Cancel = True
'                SendKeys "{ESC}"
'            End If
'        End If
'    End If
''▲ 2011/06/18 tajima
'' 2012/10/25 M.HONDA END

    '' 2013/06/04 M.HONDA START
    If cmbContainerHosyoCompany = "970060" Then
        MsgBox "RMトラストは利用できません。", vbInformation, P_KAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If
    '' 2013/06/04 M.HONDA END

    '2017/12/20 add EGL
    Call subHoshouShow(1)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyoCompany_Change
'        機能             :cmbContainerHosyoCompany_Change イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
Private Sub cmbContainerHosyoCompany_Change()

    Dim SEQ As Variant
    'If txtRcptNo.VALUE = "" Then                                               'DELETE 2022/02/14 N.IMAI
    If txtRcptNo.VALUE = "" And txtMoveKbn = "0" Then                           'INSERT 2022/02/14 N.IMAI
        SEQ = fncGETHOSESEQ(cmbContainerHosyoCompany.VALUE)
        If SEQ = "0" Then
            txtSecurityHosyB.Enabled = True
        Else
            txtSecurityHosyB.Enabled = False
            txtSecurityHosyB.VALUE = SEQ
        End If
    End If
        
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyoCompany_GotFocus
'        機能             :cmbContainerHosyoCompany_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/23追加
Private Sub cmbContainerHosyoCompany_GotFocus()
    blncmbContainerHosyoCompany = True 'カーソルON
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyoCompany_KeyPress
'        機能             :cmbContainerHosyoCompany_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/01 hirano 保証会社コード追加
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyoCompany_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(cmbContainerHosyoCompany, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyoCompany_LostFocus
'        機能             :cmbContainerHosyoCompany_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'2009/04/23 追加
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyoCompany_LostFocus()
    blncmbContainerHosyoCompany = False 'カーソルOFF
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbJIRE_AfterUpdate
'        機能             :cmbJIRE_AfterUpdate イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbJIREI_AfterUpdate()

   txtJIREI.VALUE = fncAfterUpdateCommon(cmbJIREI, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbJIREI_BeforeUpdate
'        機能             :cmbJIREI_BeforeUpdate イベント
'        2015/03/24
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbJIREI_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbJIREI, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbPage_AfterUpdate
'        機能             :cmbPage_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbPage_AfterUpdate()

    If cmbPage.ListIndex = -1 Then
        Exit Sub
    End If
    
    '指定されたラベルを先頭に持っていく
    Call GoToPage(1, , varPageTop(cmbPage.ListIndex))
    
    '参照モード時はフォーカス設定はしない
    If varSetFocus(cmbPage.ListIndex).Enabled Then
        'フォーカスを設定する
        varSetFocus(cmbPage.ListIndex).SetFocus
        SendKeys "{TAB}"
    End If
    
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 毎月回収方法コンボ AfterUpdate
'       MODULE_ID       : cmbRcptKaihi_AfterUpdate
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptKaihi_AfterUpdate()
    On Error GoTo ErrorHandler
    txtRcptKaihi.VALUE = fncAfterUpdateCommon(cmbRcptKaihi, 1, P_CONTROL_数値)
    
    '▼ 2011/06/18 tajima
    'ネット予約且つクレジットカードならば保証会社を変える
'2020/02/25 UPDATE If blnネット予約 = True And Nz(cmbRcptKaihi, "") = "2" And strネット予約初期保証会社Cd <> "" Then
    If blnネット予約 = True And Nz(cmbRcptKaihi, "") = "2" And strネット予約初期保証会社Cd <> "" And cmbContainerHosyicd <> 2 Then
        cmbContainerHosyoCompany = strネット予約初期保証会社Cd
        Call cmbContainerHosyoCompany_AfterUpdate
    End If
    '▲ 2011/06/18 tajima
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmbRcptKaihi_AfterUpdate", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 毎月回収方法コンボ BeforeUpdate
'       MODULE_ID       : cmbRcptKaihi_BeforeUpdate
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptKaihi_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    Call fncBeforeUpdateCommon(cmbRcptKaihi, 1, P_CONTROL_数値, Cancel)
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmbRcptKaihi_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 毎月回収方法コンボ KeyPress
'       MODULE_ID       : cmbRcptKaihi_KeyPress
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'       PARAM           : KeyAscii              アスキーコード(I/O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptKaihi_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler

    Call CmKeyPress_For_Num1(cmbRcptKaihi, KeyAscii, 1)
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmbRcptKaihi_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : リストにない区分を選択したとき
'       MODULE_ID       : cmbRcptKaihi_NotInList
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'       PARAM           : NewData               新データ(I/O)
'                       : Response              レスポンス(I/O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptKaihi_NotInList(NewData As String, Response As Integer)
    On Error GoTo ErrorHandler
    
    MsgBox "リストから選択してください", vbOKOnly + vbInformation, Me.Caption
    Response = acDataErrContinue
    cmbRcptKaihi.Undo
Exit Sub
    
ErrorHandler:       '↓自分の関数名
    Call MSZZ024_M00("cmbRcptKaihi_NotInList", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 請求書発行区分コンボ AfterUpdate
'       MODULE_ID       : cmbRcptSeiki_AfterUpdate
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptSeiki_AfterUpdate()
    On Error GoTo ErrorHandler

    txtRcptSeiki.VALUE = fncAfterUpdateCommon(cmbRcptSeiki, 1, P_CONTROL_数値)
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmbRcptSeiki_AfterUpdate", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 請求書発行区分コンボ BeforeUpdate
'       MODULE_ID       : cmbRcptSeiki_BeforeUpdate
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptSeiki_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    Call fncBeforeUpdateCommon(cmbRcptSeiki, 1, P_CONTROL_数値, Cancel)
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmbRcptSeiki_BeforeUpdate", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 請求書発行区分コンボ KeyPress
'       MODULE_ID       : cmbRcptSeiki_KeyPress
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'       PARAM           : KeyAscii              アスキーコード(I/O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptSeiki_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorHandler
    
    Call CmKeyPress_For_Num1(cmbRcptSeiki, KeyAscii, 1)
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmbRcptSeiki_KeyPress", True)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : リストにない区分を選択したとき
'       MODULE_ID       : cmbRcptSeiki_NotInList
'       CREATE_DATE     : 2010/03/29            K.ISHIZAKA
'       PARAM           : NewData               新データ(I/O)
'                       : Response              レスポンス(I/O)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbRcptSeiki_NotInList(NewData As String, Response As Integer)
    On Error GoTo ErrorHandler
    
    MsgBox "リストから選択してください", vbOKOnly + vbInformation, Me.Caption
    Response = acDataErrContinue
    cmbRcptSeiki.Undo
Exit Sub
    
ErrorHandler:       '↓自分の関数名
    Call MSZZ024_M00("cmbRcptSeiki_NotInList", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSEIKI_AfterUpdate
'        機能             :cmbSEIKI_AfterUpdate イベント
'        2015/03/24
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSEIKI_AfterUpdate()

   txtnenSEIKI.VALUE = fncAfterUpdateCommon(cmbSEIKI, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSEIKI_BeforeUpdate
'        機能             :cmbSEIKI_BeforeUpdate イベント
'        2015/03/24
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSEIKI_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbSEIKI, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserCode_AfterUpdate
'        機能             :cmbUserCode_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserCode_AfterUpdate()
On Error GoTo ErrorHandler

    Call fncAfterUpdateCommon(cmbUserCode, P_CODE_LENGTH, P_CONTROL_数値)

    If cmbUserCode = "" Then
        GoTo ExitRtn
    End If

    If Not fncUserRefresh Then
        '該当顧客無し
        cmbUserCode = strUserCodeOldValue
        cmdEnd.SetFocus
        cmbUserCode.SetFocus
        Exit Sub
    End If

    cmbUserCode.Requery

    GoTo ExitRtn

ErrorHandler:
    Call MSZZ024_M00("cmbUserCode_AfterUpdate", True)
    Err.Clear

ExitRtn:
    strUserCodeOldValue = cmbUserCode

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserCode_BeforeUpdate
'        機能             :cmbUserCode_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserCode_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbUserCode, P_CODE_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserCode_KeyPress
'        機能             :cmbUserCode_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserCode_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbUserCode, KeyAscii, P_CODE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserKkbn_AfterUpdate
'        機能             :cmbUserKkbn_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserKkbn_AfterUpdate()

    txtUserKkbn = fncAfterUpdateCommon(cmbUserKkbn, P_DIVISION_LENGTH, P_CONTROL_数値)
    If cmbUserKkbn.VALUE = P_USER_KKBN_法人 Then
        txtUserKkbn.ForeColor = P_RED
    Else
        txtUserKkbn.ForeColor = P_BLACK
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserKkbn_BeforeUpdate
'        機能             :cmbUserKkbn_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserKkbn_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbUserKkbn, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserKkbn_KeyPress
'        機能             :cmbUserKkbn_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserKkbn_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbUserKkbn, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserSkbn_AfterUpdate
'        機能             :cmbUserSkbn_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserSkbn_AfterUpdate()
    
    txtUserSkbn = fncAfterUpdateCommon(cmbUserSkbn, P_DIVISION_LENGTH, P_CONTROL_数値)
    
    'コンビニならば未請求にする                                                 'INSERT START 2010/03/29 K.ISHIZAKA
    If Nz(cmbUserSkbn.VALUE) = P_USER_SKBN_コンビニ Then
        cmbRcptSeiki.VALUE = P_SEIKI_未発行
        Call cmbRcptSeiki_AfterUpdate
    End If                                                                      'INSERT END   2010/03/29 K.ISHIZAKA

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserSkbn_BeforeUpdate
'        機能             :cmbUserSkbn_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserSkbn_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbUserSkbn, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbUserSkbn_KeyPress
'        機能             :cmbUserSkbn_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbUserSkbn_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbUserSkbn, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbAppFormSendCd_AfterUpdate
'        機能             :cmbAppFormSendCd_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbAppFormSendCd_AfterUpdate()

    txtAppFormSendCd = fncAfterUpdateCommon(cmbAppFormSendCd, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbAppFormSendCd_BeforeUpdate
'        機能             :cmbAppFormSendCd_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbAppFormSendCd_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbAppFormSendCd, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbAppFormSendCd_KeyPress
'        機能             :cmbAppFormSendCd_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbAppFormSendCd_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbAppFormSendCd, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbAppFormButu_AfterUpdate
'        機能             :cmbAppFormButu_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbAppFormButu_AfterUpdate()

    txtAppFormButu = fncAfterUpdateCommon(cmbAppFormButu, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbAppFormButu_BeforeUpdate
'        機能             :cmbAppFormButu_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbAppFormButu_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbAppFormButu, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbAppFormButu_KeyPress
'        機能             :cmbAppFormButu_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbAppFormButu_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbAppFormButu, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillMituCd_AfterUpdate
'        機能             :cmbBillMituCd_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillMituCd_AfterUpdate()

    txtBillMituCd = fncAfterUpdateCommon(cmbBillMituCd, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillMituCd_BeforeUpdate
'        機能             :cmbBillMituCd_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillMituCd_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbBillMituCd, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillMituCd_KeyPress
'        機能             :cmbBillMituCd_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillMituCd_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbBillMituCd, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillSeikyuCd_AfterUpdate
'        機能             :cmbBillSeikyuCd_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillSeikyuCd_AfterUpdate()

    txtBillSeikyuCd = fncAfterUpdateCommon(cmbBillSeikyuCd, P_DIVISION_LENGTH, P_CONTROL_数値)
    
    txtSendSlipNo.Enabled = (cmbBillSeikyuCd = P_SEIKYUCD_代引き)
    txtSendArrivalDate.Enabled = txtSendSlipNo.Enabled
    If txtSendSlipNo.Enabled Then
        Call YellowControl(txtReceiptPaydate, False)
        txtReceiptPaydate = txtSendArrivalDate
    Else
        Call YellowControl(txtReceiptPaydate, False)
        txtSendSlipNo = ""
        txtSendArrivalDate = ""
    End If
    '2009/04/01 ADD <S> hirano カーソルONの場合、見積もり請求区分のデフォルトセット
    If blncmbBillSeikyuCd Then
        Select Case cmbBillSeikyuCd
        Case P_SEIKYUCD_代引き
            cmbBillMituCd = P_MITSUCD_明細書
        Case P_SEIKYUCD_振込み, P_SEIKYUCD_来社現金
            cmbBillMituCd = P_MITSUCD_請求書
        End Select
        Call cmbBillMituCd_AfterUpdate
    End If
    '2009/04/01 ADD <E>
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillSeikyuCd_BeforeUpdate
'        機能             :cmbBillSeikyuCd_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillSeikyuCd_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbBillSeikyuCd, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbBillSeikyuCd_KeyPress
'        機能             :cmbBillSeikyuCd_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbBillSeikyuCd_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbBillSeikyuCd, KeyAscii, P_DIVISION_LENGTH)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :cmd_Atena_Click_KeyPress
'        機能             :cmd_Atena_Click_KeyPress イベント
'        IN               :2019/04/08 M.HONDA
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_Atena_Click()

    On Error GoTo Err_Atena
    
    Dim strSQL As String
    Dim dbs             As Database
    
    Set dbs = CurrentDb()
    
    strSQL = "DELETE FROM WK_LQ0052_TEMP; "
    dbs.Execute strSQL, dbFailOnError

    strSQL = ""
    strSQL = strSQL & " INSERT INTO WK_LQ0052_TEMP "
    strSQL = strSQL & " ( "
    strSQL = strSQL & " BUMOC,  "
    strSQL = strSQL & " KOKYC,  "
    strSQL = strSQL & " FLGSEL "
    strSQL = strSQL & ") "
    strSQL = strSQL & "VALUES ("
    strSQL = strSQL & " '" & strBumonCode & "'                          ,"
    strSQL = strSQL & " '" & Format(Me.cmbUserCode, "000000") & "', "
    strSQL = strSQL & " TRUE                                         "
    strSQL = strSQL & ")"
    strSQL = strSQL & ";"
    
    dbs.Execute strSQL, dbFailOnError
    
    
    'ワークテーブルを削除
    strSQL = "DROP TABLE WK_LQ0052_TEMP1 ; "
    dbs.Execute strSQL, dbFailOnError
    

    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,              "
    strSQL = strSQL & "    SEL_KOKY.USER_RNAME &" & "'" & "　様" & "' ,       "
    strSQL = strSQL & "IIF(SEL_KOKY.USER_TANM IS NULL OR SEL_KOKY.USER_TANM = '　' OR SEL_KOKY.USER_TANM = ' '"
    strSQL = strSQL & "   ,SEL_KOKY.USER_NAME &" & "'" & "　様" & "'  "
    strSQL = strSQL & "   ,SEL_KOKY.USER_NAME)                          "
    strSQL = strSQL & " )                                               "
    strSQL = strSQL & "                      AS KOKYN,                  "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> '' , ''  , "
    strSQL = strSQL & "     IIF(SEL_KOKY.USER_TANM IS NULL OR SEL_KOKY.USER_TANM = '　' OR SEL_KOKY.USER_TANM = ' '"
    strSQL = strSQL & "        ,NULL                                    "
    strSQL = strSQL & "        ,SEL_KOKY.USER_TANM &" & "'" & "　様" & " ') "
    strSQL = strSQL & " )                                               "
    strSQL = strSQL & "                      AS TANTN, "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RPOST  , SEL_KOKY.USER_YUBINO ) AS YUBIB, "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RADR_1 , SEL_KOKY.USER_ADR_1  ) AS ADR1N, "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RADR_2 , SEL_KOKY.USER_ADR_2  ) AS ADR2N, "
    strSQL = strSQL & " IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> ''  ,                 "
    strSQL = strSQL & "    SEL_KOKY.USER_RADR_3 , SEL_KOKY.USER_ADR_3  ) AS ADR3N, "
    strSQL = strSQL & " WK_LQ0052_TEMP.BUMOC & WK_LQ0052_TEMP.KOKYC AS PRINC "
    strSQL = strSQL & " ,IIF(TRIM(NZ(SEL_KOKY.USER_RNAME)) <> '', '　様', IIF(SEL_KOKY.USER_TANM IS NULL OR SEL_KOKY.USER_TANM = '　' OR SEL_KOKY.USER_TANM = ' ', '　様', '')) AS KEISN "  'INSERT 2026/04/22 N.IMAI
    strSQL = strSQL & " ,SEL_KOKY.* "                                                   'INSERT 2026/04/22 N.IMAI
    strSQL = strSQL & " INTO "
    strSQL = strSQL & "WK_LQ0052_TEMP1 "
    strSQL = strSQL & " FROM WK_LQ0052_TEMP "
    strSQL = strSQL & " INNER JOIN SEL_KOKY ON WK_LQ0052_TEMP.KOKYC = SEL_KOKY.KOKYC "
    strSQL = strSQL & " WHERE WK_LQ0052_TEMP.FLGSEL = TRUE               "
    strSQL = strSQL & ";"
    
    dbs.Execute strSQL, dbFailOnError

    doCmd.OpenReport "LR0052_M00", acPreview

    dbs.Close
    
    Set dbs = Nothing

Exit Sub


Err_Atena:
    dbs.Close
    Set dbs = Nothing
    Resume Err_Atena
End Sub


'==============================================================================*
'
'        MODULE_NAME      : cmd_RCPT_SHOUMEI_KDATE_Click
'        機能             : 証明書確認
'        IN               :
'        OUT              :
'       CREATE_DATE     : 2015/07/16            K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd_RCPT_SHOUMEI_KDATE_Click()
    Dim strUser             As String
    Dim strDate             As String
    Dim strTime             As String
    Dim strSQL              As String
    Dim lngCnt              As Long
    Dim objCon              As Object                                           'INSERT 2015/07/26 K.ISHIZAKA
    On Error GoTo ErrorHandler
    
    strUser = LsGetUserName()
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhnnss")
    
    strSQL = strSQL & "UPDATE"
    strSQL = strSQL & " RCPT_TRAN "
    strSQL = strSQL & "SET"
    strSQL = strSQL & " RCPT_SHOUMEI_KDATE = CONVERT(DATETIME,'" & strDate & "',112),"
    strSQL = strSQL & " RCPT_UPDAD = '" & strDate & "',"
    strSQL = strSQL & " RCPT_UPDAJ = '" & strTime & "',"
    strSQL = strSQL & " RCPT_UPDPB = '" & Me.NAME & "',"
    strSQL = strSQL & " RCPT_UPDUB = '" & strUser & "' "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " RCPT_NO = '" & txtYoukNo.VALUE & "' "

'    lngCnt = ADODB_Execute(strSQL)                                             'DELETE 2015/07/26 K.ISHIZAKA
    If strBumonCode <> "" Then                                                  'INSERT START 2015/07/26 K.ISHIZAKA
        Call fncGetConnectString
    End If
    Set objCon = ADODB_Connection(strBumonCode)
    On Error GoTo ErrorHandler1
    lngCnt = ADODB_Execute(strSQL, objCon)
    objCon.Close
    On Error GoTo ErrorHandler                                                  'INSERT END   2015/07/26 K.ISHIZAKA
    
    If cmdComplete.Enabled Then                                                 'INSERT 2015/07/26 K.ISHIZAKA
        cmdComplete.SetFocus
    Else                                                                        'INSERT START 2015/07/26 K.ISHIZAKA
        cmdEnd.SetFocus
    End If                                                                      'INSERT END   2015/07/26 K.ISHIZAKA
    cmd_RCPT_SHOUMEI_KDATE.Enabled = False
Exit Sub

ErrorHandler1:                                                                  'INSERT 2015/07/26 K.ISHIZAKA
    objCon.Close                                                                'INSERT 2015/07/26 K.ISHIZAKA
ErrorHandler:
    Call MSZZ024_M00("cmd_RCPT_SHOUMEI_KDATE_Click", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdAddPrint_Click
'        機能             :住所変更届
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdAddPrint_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub

    If fncIsPrinting = False Then Exit Sub

    Call subPrintHelper(P_PRINT_住所変更)
    
    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdCancelPrint_Click", True)
    Err.Clear
  
ExitRtn:
    cmdCancelPrint.SetFocus

End Sub

'' 20101221 M.HONDA
'==============================================================================*
'
'        MODULE_NAME      :cmdBankAccount_Click
'        機能             :口座振替印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdBankAccount_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub

    If fncIsPrinting = False Then Exit Sub

    Call subPrintHelper(P_PRINT_口座)
    
    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdBankAccount_Click", True)
    Err.Clear
  
ExitRtn:
    cmdCancelPrint.SetFocus

End Sub


'2006/06/01 SHIBAZAKI セキュリティカード対応
'==============================================================================*
'
'        MODULE_NAME      :cmdBillSCardMnt_Click
'        機能             :cmdBillSCardMnt_Click イベントハンドラ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdBillSCardMnt_Click()

    Call cmdSCardMnt_Click
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdCancelPrint_Click
'        機能             :解約書印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdCancelPrint_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    If fncIsPrinting = False Then Exit Sub  '2007/01/12 add tajima

    Call subPrintHelper(P_PRINT_解約書)
    
    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdCancelPrint_Click", True)
    Err.Clear
  
ExitRtn:
    cmdCancelPrint.SetFocus

End Sub

'' 20101221 M.HONDA
'==============================================================================*
'
'        MODULE_NAME      :cmdCardPrint_Click
'        機能             :クレジットカード申込書印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdCardPrint_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub

    If fncIsPrinting = False Then Exit Sub

    Call subPrintHelper(P_PRINT_クレジット)
    
    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdCancelPrint_Click", True)
    Err.Clear
  
ExitRtn:
    cmdCardPrint.SetFocus

End Sub

'2006/06/01 SHIBAZAKI 配置図ボタン追加
'==============================================================================*
'
'        MODULE_NAME      :cmdHaichi_Click
'        機能             :配置図
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdHaichi_Click()
On Error GoTo ErrorHandler

    Dim strMapFile      As String

    '[コントロールマスタ].EXCEL情報パス のディレクトリにファイルが存在するか検索する。
    'EXCELのファイル名は "(ヤードコード).XLS"とする。
    strMapFile = Dir(strYardMapPath & Format(txtContainerYcode, "000000") & "*.XLS")

    If strMapFile = "" Then
        MsgBox "ヤード地図ファイルが存在しません。", vbInformation, P_KAKUNIN
    Else
        Call Y_Exec(Me.hwnd, strYardMapPath & strMapFile, "", "", SW_SHOWMAXIMIZED)
    End If

    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdHaichi_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdKeyOpenPrint_Click
'        機能             :鍵の開け方印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdKeyOpenPrint_Click()
On Error GoTo ErrorHandler

    Call subPrintHelper(P_PRINT_鍵の開け方)
    
    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdKeyOpenPrint_Click", True)
    Err.Clear
  
ExitRtn:
    cmdKeyOpenPrint.SetFocus

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdKeyPrint_Click
'        機能             :鍵同封書類印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdKeyPrint_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    If fncIsPrinting = False Then Exit Sub  '2007/01/12 add tajima

    Call subPrintHelper(P_PRINT_鍵書類)
    
    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdKeyPrint_Click", True)
    Err.Clear
  
ExitRtn:
    cmdKeyPrint.SetFocus

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdQRPrint_Click
'        機能             :物件解錠用QR add 2021/06/02
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdQRPrint_Click()
On Error GoTo ErrorHandler
    If cmbKeyKagiicd.Column(0) <> 16 Then
        MsgBox "鍵区分が「SL+QR自動+ダイ自動」以外では「物件開錠用QR」は印刷できません。", vbInformation, P_KAKUNIN
        Exit Sub
    End If
    Call subPrintHelper(P_PRINT_QR)
    GoTo ExitRtn
ErrorHandler:
    Call MSZZ024_M00("cmdKeyPrint_Click", True)
    Err.Clear
ExitRtn:
    cmdKeyPrint.SetFocus
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdRcptCancel_Click
'        機能             :受付取り消し
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdRcptCancel_Click()
On Error GoTo ErrorHandler
    
    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    If cmdSave.Enabled = True Then                                              'INSERT 2021/08/18 N.IMAI
        Call fncDataSave(P_SAVEMODE_取り消し, True)
    End If                                                                      'INSERT 2021/08/18 N.IMAI
    
    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdRcptCancel_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdReCal_Click
'        機能             :cmdReCal_Click
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdReCal_Click()

    '日割計算
    'txtBillFirstkg = fncDailyRate(txtBillRentkg)
    '2019/08/08 M.HONDA INS
    If Replace(txtBillKisanDate, "/", "") <= tax_to Then
        txtBillFirstkg = fncDailyRate(txtBillFirstkg_sub)
    Else
        txtBillFirstkg = fncDailyRate(txtBillRentkg)
    End If
    '2019/08/08 M.HONDA INS
    
    '' 2012/12/27 M.HONDA INS
    '' 毎月雑費が保険料相当額の場合には、毎月雑費を計算させない
    If txtEzappiCode <> "68" Then
        txtBillEzappiDaily = fncDailyRate(txtBillEzappi)
    End If
    '' 2012/12/27 M.HONDA INS
    
    txtBillAddEzappiDaily1 = fncDailyRate(txtBillAddEzappi1)
    txtBillAddEzappiDaily2 = fncDailyRate(txtBillAddEzappi2)
    
    Call subEditMode(True)
    Call sub割引再計算(P_MODE_仮計算)                                           'I:20081001 N.Imai
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncDailyRate
'        機能             :日割計算
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDailyRate(strSource As Variant) As Variant

    Dim dblDailyRate        As Double
    Dim dblTemp             As Double
    Dim intLastOfMonth      As Integer
    
    fncDailyRate = Null
    
    If Nz(txtBillDailyRate) = "" Or Nz(strSource) = "" Then
        Exit Function
    End If
    
    dblTemp = CDbl(Nz(strSource, 0))
    If txtBillDailyRate = P_LABEL_日割りなし Then
        '月額使用料
        fncDailyRate = dblTemp
    Else
        '日割り日数÷起算日の月の最終日の日にち×月額(１０の位以下切捨て)
        '起算日の月の最終日の日にち
        intLastOfMonth = Day(DateSerial(Year(txtBillStDate), Month(txtBillStDate) + 1, 0))
        '2016/12/14 M.HONDA
        'fncDailyRate = Fix((txtBillDailyRate / intLastOfMonth * dblTemp) / 100) * 100
        fncDailyRate = Fix((txtBillDailyRate / intLastOfMonth * dblTemp))
    End If


End Function

'2006/06/01 SHIBAZAKI セキュリティカード対応
'==============================================================================*
'
'        MODULE_NAME      :cmdSCardMnt_Click
'        機能             :cmdSCardMnt_Click イベントハンドラ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdSCardMnt_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    'If blnEditMode Then                                                        'DELETE 2021/08/18 N.IMAI
    If blnEditMode And cmdSave.Enabled = True Then                              'INSERT 2021/08/18 N.IMAI
        '2009/02/23 ADD N.IMAI 割引更新確認メッセージ表示
        '2009/02/25 MOD N.IMAI
        '2009/03/09 MOD hirano  Nz付加
        'If (strSave月額料金 <> txtBillMonthCharge _
        ’Or strSave初回料金 <> txtBillFirstkg
        If (strSave月額料金 <> Nz(txtBillMonthCharge, "") _
        Or strSave初回料金 <> Nz(txtBillFirstkg, "") _
        Or strSave契約開始日 <> txtBillStDate _
        Or strSave解約予定日 <> txtReceiptCyDate _
        Or strSave利用予定期間 <> cmb_利用予定期間 _
        Or strSave請求期間 <> txtBillSeikyuKikan _
        Or strSave保証会社 <> cmbContainerHosyoCompany _
        Or strSave毎月回収方法 <> cmbRcptKaihi _
        Or blnSave割引再計算 <> False _
        Or blnSave割引する <> chk割引有効可否.VALUE) _
        And (chk割引有効可否.VALUE = True) Then
            If MsgBox("キャンペーン割引を再計算後保存します。" + vbCrLf + "再計算によって適用割引が変わる場合があります。" + vbCrLf + "よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
                If Not fncDataSave(P_SAVEMODE_保存, False, False) Then
                    Exit Sub
                End If
            End If
        Else
            If MsgBox("更新しますか？" & vbCrLf & _
                  "※更新しなければセキュリティーカード追加画面を表示できません。", _
                  vbQuestion + vbYesNo) = vbYes _
            Then
                If Not fncDataSave(P_SAVEMODE_保存, False, False) Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
        '2009/02/23 ADD N.IMAI 割引更新確認メッセージ表示
    End If
    
    
    'doCmd.OpenForm P_FORM_セキュリティーカード追加, , , , , acDialog, _
    '2007/02/08 ↑旧↓新 chg tajima
    doCmd.OpenForm P_FORM_セキュリティーカード追加, , , , , , _
            Me.NAME & "," & _
            txtContainerYcode & "," & _
            txtContainerCno & "," & _
            txtAcptNo
    
    Exit Sub
  
ErrorHandler:
    Call MSZZ024_M00("cmdSCardMnt_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdSecuNg_Click
'        機能             :保証会社NG
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdSecuNg_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    If cmdSave.Enabled = True Then                                              'INSERT 2021/08/18 N.IMAI
        Call fncDataSave(P_SAVEMODE_保証会社NG, True)
    End If                                                                      'INSERT 2021/08/18 N.IMAI
    
    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdSecuNg_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdSei2Print_Click
'        機能             :請求書２印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdSei2Print_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    Call subPrintHelper(P_PRINT_請求書2)

    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdSei2Print_Click", True)
    Err.Clear
  
ExitRtn:
    cmdSei2Print.SetFocus

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUkePrint_Click
'        機能             :受付表印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUkePrint_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    Call subPrintHelper(P_PRINT_受付管理表)

    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdUkePrint_Click", True)
    Err.Clear
  
ExitRtn:
    cmdUkePrint.SetFocus

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdMouPrint_Click
'        機能             :申込書印刷
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdMouPrint_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima
    If fncIsPrinting = False Then Exit Sub  '2007/01/12 add tajima

    Call subPrintHelper(P_PRINT_申込書)

    GoTo ExitRtn
  
ErrorHandler:
    Call MSZZ024_M00("cmdMouPrint_Click", True)
    Err.Clear
  
ExitRtn:
    'cmdMouPrint.SetFocus

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subPrintHelper
'        機能             :ヘルパー連動印刷
'        IN               :定数宣言[印刷種別]参照
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subPrintHelper(intPrintKind As Integer)
On Error GoTo ErrorHandler

    Dim oApp                As Object
    Dim oBook               As Object
    Dim strHelperFileName   As String
    Dim varPrintSeets       As Variant
    Dim blnError            As Boolean
    Dim j                   As Integer
    Dim DCNTM_PERIOD        As Integer
    
    blnError = False
    
    '↓INSERT 2008/09/25 SHIBAZAKI
    Select Case intPrintKind
'    Case P_PRINT_申込書, P_PRINT_鍵書類        'DELETE 2008/09/26 SHIBAZAKI
    Case P_PRINT_申込書                         'INSERT 2008/09/26 SHIBAZAKI
        If cmbContainerHosyicd = P_HOSYICD_リプラス Then
            MsgBox "使用できない保証区分です。", vbInformation, P_KAKUNIN
            If cmbContainerHosyicd.Enabled Then
                cmbContainerHosyicd.SetFocus
            End If
            Exit Sub
        End If
    End Select
    
    '2018/06/18 M.HONDA INS
    If txtReceiptYukbnCode <> P_YUKBN_契約 Then
    
        '↑INSERT 2008/09/25 SHIBAZAKI
        'If blnEditMode Then                                                    'DELETE 2021/08/18 N.IMAI
        If blnEditMode And cmdSave.Enabled = True Then                          'INSERT 2021/08/18 N.IMAI
        ''2013/10/25 M.HONDA UPD
            '2009/02/23 ADD N.IMAI 割引更新確認メッセージ表示
            '2009/02/25 MOD N.IMAI
    '        If (strSave月額料金 <> Nz(txtBillMonthCharge, "") _
    '        Or strSave初回料金 <> Nz(txtBillFirstkg, "") _
    '        Or strSave契約開始日 <> txtBillStDate _
    '        Or strSave解約予定日 <> txtReceiptCyDate _
    '        Or strSave利用予定期間 <> cmb_利用予定期間 _
    '        Or strSave請求期間 <> txtBillSeikyuKikan _
    '        Or strSave保証会社 <> cmbContainerHosyoCompany _
    '        Or strSave毎月回収方法 <> cmbRcptKaihi _
    '        Or blnSave割引再計算 <> False _
    '        Or blnSave割引する <> chk割引有効可否.VALUE) _
    '        And (chk割引有効可否.VALUE = True) Then
            If (strSave月額料金 <> Nz(txtBillMonthCharge, "") _
            Or strSave初回料金 <> Nz(txtBillFirstkg, "") _
            Or strSave契約開始日 <> txtBillStDate _
            Or strSave解約予定日 <> txtReceiptCyDate _
            Or strSave利用予定期間 <> cmb_利用予定期間 _
            Or strSave請求期間 <> txtBillSeikyuKikan _
            Or strSave保証会社 <> cmbContainerHosyoCompany _
            Or strSave毎月回収方法 <> cmbRcptKaihi _
            Or blnSave割引再計算 <> False _
            Or strSaveトリプル <> chkCampaign _
            Or blnSave割引する <> chk割引有効可否.VALUE) _
            And (chk割引有効可否.VALUE = True) Then
        ''2013/10/25 M.HONDA UPD
            
                If MsgBox("キャンペーン割引を再計算後保存します。" + vbCrLf + "再計算によって適用割引が変わる場合があります。" + vbCrLf + "よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbYes Then
                    If Not fncDataSave(P_SAVEMODE_保存, False) Then
                        Exit Sub
                    End If
                Else                                                                'I:20090225 N.Imai
                    Exit Sub                                                        'I:20090225 N.Imai
                End If
            Else
                If MsgBox("更新しますか？" & vbCrLf & _
                      "※更新しなければ印刷は実行できません。", _
                      vbQuestion + vbYesNo) = vbYes _
                Then
                    '保存
                    If Not fncDataSave(P_SAVEMODE_保存, False) Then
                        Exit Sub
                End If
                Else
                    Exit Sub
                End If
            End If
            '2009/02/23 ADD N.IMAI 割引更新確認メッセージ表示
        End If
        
    End If
    doCmd.Hourglass True
    
    'INSERT 2023/11/09 N.IMAI Start
    If intPrintKind = P_PRINT_申込書 _
    Or intPrintKind = P_PRINT_請求書2 Then
        Call fncInsertKouzMast
    End If
    'INSERT 2023/11/09 N.IMAI End
     
    '↓INSERT 2007/11/15 SHIBAZAKI
    '統合ヘルパーは共通関数を使用して出力する
    Select Case intPrintKind
    Case P_PRINT_受付管理表
    Case P_PRINT_請求書2
    Case Else
        'UPDATE 2008/09/25 SHIBAZAKI
'        If txtBillSecukgLabel = P_LABEL_保証委託料 Or _
'            MSZZ039.P_USAGE_バーチャルオフィス = Me.cmbContainerUsageCode.Value Then '2008/03/23 add  tajima
        'UPDATE 2008/11/13 SHIBAZAKI
'        If txtBillSecukgLabel = P_LABEL_保証委託料 Or _
'           txtBillSecukgLabel = P_LABEL_事務手数料 Or _
'           MSZZ039.P_USAGE_バーチャルオフィス = Me.cmbContainerUsageCode.Value Then

'-----**********----ryu--------update--------20100109------------************------<s>
''        If txtContainerHosyiFrom = P_HOSYIFROM_保証委託 Or _
''           txtContainerHosyiFrom = P_HOSYIFROM_事務手数料 Or _
''           MSZZ039.P_USAGE_バーチャルオフィス = Me.cmbContainerUsageCode.VALUE Then
           If txtContainerHosyiFrom = P_HOSYIFROM_保証委託 Or _
           txtContainerHosyiFrom = P_HOSYIFROM_事務手数料 Or _
           txtContainerHosyiFrom = P_HOSYIFROM_保証金 Or _
           MSZZ039.P_USAGE_バーチャルオフィス = Me.cmbContainerUsageCode.VALUE Then
'-----**********----ryu--------update--------20100109------------************------<e>

'2008/10/01↓ 割引き対応コメント不足
'            Call MSZZ039.HelperPrintPreview(strBumonCode, Me.txtRcptNo.VALUE, intPrintKind)
            '2009/02/13 MOD <S> hirano 毎月割引額の追加
            'Call MSZZ039.HelperPrintPreview(strBumonCode, Me.txtRcptNo.VALUE, wk割付情報(), intPrintKind)
            
            '--------2010/05/28------update-------ryu--------修正：「txt月額分割引」にNZ判断を追加----
            ''Call MSZZ039.HelperPrintPreview(strBumonCode, Me.txtRcptNo.VALUE, wk割付情報(), intPrintKind, txt月額分割引)
            
            '' 2013/03/30 M.HONDA 割引情報を追加
            '' 割引期間の最大値を取得
            '' 複数の割引が重なるケースがあるので全ての割引を調べる
            DCNTM_PERIOD = 0
            For j = 1 To wk割引情報.件数
                '' 事務手数料割引は除く
                If wk割引情報.TRAN_DATA(j - 1).DCNTM_TYPE <> 10 Then
                    If DCNTM_PERIOD < wk割引情報.TRAN_DATA(j - 1).DCNTM_PERIOD Then
                        DCNTM_PERIOD = wk割引情報.TRAN_DATA(j - 1).DCNTM_PERIOD
                    End If
                End If
            Next
            
            ''Call MSZZ039.HelperPrintPreview(strBumonCode, Me.txtRcptNo.VALUE, wk割付情報(), intPrintKind, Nz(txt月額分割引, 0))
                Call MSZZ039.HelperPrintPreview(strBumonCode, Me.txtRcptNo.VALUE, wk割付情報(), intPrintKind, Nz(txt月額分割引, 0), DCNTM_PERIOD)
            '2009/02/13 MOD <E> hirano
'2008/10/01↑
            GoTo ExitRtn
        End If
    End Select
    '↑INSERT 2007/11/15 SHIBAZAKI
    
    Set oApp = CreateObject("Excel.Application")

    'ヘルパーファイル名作成
    strHelperFileName = fncGetHelperFileName(intPrintKind, varPrintSeets)
    If strHelperFileName = "" Then
        GoTo ExitRtn
    End If

    'strHelperFileName = "\\kasegroup\kasesys3\テンプレート\テスト開発環境\Helper\test臨時請求書.xlsm" 'TODO
    Set oBook = oApp.Workbooks.Open(strHelperFileName)

    Select Case intPrintKind
    Case P_PRINT_受付管理表
        '受付管理表値連動
        Call subSetUkeBook(oBook.Worksheets("管理表"))
    '2005/05/22 SHIBAZAKI 請求書２対応
    Case P_PRINT_請求書2
        oApp.DisplayAlerts = False                                              'INSERT 2023/10/26 N.IMAI
        Call subSetSei2Book(oBook.Worksheets(varPrintSeets(0)))
    Case Else
        ' ▼2007/01/12 tajima
        ' 保証金ならば保証金ヘルパーにて申込書連動とする※従来の動作
        'If txtBillSecukgLabel = P_LABEL_保証金 Then            'DELETE 2008/11/13 SHIBAZAKI
        If txtContainerHosyiFrom = P_HOSYIFROM_保証金 Then      'INSERT 2008/11/13 SHIBAZAKI
            Call subSetMouBook(oBook.Worksheets("基本入力"))
'        ' 保証金でなければ統合ヘルパーへ連動する                                       'DELETE 2007/11/15 SHIBAZAKI
'        Else                                                                           'DELETE 2007/11/15 SHIBAZAKI
'            Call subSetIntegrationBook(intPrintKind, oBook.worksheets("基本入力"))     'DELETE 2007/11/15 SHIBAZAKI
'            Call oBook.sub加瀬ロゴ印刷設定(fncIsRogoPrint)                             'DELETE 2007/11/15 SHIBAZAKI
        End If
        ' ▲2007/01/12 tajima
    End Select


    On Error Resume Next

    oApp.Visible = True

    '2005/05/22 SHIBAZAKI 請求書２対応
    If intPrintKind = P_PRINT_請求書2 Then
        Call oApp.Sheets(varPrintSeets).Select
        
        '----20100109----M.RYU----add---'社印を追加
        Call subInStamp(oBook, oBook.Worksheets(varPrintSeets(0)), P_CELL_各印)
        
        Call oBook.Worksheets(varPrintSeets(0)).Range("B2").Select
    Else
        '印刷プレビュー表示
        Call oApp.Sheets(varPrintSeets).PrintPreview
    End If
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    '2005/05/22 SHIBAZAKI 請求書２対応
    If intPrintKind <> P_PRINT_請求書2 Or blnError Then
        If Not oBook Is Nothing Then Set oBook = Nothing
        If Not oApp Is Nothing Then
            oApp.DisplayAlerts = False
            oApp.Quit
            Set oApp = Nothing
        End If
    End If
    
    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "subPrintHelper" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncGetHelperFileName
'        機能             :ヘルパーファイル名作成
'        IN               :True=受付管理表のファイル名 False=それ以外のファイル名
'        OUT              :ヘルパーファイル名(フルパス)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetHelperFileName(intPrintKind As Integer, ByRef varPrintSeets As Variant) As String

    Dim strFile         As String
    Dim is印刷シート取得済  As Boolean
    
    fncGetHelperFileName = ""
    is印刷シート取得済 = False  ' 2007/01/12 add tajima
    
    If intPrintKind = P_PRINT_受付管理表 Then
        strFile = P_HELPER_受付管理表
    ElseIf intPrintKind = P_PRINT_請求書2 Then      '2005/05/22 SHIBAZAKI 請求書２対応
        strFile = P_HELPER_請求書2
    'ElseIf txtBillSecukgLabel = P_LABEL_保証金 Then            'DELETE 2008/11/13 SHIBAZAKI
    ElseIf txtContainerHosyiFrom = P_HOSYIFROM_保証金 Then      'INSERT 2008/11/13 SHIBAZAKI
        strFile = P_HELPER_保証金用
    Else
' 2007/01/12 Change tajiima
'        strFile = P_HELPER_統合                                        'DELETE 2007/11/15 SHIBAZAKI
'        varPrintSeets = fncGetIntegrationHelperSeets(intPrintKind)     'DELETE 2007/11/15 SHIBAZAKI
'        is印刷シート取得済 = True                                      'DELETE 2007/11/15 SHIBAZAKI

'        Select Case Nz(cmbContainerUsageCode)
'        Case P_USAGE_レンタルオフィス
'            strFile = P_HELPER_オフィス用
'        Case P_USAGE_バイク
'            strFile = P_HELPER_バイク用
'        Case P_USAGE_レンタルボックス
'            '会社コードによっては横浜鋼管用を使用する
'            If strCampCode = P_CAMP_横浜鋼管 Then
'                strFile = P_HELPER_横浜鋼管用
'            Else
'                strFile = P_HELPER_コンテナ用
'            End If
'        Case P_USAGE_トランクルーム
'            strFile = P_HELPER_トランク用
'        Case Else
'            Exit Function
'        End Select
' 2007/01/12 Change End
    
    End If
    
    If Not is印刷シート取得済 Then  ' 2007/01/12 add tajima
      '印刷シート取得
      varPrintSeets = fncGetPrintSeets(intPrintKind, strFile)
    End If                          ' 2007/01/12 add tajima
    
    fncGetHelperFileName = strHelperPath & strFile
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetPrintSeets
'        機能             :印刷対象シート配列取得
'        IN               :intPrintKind 印刷種別
'                         :strHelper ヘルパーファイル名
'        OUT              :印刷対象シート配列
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetPrintSeets(intPrintKind As Integer, strHelper As String) As Variant
    
    Dim strTemp     As String
    Dim varArray    As Variant
    
    Select Case intPrintKind
    Case P_PRINT_受付管理表
        varArray = Array("管理表")
    Case P_PRINT_申込書
        If strHelper = P_HELPER_保証金用 Then
            If strCampCode = P_CAMP_横浜鋼管 Then
                strTemp = "横浜鋼管個人情報"
            Else
                strTemp = "個人情報"
            End If
            varArray = Array("案内", "承諾書", "申込書", "請求書", strTemp, "支払案内")
'↓DELETE 2007/11/15 SHIBAZAKI
'        ElseIf strHelper = P_HELPER_オフィス用 Then
'            If cmbUserKkbn = P_USER_KKBN_法人 Then
'                strTemp = "申込書法人"
'            Else
'                strTemp = "申込書個人"
'            End If
'            varArray = Array("案内１", "マイオフィス一時使用約款", strTemp, "個人情報", "請求書")
'        Else
'            varArray = Array("FAX")
'↑DELETE 2007/11/15 SHIBAZAKI
        End If
        
'        '2006/04/28 S.SHIBAZAKI 乗用具チェック追加。チェックされている場合、申込書に誓約書添付
'        Select Case strHelper
'        Case P_HELPER_コンテナ用, P_HELPER_トランク用, P_HELPER_横浜鋼管用
'            If Nz(chkAppFormVehicle.Value, 0) = -1 Then
'                ReDim Preserve varArray(UBound(varArray) + 1)
'                varArray(UBound(varArray)) = "誓約書"
'            End If
'        End Select
        
        
        '解約日が入ってる場合は解約書を同時に印刷
        If Nz(txtReceiptCyDate) <> "" Then
            ReDim Preserve varArray(UBound(varArray) + 1)
            varArray(UBound(varArray)) = "解約"
        End If
'↓DELETE 2007/11/15 SHIBAZAKI
'    Case P_PRINT_解約書
'        varArray = Array("解約")
'    Case P_PRINT_鍵書類
'        If strHelper = P_HELPER_保証金用 Then
'            '2006/04/28 S.SHIBAZAKI 解約日が入っている場合は、鍵同封書類に解約通知を含めない。
'            'varArray = Array("経理用", "案内２", "月々の支払い", "住所変更", "解約")
'            varArray = Array("経理用", "案内２", "月々の支払い", "住所変更")
'        ElseIf strHelper = P_HELPER_オフィス用 Then
'            '2006/04/28 S.SHIBAZAKI 解約日が入っている場合は、鍵同封書類に解約通知を含めない。
'            'varArray = Array("経理用", "案内２", "月々の支払い", "住所変更", "解約")
'            varArray = Array("経理用", "案内２", "月々の支払い", "住所変更")
'        Else
'            '2006/04/28 S.SHIBAZAKI 解約日が入っている場合は、鍵同封書類に解約通知を含めない。
'            'varArray = Array("請求書控え", "経理用", "案内２", "月々の支払い", "住所変更", "解約")
'            varArray = Array("請求書控え", "経理用", "案内２", "月々の支払い", "住所変更")
'        End If
'
'        '2006/04/28 S.SHIBAZAKI 解約日が入っている場合は、鍵同封書類に解約通知を含めない。
'        If Nz(txtReceiptCyDate) = "" Then
'            '解約日が「入っていない」場合は解約書を同時に印刷
'            ReDim Preserve varArray(UBound(varArray) + 1)
'            varArray(UBound(varArray)) = "解約"
'        End If
'↑DELETE 2007/11/15 SHIBAZAKI
    Case P_PRINT_請求書2
        '2005/05/22 SHIBAZAKI 請求書２対応
'        If strHelper = P_HELPER_保証金用 Then
'            varArray = Array("請求書")
'        Else
            varArray = Array("請求書2")
'        End If
'    Case P_PRINT_鍵の開け方
'        Select Case strHelper
'        Case P_HELPER_コンテナ用, P_HELPER_横浜鋼管用
'            fncGetPrintSeets = Array("ドアキーの開け方")
'        Case P_HELPER_トランク用
'            fncGetPrintSeets = Array("セキュリティ")
'        Case P_HELPER_バイク用
'            fncGetPrintSeets = Array("中原田尻セキュリティ")    'TODO
'        Case P_HELPER_保証金用
'            fncGetPrintSeets = Array("鍵の開け方")
'        End Select
    End Select

    fncGetPrintSeets = varArray
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subSetMouBook
'        機能             :申込書値連動
'        IN               :値設定するシート
'        OUT              :
'
'        MEMO             :2007/11/15 SHIBAZAKI記
'                          保証金用のみここで出力
'                          保証委託料のヘルパーは共通関数で出力
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetMouBook(objSeet As Object)

    '申込書値連動
    With objSeet
        'If txtBillSecukgLabel = P_LABEL_保証金 Then            'DELETE 2008/11/13 SHIBAZAKI
        If txtContainerHosyiFrom = P_HOSYIFROM_保証金 Then      'INSERT 2008/11/13 SHIBAZAKI
            '保証金用のみ
            .Range("申込日").VALUE = txtReceiptUkdate
            .Range("保証金").VALUE = Nz(txtBillSecukg, 0)
            .Range("会社コード").VALUE = strCampCode
            .Range("用途").VALUE = Nz(txtContainerUsage)
'↓DELETE 2007/11/15 SHIBAZAKI
'        Else
'            '保証金用以外
'            .Range("受付日").Value = txtReceiptUkdate
'
'            If Nz(txtBillMituCd) <> "" Then
'                .Range("見積請求").Value = Nz(txtBillMituCd, "")
'            End If
'
'            .Range("初回支払").Value = txtBillSeikyuCd
'
'            'オフィス用と保証金用以外
'            If Nz(cmbContainerUsageCode) <> P_USAGE_レンタルオフィス Then
'                .Range("契約期間").Value = "一般"
'
'                .Range("保証料").Value = Nz(txtBillSecukg, 0)
'
'                .Range("送付方法").Value = Nz(txtAppFormSendCd, "")
'                .Range("収納物").Value = Nz(txtAppFormButu, "")
'
'                .Range("契約代表者名").Value = IIf(Trim(Nz(txtUserRname)) = "", Trim(Nz(txtUserTanm, "")), Trim(Nz(txtUserRname, "")))
'                .Range("備考").Value = Nz(txtReceiptUknote, "")
'                .Range("起算加算日数").Value = "8"  'TODO 一部の文言に影響あり　確認中
'                .Range("カギ本数").Value = Nz(txtKeyKagiLendnum, "")
'                .Range("カードNo").Value = Nz(txtKeyKagiNo, "")
'
'                .Range("敬称2").Value = "様"
'            End If
'↑DELETE 2007/11/15 SHIBAZAKI
        End If
        
        'オフィス用以外
        If Nz(cmbContainerUsageCode) <> P_USAGE_レンタルオフィス Then
            .Range("上下").VALUE = Nz(txtContainerStep, "")
            .Range("個人法人").VALUE = txtUserKkbn
            .Range("見積り有効期限").VALUE = txtAppFormLimitDate
        End If
        
        .Range("担当").VALUE = txtReceiptUktanto
        .Range("受付SYSNo").VALUE = txtRcptNo
        
        If Nz(txtBillSeikyuKikan) = "" Then
            '請求期間が空ならば、起算日を格納する
            .Range("支払済み月日").VALUE = txtBillKisanDate
        Else
            .Range("支払済み月日").VALUE = fncFirstDay(txtBillSeikyuKikan)
        End If
        
        .Range("ヤード番号").VALUE = txtContainerYcode
        .Range("ヤード名").VALUE = txtContainerYname
        .Range("所在地").VALUE = Nz(txtContainerYadd, "")

        .Range("ボックス番号").VALUE = txtContainerCno
        .Range("広さ").VALUE = Nz(txtContainerSize, "")
        .Range("起算日").VALUE = txtBillKisanDate

        .Range("日割り").VALUE = txtBillDailyRate
        .Range("解約日").VALUE = txtReceiptCyDate
        If Nz(txtReceiptCyDate) <> "" Then
            .Range("解約予定").VALUE = P_HELPER_CANCEL_STRING
        End If
        
        .Range("使用料").VALUE = txtBillMonthCharge
        .Range("鍵種別").VALUE = txtKeyKagiicd
        
        .Range("顧客コード").VALUE = cmbUserCode
        .Range("フリガナ").VALUE = txtUserKana
        .Range("契約者名").VALUE = txtUserName
        
        '顧客敬称 法人なら｢御中」　個人なら「様」
        If cmbUserKkbn = P_USER_KKBN_法人 Then
            .Range("敬称").VALUE = "御中"
        Else
            .Range("敬称").VALUE = "様"
        End If
        
        If Trim(Nz(txtUserRname)) = "" Then
            '顧客情報を適用
            .Range("郵便番号").VALUE = txtUserYubino
            .Range("住所1").VALUE = txtUserAdr1
            .Range("住所2").VALUE = txtUserAdr2 & "　" & Nz(txtUserAdr3, "")
            .Range("TEL").VALUE = txtUserTel
            .Range("FAX").VALUE = Nz(txtUserFax, "")
            .Range("携帯").VALUE = Nz(txtUserKeitai, "")
        Else
            '書類送付先を適用
            .Range("郵便番号").VALUE = txtUserRpost
            .Range("住所1").VALUE = txtUserRadr1
            .Range("住所2").VALUE = txtUserRadr2 & "　" & Nz(txtUserRadr3, "")
            .Range("TEL").VALUE = txtUserRtel
            .Range("FAX").VALUE = Nz(txtUserRfax, "")
            .Range("携帯").VALUE = Nz(txtUserRkeitai, "")
        End If
        
        '2006/05/22 SHIBAZAKI 申込書に契約開始日を印字する
        .Range("契約開始日").VALUE = Nz(txtBillStDate, "")
    
    End With
    
    '申込書明細設定
    Call subSetMousikomiMeisai(objSeet)
    
    '請求明細設定
    Call subSetSeikyuMeisai(objSeet)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetMousikomiMeisai
'        機能             :申込書　申込書明細設定
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetMousikomiMeisai(objSeet As Object)

    Dim intCount        As Integer
    
    With objSeet
        .Range("月額使用料") = txtBillRentkgKei
        .Range("初回使用料") = txtBillFirstkgKei
        intCount = 0
        '毎月雑費
        Call subSetZappi(objSeet, intCount, Nz(txtBillEzappiName), fncZeroByteToNull(txtBillEzappi), "毎月雑費名", "毎月雑費")
        '追加毎月雑費１
        Call subSetZappi(objSeet, intCount, Nz(txtBillAddEzappiName1), fncZeroByteToNull(txtBillAddEzappi1), "毎月雑費名", "毎月雑費")
        '追加毎月雑費２
        Call subSetZappi(objSeet, intCount, Nz(txtBillAddEzappiName2), fncZeroByteToNull(txtBillAddEzappi2), "毎月雑費名", "毎月雑費")
        
        intCount = 0
        '初回雑費
        Call subSetZappi(objSeet, intCount, Nz(txtBillFzappiName), fncZeroByteToNull(txtBillFzappi), "初回雑費名", "初回雑費")
        '追加初回雑費１
        Call subSetZappi(objSeet, intCount, Nz(txtBillAddFzappiName1), fncZeroByteToNull(txtBillAddFzappi1), "初回雑費名", "初回雑費")
        '追加初回雑費２
        '2006/06/01 SHIBAZAKI セキュリティカード対応
'        Call subSetZappi(objSeet, intCount, Nz(txtBillAddFzappiName2), fncZeroByteToNull(txtBillAddFzappi2), "初回雑費名", "初回雑費")
        Call subSetZappi(objSeet, intCount, lblBillAddFzappiName2.Caption, fncZeroByteToNull(txtBillAddFzappi2), "初回雑費名", "初回雑費")
    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetSeikyuMeisai
'        機能             :申込書　請求明細設定
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetSeikyuMeisai(objSeet As Object)

    Dim intCount    As Integer
    Dim strTemp     As String
    Dim blnNetNew   As Boolean '2017/12/20 ADD EGL
    
    '2006/05/10 SHIBAZAKI 明細書に日割日数を印字する
    Dim strDaily    As String
    
    intCount = 1
    
    '2017/12/20 ADD EGL ↓↓↓
    If blnネット予約 = True Then
        'ネット契約
        If FVS400_fncChkHoshoNewOld(strBumonCode, cmbContainerHosyoCompany) = False Then
            '新保証会社
            blnNetNew = True
        End If
    End If
    '2017/12/20 ADD EGL ↑↑↑
    
    With objSeet
        strTemp = Format(txtBillKisanDate, "m") & "月分"
    
        '2006/05/10 SHIBAZAKI 明細書に日割日数を印字する
        strDaily = Format(txtBillKisanDate, " ( m月d日 より ")
        If txtBillDailyRate = P_LABEL_日割りなし Then
            strDaily = strDaily & "１ヶ月分 )"
        Else
            strDaily = strDaily & _
                       Day(DateSerial(Year(txtBillKisanDate), Month(txtBillKisanDate) + 1, 0)) & _
                       "日割り " & _
                       txtBillDailyRate & _
                       "日分 )"
        End If
        
        '初月使用料
        '2006/05/10 SHIBAZAKI 明細書に日割日数を印字する
'        .Range("明細1") = strTemp & "使用料"
        .Range("明細1") = strTemp & "使用料" & strDaily
        .Range("明細金額1") = txtBillFirstkgKei
        
        '初回雑費
        Call subSetZappi(objSeet, intCount, Nz(txtBillFzappiName), fncZeroByteToNull(txtBillFzappi), "明細", "明細金額")
        '追加初回雑費１
        Call subSetZappi(objSeet, intCount, Nz(txtBillAddFzappiName1), fncZeroByteToNull(txtBillAddFzappi1), "明細", "明細金額")
        '追加初回雑費２
        '2006/06/01 SHIBAZAKI セキュリティカード対応
'        Call subSetZappi(objSeet, intCount, Nz(txtBillAddFzappiName2), fncZeroByteToNull(txtBillAddFzappi2), "明細", "明細金額")
        Call subSetZappi(objSeet, intCount, lblBillAddFzappiName2.Caption, fncZeroByteToNull(txtBillAddFzappi2), "明細", "明細金額")
    
        '日割毎月雑費
        '2006/05/10 SHIBAZAKI 明細書に日割日数を印字する
'        Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillEzappiName), fncZeroByteToNull(txtBillEzappiDaily), "明細", "明細金額")
        Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillEzappiName) & strDaily, fncZeroByteToNull(txtBillEzappiDaily), "明細", "明細金額")
        '日割追加毎月雑費１
        '2006/05/10 SHIBAZAKI 明細書に日割日数を印字する
'        Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillAddEzappiName1), fncZeroByteToNull(txtBillAddEzappiDaily1), "明細", "明細金額")
        Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillAddEzappiName1) & strDaily, fncZeroByteToNull(txtBillAddEzappiDaily1), "明細", "明細金額")
        '日割追加毎月雑費２
        '2006/05/10 SHIBAZAKI 明細書に日割日数を印字する
'        Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillAddEzappiName2), fncZeroByteToNull(txtBillAddEzappiDaily2), "明細", "明細金額")
        Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillAddEzappiName2) & strDaily, fncZeroByteToNull(txtBillAddEzappiDaily2), "明細", "明細金額")
        
        '前払い分
        If txtBillPremthSum > 0 Then
            strTemp = Format(DateSerial(Format(txtBillKisanDate, "yyyy"), Format(txtBillKisanDate, "m") + 1, 1), "yyyy年m月")
            
            '2006/04/28 S.SHIBAZAKI 前払い月数が一ヶ月の場合に"yyyy年mm月～yyyy年mm月"と明細書に出さないようにする
            If txtBillPremthSum >= 2 Then
                strTemp = strTemp & "～"
                strTemp = strTemp & Format(fncFirstDay(txtBillSeikyuKikan), "yyyy年m月")
                strTemp = strTemp & "(" & txtBillPremthSum & "ヶ月分)"
            Else
                strTemp = strTemp & "分"
            End If
            
            intCount = intCount + 1
            .Range("明細" & intCount) = strTemp & "使用料"
            .Range("明細金額" & intCount) = txtBillRentkgKei * txtBillPremthSum
            
            '毎月雑費
            Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillEzappiName), fncZeroByteToNull(txtBillEzappi) * txtBillPremthSum, "明細", "明細金額")
            '追加毎月雑費１
            Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillAddEzappiName1), fncZeroByteToNull(txtBillAddEzappi1) * txtBillPremthSum, "明細", "明細金額")
            '追加毎月雑費２
            Call subSetZappi(objSeet, intCount, strTemp & Nz(txtBillAddEzappiName2), fncZeroByteToNull(txtBillAddEzappi2) * txtBillPremthSum, "明細", "明細金額")
        End If
        
        '保証委託料
        'DELETE 2008/11/13 SHIBAZAKI
        'strTemp = IIf(txtBillSecukgLabel = P_LABEL_保証委託料, "初回" & txtBillSecukgLabel, txtBillSecukgLabel)
        'INSERT 2008/11/13 SHIBAZAKI
        strTemp = IIf(txtContainerHosyiFrom = P_HOSYIFROM_保証委託, "初回" & txtBillSecukgLabel, txtBillSecukgLabel)
        Call subSetZappi(objSeet, intCount, strTemp, fncZeroByteToNull(txtBillSecukgKei), "明細", "明細金額")
        '入金済み
        Call subSetZappi(objSeet, intCount, "入金済み", fncZeroByteToNull(txtBillSumikgKei), "明細", "明細金額")
    End With
    
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
Private Sub subSetZappi(objSeet As Object, _
                        ByRef intCount As Integer, _
                        strTekiyou As String, _
                        varCharge As Variant, _
                        strTekiyouRange As String, _
                        strChargeRange As String)

    If Nz(varCharge, 0) <> 0 Then
        intCount = intCount + 1
        objSeet.Range(strTekiyouRange & intCount) = Nz(strTekiyou)
        objSeet.Range(strChargeRange & intCount) = varCharge
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetUkeBook
'        機能             :受付管理表値連動
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetUkeBook(objSeet As Object)

    Dim strTemp         As String
    
    '2006/06/01
    Dim intCount        As Integer
    
    With objSeet
        .Range("CNS") = cmbUserCode     '顧客コード
        .Range("CNR") = txtUserKkbn     '顧客区分
        .Range("CNB") = txtRcptNo       '受付番号
        
        .Range("CNT") = txtUserKana     '顧客かな
        .Range("CNU") = txtUserName     '顧客名
        '顧客敬称 法人なら｢御中」　個人なら「様」
        If cmbUserKkbn = P_USER_KKBN_法人 Then
            .Range("CNV") = "御中"
        Else
            .Range("CNV") = "様"
        End If
        
        If Trim(Nz(txtUserRname)) = "" Then
            '顧客情報を適用
            .Range("CNAB") = txtUserTel     '顧客TEL
            .Range("CNAD") = txtUserKeitai  '顧客携帯
            .Range("CNAC") = txtUserFax     '顧客FAX
            
            .Range("CNY") = txtUserYubino   '顧客郵便番号
            .Range("CNZ").VALUE = txtUserAdr1   '顧客住所１
            .Range("CNAA").VALUE = txtUserAdr2 & "　" & Nz(txtUserAdr3, "")    '顧客住所２
            
'            .Range("CNW") = txtUserTanm     '担当者名                              'UPDATE 2011/02/28 M.RYU---<s>
            If Nz(txtUserTaka) <> "" Then
                .Range("CNW") = txtUserTanm & "（" & txtUserTaka & "）" '担当者名
            Else
                .Range("CNW") = txtUserTanm
            End If                                                                  'UPDATE 2011/02/28 M.RYU---<e>
            
            .Range("CNX") = "様"            '担当者敬称
        Else
            '書類送付先を適用
            .Range("CNAB") = txtUserRtel     '顧客TEL
            .Range("CNAD") = txtUserRkeitai  '顧客携帯
            .Range("CNAC") = txtUserRfax     '顧客FAX
            
            .Range("CNY") = txtUserRpost    '顧客郵便番号
            .Range("CNZ").VALUE = txtUserRadr1   '顧客住所１
            .Range("CNAA").VALUE = txtUserRadr2 & "　" & Nz(txtUserRadr3, "")    '顧客住所２
            
'            .Range("CNW") = txtUserRname     '担当者名                             'UPDATE 2011/02/28 M.RYU---<s>
            If Nz(txtUserRkana) <> "" Then
                .Range("CNW") = txtUserRname & "（" & txtUserRkana & "）"   '担当者名
            Else
                .Range("CNW") = txtUserRname
            End If                                                                  'UPDATE 2011/02/28 M.RYU---<e>
            
            .Range("CNX") = "様"            '担当者敬称
        End If
        
        .Range("G12") = txtUserMail         'メールアドレス 2006/01/12 Add tajima
        
        .Range("CNF") = "一般"              '期間
    
        .Range("CNI") = cmbContainerUsageCode.Column(1)
        
        .Range("CNG") = txtContainerYcode   '物件番号
        .Range("CNH") = txtContainerCno     'ボックス番号
    
        .Range("CNJ") = txtContainerYname   '物件名
        .Range("CNK") = txtContainerSize    'サイズ
        
        '▼2020/07/16 CHG .Range("CNK2") = txtContainerStep   '段
        .Range("_CNK2") = txtContainerStep   '段
        '▲2020/07/16 CHG
        
        '2009/06/30 MOD <S> hirano 受付管理表　修正
        '.Range("CNO") = txtBillMonthCharge  '月額使用料
        .Range("CNO") = txtBillMonthCharge + Nz(txt月額分割引, 0) '月額使用料
        '2009/06/30 MOD <E> hirano
        .Range("CNAV") = "0"                '使用料消費税
        .Range("CNAH") = txtBillFirstAmount '初回請求金額
        
        .Range("CNAF") = txtBillKisanDate   '請求起算日
        
        If Nz(txtBillSeikyuKikan) = "" Then
            strTemp = Format(txtBillKisanDate, "yyyy/mm")
        Else
            strTemp = txtBillSeikyuKikan
        End If
        .Range("CNAF2") = strTemp & "まで"   '請求期間
    
        .Range("CNP") = Nz(txtBillSecukg, 0) + Nz(txt保証金割引, 0)  '初回保証委託料 2011/06/18 保証金割引追加
        .Range("CNAI") = "0"                '初回請求分消費税
    
        .Range("CNL") = txtKeyKagiicd       'カギ種別
        .Range("CNM") = txtKeyKagiLendnum   '鍵本数
        .Range("CNBK") = txtKeyKagiNo       '鍵番号
        
        .Range("CNN") = txtBillMituCd       '見積もり請求区分
        
        '' 2012/07/04 M.HONDA START
        If txtReceiptYukbnCode <> P_YUKBN_受付 Then
            .Range("CNAJ") = ""    '予定収納物
        Else
            .Range("CNAJ") = txtAppFormButu     '予定収納物
        End If
        '' 2012/07/04 M.HONDA END
        
        .Range("CNAS") = txtSecurityHosyB   'リプラス承認番号

'        .Range("CNA") = txtReceiptUkdate    '受付日                            'DELETE 2013/06/21 K.ISHIZAKA
        .Range("CNA") = txtReceiptUkdatetime '受付日                            'INSERT 2013/06/21 K.ISHIZAKA
        .Range("CNE") = txtReceiptUktanto   '受付担当

        .Range("CNAW") = txtSendKeySendDate '鍵発送日
        .Range("CNAT") = txtSecurityHosyD   'リプラス戻り日

'        .Range("CNBB") = txtSendArrivalDate     '配送状況日付（配送完了日）  --20100622---del--ryu--
'        .Range("CNBC") = txtSendDeriveryStatus  '配送状況                    --20100622---del--ryu--

        .Range("CNBD") = txtSendName                '送付先名称
        .Range("CNBE") = txtSendHonorific           '送付先敬称
        .Range("CNBF") = txtSendYubino              '送付先郵便番号
        .Range("CNBG") = txtSendAdr1 & txtSendAdr2 & "　" & txtSendAdr3  '送付先住所
        .Range("CNBH") = txtSendTel                 '送付先電話番号
        .Range("CNAG") = txtBillSeikyuCd            '初回支払い区分
        .Range("CNAY") = txtSendArrivalExpect       '到着予定日
'        .Range("CNAZ") = txtSendSpecifyTime         '配達指定時間          --20100622---del--ryu--
        
        .Range("CNBA") = txtSendSlipNo              '送り状NO
        
        '2006/04/27 S.SHIBAZAKI 受付管理表の備考にコメントを出力
        '2006/06/01 SHIBAZAKI セキュリティカード対応
'        .Range("CNBJ") = Nz(txtReceiptUknote) & vbCrLf & Nz(txtReceiptComment1) & vbCrLf & Nz(txtReceiptComment2)           '備考
        .Range("CNBJ") = Nz(txtReceiptUknote)            '備考
        
        '2006/06/01 SHIBAZAKI セキュリティカード対応↓
        .Range("コメント1") = Nz(txtReceiptComment1)     'コメント１
        .Range("コメント2") = Nz(txtReceiptComment2)     'コメント２
        
        .Range("カード番号") = ""
        If lstSCardNo.ListCount <> 0 Then
            For intCount = 0 To lstSCardNo.ListCount - 1
                .Range("カード番号") = .Range("カード番号") & lstSCardNo.ItemData(intCount) & ","
            Next
            .Range("カード番号") = Left(.Range("カード番号"), Len(.Range("カード番号")) - 1)
        End If
        '2006/06/01 SHIBAZAKI セキュリティカード対応↑
        
        .Range("CXI") = txtReceiptComment1          '受付コメント１
        .Range("CXJ") = txtReceiptComment2          '受付コメント２
        
        '' ▼20101227 M.HONDA
        .Range("コメント3") = txtReceiptYoyakuComment          '予約コメント
        '' ▲20101227 M.HONDA
        
        If Nz(txtReceiptCyDate) <> "" Then
            .Range("解約予定").VALUE = P_HELPER_CANCEL_STRING
        End If
        
        '2006/04/28 S.SHIBAZAKI 受付管理表に書類送付方法を出力する
        .Range("CND") = Nz(txtAppFormSendCd)          '書類送付方法
        
        '2020/12/01 [EGL] UPD Start
        '2014/10/09 M.HONDA INS
        '.Range("コメント4") = txtContainerReason           '差額理由
        If cmb_クーポンCD.VALUE <> "" Then
            .Range("コメント4") = txtContainerReason & " " & cmb_クーポンCD.VALUE & "/" & txt_紹介会社名.VALUE
        Else
            .Range("コメント4") = txtContainerReason        '差額理由
        End If
        '2020/12/01 [EGL] UPD End
        
        '2015/02/13 M.HONDA INS ' TPOINT
        '.Range("TPOINT") = txtTpoint           'TPOINT
        
        .Range("HOMEI") = txtContainerHosyoCampany           '保証会社名称
        'INSERT 2020/04/01 N.IMAI Start
        If strBumonCode = "8" And cmbContainerHosyicd = "02" Then
            .Range("HOMEI") = txtContainerHosyicd            '保証会社名称
        End If
        'INSERT 2020/04/01 N.IMAI End
        
        .Range("AcptNo") = txtAcptNo           '契約No
        

    End With

End Sub

'2005/05/22 SHIBAZAKI 請求書２対応
'==============================================================================*
'
'        MODULE_NAME      :subSetSei2Book
'        機能             :請求書２値連動
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetSei2Book(objSeet As Object)

    Dim strTemp         As String
    
    With objSeet
        
        '---20100326---M.RYU---add------タイトル設定----
        .Range("TITLE") = txtBillMituCd                         'タイトル
            
        ''2012/08/30 M.HONDA START
        '’If Trim(Nz(txtUserRname)) = "" Then
        If Trim(Nz(txtUserRadr1)) = "" Then
        ''2012/08/30 M.HONDA END
            '顧客情報を適用
            .Range("顧客郵便番号") = txtUserYubino              '顧客郵便番号
            .Range("顧客住所１").VALUE = txtUserAdr1            '顧客住所１
            .Range("顧客住所２").VALUE = Nz(txtUserAdr2, "")    '顧客住所２
            .Range("顧客住所３").VALUE = Nz(txtUserAdr3, "")    '顧客住所３
            
            .Range("代表者名") = txtUserTanm     '担当者名
        Else
            '書類送付先を適用
            .Range("顧客郵便番号") = txtUserRpost               '顧客郵便番号
            .Range("顧客住所１").VALUE = txtUserRadr1           '顧客住所１
            .Range("顧客住所２").VALUE = Nz(txtUserRadr2, "")   '顧客住所２
            .Range("顧客住所３").VALUE = Nz(txtUserRadr3, "")   '顧客住所３
            
            .Range("代表者名") = txtUserRname     '担当者名
        End If
        If Trim(Nz(.Range("代表者名"), "")) <> "" Then
            .Range("代表者名") = .Range("代表者名") & "　様"            '担当者敬称
        End If
        
        .Range("顧客郵便番号") = "〒" & .Range("顧客郵便番号")
        
'        .range("お客様コード") = strBumonCode & cmbUserCode     '顧客コード        'DELETE 20110129 M.RYU
        .Range("お客様コード") = strBumonCode & "-" & cmbUserCode     '顧客コード   'INSERT 20110129 M.RYU
        .Range("顧客名") = txtUserName     '顧客名
        
        If Trim(Nz(.Range("代表者名"), "")) = "" Then
        
        '顧客敬称 法人なら｢御中」　個人なら「様」
        If cmbUserKkbn = P_USER_KKBN_法人 Then
            .Range("顧客名") = .Range("顧客名") & "　御中"
        Else
            .Range("顧客名") = .Range("顧客名") & "　様"
        End If
        
        End If
        
        'INSERT 2019/08/15 Y.WADA Start
        Dim strDate1        As String
        Dim dblZeiRitu1     As Double
        Dim strRoundKbn1    As String
        Dim lngZeikomi1     As Long
        Dim lngPrice1       As Long
        Dim lngTax1         As Long
        
        '税１
        strDate1 = Replace(Nz(txtBillKisanDate), "/", "")
        dblZeiRitu1 = MSZZ004_M20(strDate1, strRoundKbn1)
        lngZeikomi1 = Nz(txtBillRentkg, 0)
        Call MSZZ004_M10(lngZeikomi1, strDate1, "2", lngPrice1, lngTax1)
        
        .Range("税率1").VALUE = dblZeiRitu1 * 100
        '.Range("対象税込額1").VALUE = lngZeikomi1                              'INSERT 2023/08/24 N.IMAI
        .Range("消費税1").VALUE = lngTax1
    
        .Range("税率2").VALUE = ""
        '.Range("対象税込額2").VALUE = ""                                       'INSERT 2023/08/24 N.IMAI
        .Range("消費税2").VALUE = ""
        'INSERT 2019/08/15 Y.WADA End
        
        '----20100315----M.RYU----add-------<s>--'明細部に月額使用料を書き込む
        .Range("摘要1") = "    " & txtContainerYname & "  ( " & txtContainerCno & " ) "
        '.Range("金額1") = txtBillRentkg                                        'DELETE 2023/08/24 N.IMAI
        .Range("備考1") = ""
        '----20100315----M.RYU----add-------<e>
        'INSERT 2023/08/24 N.IMAI
        .Range("金額1") = lngPrice1                                             'INSERT 2023/08/24 N.IMAI
        
        'DELETE 2022/08/22 N.IMAI Start
        '↓↓---2018/09/25 add EGL---分社化対応、部門H以外は補足出力
'        If strBumonCode = "H" Then
'            .Range("補足文言") = ""
'        Else
'            .Range("補足文言") = "" '2020/05/01 KTS->KAS対応漏れ..."※口座名義[株式会社 加瀬倉庫]は収納代行会社です。"
'        End If
        '↑↑---2018/09/25 add EGL---
        'DELETE 2022/08/22 N.IMAI End
        
        'コントロールマスタから必要情報を取得して、シートに値を設定
        Call subContToSei2(objSeet)
        
        '加瀬DBから必要情報を取得して、シートに値を設定
        Call subKagtosToSei2(objSeet)
'DELETE 2023/11/09 N.IMAI Start
'        'INSERT 2023/10/26 N.IMAI Start
'        If objSeet.Range("金融機関") = "" Then
'            '振込先が無い場合は仮想口座の割当てを行い再度セット
'            Call fncInsertKouzMast
'            Call subKagtosToSei2(objSeet)
'        End If
'        'INSERT 2023/10/26 N.IMAI End
'DELETE 2023/11/09 N.IMAI End

    End With

End Sub

'2005/05/22 SHIBAZAKI 請求書２対応
'==============================================================================*
'
'        MODULE_NAME      :subContToSei2
'        機能             :コントロールマスタから請求書２に値連動
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subContToSei2(objSeet As Object)

    On Error GoTo ErrorHandler
    
    Dim dbSQLServer     As Database
    Dim rsData          As Recordset
    Dim blnError        As Boolean
    
    blnError = False
        
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    Set rsData = dbSQLServer.OpenRecordset(fncMakeContMastSql(), dbOpenSnapshot, dbSQLPassThrough)
    
    If rsData.EOF Then
        'データが存在しない
        'データを検索できないのは何らかの問題が発生しているはずなので、不正終了する
        Call MSZZ024_M10("OpenRecordset", "データを参照できませんでした。")
    End If

    'データ設定
    Call subContRsToSei2(rsData, objSeet)
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "subContToSei2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'2005/05/22 SHIBAZAKI 請求書２対応
'==============================================================================*
'
'        MODULE_NAME      :subContRsToSei2
'        機能             :コントロールマスタRecordSetから請求書２に値連動
'        IN               :①コントロールマスタRecordSet
'                         :②値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subContRsToSei2(rsData As Recordset, objSeet As Object)

    With objSeet
        .Range("会社名") = rsData.Fields("CONT_KAISYA")
        .Range("会社郵便番号") = "〒" & rsData.Fields("CONT_YUBINO")
        .Range("会社住所") = rsData.Fields("CONT_ADDR_1") & rsData.Fields("CONT_ADDR_2")
        .Range("会社TEL") = rsData.Fields("CONT_TEL_NO")
        .Range("会社FAX") = rsData.Fields("CONT_FAX_NO")
        .Range("登録番号") = rsData.Fields("CONT_SEIKYUSYO_TOUROKU_NO")     'INSERT 2019/08/15 Y.WADA
    
    End With
    
End Sub

'2005/05/22 SHIBAZAKI 請求書２対応
'==============================================================================*
'
'        MODULE_NAME      :subKagtosToSei2
'        機能             :加瀬DBから請求書２に値連動
'        IN               :値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKagtosToSei2(objSeet As Object)

    On Error GoTo ErrorHandler
    
    Dim dbSQLServer     As Database
    Dim rsData          As Recordset
    Dim blnError        As Boolean
    Dim strKaseSqlSvr   As String
    Dim strKaseConnect  As String
    
    blnError = False
        
    'SQLサーバー接続
    'SQL-Server名
        strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    If strKaseSqlSvr = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("DLookup(""SETUT_SETUN"", ""SETU_TABL"", ""SETUT_SETUB='ODBC_DATA_SOURCE_NAME')", "テーブル[SETU_TABL]の設定不正。")
    End If
    
    '接続文字列取得
    strKaseConnect = MSZZ007_M00()
    If strKaseConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "テーブル[SETU_TABL]の設定不正。")
    End If
    Set dbSQLServer = Workspaces(0).OpenDatabase(strKaseSqlSvr, dbDriverNoPrompt, False, strKaseConnect)

    Set rsData = dbSQLServer.OpenRecordset(fncMakeKagtosSql(), dbOpenSnapshot, dbSQLPassThrough)
    
    If rsData.EOF Then
        'データが存在しない
        'データを検索できないのは何らかの問題が発生しているはずなので、不正終了する
        Call MSZZ024_M10("OpenRecordset", "データを参照できませんでした。")
    End If

    'データ設定
    Call subKagtosRsToSei2(rsData, objSeet)
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "subKagtosToSei2" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub

'2005/05/22 SHIBAZAKI 請求書２対応
'==============================================================================*
'
'        MODULE_NAME      :fncMakeKagtosSql
'        機能             :請求書２用KAGTOS検索SQL作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeKagtosSql() As String

    Dim strSQL      As String
    
    strSQL = ""
'DELETE 2018/02/08 N.IMAI Start
'    strSQL = strSQL & " SELECT BANKT_KINYN AS 金融機関名, "
'    strSQL = strSQL & "        NYKOM_KINYC AS 金融機関コード, "
'    strSQL = strSQL & "        BANKT_SHITN AS 支店名, "
'    strSQL = strSQL & "        NYKOM_SHITC AS 支店コード, "
'    strSQL = strSQL & "        CODET_NAMEN AS 預金種別, "
'    strSQL = strSQL & "        NYKOM_KOUZB AS 口座番号, "
'DELETE 2018/02/08 N.IMAI End
    
    'INSERT 2018/02/08 N.IMAI Start
    strSQL = strSQL & " SELECT ISNULL(BANK_TAB2.BANKT_KINYN, BANK_TABL.BANKT_KINYN) AS 金融機関名, "
    strSQL = strSQL & "        ISNULL(KOUZ_MAST.KOUZM_KINYC, NYKO_MAST.NYKOM_KINYC) AS 金融機関コード, "
    strSQL = strSQL & "        ISNULL(BANK_TAB2.BANKT_SHITN, BANK_TABL.BANKT_SHITN) AS 支店名, "
    strSQL = strSQL & "        ISNULL(KOUZ_MAST.KOUZM_SHITC, NYKO_MAST.NYKOM_SHITC) AS 支店コード, "
    strSQL = strSQL & "        ISNULL(CODE_TAB2.CODET_NAMEN, CODE_TABL.CODET_NAMEN) AS 預金種別, "
    strSQL = strSQL & "        ISNULL(KOUZ_MAST.KOUZM_KOUZB, NYKO_MAST.NYKOM_KOUZB) AS 口座番号, "
    strSQL = strSQL & "        ISNULL(KOUZM_BUMOC, '') AS 仮想口座部門, "
    'INSERT 2018/02/08 N.IMAI End
    
    'strSQL = strSQL & "        NYKOM_KOUZN AS 口座名義人, "                                                    'DELETE 2022/08/22 N.IMAI
    strSQL = strSQL & "        IIF(KOUZ_MAST.KOUZM_KINYC IS NULL, NYKOM_KOUZN, CAMPM_CAMPN) AS 口座名義人, "    'INSERT 2022/08/22 N.IMAI
    strSQL = strSQL & "        CAMPM_CASEQ AS 第二会社コード, "
    strSQL = strSQL & "        ( SELECT BUMOM_BUSEQ "
    strSQL = strSQL & "            FROM BUMO_MAST "
    strSQL = strSQL & "           WHERE BUMOM_BUMOC = '" & strBumonCode & "' ) "
    strSQL = strSQL & "                    AS 第二部門コード "
    
    strSQL = strSQL & "        ,CONFT_SYKBI "                                   'INSERT 2023/08/24 N.IMAI
    
'DELETE 2022/11/30 N.IMAI Start
'    strSql = strSql & "   FROM CAMP_MAST "
'    strSql = strSql & "        INNER JOIN NYKO_MAST "
'    strSql = strSql & "           ON NYKOM_BUMOC = '" & strBumonCode & "' "
'    strSql = strSql & "          AND NYKOM_CAMPC = CAMPM_CAMPC "
'    'strSql = strSql & "          AND NYKOM_KINYC = CAMPM_KINYC "               'DELETE 20060530 N.MIURA
'    strSql = strSql & "          AND NYKOM_KINYC = '" & strKINYCode & "' "      'INSERT 20060530 N.MIURA
'
'    strSql = strSql & "        INNER JOIN BANK_TABL "
'    strSql = strSql & "           ON BANKT_KINYC = NYKOM_KINYC "
'    strSql = strSql & "          AND BANKT_SHITC = NYKOM_SHITC "
'
'    strSql = strSql & "        INNER JOIN CODE_TABL "
'    strSql = strSql & "           ON CODET_SIKBC = '" & P_CODE_MAST_預金種別 & "' "
'    strSql = strSql & "          AND CODET_CODEC = NYKOM_YOKII "
'
'    'INSERT 2018/02/08 N.IMAI Start
'                                                            '優先口座を取得するため、KOUZ_MAST Viewを使用   'INSERT 2019/05/01 Y.WADA
'    strSql = strSql & "        LEFT OUTER JOIN KOUZ_MAST "
'    strSql = strSql & "           ON KOUZM_BUMOC = '" & strBumonCode & "' "
''2018/08/06 EGL del    strSQL = strSQL & "          AND KOUZM_KOKYC = '900001' "   'TODO
'    'TODO 2018/02/09 strSQL = strSQL & "          AND KOUZM_KOKYC = '" & Format$(cmbUserCode, "000000") & "' "
'    strSql = strSql & "          AND KOUZM_KOKYC = '" & Format$(cmbUserCode, "000000") & "' " '2018/08/06 EGL add
'
'    strSql = strSql & "        LEFT OUTER JOIN BANK_TABL AS BANK_TAB2 "
'    strSql = strSql & "           ON BANK_TAB2.BANKT_KINYC = KOUZM_KINYC "
'    strSql = strSql & "          AND BANK_TAB2.BANKT_SHITC = KOUZM_SHITC "
'
'    strSql = strSql & "        LEFT OUTER JOIN CODE_TABL AS CODE_TAB2 "
'    strSql = strSql & "           ON CODE_TAB2.CODET_SIKBC = '" & P_CODE_MAST_預金種別 & "' "
'    strSql = strSql & "          AND CODE_TAB2.CODET_CODEC = KOUZM_YOKII "
'    'INSERT 2018/02/08 N.IMAI End
'DELETE 2022/11/30 N.IMAI End
    
    'INSERT 2022/11/30 N.IMAI Start
    strSQL = strSQL & "   FROM CAMP_MAST "

    strSQL = strSQL & "        LEFT OUTER JOIN KOUZ_MAST "
    strSQL = strSQL & "           ON KOUZM_BUMOC = '" & strBumonCode & "' "
    strSQL = strSQL & "          AND KOUZM_KOKYC = '" & Format$(cmbUserCode, "000000") & "' "
    strSQL = strSQL & "          AND KOUZM_KINYC != '0005' "                    'INSERT 2023/10/26 N.IMAI

    strSQL = strSQL & "        LEFT OUTER JOIN NYKO_MAST "
    strSQL = strSQL & "           ON NYKOM_BUMOC = '" & strBumonCode & "' "
    strSQL = strSQL & "          AND NYKOM_CAMPC = CAMPM_CAMPC "
    strSQL = strSQL & "          AND NYKOM_KINYC = '" & strKINYCode & "' "
    strSQL = strSQL & "          AND NYKOM_RIYOI = '0' "                        'INSERT 2023/10/26 N.IMAI
    
    strSQL = strSQL & "        LEFT OUTER JOIN BANK_TABL "
    strSQL = strSQL & "           ON BANKT_KINYC = NYKOM_KINYC "
    strSQL = strSQL & "          AND BANKT_SHITC = NYKOM_SHITC "
    
    strSQL = strSQL & "        LEFT OUTER JOIN CODE_TABL "
    strSQL = strSQL & "           ON CODET_SIKBC = '" & P_CODE_MAST_預金種別 & "' "
    strSQL = strSQL & "          AND CODET_CODEC = NYKOM_YOKII "
    
    strSQL = strSQL & "        LEFT OUTER JOIN BANK_TABL AS BANK_TAB2 "
    strSQL = strSQL & "           ON BANK_TAB2.BANKT_KINYC = KOUZM_KINYC "
    strSQL = strSQL & "          AND BANK_TAB2.BANKT_SHITC = KOUZM_SHITC "
    
    strSQL = strSQL & "        LEFT OUTER JOIN CODE_TABL AS CODE_TAB2 "
    strSQL = strSQL & "           ON CODE_TAB2.CODET_SIKBC = '" & P_CODE_MAST_預金種別 & "' "
    strSQL = strSQL & "          AND CODE_TAB2.CODET_CODEC = KOUZM_YOKII "
    'INSERT 2022/11/30 N.IMAI End
    strSQL = strSQL & "        CROSS JOIN CONF_TABL "                                       'INSERT 2023/08/24 N.IMAI
    
    'strSQL = strSQL & "  WHERE CAMPM_CAMPC = '" & strCampCode & "' "                       'DELETE 2022/08/22 N.IMAI
    strSQL = strSQL & "  WHERE CAMPM_CAMPC = ISNULL(KOUZM_CAMPC,'" & strCampCode & "') "    'INSERT 2022/08/22 N.IMAI

    fncMakeKagtosSql = strSQL
    
End Function

'2005/05/22 SHIBAZAKI 請求書２対応
'==============================================================================*
'
'        MODULE_NAME      :subKagtosRsToSei2
'        機能             :加瀬DBのRecordSetから請求書２に値連動
'        IN               :①加瀬DBのRecordSet
'                         :②値設定するシート
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKagtosRsToSei2(rsData As Recordset, objSeet As Object)

    Dim strTemp     As String
    
    With objSeet
        .Range("金融機関") = ""
        .Range("口座番号") = ""
        .Range("口座名義") = ""
        .Range("振込人名") = ""
        If Nz(rsData.Fields("金融機関コード")) <> "" Then                       'INSERT 2023/10/26 N.IMAI
            '銀行・支店の文字を追加                                                     ' INSERT 2011/03/30 M.RYU------<S>
            If DLookup("BANKT_KINYI", "BANK_TABL", _
                "BANKT_KINYC = '" & rsData.Fields("金融機関コード") & "' " & _
                " AND BANKT_SHITC = '" & rsData.Fields("支店コード") & "'") = "2" Then
                
                .Range("金融機関") = rsData.Fields("金融機関名") & _
                                 "銀行(" & rsData.Fields("金融機関コード") & ")  " & _
                                 rsData.Fields("支店名") & _
                                 "支店(" & rsData.Fields("支店コード") & ")"
            Else                                                                        ' INSERT 2011/03/30 M.RYU------<E>
                .Range("金融機関") = rsData.Fields("金融機関名") & _
                                     "(" & rsData.Fields("金融機関コード") & ")  " & _
                                     rsData.Fields("支店名") & _
                                     "(" & rsData.Fields("支店コード") & ")"
            End If
        
            .Range("口座番号") = rsData.Fields("預金種別") & "  " & _
                                 "No." & rsData.Fields("口座番号")
            .Range("口座名義") = rsData.Fields("口座名義人")
            
            If rsData.Fields("仮想口座部門") = "" Then                                              'INSERT 2018/02/08 N.IMAI
                strTemp = Replace(P_STRING_振込人名, "@CASEQ@", rsData.Fields("第二会社コード"))
                strTemp = Replace(strTemp, "@BUSEQ@", rsData.Fields("第二部門コード"))
                strTemp = Replace(strTemp, "@UCODE@", cmbUserCode)
                .Range("振込人名") = Replace(strTemp, "@UKANA@", txtUserKana)
            Else                                                                                    'INSERT 2018/02/08 N.IMAI
                .Range("振込人名") = txtUserKana                                                    'INSERT 2018/02/08 N.IMAI
            End If                                                                                  'INSERT 2018/02/08 N.IMAI
        End If
    End With
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdSendPostConv_Click
'        機能             :cmdSendPostConv_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdSendPostConv_Click()

    Call MSZ0000_YUBIB(Me.NAME, Me.txtSendYubino.NAME, Me.txtSendAdr1.NAME, _
                                Nz(Me.txtSendYubino.VALUE), Nz(Me.txtSendAdr1.VALUE))

    Call txtSendAdr1_AfterUpdate
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserCopy_Click
'        機能             :cmdUserCopy_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserCopy_Click()

    Call subUserToSendAdd
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserKpostConv_Click
'        機能             :cmdUserKpostConv_Click イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserKpostConv_Click()

    Call MSZ0000_YUBIB(Me.NAME, Me.txtUserKpost.NAME, Me.txtUserKadr1.NAME, _
                                Nz(Me.txtUserKpost.VALUE), Nz(Me.txtUserKadr1.VALUE))

    Call txtUserKadr1_AfterUpdate
    

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserKrenPostConv_Click
'        機能             :cmdUserKrenPostConv_Click イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserKrenPostConv_Click()

    Call MSZ0000_YUBIB(Me.NAME, Me.txtUserKrenPost.NAME, Me.txtUserKrenAdr1.NAME, _
                                Nz(Me.txtUserKrenPost.VALUE), Nz(Me.txtUserKrenAdr1.VALUE))

    Call txtUserKrenAdr1_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserMove_Click
'        機能             :一時使用変更承諾書
'        CREAT            :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserMove_Click()

    '処理実行
    If Not Print231.PrintUserMoveList(1, "", "", Nz(Me.cmbUserCode.VALUE, ""), Nz(Me.cmbUserCode.VALUE, "")) Then
        MsgBox "指定した条件に該当するデータはありませんでした。", vbInformation, "確認"
    End If



End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserRcopy_Click
'        機能             :cmdUserRcopy_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserRcopy_Click()

    If Trim(Nz(txtUserRname)) = "" Then
        Call subUserToSendAdd
    Else
        Call subUserSendToSendAdd
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserRpostConv_Click
'        機能             :cmdUserRpostConv_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserRpostConv_Click()

    Call MSZ0000_YUBIB(Me.NAME, Me.txtUserRpost.NAME, Me.txtUserRadr1.NAME, _
                                Nz(Me.txtUserRpost.VALUE), Nz(Me.txtUserRadr1.VALUE))

    Call txtUserRadr1_AfterUpdate
    
End Sub

'2006/06/01 SHIBAZAKI ヤード地図をクリックするとWEB地図を開くようにする
'==============================================================================*
'
'        MODULE_NAME      :cmdYardMap_Click
'        機能             :cmdYardMap_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdYardMap_Click()
On Error GoTo ErrorHandler

    If WebYardCode(txtContainerYcode) = False Then
        MsgBox "ヤードの位置情報が未設定のため開けませんでした。", vbInformation, P_KAKUNIN
    End If
    
    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdYardMap_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :フォームクローズ
'        MODULE_ID        :Form_Close
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Close()
    Call MSHS010_M00(PROG_ID, "2")
End Sub

Private Sub Form_Timer()
Dim StrTmp As String
'    StrTmp = Me.lbl_msg.Caption
'    Me.lbl_msg.Caption = Mid$(StrTmp, 2) & Left$(StrTmp, 1)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : opgRentkg_AfterUpdate
'        機能             : 使用料・雑費を[引き継ぐ/引き継がない]
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub opgRentkg_AfterUpdate()
    On Error GoTo ErrorHandler
    
    If opgRentkg.VALUE = opbRentkgHIKI.OptionValue Then
        txtBillRentkg.VALUE = txtBillRentkgMoto.VALUE
        txtBillEzappi.VALUE = txtBillEzappiMoto.VALUE
        txtEzappiCode.VALUE = txtBillEzappiCodeMoto.VALUE
        txtBillEzappiName.VALUE = txtBillEzappiNameMoto.VALUE
    Else
        txtBillRentkg.VALUE = txtContainerPrice.VALUE
        txtBillEzappi.VALUE = txtContainerEzapii.VALUE
        txtEzappiCode.VALUE = txtContainerEzapiiCode.VALUE
        txtBillEzappiName.VALUE = txtContainerEzapiiName.VALUE
    End If
    Call cmdReCal_Click
    Call subDispSecukgMessage
    Call subDispEzappiMessage
Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("opgRentkg_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi2_AfterUpdate
'        機能             :txtBillAddEzappi2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi2_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillAddEzappi2, P_PRICE_LENGTH, P_CONTROL_金額)
    If Nz(txtBillAddEzappiDaily2) = "" Then
        txtBillAddEzappiDaily2 = fncDailyRate(txtBillAddEzappi2)
    End If
    '▼2008/12/04 Mod hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。
    'txtBillSecukg = txtBillRentkg + txtBillEzappiKei  '2007/02/07 add tajima 保証料は月額の金額全て
    '▼2008/12/20 Mod hirano 保証料は、基本料＋毎月雑費から月額分割割引合計（月数で割る）をマイナスする。
    'txtBillSecukg = txtBillRentkg + txtBillEzappiKei + txt月額分割引合計
    If Not chk保証金強制変更.VALUE Then
        '新旧プラン判定
        If FVS400_fncChkHoshoNewOld(strBumonCode, cmbContainerHosyoCompany) = False Then '2017/12/20 ADD
            'del 2018/10/02 EGL -->If chk割引有効可否.VALUE Then txtBillSecukg = txtBillMonthCharge + txt月額分割引
            Call subCompSecukgKase                                                  'INSERT 2013/04/10 K.ISHIZAKA
        End If
    End If  '2017/12/20 ADD
    '▲2008/12/20 Mod hirano
    '▲2008/12/04 Mod hirano
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi2_BeforeUpdate
'        機能             :txtBillAddEzappi2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi2_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillAddEzappi2, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi2_GotFocus
'        機能             :txtBillAddEzappi2_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi2_GotFocus()

    txtBillAddEzappi2.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi2_KeyPress
'        機能             :txtBillAddEzappi2_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi2_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillAddEzappi2, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi2_LostFocus
'        機能             :txtBillAddEzappi2_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi2_LostFocus()

    txtBillAddEzappi2.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily2_AfterUpdate
'        機能             :txtBillAddEzappiDaily2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily2_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillAddEzappiDaily2, P_PRICE_LENGTH, P_CONTROL_金額)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily2_BeforeUpdate
'        機能             :txtBillAddEzappiDaily2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily2_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillAddEzappiDaily2, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily2_GotFocus
'        機能             :txtBillAddEzappiDaily2_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily2_GotFocus()

    txtBillAddEzappiDaily2.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily2_KeyPress
'        機能             :txtBillAddEzappiDaily2_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily2_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillAddEzappiDaily2, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily2_LostFocus
'        機能             :txtBillAddEzappiDaily2_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily2_LostFocus()

    txtBillAddEzappiDaily2.Format = P_FORMAT_STANDARD

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily1_AfterUpdate
'        機能             :txtBillAddEzappiDaily1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily1_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillAddEzappiDaily1, P_PRICE_LENGTH, P_CONTROL_金額)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily1_BeforeUpdate
'        機能             :txtBillAddEzappiDaily1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily1_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillAddEzappiDaily1, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily1_GotFocus
'        機能             :txtBillAddEzappiDaily1_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily1_GotFocus()

    txtBillAddEzappiDaily1.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily1_KeyPress
'        機能             :txtBillAddEzappiDaily1_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily1_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillAddEzappiDaily1, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappiDaily1_LostFocus
'        機能             :txtBillAddEzappiDaily1_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappiDaily1_LostFocus()

    txtBillAddEzappiDaily1.Format = P_FORMAT_STANDARD

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappi_AfterUpdate
'        機能             :txtBillEzappi_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappi_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillEzappi, P_PRICE_LENGTH, P_CONTROL_金額)
    '▼2008/12/04 Add hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。
    'txtBillSecukg = txtBillRentkg + txtBillEzappiKei  '2007/02/07 add tajima 保証料は月額の金額全て
    '▼2008/12/20 Add hirano 保証料は、基本料＋毎月雑費から月額分割割引合計(月数で割る)をマイナスする。
    'txtBillSecukg = txtBillRentkg + txtBillEzappiKei + txt月額分割引合計
    If Not chk保証金強制変更.VALUE Then
        'del 2018/10/02 EGL -->If chk割引有効可否.VALUE Then txtBillSecukg = txtBillMonthCharge + txt月額分割引
        Call subCompSecukgKase                                                  'INSERT 2013/04/10 K.ISHIZAKA
    End If
    '▲2008/12/20 Add hirano
    '▲2008/12/04 Add hirano
    
    Call subDispEzappiMessage                                                   'INSERT 2013/04/10 K.ISHIZAKA
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappi_BeforeUpdate
'        機能             :txtBillEzappi_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappi_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillEzappi, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappi_GotFocus
'        機能             :txtBillEzappi_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappi_GotFocus()

    txtBillEzappi.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappi_KeyPress
'        機能             :txtBillEzappi_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappi_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillEzappi, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappi_LostFocus
'        機能             :txtBillEzappi_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappi_LostFocus()

    txtBillEzappi.Format = P_FORMAT_STANDARD

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappiDaily_AfterUpdate
'        機能             :txtBillEzappiDaily_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappiDaily_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillEzappiDaily, P_PRICE_LENGTH, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappiDaily_BeforeUpdate
'        機能             :txtBillEzappiDaily_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappiDaily_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillEzappiDaily, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappiDaily_GotFocus
'        機能             :txtBillEzappiDaily_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappiDaily_GotFocus()

    txtBillEzappiDaily.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappiDaily_KeyPress
'        機能             :txtBillEzappiDaily_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappiDaily_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillEzappiDaily, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillEzappiDaily_LostFocus
'        機能             :txtBillEzappiDaily_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillEzappiDaily_LostFocus()

    txtBillEzappiDaily.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFirstkg_AfterUpdate
'        機能             :txtBillFirstkg_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFirstkg_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillFirstkg, P_PRICE_LENGTH, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFirstkg_BeforeUpdate
'        機能             :txtBillFirstkg_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFirstkg_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillFirstkg, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFirstkg_GotFocus
'        機能             :txtBillFirstkg_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFirstkg_GotFocus()

    txtBillFirstkg.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFirstkg_KeyPress
'        機能             :txtBillFirstkg_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFirstkg_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillFirstkg, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFirstkg_LostFocus
'        機能             :txtBillFirstkg_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFirstkg_LostFocus()

    txtBillFirstkg.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSecukg_AfterUpdate
'        機能             :txtBillSecukg_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukg_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillSecukg, P_PRICE_LENGTH, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSecukg_BeforeUpdate
'        機能             :txtBillSecukg_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukg_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillSecukg, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSecukg_GotFocus
'        機能             :txtBillSecukg_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukg_GotFocus()

    txtBillSecukg.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSecukg_KeyPress
'        機能             :txtBillSecukg_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukg_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillSecukg, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSecukg_LostFocus
'        機能             :txtBillSecukg_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukg_LostFocus()

    txtBillSecukg.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      : txtBillSecukgKase_AfterUpdate
'        機能             : txtBillSecukgKase_AfterUpdate イベント
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukgKase_AfterUpdate()
    Call fncAfterUpdateCommon(txtBillSecukgKase, P_PRICE_LENGTH, P_CONTROL_金額)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : txtBillSecukgKase_BeforeUpdate
'        機能             : txtBillSecukgKase_BeforeUpdate イベント
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukgKase_BeforeUpdate(Cancel As Integer)
    Call fncBeforeUpdateCommon(txtBillSecukgKase, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : txtBillSecukgKase_GotFocus
'        機能             : txtBillSecukgKase_GotFocus イベント
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukgKase_GotFocus()
    txtBillSecukgKase.Format = P_FORMAT_FIXED
End Sub

'==============================================================================*
'
'        MODULE_NAME      : txtBillSecukgKase_KeyPress
'        機能             : txtBillSecukgKase_KeyPress イベント
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukgKase_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(txtBillSecukgKase, KeyAscii, P_PRICE_LENGTH)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : txtBillSecukgKase_LostFocus
'        機能             : txtBillSecukgKase_LostFocus イベント
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSecukgKase_LostFocus()
    txtBillSecukgKase.Format = P_FORMAT_STANDARD
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSumikg_AfterUpdate
'        機能             :txtBillSumikg_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSumikg_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillSumikg, P_PRICE_LENGTH, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSumikg_BeforeUpdate
'        機能             :txtBillSumikg_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSumikg_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillSumikg, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSumikg_GotFocus
'        機能             :txtBillSumikg_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSumikg_GotFocus()

    txtBillSumikg.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSumikg_KeyPress
'        機能             :txtBillSumikg_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSumikg_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillSumikg, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSumikg_LostFocus
'        機能             :txtBillSumikg_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSumikg_LostFocus()

    txtBillSumikg.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbKeyKagiicd_AfterUpdate
'        機能             :cmbKeyKagiicd_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbKeyKagiicd_AfterUpdate()

    txtKeyKagiicd = fncAfterUpdateCommon(cmbKeyKagiicd, P_DIVISION_LENGTH, P_CONTROL_数値)
'↓ INSERT 2014/03/11 MIYAMOTO
    '自動鍵タイプの場合(1=自動鍵)
'↓UPDATE  2016/05/17 SUGIMURA
    'If Nz(cmbKeyKagiicd.Column(2)) = "1" Then
    
'↓ insert 2021/06/02 EGL
    txtKeyKagiNo.InputMask = ""
    If cmbKeyKagiicd.Column(0) = 16 Then 'insert 2016/06/03 EGL
        ラベル858.Caption = "物件鍵NO/部屋鍵NO"
        txtKaijoNo.Visible = True
        txtKaijoNo.VALUE = strKaijoNo
        txtKeyKagiNo = strKaijoNo_QR
        If pvONLET_ACCEI = 0 Then
            '権限無し
            txtKeyKagiNo.InputMask = "Password"
        Else
            txtKeyKagiNo.InputMask = ""
        End If
        Exit Sub
'↑ insert 2021/06/02 EGL

    ElseIf MSZZ070_KaiKey(txtYardTihoKbn, Nz(cmbKeyKagiicd.Column(2))) <> "0" Then
'↑UPDATE  2016/05/17 SUGIMURA
        ラベル858.Caption = "鍵番号/解除番号"
        txtKaijoNo.Visible = True
        txtKaijoNo.VALUE = strKaijoNo
    Else
        ラベル858.Caption = "鍵番号"
        txtKaijoNo.Visible = False
        txtKaijoNo.VALUE = ""
    End If
'↑ INSERT 2014/03/11 MIYAMOTO

'↓ insert 2021/06/02 EGL
    If pvONLET_ACCEI = 0 Then
        '権限無し
        txtKaijoNo.InputMask = "Password"
    Else
        txtKaijoNo.InputMask = ""
    End If
'↑ insert 2021/06/02 EGL
End Sub

'==============================================================================*
'
'        MODULE_NAME      :解除番号入力権限取得
'        MODULE_ID        :fncGetKaiKeyAdmin
'        CREATE_DATE      :2016/05/16
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetKaiKeyAdmin() As String

    Dim isException     As Boolean
    Dim strSQL          As String
    Dim ojDyn           As Recordset
    Dim vCode(0 To 1)   As Variant
    Dim userId          As String
    Dim dbKASE_DB       As Database
    
    isException = False
    On Error GoTo Exception
    
    ' 加瀬DB接続
    Set dbKASE_DB = Workspaces(0).OpenDatabase(Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'")), dbDriverNoPrompt, False, MSZZ007_M00())
    
    'ユーザ情報取得
    userId = MSZZ000.LsGetUserName()
    
    strSQL = ""
    strSQL = "SELECT ONLET_ACCEI "
    strSQL = strSQL & "FROM ONLE_TABL "
    strSQL = strSQL & "WHERE ONLET_PROGB = 'FVS400' "
    strSQL = strSQL & "AND ONLET_USERB = '" & userId & "'; "
    Set ojDyn = dbKASE_DB.OpenRecordset(strSQL, dbOpenDynaset)
    
    If Not ojDyn.EOF Then
        fncGetKaiKeyAdmin = ojDyn!ONLET_ACCEI
    Else
        fncGetKaiKeyAdmin = "0"
    End If
    
    GoTo Finally
Exception:
    isException = True
Finally:
    If Not ojDyn Is Nothing Then ojDyn.Close: Set ojDyn = Nothing
    If Not dbKASE_DB Is Nothing Then dbKASE_DB.Close: Set dbKASE_DB = Nothing
    If isException Then
        Call Err.Raise(Err.Number, "fncGetKaiKeyAdmin" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Function

'==============================================================================*
'
'        MODULE_NAME      :cmbKeyKagiicd_BeforeUpdate
'        機能             :cmbKeyKagiicd_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbKeyKagiicd_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbKeyKagiicd, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
    '-----20091205------M.RYU-------update------------<s>-'画面で南京錠を使えるように修正
''    '▼2008/03/23 add tajima
''    If Me.cmbKeyKagiicd.VALUE = 40 Or Me.cmbKeyKagiicd.VALUE = 50 Then
''        MsgBox "【南京錠】及び【番線 (カギなし)】を選択することは出来ません｡ ", vbInformation, P_ERROR
''        Cancel = True
''        SendKeys "{ESC}"
''    End If
''    '▲2008/03/23 add tajima
    '-----20120725------M.HONDA-------update------------<s>-'
''    If Me.cmbKeyKagiicd.VALUE = 50 Then
''        MsgBox "【番線 (カギなし)】を選択することは出来ません｡ ", vbInformation, P_ERROR
''        Cancel = True
''        SendKeys "{ESC}"
''    End If
    '-----20120725------M.HONDA-------update------------<e>
    '-----20091205------M.RYU-------update------------<e>

    '↓ INSERT 2014/03/11 MIYAMOTO
'↓UPDATE  2016/05/17 SUGIMURA
'   If Nz(cmbKeyKagiicd.Column(2)) = "1" Then
'　　　 MsgBox "自動鍵の鍵区分は指定出来ません。", vbInformation, P_ERROR
    If MSZZ070_KaiKey(txtYardTihoKbn, Nz(cmbKeyKagiicd.Column(2))) <> "0" Then
        MsgBox "解除番号付き鍵への鍵区分は指定出来ません。", vbInformation, P_ERROR
'↑UPDATE  2016/05/17 SUGIMURA
        Cancel = True
        SendKeys "{ESC}"
    End If
    '↑ INSERT 2014/03/11 MIYAMOTO
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbKeyKagiicd_KeyPress
'        機能             :cmbKeyKagiicd_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbKeyKagiicd_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbKeyKagiicd, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSendHonorific_AfterUpdate
'        機能             :cmbSendHonorific_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSendHonorific_AfterUpdate()

    txtSendHonorific = fncAfterUpdateCommon(cmbSendHonorific, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSendHonorific_BeforeUpdate
'        機能             :cmbSendHonorific_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSendHonorific_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbSendHonorific, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSendHonorific_KeyPress
'        機能             :cmbSendHonorific_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSendHonorific_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbSendHonorific, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptPaydate_AfterUpdate
'        機能             :txtReceiptPaydate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptPaydate_AfterUpdate()

    txtReceiptPaydate = fncAfterUpdateCommon(txtReceiptPaydate, P_DATE_LENGTH, P_CONTROL_日付)
    Call subDailyRate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptPaydate_BeforeUpdate
'        機能             :txtReceiptPaydate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptPaydate_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtReceiptPaydate, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptPaydate_GotFocus
'        機能             :txtReceiptPaydate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptPaydate_GotFocus()

    txtReceiptPaydate = Format(txtReceiptPaydate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptPaydate_KeyPress
'        機能             :txtReceiptPaydate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptPaydate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtReceiptPaydate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptPaydate_LostFocus
'        機能             :txtReceiptPaydate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptPaydate_LostFocus()

    txtReceiptPaydate = Format(txtReceiptPaydate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptUknote_AfterUpdate
'        機能             :txtReceiptUknote_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptUknote_AfterUpdate()

    Call fncAfterUpdateCommon(txtReceiptUknote, P_NOTE_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptUknote_BeforeUpdate
'        機能             :txtReceiptUknote_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptUknote_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtReceiptUknote, P_NOTE_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptComment1_AfterUpdate
'        機能             :txtReceiptComment1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptComment1_AfterUpdate()

    Call fncAfterUpdateCommon(txtReceiptComment1, P_NOTE_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptComment1_BeforeUpdate
'        機能             :txtReceiptComment1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptComment1_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtReceiptComment1, P_NOTE_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptComment2_AfterUpdate
'        機能             :txtReceiptComment2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptComment2_AfterUpdate()

    Call fncAfterUpdateCommon(txtReceiptComment2, P_NOTE_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptComment2_BeforeUpdate
'        機能             :txtReceiptComment2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptComment2_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtReceiptComment2, P_NOTE_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalExpect_AfterUpdate
'        機能             :txtSendArrivalExpect_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalExpect_AfterUpdate()

    txtSendArrivalExpect = fncAfterUpdateCommon(txtSendArrivalExpect, P_DATE_LENGTH, P_CONTROL_日付)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalExpect_BeforeUpdate
'        機能             :txtSendArrivalExpect_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalExpect_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSendArrivalExpect, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalExpect_GotFocus
'        機能             :txtSendArrivalExpect_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalExpect_GotFocus()

    txtSendArrivalExpect = Format(txtSendArrivalExpect, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalExpect_KeyPress
'        機能             :txtSendArrivalExpect_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalExpect_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtSendArrivalExpect, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalExpect_LostFocus
'        機能             :txtSendArrivalExpect_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalExpect_LostFocus()

    txtSendArrivalExpect = Format(txtSendArrivalExpect, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalDate_AfterUpdate
'        機能             :txtSendArrivalDate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalDate_AfterUpdate()

    txtSendArrivalDate = fncAfterUpdateCommon(txtSendArrivalDate, P_DATE_LENGTH, P_CONTROL_日付)

    '入金確認日にコピー
    txtReceiptPaydate = txtSendArrivalDate
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalDate_BeforeUpdate
'        機能             :txtSendArrivalDate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalDate_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSendArrivalDate, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalDate_GotFocus
'        機能             :txtSendArrivalDate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalDate_GotFocus()

    txtSendArrivalDate = Format(txtSendArrivalDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalDate_KeyPress
'        機能             :txtSendArrivalDate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalDate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtSendArrivalDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendArrivalDate_LostFocus
'        機能             :txtSendArrivalDate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendArrivalDate_LostFocus()

    txtSendArrivalDate = Format(txtSendArrivalDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendKeySendDate_AfterUpdate
'        機能             :txtSendKeySendDate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendKeySendDate_AfterUpdate()

    txtSendKeySendDate = fncAfterUpdateCommon(txtSendKeySendDate, P_DATE_LENGTH, P_CONTROL_日付)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendKeySendDate_BeforeUpdate
'        機能             :txtSendKeySendDate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendKeySendDate_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSendKeySendDate, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendKeySendDate_GotFocus
'        機能             :txtSendKeySendDate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendKeySendDate_GotFocus()

    txtSendKeySendDate = Format(txtSendKeySendDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendKeySendDate_KeyPress
'        機能             :txtSendKeySendDate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendKeySendDate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtSendKeySendDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendKeySendDate_LostFocus
'        機能             :txtSendKeySendDate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendKeySendDate_LostFocus()

    txtSendKeySendDate = Format(txtSendKeySendDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSendSpecifyTime_AfterUpdate
'        機能             :cmbSendSpecifyTime_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSendSpecifyTime_AfterUpdate()

    txtSendSpecifyTime = fncAfterUpdateCommon(cmbSendSpecifyTime, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSendSpecifyTime_BeforeUpdate
'        機能             :cmbSendSpecifyTime_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSendSpecifyTime_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbSendSpecifyTime, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbSendSpecifyTime_KeyPress
'        機能             :cmbSendSpecifyTime_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbSendSpecifyTime_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbSendSpecifyTime, KeyAscii, P_DIVISION_LENGTH)

End Sub

'2019.03.22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
'==============================================================================*
'
'        MODULE_NAME      :txtSendYamatoDate_AfterUpdate
'        機能             :txtSendYamatoDate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYamatoDate_AfterUpdate()

    txtSendYamatoDate = fncAfterUpdateCommon(txtSendYamatoDate, P_DATE_LENGTH, P_CONTROL_日付)

End Sub

'2019.03.22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
'==============================================================================*
'
'        MODULE_NAME      :txtSendYamatoDate_BeforeUpdate
'        機能             :txtSendYamatoDate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYamatoDate_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSendYamatoDate, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'2019.03.22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
'==============================================================================*
'
'        MODULE_NAME      :txtSendYamatoDate_GotFocus
'        機能             :txtSendYamatoDate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYamatoDate_GotFocus()

    txtSendYamatoDate = Format(txtSendYamatoDate, "yy/mm/dd")

End Sub

'2019.03.22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
'==============================================================================*
'
'        MODULE_NAME      :txtSendYamatoDate_KeyPress
'        機能             :txtSendYamatoDate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYamatoDate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtSendYamatoDate, KeyAscii, P_DATE_LENGTH)

End Sub

'2019.03.22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
'==============================================================================*
'
'        MODULE_NAME      :txtSendYamatoDate_LostFocus
'        機能             :txtSendYamatoDate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYamatoDate_LostFocus()

    txtSendYamatoDate = Format(txtSendYamatoDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbReceiptUktanto_AfterUpdate
'        機能             :cmbReceiptUktanto_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReceiptUktanto_AfterUpdate()

    txtReceiptUktanto = fncAfterUpdateCommon(cmbReceiptUktanto, P_PERSON_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbReceiptUktanto_BeforeUpdate
'        機能             :cmbReceiptUktanto_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReceiptUktanto_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbReceiptUktanto, P_PERSON_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbReceiptUktanto_KeyPress
'        機能             :cmbReceiptUktanto_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReceiptUktanto_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbReceiptUktanto, KeyAscii, P_PERSON_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbReceiptCreason_AfterUpdate
'        機能             :cmbReceiptCreason_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReceiptCreason_AfterUpdate()

    txtReceiptCreason = fncAfterUpdateCommon(cmbReceiptCreason, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbReceiptCreason_BeforeUpdate
'        機能             :cmbReceiptCreason_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReceiptCreason_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbReceiptCreason, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbReceiptCreason_KeyPress
'        機能             :cmbReceiptCreason_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbReceiptCreason_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbReceiptCreason, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtAppFormButu_AfterUpdate
'        機能             :txtAppFormButu_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtAppFormButu_AfterUpdate()

    Call fncAfterUpdateCommon(txtAppFormButu, P_STRAGE_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtAppFormButu_BeforeUpdate
'        機能             :txtAppFormButu_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtAppFormButu_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtAppFormButu, P_STRAGE_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtAppFormLimitDate_AfterUpdate
'        機能             :txtAppFormLimitDate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtAppFormLimitDate_AfterUpdate()

    txtAppFormLimitDate = fncAfterUpdateCommon(txtAppFormLimitDate, P_DATE_LENGTH, P_CONTROL_日付)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtAppFormLimitDate_BeforeUpdate
'        機能             :txtAppFormLimitDate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtAppFormLimitDate_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtAppFormLimitDate, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtAppFormLimitDate_GotFocus
'        機能             :txtAppFormLimitDate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtAppFormLimitDate_GotFocus()

    txtAppFormLimitDate = Format(txtAppFormLimitDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtAppFormLimitDate_KeyPress
'        機能             :txtAppFormLimitDate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtAppFormLimitDate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtAppFormLimitDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtAppFormLimitDate_LostFocus
'        機能             :txtAppFormLimitDate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtAppFormLimitDate_LostFocus()

    txtAppFormLimitDate = Format(txtAppFormLimitDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillStDate_AfterUpdate
'        機能             :txtBillStDate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillStDate_AfterUpdate()

    txtBillStDate = fncAfterUpdateCommon(txtBillStDate, P_DATE_LENGTH, P_CONTROL_日付)
    txt割引適用開始月 = Mid(txtBillStDate, 1, 7)
    Call subDailyRate
    Call chkContainerSevExpire_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillStDate_BeforeUpdate
'        機能             :txtBillStDate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillStDate_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillStDate, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillStDate_GotFocus
'        機能             :txtBillStDate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillStDate_GotFocus()

    txtBillStDate = Format(txtBillStDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillStDate_KeyPress
'        機能             :txtBillStDate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillStDate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtBillStDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillStDate_LostFocus
'        機能             :txtBillStDate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillStDate_LostFocus()

    txtBillStDate = Format(txtBillStDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSeikyuKikan_AfterUpdate
'        機能             :txtBillSeikyuKikan_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSeikyuKikan_AfterUpdate()

    txtBillSeikyuKikan = fncAfterUpdateCommon(txtBillSeikyuKikan, P_YEARMONTH_LENGTH, P_CONTROL_年月)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSeikyuKikan_BeforeUpdate
'        機能             :txtBillSeikyuKikan_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSeikyuKikan_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillSeikyuKikan, P_YEARMONTH_LENGTH, P_CONTROL_年月, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSeikyuKikan_GotFocus
'        機能             :txtBillSeikyuKikan_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSeikyuKikan_GotFocus()
    
    txtBillSeikyuKikan = Format(fncFirstDay(txtBillSeikyuKikan), "yy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSeikyuKikan_KeyPress
'        機能             :txtBillSeikyuKikan_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSeikyuKikan_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtBillSeikyuKikan, KeyAscii, P_YEARMONTH_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillSeikyuKikan_LostFocus
'        機能             :txtBillSeikyuKikan_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillSeikyuKikan_LostFocus()

    txtBillSeikyuKikan = Format(fncFirstDay(txtBillSeikyuKikan), "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :FVS400_fncChkHoshoNewOld
'        機能             :保証プラン(新・旧の判断) (2017/12/20 add)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function FVS400_fncChkHoshoNewOld(strBumonCode As String, strHoshoCode As String) As Boolean
 
    Dim dbSQLServer     As Database
    Dim rsData          As Recordset
    Dim blnError        As Boolean
    Dim strKaseSqlSvr   As String
    Dim strSQL          As String
    
    On Error GoTo ErrorHandler
    
    FVS400_fncChkHoshoNewOld = False
    blnError = False
        
    'SQL-Server名
'   strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))'DELETE 2019/03/07 Y.WADA
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'")) 'INSERT 2019/03/07 Y.WADA

    '保証会社
    strSQL = strSQL & " select count(*) as CNT"
    strSQL = strSQL + " from " & strKaseSqlSvr & ".dbo.SHIR_MAST"
    strSQL = strSQL & " where"
    strSQL = strSQL & "     SHIRM_SHIRI = '8'"
    strSQL = strSQL & " and"
    strSQL = strSQL & "     SHIRM_HOJYC = '1'"
    strSQL = strSQL & " and"
    strSQL = strSQL & "     SHIRM_BUMOC = '" & strBumonCode & "'"
    strSQL = strSQL & " and"
    strSQL = strSQL & "     SHIRM_SHIRC = '" & strHoshoCode & "'"
        
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    Set rsData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    If rsData.Fields("CNT") > 0 Then
        FVS400_fncChkHoshoNewOld = True
    End If
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "FVS400_fncChkHoshoNewOld" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncGetEZappi
'        機能             :毎月雑費金取得 (2017/12/20 add)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subGetEZappi(strRCPT_NO As String)

    Dim dbSQLServer     As Database
    'Dim rsData          As Recordset                                           'DELETE 2020/10/30
    Dim rsData          As Object                                               'INSERT 2020/10/30
    Dim blnError        As Boolean
    Dim strSQL          As String
    
    On Error GoTo ErrorHandler

    blnError = False
    
    'init
    txtBillEzappiName = ""
    txtBillEzappi = 0
    
    'SQL
    strSQL = strSQL & " select"
    strSQL = strSQL & "      NAME_NAME"
    strSQL = strSQL & "     ,PRIC_EZAPPI"
    strSQL = strSQL & " from"
    strSQL = strSQL & "     YOUK_TRAN"
    strSQL = strSQL & " inner join CNTA_MAST"
    strSQL = strSQL & "  on CNTA_CODE = YOUKT_YCODE"
    strSQL = strSQL & " and CNTA_NO   = YOUKT_NO"
    strSQL = strSQL & " inner join PRIC_TABL"
    strSQL = strSQL & "  on PRIC_YCODE = CNTA_CODE"
    strSQL = strSQL & " and PRIC_USAGE = CNTA_USAGE"
    strSQL = strSQL & " and PRIC_SIZE  = CNTA_SIZE"
    strSQL = strSQL & " and PRIC_STEP  = CNTA_STEP"
    strSQL = strSQL & " LEFT OUTER JOIN NAME_MAST"
    strSQL = strSQL & " on  NAME_ID = '096'"
    strSQL = strSQL & " and NAME_CODE = PRIC_TABL.PRIC_EZAPPI_CODE"
    strSQL = strSQL & " where"
    strSQL = strSQL & "     YOUKT_UKNO = '" & strRCPT_NO & "'"
        
    'SQLサーバー接続
'    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)    'DELETE 2020/10/30
'    Set rsData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)                   'DELETE 2020/10/30
    Set rsData = ADODB_Recordset(strSQL, ADODB_Connection(strBumonCode))                                'INSERT 2020/10/30
    If Not rsData.EOF Then
        '毎月雑費名
        txtBillEzappiName = Nz(rsData.Fields("NAME_NAME"), "")
        '毎月雑費金
        txtBillEzappi = Nz(rsData.Fields("PRIC_EZAPPI"), 0)
    End If

    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
'    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing                    'DELETE 2020/10/30

    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "FVS400_fncChkHoshoNewOld" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subHoshouShow
'        機能             :割引処理 (2017/12/20 add)
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subHoshouShow(Optional intNotInitProc As Integer = 0)

    Dim blnHoshoNew As Boolean
  
    On Error GoTo ErrorHandler

    '新旧判定
'2018/08/30 del 旧保証会社の処理は消す
'2018/08/30 del   If FVS400_fncChkHoshoNewOld(strBumonCode, cmbContainerHosyoCompany) = False Then 'Trueが新
'2018/08/30 del        If intNotInitProc = 0 Then
'2018/08/30 del        '2017/12/20 EGL
'2018/08/30 del'            '■旧フラン(初期)
'2018/08/30 del'            '毎月雑費名
'2018/08/30 del'            txtBillEzappiName = "補償加入費"
'2018/08/30 del'            '毎月雑費金
'2018/08/30 del'            txtBillEzappi = 480
'2018/08/30 del'            '保証委託料１
'2018/08/30 del            txtBillSecukgKei = IIf(txtBillSecukg = "" Or IsNull(txtBillSecukg), 0, txtBillSecukg)
'2018/08/30 del'            '保証委託料２(上部)
'2018/08/30 del            txtBillSecukg = txtBillSecukg
'2018/08/30 del'            '事務手数料
'2018/08/30 del'            If cmbBillAddFzappiCode1 = "" Then
'2018/08/30 del'                '登録されていない場合
'2018/08/30 del'                txtBillAddFzappi1 = P_JIMUTESURYO
'2018/08/30 del'            End If
'2018/08/30 del        '2017/12/20 EGL
'2018/08/30 del        End If
'2018/08/30 del        '保証金強制変更ﾁｪｯｸ
'2018/08/30 del '       chk保証金強制変更.Visible = True
'2018/08/30 del    Else
        '■新フラン
        '毎月雑費名 & 毎月雑費金
'        Call subGetEZappi(txtYou/No)                                           'DELETE 2018/02/19 N.IMAI
        Call subDispEzappiMessage                                               'INSERT 2018/02/19 N.IMAI
        If Val(Nz(txtContainerEzapii)) = Val(Nz(txtBillEzappi)) Then            'INSERT 2018/02/19 N.IMAI
            Call subGetEZappi(txtYoukNo)                                        'INSERT 2018/02/19 N.IMAI
        End If                                                                  'INSERT 2018/02/19 N.IMAI
        '毎月雑費の日割
  '      txtBillEzappiDaily = 0  '共益費は日割りはなし fncDailyRate(txtBillEzappi)
        '保証委託料１
        '▼2018/08/06 EGL ADD
        If P_HOSPLN_移行対象 = strNewHosplnIkou Then
            txtBillSecukgKei.ControlSource = ""                 'ADD 2018/10/23 tajima ※コントロールソースを無しにしないと値は設定できない
            txtBillSecukgKei = lngSecuKg
            txtBillSecukg = 0
        Else
        '▲2018/08/06 EGL ADD
            txtBillSecukgKei.ControlSource = "=[txtBillSecukg]" 'ADD 2018/10/23 tajima ※コントロールソース再設定
        '▼2018/10/02 EGL DEL
            'txtBillSecukgKei = 0
            '保証委託料２(上部)
            'txtBillSecukg = 0
            '事務手数料(電話の場合のみ且つ移動でない場合)
            '                           且つ新規の場合を条件に追加 2018/01/26
            'If blnネット予約 = False And Nz(txtMoveKbn.VALUE, 0) = P_MOVEKBN_無し And Nz(txtRcptNo, "") = "" Then
            '    Select Case strBumonCode
            '    Case "H" 'コンテナ
            '        txtBillAddFzappi1 = txtBillRentkgKei
            '    Case "8" 'トランク
            '        txtBillAddFzappi1 = 5400
            '    Case "1" 'オフィス
            '        If cmbContainerUsageCode.VALUE = 7 Then
            '            txtBillAddFzappi1 = 5400    'バーチャルオフィス(7)ならば add 2018/01/19
            '        Else
            '            txtBillAddFzappi1 = 21600   'バーチャルオフィス以外の場合
            '        End If
            '    End Select
            'End If
        '▲2018/10/02 EGL DEL
        '▼2018/08/06 EGL ADD
        End If
        '▲2018/08/06 EGL ADD
        '保証金強制変更ﾁｪｯｸ
        chk保証金強制変更.Visible = False
'2018/08/30 del    End If
    
    '◆初期処理
    If intNotInitProc = 0 Then
        '雑費コード
        cmbBillAddFzappiCode1 = P_ZAPPI_事務手数料
        '雑費コードオブジェクト設定
        Call cmbBillAddFzappiCode1_AfterUpdate
    End If
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subHoshouShow" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyicd_AfterUpdate
'        機能             :cmbContainerHosyicd_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyicd_AfterUpdate()

    txtContainerHosyicd = fncAfterUpdateCommon(cmbContainerHosyicd, P_DIVISION_LENGTH, P_CONTROL_数値)
    
    '↓INSERT 2008/11/13 SHIBAZAKI
    With cmbContainerHosyicd
        If Nz(txtContainerHosyicd, "") = "" Then
            txtContainerHosyiFrom = ""
            txtContainerHosyiTo = ""
            txtBillSecukgLabel = ""
        Else
            txtContainerHosyiFrom = Nz(.Column(2, .ListIndex), "")
            txtContainerHosyiTo = Nz(.Column(3, .ListIndex), "")
            txtBillSecukgLabel = Nz(.Column(4, .ListIndex), "")
        End If
    End With
    
    If txtContainerHosyiFrom = P_HOSYIFROM_事務手数料 Then
        cmbBillAddFzappiCode1 = Null
        Call cmbBillAddFzappiCode1_AfterUpdate
    Else
      '2017/12/20 delete start
'        If Nz(cmbBillAddFzappiCode1, "") = "" Then
'            cmbBillAddFzappiCode1 = P_ZAPPI_事務手数料
'            If IsNull(txtBillAddFzappi1) Then
'                txtBillAddFzappi1 = P_JIMUTESURYO
'                Call cmbBillAddFzappiCode1_AfterUpdate
'            End If
'        End If
       '2017/12/20 delete end
    End If
    '↑INSERT 2008/11/13 SHIBAZAKI
    
    '保証会社区分によって情報を変える
'    Select Case cmbContainerHosyicd                                        'DELETE 2008/11/13 SHIBAZAKI
'    Case P_HOSYICD_セーフティー, P_HOSYICD_トリオ, P_HOSYICD_リプラス      'DELETE 2008/11/13 SHIBAZAKI
    Select Case txtContainerHosyiFrom                                       'INSERT 2008/11/13 SHIBAZAKI
    Case P_HOSYIFROM_保証委託                                               'INSERT 2008/11/13 SHIBAZAKI
        txtSecurityHosyD.Enabled = True
        txtSecurityHosyB.Enabled = True
        Call YellowControl(txtBillSecukg, False) 'chg 2018/10/02 EGL 元 は true
        '▼ 2008/03/23 chg tajima 満了更新と自社理由なら保証委託料は０
        If Nz(txtMoveKbn, "") = P_MOVEKBN_満了更新 Or _
           (Nz(txtMoveKbn, "") = P_MOVEKBN_自社理由 And bln元ヤード解約 = True) Then
            txtBillSecukg = 0
        Else
'            txtBillSecukg = txtBillRentkg + txtBillEzappiKei  '2007/02/07 add tajima 保証料は月額の金額全て
            
            '▼2008/12/04 Mod hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。
            'txtBillSecukg = txtBillRentkg + get毎月雑費計
            '▼2008/12/20 Add hirano 保証料は、毎月額から月額割引をマイナスする。(chk保証金強制変更がチェックされていない場合)
            'txtBillSecukg = txtBillRentkg + get毎月雑費計 + txt月額分割引合計
            If Not chk保証金強制変更.VALUE Then
                'del 2018/10/02 EGL -->If chk割引有効可否.VALUE Then txtBillSecukg = txtBillMonthCharge + txt月額分割引
                Call subCompSecukgKase                                          'INSERT 2013/04/10 K.ISHIZAKA
            End If
            '▲2008/12/20 Mod hirano
            '▲2008/12/04 Mod hirano
        End If
        '▲ 2008/03/23 chg tajima
'        txtBillSecukgLabel = P_LABEL_保証委託料    'DELETE 2008/11/13 SHIBAZAKI
    Case Else
        txtSecurityHosyD = ""
        txtSecurityHosyD.Enabled = False
        txtSecurityHosyB = ""
        txtSecurityHosyB.Enabled = False
        Call YellowControl(txtBillSecukg, False)
'        txtBillSecukgLabel = P_LABEL_保証金        'DELETE 2008/11/13 SHIBAZAKI
    
        '↓INSERT 2008/09/25 SHIBAZAKI-> DELETE 2008/11/13 SHIBAZAKI
'        If cmbContainerHosyicd = P_HOSYICD_臨時 Then
'            txtBillSecukgLabel = P_LABEL_事務手数料
'        End If
        '↑INSERT 2008/09/25 SHIBAZAKI-> DELETE 2008/11/13 SHIBAZAKI
    End Select
    '2009/04/01 ADD <S> hirano 保証会社コード追加
    Select Case cmbContainerHosyicd
    'Case P_HOSYICD_保証会社無, P_HOSYICD_臨時, P_HOSYICD_その他 'DEL 2020/02/25 EGL
    Case P_HOSYICD_保証会社無, P_HOSYICD_臨時, P_HOSYICD_その他, P_HOSYICD_保証会社無_自社
        cmbContainerHosyoCompany = ""
        Call cmbContainerHosyoCompany_AfterUpdate
        cmbContainerHosyoCompany.Enabled = False
    Case P_HOSYICD_保証会社有
        cmbContainerHosyoCompany.Enabled = True
    Case Else
        cmbContainerHosyoCompany.Enabled = False
    End Select
    '2009/04/01 ADD <E>
    
    '2013/10/16 M.HONDA INS
    Select Case cmbContainerHosyicd
    Case P_HOSYICD_保証会社無
        Call YellowControl(cmb_更新区分, False)
    Case Else
        Call YellowControl(cmb_更新区分, True)
        '強制的に可へ
        cmb_更新区分 = 0
        Call cmb_更新区分_AfterUpdate
    End Select
    '2013/10/16 M.HONDA INS
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmbContainerHosyicd_BeforeUpdate
'        機能             :cmbContainerHosyicd_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyicd_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmbContainerHosyicd, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
   '2009/04/01 ADD <S> hirano 保証会社コード追加
    If blncmbContainerHosyicd Then

    'INSERT 2023/05/01 N.IMAI Start
    If cmbContainerHosyicd = "02" Then
        MsgBox "自社保証は選択できません。", vbInformation, P_KAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If
    'INSERT 2023/05/01 N.IMAI End
        'UPDATE 2021/11/30 M.HONDA
        'INSERT START 2020/02/25 EGL
'        If strHosyicd_Before = "02" Then
'            If cmbContainerHosyicd <> "02" Then
'                MsgBox "保証区分:02以外には変更できません。", vbInformation, P_KAKUNIN
'                Cancel = True
'                SendKeys "{ESC}"
'            End If
'        Else
'DELETE 2021/05/10 N.IMAI Start
'            MsgBox "保証区分:02に変更できません。", vbInformation, P_KAKUNIN
'            Cancel = True
'            SendKeys "{ESC}"
'DELETE 2021/05/10 N.IMAI End
        'UPDATE 2021/11/30 M.HONDA

'DELETE 2022/02/14 N.IMAI Start
'            'INSERT 2021/05/10 N.IMAI Start
''            If cmbContainerHosyicd = "02" Then
'                If ((cmbUserSkbn = "4" And cmbUserKkbn = "02") _
'                And (Nz(txtMoveKbn, "") = P_MOVEKBN_自社理由 _
'                    Or Nz(txtMoveKbn, "") = P_MOVEKBN_顧客理由 _
'                    Or Nz(txtMoveKbn, "") = P_MOVEKBN_未使用移動 _
'                    Or Nz(txtMoveKbn, "") = P_MOVEKBN_名義変更)) Then
'                    If MsgBox("保証区分:02に変更しますか？", vbQuestion + vbDefaultButton2 + vbYesNo, P_KAKUNIN) <> vbYes Then
'                        Cancel = True
'                        SendKeys "{ESC}"
'                    End If
''                Else
''                    MsgBox "保証区分:02に変更できません。", vbInformation, P_KAKUNIN
''                    Cancel = True
''                    SendKeys "{ESC}"
''                End If
'            End If
'            'INSERT 2021/05/10 N.IMAI End
''        End If
'        'INSERT START 2020/02/25 EGL
'DELETE 2022/02/14 N.IMAI End
        
        '画面で変更した場合
        '200904/23 ADD <S> hirano ネット予約は変更不可
        If blnネット予約 Then
            'ネット予約
            If strHosyicd_Before = cmbContainerHosyicd Then
                Cancel = True
            Else
                MsgBox "変更できません。", vbInformation, P_KAKUNIN
                Cancel = True
                SendKeys "{ESC}"
            End If
        Else
        '2009/04/23 ADD <E>
            Select Case strHosyicd_Before
            Case P_HOSYICD_リプラス, P_HOSYICD_リプラス包括
                Select Case cmbContainerHosyicd
                Case P_HOSYICD_リプラス, P_HOSYICD_リプラス包括
                    If strHosyicd_Before <> cmbContainerHosyicd Then
                        MsgBox "変更できません。", vbInformation, P_KAKUNIN
                        Cancel = True
                        SendKeys "{ESC}"
                    End If
                End Select
            Case Else
            
                '' 2013/06/04 M.HONDA START
                If txtMoveKbn.VALUE <> P_MOVEKBN_紹介者金額変更 Then
                    Select Case cmbContainerHosyicd
                    Case P_HOSYICD_リプラス, P_HOSYICD_リプラス包括
                        MsgBox "変更できません。", vbInformation, P_KAKUNIN
                        Cancel = True
                        SendKeys "{ESC}"
                    End Select
                End If
            End Select
                '' 2013/06/04 M.HONDA START
        End If '200904/23 INS
    End If
   '2009/04/01 ADD <E> hirano
End Sub

'==============================================================================*
's
'        機能             :cmbContainerHosyicd_KeyPres
'        MODULE_NAME      :cmbContainerHosyicd_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbContainerHosyicd_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmbContainerHosyicd, KeyAscii, P_DIVISION_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdCancel_Click
'        機能             :cmdCancel_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdCancel_Click()
    
    If Not blnEditMode Then
        Exit Sub
    End If
    
    If MsgBox("データ入力を取り消します。よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
        Exit Sub
    End If
    
    Call subFormRefresh
    
    '割引再計算
    'Call sub割引再計算(P_MODE_本計算)                                           'I:20081001 N.Imai
    '▼ 2008/11/28 hirano Mod 新規の場合は、仮計算、以外は、本計算
    If lblNewExist.Caption = P_LABEL_新規受付 Then
        Call sub割引再計算(P_MODE_仮計算)
    Else
        Call sub割引再計算(P_MODE_本計算)
    End If
    '▲ 2008/11/28 hirano Mod

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdComplete_Click
'        機能             :cmdComplete_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdComplete_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    If cmdSave.Enabled = True Then                                              'INSERT 2021/08/18 N.IMAI
        Call fncDataSave(P_SAVEMODE_完了, True)
    End If                                                                      'INSERT 2021/08/18 N.IMAI
    
    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdComplete_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdEnd_Click
'        機能             :cmdEnd_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdEnd_Click()

    Dim strMessage  As String
    Dim varDefBtn   As Variant
    
    '' 2013/06/04 M.HONDA START
'    strMessage = IIf(blnEditMode, "修正中ですが保存せず終了しますか？", "終了します。よろしいですか？")
'    varDefBtn = vbYesNo + vbQuestion + IIf(blnEditMode, vbDefaultButton2, vbDefaultButton1)
'
'    If MsgBox(strMessage, varDefBtn) = vbNo Then
'        Exit Sub
'    End If
'    On Error Resume Next
    '' 2013/06/04 M.HONDA END
    doCmd.Close acForm, Me.FormName

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdSave_Click
'        機能             :cmdSave_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdSave_Click()
On Error GoTo ErrorHandler

    If isネット予約担当 = True Then Exit Sub   '2007/08/23 add tajima

    If cmdSave.Enabled = True Then                                              'INSERT 2021/08/18 N.IMAI
        Call fncDataSave(P_SAVEMODE_保存, True)
    End If                                                                      'INSERT 2021/08/18 N.IMAI
    
    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdSave_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserMastMnt_Click
'        機能             :cmdUserMastMnt_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserMastMnt_Click()
    
    doCmd.OpenForm "FKS430", , , , , acDialog, Me.cmbUserCode & "," & Me.NAME

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmdUserPostConv_Click
'        機能             :cmdUserPostConv_Click イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdUserPostConv_Click()

    Call MSZ0000_YUBIB(Me.NAME, Me.txtUserYubino.NAME, Me.txtUserAdr1.NAME, _
                                Nz(Me.txtUserYubino.VALUE), Nz(Me.txtUserAdr1.VALUE))

    Call txtUserAdr1_AfterUpdate
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Open
'        機能             :Form_Open イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Open(Cancel As Integer)
On Error GoTo ErrorHandler
    
    Dim strParam()  As String

    Call MSHS010_M00(PROG_ID, "1", Nz(Me.OpenArgs))

    strFormCaption = Me.Caption
    strFormTitle = lblTitle.Caption
    strUserCodeOldValue = ""
    blnUpdate = False
    blncmbBillSeikyuCd = False  '2009/04/01 初期化
#If DEBUG_MODE Then
'    strParam = Split("FVS500,UH103055693,,H088057245", ",") '顧客移動
'    strParam = Split("FVS500,EH153025766,UH103055693,H088057245", ",") '顧客移動
    strParam = Split("FVS500,EH101005005,EH101005005,", ",")
#Else
    'パラメータ取得
    strParam = Split(Nz(Me.OpenArgs), ",")
#End If
    
    '権限取得
    pvONLET_ACCEI = fncGetKaiKeyAdmin() '2021/06/02 insert EGL

     '消費税の切替日を取得
    tax_from = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & Me.NAME & "' AND INTIF_RECFB='TAX_FROM'")) '2019/08/08 M.HONDA INS
    tax_to = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & Me.NAME & "' AND INTIF_RECFB='TAX_TO'"))   '2019/08/08 M.HONDA INS
    
    If UBound(strParam) <> P_パラメータ数 - 1 Then
        Call MSZZ024_M10("OpenArgs", "パラメータが正しくありません。")
        Exit Sub
    End If
    
    Call Show_RCPT_SHOUMEI(strParam(P_パラメータ_予約番号))                     'INSERT 2015/07/16 K.ISHIZAKA
    
    strFormFrom = strParam(P_パラメータ_呼び出し元画面)
    txtYoukNo = strParam(P_パラメータ_予約番号)
    
    blnEveryReload = False
    If Right(strFormFrom, 1) = "@" Then
        '取置き履歴から遷移した場合、呼び出し元（取置き履歴の呼び出し元）は
        '必ず再読み込みさせる
        strFormFrom = Mid(strFormFrom, 1, Len(strFormFrom) - 1)
        blnEveryReload = True
    End If
    txtRcptNo = strParam(P_パラメータ_受付番号)
    txtMotoAcptNo = strParam(P_パラメータ_元契約番号)
    
    'ヘルパー格納パス取得
'    strHelperPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""FVS400"" AND INTIF_RECFB = ""HELPER_PATH"""))
'    If Nz(strHelperPath) = "" Then
'        Call MSZZ024_M10("DLookup(""INTIF_RECDB"", ""INTI_FILE"", ""INTIF_PROGB=""FVS400"" AND INTIF_RECFB = ""HELPER_PATH"")", _
'                         "ヘルパーの格納パスを取得できませんでした。")
'    End If
    strHelperPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB=""FVS400"" AND INTIF_RECFB = ""HELPER_PATH_" & Trim(Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))) & """"))
    If Nz(strHelperPath) = "" Then
        Call MSZZ024_M10("DLookup(""INTIF_RECDB"", ""INTI_FILE"", ""INTIF_PROGB=""FVS400"" AND INTIF_RECFB = ""HELPER_PATH_" & Trim(Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))) & """)", _
                         "ヘルパーの格納パスを取得できませんでした。")
    End If
    If Right(strHelperPath, 1) <> "\" Then
        strHelperPath = strHelperPath & "\"
    End If
    
    'ページジャンプ設定
    Call subPageJumpInit
    
    '接続文字列取得
    Call fncGetConnectString
    
    '会社コード取得
    strCampCode = Trim(Nz(DLookup("CONT_SENDEN1", "dbo_CONT_MAST", "CONT_KEY=1")))
    
    '金融機関コード取得
    strKINYCode = Trim(Nz(DLookup("CONT_SENDEN2", "dbo_CONT_MAST", "CONT_KEY=1"))) 'INSERT 20060530 N.MIURA
    
    '地図ファイル格納パス取得
    strYardMapPath = Nz(DLookup("CONT_EXCEL_PATH", "dbo_CONT_MAST", "CONT_KEY=1"))
    If Nz(strYardMapPath) = "" Then
        Call MSZZ024_M10("DLookup(""CONT_EXCEL_PATH"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", _
                         "ヤード地図の格納パスを取得できませんでした。")
    End If
    If Right(strYardMapPath, 1) <> "\" Then
        strYardMapPath = strYardMapPath & "\"
    End If
    
    '割引上限取得
    lngDCMax = Trim(Nz(DLookup("CONT_DC_MAX", "dbo_CONT_MAST", "CONT_KEY=1"), 0)) 'I:20081001 N.Imai
    
    '保証会社初期値取得
    strネット予約初期保証会社Cd = getネット予約クレカ保証会社 '2011/06/18 add tajima
    
    '受付トランを検索
    Call subFormRefresh
    
    '割引再計算
    
'    ''▼ 2008/11/19 hirano Mod 新規の場合は、仮計算、以外は、本計算
'    ''Call sub割引再計算(P_MODE_本計算)                                           'I:20081001 N.Imai

'    '2009/11/10 MOD <S> 移動の場合でも　新規以外は、再計算する
'    'If Nz(txtMotoAcptNo, "") = "" Then '移動の場合は初っ端で割引計算はしない
'    If Nz(txtMotoAcptNo, "") = "" Or (Nz(txtMotoAcptNo, "") <> "" And chk割引有効可否.VALUE) Then  '----2010/05/28--del--ryu--
'    '2009/11/10 MOD <E>
        
    '--2010/05/28--: sub割引再計算（）に割引有効可否の判断があるから、こちらで判断必要ないです。
    '--2010/05/28--: ⇒上記2009/11/10追加したIF条件を削除
    If lblNewExist.Caption = P_LABEL_新規受付 Then
        Call sub割引再計算(P_MODE_仮計算)
    Else
        Call sub割引再計算(P_MODE_本計算)
    End If
    
'    End If      '----2010/05/28--del--ryu--
'    ''▲ 2008/11/19 hirano Mod
    
    '' 2013/06/11 M.HONDA START
    ' 他の契約を検索
    If cmbUserCode <> "" Then
        Call subOtherCarg
    End If
    
    '2015/03/20 M.HONDAINS
    'Me.lbl_msg.Caption = Me.lbl_msg.Caption & StrConv(Space(60), vbWide) _
                              & "Tポイントカード番号・承認番号の入力漏れに注意してください。"
    'Me.TimerInterval = 100
    '2015/03/20 M.HONDAINS
    '' 2013/06/11 M.HONDA END
    
    '' ﾄﾘﾌﾟﾙｷｬﾝﾍﾟｰﾝが適用されているかﾁｪｯｸ
    chkCampaign.VALUE = fncGetトリプル(txtAcptNo)
    strSaveトリプル = chkCampaign.VALUE
    
    
    If Me.chkNEN_KBN.VALUE = -1 Then
        If strBumonCode <> "H" Then
            If wk割引情報.件数 = 0 Then
                txt年払い = (txtBillRentkgKei + txtBillEzappiKei) * -1
            Else
                If wk割引情報.TRAN_DATA(0).DCNTM_PERIOD >= 12 Then
                    txt年払い = (txtBillRentkgKei + txtBillEzappiKei + wk割引情報.TRAN_DATA(0).VALUE_TRUE) * -1
                Else
                    txt年払い = (txtBillRentkgKei + txtBillEzappiKei) * -1
                End If
            End If
        End If
    End If

'    '2018/07/17 M.HONDA INS
'    If txtMoveKbn = "1" Or txtMoveKbn = "5" Or txtMoveKbn = "7" Or txtMoveKbn = "2" Or txtMoveKbn = "3" Or txtMoveKbn = "10" Then
'
'    Else
'
'        Dim SEQ As Variant
'        'INSERT 2018-01-20 add k.sato
'
'
'        If txtRcptNo.VALUE = "" Then
'            txtSecurityHosyB.Enabled = False
'            SEQ = fncGETHOSESEQ(cmbContainerHosyoCompany.VALUE)
'            If SEQ = "0" Then
'            txtSecurityHosyB.Enabled = True
'            Else
'                txtSecurityHosyB.VALUE = SEQ
'            End If
'        Else
'            txtSecurityHosyB.Enabled = True
'        End If
'    End If

    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("Form_Open", True)
    Err.Clear

End Sub
'==============================================================================*
'
'       MODULE_NAME     : ﾄﾘﾌﾟﾙｷｬﾝﾍﾟｰﾝ取得
'       MODULE_ID       : fncGetトリプル
'       CREATE_DATE     :
'       PARAM           :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncGetトリプル(strACPTNO As String) As Integer

    Dim objAdoDbConnection  As Object
    Dim rsData              As Object
    Dim strSQL              As String
    Dim iIndex              As Integer
    Dim strDcntNo           As String
    Dim strKaseSqlSvr       As String

    On Error GoTo Exception
    
    'SQL-Server名
'   strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))'DELETE 2019/03/07 Y.WADA
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'")) 'INSERT 2019/03/07 Y.WADA

    '適用割引割付情報取得
    'ﾄﾘﾌﾟﾙｷｬﾝﾍﾟｰﾝの割付があるかﾁｪｯｸ
    strSQL = ""
    strSQL = strSQL + "SELECT * "
    strSQL = strSQL + "  FROM DCRA_TRAN "
    strSQL = strSQL + "      ," & strKaseSqlSvr & ".dbo.DCNT_MAST "
    strSQL = strSQL + " WHERE DCRAT_ACPTNO     = '" & strACPTNO & "' "
    strSQL = strSQL + "   AND DCNTM_NO         = DCRA_TRAN.DCRAT_DCNT_NO "
    strSQL = strSQL + "   AND DCNTM_BUMOC      = '" & strBumonCode & "' "
    strSQL = strSQL + "   AND DCRAT_DCNT_NO    = '000009' "
    strSQL = strSQL + " ORDER BY DCNTM_NO, DCRAT_FROM "
    Set objAdoDbConnection = MSZZ025.ADODB_Connection(strBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL, objAdoDbConnection)
    '初期値設定
    iIndex = 0
    If rsData.EOF = False Then
        fncGetトリプル = -1
    Else
        fncGetトリプル = 0
    End If
    
Exception:

    If Not rsData Is Nothing Then: rsData.Close: Set rsData = Nothing
    If Not objAdoDbConnection Is Nothing Then: objAdoDbConnection.Close: Set objAdoDbConnection = Nothing
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncGetトリプル" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :GetOtherCargDataSql
'        機能             :データ検索SQL作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetOtherCargDataSql(userCode As String, RCPTNO As String) As String

    Dim strSQL      As String
    Dim today       As String
    
    today = Replace(DATE, "/", "")
    
    strSQL = "          select COUNT(*) CNT "
    strSQL = strSQL & " from   CARG_FILE, "
    strSQL = strSQL & " RCPT_TRAN"
    strSQL = strSQL & " where  CARG_UCODE = " & userCode
    strSQL = strSQL & " and    CARG_ACPTNO = RCPT_CARG_ACPTNO "
    strSQL = strSQL & " and    RCPT_NO <> '" & RCPTNO & "'"
    strSQL = strSQL & " and    CARG_HOSYO_CD IN ('970031') "
    strSQL = strSQL & " and    isnull(CONVERT(VARCHAR,CARG_KYDATE,112),'29991231') > '" & today & "'"

    GetOtherCargDataSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subOtherCarg
'        機能             :データ検索
'        IN               :
'                         :
'                         :
'                         :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function subOtherCarg() As Boolean

    On Error GoTo ErrorHandler
    
    Dim dbSQLServer     As Database
    Dim rsData          As Recordset
    Dim blnError        As Boolean
    
    blnError = False
        
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    Set rsData = dbSQLServer.OpenRecordset(GetOtherCargDataSql(cmbUserCode, txtYoukNo), dbOpenSnapshot, dbSQLPassThrough)
    
    If rsData.EOF Then
        subOtherCarg = True
        Exit Function
    End If
    
    If rsData.Fields("CNT") <> 0 Then
'        Me.lbl_msg.Caption = StrConv(Space(60), vbWide) _
           & "既にパルマ包括契約データが存在します。 請求は合算できないので注意してください。"
    End If
        
    subOtherCarg = True
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "fncGetData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :Form_Activate
'        機能             :Form_Activate イベント
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
'        MODULE_NAME      :Form_KeyDown
'        機能             :Form_KeyDown イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

   ' ▼2006/06/24 Add tajima
   Call CmReprMod.OpenFVS90n(KeyCode, _
                             Me.txtContainerYcode, _
                             Me.txtContainerCno, _
                             Me.cmbUserCode, _
                             Me.txtAcptNo _
                             )
  ' ▲2006/06/24 Add tajima

    'F1=「終了」 F2=「保存」
    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            Me.cmdEnd.SetFocus
        Case vbKeyF2
            KeyCode = 0
            Me.cmdSave.SetFocus
        Case vbKeyF3
            KeyCode = 0
            Me.cmdCancel.SetFocus
        Case vbKeyF5 To vbKeyF12
            KeyCode = 0
    End Select

End Sub

'==============================================================================*
'
'        MODULE_NAME      :Form_Unload
'        機能             :Form_Unload イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    Dim objForm         As Object

    Set objForm = Forms(strFormFrom)
    If Not objForm Is Nothing Then
        Call objForm.subReload(IIf(blnUpdate Or blnEveryReload, 1, 0))
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSendAddressClear
'        機能             :鍵送付先クリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSendAddressClear()

    txtSendName = ""
    txtSendYubino = ""
    txtSendAdr1 = ""
    txtSendAdr2 = ""
    txtSendAdr3 = ""
    txtSendTel = ""
    cmbSendHonorific = ""
    Call cmbSendHonorific_AfterUpdate
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSendAddressYellow
'        機能             :鍵送付先使用可不可設定
'        IN               :True=黄色　False=白
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSendAddressYellow(blnYellow As Boolean)

    Call YellowControl(txtSendName, blnYellow)
    Call YellowControl(txtSendYubino, blnYellow)
    Call YellowControl(txtSendAdr1, blnYellow)
    Call YellowControl(txtSendAdr2, blnYellow)
    Call YellowControl(txtSendAdr3, blnYellow)
    Call YellowControl(txtSendTel, blnYellow)
    Call YellowControl(cmbSendHonorific, blnYellow)
    
    cmdSendPostConv.Enabled = Not blnYellow

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserToSendAdd
'        機能             :顧客情報から鍵送付先へコピー
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserToSendAdd()

    txtSendName = Trim(Nz(txtUserName))
    txtSendYubino = Trim(Nz(txtUserYubino))
    txtSendAdr1 = Trim(Nz(txtUserAdr1))
    txtSendAdr2 = Trim(Nz(txtUserAdr2))
    txtSendAdr3 = Trim(Nz(txtUserAdr3))
    
    If Trim(Nz(txtUserTel)) <> "" Then
        txtSendTel = Trim(Nz(txtUserTel))
    ElseIf Trim(Nz(txtUserKeitai)) <> "" Then
        txtSendTel = Trim(Nz(txtUserKeitai))
    Else
        txtSendTel = Trim(Nz(txtUserFax))
    End If
    
    cmbSendHonorific = IIf(cmbUserKkbn = P_USER_KKBN_法人, P_HONORIFIC_御中, P_HONORIFIC_様)
    Call cmbSendHonorific_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserSendToSendAdd
'        機能             :顧客送付先情報から鍵送付先へコピー
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserSendToSendAdd()

    txtSendName = Trim(Nz(txtUserRname))
    txtSendYubino = Trim(Nz(txtUserRpost))
    txtSendAdr1 = Trim(Nz(txtUserRadr1))
    txtSendAdr2 = Trim(Nz(txtUserRadr2))
    txtSendAdr3 = Trim(Nz(txtUserRadr3))
    
    If Trim(Nz(txtUserRtel)) <> "" Then
        txtSendTel = Trim(Nz(txtUserRtel))
    ElseIf Trim(Nz(txtUserRkeitai)) <> "" Then
        txtSendTel = Trim(Nz(txtUserRkeitai))
    Else
        txtSendTel = Trim(Nz(txtUserRfax))
    End If
    
    cmbSendHonorific = P_HONORIFIC_様
    Call cmbSendHonorific_AfterUpdate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fraUserNewExist_AfterUpdate
'        機能             :fraUserNewExist_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub fraUserNewExist_AfterUpdate()
    

    'True=既存 False=新規
    Dim blnNew      As Boolean
    
    blnNew = IIf(fraUserNewExist.VALUE = P_USER_OPT_新規, True, False)
    
    'ここで顧客情報クリア
    Call subUserClear
    
    'フリガナ／コードは既存のときのみ使用可
    txtUserExistKana.Enabled = Not blnNew
    cmbUserCode.Enabled = Not blnNew
    cmdUserMastMnt.Enabled = Not blnNew
    
    chkUserRinput.Enabled = blnNew
    

    Call subUserInfoEnable(blnNew)             '顧客情報使用可不可設定
    
    chkUserRinput.VALUE = 0
    Call chkUserRinput_AfterUpdate
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserClear
'        機能             :顧客情報クリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserClear()

    txtUserExistKana = Null
    Call txtUserExistKana_AfterUpdate
    
    txtUserName = ""
    txtUserKana = ""
    txtUserYubino = ""
    txtUserAdr1 = ""
    txtUserAdr2 = ""
    txtUserAdr3 = ""
    txtUserTanm = ""
    txtUserTaka = ""
    txtUserTel = ""
    txtUserKeitai = ""
    txtUserFax = ""
    txtUserMail = ""
    txtUserBirthday = ""                                'INSERT 2010/04/20 SHIBAZAKI
    
    cmbUserKkbn = ""
    Call cmbUserKkbn_AfterUpdate
    
    cmbUserSkbn = ""
    Call cmbUserSkbn_AfterUpdate
    
    txtUserSday = ""
    txtUserBiko = ""
    
    Call subRenrakusakiClear
    Call subKinmusakiClear                              'INSERT 2010/04/20 SHIBAZAKI
    
    '↓INSERT 2014/10/14 tajima
    cmb_性別 = ""
    Call cmb_性別_AfterUpdate
    '↑INSERT 2014/10/14 tajima

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subRenrakusakiClear
'        機能             :顧客情報書類送付先クリア
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subRenrakusakiClear()

    txtUserRname = ""
    txtUserRkana = ""
    txtUserRpost = ""
    txtUserRadr1 = ""
    txtUserRadr2 = ""
    txtUserRadr3 = ""
    txtUserRtel = ""
    txtUserRkeitai = ""
    txtUserRfax = ""

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserInfoEnable
'        機能             :顧客情報使用可不可設定処理
'        IN               :True=使用可に設定 False=使用不可に設定
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserInfoEnable(blnEnable As Boolean)

    Call YellowControl(txtUserName, Not blnEnable)
    Call YellowControl(txtUserKana, Not blnEnable)
    Call YellowControl(txtUserYubino, Not blnEnable)
    
    cmdUserPostConv.Enabled = blnEnable
    
    Call YellowControl(txtUserAdr1, Not blnEnable)
    Call YellowControl(txtUserAdr2, Not blnEnable)
    Call YellowControl(txtUserAdr3, Not blnEnable)
    Call YellowControl(cmbUserKkbn, Not blnEnable)
    Call YellowControl(cmbUserSkbn, Not blnEnable)
    Call YellowControl(txtUserTanm, Not blnEnable)
    Call YellowControl(txtUserTaka, Not blnEnable)
    Call YellowControl(txtUserTel, Not blnEnable)
    Call YellowControl(txtUserKeitai, Not blnEnable)
    Call YellowControl(txtUserFax, Not blnEnable)
    Call YellowControl(txtUserMail, Not blnEnable)
    Call YellowControl(txtUserSday, Not blnEnable)
    Call YellowControl(txtUserBiko, Not blnEnable)
    Call YellowControl(txtUserBirthday, Not blnEnable)                  'INSERT 2010/04/21 SHIBAZAKI
    Call YellowControl(cmb_性別, Not blnEnable)                         'INSERT 2014/10/14 tajima
    
    Call subUserKinmusakiEnable(blnEnable)                              'INSERT 2010/04/21 SHIBAZAKI
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserRenrakusakiEnable
'        機能             :顧客情報使用可不可設定処理
'        IN               :True=使用可に設定 False=使用不可に設定
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserRenrakusakiEnable(blnEnable As Boolean)

    Call YellowControl(txtUserRname, Not blnEnable)
    Call YellowControl(txtUserRkana, Not blnEnable)
    Call YellowControl(txtUserRpost, Not blnEnable)
    
    cmdUserRpostConv.Enabled = blnEnable
    
    Call YellowControl(txtUserRadr1, Not blnEnable)
    Call YellowControl(txtUserRadr2, Not blnEnable)
    Call YellowControl(txtUserRadr3, Not blnEnable)
    Call YellowControl(txtUserRtel, Not blnEnable)
    Call YellowControl(txtUserRkeitai, Not blnEnable)
    Call YellowControl(txtUserRfax, Not blnEnable)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :YellowControl
'        機能             :黄色／白色処理
'        IN               :True=黄色に設定 False=白色に設定
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub YellowControl(ctrControl As Control, blnYellow As Boolean)

    '表示のみは黄色。入力可は白色。
    ctrControl.Enabled = Not blnYellow
    ctrControl.Locked = blnYellow
    If blnYellow Then
        ctrControl.BackColor = P_YELLOW
    Else
        ctrControl.BackColor = P_WHITE
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subFormRefresh
'        機能             :画面再更新
'        IN               :①True=リフレッシュ後、先頭ページに移動
'                         :②True:セキュリティーカード情報表示 False:表示しない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/06/01 SHIBAZAKI セキュリティカード対応
'Private Sub subFormRefresh()
Private Sub subFormRefresh(Optional blnPageTop As Boolean = True, Optional blnDispSCard As Boolean = True)


    lblNewExist.Caption = IIf(txtRcptNo = "", P_LABEL_新規受付, P_LABEL_既存受付)

    Call fncGetData(lblNewExist.Caption = P_LABEL_既存受付, txtMotoAcptNo <> "", blnPageTop, blnDispSCard)
    
    '変更モード設定　新規受付の場合は編集中とする
    Call subEditMode(lblNewExist.Caption = P_LABEL_新規受付)
    '2009/04/23 ネット予約の初回の場合は、保存させるため、編集中にさせる。
    If blnネット予約初回 Then
        Call subEditMode(blnネット予約初回)
    End If
    '2006/04/28 S.SHIBAZAKI 新規受付時は、「受付取消」と「保証会社NG」ボタンを使用不可にする。
    cmdRcptCancel.Enabled = (lblNewExist.Caption = P_LABEL_既存受付)
    cmdSecuNg.Enabled = (lblNewExist.Caption = P_LABEL_既存受付)
    
    '▼2008/12/04 Mod hirano 保証金変更フラグ　受付トラン保持
    'Call YellowControl(txtBillSecukg, True)  '2007/02/07 add tajima
    'chk保証金強制変更.VALUE = False '2007/02/07 add tajima
    If lblNewExist.Caption = P_LABEL_新規受付 Then
        'Call YellowControl(txtBillSecukg, True) del 2018/10/02 EGL.t
        chk保証金強制変更.VALUE = False
    End If
    '▲2008/12/04 Mod hirano
    
    '初期表示データを保存
    '2009/03/09 MOD <S> Nz付加
    'strSave月額料金 = txtBillMonthCharge                                      'I:20090223 N.Imai
    'strSave初回料金 = txtBillFirstkg                                          'I:20090225 N.Imai
    strSave月額料金 = Nz(txtBillMonthCharge, "")                                      'I:20090223 N.Imai
    strSave初回料金 = Nz(txtBillFirstkg, "")                                          'I:20090225 N.Imai
    '2009/03/09 MOD <E>
        strSave契約開始日 = txtBillStDate                                           'I:20090223 N.Imai
    strSave解約予定日 = txtReceiptCyDate                                        'I:20090223 N.Imai
    strSave利用予定期間 = cmb_利用予定期間                                       'I:20090223 N.Imai
    strSave請求期間 = txtBillSeikyuKikan                                        'I:20090223 N.Imai
    blnSave割引する = chk割引有効可否                                            'I:20090223 N.Imai
    blnSave割引再計算 = False                                                   'I:20090223 N.Imai
    strSave保証会社 = cmbContainerHosyoCompany      '2011/06/18 add tajima
'↓ UPDATE 2014/04/17 MIYAMOTO
'    strSave毎月回収方法 = cmbRcptKaihi              '2011/06/18 add tajima
    strSave毎月回収方法 = Nz(cmbRcptKaihi)
'↑ UPDATE 2014/04/17 MIYAMOTO
    strSaveトリプル = chkCampaign.VALUE             '2013/10/25 add M.HONDA
    
    Call sub割引再計算(P_MODE_仮計算)                                           'INSERT 2026/03/31 N.IMAI
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncGetData
'        機能             :データ検索
'        IN               :①True=受付トランを検索 False=予約トランを検索
'                         :②True:移動予約 False:通常予約
'                         :③True:表示後先頭ページへ移動 False:ページ切り替えしない
'                         :④True:セキュリティーカード情報表示 False:表示しない
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/06/01 SHIBAZAKI セキュリティカード対応
'Private Function fncGetData(blnRcpt As Boolean, blnMove As Boolean) As Boolean
Private Function fncGetData(blnRcpt As Boolean, _
                            blnMove As Boolean, _
                            Optional blnPageTop As Boolean = True, _
                            Optional blnDispSCard As Boolean = True) As Boolean

    On Error GoTo ErrorHandler
    
    Dim dbSQLServer     As Database
    Dim rsData          As Recordset
    Dim blnError        As Boolean
    
    '2006/09/06 SHIBAZAKI 顧客とコンテナの修繕情報を表示する
    Dim dbKagtos                As Database
    Dim strKagtosSqlserver      As String
    Dim strKagtosConnect        As String
    Dim strHOKAC                As String                                       'INSERT 2013/04/10 K.ISHIZAKA
    Dim varHOKAC                As Variant                                      'INSERT 2013/04/10 K.ISHIZAKA
    
' ***** Start 2017/07/31 add by ysuzuki
    Dim strHasseiKbn            As String   '発生区分
    Dim strShoki_Seikyu         As String   '初期費用請求方法
    Dim lngUcode                As Long     '顧客コード
    Dim strACPTNO               As String   '契約番号
    
' ***** end 2017/07/31 add by ysuzuki
    
    Dim objCon                  As Object                                       'INSERT 2021/03/31 N.IMAI
    
    blnError = False
    fncGetData = False
    blnネット予約初回 = False   '2009/04/23
    doCmd.Hourglass True
        
    '2020/12/01 [EGL] INS Start
    Call subクーポンCDコンボ値セット
    '2020/12/01 [EGL] INS End
    
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    Set rsData = dbSQLServer.OpenRecordset(fncMakeGetDataSql(blnRcpt, blnMove), dbOpenSnapshot, dbSQLPassThrough)
    
    If rsData.EOF Then
        'データが存在しない
        'データを検索できないのは何らかの問題が発生しているはずなので、不正終了する
        Call MSZZ024_M10("OpenRecordset", "データを参照できませんでした。")
    End If
    
    '▼2018/08/30 EG
    lngFZAPPI_CODE = Nz(rsData.Fields("FZAPPI_CODE"), 0) '初回雑費コード
    lngEZAPPI_CODE = Nz(rsData.Fields("EZAPPI_CODE"), 0) '毎月雑費コード
    '▲2018/08/30 EGL
    
    
    '▼2018/08/06 EGL INS<S>
    strNewHosplnIkou = Nz(rsData.Fields("NEW_HOSPLN_IKOU"), P_HOSPLN_移行対象外)
    lngSecuKg = Nz(rsData.Fields("SECUKG"), 0)
    '▲2018/08/06 EGL INS<E>
    
    strプランCd = Nz(rsData.Fields("RCPT_PLAN_CD"), "0000") '2020/09/28 EGL add

    'データ表示
    Call subDispData(rsData, blnRcpt, blnMove)
    
    '2006/09/06 SHIBAZAKI 顧客とコンテナの修繕情報を表示する↓
    '加瀬ＤＢに接続
    strKagtosSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    strKagtosConnect = MSZZ007_M00()
    Set dbKagtos = Workspaces(0).OpenDatabase(strKagtosSqlserver, dbDriverNoPrompt, False, strKagtosConnect)
    
    If Nz(cmbUserCode, "") <> "" Then
        '顧客の直近修繕（コメント／クレーム）
        txtUserMntInfo = GetUserLatestKotoText(dbKagtos, strBumonCode, cmbUserCode)
    End If
    
    'コンテナの直近修繕
    txtContainerMntInfo = GetRoomLatestReprText(dbKagtos, strBumonCode, txtContainerYcode, txtContainerCno)
    
    If P_HOSPLN_移行対象外 = strNewHosplnIkou Then  '2018/08/06 EGL INS
        '2009/04/01 ADD<S>
        '保証会社情報取得
        If Not blnRcpt Then
            '新規
            If Not funcGetShiireInfo(dbSQLServer, dbKagtos, strBumonCode, rsData, blnRcpt) Then
                GoTo ExitRtn
            End If
        Else
            '変更
            '2009/04/22 ネット予約を考慮
            If cmbContainerHosyicd = Format$(P_HOSYICD_パルマ, "00") Or (cmbContainerHosyoCompany <> "" And txtContainerHosyoCampany = "") Then
                If Not funcGetShiireInfo(dbSQLServer, dbKagtos, strBumonCode, rsData, blnRcpt) Then
                    GoTo ExitRtn
                End If
            End If
        End If
        '2009/04/01 ADD<E>
    End If  '2018/08/06 TAKENOUCHI INS
    
    strHosyo_cd_Before = cmbContainerHosyoCompany   '2009/04/23 ADD
    strHosyicd_Before = cmbContainerHosyicd         '2009/04/23 ADD
    dbKagtos.Close
    Set dbKagtos = Nothing
    '2006/09/06 SHIBAZAKI 顧客とコンテナの修繕情報を表示する↑
    
    If blnRcpt Then                                                             'INSERT START 2013/04/10 K.ISHIZAKA
        txtBillSecukgKase.VALUE = Nz(rsData.Fields("SECUKG_KASE"), 0)
        chkSecukgKase.VALUE = (Val(txtBillSecukgKase.VALUE) <> 0)
    Else
        chkSecukgKase.VALUE = False
        '2015/07/24 M.HONDA DEL
        'If txtMoveKbn.VALUE = P_MOVEKBN_顧客理由 Then
        '    strHOKAC = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='FVS400' AND INTIF_RECFB = 'HOKAC'"))
        '    If strHOKAC <> "" Then
        '        For Each varHOKAC In Split(strHOKAC, ",")
        '            If Nz(cmbContainerHosyoCompany.VALUE) = varHOKAC Then
        '                chkSecukgKase.VALUE = True
        '                Call subCompSecukgKase
        '                Exit For
        '            End If
        '        Next
        '    End If
        'End If
    End If
    txtBillSecukg.Visible = Not chkSecukgKase.VALUE
    txtBillSecukgKase.Visible = chkSecukgKase.VALUE
    Call subDispSecukgMessage
    Call subDispEzappiMessage
    If (Val(Nz(txtBillEzappi)) = Val(Nz(txtBillEzappiMoto))) _
        And (Val(Nz(txtBillRentkg)) = Val(Nz(txtBillRentkgMoto))) Then
        opgRentkg.VALUE = opbRentkgHIKI.OptionValue
    Else
        opgRentkg.VALUE = opbRentkgSHIN.OptionValue
    End If                                                                      'INSERT END   2013/04/10 K.ISHIZAKA
    
    '2006/06/01 SHIBAZAKI セキュリティカード対応
    If blnDispSCard Then
        Call fncDispSCard(dbSQLServer, blnRcpt)
    End If
    
    '2006/06/01 SHIBAZAKI セキュリティカード対応
    If blnPageTop Then
        '先頭にスクロール
        cmbPage = cmbPage.Column(0, 0)
        Call cmbPage_AfterUpdate
    End If
    
    'If Nz(txtMotoAcptNo.VALUE) = "" Then                                       'DELETE 2021/08/18 N.IMAI 'INSERT 2015/07/26 K.ISHIZAKA
    If Nz(txtMotoAcptNo.VALUE) = "" And cmd_RCPT_SHOUMEI_KDATE.Enabled Then     'INSERT 2021/08/18 N.IMAI
        cmd_RCPT_SHOUMEI_KDATE.Enabled = (Nz(rsData.Fields("RCPT_SHOUMEI_KDATE"), "") = "") 'INSERT 2015/07/16 K.ISHIZAKA
    Else                                                                        'INSERT START 2015/07/26 K.ISHIZAKA
        cmd_RCPT_SHOUMEI_KDATE.Enabled = False
    End If                                                                      'INSERT END   2015/07/26 K.ISHIZAKA
    
' ***** Start 2017/07/31 add by ysuzuki
    
    ' 画面の値を取得
    strHasseiKbn = IIf(Not blnネット予約, 0, 1)             '発生区分
    strShoki_Seikyu = Nz(Me.cmbBillSeikyuCd.VALUE, "")      '初期費用請求方法
    lngUcode = Val(Me.cmbUserCode.VALUE)                    '顧客コード
    strACPTNO = Nz(Me.txtAcptNo.VALUE, "")                  '契約番号
    
    If Nz(txtMotoAcptNo.VALUE) = "" Then  '2017/10/07 add
        Set objCon = ADODB_Connection(strBumonCode)
        'コンテナ契約ファイル、請求ファイルへの登録対象かどうか取得
        'lngStRtn = CmKomsMod.CheckRegiTarget(dbSQLServer, strHasseiKbn, strShoki_Seikyu, lngUcode, strACPTNO)  'DELETE 2021/03/31 N.IMAI
        lngStRtn = CmKomsMod.CheckRegiTarget(objCon, strHasseiKbn, strShoki_Seikyu, lngUcode, strACPTNO)        'INSERT 2021/03/31 N.IMAI
    Else                            '2017/10/07 add
        lngStRtn = 0                '2017/10/07 add
    End If                          '2017/10/07 add
    
' ***** end 2017/07/31 add by ysuzuki
    
'↓ADD 2019/06/24 egl.YSUZUKI
    If blnRcpt Then
        '証明アップロードなしで、フォロー一回以上ありの場合
        If Nz(rsData.Fields("RCPT_SHOUMEI_UP_UM"), "0") = 0 And Nz(rsData.Fields("RCPT_SHOUMEI_FCNT"), 0) > 0 Then
            '督促メッセージを表示
            txtTokusokuDisp.VALUE = "自動督促:" & Nz(rsData.Fields("RCPT_SHOUMEI_FCNT"), 0) & "回,ﾒｰﾙ送信:" & Format(rsData.Fields("RCPT_SHOUMEI_FDATE"), "mm/dd HH:mm")
            txtTokusokuDisp.Visible = True
        Else
            txtTokusokuDisp.Visible = False
        End If
    End If

'↑ADD 2019/06/24 egl.YSUZUKI
    
    fncGetData = True
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True
ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing
    If Not dbKagtos Is Nothing Then dbKagtos.Close: Set dbKagtos = Nothing

    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "fncGetData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeGetDataSql
'        機能             :データ検索SQL作成
'        IN               :①True=受付トランを検索 False=予約トランを検索
'                         :②True:移動予約 False:通常予約
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeGetDataSql(blnRcpt As Boolean, blnMove As Boolean) As String

    Dim strSQL      As String
    
    strSQL = " SELECT " & fncMakeCommonSelect(blnMove)
    
    'SELECT句
    If blnRcpt Then
        strSQL = strSQL & fncMakeRcptSelect(blnMove)
    Else
        strSQL = strSQL & fncMakeYoukSelect(blnMove)
    End If
    
    'FROM句
    strSQL = strSQL & fncMakeFrom(blnRcpt, blnMove)
    
    'WHERE句
    If blnRcpt Then
        strSQL = strSQL & "  WHERE RCPT_TRAN.RCPT_NO = '" & txtRcptNo & "' "
    Else
        strSQL = strSQL & "  WHERE YOUKT_UKNO = '" & txtYoukNo & "' "
    End If
    
    fncMakeGetDataSql = strSQL
'    Debug.Print (strSql)
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeYoukSelect
'        機能             :データ検索SQL作成(新規受付SELECT句)
'        IN               :True=元契約をJOINする　False=元契約はない
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeYoukSelect(blnMove As Boolean) As String


    Dim strSQL      As String
    Dim str契約移動条件sql As String   '2018/08/06 EGL INS
     
    '▼2008/03/23 tajima chg&add
    '移動で満了更新or解約ヤードからの加瀬移動の場合は元受付から取得
    Dim str元受取条件sql As String
    If blnMove Then
        'キャンペーン情報は満了更新のときのみ
        str元受取条件sql = " YOUKT_MOVEKBN = " & P_MOVEKBN_満了更新

        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_SEV1N ELSE NOW_YARD.YARD_SEV1N END AS YARD_SEV1N,"
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_SEV2N ELSE NOW_YARD.YARD_SEV2N END AS YARD_SEV2N, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_SEV3N ELSE NOW_YARD.YARD_SEV3N END AS YARD_SEV3N, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_ENDEN ELSE NOW_YARD.YARD_ENDEN END AS YARD_ENDEN, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_SEV_EXPIRE_MONTH ELSE NOW_YARD.YARD_SEV_EXMONTH END AS YARD_SEV_EXMONTH, "
    
    Else
    '▼ 2007/08/23 tajima add
        strSQL = strSQL & "        NOW_YARD.YARD_SEV1N, "
        strSQL = strSQL & "        NOW_YARD.YARD_SEV2N, "
        strSQL = strSQL & "        NOW_YARD.YARD_SEV3N, "
        strSQL = strSQL & "        NOW_YARD.YARD_ENDEN, "
        strSQL = strSQL & "        NOW_YARD.YARD_SEV_EXMONTH, "
    '▲ 2007/08/23 tajima add
    End If
    '▲ 2008/03/23 tajima chg&add
    
    '▼2018/08/06 EGL INS<S>
    If blnMove Then
        str契約移動条件sql = "YOUKT_MOVEKBN IN(" & P_MOVEKBN_満了更新 & "," & P_MOVEKBN_自社理由 & ") AND MOTO_RCPT.RCPT_NEW_HOSPLN_IKOU = '" & P_HOSPLN_移行対象 & "'"
        
        strSQL = strSQL & " CASE WHEN " & str契約移動条件sql & " THEN MOTO_RCPT.RCPT_NEW_HOSPLN_IKOU ELSE '" & P_HOSPLN_移行対象外 & "' END AS NEW_HOSPLN_IKOU, "
    Else
        strSQL = strSQL & "'" & P_HOSPLN_移行対象外 & "' AS NEW_HOSPLN_IKOU, "
    End If
    '▲2018/08/06 EGL INS<E>
    
    strSQL = strSQL & "        YOUKT_YUKBN AS YUKBN, "
    strSQL = strSQL & "        YOUKT_BUTU AS BUTU, "
    strSQL = strSQL & "        YOUKT_UCODE AS UCODE, "
    strSQL = strSQL & "        YOUKT_YCODE AS YCODE, "
    strSQL = strSQL & "        YOUKT_NO AS CNTANO, "
    
    '請求金額関係
    '▼2008/03/23 tajima chg&add
    '移動で満了更新or解約ヤードからの加瀬移動の場合は元受付から取得
    If blnMove Then


        '▼2008/06/10 tajima chg
        'str元受取条件sql = " YOUKT_MOVEKBN = " & P_MOVEKBN_満了更新 & _
        '                   " OR (YOUKT_MOVEKBN = " & P_MOVEKBN_自社理由 & _
        '                   "     AND ISNULL(MOTO_YARD.YARD_RENTEND_DAY,'2999/12/31') <= GETDATE() )"
        '↓
        '自社理由＆満了更新ならば（旧契約からelse新ヤードから）
        str元受取条件sql = " YOUKT_MOVEKBN in(" & P_MOVEKBN_満了更新 & "," & P_MOVEKBN_自社理由 & ")"
        '▲2008/06/10
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_RENTKG ELSE ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) END AS RENTKG, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_RENTKG ELSE ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) END AS FRENTKG, "
         '2019/08/08 M.HONDA INS
     '   strSQL = strSQL & "     CASE WHEN YOUKT_KKDATE between '" & tax_from & "' and '" & tax_to & "' THEN "
     '   strSQL = strSQL & "         ISNULL(NOW_PRIC.PRIC_FPRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) "
       ' strSQL = strSQL & "     Else "
    ' '   strSQL = strSQL & "         ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) "
       ' strSQL = strSQL & "     END AS FRENTKG ,"
     ''   '2019/08/08 M.HONDA INS
        strSQL = strSQL & " '' AS FIRSTKG, "
        '▼2018/08/30 EGL CHG
        'strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN 0 ELSE NOW_PRIC.PRIC_EZAPPI_CODE END AS EZAPPI_CODE, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_EZAPPI_CODE  ELSE NOW_PRIC.PRIC_EZAPPI_CODE  END AS EZAPPI_CODE, "
        '▲2018/08/30 EGL CHG
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_EZAPPI       ELSE NOW_PRIC.PRIC_EZAPPI       END AS EZAPPI, "
        
        '▼2018/08/30 EGL CHG
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_ADD_EZAPPI_CODE1 ELSE NOW_PRIC.PRIC_EZAPPI_CODE2 END AS EZAPPI_CODE2, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_ADD_EZAPPI1      ELSE NOW_PRIC.PRIC_EZAPPI2      END AS EZAPPI2, "
        '▲2018/08/30 EGL CHG
        
        strSQL = strSQL & " '' AS EZAPPI_DAILY, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN convert(varchar,MOTO_RCPT.RCPT_ADD_EZAPPI_CODE1) ELSE '' END AS ADD_EZAPPI_CODE1, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_ADD_EZAPPI1 ELSE 0 END AS ADD_EZAPPI1, "
        strSQL = strSQL & " '' AS ADD_EZAPPI_DAILY1, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN convert(varchar,MOTO_RCPT.RCPT_ADD_EZAPPI_CODE2) ELSE '' END AS ADD_EZAPPI_CODE2, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_ADD_EZAPPI2 ELSE 0 END AS ADD_EZAPPI2, "
        strSQL = strSQL & " '' AS ADD_EZAPPI_DAILY2, "
        
        '▼2008/06/10 tajima chg 初回雑費取得不具合対応
'        strSQL = strSQL & " '' AS FZAPPI, "
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN 0 ELSE NOW_PRIC.PRIC_FZAPPI       END AS FZAPPI, "
        '▲2008/06/10
        
        '▼2018/08/30 EGL CHG
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN 0 ELSE NOW_PRIC.PRIC_FZAPPI_CODE  END AS FZAPPI_CODE, "
        '▲2018/08/30 EGL CHG
        
'        strSQL = strSQL & " '' AS ADD_FZAPPI_CODE1, "                      'DELETE 2008/11/13 SHIBAZAKI
'        strSQL = strSQL & " '' AS ADD_FZAPPI1, "                           'DELETE 2008/11/13 SHIBAZAKI

        strSQL = strSQL & P_ZAPPI_事務手数料 & " AS ADD_FZAPPI_CODE1, "     'INSERT 2008/11/13 SHIBAZAKI
        
        ' 2015/07/24 M.HONDA
        ' 顧客理由移動時には、事務手数料3240円とする。
        ' 2013/06/04 M.HONDA START
        ' 保証会社変更以外は事務手数料は0円とする。
        'strSQL = strSQL & P_JIMUTESURYO & "      AS ADD_FZAPPI1, "          'INSERT 2008/11/13 SHIBAZAKI
        'strSQL = strSQL & " CASE WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_保障会社変更 & " THEN " & P_JIMUTESURYO & " ELSE 0 END AS ADD_FZAPPI1,"
        '2018/07/09 M.HONDA UPDA
        'オフィスの場合には21600円とする。
        If strBumonCode = "1" Then
           ' strSQL = strSQL & " CASE WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_保障会社変更 & " THEN " & P_JIMUTESURYO & "WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_顧客理由 & " THEN " & P_KMOVEJIMUTE & " ELSE 21600 END AS ADD_FZAPPI1,"
            strSQL = strSQL & " CASE WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_保障会社変更 & " THEN " & P_JIMUTESURYO & "WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_顧客理由 & " THEN " & P_KMOVEJIMUTE & " ELSE 22000 END AS ADD_FZAPPI1,"
            
        Else
            strSQL = strSQL & " CASE WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_保障会社変更 & " THEN " & P_JIMUTESURYO & "WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_顧客理由 & " THEN " & P_KMOVEJIMUTE & " ELSE 0 END AS ADD_FZAPPI1,"
        End If
        '2018/07/09 M.HONDA UPDA
        
        ' 2013/06/04 M.HONDA END
        'strSQL = strSQL & " '' AS ADD_FZAPPI_CODE2, "
        strSQL = strSQL & " 0 AS ADD_FZAPPI_CODE2, "
        strSQL = strSQL & " '' AS ADD_FZAPPI2, "
    
        '移動予約ならば前契約の価格情報取得を考慮
'        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " AND  MOTO_PRIC.PRIC_EZAPPI_CODE <> " & P_ZAPPI_無し & " THEN MOTO_EZAPPINAME.NAME_NAME " _
                        & " WHEN NOW_PRIC.PRIC_EZAPPI_CODE <> " & P_ZAPPI_無し & " THEN EZAPPINAME.NAME_NAME " _
                        & " ELSE '' END AS EZAPPI_NAME,"                        'DELETE 2013/04/10 K.ISHIZAKA
'        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " AND  MOTO_PRIC.PRIC_FZAPPI_CODE <> " & P_ZAPPI_無し & " THEN MOTO_FZAPPINAME.NAME_NAME " _
                        & " WHEN NOW_PRIC.PRIC_FZAPPI_CODE <> " & P_ZAPPI_無し & " THEN FZAPPINAME.NAME_NAME " _
                        & " ELSE '' END AS FZAPPI_NAME,"                        'DELETE 2013/04/10 K.ISHIZAKA
    
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " AND  MOTO_PRIC.PRIC_EZAPPI_CODE <> " & P_ZAPPI_無し & " THEN MOTO_EZAPPINAME.NAME_NAME " _
                        & " WHEN " & str元受取条件sql & " AND  MOTO_PRIC.PRIC_EZAPPI_CODE = " & P_ZAPPI_無し & " THEN '' " _
                        & " WHEN NOW_PRIC.PRIC_EZAPPI_CODE <> " & P_ZAPPI_無し & " THEN EZAPPINAME.NAME_NAME " _
                        & " ELSE '' END AS EZAPPI_NAME,"                        'INSERT 2013/04/10 K.ISHIZAKA
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " AND  MOTO_PRIC.PRIC_FZAPPI_CODE <> " & P_ZAPPI_無し & " THEN MOTO_FZAPPINAME.NAME_NAME " _
                        & " WHEN " & str元受取条件sql & " AND  MOTO_PRIC.PRIC_FZAPPI_CODE = " & P_ZAPPI_無し & " THEN '' " _
                        & " WHEN NOW_PRIC.PRIC_FZAPPI_CODE <> " & P_ZAPPI_無し & " THEN FZAPPINAME.NAME_NAME " _
                        & " ELSE '' END AS FZAPPI_NAME,"                        'INSERT 2013/04/10 K.ISHIZAKA
                               

        '▼2008/06/10 tajima chg 保証金初期値再整理
        '保証金 保証会社変更ならば新金額、満了更新及び解約ヤードの為の自社理由移動でリプラスならば０、それ以外なら前回の金額
        'strSQL = strSQL & " CASE WHEN YOUKT_MOVEKBN = " & P_MOVEKBN_保障会社変更 & " THEN ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) " _
        '                & " ELSE CASE WHEN MOTO_CARG.CARG_HOSYICD = '" & P_HOSYICD_リプラス & "' AND (" _
        '                & str元受取条件sql & ") THEN 0 ELSE MOTO_CARG.CARG_SECUKG END " _
        '                & " END AS SECUKG,"
        ' ↓
        '満了更新（旧契約がリプラス）ならば０
        '自社理由＆満了更新(旧契約がリプラス以外）ならば旧契約の金額
        'それ以外は新ヤードの月額（使用料＋雑費計)
        'DELETE 2008/11/13 SHIBAZAKI
'        strSQL = strSQL & " CASE WHEN MOTO_CARG.CARG_HOSYICD = '" & P_HOSYICD_リプラス & "' AND YOUKT_MOVEKBN = " & P_MOVEKBN_満了更新 & " THEN 0 " _
'                        & " ELSE CASE WHEN " & str元受取条件sql & " THEN MOTO_CARG.CARG_SECUKG " _
'                        & " ELSE ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) + ISNULL(NOW_PRIC.PRIC_EZAPPI,0) END " _
'                        & " END AS SECUKG,"
        'INSERT 2008/11/13 SHIBAZAKI
        '満了更新（旧契約が【保証委託】）ならば０
        '自社理由＆満了更新(旧契約が【保証委託以外】）ならば旧契約の金額
        'それ以外は新ヤードの月額（使用料＋雑費計)
        '▼2018/08/06 EGL UPD<S>
        'strSQL = strSQL & " CASE WHEN HOSYORYAKU.NAME_VALUE_FROM = '" & P_HOSYIFROM_保証委託 & "' AND YOUKT_MOVEKBN = " & P_MOVEKBN_満了更新 & " THEN 0 " _
        '                & " ELSE CASE WHEN " & str元受取条件sql & " THEN MOTO_CARG.CARG_SECUKG " _
        '                & " ELSE ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) + ISNULL(NOW_PRIC.PRIC_EZAPPI,0) END " _
        '                & " END AS SECUKG,"
        'strSql = strSql & " CASE WHEN HOSYORYAKU.NAME_VALUE_FROM = '" & P_HOSYIFROM_保証委託 & "' AND YOUKT_MOVEKBN = " & P_MOVEKBN_満了更新 & " THEN 0 " _
                         & " WHEN " & str契約移動条件sql & " THEN MOTO_RCPT.RCPT_SECUKG " _
                         & " ELSE CASE WHEN " & str元受取条件sql & " THEN MOTO_CARG.CARG_SECUKG " _
                         & " ELSE ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) + ISNULL(NOW_PRIC.PRIC_EZAPPI,0) END " _
                         & " END AS SECUKG,"
        '▲2018/08/06 EGL UPD<E>
        '▲2008/06/10
        '20230116 M.HODNA UPD
        strSQL = strSQL & " 0 AS SECUKG,"
        '▲2018/08/06 EGL UPD<E>
        '▲2008/06/10
        
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN MOTO_RCPT.RCPT_EZAPPI ELSE NOW_PRIC.PRIC_EZAPPI END AS FEZAPPI, "
        
        '▼2020/09/28 tajima add --自社理由と満了更新ならば移動元を、それ以外ならば受付日を元に適用するプランコードを取得する
        strSQL = strSQL & " CASE WHEN " & str元受取条件sql & " THEN ISNULL(MOTO_RCPT.RCPT_PLAN_CD,'0000') " _
                        & " ELSE ISNULL((select top 1 CTRLT_VALUE01 from CTRL_TABL where CTRLT_CODE like 'PLAN_PTN_%' and FORMAT(YOUKT_KUDATE,'yyyyMMdd') between CTRLT_KEY01 and CTRLT_KEY02),'0000') " _
                        & " END AS RCPT_PLAN_CD, "
        '▲2020/09/28 tajima add
        
    Else '移動でなく新規ならば
        strSQL = strSQL & " ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) AS RENTKG, "
        '2019/08/08 M.HONDA INS
        strSQL = strSQL & "     CASE WHEN YOUKT_KKDATE <= '" & tax_to & "' THEN "
        strSQL = strSQL & "         NOW_PRIC.PRIC_FEZAPPI "
        strSQL = strSQL & "     Else "
        strSQL = strSQL & "         NOW_PRIC.PRIC_EZAPPI "
        strSQL = strSQL & "     END AS FEZAPPI ,"
        '2019/08/08 M.HONDA INS
        
         '2019/08/08 M.HONDA INS
        'strSQL = strSQL & "     CASE WHEN YOUKT_KKDATE between '" & tax_from & "' and '" & tax_to & "' THEN "
        strSQL = strSQL & "     CASE WHEN YOUKT_KKDATE <= '" & tax_to & "' THEN "
        strSQL = strSQL & "         ISNULL(NOW_PRIC.PRIC_FPRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) "
        strSQL = strSQL & "     Else "
        strSQL = strSQL & "         ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) "
        strSQL = strSQL & "     END AS FRENTKG ,"
        '2019/08/08 M.HONDA INS
        strSQL = strSQL & " '' AS FIRSTKG, "
        
        strSQL = strSQL & " NOW_PRIC.PRIC_EZAPPI_CODE AS EZAPPI_CODE, "         '' 2012/12/27 M.HONDA INS
        strSQL = strSQL & " NOW_PRIC.PRIC_EZAPPI AS EZAPPI, "
       
        '▼2018/08/30 EGL CHG
        strSQL = strSQL & " NOW_PRIC.PRIC_FZAPPI_CODE AS FZAPPI_CODE, "
        strSQL = strSQL & " NOW_PRIC.PRIC_FZAPPI AS FZAPPI, "
        '▲2018/08/30 EGL CHG
        
        strSQL = strSQL & " '' AS EZAPPI_DAILY, "
        strSQL = strSQL & " '' AS ADD_EZAPPI_CODE1, "
        strSQL = strSQL & " '0' AS ADD_EZAPPI1, "
        strSQL = strSQL & " '' AS ADD_EZAPPI_DAILY1, "
        strSQL = strSQL & " '' AS ADD_EZAPPI_CODE2, "
        strSQL = strSQL & " '0' AS ADD_EZAPPI2, "
        strSQL = strSQL & " '' AS ADD_EZAPPI_DAILY2, "
        strSQL = strSQL & " NOW_PRIC.PRIC_FZAPPI AS FZAPPI, "
'        strSQL = strSQL & " '' AS ADD_FZAPPI_CODE1, "                          'DELETE 2008/11/13 SHIBAZAKI
'        strSQL = strSQL & " '' AS ADD_FZAPPI1, "                               'DELETE 2008/11/13 SHIBAZAKI
        '' 2011/07/21 M.HONDA START
        ''strSQL = strSQL & P_ZAPPI_事務手数料 & " AS ADD_FZAPPI_CODE1, "         'INSERT 2008/11/13 SHIBAZAKI
        ''strSQL = strSQL & P_JIMUTESURYO & "      AS ADD_FZAPPI1, "              'INSERT 2008/11/13 SHIBAZAKI
        '----↓↓↓↓----2018/10/02 CHG EGL.T-----↓↓↓↓---<s>
'        strSQL = strSQL & P_ZAPPI_事務手数料 & " AS ADD_FZAPPI_CODE1, "
        '----↓↓↓↓----20110829--M.HONDA--Ins-------↓↓↓↓---<s>
'        Dim Add_FZappi1 As Integer
'        Add_FZappi1 = getFzappi1
'        strSQL = strSQL & Add_FZappi1 & "  AS ADD_FZAPPI1, "
        '----↑↑↑↑----20110829--M.HONDA--Ins-------↑↑↑↑---<e>
        
'        strSQL = strSQL & " CASE WHEN YOUKT_GENKBN = 1 THEN "
'        '' 2013/11/13 M.HONDA UPD
'        '' strSQL = strSQL & " '' "
'        strSQL = strSQL & P_JIMUTESURYO
'        '' 2013/11/13 M.HONDA UPD
'        strSQL = strSQL & " ELSE " & P_JIMUTESURYO & " END    AS ADD_FZAPPI1, "
        '' 2011/07/21 M.HONDA END
        '以下が追加初回雑費１を価格表の初回雑費２から取得
        strSQL = strSQL & "NOW_PRIC.PRIC_FZAPPI_CODE2 AS ADD_FZAPPI_CODE1," 'コード
        strSQL = strSQL & " NOW_PRIC.PRIC_FZAPPI2 AS ADD_FZAPPI1, " '雑費金額
        '----↑↑↑↑----2018/10/02 CHG EGL.T-----↑↑↑↑---<e>
        
        'strSQL = strSQL & " '' AS ADD_FZAPPI_CODE2, "
        strSQL = strSQL & " 0 AS ADD_FZAPPI_CODE2, "
        strSQL = strSQL & " '' AS ADD_FZAPPI2, "
        strSQL = strSQL & " NOW_PRIC.PRIC_PRICE, "
        strSQL = strSQL & " NOW_PRIC.PRIC_FPRICE,  " '2019/08/08 M.HONDA INS
        
        strSQL = strSQL & " CASE WHEN NOW_PRIC.PRIC_EZAPPI_CODE = " & P_ZAPPI_無し
        strSQL = strSQL & "  THEN '' "
        strSQL = strSQL & "  ELSE EZAPPINAME.NAME_NAME "
        strSQL = strSQL & " END AS EZAPPI_NAME, "
        strSQL = strSQL & " CASE WHEN NOW_PRIC.PRIC_FZAPPI_CODE = " & P_ZAPPI_無し
        strSQL = strSQL & "  THEN '' "
        strSQL = strSQL & "  ELSE FZAPPINAME.NAME_NAME "
        strSQL = strSQL & " END AS FZAPPI_NAME, "
        
        '▼ 2018/10/02 EGL tajima chg
        '▼ 2008/06/10 tajima add
        '新規の場合は月額総金額
'       strSQL = strSQL & " ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0) + ISNULL(NOW_PRIC.PRIC_EZAPPI,0) AS SECUKG, "
'        strSQL = strSQL & fncMakeCargCase(blnMove, _
'                                      Array(P_MOVEKBN_保障会社変更), _
'                                      "ISNULL(NOW_PRIC.PRIC_PRICE,0) + ISNULL(NOW_CNTA.CNTA_PRICE_DIFF,0)", _
'                                      "MOTO_CARG.CARG_SECUKG", _
'                                      "SECUKG") & ","
        '▲ 2008/06/10 tajima add
        '新規の場合の保証金額は0
        strSQL = strSQL & " 0 AS SECUKG, "
        '▲ 2018/10/02 EGL tajima chg
    
        '▼2020/09/28 tajima add --受付日を元に適用するプランコードを取得する
         strSQL = strSQL & "ISNULL((select top 1 CTRLT_VALUE01 from CTRL_TABL where CTRLT_CODE like 'PLAN_PTN_%' and FORMAT(YOUKT_KUDATE,'yyyyMMdd') between CTRLT_KEY01 and CTRLT_KEY02),'0000') AS RCPT_PLAN_CD, "
        '▲2020/09/28 tajima add
    End If
    '▲ 2008/03/23 tajima chg&add
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
    strSQL = strSQL & "        0 AS SECUKG_KASE, "                              'INSERT 2013/04/10 K.ISHIZAKA
    
    '保証金
    '▼2008/12/04 hirano Add 保証金変更フラグ取得
    '保証金変更フラグ
    strSQL = strSQL & "        0 AS SECUKG_FLG, "
    '▲2008/12/04 hirano Add
    
    strSQL = strSQL & "        '' AS SEIKYUKIKAN, "
        
    If blnMove Then
        strSQL = strSQL & "        MOTO_CARG.CARG_FRST_BILL  AS SUMI_RENTKG, "
    Else
        strSQL = strSQL & "        0  AS SUMI_RENTKG, "
    End If
    
    strSQL = strSQL & "        YOUKT_KUDATE AS UKDATE, "
    strSQL = strSQL & "        YOUKT_UKTANTO AS UKTANTO, "
    
    '顧客コードが格納されていない場合、予約トランの値もしくは初期値を採用(第１パラメータ)
    '顧客コードが格納されている場合、顧客マスタの値を採用(第２パラメータ)
    '第３パラメータは項目の別名
    strSQL = strSQL & fncMakeUserCase("YOUKT_NAME", "YOUKTUSER.USER_NAME", "UNAME") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_KANA", "YOUKTUSER.USER_KANA", "UKANA") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_TANM", "YOUKTUSER.USER_TANM", "UTANM") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_TAKA", "YOUKTUSER.USER_TAKA", "UTAKA") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_YUBINO", "YOUKTUSER.USER_YUBINO", "UYUBINO") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_ADR_1", "YOUKTUSER.USER_ADR_1", "UADR1") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_ADR_2", "YOUKTUSER.USER_ADR_2", "UADR2") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_ADR_3", "YOUKTUSER.USER_ADR_3", "UADR3") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_TEL", "YOUKTUSER.USER_TEL", "UTEL") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_KEITAI", "YOUKTUSER.USER_KEITAI", "UKEITAI") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_FAX", "YOUKTUSER.USER_FAX", "UFAX") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_KKBN", "YOUKTUSER.USER_KKBN", "USERKKBN") & ","
    strSQL = strSQL & fncMakeUserCase("YOUKT_MAIL", "YOUKTUSER.USER_MAIL", "UMAIL") & ","
    strSQL = strSQL & fncMakeUserCase("31", "YOUKTUSER.USER_SDAY", "USDAY") & ","
    '' 2013/0919 M.HONDA INS
    'strSQL = strSQL & fncMakeUserCase(P_USER_SKBN_コンビニ, "YOUKTUSER.USER_SKBN", "USKBN") & ","
    strSQL = strSQL & fncMakeUserCase(P_USER_SKBN_振込, "YOUKTUSER.USER_SKBN", "USKBN") & ","     '' 2013/09/19 M.HONDA INS
    '' 2013/0919 M.HONDA INS
    
    
    strSQL = strSQL & "        YOUKTUSER.USER_BIKO AS UBIKO, "
    
    '移動予約の場合は元契約から情報を取得する
    '保証区分
    '↓INSERT 2008/11/13 SHIBAZAKI
    strSQL = strSQL & fncMakeCargCase(blnMove, _
                                      Array(P_MOVEKBN_保障会社変更), _
                                      "'" & P_HOSYICD_パルマ & "'", _
                                      "CASE WHEN MOTO_CARG.CARG_HOSYICD = '2' THEN '1' ELSE MOTO_CARG.CARG_HOSYICD END ", _
                                      "HOSYICD") & ","
    '↑INSERT 2008/11/13 SHIBAZAKI
    '↓INSERT 2008/09/25 SHIBAZAKI-> DELETE 2008/11/13 SHIBAZAKI
'    strSQL = strSQL & fncMakeCargCase(blnMove, _
'                                      Array(P_MOVEKBN_保障会社変更), _
'                                      "'" & P_HOSYICD_臨時 & "'", _
'                                      "MOTO_CARG.CARG_HOSYICD", _
'                                      "HOSYICD") & ","
    '↑INSERT 2008/09/25 SHIBAZAKI-> DELETE 2008/11/13 SHIBAZAKI
    '↓DELETE 2008/09/25 SHIBAZAKI
'    strSQL = strSQL & fncMakeCargCase(blnMove, _
'                                      Array(P_MOVEKBN_保障会社変更), _
'                                      "'" & P_HOSYICD_リプラス & "'", _
'                                      "MOTO_CARG.CARG_HOSYICD", _
'                                      "HOSYICD") & ","
    '↑DELETE 2008/09/25 SHIBAZAKI
   '▼2018/08/06 EGL UPD<S>
    '2009/04/01 ADD <S> hirano 保障会社コード
    'strSQL = strSQL & "        '' AS HOSYO_KAISHA_CD, "
    '2009/04/01 ADD <E>
    If blnMove Then
        strSQL = strSQL & " CASE WHEN " & str契約移動条件sql & " THEN MOTO_RCPT.RCPT_HOSYO_CD ELSE '' END AS HOSYO_KAISHA_CD, "
    Else
        strSQL = strSQL & " '' AS HOSYO_KAISHA_CD, "
    End If
    '▲2018/08/06 EGL UPD<E>
    
    If blnMove Then
        '2006/06/01 SHIBAZAKI 鍵情報の初期値はコンテナマスタから取得する
'        strSql = strSql & "        MOTO_CARG.CARG_KAGIICD AS KAGIICD, "
'        strSql = strSql & "        MOTO_CARG.CARG_KYDATE AS CYDATE, "  'DELETE 2019/03/14 Y.WADA
        strSQL = strSQL & "        '' AS CYDATE, "                      'INSERT 2019/03/14 Y.WADA
        strSQL = strSQL & "        '99' AS UKKBN, "
    Else
        '2006/06/01 SHIBAZAKI 鍵情報の初期値はコンテナマスタから取得する
'        strSql = strSql & "        '' AS KAGIICD, "
        strSQL = strSQL & "        '' AS CYDATE, "
        strSQL = strSQL & "        YOUKT_UKKBN AS UKKBN, "
    End If
    
'    '2006/04/28 S.SHIBAZAKI 乗用具チェック追加。チェックされている場合、申込書に誓約書添付
'    strSql = strSql & "        0 AS VEHICLE, "
    
    '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
    strSQL = strSQL & "        '' AS CARGACPTNO, "
    
    strSQL = strSQL & "        '' AS CARGACPTNO, "
    
    '2006/06/01 SHIBAZAKI 鍵情報の初期値はコンテナマスタから取得する
    strSQL = strSQL & "        NOW_CNTA.CNTA_KEY_NO AS KAGINO, "
    strSQL = strSQL & "        NOW_CNTA.CNTA_KEY_NUM AS KAGILENDNUM, "
    strSQL = strSQL & "        NOW_CNTA.CNTA_KEY_TYPE AS KAGIICD, "
    
'↓ INSERT 2014/03/11 MIYAMOTO
    strSQL = strSQL & "        ( "
    strSQL = strSQL & "            SELECT      AKKNT.AKKNT_KAINO "
    strSQL = strSQL & "            FROM        ( "
    strSQL = strSQL & "                            SELECT      AKKNT_YARD "
    strSQL = strSQL & "                            ,           AKKNT_CTNO "
    strSQL = strSQL & "                            ,           AKKNT_KTYPE "
    strSQL = strSQL & "                            ,           AKKNT_HATUD "
'    strSql = strSql & "                            ,           ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE ORDER BY AKKNT_HATUD DESC) AS SEQ "'UPDATE 2021/06/02 EGL
    strSQL = strSQL & "                            ,           ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE, AKKNT_STYPE ORDER BY AKKNT_HATUD DESC) AS SEQ " 'UPDATE 2021/06/02 EGL
    strSQL = strSQL & "                            ,           AKKNT_KAINO "
    strSQL = strSQL & "                            FROM        AKKN_TRAN "
    strSQL = strSQL & "                            WHERE       AKKNT_YARD = NOW_CNTA.CNTA_CODE "
    strSQL = strSQL & "                            AND         AKKNT_CTNO = NOW_CNTA.CNTA_NO "
    strSQL = strSQL & "                            AND         AKKNT_KTYPE = NOW_CNTA.CNTA_KEY_TYPE "
    strSQL = strSQL & "                            AND         AKKNT_STYPE = '4' " ' INSERT 2021/06/02 EGL
    strSQL = strSQL & "                        ) AKKNT "
    strSQL = strSQL & "            WHERE       AKKNT.SEQ = 1 "
    strSQL = strSQL & "        ) AS KAIJONO, "
'↑ INSERT 2014/03/11 MIYAMOTO

'↓ INSERT 2021/06/02 EGL
    strSQL = strSQL & "        ( "
    strSQL = strSQL & "            SELECT      AKKNT.AKKNT_KAINO "
    strSQL = strSQL & "            FROM        ( "
    strSQL = strSQL & "                            SELECT      AKKNT_YARD "
    strSQL = strSQL & "                            ,           AKKNT_CTNO "
    strSQL = strSQL & "                            ,           AKKNT_KTYPE "
    strSQL = strSQL & "                            ,           AKKNT_HATUD "
    strSQL = strSQL & "                            ,           ROW_NUMBER() OVER(PARTITION BY AKKNT_YARD, AKKNT_CTNO, AKKNT_KTYPE, AKKNT_STYPE ORDER BY AKKNT_HATUD DESC) AS SEQ "
    strSQL = strSQL & "                            ,           AKKNT_KAINO "
    strSQL = strSQL & "                            FROM        AKKN_TRAN "
    strSQL = strSQL & "                            WHERE       AKKNT_YARD = NOW_CNTA.CNTA_CODE "
    strSQL = strSQL & "                            AND         AKKNT_CTNO = NOW_CNTA.CNTA_NO "
    strSQL = strSQL & "                            AND         AKKNT_KTYPE = NOW_CNTA.CNTA_KEY_TYPE "
    strSQL = strSQL & "                            AND         AKKNT_STYPE = '6' "
    strSQL = strSQL & "                        ) AKKNT "
    strSQL = strSQL & "            WHERE       AKKNT.SEQ = 1 "
    strSQL = strSQL & "        ) AS RCPT_BUIL_KAGI_NO, "
'↑ INSERT 2021/06/02 EGL
    
    '2006/09/06 SHIBAZAKI 契約開始日は予約トラン・希望開始日を初期値とする
    strSQL = strSQL & "        YOUKT_KKDATE AS STDATE, "
    
    strSQL = strSQL & "        '' AS DC_STDATE, "                               'I:20081001 N.Imai
    strSQL = strSQL & "        0  AS DC_ENABLE, "                               'I:20081001 N.Imai
    
    'ここより下は初期値としてはNULL
    strSQL = strSQL & "        '' AS SEV_EXPIRE, "
    '2009/04/01 UPD <S> hirano 申込書送付方を初期セットなければ、FAX初期セット
    'strSQL = strSQL & "        '' AS SENDCD, "
    strSQL = strSQL & "        ISNULL(YOUKT_SEND_CD,1) AS SENDCD, "
    '2009/04/01 UPD <E> hirano
    '2006/09/06 SHIBAZAKI 契約開始日は予約トラン・希望開始日を初期値とする
'    strSql = strSql & "        '' AS STDATE, "
    '2009/04/01 UPD <S> hirano　見積もり請求区分　明細書デフォルト
    'strSQL = strSQL & "        '' AS MITUCD, "
    strSQL = strSQL & "        '01' AS MITUCD, "
    '2009/04/01 UPD <E> hirano
    '2009/04/01 UPD <S> hirano　初期費用請求方法　代引きデフォルト
    'strSQL = strSQL & "        '' AS SEIKYUCD, "
    strSQL = strSQL & "        '02' AS SEIKYUCD, "
    '2009/04/01 UPD <E> hirano
    '2006/06/01 SHIBAZAKI 鍵情報の初期値はコンテナマスタから取得する
'    strSql = strSql & "        '' AS KAGINO, "
    '2006/06/01 SHIBAZAKI セキュリティカード対応
'    strSQL = strSQL & "        '' AS SECURITYCARD, "
    '2006/06/01 SHIBAZAKI 鍵情報の初期値はコンテナマスタから取得する
'    strSql = strSql & "        '' AS KAGILENDNUM, "
    
    '↓↓↓↓------20110401---------M.HONDA----------ADD------------↓↓↓↓-------------<s>
    'strSQL = strSQL & "        '' AS HOSYD, "
    'strSQL = strSQL & "        '' AS HOSYB, "
    If blnMove Then
        strSQL = strSQL & "         CASE"
'        strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5)  THEN"      'DELETE 2011/09/27 M.RYU
        'strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5,7)  THEN"     'INSERT 2011/09/27 M.RYU
        'strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5,7,2)  THEN"     'INSERT 2015/08/03 M.HONDA
        strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5,7,2,3)  THEN"     'INSERT 2018/07/17 M.HONDA
        
        strSQL = strSQL & "             MOTO_CARG.CARG_HOSYD "
        strSQL = strSQL & "         END AS HOSYD, "
        strSQL = strSQL & "         CASE"
'        strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5)  THEN"      'DELETE 2011/09/27 M.RYU
        'strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5,7)  THEN"     'INSERT 2011/09/27 M.RYU
        'strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5,7,2)  THEN"    'INSERT 2015/08/03 M.RYU
        strSQL = strSQL & "         WHEN YOUKT_MOVEKBN IN (10,1,5,7,2,3)  THEN"    'INSERT 2018/07/17 M.HONDA
        strSQL = strSQL & "             MOTO_CARG.CARG_HOSYB "
        strSQL = strSQL & "         ELSE "
        strSQL = strSQL & "             ''  "
        strSQL = strSQL & "         END AS HOSYB,"
    Else
        strSQL = strSQL & "         '' HOSYD, "
        strSQL = strSQL & "         '' HOSYB,"
    End If
    '↑↑↑↑------20110401---------M.HONDA----------ADD------------↑↑↑↑-------------<e>
    
    '▼2008/03/23 tajima chg&add
    '移動の場合は前回の受付情報があればそこから取得
    If blnMove Then
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEND_NAME AS SENDNAME, "
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEND_HONORIFIC AS SENDHONORIFIC, "
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEND_YUBINO AS SENDYUBINO, "
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEND_ADR_1 AS SENDADR1,"
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEND_ADR_2 AS SENDADR2,"
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEND_ADR_3 AS SENDADR3,"
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEND_TEL AS SENDTEL, "
    
'        strSQL = strSQL & "        MOTO_RCPT.RCPT_SEIKI AS SEIKI,"             'DELETE 2010/04/10 K.ISHIZAKA 'INSERT 2010/03/29 K.ISHIZAKA
'       strSQL = strSQL & "        MOTO_CARG.CARG_SEIKI AS SEIKI,"              'INSERT 2010/04/10 K.ISHIZAKA
        strSQL = strSQL & "        2 AS SEIKI,"
        
        strSQL = strSQL & "        MOTO_RCPT.RCPT_KAIHI AS KAIHI,"              'INS 2012/10/03 M.HONDA
        strSQL = strSQL & "        MOTO_RCPT.RCPT_TPOINT AS RCPT_TPOINT, "      ''2013/05/21 M.HONDA
        strSQL = strSQL & "        NULL AS RCPT_KOUKBN, "                       ''2013/10/16 M.HONDA
        
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SHOUMEI_KDATE,"         'INSERT 2015/07/16 K.ISHIZAKA
        strSQL = strSQL & "        MOTO_RCPT.RCPT_NEN_KBN, "              'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "        MOTO_RCPT.RCPT_NEN_JIREI, "            'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "        MOTO_RCPT.RCPT_NEN_SEIKI, "            'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "        MOTO_RCPT.RCPT_NEN_BIKO, "             'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "        MOTO_RCPT.RCPT_SYOKISAN_DATE, "        '2018/03/16 M.HONDA INS
        
        
    Else
        strSQL = strSQL & "        '' AS SENDNAME, "
        strSQL = strSQL & "        '' AS SENDHONORIFIC, "
        strSQL = strSQL & "        '' AS SENDYUBINO, "
        strSQL = strSQL & "        '' AS SENDADR1,"
        strSQL = strSQL & "        '' AS SENDADR2,"
        strSQL = strSQL & "        '' AS SENDADR3,"
        strSQL = strSQL & "        '' AS SENDTEL, "
        
        strSQL = strSQL & "        ISNULL("                                     'INSERT START 2010/03/29 K.ISHIZAKA
        strSQL = strSQL & "("
        strSQL = strSQL & " SELECT TOP 1"
'        strSQL = strSQL & "  RCPT_SEIKI"                                       'DELETE 2010/04/10 K.ISHIZAKA
        strSQL = strSQL & "  CARG_SEIKI"                                        'INSERT 2010/04/10 K.ISHIZAKA
        strSQL = strSQL & " FROM"
'        strSQL = strSQL & "  RCPT_TRAN"                                        'DELETE 2010/04/10 K.ISHIZAKA
'        strSQL = strSQL & " INNER JOIN"                                        'DELETE 2010/04/10 K.ISHIZAKA
        strSQL = strSQL & "  CARG_FILE"
'        strSQL = strSQL & " ON"                                                'DELETE 2010/04/10 K.ISHIZAKA
'        strSQL = strSQL & "  CARG_ACPTNO = RCPT_CARG_ACPTNO"                   'DELETE 2010/04/10 K.ISHIZAKA
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & "  CARG_UCODE = YOUKT_UCODE"
        strSQL = strSQL & " AND"
        strSQL = strSQL & "  GETDATE() BETWEEN CARG_STDATE AND ISNULL(CARG_KYDATE,'9999-12-31')"
        strSQL = strSQL & ")        "
        strSQL = strSQL & "        ," & P_SEIKI_未発行 & ") AS SEIKI,"          'INSERT END   2010/03/29 K.ISHIZAKA
        'strSQL = strSQL & "        ," & P_SEIKI_発行 & ") AS SEIKI,"             '2015/10/28 M.HONDA INS
        strSQL = strSQL & "      NULL AS KAIHI,"                                'INS 2012/10/03 M.HONDA
        strSQL = strSQL & "      NULL AS RCPT_TPOINT, "                         '2013/05/21 M.HONDA
        strSQL = strSQL & "      NULL AS RCPT_KOUKBN, "                         '2013/10/16 M.HONDA
        strSQL = strSQL & "      NULL AS RCPT_SHOUMEI_KDATE,"         'INSERT 2015/07/16 K.ISHIZAKA
        strSQL = strSQL & "      NULL AS RCPT_NEN_KBN, "              'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "      NULL AS RCPT_NEN_JIREI, "            'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "      NULL AS RCPT_NEN_SEIKI, "            'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "      NULL AS RCPT_NEN_BIKO, "             'INSERT 2015/03/24 M.HONDA
        strSQL = strSQL & "      ''   AS RCPT_SYOKISAN_DATE, "        '2018/03/16 M.HONDA INS
    End If
    '▲2008/03/23 tajima chg&add
    
    strSQL = strSQL & "        '' AS KEY_SEND_DATE, "
    strSQL = strSQL & "        '' AS ARRIVALEXPECT, "
    strSQL = strSQL & "        '' AS SPECIFYTIME, "
    strSQL = strSQL & "        '' AS SLIPNO, "
    strSQL = strSQL & "        '' AS DERIVERYSTATUS, "
    strSQL = strSQL & "        '' AS ARRIVALDATE, "
    strSQL = strSQL & "        '' AS RCPT_YAMATO_DATE, "              'INSERT 2019/03/22 TFC_Kurosu
    strSQL = strSQL & "        '' AS CREASON, "
    strSQL = strSQL & "        '' AS PAYDATE, "
    strSQL = strSQL & "        '' AS UKNOTE, "
    'strSQL = strSQL & "      NULL AS KAIHI,"                                    'INSERT 2010/03/29 K.ISHIZAKA　DEL 2012/10/03 M.HONDA
    
    '↓INSERT 2010/04/20 SHIBAZAKI
    strSQL = strSQL & "        NULL AS RCPT_KREN_NAME, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_KANA, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_YUBINO, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_ADR_1, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_ADR_2, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_ADR_3, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_TEL, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_KEITAI, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_FAX, "
    strSQL = strSQL & "        NULL AS RCPT_KREN_ZOKUGARA, "
    '↑INSERT 2010/04/20 SHIBAZAKI
    strSQL = strSQL & "        NULL AS RCPT_SLOPE, "   '' 2012/04/12 M.HONDA
    strSQL = strSQL & "        '' AS COMMENT1, "
    strSQL = strSQL & "        '' AS COMMENT2, "
    strSQL = strSQL & "        '' AS ADD_FZAPPI2_NAME "
    
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_PLAN_CD "                 'add 2020/09/28 プランコード
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_SHOUMEI_ADRSAI "          'add 2020/11/16 証明書住所差異

    fncMakeYoukSelect = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeUserCase
'        機能             :契約トラン条件分岐文作成
'                         : strThenStatement 顧客コードがNULLのステートメント
'                         : strElseStatement 顧客コードがNOT NULLのステートメント
'                         : strAlias カラムエイリアス
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeUserCase(strThenStatement As String, _
                                 strElseStatement As String, _
                                 strAlias As String) As String

    Dim strSQL      As String
    
    strSQL = strSQL & "        CASE WHEN YOUKTUSER.USER_CODE IS NULL "
    strSQL = strSQL & "            THEN " & strThenStatement
    strSQL = strSQL & "            ELSE " & strElseStatement
    strSQL = strSQL & "        END AS " & strAlias

    fncMakeUserCase = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeCargCase
'        機能             :契約トラン条件分岐文作成
'        IN               : blnCarg True=条件分岐あり False=なし(THENステートメントだけでSQL作成)
'                         : varCondition 移動区分の条件値配列
'                         : strThenStatement 条件真のステートメント
'                         : strElseStatement 条件偽のステートメント
'                         : strAlias カラムエイリアス
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeCargCase(blnCarg As Boolean, _
                                 varCondition As Variant, _
                                 strThenStatement As String, _
                                 strElseStatement As String, _
                                 strAlias As String) As String

    Dim strSQL      As String
    Dim varTemp     As Variant
    Dim strTemp     As String
    
    fncMakeCargCase = ""
    
    If blnCarg Then
        If UBound(varCondition) < 0 Then
            Exit Function
        End If
        
        '指定された値の分、移動区分との一致条件をORでつなげる
        strSQL = strSQL & " CASE WHEN "
        
        strTemp = ""
        For Each varTemp In varCondition
            strTemp = strTemp & IIf(strTemp <> "", " OR ", "") & _
                      " YOUKT_MOVEKBN = " & varTemp
        Next
        strSQL = strSQL & strTemp & " THEN "
    End If
    
    'THENステートメント
    strSQL = strSQL & strThenStatement
    
    If blnCarg Then
        'ELSEステートメント
        strSQL = strSQL & " ELSE "
        strSQL = strSQL & strElseStatement
        strSQL = strSQL & " END "
    End If
    
    strSQL = strSQL & " AS " & strAlias

    fncMakeCargCase = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeRcptSelect
'        機能             :データ検索SQL作成(既存受付SELECT句)
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeRcptSelect(blnMove As Boolean) As String

    Dim strSQL      As String
    
    strSQL = strSQL & "     CASE WHEN YOUKT_KKDATE <= '" & tax_to & "' THEN "
    strSQL = strSQL & "         NOW_PRIC.PRIC_FEZAPPI "
    strSQL = strSQL & "     Else "
    strSQL = strSQL & "         NOW_PRIC.PRIC_EZAPPI "
    strSQL = strSQL & "     END AS FEZAPPI ,"

'▼2018/8/30 EGL

' 2008/03/23 tajima add　fncMakeCommonSelect から移動、現物件のを取得
'    strSQL = strSQL & "        NOW_PRIC.PRIC_EZAPPI_CODE AS EZAPPI_CODE, "     '' INS 2012/12/27 M.HDONDA
'    strSQL = strSQL & "        CASE WHEN NOW_PRIC.PRIC_EZAPPI_CODE = " & P_ZAPPI_無し
'    strSQL = strSQL & "            THEN '' "
'    strSQL = strSQL & "            ELSE EZAPPINAME.NAME_NAME "
'    strSQL = strSQL & "        END AS EZAPPI_NAME, "
'    strSQL = strSQL & "        NOW_PRIC.PRIC_FZAPPI_CODE AS FZAPPI_CODE, "     '' INS 2012/12/27 M.HDONDA
'    strSQL = strSQL & "        CASE WHEN NOW_PRIC.PRIC_FZAPPI_CODE = " & P_ZAPPI_無し
'    strSQL = strSQL & "            THEN '' "
'    strSQL = strSQL & "            ELSE FZAPPINAME.NAME_NAME "
'    strSQL = strSQL & "        END AS FZAPPI_NAME, "
' 2008/03/23 tajima add
    
    '1
    strSQL = strSQL & "        RCPT_TRAN.RCPT_EZAPPI_CODE AS EZAPPI_CODE, "
    strSQL = strSQL & "        CASE WHEN RCPT_TRAN.RCPT_EZAPPI_CODE = " & P_ZAPPI_無し
    strSQL = strSQL & "            THEN '' "
    strSQL = strSQL & "            ELSE EZAPPINAME.NAME_NAME "
    strSQL = strSQL & "        END AS EZAPPI_NAME, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_FZAPPI_CODE AS FZAPPI_CODE, "
    strSQL = strSQL & "        CASE WHEN RCPT_TRAN.RCPT_FZAPPI_CODE = " & P_ZAPPI_無し
    strSQL = strSQL & "            THEN '' "
    strSQL = strSQL & "            ELSE FZAPPINAME.NAME_NAME "
    strSQL = strSQL & "        END AS FZAPPI_NAME, "
  
    '2
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_EZAPPI_CODE1 AS EZAPPI_CODE2, "
    strSQL = strSQL & "        CASE WHEN RCPT_TRAN.RCPT_ADD_EZAPPI_CODE1 = " & P_ZAPPI_無し
    strSQL = strSQL & "            THEN '' "
    strSQL = strSQL & "            ELSE EZAPPINAME2.NAME_NAME"
    strSQL = strSQL & "        END AS EZAPPI_NAME2, "
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_FZAPPI_CODE1 AS FZAPPI_CODE2, "
    strSQL = strSQL & "        CASE WHEN RCPT_TRAN.RCPT_ADD_FZAPPI_CODE1 = " & P_ZAPPI_無し
    strSQL = strSQL & "            THEN '' "
    strSQL = strSQL & "            ELSE FZAPPINAME2.NAME_NAME"
    strSQL = strSQL & "        END AS FZAPPI_NAME2, "
'▲2018/8/30 EGL
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_BUTU AS BUTU, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_UCODE AS UCODE, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_YCODE AS YCODE, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_CNO AS CNTANO, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEV_EXPIRE AS SEV_EXPIRE, "
    
'▼ 2007/08/23 tajima add
    'キャンペーン情報
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEV1N AS YARD_SEV1N, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEV2N AS YARD_SEV2N, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEV3N AS YARD_SEV3N, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ENDEN AS YARD_ENDEN, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEV_EXPIRE_MONTH AS YARD_SEV_EXMONTH, "
'▲ 2007/08/23 tajima add
    
    '請求金額関係
    strSQL = strSQL & "        RCPT_TRAN.RCPT_RENTKG AS RENTKG, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_FIRSTKG AS FIRSTKG, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_EZAPPI AS EZAPPI, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_EZAPPI_DAILY AS EZAPPI_DAILY, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_EZAPPI_CODE1 AS ADD_EZAPPI_CODE1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_EZAPPI1 AS ADD_EZAPPI1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_EZAPPI_DAILY1 AS ADD_EZAPPI_DAILY1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_EZAPPI_CODE2 AS ADD_EZAPPI_CODE2, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_EZAPPI2 AS ADD_EZAPPI2, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_EZAPPI_DAILY2 AS ADD_EZAPPI_DAILY2, "
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_FZAPPI AS FZAPPI, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_FZAPPI_CODE1 AS ADD_FZAPPI_CODE1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_FZAPPI1 AS ADD_FZAPPI1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_FZAPPI_CODE2 AS ADD_FZAPPI_CODE2, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ADD_FZAPPI2 AS ADD_FZAPPI2, "
    strSQL = strSQL & "        FZAPPI2NAME.NAME_NAME AS ADD_FZAPPI2_NAME, " '2015/01/08 add 初回雑費２の名称
    
    '保証金
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SECUKG + RCPT_TRAN.RCPT_HOSHOU_WARIBIKI AS SECUKG, " '2011/06/18 割引前に戻す
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SECUKG_KASE AS SECUKG_KASE, "     'INSERT 2013/04/10 K.ISHIZAKA
    '▼2008/12/04 hirano Add 保証金変更フラグ取得
    '保証金変更フラグ
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SECUKG_FLG AS SECUKG_FLG, "
    '▲2008/12/04 hirano Add
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SUMIKG AS SUMI_RENTKG, "
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEIKYU_KIKAN AS SEIKYUKIKAN, "
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_UKDATE AS UKDATE, "
    
    strSQL = strSQL & "        YOUKTUSER.USER_NAME AS UNAME, "
    strSQL = strSQL & "        YOUKTUSER.USER_KANA AS UKANA, "
    strSQL = strSQL & "        YOUKTUSER.USER_TANM AS UTANM, "
    strSQL = strSQL & "        YOUKTUSER.USER_TAKA AS UTAKA, "
    strSQL = strSQL & "        YOUKTUSER.USER_YUBINO AS UYUBINO, "
    strSQL = strSQL & "        YOUKTUSER.USER_ADR_1 AS UADR1, "
    strSQL = strSQL & "        YOUKTUSER.USER_ADR_2 AS UADR2, "
    strSQL = strSQL & "        YOUKTUSER.USER_ADR_3 AS UADR3, "
    strSQL = strSQL & "        YOUKTUSER.USER_TEL AS UTEL, "
    strSQL = strSQL & "        YOUKTUSER.USER_KEITAI AS UKEITAI, "
    strSQL = strSQL & "        YOUKTUSER.USER_FAX AS UFAX, "
    strSQL = strSQL & "        YOUKTUSER.USER_KKBN AS USERKKBN, "
    strSQL = strSQL & "        YOUKTUSER.USER_MAIL AS UMAIL, "
    strSQL = strSQL & "        YOUKTUSER.USER_SDAY AS USDAY, "
    strSQL = strSQL & "        YOUKTUSER.USER_BIKO AS UBIKO, "
    strSQL = strSQL & "        YOUKTUSER.USER_SKBN AS USKBN, "
    
    strSQL = strSQL & "        YOUKT_UKKBN AS UKKBN, "
    
'    '2006/04/28 S.SHIBAZAKI 乗用具チェック追加。チェックされている場合、申込書に誓約書添付
'    strSql = strSql & "        RCPT_VEHICLE AS VEHICLE, "
    
    '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
    strSQL = strSQL & "        RCPT_TRAN.RCPT_CARG_ACPTNO AS CARGACPTNO, "
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_HOSYICD AS HOSYICD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_HOSYO_CD AS HOSYO_KAISHA_CD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_NEW_HOSPLN_IKOU AS NEW_HOSPLN_IKOU, " '2018/08/06 EGL INS
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_CD AS SENDCD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_STDATE AS STDATE, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_DC_STDATE AS DC_STDATE, "         'I:20081001 N.Imai
    strSQL = strSQL & "        RCPT_TRAN.RCPT_DC_ENABLE AS DC_ENABLE, "         'I:20081001 N.Imai
    strSQL = strSQL & "        RCPT_TRAN.RCPT_MITU_CD AS MITUCD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEIKYU_CD AS SEIKYUCD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KAGIICD AS KAGIICD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KAGI_NO AS KAGINO, "
    '2006/06/01 SHIBAZAKI セキュリティカード対応
'    strSQL = strSQL & "        RCPT_TRAN.RCPT_SECURITY_CARD AS SECURITYCARD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KAGI_LENDNUM AS KAGILENDNUM, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KAINO AS KAIJONO, "               'INSERT 2014/03/11 MIYAMOTO
    strSQL = strSQL & "        RCPT_TRAN.RCPT_BUIL_KAGI_NO AS RCPT_BUIL_KAGI_NO, " 'INSERT 2021/06/02 EGL
    strSQL = strSQL & "        RCPT_TRAN.RCPT_HOSYD AS HOSYD, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_HOSYB AS HOSYB, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_NAME AS SENDNAME, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_HONORIFIC AS SENDHONORIFIC, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_YUBINO AS SENDYUBINO, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_ADR_1 AS SENDADR1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_ADR_2 AS SENDADR2, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_ADR_3 AS SENDADR3, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEND_TEL AS SENDTEL, "
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SEIKI AS SEIKI,"                  'INSERT 2010/03/29 K.ISHIZAKA
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KEY_SEND_DATE AS KEY_SEND_DATE, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ARRIVAL_EXPECT AS ARRIVALEXPECT , "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SPECIFY_TIME AS SPECIFYTIME, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SLIP_NO AS SLIPNO, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_DERIVERY_STATUS AS DERIVERYSTATUS, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_ARRIVAL_DATE AS ARRIVALDATE, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_YAMATO_DATE AS RCPT_YAMATO_DATE, " 'INSERT 2019/03/22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
    strSQL = strSQL & "        RCPT_TRAN.RCPT_UKTANTO AS UKTANTO, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_CYDATE AS CYDATE, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_CREASON AS CREASON, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_PAYDATE AS PAYDATE, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_UKNOTE AS UKNOTE, "
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KAIHI AS KAIHI,"                  'INSERT 2010/03/29 K.ISHIZAKA
    
    '↓INSERT 2010/04/20 SHIBAZAKI
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_NAME     AS RCPT_KREN_NAME, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_KANA     AS RCPT_KREN_KANA, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_YUBINO   AS RCPT_KREN_YUBINO, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_ADR_1    AS RCPT_KREN_ADR_1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_ADR_2    AS RCPT_KREN_ADR_2, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_ADR_3    AS RCPT_KREN_ADR_3, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_TEL      AS RCPT_KREN_TEL, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_KEITAI   AS RCPT_KREN_KEITAI, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_FAX      AS RCPT_KREN_FAX, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KREN_ZOKUGARA AS RCPT_KREN_ZOKUGARA, "
    '↑INSERT 2010/04/20 SHIBAZAKI
    strSQL = strSQL & "        RCPT_TRAN.RCPT_COMMENT_1 AS COMMENT1, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_COMMENT_2 AS COMMENT2, "
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SLOPE     AS RCPT_SLOPE, "              'INSERT 2011/04/11 M.HONDA
    strSQL = strSQL & "        RCPT_TRAN.RCPT_TPOINT    AS RCPT_TPOINT, "             'INSERT 2013/05/21 M.HONDA
    strSQL = strSQL & "        RCPT_TRAN.RCPT_KOUKBN    AS RCPT_KOUKBN, "             'INSERT 2013/10/16 M.HONDA
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SHOUMEI_KDATE,"         'INSERT 2015/07/16 K.ISHIZAKA
    strSQL = strSQL & "        RCPT_TRAN.RCPT_NEN_KBN, "              'INSERT 2015/03/24 M.HONDA
    strSQL = strSQL & "        RCPT_TRAN.RCPT_NEN_JIREI, "            'INSERT 2015/03/24 M.HONDA
    strSQL = strSQL & "        RCPT_TRAN.RCPT_NEN_SEIKI, "            'INSERT 2015/03/24 M.HONDA
    strSQL = strSQL & "        RCPT_TRAN.RCPT_NEN_BIKO, "             'INSERT 2015/03/24 M.HONDA
    
    strSQL = strSQL & "        RCPT_TRAN.RCPT_SYOKISAN_DATE "         '2018/03/16 M.HONDA INS
    
'↓ADD 2019/06/24 egl.YSUZUKI
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_SHOUMEI_UP_UM "                 '証明書のアップロード
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_SHOUMEI_FCNT "                  'フォロー回数
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_SHOUMEI_FDATE "
'↑ADD 2019/06/24 egl.YSUZUKI

    strSQL = strSQL & "      , RCPT_TRAN.RCPT_PLAN_CD "                 'add 2020/09/28 プランコード
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_SHOUMEI_ADRSAI "          'add 2020/11/16 証明書住所差異

    '2020/12/01 [EGL] INS Start
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_COUPON_CD "             'クーポンCD
    strSQL = strSQL & "      , PGPA_TABL.PGPAT_BIKON AS COUPON_TXT "  '紹介会社名
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_USER_COUPON_CD "        'ユーザ入力クーポンCD
    strSQL = strSQL & "      , RCPT_TRAN.RCPT_USER_COUPON_TXT "       'ユーザ入力紹介会社名
    '2020/12/01 [EGL] INS End

    fncMakeRcptSelect = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeCommonSelect
'        機能             :データ検索SQL作成(共通SELECT句)
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeCommonSelect(blnMove As Boolean) As String

    Dim strSQL      As String
    Dim str元受取条件sql As String

' ▼ 2008/03/23 tajima chg-all SQLにて NOW_YARD,NOW_PRIC,NOW_CNTA 付加
    strSQL = strSQL & "        NOW_YARD.YARD_NAME, "
    strSQL = strSQL & "        NOW_YARD.YARD_ADDR_1 + NOW_YARD.YARD_ADDR_2 + NOW_YARD.YARD_ADDR_3 AS YADDR, "
    strSQL = strSQL & "        NOW_YARD.YARD_BIKON,"                           'INSERT 20081104 N.MIURA
' ▼ 2007/08/23 tajima del
'    strSQL = strSQL & "        YARD_SEV1N, "
'    strSQL = strSQL & "        YARD_SEV2N, "
'    strSQL = strSQL & "        YARD_SEV3N, "
'    strSQL = strSQL & "        YARD_ENDEN, "
'    strSQL = strSQL & "        YARD_SEV_EXMONTH, "
' ▲ 2007/08/23 tajima del
    
    '2006/09/06 SHIBAZAKI 営業期間追加
    strSQL = strSQL & "        NOW_YARD.YARD_BEGIN_DAY, "
    strSQL = strSQL & "        NOW_YARD.YARD_RENTEND_DAY, "
    strSQL = strSQL & "        NOW_YARD.YARD_TIHO_KBN AS YARD_TIHO_KBN, "       'INSERT 2014/06/20 MIYAMOTO
    strSQL = strSQL & "        STEPNAME.NAME_NAME AS STEP_NAME, "
    strSQL = strSQL & "        SYUNAME.NAME_NAME AS SYU_NAME, "
    strSQL = strSQL & "        NOW_CNTA.CNTA_USAGE, "
    strSQL = strSQL & "        USAGENAME.NAME_NAME AS USAGE_NAME, "
    strSQL = strSQL & "        NOW_CNTA.CNTA_SIZE, "
    strSQL = strSQL & "        NOW_CNTA.CNTA_PRICE_DIFF, "
    strSQL = strSQL & "        NOW_CNTA.CNTA_REASON, "
' ▼ 2008/11/19 hirano Add 割引き情報 割引対象フラグ取得
    strSQL = strSQL & "        NOW_CNTA.CNTA_NEBIKI_DISABLE, "
' ▲ 2008/11/19 hirano

' ▼ 2018/09/21 egl add
    strSQL = strSQL & "        NOW_CNTA.CNTA_FLOOR,"
' ▲ 2018/09/21 egl add

    strSQL = strSQL & "        NOW_PRIC.PRIC_PRICE, "
    
   
    strSQL = strSQL & "        NOW_PRIC.PRIC_FPRICE, "  '2019/08/08 M.HONDA INS
'    strSQL = strSQL & "        NOW_PRIC.PRIC_FEZAPPI AS FEZAPPI , "                         '2019/08/08 M.HONDA INS


    
    strSQL = strSQL & "        NOW_PRIC.PRIC_EZAPPI, "                          'INSERT START 2013/04/10 K.ISHIZAKA
    
    
    strSQL = strSQL & "        NOW_PRIC.PRIC_EZAPPI_CODE, "
    strSQL = strSQL & "        CASE WHEN ISNULL(NOW_PRIC.PRIC_EZAPPI_CODE," & P_ZAPPI_無し & ") = " & P_ZAPPI_無し
    strSQL = strSQL & "        THEN '' "
    strSQL = strSQL & "        ELSE EZAPPINAME.NAME_NAME "
    strSQL = strSQL & "        END AS PRIC_EZAPPI_NAME, "                       'INSERT END   2013/04/10 K.ISHIZAKA
    

    strSQL = strSQL & "        YOUKTUSER.USER_RNAME AS USER_RNAME, "
    strSQL = strSQL & "        YOUKTUSER.USER_RKANA AS USER_RKANA, "
    strSQL = strSQL & "        YOUKTUSER.USER_RPOST AS USER_RPOST, "
    strSQL = strSQL & "        YOUKTUSER.USER_RADR_1 AS USER_RADR_1, "
    strSQL = strSQL & "        YOUKTUSER.USER_RADR_2 AS USER_RADR_2, "
    strSQL = strSQL & "        YOUKTUSER.USER_RADR_3 AS USER_RADR_3, "
    strSQL = strSQL & "        YOUKTUSER.USER_RTEL AS USER_RTEL, "
    strSQL = strSQL & "        YOUKTUSER.USER_RKEITAI AS USER_RKEITAI, "
    strSQL = strSQL & "        YOUKTUSER.USER_RFAX AS USER_RFAX, "
    
    '↓INSERT 2010/04/20 SHIBAZAKI
    strSQL = strSQL & "        YOUKTUSER.USER_BIRTHDAY    AS USER_BIRTHDAY, "
    strSQL = strSQL & "        YOUKTUSER.USER_KNAME       AS USER_KNAME, "
    strSQL = strSQL & "        YOUKTUSER.USER_KKANA       AS USER_KKANA, "
    strSQL = strSQL & "        YOUKTUSER.USER_JOB         AS  USER_JOB, "
    strSQL = strSQL & "        YOUKTUSER.USER_KPOST       AS USER_KPOST, "
    strSQL = strSQL & "        YOUKTUSER.USER_KADR_1      AS USER_KADR_1, "
    strSQL = strSQL & "        YOUKTUSER.USER_KADR_2      AS USER_KADR_2, "
    strSQL = strSQL & "        YOUKTUSER.USER_KADR_3      AS USER_KADR_3, "
    strSQL = strSQL & "        YOUKTUSER.USER_KTEL        AS USER_KTEL, "
    strSQL = strSQL & "        YOUKTUSER.USER_KHOUR_START AS USER_KHOUR_START, "
    strSQL = strSQL & "        YOUKTUSER.USER_KHOUR_END   AS USER_KHOUR_END, "
    strSQL = strSQL & "        YOUKTUSER.USER_NENSHU      AS USER_NENSHU, "
    '↑INSERT 2010/04/20 SHIBAZAKI
        
    strSQL = strSQL & "        YOUKTUSER.USER_SEIBETU     AS USER_SEIBETU, "  '2014/10/14 add tajima
     
    strSQL = strSQL & "        YOUKT_YUKBN AS YUKBN, "
    strSQL = strSQL & "        YOUKT_MOVEKBN, "
    strSQL = strSQL & "        YOUKT_MOVE_TANTO, "  '2007/04/13 add tajima
    strSQL = strSQL & "        YOUKT_GENKBN, "      '2007/08/23 add tajima
    strSQL = strSQL & "        YOUKT_USE_PERIOD, "  '2007/09/25 add tajima
    
    strSQL = strSQL & "        YUKBNNAME.NAME_NAME AS YUKBN_NAME, "
    
    strSQL = strSQL & "        YOUKT_TKDATE AS LIMITDATE, "
    '2009/04/01 ADD <S> hirano 予約備考追加
    strSQL = strSQL & "        YOUKT_BIKO AS YOUKT_COMMENT, "
    '2009/04/01 ADD <E> hirano
    '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
'    strSql = strSql & "        CARGFILE.CARG_ACPTNO AS CARGACPTNO, "
        
    strSQL = strSQL & "        MOVEKBNNAME.NAME_NAME AS MOVEKBN_NAME, "
    
    If blnMove Then
    
        strSQL = strSQL & "        MOTO_CARG.CARG_ACPTNO AS MOTO_CARG_ACPTNO, "
        
        
        '2006/05/10 SHIBAZAKI 契約変更系の受付の場合、契約開始日は移動元契約の契約開始日の翌月以降であること
        strSQL = strSQL & "        MOTO_CARG.CARG_STDATE AS MOTO_STDATE, "
        
        '2006/06/01 SHIBAZAKI 鍵情報の初期値はコンテナマスタから取得する
'        strSQL = strSQL & "        MOTO_CARG.CARG_KAGIICD AS KAGIICD, "
        
        strSQL = strSQL & "        MOTO_CARG.CARG_KAIUKEDATE AS MOTO_KAIUKEDATE, "
        strSQL = strSQL & "        MOTO_CARG.CARG_KAITANTO AS MOTO_KAITANTO, "
        strSQL = strSQL & "        MOTO_CARG.CARG_KAICD AS MOTO_KAICD, "
        strSQL = strSQL & "        MOTO_CARG.CARG_KAIRIYUCD AS MOTO_KAIRIYUCD, "
        strSQL = strSQL & "        MOTO_CARG.CARG_UCODE AS MOTO_UCODE, "
        strSQL = strSQL & "        MOTO_CARG.CARG_UKTANTO AS MOTO_UKTANTO, "
        strSQL = strSQL & "        MOTO_CARG.CARG_YCODE AS MOTO_YCODE, "
        strSQL = strSQL & "        MOTO_CARG.CARG_NO AS MOTO_CARGNO, "
        strSQL = strSQL & "        MOTO_CARG.CARG_KAGIICD AS MOTO_KAGIICD, "    'INSERT 2014/04/17 MIYAMOTO
        strSQL = strSQL & "        MOTO_CARG.CARG_RENTKG+MOTO_CARG.CARG_SYOZEI AS MOTO_RENTKG, "
        strSQL = strSQL & "        MOTO_YARD.YARD_TIHO_KBN AS MOTO_YARD_YARD_TIHO_KBN, "    'INSERT 2014/06/20 MIYAMOTO
        strSQL = strSQL & "        MOTOUSER.USER_KANA AS MOTO_USERKANA, "
        strSQL = strSQL & "        HOSYORYAKU.NAME_NAME AS HOSYORYAKU_NAME, "
    '▼2008/03/23 tajima chg 前ヤードが解約ヤードかチェック
        strSQL = strSQL & " CASE WHEN ISNULL(MOTO_YARD.YARD_RENTEND_DAY,'2999/12/31') <= GETDATE() THEN '1' ELSE '0' END AS MOTO_YARD_END_FLG,"
        strSQL = strSQL & " MOTO_RCPT.RCPT_RENTKG, "                            'INSERT START 2013/04/10 K.ISHIZAKA
        strSQL = strSQL & " MOTO_RCPT.RCPT_EZAPPI, "
        strSQL = strSQL & " MOTO_PRIC.PRIC_EZAPPI_CODE AS RCPT_EZAPPI_CODE, "
        strSQL = strSQL & " CASE WHEN ISNULL(MOTO_PRIC.PRIC_EZAPPI_CODE," & P_ZAPPI_無し & ") = " & P_ZAPPI_無し
        strSQL = strSQL & " THEN '' "
        strSQL = strSQL & " ELSE MOTO_EZAPPINAME.NAME_NAME "
        strSQL = strSQL & " END AS RCPT_EZAPPI_NAME, "                          'INSERT END   2013/04/10 K.ISHIZAKA
    Else
    
        strSQL = strSQL & " ''  AS MOTO_CARG_ACPTNO, "
    
        strSQL = strSQL & " '0' AS MOTO_YARD_END_FLG,"
    '▲2008/03/23 tajima chg
        strSQL = strSQL & " 0 AS RCPT_RENTKG, "                                 'INSERT START 2013/04/10 K.ISHIZAKA
        strSQL = strSQL & " 0 AS RCPT_EZAPPI, "
        strSQL = strSQL & " " & P_ZAPPI_無し & " AS RCPT_EZAPPI_CODE, "
        strSQL = strSQL & " '' AS RCPT_EZAPPI_NAME, "                           'INSERT END   2013/04/10 K.ISHIZAKA
    End If
    
    '前回の契約データの完了日と解約日
    strSQL = strSQL & "        ( SELECT TOP 1 CAST(ISNULL(CARG_KYDATE,'') AS CHAR) + ',' "
    strSQL = strSQL & "                     + ISNULL(CAST(CARG_MNT_CMPDATE AS CHAR),'') "
    '2006/05/10 SHIBAZAKI システム日付が前回の契約データの初回契約日以降じゃないと受付完了させない
    strSQL = strSQL & "                     + ','+ ISNULL(CAST(CARG_FSDATE AS CHAR),'') "
    strSQL = strSQL & "            FROM CARG_FILE "
    strSQL = strSQL & "           WHERE CARG_YCODE=YOUKT_YCODE "
    strSQL = strSQL & "             AND CARG_NO=YOUKT_NO "
    strSQL = strSQL & "             AND CARG_ACPTNO <> ISNULL(RCPT_TRAN.RCPT_CARG_ACPTNO, '') " 'INSERT 2020/03/17 N.IMAI
    strSQL = strSQL & "           ORDER BY CARG_FSDATE DESC ) AS CARGDATES, "
    
    '20150108 M.HONDA INS
    strSQL = strSQL & "        DLVRT_OUT_ADR, "
    strSQL = strSQL & "        DLVRT_PRICE, "
    strSQL = strSQL & "        DLVRT_KYORI, "
    strSQL = strSQL & "        DLVRT_KIBO_DATE1, "
    strSQL = strSQL & "        TIMECD1.NAME_NAME as TIMECD1, "
    strSQL = strSQL & "        DLVRT_KIBO_DATE2, "
    strSQL = strSQL & "        TIMECD2.NAME_NAME as TIMECD2, "
    strSQL = strSQL & "        DLVRT_KIBO_DATE3, "
    strSQL = strSQL & "        TIMECD3.NAME_NAME as TIMECD3, "
    '20150108 M.HONDA INS
    '20150421 N.IMAI INS
    strSQL = strSQL & "        DLVRT_BUTU_DANBORU, "
    strSQL = strSQL & "        DLVRT_BUTU_KESU, "
    strSQL = strSQL & "        DLVRT_BUTU_KAGU, "
    strSQL = strSQL & "        DLVRT_BUTU_SEIKATU, "
    strSQL = strSQL & "        DLVRT_BUTU_REJYA, "
    strSQL = strSQL & "        DLVRT_BUTU_ETE, "
    '20150421 N.IMAI INS

    fncMakeCommonSelect = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeFrom
'        機能             :データ検索SQL作成(FROM句)
'        IN               :①True=既存受付を検索 False=新規受付を検索
'                         :②True=元契約をJOINする　False=元契約はない
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeFrom(blnRcpt As Boolean, blnMove As Boolean) As String

    Dim strSQL      As String
    Dim strDbName   As String   '2020/12/01 [EGL] INS
    
    strSQL = strSQL & "   FROM YOUK_TRAN "

' ▼ 2008/03/23 tajima chg-all SQLにて NOW_YARD,NOW_PRIC,NOW_CNTA 付加
    strSQL = strSQL & "  INNER JOIN YARD_MAST NOW_YARD "
    strSQL = strSQL & "     ON NOW_YARD.YARD_CODE = YOUKT_YCODE " & vbCrLf
    
    strSQL = strSQL & "  INNER JOIN CNTA_MAST NOW_CNTA"
    strSQL = strSQL & "     ON NOW_CNTA.CNTA_CODE = YOUKT_YCODE "
    strSQL = strSQL & "    AND NOW_CNTA.CNTA_NO = YOUKT_NO " & vbCrLf
    
    
    '20150108 M.HONDA INS
    strSQL = strSQL & "   LEFT OUTER JOIN "
    strSQL = strSQL & "       DLVR_TRAN ON "
    strSQL = strSQL & "       YOUKT_UKNO = DLVRT_NO " & vbCrLf
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST TIMECD1 "
    strSQL = strSQL & "     ON TIMECD1.NAME_ID = '016' "
    strSQL = strSQL & "    AND TIMECD1.NAME_CODE = DLVRT_KIBO_TIMECD1 " & vbCrLf
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST TIMECD2 "
    strSQL = strSQL & "     ON TIMECD2.NAME_ID = '016' "
    strSQL = strSQL & "    AND TIMECD2.NAME_CODE = DLVRT_KIBO_TIMECD2 " & vbCrLf
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST TIMECD3 "
    strSQL = strSQL & "     ON TIMECD3.NAME_ID = '016' "
    strSQL = strSQL & "    AND TIMECD3.NAME_CODE = DLVRT_KIBO_TIMECD3 " & vbCrLf
    '20150108 M.HONDA INS
    
    strSQL = strSQL & "  INNER JOIN NAME_MAST STEPNAME "
    strSQL = strSQL & "     ON STEPNAME.NAME_ID = '090' "
    strSQL = strSQL & "    AND STEPNAME.NAME_CODE = NOW_CNTA.CNTA_STEP " & vbCrLf
    
    strSQL = strSQL & "  INNER JOIN NAME_MAST SYUNAME "
    strSQL = strSQL & "     ON SYUNAME.NAME_ID = '040' "
    strSQL = strSQL & "    AND SYUNAME.NAME_CODE = NOW_CNTA.CNTA_SYU " & vbCrLf
    
    If blnRcpt Then
        '受付トランを内部ジョイン
        strSQL = strSQL & "  INNER JOIN RCPT_TRAN "
        strSQL = strSQL & "     ON RCPT_NO = YOUKT_RCPT_NO " & vbCrLf
    Else
        'INSERT 2020/03/17 N.IMAI Start
        strSQL = strSQL & "  LEFT JOIN RCPT_TRAN "
        strSQL = strSQL & "     ON RCPT_NO = YOUKT_RCPT_NO " & vbCrLf
        'INSERT 2020/03/17 N.IMAI End
    End If
    
    If blnMove Then
        '契約トランを内部ジョイン
        strSQL = strSQL & "  INNER JOIN CARG_FILE MOTO_CARG "
        strSQL = strSQL & "     ON MOTO_CARG.CARG_ACPTNO = YOUKT_MOTO_ACPTNO " & vbCrLf
        
        strSQL = strSQL & "  INNER JOIN NAME_MAST HOSYORYAKU "
        strSQL = strSQL & "     ON HOSYORYAKU.NAME_ID = '200' "
        'strSQL = strSQL & "    AND HOSYORYAKU.NAME_CODE = MOTO_CARG.CARG_HOSYICD " & vbCrLf
        strSQL = strSQL & "    AND HOSYORYAKU.NAME_CODE = CASE WHEN MOTO_CARG.CARG_HOSYICD = '2' THEN '1' ELSE MOTO_CARG.CARG_HOSYICD END " & vbCrLf    'INSERT 2023/05/01 N.IMAI
        
        strSQL = strSQL & "  INNER JOIN USER_MAST MOTOUSER "
        strSQL = strSQL & "     ON MOTOUSER.USER_CODE = MOTO_CARG.CARG_UCODE " & vbCrLf
    
        '▼2008/03/23 add tajima
        '元ヤードを内部結合
        strSQL = strSQL & "  INNER JOIN YARD_MAST MOTO_YARD ON MOTO_CARG.CARG_YCODE = MOTO_YARD.YARD_CODE" & vbCrLf
        '初回時のみ
'        If blnRcpt = False Then                                                'DELETE 2013/04/10 K.ISHIZAKA
            '元コンテナを内部結合（元の価格情報を取るため)
            strSQL = strSQL & "  INNER JOIN CNTA_MAST MOTO_CNTA ON MOTO_CNTA.CNTA_CODE = MOTO_CARG.CARG_YCODE AND MOTO_CNTA.CNTA_NO = MOTO_CARG.CARG_NO" & vbCrLf
            '元価格表を外部結合（元の初回毎月雑費を得たいためだけ)
            strSQL = strSQL & "   LEFT OUTER JOIN PRIC_TABL MOTO_PRIC ON MOTO_PRIC.PRIC_YCODE = MOTO_CNTA.CNTA_CODE AND MOTO_PRIC.PRIC_USAGE = MOTO_CNTA.CNTA_USAGE AND MOTO_PRIC.PRIC_SIZE = MOTO_CNTA.CNTA_SIZE  AND MOTO_PRIC.PRIC_STEP = MOTO_CNTA.CNTA_STEP" & vbCrLf
            '元受付を外部結合
            strSQL = strSQL & "  LEFT OUTER JOIN RCPT_TRAN MOTO_RCPT ON MOTO_CARG.CARG_UKNO = MOTO_RCPT.RCPT_NO" & vbCrLf
            '元初回雑費名＆毎月雑費名の取得JOIN
            strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST MOTO_FZAPPINAME "
            strSQL = strSQL & "     ON MOTO_FZAPPINAME.NAME_ID = '096' "
            strSQL = strSQL & "    AND MOTO_FZAPPINAME.NAME_CODE = MOTO_PRIC.PRIC_FZAPPI_CODE" & vbCrLf
            strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST MOTO_EZAPPINAME "
            strSQL = strSQL & "     ON MOTO_EZAPPINAME.NAME_ID = '096' "
            strSQL = strSQL & "    AND MOTO_EZAPPINAME.NAME_CODE = MOTO_PRIC.PRIC_EZAPPI_CODE" & vbCrLf
'        End If                                                                 'DELETE 2013/04/10 K.ISHIZAKA
        '▲2008/03/23 add tajima
    End If
    
    strSQL = strSQL & "  INNER JOIN NAME_MAST YUKBNNAME "
    strSQL = strSQL & "     ON YUKBNNAME.NAME_ID = '080' "
    strSQL = strSQL & "    AND YUKBNNAME.NAME_CODE = YOUKT_YUKBN " & vbCrLf
    
    strSQL = strSQL & IIf(blnRcpt, " INNER ", " LEFT OUTER ") & " JOIN USER_MAST YOUKTUSER "
    strSQL = strSQL & "     ON YOUKTUSER.USER_CODE = " & IIf(blnRcpt, "RCPT_TRAN.RCPT_UCODE", "YOUKT_UCODE") & vbCrLf
    
    '▼2008/03/23 del tajima CARG_FILEはどこでも使ってない
'    strSQL = strSQL & "   LEFT OUTER JOIN CARG_FILE CARGFILE "
'    strSQL = strSQL & "     ON CARGFILE.CARG_UKNO = YOUKT_RCPT_NO "
    '▲2008/03/23 del tajima
    
    strSQL = strSQL & "   LEFT OUTER JOIN PRIC_TABL NOW_PRIC"
    strSQL = strSQL & "     ON NOW_PRIC.PRIC_YCODE = NOW_CNTA.CNTA_CODE "
    strSQL = strSQL & "    AND NOW_PRIC.PRIC_USAGE = NOW_CNTA.CNTA_USAGE"
    strSQL = strSQL & "    AND NOW_PRIC.PRIC_SIZE = NOW_CNTA.CNTA_SIZE"
    strSQL = strSQL & "    AND NOW_PRIC.PRIC_STEP = NOW_CNTA.CNTA_STEP" & vbCrLf
    
    '▼2018/08/30 CHG EGL 元コードを価格→受付
    'strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST FZAPPINAME "
    'strSQL = strSQL & "     ON FZAPPINAME.NAME_ID = '096' "
    'strSQL = strSQL & "    AND FZAPPINAME.NAME_CODE = NOW_PRIC.PRIC_FZAPPI_CODE" & vbCrLf
    
    'strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST EZAPPINAME "
    'strSQL = strSQL & "     ON EZAPPINAME.NAME_ID = '096' "
    'strSQL = strSQL & "    AND EZAPPINAME.NAME_CODE = NOW_PRIC.PRIC_EZAPPI_CODE" & vbCrLf
    
    If blnRcpt = False Then
        strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST FZAPPINAME "
        strSQL = strSQL & "     ON FZAPPINAME.NAME_ID = '096' "
        strSQL = strSQL & "    AND FZAPPINAME.NAME_CODE = NOW_PRIC.PRIC_FZAPPI_CODE" & vbCrLf
        
        strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST EZAPPINAME "
        strSQL = strSQL & "     ON EZAPPINAME.NAME_ID = '096' "
        strSQL = strSQL & "    AND EZAPPINAME.NAME_CODE = NOW_PRIC.PRIC_EZAPPI_CODE" & vbCrLf
    Else
    
        strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST FZAPPINAME "
        strSQL = strSQL & "     ON FZAPPINAME.NAME_ID = '096' "
        strSQL = strSQL & "    AND FZAPPINAME.NAME_CODE = RCPT_TRAN.RCPT_FZAPPI_CODE" & vbCrLf
        
        strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST EZAPPINAME "
        strSQL = strSQL & "     ON EZAPPINAME.NAME_ID = '096' "
        strSQL = strSQL & "    AND EZAPPINAME.NAME_CODE = RCPT_TRAN.RCPT_EZAPPI_CODE" & vbCrLf
        
        strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST FZAPPINAME2"
        strSQL = strSQL & "   ON FZAPPINAME.NAME_ID = '096'"
        strSQL = strSQL & "   AND FZAPPINAME2.NAME_CODE = RCPT_TRAN.RCPT_ADD_FZAPPI_CODE1" & vbCrLf
        '2
        strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST EZAPPINAME2 "
        strSQL = strSQL & "   ON EZAPPINAME.NAME_ID = '096'"
        strSQL = strSQL & "   AND EZAPPINAME2.NAME_CODE = RCPT_TRAN.RCPT_ADD_EZAPPI_CODE1" & vbCrLf
    End If
    '▲2018/08/30 CHG EGL
    
    strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST USAGENAME "
    strSQL = strSQL & "     ON USAGENAME.NAME_ID = '086' "
    strSQL = strSQL & "    AND USAGENAME.NAME_CODE = NOW_CNTA.CNTA_USAGE" & vbCrLf

    strSQL = strSQL & "  LEFT OUTER JOIN NAME_MAST MOVEKBNNAME "
    strSQL = strSQL & "     ON MOVEKBNNAME.NAME_ID = '098' "
    strSQL = strSQL & "    AND MOVEKBNNAME.NAME_CODE = YOUKT_MOVEKBN "
    strSQL = strSQL & "    AND MOVEKBNNAME.NAME_CODE <> " & P_MOVEKBN_無し & vbCrLf
    
    '▼2015/01/08 add 初回雑費2の名称を求める
    If blnRcpt Then
        strSQL = strSQL & "   LEFT OUTER JOIN NAME_MAST FZAPPI2NAME "
        strSQL = strSQL & "     ON FZAPPI2NAME.NAME_ID = '096' "
        strSQL = strSQL & "    AND FZAPPI2NAME.NAME_CODE = RCPT_TRAN.RCPT_ADD_FZAPPI_CODE2" & vbCrLf
    End If
    '▲2015/01/08 add
    
    '2020/12/01 [EGL] INS Start
    If blnRcpt Then
        strDbName = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'"), "")
        
        strSQL = strSQL & "   LEFT OUTER JOIN " & strDbName & ".dbo.PGPA_TABL "
        strSQL = strSQL & "     ON PGPAT_PGP1B = 'FVS400_COUPON' "
'        strSql = strSql & "    AND PGPAT_PGP2B = RCPT_COUPON_CD" & vbCrLf                              'DELETE 2020/12/15 K.KINEBUCHI
        strSQL = strSQL & "    AND PGPAT_PGP2B = RCPT_TRAN.RCPT_COUPON_CD" & vbCrLf                     'INSERT 2020/12/15 K.KINEBUCHI
    End If
    '2020/12/01 [EGL] INS End

    fncMakeFrom = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subDispData
'        機能             :検索結果画面表示
'        IN               :rsData=表示データ
'                         :blnExist=True:既存受付 False:新規受付
'                         :blnMove=True:移動予約 False:通常予約
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispData(rsData As Recordset, blnExist As Boolean, blnMove As Boolean)
    On Error GoTo ErrorHandler
    
    Dim strMoveNote     As String
    Dim strTemp()       As String
    
    txtAcptNo = rsData.Fields("CARGACPTNO")
    
    '▼ 2007/08/23 add tajima
    If Nz(rsData.Fields("YOUKT_GENKBN"), 0) = 0 Then
        strFormTitle = "受付入力：電話"
        blnネット予約 = False
    Else
        strFormTitle = "受付入力：ネット"
        blnネット予約 = True
    End If
    lblTitle.Caption = strFormTitle
    '▲ 2007/08/23 add tajima
    bln元ヤード解約 = False '2008/03/23 tajima add
    
    If Nz(rsData.Fields("MOVEKBN_NAME"), "") <> "" Then
        lblTitle.Caption = strFormTitle & "(" & rsData.Fields("MOVEKBN_NAME") & ")"
        txtMoveKbn = rsData.Fields("YOUKT_MOVEKBN")
        txtMoveKbnName = rsData.Fields("MOVEKBN_NAME")
    End If
    
    Call subDispUser(rsData)                    '顧客情報
    Call subDispContainer(rsData)               'コンテナ情報
    '2006/05/10 S.SHIBAZAKI 移動予約の場合、移動元が「保証金」の場合のみ保証金を引き継ぐ
    'Call subDispBill(rsData)                    '請求情報
    Call subDispBill(rsData, blnExist, blnMove) '請求情報
    
    '2020/12/01 [EGL] UPD Start
    'Call subDispReceipt(rsData)                 '受付情報
    Call subDispReceipt(rsData, blnExist)        '受付情報
    '2020/12/01 [EGL] UPD End
    
    Call subDispAppForm(rsData)                 '申込書情報
    Call subDispKey(rsData)                     '鍵情報
    Call subDispSecurity(rsData)                '保証情報
    Call subDispSend(rsData)                    '鍵送付情報
    Call subHoshouShow(0)                       '2017/12/20 add
    
    Call subDispDlvr(rsData)                    '運搬サービス情報
    
    If blnMove Then
        Call subDispMotoCarg(rsData)            '移動元契約情報
    End If
    
    '契約変更以外の場合、前回の契約データの完了日or解約日、以前の契約開始日は入力出来なくするために
    '前回の契約データの完了日と解約日を取得
'2017/07/31 ysuzuki start
'del    If Nz(Trim(rsData.Fields("CARGDATES"))) = "" Then
    If Nz(Trim(rsData.Fields("CARGDATES"))) = "" Or Nz(txtMotoAcptNo) = "" Then
'2017/07/31 ysuzuki end
        txtCargKyDate = ""
        txtCargMntCmpDate = ""
        '2006/05/10 SHIBAZAKI システム日付が前回の契約データの初回契約日以降じゃないと受付完了させない
        txtCargFsDate = ""
    Else
        strTemp = Split(rsData.Fields("CARGDATES"), ",")
        txtCargKyDate = Format(Trim(strTemp(0)), "yyyy/mm/dd")
        txtCargMntCmpDate = Format(Trim(strTemp(1)), "yyyy/mm/dd")
        '2006/05/10 SHIBAZAKI システム日付が前回の契約データの初回契約日以降じゃないと受付完了させない
        txtCargFsDate = Format(Trim(strTemp(2)), "yyyy/mm/dd")
    End If
    
    If blnExist Then
        '既存受付
        '日割り日数表示
        Call subDailyRate
        '既存受付の場合、顧客情報は表示のみ
        Call subUserReference
    Else
        '新規受付
        If blnMove Then
            '移動予約備考作成
            strMoveNote = ""
            Call CreateNewMoveMemo(rsData.Fields("YOUKT_MOVEKBN"), _
                                   rsData.Fields("MOVEKBN_NAME"), _
                                   rsData.Fields("MOTO_YCODE"), _
                                   rsData.Fields("MOTO_CARGNO"), _
                                   rsData.Fields("MOTO_UCODE"), _
                                   rsData.Fields("MOTO_USERKANA"), _
                                   rsData.Fields("MOTO_RENTKG"), _
                                   rsData.Fields("HOSYORYAKU_NAME"), _
                                   txtMotoAcptNo, _
                                   strMoveNote)
            txtReceiptUknote = strMoveNote
            '▼ 2008/03/23 add tajima
            If Nz(rsData.Fields("YOUKT_MOVEKBN"), "") = P_MOVEKBN_自社理由 And _
               Nz(rsData.Fields("MOTO_YARD_END_FLG"), "0") = 1 Then
                bln元ヤード解約 = True
                lblTitle.Caption = lblTitle.Caption & ":解約ﾔｰﾄﾞ"
            End If
            '▲ 2008/03/23 add tajima
            
        '2006/09/06 SHIBAZAKI 移動予約の場合も契約開始日のデフォルト値を表示する
'        Else
'            '新規受付(Not移動)の場合、契約開始日と起算日のデフォルト値を表示する。
'            Call subStDateDefault
        End If
        
        '2006/09/06 SHIBAZAKI 移動予約の場合も契約開始日のデフォルト値を表示する
        '新規受付の場合、契約開始日と起算日のデフォルト値を表示する。
        Call subStDateDefault
        
        '新規受付の場合、月額を日割計算
        '2008/12/23 ↓以下を抜き出し呼ぶように  Call cmdReCal_Click
        'txtBillFirstkg = fncDailyRate(txtBillRentkg)
        txtBillFirstkg = fncDailyRate(rsData.Fields("FRENTKG"))      '2019/08/08 M.HONDA INS
        txtBillFirstkg_sub = rsData.Fields("FRENTKG")  '2019/08/08 M.HONDA INS
         
        
        '' 2012/12/27 M.HONDA START
        '' 毎月雑費が保険料相当額の際には日割りは０円とする。
        If rsData.Fields("EZAPPI_CODE") = "68" Then
            txtBillEzappiDaily = txtBillEzappiDaily
            txtBillEzappiDaily = txtBillEzappiDaily_sub '2019/08/08 M.HONDA INS
        Else
            txtBillEzappiDaily = fncDailyRate(txtBillEzappi)
        End If
        '' 2012/12/27 M.HONDA END
        
        
        txtBillAddEzappiDaily1 = fncDailyRate(txtBillAddEzappi1)
        txtBillAddEzappiDaily2 = fncDailyRate(txtBillAddEzappi2)
    End If
    
    '予約状態が「受付」以外は画面を参照モードにする
    If txtReceiptYukbnCode <> P_YUKBN_受付 Then
        Call subReferenceMode
    '' 2012/07/04 M.HONDA START
        txtAppFormButu.ForeColor = P_YELLOW
    Else
                                                                                'INSERT START 2015/07/16 K.ISHIZAKA
        'ネット予約で初回請求分クレジット払いの場合、請求情報を参照モードにする
        If blnネット予約 And Nz(cmbBillSeikyuCd.VALUE, 0) = P_SEIKYUCD_クレジット Then
            Call subBillReference       '請求情報
        End If                                                                  'INSERT END   2015/07/16 K.ISHIZAKA
        txtAppFormButu.ForeColor = P_BLACK
    '' 2012/07/04 M.HONDA END
    End If
    
    '契約済みデータの場合、「受付完了」と表示
    If txtReceiptYukbnCode = P_YUKBN_契約 Then
        lblNewExist.Caption = P_LABEL_受付完了
    End If
    
    '2020/11/16 [EGL] INS Start
    cmbShoumeiAdrsai = rsData.Fields("RCPT_SHOUMEI_ADRSAI")
    Call cmbShoumeiAdrsai_AfterUpdate
    '2020/11/16 [EGL] INS End
    
    ''＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠
    '' 起算日確定後じゃないと処理できない。
    If chkContainerSevExpire.Visible = True Then
        If chk割引有効可否.VALUE <> 0 Then
            chkContainerSevExpire.VALUE = True
            chkContainerSevExpire_AfterUpdate
        End If
    End If
    ''＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠
    

    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subDispData" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispUser
'        機能             :顧客情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispUser(rsData As Recordset)
    On Error GoTo ErrorHandler
    Dim objField            As Object                           'INSERT 2010/04/28 SHIBAZAKI
    
    With rsData
        fraUserNewExist.VALUE = IIf(Nz(.Fields("UCODE")) = "", P_USER_OPT_新規, P_USER_OPT_既存)
        Call fraUserNewExist_AfterUpdate
        
        txtUserExistKana = .Fields("UKANA")
        cmbUserCode = Format(.Fields("UCODE"), "000000")
        
        '顧客情報
        txtUserName = .Fields("UNAME")
        txtUserKana = .Fields("UKANA")
        txtUserYubino = .Fields("UYUBINO")
        txtUserAdr1 = .Fields("UADR1")
        txtUserAdr2 = .Fields("UADR2")
        txtUserAdr3 = .Fields("UADR3")
        txtUserTanm = .Fields("UTANM")
        txtUserTaka = .Fields("UTAKA")
        txtUserTel = .Fields("UTEL")
        txtUserKeitai = .Fields("UKEITAI")
        txtUserFax = .Fields("UFAX")
        txtUserMail = .Fields("UMAIL")
        
        txtUserSday = .Fields("USDAY")
        txtUserBiko = .Fields("UBIKO")
        
        cmbUserKkbn = .Fields("USERKKBN")
        Call cmbUserKkbn_AfterUpdate
        cmbUserSkbn = .Fields("USKBN")
        strSkkbn = .Fields("USKBN")     '' 2013/09/19 M.HONDA INS
        Call cmbUserSkbn_AfterUpdate
        
        cmb_性別 = .Fields("USER_SEIBETU") '2014/10/14 add tajima
        Call cmb_性別_AfterUpdate          '2014/10/14 add tajima

        '書類送付先
        txtUserRname = .Fields("USER_RNAME")
        txtUserRkana = .Fields("USER_RKANA")
        txtUserRpost = .Fields("USER_RPOST")
        txtUserRadr1 = .Fields("USER_RADR_1")
        txtUserRadr2 = .Fields("USER_RADR_2")
        txtUserRadr3 = .Fields("USER_RADR_3")
        txtUserRtel = .Fields("USER_RTEL")
        txtUserRkeitai = .Fields("USER_RKEITAI")
        txtUserRfax = .Fields("USER_RFAX")
        
        '↓INSERT 2010/04/20 SHIBAZAKI
        If Trim(Nz(.Fields("USER_BIRTHDAY"), "")) = "" Then
            txtUserBirthday = .Fields("USER_BIRTHDAY")
        Else
            txtUserBirthday = Left(.Fields("USER_BIRTHDAY"), 4) & "/" & Mid(.Fields("USER_BIRTHDAY"), 5, 2) & "/" & Right(.Fields("USER_BIRTHDAY"), 2)
        End If
        
        '勤務先
        txtUserKname = .Fields("USER_KNAME")
        txtUserKkana = .Fields("USER_KKANA")
        txtUserJob = .Fields("USER_JOB")
        txtUserKpost = .Fields("USER_KPOST")
        txtUserKadr1 = .Fields("USER_KADR_1")
        txtUserKadr2 = .Fields("USER_KADR_2")
        txtUserKadr3 = .Fields("USER_KADR_3")
        txtUserKtel = .Fields("USER_KTEL")
        txtUserKhourStart = .Fields("USER_KHOUR_START")
        txtUserKhourEnd = .Fields("USER_KHOUR_END")
        txtUserNenshu = .Fields("USER_NENSHU")
        
        '緊急連絡先
        '↓INSERT 2010/04/28 SHIBAZAKI
        Set objField = fncContainsField("RCPT_KREN_NAME", rsData)
        If Not (objField Is Nothing) Then
            Set objField = Nothing
            txtUserKrenName = .Fields("RCPT_KREN_NAME")
            txtUserKrenKana = .Fields("RCPT_KREN_KANA")
            txtUserKrenPost = .Fields("RCPT_KREN_YUBINO")
            txtUserKrenAdr1 = .Fields("RCPT_KREN_ADR_1")
            txtUserKrenAdr2 = .Fields("RCPT_KREN_ADR_2")
            txtUserKrenAdr3 = .Fields("RCPT_KREN_ADR_3")
            txtUserKrenTel = .Fields("RCPT_KREN_TEL")
            txtUserKrenKeitai = .Fields("RCPT_KREN_KEITAI")
            txtUserKrenFax = .Fields("RCPT_KREN_FAX")
            txtUserKrenZokugara = .Fields("RCPT_KREN_ZOKUGARA")
        End If
        '↑INSERT 2010/04/28 SHIBAZAKI
        '↓DELETE 2010/04/28 SHIBAZAKI
'        txtUserKrenName = .Fields("RCPT_KREN_NAME")
'        txtUserKrenkana = .Fields("RCPT_KREN_KANA")
'        txtUserKrenPost = .Fields("RCPT_KREN_YUBINO")
'        txtUserKrenAdr1 = .Fields("RCPT_KREN_ADR_1")
'        txtUserKrenAdr2 = .Fields("RCPT_KREN_ADR_2")
'        txtUserKrenAdr3 = .Fields("RCPT_KREN_ADR_3")
'        txtUserKrenTel = .Fields("RCPT_KREN_TEL")
'        txtUserKrenKeitai = .Fields("RCPT_KREN_KEITAI")
'        txtUserKrenFax = .Fields("RCPT_KREN_FAX")
'        txtUserKrenzokugara = .Fields("RCPT_KREN_ZOKUGARA")
        '↑DELETE 2010/04/28 SHIBAZAKI
        '↑INSERT 2010/04/20 SHIBAZAKI
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispUser" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispContainer
'        機能             :コンテナ情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispContainer(rsData As Recordset)
    On Error GoTo ErrorHandler

    Dim strCampaign     As String
    
    With rsData
        txtContainerYcode = Format(.Fields("YCODE"), "000000")
        txtContainerYname = .Fields("YARD_NAME")
        txtContainerYadd = .Fields("YADDR")
        
        'キャンペーンを改行で結合
        strCampaign = ""
        Call subCombineString(Nz(.Fields("YARD_SEV1N")), strCampaign, vbCrLf)
        Call subCombineString(Nz(.Fields("YARD_SEV2N")), strCampaign, vbCrLf)
        Call subCombineString(Nz(.Fields("YARD_SEV3N")), strCampaign, vbCrLf)
        Call subCombineString(Nz(.Fields("YARD_ENDEN")), strCampaign, vbCrLf)
        
        txtContainerSev1n = .Fields("YARD_SEV1N")
        txtContainerSev2n = .Fields("YARD_SEV2N")
        txtContainerSev3n = .Fields("YARD_SEV3N")
        txtContainerEnden = .Fields("YARD_ENDEN")
        
        txtContainerBikon = .Fields("YARD_BIKON")                              'INSERT 20081104 N.MIURA
        
        'キャンペーン満了月数が入っていないか０以下の場合はチェックボックスを非表示にする
'        txtSevExmonth = .Fields("YARD_SEV_EXMONTH")
'        If Nz(.Fields("YARD_SEV_EXMONTH"), 0) <= 0 Or _
'            blnネット予約 = True Then               '2007/09/25 add tajima
'            chkContainerSevExpire.Visible = False
'        Else
'            chkContainerSevExpire.Visible = True
'        End If
        
        '' ＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠
        txtSevExmonth = .Fields("YARD_SEV_EXMONTH")
        If Nz(.Fields("YARD_SEV_EXMONTH"), 0) <= 0 Then
            chkContainerSevExpire.Visible = False
        Else
            chkContainerSevExpire.Visible = True
        End If
        '' ＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠
        
        txtContainerCampaign = strCampaign
        
        txtContainerCno = Format(.Fields("CNTANO"), "000000")
        
' ▼ 2018/09/21 egl del
        'txtContainerStep = .Fields("STEP_NAME")
' ▲ 2018/09/21 egl del
' ▼ 2018/09/21 egl add...2018/10/10 以下範囲のコメント解除しリリースとする
        If strBumonCode = "H" Then
            '部門がコンテナの場合は段区分を表示
            lblContainerStep.Caption = "段区分"
            txtContainerStep = .Fields("STEP_NAME")
        Else
            '部門がコンテナ以外の場合は階数を表示
            lblContainerStep.Caption = "階数"
            txtContainerStep = MSZZ039.fncGetKaiName(Nz(.Fields("CNTA_FLOOR"), 0))
        End If
' ▲ 2018/09/21 egl add
        
        txtContainerType = .Fields("SYU_NAME")
        txtContainerUsage = .Fields("USAGE_NAME")
        cmbContainerUsageCode = .Fields("CNTA_USAGE")
        txtContainerSize = .Fields("CNTA_SIZE")
        
        If Nz(.Fields("SEV_EXPIRE")) = "" Then
            chkContainerSevExpire.VALUE = 0
            txtContainerSevExpire = ""
        Else
            chkContainerSevExpire.VALUE = -1
            txtContainerSevExpire = Left(.Fields("SEV_EXPIRE"), 4) & "/" & Right(.Fields("SEV_EXPIRE"), 2)
        End If
        
        cmbContainerHosyicd = .Fields("HOSYICD")
        Call cmbContainerHosyicd_AfterUpdate    '保証区分によるコントロール表示変更
        '2009/04/01 ADD <S> hirano 保障会社コード追加対応
        strHosyicd_Before = .Fields("HOSYICD")
        cmbContainerHosyoCompany = .Fields("HOSYO_KAISHA_CD")
        Call cmbContainerHosyoCompany_AfterUpdate    '保証会社コードによるコントロール表示変更
        '2009/04/01 ADD <E> hirano
        strHosyo_cd_Before = Nz(.Fields("HOSYO_KAISHA_CD"), "") '2009/04/23 ADD
        txtContainerPrice = .Fields("PRIC_PRICE")
        txtContainerEzapii = .Fields("PRIC_EZAPPI")                             'INSERT 2013/04/10 K.ISHIZAKA
        txtContainerEzapiiCode = .Fields("PRIC_EZAPPI_CODE")                    'INSERT 2013/04/10 K.ISHIZAKA
        txtContainerEzapiiName = .Fields("PRIC_EZAPPI_NAME")                    'INSERT 2013/04/10 K.ISHIZAKA
        txtContainerReason = .Fields("CNTA_REASON")
        txtContainerPriceDiff = .Fields("CNTA_PRICE_DIFF")
        
        '2006/09/06 SHIBAZAKI 営業期間追加
        txtContainerBeginDay = Nz(.Fields("YARD_BEGIN_DAY"), "")
        txtContainerRentEndDay = Nz(.Fields("YARD_RENTEND_DAY"), "")
        If txtContainerBeginDay <> "" & txtContainerRentEndDay <> "" Then
            txtContainerKikan = txtContainerBeginDay & "～" & txtContainerRentEndDay
        End If
    
        txtYardTihoKbn = Nz(.Fields("YARD_TIHO_KBN"))                           'INSERT 2014/06/20 MIYAMOTO
    
        '2007/04/13 tajima add 荷物移動情報追加
        If Nz(.Fields("YOUKT_MOVEKBN"), 0) = P_MOVEKBN_自社理由 Or _
           Nz(.Fields("YOUKT_MOVEKBN"), 0) = P_MOVEKBN_顧客理由 Then
            lbl_荷物移動担当.Visible = True
            cmb_荷物移動担当.Visible = True
            txt_荷物移動担当.Visible = True
            cmb_荷物移動担当.VALUE = .Fields("YOUKT_MOVE_TANTO")
            Call cmb_荷物移動担当_AfterUpdate
        Else
            lbl_荷物移動担当.Visible = False
            cmb_荷物移動担当.Visible = False
            txt_荷物移動担当.Visible = False
        End If
        
        '' 20120411 M.HONDA START
        cmb_スロープ.VALUE = Nz(.Fields("RCPT_SLOPE"), 0)
        Call cmb_スロープ_AfterUpdate
        '' 20120411 M.HONDA END
        
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispContainer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispAppForm
'        機能             :申込書情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispAppForm(rsData As Recordset)
    On Error GoTo ErrorHandler
    
    With rsData
        txtAppFormButu = .Fields("BUTU")
    
'        '2006/04/28 S.SHIBAZAKI 乗用具チェック追加。チェックされている場合、申込書に誓約書添付
'        chkAppFormVehicle.Value = Nz(.Fields("VEHICLE"), 0)
        
        cmbAppFormSendCd = .Fields("SENDCD")
        Call cmbAppFormSendCd_AfterUpdate
        
        txtAppFormLimitDate = Format(.Fields("LIMITDATE"), "yyyy/mm/dd")
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispAppForm" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispBill
'        機能             :請求書情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/05/10 S.SHIBAZAKI 移動予約の場合、移動元が「保証金」の場合のみ保証金を引き継ぐ
'Private Sub subDispBill(rsData As Recordset)
Private Sub subDispBill(rsData As Recordset, blnExist As Boolean, blnMove As Boolean)
    On Error GoTo ErrorHandler
    
    With rsData
        '契約開始日
        txtBillStDate = Format(.Fields("STDATE"), "yyyy/mm/dd")
        
        txtShoStDate = .Fields("RCPT_SYOKISAN_DATE")  '2018/03/16 M.HONDA INS
        
        'I:20081001 N.Imai Start
        '割引適用開始月
        If Nz(.Fields("DC_STDATE")) = "" Then
            txt割引適用開始月 = ""
        Else
            txt割引適用開始月 = Left(.Fields("DC_STDATE"), 4) & "/" & Right(.Fields("DC_STDATE"), 2)
        End If
        
        '割引有効可否
        If lblNewExist.Caption = P_LABEL_既存受付 Then
            chk割引有効可否.VALUE = Nz(.Fields("DC_ENABLE"))
        '▼ 2008/11/19 hirano add コンテナ割引き対象、ラベル追加適用
            chk割引有効可否.Enabled = True
        Else
'            If Nz(.Fields("CNTA_NEBIKI_DISABLE")) = 1 Then
            'コンテナNGの場合 .VALUE = False .Enabled = False
            'コンテナOKの場合 .VALUE = True .Enabled = True
            If blnMove Then
                ' 新規且つ、移動ならば値引きはしない状態にしておく
                chk割引有効可否.VALUE = False
                chk割引有効可否.Enabled = True
            Else
               ' 新規且つ、移動でなけれな値引する態にしておく
                chk割引有効可否.VALUE = Not -1 * Nz(.Fields("CNTA_NEBIKI_DISABLE"), 0)  'False
                chk割引有効可否.Enabled = Not -1 * Nz(.Fields("CNTA_NEBIKI_DISABLE"), 0)
            End If

'            Else
'                'コンテナOKの場合
'                chk割引有効可否.VALUE = True
'                chk割引有効可否.Enabled = True
'            End If
        '▲ 2008/11/19 hirano
        End If
        cmd割引再計算.Enabled = chk割引有効可否.VALUE
        'I:20081001 N.Imai Start
        '▼ 2008/11/19 hirano add 明細を見るボタンの制御追加、コンテナ割引き対象、ラベル追加適用
        cmd明細を見る.Enabled = chk割引有効可否.VALUE
        lblコンテナ割引き対象外.Visible = Nz(.Fields("CNTA_NEBIKI_DISABLE"), 0)
        'txt割引適用開始月.Enabled = Not lblコンテナ割引き対象外.Visible
        '▲ 2008/11/19 hirano
        
        '見積もり請求区分
        cmbBillMituCd = .Fields("MITUCD")
        Call cmbBillMituCd_AfterUpdate
        '初期費用請求方法
        cmbBillSeikyuCd = .Fields("SEIKYUCD")
        Call cmbBillSeikyuCd_AfterUpdate
        '請求期間
        If Nz(.Fields("SEIKYUKIKAN")) = "" Then
            txtBillSeikyuKikan = ""
        Else
            txtBillSeikyuKikan = Left(.Fields("SEIKYUKIKAN"), 4) & "/" & Right(.Fields("SEIKYUKIKAN"), 2)
        End If
        '▼2008/12/04 hirano Add 保証金変更フラグ保持
        '保証金変更フラグ
        chk保証金強制変更.VALUE = Nz(.Fields("SECUKG_FLG"))
        Call YellowControl(txtBillSecukg, Not chk保証金強制変更.VALUE)
        '▲2008/12/04 hirano Add
        '保証金
            txtBillSecukg = .Fields("SECUKG")
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
        Select Case txtMoveKbn.VALUE                                            'INSERT START 2013/04/10 K.ISHIZAKA
        ''2013/10/29 M.HONDA UPD
        ''Case P_MOVEKBN_自社理由
        Case P_MOVEKBN_自社理由, P_MOVEKBN_保障会社変更
        ''2013/10/29 M.HONDA UPD
            opgRentkg.Visible = True
            chkSecukgKase.Visible = False
        ''2013/10/29 M.HONDA UPD
        ''Case P_MOVEKBN_顧客理由, P_MOVEKBN_保障会社変更
        Case P_MOVEKBN_顧客理由
        ''2013/10/29 M.HONDA UPD
            opgRentkg.Visible = False
            '2015/07/24 M.HONDA UPD
            'chkSecukgKase.Visible = True
            chkSecukgKase.Visible = False
            
        Case Else
            opgRentkg.Visible = False
            chkSecukgKase.Visible = False
        End Select
        txtBillRentkgMoto = .Fields("RCPT_RENTKG")
        txtBillEzappiMoto = .Fields("RCPT_EZAPPI")
        txtBillEzappiCodeMoto = .Fields("RCPT_EZAPPI_CODE")
        txtBillEzappiNameMoto = .Fields("RCPT_EZAPPI_NAME")                     'INSERT END   2013/04/10 K.ISHIZAKA
    
        '使用料
        txtBillRentkg = .Fields("RENTKG")
        txtBillFirstkg = .Fields("FIRSTKG")
        txtBillFirstkg_sub = .Fields("PRIC_FPRICE")
        
    
        '毎月雑費
        txtBillEzappiName = .Fields("EZAPPI_NAME")
        '▼ 2008/03/23 chg tajima
        If Nz(.Fields("EZAPPI"), "0") <> "0" Then
            txtBillEzappi = CLng(.Fields("EZAPPI"))
            txtBillEzappi = .Fields("EZAPPI")
        Else
            txtBillEzappi = Null
        End If
        '▲ 2008/03/23 chg tajima
        txtBillEzappiDaily = .Fields("EZAPPI_DAILY")
        
        txtBillEzappiDaily_sub = .Fields("FEZAPPI")  '2019/08/08 M.HONDA INS
        
        txtEzappiCode = .Fields("EZAPPI_CODE")  '' 2012/12/27 M.HONDA INS
        
        txtBillEzappi.Enabled = (Nz(txtBillEzappiName) <> "")
        txtBillEzappiDaily.Enabled = txtBillEzappi.Enabled
        
        '追加毎月雑費１
        cmbBillAddEzappiCode1 = .Fields("ADD_EZAPPI_CODE1")
        Call cmbBillAddEzappiCode1_AfterUpdate
        '▼ 2008/03/23 chg tajima
        If Nz(.Fields("ADD_EZAPPI1"), "0") <> "0" Then
            txtBillAddEzappi1 = CLng(.Fields("ADD_EZAPPI1"))
            txtBillAddEzappi1 = .Fields("ADD_EZAPPI1")
        Else
            txtBillAddEzappi1 = Null
        End If
        '▲ 2008/03/23 chg tajima
        
        txtBillAddEzappiDaily1 = .Fields("ADD_EZAPPI_DAILY1")
        
        '▼ 2007/08/23 add tajima
        '   ﾈｯﾄ割引暫定処置 自動割引は変更不可！
        If Nz(rsData.Fields("YOUKT_GENKBN"), 0) = 1 Then
            Call YellowControl(cmbBillAddEzappiCode1, True)
            Call YellowControl(txtBillAddEzappi1, True)
            Call YellowControl(txtBillAddEzappiDaily1, True)
        End If
        '▲ 2007/08/23 add tajima
        
        '追加毎月雑費２
        cmbBillAddEzappiCode2 = .Fields("ADD_EZAPPI_CODE2")
        Call cmbBillAddEzappiCode2_AfterUpdate
        '▼ 2008/03/23 chg tajima
        If Nz(.Fields("ADD_EZAPPI2"), "0") <> "0" Then
            txtBillAddEzappi2 = CLng(.Fields("ADD_EZAPPI2"))
            txtBillAddEzappi2 = .Fields("ADD_EZAPPI2")
        Else
            txtBillAddEzappi2 = Null
        End If
        '▲ 2008/03/23 chg tajima
        
        txtBillAddEzappiDaily2 = .Fields("ADD_EZAPPI_DAILY2")
            
        '初回雑費
        txtBillFzappiName = .Fields("FZAPPI_NAME")
        If Nz(txtBillFzappiName) = "" Then
            txtBillFzappi = ""
        Else
            txtBillFzappi = .Fields("FZAPPI")
        End If
        txtBillFzappi.Enabled = (Nz(txtBillFzappiName) <> "")
        
        '追加初回雑費
        cmbBillAddFzappiCode1 = .Fields("ADD_FZAPPI_CODE1")
        Call cmbBillAddFzappiCode1_AfterUpdate
        txtBillAddFzappi1 = .Fields("ADD_FZAPPI1")
        
        '▼2015/01/08 add 追加初回雑費２の名称
        If Nz(.Fields("ADD_FZAPPI_CODE2"), 0) = P_ZAPPI_追加カード代 Or Nz(.Fields("ADD_FZAPPI_CODE2"), "0") = "0" Then
            lblBillAddFzappiName2.Caption = "追加カード代"
                int初回雑費2Cd = P_ZAPPI_追加カード代
            cmdBillSCardMnt.Enabled = True
        Else
            lblBillAddFzappiName2.Caption = .Fields("ADD_FZAPPI2_NAME")
            int初回雑費2Cd = .Fields("ADD_FZAPPI_CODE2")
            cmdBillSCardMnt.Enabled = False 'カード追加ボタンは無効化とする
        End If
        '▲2015/01/08 add
        
        '↓INSERT 2008/11/13 SHIBAZAKI
        '移動理由が自社理由or満了更新の場合、追加初回雑費として事務手数料は発生しない。
        If Not blnExist And blnMove Then
            Select Case txtMoveKbn
            '' 2013/06/04 M.HONDA START
            '' 2013/03/26 下記の場合に事務手数料を０円へ
            ''Case P_MOVEKBN_自社理由, P_MOVEKBN_満了更新, P_MOVEKBN_顧客理由, P_MOVEKBN_金額変更, P_MOVEKBN_紹介者金額変更
            '' Case P_MOVEKBN_自社理由, P_MOVEKBN_満了更新
            Case P_MOVEKBN_自社理由, P_MOVEKBN_満了更新, P_MOVEKBN_顧客理由, P_MOVEKBN_金額変更, P_MOVEKBN_紹介者金額変更
                'cmbBillAddFzappiCode1 = ""
            '' 2013/06/04 M.HONDA END
                Call cmbBillAddFzappiCode1_AfterUpdate
            Case Else
                If txtContainerHosyiFrom = P_HOSYIFROM_事務手数料 Then
                    cmbBillAddFzappiCode1 = ""
                    Call cmbBillAddFzappiCode1_AfterUpdate
                End If
            End Select
        End If
        '↑INSERT 2008/11/13 SHIBAZAKI
        
        '2006/06/01 SHIBAZAKI セキュリティカード対応↓
'        cmbBillAddFzappiCode2 = .Fields("ADD_FZAPPI_CODE2")
'        Call cmbBillAddFzappiCode2_AfterUpdate
        txtBillAddFzappi2 = .Fields("ADD_FZAPPI2")
    
        lblSetumei.Caption = ""
        
        '未使用移動時は移動元契約の初回請求金額を入金済み金額とする
        If txtMoveKbn = P_MOVEKBN_未使用移動 Then
            txtBillSumikg.Enabled = True
            txtBillSumikg = Val(Nz(.Fields("SUMI_RENTKG"), 0))
        Else
            txtBillSumikg = 0
            txtBillSumikg.Enabled = False
        End If
        
        txt月額分割引合計 = 0                                                   'I:20081001 N.Imai
        txt初回分割引合計 = 0                                                   'I:20081001 N.Imai
        txt保証金割引 = 0    '2011/06/18 add tajima
        
'▼ 2008/03/23 del tajima SQLで上手く取得するようにしたので以下判断不要也
        '2006/05/10 S.SHIBAZAKI 移動予約の場合、移動元が「保証金」の場合のみ保証金を引き継ぐ
'        If Not blnExist And blnMove Then
'            '新規受付かつ移動予約
'            If txtBillSecukgLabel = P_LABEL_保証委託料 Then
'                '保証委託料の場合は月額料金と保証委託料は同一
'                Call cmbContainerHosyicd_AfterUpdate
'            End If
'        End If
'
'        '▼ 2007/02/07 add tajima 新規受付の場合、保証料＝月額使用料＋毎月雑費とす
'        If blnExist = False Then
'            txtBillSecukg = txtBillRentkg + txtBillEzappiKei
'        End If
'        '▲ 2007/02/07 add
'▲ 2008/03/23 del
        
        '2015/03/24 M.HONDA INS 年払い情報
        chkNEN_KBN = Nz(.Fields("RCPT_NEN_KBN"), 0)
        If Nz(.Fields("RCPT_NEN_KBN"), 0) = 0 Then
            Call YellowControl(cmbJIREI, True)
            Call YellowControl(cmbSEIKI, True)
            Call YellowControl(txtSEIKI, True)
        Else
            Call YellowControl(cmbJIREI, False)
            Call YellowControl(cmbSEIKI, False)
            Call YellowControl(txtSEIKI, False)
       End If
        
        cmbJIREI = .Fields("RCPT_NEN_JIREI")
        Call cmbJIREI_AfterUpdate
        cmbSEIKI = .Fields("RCPT_NEN_SEIKI")
        Call cmbSEIKI_AfterUpdate
        txtSEIKI = .Fields("RCPT_NEN_BIKO")
        
        '▼ 2018/08/06 EGL<S>
        If P_HOSPLN_移行対象 = strNewHosplnIkou Then
            '既存顧客新保証会社プラン移行対象データの場合、下記画面制御を行う
            txtBillSecPlanIkouGdMsg = "共益費に含まれる毎月分割分の"
            txtBillSecukgKei.BackColor = P_ORANGE
            txtBillGoukei.ControlSource = "=[txtBillGetuKei]+[txtBillSyoKei]+[txt保証金割引]"
        End If
        '▲ 2018/08/06 TAKENOUCHI INS<E>
    End With
    
    Call YellowControl(txtBillSecukg, False)  'add 2018/10/02 EGL 保証金額は常に入力可としておく
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispBill" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispKey
'        機能             :鍵情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispKey(rsData As Recordset)
    On Error GoTo ErrorHandler
    
    With rsData
        cmbKeyKagiicd = .Fields("KAGIICD")
        Call cmbKeyKagiicd_AfterUpdate
        
        txtKeyKagiNo = .Fields("KAGINO")
        
        '2006/06/01 SHIBAZAKI セキュリティカード対応
'        txtKeySecurityCard = .Fields("SECURITYCARD")
        
        txtKeyKagiLendnum = .Fields("KAGILENDNUM")
        
        '↓ INSERT 2014/03/11 MIYAMOTO
        strKaijoNo = Nz(.Fields("KAIJONO"))
        txtKaijoNo = strKaijoNo
        
        '↓ INSERT 2021/06/02 EGL
        txtKeyKagiNo.InputMask = ""
        If cmbKeyKagiicd.Column(0) = 16 Then
            strKaijoNo_QR = Nz(.Fields("RCPT_BUIL_KAGI_NO"))
            txtKeyKagiNo = strKaijoNo_QR
            If pvONLET_ACCEI = 0 Then
                '権限無し
                txtKeyKagiNo.InputMask = "Password"
            Else
                txtKeyKagiNo.InputMask = ""
            End If
        End If
        If pvONLET_ACCEI = 0 Then
            '権限無し
            txtKaijoNo.InputMask = "Password"
        Else
            txtKaijoNo.InputMask = ""
        End If
        '↑ INSERT 2021/06/02 EGL
        
        
'↓UPDATE  2016/05/17 SUGIMURA
'       If Nz(cmbKeyKagiicd.Column(2)) = "1" Then
        If MSZZ070_KaiKey(txtYardTihoKbn, Nz(cmbKeyKagiicd.Column(2))) <> "0" Then
'↑UPDATE  2016/05/17 SUGIMURA
            cmbKeyKagiicd.Enabled = False
            cmbKeyKagiicd.Locked = True
            cmbKeyKagiicd.BackColor = 14680063
            txtKeyKagiNo.Enabled = False
            txtKeyKagiNo.Locked = True
            txtKeyKagiNo.BackColor = 14680063
        Else
            cmbKeyKagiicd.Enabled = True
            cmbKeyKagiicd.Locked = False
            cmbKeyKagiicd.BackColor = 16755370
            txtKeyKagiNo.Enabled = True
            txtKeyKagiNo.Locked = False
            txtKeyKagiNo.BackColor = 16777215
        End If
        '↑ INSERT 2014/03/11 MIYAMOTO
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispKey" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispSecurity
'        機能             :保証情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispSecurity(rsData As Recordset)
    On Error GoTo ErrorHandler
    
    With rsData
        txtSecurityHosyD = .Fields("HOSYD")
        txtSecurityHosyB = .Fields("HOSYB")
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispSecurity" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispDlvr
'        機能             :配送サービス情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispDlvr(rsData As Recordset)
    On Error GoTo ErrorHandler
    
    With rsData
        Me.txt_OUT_ADR = .Fields("DLVRT_OUT_ADR")
        Me.txt_PRICE = .Fields("DLVRT_PRICE")
        Me.txt_KYORI = .Fields("DLVRT_KYORI")
' 2015/04/21 N.IMAI START
'        Me.txt_KIBO_DATE1 = .Fields("DLVRT_KIBO_DATE1")
'        Me.txt_KIBO_TIMECD1 = .Fields("TIMECD1")
'        Me.txt_KIBO_DATE2 = .Fields("DLVRT_KIBO_DATE2")
'        Me.txt_KIBO_TIMECD2 = .Fields("TIMECD2")
'        Me.txt_KIBO_DATE3 = .Fields("DLVRT_KIBO_DATE2")
'        Me.txt_KIBO_TIMECD3 = .Fields("TIMECD3")
        Me.txt_KIBO_DATE1 = .Fields("DLVRT_KIBO_DATE1") & Space(1) & .Fields("TIMECD1")
        Me.txt_KIBO_DATE2 = .Fields("DLVRT_KIBO_DATE2") & Space(1) & .Fields("TIMECD2")
        Me.txt_KIBO_DATE3 = .Fields("DLVRT_KIBO_DATE3") & Space(1) & .Fields("TIMECD3")
        Me.txt_BUTU_DANBORU = .Fields("DLVRT_BUTU_DANBORU")
        Me.txt_BUTU_KESU = .Fields("DLVRT_BUTU_KESU")
        Me.txt_BUTU_KAGU = .Fields("DLVRT_BUTU_KAGU")
        Me.txt_BUTU_SEIKATU = .Fields("DLVRT_BUTU_SEIKATU")
        Me.txt_BUTU_REJYA = .Fields("DLVRT_BUTU_REJYA")
        Me.txt_BUTU_ETE = .Fields("DLVRT_BUTU_ETE")
' 2015/04/21 N.IMAI END
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispDlvr" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispSend
'        機能             :鍵送付情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispSend(rsData As Recordset)
    On Error GoTo ErrorHandler
    
    With rsData
        '送付先
        If Nz(.Fields("SENDNAME")) <> "" Then
            txtSendName = .Fields("SENDNAME")
            
            cmbSendHonorific = .Fields("SENDHONORIFIC")
            Call cmbSendHonorific_AfterUpdate
            
            txtSendYubino = .Fields("SENDYUBINO")
            txtSendAdr1 = .Fields("SENDADR1")
            txtSendAdr2 = .Fields("SENDADR2")
            txtSendAdr3 = .Fields("SENDADR3")
            
            txtSendTel = .Fields("SENDTEL")
        Else
            Call cmdUserCopy_Click
        End If
        
        txtSendKeySendDate = .Fields("KEY_SEND_DATE")   '送付日
        txtSendArrivalExpect = .Fields("ARRIVALEXPECT") '到着予定日
        
        cmbSendSpecifyTime = .Fields("SPECIFYTIME")     '配達指定時間
        Call cmbSendSpecifyTime_AfterUpdate
        
        txtSendSlipNo = .Fields("SLIPNO")               '伝票番号
        txtSendDeriveryStatus = .Fields("DERIVERYSTATUS") '配送状況
        
        txtSendArrivalDate = Format(.Fields("ARRIVALDATE"), "yyyy/mm/dd")   '配達完了日
        
        txtSendYamatoDate = .Fields("RCPT_YAMATO_DATE") 'メール発行日   INSERT 2019/03/22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispSend" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispReceipt
'        機能             :受付情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2020/12/01 [EGL] UPD Start
'Private Sub subDispReceipt(rsData As Recordset)
Private Sub subDispReceipt(rsData As Recordset, blnExist As Boolean)
'2020/12/01 [EGL] UPD End
    On Error GoTo ErrorHandler
    
    With rsData
        cmbReceiptUktanto = .Fields("UKTANTO")
        Call cmbReceiptUktanto_AfterUpdate
        
        If Nz(.Fields("UKDATE")) = "" Then
            txtReceiptUkdate = Format(DATE, "yyyy/mm/dd")
            txtReceiptUkdatetime.VALUE = txtReceiptUkdate.VALUE                 'INSERT 2013/06/21 K.ISHIZAKA
        Else
'            txtReceiptUkdate = .Fields("UKDATE")                               'DELETE 2013/06/21 K.ISHIZAKA
            txtReceiptUkdate.VALUE = Format(.Fields("UKDATE"), "yyyy/mm/dd")    'INSERT 2013/06/21 K.ISHIZAKA
            txtReceiptUkdatetime.VALUE = Format(.Fields("UKDATE"), "yyyy/mm/dd hh:nn:ss") 'INSERT 2013/06/21 K.ISHIZAKA
        End If
        lblReceiptUkdate.Caption = IIf(blnネット予約, "受付日時", "受付日")     'INSERT 2013/06/21 K.ISHIZAKA
        txtReceiptUkdate.Visible = Not blnネット予約                            'INSERT 2013/06/21 K.ISHIZAKA
        txtReceiptUkdatetime.Visible = blnネット予約                            'INSERT 2013/06/21 K.ISHIZAKA
        
        txtReceiptCyDate = Format(.Fields("CYDATE"), "yyyy/mm/dd")
            
        cmbReceiptCreason = .Fields("CREASON")
        Call cmbReceiptCreason_AfterUpdate
        
        txtReceiptYukbn = .Fields("YUKBN_NAME")
        txtReceiptYukbnCode = .Fields("YUKBN")
        
        txtReceiptPaydate = .Fields("PAYDATE")
        txtReceiptUknote = .Fields("UKNOTE")
        txtReceiptComment1 = .Fields("COMMENT1")
        txtReceiptComment2 = .Fields("COMMENT2")
        txtReceiptUkkbn = .Fields("UKKBN")
    '▼ 2007/09/25 tajima add
        Me.cmb_利用予定期間 = Nz(.Fields("YOUKT_USE_PERIOD"), "00")
        Me.txt_利用予定期間.VALUE = fncAfterUpdateCommon(Me.cmb_利用予定期間, P_DIVISION_LENGTH, P_CONTROL_数値)
    '▲ 2007/09/25 tajima add
        '2009/04/01 add <S> hirano 予約コメント追加
        txtReceiptYoyakuComment = .Fields("YOUKT_COMMENT")
        '2009/04/01 add <E>
        
        'コンビニならば未請求にする                                             'INSERT START 2010/03/29 K.ISHIZAKA
        If Nz(cmbUserSkbn.VALUE) = P_USER_SKBN_コンビニ Then
            cmbRcptSeiki.VALUE = P_SEIKI_未発行
        Else
            '2015/10/28 M.HONDA UPD
            'cmbRcptSeiki.VALUE = Nz(.Fields("SEIKI"), P_SEIKI_未発行)
            cmbRcptSeiki.VALUE = Nz(.Fields("SEIKI"), P_SEIKI_発行)
        End If
        Call cmbRcptSeiki_AfterUpdate
        cmbRcptKaihi.VALUE = .Fields("KAIHI")
        '2011/06/18 代わりに↓を呼ぶ...Call cmbRcptKaihi_AfterUpdate            'INSERT END   2010/03/29 K.ISHIZAKA
        txtRcptKaihi.VALUE = fncAfterUpdateCommon(cmbRcptKaihi, 1, P_CONTROL_数値)
        
        '2020/12/01 [EGL] UPD Start
        'txtTpoint.VALUE = .Fields("RCPT_TPOINT")        ' 2013/05/21 INS M.HONDA
        If blnExist Then
            txt_ユーザー入力クーポンCD.VALUE = .Fields("RCPT_USER_COUPON_CD")
            txt_ユーザー入力紹介会社名.VALUE = .Fields("RCPT_USER_COUPON_TXT")
            txt_紹介会社名.VALUE = .Fields("COUPON_TXT")
            cmb_クーポンCD.VALUE = .Fields("RCPT_COUPON_CD")
            Call cmb_クーポンCD_AfterUpdate
        End If
        '2020/12/01 [EGL] UPD End
        
        cmb_更新区分 = Nz(.Fields("RCPT_KOUKBN"), 0)     ' 2013/10/16 INS M.HONDA
        Call cmb_更新区分_AfterUpdate                    ' 2013/10/16 INS M.HONDA
    
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispReceipt" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subDispMotoCarg
'        機能             :移動元契約情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispMotoCarg(rsData As Recordset)
    On Error GoTo ErrorHandler
    
    With rsData
        txtMotoKaiUkeDate = .Fields("MOTO_KAIUKEDATE")
        txtMotoKaiTanto = .Fields("MOTO_KAITANTO")
        txtMotoUserCode = .Fields("MOTO_UCODE")
        txtMotoKaicd = .Fields("MOTO_KAICD")
        txtMotoKaiRiyuCd = .Fields("MOTO_KAIRIYUCD")
        txtMotoYcode = .Fields("MOTO_YCODE")
        txtMotoCno = .Fields("MOTO_CARGNO")
        txtMotoKagiiCd = .Fields("MOTO_KAGIICD")                'INSERT 2014/04/17 MIYAMOTO
        txtMotoUkTanto = .Fields("MOTO_UKTANTO")
        txtMotoYardTihoKbn = Nz(.Fields("MOTO_YARD_YARD_TIHO_KBN"))     'INSERT 2014/06/20 MIYAMOTO
        
        '2006/05/10 SHIBAZAKI 契約変更系の受付の場合、契約開始日は移動元契約の初回契約日の翌月以降であること
        txtMotoStDate = .Fields("MOTO_STDATE")
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "subDispSend" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subCombineString
'        機能             :文字列結合
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCombineString(strSource As String, ByRef strDestination As String, varSeparate As Variant)

    If strSource <> "" And strDestination <> "" Then
        strDestination = strDestination & varSeparate
    End If
    strDestination = strDestination & strSource

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncGetConnectString
'        機能             :接続文字列取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetConnectString() As Boolean
    On Error GoTo ErrorHandler
    
    fncGetConnectString = False
    
    '部門コード取得
    strBumonCode = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
    If strBumonCode = "" Then
        'テーブル[CONT_MAST]の設定不正
        Call MSZZ024_M10("DLookup(""CONT_BUMOC"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", "テーブル[CONT_MAST]の設定不正。")
    End If
    
    'SQL-Server名
    strSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME_" & strBumonCode & "'"))
    If strSqlserver = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("DLookup(""SETUT_SETUN"", ""SETU_TABL"", ""SETUT_SETUB='ODBC_DATA_SOURCE_NAME_" & strBumonCode & "')", "テーブル[SETU_TABL]の設定不正。")
    End If
    
    '接続文字列取得
    strConnect = MSZZ007_M00(strBumonCode)
    If strConnect = "" Then
        'テーブル[SETU_TABL]の設定不正
        Call MSZZ024_M10("MSZZ007_M00", "テーブル[SETU_TABL]の設定不正。")
    End If

    fncGetConnectString = True
    
    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncGetConnectString" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncBeforeUpdateCommon
'        機能             :更新前イベント時の処理
'                         :入力チェックを行う
'        IN               :①対象コントロール
'                         :②入力最大桁数
'                         :③定数値「コントロールの種類」参照
'                         :④更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncBeforeUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer, ByRef Cancel As Integer) As Boolean

    Dim strTemp     As String
    Dim blnRet      As Boolean
    
    If Left(objTarget.NAME, 3) = "cmb" Then
        'コンボボックス入力チェック
        fncBeforeUpdateCommon = fncChkComboInput(objTarget, intLength, Cancel)
    ElseIf intKind = P_CONTROL_日付 Then
        '日付チェック
        fncBeforeUpdateCommon = fncDateCheck(objTarget, Cancel)
    ElseIf intKind = P_CONTROL_年月 Then
        '一日の日付にして日付チェック
        fncBeforeUpdateCommon = fncDateCheck(fncFirstDay(objTarget), Cancel)
    ElseIf intKind = P_CONTROL_数値 Or intKind = P_CONTROL_金額 Or intKind = P_CONTROL_頭ゼロ消去 Then
        '数値入力チェック
        fncBeforeUpdateCommon = fncNumericCheck(objTarget, intLength, Cancel)
    ElseIf intKind = P_CONTROL_半角 Then
        '半角入力チェック
        fncBeforeUpdateCommon = fncChkSingleByte(objTarget, intLength, Cancel)
    Else
        '入力桁数チェック
        fncBeforeUpdateCommon = fncChkTextInput(objTarget, intLength, Cancel)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncAfterUpdateCommon
'        機能             :更新後イベント時の処理
'                         :数値項目は０埋めまたは０消しした値を対象コントロールに設定する
'        IN               :①対象コントロール
'                         :②入力最大桁数(数値項目の場合のみ意味を持つ)
'                         :③定数値「コントロールの種類」参照
'        OUT              :日付コントロールは"yyyy/mm/dd"形式
'                         :年月は"yyyy/mm"形式
'                         :コンボボックスは名称
'                         :上記以外は内容をそのまま
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncAfterUpdateCommon(objTarget As Object, intLength As Integer, intKind As Integer) As String

    fncAfterUpdateCommon = ""
    
    '体裁を整える
    Select Case intKind
    Case P_CONTROL_数値 '数値項目は０埋め
        objTarget = Format(objTarget, Mid(P_ZERO_SUPPRESS, 1, intLength))
    Case P_CONTROL_日付
        fncAfterUpdateCommon = Format(VarChangeDate, "yyyy/mm/dd")
    Case P_CONTROL_年月
        fncAfterUpdateCommon = Format(VarChangeDate, "yyyy/mm")
    Case P_CONTROL_頭ゼロ消去
        objTarget = Format(objTarget, "#0")
    Case Else
        fncAfterUpdateCommon = Nz(objTarget)
    End Select
    
    'コンボボックスの場合、名称を返却する
    If Left(objTarget.NAME, 3) = "cmb" Then
        If Nz(objTarget) <> "" And _
           objTarget.ListCount <> 0 _
        Then
            fncAfterUpdateCommon = Nz(objTarget.Column(1))
        End If
    End If
    
    Call subEditMode(True)

End Function

'==============================================================================*
'
'        MODULE_NAME      :subEditMode
'        機能             :変更モード設定
'        IN               :True=変更モードに設定
'                          False=変更モード解除
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subEditMode(blnEditing As Boolean)

    If blnEditing Then
        '編集モードに設定（初回のみ）
        If Not blnEditMode Then
            Me.Caption = Me.Caption & P_CAPTION_EDITING
        End If
    Else
        Me.Caption = strFormCaption
    End If
    blnEditMode = blnEditing
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncDateCheck
'        機能             :日付入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDateCheck(txtTarget As Variant, ByRef Cancel As Integer) As Boolean

    fncDateCheck = True
    
    VarChangeDate = CmDateChecker(txtTarget)
    If IsNull(VarChangeDate) Then
        MsgBox "日付を正しく入力してください。", vbInformation, P_ERROR
        Cancel = True
        SendKeys "{ESC}"
        fncDateCheck = False
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkComboInput
'        機能             :コンボボックス入力チェック
'        IN               :第一引数：対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :0=正常
'                          1=入力規則に反している
'                          2=入力値がリストに存在しない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkComboInput(cmbCheck As ComboBox, intMaxLength As Integer, ByRef Cancel As Integer) As Integer

    Dim strCode     As String
    Dim intRet      As Integer
    Dim intCount    As Integer  '06.01.22 SHIBAZAKI 追加
    Dim strValue    As String   '06.01.22 SHIBAZAKI 追加
    Dim blnFound    As Boolean  '06.01.22 SHIBAZAKI 追加
    
    intRet = 0
    
    strCode = Nz(cmbCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo ChkComboInput_End
    End If
    
    If Not fncNumericCheck(cmbCheck, intMaxLength, Cancel) Then
        intRet = 1
        GoTo ChkComboInput_End
    End If
    
    '入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        '06.01.22 SHIBAZAKI 追加↓
        blnFound = False
        '入力値を０サプレス
        strValue = Format(cmbCheck, Mid(P_ZERO_SUPPRESS, 1, intMaxLength))
        'コンボボックス内にコードが存在してる？
        For intCount = 0 To cmbCheck.ListCount - 1
            'コード値比較
            If cmbCheck.Column(0, intCount) = strValue Then
                blnFound = True
                Exit For
            End If
        Next
        '06.01.22 SHIBAZAKI 追加↑
        
        '06.01.22 SHIBAZAKI 変更
        'If cmbCheck.ListIndex = -1 Then
        If Not blnFound Then
            MsgBox "該当するコードがありません。", vbInformation, P_KAKUNIN
            SendKeys "{ESC}"
            intRet = 2
            GoTo ChkComboInput_End
        End If
    End If
    
ChkComboInput_End:
    If intRet <> 0 Then
        Cancel = True
    End If
    
    fncChkComboInput = intRet

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
Private Function fncChkTextInput(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean
    
    If Nz(txtCheck) = "" Then
        fncChkTextInput = True
        Exit Function
    End If
    
    If CmLenb(txtCheck) > intMaxLength Then
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        fncChkTextInput = False
    Else
        fncChkTextInput = True
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncChkSingleByte
'        機能             :半角文字入力チェック
'        IN               :第一引数:対象コンボボックス
'                         :第二引数：入力可能最大桁数
'                         :第三引数：更新後処理のキャンセル
'        OUT              :True:OK False:NG
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkSingleByte(txtCheck As TextBox, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean
    
    fncChkSingleByte = False
    
    If Nz(txtCheck) = "" Then
        fncChkSingleByte = True
        Exit Function
    End If
    
    '桁数チェック
    If Not fncChkTextInput(txtCheck, intMaxLength, Cancel) Then
        Exit Function
    End If
    
    '半角文字チェック
    If CmChkCher(Nz(txtCheck.VALUE), "A") Then
        fncChkSingleByte = True
    Else
        MsgBox "半角英数字で入力してください。", vbExclamation, P_KAKUNIN
        Cancel = True
        SendKeys "{ESC}"
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncNumericCheck
'        機能             :数値入力チェック
'        IN               :対象テキスト
'                          更新後処理のキャンセル
'        OUT              :True=正常 False=入力エラー
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncNumericCheck(objTarget As Object, intMaxLength As Integer, ByRef Cancel As Integer) As Boolean

    Dim strNumber       As String
    Dim lngCompare      As Long
    
    fncNumericCheck = False
    
    strNumber = Nz(objTarget, "")
    '空白なら入力チェックしない
    If strNumber = "" Then
        fncNumericCheck = True
        Exit Function
    End If
    
    lngCompare = Val(Mid(P_ALL_NINE, 1, intMaxLength))
    If Val(Nz(objTarget, 0)) > lngCompare Then
        '桁数ｴﾗｰ
        MsgBox "桁数オーバーしています。", vbInformation, P_ERROR
        Cancel = True
        Exit Function
    End If

    fncNumericCheck = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :txtBillFzappi_AfterUpdate
'        機能             :txtBillFzappi_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFzappi_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillFzappi, P_PRICE_LENGTH, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFzappi_BeforeUpdate
'        機能             :txtBillFzappi_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFzappi_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillFzappi, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFzappi_GotFocus
'        機能             :txtBillFzappi_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFzappi_GotFocus()

    txtBillFzappi.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFzappi_KeyPress
'        機能             :txtBillFzappi_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFzappi_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillFzappi, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillFzappi_LostFocus
'        機能             :txtBillFzappi_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillFzappi_LostFocus()

    txtBillFzappi.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi1_AfterUpdate
'        機能             :txtBillAddFzappi1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi1_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillAddFzappi1, P_PRICE_LENGTH, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi1_BeforeUpdate
'        機能             :txtBillAddFzappi1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi1_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillAddFzappi1, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi1_GotFocus
'        機能             :txtBillAddFzappi1_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi1_GotFocus()

    txtBillAddFzappi1.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi1_KeyPress
'        機能             :txtBillAddFzappi1_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi1_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillAddFzappi1, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi1_LostFocus
'        機能             :txtBillAddFzappi1_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi1_LostFocus()

    txtBillAddFzappi1.Format = P_FORMAT_STANDARD

End Sub


'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi2_AfterUpdate
'        機能             :txtBillAddFzappi2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi2_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillAddFzappi2, P_PRICE_LENGTH, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi2_BeforeUpdate
'        機能             :txtBillAddFzappi2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi2_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillAddFzappi2, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi2_GotFocus
'        機能             :txtBillAddFzappi2_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi2_GotFocus()

    txtBillAddFzappi2.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi2_KeyPress
'        機能             :txtBillAddFzappi2_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi2_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillAddFzappi2, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddFzappi2_LostFocus
'        機能             :txtBillAddFzappi2_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddFzappi2_LostFocus()

    txtBillAddFzappi2.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi1_AfterUpdate
'        機能             :txtBillAddEzappi1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi1_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillAddEzappi1, P_PRICE_LENGTH, P_CONTROL_金額)
    If Nz(txtBillAddEzappiDaily1) = "" Then
        txtBillAddEzappiDaily1 = fncDailyRate(txtBillAddEzappi1)
    End If
    '▼2008/12/04 Mod hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。
    'txtBillSecukg = txtBillRentkg + txtBillEzappiKei  '2007/02/07 add tajima 保証料は月額の金額全て
    '▼2008/12/20 Add hirano 保証料は、毎月額から月額割引をマイナスする。(chk保証金強制変更がチェックされていない場合)
    'txtBillSecukg = txtBillRentkg + txtBillEzappiKei + (txt月額分割引合計 / IIf(txtBillPremthSum = 0, 1, txtBillPremthSum))
    If Not chk保証金強制変更.VALUE Then
        'del 2018/10/02 EGL -->If chk割引有効可否.VALUE Then txtBillSecukg = txtBillMonthCharge + txt月額分割引
        Call subCompSecukgKase                                                  'INSERT 2013/04/10 K.ISHIZAKA
    End If
    '▲2008/12/20 Mod hirano
    '▲2008/12/04 Mod hirano
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi1_BeforeUpdate
'        機能             :txtBillAddEzappi1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi1_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillAddEzappi1, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi1_GotFocus
'        機能             :txtBillAddEzappi1_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi1_GotFocus()

    txtBillAddEzappi1.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi1_KeyPress
'        機能             :txtBillAddEzappi1_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi1_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillAddEzappi1, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillAddEzappi1_LostFocus
'        機能             :txtBillAddEzappi1_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillAddEzappi1_LostFocus()

    txtBillAddEzappi1.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillRentkg_AfterUpdate
'        機能             :txtBillRentkg_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillRentkg_AfterUpdate()

    Call fncAfterUpdateCommon(txtBillRentkg, P_PRICE_LENGTH, P_CONTROL_金額)

    '保障委託料は月額使用料と同じ
'    If txtBillSecukgLabel = P_LABEL_保証委託料 Then        'DELETE 2008/11/13 SHIBAZAKI
    If txtContainerHosyiFrom = P_HOSYIFROM_保証委託 Then    'INSERT 2008/11/13 SHIBAZAKI
        '新旧プラン判定
        If FVS400_fncChkHoshoNewOld(strBumonCode, cmbContainerHosyoCompany) = False Then '2017/12/20 ADD
            '▼2008/12/04 Mod hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。
            'txtBillSecukg = txtBillRentkg + txtBillEzappiKei  '2007/02/07 add tajima 保証料は月額の金額全て
            '▼2008/12/20 Add hirano 保証料は、毎月額から月額割引をマイナスする。(chk保証金強制変更がチェックされていない場合)
            'txtBillSecukg = txtBillRentkg + txtBillEzappiKei + txt月額分割引合計
            If Not chk保証金強制変更.VALUE Then
                'del 2018/10/02 EGL -->If chk割引有効可否.VALUE Then txtBillSecukg = txtBillMonthCharge + txt月額分割引
                Call subCompSecukgKase                                              'INSERT 2013/04/10 K.ISHIZAKA
            End If
            '▲2008/12/20 Mod hirano
            '▲2008/12/04 Mod hirano
        End If
    End If '2017/12/20 ADD
    
    Call subDispSecukgMessage                                                   'INSERT 2013/04/10 K.ISHIZAKA
End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillRentkg_BeforeUpdate
'        機能             :txtBillRentkg_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillRentkg_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtBillRentkg, P_PRICE_LENGTH, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillRentkg_GotFocus
'        機能             :txtBillRentkg_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillRentkg_GotFocus()

    txtBillRentkg.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillRentkg_KeyPress
'        機能             :txtBillRentkg_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillRentkg_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtBillRentkg, KeyAscii, P_PRICE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtBillRentkg_LostFocus
'        機能             :txtBillRentkg_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtBillRentkg_LostFocus()

    txtBillRentkg.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiNo_AfterUpdate
'        機能             :txtKeyKagiNo_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiNo_AfterUpdate()

    Call fncAfterUpdateCommon(txtKeyKagiNo, P_KEYNO_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiNo_BeforeUpdate
'        機能             :txtKeyKagiNo_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiNo_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtKeyKagiNo, P_KEYNO_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiNo_KeyPress
'        機能             :txtKeyKagiNo_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiNo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char(txtKeyKagiNo, KeyAscii, P_KEYNO_LENGTH)

End Sub

'2006/06/01 SHIBAZAKI セキュリティカード対応↓
''==============================================================================*
''
''        MODULE_NAME      :txtKeySecurityCard_AfterUpdate
''        機能             :txtKeySecurityCard_AfterUpdate イベント
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txtKeySecurityCard_AfterUpdate()
'
'    Call fncAfterUpdateCommon(txtKeySecurityCard, P_KEYNO_LENGTH, P_CONTROL_半角)
'
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :txtKeySecurityCard_BeforeUpdate
''        機能             :txtKeySecurityCard_BeforeUpdate イベント
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txtKeySecurityCard_BeforeUpdate(Cancel As Integer)
'
'    Call fncBeforeUpdateCommon(txtKeySecurityCard, P_KEYNO_LENGTH, P_CONTROL_半角, Cancel)
'
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :txtKeySecurityCard_KeyPress
''        機能             :txtKeySecurityCard_KeyPress イベント
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub txtKeySecurityCard_KeyPress(KeyAscii As Integer)
'
'    Call CmKeyPress_For_Char(txtKeySecurityCard, KeyAscii, P_KEYNO_LENGTH)
'
'End Sub
'2006/06/01 SHIBAZAKI セキュリティカード対応↑

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiLendnum_AfterUpdate
'        機能             :txtKeyKagiLendnum_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiLendnum_AfterUpdate()

    Call fncAfterUpdateCommon(txtKeyKagiLendnum, P_LENDNUM_LENGTH, P_CONTROL_頭ゼロ消去)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiLendnum_BeforeUpdate
'        機能             :txtKeyKagiLendnum_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiLendnum_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtKeyKagiLendnum, P_LENDNUM_LENGTH, P_CONTROL_頭ゼロ消去, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtKeyKagiLendnum_KeyPress
'        機能             :txtKeyKagiLendnum_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtKeyKagiLendnum_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtKeyKagiLendnum, KeyAscii, P_LENDNUM_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptCyDate_AfterUpdate
'        機能             :txtReceiptCyDate_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptCyDate_AfterUpdate()

    txtReceiptCyDate = fncAfterUpdateCommon(txtReceiptCyDate, P_DATE_LENGTH, P_CONTROL_日付)
    Call subDailyRate

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptCyDate_BeforeUpdate
'        機能             :txtReceiptCyDate_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptCyDate_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtReceiptCyDate, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptCyDate_GotFocus
'        機能             :txtReceiptCyDate_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptCyDate_GotFocus()

    txtReceiptCyDate = Format(txtReceiptCyDate, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptCyDate_KeyPress
'        機能             :txtReceiptCyDate_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptCyDate_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtReceiptCyDate, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtReceiptCyDate_LostFocus
'        機能             :txtReceiptCyDate_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtReceiptCyDate_LostFocus()

    txtReceiptCyDate = Format(txtReceiptCyDate, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyD_AfterUpdate
'        機能             :txtSecurityHosyD_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyD_AfterUpdate()

    txtSecurityHosyD = fncAfterUpdateCommon(txtSecurityHosyD, P_DATE_LENGTH, P_CONTROL_日付)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyD_BeforeUpdate
'        機能             :txtSecurityHosyD_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyD_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSecurityHosyD, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyD_GotFocus
'        機能             :txtSecurityHosyD_GotFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyD_GotFocus()

    txtSecurityHosyD = Format(txtSecurityHosyD, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyD_KeyPress
'        機能             :txtSecurityHosyD_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyD_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtSecurityHosyD, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyD_LostFocus
'        機能             :txtSecurityHosyD_LostFocus イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyD_LostFocus()

    txtSecurityHosyD = Format(txtSecurityHosyD, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyB_AfterUpdate
'        機能             :txtSecurityHosyB_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyB_AfterUpdate()

    Call fncAfterUpdateCommon(txtSecurityHosyB, P_HOSYOUNO_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyB_BeforeUpdate
'        機能             :txtSecurityHosyB_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyB_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSecurityHosyB, P_HOSYOUNO_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSecurityHosyB_KeyPress
'        機能             :txtSecurityHosyB_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSecurityHosyB_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char(txtSecurityHosyB, KeyAscii, P_HOSYOUNO_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendName_AfterUpdate
'        機能             :txtSendName_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendName_AfterUpdate()

    Call fncAfterUpdateCommon(txtSendName, P_NAME_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendName_BeforeUpdate
'        機能             :txtSendName_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendName_BeforeUpdate(Cancel As Integer)

    '2006/04/28 S.SHIBAZAKI↓
    ' 鍵送付先名称桁数と住所桁数（１～３全ての桁数合計）をB2仕様に合わせてチェックする
'    Call fncBeforeUpdateCommon(txtSendName, P_NAME_LENGTH, P_CONTROL_テキスト, Cancel)
    Call fncBeforeUpdateCommon(txtSendName, P_KEY_SEND_NAME, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendSlipNo_AfterUpdate
'        機能             :txtSendSlipNo_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendSlipNo_AfterUpdate()

    Call fncAfterUpdateCommon(txtSendSlipNo, P_SLIPNO_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendSlipNo_BeforeUpdate
'        機能             :txtSendSlipNo_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendSlipNo_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSendSlipNo, P_SLIPNO_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendSlipNo_KeyPress
'        機能             :txtSendSlipNo_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendSlipNo_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char(txtSendSlipNo, KeyAscii, P_SLIPNO_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendTel_AfterUpdate
'        機能             :txtSendTel_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendTel_AfterUpdate()

    Call fncAfterUpdateCommon(txtSendTel, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendTel_BeforeUpdate
'        機能             :txtSendTel_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendTel_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSendTel, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendTel_KeyPress
'        機能             :txtSendTel_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendTel_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtSendTel, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserAdr1_AfterUpdate
'        機能             :txtUserAdr1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserAdr1_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserAdr1, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserBirthday_AfterUpdate
'        機能             :txtUserBirthday_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserBirthday_AfterUpdate()

    txtUserBirthday = fncAfterUpdateCommon(txtUserBirthday, P_DATE_LENGTH, P_CONTROL_日付)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserBirthday_BeforeUpdate
'        機能             :txtUserBirthday_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserBirthday_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserBirthday, P_DATE_LENGTH, P_CONTROL_日付, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserBirthday_GotFocus
'        機能             :txtUserBirthday_GotFocus イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserBirthday_GotFocus()

    txtUserBirthday = Format(txtUserBirthday, "yy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserBirthday_KeyPress
'        機能             :txtUserBirthday_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserBirthday_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txtUserBirthday, KeyAscii, P_DATE_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserBirthday_LostFocus
'        機能             :txtUserBirthday_LostFocus イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserBirthday_LostFocus()

    txtUserBirthday = Format(txtUserBirthday, "yyyy/mm/dd")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserJob_BeforeUpdate
'        機能             :txtUserJob_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserJob_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserJob, 50, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserJob_AfterUpdate
'        機能             :txtUserJob_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserJob_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserJob, 50, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKadr1_BeforeUpdate
'        機能             :txtUserKadr1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKadr1_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserKadr1, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKadr1_AfterUpdate
'        機能             :txtUserKadr1_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKadr1_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKadr1, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKadr2_BeforeUpdate
'        機能             :txtUserKadr2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKadr2_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserKadr2, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKadr2_AfterUpdate
'        機能             :txtUserKadr2_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKadr2_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKadr2, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKadr3_BeforeUpdate
'        機能             :txtUserKadr3_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKadr3_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserKadr3, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKadr3_AfterUpdate
'        機能             :txtUserKadr3_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKadr3_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKadr3, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenAdr1_BeforeUpdate
'        機能             :txtUserKrenAdr1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenAdr1_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserKrenAdr1, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenAdr1_AfterUpdate
'        機能             :txtUserKrenAdr1_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenAdr1_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenAdr1, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenAdr2_BeforeUpdate
'        機能             :txtUserKrenAdr2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenAdr2_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserKrenAdr2, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenAdr2_AfterUpdate
'        機能             :txtUserKrenAdr2_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenAdr2_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenAdr2, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenAdr3_BeforeUpdate
'        機能             :txtUserKrenAdr3_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenAdr3_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserKrenAdr3, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenAdr3_AfterUpdate
'        機能             :txtUserKrenAdr3_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenAdr3_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenAdr3, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKhourStart_AfterUpdate
'        機能             :txtUserKhourStart_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKhourStart_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKhourStart, 5, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKhourStart_BeforeUpdate
'        機能             :txtUserKhourStart_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKhourStart_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKhourStart, 5, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKhourEnd_AfterUpdate
'        機能             :txtUserKhourEnd_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKhourEnd_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKhourEnd, 5, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKhourEnd_BeforeUpdate
'        機能             :txtUserKhourEnd_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKhourEnd_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKhourEnd, 5, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKname_BeforeUpdate
'        機能             :txtUserKname_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKname_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKname, 50, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKname_AfterUpdate
'        機能             :txtUserKname_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKname_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKname, 50, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenName_BeforeUpdate
'        機能             :txtUserKrenName_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenName_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKrenName, 50, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenName_AfterUpdate
'        機能             :txtUserKrenName_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenName_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenName, 50, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKkana_BeforeUpdate
'        機能             :txtUserKkana_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKkana_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKkana, 40, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKkana_AfterUpdate
'        機能             :txtUserKkana_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKkana_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKkana, 40, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenKana_BeforeUpdate
'        機能             :txtUserKrenKana_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenKana_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKrenKana, 40, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenKana_AfterUpdate
'        機能             :txtUserKrenKana_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenKana_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenKana, 40, P_CONTROL_テキスト)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txtUserKpost_AfterUpdate
'        機能             :txtUserKpost_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKpost_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKpost, P_POST_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKpost_BeforeUpdate
'        機能             :txtUserKpost_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKpost_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKpost, P_POST_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKpost_KeyPress
'        機能             :txtUserKpost_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKpost_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserKpost, KeyAscii, P_POST_LENGTH)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenPost_AfterUpdate
'        機能             :txtUserKrenPost_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenPost_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenPost, P_POST_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenPost_BeforeUpdate
'        機能             :txtUserKrenPost_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenPost_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKrenPost, P_POST_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenPost_KeyPress
'        機能             :txtUserKrenPost_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenPost_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserKrenPost, KeyAscii, P_POST_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenTel_AfterUpdate
'        機能             :txtUserKrenTel_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenTel_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenTel, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenTel_BeforeUpdate
'        機能             :txtUserKrenTel_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenTel_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKrenTel, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenTel_KeyPress
'        機能             :txtUserKrenTel_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenTel_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserKrenTel, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenKeitai_AfterUpdate
'        機能             :txtUserKrenKeitai_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenKeitai_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenKeitai, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenKeitai_BeforeUpdate
'        機能             :txtUserKrenKeitai_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenKeitai_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKrenKeitai, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenKeitai_KeyPress
'        機能             :txtUserKrenKeitai_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenKeitai_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserKrenKeitai, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenFax_AfterUpdate
'        機能             :txtUserKrenFax_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenFax_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenFax, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenFax_BeforeUpdate
'        機能             :txtUserKrenFax_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenFax_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKrenFax, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenFax_KeyPress
'        機能             :txtUserKrenFax_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenFax_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserKrenFax, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenZokugara_BeforeUpdate
'        機能             :txtUserKrenZokugara_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenZokugara_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKrenZokugara, 20, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKrenZokugara_AfterUpdate
'        機能             :txtUserKrenZokugara_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKrenZokugara_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKrenZokugara, 20, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKtel_AfterUpdate
'        機能             :txtUserKtel_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKtel_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKtel, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKtel_BeforeUpdate
'        機能             :txtUserKtel_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKtel_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKtel, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKtel_KeyPress
'        機能             :txtUserKtel_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKtel_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserKtel, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserNenshu_AfterUpdate
'        機能             :txtUserNenshu_AfterUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserNenshu_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserNenshu, 9, P_CONTROL_金額)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserNenshu_BeforeUpdate
'        機能             :txtUserNenshu_BeforeUpdate イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserNenshu_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserNenshu, 9, P_CONTROL_金額, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserNenshu_GotFocus
'        機能             :txtUserNenshu_GotFocus イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserNenshu_GotFocus()

    txtUserNenshu.Format = P_FORMAT_FIXED

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserNenshu_KeyPress
'        機能             :txtUserNenshu_KeyPress イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserNenshu_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtUserNenshu, KeyAscii, 9)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserNenshu_LostFocus
'        機能             :txtUserNenshu_LostFocus イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserNenshu_LostFocus()

    txtUserNenshu.Format = P_FORMAT_STANDARD

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRadr1_BeforeUpdate
'        機能             :txtUserRadr1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRadr1_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserRadr1, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRadr1_AfterUpdate
'        機能             :txtUserRadr1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRadr1_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRadr1, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRadr2_BeforeUpdate
'        機能             :txtUserRadr2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRadr2_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserRadr2, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRadr2_AfterUpdate
'        機能             :txtUserRadr2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRadr2_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRadr2, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRadr3_BeforeUpdate
'        機能             :txtUserRadr3_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRadr3_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserRadr3, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRadr3_AfterUpdate
'        機能             :txtUserRadr3_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRadr3_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRadr3, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserAdr1_BeforeUpdate
'        機能             :txtUserAdr1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserAdr1_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserAdr1, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserAdr2_AfterUpdate
'        機能             :txtUserAdr2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserAdr2_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserAdr2, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserAdr2_BeforeUpdate
'        機能             :txtUserAdr2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserAdr2_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserAdr2, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserAdr3_AfterUpdate
'        機能             :txtUserAdr3_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserAdr3_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserAdr3, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserAdr3_BeforeUpdate
'        機能             :txtUserAdr3_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserAdr3_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserAdr3, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendAdr1_AfterUpdate
'        機能             :txtSendAdr1_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendAdr1_AfterUpdate()

    Call fncAfterUpdateCommon(txtSendAdr1, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendAdr1_BeforeUpdate
'        機能             :txtSendAdr1_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendAdr1_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtSendAdr1, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendAdr2_AfterUpdate
'        機能             :txtSendAdr2_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendAdr2_AfterUpdate()

    Call fncAfterUpdateCommon(txtSendAdr2, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendAdr2_BeforeUpdate
'        機能             :txtSendAdr2_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendAdr2_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtSendAdr2, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendAdr3_AfterUpdate
'        機能             :txtSendAdr3_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendAdr3_AfterUpdate()

    Call fncAfterUpdateCommon(txtSendAdr3, P_ADR_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendAdr3_BeforeUpdate
'        機能             :txtSendAdr3_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendAdr3_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtSendAdr3, P_ADR_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserBiko_AfterUpdate
'        機能             :txtUserBiko_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserBiko_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserBiko, P_NOTE_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserBiko_BeforeUpdate
'        機能             :txtUserBiko_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserBiko_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserBiko, P_NOTE_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserExistKana_AfterUpdate
'        機能             :txtUserExistKana_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserExistKana_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserExistKana, P_KANA_LENGTH, P_CONTROL_テキスト)

    cmbUserCode = ""
    cmbUserCode.Requery

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserExistKana_BeforeUpdate
'        機能             :txtUserExistKana_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserExistKana_BeforeUpdate(Cancel As Integer)
    
    Call fncBeforeUpdateCommon(txtUserExistKana, P_KANA_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserExistKana_KeyPress
'        機能             :txtUserExistKana_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserExistKana_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Char(txtUserExistKana, KeyAscii, P_KANA_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKana_AfterUpdate
'        機能             :txtUserKana_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKana_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKana, P_KANA_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRkana_BeforeUpdate
'        機能             :txtUserRkana_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRkana_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserRkana, P_KANA_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRkana_AfterUpdate
'        機能             :txtUserRkana_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRkana_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRkana, P_KANA_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKana_BeforeUpdate
'        機能             :txtUserKana_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKana_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKana, P_KANA_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserMail_AfterUpdate
'        機能             :txtUserMail_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserMail_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserMail, P_MAIL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserMail_BeforeUpdate
'        機能             :txtUserMail_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserMail_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserMail, P_MAIL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserMail_KeyDown
'        機能             :txtUserMail_KeyDown イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserMail_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim intShiftDown As Integer
    Dim intAltDown As Integer
    Dim intCtrlDown As Integer
    
    ' ビット マスクを使って、押されたキーを調べます。
    intShiftDown = (Shift And acShiftMask) > 0
    intAltDown = (Shift And acAltMask) > 0
    
    If intShiftDown Or intAltDown Then
        Exit Sub
    End If
    
    '↓DELETE 2010/04/20 SHIBAZAKI
'    On Error Resume Next
'    'なぜか変更後にタブかエンターを押すとカーソルが意図したとおりに動かないので制御
'    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
'        KeyCode = 0
'        cmbContainerHosyicd.SetFocus
'    End If
    '↑DELETE 2010/04/20 SHIBAZAKI

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserSday_AfterUpdate
'        機能             :txtUserSday_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserSday_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserSday, P_DAY_LENGTH, P_CONTROL_頭ゼロ消去)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserSday_BeforeUpdate
'        機能             :txtUserSday_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserSday_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserSday, P_DAY_LENGTH, P_CONTROL_頭ゼロ消去, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserSday_KeyPress
'        機能             :txtUserSday_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserSday_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(txtUserSday, KeyAscii, P_DAY_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserTaka_AfterUpdate
'        機能             :txtUserTaka_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserTaka_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserTaka, P_KANA_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserTaka_BeforeUpdate
'        機能             :txtUserTaka_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserTaka_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserTaka, P_KANA_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserName_AfterUpdate
'        機能             :txtUserName_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserName_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserName, P_NAME_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserName_BeforeUpdate
'        機能             :txtUserName_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserName_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserName, P_NAME_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRname_BeforeUpdate
'        機能             :txtUserRname_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRname_BeforeUpdate(Cancel As Integer)

    '2006/05/10 SHIBAZAKI 書類送付先名の桁を３６⇒１８に変更
    'Call fncBeforeUpdateCommon(txtUserRname, P_NAME_LENGTH, P_CONTROL_テキスト, Cancel)
    Call fncBeforeUpdateCommon(txtUserRname, P_DAI_NAME_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRname_AfterUpdate
'        機能             :txtUserRname_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRname_AfterUpdate()

    '2006/05/10 SHIBAZAKI 書類送付先名の桁を３６⇒１８に変更
    'Call fncAfterUpdateCommon(txtUserRname, P_NAME_LENGTH, P_CONTROL_テキスト)
    Call fncAfterUpdateCommon(txtUserRname, P_DAI_NAME_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserTanm_AfterUpdate
'        機能             :txtUserTanm_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserTanm_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserTanm, P_DAI_NAME_LENGTH, P_CONTROL_テキスト)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserTanm_BeforeUpdate
'        機能             :txtUserTanm_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserTanm_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserTanm, P_DAI_NAME_LENGTH, P_CONTROL_テキスト, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserTel_AfterUpdate
'        機能             :txtUserTel_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserTel_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserTel, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserTel_BeforeUpdate
'        機能             :txtUserTel_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserTel_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserTel, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserTel_KeyPress
'        機能             :txtUserTel_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserTel_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserTel, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRtel_AfterUpdate
'        機能             :txtUserRtel_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRtel_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRtel, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRtel_BeforeUpdate
'        機能             :txtUserRtel_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRtel_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserRtel, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRtel_KeyPress
'        機能             :txtUserRtel_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRtel_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserRtel, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRfax_AfterUpdate
'        機能             :txtUserRfax_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRfax_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRfax, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRfax_BeforeUpdate
'        機能             :txtUserRfax_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRfax_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserRfax, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRfax_KeyDown
'        機能             :txtUserRfax_KeyDown イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRfax_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intShiftDown As Integer
    Dim intAltDown As Integer
    Dim intCtrlDown As Integer
    
    ' ビット マスクを使って、押されたキーを調べます。
    intShiftDown = (Shift And acShiftMask) > 0
    intAltDown = (Shift And acAltMask) > 0
    
    If intShiftDown Or intAltDown Then
        Exit Sub
    End If
    
    On Error Resume Next
    'なぜか変更後にタブかエンターを押すとカーソルが意図したとおりに動かないので制御
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmbContainerHosyicd.SetFocus
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRfax_KeyPress
'        機能             :txtUserRfax_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRfax_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserRfax, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRkeitai_AfterUpdate
'        機能             :txtUserRkeitai_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRkeitai_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRkeitai, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRkeitai_BeforeUpdate
'        機能             :txtUserRkeitai_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRkeitai_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserRkeitai, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRkeitai_KeyPress
'        機能             :txtUserRkeitai_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRkeitai_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserRkeitai, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKeitai_AfterUpdate
'        機能             :txtUserKeitai_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKeitai_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserKeitai, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKeitai_BeforeUpdate
'        機能             :txtUserKeitai_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKeitai_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserKeitai, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserKeitai_KeyPress
'        機能             :txtUserKeitai_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserKeitai_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserKeitai, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserFax_AfterUpdate
'        機能             :txtUserFax_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserFax_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserFax, P_TEL_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserFax_BeforeUpdate
'        機能             :txtUserFax_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserFax_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserFax, P_TEL_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserFax_KeyPress
'        機能             :txtUserFax_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserFax_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserFax, KeyAscii, P_TEL_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserYubino_AfterUpdate
'        機能             :txtUserYubino_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserYubino_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserYubino, P_POST_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserYubino_BeforeUpdate
'        機能             :txtUserYubino_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserYubino_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserYubino, P_POST_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserYubino_KeyPress
'        機能             :txtUserYubino_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserYubino_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserYubino, KeyAscii, P_POST_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendYubino_AfterUpdate
'        機能             :txtSendYubino_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYubino_AfterUpdate()

    Call fncAfterUpdateCommon(txtSendYubino, P_POST_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendYubino_BeforeUpdate
'        機能             :txtSendYubino_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYubino_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtSendYubino, P_POST_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtSendYubino_KeyPress
'        機能             :txtSendYubino_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtSendYubino_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtSendYubino, KeyAscii, P_POST_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRpost_AfterUpdate
'        機能             :txtUserRpost_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRpost_AfterUpdate()

    Call fncAfterUpdateCommon(txtUserRpost, P_POST_LENGTH, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRpost_BeforeUpdate
'        機能             :txtUserRpost_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRpost_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txtUserRpost, P_POST_LENGTH, P_CONTROL_半角, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txtUserRpost_KeyPress
'        機能             :txtUserRpost_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txtUserRpost_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Tel(txtUserRpost, KeyAscii, P_POST_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subStDateDefault
'        機能             :契約開始日デフォルト値設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subStDateDefault()
    
    '2006/09/06 SHIBAZAKI 移動予約時は希望開始日が空白の場合がある
    If Nz(txtBillStDate) = "" Then
        Exit Sub
    End If
    
    '2006/09/06 SHIBAZAKI 契約開始日は予約トラン・希望開始日を初期値とする
    lblSetumei.Caption = "契約開始日と起算日の初期値は「希望開始日」になります。"
'    Dim datUkdate8After     As Date
'
'    '受付日の八日後
'    datUkdate8After = DateAdd("d", 8, txtReceiptUkdate)
'
'    '受付日の八日後は月が変わっているか？
'    If Month(datUkdate8After) = Month(txtReceiptUkdate) Then
'        '受付日＋８日が翌月にならない場合、契約日と起算日のデフォルトは「受付日＋８日」
'        txtBillStDate = Format(datUkdate8After, "yyyy/mm/dd")
'        lblSetumei.Caption = "受付日＋８日が翌月にならない場合、契約開始日と起算日の初期値は「受付日＋８日」になります。"
'    Else
'        '受付日＋８日が翌月になる場合、契約日と起算日のデフォルトは「翌月一日」
'        txtBillStDate = Format(datUkdate8After, "yyyy/mm/01")
'        lblSetumei.Caption = "受付日＋８日が翌月になる場合、契約開始日と起算日の初期値は「翌月一日」になります。"
'    End If
    
    Call subDailyRate(False)    '起算日/日割日数表示
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetKisanDate
'        機能             :起算日/日割り
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDailyRate(Optional blnDispSetumei As Boolean = True)

    Dim strSetumei  As String
    
    txtBillKisanDate = ""
    txtBillDailyRate = ""

    If Nz(txtBillStDate) = "" Then
        '契約開始日が空の場合、計算しない
        Exit Sub
    End If
    
    '起算日の設定
    If Nz(txtReceiptCyDate) = "" Then
        '解約日が入っていない場合,契約日と起算日は「同じ日」
        txtBillKisanDate = txtBillStDate
        strSetumei = "解約予定日が入っていない場合、契約開始日と起算日は「同じ日」になります。"
    Else
        If Format(txtReceiptCyDate, "yyyy/mm") = Format(txtBillStDate, "yyyy/mm") Then
            '一ヶ月以内の契約は,「契約開始年月の一日」が起算日
            txtBillKisanDate = Format(txtBillStDate, "yyyy/mm/01")
            strSetumei = "一ヶ月以内の契約なので、「契約開始年月の一日」が起算日になります。"
        Else
            '解約日が翌月以降の場合,契約日と起算日は「同じ日」
            txtBillKisanDate = txtBillStDate
            strSetumei = "解約予定日が翌月以降の場合、契約開始日と起算日は「同じ日」になります。"
        End If
    End If
    
    If blnDispSetumei Then
        lblSetumei.Caption = strSetumei
    End If

    '日割り計算
    If Day(txtBillKisanDate) = 1 Then
        '起算日が一日の場合日割りなし
        txtBillDailyRate = P_LABEL_日割りなし
        
        '↓INSERT 2020/07/21 S.WATANABE
        lbl_初回日割使用料.Caption = P_LABEL_初回
        lbl_初回日割毎月雑費.Caption = P_LABEL_初回
        lbl_初回日割請求合計.Caption = P_LABEL_初回
        '↑INSERT 2020/07/21 S.WATANABE
        
        'If txtBillEzappiDaily = 0 Then                                                                     'DELETE 2018/03/28 N.IMAI
        If txtBillEzappiDaily = 0 Or (Nz(txtBillEzappiDaily, "") = "" And Nz(txtMotoAcptNo, "") = "") Then  'INSERT 2018/03/28 N.IMAI
            'txtBillEzappiDaily = txtBillEzappi    '' 2012/12/25 M.HONDA INS
            txtBillEzappiDaily = txtBillEzappiDaily_sub '2019/08/08 M.HONDA INS
        End If
        
        '2019/08/08 M.HONDA INS
        'If Replace(txtBillKisanDate, "/", "") >= tax_from And Replace(txtBillKisanDate, "/", "") <= tax_to Then
        If Replace(txtBillKisanDate, "/", "") <= tax_to Then
            txtBillEzappiDaily = txtBillEzappiDaily_sub
        Else
            txtBillEzappiDaily = txtBillEzappi
        End If
          
        
        
        
    Else
        '月の最終日－起算日の日＋１日（起算日も含めるため）
        txtBillDailyRate = Day(DateSerial(Year(txtBillStDate), Month(txtBillStDate) + 1, 0)) _
                            - Day(txtBillKisanDate) + 1
        If txtBillEzappiDaily = "" Then
            txtBillEzappiDaily = 0               '' 2012/12/25 M.HONDA INS
            txtBillEzappiDaily_sub = 0           '2019/08/08 M.HONDA INS
        End If
    
        '2019/08/08 M.HONDA INS
        If txtEzappiCode = "68" Then
            txtBillEzappiDaily = 0
        End If
    
        '↓INSERT 2020/07/21 S.WATANABE
        lbl_初回日割使用料.Caption = P_LABEL_初回日割
        lbl_初回日割毎月雑費.Caption = P_LABEL_初回日割
        lbl_初回日割請求合計.Caption = P_LABEL_初回日割
        '↑INSERT 2020/07/21 S.WATANABE
    
    End If
    
    '▼2008/03/23 add tajima
    '新規状態且つ
    If lblNewExist.Caption = P_LABEL_新規受付 Then
        '契約開始日(年月) >= 請求期間(年月) なら請求期間再設定
        'また解約日が入力済み且つ解約日以降の請求期間ならば再設定
        If Format(txtBillStDate, "yyyy/mm") >= Format(txtBillSeikyuKikan, "yyyy/mm") Or _
          (Nz(txtReceiptCyDate, "") <> "" And _
           Format(txtReceiptCyDate, "yyyy/mm") < Format(txtBillSeikyuKikan, "yyyy/mm")) Then
            
            If txtBillDailyRate = P_LABEL_日割りなし Then
                '請求期間は起算日と同年月
                txtBillSeikyuKikan = Format(txtBillKisanDate, "yyyy/mm")
            Else
                '日割りがあれば翌月まで
                txtBillSeikyuKikan = Format(DateAdd("m", 1, txtBillKisanDate), "yyyy/mm")
            End If
        End If
    End If
    '▲2008/03/23 add tajima
    
    'I:20081001 N.Imai Start
    '割引適用開始月の初期値をセット
    If Nz(txt割引適用開始月) = "" Then
        txt割引適用開始月 = Left$(txtBillStDate, 7)
    End If
    'I:20081001 N.Imai End
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subPageJumpInit
'        機能             :ページ切り替え初期設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subPageJumpInit()

    'ページ切り替えコンボボックス設定
    '2006/06/01 SHIBAZAKI セキュリティカード対応
    cmbPage.RowSource = "先頭;顧客情報;" & _
                             "コンテナ情報;" & _
                             "請求情報;" & _
                             "申込書情報;" & _
                             "鍵情報;" & _
                             "セキュリティカード情報;" & _
                             "保証情報;" & _
                             "鍵発送情報;" & _
                             "受付情報"
    
    'ページ切り替えの先頭位置設定
    '2006/06/01 SHIBAZAKI セキュリティカード対応
    varPageTop = Array(0, lbl顧客情報.Top, _
                          lblコンテナ情報.Top, _
                          lbl請求情報.Top, _
                          lbl申込書情報.Top, _
                          lbl鍵情報.Top, _
                          lblセキュリティカード情報.Top, _
                          lbl保証情報.Top, _
                          lbl鍵発送情報.Top, _
                          lbl受付情報.Top)
    
    
    'ページ切り替えのフォーカス位置設定
    'このコントロールは○○情報ラベルの裏に隠れている。
    'ページ切り替え時にフォーカスをセットして、TABキーをSENDKEYしている。
    '2006/06/01 SHIBAZAKI セキュリティカード対応
    varSetFocus = Array(txtUserFocus, txtUserFocus, _
                                         txtContainerFocus, _
                                         txtBillFocus, _
                                         txtAppFormFocus, _
                                         txtKeyFocus, _
                                         txtSCardFocus, _
                                         txtSecurityFocus, _
                                         txtSendFocus, _
                                         txtReceiptFocus)

End Sub


'==============================================================================*
'
'        MODULE_NAME      :外部から再表示させる
'        MODULE_ID        :subReload
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subReload(Optional intFlg As Integer = 0)
On Error GoTo ErrorHandler

    '顧客情報だけリフレッシュする
    If intFlg = 1 Then
        Call fncUserRefresh
    End If
    
    Exit Sub

ErrorHandler:
    Call MSZZ024_M00("subReload", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncUserRefresh
'        機能             :顧客情報再表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUserRefresh() As Boolean

    On Error GoTo ErrorHandler
    
    Dim dbSQLServer     As Database
    Dim rsData          As Recordset
    Dim blnError        As Boolean
    Dim blnEditing      As Boolean
    
    blnError = False
    fncUserRefresh = False
        
    doCmd.Hourglass True
        
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)

    Set rsData = dbSQLServer.OpenRecordset(fncMakeUserMastSql(), dbOpenSnapshot, dbSQLPassThrough)
    
    If rsData.EOF Then
        'データが存在しない
        MsgBox "顧客マスタを参照できませんでした。"
        GoTo ExitRtn
    End If

    'データ表示
    blnEditing = blnEditMode
    Call subDispUser(rsData)
    If lblNewExist.Caption <> P_LABEL_新規受付 Then
        Call subUserReference
    End If
    Call subEditMode(blnEditing)
    
    fncUserRefresh = True
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "fncUserRefresh" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeUserMastSql
'        機能             :ユーザーマスタ検索SQL作成
'        IN               :
'        OUT              :SQL文
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeUserMastSql() As String

    Dim strSQL  As String
    
    strSQL = strSQL & " SELECT USER_CODE AS UCODE, "
    strSQL = strSQL & "        USER_NAME AS UNAME, "
    strSQL = strSQL & "        USER_KANA AS UKANA, "
    strSQL = strSQL & "        USER_TANM AS UTANM, "
    strSQL = strSQL & "        USER_TAKA AS UTAKA, "
    strSQL = strSQL & "        USER_YUBINO AS UYUBINO, "
    strSQL = strSQL & "        USER_ADR_1 AS UADR1, "
    strSQL = strSQL & "        USER_ADR_2 AS UADR2, "
    strSQL = strSQL & "        USER_ADR_3 AS UADR3, "
    strSQL = strSQL & "        USER_TEL AS UTEL, "
    strSQL = strSQL & "        USER_KEITAI AS UKEITAI, "
    strSQL = strSQL & "        USER_FAX AS UFAX, "
    strSQL = strSQL & "        USER_KKBN AS USERKKBN, "
    strSQL = strSQL & "        USER_MAIL AS UMAIL, "
    strSQL = strSQL & "        USER_SDAY AS USDAY, "
    strSQL = strSQL & "        USER_BIKO AS UBIKO, "
    strSQL = strSQL & "        USER_SKBN AS USKBN, "
    strSQL = strSQL & "        USER_RNAME, "
    strSQL = strSQL & "        USER_RKANA, "
    strSQL = strSQL & "        USER_RPOST, "
    strSQL = strSQL & "        USER_RADR_1, "
    strSQL = strSQL & "        USER_RADR_2, "
    strSQL = strSQL & "        USER_RADR_3, "
    strSQL = strSQL & "        USER_RTEL, "
    strSQL = strSQL & "        USER_RKEITAI, "
    strSQL = strSQL & "        USER_RFAX "
    
    '↓INSERT 2010/04/20 SHIBAZAKI
    strSQL = strSQL & "        , "
    strSQL = strSQL & "        USER_BIRTHDAY, "
    strSQL = strSQL & "        USER_KNAME, "
    strSQL = strSQL & "        USER_KKANA, "
    strSQL = strSQL & "        USER_JOB, "
    strSQL = strSQL & "        USER_KPOST, "
    strSQL = strSQL & "        USER_KADR_1, "
    strSQL = strSQL & "        USER_KADR_2, "
    strSQL = strSQL & "        USER_KADR_3, "
    strSQL = strSQL & "        USER_KTEL, "
    strSQL = strSQL & "        USER_KHOUR_START, "
    strSQL = strSQL & "        USER_KHOUR_END, "
    strSQL = strSQL & "        USER_NENSHU "
    '↑INSERT 2010/04/20 SHIBAZAKI
    
    strSQL = strSQL & "       ,USER_SEIBETU " '2014/10/14 add tajima
        
    strSQL = strSQL & "   FROM USER_MAST "
    strSQL = strSQL & "  WHERE USER_CODE = " & cmbUserCode
    
    fncMakeUserMastSql = strSQL
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncDataSave
'        機能             :保存処理
'        IN               :[intMode]定数宣言「保存処理区分」参照
'                         :[blnConfirm]True=保存確認メッセージ表示
'                         :[blnPageTop]True=保存後の画面リフレッシュで先頭ページに移動
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/06/01 SHIBAZAKI セキュリティカード対応
'Private Function fncDataSave(intMode As Integer, blnConfirm As Boolean) As Boolean
'Private Function fncDataSave(intMode As Integer, blnConfirm As Boolean, Optional blnPageTop As Boolean = True) As Boolean
Public Function fncDataSave(intMode As Integer, blnConfirm As Boolean, Optional blnPageTop As Boolean = True) As Boolean    'UPDATE 2026/04/06 N.IMAI
    On Error GoTo ErrorHandler

    Dim intSaveMode         As Integer
    Dim dbSQLServer         As Database
    Dim objWs               As Workspace
    Dim strMsg              As String
    Dim blnError            As Boolean
    Dim strUserID           As String
    
    '2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
    Dim lngUserNo           As Long
    Dim strACPTNO           As String
    Dim lngReqNo            As Long
    
    Dim lngRTN              As Long     '戻り値 2017/07/31 add by ysuzuki
    
    Dim strucCarg   As Type_CARG_FILE

    blnError = False

    fncDataSave = False
    
    intSaveMode = intMode

    '必須入力チェック
    If Not fncRequireCheck(intSaveMode = P_SAVEMODE_完了) Then
        GoTo ExitRtn
    End If
    
    '入力整合性チェック
    '2006/05/10 SHIBAZAKI システム日付が前回の契約データの初回契約日以降じゃないと受付完了させない
    'If Not fncProprietyCheck() Then
    If Not fncProprietyCheck(intSaveMode = P_SAVEMODE_完了) Then
        GoTo ExitRtn
    End If
    
    '会社コード取得処理
    strCmCampCd = DLookup("CONT_CAMPC", "dbo_CONT_MAST")          '2018/09/25 EGL INS
    
'    If blnConfirm Then 2011/06/18 del tajima 怒対応！blnConfirmは確認メッセージを出したか確認するだけ！割引再計算が先にされない問題があった
        '2009/02/23 ADD N.IMAI 割引更新確認メッセージ表示
        '2009/02/25 MOD N.IMAI
        '2009/03/09 MOD hirano Nz付加
        'If (strSave月額料金 <> txtBillMonthCharge _
        'Or strSave初回料金 <> txtBillFirstkg
        ''2013/10/25 M.HONDA UPD
'        If (strSave月額料金 <> Nz(txtBillMonthCharge, "") _
'        Or strSave初回料金 <> Nz(txtBillFirstkg, "") _
'        Or strSave契約開始日 <> txtBillStDate _
'        Or strSave解約予定日 <> txtReceiptCyDate _
'        Or strSave利用予定期間 <> cmb_利用予定期間 _
'        Or strSave請求期間 <> txtBillSeikyuKikan _
'        Or strSave保証会社 <> cmbContainerHosyoCompany _
'        Or strSave毎月回収方法 <> cmbRcptKaihi _
'        Or blnSave割引再計算 <> False _
'        Or blnSave割引する <> chk割引有効可否.VALUE) _
'        And (chk割引有効可否.VALUE = True) Then
        If (strSave月額料金 <> Nz(txtBillMonthCharge, "") _
        Or strSave初回料金 <> Nz(txtBillFirstkg, "") _
        Or strSave契約開始日 <> txtBillStDate _
        Or strSave解約予定日 <> txtReceiptCyDate _
        Or strSave利用予定期間 <> cmb_利用予定期間 _
        Or strSave請求期間 <> txtBillSeikyuKikan _
        Or strSave保証会社 <> cmbContainerHosyoCompany _
        Or strSave毎月回収方法 <> cmbRcptKaihi _
        Or blnSave割引再計算 <> False _
        Or strSaveトリプル <> chkCampaign.VALUE _
        Or blnSave割引する <> chk割引有効可否.VALUE) _
        And (chk割引有効可否.VALUE = True) Then
        ''2013/10/25 M.HONDA UPD
        
                If blnConfirm = True Then '2011/06/18 怒対応
                    If MsgBox("キャンペーン割引を再計算後保存します。" + vbCrLf + "再計算によって適用割引が変わる場合があります。" + vbCrLf + "よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
                    fncDataSave = True
                    GoTo ExitRtn
                End If
            End If '2011/06/18 怒対応
            '一度再計算する
            Call sub割引再計算(P_MODE_仮計算)
        Else
            Select Case intMode
            Case P_SAVEMODE_保存
                strMsg = "保存します。"
            Case P_SAVEMODE_完了
                strMsg = "受付を完了して、契約データを作成します。"
            Case P_SAVEMODE_取り消し, P_SAVEMODE_保証会社NG
                strMsg = "受付を取り消します。"
            End Select
            If blnConfirm = True Then '2011/06/18 怒対応
                If MsgBox(strMsg & "よろしいですか？", vbQuestion + vbYesNo, P_KAKUNIN) = vbNo Then
                    fncDataSave = True
                    GoTo ExitRtn
                    End If
            End If '2011/06/18 怒対応
        End If
        '2009/02/23 ADD N.IMAI 割引更新確認メッセージ表示
'    End If 2011/06/18 del tajima
    strMsg = ""
    
    doCmd.Hourglass True
    doCmd.SetWarnings False
    
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    
    'トランザクション開始
    Set objWs = DBEngine.Workspaces(0)
    objWs.BeginTrans
    
    '2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
    Call fncNewNo(dbSQLServer, intSaveMode = P_SAVEMODE_完了, lngUserNo, strACPTNO, lngReqNo)
    
    '顧客マスタ更新
    '2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
'    Call fncUpdateUser(dbSqlServer, intSaveMode = P_SAVEMODE_完了)
    Call fncUpdateUser(dbSQLServer, intSaveMode = P_SAVEMODE_完了, lngUserNo)
    
        strUserID = LsGetUserName()
    
    '受付トラン更新
    '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
'    Call fncUpdateRcptTran(dbSqlServer, intMode, strUserId)
    Call fncUpdateRcptTran(dbSQLServer, intMode, strUserID, strACPTNO)
    
    '予約トラン更新
    Call fncUpdateYoukTran(dbSQLServer, _
                           lblNewExist.Caption = P_LABEL_新規受付, _
                           intMode, strUserID)
    txtRcptNo = txtYoukNo
    
    '2008/12/18 発番後に更新するように変更
    'Call fnc適用割引割付更新(dbSqlServer, strUserId, intMode)                           'I:20081001 N.Imai
    
    '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
    If lblNewExist.Caption = P_LABEL_新規受付 Then
        txtAcptNo = strACPTNO
        
        '2006/06/01 SHIBAZAKI セキュリティカード対応
        If fncIsKeiyakuChange() Then
            '契約変更の新規受付（名変、金変、保変）
                Call SetChangeSCards(dbSQLServer, txtMotoAcptNo, txtAcptNo, strUserID, Me.NAME)
        Else
            'それ以外の新規受付
            Call SetContractSCards(dbSQLServer, txtContainerYcode, txtContainerCno, txtAcptNo, strUserID, Me.NAME)
        End If
    End If
    
    '紹介待ち作成
    If intMode = P_SAVEMODE_取り消し Or intMode = P_SAVEMODE_保証会社NG Then
        If IntroWaitSingle(dbSQLServer, txtYoukNo, P_INTROKBN_キャンセル, strMsg, _
                           CLng(Me.txtContainerYcode), CLng(Me.txtContainerCno), P_INTROKBN_受付) = False Then
            Call MSZZ024_M10("IntroWaitSingle", strMsg)
        End If
        
        '2006/06/01 SHIBAZAKI セキュリティカード対応
        If fncIsKeiyakuChange() Then
            '契約変更
            Call SetChangeSCards(dbSQLServer, txtAcptNo, txtMotoAcptNo, strUserID, Me.NAME)
        Else
            'それ以外
            Call CancelContractSCards(dbSQLServer, txtAcptNo, strUserID, Me.NAME)
        End If
    End If
    
    '2008/12/18 発番後に更新するように変更
    If blnEditMode Then  '2009/02/03 ADD hirano 変更フラグの更新
        '2013/10/25M.HONDA UPD
        '2009/02/23 ADD N.IMAI 必要なときのみ更新する
        '2009/03/09 MOD hirano Nz付加
        'If lblNewExist.Caption = P_LABEL_新規受付 _
        'Or strSave月額料金 <> txtBillMonthCharge _
        'Or strSave初回料金 <> txtBillFirstkg
'        If lblNewExist.Caption = P_LABEL_新規受付 _
'        Or strSave月額料金 <> Nz(txtBillMonthCharge, "") _
'        Or strSave初回料金 <> Nz(txtBillFirstkg, "") _
'        Or strSave契約開始日 <> txtBillStDate _
'        Or strSave解約予定日 <> txtReceiptCyDate _
'        Or strSave利用予定期間 <> cmb_利用予定期間 _
'        Or strSave請求期間 <> txtBillSeikyuKikan _
'        Or strSave保証会社 <> cmbContainerHosyoCompany _
'        Or strSave毎月回収方法 <> cmbRcptKaihi _
'        Or blnSave割引する <> chk割引有効可否 _
'        Or blnSave割引再計算 <> False Then
        If lblNewExist.Caption = P_LABEL_新規受付 _
        Or strSave月額料金 <> Nz(txtBillMonthCharge, "") _
        Or strSave初回料金 <> Nz(txtBillFirstkg, "") _
        Or strSave契約開始日 <> txtBillStDate _
        Or strSave解約予定日 <> txtReceiptCyDate _
        Or strSave利用予定期間 <> cmb_利用予定期間 _
        Or strSave請求期間 <> txtBillSeikyuKikan _
        Or strSave保証会社 <> cmbContainerHosyoCompany _
        Or strSaveトリプル <> chkCampaign.VALUE _
        Or strSave毎月回収方法 <> cmbRcptKaihi _
        Or blnSave割引する <> chk割引有効可否 _
        Or blnSave割引再計算 <> False Then
        '2013/10/25M.HONDA UPD
            Call fnc適用割引割付更新(dbSQLServer, strUserID, intMode)                           'I:20081218 N.Imai
            End If
        '2009/02/23 ADD N.IMAI
        End If               '2009/02/03 ADD hirano
    
    '完了ボタンによる更新の場合は完了処理を行う
    If intSaveMode = P_SAVEMODE_完了 Then
        '完了処理
        mstrFaceAuthError = "" ' エラーメッセージ初期化
        Call fncComplete(dbSQLServer, strUserID, strACPTNO, lngReqNo)
' ***** Start 2017/07/31 ad d by ysuzuki
'   ElseIf blnネット予約 = False Then
'2018/03/19    ElseIf blnネット予約 = False And Nz(txtMotoAcptNo.VALUE) = "" Then '2017/10/07 移動でないときを条件に追加
        '電話受付の場合、コンテナ契約ファイル、請求ファイルの登録処理の可能性あり
        '登録の正確な判定は以下メソッドにて行い且つ、登録が必要ならば契約・請求ファイルに対する登録など行う
        'lngRtn = MODIFY_CARG_REQU_FILE(dbSQLServer, lngUserNo, strACPTNO, strUserID)
' ***** end 2017/07/31 add by ysuzuki
'▼2018/03/19 EGL
    Else
        '▼▼▼2020/02/15 chg 特定の移動のとき以外はどんな時でも売上を上げるため請求ファイルと契約トランを作成
        ' 初回支払方法が振込のとき
        '    If CInt(Nz(cmbBillSeikyuCd.VALUE, 0)) = P_SEIKYUCD_振込み Or CInt(Nz(cmbBillSeikyuCd.VALUE, 0)) = P_SEIKYUCD_来社現金 Then
        '    '入金確認日に値が入っているならば
        '    If Nz(txtReceiptPaydate, "") <> "" Then
        '       '入力確認日の入力が発生したならば請求ファイル作成する（洗い替え）
        If Nz(txtMotoAcptNo) <> "" And (txtMoveKbn = P_MOVEKBN_自社理由 Or txtMoveKbn = P_MOVEKBN_金額変更 Or txtMoveKbn = P_MOVEKBN_保障会社変更) Then
            '★自社理由、金額変更、保証会社変更の場合は受付完了まで請求ファイルと契約トランを作らない
        Else
        '▲▲▲2020/02/15 chg
                If CmKomsMod.DeleteRequFile(dbSQLServer, strACPTNO) = False Then
                    GoTo ErrorHandler
                    End If
                    If CmKomsMod.InsertRequFile2(dbSQLServer _
                                                      , lngReqNo _
                                                      , txtBillStDate _
                                                      , Format(txtBillStDate, "yyyy/mm/01") _
                                                      , cmbUserCode _
                                                      , txtContainerYcode _
                                                      , txtContainerCno _
                                                      , Val(txtBillSyoKei + txt初回分割引合計), 0, Val(Nz(txtBillSecukg, 0)) + Val(Nz(txt保証金割引, 0)) _
                                                      , strACPTNO _
                                                      , strMsg) = False Then
                                                      
                                                      
                    GoTo ErrorHandler
                End If
                'さらに顧客理由移動ならば元契約を解約状態にする
                If Nz(txtMotoAcptNo) <> "" And txtMoveKbn = P_MOVEKBN_顧客理由 Then
                    If Not CmKomsMod.UpdateMoveOldCARG(dbSQLServer, _
                                             txtMotoAcptNo, _
                                             Format(txtMoveKbn, "00"), _
                                             txtBillStDate, _
                                             txtReceiptUkdate, _
                                             cmbReceiptUktanto, _
                                             "", _
                                             Me.NAME, _
                                             strUserID, _
                                             strMsg _
                                             ) = pcintCHECK_OK _
                    Then
                        GoTo ErrorHandler
                    End If
                
                
                    End If
                '2018/05/07 N.HONDA INS
                '' 契約トランを作成する
                    Call subSetCargStruc(strucCarg, strACPTNO, -1)
                strucCarg.FSDATE = txtBillStDate                                  '初回契約日
                If Not InsertCargFile(dbSQLServer, strucCarg, Me.NAME, strUserID, strMsg) Then
                    Call MSZZ024_M10("InsertCargFile", strMsg)
                End If
                 '2018/05/07 N.HONDA INS
                 
            '▼▼▼2020/02/15 EGL DEL 移動のとき以外はどんな時でも売上を上げるため請求ファイルと契約トランを作成
            'Else
            '    '入力確認日が消されたら請求削除へ移動し
            '    If CmKomsMod.InsertRequDele(dbSQLServer, strACPTNO, strUserID, Me.NAME) = False Then
            '        GoTo ErrorHandler
            '    End If
            '    '請求ファイルを削除する
            '    If CmKomsMod.DeleteRequFile(dbSQLServer, strACPTNO) = False Then
            '        GoTo ErrorHandler
            '        End If
            '
            '    End If
            '▲▲▲2020/02/15 del
                End If  ' Nz(txtMotoAcptNo) <> "".....
    '▲2018/03/19 EGL
            End If

    'コミット
        objWs.CommitTrans
    Set objWs = Nothing
    
        ' 顔認証連携でエラーが発生していた場合は画面表示
    If mstrFaceAuthError <> "" Then
        MsgBox "顔認証システムの連携でエラーが発生しました。" & vbCrLf & _
               "お手数ですが、管理者へ連絡してください。" & vbCrLf & vbCrLf & _
               "【エラー内容】" & vbCrLf & mstrFaceAuthError, vbExclamation, "顔認証連携エラー"
        mstrFaceAuthError = ""
    End If
    
    blnUpdate = True
    
    dbSQLServer.Close
    Set dbSQLServer = Nothing
    
    '2013/10/31 M.HONDA UPD
    '自社移動理由の場合に、解約の為、貸禁を作らない。
    '解約ヤードは営業終了日が入ってるはずなので、貸禁契約を作らなくても問題ない。
    '自社移動理由の場合
    'If intSaveMode = P_SAVEMODE_完了 And txtMoveKbn = P_MOVEKBN_自社理由 Then
    '    'SQLサーバー接続
    '    Set dbSqlServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    '    If Not AddnReserveType1forCancelYard(dbSqlServer, txtMotoYcode, txtMotoCno, txtMotoUkTanto, Me.NAME, strUserId, strMsg) Then
    '        Call MSZZ024_M10("AddnReserveType1forCancelYard", strMsg)
    '    End If
    '    dbSqlServer.Close
    '    Set dbSqlServer = Nothing
    'End If
    '2013/10/31 M.HONDA UPD
    
    '統合ＤＢへの更新
    If Not MSTGP01_M10(txtContainerYcode.VALUE, txtContainerCno.VALUE) Then     'INSERT 2019/09/06 K.ISHIZAKA 'INSERT 2011/02/21 K.ISHIZAKA
        Call MSZZ024_M10("MSTGP01_M10", "統合ＤＢへの更新に失敗しました。")     'INSERT 2019/09/06 K.ISHIZAKA 'INSERT 2011/02/21 K.ISHIZAKA
    End If                                                                      'INSERT 2019/09/06 K.ISHIZAKA 'INSERT 2011/02/21 K.ISHIZAKA

    If blnConfirm Then
            If intSaveMode = P_SAVEMODE_完了 Then
            MsgBox "契約データを作成しました。"
        ElseIf intSaveMode = P_SAVEMODE_保存 Then
            MsgBox "受付データを更新しました。"
        Else
            MsgBox "受付を取り消しました。"
            End If
    End If

    '画面再表示
    Call subFormRefresh(blnPageTop)

    '割引再計算
    Call sub割引再計算(P_MODE_本計算)                                           'I:20081001 N.Imai
    
    fncDataSave = True
    GoTo ExitRtn

ErrorHandler:
    'ロールバック
    If Not objWs Is Nothing Then
        objWs.Rollback
    End If
    blnError = True

ExitRtn:
    If Not objWs Is Nothing Then Set objWs = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    doCmd.SetWarnings True
    doCmd.Hourglass False

    If blnError Then
        Call Err.Raise(Err.Number, "fncDataSave" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'2006/06/01 SHIBAZAKI セキュリティカード対応
'==============================================================================*
'
'        MODULE_NAME      :fncIsKeiyakuChange
'        機能             :契約変更系ならばTrueが返る
'        IN               :
'        OUT              :True=契約変更系 False=それ以外
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncIsKeiyakuChange() As Boolean

'    fncIsKeiyakuChange = (Nz(txtMoveKbn, "") = P_MOVEKBN_名義変更 Or _
'                          Nz(txtMoveKbn, "") = P_MOVEKBN_金額変更 Or _
'                          Nz(txtMoveKbn, "") = P_MOVEKBN_保障会社変更 Or _
'                          Nz(txtMoveKbn, "") = P_MOVEKBN_満了更新) '2008/03/23 add tajima  'DELETE 2011/09/27 M.RYU

    fncIsKeiyakuChange = (Nz(txtMoveKbn, "") = P_MOVEKBN_名義変更 Or _
                          Nz(txtMoveKbn, "") = P_MOVEKBN_金額変更 Or _
                          Nz(txtMoveKbn, "") = P_MOVEKBN_保障会社変更 Or _
                          Nz(txtMoveKbn, "") = P_MOVEKBN_紹介者金額変更 Or _
                          Nz(txtMoveKbn, "") = P_MOVEKBN_満了更新)                          'INSERT 2011/09/27 M.RYU
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncNewNo
'        機能             :顧客番号／受注契約番号／請求番号発番
'        IN               :objDb　      データベースオブジェクト
'                         :blnComplete　True=完了処理
'                         :lngUserNo　  顧客番号　新規顧客の場合は新規発番される
'                         :strAcptNo　  受注契約番号　完了処理の場合は新規発番される
'                         :lngReqNo　   請求番号発番　完了処理の場合は新規発番される
'        OUT              :True=正常終了 False=異常終了
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncNewNo(objDb As Database, _
                          blnComplete As Boolean, _
                          ByRef lngUserNo As Long, _
                          ByRef strACPTNO As String, _
                          ByRef lngReqNo As Long) As Boolean
On Error GoTo ErrorHandler

    Dim rsData      As Recordset
    Dim lngKyno     As Long
    Dim blnError    As Boolean
    Dim blnUpdate   As Boolean
    
    blnError = False
    
    fncNewNo = False

    ' コントロールマスタより取得
    Set rsData = objDb.OpenRecordset(fncMakeContMastSql(), dbOpenDynaset)

    If rsData.EOF Then
        Call MSZZ024_M10("OpenRecordset", "コントロールマスタを参照できませんでした。")
    End If

    blnUpdate = False
    With rsData
        '新規顧客ならば顧客番号発番、そうでなければ現在の顧客番号
        If fraUserNewExist.VALUE = P_USER_OPT_新規 Then
            lngUserNo = Nz(.Fields("CONT_USER_NO"), 0) + 1
            blnUpdate = True
        Else
            lngUserNo = cmbUserCode
        End If
        
        '完了処理ならば契約番号と請求番号を発番
        '2018/03/19 ↑完了でなくても振込で入金確認日が入力されていれば請求番号は発番する
        If blnComplete Or (Nz(txtReceiptPaydate, "") <> "" And CInt(Nz(cmbBillSeikyuCd.VALUE, 0)) = P_SEIKYUCD_振込み) Then
'       If blnComplete Then
            '請求NO
            lngReqNo = Nz(.Fields("CONT_REQNO"), 0) + 1
            
            '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する
'            '契約NO
'            lngKyno = Nz(.Fields("CONT_KYNO"), 0) + 1
'            strAcptNo = fncMakeAcptNo(lngKyno)
            
            blnUpdate = True
        Else
            lngReqNo = 0
            '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する
'            strAcptNo = ""
        End If
        
        '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する
        If lblNewExist.Caption = P_LABEL_新規受付 Then
            '契約NO
            lngKyno = Nz(.Fields("CONT_KYNO"), 0) + 1
            strACPTNO = fncMakeAcptNo(lngKyno)
        
            blnUpdate = True
        Else
            strACPTNO = Nz(txtAcptNo)
        End If
        
        '*** データ更新（コントロールマスタ）***
        If blnUpdate Then
            .Edit
            If fraUserNewExist.VALUE = P_USER_OPT_新規 Then
                .Fields("CONT_USER_NO") = lngUserNo   ' 顧客No
            End If
            '2018/03/19 ↑完了でなくても振込で入金確認日が入力されていれば請求番号は発番する
            If blnComplete Or (Nz(txtReceiptPaydate, "") <> "" And CInt(Nz(cmbBillSeikyuCd.VALUE, 0)) = P_SEIKYUCD_振込み) Then
'           If blnComplete Then
'                .Fields("CONT_KYNO") = lngKyno         ' 契約No
                .Fields("CONT_REQNO") = lngReqNo       ' 請求NO
                .Fields("CONT_SKDATE") = Format(DATE, "yyyy/mm/dd")     ' 更新日
            End If
            '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する
            If lblNewExist.Caption = P_LABEL_新規受付 Then
                .Fields("CONT_KYNO") = lngKyno         ' 契約No
            End If
            .Fields("CONT_UPDATE") = Now                            ' 更新日
            .UPDATE
        End If
        .Close
    End With

    Set rsData = Nothing

    fncNewNo = True
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncNewNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncMakeAcptNo
'        機能             :受注契約番号作成
'        IN               :lngKyno　    契約番号
'        OUT              :受注契約番号
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeAcptNo(lngKyno As Long) As String

    Dim strYY       As String
    Dim strMM       As String
    
    '年取得
    strYY = ""
    Select Case Format(txtBillStDate, "yyyy")
        Case "1998"
            strYY = "&&"
        Case "1999"
            strYY = "**"
        Case Else
            If Mid(Trim(Format(Me.txtBillStDate, "yyyy")), 1, 2) = "20" Then
                strYY = Mid(Trim(Format(Me.txtBillStDate, "yyyy")), 3)
            End If
    End Select

    '月取得
    strMM = Hex(Trim(Format(Me.txtBillStDate, "mm")))

    ' 戻り値｢受注契約番号｣セット
    fncMakeAcptNo = strBumonCode + strYY & strMM & Trim(Format$(lngKyno, "000000"))

End Function

'==============================================================================*
'
'        MODULE_NAME      :fncRequireCheck
'        機能             :必須入力チェック
'        IN               :True=完了の必須チェック False=完了以外の必須チェック
'        OUT              :True=必須入力が全て入力されている False=されていない
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRequireCheck(blnComplete As Boolean) As Boolean

    fncRequireCheck = False
    
    If fraUserNewExist.VALUE = P_USER_OPT_新規 Then
        '新規顧客の場合の必須チェック
        If Not fncRequireControl(txtUserName, "顧客名") Then Exit Function
        If Not fncRequireControl(txtUserKana, "フリガナ") Then Exit Function
        If Not fncRequireControl(txtUserYubino, "郵便番号") Then Exit Function
        If Not fncRequireControl(txtUserAdr1, "都道府県市町村") Then Exit Function
        If Not fncRequireControl(cmbUserKkbn, "顧客区分") Then Exit Function
        If Not fncRequireControl(cmbUserSkbn, "集金区分") Then Exit Function
        If Not fncRequireControl(txtUserSday, "締日") Then Exit Function
        
        If (txtUserTel.Enabled And Nz(txtUserTel.VALUE) = "") And _
           (txtUserKeitai.Enabled And Nz(txtUserKeitai.VALUE) = "") _
        Then
            MsgBox "電話番号、もしくは携帯番号は必須入力です。", vbInformation, P_KAKUNIN
            txtUserTel.SetFocus
            Exit Function
        End If
        
        If chkUserRinput.VALUE = -1 Then
            'その他連絡先必須チェック
            If Not fncRequireControl(txtUserRname, "担当者名") Then Exit Function
            If Not fncRequireControl(txtUserRkana, "担当者名カナ") Then Exit Function
            If Not fncRequireControl(txtUserRpost, "郵便番号") Then Exit Function
            If Not fncRequireControl(txtUserRadr1, "都道府県市町村") Then Exit Function
            If Not fncRequireControl(txtUserRtel, "電話番号") Then Exit Function
        End If
    Else
        '既存顧客の場合の必須チェック
        If Not fncRequireControl(cmbUserCode, "顧客コード") Then Exit Function
    End If
    
    If Not fncRequireControl(cmbRcptSeiki, "請求書発行区分") Then Exit Function 'INSERT 2010/03/29 K.ISHIZAKA
    
    If Not fncRequireControl(txtBillRentkg, "月額(固定)使用料") Then Exit Function
    'If Not fncRequireControl(txtBillFirstkg, "初回(日割)使用料") Then Exit Function 'DELETE 2020/07/21 S.WATANABE
    If Not fncRequireControl(txtBillFirstkg, lbl_初回日割使用料.Caption + "使用料") Then Exit Function 'INSERT 2020/07/21 S.WATANABE
    
    '2006/06/01 SHIBAZAKI セキュリティカード対応
    If Val(Nz(txtSCardAddNum, 0)) > 0 Then
        If Not fncRequireControl(txtBillAddFzappi2, "追加カード代") Then Exit Function
    End If
    
    If Not fncRequireControl(cmbContainerHosyicd, "保証区分") Then Exit Function
    
    If Not fncRequireControl(txtAppFormLimitDate, "期日") Then Exit Function
    
    If Not fncRequireControl(txtBillStDate, "契約開始日") Then Exit Function
    If Not fncRequireControl(cmbBillSeikyuCd, "初期費用請求方法") Then Exit Function
    
    If Not fncRequireControl(cmbKeyKagiicd, "鍵区分") Then Exit Function
    
    '送付先必須チェック
    If Not fncRequireControl(txtSendName, "鍵送付先名") Then Exit Function
    If Not fncRequireControl(cmbSendHonorific, "鍵送付先敬称") Then Exit Function
    If Not fncRequireControl(txtSendYubino, "鍵送付先郵便番号") Then Exit Function
    If Not fncRequireControl(txtSendAdr1, "鍵送付先都道府県市町村") Then Exit Function
    If Not fncRequireControl(txtSendTel, "鍵送付先電話番号") Then Exit Function
    
    If Not fncRequireControl(cmbReceiptUktanto, "受付担当") Then Exit Function
    
    '2020/06/29 Takenouchi INS Start
    '毎月回収方法の必須チェック
    If Not fncRequireControl(cmbRcptKaihi, "毎月回収方法") Then Exit Function
    '2020/06/29 Takenouchi INS End
    
    '完了時の必須チェック
    If blnComplete Then
        '2006/05/10 SHIBAZAKI 自社理由のリプラス保証での受付時、承認番号が入力されなくても受付完了を可能にする
        'If Nz(txtMoveKbn) <> P_MOVEKBN_自社理由 Then                                                   'DELETE 2018/02/15 N.IMAI
        If Nz(txtMoveKbn) <> P_MOVEKBN_自社理由 And cmbContainerHosyicd <> P_HOSYICD_保証会社無 Then    'INSERT 2018/02/15 N.IMAI
            If Not fncRequireControl(txtSecurityHosyB, "承認番号") Then Exit Function
        End If
        If Not fncRequireControl(txtSendSlipNo, "伝票NO") Then Exit Function
        If txtSendArrivalDate.Enabled Then
            If Not fncRequireControl(txtSendArrivalDate, "配送完了日") Then Exit Function
        Else
            If Not fncRequireControl(txtReceiptPaydate, "入金確認日") Then Exit Function
        End If
    End If
    
    
    fncRequireCheck = True
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncRequireControl
'        機能             :必須チェック
'        IN               :①対象コントロール　②エラーメッセージに出す文言
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncRequireControl(objTarget As Object, strMsgText As String) As Boolean

    fncRequireControl = True
    
    If objTarget.Enabled Then
        If Nz(objTarget.VALUE) = "" Then
            MsgBox strMsgText & "は必須入力です。", vbInformation, P_KAKUNIN
            objTarget.SetFocus
            fncRequireControl = False
        End If
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncProprietyCheck
'        機能             :入力妥当性チェック
'        IN               :True=完了のチェック False=完了以外のチェック
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/05/10 SHIBAZAKI システム日付が前回の契約データの初回契約日以降じゃないと受付完了させない
'Private Function fncProprietyCheck() As Boolean
Private Function fncProprietyCheck(blnComplete As Boolean) As Boolean

    Dim intDummyCancel  As Integer
    
    '2006/09/06 SHIBAZAKI 契約開始日の入力は営業期間内とする
    Dim strBeginDay     As String
    Dim strEndDay       As String
    Dim isHoshouKbnErrH As Boolean      '2020/11/16 INS [EGL]
    
    fncProprietyCheck = False
    
    '2018/05/21 M.HONDA DEL
    '2015/08/05M.HONDA DEL
'    If Left(txtYoukNo, 1) <> "E" Then
'        '2006/05/10 SHIBAZAKI システム日付が前回の契約データの初回契約日以降じゃないと受付完了させない
'        If blnComplete And Nz(txtCargFsDate, "") <> "" Then
'            If DateDiff("d", Format(DATE, "yyyy/mm/dd"), Format(txtCargFsDate, "yyyy/mm/dd")) >= 0 Then
'                MsgBox "本日既に受付完了をしています。明日以降行って下さい。", vbInformation, P_KAKUNIN
'                Exit Function
'            End If
'        End If
'    End If
    '2018/05/21 M.HONDA DEL
    
    '顧客カナの桁数チェック
    If Not fncChkTextInput(txtUserKana, P_KANA_LENGTH, intDummyCancel) Then
        txtUserKana.SetFocus
        Exit Function
    End If
    
    '代表者カナの桁数チェック
    If Not fncChkTextInput(txtUserTaka, P_KANA_LENGTH, intDummyCancel) Then
        txtUserTaka.SetFocus
        Exit Function
    End If
    
    '書類送付先カナの桁数チェック
    If Not fncChkTextInput(txtUserRkana, P_KANA_LENGTH, intDummyCancel) Then
        txtUserRkana.SetFocus
        Exit Function
    End If
    
    '↓INSERT 2010/04/20 SHIBAZAKI
    '緊急連絡先カナの桁数チェック
    If Not fncChkTextInput(txtUserKrenKana, 40, intDummyCancel) Then
        txtUserKrenKana.SetFocus
        Exit Function
    End If
    '緊急連絡先住所１の桁数チェック
    If Not fncChkTextInput(txtUserKrenAdr1, P_ADR_LENGTH, intDummyCancel) Then
        txtUserKrenAdr1.SetFocus
        Exit Function
    End If
    '↑INSERT 2010/04/20 SHIBAZAKI
    
    '2020/12/01 [EGL] DEL Start
    '’2013/05/21 M.HONDA
    ' Tﾎﾟｲﾝﾄ対応
    'If txtTpoint <> "" Then
    '    '' 入力されている際には１６桁が必須
    '    If Len(txtTpoint) <> P_TPOINT_LENGTH Then
    '        MsgBox ("Tポイントは16桁の数値を入力してください。")
    '        txtTpoint.SetFocus
    '        Exit Function
    '    End If
    'End If
    '’2013/05/21 M.HONDA
    '2020/12/01 [EGL] DEL End
    
    '2006/05/10 SHIBAZAKI 住所１の桁数を登録時にチェックする↓
    If fraUserNewExist.VALUE = P_USER_OPT_新規 Then
        '新規顧客の場合
        '顧客住所１
        If Not fncChkTextInput(txtUserAdr1, P_ADR_LENGTH, intDummyCancel) Then
            txtUserAdr1.SetFocus
            Exit Function
        End If
        If chkUserRinput.VALUE = -1 Then
            'その他連絡先住所１
            If Not fncChkTextInput(txtUserRadr1, P_ADR_LENGTH, intDummyCancel) Then
                txtUserRadr1.SetFocus
                Exit Function
            End If
        End If
        '↓INSERT 2010/04/20 SHIBAZAKI
        '勤務先カナの桁数チェック
        If Not fncChkTextInput(txtUserKkana, 40, intDummyCancel) Then
            txtUserKkana.SetFocus
            Exit Function
        End If
        '勤務先住所１の桁数チェック
        If Not fncChkTextInput(txtUserKadr1, P_ADR_LENGTH, intDummyCancel) Then
            txtUserKadr1.SetFocus
            Exit Function
        End If
        '↑INSERT 2010/04/20 SHIBAZAKI
    End If
    
    '集金区分＝コンビニならば未発行でないといけない                             'INSERT START 2010/03/29 K.ISHIZAKA
    If (Nz(cmbUserSkbn.VALUE) = P_USER_SKBN_コンビニ) And (cmbRcptSeiki.VALUE <> P_SEIKI_未発行) Then
        MsgBox "集金区分がコンビニのときは未発行を選択してください。", vbInformation, P_KAKUNIN
        cmbRcptSeiki.SetFocus
        Exit Function
    End If                                                                      'INSERT END   2010/03/29 K.ISHIZAKA
    
    '2009/04/01 保証会社チェック
    If cmbContainerHosyicd = 1 And cmbContainerHosyoCompany = "" Then
        '保証会社有は、保証会社コード指定
        MsgBox "保証会社を選択してください。", vbInformation, P_KAKUNIN
        cmbContainerHosyoCompany.SetFocus
        Exit Function
    End If

    '2017/12/20 ADD EGL
    If cmbContainerHosyoCompany <> "" And bln保証会社新旧 = True Then
        '新旧判定
        If FVS400_fncChkHoshoNewOld(strBumonCode, cmbContainerHosyoCompany) = False Then
            '新→旧のため警告
            MsgBox "新保証会社(新プラン)から旧保証会社(旧プラン)への変更はできません。", vbInformation, P_KAKUNIN
            cmbContainerHosyoCompany.SetFocus
            Exit Function
        End If
    End If
    
    '鍵送付先住所１
    If Not fncChkTextInput(txtSendAdr1, P_ADR_LENGTH, intDummyCancel) Then
        txtSendAdr1.SetFocus
        Exit Function
    End If
    '2006/05/10 SHIBAZAKI 住所１の桁数を登録時にチェックする↑

    '初回請求額は７桁まで（一千万円未満）しか許可しない。
    If txtBillFirstAmount >= 10000000 Then
        MsgBox "初回請求額が一千万円以上になっています。" & vbCrLf & "金額を再調整してください。", vbInformation, P_KAKUNIN
        txtBillRentkg.SetFocus
        Exit Function
    End If
    
    '2006/04/28 S.SHIBAZAKI↓
    ' 鍵送付先名称桁数と住所桁数（１～３全ての桁数合計）をB2仕様に合わせてチェックする
    If CmLenb(txtSendName) > P_KEY_SEND_NAME Then
        MsgBox "お届け先名は全角１５／半角３０文字以内でなければいけません。"
        txtSendName.SetFocus
        Exit Function
    End If
    
    Dim strTemp     As String
    strTemp = Nz(txtSendAdr1) & Nz(txtSendAdr2)
    If Nz(txtSendAdr3) <> "" Then
        '番地と建物名の間に全角空白を入れる
        strTemp = strTemp & "　" & Nz(txtSendAdr3)
    End If
    
    If CmLenb(strTemp) > P_KEY_SEND_ADD Then
        MsgBox "お届け先住所は全角４９／半角９８文字以内でなければいけません。" & vbCrLf & "（ｱﾊﾟｰﾄ・ﾏﾝｼｮﾝ・ﾋﾞﾙ名の記載がある場合、「町・番地」と「ｱﾊﾟｰﾄ・ﾏﾝｼｮﾝ・ﾋﾞﾙ名」の間に全角空白1文字が入ります。）"
        txtSendAdr1.SetFocus
        Exit Function
    End If
    '2006/04/28 S.SHIBAZAKI↑
    
    If fncIsKeiyakuChange() Then
        '契約変更
        '2006/05/10 SHIBAZAKI 契約変更系の受付の場合、契約開始日は移動元契約の契約開始日の翌月以降であること
        If txtBillStDate.Enabled And Nz(txtMotoStDate) <> "" Then
            If DateDiff("d", Format(txtMotoStDate, "yyyy/mm/01"), Format(txtBillStDate, "yyyy/mm/01")) <= 0 Then
                MsgBox "契約開始日は、移動元契約開始日 " & Format(txtMotoStDate, "yyyy/mm/dd") & " の翌月以降でなければいけません。", vbInformation, P_KAKUNIN
                txtBillStDate.SetFocus
                Exit Function
            End If
        End If
    Else
        'それ以外
        '契約開始日は前回の契約のメンテ完了日より大きくなくてはならない
        If txtBillStDate.Enabled And Nz(txtCargMntCmpDate) <> "" Then
            If DateDiff("d", txtCargMntCmpDate, txtBillStDate) <= 0 Then
                MsgBox "前回契約の完了日 " & Format(txtCargMntCmpDate, "yyyy/mm/dd") & " 以降の日付で入力してください。", vbInformation, P_KAKUNIN
                txtBillStDate.SetFocus
                Exit Function
            End If
        End If
        
        '契約開始日は前回の契約の解約日より大きくなくてはならない
        If txtBillStDate.Enabled And Nz(txtCargKyDate) <> "" Then
            If DateDiff("d", txtCargKyDate, txtBillStDate) <= 0 Then
                '↓↓↓↓------20110326---------M.HONDA----------DEL------------↓↓↓↓-------------<s>
                'メンテナンス完了している場合には、解約日は契約日よりも大きくても問題ない。
                'メンテ完了日と契約開始日の関係は上でﾁｪｯｸしているので問題なし。
                If Nz(txtCargMntCmpDate) = "" Then
                    MsgBox "前回契約の解約日 " & Format(txtCargKyDate, "yyyy/mm/dd") & " 以降の日付で入力してください。", vbInformation, P_KAKUNIN
                    txtBillStDate.SetFocus
                    Exit Function
                End If
                '↑↑↑↑------20110326---------M.HONDA----------DEL------------↑↑↑↑-------------<s>
            End If
        End If
    End If
    
    '2006/09/06 SHIBAZAKI 契約開始日の入力は営業期間内とする↓
    If txtBillStDate.Enabled Then
        strBeginDay = IIf(txtContainerBeginDay = "", "0001/01/01", txtContainerBeginDay)
        strEndDay = IIf(txtContainerRentEndDay = "", "9999/12/31", txtContainerRentEndDay)
        
        If DateDiff("d", strBeginDay, txtBillStDate) < 0 Or _
           DateDiff("d", strEndDay, txtBillStDate) > 0 _
        Then
            MsgBox "契約開始日は営業期間内 [" & txtContainerKikan & "] で入力してください。", vbInformation, P_KAKUNIN
            txtBillStDate.SetFocus
            Exit Function
        End If
    End If
    '2006/09/06 SHIBAZAKI 契約開始日の入力は営業期間内とする↑
    
    '請求期間は起算月以降で無ければならない
    If txtBillSeikyuKikan.Enabled And Nz(txtBillSeikyuKikan) <> "" Then
        If DateDiff("d", Format(txtBillKisanDate, "yyyy/mm/01"), fncFirstDay(txtBillSeikyuKikan)) < 0 Then
            MsgBox "請求期間は起算月以降で無ければいけません。", vbInformation, P_KAKUNIN
            txtBillSeikyuKikan.SetFocus
            Exit Function
        End If
    End If
    
    '解約日は契約開始日以降で無ければならない
    If txtReceiptCyDate.Enabled And Nz(txtReceiptCyDate) <> "" Then
        If DateDiff("d", txtBillStDate, txtReceiptCyDate) < 0 Then
            MsgBox "解約日は契約開始日以降で無ければいけません。", vbInformation, P_KAKUNIN
            txtReceiptCyDate.SetFocus
            Exit Function
        End If
    End If
    
    '2006/09/06 SHIBAZAKI 解約予定日の入力は営業終了日までとする↓
    If txtReceiptCyDate.Enabled And Nz(txtReceiptCyDate) <> "" And txtContainerRentEndDay <> "" Then
        If DateDiff("d", txtContainerRentEndDay, txtReceiptCyDate) > 0 Then
            MsgBox "解約日は営業終了日 [" & txtContainerRentEndDay & "] 以前で無ければいけません。", vbInformation, P_KAKUNIN
            txtReceiptCyDate.SetFocus
            Exit Function
        End If
    End If
    '2006/09/06 SHIBAZAKI 解約予定日の入力は営業終了日までとする↑
    
    '2006/05/10 SHIBAZAKI 解約月以降の請求期間は登録できないようにする↓
    If txtBillSeikyuKikan.Enabled And Nz(txtBillSeikyuKikan) <> "" And Nz(txtReceiptCyDate) <> "" Then
        If DateDiff("d", fncFirstDay(txtBillSeikyuKikan), Format(txtReceiptCyDate, "yyyy/mm/01")) < 0 Then
            MsgBox "請求期間は解約予定月以前で無ければいけません。", vbInformation, P_KAKUNIN
            txtBillSeikyuKikan.SetFocus
            Exit Function
        End If
    End If
    '2006/05/10 SHIBAZAKI 解約月以降の請求期間は登録できないようにする↑
    
    '2013/10/16 M.HONDA INS
    '更新不可の際には解約日は必須
    If cmb_更新区分 = "01" Then
        If IsNull(txtReceiptCyDate) Or txtReceiptCyDate = "" Then
            MsgBox "更新不可の場合、解約日は必須となります。", vbInformation, P_KAKUNIN
            txtReceiptCyDate.SetFocus
            Exit Function
        End If
    End If
    '2013/10/16 M.HONDA INS
    
    '2019/11/20 M.HONDA UPD
    If strBumonCode = "H" Then
        If Left(txtYoukNo, 1) = "U" Or cmbBillSeikyuCd = "02" Then
            If blnComplete And cmd_RCPT_SHOUMEI_KDATE.Enabled Then
                MsgBox "証明書の確認を行ってください。", vbInformation, P_KAKUNIN
                Exit Function
            End If
        End If
    Else
        If blnComplete And cmd_RCPT_SHOUMEI_KDATE.Enabled Then                      'INSERT START 2015/07/16 K.ISHIZAKA
            MsgBox "証明書の確認を行ってください。", vbInformation, P_KAKUNIN
            'cmd_RCPT_SHOUMEI_KDATE.SetFocus
            Exit Function
        End If                                                                      'INSERT END   2015/07/16 K.ISHIZAKA
    End If
    '2019/11/20 M.HONDA UPD
    
    'INSERT START 2020/02/25 EGL
    'del 2020/11 Const strErrMsgHosyo = "保証区分が02保証会社無(自社)選択時は" & vbCrLf & "顧客区分「個人」" & vbCrLf & "毎月回収方法「2:クレジットカード」" & vbCrLf & "契約者の都道府県が「東京、神奈、埼玉、千葉」" & vbCrLf & "でなければ保存できません。"
    '▼add 2020/11/16 egl コンテナ部門
    'Const strErrMsgHosyo = "保証区分が「02：自社保証」選択時は" & vbCrLf & "顧客区分「個人」" & vbCrLf & "毎月回収方法「2:クレジットカード」" & vbCrLf & "契約者の都道府県が「東京、神奈、千葉、埼玉」" & vbCrLf & "サイズが４帖以下" & vbCrLf & "携帯番号の入力あり" & vbCrLf & "でなければ保存できません。"  'DELETE 2021/12/01 N.IMAI
    'Const strErrMsgHosyoH = "保証区分が「02：自社保証」選択時は" & vbCrLf & "顧客区分「個人」" & vbCrLf & "毎月回収方法「2:クレジットカード」" & vbCrLf & "契約者の都道府県が「東京、神奈」" & vbCrLf & "サイズが４帖以下" & vbCrLf & "携帯番号の入力あり" & vbCrLf & "でなければ保存できません。"             'DELETE 2021/05/10 N.IMAI
    'Const strErrMsgHosyoH = "保証区分が「02：自社保証」選択時は" & vbCrLf & "顧客区分「個人」" & vbCrLf & "毎月回収方法「2:クレジットカード」" & vbCrLf & "契約者の都道府県が「東京、神奈、千葉、埼玉」" & vbCrLf & "サイズが４帖以下" & vbCrLf & "携帯番号の入力あり" & vbCrLf & "でなければ保存できません。" 'DELETE 2021/12/01 N.IMAI 'INSERT 2021/05/10 N.IMAI
    Const strErrMsgHosyo = "保証区分が「02：自社保証」選択時は" & vbCrLf & "顧客区分「個人」" & vbCrLf & "毎月回収方法「2:クレジットカード」" & vbCrLf & "月額が\50,000以下" & vbCrLf & "携帯番号の入力あり" & vbCrLf & "でなければ保存できません。"
    Const strErrMsgHosyoH = strErrMsgHosyo
    '▲add 2020/11/16 egl コンテナ部門
    '保証会社無し対応
    If cmbContainerHosyicd = "02" Then  '保証区分
        '▼add 2020/11/16 egl コンテナ部門
        If strBumonCode = "H" Then
            isHoshouKbnErrH = False
            '顧客区分が「個人」以外の場合
            If cmbUserKkbn <> "02" Then
                isHoshouKbnErrH = True
            End If
            '毎月回収方法が「2:クレジットカード」以外の場合
            If cmbRcptKaihi <> "2" Then
                isHoshouKbnErrH = True
            End If
            '契約者の都道府県が「東京、神奈」以外で始まる場合
            'If Mid(txtUserAdr1, 1, 2) <> "東京" And Mid(txtUserAdr1, 1, 2) <> "神奈" Then                                                                          'DELETE 2021/05/10 N.IMAI
            'DELETE 2021/12/01 N.IMAI Start
'            If Mid(txtUserAdr1, 1, 2) <> "東京" And Mid(txtUserAdr1, 1, 2) <> "神奈" And Mid(txtUserAdr1, 1, 2) <> "埼玉" And Mid(txtUserAdr1, 1, 2) <> "千葉" Then 'INSERT 2021/05/10 N.IMAI
'                isHoshouKbnErrH = True
'            End If
            'DELETE 2021/12/01 N.IMAI End
            '携帯番号が「入力なし」の場合
            If Nz(txtUserKeitai) = "" Then
                isHoshouKbnErrH = True
            End If
            'DELETE 2021/12/01 N.IMAI Start
'            'サイズが４帖より大きい場合
'            If Val(txtContainerSize) > 4 Then
'                isHoshouKbnErrH = True
'            End If
            'DELETE 2021/12/01 N.IMAI End
            'INSERT 2021/12/01 N.IMAI Start
            If CLng(Nz(txtBillRentkg, "0")) > 50000 Then
                isHoshouKbnErrH = True
            End If
            If Mid(txtContainerYadd, 1, 2) <> "東京" And Mid(txtContainerYadd, 1, 2) <> "神奈" And Mid(txtContainerYadd, 1, 2) <> "埼玉" And Mid(txtContainerYadd, 1, 2) <> "千葉" Then
                isHoshouKbnErrH = True
            End If
            'INSERT 2021/12/01 N.IMAI End
            If isHoshouKbnErrH Then
                MsgBox strErrMsgHosyoH, vbInformation, P_KAKUNIN
                Exit Function
            End If
        '別の部門
        '▲add 2020/11/16 egl コンテナ部門
        Else
            If cmbUserKkbn <> "02" Then        '顧客区分
                '個人以外
                MsgBox strErrMsgHosyo, vbInformation, P_KAKUNIN
                Exit Function
            End If
            If cmbRcptKaihi <> "2" Then       '回収方法
                'クレジットカード以外
                MsgBox strErrMsgHosyo, vbInformation, P_KAKUNIN
                Exit Function
            End If
            'DELETE 2021/12/01 N.IMAI Start
'            Select Case Mid(txtUserAdr1, 1, 2)
'            Case "東京", "神奈", "埼玉", "千葉"
'            Case Else
'                '東京､神奈､埼玉､千葉 以外
'                MsgBox strErrMsgHosyo, vbInformation, P_KAKUNIN
'                Exit Function
'            End Select
            'DELETE 2021/12/01 N.IMAI End
            '▼add 2020/11/16 egl 追加条件
            '携帯番号が「入力なし」の場合
            If Nz(txtUserKeitai) = "" And txtReceiptUkdate > "2020/11/16" Then
                MsgBox strErrMsgHosyo, vbInformation, P_KAKUNIN
            End If
            'DELETE 2021/12/01 N.IMAI Start
'            'サイズが４帖より大きい場合
'            If Val(txtContainerSize) > 4 And txtReceiptUkdate > "2020/11/16" Then
'                MsgBox strErrMsgHosyo, vbInformation, P_KAKUNIN
'            End If
            'DELETE 2021/12/01 N.IMAI End
            '▲add 2020/11/16 egl 追加条件
            'INSERT 2021/12/01 N.IMAI Start
            If CLng(Nz(txtBillRentkg, "0")) > 50000 Then
                isHoshouKbnErrH = True
            End If
            If Mid(txtContainerYadd, 1, 2) <> "東京" And Mid(txtContainerYadd, 1, 2) <> "神奈" And Mid(txtContainerYadd, 1, 2) <> "埼玉" And Mid(txtContainerYadd, 1, 2) <> "千葉" Then
                isHoshouKbnErrH = True
            End If
            'INSERT 2021/12/01 N.IMAI End
        End If
    End If
    'INSERT END   2020/02/25 EGL
    
    '2020/12/01 [EGL] INS Start
    ' 入力されたクーポンCDがエラーコード「00000000」の場合はエラーにする
    If cmb_クーポンCD.VALUE = "00000000" Then
        MsgBox "紹介コード「00000000」は保存はできません。", vbInformation, P_KAKUNIN
        cmb_クーポンCD.SetFocus
        Exit Function
    End If
    ' 入力されたクーポンCDの有効値チェックを行い、無効の場合はエラーとする
    If fncChkExistCmbInput(cmb_クーポンCD) < 0 Then
        MsgBox "入力された紹介コードは無効なコードです。", vbInformation, P_KAKUNIN
        cmb_クーポンCD.SetFocus
        Exit Function
    End If
    '2020/12/01 [EGL] INS End
    
    'INSERT 2023/09/21 N.IMAI Start
    If cmbContainerHosyoCompany = "970170" _
    And Nz(txtMoveKbn, "") <> P_MOVEKBN_自社理由 Then
        '自社理由移動以外はSumicaチェック選択不可   INSERT 2023/09/26 N.IMAI
        If lblNewExist.Caption <> P_LABEL_受付完了 Then
            MsgBox "保証会社「Sumicaチェック」は選択出来ません。", vbExclamation, P_KAKUNIN
            cmbContainerHosyoCompany.SetFocus
            Exit Function
        End If
    End If
    'INSERT 2023/09/21 N.IMAI End
    
    fncProprietyCheck = True

End Function

'==============================================================================*
'
'        MODULE_NAME      :新規顧客追加
'        MODULE_ID        :fncUpdateUser
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
'Private Function fncUpdateUser(objDb As Database, blnComplete As Boolean) As Boolean
Private Function fncUpdateUser(objDb As Database, blnComplete As Boolean, lngUserNo As Long) As Boolean
On Error GoTo ErrorHandler

    Dim rsData      As Recordset
    Dim strSQL      As String
    Dim blnError    As Boolean
    Dim strMsg      As String
    
    blnError = False
    
    fncUpdateUser = False
    
'2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
'    '新規顧客ならば顧客番号発番、そうでなければ何もしない
'    If fraUserNewExist.Value = P_USER_OPT_新規 Then
'        lngUserNo = fncNewUserNo(objDb)
'    Else
'        lngUserNo = cmbUserCode
'    End If
    
    '完了処理かつ名義変更の場合、顧客マスタ更新
    If blnComplete And txtMoveKbn = P_MOVEKBN_名義変更 Then
        If Not UpdateOldUserMast(objDb, lngUserNo, txtMotoUserCode, strMsg) Then
            Call MSZZ024_M10("UpdateOldUserMast", strMsg)
        End If
    End If
    
    If fraUserNewExist.VALUE = P_USER_OPT_新規 Then
        '顧客マスタ検索
        strSQL = "SELECT * FROM USER_MAST WHERE USER_CODE = " & lngUserNo
        Set rsData = objDb.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)
    
        '検索できた場合は、顧客コードが重複している
        If Not rsData.EOF Then
            Call MSZZ024_M10("OpenRecordset", "顧客コードが重複しています。")
        End If
    
        ' 新規に発番した顧客コードをコンボボックスに設定する
        cmbUserCode = Format(lngUserNo, "000000")
        
        '顧客マスタ追加
        Call subInsUser(rsData)
        
        'INSERT 2018/01/29 N.IMAI End
        '仮想口座追加
        strSQL = ""
        strSQL = strSQL & " SELECT *"
'        strSql = strSql & "   FROM KOUZ_MAST"  'DELETE 2019/05/01 Y.WADA
'        strSQL = strSQL & "   FROM KOUZ_MAST_J" 'INSERT 2019/05/01 Y.WADA
        strSQL = strSQL & "   FROM KOUZ_MAST"                                   'INSERT 2023/03/20 N.IMAI
        strSQL = strSQL & "  WHERE KOUZM_BUMOC = '" & strBumonCode & "'"
        strSQL = strSQL & "    AND NZ(KOUZM_KOKYC, '') = '' "
        Set rsData = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
        If rsData.EOF = False Then
            'INSERT 2023/03/27 N.IMAI Start
            strSQL = ""
            strSQL = strSQL & "UPDATE KOUZ_MAST_J"
            strSQL = strSQL & "   SET KOUZM_KOKYC = '" & Format(cmbUserCode, "000000") & "'"
            strSQL = strSQL & "      ,KOUZM_UPDAD = '" & Format(Now(), "yyyymmdd") & "'"
            strSQL = strSQL & "      ,KOUZM_UPDAJ = '" & Format(Now(), "hhnnss") & "'"
            strSQL = strSQL & "      ,KOUZM_UPDPB = '" & Me.NAME & "'"
            strSQL = strSQL & "      ,KOUZM_UPDUB = '" & LsGetUserName() & "'"
            'strSQL = strSQL & " WHERE KOUZM_KOUZB = '" & rsData.Fields("KOUZM_KOKYC") & "'"    'DELETE 2023/04/21 N.IMAI
            strSQL = strSQL & " WHERE KOUZM_KOUZB = '" & rsData.Fields("KOUZM_KOUZB") & "'"     'INSERT 2023/04/21 N.IMAI
            strSQL = strSQL & "   AND KOUZM_KINYC = '" & rsData.Fields("KOUZM_KINYC") & "'"
            strSQL = strSQL & "   AND KOUZM_SHITC = '" & rsData.Fields("KOUZM_SHITC") & "'"
            Call ADODB_Execute(strSQL)
            'INSERT 2023/03/27 N.IMAI End
'DELETE 2023/03/27 N.IMAI Start
'            rsData.Edit
'            rsData.Fields("KOUZM_KOKYC") = Format(cmbUserCode, "000000")
'            rsData.Fields("KOUZM_UPDAD") = Format(Now(), "yyyymmdd")
'            rsData.Fields("KOUZM_UPDAJ") = Format(Now(), "hhnnss")
'            rsData.Fields("KOUZM_UPDPB") = Me.NAME
'            rsData.Fields("KOUZM_UPDUB") = LsGetUserName()
'            rsData.UPDATE
'DELETE 2023/03/27 N.IMAI End
        End If
        'INSERT 2018/01/29 N.IMAI End
        End If
    
    fncUpdateUser = True
    
        GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncUpdateUser" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
''==============================================================================*
''
''        MODULE_NAME      :fncNewUserNo
''        機能             :顧客番号発番
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncNewUserNo(objDb As Database) As Long
'On Error GoTo ErrorHandler
'
'    Dim rsData      As Recordset
'    Dim blnError    As Boolean
'
'    blnError = False
'
'    ' コントロールマスタより顧客Noを取得
'    Set rsData = objDb.OpenRecordset(fncMakeContMastSql(), dbOpenDynaset)
'
'    If rsData.EOF Then
'        Call MSZZ024_M10("OpenRecordset", "コントロールマスタを参照できませんでした。")
'    End If
'
'    With rsData
'        .Edit
'        .Fields("CONT_USER_NO") = .Fields("CONT_USER_NO") + 1        ' 顧客No
'        .Fields("CONT_UPDATE") = Now  ' 更新日
'        .UPDATE
'    End With
'
'    fncNewUserNo = rsData.Fields("CONT_USER_NO")
'
'    GoTo ExitRtn
'
'ErrorHandler:
'    blnError = True
'
'ExitRtn:
'    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
'    If blnError Then
'        Call Err.Raise(Err.Number, "fncNewUserNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :コントロールマスタ検索SQL作成
'        MODULE_ID        :fncMakeContMastSql
'        CREATE_DATE      :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncMakeContMastSql() As String

    Dim strSQL      As String

    strSQL = "SELECT * "
    strSQL = strSQL & "  FROM CONT_MAST "
    strSQL = strSQL & " WHERE CONT_KEY = " & P_CONT_KEY

    fncMakeContMastSql = strSQL

End Function

'==============================================================================*
'
'        MODULE_NAME      :subInsUser
'        機能             :顧客マスタ新規追加
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subInsUser(rsUser As Recordset)
On Error GoTo ErrorHandler

    With rsUser
        .AddNew
        .Fields("USER_CODE") = cmbUserCode
        .Fields("USER_NAME") = txtUserName
        .Fields("USER_KANA") = txtUserKana
        .Fields("USER_TANM") = txtUserTanm
        .Fields("USER_TAKA") = txtUserTaka
        .Fields("USER_YUBINO") = txtUserYubino
        .Fields("USER_ADR_1") = txtUserAdr1
        .Fields("USER_ADR_2") = txtUserAdr2
        .Fields("USER_ADR_3") = txtUserAdr3
        .Fields("USER_TEL") = txtUserTel
        .Fields("USER_KEITAI") = txtUserKeitai
        .Fields("USER_FAX") = txtUserFax
        .Fields("USER_SDAY") = IIf(Nz(txtUserSday) <> "", txtUserSday, 0)
        .Fields("USER_KKBN") = IIf(Nz(cmbUserKkbn) <> "", cmbUserKkbn, 0)
        .Fields("USER_SKBN") = IIf(Nz(cmbUserSkbn) <> "", cmbUserSkbn, 0)
        .Fields("USER_BIKO") = txtUserBiko
        .Fields("USER_MAIL") = txtUserMail
        .Fields("USER_RNAME") = txtUserRname
        .Fields("USER_RKANA") = txtUserRkana
        .Fields("USER_RPOST") = txtUserRpost
        .Fields("USER_RADR_1") = txtUserRadr1
        .Fields("USER_RADR_2") = txtUserRadr2
        .Fields("USER_RADR_3") = txtUserRadr3
        .Fields("USER_RTEL") = txtUserRtel
        .Fields("USER_RKEITAI") = txtUserRkeitai
        .Fields("USER_RFAX") = txtUserRfax
        .Fields("USER_UPDATE") = DATE
        
        '↓INSERT 2010/04/20 SHIBAZAKI
        If Trim(Nz(txtUserBirthday, "")) <> "" Then
            .Fields("USER_BIRTHDAY") = Replace(txtUserBirthday, "/", "")
        End If
        .Fields("USER_KNAME") = txtUserKname
        .Fields("USER_KKANA") = txtUserKkana
        .Fields("USER_JOB") = txtUserJob
        .Fields("USER_KPOST") = txtUserKpost
        .Fields("USER_KADR_1") = txtUserKadr1
        .Fields("USER_KADR_2") = txtUserKadr2
        .Fields("USER_KADR_3") = txtUserKadr3
        .Fields("USER_KTEL") = txtUserKtel
        .Fields("USER_KHOUR_START") = txtUserKhourStart
        .Fields("USER_KHOUR_END") = txtUserKhourEnd
        .Fields("USER_NENSHU") = txtUserNenshu
        '↑INSERT 2010/04/20 SHIBAZAKI
        
        '↓INSERT 2014/10/14 tajima
        If Nz(cmb_性別) <> "" Then
            .Fields("USER_SEIBETU") = cmb_性別
        End If
        '↑INSERT 2014/10/14 tajima
        
        .UPDATE
    End With
    
    Exit Sub

ErrorHandler:
    Call Err.Raise(Err.Number, "fncInsUser" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncUpdateRcptTran
'        機能             :受付トラン更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
'Private Function fncUpdateRcptTran(objDb As Database, intMode As Integer, strUserId As String) As Boolean
    Private Function fncUpdateRcptTran(objDb As Database, intMode As Integer, strUserID As String, strACPTNO As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim strSQL      As String
    Dim rsData      As Recordset
    Dim blnError    As Boolean
    
    blnError = False
    
    fncUpdateRcptTran = False
    
    strSQL = "SELECT * FROM RCPT_TRAN WHERE RCPT_NO = '" & txtYoukNo & "'"
    Set rsData = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    '新規で登録時に既に同受付番号が存在している
    If (lblNewExist.Caption = P_LABEL_新規受付 And Not rsData.EOF) Then
        Call MSZZ024_M10("OpenRecordset", "受付トランが重複しています。")
    End If
    
    '既存データ変更時に受付トランを検索できなかった
    If (lblNewExist.Caption <> P_LABEL_新規受付 And rsData.EOF) Then
        Call MSZZ024_M10("OpenRecordset", "受付トランが検索できませんでした。")
    End If
    
    With rsData
        If .EOF Then
            '新規受付
            .AddNew
            
            .Fields("RCPT_NO") = txtYoukNo

            '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
            .Fields("RCPT_CARG_ACPTNO") = strACPTNO

            .Fields("RCPT_INSED") = Format$(DATE, "yyyymmdd")
            .Fields("RCPT_INSEJ") = Format$(time, "hhmmss")
            .Fields("RCPT_INSPB") = Me.NAME
            .Fields("RCPT_INSUB") = strUserID
            .Fields("RCPT_NEW_HOSPLN_IKOU") = strNewHosplnIkou    '2018/08/06 EGL INS
            .Fields("RCPT_CAMPC") = strCmCampCd                   '2018/09/25 EGL INS
        Else
            '既存受付
            .Edit
            
            .Fields("RCPT_UPDAD") = Format$(DATE, "yyyymmdd")
            .Fields("RCPT_UPDAJ") = Format$(time, "hhmmss")
            .Fields("RCPT_UPDPB") = Me.NAME
            .Fields("RCPT_UPDUB") = strUserID
        End If
        Call subSetUser(rsData)
        Call subSetContainer(rsData)
        Call subSetAppForm(rsData)
        Call subSetBill(rsData)
        Call subSetKey(rsData)
        Call subSetSecurity(rsData)
        Call subSetSend(rsData)
        Call subSetReceipt(rsData)
        
        '取り消し理由
        If intMode = P_SAVEMODE_取り消し Then
            .Fields("RCPT_CREASON") = P_CANCEL_REASON_手動取り消し
        ElseIf intMode = P_SAVEMODE_保証会社NG Then
            .Fields("RCPT_CREASON") = P_CANCEL_REASON_保証会社NG
        End If
        
        '2018/03/16 M.HONDA INS
        '受付完了の場合のみ初回契約開始日を入れる。
        '保存では初回契約開始日は入れない。
        '移動を考慮すると初回契約開始日が入っている=契約完了が行われたデータ
        'なので、初回契約開始日を引継ぐ
        If cmbContainerHosyoCompany = "970170" Or cmbContainerHosyoCompany = "970035" Or cmbContainerHosyoCompany = "970025" Then
            If intMode = P_SAVEMODE_完了 Then
                If Nz(txtShoStDate, "") = "" Then
                    .Fields("RCPT_SYOKISAN_DATE") = txtBillStDate
                Else
                    .Fields("RCPT_SYOKISAN_DATE") = txtShoStDate
                End If
            Else
                If Nz(txtShoStDate, "") <> "" Then
                    .Fields("RCPT_SYOKISAN_DATE") = txtShoStDate
                End If
            End If
        End If
        '2018/03/16 M.HONDA INS
        
        '▼2018/08/30 EGL
        .Fields("RCPT_FZAPPI_CODE") = lngFZAPPI_CODE '初回雑費
        .Fields("RCPT_EZAPPI_CODE") = lngEZAPPI_CODE '毎月雑費
        '▲2018/08/30 EGL
        
        .Fields("RCPT_PLAN_CD") = strプランCd '2020/09/28 EGL add
        .Fields("RCPT_SHOUMEI_ADRSAI") = fncZeroByteToNull(cmbShoumeiAdrsai) '2020/11/16 EGL add
       
        .UPDATE
    End With
    
    fncUpdateRcptTran = True
    
        GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncUpdateRcptTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subSetReceipt
'        機能             :受付情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetReceipt(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_UKTANTO") = cmbReceiptUktanto
'        .Fields("RCPT_UKDATE") = txtReceiptUkdate                              'DELETE 2013/06/21 K.ISHIZAKA
        .Fields("RCPT_UKDATE") = CDate(txtReceiptUkdatetime.VALUE)              'INSERT 2013/06/21 K.ISHIZAKA
        .Fields("RCPT_CYDATE") = fncZeroByteToNull(txtReceiptCyDate)
        .Fields("RCPT_CREASON") = fncZeroByteToNull(cmbReceiptCreason)
        .Fields("RCPT_PAYDATE") = fncZeroByteToNull(txtReceiptPaydate)
        .Fields("RCPT_UKNOTE") = txtReceiptUknote
'        .Fields("RCPT_KAIHI") = cmbRcptKaihi.VALUE                             'DELETE 2010/04/20 SHIBAZAKI 'INSERT 2010/03/29 K.ISHIZAKA
        .Fields("RCPT_KAIHI") = IIf(Nz(cmbRcptKaihi.VALUE, "") = "", Null, cmbRcptKaihi.VALUE)  'INSERT 2010/04/20 SHIBAZAKI
        .Fields("RCPT_COMMENT_1") = txtReceiptComment1
        .Fields("RCPT_COMMENT_2") = txtReceiptComment2
        
        '2020/12/01 [EGL] DEL Start
        '.Fields("RCPT_TPOINT") = txtTpoint                                      '' 2013/05/21 INS M.HONDA
        'If Nz(txtTpoint.VALUE, "") <> "" Then                                   'INSERT 2013/06/21 K.ISHIZAKA
        '    .Fields("RCPT_TPOINT_RITU") = GetTPointRitu(.Fields("RCPT_UKDATE")) 'INSERT 2013/06/21 K.ISHIZAKA
        'End If                                                                  'INSERT 2013/06/21 K.ISHIZAKA
        '上記コードを削除していいか要確認
        '2020/12/01 [EGL] DEL End
        
        .Fields("RCPT_KOUKBN") = cmb_更新区分                                   'INSERT 2013/10/16 M.HONDA
        .Fields("RCPT_COUPON_CD") = cmb_クーポンCD.VALUE                        'INSERT 2020/12/04 Takenouchi
        
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetReceipt" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetUser
'        機能             :顧客情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetUser(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_UCODE") = cmbUserCode
        .Fields("RCPT_SEIKI") = Nz(cmbRcptSeiki.VALUE, P_SEIKI_未発行)         'INSERT 2010/03/29 K.ISHIZAKA
        '↓INSERT 2010/04/21 SHIBAZAKI
        .Fields("RCPT_KREN_NAME") = txtUserKrenName
        .Fields("RCPT_KREN_KANA") = txtUserKrenKana
        .Fields("RCPT_KREN_YUBINO") = txtUserKrenPost
        .Fields("RCPT_KREN_ADR_1") = txtUserKrenAdr1
        .Fields("RCPT_KREN_ADR_2") = txtUserKrenAdr2
        .Fields("RCPT_KREN_ADR_3") = txtUserKrenAdr3
        .Fields("RCPT_KREN_TEL") = txtUserKrenTel
        .Fields("RCPT_KREN_KEITAI") = txtUserKrenKeitai
        .Fields("RCPT_KREN_FAX") = txtUserKrenFax
        .Fields("RCPT_KREN_ZOKUGARA") = txtUserKrenZokugara
        '↑INSERT 2010/04/21 SHIBAZAKI
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetUser" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetAppForm
'        機能             :申込書情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetAppForm(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_SEND_CD") = fncZeroByteToNull(cmbAppFormSendCd)
        .Fields("RCPT_BUTU") = txtAppFormButu
        
'        '2006/04/28 S.SHIBAZAKI 乗用具チェック追加。チェックされている場合、申込書に誓約書添付
'        .Fields("RCPT_VEHICLE") = Nz(chkAppFormVehicle.Value, 0)
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetAppForm" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetContainer
'        機能             :コンテナ情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetContainer(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_YCODE") = txtContainerYcode
        .Fields("RCPT_CNO") = txtContainerCno
        .Fields("RCPT_SEV_EXPIRE") = Replace(txtContainerSevExpire, "/", "")
        .Fields("RCPT_SEV_EXPIRE_MONTH") = fncZeroByteToNull(txtSevExmonth)
        .Fields("RCPT_SEV1N") = txtContainerSev1n
        .Fields("RCPT_SEV2N") = txtContainerSev2n
        .Fields("RCPT_SEV3N") = txtContainerSev3n
        .Fields("RCPT_ENDEN") = txtContainerEnden
        .Fields("RCPT_HOSYICD") = cmbContainerHosyicd
        '2009/04/01 ADD <S> hirano 保障会社コード追加
        .Fields("RCPT_HOSYO_CD") = cmbContainerHosyoCompany
        '2009/04/01 ADD <E> hirano
        .Fields("RCPT_SLOPE") = cmb_スロープ    '' 20120111 M.HONFA
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetContainer" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetBill
'        機能             :請求情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetBill(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_MITU_CD") = fncZeroByteToNull(cmbBillMituCd)
        .Fields("RCPT_SEIKYU_CD") = cmbBillSeikyuCd
        .Fields("RCPT_STDATE") = txtBillStDate
        'I:20081001 N.Imai Start
        If Nz(txt割引適用開始月) = "" Then
            '割引開始月が空ならば、適用開始日の年月を格納する
            .Fields("RCPT_DC_STDATE") = Format(txtBillStDate, "yyyymm")
        Else
            .Fields("RCPT_DC_STDATE") = Replace(txt割引適用開始月, "/", "")
        End If
        .Fields("RCPT_DC_ENABLE") = chk割引有効可否.VALUE
        'I:20081001 N.Imai End
        .Fields("RCPT_KISAN_DATE") = txtBillKisanDate
        If Nz(txtBillSeikyuKikan) = "" Then
            '請求期間が空ならば、起算日の年月を格納する
            .Fields("RCPT_SEIKYU_KIKAN") = Format(txtBillKisanDate, "yyyymm")
        Else
            .Fields("RCPT_SEIKYU_KIKAN") = Replace(txtBillSeikyuKikan, "/", "")
        End If
        
        '▼2018/08/06 EGL UPD<S>
        If P_HOSPLN_移行対象 = strNewHosplnIkou Then
            .Fields("RCPT_SECUKG") = Nz(txtBillSecukgKei, 0)
        Else
            .Fields("RCPT_SECUKG") = Nz(txtBillSecukg, 0) + Nz(txt保証金割引, 0) '2011/06/18 chg tajima
        End If
        '▲2018/08/06 EGL UPD<E>
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
        .Fields("RCPT_SECUKG_KASE") = Nz(txtBillSecukgKase, 0)                  'INSERT 2013/04/10 K.ISHIZAKA
        '▼2008/12/04 hirano Add 保証金変更フラグ保持
        .Fields("RCPT_SECUKG_FLG") = chk保証金強制変更.VALUE
        '▲2008/12/04 hirano Add
        .Fields("RCPT_RENTKG") = txtBillRentkg
        .Fields("RCPT_FIRSTKG") = txtBillFirstkg
        
        .Fields("RCPT_SUMIKG") = txtBillSumikg
        
        .Fields("RCPT_EZAPPI") = fncZeroByteToNull(txtBillEzappi)
        .Fields("RCPT_EZAPPI_DAILY") = fncZeroByteToNull(txtBillEzappiDaily)
        .Fields("RCPT_ADD_EZAPPI_CODE1") = fncZeroByteToNull(cmbBillAddEzappiCode1)
        .Fields("RCPT_ADD_EZAPPI1") = fncZeroByteToNull(txtBillAddEzappi1)
        .Fields("RCPT_ADD_EZAPPI_DAILY1") = fncZeroByteToNull(txtBillAddEzappiDaily1)
        .Fields("RCPT_ADD_EZAPPI_CODE2") = fncZeroByteToNull(cmbBillAddEzappiCode2)
        .Fields("RCPT_ADD_EZAPPI2") = fncZeroByteToNull(txtBillAddEzappi2)
        .Fields("RCPT_ADD_EZAPPI_DAILY2") = fncZeroByteToNull(txtBillAddEzappiDaily2)
        
        .Fields("RCPT_FZAPPI") = fncZeroByteToNull(txtBillFzappi)
        .Fields("RCPT_ADD_FZAPPI_CODE1") = fncZeroByteToNull(cmbBillAddFzappiCode1)
        .Fields("RCPT_ADD_FZAPPI1") = fncZeroByteToNull(txtBillAddFzappi1)
        '2006/06/01 SHIBAZAKI セキュリティカード対応
'        .Fields("RCPT_ADD_FZAPPI_CODE2") = fncZeroByteToNull(cmbBillAddFzappiCode2)
        .Fields("RCPT_ADD_FZAPPI_CODE2") = int初回雑費2Cd '2015/01/08 chg 元は→ =P_ZAPPI_追加カード代
        .Fields("RCPT_ADD_FZAPPI2") = fncZeroByteToNull(txtBillAddFzappi2)
        '2009/01/25 add tajima
        .Fields("RCPT_FWARIBIKI") = txt初回分割引合計
        .Fields("RCPT_EWARIBIKI") = txt月額分割引合計
        .Fields("RCPT_HOSHOU_WARIBIKI") = 0 - Nz(txt保証金割引, 0) '画面OBJはマイナスなのでプラスに戻してから 2011/06/18 tajima
        
        '2015/03/24 M.HONDA INS 年払い対応
        .Fields("RCPT_NEN_KBN") = chkNEN_KBN
        If cmbJIREI = "" Then
            .Fields("RCPT_NEN_JIREI") = Null
        Else
            .Fields("RCPT_NEN_JIREI") = cmbJIREI
        End If
        If cmbSEIKI = "" Then
            .Fields("RCPT_NEN_SEIKI") = Null
        Else
            .Fields("RCPT_NEN_SEIKI") = cmbSEIKI
        End If
        .Fields("RCPT_NEN_BIKO") = txtSEIKI
 
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetBill" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetKey
'        機能             :鍵情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetKey(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_KAGIICD") = fncZeroByteToNull(cmbKeyKagiicd)
        .Fields("RCPT_KAGI_NO") = txtKeyKagiNo
        '2006/06/01 SHIBAZAKI セキュリティカード対応
'        .Fields("RCPT_SECURITY_CARD") = txtKeySecurityCard
        .Fields("RCPT_KAGI_LENDNUM") = fncZeroByteToNull(txtKeyKagiLendnum)
        .Fields("RCPT_KAINO") = txtKaijoNo          'INSERT 2014/03/11 MIYAMOTO
        If cmbKeyKagiicd.Column(0) = 16 Then                'INSERT 2021/06/02 EGL
            .Fields("RCPT_BUIL_KAGI_NO") = strKaijoNo_QR    'INSERT 2021/06/02 EGL
            .Fields("RCPT_KAGI_NO") = ""                    'INSERT 2021/06/02 EGL
        End If                                              'INSERT 2021/06/02 EGL
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetKey" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetSecurity
'        機能             :保証情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetSecurity(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_HOSYD") = fncZeroByteToNull(txtSecurityHosyD)
        .Fields("RCPT_HOSYB") = txtSecurityHosyB
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetSecurity" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetSend
'        機能             :鍵送付情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetSend(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("RCPT_SEND_NAME") = txtSendName
        .Fields("RCPT_SEND_HONORIFIC") = fncZeroByteToNull(cmbSendHonorific)
        .Fields("RCPT_SEND_YUBINO") = txtSendYubino
        .Fields("RCPT_SEND_ADR_1") = txtSendAdr1
        .Fields("RCPT_SEND_ADR_2") = txtSendAdr2
        .Fields("RCPT_SEND_ADR_3") = txtSendAdr3
        .Fields("RCPT_SEND_TEL") = txtSendTel
        .Fields("RCPT_KEY_SEND_DATE") = fncZeroByteToNull(txtSendKeySendDate)
        .Fields("RCPT_ARRIVAL_EXPECT") = fncZeroByteToNull(txtSendArrivalExpect)
        .Fields("RCPT_SPECIFY_TIME") = fncZeroByteToNull(cmbSendSpecifyTime)
        .Fields("RCPT_SLIP_NO") = txtSendSlipNo
        .Fields("RCPT_ARRIVAL_DATE") = fncZeroByteToNull(txtSendArrivalDate)
        .Fields("RCPT_YAMATO_DATE") = fncZeroByteToNull(txtSendYamatoDate)  'INSERT 2019/03/22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetSend" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSetSend
'        機能             :鍵送付情報セット
'        IN               :対象
'        OUT              :０バイト文字かNULLならばNULLを返却
'                         :それ以外はそのままを返却
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncZeroByteToNull(objTarget As Object) As Variant

    If Nz(objTarget, "") = "" Then
        fncZeroByteToNull = Null
    Else
        fncZeroByteToNull = objTarget.VALUE
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncUpdateYoukTran
'        機能             :予約トラン更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateYoukTran(objDb As Database, blnNewRcpt As Boolean, intMode As Integer, strUserID As String) As Boolean
On Error GoTo ErrorHandler
    
    Dim strSQL      As String
    Dim rsData      As Recordset
    Dim blnError    As Boolean
    
    blnError = False
    
    fncUpdateYoukTran = False
    
    strSQL = "SELECT * FROM YOUK_TRAN WHERE YOUKT_UKNO = '" & txtYoukNo & "'"
    Set rsData = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    '見つからない場合エラー
    If rsData.EOF Then
        Call MSZZ024_M10("OpenRecordset", "予約トランが検索できませんでした。")
    End If
    
    With rsData
        .Edit
        
        '新規受付時は受付番号設定
        If blnNewRcpt Then
            .Fields("YOUKT_RCPT_NO") = txtYoukNo
        End If
        
        '受付キャンセル
        If intMode = P_SAVEMODE_取り消し Or intMode = P_SAVEMODE_保証会社NG Then
            .Fields("YOUKT_YUKBN") = P_YUKBN_受付キャンセル
            .Fields("YOUKT_YTDATE") = DATE
        End If
        
        '完了
        If intMode = P_SAVEMODE_完了 Then
            .Fields("YOUKT_YUKBN") = P_YUKBN_契約
        End If
        
        .Fields("YOUKT_TKDATE") = txtAppFormLimitDate
        
        '顧客情報更新
        Call subSetYoukUser(rsData)
        
        '▼ 2007/04/13 add tajima
        If txtMoveKbn = P_MOVEKBN_自社理由 Or txtMoveKbn = P_MOVEKBN_顧客理由 Then
            .Fields("YOUKT_MOVE_TANTO") = Nz(fncZeroByteToNull(Me.cmb_荷物移動担当), 0)
        End If
        '▲ 2007/04/13 add tajima
        
        '▼ 2007/09/25 add tajima
        .Fields("YOUKT_USE_PERIOD") = Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0)
        '▲ 2007/09/25 add tajima
        
        
        .Fields("YOUKT_UPDAJ") = Format$(time, "hhmmss")
        .Fields("YOUKT_UPDPB") = Me.NAME
        .Fields("YOUKT_UPDUB") = strUserID
        
        .UPDATE
    End With
    
    fncUpdateYoukTran = True
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncUpdateYoukTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subSetYoukUser
'        機能             :予約トラン顧客情報セット
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSetYoukUser(rsData As Recordset)
On Error GoTo ErrorHandler

    With rsData
        .Fields("YOUKT_UCODE") = cmbUserCode
        .Fields("YOUKT_NAME") = txtUserName
        .Fields("YOUKT_KANA") = txtUserKana
        .Fields("YOUKT_TANM") = txtUserTanm
        .Fields("YOUKT_TAKA") = txtUserTaka
        .Fields("YOUKT_YUBINO") = txtUserYubino
        .Fields("YOUKT_ADR_1") = txtUserAdr1
        .Fields("YOUKT_ADR_2") = txtUserAdr2
        .Fields("YOUKT_ADR_3") = txtUserAdr3
        .Fields("YOUKT_TEL") = txtUserTel
        .Fields("YOUKT_KEITAI") = txtUserKeitai
        .Fields("YOUKT_FAX") = txtUserFax
        .Fields("YOUKT_KKBN") = cmbUserKkbn
        .Fields("YOUKT_MAIL") = txtUserMail
    End With

    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subSetYoukUser" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fncComplete
'        機能             :受付完了処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
'Private Function fncComplete(objDb As Database, strUserId As String) As Boolean
Private Function fncComplete(objDb As Database, strUserID As String, strACPTNO As String, lngReqNo As Long) As Boolean
On Error GoTo ErrorHandler

'2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
'    Dim lngReqNo    As Long
'    Dim strAcptNo   As String
    Dim strMsg      As String
    Dim strucCarg   As Type_CARG_FILE
    Dim strMoveNote As String
'↓ INSERT 2014/06/20 MIYAMOTO
    Dim intMotoNameTo           As Integer
    Dim intSakiNameTo           As Integer
    Dim strMotoConvKagiKbn      As String
    Dim strSakiConvKagiKbn      As String
    Dim intMotoKeyChgRiyuCd     As Integer
    Dim intSakiKeyChgRiyuCd     As Integer
'↑ INSERT 2014/06/20 MIYAMOTO
    
    fncComplete = False
    
    '1.紹介トランの紹介区分を「契約」に更新
    If Not CmnUpdateIntrTran(objDb, txtYoukNo, P_INTROKBN_受付, P_INTROKBN_契約, strMsg) Then
        Call MSZZ024_M10("CmnUpdateIntrTran", strMsg)
    End If
    
    '2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
'    '2.請求番号、契約番号発番
'    Call fncNewAcptNoAndReqNo(objDb, strAcptNo, lngReqNo)
    
    '3.請求ファイルを作成する
    '2008/10/01 割引考慮※txt初回分割引合計を引くのを追加
    '2011/06/18 保証委託割引考慮※txttxt保証金割引を引くのを追加
'2017/10/07 del    If blnネット予約 = True Then    '***** INSERT 2017/07/31 add by ysuzuki
    If blnネット予約 = True Or (blnネット予約 = False And Nz(txtMotoAcptNo.VALUE) <> "") Then   '2017/10/07 add
        If Not InsertRequFile(objDb, lngReqNo, txtBillStDate, _
                            Format(txtBillStDate, "yyyy/mm/01"), cmbUserCode, txtContainerYcode, txtContainerCno, _
                            Val(txtBillSyoKei + txt初回分割引合計), 0, Val(Nz(txtBillSecukg, 0)) + Val(Nz(txt保証金割引, 0)), strACPTNO, strMsg) _
        Then
            Call MSZZ024_M10("InsertRequFile", strMsg)
        End If
' ***** Start 2017/07/31 add by ysuzuki
    Else
        ' 電話契約の場合は請求ファイルがあれば作り直しをします。これは他契約ありで初回支払が振込だったとき請求ファイルが作成済みのためそれを削除する
        If Not CmKomsMod.DeleteRequFile(objDb, strACPTNO) Then
            Call MSZZ024_M10("DeleteRequFile", "請求ファイルの削除に失敗しました。")  'このエラーは請求ファイル削除失敗時
        End If
        If Not CmKomsMod.InsertRequFile2(objDb _
                                         , lngReqNo _
                                         , txtBillStDate _
                                         , Format(txtBillStDate, "yyyy/mm/01") _
                                         , cmbUserCode _
                                         , txtContainerYcode _
                                         , txtContainerCno _
                                         , Val(txtBillSyoKei + txt初回分割引合計), 0, Val(Nz(txtBillSecukg, 0)) + Val(Nz(txt保証金割引, 0)) _
                                         , strACPTNO _
                                         , strMsg) Then
            Call MSZZ024_M10("InsertRequFile2", strMsg)
        End If
    End If
' ***** End 2017/07/31 add by ysuzuki

    '4.移動の場合の処理
    If Nz(txtMotoAcptNo) <> "" Then
'↓ INSERT 2014/06/20 MIYAMOTO
        intMotoKeyChgRiyuCd = -1
        '移動元契約トランに該当する自動鍵解除番号発番トランの登録(移動(自社理由、顧客理由、未使用移動)による解約予定の場合)
        If txtMoveKbn = P_MOVEKBN_自社理由 Or _
           txtMoveKbn = P_MOVEKBN_顧客理由 Or _
           txtMoveKbn = P_MOVEKBN_未使用移動 Then
            '移動元鍵区分から変換鍵区分を取得
'↓ UPDATE 2016/05/17 SUGIMURA
'            intMotoNameTo = DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txtMotoKagiiCd)
'   UPDATE 2021/06/02            intMotoNameTo = MSZZ070_KaiKey(txtYardTihoKbn, DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txtMotoKagiiCd))
'   UPDATE 2021/08/04            intMotoNameTo = MSZZ070_KaiKey(txtYardTihoKbn, DLookup("NAME_VALUE_FROM", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txtMotoKagiiCd))
            intMotoNameTo = MSZZ070_KaiKey(txtYardTihoKbn, DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & txtMotoKagiiCd))
'↑ UPDATE 2016/05/17 SUGIMURA
            '取得した変換鍵区分が-1の場合、移動元鍵区分を使用
'↓ UPDATE 2016/05/17 SUGIMURA
'            If intMotoNameTo = -1 Then
            If intMotoNameTo = 0 Then
'↑ UPDATE 2016/05/17 SUGIMURA
                strMotoConvKagiKbn = txtMotoKagiiCd
            '取得した変換鍵区分が-1以外の場合
            Else
                strMotoConvKagiKbn = intMotoNameTo
                '移動元データが地方ヤード以外の場合
                If txtMotoYardTihoKbn <> -1 Then
                    '鍵変更理由コードの設定("3"をセットする)
                    intMotoKeyChgRiyuCd = 3
'↓ DELETE 2016/05/17 SUGIMURA (上記IF文の外に移動)
'                    '変換鍵区分を使用し、コンテナマスタの更新を行う
'                    Call fncUpdateCntaMast(objDb, "MOTO", txtMotoYcode, txtMotoCno, strMotoConvKagiKbn)
'↑ DELETE 2016/05/17 SUGIMURA (上記IF文の外に移動)
                End If
                '変換鍵区分を使用し、コンテナマスタの更新を行う
                Call fncUpdateCntaMast(objDb, "MOTO", txtMotoYcode, txtMotoCno, strMotoConvKagiKbn)
            End If
            '上記で取得した変換鍵区分が-1、又は変換鍵区分が-1以外で地方ヤード以外の場合
'↓ UPDATE 2016/05/17 SUGIMURA
'            If intMotoNameTo = -1 Or _
'               intMotoNameTo <> -1 And txtMotoYardTihoKbn <> -1 Then
'del 2021/08/04 EGL            If intMotoNameTo <> 0 Then   元バグ、このIF文では移動時に発番されない
'↑ UPDATE 2016/05/17 SUGIMURA
                '移動元契約トランに該当する自動鍵解除番号発番トランの登録
                Call fncInsAkknTran(objDb, _
                                    Nz(txtMotoYcode), _
                                    Nz(txtMotoCno), _
                                    Nz(strMotoConvKagiKbn), _
                                    strUserID)
'del 2021/08/04 EGL                   End If
        End If
'↑ INSERT 2014/06/20 MIYAMOTO

        ' 部屋移動API連携（3ステップフロー：存在確認 -> 登録/有効化 -> 移動）
        ' 移動理由区分が 01:自社理由、02:顧客理由、03:未使用移動 の場合に実行
        If txtMoveKbn = P_MOVEKBN_自社理由 Or _
           txtMoveKbn = P_MOVEKBN_顧客理由 Or _
           txtMoveKbn = P_MOVEKBN_未使用移動 Then
            
            mstrFaceAuthError = PerformFaceAuthTransferFlow(txtMotoYcode, txtMotoCno, txtContainerYcode, txtContainerCno, cmbUserCode)
        End If
        
        Call CreateOldMoveMemo(txtMoveKbn, _
                               txtMoveKbnName, _
                               txtContainerYcode, _
                               txtContainerCno, _
                               cmbUserCode, _
                               txtUserKana, _
                               txtBillRentkg, _
                               txtContainerHosyicd, _
                               strACPTNO, _
                               txtBillStDate, _
                               strMoveNote)
        '移動元契約トラン更新
'↓ UPDATE 2014/06/20 MIYAMOTO
'        If Not UpdateMoveOldCARG(objDb, _
'                                 txtMotoAcptNo, _
'                                 Format(txtMoveKbn, "00"), _
'                                 txtBillStDate, _
'                                 txtReceiptUkdate, _
'                                 cmbReceiptUktanto, _
'                                 strMoveNote, _
'                                 Me.NAME, _
'                                 strUserID, _
'                                 strMsg) = pcintCHECK_OK _
'        Then
        If Not UpdateMoveOldCARG(objDb, _
                                 txtMotoAcptNo, _
                                 Format(txtMoveKbn, "00"), _
                                 txtBillStDate, _
                                 txtReceiptUkdate, _
                                 cmbReceiptUktanto, _
                                 strMoveNote, _
                                 Me.NAME, _
                                 strUserID, _
                                 strMsg, _
                                 intMotoKeyChgRiyuCd) = pcintCHECK_OK _
        Then
'↑ UPDATE 2014/06/20 MIYAMOTO
            Call MSZZ024_M10("UpdateMoveOldCARG", strMsg)
        End If
'↓ DELETE 2014/06/20 MIYAMOTO
''↓ INSERT 2014/04/17 MIYAMOTO
'        '移動元契約トランに該当する自動鍵解除番号発番トランの登録(移動(自社理由、顧客理由、未使用移動)による解約予定の場合)
'        If txtMoveKbn = P_MOVEKBN_自社理由 Or _
'           txtMoveKbn = P_MOVEKBN_顧客理由 Or _
'           txtMoveKbn = P_MOVEKBN_未使用移動 Then
'            Call fncInsAkknTran(objDb, _
'                                Nz(txtMotoYcode), _
'                                Nz(txtMotoCno), _
'                                Nz(txtMotoKagiiCd), _
'                                strUserID)
'        End If
''↑ INSERT 2014/04/17 MIYAMOTO
'↑ DELETE 2014/06/20 MIYAMOTO
    End If

'↓ INSERT 2014/06/20 MIYAMOTO
    intSakiKeyChgRiyuCd = -1
    '上記契約トランが解約予定(解約予定日が入力されている)の場合、該当する自動鍵解除番号発番トランの登録(解約予定の場合)
    If Nz(txtReceiptCyDate) <> "" Then
        '移動先鍵区分から変換鍵区分を取得
'↓ UPDATE 2016/05/17 SUGIMURA
'        intSakiNameTo = DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & cmbKeyKagiicd)
        intSakiNameTo = MSZZ070_KaiKey(txtYardTihoKbn, DLookup("NAME_VALUE_TO", "dbo_NAME_MAST", "NAME_ID='060' AND NAME_CODE=" & cmbKeyKagiicd))
'↑ UPDATE 2016/05/17 SUGIMURA
        '取得した変換鍵区分が-1の場合、移動先鍵区分を使用
'↓ UPDATE 2016/05/17 SUGIMURA
'        If intSakiNameTo = -1 Then
        If intSakiNameTo = 0 Then
'↑ UPDATE 2016/05/17 SUGIMURA
            strSakiConvKagiKbn = cmbKeyKagiicd
        '取得した変換鍵区分が-1以外の場合
        Else
'↓ UPDATE 2016/05/17 SUGIMURA
'            '移動先データが地方ヤード以外の場合
'            If txtYardTihoKbn = -1 Then
'                strSakiConvKagiKbn = cmbKeyKagiicd
'            Else
'                strSakiConvKagiKbn = intSakiNameTo
'                '鍵変更理由コードの設定("3"をセットする)
'                intSakiKeyChgRiyuCd = 3
'            End If
            '移動先データが地方ヤード以外の場合
            If txtYardTihoKbn <> -1 Then
                '鍵変更理由コードの設定("3"をセットする)
                intSakiKeyChgRiyuCd = 3
            End If
            '変換鍵区分に変更
            strSakiConvKagiKbn = intSakiNameTo
'↑ UPDATE 2016/05/17 SUGIMURA
        End If
        '上記で取得した変換鍵区分が-1、又は変換鍵区分が-1以外で地方ヤード以外の場合
'↓ UPDATE 2016/05/17 SUGIMURA
'        If intSakiNameTo = -1 Or _
'           intSakiNameTo <> -1 And txtYardTihoKbn <> -1 Then
        If intSakiNameTo <> 0 Then
'↑ UPDATE 2016/05/17 SUGIMURA
            '移動先契約トランに該当する自動鍵解除番号発番トランの登録
            Call fncInsAkknTran(objDb, _
                                txtContainerYcode, _
                                txtContainerCno, _
                                Val(strSakiConvKagiKbn), _
                                strUserID)
        End If
    Else
        strSakiConvKagiKbn = cmbKeyKagiicd
    End If
'↑ INSERT 2014/06/20 MIYAMOTO
    
    '5.契約トランを作成する
'↓ UPDATE 2014/06/20 MIYAMOTO
'    Call subSetCargStruc(strucCarg, strAcptNo)
    Call subSetCargStruc(strucCarg, strACPTNO, intSakiKeyChgRiyuCd)
'↑ UPDATE 2014/06/20 MIYAMOTO
    If Not InsertCargFile(objDb, strucCarg, Me.NAME, strUserID, strMsg) Then
        Call MSZZ024_M10("InsertCargFile", strMsg)
    End If
    
'↓ DELETE 2014/06/20 MIYAMOTO
''↓ INSERT 2014/04/17 MIYAMOTO
'    '上記契約トランが解約予定の場合、該当する自動鍵解除番号発番トランの登録(解約予定の場合)
'    If strucCarg.AGRE = P_AGRE_解約予定 Then
'        Call fncInsAkknTran(objDb, _
'                            txtContainerYcode, _
'                            txtContainerCno, _
'                            Val(cmbKeyKagiicd), _
'                            strUserID)
'    End If
''↑ INSERT 2014/04/17 MIYAMOTO
'↑ DELETE 2014/06/20 MIYAMOTO
    '2006/06/01 SHIBAZAKI 受付完了時に鍵情報をコンテナマスタに更新する
    'コンテナマスタ更新
'↓ UPDATE 2014/06/20 MIYAMOTO
'    Call fncUpdateCntaMast(objDb)
    Call fncUpdateCntaMast(objDb, "SAKI", txtContainerYcode, txtContainerCno, strSakiConvKagiKbn)
'↑ UPDATE 2014/06/20 MIYAMOTO

    '2006/05/31 SHIBAZAKI 契約番号は新規受付の保存時に発番する。
'    txtAcptNo = strAcptNo
    
    
    fncComplete = True
    
    Exit Function

ErrorHandler:
    'Resume 2009/04/22
    Call Err.Raise(Err.Number, "fncComplete" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'↓ INSERT 2014/06/20 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :fncUpdateCargFile
'        機能             :自動鍵解除番号発番トラン登録処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdateCargFile(objDb As Database, _
                                   ByVal StrYard As String, _
                                   ByVal strCno As String, _
                                   ByVal strUserID As String) As Boolean
    Dim blnError                As Boolean
    Dim strSQL                  As String
    Dim rsCarg                  As Recordset

On Error GoTo ErrorHandler

    blnError = False
    fncUpdateCargFile = False

    strSQL = ""
    strSQL = strSQL & "SELECT       CARG_YCODE" & vbCrLf
    strSQL = strSQL & ",            CARG_NO" & vbCrLf
    strSQL = strSQL & ",            CARG_KEY_CHGRIYUCD" & vbCrLf
    strSQL = strSQL & "FROM         CARG_FILE" & vbCrLf
    strSQL = strSQL & "WHERE        CARG_YCODE = " & Nz(StrYard) & vbCrLf
    strSQL = strSQL & "AND          CARG_NO =" & Nz(strCno) & vbCrLf

    Set rsCarg = objDb.OpenRecordset(strSQL, dbOpenDynaset)

    '見つからない場合エラー
    If rsCarg.EOF Then
        Call MSZZ024_M10("OpenRecordset", "該当するコンテナ契約トランが検索できませんでした。")
    End If

    With rsCarg
        .Edit

        .Fields("CARG_KEY_CHGRIYUCD") = "3"                 '鍵変更理由コードに"3"をセット
        .Fields("CARG_UPDAD") = Format$(DATE, "YYYYMMDD")    '更新日付
        .Fields("CARG_UPDAJ") = Format$(time, "HHMMSS")      '更新時刻
        .Fields("CARG_UPDPB") = Me.NAME                     '更新プログラムID
        .Fields("CARG_UPDUB") = strUserID                   '更新ユーザーID

        .UPDATE
    End With

    fncUpdateCargFile = True
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsCarg Is Nothing Then rsCarg.Close: Set rsCarg = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncUpdateCargFile" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'↑ INSERT 2014/06/20 MIYAMOTO

'↓ INSERT 2014/04/17 MIYAMOTO
'==============================================================================*
'
'        MODULE_NAME      :fncInsAkknTran
'        機能             :自動鍵解除番号発番トラン登録処理
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncInsAkknTran(objDb As Database, _
                                ByVal StrYard As String, _
                                ByVal strCno As String, _
                                ByVal strKagiiCd As String, _
                                ByVal strUserID) As Boolean

    Dim strSQL                  As String
    Dim rsName                  As Recordset
    'Dim rsAkkn                  As Recordset
    Dim rsAkkn                  As Object   'update 2021/06/02 EGL
    Dim blnError                As Boolean
    Dim strNewKaiNo             As String
    Dim dbSQLServer             As Object   'insert 2021/06/02 EGL

On Error GoTo ErrorHandler
    
    blnError = False
    fncInsAkknTran = False
    
    '移動元の鍵区分が自動鍵か確認する
    strSQL = ""
    strSQL = strSQL & "SELECT       NAME_MAST.NAME_VALUE_FROM" & vbCrLf
    strSQL = strSQL & "FROM         NAME_MAST" & vbCrLf
    strSQL = strSQL & "WHERE        NAME_MAST.NAME_ID = '060'" & vbCrLf
    strSQL = strSQL & "AND          NAME_MAST.NAME_CODE = " & Nz(strKagiiCd) & vbCrLf

    Set rsName = objDb.OpenRecordset(strSQL, dbOpenDynaset)

    '見つからない場合エラー
    If rsName.EOF Then
        Call MSZZ024_M10("OpenRecordset", "名称マスタが検索できませんでした。")
    End If
    
    Set dbSQLServer = ADODB_Connection(strBumonCode) 'insert 2021/06/02 EGL

    '自動鍵の場合
'↓ UPDATE 2016/05/17 SUGIMURA
'    If Nz(rsName.Fields("NAME_VALUE_FROM")) = "1" Then
    If MSZZ070_KaiKey(txtYardTihoKbn, Nz(rsName.Fields("NAME_VALUE_FROM"))) = "1" Then
'↑ UPDATE 2016/05/17 SUGIMURA
        '新解除番号の発番
        strNewKaiNo = MSZZ070_M00(strBumonCode, Nz(StrYard), Nz(strCno), Nz(strKagiiCd))

        '自動鍵解除番号発番トランの登録
        strSQL = ""
        strSQL = strSQL & "SELECT      *" & vbCrLf
        strSQL = strSQL & "FROM        AKKN_TRAN" & vbCrLf
        strSQL = strSQL & "WHERE       AKKNT_YARD = " & Nz(StrYard) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_CTNO = " & Nz(strCno) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_KTYPE = '" & Nz(strKagiiCd) & "'" & vbCrLf
        strSQL = strSQL & "AND         AKKNT_STYPE = '4'" 'insert 2021/06/02 EGL
        'Set rsAkkn = objDb.OpenRecordset(strSql, dbOpenDynaset)
        Set rsAkkn = ADODB_Recordset(strSQL, dbSQLServer, adoAppendOnly) 'update 2021/06/02 EGL
        
        With rsAkkn
            .AddNew

            .Fields("AKKNT_YARD") = Nz(StrYard)                                                 ' ヤードコード
            .Fields("AKKNT_CTNO") = Nz(strCno)                                                  ' コンテナ番号
            .Fields("AKKNT_KTYPE") = Nz(strKagiiCd)                                             ' 鍵区分
            .Fields("AKKNT_STYPE") = "4"                                                        ' 鍵区分(4桁) insert 2021/06/02 EGL
            .Fields("AKKNT_HATUD") = Format$(DATE, "YYYYMMDD") & Format$(time, "HHMMSS")        ' 解除番号発番日時
            .Fields("AKKNT_KAINO") = Nz(strNewKaiNo)                                            ' 解除番号
            .Fields("AKKNT_INSED") = Format$(DATE, "YYYYMMDD")                                  ' 作成日付
            .Fields("AKKNT_INSEJ") = Format$(time, "HHMMSS")                                    ' 作成時刻
            .Fields("AKKNT_INSPB") = Me.NAME                                                    ' 作成プログラムID
            .Fields("AKKNT_INSUB") = strUserID                                                  ' 作成ユーザーID

            .UPDATE
        End With
    End If
    
    '↓ insert 2021/06/02 EGL
    If strKagiiCd = 16 Then
        '新解除番号の発番
        strNewKaiNo = MSZZ070_M01(strBumonCode, Nz(strKagiiCd))
        '自動鍵解除番号発番トランの登録
        strSQL = ""
        strSQL = strSQL & "SELECT      *" & vbCrLf
        strSQL = strSQL & "FROM        AKKN_TRAN" & vbCrLf
        strSQL = strSQL & "WHERE       AKKNT_YARD = " & Nz(StrYard) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_CTNO = " & Nz(strCno) & vbCrLf
        strSQL = strSQL & "AND         AKKNT_KTYPE = '" & Nz(strKagiiCd) & "'" & vbCrLf
        strSQL = strSQL & "AND         AKKNT_STYPE = '6'"
        Set rsAkkn = ADODB_Recordset(strSQL, dbSQLServer, adoAppendOnly)
        With rsAkkn
            .AddNew
            .Fields("AKKNT_YARD") = Nz(StrYard)                                                 ' ヤードコード
            .Fields("AKKNT_CTNO") = Nz(strCno)                                                  ' コンテナ番号
            .Fields("AKKNT_KTYPE") = Nz(strKagiiCd)                                             ' 鍵区分
            .Fields("AKKNT_STYPE") = "6"                                                        ' 鍵区分(6桁)
            .Fields("AKKNT_HATUD") = Format$(DATE, "YYYYMMDD") & Format$(time, "HHMMSS")        ' 解除番号発番日時
            .Fields("AKKNT_KAINO") = Nz(strNewKaiNo)                                            ' 解除番号
            .Fields("AKKNT_INSED") = Format$(DATE, "YYYYMMDD")                                  ' 作成日付
            .Fields("AKKNT_INSEJ") = Format$(time, "HHMMSS")                                    ' 作成時刻
            .Fields("AKKNT_INSPB") = Me.NAME                                                    ' 作成プログラムID
            .Fields("AKKNT_INSUB") = strUserID                                                  ' 作成ユーザーID
            .UPDATE
        End With
    End If
    '↑ insert 2021/06/02 EGL
    fncInsAkknTran = True
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsName Is Nothing Then rsName.Close: Set rsName = Nothing
    If Not rsAkkn Is Nothing Then rsAkkn.Close: Set rsAkkn = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncInsAkknTran" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'↑ INSERT 2014/04/17 MIYAMOTO

'==============================================================================*
'
'        MODULE_NAME      :subSetCargStruc
'        機能             :契約トラン作成構造体設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓ UPDATE 2014/06/20 MIYAMOTO
'Private Sub subSetCargStruc(ByRef strucCarg As Type_CARG_FILE, strAcptNo As String)
Private Sub subSetCargStruc(ByRef strucCarg As Type_CARG_FILE, strACPTNO As String, intKeyChgRiyuCd As Integer)
'↑ UPDATE 2014/06/20 MIYAMOTO

    Dim varTemp     As Variant
    Dim intTemp     As Integer
    Dim var月毎割引額 As Variant    '2008/10/01 add
    
    With strucCarg
        .YCODE = txtContainerYcode  'ヤードコード
        .No = txtContainerCno       'コンテナ番号
        .UCODE = cmbUserCode        '顧客コード
        .CAMPC = strCmCampCd        '会社コード 2018/09/25 EGL INS
        
        '解約予定(解約日が入ってる)ならば"02:解約予定"それ以外は"01:契約"
        If Nz(txtReceiptCyDate) = "" Then
            .AGRE = P_AGRE_契約 '契約状態
            .CYDATE = Null      '解約予定日
            
            .KAIUKEDATE = Null  '解約受付日
            .KAITANTO = ""      '解約担当
            .KAICD = P_KAICD_通常   '解約区分
            .KAIRIYUCD = Null       '解約理由
        
        Else
            .AGRE = P_AGRE_解約予定                 '契約状態
            .CYDATE = txtReceiptCyDate              '解約予定日
        
            If Nz(txtMotoKaiUkeDate) <> "" Then
                '移動予約　移動元契約の内容を適用
                .KAIUKEDATE = txtMotoKaiUkeDate     '解約受付日
                .KAITANTO = Nz(txtMotoKaiTanto, "") '解約担当
                .KAICD = txtMotoKaicd               '解約区分
                .KAIRIYUCD = txtMotoKaiRiyuCd       '解約理由
            Else
                '通常予約
                .KAIUKEDATE = DATE                  '解約受付日
                .KAITANTO = cmbReceiptUktanto       '解約担当
                .KAICD = P_KAICD_通常               '解約区分
                .KAIRIYUCD = P_KAIRIYUCD_限定契約   '解約理由
            End If
        End If
        
        .KYDATE = .CYDATE                               '解約日
        
        .FSDATE = DATE                                  '初回契約日
        .STDATE = txtBillStDate                         '契約開始日
        .EDDATE = DateSerial(Format(txtBillStDate, "yyyy") + 1, Format(txtBillStDate, "mm"), 0) '満了日
        
        .DOCU1 = -1                                     '新契約書切替区分
        .DOCU2 = -1                                     '代引き区分
        
        '▼2008/10/01 割引き考慮
        '.RENTKG = txtBillMonthCharge                    '月額料金 = 月額使用料＋毎月雑費計
        '2009/02/13 MOD <S> hirano 割引金額修正　1ヶ月分の割引額をセット
        'If CInt(txtBillPremthSum) > 0 Then
        '    var月毎割引額 = CInt(txt月額分割引合計) / CInt(txtBillPremthSum)
        'Else
        '    var月毎割引額 = txt月額分割引合計
        'End If
        '.RENTKG = txtBillMonthCharge + var月毎割引額
        '2009/11/10 MOD <S> hirano
        '.RENTKG = txtBillMonthCharge + txt月額分割引
        .RENTKG = txtBillMonthCharge + Nz(txt月額分割引, 0)
        '2009/11/10 MOD <E>
        '2009/02/13 MOD <E>
        '▲2008/10/01
        .SYOZEI = 0                                     '消費税
        '2008/10/01 割引き考慮
'        .FRSTKG = txtBillSyoKei                         '初回金額 = 初回使用料＋初回雑費計（全ての毎月雑費日割分＋初回雑費）
        '2009/11/10 MOD <S> hirano
        '.FRSTKG = txtBillSyoKei + txt初回分割引合計     '20008/10/01 new
        .FRSTKG = txtBillSyoKei + Nz(txt初回分割引合計, 0)
        '2009/11/10 MOD <E>
        .FSYOZEI = 0                                    '初回金額消費税
        .SECUKG = Val(Nz(txtBillSecukg, 0)) + Val(Nz(txt保証金割引, 0))  '保証金 + 2011/06/18 tajima
        
        .BIKO = Nz(txtReceiptUknote)                    '備考
        .UPDATE = DATE                                  '更新日
        
        '2006/06/01 SHIBAZAKI 契約トラン作成時に代引伝票送付日に到着日ではなく発送日を設定する
        '.MDATE = fncZeroByteToNull(txtSendArrivalDate)  '代引伝票送付日
        .MDATE = fncZeroByteToNull(txtSendKeySendDate)  '代引伝票送付日
        
        .ACPTNO = strACPTNO                             '受注契約番号
        
        .HOSYICD = Val(cmbContainerHosyicd)             '保証区分
        .HOSYO_CD = IIf(Val(cmbContainerHosyoCompany) = 0, "", Val(cmbContainerHosyoCompany))   '保証会社コード 2009/04/01 Add hirano
        .KAGIICD = Val(cmbKeyKagiicd)                   '鍵区分
        
        .DAHIB = Nz(txtSendSlipNo)                      '代引き伝票番号
        
        '2006/04/28 S.SHIBAZAKI　契約トラン作成時に鍵番号にNULLが入るオペレーションが行われた際、不正終了する不具合修正
'        .KAIJB = txtKeyKagiNo                           '鍵解除番号
        If cmbKeyKagiicd.Column(0) <> 16 Then           'INSERT 2021/06/02 EGL
           .KAIJB = Nz(txtKeyKagiNo, "")                   '鍵解除番号
        End If                                          'INSERT 2021/06/02 EGL
        .HOKAI = 9                                      '保証会社区分
        
        .HOSYD = fncZeroByteToNull(txtSecurityHosyD)    '保証承認日
        .HOSYB = Nz(txtSecurityHosyB)                   '保証承認番号
        .ADMEI = txtReceiptUkkbn                        '広告媒体
        
        .KEITI = 99                                     '契約形態
        .KAGIA = 0                                      '鍵代
        
        .UKNO = txtYoukNo                               '予約受付番号
        .UKTANTO = cmbReceiptUktanto                    '契約受付担当者コード
        '2008/10/01 割引き考慮
'        .FRST_BILL = txtBillGoukei                     '初回請求金額
        '2009/11/10 MOD <S> hirano
        '.FRST_BILL = txtBillGoukei + txt初回分割引合計 + txt月額分割引合計 '2008/10/01 new '初回請求金額
        
        '2017/02/04 M.HONDA UPD
        '.FRST_BILL = txtBillGoukei + Nz(txt初回分割引合計, 0) + Nz(txt月額分割引合計, 0)    '初回請求金額
        .FRST_BILL = txtBillGoukei + Nz(txt初回分割引合計, 0) + Nz(txt月額分割引合計, 0) + Nz(txt年払い, 0)    '初回請求金額
        '2017/02/04 M.HONDA UPD
        
        
        '2009/11/10 MOD <E>
        .PREMTH_SUM = txtBillPremthSum                  '前払月数
        
        'キャンペーン適用満了日
        If Nz(txtContainerSevExpire) <> "" Then
            .CMP_EXDATE = DateSerial(Left(txtContainerSevExpire, 4), Right(txtContainerSevExpire, 2) + 1, 0)
        End If
        
        '2006/06/01 SHIBAZAKI 契約トランに鍵貸し出し本数追加
        .KEY_LENTNUM = fncZeroByteToNull(txtKeyKagiLendnum)
        
        '2010/04/10 K.ISHIZAKA 請求書発行区分追加
        .SEIKI = cmbRcptSeiki.VALUE                                             'INSERT 2010/04/10 K.ISHIZAKA
        
        .USAGE = cmbContainerUsageCode.VALUE '' 2013/07/25 M.HONDA

        .KEY_CHGRIYUCD = intKeyChgRiyuCd                                        'INSERT 2014/06/20 MIYAMOTO
        
        .PLANCD = strプランCd   'Add 2020/09/28 tajima
        
    End With
    

End Sub

'2006/04/28 S.SHIBAZAKI コントロールマスタの更新を一箇所にする
''==============================================================================*
''
''        MODULE_NAME      :fncNewAcptNoAndReqNo
''        機能             :契約番号／請求番号発番
''        IN               :
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncNewAcptNoAndReqNo(objDb As Database, ByRef strNO As String, lngReqNo As Long) As Boolean
'On Error GoTo ErrorHandler
'
'    Dim rsData      As Recordset
'    Dim lngYKNO     As Long       ' 受注契約番号
'    Dim strYY       As String
'    Dim strMM       As String
'    Dim blnError    As Boolean
'
'    blnError = False
'
'    fncNewAcptNoAndReqNo = False
'
'    ' コントロールマスタより取得
'    Set rsData = objDb.OpenRecordset(fncMakeContMastSql(), dbOpenDynaset)
'
'    If rsData.EOF Then
'        Call MSZZ024_M10("OpenRecordset", "コントロールマスタを参照できませんでした。")
'    End If
'
'    '*** データ更新（コントロールマスタ）***
'    With rsData
'        .Edit
'        .Fields("CONT_KYNO") = .Fields("CONT_KYNO") + 1      ' 顧客No
'        .Fields("CONT_REQNO") = .Fields("CONT_REQNO") + 1    ' 請求NO
'        .Fields("CONT_SKDATE") = Format(DATE, "yyyy/mm/dd")  ' 更新日
'        .UPDATE
'        '契約NO
'        lngYKNO = .Fields("CONT_KYNO")
'        '請求NO
'        lngReqNo = .Fields("CONT_REQNO")
'        .Close
'    End With
'    Set rsData = Nothing
'
'    '年取得
'    strYY = ""
'    Select Case Format(txtBillStDate, "yyyy")
'        Case "1998"
'            strYY = "&&"
'        Case "1999"
'            strYY = "**"
'        Case Else
'            If Mid(Trim(Format(Me.txtBillStDate, "yyyy")), 1, 2) = "20" Then
'                strYY = Mid(Trim(Format(Me.txtBillStDate, "yyyy")), 3)
'            End If
'    End Select
'
'    '月取得
'    strMM = Hex(Trim(Format(Me.txtBillStDate, "mm")))
'
'    ' 戻り値｢受注契約番号｣セット
'    strNO = strBumonCode + strYY & strMM & Trim(Format$(lngYKNO, "000000"))
'
'    fncNewAcptNoAndReqNo = True
'    GoTo ExitRtn
'
'ErrorHandler:
'    blnError = True
'
'ExitRtn:
'    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
'    If blnError Then
'        Call Err.Raise(Err.Number, "fncNewAcptNoAndReqNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
'
'End Function

'==============================================================================*
'
'        MODULE_NAME      :fncFirstDay
'        機能             :月初日
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncFirstDay(objTarget As Object) As String

    If Nz(objTarget) = "" Then
        fncFirstDay = ""
    Else
        fncFirstDay = objTarget & "/01"
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :subReferenceMode
'        機能             :画面参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subReferenceMode()

    Dim varFocusControl     As Variant
    
    cmdEnd.SetFocus
    
    '顧客情報は既存受付の場合、参照のみになってるのでここではボタンだけ不可にする
    cmdUserMastMnt.Enabled = False
    
    Call subUserKRenEnable(False)                       'INSERT 2010/04/20 SHIBAZAKI
    Call subContainerReference  'コンテナ情報
    Call subAppFormReference    '申込書情報
    Call subBillReference       '請求情報
    Call subKeyReference        '鍵情報
    Call subSecurityReference   '保証情報
    Call subSendReference       '鍵送付情報
    Call subReceiptReference    '申し込み情報
    
    Call YellowControl(cmbRcptSeiki, True)              'INSERT 2010/04/20 SHIBAZAKI
    Call YellowControl(cmbRcptKaihi, True)              'INSERT 2010/04/20 SHIBAZAKI
    Call YellowControl(cmb_スロープ, True)              'INSERT 2012/04/11 M.HONDA
    
    
    
    Call YellowControl(cmbJIREI, True)              'INSERT 2015/03/24 M.HONDA
    Call YellowControl(cmbSEIKI, True)              'INSERT 2015/03/24 M.HONDA
    Call YellowControl(txtSEIKI, True)              'INSERT 2015/03/24 M.HONDA
    Me.chkNEN_KBN.Enabled = False                   'INSERT 2015/03/24 M.HONDA

    'フォーカス設定用のコントロールを使用不可にする
    For Each varFocusControl In varSetFocus
        varFocusControl.Enabled = False
    Next
    
    '終了ボタン以外を使用不可にする
    '2006/04/27 S.SHIBAZAKI 受付完了後も印刷系のボタンは押下可
'    cmdUkePrint.Enabled = False
'    cmdMouPrint.Enabled = False
'    cmdCancelPrint.Enabled = False
'    cmdKeyPrint.Enabled = False
'    cmdSei2Print.Enabled = False
'    cmdKeyOpenPrint.Enabled = False
'    cmdYardMap.Enabled = False
    cmdRcptCancel.Enabled = False
    cmdSecuNg.Enabled = False
    cmdComplete.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmd_RCPT_SHOUMEI_KDATE.Enabled = False                                      'INSERT 2021/08/18 N.IMAI

    '2006/06/01 SHIBAZAKI セキュリティカード対応
    cmdSCardMnt.Enabled = False
    cmdBillSCardMnt.Enabled = False

    chk保証金強制変更.Enabled = False '2007/02/07 add tajima
    chk割引有効可否.Enabled = False                                             'I:20081001 N.Imai
    cmd割引再計算.Enabled = False                                               'I:20081001 N.Imai
    cmd明細を見る.Enabled = False                                               'I:20081001 N.Imai
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserReference
'        機能             :顧客情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserReference()

    Call YellowControl(fraUserNewExist, True)
    Call YellowControl(txtUserExistKana, True)
    Call YellowControl(cmbUserCode, True)
    cmdUserMastMnt.Enabled = True
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subContainerReference
'        機能             :コンテナ情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subContainerReference()

    Call YellowControl(cmbContainerHosyicd, True)
    Call YellowControl(cmbContainerHosyoCompany, True) '2009/04/01 add hirano
    chkContainerSevExpire.Enabled = False
    Call YellowControl(cmb_荷物移動担当, True)    '2007/04/13 add tajima

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subAppFormReference
'        機能             :申込書情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subAppFormReference()

    Call YellowControl(cmbAppFormSendCd, True)
    Call YellowControl(cmbAppFormButu, True)
    Call YellowControl(txtAppFormButu, True)
    Call YellowControl(txtBillFzappi, True)
    Call YellowControl(txtAppFormLimitDate, True)
    
'    '2006/04/28 S.SHIBAZAKI 乗用具チェック追加。チェックされている場合、申込書に誓約書添付
'    chkAppFormVehicle.Enabled = False
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subBillReference
'        機能             :請求情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subBillReference()

    Call YellowControl(cmbBillMituCd, True)
    Call YellowControl(cmbBillSeikyuCd, True)
    Call YellowControl(txtBillStDate, True)
    
    Call YellowControl(txtBillSecukg, True)
    Call YellowControl(txtBillSumikg, True)
    
    '' 2013/03/28 M.HONDA START
    txtBillSeikyuKikan.BackColor = 16751052
    txtBillEzappi.BackColor = 16764057
    txtBillEzappiDaily.BackColor = 16764057
    ''Call YellowControl(txtBillSeikyuKikan, True)
    ''Call YellowControl(txtBillEzappi, True)
    ''Call YellowControl(txtBillEzappiDaily, True)

    '' 2013/03/28 M.HONDA END
    
    Call YellowControl(txtBillRentkg, True)
    Call YellowControl(txtBillFirstkg, True)
    
    
    Call YellowControl(cmbBillAddEzappiCode1, True)
    Call YellowControl(txtBillAddEzappi1, True)
    Call YellowControl(txtBillAddEzappiDaily1, True)
    Call YellowControl(cmbBillAddEzappiCode2, True)
    Call YellowControl(txtBillAddEzappi2, True)
    Call YellowControl(txtBillAddEzappiDaily2, True)
    
    Call YellowControl(txtBillFzappi, True)
    Call YellowControl(cmbBillAddFzappiCode1, True)
    Call YellowControl(txtBillAddFzappi1, True)
    '2006/06/01 SHIBAZAKI セキュリティカード対応
'    Call YellowControl(cmbBillAddFzappiCode2, True)
    Call YellowControl(txtBillAddFzappi2, True)
    
    Call YellowControl(txt割引適用開始月, True)                                 'I:20081001 N.Imai
    
    chkSecukgKase.Enabled = False                                               'INSERT 2013/04/10 K.ISHIZAKA
    opgRentkg.Enabled = False                                                   'INSERT 2013/04/10 K.ISHIZAKA
    cmdReCal.Enabled = False
    
    Call YellowControl(cmb_スロープ, True)                                      'INSERT 2019/03/08 N.IMAI
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKeyReference
'        機能             :鍵情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKeyReference()

    Call YellowControl(cmbKeyKagiicd, True)
    Call YellowControl(txtKeyKagiNo, True)
    '2006/06/01 SHIBAZAKI セキュリティカード対応↓
'    Call YellowControl(txtKeySecurityCard, True)
    Call YellowControl(txtKeyKagiLendnum, True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSecurityReference
'        機能             :保証情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSecurityReference()

    Call YellowControl(txtSecurityHosyD, True)
    Call YellowControl(txtSecurityHosyB, True)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subSendReference
'        機能             :鍵送付情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subSendReference()

    '鍵送付先黄色
    Call subSendAddressYellow(True)
    
    Call YellowControl(txtSendKeySendDate, True)
    Call YellowControl(txtSendArrivalExpect, True)
    Call YellowControl(cmbSendSpecifyTime, True)
    Call YellowControl(txtSendSlipNo, True)
    Call YellowControl(txtSendArrivalDate, True)
    Call YellowControl(txtSendYamatoDate, True)     'INSERT 2019/03/22 TFC_Kurosu 鍵送付先情報のメール発行日追加対応
    
    cmdUserCopy.Enabled = False
    cmdUserRcopy.Enabled = False
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :subReceiptReference
'        機能             :受付情報参照のみ
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subReceiptReference()

    Call YellowControl(cmbReceiptUktanto, True)
    Call YellowControl(txtReceiptCyDate, True)
    Call YellowControl(cmbReceiptCreason, True)
    Call YellowControl(txtReceiptPaydate, True)
    Call YellowControl(txtReceiptUknote, True)
    Call YellowControl(txtReceiptComment1, True)
    Call YellowControl(txtReceiptComment2, True)
    Call YellowControl(Me.cmb_利用予定期間, True)   '2007/09/25 add tajima
    
    '2020/12/01 [EGL] DEL Start
    'Call YellowControl(txtTpoint, True)             '2013/05/21 INS M.HONDA
    '2020/12/01 [EGL] DEL End
    
    Call YellowControl(cmb_更新区分, True)          '2013/10/16 INS M.HONDA
    Call YellowControl(txt_更新区分, True)          '2013/10/16 INS M.HONDA
    
End Sub

'2006/06/01 SHIBAZAKI 受付完了時に鍵情報をコンテナマスタに更新する
'==============================================================================*
'
'        MODULE_NAME      :fncUpdateCntaMast
'        機能             :コンテナマスタ更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'↓ UPDATE 2014/06/20 MIYAMOTO
'Private Function fncUpdateCntaMast(objDb As Database) As Boolean
Private Function fncUpdateCntaMast(objDb As Database, _
                                   strIdoKbn As String, _
                                   strYcode As String, _
                                   strCno As String, _
                                   strKagiKbn As String) As Boolean
'↑ UPDATE 2014/06/20 MIYAMOTO
On Error GoTo ErrorHandler
    
    Dim strSQL      As String
    Dim rsData      As Recordset
    Dim blnError    As Boolean
    
    blnError = False
    
    fncUpdateCntaMast = False
    
    strSQL = " SELECT * FROM CNTA_MAST "
'↓ UPDATE 2014/06/20 MIYAMOTO
'    strSQL = strSQL & " WHERE CNTA_CODE = " & txtContainerYcode
'    strSQL = strSQL & "   AND CNTA_NO = " & txtContainerCno
    strSQL = strSQL & " WHERE CNTA_CODE = " & strYcode
    strSQL = strSQL & "   AND CNTA_NO = " & strCno
'↑ UPDATE 2014/06/20 MIYAMOTO
    
    Set rsData = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    '見つからない場合エラー
    If rsData.EOF Then
        Call MSZZ024_M10("OpenRecordset", "コンテナマスタが検索できませんでした。")
    End If
    
    With rsData
        .Edit
        
'↓ UPDATE 2014/06/20 MIYAMOTO
'        .Fields("CNTA_KEY_TYPE") = cmbKeyKagiicd
'        .Fields("CNTA_KEY_NO") = txtKeyKagiNo
'        .Fields("CNTA_KEY_NUM") = fncZeroByteToNull(txtKeyKagiLendnum)
        .Fields("CNTA_KEY_TYPE") = strKagiKbn
        If strIdoKbn = "SAKI" Then
            If cmbKeyKagiicd.Column(0) = 16 Then 'insert 2016/06/03 EGL
                .Fields("CNTA_KEY_NO") = txtKeyKagiNo
            End If                               'insert 2016/06/03 EGL
            .Fields("CNTA_KEY_NUM") = fncZeroByteToNull(txtKeyKagiLendnum)
        End If
'↑ UPDATE 2014/06/20 MIYAMOTO
        
        .UPDATE
    End With
    
    fncUpdateCntaMast = True
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If blnError Then
        Call Err.Raise(Err.Number, "fncUpdateCntaMast" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'2006/06/01 SHIBAZAKI セキュリティカード対応
'==============================================================================*
'
'        MODULE_NAME      :fncDispSCard
'        機能             :セキュリティカード情報画面表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncDispSCard(objDb As Database, blnExist As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim intBasicNum     As Integer
    Dim intAddNum       As Integer
    Dim strCardNo       As String
    
    fncDispSCard = False
    
    'セキュリティーカード情報取得
    If blnExist Then
        '既存受付
        Call GetContractSCards(objDb, txtAcptNo, intBasicNum, intAddNum, strCardNo)
    Else
        '新規受付
        Call GetRoomAllotSCards(objDb, txtContainerYcode, txtContainerCno, intBasicNum, intAddNum, strCardNo)
    End If
    
    '値設定
    Call subDispSCardField(intBasicNum, intAddNum, strCardNo)
    
    fncDispSCard = True
    
    Exit Function

ErrorHandler:
    Call Err.Raise(Err.Number, "fncDispSCard" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'2006/06/01 SHIBAZAKI セキュリティカード対応
'==============================================================================*
'
'        MODULE_NAME      :subSCardRefresh
'        機能             :セキュリティカード情報表示更新
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub subSCardRefresh(intBasicNum As Integer, intAddNum As Integer, strCardNo As String, blnUpdate As Boolean)
    
    If blnUpdate Then
        Call subFormRefresh(False, False)
    End If
    
    '値設定
    Call subDispSCardField(intBasicNum, intAddNum, strCardNo)
    
End Sub

'2006/06/01 SHIBAZAKI セキュリティカード対応
'==============================================================================*
'
'        MODULE_NAME      :subDispSCardField
'        機能             :セキュリティカード値表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispSCardField(intBasicNum As Integer, intAddNum As Integer, strCardNo As String)

    lstSCardNo.RowSource = strCardNo
    txtSCardBasicNum = intBasicNum
    txtSCardAddNum = intAddNum
    
    txtBillAddFzappi2.Enabled = intAddNum > 0

End Sub

'2006/09/28 tajima 追加
'==============================================================================*
'
'        MODULE_NAME      :SubRefreshfromFVS90n
'        機能             :修繕関連の再表示
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Sub SubRefreshfromFVS90n()
    
    Dim dbKagtos                As Database
    Dim strKagtosSqlserver      As String
    Dim strKagtosConnect        As String
  
  On Error GoTo ErrorHandler
    
    '加瀬ＤＢに接続
    strKagtosSqlserver = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
    strKagtosConnect = MSZZ007_M00()
    Set dbKagtos = Workspaces(0).OpenDatabase(strKagtosSqlserver, dbDriverNoPrompt, False, strKagtosConnect)
    
    If Nz(cmbUserCode, "") <> "" Then
        '顧客の直近修繕（コメント／クレーム）
        txtUserMntInfo = GetUserLatestKotoText(dbKagtos, strBumonCode, cmbUserCode)
    End If
    
    'コンテナの直近修繕
    txtContainerMntInfo = GetRoomLatestReprText(dbKagtos, strBumonCode, txtContainerYcode, txtContainerCno)
    
    dbKagtos.Close
    Set dbKagtos = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Not dbKagtos Is Nothing Then dbKagtos.Close: Set dbKagtos = Nothing
    Call Err.Raise(Err.Number, "SubRefreshfromFVS90n" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub

'2007/01/12 tajima 追加 New Method!!
'==============================================================================*
'
'        MODULE_NAME      :fncIsPrinting
'        機能             :印刷対象可否
'        IN               :Nothing
'        OUT              :TRUE...印刷可、FALSE...印刷不可
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncIsPrinting() As Boolean

    '2007/11/15 SHIBAZAKI 個室オフィス、バーチャルオフィス追加
    '2009/10/02 レンタル用途33 :コンテナ（バイク） 追加
    '2011/03/14 SHIBAZAKI
    '2011/06/01 M.RYU レンタル用途08：レンタル自習室を追加：P_USAGE_自習室
    '2020/10/15 tajima 6:シェアオフィス追加
    If cmbContainerUsageCode = P_USAGE_レンタルボックス Or _
       cmbContainerUsageCode = P_USAGE_トランクルーム Or _
       cmbContainerUsageCode = P_USAGE_レンタルオフィス Or _
       cmbContainerUsageCode = 9 Or _
       cmbContainerUsageCode = 31 Or _
       cmbContainerUsageCode = 32 Or _
       cmbContainerUsageCode = 33 Or _
       cmbContainerUsageCode = P_USAGE_バイクコンテナ Or _
       cmbContainerUsageCode = P_USAGE_個室オフィス Or _
       cmbContainerUsageCode = P_USAGE_自習室 Or _
       cmbContainerUsageCode = P_USAGE_バーチャルオフィス Or _
       cmbContainerUsageCode = 6 Or _
       cmbContainerUsageCode = P_USAGE_バイク Then      '2007/05/21 fix tajima
       
          fncIsPrinting = True
    Else
      ' 上記以外の用途は印刷出来ない
          MsgBox "このスペースの用途【" & txtContainerUsage & "】では印刷は出来ません。", vbInformation, P_KAKUNIN
          fncIsPrinting = False
    End If

End Function

'↓DELETE 2007/11/15 SHIBAZAKI
''==============================================================================*
''
''        MODULE_NAME      :fncGetIntegrationHelperSeets
''        機能             :印刷対象シート配列取得
''        IN               :intPrintKind 印刷種別
''                         :strHelper ヘルパーファイル名
''        OUT              :印刷対象シート配列
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Function fncGetIntegrationHelperSeets(a押下ボタン As Integer) As Variant
'
'    Dim varArray    As Variant
'
'    Select Case a押下ボタン
'        Case P_PRINT_受付管理表
'            varArray = Array("管理表")
'
'        Case P_PRINT_申込書
'
'          ' システム種別毎に使用するシートを指定
'          Select Case strBumonCode
'              'オフィスの場合
'              Case "1"
'                  If cmbUserKkbn = P_USER_KKBN_法人 Then
'                      varArray = Array("契約の案内", "約款", "申込書法人用（オフィス）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書", "毎月の支払い")
'                  Else
'                      varArray = Array("契約の案内", "約款", "申込書個人用（オフィス）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書", "毎月の支払い")
'                  End If
'              'トランクの場合
'              Case "8"
'                  If cmbUserKkbn = P_USER_KKBN_法人 Then
'                      varArray = Array("契約の案内", "約款", "申込書法人用（コンテナ・トランク・バイク）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書", "毎月の支払い")
'                  Else
'                      varArray = Array("契約の案内", "約款", "申込書個人用（コンテナ・トランク・バイク）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書", "毎月の支払い")
'                  End If
'              ' コンテナの場合
'              Case "H"
'                  If cmbUserKkbn = P_USER_KKBN_法人 Then
'                      varArray = Array("契約の案内", "約款", "申込書法人用（コンテナ・トランク・バイク）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書", "毎月の支払い")
'                  Else
'                      varArray = Array("契約の案内", "約款", "申込書個人用（コンテナ・トランク・バイク）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書", "毎月の支払い")
'                  End If
'              ' 横浜鋼管の場合
'              Case "T"
'                  If cmbUserKkbn = P_USER_KKBN_法人 Then
'                      varArray = Array("契約の案内", "約款", "申込書法人用（コンテナ・トランク・バイク）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書")
'                  Else
'                      varArray = Array("契約の案内", "約款", "申込書個人用（コンテナ・トランク・バイク）", "保証委託契約書", "個人情報取扱", "保証会社個人情報取扱", "明細書")
'                  End If
'          End Select
'
'          'レンタル用途がバイク関連ならば車両収納誓約書を同時に印刷
'          If cmbContainerUsageCode = 31 Or _
'             cmbContainerUsageCode = 32 Or _
'             cmbContainerUsageCode = P_USAGE_バイク Then  '2007/05/21 fix tajma
'              ReDim Preserve varArray(UBound(varArray) + 1)
'              varArray(UBound(varArray)) = "車両収納誓約書"
'          End If
'
'          '解約日が入ってる場合は解約書を同時に印刷
'          If Nz(txtReceiptCyDate) <> "" Then
'              ReDim Preserve varArray(UBound(varArray) + 1)
'              varArray(UBound(varArray)) = "解約通知書"
'          End If
'
'    Case P_PRINT_解約書
'        varArray = Array("解約通知書")
'
'    Case P_PRINT_鍵書類
'
'          ' システム種別毎に使用するシートを指定
'          Select Case strBumonCode
'              'オフィスの場合
'              Case "1"
'                  varArray = Array("明細書控え", "明細書（経理用）", "書類送付の案内", "住所変更", "毎月の支払いについて")
'              'トランクの場合
'              Case "8"
'                  varArray = Array("明細書控え", "明細書（経理用）", "書類送付の案内", "住所変更", "毎月の支払いについて")
'              ' コンテナの場合
'              Case "H"
'                  varArray = Array("明細書控え", "明細書（経理用）", "書類送付の案内", "住所変更", "毎月の支払いについて")
'              ' 横浜鋼管の場合
'              Case "T"
'                  varArray = Array("明細書控え", "明細書（経理用）", "書類送付の案内", "住所変更")
'          End Select
'
'
'        If Nz(cmbKeyKagiicd.Value) = "30" Then
'            '鍵種別が「ドアキー」の場合は「ドアキーの開け方」を同時に印刷
'            ReDim Preserve varArray(UBound(varArray) + 1)
'            varArray(UBound(varArray)) = "ドアキーの開け方"
'        End If
'
'        If Nz(txtReceiptCyDate.Value) = "" Then
'            '解約日が「入っていない」場合は解約書を同時に印刷
'            ReDim Preserve varArray(UBound(varArray) + 1)
'            varArray(UBound(varArray)) = "解約通知書"
'        End If
'
'    Case P_PRINT_請求書2
'            varArray = Array("請求書2")
'    End Select
'
'    fncGetIntegrationHelperSeets = varArray
'
'End Function
'
''==============================================================================*
''
''        MODULE_NAME      :subSetIntegrationBook
''        機能             :統合シート連動（値設定）
''        IN               :印刷種別（押下ボタン）、値設定するシート
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub subSetIntegrationBook(a印刷種別 As Integer, objSeet As Object)
'
'  On Error GoTo Exception
'
'    ' コントロールマスタの値を設定
'    Call subContToIntegrationBook(objSeet)
'    ' KASE-DBの値(振込口座情報)を設定
'    Call subKaseDbToIntegrationBook(objSeet)
'    ' 既存ルーチンを利用して金額関係を設定する
'    Call subSetMousikomiMeisai(objSeet)
'
'    With objSeet
'
'      ' コントロールマスタの部門コードを元にシステム種別を決める
'      ' このシステム種別名はヘルパーと密接な関係を持っているので変更には注意！
'      Select Case strBumonCode
'          'Case "0"
'          '     .Range("システム種別名").Value = "看板"
'          Case "1"
'               .Range("システム種別名").Value = "オフィス"
'          Case "8"
'               .Range("システム種別名").Value = "トランク"
'          Case "H"
'               .Range("システム種別名").Value = "コンテナ"
'          Case "T"
'               .Range("システム種別名").Value = "横浜鋼管"
'          Case Else
'              '上記以外の部門コードは例外終了する
'                Call MSZZ024_M10("subSetIntegrationBook", "部門コード[" & strBumonCode & "]は統合ヘルパーには対応していません。")
'      End Select
'
'      ' 押下ボタンによって印刷種別を設定する
'      ' この印刷種別名はヘルパーの動作に関係するので変更には注意！
'      Select Case a印刷種別
'          Case P_PRINT_申込書
'              .Range("印刷種別").Value = "申込書"
'          Case P_PRINT_解約書
'              .Range("印刷種別").Value = "解約書"
'          Case P_PRINT_鍵書類
'              .Range("印刷種別").Value = "鍵同封書"
'          Case Else
'              '上記以外の印刷種別は例外終了する
'                Call MSZZ024_M10("subSetIntegrationBook", "印刷種別[" & a印刷種別 & "]は統合ヘルパーには対応していません。")
'      End Select
'
'    ' 受付画面の値を設定
'      '★契約情報
'      .Range("ヤード名").Value = txtContainerYname.Value
'      .Range("ヤードコード").Value = txtContainerYcode.Value
'      .Range("ヤード住所").Value = txtContainerYadd.Value
'      .Range("スペースコード").Value = txtContainerCno.Value
'      .Range("スペースサイズ").Value = txtContainerSize.Value
'      ' レンタル用途が コンテナのときのみの上下段を設定する
'      If cmbContainerUsageCode.Value = P_USAGE_レンタルボックス Then
'          .Range("上下段").Value = txtContainerStep.Value
'      Else
'          .Range("上下段").Value = ""
'      End If
'      ' レンタル用途が バイク関連の場合はバイクを設定する 2007/05/21 fix tajima
'      If cmbContainerUsageCode = 31 Or _
'         cmbContainerUsageCode = 32 Or _
'         cmbContainerUsageCode = P_USAGE_バイク Then
'        .Range("レンタル用途").Value = "バイク"
'      Else
'        .Range("レンタル用途").Value = txtContainerUsage.Value
'      End If
'      .Range("鍵種別").Value = txtKeyKagiicd
'      .Range("受付日").Value = txtReceiptUkdate.Value
'      .Range("受付番号").Value = txtRcptNo.Value
'      .Range("解約日").Value = txtReceiptCyDate.Value
'      .Range("受付担当者名").Value = txtReceiptUktanto.Value
'      .Range("契約開始日").Value = txtBillStDate.Value
'      '.Range("契約受付期限日数").Value = コントロールマスタCONT_KEIUKE_TLDを別メソッドで設定済み
'      .Range("契約受付期限日").Value = txtAppFormLimitDate.Value
'      .Range("起算加算日数").Value = "8"
'      .Range("起算日").Value = txtBillKisanDate.Value
'      .Range("初期費用請求方法").Value = txtBillSeikyuCd.Value
'      If Nz(txtBillMituCd.Value, "") <> "" Then '明細種別は未入力のときはシートのデフォルト値を使うように
'        .Range("明細種別").Value = Nz(txtBillMituCd.Value, "")
'      End If
'      ' 支払済月日の設定
'      If Nz(txtBillSeikyuKikan) = "" Then
'          '請求期間が空ならば、起算日を格納する
'          .Range("支払済月日").Value = txtBillKisanDate
'      Else
'          .Range("支払済月日").Value = fncFirstDay(txtBillSeikyuKikan)
'      End If
'      '★契約者情報の設定
'      .Range("契約者郵便番号").Value = txtUserYubino.Value
'      .Range("契約者住所１").Value = txtUserAdr1.Value
'      .Range("契約者住所２").Value = txtUserAdr2.Value & "　" & Nz(txtUserAdr3.Value, "")
'      .Range("契約者名フリガナ").Value = txtUserKana.Value
'      .Range("契約者名").Value = txtUserName.Value
'      '顧客敬称 法人なら｢御中」　個人なら「様」
'      If cmbUserKkbn = P_USER_KKBN_法人 Then
'          .Range("契約者敬称").Value = "御中"
'      Else
'          .Range("契約者敬称").Value = "様"
'      End If
'      .Range("契約代表者名").Value = txtUserTanm  '担当者名
'      .Range("顧客コード").Value = cmbUserCode.Value
'
'      '▼2007/06/09 add tajima
'      .Range("契約者TEL").Value = txtUserTel.Value
'      .Range("契約者FAX").Value = txtUserFax.Value
'      .Range("契約者CEL").Value = txtUserKeitai.Value
'      .Range("契約番号").Value = txtAcptNo.Value
'      .Range("承認番号").Value = txtSecurityHosyB.Value
'      '▲2007/06/09 add tajima
'
'    End With
'
'    '請求明細設定
'    Call subSetSeikyuMeisai(objSeet)
'
'    Exit Sub
'
'Exception:
'    Call Err.Raise(Err.Number, "subSetIntegrationBook" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :subContToIntegrationBook
''        機能             :コントロールマスタから統合ヘルパーに値連動
''        IN               :値設定するシート
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub subContToIntegrationBook(objSeet As Object)
'
'    On Error GoTo ErrorHandler
'
'    Dim dbSqlServer     As Database
'    Dim rsData          As Recordset
'    Dim blnError        As Boolean
'    Dim strSQL          As String
'
'    blnError = False
'
'    'SQLサーバー接続
'    Set dbSqlServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
'
'    strSQL = "SELECT * "
'    strSQL = strSQL & "  FROM CONT_MAST "
'    strSQL = strSQL & " WHERE CONT_KEY = " & P_CONT_KEY
'    Set rsData = dbSqlServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
'
'    If rsData.EOF Then
'        'データが存在しない
'        'データを検索できないのは何らかの問題が発生しているはずなので、不正終了する
'        Call MSZZ024_M10("OpenRecordset", "データを参照できませんでした。")
'    End If
'
'    'データ設定
'    With objSeet
'        .Range("貸主会社名").Value = rsData.Fields("CONT_KAISYA")
'        .Range("貸主郵便番号").Value = rsData.Fields("CONT_YUBINO")
'        .Range("貸主住所１").Value = rsData.Fields("CONT_ADDR_1")
'        .Range("貸主住所２").Value = rsData.Fields("CONT_ADDR_2")
'        .Range("貸主TEL").Value = rsData.Fields("CONT_TEL_NO")
'        .Range("貸主FAX").Value = rsData.Fields("CONT_FAX_NO")
'        .Range("契約受付期限日数").Value = rsData.Fields("CONT_KEIUKE_TLD")
'    End With
'
'    GoTo ExitRtn
'
'ErrorHandler:
'    blnError = True
'
'ExitRtn:
'    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
'    If Not dbSqlServer Is Nothing Then dbSqlServer.Close: Set dbSqlServer = Nothing
'
'    If blnError Then
'        Call Err.Raise(Err.Number, "subContToIntegrationBook" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
'
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :subKaseDbToIntegrationBook
''        機能             :加瀬DBから統合ヘルパーに値連動
''        IN               :値設定するシート
''        OUT              :
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub subKaseDbToIntegrationBook(objSeet As Object)
'
'    On Error GoTo ErrorHandler
'
'    Dim dbSqlServer     As Database
'    Dim rsData          As Recordset
'    Dim blnError        As Boolean
'    Dim strKaseSqlSvr   As String
'    Dim strKaseConnect  As String
'
'    blnError = False
'
'    'SQLサーバー接続
'    'SQL-Server名
'    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))
'    If strKaseSqlSvr = "" Then
'        'テーブル[SETU_TABL]の設定不正
'        Call MSZZ024_M10("DLookup(""SETUT_SETUN"", ""SETU_TABL"", ""SETUT_SETUB='ODBC_DATA_SOURCE_NAME')", "テーブル[SETU_TABL]の設定不正。")
'    End If
'
'    '接続文字列取得
'    strKaseConnect = MSZZ007_M00()
'    If strKaseConnect = "" Then
'        'テーブル[SETU_TABL]の設定不正
'        Call MSZZ024_M10("MSZZ007_M00", "テーブル[SETU_TABL]の設定不正。")
'    End If
'    Set dbSqlServer = Workspaces(0).OpenDatabase(strKaseSqlSvr, dbDriverNoPrompt, False, strKaseConnect)
'
'    Set rsData = dbSqlServer.OpenRecordset(fncMakeKagtosSql(), dbOpenSnapshot, dbSQLPassThrough)
'    '                                         ↑請求書２と同じＳＱＬ
'    If rsData.EOF Then
'        'データが存在しない
'        'データを検索できないのは何らかの問題が発生しているはずなので、不正終了する
'        Call MSZZ024_M10("OpenRecordset", "データを参照できませんでした。")
'    End If
'
'    'データ設定
'    objSeet.Range("金融機関名") = rsData.Fields("金融機関名")
'    objSeet.Range("金融機関支店名") = rsData.Fields("支店名")
'    objSeet.Range("預金項目") = rsData.Fields("預金種別")
'    objSeet.Range("口座番号") = rsData.Fields("口座番号")
'    objSeet.Range("口座名義") = rsData.Fields("口座名義人")
'
'    GoTo ExitRtn
'
'ErrorHandler:
'    blnError = True
'
'ExitRtn:
'    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
'    If Not dbSqlServer Is Nothing Then dbSqlServer.Close: Set dbSqlServer = Nothing
'
'    If blnError Then
'        Call Err.Raise(Err.Number, "subKaseDbToIntegrationBook" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If
'
'End Sub
'
''==============================================================================*
''
''        MODULE_NAME      :fncIsRogoPrint
''        機能             :加瀬ロゴの印刷可否判定
''        IN               :Nothing...必要な情報はグローバル変数を参照
''        OUT              :TRUE...出力する、FALSE...出力しない
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Function fncIsRogoPrint() As Boolean
'
'    ' システムが横浜鋼管ならばFALSE
'    If strBumonCode = "T" Then
'        fncIsRogoPrint = False
'
'    ' 上記以外ならTRUE
'    Else
'        fncIsRogoPrint = True
'    End If
'
'End Function
'↑DELETE 2007/11/15 SHIBAZAKI

'==============================================================================*
'
'        MODULE_NAME      :chk保証金強制変更_AfterUpdate
'        機能             :chk保証金強制変更_AfterUpdate イベント
'        IN               :
'        OUT              :
'        2007/02/07 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk保証金強制変更_AfterUpdate()
    'チェックされたら保証金を入力可にする
    Call YellowControl(txtBillSecukg, Not chk保証金強制変更.VALUE)
    Call YellowControl(txtBillSecukgKase, Not chk保証金強制変更.VALUE)          'INSERT 2013/04/10 K.ISHIZAKA
    '▼2008/12/04 hirano Add 保証金変更フラグ保存に伴う追加
    Call subEditMode(True)
    '▲2008/12/04 hirano Add
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_スロープ_AfterUpdate
'        機能             :cmb_スロープ_AfterUpdate イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_スロープ_AfterUpdate()

    txt_スロープ.VALUE = fncAfterUpdateCommon(cmb_スロープ, P_DIVISION_LENGTH, P_CONTROL_数値)
        
    '' 2012/04/24 M.HONDA
    If cmb_スロープ = "01" Then
    
        '' 2012/06/30 M.HONDA
        '' 地方物件の際には10,000円とする(頭二桁が11～19まで)
        If (Left(txtContainerYcode, 2) >= 11 And Left(txtContainerYcode, 2) <= 19) Then
        
            'txtBillFzappi.VALUE = 10000 2014/02/28 M.HONDA DEL 消費税対応
            'txtBillFzappi.VALUE = 10800 '2014/02/28 M.HONDA INS 消費税対応
            txtBillFzappi.VALUE = 11000  '2019/09/26 M.HONDA
            
            
        Else
            'txtBillFzappi.VALUE = 2000  2014/02/28 M.HONDA DEL 消費税対応
            'txtBillFzappi.VALUE = 2160 '2014/02/28 M.HONDA INS 消費税対応
            'txtBillFzappi.VALUE = 3240  '2017/09/26 M.HONDA INS
            txtBillFzappi.VALUE = 3300   '2019/09/26 M.HONDA
            
        End If
        'INSERT 2019/02/06 N.IMAI Start
        If strBumonCode = "H" And lngFZAPPI_CODE = 0 Then
            lngFZAPPI_CODE = 95
        End If
        'INSERT 2019/02/06 N.IMAI End
    Else
        txtBillFzappi.VALUE = 0
        'INSERT 2019/02/06 N.IMAI Start
        If strBumonCode = "H" Then
            lngFZAPPI_CODE = 0
        End If
    End If
   

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_スロープ_KeyPress
'        機能             :cmb_スロープ_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_スロープ_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmb_スロープ, KeyAscii, P_DIVISION_LENGTH)
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_利用予定期間_KeyPress
'        機能             :cmb_利用予定期間_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_スロープ_BeforeUpdate(Cancel As Integer)
    Call fncBeforeUpdateCommon(cmb_スロープ, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_荷物移動担当_AfterUpdate
'        機能             :cmb_荷物移動担当_AfterUpdate イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_荷物移動担当_AfterUpdate()

   txt_荷物移動担当.VALUE = fncAfterUpdateCommon(cmb_荷物移動担当, P_DIVISION_LENGTH, P_CONTROL_数値)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_荷物移動担当_BeforeUpdate
'        機能             :cmb_荷物移動担当_BeforeUpdate イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_荷物移動担当_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(cmb_荷物移動担当, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_荷物移動担当_KeyPress
'        機能             :cmb_荷物移動担当_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_荷物移動担当_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Num1(cmb_荷物移動担当, KeyAscii, P_DIVISION_LENGTH)
    
End Sub
'==============================================================================*
'
'        MODULE_NAME      :isネット予約担当
'        機能             :受付担当がネット予約担当可否＆処理
'        2007/08/23 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function isネット予約担当()

    If Me.cmbReceiptUktanto.VALUE = 888 Then
        MsgBox "受付担当を変えて下さい。", vbInformation, P_KAKUNIN
        Me.cmbReceiptUktanto.SetFocus
        isネット予約担当 = True
    Else
        isネット予約担当 = False
    End If

End Function

'★★★★★★★★★★★★★★ 2007/09/25 add tajima ★★★★★★★★★★★★★★
'==============================================================================*
'
'        MODULE_NAME      :cmb_利用予定期間_KeyPress
'        機能             :cmb_利用予定期間_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用予定期間_AfterUpdate()
    txt_利用予定期間.VALUE = fncAfterUpdateCommon(cmb_利用予定期間, P_DIVISION_LENGTH, P_CONTROL_数値)
   
    ' ネット予約の場合
    If blnネット予約 = True Then
        ' 利用予定期間が二ヶ月以下ならば割引無し
        If Me.cmb_利用予定期間.VALUE = "01" Then
            Me.cmbBillAddEzappiCode1.VALUE = Null
            Me.txtContainerSevExpire.VALUE = ""
        '2009/09/11 DEL <S> hirano
        '' それ以外はネット割引サービス適応割引
        'Else
        '    Me.cmbBillAddEzappiCode1.VALUE = 57
        '    Me.txtBillAddEzappi1 = -2000
        '    Me.txtBillAddEzappiDaily1 = fncDailyRate(txtBillAddEzappi1)
        '    Me.txtContainerSevExpire = Format(DateSerial(Format(txtBillKisanDate, "yyyy"), Format(txtBillKisanDate, "m") + Val(Nz(txtSevExmonth, 0)), 1), "yyyy/mm")
        '2009/09/11 DEL <E> hirano
        End If
        Call cmbBillAddEzappiCode1_AfterUpdate
   End If
   
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_利用予定期間_KeyPress
'        機能             :cmb_利用予定期間_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用予定期間_BeforeUpdate(Cancel As Integer)
    Call fncBeforeUpdateCommon(cmb_利用予定期間, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_利用予定期間_KeyPress
'        機能             :cmb_利用予定期間_KeyPress イベント
'        2007/04/13 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_利用予定期間_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(cmb_利用予定期間, KeyAscii, P_DIVISION_LENGTH)
End Sub


'==============================================================================*
'
'        MODULE_NAME      :get毎月雑費計
'        機能             :txtBillEzappiKei のコントロールソースの代わり
'        2008/03/23 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function get毎月雑費計()

    Dim ret As Long
    
  On Error GoTo Exception
    ret = 0
    
    If Nz(Me.txtBillEzappi, "") <> "" Then
        ret = CLng(Me.txtBillEzappi)
    End If

    If Nz(Me.txtBillAddEzappi1, "") <> "" Then
        ret = ret + CLng(Me.txtBillAddEzappi1)
    End If

    If Nz(Me.txtBillAddEzappi2, "") <> "" Then
        ret = ret + CLng(Me.txtBillAddEzappi2)
    End If

    get毎月雑費計 = ret
    
    Exit Function

Exception:
    Call Err.Raise(Err.Number, "get毎月雑費計" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Function

'==============================================================================*
'
'        MODULE_NAME      :txt割引適用開始月_AfterUpdate
'        機能             :txt割引適用開始月_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt割引適用開始月_AfterUpdate()

    txt割引適用開始月 = fncAfterUpdateCommon(txt割引適用開始月, P_YEARMONTH_LENGTH, P_CONTROL_年月)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt割引適用開始月_BeforeUpdate
'        機能             :txt割引適用開始月_BeforeUpdate イベント
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt割引適用開始月_BeforeUpdate(Cancel As Integer)

    Call fncBeforeUpdateCommon(txt割引適用開始月, P_YEARMONTH_LENGTH, P_CONTROL_年月, Cancel)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt割引適用開始月_GotFocus
'        機能             :txt割引適用開始月_GotFocus イベント
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt割引適用開始月_GotFocus()

    txt割引適用開始月 = Format(fncFirstDay(txt割引適用開始月), "yy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt割引適用開始月_KeyPress
'        機能             :txt割引適用開始月_KeyPress イベント
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt割引適用開始月_KeyPress(KeyAscii As Integer)

    Call CmKeyPress_For_Date(txt割引適用開始月, KeyAscii, P_YEARMONTH_LENGTH)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :txt割引適用開始月_LostFocus
'        機能             :txt割引適用開始月_LostFocus イベント
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub txt割引適用開始月_LostFocus()

    txt割引適用開始月 = Format(fncFirstDay(txt割引適用開始月), "yyyy/mm")

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd割引再計算_Click
'        機能             :cmd割引再計算_Click イベント
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd割引再計算_Click()
    Dim str_txt初回分割引合計 As String
    Dim str_txt月額分割引合計 As String
    Dim str_txt保証金割引     As String '2011/06/18 add tajima
    '2009/02/03 ADD Start hirano
    str_txt初回分割引合計 = CStr(txt初回分割引合計)
    str_txt月額分割引合計 = CStr(txt月額分割引合計)
    '2009/02/03 ADD End hirano
    str_txt保証金割引 = CStr(txt保証金割引) '2011/06/18 add tajima
    
    Call sub割引再計算(P_MODE_仮計算)
    
    '2009/02/03 ADD Start hirano
    If Not (str_txt初回分割引合計 = CStr(txt初回分割引合計) _
    And str_txt月額分割引合計 = CStr(txt月額分割引合計)) _
    And str_txt保証金割引 = CStr(txt保証金割引) Then        '2011/06/18 chg&add tajima
        blnSave割引再計算 = True                                                'I:20090223 N.Imai
        Call subEditMode(True)
    End If
    '2009/02/03 ADD End hirano
End Sub

'==============================================================================*
'
'        MODULE_NAME      :chk割引有効可否_Click
'        機能             :chk割引有効可否_Click イベント
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub chk割引有効可否_Click()

    cmd割引再計算.Enabled = chk割引有効可否.VALUE
    '▼ 2008/11/19 hirano add 明細を見るボタン、割引適用開始月制御追加
    cmd明細を見る.Enabled = chk割引有効可否.VALUE
    'txt割引適用開始月.Enabled = chk割引有効可否.VALUE
    '▲ 2008/11/19 hirano
    Call sub割引再計算(P_MODE_仮計算)
    Call subEditMode(True)

    ''＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠
    If chkContainerSevExpire.Visible = True Then
        If chk割引有効可否.VALUE = True Then
            chkContainerSevExpire.VALUE = True
            chkContainerSevExpire_AfterUpdate
        Else
            chkContainerSevExpire.VALUE = False
            chkContainerSevExpire_AfterUpdate
        End If
    End If
    ''＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠＠

End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmd明細を見る_Click
'        機能             :cmd明細を見る_Click イベント
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd明細を見る_Click()

    Dim blnMode As Boolean
    '▼ 2008/11/19 hirano del 保存を促し　必要であれば保存させる。
    ''無理やり編集中ではない状態にする
    'If lblNewExist.Caption = P_LABEL_既存受付 Then
    '    blnMode = blnEditMode
    '    blnEditMode = False
    'End If
    '▲ 2008/11/19 hirano del
    MSZZ039.P_CMD明細 = 1
    Call cmdMouPrint_Click
    MSZZ039.P_CMD明細 = 0
    
    '▼ 2008/11/19 hirano del 保存を促し　必要であれば保存させる。
    ''元に戻す
    'If lblNewExist.Caption = P_LABEL_既存受付 Then
    '    blnEditMode = blnMode
    'End If
    '▲ 2008/11/19 hirano del
    
End Sub
''==============================================================================*
''
''        MODULE_NAME      :sub割引再計算
''        機能             :割引再計算処理
''        2008/10/01 new add
''
''==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'Private Sub sub割引再計算(intMode As Integer)
'
'    Dim strDate             As String
'    Dim strTo               As String
'    Dim lngKingaku          As Long
'    Dim lngSetKingaku       As Long
'    Dim intCount            As Integer
'    Dim blnCalc             As Boolean
'    Dim str割引適用開始月   As String
'    Dim strBillKisanDate    As String
'    Dim strDcntNo()         As String
'
'    On Error GoTo Exception
'
'    doCmd.Hourglass True
'
'    '初期値設定
'    ReDim wk割付情報(0)
'    wk割付情報(0).件数 = 0
'    ReDim strDcntNo(0)
'    txt初回分割引合計 = 0
'    txt月額分割引合計 = 0
'    '▼2008/12/04 Add hirano 保証料は、基本料＋毎月雑費（初期化）
'    If Not chk保証金強制変更.VALUE Then
'        txtBillSecukg = txtBillRentkg + get毎月雑費計
'    End If
'    '▲2008/12/04 Add hirano
'    str割引適用開始月 = txt割引適用開始月
'    strBillKisanDate = txtBillKisanDate
'    If Nz(str割引適用開始月) = "" Then
'        str割引適用開始月 = Left$(txtBillStDate, 7)
'    End If
'
'    If chk割引有効可否.VALUE And str割引適用開始月 <> "" And strBillKisanDate <> "" Then
'        '割引有効のみ計算
'        lngSetKingaku = 0
'        If Left$(strBillKisanDate, 7) = str割引適用開始月 Then
'            lngSetKingaku = txtBillFirstkg
'        Else
'            lngSetKingaku = txtBillRentkg
'        End If
'        '初回分割引合計取得
'        txt初回分割引合計 = 0
'        intCount = 0
'        ReDim wk割付情報(0)
'        strDate = Left$(Replace(strBillKisanDate, "/", ""), 6)
'        If str割引適用開始月 = Left$(strBillKisanDate, 7) Then
'            '起算日の年月＝割引適用開始月の場合のみ算出
'            If Replace(str割引適用開始月, "/", "") <> strDate Then
'                strDate = str割引適用開始月
'            End If
'            If intMode = P_MODE_本計算 Then
'                '2008/12/20 Start
'                'blnCalc = MSZZ045_fncGetTekiyo(wk割付情報(0), txtYoukNo, strDate, strBumonCode, lngDCMax)
'                blnCalc = MSZZ045_fncGetTekiyo(wk割付情報(0), txtAcptNo, strDate, strBumonCode, lngDCMax)
'                '2008/12/20 End
'            Else
'                '2008/12/20 Start
'                'blnCalc = MSZZ045_fncGetKariNebiki(wk割付情報(0), txtYoukNo, strDate, txtContainerYcode, txtContainerCno, lngSetKingaku, strBillKisanDate, str割引適用開始月, strBumonCode, lngDCMax, Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), strDcntNo, blnネット予約)
'                blnCalc = MSZZ045_fncGetKariNebiki(wk割付情報(0), txtAcptNo, strDate, txtContainerYcode, txtContainerCno, lngSetKingaku, strBillKisanDate, str割引適用開始月, strBumonCode, lngDCMax, Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), strDcntNo, blnネット予約)
'                '2008/12/20 End
'            End If
'            If blnCalc Then
'                If wk割付情報(0).件数 > 0 Then
'                    If wk割付情報(0).設定金額 = 0 Then
'                        txt初回分割引合計 = wk割付情報(0).割引合計
'                    Else
'                        txt初回分割引合計 = wk割付情報(0).設定金額 - txtBillFirstkg
'                    End If
'                End If
'            End If
'        End If
'        '月額分割引合計取得
'        strDate = Format$(DateAdd("M", 1, CDate(Left$(strDate, 4) & "/" & Mid$(strDate, 5, 2) & "/01")), "yyyyMM")
'        strTo = Replace(txtBillSeikyuKikan, "/", "")
'        Do Until (strDate > strTo)
'            intCount = intCount + 1
'            ReDim Preserve wk割付情報(intCount)
'            If intMode = P_MODE_本計算 Then
'                '2008/12/20 Start
'                'blnCalc = MSZZ045_fncGetTekiyo(wk割付情報(intCount), txtYoukNo, strDate, strBumonCode, lngDCMax)
'                blnCalc = MSZZ045_fncGetTekiyo(wk割付情報(intCount), txtAcptNo, strDate, strBumonCode, lngDCMax)
'                '2008/12/20 End
'            Else
'                '2008/12/20 Start
'                'blnCalc = MSZZ045_fncGetKariNebiki(wk割付情報(intCount), txtYoukNo, strDate, txtContainerYcode, txtContainerCno, txtBillRentkg, strBillKisanDate, str割引適用開始月, strBumonCode, lngDCMax, Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), strDcntNo, blnネット予約)
'                blnCalc = MSZZ045_fncGetKariNebiki(wk割付情報(intCount), txtAcptNo, strDate, txtContainerYcode, txtContainerCno, txtBillRentkg, strBillKisanDate, str割引適用開始月, strBumonCode, lngDCMax, Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), strDcntNo, blnネット予約)
'                '2008/12/20 End
'            End If
'            If blnCalc Then
'                If wk割付情報(intCount).件数 > 0 Then
'                    If wk割付情報(intCount).設定金額 = 0 Then
'                        lngKingaku = wk割付情報(intCount).割引合計 + lngKingaku
'                    Else
'                        lngKingaku = wk割付情報(intCount).設定金額 - txtBillRentkg + lngKingaku
'                    End If
'                End If
'            Else
'            End If
'            strDate = Format$(DateAdd("M", 1, CDate(Left$(strDate, 4) & "/" & Mid$(strDate, 5, 2) & "/01")), "yyyyMM")
'        Loop
'        txt月額分割引合計 = lngKingaku
'        '▼2008/12/04 Add hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。(chk保証金強制変更がチェックされていない場合)
'        If Not chk保証金強制変更.VALUE Then
'            txtBillSecukg = txtBillRentkg + get毎月雑費計 + lngKingaku
'        End If
'        '▲2008/12/04 Add hirano
'    End If
'
'Exception:
'
'    If Err <> 0 Then
'        Call MSZZ024_M00("sub割引再計算", True)
'    End If
'
'    doCmd.Hourglass False
'
'End Sub

'==============================================================================*
'
'        MODULE_NAME      :sub割引再計算
'        機能             :割引再計算処理
'        2008/12/20 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub割引再計算(intMode As Integer)

    Dim i                   As Integer
    Dim strDate             As String
    Dim strTo               As String
    Dim lngKingaku          As Long
    Dim intCount            As Integer
    Dim blnCalc             As Boolean
    Dim str割引適用開始月   As String
    Dim strBillKisanDate    As String
    
    Dim chkForcing          As Integer   ''INS 2013/10/24 M.HONDA
    
    On Error GoTo Exception

    chkForcing = Me.chkCampaign.VALUE    ''INS 2013/10/24 M.HONDA


    ' 起算日(契約開始日）が未入力ならば計算はしない
    If Nz(txtBillKisanDate, "") = "" Then Exit Sub

    doCmd.Hourglass True

    '初期値設定
    txt初回分割引合計 = 0
    txt月額分割引合計 = 0
    txt保証金割引 = 0
    '▼2008/12/19 Add hirano 月額追加　保証料対象割引の値がセットされる
    txt月額分割引 = 0
    '▲2008/12/19 Add hirano
    '▼2008/12/20 Add hirano 保証料は、月額（初期化）
    If Not chk保証金強制変更.VALUE Then
        'del 2018/10/02 EGL -->txtBillSecukg = Nz(txtBillMonthCharge, 0)
        Call subCompSecukgKase                                                  'INSERT 2013/04/10 K.ISHIZAKA
    End If
    '▲2008/12/20 Add hirano
    str割引適用開始月 = Replace(txt割引適用開始月, "/", "")
    strBillKisanDate = Replace(txtBillKisanDate, "/", "")
        
    If Nz(str割引適用開始月) = "" Then
        str割引適用開始月 = strBillKisanDate
    Else
        If Left$(str割引適用開始月, 6) = Left$(strBillKisanDate, 6) Then
            str割引適用開始月 = strBillKisanDate
        Else
            str割引適用開始月 = str割引適用開始月 & "01"
        End If
    End If
     
    ReDim wk割付情報(0)
    wk割付情報(0).割引合計 = 0
    wk割付情報(0).件数 = 0
    wk割付情報(0).設定金額 = 0
    wk割付情報(0).年月 = ""
            If chk割引有効可否.VALUE And str割引適用開始月 <> "" And strBillKisanDate <> "" Then
        If intMode = P_MODE_本計算 Then
            'データをDBより取得
            blnCalc = MSZZ045_fncGetDCRA_TRAN(wk割引情報, txtAcptNo, strBumonCode)
        Else
            'データを再計算
'            blnCalc = MSZZ045_fncGetNebikiData(wk割引情報, _       '---UPDATE 2011/08/25 M.RYU----<S>
'                                                txtContainerYcode, _
'                                                txtContainerCno, _
'                                                Nz(txtBillFirstkg, 0), _
'                                                Nz(txtBillRentkg, 0), _
'                                                txtBillKisanDate, _
'                                                Left$(str割引適用開始月, 6), _
'                                                strBumonCode, _
'                                                lngDCMax, _
'                                                Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), _
'                                                blnネット予約)
'            blnCalc = MSZZ045_fncGetNebikiData(wk割引情報, _
'                                                txtContainerYcode, _
'                                                txtContainerCno, _
'                                                Nz(txtBillFirstkg, 0), _
'                                                Nz(txtBillRentkg, 0), _
'                                                Nz(txtReceiptUkdate, 0), _
'                                                Nz(txtBillKisanDate, 0), _
'                                                Left$(str割引適用開始月, 6), _
'                                                strBumonCode, _
'                                                lngDCMax, _
'                                                Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), _
'                                                blnネット予約)      '---UPDATE 2011/08/25 M.RYU----<E>
            '---UPDATE 2012/10/24 M.HONDA----<S>
            blnCalc = MSZZ045_fncGetNebikiData(wk割引情報, _
                                                txtContainerYcode, _
                                                txtContainerCno, _
                                                Nz(txtBillFirstkg, 0), _
                                                Nz(txtBillRentkg, 0), _
                                                Nz(txtReceiptUkdate, 0), _
                                                Nz(txtBillKisanDate, 0), _
                                                Left$(str割引適用開始月, 6), _
                                                strBumonCode, _
                                                lngDCMax, _
                                                Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), _
                                                blnネット予約, _
                                                chkForcing)      '---UPDATE 2012/10/24 M.HONDA----<E>



        End If
        If blnCalc = True Then
            '月額分割引合計取得
            strDate = str割引適用開始月
            If txtBillSeikyuKikan = "" Then
                strTo = Left$(Replace(txtBillKisanDate, "/", ""), 6)
            Else
                strTo = Replace(txtBillSeikyuKikan, "/", "")
            End If
            intCount = 0
            Do Until (Left$(strDate, 6) > strTo)
                ReDim Preserve wk割付情報(intCount)
                If MSZZ045_fncNebikiCalc(wk割付情報(intCount), wk割引情報, strDate, Left$(str割引適用開始月, 6)) Then
                    If wk割付情報(intCount).件数 > 0 Then
'                        If Left$(strDate, 6) = Left$(str割引適用開始月, 6) _
'                        And Left$(str割引適用開始月, 6) = Left$(strBillKisanDate, 6) Then
'                            If wk割付情報(0).設定金額 = 0 Then
'                                txt初回分割引合計 = wk割付情報(intCount).割引合計
'                            Else
'                                txt初回分割引合計 = wk割付情報(intCount).設定金額
'                            End If
'                        Else
'                            If wk割付情報(intCount).設定金額 = 0 Then
'                                lngKingaku = wk割付情報(intCount).割引合計 + lngKingaku
'                            Else
'                                lngKingaku = wk割付情報(intCount).設定金額 + lngKingaku
'                            End If
'                        End If
                        '2017/07/06 M.HONDA INS
                        If Left$(strDate, 6) = Left$(str割引適用開始月, 6) _
                        And Left$(str割引適用開始月, 6) = Left$(strBillKisanDate, 6) Then
                            If wk割付情報(0).設定金額 = 0 Then
                                txt初回分割引合計 = wk割付情報(intCount).割引合計
                            Else
                                If txtContainerYcode = "MHNIKK" And Right$(strBillKisanDate, 2) = "01" Then
                                    txt初回分割引合計 = wk割付情報(intCount).設定金額
                                ElseIf txtContainerYcode = "MHNIKK" Then
                                    txt初回分割引合計 = 0
                                Else
                                    txt初回分割引合計 = wk割付情報(intCount).設定金額
                                End If
                                '2017/07/03 M.HODNA INS
                            End If
                        Else
                            If wk割付情報(intCount).設定金額 = 0 Then
                                lngKingaku = wk割付情報(intCount).割引合計 + lngKingaku
                            Else
                                If txtContainerYcode = "MHNIKK" And Right$(strBillKisanDate, 2) = "01" Then
                                    lngKingaku = 0
                                Else
                                    lngKingaku = wk割付情報(intCount).設定金額 + lngKingaku
                                End If
                            End If
                        End If
                        '2017/07/06 M.HONDA INS
                    End If
                End If
                strDate = Format$(DateAdd("M", 1, CDate(Left$(strDate, 4) & "/" & Mid$(strDate, 5, 2) & "/01")), "yyyyMM")
                intCount = intCount + 1
            Loop
            txt月額分割引合計 = lngKingaku
            '▼2008/12/04 Add hirano 保証料は、基本料＋毎月雑費から月額分割割引合計をマイナスする。(chk保証金強制変更がチェックされていない場合)
            '2009/02/02 hirano del Start txt月額分割引再計算
            'If Not chk保証金強制変更.VALUE Then
            '2009/02/02 hirano del End
            '初期化
            ReDim wk割付情報2(0)
            wk割付情報2(0).割引合計 = 0
            wk割付情報2(0).件数 = 0
            wk割付情報2(0).設定金額 = 0
            wk割付情報2(0).年月 = ""
            intCount = 0
            ReDim Preserve wk割付情報2(intCount)
            If MSZZ045_fncNebikiMonCalc(wk割付情報2(intCount), wk割引情報) Then
                If wk割付情報2(intCount).件数 > 0 Then
                  For i = 0 To wk割付情報2(intCount).件数 - 1
                    txt月額分割引 = txt月額分割引 + wk割付情報2(intCount).TRAN_DATA(i).VALUE
                  Next
                End If
            End If
            '2009/02/02 hirano Ass Start txt月額分割引再計算
            If Not chk保証金強制変更.VALUE Then
            '2009/02/02 hirano Add End
                '▼2008/12/20 Add hirano 保証料は、毎月額から月額割引をマイナスする。(chk保証金強制変更がチェックされていない場合)
                'txtBillSecukg = txtBillRentkg + get毎月雑費計 + lngKingaku
                'txtBillSecukg = txtBillRentkg + get毎月雑費計 + (lngKingaku / IIf(txtBillPremthSum = 0, 1, txtBillPremthSum))
                'del 2018/10/02 EGL -->txtBillSecukg = txtBillMonthCharge + txt月額分割引
                Call subCompSecukgKase                                          'INSERT 2013/04/10 K.ISHIZAKA
                '▲2008/12/20 Add hirano
            End If
            '▲2008/12/04 Add hirano
        End If
    
        ' ネット予約の場合には、保険料相当額を含んだ金額の割引となる
        '▼2013/04/02 M.HONDA UPDATE
        ' 割引再計算時の保証委託料の金額を月額固定の金額で計算させる。
        ' 保険料相当額は保証委託料割引の対象外となるため。
        '▼2013/01/17 M.HONDA UPDATE
        '▼2011/06/18 Add tajima
        Dim BillSecukg As Long
        
        If blnネット予約 = True Then
           BillSecukg = txtBillSecukg
        Else
            If txtEzappiCode = "68" Then
               BillSecukg = txtBillSecukg - Nz(txtBillEzappi, 0)
            Else
               BillSecukg = txtBillSecukg
            End If
        End If
            
        Call sub雑費割引再計算(intMode, BillSecukg)
        
        '' Call sub雑費割引再計算(intMode, txtBillSecukg)
        '▲2011/16/18 Add tajima
        '▲2013/01/17 M.HONDA UPDATE
        
        '----↓↓↓↓----20110829--M.HONDA--Ins-------↓↓↓↓---<s>
        '' 事務手数料割引取得
        ''Me.txtBillAddFzappi1.VALUE = getFzappi1(Me.cmb_利用予定期間.VALUE)
        '----↑↑↑↑----20110829--M.HONDA--Ins-------↑↑↑↑---<e>
        
    Else
        wk割引情報.件数 = 0
    End If

     Call subHoshouShow(0)  '2017/12/20 add
    
Exception:

    If Err <> 0 Then
        Call MSZZ024_M00("sub割引再計算", True)
    End If

    doCmd.Hourglass False

End Sub

'==============================================================================*
'
'        MODULE_NAME      :fnc適用割引割付更新
'        機能             :適用割引割付更新
'        2008/10/01 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'▼2008/11/19 MOD hirano モード追加
'Private Function fnc適用割引割付更新(objDb As Database, strUserId As String) As Boolean
Private Function fnc適用割引割付更新(objDb As Database, strUserID As String, Optional intMode As Integer = 0) As Boolean
'▲2008/11/19 MOD
    
    On Error GoTo Exception
    
    Dim i           As Integer
    Dim strSQL      As String
    Dim rsData      As Recordset
    Dim blnError    As Boolean
    Dim strTo       As String
    Dim lngTotal    As Long                                                     'I:2008/12/22
    Dim chkForcing          As Integer   ''INS 2013/10/24 M.HONDA
    
    blnError = False
    
    chkForcing = Me.chkCampaign.VALUE    ''INS 2013/10/24 M.HONDA
    fnc適用割引割付更新 = False
    lngTotal = 0                                                                'I:2008/12/22
    
    '一旦削除
    strSQL = "SELECT * FROM DCRA_TRAN WHERE DCRAT_ACPTNO = '" & txtAcptNo & "'"
    Set rsData = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    Do Until (rsData.EOF = True)
        rsData.Delete
        rsData.MoveNext
    Loop
    '▼2008/11/19 ADD hirano モード追加 保証会社NGの場合は、削除する。
    If Not intMode = P_SAVEMODE_保証会社NG Then
    '▲2008/11/19 ADD
        If chk割引有効可否.VALUE Then
            'データ追加
'            If MSZZ045_fncGetNebikiData(wk割引情報, _                  '---UPDATE 2011/08/25 M.RYU----<S>
'                                        txtContainerYcode, _
'                                        txtContainerCno, _
'                                        txtBillFirstkg, _
'                                        Nz(txtBillRentkg, 0), _
'                                        Nz(txtBillKisanDate, 0), _
'                                        txt割引適用開始月, _
'                                        strBumonCode, _
'                                        lngDCMax, _
'                                        Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), _
'                                        blnネット予約) = True Then
'            If MSZZ045_fncGetNebikiData(wk割引情報, _
'                                        txtContainerYcode, _
'                                        txtContainerCno, _
'                                        txtBillFirstkg, _
'                                        Nz(txtBillRentkg, 0), _
'                                        Nz(txtReceiptUkdate, 0), _
'                                        Nz(txtBillKisanDate, 0), _
'                                        txt割引適用開始月, _
'                                        strBumonCode, _
'                                        lngDCMax, _
'                                        Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), _
'                                        blnネット予約) = True Then      '---UPDATE 2011/08/25 M.RYU----<E>
                                        
            If MSZZ045_fncGetNebikiData(wk割引情報, _
                                        txtContainerYcode, _
                                        txtContainerCno, _
                                        txtBillFirstkg, _
                                        Nz(txtBillRentkg, 0), _
                                        Nz(txtReceiptUkdate, 0), _
                                        Nz(txtBillKisanDate, 0), _
                                        txt割引適用開始月, _
                                        strBumonCode, _
                                        lngDCMax, _
                                        Nz(fncZeroByteToNull(Me.cmb_利用予定期間), 0), _
                                        blnネット予約, _
                                        chkForcing) = True Then      '---UPDATE 2011/08/25 M.RYU----<E>
                                        
                    For i = 0 To wk割引情報.件数 - 1
                    '2008/12/22 Start
                    strTo = ""
                    '初月分
                    If Left$(wk割引情報.TRAN_DATA(i).RLDNT_FROM, 6) = Left$(Replace(txtBillKisanDate, "/", ""), 6) Then
                        strTo = Replace(DateAdd("d", -1, DateAdd("M", 1, Left$(wk割引情報.TRAN_DATA(i).RLDNT_FROM, 4) & "/" & Mid$(wk割引情報.TRAN_DATA(i).RLDNT_FROM, 5, 2) & "/01")), "/", "") 'UPD 2017/07/06 M.HONDA
                        If (txtContainerYcode = "MHNIKK" And Right$(Replace(txtBillKisanDate, "/", ""), 2) = "01") Then 'INS 2017/07/06 M.HONDA
                            '浜松町は1日起算日の場合のみ初回の割引を作成、1日起算日意外は割引を作成しない。
                            rsData.AddNew
                            rsData.Fields("DCRAT_ACPTNO") = txtAcptNo                   'I:2008/12/20
                            rsData.Fields("DCRAT_DCNT_NO") = wk割引情報.TRAN_DATA(i).RLDNT_NO
                            rsData.Fields("DCRAT_ENABLE") = wk割引情報.TRAN_DATA(i).RLDNT_ENABLE
                            rsData.Fields("DCRAT_FROM") = wk割引情報.TRAN_DATA(i).RLDNT_FROM
                            rsData.Fields("DCRAT_TO") = strTo
                            rsData.Fields("DCRAT_PRICE") = wk割引情報.TRAN_DATA(i).VALUE
                            rsData.Fields("DCRAT_TEXT") = wk割引情報.TRAN_DATA(i).RLDNT_TEXT
                            rsData.Fields("DCRAT_SEV1N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV1N '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_SEV2N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV2N '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_SEV3N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV3N '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_ENDEN") = wk割引情報.TRAN_DATA(i).DCRAT_ENDEN '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_EXMONTH") = wk割引情報.TRAN_DATA(i).DCRAT_EXMONTH '2015/10/14 M.HONDA INS
                            rsData.Fields("DCRAT_IYAKU_SEIKYU") = "0"
                            rsData.Fields("DCRAT_SEIKYU_KBN") = "0"
                            rsData.Fields("DCRAT_INSED") = Format$(DATE, "yyyymmdd")
                            rsData.Fields("DCRAT_INSEJ") = Format$(time, "hhmmss")
                            rsData.Fields("DCRAT_INSPB") = Me.NAME
                            rsData.Fields("DCRAT_INSUB") = strUserID
                            rsData.UPDATE
                        ElseIf txtContainerYcode <> "MHNIKK" Then
                            rsData.AddNew
                            rsData.Fields("DCRAT_ACPTNO") = txtAcptNo                   'I:2008/12/20
                            rsData.Fields("DCRAT_DCNT_NO") = wk割引情報.TRAN_DATA(i).RLDNT_NO
                            rsData.Fields("DCRAT_ENABLE") = wk割引情報.TRAN_DATA(i).RLDNT_ENABLE
                            rsData.Fields("DCRAT_FROM") = wk割引情報.TRAN_DATA(i).RLDNT_FROM
                            rsData.Fields("DCRAT_TO") = strTo
                            rsData.Fields("DCRAT_PRICE") = wk割引情報.TRAN_DATA(i).VALUE
                            rsData.Fields("DCRAT_TEXT") = wk割引情報.TRAN_DATA(i).RLDNT_TEXT
                            rsData.Fields("DCRAT_SEV1N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV1N '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_SEV2N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV2N '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_SEV3N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV3N '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_ENDEN") = wk割引情報.TRAN_DATA(i).DCRAT_ENDEN '2015/09/28 M.HONDA INS
                            rsData.Fields("DCRAT_EXMONTH") = wk割引情報.TRAN_DATA(i).DCRAT_EXMONTH '2015/10/14 M.HONDA INS
                            rsData.Fields("DCRAT_IYAKU_SEIKYU") = "0"
                            rsData.Fields("DCRAT_SEIKYU_KBN") = "0"
                            rsData.Fields("DCRAT_INSED") = Format$(DATE, "yyyymmdd")
                            rsData.Fields("DCRAT_INSEJ") = Format$(time, "hhmmss")
                            rsData.Fields("DCRAT_INSPB") = Me.NAME
                            rsData.Fields("DCRAT_INSUB") = strUserID
                            rsData.UPDATE
                        End If
                    End If
                    '翌月以降分
                    If strTo <> wk割引情報.TRAN_DATA(i).RLDNT_TO Then
                        If wk割引情報.TRAN_DATA(i).DCNTM_TYPE = 1 Or lngDCMax >= -(lngTotal + wk割引情報.TRAN_DATA(i).VALUE_TRUE) Then
                            If (txtContainerYcode = "MHNIKK" And Right$(Replace(txtBillKisanDate, "/", ""), 2) <> "01") Then
                                 '浜松町は1日起算日以外の場合のみ翌月の割引を作成する。
                                rsData.AddNew
                                rsData.Fields("DCRAT_ACPTNO") = txtAcptNo                   'I:2008/12/20
                                rsData.Fields("DCRAT_DCNT_NO") = wk割引情報.TRAN_DATA(i).RLDNT_NO
                                rsData.Fields("DCRAT_ENABLE") = wk割引情報.TRAN_DATA(i).RLDNT_ENABLE
                                If strTo <> "" Then
                                    rsData.Fields("DCRAT_FROM") = Replace(DateAdd("M", 1, Left$(wk割引情報.TRAN_DATA(i).RLDNT_FROM, 4) & "/" & Mid$(wk割引情報.TRAN_DATA(i).RLDNT_FROM, 5, 2) & "/01"), "/", "")
                                Else
                                    rsData.Fields("DCRAT_FROM") = wk割引情報.TRAN_DATA(i).RLDNT_FROM
                                End If
                                rsData.Fields("DCRAT_TO") = wk割引情報.TRAN_DATA(i).RLDNT_TO
                                    rsData.Fields("DCRAT_PRICE") = wk割引情報.TRAN_DATA(i).VALUE_TRUE
                                rsData.Fields("DCRAT_TEXT") = wk割引情報.TRAN_DATA(i).RLDNT_TEXT
                                rsData.Fields("DCRAT_SEV1N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV1N '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_SEV2N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV2N '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_SEV3N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV3N '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_ENDEN") = wk割引情報.TRAN_DATA(i).DCRAT_ENDEN '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_EXMONTH") = wk割引情報.TRAN_DATA(i).DCRAT_EXMONTH '2015/10/14 M.HONDA INS
                                rsData.Fields("DCRAT_IYAKU_SEIKYU") = "0"
                                rsData.Fields("DCRAT_SEIKYU_KBN") = "1"
                                rsData.Fields("DCRAT_INSED") = Format$(DATE, "yyyymmdd")
                                rsData.Fields("DCRAT_INSEJ") = Format$(time, "hhmmss")
                                rsData.Fields("DCRAT_INSPB") = Me.NAME
                                rsData.Fields("DCRAT_INSUB") = strUserID
                                rsData.UPDATE
                                If wk割引情報.TRAN_DATA(i).DCNTM_TYPE <> 1 Then
                                    lngTotal = lngTotal + wk割引情報.TRAN_DATA(i).VALUE_TRUE
                                End If
                            ElseIf txtContainerYcode <> "MHNIKK" Then
                                rsData.AddNew
                                rsData.Fields("DCRAT_ACPTNO") = txtAcptNo                   'I:2008/12/20
                                rsData.Fields("DCRAT_DCNT_NO") = wk割引情報.TRAN_DATA(i).RLDNT_NO
                                rsData.Fields("DCRAT_ENABLE") = wk割引情報.TRAN_DATA(i).RLDNT_ENABLE
                                If strTo <> "" Then
                                    rsData.Fields("DCRAT_FROM") = Replace(DateAdd("M", 1, Left$(wk割引情報.TRAN_DATA(i).RLDNT_FROM, 4) & "/" & Mid$(wk割引情報.TRAN_DATA(i).RLDNT_FROM, 5, 2) & "/01"), "/", "")
                                Else
                                    rsData.Fields("DCRAT_FROM") = wk割引情報.TRAN_DATA(i).RLDNT_FROM
                                End If
                                rsData.Fields("DCRAT_TO") = wk割引情報.TRAN_DATA(i).RLDNT_TO
                                    rsData.Fields("DCRAT_PRICE") = wk割引情報.TRAN_DATA(i).VALUE_TRUE
                                rsData.Fields("DCRAT_TEXT") = wk割引情報.TRAN_DATA(i).RLDNT_TEXT
                                rsData.Fields("DCRAT_SEV1N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV1N '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_SEV2N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV2N '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_SEV3N") = wk割引情報.TRAN_DATA(i).DCRAT_SEV3N '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_ENDEN") = wk割引情報.TRAN_DATA(i).DCRAT_ENDEN '2015/09/28 M.HONDA INS
                                rsData.Fields("DCRAT_EXMONTH") = wk割引情報.TRAN_DATA(i).DCRAT_EXMONTH '2015/10/14 M.HONDA INS
                                rsData.Fields("DCRAT_IYAKU_SEIKYU") = "0"
                                rsData.Fields("DCRAT_SEIKYU_KBN") = "1"
                                rsData.Fields("DCRAT_INSED") = Format$(DATE, "yyyymmdd")
                                rsData.Fields("DCRAT_INSEJ") = Format$(time, "hhmmss")
                                rsData.Fields("DCRAT_INSPB") = Me.NAME
                                rsData.Fields("DCRAT_INSUB") = strUserID
                                rsData.UPDATE
                                If wk割引情報.TRAN_DATA(i).DCNTM_TYPE <> 1 Then
                                    lngTotal = lngTotal + wk割引情報.TRAN_DATA(i).VALUE_TRUE
                                End If
                            End If
                        End If
                    End If
                    '2008/12/22 End
                Next i
            End If
            '▼2011/06/18 ADD tajima 雑費値引の実行
            Call sub雑費適用割引割付更新(rsData, strUserID)
            '▲2011/06/18 ADD
            End If
    '▼2008/11/19 ADD hirano モード追加
    End If
    '▲2008/11/19 ADD

Exception:

    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fnc適用割引割付更新" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :仕入情報取得
'        MODULE_ID        :funcGetShiireInfo
'        Parameter        :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(String)  = a部門コード
'                         :第3引数(Recordset) = 取得情報
'                         :第4引数(Recordset) = 新規変更
'        Retrun           :保証会社情報の取得
'        NOTE             :
'        CREATE_DATE      :2009/04/01
'        UPDATE_DATE      :
'==============================================================================*
''---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function funcGetShiireInfo(dbSqlServerKONT As DAO.Database, _
                                   dbSqlServerkase As DAO.Database, _
                                    a部門コード As String, _
                                     arsData As Recordset, ablnRcpt As Boolean) As Boolean
    Dim strSQL     As String
    Dim returnText As String
    Dim blnDataGet As Boolean 'データ取得フラグ
    Dim rsObject   As Recordset
    Dim strHOSHO_KBN As String          '保証区分
    Dim strHOSHOKAISHA_CD As String     '保証会社コード
    '2009/10/23 MOD <S> hirano
    Dim blnNetHosho As Boolean          '保証会社マスタ適用
    Dim intNetHoshoFlg As Integer       '保証会社切替フラグ
    Dim strNetHoshoCD As String         'ネット予約統一保証会社コード
    '2009/10/23 MOD <E> hirano
    funcGetShiireInfo = False
    blnDataGet = False
    On Error GoTo Exception
    If Not ablnRcpt Then
        '新規
        If Nz(arsData.Fields("UCODE"), "") <> "" Then '顧客コードがない場合は､新規顧客として過去情報を取得しない
        
        
            '' 2012/04/12 M.HONDA START
            '' 自社理由移動・金額変更の場合には保証会社を旧契約から引き継ぐ
            ''If Nz(arsData.Fields("YOUKT_MOVEKBN"), "") = "1" Or Nz(arsData.Fields("YOUKT_MOVEKBN"), "") = "5" Then
            ''2018/03/16 M.HONDA UPD
            '2018/07/17 M.HONDA UPD
            '自社理由移動・金額変更、未使用移動の場合には保証会社を旧契約から引き継ぐ
            If Nz(arsData.Fields("YOUKT_MOVEKBN"), "") = "1" Or Nz(arsData.Fields("YOUKT_MOVEKBN"), "") = "2" Or Nz(arsData.Fields("YOUKT_MOVEKBN"), "") = "5" Or Nz(arsData.Fields("YOUKT_MOVEKBN"), "") = "3" Then
            '2018/07/17 M.HONDA UPD
            
                strSQL = "select top 1 CARG_HOSYICD ,CARG_HOSYO_CD ,CARG_AGRE,CARG_STDATE "
                strSQL = strSQL & "from CARG_FILE where CARG_ACPTNO = '" & Nz(arsData.Fields("MOTO_CARG_ACPTNO"), "") & "' "        'INSERT 2023/10/04 N.IMAI
'                strSQL = strSQL & "and (CARG_HOSYO_CD !='970025' or '" & Nz(txtMoveKbn, "") & "' = '" & P_MOVEKBN_自社理由 & "') "  'INSERT 2023/09/06 N.IMAI
'                strSQL = strSQL & "and CARG_HOSYO_CD !='970025' "                  'INSERT 2023/09/11 N.IMAI
                'strSQL = strSQL & "and CARG_HOSYO_CD NOT IN ('970025','970170') "   'INSERT 2023/09/21 N.IMAI
                strSQL = strSQL & "and (CARG_HOSYO_CD NOT IN ('970025','970170') or '" & Nz(txtMoveKbn, "") & "' = '" & P_MOVEKBN_自社理由 & "') "  'INSERT 2023/09/26 N.IMAI
                ' 組み立てたSQLでデータ取得
                Set rsObject = dbSqlServerKONT.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
                If rsObject.EOF = False Then
                    With rsObject
                        strHOSHO_KBN = Nz(.Fields("CARG_HOSYICD"), "")          '保障区分
                        strHOSHOKAISHA_CD = Nz(.Fields("CARG_HOSYO_CD"), "")    '保障会社コード
                        'INSERT 2023/05/01 N.IMAI Start
                        If strHOSHO_KBN = "2" Then
                            strHOSHO_KBN = "1"
                            strHOSHOKAISHA_CD = ""
                        End If
                        'INSERT 2023/05/01 N.IMAI End
                        blnDataGet = True
                    End With
                End If
                rsObject.Close
                Set rsObject = Nothing
            End If
            '' 2012/04/12 M.HONDA END
            '2017/12/20 M.HONDA DEL
            If Not blnDataGet Then
                '顧客コードで契約トランを検索　契約状態が００，０９，１０　以外　保証区分、保証会社コードを取得
                strSQL = "select top 1 CARG_HOSYICD ,CARG_HOSYO_CD ,CARG_AGRE,CARG_STDATE "
                strSQL = strSQL & "from CARG_FILE where CARG_UCODE = '" & Nz(arsData.Fields("UCODE"), "") & "' and "    'UCODE 顧客コード
                'strSQL = strSQL & "(CARG_HOSYO_CD !='970025' or '" & Nz(txtMoveKbn, "") & "' = '" & P_MOVEKBN_自社理由 & "') and "  'INSERT 2023/09/06 N.IMAI
                'strSQL = strSQL & " CARG_HOSYO_CD !='970025' and "                 'INSERT 2023/09/11 N.IMAI
                'strSQL = strSQL & " CARG_HOSYO_CD NOT IN ('970025','970170') and "  'INSERT 2023/09/21 N.IMAI
                strSQL = strSQL & " (CARG_HOSYO_CD NOT IN ('970025','970170') or '" & Nz(txtMoveKbn, "") & "' = '" & P_MOVEKBN_自社理由 & "') and " 'INSERT 2023/09/26 N.IMAI
                strSQL = strSQL & "CARG_AGRE  not in (0,9,10) " '申込,完了,申込キャンセル　以外のデータ
                strSQL = strSQL & "order by CARG_STDATE desc"

                ' 組み立てたSQLでデータ取得
                Set rsObject = dbSqlServerKONT.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
                If rsObject.EOF = False Then
                    With rsObject
                        '保証会社無し対応----------------------------        'INSERT EGL 2020/02/25 ここから
                        '元契約が保証会社無(自社契約)の場合は､
                        '保証区分を"1"に設定し、その顧客の最新の保証会社コードを引き継ぐ。
                        If Nz(arsData.Fields("YOUKT_MOVEKBN"), "0") = 0 Then
                            If Nz(.Fields("CARG_HOSYICD"), "") = "2" Then
                                strHOSHO_KBN = "2"
                                blnDataGet = True
                                GoTo lbl_Exit:
                            End If
                        End If
                        '---------------------------------------------       'INSERT EGL 2020/01/24 ここまで
                        strHOSHO_KBN = Nz(.Fields("CARG_HOSYICD"), "")          '保障区分
                        strHOSHOKAISHA_CD = Nz(.Fields("CARG_HOSYO_CD"), "")    '保障会社コード
                        'INSERT 2023/05/01 N.IMAI Start
                        If strHOSHO_KBN = "2" Then
                            strHOSHO_KBN = "1"
                            strHOSHOKAISHA_CD = ""
                        End If
                        'INSERT 2023/05/01 N.IMAI End
                        blnDataGet = True
                    End With
                    'ﾘﾌﾟﾗｽは、初期値にしない
                    Select Case strHOSHO_KBN
                    Case P_HOSYICD_リプラス, P_HOSYICD_リプラス包括
                        '2009/04/30 MOD <S> HIRANO 自社理由、金額変更移動の場合　リプラスを引き継ぐ
                        'blnDataGet = False
                        Select Case Nz(arsData.Fields("YOUKT_MOVEKBN"), "")
    '                    Case P_MOVEKBN_自社理由, P_MOVEKBN_金額変更                            'DELETE 2011/09/27 M.RYU
                        Case P_MOVEKBN_自社理由, P_MOVEKBN_金額変更, P_MOVEKBN_紹介者金額変更   'INSERT 2011/09/27 M.RYU
                            blnDataGet = True
                        Case Else
                            blnDataGet = False
                        End Select
                        '2009/04/30 MOD <S>
                    Case P_HOSYICD_保証会社有
                        If strHOSHOKAISHA_CD = "970025" Or strHOSHOKAISHA_CD = "970035" Or strHOSHOKAISHA_CD = "970170" Then
                            blnDataGet = True
                        Else
                            '各保証会社毎に纏める。
                            If strHOSHOKAISHA_CD = "970020" Then
                                strHOSHOKAISHA_CD = "970025"
                                blnDataGet = True
                            ElseIf strHOSHOKAISHA_CD = "970030" Or strHOSHOKAISHA_CD = "970031" Or strHOSHOKAISHA_CD = 970032 Or strHOSHOKAISHA_CD = 970033 Then
                                strHOSHOKAISHA_CD = "970035"
                                blnDataGet = True
                            Else
                                blnDataGet = False
                            End If
                        End If
                    '2020/02/25 strat 自社トランク(2)で特定の移動理由の場合
                    Case P_HOSYICD_保証会社無_自社
                        Select Case Nz(arsData.Fields("YOUKT_MOVEKBN"), "")
                            Case P_MOVEKBN_名義変更, P_MOVEKBN_保障会社変更, P_MOVEKBN_紹介者金額変更, P_MOVEKBN_満了更新
                                blnDataGet = False 'falseにすることにより後述で保証会社ありにかえる
                        End Select
                    '2020/02/25 end
                    End Select
                End If
lbl_Exit: 'INSERT EGL 2020/02/25 保証会社無し対応
                rsObject.Close
                Set rsObject = Nothing
            End If


            '顧客コードで受付トランを検索　保証区分、保証会社コードを取得
'            If Not blnDataGet Then
'                strSql = "SELECT TOP 1 RCPT_HOSYICD,RCPT_HOSYO_CD FROM RCPT_TRAN "
'                strSql = strSql & "INNER JOIN YOUK_TRAN ON ( YOUKT_UKNO = RCPT_NO) "
'                strSql = strSql & "WHERE RCPT_UCODE = '" & Nz(arsData.Fields("UCODE"), "") & "' "
'                strSql = strSql & "AND YOUKT_YUKBN <> '53' "    '受付キャンセル以外
'                strSql = strSql & "order by RCPT_UKDATE desc "
'                ' 組み立てたSQLでデータ取得
'                Set rsObject = dbSqlServerKONT.OpenRecordset(strSql, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
'                If rsObject.EOF = False Then
'                    With rsObject
'                        strHOSHO_KBN = Nz(.Fields("RCPT_HOSYICD"), "")          '保障区分
'                        strHOSHOKAISHA_CD = Nz(.Fields("RCPT_HOSYO_CD"), "")    '保障会社コード
'                    End With
'                    'ﾘﾌﾟﾗｽは、初期値にしない
'                    Select Case strHOSHO_KBN
'                    Case P_HOSYICD_リプラス, P_HOSYICD_リプラス包括
'                        blnDataGet = False
'                    Case Else
'                        blnDataGet = True
'                    End Select
'                End If
'                rsObject.Close
'                Set rsObject = Nothing
'            End If
'            '2017/12/20 M.HONDA DEL
            
        End If
        'ヤードマスタをヤードコードで検索 保証会社コードを取得
        If Not blnDataGet Then
            strSQL = "SELECT YARD_HOSYO_CD  from yard_mast where yard_code = '" & Nz(arsData.Fields("YCODE"), "") & "' "
            ' 組み立てたSQLでデータ取得
            Set rsObject = dbSqlServerKONT.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
            If rsObject.EOF = False Then
                With rsObject
                    strHOSHO_KBN = "1"                                      '保障区分 1固定
                    strHOSHOKAISHA_CD = Nz(.Fields("YARD_HOSYO_CD"), "")    '保障会社コード
                End With
                blnDataGet = True
            End If
            rsObject.Close
            Set rsObject = Nothing
        End If
        
        'ネット予約か保証区分が８の場合    パルマデータの取得　保証区分：１　保証会社
        '2009/10/23 MOD <S> hirano 保証会社自動切替対応
        'If (Nz(arsData.Fields("YOUKT_GENKBN"), 0) = 1) Or (Not blnDataGet And (Nz(arsData.Fields("HOSYICD"), "") = "8" Or strHOSHO_KBN = "8")) Then
        If (Nz(arsData.Fields("YOUKT_GENKBN"), 0) = 1) Then
          If fncGetNetHosyoInfo(dbSqlServerkase, blnNetHosho, intNetHoshoFlg, strNetHoshoCD) Then
            If blnNetHosho = True Then
                If intNetHoshoFlg = 0 Then
                    'ネット予約統一
                    strHOSHO_KBN = "1"
                    strHOSHOKAISHA_CD = strNetHoshoCD
                ElseIf intNetHoshoFlg = 1 Then
                    'ヤード保証会社
                    strSQL = "SELECT YARD_HOSYO_CD  from yard_mast where yard_code = '" & Nz(arsData.Fields("YCODE"), "") & "' "
                    ' 組み立てたSQLでデータ取得
                    Set rsObject = dbSqlServerKONT.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
                    If rsObject.EOF = False Then
                        With rsObject
                            strHOSHO_KBN = "1"                                      '保障区分 1固定
                            strHOSHOKAISHA_CD = Nz(.Fields("YARD_HOSYO_CD"), "")    '保障会社コード
                        End With
                        blnDataGet = True
                    End If
                    rsObject.Close
                    Set rsObject = Nothing
                    'ヤードマスタに未設定の場合、ネット予約統一保証会社コードをセットする。
                    If strHOSHOKAISHA_CD = "" Then strHOSHOKAISHA_CD = strNetHoshoCD
                End If
            Else
                '移行前
                strHOSHOKAISHA_CD = ""
            End If
          End If
          If (strHOSHOKAISHA_CD = "") Then
        '2009/10/23 MOD <E> hirano
            strSQL = "SELECT SHIRM_SHIRC FROM SHIR_MAST "
            strSQL = strSQL & "WHERE SHIRM_BUMOC = '" & a部門コード & "' "
            strSQL = strSQL & "AND SHIRM_SHIRN LIKE '%パルマ%' "
            ' 組み立てたSQLでデータ取得
            Set rsObject = dbSqlServerkase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
            If rsObject.EOF = False Then
                With rsObject
                    strHOSHO_KBN = "1"                                      '保障区分 1固定
                    strHOSHOKAISHA_CD = Nz(.Fields("SHIRM_SHIRC"), "")    '保障会社コード
                End With
                blnDataGet = True
            End If
            rsObject.Close
            Set rsObject = Nothing
          '2009/10/23 MOD <S> hirano 保証会社自動切替対応
          End If
          '2009/10/23 MOD <E> hirano
        End If
        
        If blnDataGet = True Then
            '取得された場合
            cmbContainerHosyicd = IIf(strHOSHO_KBN = "", "", Format(strHOSHO_KBN, "00"))
            Call cmbContainerHosyicd_AfterUpdate    '保証区分によるコントロール表示変更
            cmbContainerHosyoCompany = strHOSHOKAISHA_CD
            Call cmbContainerHosyoCompany_AfterUpdate    '保証会社コードによるコントロール表示変更
            '2009/04/30 INS <S> hirano
            If strHOSHOKAISHA_CD <> "" And txtContainerHosyoCampany = "" Then
                strSQL = "SELECT SHIRM_SHIRN FROM SHIR_MAST "
                strSQL = strSQL & "WHERE SHIRM_SHIRC = '" & strHOSHOKAISHA_CD & "' "
                ' 組み立てたSQLでデータ取得
                Set rsObject = dbSqlServerkase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
                If rsObject.EOF = False Then
                    With rsObject
                        txtContainerHosyoCampany = Nz(.Fields("SHIRM_SHIRN"), "")    '保障会社コード
                    End With
                End If
                rsObject.Close
                Set rsObject = Nothing
            End If
            '2009/04/30 INS <E> hirano
        End If
    Else
        '変更
        '2009/04/23 ADD <S>
        'ネット予約か保証区分が８の場合    パルマデータの取得　保証区分：１　保証会社
        If (Nz(arsData.Fields("YOUKT_GENKBN"), 0) = 1) And Nz(arsData.Fields("HOSYICD"), "") = "8" Then
            strSQL = "SELECT SHIRM_SHIRC FROM SHIR_MAST "
            strSQL = strSQL & "WHERE SHIRM_BUMOC = '" & a部門コード & "' "
            strSQL = strSQL & "AND SHIRM_SHIRN LIKE '%パルマ%' "
            ' 組み立てたSQLでデータ取得
            Set rsObject = dbSqlServerkase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
            If rsObject.EOF = False Then
                With rsObject
                    strHOSHO_KBN = "1"                                      '保障区分 1固定
                    strHOSHOKAISHA_CD = Nz(.Fields("SHIRM_SHIRC"), "")    '保障会社コード
                End With
                blnDataGet = True
            End If
            rsObject.Close
            Set rsObject = Nothing
            blnネット予約初回 = True        '2009/04/23 ネット予約初回を意味する
        End If
        If blnDataGet Then
            '取得された場合
            cmbContainerHosyicd = IIf(strHOSHO_KBN = "", "", Format(strHOSHO_KBN, "00"))
            Call cmbContainerHosyicd_AfterUpdate    '保証区分によるコントロール表示変更
            cmbContainerHosyoCompany = strHOSHOKAISHA_CD
            Call cmbContainerHosyoCompany_AfterUpdate    '保証会社コードによるコントロール表示変更
        Else
        '2009/04/23 ADD <E>
            strSQL = "SELECT SHIRM_SHIRN FROM SHIR_MAST "
            strSQL = strSQL & "WHERE SHIRM_SHIRC = '" & Nz(arsData.Fields("HOSYO_KAISHA_CD"), "") & "' "
            ' 組み立てたSQLでデータ取得
            Set rsObject = dbSqlServerkase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
            If rsObject.EOF = False Then
                With rsObject
                    txtContainerHosyoCampany = Nz(.Fields("SHIRM_SHIRN"), "")    '保障会社コード
                End With
            End If
            rsObject.Close
            Set rsObject = Nothing
        End If '2009/04/23 ADD
    End If
    funcGetShiireInfo = True
    Exit Function
                              
Exception:
  
  If Not rsObject Is Nothing Then rsObject.Close: Set rsObject = Nothing
  Call Err.Raise(Err.Number, "GetRoomLatestReprText" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)


End Function
'==============================================================================*
'
'       MODULE_NAME       : ネット予約保証会社情報取得
'       MODULE_ID         : fncGetNetHosyoInfo
'       CREATE_DATE       : 2009/10/23            hirano
'       PARAM             :第1引数(dao.DataBase) = SqlServerにDAO接続したDataBase
'                         :第2引数(Boolean)  = pblnNetHosho マスタ適用
'                         :第3引数(Integer)  = pintNetHoshoFlg 保証会社切替フラグ
'                         :第4引数(String)   = pstrNetHoshoCD 保証会社コード
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncGetNetHosyoInfo(dbSqlServerkase As DAO.Database, pblnNetHosho As Boolean, pintNetHoshoFlg As Integer, pstrNetHoshoCD As String) As Boolean
    Dim rsObject            As Recordset
    Dim strSQL              As String
    Dim strKey              As String
    Dim strData             As String
    On Error GoTo Exception
    ' 初期化
    pblnNetHosho = False
    ' SQL生成
    strSQL = ""
    strSQL = "SELECT * FROM PGPA_TABL WHERE PGPAT_PGP1B LIKE 'MNR210%' ORDER BY 1 "
    Set rsObject = dbSqlServerkase.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough, dbReadOnly)
    If rsObject.EOF Then
        fncGetNetHosyoInfo = False
    Else
        Do Until rsObject.EOF
            strKey = Nz(rsObject.Fields("PGPAT_PGP2B"))
            strData = Nz(rsObject.Fields("PGPAT_PARAN"))
            If strKey = "NET_HOSHO_SDATE" Then
                If strData <= Format$(DATE, "yyyymmddhh24miss") Then
                    'マスタ適用
                    pblnNetHosho = True
                Else
                    '適用前（パルマ固定）
                    pblnNetHosho = False
                End If
            ElseIf strKey = "NET_HOSHO_FLG" Then
                pintNetHoshoFlg = CInt(strData)
            ElseIf strKey = "NET_HOSHO_CD" Then
                pstrNetHoshoCD = strData
            End If
            rsObject.MoveNext
        Loop
        fncGetNetHosyoInfo = True
    End If
    rsObject.Close
    Set rsObject = Nothing
    Exit Function
Exception:
  Call Err.Raise(Err.Number, "fncGetNetHosyoInfo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

Private Sub cmb_Test_Click()
'    Dim oApp            As Object
'    Dim oBook           As Object
'    Dim varPrintSeets   As Variant
'    Dim stCONT          As MSZZ039.Type_CONT_MAST
'    Dim stBANK          As MSZZ039.Type_BANK_INF
'    Dim stMAIN          As MSZZ039.Type_MAIN_INF
'    Dim strHelperPath   As String
'    Dim isException     As Boolean
'
'    On Error GoTo Exception
'    isException = False
'
'    Call MSZZ039.HelperPrintPreview(strBumonCode, Me.txtRcptNo.VALUE)
'
'
'    GoTo Finally
'
'    Exit Sub
'
'Exception:
'    isException = True
'
'Finally:
'
'    If isException Then
'        Call Err.Raise(Err.Number, "cmb_Test_Click!" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subKinmusakiClear
'        機能             :顧客情報勤務先クリア
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subKinmusakiClear()

    txtUserKname = ""
    txtUserKkana = ""
    txtUserKpost = ""
    txtUserKadr1 = ""
    txtUserKadr2 = ""
    txtUserKadr3 = ""
    txtUserKtel = ""
    txtUserJob = ""
    txtUserKhourStart = ""
    txtUserKhourEnd = ""
    txtUserNenshu = ""

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserKinmusakiEnable
'        機能             :勤務先使用可不可設定処理
'        IN               :True=使用可に設定 False=使用不可に設定
'        OUT              :
'
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserKinmusakiEnable(blnEnable As Boolean)

    Call YellowControl(txtUserKname, Not blnEnable)
    Call YellowControl(txtUserKkana, Not blnEnable)
    Call YellowControl(txtUserKpost, Not blnEnable)
    
    cmdUserKpostConv.Enabled = blnEnable
    
    Call YellowControl(txtUserKadr1, Not blnEnable)
    Call YellowControl(txtUserKadr2, Not blnEnable)
    Call YellowControl(txtUserKadr3, Not blnEnable)
    Call YellowControl(txtUserKtel, Not blnEnable)
    Call YellowControl(txtUserJob, Not blnEnable)
    Call YellowControl(txtUserKhourStart, Not blnEnable)
    Call YellowControl(txtUserKhourEnd, Not blnEnable)
    Call YellowControl(txtUserNenshu, Not blnEnable)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :subUserKRenEnable
'        機能             :緊急連絡先使用可不可設定処理
'        IN               :True=使用可に設定 False=使用不可に設定
'        OUT              :
'
'        CREATE_DATE      : 2010/04/20            SHIBAZAKI
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subUserKRenEnable(blnEnable As Boolean)

    Call YellowControl(txtUserKrenName, Not blnEnable)
    Call YellowControl(txtUserKrenKana, Not blnEnable)
    Call YellowControl(txtUserKrenZokugara, Not blnEnable)
    Call YellowControl(txtUserKrenPost, Not blnEnable)
    
    cmdUserKrenPostConv.Enabled = blnEnable
    
    Call YellowControl(txtUserKrenAdr1, Not blnEnable)
    Call YellowControl(txtUserKrenAdr2, Not blnEnable)
    Call YellowControl(txtUserKrenAdr3, Not blnEnable)
    Call YellowControl(txtUserKrenTel, Not blnEnable)
    Call YellowControl(txtUserKrenFax, Not blnEnable)
    Call YellowControl(txtUserKrenKeitai, Not blnEnable)
    
    cmd緊急連絡先複写.Enabled = blnEnable   '2018/08/06 EGL add

End Sub

'==============================================================================*
'
'       MODULE_NAME     : フィールド存在確認
'       MODULE_ID       : fncContainsControls
'       CREATE_DATE     : 2010/04/28 SHIBAZAKI
'       PARAM           : strFieldName          フィールド名(I)
'                       : objRst                レコードセット
'       RETURN          : フィールドOBJECT
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncContainsField(ByVal strFieldName As String, ByVal objRst As Object) As Object
    On Error Resume Next
    Set fncContainsField = objRst.Fields(strFieldName)
    If Err.Number = 0 Then
        Exit Function
    End If
    Err.Clear
    Set fncContainsField = Nothing
End Function
'==============================================================================*
'
'        MODULE_NAME      :sub雑費割引再計算
'        機能             :雑費割引再計算
'        2011/06/18 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub雑費割引再計算(ByVal aMode As Integer, ByVal a元保証委託料 As Long)

    Dim st割引情報      As MSZZ068Type_DCRA_TRAN_INF
    Dim is割引有無      As Boolean
    Dim str部門コード   As String
    
    On Error GoTo Exception

    If aMode = P_MODE_本計算 Then
        '部門コードを取得
        str部門コード = Nz(DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1"))
        If str部門コード = "" Then
            'テーブル[CONT_MAST]の設定不正
            Call MSZZ024_M10("DLookup(""CONT_BUMOC"", ""dbo_CONT_MAST"", ""CONT_KEY=1"")", "テーブル[CONT_MAST]の設定不正。")
        End If
        '保証委託料割引を既にある割付から取得
        is割引有無 = MSZZ045_getHoshoWaribikiDCRA_TRAN(str部門コード, txtAcptNo, st割引情報)
    
    Else
        '保証委託料割引の再計算
        ' とりあえずネット予約のときのみの摘要とする
        If blnネット予約 = True Then
            is割引有無 = MSZZ068_getHoshoItakuWaribiki(cmbRcptKaihi, _
                                                    cmbContainerHosyoCompany, _
                                                    Replace(txtReceiptUkdate, "/", ""), _
                                                    a元保証委託料, _
                                                    st割引情報)
        End If
    End If
    
    ' 割引有ならば保証委託料割引を摘要する
    If is割引有無 = True Then
        txt保証金割引 = st割引情報.DCRAT_PRICE
    Else
        txt保証金割引 = 0
    End If
    
    Exit Sub
    
Exception:
    Call Err.Raise(Err.Number, "sub雑費割引再計算" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :sub雑費適用割引割付更新
'        機能             :雑費適用割引割付更新
'        2011/06/18 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub雑費適用割引割付更新(aRsObj As Recordset, aUserId As String)
    
    Dim st割引情報      As MSZZ068Type_DCRA_TRAN_INF
    Dim is割引有無      As Boolean
    Dim isException     As Boolean
    Dim lng元保証委託料 As Long
    Dim BillSecukg As Long
    On Error GoTo Exception
    
    ' とりあえずネット予約のときのみの摘要とする
    If blnネット予約 = False Then
        Exit Sub
    End If
        
        
    
    If blnネット予約 = True Then
        ' 2013/04/05 M.HONDA
        'BillSecukg = txtBillSecukg
        lng元保証委託料 = txtBillSecukg
    Else
        If txtEzappiCode = "68" Then
           lng元保証委託料 = txtBillSecukg - Nz(txtBillEzappi, 0)
        Else
           lng元保証委託料 = txtBillSecukg
        End If
    End If
    '' 2013/01/28 M.HONDA START
    'lng元保証委託料 = txtBillRentkg
    'lng元保証委託料 = txtBillSecukg
    '' 2013/01/28 M.HONDA END
    
    '保証委託料の再計算
    is割引有無 = MSZZ068_getHoshoItakuWaribiki(cmbRcptKaihi, _
                                            cmbContainerHosyoCompany, _
                                            Replace(txtReceiptUkdate, "/", ""), _
                                            lng元保証委託料, _
                                            st割引情報)
    ' 割引有ならば保証委託料割引の割付を保存する
        If is割引有無 = True Then
    
        aRsObj.AddNew
        aRsObj.Fields("DCRAT_ACPTNO") = txtAcptNo
        aRsObj.Fields("DCRAT_DCNT_NO") = st割引情報.DCRAT_DCNT_NO
        aRsObj.Fields("DCRAT_ENABLE") = st割引情報.DCRAT_ENABLE
        aRsObj.Fields("DCRAT_FROM") = st割引情報.DCRAT_FROM
        aRsObj.Fields("DCRAT_TO") = st割引情報.DCRAT_TO
        aRsObj.Fields("DCRAT_PRICE") = st割引情報.DCRAT_PRICE
        aRsObj.Fields("DCRAT_TEXT") = st割引情報.DCRAT_TEXT
        aRsObj.Fields("DCRAT_IYAKU_SEIKYU") = st割引情報.DCRAT_IYAKU_SEIKYU
        aRsObj.Fields("DCRAT_SEIKYU_KBN") = st割引情報.DCRAT_SEIKYU_KBN
        aRsObj.Fields("DCRAT_JITSU_KEIHI") = st割引情報.DCRAT_JITSU_KEIHI
        aRsObj.Fields("DCRAT_INSED") = Format$(DATE, "yyyymmdd")
        aRsObj.Fields("DCRAT_INSEJ") = Format$(time, "hhmmss")
        aRsObj.Fields("DCRAT_INSPB") = Me.NAME
            aRsObj.Fields("DCRAT_INSUB") = aUserId
        aRsObj.UPDATE
    
    End If
    Exit Sub
  
Exception:
    If Not aRsObj Is Nothing Then aRsObj.Close: Set aRsObj = Nothing
    Call Err.Raise(Err.Number, "sub雑費適用割引割付更新" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)

End Sub
'==============================================================================*
'
'        MODULE_NAME      :getネット予約クレカ保証会社
'        機能             :ネット予約クレカ保証会社取得
'        2011/06/18 new add
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getネット予約クレカ保証会社() As String

    Dim isException As Boolean  'エラー発生検知フラグ
    Dim objCon      As Object
    Dim objRst      As Object
    Dim sqlText     As String
    Dim str保証会社CD As String
    
    isException = False
    
    On Error GoTo Exception
    
    '初期化
    str保証会社CD = ""
    
    ' SQL生成
    sqlText = "SELECT * FROM PGPA_TABL WHERE PGPAT_PGP1B LIKE 'MSZZ0068_%' AND PGPAT_PGP2B = 'HO_WARI_2_TEKIYO_HCD' "
    
    ' SQL実行
    Set objCon = ADODB_Connection() 'KASE接続
    Set objRst = ADODB_Recordset(sqlText, objCon)
    If objRst.EOF Then
        'パラメータテーブルに登録なし
        str保証会社CD = ""
    Else
        str保証会社CD = Nz(objRst.Fields("PGPAT_PARAN"))
    End If
    
    GoTo Finally
    
Exception:
    isException = True
    
Finally:
    getネット予約クレカ保証会社 = str保証会社CD
    If Not objRst Is Nothing Then objRst.Close: Set objRst = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    If isException = True Then
        Call Err.Raise(Err.Number, "getネット予約クレカ保証会社" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function

'==============================================================================*
'
'        MODULE_NAME      :getFzappi1
'        機能             :事務手数料割引取得
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function getFzappi1(Optional txtUsePerido As Integer = 99) As Integer

    ' 割引マスタを参照し事務手数料が設定されていれば事務手数料割引の値をセットする。
    Dim strSQL2 As String
    Dim objAdoDbConnection   As Object
    Dim rsData               As Object       'ADODB.Recordset
    Dim blnFLG               As Boolean
    Dim strKaseSqlSvr        As String
    
'   strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATA_SOURCE_NAME'"))'DELETE 2019/03/07 Y.WADA
    strKaseSqlSvr = Nz(DLookup("SETUT_SETUN", "SETU_TABL", "SETUT_SETUB = 'ODBC_DATABASE_NAME'")) 'INSERT 2019/03/07 Y.WADA
    
    strSQL2 = ""
    strSQL2 = "SELECT * "
    strSQL2 = strSQL2 & " FROM YOUK_TRAN "
    strSQL2 = strSQL2 & " INNER JOIN CNTA_MAST ON "
    strSQL2 = strSQL2 & " YOUKT_YCODE = CNTA_CODE AND "
    strSQL2 = strSQL2 & " YOUKT_NO = CNTA_NO "
    strSQL2 = strSQL2 & " INNER JOIN RLDN_TRAN ON "
    strSQL2 = strSQL2 & " YOUKT_YCODE = RLDNT_YCODE AND "
    strSQL2 = strSQL2 & " RLDNT_ENABLE = '1'  AND "
    strSQL2 = strSQL2 & " RLDN_TRAN.RLDNT_FROM   <=  convert(varchar,YOUKT_UKDATE,112) AND "
    strSQL2 = strSQL2 & " RLDN_TRAN.RLDNT_TO     >=  convert(varchar,YOUKT_UKDATE,112) "
    strSQL2 = strSQL2 & " INNER JOIN " & strKaseSqlSvr & ".DBO.DCNT_MAST ON "
    strSQL2 = strSQL2 & " RLDNT_NO = DCNTM_NO AND "
    strSQL2 = strSQL2 & " RLDNT_BUMOC = DCNTM_BUMOC AND "
    strSQL2 = strSQL2 & " DCNTM_TYPE = 10 "
    strSQL2 = strSQL2 & " WHERE YOUKT_UKNO = '" & Me.txtYoukNo & "'"

    Set objAdoDbConnection = MSZZ025.ADODB_Connection(strBumonCode)
    Set rsData = MSZZ025.ADODB_Recordset(strSQL2, objAdoDbConnection)
        
    If rsData.EOF Then
        'データが存在しない
        blnFLG = False
    Else
        '割引の条件に合致するかチェック
        blnFLG = True
        
        '段
        If Not IsNull(rsData.Fields("DCNTM_FLOOR")) Then
            If rsData.Fields("DCNTM_FLOOR") <> rsData.Fields("CNTA_STEP") Then
                blnFLG = False
            End If
        End If
        
        'サイズ
        If Not IsNull(rsData.Fields("DCNTM_SIZE_FROM")) Then
            If rsData.Fields("DCNTM_SIZE_FROM") > rsData.Fields("CNTA_SIZE") Then
                blnFLG = False
            End If
        End If
        If Not IsNull(rsData.Fields("DCNTM_SIZE_TO")) Then
            If rsData.Fields("DCNTM_SIZE_TO") < rsData.Fields("CNTA_SIZE") Then
                blnFLG = False
            End If
        End If
        
        '使用区分
        If Not IsNull(rsData.Fields("DCNTM_USAGE")) Then
            If rsData.Fields("DCNTM_USAGE") <> rsData.Fields("CNTA_USAGE") Then
                blnFLG = False
            End If
        End If
        
        '利用期間
        If rsData.Fields("DCNTM_USE_PERIOD") <> 0 Then
            ' 割引再計算時は画面の利用期間
            If txtUsePerido <> 99 Then
                If rsData.Fields("DCNTM_USE_PERIOD") > txtUsePerido Then
                    blnFLG = False
                End If
            Else
                If rsData.Fields("DCNTM_USE_PERIOD") > rsData.Fields("YOUKT_USE_PERIOD") Then
                    blnFLG = False
                End If
            End If
        End If
        
        '種別チェック
        If rsData.Fields("DCNTM_GENKBN") <> 99 Then
            If rsData.Fields("DCNTM_GENKBN") <> rsData.Fields("YOUKT_GENKBN") Then
                blnFLG = False
            End If
        End If
    End If
    
    If blnFLG = False Then
        getFzappi1 = P_JIMUTESURYO
    Else
        getFzappi1 = rsData.Fields("DCNTM_VALUE")
    End If
    
    GoTo Finally
    
Exception:
    'エラーの場合にはP_JIMUTESURYOをセットする。
    getFzappi1 = P_JIMUTESURYO
    
Finally:
   
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not objAdoDbConnection Is Nothing Then objAdoDbConnection.Close: Set objAdoDbConnection = Nothing

End Function

'==============================================================================*
'
'        MODULE_NAME      :cmdWebPay_Click
'        機能             :WEB決済
'        CREAT            :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmdWebPay_Click()
On Error GoTo ErrorHandler
    'IEの起動
    Dim objIE As Object '変数を定義します。
    Set objIE = CreateObject("InternetExplorer.Application") 'オブジェクトを作成します。
    objIE.Visible = True      '可視、Trueで見えるようにします。
    
    '処理したいページを表示します。
    '' テスト環境へ
    '' objIE.Navigate "http://test3535.com/payment/CardEntry.php?&mode=confirm"
    
    '' 本番環境へ
    objIE.Navigate "https://kasem969.securesites.net/~kase3535/kase3535.com/payment/CardEntry.php?&mode=confirm"
    
    
    objIE.FullScreen = False  'メニューも全て無くした最大化になる
    objIE.Toolbar = True      'ToolBar On/Off
    objIE.MenuBar = True     'MenuBar On/Off
    
    
    'ページの表示完了を待ちます。
    While objIE.ReadyState <> 4 Or objIE.Busy = True '.ReadyState <> 4の間まわる。
        DoEvents
    Wend
    
    'google検索項目(name=q)にセットする。
    objIE.Document.getElementsByName("Amount")(0).VALUE = txtBillFirstAmount
    objIE.Document.getElementsByName("OrderID")(0).VALUE = Format(Now(), "YYYYMMDDhhmmss")
                                    
                                    
    objIE.Document.getElementsByName("UserCd")(0).VALUE = cmbUserCode
    objIE.Document.getElementsByName("UserName")(0).VALUE = txtUserName
    objIE.Document.getElementsByName("Yard")(0).VALUE = txtContainerYname
    objIE.Document.getElementsByName("No")(0).VALUE = txtContainerCno
    objIE.Document.getElementsByName("Contents")(0).VALUE = txtContainerUsage
    
    objIE.Document.getElementsByName("ClientField3")(0).VALUE = DLookup("CONT_BUMOC", "dbo_CONT_MAST", "CONT_KEY=1") & _
                                    cmbUserCode & txtAcptNo
                                    
    Exit Sub
    
ErrorHandler:
    Call MSZZ024_M00("cmdWebPay_Click", True)
    Err.Clear

End Sub

'==============================================================================*
'
'        MODULE_NAME      : subCompSecukgKase
'        機能             : 保証委託料加瀬負担分再計算
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subCompSecukgKase()
    On Error GoTo ErrorHandler

    If chkSecukgKase.VALUE Then
        txtBillSecukgKase.VALUE = txtBillMonthCharge.VALUE + Nz(txt月額分割引.VALUE, 0)
        txtBillSecukg.VALUE = 0
    Else
        txtBillSecukgKase.VALUE = 0
    End If
Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subCompSecukgKase" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : subDispSecukgMessage
'        機能             : 月額使用料差異表示
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispSecukgMessage()
    Dim bVisible            As Boolean
    On Error GoTo ErrorHandler
    
    If txtReceiptYukbnCode = P_YUKBN_受付 Then
        bVisible = (Val(Nz(txtContainerPrice)) <> Val(Nz(txtBillRentkg)))
    Else
        bVisible = False
    End If
    txtSecukgMessage.Visible = bVisible
    txtSecukgMessage.VALUE = "月額使用料(" & Format(Nz(txtContainerPrice.VALUE), "#,##0") & ")と差異があります。"
    lineLeft.Visible = bVisible
    lineCenter.Visible = bVisible
    lineRight.Visible = bVisible
Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subDispSecukgMessage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'        MODULE_NAME      : subDispEzappiMessage
'        機能             : 月額雑費差異表示
'        CREAT            : 2013/04/10          K.ISHIZAKA
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subDispEzappiMessage()
    Dim bVisible            As Boolean
    On Error GoTo ErrorHandler
    
    '▼2018/08/06 EGL INS<S>
    '月額雑費差異表示処理は既存顧客新保証会社プラン移行データの場合は実施しない
    If P_HOSPLN_移行対象 = strNewHosplnIkou Then
        Exit Sub
    End If
    '▲2018/08/06 EGL INS<E>
    
    If txtReceiptYukbnCode = P_YUKBN_受付 Then
        bVisible = (Val(Nz(txtContainerEzapii)) <> Val(Nz(txtBillEzappi)))
    Else
        bVisible = False
    End If
    txtEzappiMessage.Visible = bVisible
    txtEzappiMessage.VALUE = Nz(txtContainerEzapiiName.VALUE, "月額雑費") & "(" & Format(Nz(txtContainerEzapii.VALUE), "#,##0") & ")と差異があります。"
    lineLeft2.Visible = bVisible
    lineCenter2.Visible = bVisible
    lineRight2.Visible = bVisible
Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number, "subDispEzappiMessage" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : Ｔポイント還元率取得
'       MODULE_ID       : GetTPointRitu
'       CREATE_DATE     : 2013/06/21            K.ISHIZAKA
'       PARAM           : dUKDATE               受付日時(I)
'       RETURN          : ＳＱＬ文(String)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function GetTPointRitu(ByVal dUKDATE As Date) As Double
    Dim strSQL              As String
    Dim objCon              As Object
    On Error GoTo ErrorHandler
    
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " TOP 1 "
    strSQL = strSQL & " RITU "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " RITU_KEY,"
    strSQL = strSQL & " RITU,"
    strSQL = strSQL & " DATE_FROM,"
    strSQL = strSQL & " DATE_TO,"
    strSQL = strSQL & " TIME_FROM,"
    strSQL = strSQL & " ISNULL(PGPAT_PARAN,'235959') AS TIME_TO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " RITU_KEY,"
    strSQL = strSQL & " RITU,"
    strSQL = strSQL & " DATE_FROM,"
    strSQL = strSQL & " DATE_TO,"
    strSQL = strSQL & " ISNULL(REPLACE(PGPAT_PGP2B, RITU_KEY + '.TIME_FROM',''),'') AS TIME_KEY,"
    strSQL = strSQL & " ISNULL(PGPAT_PARAN,'000000') AS TIME_FROM "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " RITU_KEY,"
    strSQL = strSQL & " RITU,"
    strSQL = strSQL & " DATE_FROM,"
    strSQL = strSQL & " ISNULL(PGPAT_PARAN,'99991231') AS DATE_TO "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " RITU_KEY,"
    strSQL = strSQL & " RITU,"
    strSQL = strSQL & " ISNULL(REPLACE(PGPAT_PGP2B, RITU_KEY + '.DATE_FROM',''),'') AS DATE_KEY,"
    strSQL = strSQL & " ISNULL(PGPAT_PARAN,'00000000') AS DATE_FROM "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & "("
    strSQL = strSQL & "SELECT"
    strSQL = strSQL & " REPLACE(PGPAT_PGP2B,'RITU.','') AS RITU_KEY,"
    strSQL = strSQL & " PGPAT_PARAN AS RITU "
    strSQL = strSQL & "FROM"
    strSQL = strSQL & " PGPA_TABL "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " PGPAT_PGP1B like 'FVS400_TPOINT.%' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PGPAT_PGP2B like 'RITU." & IIf(blnネット予約, "%", "NORMAL") & "' "
    strSQL = strSQL & ") ritu "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " PGPA_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PGPAT_PGP1B like 'FVS400_TPOINT.%' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PGPAT_PGP2B like RITU_KEY + '.DATE_FROM%' "
    strSQL = strSQL & ") datefrom "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " PGPA_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PGPAT_PGP1B like 'FVS400_TPOINT.%' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PGPAT_PGP2B like RITU_KEY + '.DATE_TO' + DATE_KEY "
    strSQL = strSQL & ") dateto "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " PGPA_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PGPAT_PGP1B like 'FVS400_TPOINT.%' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PGPAT_PGP2B like RITU_KEY + '.TIME_FROM%' "
    strSQL = strSQL & ") timeform "
    strSQL = strSQL & "LEFT JOIN"
    strSQL = strSQL & " PGPA_TABL "
    strSQL = strSQL & "ON"
    strSQL = strSQL & " PGPAT_PGP1B like 'FVS400_TPOINT.%' "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " PGPAT_PGP2B like RITU_KEY + '.TIME_TO' + TIME_KEY "
    strSQL = strSQL & ") timeto "
    strSQL = strSQL & "WHERE"
    strSQL = strSQL & " '" & Format(dUKDATE, "yyyymmdd") & "' BETWEEN DATE_FROM AND DATE_TO "
    strSQL = strSQL & "AND"
    strSQL = strSQL & " '" & Format(dUKDATE, "hhnnss") & "' BETWEEN TIME_FROM AND TIME_TO "
    strSQL = strSQL & "ORDER BY"
    strSQL = strSQL & " RITU desc "
    
    Set objCon = ADODB_Connection()
    On Error GoTo ErrorHandler1
    GetTPointRitu = CDbl(ADODB_ExecGetVariant(strSQL, objCon, "0"))
    On Error GoTo ErrorHandler
    objCon.Close
Exit Function

ErrorHandler1:
    objCon.Close
ErrorHandler:          '↓自分の関数名
    Call Err.Raise(Err.Number, "GetTPointRitu" & vbRightAllow & Err.Source, Err.Description & vbCrLf & "SQL:" & vbCrLf & strSQL, Err.HelpFile, Err.HelpContext)
End Function
'==============================================================================*
'
'        MODULE_NAME      :cmb_性別_AfterUpdate
'        機能             :cmb_性別_AfterUpdate イベント
'        2014/10/14 new add
'==============================================================================*
Private Sub cmb_性別_AfterUpdate()
    On Error GoTo ErrorHandler
    txt_性別.VALUE = fncAfterUpdateCommon(cmb_性別, 1, P_CONTROL_数値)
Exit Sub

ErrorHandler:
    Call MSZZ024_M00("cmb_性別_AfterUpdate", True)
End Sub

'==============================================================================*
'
'        MODULE_NAME      :cmb_性別_BeforeUpdate
'        機能             :cmb_性別_BeforeUpdate イベント
'        2014/10/14 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_性別_BeforeUpdate(Cancel As Integer)
    Call fncBeforeUpdateCommon(cmb_性別, 1, P_CONTROL_数値, Cancel)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmb_性別_KeyPress
'        機能             :cmb_性別_KeyPress イベント
'        2014/10/14 new add
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_性別_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(cmb_性別, KeyAscii, 1)
End Sub

'==============================================================================*
'
'       MODULE_NAME     : 身分証明書の表示
'       MODULE_ID       : Show_RCPT_SHOUMEI
'       CREATE_DATE     : 2015/07/16            K.ISHIZAKA
'       PARAM           : strYoukNo             予約番号(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub Show_RCPT_SHOUMEI(ByVal strYoukNo As String)
    Dim strPath             As String
    Dim strFile             As String
    Dim i                   As Long
    
    'On Error GoTo ErrorHandler
    On Error Resume Next
    
    
    i = 0
'    strPath = GetNetReceptPath(NetReceptType.LocalPath, NetReceptPath.P_USERIMG)
    strPath = Nz(DLookup("INTIF_RECDB", "INTI_FILE", "INTIF_PROGB='" & Me.NAME & "' AND INTIF_RECFB='USERIMG_PATH'"))
    If Nz(strPath) = "" Then
        Call MSZZ024_M10("DLookup(""INTIF_RECDB"", ""INTI_FILE"", ""INTIF_PROGB='" & Me.NAME & "' AND INTIF_RECFB='USERIMG_PATH'"")", _
                         "ヘルパーの格納パスを取得できませんでした。")
    End If
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
'    strFile = Dir(strPath & strYoukNo & "*.JPG")                               'DELETE 2015/07/26 K.ISHIZAKA
    strFile = Dir(strPath & strYoukNo & "*.*")                                  'INSERT 2015/07/26 K.ISHIZAKA
    While strFile <> ""
        i = i + 1
'        With Me.Controls("img_RCPT_SHOUMEI_" & Format(i))                       'DELETE START 2015/07/26 K.ISHIZAKA
'            .Picture = strPath & strFile
'            .HyperlinkAddress = strPath & strFile
'        End With                                                                'DELETE END   2015/07/26 K.ISHIZAKA
        Call SetHyperlinkImage(Me.Controls("img_RCPT_SHOUMEI_" & Format(i)), strPath, strFile) 'INSERT 2015/07/26 K.ISHIZAKA
        If i < 5 Then
            strFile = Dir()
        Else
            Exit Sub
        End If
    Wend
Exit Sub

'ErrorHandler:
'    Call Err.Raise(Err.Number, "Show_RCPT_SHOUMEI" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

' ***** Start 2017/07/31 add by ysuzuki

'==============================================================================*
'   MODULE_NAME : コンテナ契約ファイルと請求ファイルの改修

'   MODULE_ID   : MODIFY_CARG_REQU_FILE
'   Parameter   : 第1引数(dao.DataBase) : dbSqlServer = SqlServerにDAO接続したDataBase
'                 第2引数(String)       : strUCODE = 顧客コード
'                 第3引数(String)       : strACPTNO = 契約番号
'                 第5引数(String)       : strUserID = ユーザーID
'   Return      : Long = ( 異常終了=-1、正常終了=0(削除も登録もなし)、正常終了(削除のみ)=1、正常終了(DELETE-INSERT)=2 )
'
'   CREATE_DATE : 2017/07/28
'   Note        :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function MODIFY_CARG_REQU_FILE(dbSQLServer As DAO.Database _
                                      , ByVal lngUcode As Long _
                                      , ByVal strACPTNO As String _
                                      , ByVal strUserID As String) As Long
    ' 変数宣言
    Dim strHassei_Kbn       As String           ' 発生区分
    Dim strShoki_Seikyu     As String           ' 初期費用請求方法
    Dim lngReqNo            As Long             ' 請求番号
    Dim strMsg              As String           ' (OUT用)エラーメッセージ
    Dim strucCarg           As Type_CARG_FILE   ' コンテナ契約ファイル構造体
    Dim strSQL              As String
    Dim objRs               As Recordset
    '戻り値
    Dim lngRtnRT            As Long
    Dim lngRTN              As Long
    Dim blnRTN              As Boolean
    
    MODIFY_CARG_REQU_FILE = -1
    
    ' 画面の値を取得
    strHassei_Kbn = IIf(Not blnネット予約, 0, 1)            '発生区分
    strShoki_Seikyu = Nz(Me.cmbBillSeikyuCd.VALUE, "")      '初期費用請求方法

    ' コンテナ契約ファイル、請求ファイルへの登録対象かどうか取得
    lngRtnRT = CmKomsMod.CheckRegiTarget(dbSQLServer, strHassei_Kbn, strShoki_Seikyu, lngUcode, strACPTNO)
    If 0 > lngRtnRT Then GoTo Exception
    
    '
    If 0 = lngStRtn And 0 = lngRtnRT Then
        '戻り値を設定して、処理を終了
        MODIFY_CARG_REQU_FILE = 0
            Exit Function
    End If
    
    ' コンテナ契約ファイルと請求ファイルへの登録対象、または登録対象から外れた場合
    If 1 = lngRtnRT Or (1 = lngStRtn And 0 = lngRtnRT) Then
        ' 請求ファイルの削除
        blnRTN = CmKomsMod.DeleteRequFile(dbSQLServer, strACPTNO)
        If Not blnRTN Then GoTo Exception
        
    End If
    
    ' ****************************************
    ' コンテナ契約ファイル登録処理
    ' ****************************************
    
    ' コンテナ契約ファイルデータ抽出
    strSQL = "SELECT * FROM CARG_FILE WHERE CARG_ACPTNO = '" & strACPTNO & "'"
    Set objRs = dbSQLServer.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' 該当データ在り、または該当データはないが登録対象
    If Not objRs.EOF Or (objRs.EOF And 1 = lngRtnRT) Then
        ' コンテナ契約ファイル作成構造体設定
        Call subSetCargStruc(strucCarg, strACPTNO, "3")
        
        ' コンテナ契約ファイルへ登録
        blnRTN = InsertCargFile(dbSQLServer, strucCarg, Me.NAME, strUserID, strMsg)
        If Not blnRTN Then
            Call MSZZ024_M10("InsertCargFile", strMsg)
            Exit Function
        End If
        
    End If

    ' ****************************************
    ' 請求ファイルの登録処理
    ' ****************************************
    ' 登録対象の場合
    If 1 = lngRtnRT Then
        ' 請求番号を採番
        lngReqNo = fncNewRequNo(dbSQLServer)
        If 0 > lngReqNo Then
            Call MSZZ024_M10("fncNewRequNo", "請求番号を採番できませんでした。")
            Exit Function
            
        End If
        
        ' 請求ファイルの登録
            If Not CmKomsMod.InsertRequFile2(dbSQLServer _
                                        , lngReqNo _
                                        , txtBillStDate _
                                        , Format(txtBillStDate, "yyyy/mm/01") _
                                        , cmbUserCode _
                                        , txtContainerYcode _
                                        , txtContainerCno _
                                        , Val(txtBillSyoKei + txt初回分割引合計), 0, Val(Nz(txtBillSecukg, 0)) + Val(Nz(txt保証金割引, 0)) _
                                        , strACPTNO _
                                        , strMsg) Then
            Call MSZZ024_M10("InsertRequFile2", strMsg)
            Exit Function
            
        End If
        
    End If
    
    MODIFY_CARG_REQU_FILE = 2
    Exit Function
    
Exception:
    If -1 = MODIFY_CARG_REQU_FILE Then
        'エラー処理
        Call Err.Raise(Err.Number, "MODIFY_CARG_REQU_FILE" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'   MODULE_NAME : 請求番号発番
'   MODULE_ID   : fncNewRequNo
'   Parameter   : 第1引数(dao.DataBase)  : dbSqlServer = SqlServerにDAO接続したDataBase
'   Return      : Long = ( 異常終了=-1、正常終了=(新)請求番号 )
'
'   CREATE_DATE : 2017/08/04
'   Note        :
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncNewRequNo(objDb As Database) As Long

    Dim rsData      As Recordset
    Dim lngReqNo    As Long
    
    
    On Error GoTo Exception
    
    fncNewRequNo = -1
    
    ' コントロールマスタより取得
    Set rsData = objDb.OpenRecordset(fncMakeContMastSql(), dbOpenDynaset)

    If rsData.EOF Then
        Call MSZZ024_M10("OpenRecordset", "コントロールマスタを参照できませんでした。")
        If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
        Exit Function
    End If

    With rsData
        '請求NO
        lngReqNo = Nz(.Fields("CONT_REQNO"), 0) + 1
        
        .Edit
        
        .Fields("CONT_REQNO") = lngReqNo                        ' 請求NO
        .Fields("CONT_SKDATE") = Format(DATE, "yyyy/mm/dd")     ' 更新日
        .Fields("CONT_UPDATE") = Now                            ' 更新日
        
        .UPDATE
        
    End With

    fncNewRequNo = lngReqNo

Exception:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If -1 = fncNewRequNo Then
        Call Err.Raise(Err.Number, "fncNewRequNo" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        
    End If

End Function
' ***** end 2017/07/31 add by ysuzuki

' ***** start 2018/01/20 add by sato
'==============================================================================*
'
'        MODULE_NAME      :fncUpdHoseTabl
'        機能             :保証会社番号更新
'        IN               :保証会社コード
'        OUT              :保証会社番号
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdHoseTabl(intSHIRC As String)
On Error GoTo ErrorHandler
    
    Dim strSQL      As String
    Dim rsData      As Recordset
    Dim blnError    As Boolean
    Dim iNo         As Integer
    Dim strConnectKASE As String
    
    Dim objDb               As Database
    Dim objWs               As Workspace

    Dim objCon          As Object
    Dim objRst          As Object

    blnError = False
    
    fncUpdHoseTabl = False
    
    strConnectKASE = MSZZ007_M00("")
    
    'SQLサーバー接続
            Set objDb = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnectKASE)

    'トランザクション開始
    Set objWs = DBEngine.Workspaces(0)
    objWs.BeginTrans

    strSQL = "SELECT * FROM HOSE_TABL WHERE HOSET_SHIRC = '" & intSHIRC & "'"
    Set rsData = objDb.OpenRecordset(strSQL, dbOpenDynaset)
    
    'Set objCon = ADODB_Connection("")
    'Call ADODB_Execute(strSql, objCon)
    
    With rsData

        If rsData.EOF Then
            .AddNew
            .Fields("HOSET_SHIRC") = intSHIRC
            iNo = 1
        Else
            .Edit
            iNo = .Fields("HOSET_SEQ") + 1
        End If

        .Fields("HOSET_SEQ") = iNo

MsgBox "stop " & iNo

        .UPDATE

    End With

    fncUpdHoseTabl = iNo
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True
    
    'ロールバック
    If Not objWs Is Nothing Then
        objWs.Rollback
    End If

ExitRtn:
    objWs.CommitTrans
    Set objWs = Nothing
    
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, "fncUpdHoseTabl" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function

'==============================================================================*
'
'        MODULE_NAME      :fncUpdHoseTabl2
'        機能             :保証会社番号更新
'        IN               :保証会社コード
'        OUT              :保証会社番号
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncUpdHoseTabl2(intSHIRC As String)
On Error GoTo ErrorHandler
    
    Dim strSQL      As String
    Dim rsData      As Recordset
    Dim blnError    As Boolean
    Dim iNo         As Integer
    Dim strConnectKASE As String
    
    Dim objDb               As Database
    Dim objWs               As Workspace
    
    Dim ret As Variant

    Dim objCon          As Object
    Dim objRst          As Object

    blnError = False
    
    Set objCon = ADODB_Connection("")

    strSQL = "SET LOCK_TIMEOUT 1000 SELECT HOSET_SEQ FROM HOSE_TABL WITH(XLOCK) WHERE HOSET_SHIRC = '" & intSHIRC & "'"
    iNo = ADODB_ExecGetLong(strSQL, objCon)
    
    If Nz(iNo, -1) = -1 Then
        iNo = 1
        strSQL = "insert into HOSE_TABL values('" & intSHIRC & "'," & iNo & ")"
    Else
        iNo = iNo + 1
        strSQL = "update HOSE_TABL  set HOSET_SEQ = " & iNo & " WHERE HOSET_SHIRC = '" & intSHIRC & "'"
    End If
    
MsgBox "stop " & iNo
    
    ret = ADODB_Execute(strSQL, objCon)
    
    'Set objCon = ADODB_Connection("")
    'Call ADODB_Execute(strSql, objCon)
    




    fncUpdHoseTabl2 = iNo
    
    GoTo ExitRtn

ErrorHandler:
    blnError = True
    

ExitRtn:
    
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not objDb Is Nothing Then objDb.Close: Set objDb = Nothing
    
    If blnError Then
        Call Err.Raise(Err.Number, "fncUpdHoseTabl" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Function


'==============================================================================*
'
'       MODULE_NAME     : 保証会社番号発番
'       MODULE_ID       : fncGETHOSESEQ
'       CREATE_DATE     : 2018/01/22            K.SATO
'       PARAM           : strSHIRC              保証会社コード(I)
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Function fncGETHOSESEQ(ByVal strSHIRC As String)
    Dim objCmd              As Object   'ADODB.Command
    Dim objRst              As Object
    Dim blnError    As Boolean
    
    Const adCmdStoredProc = 4
    Const adParamInput = 1
    Const adParamOutput = 2
    
    On Error GoTo ErrorHandler
    
    blnError = False
    fncGETHOSESEQ = ""
    
    Set objCmd = CreateObject("ADODB.Command")
    With objCmd
        .CommandType = adCmdStoredProc
        .CommandText = "GETHOSESEQ"
        .Parameters.Append .CreateParameter("SHIRC", adVarChar, adParamInput, 6)
        .Parameters.Append .CreateParameter("SEQ", adVarChar, adParamOutput, 20)
        .Parameters.Append .CreateParameter("RET", adVarChar, adParamOutput, 256)
        .Parameters.Append .CreateParameter("MSG", adVarChar, adParamOutput, 256)
    End With
    
    Set objCmd.ActiveConnection = ADODB_Connection()
    
    On Error GoTo ErrorHandler1
    
    objCmd.Parameters("SHIRC").VALUE = strSHIRC
    objCmd.Execute
        
    objCmd.ActiveConnection.Close
    
    fncGETHOSESEQ = objCmd.Parameters("SEQ").VALUE

    GoTo ExitRtn

ErrorHandler:

    blnError = True

ErrorHandler1:

    blnError = True
    objCmd.ActiveConnection.Close
    
ExitRtn:
    
    If blnError Then
        Call Err.Raise(Err.Number, "GETHOSESEQ" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
' ***** 2018/08/06 EGL INS<S>
'==============================================================================*
'
'        MODULE_NAME      :cmd緊急連絡先複写_Click
'        機能             :cmd緊急連絡先複写_Click イベント
'        IN               :
'        OUT              :
'        CREATE_DATE      : 2018/08/06            TAKENOUCHI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmd緊急連絡先複写_Click()

    On Error GoTo Exception
    
    Call sub緊急連絡先複写(cmbUserCode.VALUE, txtYoukNo.VALUE)
    Exit Sub

Exception:

    Call MSZZ024_M00("cmd緊急連絡先複写_Click", True)

End Sub

'==============================================================================*
'
'        MODULE_NAME      :sub緊急連絡先複写
'        機能             :顧客コードを取得キーとして受付トランから起算日が
'                          直近（自予約以外で）の緊急連絡先を取得する
'        IN               :顧客コード,予約番号
'        OUT              :
'        CREATE_DATE      : 2018/08/06            TAKENOUCHI
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub sub緊急連絡先複写(ByVal strCustCd As String, ByVal str予約番号 As String)
 
    On Error GoTo ErrorHandler
    
    Dim dbSQLServer     As Database
    Dim rsData          As Recordset
    Dim blnError        As Boolean
    Dim strSQL          As String
    
    blnError = False
        
    strSQL = strSQL & " select "
    strSQL = strSQL & "   RT1.RCPT_KREN_NAME      AS RCPT_KREN_NAME "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_KANA      AS RCPT_KREN_KANA "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_YUBINO    AS RCPT_KREN_YUBINO "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_ADR_1     AS RCPT_KREN_ADR_1 "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_ADR_2     AS RCPT_KREN_ADR_2 "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_ADR_3     AS RCPT_KREN_ADR_3 "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_TEL       AS RCPT_KREN_TEL "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_KEITAI    AS RCPT_KREN_KEITAI "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_FAX       AS RCPT_KREN_FAX "
    strSQL = strSQL & "  ,RT1.RCPT_KREN_ZOKUGARA  AS RCPT_KREN_ZOKUGARA "
    strSQL = strSQL & " from RCPT_TRAN AS RT1 "
    strSQL = strSQL & " where RT1.RCPT_UCODE = " & strCustCd
    strSQL = strSQL & "   and RT1.RCPT_KISAN_DATE = "
    strSQL = strSQL & "      ( select max(RCPT_KISAN_DATE) from RCPT_TRAN AS RT2 where RT2.RCPT_KREN_NAME is not null and RT2.RCPT_UCODE = " & strCustCd & " and RT2.RCPT_NO != '" & str予約番号 & "' ) "
    
    'SQLサーバー接続
    Set dbSQLServer = Workspaces(0).OpenDatabase(strSqlserver, dbDriverNoPrompt, False, strConnect)
    Set rsData = dbSQLServer.OpenRecordset(strSQL, dbOpenSnapshot, dbSQLPassThrough)
    
    If Not rsData.EOF Then
        'データ表示
        With rsData
            txtUserKrenName = .Fields("RCPT_KREN_NAME")
            txtUserKrenKana = .Fields("RCPT_KREN_KANA")
            txtUserKrenPost = .Fields("RCPT_KREN_YUBINO")
            txtUserKrenAdr1 = .Fields("RCPT_KREN_ADR_1")
            txtUserKrenAdr2 = .Fields("RCPT_KREN_ADR_2")
            txtUserKrenAdr3 = .Fields("RCPT_KREN_ADR_3")
            txtUserKrenTel = .Fields("RCPT_KREN_TEL")
            txtUserKrenKeitai = .Fields("RCPT_KREN_KEITAI")
            txtUserKrenFax = .Fields("RCPT_KREN_FAX")
            txtUserKrenZokugara = .Fields("RCPT_KREN_ZOKUGARA")
        End With
    End If
    
    GoTo ExitRtn
    
ErrorHandler:
    blnError = True

ExitRtn:
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not dbSQLServer Is Nothing Then dbSQLServer.Close: Set dbSQLServer = Nothing

    If blnError Then
        Call Err.Raise(Err.Number, "fncUserRefresh" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Sub
' ***** 2018/08/06 EGL INS<E>

'2020/11/16 [EGL] INS Start
'==============================================================================*
'
'        MODULE_NAME      :cmbShoumeiAdrsai_AfterUpdate
'        機能             :cmbShoumeiAdrsai_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbShoumeiAdrsai_AfterUpdate()
    txtShoumeiAdrsai = fncAfterUpdateCommon(cmbShoumeiAdrsai, P_DIVISION_LENGTH, P_CONTROL_数値)
End Sub
'==============================================================================*
'
'        MODULE_NAME      :cmbShoumeiAdrsai_BeforeUpdate
'        機能             :cmbShoumeiAdrsai_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbShoumeiAdrsai_BeforeUpdate(Cancel As Integer)
    Call fncBeforeUpdateCommon(cmbShoumeiAdrsai, P_DIVISION_LENGTH, P_CONTROL_数値, Cancel)
End Sub
'==============================================================================*
'
'        機能             :cmbShoumeiAdrsai_KeyPress
'        MODULE_NAME      :cmbShoumeiAdrsai_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmbShoumeiAdrsai_KeyPress(KeyAscii As Integer)
    Call CmKeyPress_For_Num1(cmbShoumeiAdrsai, KeyAscii, P_DIVISION_LENGTH)
End Sub
'2020/11/16 [EGL] INS End

'2020/12/01 [EGL] INS Start
'==============================================================================*
'
'        MODULE_NAME      : cmb_クーポンCD_AfterUpdate
'        機能             : cmb_クーポンCD_AfterUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_クーポンCD_AfterUpdate()
    
    txt_紹介会社名 = fncAfterUpdateCommon(cmb_クーポンCD, 8, P_CONTROL_半角)

End Sub

'==============================================================================*
'
'        MODULE_NAME      : cmb_クーポンCD_BeforeUpdate
'        機能             : cmb_クーポンCD_BeforeUpdate イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_クーポンCD_BeforeUpdate(Cancel As Integer)
    
    If fncChkExistCmbInput(cmb_クーポンCD) < 0 Then
        MsgBox "入力された紹介コードは無効なコードです。", vbInformation, P_KAKUNIN
        Cancel = True
    End If

End Sub

'==============================================================================*
'
'        MODULE_NAME      : cmb_クーポンCD_KeyPress
'        機能             : cmb_クーポンCD_KeyPress イベント
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub cmb_クーポンCD_KeyPress(KeyAscii As Integer)
    
    Call CmKeyPress_For_Char(cmb_クーポンCD, KeyAscii, 8)

End Sub

'==============================================================================*
'
'        MODULE_NAME      : subクーポンCDコンボ値セット
'        機能             : クーポンCDコンボボックスにリストデータを設定
'        IN               :
'        OUT              :
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub subクーポンCDコンボ値セット()

On Error GoTo Exception

    Dim objCon          As Object
    Dim objRsData       As Object
    Dim strSQL          As String
    Dim isException     As Boolean
    
    isException = False
    
    strSQL = "select PGPAT_PGP2B,PGPAT_BIKON from PGPA_TABL where PGPAT_PGP1B = 'FVS400_COUPON' order by PGPAT_PGP2B"
    Set objCon = ADODB_Connection()     'KASE接続
    Set objRsData = ADODB_Recordset(strSQL, objCon)

    cmb_クーポンCD.RowSource = ""
    While Not objRsData.EOF
        cmb_クーポンCD.AddItem objRsData.Fields("PGPAT_PGP2B") & ";" & objRsData.Fields("PGPAT_BIKON")
        objRsData.MoveNext
    Wend

    GoTo Finally
    
Exception:
    isException = True

Finally:
    
    If Not objRsData Is Nothing Then objRsData.Close: Set objRsData = Nothing
    If Not objCon Is Nothing Then objCon.Close: Set objCon = Nothing
    
    If isException Then
        Call Err.Raise(Err.Number, "subクーポンCDコンボ値セット" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
    
End Sub

'==============================================================================*
'
'        MODULE_NAME      : fncChkExistCmbInput
'        機能             : チェック対象データがコンボボックスリスト内に存在
'                           するかチェックする
'
'        IN               : コンボボックスオブジェクト
'        OUT              : 0：正常、-1：エラー
'        CREATE_DATE      : 2020/12/01    [EGL]
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Function fncChkExistCmbInput(cmbCheck As ComboBox) As Integer
 
On Error GoTo Exception
    
    Dim isException   As Boolean
    Dim intCnt        As Integer
    Dim isFound       As Boolean
    Dim intRet        As Integer
    Dim strCode       As String
    
    intRet = 0
    isException = False
    
    strCode = Nz(cmbCheck, "")
    
    '空白なら入力チェックしない
    If strCode = "" Then
        GoTo Finally
    End If
    
    '入力値がリストに存在しているかチェック
    If cmbCheck.ListCount <> 0 Then
        isFound = False
        'コンボボックス内に該当コードが存在するかチェックする
        For intCnt = 0 To cmbCheck.ListCount - 1
            'コード値比較
            If cmbCheck.Column(0, intCnt) = strCode Then
                isFound = True
                Exit For
            End If
        Next
        
        If Not isFound Then
            intRet = -1
        End If
    End If
    
    GoTo Finally
    
Exception:
    isException = True

Finally:
    fncChkExistCmbInput = intRet

    If isException Then
        Call Err.Raise(Err.Number, "fncChkExistCmbInput" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If

End Function
'2020/12/01 [EGL] INS End

'==============================================================================*
'
'        MODULE_NAME      : fncInsertKouzMast
'        機能             : 仮想口座割当
'        CREATE_DATE      : 2023/10/26
'
'==============================================================================*
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Sub fncInsertKouzMast()
 
On Error GoTo ErrorHandler

    Dim strSQL  As String
    Dim rsData  As Recordset

    '仮想口座追加
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & "   FROM KOUZ_MAST"
    strSQL = strSQL & "  WHERE KOUZM_BUMOC = '" & strBumonCode & "'"
    strSQL = strSQL & "    AND NZ(KOUZM_KOKYC, '') = '' "
    'strSQL = strSQL & "    AND NOT EXISTS(SELECT * FROM KOUZ_MAST WHERE KOUZM_BUMOC = '" & strBumonCode & "' AND KOUZM_KOKYC = '" & Format(cmbUserCode, "000000") & "')"    'INSERT 2023/11/09 N.IMAI
    strSQL = strSQL & "    AND NOT EXISTS(SELECT * FROM KOUZ_MAST WHERE KOUZM_BUMOC = '" & strBumonCode & "' AND KOUZM_KOKYC = '" & Format(cmbUserCode, "000000") & "' AND KOUZM_CAMPC = '" & Trim(strCampCode) & "')"  'INSERT 2024/02/19 N.IMAI
    Set rsData = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If rsData.EOF = False Then
        strSQL = ""
        strSQL = strSQL & "UPDATE KOUZ_MAST_J"
        strSQL = strSQL & "   SET KOUZM_KOKYC = '" & Format(cmbUserCode, "000000") & "'"
        strSQL = strSQL & "      ,KOUZM_UPDAD = '" & Format(Now(), "yyyymmdd") & "'"
        strSQL = strSQL & "      ,KOUZM_UPDAJ = '" & Format(Now(), "hhnnss") & "'"
        strSQL = strSQL & "      ,KOUZM_UPDPB = '" & Me.NAME & "'"
        strSQL = strSQL & "      ,KOUZM_UPDUB = '" & LsGetUserName() & "'"
        strSQL = strSQL & " WHERE KOUZM_KOUZB = '" & rsData.Fields("KOUZM_KOUZB") & "'"
        strSQL = strSQL & "   AND KOUZM_KINYC = '" & rsData.Fields("KOUZM_KINYC") & "'"
        strSQL = strSQL & "   AND KOUZM_SHITC = '" & rsData.Fields("KOUZM_SHITC") & "'"
        Call ADODB_Execute(strSQL)
    End If

ErrorHandler:
    If Err <> 0 Then
        Call Err.Raise(Err.Number, "fncInsertKouzMast" & vbRightAllow & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub


'==============================================================================*
'        MODULE_NAME      : PerformFaceAuthTransferFlow
'        機能             : 顔認証システムの部屋移動一連のフローを実行
'        IN               : 移動元ヤード, 移動元部屋, 移動先ヤード, 移動先部屋, お客様番号
'        OUT              : エラーメッセージ（成功時は空文字）
'==============================================================================*
Private Function PerformFaceAuthTransferFlow(sourceYard As String, sourceRoom As String, destYard As String, destRoom As String, userCode As String) As String
    Dim strError As String
    Dim blnExists As Boolean
    Dim blnSuccess As Boolean
    
    On Error GoTo ErrHandler
    
    ' 1. 存在確認 (GET /check)
    blnExists = CallRoomCheckAPI(destYard, destRoom, strError)
    If strError <> "" Then
        PerformFaceAuthTransferFlow = "存在確認失敗: " & strError
        Exit Function
    End If
    
    If Not blnExists Then
        ' 2-A. 存在しない場合：新規登録 (POST /rooms)
        blnSuccess = CallRoomRegisterAPI(destYard, destRoom, userCode, strError)
    Else
        ' 2-B. 存在する場合：有効化・更新 (PUT /rooms)
        blnSuccess = CallRoomUpdateAPI(destYard, destRoom, userCode, strError)
    End If
    
    If Not blnSuccess Then
        PerformFaceAuthTransferFlow = "部屋登録/更新失敗: " & strError
        Exit Function
    End If
    
    ' 3. 移動実行 (POST /transfer)
    blnSuccess = CallRoomTransferAPI_New(sourceYard, sourceRoom, destYard, destRoom, userCode, strError)
    If Not blnSuccess Then
        PerformFaceAuthTransferFlow = "部屋移動実行失敗: " & strError
        Exit Function
    End If
    
    PerformFaceAuthTransferFlow = "" ' 成功
    Exit Function

ErrHandler:
    PerformFaceAuthTransferFlow = "フロー実行エラー: " & Err.Description
End Function

'--- 各種API呼び出し子関数 ---

Private Function CallRoomCheckAPI(yardCode As String, roomCode As String, ByRef strErr As String) As Boolean
    Dim res As Variant
    Dim regEx As Object
    
    res = SendFaceAuthRequest("GET", "/api/v1/external/system/rooms/check?yard_code=" & yardCode & "&room_code=" & roomCode)
    
    If VarType(res) = vbString Then
        ' 正規表現で "exists": true を検索 (スペースの有無を許容)
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.Pattern = """exists""\s*:\s*true"
        
        CallRoomCheckAPI = regEx.test(res)
        Set regEx = Nothing
    Else
        strErr = "HTTP Status " & res
        CallRoomCheckAPI = False
    End If
End Function

Private Function CallRoomRegisterAPI(yardCode As String, roomCode As String, userCode As String, ByRef strErr As String) As Boolean
    Dim body As String
    Dim res As Variant
    
    body = "{""yard_code"":""" & yardCode & """,""room_code"":""" & roomCode & """,""user_code"":""" & userCode & """}"
    res = SendFaceAuthRequest("POST", "/api/v1/external/system/rooms", body)
    
    If VarType(res) = vbString Then
        CallRoomRegisterAPI = True
    Else
        strErr = "HTTP Status " & res
        CallRoomRegisterAPI = False
    End If
End Function

Private Function CallRoomUpdateAPI(yardCode As String, roomCode As String, userCode As String, ByRef strErr As String) As Boolean
    Dim body As String
    Dim res As Variant
    
    ' お客様番号の更新と有効化(is_active: true)
    body = "{""yard_code"":""" & yardCode & """,""room_code"":""" & roomCode & """,""user_code"":""" & userCode & """,""is_active"":true}"
    res = SendFaceAuthRequest("PUT", "/api/v1/external/system/rooms", body)
    
    If VarType(res) = vbString Then
        CallRoomUpdateAPI = True
    Else
        strErr = "HTTP Status " & res
        CallRoomUpdateAPI = False
    End If
End Function

Private Function CallRoomTransferAPI_New(sourceYard As String, sourceRoom As String, destYard As String, destRoom As String, destUser As String, ByRef strErr As String) As Boolean
    Dim body As String
    Dim res As Variant
    
    body = "{" & _
           """source_yard_code"":""" & sourceYard & """," & _
           """source_room_code"":""" & sourceRoom & """," & _
           """dest_yard_code"":""" & destYard & """," & _
           """dest_room_code"":""" & destRoom & """," & _
           """dest_user_code"":""" & destUser & """" & _
           "}"
    res = SendFaceAuthRequest("POST", "/api/v1/external/system/transfer", body)
    
    If VarType(res) = vbString Then
        CallRoomTransferAPI_New = True
    Else
        strErr = "HTTP Status " & res
        CallRoomTransferAPI_New = False
    End If
End Function

'--- 共通送信処理 ---

Private Function SendFaceAuthRequest(method As String, path As String, Optional body As String = "") As Variant
    Dim req As Object
    Dim url As String
    
    On Error GoTo ErrHandler
    
    url = "https://dev-facekey-admin.kase3535.com" & path
    
    ' ServerXMLHTTPを使用してタイムアウトを設定
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    ' setTimeouts(resolve, connect, send, receive) - ミリ秒指定
    req.setTimeouts 2000, 2000, 2000, 2000
    
    req.Open method, url, False
    req.setRequestHeader "Content-Type", "application/json"
    req.setRequestHeader "x-api-key", "ff246566d3981d6d571f0011cbc1668f15172d6e73d6a65b827e681b6c6de492"
    
    req.send body
    
    If req.Status >= 200 And req.Status < 300 Then
        SendFaceAuthRequest = req.responseText
    Else
        SendFaceAuthRequest = req.Status ' エラー時はステータスコードを返す
    End If
    
    Set req = Nothing
    Exit Function

ErrHandler:
    SendFaceAuthRequest = "Error: " & Err.Description
    Set req = Nothing
End Function

'****************************  ended of program ********************************

