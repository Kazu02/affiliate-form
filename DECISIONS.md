# Decisions

## 2026-06-12: Adopt File-Based Project Memory

- Decision: Store durable project context in `PROJECT_CONTEXT.md`, `DECISIONS.md`, and `ROADMAP.md`.
- Reason: Claude and Codex sessions do not reliably share complete conversation history.
- Consequence: Important conclusions from project chats must be summarized into these files.
- Safety: Do not copy credentials, personal data, confidential raw data, or full chat transcripts.

---

## 2026-06-12: 広告主シートの列構成を5列→6列に拡張（広告名追加）

- Decision: 広告主成果管理シートにB列「広告名」（フォーム表示名）を追加。列構成をA=受信日時, B=広告名, C=お名前, D=紹介者名, E=スクショURL, G=承認とした。
- Reason: 広告主から「どの案件の申請か」を判別できるようにする要件が発生。
- Consequence: `writeToAdvertiserSheet` の引数と `importExistingToAdvertiserSheet` のバッファ構造を変更。既存データは再インポートで上書き。

## 2026-06-12: 広告主シートの承認列をF→G列に変更

- Decision: 承認区分（GAS書き込み）をF列からG列に移動。
- Reason: 広告主がF列に「トラッキング漏れ」列を追加したため、GASの書き込み対象をG列にずらした。
- Consequence: `writeToAdvertiserSheet` および `importExistingToAdvertiserSheet` 内の列番号を6→7に変更。

## 2026-06-29: 広告主シートのトラッキング漏れを承認記号から反映

- Decision: 既存データを広告主成果管理SSへインポートする際、メインSSの承認列が⭕/○系ならG列「承認区分」、❌/✖/×系ならF列「トラッキング漏れ」にチェックする。
- Reason: 広告主提出用シートで、非承認のうちトラッキング漏れを明確に区別する必要があるため。
- Consequence: △や空欄など他の値は無視し、F/Gともfalseにする。リアルタイム新規書き込みは承認情報がまだないためF/Gともfalseで追加する。
- Operation: 2026-06-29に広告主成果管理SSをバックアップ後、202604=18件、202605=163件、202606=160件を再インポート。合計341件、承認139件、トラッキング漏れ57件、無視145件。

## 2026-06-12: 広告主シートインポートの行構造を確定（行2=ヘッダー、行3以降=データ）

- Decision: 広告主シートは行1がタイトル、行2がヘッダー、行3以降がデータ。インポート時は行3以降を `clearContent` してから書き込む。
- Reason: 当初 `deleteRows` を使ったところ「固定されていない行をすべて削除できない」エラーが発生。また `startRow=2` で上書きするとヘッダーが消えた。
- Consequence: `clearContent` を使い、書き込み開始行は `getAdvertiserNextRow()` で動的に決定（ヘッダーのA列が非空なので常に3行目を返す）。

## 2026-06-12: 月別シート自動作成機能の追加

- Decision: `importExistingToAdvertiserSheet` 実行時、対象月のシートが存在しない場合は既存の最古シートをコピーして自動作成する（`createAdvertiserMonthSheet`）。
- Reason: 202604シートを手動作成せずにインポートできるようにするため。
- Consequence: テンプレートコピー後に行3以降を削除し、月別シートを昇順に並べる処理も含む。

## 2026-06-12: 150件クエスト実装

- Decision: 全自社フォーム合計150件（2026/06末まで）を新クエストとして追加。3名（岩本拓也・菅原貴博・村井亮介）に各50件個人ノルマ。毎朝8時にLINE報告。
- Reason: 30件クエスト（4案件×30件）から拡張された新しい目標。
- Consequence: `quest150Report` 関数と専用トリガー `ensureQuest150Trigger` を追加。

## 2026-06-12: 顧客管理シートの自動連携

- Decision: フォーム送信時（doPost）に紹介者名が存在する自社フォームは、顧客管理シートへ自動upsertする。
- Reason: 営業マン別の顧客管理を手動入力なしで維持するため。
- Consequence: `upsertCustomerRow` を doPost から呼び出す。代理店フォームは対象外。

## 2026-06-12: 回答ヘッダー同期機能（fixAnswerHeaders）

- Decision: フォームにフィールドを追加した場合、`fixAnswerHeaders` を手動実行してヘッダーと既存データ行の列を修正する。
- Reason: リビンマッチに「地域」フィールドを追加した後、スクショURLが地域列に格納されていた問題が発覚。
- Consequence: 旧フォーマット行は新フォーマットに変換（空列を挿入してスクショURLを正しい位置に移動）。

## 2026-06-15: 新案件（自社フォーム）を顧客管理シートへ自動反映

- Decision: 顧客管理シートの案件列はシート作成時のみ確定していたため、新しい自社フォーム（案件）追加後も既存の顧客管理シートに反映されなかった。これを解消する。
- Reason: アフィリエイト管理で案件を増やしても、`upsertCustomerRow` が該当案件列を見つけられず（`caseCol <= 0` で早期 return）申請が反映されない問題があった。
- Consequence:
  - `syncCustomerManagementCases()` を新設（メニュー「顧客管理の案件列を同期」）。各営業マンシートのヘッダーと最新案件（`getJishaForms()`）を突き合わせ、不足案件列を書式・幅・入力規則付きで追加。
  - `upsertCustomerRow` を改修し、案件列が無ければ自動追加（`addCaseColumnToSheet_`）してから書き込む。これで以後の新案件はフォーム送信時に自動反映される。
  - デプロイは shinhogle@gmail.com で `clasp push`→`clasp redeploy AKfycbznoqLywTwLGrictq4dTKkbx5kcfn8g8PF60QpRdjgGaOCqUTuQLlfvE3hiWkYrLBlr`。誤アカウント（3s3.cube）でのデプロイは executeAs の都合で本番破壊につながるため厳禁。

## 2026-06-22: 入力途中保存と紹介者固定リンクを全案件共通にする

- Decision: 静的フォーム側で入力途中データを `localStorage` に自動保存し、`?form=<フォーム記号>-s/-m/-i` の紹介者固定リンクを全フォーム共通で解釈する。
- Reason: ユーザーがアフィリエイトリンク先で10分以上作業して戻った際に入力内容が消える問題と、紹介者選択ミスの多発を防ぐため。
- Consequence:
  - `-s` は `菅原貴博`、`-m` は `村井亮介`、`-i` は `岩本拓也` として紹介者欄を固定する。
  - GASへ設定取得するときは末尾サフィックスを外したフォーム記号を使うため、新規案件でも同じURLルールが使える。
  - スクショファイルはブラウザ仕様で復元不可。送信前に再選択が必要。
  - 本番は GitHub Pages (`origin/master`) に `c6e0b07` として反映済み。

## 2026-07-16: アフィリエイト報酬 振込先フォームを新設（独立GASプロジェクト）

- Decision: お客様から振込先情報を集める Google フォームを新規作成し、送信ごとに統合顧客管理ブック(SS1 = `1aaiCIDQIkrp_Ado5aKua_PTEQq4jr1UWqpuLQXwpemI`)の各営業担当タブ「総合_<担当>」とマスター「統合顧客管理」の、`持ち家かどうか`(K列)の直後に自動追記する。
- 実装場所: **既存の本番GASプロジェクトとは分離した独立スタンドアロンプロジェクト**（scriptId `1X30nki2jpev7aD7EjVgg2xrpmwEBOCeW5a0ZgR07RD-KDLpA3tpuWHj_`、ローカル `payout-form/`）。本番Webアプリ(申請フォーム側)を触らずに済むため。
- Reason: アフィリエイト報酬の支払いに振込先が必要。営業マンが手入力せず、お客様のフォーム回答から自動で顧客行に紐づける。
- 照合方式: フォームのフルネームを `normalizeName`(全角/半角スペース・不可視文字除去＋ひらがな化＋小文字化)で正規化して突合。完全一致→先頭一致(苗字のみ等)→レーベンシュタイン距離1、の順であいまい照合。紹介者(営業担当)は8名から選択式で入力させ、同名顧客の絞り込みに使う。
- 追記列(6列): `金融機関名 / 支店名 / 預金種目 / 口座番号 / 口座名義(カナ) / 振込先登録日時`。口座番号・店番の**先頭ゼロ落ちを防ぐため書き込み前にセル書式を`@`(テキスト)にする**。
- 耐久性: 「総合_<担当>」は `buildIntegratedRepSheets()` が `tab.clear()` で再生成するため、振込先が消えないよう**マスター(統合顧客管理)にも同時書き込み**する。マスターに列があれば再生成時に総合_タブへ複写されて保持される。既存の `buildIntegratedRepSheets` は列名でインデックスを引くため、6列追加でも破綻しない。
- 照合できない回答は SS1 の「振込先_未照合ログ」タブに記録する。
- 認証の要点: **clasp login はForms/Sheets等のセンシティブスコープをGoogleがブロックする**ため使用不可。スクリプトエディタで `runSetupAndSelfTest` を1回実行してオーナー(shinhogle)承認。検証・後始末はキー保護の一時Webアプリ経由で自走し、完了後にデプロイ削除・doGet/webapp設定をソースから除去済み。
- フォームURL(回答用): `https://docs.google.com/forms/d/e/1FAIpQLSd3AXOJrEwker1hKLATVtQpIbSZ4wWvAS5cAl-bbYQ46UKbZg/viewform`
- E2E検証済み: 「田中 太郎」(スペース入り)→「田中太郎」照合成功、マスター＋総合_藤森宣哉の両方に先頭ゼロ保持で書込→クリア確認。トリガー(onFormSubmit, shinhogle所有)発火も確認。

## 2026-07-16: 営業名簿を単一の変更点に集約し 江口裕人 を追加（本番反映済み）

- Decision: 営業担当（紹介者）の一覧を1箇所に集約し、メンバー増減を最小操作で行えるようにする。
  - 本体GAS: 定数 `JISHA_REFERRER_OPTIONS`（Code.gs）を唯一の名簿とする。`repStatusRepAliasMap_()` は名簿から正規名の自己対応を自動生成し、苗字・表記ゆれのみ手動 `variants` で補う方式に変更。
  - 振込先フォーム: `payout-form/Code.gs` の `SALESPEOPLE` を名簿とし、`setup()` 再実行で既存Googleフォームの「紹介営業担当」選択肢を `syncFormChoices_()` で同期。
  - フロント: `index.html` の `REFERRER_SUFFIX_MAP` に `-e → 江口裕人` を追加（固定紹介者リンク用）。
- 追加メンバー: `江口裕人`（本体名簿・振込先名簿とも計9名）。
- 運用: メンバー増減は名簿（本体は `JISHA_REFERRER_OPTIONS`、振込先は `SALESPEOPLE`）を編集し、本体はメニュー「フォーム管理 > 営業担当を同期」(`syncSalesRoster()`)を1回、振込先はエディタ or 一時Webアプリで `setup()` を1回実行する。`syncSalesRoster()` は非破壊で (1)全自社フォームの紹介者選択肢を再適用 (2)顧客管理SSに担当タブ追加 (3)SS2に担当タブ追加 を行う。onOpen でも名簿変更を自動検知して紹介者選択肢を再適用。
- 本番反映（2026-07-16, shinhogle）:
  - 本体GAS `clasp push`→`redeploy AKfycbznoqLywTwLGrictq4dTKkbx5kcfn8g8PF60QpRdjgGaOCqUTuQLlfvE3hiWkYrLBlr`（URL維持, version 99）。
  - `syncSalesRoster()` を一時メンテフック経由で実行し、紹介者選択肢を14フォーム更新・顧客管理SSに「江口裕人」タブ追加を確認（SS2は9タブ既存で追加なし）。実フォーム `?form=ouchikuraberu` の referrer が9択・江口裕人含むことを確認。一時フックはソース・本番とも除去済み。
  - 振込先フォームGAS `clasp push`、一時Webアプリ経由で `setup()` 実行→紹介営業担当が9択・江口裕人含むことを確認。一時Webアプリはundeploy・webapp設定/ doGet はソースから除去済み。
- 注意: 本コミットの `gas-project/Code.gs` には、前セッションで未コミットだった営業担当別ステータス表機能（`buildSalesRepStatusSheets`/`buildIntegratedRepSheets`/`repStatusRepAliasMap_` 等・SS2/SS1生成）が含まれる。当作業はこの `repStatusRepAliasMap_` を名簿駆動へ改修する形で同ファイルに乗るため分離不能。今回のデプロイ(version 99)でこれらの関数も本番に載ったが、いずれもトリガー無しの手動関数で、`buildSalesRepStatusSheets`（ゲート付きレビュー/実行待ち）は今回実行していない（休眠状態）。
