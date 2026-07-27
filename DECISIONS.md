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
- 運用: メンバー増減は名簿（本体は `JISHA_REFERRER_OPTIONS`、振込先は `SALESPEOPLE`）を編集し、本体はメニュー「フォーム管理 > 営業担当を同期」(`syncSalesRoster()`)を1回、振込先はエディタ or 一時Webアプリで `setup()` を1回実行する。`syncSalesRoster()` は非破壊で (1)全自社フォームの紹介者選択肢を再適用 (2)顧客管理SSに担当タブ追加 (3)SS2に担当タブ追加 (4)SS1に「総合_<担当>」タブ追加 を行う（タブ内データの再生成はしない＝空タブで用意し、`buildSalesRepStatusSheets`/`buildIntegratedRepSheets` の別実行で埋まる）。onOpen でも名簿変更を自動検知して紹介者選択肢を再適用。
- 江口裕人 追加時の本番反映（2026-07-16, version 101）: 上記4系統すべてに江口タブ/選択肢が揃ったことを確認済み。顧客管理SS・SS2・SS1(総合_江口裕人)・全14フォームの紹介者選択肢・振込先フォーム選択肢、いずれも江口裕人を含む。SS1の`総合_江口裕人`は空タブ（初案件が入り `buildIntegratedRepSheets` を実行すると埋まる）。
- 本番反映（2026-07-16, shinhogle）:
  - 本体GAS `clasp push`→`redeploy AKfycbznoqLywTwLGrictq4dTKkbx5kcfn8g8PF60QpRdjgGaOCqUTuQLlfvE3hiWkYrLBlr`（URL維持, version 99）。
  - `syncSalesRoster()` を一時メンテフック経由で実行し、紹介者選択肢を14フォーム更新・顧客管理SSに「江口裕人」タブ追加を確認（SS2は9タブ既存で追加なし）。実フォーム `?form=ouchikuraberu` の referrer が9択・江口裕人含むことを確認。一時フックはソース・本番とも除去済み。
  - 振込先フォームGAS `clasp push`、一時Webアプリ経由で `setup()` 実行→紹介営業担当が9択・江口裕人含むことを確認。一時Webアプリはundeploy・webapp設定/ doGet はソースから除去済み。
- 注意: 本コミットの `gas-project/Code.gs` には、前セッションで未コミットだった営業担当別ステータス表機能（`buildSalesRepStatusSheets`/`buildIntegratedRepSheets`/`repStatusRepAliasMap_` 等・SS2/SS1生成）が含まれる。当作業はこの `repStatusRepAliasMap_` を名簿駆動へ改修する形で同ファイルに乗るため分離不能。今回のデプロイ(version 99)でこれらの関数も本番に載ったが、いずれもトリガー無しの手動関数で、`buildSalesRepStatusSheets`（ゲート付きレビュー/実行待ち）は今回実行していない（休眠状態）。

## 2026-07-24: メンバー追加を「名簿編集＋同期1回」で完結させる（総合_江口裕人 が空だった件の恒久対策）

- 症状: SS1 の `総合_江口裕人` が空のまま更新されていなかった（ヘッダー行のみ）。`統合顧客管理` には 営業担当=江口裕人 の行が35件あった。
- 原因（コードのバグではない・手順の穴）: `syncSalesRoster()` は非破壊設計で**空タブを作るだけ**で、中身は `buildSalesRepStatusSheets()` / `buildIntegratedRepSheets()` の**別実行**が必要だった。この2関数はメニューにもトリガーにも無く、スクリプトエディタからの手動実行のみ。2026-07-16 に江口を追加した際に空タブが作られたあと再生成が走らず、そのまま放置されていた（ROADMAP の「残タスク=江口裕人 追加後の再生成」が未消化）。実際には江口だけでなく**全担当タブが古い**状態だった（例: 岩本 124行→169行、菅原 43→46、村井 16→17、柳沢 11→14）。
- 調査で確認した「反映済み」の範囲: 本体GAS・振込先フォームGAS・フォーム顧客管理GAS はいずれも**本番=ローカルで一致し 江口裕人 を含む**（`clasp pull` で差分0を確認）。`index.html` の `REFERRER_SUFFIX_MAP` の `-e` も反映済み。フロント2リポジトリも push 済み。つまり欠けていたのは「担当別データの再生成」だけ。
- Decision 1: `syncSalesRoster()` が最後に `buildSalesRepStatusSheets()` と `buildIntegratedRepSheets()` を続けて実行するようにした。メンバー増減は「名簿を編集 → メニュー『営業担当を同期』を1回」で完結する。各 build は個別に try/catch し、片方が落ちても他方と (1)〜(4) の結果を失わない。実行結果（担当別行数）はアラートとログに出す。
- Decision 2: 2つの再生成関数をメニューに追加し、エディタを開かずに単体実行できるようにした（`担当別ステータス表を再生成（SS2）` / `総合_担当タブを再生成（SS1）`）。
- Decision 3（副次的に発見した2つ目の穴）: SS2 の担当タブは**案件列の横幅がバラバラ**だった（柳沢/大島/小椋/細川/藤森/江口=17列、岩本=18列、菅原/村井=20列）。`buildSalesRepStatusSheets` は**そのタブ自身のヘッダー**を見て案件列を引くため、列が無い案件の実績は `unmatchedCase` に計上されて**黙って捨てられていた**。対策として:
  - `ensureRepStatusCaseColumns_()` を新設。メインSSの自社フォームから案件表示名を集め、名簿の全担当タブに不足案件列を追記して横幅を揃える。`syncSalesRoster()` から再生成の前に呼ぶ。メニュー単体実行用に `syncRepStatusCaseColumns()` も追加（新フォーム追加後に使う）。
  - `ensureRepStatusTabs_()` の見本タブ選択を「先頭に見つかったタブ」から「**案件列が最も多いタブ**」に変更。従来は古い狭いタブを複製するため、新メンバーだけ案件列が欠けた状態で作られていた（江口が17列だった原因）。
- 本番反映（2026-07-24, shinhogle, 本体GAS）: `clasp push` 後、**一時デプロイを新規作成して**そこから `syncSalesRoster()` をキー付きメンテフック経由で実行。本番デプロイ `AKfycbznoqLyw…`(@101) は**一切触っていない**（@HEAD デプロイは匿名アクセス不可だったため、使い捨てデプロイ @102/@103 を作成→実行→`undeploy`）。フック・キーはソースからも本番からも除去済み（`clasp pull` で live==local・`maint` 参照0を確認）。デプロイ一覧も元の7件に復帰。
- 実行結果: SS1 `総合_江口裕人` = 32行（0→32）。SS2/SS1 の全担当タブを再生成し、`unmatchedCase` は `{"出会えるエージェント":1}` → `{}` に解消。SS2 全9タブの案件列集合が完全一致（20列）になったことを確認。
- 未解決（ユーザー判断待ち）:
  - `excludedReferrers: {"藤井勇大": 3}` — 自社フォーム回答に紹介者名「藤井勇大」が3件あるが名簿に無く、SS2 で除外されている。新メンバーなら `JISHA_REFERRER_OPTIONS` に追加、誤入力なら回答側を修正する必要がある。
  - SS2 の `江口裕人` は0行のままだが、これは正常。SS2 は**フォーム回答の紹介者名**を集計する表で、江口を紹介者とする回答がまだ無いため。江口の32件は `統合顧客管理` の営業担当列由来で、SS1 側に正しく出ている。

## 2026-07-24: 藤井勇大 を営業名簿に追加（計10名・新手順の初適用）

- 背景: 同日の再生成で `excludedReferrers: {"藤井勇大": 3}` を検出。自社フォーム回答に紹介者名「藤井勇大」が3件あるが名簿に無く、SS2 から除外されていた。ユーザー判断により正式メンバーとして追加。
- 変更箇所（名簿の単一の変更点のみ）:
  - `gas-project/Code.gs`: `JISHA_REFERRER_OPTIONS` に `藤井勇大` を追加。`repStatusRepAliasMap_()` の `variants` に `"藤井勇大": ["藤井"]` を追加。**藤森宣哉 と苗字が近いが、`resolveRepCanonical_` は `normalizeName` の完全一致のみで引く**ため衝突しない（前方一致・あいまい照合は使っていない）。
  - `payout-form/Code.gs`: `SALESPEOPLE` に追加。
  - `フォーム顧客管理/main.js` と `index.html`: `SALES_STAFF` に追加、`SALES_STAFF_ALIASES` に `'藤井': '藤井勇大'` を追加（2ファイル同一内容を維持）。
  - `index.html`(アフィリンク フロント) の `REFERRER_SUFFIX_MAP` は**未変更**。固定紹介者リンクは現状4名(-s/-m/-i/-e)のみで、必要になったら記号を割り当てる。
- 実施手順（改修後の新手順を初めて通しで適用）: 名簿を編集 → `clasp push` → メニュー相当の `syncSalesRoster()` を1回実行、で全系統が揃った。旧手順のような再生成の別実行は不要になっている。
- 本番反映（2026-07-24, shinhogle）:
  - アフィリンク本体: `clasp push` 後、使い捨てデプロイ(@104)経由で `syncSalesRoster()` を実行し `undeploy`。**本番 `AKfycbznoqLyw…`(@101) は未変更**。フォームの紹介者選択肢は設定シートのデータ由来（`readConfig` が列6を読む）で、doGet/doPost は `JISHA_REFERRER_OPTIONS` を参照しないため、本番の再デプロイは不要と判断した。
  - 振込先フォーム: `appsscript.json` に webapp 設定と鍵付き `doGet` を一時追加 → `clasp push` → 使い捨てデプロイ(@4)で `setup()` 実行 → `undeploy`、フックと webapp 設定をソースから除去して再 push。デプロイは元の1件(@HEAD)に復帰。
  - フォーム顧客管理: `clasp push` → `clasp redeploy AKfycbxqy6u9…`(@18→@19、URL維持)。index.html は GAS が配信するため再デプロイが必要。
- 結果:
  - 紹介者選択肢17フォーム更新。顧客管理SS に `藤井勇大` タブ追加、SS1 に `総合_藤井勇大` タブ追加（SS2 の担当タブは既存だった）。
  - SS2 `藤井勇大` = 1行 / SS1 `総合_藤井勇大` = 1行。**`excludedReferrers` が `{}` に解消**（3件の回答は同一顧客のため集計行は1）。`unmatchedCase` も `{}` のまま。
  - SS1 `柳沢悠貴` が 14→13行。該当顧客のアフィリンク紹介者が藤井勇大のため、担当がアフィリンク優先で付け替わったもので設計どおり。
  - 振込先フォーム実物の選択肢が10名になったことを公開URLのHTMLで確認。
  - 3プロジェクトとも `clasp pull` で live==local・フック残骸0を確認。

## 2026-07-27: 特別緊急クエスト（ノムコム・30件）を緊急クエストと並走で開始

- 背景: 同日開始の緊急クエスト（5案件・全体80件・7/27 10時〜7/31）に加え、ノムコム単独の特別枠を7/31まで走らせる依頼。全体30件、内訳は 岩本拓也15 / 菅原貴博10 / 江口裕人3 / 藤井勇大2（合計がちょうど30）。
- Decision 1: 既存の緊急クエストに項目を足し込まず、**独立したレポート**（`specialQuestReport` / `buildSpecialQuestMessage_` / `ensureSpecialQuestTriggers`）として実装した。1通のメッセージに両クエストを詰めると長大になり、終了後の撤去も互いに絡むため。送信時刻は緊急クエストと同じ毎日13時・20時で、LINEには2通が並ぶ。
- Decision 2: 集計対象は `SPECIAL_QUEST_FORM = "nomukomu"`（表示名 ノムコム、自社=house）の1シートのみ。`getConfigSheetByCode` で直接引く。期間は `SPECIAL_QUEST_START_AT = 2026/07/27 00:00:00` 〜 `SPECIAL_QUEST_END_STR = 2026/07/31 23:59:59`。**本日分から数える**ため、それ以前の既存11件は対象外。
- Decision 3: 紹介者名の照合は緊急クエスト（`normalizeName` 完全一致のみ）ではなく `repStatusRepAliasMap_()` + `resolveRepCanonical_()` を再利用した。「岩本」など苗字のみ、「松田恵美（岩本拓也）」など括弧内表記も正規名へ寄る。名簿10名の正規名以外は個人ノルマに載らず、全体件数のみに入って「その他メンバー」行に出る。
- Decision 4: 管理ルート（`?quest_admin=<キー>`）に `quest=special` を追加し、`preview|send|setup|status` を緊急/特別で切り替えられるようにした。二重送信ガードのプロパティも `SPECIAL_QUEST_LAST_SENT` で独立。
- 本番反映（2026-07-27, shinhogle）: `clasp push` 後、**一時デプロイ `AKfycbwn--vmqogeWbeonVQ…` を @108→@109 に redeploy**（管理ルート専用の使い捨て。本番 `AKfycbznoqLyw…`(@101) は未変更）。旧@108は古いコードを配信していてクエスト切替が効かなかったため、redeployが必須だった。
- 検証: `action=preview&quest=special` で本文生成（622文字・改行20・リテラル `\n` 0件）を確認 → `action=setup&quest=special` でトリガー2本作成 → `action=send&quest=special` でLINEグループへ初回送信成功。緊急クエスト側のトリガー2本・`lastSent` が無傷であることを `action=status`（quest未指定）で確認。
- 初回時点の実績: 全体3件（岩本拓也1・菅原貴博2）。
- 撤去予定: 7/31終了後、両クエストのトリガーは初回実行時に自動削除される（`todayJst > END` で `remove…Triggers_()`）。管理ルート・一時デプロイ・定数ブロックは手動で撤去する。
