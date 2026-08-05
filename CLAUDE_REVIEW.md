# CLAUDE_REVIEW — 営業担当別 案件ステータス表（rep-status 機能）

- レビュー種別: **手動・事後レビュー**（Claude 対話セッション。オーケストレーターの review フェーズではなく、営業名簿集約作業(2026-07-16)に伴い実施）
- 対象: `gas-project/Code.gs` の `buildSalesRepStatusSheets()` / `buildIntegratedRepSheets()` / `repStatusRepAliasMap_()`(名簿駆動へ改修済) / `resolveRepCanonical_()` / `cleanupIntegratedPhantomRows()`、および今回追加の `syncSalesRoster()` / `ensureCustomerMgmtTabs_()` / `ensureRepStatusTabs_()` / `applyReferrerSelectToJishaSheets(force)`
- 前提: 本コードは Codex により実装済み・本番デプロイ済み(version 99)・実行実績あり（`CODEX_REPORT.md`: SS2=承認183/申請85/非承認88、SS1 総合_ タブ9枚生成、幽霊行3件クリーンアップ）。本レビューは事後の妥当性確認。

## 良い点（設計上の安全策が効いている）

1. **誤ブック書き込み防止**: `buildSalesRepStatusSheets` / `buildIntegratedRepSheets` は書き込み前に `getOrCreateSpreadsheet().getId()` を `REP_STATUS_MAIN_ID` と assert し、不一致なら中止。ID ハードコードの `openById` に依存せず既存アクセサ経由。
2. **列は名前で特定**（`受信日時`/`お名前`/`紹介者名`/`承認`/`顧客名`/案件表示名）。列位置固定に依存しないため、列追加でも破綻しない。
3. **代理店フォーム除外**が `agencyCode !== AGENCY_DEFAULT("house")` の実データ判定で、タブ名の文字列一致に依存しない（正しい手段）。
4. **非破壊更新**: SS2 は対象タブの2行目以降を `clearContent` して再書き込み（他タブ・他ブックは読み取りのみ）。SS1 の総合_タブは `clear()`→再生成。いずれも入力規則を `clearDataValidations()` してから `{月}月{状態}` を書く。
5. **最新優先**: 同一顧客×案件は受信日時(rtKey)の新しい方を採用。名寄せは `normalizeName`。
6. **担当解決**: `resolveRepCanonical_` は別名マップ＋括弧内担当名（「松田恵美（岩本拓也）」等）も救済。名簿(`JISHA_REFERRER_OPTIONS`)外の担当も取りこぼさない設計（SS1側は生値でタブ化）。
7. **`cleanupIntegratedPhantomRows` の安全設計**: 名前＋担当＋「アフィリンクのみ」の内容一致が**ちょうど1件**の行だけ削除、候補>8で中止、削除前スナップショットをログ、フィルタ除去→削除→`flush()`→フィルタ再作成（`deleteRow` がフィルタ有時に無言失敗する GAS の罠に対処）、下から削除で行ズレ防止。副作用のない「アフィリンクのみ」行に限定＝他CRM情報の損失なし。

## 指摘・改善提案（いずれもブロッカーではない）

1. **[minor] SS2 は担当タブが実在しないとスキップ**（`buildSalesRepStatusSheets`）。名簿に新メンバーを足すと、その担当の SS2 タブが無い限りデータが載らない。現状は今回追加した `ensureRepStatusTabs_()`（メニュー「営業担当を同期」）で先にタブを作る運用でカバー済み。将来ハードニングするなら `buildSalesRepStatusSheets` 冒頭で `ensureRepStatusTabs_()` を呼び自己完結にすると、手順が1つ減る。SS1 側(`buildIntegratedRepSheets`)は総合_タブを自動生成するため既に自己完結。
2. **[minor] `cleanupIntegratedPhantomRows` は 2026-07-06 調査の3件をハードコード**した一回限りのクリーンアップ。既に実行済み（再実行しても一致1件でなければSKIPするので安全だが）、恒常運用の関数ではない旨をコメントで明示すると誤用防止になる。
3. **[note] 状態ラベルの決定**は `getAdvertiserApprovalFlags`（承認=⭕系→承認 / トラッキング漏れ=❌系→非承認 / それ以外→申請）に依存。広告主シート取り込みロジックと同一基準で一貫。仕様変更時は両者を同時に見直すこと。

## 今回追加分（名簿集約）のレビュー

- `applyReferrerSelectToJishaSheets(force)`: `REFERRER_OPTIONS_APPLIED` に適用済み名簿値を保持し、`force` か値変更時のみ再適用。onOpen 自動再適用が冪等で安全。旧フラグ `REFERRER_SELECT_APPLIED` も後方互換で維持。
- `ensureCustomerMgmtTabs_()` / `ensureRepStatusTabs_()`: 既存タブは触らず、名簿に無いタブのみ見本タブ（先頭/既存担当タブ）からヘッダー・書式・列幅・入力規則を複製して追加。非破壊。テンプレートが無い場合(SS2に担当タブ皆無)も `insertSheet` だけは行いログに残る。
- `syncSalesRoster()`: 上記3つを束ねた1クリック同期。`SpreadsheetApp.getUi()` はWeb実行時に例外になるが try/catch 済みで問題なし（今回の一時メンテフック経由実行でも検証済み）。

## VERDICT: APPROVED

事後レビューとして承認。既に本番稼働・結果検証済みで、設計上の安全策（ID assert・名前ベース列・非破壊・削除の多重ガード）が妥当。改善提案1（`buildSalesRepStatusSheets` の SS2 タブ自己生成）だけ、次に本機能へ触れる際の軽微なハードニング候補として記録する。

## 残タスク（Codex へ引き継ぐ場合の指示）

- 江口裕人 追加後の **再生成・検証**（初案件が入ってから）: `ensureRepStatusTabs_()`→`buildSalesRepStatusSheets()`→`buildIntegratedRepSheets()` を実行し、`総合_江口裕人`/SS2 江口タブに `{月}月{状態}` が正しく出るか検証。**本番SS2/SS1書き込みのゲート付き実行**のため、ユーザーGO必須。
- 実行は必ず `npm.cmd run agents:queue -- affililink`（オーケストレーター経由＝実Codex）で行うこと。Claude 対話での代行は不可（役割分離）。
