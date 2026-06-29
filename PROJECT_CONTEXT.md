# Project Context

## Identity

- Project: アフィリンク
- Organization: 市場作り
- Workspace path: C:\Users\shinh\OneDrive\デスクトップ\AI\プロジェクト\市場作り\アフィリンク

## Purpose

不動産アフィリエイトの申請フォームを管理し、フォーム送信ごとに回答を記録・LINE通知・顧客管理・広告主への成果報告を行うシステム。営業マンが紹介した顧客がフォームから申請すると、すべての処理が自動で走る。

## Users And Stakeholders

- **管理者**: shinhogle@gmail.com（市場作りプロジェクト全般）
- **営業マン（自社）**: 柳沢悠貴, 岩本拓也, 菅原貴博, 村井亮介, 大島雅史, 小椋裕也, 細川貴弘, 藤森宣哉
- **150件クエスト対象営業3名**: 岩本拓也, 菅原貴博, 村井亮介（各50件ノルマ、期限2026/06/30）
- **広告主**: 成果管理シート（ADVERTISER_SS_ID）にアクセスする第三者

## System Scope

### フォーム一覧（代表的な自社フォーム）

- ouchikuraberu（おうちクラベル）
- home4ufudosan（HOME4U不動産）
- home4utochi（HOME4U土地活用）
- srerealty（SREリアリティー）
- リビンマッチ（地域フィールド追加済み）
- その他複数の自社フォーム（代理店コードなし = AGENCY_DEFAULT = "house"）

代理店コードが設定されているフォームは顧客管理・広告主シート・クエスト集計の対象外。

### 入力保持・紹介者固定リンク

- フォーム入力中の値はブラウザの `localStorage` に自動保存される。
- アフィリエイトリンク先で作業して戻った場合も、名前・紹介者・クリック時刻などは復元される。
- スクショファイルはブラウザ仕様上復元できないため、戻った後に再選択が必要。
- 送信完了後、保存済みの入力途中データは削除される。
- 紹介者固定リンクは全フォーム共通で使える。新規案件も同じルール。
  - `?form=<フォーム記号>-s` → `菅原貴博`
  - `?form=<フォーム記号>-m` → `村井亮介`
  - `?form=<フォーム記号>-i` → `岩本拓也`
- 例: `https://kazu02.github.io/affiliate-form/?form=ouchikuraberu-s`
- 固定リンクでは紹介者欄は自動選択され、ユーザーは変更できない。

### スプレッドシート構成

| 種別 | 説明 |
|------|------|
| アフィリエイト管理SS | メインSS。設定シート（`設定_*`）に設定とフォーム回答を両方格納 |
| 顧客管理SS | 営業マン別シート。顧客×案件のステータス管理 |
| 広告主成果管理SS | `1bnERIRl4-VmQ2QP9IwxuPco64huKzNC5qxHfyvCBUVg`。月別シート（YYYYMM形式） |
| 代理店SS | 代理店ごとに別SS。代理店コードで管理 |

## Architecture And Operations

### GAS Web App

- Script ID: `1OqsufxjJqAfj0nvAmE20miZCMlZ5BCI0WxSElMuAqkpwMUt6YuWCPgM4`
- Deploy ID: `AKfycbznoqLywTwLGrictq4dTKkbx5kcfn8g8PF60QpRdjgGaOCqUTuQLlfvE3hiWkYrLBlr`
- executeAs: USER_DEPLOYING, access: ANYONE_ANONYMOUS
- ソースファイル: `gas-project/Code.gs`

### デプロイ手順（必須2ステップ）

```
clasp push
clasp deploy -i <Deploy ID>
```

- `clasp run` は使用不可
- テスト用: `REDIRECT=$(curl -s -o /dev/null -w "%{redirect_url}" "$URL") && curl -s "$REDIRECT"`
- GASアカウント: shinhogle@gmail.com（毎回 `clasp logout` → `clasp login` で確認すること）

### フォーム送信時の処理フロー（doPost）

1. 回答をメインSSの設定シート（G列以降）に記録
2. 代理店SSにも同期
3. LINEグループ通知
4. 顧客管理シートにupsert（自社フォーム・紹介者名ありの場合）
5. 広告主成果管理シートにリアルタイム書き込み

### 広告主成果管理シートの列構成（現行）

| 列 | 内容 |
|----|------|
| A | 受信日時 |
| B | 広告名（フォーム表示名） |
| C | お名前 |
| D | 紹介者名 |
| E | スクショURL |
| F | トラッキング漏れ（チェックボックス。承認列が❌/✖/×系の場合、GASインポートがチェック） |
| G | 承認区分（チェックボックス。承認列が⭕/○系の場合、GASインポートがチェック） |

- 行1: タイトル行（空またはラベル）
- 行2: ヘッダー行
- 行3以降: データ行

### ScriptProperties（主要キー）

- `LINE_CHANNEL_TOKEN`: LINEチャンネルアクセストークン
- `LINE_GROUP_ID`: 通知先LINEグループID
- `SPREADSHEET_ID`: メインSSのID
- `SCREENSHOT_FOLDER_ID`: スクショ保存Driveフォルダ

### トリガー

- 毎朝8時: `dailyReport`（前日の申請をLINE通知）
- 毎朝8時: `campaignReport`（30件クエスト進捗）
- 毎朝8時: `quest150Report`（150件クエスト進捗）

### 重要な定数（Code.gs）

```javascript
const ANSWER_START_COL    = 7;          // G列から回答記録
const AGENCY_DEFAULT      = "house";    // 自社フォームの識別コード
const ADVERTISER_SS_ID    = "1bnERIRl4-VmQ2QP9IwxuPco64huKzNC5qxHfyvCBUVg";
const QUEST150_MONTH      = "2026/06";
const QUEST150_SALESPEOPLE = ["岩本拓也", "菅原貴博", "村井亮介"];
```

## Constraints

- GASアカウントは必ず shinhogle@gmail.com を使うこと（3s3.cube@gmail.com と混同しないよう注意）
- `clasp run` は使用不可。テストはcurl 2ステップで行う
- 広告主シートのF列はトラッキング漏れ、G列は承認区分。既存データ一括インポート時は承認列の記号を見て、⭕/○系ならG列、❌/✖/×系ならF列をチェックする。△や空欄など他の値は両方false。
- 広告主シートの行2はヘッダー行。データは行3以降に書き込む
- インポート関数は行3以降をclearContentしてから再書き込みする（deleteRowsは「全非固定行削除」エラーになる場合がある）
- 代理店コードありのフォームは顧客管理・広告主シート・クエスト進捗の集計対象外
- セキュリティ: 個人情報・認証情報・生の顧客データはメモリファイルに記録しない

## Sources Of Truth

- Current code and configuration
- Git history, when available
- `DECISIONS.md`
- `ROADMAP.md`
- Current `CODEX_TASK.md`, when present
