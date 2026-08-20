// =============================================
// アフィリエイトフォーム - メインスクリプト
// =============================================

const DRIVE_FOLDER        = "アフィリエイト_スクショ";
const CONFIG_PREFIX       = "設定_";
const FORM_BASE_URL       = "https://kazu02.github.io/affiliate-form/";
const ANSWER_START_COL    = 7; // G列から回答を記録
const MANAGEMENT_SHEET    = "管理"; // 管理シート名
const LINE_PUSH_API       = "https://api.line.me/v2/bot/message/push";
const ADVERTISER_SS_ID    = "1bnERIRl4-VmQ2QP9IwxuPco64huKzNC5qxHfyvCBUVg";
const ADVERTISER_DATA_START_ROW = 3; // 行1=タイトル / 行2=ヘッダー / 行3以降=データ
// 広告主シート保守ルート（doGet の adv_admin）のキー。月指定の再生成と件数確認だけを許可する。
// このリポジトリは GitHub Pages の配信元のため public。キーをソースに書くと
// 誰でも読めてしまうので、値は ScriptProperties に置く。
// 未設定のあいだはルートを一切通さない（fail-closed）。
const ADVERTISER_ADMIN_KEY_PROPERTY = "ADVERTISER_ADMIN_KEY";

function advertiserAdminKey_() {
  return PropertiesService.getScriptProperties().getProperty(ADVERTISER_ADMIN_KEY_PROPERTY) || "";
}

// 代理店関連
const AGENCY_KEY          = "代理店コード";
const AGENCY_DEFAULT      = "house";        // 自社直営業の内部コード
const AGENCY_DEFAULT_NAME = "自社直営業";   // 自社直営業のSS表示名
const AGENCY_PREFIX       = "代理店_";      // 代理店SSの名前プレフィックス
const AGENCY_FOLDER       = "代理店スプシ"; // Drive保管フォルダ名
const AGENCY_PATTERN      = /^[a-zA-Z0-9_]+$/;
const AGENCY_PROP_PREFIX  = "AGENCY_SS_";   // ScriptProperties キー prefix

// フォーム名（表示用・日本語）
const FORM_NAME_KEY     = "フォーム名";    // 設定シートのキー
const FORM_CODE_HEADER  = "フォーム記号";  // 回答データG1ヘッダー
const FORM_CODE_PATTERN = /^[a-zA-Z0-9_]+$/;
// ===== 営業担当（紹介者）名簿：単一の変更点 =====
// メンバーの増減はこのカンマ区切り文字列を編集し、スプレッドシートのメニュー
// 「フォーム管理 > 営業担当を同期」を1回実行するだけ。
// これで (1)全自社フォームの紹介者選択肢 (2)顧客管理SS/SS2の担当タブ が揃う。
const JISHA_REFERRER_OPTIONS = "柳沢悠貴,岩本拓也,菅原貴博,村井亮介,大島雅史,小椋裕也,細川貴弘,藤森宣哉,江口裕人,藤井勇大";

// 150件クエスト設定
const QUEST150_TARGET_TOTAL = 150;
const QUEST150_TARGET_INDIV = 50;
const QUEST150_END_STR      = "2026/06/30";
const QUEST150_MONTH        = "2026/06";
const QUEST150_SALESPEOPLE  = ["岩本拓也", "菅原貴博", "村井亮介"];

// 30件クエスト キャンペーン設定
const CAMPAIGN_FORMS = ["ouchikuraberu", "home4ufudosan", "home4utochi", "srerealty"];
const CAMPAIGN_FORM_NAMES = {
  "ouchikuraberu": "おうちクラベル",
  "home4ufudosan":  "HOME4U不動産",
  "home4utochi":    "HOME4U土地活用",
  "srerealty":      "SREリアリティー"
};
const CAMPAIGN_TARGET  = 30;
const CAMPAIGN_END_STR = "2026/05/31";

// 緊急クエスト設定（2026/07/27〜07/31・毎日13時/20時にLINE報告）
// 対象5案件の7月合計を 岩本拓也40件/その他40件/全体80件 で追いかける
const EMERGENCY_FORMS = ["iekatsu", "livingmuch", "iei", "hokenmammoth", "deaeru"];
const EMERGENCY_FORM_NAMES = {
  "iekatsu":      "いえカツLIFE",
  "livingmuch":   "リビンマッチ",
  "iei":          "不動産一括査定イエイ",
  "hokenmammoth": "保険マンモス",
  "deaeru":       "出会えるエージェント"
};
const EMERGENCY_TARGET_TOTAL   = 80;
const EMERGENCY_TARGET_IWAMOTO = 40;
const EMERGENCY_TARGET_OTHERS  = 40;
const EMERGENCY_IWAMOTO  = "岩本拓也";
const EMERGENCY_END_STR  = "2026/07/31";
const EMERGENCY_START_AT = "2026/07/27 10:00:00"; // この日時(JST)以降の受信のみカウント（過去分は含めない）
// EMERGENCY_ADMIN_KEY（quest_admin ルート用）はクエスト終了(2026/07/31)に伴い撤去した。
// 値は public リポジトリに露出していたため、再開する場合も同じ値は使わない。

// 特別緊急クエスト設定（2026/07/27〜07/31・毎日13時/20時にLINE報告）
// ノムコム1案件を 全体30件（岩本15/菅原10/江口3/藤井2）で追いかける
const SPECIAL_QUEST_FORM      = "nomukomu";
const SPECIAL_QUEST_FORM_NAME = "ノムコム";
const SPECIAL_QUEST_TARGET_TOTAL = 30;
const SPECIAL_QUEST_MEMBERS = [
  { name: "岩本拓也", target: 15 },
  { name: "菅原貴博", target: 10 },
  { name: "江口裕人", target: 3 },
  { name: "藤井勇大", target: 2 }
];
const SPECIAL_QUEST_END_STR  = "2026/07/31";
const SPECIAL_QUEST_START_AT = "2026/07/27 00:00:00"; // この日時(JST)以降の受信のみカウント（過去分は含めない）

// フォーム記号を設定シートの内部データ（A列）から取得
// フォーム記号行がない場合はシート名サフィックスをフォールバックとして返す
function getFormCodeFromSheet(sheet) {
  const values = sheet.getDataRange().getValues();
  for (const row of values) {
    if (String(row[0]) === FORM_CODE_HEADER) {
      const v = String(row[1] || "").trim();
      if (v) return v;
    }
  }
  // フォールバック: シート名からフォーム記号を取得
  const name = sheet.getName();
  if (name.startsWith(CONFIG_PREFIX)) return name.replace(CONFIG_PREFIX, "");
  return null;
}

// フォーム記号でスプレッドシート内の設定シートを探す
function getConfigSheetByCode(ss, formCode) {
  for (const sheet of ss.getSheets()) {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) continue;
    if (getFormCodeFromSheet(sheet) === formCode) return sheet;
  }
  // フォールバック: シート名で直接検索
  return ss.getSheetByName(CONFIG_PREFIX + formCode) || null;
}

// ---- GET: フォーム設定を返す ----
function doGet(e) {
  try {
    // 緊急クエスト管理ルート（quest_admin）はクエスト終了(2026/07/31)に伴い撤去した。
    // 広告主成果管理シートの保守ルート（キー保護。月指定の再生成と件数確認のみ）。
    // キーは ScriptProperties から読む。未設定なら誰も通さない。
    const advAdminKey = advertiserAdminKey_();
    if (advAdminKey && e && e.parameter && e.parameter.adv_admin === advAdminKey) {
      return handleAdvertiserAdmin_(e.parameter.action || "preview", e.parameter.months || "",
                                    e.parameter.backup === "1");
    }
// 代理店専用リンク集。開くたびに呼ばれるので稼働状況がリアルタイムに反映される。
    if (e && e.parameter && e.parameter.agency_links) {
      return ContentService
        .createTextOutput(JSON.stringify(agencyLinksPayload_(e.parameter.agency_links)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss       = getOrCreateSpreadsheet();
    const formName = (e && e.parameter && e.parameter.form) ? e.parameter.form : getFirstFormCode(ss);
    const config   = readConfig(ss, formName);
    config.formName = formName;

    // 稼働していない案件は申請させない。停止中のアフィリンクへ顧客を送ると
    // 「このキャンペーンは終了しました」に着地して申請できずに終わるため。
    config.suspended = !isCaseActive_(formName);

    // 代理店リンク（?ag=<代理店コード>）で開かれた場合。
    // 代理店の紹介者はその担当者ひとりに決まるので、紹介者欄は出さず自動で入れる。
    const agParam = (e && e.parameter && e.parameter.ag) ? String(e.parameter.ag).trim() : "";
    if (agParam) {
      const agency = findAgencyByCode_(agParam);
      if (agency && agency.status === AGENCY_STATUS_ACTIVE) {
        config.agencyCode   = agency.code;
        config.agencyName   = agency.name;
        config.fixedReferrer = agency.person;
      } else {
        config.suspended    = true;
        config.agencyError  = "この代理店リンクは現在ご利用いただけません。";
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify(config))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- POST: フォーム回答を受信・保存 / LINE Webhook ----
function doPost(e) {
  let bridgeIdempotencyKey = '';
  try {
    const rawBody = e && e.postData ? e.postData.contents : '';
    let data = JSON.parse(rawBody);
    const bridge = _verifyCustomerLineBridgeRequest(e, rawBody, data);
    if (bridge.present) {
      if (!bridge.accepted) return _customerLineBridgeResponse(false, bridge.code);
      if (bridge.test) return _customerLineBridgeResponse(true, 'bridge_test');
      if (bridge.duplicate) return _customerLineBridgeResponse(true, 'duplicate');
      bridgeIdempotencyKey = bridge.idempotencyKey;
      handleLineWebhook({ events: [bridge.event] });
      _customerLineBridgeFinishEvent(bridgeIdempotencyKey, true);
      return _customerLineBridgeResponse(true, 'delivered');
    }

    // LINE Webhook イベント（eventsプロパティ存在で判定）
    if (data.events !== undefined) {
      if (!_customerLineDirectWebhookEnabled()) {
        return _customerLineBridgeResponse(false, 'direct_webhook_disabled');
      }
      return handleLineWebhook(data);
    }

    // 代理店の新規登録（代理店自身が公開フォームから申し込む）。
    // 公開経路なので合言葉を検証する（registerAgency_ の第2引数）。
    if (data.action === "registerAgency") {
      let out;
      try {
        out = registerAgency_(data, true);
      } catch (regErr) {
        out = { result: "error", message: String(regErr && regErr.message ? regErr.message : regErr) };
      }
      return ContentService
        .createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss       = getOrCreateSpreadsheet();
    const formName = data.formName || getFirstFormCode(ss);
    const config   = readConfig(ss, formName);

    const sheet = getConfigSheetByCode(ss, formName);
    if (!sheet) throw new Error("設定シート（フォーム記号: " + formName + "）が見つかりません。");

    // 停止中の案件は受け付けない（画面側でも止めているが、直接POSTされた場合の防御）
    if (!isCaseActive_(formName)) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: "error", message: "この案件は現在受付を停止しています。" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 代理店経由の申請。紹介者はその代理店の担当者に決まるので上書きする。
    let agencyName = "";
    const agencyCode = String(data.agencyCode || "").trim();
    if (agencyCode) {
      const agency = findAgencyByCode_(agencyCode);
      if (!agency || agency.status !== AGENCY_STATUS_ACTIVE) {
        return ContentService
          .createTextOutput(JSON.stringify({ result: "error", message: "この代理店リンクは現在ご利用いただけません。" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      agencyName    = agency.name;
      data.referrer = agency.person;
    }

    // G1にヘッダーがなければ初期化
    if (!sheet.getRange(1, ANSWER_START_COL).getValue()) {
      const headers = buildHeaders(config);
      const range   = sheet.getRange(1, ANSWER_START_COL, 1, headers.length);
      range.setValues([headers]);
      range.setFontWeight("bold");
      range.setBackground("#4f46e5");
      range.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    const screenshotUrl = data.screenshot
      ? saveScreenshot(data.screenshot, data.screenshotName, data)
      : "";

    const rowData = buildRow(data, config, screenshotUrl, formName, agencyName);
    const nextRow = findNextAnswerRow(sheet);
    sheet.getRange(nextRow, ANSWER_START_COL, 1, rowData.length).setValues([rowData]);

    // 代理店SSにも書き込み
    try {
      const code = getAgencyCode(sheet);
      const agencySS = getOrCreateAgencySpreadsheet(code);
      let agencySheet = getConfigSheetByCode(agencySS, formName);
      if (!agencySheet) {
        syncFormSheetToAgency(ss, agencySS, formName);
        agencySheet = getConfigSheetByCode(agencySS, formName);
      }
      if (agencySheet) {
        if (!agencySheet.getRange(1, ANSWER_START_COL).getValue()) {
          const headers = buildHeaders(config);
          const r = agencySheet.getRange(1, ANSWER_START_COL, 1, headers.length);
          r.setValues([headers]);
          r.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
          agencySheet.setFrozenRows(1);
        }
        const agencyNext = findNextAnswerRow(agencySheet);
        agencySheet.getRange(agencyNext, ANSWER_START_COL, 1, rowData.length).setValues([rowData]);
      }
    } catch (agencyErr) {
      Logger.log("代理店SS書き込みエラー: " + agencyErr);
    }

    // LINE グループ通知
    try {
      notifyLineGroup(buildLineMessage(config, rowData, formName));
    } catch (lineErr) {
      Logger.log("LINE通知エラー: " + lineErr);
    }

    // 申請状況一覧へ即時反映する。日次生成だけだとリンク集の件数が翌朝まで増えない。
    try {
      appendApplicationToStatusSheet_(
        config.formDisplayName || config.formTitle || formName,
        data["name"] || "", data["referrer"] || "", agencyName, screenshotUrl);
    } catch (statusErr) {
      Logger.log("申請状況一覧への即時反映エラー: " + statusErr);
    }

    // 代理店経由なら、その代理店へも新規申請を知らせる
    try {
      if (agencyName) {
        notifyAgencyOfApplication_(agencyName,
          config.formDisplayName || config.formTitle || formName,
          data["name"] || "");
      }
    } catch (agencyMailErr) {
      Logger.log("代理店への新規申請通知エラー: " + agencyMailErr);
    }

    // 顧客管理シートに自動追加（紹介者名があるとき）。
    // 代理店経由の申請は上で referrer を代理店の担当者名に置き換えてあるので、
    // 自社・代理店を問わず同じ経路で顧客管理へ入る＝顧客の一元管理になる。
    try {
      const referrer = data["referrer"] || "";
      if (referrer && referrer !== "__other__") {
        upsertCustomerRow(referrer, data["name"] || "", formName);
      }
    } catch (custErr) {
      Logger.log("顧客管理シート更新エラー: " + custErr);
    }

    // 広告主成果管理シートに書き込み
    try {
      writeToAdvertiserSheet(rowData[1], config.formDisplayName || config.formTitle || formName, data["name"] || "", data["referrer"] || "", screenshotUrl, false);
    } catch (advErr) {
      Logger.log("広告主シート書き込みエラー: " + advErr);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    if (bridgeIdempotencyKey) {
      try {
        _customerLineBridgeFinishEvent(bridgeIdempotencyKey, false);
      } catch (finishError) {
        console.error('bridge rollback error:', finishError);
      }
    }
    Logger.log(err);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- G列の次の空き行を返す（1行目はヘッダー）----
function findNextAnswerRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 2;
  const values = sheet.getRange(1, ANSWER_START_COL, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") return i + 2;
  }
  return 2;
}

// ---- スプレッドシート取得 or 作成 ----
function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId    = props.getProperty("SPREADSHEET_ID");

  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch (e) {}
  }

  const ss = SpreadsheetApp.create("アフィリエイト管理");
  props.setProperty("SPREADSHEET_ID", ss.getId());

  const configSheet = ss.getSheets()[0];
  configSheet.setName(CONFIG_PREFIX + "申請フォーム");
  initConfigSheet(configSheet, "申請フォーム", "", "form1");

  return ss;
}

// ---- 設定シートが存在する最初のフォーム記号を返す ----
function getFirstFormCode(ss) {
  for (const sheet of ss.getSheets()) {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) continue;
    const code = getFormCodeFromSheet(sheet);
    if (code) return code;
  }
  throw new Error("設定シートが見つかりません。スプレッドシートに「" + CONFIG_PREFIX + "フォーム名」シートを作成してください。");
}

// ---- 設定シートの初期データ ----
function initConfigSheet(sheet, title, displayName, formCode) {
  title       = title || "申請フォーム";
  displayName = displayName || "";
  formCode    = formCode || sheet.getName().replace(CONFIG_PREFIX, "");
  const formUrl  = FORM_BASE_URL + "?form=" + encodeURIComponent(formCode);

  const data = [
    ["フォームタイトル",    title],
    ["フォーム説明文",      "以下の手順に従って入力・作業を行ってください。"],
    ["アフィリエイトURL",   "https://ここにアフィリエイトリンクを入力"],
    ["ボタンテキスト",      "アフィリエイトリンクを開く（必ずここから！）"],
    ["フォームURL（自動）", formUrl],
    [AGENCY_KEY,            ""],
    [FORM_NAME_KEY,         displayName],
    [FORM_CODE_HEADER,      formCode],
    ["＝＝ フォーム項目（行を追加・削除で変更可） ＝＝", ""],
    ["フィールドID", "ラベル", "タイプ(text/textarea/select)", "必須(TRUE/FALSE)", "プレースホルダー", "選択肢（select時カンマ区切り）"],
    ["name",     "お名前",   "text", "TRUE", "例：山田太郎", ""],
    ["referrer", "紹介者名", "text", "TRUE", "",              ""],
  ];
  sheet.getRange(1, 1, data.length, 6).setValues(
    data.map(row => { while (row.length < 6) row.push(""); return row; })
  );

  // フォームURL行（5行目）
  const urlRange = sheet.getRange(5, 1, 1, 2);
  urlRange.setBackground("#e8f5e9").setFontColor("#1b5e20").setFontWeight("bold");
  sheet.getRange(5, 2).setFontStyle("italic");

  // 代理店コード行（6行目）
  const agencyRange = sheet.getRange(6, 1, 1, 2);
  agencyRange.setBackground("#e0f2fe").setFontColor("#075985").setFontWeight("bold");

  // フォーム名行（7行目・表示用）
  const displayRange = sheet.getRange(7, 1, 1, 2);
  displayRange.setBackground("#fff3e0").setFontColor("#e65100").setFontWeight("bold");

  // フォーム記号行（8行目）
  const codeRange = sheet.getRange(8, 1, 1, 2);
  codeRange.setBackground("#e8eaf6").setFontColor("#1a237e").setFontWeight("bold");

  // フォーム項目ヘッダー行（10行目）
  const headerRange = sheet.getRange(10, 1, 1, 5);
  headerRange.setFontWeight("bold").setBackground("#e8eaf6").setFontColor("#1a237e");

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 240);
}

// ---- 設定シートを読み込む ----
function readConfig(ss, formName) {
  const sheet = getConfigSheetByCode(ss, formName);
  if (!sheet) {
    throw new Error("設定シート（フォーム記号: " + formName + "）が見つかりません。");
  }

  // フォームURLを最新のシート名で自動更新
  const formUrl   = FORM_BASE_URL + "?form=" + encodeURIComponent(formName);
  const values    = sheet.getDataRange().getValues();
  let   urlRowIdx = -1;
  values.forEach((row, i) => {
    if (String(row[0]).includes("フォームURL")) urlRowIdx = i;
  });
  if (urlRowIdx >= 0) {
    sheet.getRange(urlRowIdx + 1, 2).setValue(formUrl);
  } else {
    sheet.insertRowBefore(5);
    const urlRange = sheet.getRange(5, 1, 1, 2);
    urlRange.setValues([["フォームURL（自動）", formUrl]]);
    urlRange.setBackground("#e8f5e9").setFontColor("#1b5e20").setFontWeight("bold");
    sheet.getRange(5, 2).setFontStyle("italic");
  }

  const config = {
    formTitle: "", formDisplayName: "", formDescription: "", affiliateUrl: "", affiliateButtonText: "", fields: []
  };
  const keyMap = {
    "フォームタイトル":  "formTitle",
    [FORM_NAME_KEY]:     "formDisplayName",
    "フォーム説明文":    "formDescription",
    "アフィリエイトURL": "affiliateUrl",
    "ボタンテキスト":    "affiliateButtonText"
  };

  const allValues      = sheet.getDataRange().getValues();
  let   fieldsStartIdx = -1;
  for (let i = 0; i < allValues.length; i++) {
    if (String(allValues[i][0]) === "フィールドID") { fieldsStartIdx = i + 1; break; }
  }

  const configRows = fieldsStartIdx >= 0 ? allValues.slice(0, fieldsStartIdx - 1) : allValues.slice(0, 7);
  configRows.forEach(row => {
    if (keyMap[row[0]]) config[keyMap[row[0]]] = String(row[1] || "");
  });

  if (fieldsStartIdx >= 0) {
    allValues.slice(fieldsStartIdx).forEach(row => {
      const id = String(row[0] || "").trim();
      if (!id || id.startsWith("＝")) return;
      config.fields.push({
        id:          id,
        label:       String(row[1] || ""),
        type:        String(row[2] || "text"),
        required:    String(row[3]).toUpperCase() === "TRUE",
        placeholder: String(row[4] || ""),
        options:     String(row[5] || "").split(",").map(s => s.trim()).filter(Boolean)
      });
    });
  }
  return config;
}

// ---- ヘッダー行を組み立て ----
// 「代理店」は末尾に足す。末尾なので既存列の位置が動かず、既存データと
// 既存処理（承認列の読み取り等）に影響しない。自社経由の申請では空になる。
function buildHeaders(config) {
  const fieldLabels = config.fields.map(f => f.label);
  return [FORM_CODE_HEADER, "受信日時", "クリック日時", "送信日時", ...fieldLabels, "スクショURL", "承認", AGENCY_COLUMN_LABEL];
}

// ---- データ行を組み立て ----
function buildRow(data, config, screenshotUrl, formName, agencyName) {
  const receivedAt  = formatJST(new Date());
  const clickAt     = data.clickTime  ? formatJST(new Date(data.clickTime))  : "";
  const submitAt    = data.submitTime ? formatJST(new Date(data.submitTime)) : "";
  const fieldValues = config.fields.map(f => data[f.id] || "");
  return [formName, receivedAt, clickAt, submitAt, ...fieldValues, screenshotUrl, "", agencyName || ""];
}

// ---- スクショ保存フォルダをIDで取得（IDが無効なら名前検索してID保存） ----
function getSaveFolder() {
  const props   = PropertiesService.getScriptProperties();
  const savedId = props.getProperty("SCREENSHOT_FOLDER_ID");
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {
      props.deleteProperty("SCREENSHOT_FOLDER_ID");
    }
  }
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
  const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);
  props.setProperty("SCREENSHOT_FOLDER_ID", folder.getId());
  return folder;
}

// ---- スクショ保存フォルダを手動で再登録（メニューから実行） ----
function resetScreenshotFolder() {
  const props   = PropertiesService.getScriptProperties();
  props.deleteProperty("SCREENSHOT_FOLDER_ID");
  const folder = getSaveFolder();
  SpreadsheetApp.getUi().alert("スクショフォルダを再登録しました。\nフォルダ名: " + folder.getName() + "\nID: " + folder.getId());
}

// ---- スクショ保存 ----
function saveScreenshot(base64Data, fileName, data) {
  try {
    const folder  = getSaveFolder();
    const base64  = base64Data.split(",")[1];
    const mime    = base64Data.split(";")[0].split(":")[1];
    const blob    = Utilities.newBlob(Utilities.base64Decode(base64), mime, fileName);
    const name    = data.name || "不明";
    const dt      = data.clickTime ? formatJSTforFilename(new Date(data.clickTime)) : formatJSTforFilename(new Date());
    blob.setName(`${dt}_${name}_${fileName || "screenshot"}`);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      Logger.log("setSharing 失敗（無視して続行）: " + e);
    }
    return file.getUrl();
  } catch (err) {
    return "保存エラー: " + err.toString();
  }
}

// ---- 日時フォーマット ----
function formatJST(date) {
  const jst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  const p   = n => String(n).padStart(2, "0");
  return `${jst.getUTCFullYear()}/${p(jst.getUTCMonth()+1)}/${p(jst.getUTCDate())} ${p(jst.getUTCHours())}:${p(jst.getUTCMinutes())}:${p(jst.getUTCSeconds())}`;
}
function formatJSTforFilename(date) {
  const jst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  const p   = n => String(n).padStart(2, "0");
  return `${jst.getUTCFullYear()}${p(jst.getUTCMonth()+1)}${p(jst.getUTCDate())}_${p(jst.getUTCHours())}${p(jst.getUTCMinutes())}`;
}

// ---- スプレッドシートを開いたときに各種更新 ----
function onOpen() {
  updateAllFormUrls();
  initAllAnswerHeaders();
  migrateAddFormNameRow();
  updateManagementSheet();
  ensureDailyReportTrigger();
  ensureCampaignReportTrigger();
  ensureQuest150Trigger();
  ensureEmergencyQuestTriggers();
  applyReferrerSelectToJishaSheets();
  try { syncCaseMaster(); } catch (e) { Logger.log("案件マスタ同期エラー: " + e); }
  try { ensureAgencyNotifyTrigger(); } catch (e) { Logger.log("代理店通知トリガー登録エラー: " + e); }
  try { ensureAppStatusTrigger(); } catch (e) { Logger.log("申請状況トリガー登録エラー: " + e); }
  SpreadsheetApp.getUi().createMenu("フォーム管理")
    .addItem("新規フォーム作成",       "showCreateFormDialog")
    .addItem("管理シートを更新",       "updateManagementSheet")
    .addSeparator()
    .addItem("初回セットアップ（1回だけ）",     "setupCaseAgencyFeatureFromMenu")
    .addItem("案件マスタを同期＋稼働を反映",   "syncCaseMasterAndApply")
    .addItem("稼働状況をシート表示へ反映",     "applyCaseVisibilityFromMenu")
    .addItem("代理店を登録（リンク集をメール送信）", "showAgencyRegisterPrompt")
    .addItem("代理店を削除",                   "showAgencyDeletePrompt")
    .addItem("全代理店へリンク集を送り直す",   "resendAllAgencyLinks")
    .addItem("代理店別の取扱案件を同期",       "syncAgencyCaseMatrixFromMenu")
    .addItem("申請状況一覧を作り直す（自社＋全代理店）", "buildApplicationStatusSheetFromMenu")
    .addItem("代理店別サマリーを作り直す（件数・承認率）", "buildAgencySummarySheetFromMenu")
    .addSeparator()
    .addItem("ASP獲得ログと突合する",         "reconcileAspLogFromMenu")
    .addItem("ASP突合の修正を反映",           "applyAspReconciliationFromMenu")
    .addItem("請求漏れ10件を承認へ直す（1回だけ）", "aspFix20260820FromMenu")
    .addItem("稼働の変更を代理店へ通知",       "notifyAgencyCaseChangesFromMenu")
    .addItem("稼働変更の日次通知を有効化",     "ensureAgencyNotifyTriggerFromMenu")
    .addItem("回答シートに「代理店」列を追加", "ensureAgencyColumnFromMenu")
    .addItem("代理店割り当て更新",     "rebuildAllAgencySpreadsheets")
    .addItem("旧共有SSをゴミ箱へ",     "deleteAllOldSharingSpreadsheets")
    .addSeparator()
    .addItem("日次レポート（テスト送信）",           "dailyReport")
    .addItem("顧客LINE登録集計を確認",             "checkCustomerLineStatsFromMenu")
    .addItem("30件クエスト進捗（テスト送信）",       "campaignReport")
    .addItem("150件クエスト進捗（テスト送信）",     "quest150Report")
    .addItem("緊急クエスト進捗（テスト送信）",       "emergencyQuestReportTest")
    .addItem("特別緊急クエスト進捗（テスト送信）",   "specialQuestReportTest")
    .addSeparator()
    .addItem("回答ヘッダーを最新フィールドに同期", "fixAnswerHeaders")
    .addItem("シート名をフォーム名に変換（移行）", "migrateSheetNamesToDisplayName")
    .addItem("フォーム記号を修復",               "repairFormCodeRows")
    .addSeparator()
    .addItem("スクショフォルダを再登録",   "resetScreenshotFolder")
    .addSeparator()
    .addItem("営業担当を同期（選択肢＋担当タブ＋データ再生成）", "syncSalesRoster")
    .addItem("顧客管理シートを作成",           "createCustomerManagementSheet")
    .addItem("顧客管理の案件列を同期",         "syncCustomerManagementCases")
    .addItem("既存データを顧客管理シートへインポート", "importExistingToCustomerSheet")
    .addItem("既存データを広告主シートへインポート",   "importExistingToAdvertiserSheet")
    .addItem("広告主シートを月指定で再生成",           "importAdvertiserMonths")
    .addSeparator()
    .addItem("担当別ステータス表の案件列を同期（SS2）", "syncRepStatusCaseColumns")
    .addItem("担当別ステータス表を再生成（SS2）",       "buildSalesRepStatusSheets")
    .addItem("総合_担当タブを再生成（SS1）",           "buildIntegratedRepSheets")
    .addToUi();
}

// ---- 自社シートの紹介者フィールドをselectに切り替え＋選択肢を名簿へ同期 ----
// onOpen から呼ばれる。名簿(JISHA_REFERRER_OPTIONS)が前回適用時から変わっていれば
// 自動で再適用する（メンバー増減が次回シート起動時に反映される）。force=true で常に適用。
function applyReferrerSelectToJishaSheets(force) {
  const props = PropertiesService.getScriptProperties();
  const applied = props.getProperty("REFERRER_OPTIONS_APPLIED");
  // 旧フラグ("REFERRER_SELECT_APPLIED"=="1")からの移行：値未記録なら「未同期」とみなして一度適用する
  if (!force && applied === JISHA_REFERRER_OPTIONS) return { updated: 0, skipped: true };

  const ss = getOrCreateSpreadsheet();
  let updated = 0;
  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const values = sheet.getDataRange().getValues();

    // 代理店コードを直接読む（行挿入の副作用を避けるため）
    let agencyCode = AGENCY_DEFAULT;
    for (const row of values) {
      if (String(row[0]) === AGENCY_KEY) {
        const v = String(row[1] || "").trim();
        agencyCode = v || AGENCY_DEFAULT;
        break;
      }
    }
    if (agencyCode !== AGENCY_DEFAULT) return; // 代理店はスキップ

    let fieldsStartIdx = -1;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === "フィールドID") { fieldsStartIdx = i + 1; break; }
    }
    if (fieldsStartIdx < 0) return;

    for (let i = fieldsStartIdx; i < values.length; i++) {
      if (String(values[i][0]).trim() === "referrer") {
        sheet.getRange(i + 1, 3).setValue("select");
        sheet.getRange(i + 1, 6).setValue(JISHA_REFERRER_OPTIONS);
        updated++;
        break;
      }
    }
  });

  props.setProperty("REFERRER_OPTIONS_APPLIED", JISHA_REFERRER_OPTIONS);
  props.setProperty("REFERRER_SELECT_APPLIED", "1"); // 後方互換
  Logger.log("自社シートの紹介者選択肢を同期しました（更新 " + updated + " シート）: " + JISHA_REFERRER_OPTIONS);
  return { updated: updated, skipped: false };
}

// ---- 名簿(JISHA_REFERRER_OPTIONS)を一括同期（メニュー「営業担当を同期」）----
// (1)全自社フォームの紹介者選択肢を再適用 (2)顧客管理SSに担当タブを追加
// (3)SS2(担当別ステータス表)に担当タブを追加 (4)SS1(統合顧客管理)に「総合_<担当>」タブを追加
// (5)SS2/SS1 の担当別データを再生成し、(3)(4)で作った空タブに実データを流し込む。
// (1)〜(4) は非破壊（既存タブ・行・データは保持）。(5) は担当別タブ＝生成物のみ作り直す。
// 営業メンバーを増減したら、この関数を1回実行するだけで全フォーム・全シートが揃う。
// ※ (5) を省くと「タブはあるが中身が空」の担当が残るので、ここで必ず続けて実行する。
function syncSalesRoster() {
  const referrer = applyReferrerSelectToJishaSheets(true);
  const cmt = ensureCustomerMgmtTabs_();
  const rep = ensureRepStatusTabs_();
  const cols = ensureRepStatusCaseColumns_(); // 案件列の横幅を全担当タブで揃えてから書き込む
  const integ = ensureIntegratedRepTabs_();

  // タブを作っただけではデータが入らないので、続けて担当別データを再生成する。
  // 片方が失敗しても (1)〜(4) の結果と他方の再生成は失わないよう、個別に捕捉して報告に載せる。
  let repBuild = null, repBuildErr = "";
  try { repBuild = buildSalesRepStatusSheets(); }
  catch (e) { repBuildErr = String(e); Logger.log("buildSalesRepStatusSheets 失敗: " + e); }
  let integBuild = null, integBuildErr = "";
  try { integBuild = buildIntegratedRepSheets(); }
  catch (e) { integBuildErr = String(e); Logger.log("buildIntegratedRepSheets 失敗: " + e); }

  const fmt = function (r) { return (r.added && r.added.length) ? r.added.join(", ") : "なし"; };
  const fmtCols = function (r) {
    const ks = Object.keys(r.added || {});
    return ks.length ? ks.map(function (k) { return k + "(" + r.added[k].join("/") + ")"; }).join(", ") : "なし";
  };
  const rows = function (b, err) {
    if (err) return "失敗（" + err + "）";
    if (!b || !b.perRepRows) return "結果なし";
    const keys = Object.keys(b.perRepRows);
    if (!keys.length) return "0件";
    return keys.map(function (k) { return k + "=" + b.perRepRows[k]; }).join(", ");
  };
  const msg = "営業担当ロスターを同期しました。\n\n"
    + "名簿: " + JISHA_REFERRER_OPTIONS + "\n\n"
    + "紹介者選択肢を更新: " + referrer.updated + " フォーム\n"
    + "顧客管理SS 追加タブ: " + fmt(cmt) + (cmt.error ? "（" + cmt.error + "）" : "") + "\n"
    + "SS2(担当別) 追加タブ: " + fmt(rep) + (rep.error ? "（" + rep.error + "）" : "") + "\n"
    + "SS2 追加した案件列: " + fmtCols(cols) + (cols.error ? "（" + cols.error + "）" : "") + "\n"
    + "SS1(総合_) 追加タブ: " + fmt(integ) + (integ.error ? "（" + integ.error + "）" : "") + "\n\n"
    + "SS2 担当別データ再生成: " + rows(repBuild, repBuildErr) + "\n"
    + "SS1 総合_データ再生成: " + rows(integBuild, integBuildErr);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return {
    referrer: referrer, customerMgmt: cmt, repStatus: rep,
    repStatusCaseColumns: cols, integrated: integ,
    repStatusBuild: repBuild, repStatusBuildError: repBuildErr,
    integratedBuild: integBuild, integratedBuildError: integBuildErr
  };
}

// ---- 顧客管理SSに名簿の担当タブが揃っているか確認し、無ければ追加（非破壊）----
// 見本タブ（先頭シート）のヘッダー・書式を複製して作る。案件列の入力規則も引き継ぐ。
function ensureCustomerMgmtTabs_() {
  const css = getCustomerManagementSS();
  if (!css) return { added: [], error: "顧客管理SSが見つかりません（先に『顧客管理シートを作成』を実行）" };
  const roster   = JISHA_REFERRER_OPTIONS.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
  const existing = css.getSheets().map(function (s) { return normalizeName(s.getName()); });
  const template = css.getSheets()[0];
  const lastCol  = template.getLastColumn();
  const headerVals = template.getRange(1, 1, 1, lastCol).getValues();
  const STATUS_OPTIONS = ["申請中", "申請済", "完了", "不参加"];
  const BASE_COLS = 3; // 顧客名/電話番号/顧客ID
  const added = [];
  roster.forEach(function (name) {
    if (existing.indexOf(normalizeName(name)) >= 0) return;
    const sheet = css.insertSheet(name);
    sheet.getRange(1, 1, 1, lastCol).setValues(headerVals);
    template.getRange(1, 1, 1, lastCol).copyTo(sheet.getRange(1, 1, 1, lastCol), { formatOnly: true });
    sheet.setFrozenRows(1);
    for (let c = 1; c <= lastCol; c++) sheet.setColumnWidth(c, template.getColumnWidth(c));
    if (lastCol > BASE_COLS) {
      const body = sheet.getRange(2, BASE_COLS + 1, 200, lastCol - BASE_COLS);
      body.setBackground("#ede9fe").setHorizontalAlignment("center");
      body.setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS).setAllowInvalid(true).build()
      );
    }
    added.push(name);
  });
  Logger.log("ensureCustomerMgmtTabs_: 追加=" + JSON.stringify(added));
  return { added: added };
}

// ---- SS2(担当別ステータス表)に名簿の担当タブが揃っているか確認し、無ければ追加（非破壊）----
// 既存の担当タブのヘッダー・書式を見本に複製。データは buildSalesRepStatusSheets 実行時に入る。
function ensureRepStatusTabs_() {
  let outSS;
  try { outSS = SpreadsheetApp.openById(REP_STATUS_SS2_ID); }
  catch (e) { return { added: [], error: "SS2を開けません: " + e }; }
  const roster = JISHA_REFERRER_OPTIONS.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
  // 見本は「案件列が最も揃っているタブ」を選ぶ。先頭一致で選ぶと、案件列が増える前に作られた
  // 古い(狭い)タブを複製してしまい、新メンバーだけ案件列が欠けた状態で作られる。
  let template = null;
  roster.forEach(function (name) {
    const t = outSS.getSheetByName(name);
    if (t && (!template || t.getLastColumn() > template.getLastColumn())) template = t;
  });
  const added = [];
  roster.forEach(function (name) {
    if (outSS.getSheetByName(name)) return;
    const sheet = outSS.insertSheet(name);
    if (template) {
      const lastCol = template.getLastColumn();
      sheet.getRange(1, 1, 1, lastCol).setValues(template.getRange(1, 1, 1, lastCol).getValues());
      template.getRange(1, 1, 1, lastCol).copyTo(sheet.getRange(1, 1, 1, lastCol), { formatOnly: true });
      sheet.setFrozenRows(1);
      for (let c = 1; c <= lastCol; c++) sheet.setColumnWidth(c, template.getColumnWidth(c));
    }
    added.push(name);
  });
  Logger.log("ensureRepStatusTabs_: 追加=" + JSON.stringify(added));
  return { added: added };
}

// ---- SS2(担当別ステータス表)の各担当タブに自社フォームの案件列が揃っているか確認し、無ければ追加（非破壊）----
// buildSalesRepStatusSheets は「そのタブ自身のヘッダー」を見て案件列を引くので、列が無い案件の実績は
// 黙って捨てられる（unmatchedCase に計上されるだけ）。新フォーム追加時・新メンバー追加時の
// 取りこぼしを防ぐため、名簿の全タブに不足案件列を追記して横幅を揃える。
function ensureRepStatusCaseColumns_() {
  let outSS;
  try { outSS = SpreadsheetApp.openById(REP_STATUS_SS2_ID); }
  catch (e) { return { added: {}, error: "SS2を開けません: " + e }; }

  const mainSS = getOrCreateSpreadsheet();
  if (mainSS.getId() !== REP_STATUS_MAIN_ID) {
    return { added: {}, error: "メインSSのID不一致 実ID=" + mainSS.getId() };
  }
  // 自社(house)フォームの案件表示名を、buildSalesRepStatusSheets と同じ条件で集める
  const caseNames = [];
  mainSS.getSheets().forEach(function (sheet) {
    if (sheet.getName().indexOf(CONFIG_PREFIX) !== 0) return;
    const vals = sheet.getDataRange().getValues();
    let agencyCode = AGENCY_DEFAULT;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === AGENCY_KEY) { const c = String(vals[i][1] || "").trim(); if (c) agencyCode = c; break; }
    }
    if (agencyCode !== AGENCY_DEFAULT) return;
    const cn = getFormDisplayName(sheet, getFormCodeFromSheet(sheet) || "");
    if (cn && caseNames.indexOf(cn) < 0) caseNames.push(cn);
  });
  if (!caseNames.length) return { added: {}, error: "自社フォームの案件名を取得できませんでした" };

  const roster = JISHA_REFERRER_OPTIONS.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
  const BASE_COLS = 3; // 顧客名/電話番号/顧客ID
  const added = {};
  roster.forEach(function (name) {
    const tab = outSS.getSheetByName(name);
    if (!tab) return;
    const lastCol = tab.getLastColumn();
    const header = tab.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    const missing = caseNames.filter(function (cn) { return header.indexOf(cn) < 0; });
    if (!missing.length) return;
    if (tab.getMaxColumns() < lastCol + missing.length) {
      tab.insertColumnsAfter(lastCol, lastCol + missing.length - tab.getMaxColumns());
    }
    tab.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    // 既存の案件列(4列目)のヘッダー書式を見本に揃える
    if (lastCol > BASE_COLS) {
      tab.getRange(1, BASE_COLS + 1, 1, 1)
         .copyTo(tab.getRange(1, lastCol + 1, 1, missing.length), { formatOnly: true });
    }
    for (let i = 0; i < missing.length; i++) tab.setColumnWidth(lastCol + 1 + i, 130);
    added[name] = missing;
  });
  Logger.log("ensureRepStatusCaseColumns_: 追加列=" + JSON.stringify(added));
  return { added: added };
}

// メニューから単体実行する用（新しいフォームを追加したあとに走らせる）
function syncRepStatusCaseColumns() {
  const r = ensureRepStatusCaseColumns_();
  const names = Object.keys(r.added || {});
  const msg = r.error
    ? "SS2 案件列の同期に失敗: " + r.error
    : (names.length
        ? "SS2 案件列を追加しました:\n" + names.map(function (n) { return "・" + n + ": " + r.added[n].join(", "); }).join("\n")
        : "SS2 の案件列は全担当タブで揃っています。");
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return r;
}

// ---- SS1(統合顧客管理)に名簿の「総合_<担当>」タブが揃っているか確認し、無ければ追加（非破壊）----
// 既存の 総合_ タブのヘッダー・書式を見本に空タブを複製。データは buildIntegratedRepSheets 実行時に入る。
function ensureIntegratedRepTabs_() {
  let ss1;
  try { ss1 = SpreadsheetApp.openById(REP_STATUS_SS1_ID); }
  catch (e) { return { added: [], error: "SS1を開けません: " + e }; }
  const roster = JISHA_REFERRER_OPTIONS.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
  const PREFIX = "総合_";
  let template = null;
  const sheets = ss1.getSheets();
  for (const s of sheets) { if (s.getName().indexOf(PREFIX) === 0) { template = s; break; } }
  const added = [];
  roster.forEach(function (name) {
    const tabName = PREFIX + name;
    if (ss1.getSheetByName(tabName)) return;
    const sheet = ss1.insertSheet(tabName);
    if (template) {
      const lastCol = template.getLastColumn();
      sheet.getRange(1, 1, 1, lastCol).setValues(template.getRange(1, 1, 1, lastCol).getValues());
      template.getRange(1, 1, 1, lastCol).copyTo(sheet.getRange(1, 1, 1, lastCol), { formatOnly: true });
      sheet.setFrozenRows(1);
      for (let c = 1; c <= lastCol; c++) sheet.setColumnWidth(c, template.getColumnWidth(c));
    }
    added.push(tabName);
  });
  Logger.log("ensureIntegratedRepTabs_: 追加=" + JSON.stringify(added));
  return { added: added };
}

// ---- 全設定シートの回答ヘッダーをG1に初期化（未設定のシートのみ）----
function initAllAnswerHeaders() {
  const ss = getOrCreateSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const g1 = sheet.getRange(1, ANSWER_START_COL).getValue();

    // 旧ヘッダー「フォーム名」→「フォーム記号」へ移行
    if (g1 === "フォーム名") {
      sheet.getRange(1, ANSWER_START_COL).setValue(FORM_CODE_HEADER);
      return;
    }
    if (g1) return;

    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    const config   = readConfig(ss, formCode);
    const headers  = buildHeaders(config);
    const range    = sheet.getRange(1, ANSWER_START_COL, 1, headers.length);
    range.setValues([headers]);
    range.setFontWeight("bold");
    range.setBackground("#4f46e5");
    range.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  });
}

// ---- 回答ヘッダーを現在のフォーム設定に同期（新フィールド追加後に実行）----
// 旧データ行: スクショURLを旧位置→新位置に移動して空フィールドを挿入
// 新データ行: 既に正しいフォーマット（URL位置で判定）なのでそのまま保持
function fixAnswerHeaders() {
  const ss = getOrCreateSpreadsheet();
  const results = [];
  const BASE_FIXED = 4; // フォーム記号 / 受信日時 / クリック日時 / 送信日時

  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < ANSWER_START_COL) return;
    const colCount = lastCol - ANSWER_START_COL + 1;

    const actualHeaders = sheet.getRange(1, ANSWER_START_COL, 1, colCount).getValues()[0].map(String);
    const ssUrlOldIdx   = actualHeaders.indexOf("スクショURL"); // 0-based from ANSWER_START_COL
    if (ssUrlOldIdx < 0) return;

    const oldFieldCount = ssUrlOldIdx - BASE_FIXED;

    let config;
    try { config = readConfig(ss, formCode); } catch (e) { return; }

    const newFieldLabels = config.fields.map(f => f.label);
    const oldFieldLabels = actualHeaders.slice(BASE_FIXED, ssUrlOldIdx);

    if (JSON.stringify(newFieldLabels) === JSON.stringify(oldFieldLabels)) return; // 変更なし

    const addedCount   = newFieldLabels.length - oldFieldCount;
    if (addedCount <= 0) return; // フィールドが増えていない場合はスキップ

    const ssUrlNewIdx = BASE_FIXED + newFieldLabels.length; // 新フォーマットでのスクショURL位置（0-based）

    // --- データ行の修正 ---
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, colCount).getValues();

      const fixedData = data.map(row => {
        // 新フォーマット判定: 新しいスクショURL位置に URL があるか
        const valAtNewPos = String(row[ssUrlNewIdx] || "");
        if (valAtNewPos.startsWith("http")) return row; // 新フォーマット: そのまま

        // 旧フォーマット: スクショURLを旧位置から取り出し、空フィールドを挿入して再組み立て
        const fixedHead    = row.slice(0, BASE_FIXED);
        const oldFields    = row.slice(BASE_FIXED, ssUrlOldIdx);
        const screenshotUrl = String(row[ssUrlOldIdx] || "");
        const approval      = String(row[ssUrlOldIdx + 1] || "");
        const padding       = Array(addedCount).fill("");
        return [...fixedHead, ...oldFields, ...padding, screenshotUrl, approval];
      });

      const newColCount = BASE_FIXED + newFieldLabels.length + 2;
      const maxCols     = Math.max(colCount, newColCount);
      const normalized  = fixedData.map(row => {
        const r = [...row];
        while (r.length < maxCols) r.push("");
        return r.slice(0, maxCols);
      });
      sheet.getRange(2, ANSWER_START_COL, lastRow - 1, maxCols).setValues(normalized);
    }

    // --- ヘッダー更新 ---
    const expectedHeaders = buildHeaders(config);
    const headerRange     = sheet.getRange(1, ANSWER_START_COL, 1, expectedHeaders.length);
    headerRange.setValues([expectedHeaders]);
    headerRange.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");

    results.push(sheet.getName() + ": 追加=" + newFieldLabels.slice(oldFieldCount).join("・"));
    Logger.log("fixAnswerHeaders [" + sheet.getName() + "]: " + newFieldLabels.slice(oldFieldCount).join(", "));
  });

  if (results.length === 0) {
    SpreadsheetApp.getUi().alert("修正が必要なシートはありませんでした。");
  } else {
    SpreadsheetApp.getUi().alert("ヘッダー修正完了！\n\n" + results.join("\n"));
  }
}

// ---- 設定シートからフォーム表示名を取得（無ければformCodeを返す）----
function getFormDisplayName(sheet, formCode) {
  const values = sheet.getDataRange().getValues();
  for (const row of values) {
    if (String(row[0]) === FORM_NAME_KEY) {
      const v = String(row[1] || "").trim();
      if (v) return v;
      break;
    }
  }
  return formCode;
}

// ---- 既存設定シートに FORM_NAME_KEY 行を補完（行シフトなし） ----
function migrateAddFormNameRow() {
  const ss = getOrCreateSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const values = sheet.getDataRange().getValues();

    // 既に行があればスキップ
    for (const row of values) {
      if (String(row[0]) === FORM_NAME_KEY) return;
    }

    // 代理店コード行の直下が空白なら、その行に書き込む（行シフトなし）
    let agencyRowIdx = -1;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === AGENCY_KEY) { agencyRowIdx = i; break; }
    }
    if (agencyRowIdx < 0) return;

    const candidateIdx = agencyRowIdx + 1;
    if (candidateIdx >= values.length) return;
    const candidate = values[candidateIdx];
    if (String(candidate[0]) !== "" || String(candidate[1]) !== "") return;

    const range = sheet.getRange(candidateIdx + 1, 1, 1, 2);
    range.setValues([[FORM_NAME_KEY, ""]]);
    range.setBackground("#fff3e0").setFontColor("#e65100").setFontWeight("bold");
  });
}

// ---- 全設定シートのフォームURLを更新 ----
function updateAllFormUrls() {
  const ss     = getOrCreateSpreadsheet();
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    const formUrl  = FORM_BASE_URL + "?form=" + encodeURIComponent(formCode);
    const values   = sheet.getDataRange().getValues();
    let   urlRowIdx = -1;
    values.forEach((row, i) => {
      if (String(row[0]).includes("フォームURL")) urlRowIdx = i;
    });
    if (urlRowIdx >= 0) {
      sheet.getRange(urlRowIdx + 1, 2).setValue(formUrl);
    } else {
      sheet.insertRowBefore(5);
      const urlRange = sheet.getRange(5, 1, 1, 2);
      urlRange.setValues([["フォームURL（自動）", formUrl]]);
      urlRange.setBackground("#e8f5e9").setFontColor("#1b5e20").setFontWeight("bold");
      sheet.getRange(5, 2).setFontStyle("italic");
    }
  });
}

// ---- onOpenトリガーをインストール（初回1回だけ実行） ----
function installTrigger() {
  const ss = getOrCreateSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onOpen") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onOpen").forSpreadsheet(ss).onOpen().create();
  installMainEditTrigger();
  ensureDailyReportTrigger();
}

// ---- スプレッドシートURLの確認用 ----
function getSpreadsheetUrl() {
  Logger.log(getOrCreateSpreadsheet().getUrl());
}

// ---- Google Sheets URLからスプレッドシートIDを抽出 ----
function extractSsIdFromUrl(url) {
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

// =============================================
// 代理店スプレッドシート管理
// =============================================

// ---- 設定シートから代理店コードを取得（行が無ければ自動で挿入）----
function getAgencyCode(sheet) {
  const values = sheet.getDataRange().getValues();
  let agencyRowIdx = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === AGENCY_KEY) {
      agencyRowIdx = i;
      break;
    }
  }

  if (agencyRowIdx < 0) {
    // 行が無いので「フォームURL」行の直下に追加
    let formUrlRowIdx = -1;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).includes("フォームURL")) formUrlRowIdx = i;
    }
    const insertAt = formUrlRowIdx >= 0 ? formUrlRowIdx + 2 : 6;
    if (formUrlRowIdx >= 0) sheet.insertRowAfter(formUrlRowIdx + 1);
    const range = sheet.getRange(insertAt, 1, 1, 2);
    range.setValues([[AGENCY_KEY, ""]]);
    range.setBackground("#e0f2fe").setFontColor("#075985").setFontWeight("bold");
    return AGENCY_DEFAULT;
  }

  const code = String(values[agencyRowIdx][1] || "").trim();
  if (!code) return AGENCY_DEFAULT;
  if (!AGENCY_PATTERN.test(code)) {
    throw new Error("代理店コード「" + code + "」は半角英数字とアンダースコアのみ使用できます。");
  }
  return code;
}

// ---- 代理店コードからSS表示名を生成 ----
function getAgencySpreadsheetName(code) {
  if (code === AGENCY_DEFAULT) return AGENCY_DEFAULT_NAME;
  return AGENCY_PREFIX + code;
}

// ---- 代理店スプシ保管フォルダを取得・作成 ----
function getOrCreateAgencyFolder() {
  const iter = DriveApp.getFoldersByName(AGENCY_FOLDER);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(AGENCY_FOLDER);
}

// ---- 代理店SSのIDキー ----
function agencyPropKey(code) { return AGENCY_PROP_PREFIX + code; }

// ---- 代理店SSを取得・作成 ----
function getOrCreateAgencySpreadsheet(code) {
  const props = PropertiesService.getScriptProperties();
  const key   = agencyPropKey(code);
  const ssId  = props.getProperty(key);
  let agencySS = null;

  if (ssId) {
    try { agencySS = SpreadsheetApp.openById(ssId); } catch (e) {}
  }

  if (!agencySS) {
    const name = getAgencySpreadsheetName(code);
    agencySS = SpreadsheetApp.create(name);
    props.setProperty(key, agencySS.getId());

    // フォルダへ移動
    try {
      const file    = DriveApp.getFileById(agencySS.getId());
      const parents = file.getParents();
      while (parents.hasNext()) parents.next().removeFile(file);
      getOrCreateAgencyFolder().addFile(file);
    } catch (e) { Logger.log("フォルダ移動エラー: " + e); }

    // 1枚目を「管理」シートに
    const firstSheet = agencySS.getSheets()[0];
    firstSheet.setName(MANAGEMENT_SHEET);
  }

  // トリガー設置（毎回確認）
  installAgencyTrigger(agencySS);

  return agencySS;
}

// ---- 代理店SSにフォーム設定シート全体（A〜最終列）をコピー ----
function syncFormSheetToAgency(ss, agencySS, formCode) {
  const mainSheet = getConfigSheetByCode(ss, formCode);
  if (!mainSheet) return;
  const displayName = mainSheet.getName().replace(CONFIG_PREFIX, "");

  let agencySheet = getConfigSheetByCode(agencySS, formCode);
  if (!agencySheet) {
    const oldSheet = agencySS.getSheetByName(CONFIG_PREFIX + formCode);
    if (oldSheet) {
      oldSheet.setName(CONFIG_PREFIX + displayName);
      agencySheet = oldSheet;
    } else {
      agencySheet = agencySS.insertSheet(CONFIG_PREFIX + displayName);
    }
  }
  agencySheet.clearContents();
  agencySheet.clearFormats();

  const lastRow = mainSheet.getLastRow();
  const lastCol = mainSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  const allData = mainSheet.getRange(1, 1, lastRow, lastCol).getValues();
  agencySheet.getRange(1, 1, lastRow, lastCol).setValues(allData);

  // スタイル適用
  for (let i = 0; i < allData.length; i++) {
    const key = String(allData[i][0]);
    if (key.includes("フォームURL")) {
      agencySheet.getRange(i + 1, 1, 1, 2).setBackground("#e8f5e9").setFontColor("#1b5e20").setFontWeight("bold");
      agencySheet.getRange(i + 1, 2).setFontStyle("italic");
    }
    if (key === AGENCY_KEY) {
      agencySheet.getRange(i + 1, 1, 1, 2).setBackground("#e0f2fe").setFontColor("#075985").setFontWeight("bold");
    }
    if (key === FORM_NAME_KEY) {
      agencySheet.getRange(i + 1, 1, 1, 2).setBackground("#fff3e0").setFontColor("#e65100").setFontWeight("bold");
    }
    if (key === "フィールドID") {
      agencySheet.getRange(i + 1, 1, 1, 5).setFontWeight("bold").setBackground("#e8eaf6").setFontColor("#1a237e");
    }
  }

  // 回答ヘッダー（G1）スタイル
  if (lastCol >= ANSWER_START_COL && mainSheet.getRange(1, ANSWER_START_COL).getValue()) {
    const answerCols = lastCol - ANSWER_START_COL + 1;
    agencySheet.getRange(1, ANSWER_START_COL, 1, answerCols)
      .setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  }
  agencySheet.setFrozenRows(1);

  agencySheet.setColumnWidth(1, 180);
  agencySheet.setColumnWidth(2, 320);
  agencySheet.setColumnWidth(3, 200);
  agencySheet.setColumnWidth(4, 120);
  agencySheet.setColumnWidth(5, 200);
}

// ---- 代理店割り当て更新（手動メニュー）----
function rebuildAllAgencySpreadsheets() {
  // 旧 onSharingEdit トリガーを先に削除してトリガー枠を確保
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onSharingEdit") ScriptApp.deleteTrigger(t);
  });

  const ss    = getOrCreateSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // フォームを代理店コードでグループ化
  const formByAgency = {};
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    let code;
    try { code = getAgencyCode(sheet); }
    catch (e) {
      Logger.log("代理店コードエラー(" + formCode + "): " + e);
      return;
    }
    if (!formByAgency[code]) formByAgency[code] = [];
    formByAgency[code].push(formCode);
  });
  Logger.log("代理店割り当て: " + JSON.stringify(formByAgency));

  // 既存の全代理店コード（プロパティ＋現在使用中）を集計
  const allCodes = new Set();
  Object.keys(formByAgency).forEach(c => allCodes.add(c));
  props.getKeys().forEach(key => {
    if (key.startsWith(AGENCY_PROP_PREFIX)) {
      allCodes.add(key.replace(AGENCY_PROP_PREFIX, ""));
    }
  });

  allCodes.forEach(code => {
    const agencySS  = getOrCreateAgencySpreadsheet(code);
    const formNames = formByAgency[code] || [];

    // 必要なフォームシートをコピー
    formNames.forEach(formName => {
      syncFormSheetToAgency(ss, agencySS, formName);
    });

    // 不要な設定シートを削除
    agencySS.getSheets().forEach(sheet => {
      const name = sheet.getName();
      if (!name.startsWith(CONFIG_PREFIX)) return;
      const fCode = getFormCodeFromSheet(sheet) || name.replace(CONFIG_PREFIX, "");
      if (formNames.indexOf(fCode) < 0) {
        agencySS.deleteSheet(sheet);
        Logger.log(code + ": 不要シート削除 - " + name);
      }
    });

    // 代理店SSの管理シートを更新
    updateAgencyManagementSheet(agencySS, code, formNames, ss);
  });

  updateManagementSheet();
  Logger.log("代理店割り当て更新 完了");
}

// ---- 代理店SSの管理シートを更新 ----
function updateAgencyManagementSheet(agencySS, code, formCodes, mainSS) {
  let mgSheet = agencySS.getSheetByName(MANAGEMENT_SHEET);
  if (!mgSheet) {
    mgSheet = agencySS.insertSheet(MANAGEMENT_SHEET, 0);
  } else if (mgSheet.getIndex() !== 1) {
    agencySS.setActiveSheet(mgSheet);
    agencySS.moveActiveSheet(1);
  }
  mgSheet.clearContents();
  mgSheet.clearFormats();

  const headers = ["フォーム記号", "フォーム名", "フォームURL", "代理店コード", "回答数", "最終回答日時"];
  const hRange  = mgSheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  mgSheet.setFrozenRows(1);

  const rows = [];
  formCodes.forEach(formCode => {
    const sheet = getConfigSheetByCode(mainSS, formCode);
    if (!sheet) return;
    const values = sheet.getDataRange().getValues();
    let formUrl     = "";
    let displayName = "";
    for (const row of values) {
      if (String(row[0]).includes("フォームURL")) formUrl     = String(row[1] || "");
      if (String(row[0]) === FORM_NAME_KEY)      displayName = String(row[1] || "");
    }

    let answerCount  = 0;
    let lastAnswerAt = "";
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow >= 2 && lastCol >= ANSWER_START_COL) {
      const answerCols = lastCol - ANSWER_START_COL + 1;
      const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, answerCols).getValues();
      const hdrs = sheet.getRange(1, ANSWER_START_COL, 1, answerCols).getValues()[0];
      const rtOff = hdrs.indexOf("受信日時");
      data.forEach(row => {
        if (row.some(c => c !== "")) {
          answerCount++;
          if (rtOff >= 0 && row[rtOff]) lastAnswerAt = String(row[rtOff]);
        }
      });
    }
    rows.push([formCode, displayName, formUrl, code, answerCount, lastAnswerAt]);
  });

  if (rows.length > 0) {
    mgSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    rows.forEach((row, i) => {
      if (row[2]) mgSheet.getRange(i + 2, 3).setFontColor("#1155cc").setFontStyle("italic");
      if (row[3]) mgSheet.getRange(i + 2, 4).setFontColor("#075985").setFontWeight("bold");
    });
    rows.forEach((_, i) => {
      if (i % 2 === 1) mgSheet.getRange(i + 2, 1, 1, headers.length).setBackground("#f3f4f6");
    });
  }

  mgSheet.setColumnWidth(1, 120);
  mgSheet.setColumnWidth(2, 200);
  mgSheet.setColumnWidth(3, 360);
  mgSheet.setColumnWidth(4, 120);
  mgSheet.setColumnWidth(5, 80);
  mgSheet.setColumnWidth(6, 160);
}

// ---- 代理店SS用 onEditトリガーをインストール ----
function installAgencyTrigger(agencySS) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onAgencyEdit" && t.getTriggerSourceId() === agencySS.getId()) return;
  }
  ScriptApp.newTrigger("onAgencyEdit").forSpreadsheet(agencySS).onEdit().create();
}

// ---- 代理店SS側で承認列編集 → メインSSに反映 ----
function onAgencyEdit(e) {
  try {
    const editedSheet = e.source.getActiveSheet();
    const sheetName   = editedSheet.getName();
    if (!sheetName.startsWith(CONFIG_PREFIX)) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1 || col < ANSWER_START_COL) return;

    const lastCol = editedSheet.getLastColumn();
    const headers = editedSheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    if (headers[col - ANSWER_START_COL] !== "承認") return;

    const newValue      = e.range.getValue();
    const receivedAtOff = headers.indexOf("受信日時");
    const receivedAt    = String(editedSheet.getRange(row, ANSWER_START_COL + receivedAtOff).getValue());

    const formCode = getFormCodeFromSheet(editedSheet);
    if (!formCode) return;
    const ss       = getOrCreateSpreadsheet();
    syncApprovalToMain(ss, formCode, receivedAt, newValue);
  } catch (err) {
    Logger.log("onAgencyEdit error: " + err);
  }
}

// ---- メインSS用 onEditトリガーをインストール ----
function installMainEditTrigger() {
  const ss       = getOrCreateSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onMainEdit" && t.getTriggerSourceId() === ss.getId()) return;
  }
  ScriptApp.newTrigger("onMainEdit").forSpreadsheet(ss).onEdit().create();
}

// ---- メインSS側で承認列・管理シート編集 → 反映 ----
function onMainEdit(e) {
  try {
    const editedSheet = e.source.getActiveSheet();
    const sheetName   = editedSheet.getName();

    // 管理シートでフォーム名/代理店コードを編集 → 設定シートへ反映
    if (sheetName === MANAGEMENT_SHEET) {
      handleManagementSheetEdit(e, editedSheet);
      return;
    }

    if (!sheetName.startsWith(CONFIG_PREFIX)) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1 || col < ANSWER_START_COL) return;

    const lastCol = editedSheet.getLastColumn();
    const headers = editedSheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    if (headers[col - ANSWER_START_COL] !== "承認") return;

    const newValue      = e.range.getValue();
    const receivedAtOff = headers.indexOf("受信日時");
    const receivedAt    = String(editedSheet.getRange(row, ANSWER_START_COL + receivedAtOff).getValue());

    const formCode = getFormCodeFromSheet(editedSheet);
    if (!formCode) return;
    let code;
    try { code = getAgencyCode(editedSheet); }
    catch (err) { return; }
    const agencySS = getOrCreateAgencySpreadsheet(code);
    syncApprovalToAgency(agencySS, formCode, receivedAt, newValue);
  } catch (err) {
    Logger.log("onMainEdit error: " + err);
  }
}

// ---- 管理シートでフォーム名/代理店コードを編集 → 設定シートへ反映 ----
function handleManagementSheetEdit(e, mgSheet) {
  const row = e.range.getRow();
  if (row <= 1) return;

  const lastCol = mgSheet.getLastColumn();
  const col     = e.range.getColumn();
  if (col > lastCol) return;

  const headers    = mgSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerName = String(headers[col - 1] || "");

  // フォーム名・代理店コードのみ反映
  if (headerName !== FORM_NAME_KEY && headerName !== AGENCY_KEY) return;

  const formCode = String(mgSheet.getRange(row, 1).getValue() || "").trim();
  if (!formCode) return;

  const ss          = getOrCreateSpreadsheet();
  const configSheet = getConfigSheetByCode(ss, formCode);
  if (!configSheet) return;

  const newValue = String(e.range.getValue() || "").trim();

  if (headerName === AGENCY_KEY) {
    if (newValue && !AGENCY_PATTERN.test(newValue)) {
      Logger.log("無効な代理店コード: " + newValue);
      return;
    }
  }

  setConfigValue(configSheet, headerName, newValue);
}

// ---- 設定シートのキー行に値を書き込む（無ければ何もしない）----
function setConfigValue(sheet, key, value) {
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

// ---- 承認値をメインSSに同期 ----
function syncApprovalToMain(ss, formName, receivedAt, value) {
  const sheet = getConfigSheetByCode(ss, formName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers       = sheet.getRange(1, ANSWER_START_COL, 1, sheet.getLastColumn() - ANSWER_START_COL + 1).getValues()[0];
  const receivedAtCol = ANSWER_START_COL + headers.indexOf("受信日時");
  const approvalCol   = ANSWER_START_COL + headers.indexOf("承認");
  if (receivedAtCol < ANSWER_START_COL || approvalCol < ANSWER_START_COL) return;

  const values = sheet.getRange(2, receivedAtCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === receivedAt) {
      sheet.getRange(i + 2, approvalCol).setValue(value);
      return;
    }
  }
}

// ---- 承認値を代理店SSに同期 ----
function syncApprovalToAgency(agencySS, formName, receivedAt, value) {
  const sheet = getConfigSheetByCode(agencySS, formName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers       = sheet.getRange(1, ANSWER_START_COL, 1, sheet.getLastColumn() - ANSWER_START_COL + 1).getValues()[0];
  const receivedAtCol = ANSWER_START_COL + headers.indexOf("受信日時");
  const approvalCol   = ANSWER_START_COL + headers.indexOf("承認");
  if (receivedAtCol < ANSWER_START_COL || approvalCol < ANSWER_START_COL) return;

  const values = sheet.getRange(2, receivedAtCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === receivedAt) {
      sheet.getRange(i + 2, approvalCol).setValue(value);
      return;
    }
  }
}

// ---- 旧共有SSをゴミ箱へ（手動メニュー）----
function deleteAllOldSharingSpreadsheets() {
  const ss = getOrCreateSpreadsheet();

  // 設定シートの「共有シートURL」「共有シートID」行から対象IDを収集
  const idsToTrash = new Set();
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const values = sheet.getDataRange().getValues();
    for (const row of values) {
      if (String(row[0]) === "共有シートURL" && row[1]) {
        const id = extractSsIdFromUrl(String(row[1]));
        if (id) idsToTrash.add(id);
      }
      if (String(row[0]) === "共有シートID" && row[1]) {
        idsToTrash.add(String(row[1]).trim());
      }
    }
  });

  // 旧フォルダ「アフィ共有スプシ」内のSSも対象に
  try {
    const iter = DriveApp.getFoldersByName("アフィ共有スプシ");
    if (iter.hasNext()) {
      const folder = iter.next();
      const files  = folder.getFiles();
      while (files.hasNext()) idsToTrash.add(files.next().getId());
    }
  } catch (e) {}

  let trashed = 0;
  idsToTrash.forEach(id => {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      trashed++;
    } catch (e) { Logger.log("削除エラー: " + id + " / " + e); }
  });

  // 設定シートから「共有シートURL」「共有シートID」行を削除
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const values = sheet.getDataRange().getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (String(values[i][0]) === "共有シートURL" || String(values[i][0]) === "共有シートID") {
        sheet.deleteRow(i + 1);
      }
    }
  });

  // 旧 onSharingEdit トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onSharingEdit") ScriptApp.deleteTrigger(t);
  });

  Logger.log("旧共有SS削除: " + trashed + "件");
}

// =============================================
// 管理シート（メインSS）
// =============================================

function updateManagementSheet() {
  const ss    = getOrCreateSpreadsheet();
  let mgSheet = ss.getSheetByName(MANAGEMENT_SHEET);
  if (!mgSheet) {
    mgSheet = ss.insertSheet(MANAGEMENT_SHEET, 0);
  }
  mgSheet.clearContents();
  mgSheet.clearFormats();

  const headers = ["フォーム記号", "フォーム名", "フォームURL", "代理店コード", "回答数", "最終回答日時"];
  const hRange  = mgSheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  mgSheet.setFrozenRows(1);

  const rows = [];
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    const values   = sheet.getDataRange().getValues();

    let formUrl     = "";
    let agencyCode  = "";
    let displayName = "";
    for (const row of values) {
      if (String(row[0]).includes("フォームURL")) formUrl     = String(row[1] || "");
      if (String(row[0]) === AGENCY_KEY)         agencyCode  = String(row[1] || "");
      if (String(row[0]) === FORM_NAME_KEY)      displayName = String(row[1] || "");
    }

    let answerCount  = 0;
    let lastAnswerAt = "";
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow >= 2 && lastCol >= ANSWER_START_COL) {
      const answerCols = lastCol - ANSWER_START_COL + 1;
      const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, answerCols).getValues();
      const hdrs = sheet.getRange(1, ANSWER_START_COL, 1, answerCols).getValues()[0];
      const rtOff = hdrs.indexOf("受信日時");
      data.forEach(row => {
        if (row.some(c => c !== "")) {
          answerCount++;
          if (rtOff >= 0 && row[rtOff]) lastAnswerAt = String(row[rtOff]);
        }
      });
    }
    rows.push([formCode, displayName, formUrl, agencyCode, answerCount, lastAnswerAt]);
  });

  if (rows.length > 0) {
    mgSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    rows.forEach((row, i) => {
      if (row[2]) mgSheet.getRange(i + 2, 3).setFontColor("#1155cc").setFontStyle("italic");
      if (row[3]) mgSheet.getRange(i + 2, 4).setFontColor("#075985").setFontWeight("bold");
    });
    rows.forEach((_, i) => {
      if (i % 2 === 1) mgSheet.getRange(i + 2, 1, 1, headers.length).setBackground("#f3f4f6");
    });
  }

  mgSheet.setColumnWidth(1, 120);
  mgSheet.setColumnWidth(2, 200);
  mgSheet.setColumnWidth(3, 360);
  mgSheet.setColumnWidth(4, 120);
  mgSheet.setColumnWidth(5, 80);
  mgSheet.setColumnWidth(6, 160);
}

// =============================================
// LINE 通知
// =============================================

// ---- LINE Webhook 受信 ----
function handleLineWebhook(data) {
  const events = data.events || [];
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty("LINE_CHANNEL_TOKEN");

  if (token && events.length > 0) {
    events.forEach(event => {
      if (event.source && event.source.type === "group") {
        const groupId = event.source.groupId;
        // グループIDをスクリプトプロパティに保存（確認用）
        props.setProperty("DETECTED_GROUP_ID", groupId);

        // スプレッドシートにも記録
        try {
          const ss    = getOrCreateSpreadsheet();
          let logSheet = ss.getSheetByName("グループIDログ");
          if (!logSheet) logSheet = ss.insertSheet("グループIDログ");
          logSheet.getRange(1, 1).setValue(groupId);
          logSheet.getRange(2, 1).setValue(new Date());
        } catch (e) { Logger.log("グループIDログエラー: " + e); }

        // グループIDコマンドへの返信も試みる
        if (event.type === "message" && event.message && event.message.type === "text" &&
            event.message.text.trim() === "グループID") {
          try {
            UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
              method: "post",
              headers: {
                "Content-Type":  "application/json",
                "Authorization": "Bearer " + token
              },
              payload: JSON.stringify({
                replyToken: event.replyToken,
                messages: [{ type: "text", text: "グループID: " + groupId }]
              }),
              muteHttpExceptions: true
            });
          } catch (e) { Logger.log("返信エラー: " + e); }
        }
      }
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- 保存済みグループIDを確認（GASエディタから実行）----
function checkDetectedGroupId() {
  const props   = PropertiesService.getScriptProperties();
  const groupId = props.getProperty("DETECTED_GROUP_ID");
  Logger.log("検出されたグループID: " + (groupId || "まだ受信なし"));
}

// ---- LINE グループへプッシュ通知 ----
function notifyLineGroup(message) {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty("LINE_CHANNEL_TOKEN");
  const groupId = props.getProperty("LINE_GROUP_ID");
  if (!token || !groupId) return;
  UrlFetchApp.fetch(LINE_PUSH_API, {
    method: "post",
    headers: {
      "Content-Type":  "application/json",
      "Authorization": "Bearer " + token
    },
    payload: JSON.stringify({
      to: groupId,
      messages: [{ type: "text", text: message.length > 4990 ? message.substring(0, 4990) + "..." : message }]
    }),
    muteHttpExceptions: true
  });
}

// ---- 通知メッセージを組み立て ----
function buildLineMessage(config, rowData, formCode) {
  const headline = config.formDisplayName || config.formTitle || formCode;
  const lines = ["【新規申請】" + headline];
  lines.push("受信日時: " + rowData[1]);
  config.fields.forEach((field, i) => {
    const val = rowData[4 + i];
    if (val !== "" && val !== undefined) lines.push(field.label + ": " + val);
  });
  // スクショURLの位置は buildHeaders の並びから引く。末尾からの相対位置で取ると、
  // 末尾に列が増えたときに黙って別の列を読む（2026-08-19 に「代理店」を末尾へ足した結果、
  // 承認列（常に空）を読んでスクショ行がLINE通知から消えていた）。
  const shotIdx = buildHeaders(config).lastIndexOf("スクショURL");
  const shotVal = shotIdx >= 0 ? String(rowData[shotIdx] || "").trim() : "";
  // スクショはフォーム側で必須にしてあるので「URLが無い」＝異常。行ごと落とすと今回と同じ
  // 「気づけない消え方」になるため、保存失敗も取得不能も1行として必ず出す。
  if (shotVal.startsWith("http")) {
    lines.push("スクショ: " + shotVal);
  } else if (shotVal) {
    lines.push("スクショ: 保存に失敗しました（" + shotVal.substring(0, 120) + "）");
  } else {
    lines.push("スクショ: 取得できませんでした（要確認）");
  }
  return lines.join("\n");
}

// ---- LINE設定をスクリプトプロパティに保存（GASエディタから手動実行） ----
function setLineGroupId(groupId) {
  PropertiesService.getScriptProperties().setProperty("LINE_GROUP_ID", groupId);
  Logger.log("LINE_GROUP_ID を設定しました: " + groupId);
}

// =============================================
// 日次レポート（毎朝9時に前日結果をLINE通知）
// =============================================

// ---- 前日のJST日付文字列を返す（yyyy/MM/dd） ----
function getYesterdayJSTDateStr() {
  const now          = new Date();
  const jstNow       = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const jstYesterday = new Date(jstNow.getTime() - 24 * 60 * 60 * 1000);
  const p = n => String(n).padStart(2, "0");
  return jstYesterday.getUTCFullYear() + "/" +
         p(jstYesterday.getUTCMonth() + 1) + "/" +
         p(jstYesterday.getUTCDate());
}

// ---- セル値（Date or 文字列）からJST日付文字列「yyyy/MM/dd」を抽出 ----
function toJSTDateStr(value) {
  if (!value) return "";
  if (value instanceof Date) {
    const jst = new Date(value.getTime() + 9 * 60 * 60 * 1000);
    const p   = n => String(n).padStart(2, "0");
    return jst.getUTCFullYear() + "/" +
           p(jst.getUTCMonth() + 1) + "/" +
           p(jst.getUTCDate());
  }
  return String(value).substring(0, 10);
}

// ---- 今月のJST年月文字列を返す（yyyy/MM） ----
function getThisMonthJSTStr() {
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p   = n => String(n).padStart(2, "0");
  return jst.getUTCFullYear() + "/" + p(jst.getUTCMonth() + 1);
}

// ---- カタカナ→ひらがな変換 ----
function toHiragana(str) {
  return str.replace(/[ァ-ヶ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0x60));
}

// ---- 名前の正規化（あらゆるスペース除去・カタカナ→ひらがな・小文字化） ----
function normalizeName(name) {
  let s = String(name || "").trim();
  // 全種類のスペース・不可視文字を除去
  s = s.replace(/[\s ​‌‍ -‏  　﻿]+/g, "");
  return toHiragana(s).toLowerCase();
}

// ---- レーベンシュタイン距離 ----
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, (_, i) => [i]);
  for (let j = 1; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[m][n];
}

// ---- 紹介者名を名寄せしてグループ化する ----
// ルール1: 正規化後に完全一致 → 同一人物
// ルール2: 短い方が長い方の「先頭」に一致（2文字以上） → 同一人物（苗字のみ vs フルネーム）
// ルール3: 正規化後の長さ4以上かつレーベンシュタイン距離1以内 → 同一人物（漢字表記ゆれ対応）
function groupReferrers(rawCounts) {
  // Step1: 正規化マップ
  const normMap = {};
  Object.entries(rawCounts).forEach(([raw, cnt]) => {
    const norm = normalizeName(raw);
    if (!norm) return;
    if (!normMap[norm]) normMap[norm] = { count: 0, raws: {} };
    normMap[norm].count += cnt;
    normMap[norm].raws[raw] = (normMap[norm].raws[raw] || 0) + cnt;
  });

  const entries = Object.entries(normMap)
    .map(([norm, d]) => ({ norm, count: d.count, raws: d.raws }))
    .sort((a, b) => b.norm.length - a.norm.length);

  const groups = [];
  const used   = new Set();

  entries.forEach(entry => {
    if (used.has(entry.norm)) return;
    const group = { count: entry.count, raws: { ...entry.raws } };
    used.add(entry.norm);

    entries.forEach(other => {
      if (used.has(other.norm)) return;
      const longer  = entry.norm.length >= other.norm.length ? entry.norm : other.norm;
      const shorter = entry.norm.length <  other.norm.length ? entry.norm : other.norm;

      // 部分一致（どの位置でもOK）または表記ゆれ（Levenshtein距離1）
      const isSubstr = shorter.length >= 2 && longer.includes(shorter);
      const isTypo   = shorter.length >= 4 && levenshtein(entry.norm, other.norm) <= 1;

      if (isSubstr || isTypo) {
        group.count += other.count;
        Object.entries(other.raws).forEach(([r, c]) => {
          group.raws[r] = (group.raws[r] || 0) + c;
        });
        used.add(other.norm);
      }
    });

    // 表示名：最も出現の多い表記を代表にする
    const mostFreqRaw = Object.entries(group.raws).sort((a, b) => b[1] - a[1])[0][0];
    const canonical = mostFreqRaw;
    const variants  = Object.keys(group.raws).filter(r => r !== mostFreqRaw);

    group.canonical = canonical;
    group.variants  = variants;
    groups.push(group);
  });

  return groups.sort((a, b) => b.count - a.count);
}

// ---- 営業マン別 今月累計セクションを生成 ----
function buildMonthlyReferrerSection(ss, thisMonth) {
  const rawCounts = {};

  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx   = headers.indexOf("受信日時");
    const refIdx  = headers.indexOf("紹介者名");
    if (rtIdx < 0 || refIdx < 0) return;

    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
      .forEach(row => {
        if (!toJSTDateStr(row[rtIdx]).startsWith(thisMonth)) return;
        const ref = String(row[refIdx] || "").trim();
        if (!ref) return;
        rawCounts[ref] = (rawCounts[ref] || 0) + 1;
      });
  });

  if (Object.keys(rawCounts).length === 0) return "";

  const groups = groupReferrers(rawCounts);
  const lines  = ["▼ 営業マン別 今月累計"];
  groups.forEach(g => {
    const label = g.variants.length > 0
      ? g.canonical + "（" + g.variants.join("・") + "）"
      : g.canonical;
    lines.push("・" + label + ": " + g.count + "件");
  });
  return lines.join("\n");
}

// ---- 日次レポート本体 ----
function dailyReport() {
  const ss        = getOrCreateSpreadsheet();
  const yesterday = getYesterdayJSTDateStr();
  const thisMonth = getThisMonthJSTStr();

  const reports = [];
  let totalDaily   = 0;
  let totalMonthly = 0;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const formCode    = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    const displayName = getFormDisplayName(sheet, formCode);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx   = headers.indexOf("受信日時");
    const refIdx  = headers.indexOf("紹介者名");
    if (rtIdx < 0) return;

    const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues();

    let dailyCount   = 0;
    let monthlyCount = 0;
    const referrerMap = {};

    data.forEach(row => {
      const dateStr = toJSTDateStr(row[rtIdx]);
      if (!dateStr) return;
      if (dateStr.startsWith(thisMonth)) monthlyCount++;
      if (dateStr === yesterday) {
        dailyCount++;
        if (refIdx >= 0) {
          const ref = String(row[refIdx] || "").trim() || "不明";
          referrerMap[ref] = (referrerMap[ref] || 0) + 1;
        }
      }
    });

    if (dailyCount === 0 && monthlyCount === 0) return;

    let line = "・" + displayName + ": " + dailyCount + "件（月計: " + monthlyCount + "件）";
    if (dailyCount > 0 && refIdx >= 0) {
      const refLines = Object.entries(referrerMap)
        .sort((a, b) => b[1] - a[1])
        .map(([ref, cnt]) => "  　" + ref + ": " + cnt + "件");
      line += "\n" + refLines.join("\n");
    }

    reports.push(line);
    totalDaily   += dailyCount;
    totalMonthly += monthlyCount;
  });

  const referrerSection = buildMonthlyReferrerSection(ss, thisMonth);

  // 顧客向け公式LINEの「営業担当ごとの登録数」。統合アプリから取る。
  // 未設定・取得失敗なら空文字が返るので、日次レポート自体は止まらない。
  let customerLineSection = "";
  try {
    customerLineSection = buildCustomerLineStatsSection_(yesterday);
  } catch (e) {
    Logger.log("顧客LINE登録の節を作れませんでした: " + e);
  }

  let message;
  if (reports.length === 0) {
    message = "【日次レポート】" + yesterday + "\n\n申請はありませんでした。\n\n今月累計: 0件";
    if (referrerSection) message += "\n\n" + referrerSection;
  } else {
    message = "【日次レポート】" + yesterday + "\n\n"
            + reports.join("\n\n")
            + "\n\n昨日合計: " + totalDaily + "件"
            + "\n今月累計: " + totalMonthly + "件";
    if (referrerSection) message += "\n\n" + referrerSection;
  }
  if (customerLineSection) message += "\n\n" + customerLineSection;

  notifyLineGroup(message);
  Logger.log(message);
}

// ---- 日次レポート用トリガー設置（無ければ作成） ----
function ensureDailyReportTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction() === "dailyReport") return;
    }
    ScriptApp.newTrigger("dailyReport")
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    Logger.log("dailyReport トリガーを設置しました（毎朝9時JST）");
  } catch (e) {
    Logger.log("dailyReport トリガー設置失敗: " + e);
  }
}

// ---- LINE通知テスト（GASエディタから手動実行）----
// ---- OAuth スコープ承認用（GASエディタから手動実行） ----
function authorizeScopes() {
  DriveApp.getRootFolder();
  Logger.log("Drive スコープ承認済み");
}

function testLineNotification() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty("LINE_CHANNEL_TOKEN");
  const groupId = props.getProperty("LINE_GROUP_ID");
  Logger.log("TOKEN: " + (token ? token.substring(0, 10) + "..." : "未設定"));
  Logger.log("GROUP_ID: " + (groupId || "未設定"));
  if (!token || !groupId) { Logger.log("プロパティ未設定"); return; }

  const res = UrlFetchApp.fetch(LINE_PUSH_API, {
    method: "post",
    headers: { "Content-Type": "application/json", "Authorization": "Bearer " + token },
    payload: JSON.stringify({ to: groupId, messages: [{ type: "text", text: "テスト通知です" }] }),
    muteHttpExceptions: true
  });
  Logger.log("HTTP: " + res.getResponseCode());
  Logger.log("Body: " + res.getContentText());
}

// =============================================
// 新規フォーム作成ダイアログ
// =============================================

function showCreateFormDialog() {
  const html = HtmlService.createHtmlOutputFromFile("dialog")
    .setWidth(480)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, "新規フォーム作成");
}

function createFormFromDialog(data) {
  const formCode        = String(data.formCode        || data.formName || "").trim();
  const formDisplayName = String(data.formDisplayName || "").trim();
  const formTitle       = String(data.formTitle       || "").trim();
  const affiliateUrl    = String(data.affiliateUrl    || "").trim();
  const agencyCode      = String(data.agencyCode      || "").trim();

  if (!formCode)     throw new Error("フォーム記号を入力してください。");
  if (!FORM_CODE_PATTERN.test(formCode)) {
    throw new Error("フォーム記号は半角英数字とアンダースコアのみ使用できます。");
  }
  if (!formTitle)    throw new Error("フォームタイトルを入力してください。");
  if (!affiliateUrl) throw new Error("アフィリエイトURLを入力してください。");
  if (agencyCode && !AGENCY_PATTERN.test(agencyCode)) {
    throw new Error("代理店コードは半角英数字とアンダースコアのみ使用できます。");
  }

  const ss = getOrCreateSpreadsheet();
  if (getConfigSheetByCode(ss, formCode)) {
    throw new Error("「" + formCode + "」は既に存在します。");
  }

  const sheetDisplayName = formDisplayName || formCode;
  const configSheet = ss.insertSheet(CONFIG_PREFIX + sheetDisplayName);
  initConfigSheet(configSheet, formTitle, formDisplayName, formCode);

  // アフィリエイトURL・代理店コード書き込み
  const values = configSheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === "アフィリエイトURL") {
      configSheet.getRange(i + 1, 2).setValue(affiliateUrl);
    }
    if (String(values[i][0]) === AGENCY_KEY && agencyCode) {
      configSheet.getRange(i + 1, 2).setValue(agencyCode);
    }
  }

  // 回答ヘッダー初期化
  const config = readConfig(ss, formCode);
  const hdrs   = buildHeaders(config);
  const hRange = configSheet.getRange(1, ANSWER_START_COL, 1, hdrs.length);
  hRange.setValues([hdrs]);
  hRange.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  configSheet.setFrozenRows(1);

  // 代理店SSに同期
  const code     = agencyCode || AGENCY_DEFAULT;
  const agencySS = getOrCreateAgencySpreadsheet(code);
  syncFormSheetToAgency(ss, agencySS, formCode);
  // 代理店SSの管理シートも更新
  const formCodes = [];
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    try {
      if (getAgencyCode(sheet) === code) {
        const c = getFormCodeFromSheet(sheet);
        if (c) formCodes.push(c);
      }
    } catch (e) {}
  });
  updateAgencyManagementSheet(agencySS, code, formCodes, ss);

  // メインSSの管理シート更新
  updateManagementSheet();

  return {
    formCode: formCode,
    formDisplayName: formDisplayName,
    formUrl: FORM_BASE_URL + "?form=" + encodeURIComponent(formCode)
  };
}

// ---- 既存シートをフォーム名ベースのシート名に移行（手動メニューから実行）----
function migrateSheetNamesToDisplayName() {
  const ss = getOrCreateSpreadsheet();
  let renamed = 0;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;

    const values = sheet.getDataRange().getValues();
    let hasCodeRow  = false;
    let displayName = "";
    let displayRowIdx = -1;

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === FORM_CODE_HEADER) { hasCodeRow = true; }
      if (String(values[i][0]) === FORM_NAME_KEY) {
        displayName    = String(values[i][1] || "").trim();
        displayRowIdx  = i;
      }
    }

    const formCode = name.replace(CONFIG_PREFIX, "");

    if (!hasCodeRow) {
      // フォーム名行の直下に フォーム記号 行を追加
      if (displayRowIdx >= 0) {
        const insertAt = displayRowIdx + 2;
        sheet.insertRowAfter(displayRowIdx + 1);
        const range = sheet.getRange(insertAt, 1, 1, 2);
        range.setValues([[FORM_CODE_HEADER, formCode]]);
        range.setBackground("#e8eaf6").setFontColor("#1a237e").setFontWeight("bold");
      }
    }

    // シート名をフォーム名ベースに変更（まだ旧フォーマットの場合）
    if (displayName && displayName !== formCode) {
      const newName = CONFIG_PREFIX + displayName;
      if (!ss.getSheetByName(newName)) {
        sheet.setName(newName);
        renamed++;
      }
    }
  });

  SpreadsheetApp.getUi().alert("移行完了。" + renamed + " 件のシートを変更しました。\n変更後は「代理店割り当て更新」も実行してください。");
}

// ---- フォーム記号が崩れたシートをG列の申請データから復元 ----
function repairFormCodeRows() {
  const ss = getOrCreateSpreadsheet();
  const report = [];
  let fixed = 0;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;

    // フォーム記号行を探す
    const values = sheet.getDataRange().getValues();
    let codeRowIdx = -1;
    let currentCode = "";
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === FORM_CODE_HEADER) {
        codeRowIdx = i;
        currentCode = String(values[i][1] || "").trim();
        break;
      }
    }

    // 既に半角英数字なら正常
    if (currentCode && FORM_CODE_PATTERN.test(currentCode)) {
      report.push("✓ " + name + "  記号: " + currentCode);
      return;
    }

    // G列の申請データから半角英数字のフォーム記号を探して復元
    const lastRow = sheet.getLastRow();
    let recovered = null;
    if (lastRow >= 2) {
      const gData = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, 1).getValues();
      for (const row of gData) {
        const candidate = String(row[0] || "").trim();
        if (candidate && FORM_CODE_PATTERN.test(candidate)) {
          recovered = candidate;
          break;
        }
      }
    }

    if (recovered) {
      if (codeRowIdx >= 0) {
        sheet.getRange(codeRowIdx + 1, 2).setValue(recovered);
      }
      fixed++;
      report.push("✓修正: " + name + "  記号: " + recovered);
    } else {
      report.push("✗要手動: " + name + "  (現在: " + (currentCode || "未設定") + ")");
    }
  });

  if (fixed > 0) {
    updateAllFormUrls();
    updateManagementSheet();
  }

  SpreadsheetApp.getUi().alert(
    "フォーム記号の修復完了（" + fixed + "件修正）\n\n" + report.join("\n")
  );
}

// =============================================
// 30件クエスト キャンペーンレポート
// =============================================

function campaignReport() {
  const now    = new Date();
  const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p      = n => String(n).padStart(2, "0");
  const todayJst     = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());
  const thisMonth    = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1);
  const yesterdayJst = getYesterdayJSTDateStr();

  if (todayJst > CAMPAIGN_END_STR) return;

  const todayMs  = new Date(todayJst.replace(/\//g, "-")).getTime();
  const endMs    = new Date(CAMPAIGN_END_STR.replace(/\//g, "-")).getTime();
  const daysLeft = Math.max(1, Math.round((endMs - todayMs) / 86400000) + 1);

  const ss             = getOrCreateSpreadsheet();
  const rawReferrers   = {};  // 今月累計
  const yesterdayRefs  = {};  // 前日のみ
  const formCounts     = {};
  CAMPAIGN_FORMS.forEach(fc => { formCounts[fc] = 0; });

  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode || !CAMPAIGN_FORMS.includes(formCode)) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx   = headers.indexOf("受信日時");
    const refIdx  = headers.indexOf("紹介者名");
    if (rtIdx < 0) return;
    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
      .forEach(row => {
        const dateStr = toJSTDateStr(row[rtIdx]);
        if (!dateStr || !dateStr.startsWith(thisMonth)) return;
        formCounts[formCode]++;
        if (refIdx >= 0) {
          const ref = String(row[refIdx] || "").trim();
          if (ref) {
            rawReferrers[ref] = (rawReferrers[ref] || 0) + 1;
            if (dateStr === yesterdayJst) {
              yesterdayRefs[ref] = (yesterdayRefs[ref] || 0) + 1;
            }
          }
        }
      });
  });

  const weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  const dow      = weekDays[jstNow.getUTCDay()];
  const lines    = [];
  lines.push("⚔️【30件クエスト】進捗レポート");
  lines.push("📅 " + (jstNow.getUTCMonth()+1) + "/" + jstNow.getUTCDate() + "（" + dow + "） 残り" + daysLeft + "日");
  lines.push("");

  let allCleared = true;

  CAMPAIGN_FORMS.forEach(formCode => {
    const count    = formCounts[formCode] || 0;
    const remain   = Math.max(0, CAMPAIGN_TARGET - count);
    const pct      = count / CAMPAIGN_TARGET;
    const formName = CAMPAIGN_FORM_NAMES[formCode] || formCode;
    const formUrl  = FORM_BASE_URL + "?form=" + formCode;
    const filledBlocks = Math.min(10, Math.round(count / CAMPAIGN_TARGET * 10));
    const bar          = "🟩".repeat(filledBlocks) + "🟥".repeat(10 - filledBlocks);
    const paceStr  = remain > 0
      ? "あと" + remain + "件（1日" + Math.ceil(remain / daysLeft) + "件ペース）"
      : "COMPLETE！";

    let icon, comment;
    if (count >= CAMPAIGN_TARGET) { icon = "✅"; comment = "クリア！完璧だ！おめでとう！"; }
    else if (pct >= 0.8)          { icon = "🔥"; comment = "もうすぐだ！全力ラストスパート！"; }
    else if (pct >= 0.6)          { icon = "⚡"; comment = "いい調子！このまま突き進め！"; }
    else if (pct >= 0.3)          { icon = "🌱"; comment = "加速しろ！まだ十分間に合う！"; }
    else if (count > 0)           { icon = "🚨"; comment = "ギアを上げろ！総力戦だ！"; }
    else                          { icon = "💀"; comment = "DANGER！今すぐ動け！"; }

    if (count < CAMPAIGN_TARGET) allCleared = false;

    lines.push(icon + " " + formName);
    lines.push("　" + formUrl);
    lines.push("　" + count + " / " + CAMPAIGN_TARGET + "件　" + paceStr);
    lines.push("　" + bar);
    lines.push("　💬 " + comment);
    lines.push("");
  });

  const totalCount  = Object.values(formCounts).reduce((a, b) => a + b, 0);
  const totalTarget = CAMPAIGN_TARGET * CAMPAIGN_FORMS.length;
  const totalPct    = Math.round(totalCount / totalTarget * 100);
  lines.push("━━━━━━━━━━━━━━");
  lines.push("📊 全体: " + totalCount + " / " + totalTarget + "件 (" + totalPct + "%)");
  if (allCleared) lines.push("🎊 全案件クリア！伝説の営業チームだ！");

  if (Object.keys(rawReferrers).length > 0) {
    const groups  = groupReferrers(rawReferrers);
    const topList = groups.slice(0, 5);
    const medals  = ["🥇", "🥈", "🥉", "4位", "5位"];
    lines.push("");
    lines.push("🏆 MVP ランキング");
    topList.forEach((g, i) => {
      const delta    = Object.keys(g.raws).reduce((sum, raw) => sum + (yesterdayRefs[raw] || 0), 0);
      const trendStr = delta > 0 ? "（前日+" + delta + "📈）" : "";
      lines.push(medals[i] + " " + g.canonical + "　" + g.count + "件" + trendStr);
      lines.push("　" + "▰".repeat(g.count));
    });
  }

  const message = lines.join("\n");
  notifyLineGroup(message);
  Logger.log(message);
}

// ---- キャンペーンレポート用トリガー設置（毎朝8時）----
function ensureCampaignReportTrigger() {
  try {
    const now = new Date();
    const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    const p = n => String(n).padStart(2, "0");
    const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());
    if (todayJst > CAMPAIGN_END_STR) return;
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction() === "campaignReport") return;
    }
    ScriptApp.newTrigger("campaignReport")
      .timeBased().everyDays(1).atHour(8).create();
    Logger.log("campaignReport トリガーを設置しました（毎朝8時）");
  } catch (e) {
    Logger.log("campaignReport トリガー設置失敗: " + e);
  }
}

// ---- 現状診断（GASエディタのRunボタンから実行 → Logsで確認）----
function diagnose() {
  const ss = getOrCreateSpreadsheet();
  Logger.log("=== スプレッドシート診断 ===");
  Logger.log("SS名: " + ss.getName());
  Logger.log("SS URL: " + ss.getUrl());
  Logger.log("");

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    Logger.log("--- シート: " + name + " ---");

    if (!name.startsWith(CONFIG_PREFIX)) {
      Logger.log("  (設定シートではない)");
      return;
    }

    const values = sheet.getDataRange().getValues();
    let formCode    = "(未設定)";
    let displayName = "(未設定)";
    let formUrl     = "(未設定)";
    let agencyCode  = "(未設定)";

    for (const row of values) {
      const key = String(row[0] || "").trim();
      const val = String(row[1] || "").trim();
      if (key === FORM_CODE_HEADER) formCode    = val || "(空)";
      if (key === FORM_NAME_KEY)    displayName = val || "(空)";
      if (key.includes("フォームURL")) formUrl  = val || "(空)";
      if (key === AGENCY_KEY)       agencyCode  = val || "(空)";
    }

    const isValidCode = FORM_CODE_PATTERN.test(formCode);
    Logger.log("  フォーム記号: " + formCode + (isValidCode ? " ✓" : " ✗ (半角英数字でない)"));
    Logger.log("  フォーム名: " + displayName);
    Logger.log("  フォームURL: " + formUrl);
    Logger.log("  代理店: " + agencyCode);

    // G列の状態
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const g1val = sheet.getRange(1, ANSWER_START_COL).getValue();
      const g2val = sheet.getRange(2, ANSWER_START_COL).getValue();
      Logger.log("  G1(ヘッダー): " + g1val);
      Logger.log("  G2(最初の申請フォーム記号): " + g2val);
    } else {
      Logger.log("  G列: データなし");
    }
    Logger.log("");
  });
  Logger.log("=== 診断終了 ===");
}

// ---- 自社（代理店コードなし）フォームを全件取得: [{code, displayName}] ----
function getJishaForms() {
  const ss    = getOrCreateSpreadsheet();
  const forms = [];
  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    // 代理店コード行を読み取り専用でチェック（getAgencyCodeは副作用があるため直接読む）
    const values = sheet.getDataRange().getValues();
    let agencyCode = AGENCY_DEFAULT;
    for (const row of values) {
      if (String(row[0]) === AGENCY_KEY) {
        const code = String(row[1] || "").trim();
        if (code) agencyCode = code;
        break;
      }
    }
    if (agencyCode !== AGENCY_DEFAULT) return; // 代理店フォームは除外
    forms.push({ code: formCode, displayName: getFormDisplayName(sheet, formCode) });
  });
  return forms;
}

// ---- formCodeの表示名をCSSのヘッダーから動的に解決 ----
function resolveFormDisplayName(formCode) {
  if (CAMPAIGN_FORM_NAMES[formCode]) return CAMPAIGN_FORM_NAMES[formCode];
  const ss = getOrCreateSpreadsheet();
  for (const sheet of ss.getSheets()) {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) continue;
    if (getFormCodeFromSheet(sheet) === formCode) return getFormDisplayName(sheet, formCode);
  }
  return formCode;
}

// ---- 顧客管理スプレッドシートを新規作成 ----
function createCustomerManagementSheet() {
  const SALESPEOPLE    = JISHA_REFERRER_OPTIONS.split(",").map(s => s.trim()).filter(Boolean);
  const jishaForms     = getJishaForms();
  const CASE_NAMES     = jishaForms.map(f => f.displayName);
  const STATUS_OPTIONS = ["申請中", "申請済", "完了", "不参加"];
  const BASE_HEADERS   = ["顧客名", "電話番号", "顧客ID"];

  if (CASE_NAMES.length === 0) {
    SpreadsheetApp.getUi().alert("自社フォームが見つかりませんでした。");
    return;
  }

  const ss           = SpreadsheetApp.create("顧客管理_アフィリエイト");
  const defaultSheet = ss.getSheets()[0];

  SALESPEOPLE.forEach((name, idx) => {
    const sheet = idx === 0 ? defaultSheet : ss.insertSheet(name);
    if (idx === 0) sheet.setName(name);

    const headers = [...BASE_HEADERS, ...CASE_NAMES];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, BASE_HEADERS.length)
         .setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(1, BASE_HEADERS.length + 1, 1, CASE_NAMES.length)
         .setBackground("#312e81").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 110);
    CASE_NAMES.forEach((_, i) => sheet.setColumnWidth(BASE_HEADERS.length + 1 + i, 130));

    const caseBodyRange = sheet.getRange(2, BASE_HEADERS.length + 1, 200, CASE_NAMES.length);
    caseBodyRange.setBackground("#ede9fe").setHorizontalAlignment("center");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_OPTIONS).setAllowInvalid(true).build();
    caseBodyRange.setDataValidation(rule);
  });

  PropertiesService.getScriptProperties().setProperty("CUSTOMER_MGMT_SS_ID", ss.getId());

  const url = ss.getUrl();
  Logger.log("顧客管理シート作成完了: " + url);
  SpreadsheetApp.getUi().alert(
    "顧客管理シートを作成しました！\n案件数: " + CASE_NAMES.length + "件\n\n" + url
  );
}

// ---- 顧客管理SSを取得（IDで取得、なければ名前で検索してIDを保存）----
function getCustomerManagementSS() {
  const props = PropertiesService.getScriptProperties();
  const id    = props.getProperty("CUSTOMER_MGMT_SS_ID");
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {
      props.deleteProperty("CUSTOMER_MGMT_SS_ID");
    }
  }
  const files = DriveApp.getFilesByName("顧客管理_アフィリエイト");
  if (files.hasNext()) {
    const file = files.next();
    props.setProperty("CUSTOMER_MGMT_SS_ID", file.getId());
    return SpreadsheetApp.openById(file.getId());
  }
  return null;
}

// ---- 顧客管理SSの指定営業マンシートに顧客を追加/更新 ----
// ヘッダーを動的に参照するので案件追加時も自動対応
function upsertCustomerRow(salesName, customerName, formCode) {
  const css = getCustomerManagementSS();
  if (!css) return;

  const normalizedSales = normalizeName(salesName);
  let sheet = null;
  for (const s of css.getSheets()) {
    if (normalizeName(s.getName()) === normalizedSales) { sheet = s; break; }
  }
  if (!sheet) return;

  // ヘッダーからformCodeの表示名を探して列番号を決定
  const lastCol    = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers    = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const targetName = resolveFormDisplayName(formCode);
  let caseCol      = headers.indexOf(targetName) + 1; // 1-based (0 → not found)
  if (caseCol <= 0) {
    // 案件列が未作成なら自動追加（新案件が顧客管理に反映されるように）
    caseCol = lastCol + 1;
    addCaseColumnToSheet_(sheet, caseCol, targetName);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    for (let i = 0; i < names.length; i++) {
      if (normalizeName(String(names[i])) === normalizeName(customerName)) {
        const cell = sheet.getRange(i + 2, caseCol);
        if (!cell.getValue()) cell.setValue("申請済");
        return;
      }
    }
  }

  const newRow = Math.max(2, lastRow + 1);
  sheet.getRange(newRow, 1).setValue(customerName);
  sheet.getRange(newRow, caseCol).setValue("申請済");
}

// ---- 案件列を1つ追加（ヘッダー書式・幅・入力規則）----
function addCaseColumnToSheet_(sheet, col, caseName) {
  const STATUS_OPTIONS = ["申請中", "申請済", "完了", "不参加"];
  sheet.getRange(1, col).setValue(caseName)
       .setBackground("#312e81").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setColumnWidth(col, 130);
  const bodyRange = sheet.getRange(2, col, 200, 1);
  bodyRange.setBackground("#ede9fe").setHorizontalAlignment("center");
  bodyRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS).setAllowInvalid(true).build()
  );
}

// ---- 顧客管理シートに不足している案件列を同期追加（メニュー実行）----
function syncCustomerManagementCases() {
  const css = getCustomerManagementSS();
  if (!css) {
    SpreadsheetApp.getUi().alert("顧客管理シートが見つかりません。先に「顧客管理シートを作成」を実行してください。");
    return;
  }
  const CASE_NAMES = getJishaForms().map(f => f.displayName);
  if (CASE_NAMES.length === 0) {
    SpreadsheetApp.getUi().alert("自社フォーム（案件）が見つかりませんでした。");
    return;
  }

  let addedTotal = 0;
  const details  = [];
  css.getSheets().forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h));
    if (headers[0] !== "顧客名") return; // 営業マンシートのみ対象

    const missing = CASE_NAMES.filter(name => headers.indexOf(name) === -1);
    if (missing.length === 0) return;

    let col = lastCol;
    missing.forEach(name => { col += 1; addCaseColumnToSheet_(sheet, col, name); });
    addedTotal += missing.length;
    details.push(sheet.getName() + ": +" + missing.length + "列（" + missing.join("、") + "）");
  });

  SpreadsheetApp.getUi().alert(
    addedTotal === 0
      ? "すべての案件列は最新です。追加はありません。"
      : "案件列を同期しました。\n追加列数: " + addedTotal + "\n\n" + details.join("\n")
  );
}

// ---- 既存フォーム回答データを顧客管理SSに一括インポート ----
function importExistingToCustomerSheet() {
  const css = getCustomerManagementSS();
  if (!css) {
    SpreadsheetApp.getUi().alert("顧客管理シートが見つかりません。先に「顧客管理シートを作成」を実行してください。");
    return;
  }

  const ss        = getOrCreateSpreadsheet();
  let importCount = 0;

  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    // 代理店フォームは除外（読み取り専用チェック）
    const vals = sheet.getDataRange().getValues();
    let agencyCode = AGENCY_DEFAULT;
    for (const row of vals) {
      if (String(row[0]) === AGENCY_KEY) {
        const code = String(row[1] || "").trim();
        if (code) agencyCode = code;
        break;
      }
    }
    if (agencyCode !== AGENCY_DEFAULT) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const nameIdx = headers.indexOf("お名前");
    const refIdx  = headers.indexOf("紹介者名");
    if (nameIdx < 0 || refIdx < 0) return;

    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
      .forEach(row => {
        const customerName = String(row[nameIdx] || "").trim();
        const salesName    = String(row[refIdx]  || "").trim();
        if (!customerName || !salesName) return;
        upsertCustomerRow(salesName, customerName, formCode);
        importCount++;
      });
  });

  SpreadsheetApp.getUi().alert("インポート完了！\n処理件数: " + importCount + "件");
}

// ---- 150件クエスト 進捗レポート（毎朝8時自動送信）----
function quest150Report() {
  const now    = new Date();
  const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p      = n => String(n).padStart(2, "0");
  const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());

  if (todayJst > QUEST150_END_STR) return;

  const todayMs  = new Date(todayJst.replace(/\//g, "-")).getTime();
  const endMs    = new Date(QUEST150_END_STR.replace(/\//g, "-")).getTime();
  const daysLeft = Math.max(1, Math.round((endMs - todayMs) / 86400000) + 1);

  const ss = getOrCreateSpreadsheet();
  let totalCount = 0;
  const salesCounts = {};
  QUEST150_SALESPEOPLE.forEach(s => { salesCounts[s] = 0; });

  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode) return;
    // 代理店フォームは除外（読み取り専用チェック）
    const vals = sheet.getDataRange().getValues();
    let agencyCode = AGENCY_DEFAULT;
    for (const row of vals) {
      if (String(row[0]) === AGENCY_KEY) {
        const code = String(row[1] || "").trim();
        if (code) agencyCode = code;
        break;
      }
    }
    if (agencyCode !== AGENCY_DEFAULT) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx   = headers.indexOf("受信日時");
    const refIdx  = headers.indexOf("紹介者名");
    if (rtIdx < 0) return;

    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
      .forEach(row => {
        const dateStr = toJSTDateStr(row[rtIdx]);
        if (!dateStr || !dateStr.startsWith(QUEST150_MONTH)) return;
        totalCount++;
        if (refIdx >= 0) {
          const ref = String(row[refIdx] || "").trim();
          for (const sales of QUEST150_SALESPEOPLE) {
            if (normalizeName(ref) === normalizeName(sales)) {
              salesCounts[sales]++;
              break;
            }
          }
        }
      });
  });

  const weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  const dow = weekDays[jstNow.getUTCDay()];
  const lines = [];

  lines.push("⚔️【150件クエスト】進捗レポート");
  lines.push("📅 " + (jstNow.getUTCMonth()+1) + "/" + jstNow.getUTCDate() + "（" + dow + "） 残り" + daysLeft + "日");
  lines.push("");

  // 全体進捗
  const totalRemain = Math.max(0, QUEST150_TARGET_TOTAL - totalCount);
  const totalFilled = Math.min(10, Math.round(totalCount / QUEST150_TARGET_TOTAL * 10));
  const totalBar    = "🟩".repeat(totalFilled) + "🟥".repeat(10 - totalFilled);
  const totalPace   = totalRemain > 0
    ? "あと" + totalRemain + "件（1日" + Math.ceil(totalRemain / daysLeft) + "件ペース）"
    : "COMPLETE！";

  lines.push("📊 全体進捗");
  lines.push("　" + totalCount + " / " + QUEST150_TARGET_TOTAL + "件　" + totalPace);
  lines.push("　" + totalBar);
  lines.push("");

  // 個人ノルマ
  lines.push("━━━━━━━━━━━━━━");
  lines.push("🎯 個人ノルマ（各50件）");
  lines.push("");

  QUEST150_SALESPEOPLE.forEach(sales => {
    const count  = salesCounts[sales] || 0;
    const remain = Math.max(0, QUEST150_TARGET_INDIV - count);
    const pct    = count / QUEST150_TARGET_INDIV;
    const filled = Math.min(10, Math.round(pct * 10));
    const bar    = "🟩".repeat(filled) + "🟥".repeat(10 - filled);
    const pace   = remain > 0
      ? "あと" + remain + "件（1日" + Math.ceil(remain / daysLeft) + "件ペース）"
      : "COMPLETE！";

    let icon;
    if (count >= QUEST150_TARGET_INDIV) icon = "✅";
    else if (pct >= 0.8) icon = "🔥";
    else if (pct >= 0.6) icon = "⚡";
    else if (pct >= 0.3) icon = "🌱";
    else if (count > 0)  icon = "🚨";
    else                 icon = "💀";

    lines.push(icon + " " + sales);
    lines.push("　" + count + " / " + QUEST150_TARGET_INDIV + "件　" + pace);
    lines.push("　" + bar);
    lines.push("");
  });

  // 全体コメント
  const overallPct = totalCount / QUEST150_TARGET_TOTAL;
  let comment;
  if (totalCount >= QUEST150_TARGET_TOTAL) comment = "🎊 全員クリア！伝説の営業チームだ！";
  else if (overallPct >= 0.8) comment = "💬 もうすぐだ！全力ラストスパート！";
  else if (overallPct >= 0.6) comment = "💬 いい調子！このまま突き進め！";
  else if (overallPct >= 0.3) comment = "💬 加速しろ！まだ十分間に合う！";
  else if (totalCount > 0)    comment = "💬 ギアを上げろ！総力戦だ！";
  else                        comment = "💬 DANGER！今すぐ動け！";
  lines.push(comment);

  const message = lines.join("\n");
  notifyLineGroup(message);
  Logger.log(message);
}

// ---- 150件クエスト トリガー設置（毎朝8時）----
function ensureQuest150Trigger() {
  try {
    const now    = new Date();
    const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    const p      = n => String(n).padStart(2, "0");
    const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());
    if (todayJst > QUEST150_END_STR) return;
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction() === "quest150Report") return;
    }
    ScriptApp.newTrigger("quest150Report")
      .timeBased().everyDays(1).atHour(8).create();
    Logger.log("quest150Report トリガーを設置しました（毎朝8時）");
  } catch (e) {
    Logger.log("quest150Report トリガー設置失敗: " + e);
  }
}

// =============================================
// 緊急クエスト（2026/07/27〜07/31・毎日13時/20時）
// =============================================

// ---- 進捗レポート送信 ----
// トリガーから呼ばれると第1引数はイベントオブジェクト（force扱いにならない）
function emergencyQuestReport(force) {
  const now    = new Date();
  const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p      = n => String(n).padStart(2, "0");
  const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());

  // 期限後はトリガーを自動撤去して終了
  if (todayJst > EMERGENCY_END_STR) {
    removeEmergencyQuestTriggers_();
    return null;
  }

  // 二重送信ガード（トリガーと手動送信の重複防止・50分以内はスキップ）
  const props = PropertiesService.getScriptProperties();
  if (force !== true) {
    const last = Number(props.getProperty("EMERGENCY_QUEST_LAST_SENT") || 0);
    if (new Date().getTime() - last < 50 * 60 * 1000) return null;
  }

  const message = buildEmergencyQuestMessage_();
  notifyLineGroup(message);
  props.setProperty("EMERGENCY_QUEST_LAST_SENT", String(new Date().getTime()));
  Logger.log(message);
  return message;
}

function emergencyQuestReportTest() {
  const msg = emergencyQuestReport(true);
  SpreadsheetApp.getUi().alert("送信しました\n\n" + String(msg || "").substring(0, 800));
}

// ---- レポート本文組み立て（送信なし・プレビュー可能） ----
function buildEmergencyQuestMessage_() {
  const now    = new Date();
  const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p      = n => String(n).padStart(2, "0");
  const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());

  const todayMs  = new Date(todayJst.replace(/\//g, "-")).getTime();
  const endMs    = new Date(EMERGENCY_END_STR.replace(/\//g, "-")).getTime();
  const daysLeft = Math.max(1, Math.round((endMs - todayMs) / 86400000) + 1);

  const ss = getOrCreateSpreadsheet();
  const formCounts = {};
  EMERGENCY_FORMS.forEach(fc => { formCounts[fc] = 0; });
  let totalCount = 0;
  let iwamotoCount = 0;
  const startMs  = emergencyToEpochMs_(EMERGENCY_START_AT);
  const endLimit = emergencyToEpochMs_(EMERGENCY_END_STR + " 23:59:59");

  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const formCode = getFormCodeFromSheet(sheet);
    if (!formCode || !EMERGENCY_FORMS.includes(formCode)) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx   = headers.indexOf("受信日時");
    const refIdx  = headers.indexOf("紹介者名");
    if (rtIdx < 0) return;
    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
      .forEach(row => {
        const t = emergencyToEpochMs_(row[rtIdx]);
        if (isNaN(t) || t < startMs || t > endLimit) return;
        formCounts[formCode]++;
        totalCount++;
        if (refIdx >= 0 && normalizeName(String(row[refIdx] || "").trim()) === normalizeName(EMERGENCY_IWAMOTO)) {
          iwamotoCount++;
        }
      });
  });

  const othersCount = totalCount - iwamotoCount;

  const weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  const dow      = weekDays[jstNow.getUTCDay()];
  const lines    = [];

  lines.push("🚨【緊急クエスト】進捗レポート");
  lines.push("📅 " + (jstNow.getUTCMonth()+1) + "/" + jstNow.getUTCDate() + "（" + dow + "）" +
             p(jstNow.getUTCHours()) + ":" + p(jstNow.getUTCMinutes()) + "時点　残り" + daysLeft + "日");
  lines.push("");

  // 全体進捗
  const totalRemain = Math.max(0, EMERGENCY_TARGET_TOTAL - totalCount);
  const totalFilled = Math.min(10, Math.round(totalCount / EMERGENCY_TARGET_TOTAL * 10));
  const totalBar    = "🟩".repeat(totalFilled) + "🟥".repeat(10 - totalFilled);
  const totalPace   = totalRemain > 0
    ? "あと" + totalRemain + "件（1日" + Math.ceil(totalRemain / daysLeft) + "件ペース）"
    : "COMPLETE！";
  lines.push("📊 全体進捗（5案件・7/27 10時〜）");
  lines.push("　" + totalCount + " / " + EMERGENCY_TARGET_TOTAL + "件　" + totalPace);
  lines.push("　" + totalBar);
  lines.push("");

  // 個人ノルマ（岩本 / その他）
  lines.push("━━━━━━━━━━━━━━");
  lines.push("🎯 個人ノルマ");
  lines.push("");
  [
    { label: EMERGENCY_IWAMOTO,   count: iwamotoCount, target: EMERGENCY_TARGET_IWAMOTO },
    { label: "その他メンバー計",  count: othersCount,  target: EMERGENCY_TARGET_OTHERS }
  ].forEach(entry => {
    const remain = Math.max(0, entry.target - entry.count);
    const pct    = entry.count / entry.target;
    const filled = Math.min(10, Math.round(pct * 10));
    const bar    = "🟩".repeat(filled) + "🟥".repeat(10 - filled);
    const pace   = remain > 0
      ? "あと" + remain + "件（1日" + Math.ceil(remain / daysLeft) + "件ペース）"
      : "COMPLETE！";
    let icon;
    if (entry.count >= entry.target) icon = "✅";
    else if (pct >= 0.8) icon = "🔥";
    else if (pct >= 0.6) icon = "⚡";
    else if (pct >= 0.3) icon = "🌱";
    else if (entry.count > 0) icon = "🚨";
    else icon = "💀";
    lines.push(icon + " " + entry.label);
    lines.push("　" + entry.count + " / " + entry.target + "件　" + pace);
    lines.push("　" + bar);
    lines.push("");
  });

  // 案件別内訳
  lines.push("━━━━━━━━━━━━━━");
  lines.push("📋 案件別（7/27 10時〜）");
  EMERGENCY_FORMS.forEach(fc => {
    const name = EMERGENCY_FORM_NAMES[fc] || fc;
    lines.push("・" + name + "：" + (formCounts[fc] || 0) + "件");
    lines.push("　" + FORM_BASE_URL + "?form=" + fc);
  });
  lines.push("");

  // 全体コメント
  const overallPct = totalCount / EMERGENCY_TARGET_TOTAL;
  let comment;
  if (totalCount >= EMERGENCY_TARGET_TOTAL) comment = "🎊 目標達成！伝説の営業チームだ！";
  else if (overallPct >= 0.8) comment = "💬 もうすぐだ！全力ラストスパート！";
  else if (overallPct >= 0.6) comment = "💬 いい調子！このまま突き進め！";
  else if (overallPct >= 0.3) comment = "💬 加速しろ！まだ十分間に合う！";
  else if (totalCount > 0)    comment = "💬 ギアを上げろ！総力戦だ！";
  else                        comment = "💬 DANGER！今すぐ動け！";
  lines.push(comment);

  return lines.join("\n");
}

// ---- 受信日時をエポックmsへ（Dateセル / "yyyy/MM/dd HH:mm:ss"文字列(JST)の両対応） ----
function emergencyToEpochMs_(value) {
  if (!value) return NaN;
  if (value instanceof Date) return value.getTime();
  const m = String(value).trim()
    .match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (!m) return NaN;
  // JST文字列なので9時間引いてUTCエポックへ
  return Date.UTC(+m[1], +m[2] - 1, +m[3], (+(m[4] || 0)) - 9, +(m[5] || 0), +(m[6] || 0));
}

// ---- トリガー設置（毎日13時・20時） ----
function ensureEmergencyQuestTriggers() {
  try {
    const now    = new Date();
    const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    const p      = n => String(n).padStart(2, "0");
    const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());

    const existing = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === "emergencyQuestReport");

    if (todayJst > EMERGENCY_END_STR) {
      existing.forEach(t => ScriptApp.deleteTrigger(t));
      return { removed: existing.length, created: 0 };
    }

    // 2本（13時/20時）揃っていればそのまま。それ以外は作り直す
    if (existing.length === 2) return { removed: 0, created: 0, existing: 2 };
    existing.forEach(t => ScriptApp.deleteTrigger(t));
    [13, 20].forEach(hour => {
      ScriptApp.newTrigger("emergencyQuestReport")
        .timeBased().everyDays(1).atHour(hour).create();
    });
    Logger.log("emergencyQuestReport トリガーを設置しました（毎日13時・20時）");
    return { removed: existing.length, created: 2 };
  } catch (e) {
    Logger.log("emergencyQuestReport トリガー設置失敗: " + e);
    return { error: String(e) };
  }
}

function removeEmergencyQuestTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "emergencyQuestReport")
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// =============================================
// 特別緊急クエスト（ノムコム・2026/07/27〜07/31・毎日13時/20時）
// =============================================

// ---- 進捗レポート送信 ----
// トリガーから呼ばれると第1引数はイベントオブジェクト（force扱いにならない）
function specialQuestReport(force) {
  const now    = new Date();
  const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p      = n => String(n).padStart(2, "0");
  const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());

  // 期限後はトリガーを自動撤去して終了
  if (todayJst > SPECIAL_QUEST_END_STR) {
    removeSpecialQuestTriggers_();
    return null;
  }

  // 二重送信ガード（トリガーと手動送信の重複防止・50分以内はスキップ）
  const props = PropertiesService.getScriptProperties();
  if (force !== true) {
    const last = Number(props.getProperty("SPECIAL_QUEST_LAST_SENT") || 0);
    if (new Date().getTime() - last < 50 * 60 * 1000) return null;
  }

  const message = buildSpecialQuestMessage_();
  notifyLineGroup(message);
  props.setProperty("SPECIAL_QUEST_LAST_SENT", String(new Date().getTime()));
  Logger.log(message);
  return message;
}

function specialQuestReportTest() {
  const msg = specialQuestReport(true);
  SpreadsheetApp.getUi().alert("送信しました\n\n" + String(msg || "").substring(0, 800));
}

// ---- レポート本文組み立て（送信なし・プレビュー可能） ----
function buildSpecialQuestMessage_() {
  const now    = new Date();
  const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p      = n => String(n).padStart(2, "0");
  const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());

  const todayMs  = new Date(todayJst.replace(/\//g, "-")).getTime();
  const endMs    = new Date(SPECIAL_QUEST_END_STR.replace(/\//g, "-")).getTime();
  const daysLeft = Math.max(1, Math.round((endMs - todayMs) / 86400000) + 1);

  const ss       = getOrCreateSpreadsheet();
  const aliasMap = repStatusRepAliasMap_();
  const counts   = {};
  SPECIAL_QUEST_MEMBERS.forEach(m => { counts[m.name] = 0; });
  let totalCount = 0;
  let namedCount = 0;
  const startMs  = emergencyToEpochMs_(SPECIAL_QUEST_START_AT);
  const endLimit = emergencyToEpochMs_(SPECIAL_QUEST_END_STR + " 23:59:59");

  const sheet = getConfigSheetByCode(ss, SPECIAL_QUEST_FORM);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow >= 2 && lastCol >= ANSWER_START_COL) {
      const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
      const rtIdx   = headers.indexOf("受信日時");
      const refIdx  = headers.indexOf("紹介者名");
      if (rtIdx >= 0) {
        sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
          .forEach(row => {
            const t = emergencyToEpochMs_(row[rtIdx]);
            if (isNaN(t) || t < startMs || t > endLimit) return;
            totalCount++;
            if (refIdx < 0) return;
            // 「岩本」「松田恵美（岩本拓也）」等の表記ゆれも正規名へ寄せる
            const canon = resolveRepCanonical_(String(row[refIdx] || "").trim(), aliasMap);
            if (canon && counts.hasOwnProperty(canon)) {
              counts[canon]++;
              namedCount++;
            }
          });
      }
    }
  }

  const othersCount = Math.max(0, totalCount - namedCount);

  const weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  const dow      = weekDays[jstNow.getUTCDay()];
  const lines    = [];

  lines.push("⚡【特別緊急クエスト】" + SPECIAL_QUEST_FORM_NAME);
  lines.push("📅 " + (jstNow.getUTCMonth()+1) + "/" + jstNow.getUTCDate() + "（" + dow + "）" +
             p(jstNow.getUTCHours()) + ":" + p(jstNow.getUTCMinutes()) + "時点　残り" + daysLeft + "日");
  lines.push("");

  // 全体進捗
  const totalRemain = Math.max(0, SPECIAL_QUEST_TARGET_TOTAL - totalCount);
  const totalFilled = Math.min(10, Math.round(totalCount / SPECIAL_QUEST_TARGET_TOTAL * 10));
  const totalBar    = "🟩".repeat(totalFilled) + "🟥".repeat(10 - totalFilled);
  const totalPace   = totalRemain > 0
    ? "あと" + totalRemain + "件（1日" + Math.ceil(totalRemain / daysLeft) + "件ペース）"
    : "COMPLETE！";
  lines.push("📊 全体進捗（7/27〜7/31）");
  lines.push("　" + totalCount + " / " + SPECIAL_QUEST_TARGET_TOTAL + "件　" + totalPace);
  lines.push("　" + totalBar);
  lines.push("");

  // 個人ノルマ
  lines.push("━━━━━━━━━━━━━━");
  lines.push("🎯 個人ノルマ");
  lines.push("");
  SPECIAL_QUEST_MEMBERS.forEach(entry => {
    const count  = counts[entry.name] || 0;
    const remain = Math.max(0, entry.target - count);
    const pct    = count / entry.target;
    const filled = Math.min(10, Math.round(pct * 10));
    const bar    = "🟩".repeat(filled) + "🟥".repeat(10 - filled);
    const pace   = remain > 0
      ? "あと" + remain + "件（1日" + Math.ceil(remain / daysLeft) + "件ペース）"
      : "COMPLETE！";
    let icon;
    if (count >= entry.target) icon = "✅";
    else if (pct >= 0.8) icon = "🔥";
    else if (pct >= 0.6) icon = "⚡";
    else if (pct >= 0.3) icon = "🌱";
    else if (count > 0)  icon = "🚨";
    else icon = "💀";
    lines.push(icon + " " + entry.name);
    lines.push("　" + count + " / " + entry.target + "件　" + pace);
    lines.push("　" + bar);
    lines.push("");
  });
  if (othersCount > 0) {
    lines.push("（その他メンバー：" + othersCount + "件）");
    lines.push("");
  }

  // エントリー先
  lines.push("━━━━━━━━━━━━━━");
  lines.push("📋 " + SPECIAL_QUEST_FORM_NAME + " エントリーはこちら");
  lines.push("　" + FORM_BASE_URL + "?form=" + SPECIAL_QUEST_FORM);
  lines.push("");

  // 全体コメント
  const overallPct = totalCount / SPECIAL_QUEST_TARGET_TOTAL;
  let comment;
  if (totalCount >= SPECIAL_QUEST_TARGET_TOTAL) comment = "🎊 30件達成！特別クエストCLEAR！";
  else if (overallPct >= 0.8) comment = "💬 ゴールは目前！最後まで走り切れ！";
  else if (overallPct >= 0.6) comment = "💬 いいペース！このまま押し切れ！";
  else if (overallPct >= 0.3) comment = "💬 折り返しへ！ノムコムに集中！";
  else if (totalCount > 0)    comment = "💬 スタートダッシュだ！ここから伸ばそう！";
  else                        comment = "💬 まずは1件！ノムコムを動かせ！";
  lines.push(comment);

  return lines.join("\n");
}

// ---- トリガー設置（毎日13時・20時） ----
function ensureSpecialQuestTriggers() {
  try {
    const now    = new Date();
    const jstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    const p      = n => String(n).padStart(2, "0");
    const todayJst = jstNow.getUTCFullYear() + "/" + p(jstNow.getUTCMonth()+1) + "/" + p(jstNow.getUTCDate());

    const existing = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === "specialQuestReport");

    if (todayJst > SPECIAL_QUEST_END_STR) {
      existing.forEach(t => ScriptApp.deleteTrigger(t));
      return { removed: existing.length, created: 0 };
    }

    // 2本（13時/20時）揃っていればそのまま。それ以外は作り直す
    if (existing.length === 2) return { removed: 0, created: 0, existing: 2 };
    existing.forEach(t => ScriptApp.deleteTrigger(t));
    [13, 20].forEach(hour => {
      ScriptApp.newTrigger("specialQuestReport")
        .timeBased().everyDays(1).atHour(hour).create();
    });
    Logger.log("specialQuestReport トリガーを設置しました（毎日13時・20時）");
    return { removed: existing.length, created: 2 };
  } catch (e) {
    Logger.log("specialQuestReport トリガー設置失敗: " + e);
    return { error: String(e) };
  }
}

function removeSpecialQuestTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "specialQuestReport")
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ---- 管理ルート（quest_admin）はクエスト終了(2026/07/31)に伴い撤去した ----
// handleEmergencyAdmin_ も到達不能になったため削除した。再開する場合は
// キーをソースに書かず ScriptProperties から読むこと（このリポジトリは public）。

// ---- 既存フォーム回答を広告主成果管理シートへ一括インポート ----
// 月ごとにバッファしてまとめて書き込む。同じシートに重複して実行しないこと。
function importExistingToAdvertiserSheet() {
  const result = importExistingToAdvertiserSheetCore_();
  SpreadsheetApp.getUi().alert(
    "インポート完了！\n追記: " + result.importCount + "件" +
    "\n承認: " + result.approvedCount + "件" +
    "\nトラッキング漏れ: " + result.trackingMissingCount + "件" +
    (result.skipCount > 0 ? "\nスキップ（シートなし）: " + result.skipCount + "件" : "")
  );
}

// ---- 月を指定して広告主成果管理シートを再生成（メニュー用）----
// 指定しなかった月のシートには一切触れないので、提出済みの過去月を巻き込まない。
function importAdvertiserMonths() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "広告主シートを月指定で再生成",
    "対象月をYYYYMM形式でカンマ区切り入力（例: 202607,202608）\n" +
    "指定した月だけを作り直します。該当データが0件の月は空シートを用意します。",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  let result;
  try {
    result = importAdvertiserMonthsCore_(res.getResponseText());
  } catch (err) {
    ui.alert("エラー: " + err);
    return;
  }
  const lines = result.months.map(m => {
    const c = result.monthCounts[m];
    return "  " + m + ": " + c.rows + "件（承認 " + c.approved +
           " / トラッキング漏れ " + c.trackingMissing + "）";
  });
  ui.alert(
    "再生成しました\n" + lines.join("\n") +
    (result.created.length ? "\n\n新規作成: " + result.created.join(", ") : "")
  );
}

// ---- "202607,202608" や ["2026/07"] を昇順・重複なしのYYYYMM配列へ正規化 ----
function normalizeAdvertiserMonths_(months) {
  const list = Array.isArray(months) ? months : String(months || "").split(",");
  const seen = {};
  const out  = [];
  list.forEach(m => {
    const v = String(m || "").replace(/[^\d]/g, "");
    if (!/^\d{6}$/.test(v)) return;
    const mm = Number(v.substring(4, 6));
    if (mm < 1 || mm > 12) return;
    if (seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out.sort();
}

// ---- メインSSの自社フォーム回答を月別(YYYYMM)に集める ----
// 全月インポートと月指定インポートで共有する収集処理。広告主シートには触れない。
function collectAdvertiserRowsByMonth_() {
  const mainSS = getOrCreateSpreadsheet();
  const buffer = {}; // yyyymm → [[受信日時, 広告名, お名前, 紹介者名, スクショURL, トラッキング漏れ, 承認], ...]
  const stats  = {}; // yyyymm → { approved, trackingMissing, ignored }

  mainSS.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    // 代理店フォームは除外
    const vals = sheet.getDataRange().getValues();
    let agencyCode = AGENCY_DEFAULT;
    for (const row of vals) {
      if (String(row[0]) === AGENCY_KEY) {
        const code = String(row[1] || "").trim();
        if (code) agencyCode = code;
        break;
      }
    }
    if (agencyCode !== AGENCY_DEFAULT) return;

    const formCode = getFormCodeFromSheet(sheet);
    const formDisplayName = getFormDisplayName(sheet, formCode || "");

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const headers    = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx      = headers.indexOf("受信日時");
    const nameIdx    = headers.indexOf("お名前");
    const refIdx     = headers.indexOf("紹介者名");
    const ssUrlIdx   = headers.indexOf("スクショURL") >= 0
      ? headers.indexOf("スクショURL")
      : headers.indexOf("スクリーンショットURL");
    const shoninIdx  = headers.indexOf("承認");
    if (rtIdx < 0) return;

    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
      .forEach(row => {
        const dateStr = toJSTDateStr(row[rtIdx]);
        if (!dateStr) return;
        const yyyymm = dateStr.substring(0, 7).replace("/", ""); // "2026/05" → "202605"
        if (!buffer[yyyymm]) {
          buffer[yyyymm] = [];
          stats[yyyymm]  = { approved: 0, trackingMissing: 0, ignored: 0 };
        }
        const approvalFlags = shoninIdx >= 0
          ? getAdvertiserApprovalFlags(row[shoninIdx])
          : { approved: false, trackingMissing: false };
        if (approvalFlags.approved) stats[yyyymm].approved++;
        else if (approvalFlags.trackingMissing) stats[yyyymm].trackingMissing++;
        else stats[yyyymm].ignored++;
        buffer[yyyymm].push([
          String(row[rtIdx]                      || ""),
          formDisplayName,
          nameIdx  >= 0 ? String(row[nameIdx]  || "").trim() : "",
          refIdx   >= 0 ? String(row[refIdx]   || "").trim() : "",
          ssUrlIdx >= 0 ? String(row[ssUrlIdx] || "").trim() : "",
          approvalFlags.trackingMissing,
          approvalFlags.approved
        ]);
      });
  });

  return { buffer: buffer, stats: stats };
}

// ---- 1か月分を広告主シートへ書き戻す（行3以降を消してから書き直す）----
// シートが無ければ作る。rows が空でもシートだけは用意する。
function writeAdvertiserMonthSheet_(advertiserSS, yyyymm, rows) {
  let sheet   = advertiserSS.getSheetByName(yyyymm);
  let created = false;
  if (!sheet) {
    sheet   = createAdvertiserMonthSheet(advertiserSS, yyyymm);
    created = true;
  }
  if (!sheet) return { written: 0, created: false, skipped: true };

  // 既存データ行をクリア（1行目タイトル・2行目ヘッダーは保持し3行目から消去）
  const clearLastRow = sheet.getLastRow();
  const clearLastCol = Math.max(sheet.getLastColumn(), 7);
  if (clearLastRow >= ADVERTISER_DATA_START_ROW) {
    sheet.getRange(ADVERTISER_DATA_START_ROW, 1,
                   clearLastRow - ADVERTISER_DATA_START_ROW + 1, clearLastCol).clearContent();
  }
  if (rows.length) {
    // 直前に行3以降を消しているので開始行は常に3で確定（クリア直後の getLastRow は当てにしない）
    sheet.getRange(ADVERTISER_DATA_START_ROW, 1, rows.length, 5).setValues(rows.map(r => r.slice(0, 5)));
    sheet.getRange(ADVERTISER_DATA_START_ROW, 6, rows.length, 2).setValues(rows.map(r => [r[5], r[6]]));
  }
  return { written: rows.length, created: created, skipped: false };
}

function importExistingToAdvertiserSheetCore_() {
  const advertiserSS = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  const collected    = collectAdvertiserRowsByMonth_();

  let importCount = 0;
  let skipCount = 0;
  let approvedCount = 0;
  let trackingMissingCount = 0;
  let ignoredApprovalCount = 0;
  const monthCounts = {};

  Object.keys(collected.buffer).sort().forEach(yyyymm => {
    const rows = collected.buffer[yyyymm];
    const s    = collected.stats[yyyymm];
    const res  = writeAdvertiserMonthSheet_(advertiserSS, yyyymm, rows);
    if (res.skipped) {
      skipCount += rows.length;
      Logger.log("広告主シート「" + yyyymm + "」なし → " + rows.length + "件スキップ");
      return;
    }
    monthCounts[yyyymm]   = rows.length;
    importCount          += rows.length;
    approvedCount        += s.approved;
    trackingMissingCount += s.trackingMissing;
    ignoredApprovalCount += s.ignored;
    Logger.log("広告主シート「" + yyyymm + "」に " + rows.length + "件追記");
  });

  return {
    importCount: importCount,
    skipCount: skipCount,
    approvedCount: approvedCount,
    trackingMissingCount: trackingMissingCount,
    ignoredApprovalCount: ignoredApprovalCount,
    monthCounts: monthCounts
  };
}

// ---- 月を指定して広告主成果管理シートを再生成 ----
// months は "202607,202608" でも ["202607","202608"] でも可。
// 指定月のシートが無ければ作り、該当データが0件でも空シートとして用意する
// （writeToAdvertiserSheet はシートが無い月の成果を捨てるため、月初までに存在させる）。
function importAdvertiserMonthsCore_(months) {
  const targets = normalizeAdvertiserMonths_(months);
  if (!targets.length) throw new Error("対象月をYYYYMM形式で指定してください（例: 202607,202608）");

  const advertiserSS = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  const collected    = collectAdvertiserRowsByMonth_();

  const monthCounts = {};
  const created     = [];
  let importCount = 0;
  let approvedCount = 0;
  let trackingMissingCount = 0;

  targets.forEach(yyyymm => {
    const rows = collected.buffer[yyyymm] || [];
    const s    = collected.stats[yyyymm]  || { approved: 0, trackingMissing: 0, ignored: 0 };
    const res  = writeAdvertiserMonthSheet_(advertiserSS, yyyymm, rows);
    if (res.skipped) throw new Error("広告主シート「" + yyyymm + "」を作成できませんでした");
    if (res.created) created.push(yyyymm);

    monthCounts[yyyymm] = {
      rows: rows.length,
      approved: s.approved,
      trackingMissing: s.trackingMissing,
      ignored: s.ignored
    };
    importCount          += rows.length;
    approvedCount        += s.approved;
    trackingMissingCount += s.trackingMissing;
    Logger.log("広告主シート「" + yyyymm + "」を再生成: " + rows.length + "件" +
               (res.created ? "（新規作成）" : ""));
  });

  return {
    months: targets,
    created: created,
    importCount: importCount,
    approvedCount: approvedCount,
    trackingMissingCount: trackingMissingCount,
    monthCounts: monthCounts
  };
}

// ---- 書き込まずに件数と現状だけ返す（再生成の事前確認用）----
function previewAdvertiserMonths_(months) {
  const targets      = normalizeAdvertiserMonths_(months);
  const advertiserSS = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  const collected    = collectAdvertiserRowsByMonth_();

  const monthCounts = {};
  targets.forEach(yyyymm => {
    const rows  = collected.buffer[yyyymm] || [];
    const s     = collected.stats[yyyymm]  || { approved: 0, trackingMissing: 0, ignored: 0 };
    const sheet = advertiserSS.getSheetByName(yyyymm);
    monthCounts[yyyymm] = {
      rows: rows.length,
      approved: s.approved,
      trackingMissing: s.trackingMissing,
      sheetExists: !!sheet,
      // getLastRow は書式やチェックボックスだけの行も拾うので、A列に値がある行を実データとして数える
      current: sheet ? countAdvertiserSheetRows_(sheet) : null
    };
  });

  return {
    months: targets,
    monthCounts: monthCounts,
    monthsInMainSS: Object.keys(collected.buffer).sort(),
    advertiserSheets: advertiserSS.getSheets()
      .map(sh => sh.getName()).filter(n => /^\d{6}$/.test(n))
  };
}

// ---- 広告主シート1枚の実データ行数とチェック数を数える（PIIは返さない）----
function countAdvertiserSheetRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < ADVERTISER_DATA_START_ROW) {
    return { lastRow: lastRow, rows: 0, approved: 0, trackingMissing: 0 };
  }
  const n    = lastRow - ADVERTISER_DATA_START_ROW + 1;
  const vals = sheet.getRange(ADVERTISER_DATA_START_ROW, 1, n, 7).getValues();
  let rows = 0, approved = 0, trackingMissing = 0;
  vals.forEach(r => {
    if (String(r[0]).trim() === "") return;
    rows++;
    if (r[5] === true) trackingMissing++;
    if (r[6] === true) approved++;
  });
  return { lastRow: lastRow, rows: rows, approved: approved, trackingMissing: trackingMissing };
}

// ---- 広告主成果管理SSをまるごとDriveに複製してバックアップする ----
// 上書きを伴う再生成の前に必ず取る（2026-06-29の再インポート時と同じ運用）
function backupAdvertiserSpreadsheet_(label) {
  const src  = DriveApp.getFileById(ADVERTISER_SS_ID);
  const jst  = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
  const p    = n => String(n).padStart(2, "0");
  const stamp = jst.getUTCFullYear() + p(jst.getUTCMonth() + 1) + p(jst.getUTCDate()) +
                "-" + p(jst.getUTCHours()) + p(jst.getUTCMinutes());
  const name = src.getName() + " バックアップ " + stamp + (label ? " " + label : "");
  const copy = src.makeCopy(name);
  Logger.log("広告主SSをバックアップ: " + name + " / " + copy.getId());
  return { id: copy.getId(), name: name, url: "https://docs.google.com/spreadsheets/d/" + copy.getId() + "/edit" };
}

// ---- 当月と翌月の広告主シートを先回りして用意する ----
// writeToAdvertiserSheet はシートが無い月の成果を黙って捨てるため、月が変わる前に作っておく。
// （2026年7月は 202607 が無いまま月が進み、リアルタイム書き込みが1か月分丸ごと落ちた）
function ensureAdvertiserMonthSheets() {
  const ss  = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  const jst = new Date(new Date().getTime() + 9 * 60 * 60 * 1000);
  const p   = n => String(n).padStart(2, "0");

  const created = [];
  for (let i = 0; i <= 1; i++) { // 当月・翌月
    const d    = new Date(Date.UTC(jst.getUTCFullYear(), jst.getUTCMonth() + i, 1));
    const name = d.getUTCFullYear() + p(d.getUTCMonth() + 1);
    if (ss.getSheetByName(name)) continue;
    createAdvertiserMonthSheet(ss, name);
    created.push(name);
  }
  if (created.length) Logger.log("広告主シートを先回り作成: " + created.join(", "));
  return created;
}

// ---- 上記を毎月25日に走らせるトリガーを登録（重複登録しない）----
// 25日にしているのは、月初0時の書き込みとシート作成が競合しないよう余裕を持たせるため。
function ensureAdvertiserMonthTrigger() {
  const existing = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "ensureAdvertiserMonthSheets");
  if (existing.length) return { created: false, triggers: existing.length };
  ScriptApp.newTrigger("ensureAdvertiserMonthSheets")
    .timeBased().onMonthDay(25).atHour(3).create();
  return { created: true, triggers: 1 };
}

// ---- 広告主シート保守ルート（doGet から adv_admin キー付きで呼ばれる）----
// action=preview … 書き込まずに件数と現状を返す / action=rebuild … 月指定で再生成
// action=backup … SSまるごとDriveに複製 / action=ensuremonths … 当月・翌月シート＋月次トリガー
// rebuild に &backup=1 を付けると、書き込む直前に同じ実行でバックアップを取る
function handleAdvertiserAdmin_(action, months, backup) {
  const out = { action: action, months: months };
  try {
    if (action === "preview") {
      out.result = previewAdvertiserMonths_(months);
    } else if (action === "backup") {
      out.result = backupAdvertiserSpreadsheet_("");
    } else if (action === "rebuild") {
      if (backup) out.backup = backupAdvertiserSpreadsheet_("(" + normalizeAdvertiserMonths_(months).join("_") + "再生成前)");
      out.result = importAdvertiserMonthsCore_(months);
    } else if (action === "ensuremonths") {
      out.result = {
        created: ensureAdvertiserMonthSheets(),
        trigger: ensureAdvertiserMonthTrigger()
      };
    } else {
      out.error = "unknown action";
    }
  } catch (e) {
    out.error = String(e);
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- 広告主シートに月別シートを新規作成（既存シートをテンプレートにコピー）----
function createAdvertiserMonthSheet(advertiserSS, yyyymm) {
  const monthSheets = advertiserSS.getSheets()
    .filter(s => /^\d{6}$/.test(s.getName()))
    .sort((a, b) => a.getName().localeCompare(b.getName()));

  let newSheet;
  if (monthSheets.length > 0) {
    // 最新の月シートを書式テンプレートに使う（列構成を変えた場合、新しい月にも引き継がれる）
    const template = monthSheets[monthSheets.length - 1];
    newSheet = template.copyTo(advertiserSS);
    newSheet.setName(yyyymm);
    const lastRow = newSheet.getLastRow();
    if (lastRow > 2) newSheet.deleteRows(3, lastRow - 2);
  } else {
    newSheet = advertiserSS.insertSheet(yyyymm);
  }

  // 月別シートが昇順に並ぶよう位置を調整
  const allSheets = advertiserSS.getSheets(); // newSheet は末尾にある
  let insertPos = allSheets.length;
  for (let i = 0; i < allSheets.length - 1; i++) {
    const name = allSheets[i].getName();
    if (/^\d{6}$/.test(name) && name > yyyymm) {
      insertPos = i + 1; // 1-based: このシートの直前に挿入
      break;
    }
  }
  try {
    advertiserSS.setActiveSheet(newSheet);
    advertiserSS.moveActiveSheet(insertPos);
  } catch (e) {
    Logger.log("シート位置の移動エラー（無視）: " + e);
  }

  Logger.log("広告主シート「" + yyyymm + "」を新規作成しました");
  return newSheet;
}

// ---- 広告主シートのA列で実際にデータがある最終行の次を返す ----
// getLastRow()は書式だけの行も含むため、A列の値で実データ末尾を探す
function getAdvertiserNextRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 1;
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim() !== "") return i + 2;
  }
  return 1;
}

function getAdvertiserApprovalFlags(value) {
  const mark = String(value || "")
    .replace(/\uFE0F/g, "")
    .replace(/\s+/g, "")
    .trim();

  const approvedMarks = ["⭕", "○", "◯", "〇"];
  const trackingMissingMarks = ["❌", "❎", "✖", "✕", "×", "☓", "✗", "✘"];

  return {
    approved: approvedMarks.indexOf(mark) >= 0,
    trackingMissing: trackingMissingMarks.indexOf(mark) >= 0
  };
}

// ---- 広告主成果管理シートに行を追記 ----
// シート名は YYYYMM 形式（例: 202606）。対象シートが存在しない場合はログのみ。
function writeToAdvertiserSheet(receivedAt, formDisplayName, customerName, referrerName, screenshotUrl, approved, trackingMissing) {
  const ss   = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  const now  = new Date();
  const jst  = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p    = n => String(n).padStart(2, "0");
  const sheetName = jst.getUTCFullYear() + p(jst.getUTCMonth() + 1);

  // 月初にシートが無いと、その月の成果が丸ごと記録されないまま流れてしまうので自動で作る
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    try {
      sheet = createAdvertiserMonthSheet(ss, sheetName);
    } catch (e) {
      Logger.log("広告主シート「" + sheetName + "」の自動作成に失敗: " + e);
    }
  }
  if (!sheet) {
    Logger.log("広告主シート「" + sheetName + "」が見つかりません。スキップします。");
    return;
  }

  const nextRow = getAdvertiserNextRow(sheet);
  sheet.getRange(nextRow, 1, 1, 5).setValues([[receivedAt, formDisplayName, customerName, referrerName, screenshotUrl]]);
  sheet.getRange(nextRow, 6, 1, 2).setValues([[trackingMissing || false, approved || false]]);
  Logger.log("広告主シート「" + sheetName + "」に追記: " + customerName);
}

// =============================================
// 営業担当別 案件ステータス表を SS2 に生成
// 各セル = {申請月}月{状態}（例「5月承認」「6月申請」「5月非承認」）
// 正データ = C(アフィリエイト管理SS=メインSS)の 設定_ タブ（代理店/小中様は agencyCode で除外）
// 出力 = SS2 の8営業担当タブのみ（非破壊：2行目以降 clearContent → 再書き込み）
// 実行: Apps Scriptエディタから buildSalesRepStatusSheets を Run（clasp run 不可のため）
// =============================================
const REP_STATUS_SS2_ID  = "1TARRQ_hqnptRGeEu5WzH1hlVZ2jqlfGunsHRcxbTWsg";
const REP_STATUS_MAIN_ID  = "1JOMT_Uuoq3H6O9lZcAKuSpzC2ehUJ35k53mVvnZRwY0"; // C の ID ガード用
const REP_STATUS_SS1_ID   = "1aaiCIDQIkrp_Ado5aKua_PTEQq4jr1UWqpuLQXwpemI"; // 統合顧客管理（フォーム顧客管理）ブック

// 紹介者名(別名)→ 名簿の正規名(SS2タブ名)。normalizeName(別名) をキーに引く（漢字表記ゆれ吸収）
// 正規名の自己対応は JISHA_REFERRER_OPTIONS から自動生成。表記ゆれ・苗字のみだけ手動で足す。
function repStatusRepAliasMap_() {
  const variants = {
    "柳沢悠貴": ["柳澤悠貴", "柳沢", "柳澤", "橋沢悠貴", "橋澤悠貴", "橋沢", "橋澤"],
    "岩本拓也": ["岩本"],
    "菅原貴博": ["菅原"],
    "村井亮介": ["村井"],
    "大島雅史": ["大島"],
    "小椋裕也": ["小椋"],
    "細川貴弘": ["細川"],
    "藤森宣哉": ["藤森"],
    "江口裕人": ["江口"],
    "藤井勇大": ["藤井"]
  };
  const map = {};
  JISHA_REFERRER_OPTIONS.split(",").map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (canon) {
    map[normalizeName(canon)] = canon; // 正規名の自己対応
    (variants[canon] || []).forEach(function (alias) { map[normalizeName(alias)] = canon; });
  });
  return map;
}

// 紹介者名 → 8名の正規名。括弧内（「松田恵美（岩本拓也）」等）の担当名も救済する
function resolveRepCanonical_(rawRef, aliasMap) {
  let c = aliasMap[normalizeName(rawRef)];
  if (c) return c;
  const m = String(rawRef).match(/[（(]([^）)]+)[）)]/);
  if (m) { c = aliasMap[normalizeName(m[1])]; if (c) return c; }
  return null;
}

function buildSalesRepStatusSheets() {
  const mainSS = getOrCreateSpreadsheet();
  const mainId = mainSS.getId();
  if (mainId !== REP_STATUS_MAIN_ID) {
    const msg = "中止: メインSSのIDが想定外。実ID=" + mainId + " 名前=" + mainSS.getName() +
                " / 期待=" + REP_STATUS_MAIN_ID + "（誤ブックへの書込防止）";
    Logger.log(msg);
    throw new Error(msg);
  }

  const aliasMap = repStatusRepAliasMap_();
  const canonicalReps = JISHA_REFERRER_OPTIONS.split(",").map(function (s) { return s.trim(); });

  // agg: rep -> custKey -> { name, cases: { 案件表示名 -> {cell, status, rtKey} } }
  const agg = {};
  const targetForms = [];
  const excludedReferrers = {};
  let monthMissing = 0;

  mainSS.getSheets().forEach(function (sheet) {
    if (sheet.getName().indexOf(CONFIG_PREFIX) !== 0) return;
    const vals = sheet.getDataRange().getValues();
    // 代理店フォーム除外（agencyCode が house 以外は対象外）
    let agencyCode = AGENCY_DEFAULT;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === AGENCY_KEY) {
        const code = String(vals[i][1] || "").trim();
        if (code) agencyCode = code;
        break;
      }
    }
    if (agencyCode !== AGENCY_DEFAULT) return;

    const caseName = getFormDisplayName(sheet, getFormCodeFromSheet(sheet) || "");
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx = headers.indexOf("受信日時");
    const nameIdx = headers.indexOf("お名前");
    const refIdx = headers.indexOf("紹介者名");
    const shoninIdx = headers.indexOf("承認");
    if (rtIdx < 0 || nameIdx < 0 || refIdx < 0) return;
    targetForms.push(caseName);

    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues()
      .forEach(function (row) {
        const custName = String(row[nameIdx] || "").trim();
        const rawRef = String(row[refIdx] || "").trim();
        if (!custName || !rawRef) return;
        const canon = resolveRepCanonical_(rawRef, aliasMap);
        if (!canon) { excludedReferrers[rawRef] = (excludedReferrers[rawRef] || 0) + 1; return; }
        const dateStr = toJSTDateStr(row[rtIdx]); // "YYYY/MM/DD"
        if (!dateStr || dateStr.length < 7) { monthMissing++; return; }
        const month = parseInt(dateStr.substring(5, 7), 10);
        if (!month) { monthMissing++; return; }
        const flags = shoninIdx >= 0 ? getAdvertiserApprovalFlags(row[shoninIdx]) : { approved: false, trackingMissing: false };
        const status = flags.approved ? "承認" : (flags.trackingMissing ? "非承認" : "申請");
        const cell = month + "月" + status;
        const rtRaw = row[rtIdx];
        const rtKey = (rtRaw instanceof Date) ? rtRaw.getTime() : (Date.parse(String(rtRaw)) || 0);

        const custKey = normalizeName(custName);
        if (!agg[canon]) agg[canon] = {};
        if (!agg[canon][custKey]) agg[canon][custKey] = { name: custName, cases: {} };
        const ex = agg[canon][custKey].cases[caseName];
        if (!ex || rtKey >= ex.rtKey) {
          agg[canon][custKey].cases[caseName] = { cell: cell, status: status, rtKey: rtKey };
        }
      });
  });

  // ---- SS2 へ非破壊書き込み（8タブのみ）----
  const outSS = SpreadsheetApp.openById(REP_STATUS_SS2_ID);
  const perRepRows = {};
  const unmatchedCase = {};
  let cApproved = 0, cApplied = 0, cRejected = 0;

  canonicalReps.forEach(function (rep) {
    const tab = outSS.getSheetByName(rep);
    if (!tab) { Logger.log("SS2タブなし: " + rep + " → スキップ"); return; }
    // 入力規則（データ検証）を全面解除（{月}月{状態} が既存規則に違反するため）
    tab.getRange(1, 1, tab.getMaxRows(), tab.getMaxColumns()).clearDataValidations();
    const lastCol = tab.getLastColumn();
    const header = tab.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    let nameCol = header.indexOf("顧客名"); if (nameCol < 0) nameCol = 0;
    const caseCol = {};
    header.forEach(function (h, idx) { if (h && idx >= 3) caseCol[h] = idx; }); // 案件列=4列目(index3)以降

    const lastRow = tab.getLastRow();
    if (lastRow > 1) {
      tab.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      tab.getRange(2, 1, lastRow - 1, lastCol).setBackground("#ffffff"); // 古い色分けを消す
    }

    const custs = agg[rep] || {};
    const keys = Object.keys(custs).sort(function (a, b) {
      const na = custs[a].name, nb = custs[b].name;
      return na < nb ? -1 : (na > nb ? 1 : 0);
    });
    if (keys.length === 0) { tab.getRange(1, 1, 1, lastCol).setBorder(true, true, true, true, true, true); perRepRows[rep] = 0; return; }
    const bgRows = [];
    const out = keys.map(function (k) {
      const c = custs[k];
      const arr = [];
      const bg = [];
      for (let i = 0; i < lastCol; i++) { arr.push(""); bg.push("#ffffff"); }
      arr[nameCol] = c.name;
      Object.keys(c.cases).forEach(function (cn) {
        const cc = c.cases[cn];
        if (caseCol[cn] !== undefined) {
          arr[caseCol[cn]] = cc.cell;
          // 色分け: 承認=緑 / 非承認=赤 / 申請=黄
          bg[caseCol[cn]] = cc.status === "承認" ? "#d9ead3" : (cc.status === "非承認" ? "#f4cccc" : "#fff2cc");
          if (cc.status === "承認") cApproved++; else if (cc.status === "非承認") cRejected++; else cApplied++;
        } else {
          unmatchedCase[cn] = (unmatchedCase[cn] || 0) + 1;
        }
      });
      bgRows.push(bg);
      return arr;
    });
    tab.getRange(2, 1, out.length, lastCol).setValues(out);
    tab.getRange(2, 1, out.length, lastCol).setBackgrounds(bgRows);
    tab.getRange(1, 1, out.length + 1, lastCol).setBorder(true, true, true, true, true, true); // 枠線
    perRepRows[rep] = out.length;
  });

  Logger.log("=== buildSalesRepStatusSheets 結果 ===");
  Logger.log("C(メインSS) ID=" + mainId + " ガードOK");
  Logger.log("対象 設定_ タブ数=" + targetForms.length + " : " + targetForms.join(", "));
  Logger.log("担当別 出力行数: " + JSON.stringify(perRepRows));
  Logger.log("最終セル状態別: 承認=" + cApproved + " 申請=" + cApplied + " 非承認=" + cRejected);
  Logger.log("月不明で除外=" + monthMissing);
  Logger.log("8名に正規化できず除外した紹介者名: " + JSON.stringify(excludedReferrers));
  Logger.log("SS2ヘッダーに突合できなかった案件表示名: " + JSON.stringify(unmatchedCase));

  return {
    targetForms: targetForms, perRepRows: perRepRows,
    approved: cApproved, applied: cApplied, rejected: cRejected,
    monthMissing: monthMissing, excludedReferrers: excludedReferrers, unmatchedCase: unmatchedCase
  };
}

// =============================================
// 統合顧客管理（SS1）を営業担当別に分割 → SS1 に「総合_<担当>」タブを生成
// 統合の全列を保持し、案件セルを {月}月{状態}＋色分けに置換。担当はアフィリンク優先。
// =============================================
function buildIntegratedRepSheets() {
  const cSS = getOrCreateSpreadsheet();
  if (cSS.getId() !== REP_STATUS_MAIN_ID) {
    throw new Error("中止: メインSS(C)のID不一致 実ID=" + cSS.getId());
  }
  const aliasMap = repStatusRepAliasMap_();

  // ---- C から: custKey → {案件 → {cell,status,rtKey}}、custKey → {rep,rtKey}(アフィリンク担当) ----
  const byCust = {};
  const affRepByCust = {};
  cSS.getSheets().forEach(function (sheet) {
    if (sheet.getName().indexOf(CONFIG_PREFIX) !== 0) return;
    const vals = sheet.getDataRange().getValues();
    let agencyCode = AGENCY_DEFAULT;
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === AGENCY_KEY) { const code = String(vals[i][1] || "").trim(); if (code) agencyCode = code; break; }
    }
    if (agencyCode !== AGENCY_DEFAULT) return;
    const caseName = getFormDisplayName(sheet, getFormCodeFromSheet(sheet) || "");
    const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, lastCol - ANSWER_START_COL + 1).getValues()[0];
    const rtIdx = headers.indexOf("受信日時"), nameIdx = headers.indexOf("お名前"), refIdx = headers.indexOf("紹介者名"), shoninIdx = headers.indexOf("承認");
    if (rtIdx < 0 || nameIdx < 0) return;
    sheet.getRange(2, ANSWER_START_COL, lastRow - 1, lastCol - ANSWER_START_COL + 1).getValues().forEach(function (row) {
      const custName = String(row[nameIdx] || "").trim();
      if (!custName) return;
      const dateStr = toJSTDateStr(row[rtIdx]);
      if (!dateStr || dateStr.length < 7) return;
      const month = parseInt(dateStr.substring(5, 7), 10);
      if (!month) return;
      const flags = shoninIdx >= 0 ? getAdvertiserApprovalFlags(row[shoninIdx]) : { approved: false, trackingMissing: false };
      const status = flags.approved ? "承認" : (flags.trackingMissing ? "非承認" : "申請");
      const cell = month + "月" + status;
      const rtRaw = row[rtIdx];
      const rtKey = (rtRaw instanceof Date) ? rtRaw.getTime() : (Date.parse(String(rtRaw)) || 0);
      const custKey = normalizeName(custName);
      if (!byCust[custKey]) byCust[custKey] = {};
      const ex = byCust[custKey][caseName];
      if (!ex || rtKey >= ex.rtKey) byCust[custKey][caseName] = { cell: cell, status: status, rtKey: rtKey };
      const canon = refIdx >= 0 ? resolveRepCanonical_(String(row[refIdx] || "").trim(), aliasMap) : null;
      if (canon) { const e2 = affRepByCust[custKey]; if (!e2 || rtKey >= e2.rtKey) affRepByCust[custKey] = { rep: canon, rtKey: rtKey }; }
    });
  });
  const allCases = {};
  Object.keys(byCust).forEach(function (k) { Object.keys(byCust[k]).forEach(function (cn) { allCases[cn] = true; }); });

  // ---- SS1 統合顧客管理 を読む ----
  const ss1 = SpreadsheetApp.openById(REP_STATUS_SS1_ID);
  const src = ss1.getSheetByName("統合顧客管理");
  if (!src) throw new Error("統合顧客管理 タブが見つかりません");
  const data = src.getDataRange().getValues();
  if (data.length < 2) throw new Error("統合顧客管理 が空です");
  const header = data[0].map(function (h) { return String(h).trim(); });
  const w = header.length;
  const repColIdx = header.indexOf("営業担当");
  const affNameColIdx = header.indexOf("アフィリンク顧客名");
  const nameColIdx = header.indexOf("名前");
  const caseColIdxs = {};
  header.forEach(function (h, idx) { if (allCases[h]) caseColIdxs[h] = idx; });

  // ---- 担当別に振り分け ----
  const perRep = {};
  const excluded = {};
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row.join("")).trim() === "") continue;
    const affName = affNameColIdx >= 0 ? String(row[affNameColIdx] || "").trim() : "";
    const custKey = affName ? normalizeName(affName) : (nameColIdx >= 0 ? normalizeName(String(row[nameColIdx] || "")) : "");
    let rep = (custKey && affRepByCust[custKey]) ? affRepByCust[custKey].rep : null;
    if (!rep && repColIdx >= 0) {
      const raw = String(row[repColIdx] || "").trim();
      rep = resolveRepCanonical_(raw, aliasMap) || (raw || null); // 8名に無い担当は営業担当の生値でタブ化（取りこぼし防止）
    }
    if (!rep) { excluded["(空欄)"] = (excluded["(空欄)"] || 0) + 1; continue; } // 営業担当も紹介者も空の行のみ除外
    const outRow = row.slice(0, w);
    while (outRow.length < w) outRow.push("");
    const cc = custKey ? (byCust[custKey] || {}) : {};
    Object.keys(caseColIdxs).forEach(function (cn) {
      if (cc[cn]) outRow[caseColIdxs[cn]] = cc[cn].cell; // C に状態があれば {月}月{状態}、無ければ統合の元値を保持
    });
    if (!perRep[rep]) perRep[rep] = [];
    perRep[rep].push(outRow);
  }

  // ---- SS1 に「総合_<担当>」タブを書き込み ----
  const canonicalReps = JISHA_REFERRER_OPTIONS.split(",").map(function (s) { return s.trim(); });
  Object.keys(perRep).forEach(function (r) { if (canonicalReps.indexOf(r) < 0) canonicalReps.push(r); }); // 8名以外の担当も追加
  const perRepRows = {};
  canonicalReps.forEach(function (rep) {
    const tabName = "総合_" + rep;
    let tab = ss1.getSheetByName(tabName);
    if (!tab) tab = ss1.insertSheet(tabName);
    tab.clear();
    tab.getRange(1, 1, tab.getMaxRows(), tab.getMaxColumns()).clearDataValidations();
    const rows = perRep[rep] || [];
    const all = [header.slice(0, w)].concat(rows);
    tab.getRange(1, 1, all.length, w).setValues(all);
    tab.getRange(1, 1, all.length, w).setBorder(true, true, true, true, true, true); // 枠線
    if (rows.length > 0) {
      const bg = rows.map(function (rw) {
        const b = []; for (let i = 0; i < w; i++) b.push("#ffffff");
        Object.keys(caseColIdxs).forEach(function (cn) {
          const idx = caseColIdxs[cn]; const v = String(rw[idx] || "");
          if (v.indexOf("承認") >= 0 && v.indexOf("非承認") < 0) b[idx] = "#d9ead3";
          else if (v.indexOf("非承認") >= 0) b[idx] = "#f4cccc";
          else if (v.indexOf("申請") >= 0) b[idx] = "#fff2cc";
        });
        return b;
      });
      tab.getRange(2, 1, rows.length, w).setBackgrounds(bg);
    }
    perRepRows[rep] = rows.length;
  });

  Logger.log("=== buildIntegratedRepSheets 結果 ===");
  Logger.log("担当別 行数: " + JSON.stringify(perRepRows));
  Logger.log("担当不明で除外: " + JSON.stringify(excluded));
  return { perRepRows: perRepRows, excluded: excluded };
}

// =============================================
// 統合顧客管理（SS1）から、元データC(設定_フォーム群)に実在しない
// 「アフィリンクのみ」の幽霊/重複行を削除する。
// 安全策: 氏名＋担当＋状態で内容一致した行のみ削除。各条件で一致1件でなければSKIP。
// 削除前に行内容をログ化して返す。
// =============================================
function cleanupIntegratedPhantomRows() {
  const ss1 = SpreadsheetApp.openById(REP_STATUS_SS1_ID);
  const sh = ss1.getSheetByName("統合顧客管理");
  if (!sh) throw new Error("統合顧客管理 タブが見つかりません");
  const data = sh.getDataRange().getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });
  const nameCol = header.indexOf("名前");
  const repCol = header.indexOf("営業担当");
  const stateCol = header.indexOf("状態");
  if (nameCol < 0 || repCol < 0 || stateCol < 0) throw new Error("必要列(名前/営業担当/状態)が見つかりません");

  // C照合で確定した幽霊/重複行（2026-07-06 調査）
  const targets = [
    { name: "高橋祐樹",   rep: "柳沢悠貴", reason: "Cは髙橋祐樹＝菅原(別行が正)" },
    { name: "榎本",       rep: "柳沢悠貴", reason: "Cは榎本彩人＝岩本(別行が正)" },
    { name: "岩鼻ひかる", rep: "菅原貴博", reason: "Cに該当者なし" }
  ];

  const toDelete = [];   // {row, snapshot}
  const report = [];
  targets.forEach(function (t) {
    const hits = [];
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][nameCol]).trim() === t.name &&
          String(data[r][repCol]).trim() === t.rep &&
          String(data[r][stateCol]).indexOf("アフィリンクのみ") >= 0) {
        hits.push(r);
      }
    }
    if (hits.length === 1) {
      const r = hits[0];
      const snap = [];
      for (let c = 0; c < header.length; c++) { const v = data[r][c]; if (v !== "" && v !== null) snap.push(header[c] + "=" + v); }
      toDelete.push({ row: r + 1, snapshot: snap.join(" | ") }); // 1-indexed シート行
      report.push("DELETE r" + (r + 1) + " " + t.name + "/" + t.rep + " [" + t.reason + "] :: " + snap.join(" | "));
    } else {
      report.push("SKIP " + t.name + "/" + t.rep + " (一致=" + hits.length + "件のため保留)");
    }
  });

  // 安全上限: 想定は3行。8行超なら誤検知の恐れがあるため中止。
  if (toDelete.length > 8) throw new Error("削除候補が多すぎます(" + toDelete.length + ")。中止。");

  // ★重要: フィルタがあると deleteRow が無言で失敗する。一旦外してから削除し、後で再作成する。
  var hadFilter = false;
  var existing = sh.getFilter();
  if (existing) { existing.remove(); hadFilter = true; }

  // 下から順に削除（行番号ズレ防止）
  toDelete.map(function (d) { return d.row; }).sort(function (a, b) { return b - a; }).forEach(function (rn) { sh.deleteRow(rn); });
  SpreadsheetApp.flush(); // 削除を確定させてから後続(再生成)が最新を読むように

  // フィルタを元のデータ範囲に再作成（フィルタ条件は保持されず、素のフィルタを復元）
  if (hadFilter && sh.getLastRow() > 1) {
    sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).createFilter();
  }

  Logger.log("=== cleanupIntegratedPhantomRows ===");
  report.forEach(function (l) { Logger.log(l); });
  return { deleted: toDelete.length, detail: report };
}
