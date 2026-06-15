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
const JISHA_REFERRER_OPTIONS = "柳沢悠貴,岩本拓也,菅原貴博,村井亮介,大島雅史,小椋裕也,細川貴弘,藤森宣哉";

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
const ss       = getOrCreateSpreadsheet();
    const formName = (e && e.parameter && e.parameter.form) ? e.parameter.form : getFirstFormCode(ss);
    const config   = readConfig(ss, formName);
    config.formName = formName;
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
  try {
    const data = JSON.parse(e.postData.contents);

    // LINE Webhook イベント（eventsプロパティ存在で判定）
    if (data.events !== undefined) {
      return handleLineWebhook(data);
    }

    const ss       = getOrCreateSpreadsheet();
    const formName = data.formName || getFirstFormCode(ss);
    const config   = readConfig(ss, formName);

    const sheet = getConfigSheetByCode(ss, formName);
    if (!sheet) throw new Error("設定シート（フォーム記号: " + formName + "）が見つかりません。");

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

    const rowData = buildRow(data, config, screenshotUrl, formName);
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

    // 顧客管理シートに自動追加（自社フォームのみ・紹介者名があるとき）
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
function buildHeaders(config) {
  const fieldLabels = config.fields.map(f => f.label);
  return [FORM_CODE_HEADER, "受信日時", "クリック日時", "送信日時", ...fieldLabels, "スクショURL", "承認"];
}

// ---- データ行を組み立て ----
function buildRow(data, config, screenshotUrl, formName) {
  const receivedAt  = formatJST(new Date());
  const clickAt     = data.clickTime  ? formatJST(new Date(data.clickTime))  : "";
  const submitAt    = data.submitTime ? formatJST(new Date(data.submitTime)) : "";
  const fieldValues = config.fields.map(f => data[f.id] || "");
  return [formName, receivedAt, clickAt, submitAt, ...fieldValues, screenshotUrl, ""];
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
  applyReferrerSelectToJishaSheets();
  SpreadsheetApp.getUi().createMenu("フォーム管理")
    .addItem("新規フォーム作成",       "showCreateFormDialog")
    .addItem("管理シートを更新",       "updateManagementSheet")
    .addItem("代理店割り当て更新",     "rebuildAllAgencySpreadsheets")
    .addItem("旧共有SSをゴミ箱へ",     "deleteAllOldSharingSpreadsheets")
    .addSeparator()
    .addItem("日次レポート（テスト送信）",           "dailyReport")
    .addItem("30件クエスト進捗（テスト送信）",       "campaignReport")
    .addItem("150件クエスト進捗（テスト送信）",     "quest150Report")
    .addSeparator()
    .addItem("回答ヘッダーを最新フィールドに同期", "fixAnswerHeaders")
    .addItem("シート名をフォーム名に変換（移行）", "migrateSheetNamesToDisplayName")
    .addItem("フォーム記号を修復",               "repairFormCodeRows")
    .addSeparator()
    .addItem("スクショフォルダを再登録",   "resetScreenshotFolder")
    .addSeparator()
    .addItem("顧客管理シートを作成",           "createCustomerManagementSheet")
    .addItem("顧客管理の案件列を同期",         "syncCustomerManagementCases")
    .addItem("既存データを顧客管理シートへインポート", "importExistingToCustomerSheet")
    .addItem("既存データを広告主シートへインポート",   "importExistingToAdvertiserSheet")
    .addToUi();
}

// ---- 自社シートの紹介者フィールドをselectに切り替え（初回のみ自動実行）----
function applyReferrerSelectToJishaSheets() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("REFERRER_SELECT_APPLIED") === "1") return;

  const ss = getOrCreateSpreadsheet();
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
        break;
      }
    }
  });

  props.setProperty("REFERRER_SELECT_APPLIED", "1");
  Logger.log("自社シートの紹介者フィールドをselectに更新しました");
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
  const screenshotUrl = rowData[rowData.length - 2];
  if (screenshotUrl && String(screenshotUrl).startsWith("http")) {
    lines.push("スクショ: " + screenshotUrl);
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

// ---- 既存フォーム回答を広告主成果管理シートへ一括インポート ----
// 月ごとにバッファしてまとめて書き込む。同じシートに重複して実行しないこと。
function importExistingToAdvertiserSheet() {
  const advertiserSS = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  const mainSS       = getOrCreateSpreadsheet();

  const buffer = {}; // yyyymm → [[receivedAt, name, referrer, ssUrl], ...]
  let skipCount = 0;

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
        if (!buffer[yyyymm]) buffer[yyyymm] = [];
        const approved = shoninIdx >= 0 && String(row[shoninIdx] || "").trim() === "⭕️";
        buffer[yyyymm].push([
          String(row[rtIdx]                      || ""),
          formDisplayName,
          nameIdx  >= 0 ? String(row[nameIdx]  || "").trim() : "",
          refIdx   >= 0 ? String(row[refIdx]   || "").trim() : "",
          ssUrlIdx >= 0 ? String(row[ssUrlIdx] || "").trim() : "",
          approved
        ]);
      });
  });

  let importCount = 0;
  Object.keys(buffer).sort().forEach(yyyymm => {
    let targetSheet = advertiserSS.getSheetByName(yyyymm);
    if (!targetSheet) {
      targetSheet = createAdvertiserMonthSheet(advertiserSS, yyyymm);
    }
    if (!targetSheet) {
      skipCount += buffer[yyyymm].length;
      Logger.log("広告主シート「" + yyyymm + "」なし → " + buffer[yyyymm].length + "件スキップ");
      return;
    }
    // 既存データ行をクリア（1行目タイトル・2行目ヘッダーは保持し3行目から消去）
    const clearLastRow = targetSheet.getLastRow();
    const clearLastCol = Math.max(targetSheet.getLastColumn(), 6);
    if (clearLastRow > 2) {
      targetSheet.getRange(3, 1, clearLastRow - 2, clearLastCol).clearContent();
    }

    const rows = buffer[yyyymm];
    const startRow = getAdvertiserNextRow(targetSheet); // ヘッダー(2行目)の次=3行目
    targetSheet.getRange(startRow, 1, rows.length, 5).setValues(rows.map(r => r.slice(0, 5)));
    targetSheet.getRange(startRow, 7, rows.length, 1).setValues(rows.map(r => [r[5]]));
    importCount += rows.length;
    Logger.log("広告主シート「" + yyyymm + "」に " + rows.length + "件追記");
  });

  SpreadsheetApp.getUi().alert(
    "インポート完了！\n追記: " + importCount + "件" +
    (skipCount > 0 ? "\nスキップ（シートなし）: " + skipCount + "件" : "")
  );
}

// ---- 広告主シートに月別シートを新規作成（既存シートをテンプレートにコピー）----
function createAdvertiserMonthSheet(advertiserSS, yyyymm) {
  const monthSheets = advertiserSS.getSheets()
    .filter(s => /^\d{6}$/.test(s.getName()))
    .sort((a, b) => a.getName().localeCompare(b.getName()));

  let newSheet;
  if (monthSheets.length > 0) {
    const template = monthSheets[0]; // 最古のシートを書式テンプレートに使用
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

// ---- 広告主成果管理シートに行を追記 ----
// シート名は YYYYMM 形式（例: 202606）。対象シートが存在しない場合はログのみ。
function writeToAdvertiserSheet(receivedAt, formDisplayName, customerName, referrerName, screenshotUrl, approved) {
  const ss   = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  const now  = new Date();
  const jst  = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const p    = n => String(n).padStart(2, "0");
  const sheetName = jst.getUTCFullYear() + p(jst.getUTCMonth() + 1);

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("広告主シート「" + sheetName + "」が見つかりません。スキップします。");
    return;
  }

  const nextRow = getAdvertiserNextRow(sheet);
  sheet.getRange(nextRow, 1, 1, 5).setValues([[receivedAt, formDisplayName, customerName, referrerName, screenshotUrl]]);
  sheet.getRange(nextRow, 7, 1, 1).setValue(approved || false);
  Logger.log("広告主シート「" + sheetName + "」に追記: " + customerName);
}
