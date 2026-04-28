// =============================================
// アフィリエイトフォーム - メインスクリプト
// =============================================

const DRIVE_FOLDER        = "アフィリエイト_スクショ";
const CONFIG_PREFIX       = "設定_";
const FORM_BASE_URL       = "https://kazu02.github.io/affiliate-form/";
const ANSWER_START_COL    = 7; // G列から回答を記録
const MANAGEMENT_SHEET    = "管理"; // 管理シート名
const LINE_PUSH_API       = "https://api.line.me/v2/bot/message/push";

// 代理店関連
const AGENCY_KEY          = "代理店コード";
const AGENCY_DEFAULT      = "house";        // 自社直営業の内部コード
const AGENCY_DEFAULT_NAME = "自社直営業";   // 自社直営業のSS表示名
const AGENCY_PREFIX       = "代理店_";      // 代理店SSの名前プレフィックス
const AGENCY_FOLDER       = "代理店スプシ"; // Drive保管フォルダ名
const AGENCY_PATTERN      = /^[a-zA-Z0-9_]+$/;
const AGENCY_PROP_PREFIX  = "AGENCY_SS_";   // ScriptProperties キー prefix

// ---- GET: フォーム設定を返す ----
function doGet(e) {
  try {
    const ss       = getOrCreateSpreadsheet();
    const formName = (e && e.parameter && e.parameter.form) ? e.parameter.form : getFirstFormName(ss);
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
    const formName = data.formName || getFirstFormName(ss);
    const config   = readConfig(ss, formName);

    const sheet = ss.getSheetByName(CONFIG_PREFIX + formName);
    if (!sheet) throw new Error("設定シート「" + CONFIG_PREFIX + formName + "」が見つかりません。");

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
      let agencySheet = agencySS.getSheetByName(CONFIG_PREFIX + formName);
      if (!agencySheet) {
        syncFormSheetToAgency(ss, agencySS, formName);
        agencySheet = agencySS.getSheetByName(CONFIG_PREFIX + formName);
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
  configSheet.setName(CONFIG_PREFIX + "フォーム1");
  initConfigSheet(configSheet, "申請フォーム");

  return ss;
}

// ---- 設定シートが存在する最初のフォーム名を返す ----
function getFirstFormName(ss) {
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.startsWith(CONFIG_PREFIX)) {
      return name.replace(CONFIG_PREFIX, "");
    }
  }
  throw new Error("設定シートが見つかりません。スプレッドシートに「" + CONFIG_PREFIX + "フォーム名」シートを作成してください。");
}

// ---- 設定シートの初期データ ----
function initConfigSheet(sheet, title) {
  title = title || "申請フォーム";
  const formName = sheet.getName().replace(CONFIG_PREFIX, "");
  const formUrl  = FORM_BASE_URL + "?form=" + encodeURIComponent(formName);

  const data = [
    ["フォームタイトル",    title],
    ["フォーム説明文",      "以下の手順に従って入力・作業を行ってください。"],
    ["アフィリエイトURL",   "https://ここにアフィリエイトリンクを入力"],
    ["ボタンテキスト",      "アフィリエイトリンクを開く（必ずここから！）"],
    ["フォームURL（自動）", formUrl],
    [AGENCY_KEY,            ""],
    ["", ""],
    ["＝＝ フォーム項目（行を追加・削除で変更可） ＝＝", ""],
    ["フィールドID", "ラベル", "タイプ(text/textarea/select)", "必須(TRUE/FALSE)", "プレースホルダー"],
    ["name",     "お名前",   "text", "TRUE", "例：山田太郎"],
    ["referrer", "紹介者名", "text", "TRUE", "例：田中花子"],
  ];
  sheet.getRange(1, 1, data.length, 5).setValues(
    data.map(row => { while (row.length < 5) row.push(""); return row; })
  );

  // フォームURL行（5行目）
  const urlRange = sheet.getRange(5, 1, 1, 2);
  urlRange.setBackground("#e8f5e9").setFontColor("#1b5e20").setFontWeight("bold");
  sheet.getRange(5, 2).setFontStyle("italic");

  // 代理店コード行（6行目）
  const agencyRange = sheet.getRange(6, 1, 1, 2);
  agencyRange.setBackground("#e0f2fe").setFontColor("#075985").setFontWeight("bold");

  // フォーム項目ヘッダー行（9行目）
  const headerRange = sheet.getRange(9, 1, 1, 5);
  headerRange.setFontWeight("bold").setBackground("#e8eaf6").setFontColor("#1a237e");

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 200);
}

// ---- 設定シートを読み込む ----
function readConfig(ss, formName) {
  const sheetName = CONFIG_PREFIX + formName;
  const sheet     = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("設定シート「" + sheetName + "」が見つかりません。");
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
    formTitle: "", formDescription: "", affiliateUrl: "", affiliateButtonText: "", fields: []
  };
  const keyMap = {
    "フォームタイトル":  "formTitle",
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
        placeholder: String(row[4] || "")
      });
    });
  }
  return config;
}

// ---- ヘッダー行を組み立て ----
function buildHeaders(config) {
  const fieldLabels = config.fields.map(f => f.label);
  return ["フォーム名", "受信日時", "クリック日時", "送信日時", ...fieldLabels, "スクショURL", "承認"];
}

// ---- データ行を組み立て ----
function buildRow(data, config, screenshotUrl, formName) {
  const receivedAt  = formatJST(new Date());
  const clickAt     = data.clickTime  ? formatJST(new Date(data.clickTime))  : "";
  const submitAt    = data.submitTime ? formatJST(new Date(data.submitTime)) : "";
  const fieldValues = config.fields.map(f => data[f.id] || "");
  return [formName, receivedAt, clickAt, submitAt, ...fieldValues, screenshotUrl, ""];
}

// ---- スクショ保存 ----
function saveScreenshot(base64Data, fileName, data) {
  try {
    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);
    const base64  = base64Data.split(",")[1];
    const mime    = base64Data.split(";")[0].split(":")[1];
    const blob    = Utilities.newBlob(Utilities.base64Decode(base64), mime, fileName);
    const name    = data.name || "不明";
    const dt      = data.clickTime ? formatJSTforFilename(new Date(data.clickTime)) : formatJSTforFilename(new Date());
    blob.setName(`${dt}_${name}_${fileName || "screenshot"}`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
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
  updateManagementSheet();
  SpreadsheetApp.getUi().createMenu("フォーム管理")
    .addItem("新規フォーム作成",       "showCreateFormDialog")
    .addItem("管理シートを更新",       "updateManagementSheet")
    .addItem("代理店割り当て更新",     "rebuildAllAgencySpreadsheets")
    .addItem("旧共有SSをゴミ箱へ",     "deleteAllOldSharingSpreadsheets")
    .addToUi();
}

// ---- 全設定シートの回答ヘッダーをG1に初期化（未設定のシートのみ）----
function initAllAnswerHeaders() {
  const ss = getOrCreateSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    if (sheet.getRange(1, ANSWER_START_COL).getValue()) return;
    const formName = name.replace(CONFIG_PREFIX, "");
    const config   = readConfig(ss, formName);
    const headers  = buildHeaders(config);
    const range    = sheet.getRange(1, ANSWER_START_COL, 1, headers.length);
    range.setValues([headers]);
    range.setFontWeight("bold");
    range.setBackground("#4f46e5");
    range.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  });
}

// ---- 全設定シートのフォームURLを更新 ----
function updateAllFormUrls() {
  const ss     = getOrCreateSpreadsheet();
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const formName = name.replace(CONFIG_PREFIX, "");
    const formUrl  = FORM_BASE_URL + "?form=" + encodeURIComponent(formName);
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
function syncFormSheetToAgency(ss, agencySS, formName) {
  const mainSheet = ss.getSheetByName(CONFIG_PREFIX + formName);
  if (!mainSheet) return;

  let agencySheet = agencySS.getSheetByName(CONFIG_PREFIX + formName);
  if (!agencySheet) {
    agencySheet = agencySS.insertSheet(CONFIG_PREFIX + formName);
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
    const formName = name.replace(CONFIG_PREFIX, "");
    let code;
    try { code = getAgencyCode(sheet); }
    catch (e) {
      Logger.log("代理店コードエラー(" + formName + "): " + e);
      return;
    }
    if (!formByAgency[code]) formByAgency[code] = [];
    formByAgency[code].push(formName);
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
      const formName = name.replace(CONFIG_PREFIX, "");
      if (formNames.indexOf(formName) < 0) {
        agencySS.deleteSheet(sheet);
        Logger.log(code + ": 不要シート削除 - " + formName);
      }
    });

    // 代理店SSの管理シートを更新
    updateAgencyManagementSheet(agencySS, code, formNames, ss);
  });

  updateManagementSheet();
  Logger.log("代理店割り当て更新 完了");
}

// ---- 代理店SSの管理シートを更新 ----
function updateAgencyManagementSheet(agencySS, code, formNames, mainSS) {
  let mgSheet = agencySS.getSheetByName(MANAGEMENT_SHEET);
  if (!mgSheet) {
    mgSheet = agencySS.insertSheet(MANAGEMENT_SHEET, 0);
  } else if (mgSheet.getIndex() !== 1) {
    agencySS.setActiveSheet(mgSheet);
    agencySS.moveActiveSheet(1);
  }
  mgSheet.clearContents();
  mgSheet.clearFormats();

  const headers = ["フォーム名", "フォームURL", "代理店コード", "回答数", "最終回答日時"];
  const hRange  = mgSheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  mgSheet.setFrozenRows(1);

  const rows = [];
  formNames.forEach(formName => {
    const sheet = mainSS.getSheetByName(CONFIG_PREFIX + formName);
    if (!sheet) return;
    const values = sheet.getDataRange().getValues();
    let formUrl = "";
    for (const row of values) {
      if (String(row[0]).includes("フォームURL")) formUrl = String(row[1] || "");
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
    rows.push([formName, formUrl, code, answerCount, lastAnswerAt]);
  });

  if (rows.length > 0) {
    mgSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    rows.forEach((row, i) => {
      if (row[1]) mgSheet.getRange(i + 2, 2).setFontColor("#1155cc").setFontStyle("italic");
      if (row[2]) mgSheet.getRange(i + 2, 3).setFontColor("#075985").setFontWeight("bold");
    });
    rows.forEach((_, i) => {
      if (i % 2 === 1) mgSheet.getRange(i + 2, 1, 1, headers.length).setBackground("#f3f4f6");
    });
  }

  mgSheet.setColumnWidth(1, 160);
  mgSheet.setColumnWidth(2, 360);
  mgSheet.setColumnWidth(3, 120);
  mgSheet.setColumnWidth(4, 80);
  mgSheet.setColumnWidth(5, 160);
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

    const formName = sheetName.replace(CONFIG_PREFIX, "");
    const ss       = getOrCreateSpreadsheet();
    syncApprovalToMain(ss, formName, receivedAt, newValue);
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

// ---- メインSS側で承認列編集 → 代理店SSに反映 ----
function onMainEdit(e) {
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

    const formName = sheetName.replace(CONFIG_PREFIX, "");
    let code;
    try { code = getAgencyCode(editedSheet); }
    catch (err) { return; }
    const agencySS = getOrCreateAgencySpreadsheet(code);
    syncApprovalToAgency(agencySS, formName, receivedAt, newValue);
  } catch (err) {
    Logger.log("onMainEdit error: " + err);
  }
}

// ---- 承認値をメインSSに同期 ----
function syncApprovalToMain(ss, formName, receivedAt, value) {
  const sheet = ss.getSheetByName(CONFIG_PREFIX + formName);
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
  const sheet = agencySS.getSheetByName(CONFIG_PREFIX + formName);
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

  const headers = ["フォーム名", "フォームURL", "代理店コード", "回答数", "最終回答日時"];
  const hRange  = mgSheet.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  mgSheet.setFrozenRows(1);

  const rows = [];
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    const formName = name.replace(CONFIG_PREFIX, "");
    const values   = sheet.getDataRange().getValues();

    let formUrl    = "";
    let agencyCode = "";
    for (const row of values) {
      if (String(row[0]).includes("フォームURL")) formUrl    = String(row[1] || "");
      if (String(row[0]) === AGENCY_KEY)         agencyCode = String(row[1] || "");
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
    rows.push([formName, formUrl, agencyCode, answerCount, lastAnswerAt]);
  });

  if (rows.length > 0) {
    mgSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    rows.forEach((row, i) => {
      if (row[1]) mgSheet.getRange(i + 2, 2).setFontColor("#1155cc").setFontStyle("italic");
      if (row[2]) mgSheet.getRange(i + 2, 3).setFontColor("#075985").setFontWeight("bold");
    });
    rows.forEach((_, i) => {
      if (i % 2 === 1) mgSheet.getRange(i + 2, 1, 1, headers.length).setBackground("#f3f4f6");
    });
  }

  mgSheet.setColumnWidth(1, 160);
  mgSheet.setColumnWidth(2, 360);
  mgSheet.setColumnWidth(3, 120);
  mgSheet.setColumnWidth(4, 80);
  mgSheet.setColumnWidth(5, 160);
}

// =============================================
// LINE 通知
// =============================================

// ---- LINE Webhook 受信 ----
function handleLineWebhook(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
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
function buildLineMessage(config, rowData, formName) {
  const lines = ["【新規申請】" + (config.formTitle || formName)];
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

// ---- LINE通知テスト（GASエディタから手動実行）----
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
  const formName     = String(data.formName     || "").trim();
  const formTitle    = String(data.formTitle    || "").trim();
  const affiliateUrl = String(data.affiliateUrl || "").trim();
  const agencyCode   = String(data.agencyCode   || "").trim();

  if (!formName)     throw new Error("フォーム名を入力してください。");
  if (!formTitle)    throw new Error("フォームタイトルを入力してください。");
  if (!affiliateUrl) throw new Error("アフィリエイトURLを入力してください。");
  if (agencyCode && !AGENCY_PATTERN.test(agencyCode)) {
    throw new Error("代理店コードは半角英数字とアンダースコアのみ使用できます。");
  }

  const ss = getOrCreateSpreadsheet();
  if (ss.getSheetByName(CONFIG_PREFIX + formName)) {
    throw new Error("「" + formName + "」は既に存在します。");
  }

  const configSheet = ss.insertSheet(CONFIG_PREFIX + formName);
  initConfigSheet(configSheet, formTitle);

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
  const config = readConfig(ss, formName);
  const hdrs   = buildHeaders(config);
  const hRange = configSheet.getRange(1, ANSWER_START_COL, 1, hdrs.length);
  hRange.setValues([hdrs]);
  hRange.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
  configSheet.setFrozenRows(1);

  // 代理店SSに同期
  const code     = agencyCode || AGENCY_DEFAULT;
  const agencySS = getOrCreateAgencySpreadsheet(code);
  syncFormSheetToAgency(ss, agencySS, formName);
  // 代理店SSの管理シートも更新
  const formNames = [];
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(CONFIG_PREFIX)) return;
    try {
      if (getAgencyCode(sheet) === code) {
        formNames.push(name.replace(CONFIG_PREFIX, ""));
      }
    } catch (e) {}
  });
  updateAgencyManagementSheet(agencySS, code, formNames, ss);

  // メインSSの管理シート更新
  updateManagementSheet();

  return { formName: formName, formUrl: FORM_BASE_URL + "?form=" + encodeURIComponent(formName) };
}
