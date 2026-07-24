/**
 * アフィリエイト報酬 振込先フォーム
 * ------------------------------------------------------------------
 * - setup(): Google フォームを生成し、送信トリガーを登録する。
 * - onSubmit(e): 回答を受け取り、統合顧客管理ブック(SS1)の各営業担当シートで
 *   フルネームを名寄せ照合し、「持ち家かどうか」列の直後に振込先情報を書き込む。
 * - inspect(): SS1 の構造(タブ名・ヘッダー・持ち家/名前列位置)を調べて返す。
 *
 * 実行アカウント: shinhogle@gmail.com
 */

// フォームのテスト回答を全削除する（検証用。エディタから実行）
function purgeFormResponses() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_FORM_ID);
  if (!id) return { error: "no form id" };
  var form = FormApp.openById(id);
  var n = form.getResponses().length;
  form.deleteAllResponses();
  return { deleted: n };
}

// 検証用ヘルパーシートを削除する
function cleanupHelperSheets() {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var removed = [];
  ["_payout_meta", "_payout_inspect"].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh) { ss.deleteSheet(sh); removed.push(n); }
  });
  return { removed: removed };
}

// 名前でマスター＋担当タブの振込先セルを読み取る（検証用）
function readRowByName_(name) {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var targets = [];
  var master = ss.getSheetByName(MASTER_SHEET);
  if (master) targets.push(master);
  getSalespersonSheets_().forEach(function (sh) { targets.push(sh); });
  var res = [];
  targets.forEach(function (sheet) {
    var row = findMatchRow_(sheet, name, null);
    if (row <= 0) return;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var col = headers.indexOf(PAYOUT_HEADERS[0]);
    if (col < 0) return;
    var v = sheet.getRange(row, col + 1, 1, PAYOUT_HEADERS.length).getValues()[0];
    res.push({ sheet: sheet.getName(), row: row, values: v });
  });
  return res;
}

// ===== 定数 =====
const SS1_ID = "1aaiCIDQIkrp_Ado5aKua_PTEQq4jr1UWqpuLQXwpemI"; // 統合顧客管理ブック
// ===== 営業担当 名簿：単一の変更点 =====
// メンバーの増減はこの配列を編集し、エディタで setup() を1回実行するだけ。
// setup() が既存フォームの「紹介営業担当」選択肢をこの配列へ同期する。
const SALESPEOPLE = ["柳沢悠貴", "岩本拓也", "菅原貴博", "村井亮介", "大島雅史", "小椋裕也", "細川貴弘", "藤森宣哉", "江口裕人", "藤井勇大"];

const OWNERSHIP_HEADER = "持ち家かどうか";
const NAME_HEADER_CANDIDATES = ["名前", "顧客名", "お名前", "氏名"];

// フォームの設問タイトル（onSubmit のマッピングキー）
const Q_NAME     = "お客様のお名前（フルネーム）";
const Q_REFERRER = "ご紹介いただいた営業担当者";
const Q_BANK     = "金融機関名（銀行名）";
const Q_BRANCH   = "支店名";
const Q_ACCTTYPE = "預金種目";
const Q_ACCTNUM  = "口座番号";
const Q_HOLDER   = "口座名義（カナ）";

// 「持ち家かどうか」の直後に挿入する振込先列（この順で6列）
const PAYOUT_HEADERS = ["金融機関名", "支店名", "預金種目", "口座番号", "口座名義(カナ)", "振込先登録日時"];

// 統合顧客管理（マスター）タブ名。総合_<担当> はここから再生成されるため、
// 振込先を確実に残すにはマスターにも書き込む。
const MASTER_SHEET = "統合顧客管理";
const REFERRER_HEADER_CANDIDATES = ["営業担当", "紹介者名", "紹介者"];

// ScriptProperties キー
const PROP_FORM_ID = "PAYOUT_FORM_ID";
const PROP_UNMATCHED_SHEET = "振込先_未照合ログ";

// ===== 名寄せユーティリティ =====
function toHiragana(str) {
  return String(str || "").replace(/[ァ-ヶ]/g, function (ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0x60);
  });
}
function normalizeName(name) {
  var s = String(name || "").trim();
  s = s.replace(/[\s\u00a0\u200b-\u200f\u202f\u205f\u3000\ufeff]+/g, "");
  return toHiragana(s).toLowerCase();
}
function levenshtein(a, b) {
  var m = a.length, n = b.length;
  var dp = [];
  for (var i = 0; i <= m; i++) { dp.push([i]); }
  for (var j = 1; j <= n; j++) { dp[0][j] = j; }
  for (var i2 = 1; i2 <= m; i2++) {
    for (var j2 = 1; j2 <= n; j2++) {
      dp[i2][j2] = a[i2 - 1] === b[j2 - 1] ? dp[i2 - 1][j2 - 1]
        : 1 + Math.min(dp[i2 - 1][j2], dp[i2][j2 - 1], dp[i2 - 1][j2 - 1]);
    }
  }
  return dp[m][n];
}
// フルネーム同士の一致判定（完全一致→先頭一致→編集距離1）
function namesMatch(a, b) {
  var na = normalizeName(a), nb = normalizeName(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  // 一方が他方の先頭に一致（苗字のみ等）。ただし2文字以上の重なりが必要。
  var shorter = na.length <= nb.length ? na : nb;
  var longer = na.length <= nb.length ? nb : na;
  if (shorter.length >= 2 && longer.indexOf(shorter) === 0) return true;
  // 表記ゆれ（4文字以上で編集距離1以内）
  if (na.length >= 4 && nb.length >= 4 && levenshtein(na, nb) <= 1) return true;
  return false;
}

// ===== SS1 構造ヘルパー =====
// SS1 の営業担当シート一覧（タブ名が営業担当名を含むもの）
function getSalespersonSheets_() {
  var ss = SpreadsheetApp.openById(SS1_ID);
  return ss.getSheets().filter(function (sh) {
    var name = sh.getName();
    return SALESPEOPLE.some(function (rep) { return name.indexOf(rep) >= 0; });
  });
}
// 指定営業担当に対応するシート（無ければ null）
function getSheetForReferrer_(referrer) {
  var sheets = getSalespersonSheets_();
  var norm = normalizeName(referrer);
  for (var i = 0; i < sheets.length; i++) {
    if (normalizeName(sheets[i].getName()).indexOf(norm) >= 0 ||
        norm.indexOf(normalizeName(sheets[i].getName())) >= 0) {
      return sheets[i];
    }
  }
  // タブ名一致しなければ、タブ名に営業担当名が含まれるかで再探索
  for (var j = 0; j < sheets.length; j++) {
    if (sheets[j].getName().indexOf(referrer) >= 0) return sheets[j];
  }
  return null;
}
// ヘッダー行(1行目)から列 index を得る
function findHeaderIndex_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    for (var c = 0; c < candidates.length; c++) {
      if (h === candidates[c]) return i;
    }
  }
  return -1;
}
function findNameCol_(headers) { return findHeaderIndex_(headers, NAME_HEADER_CANDIDATES); }
function findOwnershipCol_(headers) { return findHeaderIndex_(headers, [OWNERSHIP_HEADER]); }

// ===== inspect: 構造調査 =====
function inspect() {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var out = { title: ss.getName(), sheets: [] };
  ss.getSheets().forEach(function (sh) {
    var lastCol = sh.getLastColumn();
    var lastRow = sh.getLastRow();
    var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
    out.sheets.push({
      name: sh.getName(),
      rows: lastRow,
      cols: lastCol,
      isSalesperson: SALESPEOPLE.some(function (r) { return sh.getName().indexOf(r) >= 0; }),
      nameCol: findNameCol_(headers),
      ownershipCol: findOwnershipCol_(headers),
      afterOwnership: (function () {
        var oc = findOwnershipCol_(headers);
        return oc >= 0 ? headers.slice(oc, oc + 8) : [];
      })(),
      headers: headers
    });
  });
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// inspect の結果を SS1 のシート "_payout_inspect" に書き出す（gviz で読み取り可能にする）
function inspectAndReport() {
  var data = inspect();
  var ss = SpreadsheetApp.openById(SS1_ID);
  var name = "_payout_inspect";
  var sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(name, 0);
  var rows = [["tab", "rows", "cols", "isSalesperson", "nameColIdx(0based)", "ownershipColIdx(0based)", "afterOwnership", "headers"]];
  data.sheets.forEach(function (s) {
    rows.push([
      s.name, s.rows, s.cols, s.isSalesperson ? "Y" : "",
      s.nameCol, s.ownershipCol,
      (s.afterOwnership || []).join(" | "),
      (s.headers || []).join(" | ")
    ]);
  });
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return name;
}

// ===== 振込先列の用意（各営業担当シートに1回だけ挿入） =====
// 「持ち家かどうか」の直後に PAYOUT_HEADERS を挿入。既に存在すればスキップ。
// 戻り値: シート名 -> 振込先ブロック開始列(1-based) のマップ
function ensurePayoutColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var ownIdx = findOwnershipCol_(headers); // 0-based
  if (ownIdx < 0) return -1; // 持ち家列が無いシートは対象外

  // 既に金融機関名が入っていれば、その列を返す（冪等）
  var bankIdx = headers.indexOf(PAYOUT_HEADERS[0]);
  if (bankIdx >= 0) return bankIdx + 1;

  var insertAfter = ownIdx + 1; // 1-based の持ち家列
  sheet.insertColumnsAfter(insertAfter, PAYOUT_HEADERS.length);
  var startCol = insertAfter + 1;
  sheet.getRange(1, startCol, 1, PAYOUT_HEADERS.length).setValues([PAYOUT_HEADERS]);
  sheet.getRange(1, startCol, 1, PAYOUT_HEADERS.length).setFontWeight("bold").setBackground("#e6e6e6");
  return startCol;
}

function ensurePayoutColumnsAllSheets() {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var sheets = [];
  var master = ss.getSheetByName(MASTER_SHEET);
  if (master) sheets.push(master);
  getSalespersonSheets_().forEach(function (sh) { sheets.push(sh); });

  var report = {};
  sheets.forEach(function (sh) {
    report[sh.getName()] = ensurePayoutColumns_(sh);
  });
  Logger.log("ensurePayoutColumnsAllSheets: " + JSON.stringify(report));
  return report;
}

// ===== 推奨エントリポイント =====
// setup() + selfTest() を実行し、全結果を _payout_meta に書き出す（1回の実行で設定と検証を完了）。
function runSetupAndSelfTest() {
  var report = { status: "", steps: {} };
  try {
    report.steps.setup = setup();
    report.status = "SETUP_OK";
    report.steps.selfTest = selfTest();
    report.status = "OK";
  } catch (err) {
    report.status = "ERROR";
    report.error = String(err) + " | " + (err && err.stack ? err.stack : "");
    Logger.log("runSetupAndSelfTest ERROR: " + report.error);
  }
  writeMetaReport_(report);
  return report;
}

// マスターから「営業担当が8名のいずれか」の実顧客を1件選び、振込先の書き込み→読み戻し→クリア を検証する。
function selfTest() {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var master = ss.getSheetByName(MASTER_SHEET);
  if (!master) return { ok: false, reason: "マスター(" + MASTER_SHEET + ")が見つからない" };
  var lastRow = master.getLastRow(), lastCol = master.getLastColumn();
  var headers = master.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var nameCol = findNameCol_(headers);
  var refCol = findHeaderIndex_(headers, REFERRER_HEADER_CANDIDATES);
  if (nameCol < 0) return { ok: false, reason: "名前列が見つからない" };
  var data = master.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var picked = null;
  for (var r = 0; r < data.length; r++) {
    var nm = String(data[r][nameCol] || "").trim();
    var rf = refCol >= 0 ? String(data[r][refCol] || "").trim() : "";
    if (nm && SALESPEOPLE.indexOf(rf) >= 0) { picked = { name: nm, ref: rf }; break; }
  }
  if (!picked) return { ok: false, reason: "8名担当の実顧客が見つからない" };

  var payout = ["__SELFTEST__", "__SELFTEST__", "普通", "0000000", "セルフテスト", "(selftest)"];
  var wr = null, verify = [], cleared = [];
  try {
    wr = writePayoutAll_(picked.name, picked.ref, payout);
    (wr.wrote || []).forEach(function (w) {
      var sh = ss.getSheetByName(w.sheet);
      var v = sh.getRange(w.row, w.startCol, 1, PAYOUT_HEADERS.length).getValues()[0];
      verify.push(w.sheet + "#" + w.row + "=[" + v.join("|") + "]");
    });
  } finally {
    if (wr) (wr.wrote || []).forEach(function (w) {
      var sh = ss.getSheetByName(w.sheet);
      sh.getRange(w.row, w.startCol, 1, PAYOUT_HEADERS.length).clearContent();
      cleared.push(w.sheet + "#" + w.row);
    });
  }
  return { ok: true, picked: picked, wroteCount: (wr && wr.wrote ? wr.wrote.length : 0), verify: verify, cleared: cleared };
}

// レポートを _payout_meta に複数行で書き出す（gviz で回収）。
function writeMetaReport_(report) {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var name = "_payout_meta";
  var sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(name, 0);
  var s = report.steps || {};
  var setupR = s.setup || {};
  var st = s.selfTest || {};
  var rows = [
    ["status", report.status || ""],
    ["error", report.error || ""],
    ["formId", setupR.formId || ""],
    ["publishedUrl", setupR.publishedUrl || ""],
    ["editUrl", setupR.editUrl || ""],
    ["reused", setupR.reused ? "Y" : ""],
    ["selfTest.ok", st.ok ? "Y" : (st.reason || "")],
    ["selfTest.picked", st.picked ? (st.picked.name + " / " + st.picked.ref) : ""],
    ["selfTest.wroteCount", String(st.wroteCount || 0)],
    ["selfTest.verify", (st.verify || []).join("  ||  ")],
    ["selfTest.cleared", (st.cleared || []).join(", ")]
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
}

// ===== エディタから実行するエントリポイント（setup のみ） =====
// setup() を実行し、結果またはエラーを _payout_meta に書き出す（外部から gviz で確認するため）。
function runSetup() {
  try {
    var res = setup();
    Logger.log("runSetup OK: " + JSON.stringify(res));
    return res;
  } catch (err) {
    var msg = String(err) + " | " + (err && err.stack ? err.stack : "");
    Logger.log("runSetup ERROR: " + msg);
    try {
      writeMeta_({ formId: "(ERROR)", publishedUrl: msg.slice(0, 500), editUrl: "", reused: false });
    } catch (e2) {}
    throw err;
  }
}

// ===== フォーム生成 + トリガー登録 =====
function setup() {
  // 既存フォームがあれば作り直さない
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(PROP_FORM_ID);
  if (existing) {
    var f = FormApp.openById(existing);
    Logger.log("既存フォームを再利用: " + f.getPublishedUrl());
    syncFormChoices_(f); // 紹介営業担当の選択肢を名簿(SALESPEOPLE)へ同期
    ensurePayoutColumnsAllSheets();
    ensureSubmitTrigger_(existing);
    var reusedResult = { formId: existing, publishedUrl: f.getPublishedUrl(), editUrl: f.getEditUrl(), reused: true };
    writeMeta_(reusedResult);
    return reusedResult;
  }

  var form = FormApp.create("アフィリエイト報酬 振込先のご登録");
  form.setDescription(
    "この度はアフィリエイトプログラムにご協力いただきありがとうございます。\n" +
    "報酬のお振込のため、下記の振込先情報をご登録ください。\n" +
    "※ お名前は、お申し込み時と同じフルネームをご入力ください（照合に使用します）。"
  );
  form.setCollectEmail(false);
  form.setProgressBar(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);

  // Q1 お名前
  form.addTextItem().setTitle(Q_NAME)
    .setHelpText("お申し込み時と同じフルネームをご入力ください（例: 山田 太郎）").setRequired(true);

  // Q2 紹介営業担当（選択式）
  form.addListItem().setTitle(Q_REFERRER)
    .setHelpText("ご紹介いただいた担当者をお選びください")
    .setChoiceValues(SALESPEOPLE).setRequired(true);

  // Q3 金融機関名
  form.addTextItem().setTitle(Q_BANK).setHelpText("例: 三菱UFJ銀行 / ゆうちょ銀行").setRequired(true);

  // Q4 支店名
  form.addTextItem().setTitle(Q_BRANCH)
    .setHelpText("例: 新宿支店（ゆうちょ銀行の場合は「店番3桁」をご記入ください）").setRequired(true);

  // Q5 預金種目
  form.addMultipleChoiceItem().setTitle(Q_ACCTTYPE)
    .setChoiceValues(["普通", "当座"]).setRequired(true);

  // Q6 口座番号
  form.addTextItem().setTitle(Q_ACCTNUM)
    .setHelpText("例: 1234567（ゆうちょ銀行の場合は「記号-番号」）").setRequired(true);

  // Q7 口座名義（カナ）
  form.addTextItem().setTitle(Q_HOLDER)
    .setHelpText("例: ヤマダ タロウ（カタカナでご入力ください）").setRequired(true);

  form.setConfirmationMessage("ご登録ありがとうございました。振込先情報を承りました。");

  var formId = form.getId();
  props.setProperty(PROP_FORM_ID, formId);

  // 振込先列を各営業担当シートに用意
  ensurePayoutColumnsAllSheets();

  // 送信トリガー登録
  ensureSubmitTrigger_(formId);

  var result = { formId: formId, publishedUrl: form.getPublishedUrl(), editUrl: form.getEditUrl(), reused: false };
  writeMeta_(result);
  Logger.log("setup 完了: " + JSON.stringify(result));
  return result;
}

// setup 結果を SS1 の "_payout_meta" に書き出す（gviz で URL を回収）
function writeMeta_(result) {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var name = "_payout_meta";
  var sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(name, 0);
  sh.getRange(1, 1, 4, 2).setValues([
    ["formId", result.formId],
    ["publishedUrl", result.publishedUrl],
    ["editUrl", result.editUrl],
    ["reused", result.reused ? "Y" : ""]
  ]);
}

function ensureSubmitTrigger_(formId) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onSubmit") return; // 既存
  }
  var form = FormApp.openById(formId);
  ScriptApp.newTrigger("onSubmit").forForm(form).onFormSubmit().create();
  Logger.log("onSubmit トリガーを登録しました");
}

// 「紹介営業担当」リスト項目の選択肢を名簿(SALESPEOPLE)へ同期する（既存フォームの更新用）。
// メンバー増減後に setup() を再実行すると、この関数で選択肢が最新化される。
function syncFormChoices_(form) {
  var items = form.getItems(FormApp.ItemType.LIST);
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle().trim() === Q_REFERRER) {
      items[i].asListItem().setChoiceValues(SALESPEOPLE);
      Logger.log("紹介営業担当の選択肢を同期: " + SALESPEOPLE.join(", "));
      return true;
    }
  }
  Logger.log("紹介営業担当のリスト項目が見つかりませんでした");
  return false;
}

// ===== フォーム送信ハンドラ =====
function onSubmit(e) {
  try {
    var answers = {};
    var items = e.response.getItemResponses();
    for (var i = 0; i < items.length; i++) {
      answers[items[i].getItem().getTitle().trim()] = String(items[i].getResponse() || "").trim();
    }
    var fullName = answers[Q_NAME] || "";
    var referrer = answers[Q_REFERRER] || "";
    var payout = [
      answers[Q_BANK] || "",
      answers[Q_BRANCH] || "",
      answers[Q_ACCTTYPE] || "",
      answers[Q_ACCTNUM] || "",
      answers[Q_HOLDER] || "",
      formatTimestamp_(new Date())
    ];

    var result = writePayoutAll_(fullName, referrer, payout);
    if (!result.wrote.length) {
      logUnmatched_(fullName, referrer, payout, result.reason);
    }
    Logger.log("onSubmit: name=" + fullName + " ref=" + referrer + " -> " + JSON.stringify(result));
  } catch (err) {
    Logger.log("onSubmit エラー: " + err + "\n" + (err && err.stack));
    try { logUnmatched_("(error)", "(error)", [String(err)], "exception"); } catch (e2) {}
  }
}

// マスター(統合顧客管理)と該当営業担当タブ(総合_<担当>)の両方に振込先を書き込む。
// 総合_ タブは統合顧客管理から再生成されるため、マスターにも書くことで消失を防ぐ。
function writePayoutAll_(fullName, referrer, payout) {
  if (!fullName) return { wrote: [], reason: "氏名未入力" };
  var ss = SpreadsheetApp.openById(SS1_ID);
  var wrote = [];

  // 1) マスター（統合顧客管理）: 名前一致。同名複数なら 営業担当==紹介者 で絞り込む。
  var master = ss.getSheetByName(MASTER_SHEET);
  if (master) {
    var mrow = findMatchRow_(master, fullName, referrer);
    if (mrow > 0) {
      var mcol = ensurePayoutColumns_(master);
      if (mcol > 0) {
        writePayoutRow_(master, mrow, mcol, payout);
        wrote.push({ sheet: master.getName(), row: mrow, startCol: mcol });
      }
    }
  }

  // 2) 営業担当タブ: 紹介者のタブを優先し、無ければ全担当タブを探索（先頭一致1件）。
  var ordered = [];
  var refSheet = referrer ? getSheetForReferrer_(referrer) : null;
  if (refSheet) ordered.push(refSheet);
  getSalespersonSheets_().forEach(function (sh) {
    if (!refSheet || sh.getName() !== refSheet.getName()) ordered.push(sh);
  });
  for (var s = 0; s < ordered.length; s++) {
    var sheet = ordered[s];
    var row = findMatchRow_(sheet, fullName, null);
    if (row <= 0) continue;
    var col = ensurePayoutColumns_(sheet);
    if (col <= 0) continue;
    writePayoutRow_(sheet, row, col, payout);
    wrote.push({ sheet: sheet.getName(), row: row, startCol: col });
    break;
  }

  return { wrote: wrote, reason: wrote.length ? "" : "顧客名の照合先が見つかりません" };
}

// 振込先6セルを書き込む。口座番号・店番の先頭ゼロ落ちを防ぐため、書き込み前にテキスト書式("@")にする。
function writePayoutRow_(sheet, row, startCol, payout) {
  var range = sheet.getRange(row, startCol, 1, PAYOUT_HEADERS.length);
  range.setNumberFormat("@");
  range.setValues([payout]);
}

// シート内で fullName に一致する行(1-based)を返す。無ければ -1。
// 完全一致(正規化)を優先し、無ければあいまい一致。
// preferReferrer 指定かつ候補複数なら 営業担当列==preferReferrer で絞り込む（同名対策）。
function findMatchRow_(sheet, fullName, preferReferrer) {
  var lastCol = sheet.getLastColumn(), lastRow = sheet.getLastRow();
  if (lastRow < 2 || lastCol < 1) return -1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var nameCol = findNameCol_(headers);
  if (nameCol < 0) return -1;
  var refCol = findHeaderIndex_(headers, REFERRER_HEADER_CANDIDATES);
  var width = Math.max(nameCol, refCol) + 1;
  var data = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  var target = normalizeName(fullName);
  var exactRows = [], fuzzyRows = [];
  for (var r = 0; r < data.length; r++) {
    var nm = data[r][nameCol];
    if (!nm) continue;
    if (normalizeName(nm) === target) exactRows.push(r);
    else if (namesMatch(nm, fullName)) fuzzyRows.push(r);
  }
  var candidates = exactRows.length ? exactRows : fuzzyRows;
  if (!candidates.length) return -1;
  if (preferReferrer && refCol >= 0 && candidates.length > 1) {
    for (var c = 0; c < candidates.length; c++) {
      if (namesMatch(data[candidates[c]][refCol], preferReferrer)) return candidates[c] + 2;
    }
  }
  return candidates[0] + 2;
}

function logUnmatched_(fullName, referrer, payout, reason) {
  var ss = SpreadsheetApp.openById(SS1_ID);
  var sh = ss.getSheetByName(PROP_UNMATCHED_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROP_UNMATCHED_SHEET);
    sh.appendRow(["受信日時", "入力フルネーム", "紹介営業担当", "金融機関名", "支店名", "預金種目", "口座番号", "口座名義(カナ)", "理由"]);
  }
  var row = [formatTimestamp_(new Date()), fullName, referrer].concat(payout.slice(0, 5)).concat([reason]);
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setNumberFormat("@").setValues([row]); // 口座番号の先頭ゼロ保持
}

function formatTimestamp_(d) {
  return Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
}

// ===== テスト用 =====
// 実際のフォーム送信なしで書き込みロジックを検証する。
// 引数はスクリプトプロパティ TEST_NAME / TEST_REFERRER で上書き可能。
function testWrite() {
  var p = PropertiesService.getScriptProperties();
  var name = p.getProperty("TEST_NAME") || "須川一輝";
  var ref = p.getProperty("TEST_REFERRER") || "";
  var res = writePayoutAll_(name, ref, ["テスト銀行", "テスト支店", "普通", "9999999", "テスト メイギ", formatTimestamp_(new Date())]);
  var out = { name: name, referrer: ref, result: res };
  Logger.log("testWrite: " + JSON.stringify(out));
  writeMeta_({ formId: "(testWrite)", publishedUrl: JSON.stringify(out), editUrl: "", reused: true });
  return out;
}

// テスト書き込みの取り消し: 指定名の行の振込先6セルをクリアする（マスター＋全担当タブ）。
function clearPayoutForName() {
  var p = PropertiesService.getScriptProperties();
  var name = p.getProperty("TEST_NAME") || "";
  if (!name) { Logger.log("TEST_NAME 未設定"); return; }
  var ss = SpreadsheetApp.openById(SS1_ID);
  var targets = [];
  var master = ss.getSheetByName(MASTER_SHEET);
  if (master) targets.push(master);
  getSalespersonSheets_().forEach(function (sh) { targets.push(sh); });
  var cleared = [];
  targets.forEach(function (sheet) {
    var row = findMatchRow_(sheet, name, null);
    if (row <= 0) return;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
    var col = headers.indexOf(PAYOUT_HEADERS[0]);
    if (col < 0) return;
    sheet.getRange(row, col + 1, 1, PAYOUT_HEADERS.length).clearContent();
    cleared.push(sheet.getName() + "#" + row);
  });
  Logger.log("clearPayoutForName: " + name + " -> " + JSON.stringify(cleared));
  return cleared;
}

