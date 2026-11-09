// =============================================
// 同じ人が同じ案件へ2回以上申請したとき、1件にまとめる（2026-11-09 本人依頼）
//
// 「同じ人」= 同じ案件シート × お名前（空白・カナかな・全角半角・末尾の「様/さん」を無視）が一致。
// ただし電話番号の列があり、両方に番号が入っていて食い違うときは別人として扱う。
//
// まとめ方: 最初の申請の行を残す。残す行の空欄は後の申請の値で埋め、承認は一番強い印
// （⭕ > その他の印 > ✖ > 空欄・「重複」「ダブり」などの重複メモ）にする。
// 後の申請の行は 設定_ シートから外し、メインSSの「重複申請_統合済み」タブへ全項目を残す（戻せる）。
//
// 集計はすべて 設定_ シートの行から作っているので、ここで1行にすれば
// 日次レポート・申請状況一覧・代理店別サマリー・SS2/SS1・広告主シート・ダッシュボードが全部1件になる。
//
// 2つの入口:
//   1) 申請時（doPost → dupMergeIncoming_）: 既に同じ人の行があれば新しい行を足さずにまとめる。
//   2) 既存分（メニュー「重複申請を1件にまとめる」→ mergeDuplicateApplications）。
// =============================================

const DUP_ARCHIVE_TAB_ = "重複申請_統合済み";
const DUP_ARCHIVE_HEADERS_ = ["統合日時", "経路", "案件シート", "案件記号", "お名前", "紹介者名",
  "この申請の受信日時", "残した行の受信日時", "この申請の承認", "元の全項目(JSON)"];
const DUP_MEMO_PATTERN_ = /重複|ダブ|二重|2回目|２回目|ミス/;

function dupNormName_(name) {
  let s = String(name == null ? "" : name).normalize("NFKC").replace(/[\s　]/g, "");
  s = s.replace(/[ァ-ヶ]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0x60); });
  s = s.replace(/(様|さま|さん)$/, "");
  return s.toLowerCase();
}

function dupPhone_(v) {
  return String(v == null ? "" : v).normalize("NFKC").replace(/\D/g, "").replace(/^81/, "0");
}

// 承認の強さ。高いほうを残す。
function dupApprovalRank_(v) {
  const s = String(v == null ? "" : v).trim();
  if (!s) return 0;
  const f = getAdvertiserApprovalFlags(s);
  if (f.approved) return 3;
  if (DUP_MEMO_PATTERN_.test(s)) return 0;   // 重複メモは承認の情報ではない
  if (f.trackingMissing) return 1;
  return 2;                                 // 🔺 など
}

function dupIsBlank_(v) { return v === "" || v === null || v === undefined; }

// 見出しから列の位置を引く
function dupColumns_(headers) {
  const h = headers.map(String);
  return {
    name: h.indexOf("お名前"),
    ref:  h.indexOf("紹介者名"),
    recv: h.indexOf("受信日時"),
    appr: h.indexOf("承認"),
    shot: h.lastIndexOf("スクショURL"),
    code: h.indexOf(FORM_CODE_HEADER),
    phones: h.map(function (x, i) { return /電話/.test(x) ? i : -1; }).filter(function (i) { return i >= 0; })
  };
}

// 2行が同じ人か（名前は呼び出し側で一致済み）。電話が両方あって違えば別人。
function dupSamePerson_(cols, a, b) {
  for (let k = 0; k < cols.phones.length; k++) {
    const i = cols.phones[k];
    const pa = dupPhone_(a[i]), pb = dupPhone_(b[i]);
    if (pa && pb && pa !== pb) return false;
  }
  return true;
}

function dupReadAgencyCode_(sheet) {
  const v = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 40), 2).getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]) === AGENCY_KEY) return String(v[i][1] || "").trim() || AGENCY_DEFAULT;
  }
  return AGENCY_DEFAULT;
}

function dupArchiveSheet_(ss) {
  let sh = ss.getSheetByName(DUP_ARCHIVE_TAB_);
  if (!sh) {
    sh = ss.insertSheet(DUP_ARCHIVE_TAB_);
    sh.getRange(1, 1, 1, DUP_ARCHIVE_HEADERS_.length).setValues([DUP_ARCHIVE_HEADERS_])
      .setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  return sh;
}

function dupArchiveRows_(ss, entries) {
  if (!entries.length) return;
  const sh = dupArchiveSheet_(ss);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, entries.length, DUP_ARCHIVE_HEADERS_.length).setNumberFormat("@").setValues(entries);
}

function dupArchiveEntry_(route, sheet, cols, headers, row, keptRecv) {
  const obj = {};
  headers.forEach(function (h, i) { if (String(h)) obj[String(h)] = row[i] instanceof Date ? formatJST(row[i]) : row[i]; });
  const s = function (i) { return i >= 0 ? (row[i] instanceof Date ? formatJST(row[i]) : String(row[i] == null ? "" : row[i])) : ""; };
  return [formatJST(new Date()), route, sheet.getName(), s(cols.code), s(cols.name), s(cols.ref),
    s(cols.recv), keptRecv instanceof Date ? formatJST(keptRecv) : String(keptRecv || ""), s(cols.appr), JSON.stringify(obj)];
}

// ---- 申請時: 既に同じ人の行があればそこへまとめる ----
// 戻り値: null = 重複なし（通常どおり行を足す） / {row, keptRecv, referrer} = まとめた
function dupMergeIncoming_(ss, sheet, formCode, rowData, config) {
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < ANSWER_START_COL) return null;
  const width = lastCol - ANSWER_START_COL + 1;
  const headers = sheet.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
  const cols = dupColumns_(headers);
  if (cols.name < 0) return null;

  // 新しい申請を、シートの見出しの並びへ置き直す
  const newHeaders = buildHeaders(config);
  const incoming = headers.map(function (h) {
    const j = newHeaders.indexOf(h);
    return j >= 0 ? rowData[j] : "";
  });
  const key = dupNormName_(incoming[cols.name]);
  if (!key) return null;

  const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, width).getValues();
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (dupNormName_(r[cols.name]) !== key) continue;
    if (!dupSamePerson_(cols, r, incoming)) continue;
    const rowNo = i + 2;
    // 残す行の空欄だけを埋める（承認と案件記号は触らない）
    incoming.forEach(function (v, c) {
      if (c === cols.appr || c === cols.code || dupIsBlank_(v) || !dupIsBlank_(r[c])) return;
      sheet.getRange(rowNo, ANSWER_START_COL + c).setValue(v);
    });
    dupArchiveRows_(ss, [dupArchiveEntry_("申請時", sheet, cols, headers, incoming, r[cols.recv])]);
    return { row: rowNo, keptRecv: r[cols.recv], referrer: cols.ref >= 0 ? String(r[cols.ref] || "") : "" };
  }
  return null;
}

function dupIncomingLineMessage_(caseName, name, referrer, keptRecv) {
  return ["【重複申請を1件にまとめました】" + caseName,
    "お名前: " + name,
    "紹介者: " + (referrer || "（なし）"),
    "最初の申請: " + (keptRecv instanceof Date ? formatJST(keptRecv) : String(keptRecv || "")),
    "同じ方の同じ案件への2回目の申請です。件数・集計には1件として数えます。"].join("\n");
}

// ---- 既存分をまとめる ----
// dryRun=true なら何も書かずに件数だけ返す。
function mergeDuplicateApplicationsCore_(dryRun) {
  const ss = getOrCreateSpreadsheet();
  const mainTz = ss.getSpreadsheetTimeZone ? ss.getSpreadsheetTimeZone() : "";
  const now = new Date();
  const result = { groups: 0, removed: 0, approvalRaised: 0, sheets: [], advertiserRemoved: 0, advertiserMissing: 0, details: [] };
  const advertiserTargets = []; // {month, caseName, name, removeMs:[...], keptMs, keptApproved}
  const archive = [];
  const touchedAgency = {};

  ss.getSheets().forEach(function (sheet) {
    if (!sheet.getName().startsWith(CONFIG_PREFIX)) return;
    const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 3 || lastCol < ANSWER_START_COL) return;
    const width = lastCol - ANSWER_START_COL + 1;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const cols = dupColumns_(headers);
    if (cols.name < 0) return;

    const range = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, width);
    const values = range.getValues();
    const formats = range.getNumberFormats();
    const formulas = range.getFormulas();
    if (formulas.some(function (r) { return r.some(function (f) { return f; }); })) {
      throw new Error("回答欄に数式があるため止めました: " + sheet.getName());
    }

    // 名前ごとにまとめる（行の順＝申請順。受信日時が読めれば早い順に並べ替える）
    const byKey = {};
    values.forEach(function (r, i) {
      const k = dupNormName_(r[cols.name]);
      if (!k) return;
      (byKey[k] = byKey[k] || []).push(i);
    });
    const removeIdx = {};
    const formCode = getFormCodeFromSheet(sheet) || "";
    const caseName = getFormDisplayName(sheet, formCode);
    let sheetGroups = 0;

    Object.keys(byKey).forEach(function (k) {
      let idxs = byKey[k];
      if (idxs.length < 2) return;
      idxs = idxs.slice().sort(function (a, b) {
        const ma = cols.recv >= 0 ? dupTimeMs_(values[a][cols.recv], mainTz) : 0;
        const mb = cols.recv >= 0 ? dupTimeMs_(values[b][cols.recv], mainTz) : 0;
        return (ma && mb && ma !== mb) ? ma - mb : a - b;
      });
      // 電話が食い違う行は別人なので、残す行と同じ人のものだけをまとめる
      const keep = idxs[0];
      const dups = idxs.slice(1).filter(function (i) { return dupSamePerson_(cols, values[keep], values[i]); });
      if (!dups.length) return;

      const kept = values[keep];
      const beforeAppr = cols.appr >= 0 ? kept[cols.appr] : "";
      dups.forEach(function (i) {
        const d = values[i];
        d.forEach(function (v, c) {
          if (c === cols.appr || c === cols.code) return;
          if (dupIsBlank_(kept[c]) && !dupIsBlank_(v)) { kept[c] = v; formats[keep][c] = formats[i][c]; }
        });
        if (cols.appr >= 0 && dupApprovalRank_(d[cols.appr]) > dupApprovalRank_(kept[cols.appr])) {
          kept[cols.appr] = d[cols.appr];
        }
        removeIdx[i] = true;
        archive.push(dupArchiveEntry_("既存の統合", sheet, cols, headers, d, kept[cols.recv]));
      });
      if (cols.appr >= 0 && String(beforeAppr) !== String(kept[cols.appr])) result.approvalRaised++;
      sheetGroups++;
      result.groups++;
      result.removed += dups.length;
      result.details.push(sheet.getName() + " 行" + (keep + 2) + " ← 行" + dups.map(function (i) { return i + 2; }).join("・"));

      advertiserTargets.push({
        caseName: caseName,
        name: kept[cols.name],
        keptMs: cols.recv >= 0 ? dupTimeMs_(kept[cols.recv], mainTz) : 0,
        keptApproval: cols.appr >= 0 ? kept[cols.appr] : "",
        keptUrl: cols.shot >= 0 ? String(kept[cols.shot] || "") : "",
        removeMs: dups.map(function (i) { return cols.recv >= 0 ? dupTimeMs_(values[i][cols.recv], mainTz) : 0; })
      });
    });

    if (!sheetGroups) return;
    const agencyCode = dupReadAgencyCode_(sheet);
    result.sheets.push(sheet.getName() + " " + sheetGroups + "組");
    if (dryRun) return;

    // 回答欄（G列以降）だけを詰め直す。A〜F列は設定なので行ごと消さない。
    const keptValues = [], keptFormats = [];
    values.forEach(function (r, i) { if (!removeIdx[i]) { keptValues.push(r); keptFormats.push(formats[i]); } });
    const blankRow = headers.map(function () { return ""; });
    while (keptValues.length < values.length) { keptValues.push(blankRow.slice()); keptFormats.push(formats[formats.length - 1].slice()); }
    range.setNumberFormats(keptFormats);
    range.setValues(keptValues);
    touchedAgency[agencyCode + "\t" + formCode] = { code: agencyCode, form: formCode };
  });

  if (dryRun) return result;

  dupArchiveRows_(ss, archive);
  SpreadsheetApp.flush();

  // 代理店SS（各案件シートの写し）を取り直す
  Object.keys(touchedAgency).forEach(function (k) {
    const t = touchedAgency[k];
    try { syncFormSheetToAgency(ss, getOrCreateAgencySpreadsheet(t.code), t.form); }
    catch (e) { Logger.log("代理店SSの取り直しに失敗: " + k + " / " + e); }
  });

  // 広告主成果管理SSの月別シートから、外した申請の行を消す
  try { dupMergeAdvertiserSheets_(advertiserTargets, result); }
  catch (e) { Logger.log("広告主シートの統合に失敗: " + e); result.advertiserError = String(e); }

  // 生成物を作り直す（どれも 設定_ シートから作る）
  const rebuilt = [];
  [["管理シート", updateManagementSheet],
   ["申請状況一覧", buildApplicationStatusSheet],
   ["代理店別サマリー", buildAgencySummarySheet]].forEach(function (p) {
    try { p[1](); rebuilt.push(p[0]); } catch (e) { Logger.log(p[0] + " の作り直しに失敗: " + e); }
  });
  result.rebuilt = rebuilt;
  result.at = formatJST(now);
  return result;
}

// 日時の値を「JSTの瞬間(ms)」へ。広告主シートは月によって形が違う:
//  - 申請時の追記: 文字 "2026/09/01 22:49:42" がシートのタイムゾーンで日時値になる → 壁時計で読む
//  - 月指定の再生成: String(Date) の文字 "Mon Jul 27 2026 14:24:15 GMT+0900 (日本標準時)"
function dupTimeMs_(v, tz) {
  if (v instanceof Date) {
    if (tz && typeof Utilities !== "undefined" && Utilities.formatDate) {
      return aspToMillis_(Utilities.formatDate(v, tz, "yyyy/MM/dd HH:mm:ss"));
    }
    return v.getTime();
  }
  const s = String(v == null ? "" : v).trim();
  const a = aspToMillis_(s);
  if (a) return a;
  const p = Date.parse(s.replace(/\s*\(.*\)\s*$/, ""));
  return isNaN(p) ? 0 : p;
}

function dupMergeAdvertiserSheets_(targets, result) {
  if (!targets.length) return;
  const adv = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  // 広告主SSはタイムゾーンが空で返り、日時値はUTCの壁時計で入っている（2026-11-09 実測）
  const advTz = (adv.getSpreadsheetTimeZone && adv.getSpreadsheetTimeZone()) || "GMT";
  // 月のシートは申請月と違うこともある（202611 に10月分が載っている）ので、YYYYMM のシートを全部見る
  const cache = {};
  adv.getSheets().forEach(function (sh) {
    const nm = sh.getName();
    if (!/^\d{6}$/.test(nm) || sh.getLastRow() < ADVERTISER_DATA_START_ROW) return;
    const rows = sh.getRange(ADVERTISER_DATA_START_ROW, 1, sh.getLastRow() - ADVERTISER_DATA_START_ROW + 1, 7).getValues();
    cache[nm] = { sheet: sh, rows: rows, times: rows.map(function (r) { return dupTimeMs_(r[0], advTz); }), del: {} };
  });
  const near = function (t, list) {
    let best = Infinity;
    list.forEach(function (x) { if (x) best = Math.min(best, Math.abs(t - x)); });
    return best;
  };

  targets.forEach(function (t) {
    const key = dupNormName_(t.name);
    const keptHits = [], dupHits = [];
    Object.keys(cache).forEach(function (nm) {
      const m = cache[nm];
      m.rows.forEach(function (r, i) {
        if (m.del[i] || String(r[1]) !== String(t.caseName) || dupNormName_(r[2]) !== key) return;
        const at = m.times[i];
        if (!at) return;
        const dk = t.keptMs ? Math.abs(at - t.keptMs) : Infinity;
        const dd = near(at, t.removeMs);
        // 60秒以内で、残す申請より外す申請のほうに近い行だけを消す
        if (dd <= 60000 && dd < dk) dupHits.push({ m: m, i: i });
        else if (dk <= 60000) keptHits.push({ m: m, i: i });
      });
    });
    // 残す申請の行がどこにも無い（前の版で取り違えて消した）ときは、外す行の1つを残す申請の内容に直して残す
    if (!keptHits.length && dupHits.length && t.keptMs) {
      const fix = dupHits.shift();
      fix.m.sheet.getRange(ADVERTISER_DATA_START_ROW + fix.i, 1).setValue(formatJST(new Date(t.keptMs)));
      if (t.keptUrl) fix.m.sheet.getRange(ADVERTISER_DATA_START_ROW + fix.i, 5).setValue(t.keptUrl);
      keptHits.push(fix);
      result.advertiserFixed = (result.advertiserFixed || 0) + 1;
    }
    dupHits.forEach(function (h) { h.m.del[h.i] = true; result.advertiserRemoved++; });
    result.advertiserMissing += Math.max(0, t.removeMs.length - dupHits.length);
    // 残した行の承認を、まとめた後の承認にそろえる
    const f = getAdvertiserApprovalFlags(t.keptApproval);
    if (f.approved) keptHits.forEach(function (h) {
      if (h.m.rows[h.i][6] !== true) h.m.sheet.getRange(ADVERTISER_DATA_START_ROW + h.i, 6, 1, 2).setValues([[false, true]]);
    });
  });

  // 下の行から消す（上から消すと行番号がずれる）
  Object.keys(cache).forEach(function (k) {
    const m = cache[k];
    if (!m || !Object.keys(m.del).length) return;
    // フィルタがあると deleteRow が無言で失敗するので、外して消してから付け直す
    const filter = m.sheet.getFilter ? m.sheet.getFilter() : null;
    const filterA1 = filter ? filter.getRange().getA1Notation() : "";
    if (filter) filter.remove();
    Object.keys(m.del).map(Number).sort(function (a, b) { return b - a; }).forEach(function (i) {
      m.sheet.deleteRow(ADVERTISER_DATA_START_ROW + i);
    });
    SpreadsheetApp.flush();
    if (filterA1) {
      try { m.sheet.getRange(filterA1).createFilter(); } catch (e) { Logger.log("フィルタの付け直しに失敗: " + k + " / " + e); }
    }
  });
}

// ---- 既に 設定_ 側をまとめた分について、広告主シートだけを後から片づける ----
// 「重複申請_統合済み」タブ（経路=既存の統合）から対象を組み立てる。
// 消すのは「外した申請と同じ案件・同じ名前・受信日時が60秒以内」の行だけなので、2回押しても消えすぎない。
function mergeDuplicateAdvertiserFromArchive() {
  const ss = getOrCreateSpreadsheet();
  const arch = ss.getSheetByName(DUP_ARCHIVE_TAB_);
  const result = { advertiserRemoved: 0, advertiserMissing: 0, mainTz: ss.getSpreadsheetTimeZone() };
  if (!arch || arch.getLastRow() < 2) return result;
  const rows = arch.getRange(2, 1, arch.getLastRow() - 1, DUP_ARCHIVE_HEADERS_.length).getValues();
  const groups = {};
  const nameCache = {};
  rows.forEach(function (r) {
    if (String(r[1]) !== "既存の統合") return;
    const sheetName = String(r[2]);
    const sh = ss.getSheetByName(sheetName);
    if (!sh || dupReadAgencyCode_(sh) !== AGENCY_DEFAULT) return; // 広告主シートは自社分だけ
    if (nameCache[sheetName] === undefined) nameCache[sheetName] = getFormDisplayName(sh, getFormCodeFromSheet(sh) || "");
    const key = sheetName + "	" + dupNormName_(r[4]) + "	" + String(r[7]);
    const g = groups[key] = groups[key] || { caseName: nameCache[sheetName], name: r[4], keptMs: dupTimeMs_(r[7]), keptApproval: "", keptUrl: "", removeMs: [], sheet: sh };
    g.removeMs.push(dupTimeMs_(r[6]));
  });
  // 残した行のいまの承認を 設定_ シートから引く
  Object.keys(groups).forEach(function (k) {
    const g = groups[k], sh = g.sheet;
    const width = sh.getLastColumn() - ANSWER_START_COL + 1;
    const headers = sh.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const cols = dupColumns_(headers);
    if (cols.appr < 0 || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, ANSWER_START_COL, sh.getLastRow() - 1, width).getValues();
    const key = dupNormName_(g.name);
    const hit = data.find(function (r) { return dupNormName_(r[cols.name]) === key && Math.abs(dupTimeMs_(r[cols.recv], result.mainTz) - g.keptMs) <= 60000; });
    if (hit) { g.keptApproval = hit[cols.appr]; g.keptUrl = cols.shot >= 0 ? String(hit[cols.shot] || "") : ""; }
  });
  const targets = Object.keys(groups).map(function (k) { return groups[k]; });
  dupMergeAdvertiserSheets_(targets, result);
  const msg = "広告主シートの重複 " + result.advertiserRemoved + "行を削除・取り違えの直し " + (result.advertiserFixed || 0) + "・見つからず " + result.advertiserMissing + "（設定_側は統合済み）";
  Logger.log(msg + " tz=" + result.mainTz);
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg + " 完了", "重複申請", 60); } catch (e) {}
  return result;
}

// ---- メニュー ----
function previewDuplicateApplicationsFromMenu() {
  const r = mergeDuplicateApplicationsCore_(true);
  const msg = "重複 " + r.groups + "組・外す行 " + r.removed + "件（下見・まだ何も変えていません）";
  Logger.log(msg + "\n" + r.details.join("\n"));
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "重複申請の下見", 30);
  return r;
}

function mergeDuplicateApplications() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const r = mergeDuplicateApplicationsCore_(false);
    const msg = "重複申請の統合 完了: " + r.groups + "組・" + r.removed + "行を1件へ（承認を引き継いだ組 " + r.approvalRaised +
      "）。広告主シート " + r.advertiserRemoved + "行削除・未発見 " + r.advertiserMissing + "。作り直し: " + (r.rebuilt || []).join("・");
    Logger.log(msg + "\n" + r.details.join("\n"));
    try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, "重複申請", 60); } catch (e) {}
    return r;
  } finally {
    lock.releaseLock();
  }
}
