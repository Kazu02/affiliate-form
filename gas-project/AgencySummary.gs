// =============================================
// 代理店別サマリー（件数・承認率・稼働状況の数値化）
// （2026-08-20 追加）
//
// 「代理店ごとの件数・稼働率の数値化」への対応。申請状況一覧と同じ元データ
// （collectApplications_）から集計するので、二重に数え方を持たない。
// 生成物なので手で編集しない。
// =============================================

const AGENCY_SUMMARY_SHEET = "代理店別サマリー";
const RECENT_WINDOW_DAYS = 30;
const JISHA_LABEL = "自社";

function getAgencySummarySheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(AGENCY_SUMMARY_SHEET);
  if (!sh) sh = ss.insertSheet(AGENCY_SUMMARY_SHEET, 5);
  return sh;
}

// 申請1件を「自社」か代理店名かに振り分ける
function partyOf_(row) {
  const ag = String(row.values[4] || "").trim();
  return ag || JISHA_LABEL;
}

function buildAgencySummarySheet() {
  const sh = getAgencySummarySheet_();
  const rows = collectApplications_();
  const now = new Date().getTime();
  const recentFrom = now - RECENT_WINDOW_DAYS * 24 * 60 * 60 * 1000;

  // ---- 区分ごとの集計 ----
  const stat = {};   // party -> {n, approved, missing, pending, recent, first, last}
  const byCase = {}; // caseName -> {party -> n}
  const partySet = {};
  const caseOrder = [];

  rows.forEach(function (r) {
    const party = partyOf_(r);
    partySet[party] = true;
    if (!stat[party]) {
      stat[party] = { n: 0, approved: 0, missing: 0, pending: 0, recent: 0, first: 0, last: 0 };
    }
    const s = stat[party];
    s.n++;

    const flags = getAdvertiserApprovalFlags(r.values[5]);
    if (flags.approved) s.approved++;
    else if (flags.trackingMissing) s.missing++;
    else s.pending++;

    if (r.sortKey) {
      if (r.sortKey >= recentFrom) s.recent++;
      if (!s.first || r.sortKey < s.first) s.first = r.sortKey;
      if (r.sortKey > s.last) s.last = r.sortKey;
    }

    const caseName = String(r.values[1] || "");
    if (!byCase[caseName]) { byCase[caseName] = {}; caseOrder.push(caseName); }
    byCase[caseName][party] = (byCase[caseName][party] || 0) + 1;
  });

  // 区分の並び: 自社を先頭、あとは申請数の多い順
  const parties = Object.keys(partySet).sort(function (a, b) {
    if (a === JISHA_LABEL) return -1;
    if (b === JISHA_LABEL) return 1;
    return stat[b].n - stat[a].n;
  });

  // 代理店マスタにあるが申請ゼロの代理店も行として出す（稼働していないことが分かるように）
  readAgencies_().forEach(function (a) {
    if (!a.name || partySet[a.name]) return;
    partySet[a.name] = true;
    stat[a.name] = { n: 0, approved: 0, missing: 0, pending: 0, recent: 0, first: 0, last: 0 };
    parties.push(a.name);
  });

  sh.clear();

  const fmt = function (ms) { return ms ? formatJST(new Date(ms - 9 * 60 * 60 * 1000)).slice(0, 10) : ""; };

  // ---- ブロック1: 区分別 ----
  const h1 = ["区分", "申請数", "承認", "トラッキング漏れ", "未判定", "承認率",
              "直近" + RECENT_WINDOW_DAYS + "日", "初回申請", "直近申請"];
  const r1 = sh.getRange(1, 1, 1, h1.length);
  r1.setValues([h1]);
  r1.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff").setWrap(true);

  const body1 = parties.map(function (p) {
    const s = stat[p];
    const rate = s.n ? Math.round((s.approved / s.n) * 1000) / 10 : 0;
    return [p, s.n, s.approved, s.missing, s.pending, s.n ? rate + "%" : "-",
            s.recent, fmt(s.first), fmt(s.last)];
  });
  // 合計行
  const tot = parties.reduce(function (acc, p) {
    const s = stat[p];
    acc.n += s.n; acc.a += s.approved; acc.m += s.missing; acc.p += s.pending; acc.r += s.recent;
    return acc;
  }, { n: 0, a: 0, m: 0, p: 0, r: 0 });
  body1.push(["合計", tot.n, tot.a, tot.m, tot.p,
              tot.n ? (Math.round((tot.a / tot.n) * 1000) / 10) + "%" : "-", tot.r, "", ""]);

  if (body1.length) {
    sh.getRange(2, 1, body1.length, h1.length).setValues(body1);
    sh.getRange(2 + body1.length - 1, 1, 1, h1.length)
      .setFontWeight("bold").setBackground("#eef2ff");
  }

  // ---- ブロック2: 案件 × 区分 の申請数 ----
  const start = 2 + body1.length + 2;
  sh.getRange(start - 1, 1).setValue("案件別の申請数").setFontWeight("bold");

  const h2 = ["案件名"].concat(parties).concat(["合計"]);
  const r2 = sh.getRange(start, 1, 1, h2.length);
  r2.setValues([h2]);
  r2.setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff").setWrap(true);

  const body2 = caseOrder.map(function (cn) {
    const row = [cn];
    let sum = 0;
    parties.forEach(function (p) {
      const v = (byCase[cn] && byCase[cn][p]) || 0;
      row.push(v); sum += v;
    });
    row.push(sum);
    return row;
  });
  body2.sort(function (a, b) { return b[b.length - 1] - a[a.length - 1]; }); // 合計の多い順

  if (body2.length) {
    sh.getRange(start + 1, 1, body2.length, h2.length).setValues(body2);
  }

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 240);
  for (let i = 2; i <= Math.max(h1.length, h2.length); i++) sh.setColumnWidth(i, 110);

  return { parties: parties.length, applications: tot.n, cases: body2.length };
}

function buildAgencySummarySheetFromMenu() {
  const r = buildAgencySummarySheet();
  SpreadsheetApp.getUi().alert(
    "代理店別サマリーを作り直しました。" + "\n\n" +
    "区分: " + r.parties + " 件（自社＋代理店）" + "\n" +
    "申請: " + r.applications + " 件" + "\n" +
    "案件: " + r.cases + " 件" + "\n\n" +
    "このシートは生成物です。手で編集しても次の再生成で消えます。"
  );
}

// 申請状況一覧と同じ日次トリガーで一緒に作り直す
function rebuildAgencyReports() {
  const a = buildApplicationStatusSheet();
  const b = buildAgencySummarySheet();
  Logger.log("申請状況一覧: " + a.count + " 件 / 代理店別サマリー: " + b.parties + " 区分");
  return { status: a, summary: b };
}
