// =============================================
// 営業ダッシュボード（営業マン用の1枚画面）
// （2026-08-27 追加）
//
// 営業が普段やることは2つしかない。
//   (1) 承認確認 … 「このスクショで承認してもらえるか」を見て3択で答える
//   (2) 顧客管理 … 自分の顧客がどの案件でいまどうなっているかを見る／メモを残す
//
// これまでは SS2 の `承認確認_<担当>` タブと SS1 の `総合_<担当>` タブを開いて
// もらっていたが、**スプレッドシートは営業には難しい**（タブが多い・列が多い・
// どこを触っていいか分からない・スマホで開けない）。
// 代理店リンク集（links.html）と同じ形＝**個人ごとのURLを1本渡すだけ**にする。
//
// 設計上の約束:
//  - 画面には営業が触っていい欄しか出さない。照合キーのような機械の都合は出さない。
//  - 正データは 承認漏れ管理（メインSS）。この画面はそこへ直接書く。
//    SS2 のタブ経由（pushSalesApprovalChecks → pullSalesApprovalChecks）は
//    残してあるが、**先に答えた方が勝ち**でぶつからない
//    （push は「未確認」以外を再依頼しない／pull は「未確認」を読み飛ばす）。
//  - 一覧は 申請状況一覧（日次生成＋申請時に2行目へ差し込み）から引く。
//    設定タブを全部なめると 56 秒かかった実績があるため、案件シートは読まない。
//  - **このリポジトリは public。** 合言葉（トークン）はコードに書かず営業マスタに置く。
// =============================================

const SD_MASTER_SHEET  = "営業マスタ";
const SD_MEMO_SHEET    = "営業メモ";
const SD_PAGE_URL      = "https://kazu02.github.io/affiliate-form/sales.html";
const SD_STATUS_ACTIVE = "稼働";

const SD_HEADERS = ["営業担当", "合言葉（さわらないでください）", "状態", "発行日", "ダッシュボードURL"];
const SDC_NAME = 1, SDC_TOKEN = 2, SDC_STATUS = 3, SDC_ISSUED = 4, SDC_URL = 5;

const SD_MEMO_HEADERS = ["営業担当", "顧客名", "メモ", "更新日時"];
const SDM_REP = 1, SDM_CUST = 2, SDM_MEMO = 3, SDM_AT = 4;

// 画面から送られてくる答え。承認漏れ管理の「営業確認」に入る値と同じにする。
const SD_ANSWERS = ["OK", "要再取得", "取下げ"];

// メモの長さの上限。公開経路なので青天井にしない。
const SD_MEMO_MAX = 1000;

// ---------------------------------------------------------------
// 営業マスタ（担当者ごとの合言葉とURL）
// ---------------------------------------------------------------

function sdMasterSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(SD_MASTER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SD_MASTER_SHEET, 3);
    const h = sh.getRange(1, 1, 1, SD_HEADERS.length);
    h.setValues([SD_HEADERS]);
    h.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(SDC_NAME, 140);
    sh.setColumnWidth(SDC_TOKEN, 260);
    sh.setColumnWidth(SDC_STATUS, 80);
    sh.setColumnWidth(SDC_ISSUED, 170);
    sh.setColumnWidth(SDC_URL, 420);
  }
  return sh;
}

function sdReadReps_() {
  const sh = sdMasterSheet_();
  const out = [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  sh.getRange(2, 1, lastRow - 1, SD_HEADERS.length).getValues().forEach(function (r, i) {
    const name = String(r[SDC_NAME - 1] || "").trim();
    if (!name) return;
    out.push({
      row:    i + 2,
      name:   name,
      token:  String(r[SDC_TOKEN - 1] || "").trim(),
      status: String(r[SDC_STATUS - 1] || "").trim(),
      url:    String(r[SDC_URL - 1] || "").trim()
    });
  });
  return out;
}

// 合言葉から担当者を引く。**これが唯一の身元確認**なので、
// 空文字が空欄の行に当たらないよう token が空の行は除外する。
function sdFindByToken_(token) {
  const key = String(token || "").trim();
  if (key.length < 16) return null;
  const list = sdReadReps_();
  for (let i = 0; i < list.length; i++) {
    if (list[i].token && list[i].token === key) return list[i];
  }
  return null;
}

// 名簿（JISHA_REFERRER_OPTIONS）の全員ぶんの行・合言葉・URLを揃える。
// **非破壊**。既にある合言葉は作り直さない（配ったURLが死ぬため）。
function syncSalesDashboard() {
  const sh = sdMasterSheet_();
  const existing = sdReadReps_();
  const byName = {};
  existing.forEach(function (r) { byName[r.name] = r; });

  const roster = JISHA_REFERRER_OPTIONS.split(",")
    .map(function (s) { return s.trim(); })
    .filter(Boolean);

  let added = 0, fixed = 0;
  roster.forEach(function (name) {
    const cur = byName[name];
    if (!cur) {
      const token = Utilities.getUuid().replace(/-/g, "");
      sh.appendRow([name, token, SD_STATUS_ACTIVE, formatJST(new Date()), sdUrlFor_(token)]);
      added++;
      return;
    }
    // 途中で合言葉や状態が欠けた行を埋め直す（URLは合言葉から必ず導ける）
    let token = cur.token;
    if (!token) {
      token = Utilities.getUuid().replace(/-/g, "");
      sh.getRange(cur.row, SDC_TOKEN).setValue(token);
      sh.getRange(cur.row, SDC_ISSUED).setValue(formatJST(new Date()));
      fixed++;
    }
    if (!cur.status) { sh.getRange(cur.row, SDC_STATUS).setValue(SD_STATUS_ACTIVE); fixed++; }
    const want = sdUrlFor_(token);
    if (cur.url !== want) { sh.getRange(cur.row, SDC_URL).setValue(want); fixed++; }
  });

  SpreadsheetApp.flush();
  return { added: added, fixed: fixed, total: sdReadReps_().length };
}

function sdUrlFor_(token) {
  return SD_PAGE_URL + "?k=" + encodeURIComponent(token);
}

// メニュー: 全員ぶんのURLを揃えて、コピーできる形で見せる。
function showSalesDashboardUrls() {
  const r = syncSalesDashboard();
  const list = sdReadReps_();

  let rows = "";
  list.forEach(function (rep) {
    const active = rep.status === SD_STATUS_ACTIVE;
    rows +=
      '<tr>' +
      '<td class="nm">' + escapeHtml_(rep.name) + (active ? "" : '<span class="off">停止中</span>') + '</td>' +
      '<td><input type="text" readonly value="' + escapeHtml_(rep.url) + '" onclick="this.select()"></td>' +
      '<td><button type="button" onclick="cp(this)">コピー</button></td>' +
      '</tr>';
  });

  const html =
    '<style>' +
    'body{font-family:-apple-system,"Hiragino Sans","Noto Sans JP",sans-serif;font-size:13px;color:#0f172a;margin:0;padding:14px;}' +
    'p{line-height:1.7;margin:0 0 12px;}' +
    'table{width:100%;border-collapse:collapse;}' +
    'td{padding:5px 4px;border-bottom:1px solid #e2e8f0;vertical-align:middle;}' +
    'td.nm{white-space:nowrap;font-weight:700;padding-right:10px;}' +
    '.off{color:#b91c1c;font-weight:400;margin-left:6px;font-size:11px;}' +
    'input{width:100%;font-size:11px;padding:5px 6px;border:1px solid #cbd5e1;border-radius:5px;color:#334155;}' +
    'button{font-size:12px;padding:6px 10px;border:none;border-radius:5px;background:#4f46e5;color:#fff;cursor:pointer;white-space:nowrap;}' +
    'button.done{background:#16a34a;}' +
    '.warn{background:#fef2f2;border:1px solid #fca5a5;color:#991b1b;border-radius:6px;padding:10px 12px;margin-bottom:12px;line-height:1.7;}' +
    '</style>' +
    '<p>営業ひとりに1本のURLです。<b>本人にだけ</b>LINEなどで送り、ブックマークしてもらってください。</p>' +
    '<div class="warn">このURLを知っている人は、その営業の顧客一覧を見られます。' +
    '本人以外へ転送しないよう伝えてください。渡し間違えたときは、営業マスタの合言葉の欄を' +
    '空にしてこのメニューをもう一度実行すると、新しいURLに変わります（古いURLは使えなくなります）。</div>' +
    '<table>' + rows + '</table>' +
    '<script>' +
    'function cp(b){var i=b.parentNode.parentNode.querySelector("input");i.select();' +
    'document.execCommand("copy");b.textContent="コピーしました";b.className="done";' +
    'setTimeout(function(){b.textContent="コピー";b.className="";},1500);}' +
    '</script>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(680).setHeight(460),
    "営業ダッシュボードのURL（新規 " + r.added + " 件 / 修正 " + r.fixed + " 件）"
  );
}

// ---------------------------------------------------------------
// 画面へ返すデータ
// ---------------------------------------------------------------

// スクショURLは「保存エラー: ...」のような文字列が入っていることがある。
// リンクとして出すと押しても何も起きないので、httpで始まるものだけURLとして扱う。
function sdSafeUrl_(v) {
  const s = String(v || "").trim();
  return /^https?:\/\//i.test(s) ? s : "";
}

// 申請状況一覧から、その営業の申請を拾って顧客ごとにまとめる。
// 代理店経由（紹介者が「顧客名（担当者名）」形式）も担当営業の預かりなので含める。
function sdCustomers_(repName) {
  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(APP_STATUS_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, APP_STATUS_HEADERS.length).getValues();
  const byCust = {};
  const order = [];

  vals.forEach(function (r) {
    const custName = String(r[2] || "").trim();
    if (!custName) return;
    if (agCanonicalRep_(r[3]) !== repName) return;

    const flags = getAdvertiserApprovalFlags(r[5]);
    const status = flags.approved ? "承認" : (flags.trackingMissing ? "非承認" : "確認中");
    const recvAt = toDisplayDate_(r[0]);

    const key = normalizeName(custName);
    if (!byCust[key]) {
      byCust[key] = { name: custName, lastRecvAt: recvAt, memo: "", cases: [] };
      order.push(key);
    }
    const c = byCust[key];
    if (recvAt > c.lastRecvAt) c.lastRecvAt = recvAt;
    c.cases.push({
      caseName:   String(r[1] || ""),
      recvAt:     recvAt,
      status:     status,
      agencyName: String(r[4] || ""),
      shotUrl:    sdSafeUrl_(r[6])
    });
  });

  const memos = sdReadMemos_(repName);
  const out = order.map(function (k) {
    const c = byCust[k];
    c.memo = memos[k] || "";
    c.cases.sort(function (a, b) { return a.recvAt < b.recvAt ? 1 : -1; });
    return c;
  });
  out.sort(function (a, b) { return a.lastRecvAt < b.lastRecvAt ? 1 : -1; });
  return out;
}

// 承認漏れ管理から、その営業がまだ答えていない確認依頼を拾う。
// 対象は pushSalesApprovalChecks と同じ条件（型B・型C／リンクが生きている／未確認）。
function sdChecks_(repName) {
  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(AG_MANAGE_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();
  const out = [];
  vals.forEach(function (r) {
    const type = String(r[AGC_TYPE - 1] || "");
    if (type !== "B" && type !== "C") return;
    if (String(r[AGC_LINK - 1] || "") === "リンク切れ疑い") return;
    const chk = String(r[AGC_SALESCHK - 1] || "").trim();
    if (chk && chk !== "未確認") return;
    if (agCanonicalRep_(r[AGC_REF - 1]) !== repName) return;

    const shot = String(r[AGC_SHOT - 1] || "").trim();
    out.push({
      key:          agKeyOf_(r[AGC_CASE - 1], r[AGC_NAME - 1], aspToMillis_(r[AGC_RECV - 1])),
      caseName:     String(r[AGC_CASE - 1] || ""),
      customerName: String(r[AGC_NAME - 1] || ""),
      recvAt:       toDisplayDate_(r[AGC_RECV - 1]),
      days:         Number(r[AGC_DAYS - 1]) || 0,
      reward:       Number(r[AGC_REWARD - 1]) || 0,
      agencyName:   String(r[AGC_AGENCY - 1] || ""),
      shotUrl:      sdSafeUrl_(shot),
      shotBroken:   !!shot && !sdSafeUrl_(shot)
    });
  });
  out.sort(function (a, b) { return b.days - a.days; }); // 古い順＝待たせている順
  return out;
}

function salesDashboardPayload_(token) {
  const rep = sdFindByToken_(token);
  if (!rep) {
    return { error: "このURLでは開けません。担当者からお送りしたURLをそのまま開いてください。" };
  }
  if (rep.status !== SD_STATUS_ACTIVE) {
    return { error: "このURLは現在ご利用いただけません。担当者へご連絡ください。" };
  }

  const customers = sdCustomers_(rep.name);
  const checks    = sdChecks_(rep.name);
  const thisMonth = formatJST(new Date()).substring(0, 7); // yyyy/MM

  let applied = 0, approved = 0, waiting = 0;
  customers.forEach(function (c) {
    c.cases.forEach(function (k) {
      if (String(k.recvAt).substring(0, 7) === thisMonth) {
        applied++;
        if (k.status === "承認") approved++;
      }
      if (k.status === "確認中") waiting++;
    });
  });

  return {
    repName: rep.name,
    updatedAt: formatJST(new Date()),
    summary: {
      needCheck:      checks.length,
      appliedMonth:   applied,
      approvedMonth:  approved,
      waiting:        waiting,
      customerCount:  customers.length
    },
    checks: checks,
    customers: customers
  };
}

// ---------------------------------------------------------------
// 画面からの書き込み
// ---------------------------------------------------------------

// 承認確認の答えを 承認漏れ管理 へ直接書く。
// 行番号ではなく照合キーで引く（案件シートの編集で行が動くため）。
// **自分の行しか書けない**ように、キーが一致しても紹介者が本人でなければ弾く。
function handleSalesCheck_(data) {
  const rep = sdFindByToken_(data && data.token);
  if (!rep || rep.status !== SD_STATUS_ACTIVE) {
    return { result: "error", message: "このURLでは操作できません。担当者へご連絡ください。" };
  }
  const key    = String((data && data.key) || "").trim();
  const answer = String((data && data.answer) || "").trim();
  const comment = String((data && data.comment) || "").trim().substring(0, SD_MEMO_MAX);
  if (!key) return { result: "error", message: "対象が特定できませんでした。画面を開き直してください。" };
  if (SD_ANSWERS.indexOf(answer) < 0) {
    return { result: "error", message: "答えの種類が正しくありません。画面を開き直してください。" };
  }

  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(AG_MANAGE_SHEET);
  if (!sh) return { result: "error", message: "確認の一覧が見つかりません。担当者へご連絡ください。" };
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { result: "error", message: "対象が見つかりませんでした。" };

  const vals = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();
  let updated = 0;
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    if (agKeyOf_(r[AGC_CASE - 1], r[AGC_NAME - 1], aspToMillis_(r[AGC_RECV - 1])) !== key) continue;
    if (agCanonicalRep_(r[AGC_REF - 1]) !== rep.name) continue;
    const row = i + 2;
    sh.getRange(row, AGC_SALESCHK).setValue(answer);
    if (comment) sh.getRange(row, AGC_SALESCMT).setValue(comment);
    // OK なら広告主へ出せる。要再取得・取下げは出さない。
    sh.getRange(row, AGC_SENDOK).setValue(answer === "OK" ? "出す" : "出さない");
    updated++;
  }
  if (!updated) {
    return { result: "error", message: "対象が見つかりませんでした。画面を開き直してください。" };
  }
  SpreadsheetApp.flush();
  return { result: "ok", updated: updated };
}

// ---------------------------------------------------------------
// 顧客メモ
// ---------------------------------------------------------------

function sdMemoSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(SD_MEMO_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SD_MEMO_SHEET);
    const h = sh.getRange(1, 1, 1, SD_MEMO_HEADERS.length);
    h.setValues([SD_MEMO_HEADERS]);
    h.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(SDM_REP, 120);
    sh.setColumnWidth(SDM_CUST, 140);
    sh.setColumnWidth(SDM_MEMO, 520);
    sh.setColumnWidth(SDM_AT, 170);
  }
  return sh;
}

// その営業のメモを 正規化した顧客名 → メモ で返す。
function sdReadMemos_(repName) {
  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(SD_MEMO_SHEET);
  const out = {};
  if (!sh) return out;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return out;
  sh.getRange(2, 1, lastRow - 1, SD_MEMO_HEADERS.length).getValues().forEach(function (r) {
    if (String(r[SDM_REP - 1] || "").trim() !== repName) return;
    const cust = String(r[SDM_CUST - 1] || "").trim();
    if (!cust) return;
    out[normalizeName(cust)] = String(r[SDM_MEMO - 1] || "");
  });
  return out;
}

function handleSalesMemo_(data) {
  const rep = sdFindByToken_(data && data.token);
  if (!rep || rep.status !== SD_STATUS_ACTIVE) {
    return { result: "error", message: "このURLでは操作できません。担当者へご連絡ください。" };
  }
  const cust = String((data && data.customer) || "").trim();
  if (!cust) return { result: "error", message: "お客様が特定できませんでした。" };
  const memo = String((data && data.memo) || "").substring(0, SD_MEMO_MAX);

  const sh = sdMemoSheet_();
  const lastRow = sh.getLastRow();
  const key = normalizeName(cust);
  if (lastRow >= 2) {
    const vals = sh.getRange(2, 1, lastRow - 1, SD_MEMO_HEADERS.length).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][SDM_REP - 1] || "").trim() !== rep.name) continue;
      if (normalizeName(vals[i][SDM_CUST - 1]) !== key) continue;
      const row = i + 2;
      sh.getRange(row, SDM_MEMO).setValue(memo);
      sh.getRange(row, SDM_AT).setValue(formatJST(new Date()));
      SpreadsheetApp.flush();
      return { result: "ok" };
    }
  }
  sh.appendRow([rep.name, cust, memo, formatJST(new Date())]);
  SpreadsheetApp.flush();
  return { result: "ok" };
}
