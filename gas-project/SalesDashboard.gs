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
  const hits = sdReadReps_().filter(function (r) { return r.token && r.token === key; });
  // **同じ合言葉が2行にあると誰なのか決まらない。** 先頭を採ると、後の行の人のURLで
  // 先頭の人の顧客が見えてしまう。決まらないときは通さない（人が直すまで開けない）。
  if (hits.length !== 1) return null;
  return hits[0];
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
    // 途中で合言葉や状態が欠けた行を埋め直す（URLは合言葉から必ず導ける）。
    // **16文字未満は身元確認で弾かれる**ので、空と同じく作り直す
    // （そのままだと「URLはあるのに開けない」が続く）。
    let token = cur.token;
    if (!token || token.length < 16) {
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

  // 同じ合言葉が2行以上にあると、その合言葉のURLは（安全側に倒して）開けなくなる。
  // 手でシートを編集したときにしか起きないが、起きたら気付けるようにする。
  const after = sdReadReps_();
  const seen = {}, dup = {};
  after.forEach(function (r) {
    if (!r.token) return;
    if (seen[r.token]) dup[r.token] = true;
    seen[r.token] = true;
  });
  const dupNames = after
    .filter(function (r) { return r.token && dup[r.token]; })
    .map(function (r) { return r.name; });

  return { added: added, fixed: fixed, total: after.length, duplicated: dupNames };
}

function sdUrlFor_(token) {
  return SD_PAGE_URL + "?k=" + encodeURIComponent(token);
}

// メニュー: 全員ぶんのURLを揃えて、コピーできる形で見せる。
function showSalesDashboardUrls() {
  const r = syncSalesDashboard();
  const list = sdReadReps_();

  // 名簿（JISHA_REFERRER_OPTIONS）から外れているのに稼働のままの行を目立たせる。
  // **勝手に止めない。** 名簿に無くても実在する紹介者がいるため（例: 萩原愛也）。
  // ただし退職者の行がそのままだと、その人のURLで顧客一覧が見え続けるので必ず気付かせる。
  const roster = {};
  JISHA_REFERRER_OPTIONS.split(",").forEach(function (nm) {
    const t = nm.trim();
    if (t) roster[t] = true;
  });
  let offRoster = 0;

  let rows = "";
  list.forEach(function (rep) {
    const active = rep.status === SD_STATUS_ACTIVE;
    const stray  = active && !roster[rep.name];
    if (stray) offRoster++;
    rows +=
      '<tr' + (stray ? ' class="stray"' : '') + '>' +
      '<td class="nm">' + escapeHtml_(rep.name) +
        (active ? "" : '<span class="off">停止中</span>') +
        (stray ? '<span class="off">名簿にありません</span>' : '') + '</td>' +
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
    'tr.stray td{background:#fffbeb;}' +
    '</style>' +
    '<p>営業ひとりに1本のURLです。<b>本人にだけ</b>LINEなどで送り、ブックマークしてもらってください。</p>' +
    '<div class="warn">このURLを知っている人は、その営業の顧客一覧を見られます。' +
    '本人以外へ転送しないよう伝えてください。渡し間違えたときは、営業マスタの合言葉の欄を' +
    '空にしてこのメニューをもう一度実行すると、新しいURLに変わります（古いURLは使えなくなります）。</div>' +
    (r.duplicated && r.duplicated.length
      ? '<div class="warn"><b>合言葉が重複しています: ' + escapeHtml_(r.duplicated.join("、")) + '</b>。' +
        'この方たちのURLは安全のため開けません（誰のURLか決まらないため）。' +
        '営業マスタで重複している合言葉の欄を<b>空にして</b>、このメニューをもう一度実行してください。</div>'
      : '') +
    (offRoster
      ? '<div class="warn"><b>名簿に無い担当が ' + offRoster + ' 名います（黄色の行）。</b>' +
        '退職・異動した方であれば、営業マスタのその行の「状態」を <b>停止</b> に変えてください。' +
        'そのままだと、その方のURLで顧客一覧が見え続けます。' +
        '名簿に無いだけで在籍している方（代理店経由の紹介者など）は、そのままで構いません。</div>'
      : '') +
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

// 紹介者名 → 正規名。**1行ごとに引くので必ず覚えておく。**
// `agCanonicalRep_()` は呼ぶたびに対応表（名簿10名＋別名20件）を作り直すので、
// 申請状況一覧と承認漏れ管理で計1600行を素で回すと正規表現が十数万回走る。
// この案件は「設定シートの全行スキャンで56秒」の前科があるので、
// 営業が毎回待たされる経路では先に潰しておく。紹介者名は同じ値が何度も出るので
// 覚えておくだけでほぼ全部が当たる。
let SD_REP_MEMO = null;
function sdCanonRep_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (!SD_REP_MEMO) SD_REP_MEMO = {};
  if (SD_REP_MEMO[s] !== undefined) return SD_REP_MEMO[s];
  const v = agCanonicalRep_(s);
  SD_REP_MEMO[s] = v;
  return v;
}

// **営業が書いた文字列をシートへ入れる前に必ず通す。**
// `=` や `+` で始まる文字列は Google Sheets が数式として保存する。
// メモは画面へ読み戻すので、`=営業マスタ!B2` と書けば**他人の合言葉が読める**。
// 先頭に `'` を付けるとシートは「文字列として扱え」と解釈し、読むときは付けた `'` を返さない。
function sdPlainText_(v) {
  const s = String(v == null ? "" : v);
  return /^[=+]/.test(s) ? "'" + s : s;
}

// スクショURLは「保存エラー: ...」のような文字列が入っていることがある。
// リンクとして出すと押しても何も起きないので、httpで始まるものだけURLとして扱う。
function sdSafeUrl_(v) {
  const s = String(v || "").trim();
  return /^https?:\/\//i.test(s) ? s : "";
}

// 申請状況一覧から、その営業の申請を拾って顧客ごとにまとめる。
// 代理店経由（紹介者が「顧客名（担当者名）」形式）も担当営業の預かりなので含める。
function sdCustomers_(repName, problems) {
  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(APP_STATUS_SHEET);
  if (!sh) { if (problems) problems.push(APP_STATUS_SHEET); return []; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, APP_STATUS_HEADERS.length).getValues();
  const byCust = {};
  const order = [];

  vals.forEach(function (r) {
    const custName = String(r[2] || "").trim();
    if (!custName) return;
    if (sdCanonRep_(r[3]) !== repName) return;

    const flags = getAdvertiserApprovalFlags(r[5]);
    const status = flags.approved ? "承認" : (flags.trackingMissing ? "非承認" : "確認中");
    // **並び順と月の集計は文字列でなく時刻で決める。**
    // 受信日時は手入力や移入で `2026-08-26 10:00:00` のような書き方が混ざりうる。
    // 文字列のまま比べると、ハイフンの行が最新にならず、今月の集計からも落ちる。
    const recvMs = aspToMillis_(r[0]);
    const recvAt = toDisplayDate_(r[0]);

    const key = normalizeName(custName);
    if (!byCust[key]) {
      byCust[key] = { name: custName, lastRecvAt: recvAt, lastMs: recvMs, memo: "", cases: [] };
      order.push(key);
    }
    const c = byCust[key];
    if (recvMs > c.lastMs) { c.lastMs = recvMs; c.lastRecvAt = recvAt; }
    c.cases.push({
      caseName:   String(r[1] || ""),
      recvAt:     recvAt,
      recvMs:     recvMs,
      status:     status,
      agencyName: String(r[4] || ""),
      shotUrl:    sdSafeUrl_(r[6])
    });
  });

  const memos = sdReadMemos_(repName);
  const out = order.map(function (k) {
    const c = byCust[k];
    c.memo = memos[k] || "";
    c.cases.sort(function (a, b) { return b.recvMs - a.recvMs; });
    return c;
  });
  out.sort(function (a, b) { return b.lastMs - a.lastMs; });
  return out;
}

// 承認漏れ管理から、その営業がまだ答えていない確認依頼を拾う。
// 対象は pushSalesApprovalChecks と同じ条件（型B・型C／リンクが生きている／未確認）。
function sdChecks_(repName, problems) {
  const ss = getOrCreateSpreadsheet();
  const sh = ss.getSheetByName(AG_MANAGE_SHEET);
  if (!sh) { if (problems) problems.push(AG_MANAGE_SHEET); return []; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();
  const out = [];
  vals.forEach(function (r) {
    if (!sdIsOpenCheck_(r)) return;
    if (sdCanonRep_(r[AGC_REF - 1]) !== repName) return;

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

// 稼働中の案件のうち、その顧客がまだ申し込んでいないものを各顧客へ付ける。
//
// **停止した案件は出さない。** 出すと「もう受け付けていないもの」を勧めることになり、
// 顧客は「このキャンペーンは終了しました」に着地する。営業の信用に直接響く。
// 稼働の正は案件マスタのチェックボックスなので、そこだけを見る。
//
// 案件名の突き合わせは normalizeName で寄せる。案件マスタ側と申請状況一覧側で
// スペースや全角半角が揺れても、同じ案件を「未実施」に数えないようにするため。
function sdAttachTodo_(customers, activeCases) {
  customers.forEach(function (c) {
    const done = {};
    c.cases.forEach(function (k) { done[normalizeName(k.caseName)] = true; });
    c.todo = activeCases
      .filter(function (a) { return !done[normalizeName(a.name)]; })
      .map(function (a) { return a.name; });
  });
}

function salesDashboardPayload_(token) {
  const rep = sdFindByToken_(token);
  if (!rep) {
    return { error: "このURLでは開けません。担当者からお送りしたURLをそのまま開いてください。" };
  }
  if (rep.status !== SD_STATUS_ACTIVE) {
    return { error: "このURLは現在ご利用いただけません。担当者へご連絡ください。" };
  }

  // 読む先のシートが無いと「0件」と区別が付かない。**黙って0件にしない。**
  const problems  = [];
  const customers = sdCustomers_(rep.name, problems);
  const checks    = sdChecks_(rep.name, problems);
  const thisMonth = formatJST(new Date()).substring(0, 7); // yyyy/MM

  // 稼働中の案件。「この顧客がまだやっていないもの」を出すために使う。
  // **読めなかったときは todo を付けない。** 付けてしまうと、実際には稼働案件が
  // あるのに「未実施なし」と読める画面になり、0件と取得不能の区別が付かなくなる。
  let activeCases = [];
  let todoReady = false;
  try {
    activeCases = listActiveCases_();
    todoReady = true;
  } catch (err) {
    Logger.log("salesDashboardPayload_: 稼働案件を読めませんでした: " + err);
  }
  if (todoReady) sdAttachTodo_(customers, activeCases);

  let applied = 0, approved = 0, waiting = 0;
  customers.forEach(function (c) {
    c.cases.forEach(function (k) {
      // 受信日時の書き方が揺れても落ちないよう、時刻から月を作り直して比べる
      const ym = k.recvMs ? formatJST(new Date(k.recvMs)).substring(0, 7) : String(k.recvAt).substring(0, 7);
      if (ym === thisMonth) {
        applied++;
        if (k.status === "承認") approved++;
      }
      if (k.status === "確認中") waiting++;
    });
  });

  return {
    repName: rep.name,
    updatedAt: formatJST(new Date()),
    // 顧客の公式LINE追加の確認（SalesLineReview.gs）を出すか。
    // 設定が無いときはタブごと出さない。中身は `?sales_line=` で別に読む。
    lineEnabled: slrEnabled_(),
    // 画面はこれが入っていたら「0件」ではなく異常として出す
    notice: problems.length
      ? "いま一覧を読み込めませんでした。0件ではなく、システム側の不具合です。事務までご連絡ください。"
      : "",
    // 未実施の案件を出せる状態か。false のときは画面側もその節を出さない
    todoReady: todoReady,
    summary: {
      needCheck:      checks.length,
      appliedMonth:   applied,
      approvedMonth:  approved,
      waiting:        waiting,
      customerCount:  customers.length,
      activeCaseCount: activeCases.length
    },
    checks: checks,
    customers: customers
  };
}

// ---------------------------------------------------------------
// 画面からの書き込み
// ---------------------------------------------------------------

// 画面へ出す対象かどうか。**読むときと書くときで必ず同じ判定を使う。**
// キーだけで当てて書くと、画面に出していない行（同じ案件・顧客・同じ分の型Dなど）まで
// 「出す」に変わり、広告主への依頼へ載りうる（Codexの指摘7）。
function sdIsOpenCheck_(r) {
  const type = String(r[AGC_TYPE - 1] || "");
  if (type !== "B" && type !== "C") return false;
  if (String(r[AGC_LINK - 1] || "") === "リンク切れ疑い") return false;
  const chk = String(r[AGC_SALESCHK - 1] || "").trim();
  return !chk || chk === "未確認";
}

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

  // 二重に押したとき・別の人と同時のときに、読み→書きの間へ割り込ませない。
  // **`getDocumentLock()` はWebアプリでは null を返す**ので使えない（公式仕様）。
  // 取れなければ書かずに引き返す（黙って壊すより、もう一度押してもらう方がよい）。
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return { result: "error", message: "いま混み合っています。少しおいて、もう一度押してください。" };
    }
  } catch (e) {
    return { result: "error", message: "いま混み合っています。少しおいて、もう一度押してください。" };
  }
  try {

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { result: "error", message: "対象が見つかりませんでした。" };

  const vals = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();
  let updated = 0;
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    if (agKeyOf_(r[AGC_CASE - 1], r[AGC_NAME - 1], aspToMillis_(r[AGC_RECV - 1])) !== key) continue;
    if (sdCanonRep_(r[AGC_REF - 1]) !== rep.name) continue;
    // 画面に出していない行・すでに答えた行は書き換えない。
    // これで「先に答えた方が勝ち」になり、古い画面から上書きされない。
    if (!sdIsOpenCheck_(r)) continue;
    const row = i + 2;
    // 営業確認・営業コメント・依頼可否は隣り合う3列なので**1回で書く**。
    // 1セルずつ書くと、同じ人が二重に送ったときに書き込みが交差して
    // 「取下げ なのに 出す」のような組み合わせが残りうる（そのまま広告主へ出てしまう）。
    sh.getRange(row, AGC_SALESCHK, 1, 3).setValues([[
      answer,
      comment ? sdPlainText_(comment) : String(r[AGC_SALESCMT - 1] || ""),
      answer === "OK" ? "出す" : "出さない"   // OK なら広告主へ出せる。要再取得・取下げは出さない
    ]]);
    updated++;
  }
  if (!updated) {
    // 誰か（自分の別の画面を含む）が先に答えた場合もここへ来る。
    return { result: "error", message: "この件はすでに回答済みか、対象から外れています。「最新の状態にする」を押してご確認ください。" };
  }
  SpreadsheetApp.flush();
  return { result: "ok", updated: updated };

  } finally {
    try { lock.releaseLock(); } catch (e) { /* すでに解放されていれば何もしない */ }
  }
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
  const range = sh.getRange(2, 1, lastRow - 1, SD_MEMO_HEADERS.length);
  const vals = range.getValues();
  // **式が入っている行は中身を返さない。** 書き込み側で防いでいるが、
  // 人が手でシートへ式を入れた場合にも他のセルの中身が画面へ出ないようにする。
  const fmls = range.getFormulas();
  vals.forEach(function (r, i) {
    if (String(r[SDM_REP - 1] || "").trim() !== repName) return;
    const cust = String(r[SDM_CUST - 1] || "").trim();
    if (!cust) return;
    if (fmls[i][SDM_MEMO - 1]) { out[normalizeName(cust)] = ""; return; }
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
      sh.getRange(row, SDM_MEMO, 1, 2).setValues([[sdPlainText_(memo), formatJST(new Date())]]);
      SpreadsheetApp.flush();
      return { result: "ok" };
    }
  }
  sh.appendRow([rep.name, sdPlainText_(cust), sdPlainText_(memo), formatJST(new Date())]);
  SpreadsheetApp.flush();
  return { result: "ok" };
}
