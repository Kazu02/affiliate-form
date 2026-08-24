// =============================================
// 承認漏れの棚卸しと確認依頼
// （2026-08-24 追加）
//
// 設計の全文: 02_設計/承認漏れ確認フロー_20260824.md（public リポジトリのため gitignore 済み）
//
// 既存の AspReconcile.gs は「ASPが承認しているのに自社が未承認」を直す道具で、
// **ASP側の行を回している**。したがって「自社にあってASPに無い」行は原理的に出てこない。
// こちらは逆向きに、**自社の申請行を起点**にして承認が取れていないものを洗い出す。
//
// 「承認漏れ」は原因も・必要な証拠も・話す相手も違う4つが混ざっているので、最初に型を分ける。
//   型A 案件まるごと0承認    → 広告主が承認処理を回していない疑い。**営業確認をしない**
//   型B 同案件に承認実績あり  → 個別の取りこぼし。サンクスメールが効く
//   型C ASPに記録が無い      → トラッキング漏れ **または** リンク切れ
//   型D 否認                → 理由確認のみ
//
// **型Aを営業確認から外すのが要点。** 案件まるごと止まっているものを営業担当へ配っても、
// 取れるのは「スクショは正しい」だけで、承認が付かない理由は何ひとつ分からない。
//
// 受け渡しの決め（2026-08-24 ユーザー判断）:
//   - 広告主へは**まとめてスプレッドシートで渡す**（広告主成果管理SSに確認依頼タブを作る）
//   - 明細に**顧客名を載せる**（広告主は元々申込者情報を持っており、氏名が最も確実な照合キー）
//   - 営業確認は**SS2に担当別タブを書き出す**（営業担当はメインSSを編集できない）
// =============================================

const AG_MANAGE_SHEET  = "承認漏れ管理";   // メインSS。ここが正
const AG_REQUEST_SHEET = "承認確認依頼";   // 広告主成果管理SS。広告主が記入する
const AG_SALES_PREFIX  = "承認確認_";      // SS2。営業担当が記入する

// しきい値。**推測ではなく実測で決めた。**
// 2026-08-20 のCSV451件で、承認316件のうち304件（96%）が 承認日時 = 注文日時 だった。
// つまり承認は通常「注文と同時に自動で付く」ので、数日以上待っているものは待っても付かない。
// 手動で承認していると見られる3案件だけ実測の最大が48〜50日だったので、そこだけ延ばす。
const AG_DEFAULT_DAYS = 14;
const AG_SLOW_DAYS    = 60;
const AG_SLOW_CASES   = ["保険マンモス", "スマモニ", "ポケットリサーチ"];

// クリック日時の許容差。AspReconcile と同じ理由で±120秒（一致するときは中央値0秒）。
const AG_TOLERANCE_SEC = 120;

const AG_HEADERS = [
  "型", "案件名", "顧客名", "紹介者（営業）", "代理店", "受信日時", "クリック日時",
  "ASPステータス", "経過日数", "報酬額", "リンク生死", "スクショURL",
  "営業確認", "営業コメント", "依頼可否", "依頼ID", "依頼日",
  "広告主回答", "回答日", "結果", "対象シート", "対象行"
];
// 1始まりの列番号（並べ替えたら必ずここも直す）
const AGC_TYPE = 1, AGC_CASE = 2, AGC_NAME = 3, AGC_REF = 4, AGC_AGENCY = 5,
      AGC_RECV = 6, AGC_CLICK = 7, AGC_ASPST = 8, AGC_DAYS = 9, AGC_REWARD = 10,
      AGC_LINK = 11, AGC_SHOT = 12, AGC_SALESCHK = 13, AGC_SALESCMT = 14,
      AGC_SENDOK = 15, AGC_REQID = 16, AGC_REQDATE = 17,
      AGC_ADVANS = 18, AGC_ANSDATE = 19, AGC_RESULT = 20,
      AGC_SHEET = 21, AGC_ROW = 22;

const AG_SALES_CHOICES  = ["未確認", "OK", "要再取得", "取下げ"];
const AG_ADV_CHOICES    = ["承認", "否認", "該当なし", "確認中"];

// ---------------------------------------------------------------
// 共通ヘルパー
// ---------------------------------------------------------------

function agManageSheet_() {
  const ss = getOrCreateSpreadsheet();
  let sh = ss.getSheetByName(AG_MANAGE_SHEET);
  if (!sh) sh = ss.insertSheet(AG_MANAGE_SHEET);
  return sh;
}

// その案件のしきい値（日）。案件名の部分一致で低速案件を拾う。
// ASP側は「…（特別単価）」のような修飾が付くので、含むかどうかで見る。
function agThresholdDays_(caseName) {
  const c = String(caseName || "");
  for (let i = 0; i < AG_SLOW_CASES.length; i++) {
    if (c.indexOf(AG_SLOW_CASES[i]) >= 0) return AG_SLOW_DAYS;
  }
  return AG_DEFAULT_DAYS;
}

function agDaysBetween_(fromMs, toMs) {
  if (!fromMs) return "";
  return Math.floor((toMs - fromMs) / 86400000);
}

// 依頼日・回答日は「時刻を持たない日付」で入る。
// aspToMillis_ の正規表現は時刻を必須にしているので、日付だけの文字列だと 0 を返す。
// そのまま使うと「14日以上回答なし」の検出が**黙って一度も動かない**ので、日付だけも読む。
// （セルが日付書式なら Sheets が Date に直すため aspToMillis_ で足りるが、
//   テキスト書式のまま残ることがあり、そのときに検出が消える）
function agDateToMillis_(v) {
  const ms = aspToMillis_(v);
  if (ms) return ms;
  const m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})\s*$/.exec(String(v || "").trim());
  if (!m) return 0;
  return Date.UTC(+m[1], +m[2] - 1, +m[3], -9, 0, 0); // JSTの0時
}

// 履歴の引き継ぎキー。**行番号を使わない。**
// 行は案件シート側の編集で動くので、行番号で引き継ぐと別の申請の確認結果を持ち込む。
function agKeyOf_(caseName, customerName, recvMs) {
  const t = recvMs ? Math.floor(recvMs / 60000) : 0; // 分まで（秒のゆらぎを吸収）
  return String(caseName || "") + "|" + String(customerName || "").trim() + "|" + t;
}

// ---------------------------------------------------------------
// 自社の申請行を集める（AspReconcile の collectOwnClickRows_ より項目が多い）
// ---------------------------------------------------------------
function collectOwnApplicationRows_() {
  const ss = getOrCreateSpreadsheet();
  const out = [];

  listCaseSheets_(ss).forEach(function (c) {
    const sheet = c.sheet;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < ANSWER_START_COL) return;

    const width = lastCol - ANSWER_START_COL + 1;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const idx = function (label) { return headers.indexOf(label); };

    const iRecv = idx("受信日時");
    const iClick = idx("クリック日時");
    const iName = idx("お名前");
    const iRef  = idx("紹介者名");
    const iAgcy = idx(AGENCY_COLUMN_LABEL);
    const iAppr = idx("承認");
    // スクショURLは固定列で必ず項目群より後ろに来る。万一同名の項目が作られても
    // 取り違えないよう lastIndexOf で引く（2026-08-20 に同じ理由で Code.gs も直した）。
    const iShot = headers.lastIndexOf("スクショURL");
    if (iRecv < 0) return;

    const data = sheet.getRange(2, ANSWER_START_COL, lastRow - 1, width).getValues();
    for (let k = 0; k < data.length; k++) {
      const row = data[k];
      const hasRecv = row[iRecv] !== "" && row[iRecv] !== null && row[iRecv] !== undefined;
      const hasName = iName >= 0 && String(row[iName] || "").trim() !== "";
      if (!hasRecv && !hasName) continue;

      out.push({
        caseName: c.name,
        caseCode: c.code,
        name:     iName  >= 0 ? String(row[iName] || "").trim() : "",
        referrer: iRef   >= 0 ? String(row[iRef]  || "").trim() : "",
        agency:   iAgcy  >= 0 ? String(row[iAgcy] || "").trim() : "",
        appr:     iAppr  >= 0 ? row[iAppr] : "",
        shot:     iShot  >= 0 ? String(row[iShot] || "") : "",
        recvMs:   aspToMillis_(row[iRecv]),
        clickMs:  iClick >= 0 ? aspToMillis_(row[iClick]) : 0,
        sheetName: sheet.getName(),
        row: k + 2
      });
    }
  });
  return out;
}

// ---------------------------------------------------------------
// ASP側を案件ごとに集計する
//   - その案件に承認実績があるか（型A / 型B の分かれ目）
//   - その案件の最終クリック日時（リンク切れの判定に使う）
// ---------------------------------------------------------------
function agAspStatsByCase_(aspRows, caseNames) {
  const stats = {};
  caseNames.forEach(function (cn) {
    stats[cn] = { approved: 0, total: 0, lastClickMs: 0 };
  });
  aspRows.forEach(function (a) {
    caseNames.forEach(function (cn) {
      if (!aspNameMatches_(a.ad, cn)) return;
      const s = stats[cn];
      s.total++;
      if (a.status === "承認") s.approved++;
      if (a.ms > s.lastClickMs) s.lastClickMs = a.ms;
    });
  });
  return stats;
}

// ---------------------------------------------------------------
// Step2 + Step3: 棚卸しして型を判定する
// ---------------------------------------------------------------
function buildApprovalGapSheet() {
  const aspRows = collectAspLogRows_();
  if (!aspRows.length) {
    throw new Error("ASP獲得ログにデータがありません。先にCSVを取り込んでください。");
  }
  const own = collectOwnApplicationRows_();
  const nowMs = Date.now();

  const caseNames = [];
  own.forEach(function (o) { if (caseNames.indexOf(o.caseName) < 0) caseNames.push(o.caseName); });
  const stats = agAspStatsByCase_(aspRows, caseNames);

  // ASP行は1件につき自社1行にしか当てない。
  // 自社起点で回すと、同じASP行を近接する複数の申請が取り合う。取り合ったままだと
  // 「ASPに記録がある」申請が2件に見え、片方が誤って型Cへ落ちる。
  const used = {};

  // **案件ごとにASP行を先に索引する。** これをやらないと 自社550行 × ASP450行 の
  // 総当たりの中で毎回 aspNameMatches_（正規表現4本）を回すことになり、
  // 25万回のマッチで実行が分単位になる。代理店登録が56秒かかった件と同じ形
  // （2026-08-20 の教訓: GASで「件数 × 全走査」をやると簡単に分単位になる）。
  const aspByCase = {};
  caseNames.forEach(function (cn) {
    const list = [];
    aspRows.forEach(function (a, i) {
      if (aspNameMatches_(a.ad, cn)) list.push({ a: a, i: i });
    });
    aspByCase[cn] = list;
  });

  // 受信日時の新しい順に処理する（新しい申請ほど記憶が新しく、確認が取りやすい）
  own.sort(function (x, y) { return y.recvMs - x.recvMs; });

  const rows = [];
  const tally = { A: 0, B: 0, C: 0, D: 0, dead: 0, approved: 0, waiting: 0 };

  own.forEach(function (o) {
    // 既に承認が付いている申請は漏れではない
    if (getAdvertiserApprovalFlags(o.appr).approved) { tally.approved++; return; }

    const st = stats[o.caseName] || { approved: 0, total: 0, lastClickMs: 0 };
    const keyMs = o.clickMs || o.recvMs;

    // 同じ案件のASP行のうち、まだ使われていない中で最も近いものを取る
    const cand = aspByCase[o.caseName] || [];
    let best = null;
    for (let k = 0; k < cand.length; k++) {
      if (used[cand[k].i]) continue;
      const d = Math.abs(cand[k].a.ms - keyMs);
      if (d > AG_TOLERANCE_SEC * 1000) continue;
      if (best === null || d < best.d) best = { d: d, a: cand[k].a, i: cand[k].i };
    }

    let type = "", aspStatus = "", reward = "", link = "生存", sendOk = "";
    let ageMs = keyMs;

    if (best) {
      used[best.i] = true;
      aspStatus = best.a.status;
      reward = best.a.reward;
      if (aspStatus === "承認") {
        // ASPは承認済みだが自社の承認欄が空。これは AspReconcile.gs 側の仕事（要修正）。
        // 承認漏れではないので、ここでは扱わずに数だけ持つ。
        tally.approved++;
        return;
      }
      if (aspStatus === "否認") {
        type = "D";
        sendOk = "出さない"; // 理由確認のみ。既定では依頼に載せない
      } else {
        // 承認待ち。しきい値に達していなければ、まだ待ちの範囲なので出さない
        const days = agDaysBetween_(ageMs, nowMs);
        if (days < agThresholdDays_(o.caseName)) { tally.waiting++; return; }
        type = st.approved > 0 ? "B" : "A";
        sendOk = (type === "A") ? "出す" : ""; // 型Aは営業確認が要らないのですぐ出せる
      }
    } else {
      // ASPに記録が無い。トラッキング漏れ **または** リンク切れ。
      // その案件のASP最終クリックより後の申請は、リンクが死んでいた可能性が高い。
      // 死んだリンクを踏んだ申請はASPにクリックが計上されず、広告主にも記録が無い。
      // これを広告主へ出すと、自社の管理不備を自分から見せることになるので既定で外す。
      aspStatus = "記録なし";
      if (st.lastClickMs && keyMs > st.lastClickMs) {
        link = "リンク切れ疑い";
        type = "C";
        sendOk = "出さない";
        tally.dead++;
      } else {
        type = "C";
      }
    }

    if (type === "A") tally.A++;
    else if (type === "B") tally.B++;
    else if (type === "C" && link === "生存") tally.C++;
    else if (type === "D") tally.D++;

    rows.push({
      key: agKeyOf_(o.caseName, o.name, o.recvMs),
      values: [
        type, o.caseName, o.name, o.referrer, o.agency,
        o.recvMs ? formatJST(new Date(o.recvMs)) : "",
        o.clickMs ? formatJST(new Date(o.clickMs)) : "",
        aspStatus, agDaysBetween_(ageMs, nowMs), reward, link, o.shot,
        "未確認", "", sendOk, "", "", "", "", "", o.sheetName, o.row
      ]
    });
  });

  // 型A → 型B → 型C → 型D、同じ型の中は経過日数の長い順
  const order = { A: 0, B: 1, C: 2, D: 3 };
  rows.sort(function (x, y) {
    const d = order[x.values[0]] - order[y.values[0]];
    if (d !== 0) return d;
    return (y.values[AGC_DAYS - 1] || 0) - (x.values[AGC_DAYS - 1] || 0);
  });

  // 既存の確認結果を引き継ぐ。**ここを作り直しにすると、営業担当に同じ確認を毎月させることになる。**
  const sh = agManageSheet_();
  const carry = {};
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const old = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();
    old.forEach(function (r) {
      const k = agKeyOf_(r[AGC_CASE - 1], r[AGC_NAME - 1], aspToMillis_(r[AGC_RECV - 1]));
      carry[k] = {
        salesChk: r[AGC_SALESCHK - 1], salesCmt: r[AGC_SALESCMT - 1],
        sendOk:   r[AGC_SENDOK - 1],
        reqId:    r[AGC_REQID - 1],   reqDate: r[AGC_REQDATE - 1],
        advAns:   r[AGC_ADVANS - 1],  ansDate: r[AGC_ANSDATE - 1],
        result:   r[AGC_RESULT - 1]
      };
    });
  }
  let carried = 0;
  rows.forEach(function (r) {
    const c = carry[r.key];
    if (!c) return;
    carried++;
    if (c.salesChk) r.values[AGC_SALESCHK - 1] = c.salesChk;
    if (c.salesCmt) r.values[AGC_SALESCMT - 1] = c.salesCmt;
    if (c.sendOk)   r.values[AGC_SENDOK  - 1] = c.sendOk;
    if (c.reqId)    r.values[AGC_REQID   - 1] = c.reqId;
    if (c.reqDate)  r.values[AGC_REQDATE - 1] = c.reqDate;
    if (c.advAns)   r.values[AGC_ADVANS  - 1] = c.advAns;
    if (c.ansDate)  r.values[AGC_ANSDATE - 1] = c.ansDate;
    if (c.result)   r.values[AGC_RESULT  - 1] = c.result;
  });

  sh.clear();
  const hr = sh.getRange(1, 1, 1, AG_HEADERS.length);
  hr.setValues([AG_HEADERS]);
  hr.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff").setWrap(true);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  if (rows.length) {
    sh.getRange(2, 1, rows.length, AG_HEADERS.length)
      .setValues(rows.map(function (r) { return r.values; }));

    sh.getRange(2, AGC_SALESCHK, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(AG_SALES_CHOICES, true).build());
    sh.getRange(2, AGC_SENDOK, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["出す", "出さない"], true).build());
    sh.getRange(2, AGC_ADVANS, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(AG_ADV_CHOICES, true).build());

    // 型ごとに色を付ける。型A＝営業確認が要らない（＝すぐ出せる）ことが一目で分かるように。
    const tint = { A: "#dcfce7", B: "#fef9c3", C: "#e0f2fe", D: "#f1f5f9" };
    const bg = rows.map(function (r) {
      const c = r.values[AGC_LINK - 1] === "リンク切れ疑い" ? "#f5f5f4" : (tint[r.values[0]] || "#ffffff");
      return AG_HEADERS.map(function () { return c; });
    });
    sh.getRange(2, 1, rows.length, AG_HEADERS.length).setBackgrounds(bg);
  }

  sh.setColumnWidth(AGC_TYPE, 40);
  sh.setColumnWidth(AGC_CASE, 220);
  sh.setColumnWidth(AGC_NAME, 120);
  sh.setColumnWidth(AGC_REF, 120);
  sh.setColumnWidth(AGC_RECV, 150);
  sh.setColumnWidth(AGC_CLICK, 150);
  sh.setColumnWidth(AGC_SHOT, 240);
  sh.setColumnWidth(AGC_SALESCMT, 240);

  return {
    A: tally.A, B: tally.B, C: tally.C, D: tally.D,
    dead: tally.dead, approved: tally.approved, waiting: tally.waiting,
    total: rows.length, carried: carried, ownTotal: own.length
  };
}

// ---------------------------------------------------------------
// Step4: 営業担当へ確認を依頼する（SS2へ担当別タブを書き出す）
// ---------------------------------------------------------------
const AG_SALES_HEADERS = [
  "案件名", "顧客名", "受信日時", "スクショURL", "確認結果", "コメント", "照合キー"
];

function pushSalesApprovalChecks() {
  const sh = agManageSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { reps: 0, rows: 0, unassigned: 0 };
  const data = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();

  // 営業確認が要るのは 型B と、リンクが生きている型C だけ。型A・型Dは回さない。
  const byRep = {};
  let unassigned = 0;
  data.forEach(function (r) {
    const type = String(r[AGC_TYPE - 1] || "");
    if (type !== "B" && type !== "C") return;
    if (String(r[AGC_LINK - 1] || "") === "リンク切れ疑い") return;
    const chk = String(r[AGC_SALESCHK - 1] || "");
    if (chk && chk !== "未確認") return; // 済みは再依頼しない
    const rep = String(r[AGC_REF - 1] || "").trim();
    if (!rep) { unassigned++; return; }
    if (!byRep[rep]) byRep[rep] = [];
    byRep[rep].push([
      r[AGC_CASE - 1], r[AGC_NAME - 1], r[AGC_RECV - 1], r[AGC_SHOT - 1],
      "未確認", "", agKeyOf_(r[AGC_CASE - 1], r[AGC_NAME - 1], aspToMillis_(r[AGC_RECV - 1]))
    ]);
  });

  let outSS;
  try { outSS = SpreadsheetApp.openById(REP_STATUS_SS2_ID); }
  catch (e) { throw new Error("SS2（担当別ステータス表）を開けません: " + e); }

  let reps = 0, total = 0;
  Object.keys(byRep).forEach(function (rep) {
    const name = AG_SALES_PREFIX + rep;
    let tab = outSS.getSheetByName(name);
    if (!tab) tab = outSS.insertSheet(name);
    tab.clear();
    const hr = tab.getRange(1, 1, 1, AG_SALES_HEADERS.length);
    hr.setValues([AG_SALES_HEADERS]);
    hr.setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    tab.setFrozenRows(1);

    const list = byRep[rep];
    tab.getRange(2, 1, list.length, AG_SALES_HEADERS.length).setValues(list);
    tab.getRange(2, 5, list.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(AG_SALES_CHOICES, true).build());
    // 照合キーは機械が読み戻すためのもの。触られると反映できなくなるので隠す。
    tab.hideColumns(7);
    tab.setColumnWidth(1, 220);
    tab.setColumnWidth(2, 120);
    tab.setColumnWidth(3, 150);
    tab.setColumnWidth(4, 260);
    tab.setColumnWidth(6, 260);
    // 何を見るかをタブの中に書いておく（別の場所の手順書を読ませない）
    tab.getRange(list.length + 3, 1).setValue(
      "確認していただきたいのは「サンクスメールのスクショが証拠として成立しているか」の1点です。" +
      "(1)その案件の申込完了メールか (2)申込日時が写っているか (3)顧客のメールアドレスが写っているか " +
      "(4)顧客名が申請と一致するか (5)仮登録で止まっていないか。" +
      "問題なければ「OK」、スクショを取り直す必要があれば「要再取得」、" +
      "申込が成立していなければ「取下げ」を選んでください。");
    reps++; total += list.length;
  });

  return { reps: reps, rows: total, unassigned: unassigned };
}

// 担当別タブの記入を 承認漏れ管理 へ読み戻す
function pullSalesApprovalChecks() {
  let outSS;
  try { outSS = SpreadsheetApp.openById(REP_STATUS_SS2_ID); }
  catch (e) { throw new Error("SS2を開けません: " + e); }

  const answers = {};
  outSS.getSheets().forEach(function (tab) {
    const nm = tab.getName();
    if (nm.indexOf(AG_SALES_PREFIX) !== 0) return;
    const last = tab.getLastRow();
    if (last < 2) return;
    const v = tab.getRange(2, 1, last - 1, AG_SALES_HEADERS.length).getValues();
    v.forEach(function (r) {
      const key = String(r[6] || "");
      const chk = String(r[4] || "").trim();
      if (!key || !chk || chk === "未確認") return;
      answers[key] = { chk: chk, cmt: String(r[5] || "") };
    });
  });

  const sh = agManageSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { updated: 0 };
  const data = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();
  let updated = 0;
  data.forEach(function (r, i) {
    const key = agKeyOf_(r[AGC_CASE - 1], r[AGC_NAME - 1], aspToMillis_(r[AGC_RECV - 1]));
    const a = answers[key];
    if (!a) return;
    if (String(r[AGC_SALESCHK - 1] || "") === a.chk) return;
    const row = i + 2;
    sh.getRange(row, AGC_SALESCHK).setValue(a.chk);
    if (a.cmt) sh.getRange(row, AGC_SALESCMT).setValue(a.cmt);
    // OK なら広告主へ出せる。取下げ・要再取得は出さない。
    sh.getRange(row, AGC_SENDOK).setValue(a.chk === "OK" ? "出す" : "出さない");
    updated++;
  });
  SpreadsheetApp.flush();
  return { updated: updated };
}

// ---------------------------------------------------------------
// Step5: 広告主へまとめて渡す確認依頼シートを作る
//        （広告主成果管理SSに追記する。広告主はここへ結果を記入する）
// ---------------------------------------------------------------
const AG_REQ_HEADERS = [
  "依頼ID", "依頼日", "広告名", "申込日時", "クリック日時", "お名前", "紹介者名",
  "スクショ（申込完了メール）", "当方の記録", "ASPステータス",
  "確認結果【広告主記入】", "理由・コメント【広告主記入】", "回答日【広告主記入】"
];

function agRequestSheet_() {
  const ss = SpreadsheetApp.openById(ADVERTISER_SS_ID);
  let sh = ss.getSheetByName(AG_REQUEST_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AG_REQUEST_SHEET, 0);
    sh.getRange(1, 1).setValue(
      "承認状況の確認依頼です。右の3列（確認結果・理由・回答日）へご記入ください。" +
      "当方では申込完了メール（サンクスメール）の控えを確認しています。");
    sh.getRange(1, 1).setFontWeight("bold");
    const hr = sh.getRange(2, 1, 1, AG_REQ_HEADERS.length);
    hr.setValues([AG_REQ_HEADERS]);
    hr.setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff").setWrap(true);
    sh.setFrozenRows(2);
    // 記入してもらう3列を色で区別する
    sh.getRange(2, 11, 1, 3).setBackground("#b45309");
    sh.setColumnWidth(3, 220);
    sh.setColumnWidth(4, 150);
    sh.setColumnWidth(5, 150);
    sh.setColumnWidth(8, 260);
    sh.setColumnWidth(9, 220);
    sh.setColumnWidth(12, 260);
  }
  return sh;
}

// **追記する。既存行は消さない。** 広告主が記入した回答を消してしまうため。
function pushAdvertiserRequests() {
  const sh = agManageSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { added: 0, skipped: 0 };
  const data = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();

  const req = agRequestSheet_();
  const reqLast = req.getLastRow();

  const today = formatJST(new Date()).slice(0, 10);
  const stamp = today.replace(/[^0-9]/g, "");
  let seq = reqLast >= 3 ? reqLast - 2 : 0;

  const toAdd = [];
  const idByRow = {};
  data.forEach(function (r, i) {
    if (String(r[AGC_SENDOK - 1] || "") !== "出す") return;
    // **既に依頼IDが付いている行は二度と載せない。** 同じ案件が2行並ぶと広告主側で
    // 「どちらに答えたか」が分からなくなる。出し直したいときは依頼IDを手で消す。
    if (String(r[AGC_REQID - 1] || "")) return;
    const type = String(r[AGC_TYPE - 1] || "");
    seq++;
    const id = "REQ-" + stamp + "-" + ("000" + seq).slice(-4);
    const note = (type === "A")
      ? "同案件でこの期間の成果が一件も承認されていません"
      : "当方で申込完了メールを確認済みです";
    toAdd.push([
      id, today, r[AGC_CASE - 1], r[AGC_RECV - 1], r[AGC_CLICK - 1],
      r[AGC_NAME - 1], r[AGC_REF - 1], r[AGC_SHOT - 1], note, r[AGC_ASPST - 1],
      "", "", ""
    ]);
    idByRow[i + 2] = id;
  });

  if (!toAdd.length) return { added: 0, skipped: 0 };

  // 広告名でまとめて並べる（案件ごとに見てもらうため）
  toAdd.sort(function (x, y) { return String(x[2]).localeCompare(String(y[2])); });

  const start = Math.max(reqLast + 1, 3);
  req.getRange(start, 1, toAdd.length, AG_REQ_HEADERS.length).setValues(toAdd);
  req.getRange(start, 11, toAdd.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(AG_ADV_CHOICES, true).build());
  req.getRange(start, 11, toAdd.length, 3).setBackground("#fff7ed");

  // 依頼IDと依頼日を 承認漏れ管理 へ書き戻す
  Object.keys(idByRow).forEach(function (rowNum) {
    sh.getRange(Number(rowNum), AGC_REQID).setValue(idByRow[rowNum]);
    sh.getRange(Number(rowNum), AGC_REQDATE).setValue(today);
  });
  SpreadsheetApp.flush();

  const byCase = {};
  toAdd.forEach(function (r) { byCase[r[2]] = (byCase[r[2]] || 0) + 1; });
  return { added: toAdd.length, cases: byCase, url: SpreadsheetApp.openById(ADVERTISER_SS_ID).getUrl() };
}

// 広告主の記入を 承認漏れ管理 へ読み戻す
function pullAdvertiserAnswers() {
  const req = agRequestSheet_();
  const reqLast = req.getLastRow();
  if (reqLast < 3) return { updated: 0 };
  const rows = req.getRange(3, 1, reqLast - 2, AG_REQ_HEADERS.length).getValues();

  const ans = {};
  rows.forEach(function (r) {
    const id = String(r[0] || "");
    const result = String(r[10] || "").trim();
    if (!id || !result || result === "確認中") return;
    ans[id] = { result: result, note: String(r[11] || ""), date: r[12] };
  });

  const sh = agManageSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { updated: 0 };
  const data = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();
  let updated = 0;
  data.forEach(function (r, i) {
    const id = String(r[AGC_REQID - 1] || "");
    const a = ans[id];
    if (!a) return;
    if (String(r[AGC_RESULT - 1] || "") === a.result) return;
    const row = i + 2;
    sh.getRange(row, AGC_ADVANS).setValue(a.note || a.result);
    sh.getRange(row, AGC_ANSDATE).setValue(a.date || formatJST(new Date()).slice(0, 10));
    sh.getRange(row, AGC_RESULT).setValue(a.result);
    updated++;
  });
  SpreadsheetApp.flush();
  return { updated: updated };
}

// ---------------------------------------------------------------
// Step6: 承認された分だけ自社シートへ ⭕ を書く
//        **一方向だけ。承認を消す操作はしない**（AspReconcile と同じ方針・同じ理由）
// ---------------------------------------------------------------
function applyApprovalGapApproved(dryRun) {
  const ss = getOrCreateSpreadsheet();
  const sh = agManageSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { applied: 0, skipped: 0, reasons: [] };
  const data = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();

  let applied = 0, skipped = 0;
  const reasons = [];

  data.forEach(function (r) {
    if (String(r[AGC_RESULT - 1] || "") !== "承認") return;

    const sheetName = String(r[AGC_SHEET - 1] || "");
    const row = Number(r[AGC_ROW - 1] || 0);
    const label = String(r[AGC_CASE - 1]) + " / " + String(r[AGC_NAME - 1]);
    if (!sheetName || !row) { skipped++; reasons.push(label + ": 対象行が特定できない"); return; }

    const target = ss.getSheetByName(sheetName);
    if (!target) { skipped++; reasons.push(label + ": シートが無い（" + sheetName + "）"); return; }

    const lastCol = target.getLastColumn();
    const width = lastCol - ANSWER_START_COL + 1;
    const headers = target.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const iAppr = headers.indexOf("承認");
    const iName = headers.indexOf("お名前");
    const iRecv = headers.indexOf("受信日時");
    if (iAppr < 0) { skipped++; reasons.push(label + ": 承認列が無い"); return; }

    // **行番号だけを信じて書かない。** 行は案件シート側の編集で動く。
    // 依頼したときの顧客名・受信日時と一致するかを見てから書く。
    const cur = target.getRange(row, ANSWER_START_COL, 1, width).getValues()[0];
    if (iName >= 0 && String(cur[iName] || "").trim() !== String(r[AGC_NAME - 1] || "").trim()) {
      skipped++; reasons.push(label + ": 顧客名が一致しない（行がずれている）"); return;
    }
    if (iRecv >= 0) {
      const a = aspToMillis_(cur[iRecv]), b = aspToMillis_(r[AGC_RECV - 1]);
      if (a && b && Math.abs(a - b) > 60000) {
        skipped++; reasons.push(label + ": 受信日時が一致しない（行がずれている）"); return;
      }
    }
    if (getAdvertiserApprovalFlags(cur[iAppr]).approved) {
      skipped++; reasons.push(label + ": 既に承認になっている"); return;
    }

    if (!dryRun) target.getRange(row, ANSWER_START_COL + iAppr).setValue("⭕");
    applied++;
  });
  if (!dryRun) SpreadsheetApp.flush();
  return { applied: applied, skipped: skipped, reasons: reasons };
}

// ---------------------------------------------------------------
// Step7: 放置の検出（日次レポートから呼ぶ。fail-closed で本体を止めない）
// ---------------------------------------------------------------
function buildApprovalGapSection_() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sh = ss.getSheetByName(AG_MANAGE_SHEET);
    if (!sh) return "";
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return "";
    const data = sh.getRange(2, 1, lastRow - 1, AG_HEADERS.length).getValues();

    const nowMs = Date.now();
    const overdueByRep = {};
    let noAnswer = 0;

    data.forEach(function (r) {
      // 営業確認の期限超過（3営業日 ≒ 5日で見る）
      const type = String(r[AGC_TYPE - 1] || "");
      const chk  = String(r[AGC_SALESCHK - 1] || "");
      if ((type === "B" || type === "C") &&
          String(r[AGC_LINK - 1] || "") !== "リンク切れ疑い" &&
          (!chk || chk === "未確認")) {
        const rep = String(r[AGC_REF - 1] || "").trim() || "（担当なし）";
        overdueByRep[rep] = (overdueByRep[rep] || 0) + 1;
      }
      // 依頼したまま14日回答が無い
      const reqDate = agDateToMillis_(r[AGC_REQDATE - 1]);
      if (reqDate && !String(r[AGC_RESULT - 1] || "") &&
          agDaysBetween_(reqDate, nowMs) >= 14) noAnswer++;
    });

    const reps = Object.keys(overdueByRep);
    if (!reps.length && !noAnswer) return "";

    let s = "\n■ 承認漏れの確認";
    if (reps.length) {
      s += "\n・営業確認の未対応: ";
      s += reps.sort(function (a, b) { return overdueByRep[b] - overdueByRep[a]; })
               .map(function (r) { return r + " " + overdueByRep[r] + "件"; }).join(" / ");
    }
    if (noAnswer) s += "\n・広告主へ依頼したまま14日以上回答なし: " + noAnswer + "件";
    return s;
  } catch (e) {
    Logger.log("承認漏れ節の生成に失敗（日次レポートは続行）: " + e);
    return "";
  }
}

// ---------------------------------------------------------------
// メニュー
// ---------------------------------------------------------------
function buildApprovalGapFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = buildApprovalGapSheet();
    ui.alert(
      "承認漏れを棚卸ししました。\n\n" +
      "自社の申請 " + r.ownTotal + "件を見て、承認が取れていないものを型で分けました。\n\n" +
      "型A（案件まるごと0承認・営業確認は不要）: " + r.A + " 件\n" +
      "型B（同案件に承認実績あり・個別の取り残し）: " + r.B + " 件\n" +
      "型C（ASPに記録なし・トラッキング漏れ疑い）: " + r.C + " 件\n" +
      "型D（否認）: " + r.D + " 件\n" +
      "リンク切れ疑い（依頼の対象外）: " + r.dead + " 件\n\n" +
      "承認済み: " + r.approved + " 件 / まだ待ちの範囲（しきい値内）: " + r.waiting + " 件\n" +
      "前回の確認結果を引き継いだ行: " + r.carried + " 件\n\n" +
      "型Aは営業確認が要らないので「広告主への確認依頼を作る」をすぐ実行できます。\n" +
      "型B・型Cは先に「営業担当へ確認を依頼」を実行してください。");
  } catch (e) {
    ui.alert("棚卸しできませんでした。\n\n" + e);
  }
}

function pushSalesApprovalChecksFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = pushSalesApprovalChecks();
    if (!r.rows) { ui.alert("営業確認が必要な行はありませんでした。"); return; }
    ui.alert("担当別ステータス表（SS2）へ「" + AG_SALES_PREFIX + "＜担当名＞」タブを書き出しました。\n\n" +
             "担当: " + r.reps + "名 / 依頼: " + r.rows + "件\n" +
             (r.unassigned ? "紹介者が空で割り当てられなかった行: " + r.unassigned + "件\n" : "") +
             "\n記入が終わったら「営業の確認結果を取り込む」を実行してください。");
  } catch (e) {
    ui.alert("書き出せませんでした。\n\n" + e);
  }
}

function pullSalesApprovalChecksFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = pullSalesApprovalChecks();
    ui.alert("営業の確認結果を取り込みました。\n\n更新: " + r.updated + " 件\n\n" +
             "OK になったものは「依頼可否」が『出す』になります。");
  } catch (e) {
    ui.alert("取り込めませんでした。\n\n" + e);
  }
}

function pushAdvertiserRequestsFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert(
    "「依頼可否」が『出す』の行を、広告主成果管理SSの「" + AG_REQUEST_SHEET + "」へ追記します。\n" +
    "既存の行は消しません（広告主の記入を保持するため）。よろしいですか。",
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  try {
    const r = pushAdvertiserRequests();
    if (!r.added) { ui.alert("追記する行はありませんでした。\n\n" +
      "型Aは棚卸し直後から『出す』になります。型B・型Cは営業確認がOKになると『出す』に変わります。"); return; }
    const lines = Object.keys(r.cases).map(function (k) { return "  " + k + ": " + r.cases[k] + "件"; });
    ui.alert("確認依頼を追記しました。\n\n合計 " + r.added + " 件\n" + lines.join("\n") +
             "\n\n広告主成果管理SSの「" + AG_REQUEST_SHEET + "」タブを共有してください。");
  } catch (e) {
    ui.alert("追記できませんでした。\n\n" + e);
  }
}

function pullAdvertiserAnswersFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = pullAdvertiserAnswers();
    ui.alert("広告主の回答を取り込みました。\n\n更新: " + r.updated + " 件\n\n" +
             "『承認』になったものは「承認された分を自社シートへ反映」で ⭕ を書けます。");
  } catch (e) {
    ui.alert("取り込めませんでした。\n\n" + e);
  }
}

function applyApprovalGapFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const dry = applyApprovalGapApproved(true);
  const head = "下見の結果です。\n\n書き込み予定: " + dry.applied + " 件\nスキップ: " + dry.skipped + " 件";
  const detail = dry.reasons.length ? "\n\n" + dry.reasons.slice(0, 15).join("\n") : "";
  if (!dry.applied) { ui.alert(head + detail); return; }
  const ok = ui.alert(head + detail + "\n\n自社シートの承認欄へ ⭕ を書き込みます。" +
                      "承認を消す操作は行いません。よろしいですか。", ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const r = applyApprovalGapApproved(false);
  ui.alert("反映しました。\n\n書き込み: " + r.applied + " 件 / スキップ: " + r.skipped + " 件\n\n" +
           "広告主シートへ提出する前に、対象月を再生成してください。");
}
