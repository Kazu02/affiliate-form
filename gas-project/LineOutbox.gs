// ===== 公式LINE の送信と「届かなかった分」の後送り（2026-10-01 追加） =====
//
// 公式LINE（市場作り管理BOT）はライトプラン＝月5,000通で、アフィリンクとアポ報告（フォーム顧客管理）の
// 2つの GAS が同じ枠を使う。グループへの push は「グループの人数 × 1通」で数え、
// 1回の push に吹き出しを5つまで入れても1通のまま。
// 枠が尽きると push は 429 を返す。以前の送信関数は応答を見ずに捨てていた（2026-09-29 に発生）。
//
//   1. 送れなかった本文はシート「LINE送信待ち」に残す
//   2. 1時間おきの flushLineOutbox() が、待ちがあるときだけ残り枠を確かめ、
//      受付時刻を添えて5つずつまとめて送る。月初0時に枠が戻ると、その直後の回で届く
//   3. 今月の送信数と、送信待ち・送信不可の件数を日次レポートに添える（同じ吹き出しに足すので通数は増えない）
//
// 「LINE送信待ち」は行番号で印を付けるので、行を消したり並べ替えたりしない。
// このファイルはアフィリンク（LineOutbox.gs）とフォーム顧客管理（line-outbox.js）で同じ本文。
// 違うのは「プロジェクトごとの設定」の2か所だけ。

// ---- プロジェクトごとの設定 ----
var LO_GROUP_PROPERTY = 'LINE_GROUP_ID';

// getOrCreateSpreadsheet() は開けないと新しいスプレッドシートを作るので使わない。
function loSpreadsheet_() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
}
// ---- ここまで ----

var LO_API              = 'https://api.line.me';
var LO_SHEET_NAME       = 'LINE送信待ち';
var LO_HEADERS          = ['受付日時', '状態', '本文', '失敗の理由', '送信日時'];
var LO_WAITING          = '待ち';
var LO_SENT             = '送信済み';
var LO_REJECTED         = '送信不可'; // LINE が本文を受け付けない（400）。送り直しても通らないので人が見る
var LO_FLUSH_HANDLER    = 'flushLineOutbox';
var LO_LEASE_KEY        = 'LINE_OUTBOX_FLUSH_UNTIL';
var LO_BUBBLES_PER_PUSH = 5;    // push 1回に入れられる吹き出しの上限
var LO_MAX_TEXT         = 4990; // 吹き出し1つは5,000文字まで
var LO_WARN_RATIO       = 0.8;
var LO_REJECTED_DAYS    = 7;    // 日次レポートで「送信不可」を数える期間

function loConfig_() {
  var props   = PropertiesService.getScriptProperties();
  var token   = props.getProperty('LINE_CHANNEL_TOKEN');
  var groupId = props.getProperty(LO_GROUP_PROPERTY);
  return token && groupId ? { token: token, groupId: groupId } : null;
}

// 絵文字（サロゲートペア）の途中では切らない。
function loClip_(text) {
  var t = String(text);
  if (t.length <= LO_MAX_TEXT) return t;
  var cut = LO_MAX_TEXT;
  var c = t.charCodeAt(cut - 1);
  if (c >= 0xD800 && c <= 0xDBFF) cut--;
  return t.substring(0, cut) + '...';
}

function loPush_(cfg, texts, retryKey) {
  var res = UrlFetchApp.fetch(LO_API + '/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + cfg.token,
      'X-Line-Retry-Key': retryKey
    },
    payload: JSON.stringify({
      to: cfg.groupId,
      messages: texts.map(function(t) { return { type: 'text', text: loClip_(t) }; })
    }),
    muteHttpExceptions: true
  });
  return { code: res.getResponseCode(), body: res.getContentText().substring(0, 200) };
}

// 応答が切れたら同じ再送キーで1回だけ送り直す。LINE 側で受付済みなら 409 が返り、二重には届かない。
// 2回目も切れたら例外を投げる。
function loSend_(cfg, texts) {
  var key = Utilities.getUuid();
  try {
    return loPush_(cfg, texts, key);
  } catch (e) {
    return loPush_(cfg, texts, key);
  }
}

function loDelivered_(r) { return r.code === 200 || r.code === 409; }

// 宛先（グループID）の誤りによる 400。本文のせいではないので「送信不可」にしない。
function loDestinationError_(r) { return r.code === 400 && r.body.indexOf("'to'") >= 0; }

function loGetJson_(cfg, path) {
  var res = UrlFetchApp.fetch(LO_API + path, {
    headers: { 'Authorization': 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error(path + ': HTTP ' + res.getResponseCode());
  return JSON.parse(res.getContentText());
}

// ---- グループへ送る ----
// 5つまでは1回の push（通数は1回分）。届かなかった本文は「LINE送信待ち」へ残し、false を返す。
// トークンかグループIDが無いときは従来どおり何もしない。
function sendLineGroup_(texts) {
  var cfg = loConfig_();
  if (!cfg) return false;
  var list = texts.filter(function(t) { return t !== null && t !== undefined && String(t) !== ''; });
  var ok = true;
  for (var i = 0; i < list.length; i += LO_BUBBLES_PER_PUSH) {
    var chunk = list.slice(i, i + LO_BUBBLES_PER_PUSH);
    var reason;
    try {
      var r = loSend_(cfg, chunk);
      if (loDelivered_(r)) continue;
      reason = 'HTTP ' + r.code + ' ' + r.body;
    } catch (e) {
      // 2回とも応答が切れた。LINE 側では届いている可能性もあるが、落とすよりは二重を選ぶ
      reason = '応答なし: ' + e;
    }
    ok = false;
    lineOutboxAdd_(chunk, reason);
  }
  return ok;
}

// ---- 送信待ちシート ----
function lineOutboxSheet_(ss) {
  var sheet = ss.getSheetByName(LO_SHEET_NAME);
  if (sheet) return sheet;
  try {
    sheet = ss.insertSheet(LO_SHEET_NAME);
  } catch (e) {
    sheet = ss.getSheetByName(LO_SHEET_NAME); // 同時に別の実行が作った
    if (sheet) return sheet;
    throw e;
  }
  sheet.getRange(1, 1, 1, LO_HEADERS.length).setValues([LO_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

// = + - @ で始まる文字列は appendRow が数式にするので、先頭に ' を付けて文字のまま残す。
function loCellText_(t) {
  var s = String(t);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function loFromCell_(v) {
  var s = String(v);
  return /^'[=+\-@]/.test(s) ? s.substring(1) : s;
}

// 呼び出し元（フォーム受付など）を止めないよう、ここでは例外を外へ出さない。
function lineOutboxAdd_(texts, reason) {
  try {
    var sheet = lineOutboxSheet_(loSpreadsheet_());
    var now = new Date();
    texts.forEach(function(t) {
      sheet.appendRow([now, LO_WAITING, loCellText_(t), loCellText_(String(reason).substring(0, 300)), '']);
    });
  } catch (e) {
    console.error('LINE送信待ちへ残せませんでした: ' + e + ' / ' + reason);
  }
}

// 「待ち」の行を返す。送信済みの行は増える一方なので、まず状態の列だけ読み、最初の待ちから下を読む。
function lineOutboxWaiting_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var last = sheet.getLastRow();
  var states = sheet.getRange(2, 2, last - 1, 1).getValues();
  var first = -1;
  for (var i = 0; i < states.length; i++) {
    if (states[i][0] === LO_WAITING) { first = i; break; }
  }
  if (first < 0) return [];
  var waiting = [];
  sheet.getRange(first + 2, 1, last - first - 1, LO_HEADERS.length).getValues().forEach(function(r, j) {
    if (r[1] === LO_WAITING) {
      waiting.push({ row: first + 2 + j, at: r[0], text: loFromCell_(r[2]), reason: loFromCell_(r[3]) });
    }
  });
  return waiting;
}

// 見出し行の直後に受付時刻を差し込む（新しい報告と取り違えないため）。
function loWithTime_(text, at) {
  if (!(at instanceof Date)) return text;
  var t = Utilities.formatDate(at, 'Asia/Tokyo', 'M/d HH:mm');
  var i = text.indexOf('\n');
  return i < 0 ? text + '\n（' + t + ' 受付分）' : text.substring(0, i) + '（' + t + ' 受付分）' + text.substring(i);
}

// 先頭に事情の説明を1つ置き、5つずつ push に分ける。items[k].row はその吹き出しの待ち行（説明は 0）。
// remaining は今回送らずに残す件数（残り枠が足りないとき）。
function loBuildPushes_(waiting, remaining) {
  var byQuota = waiting.every(function(w) { return w.reason.indexOf('HTTP 429') === 0; });
  var intro = '【未送信分の再送】\n' +
    (byQuota ? '公式LINEの月間送信数が上限に達していたため届かなかった通知 '
             : '公式LINEの送信に失敗していた通知 ') +
    waiting.length + '件を、受付順にまとめて送ります。\n' +
    (remaining > 0 ? '（残りの ' + remaining + '件は、送信枠が空きしだい送ります）\n' : '') +
    '新しい報告ではありません（受付時刻をご確認ください）。';
  var items = [{ text: intro, row: 0 }].concat(waiting.map(function(w) {
    return { text: loWithTime_(w.text, w.at), row: w.row };
  }));
  var pushes = [];
  for (var i = 0; i < items.length; i += LO_BUBBLES_PER_PUSH) {
    pushes.push(items.slice(i, i + LO_BUBBLES_PER_PUSH));
  }
  return pushes;
}

function loMark_(sheet, rows, state, note) {
  var now = new Date();
  rows.forEach(function(row) {
    sheet.getRange(row, 2).setValue(state);
    if (note) sheet.getRange(row, 4).setValue(loCellText_(note));
    sheet.getRange(row, 5).setValue(now);
  });
  SpreadsheetApp.flush();
}

// ---- 今月の送信数 ----
function lineQuota_(cfg) {
  var q = loGetJson_(cfg, '/v2/bot/message/quota');
  if (q.type !== 'limited') return { limited: false };
  var used = loGetJson_(cfg, '/v2/bot/message/quota/consumption').totalUsage;
  return { limited: true, limit: Number(q.value), used: Number(used) };
}

// あと何回 push できるか。分からないときは送ってみる（失敗すれば待ちに残るだけ）。
function linePushBudget_(cfg) {
  try {
    var q = lineQuota_(cfg);
    if (!q.limited) return Infinity;
    var members = Number(loGetJson_(cfg, '/v2/bot/group/' + cfg.groupId + '/members/count').count);
    if (!(members > 0)) return Infinity;
    return Math.floor((q.limit - q.used) / members);
  } catch (e) {
    Logger.log('残り枠を確かめられませんでした: ' + e);
    return Infinity;
  }
}

// ---- 1時間おきのトリガーで実行 ----
// 待ちが無ければシートの状態の列を読むだけで終わる（LINE の API は呼ばない）。
// 顧客LINE連携などが短い待ち時間でスクリプトロックを取るので、ここではロックを使わず
// ScriptProperties の期限で二重起動だけ防ぐ。
function flushLineOutbox() {
  var sheet = loSpreadsheet_().getSheetByName(LO_SHEET_NAME);
  var waiting = lineOutboxWaiting_(sheet);
  if (!waiting.length) return;
  var cfg = loConfig_();
  if (!cfg) return;

  var props = PropertiesService.getScriptProperties();
  if (Number(props.getProperty(LO_LEASE_KEY) || 0) > Date.now()) return;
  props.setProperty(LO_LEASE_KEY, String(Date.now() + 5 * 60 * 1000));
  try {
    var left = linePushBudget_(cfg);
    if (left < 1) {
      Logger.log('残り枠が足りないので待つ（待ち ' + waiting.length + '件）');
      return;
    }
    // 説明の吹き出しを含めて left 回の push に入る件数だけ送る
    var sendable = left === Infinity ? waiting : waiting.slice(0, left * LO_BUBBLES_PER_PUSH - 1);
    var pushes = loBuildPushes_(sendable, waiting.length - sendable.length);
    for (var p = 0; p < pushes.length; p++) {
      if (left < 1) break;
      var items = pushes[p];
      var rows = items.filter(function(it) { return it.row > 0; }).map(function(it) { return it.row; });
      var r = loSend_(cfg, items.map(function(it) { return it.text; }));
      left--;
      Logger.log('再送 push ' + (p + 1) + '/' + pushes.length + ': HTTP ' + r.code);
      if (loDelivered_(r)) { loMark_(sheet, rows, LO_SENT, ''); continue; }
      if (r.code !== 400 || loDestinationError_(r)) {
        if (r.code === 400) console.error('LINE の宛先の設定を確認してください: ' + r.body);
        break; // 枠切れ・一時的な障害・設定の誤り。残りは次の回へ
      }
      left = loIsolateRejected_(cfg, sheet, items, left);
      if (left < 0) break;
    }
  } finally {
    props.deleteProperty(LO_LEASE_KEY);
  }
}

// 束が 400 のとき、1件ずつ送り直して受け付けない本文だけ「送信不可」にする。戻り値は残りの push 回数。
// 途中で 400 以外に当たったときと、2件以上がすべて 400 のとき（本文ではなく別の原因とみる）は
// 届いた分だけ印を付け、残りは「待ち」のまま -1 を返して止める。
function loIsolateRejected_(cfg, sheet, items, left) {
  var rows = items.filter(function(it) { return it.row > 0; });
  var results = [];
  var stopped = false;
  for (var k = 0; k < rows.length; k++) {
    if (left < 1) break;
    var r = loSend_(cfg, [rows[k].text]);
    left--;
    if (!loDelivered_(r) && (r.code !== 400 || loDestinationError_(r))) { stopped = true; break; }
    results.push({ row: rows[k].row, r: r });
  }
  var allRejected = rows.length > 1 && results.length === rows.length &&
    !results.some(function(x) { return loDelivered_(x.r); });
  if (allRejected) console.error('LINE が本文をすべて受け付けません（400）: ' + results[0].r.body);
  results.forEach(function(x) {
    if (loDelivered_(x.r)) loMark_(sheet, [x.row], LO_SENT, '');
    else if (!stopped && !allRejected) loMark_(sheet, [x.row], LO_REJECTED, 'HTTP 400 ' + x.r.body);
  });
  return stopped || allRejected ? -1 : left;
}

// ---- 日次レポートに添える節 ----
// always=false のときは、8割を超えたか、直近の「送信不可」があるときだけ返す。
// 取れなければ空文字（レポート自体は止めない）。
function lineQuotaReportSection_(always) {
  var q = null;
  var box = { waiting: 0, rejected: 0 };
  try {
    var cfg = loConfig_();
    if (!cfg) return '';
    try { box = loOutboxCounts_(new Date()); } catch (e) { Logger.log('LINE送信待ちを数えられませんでした: ' + e); }
    try {
      q = lineQuota_(cfg);
      if (!q.limited || !(q.limit > 0)) q = null;
    } catch (e) {
      Logger.log('公式LINEの送信数を取れませんでした: ' + e);
    }
    var warn = !!q && q.used / q.limit >= LO_WARN_RATIO;
    if (!always && !warn && !box.rejected) return '';

    var lines = [];
    if (q) {
      lines.push('■公式LINEの送信数（今月・アフィリンクとアポ報告の合計）');
      lines.push(loNum_(q.used) + ' / ' + loNum_(q.limit) + '通（' + Math.floor(q.used / q.limit * 100) + '%）');
      var pace = loPaceLine_(q, new Date());
      if (pace) lines.push(pace);
      if (warn) lines.push('【注意】8割を超えました。上限に達したあとの通知は、翌月1日の0時台にまとめて届きます。');
    }
    if (box.waiting || box.rejected) {
      if (!lines.length) lines.push('■公式LINEの送信');
      if (box.waiting) lines.push('送れずに待っている通知: ' + box.waiting + '件（送信枠が空きしだい送ります）');
      if (box.rejected) lines.push('【要確認】LINE が受け付けなかった通知: ' + box.rejected + '件（直近' + LO_REJECTED_DAYS + '日・シート「' + LO_SHEET_NAME + '」）');
    }
    return lines.join('\n');
  } catch (e) {
    Logger.log('公式LINEの節を作れませんでした: ' + e);
    return '';
  }
}

function loOutboxCounts_(now) {
  var counts = { waiting: 0, rejected: 0 };
  var sheet = loSpreadsheet_().getSheetByName(LO_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return counts;
  var since = now.getTime() - LO_REJECTED_DAYS * 864e5;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
    if (r[1] === LO_WAITING) counts.waiting++;
    if (r[1] === LO_REJECTED && r[0] instanceof Date && r[0].getTime() >= since) counts.rejected++;
  });
  return counts;
}

function loNum_(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// 今月ここまでの平均ペースで、上限に届く日を見積もる。月初2日間は再送分が混じるので出さない。
function loPaceLine_(q, now) {
  if (q.used >= q.limit) return '上限に達しています。';
  var ym = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM').split('-');
  var y = Number(ym[0]), m = Number(ym[1]);
  var start = new Date(y + '-' + ('0' + m).slice(-2) + '-01T00:00:00+09:00');
  var end = m === 12 ? new Date((y + 1) + '-01-01T00:00:00+09:00')
                     : new Date(y + '-' + ('0' + (m + 1)).slice(-2) + '-01T00:00:00+09:00');
  var days = (now.getTime() - start.getTime()) / 864e5;
  if (days < 2 || q.used <= 0) return '';
  var hit = new Date(now.getTime() + (q.limit - q.used) / (q.used / days) * 864e5);
  return hit.getTime() >= end.getTime()
    ? 'このペースなら月末まで足ります。'
    : 'このペースだと ' + Utilities.formatDate(hit, 'Asia/Tokyo', 'M/d') + ' ごろ上限に達します。';
}

// ---- 初回に1回だけ実行（GASエディタ）: シートと1時間おきのトリガーを用意する ----
function setupLineOutbox() {
  lineOutboxSheet_(loSpreadsheet_());
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === LO_FLUSH_HANDLER) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(LO_FLUSH_HANDLER).timeBased().everyHours(1).create();
  Logger.log('シート「' + LO_SHEET_NAME + '」と1時間おきのトリガーを用意しました');
  previewLineOutbox();
}

// ---- 送らずに状態だけ見る（GASエディタから実行）----
function previewLineOutbox() {
  var waiting = lineOutboxWaiting_(loSpreadsheet_().getSheetByName(LO_SHEET_NAME));
  Logger.log('送信待ち: ' + waiting.length + '件');
  var cfg = loConfig_();
  if (!cfg) { Logger.log('LINE の設定がありません'); return; }
  Logger.log('あと送れる push 回数: ' + linePushBudget_(cfg));
  Logger.log(lineQuotaReportSection_(true) || '（今月の送信数は取れませんでした）');
}
