// =============================================
// 2026-08-20 の請求漏れ10件を承認へ直す（1回だけ実行する使い捨て）
//
// ASP獲得ログとの突合で「ASPは承認・自社は未承認」と判明した10件。
// 報酬合計 97,000円。差0〜2秒・案件名一致で同一成果と確認済み。
//
// 汎用の突合は AspReconcile.gs（ASP獲得ログを取り込んで reconcileAspLog）。
// こちらは検証済みの10件を人手のチェック無しで直すための一度きりの関数。
//
// **書く前に必ず検証する。** 行番号だけを信じて書かない。
// (1) その行のクリック日時がASPの記録と120秒以内か
// (2) いま承認になっていないか
// どちらか外れたら書かずにスキップし、理由を返す。
// =============================================

// [シート名, 行, ASPのクリック日時(JST), ASPの広告名, 報酬額]
const ASP_FIX_20260820 = [
  ["設定_不動産一括査定イエイ", 74, "2026-07-30 18:42:15", "イエイ",                 10000],
  ["設定_いえカツLIFE",        13, "2026-07-29 22:44:26", "いえカツLIFE",           15000],
  ["設定_リビンマッチ",         51, "2026-07-21 22:24:05", "リビンマッチ",            8000],
  ["設定_不動産一括査定イエイ", 51, "2026-07-17 15:37:10", "イエイ",                 10000],
  ["設定_不動産一括査定イエイ", 50, "2026-07-13 18:21:43", "イエイ",                 10000],
  ["設定_リビンマッチ",         46, "2026-07-13 18:08:04", "リビンマッチ",            8000],
  ["設定_リビンマッチ",         45, "2026-07-12 20:10:17", "リビンマッチ",            8000],
  ["設定_リビンマッチ",         23, "2026-06-25 21:22:19", "リビンマッチ",            8000],
  ["設定_ポケットリサーチ",      3, "2026-05-13 22:50:16", "ポケットリサーチ（特別単価）", 10000],
  ["設定_スマモニ",            10, "2026-05-12 16:55:28", "スマモニ（特別単価）",       10000]
];

// dryRun=true なら何も書かずに、書く対象と検証結果だけ返す。
function applyAspFix20260820(dryRun) {
  const ss = getOrCreateSpreadsheet();
  const report = [];
  let applied = 0, skipped = 0;

  ASP_FIX_20260820.forEach(function (spec) {
    const sheetName = spec[0], row = spec[1], aspClick = spec[2], adName = spec[3], reward = spec[4];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { skipped++; report.push("スキップ " + sheetName + " 行" + row + ": シートが無い"); return; }

    const lastCol = sheet.getLastColumn();
    if (lastCol < ANSWER_START_COL) { skipped++; report.push("スキップ " + sheetName + " 行" + row + ": 回答列が無い"); return; }
    const width = lastCol - ANSWER_START_COL + 1;
    const headers = sheet.getRange(1, ANSWER_START_COL, 1, width).getValues()[0].map(String);
    const iClick = headers.indexOf("クリック日時");
    const iAppr  = headers.indexOf("承認");
    const iName  = headers.indexOf("お名前");
    if (iClick < 0 || iAppr < 0) { skipped++; report.push("スキップ " + sheetName + " 行" + row + ": 見出しが揃わない"); return; }

    const vals = sheet.getRange(row, ANSWER_START_COL, 1, width).getValues()[0];

    // 検証(1) クリック日時がASPの記録と120秒以内か
    const ownMs = aspToMillis_(vals[iClick]);
    const aspMs = aspToMillis_(aspClick);
    const diff = Math.abs(ownMs - aspMs) / 1000;
    if (!ownMs || diff > 120) {
      skipped++;
      report.push("スキップ " + sheetName + " 行" + row + ": クリック日時が合わない（差 " +
                  (ownMs ? Math.round(diff) + "秒" : "空") + "）。行がずれた可能性がある");
      return;
    }

    // 検証(2) 既に承認になっていないか
    const cur = aspApprovalOf_(vals[iAppr]);
    if (cur === "承認") {
      skipped++;
      report.push("スキップ " + sheetName + " 行" + row + ": すでに承認済み");
      return;
    }

    const who = iName >= 0 ? String(vals[iName] || "") : "";
    if (dryRun) {
      report.push("対象 " + sheetName + " 行" + row + " (" + who + ") " + cur + " → 承認 / " +
                  adName + " " + reward + "円 / 差" + Math.round(diff) + "秒");
      applied++;
      return;
    }
    sheet.getRange(row, ANSWER_START_COL + iAppr).setValue("⭕");
    applied++;
    report.push("修正 " + sheetName + " 行" + row + " (" + who + ") " + cur + " → 承認");
  });

  if (!dryRun) SpreadsheetApp.flush();
  const msg = (dryRun ? "【下見】" : "【適用】") + " 対象 " + applied + " 件 / スキップ " + skipped + " 件\n" +
              report.join("\n");
  Logger.log(msg);
  return msg;
}

// まず下見（何も書かない）
function previewAspFix20260820() {
  return applyAspFix20260820(true);
}

// 実際に直す
function runAspFix20260820() {
  return applyAspFix20260820(false);
}

function aspFix20260820FromMenu() {
  const ui = SpreadsheetApp.getUi();
  const pre = applyAspFix20260820(true);
  const ok = ui.alert("請求漏れ10件を承認へ直します。まず下見の結果です。\n\n" + pre +
                      "\n\n適用してよろしいですか。", ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  ui.alert(applyAspFix20260820(false) +
           "\n\n続けて、広告主シートの 202605 / 202606 / 202607 を再生成してください。");
}
