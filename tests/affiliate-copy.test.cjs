// アフィリエイトURLは「リンクで開く」のでなく「コピーしてブラウザに貼る」（2026-11-16）。
// このページから踏ませると参照元がASPへ渡り、成果が付かないことがあるため。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { test } = require('node:test');
const { chromium } = require('playwright');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const AFF = 'https://affiliate.example.test/link?a=1';

async function openForm(browser, { clipboard = true } = {}) {
  const context = await browser.newContext();
  if (clipboard) await context.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: 'https://form.example.test' });
  const page = await context.newPage();
  const pings = [];
  await page.route('https://form.example.test/**', route => {
    if (route.request().url().includes('config.js')) return route.fulfill({ contentType: 'text/javascript', body: 'const CONFIG = { gasUrl: "https://api.example.test/exec" };' });
    return route.fulfill({ contentType: 'text/html', body: html });
  });
  await page.route('https://api.example.test/**', route => {
    if (route.request().method() === 'POST') {
      pings.push(JSON.parse(route.request().postData() || '{}'));
      return route.fulfill({ status: 200, body: '{}' });
    }
    return route.fulfill({ contentType: 'application/json', body: JSON.stringify({
      formName: 'nomukomu', formTitle: 'Test form', affiliateUrl: AFF, affiliateButtonText: '申し込む',
      fields: [{ id: 'name', label: '氏名', type: 'text' }], suspended: false
    }) });
  });
  await page.goto('https://form.example.test/?form=nomukomu');
  await page.waitForFunction(() => document.getElementById('loading').classList.contains('hidden'));
  return { context, page, pings };
}

test('URL is shown for copying, not opened as a link, and the copy is recorded once', async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const { context, page, pings } = await openForm(browser);
    // 画面にASPへのリンクが1本も無い（踏ませない）
    assert.equal(await page.locator('a[href*="affiliate.example.test"]').count(), 0);
    const btn = page.locator('#affiliateBtn');
    assert.equal(await btn.evaluate(el => el.tagName), 'BUTTON');
    assert.equal(await btn.textContent(), 'タップしてURLをコピー');

    const shown = await page.locator('#affiliateUrlBox').inputValue();
    const clickId = await page.locator('#clickId').inputValue();
    assert.match(clickId, /^[0-9a-f]{32}$/);
    assert.equal(shown, AFF + '&param1=' + clickId);

    await btn.click();
    assert.equal(await page.evaluate(() => navigator.clipboard.readText()), shown);
    assert.match(await btn.textContent(), /コピーしました/);
    assert.equal(await page.locator('#copyFail').isVisible(), false);
    assert.equal(await page.locator('#clickTimeDisplay').isVisible(), true);
    const firstTime = await page.locator('#clickTime').inputValue();
    assert.ok(firstTime);
    await page.waitForTimeout(100);
    assert.equal(pings.length, 1);
    assert.equal(pings[0].action, 'clickPing');
    assert.equal(pings[0].clickId, clickId);

    // 2回目のコピーは時刻を上書きせず、記録も増やさない
    await page.waitForTimeout(20);
    await btn.click();
    assert.equal(await page.locator('#clickTime').inputValue(), firstTime);
    await page.waitForTimeout(100);
    assert.equal(pings.length, 1);

    // ブラウザへ行って戻ってきた（再読み込み）でも同じURL・同じ記録のまま
    await page.reload();
    await page.waitForFunction(() => document.getElementById('loading').classList.contains('hidden'));
    assert.equal(await page.locator('#affiliateUrlBox').inputValue(), shown);
    assert.equal(await page.locator('#clickTime').inputValue(), firstTime);
    await context.close();
  } finally {
    await browser.close();
  }
});

test('when copying fails the customer is told to copy by hand, and the press is still recorded', async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const { context, page } = await openForm(browser, { clipboard: false });
    await page.evaluate(() => {
      Object.defineProperty(navigator, 'clipboard', { value: { writeText: () => Promise.reject(new Error('denied')) } });
      document.execCommand = () => false;
    });
    await page.locator('#affiliateBtn').click();
    assert.equal(await page.locator('#copyFail').isVisible(), true);
    assert.equal(await page.locator('#affiliateBtn').textContent(), 'タップしてURLをコピー');
    assert.ok(await page.locator('#clickTime').inputValue());
    await context.close();
  } finally {
    await browser.close();
  }
});

test('submitting before copying asks the customer to copy the URL first', async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const { context, page } = await openForm(browser);
    await page.locator('#name').fill('テスト');
    let message = '';
    page.once('dialog', d => { message = d.message(); d.dismiss(); });
    await page.locator('#submitBtn').click();
    await page.waitForTimeout(100);
    assert.match(message, /URLをコピー/);
    await context.close();
  } finally {
    await browser.close();
  }
});
