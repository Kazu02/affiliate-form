// 公式LINEから来た申請は、URLの lc（参照）を送信データに載せる（2026-11-16）。
// 名前に頼らず「誰がどの案件をやったか」を公式LINE側で分かるようにするため。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { test } = require('node:test');
const { chromium } = require('playwright');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const REF = 'a'.repeat(32) + '0123456789abcdef';
// 1x1 の PNG
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=', 'base64');

async function submitWith(browser, query) {
  const context = await browser.newContext();
  await context.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: 'https://form.example.test' });
  const page = await context.newPage();
  const posts = [];
  await page.route('https://form.example.test/**', route => {
    if (route.request().url().includes('config.js')) return route.fulfill({ contentType: 'text/javascript', body: 'const CONFIG = { gasUrl: "https://api.example.test/exec" };' });
    return route.fulfill({ contentType: 'text/html', body: html });
  });
  await page.route('https://api.example.test/**', route => {
    if (route.request().method() === 'POST') {
      posts.push(JSON.parse(route.request().postData() || '{}'));
      return route.fulfill({ status: 200, body: '{}' });
    }
    return route.fulfill({ contentType: 'application/json', body: JSON.stringify({
      formName: 'nomukomu', formTitle: 'Test form', affiliateUrl: 'https://affiliate.example.test/link?a=1', affiliateButtonText: '申し込む',
      fields: [{ id: 'name', label: '氏名', type: 'text' }], suspended: false
    }) });
  });
  await page.goto('https://form.example.test/?form=nomukomu' + query);
  await page.waitForFunction(() => document.getElementById('loading').classList.contains('hidden'));
  await page.locator('#name').fill('テスト');
  await page.locator('#affiliateBtn').click();
  await page.locator('#screenshotInput').setInputFiles({ name: 'done.png', mimeType: 'image/png', buffer: PNG });
  await page.locator('#submitBtn').click();
  await page.waitForFunction(() => document.getElementById('successScreen').classList.contains('show'));
  await context.close();
  return posts.filter(p => p.action !== 'clickPing');
}

test('a valid lc reference is sent with the application', async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const sent = await submitWith(browser, '&lc=' + REF.toUpperCase());
    assert.equal(sent.length, 1);
    assert.equal(sent[0].lineRef, REF);
  } finally { await browser.close(); }
});

test('a missing or malformed lc is not sent', async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    for (const q of ['', '&lc=abc', '&lc=' + REF + 'zz', '&lc=%3Cscript%3E']) {
      const sent = await submitWith(browser, q);
      assert.equal(sent.length, 1);
      assert.equal('lineRef' in sent[0], false, q);
    }
  } finally { await browser.close(); }
});
