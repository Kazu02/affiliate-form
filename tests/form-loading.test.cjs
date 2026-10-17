const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const { chromium } = require('playwright');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const script = html.match(/<script>\s*([\s\S]*?)<\/script>/)[1].replace(/\binit\(\);\s*$/, '');

function loader(fetch) {
  // Accelerate the production timeout and retry delay; exercise the real loader.
  const context = vm.createContext({
    fetch, AbortController, clearTimeout,
    document: { getElementById: () => ({ addEventListener() {} }) },
    window: { addEventListener() {} },
    setTimeout: (fn, ms) => setTimeout(fn, ms === 25000 ? 15 : 1)
  });
  vm.runInContext(script, context);
  return context.loadFormConfig;
}
const response = (status, body = { fields: [] }) => ({ ok: status === 200, status, json: async () => body });

test('transient 404 retries the same form URL and returns the fresh config', async () => {
  const calls = [];
  const config = { fields: [], affiliateUrl: 'https://example.test/new' };
  const load = loader(async (url, options) => {
    calls.push({ url, options });
    return calls.length === 1 ? response(404) : response(200, config);
  });
  assert.equal(await load('https://example.test/config?form=nomukomu&ag=ag01'), config);
  assert.equal(calls.length, 2);
  assert.equal(calls[0].url, calls[1].url);
  assert.ok(calls.every(c => c.options.method === 'GET' && c.options.cache === 'no-store'));
});

test('a stalled fetch is aborted and retried once', async () => {
  const signals = [];
  const load = loader(async (url, options) => {
    signals.push(options.signal);
    return signals.length === 1 ? new Promise(() => {}) : response(200);
  });
  assert.ok((await load('https://example.test/config')).fields);
  assert.equal(signals.length, 2);
  assert.equal(signals[0].aborted, true);
});

test('stalled response bodies stop after two attempts', async () => {
  let calls = 0;
  const load = loader(async () => {
    calls++;
    return { ok: true, json: () => new Promise(() => {}) };
  });
  await assert.rejects(load('https://example.test/config'), /タイムアウト/);
  assert.equal(calls, 2);
});

test('network failures retry, but permanent HTTP errors do not', async () => {
  let calls = 0;
  const load = loader(async () => { calls++; throw new TypeError('network'); });
  await assert.rejects(load('https://example.test/config'), /network/);
  assert.equal(calls, 2);
  calls = 0;
  await assert.rejects(loader(async () => { calls++; return response(403); })('https://example.test/config'));
  assert.equal(calls, 1);
});

test('suspended and application-error responses are returned unchanged', async () => {
  for (const config of [{ suspended: true, fields: [] }, { error: 'unknown form' }]) {
    let calls = 0;
    assert.equal(await loader(async () => { calls++; return response(200, config); })('https://example.test/config'), config);
    assert.equal(calls, 1);
  }
});

test('browser recovers after 404, keeps referral parameters, and exits loading after repeated failure', async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    for (const recover of [true, false]) {
      const page = await browser.newPage();
      let requests = 0;
      const urls = [];
      await page.addInitScript(() => {
        const original = window.setTimeout.bind(window);
        window.setTimeout = (fn, ms, ...args) => original(fn, ms === 25000 ? 100 : ms === 1000 ? 1 : ms, ...args);
      });
      await page.route('https://form.example.test/**', route => {
        if (route.request().url().includes('config.js')) return route.fulfill({ contentType: 'text/javascript', body: 'const CONFIG = { gasUrl: "https://api.example.test/config" };' });
        return route.fulfill({ contentType: 'text/html', body: html });
      });
      await page.route('https://api.example.test/**', route => {
        requests++;
        urls.push(route.request().url());
        if (!recover || requests === 1) return route.fulfill({ status: 404, body: 'unavailable' });
        return route.fulfill({ contentType: 'application/json', body: JSON.stringify({
          formName: 'nomukomu', formTitle: 'Test form', affiliateUrl: 'https://affiliate.example.test/',
          fields: [{ id: 'referrer', label: '紹介者', type: 'text' }], suspended: false
        }) });
      });
      await page.goto('https://form.example.test/?form=nomukomu-s&ag=ag01');
      await page.waitForFunction(() => document.getElementById('loading').classList.contains('hidden'));
      assert.equal(requests, 2);
      assert.equal(new URL(urls[1]).searchParams.get('form'), 'nomukomu');
      assert.equal(new URL(urls[1]).searchParams.get('ag'), 'ag01');
      if (recover) {
        assert.equal(await page.locator('#formWrap').isVisible(), true);
        assert.equal(await page.locator('#referrer').inputValue(), '菅原貴博');
      } else {
        assert.equal(await page.locator('#errorBanner').isVisible(), true);
        assert.equal(await page.locator('#formWrap').isVisible(), false);
      }
      await page.close();
    }
  } finally {
    await browser.close();
  }
});
