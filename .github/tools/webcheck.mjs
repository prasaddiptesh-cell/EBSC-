// Opens the live website in a real browser (Chromium), prints what went wrong and saves screenshots.
import { chromium } from 'playwright';
const url = process.env.WEB_URL || 'https://easybuys-web.easybuys.workers.dev';
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1280, height: 860 } });
const log = [];
page.on('console', (m) => log.push(`console.${m.type()}: ${m.text()}`.slice(0, 400)));
page.on('pageerror', (e) => log.push(`PAGE ERROR: ${String(e && e.stack || e).slice(0, 1500)}`));
page.on('requestfailed', (r) => log.push(`request failed: ${r.url().slice(0, 120)} ${r.failure() && r.failure().errorText}`));
const res = await page.goto(url, { waitUntil: 'load', timeout: 60000 });
log.push(`status ${res && res.status()} headers coop=${res && res.headers()['cross-origin-opener-policy']} coep=${res && res.headers()['cross-origin-embedder-policy']}`);
log.push(`crossOriginIsolated=${await page.evaluate(() => self.crossOriginIsolated)}`);
await page.waitForTimeout(20000);
await page.screenshot({ path: 'shot1.png' });
const text = await page.evaluate(() => document.body.innerText.slice(0, 1500));
log.push('TEXT: ' + text.replace(/\n+/g, ' | '));
console.log(log.join('\n'));
await browser.close();
