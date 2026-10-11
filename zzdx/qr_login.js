// 转转学堂 飞书扫码登录：终端打印二维码 → 扫码成功后保存 cookie
// 由 抓错题.py 在登录态失效时调用
const { chromium } = require('playwright-core');
const fs = require('fs');
const path = require('path');
const QRCode = require('qrcode');

const ZDIR = __dirname;
const EXE = process.env.HOME + '/Library/Caches/ms-playwright/chromium-1234/chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing';
const URL = 'https://zzsso.zhuanspirit.com/user/login?host=zzsso.zhuanspirit.com&appType=1&scheme=https&targetHost=wxdaxue.zhuanspirit.com&urlParam=%2Fredirect%3Fredirect%3Dhttps%3A%2F%2Fwxdaxue.zhuanspirit.com%252F';
const STATE = path.join(ZDIR, 'state.json');
const COOKIE = path.join(ZDIR, 'cookie.txt');

function printQR(payload) {
  const qr = QRCode.create(payload, { errorCorrectionLevel: 'L' });
  const size = qr.modules.size;
  const M = 2;
  const blank = ' '.repeat(size + M * 2);
  const rows = [];
  for (let y = 0; y < size; y += 2) {
    let line = ' '.repeat(M);
    for (let x = 0; x < size; x++) {
      const top = qr.modules.get(x, y);
      const bot = y + 1 < size ? qr.modules.get(x, y + 1) : false;
      line += top && bot ? '█' : top ? '▀' : bot ? '▄' : ' ';
    }
    line += ' '.repeat(M);
    rows.push(line);
  }
  console.log('\n请用飞书 App 扫码登录转转学堂（8 分钟内有效）：\n');
  console.log(blank);
  rows.forEach(r => console.log(r));
  console.log(blank);
}

(async () => {
  const browser = await chromium.launch({ executablePath: EXE, headless: true });
  const ctx = await browser.newContext({ viewport: { width: 1200, height: 900 } });
  const page = await ctx.newPage();
  let printed = false;
  // 拦截飞书 qrlogin init 响应，直接拿 token 生成二维码（比截图更快更及时）
  page.on('response', async r => {
    try {
      if (!printed && /accounts\/qrlogin\/init/.test(r.url())) {
        const body = await r.text();
        const m = body.match(/"token"\s*:\s*"([^"]+)"/);
        if (m) {
          printQR(JSON.stringify({ qrlogin: { token: m[1] } }));
          printed = true;
        }
      }
    } catch {}
  });

  await page.goto(URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(5000);
  if (!printed) console.log('未捕获二维码 token，登录页可能已改版');

  const deadline = Date.now() + 8 * 60 * 1000;
  let ok = false;
  while (Date.now() < deadline) {
    const u = page.url();
    if (/wxdaxue\.zhuanspirit\.com/.test(u) && !/zzsso/.test(u)) { ok = true; break; }
    await page.waitForTimeout(3000);
  }
  if (!ok) { console.log('SCAN TIMEOUT（8 分钟未扫码）'); await browser.close(); process.exit(2); }
  await page.waitForTimeout(8000);
  await ctx.storageState({ path: STATE });
  const cookies = await ctx.cookies();
  const zz = cookies.filter(c => c.domain.includes('zhuanspirit'));
  fs.writeFileSync(COOKIE, zz.map(c => `${c.name}=${c.value}`).join('; '));
  console.log('登录成功，cookie 已保存（' + zz.length + ' 条）');
  await browser.close();
})().catch(e => { console.error('FATAL', e); process.exit(1); });
