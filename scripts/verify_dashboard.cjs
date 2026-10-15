/**
 * W4 · 看板功能验证（真实浏览器驱动，非静态断言）
 * 1) 打开看板，确认 8 个图表 canvas 全部渲染
 * 2) 读取 KPI，与 Python 口径逐项比对
 * 3) 真实点击筛选器，验证联动生效
 * 4) 截图留证
 */
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..');
const CFG_PATH = path.join(ROOT, 'config.json');
const PREFIX = (fs.existsSync(CFG_PATH) && JSON.parse(fs.readFileSync(CFG_PATH, 'utf-8'))['输出前缀']) || '超市运营分析';
const FILE = path.join(ROOT, `${PREFIX}_交互看板.html`);
const SHOT = path.join(ROOT, 'outputs', 'dashboard_preview.png');
const METRICS = JSON.parse(fs.readFileSync(path.join(ROOT, 'outputs', 'metrics.json'), 'utf-8'));

const toURL = p => 'file:///' + p.replace(/\\/g, '/').split('/').map((s, i) => i === 0 ? s : encodeURIComponent(s)).join('/');

let pass = 0, fail = 0;
function check(name, actual, expected, tol = 0.06) {
  const a = String(actual).replace(/[^\d.\-]/g, '');
  const ok = expected === null ? true : Math.abs(parseFloat(a) - expected) <= tol;
  console.log(`  ${ok ? '✅' : '❌'} ${name}：页面 = ${actual}${expected === null ? '' : ` | 期望 ≈ ${expected}`}`);
  ok ? pass++ : fail++;
}

(async () => {
  const CHROME = process.env.PW_CHROME ||
    path.join(process.env.LOCALAPPDATA || '', 'ms-playwright', 'chromium-1234', 'chrome-win64', 'chrome.exe');
  if (!fs.existsSync(CHROME)) { console.error('未找到浏览器：' + CHROME); process.exit(2); }
  const browser = await chromium.launch({ executablePath: CHROME });
  const page = await browser.newPage({ viewport: { width: 1480, height: 1000 } });
  const errors = [];
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('pageerror', e => errors.push('pageerror: ' + e.message));

  await page.goto(toURL(FILE), { waitUntil: 'load' });
  await page.waitForSelector('#kpis .kpi', { timeout: 15000 });
  await page.waitForTimeout(2500);

  const B = METRICS.baseline;
  console.log('='.repeat(72));
  console.log('W4 · 看板功能验证（真实浏览器）');
  console.log('='.repeat(72));

  // ---- 1. 图表渲染 ----
  const canvasCount = await page.$$eval('.chart canvas', els => els.length);
  console.log(`\n【1】图表渲染：${canvasCount} 个 canvas`);
  check('8 个图表全部渲染', canvasCount, 8, 0);

  // ---- 2. KPI 与 Python 口径比对（全量筛选）----
  const kpis = await page.$$eval('#kpis .kpi', els => els.map(e => ({
    k: e.querySelector('.k').textContent.trim(),
    v: e.querySelector('.v').textContent.trim()
  })));
  const K = Object.fromEntries(kpis.map(x => [x.k, x.v]));
  console.log('\n【2】KPI 与 Python 口径比对（全量）');
  check('销售额（万）', K['销售额'], +(B.销售额 / 10000).toFixed(1));
  check('利润（万）', K['利润'], +(B.利润 / 10000).toFixed(1));
  check('利润率', K['利润率'], B.利润率, 0.02);
  check('订单数', K['订单数'], B.订单数, 0);
  check('客户数', K['客户数'], B.客户数, 0);
  check('明细行', K['明细行'], B.明细行, 0);
  check('退货率', K['退货率'], METRICS.M6_return.真实退货率, 0.02);

  // ---- 3. 筛选联动（真实点击）----
  console.log('\n【3】筛选联动（真实点击：类别 → 仅「办公用品」）');
  // 注意：每次 render 会重建筛选器 DOM，因此每点一次都要重新定位
  const catChip = t => page.locator('#filters .frow').nth(2).locator('.chip').filter({ hasText: t }).first();
  for (const t of ['家具', '技术']) {
    await catChip(t).click();
    await page.waitForTimeout(600);
  }
  await page.waitForTimeout(700);
  const k2 = await page.$$eval('#kpis .kpi', els => els.map(e => ({
    k: e.querySelector('.k').textContent.trim(), v: e.querySelector('.v').textContent.trim()
  })));
  const K2 = Object.fromEntries(k2.map(x => [x.k, x.v]));
  const office = METRICS.M3_category.find(x => x.类别 === '办公用品');
  check('筛选后销售额（万）', K2['销售额'], +(office.销售额 / 10000).toFixed(1));
  check('筛选后利润率', K2['利润率'], office.利润率, 0.05);

  // 还原
  await page.click('#filters .reset');
  await page.waitForTimeout(800);
  const k3 = await page.$$eval('#kpis .kpi', els => els.map(e => e.querySelector('.v').textContent.trim()));
  console.log('\n【4】重置筛选');
  check('重置后销售额（万）回到全量', k3[0], +(B.销售额 / 10000).toFixed(1));

  // ---- 5. 亏损客户表 ----
  const lossRows = await page.$$eval('#losstbl tbody tr', els => els.length);
  console.log('\n【5】亏损客户表');
  check('Top 15 行已渲染', lossRows, 15, 0);

  // ---- 6. 截图 ----
  await page.screenshot({ path: SHOT, fullPage: true });
  console.log(`\n【6】截图：${SHOT}`);

  // ---- 7. 控制台错误 ----
  console.log('\n【7】控制台错误');
  if (errors.length) {
    console.log('  ❌ 发现 ' + errors.length + ' 条：');
    errors.slice(0, 8).forEach(e => console.log('     - ' + e.slice(0, 160)));
    fail++;
  } else {
    console.log('  ✅ 无 JS 报错');
    pass++;
  }

  await browser.close();
  console.log('\n' + '='.repeat(72));
  console.log(`结果：通过 ${pass} 项，失败 ${fail} 项`);
  console.log('='.repeat(72));
  process.exit(fail ? 1 : 0);
})();
