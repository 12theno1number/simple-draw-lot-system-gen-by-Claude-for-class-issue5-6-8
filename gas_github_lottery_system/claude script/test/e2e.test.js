// P0 端對端測試（以真實瀏覽器操作 index.html）
// 需要 Playwright：npm i -D playwright && npx playwright install chromium
// 執行：node --test "claude script/test/e2e.test.js"
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
const { pathToFileURL } = require('node:url');
const { chromium } = require('playwright');

const PAGE = pathToFileURL(path.join(__dirname, '..', 'project', 'index.html')).href;
const FIXTURE = path.join(__dirname, 'fixtures', '測試用抽獎後台.xlsx');

let browser, ctx, page;
test.before(async () => { browser = await chromium.launch(); });
test.after(async () => { await browser.close(); });
test.beforeEach(async () => {
  ctx = await browser.newContext({ acceptDownloads: true });
  page = await ctx.newPage();
  page.setDefaultTimeout(5000);
  page.on('dialog', d => { throw new Error('不應出現瀏覽器彈窗：' + d.message()); });
  await page.goto(PAGE);
});
test.afterEach(async () => { await ctx.close(); });

// 以檔案內容傳入（Playwright 在部分環境下無法用中文路徑觸發 change 事件）
const importFixture = () => page.setInputFiles('#file', {
  name: path.basename(FIXTURE),
  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  buffer: fs.readFileSync(FIXTURE),
});
const prizeOptions = () => page.$$eval('#prizeSel option', os => os.map(o => o.textContent));
const recordCount = () => page.$$eval('#results li', l => l.length);

test('匯入 Excel：自動讀到 18 人與 4 個獎項', async () => {
  await importFixture();
  await page.waitForFunction(() => document.querySelectorAll('#prizeSel option').length === 4);
  const opts = await prizeOptions();
  assert.match(opts[0], /^頭獎.*（剩 10 \/ 10）$/);
  assert.match(await page.textContent('#count'), /18 位/);
  assert.equal(await page.$eval('#items', e => e.readOnly), true);
});

test('抽獎後名額減少、紀錄顯示獎項與完整日期', async () => {
  await importFixture();
  await page.waitForSelector('#prizeSel option', { state: 'attached' });
  await page.uncheck('#anim');
  await page.fill('#n', '3');
  await page.click('#drawBtn');
  assert.equal(await recordCount(), 3);
  assert.match((await prizeOptions())[0], /剩 7 \/ 10/);
  assert.match(await page.textContent('#results li .t'), /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/);
});

test('B1/B2：動畫進行中，清除紀錄、匯入、切換獎項、名單都被鎖住', async () => {
  await importFixture();
  await page.waitForSelector('#prizeSel option', { state: 'attached' });
  await page.uncheck('#anim'); await page.click('#drawBtn');  // 先有一筆紀錄
  await page.check('#anim');
  await page.click('#drawBtn');
  assert.equal(await page.isDisabled('#resetBtn'), true);
  assert.equal(await page.isDisabled('#prizeSel'), true);
  assert.equal(await page.isDisabled('#drawBtn'), true);
  assert.equal(await page.$eval('#drop', e => e.classList.contains('disabled')), true);
  await page.waitForFunction(() => !document.getElementById('stage').classList.contains('rolling'));
  assert.equal(await recordCount(), 2);
  assert.equal(await page.isDisabled('#resetBtn'), false);
});

test('B3：清除紀錄需要頁面內二次確認，按取消不會清除', async () => {
  await importFixture();
  await page.waitForSelector('#prizeSel option', { state: 'attached' });
  await page.uncheck('#anim'); await page.click('#drawBtn');
  await page.click('#resetBtn');
  assert.equal(await page.isVisible('#resetConfirm'), true);
  await page.click('#resetNo');
  assert.equal(await recordCount(), 1);
  await page.click('#resetBtn'); await page.click('#resetYes');
  assert.equal(await recordCount(), 0);
  assert.match((await prizeOptions())[0], /剩 10 \/ 10/);
});

test('G3：重新整理後紀錄與名額都還在', async () => {
  await importFixture();
  await page.waitForSelector('#prizeSel option', { state: 'attached' });
  await page.uncheck('#anim'); await page.fill('#n', '2'); await page.click('#drawBtn');
  await page.reload();
  assert.equal(await recordCount(), 2);
  assert.match((await prizeOptions())[0], /剩 8 \/ 10/);
  assert.match(await page.textContent('#msg'), /已還原上次的 2 筆/);
});

test('已有紀錄時名單鎖定，不能換檔案', async () => {
  await importFixture();
  await page.waitForSelector('#prizeSel option', { state: 'attached' });
  await page.uncheck('#anim'); await page.click('#drawBtn');
  assert.equal(await page.isDisabled('#clearItems'), true);
  await importFixture(); // 直接對 input 設檔案也要被擋下
  assert.match(await page.textContent('#msg'), /請先「清除全部紀錄」/);
  assert.equal(await recordCount(), 1);
});

test('匯出 CSV：含新欄位與 BOM', async () => {
  await importFixture();
  await page.waitForSelector('#prizeSel option', { state: 'attached' });
  await page.uncheck('#anim'); await page.click('#drawBtn');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#exportBtn')]);
  const text = fs.readFileSync(await dl.path(), 'utf8');
  assert.ok(text.startsWith('﻿序號,獎項,姓名,系所,年級,學號,抽出時間'));
  assert.equal(text.trim().split(/\r?\n/).length, 2);
});

test('手動輸入名單：同名的兩人都抽得到，抽完即停', async () => {
  await page.fill('#items', '王小明\n王小明\n陳大文');
  await page.uncheck('#anim');
  await page.fill('#n', '5');
  await page.click('#drawBtn');
  assert.equal(await recordCount(), 3);
  assert.equal(await page.isDisabled('#drawBtn'), true);
  const names = await page.$$eval('#results li', l => l.map(x => x.textContent));
  assert.equal(names.filter(n => n.includes('王小明')).length, 2);
});
