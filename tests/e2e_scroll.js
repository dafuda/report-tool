/* 修改分類後清單不跳回最上面：node tests/e2e_scroll.js */
const path = require('path');
const { chromium } = require('playwright');
let failed = 0;
const ok = (c, m) => {
  console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`);
  if (!c) failed++;
};
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1300, height: 700 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('file://' + path.resolve(__dirname, '..', 'index.html'));
  await page.setInputFiles('.dropzone input[type=file]', path.join(__dirname, 'fixtures', 'sample_master.docx'));
  await page.waitForFunction(() => RG.app.S.master);
  await page.evaluate(() => RG.app.go('classify'));
  const wrapSel = 'main > div > .card:last-child .tbl-wrap';
  await page.$eval(wrapSel, (w) => (w.scrollTop = Math.floor(w.scrollHeight / 2)));
  await page.evaluate(() => window.scrollTo(0, 300));
  const before = await page.$eval(wrapSel, (w) => w.scrollTop);
  const winBefore = await page.evaluate(() => window.scrollY);
  ok(before > 100, `清單已往下捲（${before}px）`);
  // 改清單可視範圍中間那一列的分類（模擬使用者點選）
  const idx = await page.$eval(wrapSel, (w) => {
    const box = w.getBoundingClientRect();
    const sels = [...w.querySelectorAll('tbody tr select')];
    return sels.findIndex((s) => {
      const r = s.getBoundingClientRect();
      return r.top > box.top + box.height / 3 && r.bottom < box.bottom;
    });
  });
  const last = (await page.$$(`${wrapSel} tbody tr select`))[idx];
  await last.focus();
  await last.scrollIntoViewIfNeeded(); // Playwright 操作前本來就會捲到元素，先捲好再量
  const winBefore2 = await page.evaluate(() => window.scrollY);
  const fk = await last.getAttribute('data-fk');
  const cur = await last.inputValue();
  await last.selectOption(cur === 'T' ? 'M' : 'T');
  const after = await page.$eval(wrapSel, (w) => w.scrollTop);
  const winAfter = await page.evaluate(() => window.scrollY);
  ok(Math.abs(after - before) <= 2, `修改分類後清單位置不變（${before} → ${after}）`);
  ok(Math.abs(winAfter - winBefore2) <= 2, `修改分類後整頁位置不變（${winBefore2} → ${winAfter}）`);
  ok((await page.evaluate(() => document.activeElement && document.activeElement.getAttribute('data-fk'))) === fk, '焦點留在剛修改的欄位');
  // 改標籤
  const inp = (await page.$$(`${wrapSel} tbody tr input[type=text]`))[idx + 1];
  await inp.focus();
  await inp.fill('電梯');
  await inp.evaluate((el) => el.dispatchEvent(new Event('change', { bubbles: true })));
  const after2 = await page.$eval(wrapSel, (w) => w.scrollTop);
  ok(Math.abs(after2 - before) <= 2, `修改標籤後清單位置不變（${after2}）`);
  // 換篩選條件時回到清單頂端
  await page.selectOption('select[data-filter="cat"]', 'C');
  const after3 = await page.$eval(wrapSel, (w) => w.scrollTop);
  ok(after3 === 0, '換篩選條件時清單回到頂端');
  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
