/* 標籤防呆、鍵盤快速分類、表頭篩選與批次處理：node tests/e2e_classify_tools.js */
const path = require('path');
const { chromium } = require('playwright');
let failed = 0;
const ok = (c, m) => {
  console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`);
  if (!c) failed++;
};
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('dialog', (d) => d.accept());
  await page.goto('file://' + path.resolve(__dirname, '..', 'index.html'));
  await page.setInputFiles('.dropzone input[type=file]', path.join(__dirname, 'fixtures', 'sample_master.docx'));
  await page.waitForFunction(() => RG.app.S.master);
  await page.evaluate(() => RG.app.go('classify'));
  const eff = (id) => page.evaluate((id) => { const b = RG.app.S.master.blocks.find((x) => x.id === id); return RG.kb.effective(b, RG.app.S.overrides); }, id);
  const rowsCount = () => page.$$eval('main > div > .card:last-child tbody tr[data-row]', (t) => t.length);

  // ---------- 1. 表頭篩選 ----------
  await page.selectOption('select[data-filter="kind"]', 'figure');
  const nFig = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => b.kind === 'figure').length);
  ok((await rowsCount()) === nFig, `表頭「類型」篩選：只顯示圖 ${nFig} 筆`);
  ok(await page.$eval('select[data-filter="kind"]', (s) => s.classList.contains('filter-on')), '使用中的篩選欄以藍底標示');
  await page.selectOption('select[data-filter="kind"]', 'imgtable');
  ok((await rowsCount()) === 2, '「表格截圖」篩選出 2 筆');
  await page.click('button:has-text("清除所有篩選")');
  await page.selectOption('select[data-filter="autoCat"]', 'C');
  const autoC = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => !['empty', 'other'].includes(b.kind) && b.auto.cat === 'C').length);
  ok((await rowsCount()) === autoC, `表頭「自動判斷」篩選：自動判為 C 的 ${autoC} 筆`);
  await page.click('button:has-text("清除所有篩選")');
  await page.fill('input[data-filter="q"]', '表4-1');
  await page.press('input[data-filter="q"]', 'Enter');
  await page.dispatchEvent('input[data-filter="q"]', 'change');
  ok((await rowsCount()) >= 1, '表頭「內容」篩選：搜尋表號');
  await page.click('button:has-text("清除所有篩選")');

  // ---------- 2. 批次處理篩選結果 ----------
  await page.selectOption('select[data-filter="kind"]', 'figure');
  await page.selectOption('select[data-bulk="cat"]', 'T');
  await page.click('.bulkbar button:has-text("套用")');
  const figCats = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => b.kind === 'figure').map((b) => RG.kb.effective(b, RG.app.S.overrides).cat));
  ok(figCats.every((c) => c === 'T'), '批次：篩選出的圖全部設為 T');
  await page.selectOption('select[data-bulk="cat"]', 'C');
  await page.click('.bulkbar button:has-text("套用")');
  await page.fill('input[data-bulk="tag"]', '電梯');
  await page.click('.bulkbar button:has-text("加入標籤")');
  let figTags = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => b.kind === 'figure').map((b) => RG.kb.effective(b, RG.app.S.overrides).tags.join()));
  ok(figTags.every((t) => t.includes('電梯')), '批次：加入標籤');
  await page.fill('input[data-bulk="tag"]', '電梯');
  await page.click('.bulkbar button:has-text("移除標籤")');
  figTags = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => b.kind === 'figure').map((b) => RG.kb.effective(b, RG.app.S.overrides).tags.join()));
  ok(figTags.every((t) => !t.includes('電梯')), '批次：移除標籤');
  await page.click('.bulkbar button:has-text("全部確認")');
  const figMem = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => b.kind === 'figure' && RG.mem.memorable(b)).every((b) => b.mem.state === 'confirmed'));
  ok(figMem, '批次：全部確認並記入記憶');
  await page.click('button:has-text("清除所有篩選")');

  // ---------- 3. 標籤防呆 ----------
  const t31 = await page.evaluate(() => RG.app.S.master.blocks.find((b) => b.kind === 'table' && b.capNo === '3-1').id);
  await page.fill('input[data-filter="q"]', '表3-1');
  await page.dispatchEvent('input[data-filter="q"]', 'change');
  const tagSel = `input[data-fk="${t31}:tags"]`;
  await page.fill(tagSel, '擋土柱斜撐');
  await page.dispatchEvent(tagSel, 'change');
  ok(await page.$eval(tagSel, (el) => el.classList.contains('bad-input')), '打錯字的標籤：輸入欄標紅');
  const warn = await page.textContent(`tr[data-row="${t31}"] .tag-warn`);
  ok(warn.includes('不是已知標籤'), '打錯字的標籤：顯示警告');
  ok(/有 \d+ 筆使用了工具不認識的標籤/.test(await page.textContent('main')), '上方顯示不認識標籤的筆數');
  await page.click(`tr[data-row="${t31}"] button:has-text("改為「擋土柱+斜撐」")`);
  ok((await eff(t31)).tags.join() === '擋土柱+斜撐', '一鍵改為建議的正確標籤');
  await page.fill(tagSel, '鷹架');
  await page.dispatchEvent(tagSel, 'change');
  await page.click(`tr[data-row="${t31}"] button:has-text("登錄為自訂標籤")`);
  ok(!(await page.$(`tr[data-row="${t31}"] .tag-warn`)), '登錄為自訂標籤後不再警告');
  ok((await page.evaluate(() => RG.app.S.rules.customTags)).includes('鷹架'), '自訂標籤已存入規則');
  await page.evaluate(() => RG.app.go('project'));
  ok(await page.$('button.chip:has-text("鷹架")'), '專案資料的條件標籤可以啟用自訂標籤');
  await page.evaluate(() => RG.app.go('classify'));
  await page.click('button:has-text("清除所有篩選")');

  // ---------- 4. 鍵盤快速分類 ----------
  const ids = await page.evaluate(() => RG.app.S.ui.clsPageIds);
  await page.click(`tr[data-row="${ids[2]}"] td:nth-child(2)`);
  ok((await page.evaluate(() => RG.app.S.ui.cls.cur)) === ids[2], '點一列成為目前列');
  await page.keyboard.press('ArrowDown');
  ok((await page.evaluate(() => RG.app.S.ui.cls.cur)) === ids[3], '↓ 移到下一列');
  ok(await page.$(`tr.cur[data-row="${ids[3]}"]`), '目前列有標示');
  await page.keyboard.press('m');
  ok((await eff(ids[3])).cat === 'M', '按 M 設為混合');
  await page.keyboard.press('Enter');
  const st = await page.evaluate((id) => ({ cur: RG.app.S.ui.cls.cur, mem: RG.app.S.master.blocks.find((b) => b.id === id).mem.state }), ids[3]);
  ok(st.mem === 'confirmed' && st.cur === ids[4], 'Enter 確認並跳到下一列');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowUp');
  ok((await page.evaluate(() => RG.app.S.ui.cls.cur)) === ids[2], '↑ 移到上一列');
  // 找一列沒有標籤的，按 C 會進入標籤欄
  const noTagIdx = await page.evaluate(() => { const S = RG.app.S; return S.ui.clsPageIds.findIndex((id, i) => i > 4 && !RG.kb.effective(S.master.blocks.find((b) => b.id === id), S.overrides).tags.length); });
  const target = ids[noTagIdx];
  await page.click(`tr[data-row="${target}"] td:nth-child(2)`);
  await page.keyboard.press('c');
  ok((await eff(target)).cat === 'C', '按 C 設為條件式');
  ok((await page.evaluate(() => document.activeElement.getAttribute('data-fk'))) === `${target}:tags`, 'C 沒有標籤時自動進入標籤欄');
  await page.keyboard.type('電梯');
  await page.keyboard.press('Enter');
  ok((await eff(target)).tags.join() === '電梯', '在標籤欄輸入後按 Enter 儲存');
  ok((await page.evaluate(() => document.activeElement.tagName)) !== 'INPUT', 'Enter 後離開輸入欄，可繼續用鍵盤');
  await page.keyboard.press('ArrowDown');
  ok((await page.evaluate(() => RG.app.S.ui.cls.cur)) === ids[noTagIdx + 1], '離開標籤欄後 ↓ 仍可移動');
  await page.keyboard.press('e');
  ok((await page.evaluate(() => document.activeElement.getAttribute('data-fk'))) === `${ids[noTagIdx + 1]}:tags`, '按 E 編輯標籤');
  await page.keyboard.press('Escape');
  ok((await page.evaluate(() => document.activeElement.tagName)) !== 'INPUT', 'Esc 離開輸入欄');
  await page.screenshot({ path: path.join(process.env.OUT_DIR || path.join(__dirname, 'out'), 'classify-tools.png') }).catch(() => {});

  // ---------- 5. 工項規劃的表頭篩選與批次 ----------
  await page.evaluate(() => RG.app.go('plan'));
  await page.selectOption('select[data-filter="kind"]', 'figure');
  const planFig = await page.$$eval('main > div > .card:last-child tbody tr', (t) => t.length);
  ok(planFig === nFig, `工項規劃表頭篩選：圖 ${planFig} 筆`);
  await page.click('.bulkbar button:has-text("全部刪除")');
  const drops = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => b.kind === 'figure').every((b) => RG.app.S.planOverrides[b.id] === 'drop'));
  ok(drops, '工項規劃批次：篩選結果全部刪除');
  await page.selectOption('select[data-filter="kind"]', '');
  await page.selectOption('select[data-filter="decision"]', 'drop');
  ok((await page.$$eval('main > div > .card:last-child tbody tr', (t) => t.length)) === nFig, '「決定」篩選：指定刪除的項目');
  await page.click('.bulkbar button:has-text("全部改回自動")');
  ok((await page.evaluate(() => Object.keys(RG.app.S.planOverrides).length)) === 0, '工項規劃批次：全部改回自動');

  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
