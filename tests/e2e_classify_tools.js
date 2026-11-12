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

  // ---------- 4b. 複選 ----------
  await page.evaluate(() => { RG.app.S.ui.cls.sel = {}; RG.app.render(); });
  const pid = await page.evaluate(() => RG.app.S.ui.clsPageIds);
  const selIds = () => page.evaluate(() => Object.keys(RG.app.S.ui.cls.sel || {}).filter((k) => RG.app.S.ui.cls.sel[k]));
  const cell = (id) => `tr[data-row="${id}"] td.selcell`;
  await page.click(cell(pid[10]));
  await page.click(cell(pid[13]), { modifiers: ['Shift'] });
  let sel = await selIds();
  ok(sel.length === 4 && [10, 11, 12, 13].every((i) => sel.includes(pid[i])), '點選＋Shift 點選：選取一段範圍（4 筆）');
  ok((await page.textContent('[data-bulk-label]')).includes('已選取的 4 筆'), '批次列改為「對已選取的 4 筆」');
  ok(await page.$eval(`tr[data-row="${pid[11]}"]`, (tr) => tr.classList.contains('selected') && tr.querySelector('input.selbox').checked), '已選取的列有勾選與底色');
  const beforeOthers = await page.evaluate((ids) => ids.map((id) => RG.kb.effective(RG.app.S.master.blocks.find((b) => b.id === id), RG.app.S.overrides).cat), [pid[9], pid[14]]);
  await page.selectOption('select[data-bulk="cat"]', 'M');
  await page.click('.bulkbar button:has-text("套用")');
  const cats = await page.evaluate((ids) => ids.map((id) => RG.kb.effective(RG.app.S.master.blocks.find((b) => b.id === id), RG.app.S.overrides).cat), [pid[10], pid[11], pid[12], pid[13], pid[9], pid[14]]);
  ok(cats.slice(0, 4).every((c) => c === 'M'), '批次分類只套用到已選取的 4 筆');
  ok(cats[4] === beforeOthers[0] && cats[5] === beforeOthers[1], '未選取的列不受影響');
  ok((await selIds()).length === 4, '套用後勾選保留，可以繼續做其他批次處理');
  await page.click('button:has-text("取消選取")');
  ok((await selIds()).length === 0, '取消選取');
  // 拖曳
  const box = async (id) => (await page.$(cell(id))).boundingBox();
  await page.$eval(cell(pid[22]), (el) => el.scrollIntoView({ block: 'center' })); // 讓 20～24 列都在清單可視範圍內
  const b1 = await box(pid[20]);
  await page.mouse.move(b1.x + b1.width / 2, b1.y + b1.height / 2);
  await page.mouse.down();
  for (const i of [21, 22, 23, 24]) {
    const bb = await box(pid[i]);
    await page.mouse.move(bb.x + bb.width / 2, bb.y + bb.height / 2, { steps: 3 });
  }
  await page.mouse.up();
  sel = await selIds();
  ok(sel.length === 5 && [20, 21, 22, 23, 24].every((i) => sel.includes(pid[i])), '按住拖曳連續選取（5 筆）');
  // 表頭全選
  await page.click('input[data-sel-all]');
  ok((await selIds()).length === pid.length, '表頭勾選框選取本頁全部');
  await page.click('button:has-text("取消選取")');
  // 鍵盤：空白鍵、Shift+↓
  await page.click(`tr[data-row="${pid[30]}"] td:nth-child(2)`);
  await page.keyboard.press(' ');
  ok((await selIds()).join() === pid[30], '空白鍵勾選目前列');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  sel = await selIds();
  ok(sel.length === 3 && [30, 31, 32].every((i) => sel.includes(pid[i])), 'Shift＋↓ 一邊移動一邊勾選');
  await page.click('button:has-text("選取全部篩選結果")');
  ok((await selIds()).length === (await page.evaluate(() => RG.app.S.ui.clsListCount)), '選取全部篩選結果');
  await page.click('button:has-text("取消選取")');
  // 說明區塊
  const legend = await page.textContent('details.card');
  ok(legend.includes('鍵盤快速分類') && legend.includes('Shift') && legend.includes('複選與批次處理'), '說明區塊含鍵盤快速分類與複選說明');

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

  // ---------- 6. 工項規劃的複選 ----------
  await page.click('button:has-text("清除所有篩選")');
  const ppid = await page.evaluate(() => RG.app.S.ui.planPageIds);
  ok(ppid.length >= 6, `工項規劃清單有 ${ppid.length} 筆可測試`);
  const psel = () => page.evaluate(() => Object.keys(RG.app.S.ui.plan.sel || {}).filter((k) => RG.app.S.ui.plan.sel[k]));
  await page.click(cell(ppid[1]));
  await page.click(cell(ppid[3]), { modifiers: ['Shift'] });
  let ps = await psel();
  ok(ps.length === 3 && [1, 2, 3].every((i) => ps.includes(ppid[i])), '工項規劃：Shift＋點選選取範圍（3 筆）');
  ok((await page.textContent('[data-bulk-label]')).includes('已選取的 3 筆'), '工項規劃：批次列改為「對已選取的 3 筆」');
  await page.click('.bulkbar button:has-text("全部刪除")');
  const po = await page.evaluate(() => RG.app.S.planOverrides);
  ok(Object.keys(po).length === 3 && [1, 2, 3].every((i) => po[ppid[i]] === 'drop'), '工項規劃：只有已選取的 3 筆設為刪除');
  ok((await psel()).length === 3, '工項規劃：套用後勾選保留');
  await page.click('button:has-text("取消選取")');
  ok((await psel()).length === 0, '工項規劃：取消選取');
  await page.$eval(cell(ppid[Math.min(4, ppid.length - 1)]), (el) => el.scrollIntoView({ block: 'center' }));
  const pb = async (id) => (await page.$(cell(id))).boundingBox();
  const p0 = await pb(ppid[3]);
  await page.mouse.move(p0.x + p0.width / 2, p0.y + p0.height / 2);
  await page.mouse.down();
  for (const i of [4, 5]) {
    const bb = await pb(ppid[i]);
    await page.mouse.move(bb.x + bb.width / 2, bb.y + bb.height / 2, { steps: 3 });
  }
  await page.mouse.up();
  ps = await psel();
  ok(ps.length === 3 && [3, 4, 5].every((i) => ps.includes(ppid[i])), '工項規劃：按住拖曳連續選取（3 筆）');
  await page.click('.bulkbar button:has-text("全部保留")');
  const po2 = await page.evaluate(() => RG.app.S.planOverrides);
  ok([3, 4, 5].every((i) => po2[ppid[i]] === 'keep'), '工項規劃：已選取的設為保留');
  await page.click('button:has-text("取消選取")');
  await page.click('input[data-sel-all]');
  ok((await psel()).length === (await page.evaluate(() => RG.app.S.ui.planPageIds.length)), '工項規劃：表頭勾選框選取本頁全部');
  await page.click('.bulkbar button:has-text("全部改回自動")');
  ok((await page.evaluate(() => Object.keys(RG.app.S.planOverrides).length)) === 0, '工項規劃：已選取的全部改回自動');
  // 分類頁與工項規劃頁的勾選互不影響
  await page.evaluate(() => RG.app.go('classify'));
  ok((await page.evaluate(() => Object.keys(RG.app.S.ui.cls.sel || {}).filter((k) => RG.app.S.ui.cls.sel[k]).length)) === 0, '分類頁的勾選不受工項規劃頁影響');

  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
