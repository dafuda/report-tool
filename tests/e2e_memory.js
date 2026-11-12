/* 依文字內容記憶的端對端測試。
 * 用法：
 *   python tests/make_sample_docx.py tests/fixtures/sample_master.docx
 *   python tests/make_sample_docx.py tests/fixtures/sample_master_v2.docx --v2
 *   node tests/e2e_memory.js */
const path = require('path');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..');
const v1 = path.join(__dirname, 'fixtures', 'sample_master.docx');
const v2 = path.join(__dirname, 'fixtures', 'sample_master_v2.docx');
let failed = 0;
const ok = (cond, msg) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${msg}`);
  if (!cond) failed++;
};

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  const url = 'file://' + path.join(root, 'index.html');
  await page.goto(url);
  const load = async (file) => {
    await page.evaluate(() => RG.app.go('master'));
    await page.setInputFiles('main input[type=file][accept=".docx"] >> nth=0', file);
    await page.waitForFunction((n) => RG.app.S.master && RG.app.S.master.name === n, path.basename(file));
  };
  const find = (s) => page.evaluate((s) => {
    const b = RG.app.S.master.blocks.find((x) => x.text.includes(s));
    const e = RG.kb.effective(b, RG.app.S.overrides);
    return { id: b.id, mem: b.mem && b.mem.state, sim: b.mem && b.mem.sim, cat: e.cat, tags: e.tags, plan: RG.app.S.planOverrides[b.id] || '' };
  }, s);

  // 1. 第一次讀入：記憶庫是空的，全部是新段落
  await load(v1);
  let r = await page.evaluate(() => RG.app.S.memResult);
  ok(r.exact === 0 && r.fuzzy === 0 && r.fresh > 0, `第一次讀入全部為新段落（${r.fresh} 筆）`);

  // 2. 修改幾筆，然後確認全部
  await page.evaluate(() => {
    const { S, remember } = RG.app;
    const B = S.master.blocks;
    const by = (s) => B.find((x) => x.text.includes(s));
    S.overrides[by('本案為臺中市').id] = { cat: 'T' }; // 刻意改成 T，用來驗證依文字對應
    S.overrides[by('營造安全衛生設施標準').id] = { cat: 'M' };
    S.overrides[B.find((x) => x.kind === 'table' && x.capNo === '3-2').id] = { tags: ['電梯', 'RC'] };
    const g = B.find((x) => x.kind === 'rows' && x.title.includes('骨櫃'));
    S.planOverrides[g.id] = 'keep';
    remember(B);
  });
  const memCount = await page.evaluate(() => RG.mem.stats(RG.app.S.memory).count);
  ok(memCount > 10, `確認後記憶庫有 ${memCount} 筆`);

  // 3. 換成改版母本（開頭多一段、一段略改、一段刪除）
  await load(v2);
  r = await page.evaluate(() => { const r = RG.app.S.memResult; return { exact: r.exact, fuzzy: r.fuzzy, fresh: r.fresh, missing: r.missing.map((e) => e.pv) }; });
  console.log('v2', JSON.stringify(r));
  const intro = await find('本案為臺中市');
  ok(intro.mem === 'exact' && intro.cat === 'T', `位置改變的段落依文字沿用記憶（${intro.id}，分類 ${intro.cat}）`);
  const added = await find('殯葬設施整體規劃');
  ok(added.mem === 'new', `新增的段落標為「新段落」（${added.id}，佔用了舊編號也不會套錯）`);
  ok(added.cat !== 'T', '新段落不會誤用舊編號的分類');
  const law = await find('營造安全衛生設施標準');
  ok(law.mem === 'fuzzy' && law.cat === 'M', `略改的段落標為「推定 ${Math.round(law.sim * 100)}%」並沿用分類 M`);
  const t32 = await page.evaluate(() => { const b = RG.app.S.master.blocks.find((x) => x.kind === 'table' && x.capNo === '3-2'); return { mem: b.mem.state, tags: RG.kb.effective(b, RG.app.S.overrides).tags }; });
  ok(t32.mem === 'exact' && t32.tags.join('、') === '電梯、RC', '表格的確認標籤沿用');
  const grp = await page.evaluate(() => { const b = RG.app.S.master.blocks.find((x) => x.kind === 'rows' && x.title.includes('骨櫃')); return RG.app.S.planOverrides[b.id] || ''; });
  ok(grp === 'keep', '工項規劃的使用者指定沿用');
  ok(r.missing.some((t) => t.includes('國豐街253巷')), '被刪除的段落列在「前一份有、這份沒有」');

  // 4. 只看需確認的 → 確認 → 重新整理後同一份母本，全部沿用
  await page.evaluate(() => RG.app.go('classify'));
  await page.click('button:has-text("只看需確認的")');
  const todoRows = await page.$$eval('main > div > .card:last-child table.data tbody tr', (trs) => trs.length);
  ok(todoRows === r.fuzzy + r.fresh, `「只看需確認的」篩選出 ${todoRows} 筆`);
  page.once('dialog', (d) => d.accept());
  await page.click('button:has-text("確認目前篩選的")');
  await page.screenshot({ path: path.join(process.env.OUT_DIR || path.join(__dirname, 'out'), 'memory-classify.png'), fullPage: false }).catch(() => {});
  await page.reload();
  await load(v2);
  r = await page.evaluate(() => RG.app.S.memResult);
  ok(r.fuzzy === 0 && r.fresh === 0, `重新開啟後同一份母本全部沿用（${r.exact} 筆）`);
  const law2 = await find('營造安全衛生設施標準');
  ok(law2.cat === 'M', '確認過的推定結果已記住');

  // 5. 重新整理後讀回第一版：仍全部依文字對應
  await page.reload();
  await load(v1);
  r = await page.evaluate(() => RG.app.S.memResult);
  const intro1 = await find('本案為臺中市');
  ok(intro1.cat === 'T' && r.fresh === 0, `切回第一版仍正確沿用（推定 ${r.fuzzy}、新段落 ${r.fresh}）`);

  // 6. 關鍵字建議：在規則設定頁分析、加入，已確認段落不受影響
  const before = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => RG.mem.memorable(b)).map((b) => b.id + ':' + JSON.stringify(RG.kb.effective(b, RG.app.S.overrides))));
  await page.evaluate(() => RG.app.go('rules'));
  await page.click('button:has-text("分析記憶庫")');
  const sugg = await page.evaluate(() => RG.app.S.kwSuggest.map((r) => ({ tag: r.tag, add: r.add.map((a) => a.term), remove: r.remove.map((m) => m.kw) })));
  console.log('suggest', JSON.stringify(sugg));
  ok(Array.isArray(sugg), '規則設定頁可以分析記憶庫');
  await page.screenshot({ path: path.join(process.env.OUT_DIR || path.join(__dirname, 'out'), 'keyword-suggest.png'), fullPage: false }).catch(() => {});
  const addBtn = await page.$('button.primary:has-text("加入")');
  if (addBtn) {
    const kwBefore = await page.evaluate(() => JSON.stringify(RG.app.S.rules.tagKeywords));
    await addBtn.click();
    const kwAfter = await page.evaluate(() => JSON.stringify(RG.app.S.rules.tagKeywords));
    ok(kwBefore !== kwAfter, '按「加入」後關鍵字已更新');
    await page.waitForSelector('.badge.keep:has-text("已加入")');
  }
  // 讓所有段落的自動判斷大改：把「電梯」關鍵字改成會抓到一般字的詞，再重新初判
  await page.evaluate(() => { RG.app.S.rules.tagKeywords['電梯'] = ['作業', '設置']; RG.app.go('classify'); });
  await page.click('button:has-text("重新初判")');
  const after = await page.evaluate(() => RG.app.S.master.blocks.filter((b) => RG.mem.memorable(b)).map((b) => b.id + ':' + JSON.stringify(RG.kb.effective(b, RG.app.S.overrides))));
  ok(JSON.stringify(before) === JSON.stringify(after), '關鍵字改變後重新初判，已確認的段落維持原確認結果');

  // 7. 清除記憶庫
  page.once('dialog', (d) => d.accept());
  await page.evaluate(() => RG.app.go('classify'));
  await page.click('button:has-text("清除記憶庫")');
  const cnt = await page.evaluate(() => RG.mem.stats(RG.app.S.memory).count);
  ok(cnt === 0, '清除記憶庫');

  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
