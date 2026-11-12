/* 記憶庫管理頁的端對端測試：node tests/e2e_memory_page.js
 * 需要 tests/fixtures/sample_master.docx 與 python3 + openpyxl（修改匯出的 Excel） */
const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..');
const v1 = path.join(__dirname, 'fixtures', 'sample_master.docx');
const outDir = process.env.OUT_DIR || path.join(__dirname, 'out');
fs.mkdirSync(outDir, { recursive: true });
let failed = 0;
const ok = (c, m) => {
  console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`);
  if (!c) failed++;
};

(async () => {
  const browser = await chromium.launch();
  const page = await (await browser.newContext({ acceptDownloads: true })).newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('dialog', (d) => d.accept());
  await page.goto('file://' + path.join(root, 'index.html'));
  await page.setInputFiles('.dropzone input[type=file]', v1);
  await page.waitForFunction(() => RG.app.S.master);

  // 準備：全部確認，並刻意做出兩種有問題的記錄
  await page.evaluate(() => {
    const { S, remember } = RG.app;
    const by = (s) => S.master.blocks.find((b) => b.text.includes(s));
    S.overrides[by('本報告依據公共工程').id] = { cat: 'C', tags: [] }; // C 類沒有標籤
    S.overrides[by('本案為臺中市').id] = { tags: ['電梯'] }; // 有標籤但不是 C
    remember(S.master.blocks);
  });
  const total = await page.evaluate(() => RG.mem.stats(RG.app.S.memory).count);
  await page.evaluate(() => RG.app.go('memory'));
  await page.waitForSelector('text=完整性檢查');
  await page.screenshot({ path: path.join(outDir, 'memory-page.png'), fullPage: true });
  const kpiText = await page.textContent('.kpis');
  ok(kpiText.includes(String(total)), `概況顯示記憶筆數 ${total}`);
  const checks = await page.textContent('.checklist');
  ok(/C 類沒有標籤 1 筆/.test(checks), '完整性檢查：C 類沒有標籤 1 筆');
  ok(/有標籤但不是 C 1 筆/.test(checks), '完整性檢查：有標籤但不是 C 1 筆');
  ok(/目前母本「sample_master.docx」共 \d+ 個可記憶的段落／表格，其中 \d+ 個已在記憶庫（100%）/.test(await page.textContent('main')), '顯示目前母本的記憶涵蓋率 100%');

  const listRows = () => page.$$eval('main > div > .card:last-child table.data tbody tr', (t) => t.length);
  // 只看 C 類沒有標籤 → 直接改分類
  await page.click('li:has-text("C 類沒有標籤") button:has-text("只看這些")');
  ok((await listRows()) === 1, '「只看這些」篩選出 1 筆');
  await page.selectOption('main > div > .card:last-child table.data tbody tr select >> nth=0', 'T');
  const afterEdit = await page.evaluate(() => {
    const b = RG.app.S.master.blocks.find((x) => x.text.includes('本報告依據公共工程'));
    const e = RG.app.S.memory.entries[RG.mem.keyOf(b)];
    return { mem: e.cat, eff: RG.kb.effective(b, RG.app.S.overrides).cat };
  });
  ok(afterEdit.mem === 'T' && afterEdit.eff === 'T', '在記憶庫改分類，記憶與目前母本同步更新');

  // 有標籤但不是 C → 批次改為 C
  await page.click('button:has-text("清除篩選")');
  await page.selectOption('main > div > .card:nth-last-child(2) select >> nth=2', 'tagNotC');
  await page.click('button:has-text("選取全部篩選結果")');
  await page.selectOption('main > div > .card:nth-last-child(2) select >> nth=3', 'C');
  await page.click('button:has-text("套用")');
  const tagNotC = await page.evaluate(() => Object.values(RG.app.S.memory.entries).filter((e) => e.cat !== 'C' && e.tags.length).length);
  ok(tagNotC === 0, '批次改分類後沒有「有標籤但不是 C」');

  // 刪除單筆
  await page.click('button:has-text("清除篩選")');
  await page.fill('input[type=search]', '本案為臺中市');
  await page.press('input[type=search]', 'Enter');
  await page.dispatchEvent('input[type=search]', 'change');
  ok((await listRows()) === 1, '搜尋找到 1 筆');
  await page.click('main > div > .card:last-child table.data tbody tr button:has-text("刪除")');
  const del = await page.evaluate(() => {
    const b = RG.app.S.master.blocks.find((x) => x.text.includes('本案為臺中市'));
    return { count: RG.mem.stats(RG.app.S.memory).count, mem: b.mem.state, ov: !!RG.app.S.overrides[b.id] };
  });
  ok(del.count === total - 1 && del.mem === 'new' && !del.ov, '刪除單筆後，目前母本的對應段落改回自動判斷');

  // 衝突標為已解決
  await page.evaluate(() => { const e = Object.values(RG.app.S.memory.entries)[0]; e.conflict = true; RG.app.render(); });
  await page.click('button:has-text("清除篩選")');
  await page.click('li:has-text("衝突") button:has-text("只看這些")');
  await page.click('button:has-text("選取全部篩選結果")');
  await page.click('button:has-text("衝突標為已解決")');
  ok((await page.evaluate(() => RG.mem.stats(RG.app.S.memory).conflicts)) === 0, '衝突標為已解決');

  // 匯出 Excel → 用 Python 修改 → 匯入
  await page.click('button:has-text("清除篩選")');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('button:has-text("匯出 Excel")')]);
  const xlsx = path.join(outDir, 'memory_export.xlsx');
  await dl.saveAs(xlsx);
  const edited = path.join(outDir, 'memory_edited.xlsx');
  const py = `
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1]); ws = wb['記憶庫']
hdr = [c.value for c in ws[1]]; col = {h: i + 1 for i, h in enumerate(hdr)}
rows = list(range(2, ws.max_row + 1))
target = None
for r in rows:
    if '營造安全衛生設施標準' in str(ws.cell(r, col['內容（前 120 字）']).value):
        target = r
ws.cell(target, col['分類']).value = 'M'
ws.cell(target, col['工項規劃']).value = '刪除'
others = [r for r in rows if r != target]
ws.cell(others[0], col['刪除']).value = '刪除'
ws.cell(others[1], col['分類']).value = 'X'
wb.save(sys.argv[2])
`;
  execFileSync('python3', ['-c', py, xlsx, edited]);
  const before = await page.evaluate(() => RG.mem.stats(RG.app.S.memory).count);
  await page.setInputFiles('label:has-text("匯入 Excel 修改") input', edited);
  await page.waitForSelector('.toast:has-text("已匯入")');
  const toastText = await page.textContent('.toast:has-text("已匯入")');
  console.log(toastText);
  const imp = await page.evaluate(() => {
    const b = RG.app.S.master.blocks.find((x) => x.text.includes('營造安全衛生設施標準'));
    const e = RG.app.S.memory.entries[RG.mem.keyOf(b)];
    return { cat: e.cat, plan: e.plan, eff: RG.kb.effective(b, RG.app.S.overrides).cat, po: RG.app.S.planOverrides[b.id], count: RG.mem.stats(RG.app.S.memory).count };
  });
  ok(imp.cat === 'M' && imp.plan === 'drop', 'Excel 修改的分類與工項規劃已匯入');
  ok(imp.eff === 'M' && imp.po === 'drop', 'Excel 修改已套用到目前母本');
  ok(imp.count === before - 1, 'Excel 標「刪除」的列已刪除');
  ok(/1 筆分類或工項規劃填寫錯誤已略過/.test(toastText), '填錯的分類（X）會略過並提示');

  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
