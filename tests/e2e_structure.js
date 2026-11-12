/* 依架構由大到小的分類、各章工項連動、工項前後對照、舊記憶移轉：node tests/e2e_structure.js */
const path = require('path');
const { chromium } = require('playwright');
let failed = 0;
const ok = (c, m) => {
  console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`);
  if (!c) failed++;
};
(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  const url = 'file://' + path.resolve(__dirname, '..', 'index.html');

  // ---------- 1. 舊記憶移轉 ----------
  await page.goto(url);
  await page.evaluate(() => {
    localStorage.setItem('rg-memory-v1', JSON.stringify({ version: 1, threshold: 0.9, entries: { 'p:abc:3': { kind: 'p', len: 3, nh: '舊段落', pv: '舊段落', cat: 'T', tags: [], plan: '', masters: ['old.docx:1'], conflict: false, updated: 1 } } }));
    const st = JSON.parse(localStorage.getItem('rg-report-tool-v1') || '{}');
    st.overrides = { b10: { cat: 'T' } };
    localStorage.setItem('rg-report-tool-v1', JSON.stringify(st));
  });
  await page.reload();
  const mig = await page.evaluate(() => ({ n: Object.keys(RG.app.S.memory.entries).length, v: RG.app.S.memory.classifyVersion, ov: Object.keys(RG.app.S.overrides).length }));
  ok(mig.n === 0 && mig.v === 2 && mig.ov === 0, '舊原則的記憶庫與分類確認已清空');
  await page.setInputFiles('.dropzone input[type=file]', path.join(__dirname, 'fixtures', 'sample_master.docx'));
  await page.waitForFunction(() => RG.app.S.master);
  await page.evaluate(() => RG.app.go('classify'));
  ok(/分類原則已更新/.test(await page.textContent('main')), '分類頁顯示原則更新提示');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('button:has-text("下載舊資料備份")')]);
  const fs = require('fs');
  const tmp = path.join(process.env.OUT_DIR || path.join(__dirname, 'out'), 'legacy.json');
  fs.mkdirSync(path.dirname(tmp), { recursive: true });
  await dl.saveAs(tmp);
  const legacy = JSON.parse(fs.readFileSync(tmp, 'utf8'));
  ok(legacy.memory.entries['p:abc:3'] && legacy.overrides.b10, '舊資料備份含原本的記憶與分類確認');
  await page.click('button:has-text("關閉提示")');
  ok(!/分類原則已更新/.test(await page.textContent('main')), '關閉提示後不再顯示');

  // ---------- 1b. 前置頁：內容 F、頁面標題 T ----------
  const front = await page.evaluate(() => Object.fromEntries(RG.app.S.master.blocks.filter((b) => b.chapter === '前置' && b.kind === 'p').map((b) => [b.text.trim(), b.auto.cat])));
  console.log(JSON.stringify(front));
  ok(front['設計階段施工風險評估報告'] === 'T', '前置頁的頁面標題（…風險評估報告）沿用 T');
  ok(front['中華民國115年6月'] === 'F' && front['代辨機關：臺中市政府'] === 'F', '前置頁的日期、機關等內容為 F');
  // 已存的章節原則（更新前）會自動補上「前置」
  const upgraded = await page.evaluate(() => {
    const st = JSON.parse(localStorage.getItem('rg-report-tool-v1'));
    st.rules.sectionPolicy = st.rules.sectionPolicy.filter((r) => !r.chapter.includes('前置'));
    delete st.rules.sectionPolicyRev;
    localStorage.setItem('rg-report-tool-v1', JSON.stringify(st));
    return true;
  });
  await page.reload();
  const sp = await page.evaluate(() => RG.app.S.rules.sectionPolicy.map((r) => r.chapter));
  ok(upgraded && sp[0] === '前置', '已存的章節原則會自動補上「前置」');
  await page.setInputFiles('.dropzone input[type=file]', path.join(__dirname, 'fixtures', 'sample_master.docx'));
  await page.waitForFunction(() => RG.app.S.master);

  // ---------- 2. 各章工項連動、標題隨內容刪除 ----------
  const plan = (po) => page.evaluate((po) => {
    const S = RG.app.S;
    S.planOverrides = po;
    const res = RG.kb.plan(S.master.blocks, S.overrides, new Set(['擋土柱+斜撐', '電梯', '施工']), S.planOverrides);
    const by = (f) => S.master.blocks.find(f);
    const r = (b) => res[b.id];
    return {
      t41: r(by((b) => b.kind === 'table' && b.capNo === '4-1')),
      c41: r(by((b) => b.kind === 'caption' && b.capNo === '4-1')),
      gB: r(by((b) => b.kind === 'rows' && b.title.includes('擋土'))),
      gA: r(by((b) => b.kind === 'rows' && b.title.includes('假設'))),
      h1: r(by((b) => b.kind === 'heading' && b.text.includes('（一）擋土設施'))),
      h2: r(by((b) => b.kind === 'heading' && b.text.includes('（二）電梯組立'))),
      t32: r(by((b) => b.kind === 'table' && b.capNo === '3-2')),
    };
  }, po);
  const t31 = await page.evaluate(() => RG.app.S.master.blocks.find((b) => b.kind === 'table' && b.capNo === '3-1').id);
  let r = await plan({});
  ok(r.t41.include && r.gB.include && r.h1.include, '條件符合時，擋土相關的各章表格與標題都保留');
  r = await plan({ [t31]: 'drop' });
  console.log(JSON.stringify({ t41: r.t41, gB: r.gB, h1: r.h1 }));
  ok(!r.t41.include && /各章工項連動/.test(r.t41.reason), '第三章刪除「擋土設施」→ 第四章「擋土支撐工程作業拆解表」連動刪除');
  ok(!r.c41.include, '連動刪除的表格，表名一起刪除');
  ok(!r.gB.include && /各章工項連動/.test(r.gB.reason), '→ 第四章風險評估表的「擋土支撐工程」子表格連動刪除');
  ok(r.gA.include, '其他工項（假設工程）不受影響');
  ok(!r.h1.include && /標題底下的內容全部刪除/.test(r.h1.reason), '「（一）擋土設施」底下內容全刪 → 標題也刪除');
  const t41id = await page.evaluate(() => RG.app.S.master.blocks.find((b) => b.kind === 'table' && b.capNo === '4-1').id);
  r = await plan({ [t31]: 'drop', [t41id]: 'keep' });
  ok(r.t41.include, '使用者指定保留的不會被連動刪除');
  // 沒有電梯條件 → 表3-2 刪除、（二）電梯組立 標題也刪除
  const r2 = await page.evaluate(() => {
    const S = RG.app.S;
    const res = RG.kb.plan(S.master.blocks, S.overrides, new Set(['擋土柱+斜撐', '施工']), {});
    const by = (f) => S.master.blocks.find(f);
    return { t32: res[by((b) => b.kind === 'table' && b.capNo === '3-2').id], h2: res[by((b) => b.kind === 'heading' && b.text.includes('（二）電梯組立')).id] };
  });
  ok(!r2.t32.include && !r2.h2.include, '沒有電梯條件 → 表3-2 與「（二）電梯組立」標題一起刪除');

  // ---------- 3. 工項前後對照 ----------
  await page.evaluate(() => RG.app.go('check'));
  await page.click('button:has-text("檢查母本")');
  await page.waitForSelector('text=檢查結果');
  const iss = await page.evaluate(() => RG.app.S.lastCheck.issues.filter((i) => i.item === '工項前後不一致').map((i) => i.textA + '｜' + i.desc));
  console.log(iss.join('\n'));
  ok(iss.some((x) => x.startsWith('骨櫃工程') && x.includes('作業拆解表')), '風險評估表有「骨櫃工程」但沒有作業拆解表 → 列出');
  ok(iss.some((x) => x.startsWith('電梯組立')), '第三章評選「電梯組立」在第四章找不到 → 列出');
  ok(!iss.some((x) => x.startsWith('擋土') || x.startsWith('假設工程')), '各章都有的工項（擋土、假設工程）不會誤報');

  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
