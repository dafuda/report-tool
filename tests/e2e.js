/* 端對端測試：用 Playwright 開啟 index.html，跑完整流程。
 * 用法：python tests/make_sample_docx.py tests/fixtures/sample_master.docx && node tests/e2e.js */
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..');
const sample = path.join(__dirname, 'fixtures', 'sample_master.docx');
const outDir = process.env.OUT_DIR || path.join(__dirname, 'out');
fs.mkdirSync(outDir, { recursive: true });

let failed = 0;
const ok = (cond, msg) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${msg}`);
  if (!cond) failed++;
};

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ acceptDownloads: true, viewport: { width: 1400, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('file://' + path.join(root, 'index.html'));

  // 1. 母本
  await page.setInputFiles('.drop input[type=file]', sample);
  await page.waitForSelector('.kpis');
  const rep = await page.evaluate(() => {
    const r = RG.app.S.master.report;
    return { ch: r.chapters.length, tables: r.tableCount, caps: r.tableCaptionCount, groups: r.groupCount, figs: r.figureCaptionCount, sect: r.sectionBreaks, seq: r.fields.SEQ, noCap: r.captionsWithoutTable };
  });
  console.log('structure', JSON.stringify(rep));
  ok(rep.ch === 6, '辨識出 6 章');
  ok(rep.tables === 5 && rep.caps === 5, '5 張表格與 5 個表名');
  ok(rep.noCap.length === 0, '每個表名都對應到表格');
  ok(rep.groups === 3, '表4-1 切出 3 個分項工程子表格');
  ok(rep.figs === 1, '1 個圖名');
  ok(rep.seq === 6, 'SEQ 功能變數 6 個');
  await page.screenshot({ path: path.join(outDir, '1-master.png'), fullPage: true });

  // 分類
  const cls = await page.evaluate(() => {
    const B = RG.app.S.master.blocks;
    const find = (s) => B.find((b) => b.text.includes(s));
    return {
      law: find('營造安全衛生設施標準').auto.cat,
      t31: B.find((b) => b.kind === 'table' && b.capNo === '3-1').auto,
      gB: B.find((b) => b.kind === 'rows' && b.title.includes('擋土')).auto.tags,
      gC: B.find((b) => b.kind === 'rows' && b.title.includes('骨櫃')).auto.tags,
      intro: find('本案為臺中市').auto.cat,
    };
  });
  console.log('classify', JSON.stringify(cls));
  ok(cls.law === 'T', '法規節為 T');
  ok(cls.t31.cat === 'C' && cls.t31.tags.includes('擋土柱+斜撐'), '表3-1 為 C 並標擋土柱+斜撐');
  ok(cls.gB.includes('擋土柱+斜撐'), '擋土子表格標籤正確');
  ok(cls.gC.includes('納骨塔/骨櫃'), '骨櫃子表格標籤正確');
  ok(['F', 'M'].includes(cls.intro), '含專案事實的段落為 F／M（不因工程名稱含「納骨塔」被判為 C）');

  // 3. 檢查母本
  await page.evaluate(() => RG.app.go('check'));
  await page.click('button:has-text("檢查母本")');
  await page.waitForSelector('text=檢查結果');
  const iss = await page.evaluate(() => RG.app.S.lastCheck.issues.map((i) => i.item + '｜' + i.desc));
  console.log(iss.join('\n'));
  const has = (s) => iss.some((x) => x.includes(s));
  ok(has('表號引用錯誤') && has('表3-7'), '抓到不存在的表3-7');
  ok(has('表名引用不符'), '抓到表名引用不符');
  ok(has('風險值錯誤'), '抓到風險值錯誤');
  ok(has('生方向指示燈') || iss.some((x) => x.startsWith('錯字')), '抓到已知錯字');
  ok(has('年雨量'), '抓到年雨量量級問題');
  ok(has('日期'), '抓到評估日期與封面不一致');
  await page.screenshot({ path: path.join(outDir, '6-check-master.png'), fullPage: true });

  // 還原測試
  await page.evaluate(() => RG.app.go('generate'));
  await page.click('button:has-text("執行還原測試")');
  await page.waitForSelector('text=還原測試結果');
  const rs = await page.evaluate(() => { const r = RG.app.S.restore; return { ratio: r.ratio, t: [r.tablesMaster, r.tablesOut], c: [r.chaptersMaster, r.chaptersOut], diffs: r.diffs.slice(0, 5) }; });
  console.log('restore', JSON.stringify(rs));
  ok(rs.ratio >= 0.98, '還原測試 ≥ 98%');
  ok(rs.t[0] === rs.t[1] && rs.c[0] === rs.c[1], '還原測試表格數、章數相同');

  // 新案：無地下室、平地、SC，換名稱
  await page.evaluate(() => {
    const S = RG.app.S;
    Object.values(S.facts).forEach((f) => { f.newValue = ''; });
    Object.assign(S.facts['工程名稱'], { newValue: '新竹縣竹北市虛構活動中心新建工程' });
    Object.assign(S.facts['構造形式'], { newValue: 'SC' });
    Object.assign(S.facts['地上層數'], { newValue: '3' });
    Object.assign(S.facts['地下層數'], { newValue: '0' });
    Object.assign(S.facts['基地條件'], { newValue: '平地' });
    Object.assign(S.facts['設施用途'], { newValue: '活動中心' });
    Object.assign(S.facts['地址'], { masterLiterals: '松竹路', newValue: '' }); // 不編造：沒提供就待填
    S.confirmed = true;
    RG.app.go('plan');
  });
  await page.screenshot({ path: path.join(outDir, '4-plan.png'), fullPage: true });
  await page.evaluate(() => RG.app.go('generate'));
  const [dl] = await Promise.all([
    page.waitForEvent('download').catch(() => null),
    (async () => {
      await page.click('button:has-text("產出新案報告")');
      await page.waitForSelector('text=產出結果');
      await page.click('button:has-text("下載報告書 .docx")');
    })(),
  ]);
  if (dl) await dl.saveAs(path.join(outDir, 'new_case.docx'));
  const outInfo = await page.evaluate(() => {
    const o = RG.app.S.lastOutput;
    const B = o.blocks;
    return {
      text: B.map((b) => b.text).join('\n'),
      caps: B.filter((b) => b.kind === 'caption').map((b) => b.text),
      types: o.diffs.reduce((m, d) => ((m[d.type] = (m[d.type] || 0) + 1), m), {}),
      review: o.diffs.filter((d) => d.type === '待審').map((d) => d.reason),
      fTodo: o.diffs.filter((d) => d.type === '待填' && d.reason.startsWith('F 專案事實')).map((d) => d.after),
      removed: o.removedCapNos,
    };
  });
  console.log('caps', outInfo.caps.join(' / '));
  console.log('types', JSON.stringify(outInfo.types));
  ok(!outInfo.text.includes('擋土柱'), '擋土相關內容已移除');
  ok(!outInfo.text.includes('骨櫃'), '骨櫃子表格已移除');
  ok(!outInfo.caps.some((c) => c.includes('擋土設施')), '表3-1 表名已移除');
  ok(!outInfo.text.includes('表3-1 擋土'), '表5-1 中引用已刪除表3-1 的列已移除');
  ok(outInfo.text.includes('新竹縣竹北市虛構活動中心新建工程'), '工程名稱已替換');
  ok(!outInfo.text.includes('臺中市北屯區第28公墓新建納骨塔工程'), '母本工程名稱無殘留');
  ok(outInfo.text.includes('本案為新竹縣竹北市虛構活動中心新建工程'), '含專案事實的段落沒有被誤刪');
  ok(outInfo.text.includes('地上3層'), '地上層數已替換');
  ok(outInfo.text.includes('【待填：地址】') && !outInfo.text.includes('松竹路'), '未提供的地址標為待填，不帶入母本值');
  ok(outInfo.text.includes('逃生方向指示燈'), '錯字已自動更正');
  ok(!outInfo.text.includes('代辨機關'), '代辨機關已更正');
  console.log('review', JSON.stringify(outInfo.review));
  console.log('fTodo', JSON.stringify(outInfo.fTodo));
  ok(outInfo.review.some((r) => r.includes('地下層數') && r.includes('0')), 'M 類：地下層數為 0 時標「待審」');
  ok(outInfo.fTodo.includes('【待填：專案事實】'), 'F 類：沒有對應欄位的段落整段標【待填】');
  ok(!outInfo.text.includes('115年6月'), 'F 類：母本日期不帶入新案');
  ok(outInfo.text.includes('本案為新竹縣竹北市虛構活動中心新建工程'), 'M 類：句型保留');

  // 產出檔檢查
  await page.click('button:has-text("對產出檔執行一致性檢查")');
  await page.waitForSelector('text=檢查結果');
  const iss2 = await page.evaluate(() => RG.app.S.lastCheck.issues.map((i) => i.item + '｜' + i.textA));
  console.log(iss2.join('\n'));
  ok(iss2.some((x) => x.startsWith('母本殘留')), '產出檔殘留檢查有結果（北屯區／松竹路等未提供新值）');
  ok(iss2.some((x) => x.startsWith('待填')), '列出待填項目');
  ok(!iss2.some((x) => x.startsWith('風險值錯誤')), '產出檔風險值已全部正確');
  await page.screenshot({ path: path.join(outDir, '6-check-output.png'), fullPage: true });

  for (const v of ['classify', 'project', 'rules', 'help']) {
    await page.evaluate((x) => RG.app.go(x), v);
    await page.screenshot({ path: path.join(outDir, `view-${v}.png`), fullPage: false });
  }
  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
