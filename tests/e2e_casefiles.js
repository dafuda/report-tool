/* 新案資料檔讀取（PowerPoint、Excel）：
 *   python tests/make_case_files.py tests/fixtures
 *   node tests/e2e_casefiles.js */
const path = require('path');
const { chromium } = require('playwright');
let failed = 0;
const ok = (c, m) => {
  console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`);
  if (!c) failed++;
};
(async () => {
  const fx = path.join(__dirname, 'fixtures');
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1300, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('file://' + path.resolve(__dirname, '..', 'index.html'));
  await page.evaluate(() => RG.app.go('project'));
  // 以「檔案內容＋檔名」上傳：測試環境中 Playwright 以路徑上傳中文檔名會失敗，實際瀏覽器沒有這個問題
  const fs = require('fs');
  const payload = (n) => ({ name: n, mimeType: 'application/octet-stream', buffer: fs.readFileSync(path.join(fx, n)) });
  await page.setInputFiles('label:has-text("選擇新案檔案") input', [payload('新案_設計說明簡報.pptx'), payload('新案_基本資料.xlsx'), payload('舊版簡報.ppt')]);
  await page.waitForFunction(() => RG.app.S.caseFiles.length === 3);
  const files = await page.evaluate(() => RG.app.S.caseFiles.map((f) => ({ name: f.name, len: f.text.length, error: f.error || '' })));
  console.log(JSON.stringify(files));
  const pp = files.find((f) => f.name.endsWith('.pptx'));
  ok(pp && pp.len > 0 && !pp.error, 'PowerPoint（.pptx）讀取成功');
  ok(files.find((f) => f.name.endsWith('.xlsx')).len > 0, 'Excel（.xlsx）讀取成功');
  ok(/另存新檔/.test(files.find((f) => f.name.endsWith('.ppt')).error), '舊版 .ppt 提示另存為 .pptx');
  const sg = await page.evaluate(() => Object.fromEntries(Object.entries(RG.app.S.suggestions).map(([k, v]) => [k, v[0]])));
  console.log(JSON.stringify(sg, null, 1));
  ok(sg['工程名稱'] && sg['工程名稱'].value === '新竹縣竹北市虛構活動中心新建工程' && sg['工程名稱'].src.includes('第 1 張投影片'), '簡報文字方塊：工程名稱，來源標到第 1 張投影片');
  ok(sg['構造形式'] && sg['構造形式'].value === 'SC' && sg['構造形式'].src.includes('第 2 張投影片'), '簡報表格：構造形式，來源標到第 2 張投影片');
  ok(sg['地下層數'] && sg['地下層數'].value === '0', '簡報表格：地下層數');
  ok(sg['基地面積'] && sg['基地面積'].value.startsWith('2,500') && sg['基地面積'].src.includes('備註'), '簡報備註：基地面積，來源標到投影片備註');
  ok(sg['工期'] && sg['工期'].value === '540日曆天', '簡報多行文字：工期');
  ok(sg['主辦機關'] && sg['主辦機關'].value === '新竹縣政府' && sg['主辦機關'].src.includes('工作表「基本資料」】第2列'), 'Excel：主辦機關，來源標到工作表與列號（第 2 列）');
  await page.click('button:has-text("空白欄位全部採用建議")');
  const v = await page.evaluate(() => ({ name: RG.app.S.facts['工程名稱'].newValue, src: RG.app.S.facts['工程名稱'].source }));
  ok(v.name === '新竹縣竹北市虛構活動中心新建工程' && v.src.includes('投影片'), '採用建議後，欄位填入新案值並記錄來源');
  ok(errors.length === 0, '頁面沒有 JavaScript 錯誤' + (errors.length ? '：' + errors.join('; ') : ''));
  await browser.close();
  console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
  process.exit(failed ? 1 : 0);
})();
