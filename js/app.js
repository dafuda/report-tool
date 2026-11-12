/* 畫面與流程控制 */
(function () {
  'use strict';
  const RG = window.RG;
  const D = RG.docx;
  const K = RG.kb;
  const STORE_KEY = 'rg-report-tool-v1';

  /* ---------- 狀態 ---------- */
  const S = {
    master: null,
    facts: RG.newFacts(),
    rules: RG.clone(RG.DEFAULT_RULES),
    overrides: {},
    planOverrides: {},
    manualTags: {},
    options: { clean: false, autoFix: true, recomputeRisk: true, updateFields: true, fillMissingF: true, markReview: true, figurePlaceholder: true },
    confirmed: false,
    masterSig: '',
    memory: RG.mem.empty(),
    memResult: null,
    kwSuggest: null,
    kwDone: {},
    projectName: '新案A',
    caseFiles: [],
    suggestions: {},
    lastOutput: null,
    lastCheck: null,
    restore: null,
    view: 'master',
    ui: { cls: { page: 0, chapter: '', cat: '', kind: '', q: '', onlyFlag: false, onlyMod: false, mem: '' }, memv: { page: 0, q: '', cat: '', tag: '', master: '', kind: '', check: '', sel: {} }, plan: { page: 0, filter: 'all', q: '' }, showMaster: false, diffType: '', checkItem: '' },
  };
  const PERSIST = ['facts', 'rules', 'overrides', 'planOverrides', 'manualTags', 'options', 'confirmed', 'masterSig', 'projectName'];

  function save() {
    try {
      const o = {};
      PERSIST.forEach((k) => (o[k] = S[k]));
      localStorage.setItem(STORE_KEY, JSON.stringify(o));
    } catch (e) {
      /* 無法儲存時忽略（例如私密瀏覽） */
    }
  }
  function applyLoaded(o) {
    if (!o) return;
    for (const k of PERSIST) if (o[k] !== undefined) S[k] = o[k];
    const base = RG.newFacts();
    for (const k of Object.keys(base)) S.facts[k] = Object.assign(base[k], S.facts[k] || {});
    S.rules = Object.assign(RG.clone(RG.DEFAULT_RULES), S.rules || {});
    S.options = Object.assign({ clean: false, autoFix: true, recomputeRisk: true, updateFields: true, fillMissingF: true, markReview: true, figurePlaceholder: true }, S.options || {});
  }
  try {
    applyLoaded(JSON.parse(localStorage.getItem(STORE_KEY) || 'null'));
  } catch (e) {
    /* ignore */
  }

  /* ---------- 記憶庫（依文字內容記住分類與標籤，跨母本、跨專案共用） ---------- */
  const MEM_KEY = 'rg-memory-v1';
  try {
    const m = JSON.parse(localStorage.getItem(MEM_KEY) || 'null');
    if (m && m.entries) S.memory = Object.assign(RG.mem.empty(), m);
  } catch (e) {
    /* ignore */
  }
  let memWarned = false;
  function saveMemory() {
    try {
      localStorage.setItem(MEM_KEY, JSON.stringify(S.memory));
    } catch (e) {
      if (!memWarned) toast('記憶庫太大，無法存在瀏覽器中。請用「匯出記憶庫」另存備份。', 'bad');
      memWarned = true;
    }
  }
  const memCtx = () => ({ overrides: S.overrides, planOverrides: S.planOverrides, sig: S.masterSig });
  /* 把區塊目前的分類／標籤／工項規劃記入記憶，並標為已確認 */
  function remember(blocks) {
    const list = blocks.filter((b) => RG.mem.memorable(b));
    RG.mem.recordMany(S.memory, list, memCtx());
    list.forEach((b) => (b.mem = { state: 'confirmed' }));
    saveMemory();
    return list.length;
  }
  const needsReview = (b) => b.mem && (b.mem.state === 'fuzzy' || b.mem.state === 'new');

  /* ---------- 小工具 ---------- */
  function h(tag, props, ...children) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(props || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else if (['value', 'checked', 'selected', 'disabled', 'open', 'hidden', 'multiple'].includes(k)) el[k] = v;
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const c of children.flat(Infinity)) {
      if (c == null || c === false) continue;
      el.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
    }
    return el;
  }
  const $ = (s) => document.querySelector(s);
  const fmtSize = (n) => (n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : (n / 1024).toFixed(0) + ' KB');
  const short = (s, n = 120) => (s && s.length > n ? s.slice(0, n) + '…' : s || '');

  function toast(msg, type = '') {
    const t = h('div', { class: 'toast ' + type }, msg);
    $('#toasts').appendChild(t);
    setTimeout(() => t.remove(), type === 'bad' ? 7000 : 3500);
  }
  async function busy(msg, fn) {
    $('#busyMsg').textContent = msg;
    $('#busy').hidden = false;
    await new Promise((r) => setTimeout(r, 30));
    try {
      return await fn();
    } catch (e) {
      console.error(e);
      toast('發生錯誤：' + e.message, 'bad');
    } finally {
      $('#busy').hidden = true;
    }
  }
  function download(name, data, mime = 'application/octet-stream') {
    const blob = data instanceof Blob ? data : new Blob([data], { type: mime });
    const a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  }
  function xlsxDownload(name, sheets) {
    const wb = XLSX.utils.book_new();
    for (const s of sheets) {
      const ws = XLSX.utils.aoa_to_sheet(s.rows);
      if (s.cols) ws['!cols'] = s.cols.map((w) => ({ wch: w }));
      ws['!autofilter'] = s.rows.length > 1 ? { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: s.rows.length - 1, c: s.rows[0].length - 1 } }) } : undefined;
      XLSX.utils.book_append_sheet(wb, ws, s.name);
    }
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    download(name, out, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }
  async function readXlsxRows(file, sheetName) {
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const ws = wb.Sheets[sheetName] || wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: '' });
  }
  const filePicker = (accept, onFile, multiple = false) => {
    const inp = h('input', { type: 'file', accept, hidden: true, multiple });
    inp.addEventListener('change', () => {
      if (inp.files.length) onFile(multiple ? Array.from(inp.files) : inp.files[0]);
      inp.value = '';
    });
    return inp;
  };
  const pickBtn = (label, accept, onFile, cls = 'btn', multiple = false) => h('label', { class: cls }, label, filePicker(accept, onFile, multiple));

  /* ---------- 衍生資料 ---------- */
  const tagInfo = () => K.projectTags(S);
  const planRes = (opts) => K.plan(S.master.blocks, S.overrides, tagInfo().set, S.planOverrides, opts);
  const posOf = (b) => `${b.chapter}${b.capNo ? ' ' + (b.capType || '表') + b.capNo : ''}`;
  function invalidateOutputs() {
    S.lastOutput = null;
    S.restore = null;
  }

  /* ---------- 導覽 ---------- */
  const VIEWS = [
    { id: 'master', n: 1, label: '母本解析', sub: '讀入母本、結構驗證', render: viewMaster },
    { id: 'classify', n: 2, label: '分類與標籤確認', sub: 'F／T／C／M 與條件標籤', render: viewClassify },
    { id: 'project', n: 3, label: '專案資料', sub: '新案事實與條件標籤', render: viewProject },
    { id: 'plan', n: 4, label: '工項規劃', sub: '選出適用的內容', render: viewPlan },
    { id: 'generate', n: 5, label: '產出報告', sub: '報告書＋差異清單', render: viewGenerate },
    { id: 'check', n: 6, label: '一致性檢查', sub: '跨章節矛盾與殘留', render: viewCheck },
    { sep: true },
    { id: 'memory', n: '≡', label: '記憶庫', sub: '檢視、修改、完整性檢查', render: viewMemory },
    { id: 'rules', n: '⚙', label: '規則設定', sub: '標籤、更正、檢查規則', render: viewRules },
    { id: 'help', n: '?', label: '使用說明', sub: '流程與注意事項', render: viewHelp },
  ];
  function isDone(id) {
    if (id === 'master') return !!S.master;
    if (id === 'classify') return !!S.master && Object.keys(S.overrides).length > 0;
    if (id === 'project') return S.confirmed;
    if (id === 'generate') return !!S.lastOutput;
    if (id === 'check') return !!S.lastCheck;
    return false;
  }
  function renderNav() {
    const nav = $('#nav');
    nav.innerHTML = '';
    for (const v of VIEWS) {
      if (v.sep) {
        nav.appendChild(h('div', { class: 'nav-sep' }));
        continue;
      }
      nav.appendChild(
        h('button', { class: `nav-item ${S.view === v.id ? 'active' : ''} ${isDone(v.id) ? 'done' : ''}`, onclick: () => go(v.id) },
          h('span', { class: 'nav-num' }, isDone(v.id) && typeof v.n === 'number' ? '✓' : v.n),
          h('span', { class: 'nav-label' }, v.label, h('span', { class: 'nav-sub' }, v.sub)))
      );
    }
  }
  function go(id) {
    S.view = id;
    render();
    window.scrollTo(0, 0);
  }
  function render() {
    renderNav();
    const main = $('#main');
    const scroll = window.scrollY;
    main.innerHTML = '';
    const v = VIEWS.find((x) => x.id === S.view) || VIEWS[0];
    main.appendChild(v.render());
    window.scrollTo(0, scroll);
  }
  const needMaster = () =>
    h('div', { class: 'card' }, h('div', { class: 'note warn' }, '請先在「1 母本解析」讀入母本 .docx。'), h('button', { class: 'btn primary', onclick: () => go('master') }, '前往母本解析'));

  /* ====================================================================
   * 1. 母本解析
   * ==================================================================== */
  async function readRels(zip) {
    const f = zip.file('word/_rels/document.xml.rels');
    if (!f) return {};
    const d = D.parseXml(await f.async('string'));
    const map = {};
    for (const r of d.getElementsByTagName('Relationship')) map[r.getAttribute('Id')] = r.getAttribute('Target');
    return map;
  }
  /* 圖片縮圖（png／jpg 等可預覽；emf／wmf 只顯示檔名） */
  function thumb(b) {
    const m = S.master;
    if (!m || !m.zip || !b.images || !b.images.length) return null;
    const target = m.rels[b.images[0]] || '';
    const path = target.startsWith('/') ? target.slice(1) : 'word/' + target.replace(/^\.\//, '');
    const ext = (path.split('.').pop() || '').toLowerCase();
    const more = b.images.length > 1 ? `（共 ${b.images.length} 張）` : '';
    if (!['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'].includes(ext)) return h('div', { class: 'reason' }, `${path.split('/').pop()}（無法預覽）${more}`);
    S.thumbs = S.thumbs || {};
    const img = h('img', { alt: '', style: { width: '140px', height: '80px', objectFit: 'contain', background: '#f4f5f7', border: '1px solid var(--line)', borderRadius: '4px', display: 'block', marginTop: '4px', imageRendering: 'auto' } });
    if (S.thumbs[path]) img.src = S.thumbs[path];
    else if (m.zip.file(path)) m.zip.file(path).async('blob').then((blob) => { S.thumbs[path] = URL.createObjectURL(blob); img.src = S.thumbs[path]; });
    return h('div', {}, img, more ? h('div', { class: 'reason' }, more) : null);
  }
  const blockText = (b) => {
    if (b.kind === 'figure') return `［${b.capType === '表' ? '表格截圖' : '圖'}］${b.capNo ? (b.capType || '圖') + b.capNo + ' ' : ''}${b.capName || '（沒有圖名）'}`;
    return b.kind === 'rows' ? b.title + '\n' + b.text : b.text || '（空白）';
  };
  const kindLabel = (b) => (b.kind === 'figure' && b.capType === '表' ? '表格截圖' : b.boxText && b.kind === 'p' ? '段落（含文字方塊）' : K.KIND_LABEL[b.kind] || b.kind);

  async function loadMaster(file) {
    await busy(`解析母本中（${fmtSize(file.size)}）…`, async () => {
      const t0 = performance.now();
      const buffer = await file.arrayBuffer();
      const { zip, doc, styles } = await D.load(buffer);
      const blocks = D.parseBlocks(doc, styles);
      const hfBlocks = await D.headerFooterBlocks(zip);
      const rels = await readRels(zip);
      K.classify(blocks, S.rules, S.facts);
      const report = D.structureReport(doc, styles, blocks);
      const sig = `${file.name}:${file.size}`;
      const prevSig = S.masterSig;
      const sameMaster = prevSig === sig;
      // 目前畫面上的母本：先把修改記入記憶，避免遺失
      if (S.master) {
        const touched = S.master.blocks.filter((b) => S.overrides[b.id] || S.planOverrides[b.id]);
        if (touched.length) RG.mem.recordMany(S.memory, touched, memCtx());
      }
      if (sameMaster) {
        // 同一份母本：區塊編號可信，舊的修改直接保留，並補記入記憶（相容舊版本的資料）
        const touched = blocks.filter((b) => S.overrides[b.id] || S.planOverrides[b.id]);
        if (touched.length) RG.mem.recordMany(S.memory, touched, { overrides: S.overrides, planOverrides: S.planOverrides, sig });
      } else {
        // 不同母本：區塊編號對不上，改用文字內容對應記憶
        S.overrides = {};
        S.planOverrides = {};
      }
      const res = RG.mem.apply(S.memory, blocks, { overrides: S.overrides, planOverrides: S.planOverrides, sig, prevSig, keepExisting: sameMaster });
      S.memResult = Object.assign(res, { name: file.name, prevName: prevSig && !sameMaster ? prevSig.replace(/:\d+$/, '') : '', sameMaster });
      saveMemory();
      report.headerFooter = hfBlocks.length;
      S.thumbs = {};
      S.master = { name: file.name, size: file.size, buffer, blocks, hfBlocks, zip, rels, report, ms: Math.round(performance.now() - t0) };
      S.masterSig = sig;
      S.ui.cls.mem = '';
      invalidateOutputs();
      S.lastCheck = null;
      save();
      toast(`母本解析完成：${blocks.filter((b) => !b.parentId).length} 個區塊。記憶對應：沿用 ${res.exact}、推定 ${res.fuzzy}、新段落 ${res.fresh}`, 'ok');
      render();
    });
  }

  function viewMaster() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '1 母本解析與結構驗證'),
      h('p', { class: 'view-desc' }, '讀入母本（設計階段施工風險評估報告 .docx）。工具會拆出每個段落與表格，辨識章節、表名、圖名與子表格，並和規劃書第 3 節的預期結構比對。檔案只在您的電腦上處理，不會上傳。'));
    const inp = filePicker('.docx', loadMaster);
    const drop = h('div', { class: 'dropzone', onclick: () => inp.click() }, h('div', { class: 'big' }, S.master ? '重新選擇母本' : '選擇或拖曳母本 .docx 到這裡'), h('div', { class: 'muted' }, '例：28公墓-風險評估報告(進版)1150615 - 測試.docx'), inp);
    drop.addEventListener('dragover', (e) => {
      e.preventDefault();
      drop.classList.add('over');
    });
    drop.addEventListener('dragleave', () => drop.classList.remove('over'));
    drop.addEventListener('drop', (e) => {
      e.preventDefault();
      drop.classList.remove('over');
      const f = e.dataTransfer.files[0];
      if (f) loadMaster(f);
    });
    if (!S.master) {
      wrap.appendChild(h('div', { class: 'card' }, drop));
      return wrap;
    }
    const m = S.master;
    const r = m.report;
    wrap.appendChild(h('div', { class: 'card' }, h('div', { class: 'row between' }, h('div', {}, h('b', {}, m.name), h('div', { class: 'muted' }, `${fmtSize(m.size)} · 解析 ${m.ms} ms`)),
      h('div', { class: 'row' }, h('button', { class: 'btn', onclick: () => exportStructure() }, '匯出結構驗證報告 .xlsx'), h('button', { class: 'btn', onclick: () => inp.click() }, '重新選擇母本'), inp))));

    const kpi = (k, v, s, cls = '') => h('div', { class: 'kpi ' + cls }, h('div', { class: 'k' }, k), h('div', { class: 'v' }, v), s ? h('div', { class: 's' }, s) : null);
    wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, '結構摘要'),
      h('div', { class: 'kpis' },
        kpi('章', r.chapters.length, '預期 6 章', r.missingChapters.length ? 'warn' : 'ok'),
        kpi('節', r.chapters.reduce((n, c) => n + c.sections.length, 0)),
        kpi('表格', r.tableCount, `表名 ${r.tableCaptionCount} 個`),
        kpi('子表格（分項工程）', r.groupCount, '表4-13 等'),
        kpi('圖名', r.figureCaptionCount, `圖 ${r.figureCount} 張`),
        kpi('表格截圖', r.imageTables.length, r.imageTables.length ? r.imageTables.slice(0, 4).join('、') + (r.imageTables.length > 4 ? '…' : '') : '表名下方是圖片'),
        kpi('文字方塊', r.boxParagraphs, '含文字方塊的段落'),
        kpi('頁首頁尾', r.headerFooter || 0, '有文字的頁首頁尾'),
        kpi('分節', r.sectionBreaks, '母本約 57 個'),
        kpi('段落', r.paragraphCount),
        kpi('SEQ 功能變數', r.fields.SEQ || 0, `TOC ${r.fields.TOC || 0}・PAGEREF ${r.fields.PAGEREF || 0}`))));

    const li = (ok, text, warnOnly) => h('li', {}, h('span', { class: ok ? 'ic-ok' : warnOnly ? 'ic-warn' : 'ic-bad' }, ok ? '✓' : '!'), h('span', {}, text));
    wrap.appendChild(h('div', { class: 'grid2' },
      h('div', { class: 'card' }, h('h2', {}, '與規劃書預期結構比對'),
        h('ul', { class: 'checklist' },
          li(!r.missingChapters.length, r.missingChapters.length ? `缺少章：${r.missingChapters.join('、')}` : '第一章～第六章齊全'),
          ...r.expectedStyles.map((s) => li(s.defined, `樣式「${s.name}」：${s.defined ? '已定義' : '未定義'}，用於 ${s.used} 個段落`, true)),
          li(!r.captionsWithoutTable.length, r.captionsWithoutTable.length ? `找不到對應表格的表名：${r.captionsWithoutTable.slice(0, 12).join('、')}${r.captionsWithoutTable.length > 12 ? '…' : ''}` : '每個表名都有對應的表格', true),
          li(!r.tablesWithoutCaption, `沒有表名的表格：${r.tablesWithoutCaption} 張`, true),
          li(!r.numberGaps.length, r.numberGaps.length ? `表號不連續（缺）：${r.numberGaps.slice(0, 15).join('、')}${r.numberGaps.length > 15 ? '…' : ''}` : '表號連續', true),
          li(!r.duplicateCaptions.length, r.duplicateCaptions.length ? `重複表號：${r.duplicateCaptions.join('、')}` : '無重複表號', true),
          li((r.fields.SEQ || 0) > 0, `表名／圖名的 SEQ 功能變數：${r.fields.SEQ || 0} 個`, true))),
      h('div', { class: 'card' }, h('h2', {}, '章節大綱'),
        h('div', { class: 'outline' }, r.chapters.map((c) => h('details', {}, h('summary', {}, c.title, h('span', { class: 'muted' }, `　${c.sections.length} 節 · ${c.blocks} 個區塊`)), h('ol', {}, c.sections.map((s) => h('li', {}, s)))))))));

    wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, '表格清單 ', h('span', { class: 'muted' }, `${r.tables.length} 張`)),
      h('div', { class: 'tbl-wrap' }, h('table', { class: 'data' },
        h('thead', {}, h('tr', {}, ['區塊', '章', '表號', '表名', '列數', '子表格', '合併格'].map((x) => h('th', {}, x)))),
        h('tbody', {}, r.tables.map((t) => h('tr', {}, h('td', {}, t.id), h('td', {}, t.chapter), h('td', {}, t.capNo ? '表' + t.capNo : h('span', { class: 'muted' }, '（無）')), h('td', {}, t.capName), h('td', { class: 'num' }, t.rows), h('td', { class: 'num' }, t.groups || ''), h('td', { class: 'num' }, t.merged || ''))))))));
    return wrap;
  }
  function exportStructure() {
    const r = S.master.report;
    xlsxDownload('母本結構驗證報告.xlsx', [
      { name: '摘要', cols: [28, 60], rows: [['項目', '結果'], ['檔名', S.master.name], ['章數', r.chapters.length], ['缺少章', r.missingChapters.join('、')], ['表格數', r.tableCount], ['表名數', r.tableCaptionCount], ['子表格數', r.groupCount], ['圖名數', r.figureCaptionCount], ['分節數', r.sectionBreaks], ['SEQ', r.fields.SEQ || 0], ['TOC', r.fields.TOC || 0], ['PAGEREF', r.fields.PAGEREF || 0], ['找不到表格的表名', r.captionsWithoutTable.join('、')], ['表號缺號', r.numberGaps.join('、')], ['重複表號', r.duplicateCaptions.join('、')], ...r.expectedStyles.map((s) => ['樣式 ' + s.name, `${s.defined ? '已定義' : '未定義'}，使用 ${s.used}`])] },
      { name: '章節', cols: [10, 40, 60], rows: [['章', '章名', '節'], ...r.chapters.map((c) => [c.key, c.title, c.sections.join('\n')])] },
      { name: '表格', cols: [8, 8, 10, 40, 8, 8, 8], rows: [['區塊', '章', '表號', '表名', '列數', '子表格', '合併格'], ...r.tables.map((t) => [t.id, t.chapter, t.capNo, t.capName, t.rows, t.groups, t.merged])] },
    ]);
  }

  /* ====================================================================
   * 2. 分類與標籤確認
   * ==================================================================== */
  const PAGE = 100;
  function viewClassify() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '2 分類與標籤確認'),
      h('p', { class: 'view-desc' }, '每個段落與表格都先由工具初判為 F 專案事實、T 通用範本、C 條件式、M 混合，並標上適用條件標籤。請確認或修改（只確認有疑義的也可以）。修改過的列會以淺黃底顯示。'));
    wrap.appendChild(classifyLegend());
    if (!S.master) return wrap.appendChild(needMaster()), wrap;
    const f = S.ui.cls;
    const blocks = S.master.blocks.filter((b) => !['empty', 'other'].includes(b.kind));
    const counts = { F: 0, T: 0, C: 0, M: 0 };
    blocks.forEach((b) => counts[K.effective(b, S.overrides).cat]++);
    const chapters = [...new Set(blocks.map((b) => b.chapter))];
    const q = f.q.trim();
    const list = blocks.filter((b) => {
      const e = K.effective(b, S.overrides);
      if (f.chapter && b.chapter !== f.chapter) return false;
      if (f.cat && e.cat !== f.cat) return false;
      if (f.kind && b.kind !== f.kind) return false;
      if (f.onlyFlag && e.cat === 'T') return false;
      if (f.onlyMod && !S.overrides[b.id]) return false;
      if (f.mem === 'todo' && !needsReview(b)) return false;
      if (f.mem && f.mem !== 'todo' && !(b.mem && b.mem.state === f.mem)) return false;
      if (q && !(b.text.includes(q) || e.tags.join(' ').includes(q) || b.id === q || ('表' + (b.capNo || '')) === q)) return false;
      return true;
    });
    const pages = Math.max(1, Math.ceil(list.length / PAGE));
    f.page = Math.min(f.page, pages - 1);
    const upd = (k, v) => {
      f[k] = v;
      f.page = 0;
      render();
    };
    const sel = (k, opts) => h('select', { onchange: (e) => upd(k, e.target.value) }, opts.map(([v, l]) => h('option', { value: v, selected: f[k] === v }, l)));

    wrap.appendChild(memoryCard(blocks, list));
    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: 'row between' },
        h('div', { class: 'row' }, Object.entries(counts).map(([c, n]) => h('span', {}, h('span', { class: 'badge cat cat-' + c }, c), ` ${K.CATS[c]} ${n}`)), h('span', { class: 'muted' }, `已修改 ${Object.keys(S.overrides).length} 筆`)),
        h('div', { class: 'row' },
          h('button', { class: 'btn', onclick: exportClassify }, '匯出分類確認表 .xlsx'),
          pickBtn('匯入確認結果', '.xlsx', importClassify),
          h('button', { class: 'btn', onclick: reclassify, title: '依目前的規則重新初判（保留您的修改）' }, '重新初判'))),
      h('div', { class: 'row', style: { marginTop: '10px' } },
        sel('chapter', [['', '全部章節'], ...chapters.map((c) => [c, c])]),
        sel('cat', [['', '全部分類'], ...Object.entries(K.CATS).map(([c, l]) => [c, `${c} ${l}`])]),
        sel('kind', [['', '全部類型'], ...Object.entries(K.KIND_LABEL).filter(([k]) => !['empty', 'other'].includes(k)).map(([k, l]) => [k, l])]),
        h('input', { type: 'search', placeholder: '搜尋文字、標籤、表號（如 表4-13）', value: f.q, style: { width: '260px' }, onchange: (e) => upd('q', e.target.value) }),
        h('label', { class: 'row' }, h('input', { type: 'checkbox', checked: f.onlyFlag, onchange: (e) => upd('onlyFlag', e.target.checked) }), '只看 F／C／M'),
        h('label', { class: 'row' }, h('input', { type: 'checkbox', checked: f.onlyMod, onchange: (e) => upd('onlyMod', e.target.checked) }), '只看已修改'),
        sel('mem', [['', '全部記憶狀態'], ['todo', '需確認（推定＋新段落）'], ['exact', '完全相同沿用'], ['fuzzy', '依前版推定'], ['new', '新段落'], ['confirmed', '本次已確認']]))));

    const rows = list.slice(f.page * PAGE, (f.page + 1) * PAGE).map((b) => {
      const e = K.effective(b, S.overrides);
      const o = S.overrides[b.id];
      const setO = (patch) => {
        const cur = Object.assign({}, S.overrides[b.id] || {}, patch);
        if (cur.cat === b.auto.cat || !cur.cat) delete cur.cat;
        if (cur.tags && cur.tags.join('、') === b.auto.tags.join('、')) delete cur.tags;
        if (Object.keys(cur).length) S.overrides[b.id] = cur;
        else delete S.overrides[b.id];
        remember([b]);
        invalidateOutputs();
        save();
        render();
      };
      return h('tr', { class: o ? 'modified' : '' },
        h('td', {}, h('div', {}, b.id), h('div', { class: 'reason' }, posOf(b)), memBadge(b),
          needsReview(b) ? h('div', {}, h('button', { class: 'btn small', title: '目前的分類與標籤正確，記入記憶', onclick: () => { remember([b]); render(); } }, '✓ 確認')) : null),
        h('td', {}, h('span', { class: 'badge' }, kindLabel(b))),
        h('td', {}, h('div', { class: 'clamp', title: '點一下展開／收合', onclick: (ev) => ev.currentTarget.classList.toggle('open') }, blockText(b)), b.kind === 'figure' ? thumb(b) : null),
        h('td', {}, h('span', { class: 'badge cat cat-' + b.auto.cat }, b.auto.cat), ' ', b.auto.tags.map((t) => h('span', { class: 'badge tag' }, t)), h('div', { class: 'reason' }, b.auto.reason)),
        h('td', {}, h('select', { onchange: (ev) => setO({ cat: ev.target.value }) }, Object.keys(K.CATS).map((c) => h('option', { value: c, selected: e.cat === c }, `${c} ${K.CATS[c]}`)))),
        h('td', {}, h('input', { type: 'text', value: e.tags.join('、'), placeholder: e.cat === 'C' ? '例：擋土柱+斜撐、地下1層' : '僅 C 類使用', title: e.cat === 'C' ? '新案必須具備全部標籤，這段才會保留' : '只有 C 條件式會用到標籤；此分類填了也不影響產出', style: { width: '200px', opacity: e.cat === 'C' ? 1 : 0.5 }, onchange: (ev) => setO({ tags: ev.target.value.split(/[、,，;；]+/).map((s) => s.trim()).filter(Boolean) }) }),
          o ? h('div', {}, h('button', { class: 'btn small', onclick: () => { delete S.overrides[b.id]; remember([b]); save(); render(); } }, '還原自動判斷')) : null));
    });
    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: 'tbl-wrap' }, h('table', { class: 'data' }, h('thead', {}, h('tr', {}, ['區塊', '類型', '內容', '自動判斷', '確認分類', '確認標籤'].map((x) => h('th', {}, x)))), h('tbody', {}, rows))),
      pager(f, pages, list.length)));
    return wrap;
  }
  function memBadge(b) {
    const m = b.mem;
    if (!m) return null;
    if (m.state === 'exact') return h('div', {}, h('span', { class: 'badge keep', title: '文字與記憶庫中的段落完全相同，已沿用先前的確認結果' }, '沿用記憶'));
    if (m.state === 'confirmed') return h('div', {}, h('span', { class: 'badge keep' }, '已確認'));
    if (m.state === 'fuzzy') return h('div', {}, h('span', { class: 'badge user', title: (m.note ? m.note + '\n' : '') + '前版文字：' + (m.from || '') }, `推定 ${Math.round((m.sim || 0) * 100)}%`));
    return h('div', {}, h('span', { class: 'badge drop', title: '記憶庫中沒有相同或相似的段落，目前為自動判斷' }, '新段落'));
  }
  function memoryCard(blocks, list) {
    const st = RG.mem.stats(S.memory);
    const r = S.memResult;
    const todo = blocks.filter(needsReview).length;
    const toConfirm = list.filter((b) => RG.mem.memorable(b) && b.mem && b.mem.state !== 'confirmed');
    const thrInput = h('input', { type: 'number', min: 70, max: 100, step: 1, value: Math.round((S.memory.threshold || 0.9) * 100), style: { width: '64px' } });
    const card = h('div', { class: 'card' },
      h('div', { class: 'row between' },
        h('h2', {}, '記憶庫 ', h('span', { class: 'muted' }, `${st.count} 筆（來自 ${st.masters} 份母本）${st.conflicts ? `・衝突 ${st.conflicts} 筆` : ''}`)),
        h('div', { class: 'row' },
          h('button', { class: 'btn', onclick: () => go('memory') }, '管理記憶庫'),
          h('button', { class: 'btn', title: '依記憶庫歸納條件標籤的關鍵字', onclick: () => { runKeywordSuggest(); go('rules'); } }, '關鍵字建議'),
          h('button', { class: 'btn', onclick: () => xlsxOrJson() }, '匯出記憶庫'),
          pickBtn('匯入記憶庫（合併）', '.json', importMemory),
          h('button', { class: 'btn ghost', onclick: () => {
            if (!confirm('清除記憶庫中所有記住的分類與標籤？（目前畫面上的分類不受影響）')) return;
            S.memory = Object.assign(RG.mem.empty(), { threshold: S.memory.threshold });
            S.master.blocks.forEach((b) => RG.mem.memorable(b) && (b.mem = { state: 'new' }));
            saveMemory();
            render();
          } }, '清除記憶庫'))),
      h('div', { class: 'muted' }, '你確認或修改過的分類與標籤，會依「段落文字」記住。換母本、母本改版或讀入其他案報告時，相同的段落自動沿用，相似的段落標為「推定」，只需核對推定與新段落。'));
    if (r) {
      card.appendChild(h('div', { class: 'note ' + (todo ? 'warn' : 'ok'), style: { marginTop: '10px' } },
        `讀入「${r.name}」${r.prevName ? `（前一份：${r.prevName}）` : ''}的對應結果：完全相同沿用 ${r.exact} 筆、依前版推定 ${r.fuzzy} 筆、新段落 ${r.fresh} 筆。`,
        todo ? ` 目前還有 ${todo} 筆需要確認。` : ' 全部都已確認。'));
    }
    card.appendChild(h('div', { class: 'row', style: { marginTop: '8px' } },
      todo ? h('button', { class: 'btn primary', onclick: () => { S.ui.cls.mem = 'todo'; S.ui.cls.page = 0; render(); } }, `只看需確認的 ${todo} 筆`) : null,
      h('button', { class: 'btn', disabled: !toConfirm.length, title: '目前篩選結果中的分類與標籤都正確，全部記入記憶', onclick: () => {
        if (!confirm(`將目前篩選出的 ${toConfirm.length} 筆標為「已確認」並記入記憶庫？`)) return;
        const n = remember(toConfirm);
        toast(`已確認並記入 ${n} 筆`, 'ok');
        render();
      } }, `確認目前篩選的 ${toConfirm.length} 筆`),
      h('span', { class: 'muted' }, '推定的相似度門檻'), thrInput, h('span', { class: 'muted' }, '%'),
      h('button', { class: 'btn small', onclick: () => {
        const v = Math.min(100, Math.max(70, parseInt(thrInput.value, 10) || 90));
        S.memory.threshold = v / 100;
        saveMemory();
        S.memResult = Object.assign(RG.mem.apply(S.memory, S.master.blocks, { overrides: S.overrides, planOverrides: S.planOverrides, sig: S.masterSig, keepExisting: false }), { name: S.master.name, prevName: '', sameMaster: true });
        invalidateOutputs();
        save();
        render();
        toast(`已用 ${v}% 重新套用記憶`, 'ok');
      } }, '套用並重新對應')));
    if (r && r.missing && r.missing.length) {
      card.appendChild(h('details', { style: { marginTop: '10px' } },
        h('summary', { style: { cursor: 'pointer' } }, `前一份母本有、這份沒有的段落 ${r.missing.length} 筆（不會套到其他段落，僅供參考）`),
        h('div', { class: 'tbl-wrap', style: { maxHeight: '300px', marginTop: '6px' } }, h('table', { class: 'data' },
          h('thead', {}, h('tr', {}, h('th', {}, '類型'), h('th', {}, '內容（前 120 字）'), h('th', {}, '分類'), h('th', {}, '標籤'))),
          h('tbody', {}, r.missing.slice(0, 300).map((e) => h('tr', {}, h('td', {}, K.KIND_LABEL[e.kind] || e.kind), h('td', {}, e.pv), h('td', {}, h('span', { class: 'badge cat cat-' + e.cat }, e.cat)), h('td', {}, (e.tags || []).join('、')))))))));
    }
    return card;
  }
  function xlsxOrJson() {
    download(`分類記憶庫_${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(S.memory), 'application/json');
  }
  async function importMemory(file) {
    try {
      const m = JSON.parse(await file.text());
      const n = RG.mem.merge(S.memory, m.memory || m);
      saveMemory();
      if (S.master) {
        S.memResult = Object.assign(RG.mem.apply(S.memory, S.master.blocks, { overrides: S.overrides, planOverrides: S.planOverrides, sig: S.masterSig, keepExisting: true }), { name: S.master.name, prevName: '', sameMaster: true });
        save();
      }
      toast(`已合併 ${n} 筆記憶`, 'ok');
      render();
    } catch (e) {
      toast('記憶庫檔案格式錯誤：' + e.message, 'bad');
    }
  }
  function classifyLegend() {
    const row = (c, what, result) => h('tr', {}, h('td', {}, h('span', { class: 'badge cat cat-' + c }, c), ' ', K.CATS[c]), h('td', {}, what), h('td', {}, result));
    return h('details', { class: 'card', open: !S.master || !Object.keys(S.overrides).length },
      h('summary', { style: { cursor: 'pointer', fontWeight: 600 } }, '分類代號與「確認標籤」說明'),
      h('table', { class: 'data', style: { marginTop: '10px' } },
        h('thead', {}, h('tr', {}, h('th', {}, '分類'), h('th', {}, '指的是'), h('th', {}, '產出時的處理'))),
        h('tbody', {},
          row('F', '只屬於本案的資料：工程名稱、地址、面積、日期、人員、表1-1 等', '有設定母本寫法的欄位 → 換成新案值；整段（表格為整格）找不到任何對應欄位 → 改為紅字【待填】，絕不帶入母本數字'),
          row('T', '與案子無關的通用文字：前言、方法說明、法規、災害案例、風險矩陣', '原文照抄，不替換專案事實；只做已知錯字更正'),
          row('C', '依構造、工項、環境條件決定是否出現：擋土、開挖、電梯、骨櫃…', '「確認標籤」全部符合新案條件才保留（並替換其中的專案事實），否則整段／整表刪除'),
          row('M', '固定句型中嵌入專案事實：「本案為地下1層、地上4層RC造…」', '只替換句中的專案事實、保留句型；若替換後可能不通（如地下0層）或句中提到新案沒有的工項 → 淺藍底＋註解「待審」'))),
      h('div', { class: 'muted', style: { marginTop: '8px' } }, '圖片與表格截圖（內容是圖、無法編輯的表格）：F 產出時換成紅框【待補圖：圖x-x 名稱】、T 沿用原圖、C 依確認標籤保留或刪除；依圖名自動初判，沒有圖名或無法判斷時預設為 F。文字方塊內的文字與頁首頁尾，比照所在段落處理專案事實替換與錯字更正。'),
      h('div', { class: 'muted', style: { marginTop: '8px' } }, '確認標籤：只有 C 類會用到。填寫這段內容適用的條件（以頓號分隔），新案必須「全部」具備才保留；例如「擋土柱+斜撐、地下1層」表示有地下室且採擋土柱工法的案子才會出現。新案具備哪些標籤，在「3 專案資料」最下方決定。F／T／M 類填了標籤不會有作用。'));
  }
  function pager(f, pages, total) {
    return h('div', { class: 'pager' },
      h('button', { class: 'btn small', disabled: f.page <= 0, onclick: () => { f.page--; render(); } }, '‹ 上一頁'),
      h('span', { class: 'muted' }, `第 ${f.page + 1} / ${pages} 頁，共 ${total} 筆`),
      h('button', { class: 'btn small', disabled: f.page >= pages - 1, onclick: () => { f.page++; render(); } }, '下一頁 ›'));
  }
  /* 依目前規則重新初判，再用記憶庫還原你確認過的結果（避免自動結果改變時，已確認的段落跟著變） */
  function reclassifyKeepMemory() {
    if (!S.master) return;
    K.classify(S.master.blocks, S.rules, S.facts);
    const r = RG.mem.apply(S.memory, S.master.blocks, { overrides: S.overrides, planOverrides: S.planOverrides, sig: S.masterSig, keepExisting: false });
    S.memResult = Object.assign(r, { name: S.master.name, prevName: '', sameMaster: true });
    invalidateOutputs();
    save();
  }
  function reclassify() {
    reclassifyKeepMemory();
    toast('已依目前規則重新初判（已確認的段落維持你的確認結果）', 'ok');
    render();
  }
  function exportClassify() {
    const MEM_LABEL = { exact: '沿用記憶', fuzzy: '推定', new: '新段落', confirmed: '已確認' };
    const rows = [['區塊', '章', '節', '類型', '表號', '內容（前 300 字）', '自動分類', '自動標籤', '判斷理由', '記憶對應', '確認分類', '確認標籤']];
    for (const b of S.master.blocks) {
      if (['empty', 'other'].includes(b.kind)) continue;
      const e = K.effective(b, S.overrides);
      rows.push([b.id, b.chapter, b.section || '', K.KIND_LABEL[b.kind], b.capNo ? (b.capType || '表') + b.capNo : '', short(blockText(b), 300), b.auto.cat, b.auto.tags.join('、'), b.auto.reason, b.mem ? MEM_LABEL[b.mem.state] + (b.mem.state === 'fuzzy' ? ` ${Math.round(b.mem.sim * 100)}%` : '') : '', e.cat, e.tags.join('、')]);
    }
    xlsxDownload('分類確認表.xlsx', [{ name: '分類確認', cols: [9, 8, 18, 8, 8, 60, 8, 22, 26, 10, 9, 24], rows }, { name: '說明', cols: [80], rows: [['匯入時，檔案中出現的每一列都會視為「已確認」並記入記憶庫。'], ['請修改「確認分類」（F／T／C／M）與「確認標籤」（以頓號分隔）兩欄，存檔後在工具中按「匯入確認結果」。'], ['F 專案事實：只屬於本案的資料，從新案資料擷取'], ['T 通用範本：與案子無關的方法論、法規，直接沿用'], ['C 條件式：依構造、工項、環境條件決定是否出現'], ['M 混合：範本句型中嵌入專案事實']] }]);
  }
  async function importClassify(file) {
    await busy('匯入分類確認表…', async () => {
      const rows = await readXlsxRows(file, '分類確認');
      const byId = Object.fromEntries(S.master.blocks.map((b) => [b.id, b]));
      let n = 0;
      const seen = [];
      for (const r of rows) {
        const b = byId[String(r['區塊']).trim()];
        if (!b) continue;
        seen.push(b);
        const cat = String(r['確認分類'] || '').trim().toUpperCase();
        const tags = String(r['確認標籤'] || '').split(/[、,，;；]+/).map((s) => s.trim()).filter(Boolean);
        const o = {};
        if (K.CATS[cat] && cat !== b.auto.cat) o.cat = cat;
        if (tags.join('、') !== b.auto.tags.join('、')) o.tags = tags;
        if (Object.keys(o).length) {
          S.overrides[b.id] = o;
          n++;
        } else delete S.overrides[b.id];
      }
      const m = remember(seen);
      invalidateOutputs();
      save();
      toast(`已匯入，${n} 筆與自動判斷不同；${m} 筆記入記憶庫`, 'ok');
      render();
    });
  }

  /* ====================================================================
   * 3. 專案資料
   * ==================================================================== */
  function viewProject() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '3 專案資料（新案事實）'),
      h('p', { class: 'view-desc' }, '填入新案的專案事實。可以先讀入新案資料（docx／xlsx），工具會依欄位名稱找出建議值與來源。找不到的欄位請留空，產出時會標成紅字【待填】——絕不會帶入母本的數字。全部確認後請勾選最下方的「已確認」，才能產出報告。'));
    const ti = tagInfo();

    // 新案資料檔
    wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, '新案資料檔 ', h('span', { class: 'muted' }, '設計說明書、興建計畫書、地質報告、預算書…')),
      h('div', { class: 'row' }, pickBtn('選擇新案檔案（可多選）', '.docx,.xlsx,.xls,.csv,.txt,.pdf', addCaseFiles, 'btn primary', true),
        S.caseFiles.length ? h('button', { class: 'btn', onclick: () => { runSuggest(); render(); } }, '重新尋找建議值') : null,
        S.caseFiles.length ? h('button', { class: 'btn', onclick: acceptAllSuggestions }, '空白欄位全部採用建議') : null,
        S.caseFiles.length ? h('button', { class: 'btn ghost', onclick: () => { S.caseFiles = []; S.suggestions = {}; render(); } }, '清除檔案') : null),
      S.caseFiles.length ? h('ul', { class: 'checklist', style: { marginTop: '8px' } }, S.caseFiles.map((f) => h('li', {}, h('span', { class: f.error ? 'ic-warn' : 'ic-ok' }, f.error ? '!' : '✓'), h('span', {}, f.name, h('span', { class: 'muted' }, f.error ? `　${f.error}` : `　${f.text.length.toLocaleString()} 字`))))) : h('div', { class: 'muted', style: { marginTop: '6px' } }, '掃描 PDF 與 PDF 文字擷取、Claude API 結構化擷取將在後續版本（Python 版）提供；目前請先將 PDF 轉為 Word，或直接手動填寫。')));

    // 欄位表
    const card = h('div', { class: 'card' }, h('div', { class: 'row between' }, h('h2', {}, '專案事實欄位'),
      h('div', { class: 'row' },
        h('label', { class: 'row' }, h('input', { type: 'checkbox', checked: S.ui.showMaster, onchange: (e) => { S.ui.showMaster = e.target.checked; render(); } }), '顯示母本設定欄'),
        h('button', { class: 'btn', onclick: fillFromMaster, title: '還原測試用：把母本值當作新案值' }, '以母本資料填入'),
        h('button', { class: 'btn', onclick: () => { if (confirm('清空所有新案值？')) { Object.values(S.facts).forEach((f) => { f.newValue = ''; f.source = ''; }); S.confirmed = false; save(); render(); } } }, '清空新案值'),
        h('button', { class: 'btn', onclick: exportFacts }, '匯出專案資料確認表 .xlsx'),
        pickBtn('匯入確認表', '.xlsx', importFacts))));
    if (S.ui.showMaster) card.appendChild(h('div', { class: 'note' }, '母本設定：「母本值」用於還原測試；「母本出現寫法」是母本中實際出現的文字（一行一種），產出時會被替換成「替換格式」，{v} 代表新案值。例如母本寫法「地上4層」、格式「地上{v}層」。'));
    const groups = [...new Set(RG.FIELD_DEFS.map((d) => d.group))];
    for (const g of groups) {
      const defs = RG.FIELD_DEFS.filter((d) => d.group === g);
      card.appendChild(h('div', { class: 'fact-group' }, h('h3', {}, g),
        h('table', { class: 'data facts' }, h('tbody', {}, defs.map((d) => factRow(d))))));
    }
    wrap.appendChild(card);

    // 條件標籤
    wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, '條件標籤 ', h('span', { class: 'muted' }, '綠色＝由專案事實推導；藍色＝手動加入；紅色刪除線＝手動排除。點一下切換。')), tagPanel(ti)));

    // 確認
    const missing = RG.FIELD_DEFS.filter((d) => !String(S.facts[d.label].newValue || '').trim()).length;
    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: missing ? 'note warn' : 'note ok' }, missing ? `尚有 ${missing} 個欄位空白，產出時會標成【待填：欄位名稱】。` : '所有欄位都已填寫。'),
      h('label', { class: 'switch' }, h('input', { type: 'checkbox', checked: S.confirmed, onchange: (e) => { S.confirmed = e.target.checked; save(); render(); } }),
        h('span', {}, h('b', {}, '我已確認專案資料與條件標籤'), h('span', { class: 'muted' }, '未經確認不可產出報告（規劃書第 6.2 節）。修改任何欄位後需要重新確認。')))));
    return wrap;
  }
  function factRow(d) {
    const f = S.facts[d.label];
    const sg = S.suggestions[d.label] || [];
    const setF = (k, v) => {
      f[k] = v;
      if (k === 'newValue') f.source = f.source && f.source.startsWith('手動') ? f.source : v ? '手動輸入' : '';
      if (S.confirmed) {
        S.confirmed = false;
        toast('專案資料已修改，請重新確認');
      }
      invalidateOutputs();
      save();
    };
    const input = d.multiline
      ? h('textarea', { rows: 2, value: f.newValue, placeholder: d.hint || '', onchange: (e) => { setF('newValue', e.target.value); render(); } })
      : h('input', { type: 'text', value: f.newValue, placeholder: d.hint || '', onchange: (e) => { setF('newValue', e.target.value); render(); } });
    const cells = [h('td', { class: 'lbl' }, d.label),
      h('td', { class: String(f.newValue || '').trim() ? '' : 'missing' }, input,
        f.source ? h('div', { class: 'src' }, '來源：' + f.source) : null,
        sg.filter((s) => s.value !== f.newValue).slice(0, 3).map((s) => h('div', { class: 'sugg' }, '建議：', h('b', {}, short(s.value, 60)), h('span', { class: 'muted' }, s.src),
          h('button', { class: 'btn small', onclick: () => { f.newValue = s.value; f.source = s.src + '：' + short(s.snippet, 60); S.confirmed = false; invalidateOutputs(); save(); render(); } }, '採用'))))];
    if (S.ui.showMaster) {
      cells.push(h('td', { style: { width: '150px' } }, h('div', { class: 'muted' }, '母本值'), h('input', { type: 'text', value: f.masterValue, onchange: (e) => setF('masterValue', e.target.value) })));
      cells.push(h('td', { style: { width: '210px' } }, h('div', { class: 'muted' }, '母本出現寫法（一行一種）'), h('textarea', { rows: 2, value: f.masterLiterals, onchange: (e) => setF('masterLiterals', e.target.value) })));
      cells.push(h('td', { style: { width: '120px' } }, h('div', { class: 'muted' }, '替換格式'), h('input', { type: 'text', value: f.format, onchange: (e) => setF('format', e.target.value || '{v}') })));
    }
    return h('tr', {}, cells);
  }
  function tagPanel(ti) {
    const all = new Set(Object.values(RG.TAG_GROUPS).flat());
    if (S.master) S.master.blocks.forEach((b) => K.effective(b, S.overrides).tags.forEach((t) => all.add(t)));
    ti.derived.tags.forEach((t) => all.add(t));
    const grouped = Object.assign({}, RG.TAG_GROUPS);
    const known = new Set(Object.values(RG.TAG_GROUPS).flat());
    const other = [...all].filter((t) => !known.has(t));
    if (other.length) grouped['其他'] = other;
    const derived = new Set(ti.derived.tags);
    const chip = (t) => {
      const man = S.manualTags[t];
      const cls = derived.has(t) ? (man === 'off' ? 'off' : 'derived') : man === 'on' ? 'on' : '';
      return h('button', { class: 'chip ' + cls, title: derived.has(t) ? '推導自 ' + ti.derived.why[t] : '', onclick: () => {
        if (derived.has(t)) S.manualTags[t] = man === 'off' ? undefined : 'off';
        else S.manualTags[t] = man === 'on' ? undefined : 'on';
        if (!S.manualTags[t]) delete S.manualTags[t];
        S.confirmed = false;
        invalidateOutputs();
        save();
        render();
      } }, t);
    };
    return h('div', {}, Object.entries(grouped).map(([g, tags]) => h('div', { class: 'tag-group' }, h('div', { class: 'g' }, g), h('div', { class: 'chips' }, tags.map(chip)))),
      h('div', { class: 'muted', style: { marginTop: '6px' } }, '目前生效：', [...ti.set].join('、') || '（無）'));
  }
  async function addCaseFiles(files) {
    await busy('讀取新案資料…', async () => {
      for (const file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        const item = { name: file.name, text: '' };
        try {
          if (ext === 'docx') {
            const { doc } = await D.load(await file.arrayBuffer());
            item.text = D.toPlainText(doc);
          } else if (['xlsx', 'xls', 'csv'].includes(ext)) {
            const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
            item.text = wb.SheetNames.map((n) => `【工作表 ${n}】\n` + XLSX.utils.sheet_to_csv(wb.Sheets[n], { FS: '\t', blankrows: false })).join('\n');
          } else if (ext === 'txt') item.text = await file.text();
          else item.error = 'PDF 目前無法在瀏覽器版讀取，請另存為 Word 或手動填寫';
        } catch (e) {
          item.error = '讀取失敗：' + e.message;
        }
        S.caseFiles = S.caseFiles.filter((f) => f.name !== item.name).concat(item);
      }
      runSuggest();
      const n = Object.values(S.suggestions).filter((a) => a.length).length;
      toast(`已讀取 ${files.length} 個檔案，找到 ${n} 個欄位的建議值`, 'ok');
      render();
    });
  }
  const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  function runSuggest() {
    S.suggestions = {};
    for (const d of RG.FIELD_DEFS) {
      const out = [];
      for (const file of S.caseFiles) {
        if (!file.text) continue;
        for (const label of d.labels) {
          const re = new RegExp('(?:^|[\\n\\t\\s　])' + escRe(label) + '\\s*[:：\\t]\\s*([^\\n\\t]{1,120})', 'g');
          let m;
          while ((m = re.exec(file.text))) {
            const v = m[1].trim().replace(/[。；;，,]+$/, '');
            if (!v || out.some((x) => x.value === v)) continue;
            const line = file.text.slice(0, m.index).split('\n').length;
            out.push({ value: v, src: `${file.name} 第${line}行`, snippet: m[0].trim() });
            if (out.length >= 3) break;
          }
          if (out.length >= 3) break;
        }
      }
      if (out.length) S.suggestions[d.label] = out;
    }
  }
  function acceptAllSuggestions() {
    let n = 0;
    for (const [label, arr] of Object.entries(S.suggestions)) {
      const f = S.facts[label];
      if (!String(f.newValue || '').trim() && arr.length) {
        f.newValue = arr[0].value;
        f.source = arr[0].src + '：' + short(arr[0].snippet, 60);
        n++;
      }
    }
    S.confirmed = false;
    invalidateOutputs();
    save();
    toast(`已採用 ${n} 個建議值，請逐一核對`, 'ok');
    render();
  }
  function fillFromMaster() {
    if (!confirm('把每個欄位的「母本值」填入新案值？（用於還原測試；母本值空白的欄位會維持空白）')) return;
    for (const f of Object.values(S.facts)) {
      f.newValue = f.masterValue;
      f.source = f.masterValue ? '母本值' : '';
    }
    S.confirmed = false;
    invalidateOutputs();
    save();
    render();
  }
  function exportFacts() {
    const rows = [['群組', '欄位', '新案值（擷取值）', '來源', '使用者修正', '母本值', '母本出現寫法', '替換格式']];
    for (const d of RG.FIELD_DEFS) {
      const f = S.facts[d.label];
      rows.push([d.group, d.label, f.newValue, f.source, '', f.masterValue, f.masterLiterals, f.format]);
    }
    const ti = tagInfo();
    const tagRows = [['條件標籤', '狀態', '推導依據']];
    const all = new Set([...Object.values(RG.TAG_GROUPS).flat(), ...ti.set]);
    for (const t of all) tagRows.push([t, ti.set.has(t) ? '適用' : '不適用', ti.derived.why[t] || (S.manualTags[t] ? '手動' : '')]);
    xlsxDownload(`${S.projectName}_專案資料確認表.xlsx`, [{ name: '專案資料', cols: [10, 16, 40, 36, 30, 24, 24, 14], rows }, { name: '條件標籤', cols: [18, 10, 30], rows: tagRows }]);
  }
  async function importFacts(file) {
    await busy('匯入專案資料確認表…', async () => {
      const rows = await readXlsxRows(file, '專案資料');
      let n = 0;
      for (const r of rows) {
        const f = S.facts[String(r['欄位']).trim()];
        if (!f) continue;
        const fix = String(r['使用者修正'] || '').trim();
        const v = fix || String(r['新案值（擷取值）'] || '').trim();
        if (v !== f.newValue) {
          f.newValue = v;
          f.source = fix ? '使用者修正（確認表）' : String(r['來源'] || '');
          n++;
        }
        if (r['母本值'] !== undefined) f.masterValue = String(r['母本值']);
        if (r['母本出現寫法'] !== undefined) f.masterLiterals = String(r['母本出現寫法']);
        if (r['替換格式']) f.format = String(r['替換格式']);
      }
      S.confirmed = false;
      invalidateOutputs();
      save();
      toast(`已匯入，${n} 個欄位有變更，請重新確認`, 'ok');
      render();
    });
  }

  /* ====================================================================
   * 4. 工項規劃
   * ==================================================================== */
  function viewPlan() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '4 工項規劃'),
      h('p', { class: 'view-desc' }, '依專案的條件標籤，決定每個條件式內容（工項、子表格、評選表、風險條目）要保留或刪除。表名、「分項工程」標題會跟著所屬表格一起處理。可以逐筆改成「保留」或「刪除」。'));
    if (!S.master) return wrap.appendChild(needMaster()), wrap;
    const res = planRes();
    const f = S.ui.plan;
    const ti = tagInfo();
    const relevant = S.master.blocks.filter((b) => {
      const e = K.effective(b, S.overrides);
      return (e.cat === 'C' && e.tags.length) || b.kind === 'rows' || b.kind === 'figure' || !res[b.id].include || S.planOverrides[b.id];
    });
    const q = f.q.trim();
    const list = relevant.filter((b) => {
      const r = res[b.id];
      if (f.filter === 'keep' && !r.include) return false;
      if (f.filter === 'drop' && r.include) return false;
      if (f.filter === 'user' && !S.planOverrides[b.id]) return false;
      if (q && !(b.text.includes(q) || ('表' + (b.capNo || '')) === q)) return false;
      return true;
    });
    const nDrop = S.master.blocks.filter((b) => !res[b.id].include && !['empty', 'other'].includes(b.kind)).length;
    const pages = Math.max(1, Math.ceil(list.length / PAGE));
    f.page = Math.min(f.page, pages - 1);
    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: 'row' }, h('b', {}, '目前條件標籤：'), [...ti.set].map((t) => h('span', { class: 'badge tag' }, t)), h('button', { class: 'btn small', onclick: () => go('project') }, '修改')),
      h('div', { class: 'row', style: { marginTop: '10px' } },
        h('span', { class: 'badge keep' }, `條件式內容 ${relevant.length} 筆`), h('span', { class: 'badge drop' }, `將刪除 ${nDrop} 個區塊`), h('span', { class: 'badge user' }, `使用者指定 ${Object.keys(S.planOverrides).length} 筆`)),
      h('div', { class: 'row', style: { marginTop: '10px' } },
        h('select', { onchange: (e) => { f.filter = e.target.value; f.page = 0; render(); } }, [['all', '全部'], ['keep', '只看保留'], ['drop', '只看刪除'], ['user', '只看使用者指定']].map(([v, l]) => h('option', { value: v, selected: f.filter === v }, l))),
        h('input', { type: 'search', placeholder: '搜尋內容或表號', value: f.q, onchange: (e) => { f.q = e.target.value; f.page = 0; render(); } }),
        h('button', { class: 'btn', onclick: exportPlan }, '匯出工項規劃表 .xlsx'),
        pickBtn('匯入工項規劃表', '.xlsx', importPlan),
        Object.keys(S.planOverrides).length ? h('button', { class: 'btn ghost', onclick: () => { S.planOverrides = {}; invalidateOutputs(); save(); render(); } }, '清除所有指定') : null)));
    const rows = list.slice(f.page * PAGE, (f.page + 1) * PAGE).map((b) => {
      const r = res[b.id];
      const e = K.effective(b, S.overrides);
      const po = S.planOverrides[b.id] || '';
      return h('tr', { class: po ? 'modified' : '' },
        h('td', {}, b.id, h('div', { class: 'reason' }, posOf(b))),
        h('td', {}, h('span', { class: 'badge' }, kindLabel(b))),
        h('td', {}, h('div', { class: 'clamp', onclick: (ev) => ev.currentTarget.classList.toggle('open') }, blockText(b)), b.kind === 'figure' ? thumb(b) : null,
          b.kind === 'figure' && r.include ? h('div', { class: 'reason' }, e.cat === 'F' && S.options.figurePlaceholder ? '產出時換成【待補圖】' : '產出時沿用原圖') : null),
        h('td', {}, e.tags.map((t) => h('span', { class: 'badge tag', style: { opacity: ti.set.has(t) ? 1 : 0.55 } }, t))),
        h('td', {}, h('span', { class: 'badge ' + (r.include ? 'keep' : 'drop') }, r.include ? '保留' : '刪除'), h('div', { class: 'reason' }, r.reason)),
        h('td', {}, h('select', { onchange: (ev) => { if (ev.target.value) S.planOverrides[b.id] = ev.target.value; else delete S.planOverrides[b.id]; if (RG.mem.memorable(b)) remember([b]); invalidateOutputs(); save(); render(); } },
          [['', '自動'], ['keep', '保留'], ['drop', '刪除']].map(([v, l]) => h('option', { value: v, selected: po === v }, l)))));
    });
    wrap.appendChild(h('div', { class: 'card' }, h('div', { class: 'tbl-wrap' }, h('table', { class: 'data' }, h('thead', {}, h('tr', {}, ['區塊', '類型', '內容', '條件標籤', '判定', '決定'].map((x) => h('th', {}, x)))), h('tbody', {}, rows))), pager(f, pages, list.length)));
    return wrap;
  }
  function exportPlan() {
    const res = planRes();
    const rows = [['區塊', '章', '表號', '類型', '內容（前 200 字）', '條件標籤', '自動判定', '理由', '使用者決定（保留／刪除）']];
    for (const b of S.master.blocks) {
      const e = K.effective(b, S.overrides);
      if (!((e.cat === 'C' && e.tags.length) || b.kind === 'rows' || b.kind === 'figure' || !res[b.id].include || S.planOverrides[b.id])) continue;
      const po = S.planOverrides[b.id];
      rows.push([b.id, b.chapter, b.capNo ? '表' + b.capNo : '', K.KIND_LABEL[b.kind], short(b.kind === 'rows' ? b.title + ' ' + b.text : b.text, 200), e.tags.join('、'), res[b.id].include ? '保留' : '刪除', res[b.id].reason, po === 'keep' ? '保留' : po === 'drop' ? '刪除' : '']);
    }
    xlsxDownload(`${S.projectName}_工項規劃表.xlsx`, [{ name: '工項規劃', cols: [9, 8, 8, 8, 60, 24, 9, 28, 14], rows }]);
  }
  async function importPlan(file) {
    await busy('匯入工項規劃表…', async () => {
      const rows = await readXlsxRows(file, '工項規劃');
      const byIdPlan = Object.fromEntries(S.master.blocks.map((b) => [b.id, b]));
      const touchedPlan = [];
      let n = 0;
      for (const r of rows) {
        const id = String(r['區塊']).trim();
        const v = String(r['使用者決定（保留／刪除）'] || '').trim();
        if (byIdPlan[id] && RG.mem.memorable(byIdPlan[id])) touchedPlan.push(byIdPlan[id]);
        if (v === '保留') S.planOverrides[id] = 'keep';
        else if (v === '刪除') S.planOverrides[id] = 'drop';
        else {
          delete S.planOverrides[id];
          continue;
        }
        n++;
      }
      if (touchedPlan.length) RG.mem.recordMany(S.memory, touchedPlan, memCtx());
      saveMemory();
      invalidateOutputs();
      save();
      toast(`已匯入 ${n} 筆使用者決定`, 'ok');
      render();
    });
  }

  /* ====================================================================
   * 5. 產出
   * ==================================================================== */
  const TYPE_COLORS = { 沿用母本: '#8a94a8', 新案資料填入: '#1f5fbf', 待填: '#c62828', 待補圖: '#b5523b', 待審: '#2b7bb9', 自動更正: '#0f7f7a', 母本疑義: '#e08a00', 已刪除: '#5a3d8a', 自動計算風險值: '#2e9b5f' };
  async function doGenerate() {
    await busy('產出報告中…', async () => {
      const res = planRes();
      const out = await RG.gen.generate(S.master.buffer, { blocks: S.master.blocks, overrides: S.overrides, facts: S.facts, rules: S.rules, options: S.options, planRes: res, tagSet: tagInfo().set });
      const loaded = await D.load(out.data);
      out.blocks = D.parseBlocks(loaded.doc, loaded.styles).concat(await D.headerFooterBlocks(loaded.zip));
      out.name = `${S.projectName}_報告書.docx`;
      S.lastOutput = out;
      toast('報告產出完成，請下載報告書與差異清單', 'ok');
      render();
    });
  }
  async function doRestore() {
    await busy('執行還原測試…', async () => {
      const facts = RG.clone(S.facts);
      Object.values(facts).forEach((f) => (f.newValue = f.masterValue));
      const res = K.plan(S.master.blocks, S.overrides, new Set(), {}, { keepAll: true });
      const out = await RG.gen.generate(S.master.buffer, { blocks: S.master.blocks, overrides: S.overrides, facts, rules: S.rules, options: { clean: true, autoFix: true, recomputeRisk: false, updateFields: true, fillMissingF: false, markReview: false, figurePlaceholder: false }, planRes: res });
      const { doc, styles } = await D.load(out.data);
      const blocks = D.parseBlocks(doc, styles);
      S.restore = Object.assign(RG.gen.compareRestore(S.master.blocks, blocks, S.rules), { data: out.data, uncoveredF: out.stats.uncoveredF });
      render();
    });
  }
  function exportDiff() {
    const o = S.lastOutput;
    const rows = [['編號', '章節', '表號／段落', '來源類型', '來源依據', '母本原文', '新案文字', '判定理由', '需人工確認', '確認結果']];
    for (const d of o.diffs) rows.push([d.no, d.chapter + (d.section ? ' ' + d.section : ''), d.loc, d.type, d.src, d.before, d.after, d.reason, d.needCheck, '']);
    const st = diffStats(o.diffs);
    const statRows = [['章', ...st.types, '合計'], ...st.chapters.map((c) => [c, ...st.types.map((t) => st.m[c][t] || 0), st.types.reduce((n, t) => n + (st.m[c][t] || 0), 0)])];
    statRows.push([], ['說明：可選取上表後在 Excel 插入「堆疊直條圖」。']);
    xlsxDownload(`${S.projectName}_差異清單.xlsx`, [{ name: '差異清單', cols: [6, 18, 26, 12, 12, 50, 50, 36, 10, 12], rows }, { name: '統計', cols: [12, ...st.types.map(() => 12), 8], rows: statRows }]);
  }
  function diffStats(diffs) {
    const types = Object.keys(TYPE_COLORS).filter((t) => diffs.some((d) => d.type === t));
    const chapters = [...new Set(diffs.map((d) => d.chapter))];
    const m = {};
    for (const c of chapters) m[c] = {};
    for (const d of diffs) m[d.chapter][d.type] = (m[d.chapter][d.type] || 0) + 1;
    return { types, chapters, m };
  }
  function viewGenerate() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '5 產出報告'),
      h('p', { class: 'view-desc' }, '以母本 .docx 為樣板複製，保留樣式、頁首頁尾、分節、目錄與表格格式；替換專案事實、刪除不適用內容，並同步產出差異清單。開啟產出的 Word 時，請選「是」更新功能變數，目錄、表號、圖號會自動重排。'));
    if (!S.master) return wrap.appendChild(needMaster()), wrap;
    const opt = (k, label, desc) => h('label', { class: 'switch' }, h('input', { type: 'checkbox', checked: S.options[k], onchange: (e) => { S.options[k] = e.target.checked; invalidateOutputs(); save(); } }), h('span', {}, h('b', {}, label), h('span', { class: 'muted' }, desc)));
    wrap.appendChild(h('div', { class: 'grid2' },
      h('div', { class: 'card' }, h('h2', {}, '產出選項'),
        opt('clean', '乾淨版（--clean）', '不加母本疑義的橘底、待審的淺藍底與註解；【待填】紅字仍保留，以免遺漏'),
        opt('fillMissingF', 'F 類找不到新案值時整段／整格標【待填】', '關閉時 F 類只替換有設定的欄位，其餘沿用母本原文（可能殘留母本數字）'),
        opt('figurePlaceholder', 'F 類圖片與表格截圖換成【待補圖】佔位框', '母本的現場照片、位置圖、配置圖等不帶入新案；圖名保留，請再放入新案的圖。T 類與條件符合的 C 類圖片沿用原圖'),
        opt('markReview', 'M 類句型需改寫時標「待審」', '例如新案地下層數為 0、句中提到新案沒有的工項；以淺藍底＋Word 註解標示，並列入差異清單'),
        opt('autoFix', '自動更正已知錯字', '依「規則設定」的自動更正清單，例如「生方向指示燈」→「逃生方向指示燈」'),
        opt('recomputeRisk', '自動計算風險值', '依風險矩陣（可能性 × 嚴重度）重算風險值與 H／M／L 等級'),
        opt('updateFields', '開檔時更新功能變數', '在 settings.xml 設定 updateFields，Word 開檔會詢問是否更新目錄')),
      h('div', { class: 'card' }, h('h2', {}, '產出前檢查'),
        h('ul', { class: 'checklist' },
          h('li', {}, h('span', { class: 'ic-ok' }, '✓'), h('span', {}, '母本：' + S.master.name)),
          h('li', {}, h('span', { class: S.confirmed ? 'ic-ok' : 'ic-bad' }, S.confirmed ? '✓' : '!'), h('span', {}, S.confirmed ? '專案資料已確認' : '專案資料尚未確認（請到「3 專案資料」勾選已確認）')),
          h('li', {}, h('span', { class: 'ic-ok' }, '•'), h('span', {}, '條件標籤：' + ([...tagInfo().set].join('、') || '（無）')))),
        h('div', { class: 'row', style: { marginTop: '12px' } },
          h('button', { class: 'btn primary', disabled: !S.confirmed, onclick: doGenerate }, '產出新案報告'),
          h('button', { class: 'btn', onclick: doRestore, title: '以母本自己的資料產出，比對逐字相同比例（驗收測試）' }, '執行還原測試')))));

    if (S.restore) {
      const r = S.restore;
      const pass = r.ratio >= 0.98 && r.tablesMaster === r.tablesOut && r.chaptersMaster === r.chaptersOut;
      wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, '還原測試結果'),
        h('div', { class: 'note ' + (pass ? 'ok' : 'bad') }, pass ? '通過：正文段落與表格儲存格 ≥ 98% 逐字相同，表格數與章數相同。' : '未通過，請檢查下列差異。'),
        r.uncoveredF ? h('div', { class: 'note warn' }, `有 ${r.uncoveredF} 段／格 F 專案事實沒有對應到任何欄位的母本寫法。還原測試中先沿用原文；正式產出時會標成【待填】。請到「3 專案資料」補齊母本設定，或把這些內容改分類為 T。`) : null,
        h('div', { class: 'kpis' },
          h('div', { class: 'kpi' }, h('div', { class: 'k' }, '逐字相同'), h('div', { class: 'v' }, (r.ratio * 100).toFixed(2) + '%'), h('div', { class: 's' }, `${r.same} / ${r.total} 單位`)),
          h('div', { class: 'kpi' }, h('div', { class: 'k' }, '表格數'), h('div', { class: 'v' }, `${r.tablesOut} / ${r.tablesMaster}`)),
          h('div', { class: 'kpi' }, h('div', { class: 'k' }, '章數'), h('div', { class: 'v' }, `${r.chaptersOut} / ${r.chaptersMaster}`))),
        h('div', { class: 'row', style: { marginTop: '10px' } }, h('button', { class: 'btn', onclick: () => download('還原測試.docx', r.data) }, '下載還原測試 .docx')),
        r.diffs.length ? h('div', { class: 'tbl-wrap', style: { marginTop: '10px', maxHeight: '320px' } }, h('table', { class: 'data' }, h('thead', {}, h('tr', {}, h('th', {}, '位置'), h('th', {}, '母本（自動更正後）'), h('th', {}, '產出'))), h('tbody', {}, r.diffs.slice(0, 100).map((d) => h('tr', {}, h('td', {}, d.key), h('td', {}, short(d.master, 200)), h('td', {}, short(d.out, 200))))))) : null));
    }

    const o = S.lastOutput;
    if (o) {
      const st = diffStats(o.diffs);
      const max = Math.max(1, ...st.chapters.map((c) => st.types.reduce((n, t) => n + (st.m[c][t] || 0), 0)));
      const count = (t) => o.diffs.filter((d) => d.type === t).length;
      const shown = o.diffs.filter((d) => (S.ui.diffType ? d.type === S.ui.diffType : d.type !== '沿用母本'));
      wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, '產出結果'),
        h('div', { class: 'row' },
          h('button', { class: 'btn primary', onclick: () => download(o.name, o.data, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') }, '下載報告書 .docx'),
          h('button', { class: 'btn', onclick: exportDiff }, '下載差異清單 .xlsx'),
          h('button', { class: 'btn', onclick: () => { runCheck(o.blocks, '產出檔：' + o.name, false); go('check'); } }, '對產出檔執行一致性檢查')),
        h('div', { class: 'kpis', style: { marginTop: '12px' } },
          ...['沿用母本', '新案資料填入', '待填', '待補圖', '待審', '自動更正', '母本疑義', '已刪除', '自動計算風險值'].map((t) => h('div', { class: 'kpi' + (['待填', '待補圖', '待審'].includes(t) && count(t) ? ' warn' : '') }, h('div', { class: 'k' }, t), h('div', { class: 'v' }, count(t))))),
        h('h2', { style: { marginTop: '16px' } }, '各章來源統計'),
        h('div', { class: 'bars' }, st.chapters.map((c) => {
          const tot = st.types.reduce((n, t) => n + (st.m[c][t] || 0), 0);
          return h('div', { class: 'bar-row' }, h('span', {}, c), h('div', { class: 'bar', style: { width: (tot / max) * 100 + '%' } }, st.types.map((t) => (st.m[c][t] ? h('span', { title: `${t}：${st.m[c][t]}`, style: { width: (st.m[c][t] / tot) * 100 + '%', background: TYPE_COLORS[t] } }) : null))), h('span', { class: 'muted' }, tot));
        })),
        h('div', { class: 'legend' }, st.types.map((t) => h('span', {}, h('i', { style: { background: TYPE_COLORS[t] } }), t)))));
      wrap.appendChild(h('div', { class: 'card' }, h('div', { class: 'row between' }, h('h2', {}, '差異清單預覽 ', h('span', { class: 'muted' }, `${shown.length} 筆（最多顯示 300 筆）`)),
        h('select', { onchange: (e) => { S.ui.diffType = e.target.value; render(); } }, [['', '全部（不含沿用母本）'], ...Object.keys(TYPE_COLORS).map((t) => [t, t])].map(([v, l]) => h('option', { value: v, selected: S.ui.diffType === v }, l)))),
        h('div', { class: 'tbl-wrap' }, h('table', { class: 'data' }, h('thead', {}, h('tr', {}, ['#', '章節', '位置', '來源類型', '母本原文', '新案文字', '判定理由'].map((x) => h('th', {}, x)))),
          h('tbody', {}, shown.slice(0, 300).map((d) => h('tr', {}, h('td', {}, d.no), h('td', {}, d.chapter), h('td', {}, d.loc), h('td', {}, h('span', { class: 'badge', style: { background: TYPE_COLORS[d.type], color: '#fff' } }, d.type)), h('td', {}, h('div', { class: 'clamp', onclick: (ev) => ev.currentTarget.classList.toggle('open') }, d.before)), h('td', {}, h('div', { class: 'clamp', onclick: (ev) => ev.currentTarget.classList.toggle('open') }, d.after)), h('td', { class: 'reason' }, d.reason))))))));
    }
    return wrap;
  }

  /* ====================================================================
   * 6. 一致性檢查
   * ==================================================================== */
  function runCheck(blocks, label, isMaster) {
    const issues = RG.check.run(blocks, { rules: S.rules, facts: S.facts, tagSet: tagInfo().set, isMaster });
    S.lastCheck = { label, issues, at: new Date().toLocaleString('zh-TW') };
    S.ui.checkItem = '';
  }
  async function checkOtherFile(file) {
    await busy('檢查中…', async () => {
      const { zip, doc, styles } = await D.load(await file.arrayBuffer());
      runCheck(D.parseBlocks(doc, styles).concat(await D.headerFooterBlocks(zip)), file.name, false);
      render();
    });
  }
  function viewCheck() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '6 一致性檢查'),
      h('p', { class: 'view-desc' }, '檢查跨章節同一事實不一致、表號與表名引用、日期、數值合理性、風險值、母本殘留名詞、已知錯字與待填項目。頁碼欄目前留空（需轉 PDF 比對，於 Python 版提供）。'));
    wrap.appendChild(h('div', { class: 'card' }, h('div', { class: 'row' },
      h('button', { class: 'btn', disabled: !S.master, onclick: () => busy('檢查母本…', async () => { runCheck(S.master.blocks.concat(S.master.hfBlocks || []), '母本：' + S.master.name, true); render(); }) }, '檢查母本'),
      h('button', { class: 'btn primary', disabled: !S.lastOutput, onclick: () => busy('檢查產出檔…', async () => { runCheck(S.lastOutput.blocks, '產出檔：' + S.lastOutput.name, false); render(); }) }, '檢查最近產出的報告'),
      pickBtn('選擇其他 .docx 檢查', '.docx', checkOtherFile))));
    const c = S.lastCheck;
    if (!c) return wrap;
    const sum = RG.check.summary(c.issues);
    const shown = c.issues.filter((i) => !S.ui.checkItem || i.item.split('：')[0] === S.ui.checkItem);
    wrap.appendChild(h('div', { class: 'card' }, h('div', { class: 'row between' }, h('h2', {}, '檢查結果 ', h('span', { class: 'muted' }, `${c.label} · ${c.at}`)),
      h('button', { class: 'btn', onclick: () => xlsxDownload('一致性檢查報告.xlsx', [{ name: '一致性檢查', cols: [20, 34, 40, 30, 40, 40, 8], rows: [['項目', '位置A', 'A的寫法', '位置B', 'B的寫法', '問題說明', '頁碼'], ...c.issues.map((i) => [i.item, i.locA, i.textA, i.locB, i.textB, i.desc, i.page])] }]) }, '匯出一致性檢查報告 .xlsx')),
      c.issues.length ? h('div', { class: 'chips', style: { margin: '8px 0 12px' } },
        h('button', { class: 'chip ' + (S.ui.checkItem ? '' : 'on'), onclick: () => { S.ui.checkItem = ''; render(); } }, `全部 ${c.issues.length}`),
        Object.entries(sum).map(([k, n]) => h('button', { class: 'chip ' + (S.ui.checkItem === k ? 'on' : ''), onclick: () => { S.ui.checkItem = k; render(); } }, `${k} ${n}`))) : h('div', { class: 'note ok' }, '沒有發現問題。'),
      shown.length ? h('div', { class: 'tbl-wrap' }, h('table', { class: 'data' }, h('thead', {}, h('tr', {}, ['項目', '位置A', 'A的寫法', '位置B', 'B的寫法', '問題說明'].map((x) => h('th', {}, x)))),
        h('tbody', {}, shown.slice(0, 500).map((i) => h('tr', {}, h('td', {}, i.item), h('td', { class: 'reason' }, i.locA), h('td', {}, i.textA), h('td', { class: 'reason' }, i.locB), h('td', {}, i.textB), h('td', {}, i.desc)))))) : null));
    return wrap;
  }


  /* ====================================================================
   * 記憶庫管理
   * ==================================================================== */
  const PLAN_LABEL = { keep: '保留', drop: '刪除', '': '' };
  function currentKeys() {
    if (!S.master) return new Set();
    return new Set(S.master.blocks.filter((b) => RG.mem.memorable(b)).map((b) => RG.mem.keyOf(b)));
  }
  /* 記憶修改後，讓目前母本立即反映 */
  function refreshAfterMemoryEdit(deletedKeys) {
    if (S.master) {
      if (deletedKeys && deletedKeys.size) {
        for (const b of S.master.blocks) {
          if (RG.mem.memorable(b) && deletedKeys.has(RG.mem.keyOf(b))) {
            delete S.overrides[b.id];
            delete S.planOverrides[b.id];
            b.mem = { state: 'new' };
          }
        }
      }
      const r = RG.mem.apply(S.memory, S.master.blocks, { overrides: S.overrides, planOverrides: S.planOverrides, sig: S.masterSig, keepExisting: false });
      S.memResult = Object.assign(r, { name: S.master.name, prevName: '', sameMaster: true });
      invalidateOutputs();
      save();
    }
    saveMemory();
  }
  const MEM_CHECKS = {
    conflict: { label: '衝突', desc: '同一段文字在同一份母本中被標成不同分類或標籤', test: (e) => e.conflict },
    cNoTag: { label: 'C 類沒有標籤', desc: '條件式卻沒有標籤，產出時會一律保留', test: (e) => e.cat === 'C' && !(e.tags || []).length },
    tagNotC: { label: '有標籤但不是 C', desc: '只有 C 類會用到標籤，這些標籤不會有作用', test: (e) => e.cat !== 'C' && (e.tags || []).length > 0 },
    planOnT: { label: '有工項規劃決定但不是 C', desc: '非 C 類一律保留，除非指定刪除；請確認是否有意', test: (e) => e.cat !== 'C' && e.plan === 'drop' },
    notCurrent: { label: '目前母本沒有', desc: '記憶中有、但目前讀入的母本沒有這段文字', test: (e, ctx) => ctx.hasMaster && !ctx.cur.has(e._key) },
    inCurrent: { label: '目前母本有', desc: '目前讀入的母本中有這段文字', test: (e, ctx) => ctx.cur.has(e._key) },
  };
  function viewMemory() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '記憶庫'),
      h('p', { class: 'view-desc' }, '這裡列出所有記住的分類、標籤與工項規劃決定（依段落文字記錄，所有專案共用）。可以直接修改、刪除，或匯出 Excel 檢查後再匯入。修改會立即套用到目前讀入的母本。'));
    const f = S.ui.memv;
    const cur = currentKeys();
    const ctx = { cur, hasMaster: !!S.master };
    const all = Object.entries(S.memory.entries).map(([k, e]) => Object.assign({}, e, { _key: k }));
    const st = RG.mem.stats(S.memory);

    // ---- 統計 ----
    const catCount = { F: 0, T: 0, C: 0, M: 0 };
    const tagCount = {};
    const masterCount = {};
    for (const e of all) {
      catCount[e.cat] = (catCount[e.cat] || 0) + 1;
      (e.tags || []).forEach((t) => (tagCount[t] = (tagCount[t] || 0) + 1));
      (e.masters || []).forEach((m) => (masterCount[m] = (masterCount[m] || 0) + 1));
    }
    let bytes = 0;
    try {
      bytes = JSON.stringify(S.memory).length * 2;
    } catch (e) {
      /* ignore */
    }
    const last = all.reduce((m, e) => Math.max(m, e.updated || 0), 0);
    const kpi = (k, v, sub, cls = '') => h('div', { class: 'kpi ' + cls }, h('div', { class: 'k' }, k), h('div', { class: 'v' }, v), sub ? h('div', { class: 's' }, sub) : null);
    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: 'row between' }, h('h2', {}, '概況'),
        h('div', { class: 'row' },
          h('button', { class: 'btn', disabled: !all.length, onclick: exportMemoryXlsx }, '匯出 Excel'),
          pickBtn('匯入 Excel 修改', '.xlsx', importMemoryXlsx),
          h('button', { class: 'btn', disabled: !all.length, onclick: xlsxOrJson }, '匯出備份 .json'),
          pickBtn('匯入 .json（合併）', '.json', importMemory))),
      h('div', { class: 'kpis' },
        kpi('記憶筆數', st.count, `約 ${Math.max(1, Math.round(bytes / 1024))} KB／瀏覽器上限約 5,000 KB`),
        ...['F', 'T', 'C', 'M'].map((c) => kpi(`${c} ${K.CATS[c]}`, catCount[c] || 0)),
        kpi('來源母本', st.masters),
        kpi('衝突', st.conflicts, '', st.conflicts ? 'warn' : ''),
        kpi('最後修改', last ? new Date(last).toLocaleDateString('zh-TW') : '—')),
      Object.keys(tagCount).length ? h('div', { style: { marginTop: '10px' } }, h('span', { class: 'muted' }, '各標籤筆數：'),
        h('div', { class: 'chips', style: { marginTop: '4px' } }, Object.entries(tagCount).sort((a, b) => b[1] - a[1]).map(([t, n]) => h('button', { class: 'chip ' + (f.tag === t ? 'on' : ''), onclick: () => { f.tag = f.tag === t ? '' : t; f.page = 0; render(); } }, `${t} ${n}`)))) : null,
      Object.keys(masterCount).length ? h('div', { style: { marginTop: '8px' } }, h('span', { class: 'muted' }, '來源母本：'),
        h('div', { class: 'chips', style: { marginTop: '4px' } }, Object.entries(masterCount).map(([m, n]) => h('button', { class: 'chip ' + (f.master === m ? 'on' : ''), title: m, onclick: () => { f.master = f.master === m ? '' : m; f.page = 0; render(); } }, `${RG.mem.masterName(m)}${m === S.masterSig ? '（目前）' : ''} ${n}`)))) : null));

    if (!all.length) {
      wrap.appendChild(h('div', { class: 'card' }, h('div', { class: 'note' }, '記憶庫是空的。在「2 分類與標籤確認」確認或修改分類與標籤後，就會記到這裡；也可以匯入同事的記憶庫 .json。')));
      return wrap;
    }

    // ---- 完整性檢查 ----
    const checkRows = Object.entries(MEM_CHECKS).filter(([k]) => S.master || !['notCurrent', 'inCurrent'].includes(k)).map(([k, c]) => {
      const n = all.filter((e) => c.test(e, ctx)).length;
      const isInfo = k === 'inCurrent' || k === 'notCurrent';
      return h('li', {}, h('span', { class: n && !isInfo ? 'ic-warn' : 'ic-ok' }, n && !isInfo ? '!' : '✓'),
        h('span', {}, h('b', {}, c.label), ` ${n} 筆`, h('span', { class: 'muted' }, `　${c.desc}`), ' ',
          n ? h('button', { class: 'btn small', onclick: () => { f.check = k; f.page = 0; render(); } }, '只看這些') : null));
    });
    let coverage = null;
    if (S.master) {
      const memorable = S.master.blocks.filter((b) => RG.mem.memorable(b));
      const inMem = memorable.filter((b) => S.memory.entries[RG.mem.keyOf(b)]).length;
      coverage = h('div', { class: 'note ' + (inMem === memorable.length ? 'ok' : 'warn'), style: { marginTop: '10px' } },
        `目前母本「${S.master.name}」共 ${memorable.length} 個可記憶的段落／表格，其中 ${inMem} 個已在記憶庫（${memorable.length ? Math.round((inMem / memorable.length) * 100) : 0}%）。`,
        inMem < memorable.length ? h('span', {}, ' 其餘尚未確認，', h('button', { class: 'btn small', onclick: () => { S.ui.cls.mem = 'todo'; S.ui.cls.page = 0; go('classify'); } }, '到分類頁確認')) : null);
    }
    wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, '完整性檢查'), h('ul', { class: 'checklist' }, checkRows), coverage));

    // ---- 篩選與清單 ----
    const q = M_norm(f.q);
    const list = all.filter((e) => {
      if (f.cat && e.cat !== f.cat) return false;
      if (f.kind && e.kind !== f.kind) return false;
      if (f.tag && !(e.tags || []).includes(f.tag)) return false;
      if (f.master && !(e.masters || []).includes(f.master)) return false;
      if (f.check && !MEM_CHECKS[f.check].test(e, ctx)) return false;
      if (q && !(e.nh.includes(q) || (e.tags || []).join('、').includes(f.q.trim()))) return false;
      return true;
    }).sort((a, b) => (b.updated || 0) - (a.updated || 0));
    const pages = Math.max(1, Math.ceil(list.length / PAGE));
    f.page = Math.min(f.page, pages - 1);
    const shown = list.slice(f.page * PAGE, (f.page + 1) * PAGE);
    const selKeys = Object.keys(f.sel).filter((k) => f.sel[k] && S.memory.entries[k]);
    const upd = (k, v) => {
      f[k] = v;
      f.page = 0;
      render();
    };
    const sel = (k, opts) => h('select', { onchange: (e) => upd(k, e.target.value) }, opts.map(([v, l]) => h('option', { value: v, selected: f[k] === v }, l)));
    const anyFilter = f.q || f.cat || f.tag || f.master || f.kind || f.check;
    const bulkCat = h('select', {}, h('option', { value: '' }, '改分類為…'), Object.keys(K.CATS).map((c) => h('option', { value: c }, `${c} ${K.CATS[c]}`)));
    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: 'row' },
        h('input', { type: 'search', placeholder: '搜尋內容或標籤', value: f.q, style: { width: '220px' }, onchange: (e) => upd('q', e.target.value) }),
        sel('cat', [['', '全部分類'], ...Object.entries(K.CATS).map(([c, l]) => [c, `${c} ${l}`])]),
        sel('kind', [['', '全部類型'], ...RG.mem.KINDS.map((k) => [k, K.KIND_LABEL[k]])]),
        sel('check', [['', '不限檢查項目'], ...Object.entries(MEM_CHECKS).filter(([k]) => S.master || !['notCurrent', 'inCurrent'].includes(k)).map(([k, c]) => [k, c.label])]),
        f.tag ? h('span', { class: 'badge tag' }, '標籤：' + f.tag) : null,
        f.master ? h('span', { class: 'badge' }, '母本：' + RG.mem.masterName(f.master)) : null,
        anyFilter ? h('button', { class: 'btn small', onclick: () => { Object.assign(f, { q: '', cat: '', tag: '', master: '', kind: '', check: '', page: 0 }); render(); } }, '清除篩選') : null),
      h('div', { class: 'row', style: { marginTop: '10px' } },
        h('span', { class: 'muted' }, `已選 ${selKeys.length} 筆`),
        h('button', { class: 'btn small', onclick: () => { shown.forEach((e) => (f.sel[e._key] = true)); render(); } }, '選取本頁'),
        h('button', { class: 'btn small', onclick: () => { list.forEach((e) => (f.sel[e._key] = true)); render(); } }, `選取全部篩選結果（${list.length}）`),
        selKeys.length ? h('button', { class: 'btn small', onclick: () => { f.sel = {}; render(); } }, '取消選取') : null,
        bulkCat,
        h('button', { class: 'btn small', disabled: !selKeys.length, onclick: () => {
          if (!bulkCat.value) return toast('請先選擇要改成的分類');
          selKeys.forEach((k) => RG.mem.update(S.memory, k, { cat: bulkCat.value }));
          refreshAfterMemoryEdit();
          toast(`已將 ${selKeys.length} 筆改為 ${bulkCat.value}`, 'ok');
          render();
        } }, '套用'),
        h('button', { class: 'btn small', disabled: !selKeys.some((k) => S.memory.entries[k].conflict), onclick: () => {
          let n = 0;
          selKeys.forEach((k) => { if (S.memory.entries[k].conflict) { RG.mem.update(S.memory, k, {}); n++; } });
          refreshAfterMemoryEdit();
          toast(`已將 ${n} 筆衝突標為已解決`, 'ok');
          render();
        } }, '衝突標為已解決'),
        h('button', { class: 'btn small danger', disabled: !selKeys.length, onclick: () => {
          if (!confirm(`刪除選取的 ${selKeys.length} 筆記憶？目前母本中對應的段落會改回自動判斷。`)) return;
          const del = new Set(selKeys);
          RG.mem.remove(S.memory, selKeys);
          f.sel = {};
          refreshAfterMemoryEdit(del);
          toast(`已刪除 ${del.size} 筆記憶`, 'ok');
          render();
        } }, '刪除選取'))));

    const rows = shown.map((e) => {
      const k = e._key;
      const edit = (patch) => {
        RG.mem.update(S.memory, k, patch);
        refreshAfterMemoryEdit();
        render();
      };
      return h('tr', { class: e.conflict ? 'modified' : '' },
        h('td', {}, h('input', { type: 'checkbox', checked: !!f.sel[k], onchange: (ev) => { f.sel[k] = ev.target.checked; render(); } })),
        h('td', {}, h('div', { class: 'clamp', title: '點一下展開正規化後的全文（前 600 字）', onclick: (ev) => { const el = ev.currentTarget; el.classList.toggle('open'); el.textContent = el.classList.contains('open') ? e.nh.replace(/\|/g, '\n') : e.pv; } }, e.pv),
          h('div', { class: 'reason' }, K.KIND_LABEL[e.kind] || e.kind, `・${e.len} 字`, cur.has(k) ? '・目前母本有' : ''),
          e.conflict ? h('span', { class: 'badge user' }, '衝突') : null),
        h('td', {}, h('select', { onchange: (ev) => edit({ cat: ev.target.value }) }, Object.keys(K.CATS).map((c) => h('option', { value: c, selected: e.cat === c }, `${c} ${K.CATS[c]}`)))),
        h('td', {}, h('input', { type: 'text', value: (e.tags || []).join('、'), placeholder: e.cat === 'C' ? '例：擋土柱+斜撐' : '僅 C 類使用', style: { width: '170px', opacity: e.cat === 'C' ? 1 : 0.5 }, onchange: (ev) => edit({ tags: ev.target.value.split(/[、,，;；]+/).map((x) => x.trim()).filter(Boolean) }) })),
        h('td', {}, h('select', { onchange: (ev) => edit({ plan: ev.target.value }) }, [['', '自動'], ['keep', '保留'], ['drop', '刪除']].map(([v, l]) => h('option', { value: v, selected: (e.plan || '') === v }, l)))),
        h('td', { class: 'reason' }, (e.masters || []).map((m) => h('div', {}, RG.mem.masterName(m)))),
        h('td', { class: 'reason' }, e.updated ? new Date(e.updated).toLocaleString('zh-TW') : ''),
        h('td', {}, h('button', { class: 'btn small danger', onclick: () => {
          if (!confirm('刪除這筆記憶？')) return;
          RG.mem.remove(S.memory, [k]);
          delete f.sel[k];
          refreshAfterMemoryEdit(new Set([k]));
          render();
        } }, '刪除')));
    });
    wrap.appendChild(h('div', { class: 'card' },
      h('div', { class: 'tbl-wrap' }, h('table', { class: 'data' }, h('thead', {}, h('tr', {}, ['', '內容', '分類', '標籤', '工項規劃', '來源母本', '最後修改', ''].map((x) => h('th', {}, x)))), h('tbody', {}, rows))),
      pager(f, pages, list.length)));
    return wrap;
  }
  const M_norm = (s) => RG.mem.norm(String(s || '').trim());

  function exportMemoryXlsx() {
    const rows = [['記憶鍵', '類型', '內容（前 120 字）', '正規化全文（前 600 字）', '分類', '標籤', '工項規劃', '來源母本', '最後修改', '衝突', '刪除']];
    const all = Object.entries(S.memory.entries).sort((a, b) => (b[1].updated || 0) - (a[1].updated || 0));
    for (const [k, e] of all) rows.push([k, K.KIND_LABEL[e.kind] || e.kind, e.pv, e.nh.replace(/\|/g, '\n'), e.cat, (e.tags || []).join('、'), PLAN_LABEL[e.plan || ''], (e.masters || []).map(RG.mem.masterName).join('\n'), e.updated ? new Date(e.updated).toLocaleString('zh-TW') : '', e.conflict ? '是' : '', '']);
    xlsxDownload(`分類記憶庫_${new Date().toISOString().slice(0, 10)}.xlsx`, [
      { name: '記憶庫', cols: [26, 8, 50, 50, 6, 24, 9, 24, 18, 6, 6], rows },
      { name: '說明', cols: [90], rows: [['可修改「分類」（F／T／C／M）、「標籤」（以頓號分隔）、「工項規劃」（保留／刪除／空白）三欄；要刪除的列，在「刪除」欄填「刪除」。'], ['「記憶鍵」請勿修改，匯入時依此對應。其他欄位修改不會匯入。'], ['存檔後到工具的「記憶庫」頁按「匯入 Excel 修改」。修改過的列會視為已解決衝突。']] },
    ]);
  }
  async function importMemoryXlsx(file) {
    await busy('匯入記憶庫 Excel…', async () => {
      const rows = await readXlsxRows(file, '記憶庫');
      let changed = 0;
      let removed = 0;
      let skipped = 0;
      const bad = [];
      const del = new Set();
      const planMap = { 保留: 'keep', 刪除: 'drop', '': '' };
      for (const r of rows) {
        const k = String(r['記憶鍵'] || '').trim();
        const e = S.memory.entries[k];
        if (!e) {
          if (k) skipped++;
          continue;
        }
        if (String(r['刪除'] || '').trim() === '刪除') {
          del.add(k);
          continue;
        }
        const cat = String(r['分類'] || '').trim().toUpperCase();
        const tags = String(r['標籤'] || '').split(/[、,，;；]+/).map((x) => x.trim()).filter(Boolean);
        const planTxt = String(r['工項規劃'] || '').trim();
        if (!K.CATS[cat] || !(planTxt in planMap)) {
          bad.push(k);
          continue;
        }
        const plan = planMap[planTxt];
        if (cat !== e.cat || tags.join('、') !== (e.tags || []).join('、') || plan !== (e.plan || '')) {
          RG.mem.update(S.memory, k, { cat, tags, plan });
          changed++;
        }
      }
      removed = RG.mem.remove(S.memory, [...del]);
      refreshAfterMemoryEdit(del);
      toast(`已匯入：修改 ${changed} 筆、刪除 ${removed} 筆` + (skipped ? `、${skipped} 筆找不到記憶鍵已略過` : '') + (bad.length ? `、${bad.length} 筆分類或工項規劃填寫錯誤已略過` : ''), bad.length ? '' : 'ok');
      render();
    });
  }

  /* ====================================================================
   * 規則設定
   * ==================================================================== */
  function factLiterals() {
    const out = [];
    for (const f of Object.values(S.facts)) out.push(...RG.literalsOf(f));
    return out;
  }
  function runKeywordSuggest() {
    S.kwSuggest = K.suggestKeywords(S.memory, S.rules, factLiterals());
    S.kwDone = {};
  }
  const escRe2 = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  function applyKeyword(kind, tag, value) {
    const list = (S.rules.tagKeywords[tag] = (S.rules.tagKeywords[tag] || []).slice());
    if (kind === 'add') {
      const kw = escRe2(value);
      if (!list.includes(kw)) list.push(kw);
    } else {
      S.rules.tagKeywords[tag] = list.filter((k) => k !== value);
      if (!S.rules.tagKeywords[tag].length) delete S.rules.tagKeywords[tag];
    }
    S.kwDone[kind + '|' + tag + '|' + value] = kind === 'add' ? '已加入' : '已移除';
  }
  function afterKeywordChange(msg) {
    save();
    reclassifyKeepMemory();
    toast(msg + (S.master ? '；目前母本已重新初判，已確認的段落不受影響' : ''), 'ok');
    render();
  }
  function keywordCard() {
    const st = RG.mem.stats(S.memory);
    const card = h('div', { class: 'card' },
      h('div', { class: 'row between' },
        h('h2', {}, '關鍵字建議 ', h('span', { class: 'muted' }, `依記憶庫 ${st.count} 筆確認結果歸納`)),
        h('div', { class: 'row' },
          h('button', { class: 'btn primary', disabled: !st.count, onclick: () => { runKeywordSuggest(); render(); } }, S.kwSuggest ? '重新分析' : '分析記憶庫'),
          S.kwSuggest && S.kwSuggest.length ? h('button', { class: 'btn', onclick: () => {
            let n = 0;
            for (const r of S.kwSuggest) {
              for (const a of r.add) if (!S.kwDone['add|' + r.tag + '|' + a.term]) { applyKeyword('add', r.tag, a.term); n++; }
              for (const m of r.remove) if (!S.kwDone['remove|' + r.tag + '|' + m.kw]) { applyKeyword('remove', r.tag, m.kw); n++; }
            }
            afterKeywordChange(`已套用 ${n} 項建議`);
          } }, '全部採用') : null)),
      h('div', { class: 'muted' }, '工具會找出「你標了某個條件標籤、但目前關鍵字抓不到」的段落，歸納它們共同出現、又很少出現在其他段落的詞，建議加入該標籤的關鍵字；也會找出常誤判的關鍵字，建議移除。建議只有在你按「加入」或「移除」後才會生效。記憶庫越多、確認得越完整，建議越準。'));
    if (!st.count) card.appendChild(h('div', { class: 'note', style: { marginTop: '10px' } }, '記憶庫還是空的。請先在「2 分類與標籤確認」確認或修改分類與標籤。'));
    if (!S.kwSuggest) return card;
    if (!S.kwSuggest.length) {
      card.appendChild(h('div', { class: 'note ok', style: { marginTop: '10px' } }, '目前沒有建議：你標過的條件標籤，現有關鍵字都抓得到，也沒有常誤判的關鍵字。'));
      return card;
    }
    const doneBadge = (k) => (S.kwDone[k] ? h('span', { class: 'badge keep' }, S.kwDone[k]) : null);
    for (const r of S.kwSuggest) {
      const box = h('div', { style: { borderTop: '1px solid var(--line)', paddingTop: '10px', marginTop: '10px' } },
        h('div', {}, h('span', { class: 'badge tag' }, r.tag), ' ',
          h('span', { class: 'muted' }, `你標了 ${r.pos} 筆` + (r.missed ? `，其中 ${r.missed} 筆目前的關鍵字抓不到` : '') + `。目前關鍵字：${((S.rules.tagKeywords || {})[r.tag] || []).join('、') || '（無）'}`)));
      for (const a of r.add) {
        const k = 'add|' + r.tag + '|' + a.term;
        box.appendChild(h('div', { class: 'row', style: { marginTop: '6px' } },
          h('span', {}, '建議加入 ', h('b', {}, `「${a.term}」`)),
          h('span', { class: 'muted' }, `可多抓 ${a.gain} 筆` + (a.fp ? `，但會誤抓 ${a.fp} 筆沒有此標籤的段落` : '，不會誤抓')),
          h('span', { class: 'muted', title: '會多抓到的段落：\n' + a.examples.join('\n') + (a.fpExamples.length ? '\n\n可能誤抓：\n' + a.fpExamples.join('\n') : ''), style: { textDecoration: 'underline dotted', cursor: 'help' } }, '例子'),
          S.kwDone[k] ? doneBadge(k) : h('button', { class: 'btn small primary', onclick: () => { applyKeyword('add', r.tag, a.term); afterKeywordChange(`已把「${a.term}」加入「${r.tag}」的關鍵字`); } }, '加入'),
          S.kwDone[k] ? null : h('button', { class: 'btn small', onclick: () => { S.kwDone[k] = '已略過'; render(); } }, '略過')));
      }
      for (const m of r.remove) {
        const k = 'remove|' + r.tag + '|' + m.kw;
        box.appendChild(h('div', { class: 'row', style: { marginTop: '6px' } },
          h('span', {}, '建議移除 ', h('b', {}, `「${m.kw}」`)),
          h('span', { class: 'muted' }, `它抓到的 ${m.tp + m.fp} 筆 C 類段落中，有 ${m.fp} 筆你沒有標「${r.tag}」`),
          h('span', { class: 'muted', title: '誤判的段落：\n' + m.examples.join('\n'), style: { textDecoration: 'underline dotted', cursor: 'help' } }, '例子'),
          S.kwDone[k] ? doneBadge(k) : h('button', { class: 'btn small danger', onclick: () => { applyKeyword('remove', r.tag, m.kw); afterKeywordChange(`已從「${r.tag}」移除關鍵字「${m.kw}」`); } }, '移除'),
          S.kwDone[k] ? null : h('button', { class: 'btn small', onclick: () => { S.kwDone[k] = '已略過'; render(); } }, '略過')));
      }
      card.appendChild(box);
    }
    return card;
  }
  function viewRules() {
    const wrap = h('div', {}, h('h1', { class: 'view-title' }, '規則設定'),
      h('p', { class: 'view-desc' }, '以下規則對應規劃書中的 rules/*.yaml，以 JSON 格式編輯。修改後按「儲存」，會套用到分類初判、產出與檢查。「關鍵字」與「match」欄位是正規表示式。'));
    wrap.appendChild(keywordCard());
    const items = [
      ['tagKeywords', '條件標籤關鍵字（建庫初判用）', '格式：{"標籤": ["關鍵字1", "關鍵字2"]}'],
      ['tagRules', '專案事實 → 條件標籤 推導規則（tag_rules）', '格式：[{"field": "欄位", "match": "正規表示式", "tags": ["標籤"]}]；always: true 表示一律加入'],
      ['knownFixes', '自動更正清單（known_fixes）', '格式：[{"find": "正規表示式", "replace": "更正後", "note": "說明"}]'],
      ['knownIssues', '母本疑義標示', '格式：[{"find": "文字", "note": "疑義說明"}]，產出時加橘底與註解'],
      ['residueExtra', '母本殘留名詞', '格式：[{"term": "名詞", "unlessTag": "新案有此標籤時不檢查"}]；專案事實的母本寫法會自動加入'],
      ['extractors', '跨章節一致性擷取規則', '格式：[{"name": "名稱", "re": "正規表示式，第 1 個括號為值"}]'],
      ['templateSections', '視為通用範本的節名（正規表示式）', ''],
      ['templateCaptions', '視為通用範本的表名（正規表示式）', ''],
      ['factTableCaptions', '視為專案事實表的表名（正規表示式）', ''],
      ['figureFactCaptions', '視為 F 專案圖的圖名／表名（正規表示式）', '圖片或表格截圖的名稱符合時，產出換成【待補圖】；沒有名稱或都不符合時也預設為 F'],
      ['figureTemplateCaptions', '視為 T 通用圖的圖名（正規表示式）', '符合時沿用原圖；出現條件關鍵字的圖會判為 C'],
      ['factValueHeaders', 'F 專案事實表的「值」欄表頭（正規表示式）', '表頭符合的欄位，找不到新案值時整格標【待填】；都不符合時以最後一欄為值欄'],
      ['refTableCaption', '表5-1（依實際表格同步）的表名（正規表示式）', ''],
    ];
    for (const [key, title, hint] of items) {
      const isStr = typeof S.rules[key] === 'string';
      const ta = h('textarea', { class: 'code', rows: isStr ? 2 : 10, style: isStr ? { minHeight: '0' } : null, value: isStr ? S.rules[key] : JSON.stringify(S.rules[key], null, 2) });
      wrap.appendChild(h('div', { class: 'card' }, h('h2', {}, title), hint ? h('div', { class: 'muted', style: { marginBottom: '6px' } }, hint) : null, ta,
        h('div', { class: 'row', style: { marginTop: '8px' } },
          h('button', { class: 'btn primary', onclick: () => {
            try {
              S.rules[key] = isStr ? ta.value.trim() : JSON.parse(ta.value);
              save();
              invalidateOutputs();
              if (S.master && ['tagKeywords', 'templateSections', 'templateCaptions', 'factTableCaptions', 'figureFactCaptions', 'figureTemplateCaptions'].includes(key)) reclassifyKeepMemory();
              toast('已儲存：' + title, 'ok');
            } catch (e) {
              toast('JSON 格式錯誤：' + e.message, 'bad');
            }
          } }, '儲存'),
          h('button', { class: 'btn', onclick: () => { S.rules[key] = RG.clone(RG.DEFAULT_RULES[key]); save(); render(); toast('已還原預設'); } }, '還原預設'))));
    }
    return wrap;
  }

  /* ====================================================================
   * 使用說明
   * ==================================================================== */
  function viewHelp() {
    const step = (n, t, d) => h('li', {}, h('b', {}, `${n}. ${t}`), h('div', { class: 'muted' }, d));
    return h('div', {}, h('h1', { class: 'view-title' }, '使用說明'),
      h('div', { class: 'card' }, h('h2', {}, '操作流程'), h('ol', { class: 'checklist' },
        step(1, '母本解析', '選擇母本 .docx，檢查結構驗證結果（章、表、圖、分節、樣式）。'),
        step(2, '分類與標籤確認', '確認每段的 F／T／C／M 分類與條件標籤；可匯出 Excel 修改後再匯入。'),
        step(3, '專案資料', '讀入新案資料取得建議值，填寫或修正專案事實，確認條件標籤，最後勾選「已確認」。'),
        step(4, '工項規劃', '檢視依條件標籤自動決定的保留／刪除，必要時逐筆指定。'),
        step(5, '產出報告', '下載新案報告書 .docx 與差異清單 .xlsx。開啟 Word 時選「是」更新功能變數。'),
        step(6, '一致性檢查', '檢查產出檔是否有跨章節矛盾、母本殘留、待填項目等。'))),
      h('div', { class: 'card' }, h('h2', {}, '文件中的標示'), h('ul', { class: 'checklist' },
        h('li', {}, h('span', { style: { color: '#c00', fontWeight: 700 } }, '紅字【待填：欄位】'), '　新案資料沒有提供的專案事實'),
        h('li', {}, h('span', { style: { background: '#DDEBF7', padding: '0 4px' } }, '淺藍底＋註解「待審」'), '　M 類替換後句型可能不通，需人工改寫（乾淨版不標示）'),
        h('li', {}, h('span', { style: { background: '#FFC000', padding: '0 4px' } }, '橘底＋註解'), '　母本疑義（乾淨版不標示）'),
        h('li', {}, h('span', {}, '沿用母本、新案資料填入、自動更正'), '　文件中不標示，請看差異清單'))),
      h('div', { class: 'card' }, h('h2', {}, '這個網頁版目前的範圍'), h('ul', { class: 'checklist' },
        h('li', {}, h('span', { class: 'ic-ok' }, '✓'), '母本解析與結構驗證、F／T／C／M 分類與標籤（M1）'),
        h('li', {}, h('span', { class: 'ic-ok' }, '✓'), '以母本為樣板產出、還原測試（M2）'),
        h('li', {}, h('span', { class: 'ic-ok' }, '✓'), '一致性檢查（M3）'),
        h('li', {}, h('span', { class: 'ic-ok' }, '✓'), '新案資料讀取（docx／xlsx）與欄位建議、確認表（M4，不含 PDF／OCR／Claude 擷取）'),
        h('li', {}, h('span', { class: 'ic-ok' }, '✓'), '依條件標籤規劃工項、表5-1 同步刪列（M5）'),
        h('li', {}, h('span', { class: 'ic-ok' }, '✓'), '差異清單與統計（M8，Excel 圖表需手動插入）'),
        h('li', {}, h('span', { class: 'ic-ok' }, '✓'), '圖片與表格截圖分類（F 換成【待補圖】佔位框）、文字方塊與頁首頁尾的專案事實替換與檢查'),
        h('li', {}, h('span', { class: 'ic-warn' }, '…'), '用語資料庫匯入（M6）、Claude 起草（M7）、標準用語統一、DOCPROPERTY 綁定、頁碼對照、上傳新案圖片替換：尚未實作'))),
      h('div', { class: 'card' }, h('h2', {}, '資料保存'), h('p', { class: 'muted' }, '所有檔案都只在瀏覽器中處理，不會上傳。專案資料、分類確認與規則會自動存在這台電腦的瀏覽器中；換電腦或備份時，請用上方「儲存專案檔」。母本 .docx 不會被儲存，每次開啟工具需重新選擇。'),
        h('p', { class: 'muted' }, '記憶庫：你確認或修改過的分類、標籤與工項規劃，會依「段落文字」記在記憶庫中（與專案無關，所有專案共用）。換母本或母本改版時不會整份清除：文字完全相同的段落自動沿用，相似的標為「推定」，新的標為「新段落」，只需核對後兩類。左側「記憶庫」頁可以檢視全部記憶、完整性檢查、逐筆或批次修改與刪除，並可匯出 Excel 修改後匯入；也會一併存入專案檔。')));
  }

  /* ---------- 頂端列 ---------- */
  $('#projectName').value = S.projectName;
  $('#projectName').addEventListener('change', (e) => {
    S.projectName = e.target.value.trim() || '新案A';
    save();
  });
  $('#btnSaveProject').addEventListener('click', () => {
    const o = { app: 'report-tool', version: 1, savedAt: new Date().toISOString() };
    PERSIST.forEach((k) => (o[k] = S[k]));
    o.memory = S.memory;
    download(`${S.projectName}_專案檔.json`, JSON.stringify(o, null, 2), 'application/json');
  });
  $('#fileProject').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    try {
      const proj = JSON.parse(await f.text());
      applyLoaded(proj);
      if (proj.memory) {
        RG.mem.merge(S.memory, proj.memory);
        saveMemory();
      }
      $('#projectName').value = S.projectName;
      if (S.master) K.classify(S.master.blocks, S.rules, S.facts);
      invalidateOutputs();
      save();
      toast('已開啟專案檔', 'ok');
      render();
    } catch (err) {
      toast('專案檔格式錯誤：' + err.message, 'bad');
    }
  });
  $('#btnReset').addEventListener('click', () => {
    if (!confirm('清除所有專案資料、分類確認、工項規劃與規則設定，回到預設值？\n（記憶庫不會清除；要清除請到「2 分類與標籤確認」按「清除記憶庫」）')) return;
    S.facts = RG.newFacts();
    S.rules = RG.clone(RG.DEFAULT_RULES);
    S.overrides = {};
    S.planOverrides = {};
    S.manualTags = {};
    S.confirmed = false;
    if (S.master) K.classify(S.master.blocks, S.rules, S.facts);
    invalidateOutputs();
    save();
    render();
  });

  window.RG.app = { S, render, go, loadMaster, remember, runKeywordSuggest };
  render();
})();
