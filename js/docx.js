/* docx 讀寫工具：解析母本結構、在段落中替換文字（保留格式）、標示顏色與註解 */
(function () {
  'use strict';
  const RG = (window.RG = window.RG || {});
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XMLNS = 'http://www.w3.org/XML/1998/namespace';
  const D = (RG.docx = { W });

  const isW = (n, name) => n && n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name);
  D.isW = isW;
  D.kids = (el, name) => {
    const out = [];
    for (let n = el.firstChild; n; n = n.nextSibling) if (isW(n, name)) out.push(n);
    return out;
  };
  D.kid = (el, name) => {
    if (!el) return null;
    for (let n = el.firstChild; n; n = n.nextSibling) if (isW(n, name)) return n;
    return null;
  };
  D.attr = (el, name) => (el ? el.getAttributeNS(W, name) || el.getAttribute('w:' + name) : null);
  D.all = (el, name) => Array.from(el.getElementsByTagNameNS(W, name));

  D.parseXml = function (xml) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('XML 解析失敗：' + err.textContent.slice(0, 200));
    return doc;
  };
  D.serialize = (doc) => {
    let s = new XMLSerializer().serializeToString(doc);
    if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
    return s;
  };

  D.load = async function (buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const f = zip.file('word/document.xml');
    if (!f) throw new Error('不是有效的 Word（.docx）檔案：找不到 word/document.xml');
    const doc = D.parseXml(await f.async('string'));
    const sf = zip.file('word/styles.xml');
    const styles = D.parseStyles(sf ? await sf.async('string') : '');
    return { zip, doc, styles };
  };

  D.parseStyles = function (xml) {
    const map = {};
    if (!xml) return map;
    const d = D.parseXml(xml);
    for (const s of d.getElementsByTagNameNS(W, 'style')) {
      const id = D.attr(s, 'styleId');
      const n = D.kid(s, 'name');
      map[id] = { name: n ? D.attr(n, 'val') : id, type: D.attr(s, 'type') };
    }
    return map;
  };

  D.body = (doc) => D.kid(doc.documentElement, 'body');
  /* 本文的頂層元素（不含最後的 sectPr）。解析與產出都用同一個函式，索引才會對得上。 */
  D.bodyChildren = (doc) => D.kids(D.body(doc)).filter((n) => n.localName !== 'sectPr');

  D.closestP = function (n) {
    for (let x = n.parentNode; x; x = x.parentNode) if (isW(x, 'p')) return x;
    return null;
  };
  D.runsOf = (p) => D.all(p, 'r').filter((r) => D.closestP(r) === p);

  /* 段落的文字片段（每個 w:t / tab / br 一段），用來對應字元位置與 XML 節點 */
  D.segs = function (p) {
    const out = [];
    let pos = 0;
    for (const r of D.runsOf(p)) {
      for (let c = r.firstChild; c; c = c.nextSibling) {
        if (!isW(c)) continue;
        const l = c.localName;
        let txt = null;
        if (l === 't') txt = c.textContent;
        else if (l === 'tab') txt = '\t';
        else if (l === 'br' || l === 'cr') txt = '\n';
        else if (l === 'noBreakHyphen') txt = '-';
        if (txt === null) continue;
        out.push({ r, c, start: pos, end: pos + txt.length, text: txt });
        pos += txt.length;
      }
    }
    return out;
  };
  D.paraText = (p) => D.segs(p).map((s) => s.text).join('');
  D.cellText = (tc) => D.all(tc, 'p').map(D.paraText).join('\n');
  D.cellSpan = (tc) => {
    const g = D.kid(D.kid(tc, 'tcPr'), 'gridSpan');
    return g ? parseInt(D.attr(g, 'val'), 10) || 1 : 1;
  };
  D.vMerge = (tc) => {
    const v = D.kid(D.kid(tc, 'tcPr'), 'vMerge');
    if (!v) return null;
    return D.attr(v, 'val') === 'restart' ? 'restart' : 'continue';
  };

  D.setT = function (t, s) {
    t.textContent = s;
    t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
  };

  /* 把含多個內容的 run 拆成一個 run 一個內容，方便精準替換 */
  D.atomize = function (p) {
    for (const r of D.runsOf(p)) {
      const content = [];
      for (let c = r.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && !isW(c, 'rPr')) content.push(c);
      if (content.length <= 1) continue;
      const rPr = D.kid(r, 'rPr');
      let after = r;
      for (let i = 1; i < content.length; i++) {
        const nr = r.cloneNode(false);
        if (rPr) nr.appendChild(rPr.cloneNode(true));
        nr.appendChild(content[i]);
        after.parentNode.insertBefore(nr, after.nextSibling);
        after = nr;
      }
    }
  };

  D.splitRunAt = function (seg, offset) {
    const r = seg.r;
    const txt = seg.c.textContent;
    const r2 = r.cloneNode(true);
    r.parentNode.insertBefore(r2, r.nextSibling);
    D.setT(seg.c, txt.slice(0, offset));
    D.setT(D.kid(r2, 't'), txt.slice(offset));
  };

  D.makeRun = function (doc, rPr, text) {
    const r = doc.createElementNS(W, 'w:r');
    if (rPr) r.appendChild(rPr.cloneNode(true));
    const parts = String(text).split(/(\n|\t)/);
    for (const part of parts) {
      if (part === '') continue;
      if (part === '\n') r.appendChild(doc.createElementNS(W, 'w:br'));
      else if (part === '\t') r.appendChild(doc.createElementNS(W, 'w:tab'));
      else {
        const t = doc.createElementNS(W, 'w:t');
        D.setT(t, part);
        r.appendChild(t);
      }
    }
    return r;
  };

  /* 在段落中依字元位置替換。matches: [{start,end,repl,mark,comment}]，不可重疊。
   * onRun(run, match) 讓呼叫者加註解。 */
  D.replaceInParagraph = function (p, matches, onRun) {
    if (!matches.length) return 0;
    D.atomize(p);
    const split = (pos) => {
      const s = D.segs(p).find((x) => x.start < pos && pos < x.end);
      if (s && s.c.localName === 't') D.splitRunAt(s, pos - s.start);
    };
    let n = 0;
    for (const m of matches.slice().sort((a, b) => b.start - a.start)) {
      split(m.end);
      split(m.start);
      const inRange = D.segs(p).filter((s) => s.start >= m.start && s.end <= m.end && s.end > s.start);
      if (!inRange.length) continue;
      const first = inRange[0].r;
      let nr = null;
      if (m.repl !== '') {
        nr = D.makeRun(p.ownerDocument, D.kid(first, 'rPr'), m.repl);
        first.parentNode.insertBefore(nr, first);
      }
      new Set(inRange.map((s) => s.r)).forEach((r) => r.parentNode && r.parentNode.removeChild(r));
      if (nr && m.mark) D.markRun(nr, m.mark);
      if (nr && onRun) onRun(nr, m);
      n++;
    }
    return n;
  };

  D.setCellText = function (tc, text) {
    const ps = D.kids(tc, 'p');
    if (!ps.length) return;
    const p = ps[0];
    const runs = D.runsOf(p);
    const rPr = runs.length ? D.kid(runs[0], 'rPr') : null;
    runs.forEach((r) => r.parentNode.removeChild(r));
    p.appendChild(D.makeRun(p.ownerDocument, rPr, text));
    for (const q of ps.slice(1)) D.runsOf(q).forEach((r) => r.parentNode.removeChild(r));
  };

  /* ---------- 字元格式（rPr 子元素需依 schema 順序） ---------- */
  const RPR_ORDER = 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath'.split(' ');
  D.setRprProp = function (r, name, attrs) {
    const doc = r.ownerDocument;
    let rPr = D.kid(r, 'rPr');
    if (!rPr) {
      rPr = doc.createElementNS(W, 'w:rPr');
      r.insertBefore(rPr, r.firstChild);
    }
    const old = D.kid(rPr, name);
    if (old) rPr.removeChild(old);
    const el = doc.createElementNS(W, 'w:' + name);
    for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(W, 'w:' + k, v);
    const idx = RPR_ORDER.indexOf(name);
    let before = null;
    for (let c = rPr.firstChild; c; c = c.nextSibling) {
      if (isW(c) && RPR_ORDER.indexOf(c.localName) > idx) {
        before = c;
        break;
      }
    }
    rPr.insertBefore(el, before);
  };
  D.MARKS = {
    todo: { color: 'FF0000' }, // 待填：紅字
    issue: { shd: 'FFC000' }, // 母本疑義：橘底
    library: { shd: 'C6EFCE' }, // 取自資料庫：淺綠底
    draft: { highlight: 'yellow' }, // Claude 起草：黃底
  };
  D.markRun = function (r, mark) {
    const m = D.MARKS[mark];
    if (!m) return;
    if (m.color) D.setRprProp(r, 'color', { val: m.color });
    if (m.highlight) D.setRprProp(r, 'highlight', { val: m.highlight });
    if (m.shd) D.setRprProp(r, 'shd', { val: 'clear', color: 'auto', fill: m.shd });
  };

  /* ---------- Word 註解 ---------- */
  D.CommentStore = class {
    constructor(zip, docDom) {
      this.zip = zip;
      this.docDom = docDom;
      this.items = [];
      this.nextId = 0;
    }
    async init() {
      const f = this.zip.file('word/comments.xml');
      if (f) {
        this.cdoc = D.parseXml(await f.async('string'));
        for (const c of this.cdoc.getElementsByTagNameNS(W, 'comment')) this.nextId = Math.max(this.nextId, (parseInt(D.attr(c, 'id'), 10) || 0) + 1);
      }
      for (const tag of ['commentRangeStart', 'commentReference']) {
        for (const c of this.docDom.getElementsByTagNameNS(W, tag)) this.nextId = Math.max(this.nextId, (parseInt(D.attr(c, 'id'), 10) || 0) + 1);
      }
    }
    add(run, text) {
      const doc = this.docDom;
      const id = String(this.nextId++);
      const s = doc.createElementNS(W, 'w:commentRangeStart');
      s.setAttributeNS(W, 'w:id', id);
      const e = doc.createElementNS(W, 'w:commentRangeEnd');
      e.setAttributeNS(W, 'w:id', id);
      const refRun = doc.createElementNS(W, 'w:r');
      const ref = doc.createElementNS(W, 'w:commentReference');
      ref.setAttributeNS(W, 'w:id', id);
      refRun.appendChild(ref);
      run.parentNode.insertBefore(s, run);
      run.parentNode.insertBefore(e, run.nextSibling);
      e.parentNode.insertBefore(refRun, e.nextSibling);
      this.items.push({ id, text });
    }
    async finalize() {
      if (!this.items.length) return;
      if (!this.cdoc) {
        this.cdoc = D.parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W}"/>`);
        const relPath = 'word/_rels/document.xml.rels';
        const rels = D.parseXml(await this.zip.file(relPath).async('string'));
        const RNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
        const root = rels.documentElement;
        let max = 0;
        for (const r of root.getElementsByTagNameNS(RNS, 'Relationship')) max = Math.max(max, parseInt((r.getAttribute('Id') || '').replace(/\D/g, ''), 10) || 0);
        const rel = rels.createElementNS(RNS, 'Relationship');
        rel.setAttribute('Id', 'rId' + (max + 1000));
        rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments');
        rel.setAttribute('Target', 'comments.xml');
        root.appendChild(rel);
        this.zip.file(relPath, D.serialize(rels));
        const ctPath = '[Content_Types].xml';
        const ct = D.parseXml(await this.zip.file(ctPath).async('string'));
        const CNS = 'http://schemas.openxmlformats.org/package/2006/content-types';
        const ov = ct.createElementNS(CNS, 'Override');
        ov.setAttribute('PartName', '/word/comments.xml');
        ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
        ct.documentElement.appendChild(ov);
        this.zip.file(ctPath, D.serialize(ct));
      }
      const root = this.cdoc.documentElement;
      const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
      for (const it of this.items) {
        const c = this.cdoc.createElementNS(W, 'w:comment');
        c.setAttributeNS(W, 'w:id', it.id);
        c.setAttributeNS(W, 'w:author', '報告產生工具');
        c.setAttributeNS(W, 'w:initials', 'RG');
        c.setAttributeNS(W, 'w:date', now);
        const p = this.cdoc.createElementNS(W, 'w:p');
        p.appendChild(D.makeRun(this.cdoc, null, it.text));
        c.appendChild(p);
        root.appendChild(c);
      }
      this.zip.file('word/comments.xml', D.serialize(this.cdoc));
    }
  };

  /* settings.xml 加上 updateFields，開檔時詢問是否更新目錄與功能變數 */
  const SETTINGS_AFTER = 'hdrShapeDefaults footnotePr endnotePr compat docVars rsids mathPr attachedSchema themeFontLang clrSchemeMapping doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade captions readModeInkLockDown smartTagType schemaLibrary shapeDefaults doNotEmbedSmartTags decimalSymbol listSeparator'.split(' ');
  D.setUpdateFields = async function (zip) {
    const f = zip.file('word/settings.xml');
    if (!f) return false;
    const s = D.parseXml(await f.async('string'));
    const root = s.documentElement;
    let el = D.kid(root, 'updateFields');
    if (!el) {
      el = s.createElementNS(W, 'w:updateFields');
      let before = null;
      for (let c = root.firstChild; c; c = c.nextSibling) {
        if (isW(c) && SETTINGS_AFTER.includes(c.localName)) {
          before = c;
          break;
        }
      }
      root.insertBefore(el, before);
    }
    el.setAttributeNS(W, 'w:val', 'true');
    zip.file('word/settings.xml', D.serialize(s));
    return true;
  };

  /* ---------- 母本結構解析 ---------- */
  const toHalf = (s) => String(s).replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
  D.toHalf = toHalf;
  const CH_RE = /^第\s*[一二三四五六七八九十]+\s*章/;
  const SEC_RE = /^第\s*[一二三四五六七八九十]+\s*節/;
  const CAP_RE = /^(表|圖)\s*([0-9０-９]+)\s*[-－–.．]\s*([0-9０-９]+)\s*[、：:．.\s]*(.*)$/;
  const TOC_STYLE = /toc|目錄|table of figures|圖表目錄|index/i;

  D.parseCaption = function (t) {
    const m = t.replace(/\s+/g, ' ').trim().match(CAP_RE);
    if (!m) return null;
    return { capType: m[1], capNo: toHalf(m[2]) + '-' + toHalf(m[3]), capName: m[4].trim() };
  };

  D.parseBlocks = function (doc, styles) {
    const kids = D.bodyChildren(doc);
    const blocks = [];
    let chapter = '前置';
    let section = '';
    kids.forEach((el, idx) => {
      const b = { id: 'b' + idx, idx, chapter, section, kind: 'other', text: '', style: '' };
      if (el.localName === 'p') {
        const pPr = D.kid(el, 'pPr');
        const ps = D.kid(pPr, 'pStyle');
        const sid = ps ? D.attr(ps, 'val') : '';
        const sname = (styles[sid] && styles[sid].name) || sid || '';
        b.style = sname;
        b.hasSectPr = !!D.kid(pPr, 'sectPr');
        const text = D.paraText(el);
        b.text = text;
        const t = text.replace(/\t+\d*\s*$/, '').trim();
        const isToc = TOC_STYLE.test(sname) || /\t\s*[0-9ivxlcIVXLC]+\s*$/.test(text);
        const hasDrawing = el.getElementsByTagNameNS(W, 'drawing').length + el.getElementsByTagNameNS(W, 'pict').length > 0;
        const cap = !isToc && t.length < 90 ? D.parseCaption(t) : null;
        if (isToc) b.kind = 'toc';
        else if (sname === '章節' || (CH_RE.test(t) && t.length < 40)) {
          b.kind = 'heading';
          b.level = 1;
          const m = t.match(CH_RE);
          chapter = m ? m[0].replace(/\s+/g, '') : t.slice(0, 10);
          section = '';
          b.chapter = chapter;
          b.section = '';
          b.title = t;
        } else if (sname === '1-1' || (SEC_RE.test(t) && t.length < 40)) {
          b.kind = 'heading';
          b.level = 2;
          section = t.slice(0, 30);
          b.section = section;
          b.title = t;
        } else if (cap && (/表名|圖名|caption|標號/i.test(sname) || t.length < 50)) {
          b.kind = 'caption';
          Object.assign(b, cap);
        } else if (!t && hasDrawing) b.kind = 'figure';
        else if (!t) b.kind = 'empty';
        else b.kind = 'p';
      } else if (el.localName === 'tbl') {
        b.kind = 'table';
        const rows = D.kids(el, 'tr').map((tr) =>
          D.kids(tr, 'tc').map((tc) => ({ text: D.cellText(tc), span: D.cellSpan(tc), vm: D.vMerge(tc) }))
        );
        b.rows = rows.map((r) => r.map((c) => c.text));
        b.text = b.rows.map((r) => r.join(' | ')).join('\n');
        b.merged = rows.reduce((n, r) => n + r.filter((c) => c.span > 1 || c.vm === 'continue').length, 0);
        // 依「分項工程」列切分子表格（表4-13 等）
        const heads = [];
        b.rows.forEach((r, i) => {
          const first = (r[0] || '').trim();
          const all = r.join('\t').trim();
          if (/^分項工程/.test(first) || /^分項工程\s*[：:]/.test(all)) heads.push(i);
        });
        if (heads.length) {
          b.groups = heads.map((start, k) => {
            const end = k + 1 < heads.length ? heads[k + 1] - 1 : b.rows.length - 1;
            return { id: `${b.id}.g${k}`, rowStart: start, rowEnd: end, title: b.rows[start].filter(Boolean)[0] || '' };
          });
        }
      } else if (el.localName === 'sdt') {
        const ps = D.all(el, 'p');
        b.text = ps.map(D.paraText).join('\n');
        const instr = D.all(el, 'instrText').map((x) => x.textContent).join(' ');
        const gal = D.all(el, 'docPartGallery').map((x) => D.attr(x, 'val')).join(' ');
        b.kind = /TOC/.test(instr) || /Contents/i.test(gal) ? 'toc' : b.text.trim() ? 'sdt' : 'other';
      }
      blocks.push(b);
      if (b.groups) {
        for (const g of b.groups) {
          const rowsText = b.rows.slice(g.rowStart, g.rowEnd + 1).map((r) => r.join(' | ')).join('\n');
          blocks.push({
            id: g.id,
            idx,
            parentId: b.id,
            kind: 'rows',
            chapter,
            section,
            text: rowsText,
            title: g.title,
            rowStart: g.rowStart,
            rowEnd: g.rowEnd,
            tagText: b.rows.slice(g.rowStart, Math.min(g.rowEnd, g.rowStart + 2) + 1).map((r) => r.join(' ')).join('\n'),
          });
        }
      }
    });

    // 表名 ↔ 表格、圖名 ↔ 圖、「分項工程：X」段落 ↔ 表格
    const top = blocks.filter((b) => !b.parentId);
    const skip = (b) => b.kind === 'empty' || b.kind === 'other';
    top.forEach((b, i) => {
      if (b.kind === 'caption') {
        const want = b.capType === '表' ? 'table' : 'figure';
        const order = b.capType === '表' ? [1, 2, 3, -1, -2] : [-1, -2, -3, 1, 2];
        for (const d of order) {
          const step = d > 0 ? 1 : -1;
          let j = i,
            n = 0;
          let found = null;
          while (n < Math.abs(d)) {
            j += step;
            if (j < 0 || j >= top.length) break;
            if (skip(top[j])) continue;
            n++;
          }
          if (j >= 0 && j < top.length && top[j].kind === want && !top[j].captionId) found = top[j];
          if (found) {
            b.linkTo = found.id;
            found.captionId = b.id;
            found.capNo = b.capNo;
            found.capName = b.capName;
            break;
          }
        }
      } else if (b.kind === 'p' && /^\s*分項工程\s*[：:]/.test(b.text)) {
        for (let j = i + 1; j < Math.min(top.length, i + 4); j++) {
          if (skip(top[j])) continue;
          if (top[j].kind === 'table') b.linkTo = top[j].id;
          break;
        }
      }
    });
    for (const b of blocks) if (b.parentId) {
      const p = blocks.find((x) => x.id === b.parentId);
      b.capNo = p.capNo;
      b.capName = p.capName;
    }
    return blocks;
  };

  D.structureReport = function (doc, styles, blocks) {
    const fields = {};
    const instrs = D.all(doc, 'instrText').map((x) => x.textContent);
    D.all(doc, 'fldSimple').forEach((x) => instrs.push(D.attr(x, 'instr') || ''));
    for (const s of instrs) {
      const k = (s.trim().split(/\s+/)[0] || '').toUpperCase();
      if (k) fields[k] = (fields[k] || 0) + 1;
    }
    const chapters = [];
    for (const b of blocks) {
      if (b.kind === 'heading' && b.level === 1) chapters.push({ key: b.chapter, title: b.title, sections: [], blocks: 0 });
      else if (b.kind === 'heading' && b.level === 2 && chapters.length) chapters[chapters.length - 1].sections.push(b.title);
    }
    for (const c of chapters) c.blocks = blocks.filter((b) => b.chapter === c.key && !b.parentId).length;
    const top = blocks.filter((b) => !b.parentId);
    const tables = top.filter((b) => b.kind === 'table');
    const tableCaps = top.filter((b) => b.kind === 'caption' && b.capType === '表');
    const figCaps = top.filter((b) => b.kind === 'caption' && b.capType === '圖');
    const usedStyles = {};
    for (const b of top) if (b.style) usedStyles[b.style] = (usedStyles[b.style] || 0) + 1;
    const expectedStyles = ['章節', '1-1', 'Body Text', '表名', '圖名'].map((n) => ({
      name: n,
      defined: Object.values(styles).some((s) => s.name === n || s.name.toLowerCase() === n.toLowerCase()),
      used: usedStyles[n] || 0,
    }));
    const expectedCh = ['第一章', '第二章', '第三章', '第四章', '第五章', '第六章'];
    const missingCh = expectedCh.filter((c) => !chapters.some((x) => x.key === c));
    // 表號連續性
    const gaps = [];
    const byCh = {};
    for (const c of tableCaps) {
      const [a, n] = c.capNo.split('-').map(Number);
      (byCh[a] = byCh[a] || []).push(n);
    }
    for (const [a, arr] of Object.entries(byCh)) {
      const s = [...new Set(arr)].sort((x, y) => x - y);
      for (let i = 1; i <= s[s.length - 1]; i++) if (!s.includes(i)) gaps.push(`表${a}-${i}`);
    }
    const dupCaps = tableCaps.map((c) => c.capNo).filter((v, i, arr) => arr.indexOf(v) !== i);
    return {
      chapters,
      missingChapters: missingCh,
      tables: tables.map((t) => ({
        id: t.id,
        capNo: t.capNo || '',
        capName: t.capName || '',
        rows: t.rows.length,
        groups: t.groups ? t.groups.length : 0,
        merged: t.merged,
        chapter: t.chapter,
      })),
      tableCount: tables.length,
      tableCaptionCount: tableCaps.length,
      figureCaptionCount: figCaps.length,
      figureCount: top.filter((b) => b.kind === 'figure').length,
      sectionBreaks: doc.getElementsByTagNameNS(W, 'sectPr').length,
      fields,
      expectedStyles,
      captionsWithoutTable: tableCaps.filter((c) => !c.linkTo).map((c) => '表' + c.capNo),
      tablesWithoutCaption: tables.filter((t) => !t.captionId).length,
      numberGaps: gaps,
      duplicateCaptions: [...new Set(dupCaps)].map((n) => '表' + n),
      groupCount: tables.reduce((n, t) => n + (t.groups ? t.groups.length : 0), 0),
      paragraphCount: top.filter((b) => b.kind === 'p').length,
    };
  };

  /* 讀任意 docx 為純文字（新案資料用）：段落一行、表格每列以 Tab 分隔 */
  D.toPlainText = function (doc) {
    const lines = [];
    for (const el of D.bodyChildren(doc)) {
      if (el.localName === 'p') lines.push(D.paraText(el));
      else if (el.localName === 'tbl') {
        for (const tr of D.kids(el, 'tr')) lines.push(D.kids(tr, 'tc').map((tc) => D.cellText(tc).replace(/\n/g, ' ')).join('\t'));
      } else lines.push(D.all(el, 'p').map(D.paraText).join('\n'));
    }
    return lines.join('\n');
  };
})();
