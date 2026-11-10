/* 產出新案報告（第 7 節）：以母本 docx 為樣板複製，替換專案事實、刪除不適用內容、自動更正、重算風險值 */
(function () {
  'use strict';
  const RG = (window.RG = window.RG || {});
  const D = RG.docx;
  const K = RG.kb;
  const G = (RG.gen = {});

  const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const short = (s, n = 300) => (s.length > n ? s.slice(0, n) + '…' : s);

  /* 建立替換器：專案事實、已知錯字、母本疑義 */
  G.buildMatcher = function (ctx) {
    const { facts, rules, options } = ctx;
    const factLits = [];
    for (const [label, f] of Object.entries(facts)) {
      const nv = String(f.newValue || '').trim();
      if (nv && nv === String(f.masterValue || '').trim()) continue; // 新案值與母本相同：沿用
      for (const lit of RG.literalsOf(f)) {
        factLits.push({
          lit,
          repl: nv ? (f.format || '{v}').split('{v}').join(nv) : `【待填：${label}】`,
          mark: nv ? null : 'todo',
          type: nv ? '新案資料填入' : '待填',
          info: label,
        });
      }
    }
    factLits.sort((a, b) => b.lit.length - a.lit.length);
    const factRe = factLits.length ? new RegExp(factLits.map((x) => escRe(x.lit)).join('|'), 'g') : null;
    const factMap = Object.fromEntries(factLits.map((x) => [x.lit, x]));
    const fixes = options.autoFix
      ? (rules.knownFixes || []).map((f) => ({ re: K.safeRe(f.find, 'g'), f })).filter((x) => x.re)
      : [];
    const issues = !options.clean
      ? (rules.knownIssues || []).filter((x) => x.find).map((x) => ({ re: new RegExp(escRe(x.find), 'g'), x }))
      : [];

    return function match(text, useFacts) {
      const cand = [];
      if (useFacts && factRe) {
        factRe.lastIndex = 0;
        let m;
        while ((m = factRe.exec(text))) {
          const x = factMap[m[0]];
          cand.push({ start: m.index, end: m.index + m[0].length, repl: x.repl, mark: x.mark, type: x.type, info: x.info, prio: 3 });
        }
      }
      for (const { re, f } of fixes) {
        re.lastIndex = 0;
        let m;
        while ((m = re.exec(text))) {
          if (!m[0]) {
            re.lastIndex++;
            continue;
          }
          cand.push({ start: m.index, end: m.index + m[0].length, repl: f.replace, type: '自動更正', info: `${m[0]}→${f.replace}（${f.note || ''}）`, prio: 2 });
        }
      }
      for (const { re, x } of issues) {
        re.lastIndex = 0;
        let m;
        while ((m = re.exec(text))) {
          cand.push({ start: m.index, end: m.index + m[0].length, repl: m[0], mark: 'issue', comment: '母本疑義：' + x.note, type: '母本疑義', info: x.note, prio: 1 });
        }
      }
      cand.sort((a, b) => a.start - b.start || b.prio - a.prio || b.end - b.start - (a.end - a.start));
      const out = [];
      let last = -1;
      for (const c of cand) if (c.start >= last) {
        out.push(c);
        last = c.end;
      }
      return out;
    };
  };

  const TYPE_ORDER = ['待填', '新案資料填入', '自動更正', '母本疑義'];

  /* buffer: 母本 ArrayBuffer；ctx: {blocks, overrides, facts, rules, options, planRes} */
  G.generate = async function (buffer, ctx) {
    const { blocks, overrides, planRes, rules, options } = ctx;
    const zip = await JSZip.loadAsync(buffer);
    const doc = D.parseXml(await zip.file('word/document.xml').async('string'));
    const kids = D.bodyChildren(doc);
    const comments = new D.CommentStore(zip, doc);
    await comments.init();
    const match = G.buildMatcher(ctx);
    const diffs = [];
    const removeEls = [];
    const removeRows = new Map(); // tbl -> Set(rowIdx)
    const removedCapNos = new Set();
    const byId = Object.fromEntries(blocks.map((b) => [b.id, b]));
    const where = (b) => (b.capNo ? '表' + b.capNo : K.KIND_LABEL[b.kind] || b.kind);
    const addDiff = (b, loc, type, src, before, after, reason, needCheck) =>
      diffs.push({ chapter: b.chapter, section: b.section || '', loc, type, src, before: short(before || ''), after: short(after || ''), reason, needCheck: needCheck ? '是' : '否' });

    const processParagraph = (b, p, cat, loc) => {
      const before = D.paraText(p);
      if (!before.trim()) return false;
      const ms = match(before, cat !== 'T');
      if (!ms.length) return false;
      D.replaceInParagraph(p, ms, (run, m) => {
        if (m.comment) comments.add(run, m.comment);
      });
      const after = D.paraText(p);
      const types = [...new Set(ms.map((m) => m.type))].sort((a, z) => TYPE_ORDER.indexOf(a) - TYPE_ORDER.indexOf(z));
      addDiff(
        b,
        loc,
        types[0],
        '母本 ' + b.id,
        before,
        after,
        ms.map((m) => `${m.type}：${m.info}`).join('；'),
        types.includes('待填') || types.includes('母本疑義') || types.includes('自動更正')
      );
      return true;
    };

    for (const b of blocks) {
      if (b.parentId) continue;
      const el = kids[b.idx];
      const r = planRes[b.id] || { include: true };
      if (!r.include) {
        removeEls.push(el);
        if (b.kind === 'caption' && b.capType === '表') removedCapNos.add(b.capNo);
        if (b.kind === 'table' && b.capNo) removedCapNos.add(b.capNo);
        if (b.kind !== 'empty' && b.kind !== 'other') addDiff(b, where(b), '已刪除', '母本 ' + b.id, b.text, '', r.reason || '不適用', false);
        continue;
      }
      const cat = K.effective(b, overrides).cat;
      if (el.localName === 'tbl') {
        const groupOfRow = {};
        for (const g of b.groups || []) {
          const gb = byId[g.id];
          const gr = planRes[g.id] || { include: true };
          for (let i = g.rowStart; i <= g.rowEnd; i++) groupOfRow[i] = { gb, gr };
          if (!gr.include) addDiff(gb, `${where(b)} ${g.title}（第${g.rowStart + 1}～${g.rowEnd + 1}列）`, '已刪除', '母本 ' + g.id, gb.text, '', gr.reason, false);
        }
        let changed = false;
        D.kids(el, 'tr').forEach((tr, ri) => {
          const go = groupOfRow[ri];
          if (go && !go.gr.include) {
            if (!removeRows.has(el)) removeRows.set(el, new Set());
            removeRows.get(el).add(ri);
            return;
          }
          const rowCat = go ? K.effective(go.gb, overrides).cat : cat;
          D.kids(tr, 'tc').forEach((tc, ci) => {
            D.all(tc, 'p').forEach((p) => {
              if (processParagraph(go ? go.gb : b, p, rowCat, `${where(b)} 第${ri + 1}列第${ci + 1}欄`)) changed = true;
            });
          });
        });
        if (!changed && !(b.groups || []).length) addDiff(b, where(b), '沿用母本', '母本 ' + b.id, b.text, b.text, '', false);
      } else {
        const ps = el.localName === 'p' ? [el] : D.all(el, 'p');
        let changed = false;
        ps.forEach((p) => {
          if (processParagraph(b, p, cat, where(b))) changed = true;
        });
        if (!changed && b.text.trim() && !['empty', 'other', 'toc'].includes(b.kind)) addDiff(b, where(b), '沿用母本', '母本 ' + b.id, b.text, b.text, '', false);
      }
    }

    // 表5-1：刪除只引用已刪除表格的列
    const refCap = K.safeRe(rules.refTableCaption || '(?!)');
    if (removedCapNos.size && refCap) {
      for (const b of blocks) {
        if (b.kind !== 'table' || !b.capName || !refCap.test(b.capName) || !(planRes[b.id] || {}).include) continue;
        const el = kids[b.idx];
        D.kids(el, 'tr').forEach((tr, ri) => {
          const t = D.kids(tr, 'tc').map(D.cellText).join(' ');
          const refs = [...t.matchAll(/表\s*([0-9０-９]+)\s*[-－–]\s*([0-9０-９]+)/g)].map((m) => D.toHalf(m[1]) + '-' + D.toHalf(m[2]));
          if (refs.length && refs.every((n) => removedCapNos.has(n))) {
            if (!removeRows.has(el)) removeRows.set(el, new Set());
            removeRows.get(el).add(ri);
            addDiff(b, `表${b.capNo} 第${ri + 1}列`, '已刪除', '母本 ' + b.id, t, '', '引用的表格已刪除：' + refs.map((n) => '表' + n).join('、'), true);
          }
        });
      }
    }

    // 風險值自動計算
    let riskFixed = 0;
    if (options.recomputeRisk) {
      for (const b of blocks) {
        if (b.kind !== 'table' || !(planRes[b.id] || {}).include) continue;
        const el = kids[b.idx];
        D.kids(el, 'tr').forEach((tr, ri) => {
          if (removeRows.has(el) && removeRows.get(el).has(ri)) return;
          const tcs = D.kids(tr, 'tc');
          for (const item of K.scanRiskRow(tcs.map(D.cellText))) {
            for (const f of item.fixes) {
              D.setCellText(tcs[f.col], f.to);
              riskFixed++;
              addDiff(b, `${where(b)} 第${ri + 1}列第${f.col + 1}欄`, '自動計算風險值', '依風險矩陣', f.from, f.to, `可能性 ${item.a} × 嚴重度 ${item.b} = ${item.value}（${item.level}）`, true);
            }
          }
        });
      }
    }

    // 刪除列：被刪範圍後的第一列若是垂直合併的延續格，改為起始格以免合併錯位
    for (const [tbl, set] of removeRows) {
      const trs = D.kids(tbl, 'tr');
      trs.forEach((tr, i) => {
        if (!set.has(i)) return;
        const next = trs[i + 1];
        if (next && !set.has(i + 1)) {
          for (const tc of D.kids(next, 'tc')) {
            const v = D.kid(D.kid(tc, 'tcPr'), 'vMerge');
            if (v && D.attr(v, 'val') !== 'restart') v.setAttributeNS(D.W, 'w:val', 'restart');
          }
        }
      });
      trs.forEach((tr, i) => set.has(i) && tbl.removeChild(tr));
      if (!D.kids(tbl, 'tr').length) removeEls.push(tbl);
    }
    // 刪除元素（含分節符號的段落只清空內容，保留分節）
    for (const el of removeEls) {
      if (!el.parentNode) continue;
      const pPr = el.localName === 'p' ? D.kid(el, 'pPr') : null;
      if (pPr && D.kid(pPr, 'sectPr')) {
        for (const c of Array.from(el.childNodes)) if (c !== pPr) el.removeChild(c);
      } else el.parentNode.removeChild(el);
    }

    zip.file('word/document.xml', D.serialize(doc));
    await comments.finalize();
    if (options.updateFields) await D.setUpdateFields(zip);
    const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    diffs.forEach((d, i) => (d.no = i + 1));
    return { data, diffs, riskFixed, removedCapNos: [...removedCapNos], commentCount: comments.items.length };
  };

  /* 還原測試：比對產出與母本（已知錯字自動更正後）的逐字相同比例 */
  G.compareRestore = function (masterBlocks, outBlocks, rules) {
    const fixes = (rules.knownFixes || []).map((f) => ({ re: K.safeRe(f.find, 'g'), f })).filter((x) => x.re);
    const fix = (s) => fixes.reduce((t, { re, f }) => t.replace(re, f.replace), s);
    const units = (blocks) => {
      const out = [];
      for (const b of blocks) {
        if (b.parentId || ['empty', 'other', 'toc'].includes(b.kind)) continue;
        if (b.kind === 'table') b.rows.forEach((r, ri) => r.forEach((c, ci) => out.push({ key: `${b.capNo ? '表' + b.capNo : b.id} 第${ri + 1}列第${ci + 1}欄`, text: c })));
        else out.push({ key: `${b.chapter} ${b.id}`, text: b.text });
      }
      return out;
    };
    const A = units(masterBlocks);
    const B = units(outBlocks);
    let same = 0;
    const diffs = [];
    const n = Math.max(A.length, B.length);
    for (let i = 0; i < n; i++) {
      const a = A[i] ? fix(A[i].text) : null;
      const b = B[i] ? B[i].text : null;
      if (a === b) same++;
      else if (diffs.length < 200) diffs.push({ key: (A[i] || B[i]).key, master: a, out: b });
    }
    const count = (bs, f) => bs.filter((b) => !b.parentId && f(b)).length;
    return {
      total: n,
      same,
      ratio: n ? same / n : 1,
      diffs,
      tablesMaster: count(masterBlocks, (b) => b.kind === 'table'),
      tablesOut: count(outBlocks, (b) => b.kind === 'table'),
      chaptersMaster: count(masterBlocks, (b) => b.kind === 'heading' && b.level === 1),
      chaptersOut: count(outBlocks, (b) => b.kind === 'heading' && b.level === 1),
    };
  };
})();
