/* 一致性檢查（第 10 節） */
(function () {
  'use strict';
  const RG = (window.RG = window.RG || {});
  const D = RG.docx;
  const K = RG.kb;
  const C = (RG.check = {});

  const CN = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
  const normVal = (v) => {
    let s = D.toHalf(String(v)).replace(/[,，\s]/g, '');
    if (/^[一二三四五六七八九十]+$/.test(s)) {
      if (s.length === 1) return String(CN[s]);
      if (s[0] === '十') return String(10 + (CN[s[1]] || 0));
      return String(CN[s[0]] * 10 + (CN[s[2]] || 0));
    }
    const n = parseFloat(s);
    return isNaN(n) ? s.toUpperCase() : String(n);
  };
  const loc = (b) => `${b.chapter}${b.section ? ' ' + b.section.slice(0, 12) : ''} ${b.capNo ? '表' + b.capNo : K.KIND_LABEL[b.kind] || ''}（${b.id}）`;
  const ctxOf = (text, i, len = 0) => text.slice(Math.max(0, i - 15), i + len + 25).replace(/\s+/g, ' ');

  /* 可檢查的文字單位：段落、表格儲存格 */
  const units = function* (blocks) {
    for (const b of blocks) {
      if (b.parentId || ['empty', 'other', 'toc'].includes(b.kind) || (b.kind === 'figure' && !b.placeholder)) continue;
      if (b.kind === 'table') {
        for (let ri = 0; ri < b.rows.length; ri++) for (let ci = 0; ci < b.rows[ri].length; ci++) yield { b, text: b.rows[ri][ci], where: `${loc(b)} 第${ri + 1}列第${ci + 1}欄` };
      } else yield { b, text: b.text, where: loc(b) };
    }
  };

  /* ctx: {rules, facts, tagSet, isMaster} */
  C.run = function (blocks, ctx) {
    const { rules, facts } = ctx;
    const tagSet = ctx.tagSet || new Set();
    const issues = [];
    const add = (item, locA, textA, locB, textB, desc) => issues.push({ item, locA, textA, locB: locB || '', textB: textB || '', desc, page: '' });
    const all = [...units(blocks)];

    // 1. 同一事實跨章節不一致
    for (const ex of rules.extractors || []) {
      const re = K.safeRe(ex.re, 'g');
      if (!re) continue;
      const seen = new Map();
      for (const u of all) {
        re.lastIndex = 0;
        let m;
        while ((m = re.exec(u.text))) {
          const v = normVal(m[1]);
          if (!seen.has(v)) seen.set(v, { u, snippet: ctxOf(u.text, m.index, m[0].length) });
        }
      }
      if (seen.size > 1) {
        const arr = [...seen.entries()];
        for (let i = 1; i < arr.length; i++) add(`同一事實不一致：${ex.name}`, arr[0][1].u.where, arr[0][1].snippet, arr[i][1].u.where, arr[i][1].snippet, `${ex.name}出現不同的值：${arr[0][0]} 與 ${arr[i][0]}`);
      }
    }

    // 2. 表號與表名引用
    const caps = {};
    for (const b of blocks) if (b.kind === 'caption' && b.capType === '表') caps[b.capNo] = b.capName;
    for (const u of all) {
      if (u.b.kind === 'caption') continue;
      for (const m of u.text.matchAll(/表\s*([0-9０-９]+)\s*[-－–]\s*([0-9０-９]+)\s*[「『（(]?([^\s，。、；;,:：)）」』]{0,30})/g)) {
        const no = D.toHalf(m[1]) + '-' + D.toHalf(m[2]);
        if (!(no in caps)) {
          add('表號引用錯誤', u.where, ctxOf(u.text, m.index, m[0].length), '', '', `引用的「表${no}」在文件中不存在`);
          continue;
        }
        const name = (m[3] || '').trim();
        const actual = caps[no] || '';
        if (name.length >= 4 && /表$/.test(name) && actual && !actual.includes(name) && !name.includes(actual)) {
          add('表名引用不符', u.where, `表${no} ${name}`, `表${no} 實際表名`, actual, '引用的表名與實際表名不同');
        }
      }
    }

    // 3. 日期
    const DATE_RE = /(\d{2,3})\s*年\s*(\d{1,2})\s*月(?:\s*(\d{1,2})\s*日)?/g;
    let cover = null;
    const rd = facts && facts['報告日期'] && (facts['報告日期'].newValue || facts['報告日期'].masterValue);
    if (rd) {
      const m = [...D.toHalf(rd).matchAll(DATE_RE)][0];
      if (m) cover = { y: +m[1], mo: +m[2], text: rd, where: '專案資料：報告日期' };
    }
    if (!cover) {
      for (const u of all) {
        if (u.b.chapter !== '前置') continue;
        const m = [...D.toHalf(u.text).matchAll(DATE_RE)][0];
        if (m) {
          cover = { y: +m[1], mo: +m[2], text: m[0], where: u.where };
          break;
        }
      }
    }
    const evalDates = new Map();
    for (const u of all) {
      if (!/評估日期|日期[：:]/.test(u.text)) continue;
      for (const m of D.toHalf(u.text).matchAll(DATE_RE)) {
        const y = +m[1];
        const mo = +m[2];
        if (cover && (y > cover.y || (y === cover.y && mo > cover.mo))) add('日期不合理', u.where, m[0], cover.where, cover.text, '評估日期晚於封面（報告）日期');
        else if (cover && cover.y - y >= 2) add('日期不合理', u.where, m[0], cover.where, cover.text, `評估日期與封面日期相差 ${cover.y - y} 年`);
        const key = `${y}/${mo}`;
        if (!evalDates.has(key)) evalDates.set(key, { u, text: m[0] });
      }
    }
    if (evalDates.size > 1) {
      const arr = [...evalDates.values()];
      for (let i = 1; i < arr.length; i++) add('日期不一致', arr[0].u.where, arr[0].text, arr[i].u.where, arr[i].text, '各表評估日期不一致');
    }

    // 4. 數值合理性
    for (const u of all) {
      const t = D.toHalf(u.text);
      for (const m of t.matchAll(/濕度[^0-9]{0,8}([0-9.]+)\s*[%％]/g)) if (+m[1] > 100) add('數值不合理', u.where, ctxOf(t, m.index, m[0].length), '', '', '濕度大於 100%');
      for (const m of t.matchAll(/年(?:平均)?(?:降)?雨量[^0-9]{0,12}([0-9,.]+)\s*(?:毫米|mm|公釐)/g)) {
        const v = parseFloat(m[1].replace(/,/g, ''));
        if (v < 300 || v > 8000) add('數值不合理', u.where, ctxOf(t, m.index, m[0].length), '', '', `年雨量 ${v} mm 量級不合理（可能是月雨量）`);
      }
      for (const m of t.matchAll(/(?:月)?平均(?:氣)?溫[^0-9]{0,8}([0-9.]+)\s*(?:°C|℃|度)/g)) if (+m[1] > 40) add('數值不合理', u.where, ctxOf(t, m.index, m[0].length), '', '', '平均氣溫超過 40°C');
      for (const m of t.matchAll(/([0-9,.]+)\s*(?:m²|㎡|平方公尺)\s*[（(]\s*(?:約)?\s*([0-9.]+)\s*公頃/g)) {
        const a = parseFloat(m[1].replace(/,/g, ''));
        const h = parseFloat(m[2]);
        if (a && h && Math.abs(a / 10000 - h) / h > 0.05) add('數值不合理', u.where, ctxOf(t, m.index, m[0].length), '', '', `面積換算不符：${a} m² ≠ ${h} 公頃`);
      }
    }

    // 5. 風險值
    for (const b of blocks) {
      if (b.kind !== 'table') continue;
      b.rows.forEach((r, ri) => {
        for (const item of K.scanRiskRow(r)) {
          for (const f of item.fixes) add('風險值錯誤', `${loc(b)} 第${ri + 1}列第${f.col + 1}欄`, f.from, '應為', f.to, `可能性 ${item.a} × 嚴重度 ${item.b} = ${item.value}（${item.level}）`);
        }
      });
    }

    // 6. 與案無關的殘留
    if (!ctx.isMaster) {
      const terms = new Map();
      for (const [label, f] of Object.entries(facts || {})) {
        const nv = String(f.newValue || '').trim();
        if (nv && nv === String(f.masterValue || '').trim()) continue;
        for (const lit of RG.literalsOf(f)) if (!nv.includes(lit)) terms.set(lit, `母本「${label}」`);
      }
      const newVals = Object.values(facts || {}).map((f) => String(f.newValue || '')).join('\n');
      for (const r of rules.residueExtra || []) {
        if (!r.term || (r.unlessTag && tagSet.has(r.unlessTag)) || newVals.includes(r.term)) continue;
        if (!terms.has(r.term)) terms.set(r.term, '母本專屬名詞');
      }
      for (const [term, why] of terms) {
        for (const u of all) {
          const i = u.text.indexOf(term);
          if (i >= 0) add('母本殘留', u.where, ctxOf(u.text, i, term.length), '', '', `出現${why}「${term}」`);
        }
      }
    }

    // 7. 錯字
    for (const f of rules.knownFixes || []) {
      const re = K.safeRe(f.find, 'g');
      if (!re) continue;
      for (const u of all) {
        re.lastIndex = 0;
        const m = re.exec(u.text);
        if (m) add('錯字', u.where, ctxOf(u.text, m.index, m[0].length), '建議', f.replace, f.note || '已知錯字');
      }
    }

    // 8. 待填與待審
    for (const u of all) {
      for (const m of u.text.matchAll(/【(待填|待補圖|待審)[^】]*】/g)) add('待填／待審', u.where, m[0], '', '', '尚未填寫或待審');
    }
    return issues;
  };

  C.summary = function (issues) {
    const s = {};
    for (const i of issues) s[i.item.split('：')[0]] = (s[i.item.split('：')[0]] || 0) + 1;
    return s;
  };
})();
