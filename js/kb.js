/* 知識庫：內容分類（F/T/C/M）、條件標籤初判、條件標籤推導、工項規劃、風險矩陣 */
(function () {
  'use strict';
  const RG = (window.RG = window.RG || {});
  const K = (RG.kb = {});

  K.CATS = {
    F: '專案事實',
    T: '通用範本',
    C: '條件式',
    M: '混合',
  };
  K.KIND_LABEL = {
    heading: '標題',
    caption: '表/圖名',
    p: '段落',
    table: '表格',
    rows: '子表格',
    figure: '圖',
    toc: '目錄',
    sdt: '內容控制項',
    empty: '空白段',
    other: '其他',
  };

  const safeRe = (src, flags) => {
    try {
      return new RegExp(src, flags);
    } catch (e) {
      return null;
    }
  };
  K.safeRe = safeRe;

  K.compileKeywords = function (tagKeywords) {
    const out = [];
    for (const [tag, kws] of Object.entries(tagKeywords || {})) {
      const re = safeRe((kws || []).join('|'));
      if (re && kws.length) out.push({ tag, re });
    }
    return out;
  };

  const FACT_RE = /(\d[\d,.]*\s*(m²|㎡|m2|平方公尺|公頃|毫米|mm|公釐|°C|℃|[%％]))|地號|\d{2,3}\s*年\s*\d{1,2}\s*月|電話|TEL/i;

  /* 初判分類與標籤，結果寫在 b.auto */
  K.classify = function (blocks, rules, facts) {
    const kw = K.compileKeywords(rules.tagKeywords);
    const tplSec = safeRe(rules.templateSections || '(?!)');
    const tplCap = safeRe(rules.templateCaptions || '(?!)');
    const factCap = safeRe(rules.factTableCaptions || '(?!)');
    const lits = [];
    for (const f of Object.values(facts || {})) lits.push(...RG.literalsOf(f));
    const byId = Object.fromEntries(blocks.map((b) => [b.id, b]));
    const captionText = (b) => (b.captionId && byId[b.captionId] ? byId[b.captionId].text : '');

    for (const b of blocks) {
      let cat = 'T';
      let tags = [];
      let reason = '';
      if (['heading', 'caption', 'toc', 'empty', 'other', 'figure'].includes(b.kind)) {
        reason = '結構元素（標題／表名／目錄等）';
      } else {
        let tagText = b.text;
        if (b.kind === 'table') tagText = captionText(b) + '\n' + (b.rows || []).slice(0, 3).map((r) => r.join(' ')).join('\n');
        if (b.kind === 'rows') tagText = b.tagText;
        // 專案事實本身（如工程名稱含「納骨塔」、「RC造」）不算條件關鍵字，先移除再比對
        for (const l of lits) tagText = tagText.split(l).join(' ');
        const cap = b.capName || '';
        const forcedT = (tplSec && tplSec.test(b.section || '')) || (tplCap && cap && tplCap.test(cap));
        if (!forcedT && b.chapter !== '前置' && !(b.kind === 'table' && b.groups)) {
          for (const k of kw) if (k.re.test(tagText)) tags.push(k.tag);
        }
        const hasLit = lits.find((l) => b.text.includes(l));
        const hasFact = !!hasLit || FACT_RE.test(b.text);
        if (forcedT) {
          cat = 'T';
          reason = '通用範本章節或表格（法規／案例／風險矩陣）';
        } else if (b.kind === 'table' && factCap && cap && factCap.test(cap)) {
          cat = 'F';
          reason = '專案事實表：' + cap;
        } else if (tags.length) {
          cat = 'C';
          reason = '出現條件關鍵字';
        } else if (hasFact) {
          cat = b.kind === 'p' && b.text.length < 40 ? 'F' : 'M';
          reason = hasLit ? `含母本專案事實「${hasLit}」` : '含數值、日期或地號等專案事實';
        } else {
          reason = '未偵測到專案事實或條件';
        }
      }
      b.auto = { cat, tags, reason };
    }
    return blocks;
  };

  K.effective = function (b, overrides) {
    const o = (overrides && overrides[b.id]) || {};
    const a = b.auto || { cat: 'T', tags: [] };
    return { cat: o.cat || a.cat, tags: o.tags || a.tags };
  };

  /* 從專案事實推導條件標籤 */
  K.deriveTags = function (facts, rules) {
    const set = new Set();
    const why = {};
    for (const r of rules.tagRules || []) {
      if (r.always) {
        (r.tags || []).forEach((t) => {
          set.add(t);
          why[t] = '固定';
        });
        continue;
      }
      const v = facts[r.field] && facts[r.field].newValue;
      const re = safeRe(r.match);
      if (v && re && re.test(v)) {
        (r.tags || []).forEach((t) => {
          set.add(t);
          why[t] = `${r.field}「${v}」`;
        });
      }
    }
    return { tags: [...set], why };
  };

  K.projectTags = function (state) {
    const d = K.deriveTags(state.facts, state.rules);
    const set = new Set(d.tags);
    for (const [t, v] of Object.entries(state.manualTags || {})) {
      if (v === 'on') set.add(t);
      if (v === 'off') set.delete(t);
    }
    return { set, derived: d };
  };

  /* 工項規劃：決定每個區塊保留或刪除 */
  K.plan = function (blocks, overrides, tagSet, planOverrides, opts = {}) {
    const res = {};
    const byId = Object.fromEntries(blocks.map((b) => [b.id, b]));
    const decide = (b) => {
      const po = planOverrides && planOverrides[b.id];
      if (po === 'keep' || po === 'drop') return { include: po === 'keep', reason: '使用者指定', user: true };
      if (opts.keepAll) return { include: true, reason: '保留全部（還原測試）' };
      const e = K.effective(b, overrides);
      if (e.cat === 'C' && e.tags.length) {
        const miss = e.tags.filter((t) => !tagSet.has(t));
        return miss.length
          ? { include: false, reason: '缺少條件標籤：' + miss.join('、') }
          : { include: true, reason: '條件符合：' + e.tags.join('、') };
      }
      return { include: true, reason: '' };
    };
    for (const b of blocks) if (!b.linkTo) res[b.id] = decide(b);
    // 子表格全刪 → 整張表刪
    for (const b of blocks) {
      if (b.kind === 'table' && b.groups && res[b.id].include && !res[b.id].user) {
        if (b.groups.every((g) => !res[g.id].include)) res[b.id] = { include: false, reason: '所有分項工程皆不適用' };
      }
      if (b.parentId && !res[b.parentId].include) res[b.id] = { include: false, reason: '所屬表格已刪除' };
    }
    // 表名、分項工程標題跟隨其表格
    for (const b of blocks) {
      if (!b.linkTo) continue;
      const po = planOverrides && planOverrides[b.id];
      if (po === 'keep' || po === 'drop') res[b.id] = { include: po === 'keep', reason: '使用者指定', user: true };
      else {
        const t = res[b.linkTo];
        const own = decide(b);
        if (t && !t.include) res[b.id] = { include: false, reason: `隨${byId[b.linkTo].capNo ? '表' + byId[b.linkTo].capNo : '所屬表格'}刪除` };
        else res[b.id] = own;
      }
    }
    // 分項工程標題被刪 → 其表格也刪
    for (const b of blocks) {
      if (b.linkTo && b.kind === 'p' && !res[b.id].include && res[b.linkTo] && res[b.linkTo].include && !res[b.linkTo].user) {
        res[b.linkTo] = { include: false, reason: '隨分項工程標題刪除' };
        const t = byId[b.linkTo];
        if (t.captionId && !(planOverrides || {})[t.captionId]) res[t.captionId] = { include: false, reason: '隨表格刪除' };
      }
    }
    return res;
  };

  /* ---------- 風險矩陣（第 3.4 節） ---------- */
  K.riskLevel = (v) => (v >= 6 ? 'H' : v >= 3 ? 'M' : 'L');
  const int13 = (s) => {
    const t = RG.docx.toHalf(String(s || '')).trim();
    return /^[1-3]$/.test(t) ? parseInt(t, 10) : null;
  };
  /* 在一列儲存格中找出「可能性、嚴重度、風險值/等級」組合，回傳需修正的項目 */
  K.scanRiskRow = function (cells) {
    const out = [];
    const c = cells.map((x) => RG.docx.toHalf(String(x || '')).trim());
    for (let i = 0; i + 2 < c.length; i++) {
      const a = int13(c[i]);
      const b = int13(c[i + 1]);
      if (a === null || b === null) continue;
      const v = a * b;
      const lv = K.riskLevel(v);
      if (/^\d+$/.test(c[i + 2])) {
        const item = { col: i, a, b, value: v, level: lv, fixes: [] };
        if (parseInt(c[i + 2], 10) !== v) item.fixes.push({ col: i + 2, from: c[i + 2], to: String(v) });
        if (i + 3 < c.length && /^[HML]$/i.test(c[i + 3]) && c[i + 3].toUpperCase() !== lv) item.fixes.push({ col: i + 3, from: c[i + 3], to: lv });
        out.push(item);
        i += 2;
      } else if (/^[HML]$/i.test(c[i + 2])) {
        const item = { col: i, a, b, value: v, level: lv, fixes: [] };
        if (c[i + 2].toUpperCase() !== lv) item.fixes.push({ col: i + 2, from: c[i + 2], to: lv });
        out.push(item);
        i += 2;
      }
    }
    return out;
  };
})();
