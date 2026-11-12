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
          // 很短、或「欄位：值」形式的段落視為 F（整段都是事實）；完整句子視為 M（句型中嵌入事實）
          const t = b.text.trim();
          cat = b.kind === 'p' && (t.length < 25 || (t.length < 60 && /^[^：:，。]{1,15}[：:]/.test(t))) ? 'F' : 'M';
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

  /* ---------- 關鍵字建議：從記憶庫（使用者確認過的分類與標籤）歸納 ----------
   * 對每個標籤：
   *  add    你標了這個標籤、但目前關鍵字抓不到的段落中，常共同出現、且很少出現在其他段落的詞
   *  remove 目前的關鍵字中，抓到的段落大多被你拿掉這個標籤者 */
  const BAD_TERM = /[\d\p{P}\p{S}\s]/u;
  /* 以虛詞開頭或結尾的詞不建議（例如「應由」「設置之」） */
  const FUNC_EDGE = /^[應之由及與並於的或等為以其時後前內上下中而且則需須將把被對向在有是]|[應之由及與並於的或等為以其而且則需須將把被對向在是]$/;
  K.suggestKeywords = function (memory, rules, lits, opts = {}) {
    const minPrecision = opts.minPrecision || 0.6;
    const maxAdd = opts.maxAdd || 5;
    const strip = (s) => {
      let t = s;
      for (const l of lits || []) {
        const n = String(l).normalize('NFKC').replace(/\s+/g, '');
        if (n) t = t.split(n).join('｜');
      }
      return t;
    };
    const entries = Object.values((memory && memory.entries) || {})
      .filter((e) => !e.conflict && e.nh)
      .map((e) => ({ e, t: strip(e.nh), tags: new Set(e.tags || []) }));
    if (!entries.length) return [];
    const allTags = new Set();
    entries.forEach((x) => x.tags.forEach((t) => allTags.add(t)));
    Object.keys(rules.tagKeywords || {}).forEach((t) => allTags.add(t));
    const out = [];
    for (const tag of allTags) {
      const kws = (rules.tagKeywords || {})[tag] || [];
      const re = kws.length ? safeRe(kws.join('|')) : null;
      const pos = entries.filter((x) => x.tags.has(tag));
      const neg = entries.filter((x) => !x.tags.has(tag));
      const missed = pos.filter((x) => !(re && re.test(x.t)));
      const res = { tag, pos: pos.length, missed: missed.length, add: [], remove: [] };

      // 建議加入
      if (missed.length) {
        const minDf = missed.length === 1 ? 1 : 2;
        const df = new Map();
        missed.forEach((x, i) => {
          const seen = new Set();
          for (let L = 2; L <= 6; L++) {
            for (let k = 0; k + L <= x.t.length; k++) {
              const g = x.t.substr(k, L);
              if (seen.has(g) || BAD_TERM.test(g) || FUNC_EDGE.test(g)) continue;
              seen.add(g);
              df.set(g, (df.get(g) || 0) + 1);
            }
          }
        });
        const cands = [...df.entries()].filter(([, n]) => n >= minDf).sort((a, b) => b[1] - a[1] || b[0].length - a[0].length).slice(0, 400);
        const scored = [];
        for (const [term] of cands) {
          const cov = new Set();
          missed.forEach((x, i) => x.t.includes(term) && cov.add(i));
          const fp = neg.filter((x) => x.t.includes(term));
          const tp = pos.filter((x) => x.t.includes(term)).length;
          const precision = tp / (tp + fp.length);
          if (precision < minPrecision) continue;
          scored.push({ term, cov, fp: fp.length, tp, precision, fpEx: fp.slice(0, 3).map((x) => x.e.pv) });
        }
        // 涵蓋相同段落、誤抓數相同的詞只留一個（長度最接近 4 字，平手取較長）
        const groups = new Map();
        for (const s of scored) {
          const key = [...s.cov].join(',') + '|' + s.fp;
          const cur = groups.get(key);
          const d = (t) => Math.abs(t.length - 4);
          if (!cur || d(s.term) < d(cur.term) || (d(s.term) === d(cur.term) && s.term.length > cur.term.length)) groups.set(key, s);
        }
        // 貪婪選擇：每次挑能多抓最多段落者
        const pool = [...groups.values()];
        const covered = new Set();
        while (res.add.length < maxAdd) {
          let best = null;
          let bestGain = 0;
          for (const s of pool) {
            let g = 0;
            s.cov.forEach((i) => !covered.has(i) && g++);
            if (g > bestGain || (g === bestGain && best && s.fp < best.fp)) {
              best = s;
              bestGain = g;
            }
          }
          if (!best || !bestGain) break;
          best.cov.forEach((i) => covered.add(i));
          res.add.push({
            term: best.term,
            gain: bestGain,
            fp: best.fp,
            examples: missed.filter((x, i) => best.cov.has(i)).slice(0, 3).map((x) => x.e.pv),
            fpExamples: best.fpEx,
          });
          pool.splice(pool.indexOf(best), 1);
        }
      }

      // 建議移除：關鍵字抓到的段落中，被你拿掉此標籤的比例過高
      for (const kw of kws) {
        const r = safeRe(kw);
        if (!r) continue;
        // 誤判只計「確認為 C 但沒有此標籤」的段落（T 類可能是法規、案例等本來就不判條件的章節）
        const hit = entries.filter((x) => r.test(x.t) && (x.tags.has(tag) || x.e.cat === 'C'));
        if (hit.length < 2) continue;
        const wrong = hit.filter((x) => !x.tags.has(tag));
        const tp = hit.length - wrong.length;
        if (wrong.length >= 2 && wrong.length > tp) res.remove.push({ kw, tp, fp: wrong.length, examples: wrong.slice(0, 3).map((x) => x.e.pv) });
      }
      if (res.add.length || res.remove.length) out.push(res);
    }
    return out.sort((a, b) => b.missed - a.missed);
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
