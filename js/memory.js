/* 依文字內容記憶分類與標籤：換母本、母本改版、其他案報告，只要文字相同（或相似）就沿用先前的確認結果 */
(function () {
  'use strict';
  const RG = (window.RG = window.RG || {});
  const K = RG.kb;
  const M = (RG.mem = {});

  M.KINDS = ['p', 'table', 'rows', 'sdt'];
  M.HEAD = 600; // 模糊比對只用正規化後前 600 字
  M.empty = () => ({ version: 1, threshold: 0.9, entries: {} });

  M.norm = (s) => String(s || '').normalize('NFKC').replace(/\s+/g, '');
  /* cyrb53：快速字串雜湊 */
  M.hash = function (str, seed = 0) {
    let h1 = 0xdeadbeef ^ seed,
      h2 = 0x41c6ce57 ^ seed;
    for (let i = 0; i < str.length; i++) {
      const ch = str.charCodeAt(i);
      h1 = Math.imul(h1 ^ ch, 2654435761);
      h2 = Math.imul(h2 ^ ch, 1597334677);
    }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
  };
  const textOf = (b) => (b.kind === 'rows' ? b.title + '\n' + b.text : b.text);
  M.memorable = (b) => M.KINDS.includes(b.kind) && M.norm(textOf(b)).length > 0;
  M.keyOf = function (b) {
    if (b._memKey) return b._memKey;
    const n = M.norm(textOf(b));
    b._memNorm = n;
    b._memKey = b.kind + ':' + M.hash(n) + ':' + n.length;
    return b._memKey;
  };

  /* 字元雙字組 Dice 相似度 */
  const bigrams = (s) => {
    const m = new Map();
    for (let i = 0; i < s.length - 1; i++) {
      const g = s.substr(i, 2);
      m.set(g, (m.get(g) || 0) + 1);
    }
    return m;
  };
  M.bigrams = bigrams;
  /* a、b 為字串；A、B 可傳入預先算好的雙字組以加速 */
  M.dice = function (a, b, A, B) {
    if (a === b) return 1;
    if (a.length < 2 || b.length < 2) return 0;
    A = A || bigrams(a);
    B = B || bigrams(b);
    if (A.size > B.size) [A, B] = [B, A];
    let inter = 0;
    for (const [g, n] of A) if (B.has(g)) inter += Math.min(n, B.get(g));
    return (2 * inter) / (a.length - 1 + (b.length - 1));
  };

  const sameTags = (x, y) => (x || []).join('、') === (y || []).join('、');

  /* 記錄一個區塊目前的分類／標籤（以及工項規劃的使用者指定） */
  M.record = function (mem, b, ctx) {
    if (!M.memorable(b)) return null;
    const key = M.keyOf(b);
    const e = K.effective(b, ctx.overrides);
    const old = mem.entries[key] || {};
    const plan = ctx.planOverrides ? ctx.planOverrides[b.id] || '' : old.plan || '';
    mem.entries[key] = {
      kind: b.kind,
      len: b._memNorm.length,
      nh: b._memNorm.slice(0, M.HEAD),
      pv: textOf(b).replace(/\s+/g, ' ').trim().slice(0, 120),
      cat: e.cat,
      tags: e.tags.slice(),
      plan,
      masters: [...new Set([...(old.masters || []), ctx.sig].filter(Boolean))],
      conflict: false,
      updated: Date.now(),
    };
    return mem.entries[key];
  };

  /* 一次記錄多個區塊；同一份母本裡相同文字卻有不同分類者標為衝突 */
  M.recordMany = function (mem, blocks, ctx) {
    const seen = new Map();
    let n = 0;
    for (const b of blocks) {
      if (!M.memorable(b)) continue;
      const key = M.keyOf(b);
      const e = K.effective(b, ctx.overrides);
      const sigVal = e.cat + '|' + e.tags.join('、') + '|' + ((ctx.planOverrides || {})[b.id] || '');
      const ent = M.record(mem, b, ctx);
      if (seen.has(key) && seen.get(key) !== sigVal) ent.conflict = true;
      seen.set(key, sigVal);
      n++;
    }
    return n;
  };

  /* 讀入母本後套用記憶。
   * ctx: {overrides, planOverrides, sig, prevSig, keepExisting}
   * 每個可記憶區塊會得到 b.mem = {state: 'exact'|'fuzzy'|'new', sim, from} */
  M.apply = function (mem, blocks, ctx) {
    const res = { exact: 0, fuzzy: 0, fresh: 0, missing: [] };
    const thr = mem.threshold || 0.9;
    const entries = mem.entries;
    const used = new Set();
    const pending = [];
    const put = (b, ent) => {
      if (ctx.keepExisting && ctx.overrides[b.id]) return;
      const o = {};
      if (ent.cat !== b.auto.cat) o.cat = ent.cat;
      if (!sameTags(ent.tags, b.auto.tags)) o.tags = ent.tags.slice();
      if (Object.keys(o).length) ctx.overrides[b.id] = o;
      else delete ctx.overrides[b.id];
      if (ent.plan && !(ctx.keepExisting && ctx.planOverrides[b.id])) ctx.planOverrides[b.id] = ent.plan;
    };
    for (const b of blocks) {
      if (!M.memorable(b)) continue;
      const key = M.keyOf(b);
      const ent = entries[key];
      if (ent && !ent.conflict) {
        put(b, ent);
        used.add(key);
        if (ctx.sig && !ent.masters.includes(ctx.sig)) ent.masters.push(ctx.sig);
        b.mem = { state: 'exact' };
        res.exact++;
      } else if (ent && ent.conflict) {
        put(b, ent);
        used.add(key);
        b.mem = { state: 'fuzzy', sim: 1, from: ent.pv, note: '相同文字在前版有不同的分類，請確認' };
        res.fuzzy++;
      } else pending.push(b);
    }
    // 模糊比對：同類型、長度相近
    if (pending.length) {
      const byKind = {};
      // 已被完全相同對應用掉的記憶也可以再拿來推定相似段落（例如表4-13 許多相似的風險條目）
      for (const [key, ent] of Object.entries(entries)) {
        if (ent.conflict) continue;
        (byKind[ent.kind] = byKind[ent.kind] || []).push([key, ent, null]);
      }
      for (const b of pending) {
        const n = b._memNorm;
        const head = n.slice(0, M.HEAD);
        const HB = bigrams(head);
        let best = null;
        let bestSim = 0;
        for (const item of byKind[b.kind] || []) {
          const [key, ent] = item;
          const r = ent.len / n.length;
          if (r < 0.8 || r > 1.25) continue;
          if (!item[2]) item[2] = bigrams(ent.nh); // 快取
          const s = M.dice(head, ent.nh, HB, item[2]);
          if (s > bestSim) {
            bestSim = s;
            best = [key, ent];
          }
        }
        if (best && bestSim >= thr) {
          put(b, best[1]);
          used.add(best[0]);
          b.mem = { state: 'fuzzy', sim: bestSim, from: best[1].pv };
          res.fuzzy++;
        } else {
          b.mem = { state: 'new' };
          res.fresh++;
        }
      }
    }
    if (ctx.prevSig && ctx.prevSig !== ctx.sig) {
      for (const [key, ent] of Object.entries(entries)) if (!used.has(key) && ent.masters.includes(ctx.prevSig)) res.missing.push(ent);
    }
    return res;
  };

  M.merge = function (mem, other) {
    let n = 0;
    for (const [key, ent] of Object.entries((other && other.entries) || {})) {
      const cur = mem.entries[key];
      if (!cur || (ent.updated || 0) > (cur.updated || 0)) {
        mem.entries[key] = Object.assign({}, ent, { masters: [...new Set([...((cur && cur.masters) || []), ...(ent.masters || [])])] });
        n++;
      }
    }
    return n;
  };

  M.stats = function (mem) {
    const entries = Object.values(mem.entries);
    const masters = new Set();
    entries.forEach((e) => (e.masters || []).forEach((m) => masters.add(m)));
    return { count: entries.length, masters: masters.size, conflicts: entries.filter((e) => e.conflict).length };
  };
})();
