/* 關鍵字建議的單元測試（不需瀏覽器）：node tests/unit_keywords.js */
const path = require('path');
global.window = {};
require(path.join(__dirname, '..', 'js', 'rules.js'));
require(path.join(__dirname, '..', 'js', 'kb.js'));
require(path.join(__dirname, '..', 'js', 'memory.js'));
const RG = window.RG;
let failed = 0;
const ok = (c, m) => {
  console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`);
  if (!c) failed++;
};

const rules = RG.clone(RG.DEFAULT_RULES);
const mem = RG.mem.empty();
let i = 0;
const add = (text, cat, tags) => {
  const b = { id: 'b' + i++, kind: 'p', text, auto: { cat: 'T', tags: [] } };
  RG.mem.record(mem, b, { overrides: { [b.id]: { cat, tags } }, sig: 'm' });
};
// 你標了「電梯」，但用的詞是「升降機道」，目前關鍵字（電梯、昇降機）抓不到
add('升降機道內作業應設置安全網及護欄，防止人員墜落。', 'C', ['電梯']);
add('升降機道開口處應設置警示標誌與防護設施。', 'C', ['電梯']);
add('升降機道內之導軌安裝作業，應使用安全帶。', 'C', ['電梯']);
add('電梯機坑積水應先行抽排後再進入作業。', 'C', ['電梯']);
// 「天井」常誤判：你確認為 C 但標的是別的標籤
add('天井模板支撐應經計算確認其強度。', 'C', ['挑空區']);
add('天井周邊鷹架應設置交叉拉桿。', 'C', ['挑空區']);
add('天井區域高處作業應設置安全網。', 'C', ['挑空區']);
add('採光天井玻璃安裝應使用吊掛設備。', 'C', ['採光天井']);
// 其他一般段落
for (let k = 0; k < 15; k++) add(`第${k}項施工作業應依安全衛生計畫辦理，並設置必要之防護設施。`, 'T', []);
add('鷹架組立作業應由合格人員辦理，並設置上下設備。', 'C', ['鷹架']);
add('鷹架拆除時應由上而下依序拆除。', 'C', ['鷹架']);

const res = RG.kb.suggestKeywords(mem, rules, []);
console.log(JSON.stringify(res.map((r) => ({ tag: r.tag, missed: r.missed, add: r.add.map((a) => `${a.term}(+${a.gain},fp${a.fp})`), remove: r.remove.map((m) => `${m.kw}(tp${m.tp},fp${m.fp})`) })), null, 1));
const elev = res.find((r) => r.tag === '電梯');
ok(elev && elev.missed === 3, '電梯：3 筆目前關鍵字抓不到');
ok(elev && elev.add.some((a) => a.term.includes('升降機') && a.gain === 3 && a.fp === 0), '電梯：建議加入含「升降機」的詞，可多抓 3 筆、不誤抓');
const sky = res.find((r) => r.tag === '採光天井');
ok(sky && sky.remove.some((m) => m.kw === '天井' && m.fp === 3), '採光天井：建議移除常誤判的「天井」');
const scaf = res.find((r) => r.tag === '鷹架');
ok(scaf && scaf.add.some((a) => a.term.includes('鷹架')), '新標籤「鷹架」（規則裡沒有）也會建議關鍵字');
ok(!res.some((r) => r.add.some((a) => /\d|，|。/.test(a.term))), '建議詞不含數字與標點');

// 套用後再分析：電梯不再有「加入」建議
rules.tagKeywords['電梯'] = rules.tagKeywords['電梯'].concat(elev.add.map((a) => a.term));
const res2 = RG.kb.suggestKeywords(mem, rules, []);
const elev2 = res2.find((r) => r.tag === '電梯');
ok(!elev2 || !elev2.add.length, '採用後重新分析，電梯沒有新的加入建議');

// 專案事實（工程名稱）中的字不會被當成建議
const mem2 = RG.mem.empty();
i = 0;
const add2 = (text, tags) => {
  const b = { id: 'x' + i++, kind: 'p', text, auto: { cat: 'T', tags: [] } };
  RG.mem.record(mem2, b, { overrides: { [b.id]: { cat: 'C', tags } }, sig: 'm' });
};
add2('臺中市大雅區示範公園新建工程之滯留設施開挖作業。', ['滯留設施']);
add2('臺中市大雅區示範公園新建工程之滯留設施回填作業。', ['滯留設施']);
const res3 = RG.kb.suggestKeywords(mem2, rules, ['臺中市大雅區示範公園新建工程']);
const r3 = res3.find((r) => r.tag === '滯留設施');
ok(r3 && r3.add.every((a) => !'臺中市大雅區示範公園新建工程'.includes(a.term)), '工程名稱的字不會被建議成關鍵字');

// 先前版本存下的記憶（換行未保留）仍可對應
const legacy = RG.mem.empty();
const tb = { id: 't0', kind: 'table', text: '工程方案 | 備註\n電梯井道組立 | 無', auto: { cat: 'T', tags: [] } };
const ln = tb.text.normalize('NFKC').replace(/\s+/g, '');
legacy.entries['table:' + RG.mem.hash(ln) + ':' + ln.length] = { kind: 'table', len: ln.length, nh: ln, pv: '', cat: 'C', tags: ['電梯'], plan: '', masters: ['m'], conflict: false, updated: 1 };
const ov = {};
const rr = RG.mem.apply(legacy, [Object.assign({}, tb)], { overrides: ov, planOverrides: {}, sig: 'm' });
ok(rr.exact === 1 && ov.t0 && ov.t0.cat === 'C', '先前版本存下的記憶仍可完全對應');

console.log(failed ? `\n${failed} 項失敗` : '\n全部通過');
process.exit(failed ? 1 : 0);
