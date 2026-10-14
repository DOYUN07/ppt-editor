'use strict';
// ============================================================================
// Importing existing presentations
//   1) 원본 디자인 유지 수정  — edit text/pictures/tables/notes in place (engine/scripts/apply.ps1)
//   2) 새 양식으로 다시 만들기 — content -> template slides
//   3) 자유 편집으로 가져오기  — every shape becomes an editable canvas element (background kept as an image)
// ============================================================================
let imp = null;          // { id, name, dir, info, edits, sel, shape }
const SOFT = '↵';  // soft line break inside a paragraph (shown as ↵)
const shapeText = s => (s.paragraphs || []).map(p => p.text).join('\n');
const isJunk = s => [13, 15, 16].includes(s.ph) || /^(\d{1,2}|STEP\s*\d+|[—→•·\-↵]*)$/i.test(shapeText(s).trim());
const impUrl = (kind, file) => `/imports/${imp.id}/${kind}/${file}`;
const impPath = file => imp.dir + '\\media\\' + file;

async function importPpt(file) {
  const m = $('modal'); $('modalBack').classList.add('show');
  m.innerHTML = `<h3><span class="spinner"></span>PowerPoint로 여는 중…</h3><div class="hint">${esc(file.name)} — 슬라이드가 많으면 1~2분 걸릴 수 있어요.</div>`;
  try {
    const r = await fetch('/api/import?name=' + encodeURIComponent(file.name), { method: 'POST', body: file });
    const res = await r.json();
    if (!r.ok || res.error) throw new Error((res.error || '가져오지 못했어요.') + (res.log ? '\n' + res.log : ''));
    imp = { id: res.id, name: res.name, dir: res.dir, info: res.info, edits: {}, sel: 0, shape: null };
    m.innerHTML = `<h3>'${esc(res.name)}' — ${res.info.slideCount}장을 가져왔어요</h3>
      <div class="row" style="gap:6px;margin-top:6px">${res.info.slides.slice(0, 6).map(s => `<img src="${impUrl('png', s.png)}" style="width:120px;border-radius:3px;box-shadow:0 1px 4px rgba(0,0,0,.15)">`).join('')}</div>
      <div class="choice">
        <button id="chFree"><b>자유 편집으로 가져오기</b><span class="hint">모든 글상자·도형·사진을 캔버스 요소로 바꿔요. 끌어서 옮기고 글꼴·색을 바꾸고 새 요소를 더할 수 있어요.</span></button>
        <button id="chOrig"><b>원본 그대로 두고 고치기</b><span class="hint">원본 디자인을 100% 유지한 채 글자·사진·표·노트만 바꾸고 슬라이드를 지워요.</span></button>
        <button id="chNew"><b>새 양식으로 다시 만들기</b><span class="hint">내용만 뽑아서 편집기 양식(목차·카드·지표·단계…)으로 새로 디자인해요.</span></button>
      </div>
      <div class="row" style="margin-top:12px"><span class="grow"></span><button class="btn" id="closeModal">취소</button></div>`;
    $('chFree').onclick = () => { $('modalBack').classList.remove('show'); freeImport(); };
    $('chOrig').onclick = () => { $('modalBack').classList.remove('show'); openOrig(); };
    $('chNew').onclick = () => { $('modalBack').classList.remove('show'); convertToDeck(); };
  } catch (e) {
    m.innerHTML = `<h3>가져오지 못했어요</h3><div class="hint">암호가 걸린 파일이거나 PowerPoint가 다른 대화상자를 띄운 상태일 수 있어요.</div><pre class="log">${esc(e.message)}</pre><button class="btn" id="closeModal">닫기</button>`;
  }
  const cm = $('closeModal'); if (cm) cm.onclick = () => $('modalBack').classList.remove('show');
}
function replaceDeck(slides, title, fileName) {
  if (state.deck.slides.length && !confirm(`지금 편집 중인 슬라이드(${state.deck.slides.length}장)를 가져온 내용(${slides.length}장)으로 바꿀까요?\n(지금 내용이 필요하면 취소하고 먼저 '저장'을 하세요)`)) return false;
  snap();
  state.deck = { title, footer: state.deck.footer || title, size: { w: SW, h: SH }, slides };
  state.fileName = fileName; cur = 0; sel = []; $('orig').hidden = true; renderAll();
  return true;
}

// ---- 3) free edit: shapes -> canvas elements ----
const AST_TO_SHAPE = { 1: 'rect', 5: 'roundRect', 9: 'ellipse', 7: 'triangle', 8: 'rtTriangle', 4: 'diamond', 2: 'parallelogram', 3: 'trapezoid', 12: 'pentagon', 10: 'hexagon', 6: 'octagon', 91: 'star4', 92: 'star5', 147: 'star6', 11: 'plus', 21: 'heart', 33: 'arrowRight', 34: 'arrowLeft', 35: 'arrowUp', 36: 'arrowDown', 37: 'arrowLR', 52: 'chevron', 51: 'homePlate', 106: 'callout', 179: 'cloud', 18: 'donut', 22: 'lightning', 158: 'frame', 155: 'snipRect', 152: 'round2Same' };
const DASH_MAP = { 1: 'solid', 4: 'dash', 3: 'dot', 5: 'dashDot', 7: 'longDash', 2: 'dot', 6: 'dashDot', 8: 'dashDot' };
function chartKind(t) { if ([57, 58, 59, 60, -4099].includes(t)) return t === 58 ? 'barStacked' : 'bar'; if ([4, 63, 64, 65, 66, 67].includes(t)) return 'line'; if ([5, -4102, 69, 70].includes(t)) return 'pie'; if (t === -4120 || t === 80) return 'doughnut'; if ([1, 76, 77].includes(t)) return t === 76 ? 'areaStacked' : 'area'; if (t === 52 || t === 53) return 'columnStacked'; return 'column'; }
// symbol-font bullet characters (Wingdings / Symbol code points) -> unicode
const SYMBOL_BULLETS = { 0xFC: '✓', 0xFB: '✗', 0xFE: '☑', 0xA7: '▪', 0xA8: '◻', 0x6C: '●', 0x6E: '■', 0x6F: '□', 0x70: '◼', 0x71: '❑', 0x75: '◆', 0x76: '❖', 0x77: '⬥', 0xD8: '➢', 0xE0: '➔', 0xE8: '➔', 0xF0: '⇨', 0x9F: '•', 0xB7: '•', 0x9E: '·', 0xA1: '○', 0xA4: '◉', 0x4A: '☺', 0xAB: '★' };
function bulletChar(p) {
  let c = +p.bulletChar || 0x2022; const sym = /wingdings|symbol|webdings/i.test(p.bulletFont || '');
  if (c >= 0xF000 && c <= 0xF0FF) c -= 0xF000;
  if (sym) return SYMBOL_BULLETS[c] || '•';
  return c === 0x2022 ? 'disc' : (c === 0x25A0 || c === 0x25AA) ? 'square' : String.fromCodePoint(c);
}
let fontAlias = null;   // localized font name -> name the browser and PowerPoint both understand
let fontNames = null;   // every installed family name (any language), lower-case
async function loadFontAlias() {
  if (fontAlias) return fontAlias;
  try { const r = await (await fetch('/api/fontalias')).json(); fontAlias = r.alias || {}; fontNames = new Set((r.names || []).map(f => f.toLowerCase())); }
  catch (e) { fontAlias = {}; fontNames = null; }
  return fontAlias;
}
function textFromImport(sh, s) {
  const st = sh.style || {}, pts = (v, rule, size) => rule === -1 ? v * size * 1.2 : v;
  const paras = (sh.paragraphs || []).map(p => {
    const runs = (p.runs && p.runs.length ? p.runs : [{ t: p.text }]).map(r => { const o = { t: String(r.t ?? '').replace(new RegExp(SOFT, 'g'), '\n') };
      if (r.size) o.size = Math.round(r.size * s * 10) / 10; if (r.b) o.b = true; if (r.i) o.i = true; if (r.u) o.u = true; if (r.color) o.color = '#' + r.color; if (r.font) o.font = fontAlias[r.font] || r.font; return o; });
    const q = { runs }; if (p.level > 1) q.level = p.level;
    if (p.bullet === 1) { const b = bulletChar(p); if (b === 'disc' || b === 'square') q.bullet = b; else { q.bullet = 'char'; q.bulletChar = b; } }
    if (p.bullet === 2) q.bullet = 'num';
    q.align = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify' }[p.align] || 'left';
    const size = p.size || (p.runs && p.runs[0] && p.runs[0].size) || 18;
    if (p.sa > 0) q.spaceAfter = Math.round(pts(p.sa, p.saRule, size) * s * 10) / 10;
    if (p.sb > 0) q.spaceBefore = Math.round(pts(p.sb, p.sbRule, size) * s * 10) / 10;
    return q;
  });
  const first = paras[0] && paras[0].runs[0] || {}, p0 = (sh.paragraphs || [])[0] || {};
  // line spacing of the first paragraph: multiples of a line, or exact points converted to a multiple
  let ls = 1.0; if (p0.ls > 0) ls = p0.lsRule === -1 ? p0.ls : p0.ls / ((p0.size || 18) * 1.2);
  const tx = { paras, size: first.size || 18, color: 'dk1', font: 'body', bold: false, align: (paras[0] && paras[0].align) || 'left', valign: { 1: 'top', 3: 'middle', 4: 'bottom' }[st.valign] || 'top', lineSpacing: Math.round(Math.min(3, Math.max(0.7, ls)) * 100) / 100, pad: (st.pad || [3.6, 7.2, 3.6, 7.2]).map(v => v * s) };
  if (st.nowrap) tx.wrap = false;
  tx.fitPpt = true;
  return tx;
}
// fonts used by the imported runs that this PC does not have
function missingFonts(slides) {
  const have = new Set((server.fonts || []).map(f => f.toLowerCase())), miss = new Map();
  if (fontNames) fontNames.forEach(f => have.add(f));
  slides.forEach(sl => sl.extra.forEach(e => { if (!e.text) return; e.text.paras.forEach(p => p.runs.forEach(r => { if (r.font && !have.has(r.font.toLowerCase()) && !/^\+m[jn]-/.test(r.font)) miss.set(r.font, (miss.get(r.font) || 0) + 1); })); }));
  return [...miss.keys()];
}
// exported pictures already carry rotation/flip/crop; rotated ones come out at their rotated outline size
function exportedBox(sh, g, s, k, ox, oy) {
  const st = sh.style || {};
  // cut-outs carry their exact drawn bounds (already rotated / flipped)
  if (sh.snapBox) { const [x, y, w, h] = sh.snapBox; return { x: ox + x * s, y: oy + y * s, w: w * s, h: h * s }; }
  if (!st.rot || !sh.px || !k) return g;
  const w = sh.px[0] / k * s, h = sh.px[1] / k * s, cx = g.x + g.w / 2, cy = g.y + g.h / 2;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}
async function freeImport() {
  await loadFontAlias();
  const info = imp.info, s = Math.min(SW / info.width, SH / info.height), ox = (SW - info.width * s) / 2, oy = (SH - info.height * s) / 2;
  // export resolution (pixels per point), measured on unrotated exports
  const ks = []; info.slides.forEach(sl => sl.shapes.forEach(sh => { if (sh.px && !(sh.style || {}).rot && sh.w > 20) ks.push(sh.px[0] / sh.w); }));
  ks.sort((a, b) => a - b); const k = ks.length ? ks[ks.length >> 1] : 4 / 3;
  let skipped = 0;
  const slides = info.slides.map(sl => {
    const extra = []; let z = 0;
    sl.shapes.forEach(sh => {
      if (sh.kind === 'group' || (sh.kind === 'text' && sh.empty)) return;
      const st = sh.style || {}, g = { x: ox + sh.x * s, y: oy + sh.y * s, w: sh.w * s, h: sh.h * s };
      const common = Object.assign({ id: uid(), z: z++ }, g); if (st.rot) common.rot = st.rot; if (st.flipH) common.flipH = true; if (st.flipV) common.flipV = true; if (st.shadow) common.shadow = true;
      const line = st.line && st.line.width > 0 ? { color: '#' + st.line.color, width: Math.round(st.line.width * s * 10) / 10, dash: DASH_MAP[st.line.dash] || 'solid' } : null;
      const exported = sh.snap || (sh.kind === 'picture' && sh.image);
      if (exported) { const o = Object.assign({ id: common.id, z: common.z }, exportedBox(sh, g, s, k, ox, oy)); if (st.shadow && !sh.snapBox) o.shadow = true;
        extra.push(Object.assign(o, { kind: 'image', image: { src: impPath(exported), url: impUrl('media', exported), fit: 'stretch' } })); return; }
      if (sh.kind === 'picture') { skipped++; return; }
      if (sh.kind === 'table' && sh.table && sh.table.length) { const nc = sh.table[0].length; extra.push(Object.assign(common, { kind: 'table', table: { rows: sh.table, header: true, colW: Array.from({ length: nc }, () => g.w / nc), fontSize: 13, headerFill: 'dk2', headerColor: 'lt1', bandFill: 'lt2', bodyFill: 'lt1', color: 'dk1', banded: true, firstColBold: false } })); return; }
      if (sh.kind === 'chart') { if (sh.categories && sh.categories.length) extra.push(Object.assign(common, { kind: 'chart', chart: { type: chartKind(sh.chartType), categories: sh.categories, series: [Object.assign({ name: sh.seriesName || '값', values: sh.values || [] }, sh.seriesColor && !/pie|doughnut/.test(chartKind(sh.chartType)) ? { color: '#' + sh.seriesColor } : {})],
        showValues: sh.labels !== false, highlight: -1, ...(sh.pointColors ? { pointColors: sh.pointColors.map(c => '#' + c) } : {}), ...(sh.axisMin != null ? { axisMin: sh.axisMin } : {}), ...(sh.axisMax != null ? { axisMax: sh.axisMax } : {}), ...(sh.numFmt && sh.numFmt !== 'General' ? { numFmt: sh.numFmt } : {}), ...(sh.valueAxis === false ? { valueAxis: false } : {}), ...(sh.gridlines === false ? { gridlines: false } : {}) } })); else skipped++; return; }
      if (sh.kind === 'line') { extra.push(Object.assign(common, { kind: 'line', line: Object.assign({ color: '#000000', width: 1 }, line || {}, { end: st.line && st.line.endArrow > 1 ? 'arrow' : 'none', start: st.line && st.line.beginArrow > 1 ? 'arrow' : 'none' }) })); return; }
      // outline: editor preset, or the exact path read from the file (block arc, teardrop, freeform ...)
      const shapeName = sh.geom ? 'custom' : AST_TO_SHAPE[st.ast] || 'rect';
      const hasFrame = st.fill || line || st.gradient || sh.geom || (st.ast && st.ast !== 1 && st.ast !== 138);
      const tx = sh.kind === 'text' ? textFromImport(sh, s) : null;
      if (tx && !hasFrame) { const e = Object.assign(common, { kind: 'text', text: tx }); if (sh.ph === 1 || sh.ph === 3) e.isTitle = true; extra.push(e); return; }
      const e = Object.assign(common, { kind: 'shape', shape: shapeName, fill: st.fill ? '#' + st.fill : null, line });
      if (sh.geom) e.geom = sh.geom;
      if (st.fillAlpha != null) e.fillOpacity = st.fillAlpha; if (st.radius != null) e.radius = st.radius;
      // effects stay editable values
      if (st.gradient) e.gradient = { type: st.gradient.type, angle: st.gradient.angle ?? 90, stops: st.gradient.stops.map(g => ({ pos: g.pos, color: '#' + g.color, alpha: g.alpha })).sort((a, b) => a.pos - b.pos) };
      if (st.shadowFx) { const f = st.shadowFx; e.shadow = true; e.shadowFx = { color: '#' + f.color, alpha: f.alpha, blur: f.blur * s, dx: f.dx * s, dy: f.dy * s }; }
      if (st.glow) e.glow = { r: st.glow.r * s, color: '#' + st.glow.color, alpha: st.glow.alpha };
      if (st.softEdge) e.softEdge = st.softEdge * s;
      if (tx) e.text = tx;
      extra.push(e);
    });
    const out = { id: uid(), template: null, ov: {}, del: [], extra, notes: sl.notes || '', hidden: !!sl.hidden };
    if (sl.bg) out.bg = { image: impPath(sl.bg), imageUrl: impUrl('media', sl.bg) };
    return out;
  });
  if (replaceDeck(slides, imp.name, imp.name + ' (편집)')) {
    toast(`${slides.length}장을 자유 편집 요소로 가져왔어요.${skipped ? ` (${skipped}개 요소는 옮기지 못했어요)` : ''} 배경과 슬라이드 마스터 그림은 배경 이미지로 들어가 있어요.`);
    offerImportStyle(slides, missingFonts(slides));
  }
}

// ---- 2) convert to template slides ----
function columns(bs) { const cols = []; bs.forEach(b => { let c = cols.find(c => Math.abs(c.x - b.x) < 25); if (!c) { c = { x: b.x, items: [] }; cols.push(c); } c.items.push(b); }); cols.forEach(c => c.items.sort((a, b) => a.y - b.y)); return cols.sort((a, b) => a.x - b.x); }
function pairs(bs, one) { const out = [], used = new Set(); [...bs].sort((a, b) => (a.y - b.y) || (a.x - b.x)).forEach(b => { if (used.has(b)) return; const d = bs.filter(o => !used.has(o) && o !== b && Math.abs(o.x - b.x) < 25 && o.y > b.y && o.y - (b.y + b.h) < 40).sort((a, c) => a.y - c.y)[0]; used.add(b); if (d) used.add(d); out.push({ title: one(b), desc: d ? one(d) : '' }); }); return out; }
function convertToDeck() {
  const info = imp.info, W = info.width, H = info.height; let secN = 0;
  const mk = (type, data, notes) => ({ id: uid(), template: { type, data }, ov: {}, del: [], extra: [], notes: notes || '' });
  const slides = info.slides.filter(s => !s.hidden).map((s, i, arr) => {
    const texts = s.shapes.filter(x => x.kind === 'text' && !x.empty && shapeText(x).trim());
    let title = texts.find(x => x.ph === 1 || x.ph === 3);
    if (!title) title = texts.filter(x => x.y < H * 0.35 && !isJunk(x)).sort((a, b) => Math.max(0, ...b.paragraphs.map(p => p.size)) - Math.max(0, ...a.paragraphs.map(p => p.size)))[0];
    const titleText = title ? shapeText(title).replace(new RegExp(SOFT, 'g'), '\n') : '';
    const bodies = texts.filter(x => x !== title && !isJunk(x)).sort((a, b) => (a.y - b.y) || (a.x - b.x));
    const bodyText = bodies.map(shapeText).join('\n');
    const pics = s.shapes.filter(x => x.kind === 'picture' && x.image).sort((a, b) => b.w * b.h - a.w * a.h);
    const pic = pics[0] && pics[0].w * pics[0].h > W * H * 0.04 ? pics[0] : null;
    const table = s.shapes.find(x => x.kind === 'table' && x.table && x.table.length);
    const chart = s.shapes.find(x => x.kind === 'chart' && x.categories && x.categories.length);
    const img = pic ? { image: impPath(pic.image), imageUrl: impUrl('media', pic.image) } : {};
    const sub = bodies.find(x => x.ph === 4) || bodies[0];
    const one = b => shapeText(b).replace(new RegExp(SOFT, 'g'), ' ').replace(/\n/g, ' ').trim();
    if (i === 0) return mk('cover', Object.assign({ section: '표지', title: titleText || imp.name, subtitle: sub ? shapeText(sub).split('\n')[0] : '', meta: bodies.filter(x => x !== sub).map(x => shapeText(x).split('\n')[0]).join('  ·  ') }, img), s.notes);
    if (i === arr.length - 1 && /감사|thank|q\s*&\s*a|문의|끝/i.test(titleText + ' ' + bodyText.slice(0, 40))) return mk('closing', Object.assign({ section: '마무리', title: titleText || '감사합니다', subtitle: sub ? shapeText(sub).split('\n')[0] : '', meta: bodies.filter(x => x !== sub).map(x => shapeText(x).split('\n')[0]).join('  ·  ') }, img), s.notes);
    if (table) return mk('table', { title: titleText, headers: table.table[0], rows: table.table.slice(1), statusColors: { '완료': 'accent4', '진행 중': 'accent1', '예정': 'accent6' } }, s.notes);
    if (chart) { const cap = bodies.find(b => b.y + b.h <= chart.y + 8 && Math.abs(b.x - chart.x) < 80);
      return mk('chart', { title: titleText, caption: cap ? one(cap) : '', chartType: chartKind(chart.chartType), seriesName: chart.seriesName || '값', categories: chart.categories, values: chart.values || [], highlight: -1, insights: pairs(bodies.filter(b => b !== cap), one).slice(0, 3) }, s.notes); }
    if (!bodies.length && !pic) return mk('section', { number: N2(++secN), title: titleText || '섹션', desc: '' }, s.notes);
    const cols = columns(bodies), prs = pairs(bodies, one);
    if (/목차|차례|agenda|contents|index/i.test(titleText) && prs.length >= 2) return mk('agenda', { title: titleText, items: prs.slice(0, 6) }, s.notes);
    const isNum = t => /^[+\-]?[\d.,]+\s*(%|%p|배|억|만|천|원|명|건|개|x|X|점|위|년|개월|시간|분|K|M|B)?\+?$/.test(t);
    const numCols = cols.filter(c => c.items.some(b => isNum(one(b))));
    if (numCols.length >= 2 && numCols.length <= 4) {
      const used = new Set();
      const stats = numCols.map(c => { const i0 = c.items.findIndex(b => isNum(one(b))); const v = c.items[i0], l = c.items[i0 + 1], nt = c.items[i0 + 2]; [v, l, nt].forEach(b => b && used.add(b)); return { value: one(v), label: l ? one(l) : '', note: nt ? one(nt) : '' }; });
      const rest = bodies.filter(b => !used.has(b) && one(b).length > 12).sort((a, b) => one(b).length - one(a).length);
      return mk('stats', Object.assign({ title: titleText, stats }, rest[0] ? { insight: one(rest[0]) } : {}), s.notes);
    }
    if (/단계|로드맵|프로세스|절차|step|roadmap|process|timeline/i.test(titleText) && cols.length >= 3 && cols.length <= 5)
      return mk('process', { title: titleText, current: -1, steps: cols.map(c => { const t = c.items.map(one); return { title: t[0] || '', desc: t[1] || '', when: t[2] || '' }; }) }, s.notes);
    if (/비교|vs\.?|as-?is|to-?be|전후|before|after/i.test(titleText + ' ' + bodyText.slice(0, 60))) {
      const L = bodies.filter(b => b.x + b.w / 2 < W / 2).map(one).filter(t => !/^[—→\-]$/.test(t)), Rr = bodies.filter(b => b.x + b.w / 2 >= W / 2).map(one).filter(t => !/^[—→\-]$/.test(t));
      if (L.length >= 2 && Rr.length >= 2) { const side = a => { const lab = a[0].length <= 8 && a.length > 2 ? a.shift() : ''; return { label: lab, title: a.shift(), items: a }; }; return mk('compare', { title: titleText, left: side(L), right: side(Rr) }, s.notes); }
    }
    const leadShape = bodies.find(b => b.y < H * 0.25 && b.w > W * 0.5);
    const cardPairs = pairs(bodies.filter(b => b !== leadShape), one);
    if (!pic && cardPairs.length >= 2 && cardPairs.length <= 4 && cardPairs.every(p => p.title.length <= 30 && p.desc)) return mk('cards', { title: titleText, lead: leadShape ? one(leadShape) : '', cards: cardPairs }, s.notes);
    if (/구역|섹션|section/i.test(s.layout) && bodyText.length < 80 && !pic) return mk('section', { number: N2(++secN), title: titleText, desc: bodyText.replace(/\n/g, ' ') }, s.notes);
    if (pic && bodyText.length < 160) return mk('photo', Object.assign({ title: titleText, text: bodyText.replace(new RegExp(SOFT, 'g'), ' '), caption: '' }, img), s.notes);
    const bullets = [];
    bodies.forEach(b => b.paragraphs.forEach(p => { const t = p.text.replace(new RegExp(SOFT, 'g'), ' ').trim(); if (t) bullets.push((p.level > 1 ? '>' : '') + t); }));
    return mk('bullets', Object.assign({ title: titleText, bullets }, pic ? img : {}), s.notes);
  });
  if (replaceDeck(slides, slides[0] ? slides[0].template.data.title : imp.name, imp.name + ' (새 양식)')) toast(`${slides.length}장을 편집기 양식으로 옮겼어요. 왼쪽 '디자인'에서 색과 양식을 고르세요.`);
}

// ---- 1) edit the original deck in place ----
function slideEdit(i) { return imp.edits[i] || (imp.edits[i] = { shapes: {} }); }
function isModified(i) { const e = imp.edits[i]; return !!e && (e.delete || e.notes !== undefined || Object.keys(e.shapes).length > 0); }
function editableShapes(s) { return s.shapes.filter(x => (x.kind === 'text' && ![13, 16].includes(x.ph)) || (x.kind === 'picture' && x.image) || x.kind === 'table').sort((a, b) => (a.y - b.y) || (a.x - b.x)); }
function openOrig() { $('orig').hidden = false; $('origName').textContent = imp.name; $('origInfo').textContent = `${imp.info.slideCount}장 · 원본 디자인 유지`; $('origFile').value = imp.name + ' (수정본)'; renderOrig(); }
function renderOrig() {
  const list = $('origList'); list.innerHTML = '';
  imp.info.slides.forEach((s, i) => {
    const e = imp.edits[s.index] || {};
    const d = document.createElement('div'); d.className = 'it' + (i === imp.sel ? ' sel' : '') + (e.delete ? ' del' : '');
    d.innerHTML = `<img src="${impUrl('png', s.png)}"><span class="tag">${s.index}${s.hidden ? ' · 숨김' : ''}</span>${e.delete ? '<span class="tag mod">삭제</span>' : isModified(s.index) ? '<span class="tag mod">수정</span>' : ''}`;
    d.onclick = () => { imp.sel = i; imp.shape = null; renderOrig(); };
    list.appendChild(d);
  });
  renderOrigStage(); renderOrigForm();
}
function fitOrig() { const st = $('origStage'), c = $('origCanvas'), r = imp.info.width / imp.info.height; const w = Math.max(200, st.clientWidth - 40), h = Math.max(120, st.clientHeight - 40); const cw = Math.min(w, h * r); c.style.width = cw + 'px'; c.style.height = cw / r + 'px'; }
function renderOrigStage() {
  const s = imp.info.slides[imp.sel], W = imp.info.width, H = imp.info.height, e = imp.edits[s.index] || { shapes: {} };
  $('origCanvas').innerHTML = `<img src="${impUrl('png', s.png)}">` + editableShapes(s).map(x => `<div class="hit${imp.shape === x.key ? ' on' : ''}${e.shapes[x.key] ? ' mod' : ''}" data-key="${x.key}" style="left:${x.x / W * 100}%;top:${x.y / H * 100}%;width:${x.w / W * 100}%;height:${x.h / H * 100}%"></div>`).join('');
  $('origCanvas').querySelectorAll('.hit').forEach(h => h.onclick = () => { imp.shape = h.dataset.key; renderOrigStage(); renderOrigForm(true); });
  fitOrig();
}
function renderOrigForm(scroll) {
  const s = imp.info.slides[imp.sel], e = slideEdit(s.index), f = $('origForm');
  f.innerHTML = `<h3 class="sec">${s.index}번 슬라이드 <label class="row" style="gap:4px;font-weight:400"><input type="checkbox" id="oDel" ${e.delete ? 'checked' : ''}> 이 슬라이드 삭제</label></h3><div class="hint">미리보기에서 영역을 누르면 해당 칸으로 이동해요. ${SOFT} 표시는 줄바꿈(Shift+Enter)이에요.</div>`;
  $('oDel').onchange = ev => { e.delete = ev.target.checked; renderOrig(); };
  let tn = 0;
  for (const x of editableShapes(s)) {
    const fe = document.createElement('div'); fe.className = 'field' + (imp.shape === x.key ? ' on' : ''); fe.dataset.key = x.key; const curE = e.shapes[x.key] || {};
    if (x.kind === 'text') {
      const label = (x.ph === 1 || x.ph === 3) ? '제목' : x.ph === 4 ? '부제목' : x.ph === 15 ? '바닥글' : `텍스트 ${++tn}`, orig = shapeText(x);
      fe.innerHTML = `<div class="lbl"><span>${label}${x.inGroup ? ' (그룹 안)' : ''}</span><span>${curE.paragraphs ? '수정됨' : ''}</span></div><textarea rows="${Math.min(8, Math.max(2, (x.paragraphs || []).length + 1))}"></textarea>`;
      const ta = fe.querySelector('textarea'); ta.value = curE.paragraphs ? curE.paragraphs.join('\n') : orig;
      ta.oninput = () => { if (ta.value === orig) delete e.shapes[x.key]; else e.shapes[x.key] = { paragraphs: ta.value.split('\n') }; fe.querySelector('.lbl span:last-child').textContent = e.shapes[x.key] ? '수정됨' : ''; renderOrigStage(); markList(); };
      ta.onfocus = () => { if (imp.shape !== x.key) { imp.shape = x.key; renderOrigStage(); f.querySelectorAll('.field').forEach(q => q.classList.toggle('on', q.dataset.key === x.key)); } };
    } else if (x.kind === 'picture') {
      fe.innerHTML = `<div class="lbl"><span>사진${x.inGroup ? ' (그룹 안 — 바꿀 수 없어요)' : ''}</span><span>${curE.image ? '바뀜' : ''}</span></div><div class="row"><img src="${curE.imageUrl || impUrl('media', x.image)}" style="width:110px;height:70px;object-fit:cover;border-radius:4px;background:#F2F3F7">${x.inGroup ? '' : `<button class="btn sm" data-a="pick">사진 바꾸기</button>${curE.image ? '<button class="btn sm" data-a="undo">되돌리기</button>' : ''}`}</div>`;
      const pk = fe.querySelector('[data-a=pick]');
      if (pk) pk.onclick = () => { const inp = $('imgFile'); inp.value = ''; inp.onchange = async () => { const fl = inp.files[0]; if (!fl) return; try { const r = await upload(fl, 'image'); e.shapes[x.key] = { image: r.path, imageUrl: r.url }; renderOrig(); } catch (err) { toast(err.message); } }; inp.click(); };
      const un = fe.querySelector('[data-a=undo]'); if (un) un.onclick = () => { delete e.shapes[x.key]; renderOrig(); };
    } else if (x.kind === 'table') {
      const data = curE.table || x.table;
      fe.innerHTML = `<div class="lbl"><span>표 (${x.table.length}행 × ${x.table[0].length}열)</span><span>${curE.table ? '수정됨' : ''}</span></div><table>${data.map((row, r) => `<tr>${row.map((v, c) => `<td><input type="text" data-r="${r}" data-c="${c}" value="${esc(v)}"></td>`).join('')}</tr>`).join('')}</table>`;
      fe.querySelectorAll('input').forEach(inp => inp.oninput = () => { const t = (e.shapes[x.key] && e.shapes[x.key].table) || x.table.map(r => r.slice()); t[+inp.dataset.r][+inp.dataset.c] = inp.value; if (JSON.stringify(t) === JSON.stringify(x.table)) delete e.shapes[x.key]; else e.shapes[x.key] = { table: t }; renderOrigStage(); markList(); });
    }
    f.appendChild(fe);
  }
  const nf = document.createElement('div'); nf.className = 'field'; nf.innerHTML = '<div class="lbl"><span>발표자 노트</span></div><textarea rows="3"></textarea>';
  const nt = nf.querySelector('textarea'); nt.value = e.notes !== undefined ? e.notes : (s.notes || ''); nt.oninput = () => { if (nt.value === (s.notes || '')) delete e.notes; else e.notes = nt.value; markList(); };
  f.appendChild(nf);
  if (scroll && imp.shape) { const t = f.querySelector(`.field[data-key="${imp.shape}"]`); if (t) { t.scrollIntoView({ block: 'center' }); const ta = t.querySelector('textarea, input'); if (ta) ta.focus({ preventScroll: true }); } }
}
function markList() { $('origList').querySelectorAll('.it').forEach((d, i) => { const idx = imp.info.slides[i].index, e = imp.edits[idx] || {}; let tag = d.querySelector('.tag.mod'); const want = e.delete ? '삭제' : isModified(idx) ? '수정' : ''; if (want && !tag) { tag = document.createElement('span'); tag.className = 'tag mod'; d.appendChild(tag); } if (tag) { if (want) tag.textContent = want; else tag.remove(); } }); }
async function saveOrig() {
  const slides = Object.entries(imp.edits).filter(([i]) => isModified(+i)).map(([i, e]) => { const o = { index: +i }; if (e.delete) { o.delete = true; return o; } if (e.notes !== undefined) o.notes = e.notes; o.shapes = Object.entries(e.shapes).map(([key, v]) => Object.assign({ key }, v.paragraphs ? { paragraphs: v.paragraphs } : {}, v.image ? { image: v.image } : {}, v.table ? { table: v.table } : {})); return o; });
  const useBrand = $('origBrand').checked;
  if (!slides.length && !useBrand) { toast('바뀐 내용이 없어요.'); return; }
  if (slides.filter(s => s.delete).length === imp.info.slideCount) { toast('모든 슬라이드를 지울 수는 없어요.'); return; }
  const m = $('modal'); $('modalBack').classList.add('show');
  m.innerHTML = '<h3><span class="spinner"></span>수정본을 만드는 중…</h3><div class="hint">원본 파일은 그대로 두고 새 파일로 저장해요.</div>';
  const body = { id: imp.id, fileName: $('origFile').value || imp.name + ' (수정본)', editsJson: JSON.stringify({ slides }) };
  if (useBrand) body.brandJson = JSON.stringify({ colors: pal(), headFont: state.brand.headFont, bodyFont: state.brand.bodyFont });
  try { const r = await fetch('/api/apply', { method: 'POST', body: JSON.stringify(body) }); const res = await r.json(); if (res.ok) showResult(res); else { m.innerHTML = `<h3>저장하지 못했어요</h3><pre class="log">${esc(res.log || res.error || '')}</pre><button class="btn" id="closeModal">닫기</button>`; $('closeModal').onclick = () => $('modalBack').classList.remove('show'); } }
  catch (e) { m.innerHTML = '<h3>편집기 서버에 연결할 수 없어요</h3><button class="btn" id="closeModal">닫기</button>'; $('closeModal').onclick = () => $('modalBack').classList.remove('show'); }
}
function wireImport() {
  $('importBtn').onclick = () => { $('pptFile').value = ''; $('pptFile').click(); };
  $('pptFile').onchange = () => { const f = $('pptFile').files[0]; if (f) importPpt(f); };
  $('origClose').onclick = () => { const changed = Object.keys(imp.edits).some(i => isModified(+i)); if (changed && !confirm('저장하지 않은 수정 내용이 있어요. 닫을까요?')) return; $('orig').hidden = true; };
  $('origSave').onclick = saveOrig; $('origToNew').onclick = convertToDeck; $('origToFree').onclick = freeImport;
  window.addEventListener('resize', () => { if (imp && !$('orig').hidden) fitOrig(); });
  document.addEventListener('dragover', e => { if ([...e.dataTransfer.items].some(i => i.kind === 'file')) e.preventDefault(); });
  document.addEventListener('drop', e => { const f = [...e.dataTransfer.files].find(x => /\.(pptx|ppt|pptm|potx)$/i.test(x.name)); if (f) { e.preventDefault(); importPpt(f); } });
}
