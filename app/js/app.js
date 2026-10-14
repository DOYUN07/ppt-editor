'use strict';
// ============================================================================
// PPT 편집기 — canvas editor (Miricanvas-style) on top of the template model in model.js
// ============================================================================
const $ = id => document.getElementById(id);
let server = { fonts: [], layouts: [], history: [], brands: null, projects: [] };
let state = null;           // { v, fileName, brand, layout, deck, assets }
let cur = 0;                // current slide index
let sel = [];               // selected element ids
let zoom = 1, zoomFit = true;
let editing = null;         // { id, el, node }
let drawerTab = 'design';
let clipboard = null;
const undoStack = [], redoStack = [];

function toast(msg) { const t = $('toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 2400); }
function store() { try { localStorage.setItem('ppt-editor-v4', JSON.stringify(state)); } catch (e) {} }
function restore() { try { const s = JSON.parse(localStorage.getItem('ppt-editor-v4') || 'null'); if (s && s.deck && s.brand && s.layout) return s; } catch (e) {} return null; }
function pal() { return derivePalette(state.brand).colors; }

// ---------------- history ----------------
function snap() { undoStack.push(JSON.stringify({ deck: state.deck, layout: state.layout })); if (undoStack.length > 100) undoStack.shift(); redoStack.length = 0; updateUndo(); }
function undo() { if (!undoStack.length) return; finishEdit(); redoStack.push(JSON.stringify({ deck: state.deck, layout: state.layout })); const s = JSON.parse(undoStack.pop()); state.deck = s.deck; state.layout = s.layout; cur = Math.min(cur, state.deck.slides.length - 1); sel = []; renderAll(); updateUndo(); }
function redo() { if (!redoStack.length) return; finishEdit(); undoStack.push(JSON.stringify({ deck: state.deck, layout: state.layout })); const s = JSON.parse(redoStack.pop()); state.deck = s.deck; state.layout = s.layout; cur = Math.min(cur, state.deck.slides.length - 1); sel = []; renderAll(); updateUndo(); }
function updateUndo() { $('undoBtn').disabled = !undoStack.length; $('redoBtn').disabled = !redoStack.length; }

// ---------------- model access ----------------
const slide = () => state.deck.slides[cur];
let resolvedCache = null;
function resolved() { if (!resolvedCache || resolvedCache.idx !== cur) resolvedCache = { idx: cur, r: resolveSlide(state, cur) }; return resolvedCache.r; }
function invalidate() { resolvedCache = null; }
function elById(id) { return resolved().elements.find(e => e.id === id); }
const isRole = id => String(id).startsWith('r:');
// Apply a property patch to an element (generated elements keep their patch as an override).
function patchEl(id, patch) {
  const s = slide();
  if (isRole(id)) {
    const role = id.slice(2); const ov = s.ov[role] || (s.ov[role] = {});
    for (const [k, v] of Object.entries(patch)) {
      if (k === 'textStyle') ov.textStyle = Object.assign({}, ov.textStyle, v);
      else if (k === 'image' || k === 'chart' || k === 'table' || k === 'icon') ov[k] = Object.assign({}, ov[k], v);
      else if (v === undefined) delete ov[k]; else ov[k] = v;
    }
  } else {
    const e = s.extra.find(x => x.id === id); if (!e) return;
    for (const [k, v] of Object.entries(patch)) {
      if (k === 'textStyle') e.text = Object.assign({}, e.text, v);
      else if (k === 'image' || k === 'chart' || k === 'table' || k === 'icon') e[k] = Object.assign({}, e[k], v);
      else if (v === undefined) delete e[k]; else e[k] = v;
    }
  }
  invalidate();
}
// Set the full text of an element; template text roles write back into the slide data.
function setText(id, tx) {
  const s = slide();
  if (!isRole(id)) { const e = s.extra.find(x => x.id === id); if (e) e.text = Object.assign({}, e.text, { paras: tx.paras }); invalidate(); return; }
  const role = id.slice(2), ov = s.ov[role] || (s.ov[role] = {}), plain = plainText(tx);
  if (role === 'frame.footer') { state.deck.footer = plain; delete ov.text; delete ov.textForce; }
  else if (s.template && isDataRole(role)) {
    if (LIST_ROLES.includes(role)) setPath(s.template.data, role, listLines(tx));
    else setPath(s.template.data, role, plain);
    if (isRich(tx)) ov.text = { paras: clone(tx.paras) }; else delete ov.text;
    delete ov.textForce;
  } else { ov.text = { paras: clone(tx.paras) }; ov.textForce = true; }
  invalidate();
}
function deleteEls(ids) {
  const s = slide();
  ids.forEach(id => {
    if (isRole(id)) { const role = id.slice(2); if (!s.del.includes(role)) s.del.push(role); }
    else s.extra = s.extra.filter(e => e.id !== id);
  });
  invalidate();
}
function addEl(el, select = true) {
  const s = slide(); const zMax = Math.max(1000, ...resolved().elements.map(e => e.z || 0));
  const e = Object.assign({ id: uid(), z: zMax + 1 }, el); s.extra.push(e); invalidate();
  if (select) sel = [e.id];
  return e;
}
const styleOf = tx => { const st = Object.assign({}, tx); delete st.paras; return st; };

// ---------------- rendering ----------------
let vcCache = null;
function vc() { return vcCache || (vcCache = viewCtx(state)); }
function renderAll() { vcCache = null; invalidate(); renderCanvas(); renderStrip(); renderCtx(); renderProps(); renderDrawer(); $('fileName').value = state.fileName || ''; store(); applyUiAccent(); }
function renderSoft() { invalidate(); renderCanvas(); renderCtx(); renderProps(); renderStripOne(cur); store(); }
function applyUiAccent() { document.documentElement.style.setProperty('--acc', '#' + pal().dk2); }
function fit() {
  const wrap = $('canvasWrap'); const w = wrap.clientWidth - 60, h = wrap.clientHeight - 60;
  if (zoomFit) zoom = Math.max(0.2, Math.min(w / SW, h / SH));
  $('canvas').style.width = SW * zoom + 'px'; $('canvas').style.height = SH * zoom + 'px';
  $('stage').style.transform = `scale(${zoom})`; $('overlay').style.transform = `scale(${zoom})`;
  $('zoomFit').textContent = Math.round(zoom * 100) + '%';
  renderOverlay();
}
function renderCanvas() {
  const r = resolved();
  $('stage').innerHTML = slideHtml(state, cur, { vc: vc(), resolved: r, hideIds: editing ? new Set([editing.id]) : new Set() });
  $('stage').querySelectorAll('.el').forEach(n => { const e = r.elements.find(x => x.id === n.dataset.id); if (e && e.locked) n.classList.add('locked'); });
  if (editing) mountEditor();
  sel = sel.filter(id => r.elements.some(e => e.id === id));
  fit();
}
function bbox(els) {
  const xs = els.flatMap(e => [e.x, e.x + e.w]), ys = els.flatMap(e => [e.y, e.y + e.h]);
  return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
}
function renderOverlay(guides) {
  const ov = $('overlay'), z = zoom, hs = 9 / z, els = sel.map(elById).filter(Boolean);
  let h = '';
  if (guides) guides.forEach(g => { h += g.x != null ? `<div class="guide" style="left:${g.x - 0.5 / z}px;top:0;width:${1 / z}px;height:${SH}px"></div>` : `<div class="guide" style="top:${g.y - 0.5 / z}px;left:0;height:${1 / z}px;width:${SW}px"></div>`; });
  if (els.length === 1 && !editing) {
    const e = els[0];
    h += `<div class="sel-box" style="left:${e.x}px;top:${e.y}px;width:${e.w}px;height:${e.h}px;border-width:${1.5 / z}px;transform:rotate(${e.rot || 0}deg)">`;
    if (!e.locked) {
      const pts = e.kind === 'line' ? [['nw', 0, 0], ['se', 1, 1]] : [['nw', 0, 0], ['n', .5, 0], ['ne', 1, 0], ['e', 1, .5], ['se', 1, 1], ['s', .5, 1], ['sw', 0, 1], ['w', 0, .5]];
      const cursor = { nw: 'nwse-resize', se: 'nwse-resize', ne: 'nesw-resize', sw: 'nesw-resize', n: 'ns-resize', s: 'ns-resize', e: 'ew-resize', w: 'ew-resize' };
      pts.forEach(([k, fx, fy]) => { h += `<div class="hdl" data-h="${k}" style="left:${fx * e.w - hs / 2}px;top:${fy * e.h - hs / 2}px;width:${hs}px;height:${hs}px;border-width:${1.5 / z}px;cursor:${cursor[k]}"></div>`; });
      if (e.kind !== 'line') h += `<div class="hdl rot" data-h="rot" style="left:${e.w / 2 - hs / 2}px;top:${-26 / z}px;width:${hs}px;height:${hs}px;border-width:${1.5 / z}px"></div>`;
    }
    h += '</div>';
  } else if (els.length > 1) {
    els.forEach(e => { h += `<div class="sel-box multi" style="left:${e.x}px;top:${e.y}px;width:${e.w}px;height:${e.h}px;border-width:${1 / z}px;transform:rotate(${e.rot || 0}deg)"></div>`; });
    const b = bbox(els); h += `<div class="sel-box" style="left:${b.x}px;top:${b.y}px;width:${b.w}px;height:${b.h}px;border-width:${1.5 / z}px"></div>`;
  }
  ov.innerHTML = h;
}
function renderStrip() {
  const strip = $('strip'); strip.innerHTML = '';
  const v = vc(), sc = 150 / SW;
  state.deck.slides.forEach((s, i) => {
    const t = document.createElement('div'); t.className = 'thumb' + (i === cur ? ' sel' : '') + (s.hidden ? ' hidden' : ''); t.draggable = true; t.dataset.i = i;
    t.innerHTML = `<div class="fr">${slideHtml(state, i, { vc: v })}</div><div class="cap"><span>${i + 1}</span><span>${esc(s.template ? TYPES[s.template.type] || '' : '자유')}${s.hidden ? ' · 숨김' : ''}</span></div>`;
    t.querySelector('.slide').style.transform = `scale(${sc})`;
    t.onclick = () => goSlide(i);
    t.ondragstart = e => { e.dataTransfer.setData('text/slide', String(i)); };
    t.ondragover = e => { if (e.dataTransfer.types.includes('text/slide')) { e.preventDefault(); t.classList.add('dragover'); } };
    t.ondragleave = () => t.classList.remove('dragover');
    t.ondrop = e => { t.classList.remove('dragover'); const from = +e.dataTransfer.getData('text/slide'); if (isNaN(from) || from === i) return; e.preventDefault(); snap(); const [m] = state.deck.slides.splice(from, 1); state.deck.slides.splice(from < i ? i - 1 : i, 0, m); cur = from < i ? i - 1 : i; renderAll(); };
    strip.appendChild(t);
  });
  const add = document.createElement('button'); add.className = 'add-slide'; add.innerHTML = '<b style="font-size:22px;line-height:1">＋</b>새 슬라이드';
  add.onclick = () => { openDrawer('templates'); toast('왼쪽에서 추가할 슬라이드 모양을 고르세요.'); };
  strip.appendChild(add);
}
function renderStripOne(i) {
  const t = $('strip').children[i]; if (!t) return renderStrip();
  const fr = t.querySelector('.fr'); fr.innerHTML = slideHtml(state, i, { vc: vc() }); fr.querySelector('.slide').style.transform = `scale(${150 / SW})`;
}
function goSlide(i) { finishEdit(); cur = Math.max(0, Math.min(i, state.deck.slides.length - 1)); sel = []; invalidate(); renderCanvas(); renderCtx(); renderProps(); [...$('strip').children].forEach((t, j) => t.classList.toggle('sel', j === cur)); const t = $('strip').children[cur]; if (t) t.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }

// ---------------- pointer interaction ----------------
let drag = null;
function slidePt(ev) { const r = $('canvas').getBoundingClientRect(); return { x: (ev.clientX - r.left) / zoom, y: (ev.clientY - r.top) / zoom }; }
function snapCandidates(excl) {
  const xs = [0, SW / 2, SW], ys = [0, SH / 2, SH];
  resolved().elements.forEach(e => { if (excl.includes(e.id) || e.w >= SW - 1) return; xs.push(e.x, e.x + e.w / 2, e.x + e.w); ys.push(e.y, e.y + e.h / 2, e.y + e.h); });
  return { xs, ys };
}
function snapBox(b, cand) {
  const th = 6 / zoom, guides = []; let dx = 0, dy = 0, best = th;
  for (const [v, off] of [[b.x, 0], [b.x + b.w / 2, 0], [b.x + b.w, 0]]) for (const c of cand.xs) { const d = c - v; if (Math.abs(d) < best) { best = Math.abs(d); dx = d; } }
  best = th;
  for (const v of [b.y, b.y + b.h / 2, b.y + b.h]) for (const c of cand.ys) { const d = c - v; if (Math.abs(d) < best) { best = Math.abs(d); dy = d; } }
  const nb = { x: b.x + dx, y: b.y + dy, w: b.w, h: b.h };
  if (dx || cand.xs.some(c => [nb.x, nb.x + nb.w / 2, nb.x + nb.w].some(v => Math.abs(c - v) < 0.01))) cand.xs.forEach(c => { if ([nb.x, nb.x + nb.w / 2, nb.x + nb.w].some(v => Math.abs(c - v) < 0.01)) guides.push({ x: c }); });
  if (dy || cand.ys.some(c => [nb.y, nb.y + nb.h / 2, nb.y + nb.h].some(v => Math.abs(c - v) < 0.01))) cand.ys.forEach(c => { if ([nb.y, nb.y + nb.h / 2, nb.y + nb.h].some(v => Math.abs(c - v) < 0.01)) guides.push({ y: c }); });
  return { dx, dy, guides };
}
function liveEl(e) { const n = $('stage').querySelector(`.el[data-id="${CSS.escape(e.id)}"]`); if (n) n.outerHTML = elementHtml(vc(), e, { page: cur + 1 }); }
function onCanvasDown(ev) {
  if (ev.button !== 0) return;
  const p = slidePt(ev), hd = ev.target.closest('.hdl');
  if (editing) { if (ev.target.closest('.editing')) return; finishEdit(); }
  if (hd && sel.length === 1) {
    const e = elById(sel[0]); snap();
    drag = { mode: hd.dataset.h === 'rot' ? 'rot' : 'resize', h: hd.dataset.h, start: p, orig: clone(e), el: clone(e) };
    ev.preventDefault(); return;
  }
  const node = ev.target.closest('.stage .el');
  if (node) {
    const id = node.dataset.id, e = elById(id);
    if (ev.shiftKey || ev.ctrlKey) { sel = sel.includes(id) ? sel.filter(x => x !== id) : sel.concat(id); renderOverlay(); renderCtx(); renderProps(); return; }
    if (!sel.includes(id)) sel = [id];
    renderOverlay(); renderCtx(); renderProps();
    const movable = sel.map(elById).filter(x => x && !x.locked);
    if (!movable.length) return;
    drag = { mode: 'move', start: p, origs: movable.map(clone), moved: false, cand: snapCandidates(sel) };
    ev.preventDefault(); return;
  }
  if (!ev.shiftKey) sel = [];
  drag = { mode: 'marquee', start: p, base: sel.slice() };
  renderOverlay(); renderCtx(); renderProps();
}
function onMove(ev) {
  if (!drag) return;
  const p = slidePt(ev), dx = p.x - drag.start.x, dy = p.y - drag.start.y;
  if (drag.mode === 'move') {
    if (!drag.moved && Math.hypot(dx, dy) < 2 / zoom) return;
    if (!drag.moved) { snap(); drag.moved = true; }
    const b = bbox(drag.origs.map(o => Object.assign({}, o, { x: o.x + dx, y: o.y + dy })));
    const s = ev.altKey ? { dx: 0, dy: 0, guides: [] } : snapBox(b, drag.cand);
    drag.origs.forEach(o => { const e = Object.assign({}, o, { x: Math.round((o.x + dx + s.dx) * 10) / 10, y: Math.round((o.y + dy + s.dy) * 10) / 10 }); liveEl(e); drag.last = drag.last || {}; drag.last[o.id] = e; });
    renderOverlayLive(Object.values(drag.last || {}), s.guides);
  } else if (drag.mode === 'resize') {
    const o = drag.orig, h = drag.h; let { x, y, w, h: hh } = o;
    const keep = ev.shiftKey !== (['image', 'icon'].includes(o.kind) && h.length === 2);
    if (h.includes('e')) w = Math.max(4, o.w + dx); if (h.includes('s')) hh = Math.max(4, o.h + dy);
    if (h.includes('w')) { w = Math.max(4, o.w - dx); x = o.x + o.w - w; } if (h.includes('n')) { hh = Math.max(4, o.h - dy); y = o.y + o.h - hh; }
    if (o.kind === 'line') { w = o.w + (h.includes('e') ? dx : 0) - (h.includes('w') ? dx : 0); hh = o.h + (h.includes('s') ? dy : 0) - (h.includes('n') ? dy : 0); x = h.includes('w') ? o.x + dx : o.x; y = h.includes('n') ? o.y + dy : o.y; }
    if (keep && h.length === 2 && o.kind !== 'line') { const r = o.w / o.h; if (w / hh > r) w = hh * r; else hh = w / r; if (h.includes('w')) x = o.x + o.w - w; if (h.includes('n')) y = o.y + o.h - hh; }
    const e = Object.assign({}, o, { x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10, w: Math.round(w * 10) / 10, h: Math.round(hh * 10) / 10 });
    if (e.kind === 'table') { /* rows keep proportion */ }
    drag.el = e; liveEl(e); renderOverlayLive([e]);
  } else if (drag.mode === 'rot') {
    const o = drag.orig, cx = o.x + o.w / 2, cy = o.y + o.h / 2;
    let a = Math.atan2(p.y - cy, p.x - cx) * 180 / Math.PI + 90; if (ev.shiftKey) a = Math.round(a / 15) * 15; else if (Math.abs(((a % 90) + 90) % 90) < 3) a = Math.round(a / 90) * 90;
    a = ((Math.round(a) % 360) + 360) % 360; const e = Object.assign({}, o, { rot: a }); drag.el = e; liveEl(e); renderOverlayLive([e]);
  } else if (drag.mode === 'marquee') {
    const r = { x: Math.min(p.x, drag.start.x), y: Math.min(p.y, drag.start.y), w: Math.abs(dx), h: Math.abs(dy) };
    drag.rect = r;
    const hit = resolved().elements.filter(e => !(e.x > r.x + r.w || e.x + e.w < r.x || e.y > r.y + r.h || e.y + e.h < r.y)).map(e => e.id);
    sel = Array.from(new Set(drag.base.concat(hit)));
    renderOverlay(); $('overlay').insertAdjacentHTML('beforeend', `<div class="marquee" style="left:${r.x}px;top:${r.y}px;width:${r.w}px;height:${r.h}px;border-width:${1 / zoom}px"></div>`);
  }
}
function renderOverlayLive(els, guides) {
  const save = sel; const m = new Map(els.map(e => [e.id, e]));
  const orig = elById; elById = id => m.get(id) || orig(id); renderOverlay(guides); elById = orig; sel = save;
}
function onUp() {
  if (!drag) return;
  const d = drag; drag = null;
  if (d.mode === 'move' && d.moved && d.last) { Object.values(d.last).forEach(e => patchEl(e.id, { x: e.x, y: e.y })); renderSoft(); }
  else if ((d.mode === 'resize' || d.mode === 'rot') && d.el) { const e = d.el; patchEl(e.id, d.mode === 'rot' ? { rot: e.rot } : { x: e.x, y: e.y, w: e.w, h: e.h }); renderSoft(); }
  else if (d.mode === 'marquee') { renderOverlay(); renderCtx(); renderProps(); }
  else if (d.mode === 'resize' || d.mode === 'rot') { undoStack.pop(); updateUndo(); }
}
function onDbl(ev) {
  const node = ev.target.closest('.stage .el'); if (!node) return;
  const e = elById(node.dataset.id); if (!e || e.locked) return;
  const role = isRole(e.id) ? e.id.slice(2) : '';
  if (role === 'imageText' || role === 'imageBox' || e.kind === 'image') { pickImage(e); return; }
  if (e.kind === 'text' || (e.kind === 'shape')) startEdit(e);
  else if (e.kind === 'table' || e.kind === 'chart') { renderProps(); toast('오른쪽 패널에서 내용을 고치세요.'); }
}

// ---------------- text editing (contenteditable) ----------------
function startEdit(e) {
  if (e.text && e.text.field === 'page') { toast('쪽번호는 자동으로 들어가요.'); return; }
  if (!e.text) {
    snap();
    const tx = textObj('', { size: 18, color: e.fill === 'lt1' || !e.fill ? 'dk1' : 'lt1', align: 'center', valign: 'middle' });
    if (isRole(e.id)) { const ov = slide().ov[e.id.slice(2)] || (slide().ov[e.id.slice(2)] = {}); ov.text = { paras: tx.paras }; ov.textForce = true; ov.textStyle = styleOf(tx); }
    else { const x = slide().extra.find(q => q.id === e.id); x.text = tx; }
    invalidate(); e = elById(e.id);
  }
  editing = { id: e.id, before: JSON.stringify(e.text) };
  sel = [e.id]; renderCanvas(); renderCtx();
}
function mountEditor() {
  const e = elById(editing.id); if (!e) { editing = null; return; }
  const v = vc(), isShape = e.kind === 'shape';
  const wrap = document.createElement('div'); wrap.className = 'editing';
  wrap.style.cssText = `left:${e.x}px;top:${e.y}px;width:${e.w}px;height:${e.h}px;transform:rotate(${e.rot || 0}deg)`;
  const html = textHtml(v, e.text, { pad: isShape ? [4, 7, 4, 7] : undefined, page: cur + 1, editing: true });
  wrap.innerHTML = html;
  const tx = wrap.querySelector('.tx'); tx.contentEditable = 'true'; tx.spellcheck = false; tx.style.outline = 'none'; tx.style.minHeight = '100%';
  tx.querySelectorAll('.bul').forEach(b => b.contentEditable = 'false');
  $('stage').querySelector('.slide').appendChild(wrap);
  editing.node = tx; editing.wrap = wrap;
  tx.addEventListener('keydown', ev => {
    if (ev.key === 'Escape') { ev.preventDefault(); finishEdit(); }
    if (ev.key === 'Tab') { ev.preventDefault(); indentSel(ev.shiftKey ? -1 : 1); }
    ev.stopPropagation();
  });
  tx.addEventListener('input', () => fixBullets(tx));
  setTimeout(() => { tx.focus(); const r = document.createRange(); r.selectNodeContents(tx); r.collapse(false); const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); }, 0);
}
function fixBullets(tx) {
  // keep numbering markers in sync and make sure every line is a paragraph div
  [...tx.childNodes].forEach(n => { if (n.nodeType === 3 && n.textContent.trim()) { const d = document.createElement('div'); d.className = 'p'; tx.insertBefore(d, n); d.appendChild(n); } });
  let num = 0;
  tx.querySelectorAll(':scope > div').forEach(p => { p.classList.add('p'); const b = p.dataset.bullet || 'none'; let mk = p.querySelector(':scope > .bul');
    if (b === 'num') { num++; if (mk) mk.textContent = num + '.'; } else num = 0; });
}
function parseEditor(tx, base) {
  const pv = pal(), hexToToken = h => { const x = hx(h); const k = Object.keys(pv).find(t => pv[t] === x && THEME_TOKENS.includes(t)); return k || '#' + x; };
  const rgb2hex = c => { const m = String(c).match(/\d+(\.\d+)?/g); return m ? toHex(m.slice(0, 3).map(Number)) : null; };
  const baseColorHex = (base.color && pv[base.color]) ? pv[base.color] : hx(base.color || '000000');
  const paras = [];
  const blocks = [...tx.childNodes].filter(n => n.nodeType === 1 && n.matches('div') || (n.nodeType === 3 && n.textContent));
  (blocks.length ? blocks : [tx]).forEach(b => {
    const p = { runs: [] };
    if (b.nodeType === 1 && b !== tx) {
      const lvl = +(b.dataset.level || 1), bul = b.dataset.bullet || 'none', al = b.style.textAlign;
      if (lvl > 1) p.level = lvl; if (bul !== 'none') p.bullet = bul; if (bul === 'char') p.bulletChar = b.dataset.bchar || '•'; if (al && al !== (base.align || 'left')) p.align = al;
      const bp = (base.paras || [])[paras.length]; if (bp) { if (bp.spaceAfter != null) p.spaceAfter = bp.spaceAfter; if (bp.spaceBefore != null) p.spaceBefore = bp.spaceBefore; }
    }
    const walk = (n, st) => {
      if (n.nodeType === 3) { if (n.textContent) p.runs.push(Object.assign({ t: n.textContent.replace(/​/g, '') }, st)); return; }
      if (n.nodeType !== 1 || n.classList.contains('bul')) return;
      if (n.tagName === 'BR') { p.runs.push(Object.assign({ t: '\n' }, st)); return; }
      const s2 = Object.assign({}, st), cs = n.style;
      if (n.tagName === 'B' || n.tagName === 'STRONG' || /^(bold|[6-9]00)$/.test(cs.fontWeight)) s2.b = true; else if (cs.fontWeight && /^(normal|[1-4]00)$/.test(cs.fontWeight)) s2.b = false;
      if (n.tagName === 'I' || n.tagName === 'EM' || cs.fontStyle === 'italic') s2.i = true;
      const deco = (cs.textDecorationLine || cs.textDecoration || '') + ' ' + (n.tagName === 'U' ? 'underline' : '') + ' ' + (n.tagName === 'S' || n.tagName === 'STRIKE' ? 'line-through' : '');
      if (/underline/.test(deco) && !n.closest('[data-link]')) s2.u = true; if (/line-through/.test(deco)) s2.s = true;
      if (n.dataset && n.dataset.c) s2.color = n.dataset.c; else if (cs.color) { const h = rgb2hex(cs.color); if (h) s2.color = hexToToken(h); } else if (n.tagName === 'FONT' && n.getAttribute('color')) s2.color = hexToToken(n.getAttribute('color'));
      if (cs.fontSize && /px$/.test(cs.fontSize)) s2.size = Math.round(parseFloat(cs.fontSize) * 10) / 10;
      if (cs.fontFamily) { const f = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(); s2.font = f === state.brand.headFont ? 'head' : f === state.brand.bodyFont ? 'body' : f; }
      if (n.dataset && n.dataset.link) s2.link = n.dataset.link;
      n.childNodes.forEach(c => walk(c, s2));
    };
    (b.nodeType === 3 ? [b] : [...b.childNodes]).forEach(c => walk(c, {}));
    // drop props equal to the element defaults and merge neighbours
    p.runs = p.runs.map(r => { const o = { t: r.t };
      if (r.b != null && r.b !== !!base.bold) o.b = r.b; if (r.i) o.i = true; if (r.u) o.u = true; if (r.s) o.s = true;
      if (r.color && hx(r.color.replace('#', '')) !== baseColorHex && r.color !== base.color) o.color = r.color;
      if (r.size && Math.abs(r.size - (base.size || 16)) > 0.05) o.size = r.size; if (r.font && r.font !== (base.font || 'body')) o.font = r.font;
      if (r.link) o.link = r.link; return o; })
      .reduce((acc, r) => { const last = acc[acc.length - 1]; const key = o => JSON.stringify(Object.assign({}, o, { t: '' })); if (last && key(last) === key(r)) last.t += r.t; else acc.push(r); return acc; }, []);
    if (!p.runs.length) p.runs = [{ t: '' }];
    p.runs.forEach(r => { r.t = r.t.replace(/\n$/, ''); });
    paras.push(p);
  });
  return paras.length ? paras : [{ runs: [{ t: '' }] }];
}
function finishEdit() {
  if (!editing) return;
  const ed = editing; editing = null;
  const e = elById(ed.id);
  if (e && ed.node) {
    const paras = parseEditor(ed.node, e.text);
    const tx = Object.assign({}, e.text, { paras });
    if (JSON.stringify(tx) !== ed.before) { snap(); setText(ed.id, tx); }
  }
  invalidate(); renderCanvas(); renderCtx(); renderProps(); renderStripOne(cur); store();
}
function editorSelParas() {
  const s = window.getSelection(); if (!editing || !s.rangeCount) return [];
  const tx = editing.node, r = s.getRangeAt(0);
  return [...tx.querySelectorAll(':scope > div')].filter(p => r.intersectsNode(p));
}
function indentSel(d) {
  if (editing) { editorSelParas().forEach(p => { const l = Math.max(1, Math.min(5, (+(p.dataset.level || 1)) + d)); p.dataset.level = l; const b = p.dataset.bullet || 'none'; p.style.paddingLeft = ((l - 1) * 22 + (b !== 'none' ? 18 : 0)) + 'px'; const mk = p.querySelector('.bul'); if (mk) mk.style.left = (l - 1) * 22 + 'px'; }); return; }
  eachTextEl(tx => tx.paras.forEach(p => { p.level = Math.max(1, Math.min(5, (p.level || 1) + d)); }));
}
// Apply a paragraph-level change to the whole selected element(s) (when not in text editing)
function eachTextEl(fn) {
  const els = sel.map(elById).filter(e => e && e.text); if (!els.length) return;
  snap(); els.forEach(e => { const tx = clone(e.text); fn(tx); setText(e.id, tx); patchEl(e.id, { textStyle: styleOf(tx) }); }); renderSoft();
}
// Character-level style: inside the editor apply to the selection, otherwise to the whole element.
function styleText(prop, value) {
  if (editing) {
    const tx = editing.node; tx.focus();
    document.execCommand('styleWithCSS', false, true);
    if (prop === 'b') document.execCommand('bold'); else if (prop === 'i') document.execCommand('italic'); else if (prop === 'u') document.execCommand('underline'); else if (prop === 's') document.execCommand('strikeThrough');
    else if (prop === 'color') { document.execCommand('foreColor', false, '#' + (pal()[value] || hx(value))); }
    else if (prop === 'size') { document.execCommand('fontSize', false, '7'); tx.querySelectorAll('font[size="7"], span[style*="xxx-large"]').forEach(f => { const sp = document.createElement('span'); sp.style.fontSize = value + 'px'; sp.innerHTML = f.innerHTML; f.replaceWith(sp); }); }
    else if (prop === 'font') { document.execCommand('fontName', false, value === 'head' ? state.brand.headFont : value === 'body' ? state.brand.bodyFont : value); }
    else if (prop === 'link') { if (value) document.execCommand('createLink', false, value); tx.querySelectorAll('a[href]').forEach(a => { const sp = document.createElement('span'); sp.dataset.link = a.getAttribute('href'); sp.innerHTML = a.innerHTML; a.replaceWith(sp); }); }
    return;
  }
  const els = sel.map(elById).filter(e => e && e.text); if (!els.length) return;
  snap();
  els.forEach(e => {
    const st = {}, k = { b: 'bold', i: 'italic' }[prop];
    if (k) st[k] = !(e.text[k]); else if (prop === 'color') st.color = value; else if (prop === 'size') st.size = value; else if (prop === 'font') st.font = value;
    patchEl(e.id, { textStyle: st });
    // clear conflicting run-level overrides so the element-wide style shows everywhere
    const runKey = { b: 'b', i: 'i', color: 'color', size: 'size', font: 'font' }[prop];
    if (prop === 'u' || prop === 's') { const tx = clone(elById(e.id).text); const on = !tx.paras.every(p => p.runs.every(r => r[prop])); tx.paras.forEach(p => p.runs.forEach(r => { if (on) r[prop] = true; else delete r[prop]; })); setText(e.id, tx); }
    else if (runKey) { const tx = clone(elById(e.id).text); let changed = false; tx.paras.forEach(p => p.runs.forEach(r => { if (runKey in r) { delete r[runKey]; changed = true; } })); if (changed) setText(e.id, tx); }
  });
  renderSoft();
}
function paraStyle(prop, value) {
  if (editing) {
    const ps = editorSelParas();
    ps.forEach(p => {
      if (prop === 'align') p.style.textAlign = value;
      if (prop === 'bullet') { p.dataset.bullet = value; let mk = p.querySelector(':scope > .bul'); if (value === 'none') { if (mk) mk.remove(); } else { if (!mk) { mk = document.createElement('span'); mk.className = 'bul'; mk.contentEditable = 'false'; p.prepend(mk); } mk.textContent = value === 'disc' ? '●' : value === 'square' ? '■' : '1.'; mk.style.fontSize = value === 'num' ? '1em' : '0.62em'; }
        const l = +(p.dataset.level || 1); p.style.paddingLeft = ((l - 1) * 22 + (value !== 'none' ? 18 : 0)) + 'px'; }
    });
    fixBullets(editing.node); return;
  }
  if (prop === 'align') { const els = sel.map(elById).filter(e => e && e.text); snap(); els.forEach(e => { patchEl(e.id, { textStyle: { align: value } }); const tx = clone(elById(e.id).text); tx.paras.forEach(p => delete p.align); setText(e.id, tx); }); renderSoft(); return; }
  if (prop === 'bullet') eachTextEl(tx => tx.paras.forEach(p => { if (value === 'none') delete p.bullet; else p.bullet = value; }));
}

// ---------------- keyboard ----------------
function onKey(ev) {
  if (!$('show').hidden) return;
  const tag = (ev.target.tagName || '').toLowerCase();
  if (editing || tag === 'input' || tag === 'textarea' || tag === 'select' || ev.target.isContentEditable) return;
  const k = ev.key.toLowerCase(), ctrl = ev.ctrlKey || ev.metaKey;
  if (ctrl && k === 'z') { ev.preventDefault(); ev.shiftKey ? redo() : undo(); return; }
  if (ctrl && k === 'y') { ev.preventDefault(); redo(); return; }
  if (ev.key === 'F5') { ev.preventDefault(); present(ev.shiftKey ? cur : 0); return; }
  if (ctrl && k === 'a') { ev.preventDefault(); sel = resolved().elements.map(e => e.id); renderOverlay(); renderCtx(); renderProps(); return; }
  if (ctrl && k === 'c' && sel.length) { clipboard = sel.map(elById).filter(Boolean).map(stripRole); toast(`${clipboard.length}개 복사했어요.`); return; }
  if (ctrl && k === 'x' && sel.length) { clipboard = sel.map(elById).filter(Boolean).map(stripRole); snap(); deleteEls(sel); sel = []; renderSoft(); return; }
  if (ctrl && k === 'v' && clipboard) { ev.preventDefault(); pasteEls(clipboard, 16); return; }
  if (ctrl && k === 'd' && sel.length) { ev.preventDefault(); duplicateSel(); return; }
  if (!sel.length) { if (ev.key === 'PageDown' || (ev.key === 'ArrowDown' && !ctrl)) { goSlide(cur + 1); ev.preventDefault(); } if (ev.key === 'PageUp' || ev.key === 'ArrowUp') { goSlide(cur - 1); ev.preventDefault(); } return; }
  if (ev.key === 'Delete' || ev.key === 'Backspace') { ev.preventDefault(); snap(); deleteEls(sel); sel = []; renderSoft(); return; }
  if (ev.key === 'Escape') { sel = []; renderOverlay(); renderCtx(); renderProps(); return; }
  if (ev.key === 'Enter' && sel.length === 1) { const e = elById(sel[0]); if (e && (e.kind === 'text' || e.kind === 'shape')) { ev.preventDefault(); startEdit(e); } return; }
  const step = ev.shiftKey ? 10 : 1, mv = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[ev.key];
  if (mv) { ev.preventDefault(); if (!ev.repeat) snap(); sel.map(elById).filter(e => e && !e.locked).forEach(e => patchEl(e.id, { x: e.x + mv[0], y: e.y + mv[1] })); renderSoft(); }
}
function stripRole(e) { const o = clone(e); delete o.role; delete o.isTitle; o.id = uid(); return o; }
function pasteEls(list, off) { snap(); const zMax = Math.max(1000, ...resolved().elements.map(e => e.z || 0)); sel = list.map((e, i) => { const n = Object.assign(clone(e), { id: uid(), x: e.x + off, y: e.y + off, z: zMax + 1 + i }); slide().extra.push(n); return n.id; }); invalidate(); renderSoft(); }
function duplicateSel() { const list = sel.map(elById).filter(Boolean).map(stripRole); pasteEls(list, 16); }

// ---------------- color picker popover ----------------
function colorButton(current, onPick, opts = {}) {
  const b = document.createElement('button'); b.className = 'swatch-btn' + (current ? '' : ' none'); b.title = opts.title || '색';
  if (current) b.style.background = colorOf(vc(), current);
  b.onclick = ev => { ev.stopPropagation(); openColorPop(b, current, onPick, opts); };
  return b;
}
function openColorPop(anchor, current, onPick, opts) {
  closePops();
  const pv = pal(), pop = document.createElement('div'); pop.className = 'cpop'; pop.id = 'cpop';
  const tokens = THEME_TOKENS;
  const recent = (state.recentColors || []).slice(0, 8);
  pop.innerHTML = `<div class="hint">테마 색 (색 조합을 바꾸면 함께 바뀌어요)</div><div class="sws">${tokens.map(t => `<div class="sw${current === t ? ' sel' : ''}" data-c="${t}" title="${TOKEN_NAMES[t]}" style="background:#${pv[t]}"></div>`).join('')}</div>
    <div class="hint">기본 색</div><div class="sws">${['#000000', '#404040', '#7F7F7F', '#D9D9D9', '#FFFFFF', '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0', '#FF66CC'].map(c => `<div class="sw${current === c ? ' sel' : ''}" data-c="${c}" style="background:${c}"></div>`).join('')}</div>
    ${recent.length ? `<div class="hint">최근 색</div><div class="sws">${recent.map(c => `<div class="sw" data-c="${c}" style="background:${c}"></div>`).join('')}</div>` : ''}
    <div class="row" style="margin-top:6px"><input type="color" id="cpopCustom" value="${current && current.startsWith('#') ? current : '#' + (pv[current] || '333333')}"> <span class="hint">직접 고르기</span>${opts.allowNone ? '<span class="grow"></span><button class="btn sm" data-none="1">없음</button>' : ''}</div>`;
  document.body.appendChild(pop);
  const r = anchor.getBoundingClientRect(); pop.style.left = Math.min(window.innerWidth - 250, r.left) + 'px'; pop.style.top = (r.bottom + 6) + 'px';
  pop.onclick = ev => { ev.stopPropagation(); const sw = ev.target.closest('.sw'); if (sw) { onPick(sw.dataset.c); closePops(); } if (ev.target.dataset.none) { onPick(null); closePops(); } };
  $('cpopCustom').onchange = ev => { const c = ev.target.value.toUpperCase(); state.recentColors = [c].concat((state.recentColors || []).filter(x => x !== c)).slice(0, 8); onPick(c); closePops(); };
}
function closePops() { const p = $('cpop'); if (p) p.remove(); document.querySelectorAll('.menu.open').forEach(m => m.classList.remove('open')); }

// ---------------- context toolbar ----------------
function tbtn(label, title, fn, on) { const b = document.createElement('button'); b.className = 'btn ic sm' + (on ? ' on' : ''); b.innerHTML = label; b.title = title; b.onmousedown = ev => ev.preventDefault(); b.onclick = fn; return b; }
function tsep() { const s = document.createElement('span'); s.className = 'sep'; return s; }
function renderCtx() {
  const bar = $('ctxbar'); bar.innerHTML = '';
  const els = sel.map(elById).filter(Boolean);
  if (!els.length && !editing) {
    bar.append(Object.assign(document.createElement('span'), { className: 'lbl', textContent: `${cur + 1} / ${state.deck.slides.length} · ${slide().template ? TYPES[slide().template.type] : '자유 슬라이드'}` }));
    bar.append(tsep(), tbtn('＋ 글상자', '글상자 추가', () => insertText('body')), tbtn('▭ 도형', '도형 추가', () => openDrawer('elements')), tbtn('🖼 사진', '사진 추가', () => openDrawer('photos')), tbtn('▦ 표', '표 추가', () => insertTable(3, 3)), tbtn('📊 차트', '차트 추가', () => insertChart('column')));
    bar.append(tsep(), tbtn('양식 추천', '배치와 도형을 새로 추천', recommendAll));
    return;
  }
  const e0 = els[0] || elById(editing && editing.id), textEls = els.filter(e => e.text || e.kind === 'shape' || e.kind === 'text');
  const hasText = editing || textEls.length;
  if (hasText) {
    const tx = (e0 && e0.text) || {};
    const fsel = document.createElement('select'); fsel.title = '글꼴'; fsel.onmousedown = ev => { if (editing) saveRange(); };
    fsel.innerHTML = `<option value="">글꼴</option><option value="head">제목 글꼴 (${esc(state.brand.headFont)})</option><option value="body">본문 글꼴 (${esc(state.brand.bodyFont)})</option>` + (server.fonts || []).filter(f => /[가-힣]|Pretendard|Nanum|Noto|Malgun|Gothic|Myeong|Arial|Calibri|Segoe|Georgia|Times|Verdana|Tahoma|Inter|Roboto|Montserrat/.test(f)).slice(0, 120).map(f => `<option>${esc(f)}</option>`).join('');
    fsel.value = tx.font && ['head', 'body'].includes(tx.font) ? tx.font : ''; fsel.onchange = () => { restoreRange(); styleText('font', fsel.value); };
    const size = document.createElement('input'); size.type = 'number'; size.min = 6; size.max = 200; size.value = tx.size || 16; size.title = '글자 크기';
    size.onmousedown = () => { if (editing) saveRange(); };
    size.onchange = () => { restoreRange(); styleText('size', +size.value); };
    bar.append(fsel, size, tbtn('－', '작게', () => { const v = Math.max(6, (+size.value) - 2); size.value = v; styleText('size', v); }), tbtn('＋', '크게', () => { const v = (+size.value) + 2; size.value = v; styleText('size', v); }));
    bar.append(tbtn('<b>B</b>', '굵게 (Ctrl+B)', () => styleText('b'), !editing && tx.bold), tbtn('<i>I</i>', '기울임 (Ctrl+I)', () => styleText('i'), !editing && tx.italic), tbtn('<u>U</u>', '밑줄 (Ctrl+U)', () => styleText('u')), tbtn('<s>S</s>', '취소선', () => styleText('s')));
    const cb = colorButton(tx.color || 'dk1', c => { restoreRange(); styleText('color', c); }, { title: '글자 색' }); cb.onmousedown = () => { if (editing) saveRange(); };
    bar.append(cb, tsep());
    const al = tx.align || 'left';
    bar.append(tbtn('⯇', '왼쪽 정렬', () => paraStyle('align', 'left'), al === 'left'), tbtn('≡', '가운데 정렬', () => paraStyle('align', 'center'), al === 'center'), tbtn('⯈', '오른쪽 정렬', () => paraStyle('align', 'right'), al === 'right'), tbtn('☰', '양쪽 정렬', () => paraStyle('align', 'justify'), al === 'justify'));
    if (!editing) { const va = tx.valign || 'top'; bar.append(tbtn('⤒', '위쪽 맞춤', () => setStyleAll({ valign: 'top' }), va === 'top'), tbtn('⇕', '가운데 맞춤', () => setStyleAll({ valign: 'middle' }), va === 'middle'), tbtn('⤓', '아래쪽 맞춤', () => setStyleAll({ valign: 'bottom' }), va === 'bottom')); }
    bar.append(tsep(), tbtn('•', '글머리 기호', () => paraStyle('bullet', 'disc')), tbtn('1.', '번호 매기기', () => paraStyle('bullet', 'num')), tbtn('✕•', '기호 없애기', () => paraStyle('bullet', 'none')), tbtn('⇤', '내어쓰기 (Shift+Tab)', () => indentSel(-1)), tbtn('⇥', '들여쓰기 (Tab)', () => indentSel(1)));
    const lsel = document.createElement('select'); lsel.title = '줄 간격'; lsel.innerHTML = [1, 1.1, 1.2, 1.35, 1.5, 2].map(v => `<option value="${v}">줄 ${v}</option>`).join(''); lsel.value = String(tx.lineSpacing || 1.1); lsel.onchange = () => setStyleAll({ lineSpacing: +lsel.value });
    bar.append(lsel, tbtn('🔗', '링크', () => { const u = prompt('연결할 주소 (https://…)', 'https://'); if (u) { restoreRange(); styleText('link', u); } }));
    if (editing) { bar.append(tsep(), tbtn('완료', '편집 끝내기 (Esc)', finishEdit)); return; }
    bar.append(tsep());
  }
  if (els.length === 1) {
    const e = e0;
    if (e.kind === 'shape' || e.kind === 'text') bar.append(Object.assign(document.createElement('span'), { className: 'lbl', textContent: '채우기' }), colorButton(e.fill, c => { snap(); patchEl(e.id, { fill: c || null }); renderSoft(); }, { allowNone: true, title: '채우기 색' }));
    if (['shape', 'text', 'image', 'line'].includes(e.kind)) {
      bar.append(Object.assign(document.createElement('span'), { className: 'lbl', textContent: e.kind === 'line' ? '선' : '테두리' }), colorButton(e.line && e.line.color, c => { snap(); patchEl(e.id, { line: c ? Object.assign({ width: e.kind === 'line' ? 2 : 1.5 }, e.line, { color: c }) : null }); renderSoft(); }, { allowNone: e.kind !== 'line', title: '선 색' }));
      const lw = document.createElement('select'); lw.title = '선 굵기'; lw.innerHTML = [0.5, 1, 1.5, 2, 3, 4, 6, 8].map(v => `<option value="${v}">${v}pt</option>`).join(''); lw.value = String((e.line && e.line.width) || (e.kind === 'line' ? 2 : 1)); lw.onchange = () => { snap(); patchEl(e.id, { line: Object.assign({ color: 'dk1' }, e.line, { width: +lw.value }) }); renderSoft(); };
      const ld = document.createElement('select'); ld.title = '선 모양'; ld.innerHTML = [['solid', '실선'], ['dash', '파선'], ['dot', '점선'], ['dashDot', '일점쇄선'], ['longDash', '긴 파선']].map(([v, t]) => `<option value="${v}">${t}</option>`).join(''); ld.value = (e.line && e.line.dash) || 'solid'; ld.onchange = () => { snap(); patchEl(e.id, { line: Object.assign({ color: 'dk1', width: 1.5 }, e.line, { dash: ld.value }) }); renderSoft(); };
      bar.append(lw, ld);
    }
    if (e.kind === 'line') bar.append(tbtn('←', '시작 화살표', () => { snap(); patchEl(e.id, { line: Object.assign({}, e.line, { start: e.line && e.line.start === 'arrow' ? 'none' : 'arrow' }) }); renderSoft(); }, e.line && e.line.start === 'arrow'), tbtn('→', '끝 화살표', () => { snap(); patchEl(e.id, { line: Object.assign({}, e.line, { end: e.line && e.line.end === 'arrow' ? 'none' : 'arrow' }) }); renderSoft(); }, e.line && e.line.end === 'arrow'));
    if (e.kind === 'image') bar.append(tbtn('사진 바꾸기', '사진 바꾸기', () => pickImage(e)), tbtn('배경 지우기', '사진 배경 지우기', () => openBgRemove(e)));
    if (e.kind === 'icon') bar.append(Object.assign(document.createElement('span'), { className: 'lbl', textContent: '색' }), colorButton(e.icon.color || 'dk2', c => { snap(); patchEl(e.id, { icon: { color: c } }); renderSoft(); }));
    bar.append(tsep());
  }
  bar.append(tbtn('⧉', '복제 (Ctrl+D)', duplicateSel), tbtn('🗑', '삭제 (Delete)', () => { snap(); deleteEls(sel); sel = []; renderSoft(); }));
  const arr = document.createElement('div'); arr.className = 'menu';
  arr.innerHTML = `<button class="btn sm">정렬·순서 ▾</button><div class="pop" style="left:0;right:auto">
    <div class="mi" data-a="front">맨 앞으로</div><div class="mi" data-a="fwd">앞으로</div><div class="mi" data-a="bwd">뒤로</div><div class="mi" data-a="back">맨 뒤로</div><hr style="border:none;border-top:1px solid var(--line)">
    <div class="mi" data-a="al">왼쪽 맞춤</div><div class="mi" data-a="ac">가운데 맞춤</div><div class="mi" data-a="ar">오른쪽 맞춤</div><div class="mi" data-a="at">위쪽 맞춤</div><div class="mi" data-a="am">중간 맞춤</div><div class="mi" data-a="ab">아래쪽 맞춤</div>
    <div class="mi" data-a="dh">가로 간격 똑같이</div><div class="mi" data-a="dv">세로 간격 똑같이</div></div>`;
  arr.querySelector('button').onclick = ev => { ev.stopPropagation(); const open = arr.classList.contains('open'); closePops(); if (!open) arr.classList.add('open'); };
  arr.querySelector('.pop').onclick = ev => { const a = ev.target.dataset.a; if (a) { arrange(a); closePops(); } };
  bar.append(arr);
  const lk = els.every(e => e.locked);
  bar.append(tbtn(lk ? '🔒' : '🔓', lk ? '잠금 해제' : '잠그기 (움직이지 않게)', () => { snap(); els.forEach(e => patchEl(e.id, { locked: lk ? undefined : true })); renderSoft(); }, lk));
}
let savedRange = null;
function saveRange() { const s = window.getSelection(); if (s.rangeCount) savedRange = s.getRangeAt(0).cloneRange(); }
function restoreRange() { if (editing && savedRange) { editing.node.focus(); const s = window.getSelection(); s.removeAllRanges(); s.addRange(savedRange); } }
function setStyleAll(st) { const els = sel.map(elById).filter(e => e && e.text); if (!els.length) return; snap(); els.forEach(e => patchEl(e.id, { textStyle: st })); renderSoft(); }
function arrange(a) {
  const els = sel.map(elById).filter(Boolean); if (!els.length) return; snap();
  const all = resolved().elements, zs = all.map(e => e.z || 0), zMin = Math.min(...zs), zMax = Math.max(...zs);
  if (a === 'front') els.forEach((e, i) => patchEl(e.id, { z: zMax + 1 + i }));
  if (a === 'back') els.forEach((e, i) => patchEl(e.id, { z: zMin - els.length + i }));
  if (a === 'fwd' || a === 'bwd') els.forEach(e => { const idx = all.findIndex(x => x.id === e.id), o = all[idx + (a === 'fwd' ? 1 : -1)]; if (o) { const z1 = e.z, z2 = o.z; patchEl(e.id, { z: z2 }); patchEl(o.id, { z: z1 }); } });
  const b = els.length > 1 ? bbox(els) : { x: 0, y: 0, w: SW, h: SH };
  if (a === 'al') els.forEach(e => patchEl(e.id, { x: b.x })); if (a === 'ar') els.forEach(e => patchEl(e.id, { x: b.x + b.w - e.w })); if (a === 'ac') els.forEach(e => patchEl(e.id, { x: b.x + (b.w - e.w) / 2 }));
  if (a === 'at') els.forEach(e => patchEl(e.id, { y: b.y })); if (a === 'ab') els.forEach(e => patchEl(e.id, { y: b.y + b.h - e.h })); if (a === 'am') els.forEach(e => patchEl(e.id, { y: b.y + (b.h - e.h) / 2 }));
  if ((a === 'dh' || a === 'dv') && els.length > 2) {
    const k = a === 'dh' ? 'x' : 'y', d = a === 'dh' ? 'w' : 'h', s = els.slice().sort((p, q) => p[k] - q[k]);
    const total = s.reduce((t, e) => t + e[d], 0), span = s[s.length - 1][k] + s[s.length - 1][d] - s[0][k], gap = (span - total) / (s.length - 1); let pos = s[0][k];
    s.forEach(e => { patchEl(e.id, { [k]: Math.round(pos * 10) / 10 }); pos += e[d] + gap; });
  }
  renderSoft();
}

// ---------------- insert ----------------
function centerBox(w, h) { return { x: Math.round((SW - w) / 2), y: Math.round((SH - h) / 2), w, h }; }
const TEXT_PRESETS = {
  title: ['제목 추가', { size: 40, bold: true, font: 'head', color: 'dk2' }, 600, 70],
  subtitle: ['부제목 추가', { size: 26, bold: true, color: 'dk2' }, 520, 46],
  body: ['본문 텍스트를 입력하세요', { size: 16, color: 'dk1', ls: 1.3 }, 420, 60],
  caption: ['캡션 · 출처', { size: 11, color: 'accent6' }, 300, 24],
  number: ['120%', { size: 64, bold: true, color: 'accent1' }, 300, 90],
  quote: ['“인상 깊은 문장이나 고객의 목소리를 넣어 보세요.”', { size: 22, italic: true, color: 'dk2', ls: 1.3 }, 560, 90],
  bullets: ['첫 번째 항목\n두 번째 항목\n세 번째 항목', { size: 18, list: 'disc', bulletColor: 'accent1', spaceAfter: 8, ls: 1.2 }, 420, 120],
  numbered: ['첫 번째 단계\n두 번째 단계\n세 번째 단계', { size: 18, list: 'num', spaceAfter: 8, ls: 1.2 }, 420, 120]
};
function insertText(kind) {
  const p = TEXT_PRESETS[kind]; snap();
  const tx = textObj(p[0], p[1]); if (p[1].italic) tx.italic = true;
  addEl(Object.assign({ kind: 'text', text: tx }, centerBox(p[2], p[3]))); renderSoft();
}
function insertShape(shape) { snap(); addEl(Object.assign({ kind: 'shape', shape, fill: 'accent1', radius: shape === 'roundRect' ? 0.15 : undefined }, centerBox(shape.startsWith('arrow') || shape === 'chevron' || shape === 'homePlate' ? 220 : 160, shape.startsWith('arrow') && shape !== 'arrowUp' && shape !== 'arrowDown' ? 90 : 160))); renderSoft(); }
function insertLine(kind) { snap(); const ln = { color: 'dk2', width: 2, dash: kind === 'dash' ? 'dash' : 'solid', end: kind === 'arrow' || kind === 'double' ? 'arrow' : 'none', start: kind === 'double' ? 'arrow' : 'none' }; addEl({ kind: 'line', x: 330, y: 270, w: 300, h: 0, line: ln }); renderSoft(); }
function insertIcon(name) { snap(); addEl(Object.assign({ kind: 'icon', icon: { name, color: 'accent1', stroke: 1.8 } }, centerBox(96, 96))); renderSoft(); }
function insertImage(asset) { snap(); const ar = asset.aspect || 1.5; let w = 480, h = w / ar; if (h > 360) { h = 360; w = h * ar; } addEl(Object.assign({ kind: 'image', image: { src: asset.path, url: asset.url, fit: 'cover' } }, centerBox(Math.round(w), Math.round(h)))); renderSoft(); }
function insertTable(r, c) { snap(); const rows = [Array.from({ length: c }, (_, i) => '머리글 ' + (i + 1))].concat(Array.from({ length: r - 1 }, () => Array.from({ length: c }, () => ''))); addEl(Object.assign({ kind: 'table', table: { rows, header: true, fontSize: 14, headerFill: 'dk2', headerColor: 'lt1', bandFill: 'lt2', bodyFill: 'lt1', color: 'dk1', banded: true, firstColBold: false, colW: Array.from({ length: c }, () => 720 / c) } }, centerBox(720, 44 * r))); renderSoft(); }
const CHART_SAMPLES = { categories: ['1분기', '2분기', '3분기', '4분기'], series: [{ name: '올해', values: [42, 58, 66, 80] }] };
function insertChart(type) { snap(); const ch = Object.assign({ type, showValues: true, highlight: -1 }, clone(CHART_SAMPLES)); if (/Stacked|line|area|radar/.test(type)) ch.series.push({ name: '작년', values: [30, 41, 52, 61] }); addEl(Object.assign({ kind: 'chart', chart: ch }, centerBox(560, 320))); renderSoft(); }
function addSlide(type) { snap(); const s = newSlide(type); if (type === 'section') s.template.data.number = N2(state.deck.slides.filter(x => x.template && x.template.type === 'section').length + 1); state.deck.slides.splice(cur + 1, 0, s); cur++; sel = []; renderAll(); }
function changeSlideType(type) {
  const s = slide(); snap();
  const g = s.template ? toGeneric(s.template.type, s.template.data) : { title: '', items: [], text: '' };
  s.template = { type, data: fromGeneric(type, g) }; s.ov = {}; s.del = [];
  sel = []; renderAll(); toast(`'${TYPES[type]}' 모양으로 바꿨어요. 직접 추가한 요소는 그대로예요.`);
}
function detachTemplate() {
  const s = slide(); if (!s.template) return; snap();
  const r = resolved();
  s.extra = r.elements.map((e, i) => { const o = clone(e); delete o.role; o.id = uid(); o.z = i; return o; });
  s.bg = r.bg; s.template = null; s.ov = {}; s.del = [];
  sel = []; renderAll(); toast('자유 슬라이드로 바꿨어요. 이제 양식을 바꿔도 이 슬라이드는 그대로 유지돼요.');
}
function recommendAll() {
  snap();
  state.layout = recommendLayout(state.layout, server.history);
  // a new layout moves things: drop position overrides of generated elements (text, colors and formatting stay)
  state.deck.slides.forEach(s => Object.values(s.ov || {}).forEach(o => GEOM.forEach(k => delete o[k])));
  renderAll(); toast(`'${state.layout.name}' — 배치와 도형을 새로 바꿨어요. 직접 추가한 요소는 그대로예요.`);
}

// ---------------- images ----------------
async function upload(file, kind) {
  const r = await fetch('/api/upload?kind=' + kind + '&name=' + encodeURIComponent(file.name), { method: 'POST', body: file });
  const res = await r.json(); if (!r.ok || res.error) throw new Error(res.error || '업로드 실패');
  if (kind === 'image') { state.assets = state.assets || []; if (!state.assets.some(a => a.path === res.path)) state.assets.unshift({ path: res.path, url: res.url, name: res.name, aspect: res.aspect }); }
  return res;
}
function pickImage(target) {
  const inp = $('imgFile'); inp.value = '';
  inp.onchange = async () => { const f = inp.files[0]; if (f) await placeImageFile(f, target); };
  inp.click();
}
async function placeImageFile(f, target) {
  try {
    const res = await upload(f, 'image'); snap();
    if (target && target.id) {
      const role = isRole(target.id) ? target.id.slice(2) : '';
      if (role === 'image' || role === 'imageText' || role === 'imageBox') { const d = slide().template.data; d.image = res.path; d.imageUrl = res.url; invalidate(); }
      else if (target.kind === 'image') patchEl(target.id, { image: { src: res.path, url: res.url } });
      else if (target.kind === 'shape') { addEl({ kind: 'image', x: target.x, y: target.y, w: target.w, h: target.h, shape: target.shape, radius: target.radius, image: { src: res.path, url: res.url, fit: 'cover' }, z: (target.z || 0) + 0.5 }); deleteEls([target.id]); }
    } else insertImage(res);
    renderSoft(); renderDrawer(); toast('사진을 넣었어요.');
  } catch (e) { toast(e.message); }
}

// ---------------- properties panel ----------------
function field(label, el) { const w = document.createElement('div'); const l = document.createElement('label'); l.className = 'f'; l.textContent = label; w.append(l, el); return w; }
function numInput(v, onChange, step) { const i = document.createElement('input'); i.type = 'number'; i.value = Math.round(v * 10) / 10; i.step = step || 1; i.style.width = '100%'; i.onchange = () => onChange(+i.value); return i; }
function renderProps() {
  const p = $('props'); p.innerHTML = '';
  const els = sel.map(elById).filter(Boolean);
  if (!els.length) return slideProps(p);
  if (els.length > 1) { p.innerHTML = `<h3 class="sec">${els.length}개 선택됨</h3><div class="hint">위쪽 '정렬·순서'로 맞추거나 간격을 똑같이 할 수 있어요. 끌면 함께 움직여요.</div>`; return; }
  const e = els[0], role = isRole(e.id) ? e.id.slice(2) : '';
  const kindName = { text: '글상자', shape: '도형', line: '선', image: '사진', icon: '아이콘', table: '표', chart: '차트' }[e.kind];
  p.insertAdjacentHTML('beforeend', `<h3 class="sec">${kindName}${role ? ' <span class="hint">양식 요소</span>' : ''}</h3>`);
  const g = document.createElement('div'); g.className = 'grid2';
  const geo = (k, lbl) => g.append(field(lbl, numInput(e[k] || 0, v => { snap(); patchEl(e.id, { [k]: v }); renderSoft(); })));
  geo('x', '가로 위치'); geo('y', '세로 위치'); geo('w', '너비'); geo('h', '높이'); geo('rot', '회전(도)');
  const op = document.createElement('input'); op.type = 'range'; op.min = 0.1; op.max = 1; op.step = 0.05; op.value = e.opacity ?? 1; op.onchange = () => { snap(); patchEl(e.id, { opacity: +op.value >= 1 ? undefined : +op.value }); renderSoft(); };
  g.append(field('투명도', op)); p.append(g);
  const row = document.createElement('div'); row.className = 'row'; row.style.marginTop = '8px';
  row.append(tbtn('↔', '좌우 뒤집기', () => { snap(); patchEl(e.id, { flipH: e.flipH ? undefined : true }); renderSoft(); }, e.flipH), tbtn('↕', '상하 뒤집기', () => { snap(); patchEl(e.id, { flipV: e.flipV ? undefined : true }); renderSoft(); }, e.flipV));
  if (e.kind !== 'line' && e.kind !== 'table') row.append(tbtn('그림자', '그림자', () => { snap(); patchEl(e.id, { shadow: e.shadow ? undefined : true }); renderSoft(); }, e.shadow));
  p.append(row);
  if (e.kind === 'shape' || (e.kind === 'image')) {
    const ss = document.createElement('select'); ss.innerHTML = (e.kind === 'image' ? [['rect', '사각형'], ['roundRect', '둥근 사각형'], ['ellipse', '원'], ['hexagon', '육각형'], ['diamond', '마름모'], ['triangle', '삼각형'], ['star5', '별'], ['heart', '하트'], ['cloud', '구름']] : SHAPE_LIST).map(([v, t]) => `<option value="${v}">${t}</option>`).join('');
    ss.value = e.shape || 'rect'; ss.onchange = () => { snap(); patchEl(e.id, { shape: ss.value }); renderSoft(); };
    p.append(field(e.kind === 'image' ? '사진 모양' : '도형 모양', ss));
    if ((e.shape || 'rect') === 'roundRect' || e.shape === 'snipRect' || e.shape === 'round2Same' || e.shape === 'callout') { const rr = document.createElement('input'); rr.type = 'range'; rr.min = 0; rr.max = 0.5; rr.step = 0.02; rr.value = e.radius ?? 0.12; rr.onchange = () => { snap(); patchEl(e.id, { radius: +rr.value }); renderSoft(); }; p.append(field('모서리 둥글기', rr)); }
  }
  if (e.kind === 'image') {
    const fitSel = document.createElement('select'); fitSel.innerHTML = '<option value="cover">채우기 (잘라서 맞춤)</option><option value="contain">전체 보이기</option><option value="stretch">늘이기</option>'; fitSel.value = e.image.fit || 'cover'; fitSel.onchange = () => { snap(); patchEl(e.id, { image: { fit: fitSel.value } }); renderSoft(); };
    p.append(field('맞춤 방식', fitSel));
    if ((e.image.fit || 'cover') === 'cover') { const g2 = document.createElement('div'); g2.className = 'grid2'; const fx = document.createElement('input'); fx.type = 'range'; fx.min = 0; fx.max = 1; fx.step = 0.05; fx.value = e.image.focusX ?? 0.5; fx.onchange = () => { snap(); patchEl(e.id, { image: { focusX: +fx.value } }); renderSoft(); }; const fy = fx.cloneNode(); fy.value = e.image.focusY ?? 0.5; fy.onchange = () => { snap(); patchEl(e.id, { image: { focusY: +fy.value } }); renderSoft(); }; g2.append(field('보이는 위치 (좌우)', fx), field('보이는 위치 (상하)', fy)); p.append(g2); }
    const b = document.createElement('button'); b.className = 'btn sm'; b.textContent = '사진 바꾸기'; b.style.marginTop = '8px'; b.onclick = () => pickImage(e); p.append(b);
    const bb = document.createElement('button'); bb.className = 'btn sm'; bb.textContent = '배경 지우기'; bb.style.margin = '8px 0 0 6px'; bb.onclick = () => openBgRemove(e); p.append(bb);
  }
  if (e.kind === 'icon') {
    p.insertAdjacentHTML('beforeend', '<label class="f">아이콘 바꾸기</label>');
    const grid = document.createElement('div'); grid.className = 'icon-grid'; grid.innerHTML = Object.keys(ICONS).map(k => `<button data-i="${k}" title="${ICONS[k][0]}">${iconSvg(k, '#' + pal().dk2)}</button>`).join('');
    grid.onclick = ev => { const b = ev.target.closest('button'); if (b) { snap(); patchEl(e.id, { icon: { name: b.dataset.i } }); renderSoft(); } };
    const sw = document.createElement('input'); sw.type = 'range'; sw.min = 0.8; sw.max = 3; sw.step = 0.1; sw.value = e.icon.stroke || 1.8; sw.onchange = () => { snap(); patchEl(e.id, { icon: { stroke: +sw.value } }); renderSoft(); };
    p.append(field('선 굵기', sw), grid);
  }
  if (e.kind === 'table') tableProps(p, e);
  if (e.kind === 'chart') chartProps(p, e);
  animProps(p, e);
  if (role) p.insertAdjacentHTML('beforeend', `<div class="hint" style="margin-top:14px">양식이 만든 요소예요. 글자는 그대로 두고 배치를 바꾸려면 '양식 추천'을, 이 슬라이드만 완전히 자유롭게 고치려면 슬라이드 설정의 '자유 슬라이드로 바꾸기'를 쓰세요.</div>`);
}
function tableProps(p, e) {
  const t = e.table, role = isRole(e.id);
  p.insertAdjacentHTML('beforeend', '<h3 class="sec">표 내용</h3><div class="hint">칸을 눌러 바로 고치세요.</div>');
  const tbl = document.createElement('table'); tbl.className = 'data-grid';
  tbl.innerHTML = t.rows.map((r, ri) => '<tr>' + r.map((v, ci) => `<td><input type="text" data-r="${ri}" data-c="${ci}" value="${esc(v)}"></td>`).join('') + '</tr>').join('');
  tbl.onchange = ev => { const i = ev.target; const rows = clone(t.rows); rows[+i.dataset.r][+i.dataset.c] = i.value; snap(); setTableRows(e, rows); renderSoft(); };
  p.append(tbl);
  const row = document.createElement('div'); row.className = 'row'; row.style.marginTop = '6px';
  const mod = fn => () => { const rows = clone(t.rows); fn(rows); snap(); setTableRows(e, rows); renderSoft(); };
  row.append(tbtn('행＋', '행 추가', mod(rows => rows.push(rows[0].map(() => '')))), tbtn('행－', '마지막 행 삭제', mod(rows => { if (rows.length > 2) rows.pop(); })), tbtn('열＋', '열 추가', mod(rows => rows.forEach((r, i) => r.push(i ? '' : '새 열')))), tbtn('열－', '마지막 열 삭제', mod(rows => { if (rows[0].length > 1) rows.forEach(r => r.pop()); })));
  p.append(row);
  const g = document.createElement('div'); g.className = 'grid2';
  g.append(field('머리글 색', colorButton(t.headerFill || 'dk2', c => { snap(); patchEl(e.id, { table: { headerFill: c } }); renderSoft(); })), field('줄무늬 색', colorButton(t.bandFill || 'lt2', c => { snap(); patchEl(e.id, { table: { bandFill: c } }); renderSoft(); })));
  g.append(field('글자 크기', numInput(t.fontSize || 14, v => { snap(); patchEl(e.id, { table: { fontSize: v } }); renderSoft(); })));
  const al = document.createElement('select'); al.innerHTML = '<option value="left">왼쪽</option><option value="center">가운데</option><option value="right">오른쪽</option>'; al.value = t.align || 'left'; al.onchange = () => { snap(); patchEl(e.id, { table: { align: al.value } }); renderSoft(); };
  g.append(field('글자 정렬', al)); p.append(g);
}
function setTableRows(e, rows) {
  const s = slide();
  if (isRole(e.id) && s.template && s.template.type === 'table') { s.template.data.headers = rows[0]; s.template.data.rows = rows.slice(1); invalidate(); return; }
  const nc = rows[0].length, cw = Array.from({ length: nc }, () => e.w / nc);
  patchEl(e.id, { table: { rows, colW: (e.table.colW && e.table.colW.length === nc) ? e.table.colW : cw }, h: Math.max(e.h, 0) * rows.length / e.table.rows.length });
}
function chartProps(p, e) {
  const c = e.chart;
  p.insertAdjacentHTML('beforeend', '<h3 class="sec">차트</h3>');
  const ty = document.createElement('select'); ty.innerHTML = [['column', '세로 막대'], ['columnStacked', '누적 세로 막대'], ['bar', '가로 막대'], ['barStacked', '누적 가로 막대'], ['line', '꺾은선'], ['area', '영역'], ['pie', '원형'], ['doughnut', '도넛'], ['radar', '방사형']].map(([v, t]) => `<option value="${v}">${t}</option>`).join('');
  ty.value = c.type; ty.onchange = () => { snap(); setChart(e, { type: ty.value }); renderSoft(); }; p.append(field('종류', ty));
  p.insertAdjacentHTML('beforeend', '<label class="f">데이터 (첫 행: 계열 이름, 첫 열: 항목)</label>');
  const grid = [[' '].concat(c.series.map(s => s.name))].concat(c.categories.map((ct, i) => [ct].concat(c.series.map(s => (s.values || [])[i] ?? ''))));
  const tbl = document.createElement('table'); tbl.className = 'data-grid';
  tbl.innerHTML = grid.map((r, ri) => '<tr>' + r.map((v, ci) => `<td><input type="text" data-r="${ri}" data-c="${ci}" value="${esc(v)}" ${ri === 0 && ci === 0 ? 'disabled' : ''}></td>`).join('') + '</tr>').join('');
  const read = () => { const g2 = [...tbl.rows].map(r => [...r.cells].map(td => td.querySelector('input').value)); return { categories: g2.slice(1).map(r => r[0]), series: g2[0].slice(1).map((n, k) => ({ name: n, values: g2.slice(1).map(r => parseFloat(String(r[k + 1]).replace(/,/g, '')) || 0) })) }; };
  tbl.onchange = () => { snap(); setChart(e, read()); renderSoft(); };
  p.append(tbl);
  const row = document.createElement('div'); row.className = 'row'; row.style.marginTop = '6px';
  const mod = fn => () => { const d = read(); fn(d); snap(); setChart(e, d); renderSoft(); };
  row.append(tbtn('항목＋', '항목 추가', mod(d => { d.categories.push('새 항목'); d.series.forEach(s => s.values.push(0)); })), tbtn('항목－', '마지막 항목 삭제', mod(d => { if (d.categories.length > 1) { d.categories.pop(); d.series.forEach(s => s.values.pop()); } })),
    tbtn('계열＋', '계열 추가', mod(d => d.series.push({ name: '계열 ' + (d.series.length + 1), values: d.categories.map(() => 0) }))), tbtn('계열－', '마지막 계열 삭제', mod(d => { if (d.series.length > 1) d.series.pop(); })));
  p.append(row);
  const g = document.createElement('div'); g.className = 'grid2';
  const hl = numInput((c.highlight ?? -1) + 1, v => { snap(); setChart(e, { highlight: v > 0 ? v - 1 : -1 }); renderSoft(); }); g.append(field('강조할 항목 번호 (0=없음)', hl));
  const sv = document.createElement('select'); sv.innerHTML = '<option value="1">값 표시</option><option value="0">값 숨기기</option>'; sv.value = c.showValues === false ? '0' : '1'; sv.onchange = () => { snap(); setChart(e, { showValues: sv.value === '1' }); renderSoft(); }; g.append(field('데이터 레이블', sv));
  p.append(g);
}
function setChart(e, patch) {
  const s = slide();
  if (isRole(e.id) && s.template && s.template.type === 'chart') {
    const d = s.template.data;
    if (patch.type) d.chartType = patch.type; if (patch.categories) d.categories = patch.categories; if (patch.series) { d.series = patch.series; d.seriesName = patch.series[0].name; d.values = patch.series[0].values; }
    if ('highlight' in patch) d.highlight = patch.highlight; if ('showValues' in patch) patchEl(e.id, { chart: { showValues: patch.showValues } });
    invalidate(); return;
  }
  patchEl(e.id, { chart: patch });
}
const ANIM_NAMES = [['none', '없음'], ['appear', '나타내기'], ['fade', '밝기 변화'], ['fly', '날아오기'], ['float', '떠오르기'], ['wipe', '닦아내기'], ['zoom', '확대/축소'], ['split', '바운드(분할)'], ['bounce', '튀어오르기'], ['wheel', '시계 방향 회전'], ['growTurn', '커지며 회전'], ['riseUp', '솟아오르기'], ['spin', '회전(강조)'], ['teeter', '흔들기(강조)']];
function animProps(p, e) {
  const a = e.anim || {};
  p.insertAdjacentHTML('beforeend', '<h3 class="sec">애니메이션</h3>');
  const ef = document.createElement('select'); ef.innerHTML = ANIM_NAMES.map(([v, t]) => `<option value="${v}">${t}</option>`).join(''); ef.value = a.effect || 'none';
  ef.onchange = () => { snap(); const maxOrder = Math.max(0, ...resolved().elements.map(x => (x.anim && x.anim.order) || 0)); patchEl(e.id, { anim: ef.value === 'none' ? undefined : Object.assign({ trigger: 'click', dur: 0.5, order: a.order || maxOrder + 1 }, a, { effect: ef.value }) }); renderSoft(); };
  p.append(field('효과', ef));
  if (a.effect && a.effect !== 'none') {
    const g = document.createElement('div'); g.className = 'grid2';
    const tr = document.createElement('select'); tr.innerHTML = '<option value="click">클릭할 때</option><option value="with">이전과 함께</option><option value="after">이전 다음에</option>'; tr.value = a.trigger || 'click'; tr.onchange = () => { snap(); patchEl(e.id, { anim: Object.assign({}, a, { trigger: tr.value }) }); renderSoft(); };
    g.append(field('시작', tr));
    if (['fly', 'wipe'].includes(a.effect)) { const dr = document.createElement('select'); dr.innerHTML = '<option value="bottom">아래에서</option><option value="left">왼쪽에서</option><option value="right">오른쪽에서</option><option value="top">위에서</option>'; dr.value = a.dir || 'bottom'; dr.onchange = () => { snap(); patchEl(e.id, { anim: Object.assign({}, a, { dir: dr.value }) }); renderSoft(); }; g.append(field('방향', dr)); }
    g.append(field('재생 시간(초)', numInput(a.dur || 0.5, v => { snap(); patchEl(e.id, { anim: Object.assign({}, a, { dur: v }) }); renderSoft(); }, 0.1)));
    g.append(field('지연(초)', numInput(a.delay || 0, v => { snap(); patchEl(e.id, { anim: Object.assign({}, a, { delay: v }) }); renderSoft(); }, 0.1)));
    g.append(field('순서', numInput(a.order || 1, v => { snap(); patchEl(e.id, { anim: Object.assign({}, a, { order: v }) }); renderSoft(); })));
    p.append(g);
  }
}
// ---- slide panel ----
const FORMS = {
  cover: [['title', '제목', 'area'], ['subtitle', '부제목'], ['meta', '발표자 · 날짜']],
  closing: [['title', '제목', 'area'], ['subtitle', '부제목'], ['meta', '연락처']],
  section: [['number', '번호'], ['title', '섹션 제목'], ['desc', '설명', 'area']],
  agenda: [['title', '제목'], ['items', '항목 (한 줄에 하나: 제목 | 설명)', 'pairs', ['title', 'desc']]],
  bullets: [['title', '제목'], ['bullets', '글머리표 (한 줄에 하나, 앞에 > 를 붙이면 하위 항목)', 'lines'], ['imageText', '사진이 없을 때 영역 문구 (비우면 글머리표만)', 'area']],
  photo: [['title', '제목', 'area'], ['text', '설명', 'area'], ['caption', '캡션']],
  cards: [['title', '제목'], ['lead', '한 줄 요약'], ['cards', '카드 (제목 | 설명), 2~4개', 'pairs', ['title', 'desc']]],
  stats: [['title', '제목'], ['stats', '지표 (값 | 이름 | 설명), 2~4개', 'pairs', ['value', 'label', 'note']], ['insightLabel', '강조 박스 라벨'], ['insight', '강조 박스 문장 (비우면 숨김)', 'area']],
  chart: [['title', '제목'], ['caption', '차트 설명'], ['insights', '해설 (제목 | 설명), 최대 3개', 'pairs', ['title', 'desc']]],
  process: [['title', '제목'], ['steps', '단계 (제목 | 설명 | 시기), 3~5개', 'pairs', ['title', 'desc', 'when']], ['current', '강조할 단계 번호 (1부터, 0=없음)', 'num1']],
  compare: [['title', '제목'], ['left.label', '왼쪽 라벨'], ['left.title', '왼쪽 제목'], ['left.items', '왼쪽 항목 (한 줄에 하나)', 'lines'], ['right.label', '오른쪽 라벨'], ['right.title', '오른쪽 제목'], ['right.items', '오른쪽 항목 (한 줄에 하나)', 'lines']],
  table: [['title', '제목']], titleOnly: [['title', '제목']], blank: []
};
const pairsToText = (arr, keys) => (arr || []).map(o => keys.map(k => o[k] ?? '').join(' | ').replace(/( \| )+$/, '')).join('\n');
const textToPairs = (t, keys) => t.split('\n').map(l => l.trim()).filter(Boolean).map(l => { const p = l.split('|').map(s => s.trim()); const o = {}; keys.forEach((k, i) => { if (p[i]) o[k] = p[i]; }); return o; });
function slideProps(p) {
  const s = slide();
  p.insertAdjacentHTML('beforeend', `<h3 class="sec">${cur + 1}번 슬라이드</h3>`);
  const ts = document.createElement('select');
  ts.innerHTML = Object.entries(TYPES).map(([k, v]) => `<option value="${k}">${v}</option>`).join('') + '<option value="__free">자유 슬라이드 (양식 없음)</option>';
  ts.value = s.template ? s.template.type : '__free';
  ts.onchange = () => { if (ts.value === '__free') detachTemplate(); else changeSlideType(ts.value); };
  p.append(field('슬라이드 모양', ts));
  if (s.template) {
    const d = s.template.data, form = FORMS[s.template.type] || [];
    if (form.length) p.insertAdjacentHTML('beforeend', '<h3 class="sec">내용 <span class="hint">캔버스에서 글자를 더블클릭해도 고칠 수 있어요</span></h3>');
    for (const [key, label, kind, opt] of form) {
      let el;
      if (['area', 'pairs', 'lines'].includes(kind)) {
        el = document.createElement('textarea');
        el.value = kind === 'pairs' ? pairsToText(getPath(d, key), opt) : kind === 'lines' ? (getPath(d, key) || []).join('\n') : String(getPath(d, key) ?? '').replace(/\r/g, '\n');
        el.rows = kind === 'area' ? 2 : 4;
        el.onfocus = () => { el._snapped = false; };
        el.oninput = () => { if (!el._snapped) { snap(); el._snapped = true; } setPath(d, key, kind === 'pairs' ? textToPairs(el.value, opt) : kind === 'lines' ? el.value.split('\n').filter(l => l.trim()) : el.value); renderCanvasStrip(); };
      } else if (kind === 'num1') {
        el = document.createElement('input'); el.type = 'number'; el.min = 0; el.style.width = '100%'; const v = getPath(d, key); el.value = v == null || v < 0 ? 0 : Number(v) + 1;
        el.oninput = () => { snap(); const n = parseInt(el.value || '0', 10); setPath(d, key, n > 0 ? n - 1 : -1); renderCanvasStrip(); };
      } else {
        el = document.createElement('input'); el.type = 'text'; el.value = getPath(d, key) ?? '';
        el.onfocus = () => { el._snapped = false; };
        el.oninput = () => { if (!el._snapped) { snap(); el._snapped = true; } setPath(d, key, el.value); renderCanvasStrip(); };
      }
      p.append(field(label, el));
    }
    if (['bullets', 'photo', 'cover', 'closing'].includes(s.template.type)) {
      const f = document.createElement('div'); f.className = 'img-field';
      f.innerHTML = (d.imageUrl ? `<img src="${esc(d.imageUrl)}">` : '<span class="hint">사진 없음</span>') + `<span class="grow"></span><button class="btn sm" data-a="pick">${d.image ? '바꾸기' : '사진 넣기'}</button>${d.image ? '<button class="btn sm" data-a="del">빼기</button>' : ''}`;
      f.querySelector('[data-a=pick]').onclick = () => pickImage({ id: 'r:image' });
      const del = f.querySelector('[data-a=del]'); if (del) del.onclick = () => { snap(); delete d.image; delete d.imageUrl; renderSoft(); renderProps(); };
      p.append(field('사진', f));
    }
    if ((s.del || []).length || Object.keys(s.ov || {}).length) {
      const rb = document.createElement('button'); rb.className = 'btn sm'; rb.style.marginTop = '8px'; rb.textContent = '양식 요소 되살리기 (위치·삭제 초기화)';
      rb.onclick = () => { snap(); s.del = []; Object.values(s.ov).forEach(o => GEOM.forEach(k => delete o[k])); renderSoft(); };
      p.append(rb);
    }
  }
  p.insertAdjacentHTML('beforeend', '<h3 class="sec">배경</h3>');
  const r = resolved(), bgRow = document.createElement('div'); bgRow.className = 'row';
  bgRow.append(colorButton(r.bg.color || 'lt1', c => { snap(); s.bg = { color: c }; renderSoft(); }, { title: '배경 색' }));
  const bi = document.createElement('button'); bi.className = 'btn sm'; bi.textContent = '배경 사진'; bi.onclick = () => { const inp = $('imgFile'); inp.value = ''; inp.onchange = async () => { const f = inp.files[0]; if (!f) return; try { const res = await upload(f, 'image'); snap(); s.bg = { image: res.path, imageUrl: res.url }; renderSoft(); } catch (e) { toast(e.message); } }; inp.click(); };
  bgRow.append(bi);
  if (s.bg) { const rs = document.createElement('button'); rs.className = 'btn sm'; rs.textContent = '기본으로'; rs.onclick = () => { snap(); delete s.bg; renderSoft(); }; bgRow.append(rs); }
  const all = document.createElement('button'); all.className = 'btn sm'; all.textContent = '모든 슬라이드에'; all.title = '이 배경을 모든 슬라이드에 적용'; all.onclick = () => { snap(); const b = clone(r.bg); state.deck.slides.forEach(x => { x.bg = clone(b); }); renderAll(); };
  bgRow.append(all); p.append(bgRow);
  p.insertAdjacentHTML('beforeend', '<h3 class="sec">화면 전환</h3>');
  const tr = document.createElement('select');
  tr.innerHTML = [['none', '없음'], ['fade', '밝기 변화'], ['push', '밀어내기'], ['wipe', '닦아내기'], ['split', '나누기'], ['reveal', '나타내기'], ['cover', '덮기'], ['uncover', '당기기'], ['zoom', '확대/축소'], ['randomBars', '실선 무늬'], ['dissolve', '디졸브'], ['flip', '뒤집기'], ['cube', '큐브'], ['gallery', '갤러리'], ['curtains', '커튼'], ['vortex', '소용돌이'], ['ripple', '물결'], ['honeycomb', '벌집'], ['doors', '문'], ['window', '창'], ['ferris', '관람차'], ['conveyor', '컨베이어'], ['rotate', '회전'], ['orbit', '궤도'], ['fly', '날기'], ['switch', '전환'], ['pan', '이동'], ['wind', '바람'], ['prestige', '프레스티지'], ['fracture', '부서지기'], ['crush', '구기기'], ['peel', '벗기기'], ['pageCurl', '페이지 말아 넘기기'], ['airplane', '비행기'], ['origami', '종이접기'], ['random', '임의']].map(([v, t]) => `<option value="${v}">${t}</option>`).join('');
  tr.value = (s.transition && s.transition.effect) || 'none';
  tr.onchange = () => { snap(); s.transition = tr.value === 'none' ? null : { effect: tr.value, dur: (s.transition && s.transition.dur) || 0.7 }; renderSoft(); renderProps(); };
  const trRow = document.createElement('div'); trRow.className = 'grid2'; trRow.append(field('효과', tr));
  if (s.transition) trRow.append(field('시간(초)', numInput(s.transition.dur || 0.7, v => { snap(); s.transition.dur = v; store(); }, 0.1)));
  p.append(trRow);
  const ta = document.createElement('button'); ta.className = 'btn sm'; ta.textContent = '모든 슬라이드에 적용'; ta.onclick = () => { snap(); state.deck.slides.forEach(x => { x.transition = s.transition ? clone(s.transition) : null; }); toast('모든 슬라이드에 같은 화면 전환을 적용했어요.'); store(); };
  p.append(ta);
  p.insertAdjacentHTML('beforeend', '<h3 class="sec">기타</h3>');
  const hid = document.createElement('label'); hid.className = 'row'; hid.innerHTML = `<input type="checkbox" ${s.hidden ? 'checked' : ''}> 발표할 때 숨기기`; hid.querySelector('input').onchange = ev => { snap(); s.hidden = ev.target.checked; renderStrip(); store(); };
  p.append(hid);
  const nt = document.createElement('textarea'); nt.rows = 4; nt.value = s.notes || ''; nt.oninput = () => { s.notes = nt.value; store(); };
  p.append(field('발표자 노트', nt));
  const fl = document.createElement('input'); fl.type = 'text'; fl.value = state.deck.footer || ''; fl.oninput = () => { state.deck.footer = fl.value; renderCanvasStrip(); };
  p.append(field('바닥글 문구 (전체)', fl));
  const row = document.createElement('div'); row.className = 'row'; row.style.marginTop = '12px';
  row.append(tbtn('⧉ 슬라이드 복제', '슬라이드 복제', () => { snap(); const c2 = clone(s); c2.id = uid(); state.deck.slides.splice(cur + 1, 0, c2); cur++; renderAll(); }), tbtn('🗑 슬라이드 삭제', '슬라이드 삭제', () => { if (state.deck.slides.length <= 1) return toast('슬라이드가 하나뿐이에요.'); snap(); state.deck.slides.splice(cur, 1); cur = Math.max(0, cur - 1); sel = []; renderAll(); }));
  p.append(row);
}
let csTimer = null;
function renderCanvasStrip() { invalidate(); renderCanvas(); clearTimeout(csTimer); csTimer = setTimeout(() => { renderStripOne(cur); store(); }, 200); }

// ---------------- left drawer ----------------
const RAIL = [['design', '디자인', '<path d="M12 3a9 9 0 1 0 0 18c1.1 0 1.5-.8 1.5-1.6 0-1.2-1-1.6-1-2.6 0-1 .8-1.8 1.8-1.8H17a4 4 0 0 0 4-4C21 6.6 17 3 12 3z"/><circle cx="7.5" cy="11" r="1.2"/><circle cx="10.5" cy="7" r="1.2"/><circle cx="15.5" cy="7.5" r="1.2"/>'],
  ['templates', '템플릿', '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M9 9v11"/>'], ['elements', '요소', '<circle cx="7.5" cy="7.5" r="4"/><rect x="13" y="13" width="8" height="8" rx="1.5"/><path d="M16.5 3l4 7h-8z"/>'],
  ['text', '텍스트', '<path d="M5 5h14M12 5v14M9 19h6"/>'], ['photos', '사진', '<rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-9 8"/>'],
  ['data', '표·차트', '<path d="M3 20.5h18M6.5 17V11M12 17V6M17.5 17v-4"/>'], ['bg', '배경', '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 15l5-5 5 5 3-3 5 5"/>']];
function renderRail() {
  $('rail').innerHTML = RAIL.map(([k, t, ic]) => `<button data-t="${k}" class="${drawerTab === k && !$('main').classList.contains('drawer-closed') ? 'on' : ''}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ic}</svg>${t}</button>`).join('');
  $('rail').onclick = ev => { const b = ev.target.closest('button'); if (!b) return; if (drawerTab === b.dataset.t && !$('main').classList.contains('drawer-closed')) { $('main').classList.add('drawer-closed'); renderRail(); setTimeout(fit, 0); return; } openDrawer(b.dataset.t); };
}
function openDrawer(t) { drawerTab = t; $('main').classList.remove('drawer-closed'); renderRail(); renderDrawer(); setTimeout(fit, 0); }
function renderDrawer() {
  const d = $('drawer'); d.innerHTML = '';
  if ($('main').classList.contains('drawer-closed')) return;
  ({ design: drawDesign, templates: drawTemplates, elements: drawElements, text: drawText, photos: drawPhotos, data: drawData, bg: drawBg })[drawerTab](d);
}
let tplMode = 'add';
function drawTemplates(d) {
  d.insertAdjacentHTML('beforeend', `<h3 class="sec">슬라이드 모양</h3><div class="seg" id="tplMode"><button data-m="add" class="${tplMode === 'add' ? 'on' : ''}">새 슬라이드 추가</button><button data-m="change" class="${tplMode === 'change' ? 'on' : ''}">현재 슬라이드 바꾸기</button></div><div class="hint">지금 색 조합과 양식으로 미리 보여줘요.</div><div class="tile-grid" id="tplGrid" style="margin-top:8px"></div>`);
  $('tplMode').onclick = ev => { const b = ev.target.closest('button'); if (b) { tplMode = b.dataset.m; drawTemplates(($('drawer').innerHTML = '', $('drawer'))); } };
  const g = $('tplGrid'), v = vc(), sc = 125 / SW;
  const tmpState = Object.assign({}, state, { deck: Object.assign({}, state.deck, { slides: Object.keys(TYPES).map(newSlide) }) });
  Object.keys(TYPES).forEach((k, i) => {
    const t = document.createElement('div'); t.className = 'tpl';
    t.innerHTML = `<div class="fr">${slideHtml(tmpState, i, { vc: v })}</div><div class="cap">${TYPES[k]}</div>`;
    g.appendChild(t); requestAnimationFrame(() => { const fr = t.querySelector('.fr'); t.querySelector('.slide').style.transform = `scale(${fr.clientWidth / SW})`; });
    t.onclick = () => tplMode === 'add' ? addSlide(k) : changeSlideType(k);
  });
  const free = document.createElement('div'); free.className = 'tpl'; free.innerHTML = '<div class="fr" style="display:flex;align-items:center;justify-content:center;color:#9AA1AE">빈 캔버스</div><div class="cap">자유 슬라이드</div>';
  free.onclick = () => tplMode === 'add' ? (snap(), state.deck.slides.splice(cur + 1, 0, newSlide('free')), cur++, sel = [], renderAll()) : detachTemplate();
  g.appendChild(free);
}
function drawElements(d) {
  d.insertAdjacentHTML('beforeend', '<h3 class="sec">도형</h3><div class="shape-grid" id="shGrid"></div><h3 class="sec">선</h3><div class="shape-grid" id="lnGrid"></div><h3 class="sec">아이콘 <input type="text" id="icSearch" placeholder="검색" style="width:110px;padding:3px 6px"></h3><div class="icon-grid" id="icGrid"></div>');
  const c = '#' + pal().accent1;
  $('shGrid').innerHTML = SHAPE_LIST.map(([k, t]) => `<button data-s="${k}" title="${t}"><svg viewBox="-2 -2 44 44"><path d="${shapePath(k, 40, 40, 0.18)}" fill="${c}" fill-rule="evenodd"/></svg></button>`).join('');
  $('shGrid').onclick = ev => { const b = ev.target.closest('button'); if (b) insertShape(b.dataset.s); };
  const dk = '#' + pal().dk2;
  $('lnGrid').innerHTML = [['line', '선', ''], ['arrow', '화살표', 'end'], ['double', '양쪽 화살표', 'both'], ['dash', '점선', '']].map(([k, t, a]) => `<button data-l="${k}" title="${t}"><svg viewBox="0 0 40 40"><defs><marker id="mk${k}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${dk}"/></marker></defs><line x1="5" y1="35" x2="35" y2="5" stroke="${dk}" stroke-width="3" ${k === 'dash' ? 'stroke-dasharray="5 4"' : ''} ${a ? `marker-end="url(#mk${k})"` : ''} ${a === 'both' ? `marker-start="url(#mk${k})"` : ''}/></svg></button>`).join('');
  $('lnGrid').onclick = ev => { const b = ev.target.closest('button'); if (b) insertLine(b.dataset.l); };
  const drawIcons = q => { $('icGrid').innerHTML = Object.entries(ICONS).filter(([k, v]) => !q || v[0].includes(q) || k.toLowerCase().includes(q.toLowerCase())).map(([k, v]) => `<button data-i="${k}" title="${v[0]}">${iconSvg(k, dk)}</button>`).join(''); };
  drawIcons(''); $('icSearch').oninput = ev => drawIcons(ev.target.value.trim());
  $('icGrid').onclick = ev => { const b = ev.target.closest('button'); if (b) insertIcon(b.dataset.i); };
}
function drawText(d) {
  d.insertAdjacentHTML('beforeend', '<h3 class="sec">글상자 추가</h3>');
  const pv = pal();
  const items = [['title', `<b style="font-size:22px;color:#${pv.dk2}">제목 추가</b>`], ['subtitle', `<b style="font-size:17px;color:#${pv.dk2}">부제목 추가</b>`], ['body', '<span style="font-size:13px">본문 텍스트 추가</span>'], ['caption', `<span style="font-size:11px;color:#${pv.accent6}">캡션 · 출처</span>`],
    ['number', `<b style="font-size:26px;color:#${pv.accent1}">120%</b> <span class="hint">강조 숫자</span>`], ['quote', `<i style="font-size:14px;color:#${pv.dk2}">“인용구”</i>`], ['bullets', '<span>● 글머리 기호 목록</span>'], ['numbered', '<span>1. 번호 목록</span>']];
  items.forEach(([k, h]) => { const b = document.createElement('button'); b.className = 'text-preset'; b.innerHTML = h; b.onclick = () => insertText(k); d.append(b); });
  d.insertAdjacentHTML('beforeend', '<div class="hint" style="margin-top:8px">글상자를 더블클릭하면 바로 입력할 수 있고, 위쪽 도구 막대로 글꼴·크기·색·정렬을 바꿔요. Shift+Enter는 줄바꿈, Tab은 들여쓰기예요.</div>');
}
function drawPhotos(d) {
  d.insertAdjacentHTML('beforeend', '<h3 class="sec">사진</h3><button class="btn primary" id="upBtn" style="width:100%;justify-content:center">사진 올리기</button><div class="hint">캔버스에 사진 파일을 끌어다 놓아도 돼요. 사진 요소를 더블클릭하면 바꿀 수 있어요.</div><h3 class="sec">올린 사진</h3><div class="gallery" id="gal"></div>');
  $('upBtn').onclick = () => pickImage(null);
  const as = state.assets || [];
  $('gal').innerHTML = as.length ? as.map((a, i) => `<img src="${esc(a.url)}" data-i="${i}" title="${esc(a.name || '')}">`).join('') : '<div class="hint" style="grid-column:1/-1">아직 올린 사진이 없어요.</div>';
  $('gal').onclick = ev => { const im = ev.target.closest('img'); if (!im) return; const a = as[+im.dataset.i]; const one = sel.length === 1 && elById(sel[0]); if (one && one.kind === 'image') { snap(); patchEl(one.id, { image: { src: a.path, url: a.url } }); renderSoft(); } else insertImage(a); };
}
function drawData(d) {
  d.insertAdjacentHTML('beforeend', '<h3 class="sec">표</h3><div class="hint">칸 수를 고르세요 (행 × 열)</div><div id="tblPick" style="display:grid;grid-template-columns:repeat(8,22px);gap:3px;margin:6px 0"></div><div class="hint" id="tblPickLbl">3 × 3</div><h3 class="sec">차트</h3><div class="tile-grid" id="chGrid"></div>');
  const tp = $('tblPick'); let hr = 3, hc = 3;
  const draw = () => { tp.innerHTML = Array.from({ length: 64 }, (_, i) => { const r = Math.floor(i / 8) + 1, c = i % 8 + 1; return `<div data-r="${r}" data-c="${c}" style="width:22px;height:22px;border:1px solid ${r <= hr && c <= hc ? 'var(--sel)' : 'var(--line)'};background:${r <= hr && c <= hc ? '#E8EEFF' : '#fff'};border-radius:3px;cursor:pointer"></div>`; }).join(''); $('tblPickLbl').textContent = `${hr} × ${hc}`; };
  draw(); tp.onmouseover = ev => { const c = ev.target.dataset; if (c.r) { hr = +c.r; hc = +c.c; draw(); } }; tp.onclick = () => insertTable(Math.max(2, hr), hc);
  const v = vc();
  [['column', '세로 막대'], ['columnStacked', '누적 막대'], ['bar', '가로 막대'], ['line', '꺾은선'], ['area', '영역'], ['pie', '원형'], ['doughnut', '도넛'], ['radar', '방사형']].forEach(([k, t]) => {
    const ch = Object.assign({ type: k, showValues: false }, clone(CHART_SAMPLES)); if (/Stacked|line|area|radar/.test(k)) ch.series.push({ name: '작년', values: [30, 41, 52, 61] });
    const b = document.createElement('div'); b.className = 'tpl'; b.innerHTML = `<div class="fr" style="padding:6px"><div style="position:relative;width:100%;height:100%">${chartSvg(v, { w: 120, h: 62, chart: ch })}</div></div><div class="cap">${t}</div>`;
    b.onclick = () => insertChart(k); $('chGrid').append(b);
  });
}
function drawBg(d) {
  const s = slide(), r = resolved();
  d.insertAdjacentHTML('beforeend', '<h3 class="sec">배경 색</h3><div class="hint">테마 색을 고르면 색 조합을 바꿀 때 함께 바뀌어요.</div>');
  const pv = pal(); const g = document.createElement('div'); g.className = 'grid4';
  THEME_TOKENS.forEach(t => { const b = document.createElement('button'); b.className = 'swatch-btn'; b.style.cssText = `width:100%;height:40px;background:#${pv[t]}`; b.title = TOKEN_NAMES[t]; b.onclick = () => { snap(); s.bg = { color: t }; renderSoft(); }; g.append(b); });
  d.append(g);
  const row = document.createElement('div'); row.className = 'row'; row.style.marginTop = '8px';
  row.append(colorButton(r.bg.color, c => { snap(); s.bg = { color: c }; renderSoft(); }, { title: '직접 고르기' }), Object.assign(document.createElement('span'), { className: 'hint', textContent: '다른 색' }));
  d.append(row);
  d.insertAdjacentHTML('beforeend', '<h3 class="sec">배경 사진</h3>');
  const up = document.createElement('button'); up.className = 'btn'; up.textContent = '배경 사진 올리기'; up.onclick = () => { const inp = $('imgFile'); inp.value = ''; inp.onchange = async () => { const f = inp.files[0]; if (!f) return; try { const res = await upload(f, 'image'); snap(); s.bg = { image: res.path, imageUrl: res.url }; renderSoft(); renderDrawer(); } catch (e) { toast(e.message); } }; inp.click(); };
  d.append(up);
  const as = state.assets || [];
  if (as.length) { const gal = document.createElement('div'); gal.className = 'gallery'; gal.style.marginTop = '8px'; gal.innerHTML = as.map((a, i) => `<img src="${esc(a.url)}" data-i="${i}">`).join(''); gal.onclick = ev => { const im = ev.target.closest('img'); if (!im) return; const a = as[+im.dataset.i]; snap(); s.bg = { image: a.path, imageUrl: a.url }; renderSoft(); }; d.append(gal); }
  const row2 = document.createElement('div'); row2.className = 'row'; row2.style.marginTop = '12px';
  row2.append(tbtn('기본 배경으로', '양식 기본 배경', () => { snap(); delete s.bg; renderSoft(); }), tbtn('모든 슬라이드에 적용', '모든 슬라이드에 같은 배경', () => { snap(); const b = clone(resolved().bg); state.deck.slides.forEach(x => { x.bg = clone(b); }); renderAll(); }));
  d.append(row2);
}
// ---- design (colors, fonts, logo, layout) ----
function drawDesign(d) {
  const b = state.brand, { colors, notes } = derivePalette(b);
  d.innerHTML = `<h3 class="sec">색 조합</h3><select id="brandSel"></select>
    <div class="colors" style="margin-top:8px">
      <div class="cpick"><input type="color" id="cMain" value="${b.main}"><div><div class="lbl">메인 색</div><div class="hex">${hx(b.main)}</div></div></div>
      <div class="cpick"><input type="color" id="cAccent" value="${b.accent}"><div><div class="lbl">강조 색</div><div class="hex">${hx(b.accent)}</div></div></div>
      <div class="cpick"><input type="color" id="cSub" value="${b.sub}" ${b.subAuto ? 'disabled' : ''}><div><div class="lbl">보조 색</div><div class="hex">${b.subAuto ? '자동' : hx(b.sub)}</div></div></div>
      <label class="cpick" style="cursor:pointer"><input type="checkbox" id="subAuto" ${b.subAuto ? 'checked' : ''}> <span class="lbl">보조 색 자동</span></label>
    </div>
    <div class="swatches">${['dk2', 'accent1', 'accent2', 'accent3', 'lt2', 'dk1'].map(k => `<div class="sw" title="#${colors[k]}" style="background:#${colors[k]}"></div>`).join('')}</div>
    <div class="notes">${notes.map(n => n.startsWith('ok:') ? `<div class="ok">${esc(n.slice(3))}</div>` : `<div>${esc(n)}</div>`).join('')}</div>
    <div class="grid2"><div><label class="f">제목 글꼴</label><input type="text" id="headFont" list="fontList" value="${esc(b.headFont)}"></div><div><label class="f">본문 글꼴</label><input type="text" id="bodyFont" list="fontList" value="${esc(b.bodyFont)}"></div></div>
    <datalist id="fontList">${(server.fonts || []).map(f => `<option value="${esc(f)}">`).join('')}</datalist>
    <div id="fontWarn"></div>
    <h3 class="sec">로고</h3>
    <div class="logo-prev"><div id="logoLight"></div><div id="logoDark" style="background:#${colors.dk2}"></div></div>
    <div class="row" style="margin-top:6px"><button class="btn sm" id="logoUpload">로고 파일 올리기</button>${DEFAULT_LOGO ? '<button class="btn sm" id="logoDefault">기본 로고</button>' : ''}<button class="btn sm" id="logoNone">로고 없음</button></div>
    <div class="hint">투명 PNG·SVG 권장. 투명 PNG는 어두운 배경용 흰색 로고를 자동으로 만들어요.</div>
    <div class="row" style="margin-top:10px"><button class="btn sm" id="saveBrand">이 색·로고 조합 저장</button><button class="btn sm" id="delBrand">삭제</button></div>
    <h3 class="sec">양식 <button class="btn sm primary" id="recommend">다른 양식 추천</button></h3>
    <div class="hint">현재: ${esc(state.layout.name || '사용자 지정')} · 직접 옮긴 위치는 추천할 때 초기화돼요(글자·색·추가한 요소는 유지).</div>
    <div id="knobs"></div>`;
  const sel_ = $('brandSel'); sel_.add(new Option('저장된 색 조합 선택…', '')); (server.brands || []).forEach((br, i) => sel_.add(new Option(br.name, String(i))));
  const missing = [b.headFont, b.bodyFont].filter((f, i, a) => a.indexOf(f) === i && server.fonts.length && !server.fonts.includes(f));
  $('fontWarn').innerHTML = missing.length ? `<div class="warn">이 PC에 없는 글꼴: ${esc(missing.join(', '))}</div>` : '';
  const lg = b.logo;
  $('logoLight').innerHTML = lg ? `<img src="${esc(lg.colorUrl)}">` : '로고 없음';
  $('logoDark').innerHTML = lg ? (lg.whiteUrl ? `<img src="${esc(lg.whiteUrl)}">` : `<span style="background:#fff;border-radius:6px;padding:4px 8px;display:flex"><img src="${esc(lg.colorUrl)}" style="max-height:20px"></span>`) : '<span style="color:#fff">로고 없음</span>';
  const brandChanged = () => { vcCache = null; renderAll(); };
  const setC = (k, id) => $(id).addEventListener('change', ev => { state.brand[k] = ev.target.value; brandChanged(); });
  setC('main', 'cMain'); setC('accent', 'cAccent'); setC('sub', 'cSub');
  ['cMain', 'cAccent', 'cSub'].forEach(id => $(id).addEventListener('input', ev => { const k = { cMain: 'main', cAccent: 'accent', cSub: 'sub' }[id]; state.brand[k] = ev.target.value; vcCache = null; invalidate(); renderCanvas(); applyUiAccent(); }));
  $('subAuto').onchange = ev => { state.brand.subAuto = ev.target.checked; brandChanged(); };
  $('headFont').onchange = ev => { state.brand.headFont = ev.target.value.trim() || 'Malgun Gothic'; brandChanged(); };
  $('bodyFont').onchange = ev => { state.brand.bodyFont = ev.target.value.trim() || 'Malgun Gothic'; brandChanged(); };
  $('logoUpload').onclick = () => { $('logoFile').value = ''; $('logoFile').click(); };
  if ($('logoDefault')) $('logoDefault').onclick = () => { state.brand.logo = clone(DEFAULT_LOGO); brandChanged(); };
  $('logoNone').onclick = () => { state.brand.logo = null; brandChanged(); };
  sel_.onchange = ev => { const i = ev.target.value; if (i === '') return; const br = server.brands[Number(i)]; state.brand = normBrand(clone(br)); snap(); state.layout = recommendLayout(state.layout, server.history); brandChanged(); toast(`'${br.name}' 색 조합을 적용하고 양식을 새로 추천했어요.`); };
  $('saveBrand').onclick = async () => { const name = prompt('색·로고 조합 이름', state.brand.name || '내 색 조합'); if (!name) return; const list = server.brands || []; const entry = Object.assign({}, state.brand, { name }); const i = list.findIndex(x => x.name === name); if (i >= 0) list[i] = entry; else list.push(entry); server.brands = list; state.brand.name = name; await saveBrands(); renderDrawer(); toast('저장했어요.'); };
  $('delBrand').onclick = async () => { const i = $('brandSel').value; if (i === '') { toast('삭제할 색 조합을 위 목록에서 먼저 고르세요.'); return; } const br = server.brands[Number(i)]; if (!confirm(`'${br.name}' 색 조합을 삭제할까요?`)) return; server.brands.splice(Number(i), 1); await saveBrands(); renderDrawer(); };
  $('recommend').onclick = recommendAll;
  const K = [['cover', '표지 디자인', [['blocks', '사각'], ['circles', '원형'], ['split', '분할'], ['minimal', '밝은']]], ['coverTitle', '표지 제목', [['middle', '가운데'], ['low', '아래']]], ['mirror', '배치 방향', [['false', '오른쪽'], ['true', '왼쪽']]],
    ['titleAlign', '본문 제목', [['left', '왼쪽'], ['center', '가운데']]], ['kicker', '섹션 라벨', [['false', '없음'], ['true', '표시']]], ['motif', '도형 모양', [['rounded', '둥근'], ['square', '각진'], ['circle', '원']]],
    ['contentBg', '본문 배경', [['lt1', '흰색'], ['lt2', '연한 톤']]], ['agenda', '목차', [['grid', '카드'], ['list', '리스트']]], ['cards', '카드', [['outline', '테두리'], ['filled', '채움'], ['stack', '가로 줄']]], ['stats', '숫자 지표', [['plain', '큰 숫자'], ['tiles', '타일']]], ['process', '단계', [['horizontal', '가로'], ['vertical', '세로']]]];
  const kn = $('knobs');
  K.forEach(([key, label, opts]) => { const wrap = document.createElement('div'); wrap.innerHTML = `<label class="f">${label}</label><div class="seg">${opts.map(([v, t]) => `<button data-v="${v}" class="${String(state.layout[key]) === v ? 'on' : ''}">${t}</button>`).join('')}</div>`;
    wrap.querySelector('.seg').onclick = ev => { const bt = ev.target.closest('button'); if (!bt) return; snap(); let v = bt.dataset.v; if (v === 'true') v = true; else if (v === 'false') v = false; state.layout[key] = v; if (key === 'motif') state.layout.radius = v === 'circle' ? 0.5 : v === 'square' ? 0 : 0.12; if (!String(state.layout.name).endsWith('*')) state.layout.name = (state.layout.name || '양식') + '*'; renderAll(); };
    kn.append(wrap); });
}
async function saveBrands() { await fetch('/api/brands', { method: 'POST', body: JSON.stringify(server.brands, null, 2) }); }
function normBrand(b) { if (!b.logo || b.logo === 'none') b.logo = null; return b; }

// ---------------- build / export ----------------
async function build() {
  if (build.busy) return; finishEdit(); closePops();
  build.busy = true; $('buildBtn').disabled = true;
  const m = $('modal'); $('modalBack').classList.add('show');
  m.innerHTML = '<h3><span class="spinner"></span>PowerPoint에서 만드는 중…</h3><div class="hint">슬라이드 수에 따라 1~3분 걸려요. 차트가 있으면 Excel 창이 잠깐 떴다 사라질 수 있어요.</div>';
  const pv = pal(), doc = engineDoc(state, pv);
  const style = Object.assign({}, state.layout, { name: String(state.layout.name || 'custom').replace(/\*$/, '') });
  try {
    const r = await fetch('/api/render', { method: 'POST', body: JSON.stringify({ fileName: state.fileName || '새 프레젠테이션', potx: $('optPotx').checked, pdf: $('optPdf').checked, docJson: JSON.stringify(doc), styleJson: JSON.stringify(style), brand: { dk2: pv.dk2, accent1: pv.accent1, headFont: state.brand.headFont } }) });
    const res = await r.json();
    if (res.ok) { showResult(res); try { const st = await (await fetch('/api/state')).json(); server.history = Array.isArray(st.history) ? st.history : []; } catch (e) {} }
    else { m.innerHTML = `<h3>만들지 못했어요</h3><div class="hint">PowerPoint가 응답하지 않거나 내용에 문제가 있을 수 있어요.</div><pre class="log">${esc(res.log || res.error || '')}</pre><button class="btn" id="closeModal">닫기</button>`; $('closeModal').onclick = () => $('modalBack').classList.remove('show'); }
  } catch (e) { m.innerHTML = `<h3>편집기 서버에 연결할 수 없어요</h3><div class="hint">'PPT Editor' 창이 닫혔는지 확인하고 바탕화면의 'PPT 편집기'로 다시 실행하세요.</div><button class="btn" id="closeModal">닫기</button>`; $('closeModal').onclick = () => $('modalBack').classList.remove('show'); }
  build.busy = false; $('buildBtn').disabled = false;
}
function showResult(res) {
  const m = $('modal'); $('modalBack').classList.add('show');
  const extra = [res.pdf, res.potx].filter(Boolean);
  m.innerHTML = `<h3>완성됐어요</h3><div class="hint">${esc(res.out)}${extra.map(x => '<br>' + esc(x)).join('')}</div>
    <div class="row" style="margin-top:10px"><button class="btn primary" id="openFile">PowerPoint로 열기</button>${res.pdf ? '<button class="btn" id="openPdf">PDF 열기</button>' : ''}<button class="btn" id="openFolder">폴더 열기</button><span class="grow"></span><button class="btn" id="closeModal">닫기</button></div>
    ${(res.log || '').includes('WARN') ? `<div class="warn" style="margin-top:8px">일부 요소는 그대로 옮기지 못했어요:<br>${esc(res.log.split('\n').filter(l => l.includes('WARN')).join('\n')).replace(/\n/g, '<br>')}</div>` : ''}
    <div class="results">${(res.pngs || []).map(p => `<img src="${p}" alt="">`).join('')}</div>`;
  $('openFile').onclick = () => fetch('/api/open', { method: 'POST', body: JSON.stringify({ what: 'file', path: res.out }) });
  if ($('openPdf')) $('openPdf').onclick = () => fetch('/api/open', { method: 'POST', body: JSON.stringify({ what: 'file', path: res.pdf }) });
  $('openFolder').onclick = () => fetch('/api/open', { method: 'POST', body: JSON.stringify({ what: 'folder', path: res.out }) });
  $('closeModal').onclick = () => $('modalBack').classList.remove('show');
}

// ---------------- slideshow ----------------
let show = null;
function present(start) {
  finishEdit();
  const order = state.deck.slides.map((s, i) => i).filter(i => !state.deck.slides[i].hidden);
  if (!order.length) return toast('발표할 슬라이드가 없어요.');
  show = { order, pos: Math.max(0, order.indexOf(start) >= 0 ? order.indexOf(start) : 0), step: 0 };
  $('show').hidden = false; try { const p = document.documentElement.requestFullscreen(); if (p && p.catch) p.catch(() => {}); } catch (e) {}
  drawShow(true);
}
function showSteps(idx) {
  const els = resolveSlide(state, idx).elements.filter(e => e.anim && e.anim.effect && e.anim.effect !== 'none').sort((a, b) => (a.anim.order || 0) - (b.anim.order || 0));
  const steps = []; els.forEach(e => { if (!steps.length || (e.anim.trigger || 'click') === 'click') steps.push([]); steps[steps.length - 1].push(e); });
  return steps;
}
function drawShow(entering) {
  const idx = show.order[show.pos], steps = showSteps(idx), hidden = new Set(steps.slice(show.step).flat().map(e => e.id));
  const host = $('showHost'), s = Math.min(window.innerWidth / SW, window.innerHeight / SH);
  host.style.width = SW * s + 'px'; host.style.height = SH * s + 'px';
  host.innerHTML = slideHtml(state, idx, { hideIds: hidden });
  const sl = host.querySelector('.slide'); sl.style.transform = `scale(${s})`;
  const tr = state.deck.slides[idx].transition; if (entering && tr && tr.effect) sl.classList.add('fadein');
  if (show.step > 0) steps[show.step - 1].forEach(e => { const n = host.querySelector(`.el[data-id="${CSS.escape(e.id)}"]`); if (n) { const map = { fade: 'a-fade', fly: 'a-fly', zoom: 'a-zoom', wipe: 'a-wipe', float: 'a-float', riseUp: 'a-float', split: 'a-zoom', bounce: 'a-bounce', wheel: 'a-spin', growTurn: 'a-spin', spin: 'a-spin', teeter: 'a-bounce' }; n.classList.add(map[e.anim.effect] || 'a-fade'); n.style.animationDuration = (e.anim.dur || 0.5) + 's'; n.style.animationDelay = (e.anim.delay || 0) + 's'; } });
  $('showCount').textContent = `${show.pos + 1} / ${show.order.length}`;
}
function showNext() { const steps = showSteps(show.order[show.pos]); if (show.step < steps.length) { show.step++; drawShow(false); return; } if (show.pos < show.order.length - 1) { show.pos++; show.step = 0; drawShow(true); } else endShow(); }
function showPrev() { if (show.step > 0) { show.step = 0; drawShow(false); return; } if (show.pos > 0) { show.pos--; show.step = showSteps(show.order[show.pos]).length; drawShow(false); } }
function endShow() { $('show').hidden = true; show = null; try { if (document.fullscreenElement) { const p = document.exitFullscreen(); if (p && p.catch) p.catch(() => {}); } } catch (e) {} }

// ---------------- projects ----------------
function renderProjects() { const s = $('projectSel'); s.innerHTML = ''; s.add(new Option('저장된 프로젝트…', '')); server.projects.forEach(p => s.add(new Option(p, p))); }
function normalizeState(s) {
  s.brand = normBrand(s.brand || clone(DEFAULT_BRAND));
  s.layout = s.layout || recommendLayout(null, []); for (const [k, vals] of Object.entries(KNOBS)) if (s.layout[k] === undefined) s.layout[k] = vals[0];
  if (s.layout.radius === undefined) s.layout.radius = 0.12; if (!s.layout.titleSize) s.layout.titleSize = 28;
  s.deck = migrateDeck(s.deck || defaultDeck()); s.assets = s.assets || []; s.v = 4;
  return s;
}

// ---------------- wiring ----------------
function wire() {
  const cv = $('canvas');
  cv.addEventListener('pointerdown', onCanvasDown);
  window.addEventListener('pointermove', onMove);
  window.addEventListener('pointerup', onUp);
  cv.addEventListener('dblclick', onDbl);
  $('canvasWrap').addEventListener('pointerdown', ev => { if (ev.target === $('canvasWrap') || ev.target === $('canvasPad')) { finishEdit(); sel = []; renderOverlay(); renderCtx(); renderProps(); } });
  $('canvasWrap').addEventListener('wheel', ev => { if (ev.ctrlKey) { ev.preventDefault(); zoomFit = false; zoom = Math.max(0.2, Math.min(4, zoom * (ev.deltaY < 0 ? 1.1 : 0.9))); fit(); } }, { passive: false });
  document.addEventListener('keydown', onKey);
  document.addEventListener('click', ev => { if (!ev.target.closest('.cpop') && !ev.target.closest('.menu')) closePops(); });
  window.addEventListener('resize', () => { fit(); });
  $('undoBtn').onclick = undo; $('redoBtn').onclick = redo;
  $('zoomIn').onclick = () => { zoomFit = false; zoom = Math.min(4, zoom * 1.2); fit(); };
  $('zoomOut').onclick = () => { zoomFit = false; zoom = Math.max(0.2, zoom / 1.2); fit(); };
  $('zoomFit').onclick = () => { zoomFit = true; fit(); };
  $('fileName').oninput = ev => { state.fileName = ev.target.value; store(); };
  $('buildBtn').onclick = build;
  $('buildMore').onclick = ev => { ev.stopPropagation(); $('buildMenu').classList.toggle('open'); };
  $('presentBtn').onclick = () => present(cur);
  $('show').addEventListener('click', showNext);
  document.addEventListener('keydown', ev => { if ($('show').hidden) return; if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter'].includes(ev.key)) { ev.preventDefault(); showNext(); } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(ev.key)) { ev.preventDefault(); showPrev(); } else if (ev.key === 'Escape') endShow(); });
  $('newDeck').onclick = () => { if (!confirm('지금 내용을 지우고 새로 시작할까요? (저장하지 않은 내용은 사라져요)')) return; snap(); state.deck = defaultDeck(); state.fileName = '새 프레젠테이션'; state.layout = recommendLayout(state.layout, server.history); cur = 0; sel = []; renderAll(); };
  $('saveProject').onclick = async () => { finishEdit(); const name = (state.fileName || '').trim() || '새 프레젠테이션'; await fetch('/api/project?name=' + encodeURIComponent(name), { method: 'POST', body: JSON.stringify(state, null, 1) }); if (!server.projects.includes(name)) server.projects.unshift(name); renderProjects(); toast(`'${name}' 프로젝트를 저장했어요.`); };
  $('loadProject').onclick = async () => { const name = $('projectSel').value; if (!name) { toast('불러올 프로젝트를 고르세요.'); return; } const r = await fetch('/api/project?name=' + encodeURIComponent(name)); if (!r.ok) { toast('불러오지 못했어요.'); return; } snap(); state = normalizeState(await r.json()); cur = 0; sel = []; renderAll(); toast(`'${name}'을(를) 불러왔어요.`); };
  $('logoFile').onchange = async () => { const f = $('logoFile').files[0]; if (!f) return; try { const r = await upload(f, 'logo'); state.brand.logo = { name: r.name, color: r.path, colorUrl: r.url, white: r.white, whiteUrl: r.whiteUrl, aspect: r.aspect }; vcCache = null; renderAll(); toast(r.white ? '로고를 넣었어요. 어두운 배경용 흰색 로고도 만들었어요.' : '로고를 넣었어요.'); } catch (e) { toast(e.message); } };
  // drop images onto the canvas / an image element
  const wrap = $('canvasWrap');
  wrap.addEventListener('dragover', ev => { if ([...ev.dataTransfer.items].some(i => i.kind === 'file')) { ev.preventDefault(); $('canvas').classList.add('drop'); } });
  wrap.addEventListener('dragleave', () => $('canvas').classList.remove('drop'));
  wrap.addEventListener('drop', ev => { $('canvas').classList.remove('drop'); const f = [...ev.dataTransfer.files].find(x => /^image\//.test(x.type)); if (!f) return; ev.preventDefault(); ev.stopPropagation(); const node = document.elementFromPoint(ev.clientX, ev.clientY); const elN = node && node.closest('.stage .el'); const target = elN ? elById(elN.dataset.id) : null; placeImageFile(f, target && (target.kind === 'image' || ['r:imageText', 'r:imageBox', 'r:image'].includes(target.id)) ? target : null); });
  document.addEventListener('paste', async ev => { if (editing || /input|textarea/i.test(document.activeElement.tagName)) return; const f = [...(ev.clipboardData || {}).files || []].find(x => /^image\//.test(x.type)); if (f) { ev.preventDefault(); placeImageFile(f, null); } });
}

// ---------------- boot ----------------
(async function boot() {
  try { const st = await (await fetch('/api/state')).json(); server = Object.assign(server, st); } catch (e) { toast('서버 정보를 불러오지 못했어요.'); }
  if (!Array.isArray(server.history)) server.history = server.history ? [server.history] : [];
  const hasPretendard = (server.fonts || []).includes('Pretendard');
  if (!server.brands) { const b = clone(DEFAULT_BRAND); if (!hasPretendard) { b.headFont = b.bodyFont = 'Malgun Gothic'; } server.brands = [b]; try { await saveBrands(); } catch (e) {} }
  server.brands = server.brands.map(normBrand);
  let s = restore();
  if (!s) { try { const old = JSON.parse(localStorage.getItem('ppt-editor-state') || 'null'); if (old && old.deck) s = old; } catch (e) {} }
  state = normalizeState(s || { fileName: '새 프레젠테이션', brand: clone(server.brands[0] || DEFAULT_BRAND), layout: recommendLayout(null, server.history), deck: defaultDeck() });
  renderProjects(); renderRail(); wire(); renderAll(); updateUndo();
  if (typeof wireImport === 'function') wireImport();
})();
