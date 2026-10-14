'use strict';
// ============================================================================
// 1) Theme link for imported slides: the original's colours are stored relative to theme colours, so changing the
//    brand colours recolours imported slides while keeping their relations (lighter / darker / shifted hue).
// 2) My templates (내 양식): any slide can be saved and reused; its colours follow the current brand.
// ============================================================================

// ---------- colours ----------
const hueDist = (a, b) => { const d = Math.abs(hsl(a)[0] - hsl(b)[0]); return Math.min(d, 360 - d); };
// every colour value of an element: fn(value) returns the replacement (or the same value)
function eachColor(e, fn) {
  const fix = (o, k) => { if (o && typeof o[k] === 'string' && o[k]) o[k] = fn(o[k], o, k); };
  fix(e, 'fill'); if (e.line) fix(e.line, 'color'); if (e.glow) fix(e.glow, 'color'); if (e.shadowFx) fix(e.shadowFx, 'color');
  if (e.gradient) (e.gradient.stops || []).forEach(s => fix(s, 'color'));
  if (e.text) { fix(e.text, 'color'); (e.text.paras || []).forEach(p => (p.runs || []).forEach(r => fix(r, 'color'))); }
  if (e.icon) fix(e.icon, 'color');
  if (e.chart) { (e.chart.series || []).forEach(s => fix(s, 'color')); if (e.chart.pointColors) e.chart.pointColors = e.chart.pointColors.map(c => fn(c, null, null)); }
}
// how much each literal colour is used (area for fills, glyph area for text)
function colorWeights(slides) {
  const w = new Map(), add = (c, v) => { if (!c || isThemeColor(c) || !/^#?[0-9a-f]{6}$/i.test(c)) return; const h = hx(c); w.set(h, (w.get(h) || 0) + v); };
  slides.forEach(sl => (sl.extra || []).forEach(e => {
    const area = Math.max(1, e.w * e.h);
    add(e.fill, area * (e.fillOpacity ?? 1));
    if (e.gradient) (e.gradient.stops || []).forEach(s => add(s.color, area / Math.max(1, e.gradient.stops.length)));
    if (e.line) add(e.line.color, (e.w + e.h) * (e.line.width || 1));
    if (e.text) (e.text.paras || []).forEach(p => (p.runs || []).forEach(r => add(r.color || e.text.color, String(r.t || '').length * Math.pow(r.size || e.text.size || 16, 2) * 0.6)));
    if (e.chart) (e.chart.pointColors || (e.chart.series || []).map(s => s.color)).forEach(c => add(c, area / 4));
  }));
  return w;
}
// up to three hue families of clearly coloured (saturated) colours: the most used becomes the main colour,
// then accent, then sub. Each family is represented by its most saturated colour, from which the lighter /
// darker members can be mixed.
function findColorFamilies(slides) {
  const list = [...colorWeights(slides)].filter(([h]) => { const [, s, l] = hsl(h); return s >= 0.35 && l > 0.15 && l < 0.9; }).sort((a, b) => b[1] - a[1]);
  const fams = [];
  list.forEach(([h, wt]) => {
    let f = fams.find(f => hueDist(f.first, h) <= 30);
    if (!f && fams.length < 3) { f = { first: h, members: [], weight: 0 }; fams.push(f); }
    if (f) { f.members.push([h, wt]); f.weight += wt; }
  });
  fams.forEach(f => {
    const smax = Math.max(...f.members.map(([h]) => hsl(h)[1]));
    f.rep = f.members.filter(([h]) => hsl(h)[1] >= smax - 0.12).sort((a, b) => b[1] - a[1])[0][0];
  });
  return fams.sort((a, b) => b.weight - a.weight);
}
const famBrand = (fams, base) => Object.assign({}, base, { main: '#' + fams[0].rep, accent: '#' + (fams[1] || fams[0]).rep, sub: '#' + (fams[2] || fams[1] || fams[0]).rep, subAuto: !fams[2] });
const FAMILY_TOKENS = ['dk2', 'accent1', 'accent2'];
// coloured literal colours -> 'token~dh,ds,dl': the hue / saturation / lightness difference to the theme colour of
// their hue family under palette pv. Exact now, and the same relation is kept when the brand colours change.
// Greys, near-white / near-black and colours of other hues stay fixed.
function linkColors(slides, fams, pv) {
  if (!fams.length) return 0;
  const memo = new Map(); let n = 0;
  const link = h => {
    const [ch, cs, cl] = hsl(h); if (cs < 0.12 || cl > 0.985 || cl < 0.03) return null;
    let bi = -1, bd = 1e9; fams.forEach((f, i) => { const d = hueDist(f.rep, h); if (d < bd) { bd = d; bi = i; } });
    if (bd > 50) return null;
    const tk = FAMILY_TOKENS[bi], [th, ts, tl] = hsl(pv[tk]);
    let dh = ch - th; if (dh > 180) dh -= 360; if (dh < -180) dh += 360;
    return `${tk}~${dh.toFixed(1)},${(cs - ts).toFixed(3)},${(cl - tl).toFixed(3)}`;
  };
  slides.forEach(sl => (sl.extra || []).forEach(e => eachColor(e, c => {
    if (!/^#?[0-9a-f]{6}$/i.test(c)) return c;
    const h = hx(c); if (!memo.has(h)) memo.set(h, link(h));
    const t = memo.get(h); if (t) { n++; return t; } return c;
  })));
  return n;
}
// fonts -> the brand's title / body fonts (keeps bold for heavy faces)
function linkFonts(slides) {
  slides.forEach(sl => (sl.extra || []).forEach(e => {
    if (!e.text) return;
    (e.text.paras || []).forEach(p => (p.runs || []).forEach(r => {
      const size = r.size || e.text.size || 16, fw = fontWeightOf(r.font);
      if (fw && fw >= 600 && r.b == null) r.b = true;
      if (size >= 20) r.font = 'head'; else delete r.font;
    }));
  }));
}

// ---------- after a free import: choose how colours and fonts relate to the brand ----------
function offerImportStyle(slides, missing) {
  const fams = findColorFamilies(slides);
  const m = $('modal'); $('modalBack').classList.add('show');
  const sw = h => `<span style="display:inline-block;width:22px;height:22px;border-radius:5px;background:#${h};vertical-align:middle;box-shadow:0 0 0 1px rgba(0,0,0,.1)"></span>`;
  const names = ['메인', '강조', '보조'];
  const pick = (server.fonts || []).filter(f => /Pretendard|Paperlogy|Nanum|Noto Sans KR|Malgun|맑은|나눔|본고딕|Spoqa|SUIT|Gmarket/i.test(f) && !FONT_WEIGHT_RE.test(f));
  const def = pick.find(f => f === 'Pretendard') || pick[0] || 'Malgun Gothic';
  m.innerHTML = `<h3>가져온 PPT의 색·글꼴을 어떻게 할까요?</h3>
    ${fams.length ? `<div class="hint">원본에서 찾은 색: ${fams.map((f, i) => `${sw(f.rep)} ${names[i]}`).join(' &nbsp; ')}</div>` : '<div class="hint">원본에서 주요 색을 찾지 못했어요 (무채색 위주).</div>'}
    <div class="choice" id="isOpt" style="margin-top:10px">
      <button data-v="link" class="on"><b>원본 색 그대로 + 색 조합에 연결 (추천)</b><span class="hint">지금은 원본과 똑같이 보여요. 디자인 탭에서 색 조합을 바꾸면 가져온 슬라이드 색도 함께 바뀌어요. 원본 색은 '가져온 색' 조합으로 저장돼요.</span></button>
      <button data-v="brand"><b>지금 내 색 조합·글꼴로 바꾸기</b><span class="hint">'${esc(state.brand.name || '현재 색 조합')}' 색과 제목·본문 글꼴로 바로 바꿔요. 배치는 원본 그대로예요.</span></button>
      <button data-v="keep"><b>바꾸지 않기</b><span class="hint">원본 색을 고정해서 가져와요. 색 조합을 바꿔도 이 슬라이드들은 그대로예요.</span></button>
    </div>
    ${missing.length ? `<div id="isFonts" style="margin-top:12px"><div class="hint">이 PC에 없는 글꼴 ${missing.length}개: ${missing.map(esc).join(', ')} — 그대로 두면 비슷한 글꼴로 대신 보여요.</div>
      <label class="row" style="gap:6px;margin-top:6px;font-size:12.5px"><input type="checkbox" id="isSwap" checked> 없는 글꼴은 <select id="isSel">${(pick.length ? pick : [def]).map(f => `<option ${f === def ? 'selected' : ''}>${esc(f)}</option>`).join('')}</select> (으)로 바꾸기</label></div>` : ''}
    <div class="row" style="margin-top:14px"><span class="grow"></span><button class="btn primary" id="isGo">적용</button></div>`;
  let opt = 'link';
  $('isOpt').onclick = ev => { const b = ev.target.closest('button'); if (!b) return; opt = b.dataset.v; $('isOpt').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); if ($('isFonts')) $('isFonts').style.display = opt === 'brand' ? 'none' : ''; };
  $('isGo').onclick = async () => {
    const deck = state.deck.slides;
    if (opt !== 'brand' && missing.length && $('isSwap') && $('isSwap').checked) swapMissingFonts(deck, missing, $('isSel').value);
    if (opt === 'link' && fams.length) {
      const b = normBrand(famBrand(fams, { name: `가져온 색: ${imp.name}`.slice(0, 40), headFont: state.brand.headFont, bodyFont: state.brand.bodyFont, logo: state.brand.logo }));
      // keep the current colours selectable before switching
      const list = server.brands || (server.brands = []);
      if (!list.some(x => x.name === state.brand.name)) list.push(clone(state.brand));
      const i = list.findIndex(x => x.name === b.name); if (i >= 0) list[i] = b; else list.push(b);
      state.brand = clone(b); try { await saveBrands(); } catch (e) {}
      const n = linkColors(deck, fams, pal());
      toast(`원본 색 ${fams.length}가지를 색 조합에 연결했어요 (${n}곳). 디자인 탭에서 색 조합을 바꾸면 함께 바뀌어요.`);
    } else if (opt === 'brand') {
      const n = fams.length ? linkColors(deck, fams, derivePalette(famBrand(fams, state.brand)).colors) : 0;
      linkFonts(deck);
      toast(`'${state.brand.name || '현재'}' 색 조합(${n}곳)과 글꼴로 바꿨어요.`);
    }
    $('modalBack').classList.remove('show'); vcCache = null; invalidate(); renderAll();
  };
}
function swapMissingFonts(slides, missing, to) {
  const set = new Set(missing), have = new Set(server.fonts || []);
  const weightOf = f => /ExtraBlack|Black|EBK|Heavy/i.test(f) ? 'Black' : /ExtraBold|\bEB\b|EB$/i.test(f) ? 'ExtraBold' : /SemiBold|DemiBold/i.test(f) ? 'SemiBold'
    : /Bold|\bB$|[^A-Z]B$/.test(f) ? 'Bold' : /Medium|\bM$/.test(f) ? 'Medium' : /ExtraLight|Thin/i.test(f) ? 'ExtraLight' : /Light|\bL$|[^A-Z]L$/.test(f) ? 'Light' : '';
  const swap = r => { const w = weightOf(r.font); if (w && have.has(`${to} ${w}`)) r.font = `${to} ${w}`; else { r.font = to; if (/Bold|Black/.test(w)) r.b = true; } };
  slides.forEach(sl => (sl.extra || []).forEach(e => { if (e.text) e.text.paras.forEach(p => p.runs.forEach(r => { if (r.font && set.has(r.font)) swap(r); })); }));
}

// ---------- my templates ----------
let myTemplates = null;   // [{ id, name, created, bg, elements }]
async function loadMyTemplates() {
  if (myTemplates) return myTemplates;
  try { const r = await fetch('/api/templates'); myTemplates = r.ok ? await r.json() : null; } catch (e) { myTemplates = null; }
  if (!Array.isArray(myTemplates)) { myTemplates = []; myTemplates.unavailable = true; }
  return myTemplates;
}
async function storeMyTemplates() {
  const r = await fetch('/api/templates', { method: 'POST', body: JSON.stringify(myTemplates, null, 1) });
  if (!r.ok) throw new Error('저장하지 못했어요. 편집기를 다시 시작한 뒤 해 보세요.');
}
// current slide -> template; literal colours that belong to the slide's hue families are linked to the theme
async function saveSlideAsTemplate() {
  await loadMyTemplates();
  if (myTemplates.unavailable) { toast('편집기를 한 번 다시 시작해야 내 양식을 쓸 수 있어요. (작업 표시줄의 PPT Editor 창을 닫고 다시 실행)'); return; }
  const name = prompt('내 양식 이름', `양식 ${myTemplates.length + 1}`); if (!name) return;
  const r = resolveSlide(state, cur);
  const els = clone(r.elements).map(e => { const o = Object.assign({}, e); delete o.role; delete o.locked; o.id = uid(); return o; });
  const tmp = [{ extra: els }], fams = findColorFamilies(tmp);
  // families are matched to the brand order of the current palette so the template keeps its look right now
  if (fams.length) linkColors(tmp, fams, derivePalette(famBrand(fams, state.brand)).colors);
  myTemplates.unshift({ id: uid(), name, created: new Date().toISOString().slice(0, 10), bg: clone(r.bg || {}), elements: els });
  try { await storeMyTemplates(); toast(`'${name}'을(를) 내 양식에 저장했어요. 템플릿 탭 맨 위에서 쓸 수 있어요.`); if (drawerTab === 'templates') renderDrawer(); }
  catch (e) { myTemplates.shift(); toast(e.message); }
}
function addFromMyTemplate(t, replace) {
  snap();
  const s = newSlide('free');
  s.extra = clone(t.elements).map(e => Object.assign(e, { id: uid() }));
  if (t.bg && (t.bg.image || (t.bg.color && t.bg.color !== 'lt1'))) s.bg = clone(t.bg);
  if (replace) { s.notes = slide().notes || ''; state.deck.slides[cur] = s; } else { state.deck.slides.splice(cur + 1, 0, s); cur++; }
  sel = []; invalidate(); renderAll(); toast(`'${t.name}' 양식을 ${replace ? '현재 슬라이드에 적용' : '추가'}했어요. 색은 지금 색 조합을 따라가요.`);
}
async function drawMyTemplates(d, mode) {
  const box = document.createElement('div'); d.prepend(box);
  box.innerHTML = `<h3 class="sec">내 양식 <button class="btn sm" id="mtSave" style="float:right">＋ 이 슬라이드 저장</button></h3><div class="tile-grid" id="mtGrid"></div>`;
  $('mtSave').onclick = saveSlideAsTemplate;
  await loadMyTemplates();
  const g = $('mtGrid'); if (!g) return;
  if (myTemplates.unavailable) { g.outerHTML = '<div class="hint">편집기를 다시 시작하면 내 양식을 쓸 수 있어요.</div>'; return; }
  if (!myTemplates.length) { g.outerHTML = '<div class="hint">마음에 드는 슬라이드(가져온 PPT 포함)를 열고 [＋ 이 슬라이드 저장]을 누르면 여기에 모여요. 색은 그때그때 색 조합을 따라가요.</div>'; return; }
  const v = vc();
  myTemplates.forEach(t => {
    const el = document.createElement('div'); el.className = 'tpl'; el.style.position = 'relative';
    const resolved = { bg: t.bg || {}, elements: t.elements };
    el.innerHTML = `<div class="fr">${slideHtml(state, cur, { vc: v, resolved })}</div><div class="cap">${esc(t.name)}</div><button title="삭제" style="position:absolute;top:4px;right:4px;border:none;border-radius:4px;background:rgba(20,24,40,.6);color:#fff;font-size:11px;cursor:pointer;padding:1px 6px">✕</button>`;
    g.appendChild(el); requestAnimationFrame(() => { const fr = el.querySelector('.fr'); el.querySelector('.slide').style.transform = `scale(${fr.clientWidth / SW})`; });
    el.onclick = ev => {
      if (ev.target.tagName === 'BUTTON') { ev.stopPropagation(); if (!confirm(`'${t.name}' 양식을 삭제할까요?`)) return; myTemplates.splice(myTemplates.indexOf(t), 1); storeMyTemplates().catch(e => toast(e.message)); renderDrawer(); return; }
      addFromMyTemplate(t, mode === 'change');
    };
  });
}
