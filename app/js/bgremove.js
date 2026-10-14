'use strict';
// ============================================================================
// Photo background removal (runs entirely in the browser, nothing leaves the PC)
//   자동      — removes the background connected to the picture's edges (plain / studio backgrounds)
//   클릭 지우기 — magic wand: removes the area of similar colour around a clicked point
//   지우개 / 복원 붓 — manual touch-up
// The result is saved as a new transparent PNG; the original upload is kept.
// ============================================================================
let bgr = null;   // { e, img, W, H, src: Uint8ClampedArray, mask: Uint8Array, hist: [], tool, tol, brush, feather, cv, ctx }

function openBgRemove(e) {
  if (!e || e.kind !== 'image' || !(e.image && e.image.url)) return toast('사진을 먼저 고르세요.');
  const img = new Image();
  img.onload = () => {
    const sc = Math.min(1, 2400 / Math.max(img.naturalWidth, img.naturalHeight));
    const W = Math.max(1, Math.round(img.naturalWidth * sc)), H = Math.max(1, Math.round(img.naturalHeight * sc));
    const c = document.createElement('canvas'); c.width = W; c.height = H; const x = c.getContext('2d'); x.drawImage(img, 0, 0, W, H);
    let data; try { data = x.getImageData(0, 0, W, H).data; } catch (err) { return toast('이 사진은 편집할 수 없어요.'); }
    const mask = new Uint8Array(W * H); for (let i = 0; i < W * H; i++) mask[i] = data[i * 4 + 3];
    bgr = { e, W, H, src: data, mask, hist: [], tool: 'wand', tol: 40, brush: 24, feather: true };
    bgrModal();
  };
  img.onerror = () => toast('사진을 불러오지 못했어요.');
  img.src = e.image.url;
}

function bgrModal() {
  const m = $('modal'); $('modalBack').classList.add('show');
  m.innerHTML = `<h3>배경 지우기</h3>
    <div class="hint">배경이 단색이거나 단순할수록 깔끔하게 지워져요. [자동]으로 먼저 지우고, 남은 부분은 클릭하거나 붓으로 다듬으세요. 사진은 이 PC 안에서만 처리돼요.</div>
    <div style="display:flex;gap:14px;margin-top:10px;flex-wrap:wrap">
      <div id="bgrWrap" style="flex:1 1 520px;min-width:280px;height:min(62vh,560px);display:flex;align-items:center;justify-content:center;border-radius:8px;
        background:repeating-conic-gradient(#E6E8EE 0 25%,#fff 0 50%) 0 0/18px 18px;overflow:hidden;position:relative">
        <canvas id="bgrCv" style="max-width:100%;max-height:100%;cursor:crosshair;touch-action:none"></canvas>
        <div id="bgrBrush" style="position:absolute;border:1.5px solid #E8613C;border-radius:50%;pointer-events:none;display:none"></div>
      </div>
      <div style="flex:0 0 220px">
        <button class="btn primary" id="bgrAuto" style="width:100%">✨ 자동으로 배경 지우기</button>
        <label class="f">도구</label>
        <div class="seg" id="bgrTools"><button data-t="wand">클릭 지우기</button><button data-t="erase">지우개</button><button data-t="restore">복원 붓</button></div>
        <label class="f">비슷한 색 범위 <b id="bgrTolV"></b></label><input type="range" id="bgrTol" min="5" max="120" step="1" style="width:100%">
        <label class="f">붓 크기 <b id="bgrBrV"></b></label><input type="range" id="bgrBr" min="4" max="120" step="1" style="width:100%">
        <label class="row" style="gap:6px;margin-top:10px;font-size:12.5px"><input type="checkbox" id="bgrFe"> 가장자리 부드럽게</label>
        <div class="row" style="gap:6px;margin-top:12px"><button class="btn sm" id="bgrUndo">↶ 되돌리기</button><button class="btn sm" id="bgrReset">처음으로</button></div>
        <div class="hint" style="margin-top:12px;line-height:1.6">· 클릭 지우기: 누른 곳과 비슷한 색을 한 번에 지워요.<br>· 지우개/복원 붓: 끌어서 직접 지우거나 되살려요.<br>· 범위가 너무 크면 사람·물건까지 지워져요. 줄여서 다시 해 보세요.</div>
      </div>
    </div>
    <div class="row" style="margin-top:14px"><span class="grow"></span><button class="btn" id="bgrCancel">취소</button><button class="btn primary" id="bgrApply">적용</button></div>`;
  const cv = $('bgrCv'); cv.width = bgr.W; cv.height = bgr.H; bgr.cv = cv; bgr.ctx = cv.getContext('2d');
  const setTool = t => { bgr.tool = t; $('bgrTools').querySelectorAll('button').forEach(b => b.classList.toggle('on', b.dataset.t === t)); };
  setTool(bgr.tool);
  $('bgrTools').onclick = ev => { const b = ev.target.closest('button'); if (b) setTool(b.dataset.t); };
  const tol = $('bgrTol'), br = $('bgrBr'); tol.value = bgr.tol; br.value = bgr.brush; $('bgrTolV').textContent = bgr.tol; $('bgrBrV').textContent = bgr.brush;
  tol.oninput = () => { bgr.tol = +tol.value; $('bgrTolV').textContent = tol.value; };
  br.oninput = () => { bgr.brush = +br.value; $('bgrBrV').textContent = br.value; };
  $('bgrFe').checked = bgr.feather; $('bgrFe').onchange = () => { bgr.feather = $('bgrFe').checked; bgrDraw(); };
  $('bgrAuto').onclick = () => { bgrPush(); bgrAuto(); bgrDraw(); };
  $('bgrUndo').onclick = () => { if (bgr.hist.length) { bgr.mask = bgr.hist.pop(); bgrDraw(); } };
  $('bgrReset').onclick = () => { bgrPush(); for (let i = 0; i < bgr.W * bgr.H; i++) bgr.mask[i] = bgr.src[i * 4 + 3]; bgrDraw(); };
  $('bgrCancel').onclick = () => { $('modalBack').classList.remove('show'); bgr = null; };
  $('bgrApply').onclick = bgrApply;
  bgrPointer(cv);
  bgrDraw();
}

function bgrPush() { bgr.hist.push(bgr.mask.slice()); if (bgr.hist.length > 20) bgr.hist.shift(); }
const bgrDist = (d, i, r, g, b) => { const dr = d[i] - r, dg = d[i + 1] - g, db = d[i + 2] - b; return Math.sqrt(dr * dr * 0.3 + dg * dg * 0.59 + db * db * 0.11) * 1.7; };

// flood fill from (seeds) over pixels similar to each seed's colour; value 0 = remove, 255 = restore
function bgrFlood(seeds, value) {
  const { W, H, src, mask } = bgr, tol = bgr.tol, seen = new Uint8Array(W * H), stack = new Int32Array(W * H);
  for (const s0 of seeds) {
    if (seen[s0]) continue;
    if (value === 0 && mask[s0] === 0) { seen[s0] = 1; }
    const r = src[s0 * 4], g = src[s0 * 4 + 1], b = src[s0 * 4 + 2];
    let sp = 0; stack[sp++] = s0; seen[s0] = 1;
    while (sp) {
      const p = stack[--sp]; mask[p] = value === 255 ? src[p * 4 + 3] : 0;
      const x = p % W, y = (p - x) / W;
      const nb = [x > 0 ? p - 1 : -1, x < W - 1 ? p + 1 : -1, y > 0 ? p - W : -1, y < H - 1 ? p + W : -1];
      for (const q of nb) {
        if (q < 0 || seen[q]) continue;
        // transparent pixels always join; otherwise the colour must be close to the seed
        if (src[q * 4 + 3] < 16 || bgrDist(src, q * 4, r, g, b) <= tol) { seen[q] = 1; stack[sp++] = q; }
      }
    }
  }
}
function bgrAuto() {
  const { W, H } = bgr, seeds = [];
  const step = Math.max(1, Math.round(Math.min(W, H) / 200));
  for (let x = 0; x < W; x += step) { seeds.push(x, (H - 1) * W + x); }
  for (let y = 0; y < H; y += step) { seeds.push(y * W, y * W + W - 1); }
  bgrFlood(seeds, 0);
  // drop tiny islands of background colour left inside the cut (specks along the edge)
  bgrDespeckle();
}
function bgrDespeckle() {
  const { W, H, mask } = bgr, lab = new Int32Array(W * H).fill(-1), stack = new Int32Array(W * H); const min = Math.max(12, Math.round(W * H * 0.00004));
  for (let s0 = 0; s0 < W * H; s0++) {
    if (mask[s0] === 0 || lab[s0] !== -1) continue;
    const comp = []; let sp = 0; stack[sp++] = s0; lab[s0] = s0;
    while (sp) { const p = stack[--sp]; if (comp.length <= min) comp.push(p); const x = p % W;
      for (const q of [x > 0 ? p - 1 : -1, x < W - 1 ? p + 1 : -1, p - W, p + W]) if (q >= 0 && q < W * H && mask[q] !== 0 && lab[q] === -1) { lab[q] = s0; stack[sp++] = q; } }
    if (comp.length <= min) comp.forEach(p => { mask[p] = 0; });
  }
}
function bgrBrushAt(cx, cy, value) {
  const { W, H, src, mask } = bgr, r = bgr.brush * (bgr.W / bgr.cv.getBoundingClientRect().width);
  for (let y = Math.max(0, Math.floor(cy - r)); y <= Math.min(H - 1, Math.ceil(cy + r)); y++)
    for (let x = Math.max(0, Math.floor(cx - r)); x <= Math.min(W - 1, Math.ceil(cx + r)); x++)
      if ((x - cx) ** 2 + (y - cy) ** 2 <= r * r) { const p = y * W + x; mask[p] = value ? src[p * 4 + 3] : 0; }
}
function bgrPointer(cv) {
  const pos = ev => { const b = cv.getBoundingClientRect(); return [(ev.clientX - b.left) * bgr.W / b.width, (ev.clientY - b.top) * bgr.H / b.height]; };
  const ring = $('bgrBrush'); let down = false, last = null;
  const showRing = ev => { if (bgr.tool === 'wand') { ring.style.display = 'none'; return; } const wb = $('bgrWrap').getBoundingClientRect(); const d = bgr.brush * 2;
    ring.style.display = 'block'; ring.style.width = ring.style.height = d + 'px'; ring.style.left = (ev.clientX - wb.left - bgr.brush) + 'px'; ring.style.top = (ev.clientY - wb.top - bgr.brush) + 'px'; };
  cv.onpointerdown = ev => {
    const [x, y] = pos(ev); if (x < 0 || y < 0 || x >= bgr.W || y >= bgr.H) return;
    bgrPush();
    if (bgr.tool === 'wand') { bgrFlood([Math.floor(y) * bgr.W + Math.floor(x)], 0); bgrDraw(); return; }
    down = true; last = [x, y]; cv.setPointerCapture(ev.pointerId); bgrBrushAt(x, y, bgr.tool === 'restore'); bgrDraw();
  };
  cv.onpointermove = ev => {
    showRing(ev); if (!down) return;
    const [x, y] = pos(ev), steps = Math.max(1, Math.ceil(Math.hypot(x - last[0], y - last[1]) / 3));
    for (let i = 1; i <= steps; i++) bgrBrushAt(last[0] + (x - last[0]) * i / steps, last[1] + (y - last[1]) * i / steps, bgr.tool === 'restore');
    last = [x, y]; bgrDraw();
  };
  cv.onpointerup = cv.onpointercancel = () => { down = false; };
  cv.onpointerleave = () => { ring.style.display = 'none'; };
}
// final alpha: mask, optionally softened by one pixel along the cut edge
function bgrAlpha() {
  const { W, H, mask } = bgr; if (!bgr.feather) return mask;
  const out = mask.slice();
  for (let y = 1; y < H - 1; y++) for (let x = 1; x < W - 1; x++) {
    const p = y * W + x; if (mask[p] === 0) continue;
    if (mask[p - 1] && mask[p + 1] && mask[p - W] && mask[p + W]) continue;
    let s = 0; for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) s += mask[p + dy * W + dx];
    out[p] = Math.min(mask[p], Math.round(s / 9));
  }
  return out;
}
function bgrDraw() {
  const { W, H, src } = bgr, a = bgrAlpha(), im = bgr.ctx.createImageData(W, H), d = im.data;
  d.set(src); for (let i = 0; i < W * H; i++) d[i * 4 + 3] = a[i];
  bgr.ctx.putImageData(im, 0, 0);
}
async function bgrApply() {
  const btn = $('bgrApply'); btn.disabled = true; btn.textContent = '저장 중…';
  try {
    bgrDraw();
    const blob = await new Promise(r => bgr.cv.toBlob(r, 'image/png'));
    const res = await upload(new File([blob], 'bg-removed.png', { type: 'image/png' }), 'image');
    const e = bgr.e; snap();
    patchEl(e.id, { image: Object.assign({}, e.image, { src: res.path, url: res.url }) });
    $('modalBack').classList.remove('show'); bgr = null;
    renderSoft(); renderDrawer(); toast('배경을 지웠어요. 원래 사진은 사진 모음에 그대로 있어요.');
  } catch (err) { toast(err.message || '저장하지 못했어요.'); btn.disabled = false; btn.textContent = '적용'; }
}
