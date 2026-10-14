'use strict';
// Renders resolved elements to HTML. Used by the canvas, thumbnails and the slideshow, so all three match.
const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function viewCtx(state) {
  const pal = derivePalette(state.brand).colors;
  return { pal, head: `'${state.brand.headFont}', 'Malgun Gothic', sans-serif`, body: `'${state.brand.bodyFont}', 'Malgun Gothic', sans-serif` };
}
const colorOf = (vc, c) => !c ? 'transparent' : vc.pal[c] ? '#' + vc.pal[c] : isThemeColor(c) ? '#' + resolveColor(c, vc.pal) : c;
// 'Paperlogy 7 Bold', 'Pretendard ExtraBold' ... : browsers often know only the base family + a weight
const FONT_WEIGHT_RE = /\s+(?:\d\s+)?(Thin|ExtraLight|UltraLight|Light|Regular|Medium|SemiBold|DemiBold|Bold|ExtraBold|UltraBold|Black|Heavy)$/i;
const FONT_WEIGHTS = { thin: 100, extralight: 200, ultralight: 200, light: 300, regular: 400, medium: 500, semibold: 600, demibold: 600, bold: 700, extrabold: 800, ultrabold: 800, black: 900, heavy: 900 };
function fontWeightOf(f) { const m = String(f || '').match(FONT_WEIGHT_RE); return m ? FONT_WEIGHTS[m[1].toLowerCase()] : null; }
function fontOf(vc, f) {
  if (f === 'head') return vc.head; if (!f || f === 'body') return vc.body;
  const name = String(f).replace(/'/g, ''), base = name.replace(FONT_WEIGHT_RE, '');
  return `'${name}', ${base !== name ? `'${base}', ` : ''}${vc.body}`;
}

function runHtml(vc, r, base) {
  const st = [];
  if (r.size) st.push(`font-size:${r.size}px`);
  if (r.font) st.push(`font-family:${fontOf(vc, r.font)}`);
  if (r.color) st.push(`color:${colorOf(vc, r.color)}`);
  const fw = fontWeightOf(r.font);
  if (r.b != null) st.push(`font-weight:${r.b ? Math.max(700, fw || 0) : fw || 400}`); else if (fw) st.push(`font-weight:${fw}`);
  if (r.i) st.push('font-style:italic');
  const deco = [r.u ? 'underline' : '', r.s ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) st.push(`text-decoration:${deco}`);
  const t = esc(r.t).replace(/\n/g, '<br>') || '';
  const attrs = (r.color && isThemeColor(r.color) ? ` data-c="${r.color}"` : '') + (r.link ? ` data-link="${esc(r.link)}"` : '');
  const inner = r.link ? `<u>${t}</u>` : t;
  return st.length || attrs ? `<span style="${st.join(';')}"${attrs}>${inner}</span>` : inner;
}
function textHtml(vc, tx, opts = {}) {
  if (!tx) return '';
  const ls = (tx.lineSpacing || 1.1) * 1.2;
  let num = 0;
  const paras = (tx.paras || []).map((p, i) => {
    const lvl = p.level || 1, bul = p.bullet || 'none';
    if (bul === 'num') num++; else num = 0;
    const marker = bul === 'disc' ? '●' : bul === 'square' ? '■' : bul === 'num' ? num + '.' : bul === 'char' ? esc(p.bulletChar || '•') : '';
    const indent = (lvl - 1) * 22 + (bul !== 'none' ? 18 : 0);
    const mk = marker ? `<span class="bul" style="left:${(lvl - 1) * 22}px;color:${colorOf(vc, tx.bulletColor || tx.color)};font-size:${bul === 'num' ? '1em' : bul === 'char' ? '0.9em' : '0.62em'}">${marker}</span>` : '';
    let body = (p.runs || []).map(r => runHtml(vc, r, tx)).join('') || '&#8203;';
    const sa = p.spaceAfter ?? tx.spaceAfter, sb = p.spaceBefore, al = p.align || tx.align || 'left';
    // no-wrap text overflows around its alignment point like PowerPoint (centered text spills out on both sides)
    const nw = tx.wrap === false && !opts.editing ? `display:flex;justify-content:${{ center: 'center', right: 'flex-end' }[al] || 'flex-start'};` : '';
    if (nw) body = `<span>${body}</span>`;
    // imported text: browsers set Korean text ~1% wider than PowerPoint, so give lines a little room before wrapping
    const slack = tx.fitPpt ? (al === 'center' ? 'margin-left:-1.5%;margin-right:-1.5%;' : al === 'right' ? 'margin-left:-3%;' : 'margin-right:-3%;') : '';
    return `<div class="p" data-level="${lvl}" data-bullet="${bul}"${bul === 'char' ? ` data-bchar="${esc(p.bulletChar || '•')}"` : ''} style="${nw}${slack}text-align:${al};padding-left:${indent}px;${sa != null && i < tx.paras.length - 1 ? `margin-bottom:${sa}px;` : ''}${sb && i > 0 ? `margin-top:${sb}px;` : ''}">${mk}${body}</div>`;
  }).join('');
  const field = tx.field === 'page' ? `<div class="p" style="text-align:${tx.align || 'right'}">${opts.page ?? ''}</div>` : '';
  const pad = tx.pad || opts.pad || [0, 0, 0, 0];
  const jc = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[tx.valign || 'top'];
  return `<div class="tx" style="padding:${pad[0]}px ${pad[1]}px ${pad[2]}px ${pad[3]}px;justify-content:${jc};font-family:${fontOf(vc, tx.font)};font-size:${tx.size || 16}px;color:${colorOf(vc, tx.color || 'dk1')};font-weight:${tx.bold ? 700 : 400};${tx.italic ? 'font-style:italic;' : ''}${tx.wrap === false ? 'white-space:pre;' : ''}line-height:${ls}">${field || paras}</div>`;
}
function lineDash(d, w) { return d === 'dash' ? `${w * 4} ${w * 3}` : d === 'dot' ? `${w} ${w * 2}` : d === 'dashDot' ? `${w * 4} ${w * 2} ${w} ${w * 2}` : d === 'longDash' ? `${w * 8} ${w * 3}` : ''; }
function rgba(c, a) { const h = String(c).replace('#', ''); if (!/^[0-9a-f]{6}$/i.test(h)) return c; return `rgba(${parseInt(h.slice(0, 2), 16)},${parseInt(h.slice(2, 4), 16)},${parseInt(h.slice(4, 6), 16)},${a ?? 1})`; }
// gradient fill: linear (angle in degrees, clockwise from left->right like PowerPoint) or radial from the centre
function gradientDef(vc, g, id) {
  const stops = [...(g.stops || [])].sort((a, b) => a.pos - b.pos).map(s => `<stop offset="${(s.pos * 100).toFixed(1)}%" stop-color="${colorOf(vc, s.color)}" stop-opacity="${s.alpha ?? 1}"/>`).join('');
  if (g.type === 'radial') return `<radialGradient id="${id}" cx="50%" cy="50%" r="70%">${stops}</radialGradient>`;
  const a = (g.angle ?? 90) * Math.PI / 180, dx = Math.cos(a) / 2, dy = Math.sin(a) / 2;
  return `<linearGradient id="${id}" x1="${0.5 - dx}" y1="${0.5 - dy}" x2="${0.5 + dx}" y2="${0.5 + dy}">${stops}</linearGradient>`;
}
function shapeSvg(vc, e) {
  const w = Math.max(1, e.w), h = Math.max(1, e.h), ln = e.line, sw = ln && ln.color ? (ln.width || 1) : 0;
  const d = shapePath(e.shape || 'rect', w, h, e.radius, e.geom);
  let fill = e.fill ? colorOf(vc, e.fill) : 'none', defs = '';
  if (e.gradient && e.gradient.stops && e.gradient.stops.length) { const id = 'g' + Math.random().toString(36).slice(2, 8); defs = `<defs>${gradientDef(vc, e.gradient, id)}</defs>`; fill = `url(#${id})`; }
  return `<svg class="shp" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" overflow="visible">${defs}<path d="${d}" fill="${fill}" fill-opacity="${e.fillOpacity ?? 1}" fill-rule="evenodd" ${sw ? `stroke="${colorOf(vc, ln.color)}" stroke-width="${sw}" stroke-dasharray="${lineDash(ln.dash, sw)}"` : ''}/></svg>`;
}
// PowerPoint-like effects as CSS filters: shadow (colour / blur / offset), glow, soft edge
function effectFilter(vc, e) {
  const f = [];
  if (e.softEdge) f.push(`blur(${(e.softEdge * 0.45).toFixed(1)}px)`);
  if (e.glow && e.glow.r) f.push(`drop-shadow(0 0 ${(e.glow.r * 0.6).toFixed(1)}px ${rgba(colorOf(vc, e.glow.color || '#FFFFFF'), e.glow.alpha ?? 0.6)})`);
  if (e.shadow) { const s = e.shadowFx; f.push(s ? `drop-shadow(${s.dx || 0}px ${s.dy || 0}px ${((s.blur || 0) * 0.5).toFixed(1)}px ${rgba(colorOf(vc, s.color || '#000000'), s.alpha ?? 0.35)})` : 'drop-shadow(0 4px 8px rgba(0,0,0,.25))'); }
  return f.length ? `filter:${f.join(' ')}` : '';
}
function lineSvg(vc, e) {
  const w = Math.max(1, Math.abs(e.w)), h = Math.max(1, Math.abs(e.h)), ln = e.line || {}, sw = ln.width || 2, col = colorOf(vc, ln.color || 'dk1');
  let x1 = 0, y1 = 0, x2 = e.w, y2 = e.h; if (e.flipH) { x1 = e.w; x2 = 0; } if (e.flipV) { y1 = e.h; y2 = 0; }
  const id = 'm' + Math.random().toString(36).slice(2, 7);
  const marker = `<marker id="${id}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="${Math.max(3, 9 / sw * 2)}" markerHeight="${Math.max(3, 9 / sw * 2)}" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${col}"/></marker>`;
  return `<svg class="shp" width="${w}" height="${h}" overflow="visible"><defs>${marker}</defs><line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${col}" stroke-width="${sw}" stroke-dasharray="${lineDash(ln.dash, sw)}" ${ln.end === 'arrow' ? `marker-end="url(#${id})"` : ''} ${ln.start === 'arrow' ? `marker-start="url(#${id})"` : ''}/></svg>`;
}
function tableHtml(vc, e) {
  const t = e.table, rows = t.rows || [], nr = rows.length, rh = e.h / Math.max(1, nr), cw = t.colW || [];
  const sc = t.statusColors || {};
  return `<table class="tbl" style="font-size:${t.fontSize || 14}px;font-family:${vc.body}">` + rows.map((r, ri) => '<tr>' + r.map((v, ci) => {
    const head = t.header !== false && ri === 0, band = !head && t.banded !== false && ri % 2 === 0;
    const bg = head ? colorOf(vc, t.headerFill || 'dk2') : band ? colorOf(vc, t.bandFill || 'lt2') : colorOf(vc, t.bodyFill || 'lt1');
    const k = !head && sc[v];
    const col = head ? colorOf(vc, t.headerColor || 'lt1') : k ? colorOf(vc, k) : colorOf(vc, t.color || 'dk1');
    const bold = head || k || (ci === 0 && t.firstColBold !== false);
    return `<td style="width:${cw[ci] || e.w / r.length}px;height:${rh}px;background:${bg};color:${col};font-weight:${bold ? 700 : 400};text-align:${t.align || 'left'}${t.borderColor ? `;border-bottom:0.75px solid ${colorOf(vc, t.borderColor)}` : ''}">${esc(v)}</td>`;
  }).join('') + '</tr>').join('') + '</table>';
}
function niceMax(v) { const raw = Math.max(1, v) * 1.1, mag = Math.pow(10, Math.floor(Math.log10(raw))); return Math.ceil(raw / (mag / 2)) * (mag / 2); }
// Excel-style number format subset: 0.0%, #,##0, 0.00
function fmtVal(v, f) {
  if (v === '' || v == null || isNaN(+v)) return v ?? '';
  if (!f) return String(v);
  const dec = ((f.split(/[%;]/)[0].split('.')[1] || '').match(/0/g) || []).length;
  if (/%/.test(f)) return (+v * 100).toFixed(dec) + '%';
  const n = (+v).toFixed(dec);
  return /,/.test(f) ? (+n).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec }) : n;
}
function chartSvg(vc, e) {
  const c = e.chart, cats = c.categories || [], series = c.series || [], W = e.w, H = e.h, n = cats.length;
  if (!n || !series.length) return '';
  const palette = (c.colors || ['dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent6']).map(k => colorOf(vc, k));
  const kind = c.type || 'column', hl = Number(c.highlight ?? -1), dk1 = colorOf(vc, 'dk1'), ax = '#6B7280', ff = vc.body;
  const legend = (labels, cols) => `<div style="position:absolute;left:0;bottom:0;width:${W}px;text-align:center;font:11px ${ff};color:${dk1}">` + labels.map((l, i) => `<span style="margin:0 7px;white-space:nowrap"><i style="display:inline-block;width:10px;height:10px;background:${cols[i % cols.length]};margin-right:4px;vertical-align:-1px"></i>${esc(l)}</span>`).join('') + '</div>';
  if (kind === 'pie' || kind === 'doughnut') {
    const vals = (series[0].values || []).map(Number), tot = vals.reduce((a, b) => a + Math.max(0, b), 0) || 1; let acc = 0;
    const stops = vals.map((v, i) => { const a0 = acc / tot * 360; acc += Math.max(0, v); return `${palette[i % palette.length]} ${a0}deg ${acc / tot * 360}deg`; }).join(',');
    const s = Math.min(W, H - 30);
    return `<div style="position:absolute;left:${(W - s) / 2}px;top:0;width:${s}px;height:${s}px;border-radius:50%;background:conic-gradient(${stops});${kind === 'doughnut' ? `-webkit-mask:radial-gradient(circle,transparent ${s * 0.27}px,#000 ${s * 0.27 + 1}px);mask:radial-gradient(circle,transparent ${s * 0.27}px,#000 ${s * 0.27 + 1}px)` : ''}"></div>` + legend(cats, palette);
  }
  const showLegend = series.length > 1;
  const allVals = series.flatMap(s => (s.values || []).map(Number));
  const stacked = /Stacked/.test(kind);
  let maxV = c.axisMax != null ? +c.axisMax : niceMax(stacked ? Math.max(...cats.map((_, i) => series.reduce((a, s) => a + (Number((s.values || [])[i]) || 0), 0))) : Math.max(...allVals, 0));
  let minV = c.axisMin != null && !stacked ? +c.axisMin : 0;
  // like PowerPoint/Excel: when every value sits in a narrow band far from zero, the axis does not start at zero
  const lo = Math.min(...allVals), hi = Math.max(...allVals);
  if (c.axisMin == null && !stacked && lo > 0 && lo > hi * 5 / 6) {
    const d = hi - lo || hi * 0.1, mag = 10 ** Math.floor(Math.log10(d / 3)), st = [1, 2, 5, 10].map(k => k * mag).find(x => x >= d / 3) || mag * 10;
    minV = Math.max(0, Math.floor((lo - d / 2) / st) * st); if (c.axisMax == null) maxV = Math.ceil((hi + d * 0.05) / st) * st;
  }
  const scale = v => Math.max(0, v - minV) / ((maxV - minV) || 1);
  const padL = c.valueAxis === false && !kind.startsWith('bar') ? 6 : 36, padB = showLegend ? 44 : 24, pw = W - padL, ph = H - padB - 12;
  let s = `<svg width="${W}" height="${H}" style="position:absolute;left:0;top:0;font-family:${ff}" overflow="visible">`;
  if (kind.startsWith('bar')) {
    const band = ph / n, bh = band * 0.6 / (stacked ? 1 : series.length);
    cats.forEach((ct, i) => {
      s += `<text x="${padL + 20}" y="${12 + i * band + band / 2 + 4}" font-size="11" fill="${dk1}" text-anchor="end">${esc(ct)}</text>`;
      let accX = 0;
      series.forEach((se, k) => { const v = Number((se.values || [])[i]) || 0, w = (pw - 70) * scale(v), y = 12 + i * band + (band - bh * (stacked ? 1 : series.length)) / 2 + (stacked ? 0 : k * bh);
        const col = k === 0 && i === hl ? colorOf(vc, 'accent1') : k === 0 && c.pointColors && c.pointColors[i] ? colorOf(vc, c.pointColors[i]) : (se.color ? colorOf(vc, se.color) : palette[k % palette.length]);
        s += `<rect x="${padL + 28 + accX}" y="${y}" width="${w}" height="${bh}" fill="${col}"/>`;
        if (c.showValues !== false) s += `<text x="${padL + 34 + accX + w}" y="${y + bh / 2 + 4}" font-size="11" font-weight="700" fill="${dk1}">${fmtVal(v, c.numFmt)}</text>`;
        if (stacked) accX += w; });
    });
    return s + '</svg>' + (showLegend ? legend(series.map(x => x.name), palette) : '');
  }
  for (let g = 0; g <= 4; g++) { const y = 12 + ph - ph * g / 4;
    if (c.gridlines !== false || g === 0) s += `<line x1="${padL}" x2="${W}" y1="${y}" y2="${y}" stroke="#E3E7EE"/>`;
    if (c.valueAxis !== false) s += `<text x="${padL - 6}" y="${y + 4}" font-size="10" fill="${ax}" text-anchor="end">${fmtVal(minV + (maxV - minV) * g / 4, c.numFmt || '0')}</text>`; }
  const step = pw / n;
  if (kind.startsWith('line') || kind.startsWith('area') || kind === 'radar') {
    series.forEach((se, k) => {
      const col = se.color ? colorOf(vc, se.color) : palette[k % palette.length];
      const pts = cats.map((_, i) => [padL + step * i + step / 2, 12 + ph - ph * scale(Number((se.values || [])[i]) || 0)]);
      if (kind.startsWith('area')) s += `<polygon points="${padL + step / 2},${12 + ph} ${pts.map(p => p.join(',')).join(' ')} ${padL + step * (n - 1) + step / 2},${12 + ph}" fill="${col}" fill-opacity="0.35"/>`;
      s += `<polyline fill="none" stroke="${col}" stroke-width="3" points="${pts.map(p => p.join(',')).join(' ')}"/>`;
      pts.forEach((p, i) => { s += `<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="${col}"/>`; if (c.showValues !== false) s += `<text x="${p[0]}" y="${p[1] - 9}" font-size="11" font-weight="700" text-anchor="middle" fill="${dk1}">${fmtVal((se.values || [])[i], c.numFmt)}</text>`; });
    });
  } else {
    const groupW = step * 0.64, bw = stacked ? groupW : groupW / series.length;
    cats.forEach((_, i) => { let accY = 0;
      series.forEach((se, k) => { const v = Number((se.values || [])[i]) || 0, bh = stacked ? ph * v / maxV : ph * scale(v), x = padL + step * i + (step - groupW) / 2 + (stacked ? 0 : k * bw), y = 12 + ph - bh - accY;
        const col = k === 0 && i === hl ? colorOf(vc, 'accent1') : k === 0 && c.pointColors && c.pointColors[i] ? colorOf(vc, c.pointColors[i]) : (se.color ? colorOf(vc, se.color) : palette[k % palette.length]);
        s += `<rect x="${x}" y="${y}" width="${bw}" height="${bh}" fill="${col}"/>`;
        if (c.showValues !== false && !stacked) s += `<text x="${x + bw / 2}" y="${y - 5}" font-size="11" font-weight="700" text-anchor="middle" fill="${dk1}">${fmtVal(v, c.numFmt)}</text>`;
        if (stacked) accY += bh; }); });
  }
  cats.forEach((ct, i) => { s += `<text x="${padL + step * i + step / 2}" y="${12 + ph + 16}" font-size="11" text-anchor="middle" fill="${dk1}">${esc(ct)}</text>`; });
  return s + '</svg>' + (showLegend ? legend(series.map(x => x.name), palette) : '');
}
function elementHtml(vc, e, opts = {}) {
  const tf = [];
  if (e.rot) tf.push(`rotate(${e.rot}deg)`);
  if (e.kind !== 'line') { if (e.flipH) tf.push('scaleX(-1)'); if (e.flipV) tf.push('scaleY(-1)'); }
  const st = [`left:${e.x}px`, `top:${e.y}px`, `width:${Math.max(1, e.w)}px`, `height:${Math.max(1, e.h)}px`];
  if (tf.length) st.push(`transform:${tf.join(' ')}`);
  if (e.opacity != null && e.opacity < 1) st.push(`opacity:${e.opacity}`);
  const fx = effectFilter(vc, e); if (fx) st.push(fx);
  if (opts.hidden) st.push('visibility:hidden');
  let inner = '';
  const counter = e.flipH || e.flipV ? `style="transform:${e.flipH ? 'scaleX(-1)' : ''} ${e.flipV ? 'scaleY(-1)' : ''}"` : '';
  switch (e.kind) {
    case 'text': {
      const bgc = e.fill ? `background:${colorOf(vc, e.fill)};` : '', bd = e.line && e.line.color ? `outline:${e.line.width || 1}px solid ${colorOf(vc, e.line.color)};outline-offset:-${e.line.width || 1}px;` : '';
      inner = `<div class="fill" style="${bgc}${bd}" ${counter}>${textHtml(vc, e.text, { page: opts.page })}</div>`; break;
    }
    case 'shape': inner = shapeSvg(vc, e) + (e.text ? `<div class="fill" ${counter}>${textHtml(vc, e.text, { pad: [4, 7, 4, 7], page: opts.page })}</div>` : ''); break;
    case 'line': inner = lineSvg(vc, e); break;
    case 'image': {
      const im = e.image || {}, fit = im.fit === 'contain' ? 'contain' : im.fit === 'stretch' ? 'fill' : 'cover';
      const pos = `${(im.focusX ?? 0.5) * 100}% ${(im.focusY ?? 0.5) * 100}%`;
      const shape = e.shape && e.shape !== 'rect' ? (e.shape === 'roundRect' ? `border-radius:${Math.min(e.w, e.h) * (e.radius ?? 0.1)}px;` : e.shape === 'ellipse' ? 'border-radius:50%;' : `clip-path:path('${shapePath(e.shape, e.w, e.h, e.radius)}');`) : '';
      const bd = e.line && e.line.color ? `outline:${e.line.width || 1}px solid ${colorOf(vc, e.line.color)};outline-offset:-${e.line.width || 1}px;` : '';
      inner = im.url ? `<img src="${esc(im.url)}" draggable="false" style="width:100%;height:100%;object-fit:${fit};object-position:${pos};${shape}${bd}">` : `<div class="fill ph">이미지</div>`; break;
    }
    case 'icon': inner = iconSvg(e.icon.name, colorOf(vc, e.icon.color || 'dk2'), e.icon.stroke).replace('width="96" height="96"', 'width="100%" height="100%"'); break;
    case 'table': inner = tableHtml(vc, e); break;
    case 'chart': inner = `<div class="fill">${chartSvg(vc, e)}</div>`; break;
  }
  return `<div class="el k-${e.kind}" data-id="${esc(e.id || '')}" style="${st.join(';')}">${inner}</div>`;
}
function slideHtml(state, idx, opts = {}) {
  const vc = opts.vc || viewCtx(state), r = opts.resolved || resolveSlide(state, idx);
  const bg = r.bg.image ? `background:url('${esc(r.bg.imageUrl || '')}') center/cover` : `background:${colorOf(vc, r.bg.color || 'lt1')}`;
  const hide = opts.hideIds || new Set();
  return `<div class="slide" style="${bg};width:${SW}px;height:${SH}px">` + r.elements.map(e => elementHtml(vc, e, { page: idx + 1, hidden: hide.has(e.id) })).join('') + '</div>';
}
