'use strict';
// Shape catalog: SVG outlines for the preview that approximate PowerPoint's preset shapes.
// Names match engine/scripts/render.ps1 ($SHAPES), which maps them to MsoAutoShapeType.
const SHAPE_LIST = [
  ['rect', '사각형'], ['roundRect', '둥근 사각형'], ['ellipse', '원'], ['triangle', '삼각형'], ['rtTriangle', '직각삼각형'],
  ['diamond', '마름모'], ['parallelogram', '평행사변형'], ['trapezoid', '사다리꼴'], ['pentagon', '오각형'], ['hexagon', '육각형'],
  ['octagon', '팔각형'], ['star4', '별(4)'], ['star5', '별(5)'], ['star6', '별(6)'], ['plus', '십자'], ['heart', '하트'],
  ['arrowRight', '오른쪽 화살표'], ['arrowLeft', '왼쪽 화살표'], ['arrowUp', '위쪽 화살표'], ['arrowDown', '아래쪽 화살표'], ['arrowLR', '양쪽 화살표'],
  ['chevron', '갈매기형'], ['homePlate', '오각형 화살표'], ['callout', '말풍선'], ['cloud', '구름'], ['donut', '도넛'], ['lightning', '번개'],
  ['frame', '액자'], ['snipRect', '모서리 잘린 사각형'], ['round2Same', '위쪽 둥근 사각형']
];
const f1 = n => Math.round(n * 10) / 10;
function poly(pts) { return 'M' + pts.map(p => f1(p[0]) + ' ' + f1(p[1])).join(' L') + ' Z'; }
function starPts(w, h, n, inner) {
  const pts = [], cx = w / 2, cy = h / 2;
  for (let i = 0; i < n * 2; i++) { const r = i % 2 ? inner : 1, a = -Math.PI / 2 + i * Math.PI / n; pts.push([cx + Math.cos(a) * r * w / 2, cy + Math.sin(a) * r * h / 2]); }
  return poly(pts);
}
function regular(w, h, n, rot) {
  const pts = []; for (let i = 0; i < n; i++) { const a = (rot || -Math.PI / 2) + i * 2 * Math.PI / n; pts.push([w / 2 + Math.cos(a) * w / 2, h / 2 + Math.sin(a) * h / 2]); }
  return poly(pts);
}
function rrect(w, h, r) { r = Math.min(r, w / 2, h / 2); return `M${r} 0 H${w - r} A${r} ${r} 0 0 1 ${w} ${r} V${h - r} A${r} ${r} 0 0 1 ${w - r} ${h} H${r} A${r} ${r} 0 0 1 0 ${h - r} V${r} A${r} ${r} 0 0 1 ${r} 0 Z`; }
function shapePath(type, w, h, radius) {
  const m = Math.min(w, h), r = (radius == null ? 0.12 : radius) * m;
  switch (type) {
    case 'roundRect': return rrect(w, h, r);
    case 'ellipse': return `M${w / 2} 0 A${w / 2} ${h / 2} 0 1 1 ${w / 2 - 0.01} 0 Z`;
    case 'triangle': return poly([[w / 2, 0], [w, h], [0, h]]);
    case 'rtTriangle': return poly([[0, 0], [w, h], [0, h]]);
    case 'diamond': return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': { const d = m * 0.25; return poly([[d, 0], [w, 0], [w - d, h], [0, h]]); }
    case 'trapezoid': { const d = m * 0.25; return poly([[d, 0], [w - d, 0], [w, h], [0, h]]); }
    case 'pentagon': return regular(w, h, 5);
    case 'hexagon': { const d = w * 0.25; return poly([[d, 0], [w - d, 0], [w, h / 2], [w - d, h], [d, h], [0, h / 2]]); }
    case 'octagon': { const d = m * 0.29; return poly([[d, 0], [w - d, 0], [w, d], [w, h - d], [w - d, h], [d, h], [0, h - d], [0, d]]); }
    case 'star4': return starPts(w, h, 4, 0.38);
    case 'star5': return starPts(w, h, 5, 0.42);
    case 'star6': return starPts(w, h, 6, 0.55);
    case 'plus': { const a = w * 0.3, b = h * 0.3; return poly([[a, 0], [w - a, 0], [w - a, b], [w, b], [w, h - b], [w - a, h - b], [w - a, h], [a, h], [a, h - b], [0, h - b], [0, b], [a, b]]); }
    case 'heart': return `M${w / 2} ${h * 0.28} C${w * 0.5} ${-h * 0.05} ${-w * 0.08} ${h * 0.02} ${w * 0.02} ${h * 0.38} C${w * 0.1} ${h * 0.62} ${w * 0.38} ${h * 0.78} ${w / 2} ${h} C${w * 0.62} ${h * 0.78} ${w * 0.9} ${h * 0.62} ${w * 0.98} ${h * 0.38} C${w * 1.08} ${h * 0.02} ${w * 0.5} ${-h * 0.05} ${w / 2} ${h * 0.28} Z`;
    case 'arrowRight': { const t = h * 0.25, hw = Math.min(w * 0.45, h * 0.6); return poly([[0, t], [w - hw, t], [w - hw, 0], [w, h / 2], [w - hw, h], [w - hw, h - t], [0, h - t]]); }
    case 'arrowLeft': { const t = h * 0.25, hw = Math.min(w * 0.45, h * 0.6); return poly([[w, t], [hw, t], [hw, 0], [0, h / 2], [hw, h], [hw, h - t], [w, h - t]]); }
    case 'arrowUp': { const t = w * 0.25, hh = Math.min(h * 0.45, w * 0.6); return poly([[t, h], [t, hh], [0, hh], [w / 2, 0], [w, hh], [w - t, hh], [w - t, h]]); }
    case 'arrowDown': { const t = w * 0.25, hh = Math.min(h * 0.45, w * 0.6); return poly([[t, 0], [t, h - hh], [0, h - hh], [w / 2, h], [w, h - hh], [w - t, h - hh], [w - t, 0]]); }
    case 'arrowLR': { const t = h * 0.25, hw = Math.min(w * 0.3, h * 0.6); return poly([[0, h / 2], [hw, 0], [hw, t], [w - hw, t], [w - hw, 0], [w, h / 2], [w - hw, h], [w - hw, h - t], [hw, h - t], [hw, h]]); }
    case 'chevron': { const d = Math.min(w * 0.45, h * 0.5); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h], [d, h / 2]]); }
    case 'homePlate': { const d = Math.min(w * 0.45, h * 0.5); return poly([[0, 0], [w - d, 0], [w, h / 2], [w - d, h], [0, h]]); }
    case 'callout': { const bh = h * 0.78, rr = Math.min(r, bh / 2); return `M${rr} 0 H${w - rr} A${rr} ${rr} 0 0 1 ${w} ${rr} V${bh - rr} A${rr} ${rr} 0 0 1 ${w - rr} ${bh} H${w * 0.42} L${w * 0.22} ${h} L${w * 0.25} ${bh} H${rr} A${rr} ${rr} 0 0 1 0 ${bh - rr} V${rr} A${rr} ${rr} 0 0 1 ${rr} 0 Z`; }
    case 'cloud': return `M${w * 0.25} ${h * 0.85} C${w * 0.02} ${h * 0.85} ${w * 0.0} ${h * 0.5} ${w * 0.2} ${h * 0.45} C${w * 0.15} ${h * 0.15} ${w * 0.45} ${h * 0.05} ${w * 0.55} ${h * 0.25} C${w * 0.65} ${h * 0.05} ${w * 0.92} ${h * 0.15} ${w * 0.85} ${h * 0.42} C${w * 1.02} ${h * 0.5} ${w * 0.98} ${h * 0.85} ${w * 0.75} ${h * 0.85} Z`;
    case 'donut': { const t = m * 0.25; return `M${w / 2} 0 A${w / 2} ${h / 2} 0 1 1 ${w / 2 - 0.01} 0 Z M${w / 2} ${t} A${w / 2 - t} ${h / 2 - t} 0 1 0 ${w / 2 + 0.01} ${t} Z`; }
    case 'lightning': return poly([[w * 0.38, 0], [w * 0.62, h * 0.33], [w * 0.52, h * 0.38], [w * 0.82, h * 0.66], [w * 0.7, h * 0.7], [w, h], [w * 0.3, h * 0.74], [w * 0.42, h * 0.69], [0, h * 0.42], [w * 0.15, h * 0.36]]);
    case 'frame': { const t = m * 0.12; return `M0 0 H${w} V${h} H0 Z M${t} ${t} V${h - t} H${w - t} V${t} Z`; }
    case 'snipRect': { const d = r * 1.4; return poly([[0, 0], [w - d, 0], [w, d], [w, h], [0, h]]); }
    case 'round2Same': { const rr = Math.min(r, w / 2, h); return `M${rr} 0 H${w - rr} A${rr} ${rr} 0 0 1 ${w} ${rr} V${h} H0 V${rr} A${rr} ${rr} 0 0 1 ${rr} 0 Z`; }
    default: return `M0 0 H${w} V${h} H0 Z`;
  }
}
